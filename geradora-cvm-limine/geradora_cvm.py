"""
Geradora CVM Limine (Demonstrativo Trimestral) em Python.

Substitui a planilha "Geradora Limine - EMISSÃO EM LOTE.xlsb" e os macros dela:

    imprimir_exportar             ->  python geradora_cvm.py gerar
    BaixarEImportarPipefyReport   ->  python geradora_cvm.py pipefy
    listarArquivos                ->  python geradora_cvm.py listar-validados
    (abas coladas à mão na planilha) -> python geradora_cvm.py migrar --geradora <arquivo.xlsb>

    popularGeradora / populaCarteiras_ não foram portados: estão quebrados no
    layout atual do Setup (leem CNPJ na coluna do nome e usam o caminho do 2T2025).
    Veja o README.

As regras de cálculo reproduzem as fórmulas da aba CVM célula a célula; os
comentários "CVM!X" apontam a célula de origem. A formatação vem de
assets/modelo_cvm.xlsx, que é a aba CVM salva pelo próprio Excel: o XLSX é esse
arquivo com as células calculadas preenchidas, e o PDF é desenhado a partir da
mesma grade (alturas, larguras, mesclagens e estilos), como o Excel imprime.
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import io
import json
import math
import os
import posixpath
import re
import sys
import time
import urllib.request
import zipfile
from collections import defaultdict
from dataclasses import dataclass, field
from pathlib import Path
from xml.sax.saxutils import escape

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock
from openpyxl.utils import get_column_letter, range_boundaries

BASE_DIR = Path(__file__).resolve().parent
MODELO_XLSX = BASE_DIR / "assets" / "modelo_cvm.xlsx"

# Linhas da aba CVM ocultas na planilha original (seções II e III).
LINHAS_SECAO_LASTRO = range(22, 48)

# Célula de texto do gestor -> coluna do relatório Pipefy (CVM!I5:AE5 = Pipefy!A:W).
# CVM!A52 =IF(V5="",$I$8,V5) e V5 = Pipefy coluna N ("Descrição item 9"), etc.
TEXTOS_GESTOR = {
    52: "Descrição item 9",
    59: "Descrição item 1",
    66: "Descrição item 2.a",
    72: "Descrição item 2.b",
    77: "Descrição item 3",
    84: "Descrição item 4.a",
    94: "Descrição item 5",
    100: "Descrição item 6.a",
    112: "Descrição item 8",
}

# Respostas do gestor informadas à mão (entrada/respostas_gestor.xlsx): coluna -> célula do relatório.
# Têm prioridade sobre o Pipefy e sobre os textos fixos do modelo (A89 e A106).
SECOES_GESTOR = {
    "IV - Processos judiciais": 52,
    "V.a - Política de investimento": 59,
    "V.b.i - Critérios de concessão": 66,
    "V.b.ii - Alterações nos critérios": 72,
    "V.c - Garantias": 77,
    "V.d.i - Contratos relevantes": 84,
    "V.d.ii - Caráter da cessão": 89,
    "V.e - Pré-pagamento": 94,
    "V.f - Alienação": 100,
    "V.g - Descontinuidade": 106,
    "V.h - Regularidade dos fluxos": 112,
}

CARACTERES_INVALIDOS ='/\\:*?"<>|'
XML_INVALIDO = re.compile("[\x00-\x08\x0b\x0c\x0e-\x1f\ufffe\uffff]")
DATA_ZERO_EXCEL = dt.date(1899, 12, 30)


# --------------------------------------------------------------------------
# utilidades
# --------------------------------------------------------------------------
def digitos(valor) -> str:
    if valor is None:
        return ""
    if isinstance(valor, float) and valor.is_integer():
        valor = int(valor)
    return re.sub(r"\D", "", str(valor))


def norm(texto) -> str:
    """Comparação de texto no estilo do Excel (sem diferenciar maiúsculas)."""
    return str(texto).strip().casefold() if texto is not None else ""


def num(valor) -> float:
    """SUMIFS ignora o que não é número."""
    if isinstance(valor, bool):
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    return 0.0


def para_data(valor) -> dt.date | None:
    if valor is None or valor == "":
        return None
    if isinstance(valor, dt.datetime):
        return valor.date()
    if isinstance(valor, dt.date):
        return valor
    if isinstance(valor, (int, float)):  # serial do Excel
        return (dt.datetime(1899, 12, 30) + dt.timedelta(days=float(valor))).date()
    texto = str(valor).strip()
    for formato in ("%Y-%m-%d", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S"):
        try:
            return dt.datetime.strptime(texto, formato).date()
        except ValueError:
            pass
    return None


def div(a: float, b: float) -> float:
    """IFERROR(a/b, 0)."""
    return a / b if b else 0.0


def numero_excel(x: float) -> str:
    """Número concatenado em texto no Excel pt-BR (formato Geral, 15 dígitos)."""
    if float(x).is_integer():
        return str(int(x))
    return f"{x:.15g}".replace(".", ",")


def nome_arquivo_seguro(texto: str) -> str:
    for ch in CARACTERES_INVALIDOS:
        texto = texto.replace(ch, "-")
    return texto


def numero_csv(texto: str):
    """Converte o texto de uma célula de CSV em número quando ele é um número.

    Aceita "1234.56", "1234,56" e "1.234,56". Não converte texto com zero à esquerda
    (CNPJ só com dígitos, códigos) nem datas e CNPJ formatados.
    """
    t = texto.strip()
    if not t or not re.fullmatch(r"-?[\d.,]+", t) or re.match(r"-?0\d", t):
        return texto
    if "," in t and "." in t:
        t = t.replace(".", "").replace(",", ".") if t.rfind(",") > t.rfind(".") else t.replace(",", "")
    elif t.count(",") == 1:
        t = t.replace(",", ".")
    elif t.count(",") > 1 or t.count(".") > 1:
        t = t.replace(",", "").replace(".", "")
    try:
        valor = float(t)
    except ValueError:
        return texto
    return int(valor) if valor.is_integer() and "." not in t else valor


def caminho_gravavel(caminho: Path) -> Path:
    """No Windows, usa o prefixo \\\\?\\ quando o caminho passa do limite de 260 caracteres."""
    if os.name != "nt":
        return caminho
    texto = str(caminho)
    if texto.startswith("\\\\?\\") or len(texto) < 240:
        return caminho
    texto = os.path.abspath(texto)
    if texto.startswith("\\\\"):
        return Path("\\\\?\\UNC\\" + texto[2:])
    return Path("\\\\?\\" + texto)


def ler_config(caminho: Path) -> dict:
    with open(caminho, encoding="utf-8") as f:
        cfg = json.load(f)
    cfg["_dir"] = caminho.resolve().parent
    return cfg


def caminho_cfg(cfg: dict, chave: str) -> Path:
    p = Path(cfg["pastas"][chave])
    return p if p.is_absolute() else cfg["_dir"] / p


def arquivo_entrada(cfg: dict, chave: str) -> Path:
    return caminho_cfg(cfg, "entrada") / cfg["arquivos"][chave]


# --------------------------------------------------------------------------
# leitura das bases (xlsx, xlsm, xlsb ou csv; cabeçalho na 1ª linha)
# --------------------------------------------------------------------------
def ler_tabela(caminho: Path, aba: str | None = None) -> list[dict]:
    caminho = Path(caminho)
    if not caminho.exists():
        raise FileNotFoundError(f"Arquivo de entrada não encontrado: {caminho}")
    ext = caminho.suffix.lower()
    if ext == ".csv":
        with open(caminho, encoding="utf-8-sig", newline="") as f:
            amostra = f.read(4096)
            f.seek(0)
            sep = ";" if amostra.count(";") > amostra.count(",") else ","
            linhas = list(csv.reader(f, delimiter=sep))
        # no CSV tudo é texto; números viram números para entrar nas somas
        linhas = linhas[:1] + [[numero_csv(v) for v in linha] for linha in linhas[1:]]
    elif ext == ".xlsb":
        from pyxlsb import open_workbook

        with open_workbook(str(caminho)) as wb:
            nome = aba or wb.sheets[0]
            with wb.get_sheet(nome) as sh:
                linhas = []
                for row in sh.rows():
                    linha = []
                    for c in row:
                        while len(linha) < c.c:
                            linha.append(None)
                        linha.append(c.v)
                    linhas.append(linha)
    else:
        wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
        ws = wb[aba] if aba else wb.worksheets[0]
        linhas = [list(r) for r in ws.iter_rows(values_only=True)]
        wb.close()
    if not linhas:
        return []
    cab = [str(c).strip() if c is not None else f"_col{i}" for i, c in enumerate(linhas[0])]
    saida = []
    for linha in linhas[1:]:
        if linha is None or all(v in (None, "") for v in linha):
            continue
        saida.append({cab[i]: (linha[i] if i < len(linha) else None) for i in range(len(cab))})
    return saida


def coluna(linha: dict, *nomes):
    """Busca a coluna por nome, sem diferenciar maiúsculas/acentos simples."""
    chaves = {norm(k): k for k in linha}
    for n in nomes:
        k = chaves.get(norm(n))
        if k is not None:
            return linha[k]
    return None


@dataclass
class Bases:
    fundos: list[dict]
    carteira: list[dict]
    sinqia: list[dict]
    pipefy: list[dict]
    mapeamento: list[tuple[str, str]]
    avisos_gerais: list[str] = field(default_factory=list)
    respostas: list[dict] = field(default_factory=list)  # entrada/respostas_gestor.xlsx (opcional)

    def __post_init__(self):
        self.respostas_manuais: dict[str, dict] = {}
        for linha in self.respostas:  # uma linha por fundo; vale a primeira
            self.respostas_manuais.setdefault(digitos(coluna(linha, "CNPJ do Fundo", "CNPJ")), linha)
        # índices por CNPJ: cada fundo consulta só as próprias linhas
        self.carteira_por_cnpj: dict[str, list[dict]] = defaultdict(list)
        for linha in self.carteira:
            self.carteira_por_cnpj[linha["cnpj"]].append(linha)
        self.sinqia_por_doc: dict[int, list[dict]] = defaultdict(list)
        for linha in self.sinqia:
            if linha["doc"] is not None:
                self.sinqia_por_doc[linha["doc"]].append(linha)
        self.pipefy_por_cnpj: dict[str, dict] = {}
        for linha in self.pipefy:  # FILTER devolve primeiro a linha de cima
            self.pipefy_por_cnpj.setdefault(digitos(coluna(linha, "CNPJ do Fundo")), linha)
        self.nome_por_cnpj: dict[str, str] = {}
        for f in self.fundos:  # XLOOKUP devolve o primeiro da lista
            self.nome_por_cnpj.setdefault(digitos(f["cnpj"]), f["nome"])


# Destino de mapeamento para linhas que a carteira traz, mas que não entram no quadro.
FORA_DO_RELATORIO = "(fora do relatório)"


def mapear_descricao(descricao, mapeamento) -> str:
    """CarteiraInforme!G: primeira chave (na ordem da tabela I:J) contida em LOWER(Descricao).

    =XLOOKUP(1,--ISNUMBER(FIND(chaves,LOWER(C2),1)),mapeamentos,"")
    """
    texto = str(descricao).lower() if descricao is not None else ""
    for chave, destino in mapeamento:
        if chave and chave in texto:
            return destino
    return ""


MARCA_ATUALIZADO = "_atualizado_por"


def aplicar_atualizacoes_pipefy(pipefy: list[dict], atualizacoes: list[dict], fundos: list[dict],
                                avisos: list[str], origem: str) -> None:
    """Sobrepõe uma extração parcial do Pipefy à base (pipefy.xlsx), fundo a fundo.

    Só os fundos da extração mudam, e neles só as colunas "Descrição item ..." preenchidas;
    o que a extração não traz (outros fundos, outras colunas) fica como está na base.
    """
    nome_lista = {}
    cnpj_do_nome = {}
    for f in fundos:
        nome_lista.setdefault(digitos(f["cnpj"]), norm(f["nome"]))
        cnpj_do_nome.setdefault(norm(f["nome"]), digitos(f["cnpj"]))

    def nome_card(linha):
        return norm(coluna(linha, "Nome do Fundo") or coluna(linha, "Título", "Title"))

    def titulo_card(linha):
        return coluna(linha, "Título", "Title") or coluna(linha, "Nome do Fundo")

    por_cnpj: dict[str, list[dict]] = defaultdict(list)
    for linha in atualizacoes:
        cnpj = digitos(coluna(linha, "CNPJ do Fundo", "CNPJ"))
        if cnpj:
            por_cnpj[cnpj].append(linha)
    escolhidas: dict[str, dict] = {}
    for cnpj, linhas in por_cnpj.items():
        if len(linhas) == 1:
            escolhidas.setdefault(cnpj, linhas[0])
            continue
        # Vários cards com o mesmo CNPJ. Os que têm o nome de OUTRO fundo da lista vão para esse fundo;
        # entre os que sobram vale o do nome da lista e, havendo mais de um, o criado por último.
        restantes = []
        for l in linhas:
            outro = cnpj_do_nome.get(nome_card(l))
            if outro and outro != cnpj and outro not in por_cnpj:
                escolhidas[outro] = l
                avisos.append(f"{origem}: o card '{titulo_card(l)}' está com o CNPJ de outro fundo; "
                              "foi aplicado ao fundo de mesmo nome da lista")
            else:
                restantes.append(l)
        candidatos = [l for l in restantes if nome_card(l) == nome_lista.get(cnpj)] or restantes
        if not candidatos:
            continue
        escolhido = max(candidatos, key=lambda l: str(coluna(l, "Criado em") or ""))  # empate: o primeiro
        escolhidas[cnpj] = escolhido
        ignorados = [l for l in restantes if l is not escolhido]
        if ignorados:
            mesmo_fundo = all(nome_card(l) == nome_card(escolhido) for l in ignorados)
            avisos.append(f"{origem}: {len(ignorados) + 1} cards com o CNPJ {formatar_cnpj(cnpj)} "
                          f"('{titulo_card(escolhido)}'); "
                          + ("vale o criado por último" if mesmo_fundo else
                             "os de nome diferente foram ignorados (confira o CNPJ no Pipefy)"))

    base = {}
    for linha in pipefy:
        base.setdefault(digitos(coluna(linha, "CNPJ do Fundo")), linha)
    for cnpj, nova in escolhidas.items():
        campos = {k: v for k, v in nova.items()
                  if norm(k).startswith("descrição item") and v not in (None, "") and str(v).strip()}
        if not campos:
            continue
        alvo = base.get(cnpj)
        if alvo is None:
            alvo = {"Title": coluna(nova, "Título", "Title") or coluna(nova, "Nome do Fundo"),
                    "CNPJ do Fundo": formatar_cnpj(cnpj)}
            pipefy.append(alvo)
            base[cnpj] = alvo
        chaves = {norm(k): k for k in alvo}
        for k, v in campos.items():
            alvo[chaves.get(norm(k), k)] = v
        alvo[MARCA_ATUALIZADO] = origem


def carregar_bases(cfg: dict) -> Bases:
    mapeamento = [(str(k).lower(), str(v).strip()) for k, v in cfg["mapeamento_carteira"]]
    avisos = []

    fundos = []
    for r in ler_tabela(arquivo_entrada(cfg, "fundos")):
        cnpj = coluna(r, "CNPJ")
        if cnpj in (None, ""):
            continue
        fundos.append({"nome": str(coluna(r, "Nome Fundo", "Nome") or "").strip(),
                       "cnpj": str(cnpj).strip()})

    carteira = []
    sem_mapa = set()
    for r in ler_tabela(arquivo_entrada(cfg, "carteira_informe")):
        desc = coluna(r, "Descricao", "Descrição")
        cat = mapear_descricao(desc, mapeamento)
        if desc and not cat:
            sem_mapa.add(str(desc))
        carteira.append({
            "data": para_data(coluna(r, "Data Base")),
            "cnpj": digitos(coluna(r, "CNPJ Fundo", "CNPJ")),
            "descricao": desc,
            "valor": num(coluna(r, "Valor")),
            "categoria": cat,
        })
    if sem_mapa:
        avisos.append("Descrições da carteira sem mapeamento (ficam fora do relatório): "
                      + "; ".join(sorted(sem_mapa)))

    sinqia = []
    doc_invalido = set()
    for r in ler_tabela(arquivo_entrada(cfg, "resumo_sinqia")):
        # DOC_FUNDO = 13 primeiros dígitos do CNPJ, como número (o SUMIFS compara numericamente)
        doc = coluna(r, "DOC_FUNDO")
        if isinstance(doc, (int, float)) and not isinstance(doc, bool):
            doc_num = int(doc)
        elif doc is not None and str(doc).strip().isdigit():
            doc_num = int(str(doc).strip())
        else:
            doc_num = None
            if doc not in (None, ""):
                doc_invalido.add(f"{doc} ({coluna(r, 'NOME_FUNDO')}, {coluna(r, 'RELATORIO')})")
        sinqia.append({
            "situacao": coluna(r, "SITUACAO_RECEBIVEL"),
            "doc": doc_num,
            "relatorio": coluna(r, "RELATORIO"),
            "qtde": num(coluna(r, "QTDE")),
            "nominal": num(coluna(r, "VALOR_NOMINAL")),
            "presente": num(coluna(r, "VALOR_PRESENTE")),
        })
    if doc_invalido:
        avisos.append("Linhas do Resumo Sinqia com DOC_FUNDO que não é CNPJ (não casam com nenhum fundo): "
                      + "; ".join(sorted(doc_invalido)))

    pipefy = []
    caminho_pipefy = arquivo_entrada(cfg, "pipefy")
    if caminho_pipefy.exists():
        pipefy = ler_tabela(caminho_pipefy)
    else:
        avisos.append(f"Relatório Pipefy não encontrado ({caminho_pipefy}); todos os textos do gestor "
                      "sairão com o texto padrão de 'gestor não se manifestou'.")

    caminho_atual = caminho_cfg(cfg, "entrada") / cfg["arquivos"].get("pipefy_atualizacoes",
                                                                       "pipefy_atualizacoes.xlsx")
    if caminho_atual.exists():
        aplicar_atualizacoes_pipefy(pipefy, ler_tabela(caminho_atual), fundos, avisos, caminho_atual.name)

    respostas = []
    caminho_respostas = caminho_cfg(cfg, "entrada") / cfg["arquivos"].get("respostas_gestor", "respostas_gestor.xlsx")
    if caminho_respostas.exists():
        respostas = ler_tabela(caminho_respostas)
        desconhecidas = [c for c in (respostas[0] if respostas else {})
                         if c not in SECOES_GESTOR and c not in ("CNPJ do Fundo", "CNPJ", "Fundo")]
        if desconhecidas:
            avisos.append(f"Colunas de {caminho_respostas.name} que não são seções do relatório (ignoradas): "
                          + "; ".join(map(str, desconhecidas)))

    return Bases(fundos, carteira, sinqia, pipefy, mapeamento, avisos, respostas)


def ler_validacao(caminho: Path) -> list[dict]:
    """Aba Validacao_Limine (legado): status na coluna Y, nominal em M, presente em N."""
    wb = openpyxl.load_workbook(caminho, read_only=True, data_only=True)
    ws = wb.worksheets[0]
    linhas = [{"status": r[24] if len(r) > 24 else None,
               "nominal": r[12] if len(r) > 12 else None,
               "presente": r[13] if len(r) > 13 else None}
              for r in ws.iter_rows(values_only=True)]
    wb.close()
    return linhas


# --------------------------------------------------------------------------
# cálculo do relatório (equivalente à aba CVM)
# --------------------------------------------------------------------------
@dataclass
class Relatorio:
    nome: str
    cnpj: str
    data_base: dt.date
    celulas: dict  # "B3" -> valor (apenas as células calculadas)
    avisos: list[str]
    atualizado_pipefy: bool = False  # textos do gestor vieram de pipefy_atualizacoes.xlsx

    @property
    def nome_arquivo(self) -> str:
        # o Windows aceita até 255 caracteres por nome; nomes reais têm no máximo ~130
        sufixo = f"_{digitos(self.cnpj)}_TRIMESTRAL_CVM"
        return nome_arquivo_seguro(self.nome.strip()[:200 - len(sufixo)].rstrip() + sufixo)


def tem_carteira(bases: Bases, cnpj: str) -> bool:
    """Setup!O / Setup!E: COUNTIF(CarteiraInforme!B:B, CNPJ) > 0."""
    return digitos(cnpj) in bases.carteira_por_cnpj


def formatar_cnpj(cnpj) -> str:
    d = digitos(cnpj)
    return f"{d[:2]}.{d[2:5]}.{d[5:8]}/{d[8:12]}-{d[12:]}" if len(d) == 14 else str(cnpj)


def fundos_fora_da_lista(bases: Bases) -> list[tuple[str, str]]:
    """CNPJs da Carteira Informe que não estão na lista de fundos: (CNPJ formatado, nome provável).

    O nome provável vem do título do card do Pipefy, quando houver.
    """
    fora = []
    for cnpj in bases.carteira_por_cnpj:
        if cnpj and cnpj not in bases.nome_por_cnpj:
            card = bases.pipefy_por_cnpj.get(cnpj)
            nome = str(coluna(card, "Title") or "").strip() if card else ""
            fora.append((formatar_cnpj(cnpj), nome))
    return fora


def nome_do_fundo(bases: Bases, cnpj: str) -> str:
    """Setup!C7 =XLOOKUP(cnpj, H:H, G:G): primeiro nome da lista com esse CNPJ."""
    return bases.nome_por_cnpj.get(digitos(cnpj), "")


AVISO_SOMA = "Soma das linhas da composição difere do Patrimônio Líquido"


def percentuais_composicao(c: dict) -> str:
    """D13:D20 a partir de C13:C20. Devolve o aviso quando as linhas não somam o PL."""
    pl = c["C20"]
    for linha in range(13, 20):
        c[f"D{linha}"] = div(c[f"C{linha}"], pl)                       # =IFERROR(C13/$C$20,0%)
    c["D20"] = sum(c[f"D{l}"] for l in range(13, 20))                  # =SUM(D13:D19)
    soma_linhas = sum(c[f"C{l}"] for l in range(13, 20))
    if pl and abs(soma_linhas - pl) > 0.01:
        return f"{AVISO_SOMA} (soma={soma_linhas:.2f}; PL={pl:.2f})"
    return ""


def calcular(bases: Bases, cnpj: str, data_base: dt.date, rotulos_composicao: dict,
             texto_padrao: str, validacao: list[dict] | None = None) -> Relatorio:
    cnpj = cnpj.strip()
    cnpj_dig = digitos(cnpj)
    nome = nome_do_fundo(bases, cnpj)
    c: dict = {}
    avisos: list[str] = []

    # Cabeçalho
    c["B3"] = nome                                                     # =Setup!C7
    c["B4"] = cnpj                                                     # =Setup!$C$8
    c["B5"] = f"{numero_excel(data_base.month / 3)}º Trimestre/{data_base.year}"  # =MONTH(B6)/3&"º Trimestre/"&YEAR(B6)
    c["B6"] = data_base                                                # =Setup!$C$6
    c["A10"] = "A seguir segue a composição da Carteira em " + data_base.strftime("%d/%m/%Y")
    c["G3"] = nome                                                     # =fundo (auxiliar, fora da impressão)
    c["G4"] = cnpj_dig[:13]                                            # =LEFT(CNPJ sem pontuação, 13)

    # I. Composição da carteira
    # C13:C20 =SUMIFS(CarteiraInforme!D:D, G:G, rótulo, A:A, data base, B:B, CNPJ)
    linhas_fundo = [l for l in bases.carteira_por_cnpj.get(cnpj_dig, ()) if l["data"] == data_base]
    if not linhas_fundo:
        avisos.append("Fundo sem linhas na Carteira Informe para a data base")
    soma_cat: dict[str, float] = defaultdict(float)
    for l in linhas_fundo:
        soma_cat[norm(l["categoria"])] += l["valor"]
    for linha, rotulo in rotulos_composicao.items():
        c[f"C{linha}"] = soma_cat.get(norm(rotulo), 0)
    aviso_soma = percentuais_composicao(c)
    if aviso_soma:
        avisos.append(aviso_soma)
    fora = [l for l in linhas_fundo if norm(l["categoria"]) == norm(FORA_DO_RELATORIO)]
    if fora:
        avisos.append("Linhas fora do relatório: " + "; ".join(
            f"{l['descricao']} = {l['valor']:.2f}" for l in fora))
    no_quadro = {norm(r) for r in rotulos_composicao.values()} | {norm(FORA_DO_RELATORIO)}
    outras = sorted({l["categoria"] for l in linhas_fundo if l["categoria"] and norm(l["categoria"]) not in no_quadro})
    if outras:
        avisos.append("Categorias da carteira que não aparecem no quadro: " + ", ".join(outras))

    # II. Verificação de lastro (Resumo Sinqia). G4 = LEFT(CNPJ sem pontuação, 13)
    doc = int(cnpj_dig[:13]) if cnpj_dig else None
    sq = bases.sinqia_por_doc.get(doc, ())

    def soma(campo, **filtro):
        return sum(l[campo] for l in sq if all(norm(l[k]) == norm(v) for k, v in filtro.items()))

    for linha, filtro in ((27, {"situacao": "A vencer"}), (28, {"situacao": "Vencido"}),
                          (30, {"relatorio": "Liquidacao"}), (31, {"relatorio": "Aquisicao"})):
        c[f"B{linha}"] = soma("qtde", **filtro)
        c[f"C{linha}"] = soma("nominal", **filtro)
        c[f"D{linha}"] = soma("presente", **filtro)
    for col in "BCD":
        c[f"{col}29"] = c[f"{col}27"] + c[f"{col}28"]                  # =SUM(B27:B28)
    c["E27"] = div(c["B27"], c["B29"]); c["F27"] = div(c["D27"], c["D29"])
    c["E28"] = div(c["B28"], c["B29"]); c["F28"] = div(c["D28"], c["D29"])
    c["E29"] = c["E27"] + c["E28"];      c["F29"] = c["F27"] + c["F28"]
    c["E30"] = div(c["B30"], c["B29"]); c["F30"] = div(c["D30"], c["D29"])
    c["E31"] = div(c["B31"], c["B29"]); c["F31"] = div(c["D31"], c["D29"])

    # Linhas 32-34: aba Validacao_Limine (Y = status, M = nominal, N = presente)
    val = validacao or []

    def cont(pred):
        return sum(1 for v in val if pred(v["status"]))

    def somav(pred, campo):
        return sum(num(v[campo]) for v in val if pred(v["status"]))

    a_vencer = lambda s: norm(s) == norm("A vencer")
    vencido = lambda s: norm(s) == norm("Vencido")
    # B34 =COUNTIF(Y:Y, CVM!$J$23) com J23 vazio: o Excel trata critério vazio como 0
    zero = lambda s: isinstance(s, (int, float)) and not isinstance(s, bool) and s == 0
    for linha, pred in ((32, a_vencer), (33, vencido), (34, zero)):
        c[f"B{linha}"] = cont(pred)
        c[f"C{linha}"] = somav(pred, "nominal")
        c[f"D{linha}"] = somav(pred, "presente")
    c["E32"] = div(c["B32"], c["B27"]); c["F32"] = div(c["D32"], c["D27"])
    c["E33"] = div(c["B33"], c["B28"]); c["F33"] = div(c["D33"], c["D27"])  # planilha divide por D27 (não D28)
    c["E34"] = div(c["B34"], c["B30"]); c["F34"] = div(c["D34"], c["D30"])

    # IV/V. Textos do gestor: CVM!I5 =FILTER(Pipefy!A:W, Pipefy!B:B=CNPJ, "")
    resposta = bases.pipefy_por_cnpj.get(cnpj_dig)
    if resposta is None:
        avisos.append("Sem resposta do gestor no Pipefy: textos padrão usados")
    elif resposta.get(MARCA_ATUALIZADO):
        avisos.append(f"Textos do gestor atualizados por {resposta[MARCA_ATUALIZADO]}")
    c["I5"] = coluna(resposta, "Title") if resposta else None          # título do card (auxiliar)
    for linha, campo in TEXTOS_GESTOR.items():
        valor = coluna(resposta, campo) if resposta else None
        c[f"A{linha}"] = texto_padrao if valor in (None, "") else str(valor)   # =IF(K5="",$I$8,K5)

    manual = bases.respostas_manuais.get(cnpj_dig)
    if manual:
        usadas = []
        for secao, linha in SECOES_GESTOR.items():
            texto = coluna(manual, secao)
            if texto not in (None, "") and str(texto).strip():
                c[f"A{linha}"] = str(texto).strip()
                usadas.append(secao.split(" - ")[0])
        if usadas:
            if resposta is None:
                avisos[:] = [a for a in avisos if not a.startswith("Sem resposta do gestor no Pipefy")]
            avisos.append("Textos do gestor de respostas_gestor.xlsx (substituem o Pipefy): " + ", ".join(usadas))

    return Relatorio(nome, cnpj, data_base, c, avisos,
                     atualizado_pipefy=bool(resposta and resposta.get(MARCA_ATUALIZADO)))


# --------------------------------------------------------------------------
# modelo da aba CVM (assets/modelo_cvm.xlsx, salvo pelo Excel)
# --------------------------------------------------------------------------
RE_CELULA = re.compile(r'<c r="([A-Z]{1,3})(\d+)"([^>]*?)(?:/>|>(.*?)</c>)', re.S)
RE_LINHA = re.compile(r'<row r="(\d+)"([^>]*?)(/?)>')


def serial_excel(valor) -> float:
    if isinstance(valor, dt.datetime):
        return (valor - dt.datetime(1899, 12, 30)).total_seconds() / 86400
    return float((valor - DATA_ZERO_EXCEL).days)


def celula_xml(ref: str, estilo: str | None, valor) -> str:
    s = f' s="{estilo}"' if estilo else ""
    if valor is None or valor == "":
        return f'<c r="{ref}"{s}/>'
    if isinstance(valor, bool):
        return f'<c r="{ref}"{s} t="b"><v>{int(valor)}</v></c>'
    if isinstance(valor, dt.date):
        valor = serial_excel(valor)
    if isinstance(valor, (int, float)):
        x = float(valor)
        if not math.isfinite(x):
            return f'<c r="{ref}"{s} t="e"><v>#NUM!</v></c>'
        return f'<c r="{ref}"{s}><v>{int(x) if x.is_integer() else repr(x)}</v></c>'
    texto = XML_INVALIDO.sub("", str(valor).replace("\r\n", "\n").replace("\r", "\n"))
    return f'<c r="{ref}"{s} t="inlineStr"><is><t xml:space="preserve">{escape(texto)}</t></is></c>'


def _coluna_num(letras: str) -> int:
    n = 0
    for ch in letras:
        n = n * 26 + ord(ch) - 64
    return n


def _inserir_celula(xml: str, ref: str, valor) -> str:
    """Cria a célula quando o modelo não tem o elemento <c> (por exemplo, sem formatação)."""
    letras, linha = re.fullmatch(r"([A-Z]+)(\d+)", ref).groups()
    nova = celula_xml(ref, None, valor)
    m = re.search(rf'<row r="{linha}"([^>]*?)(/?)>', xml)
    if m is None:
        for outra in RE_LINHA.finditer(xml):
            if int(outra.group(1)) > int(linha):
                return xml[:outra.start()] + f'<row r="{linha}">{nova}</row>' + xml[outra.start():]
        if "<sheetData/>" in xml:
            return xml.replace("<sheetData/>", f'<sheetData><row r="{linha}">{nova}</row></sheetData>')
        return xml.replace("</sheetData>", f'<row r="{linha}">{nova}</row></sheetData>')
    attrs = re.sub(r'\s+spans="[^"]*"', "", m.group(1))
    if m.group(2):  # <row .../>
        return xml[:m.start()] + f'<row r="{linha}"{attrs}>{nova}</row>' + xml[m.end():]
    fim = xml.index("</row>", m.end())
    corpo = xml[m.end():fim]
    pos = len(corpo)
    for cel in RE_CELULA.finditer(corpo):
        if _coluna_num(cel.group(1)) > _coluna_num(letras):
            pos = cel.start()
            break
    return xml[:m.start()] + f'<row r="{linha}"{attrs}>' + corpo[:pos] + nova + corpo[pos:] + xml[fim:]


class ModeloCVM:
    """A aba CVM do modelo: o XLSX de saída é ele com as células calculadas preenchidas."""

    def __init__(self, caminho: Path = MODELO_XLSX):
        dados = Path(caminho).read_bytes()
        with zipfile.ZipFile(io.BytesIO(dados)) as z:
            self._infos = z.infolist()
            self._partes = {i.filename: z.read(i.filename) for i in self._infos}
        rid = re.search(r'<sheet [^>]*r:id="([^"]+)"', self._parte_txt("xl/workbook.xml")).group(1)
        self._parte_aba = self._destino_rel("xl/workbook.xml", rid)
        self._xml_aba = self._parte_txt(self._parte_aba)
        self.ws = openpyxl.load_workbook(io.BytesIO(dados), rich_text=True).worksheets[0]
        self._geometria = None
        self._estilos: dict = {}
        self.cache_linhas: dict = {}  # quebra de linha dos textos fixos, reaproveitada entre fundos

    # ---- partes do pacote
    def _parte_txt(self, nome: str) -> str:
        return self._partes[nome].decode("utf-8")

    def _destino_rel(self, parte: str, rid: str) -> str:
        pasta, nome = posixpath.split(parte)
        rels = self._parte_txt(posixpath.join(pasta, "_rels", nome + ".rels"))
        for rel in re.findall(r"<Relationship\b[^>]*>", rels):
            attrs = dict(re.findall(r'(\w+)="([^"]*)"', rel))
            if attrs.get("Id") == rid:
                alvo = attrs["Target"]
                return alvo.lstrip("/") if alvo.startswith("/") else posixpath.normpath(posixpath.join(pasta, alvo))
        raise KeyError(f"relacionamento {rid} não encontrado em {parte}")

    def logo_cabecalho(self):
        """Imagem do cabeçalho esquerdo (&G): (bytes, largura_pt, altura_pt) ou None."""
        m = re.search(r'<legacyDrawingHF r:id="([^"]+)"', self._xml_aba)
        if not m or "&G" not in (self.ws.oddHeader.left.text or ""):
            return None
        vml_parte = self._destino_rel(self._parte_aba, m.group(1))
        vml = self._partes[vml_parte].decode("utf-8", "replace")
        forma = re.search(r'<v:shape [^>]*id="LH".*?</v:shape>', vml, re.S)
        if not forma:
            return None
        estilo = re.search(r"style='([^']*)'", forma.group(0)).group(1)
        medida = dict(re.findall(r"(width|height):([\d.]+)pt", estilo))
        rid_img = re.search(r'o:relid="([^"]+)"', forma.group(0)).group(1)
        imagem = self._partes[self._destino_rel(vml_parte, rid_img)]
        return imagem, float(medida["width"]), float(medida["height"])

    # ---- XLSX
    def xlsx(self, celulas: dict, mostrar_lastro: bool) -> bytes:
        pendentes = dict(celulas)

        def troca(m):
            ref = m.group(1) + m.group(2)
            if ref not in pendentes:
                return m.group(0)
            estilo = re.search(r'\ss="(\d+)"', m.group(3))
            return celula_xml(ref, estilo.group(1) if estilo else None, pendentes.pop(ref))

        xml = RE_CELULA.sub(troca, self._xml_aba)
        for ref, valor in pendentes.items():
            xml = _inserir_celula(xml, ref, valor)

        marca_oculta = "" if mostrar_lastro else ' hidden="1"'

        def oculta(m):
            linha = int(m.group(1))
            if linha not in LINHAS_SECAO_LASTRO:
                return m.group(0)
            attrs = re.sub(r'\s+hidden="(?:1|true)"', "", m.group(2))
            return f'<row r="{linha}"{attrs}{marca_oculta}{m.group(3)}>'

        xml = RE_LINHA.sub(oculta, xml)
        buf = io.BytesIO()
        agora = time.localtime()[:6]
        with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as z:
            for info in self._infos:
                dados = xml.encode("utf-8") if info.filename == self._parte_aba else self._partes[info.filename]
                novo = zipfile.ZipInfo(info.filename, date_time=agora)
                novo.compress_type = zipfile.ZIP_DEFLATED
                z.writestr(novo, dados)
        return buf.getvalue()

    # ---- geometria de impressão (para o PDF)
    @property
    def geometria(self) -> "Geometria":
        if self._geometria is None:
            self._geometria = Geometria.do_modelo(self)
        return self._geometria

    def estilo(self, linha: int, col: int) -> "EstiloCelula":
        chave = (linha, col)
        if chave not in self._estilos:
            self._estilos[chave] = EstiloCelula.da_celula(self.ws.cell(linha, col))
        return self._estilos[chave]


def carregar_modelo() -> ModeloCVM:
    return ModeloCVM(MODELO_XLSX)


def rotulos_composicao(ws) -> dict:
    return {linha: str(ws[f"B{linha}"].value).strip() for linha in range(13, 21)}


def salvar_xlsx(rel: Relatorio, destino: Path, mostrar_lastro: bool, modelo: ModeloCVM) -> None:
    caminho_gravavel(destino).write_bytes(modelo.xlsx(rel.celulas, mostrar_lastro))


# --------------------------------------------------------------------------
# saída PDF: a aba CVM desenhada como o Excel imprime (A4, zoom da página)
# --------------------------------------------------------------------------
# Calibração medida no PDF que o próprio Excel exporta desta aba (zoom 68%). O Excel não
# aplica o zoom igual nos dois eixos: as colunas saem 7,5% e as linhas 1,6% maiores.
AJUSTE_X = 0.73086 / 0.68
AJUSTE_Y = 0.6906 / 0.68
PX_DIGITO = 7              # largura do dígito da fonte Normal (Aptos Narrow 11), em pixels
PASSO_FONTE = 72 / 600     # o Excel arredonda o corpo da fonte para pixels da impressora (600 dpi)
DESLOCAMENTO_Y = 0.47      # a grade começa um pouco abaixo da margem superior
# medidas de texto em pontos impressos a 68% (escalam com o zoom)
MARGEM_TEXTO = 1.1         # entre a borda esquerda da célula e o texto
MARGEM_DIREITA = 2.0       # entre o texto alinhado à direita e a borda direita
RECUO_NIVEL = 6.2          # por nível de recuo
BASE_INFERIOR = 2.3        # linha de base acima da borda de baixo (alinhamento inferior)
# O texto impresso pelo Excel sai ~0,5% mais largo que as métricas da fonte e a área útil da
# célula é um pouco menor: medido em 895 quebras de linha do Excel, ele quebra quando a linha
# passa de 522-528 pt numa célula de 531 pt úteis.
FATOR_QUEBRA = 0.99
# medidas proporcionais ao corpo da fonte impresso
ENTRELINHA = 1.257
BASE_SUPERIOR = 0.97       # linha de base abaixo da borda de cima (alinhamento superior)
BASE_CENTRO = 0.32         # linha de base abaixo do centro (alinhamento central)
ESPESSURA_BORDA = {"hair": 0.3, "thin": 0.6, "medium": 1.2, "mediumDashed": 1.2, "mediumDashDot": 1.2,
                   "mediumDashDotDot": 1.2, "slantDashDot": 1.2, "thick": 1.8, "double": 1.8}
PAPEL = {1: (612.0, 792.0), 9: (595.2756, 841.8898)}  # carta, A4


@dataclass
class EstiloCelula:
    fonte: str
    tamanho: float
    negrito: bool
    italico: bool
    cor: tuple | None
    horizontal: str | None
    vertical: str | None
    quebra: bool
    recuo: float
    formato: str
    bordas: tuple            # (esq, dir, cima, baixo): (estilo, cor) ou None
    fundo: tuple | None

    @staticmethod
    def da_celula(cel) -> "EstiloCelula":
        f, a, b = cel.font, cel.alignment, cel.border
        bordas = tuple((lado.style, _cor(lado.color)) if lado is not None and lado.style else None
                       for lado in (b.left, b.right, b.top, b.bottom))
        fundo = _cor(cel.fill.fgColor) if cel.fill is not None and cel.fill.fill_type == "solid" else None
        return EstiloCelula(f.name or "Arial", float(f.sz or 11), bool(f.b), bool(f.i), _cor(f.color),
                            a.horizontal, a.vertical, bool(a.wrap_text), float(a.indent or 0),
                            cel.number_format or "General", bordas, fundo)


def _cor(cor) -> tuple | None:
    """Cor RGB (0-1) de um objeto de cor do openpyxl; temas e índices viram preto/branco."""
    if cor is None:
        return None
    try:
        if cor.type == "rgb" and isinstance(cor.rgb, str) and len(cor.rgb) >= 6:
            h = cor.rgb[-6:]
            return tuple(int(h[i:i + 2], 16) / 255 for i in (0, 2, 4))
        if cor.type == "theme":
            return (1.0, 1.0, 1.0) if cor.theme in (0, 2) else (0.0, 0.0, 0.0)
        if cor.type == "indexed":
            return (1.0, 1.0, 1.0) if cor.indexed in (1, 9) else (0.0, 0.0, 0.0)
    except (AttributeError, ValueError):
        pass
    return None


@dataclass
class Geometria:
    pagina: tuple[float, float]
    margens: tuple[float, float, float, float]  # esquerda, direita, topo, base (pt)
    margem_cabecalho: float
    zoom: float
    area: tuple[int, int, int, int]               # col1, lin1, col2, lin2
    x_col: dict[int, float]                       # borda esquerda impressa da coluna
    larg_col: dict[int, float]                    # largura impressa da coluna
    alt_linha: dict[int, float]                   # altura da linha no Excel (pt, sem zoom)
    ocultas: set[int]
    quebras: set[int]                             # quebra manual depois destas linhas
    mesclas: dict[tuple[int, int], tuple[int, int, int, int]]   # (lin, col) -> (l1, c1, l2, c2)
    logo: tuple | None                            # (bytes, largura impressa, altura impressa)

    @property
    def sx(self) -> float:
        return self.zoom * AJUSTE_X

    @property
    def sy(self) -> float:
        return self.zoom * AJUSTE_Y

    @staticmethod
    def do_modelo(modelo: ModeloCVM) -> "Geometria":
        ws, xml = modelo.ws, modelo._xml_aba
        zoom = (ws.page_setup.scale or 100) / 100
        papel = PAPEL.get(int(ws.page_setup.paperSize or 9), PAPEL[9])
        if ws.page_setup.orientation == "landscape":
            papel = (papel[1], papel[0])
        pm = ws.page_margins
        margens = tuple(72 * (v if v is not None else d) for v, d in
                        ((pm.left, 0.7), (pm.right, 0.7), (pm.top, 0.75), (pm.bottom, 0.75)))
        area_txt = (ws.print_area or ws.dimensions).split("!")[-1].replace("$", "")
        area = range_boundaries(area_txt)

        larguras = {}
        for col in re.findall(r"<col\b[^>]*/>", xml):
            a = dict(re.findall(r'(\w+)="([^"]*)"', col))
            for c in range(int(a["min"]), int(a["max"]) + 1):
                larguras[c] = float(a.get("width", 0))
        padrao = re.search(r'defaultColWidth="([\d.]+)"', xml)
        larg_padrao = float(padrao.group(1)) if padrao else 9.140625
        sx = zoom * AJUSTE_X
        larg_col = {c: round(larguras.get(c, larg_padrao) * PX_DIGITO) * 0.75 * sx
                    for c in range(area[0], area[2] + 1)}
        total = sum(larg_col.values())
        x = margens[0]
        if ws.print_options.horizontalCentered:
            x += (papel[0] - margens[0] - margens[1] - total) / 2
        x_col = {}
        for c in range(area[0], area[2] + 1):
            x_col[c] = x
            x += larg_col[c]

        alt_padrao = ws.sheet_format.defaultRowHeight or 15
        alt_linha, ocultas = {}, set()
        for r in range(area[1], area[3] + 1):
            dim = ws.row_dimensions[r]
            alt_linha[r] = dim.height if dim.height is not None else alt_padrao
            if dim.hidden:
                ocultas.add(r)

        mesclas = {}
        for m in ws.merged_cells.ranges:
            for r in range(m.min_row, m.max_row + 1):
                for c in range(m.min_col, m.max_col + 1):
                    mesclas[(r, c)] = (m.min_row, m.min_col, m.max_row, m.max_col)

        logo = modelo.logo_cabecalho()
        if logo:
            logo = (logo[0], logo[1] * sx, logo[2] * zoom * AJUSTE_Y)
        return Geometria(papel, margens, 72 * (pm.header if pm.header is not None else 0.3), zoom, area,
                         x_col, larg_col, alt_linha, ocultas, {b.id for b in ws.row_breaks.brk}, mesclas, logo)


_FONTES: dict | None = None


def _fontes() -> dict:
    """Registra a Arial do Windows (a fonte do modelo); sem ela, usa Helvetica (mesmas métricas)."""
    global _FONTES
    if _FONTES is None:
        from reportlab.pdfbase import pdfmetrics
        from reportlab.pdfbase.ttfonts import TTFont

        sans = ("Helvetica", "Helvetica-Bold", "Helvetica-Oblique", "Helvetica-BoldOblique")
        pasta = Path(os.environ.get("WINDIR", r"C:\Windows")) / "Fonts"
        arquivos = ("arial.ttf", "arialbd.ttf", "ariali.ttf", "arialbi.ttf")
        if all((pasta / a).exists() for a in arquivos):
            try:
                nomes = ("Arial", "Arial-Bold", "Arial-Italic", "Arial-BoldItalic")
                for nome, arq in zip(nomes, arquivos):
                    pdfmetrics.registerFont(TTFont(nome, str(pasta / arq)))
                sans = nomes
            except Exception:  # fonte corrompida ou inacessível: segue com Helvetica
                pass
        _FONTES = {"sans": sans,
                   "serif": ("Times-Roman", "Times-Bold", "Times-Italic", "Times-BoldItalic"),
                   "mono": ("Courier", "Courier-Bold", "Courier-Oblique", "Courier-BoldOblique")}
    return _FONTES


def _fonte_pdf(nome: str, negrito: bool, italico: bool) -> str:
    n = (nome or "").lower()
    familia = ("serif" if any(k in n for k in ("times", "cambria", "georgia", "garamond", "book"))
               else "mono" if any(k in n for k in ("courier", "consolas", "mono")) else "sans")
    return _fontes()[familia][(1 if negrito else 0) + (2 if italico else 0)]


def _corpo_impresso(tamanho: float, zoom: float) -> float:
    return math.ceil(tamanho * zoom / PASSO_FONTE - 1e-6) * PASSO_FONTE


def _milhar(inteiro: str) -> str:
    grupos = []
    while len(inteiro) > 3:
        grupos.insert(0, inteiro[-3:])
        inteiro = inteiro[:-3]
    return ".".join([inteiro] + grupos)


def formatar_numero(valor: float, formato: str) -> tuple[str, str]:
    """Texto de um número no formato da célula (Excel pt-BR).

    Devolve (texto, reserva): `reserva` são caracteres cuja largura o Excel deixa em branco à
    direita (os "_)" e "?" do formato contábil), para alinhar os números como na planilha.
    """
    formato = formato or "General"
    if formato in ("General", "@"):
        return numero_excel(valor), ""
    sem_literais = re.sub(r'"[^"]*"|\\.', "", formato).lower()
    if re.search(r"[dmy]", sem_literais) and not re.search(r"[0#]", sem_literais):
        return para_data(valor).strftime("%d/%m/%Y"), ""   # data guardada como número
    secoes = re.split(r';(?=(?:[^"]*"[^"]*")*[^"]*$)', formato)
    casas_m = re.search(r"0\.(0+)", secoes[0])
    casas = len(casas_m.group(1)) if casas_m else 0
    pct = "%" in secoes[0].replace('"%"', "")
    exibido = round(valor * (100 if pct else 1), casas)
    if exibido == 0 and len(secoes) >= 3:
        secao, valor = secoes[2], 0.0
    elif exibido < 0 and len(secoes) >= 2:
        secao, valor = secoes[1], abs(valor)
    else:
        secao = secoes[0]
        if exibido == 0:
            valor = 0.0  # "0,00%", nunca "-0,00%"
    texto, reserva = "", ""
    i = 0
    while i < len(secao):
        ch = secao[i]
        if ch == '"':
            fim = secao.index('"', i + 1)
            texto += reserva + secao[i + 1:fim]
            reserva, i = "", fim + 1
        elif ch == "\\" and i + 1 < len(secao):
            texto += reserva + secao[i + 1]
            reserva, i = "", i + 2
        elif ch == "_" and i + 1 < len(secao):
            if texto:
                reserva += secao[i + 1]
            i += 2
        elif ch == "*":
            i += 2
        elif ch == "[":
            i = secao.index("]", i) + 1
        elif ch == "?":
            reserva += "0"
            i += 1
        elif ch in "#0":
            m = re.match(r"[#0,]+(\.[0#?]+)?", secao[i:])
            corpo = m.group(0)
            dec = len(m.group(1)) - 1 if m.group(1) else 0
            n = abs(valor) * (100 if pct else 1)
            if exibido < 0 and len(secoes) < 2:
                texto += "-"
            inteiro, _, frac = f"{n:.{dec}f}".partition(".")
            if "," in corpo:
                inteiro = _milhar(inteiro)
            texto += reserva + inteiro + ("," + frac if dec else "")
            reserva, i = "", i + len(corpo)
        else:
            texto += reserva + ch
            reserva, i = "", i + 1
    return texto.strip(), reserva


def _trechos(valor, est: EstiloCelula, corpo: float, zoom: float) -> list[tuple[str, str, float]]:
    """Valor da célula em trechos (texto, fonte do PDF, corpo impresso), respeitando o rich text."""
    if isinstance(valor, CellRichText):
        saida = []
        for bloco in valor:
            if isinstance(bloco, TextBlock):
                f = bloco.font
                neg = est.negrito if f is None or f.b is None else bool(f.b)
                ita = est.italico if f is None or f.i is None else bool(f.i)
                nome = f.rFont if f is not None and f.rFont else est.fonte
                tam = _corpo_impresso(float(f.sz), zoom) if f is not None and f.sz else corpo
                saida.append((bloco.text, _fonte_pdf(nome, neg, ita), tam))
            else:
                saida.append((str(bloco), _fonte_pdf(est.fonte, est.negrito, est.italico), corpo))
        return saida
    return [(str(valor), _fonte_pdf(est.fonte, est.negrito, est.italico), corpo)]


def _quebrar_linhas(trechos, largura: float, medir) -> list[tuple[list, float, bool]]:
    """Quebra o texto em linhas de até `largura` pt, como a quebra automática do Excel.

    Cada linha é (tokens, largura, fim_de_paragrafo); token = (texto, fonte, corpo, largura).
    """
    linhas = []
    atual: list = []
    larg = 0.0

    def fechar(fim_paragrafo):
        nonlocal atual, larg
        while atual and atual[-1][0].isspace():
            larg -= atual.pop()[3]
        linhas.append((atual, larg, fim_paragrafo))
        atual, larg = [], 0.0

    for texto, fonte, corpo in trechos:
        for tok in re.findall(r"\n|[^\S\n]+|\S+", texto):
            if tok == "\n":
                fechar(True)
                continue
            w = medir(tok, fonte, corpo)
            if tok.isspace():
                if atual:
                    atual.append((tok, fonte, corpo, w))
                    larg += w
                continue
            if larg + w > largura + 0.01 and any(not t[0].isspace() for t in atual):
                fechar(False)
            if w > largura:  # palavra maior que a linha: quebra entre caracteres
                for ch in tok:
                    wc = medir(ch, fonte, corpo)
                    if larg + wc > largura + 0.01 and atual:
                        fechar(False)
                    atual.append((ch, fonte, corpo, wc))
                    larg += wc
                continue
            atual.append((tok, fonte, corpo, w))
            larg += w
    fechar(True)
    return linhas


@dataclass
class _Texto:
    l1: int
    l2: int
    c1: int
    c2: int
    linhas: list
    horizontal: str
    vertical: str
    recuo: float
    reserva: float
    corpo: float
    justificar: bool
    cor: tuple | None
    vazias_fim: int = 0      # linhas em branco no fim do texto (o Excel as considera ao alinhar)


def _preparar_textos(rel: Relatorio, modelo: ModeloCVM, alturas: dict, cache: dict) -> list[_Texto]:
    from reportlab.pdfbase.pdfmetrics import stringWidth

    g = modelo.geometria
    c1, _, c2, _ = g.area
    k = g.zoom / 0.68
    margens = (MARGEM_TEXTO + MARGEM_DIREITA) * k

    def valor(r, c):
        ref = f"{get_column_letter(c)}{r}"
        return rel.celulas[ref] if ref in rel.celulas else modelo.ws.cell(r, c).value

    def vazia(r, c):
        v = valor(r, c)
        return (v is None or (isinstance(v, str) and not v.strip())) and (r, c) not in g.mesclas

    textos = []
    for r in alturas:
        for c in range(c1, c2 + 1):
            mescla = g.mesclas.get((r, c))
            if mescla and (mescla[0], mescla[1]) != (r, c):
                continue
            v = valor(r, c)
            if v is None or (isinstance(v, str) and not v.strip()):
                continue
            est = modelo.estilo(r, c)
            corpo = _corpo_impresso(est.tamanho, g.zoom)
            numero = isinstance(v, (int, float)) and not isinstance(v, bool)
            reserva = ""
            if numero:
                texto, reserva = formatar_numero(float(v), est.formato)
                trechos = [(texto, _fonte_pdf(est.fonte, est.negrito, est.italico), corpo)]
            elif isinstance(v, (dt.date, dt.datetime)):
                trechos = [(v.strftime("%d/%m/%Y"), _fonte_pdf(est.fonte, est.negrito, est.italico), corpo)]
            else:
                if isinstance(v, str):  # mantém as quebras de linha finais, como o Excel
                    v = XML_INVALIDO.sub("", v.replace("\r\n", "\n").replace("\r", "\n")).lstrip().rstrip(" \t")
                trechos = _trechos(v, est, corpo, g.zoom)
            h = est.horizontal or "general"
            if h == "general":
                h = "right" if numero or isinstance(v, (dt.date, dt.datetime)) else "left"
            quebra = (est.quebra or h in ("justify", "distributed")) and not numero
            l1, ca, l2, cb = mescla or (r, c, r, c)
            if not quebra and not mescla:
                largura_txt = sum(stringWidth(t, f, s) for t, f, s in trechos)
                ca, cb = _transbordar(r, c, h, largura_txt + margens, g, vazia)
            recuo = est.recuo * RECUO_NIVEL * k
            largura = sum(g.larg_col[x] for x in range(ca, cb + 1)) - margens - recuo
            larg_reserva = stringWidth(reserva, trechos[0][1], corpo) if reserva else 0.0
            chave = (r, c, repr(trechos), round(largura, 2), quebra)
            em_cache = cache.get(chave)
            if em_cache is None:
                vazias = 0
                if quebra:
                    linhas = _quebrar_linhas(trechos, largura * FATOR_QUEBRA, stringWidth)
                    while len(linhas) > 1 and not linhas[-1][0]:
                        linhas.pop()
                        vazias += 1
                else:
                    plano = [(t.replace("\n", " "), f, s) for t, f, s in trechos]
                    linhas = _quebrar_linhas(plano, float("inf"), stringWidth)
                    excesso = linhas[0][1] + larg_reserva - largura
                    if excesso > 0.01:  # não cabe e não pode transbordar: reduz a fonte até caber
                        fator = (largura - larg_reserva) / linhas[0][1]
                        linhas = [([(t, f, s * fator, w * fator) for t, f, s, w in linhas[0][0]],
                                   linhas[0][1] * fator, True)]
                em_cache = cache[chave] = (linhas, vazias)
            linhas, vazias = em_cache
            textos.append(_Texto(l1, l2, ca, cb, linhas, h, est.vertical or "bottom", recuo, larg_reserva,
                                 max((t[2] for ln in linhas for t in ln[0]), default=corpo),
                                 h in ("justify", "distributed"), est.cor, vazias))
    return textos


def _transbordar(r, c, h, largura, g: Geometria, vazia) -> tuple[int, int]:
    """Colunas ocupadas por um texto sem quebra: ele invade as vizinhas vazias, como no Excel."""
    c1, _, c2, _ = g.area
    ca = cb = c
    if h == "centerContinuous":
        while cb < c2 and vazia(r, cb + 1):
            cb += 1
        return ca, cb
    ocupada = g.larg_col[c]
    while ocupada < largura:
        if h in ("left", "fill") and cb < c2 and vazia(r, cb + 1):
            cb += 1
            ocupada += g.larg_col[cb]
        elif h == "right" and ca > c1 and vazia(r, ca - 1):
            ca -= 1
            ocupada += g.larg_col[ca]
        elif h == "center" and ((cb < c2 and vazia(r, cb + 1)) or (ca > c1 and vazia(r, ca - 1))):
            if cb < c2 and vazia(r, cb + 1):
                cb += 1
                ocupada += g.larg_col[cb]
            if ocupada < largura and ca > c1 and vazia(r, ca - 1):
                ca -= 1
                ocupada += g.larg_col[ca]
        else:
            break
    return ca, cb


def salvar_pdf(rel: Relatorio, destino: Path, mostrar_lastro: bool, modelo: ModeloCVM) -> None:
    from reportlab import rl_config
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfgen.canvas import Canvas

    rl_config.useA85 = 0  # fluxos binários: sem a codificação ASCII85 (lenta em Python puro)

    g = modelo.geometria
    larg_pag, alt_pag = g.pagina
    col1, lin1, col2, lin2 = g.area
    k = g.zoom / 0.68

    # 1) linhas visíveis, com a altura impressa
    alturas = {}
    for r in range(lin1, lin2 + 1):
        if r in g.ocultas and not (mostrar_lastro and r in LINHAS_SECAO_LASTRO):
            continue
        alturas[r] = g.alt_linha[r] * g.sy

    # 2) textos; a linha cresce quando o texto quebrado não cabe (o Excel cortaria o texto)
    textos = _preparar_textos(rel, modelo, alturas, modelo.cache_linhas)

    def altura_necessaria(linhas: int, corpo: float) -> float:
        return (linhas - 1) * ENTRELINHA * corpo + 0.905 * corpo + (BASE_INFERIOR + 0.5) * k

    for t in textos:
        if len(t.linhas) + t.vazias_fim <= 1:
            continue
        visiveis = [r for r in range(t.l1, t.l2 + 1) if r in alturas]
        disponivel = sum(alturas[r] for r in visiveis)
        if t.vazias_fim and altura_necessaria(len(t.linhas) + t.vazias_fim, t.corpo) > disponivel:
            t.vazias_fim = 0  # as linhas em branco não cabem: o Excel cortaria o texto; ignora
        falta = altura_necessaria(len(t.linhas), t.corpo) - disponivel
        if falta > 0 and visiveis:
            alturas[visiveis[-1]] += falta

    # 3) paginação: quebras manuais do modelo e automáticas; células mescladas não se dividem
    fim_mescla = {}
    for (r, c), (l1, _, l2, _) in g.mesclas.items():
        if r == l1:
            fim_mescla[r] = max(fim_mescla.get(r, r), l2)
    topo = g.margens[2] + DESLOCAMENTO_Y * k
    limite = alt_pag - g.margens[3]
    paginas: list[list[int]] = [[]]
    pos: dict[int, tuple[int, float]] = {}
    y, protegido_ate = topo, 0
    visiveis = list(alturas)
    quebrou_sozinho = False  # houve quebra automática desde a última quebra manual
    for i, r in enumerate(visiveis):
        if r > protegido_ate:
            fim = fim_mescla.get(r, r)
            bloco = sum(alturas[x] for x in range(r, fim + 1) if x in alturas)
            if y + bloco > limite + 0.01 and paginas[-1]:
                paginas.append([])
                y = topo
                quebrou_sozinho = True
            protegido_ate = fim
        pos[r] = (len(paginas) - 1, y)
        paginas[-1].append(r)
        y += alturas[r]
        if r in g.quebras and i < len(visiveis) - 1:
            # Com textos longos, o conteúdo antes da quebra manual já passou para outra página;
            # quebrar de novo deixaria essa página quase vazia. Nesse caso a quebra manual é ignorada.
            if not quebrou_sozinho:
                paginas.append([])
                y = topo
            quebrou_sozinho = False
    paginas = [p for p in paginas if p]
    pagina_de = {r: n for n, linhas in enumerate(paginas) for r in linhas}

    # 4) desenho
    cv = Canvas(str(destino), pagesize=(larg_pag, alt_pag), pageCompression=1)
    cv.setTitle(f"Demonstrativo Trimestral - {rel.nome}")
    cv.setAuthor("Limine Trust DTVM Ltda.")
    logo = None
    if g.logo:
        logo = (ImageReader(io.BytesIO(g.logo[0])), g.logo[1], g.logo[2])

    def caixa(l1, c1_, l2, c2_):
        x0 = g.x_col[c1_]
        x1 = g.x_col[c2_] + g.larg_col[c2_]
        y0 = pos[l1][1]
        ult = max(x for x in range(l1, l2 + 1) if x in alturas)
        return x0, y0, x1, pos[ult][1] + alturas[ult]

    for n, linhas in enumerate(paginas):
        if logo:
            cv.drawImage(logo[0], g.margens[0], alt_pag - g.margem_cabecalho - logo[2],
                         width=logo[1], height=logo[2], mask="auto")
        # fundos
        for r in linhas:
            for c in range(col1, col2 + 1):
                est = modelo.estilo(r, c)
                if est.fundo:
                    cv.setFillColorRGB(*est.fundo)
                    cv.rect(g.x_col[c], alt_pag - pos[r][1] - alturas[r], g.larg_col[c], alturas[r],
                            stroke=0, fill=1)
        # textos
        for t in textos:
            if pagina_de.get(t.l1) != n:
                continue
            x0, y0, x1, y1 = caixa(t.l1, t.c1, t.l2, t.c2)
            _desenhar_texto(cv, t, x0, y0, x1, y1, alt_pag, k)
        # bordas
        for r in linhas:
            for c in range(col1, col2 + 1):
                est = modelo.estilo(r, c)
                if not any(est.bordas):
                    continue
                m = g.mesclas.get((r, c))
                x0, x1 = g.x_col[c], g.x_col[c] + g.larg_col[c]
                y0, y1 = alt_pag - pos[r][1], alt_pag - pos[r][1] - alturas[r]
                lados = ((est.bordas[0], (x0, y0, x0, y1), m and c > m[1]),
                         (est.bordas[1], (x1, y0, x1, y1), m and c < m[3]),
                         (est.bordas[2], (x0, y0, x1, y0), m and r > m[0]),
                         (est.bordas[3], (x0, y1, x1, y1), m and r < m[2]))
                for borda, (xa, ya, xb, yb), interna in lados:
                    if borda and not interna:
                        estilo, cor = borda
                        espessura = ESPESSURA_BORDA.get(estilo, 0.6) * k
                        if xa == xb:  # o Excel desenha a linha vertical por dentro da borda esquerda
                            xa = xb = xa - espessura / 2
                        cv.setLineWidth(espessura)
                        cv.setStrokeColorRGB(*(cor or (0, 0, 0)))
                        if "dash" in estilo.lower() or "dot" in estilo.lower():
                            cv.setDash(2 * k, 2 * k)
                        else:
                            cv.setDash()
                        cv.line(xa, ya, xb, yb)
        cv.showPage()
    cv.save()


def _desenhar_texto(cv, t: _Texto, x0, y0, x1, y1, alt_pag, k) -> None:
    """Desenha o texto na caixa (coordenadas a partir do topo da página)."""
    n = len(t.linhas) + t.vazias_fim
    passo = ENTRELINHA * t.corpo
    if t.vertical == "top":
        base = y0 + BASE_SUPERIOR * t.corpo
    elif t.vertical in ("center", "justify", "distributed"):
        base = (y0 + y1) / 2 + BASE_CENTRO * t.corpo - (n - 1) * passo / 2
    else:
        base = y1 - BASE_INFERIOR * k - (n - 1) * passo
    cv.setFillColorRGB(*(t.cor or (0, 0, 0)))
    esq = x0 + MARGEM_TEXTO * k + (t.recuo if t.horizontal in ("left", "justify", "distributed") else 0)
    dir_ = x1 - MARGEM_DIREITA * k - (t.recuo if t.horizontal == "right" else 0) - t.reserva
    for i, (tokens, largura, fim_paragrafo) in enumerate(t.linhas):
        y = alt_pag - (base + i * passo)
        extra = 0.0
        if t.horizontal in ("right",):
            x = dir_ - largura
        elif t.horizontal in ("center", "centerContinuous"):
            x = (x0 + x1 - largura - t.reserva) / 2
        else:
            x = esq
            if t.justificar and not fim_paragrafo:
                espacos = sum(len(tok[0]) for tok in tokens if tok[0].isspace())
                if espacos:
                    extra = max(0.0, (dir_ - esq - largura) / espacos)
        if extra:
            for texto, fonte, corpo, w in tokens:
                if texto.isspace():
                    x += w + extra * len(texto)
                    continue
                cv.setFont(fonte, corpo)
                cv.drawString(x, y, texto)
                x += w
            continue
        # junta os tokens de mesma fonte num único trecho
        grupo, fonte_g, corpo_g, inicio = "", None, None, x
        for texto, fonte, corpo, w in tokens:
            if (fonte, corpo) != (fonte_g, corpo_g):
                if grupo:
                    cv.setFont(fonte_g, corpo_g)
                    cv.drawString(inicio, y, grupo)
                grupo, fonte_g, corpo_g, inicio = "", fonte, corpo, x
            grupo += texto
            x += w
        if grupo:
            cv.setFont(fonte_g, corpo_g)
            cv.drawString(inicio, y, grupo)


# --------------------------------------------------------------------------
# comandos
# --------------------------------------------------------------------------
# --------------------------------------------------------------------------
# conferência com o Informe Mensal FIDC publicado pela CVM (dados abertos)
# --------------------------------------------------------------------------
URL_CVM_PADRAO = "https://dados.cvm.gov.br/dados/FIDC/DOC/INF_MENSAL/DADOS/inf_mensal_fidc_{aaaamm}.zip"
TABELAS_CVM = ("I", "III", "IV")  # ativo, passivo e patrimônio líquido

# Linhas do quadro I conferidas (a 18, Valores a Receber, não tem campo no informe)
LINHAS_CONFERIDAS = {
    13: "Disponibilidades",
    14: "Títulos Públicos Federais",
    15: "Operações de Crédito",
    16: "Provisão para Perdas (PDD)",
    17: "Outros Ativos",
    19: "Contas a Pagar",
    20: "Patrimônio Líquido",
}

# Subpastas da pasta de saída (config.json -> "subpastas")
SUBPASTAS_PADRAO = {"revisados": "REVISADOS", "revisao_humana": "SEM INFORME", "insumos": "DRAFT EXCEL",
                    # dentro de "revisados": fundos com textos trocados por pipefy_atualizacoes.xlsx
                    "revisados_atualizados": "ALTERADOS PELA PLANILHA"}

STATUS_OK = "OK"
STATUS_CORRIGIDO = "corrigido com a CVM"
STATUS_INCONSISTENTE = "informe da CVM inconsistente"
STATUS_SEM_INFORME = "sem informe na CVM"


def _br(valor) -> str:
    return "" if valor is None else f"{valor:.2f}".replace(".", ",")


def _cfg_cvm(cfg: dict) -> dict:
    return {"url": URL_CVM_PADRAO, "tolerancia": 1.0, **cfg.get("cvm", {})}


def baixar_informe_cvm(cfg: dict, data_base: dt.date) -> Path:
    """Baixa o zip do mês da data base para entrada/cvm. Sem rede, usa a cópia já baixada."""
    aaaamm = data_base.strftime("%Y%m")
    pasta = caminho_cfg(cfg, "entrada") / "cvm"
    pasta.mkdir(parents=True, exist_ok=True)
    destino = pasta / f"inf_mensal_fidc_{aaaamm}.zip"
    url = _cfg_cvm(cfg)["url"].format(aaaamm=aaaamm)
    try:
        req = urllib.request.Request(url, headers={"User-Agent": "geradora-cvm-limine"})
        with urllib.request.urlopen(req, timeout=120) as resp:
            conteudo = resp.read()
        zipfile.ZipFile(io.BytesIO(conteudo)).testzip()
        temporario = destino.with_suffix(".tmp")
        temporario.write_bytes(conteudo)
        os.replace(temporario, destino)
        print(f"[INFO] Informe Mensal FIDC {aaaamm} baixado da CVM: {url}")
    except Exception as e:
        if not destino.exists():
            raise SystemExit(f"[ERRO] Não foi possível baixar o Informe Mensal FIDC da CVM ({url}): {e}\n"
                             f"Baixe o arquivo para {destino} ou rode com --sem-cvm.") from e
        print(f"[AVISO] Não foi possível baixar o informe da CVM ({e}); usando a cópia local {destino}")
    return destino


def ler_informe_cvm(arquivo_zip: Path, data_base: dt.date) -> dict[str, dict]:
    """CNPJ (só dígitos) -> campos das tabelas I, III e IV do informe do mês."""
    aaaamm = data_base.strftime("%Y%m")
    informe: dict[str, dict] = {}
    with zipfile.ZipFile(arquivo_zip) as z:
        nomes = {n.lower(): n for n in z.namelist()}
        for tab in TABELAS_CVM:
            nome = nomes.get(f"inf_mensal_fidc_tab_{tab}_{aaaamm}.csv".lower())
            if not nome:
                raise SystemExit(f"[ERRO] {arquivo_zip} não tem a tabela {tab} de {aaaamm}")
            texto = z.read(nome).decode("latin-1")
            vistos = set()
            for linha in csv.DictReader(io.StringIO(texto), delimiter=";"):
                cnpj = digitos(linha.get("CNPJ_FUNDO_CLASSE"))
                if not cnpj or cnpj in vistos:   # 1ª linha de cada CNPJ em cada tabela
                    continue
                vistos.add(cnpj)
                informe.setdefault(cnpj, {}).update(linha)
    return informe


def valores_cvm(reg: dict) -> dict:
    """Linhas do quadro I segundo o informe da CVM, com a mesma classificação da carteira.

    Operações de Crédito vêm brutas (com e sem aquisição substancial de riscos) e a PDD sai à parte.
    Outros Ativos é o restante do ativo, sem compromissadas e derivativos, que o quadro não mostra.
    """
    def v(campo):  # os CSVs da CVM usam ponto decimal
        try:
            return float(reg.get(campo) or 0)
        except ValueError:
            return 0.0
    pdd = v("TAB_I2A11_VL_REDUCAO_RECUP") + v("TAB_I2B11_VL_REDUCAO_RECUP")
    credito = v("TAB_I2A_VL_DIRCRED_RISCO") + v("TAB_I2B_VL_DIRCRED_SEM_RISCO")
    disp, tpf, ativo = v("TAB_I1_VL_DISP"), v("TAB_I2D_VL_TITPUB_FED"), v("TAB_I_VL_ATIVO")
    return {
        13: disp,
        14: tpf,
        15: credito + pdd,
        16: -pdd,
        17: ativo - disp - tpf - credito - v("TAB_I2F_VL_OPER_COMPROM") - v("TAB_I3_VL_POSICAO_DERIV"),
        19: -v("TAB_III_A_VL_PAGAR"),
        20: v("TAB_IV_A_VL_PL"),
        "ativo": ativo,
        "passivo": v("TAB_III_VL_PASSIVO"),
    }


@dataclass
class ConferenciaCVM:
    status: str
    divergencias: list = field(default_factory=list)  # (linha, descrição, valor da carteira, valor da CVM)
    acao: str = ""

    @property
    def revisar(self) -> bool:
        return self.status != STATUS_OK


def conferir_cvm(rel: Relatorio, informe: dict[str, dict], tolerancia: float = 1.0) -> ConferenciaCVM:
    """Compara o quadro I com o informe da CVM e corrige o relatório quando o informe é confiável.

    - todas as linhas batem (diferença <= tolerância): OK
    - o informe não fecha (ativo - passivo != PL): mantém a carteira e manda para revisão
    - senão: troca os valores divergentes pelos da CVM, recalcula os % e manda para revisão
    """
    reg = informe.get(digitos(rel.cnpj))
    if reg is None:
        rel.avisos.append("Conferência CVM: fundo sem Informe Mensal na CVM para o mês da data base")
        return ConferenciaCVM(STATUS_SEM_INFORME, acao="conferir manualmente")
    cvm = valores_cvm(reg)
    c = rel.celulas
    divergencias = [(linha, desc, c.get(f"C{linha}", 0), cvm[linha])
                    for linha, desc in LINHAS_CONFERIDAS.items()
                    if abs(c.get(f"C{linha}", 0) - cvm[linha]) > tolerancia]
    if not divergencias:
        return ConferenciaCVM(STATUS_OK)
    texto = "; ".join(f"{d}: carteira {a:.2f}, CVM {b:.2f}" for _, d, a, b in divergencias)
    fechamento = cvm["ativo"] - cvm["passivo"] - cvm[20]
    if abs(cvm["ativo"]) <= tolerancia and abs(cvm[20]) <= tolerancia:
        acao = "mantidos os valores da carteira: o informe da CVM está zerado (ativo e PL = 0)"
        rel.avisos.append(f"Conferência CVM: {texto}; {acao}")
        return ConferenciaCVM(STATUS_INCONSISTENTE, divergencias, acao)
    if abs(fechamento) > tolerancia:
        acao = (f"mantidos os valores da carteira: o informe da CVM não fecha (ativo {cvm['ativo']:.2f} "
                f"- passivo {cvm['passivo']:.2f} - PL {cvm[20]:.2f} = {fechamento:.2f})")
        rel.avisos.append(f"Conferência CVM: {texto}; {acao}")
        return ConferenciaCVM(STATUS_INCONSISTENTE, divergencias, acao)
    for linha, _, _, valor in divergencias:
        c[f"C{linha}"] = valor
    rel.avisos[:] = [a for a in rel.avisos if not a.startswith(AVISO_SOMA)]
    aviso_soma = percentuais_composicao(c)
    if aviso_soma:
        rel.avisos.append(aviso_soma)
    rel.avisos.append(f"Conferência CVM: corrigido com o informe da CVM ({texto})")
    return ConferenciaCVM(STATUS_CORRIGIDO, divergencias, "relatório corrigido com o valor da CVM")


def cmd_gerar(args, cfg):
    bases = carregar_bases(cfg)
    for a in bases.avisos_gerais:
        print(f"[AVISO] {a}")
    data_base = para_data(args.data_base or cfg["data_base"])
    saida = Path(args.saida) if args.saida else caminho_cfg(cfg, "saida")
    caminho_gravavel(saida).mkdir(parents=True, exist_ok=True)
    # Com a conferência da CVM, os arquivos vão para subpastas da saída:
    #   PDF que bate com o informe -> "revisados" (pronto para emitir)
    #   PDF que não bate           -> "revisao_humana" (com o revisao_cvm.csv)
    #   XLSX (draft) de todos      -> "insumos"
    informe, revisados, revisados_alt, insumos, revisao = None, saida, saida, saida, None
    tolerancia = float(_cfg_cvm(cfg)["tolerancia"])
    if not args.sem_cvm:
        informe = ler_informe_cvm(baixar_informe_cvm(cfg, data_base), data_base)
        subpastas = {**SUBPASTAS_PADRAO, **cfg.get("subpastas", {})}
        revisados = saida / subpastas["revisados"]
        revisados_alt = revisados / subpastas["revisados_atualizados"]
        insumos = saida / subpastas["insumos"]
        revisao = saida / subpastas["revisao_humana"]
        for pasta in (revisados, insumos, revisao):  # a subpasta dos atualizados só é criada se for usada
            caminho_gravavel(pasta).mkdir(parents=True, exist_ok=True)
    pastas_geradas = [p for p in {saida, revisados, revisados_alt, insumos, revisao} if p is not None]
    modelo = carregar_modelo()
    rotulos = rotulos_composicao(modelo.ws)
    texto_padrao = cfg["texto_gestor_sem_resposta"]
    formatos = {"pdf", "xlsx"} if args.formato == "ambos" else {args.formato}

    resumo, erros = [], 0
    if args.cnpj:
        alvo = [c.strip() for c in args.cnpj]
    else:
        fora = fundos_fora_da_lista(bases)
        if fora:
            print(f"[AVISO] {len(fora)} fundo(s) na Carteira Informe que não estão na lista de fundos "
                  "(não gerados; inclua nome e CNPJ em fundos.xlsx): "
                  + "; ".join(f"{c} {n}".strip() for c, n in fora))
        for c, n in fora:
            resumo.append({"fundo": n, "cnpj": c, "pl": "", "arquivo": "",
                           "avisos": "NÃO GERADO: está na Carteira Informe mas não na lista de fundos "
                                     "(inclua nome e CNPJ em fundos.xlsx)"})
        # imprimir_exportar: percorre a lista e gera só quem "Tem Informe?" = SIM
        alvo = [f["cnpj"] for f in bases.fundos if tem_carteira(bases, f["cnpj"])]
        pulados = [f["nome"] or f["cnpj"] for f in bases.fundos if not tem_carteira(bases, f["cnpj"])]
        if pulados:
            print(f"[INFO] {len(pulados)} fundo(s) sem carteira na base, não gerados: " + "; ".join(pulados))
    unicos, vistos = [], set()
    for c in alvo:  # 1ª ocorrência de cada CNPJ, na ordem da lista
        if digitos(c) not in vistos:
            vistos.add(digitos(c))
            unicos.append(c)
    if len(unicos) < len(alvo):
        print(f"[INFO] {len(alvo) - len(unicos)} CNPJ(s) repetido(s) na lista; cada fundo é gerado uma vez")
    alvo = unicos

    pasta_validados = Path(args.validados) if args.validados else None
    validados = list(pasta_validados.glob("*.xlsx")) if pasta_validados else []
    nao_listados = len(resumo)
    revisoes = []
    inicio = time.perf_counter()
    for cnpj in alvo:
        rel = None
        try:
            validacao = None
            achados = [p for p in validados if digitos(cnpj) in digitos(p.stem)]
            if achados:
                validacao = ler_validacao(achados[0])
            rel = calcular(bases, cnpj, data_base, rotulos, texto_padrao, validacao)
            if not rel.nome:
                rel.avisos.append("CNPJ não está na lista de fundos (nome em branco)")
            conf = conferir_cvm(rel, informe, tolerancia) if informe is not None else None
            revisar = bool(conf and conf.revisar)
            pdf_ok = revisados_alt if rel.atualizado_pipefy else revisados
            destinos = {"pdf": revisao if revisar else pdf_ok, "xlsx": insumos}
            for ext in sorted(formatos):
                destino = destinos[ext]
                for outra in pastas_geradas:  # se o fundo trocou de pasta desde a última geração, tira a cópia antiga
                    antigo = outra / f"{rel.nome_arquivo}.{ext}"
                    if outra != destino and antigo.is_file():
                        antigo.unlink()
                caminho_gravavel(destino).mkdir(parents=True, exist_ok=True)
                arquivo = destino / f"{rel.nome_arquivo}.{ext}"
                if ext == "pdf":
                    salvar_pdf(rel, caminho_gravavel(arquivo), args.com_lastro, modelo)
                else:
                    salvar_xlsx(rel, arquivo, args.com_lastro, modelo)
            print(("[REVISAR] " if revisar else "[OK] ") + rel.nome_arquivo
                  + (f"  ({conf.status})" if revisar else "")
                  + (f"  ({len(rel.avisos)} aviso(s))" if rel.avisos else ""))
            resumo.append({"fundo": rel.nome, "cnpj": rel.cnpj, "pl": rel.celulas["C20"],
                           "arquivo": rel.nome_arquivo,
                           "conferencia_cvm": conf.status if conf else "não conferido",
                           "pasta": " | ".join(sorted({str(destinos[e]) for e in formatos})), "avisos": " | ".join(rel.avisos)})
            if revisar:
                base_rev = {"fundo": rel.nome, "cnpj": rel.cnpj, "status": conf.status, "acao": conf.acao}
                for linha, desc, a, b in conf.divergencias or [("", "", None, None)]:
                    revisoes.append({**base_rev, "linha": f"C{linha}" if linha else "", "descricao": desc,
                                     "valor_carteira": _br(a), "valor_cvm": _br(b),
                                     "diferenca": _br(a - b) if a is not None else ""})
        except Exception as e:  # um fundo com problema não interrompe o lote
            erros += 1
            nome = rel.nome_arquivo if rel else cnpj
            print(f"[ERRO] {nome}: {type(e).__name__}: {e}")
            resumo.append({"fundo": rel.nome if rel else "", "cnpj": cnpj, "pl": "",
                           "arquivo": rel.nome_arquivo if rel else "", "conferencia_cvm": "", "pasta": "",
                           "avisos": f"ERRO ao gerar: {type(e).__name__}: {e}"})

    with open(caminho_gravavel(saida / "resumo_geracao.csv"), "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["fundo", "cnpj", "pl", "arquivo", "conferencia_cvm", "pasta", "avisos"],
                           delimiter=";")
        w.writeheader()
        w.writerows(resumo)
    if revisao is not None:
        with open(caminho_gravavel(revisao / "revisao_cvm.csv"), "w", encoding="utf-8-sig", newline="") as f:
            w = csv.DictWriter(f, fieldnames=["fundo", "cnpj", "status", "linha", "descricao", "valor_carteira",
                                              "valor_cvm", "diferenca", "acao"], delimiter=";")
            w.writeheader()
            w.writerows(revisoes)
    print(f"\nExportação em lote concluída: {len(resumo) - erros - nao_listados} fundo(s) em {saida} "
          f"({time.perf_counter() - inicio:.1f}s)")
    if nao_listados:
        print(f"[AVISO] {nao_listados} fundo(s) da carteira ficaram de fora por não estarem na lista; "
              "veja o resumo.")
    print(f"Resumo com avisos por fundo: {saida / 'resumo_geracao.csv'}")
    if revisao is not None:
        n_rev = len({r["cnpj"] for r in revisoes})
        n_ok = sum(1 for r in resumo if r.get("conferencia_cvm") == STATUS_OK)
        print(f"Conferência com a CVM: {n_ok} fundo(s) batem com o informe: PDF em {revisados}")
        n_alt = sum(1 for r in resumo if r.get("conferencia_cvm") == STATUS_OK and str(revisados_alt) in r["pasta"])
        if n_alt:
            print(f"                       ({n_alt} deles, com textos alterados pela planilha do Pipefy, em "
                  f"{revisados_alt})")
        print(f"                       Excel (draft) de todos os fundos em {insumos}")
        print(f"                       {n_rev} fundo(s) não batem: PDF em {revisao} (detalhes em revisao_cvm.csv)")
    if erros:
        print(f"[ERRO] {erros} fundo(s) não foram gerados; veja o resumo.")
        sys.exit(1)


def _post_graphql(url: str, token: str, query: str) -> dict:
    req = urllib.request.Request(url, data=json.dumps({"query": query}).encode("utf-8"), method="POST",
                                 headers={"Authorization": f"Bearer {token}", "Content-Type": "application/json"})
    with urllib.request.urlopen(req, timeout=60) as r:
        resposta = json.loads(r.read().decode("utf-8"))
    if resposta.get("errors"):
        raise RuntimeError(f"Erro da API do Pipefy: {resposta['errors']}")
    return resposta["data"]


def cmd_pipefy(args, cfg):
    """BaixarEImportarPipefyReport: exporta o relatório do pipe e salva como pipefy.xlsx."""
    pc = cfg["pipefy"]
    token = os.environ.get(pc["variavel_token"])
    if not token:
        sys.exit(f"Defina a variável de ambiente {pc['variavel_token']} com o token da API do Pipefy.")
    data = _post_graphql(pc["url"], token,
                         f"mutation {{ exportPipeReport(input: {{pipeId: {pc['pipe_id']}, "
                         f"pipeReportId: {pc['pipe_report_id']}}}) {{ pipeReportExport {{ id }} }} }}")
    export_id = data["exportPipeReport"]["pipeReportExport"]["id"]
    url_arquivo = None
    for _ in range(30):  # o arquivo leva alguns segundos para ficar pronto
        data = _post_graphql(pc["url"], token, f"{{ pipeReportExport(id: {export_id}) {{ fileURL }} }}")
        url_arquivo = data["pipeReportExport"]["fileURL"]
        if url_arquivo:
            break
        time.sleep(2)
    if not url_arquivo:
        sys.exit("O Pipefy não devolveu o link do relatório a tempo; tente de novo.")
    destino = arquivo_entrada(cfg, "pipefy")
    destino.parent.mkdir(parents=True, exist_ok=True)
    with urllib.request.urlopen(url_arquivo, timeout=120) as r, open(destino, "wb") as f:
        f.write(r.read())
    print(f"Relatório Pipefy salvo em {destino} ({len(ler_tabela(destino))} linhas)")


def cmd_listar(args, cfg):
    """listarArquivos: lista os .xlsx da pasta de validados."""
    pasta = Path(args.pasta) if args.pasta else caminho_cfg(cfg, "validados")
    if not pasta.is_dir():
        sys.exit(f"O caminho especificado não existe: {pasta}")
    arquivos = sorted(p.name for p in pasta.glob("*.xlsx"))
    for a in arquivos:
        print(a)
    print(f"\n{len(arquivos)} arquivo(s) em {pasta}")


def cmd_migrar(args, cfg):
    """Extrai as bases coladas numa Geradora .xlsb para a pasta de entrada."""
    from pyxlsb import open_workbook

    origem = Path(args.geradora)
    entrada = caminho_cfg(cfg, "entrada")
    entrada.mkdir(parents=True, exist_ok=True)

    def ler_aba(nome):
        with open_workbook(str(origem)) as wb:
            with wb.get_sheet(nome) as sh:
                linhas = {}
                for row in sh.rows():
                    for c in row:
                        if c.v is not None:
                            linhas.setdefault(c.r, {})[c.c] = c.v
        return linhas

    def gravar(nome_arquivo, cabecalho, linhas):
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.append(cabecalho)
        for l in linhas:
            ws.append(l)
        wb.save(entrada / nome_arquivo)
        print(f"  {nome_arquivo}: {len(linhas)} linha(s)")

    def serial_para_data(v):
        return para_data(v) if isinstance(v, (int, float)) else v

    print(f"Lendo {origem}")
    setup = ler_aba("Setup")
    fundos = [[r.get(6), r.get(7), r.get(9)] for i, r in sorted(setup.items()) if i >= 5 and r.get(7)]
    gravar(cfg["arquivos"]["fundos"], ["Nome Fundo", "CNPJ", "Gerar (x)"], fundos)

    ci = ler_aba("CarteiraInforme")
    gravar(cfg["arquivos"]["carteira_informe"],
           ["Data Base", "CNPJ Fundo", "Descricao", "Valor", "Quantidade de Cotas", "Valor da Cota"],
           [[serial_para_data(r.get(0)), r.get(1), r.get(2), r.get(3), r.get(4), r.get(5)]
            for i, r in sorted(ci.items()) if i >= 1 and r.get(1)])

    rs = ler_aba("RESUMO_SINQIA")
    gravar(cfg["arquivos"]["resumo_sinqia"],
           ["SITUACAO_RECEBIVEL", "DOC_FUNDO", "NOME_FUNDO", "RELATORIO", "QTDE", "VALOR_NOMINAL", "VALOR_PRESENTE"],
           [[r.get(c) for c in range(7)] for i, r in sorted(rs.items()) if i >= 1 and any(r.get(c) is not None for c in range(7))])

    pf = ler_aba("Pipefy")
    ncol = max((c for c in pf.get(0, {})), default=-1) + 1
    gravar(cfg["arquivos"]["pipefy"], [pf[0].get(c) for c in range(ncol)],
           [[r.get(c) for c in range(ncol)] for i, r in sorted(pf.items()) if i >= 1 and r.get(1)])

    data_base = serial_para_data(setup.get(5, {}).get(2))
    if isinstance(data_base, dt.date):
        caminho = Path(args.config)
        texto = caminho.read_text(encoding="utf-8")
        texto = re.sub(r'("data_base"\s*:\s*")[^"]*(")', rf"\g<1>{data_base.isoformat()}\g<2>", texto, count=1)
        caminho.write_text(texto, encoding="utf-8")
        print(f"  data_base do config.json atualizada para {data_base.isoformat()}")


def _saida_utf8():
    """Evita UnicodeEncodeError ao gravar a saída num arquivo/pipe com a codificação do Windows."""
    for fluxo in (sys.stdout, sys.stderr):
        try:
            if fluxo.isatty():
                fluxo.reconfigure(errors="replace")
            else:
                fluxo.reconfigure(encoding="utf-8", errors="replace")
        except (AttributeError, ValueError):
            pass


def main(argv=None):
    _saida_utf8()
    ap = argparse.ArgumentParser(description="Geradora do Demonstrativo Trimestral CVM - Limine")
    ap.add_argument("--config", default=str(BASE_DIR / "config.json"), help="arquivo de configuração")
    sub = ap.add_subparsers(dest="comando", required=True)

    g = sub.add_parser("gerar", help="gera PDF e XLSX do demonstrativo para os fundos")
    g.add_argument("--cnpj", nargs="+", help="gera só estes CNPJs (padrão: todos com carteira)")
    g.add_argument("--data-base", help="AAAA-MM-DD (padrão: config.json)")
    g.add_argument("--saida", help="pasta de saída (padrão: pastas.saida do config.json)")
    g.add_argument("--formato", choices=["ambos", "pdf", "xlsx"], default="ambos")
    g.add_argument("--com-lastro", action="store_true",
                   help="mostra as seções II e III (ocultas na planilha original)")
    g.add_argument("--validados", help="pasta com as validações por fundo (linhas 32-34; arquivo com o CNPJ no nome)")
    g.add_argument("--sem-cvm", action="store_true",
                   help="não confere com o Informe Mensal FIDC da CVM (tudo vai para a pasta de saída)")
    g.set_defaults(func=cmd_gerar)

    p = sub.add_parser("pipefy", help="baixa o relatório do Pipefy para a pasta de entrada")
    p.set_defaults(func=cmd_pipefy)

    lv = sub.add_parser("listar-validados", help="lista os arquivos .xlsx da pasta de validados")
    lv.add_argument("--pasta")
    lv.set_defaults(func=cmd_listar)

    m = sub.add_parser("migrar", help="extrai as bases de uma Geradora .xlsb para a pasta de entrada")
    m.add_argument("--geradora", required=True)
    m.set_defaults(func=cmd_migrar)

    args = ap.parse_args(argv)
    cfg = ler_config(Path(args.config))
    args.func(args, cfg)


if __name__ == "__main__":
    main()
