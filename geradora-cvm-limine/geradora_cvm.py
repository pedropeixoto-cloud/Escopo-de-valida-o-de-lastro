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
comentários "CVM!X" apontam a célula de origem.
"""
from __future__ import annotations

import argparse
import csv
import datetime as dt
import json
import os
import re
import sys
import time
import urllib.request
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl
from openpyxl.cell.rich_text import CellRichText, TextBlock

BASE_DIR = Path(__file__).resolve().parent
ASSETS = BASE_DIR / "assets"
MODELO_XLSX = ASSETS / "modelo_cvm.xlsx"
LOGO = ASSETS / "logo_limine.png"

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

CARACTERES_INVALIDOS = '/\\:*?"<>|'


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


def mapear_descricao(descricao, mapeamento) -> str:
    """CarteiraInforme!G: primeira chave (na ordem da tabela I:J) contida em LOWER(Descricao).

    =XLOOKUP(1,--ISNUMBER(FIND(chaves,LOWER(C2),1)),mapeamentos,"")
    """
    texto = str(descricao).lower() if descricao is not None else ""
    for chave, destino in mapeamento:
        if chave and chave in texto:
            return destino
    return ""


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

    return Bases(fundos, carteira, sinqia, pipefy, mapeamento, avisos)


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

    @property
    def nome_arquivo(self) -> str:
        return nome_arquivo_seguro(f"{self.nome.strip()}_{digitos(self.cnpj)}_TRIMESTRAL_CVM")


def tem_carteira(bases: Bases, cnpj: str) -> bool:
    """Setup!O / Setup!E: COUNTIF(CarteiraInforme!B:B, CNPJ) > 0."""
    d = digitos(cnpj)
    return any(l["cnpj"] == d for l in bases.carteira)


def nome_do_fundo(bases: Bases, cnpj: str) -> str:
    """Setup!C7 =XLOOKUP(cnpj, H:H, G:G): primeiro nome da lista com esse CNPJ."""
    d = digitos(cnpj)
    for f in bases.fundos:
        if digitos(f["cnpj"]) == d:
            return f["nome"]
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

    # I. Composição da carteira
    # C13:C20 =SUMIFS(CarteiraInforme!D:D, G:G, rótulo, A:A, data base, B:B, CNPJ)
    linhas_fundo = [l for l in bases.carteira if l["cnpj"] == cnpj_dig and l["data"] == data_base]
    if not linhas_fundo:
        avisos.append("Fundo sem linhas na Carteira Informe para a data base")
    for linha, rotulo in rotulos_composicao.items():
        c[f"C{linha}"] = sum(l["valor"] for l in linhas_fundo if norm(l["categoria"]) == norm(rotulo))
    pl = c["C20"]
    for linha in range(13, 20):
        c[f"D{linha}"] = div(c[f"C{linha}"], pl)                       # =IFERROR(C13/$C$20,0%)
    c["D20"] = sum(c[f"D{l}"] for l in range(13, 20))                  # =SUM(D13:D19)
    if pl and abs(sum(c[f"C{l}"] for l in range(13, 20)) - pl) > 0.01:
        avisos.append("Soma das linhas da composição difere do Patrimônio Líquido "
                      f"(soma={sum(c[f'C{l}'] for l in range(13, 20)):.2f}; PL={pl:.2f})")
    outras = sorted({l["categoria"] for l in linhas_fundo if l["categoria"]
                     and norm(l["categoria"]) not in {norm(r) for r in rotulos_composicao.values()}})
    if outras:
        avisos.append("Categorias da carteira que não aparecem no quadro: " + ", ".join(outras))

    # II. Verificação de lastro (Resumo Sinqia). G4 = LEFT(CNPJ sem pontuação, 13)
    doc = int(cnpj_dig[:13]) if cnpj_dig else None
    sq = [l for l in bases.sinqia if l["doc"] is not None and l["doc"] == doc]

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
    resposta = next((p for p in bases.pipefy if digitos(coluna(p, "CNPJ do Fundo")) == cnpj_dig), None)
    if resposta is None:
        avisos.append("Sem resposta do gestor no Pipefy: textos padrão usados")
    for linha, campo in TEXTOS_GESTOR.items():
        valor = coluna(resposta, campo) if resposta else None
        c[f"A{linha}"] = texto_padrao if valor in (None, "") else str(valor)   # =IF(K5="",$I$8,K5)

    return Relatorio(nome, cnpj, data_base, c, avisos)


# --------------------------------------------------------------------------
# saída XLSX (a partir do modelo com a formatação original da aba CVM)
# --------------------------------------------------------------------------
def carregar_modelo():
    return openpyxl.load_workbook(MODELO_XLSX, rich_text=True)


def rotulos_composicao(ws) -> dict:
    return {linha: str(ws[f"B{linha}"].value).strip() for linha in range(13, 21)}


def salvar_xlsx(rel: Relatorio, destino: Path, mostrar_lastro: bool) -> None:
    wb = carregar_modelo()
    ws = wb["CVM"]
    for ref, valor in rel.celulas.items():
        ws[ref].value = valor
    for linha in LINHAS_SECAO_LASTRO:
        ws.row_dimensions[linha].hidden = not mostrar_lastro
    wb.save(destino)


# --------------------------------------------------------------------------
# saída PDF (reportlab), imitando a impressão da aba CVM (A4, zoom 68%)
# --------------------------------------------------------------------------
def fmt_contabil(v: float, casas: int = 2) -> str:
    """Formato contábil da planilha: 1.234,56 / (1.234,56) / - para zero."""
    if round(v, casas) == 0:
        return "-"
    s = f"{abs(v):,.{casas}f}".replace(",", "X").replace(".", ",").replace("X", ".")
    return f"({s})" if v < 0 else s


def fmt_pct(v: float) -> str:
    return f"{v * 100:.2f}%".replace(".", ",")


def texto_celula(valor) -> str:
    """Converte texto (ou rich text do Excel) em marcação do reportlab."""
    from xml.sax.saxutils import escape

    if valor is None:
        return ""
    if isinstance(valor, CellRichText):
        partes = []
        for bloco in valor:
            if isinstance(bloco, TextBlock):
                t = escape(bloco.text)
                partes.append(f"<b>{t}</b>" if bloco.font is not None and bloco.font.b else t)
            else:
                partes.append(escape(str(bloco)))
        txt = "".join(partes)
    else:
        txt = escape(str(valor))
    return txt.strip().replace("\n", "<br/>")


def salvar_pdf(rel: Relatorio, destino: Path, mostrar_lastro: bool) -> None:
    from reportlab.lib import colors
    from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT
    from reportlab.lib.pagesizes import A4
    from reportlab.lib.styles import ParagraphStyle
    from reportlab.platypus import (PageBreak, Paragraph, SimpleDocTemplate, Spacer, Table,
                                    TableStyle)

    ws = carregar_modelo()["CVM"]

    def v(ref):
        return rel.celulas.get(ref, ws[ref].value)

    escala = 0.68                       # zoom de impressão da aba
    f11, f10 = 11 * escala, 10 * escala
    linha = 13.8 * escala               # altura padrão de linha
    margem_lat, margem_vert = 28.35, 85.04
    largura_util = A4[0] - 2 * margem_lat
    larg_col = [35.55, 30.66, 17.55, 15.11, 12.55, 27.22]  # larguras A:F (caracteres)
    total = sum(larg_col)
    W = [largura_util * w / total for w in larg_col]
    azul = colors.HexColor("#DEEAF6")

    normal = ParagraphStyle("n", fontName="Helvetica", fontSize=f11, leading=f11 * 1.25)
    just = ParagraphStyle("j", parent=normal, alignment=TA_JUSTIFY)
    negrito = ParagraphStyle("b", parent=normal, fontName="Helvetica-Bold")
    titulo = ParagraphStyle("t", parent=negrito, alignment=TA_CENTER)
    recuo = ParagraphStyle("i", parent=just, leftIndent=f11 * 0.9)
    cel = ParagraphStyle("c", fontName="Helvetica", fontSize=f10, leading=f10 * 1.2)
    cel_b = ParagraphStyle("cb", parent=cel, fontName="Helvetica-Bold")
    cel_bc = ParagraphStyle("cbc", parent=cel_b, alignment=TA_CENTER)

    def esp(n=1):
        return Spacer(1, linha * n)

    def p(ref, estilo=just):
        return Paragraph(texto_celula(v(ref)), estilo)

    story = [p("A1", titulo), esp()]

    # Cabeçalho (A3:F6)
    cab = [[Paragraph(texto_celula(v(f"A{r}")), cel_b),
            Paragraph(v(f"B{r}").strftime("%d/%m/%Y") if r == 6 else texto_celula(v(f"B{r}")),
                      ParagraphStyle("r", parent=cel_b, alignment=2))] for r in range(3, 7)]
    t = Table(cab, colWidths=[W[0], sum(W[1:])], rowHeights=linha)
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                           ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                           ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
    story += [t, esp(), p("A8", negrito), esp(), p("A10", normal), esp()]

    # Quadro de composição (B12:D20)
    dados = [[Paragraph(texto_celula(v(f"{c}12")), cel_bc) for c in "BCD"]]
    for r in range(13, 21):
        estilo = cel_b if r == 20 else cel
        dados.append([Paragraph(texto_celula(v(f"B{r}")), estilo),
                      Paragraph(fmt_contabil(v(f"C{r}")), ParagraphStyle("x", parent=cel, alignment=2)),
                      Paragraph(fmt_pct(v(f"D{r}")), ParagraphStyle("y", parent=estilo, alignment=2))])
    t = Table(dados, colWidths=W[1:4], rowHeights=linha)
    t.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                           ("BACKGROUND", (0, 0), (-1, 0), azul),
                           ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                           ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, -1), 1)]))
    tabela = Table([["", t]], colWidths=[W[0], sum(W[1:])])
    tabela.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0)]))
    story += [tabela, esp()]

    if mostrar_lastro:
        def quadro(r_cab, r_ini, r_fim):
            d = [[Paragraph(texto_celula(v(f"{c}{r_cab}")), cel_bc if c != "A" else cel_b) for c in "ABCDEF"]]
            for r in range(r_ini, r_fim + 1):
                bold = r == 29
                d.append([Paragraph(texto_celula(v(f"A{r}")), cel_b if bold else cel)]
                         + [Paragraph(fmt_contabil(num(v(f"B{r}")), 0 if r_cab == 26 else 2),
                                      ParagraphStyle("q", parent=cel, alignment=2))]
                         + [Paragraph(fmt_contabil(num(v(f"{c}{r}"))), ParagraphStyle("q", parent=cel, alignment=2))
                            for c in "CD"]
                         + [Paragraph(fmt_pct(num(v(f"{c}{r}"))) if r_cab == 26 else fmt_contabil(num(v(f"{c}{r}"))),
                                      ParagraphStyle("q", parent=cel_b if bold else cel, alignment=2)) for c in "EF"])
            tq = Table(d, colWidths=W)
            tq.setStyle(TableStyle([("GRID", (0, 0), (-1, -1), 0.5, colors.black),
                                    ("BACKGROUND", (0, 0), (-1, 0), azul),
                                    ("VALIGN", (0, 0), (-1, -1), "MIDDLE")]))
            return tq
        story += [p("A22", negrito), esp(), p("A24", normal), esp(), quadro(26, 27, 36), esp(2),
                  p("A39", negrito), esp(), p("A41", normal), esp(), quadro(43, 44, 46), esp()]

    story += [p("A48", negrito), esp(), p("A50", normal), esp(), p("A52"), esp(2),
              p("A55", negrito), esp(), p("A57"), esp(), p("A59"), esp(),
              p("A60"), esp(), p("A63", recuo), esp(), p("A66"), esp(2),
              p("A69", recuo), esp(), p("A72"), esp(2),
              p("A75"), esp(), p("A77"), esp(2),
              p("A80"), esp(), p("A82", recuo), esp(), p("A84"), esp(2),
              p("A87", recuo), esp(), p("A89"), esp(),
              PageBreak(),                                            # quebra manual da planilha
              p("A92"), esp(), p("A94"), esp(2),
              p("A97"), esp(), p("A100"), esp(2),
              p("A103"), esp(), p("A106"), esp(2),
              p("A109"), esp(), p("A112"), esp(3),
              p("A116", negrito)]

    def cabecalho(canvas, doc):
        if LOGO.exists():
            canvas.drawImage(str(LOGO), margem_lat, A4[1] - 28.35 - 30 * escala * 1.4,
                             width=102.6 * escala * 1.4, height=30 * escala * 1.4,
                             preserveAspectRatio=True, mask="auto")

    doc = SimpleDocTemplate(str(destino), pagesize=A4, leftMargin=margem_lat, rightMargin=margem_lat,
                            topMargin=margem_vert, bottomMargin=margem_vert,
                            title=f"Demonstrativo Trimestral - {rel.nome}", author="Limine Trust DTVM Ltda.")
    doc.build(story, onFirstPage=cabecalho, onLaterPages=cabecalho)


# --------------------------------------------------------------------------
# comandos
# --------------------------------------------------------------------------
def cmd_gerar(args, cfg):
    bases = carregar_bases(cfg)
    for a in bases.avisos_gerais:
        print(f"[AVISO] {a}")
    data_base = para_data(args.data_base or cfg["data_base"])
    saida = Path(args.saida) if args.saida else caminho_cfg(cfg, "saida")
    saida.mkdir(parents=True, exist_ok=True)
    ws = carregar_modelo()["CVM"]
    rotulos = rotulos_composicao(ws)
    texto_padrao = cfg["texto_gestor_sem_resposta"]
    formatos = {"pdf", "xlsx"} if args.formato == "ambos" else {args.formato}

    if args.cnpj:
        alvo = [c.strip() for c in args.cnpj]
    else:
        # imprimir_exportar: percorre a lista e gera só quem "Tem Informe?" = SIM
        alvo = [f["cnpj"] for f in bases.fundos if tem_carteira(bases, f["cnpj"])]
        pulados = [f["nome"] or f["cnpj"] for f in bases.fundos if not tem_carteira(bases, f["cnpj"])]
        if pulados:
            print(f"[INFO] {len(pulados)} fundo(s) sem carteira na base, não gerados: " + "; ".join(pulados))

    pasta_validados = Path(args.validados) if args.validados else None
    resumo = []
    for cnpj in alvo:
        validacao = None
        if pasta_validados:
            achados = [p for p in pasta_validados.glob("*.xlsx") if digitos(cnpj) in digitos(p.stem)]
            if achados:
                validacao = ler_validacao(achados[0])
        rel = calcular(bases, cnpj, data_base, rotulos, texto_padrao, validacao)
        if not rel.nome:
            rel.avisos.append("CNPJ não está na lista de fundos (nome em branco)")
        if "pdf" in formatos:
            salvar_pdf(rel, saida / f"{rel.nome_arquivo}.pdf", args.com_lastro)
        if "xlsx" in formatos:
            salvar_xlsx(rel, saida / f"{rel.nome_arquivo}.xlsx", args.com_lastro)
        print(f"[OK] {rel.nome_arquivo}" + (f"  ({len(rel.avisos)} aviso(s))" if rel.avisos else ""))
        resumo.append({"fundo": rel.nome, "cnpj": rel.cnpj, "pl": rel.celulas["C20"],
                       "arquivo": rel.nome_arquivo, "avisos": " | ".join(rel.avisos)})

    with open(saida / "resumo_geracao.csv", "w", encoding="utf-8-sig", newline="") as f:
        w = csv.DictWriter(f, fieldnames=["fundo", "cnpj", "pl", "arquivo", "avisos"], delimiter=";")
        w.writeheader()
        w.writerows(resumo)
    print(f"\nExportação em lote concluída: {len(resumo)} fundo(s) em {saida}")
    print(f"Resumo com avisos por fundo: {saida / 'resumo_geracao.csv'}")


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
        caminho = cfg["_dir"] / "config.json" if not args.config else Path(args.config)
        texto = Path(caminho).read_text(encoding="utf-8")
        texto = re.sub(r'("data_base"\s*:\s*")[^"]*(")', rf"\g<1>{data_base.isoformat()}\g<2>", texto, count=1)
        Path(caminho).write_text(texto, encoding="utf-8")
        print(f"  data_base do config.json atualizada para {data_base.isoformat()}")


def main(argv=None):
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
