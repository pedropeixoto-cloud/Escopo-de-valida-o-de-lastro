"""Testes da geradora com bases fictícias (as bases reais são confidenciais e ficam fora do git).

    python -m unittest discover -s tests -v
"""
import datetime as dt
import io
import json
import re
import shutil
import sys
import tempfile
import unittest
import zipfile
from contextlib import redirect_stdout
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))
import geradora_cvm as g  # noqa: E402

DATA_BASE = dt.date(2026, 9, 30)
TEXTO_PADRAO = json.loads((Path(g.BASE_DIR) / "config.json").read_text(encoding="utf-8"))["texto_gestor_sem_resposta"]

OK_DIR, INS_DIR, REV_DIR = (g.SUBPASTAS_PADRAO[k] for k in ("revisados", "insumos", "revisao_humana"))
ALFA, BETA, GAMA = "11.111.111/0001-11", "22.222.222/0001-22", "33.333.333/0001-33"


def _xlsx(caminho: Path, linhas: list[list]):
    wb = openpyxl.Workbook()
    for linha in linhas:
        wb.active.append(linha)
    wb.save(caminho)


# Informe Mensal FIDC fictício que bate com a carteira de ALFA e BETA
INFORME_OK = {
    ALFA: {"TAB_I_VL_ATIVO": 10500, "TAB_I1_VL_DISP": 1000, "TAB_I2A_VL_DIRCRED_RISCO": 8000,
           "TAB_I2A11_VL_REDUCAO_RECUP": 500, "TAB_I2D_VL_TITPUB_FED": 1500,
           "TAB_III_VL_PASSIVO": 500, "TAB_III_A_VL_PAGAR": 500, "TAB_IV_A_VL_PL": 10000},
    BETA: {"TAB_I_VL_ATIVO": 100, "TAB_I1_VL_DISP": 100, "TAB_IV_A_VL_PL": 100},
}


def montar_informe_cvm(destino: Path, fundos: dict):
    """Zip no formato dos dados abertos da CVM (tabelas I, III e IV, latin-1, separador ;)."""
    destino.parent.mkdir(parents=True, exist_ok=True)
    with zipfile.ZipFile(destino, "w") as z:
        for tab in g.TABELAS_CVM:
            campos = sorted({k for v in fundos.values() for k in v if re.match(rf"TAB_{tab}(_|\d)", k)})
            linhas = ["TP_FUNDO_CLASSE;CNPJ_FUNDO_CLASSE;DENOM_SOCIAL;DT_COMPTC;" + ";".join(campos)]
            for cnpj, v in fundos.items():
                linhas.append(f"Classe;{cnpj};FUNDO FICTÍCIO;2026-09-30;"
                              + ";".join(f"{v.get(c, 0):.2f}" for c in campos))
            z.writestr(f"inf_mensal_fidc_tab_{tab}_202609.csv", "\n".join(linhas).encode("latin-1"))


def montar_entrada(pasta: Path) -> Path:
    """Bases fictícias: ALFA completo, BETA sem Pipefy, GAMA sem carteira."""
    entrada = pasta / "entrada"
    entrada.mkdir()
    _xlsx(entrada / "fundos.xlsx", [["Nome Fundo", "CNPJ"],
                                    ["FUNDO ALFA FIDC", ALFA],
                                    ["FUNDO BETA FIDC", BETA],
                                    ["FUNDO GAMA FIDC", GAMA],
                                    ["NOME REPETIDO IGNORADO", ALFA]])
    carteira = [["Data Base", "CNPJ Fundo", "Descricao", "Valor"],
                [DATA_BASE, ALFA, "Disponibilidades", 1000.0],
                [DATA_BASE, ALFA, "Direitos Creditorios COM a Vencer", 8000.0],
                [DATA_BASE, "11111111000111", "Direitos Creditorios Vencidos", 500.0],   # CNPJ sem pontuação
                [DATA_BASE, ALFA, "Provisao para Devedores Duvidosos - PDD", -500.0],
                [DATA_BASE, ALFA, "Titulos Publicos Federais", 1500.0],
                [DATA_BASE, ALFA, "Valores a Pagar", -500.0],
                [DATA_BASE, ALFA, "Patrimonio Liquido da Subclasse Senior", 10000.0],
                [dt.date(2026, 6, 30), ALFA, "Disponibilidades", 999999.0],               # outra data base
                [DATA_BASE, BETA, "Disponibilidades", 100.0],
                [DATA_BASE, BETA, "Patrimonio Liquido da Subclasse Senior", 100.0]]
    _xlsx(entrada / "carteira_informe.xlsx", carteira)
    _xlsx(entrada / "resumo_sinqia.xlsx", [
        ["SITUACAO_RECEBIVEL", "DOC_FUNDO", "NOME_FUNDO", "RELATORIO", "QTDE", "VALOR_NOMINAL", "VALOR_PRESENTE"],
        ["A vencer", 1111111100011, "ALFA", "Estoque", 30, 3000.0, 2700.0],
        ["Vencido", 1111111100011, "ALFA", "Estoque", 10, 1000.0, 900.0],
        ["A vencer", 1111111100011, "ALFA", "Liquidacao", 5, 500.0, 450.0],
        ["A vencer", 1111111100011, "ALFA", "Aquisicao", 8, 800.0, 720.0],
        ["A vencer", "TEXTO-INVALIDO", "OUTRO", "Aquisicao", 1, 1.0, 1.0]])
    campos = ["Title", "CNPJ do Fundo"] + [f"Descrição item {i}" for i in
                                           ("1", "2.a", "2.b", "3", "4.a", "4.b", "5", "6.a", "6.b", "7", "8", "9")]
    resposta = {c: f"Resposta {c}" for c in campos[2:]}
    resposta["Descrição item 1"] = ""                       # campo vazio -> texto padrão
    resposta["Descrição item 9"] = "Processo & <teste>\n\nsegunda linha\n"
    _xlsx(entrada / "pipefy.xlsx", [campos,
                                     ["ALFA CARD", ALFA] + [resposta[c] for c in campos[2:]],
                                     ["ALFA CARD ANTIGO", ALFA] + ["ignorado"] * 12])
    publicado = pasta / "cvm_publicado"
    montar_informe_cvm(publicado / "inf_mensal_fidc_202609.zip", INFORME_OK)
    cfg = json.loads((Path(g.BASE_DIR) / "config.json").read_text(encoding="utf-8"))
    cfg["data_base"] = DATA_BASE.isoformat()
    cfg["pastas"] = {"entrada": str(entrada), "saida": str(pasta / "saida"), "validados": str(pasta / "validados")}
    cfg["cvm"] = {"url": publicado.as_uri() + "/inf_mensal_fidc_{aaaamm}.zip", "tolerancia": 1.0}
    (pasta / "config.json").write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
    return pasta / "config.json"


class Base(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        cls.tmp = Path(tempfile.mkdtemp(prefix="geradora_teste_"))
        cls.config = montar_entrada(cls.tmp)
        cls.cfg = g.ler_config(cls.config)
        cls.bases = g.carregar_bases(cls.cfg)
        cls.modelo = g.carregar_modelo()
        cls.rotulos = g.rotulos_composicao(cls.modelo.ws)

    @classmethod
    def tearDownClass(cls):
        shutil.rmtree(cls.tmp, ignore_errors=True)

    def calcular(self, cnpj):
        return g.calcular(self.bases, cnpj, DATA_BASE, self.rotulos, TEXTO_PADRAO)


class TestRegras(Base):
    def test_mapeamento_usa_a_primeira_chave_da_tabela(self):
        m = self.bases.mapeamento
        self.assertEqual(g.mapear_descricao("Direitos Creditorios COM a Vencer", m), "Operações de Crédito")
        self.assertEqual(g.mapear_descricao("Provisao - PDD", m), "Provisão para Perdas (PDD)")
        self.assertEqual(g.mapear_descricao("Titulos Publicos Federais", m), "Títulos Públicos Federais")
        self.assertEqual(g.mapear_descricao("Posicao em Derivativos - Passivo", m), g.FORA_DO_RELATORIO)
        self.assertEqual(g.mapear_descricao("Algo desconhecido", m), "")

    def test_composicao_da_carteira(self):
        c = self.calcular(ALFA).celulas
        self.assertEqual(c["C13"], 1000.0)
        self.assertEqual(c["C14"], 1500.0)
        self.assertEqual(c["C15"], 8500.0)         # a vencer + vencidos, inclusive CNPJ sem pontuação
        self.assertEqual(c["C16"], -500.0)
        self.assertEqual(c["C19"], -500.0)
        self.assertEqual(c["C20"], 10000.0)        # só a data base pedida
        self.assertAlmostEqual(c["D15"], 0.85)
        self.assertAlmostEqual(c["D20"], 1.0)
        self.assertEqual(c["B5"], "3º Trimestre/2026")
        self.assertEqual(c["A10"], "A seguir segue a composição da Carteira em 30/09/2026")

    def test_lastro_sinqia(self):
        c = self.calcular(ALFA).celulas
        self.assertEqual((c["B27"], c["C27"], c["D27"]), (43, 4300.0, 3870.0))
        self.assertEqual((c["B28"], c["D28"]), (10, 900.0))
        self.assertEqual(c["B29"], 53)
        self.assertEqual((c["B30"], c["B31"]), (5, 8))
        self.assertAlmostEqual(c["E27"], 43 / 53)
        self.assertEqual(c["F33"], 0)              # sem validação; a fórmula divide por D27

    def test_textos_do_gestor(self):
        c = self.calcular(ALFA).celulas
        self.assertEqual(c["I5"], "ALFA CARD")     # primeira linha do Pipefy com o CNPJ
        self.assertEqual(c["A59"], TEXTO_PADRAO)   # campo vazio
        self.assertEqual(c["A66"], "Resposta Descrição item 2.a")
        self.assertTrue(c["A52"].startswith("Processo & <teste>"))
        beta = self.calcular(BETA)
        self.assertTrue(all(beta.celulas[f"A{l}"] == TEXTO_PADRAO for l in g.TEXTOS_GESTOR))
        self.assertIn("Sem resposta do gestor no Pipefy: textos padrão usados", beta.avisos)

    def test_nome_e_quem_tem_carteira(self):
        self.assertEqual(g.nome_do_fundo(self.bases, "11111111000111"), "FUNDO ALFA FIDC")
        self.assertTrue(g.tem_carteira(self.bases, ALFA))
        self.assertFalse(g.tem_carteira(self.bases, GAMA))
        self.assertTrue(any("DOC_FUNDO" in a for a in self.bases.avisos_gerais))

    def test_nome_de_arquivo(self):
        rel = self.calcular(ALFA)
        self.assertEqual(rel.nome_arquivo, "FUNDO ALFA FIDC_11111111000111_TRIMESTRAL_CVM")
        rel.nome = "NOME / COM : INVALIDOS? " * 20
        self.assertLessEqual(len(rel.nome_arquivo), 200)
        self.assertTrue(rel.nome_arquivo.endswith("_11111111000111_TRIMESTRAL_CVM"))
        self.assertFalse(set('/\\:*?"<>|') & set(rel.nome_arquivo))


class TestFormatos(unittest.TestCase):
    CONTABIL = '_(* #,##0.00_);_(* \\(#,##0.00\\);_(* "-"??_);_(@_)'

    def test_contabil(self):
        self.assertEqual(g.formatar_numero(441014.18, self.CONTABIL), ("441.014,18", ")"))
        self.assertEqual(g.formatar_numero(-8150110.41, self.CONTABIL), ("(8.150.110,41)", ""))
        self.assertEqual(g.formatar_numero(0, self.CONTABIL), ("-", "00)"))
        self.assertEqual(g.formatar_numero(0.001, self.CONTABIL), ("-", "00)"))

    def test_percentual_sem_menos_zero(self):
        self.assertEqual(g.formatar_numero(0.000953, "0.00%")[0], "0,10%")
        self.assertEqual(g.formatar_numero(-0.0176, "0.00%")[0], "-1,76%")
        self.assertEqual(g.formatar_numero(-0.0000001, "0.00%")[0], "0,00%")
        self.assertEqual(g.formatar_numero(1.0000000000000809, "0.00%")[0], "100,00%")

    def test_geral_e_inteiro(self):
        self.assertEqual(g.formatar_numero(43, '_(* #,##0_);_(* \\(#,##0\\);_(* "-"??_);_(@_)')[0], "43")
        self.assertEqual(g.formatar_numero(1.5, "General")[0], "1,5")
        self.assertEqual(g.formatar_numero(46295, "mm-dd-yy")[0], "30/09/2026")

    def test_numero_csv(self):
        self.assertEqual(g.numero_csv("8280.48"), 8280.48)
        self.assertEqual(g.numero_csv("1.234,56"), 1234.56)
        self.assertEqual(g.numero_csv("1,5"), 1.5)
        self.assertEqual(g.numero_csv("1332733400011"), 1332733400011)
        self.assertEqual(g.numero_csv("-500"), -500)
        for texto in ("01234567000189", "48.169.685/0001-20", "30/09/2026", "2026-09-30", "abc", ""):
            self.assertEqual(g.numero_csv(texto), texto)


class TestSaidas(Base):
    def test_xlsx_preenche_o_modelo_do_excel(self):
        rel = self.calcular(ALFA)
        dados = self.modelo.xlsx(rel.celulas, mostrar_lastro=False)
        nomes = zipfile.ZipFile(io.BytesIO(dados)).namelist()
        self.assertIn("xl/media/image1.png", nomes)            # logo do cabeçalho de impressão
        self.assertFalse([n for n in nomes if "externalLink" in n or "webextension" in n])
        ws = openpyxl.load_workbook(io.BytesIO(dados), rich_text=True).worksheets[0]
        self.assertEqual(ws["B3"].value, "FUNDO ALFA FIDC")
        self.assertEqual(ws["B6"].value.date(), DATA_BASE)
        self.assertEqual(ws["C20"].value, 10000)
        self.assertEqual(ws["A52"].value, "Processo & <teste>\n\nsegunda linha\n")
        self.assertEqual(ws["C13"].number_format, self.modelo.ws["C13"].number_format)
        self.assertEqual(ws.oddHeader.left.text, "&G")
        self.assertTrue(all(ws.row_dimensions[r].hidden for r in g.LINHAS_SECAO_LASTRO))
        com_lastro = openpyxl.load_workbook(io.BytesIO(self.modelo.xlsx(rel.celulas, True))).worksheets[0]
        self.assertFalse(any(com_lastro.row_dimensions[r].hidden for r in g.LINHAS_SECAO_LASTRO))

    def test_xlsx_cria_celula_que_o_modelo_nao_tem(self):
        dados = self.modelo.xlsx({"Z200": "novo", "AA3": 7}, mostrar_lastro=False)
        ws = openpyxl.load_workbook(io.BytesIO(dados)).worksheets[0]
        self.assertEqual((ws["Z200"].value, ws["AA3"].value), ("novo", 7))

    def test_pdf(self):
        for cnpj in (ALFA, BETA):
            destino = self.tmp / f"teste_{g.digitos(cnpj)}.pdf"
            g.salvar_pdf(self.calcular(cnpj), destino, False, self.modelo)
            dados = destino.read_bytes()
            self.assertTrue(dados.startswith(b"%PDF"))
            self.assertEqual(len(re.findall(rb"/Type /Page\b(?!s)", dados)), 2)   # 2 páginas, sem página vazia
        g.salvar_pdf(self.calcular(ALFA), self.tmp / "com_lastro.pdf", True, self.modelo)
        self.assertGreater((self.tmp / "com_lastro.pdf").stat().st_size, 0)


class TestComandoGerar(Base):
    def gerar(self, *extra):
        saida = io.StringIO()
        with redirect_stdout(saida):
            try:
                g.main(["--config", str(self.config), "gerar", *extra])
                codigo = 0
            except SystemExit as e:
                codigo = e.code
        return codigo, saida.getvalue()

    def test_gera_quem_tem_carteira(self):
        codigo, log = self.gerar()
        self.assertEqual(codigo, 0, log)
        pasta = self.tmp / "saida"
        arquivos = sorted(p.name for p in (pasta / OK_DIR).iterdir())
        self.assertIn("FUNDO ALFA FIDC_11111111000111_TRIMESTRAL_CVM.pdf", arquivos)
        self.assertFalse([a for a in arquivos if a.endswith(".xlsx")])       # Excel vai para Insumos
        insumos = sorted(p.name for p in (pasta / INS_DIR).iterdir())
        self.assertEqual(insumos, ["FUNDO ALFA FIDC_11111111000111_TRIMESTRAL_CVM.xlsx",
                                   "FUNDO BETA FIDC_22222222000122_TRIMESTRAL_CVM.xlsx"])
        self.assertFalse([a for a in arquivos if "GAMA" in a])
        self.assertIn("1 fundo(s) sem carteira", log)
        self.assertIn("1 CNPJ(s) repetido(s)", log)
        resumo = (pasta / "resumo_geracao.csv").read_text(encoding="utf-8-sig").splitlines()
        self.assertEqual([l.split(";")[1] for l in resumo[1:]], [ALFA, BETA])  # ordem da lista, sem repetição

    def test_gerar_um_cnpj_mantem_os_outros_no_resumo(self):
        saida = self.tmp / "saida_parcial"
        self.gerar("--saida", str(saida))
        codigo, _ = self.gerar("--saida", str(saida), "--cnpj", BETA)
        self.assertEqual(codigo, 0)
        resumo = (saida / "resumo_geracao.csv").read_text(encoding="utf-8-sig").splitlines()
        self.assertEqual(sorted(l.split(";")[1] for l in resumo[1:]), [ALFA, BETA])

    def test_um_fundo_com_erro_nao_para_o_lote(self):
        bloqueio = self.tmp / "saida_erro" / OK_DIR / "FUNDO ALFA FIDC_11111111000111_TRIMESTRAL_CVM.pdf"
        bloqueio.mkdir(parents=True)          # uma pasta com o nome do PDF impede a gravação
        codigo, log = self.gerar("--saida", str(self.tmp / "saida_erro"))
        self.assertEqual(codigo, 1)
        self.assertIn("[ERRO] FUNDO ALFA FIDC", log)
        self.assertTrue((self.tmp / "saida_erro" / OK_DIR / "FUNDO BETA FIDC_22222222000122_TRIMESTRAL_CVM.pdf").exists())
        resumo = (self.tmp / "saida_erro" / "resumo_geracao.csv").read_text(encoding="utf-8-sig")
        self.assertIn("ERRO ao gerar", resumo)

    def test_avisa_fundo_da_carteira_fora_da_lista(self):
        pasta = self.tmp / "fora"
        shutil.copytree(self.tmp / "entrada", pasta / "entrada")
        delta = "44.444.444/0001-44"
        for arquivo, linha in (("carteira_informe.xlsx", [DATA_BASE, "44444444000144", "Disponibilidades", 50.0]),
                               ("pipefy.xlsx", ["DELTA CARD", delta])):
            wb = openpyxl.load_workbook(pasta / "entrada" / arquivo)
            wb.active.append(linha)
            wb.save(pasta / "entrada" / arquivo)
        cfg = json.loads(self.config.read_text(encoding="utf-8"))
        cfg["pastas"] = {"entrada": str(pasta / "entrada"), "saida": str(pasta / "saida"),
                         "validados": str(pasta / "validados")}
        config = pasta / "config.json"
        config.write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")

        self.assertEqual(g.fundos_fora_da_lista(g.carregar_bases(g.ler_config(config))), [(delta, "DELTA CARD")])
        saida = io.StringIO()
        with redirect_stdout(saida):
            g.main(["--config", str(config), "gerar", "--formato", "xlsx"])
        log = saida.getvalue()
        self.assertIn(f"[AVISO] 1 fundo(s) na Carteira Informe que não estão na lista de fundos", log)
        self.assertIn(f"{delta} DELTA CARD", log)
        self.assertIn("concluída: 2 fundo(s)", log)
        self.assertFalse([p for p in (pasta / "saida").rglob("*") if "44444444000144" in p.name])
        resumo = (pasta / "saida" / "resumo_geracao.csv").read_text(encoding="utf-8-sig")
        self.assertIn(f"DELTA CARD;{delta};;;;;NÃO GERADO", resumo)  # pl, arquivo, conferencia_cvm e pasta vazios

        # gerando só CNPJs escolhidos, não avisa
        saida = io.StringIO()
        with redirect_stdout(saida):
            g.main(["--config", str(config), "gerar", "--formato", "xlsx", "--cnpj", ALFA])
        self.assertNotIn("não estão na lista", saida.getvalue())

    def test_entrada_em_csv(self):
        pasta = self.tmp / "csv"
        shutil.copytree(self.tmp / "entrada", pasta / "entrada")
        wb = openpyxl.load_workbook(pasta / "entrada" / "carteira_informe.xlsx")
        with open(pasta / "entrada" / "carteira_informe.csv", "w", encoding="utf-8-sig", newline="") as f:
            for linha in wb.active.iter_rows(values_only=True):
                f.write(";".join(v.strftime("%d/%m/%Y") if hasattr(v, "strftime") else str(v).replace(".", ",")
                                 for v in linha) + "\n")
        cfg = dict(self.cfg)
        cfg["arquivos"] = dict(cfg["arquivos"], carteira_informe="carteira_informe.csv")
        cfg["pastas"] = dict(cfg["pastas"], entrada=str(pasta / "entrada"))
        bases = g.carregar_bases(cfg)
        rel = g.calcular(bases, ALFA, DATA_BASE, self.rotulos, TEXTO_PADRAO)
        self.assertEqual(rel.celulas["C20"], 10000.0)
        self.assertEqual(rel.celulas["C15"], 8500.0)


class TestRespostasManuais(Base):
    def test_respostas_manuais_substituem_pipefy_e_textos_fixos(self):
        pasta = self.tmp / "manuais"
        shutil.copytree(self.tmp / "entrada", pasta / "entrada")
        _xlsx(pasta / "entrada" / "respostas_gestor.xlsx", [
            ["CNPJ do Fundo", "Fundo", "V.a - Política de investimento", "V.d.ii - Caráter da cessão",
             "V.g - Descontinuidade", "V.c - Garantias"],
            ["22222222000122", "BETA", "Nova política", "Com coobrigação", "Sem efeito relevante", ""]])
        cfg = dict(self.cfg)
        cfg["pastas"] = dict(cfg["pastas"], entrada=str(pasta / "entrada"))
        bases = g.carregar_bases(cfg)
        beta = g.calcular(bases, BETA, DATA_BASE, self.rotulos, TEXTO_PADRAO)
        self.assertEqual(beta.celulas["A59"], "Nova política")
        self.assertEqual(beta.celulas["A89"], "Com coobrigação")       # texto fixo do modelo
        self.assertEqual(beta.celulas["A106"], "Sem efeito relevante")  # texto fixo do modelo
        self.assertEqual(beta.celulas["A77"], TEXTO_PADRAO)             # coluna vazia: mantém o Pipefy/padrão
        self.assertFalse(any(a.startswith("Sem resposta do gestor") for a in beta.avisos))
        self.assertTrue(any("respostas_gestor.xlsx" in a for a in beta.avisos))
        alfa = g.calcular(bases, ALFA, DATA_BASE, self.rotulos, TEXTO_PADRAO)
        self.assertNotIn("A89", alfa.celulas)                           # outros fundos não mudam


class TestAtualizacoesPipefy(Base):
    def bases_com(self, linhas):
        pasta = Path(tempfile.mkdtemp(dir=self.tmp))
        shutil.copytree(self.tmp / "entrada", pasta / "entrada")
        _xlsx(pasta / "entrada" / "pipefy_atualizacoes.xlsx",
              [["Título", "Fase atual", "Nome do Fundo", "CNPJ do Fundo", "Descrição item 1", "Descrição item 3"]]
              + linhas)
        cfg = dict(self.cfg)
        cfg["pastas"] = dict(cfg["pastas"], entrada=str(pasta / "entrada"))
        return g.carregar_bases(cfg)

    def test_so_muda_os_fundos_e_as_colunas_da_extracao(self):
        bases = self.bases_com([["ALFA", "Jurídico", "FUNDO ALFA FIDC", ALFA, "Política nova", None]])
        alfa = g.calcular(bases, ALFA, DATA_BASE, self.rotulos, TEXTO_PADRAO)
        original = self.calcular(ALFA)
        self.assertEqual(alfa.celulas["A59"], "Política nova")            # item 1 veio da extração
        for celula in ("A52", "A66", "A72", "A77", "A84", "A94", "A100", "A112"):
            self.assertEqual(alfa.celulas[celula], original.celulas[celula])  # o resto fica como na base
        self.assertTrue(any("atualizados por pipefy_atualizacoes.xlsx" in a for a in alfa.avisos))
        beta = g.calcular(bases, BETA, DATA_BASE, self.rotulos, TEXTO_PADRAO)
        self.assertEqual(beta.celulas, self.calcular(BETA).celulas)       # fundo fora da extração não muda

    def test_gerar_separa_os_pdfs_alterados_pela_planilha(self):
        pasta = self.tmp / "alterados"
        shutil.copytree(self.tmp / "entrada", pasta / "entrada")
        _xlsx(pasta / "entrada" / "pipefy_atualizacoes.xlsx",
              [["Título", "CNPJ do Fundo", "Descrição item 1"], ["ALFA", ALFA, "Política nova"]])
        cfg = json.loads(self.config.read_text(encoding="utf-8"))
        cfg["pastas"] = dict(cfg["pastas"], entrada=str(pasta / "entrada"), saida=str(pasta / "saida"))
        config = pasta / "config.json"
        config.write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
        with redirect_stdout(io.StringIO()):
            g.main(["--config", str(config), "gerar"])
        ok = pasta / "saida" / OK_DIR
        alterados = ok / g.SUBPASTAS_PADRAO["revisados_atualizados"]
        self.assertEqual([p.name for p in alterados.iterdir()], ["FUNDO ALFA FIDC_11111111000111_TRIMESTRAL_CVM.pdf"])
        self.assertTrue((ok / "FUNDO BETA FIDC_22222222000122_TRIMESTRAL_CVM.pdf").exists())
        self.assertFalse((ok / "FUNDO ALFA FIDC_11111111000111_TRIMESTRAL_CVM.pdf").exists())

    def test_dois_cards_do_mesmo_fundo_vale_o_mais_recente(self):
        pasta = Path(tempfile.mkdtemp(dir=self.tmp))
        shutil.copytree(self.tmp / "entrada", pasta / "entrada")
        _xlsx(pasta / "entrada" / "pipefy_atualizacoes.xlsx",
              [["Título", "Criado em", "Nome do Fundo", "CNPJ do Fundo", "Descrição item 1"],
               ["ALFA", dt.datetime(2026, 10, 1, 9, 0), "FUNDO ALFA FIDC", ALFA, "Resposta antiga"],
               ["ALFA", dt.datetime(2026, 11, 5, 9, 0), "FUNDO ALFA FIDC", ALFA, "Resposta nova"],
               ["ALFA", dt.datetime(2026, 10, 20, 9, 0), "FUNDO ALFA FIDC", ALFA, "Resposta do meio"]])
        cfg = dict(self.cfg)
        cfg["pastas"] = dict(cfg["pastas"], entrada=str(pasta / "entrada"))
        bases = g.carregar_bases(cfg)
        self.assertEqual(g.calcular(bases, ALFA, DATA_BASE, self.rotulos, TEXTO_PADRAO).celulas["A59"], "Resposta nova")
        self.assertTrue(any("vale o criado por último" in a for a in bases.avisos_gerais))

    def test_fundo_sem_card_na_base_ganha_os_textos(self):
        bases = self.bases_com([["BETA", "Jurídico", "FUNDO BETA FIDC", BETA, "Política do BETA", "Sem garantias"]])
        beta = g.calcular(bases, BETA, DATA_BASE, self.rotulos, TEXTO_PADRAO)
        self.assertEqual(beta.celulas["A59"], "Política do BETA")
        self.assertEqual(beta.celulas["A77"], "Sem garantias")
        self.assertEqual(beta.celulas["A66"], TEXTO_PADRAO)               # coluna que a extração não traz

    def test_card_com_cnpj_de_outro_fundo_vai_pelo_nome(self):
        bases = self.bases_com([["ALFA", "Jurídico", "FUNDO ALFA FIDC", ALFA, "Do ALFA", None],
                                ["BETA", "Jurídico", "FUNDO BETA FIDC", ALFA, "Do BETA", None]])
        self.assertEqual(g.calcular(bases, ALFA, DATA_BASE, self.rotulos, TEXTO_PADRAO).celulas["A59"], "Do ALFA")
        self.assertEqual(g.calcular(bases, BETA, DATA_BASE, self.rotulos, TEXTO_PADRAO).celulas["A59"], "Do BETA")
        self.assertTrue(any("CNPJ de outro fundo" in a for a in bases.avisos_gerais))


class TestLastroManual(Base):
    def test_lastro_manual_preenche_e_exibe_a_secao_ii_so_no_fundo(self):
        pasta = self.tmp / "lastro_manual"
        shutil.copytree(self.tmp / "entrada", pasta / "entrada")
        colunas = ["CNPJ do Fundo"] + [f"{p} - {s}" for p in g.LASTRO_MANUAL for s in ("Qtde", "Nominal", "Presente")]
        _xlsx(pasta / "entrada" / "lastro_manual.xlsx",
              [colunas, [BETA, 90, 900.0, 800.0, 10, 100.0, 90.0, 0, 0, 0, 0, 0, 0, 0, 0, 0]])
        cfg = dict(self.cfg)
        cfg["pastas"] = dict(cfg["pastas"], entrada=str(pasta / "entrada"))
        bases = g.carregar_bases(cfg)
        beta = g.calcular(bases, BETA, DATA_BASE, self.rotulos, TEXTO_PADRAO)
        c = beta.celulas
        self.assertEqual((c["B29"], c["C29"], c["D29"]), (100, 1000.0, 890.0))   # total = soma das linhas 32 a 36
        self.assertAlmostEqual(c["E32"], 0.9)
        self.assertAlmostEqual(c["F33"], 0.1)                                    # % sobre o valor nominal
        self.assertAlmostEqual(c["E29"], 1.0)
        self.assertEqual(beta.linhas_lastro, frozenset(g.LINHAS_SECAO_II))
        alfa = g.calcular(bases, ALFA, DATA_BASE, self.rotulos, TEXTO_PADRAO)
        self.assertEqual(alfa.linhas_lastro, frozenset())
        # no XLSX, a seção II aparece no BETA e continua oculta no ALFA; a seção III fica oculta nos dois
        for rel, visivel in ((beta, True), (alfa, False)):
            destino = pasta / f"{rel.nome_arquivo}.xlsx"
            g.salvar_xlsx(rel, destino, False, self.modelo)
            ws = openpyxl.load_workbook(destino).active
            self.assertEqual(not ws.row_dimensions[29].hidden, visivel)
            self.assertTrue(ws.row_dimensions[44].hidden)
        g.salvar_pdf(beta, pasta / "beta.pdf", False, self.modelo)
        self.assertTrue((pasta / "beta.pdf").stat().st_size > 1000)


class TestConferenciaCVM(Base):
    def informe(self, **mudancas):
        reg = {k: str(v) for k, v in INFORME_OK[ALFA].items()}
        reg.update({k: str(v) for k, v in mudancas.items()})
        return {g.digitos(ALFA): reg}

    def test_bate_com_a_cvm(self):
        rel = self.calcular(ALFA)
        conf = g.conferir_cvm(rel, self.informe())
        self.assertEqual(conf.status, g.STATUS_OK)
        self.assertFalse(conf.revisar)

    def test_corrige_com_a_cvm_quando_o_informe_fecha(self):
        rel = self.calcular(ALFA)
        # CVM: PL 9.800 e contas a pagar 700 (ativo 10.500 - passivo 700 = 9.800)
        conf = g.conferir_cvm(rel, self.informe(TAB_III_VL_PASSIVO=700, TAB_III_A_VL_PAGAR=700,
                                                TAB_IV_A_VL_PL=9800))
        self.assertEqual(conf.status, g.STATUS_CORRIGIDO)
        self.assertEqual({d[0] for d in conf.divergencias}, {19, 20})
        self.assertEqual(rel.celulas["C20"], 9800)
        self.assertEqual(rel.celulas["C19"], -700)
        self.assertAlmostEqual(rel.celulas["D13"], 1000 / 9800)
        self.assertAlmostEqual(rel.celulas["D20"], 1.0)
        self.assertTrue(any("corrigido com o informe da CVM" in a for a in rel.avisos))

    def test_nao_corrige_quando_o_informe_nao_fecha(self):
        rel = self.calcular(ALFA)
        conf = g.conferir_cvm(rel, self.informe(TAB_IV_A_VL_PL=0))
        self.assertEqual(conf.status, g.STATUS_INCONSISTENTE)
        self.assertEqual(rel.celulas["C20"], 10000.0)   # mantém a carteira

    def test_nao_corrige_com_informe_zerado(self):
        rel = self.calcular(ALFA)
        zerado = {k: 0 for k in INFORME_OK[ALFA]}
        conf = g.conferir_cvm(rel, self.informe(**zerado))
        self.assertEqual(conf.status, g.STATUS_INCONSISTENTE)
        self.assertEqual(rel.celulas["C20"], 10000.0)

    def test_fundo_sem_informe(self):
        rel = self.calcular(BETA)
        conf = g.conferir_cvm(rel, self.informe())
        self.assertEqual(conf.status, g.STATUS_SEM_INFORME)
        self.assertTrue(conf.revisar)

    def test_outros_ativos_sem_derivativos_e_compromissadas(self):
        v = g.valores_cvm({"TAB_I_VL_ATIVO": "1000", "TAB_I1_VL_DISP": "100", "TAB_I2F_VL_OPER_COMPROM": "50",
                           "TAB_I3_VL_POSICAO_DERIV": "300", "TAB_I2C5_VL_COTA_FIF": "550"})
        self.assertEqual(v[17], 550)

    def test_gerar_separa_os_fundos_para_revisao(self):
        pasta = self.tmp / "revisao"
        publicado = pasta / "publicado"
        divergente = dict(INFORME_OK, **{BETA: {"TAB_I_VL_ATIVO": 100, "TAB_I1_VL_DISP": 100,
                                                "TAB_IV_A_VL_PL": 90, "TAB_III_VL_PASSIVO": 10,
                                                "TAB_III_A_VL_PAGAR": 10}})
        montar_informe_cvm(publicado / "inf_mensal_fidc_202609.zip", divergente)
        cfg = json.loads(self.config.read_text(encoding="utf-8"))
        cfg["pastas"] = dict(cfg["pastas"], saida=str(pasta / "saida"))
        cfg["cvm"]["url"] = publicado.as_uri() + "/inf_mensal_fidc_{aaaamm}.zip"
        config = pasta / "config.json"
        config.write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
        # um BETA antigo na raiz do _DRAFT deve sair de lá
        (pasta / "saida").mkdir(parents=True)
        (pasta / "saida" / "FUNDO BETA FIDC_22222222000122_TRIMESTRAL_CVM.pdf").write_bytes(b"antigo")
        log = io.StringIO()
        with redirect_stdout(log):
            g.main(["--config", str(config), "gerar"])
        raiz = sorted(p.name for p in (pasta / "saida").iterdir())
        self.assertEqual(raiz, sorted([OK_DIR, INS_DIR, REV_DIR, "resumo_geracao.csv"]))
        self.assertTrue((pasta / "saida" / INS_DIR / "FUNDO ALFA FIDC_11111111000111_TRIMESTRAL_CVM.xlsx").exists())
        self.assertTrue((pasta / "saida" / INS_DIR / "FUNDO BETA FIDC_22222222000122_TRIMESTRAL_CVM.xlsx").exists())
        saida = sorted(p.name for p in (pasta / "saida" / OK_DIR).iterdir())
        rev = sorted(p.name for p in (pasta / "saida" / REV_DIR).iterdir())
        self.assertIn("FUNDO ALFA FIDC_11111111000111_TRIMESTRAL_CVM.pdf", saida)
        self.assertFalse([a for a in saida if "BETA" in a])
        self.assertIn("FUNDO BETA FIDC_22222222000122_TRIMESTRAL_CVM.pdf", rev)
        self.assertFalse([a for a in rev if a.endswith(".xlsx")])          # o Excel (draft) fica separado
        self.assertIn("[REVISAR] FUNDO BETA", log.getvalue())
        csv_rev = (pasta / "saida" / REV_DIR / "revisao_cvm.csv").read_text(encoding="utf-8-sig")
        self.assertIn(g.STATUS_CORRIGIDO, csv_rev)
        self.assertIn("Patrimônio Líquido", csv_rev)
        ws = openpyxl.load_workbook(pasta / "saida" / INS_DIR / "FUNDO BETA FIDC_22222222000122_TRIMESTRAL_CVM.xlsx").active
        self.assertEqual(ws["C20"].value, 90)
        self.assertEqual(ws["C19"].value, -10)

    def test_sem_cvm(self):
        pasta = self.tmp / "sem_cvm"
        cfg = json.loads(self.config.read_text(encoding="utf-8"))
        cfg["pastas"] = dict(cfg["pastas"], saida=str(pasta / "saida"))
        cfg["cvm"]["url"] = (pasta / "nao_existe").as_uri() + "/{aaaamm}.zip"
        pasta.mkdir()
        config = pasta / "config.json"
        config.write_text(json.dumps(cfg, ensure_ascii=False), encoding="utf-8")
        with redirect_stdout(io.StringIO()):
            g.main(["--config", str(config), "gerar", "--sem-cvm"])
        self.assertTrue((pasta / "saida" / "FUNDO BETA FIDC_22222222000122_TRIMESTRAL_CVM.pdf").exists())
        self.assertFalse((pasta / "saida" / REV_DIR).exists())


if __name__ == "__main__":
    unittest.main()
