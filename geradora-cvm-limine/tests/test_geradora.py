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

ALFA, BETA, GAMA = "11.111.111/0001-11", "22.222.222/0001-22", "33.333.333/0001-33"


def _xlsx(caminho: Path, linhas: list[list]):
    wb = openpyxl.Workbook()
    for linha in linhas:
        wb.active.append(linha)
    wb.save(caminho)


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
    cfg = json.loads((Path(g.BASE_DIR) / "config.json").read_text(encoding="utf-8"))
    cfg["data_base"] = DATA_BASE.isoformat()
    cfg["pastas"] = {"entrada": str(entrada), "saida": str(pasta / "saida"), "validados": str(pasta / "validados")}
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
        self.assertEqual(g.mapear_descricao("Posicao em Derivativos - Passivo", m), "Patrimônio Líquido")
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
        arquivos = sorted(p.name for p in pasta.iterdir())
        self.assertIn("FUNDO ALFA FIDC_11111111000111_TRIMESTRAL_CVM.pdf", arquivos)
        self.assertIn("FUNDO BETA FIDC_22222222000122_TRIMESTRAL_CVM.xlsx", arquivos)
        self.assertFalse([a for a in arquivos if "GAMA" in a])
        self.assertIn("1 fundo(s) sem carteira", log)
        self.assertIn("1 CNPJ(s) repetido(s)", log)
        resumo = (pasta / "resumo_geracao.csv").read_text(encoding="utf-8-sig").splitlines()
        self.assertEqual([l.split(";")[1] for l in resumo[1:]], [ALFA, BETA])  # ordem da lista, sem repetição

    def test_um_fundo_com_erro_nao_para_o_lote(self):
        bloqueio = self.tmp / "saida_erro" / "FUNDO ALFA FIDC_11111111000111_TRIMESTRAL_CVM.pdf"
        bloqueio.mkdir(parents=True)          # uma pasta com o nome do PDF impede a gravação
        codigo, log = self.gerar("--saida", str(self.tmp / "saida_erro"))
        self.assertEqual(codigo, 1)
        self.assertIn("[ERRO] FUNDO ALFA FIDC", log)
        self.assertTrue((self.tmp / "saida_erro" / "FUNDO BETA FIDC_22222222000122_TRIMESTRAL_CVM.pdf").exists())
        resumo = (self.tmp / "saida_erro" / "resumo_geracao.csv").read_text(encoding="utf-8-sig")
        self.assertIn("ERRO ao gerar", resumo)

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


if __name__ == "__main__":
    unittest.main()
