# Geradora CVM Limine (Python)

Substitui a planilha `Geradora Limine - EMISSÃO EM LOTE.xlsb`. Gera o **Demonstrativo
Trimestral** (PDF + XLSX) de cada fundo da Limine sem abrir o Excel, em ~8 segundos para os
179 fundos do 3T2026.

## Instalação

```bash
pip install -r requirements.txt
```

## Uso no trimestre

1. Atualize as bases em `entrada/` (ou rode o passo 0 abaixo):
   - `fundos.xlsx`: lista de fundos, com as colunas `Nome Fundo` e `CNPJ` (antiga aba Setup, colunas G:H)
   - `carteira_informe.xlsx`: colunas `Data Base`, `CNPJ Fundo`, `Descricao` e `Valor` (antiga aba CarteiraInforme)
   - `resumo_sinqia.xlsx`: colunas `SITUACAO_RECEBIVEL`, `DOC_FUNDO`, `RELATORIO`, `QTDE`,
     `VALOR_NOMINAL` e `VALOR_PRESENTE` (antiga aba RESUMO_SINQIA)
   - `pipefy.xlsx`: respostas dos gestores (passo 2)
   - Também aceita `.csv` (separador `;` ou `,`; números como `1234.56`, `1234,56` ou `1.234,56`)
     ou `.xlsb`: basta trocar o nome em `config.json` → `arquivos`.
2. Ajuste `data_base` e as pastas em `config.json`.
3. Baixe as respostas dos gestores no Pipefy (antigo botão **Atualizar Pipefy**):
   ```bash
   set PIPEFY_TOKEN=<token da API>
   python geradora_cvm.py pipefy
   ```
4. Gere os relatórios (antigo botão **Imprimir PDF - Exportar Excel**):
   ```bash
   python geradora_cvm.py gerar
   ```
   Os arquivos saem em `pastas.saida` como `<FUNDO>_<CNPJ>_TRIMESTRAL_CVM.pdf/.xlsx`,
   junto com um `resumo_geracao.csv` que traz o PL e os avisos de cada fundo. Se algum fundo
   falhar, os demais são gerados, o erro vai para o resumo e o comando termina com código 1.

**Passo 0 (opcional):** para aproveitar uma Geradora que já tem as bases coladas, rode:
```bash
python geradora_cvm.py migrar --geradora "caminho\Geradora Limine - EMISSÃO EM LOTE.xlsb"
```
Esse comando extrai as abas Setup, CarteiraInforme, RESUMO_SINQIA e Pipefy para `entrada/` e
atualiza a `data_base`.

### Opções do `gerar`

| opção | efeito |
|---|---|
| `--cnpj 50.824.501/0001-60 ...` | gera só os fundos indicados |
| `--formato pdf` / `xlsx` | gera só um dos formatos |
| `--saida <pasta>` | grava em outra pasta |
| `--data-base 2026-12-31` | troca a data base sem mexer no config |
| `--com-lastro` | exibe as seções II e III, que ficam ocultas na planilha |
| `--validados <pasta>` | preenche as linhas "analisados" (32 a 34) com o `.xlsx` cujo nome contém o CNPJ |

Outro comando: `python geradora_cvm.py listar-validados` (antigo **Listar Verificação de Lastro**).

## Como a planilha calculava (e o script reproduz)

| Relatório | Regra |
|---|---|
| Fundo / CNPJ | Nome vem da lista de fundos (primeiro registro com o CNPJ) |
| Período | `MONTH(data)/3 & "º Trimestre/" & YEAR(data)` |
| I. Composição | Soma de `Valor` da Carteira Informe por categoria, filtrando CNPJ e data base. A categoria é a **primeira** chave de `mapeamento_carteira` (na ordem) contida na descrição em minúsculas. % sobre PL = valor / linha "Patrimônio Líquido" |
| II. Lastro (oculta) | Resumo Sinqia filtrado por `DOC_FUNDO` = 13 primeiros dígitos do CNPJ: a vencer, vencidos, liquidação (substituídos) e aquisição |
| IV e V. Textos | Primeira linha do Pipefy com o CNPJ do fundo; um campo vazio (ou fundo sem resposta) recebe `texto_gestor_sem_resposta` |
| Quem é gerado | Todo fundo da lista que tem ao menos uma linha na Carteira Informe (coluna "Tem Informe?") |

### Fundos novos e encerrados

A lista de fundos (`fundos.xlsx`, antiga aba Setup) é mantida à mão.

- **Fundo encerrado:** pode continuar na lista. Sem carteira no trimestre, ele não é gerado e aparece no
  aviso `[INFO] N fundo(s) sem carteira na base, não gerados`.
- **Fundo novo:** inclua uma linha com nome e CNPJ em `fundos.xlsx`. Se ele aparecer na Carteira Informe
  e não estiver na lista, o `gerar` **não** gera o relatório, mas avisa com `[AVISO] … não estão na lista
  de fundos` (com o CNPJ e o título do card do Pipefy, quando houver) e registra o fundo no
  `resumo_geracao.csv` como "NÃO GERADO".

### Formatação: `assets/modelo_cvm.xlsx`

O modelo é a própria aba CVM **salva pelo Excel** (só valores, sem vínculos externos): fontes,
bordas, alturas de linha, larguras de coluna, configuração de página e o logo do cabeçalho de
impressão.

- O **XLSX** de cada fundo é esse arquivo com as células calculadas preenchidas.
  Abrindo no Excel e imprimindo, ele sai igual ao PDF da planilha original.
- O **PDF** é desenhado a partir da mesma grade, imitando a impressão do Excel (A4, zoom 68%):
  usa a Arial do Windows e as medidas calibradas no PDF que o próprio Excel exporta.

Para mudar um texto fixo, a formatação ou o logo, edite o modelo no Excel e salve. Não é
preciso mexer no código.

## Testes

```bash
python -m unittest discover -s tests -v
```

Usam bases fictícias criadas na hora (as bases reais são confidenciais e não vão para o git).
Cobrem as regras de cálculo, os formatos de número, a leitura de CSV, o XLSX e o PDF, e o
comando `gerar` (inclusive um fundo com erro no meio do lote).

## Validação contra a planilha (3T2026)

A Geradora original foi recalculada no Excel para os 195 fundos do Setup e o macro
`imprimir_exportar` foi reproduzido nos 179 que têm carteira (178 do macro + o 1º fundo, que ele pulava):

- **Cálculo:** as 14.235 células calculadas batem com a aba CVM, exceto os casos do "0" descritos abaixo.
- **XLSX:** valores das 640.088 células (A1:AE116) iguais à aba recalculada. Formatação, alturas,
  larguras, página e logo idênticos ao XLSX salvo pelo Excel. Impresso pelo Excel, o XLSX gerado
  dá um PDF com todas as palavras na mesma posição do PDF original em 164 fundos; nos outros 14
  a única diferença é o caso do "0".
- **PDF:** mesmo número de páginas nos 178 fundos. Das 12.460 linhas de texto dos PDFs do Excel,
  todas aparecem idênticas, com as mesmas quebras de linha, exceto os 126 "0" e 9 nomes longos que o Excel cortava.
  Nos 164 fundos sem o caso do "0", todas as linhas ficam a menos de 1 pt (0,35 mm) da posição no PDF do Excel.

## Segurança

O token da API do Pipefy estava gravado no macro `pipefy` da planilha, legível por qualquer pessoa com
acesso ao arquivo. O script lê o token da variável de ambiente `PIPEFY_TOKEN`. **Gere um token novo e
revogue o antigo.**

## Diferenças em relação à planilha

- **Card vazio no Pipefy não imprime mais "0".** Na planilha, o `FILTER` devolvia 0 para campo vazio,
  e o relatório saía com "0" no lugar do texto. Isso afetava 14 dos fundos gerados no 3T2026.
  O script usa o texto padrão de "gestor não se manifestou".
- **O XLSX leva os textos do gestor.** O macro convertia as fórmulas em valores célula a célula,
  começando pelo `FILTER` do Pipefy (I5). Com isso, os textos do gestor do XLSX saíam errados
  (vazios nos arquivos do 2T2026). O PDF da planilha saía certo, porque era exportado antes.
- **Primeiro fundo da lista agora é gerado.** O macro começava na linha 7 do Setup e pulava a linha 6.
- **CNPJ comparado só pelos dígitos** (Carteira, Pipefy, lista): pontuação ou espaço diferente não faz mais o fundo "sumir".
- **CNPJ repetido na lista é gerado uma vez só.**
- **Textos longos do gestor não são cortados no PDF.** No Excel, a célula mesclada cortava o
  texto. O script aumenta a linha até caber.
- **Nome do fundo longo aparece inteiro.** O Excel cortava o começo dos nomes que não cabiam no
  cabeçalho (por exemplo "OMERCANTIL…" em vez de "AGROMERCANTIL…"). O PDF reduz a fonte até caber.
- **Sem "-0,00%".** Percentuais muito pequenos e negativos saem como "0,00%", como no Excel.
- **O XLSX não tem vínculo externo** com a planilha "Gerador Informe Mensal Sinqia" (o Excel
  não pede mais para atualizar vínculos ao abrir) nem a aba vazia "Planilha1".
- `popularGeradora` não foi portado. Ele salvava uma cópia `.xlsb` inteira por fundo, mas no layout atual
  do Setup lia o nome do fundo como se fosse o CNPJ e buscava carteiras no caminho do 2T2025. Os `.xlsb`
  que estão em `_CVM` saíram desse macro e têm todos os valores zerados.
- Mantidos de propósito, por fidelidade (vale revisar):
  - `F33` divide por `D27`, quando o esperado seria `D28`.
  - "Posicao em Derivativos - Passivo" cai em Patrimônio Líquido (chave `derivativos`).
  - "Operações Compromissadas" é mapeada, mas não tem linha no quadro e fica de fora (o script avisa).
  - "Valores a Receber" não tem chave de mapeamento, então sai sempre zerado.
