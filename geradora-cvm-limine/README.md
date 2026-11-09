# Geradora CVM Limine (Python)

Substitui a planilha `Geradora Limine - EMISSÃO EM LOTE.xlsb`. Gera o **Demonstrativo
Trimestral** (PDF + XLSX) de cada fundo da Limine sem abrir o Excel.

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
   - Também aceita `.csv` (separador `;` ou `,`) ou `.xlsb`: basta trocar o nome em `config.json` → `arquivos`.
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
   junto com um `resumo_geracao.csv` que traz o PL e os avisos de cada fundo.

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

A formatação do XLSX vem de `assets/modelo_cvm.xlsx`, que é a própria aba CVM sem as fórmulas.
Para mudar um texto fixo do relatório, edite esse modelo; o PDF lê os mesmos textos dele.

## Validação

Na base do 3T2026, o script foi comparado com a planilha recalculada no Excel para os 195 fundos do
Setup: 13.650 células calculadas e as 114.840 células dos 179 XLSX gerados. Tudo bate, exceto os
casos do "0" descritos abaixo.

## Segurança

O token da API do Pipefy estava gravado no macro `pipefy` da planilha, legível por qualquer pessoa com
acesso ao arquivo. O script lê o token da variável de ambiente `PIPEFY_TOKEN`. **Gere um token novo e
revogue o antigo.**

## Diferenças em relação à planilha

- **Card vazio no Pipefy não imprime mais "0".** Na planilha, o `FILTER` devolvia 0 para campo vazio,
  e o relatório saía com "0" no lugar do texto. Isso afetava 14 fundos do 3T2026.
  O script usa o texto padrão de "gestor não se manifestou".
- **Primeiro fundo da lista agora é gerado.** O macro começava na linha 7 do Setup e pulava a linha 6.
- **CNPJ comparado só pelos dígitos** (Carteira, Pipefy, lista): pontuação ou espaço diferente não faz mais o fundo "sumir".
- **Textos longos do gestor quebram linha no PDF.** No Excel, a célula mesclada cortava o texto.
- O XLSX não leva o logo no cabeçalho de impressão (o openpyxl não grava imagem de cabeçalho). O PDF leva.
- `popularGeradora` não foi portado. Ele salvava uma cópia `.xlsb` inteira por fundo, mas no layout atual
  do Setup lia o nome do fundo como se fosse o CNPJ e buscava carteiras no caminho do 2T2025. Os `.xlsb`
  que estão em `_CVM` saíram desse macro e têm todos os valores zerados.
- Mantidos de propósito, por fidelidade (vale revisar):
  - `F33` divide por `D27`, quando o esperado seria `D28`.
  - "Posicao em Derivativos - Passivo" cai em Patrimônio Líquido (chave `derivativos`).
  - "Operações Compromissadas" é mapeada, mas não tem linha no quadro e fica de fora (o script avisa).
  - "Valores a Receber" não tem chave de mapeamento, então sai sempre zerado.
