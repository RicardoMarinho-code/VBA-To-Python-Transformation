# Automação de Empenhos (VBA → Python)

Conversão das macros VBA de Excel (`ProcessarEmpenhos` e `ResetarEmpenhos`) para
Python, usando [openpyxl](https://openpyxl.readthedocs.io/).

O objetivo é, a partir das planilhas de **saldos de empenho**, preencher
automaticamente a aba de trabalho (**CORH**) com a Nota de Empenho, a fonte, a
competência e os valores utilizado/remanescente de cada empenho — distribuindo o
valor de cada linha entre os empenhos disponíveis do mesmo contrato/município.

---

## Índice

- [Requisitos](#requisitos)
- [Instalação](#instalação)
- [Configuração (`.env`)](#configuração-env)
- [Como usar](#como-usar)
- [Como funciona](#como-funciona)
- [Mapeamento de colunas](#mapeamento-de-colunas)
- [Observações importantes](#observações-importantes)
- [Estrutura do projeto](#estrutura-do-projeto)

---

## Requisitos

- Python 3.8 ou superior
- [openpyxl](https://pypi.org/project/openpyxl/) (ver `requirements.txt`)

## Instalação

```bash
pip install -r requirements.txt
```

## Configuração (`.env`)

Os **nomes dos arquivos** não ficam no código — ficam em um arquivo `.env` na
pasta do projeto. O `.env` é lido pelo `config.py` (sem depender de nenhuma
biblioteca externa).

Crie o seu `.env` a partir do exemplo:

```bash
cp .env.example .env
```

E edite os valores:

```ini
# Caminho (ou nome) do arquivo Excel a ser processado.
# Se for só o nome, ele deve estar na mesma pasta destes scripts.
WORKBOOK_PATH=Planilha de Financeiro.xlsm

# Nome do arquivo de relatório gerado pelo depurador.
REPORT_FILE=relatorio_depuracao.csv
```

Detalhes:

- Se você informar só o **nome** do arquivo, ele é resolvido para um caminho
  **absoluto** relativo à pasta dos scripts — então funciona rodando de qualquer
  diretório.
- Variáveis de ambiente do sistema têm prioridade sobre o `.env` (dá para
  sobrescrever sem editar o arquivo).
- O `.env` é **ignorado pelo git** (`.gitignore`); o `.env.example` é versionado
  para documentar o que precisa ser configurado.

## Como usar

O fluxo recomendado é: **depurar → processar → (se precisar) resetar**.

### 1. Depurar (simulação, não altera nada)

```bash
python depurar_empenhos.py
```

Faz um *dry-run*: lê a planilha somente para leitura, **prevê** o que a macro faria
e aponta os problemas (contrato/município não encontrado, saldo insuficiente,
linhas já processadas). Também gera um relatório CSV (`REPORT_FILE`) com os
problemas encontrados. **Não modifica a planilha.**

### 2. Processar

```bash
python processar_empenhos.py
```

Preenche a aba CORH, insere linhas quando um valor precisa de mais de um empenho,
aplica a formatação e cria a aba `LOG` com as pendências. **Grava por cima do
arquivo** — feche o Excel antes de rodar.

### 3. Resetar (desfazer)

```bash
python resetar_empenhos.py
```

Desfaz o processamento: exclui as linhas geradas pela macro, limpa os campos
preenchidos e remove as colunas criadas (O, P, Q).

> Você também pode importar as funções em outro script:
> ```python
> from processar_empenhos import process_commitments
> from resetar_empenhos import reset_commitments
> process_commitments()   # usa o WORKBOOK_PATH do .env
> ```

## Como funciona

**Abas envolvidas** (nomes definidos no código):

- `SALDOS DE EMPENHO` — origem dos empenhos (dados a partir da linha 7).
- `ajustada CORH` — aba de trabalho a ser preenchida (cabeçalho na linha 5).
- `LOG` — criada automaticamente com as pendências.

**Lógica do processamento:**

1. Carrega os empenhos da aba de saldos, considerando **apenas os anos 2025 e
   2026** (o ano é extraído da Nota de Empenho). O saldo usado é o da coluna Q se
   for maior que zero; senão, o da coluna O.
2. Monta uma chave `contrato + município` (o município é normalizado: maiúsculas,
   sem acentos, sem hífen).
3. Percorre a aba CORH **de baixo para cima**. Para cada linha:
   - pula se já foi processada (coluna H preenchida) ou se o valor é ≤ 0;
   - se a chave não existe nos saldos → pinta de **amarelo** e registra
     "Contrato/Município não encontrado";
   - senão, **distribui o valor** entre os empenhos daquela chave, do ano mais
     antigo para o mais novo. O primeiro empenho preenche a própria linha; cada
     empenho adicional gera uma **nova linha** logo abaixo (marcada na coluna Q);
   - a **competência** é `3` para empenhos de 2026 e `2` para os anteriores;
   - se sobrar valor sem saldo → pinta de **laranja** e registra "Saldo
     insuficiente".
4. O saldo consumido é **compartilhado entre as linhas** (um mesmo empenho usado
   em várias linhas vai reduzindo até acabar).
5. No fim, aplica bordas, negrito no cabeçalho, formato numérico e ajuste de
   largura das colunas.

## Mapeamento de colunas

**Aba `SALDOS DE EMPENHO` (leitura):**

| Coluna | Conteúdo |
|:---:|---|
| A | Contrato |
| G | Fonte |
| I | Nota de Empenho |
| M | Município |
| O | Saldo (alternativa) |
| Q | Saldo (prioritário, se > 0) |

**Aba `ajustada CORH` (leitura e escrita):**

| Coluna | Conteúdo | Escrita? |
|:---:|---|:---:|
| A | Contrato | leitura |
| C | Município | leitura |
| H | Nota de Empenho | escrita |
| J | Fonte | escrita |
| L | Competência | escrita |
| M | Valor total | leitura |
| O | Saldo utilizado NE | escrita |
| P | Saldo remanescente NE | escrita |
| Q | Controle Macro (marcador) | escrita |

## Observações importantes

- **Faça backup** da planilha antes de processar — o script grava por cima do
  arquivo.
- **Feche o Excel** antes de rodar o `processar`/`resetar`, senão a gravação falha.
- **Arquivos `.xlsm`**: o carregamento usa `keep_vba=True` automaticamente para
  **preservar as macros VBA** ao salvar. Sem isso, o openpyxl removeria as macros.
- **Fórmulas**: o depurador lê o valor calculado das células (`data_only=True`),
  o que exige que a planilha tenha sido salva pelo Excel ao menos uma vez. Se a
  região de dados usar fórmulas, avalie o impacto antes de processar.
- O depurador **nunca** altera a planilha; use-o à vontade para inspecionar.

## Estrutura do projeto

```
VBA-Python-Transformation/
├── processar_empenhos.py     # processamento (equivale à macro ProcessarEmpenhos)
├── resetar_empenhos.py       # desfaz o processamento (macro ResetarEmpenhos)
├── depurar_empenhos.py       # dry-run: prevê o resultado e lista problemas
├── config.py                 # carrega o .env e resolve os caminhos
├── .env                      # nomes dos arquivos (NÃO versionado)
├── .env.example              # modelo do .env (versionado)
├── requirements.txt          # dependências (openpyxl)
├── ProcessarEmprenhos.txt    # macro VBA original (referência)
└── ResetarEmpenhos.txt       # macro VBA original (referência)
```
