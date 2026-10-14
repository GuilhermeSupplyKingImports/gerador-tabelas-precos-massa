# GERADOR DE TABELAS DE PREÇOS EM MASSA

Este app gera arquivos CSV no modelo de tabelas de preços em massa do Protheus e mantém relatórios analíticos em Excel.

A versão atual possui uma **Home** com dois módulos:

1. **Precificação Comum**
2. **Adequação ML**

## Como rodar localmente

No Windows, extraia o ZIP e clique em:

```txt
executar_app.bat
```

Ou rode pelo terminal:

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Seleção de arquivos antes da geração

Nos dois módulos, antes de processar, o usuário pode marcar exatamente quais saídas deseja gerar.

Exemplos:

- somente relatório completo;
- somente tabelas CSV para Protheus;
- relatório + log;
- pacote completo.

O ZIP final sempre inclui apenas os arquivos selecionados naquela execução.

## Seleção de SKUs antes da exportação

Depois de gerar o processamento, o app exibe a etapa **Revisar SKUs para exportação**.

Nessa etapa:

- todos os SKUs vêm marcados por padrão;
- o usuário pode desmarcar os SKUs que não devem ir para os arquivos Protheus;
- os CSVs e o ZIP são montados somente com os SKUs marcados;
- a análise continua visível para auditoria;
- o relatório Excel inclui a coluna `Exportar?`, indicando o que entrou ou não na exportação daquela execução.

Essa função evita edição manual do CSV no Excel e preserva o layout exigido pelo Protheus.

## Nome personalizado nos downloads

Depois da etapa de revisão de SKUs, o app exibe o campo **Identificação dos arquivos (opcional)**.

Quando preenchido, esse texto passa a compor o nome dos arquivos baixados, junto com o código da tabela e a data em `DDMMAAAA`.

Exemplo com identificação `FEBI` e data `17/06/2026`:

- `001 - FEBI 17062026.csv`
- `007 - FEBI 17062026.csv`
- `013 - FEBI 17062026.csv`
- `relatorio_precificacao_comum - FEBI 17062026.xlsx`
- `log_precificacao_comum - FEBI 17062026.txt`
- `pacote_precificacao_comum - FEBI 17062026.zip`

Se o campo ficar vazio, o app usa o padrão técnico anterior.

## Módulo 1 — Precificação Comum

Use este módulo quando os analistas precisarem gerar as tabelas normais de precificação.

### Tabelas-fonte existentes

Antes do upload, o usuário informa quais preços já possui na planilha:

- `001`;
- `007`;
- ou `001 + 007`.

O preço informado é sempre a autoridade. Em especial, quando a planilha já contém a 007, o app **não recalcula nem substitui essa 007** usando o desconto da 001.

Depois, o usuário escolhe separadamente quais tabelas deseja gerar. Exemplo: se já possui `001 + 007`, pode gerar somente `004`, `012` e `013`.

### Planilha de entrada

O layout depende das fontes selecionadas:

| Fontes selecionadas | Estrutura |
|---|---|
| 001 | A=SKU, B=Custo Médio, C=Preço 001 |
| 007 | A=SKU, B=Custo Médio, C=Preço 007 |
| 001 + 007 | A=SKU, B=Custo Médio, C=Preço 001, D=Preço 007 |

Os SKUs são normalizados para cinco dígitos quando aplicável.

### Regras aplicadas

| Tabela | Regra |
|---|---|
| 001 | Preço 001 informado pelo usuário; não é reconstruído a partir da 007 |
| 004 | Menor valor entre `Custo médio / 2` e `Preço final 007 / 4` |
| 012 | Menor valor entre `Custo médio / 2` e `Preço final 007 / 4` |
| 007 | Se informada, preserva o valor informado; caso contrário, usa Preço 001 com desconto configurado |
| 013 | Preço final da 007 + acréscimo configurado |

O relatório mostra também a origem do preço 007 e, quando 001 e 007 foram informadas, a diferença entre a 007 existente e a 007 teórica que seria calculada pela 001.

## Módulo 2 — Adequação ML

Use este módulo para a rotina de adequação com os preços revisados do Mercado Livre.

### Entradas

- Tabela 007 oficial
- `PREÇOS ML - REVISADO.xlsx`, com:
  - A = SKU
  - B = Preço ML
  - C = Custo Médio
- Relatório anterior opcional para comparativo

### Regras aplicadas

- Se o preço comercial 007 estiver maior que o preço ML, o sistema calcula `Preço ML x 0,90`.
- Se o preço comercial 007 já estiver menor ou igual ao ML, mantém o preço 007.
- Se o preço comercial 007 estiver como `0,01`, o sistema trata como item sem preço de venda cadastrado e calcula `Preço ML x 0,90`.
- As tabelas **001, 007 e 013 são iguais no módulo ML**; muda apenas o código da tabela na primeira linha do CSV.
- As tabelas **004 e 012** usam o menor valor entre `Custo Médio / 2` e `Preço final 007 / 4`.
- O lucro bruto é calculado pela regra `Preço final da tabela 007 - Custo Médio`.
- O sistema gera alertas de prejuízo, markup baixo, margem baixa e custos inválidos.

### Arquivos gerados no módulo ML

- `001 - IDENTIFICACAO DDMMAAAA.csv`
- `004 - IDENTIFICACAO DDMMAAAA.csv`
- `007 - IDENTIFICACAO DDMMAAAA.csv`
- `012 - IDENTIFICACAO DDMMAAAA.csv`
- `013 - IDENTIFICACAO DDMMAAAA.csv`
- `relatorio_tabela_001_004_007_012_013_ml - IDENTIFICACAO DDMMAAAA.xlsx`
- `log_tabela_001_004_007_012_013_ml - IDENTIFICACAO DDMMAAAA.txt`
- `pacote_tabela_001_004_007_012_013_ml - IDENTIFICACAO DDMMAAAA.zip`

## Formato dos arquivos Protheus

Todos os arquivos Protheus são gerados em CSV, sem cabeçalho, com separador `;` e seguem o mesmo padrão:

| Célula/Coluna | Regra |
|---|---|
| Linha 1 | Código da tabela com 3 dígitos e 4 campos no total, exemplo `007;;;` |
| Coluna A, a partir da linha 2 | SKUs normalizados com 5 dígitos, exemplo `00050` |
| Coluna B, a partir de B2 | Preço com ponto decimal, exemplo `25.99` |
| Coluna C, a partir de C2 | Data da alteração em `DD/MM/AAAA` |
| Delimitador final | Todas as linhas terminam com `;`, exemplo `00050;25.99;17/06/2026;` |
| Observação | O relatório analítico permanece em `.xlsx`; apenas os arquivos de subida Protheus são `.csv` |

## Logs

Cada módulo gera um `.txt` com intercorrências, incluindo:

- SKU vazio;
- SKU com mais de 5 dígitos;
- SKU em formato suspeito;
- SKU duplicado;
- preço inválido;
- custo inválido;
- linhas que não foram enviadas para alguma tabela por falta de dados válidos.

## Execução por terminal

### Adequação ML

```bash
python run_cli.py ml --tabela-007 "tabela 007-0526.xlsx" --ml "PREÇOS ML - REVISADO.xlsx" --saida saida_ml
```

### Precificação Comum

```bash
python run_cli.py comum --base "precificacao_comum.xlsx" --saida saida_comum
```

## Atualização V16

A Precificação Comum passou a separar explicitamente **fontes já existentes** de **tabelas a gerar**.

Exemplo operacional:

1. selecionar `001` e `007` como tabelas-fonte existentes;
2. enviar uma planilha com `SKU`, `Custo Médio`, `Preço 001` e `Preço 007`;
3. selecionar somente `004`, `012` e `013` para geração;
4. o app preserva a 007 existente e calcula as demais a partir dela.

Essa regra elimina a obrigatoriedade de sempre reconstruir todo o conjunto a partir da tabela 001.

## Atualização V15

As tabelas internas **004** e **012** acompanham promoções da tabela 007.

Para cada SKU, o app calcula dois candidatos:

1. cálculo original da venda interna: `Custo Médio / 2`;
2. novo cálculo promocional: `Preço final da tabela 007 / 4`.

O menor valor válido entre os dois é usado como preço final das tabelas **004** e **012**.
