# Organizador financeiro pessoal

Fonte de verdade: os 3 CSVs desta pasta. Cada gasto, renda ou gasto fixo que o
Bruno manda no chat é registrado aqui (commitado e enviado para o GitHub na
hora — o ambiente de execução é efêmero, então o que não for commitado se
perde). A planilha `.xlsx` do mês é sempre **gerada** a partir destes CSVs
por `scripts/gerar_planilha_mensal.py`, nunca editada à mão — os CSVs são a
única fonte de verdade.

## Arquivos

### `lancamentos.csv` — gastos variáveis, um por linha, histórico completo
Nunca é limpo ou resetado por mês; a planilha mensal filtra por `data` na hora
de gerar. Colunas:

| Coluna | Formato | Observação |
|---|---|---|
| `data` | `AAAA-MM-DD` | Data do gasto (não da mensagem) |
| `categoria` | uma das categorias fixas abaixo | |
| `descricao` | texto livre | |
| `valor` | número com ponto decimal (`123.45`) | sempre positivo |
| `forma_pagamento` | texto livre (ex.: `crédito`, `pix`, `débito`, `dinheiro`) | opcional |
| `observacao` | texto livre | opcional |

### `gastos_fixos.csv` — despesas recorrentes mensais (cadastradas uma vez)
| Coluna | Formato |
|---|---|
| `descricao` | texto livre |
| `categoria` | uma das categorias fixas abaixo |
| `valor` | número com ponto decimal |
| `dia_vencimento` | dia do mês (1-31) |
| `ativo` | `sim` / `nao` — gasto fixo inativo entra no histórico mas some da projeção do mês |

### `renda.csv` — fontes de renda
| Coluna | Formato |
|---|---|
| `descricao` | texto livre |
| `valor` | número com ponto decimal |
| `dia_recebimento` | dia do mês (1-31) |
| `tipo` | `fixo` ou `variavel` |
| `ativo` | `sim` / `nao` |

## Categorias fixas (não criar categoria nova sem avisar)
Moradia, Contas e Utilidades, Alimentação, Transporte, Saúde, Educação, Lazer
e Assinaturas, Vestuário, Investimentos, Dívidas e Financiamentos, Outros.

Taxonomia fixa por consistência: categoria livre por lançamento inviabiliza
comparar mês a mês. Se uma categoria nova for realmente necessária, ela entra
na lista acima (e neste README) antes de ser usada, não como exceção pontual.

## Gerar a planilha do mês
```
python3 scripts/gerar_planilha_mensal.py --mes 9 --ano 2026
```
Gera o `.xlsx` em `/tmp` (nunca dentro do repo — ver `.gitignore`) e a skill
`organizador-financeiro` entrega via SendUserFile.
