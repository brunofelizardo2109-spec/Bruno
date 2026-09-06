---
name: organizador-financeiro
description: "Use sempre que o usuário (Bruno) mandar um gasto avulso, um gasto fixo, uma fonte de renda, uma alteração de valor/status de algo já cadastrado, ou pedir para fechar/gerar a planilha do mês. Este é o organizador financeiro pessoal do Bruno: cada gasto que ele manda ao longo do mês é registrado em financeiro/lancamentos.csv, gastos fixos em financeiro/gastos_fixos.csv, renda em financeiro/renda.csv — commitados e enviados ao GitHub imediatamente, nunca só localmente. No fim do mês (ou quando ele pedir), gerar a planilha .xlsx com scripts/gerar_planilha_mensal.py e entregar via SendUserFile. Nunca montar a planilha manualmente ou editar os CSVs por fora do fluxo abaixo."
---

# Organizador financeiro pessoal

## Por que este fluxo existe
O ambiente de execução é efêmero (container reclamado por inatividade) — se um
lançamento só existir na conversa ou num arquivo não commitado, ele some. Os
3 CSVs em `financeiro/` são a única fonte de verdade; a planilha `.xlsx` é
sempre **gerada** a partir deles por `scripts/gerar_planilha_mensal.py`, nunca
editada à mão. Ver `financeiro/README.md` para o schema completo de cada CSV.

## Categorias fixas (nunca inventar uma nova sem avisar)
Moradia, Contas e Utilidades, Alimentação, Transporte, Saúde, Educação, Lazer
e Assinaturas, Vestuário, Investimentos, Dívidas e Financiamentos, Outros.

Palavras-chave para inferir a categoria automaticamente (não perguntar se bater
com confiança):
- **Moradia**: aluguel, condomínio, IPTU, reforma da casa
- **Contas e Utilidades**: luz, água, gás, internet, telefone, celular
- **Alimentação**: mercado, supermercado, ifood, restaurante, padaria, feira
- **Transporte**: uber, 99, combustível, gasolina, estacionamento, manutenção do carro, pedágio
- **Saúde**: plano de saúde, farmácia, consulta, exame, academia, dentista
- **Educação**: curso, faculdade, livro, material escolar
- **Lazer e Assinaturas**: netflix, spotify, cinema, bar, viagem, streaming
- **Vestuário**: roupa, tênis, calçado
- **Investimentos**: aplicação, CDB, tesouro direto, ações, previdência
- **Dívidas e Financiamentos**: fatura do cartão, empréstimo, financiamento, parcela

Se a descrição não bater com nenhuma palavra-chave com confiança, perguntar com
`AskUserQuestion` oferecendo a lista fixa de categorias — nunca chutar uma
categoria só para não perguntar.

## Registrar um gasto variável (`financeiro/lancamentos.csv`)
Bruno manda algo como "gastei 450 no mercado hoje" ou "paguei 120 de uber ontem".
1. Extrair: `valor` (obrigatório), `descricao` (obrigatório), `data` (se não
   disser, usar a data de hoje), `categoria` (inferir; perguntar só se
   ambíguo), `forma_pagamento` (se mencionado) e `observacao` (se sobrar
   contexto relevante).
2. Nunca inventar valor ou data — se o valor não vier explícito, perguntar.
3. Adicionar uma linha ao final de `financeiro/lancamentos.csv` (formato
   `AAAA-MM-DD,categoria,descricao,valor,forma_pagamento,observacao`, `valor`
   sempre com ponto decimal, nunca vírgula).
4. Commitar e dar **push imediato** (ver seção "Commit e push" abaixo) — não
   acumular vários lançamentos num commit só esperando o usuário mandar mais.
5. Confirmar em uma frase: `Registrado: R$ 450,00 em Alimentação (mercado) — 06/09.`

## Cadastrar/alterar um gasto fixo (`financeiro/gastos_fixos.csv`)
- **Novo gasto fixo**: precisa de descrição, categoria (inferir/perguntar),
  valor e dia de vencimento — se faltar o dia, perguntar (é usado na projeção
  do mês). `ativo` começa `sim`.
- **Alteração de valor** ("meu aluguel subiu pra 3200"): localizar a linha por
  `descricao` (case-insensitive, aceitar correspondência parcial óbvia); se
  houver mais de uma linha parecida, perguntar qual. Atualizar só o `valor`
  dessa linha — não duplicar.
- **Cancelamento** ("cancelei a academia"): setar `ativo=nao` na linha
  correspondente. Nunca apagar a linha — isso destrói o histórico do que já
  foi pago antes do cancelamento.
- Commitar e dar push imediato. Confirmar em uma frase.

## Cadastrar/alterar renda (`financeiro/renda.csv`)
Mesmo padrão do gasto fixo: descrição, valor, dia de recebimento, tipo (`fixo`
para salário, `variavel` para freelance/bônus). Alteração de valor atualiza a
linha existente; renda que parou de existir vira `ativo=nao`, nunca é apagada.
Commitar e dar push imediato. Confirmar em uma frase.

## Commit e push (depois de CADA registro/alteração, sem exceção)
```
git add financeiro/<arquivo>.csv
git commit -m "financeiro: <ação curta em português>"
git push -u origin <branch atual>
```
Mensagem de commit específica (ex.: `financeiro: registra gasto - mercado (R$ 450,00)`,
`financeiro: atualiza valor do aluguel para R$ 3200`), nunca genérica tipo
"update". Se o push falhar por rede, seguir a política de retry padrão (2s,
4s, 8s, 16s) antes de avisar o usuário — sem isso o lançamento existe só no
container e se perde quando ele for reciclado.

## Gerar a planilha do mês ("fecha o mês", "manda a planilha", fim do mês)
1. Se Bruno não especificar o mês, usar o mês corrente. Ele pode pedir outro
   mês/ano explicitamente.
2. Rodar:
   ```
   python3 scripts/gerar_planilha_mensal.py --mes N --ano AAAA --out /tmp/Organizador_Financeiro_<Mes>-<AAAA>.xlsx
   ```
   O script já valida categorias, calcula os totais em Python, escreve
   fórmulas reais (SUMIFS/SUMIF/IFERROR) nas células e confere sozinho que o
   valor injetado bate com a fórmula antes de terminar — se ele falhar
   (`ValueError`/`RuntimeError`), **não** tentar contornar montando a planilha
   na mão: o erro está apontando um dado inconsistente nos CSVs (categoria
   inválida, data mal formatada, valor não numérico) — corrigir a causa raiz.
3. Entregar o `.xlsx` via `SendUserFile` (nunca commitar o arquivo gerado —
   ver `.gitignore`/nota abaixo).
4. Resumir em 1-2 frases usando os números que o próprio script imprime
   (renda ativa, gastos fixos, gastos variáveis, saldo) — não recalcular na
   mão. Se o saldo do mês for negativo, dizer isso explicitamente logo no
   resumo, não deixar só na planilha.

## Por que o .xlsx nunca é commitado
Estado (fonte de verdade) e relatório (derivado) são coisas diferentes: commitar
o `.xlsx` geraria duas fontes de verdade divergentes assim que um novo
lançamento entrasse depois de gerado. O `.xlsx` é sempre reprodutível a partir
dos CSVs — se sumir, é só rodar o script de novo.

## Nota técnica (não pular a etapa 3 do script)
O LibreOffice headless deste ambiente não recalcula fórmulas de forma
confiável (timeout mesmo com 90s de espera, testado e confirmado). Por isso o
script injeta o valor calculado direto no XML da fórmula e confere o resultado
antes de terminar — mesma técnica documentada em
`.claude/skills/contratacao-fornecedor/SKILL.md`. Se algum dia for adicionar
uma fórmula nova ao script, ela precisa entrar também no dicionário de cache
em `montar_resumo()`, senão a conferência final vai falhar (o que é o
comportamento certo: melhor falhar alto do que entregar uma célula em branco).
