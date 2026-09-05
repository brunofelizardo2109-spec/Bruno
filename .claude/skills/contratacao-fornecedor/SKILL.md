---
name: contratacao-fornecedor
description: "Use sempre que o usuário (Bruno, engenheiro civil, Kasap Engenharia, obras de alto padrão) enviar 2 ou 3 orçamentos/propostas de fornecedores em PDF para um mesmo serviço/escopo e quiser fechar a contratação completa: mapa de cotação, checagem de divergência entre os orçamentos, escolha do fornecedor vencedor, ficha de dados do fornecedor, ordem de compra (OC) e contrato prontos. Dispara também quando ele pedir só uma etapa isolada (ex: \"gera a OC\", \"faz o contrato desse fornecedor\") — nesse caso execute só a etapa pedida, aproveitando os dados já anexados/informados. Sempre usar este fluxo e os templates de templates/ em vez de montar mapa/OC/contrato do zero — são os documentos oficiais da Kasap, com cláusulas e fórmulas já validadas."
---

# Pipeline de contratação de fornecedor (obra de alto padrão)

## Por que este fluxo existe
Bruno já tem 3 documentos padronizados da empresa (`templates/mapa_cotacao_modelo.xlsx`,
`templates/oc_modelo.xlsx`, `templates/contrato_modelo.docx`) com cláusulas jurídicas,
fórmulas e formatação já revisadas. O risco real não é "não saber montar uma planilha
de orçamento" — é (a) comparar orçamentos de fornecedores que na verdade não estão
cotando a mesma coisa (escopo/quantidade/unidade diferentes) e só perceber isso depois
de já ter fechado, e (b) um dado divergir entre mapa → OC → contrato (nome, CNPJ, valor,
prazo) sem ninguém notar. Este skill existe para fechar essas duas lacunas com dois
gates explícitos — checagem de divergência na entrada, conferência cruzada na saída —
mantendo o resto do trabalho manual mínimo.

## Templates (nunca redesenhar do zero)
- `templates/mapa_cotacao_modelo.xlsx` — Mapa de Cotações (3 fornecedores lado a lado).
- `templates/oc_modelo.xlsx` — Ordem de Compra. Nota: na aba `OC001`, a coluna **E é
  QUANTIDADE** e **F é UNIDADE** nas linhas de item (32-36) — os rótulos do cabeçalho
  já foram corrigidos para bater com isso; a fórmula de total é sempre `=G{linha}*E{linha}`
  (valor unitário × quantidade). Não inverter.
- `templates/contrato_modelo.docx` — Contrato de prestação de serviços. CONTRATANTE
  (Kasap Engenharia) já vem fixo; CONTRATADA fica em branco para preencher.

Para cada obra/serviço novo, copie o template para um arquivo de trabalho — nunca edite
os arquivos em `templates/` diretamente.

## Fluxo (6 etapas + 2 gates de qualidade)

### Etapa 0 — Entrada
Bruno anexa 2 ou 3 PDFs de proposta/orçamento de fornecedores para o mesmo serviço.
Se ele não tiver informado, perguntar: nome da obra e nome/escopo do serviço. Perguntar
também se ele quer pular a checagem de divergência da Etapa 1 (nem toda contratação
precisa — cotações simples e já conhecidas podem pular). Se ele não disser nada, o
padrão é **checar**.

### Etapa 1 — Gate de divergência (pular apenas se Bruno pedir explicitamente)
Ler os PDFs inteiros (todas as páginas, com o Read tool — nunca confiar em resumo).
Comparar entre os fornecedores, item a item: mesma composição/escopo técnico, mesma
quantidade, mesma unidade, prazo de entrega, validade da proposta, condições de
pagamento embutidas na proposta.

- **Convergente** (nenhuma divergência relevante): avisar em 1 frase e seguir direto
  para a Etapa 2.
- **Divergente**: não seguir sozinho. Gerar um relatório curto em tabela — item |
  fornecedor A | fornecedor B | fornecedor C | tipo de divergência (item ausente /
  quantidade diferente / unidade diferente / escopo diferente / prazo diferente) — e
  usar AskUserQuestion com duas opções:
  1. **Prosseguir mesmo assim** → segue para a Etapa 2 usando os dados como estão,
     marcando no mapa quais itens ficaram divergentes.
  2. **Vou corrigir com os fornecedores** → entregar o relatório pronto para Bruno
     mandar aos fornecedores, e parar aqui. Quando ele reanexar os orçamentos
     corrigidos, repetir a Etapa 1 do zero nesse novo lote — é um loop: só sai dele
     quando os orçamentos convergirem ou quando Bruno escolher "prosseguir mesmo
     assim".

Isso é comparar escopo, não só preço: um orçamento mais barato cotando menos
quantidade ou material diferente não é "mais barato", é outra coisa — nunca tratar
como comparável sem avisar.

### Etapa 2 — Mapa de cotação
Montar `templates/mapa_cotacao_modelo.xlsx` preenchido, um fornecedor por coluna,
seguindo a tripla conferência (mapear item a item por descrição técnica completa,
nunca aproximar "no chute"; recalcular fórmulas com LibreOffice headless antes de
entregar; reportar qualquer suposição feita). Entregar o arquivo.

### Etapa 3 — Escolha do vencedor
Perguntar qual fornecedor ganhou, se Bruno ainda não tiver dito.

### Etapa 4 — Ficha de dados do fornecedor
Gerar uma tabela simples (pode ser no próprio chat ou uma planilha pequena) pedindo
só o que falta para OC e contrato: razão social, CNPJ, endereço completo, responsável
legal (nome/RG/CPF, para o contrato), responsável técnico/contato de obra, telefone,
e-mail, prazo de entrega/execução acordado, condições de pagamento acordadas (nº de
parcelas, valores, datas), banco/agência se houver faturamento direto. Esperar Bruno
preencher antes de seguir.

### Etapa 5 — Ordem de Compra
Bruno mantém **um único arquivo de OC por obra**, com uma aba por ordem de compra
(`OC001`, `OC002`, `OC003`, ...) — nunca um arquivo novo por fornecedor. O arquivo
segue o padrão de nome `KASAP Engenharia_Ordem de Compra - Obra <NOME DA OBRA>.xlsx`
e normalmente já existe na pasta da obra no Drive (ver Etapa 7). Antes de gerar:
1. Localizar esse arquivo (Drive ou anexado por Bruno). Se não existir ainda para a
   obra, criar a partir de `templates/oc_modelo.xlsx`.
2. Achar o próximo número de OC livre (maior `OCxxx` existente + 1).
3. Duplicar a aba `OC001` (ou a última aba) dentro do próprio arquivo, renomear para
   `OCxxx`, e preencher: CONTRATANTE = dados do cliente/obra, COMPRADOR = Kasap
   Engenharia (fixo), FORNECEDOR = dados da ficha da Etapa 4, itens = os itens do
   fornecedor vencedor no mapa (descrição técnica completa, não resumida — lembrar que
   nas linhas de item a coluna E é quantidade e F é unidade), parcelas de pagamento no
   bloco de observações (texto real acordado, nunca o placeholder).

### Etapa 6 — Contrato
Preencher `templates/contrato_modelo.docx`: CONTRATADA = dados da ficha, objeto/escopo
= descrição dos itens (a mesma da OC, não reescrever genérico), valor total e parcelas
= os mesmos da OC (inclusive valor por extenso), prazos = os mesmos da ficha/OC,
inserir a relação de itens como tabela real do Word (editável/pesquisável, não
imagem) no lugar de "[Inserir aqui a imagem do demonstrativo / planilha de orçamento]".

### Gate final — Conferência cruzada (antes de entregar OC e contrato)
Nunca entregar sem checar, e nunca ignorar silenciosamente uma inconsistência encontrada:
- Valor total da OC == soma das parcelas descritas nas observações da OC.
- Razão social, CNPJ e endereço do fornecedor idênticos, literalmente, entre ficha, OC
  e contrato.
- Prazo de entrega da OC == prazo de execução no contrato.
- Escopo do contrato reflete a descrição técnica real dos itens da OC, não um texto
  genérico.

Se algo não bater, corrigir antes de entregar e avisar o que foi ajustado — não
entregar "do jeito que deu".

### Etapa 7 — Entrega e arquivamento (Drive + Trello + download)
Bruno já organiza cada obra assim (validado no caso Concórdia Corporate — usar como
referência de convenção, não copiar valores):
- **Drive**: pasta da obra (ex. "41 - Reforma lobby - Concórdia Corporate") →
  subpasta de empreiteiros/fornecedores (ex. "05 - Empreiteiros e Fornecedores") →
  subpasta numerada por fornecedor (ex. "21 - Mega furos (furos de laje)"), com
  subpastas fixas dentro: `Proposta`, `Orçamento`, `Medição`, `Contrato`. O número da
  subpasta do fornecedor é o mesmo número do cartão dele no Trello — usar sempre para
  casar os dois.
- **Trello**: um board por obra (ex. "Controle De Comprar Concórdia"), com listas de
  pipeline (`Pedido de compra` → `Aguardando aprovação` → `Coleta aprovada` →
  `Aguardando entraga/ordem de serviço`) e uma lista separada `Contratos empreiteiro`
  com um cartão por fornecedor já contratado (nome `Contrato <Fornecedor>`).

Ao entregar um documento:
1. Salvar o arquivo na subpasta correta do Drive (mapa/relatório de divergência em
   `Orçamento`, contrato em `Contrato`; a OC vai como aba dentro do arquivo
   consolidado da obra, não como arquivo novo — ver Etapa 5).
2. Comentar no cartão Trello correspondente (achar pelo número/nome do fornecedor)
   com um link para o arquivo salvo no Drive. Esta integração Trello não tem ação de
   "anexar arquivo" de verdade — um comentário com o link do Drive é o equivalente
   funcional.
   - Se for o contrato e ainda não existir um cartão em `Contratos empreiteiro` para
     esse fornecedor, criar um novo cartão lá (`Contrato <Fornecedor>`) antes de
     comentar.
   - **Não mover o cartão entre listas de pipeline sozinho** (ex. de "Coleta
     aprovada" para "Aguardando entrega/ordem de serviço") — perguntar a Bruno antes,
     é o indicador visual do andamento da obra dele e mudar sem avisar bagunça a
     leitura do board.
3. Entregar o mesmo arquivo via SendUserFile para download imediato.

### Obra nova (ainda não tem pasta/board)
Hoje só existe a obra Concórdia — tudo vai para a pasta dela no Drive e para o board
"Controle De Comprar Concórdia" no Trello. Quando aparecer uma obra nova, **criar o
mecanismo inteiro replicando o padrão da Concórdia**, sem esperar Bruno montar na mão:
- **Drive**: pasta raiz da obra, subpasta de empreiteiros/fornecedores dentro dela, e
  para cada fornecedor uma subpasta numerada com `Proposta` / `Orçamento` / `Medição`
  / `Contrato` dentro. Copiar `templates/oc_modelo.xlsx` para dentro da pasta da obra
  como o arquivo consolidado de OC (`KASAP Engenharia_Ordem de Compra - Obra <NOME>.xlsx`).
- **Trello**: um board novo para a obra com as mesmas listas da Concórdia (`Pedido de
  compra`, `Aguardando aprovação`, `Coleta aprovada`, `Aguardando entraga/ordem de
  serviço`, `Contratos empreiteiro`, `Documentação de obra`, `Projetos`).

Só perguntar a Bruno o nome/número da obra nova (para nomear pasta e board de forma
consistente com a numeração dele) — não perguntar se deve criar a estrutura, isso já
está decidido.

## Estilo obrigatório
- Técnico, direto, resumido. Sem recapitular passo a passo no resumo final.
- Nunca inventar ou aproximar um dado que não veio do PDF, da ficha ou da OC. Campo
  sem informação = deixar em branco e avisar, nunca "chutar".
- Toda suposição feita (ex.: agrupar itens, corrigir um typo) deve ser reportada
  explicitamente no resumo final, nunca fica silenciosa.
- Entregar cada arquivo via SendUserFile com nome claro, ex:
  `OC-0032_NomeFornecedor_Servico.xlsx`, `Contrato_NomeFornecedor_Servico.docx`.

## Dados sensíveis — não commitar no git
Mapa, OC e contrato preenchidos contêm CNPJ, endereço, dados pessoais (RG/CPF de
representante) e valores comerciais de terceiros. **Nunca commitar esses arquivos
preenchidos neste repositório** — só os templates em branco em `templates/` fazem
parte do versionamento. O destino oficial dos documentos preenchidos é o Google
Drive + Trello da obra (Etapa 7), sempre também entregues via SendUserFile.
