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
nunca aproximar "no chute"; reportar qualquer suposição feita). Antes de entregar,
aplicar o checklist de qualidade visual abaixo. Entregar o arquivo.

### Checklist de qualidade visual do .xlsx (mapa e OC, sempre antes de entregar)
Testado e validado no caso Prumus/item 15 — três bugs reais que passaram batido na
primeira entrega:
1. **Fórmula mostrando R$ 0,00 ou em branco**: o LibreOffice headless não funciona
   neste ambiente (`soffice --convert-to ...` falha até em arquivos triviais — não
   perder tempo tentando de novo, é limitação do sandbox). Sem recalcular, o
   `openpyxl` grava a fórmula mas nenhum valor em cache, e qualquer visualizador que
   não recalcule ao vivo (preview do Files/Quick Look no iPhone, por exemplo) mostra
   0 ou vazio em vez do resultado. Corrigir escrevendo o valor calculado direto no
   XML junto da fórmula (o próprio `openpyxl` não expõe isso na API — precisa abrir
   o `.xlsx` como zip, achar o `<c r="CELULA">...<f>...</f></c>` na
   `xl/worksheets/sheetN.xml` e inserir `<v>VALOR</v>` antes de `</c>`). Fazer isso
   para toda fórmula de subtotal/total que aparece no arquivo (mapa: `K` de cada
   item, `J22`/`L22`/`N22`, `J25`/`L25`/`N25`; OC: `H` de cada item, `H38`, `H50`).
   Depois, reabrir com `data_only=True` e conferir que os valores batem com a conta
   manual — nunca entregar confiando só na fórmula sem essa conferência.
2. **Linhas de item sem uso aparecendo em branco**: quando há menos itens do que
   linhas no template (ex.: 1 item só), **ocultar** as linhas extras
   (`ws.row_dimensions[r].hidden = True`) — nunca deletar (`delete_rows` quebra
   merged cells e desloca fórmulas sem avisar, já visto no caso Vale dos
   Cristais/KASAP) e nunca deixar visível em branco.
3. **Texto de descrição "estourando" pra fora da linha**: célula com `wrapText=True`
   mas linha com altura padrão (`row_dimensions[r].height = None`) só cabe uma
   linha de texto — uma descrição técnica longa transborda visualmente pra fora do
   desenho da célula em qualquer visualizador. Setar `row_dimensions[r].height`
   explicitamente pra um valor generoso (~15pt por linha de texto esperada; para
   descrição técnica típica de 3-5 linhas, 75-100pt) sempre que a descrição for mais
   longa que uma linha curta. Alguns templates (a OC, por exemplo) já vêm com altura
   generosa de fábrica — conferir antes de assumir que precisa mudar.

Sem uma ferramenta de renderização neste ambiente, a única forma de "ver como
ficou" é reabrir o arquivo salvo e inspecionar programaticamente: valores calculados
(`data_only=True`), `row_dimensions[r].hidden` e `.height`, `column_dimensions[c].hidden`.
Fazer essa inspeção sempre, não só quando o usuário reclamar.

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

### Etapa 7 — Entrega (download apenas, sem tocar em Drive/Trello)
Testado em produção: subir os `.xlsx`/`.docx` preenchidos direto pro Drive via
base64 não é confiável neste ambiente (arquivos de ~15-35 KB viram strings base64
de 15-45 mil caracteres, longas demais para reproduzir com garantia — uma tentativa
já corrompeu silenciosamente um arquivo real antes de ser pega pelo `fileSize` de
conferência). Não vale o risco com documentos de obra de verdade. Por decisão do
Bruno: **nunca** escrever direto no Drive nem no Trello — só gerar o arquivo e
entregar via SendUserFile. Bruno mesmo arrasta pro lugar certo.

Para o nome do arquivo entregue já ajudar Bruno a saber onde ele vai (sem que o
skill precise escrever lá), seguir a convenção que ele já usa:
- **Drive**: pasta da obra → subpasta de empreiteiros/fornecedores → subpasta
  numerada por fornecedor (mesmo número do cartão Trello dele), com subpastas fixas
  `Proposta` / `Orçamento` / `Medição` / `Contrato` dentro. Mapa e relatório de
  divergência vão em `Orçamento`; contrato em `Contrato`; a OC é aba dentro do
  arquivo consolidado da obra (ver Etapa 5), não arquivo novo.
- **Trello**: um board por obra, cartão por fornecedor nomeado
  `NN - Fornecedor (serviço)`, mais uma lista `Contratos empreiteiro` com um cartão
  `Contrato <Fornecedor>` por fornecedor já contratado.

Nomear o arquivo entregue incluindo o destino esperado, ex.:
`Orçamento_15-Prumus_Mapa-Cotacao.xlsx`, `Contrato_15-Prumus_MaoDeObraCivil.docx`,
e avisar em uma frase onde ele deveria ir (obra/fornecedor/subpasta).

## Estilo obrigatório
- Técnico, direto, resumido. Sem recapitular passo a passo no resumo final.
- Nunca inventar ou aproximar um dado que não veio do PDF, da ficha ou da OC. Campo
  sem informação = deixar em branco e avisar, nunca "chutar".
- Toda suposição feita (ex.: agrupar itens, corrigir um typo) deve ser reportada
  explicitamente no resumo final, nunca fica silenciosa.
- Entregar cada arquivo via SendUserFile com nome claro, ex:
  `OC-0032_NomeFornecedor_Servico.xlsx`, `Contrato_NomeFornecedor_Servico.docx`.

## Dados sensíveis — não commitar no git, não escrever em Drive/Trello
Mapa, OC e contrato preenchidos contêm CNPJ, endereço, dados pessoais (RG/CPF de
representante) e valores comerciais de terceiros. **Nunca commitar esses arquivos
preenchidos neste repositório** — só os templates em branco em `templates/` fazem
parte do versionamento. **Nunca escrever esses arquivos direto no Drive ou no
Trello** (ver Etapa 7) — entregar sempre via SendUserFile e deixar Bruno arquivar.
