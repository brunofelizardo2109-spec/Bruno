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
Se ele não tiver informado, perguntar: nome da obra e nome/escopo do serviço. Depois,
usar AskUserQuestion com a pergunta **"Deseja conferir se os orçamentos estão
equalizados?"** (sim/não) — nem toda contratação precisa (cotações simples e já
conhecidas podem pular direto pro mapa). Se ele não disser nada, o padrão é **sim,
checar**.
- Se **não**: pular a Etapa 1 e ir direto pra Etapa 2, sem perguntar de novo depois —
  a sequência daí pra frente é direta (mapa → pergunta OC → pergunta contrato).

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
     mandar aos fornecedores, e responder com a frase padrão **"Aguardo o envio dos
     orçamentos corrigidos/equalizados."**, e parar aqui. Quando ele reanexar os
     orçamentos atualizados, repetir a Etapa 1 do zero nesse novo lote — é um loop:
     só sai dele quando os orçamentos convergirem ou quando Bruno escolher
     "prosseguir mesmo assim".

Isso é comparar escopo, não só preço: um orçamento mais barato cotando menos
quantidade ou material diferente não é "mais barato", é outra coisa — nunca tratar
como comparável sem avisar.

### Etapa 2 — Mapa de cotação
Montar `templates/mapa_cotacao_modelo.xlsx` preenchido, um fornecedor por coluna,
seguindo a tripla conferência (mapear item a item por descrição técnica completa,
nunca aproximar "no chute"; reportar qualquer suposição feita). Antes de entregar,
aplicar o checklist de qualidade visual abaixo. Entregar o arquivo.

### Checklist de qualidade visual do .xlsx (mapa e OC, sempre antes de entregar)
Testado e validado no caso Prumus/item 15 — bugs reais que passaram batido na
primeira entrega e foram corrigidos numa segunda rodada:
1. **Fórmula mostrando R$ 0,00 ou em branco**: o LibreOffice headless não funciona
   neste ambiente (`soffice --convert-to ...` falha até em arquivos triviais — não
   perder tempo tentando de novo, é limitação do sandbox). Sem recalcular, o
   `openpyxl` grava a fórmula mas nenhum valor em cache, e qualquer visualizador que
   não recalcule ao vivo (preview do Files/Quick Look no iPhone, por exemplo) mostra
   0 ou vazio em vez do resultado. Corrigir escrevendo o valor calculado direto no
   XML junto da fórmula (o próprio `openpyxl` não expõe isso na API — precisa abrir
   o `.xlsx` como zip, achar o `<c r="CELULA">...<f>...</f></c>` na
   `xl/worksheets/sheetN.xml` e inserir `<v>VALOR</v>` antes de `</c>`). Fazer isso
   para toda fórmula de subtotal/total do arquivo. Depois, reabrir com
   `data_only=True` e conferir que os valores batem com a conta manual — nunca
   entregar confiando só na fórmula sem essa conferência.
2. **Linhas de item sem uso**: quando há menos itens do que linhas no template
   (ex.: 1 item só), **ocultar (`hidden = True`) não resolve** — o visualizador do
   iPhone/Files do Bruno ignora esse atributo e mostra as linhas em branco do mesmo
   jeito. A correção de verdade é **remover fisicamente** as linhas extras e
   deslocar o rodapé (subtotal/desconto/frete/total/condições/observação) pra cima,
   já que `delete_rows` sozinho quebra merged cells e não corrige fórmulas. O jeito
   seguro (validado no mapa e na OC do caso Prumus):
   1. Antes de mudar qualquer coisa, capturar do rodapé original: valor/fórmula de
      cada célula, estilo (`font`/`fill`/`border`/`alignment`/`number_format`),
      altura de cada linha, e todos os `merged_cells.ranges` que caem dentro do
      rodapé (guardar) e dentro das linhas de item que vão sobrar sem uso
      (só desfazer o merge, não precisa guardar).
   2. Desfazer (`unmerge_cells`) todos os merges da região que vai ser descartada
      (linhas de item não usadas + rodapé inteiro).
   3. Limpar (`.value = None`) todas as células dessa região.
   4. Reescrever o rodapé capturado nas novas linhas (`linha_original - deslocamento`,
      onde `deslocamento = primeira_linha_do_rodapé - (primeira_linha_de_item +
      qtd_itens_usados)`), reaplicando valor/estilo/altura.
   5. Nas fórmulas reescritas, ajustar as referências de linha: números de linha
      que caíam dentro do rodapé descontam o deslocamento; números de linha que
      caíam no intervalo de itens (ex. `H32:H36`) colapsam para a única linha de
      item usada (`H32:H32`); qualquer outra referência (ex. `D9`, `D23` — dados
      fixos acima da tabela) fica igual.
   6. Recriar os merges do rodapé nas linhas novas.
   7. Injetar o cache de fórmula (item 1 deste checklist) e reabrir pra conferir.
3. **Texto de descrição/condições "estourando" pra fora da linha**: célula com
   `wrapText=True` mas linha com altura padrão (`row_dimensions[r].height = None`)
   ou com `wrapText` nem definido só cabe uma linha de texto — um texto mais longo
   transborda visualmente pra fora do desenho da célula em qualquer visualizador.
   Conferir **toda** célula onde for escrever texto (descrição do item, condições
   de entrega, condições de pagamento, observação) — não só a descrição do item —
   e, se não tiver `wrapText=True` já no template, setar explicitamente
   (`Alignment(wrap_text=True, ...)`, preservando `horizontal`/`vertical` originais)
   e dar altura generosa (~15pt por linha esperada de texto). Alguns templates (a
   OC, por exemplo) já vêm com altura generosa de fábrica — conferir antes de
   assumir que precisa mudar.
4. **Cotação única (só 1 fornecedor) no mapa**: as colunas do fornecedor 2 e do
   fornecedor 3 (mapa: `L`/`M` e `N`/`O`) ficam **totalmente em branco** — sem
   `[FORNECEDOR 2]`/`[FORNECEDOR 3]`, sem `R$ 0,00` nas linhas de subtotal/total,
   sem o texto padrão de condições/observação do template. Limpar o valor de toda
   célula dessas colunas do cabeçalho até o fim do bloco de observação (não só as
   linhas de fórmula) — de novo, ocultar a coluna não é suficiente porque o
   visualizador do Bruno não respeita coluna oculta; deixar realmente em branco
   resolve nos dois casos.
5. **Texto da Observação**: manter só a condição de pagamento (parcelas, valores,
   datas) — nada de comentário extra (cotação única, validade da proposta, aviso
   de dado fictício de teste). Esse tipo de nota, quando necessário, já está
   marcado nos próprios campos do fornecedor (ex. "(FICTÍCIO)" ao lado de cada
   dado) — não precisa repetir na Observação.

Sem uma ferramenta de renderização neste ambiente, a única forma de "ver como
ficou" é reabrir o arquivo salvo e inspecionar programaticamente: valores calculados
(`data_only=True`), conteúdo célula a célula da região de itens/rodapé, altura de
linha, e `merged_cells.ranges`. Fazer essa inspeção sempre, não só quando o usuário
reclamar — e não confiar em `hidden` pra esconder linha/coluna, já provou não
funcionar no visualizador que o Bruno usa.

### Etapa 3 — Escolha do vencedor
Perguntar qual fornecedor ganhou, se Bruno ainda não tiver dito. Em seguida, usar
AskUserQuestion: **"Deseja que eu já emita a Ordem de Compra?"** (sim/não). Se
**não**, parar aqui — o mapa já foi entregue, Bruno chama de novo quando quiser a
OC. Se **sim**, seguir para a Etapa 4.

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

Depois de entregar a OC, usar AskUserQuestion: **"Deseja que eu já emita o
contrato para essa prestação de serviço?"** (sim/não). Se **não**, parar aqui. Se
**sim**, seguir para a Etapa 6.

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
