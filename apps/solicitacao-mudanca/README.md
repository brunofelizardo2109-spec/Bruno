# Agente de Solicitação de Mudança — Grupo A

Formulário próprio → Gemini sugere tipo, dimensões e resumo → linha nova no
[Log de Solicitações - Grupo A](https://docs.google.com/spreadsheets/d/1c_RtRyhsDIeHzxEcknnD4yoeySHKsbz3wkpyvEyQqgg/edit).
Mesmo fluxo do guia (n8n), rodando em Google Apps Script: sem trial, sem mensalidade, chave fora do código.
**A IA sugere; o CCB decide.**

| Guia (n8n) | Este app |
|---|---|
| On form submission | `doGet` + `Index.html` |
| HTTP Request (Gemini) | `sugerirComGemini_` |
| Edit Fields | `interpretar_` |
| Append row in sheet | `registrarSolicitacao` |
| Credencial Header Auth | Propriedade `GEMINI_API_KEY` |

## Implantar (≈10 min, uma vez)

1. Abra a planilha → **Extensões → Apps Script**.
2. Crie os arquivos com o mesmo nome e cole o conteúdo: `Code.gs`, `Index.html` (Arquivo + → HTML, nome `Index`).
3. **Configurações do projeto** → marque *Mostrar "appsscript.json"* → cole `appsscript.json`.
4. Ainda em Configurações → **Propriedades do script** → adicionar `GEMINI_API_KEY` = chave do AI Studio (Passo 1 do guia). Opcional: `GEMINI_MODEL` para trocar o modelo sem mexer no código.
5. No editor, selecione `setup` → **Executar** → autorize. Depois `testarGemini` → confira o JSON no Registro de execução.
6. **Implantar → Nova implantação → App da Web**: Executar como *Eu*; Acesso *Qualquer pessoa*. Copie a URL `/exec` — é o endereço do formulário.
7. Mudou o código? **Implantar → Gerenciar implantações → editar → Nova versão** (a URL se mantém).

## Teste 2 (critério de aceite do guia)

| Caso | Esperado |
|---|---|
| 1 · Pedido claro (Ana Souza) | Linha com tipo e dimensões coerentes |
| 2 · Pedido vago (Marcos) | Resumo aponta o que falta |
| 3 · Aspas e quebra de linha (Carlos Lima) | Texto intacto na planilha |
| Extra · Gemini fora do ar / sem chave | Linha gravada com "IA indisponível", pedido não se perde |

## Segurança

- Chave só em Propriedades do script; nunca no código, print ou chat. Vazou → apague no AI Studio, gere outra, troque a propriedade.
- Texto começando com `= + - @` é gravado como texto (bloqueia injeção de fórmula).
- Campo isca anti-robô + limite de envios por minuto + limite de caracteres.
- Colunas A–F avisam antes de edição; G–I (IA) em amarelo; Status com lista fechada.
- Plano gratuito do Gemini pode usar o conteúdo enviado: só dados fictícios/não sensíveis.

## Domínio próprio (opcional)

A URL `script.google.com/macros/s/.../exec` é longa. Para um endereço do grupo:
Google Sites (gratuito, `sites.google.com/view/grupo-a-mudancas`) com o app incorporado, ou domínio comprado
(ex.: Registro.br) apontando para o Site. Encurtadores de link não são recomendados: escondem o destino.
