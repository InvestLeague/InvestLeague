# Levantamento de afastamentos · Curso de Infantaria da EsAO

Formulário para celular e painel de disponibilidade dos 13 militares, de 01/11/2026 a 31/01/2027.
Um único arquivo (`index.html`) funciona em duas hospedagens:

| Hospedagem | Base de dados | Quem consegue responder |
|---|---|---|
| **Artifact do Claude** (já publicado) | Base compartilhada do artifact, atualizada em tempo real | Quem tiver conta claude.ai e receber o link com permissão de edição |
| **`formulario.html`** (sem login) | Nenhuma: gera uma mensagem com código que o militar envia ao Maj Modesto (WhatsApp/e-mail); o gestor cola no painel em "Respostas recebidas por mensagem" | Qualquer pessoa que consiga abrir a página |
| **Google Apps Script** (opcional) | Planilha Google, com histórico de envios | Qualquer pessoa com o link, sem login |

Aberto direto no navegador (sem hospedagem), o arquivo entra em **modo demonstração** e salva só no próprio aparelho.

## Quem vê o quê

- **Militares** abrem o link e veem **só o formulário**. Não aparecem o painel, os botões de troca de tela nem o status dos colegas. Cada um só lê a própria resposta.
- **Maj Modesto** (dono do artifact, ou quem tiver a chave do painel no Apps Script) vê o formulário e o painel completo.
- No artifact isso é garantido pelas regras da base: `respostas` só é lida pelo dono; cada conta grava e lê apenas `respostas/<id da conta>`. Nem quem tem permissão de Editor consegue ler as respostas dos outros.

## Formulário sem login (`formulario.html`)

- Gerado a partir do `index.html` com `node gerar-formulario.js` (muda só `AVULSO = true` e o título). Depois de alterar o `index.html`, gere de novo.
- Arquivo único, sem servidor: pode ser publicado como artifact com link público, hospedado em qualquer site estático (GitHub Pages, Google Sites) ou enviado como arquivo.
- No final, o militar toca em **Enviar pelo WhatsApp** (ou copia a mensagem). A mensagem traz o resumo legível e um código `AFX1.…`.
- A resposta também fica guardada no aparelho do militar; para corrigir, ele reabre o formulário, ajusta e envia nova mensagem. No painel vale a mais recente; mensagens antigas coladas de novo são ignoradas.

## Como funciona

- Cada militar tem **um registro próprio** (documento no artifact, linha na planilha). Envios simultâneos de militares diferentes nunca se sobrescrevem.
- Cada registro tem um número de versão. Se a mesma resposta for alterada em dois aparelhos, o segundo envio é recusado e o militar escolhe qual versão manter.
- As regras (dispensas de Natal/Ano-Novo, datas invertidas, campos incompletos, sobreposição, desligamento) estão em `validar()` no `index.html` e são recalculadas no painel a cada carga.
- Respostas com conflito podem ser enviadas após confirmação e aparecem no painel como **Precisa corrigir**.
- Quem não respondeu aparece como **Sem resposta** e nunca é contado como disponível. A partir da data de desligamento o militar sai do efetivo.
- Link direto para o painel: acrescente `#painel` ao link (artifact; só abre para o dono) ou `?v=painel&chave=<PAINEL_CHAVE>` (Apps Script).
- No Apps Script não há login: quem escolher um nome no formulário consegue carregar a resposta daquele nome para corrigir. O painel e a lista completa exigem a chave.

## Publicar no Google Apps Script (link sem login)

1. Crie uma Planilha Google nova (ex.: "Afastamentos Inf EsAO").
2. Menu **Extensões › Apps Script**.
3. No arquivo `Código.gs`, apague o conteúdo e cole o conteúdo de `Code.gs`.
4. Clique em **+ › HTML**, nomeie `index` e cole o conteúdo de `index.html`.
5. Selecione a função `instalar` e clique em **Executar** (autorize quando pedir). Isso cria as abas, lança os dados do Maj Modesto e gera a chave do painel (aparece em **Registro de execução**; fica também em Configurações do projeto › Propriedades do script › `PAINEL_CHAVE`).
6. **Implantar › Nova implantação › Tipo: App da Web**. Executar como: **Eu**. Quem pode acessar: **Qualquer pessoa** (ou "Qualquer pessoa com Conta do Google").
7. Copie a URL que termina em `/exec` e envie aos militares (só formulário). O seu painel fica em `URL/exec?v=painel&chave=<PAINEL_CHAVE>`; não repasse esse link.

Ao alterar o `index.html` depois, use **Implantar › Gerenciar implantações › Editar › Nova versão** para manter a mesma URL.
