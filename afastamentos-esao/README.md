# Levantamento de afastamentos · Curso de Infantaria da EsAO

Formulário para celular e painel de disponibilidade dos 13 militares, de 01/11/2026 a 31/01/2027.
Um único arquivo (`index.html`) funciona em duas hospedagens:

| Hospedagem | Base de dados | Quem consegue responder |
|---|---|---|
| **Artifact do Claude** (já publicado) | Base compartilhada do artifact, atualizada em tempo real | Quem tiver conta claude.ai e receber o link com permissão de edição |
| **Planilha Google + Apps Script** (recomendado para quem não tem Claude) | Planilha Google, com histórico de envios | Qualquer pessoa com o link, sem login |

Aberto direto no navegador (sem hospedagem), o arquivo entra em **modo demonstração** e salva só no próprio aparelho.

## Quem vê o quê

- **Militares** abrem o link e veem **só o formulário**. Não aparecem o painel, os botões de troca de tela nem o status dos colegas. Cada um só lê a própria resposta.
- **Maj Modesto** (dono do artifact, ou quem tiver a chave do painel no Apps Script) vê o formulário e o painel completo.
- No artifact isso é garantido pelas regras da base: `respostas` só é lida pelo dono; cada conta grava e lê apenas `respostas/<id da conta>`. Nem quem tem permissão de Editor consegue ler as respostas dos outros.

## Como funciona

- Cada militar tem **um registro próprio** (documento no artifact, linha na planilha). Envios simultâneos de militares diferentes nunca se sobrescrevem.
- Cada registro tem um número de versão. Se a mesma resposta for alterada em dois aparelhos, o segundo envio é recusado e o militar escolhe qual versão manter.
- As regras (dispensas de Natal/Ano-Novo, datas invertidas, campos incompletos, sobreposição, desligamento) estão em `validar()` no `index.html` e são recalculadas no painel a cada carga.
- Respostas com conflito podem ser enviadas após confirmação e aparecem no painel como **Precisa corrigir**.
- Quem não respondeu aparece como **Sem resposta** e nunca é contado como disponível. A partir da data de desligamento o militar sai do efetivo.
- Link direto para o painel: acrescente `#painel` ao link (artifact; só abre para o dono) ou `?v=painel&chave=<PAINEL_CHAVE>` (Apps Script).
- No Apps Script não há login: quem escolher um nome no formulário consegue carregar a resposta daquele nome para corrigir. O painel e a lista completa exigem a chave.

## Publicar na Planilha Google (link sem login)

Tudo vai num arquivo só: `Codigo.gs` (gerado com `node gerar-apps-script.js` a partir de `Code.gs` + `index.html`).

1. Abra **sheets.new** e dê um nome à planilha (ex.: "Afastamentos Inf EsAO").
2. Menu **Extensões › Apps Script**. Apague o conteúdo do editor, cole o conteúdo inteiro de `Codigo.gs` e salve (Ctrl+S).
3. **Implantar › Nova implantação**. Na engrenagem, escolha **App da Web**. Executar como: **Eu**. Quem pode acessar: **Qualquer pessoa**. Clique em **Implantar** e autorize (em "O Google não verificou este app", clique em **Avançado › Acessar**; o app é seu).
4. Volte à planilha e recarregue a página. No menu **Afastamentos › Ver links do formulário e do painel** aparecem:
   - o link do **formulário**, para enviar aos militares (sem login, grava direto na planilha);
   - o link do **seu painel**, com a chave (não repasse).

No primeiro acesso, o sistema cria as abas `Respostas` e `Historico`, gera a chave do painel e lança os dados do Maj Modesto.
Ao alterar algo depois, cole o novo `Codigo.gs` e use **Implantar › Gerenciar implantações › Editar › Nova versão** para manter o mesmo link.
