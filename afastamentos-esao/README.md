# Levantamento de afastamentos · Curso de Infantaria da EsAO

Formulário para celular e painel de disponibilidade dos 13 militares, de 01/11/2026 a 31/01/2027.
Um único arquivo (`index.html`) funciona em duas hospedagens:

| Hospedagem | Base de dados | Quem consegue responder |
|---|---|---|
| **Artifact do Claude** (já publicado) | Base compartilhada do artifact, atualizada em tempo real | Quem tiver conta claude.ai e receber o link com permissão de edição |
| **Google Apps Script** (opcional) | Planilha Google, com histórico de envios | Qualquer pessoa com o link, sem login |

Aberto direto no navegador (sem hospedagem), o arquivo entra em **modo demonstração** e salva só no próprio aparelho.

## Como funciona

- Cada militar tem **um registro próprio** (documento no artifact, linha na planilha). Envios simultâneos de militares diferentes nunca se sobrescrevem.
- Cada registro tem um número de versão. Se a mesma resposta for alterada em dois aparelhos, o segundo envio é recusado e o militar escolhe qual versão manter.
- As regras (dispensas de Natal/Ano-Novo, datas invertidas, campos incompletos, sobreposição, desligamento) estão em `validar()` no `index.html` e são recalculadas no painel a cada carga.
- Respostas com conflito podem ser enviadas após confirmação e aparecem no painel como **Precisa corrigir**.
- Quem não respondeu aparece como **Sem resposta** e nunca é contado como disponível. A partir da data de desligamento o militar sai do efetivo.
- Link direto para o painel: acrescente `#painel` ao link (artifact) ou `?v=painel` (Apps Script).

## Publicar no Google Apps Script (link sem login)

1. Crie uma Planilha Google nova (ex.: "Afastamentos Inf EsAO").
2. Menu **Extensões › Apps Script**.
3. No arquivo `Código.gs`, apague o conteúdo e cole o conteúdo de `Code.gs`.
4. Clique em **+ › HTML**, nomeie `index` e cole o conteúdo de `index.html`.
5. Selecione a função `instalar` e clique em **Executar** (autorize quando pedir). Isso cria as abas e lança os dados do Maj Modesto.
6. **Implantar › Nova implantação › Tipo: App da Web**. Executar como: **Eu**. Quem pode acessar: **Qualquer pessoa** (ou "Qualquer pessoa com Conta do Google").
7. Copie a URL que termina em `/exec` e envie aos militares. O painel fica em `URL/exec?v=painel`.

Ao alterar o `index.html` depois, use **Implantar › Gerenciar implantações › Editar › Nova versão** para manter a mesma URL.
