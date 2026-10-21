# Agendamento de devolução de PNR · EsAO 2026

Plataforma onde cada aluno agenda, pelo celular, a vistoria de entrega do seu PNR com o instrutor
responsável, e o instrutor confirma. Sem cadastro e sem senha: basta o link.

Mesma stack do levantamento de afastamentos (`../afastamentos-esao`): uma página única em
HTML/JavaScript servida por um Google Apps Script ligado a uma Planilha Google.

| Arquivo | O que é |
|---|---|
| `index.html` | A página inteira: tela do aluno, do instrutor e da coordenação |
| `Code.gs` | Backend: agenda, regras, e-mails e gravação na planilha |
| `gerar-apps-script.js` | Junta os dois num `Codigo.gs` só, para colar no editor |
| `Codigo.gs` | **Arquivo gerado.** É este que vai para o Apps Script |
| `instalar.sh` · `publicar.sh` | Instalação e atualização pelo terminal do Mac, sem clicar em telas |
| `appsscript.json` | Fuso, permissões e configuração do app da web |
| `COMO-RODAR-NO-MAC.md` | **Comece por aqui** se for publicar pelo terminal |

Para testar sem publicar nada, abra o `index.html` direto no navegador: ele entra em
**modo demonstração**, com dados salvos só naquele aparelho, e um botão para alternar entre as
três telas.

## Como os links circulam

A coordenação divulga **apenas aos instrutores**. Cada instrutor, na tela dele, copia o
**link dos seus orientados** e repassa à sua turma. Quem abre o link de um instrutor **só consegue
agendar com aquele instrutor**: não escolhe na lista e não cai na agenda de outro, nem alterando o
endereço, porque o servidor refaz essa amarração a cada pedido.

```
Coordenação ──(link da tela do instrutor)──► Instrutor ──(link dos orientados)──► Alunos dele
```

| Perfil | Link | Acesso |
|---|---|---|
| **Coordenação** | `.../exec?v=painel&chave=<CHAVE>` | Chave de administrador |
| **Instrutor** | `.../exec?v=instrutor&chave=<CHAVE>` | Um link pessoal por instrutor |
| **Aluno de um instrutor** | `.../exec?i=<CHAVE>` | Aberto, já amarrado àquele instrutor |
| **Link geral** | `.../exec` | Reserva: o aluno escolhe o instrutor na lista |

Cada instrutor tem **duas chaves diferentes**, nas colunas 5 e 6 da aba `Instrutores`: a da tela
dele e a do link dos orientados. A segunda não dá nenhum poder de instrutor, e a primeira nunca
aparece no painel, justamente para não ser repassada por engano.

O **link geral** existe só para quem ficou sem o link do seu instrutor. Nele o aluno escolhe na
lista e, se errar, o instrutor transfere o pedido ao responsável certo.

O aluno nunca vê o nome de outro aluno: na agenda do instrutor ele só vê que o horário está
ocupado. Instrutor e coordenação veem os nomes.

## Regras que o sistema aplica sozinho

**Os 6 tempos do dia** (vistoria de 1 h 15, com 15 min de intervalo):

| Manhã | Tarde |
|---|---|
| 08h00 – 09h15 | 13h30 – 14h45 |
| 09h30 – 10h45 | 15h00 – 16h15 |
| 11h00 – 12h15 | 16h30 – 17h45 |

- **Período:** de hoje até 30/12/2026, ajustável na aba `Config`.
- **Bloqueados:** sábados e domingos, 02/11, 20/11, 25/12 e a **tarde de 24/12**. 28/10 fica aberto.
- **Disponibilidade do instrutor:** nenhum horário aparece fora dela. Entende `Até 04/12`,
  `13/11 a 26/11`, `Até 19/11; 30/11 a 22/12` e `03/12 e 04/12`.
- **Antecedência mínima de 72 h** para o aluno agendar sozinho. Abaixo disso o horário só sai
  com o aval do instrutor.
- **Data da viagem:** o sistema **impede** marcar no dia do embarque ou depois. 2 a 3 dias antes
  aparece marcado como prazo ideal; 1 dia antes passa com aviso.
- **Um agendamento ativo por aluno**, e nunca dois alunos no mesmo tempo (gravação sob
  `LockService`, com número de versão).

## Quando o aluno esqueceu de agendar

Três saídas, todas na aba **Ferramentas** do instrutor:

1. **Pedido de encaixe** — o aluno toca no horário amarelo, justifica, e o instrutor aprova ou não.
2. **Marcar pelo aluno** — o instrutor agenda direto, já confirmado.
3. **Liberar exceção** — dispensa aquele aluno das 72 h por alguns dias.

Em qualquer uma delas continuam valendo duas regras: nunca no dia da viagem e nunca depois da
disponibilidade do instrutor.

## Lembretes

Para os instrutores que cadastrarem e-mail:

- **ao confirmar:** evento criado na Agenda Google, com convite, mais um `.ics` que entra na agenda
  do iPhone e do Android com alarmes na véspera e 1 h antes;
- **todos os dias às 18 h:** as vistorias de amanhã, com endereços;
- **todos os dias às 7 h:** as vistorias de hoje;
- **na hora:** pedido novo, encaixe, cancelamento e transferência.

O aluno também recebe, se informar um e-mail.

## Camada "à prova de erro"

- Passo a passo numerado, uma tarefa por tela, com "onde estou / o que faço / o que acontece depois".
- Horário indisponível **explica o motivo** ao toque, em português.
- Botões dizem o que fazem: *"Pedir vistoria em qui 03/12 às 09h30"*, não *"Enviar"*.
- Checklist com **"Como fazer"** em cada item, incluindo o passo a passo da Light e da Naturgy.
- Cartão **"No dia, leve e faça"** e contagem regressiva para a vistoria e para a viagem.
- Botão **"Preciso de ajuda"** em todas as telas, com as dúvidas frequentes.
- O instrutor tem um **roteiro de vistoria** no celular: marca ✓ ou ⚠ item por item, e o que ficar
  ⚠ vira pendência enviada ao aluno.
- Todo alerta do painel vem com a ação pronta, inclusive a mensagem de cobrança para o WhatsApp.

## Abas da planilha

Criadas sozinhas no primeiro acesso.

| Aba | Para que serve |
|---|---|
| `Instrutores` | Nome, e-mail, disponibilidade, total de alunos, as duas chaves e `ativo` |
| `Agendamentos` | Uma linha por agendamento, com histórico no JSON |
| `Bloqueios` | Dias e tempos bloqueados por instrutor |
| `Aceites` | Quem leu e aceitou os termos, com data, hora e versão (prova) |
| `Log` | Auditoria de tudo |
| `Config` | `inicio`, `fim` e `horasMinimas` |

**Tudo o que muda, muda na planilha, sem mexer em código:** trocar uma data de disponibilidade,
cadastrar um e-mail, tirar um instrutor da lista (`ativo` = `não`) ou mudar a antecedência mínima.
Depois de editar, basta recarregar a página.

Os 11 instrutores da escala 2026 já vêm lançados:

| Instrutor | Alunos | Disponibilidade |
|---|---|---|
| Maj Modesto | 10 | Até 04/12 |
| Maj Andrews | 10 | Até 18/12 |
| Maj Pimenta | 1 | 09/11 a 13/11 |
| Maj Calderaro | 10 | Até 19/11 |
| Maj Cesse | 11 | Até 04/12 |
| Cap Dantas | 10 | 13/11 a 26/11 |
| Cap Loan | 11 | Até 04/12 |
| Cap Santana | 11 | Até 11/12 |
| Cap Uerlei Moreira | 11 | Até 18/12 |
| Cap Cavalcanti | 11 | Até 29/12 |
| Cap Marcello | 11 | Até 19/11; 30/11 a 22/12 |

## Publicar pelo terminal (recomendado)

Dois comandos, descritos em `COMO-RODAR-NO-MAC.md`: `./instalar.sh` cria a planilha
**Entrega PNR 2026** e publica; `./publicar.sh` atualiza mantendo o mesmo endereço.

## Publicar pelo navegador (alternativa)

1. Abra **sheets.new** e nomeie a planilha **Entrega PNR 2026**.
2. Menu **Extensões › Apps Script**. Apague o conteúdo do editor, cole o **`Codigo.gs` inteiro** e
   salve com Ctrl+S.
3. Na engrenagem **Configurações do projeto**, confirme o fuso **(GMT-03:00) São Paulo**. É dele que
   dependem os horários dos lembretes.
4. **Implantar › Nova implantação**. Na engrenagem, escolha **App da Web**. Executar como: **Eu**.
   Quem pode acessar: **Qualquer pessoa**. Clique em **Implantar** e autorize. Em "O Google não
   verificou este app", clique em **Avançado › Acessar**; o app é seu.
   A autorização pede planilha, Gmail e Agenda, que são os três recursos usados.
5. Volte à planilha e recarregue a página. No menu **PNR**:
   - **Ver links** mostra o seu painel, a tela e o link dos orientados de cada instrutor, e o link geral;
   - **Ativar lembretes diários** cria os gatilhos das 18 h e das 7 h.
6. Mande a **cada instrutor**, individualmente, só a linha dele. No primeiro acesso ele já vê o
   link dos orientados com um botão para mandar no WhatsApp da turma.

Ao alterar algo depois, rode `node gerar-apps-script.js`, cole o novo `Codigo.gs` e use
**Implantar › Gerenciar implantações › Editar › Nova versão**, para manter o mesmo link.

## Antes de divulgar

Entregue o link a 3 alunos que nunca viram o sistema, com uma única instrução: "agende sua
devolução". Se os 3 terminarem sem perguntar nada a ninguém, pode divulgar para a turma. Onde
alguém travar, ajuste o texto daquela tela.

## Base legal dos termos

O Termo de Ciência e Compromisso cruza as orientações da OM com a
**Portaria C Ex nº 1.846, de 4 de outubro de 2022** (Instruções Gerais para a Administração dos PNR
do Exército), que prevê a vistoria prévia com lavratura do Termo de Desocupação, exige a
comprovação da quitação de todos os débitos na entrega e permite cobrar do permissionário os danos
causados.

**Atenção:** a Portaria admite a vistoria "na presença do permissionário **ou seu representante**".
As regras de entrega apenas pelo próprio militar e sem acompanhamento da família são **mais
restritivas** que a norma, e por isso aparecem no termo como **determinação da OM**. Vale respaldá-las
num DIEx ou nota de serviço, e conferir se a Região Militar tem normas locais com modelo próprio de
termo ou prazo mínimo de agendamento.
