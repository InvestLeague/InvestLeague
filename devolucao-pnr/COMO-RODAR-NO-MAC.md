# Rodar no seu Mac, pelo terminal

Tudo por linha de comando, sem clicar em telas de configuração. São dois scripts:
`instalar.sh` (uma vez) e `publicar.sh` (a cada alteração).

## Por que ainda usa o Google e não só o seu Mac

Os 107 alunos vão abrir o link **no celular deles**, de qualquer lugar e a qualquer hora. Para isso
o sistema precisa de um endereço na internet. Se rodasse apenas no seu Mac, só funcionaria enquanto
o computador estivesse ligado e na mesma rede, e ninguém de fora alcançaria.

Então a divisão fica assim:

| Onde | O que fica |
|---|---|
| **Seu Mac** | O código, a edição e os comandos de publicação |
| **Sua conta Google** | A hospedagem do link e a planilha com os dados |

Para ver o sistema sem publicar nada, abra o `index.html` direto no navegador: ele entra em
**modo demonstração**, com dados só naquele aparelho e um botão para alternar entre as três telas.

```bash
open index.html
```

## Instalar (uma vez)

```bash
cd <a sua pasta>
git clone -b devolucao-pnr https://github.com/InvestLeague/InvestLeague.git
cd InvestLeague/devolucao-pnr
chmod +x instalar.sh publicar.sh
./instalar.sh
```

O script faz, nesta ordem:

1. confere o Node e instala o `clasp` se faltar;
2. conecta na sua conta Google (abre o navegador **uma vez**, só para autorizar o `clasp`);
3. gera o `Codigo.gs` a partir do `index.html` e do `Code.gs`;
4. **cria a planilha "Entrega PNR 2026"** no seu Drive, junto com o projeto;
5. envia o código;
6. publica o app da web, já como "executar como eu" e "qualquer pessoa pode acessar", e imprime o
   endereço do sistema.

No fim ele mostra o endereço e o único passo que depende de você.

## O passo que só você pode dar

Depois do script, abra o endereço que ele imprimiu. O Google vai pedir autorização para a planilha,
o Gmail e a Agenda. Em "O Google não verificou este app", clique em **Avançado › Acessar**: o app é
seu. Essa tela é uma proteção da conta Google, e nenhuma automação pode respondê-la no seu lugar.
É um clique, uma única vez.

Depois disso, abra a planilha e use o menu **PNR**:

- **Ver links** — o seu painel, a tela e o link dos orientados de cada instrutor, e o link geral;
- **Ativar lembretes diários** — cria os avisos das 18 h e das 7 h.

## Publicar uma alteração

```bash
./publicar.sh
```

Mantém o **mesmo endereço**, então os links já distribuídos aos instrutores e aos alunos continuam
valendo. Nunca apague o arquivo `.deployid`: é ele que guarda essa amarração.

## Comandos úteis

```bash
cd build && clasp open              # abre o editor do Apps Script no navegador
cd build && clasp open --addon      # abre a planilha "Entrega PNR 2026"
cd build && clasp deployments       # lista as publicações
cd build && clasp logs              # erros de execução, se algo falhar
```

## Se algo der errado

| Mensagem | O que fazer |
|---|---|
| `User has not enabled the Apps Script API` | Ative em https://script.google.com/home/usersettings e rode de novo |
| `EACCES` ao instalar o clasp | `sudo npm install -g @google/clasp@2.4.2` |
| `command not found: node` | `brew install node` |
| `clasp: command not found` após instalar | Feche e abra o terminal |

Em qualquer erro, copie a saída do terminal e me mande: eu corrijo o código ou o script.

## O que não vai para o Git

`build/` e `.deployid` ficam de fora: eles contêm o vínculo com a **sua** planilha e a **sua**
publicação. Clonando o repositório em outro computador, rode `./instalar.sh` para criar os seus.
