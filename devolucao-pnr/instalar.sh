#!/usr/bin/env bash
# Cria a planilha "Entrega PNR 2026", envia o código e publica o sistema.
# Roda uma única vez. Para atualizar depois, use ./publicar.sh
#
#   chmod +x instalar.sh publicar.sh
#   ./instalar.sh

set -euo pipefail
cd "$(dirname "$0")"

TITULO="${TITULO:-Entrega PNR 2026}"
AMARELO=$'\033[33m'; VERDE=$'\033[32m'; VERM=$'\033[31m'; FIM=$'\033[0m'

passo() { printf '\n%s==> %s%s\n' "$AMARELO" "$1" "$FIM"; }
erro()  { printf '\n%sERRO: %s%s\n' "$VERM" "$1" "$FIM" >&2; exit 1; }

passo "1/6 · Conferindo o que está instalado"
command -v node >/dev/null 2>&1 || erro "Node não encontrado. Instale com: brew install node"
printf 'node %s\n' "$(node --version)"

if ! command -v clasp >/dev/null 2>&1; then
  echo "clasp não encontrado. Instalando…"
  npm install -g @google/clasp@2.4.2 || erro "Falhou instalar o clasp. Tente: sudo npm install -g @google/clasp@2.4.2"
fi
printf 'clasp %s\n' "$(clasp --version)"

passo "2/6 · Entrando na sua conta Google"
if clasp login --status >/dev/null 2>&1; then
  echo "Já está conectado: $(clasp login --status 2>/dev/null | tail -1)"
else
  echo "Vai abrir o navegador uma vez para autorizar o clasp na sua conta."
  clasp login || erro "Login no clasp não concluído."
fi

passo "3/6 · Gerando o arquivo do Apps Script"
node gerar-apps-script.js

passo "4/6 · Criando a planilha \"$TITULO\" e o projeto"
mkdir -p build
if [ -f build/.clasp.json ]; then
  echo "Projeto já existe (build/.clasp.json). Pulando a criação."
  echo "Se quiser começar de zero, apague a pasta build e rode de novo."
else
  ( cd build && clasp create --type sheets --title "$TITULO" --rootDir . ) \
    || erro "Não foi possível criar o projeto. Se a conta tiver a API do Apps Script desligada, ative em https://script.google.com/home/usersettings"
fi

# O código e as configurações são sempre os da pasta principal.
cp Codigo.gs build/Codigo.gs
cp appsscript.json build/appsscript.json

passo "5/6 · Enviando o código"
( cd build && clasp push -f ) || erro "Falhou o envio do código."

passo "6/6 · Publicando o app da web"
SAIDA="$( cd build && clasp deploy -d "versão inicial" 2>&1 )" || erro "Falhou a publicação. Saída: $SAIDA"
echo "$SAIDA"

ID="$(printf '%s' "$SAIDA" | grep -oE 'AKfyc[A-Za-z0-9_-]+' | head -1)"
if [ -z "$ID" ]; then
  echo "${AMARELO}Não consegui ler o identificador da publicação.${FIM}"
  echo "Rode: cd build && clasp deployments"
else
  printf '%s' "$ID" > .deployid
  URL="https://script.google.com/macros/s/$ID/exec"
  printf '\n%sPronto.%s\n\n' "$VERDE" "$FIM"
  echo "Endereço do sistema:"
  echo "  $URL"
  echo
  echo "FALTA UM PASSO, e só o senhor pode dar:"
  echo "  1. Abra o endereço acima no navegador."
  echo "  2. O Google vai pedir autorização. Em \"O Google não verificou este app\","
  echo "     clique em Avançado › Acessar. O app é seu."
  echo "  3. Depois abra a planilha \"$TITULO\" no seu Drive e recarregue a página."
  echo "     No menu PNR: \"Ver links\" e \"Ativar lembretes diários\"."
  echo
  echo "Para abrir a planilha agora:  cd build && clasp open --addon"
fi
