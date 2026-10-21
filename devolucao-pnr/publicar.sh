#!/usr/bin/env bash
# Publica uma alteração mantendo o MESMO endereço, para não refazer os links.
# Use sempre que o index.html ou o Code.gs mudarem.
#
#   ./publicar.sh

set -euo pipefail
cd "$(dirname "$0")"

AMARELO=$'\033[33m'; VERDE=$'\033[32m'; VERM=$'\033[31m'; FIM=$'\033[0m'
passo() { printf '\n%s==> %s%s\n' "$AMARELO" "$1" "$FIM"; }
erro()  { printf '\n%sERRO: %s%s\n' "$VERM" "$1" "$FIM" >&2; exit 1; }

[ -f build/.clasp.json ] || erro "O projeto ainda não existe. Rode ./instalar.sh primeiro."

passo "1/3 · Gerando o arquivo do Apps Script"
node gerar-apps-script.js
cp Codigo.gs build/Codigo.gs
cp appsscript.json build/appsscript.json

passo "2/3 · Enviando o código"
( cd build && clasp push -f ) || erro "Falhou o envio."

passo "3/3 · Republicando no mesmo endereço"
if [ -f .deployid ]; then
  ID="$(cat .deployid)"
  ( cd build && clasp deploy -i "$ID" -d "atualização $(date '+%d/%m %H:%M')" ) \
    || erro "Falhou a republicação."
  printf '\n%sPronto. Mesmo endereço, nova versão.%s\n' "$VERDE" "$FIM"
  echo "  https://script.google.com/macros/s/$ID/exec"
  echo
  echo "Os links já distribuídos continuam valendo."
else
  echo "${AMARELO}Sem o arquivo .deployid: vou criar uma publicação nova.${FIM}"
  echo "Atenção: isso gera um endereço DIFERENTE, e os links antigos param de funcionar."
  ( cd build && clasp deploy -d "atualização $(date '+%d/%m %H:%M')" )
  echo "Para ver as publicações:  cd build && clasp deployments"
fi
