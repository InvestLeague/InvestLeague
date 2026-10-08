// Gera formulario.html (preenchimento sem login) a partir de index.html.
// Uso: node gerar-formulario.js
const fs = require("fs");
let s = fs.readFileSync(__dirname + "/index.html", "utf8");
const troca = (a, b) => { if (!s.includes(a)) throw new Error("Trecho não encontrado: " + a); s = s.replace(a, b); };
troca("<title>Afastamentos Infantaria EsAO</title>", "<title>Formulário de Afastamentos EsAO</title>");
troca("const AVULSO = false;", "const AVULSO = true;");
fs.writeFileSync(__dirname + "/formulario.html", s);
console.log("formulario.html gerado");
