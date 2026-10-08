// Gera Codigo.gs: um único arquivo para colar no Apps Script (backend + página embutida).
// Uso: node gerar-apps-script.js   (rode de novo sempre que alterar index.html ou Code.gs)
const fs = require("fs");
const dir = __dirname;
let gs = fs.readFileSync(dir + "/Code.gs", "utf8");
const html = fs.readFileSync(dir + "/index.html", "utf8");
const de = "HtmlService.createHtmlOutputFromFile('index')";
if (!gs.includes(de)) throw new Error("doGet não encontrado em Code.gs");
gs = gs.replace(de, "HtmlService.createHtmlOutput(PAGINA)");
gs = "// ARQUIVO GERADO por gerar-apps-script.js a partir de Code.gs + index.html. Cole inteiro no Apps Script.\n" + gs +
  "\n/** Página do formulário e do painel (index.html). */\nvar PAGINA = " + JSON.stringify(html) + ";\n";
fs.writeFileSync(dir + "/Codigo.gs", gs);
console.log("Codigo.gs gerado (" + Math.round(gs.length / 1024) + " KB)");
