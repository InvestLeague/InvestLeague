/**
 * Levantamento de afastamentos — Curso de Infantaria da EsAO
 * Backend em Google Apps Script, ligado a uma Planilha Google.
 *
 * Aba "Respostas": uma linha por militar (nunca uma linha compartilhada),
 * gravada sob LockService e com controle de versão (rev). Dois envios ao
 * mesmo tempo de militares diferentes não se sobrescrevem; dois envios do
 * mesmo militar a partir de versões diferentes são recusados e o segundo
 * aparelho é avisado.
 * Aba "Historico": cada envio fica registrado (auditoria).
 *
 * Acesso: o link comum abre só o formulário. O painel (todas as respostas)
 * exige a chave do gestor: URL/exec?v=painel&chave=<PAINEL_CHAVE>, criada por instalar().
 */

var ABA = 'Respostas';
var ABA_HIST = 'Historico';
var CAB = ['id', 'militar', 'rev', 'atualizadoEm', 'dispensa', 'desligamento', 'afastamentos (leitura)', 'dados (JSON - não editar)'];
var MILITARES = {
  'barroso-magno': 'Maj Barroso Magno', 'modesto': 'Maj Modesto', 'andrews': 'Maj Andrews', 'pimenta': 'Maj Pimenta',
  'aredes': 'Maj Arêdes', 'calderaro': 'Cap Calderaro', 'cesse': 'Cap Cesse', 'dantas': 'Cap Dantas', 'loan': 'Cap Loan',
  'santana': 'Cap Santana', 'uerlei-moreira': 'Cap Uerlei Moreira', 'cavalcanti': 'Cap Cavalcanti', 'marcello': 'Cap Marcello'
};
var NOMES_CAT = {
  ferias26: 'Férias 2026', feriasant: 'Férias de ano anterior', natal: 'Dispensa de Natal', anonovo: 'Dispensa de Ano-Novo',
  instalacao: 'Instalação', dispcmt: 'Dispensa Cmt EsAO', curso: 'Curso/estágio', missao: 'Missão/serviço fora da EsAO',
  licenca: 'Licença/afastamento já previsto', outro: 'Outro afastamento autorizado'
};

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('Afastamentos Infantaria EsAO')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function planilha_() {
  var id = PropertiesService.getScriptProperties().getProperty('PLANILHA_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
}

function aba_(nome, cab) {
  var ss = planilha_();
  var sh = ss.getSheetByName(nome);
  if (!sh) {
    sh = ss.insertSheet(nome);
    sh.getRange(1, 1, 1, cab.length).setValues([cab]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** Todas as respostas: somente para o painel, com a chave do gestor. */
function listar(chave) {
  var certa = PropertiesService.getScriptProperties().getProperty('PAINEL_CHAVE');
  if (!certa || chave !== certa) throw new Error('Acesso ao painel negado.');
  return listar_();
}

/** Resposta de um militar, para ele revisar e corrigir no formulário. */
function carregar(id) {
  var r = listar_().filter(function (x) { return x.id === id; })[0];
  return r || null;
}

function listar_() {
  var sh = aba_(ABA, CAB);
  var n = sh.getLastRow();
  if (n < 2) return [];
  var vals = sh.getRange(2, 1, n - 1, CAB.length).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    var json = vals[i][7];
    if (!json) continue;
    try { out.push(JSON.parse(json)); } catch (e) { /* linha corrompida: ignora */ }
  }
  return out;
}

/** Grava a resposta de UM militar. revEsperada = versão que o aparelho carregou (0 = primeira resposta). */
function salvar(registroJson, revEsperada) {
  var rec = JSON.parse(registroJson);
  if (!rec || !MILITARES[rec.id]) throw new Error('Militar inválido.');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = aba_(ABA, CAB);
    var n = sh.getLastRow();
    var linha = -1, atual = null;
    if (n >= 2) {
      var ids = sh.getRange(2, 1, n - 1, 1).getValues();
      for (var i = 0; i < ids.length; i++) if (ids[i][0] === rec.id) { linha = i + 2; break; }
    }
    if (linha > 0) { try { atual = JSON.parse(sh.getRange(linha, 8).getValue()); } catch (e) { atual = null; } }
    var revAtual = atual ? (atual.rev || 0) : 0;
    if (revAtual !== Number(revEsperada || 0)) return { ok: false, conflict: true, current: atual };

    var agora = new Date().toISOString();
    rec.militar = MILITARES[rec.id];
    rec.rev = revAtual + 1;
    rec.atualizadoEm = agora;
    rec.enviadoEm = (atual && atual.enviadoEm) || agora;
    var row = [rec.id, rec.militar, rec.rev, agora, rec.dispensa || '', rec.temDeslig ? (rec.deslig || '') : 'não', leitura_(rec), JSON.stringify(rec)];
    if (linha > 0) sh.getRange(linha, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);
    aba_(ABA_HIST, ['quando', 'id', 'militar', 'rev', 'dados']).appendRow([agora, rec.id, rec.militar, rec.rev, JSON.stringify(rec)]);
    SpreadsheetApp.flush();
    return { ok: true, registro: rec };
  } finally {
    lock.releaseLock();
  }
}

function leitura_(rec) {
  var br = function (s) { return s ? s.slice(8, 10) + '/' + s.slice(5, 7) + '/' + s.slice(0, 4) : '?'; };
  return (rec.periodos || []).map(function (p) {
    return (NOMES_CAT[p.cat] || '?') + ': ' + br(p.ini) + ' a ' + br(p.fim) + (p.obs ? ' (' + p.obs + ')' : '');
  }).join('\n');
}

/** Rode UMA vez pelo editor (botão Executar) para criar as abas e lançar os dados do Maj Modesto. */
function instalar() {
  aba_(ABA, CAB);
  aba_(ABA_HIST, ['quando', 'id', 'militar', 'rev', 'dados']);
  var props = PropertiesService.getScriptProperties();
  if (!props.getProperty('PAINEL_CHAVE')) props.setProperty('PAINEL_CHAVE', Utilities.getUuid().replace(/-/g, '').slice(0, 16));
  Logger.log('Link do painel: <URL da implantação>?v=painel&chave=' + props.getProperty('PAINEL_CHAVE'));
  var existe = listar_().some(function (r) { return r.id === 'modesto'; });
  if (!existe) {
    salvar(JSON.stringify({
      id: 'modesto', militar: 'Maj Modesto', dispensa: 'sem', temDeslig: true, deslig: '2027-01-24',
      periodos: [
        { id: 'p-dispcmt', cat: 'dispcmt', ini: '2026-12-07', fim: '2026-12-21' },
        { id: 'p-ferias', cat: 'ferias26', ini: '2026-12-23', fim: '2027-01-21' }
      ]
    }), 0);
  }
}
