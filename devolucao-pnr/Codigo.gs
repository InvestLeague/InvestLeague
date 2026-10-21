// ARQUIVO GERADO por gerar-apps-script.js a partir de Code.gs + index.html. Cole inteiro no Apps Script.
/**
 * Agendamento de devolução de PNR — EsAO
 * Backend em Google Apps Script, ligado a uma Planilha Google.
 *
 * Três telas no mesmo link publicado:
 *   .../exec                            → aluno (aberto, sem login)
 *   .../exec?v=instrutor&chave=<CHAVE>  → instrutor (uma chave por instrutor)
 *   .../exec?v=painel&chave=<CHAVE>     → coordenação
 *
 * Abas criadas sozinhas no primeiro acesso (função instalar()):
 *   Instrutores   — nome, e-mail, disponibilidade, total de alunos, chave pessoal
 *   Agendamentos  — uma linha por agendamento, com histórico no JSON
 *   Bloqueios     — dias e tempos bloqueados por instrutor
 *   Aceites       — registro de quem leu e aceitou os termos (prova)
 *   Log           — auditoria de tudo que foi feito
 *   Config        — horários, período e regra das 72 h
 *
 * Gravação sob LockService e com número de versão (rev): dois alunos nunca
 * conseguem pegar o mesmo tempo, e um pedido alterado em dois aparelhos é
 * recusado no segundo em vez de sobrescrever o primeiro.
 */

var TZ = 'America/Sao_Paulo';
var OFFSET = '-03:00';            // o Rio não tem horário de verão desde 2019
var TERMO_VERSAO = '1.0';

var A_INST = 'Instrutores';
var A_AGEND = 'Agendamentos';
var A_BLOQ = 'Bloqueios';
var A_ACEITE = 'Aceites';
var A_LOG = 'Log';
var A_CONFIG = 'Config';

var CAB_INST = ['nome', 'e-mail (para lembretes)', 'disponibilidade', 'total de alunos',
  'chave do instrutor (só para ele)', 'chave do link dos orientados', 'ativo'];
var CAB_AGEND = ['id', 'nome de guerra', 'instrutor', 'data', 'tempo', 'hora', 'situação', 'atualizado em', 'resumo (leitura)', 'dados (JSON - não editar)'];
var CAB_BLOQ = ['instrutor (ou TODOS)', 'data', 'tempo (ou TODOS)', 'motivo', 'criado em'];
var CAB_ACEITE = ['quando', 'nome de guerra', 'instrutor', 'versão do termo', 'itens aceitos'];
var CAB_LOG = ['quando', 'quem', 'ação', 'detalhe'];

/** Os 6 tempos do dia: vistoria de 1 h 15, com 15 min de intervalo. */
var TEMPOS = [
  { k: 'M1', ini: '08:00', fim: '09:15', turno: 'manha' },
  { k: 'M2', ini: '09:30', fim: '10:45', turno: 'manha' },
  { k: 'M3', ini: '11:00', fim: '12:15', turno: 'manha' },
  { k: 'T1', ini: '13:30', fim: '14:45', turno: 'tarde' },
  { k: 'T2', ini: '15:00', fim: '16:15', turno: 'tarde' },
  { k: 'T3', ini: '16:30', fim: '17:45', turno: 'tarde' }
];

var INI_PADRAO = '2026-10-21';
var FIM_PADRAO = '2026-12-30';
var HORAS_MIN = 72;

/** Feriados e pontos facultativos. 'dia' bloqueia tudo; 'tarde' bloqueia só os 3 tempos da tarde. */
var FERIADOS = {
  '2026-11-02': { tipo: 'dia', n: 'Finados' },
  '2026-11-20': { tipo: 'dia', n: 'Dia Nacional de Zumbi e da Consciência Negra' },
  '2026-12-24': { tipo: 'tarde', n: 'Véspera de Natal (ponto facultativo à tarde)' },
  '2026-12-25': { tipo: 'dia', n: 'Natal' }
};

/** Escala de recebimento 2026. A disponibilidade pode ser editada na aba Instrutores. */
var SEED_INSTRUTORES = [
  ['Maj Modesto', 10, 'Até 04/12'],
  ['Maj Andrews', 10, 'Até 18/12'],
  ['Maj Pimenta', 1, '09/11 a 13/11'],
  ['Maj Calderaro', 10, 'Até 19/11'],
  ['Maj Cesse', 11, 'Até 04/12'],
  ['Cap Dantas', 10, '13/11 a 26/11'],
  ['Cap Loan', 11, 'Até 04/12'],
  ['Cap Santana', 11, 'Até 11/12'],
  ['Cap Uerlei Moreira', 11, 'Até 18/12'],
  ['Cap Cavalcanti', 11, 'Até 29/12'],
  ['Cap Marcello', 11, 'Até 19/11; 30/11 a 22/12']
];

/* ===================== utilidades ===================== */

function doGet() {
  garantirInstalacao_();
  return HtmlService.createHtmlOutput(PAGINA)
    .setTitle('Devolução de PNR · EsAO')
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
    sh.setColumnWidth(1, 160);
  }
  return sh;
}

function hoje_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd'); }
function agora_() { return new Date().toISOString(); }

/** Data 'yyyy-MM-dd' + hora 'HH:mm' no fuso do Rio → objeto Date. */
function quando_(data, hora) { return new Date(data + 'T' + hora + ':00' + OFFSET); }

function tempo_(k) {
  for (var i = 0; i < TEMPOS.length; i++) if (TEMPOS[i].k === k) return TEMPOS[i];
  return null;
}

function br_(d) { return d ? d.slice(8, 10) + '/' + d.slice(5, 7) + '/' + d.slice(0, 4) : ''; }

/** Compara nomes de guerra ignorando acento, caixa e espaço repetido. */
function norm_(s) {
  return String(s || '').trim().toUpperCase()
    .replace(/[ÁÀÂÃÄ]/g, 'A').replace(/[ÉÈÊË]/g, 'E').replace(/[ÍÌÎÏ]/g, 'I')
    .replace(/[ÓÒÔÕÖ]/g, 'O').replace(/[ÚÙÛÜ]/g, 'U').replace(/Ç/g, 'C')
    .replace(/\s+/g, ' ');
}

function log_(quem, acao, detalhe) {
  try { aba_(A_LOG, CAB_LOG).appendRow([agora_(), quem || '?', acao, detalhe || '']); } catch (e) {}
}

function config_() {
  var sh = aba_(A_CONFIG, ['chave', 'valor', 'o que é']);
  var n = sh.getLastRow(), out = {};
  if (n >= 2) {
    var v = sh.getRange(2, 1, n - 1, 2).getValues();
    for (var i = 0; i < v.length; i++) if (v[i][0]) out[String(v[i][0])] = v[i][1];
  }
  var ini = out.inicio ? String(out.inicio).slice(0, 10) : INI_PADRAO;
  var fim = out.fim ? String(out.fim).slice(0, 10) : FIM_PADRAO;
  var horas = Number(out.horasMinimas);
  return {
    inicio: ini > hoje_() ? ini : hoje_(),
    fim: fim,
    horasMinimas: isNaN(horas) || horas < 0 ? HORAS_MIN : horas
  };
}

/* ===================== instrutores ===================== */

/**
 * Lê a disponibilidade escrita em português e devolve períodos fechados.
 * Entende "Até 04/12", "09/11 a 13/11", "03/12 e 04/12" e combinações com ';'.
 * Mês 10 ou maior vira 2026; menor vira 2027.
 */
function janelas_(texto, cfg) {
  var t = String(texto || '').trim();
  if (!t) return [{ ini: cfg.inicio, fim: cfg.fim }];
  var data = function (s) {
    var m = String(s).match(/(\d{1,2})\s*\/\s*(\d{1,2})(?:\s*\/\s*(\d{2,4}))?/);
    if (!m) return null;
    var d = ('0' + m[1]).slice(-2), mes = ('0' + m[2]).slice(-2);
    var ano = m[3] ? (m[3].length === 2 ? '20' + m[3] : m[3]) : (Number(mes) >= 10 ? '2026' : '2027');
    return ano + '-' + mes + '-' + d;
  };
  var out = [];
  t.split(/;|\be\b(?![^\d]*a\b)/i).forEach(function (parte) {
    parte = parte.trim();
    if (!parte) return;
    var faixa = parte.match(/(\d{1,2}\s*\/\s*\d{1,2}(?:\s*\/\s*\d{2,4})?)\s*(?:a|até|a\s+até|-|–)\s*(\d{1,2}\s*\/\s*\d{1,2}(?:\s*\/\s*\d{2,4})?)/i);
    if (faixa) { out.push({ ini: data(faixa[1]), fim: data(faixa[2]) }); return; }
    if (/at[eé]/i.test(parte)) { var f = data(parte); if (f) out.push({ ini: cfg.inicio, fim: f }); return; }
    var u = data(parte);
    if (u) out.push({ ini: u, fim: u });
  });
  out = out.filter(function (r) { return r.ini && r.fim && r.ini <= r.fim; });
  if (!out.length) return [{ ini: cfg.inicio, fim: cfg.fim }];
  // Recorta pelo período da campanha.
  return out.map(function (r) {
    return { ini: r.ini < cfg.inicio ? cfg.inicio : r.ini, fim: r.fim > cfg.fim ? cfg.fim : r.fim };
  }).filter(function (r) { return r.ini <= r.fim; });
}

function instrutores_() {
  var sh = aba_(A_INST, CAB_INST);
  var n = sh.getLastRow();
  if (n < 2) return [];
  var vals = sh.getRange(2, 1, n - 1, CAB_INST.length).getValues();
  var cfg = config_(), out = [];
  for (var i = 0; i < vals.length; i++) {
    var nome = String(vals[i][0] || '').trim();
    if (!nome) continue;
    if (String(vals[i][6]).toLowerCase() === 'não' || String(vals[i][6]).toLowerCase() === 'nao') continue;
    out.push({
      nome: nome,
      email: String(vals[i][1] || '').trim(),
      dispTexto: String(vals[i][2] || '').trim(),
      janelas: janelas_(vals[i][2], cfg),
      total: Number(vals[i][3]) || 0,
      chave: String(vals[i][4] || '').trim(),
      chaveAlunos: String(vals[i][5] || '').trim(),
      linha: i + 2
    });
  }
  return out;
}

/** Descobre a quem pertence a chave: um instrutor, a coordenação, ou ninguém. */
function papel_(chave) {
  chave = String(chave || '').trim();
  if (!chave) return { papel: 'aluno' };
  var painel = PropertiesService.getScriptProperties().getProperty('PAINEL_CHAVE');
  if (painel && chave === painel) return { papel: 'painel' };
  var lista = instrutores_();
  for (var i = 0; i < lista.length; i++) if (lista[i].chave && lista[i].chave === chave) return { papel: 'instrutor', instrutor: lista[i].nome };
  return { papel: 'aluno', chaveInvalida: true };
}

/**
 * Link que o instrutor divulga aos seus orientados: amarra o aluno àquele instrutor.
 * É uma chave diferente da do instrutor, e não dá nenhum poder de instrutor.
 */
function instrutorPorTokenAlunos_(tok) {
  tok = String(tok || '').trim();
  if (!tok) return null;
  var lista = instrutores_();
  for (var i = 0; i < lista.length; i++) if (lista[i].chaveAlunos && lista[i].chaveAlunos === tok) return lista[i];
  return null;
}

/* ===================== agendamentos ===================== */

function agendamentos_() {
  var sh = aba_(A_AGEND, CAB_AGEND);
  var n = sh.getLastRow();
  if (n < 2) return [];
  var vals = sh.getRange(2, 1, n - 1, CAB_AGEND.length).getValues();
  var out = [];
  for (var i = 0; i < vals.length; i++) {
    if (!vals[i][9]) continue;
    try { var r = JSON.parse(vals[i][9]); r._linha = i + 2; out.push(r); } catch (e) {}
  }
  return out;
}

var VIVOS = { solicitado: 1, encaixe: 1, confirmado: 1, pendencias: 1 };
function vivo_(a) { return !!VIVOS[a.situacao]; }
/** Ocupa o tempo na agenda do instrutor (pedido pendente também segura o horário). */
function ocupa_(a) { return a.situacao === 'solicitado' || a.situacao === 'encaixe' || a.situacao === 'confirmado'; }

function resumo_(a) {
  var t = tempo_(a.tempo);
  var linhas = [
    br_(a.data) + ' ' + (t ? t.ini + ' às ' + t.fim : '') + ' · ' + a.instrutor,
    'Situação: ' + rotulo_(a.situacao)
  ];
  if (a.endereco) linhas.push('PNR: ' + a.endereco);
  if (a.viagem) linhas.push('Viagem: ' + br_(a.viagem));
  if (a.novaOM) linhas.push('Nova OM: ' + a.novaOM);
  if (a.pendencias && a.pendencias.length) linhas.push('Pendências: ' + a.pendencias.join('; '));
  return linhas.join('\n');
}

function rotulo_(s) {
  return ({
    solicitado: 'Aguardando confirmação do instrutor',
    encaixe: 'Pedido de encaixe (menos de 72 h)',
    confirmado: 'Confirmada',
    recusado: 'Recusada pelo instrutor',
    cancelado: 'Cancelada pelo aluno',
    entregue: 'Chaves recebidas',
    pendencias: 'Com pendências a corrigir'
  })[s] || s;
}

function gravar_(a) {
  var sh = aba_(A_AGEND, CAB_AGEND);
  var t = tempo_(a.tempo);
  var row = [a.id, a.nomeGuerra, a.instrutor, a.data, a.tempo, t ? t.ini : '', rotulo_(a.situacao),
    a.atualizadoEm, resumo_(a), JSON.stringify(a)];
  if (a._linha) sh.getRange(a._linha, 1, 1, row.length).setValues([row]);
  else {
    sh.appendRow(row);
    a._linha = sh.getLastRow();
  }
  SpreadsheetApp.flush();
}

function hist_(a, texto, quem) {
  a.historico = a.historico || [];
  a.historico.push({ quando: agora_(), quem: quem || '', texto: texto });
}

/* ===================== bloqueios ===================== */

function bloqueios_() {
  var sh = aba_(A_BLOQ, CAB_BLOQ);
  var n = sh.getLastRow();
  if (n < 2) return [];
  var vals = sh.getRange(2, 1, n - 1, 4).getValues(), out = [];
  for (var i = 0; i < vals.length; i++) {
    var d = vals[i][1];
    if (!d) continue;
    if (d instanceof Date) d = Utilities.formatDate(d, TZ, 'yyyy-MM-dd');
    out.push({
      instrutor: String(vals[i][0] || 'TODOS').trim(),
      data: String(d).slice(0, 10),
      tempo: String(vals[i][2] || 'TODOS').trim(),
      motivo: String(vals[i][3] || '').trim(),
      linha: i + 2
    });
  }
  return out;
}

/* ===================== API chamada pela página ===================== */

/**
 * Tudo o que a página precisa numa só chamada.
 * O aluno recebe apenas os tempos ocupados (sem nome de colega) e o próprio pedido.
 */
function estado(chave, nomeGuerra, tokAlunos) {
  garantirInstalacao_();
  var p = papel_(chave);
  var cfg = config_();
  var insts = instrutores_();
  var ags = agendamentos_();
  var bls = bloqueios_();

  // Aluno que veio pelo link do instrutor: fica amarrado a ele, sem escolher.
  var fixo = (p.papel === 'aluno' && tokAlunos) ? instrutorPorTokenAlunos_(tokAlunos) : null;

  var pub = {
    papel: p.papel,
    chaveInvalida: !!p.chaveInvalida,
    instrutorFixo: fixo ? fixo.nome : '',
    tokenInvalido: !!(p.papel === 'aluno' && tokAlunos && !fixo),
    eu: p.instrutor || '',
    hoje: hoje_(),
    agora: agora_(),
    // A página roda num iframe, então o endereço publicado tem de vir do servidor.
    urlBase: (function () { try { return ScriptApp.getService().getUrl() || ''; } catch (e) { return ''; } })(),
    termoVersao: TERMO_VERSAO,
    cfg: cfg,
    tempos: TEMPOS,
    feriados: FERIADOS,
    instrutores: insts.map(function (i) {
      var o = { nome: i.nome, dispTexto: i.dispTexto, janelas: i.janelas, total: i.total, temEmail: !!i.email };
      // A chave do link dos orientados só vai para a coordenação e para o próprio instrutor.
      if (p.papel === 'painel' || (p.papel === 'instrutor' && i.nome === p.instrutor)) o.tokenAlunos = i.chaveAlunos;
      return o;
    }),
    bloqueios: bls.map(function (b) { return { instrutor: b.instrutor, data: b.data, tempo: b.tempo, motivo: b.motivo }; })
  };

  // Ocupação: o aluno vê só que está ocupado; instrutor e painel veem quem é.
  pub.ocupados = ags.filter(ocupa_).map(function (a) {
    var o = { instrutor: a.instrutor, data: a.data, tempo: a.tempo };
    if (p.papel === 'painel' || (p.papel === 'instrutor' && a.instrutor === p.instrutor)) {
      o.nomeGuerra = a.nomeGuerra; o.situacao = a.situacao; o.id = a.id;
    }
    return o;
  });

  if (p.papel === 'painel') {
    pub.agendamentos = ags.map(limpo_);
  } else if (p.papel === 'instrutor') {
    pub.agendamentos = ags.filter(function (a) { return a.instrutor === p.instrutor; }).map(limpo_);
    pub.excecoes = excecoes_(p.instrutor);
  } else if (nomeGuerra) {
    var k = norm_(nomeGuerra);
    pub.meus = ags.filter(function (a) { return norm_(a.nomeGuerra) === k; }).map(limpo_);
    pub.excecoesMinhas = excecoesDe_(k);
  }
  return pub;
}

function limpo_(a) { var o = {}; for (var k in a) if (k !== '_linha') o[k] = a[k]; return o; }

/* ---- exceções à regra das 72 h (liberadas pelo instrutor) ---- */

function excecoes_(instrutor) {
  var raw = PropertiesService.getScriptProperties().getProperty('EXCECOES') || '{}';
  var all = {};
  try { all = JSON.parse(raw); } catch (e) {}
  var out = [], hoje = hoje_();
  for (var k in all) if (all[k].instrutor === instrutor && all[k].validaAte >= hoje) out.push(all[k]);
  return out;
}

function excecoesDe_(chaveNome) {
  var raw = PropertiesService.getScriptProperties().getProperty('EXCECOES') || '{}';
  var all = {};
  try { all = JSON.parse(raw); } catch (e) {}
  var hoje = hoje_();
  return all[chaveNome] && all[chaveNome].validaAte >= hoje ? [all[chaveNome]] : [];
}

/** O instrutor dispensa um aluno da regra das 72 h por alguns dias. */
function liberarExcecao(chave, nomeGuerra, dias) {
  var p = papel_(chave);
  if (p.papel !== 'instrutor') throw new Error('Só o instrutor pode liberar exceção.');
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var props = PropertiesService.getScriptProperties();
    var all = {};
    try { all = JSON.parse(props.getProperty('EXCECOES') || '{}'); } catch (e) {}
    var d = Math.max(1, Math.min(30, Number(dias) || 7));
    var ate = Utilities.formatDate(new Date(Date.now() + d * 864e5), TZ, 'yyyy-MM-dd');
    all[norm_(nomeGuerra)] = { nomeGuerra: String(nomeGuerra).trim(), instrutor: p.instrutor, validaAte: ate };
    props.setProperty('EXCECOES', JSON.stringify(all));
    log_(p.instrutor, 'liberou exceção de 72 h', nomeGuerra + ' até ' + br_(ate));
    return { ok: true, validaAte: ate };
  } finally { lock.releaseLock(); }
}

/* ---- aceite dos termos ---- */

function aceitarTermos(nomeGuerra, instrutor, itens) {
  var nome = String(nomeGuerra || '').trim();
  if (nome.length < 3) throw new Error('Informe o nome de guerra.');
  aba_(A_ACEITE, CAB_ACEITE).appendRow([agora_(), nome, instrutor || '', TERMO_VERSAO, (itens || []).join(' | ')]);
  log_(nome, 'aceitou os termos', 'versão ' + TERMO_VERSAO);
  return { ok: true, versao: TERMO_VERSAO };
}

/* ---- validação de um tempo ---- */

/**
 * Diz se o tempo pode ser marcado e, quando não pode, explica o motivo em
 * português. A mesma função serve para o aluno e para o instrutor.
 * `porInstrutor` dispensa a regra das 72 h.
 */
function checar_(dados, insts, ags, bls, cfg, porInstrutor, excecao) {
  var data = String(dados.data || '').slice(0, 10);
  var t = tempo_(dados.tempo);
  if (!t) return 'Tempo inválido.';
  if (!/^\d{4}-\d{2}-\d{2}$/.test(data)) return 'Data inválida.';
  if (data < cfg.inicio) return 'Essa data já passou.';
  if (data > cfg.fim) return 'A campanha de devolução termina em ' + br_(cfg.fim) + '.';

  var dow = new Date(data + 'T12:00:00' + OFFSET).getDay();
  if (dow === 0 || dow === 6) return 'Não há vistoria em fins de semana.';

  var fer = FERIADOS[data];
  if (fer && fer.tipo === 'dia') return fer.n + ': feriado, sem vistoria.';
  if (fer && fer.tipo === 'tarde' && t.turno === 'tarde') return fer.n + ': só há vistoria pela manhã.';

  var inst = null;
  for (var i = 0; i < insts.length; i++) if (insts[i].nome === dados.instrutor) inst = insts[i];
  if (!inst) return 'Instrutor não encontrado. Escolha um da lista.';

  var dentro = inst.janelas.some(function (j) { return data >= j.ini && data <= j.fim; });
  if (!dentro) return 'O ' + inst.nome + ' só recebe vistorias em: ' + inst.dispTexto + '.';

  for (var b = 0; b < bls.length; b++) {
    var bl = bls[b];
    if (bl.data !== data) continue;
    if (bl.instrutor !== 'TODOS' && bl.instrutor !== dados.instrutor) continue;
    if (bl.tempo !== 'TODOS' && bl.tempo !== dados.tempo) continue;
    return 'O ' + inst.nome + ' bloqueou esse horário' + (bl.motivo ? ': ' + bl.motivo : '.');
  }

  for (var a = 0; a < ags.length; a++) {
    if (ocupa_(ags[a]) && ags[a].instrutor === dados.instrutor && ags[a].data === data && ags[a].tempo === dados.tempo
      && ags[a].id !== dados.id) return 'Esse horário acabou de ser ocupado. Escolha outro.';
  }

  if (dados.viagem) {
    if (data > dados.viagem) return 'Essa data é depois da sua viagem (' + br_(dados.viagem) + ').';
    if (data === dados.viagem) return 'É o dia da sua viagem. Se houver pendência, a chave não é recebida. Escolha 2 ou 3 dias antes.';
  }

  if (!porInstrutor && !excecao) {
    var horas = (quando_(data, t.ini).getTime() - Date.now()) / 36e5;
    if (horas < cfg.horasMinimas) return 'Faltam menos de ' + cfg.horasMinimas + ' h. Use o pedido de encaixe para falar com o instrutor.';
  }
  return null;
}

/* ---- aluno solicita ---- */

function solicitar(dados) {
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var cfg = config_(), insts = instrutores_(), ags = agendamentos_(), bls = bloqueios_();
    var nome = String(dados.nomeGuerra || '').trim();
    if (nome.length < 3) return { ok: false, erro: 'Informe o seu nome de guerra.' };

    // Veio pelo link de um instrutor: o pedido é dele, e o aluno não pode desviar para outro.
    if (dados.tokAlunos) {
      var dono = instrutorPorTokenAlunos_(dados.tokAlunos);
      if (!dono) return { ok: false, erro: 'Esse link não vale mais. Peça o link atual ao seu instrutor.' };
      dados.instrutor = dono.nome;
    }
    if (!dados.endereco || String(dados.endereco).trim().length < 8) return { ok: false, erro: 'Informe o endereço do PNR (rua, bloco e apartamento).' };
    if (!dados.viagem) return { ok: false, erro: 'Informe a data da sua viagem.' };
    if (!dados.aceitou) return { ok: false, erro: 'É preciso ler e aceitar os termos.' };

    var k = norm_(nome);
    var ativo = ags.filter(function (a) { return norm_(a.nomeGuerra) === k && vivo_(a); })[0];
    if (ativo) return { ok: false, erro: 'Você já tem uma vistoria em andamento em ' + br_(ativo.data) + '. Cancele antes de marcar outra.', atual: limpo_(ativo) };

    var encaixe = !!dados.encaixe;
    var exc = excecoesDe_(k).length > 0;
    var erro = checar_(dados, insts, ags, bls, cfg, false, exc || encaixe);
    if (erro) return { ok: false, erro: erro };

    if (encaixe) {
      // No encaixe, só a regra das 72 h é dispensada; todo o resto já foi checado.
      var horas = (quando_(dados.data, tempo_(dados.tempo).ini).getTime() - Date.now()) / 36e5;
      if (horas >= cfg.horasMinimas && !exc) encaixe = false;
      if (horas < 2) return { ok: false, erro: 'Esse horário é daqui a menos de 2 h. Fale direto com o instrutor.' };
    }

    var a = {
      id: Utilities.getUuid().slice(0, 8),
      nomeGuerra: nome,
      instrutor: dados.instrutor,
      data: String(dados.data).slice(0, 10),
      tempo: dados.tempo,
      situacao: encaixe ? 'encaixe' : 'solicitado',
      endereco: String(dados.endereco).trim(),
      viagem: String(dados.viagem).slice(0, 10),
      novaOM: String(dados.novaOM || '').trim(),
      telefone: String(dados.telefone || '').trim(),
      checklist: dados.checklist || {},
      justificativa: String(dados.justificativa || '').trim(),
      termoVersao: TERMO_VERSAO,
      rev: 1,
      criadoEm: agora_(),
      atualizadoEm: agora_(),
      pendencias: [],
      historico: []
    };
    hist_(a, encaixe ? 'Pedido de encaixe enviado pelo aluno.' : 'Vistoria solicitada pelo aluno.', nome);
    gravar_(a);
    log_(nome, encaixe ? 'pediu encaixe' : 'solicitou vistoria', a.instrutor + ' ' + br_(a.data) + ' ' + a.tempo);
    avisarInstrutor_(a, encaixe ? 'encaixe' : 'novo');
    return { ok: true, agendamento: limpo_(a) };
  } finally { lock.releaseLock(); }
}

/* ---- aluno cancela ---- */

function cancelar(id, nomeGuerra, motivo) {
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var a = acharId_(id);
    if (!a) return { ok: false, erro: 'Agendamento não encontrado.' };
    if (norm_(a.nomeGuerra) !== norm_(nomeGuerra)) return { ok: false, erro: 'Para cancelar, digite o nome de guerra exatamente como no agendamento.' };
    if (!vivo_(a)) return { ok: false, erro: 'Esse agendamento já está encerrado.' };
    a.situacao = 'cancelado';
    a.motivo = String(motivo || '').trim();
    a.atualizadoEm = agora_(); a.rev = (a.rev || 0) + 1;
    hist_(a, 'Cancelado pelo aluno.' + (a.motivo ? ' Motivo: ' + a.motivo : ''), a.nomeGuerra);
    gravar_(a);
    apagarEvento_(a);
    log_(a.nomeGuerra, 'cancelou', a.instrutor + ' ' + br_(a.data));
    avisarInstrutor_(a, 'cancelado');
    return { ok: true, agendamento: limpo_(a) };
  } finally { lock.releaseLock(); }
}

function acharId_(id) {
  var ags = agendamentos_();
  for (var i = 0; i < ags.length; i++) if (ags[i].id === id) return ags[i];
  return null;
}

/* ---- ações do instrutor ---- */

function confirmar(chave, id) {
  return acaoInstrutor_(chave, id, function (a, eu) {
    if (a.situacao !== 'solicitado' && a.situacao !== 'encaixe') throw new Error('Esse pedido não está aguardando confirmação.');
    a.situacao = 'confirmado';
    hist_(a, 'Confirmada pelo instrutor.', eu);
    criarEvento_(a);
    avisarAluno_(a, 'confirmado');
    return 'confirmou';
  });
}

function recusar(chave, id, motivo, sugestao) {
  return acaoInstrutor_(chave, id, function (a, eu) {
    a.situacao = 'recusado';
    a.motivo = String(motivo || '').trim();
    a.sugestao = String(sugestao || '').trim();
    hist_(a, 'Recusada pelo instrutor.' + (a.motivo ? ' Motivo: ' + a.motivo : '') + (a.sugestao ? ' Sugestão: ' + a.sugestao : ''), eu);
    apagarEvento_(a);
    avisarAluno_(a, 'recusado');
    return 'recusou';
  });
}

/** O aluno escolheu o instrutor errado: o pedido passa para o certo, sem se perder. */
function transferir(chave, id, novoInstrutor) {
  return acaoInstrutor_(chave, id, function (a, eu) {
    var achou = instrutores_().some(function (i) { return i.nome === novoInstrutor; });
    if (!achou) throw new Error('Instrutor de destino não encontrado.');
    var antes = a.instrutor;
    a.instrutor = novoInstrutor;
    a.situacao = 'solicitado';
    hist_(a, 'Transferida de ' + antes + ' para ' + novoInstrutor + '.', eu);
    apagarEvento_(a);
    avisarInstrutor_(a, 'transferido');
    avisarAluno_(a, 'transferido');
    return 'transferiu para ' + novoInstrutor;
  });
}

/** O instrutor marca pelo aluno (quem ligou, quem esqueceu): já sai confirmado. */
function agendarPeloInstrutor(chave, dados) {
  var p = papel_(chave);
  if (p.papel !== 'instrutor') throw new Error('Acesso negado.');
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var cfg = config_(), insts = instrutores_(), ags = agendamentos_(), bls = bloqueios_();
    var nome = String(dados.nomeGuerra || '').trim();
    if (nome.length < 3) return { ok: false, erro: 'Informe o nome de guerra do aluno.' };
    dados.instrutor = p.instrutor;
    var k = norm_(nome);
    var ativo = ags.filter(function (a) { return norm_(a.nomeGuerra) === k && vivo_(a); })[0];
    if (ativo) return { ok: false, erro: nome + ' já tem vistoria em ' + br_(ativo.data) + ' com o ' + ativo.instrutor + '.' };
    var erro = checar_(dados, insts, ags, bls, cfg, true, true);
    if (erro) return { ok: false, erro: erro };
    var a = {
      id: Utilities.getUuid().slice(0, 8), nomeGuerra: nome, instrutor: p.instrutor,
      data: String(dados.data).slice(0, 10), tempo: dados.tempo, situacao: 'confirmado',
      endereco: String(dados.endereco || '').trim(), viagem: String(dados.viagem || '').slice(0, 10),
      novaOM: '', telefone: '', checklist: {}, termoVersao: '', pendencias: [],
      criadoEm: agora_(), atualizadoEm: agora_(), rev: 1, historico: [], peloInstrutor: true
    };
    hist_(a, 'Marcada pelo próprio instrutor.', p.instrutor);
    gravar_(a);
    log_(p.instrutor, 'agendou pelo aluno', nome + ' ' + br_(a.data) + ' ' + a.tempo);
    criarEvento_(a);
    return { ok: true, agendamento: limpo_(a) };
  } finally { lock.releaseLock(); }
}

/** Resultado da vistoria: chaves recebidas, ou lista do que falta corrigir. */
function registrarResultado(chave, id, recebeu, pendencias, obs) {
  return acaoInstrutor_(chave, id, function (a, eu) {
    if (recebeu) {
      a.situacao = 'entregue';
      a.pendencias = [];
      hist_(a, 'Vistoria realizada e chaves recebidas.' + (obs ? ' ' + obs : ''), eu);
      avisarAluno_(a, 'entregue');
    } else {
      a.situacao = 'pendencias';
      a.pendencias = (pendencias || []).map(function (s) { return String(s).trim(); }).filter(String);
      hist_(a, 'Vistoria com pendências: ' + a.pendencias.join('; ') + (obs ? '. ' + obs : ''), eu);
      avisarAluno_(a, 'pendencias');
    }
    a.obs = String(obs || '').trim();
    apagarEvento_(a);
    return recebeu ? 'recebeu as chaves' : 'registrou pendências';
  });
}

/** Depois de corrigir as pendências, o aluno marca a nova vistoria. */
function remarcar(id, nomeGuerra, data, tempo) {
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var a = acharId_(id);
    if (!a) return { ok: false, erro: 'Agendamento não encontrado.' };
    if (norm_(a.nomeGuerra) !== norm_(nomeGuerra)) return { ok: false, erro: 'Digite o nome de guerra exatamente como no agendamento.' };
    if (a.situacao !== 'pendencias' && a.situacao !== 'recusado') return { ok: false, erro: 'Esse agendamento não está aguardando nova data.' };
    var cfg = config_(), insts = instrutores_(), ags = agendamentos_(), bls = bloqueios_();
    var alvo = { id: a.id, instrutor: a.instrutor, data: data, tempo: tempo, viagem: a.viagem };
    var erro = checar_(alvo, insts, ags, bls, cfg, false, excecoesDe_(norm_(nomeGuerra)).length > 0);
    if (erro) return { ok: false, erro: erro };
    var antes = a.data + ' ' + a.tempo;
    a.data = String(data).slice(0, 10); a.tempo = tempo;
    a.situacao = 'solicitado';
    a.atualizadoEm = agora_(); a.rev = (a.rev || 0) + 1;
    hist_(a, 'Aluno pediu nova vistoria (antes ' + antes + ').', a.nomeGuerra);
    gravar_(a);
    avisarInstrutor_(a, 'remarcado');
    return { ok: true, agendamento: limpo_(a) };
  } finally { lock.releaseLock(); }
}

function acaoInstrutor_(chave, id, fn) {
  var p = papel_(chave);
  if (p.papel !== 'instrutor' && p.papel !== 'painel') throw new Error('Acesso negado.');
  var lock = LockService.getScriptLock(); lock.waitLock(25000);
  try {
    var a = acharId_(id);
    if (!a) return { ok: false, erro: 'Agendamento não encontrado.' };
    if (p.papel === 'instrutor' && a.instrutor !== p.instrutor) return { ok: false, erro: 'Esse pedido é de outro instrutor.' };
    var eu = p.instrutor || 'Coordenação';
    var acao;
    try { acao = fn(a, eu); } catch (e) { return { ok: false, erro: e.message }; }
    a.atualizadoEm = agora_(); a.rev = (a.rev || 0) + 1;
    gravar_(a);
    log_(eu, acao, a.nomeGuerra + ' ' + br_(a.data) + ' ' + a.tempo);
    return { ok: true, agendamento: limpo_(a) };
  } finally { lock.releaseLock(); }
}

/* ---- bloqueios do instrutor ---- */

function bloquear(chave, data, tempo, motivo) {
  var p = papel_(chave);
  if (p.papel !== 'instrutor' && p.papel !== 'painel') throw new Error('Acesso negado.');
  var quem = p.instrutor || 'TODOS';
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var ags = agendamentos_().filter(function (a) {
      return ocupa_(a) && a.instrutor === quem && a.data === data && (tempo === 'TODOS' || a.tempo === tempo);
    });
    if (ags.length) return { ok: false, erro: 'Há ' + ags.length + ' vistoria(s) marcada(s) nesse dia. Recuse ou remarque antes de bloquear.' };
    aba_(A_BLOQ, CAB_BLOQ).appendRow([quem, data, tempo || 'TODOS', String(motivo || '').trim(), agora_()]);
    log_(quem, 'bloqueou horário', data + ' ' + (tempo || 'TODOS'));
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function desbloquear(chave, data, tempo) {
  var p = papel_(chave);
  if (p.papel !== 'instrutor' && p.papel !== 'painel') throw new Error('Acesso negado.');
  var quem = p.instrutor || 'TODOS';
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var sh = aba_(A_BLOQ, CAB_BLOQ), bls = bloqueios_();
    for (var i = bls.length - 1; i >= 0; i--) {
      if (bls[i].instrutor === quem && bls[i].data === data && (tempo === 'TODOS' || bls[i].tempo === tempo || bls[i].tempo === 'TODOS')) {
        sh.deleteRow(bls[i].linha);
      }
    }
    log_(quem, 'liberou horário', data + ' ' + (tempo || 'TODOS'));
    return { ok: true };
  } finally { lock.releaseLock(); }
}

/** O instrutor cadastra ou troca o e-mail que recebe os lembretes. */
function salvarEmail(chave, email) {
  var p = papel_(chave);
  if (p.papel !== 'instrutor') throw new Error('Acesso negado.');
  var e = String(email || '').trim();
  if (e && !/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(e)) return { ok: false, erro: 'E-mail inválido.' };
  var lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    var insts = instrutores_();
    for (var i = 0; i < insts.length; i++) {
      if (insts[i].nome === p.instrutor) {
        aba_(A_INST, CAB_INST).getRange(insts[i].linha, 2).setValue(e);
        log_(p.instrutor, 'cadastrou e-mail', e ? 'ok' : 'removido');
        return { ok: true };
      }
    }
    return { ok: false, erro: 'Instrutor não encontrado.' };
  } finally { lock.releaseLock(); }
}

/* ===================== agenda e e-mail ===================== */

function emailDe_(nome) {
  var insts = instrutores_();
  for (var i = 0; i < insts.length; i++) if (insts[i].nome === nome) return insts[i].email;
  return '';
}

function titulo_(a) { return 'Vistoria de PNR · ' + a.nomeGuerra; }

function corpo_(a) {
  var t = tempo_(a.tempo);
  return [
    'Aluno: ' + a.nomeGuerra,
    'Instrutor: ' + a.instrutor,
    'Quando: ' + br_(a.data) + ', das ' + t.ini + ' às ' + t.fim,
    'PNR: ' + (a.endereco || 'não informado'),
    a.viagem ? 'Viagem do aluno: ' + br_(a.viagem) : '',
    a.novaOM ? 'Nova OM: ' + a.novaOM : '',
    a.telefone ? 'Telefone: ' + a.telefone : '',
    '',
    'A entrega é pessoal, só com o militar responsável. A família não acompanha.'
  ].filter(String).join('\n');
}

/** Evento na agenda do dono da planilha, com o instrutor como convidado. */
function criarEvento_(a) {
  try {
    var t = tempo_(a.tempo);
    var cal = CalendarApp.getDefaultCalendar();
    if (!cal) return;
    var ev = cal.createEvent(titulo_(a), quando_(a.data, t.ini), quando_(a.data, t.fim), {
      description: corpo_(a), location: a.endereco || ''
    });
    var email = emailDe_(a.instrutor);
    if (email) { try { ev.addGuest(email); } catch (e) {} }
    try { ev.addPopupReminder(60); ev.addPopupReminder(18 * 60); } catch (e) {}
    a.eventoId = ev.getId();
  } catch (e) { log_('sistema', 'falha ao criar evento', String(e)); }
}

function apagarEvento_(a) {
  if (!a.eventoId) return;
  try { var ev = CalendarApp.getDefaultCalendar().getEventById(a.eventoId); if (ev) ev.deleteEvent(); } catch (e) {}
  a.eventoId = '';
}

/** Convite .ics com alarmes na véspera e 1 h antes: entra na agenda do iPhone e do Android. */
function ics_(a) {
  var t = tempo_(a.tempo);
  var z = function (d) { return Utilities.formatDate(d, 'UTC', "yyyyMMdd'T'HHmmss'Z'"); };
  var esc = function (s) { return String(s || '').replace(/([,;\\])/g, '\\$1').replace(/\n/g, '\\n'); };
  return [
    'BEGIN:VCALENDAR', 'VERSION:2.0', 'PRODID:-//EsAO//Devolucao PNR//PT-BR', 'METHOD:REQUEST',
    'BEGIN:VEVENT',
    'UID:pnr-' + a.id + '@esao',
    'DTSTAMP:' + z(new Date()),
    'DTSTART:' + z(quando_(a.data, t.ini)),
    'DTEND:' + z(quando_(a.data, t.fim)),
    'SUMMARY:' + esc(titulo_(a)),
    'LOCATION:' + esc(a.endereco || ''),
    'DESCRIPTION:' + esc(corpo_(a)),
    'BEGIN:VALARM', 'TRIGGER:-PT1H', 'ACTION:DISPLAY', 'DESCRIPTION:Vistoria de PNR em 1 hora', 'END:VALARM',
    'BEGIN:VALARM', 'TRIGGER:-P1D', 'ACTION:DISPLAY', 'DESCRIPTION:Vistoria de PNR amanhã', 'END:VALARM',
    'END:VEVENT', 'END:VCALENDAR'
  ].join('\r\n');
}

function enviar_(para, assunto, texto, anexoAgendamento) {
  if (!para) return;
  try {
    var opt = { name: 'Devolução de PNR · EsAO' };
    if (anexoAgendamento) {
      opt.attachments = [Utilities.newBlob(ics_(anexoAgendamento), 'text/calendar', 'vistoria.ics')];
    }
    MailApp.sendEmail(para, assunto, texto, opt);
  } catch (e) { log_('sistema', 'falha ao enviar e-mail', String(e)); }
}

function avisarInstrutor_(a, tipo) {
  var email = emailDe_(a.instrutor);
  if (!email) return;
  var t = tempo_(a.tempo);
  var quando = br_(a.data) + ' às ' + t.ini;
  var m = {
    novo: ['Novo pedido de vistoria · ' + a.nomeGuerra + ' · ' + quando,
      'O ' + a.nomeGuerra + ' pediu vistoria para ' + quando + '.\n\n' + corpo_(a) + '\n\nAbra o seu link para confirmar ou recusar.'],
    encaixe: ['PEDIDO DE ENCAIXE (menos de 72 h) · ' + a.nomeGuerra + ' · ' + quando,
      'O ' + a.nomeGuerra + ' pediu um ENCAIXE para ' + quando + ', com menos de 72 h de antecedência.\n\n'
      + (a.justificativa ? 'Justificativa do aluno: ' + a.justificativa + '\n\n' : '') + corpo_(a)
      + '\n\nAbra o seu link para aprovar ou recusar.'],
    cancelado: ['Vistoria CANCELADA · ' + a.nomeGuerra + ' · ' + quando,
      'O ' + a.nomeGuerra + ' cancelou a vistoria de ' + quando + '.' + (a.motivo ? '\nMotivo: ' + a.motivo : '') + '\n\nO horário está livre de novo.'],
    transferido: ['Pedido transferido para o senhor · ' + a.nomeGuerra + ' · ' + quando,
      'Um pedido de vistoria foi transferido para o senhor.\n\n' + corpo_(a) + '\n\nAbra o seu link para confirmar ou recusar.'],
    remarcado: ['Nova vistoria após pendências · ' + a.nomeGuerra + ' · ' + quando,
      'O ' + a.nomeGuerra + ' informou que corrigiu as pendências e pediu nova vistoria para ' + quando + '.\n\n' + corpo_(a)]
  }[tipo];
  if (m) enviar_(email, m[0], m[1], tipo === 'novo' || tipo === 'transferido' ? null : null);
}

function avisarAluno_(a, tipo) {
  // O aluno é avisado na própria tela. E-mail só se ele tiver informado um.
  if (!a.email) return;
  var t = tempo_(a.tempo);
  var quando = br_(a.data) + ' às ' + t.ini;
  var m = {
    confirmado: ['Vistoria CONFIRMADA · ' + quando,
      'Sua vistoria está confirmada para ' + quando + ' com o ' + a.instrutor + '.\n\n'
      + 'No dia, leve TODAS as chaves e cópias, os controles e os comprovantes da Light e da Naturgy.\n'
      + 'Compareça pessoalmente e sozinho: a família não acompanha a vistoria.', a],
    recusado: ['Vistoria RECUSADA · ' + quando,
      'O ' + a.instrutor + ' não pôde confirmar ' + quando + '.' + (a.motivo ? '\nMotivo: ' + a.motivo : '')
      + (a.sugestao ? '\nSugestão do instrutor: ' + a.sugestao : '') + '\n\nAbra o link e escolha outro horário.'],
    transferido: ['Seu pedido passou para o ' + a.instrutor,
      'Seu pedido de vistoria foi transferido para o ' + a.instrutor + ', que é o seu instrutor responsável. Aguarde a confirmação dele.'],
    entregue: ['Chaves recebidas · PNR entregue',
      'O ' + a.instrutor + ' registrou a vistoria e recebeu as chaves do seu PNR. Processo encerrado.'],
    pendencias: ['Vistoria com PENDÊNCIAS · corrija e marque de novo',
      'A vistoria foi realizada, mas há pendências a corrigir:\n\n- ' + (a.pendencias || []).join('\n- ')
      + '\n\nAs chaves não são recebidas enquanto houver pendência. Corrija e marque nova vistoria pelo link.']
  }[tipo];
  if (m) enviar_(a.email, m[0], m[1], m[2]);
}

/* ===================== lembretes diários ===================== */

/** Cria os gatilhos das 18 h (véspera) e das 7 h (dia). Chamado pelo menu da planilha. */
function ativarLembretes() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'lembreteVespera' || t.getHandlerFunction() === 'lembreteManha') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('lembreteVespera').timeBased().atHour(18).nearMinute(0).everyDays(1).inTimezone(TZ).create();
  ScriptApp.newTrigger('lembreteManha').timeBased().atHour(7).nearMinute(0).everyDays(1).inTimezone(TZ).create();
  log_('sistema', 'ativou lembretes diários', '18 h e 7 h');
  try {
    SpreadsheetApp.getUi().alert('Lembretes ativados.\n\nTodos os dias às 18 h os instrutores recebem as vistorias de amanhã, e às 7 h as de hoje.');
  } catch (e) {}
}

function lembreteVespera() { lembrete_(1, 'amanhã'); }
function lembreteManha() { lembrete_(0, 'hoje'); }

function lembrete_(somaDias, palavra) {
  var dia = Utilities.formatDate(new Date(Date.now() + somaDias * 864e5), TZ, 'yyyy-MM-dd');
  var ags = agendamentos_().filter(function (a) { return a.situacao === 'confirmado' && a.data === dia; });
  if (!ags.length) return;
  var por = {};
  ags.forEach(function (a) { (por[a.instrutor] = por[a.instrutor] || []).push(a); });
  Object.keys(por).forEach(function (nome) {
    var email = emailDe_(nome);
    if (!email) return;
    var lista = por[nome].sort(function (x, y) { return x.tempo < y.tempo ? -1 : 1; });
    var txt = lista.map(function (a) {
      var t = tempo_(a.tempo);
      return t.ini + ' às ' + t.fim + ' · ' + a.nomeGuerra + '\n   ' + (a.endereco || 'endereço não informado')
        + (a.viagem ? '\n   Viagem do aluno: ' + br_(a.viagem) : '');
    }).join('\n\n');
    enviar_(email,
      'Vistorias de ' + palavra + ' (' + br_(dia) + '): ' + lista.length,
      'Senhor ' + nome + ',\n\n' + palavra.charAt(0).toUpperCase() + palavra.slice(1)
      + ', ' + br_(dia) + ', o senhor tem ' + lista.length + ' vistoria(s) de PNR:\n\n' + txt
      + '\n\nLembre-se: a entrega é pessoal, só com o militar responsável, e a família não acompanha.'
      + '\nSe houver pendência, a chave não é recebida.');
  });
  log_('sistema', 'enviou lembretes de ' + palavra, ags.length + ' vistoria(s) em ' + dia);
}

/* ===================== menu da planilha ===================== */

function onOpen() {
  SpreadsheetApp.getUi().createMenu('PNR')
    .addItem('Ver links (aluno, instrutores e painel)', 'verLinks')
    .addItem('Ativar lembretes diários (18 h e 7 h)', 'ativarLembretes')
    .addItem('Gerar chaves que faltam', 'gerarChaves')
    .addToUi();
}

function verLinks() {
  garantirInstalacao_();
  var url = ScriptApp.getService().getUrl();
  if (!url) {
    SpreadsheetApp.getUi().alert('O sistema ainda não foi implantado.\n\nNo Apps Script, use Implantar › Nova implantação › App da Web.');
    return;
  }
  var chave = PropertiesService.getScriptProperties().getProperty('PAINEL_CHAVE');
  var a = function (u, t) { return '<a href="' + u + '" target="_blank" style="word-break:break-all">' + (t || u) + '</a>'; };
  var html = '<p><b>1. Seu painel de coordenação</b> — não repasse:<br>' + a(url + '?v=painel&chave=' + chave) + '</p>'
    + '<p><b>2. Link de cada instrutor</b> — é isto que o senhor divulga: mande a cada instrutor,'
    + ' individualmente, apenas a linha dele. Dentro da tela, o próprio instrutor copia o link dos'
    + ' orientados dele e repassa à sua turma.</p><ul>';
  instrutores_().forEach(function (i) {
    html += '<li style="margin-bottom:10px"><b>' + i.nome + '</b>'
      + '<br><i>Tela do instrutor:</i> ' + a(url + '?v=instrutor&chave=' + i.chave)
      + '<br><i>Link dos orientados dele:</i> ' + a(url + '?i=' + i.chaveAlunos)
      + '</li>';
  });
  html += '</ul><p><b>3. Link geral</b> — só para quem ficou sem o link do instrutor. Nele o aluno'
    + ' escolhe o instrutor na lista:<br>' + a(url) + '</p>';
  SpreadsheetApp.getUi().showModalDialog(
    HtmlService.createHtmlOutput('<div style="font:14px Arial,sans-serif;line-height:1.5">' + html + '</div>').setWidth(680).setHeight(600),
    'Links do sistema de devolução de PNR');
}

function gerarChaves() {
  var sh = aba_(A_INST, CAB_INST);
  var n = sh.getLastRow();
  if (n < 2) return;
  var mudou = gerarChaves_();
  if (mudou) log_('sistema', 'gerou chaves', mudou + ' chave(s)');
  try {
    SpreadsheetApp.getUi().alert(mudou ? mudou + ' chave(s) gerada(s).' : 'Todos os instrutores já têm as duas chaves.');
  } catch (e) {}
}

/* ===================== instalação ===================== */

function garantirInstalacao_() {
  if (!PropertiesService.getScriptProperties().getProperty('INSTALADO')) instalar();
}

/** Cria as abas, lança os instrutores da escala 2026 e gera as chaves. Roda sozinha no primeiro acesso. */
function instalar() {
  var props = PropertiesService.getScriptProperties();

  var shI = aba_(A_INST, CAB_INST);
  if (shI.getLastRow() < 2) {
    var nova = function () { return Utilities.getUuid().replace(/-/g, '').slice(0, 16); };
    var rows = SEED_INSTRUTORES.map(function (s) {
      return [s[0], '', s[2], s[1], nova(), nova(), 'sim'];
    });
    shI.getRange(2, 1, rows.length, CAB_INST.length).setValues(rows);
    shI.setColumnWidth(3, 220); shI.setColumnWidth(5, 180); shI.setColumnWidth(6, 180);
    shI.getRange(1, 1, 1, CAB_INST.length).setNote(
      'Edite aqui sem mexer no código.\n'
      + '• disponibilidade: "Até 04/12", "13/11 a 26/11", "Até 19/11; 30/11 a 22/12", "03/12 e 04/12"\n'
      + '• e-mail: usado nos lembretes da véspera (18 h) e do dia (7 h)\n'
      + '• chave do instrutor: abre a tela dele. Mande só para ele.\n'
      + '• chave do link dos orientados: vai no link que o instrutor passa aos alunos dele.\n'
      + '  Quem abre esse link só consegue agendar com esse instrutor.\n'
      + '• ativo: escreva "não" para tirar o instrutor da lista');
  }

  aba_(A_AGEND, CAB_AGEND);
  aba_(A_BLOQ, CAB_BLOQ);
  aba_(A_ACEITE, CAB_ACEITE);
  aba_(A_LOG, CAB_LOG);

  var shC = aba_(A_CONFIG, ['chave', 'valor', 'o que é']);
  if (shC.getLastRow() < 2) {
    shC.getRange(2, 1, 3, 3).setValues([
      ['inicio', INI_PADRAO, 'Primeiro dia da campanha (datas anteriores nunca aparecem)'],
      ['fim', FIM_PADRAO, 'Último dia da campanha'],
      ['horasMinimas', HORAS_MIN, 'Antecedência mínima para o aluno agendar sozinho. Abaixo disso, só com encaixe aprovado pelo instrutor.']
    ]);
    shC.setColumnWidth(3, 460);
  }

  if (!props.getProperty('PAINEL_CHAVE')) props.setProperty('PAINEL_CHAVE', Utilities.getUuid().replace(/-/g, '').slice(0, 16));
  gerarChaves_();
  props.setProperty('INSTALADO', '1');
  log_('sistema', 'instalou', 'abas, instrutores e chaves');
}

/** Cada instrutor tem duas chaves: a dele (coluna 5) e a do link dos orientados (coluna 6). */
function gerarChaves_() {
  var sh = aba_(A_INST, CAB_INST);
  var n = sh.getLastRow();
  if (n < 2) return 0;
  var vals = sh.getRange(2, 1, n - 1, CAB_INST.length).getValues(), mudou = 0;
  var nova = function () { return Utilities.getUuid().replace(/-/g, '').slice(0, 16); };
  for (var i = 0; i < vals.length; i++) {
    if (!String(vals[i][0] || '').trim()) continue;
    if (!String(vals[i][4] || '').trim()) { sh.getRange(i + 2, 5).setValue(nova()); mudou++; }
    if (!String(vals[i][5] || '').trim()) { sh.getRange(i + 2, 6).setValue(nova()); mudou++; }
  }
  return mudou;
}

/** Página das três telas (index.html). */
var PAGINA = "<title>Devolução de PNR · EsAO</title>\n<link rel=\"preconnect\" href=\"https://fonts.googleapis.com\">\n<link rel=\"preconnect\" href=\"https://fonts.gstatic.com\" crossorigin>\n<link rel=\"stylesheet\" href=\"https://fonts.googleapis.com/css2?family=Barlow+Condensed:wght@500;600;700&family=Public+Sans:wght@400;500;600;700&display=swap\">\n<style>\n/* Celular primeiro: uma coluna estreita. O painel da coordenação abre em duas colunas a partir de 960px. */\n:root {\n  --bg: #f1f3ef;\n  --surface: #ffffff;\n  --surface-2: #e8ebe5;\n  --ink: #15201a;\n  --muted: #56615a;\n  --line: #d3d9d1;\n  --accent: #2c5a37;\n  --accent-ink: #ffffff;\n  --accent-soft: #dce9df;\n  --heat: #2c5a37;\n  --ok: #1a7a2a; --ok-soft: #dff1e1;\n  --warn: #8a5a00; --warn-soft: #fcefd2;\n  --bad: #b3261e; --bad-soft: #fbe2df;\n  --pend: #5b5f6b; --pend-soft: #e7e8ec;\n  --focus: #1f6feb;\n  --f-display: \"Barlow Condensed\", \"Arial Narrow\", \"Roboto Condensed\", system-ui, sans-serif;\n  --f-body: \"Public Sans\", system-ui, -apple-system, \"Segoe UI\", Roboto, sans-serif;\n  color-scheme: light;\n}\n@media (prefers-color-scheme: dark) {\n  :root:not([data-theme=\"light\"]) {\n    --bg: #0f1311; --surface: #181d1a; --surface-2: #222924; --ink: #ecf0ea; --muted: #a2ada5; --line: #303833;\n    --accent: #86c495; --accent-ink: #0d1a10; --accent-soft: #23382a; --heat: #86c495;\n    --ok: #6fd07c; --ok-soft: #1d3322; --warn: #f1bf55; --warn-soft: #3a2f14; --bad: #ff8a80; --bad-soft: #3d1f1c;\n    --pend: #b8bcc8; --pend-soft: #2a2c33; --focus: #7aa8ff;\n    color-scheme: dark;\n  }\n}\n:root[data-theme=\"dark\"] {\n  --bg: #0f1311; --surface: #181d1a; --surface-2: #222924; --ink: #ecf0ea; --muted: #a2ada5; --line: #303833;\n  --accent: #86c495; --accent-ink: #0d1a10; --accent-soft: #23382a; --heat: #86c495;\n  --ok: #6fd07c; --ok-soft: #1d3322; --warn: #f1bf55; --warn-soft: #3a2f14; --bad: #ff8a80; --bad-soft: #3d1f1c;\n  --pend: #b8bcc8; --pend-soft: #2a2c33; --focus: #7aa8ff;\n  color-scheme: dark;\n}\n\n* { box-sizing: border-box; }\n[hidden] { display: none !important; }\nhtml { -webkit-text-size-adjust: 100%; }\nbody { background: var(--bg); color: var(--ink); font: 15px/1.5 var(--f-body); margin: 0; }\nbutton, input, select, textarea { font: inherit; color: inherit; }\nbutton { cursor: pointer; }\n:focus-visible { outline: 3px solid var(--focus); outline-offset: 2px; }\nh1, h2, h3 { font-family: var(--f-display); text-wrap: balance; margin: 0; letter-spacing: .01em; }\n.num { font-variant-numeric: tabular-nums; }\n.sr { position: absolute; width: 1px; height: 1px; overflow: hidden; clip: rect(0 0 0 0); white-space: nowrap; }\n\n/* ---------- topo ---------- */\n.top { background: var(--surface); border-bottom: 1px solid var(--line); position: sticky; top: 0; z-index: 20; }\n.top-in { max-width: 1240px; margin: 0 auto; padding: 10px 16px; display: flex; align-items: center; gap: 12px; flex-wrap: wrap; }\n.brand { display: flex; flex-direction: column; min-width: 0; flex: 1 1 200px; }\n.eyebrow { font-size: 11px; text-transform: uppercase; letter-spacing: .08em; color: var(--muted); font-weight: 600; }\n.brand h1 { font-size: 21px; line-height: 1.1; font-weight: 700; }\n.views { display: flex; background: var(--surface-2); border-radius: 10px; padding: 3px; gap: 2px; flex-wrap: wrap; }\n.views button { border: 0; background: transparent; padding: 7px 13px; border-radius: 8px; font-weight: 600; font-size: 14px; color: var(--muted); }\n.views button[aria-pressed=\"true\"] { background: var(--surface); color: var(--ink); box-shadow: 0 1px 2px rgb(0 0 0 / .12); }\n.conn { font-size: 12px; color: var(--muted); display: flex; align-items: center; gap: 6px; width: 100%; }\n.conn i { width: 8px; height: 8px; border-radius: 50%; background: var(--pend); display: inline-block; flex: none; }\n.conn.on i { background: var(--ok); }\n.conn.off i { background: var(--bad); }\n.banner { max-width: 1240px; margin: 12px auto 0; padding: 0 16px; }\n.banner > div { border-radius: 10px; padding: 10px 12px; font-size: 14px; border: 1px solid var(--line); background: var(--surface); }\n.banner .w { background: var(--warn-soft); border-color: transparent; }\n.banner .b { background: var(--bad-soft); border-color: transparent; }\n\n/* ---------- blocos comuns ---------- */\n.wrap { max-width: 640px; margin: 0 auto; padding: 16px 16px 130px; }\n.wide { max-width: 1240px; }\n.card { background: var(--surface); border: 1px solid var(--line); border-radius: 14px; padding: 14px; }\n.stack { display: grid; gap: 10px; }\n.q { font-size: 27px; font-weight: 700; line-height: 1.12; margin-bottom: 6px; }\n.hint { color: var(--muted); margin: 0 0 14px; max-width: 62ch; }\n.btn { border: 1px solid var(--line); background: var(--surface); border-radius: 10px; padding: 9px 14px; font-weight: 600; font-size: 14px; }\n.btn.sm { padding: 6px 10px; font-size: 13px; }\n.btn.pri { background: var(--accent); color: var(--accent-ink); border-color: var(--accent); }\n.btn.danger { color: var(--bad); }\n.btn.big { width: 100%; padding: 14px; font-size: 16px; border-radius: 12px; }\n.btn:disabled { opacity: .5; cursor: not-allowed; }\n.row-btns { display: flex; gap: 8px; flex-wrap: wrap; }\n.pill { display: inline-flex; align-items: center; gap: 4px; font-size: 12px; font-weight: 600; border-radius: 999px; padding: 2px 9px; width: fit-content; white-space: nowrap; }\n.pill.ok { background: var(--ok-soft); color: var(--ok); }\n.pill.bad { background: var(--bad-soft); color: var(--bad); }\n.pill.pend { background: var(--pend-soft); color: var(--pend); }\n.pill.warn { background: var(--warn-soft); color: var(--warn); }\n.pill.acc { background: var(--accent-soft); color: var(--accent); }\n.field { display: grid; gap: 6px; margin-top: 14px; }\n.field label, .lab { font-weight: 700; font-size: 14px; }\n.field small { color: var(--muted); font-size: 13px; }\n.field input, .field select, .field textarea { border: 1px solid var(--line); background: var(--surface); border-radius: 10px; padding: 11px 12px; width: 100%; }\n.field textarea { min-height: 72px; resize: vertical; }\n.empty { color: var(--muted); font-size: 14px; padding: 8px 0; }\n.sep { height: 1px; background: var(--line); margin: 14px 0; }\n\n/* explicação fixa no topo de cada tela: onde estou, o que faço, o que acontece depois */\n.guide { background: var(--accent-soft); border-radius: 12px; padding: 12px 14px; margin-bottom: 14px; font-size: 14px; }\n.guide b { display: block; font-family: var(--f-display); font-size: 17px; margin-bottom: 2px; }\n.guide p { margin: 4px 0 0; }\n.guide .after { color: var(--muted); font-size: 13px; margin-top: 6px; }\n\n/* etapas do aluno */\n.steps { display: grid; grid-template-columns: repeat(6, 1fr); gap: 4px; margin-bottom: 16px; }\n.steps button { border: 0; background: none; padding: 0; text-align: left; display: flex; flex-direction: column; gap: 5px; min-width: 0; }\n.steps button:disabled { cursor: default; }\n.steps .bar { height: 4px; border-radius: 2px; background: var(--line); }\n.steps .on .bar, .steps .done .bar { background: var(--accent); }\n.steps .lbl { font-size: 10px; color: var(--muted); white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }\n.steps .on .lbl { color: var(--ink); font-weight: 700; }\n.stepn { font-size: 12px; color: var(--muted); font-weight: 600; text-transform: uppercase; letter-spacing: .06em; }\n\n/* lista de opções grandes */\n.opts { display: grid; gap: 8px; }\n.opt { border: 1px solid var(--line); background: var(--surface); border-radius: 12px; padding: 12px 14px; text-align: left; display: grid; grid-template-columns: 22px 1fr; gap: 2px 10px; align-items: start; width: 100%; }\n.opt .radio { width: 20px; height: 20px; border-radius: 50%; border: 2px solid var(--muted); margin-top: 2px; display: grid; place-items: center; }\n.opt[aria-checked=\"true\"] { border-color: var(--accent); box-shadow: 0 0 0 2px var(--accent) inset; }\n.opt[aria-checked=\"true\"] .radio { border-color: var(--accent); }\n.opt[aria-checked=\"true\"] .radio::after { content: \"\"; width: 10px; height: 10px; border-radius: 50%; background: var(--accent); }\n.opt strong { font-size: 16px; }\n.opt small, .opt .sub { grid-column: 2; color: var(--muted); font-size: 13px; }\n.opt .tags { grid-column: 2; display: flex; gap: 6px; flex-wrap: wrap; margin-top: 4px; }\n\n/* caixas de marcar */\n.chk { border: 1px solid var(--line); background: var(--surface); border-radius: 12px; padding: 12px 14px; text-align: left; display: grid; grid-template-columns: 24px 1fr; gap: 2px 10px; align-items: start; width: 100%; }\n.chk .box { width: 22px; height: 22px; border-radius: 6px; border: 2px solid var(--muted); display: grid; place-items: center; font-weight: 800; font-size: 15px; color: transparent; margin-top: 1px; }\n.chk[aria-checked=\"true\"] { border-color: var(--accent); }\n.chk[aria-checked=\"true\"] .box { border-color: var(--accent); background: var(--accent); color: var(--accent-ink); }\n.chk strong { font-size: 15px; font-weight: 700; }\n.chk .sub { grid-column: 2; color: var(--muted); font-size: 13px; }\n.chk.crit { border-left: 4px solid var(--bad); }\n\n/* termos */\n.termo { max-height: 46vh; overflow: auto; border: 1px solid var(--line); border-radius: 12px; padding: 14px; background: var(--surface); font-size: 14px; }\n.termo h3 { font-size: 17px; margin: 14px 0 4px; font-family: var(--f-body); font-weight: 700; }\n.termo h3:first-child { margin-top: 0; }\n.termo p, .termo li { margin: 4px 0; }\n.termo ul { padding-left: 20px; margin: 4px 0; }\n.termo .crit { background: var(--bad-soft); border-radius: 8px; padding: 8px 10px; margin: 8px 0; }\n.scrollmark { font-size: 13px; color: var(--muted); margin-top: 8px; display: flex; gap: 6px; align-items: center; }\n.scrollmark.done { color: var(--ok); font-weight: 600; }\n\n/* \"como fazer\" expansível */\n.howto { grid-column: 1 / -1; margin-top: 8px; background: var(--surface-2); border-radius: 10px; padding: 10px 12px; font-size: 13px; }\n.howto ol, .howto ul { padding-left: 18px; margin: 4px 0; }\n.howto a { color: var(--focus); }\n.more { grid-column: 2; justify-self: start; border: 0; background: none; color: var(--accent); font-weight: 700; font-size: 13px; padding: 4px 0; text-decoration: underline; }\n\n/* progresso do checklist */\n.prog { background: var(--surface-2); border-radius: 999px; height: 10px; overflow: hidden; margin: 6px 0 14px; }\n.prog i { display: block; height: 100%; background: var(--accent); border-radius: 999px; }\n\n/* calendário */\n.mtabs { display: grid; grid-auto-flow: column; gap: 4px; background: var(--surface-2); padding: 3px; border-radius: 10px; }\n.mtabs button { border: 0; background: transparent; padding: 8px 4px; border-radius: 8px; font-weight: 600; font-size: 14px; color: var(--muted); }\n.mtabs button[aria-pressed=\"true\"] { background: var(--surface); color: var(--ink); box-shadow: 0 1px 2px rgb(0 0 0 / .12); }\n.wk, .grid { display: grid; grid-template-columns: repeat(7, minmax(0, 1fr)); gap: 3px; }\n.wk { margin: 12px 0 4px; }\n.wk span { font-size: 11px; text-align: center; color: var(--muted); font-weight: 600; text-transform: uppercase; letter-spacing: .04em; }\n.day { position: relative; border: 1px solid transparent; background: var(--surface); border-radius: 8px; min-height: 52px; display: flex; flex-direction: column; align-items: center; justify-content: center; gap: 2px; font-weight: 600; font-variant-numeric: tabular-nums; font-size: 15px; padding: 2px; }\n.day.blank { background: transparent; }\n.day .dn { line-height: 1; }\n.day .vg { font-size: 10px; font-weight: 700; line-height: 1; }\n.day.livre { background: var(--ok-soft); color: var(--ok); border-color: color-mix(in srgb, var(--ok) 35%, transparent); }\n.day.curto { background: var(--warn-soft); color: var(--warn); border-color: color-mix(in srgb, var(--warn) 35%, transparent); }\n.day.nao { background: var(--surface-2); color: var(--muted); }\n.day.cheio { background: var(--pend-soft); color: var(--pend); }\n.day.sel { outline: 3px solid var(--ink); outline-offset: -1px; z-index: 1; }\n.day.ideal::after { content: \"\"; position: absolute; top: 3px; right: 4px; width: 6px; height: 6px; border-radius: 50%; background: var(--ok); }\n.legend { display: flex; flex-wrap: wrap; gap: 6px 14px; font-size: 12px; color: var(--muted); margin-top: 10px; }\n.legend span { display: inline-flex; gap: 5px; align-items: center; }\n.legend i { width: 14px; height: 12px; border-radius: 3px; display: inline-block; border: 1px solid var(--line); }\n\n/* lista de tempos do dia escolhido */\n.slots { display: grid; gap: 8px; margin-top: 14px; }\n.slot { border: 1px solid var(--line); background: var(--surface); border-radius: 12px; padding: 12px 14px; text-align: left; display: grid; grid-template-columns: 1fr auto; gap: 4px 10px; align-items: center; width: 100%; }\n.slot .h { font-family: var(--f-display); font-size: 20px; font-weight: 700; }\n.slot .why { grid-column: 1 / -1; font-size: 13px; color: var(--muted); }\n.slot.livre { border-color: color-mix(in srgb, var(--ok) 45%, transparent); }\n.slot.off { background: var(--surface-2); }\n.turno-h { font-size: 12px; text-transform: uppercase; letter-spacing: .07em; color: var(--muted); font-weight: 700; margin-top: 10px; }\n\n/* cartão do próximo passo */\n.next { border: 2px solid var(--accent); border-radius: 14px; padding: 14px; background: var(--surface); }\n.next.okb { border-color: var(--ok); }\n.next.badb { border-color: var(--bad); }\n.next.warnb { border-color: var(--warn); }\n.next .big { font-family: var(--f-display); font-size: 25px; font-weight: 700; line-height: 1.1; }\n.next .when { font-family: var(--f-display); font-size: 21px; font-weight: 700; margin-top: 4px; }\n.next dl { margin: 10px 0 0; display: grid; grid-template-columns: auto 1fr; gap: 3px 10px; font-size: 14px; }\n.next dt { color: var(--muted); }\n.next dd { margin: 0; }\n.leve { background: var(--warn-soft); border-radius: 10px; padding: 10px 12px; margin-top: 12px; font-size: 14px; }\n.leve ul { margin: 4px 0 0; padding-left: 20px; }\n.countdown { display: flex; gap: 8px; margin-top: 12px; }\n.countdown div { flex: 1; background: var(--surface-2); border-radius: 10px; padding: 8px 10px; text-align: center; }\n.countdown b { display: block; font-family: var(--f-display); font-size: 26px; line-height: 1; }\n.countdown span { font-size: 11px; color: var(--muted); text-transform: uppercase; letter-spacing: .05em; }\n\n/* painel / instrutor */\n.kpis { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 8px; }\n.kpi { background: var(--surface); border: 1px solid var(--line); border-radius: 12px; padding: 10px 12px; min-width: 0; }\n.kpi .v { font-family: var(--f-display); font-size: 32px; font-weight: 700; line-height: 1; }\n.kpi .l { font-size: 13px; color: var(--muted); }\n.kpi.ok .v { color: var(--ok); } .kpi.bad .v { color: var(--bad); } .kpi.pend .v { color: var(--pend); } .kpi.warn .v { color: var(--warn); }\n.sec { background: var(--surface); border: 1px solid var(--line); border-radius: 14px; padding: 14px; min-width: 0; }\n.sec-h { display: flex; justify-content: space-between; align-items: baseline; gap: 10px; flex-wrap: wrap; margin-bottom: 10px; }\n.sec-h h2 { font-size: 21px; font-weight: 700; }\n.sec-h p { margin: 0; color: var(--muted); font-size: 13px; }\n.dash { max-width: 1240px; margin: 0 auto; padding: 16px 16px 60px; display: grid; gap: 14px; }\n.ped { border: 1px solid var(--line); border-radius: 12px; padding: 12px; display: grid; gap: 8px; }\n.ped.urg { border-color: var(--warn); background: var(--warn-soft); }\n.ped h3 { font-family: var(--f-body); font-size: 16px; font-weight: 700; }\n.ped dl { margin: 0; display: grid; grid-template-columns: auto 1fr; gap: 2px 10px; font-size: 13px; }\n.ped dt { color: var(--muted); }\n.ped dd { margin: 0; }\n.ped .acts { display: flex; gap: 6px; flex-wrap: wrap; }\n.todo { background: var(--accent-soft); border-radius: 8px; padding: 8px 10px; font-size: 13px; font-weight: 600; }\n.tbl { width: 100%; border-collapse: collapse; font-size: 14px; }\n.tbl th, .tbl td { text-align: left; padding: 8px 6px; border-bottom: 1px solid var(--line); }\n.tbl th { font-size: 12px; text-transform: uppercase; letter-spacing: .05em; color: var(--muted); font-weight: 600; }\n.tbl .num { text-align: right; font-variant-numeric: tabular-nums; }\n.bar2 { height: 8px; background: var(--surface-2); border-radius: 4px; overflow: hidden; min-width: 60px; }\n.bar2 i { display: block; height: 100%; background: var(--accent); }\n.alerta { border-radius: 10px; padding: 10px 12px; font-size: 14px; display: grid; gap: 8px; margin-bottom: 8px; }\n.alerta.w { background: var(--warn-soft); }\n.alerta.b { background: var(--bad-soft); }\n.alerta b { font-weight: 700; }\n.heat { display: grid; grid-template-columns: repeat(7, minmax(0, 1fr)); gap: 3px; }\n.heat div { border-radius: 6px; min-height: 44px; padding: 3px 4px; font-size: 11px; background: color-mix(in srgb, var(--heat) var(--h, 0%), var(--surface)); border: 1px solid var(--line); }\n.heat div.hot { color: var(--accent-ink); }\n.heat b { display: block; font-family: var(--f-display); font-size: 16px; }\n.evl { list-style: none; margin: 0; padding: 0; }\n.evl li { padding: 7px 0; border-bottom: 1px solid var(--line); font-size: 13px; }\n.evl li:last-child { border-bottom: 0; }\n.evl small { color: var(--muted); display: block; }\n\n/* roteiro de vistoria */\n.rot { display: grid; gap: 6px; }\n.rot .it { border: 1px solid var(--line); border-radius: 10px; padding: 10px 12px; display: grid; grid-template-columns: 1fr auto auto; gap: 8px; align-items: center; }\n.rot .it strong { font-size: 14px; font-weight: 700; }\n.rot .it small { grid-column: 1 / -1; color: var(--muted); font-size: 12px; }\n.rot .it button { border: 1px solid var(--line); background: var(--surface); border-radius: 8px; padding: 6px 10px; font-weight: 700; font-size: 14px; }\n.rot .it .y[aria-pressed=\"true\"] { background: var(--ok-soft); color: var(--ok); border-color: var(--ok); }\n.rot .it .n[aria-pressed=\"true\"] { background: var(--bad-soft); color: var(--bad); border-color: var(--bad); }\n\n/* ajuda, modal, toast */\n.helpbtn { position: fixed; right: 14px; bottom: calc(86px + env(safe-area-inset-bottom, 0px)); z-index: 18; border: 1px solid var(--line); background: var(--surface); border-radius: 999px; padding: 10px 16px; font-weight: 700; box-shadow: 0 2px 10px rgb(0 0 0 / .12); }\n.navbar { position: fixed; left: 0; right: 0; bottom: 0; background: var(--surface); border-top: 1px solid var(--line); padding: 10px 16px calc(10px + env(safe-area-inset-bottom, 0px)); z-index: 19; }\n.navbar-in { max-width: 640px; margin: 0 auto; display: flex; gap: 10px; }\n.navbar .btn { flex: 1; padding: 13px; font-size: 16px; }\n.navbar .btn.pri { flex: 2; }\n.modal { position: fixed; inset: 0; background: rgb(0 0 0 / .45); z-index: 50; display: flex; align-items: flex-end; justify-content: center; padding: 12px; }\n.modal .card { width: 100%; max-width: 580px; max-height: 88vh; overflow: auto; }\n.modal h2 { font-size: 22px; margin-bottom: 8px; }\n.toast { position: fixed; left: 50%; transform: translateX(-50%); bottom: calc(92px + env(safe-area-inset-bottom, 0px)); background: var(--ink); color: var(--bg); padding: 11px 16px; border-radius: 10px; font-weight: 600; z-index: 60; max-width: calc(100% - 32px); }\n.faq { display: grid; gap: 6px; }\n.faq details { border: 1px solid var(--line); border-radius: 10px; padding: 10px 12px; background: var(--surface); }\n.faq summary { font-weight: 700; cursor: pointer; font-size: 14px; }\n.faq p { margin: 6px 0 0; font-size: 14px; }\n\n@media (min-width: 640px) {\n  .kpis { grid-template-columns: repeat(4, minmax(0, 1fr)); }\n  .day { min-height: 62px; }\n  .modal { align-items: center; }\n}\n@media (min-width: 960px) {\n  .dash { grid-template-columns: minmax(0, 1.3fr) minmax(0, 1fr); align-items: start; }\n  .dash .full { grid-column: 1 / -1; }\n}\n@media (prefers-reduced-motion: no-preference) {\n  .btn, .opt, .chk, .day, .slot { transition: background-color .12s, box-shadow .12s, border-color .12s; }\n}\n</style>\n\n<header class=\"top\">\n  <div class=\"top-in\">\n    <div class=\"brand\">\n      <span class=\"eyebrow\" id=\"eyebrow\">EsAO · Devolução de PNR 2026</span>\n      <h1 id=\"h1\">Agendar a entrega do PNR</h1>\n    </div>\n    <div class=\"views\" id=\"views\" role=\"group\" aria-label=\"Escolher tela\" hidden></div>\n    <div class=\"conn\" id=\"conn\"><i></i><span>Conectando…</span></div>\n  </div>\n</header>\n<div class=\"banner\" id=\"banner\" hidden></div>\n<main id=\"app\"><div class=\"wrap\"><p class=\"empty\">Carregando…</p></div></main>\n<div id=\"layer\"></div>\n\n<script>\n\"use strict\";\n\n/* ===================== textos fixos ===================== */\n\nconst TEMPOS_PADRAO = [\n  { k: \"M1\", ini: \"08:00\", fim: \"09:15\", turno: \"manha\" },\n  { k: \"M2\", ini: \"09:30\", fim: \"10:45\", turno: \"manha\" },\n  { k: \"M3\", ini: \"11:00\", fim: \"12:15\", turno: \"manha\" },\n  { k: \"T1\", ini: \"13:30\", fim: \"14:45\", turno: \"tarde\" },\n  { k: \"T2\", ini: \"15:00\", fim: \"16:15\", turno: \"tarde\" },\n  { k: \"T3\", ini: \"16:30\", fim: \"17:45\", turno: \"tarde\" }\n];\n\nconst STEPS = [\"Como funciona\", \"Quem é você\", \"Termos\", \"Checklist\", \"Seu PNR\", \"Horário\"];\n\n/* Os 3 pontos que mais geram problema: ficam em caixa separada e exigem aceite individual. */\nconst CRITICOS = [\n  { k: \"pessoal\", t: \"A entrega é só minha, pessoalmente\", d: \"Não posso mandar esposa, parente nem companheiro de turma. Não posso viajar deixando a entrega para outro militar.\" },\n  { k: \"prazo\", t: \"Não vou marcar no dia da viagem\", d: \"Vou marcar 1 a 3 dias antes, de preferência 2 ou 3. Se houver pendência, a chave não é recebida até eu corrigir.\" },\n  { k: \"sozinho\", t: \"Vou estar sozinho na vistoria\", d: \"A vistoria é feita apenas com o militar responsável. A família não pode acompanhar.\" }\n];\n\nconst TERMO_HTML = `\n<h3>Termo de Ciência e Compromisso para devolução de PNR</h3>\n<p>Declaro estar ciente de que a desocupação do Próprio Nacional Residencial é regida pelas Instruções Gerais para a Administração dos PNR do Exército (Portaria C Ex nº 1.846, de 4 de outubro de 2022) e pelas determinações desta OM, e me comprometo com o seguinte.</p>\n\n<h3>1. Estado do imóvel</h3>\n<p>A desocupação será precedida de <b>vistoria</b>, que resultará no <b>Termo de Desocupação de PNR</b>. O imóvel será comparado com o <b>Termo de Recebimento/Vistoria Inicial</b> que preenchi na ocupação. Devo entregá-lo:</p>\n<ul>\n  <li><b>pintado</b> e <b>faxinado</b>, sem pertences pessoais nem entulho;</li>\n  <li>com <b>iluminação completa</b>: lâmpadas, luminárias, interruptores e tomadas funcionando;</li>\n  <li>com os equipamentos previstos no termo inicial <b>instalados e Em Condições De uso</b>, ou seja, em boas condições e funcionando, a exemplo do ar-condicionado;</li>\n  <li>com louças, metais, portas, janelas, fechaduras e demais itens em perfeito estado;</li>\n  <li>com <b>todas as chaves e cópias</b>, além de controles e tags de acesso, se houver.</li>\n</ul>\n\n<h3>2. Débitos e concessionárias</h3>\n<p>Devo comprovar a <b>quitação de todos os débitos referentes ao PNR</b>, incluindo o <b>encerramento da conta e a quitação de débitos na Light</b> (energia) e <b>na Naturgy</b> (gás), apresentando os comprovantes no dia da vistoria.</p>\n\n<h3>3. Entrega pessoal</h3>\n<div class=\"crit\"><p>A entrega será feita <b>exclusivamente por mim, militar responsável pelo PNR</b>. Não serão aceitos parentes, cônjuges ou companheiros de turma como representantes. Portanto, <b>não poderei viajar, sair de férias ou me ausentar</b> deixando essa responsabilidade a terceiros.</p></div>\n\n<h3>4. Presença apenas do militar</h3>\n<div class=\"crit\"><p>A vistoria será acompanhada <b>somente por mim</b>. Familiares não poderão acompanhar a entrega do PNR.</p></div>\n\n<h3>5. Possibilidade de pendências e prazo de segurança</h3>\n<div class=\"crit\"><p>Estou ciente de que se trata de uma <b>verificação</b>, que pode apontar pendências como falta de pintura, item faltante ou material danificado. <b>As chaves não serão recebidas enquanto houver pendência.</b> Por isso devo agendar a vistoria com <b>1 a 3 dias de antecedência em relação à data da viagem</b>, e <b>nunca no mesmo dia do embarque</b>, para ter tempo de fazer as correções.</p></div>\n\n<h3>6. Danos</h3>\n<p>Danos ao imóvel, às instalações ou ao mobiliário que não forem sanados poderão ser <b>cobrados de mim</b>, nos termos das Instruções Gerais.</p>\n\n<h3>7. Responsabilidade</h3>\n<p>A entrega do PNR em condições adequadas e com todas as pendências sanadas é <b>responsabilidade exclusiva minha, militar ocupante</b>.</p>\n\n<h3>8. Comparecimento</h3>\n<p>Comprometo-me a comparecer no dia e na hora agendados. Se não puder, cancelo ou remarco pelo sistema com antecedência, para liberar o horário a outro companheiro.</p>\n`;\n\nconst CHECK = [\n  {\n    k: \"pintura\", t: \"Pintura em ordem\",\n    d: \"Paredes e tetos pintados, sem manchas, furos ou remendos aparentes.\",\n    como: \"<p>Compare com o <b>Termo de Recebimento</b> que você assinou ao ocupar o PNR: o imóvel deve voltar no mesmo estado.</p><ul><li>Tape e lixe os furos de quadros, suportes e prateleiras antes de pintar.</li><li>Use a mesma cor e o mesmo acabamento do original.</li><li>Não esqueça teto de banheiro e cozinha, onde costuma haver mofo.</li></ul>\"\n  },\n  {\n    k: \"faxina\", t: \"Faxina feita e imóvel vazio\",\n    d: \"Sem pertences pessoais nem entulho, inclusive área de serviço e armários.\",\n    como: \"<p>Vazio quer dizer vazio: nada de móveis velhos, tinta sobrando, caixas ou lixo na área de serviço.</p><ul><li>Limpe azulejos, rejunte, box, vidros e ralos.</li><li>Confira o alto dos armários e o quadro de luz.</li><li>Retire o entulho do condomínio, e não só do apartamento.</li></ul>\"\n  },\n  {\n    k: \"iluminacao\", t: \"Lâmpadas, tomadas e interruptores\",\n    d: \"Tudo instalado e funcionando, sem fio solto nem espelho faltando.\",\n    como: \"<p>Teste um por um, acendendo e apagando.</p><ul><li>Reponha todas as lâmpadas queimadas, inclusive de área de serviço, banheiro e varanda.</li><li>Recoloque os espelhos (tampas) de tomadas e interruptores.</li><li>Se tirou luminária ou chuveiro, devolva o item original ou um equivalente, sem deixar fio exposto.</li></ul>\"\n  },\n  {\n    k: \"ar\", t: \"Ar-condicionado instalado e Em Condições De uso\",\n    d: \"Instalado no lugar, ligando e gelando.\",\n    como: \"<p><b>Ligue o aparelho antes da vistoria</b> e confira se está gelando de verdade.</p><ul><li>Se você retirou o aparelho, ele deve estar reinstalado antes da entrega.</li><li>Limpe os filtros e confira se não há vazamento de água.</li><li>Deixe o controle junto das chaves.</li></ul>\"\n  },\n  {\n    k: \"hidro\", t: \"Louças, metais, portas e janelas\",\n    d: \"Sem vazamento, sem trinco quebrado, sem vidro rachado.\",\n    como: \"<ul><li>Abra todas as torneiras e dê descarga em todos os vasos.</li><li>Teste fechaduras, trincos e maçanetas de todas as portas e janelas.</li><li>Troque o que estiver rachado, entortado ou faltando.</li></ul>\"\n  },\n  {\n    k: \"light\", t: \"Light: conta encerrada e débitos quitados\",\n    d: \"Leve o comprovante no dia da vistoria.\",\n    como: \"<p>Peça o <b>encerramento da conta</b> (desligamento) no seu nome e quite o que estiver em aberto.</p><ol><li>Agência virtual da Light, aplicativo ou telefone.</li><li>Peça o encerramento informando a data da sua saída.</li><li>Pague a fatura final e <b>guarde o comprovante</b>, impresso ou no celular.</li></ol><p>Sem o comprovante, a vistoria pode ficar pendente.</p>\"\n  },\n  {\n    k: \"naturgy\", t: \"Naturgy: conta encerrada e débitos quitados\",\n    d: \"Leve o comprovante no dia da vistoria.\",\n    como: \"<p>Mesmo procedimento do gás.</p><ol><li>Atendimento da Naturgy, pelo aplicativo, site ou telefone.</li><li>Peça o encerramento da conta e quite as faturas.</li><li><b>Guarde o comprovante</b> para apresentar ao instrutor.</li></ol>\"\n  },\n  {\n    k: \"chaves\", t: \"Todas as chaves, cópias e controles\",\n    d: \"Inclusive as cópias que você fez, controles de portão e tags.\",\n    como: \"<p>Junte tudo num chaveiro só, para não esquecer nada no dia.</p><ul><li>Chaves de porta, portão, caixa de correio e área de serviço.</li><li><b>Todas as cópias</b> que você mandou fazer.</li><li>Controles de portão, tags de acesso e controle do ar-condicionado.</li></ul>\"\n  },\n  {\n    k: \"termo\", t: \"Termo de Recebimento em mãos\",\n    d: \"É o documento que você assinou ao ocupar o PNR, usado na comparação.\",\n    como: \"<p>A vistoria compara o estado atual com o <b>Termo de Recebimento/Vistoria Inicial</b>. Leve uma cópia para resolver qualquer dúvida na hora.</p><p>Não acha o seu? Procure a Seção de PNR antes da vistoria.</p>\"\n  }\n];\n\n/* Roteiro que o instrutor segue no celular, cômodo por cômodo, igual para todos. */\nconst ROTEIRO = [\n  { k: \"pintura\", t: \"Pintura de paredes e tetos\", d: \"Compare com o Termo de Recebimento. Sem furos, manchas ou remendo aparente.\" },\n  { k: \"faxina\", t: \"Faxina e imóvel vazio\", d: \"Sem pertences nem entulho, inclusive área de serviço e alto dos armários.\" },\n  { k: \"iluminacao\", t: \"Lâmpadas, tomadas e interruptores\", d: \"Acenda e apague tudo. Espelhos no lugar, sem fio exposto.\" },\n  { k: \"ar\", t: \"Ar-condicionado Em Condições De uso\", d: \"Ligue e confira se gela. Controle presente.\" },\n  { k: \"hidro\", t: \"Louças, metais e hidráulica\", d: \"Torneiras, descargas, vazamentos e ralos.\" },\n  { k: \"esquadrias\", t: \"Portas, janelas e fechaduras\", d: \"Trincos, maçanetas, vidros e persianas.\" },\n  { k: \"light\", t: \"Comprovante da Light\", d: \"Conta encerrada e débitos quitados.\" },\n  { k: \"naturgy\", t: \"Comprovante da Naturgy\", d: \"Conta encerrada e débitos quitados.\" },\n  { k: \"chaves\", t: \"Chaves, cópias e controles\", d: \"Confira a quantidade contra o termo inicial.\" },\n  { k: \"militar\", t: \"Entrega pessoal do responsável\", d: \"É o próprio militar, sem representante e sem família acompanhando.\" }\n];\n\nconst FAQ = [\n  [\"Minha esposa pode entregar as chaves para mim?\", \"Não. A entrega é pessoal, feita pelo militar responsável pelo PNR. Parente, cônjuge e companheiro de turma não são aceitos.\"],\n  [\"Posso agendar no mesmo dia do meu voo?\", \"Não. O sistema bloqueia o dia da viagem. Se aparecer alguma pendência, a chave não é recebida, e você não terá tempo de corrigir. Marque 2 ou 3 dias antes.\"],\n  [\"Minha família pode acompanhar a vistoria?\", \"Não. A vistoria é feita apenas com o militar responsável.\"],\n  [\"Esqueci de agendar e falta pouco tempo. E agora?\", \"Escolha o horário que você precisa e use o botão de pedido de encaixe. Escreva a justificativa. O seu instrutor decide se aprova.\"],\n  [\"E se o instrutor achar alguma pendência?\", \"A chave não é recebida naquele dia. Você recebe a lista exata do que corrigir e marca uma nova vistoria pelo mesmo link, já com prioridade.\"],\n  [\"Escolhi o instrutor errado. Perdi o pedido?\", \"Não. O instrutor transfere o seu pedido para o responsável certo, e você é avisado.\"],\n  [\"Preciso de senha ou cadastro?\", \"Não. Basta o link. Você se identifica pelo nome de guerra, e é por ele que o sistema encontra o seu agendamento depois.\"],\n  [\"Posso cancelar ou remarcar?\", \"Sim, pelo mesmo link, digitando o seu nome de guerra. Faça isso o quanto antes, para liberar o horário a outro companheiro.\"],\n  [\"O que preciso levar no dia?\", \"Todas as chaves, cópias e controles; os comprovantes da Light e da Naturgy; e uma cópia do Termo de Recebimento. E compareça sozinho.\"]\n];\n\nconst WD = [\"dom\", \"seg\", \"ter\", \"qua\", \"qui\", \"sex\", \"sáb\"];\nconst WD_CURTO = [\"D\", \"S\", \"T\", \"Q\", \"Q\", \"S\", \"S\"];\nconst MES_NOME = [\"Janeiro\", \"Fevereiro\", \"Março\", \"Abril\", \"Maio\", \"Junho\", \"Julho\", \"Agosto\", \"Setembro\", \"Outubro\", \"Novembro\", \"Dezembro\"];\nconst MES_CURTO = [\"Jan\", \"Fev\", \"Mar\", \"Abr\", \"Mai\", \"Jun\", \"Jul\", \"Ago\", \"Set\", \"Out\", \"Nov\", \"Dez\"];\n\n/* ===================== datas ===================== */\n\nconst dn = s => { const [y, m, d] = String(s).slice(0, 10).split(\"-\").map(Number); return Date.UTC(y, m - 1, d) / 864e5; };\nconst ds = n => new Date(n * 864e5).toISOString().slice(0, 10);\nconst wd = s => new Date(dn(s) * 864e5).getUTCDay();\nconst f2 = s => String(s).slice(8, 10) + \"/\" + String(s).slice(5, 7);\nconst f3 = s => f2(s) + \"/\" + String(s).slice(0, 4);\nconst fw = s => f2(s) + \" (\" + WD[wd(s)] + \")\";\nconst difDias = (a, b) => dn(b) - dn(a);\nconst mesDe = s => +String(s).slice(5, 7) - 1;\nconst anoDe = s => +String(s).slice(0, 4);\nconst esc = s => String(s ?? \"\").replace(/[&<>\"']/g, c => ({ \"&\": \"&amp;\", \"<\": \"&lt;\", \">\": \"&gt;\", '\"': \"&quot;\", \"'\": \"&#39;\" }[c]));\nconst horaDe = (data, hhmm) => new Date(String(data).slice(0, 10) + \"T\" + hhmm + \":00-03:00\");\n\n/* ===================== estado ===================== */\n\nconst S = {\n  papel: \"aluno\", eu: \"\", chave: \"\", tokAlunos: \"\", view: \"\", est: null, carregando: true,\n  step: 0, mes: null, dia: null,\n  dados: { nomeGuerra: \"\", instrutor: \"\", endereco: \"\", viagem: \"\", novaOM: \"\", telefone: \"\", email: \"\" },\n  aceites: {}, aceiteGeral: false, rolouTermo: false,\n  checklist: {}, abertos: {},\n  meu: null, erro: \"\", enviando: false,\n  abaInst: \"pedidos\", roteiro: {}, roteiroId: null,\n  mostrarComoFunciona: true\n};\n\nconst LS = \"pnr-esao\";\nfunction salvarLocal() {\n  try {\n    localStorage.setItem(LS, JSON.stringify({\n      dados: S.dados, checklist: S.checklist, aceites: S.aceites, aceiteGeral: S.aceiteGeral\n    }));\n  } catch (e) {}\n}\nfunction lerLocal() {\n  try {\n    const o = JSON.parse(localStorage.getItem(LS) || \"{}\");\n    if (o.dados) Object.assign(S.dados, o.dados);\n    if (o.checklist) S.checklist = o.checklist;\n    if (o.aceites) S.aceites = o.aceites;\n    S.aceiteGeral = !!o.aceiteGeral;\n  } catch (e) {}\n}\n\nconst conn = (cls, txt) => {\n  const el = document.getElementById(\"conn\");\n  el.className = \"conn \" + cls;\n  el.querySelector(\"span\").textContent = txt;\n};\nfunction banner(html, cls) {\n  const b = document.getElementById(\"banner\");\n  if (!html) { b.hidden = true; b.innerHTML = \"\"; return; }\n  b.hidden = false;\n  b.innerHTML = '<div class=\"' + (cls || \"\") + '\">' + html + \"</div>\";\n}\nfunction toast(msg) {\n  const old = document.getElementById(\"toast\");\n  if (old) old.remove();\n  const d = document.createElement(\"div\");\n  d.id = \"toast\"; d.className = \"toast\"; d.setAttribute(\"role\", \"status\"); d.textContent = msg;\n  document.body.appendChild(d);\n  setTimeout(() => d.remove(), 4200);\n}\n\n/* ===================== base de dados ===================== */\n\nlet API = null;\n\n/** Google Apps Script: as funções do Code.gs. */\nfunction gasAPI() {\n  const call = (fn, ...args) => new Promise((res, rej) =>\n    google.script.run.withSuccessHandler(res).withFailureHandler(e => rej(new Error(e && e.message ? e.message : String(e))))[fn](...args));\n  return { kind: \"gas\", call };\n}\n\n/** Demonstração: arquivo aberto direto no navegador. Os dados ficam só neste aparelho. */\nfunction demoAPI() {\n  const K = \"pnr-esao-demo\";\n  const ler = () => { try { return JSON.parse(localStorage.getItem(K) || \"{}\"); } catch (e) { return {}; } };\n  const gravar = o => { try { localStorage.setItem(K, JSON.stringify(o)); } catch (e) {} };\n  const INSTS = [\n    [\"Maj Modesto\", 10, \"Até 04/12\"], [\"Maj Andrews\", 10, \"Até 18/12\"], [\"Maj Pimenta\", 1, \"09/11 a 13/11\"],\n    [\"Maj Calderaro\", 10, \"Até 19/11\"], [\"Maj Cesse\", 11, \"Até 04/12\"], [\"Cap Dantas\", 10, \"13/11 a 26/11\"],\n    [\"Cap Loan\", 11, \"Até 04/12\"], [\"Cap Santana\", 11, \"Até 11/12\"], [\"Cap Uerlei Moreira\", 11, \"Até 18/12\"],\n    [\"Cap Cavalcanti\", 11, \"Até 29/12\"], [\"Cap Marcello\", 11, \"Até 19/11; 30/11 a 22/12\"]\n  ];\n  const hoje = new Date().toISOString().slice(0, 10);\n  const cfg = { inicio: hoje > \"2026-10-21\" ? hoje : \"2026-10-21\", fim: \"2026-12-30\", horasMinimas: 72 };\n  const janelas = texto => {\n    const data = s => {\n      const m = String(s).match(/(\\d{1,2})\\s*\\/\\s*(\\d{1,2})/);\n      if (!m) return null;\n      const d = (\"0\" + m[1]).slice(-2), mes = (\"0\" + m[2]).slice(-2);\n      return (+mes >= 10 ? \"2026\" : \"2027\") + \"-\" + mes + \"-\" + d;\n    };\n    const out = [];\n    String(texto).split(/;|\\be\\b(?![^\\d]*a\\b)/i).forEach(p => {\n      p = p.trim(); if (!p) return;\n      const fx = p.match(/(\\d{1,2}\\s*\\/\\s*\\d{1,2})\\s*(?:a|até|-)\\s*(\\d{1,2}\\s*\\/\\s*\\d{1,2})/i);\n      if (fx) { out.push({ ini: data(fx[1]), fim: data(fx[2]) }); return; }\n      if (/at[eé]/i.test(p)) { const f = data(p); if (f) out.push({ ini: cfg.inicio, fim: f }); return; }\n      const u = data(p); if (u) out.push({ ini: u, fim: u });\n    });\n    return out.filter(r => r.ini && r.fim && r.ini <= r.fim)\n      .map(r => ({ ini: r.ini < cfg.inicio ? cfg.inicio : r.ini, fim: r.fim > cfg.fim ? cfg.fim : r.fim }))\n      .filter(r => r.ini <= r.fim);\n  };\n  const base = () => {\n    const db = ler();\n    return {\n      papel: db.papel || \"aluno\", eu: db.eu || \"\", hoje, agora: new Date().toISOString(), termoVersao: \"1.0\",\n      cfg, tempos: TEMPOS_PADRAO,\n      feriados: {\n        \"2026-11-02\": { tipo: \"dia\", n: \"Finados\" },\n        \"2026-11-20\": { tipo: \"dia\", n: \"Dia Nacional de Zumbi e da Consciência Negra\" },\n        \"2026-12-24\": { tipo: \"tarde\", n: \"Véspera de Natal (ponto facultativo à tarde)\" },\n        \"2026-12-25\": { tipo: \"dia\", n: \"Natal\" }\n      },\n      instrutores: INSTS.map(i => ({ nome: i[0], dispTexto: i[2], janelas: janelas(i[2]), total: i[1], temEmail: false })),\n      bloqueios: db.bloqueios || [],\n      ags: db.ags || []\n    };\n  };\n  return {\n    kind: \"demo\",\n    async call(fn, ...args) {\n      const db = ler(); const b = base();\n      if (fn === \"estado\") {\n        const nome = args[1];\n        const out = Object.assign({}, b);\n        out.ocupados = b.ags.filter(a => [\"solicitado\", \"encaixe\", \"confirmado\"].includes(a.situacao))\n          .map(a => ({ instrutor: a.instrutor, data: a.data, tempo: a.tempo, nomeGuerra: a.nomeGuerra, situacao: a.situacao, id: a.id }));\n        out.agendamentos = b.ags;\n        if (nome) out.meus = b.ags.filter(a => normNome(a.nomeGuerra) === normNome(nome));\n        out.excecoes = []; out.excecoesMinhas = [];\n        delete out.ags;\n        return out;\n      }\n      if (fn === \"aceitarTermos\") return { ok: true, versao: \"1.0\" };\n      if (fn === \"solicitar\") {\n        const d = args[0];\n        const erro = motivo(b, d.instrutor, d.data, d.tempo, { viagem: d.viagem, excecao: !!d.encaixe }).txt;\n        if (erro) return { ok: false, erro };\n        const a = Object.assign({}, d, {\n          id: Math.random().toString(36).slice(2, 10),\n          situacao: d.encaixe ? \"encaixe\" : \"solicitado\",\n          criadoEm: new Date().toISOString(), atualizadoEm: new Date().toISOString(),\n          pendencias: [], historico: [{ quando: new Date().toISOString(), texto: \"Solicitada pelo aluno.\" }]\n        });\n        db.ags = (db.ags || []).concat([a]); gravar(db);\n        return { ok: true, agendamento: a };\n      }\n      const acha = id => (db.ags || []).find(a => a.id === id);\n      if (fn === \"cancelar\") {\n        const a = acha(args[0]); if (!a) return { ok: false, erro: \"Não encontrado.\" };\n        a.situacao = \"cancelado\"; a.atualizadoEm = new Date().toISOString(); gravar(db); return { ok: true, agendamento: a };\n      }\n      if (fn === \"confirmar\" || fn === \"recusar\" || fn === \"registrarResultado\" || fn === \"transferir\" || fn === \"remarcar\") {\n        const a = acha(args[1] !== undefined && fn !== \"remarcar\" ? args[1] : args[0]);\n        if (!a) return { ok: false, erro: \"Não encontrado.\" };\n        if (fn === \"confirmar\") a.situacao = \"confirmado\";\n        if (fn === \"recusar\") { a.situacao = \"recusado\"; a.motivo = args[2] || \"\"; }\n        if (fn === \"transferir\") { a.instrutor = args[2]; a.situacao = \"solicitado\"; }\n        if (fn === \"remarcar\") { a.data = args[2]; a.tempo = args[3]; a.situacao = \"solicitado\"; }\n        if (fn === \"registrarResultado\") {\n          if (args[2]) { a.situacao = \"entregue\"; a.pendencias = []; }\n          else { a.situacao = \"pendencias\"; a.pendencias = args[3] || []; }\n        }\n        a.atualizadoEm = new Date().toISOString(); gravar(db); return { ok: true, agendamento: a };\n      }\n      if (fn === \"bloquear\") {\n        db.bloqueios = (db.bloqueios || []).concat([{ instrutor: b.eu || \"TODOS\", data: args[1], tempo: args[2] || \"TODOS\", motivo: args[3] || \"\" }]);\n        gravar(db); return { ok: true };\n      }\n      if (fn === \"desbloquear\") {\n        db.bloqueios = (db.bloqueios || []).filter(x => !(x.data === args[1] && (args[2] === \"TODOS\" || x.tempo === args[2])));\n        gravar(db); return { ok: true };\n      }\n      if (fn === \"liberarExcecao\") return { ok: true, validaAte: cfg.fim };\n      if (fn === \"salvarEmail\") return { ok: true };\n      if (fn === \"agendarPeloInstrutor\") {\n        const d = args[1];\n        const a = Object.assign({}, d, { id: Math.random().toString(36).slice(2, 10), instrutor: b.eu, situacao: \"confirmado\", pendencias: [], criadoEm: new Date().toISOString(), atualizadoEm: new Date().toISOString() });\n        db.ags = (db.ags || []).concat([a]); gravar(db); return { ok: true, agendamento: a };\n      }\n      return { ok: false, erro: \"Função não disponível no modo demonstração.\" };\n    },\n    trocarPapel(papel, eu) { const db = ler(); db.papel = papel; db.eu = eu || \"\"; gravar(db); }\n  };\n}\n\nconst normNome = s => String(s || \"\").trim().toUpperCase()\n  .normalize(\"NFD\").replace(/[̀-ͯ]/g, \"\").replace(/\\s+/g, \" \");\n\nasync function carregar(silencioso) {\n  if (!API) return;\n  try {\n    const est = await API.call(\"estado\", S.chave, S.dados.nomeGuerra || \"\", S.tokAlunos);\n    S.est = est;\n    S.papel = est.papel || \"aluno\";\n    S.eu = est.eu || \"\";\n    if (!S.view) S.view = S.papel === \"painel\" ? \"painel\" : S.papel === \"instrutor\" ? \"instrutor\" : \"aluno\";\n    if (est.chaveInvalida) banner(\"<b>Essa chave não vale mais.</b> Abrindo a tela do aluno. Peça o link certo à coordenação.\", \"w\");\n    // Link do instrutor: o aluno já vem amarrado a ele e não escolhe na lista.\n    if (est.instrutorFixo) { S.dados.instrutor = est.instrutorFixo; salvarLocal(); }\n    if (est.tokenInvalido) banner(\"<b>Esse link não vale mais.</b> Peça o link atual ao seu instrutor. Por enquanto, escolha-o na lista.\", \"w\");\n    const vivos = (est.meus || []).filter(a => [\"solicitado\", \"encaixe\", \"confirmado\", \"pendencias\", \"recusado\"].includes(a.situacao));\n    S.meu = vivos.sort((a, b) => (b.atualizadoEm || \"\") > (a.atualizadoEm || \"\") ? 1 : -1)[0] || null;\n    if (!S.mes) S.mes = (est.cfg && est.cfg.inicio) ? String(est.cfg.inicio).slice(0, 7) : null;\n    S.carregando = false;\n    conn(\"on\", API.kind === \"demo\" ? \"Modo demonstração · dados só neste aparelho\"\n      : S.papel === \"painel\" ? \"Painel da coordenação\" : S.papel === \"instrutor\" ? \"Conectado · \" + S.eu : \"Conectado\");\n    render();\n  } catch (e) {\n    S.carregando = false;\n    if (!silencioso) { conn(\"off\", \"Sem conexão\"); banner(\"<b>Não foi possível falar com a base.</b> Verifique a internet e recarregue a página. \" + esc(e.message || \"\"), \"b\"); }\n    render();\n  }\n}\n\nasync function acao(fn, ...args) {\n  if (S.enviando) return null;\n  S.enviando = true; S.erro = \"\"; render();\n  try {\n    const r = await API.call(fn, ...args);\n    S.enviando = false;\n    if (r && r.ok === false) { S.erro = r.erro || \"Não foi possível concluir.\"; render(); return r; }\n    return r;\n  } catch (e) {\n    S.enviando = false;\n    S.erro = e.message || \"Falha na conexão. Tente de novo.\";\n    render();\n    return null;\n  }\n}\n\n/* ===================== regras de disponibilidade ===================== */\n\nconst tempos = () => (S.est && S.est.tempos && S.est.tempos.length ? S.est.tempos : TEMPOS_PADRAO);\nconst tempoDe = k => tempos().find(t => t.k === k) || null;\nconst instDe = nome => ((S.est && S.est.instrutores) || []).find(i => i.nome === nome) || null;\n\n/**\n * Por que este tempo não pode ser marcado. Devolve {cls, txt}; txt vazio = pode.\n * cls: livre | fds | feriado | fora | bloqueado | cheio | viagem | curto | passou | fim\n * Mesma ordem do backend, para a tela nunca prometer o que o servidor recusa.\n */\nfunction motivo(est, instrutor, data, tempoK, opts) {\n  opts = opts || {};\n  const cfg = est.cfg || {};\n  const t = (est.tempos || TEMPOS_PADRAO).find(x => x.k === tempoK);\n  if (!t) return { cls: \"nao\", txt: \"Tempo inválido.\" };\n  if (data < cfg.inicio) return { cls: \"passou\", txt: \"Essa data já passou.\" };\n  if (data > cfg.fim) return { cls: \"fim\", txt: \"A campanha termina em \" + f3(cfg.fim) + \".\" };\n\n  const d = wd(data);\n  if (d === 0 || d === 6) return { cls: \"fds\", txt: \"Fim de semana: não há vistoria de PNR.\" };\n\n  const fer = (est.feriados || {})[data];\n  if (fer && fer.tipo === \"dia\") return { cls: \"feriado\", txt: fer.n + \": feriado, sem vistoria.\" };\n  if (fer && fer.tipo === \"tarde\" && t.turno === \"tarde\") return { cls: \"feriado\", txt: fer.n + \": só há vistoria pela manhã.\" };\n\n  const inst = (est.instrutores || []).find(i => i.nome === instrutor);\n  if (!inst) return { cls: \"nao\", txt: \"Escolha um instrutor da lista.\" };\n  if (!(inst.janelas || []).some(j => data >= j.ini && data <= j.fim))\n    return { cls: \"fora\", txt: \"O \" + inst.nome + \" só recebe vistorias em: \" + inst.dispTexto + \".\" };\n\n  for (const b of (est.bloqueios || [])) {\n    if (b.data !== data) continue;\n    if (b.instrutor !== \"TODOS\" && b.instrutor !== instrutor) continue;\n    if (b.tempo !== \"TODOS\" && b.tempo !== tempoK) continue;\n    return { cls: \"bloqueado\", txt: \"O \" + inst.nome + \" bloqueou esse horário\" + (b.motivo ? \": \" + b.motivo : \".\") };\n  }\n\n  // A ocupação que o aluno recebe vem sem id; a exclusão só vale ao remarcar o próprio pedido.\n  const oc = (est.ocupados || []).find(o => o.instrutor === instrutor && o.data === data && o.tempo === tempoK\n    && !(opts.id && o.id === opts.id));\n  if (oc) return { cls: \"cheio\", txt: \"Horário já ocupado\" + (oc.nomeGuerra ? \": \" + oc.nomeGuerra : \".\") };\n\n  if (opts.viagem) {\n    if (data > opts.viagem) return { cls: \"viagem\", txt: \"É depois da sua viagem (\" + f3(opts.viagem) + \").\" };\n    if (data === opts.viagem) return { cls: \"viagem\", txt: \"É o dia da sua viagem. Se houver pendência, você não tem tempo de corrigir. Escolha 2 ou 3 dias antes.\" };\n  }\n\n  if (!opts.porInstrutor && !opts.excecao) {\n    const horas = (horaDe(data, t.ini).getTime() - Date.now()) / 36e5;\n    const min = cfg.horasMinimas == null ? 72 : cfg.horasMinimas;\n    if (horas < min) return { cls: \"curto\", txt: \"Faltam menos de \" + min + \" h. Só com pedido de encaixe ao instrutor.\" };\n  }\n  return { cls: \"livre\", txt: \"\" };\n}\n\n/** Os 6 tempos de um dia, já com o motivo de cada um. */\nfunction temposDoDia(data, instrutor, opts) {\n  return tempos().map(t => {\n    const m = motivo(S.est, instrutor, data, t.k, opts);\n    return { t, cls: m.cls, txt: m.txt, livre: !m.txt };\n  });\n}\n\n/** Resumo do dia para o calendário. */\nfunction resumoDia(data, instrutor, opts) {\n  const lst = temposDoDia(data, instrutor, opts);\n  const livres = lst.filter(x => x.livre).length;\n  if (livres) return { cls: \"livre\", n: livres, txt: livres + (livres > 1 ? \" tempos livres\" : \" tempo livre\") };\n  const curto = lst.filter(x => x.cls === \"curto\").length;\n  if (curto) return { cls: \"curto\", n: curto, txt: curto + \" tempo(s) só com encaixe\" };\n  const cheio = lst.filter(x => x.cls === \"cheio\").length;\n  if (cheio === lst.length) return { cls: \"cheio\", n: 0, txt: \"Dia cheio\" };\n  return { cls: \"nao\", n: 0, txt: lst[0] ? lst[0].txt : \"Indisponível\" };\n}\n\n/** Meses com dias dentro do período da campanha. */\nfunction mesesDaCampanha() {\n  const cfg = (S.est && S.est.cfg) || {};\n  if (!cfg.inicio) return [];\n  const out = [];\n  let a = anoDe(cfg.inicio), m = mesDe(cfg.inicio);\n  const aF = anoDe(cfg.fim), mF = mesDe(cfg.fim);\n  while (a < aF || (a === aF && m <= mF)) {\n    out.push({ ym: a + \"-\" + (\"0\" + (m + 1)).slice(-2), nome: MES_NOME[m], curto: MES_CURTO[m] + \"/\" + String(a).slice(2) });\n    m++; if (m > 11) { m = 0; a++; }\n  }\n  return out;\n}\n\nfunction diasDoMes(ym) {\n  const [a, m] = ym.split(\"-\").map(Number);\n  const out = [];\n  const ult = new Date(Date.UTC(a, m, 0)).getUTCDate();\n  for (let d = 1; d <= ult; d++) out.push(ym + \"-\" + (\"0\" + d).slice(-2));\n  return out;\n}\n\n/* ===================== render: moldura ===================== */\n\nconst app = document.getElementById(\"app\");\n\nfunction render() {\n  const h1 = document.getElementById(\"h1\"), eb = document.getElementById(\"eyebrow\");\n  const vs = document.getElementById(\"views\");\n\n  if (S.papel === \"painel\") {\n    eb.textContent = \"EsAO · Devolução de PNR 2026\";\n    h1.textContent = \"Painel da coordenação\";\n    vs.hidden = false;\n    vs.innerHTML = [[\"painel\", \"Painel\"], [\"aluno\", \"Ver tela do aluno\"]]\n      .map(([v, t]) => '<button data-act=\"view\" data-v=\"' + v + '\" aria-pressed=\"' + (S.view === v) + '\">' + t + \"</button>\").join(\"\");\n  } else if (S.papel === \"instrutor\") {\n    eb.textContent = \"EsAO · Devolução de PNR 2026\";\n    h1.textContent = S.eu;\n    vs.hidden = false;\n    vs.innerHTML = [[\"pedidos\", \"Pedidos\"], [\"hoje\", \"Hoje\"], [\"agenda\", \"Agenda\"], [\"tools\", \"Ferramentas\"]]\n      .map(([v, t]) => '<button data-act=\"aba\" data-v=\"' + v + '\" aria-pressed=\"' + (S.abaInst === v) + '\">' + t + \"</button>\").join(\"\");\n  } else {\n    eb.textContent = \"EsAO · Devolução de PNR 2026\";\n    h1.textContent = \"Agendar a entrega do PNR\";\n    vs.hidden = true; vs.innerHTML = \"\";\n  }\n\n  if (S.carregando) { app.innerHTML = '<div class=\"wrap\"><p class=\"empty\">Carregando…</p></div>'; return; }\n  if (!S.est) { app.innerHTML = '<div class=\"wrap\"><div class=\"card\"><h2>Sem conexão com a base</h2><p>Recarregue a página. Se continuar, avise a coordenação.</p></div></div>'; return; }\n\n  if (S.view === \"painel\") app.innerHTML = renderPainel();\n  else if (S.view === \"instrutor\") app.innerHTML = renderInstrutor();\n  else app.innerHTML = renderAluno();\n\n  const el = app.querySelector(\"#termo-box\");\n  if (el) {\n    el.addEventListener(\"scroll\", () => {\n      if (el.scrollTop + el.clientHeight >= el.scrollHeight - 24 && !S.rolouTermo) { S.rolouTermo = true; render(); }\n    }, { passive: true });\n    if (el.scrollHeight <= el.clientHeight + 8) S.rolouTermo = true;\n  }\n}\n\nconst erroBox = () => S.erro ? '<div class=\"alerta b\"><b>Não deu para concluir</b><span>' + esc(S.erro) + \"</span></div>\" : \"\";\n\nconst guia = (onde, oque, depois) =>\n  '<div class=\"guide\"><b>' + esc(onde) + \"</b><p>\" + oque + \"</p>\"\n  + (depois ? '<p class=\"after\">O que acontece depois: ' + depois + \"</p>\" : \"\") + \"</div>\";\n\n/* ===================== render: aluno ===================== */\n\nfunction renderAluno() {\n  // O aluno já agendado pode reabrir o checklist e voltar ao acompanhamento.\n  if (S.meu && S.verCheck) return wrapA(passoChecklist());\n  if (S.meu) return wrapA(renderStatus());\n\n  const passos = '<nav class=\"steps\" aria-label=\"Etapas\">' + STEPS.map((s, i) =>\n    '<button data-act=\"step\" data-i=\"' + i + '\" class=\"' + (i === S.step ? \"on\" : i < S.step ? \"done\" : \"\")\n    + '\"' + (i > S.step ? \" disabled\" : \"\") + '><span class=\"bar\"></span><span class=\"lbl\">' + s + \"</span></button>\").join(\"\") + \"</nav>\";\n\n  const corpo = [passoComoFunciona, passoQuem, passoTermos, passoChecklist, passoPNR, passoHorario][S.step]();\n  return wrapA(passos + '<p class=\"stepn\">Passo ' + (S.step + 1) + \" de \" + STEPS.length + \"</p>\" + corpo);\n}\n\nfunction wrapA(html) {\n  return '<div class=\"wrap\">' + erroBox() + html + \"</div>\"\n    + '<button class=\"helpbtn\" data-act=\"ajuda\">? Preciso de ajuda</button>';\n}\n\n/* --- passo 0: como funciona --- */\nfunction passoComoFunciona() {\n  const cards = [\n    [\"1\", \"Prepare o PNR\", \"Pintado, faxinado, tudo funcionando, e as contas da Light e da Naturgy encerradas e quitadas.\"],\n    [\"2\", \"Agende aqui\", \"Escolha o seu instrutor e um horário com <b>72 h ou mais</b> de antecedência, e <b>2 ou 3 dias antes da sua viagem</b>.\"],\n    [\"3\", \"Aguarde a confirmação\", \"O seu instrutor confirma. Você vê a resposta nesta mesma tela.\"],\n    [\"4\", \"Compareça sozinho\", \"Com todas as chaves e os comprovantes. A entrega é pessoal: sem parente, sem companheiro de turma e sem família acompanhando.\"]\n  ];\n  return guia(\"Como funciona a entrega do PNR\",\n    \"Leia os 4 passos abaixo. Leva menos de um minuto e evita quase todo problema.\",\n    \"você informa o seu nome de guerra e escolhe o seu instrutor.\")\n    + '<div class=\"stack\">' + cards.map(c =>\n      '<div class=\"card\"><div style=\"display:grid;grid-template-columns:34px 1fr;gap:12px;align-items:start\">'\n      + '<div style=\"font-family:var(--f-display);font-size:30px;font-weight:700;color:var(--accent);line-height:1\">' + c[0] + \"</div>\"\n      + \"<div><strong>\" + c[1] + \"</strong><p style=\\\"margin:4px 0 0;color:var(--muted)\\\">\" + c[2] + \"</p></div></div></div>\").join(\"\")\n    + '<div class=\"alerta w\" style=\"margin-top:12px\"><b>O erro mais comum</b><span>Marcar a vistoria para o mesmo dia do voo. Se o instrutor achar qualquer pendência, a chave não é recebida, e não há tempo de corrigir. Deixe 2 ou 3 dias de folga.</span></div>'\n    + '<div class=\"navbar\"><div class=\"navbar-in\"><button class=\"btn pri big\" data-act=\"go\" data-i=\"1\">Começar o agendamento</button></div></div>';\n}\n\n/* --- passo 1: quem é você --- */\nfunction passoQuem() {\n  const insts = (S.est.instrutores || []);\n  const agPor = {};\n  (S.est.ocupados || []).forEach(o => { agPor[o.instrutor] = (agPor[o.instrutor] || 0) + 1; });\n  const nome = S.dados.nomeGuerra;\n  const fixo = S.est.instrutorFixo;\n\n  return guia(\"Quem é você\",\n    fixo\n      ? \"Digite o seu <b>nome de guerra</b>. O instrutor já vem definido pelo link que ele lhe mandou.\"\n      : \"Digite o seu <b>nome de guerra</b> e toque no <b>seu instrutor responsável</b>. Não há senha nem cadastro.\",\n    \"você lê e aceita os termos da entrega.\")\n    + '<div class=\"card\"><div class=\"field\" style=\"margin-top:0\">'\n    + \"<label for=\\\"ng\\\">Seu nome de guerra</label>\"\n    + '<small>Escreva como você é tratado na OM, por exemplo \"Souza\" ou \"Lima Filho\". É por ele que você encontra o seu agendamento depois.</small>'\n    + '<input id=\"ng\" type=\"text\" data-bind=\"nomeGuerra\" value=\"' + esc(nome) + '\" autocomplete=\"off\" autocapitalize=\"words\" placeholder=\"Ex.: Souza\">'\n    + \"</div></div>\"\n    + (fixo\n      ? '<div class=\"card\" style=\"margin-top:14px;border-color:var(--accent)\">'\n      + '<span class=\"pill acc\">Seu instrutor</span>'\n      + '<div style=\"font-family:var(--f-display);font-size:24px;font-weight:700;margin-top:6px\">' + esc(fixo) + \"</div>\"\n      + '<p style=\"margin:4px 0 0;color:var(--muted);font-size:14px\">Recebe vistorias: <b>'\n      + esc((instDe(fixo) || {}).dispTexto || \"período completo\") + \"</b>.<br>\"\n      + \"Você abriu o link do \" + esc(fixo) + \", então a sua vistoria será com ele. Se ele não for o seu instrutor responsável, peça o link certo ao seu orientador.</p></div>\"\n      : '<h2 class=\"q\" style=\"margin-top:18px\">Seu instrutor responsável</h2>'\n      + '<p class=\"hint\">Se escolher o instrutor errado, não há problema: ele transfere o seu pedido para o responsável certo e você é avisado.</p>'\n      + '<div class=\"opts\">' + insts.map(i => {\n      const n = agPor[i.nome] || 0;\n      return '<button class=\"opt\" role=\"radio\" data-act=\"inst\" data-n=\"' + esc(i.nome) + '\" aria-checked=\"' + (S.dados.instrutor === i.nome) + '\">'\n        + '<span class=\"radio\" aria-hidden=\"true\"></span><strong>' + esc(i.nome) + \"</strong>\"\n        + '<span class=\"sub\">Recebe vistorias: ' + esc(i.dispTexto || \"período completo\") + \"</span>\"\n        + '<span class=\"tags\"><span class=\"pill pend\">' + i.total + \" alunos</span>\"\n        + (n ? '<span class=\"pill acc\">' + n + \" já agendado(s)</span>\" : \"\") + \"</span></button>\";\n    }).join(\"\") + \"</div>\")\n    + nav(0, S.dados.nomeGuerra.trim().length >= 3 && S.dados.instrutor ? 2 : null,\n      S.dados.nomeGuerra.trim().length < 3 ? \"Digite o seu nome de guerra\" : !S.dados.instrutor ? \"Escolha o seu instrutor\" : \"Ler os termos\");\n}\n\n/* --- passo 2: termos --- */\nfunction passoTermos() {\n  const faltam = CRITICOS.filter(c => !S.aceites[c.k]).length + (S.aceiteGeral ? 0 : 1);\n  const podeAssinar = S.rolouTermo && faltam === 0;\n  const assinou = normNome(S.dados.assinatura || \"\") === normNome(S.dados.nomeGuerra) && S.dados.nomeGuerra.trim().length >= 3;\n\n  return guia(\"Termos da entrega do PNR\",\n    \"Leia o termo até o fim, marque os 3 pontos críticos e assine com o seu nome de guerra.\",\n    \"você confere o checklist de preparação do imóvel.\")\n    + '<div class=\"termo\" id=\"termo-box\" tabindex=\"0\">' + TERMO_HTML + \"</div>\"\n    + '<div class=\"scrollmark' + (S.rolouTermo ? \" done\" : \"\") + '\">' + (S.rolouTermo ? \"✓ Você leu o termo até o fim\" : \"↓ Role o texto até o fim para continuar\") + \"</div>\"\n    + '<h2 class=\"q\" style=\"margin-top:18px\">Os 3 pontos que mais geram problema</h2>'\n    + '<p class=\"hint\">Marque um por um. São os que mais causam transtorno todos os anos.</p>'\n    + '<div class=\"stack\">' + CRITICOS.map(c =>\n      '<button class=\"chk crit\" role=\"checkbox\" data-act=\"crit\" data-k=\"' + c.k + '\" aria-checked=\"' + !!S.aceites[c.k] + '\">'\n      + '<span class=\"box\" aria-hidden=\"true\">✓</span><strong>' + esc(c.t) + \"</strong>\"\n      + '<span class=\"sub\">' + esc(c.d) + \"</span></button>\").join(\"\")\n    + '<button class=\"chk\" role=\"checkbox\" data-act=\"geral\" aria-checked=\"' + S.aceiteGeral + '\">'\n    + '<span class=\"box\" aria-hidden=\"true\">✓</span><strong>Li e aceito todo o Termo de Ciência e Compromisso</strong>'\n    + '<span class=\"sub\">Inclui pintura, faxina, iluminação, equipamentos Em Condições De uso, quitação da Light e da Naturgy, chaves e responsabilidade por danos.</span></button>'\n    + \"</div>\"\n    + '<div class=\"card\" style=\"margin-top:14px\"><div class=\"field\" style=\"margin-top:0\">'\n    + '<label for=\"assin\">Assine digitando o seu nome de guerra</label>'\n    + \"<small>Digite igual ao que você informou: <b>\" + esc(S.dados.nomeGuerra) + \"</b>. Isso fica registrado com a data e a hora.</small>\"\n    + '<input id=\"assin\" type=\"text\" data-bind=\"assinatura\" value=\"' + esc(S.dados.assinatura || \"\") + '\" autocomplete=\"off\" '\n    + (podeAssinar ? \"\" : \"disabled\") + ' placeholder=\"' + esc(S.dados.nomeGuerra) + '\">'\n    + (podeAssinar ? \"\" : '<small style=\"color:var(--warn)\">Role o termo até o fim e marque os 4 itens acima para liberar a assinatura.</small>')\n    + \"</div></div>\"\n    + nav(1, assinou ? 3 : null, !S.rolouTermo ? \"Leia o termo até o fim\" : faltam ? \"Marque os \" + faltam + \" item(ns) que faltam\" : !assinou ? \"Assine com o seu nome de guerra\" : \"Aceitar e continuar\");\n}\n\n/* --- passo 3: checklist --- */\nfunction passoChecklist() {\n  const feitos = CHECK.filter(c => S.checklist[c.k]).length;\n  const pct = Math.round(feitos / CHECK.length * 100);\n  return guia(\"Checklist de preparação\",\n    \"Marque o que já está pronto. Toque em <b>Como fazer</b> em qualquer item para ver o passo a passo.\",\n    \"você informa o endereço do PNR e a data da sua viagem.\")\n    + \"<p><b>\" + feitos + \" de \" + CHECK.length + \"</b> itens prontos</p>\"\n    + '<div class=\"prog\"><i style=\"width:' + pct + '%\"></i></div>'\n    + (feitos < CHECK.length\n      ? '<div class=\"alerta w\"><b>Pode agendar mesmo assim</b><span>O checklist não bloqueia o agendamento. Mas no dia da vistoria tudo precisa estar pronto, ou a chave não é recebida.</span></div>'\n      : '<div class=\"alerta w\" style=\"background:var(--ok-soft)\"><b>Tudo pronto</b><span>Ótimo. Guarde os comprovantes da Light e da Naturgy para apresentar no dia.</span></div>')\n    + '<div class=\"stack\">' + CHECK.map(c =>\n      '<div class=\"chk\" role=\"group\" style=\"cursor:default\">'\n      + '<button class=\"box\" data-act=\"check\" data-k=\"' + c.k + '\" aria-pressed=\"' + !!S.checklist[c.k] + '\" aria-label=\"Marcar: ' + esc(c.t) + '\" '\n      + 'style=\"' + (S.checklist[c.k] ? \"border-color:var(--accent);background:var(--accent);color:var(--accent-ink)\" : \"\") + '\">✓</button>'\n      + \"<strong>\" + esc(c.t) + \"</strong>\"\n      + '<span class=\"sub\">' + esc(c.d) + \"</span>\"\n      + '<button class=\"more\" data-act=\"howto\" data-k=\"' + c.k + '\">' + (S.abertos[c.k] ? \"Ocultar\" : \"Como fazer\") + \"</button>\"\n      + (S.abertos[c.k] ? '<div class=\"howto\">' + c.como + \"</div>\" : \"\")\n      + \"</div>\").join(\"\") + \"</div>\"\n    + (S.verCheck\n      ? '<div class=\"navbar\"><div class=\"navbar-in\"><button class=\"btn pri big\" data-act=\"voltarStatus\">Voltar ao meu agendamento</button></div></div>'\n      : nav(2, 4, \"Continuar\"));\n}\n\n/* --- passo 4: dados do PNR --- */\nfunction passoPNR() {\n  const d = S.dados;\n  const temEnd = (d.endereco || \"\").trim().length >= 8;\n  const temVia = !!d.viagem;\n  let avisoViagem = \"\";\n  if (temVia) {\n    const dias = difDias(S.est.hoje, d.viagem);\n    if (dias < 0) avisoViagem = '<div class=\"alerta b\"><b>Data no passado</b><span>Confira a data da viagem.</span></div>';\n    else if (dias <= 1) avisoViagem = '<div class=\"alerta b\"><b>Prazo apertado</b><span>Sua viagem é em ' + (dias === 0 ? \"hoje\" : \"1 dia\") + '. Praticamente não há folga para corrigir pendências. Procure o seu instrutor agora e peça um encaixe.</span></div>';\n    else if (dias <= 3) avisoViagem = '<div class=\"alerta w\"><b>Atenção ao prazo</b><span>Faltam ' + dias + \" dias para a viagem. Marque a vistoria o quanto antes e deixe pelo menos 1 dia de folga.</span></div>\";\n  }\n  return guia(\"Seu PNR e sua viagem\",\n    \"Estes dados vão para o instrutor: ele precisa saber <b>onde</b> é a vistoria e <b>quando</b> você viaja.\",\n    \"você escolhe o dia e o horário da vistoria.\")\n    + '<div class=\"card\"><div class=\"field\" style=\"margin-top:0\">'\n    + '<label for=\"end\">Endereço do PNR</label>'\n    + \"<small>Rua ou avenida, número, bloco e apartamento. Ex.: Av. Gen. Benedito da Silveira, bloco 11, apto 204, Vila Militar.</small>\"\n    + '<textarea id=\"end\" data-bind=\"endereco\" placeholder=\"Endereço completo, com bloco e apartamento\">' + esc(d.endereco) + \"</textarea>\"\n    + \"</div>\"\n    + '<div class=\"field\"><label for=\"via\">Data da sua viagem (ou apresentação na nova OM)</label>'\n    + \"<small>O sistema usa esta data para impedir que você marque a vistoria no dia do embarque.</small>\"\n    + '<input id=\"via\" type=\"date\" data-bind=\"viagem\" value=\"' + esc(d.viagem) + '\" min=\"' + esc(S.est.hoje) + '\"></div>'\n    + '<div class=\"field\"><label for=\"om\">Nova OM</label><small>Opcional. Ajuda a coordenação a organizar as prioridades.</small>'\n    + '<input id=\"om\" type=\"text\" data-bind=\"novaOM\" value=\"' + esc(d.novaOM) + '\" placeholder=\"Ex.: 1º BIS - Manaus/AM\"></div>'\n    + '<div class=\"field\"><label for=\"tel\">Seu WhatsApp</label><small>Opcional. Para o instrutor falar com você se precisar remarcar.</small>'\n    + '<input id=\"tel\" type=\"tel\" data-bind=\"telefone\" value=\"' + esc(d.telefone) + '\" placeholder=\"(21) 9.....\"></div>'\n    + '<div class=\"field\"><label for=\"mail\">Seu e-mail</label><small>Opcional. Se informar, você recebe a confirmação e os avisos por e-mail.</small>'\n    + '<input id=\"mail\" type=\"email\" data-bind=\"email\" value=\"' + esc(d.email) + '\" placeholder=\"nome@exemplo.com\"></div>'\n    + \"</div>\" + avisoViagem\n    + nav(3, temEnd && temVia ? 5 : null, !temEnd ? \"Informe o endereço do PNR\" : !temVia ? \"Informe a data da viagem\" : \"Escolher o horário\");\n}\n\n/* --- passo 5: escolher horário --- */\nfunction passoHorario() {\n  const inst = instDe(S.dados.instrutor);\n  if (!inst) return '<div class=\"alerta b\"><b>Escolha um instrutor</b><span>Volte ao passo 2.</span></div>' + nav(1, null, \"\");\n  const opts = { viagem: S.dados.viagem };\n  const meses = mesesDaCampanha();\n  if (!S.mes || !meses.some(m => m.ym === S.mes)) S.mes = meses[0] ? meses[0].ym : null;\n\n  const dias = S.mes ? diasDoMes(S.mes) : [];\n  const primeiro = dias.length ? wd(dias[0]) : 0;\n  const cel = [];\n  for (let i = 0; i < primeiro; i++) cel.push('<span class=\"day blank\"></span>');\n  let totalLivre = 0;\n  dias.forEach(d => {\n    const r = resumoDia(d, inst.nome, opts);\n    if (r.cls === \"livre\") totalLivre += r.n;\n    const dias3 = S.dados.viagem ? difDias(d, S.dados.viagem) : null;\n    const ideal = dias3 !== null && dias3 >= 2 && dias3 <= 3 && r.cls === \"livre\";\n    cel.push('<button class=\"day ' + r.cls + (S.dia === d ? \" sel\" : \"\") + (ideal ? \" ideal\" : \"\")\n      + '\" data-act=\"dia\" data-d=\"' + d + '\" aria-label=\"' + f3(d) + \": \" + esc(r.txt) + '\">'\n      + '<span class=\"dn\">' + (+d.slice(8, 10)) + \"</span>\"\n      + (r.n ? '<span class=\"vg\">' + r.n + \"</span>\" : \"\") + \"</button>\");\n  });\n\n  const livresNaJanela = inst.janelas.map(j => f3(j.ini) + \" a \" + f3(j.fim)).join(\" e \");\n\n  let detalhe = \"\";\n  if (S.dia) {\n    const lst = temposDoDia(S.dia, inst.nome, opts);\n    const dias3 = S.dados.viagem ? difDias(S.dia, S.dados.viagem) : null;\n    let aviso = \"\";\n    if (dias3 === 1) aviso = '<div class=\"alerta w\"><b>Só 1 dia antes da viagem</b><span>Dá para marcar, mas se houver pendência o prazo é curtíssimo. Se puder, escolha 2 ou 3 dias antes.</span></div>';\n    else if (dias3 !== null && dias3 >= 2 && dias3 <= 3) aviso = '<div class=\"alerta w\" style=\"background:var(--ok-soft)\"><b>Prazo ideal</b><span>' + dias3 + \" dias antes da sua viagem. Se aparecer pendência, você tem tempo de corrigir.</span></div>\";\n    const bloco = turno => '<p class=\"turno-h\">' + (turno === \"manha\" ? \"Manhã\" : \"Tarde\") + \"</p>\"\n      + lst.filter(x => x.t.turno === turno).map(x =>\n        '<button class=\"slot ' + (x.livre ? \"livre\" : \"off\") + '\" data-act=\"' + (x.livre ? \"pedir\" : x.cls === \"curto\" ? \"encaixe\" : \"porque\") + '\" data-k=\"' + x.t.k + '\">'\n        + '<span class=\"h\">' + x.t.ini + \" às \" + x.t.fim + \"</span>\"\n        + (x.livre ? '<span class=\"pill ok\">Livre · tocar para pedir</span>'\n          : x.cls === \"curto\" ? '<span class=\"pill warn\">⚡ Pedir encaixe</span>'\n            : '<span class=\"pill pend\">Indisponível</span>')\n        + (x.txt ? '<span class=\"why\">' + esc(x.txt) + \"</span>\" : '<span class=\"why\">Vistoria de 1 h 15. Toque para pedir este horário ao ' + esc(inst.nome) + \".</span>\")\n        + \"</button>\").join(\"\");\n    detalhe = '<div class=\"sep\"></div><h2 class=\"q\">' + fw(S.dia) + \"</h2>\" + aviso\n      + '<div class=\"slots\">' + bloco(\"manha\") + bloco(\"tarde\") + \"</div>\";\n  }\n\n  return guia(\"Escolher o dia e o horário\",\n    \"Toque num <b>dia verde</b> e depois no horário. Dias cinza têm o motivo explicado ao toque.\",\n    \"o \" + esc(inst.nome) + \" recebe o pedido e confirma. Você acompanha nesta tela.\")\n    + '<div class=\"card\" style=\"margin-bottom:12px\"><strong>' + esc(inst.nome) + \"</strong>\"\n    + '<p style=\"margin:4px 0 0;color:var(--muted);font-size:14px\">Recebe vistorias em ' + esc(livresNaJanela || inst.dispTexto) + \". \"\n    + (totalLivre ? \"<b>\" + totalLivre + \" tempo(s) livre(s)</b> neste mês.\" : \"Nenhum tempo livre neste mês: veja os outros meses.\") + \"</p>\"\n    + (S.est.instrutorFixo ? \"\" : '<button class=\"btn sm\" data-act=\"go\" data-i=\"1\" style=\"margin-top:10px\">Trocar de instrutor</button>') + \"</div>\"\n    + '<div class=\"mtabs\" role=\"group\">' + meses.map(m =>\n      '<button data-act=\"mes\" data-v=\"' + m.ym + '\" aria-pressed=\"' + (S.mes === m.ym) + '\">' + m.curto + \"</button>\").join(\"\") + \"</div>\"\n    + '<div class=\"wk\">' + WD_CURTO.map((w, i) => '<span>' + w + \"</span>\").join(\"\") + \"</div>\"\n    + '<div class=\"grid\">' + cel.join(\"\") + \"</div>\"\n    + '<div class=\"legend\">'\n    + '<span><i style=\"background:var(--ok-soft);border-color:var(--ok)\"></i> livre</span>'\n    + '<span><i style=\"background:var(--warn-soft);border-color:var(--warn)\"></i> só com encaixe</span>'\n    + '<span><i style=\"background:var(--pend-soft)\"></i> cheio</span>'\n    + '<span><i style=\"background:var(--surface-2)\"></i> sem vistoria</span>'\n    + '<span><i style=\"background:var(--ok);border-radius:50%;width:8px;height:8px\"></i> prazo ideal (2 a 3 dias antes da viagem)</span>'\n    + \"</div>\" + detalhe\n    + nav(4, null, \"\");\n}\n\nfunction nav(voltar, proximo, rotulo) {\n  const b = voltar !== null && voltar !== undefined\n    ? '<button class=\"btn\" data-act=\"go\" data-i=\"' + voltar + '\">Voltar</button>' : \"\";\n  const p = rotulo\n    ? '<button class=\"btn pri\" data-act=\"go\" data-i=\"' + (proximo === null ? \"\" : proximo) + '\"' + (proximo === null ? \" disabled\" : \"\") + \">\" + esc(rotulo) + \"</button>\"\n    : \"\";\n  if (!b && !p) return \"\";\n  return '<div class=\"navbar\"><div class=\"navbar-in\">' + b + p + \"</div></div>\";\n}\n\n/* --- tela de acompanhamento do aluno --- */\nfunction renderStatus() {\n  const a = S.meu, t = tempoDe(a.tempo) || { ini: \"\", fim: \"\" };\n  const faltamCheck = CHECK.filter(c => !S.checklist[c.k]);\n  const dias = difDias(S.est.hoje, a.data);\n  const diasViagem = a.viagem ? difDias(S.est.hoje, a.viagem) : null;\n\n  const cab = {\n    solicitado: [\"warnb\", \"⏳ Aguardando confirmação\", \"O \" + a.instrutor + \" recebeu o seu pedido. Enquanto ele não confirmar, o horário fica reservado para você.\"],\n    encaixe: [\"warnb\", \"⚡ Encaixe em análise\", \"Você pediu um horário com menos de 72 h. O \" + a.instrutor + \" vai decidir se aprova.\"],\n    confirmado: [\"okb\", \"✅ Vistoria CONFIRMADA\", \"Está marcado. Compareça no dia e na hora, pessoalmente e sozinho.\"],\n    recusado: [\"badb\", \"❌ Pedido recusado\", \"O \" + a.instrutor + \" não pôde atender nesse horário.\" + (a.motivo ? \" Motivo: \" + a.motivo : \"\") + (a.sugestao ? \" Sugestão dele: \" + a.sugestao : \"\")],\n    pendencias: [\"badb\", \"⚠️ Vistoria com pendências\", \"A chave não foi recebida. Corrija os itens abaixo e marque uma nova vistoria.\"],\n    entregue: [\"okb\", \"🎉 Chaves recebidas\", \"PNR entregue e processo encerrado. Bom trabalho e boa viagem.\"]\n  }[a.situacao] || [\"\", a.situacao, \"\"];\n\n  let acoes = \"\";\n  if (a.situacao === \"confirmado\" || a.situacao === \"solicitado\" || a.situacao === \"encaixe\") {\n    acoes = '<div class=\"row-btns\" style=\"margin-top:12px\">'\n      + '<button class=\"btn pri\" data-act=\"ics\">📅 Pôr na agenda do celular</button>'\n      + '<button class=\"btn danger\" data-act=\"cancelar\">Cancelar</button></div>';\n  } else if (a.situacao === \"pendencias\" || a.situacao === \"recusado\") {\n    acoes = '<div class=\"row-btns\" style=\"margin-top:12px\">'\n      + '<button class=\"btn pri big\" data-act=\"remarcar\">' + (a.situacao === \"pendencias\" ? \"Já corrigi, marcar nova vistoria\" : \"Escolher outro horário\") + \"</button></div>\";\n  }\n\n  const leve = a.situacao === \"confirmado\" ? '<div class=\"leve\"><b>No dia, leve e faça:</b><ul>'\n    + \"<li><b>Todas as chaves e cópias</b>, controles de portão e do ar-condicionado</li>\"\n    + \"<li><b>Comprovantes da Light e da Naturgy</b>, com a conta encerrada e quitada</li>\"\n    + \"<li>Cópia do <b>Termo de Recebimento</b> do PNR</li>\"\n    + \"<li>Compareça <b>pessoalmente</b> e <b>sozinho</b>: a família não acompanha</li>\"\n    + \"<li>Ligue o <b>ar-condicionado</b> antes, para confirmar que está gelando</li>\"\n    + \"</ul></div>\" : \"\";\n\n  const pend = (a.pendencias && a.pendencias.length)\n    ? '<div class=\"alerta b\" style=\"margin-top:12px\"><b>Corrija estes itens antes da nova vistoria</b><span><ul style=\"margin:4px 0 0;padding-left:20px\">'\n    + a.pendencias.map(p => \"<li>\" + esc(p) + \"</li>\").join(\"\") + \"</ul></span></div>\" : \"\";\n\n  const alertaCheck = (a.situacao === \"confirmado\" && faltamCheck.length)\n    ? '<div class=\"alerta w\" style=\"margin-top:12px\"><b>' + faltamCheck.length + \" item(ns) do checklist ainda sem marcar</b><span><ul style=\\\"margin:4px 0 0;padding-left:20px\\\">\"\n    + faltamCheck.slice(0, 5).map(c => \"<li>\" + esc(c.t) + \"</li>\").join(\"\")\n    + \"</ul>Se não estiver pronto no dia, a chave não é recebida.</span>\"\n    + '<button class=\"btn sm\" data-act=\"verCheck\">Abrir o checklist</button></div>' : \"\";\n\n  const conta = (a.situacao === \"confirmado\" && dias >= 0) ? '<div class=\"countdown\">'\n    + \"<div><b>\" + dias + \"</b><span>dias para a vistoria</span></div>\"\n    + (diasViagem !== null ? \"<div><b>\" + diasViagem + \"</b><span>dias para a viagem</span></div>\" : \"\")\n    + \"<div><b>\" + faltamCheck.length + \"</b><span>itens a concluir</span></div></div>\" : \"\";\n\n  return guia(\"Seu agendamento\",\n    \"Esta é a sua tela de acompanhamento. Guarde o link: é por aqui que você vê a resposta do instrutor.\",\n    \"\")\n    + '<div class=\"next ' + cab[0] + '\">'\n    + '<div class=\"big\">' + cab[1] + \"</div>\"\n    + '<div class=\"when\">' + fw(a.data) + \" · \" + t.ini + \" às \" + t.fim + \"</div>\"\n    + \"<p style=\\\"margin:6px 0 0;color:var(--muted)\\\">\" + esc(cab[2]) + \"</p>\"\n    + \"<dl><dt>Instrutor</dt><dd>\" + esc(a.instrutor) + \"</dd>\"\n    + \"<dt>Aluno</dt><dd>\" + esc(a.nomeGuerra) + \"</dd>\"\n    + (a.endereco ? \"<dt>PNR</dt><dd>\" + esc(a.endereco) + \"</dd>\" : \"\")\n    + (a.viagem ? \"<dt>Sua viagem</dt><dd>\" + f3(a.viagem) + \"</dd>\" : \"\")\n    + \"</dl>\" + conta + leve + pend + alertaCheck + acoes\n    + \"</div>\"\n    + '<div class=\"row-btns\" style=\"margin-top:14px\">'\n    + '<button class=\"btn\" data-act=\"ajuda\">Dúvidas frequentes</button>'\n    + '<button class=\"btn\" data-act=\"zap\">💬 Falar com o instrutor</button>'\n    + '<button class=\"btn\" data-act=\"recarregar\">Atualizar</button></div>';\n}\n\n/* ===================== render: instrutor ===================== */\n\nfunction renderInstrutor() {\n  const ags = S.est.agendamentos || [];\n  const pend = ags.filter(a => a.situacao === \"solicitado\" || a.situacao === \"encaixe\")\n    .sort((a, b) => (a.situacao === \"encaixe\" ? -1 : 1) - (b.situacao === \"encaixe\" ? -1 : 1) || (a.criadoEm > b.criadoEm ? 1 : -1));\n  const conf = ags.filter(a => a.situacao === \"confirmado\").sort((a, b) => a.data + a.tempo > b.data + b.tempo ? 1 : -1);\n  const inst = instDe(S.eu) || { total: 0, dispTexto: \"\", janelas: [] };\n  const feitos = ags.filter(a => a.situacao === \"entregue\").length;\n  const comPend = ags.filter(a => a.situacao === \"pendencias\").length;\n\n  const primeiroAcesso = !localStorage.getItem(\"pnr-inst-ok\");\n  if (primeiroAcesso) return '<div class=\"dash\"><div class=\"sec full\">' + renderOnboardingInstrutor(inst) + \"</div></div>\";\n\n  const hojeL = conf.filter(a => a.data === S.est.hoje);\n  const ultimoDia = inst.janelas.length ? inst.janelas[inst.janelas.length - 1].fim : S.est.cfg.fim;\n  const diasRest = difDias(S.est.hoje, ultimoDia);\n  const faltam = Math.max(0, (inst.total || 0) - (feitos + conf.length + pend.length));\n\n  const topo = '<div class=\"sec full\">' + guia(S.abaInst === \"pedidos\" ? \"Pedidos aguardando o senhor\"\n    : S.abaInst === \"hoje\" ? \"Vistorias de hoje\" : S.abaInst === \"agenda\" ? \"Sua agenda completa\" : \"Ferramentas\",\n    S.abaInst === \"pedidos\" ? \"Confirme ou recuse cada pedido. Os marcados com <b>⚡</b> são encaixes, com menos de 72 h.\"\n      : S.abaInst === \"hoje\" ? \"Em ordem de horário, com endereço e roteiro de vistoria.\"\n        : S.abaInst === \"agenda\" ? \"Tudo o que está marcado, e o resultado do que já foi vistoriado.\"\n          : \"Agendar por um aluno, liberar exceção, bloquear horário e cadastrar o seu e-mail.\",\n    S.abaInst === \"pedidos\" ? \"o aluno vê a sua resposta na tela dele, na hora.\" : \"\")\n    + '<div class=\"kpis\">'\n    + '<div class=\"kpi warn\"><div class=\"v\">' + pend.length + '</div><div class=\"l\">pedidos aguardando</div></div>'\n    + '<div class=\"kpi\"><div class=\"v\">' + conf.length + '</div><div class=\"l\">confirmadas</div></div>'\n    + '<div class=\"kpi ok\"><div class=\"v\">' + feitos + \"/\" + (inst.total || 0) + '</div><div class=\"l\">chaves recebidas</div></div>'\n    + '<div class=\"kpi bad\"><div class=\"v\">' + comPend + '</div><div class=\"l\">com pendências</div></div>'\n    + \"</div>\"\n    + (diasRest >= 0 && diasRest <= 10 && faltam > 0\n      ? '<div class=\"alerta w\" style=\"margin-top:12px\"><b>Sua disponibilidade termina em ' + diasRest + \" dia(s), em \" + f3(ultimoDia) + \"</b>\"\n      + \"<span>Pelo total da escala, cerca de <b>\" + faltam + \" aluno(s)</b> seus ainda não agendaram. Cobre a turma: sem agendamento, eles ficam sem instrutor.</span></div>\"\n      : \"\")\n    + \"</div>\";\n\n  let corpo = \"\";\n  if (S.abaInst === \"pedidos\") {\n    corpo = '<div class=\"sec full\"><div class=\"sec-h\"><h2>Pedidos</h2><p>' + pend.length + \" aguardando</p></div>\"\n      + (pend.length ? '<div class=\"stack\">' + pend.map(cardPedido).join(\"\") + \"</div>\"\n        : '<p class=\"empty\">Nenhum pedido aguardando. Quando um aluno solicitar, aparece aqui e o senhor recebe um e-mail, se tiver cadastrado.</p>')\n      + \"</div>\"\n      // Sem nenhum movimento ainda, o primeiro passo do instrutor é repassar o link à turma.\n      + (ags.length ? \"\" : cardLinkAlunos(inst));\n  } else if (S.abaInst === \"hoje\") {\n    corpo = '<div class=\"sec full\"><div class=\"sec-h\"><h2>Hoje · ' + f3(S.est.hoje) + \"</h2><p>\" + hojeL.length + \" vistoria(s)</p></div>\"\n      + (hojeL.length ? '<div class=\"stack\">' + hojeL.map(a => cardVistoria(a, true)).join(\"\") + \"</div>\"\n        : '<p class=\"empty\">Nada marcado para hoje.</p>')\n      + \"</div>\";\n  } else if (S.abaInst === \"agenda\") {\n    const prox = conf.filter(a => a.data >= S.est.hoje);\n    const hist = ags.filter(a => [\"entregue\", \"pendencias\", \"recusado\", \"cancelado\"].includes(a.situacao))\n      .sort((a, b) => (b.atualizadoEm || \"\") > (a.atualizadoEm || \"\") ? 1 : -1);\n    corpo = '<div class=\"sec full\"><div class=\"sec-h\"><h2>Próximas vistorias</h2><p>' + prox.length + \"</p></div>\"\n      + (prox.length ? '<div class=\"stack\">' + prox.map(a => cardVistoria(a, false)).join(\"\") + \"</div>\" : '<p class=\"empty\">Nenhuma vistoria marcada.</p>')\n      + \"</div>\"\n      + '<div class=\"sec full\"><div class=\"sec-h\"><h2>Já resolvidas</h2><p>' + hist.length + \"</p></div>\"\n      + (hist.length ? '<ul class=\"evl\">' + hist.map(a =>\n        \"<li><b>\" + esc(a.nomeGuerra) + \"</b> · \" + f3(a.data) + \" · \" + situacaoTxt(a.situacao)\n        + (a.pendencias && a.pendencias.length ? \"<small>Pendências: \" + esc(a.pendencias.join(\"; \")) + \"</small>\" : \"\")\n        + \"</li>\").join(\"\") + \"</ul>\" : '<p class=\"empty\">Nada ainda.</p>')\n      + \"</div>\";\n  } else {\n    corpo = renderFerramentas(inst);\n  }\n  return '<div class=\"dash\">' + erroBox() + topo + corpo + \"</div>\";\n}\n\nfunction renderOnboardingInstrutor(inst) {\n  return guia(\"Bem-vindo, \" + S.eu, \"Três telas rápidas sobre como o sistema funciona para o senhor.\", \"\")\n    + '<div class=\"stack\">'\n    + '<div class=\"card\"><strong>1. Primeiro, mande o link à sua turma</strong><p style=\"margin:4px 0 0;color:var(--muted)\">O senhor tem um <b>link próprio dos seus orientados</b>. Quem abrir esse link só consegue agendar com o senhor. Ele está em <b>Ferramentas</b>, com um botão para mandar no WhatsApp. Depois disso, o senhor <b>confirma ou recusa</b> cada pedido e registra o resultado da vistoria.</p></div>'\n    + '<div class=\"card\"><strong>2. Sua disponibilidade</strong><p style=\"margin:4px 0 0;color:var(--muted)\">O sistema só oferece aos alunos os dias em que o senhor recebe vistorias: <b>' + esc(inst.dispTexto || \"período completo\") + \"</b>. O senhor tem <b>\" + (inst.total || 0) + \" alunos</b> na escala. Se a disponibilidade estiver errada, avise a coordenação.</p></div>\"\n    + '<div class=\"card\"><strong>3. Lembretes</strong><p style=\"margin:4px 0 0;color:var(--muted)\">Se o senhor cadastrar um e-mail em <b>Ferramentas</b>, recebe um aviso a cada pedido novo, o convite para pôr a vistoria na agenda do celular, e um resumo às 18 h da véspera e às 7 h do dia.</p></div>'\n    + '<div class=\"card\"><strong>Regras que o sistema já aplica sozinho</strong><ul style=\"margin:6px 0 0;padding-left:20px;color:var(--muted)\"><li>Fins de semana, 02/11, 20/11, 25/12 e a tarde de 24/12 ficam bloqueados.</li><li>O aluno não consegue marcar no dia da viagem dele.</li><li>Abaixo de 72 h, só com o seu aval, pelo pedido de encaixe.</li><li>Nada é oferecido fora da sua disponibilidade.</li></ul></div>'\n    + \"</div>\" + cardLinkAlunos(inst)\n    + '<button class=\"btn pri big\" data-act=\"instOk\" style=\"margin-top:14px\">Entendi, abrir meus pedidos</button>';\n}\n\nfunction situacaoTxt(s) {\n  return ({ solicitado: \"Aguardando\", encaixe: \"⚡ Encaixe\", confirmado: \"Confirmada\", recusado: \"Recusada\", cancelado: \"Cancelada\", entregue: \"✅ Chaves recebidas\", pendencias: \"⚠️ Com pendências\" })[s] || s;\n}\n\nfunction cardPedido(a) {\n  const t = tempoDe(a.tempo) || {};\n  const urg = a.situacao === \"encaixe\";\n  const horas = Math.round((new Date(a.criadoEm || Date.now()).getTime() - Date.now()) / -36e5);\n  const diasViagem = a.viagem ? difDias(a.data, a.viagem) : null;\n  return '<div class=\"ped' + (urg ? \" urg\" : \"\") + '\">'\n    + \"<h3>\" + (urg ? \"⚡ ENCAIXE · \" : \"\") + esc(a.nomeGuerra) + \"</h3>\"\n    + '<div class=\"todo\">' + (urg\n      ? \"O aluno pede um horário com menos de 72 h. Aprove só se o senhor puder atender.\"\n      : \"Confirme ou recuse. Enquanto o senhor não responder, o horário fica reservado.\") + \"</div>\"\n    + \"<dl><dt>Quando</dt><dd><b>\" + fw(a.data) + \" · \" + t.ini + \" às \" + t.fim + \"</b></dd>\"\n    + \"<dt>PNR</dt><dd>\" + esc(a.endereco || \"não informado\") + \"</dd>\"\n    + (a.viagem ? \"<dt>Viagem</dt><dd>\" + f3(a.viagem) + (diasViagem !== null ? \" · \" + (diasViagem <= 0 ? \"⚠️ mesmo dia ou depois\" : diasViagem === 1 ? \"⚠️ só 1 dia de folga\" : diasViagem + \" dias de folga\") : \"\") + \"</dd>\" : \"\")\n    + (a.novaOM ? \"<dt>Nova OM</dt><dd>\" + esc(a.novaOM) + \"</dd>\" : \"\")\n    + (a.telefone ? \"<dt>WhatsApp</dt><dd>\" + esc(a.telefone) + \"</dd>\" : \"\")\n    + (a.justificativa ? \"<dt>Justificativa</dt><dd>\" + esc(a.justificativa) + \"</dd>\" : \"\")\n    + \"<dt>Pedido há</dt><dd>\" + (horas < 1 ? \"menos de 1 h\" : horas + \" h\") + \"</dd>\"\n    + \"</dl>\"\n    + checklistAluno(a)\n    + '<div class=\"acts\">'\n    + '<button class=\"btn pri\" data-act=\"confirmar\" data-id=\"' + a.id + '\">✅ Confirmar ' + f2(a.data) + \" \" + t.ini + \"</button>\"\n    + '<button class=\"btn\" data-act=\"recusar\" data-id=\"' + a.id + '\">Recusar / propor outro</button>'\n    + '<button class=\"btn\" data-act=\"transferir\" data-id=\"' + a.id + '\">Não é meu aluno</button>'\n    + \"</div></div>\";\n}\n\nfunction checklistAluno(a) {\n  const c = a.checklist || {};\n  const faltam = CHECK.filter(x => !c[x.k]);\n  if (!Object.keys(c).length) return \"\";\n  if (!faltam.length) return '<div class=\"todo\" style=\"background:var(--ok-soft)\">O aluno marcou todos os itens do checklist como prontos.</div>';\n  return '<div class=\"todo\" style=\"background:var(--warn-soft)\">O aluno ainda não marcou: ' + esc(faltam.map(x => x.t.toLowerCase()).join(\", \")) + \".</div>\";\n}\n\nfunction cardVistoria(a, hoje) {\n  const t = tempoDe(a.tempo) || {};\n  const mapa = \"https://www.google.com/maps/search/?api=1&query=\" + encodeURIComponent(a.endereco || \"\");\n  return '<div class=\"ped\">'\n    + \"<h3>\" + t.ini + \" às \" + t.fim + \" · \" + esc(a.nomeGuerra) + \"</h3>\"\n    + \"<dl><dt>Data</dt><dd>\" + fw(a.data) + \"</dd>\"\n    + \"<dt>PNR</dt><dd>\" + esc(a.endereco || \"não informado\") + \"</dd>\"\n    + (a.viagem ? \"<dt>Viagem do aluno</dt><dd>\" + f3(a.viagem) + \"</dd>\" : \"\")\n    + (a.telefone ? \"<dt>WhatsApp</dt><dd>\" + esc(a.telefone) + \"</dd>\" : \"\")\n    + \"</dl>\" + checklistAluno(a)\n    + '<div class=\"acts\">'\n    + '<button class=\"btn pri\" data-act=\"roteiro\" data-id=\"' + a.id + '\">Fazer a vistoria (roteiro)</button>'\n    + (a.endereco ? '<a class=\"btn\" href=\"' + esc(mapa) + '\" target=\"_blank\" rel=\"noopener\">Ver no mapa</a>' : \"\")\n    + (a.telefone ? '<a class=\"btn\" href=\"https://wa.me/55' + esc(String(a.telefone).replace(/\\D/g, \"\")) + '\" target=\"_blank\" rel=\"noopener\">💬 WhatsApp</a>' : \"\")\n    + '<button class=\"btn\" data-act=\"recusar\" data-id=\"' + a.id + '\">Cancelar vistoria</button>'\n    + \"</div></div>\";\n}\n\n/** O link que o instrutor repassa à sua turma: quem abrir só agenda com ele. */\nfunction cardLinkAlunos(inst) {\n  const tok = inst && inst.tokenAlunos;\n  if (!tok) return \"\";\n  const link = linkBase() + \"?i=\" + tok;\n  const msg = \"Senhores, para agendar a entrega do PNR comigo, usem este link: \" + link\n    + \"\\n\\nLeiam os termos com atenção. Antecedência mínima de 72 h, nunca no dia do voo. \"\n    + \"A entrega é pessoal, sem representante, e a família não acompanha a vistoria.\";\n  return '<div class=\"sec full\" style=\"border-color:var(--accent)\">'\n    + '<div class=\"sec-h\"><h2>Link dos seus orientados</h2><p>Mande este link para a sua turma</p></div>'\n    + '<p class=\"hint\">Quem abrir este link só consegue agendar <b>com o senhor</b>: não precisa escolher instrutor e não consegue cair na agenda de outro.</p>'\n    + '<div class=\"card\" style=\"word-break:break-all;font-size:13px\" id=\"linkalunos\">' + esc(link) + \"</div>\"\n    + '<div class=\"row-btns\" style=\"margin-top:10px\">'\n    + '<a class=\"btn pri\" href=\"https://wa.me/?text=' + encodeURIComponent(msg) + '\" target=\"_blank\" rel=\"noopener\">💬 Mandar no WhatsApp</a>'\n    + '<button class=\"btn\" data-act=\"copiarEl\" data-k=\"linkalunos\">Copiar o link</button>'\n    + \"</div>\"\n    + '<p class=\"hint\" style=\"margin:10px 0 0\">Não repasse o link da <b>sua</b> tela (o que a coordenação lhe mandou): nele o aluno veria os pedidos e poderia confirmar vistorias.</p>'\n    + \"</div>\";\n}\n\n/** Endereço publicado do sistema, sem os parâmetros da tela atual. */\nfunction linkBase() {\n  try {\n    if (S.est && S.est.urlBase) return S.est.urlBase;\n    const u = String(location.href).split(\"?\")[0].split(\"#\")[0];\n    return u;\n  } catch (e) { return \"\"; }\n}\n\nfunction renderFerramentas(inst) {\n  const insts = S.est.instrutores || [];\n  const bls = (S.est.bloqueios || []).filter(b => b.instrutor === S.eu).sort((a, b) => a.data > b.data ? 1 : -1);\n  return cardLinkAlunos(inst)\n    + '<div class=\"sec full\"><div class=\"sec-h\"><h2>Aluno esqueceu de agendar?</h2><p>Duas formas de resolver</p></div>'\n    + '<div class=\"stack\">'\n    + '<div class=\"card\"><strong>1. Marcar eu mesmo pelo aluno</strong>'\n    + '<p style=\"margin:4px 0 0;color:var(--muted);font-size:14px\">Para quem ligou ou falou pessoalmente. Já sai confirmado, sem a regra das 72 h.</p>'\n    + '<button class=\"btn pri\" data-act=\"agendarPor\" style=\"margin-top:10px\">Marcar por um aluno</button></div>'\n    + '<div class=\"card\"><strong>2. Liberar o aluno da regra das 72 h</strong>'\n    + '<p style=\"margin:4px 0 0;color:var(--muted);font-size:14px\">O aluno escolhe sozinho o horário que precisa, por alguns dias.</p>'\n    + '<button class=\"btn\" data-act=\"excecao\" style=\"margin-top:10px\">Liberar exceção</button></div>'\n    + \"</div></div>\"\n    + '<div class=\"sec full\"><div class=\"sec-h\"><h2>Bloquear horário</h2><p>Serviço, missão, compromisso</p></div>'\n    + \"<p class=\\\"hint\\\">O que o senhor bloquear deixa de aparecer para os alunos. Dias com vistoria marcada não podem ser bloqueados: recuse ou remarque antes.</p>\"\n    + '<button class=\"btn pri\" data-act=\"bloquear\">Bloquear um dia ou um tempo</button>'\n    + (bls.length ? '<div class=\"sep\"></div><ul class=\"evl\">' + bls.map(b =>\n      \"<li>\" + f3(b.data) + \" · \" + (b.tempo === \"TODOS\" ? \"dia inteiro\" : b.tempo) + (b.motivo ? \" · \" + esc(b.motivo) : \"\")\n      + ' <button class=\"btn sm\" data-act=\"desbloquear\" data-d=\"' + b.data + '\" data-k=\"' + b.tempo + '\">Liberar</button></li>').join(\"\") + \"</ul>\" : \"\")\n    + \"</div>\"\n    + '<div class=\"sec full\"><div class=\"sec-h\"><h2>Seu e-mail para lembretes</h2><p>' + (inst.temEmail ? \"cadastrado\" : \"não cadastrado\") + \"</p></div>\"\n    + '<p class=\"hint\">Com e-mail cadastrado o senhor recebe: aviso de pedido novo, o convite para pôr a vistoria na agenda do celular, e o resumo às 18 h da véspera e às 7 h do dia.</p>'\n    + '<div class=\"field\" style=\"margin-top:0\"><label for=\"eminst\">E-mail</label>'\n    + '<input id=\"eminst\" type=\"email\" placeholder=\"nome@exemplo.com\"></div>'\n    + '<button class=\"btn pri\" data-act=\"salvarEmail\" style=\"margin-top:10px\">Salvar e-mail</button>'\n    + \"</div>\";\n}\n\n/* ===================== render: painel ===================== */\n\nfunction renderPainel() {\n  const ags = S.est.agendamentos || [];\n  const insts = S.est.instrutores || [];\n  const totalAlunos = insts.reduce((s, i) => s + (i.total || 0), 0);\n  const pend = ags.filter(a => a.situacao === \"solicitado\" || a.situacao === \"encaixe\");\n  const conf = ags.filter(a => a.situacao === \"confirmado\");\n  const feitos = ags.filter(a => a.situacao === \"entregue\");\n  const comPend = ags.filter(a => a.situacao === \"pendencias\");\n  const vivosPorNome = {};\n  ags.forEach(a => { if ([\"solicitado\", \"encaixe\", \"confirmado\", \"entregue\", \"pendencias\"].includes(a.situacao)) vivosPorNome[normNome(a.nomeGuerra)] = a; });\n  const engajados = Object.keys(vivosPorNome).length;\n  const semAgendar = Math.max(0, totalAlunos - engajados);\n\n  // Por instrutor: quantos agendaram, dias restantes e risco.\n  const linhas = insts.map(i => {\n    const meus = ags.filter(a => a.instrutor === i.nome && [\"solicitado\", \"encaixe\", \"confirmado\", \"entregue\", \"pendencias\"].includes(a.situacao));\n    const ult = i.janelas && i.janelas.length ? i.janelas[i.janelas.length - 1].fim : S.est.cfg.fim;\n    const rest = difDias(S.est.hoje, ult);\n    const faltam = Math.max(0, (i.total || 0) - meus.length);\n    const risco = faltam > 0 && rest >= 0 ? Math.round(faltam / Math.max(1, rest + 1) * 10) / 10 : 0;\n    return { i, n: meus.length, faltam, ult, rest, risco, entregues: meus.filter(a => a.situacao === \"entregue\").length };\n  }).sort((a, b) => b.risco - a.risco || a.rest - b.rest);\n\n  const emRisco = linhas.filter(l => l.faltam > 0 && l.rest >= 0 && (l.rest <= 10 || l.risco >= 1.5));\n\n  const pendAntigos = pend.filter(a => (Date.now() - new Date(a.criadoEm || Date.now()).getTime()) > 24 * 36e5);\n\n  const dias = {};\n  conf.concat(feitos).forEach(a => { dias[a.data] = (dias[a.data] || 0) + 1; });\n  const maxDia = Math.max(1, ...Object.values(dias));\n\n  const mesesL = mesesDaCampanha();\n  if (!S.mes || !mesesL.some(m => m.ym === S.mes)) S.mes = mesesL[0] ? mesesL[0].ym : null;\n  const diasMes = S.mes ? diasDoMes(S.mes) : [];\n  const primeiro = diasMes.length ? wd(diasMes[0]) : 0;\n  const heatCels = [];\n  for (let k = 0; k < primeiro; k++) heatCels.push(\"<div style=\\\"background:transparent;border:0\\\"></div>\");\n  diasMes.forEach(d => {\n    const n = dias[d] || 0;\n    const h = Math.round(n / maxDia * 85);\n    heatCels.push('<div style=\"--h:' + h + '%\"' + (h > 55 ? ' class=\"hot\"' : \"\") + \">\"\n      + (+d.slice(8, 10)) + (n ? \"<b>\" + n + \"</b>\" : \"\") + \"</div>\");\n  });\n\n  return '<div class=\"dash\">' + erroBox()\n    + '<div class=\"sec full\">' + guia(\"Painel da coordenação\",\n      \"Acompanhe a campanha e cobre quem está atrasado. Cada alerta já vem com o que fazer.\",\n      \"\")\n    + '<div class=\"kpis\">'\n    + '<div class=\"kpi\"><div class=\"v\">' + totalAlunos + '</div><div class=\"l\">alunos na escala</div></div>'\n    + '<div class=\"kpi bad\"><div class=\"v\">' + semAgendar + '</div><div class=\"l\">ainda sem agendar</div></div>'\n    + '<div class=\"kpi warn\"><div class=\"v\">' + pend.length + '</div><div class=\"l\">pedidos aguardando instrutor</div></div>'\n    + '<div class=\"kpi ok\"><div class=\"v\">' + feitos.length + '</div><div class=\"l\">chaves recebidas</div></div>'\n    + \"</div>\"\n    + '<div class=\"kpis\" style=\"margin-top:8px\">'\n    + '<div class=\"kpi\"><div class=\"v\">' + conf.length + '</div><div class=\"l\">vistorias confirmadas</div></div>'\n    + '<div class=\"kpi bad\"><div class=\"v\">' + comPend.length + '</div><div class=\"l\">com pendências a corrigir</div></div>'\n    + '<div class=\"kpi warn\"><div class=\"v\">' + pendAntigos.length + '</div><div class=\"l\">pedidos há mais de 24 h</div></div>'\n    + '<div class=\"kpi pend\"><div class=\"v\">' + Math.round(feitos.length / Math.max(1, totalAlunos) * 100) + '%</div><div class=\"l\">da campanha concluída</div></div>'\n    + \"</div></div>\"\n\n    + '<div class=\"sec full\"><div class=\"sec-h\"><h2>O que precisa de ação agora</h2></div>'\n    + (emRisco.length || pendAntigos.length\n      ? emRisco.map(l =>\n        '<div class=\"alerta ' + (l.rest <= 5 ? \"b\" : \"w\") + '\"><b>' + esc(l.i.nome) + \": disponibilidade termina em \" + l.rest + \" dia(s), em \" + f3(l.ult) + \"</b>\"\n        + \"<span>\" + l.faltam + \" de \" + l.i.total + \" alunos dele ainda não agendaram. Cobre a turma hoje; depois dessa data ele não recebe mais vistorias.</span>\"\n        + '<div class=\"row-btns\"><button class=\"btn sm\" data-act=\"cobrar\" data-n=\"' + esc(l.i.nome) + '\">Mensagem pronta para cobrar</button></div></div>').join(\"\")\n        + (pendAntigos.length ? '<div class=\"alerta w\"><b>' + pendAntigos.length + \" pedido(s) sem resposta há mais de 24 h</b><span>\"\n          + esc(pendAntigos.map(a => a.nomeGuerra + \" → \" + a.instrutor).join(\"; \")) + \". Avise os instrutores.</span></div>\" : \"\")\n      : '<p class=\"empty\">Nada urgente. Todos os instrutores estão com folga de prazo e os pedidos estão sendo respondidos.</p>')\n    + \"</div>\"\n\n    + '<div class=\"sec full\"><div class=\"sec-h\"><h2>Por instrutor</h2><p>Agendados sobre o total da escala</p></div>'\n    + '<table class=\"tbl\"><thead><tr><th>Instrutor</th><th class=\"num\">Agendados</th><th>Progresso</th><th class=\"num\">Entregues</th><th>Disponível até</th><th class=\"num\">Dias</th></tr></thead><tbody>'\n    + linhas.map(l => \"<tr><td><b>\" + esc(l.i.nome) + \"</b><br><small style=\\\"color:var(--muted)\\\">\" + esc(l.i.dispTexto) + \"</small></td>\"\n      + '<td class=\"num\">' + l.n + \"/\" + l.i.total + \"</td>\"\n      + '<td><div class=\"bar2\"><i style=\"width:' + Math.round(l.n / Math.max(1, l.i.total) * 100) + '%\"></i></div></td>'\n      + '<td class=\"num\">' + l.entregues + \"</td>\"\n      + \"<td>\" + f3(l.ult) + \"</td>\"\n      + '<td class=\"num\">' + (l.rest < 0 ? \"encerrado\" : l.rest) + \"</td></tr>\").join(\"\")\n    + \"</tbody></table></div>\"\n\n    + '<div class=\"sec\"><div class=\"sec-h\"><h2>Ocupação por dia</h2><p>Vistorias marcadas e realizadas</p></div>'\n    + '<div class=\"mtabs\" role=\"group\">' + mesesL.map(m =>\n      '<button data-act=\"mes\" data-v=\"' + m.ym + '\" aria-pressed=\"' + (S.mes === m.ym) + '\">' + m.curto + \"</button>\").join(\"\") + \"</div>\"\n    + '<div class=\"wk\" style=\"margin-top:10px\">' + WD_CURTO.map(w => \"<span>\" + w + \"</span>\").join(\"\") + \"</div>\"\n    + '<div class=\"heat\">' + heatCels.join(\"\") + \"</div>\"\n    + '<p class=\"hint\" style=\"margin-top:10px\">Cada dia tem 6 tempos por instrutor: 3 de manhã e 3 de tarde.</p></div>'\n\n    + '<div class=\"sec full\"><div class=\"sec-h\"><h2>Links para reenviar</h2><p>Cada instrutor recebe só a linha dele</p></div>'\n    + '<p class=\"hint\">Mande a <b>tela do instrutor</b> a cada instrutor, individualmente. É ele quem repassa o <b>link dos orientados</b> à turma dele. Quem abre o link de um instrutor só consegue agendar com aquele instrutor.</p>'\n    + '<table class=\"tbl\"><thead><tr><th>Instrutor</th><th>Link dos orientados dele</th></tr></thead><tbody>'\n    + insts.map((i, k) => \"<tr><td><b>\" + esc(i.nome) + \"</b></td>\"\n      + '<td style=\"word-break:break-all;font-size:12px\"><span id=\"la' + k + '\">'\n      + esc(i.tokenAlunos ? linkBase() + \"?i=\" + i.tokenAlunos : \"—\") + \"</span>\"\n      + (i.tokenAlunos ? ' <button class=\"btn sm\" data-act=\"copiarEl\" data-k=\"la' + k + '\">Copiar</button>' : \"\")\n      + \"</td></tr>\").join(\"\")\n    + \"</tbody></table>\"\n    + '<p class=\"hint\" style=\"margin-top:10px\">Os links das <b>telas dos instrutores</b> ficam no menu <b>PNR › Ver links</b> da planilha, e não aqui, para não serem repassados por engano.</p>'\n    + \"</div>\"\n    + '<div class=\"sec\"><div class=\"sec-h\"><h2>Últimos movimentos</h2></div>'\n    + (ags.length ? '<ul class=\"evl\">' + ags.slice().sort((a, b) => (b.atualizadoEm || \"\") > (a.atualizadoEm || \"\") ? 1 : -1).slice(0, 14).map(a =>\n      \"<li><b>\" + esc(a.nomeGuerra) + \"</b> · \" + esc(a.instrutor) + \" · \" + f3(a.data)\n      + \"<small>\" + situacaoTxt(a.situacao) + (a.pendencias && a.pendencias.length ? \" · \" + esc(a.pendencias.join(\"; \")) : \"\") + \"</small></li>\").join(\"\") + \"</ul>\"\n      : '<p class=\"empty\">Nenhum agendamento ainda.</p>')\n    + \"</div></div>\";\n}\n\n/* ===================== modais ===================== */\n\nconst layer = document.getElementById(\"layer\");\nfunction fecharModal() { layer.innerHTML = \"\"; }\nfunction modal(titulo, corpo, acoes) {\n  layer.innerHTML = '<div class=\"modal\" data-act=\"fundo\"><div class=\"card\" role=\"dialog\" aria-modal=\"true\" aria-label=\"' + esc(titulo) + '\">'\n    + \"<h2>\" + esc(titulo) + \"</h2>\" + corpo\n    + '<div class=\"row-btns\" style=\"margin-top:14px\">' + acoes + '<button class=\"btn\" data-act=\"fechar\">Fechar</button></div>'\n    + \"</div></div>\";\n}\n\nfunction modalAjuda() {\n  const ondeEstou = S.view === \"painel\" ? \"Painel da coordenação\" : S.view === \"instrutor\" ? \"Tela do instrutor\" : STEPS[S.step] || \"Seu agendamento\";\n  modal(\"Ajuda · \" + ondeEstou,\n    '<p class=\"hint\">Dúvidas mais comuns. Se a sua não estiver aqui, fale com o seu instrutor.</p>'\n    + '<div class=\"faq\">' + FAQ.map(([q, a]) => \"<details><summary>\" + esc(q) + \"</summary><p>\" + esc(a) + \"</p></details>\").join(\"\") + \"</div>\"\n    + '<div class=\"sep\"></div><p style=\"font-size:14px\"><b>Resumo das regras:</b> vistoria de 1 h 15, em 6 tempos por dia (08h00, 09h30 e 11h00; 13h30, 15h00 e 16h30). Sem vistoria em fins de semana, em 02/11, 20/11 e 25/12, nem na tarde de 24/12. Antecedência mínima de 72 h, que o instrutor pode flexibilizar. Nunca no dia da viagem.</p>',\n    \"\");\n}\n\nfunction modalPorque(txt) {\n  modal(\"Esse horário não está disponível\", '<p>' + esc(txt) + \"</p>\"\n    + '<p class=\"hint\" style=\"margin-top:10px\">Escolha um horário verde no mesmo dia, ou toque em outro dia do calendário.</p>', \"\");\n}\n\nfunction modalEncaixe(k) {\n  const t = tempoDe(k) || {};\n  modal(\"Pedir encaixe ao \" + S.dados.instrutor,\n    '<div class=\"alerta w\"><b>Faltam menos de 72 h para ' + f3(S.dia) + \" às \" + t.ini + \"</b>\"\n    + \"<span>Normalmente o agendamento exige 72 h. Explique abaixo por que o senhor precisa deste horário. O instrutor decide se aprova.</span></div>\"\n    + '<div class=\"field\" style=\"margin-top:0\"><label for=\"just\">Justificativa</label>'\n    + \"<small>Seja direto: data do voo, apresentação na nova OM, motivo do atraso.</small>\"\n    + '<textarea id=\"just\" placeholder=\"Ex.: embarco dia 12 e só recebi a passagem ontem.\"></textarea></div>',\n    '<button class=\"btn pri\" data-act=\"enviarEncaixe\" data-k=\"' + k + '\">Enviar pedido de encaixe</button>');\n}\n\nfunction modalConfirmarPedido(k) {\n  const t = tempoDe(k) || {};\n  const d = S.dados;\n  const dias = d.viagem ? difDias(S.dia, d.viagem) : null;\n  modal(\"Confirme o seu pedido\",\n    '<div class=\"next\"><div class=\"when\">' + fw(S.dia) + \" · \" + t.ini + \" às \" + t.fim + \"</div>\"\n    + \"<dl><dt>Instrutor</dt><dd>\" + esc(d.instrutor) + \"</dd>\"\n    + \"<dt>Aluno</dt><dd>\" + esc(d.nomeGuerra) + \"</dd>\"\n    + \"<dt>PNR</dt><dd>\" + esc(d.endereco) + \"</dd>\"\n    + \"<dt>Sua viagem</dt><dd>\" + f3(d.viagem) + (dias !== null ? \" · \" + dias + \" dia(s) de folga\" : \"\") + \"</dd></dl></div>\"\n    + '<div class=\"leve\"><b>Ao confirmar, o senhor declara que:</b><ul>'\n    + \"<li>vai comparecer <b>pessoalmente</b>, sem representante;</li>\"\n    + \"<li>vai estar <b>sozinho</b>, sem família acompanhando;</li>\"\n    + \"<li>o PNR estará pintado, faxinado e tudo Em Condições De uso;</li>\"\n    + \"<li>levará as chaves e os comprovantes da Light e da Naturgy.</li></ul></div>\",\n    '<button class=\"btn pri\" data-act=\"enviarPedido\" data-k=\"' + k + '\">Pedir vistoria</button>');\n}\n\nfunction modalCancelar() {\n  modal(\"Cancelar a vistoria\",\n    '<div class=\"alerta w\"><b>Atenção</b><span>O horário será liberado para outro companheiro. Se precisar, você terá de marcar de novo, respeitando as 72 h.</span></div>'\n    + '<div class=\"field\" style=\"margin-top:0\"><label for=\"cnome\">Digite o seu nome de guerra para confirmar</label>'\n    + \"<small>Exatamente como no agendamento: <b>\" + esc(S.meu.nomeGuerra) + \"</b></small>\"\n    + '<input id=\"cnome\" type=\"text\" autocomplete=\"off\"></div>'\n    + '<div class=\"field\"><label for=\"cmot\">Motivo</label><small>Opcional, mas ajuda o instrutor.</small>'\n    + '<input id=\"cmot\" type=\"text\" placeholder=\"Ex.: mudou a data do voo\"></div>',\n    '<button class=\"btn danger\" data-act=\"enviarCancelar\">Cancelar a vistoria</button>');\n}\n\nfunction modalRecusar(id) {\n  const a = (S.est.agendamentos || []).find(x => x.id === id) || {};\n  modal(\"Recusar ou propor outro horário\",\n    \"<p>Pedido de <b>\" + esc(a.nomeGuerra) + \"</b> para \" + f3(a.data) + \".</p>\"\n    + '<div class=\"field\"><label for=\"rmot\">Motivo</label><small>O aluno vê este texto.</small>'\n    + '<input id=\"rmot\" type=\"text\" placeholder=\"Ex.: estarei de serviço\"></div>'\n    + '<div class=\"field\"><label for=\"rsug\">Sugestão de outro horário</label><small>Opcional. O aluno escolhe de novo no calendário.</small>'\n    + '<input id=\"rsug\" type=\"text\" placeholder=\"Ex.: dia 04/12 pela manhã\"></div>',\n    '<button class=\"btn danger\" data-act=\"enviarRecusar\" data-id=\"' + id + '\">Recusar</button>');\n}\n\nfunction modalTransferir(id) {\n  const a = (S.est.agendamentos || []).find(x => x.id === id) || {};\n  const insts = (S.est.instrutores || []).filter(i => i.nome !== S.eu);\n  modal(\"Transferir para o instrutor certo\",\n    \"<p>O <b>\" + esc(a.nomeGuerra) + \"</b> escolheu o senhor, mas não é seu aluno. Escolha o responsável correto; o pedido não se perde e o aluno é avisado.</p>\"\n    + '<div class=\"field\"><label for=\"tdest\">Instrutor responsável</label><select id=\"tdest\">'\n    + insts.map(i => '<option value=\"' + esc(i.nome) + '\">' + esc(i.nome) + \" · \" + esc(i.dispTexto) + \"</option>\").join(\"\")\n    + \"</select></div>\",\n    '<button class=\"btn pri\" data-act=\"enviarTransferir\" data-id=\"' + id + '\">Transferir</button>');\n}\n\nfunction modalRoteiro(id) {\n  const a = (S.est.agendamentos || []).find(x => x.id === id);\n  if (!a) return;\n  S.roteiroId = id;\n  const r = S.roteiro[id] || {};\n  const marcados = ROTEIRO.filter(x => r[x.k]).length;\n  const ruins = ROTEIRO.filter(x => r[x.k] === \"n\");\n  modal(\"Roteiro de vistoria · \" + a.nomeGuerra,\n    '<div class=\"guide\"><b>Como usar</b><p>Percorra o imóvel na ordem e marque <b>✓</b> ou <b>⚠</b> em cada item. O que ficar ⚠ vai direto para o aluno como pendência a corrigir.</p></div>'\n    + \"<p><b>\" + marcados + \" de \" + ROTEIRO.length + \"</b> itens avaliados\"\n    + (ruins.length ? \" · <b style=\\\"color:var(--bad)\\\">\" + ruins.length + \" com problema</b>\" : \"\") + \"</p>\"\n    + '<div class=\"rot\">' + ROTEIRO.map(x =>\n      '<div class=\"it\"><strong>' + esc(x.t) + \"</strong>\"\n      + '<button class=\"y\" data-act=\"rot\" data-k=\"' + x.k + '\" data-v=\"y\" aria-pressed=\"' + (r[x.k] === \"y\") + '\">✓</button>'\n      + '<button class=\"n\" data-act=\"rot\" data-k=\"' + x.k + '\" data-v=\"n\" aria-pressed=\"' + (r[x.k] === \"n\") + '\">⚠</button>'\n      + \"<small>\" + esc(x.d) + \"</small></div>\").join(\"\") + \"</div>\"\n    + '<div class=\"field\"><label for=\"robs\">Observações</label>'\n    + '<textarea id=\"robs\" placeholder=\"Opcional. Detalhe o que precisa ser corrigido.\"></textarea></div>'\n    + (ruins.length\n      ? '<div class=\"alerta b\"><b>Não receba a chave</b><span>Há ' + ruins.length + \" item(ns) com problema. Registre as pendências: o aluno recebe a lista e marca nova vistoria.</span></div>\"\n      : marcados === ROTEIRO.length\n        ? '<div class=\"alerta w\" style=\"background:var(--ok-soft)\"><b>Tudo em ordem</b><span>Pode receber as chaves e encerrar.</span></div>'\n        : '<div class=\"alerta w\"><b>Faltam itens</b><span>Avalie todos antes de decidir.</span></div>'),\n    ruins.length\n      ? '<button class=\"btn danger\" data-act=\"regPend\" data-id=\"' + id + '\">Registrar pendências</button>'\n      : '<button class=\"btn pri\" data-act=\"regOk\" data-id=\"' + id + '\">✅ Recebi as chaves</button>');\n}\n\nfunction modalBloquear() {\n  modal(\"Bloquear um dia ou um tempo\",\n    '<p class=\"hint\">O que o senhor bloquear deixa de aparecer para os alunos.</p>'\n    + '<div class=\"field\" style=\"margin-top:0\"><label for=\"bd\">Dia</label>'\n    + '<input id=\"bd\" type=\"date\" min=\"' + esc(S.est.cfg.inicio) + '\" max=\"' + esc(S.est.cfg.fim) + '\"></div>'\n    + '<div class=\"field\"><label for=\"bt\">O que bloquear</label><select id=\"bt\">'\n    + '<option value=\"TODOS\">O dia inteiro (6 tempos)</option>'\n    + tempos().map(t => '<option value=\"' + t.k + '\">Só ' + t.ini + \" às \" + t.fim + \"</option>\").join(\"\")\n    + \"</select></div>\"\n    + '<div class=\"field\"><label for=\"bm\">Motivo</label><small>Opcional. O aluno vê este texto ao tocar no dia.</small>'\n    + '<input id=\"bm\" type=\"text\" placeholder=\"Ex.: serviço de dia\"></div>',\n    '<button class=\"btn pri\" data-act=\"enviarBloqueio\">Bloquear</button>');\n}\n\nfunction modalExcecao() {\n  modal(\"Liberar aluno da regra das 72 h\",\n    '<p class=\"hint\">O aluno passa a poder escolher sozinho qualquer horário livre da sua agenda, inclusive de amanhã.</p>'\n    + '<div class=\"field\" style=\"margin-top:0\"><label for=\"xn\">Nome de guerra do aluno</label>'\n    + '<input id=\"xn\" type=\"text\" placeholder=\"Ex.: Souza\"></div>'\n    + '<div class=\"field\"><label for=\"xd\">Por quantos dias</label><select id=\"xd\">'\n    + [3, 5, 7, 10, 15].map(d => '<option value=\"' + d + '\"' + (d === 7 ? \" selected\" : \"\") + \">\" + d + \" dias</option>\").join(\"\")\n    + \"</select></div>\",\n    '<button class=\"btn pri\" data-act=\"enviarExcecao\">Liberar</button>');\n}\n\nfunction modalAgendarPor() {\n  modal(\"Marcar por um aluno\",\n    '<div class=\"alerta w\"><b>Use com cuidado</b><span>A vistoria já sai <b>confirmada</b> e sem a regra das 72 h. Confirme com o aluno antes.</span></div>'\n    + '<div class=\"field\" style=\"margin-top:0\"><label for=\"an\">Nome de guerra</label><input id=\"an\" type=\"text\" placeholder=\"Ex.: Souza\"></div>'\n    + '<div class=\"field\"><label for=\"ae\">Endereço do PNR</label><input id=\"ae\" type=\"text\" placeholder=\"Bloco e apartamento\"></div>'\n    + '<div class=\"field\"><label for=\"ad\">Dia</label><input id=\"ad\" type=\"date\" min=\"' + esc(S.est.cfg.inicio) + '\" max=\"' + esc(S.est.cfg.fim) + '\"></div>'\n    + '<div class=\"field\"><label for=\"at\">Tempo</label><select id=\"at\">'\n    + tempos().map(t => '<option value=\"' + t.k + '\">' + t.ini + \" às \" + t.fim + \"</option>\").join(\"\") + \"</select></div>\"\n    + '<div class=\"field\"><label for=\"av\">Data da viagem do aluno</label><input id=\"av\" type=\"date\"></div>',\n    '<button class=\"btn pri\" data-act=\"enviarAgendarPor\">Marcar e confirmar</button>');\n}\n\nfunction modalRemarcar() {\n  const a = S.meu;\n  const inst = instDe(a.instrutor);\n  const opts = { viagem: a.viagem, id: a.id };\n  const meses = mesesDaCampanha();\n  const ym = S.mes || (meses[0] && meses[0].ym);\n  const dias = diasDoMes(ym).filter(d => resumoDia(d, a.instrutor, opts).cls === \"livre\");\n  modal(a.situacao === \"pendencias\" ? \"Nova vistoria após corrigir\" : \"Escolher outro horário\",\n    (a.situacao === \"pendencias\"\n      ? '<div class=\"alerta w\"><b>Confirme antes de marcar</b><span>As pendências foram realmente corrigidas? Se o instrutor achar o mesmo problema, a chave segue retida.</span></div>'\n      : \"\")\n    + '<div class=\"field\" style=\"margin-top:0\"><label for=\"rn\">Seu nome de guerra</label>'\n    + '<input id=\"rn\" type=\"text\" autocomplete=\"off\" placeholder=\"' + esc(a.nomeGuerra) + '\"></div>'\n    + '<div class=\"field\"><label for=\"rd\">Dia</label><select id=\"rd\">'\n    + (dias.length ? dias.map(d => '<option value=\"' + d + '\">' + fw(d) + \"</option>\").join(\"\")\n      : '<option value=\"\">Sem dia livre neste mês</option>') + \"</select></div>\"\n    + '<div class=\"field\"><label for=\"rt\">Tempo</label><select id=\"rt\">'\n    + tempos().map(t => '<option value=\"' + t.k + '\">' + t.ini + \" às \" + t.fim + \"</option>\").join(\"\") + \"</select>\"\n    + \"<small>Se o tempo escolhido estiver ocupado, o sistema avisa e você escolhe outro.</small></div>\",\n    '<button class=\"btn pri\" data-act=\"enviarRemarcar\">Pedir nova vistoria</button>');\n}\n\nfunction modalCobrar(nome) {\n  const inst = instDe(nome) || {};\n  const txt = \"Senhores, a disponibilidade do \" + nome + \" para receber PNR vai até \" + f3(inst.janelas && inst.janelas.length ? inst.janelas[inst.janelas.length - 1].fim : S.est.cfg.fim)\n    + \". Quem ainda não agendou a vistoria, faça hoje pelo link. Lembrem-se: antecedência mínima de 72 h, nunca no dia do voo, entrega pessoal e sem acompanhante.\";\n  modal(\"Mensagem pronta para cobrar\",\n    '<p class=\"hint\">Copie e mande no grupo da turma.</p>'\n    + '<div class=\"card\" style=\"font-size:14px\" id=\"msgcobrar\">' + esc(txt) + \"</div>\",\n    '<a class=\"btn pri\" href=\"https://wa.me/?text=' + encodeURIComponent(txt) + '\" target=\"_blank\" rel=\"noopener\">Abrir no WhatsApp</a>'\n    + '<button class=\"btn\" data-act=\"copiar\">Copiar texto</button>');\n}\n\n/** Arquivo .ics para o aluno pôr a vistoria na agenda do celular. */\nfunction baixarIcs() {\n  const a = S.meu, t = tempoDe(a.tempo);\n  const z = d => d.toISOString().replace(/[-:]/g, \"\").replace(/\\.\\d{3}/, \"\");\n  const ics = [\n    \"BEGIN:VCALENDAR\", \"VERSION:2.0\", \"PRODID:-//EsAO//PNR//PT-BR\", \"BEGIN:VEVENT\",\n    \"UID:pnr-\" + a.id + \"@esao\", \"DTSTAMP:\" + z(new Date()),\n    \"DTSTART:\" + z(horaDe(a.data, t.ini)), \"DTEND:\" + z(horaDe(a.data, t.fim)),\n    \"SUMMARY:Entrega do PNR - vistoria com \" + a.instrutor,\n    \"LOCATION:\" + String(a.endereco || \"\").replace(/,/g, \"\\\\,\"),\n    \"DESCRIPTION:Levar TODAS as chaves e copias\\\\, controles\\\\, comprovantes da Light e da Naturgy e copia do Termo de Recebimento. Comparecer pessoalmente e sozinho.\",\n    \"BEGIN:VALARM\", \"TRIGGER:-PT2H\", \"ACTION:DISPLAY\", \"DESCRIPTION:Vistoria do PNR em 2 horas\", \"END:VALARM\",\n    \"BEGIN:VALARM\", \"TRIGGER:-P1D\", \"ACTION:DISPLAY\", \"DESCRIPTION:Vistoria do PNR amanha\", \"END:VALARM\",\n    \"END:VEVENT\", \"END:VCALENDAR\"\n  ].join(\"\\r\\n\");\n  try {\n    const url = URL.createObjectURL(new Blob([ics], { type: \"text/calendar\" }));\n    const link = document.createElement(\"a\");\n    link.href = url; link.download = \"vistoria-pnr.ics\";\n    document.body.appendChild(link); link.click(); link.remove();\n    setTimeout(() => URL.revokeObjectURL(url), 4000);\n    toast(\"Abra o arquivo baixado para pôr na sua agenda.\");\n  } catch (e) { toast(\"Não foi possível gerar o arquivo neste aparelho.\"); }\n}\n\n/* ===================== eventos ===================== */\n\ndocument.addEventListener(\"input\", e => {\n  const b = e.target.getAttribute && e.target.getAttribute(\"data-bind\");\n  if (!b) return;\n  S.dados[b] = e.target.value;\n  salvarLocal();\n  // Não re-renderiza: isso tiraria o foco do campo. O botão atualiza ao tocar.\n  if (b === \"nomeGuerra\" || b === \"assinatura\" || b === \"endereco\" || b === \"viagem\") {\n    const nb = document.querySelector(\".navbar .btn.pri\");\n    if (nb) nb.disabled = !podeAvancar();\n  }\n});\n\nfunction podeAvancar() {\n  if (S.step === 1) return S.dados.nomeGuerra.trim().length >= 3 && !!S.dados.instrutor;\n  if (S.step === 2) return S.rolouTermo && CRITICOS.every(c => S.aceites[c.k]) && S.aceiteGeral\n    && normNome(S.dados.assinatura || \"\") === normNome(S.dados.nomeGuerra);\n  if (S.step === 4) return (S.dados.endereco || \"\").trim().length >= 8 && !!S.dados.viagem;\n  return true;\n}\n\ndocument.addEventListener(\"click\", async e => {\n  const el = e.target.closest(\"[data-act]\");\n  if (!el) return;\n  const act = el.getAttribute(\"data-act\");\n  const v = el.getAttribute(\"data-v\"), k = el.getAttribute(\"data-k\"), id = el.getAttribute(\"data-id\");\n  const d = el.getAttribute(\"data-d\"), n = el.getAttribute(\"data-n\"), i = el.getAttribute(\"data-i\");\n\n  // fechar modal tocando no fundo\n  if (act === \"fundo\") { if (e.target === el) fecharModal(); return; }\n  if (act === \"fechar\") { fecharModal(); return; }\n\n  switch (act) {\n    case \"view\": S.view = v; S.erro = \"\"; render(); return;\n    case \"aba\": S.abaInst = v; S.erro = \"\"; render(); return;\n    case \"mes\": S.mes = v; render(); return;\n    case \"ajuda\": modalAjuda(); return;\n    case \"recarregar\": await carregar(); toast(\"Atualizado.\"); return;\n    case \"instOk\": try { localStorage.setItem(\"pnr-inst-ok\", \"1\"); } catch (x) {} render(); return;\n\n    case \"step\": case \"go\": {\n      const alvo = Number(i);\n      if (isNaN(alvo)) return;\n      if (alvo > S.step && !podeAvancar()) return;\n      if (S.step === 2 && alvo === 3) {\n        // registra o aceite no momento em que o aluno assina\n        await acao(\"aceitarTermos\", S.dados.nomeGuerra, S.dados.instrutor,\n          CRITICOS.map(c => c.t).concat([\"Termo completo\"]));\n      }\n      S.step = Math.max(0, Math.min(STEPS.length - 1, alvo));\n      S.erro = \"\"; window.scrollTo(0, 0); render(); return;\n    }\n    case \"inst\": S.dados.instrutor = n; S.dia = null; salvarLocal(); render(); return;\n    case \"crit\": S.aceites[k] = !S.aceites[k]; salvarLocal(); render(); return;\n    case \"geral\": S.aceiteGeral = !S.aceiteGeral; salvarLocal(); render(); return;\n    case \"check\": S.checklist[k] = !S.checklist[k]; salvarLocal(); render(); return;\n    case \"howto\": S.abertos[k] = !S.abertos[k]; render(); return;\n    case \"verCheck\": S.verCheck = true; window.scrollTo(0, 0); render(); return;\n    case \"voltarStatus\": S.verCheck = false; window.scrollTo(0, 0); render(); return;\n    case \"dia\": S.dia = d; render(); return;\n    case \"porque\": {\n      const m = motivo(S.est, S.dados.instrutor, S.dia, k, { viagem: S.dados.viagem });\n      modalPorque(m.txt || \"Horário indisponível.\");\n      return;\n    }\n    case \"encaixe\": modalEncaixe(k); return;\n    case \"pedir\": modalConfirmarPedido(k); return;\n\n    case \"enviarPedido\": case \"enviarEncaixe\": {\n      const just = document.getElementById(\"just\");\n      const r = await acao(\"solicitar\", {\n        nomeGuerra: S.dados.nomeGuerra, instrutor: S.dados.instrutor, data: S.dia, tempo: k,\n        endereco: S.dados.endereco, viagem: S.dados.viagem, novaOM: S.dados.novaOM,\n        telefone: S.dados.telefone, email: S.dados.email, checklist: S.checklist,\n        aceitou: true, encaixe: act === \"enviarEncaixe\", justificativa: just ? just.value : \"\",\n        tokAlunos: S.tokAlunos\n      });\n      fecharModal();\n      if (r && r.ok) { toast(act === \"enviarEncaixe\" ? \"Pedido de encaixe enviado.\" : \"Pedido enviado ao instrutor.\"); await carregar(); window.scrollTo(0, 0); }\n      return;\n    }\n    case \"cancelar\": modalCancelar(); return;\n    case \"enviarCancelar\": {\n      const nome = document.getElementById(\"cnome\").value;\n      const mot = document.getElementById(\"cmot\").value;\n      const r = await acao(\"cancelar\", S.meu.id, nome, mot);\n      if (r && r.ok) { fecharModal(); toast(\"Vistoria cancelada.\"); S.meu = null; S.step = 5; await carregar(); }\n      else if (S.erro) { fecharModal(); render(); }\n      return;\n    }\n    case \"remarcar\": modalRemarcar(); return;\n    case \"enviarRemarcar\": {\n      const nome = document.getElementById(\"rn\").value;\n      const dia = document.getElementById(\"rd\").value;\n      const tmp = document.getElementById(\"rt\").value;\n      if (!dia) { toast(\"Escolha um dia livre.\"); return; }\n      const r = await acao(\"remarcar\", S.meu.id, nome, dia, tmp);\n      if (r && r.ok) { fecharModal(); toast(\"Nova vistoria solicitada.\"); await carregar(); }\n      else if (S.erro) { fecharModal(); render(); }\n      return;\n    }\n    case \"ics\": baixarIcs(); return;\n    case \"zap\": {\n      const txt = \"Senhor \" + (S.meu ? S.meu.instrutor : S.dados.instrutor) + \", sou o \" + (S.meu ? S.meu.nomeGuerra : S.dados.nomeGuerra)\n        + \". Falo sobre a vistoria do meu PNR\" + (S.meu ? \" marcada para \" + f3(S.meu.data) : \"\") + \".\";\n      window.open(\"https://wa.me/?text=\" + encodeURIComponent(txt), \"_blank\");\n      return;\n    }\n\n    /* --- instrutor --- */\n    case \"confirmar\": {\n      const r = await acao(\"confirmar\", S.chave, id);\n      if (r && r.ok) { toast(\"Confirmada. O aluno já vê na tela dele.\"); await carregar(); }\n      return;\n    }\n    case \"recusar\": modalRecusar(id); return;\n    case \"enviarRecusar\": {\n      const r = await acao(\"recusar\", S.chave, id, document.getElementById(\"rmot\").value, document.getElementById(\"rsug\").value);\n      if (r && r.ok) { fecharModal(); toast(\"Recusada. O aluno foi avisado.\"); await carregar(); }\n      return;\n    }\n    case \"transferir\": modalTransferir(id); return;\n    case \"enviarTransferir\": {\n      const r = await acao(\"transferir\", S.chave, id, document.getElementById(\"tdest\").value);\n      if (r && r.ok) { fecharModal(); toast(\"Transferida.\"); await carregar(); }\n      return;\n    }\n    case \"roteiro\": modalRoteiro(id); return;\n    case \"rot\": {\n      const rid = S.roteiroId;\n      S.roteiro[rid] = S.roteiro[rid] || {};\n      S.roteiro[rid][k] = S.roteiro[rid][k] === v ? null : v;\n      const obs = document.getElementById(\"robs\");\n      const guardado = obs ? obs.value : \"\";\n      modalRoteiro(rid);\n      const novo = document.getElementById(\"robs\");\n      if (novo) novo.value = guardado;\n      return;\n    }\n    case \"regOk\": {\n      const obs = document.getElementById(\"robs\");\n      const r = await acao(\"registrarResultado\", S.chave, id, true, [], obs ? obs.value : \"\");\n      if (r && r.ok) { fecharModal(); toast(\"Chaves recebidas. Processo encerrado.\"); await carregar(); }\n      return;\n    }\n    case \"regPend\": {\n      const rid = S.roteiro[id] || {};\n      const lista = ROTEIRO.filter(x => rid[x.k] === \"n\").map(x => x.t);\n      const obs = document.getElementById(\"robs\");\n      const r = await acao(\"registrarResultado\", S.chave, id, false, lista, obs ? obs.value : \"\");\n      if (r && r.ok) { fecharModal(); toast(\"Pendências registradas. O aluno recebeu a lista.\"); await carregar(); }\n      return;\n    }\n    case \"bloquear\": modalBloquear(); return;\n    case \"enviarBloqueio\": {\n      const dia = document.getElementById(\"bd\").value;\n      if (!dia) { toast(\"Escolha o dia.\"); return; }\n      const r = await acao(\"bloquear\", S.chave, dia, document.getElementById(\"bt\").value, document.getElementById(\"bm\").value);\n      if (r && r.ok) { fecharModal(); toast(\"Horário bloqueado.\"); await carregar(); }\n      else if (S.erro) { fecharModal(); render(); }\n      return;\n    }\n    case \"desbloquear\": {\n      const r = await acao(\"desbloquear\", S.chave, d, k);\n      if (r && r.ok) { toast(\"Horário liberado.\"); await carregar(); }\n      return;\n    }\n    case \"excecao\": modalExcecao(); return;\n    case \"enviarExcecao\": {\n      const r = await acao(\"liberarExcecao\", S.chave, document.getElementById(\"xn\").value, document.getElementById(\"xd\").value);\n      if (r && r.ok) { fecharModal(); toast(\"Exceção liberada até \" + f3(r.validaAte) + \".\"); await carregar(); }\n      return;\n    }\n    case \"agendarPor\": modalAgendarPor(); return;\n    case \"enviarAgendarPor\": {\n      const r = await acao(\"agendarPeloInstrutor\", S.chave, {\n        nomeGuerra: document.getElementById(\"an\").value,\n        endereco: document.getElementById(\"ae\").value,\n        data: document.getElementById(\"ad\").value,\n        tempo: document.getElementById(\"at\").value,\n        viagem: document.getElementById(\"av\").value\n      });\n      if (r && r.ok) { fecharModal(); toast(\"Vistoria marcada e confirmada.\"); await carregar(); }\n      else if (S.erro) { fecharModal(); render(); }\n      return;\n    }\n    case \"salvarEmail\": {\n      const em = document.getElementById(\"eminst\").value;\n      const r = await acao(\"salvarEmail\", S.chave, em);\n      if (r && r.ok) { toast(\"E-mail salvo. O senhor passa a receber os lembretes.\"); await carregar(); }\n      return;\n    }\n\n    /* --- painel --- */\n    case \"cobrar\": modalCobrar(n); return;\n    case \"copiar\": case \"copiarEl\": {\n      const t = document.getElementById(act === \"copiar\" ? \"msgcobrar\" : k);\n      if (t) {\n        try { await navigator.clipboard.writeText(t.textContent); toast(\"Copiado.\"); }\n        catch (x) { toast(\"Selecione e copie o texto à mão.\"); }\n      }\n      return;\n    }\n  }\n});\n\ndocument.addEventListener(\"keydown\", e => { if (e.key === \"Escape\") fecharModal(); });\n\n/* ===================== início ===================== */\n\n(async function () {\n  lerLocal();\n  if (window.google && google.script && google.script.run) {\n    API = gasAPI();\n    const loc = await new Promise(res => { try { google.script.url.getLocation(res); } catch (e) { res({ parameter: {} }); } });\n    const p = loc.parameter || {};\n    S.chave = p.chave || \"\";\n    S.tokAlunos = p.i || \"\";          // ?i=<chave> → link que o instrutor passou aos orientados dele\n    if (p.v === \"painel\") S.view = \"painel\";\n    else if (p.v === \"instrutor\") S.view = \"instrutor\";\n    conn(\"\", \"Conectando à planilha…\");\n  } else {\n    API = demoAPI();\n    banner('<b>Modo demonstração.</b> Esta cópia não está ligada à planilha: os dados ficam só neste navegador. '\n      + 'Para testar as outras telas: <button class=\"btn sm\" data-act=\"demoAluno\">aluno</button> '\n      + '<button class=\"btn sm\" data-act=\"demoInstrutor\">instrutor</button> '\n      + '<button class=\"btn sm\" data-act=\"demoPainel\">painel</button>', \"w\");\n    document.addEventListener(\"click\", ev => {\n      const b = ev.target.closest(\"[data-act^=demo]\");\n      if (!b) return;\n      const a = b.getAttribute(\"data-act\");\n      if (a === \"demoAluno\") API.trocarPapel(\"aluno\", \"\");\n      if (a === \"demoInstrutor\") API.trocarPapel(\"instrutor\", \"Cap Dantas\");\n      if (a === \"demoPainel\") API.trocarPapel(\"painel\", \"\");\n      S.view = \"\"; S.meu = null; carregar();\n    });\n  }\n  await carregar();\n  // Atualiza sozinho quando a tela volta a ficar visível (instrutor e painel).\n  document.addEventListener(\"visibilitychange\", () => { if (!document.hidden && S.papel !== \"aluno\") carregar(true); });\n  setInterval(() => { if (!document.hidden && S.papel !== \"aluno\") carregar(true); }, 60000);\n})();\n</script>\n";
