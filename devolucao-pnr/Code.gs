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
  return HtmlService.createHtmlOutputFromFile('index')
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
