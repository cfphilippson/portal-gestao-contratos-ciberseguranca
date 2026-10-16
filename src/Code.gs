/**
 * Portal de Gestão de Contratos de Serviços em Cibersegurança
 * Backend — Google Apps Script (V8)
 *
 * Deploy: Implantar > Nova implantação > App da Web
 *   Executar como: Eu (dono da planilha) · Quem pode acessar: Qualquer pessoa em <domínio>
 *
 * Funções terminadas em "_" são privadas (não podem ser chamadas via google.script.run).
 * Toda função pública exposta ao cliente valida identidade e perfil no servidor.
 */

const APP = {
  NAME: 'Portal de Gestão de Contratos de Serviços em Cibersegurança',
  SHORT: 'Contratos Cyber',
  TZ: 'America/Sao_Paulo',
  PERFIS: ['Administrador', 'Gestor', 'Leitura'],
  STATUS_CONTRATO: ['Em vigor', 'Em renovação', 'Encerrado'],
  TEMAS: ['light', 'dark', 'system'],
};

const SH = {
  CONTRATOS: 'Contratos', USUARIOS: 'Usuarios', SOLIC: 'Solicitacoes',
  LOG: 'LogAlertas', AUDIT: 'Auditoria', CONFIG: 'Config',
};

const HEADERS = {
  Contratos: ['id', 'parceiro', 'tipo_servico', 'ferramenta', 'finalidade', 'vigencia_inicio', 'vigencia_fim',
    'indice_reajuste', 'mes_reajuste', 'sla', 'renovacao_tipo', 'valor_suporte', 'valor_orcado',
    'valor_pago_anterior', 'observacoes', 'gestor_nome', 'gestor_email', 'status_renovacao',
    'atualizado_em', 'atualizado_por'],
  Usuarios: ['email', 'nome', 'perfil', 'escopo', 'status', 'criado_em', 'aprovado_por', 'ultimo_acesso', 'tema'],
  Solicitacoes: ['id', 'email', 'nome', 'area', 'perfil_solicitado', 'justificativa', 'status', 'criada_em',
    'decidido_por', 'decidido_em', 'motivo'],
  LogAlertas: ['chave', 'contrato_id', 'parceiro', 'ferramenta', 'janela', 'vigencia_fim', 'enviado_em',
    'destinatarios', 'status', 'erro'],
  Auditoria: ['quando', 'usuario', 'acao', 'detalhe'],
  Config: ['chave', 'valor', 'descricao'],
};

const DEFAULT_CONFIG = [
  ['JANELAS', '90,60,30', 'Dias antes do fim da vigência em que o alerta é enviado (separados por vírgula)'],
  ['NOTIFICAR_ADMINS', 'SIM', 'SIM/NAO — copiar os administradores do portal nos alertas'],
  ['CC_FIXO', '', 'E-mails sempre copiados nos alertas (lista de distribuição), separados por vírgula'],
  ['REENVIO_VENCIDOS', 'SIM', 'SIM/NAO — reenviar aviso semanal enquanto o contrato estiver vencido'],
  ['DIA_REENVIO', '1', 'Dia da semana do reenvio de vencidos (1=segunda … 7=domingo)'],
  ['HORA_GATILHO', '7', 'Hora do gatilho diário de alertas (0–23, America/Sao_Paulo)'],
  ['AUTO_REFRESH_SEGUNDOS', '60', 'Intervalo (s) em que o portal verifica se os dados mudaram (mín. 15)'],
  ['LOGO_FILE_ID', '', 'Opcional: ID de um arquivo no Drive (versão para fundo escuro) com o logo da organização'],
  ['ROTULO_ORCADO', 'Orçado 2026', 'Rótulo da coluna valor_orcado'],
  ['ROTULO_PAGO', 'Pago 2025', 'Rótulo da coluna valor_pago_anterior'],
  ['DOMINIO_PERMITIDO', 'example.com', 'Somente e-mails deste domínio podem solicitar acesso (vazio = qualquer)'],
];

/* =========================================================================
 * Web app
 * ========================================================================= */

function doGet() {
  const t = HtmlService.createTemplateFromFile('Index');
  t.appName = APP.NAME;
  return t.evaluate()
    .setTitle(APP.NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/** Usado pelo template Index.html. Lista fechada para não expor outros arquivos. */
function include(name) {
  if (['Styles', 'Scripts'].indexOf(name) < 0) throw new Error('Arquivo não permitido.');
  return HtmlService.createHtmlOutputFromFile(name).getContent();
}

/* =========================================================================
 * Infra: planilha, tabelas, config, datas
 * ========================================================================= */

let _ss = null, _cfg = null, _tz = null;

function ss_() {
  if (_ss) return _ss;
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  _ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActiveSpreadsheet();
  if (!_ss) throw new Error('Planilha não configurada. Execute setup() no editor do Apps Script.');
  return _ss;
}

function tz_() { return _tz || (_tz = ss_().getSpreadsheetTimeZone() || APP.TZ); }

function sheet_(name) {
  const sh = ss_().getSheetByName(name);
  if (!sh) throw new Error('Aba "' + name + '" não encontrada. Execute setup().');
  return sh;
}

function readTable_(name) {
  const values = sheet_(name).getDataRange().getValues();
  if (values.length < 2) return [];
  const head = values[0].map(function (h) { return String(h).trim(); });
  const out = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (r.every(function (c) { return c === '' || c === null; })) continue;
    const o = { _row: i + 1 };
    head.forEach(function (h, j) { if (h) o[h] = r[j]; });
    out.push(o);
  }
  return out;
}

/** Atualiza (rowNum) ou insere (rowNum vazio) preservando colunas não informadas. */
function upsertRow_(name, rowNum, obj) {
  const sh = sheet_(name);
  const lastCol = sh.getLastColumn();
  const head = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  const row = rowNum ? sh.getRange(rowNum, 1, 1, lastCol).getValues()[0] : new Array(lastCol).fill('');
  head.forEach(function (h, i) { if (Object.prototype.hasOwnProperty.call(obj, h)) row[i] = obj[h]; });
  if (rowNum) sh.getRange(rowNum, 1, 1, lastCol).setValues([row]);
  else sh.appendRow(row);
}

function cfg_() {
  if (_cfg) return _cfg;
  const m = {};
  readTable_(SH.CONFIG).forEach(function (r) {
    if (r.chave) m[String(r.chave).trim()] = String(r.valor == null ? '' : r.valor).trim();
  });
  DEFAULT_CONFIG.forEach(function (d) { if (!(d[0] in m)) m[d[0]] = d[1]; });
  _cfg = m;
  return m;
}

function setConfig_(key, value) {
  const rows = readTable_(SH.CONFIG);
  const hit = rows.filter(function (r) { return String(r.chave).trim() === key; })[0];
  const desc = (DEFAULT_CONFIG.filter(function (d) { return d[0] === key; })[0] || [])[2] || '';
  upsertRow_(SH.CONFIG, hit ? hit._row : null, hit ? { valor: value } : { chave: key, valor: value, descricao: desc });
  _cfg = null;
}

/** Janelas de alerta em ordem decrescente, ex.: [90, 60, 30]. */
function janelas_() {
  const j = String(cfg_().JANELAS || '90,60,30').split(/[,; ]+/).map(Number)
    .filter(function (n) { return n > 0 && n <= 365; });
  return uniq_(j).sort(function (a, b) { return b - a; });
}

function pad2_(n) { return ('0' + n).slice(-2); }
function str_(v) { return v == null ? '' : String(v).trim(); }
function uniq_(a) { return a.filter(function (x, i) { return x !== '' && a.indexOf(x) === i; }); }

function num_(v) {
  if (typeof v === 'number' && isFinite(v)) return v;
  const s = str_(v);
  if (!s) return null;
  const t = s.replace(/[R$\s]/g, '');
  // "251.234,50" (pt-BR) ou "251234.50" (input numérico do navegador)
  const n = t.indexOf(',') >= 0 ? Number(t.replace(/\./g, '').replace(',', '.')) : Number(t);
  return isFinite(n) ? n : null;
}

function isoDate_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz_(), 'yyyy-MM-dd');
  const s = str_(v);
  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m) return m[1] + '-' + m[2] + '-' + m[3];
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (m) return m[3] + '-' + pad2_(m[2]) + '-' + pad2_(m[1]);
  return '';
}

function fmtDT_(v) {
  if (v instanceof Date && !isNaN(v)) return Utilities.formatDate(v, tz_(), 'dd/MM/yyyy HH:mm');
  return str_(v);
}

function brDate_(iso) { return iso ? iso.slice(8, 10) + '/' + iso.slice(5, 7) + '/' + iso.slice(0, 4) : '—'; }
function brl_(v) {
  if (v == null) return '—';
  const p = v.toFixed(2).split('.');
  return 'R$ ' + p[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.') + ',' + p[1];
}
function today_() { return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd'); }
function daysBetween_(a, b) { return Math.round((Date.parse(b + 'T00:00:00Z') - Date.parse(a + 'T00:00:00Z')) / 864e5); }
function addDays_(iso, n) { return new Date(Date.parse(iso + 'T00:00:00Z') + n * 864e5).toISOString().slice(0, 10); }
function toSheetDate_(iso) { const p = iso.split('-').map(Number); return new Date(p[0], p[1] - 1, p[2]); }

function validIso_(s) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(s)) return false;
  const d = new Date(s + 'T00:00:00Z');
  return !isNaN(d) && d.toISOString().slice(0, 10) === s;
}

/** Evita injeção de fórmula ao gravar texto vindo do navegador. */
function safeText_(v, max) {
  let s = str_(v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
  if (max && s.length > max) s = s.slice(0, max);
  if (/^[=+\-@\t\r]/.test(s)) s = "'" + s;
  return s;
}

function isEmail_(s) { return /^[^\s@,;<>"']+@[^\s@,;<>"']+\.[a-z]{2,}$/i.test(s); }
function splitEmails_(s) {
  return uniq_(str_(s).split(/[,;\s]+/).map(function (e) { return e.toLowerCase(); }).filter(isEmail_));
}

/* Versão dos dados: muda a cada escrita (portal ou edição direta na planilha). */
function getVersion_() { return PropertiesService.getScriptProperties().getProperty('DATA_VERSION') || '0'; }
function bumpVersion_() {
  const v = String(Date.now());
  PropertiesService.getScriptProperties().setProperty('DATA_VERSION', v);
  return v;
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('Portal ocupado no momento. Tente novamente em alguns segundos.');
  try { return fn(); } finally { lock.releaseLock(); }
}

function audit_(user, acao, detalhe) {
  sheet_(SH.AUDIT).appendRow([new Date(), user, safeText_(acao, 200), safeText_(detalhe, 2000)]);
}

/* =========================================================================
 * Identidade e autorização
 * ========================================================================= */

function currentEmail_() {
  return str_(Session.getActiveUser().getEmail()).toLowerCase();
}

function normUser_(r) {
  return {
    email: str_(r.email).toLowerCase(), nome: str_(r.nome), perfil: str_(r.perfil) || 'Leitura',
    escopo: str_(r.escopo) || '*', status: str_(r.status) || 'Ativo', tema: str_(r.tema) || 'system',
    criado_em: fmtDT_(r.criado_em), aprovado_por: str_(r.aprovado_por), ultimo_acesso: fmtDT_(r.ultimo_acesso),
    _row: r._row,
  };
}

function allUsers_() { return readTable_(SH.USUARIOS).filter(function (r) { return r.email; }).map(normUser_); }
function activeUsers_() { return allUsers_().filter(function (u) { return u.status === 'Ativo'; }); }
function findUser_(email) { return allUsers_().filter(function (u) { return u.email === email; })[0] || null; }
function adminEmails_() {
  return activeUsers_().filter(function (u) { return u.perfil === 'Administrador'; }).map(function (u) { return u.email; });
}

function requireUser_(perfis) {
  const email = currentEmail_();
  if (!email) throw new Error('SEM_IDENTIDADE');
  const u = findUser_(email);
  if (!u || u.status !== 'Ativo') throw new Error('ACESSO_NEGADO');
  if (perfis && perfis.indexOf(u.perfil) < 0) throw new Error('Permissão insuficiente para esta ação.');
  return u;
}

function forgetUserCache_(email) { CacheService.getScriptCache().remove('ok:' + email); }

/** Escopo: "*" = todos; senão lista de IDs (CT-001,CT-004). Gestor também vê contratos em que é o gestor. */
function canSee_(u, c) {
  if (u.perfil === 'Administrador') return true;
  if (c.gestor_email && c.gestor_email === u.email) return true;
  const esc = str_(u.escopo);
  if (!esc || esc === '*') return true;
  return esc.toUpperCase().split(/[,; ]+/).indexOf(c.id.toUpperCase()) >= 0;
}

/* =========================================================================
 * Contratos
 * ========================================================================= */

function faixa_(c, jan) {
  if (c.status_renovacao === 'Encerrado') return 'encerrado';
  if (c.dias === null) return 'sem_data';
  if (c.dias < 0) return 'vencido';
  const asc = jan.slice().sort(function (a, b) { return a - b; });
  for (let i = 0; i < asc.length; i++) if (c.dias <= asc[i]) return 'j' + asc[i];
  return 'ok';
}

function enrichContract_(r, today, jan) {
  const c = {
    id: str_(r.id), parceiro: str_(r.parceiro), tipo_servico: str_(r.tipo_servico), ferramenta: str_(r.ferramenta),
    finalidade: str_(r.finalidade), vigencia_inicio: isoDate_(r.vigencia_inicio), vigencia_fim: isoDate_(r.vigencia_fim),
    indice_reajuste: str_(r.indice_reajuste), mes_reajuste: str_(r.mes_reajuste), sla: str_(r.sla),
    renovacao_tipo: str_(r.renovacao_tipo), valor_suporte: str_(r.valor_suporte),
    valor_orcado: num_(r.valor_orcado), valor_pago_anterior: num_(r.valor_pago_anterior),
    observacoes: str_(r.observacoes), gestor_nome: str_(r.gestor_nome), gestor_email: str_(r.gestor_email).toLowerCase(),
    status_renovacao: str_(r.status_renovacao) || 'Em vigor',
    atualizado_em: fmtDT_(r.atualizado_em), atualizado_por: str_(r.atualizado_por), _row: r._row,
  };
  c.dias = c.vigencia_fim ? daysBetween_(today, c.vigencia_fim) : null;
  c.faixa = faixa_(c, jan);
  return c;
}

function allContracts_() {
  const today = today_(), jan = janelas_();
  return readTable_(SH.CONTRATOS).filter(function (r) { return str_(r.id); })
    .map(function (r) { return enrichContract_(r, today, jan); });
}

function stripRow_(o) { const c = {}; Object.keys(o).forEach(function (k) { if (k !== '_row') c[k] = o[k]; }); return c; }

function nextAlerts_(contratos, jan, today) {
  const out = [];
  contratos.forEach(function (c) {
    if (!c.vigencia_fim || c.status_renovacao === 'Encerrado') return;
    jan.forEach(function (w) {
      const d = addDays_(c.vigencia_fim, -w);
      if (d > today) out.push({ data: d, janela: w, id: c.id, ferramenta: c.ferramenta, parceiro: c.parceiro, vigencia_fim: c.vigencia_fim });
    });
  });
  return out.sort(function (a, b) { return a.data < b.data ? -1 : a.data > b.data ? 1 : 0; }).slice(0, 25);
}

function regrasPublicas_() {
  const c = cfg_();
  return {
    janelas: janelas_(), notificarAdmins: c.NOTIFICAR_ADMINS === 'SIM', ccFixo: c.CC_FIXO,
    reenvioVencidos: c.REENVIO_VENCIDOS === 'SIM', diaReenvio: Number(c.DIA_REENVIO) || 1,
    horaGatilho: Number(c.HORA_GATILHO) || 7, autoRefreshSeg: Math.max(15, Number(c.AUTO_REFRESH_SEGUNDOS) || 60),
  };
}

function lastRun_() {
  try { return JSON.parse(PropertiesService.getScriptProperties().getProperty('LAST_RUN') || 'null'); }
  catch (e) { return null; }
}

/* =========================================================================
 * API pública — leitura
 * ========================================================================= */

function getBootstrap() {
  const email = currentEmail_();
  const c = cfg_();
  const base = {
    app: { nome: APP.NAME, curto: APP.SHORT, logo: logoDataUri_(), autoRefreshSeg: Math.max(15, Number(c.AUTO_REFRESH_SEGUNDOS) || 60), dominio: c.DOMINIO_PERMITIDO, perfis: APP.PERFIS, statusContrato: APP.STATUS_CONTRATO },
    version: getVersion_(),
  };
  if (!email) { base.user = { email: '', status: 'sem_identidade' }; return base; }

  const u = findUser_(email);
  if (u && u.status === 'Ativo') {
    try { upsertRow_(SH.USUARIOS, u._row, { ultimo_acesso: new Date() }); } catch (e) { /* não bloqueia */ }
    base.user = { email: email, nome: u.nome, perfil: u.perfil, status: 'ativo', tema: u.tema };
    base.sheetUrl = u.perfil === 'Administrador' ? ss_().getUrl() : null;
    return base;
  }
  const req = latestRequest_(email);
  let status = 'sem_acesso';
  if (req && req.status === 'Pendente') status = 'pendente';
  else if (req && req.status === 'Recusada') status = 'recusado';
  base.user = { email: email, status: status };
  base.request = req ? stripRow_(req) : null;
  return base;
}

function getData() {
  const u = requireUser_();
  const today = today_(), jan = janelas_(), c = cfg_();
  const contratos = allContracts_().filter(function (x) { return canSee_(u, x); });
  const ids = {};
  contratos.forEach(function (x) { ids[x.id] = true; });
  const log = readTable_(SH.LOG).filter(function (r) { return ids[str_(r.contrato_id)]; }).map(function (r) {
    const when = r.enviado_em instanceof Date ? r.enviado_em : null;
    return {
      ordem: when ? when.getTime() : 0, enviado_em: fmtDT_(r.enviado_em), contrato_id: str_(r.contrato_id),
      parceiro: str_(r.parceiro), ferramenta: str_(r.ferramenta), janela: str_(r.janela),
      vigencia_fim: isoDate_(r.vigencia_fim), status: str_(r.status),
      destinatarios: u.perfil === 'Administrador' ? str_(r.destinatarios) : '',
    };
  }).sort(function (a, b) { return b.ordem - a.ordem; }).slice(0, 100);

  return {
    today: today, janelas: jan, version: getVersion_(),
    sincronizadoEm: Utilities.formatDate(new Date(), tz_(), 'dd/MM/yyyy HH:mm:ss'),
    rotulos: { orcado: c.ROTULO_ORCADO, pago: c.ROTULO_PAGO },
    contratos: contratos.map(stripRow_), log: log, proximos: nextAlerts_(contratos, jan, today),
    regras: regrasPublicas_(), ultimaExecucao: lastRun_(),
  };
}

/** Chamado pelo polling do auto-refresh. Barato: cache de autorização por 2 min. */
function getVersion() {
  const email = currentEmail_();
  if (!email) throw new Error('SEM_IDENTIDADE');
  const cache = CacheService.getScriptCache();
  if (!cache.get('ok:' + email)) { requireUser_(); cache.put('ok:' + email, '1', 120); }
  return getVersion_();
}

function setTheme(tema) {
  if (APP.TEMAS.indexOf(tema) < 0) throw new Error('Tema inválido.');
  const email = currentEmail_();
  const u = email && findUser_(email);
  if (!u || u.status !== 'Ativo') return false; // sem acesso: tema fica só no navegador
  upsertRow_(SH.USUARIOS, u._row, { tema: tema });
  return true;
}

/* =========================================================================
 * API pública — solicitação de acesso
 * ========================================================================= */

function latestRequest_(email) {
  const rows = readTable_(SH.SOLIC).filter(function (r) { return str_(r.email).toLowerCase() === email; });
  if (!rows.length) return null;
  const r = rows[rows.length - 1];
  return {
    id: str_(r.id), email: email, nome: str_(r.nome), area: str_(r.area), perfil_solicitado: str_(r.perfil_solicitado),
    justificativa: str_(r.justificativa), status: str_(r.status), criada_em: fmtDT_(r.criada_em),
    decidido_por: str_(r.decidido_por), decidido_em: fmtDT_(r.decidido_em), motivo: str_(r.motivo), _row: r._row,
  };
}

function submitAccessRequest(form) {
  const email = currentEmail_();
  if (!email) throw new Error('Não foi possível identificar sua conta. Acesse com a conta corporativa.');
  const dom = str_(cfg_().DOMINIO_PERMITIDO).toLowerCase();
  if (dom && !email.endsWith('@' + dom)) throw new Error('Somente contas @' + dom + ' podem solicitar acesso.');
  form = form || {};
  const nome = safeText_(form.nome, 120), area = safeText_(form.area, 120), just = safeText_(form.justificativa, 1000);
  const perfil = str_(form.perfil);
  if (nome.length < 3) throw new Error('Informe seu nome completo.');
  if (area.length < 2) throw new Error('Informe sua área.');
  if (APP.PERFIS.indexOf(perfil) < 0) throw new Error('Perfil inválido.');
  if (just.length < 10) throw new Error('Descreva a justificativa (mínimo de 10 caracteres).');

  withLock_(function () {
    const u = findUser_(email);
    if (u && u.status === 'Ativo') throw new Error('Você já possui acesso. Recarregue a página.');
    const prev = latestRequest_(email);
    if (prev && prev.status === 'Pendente') throw new Error('Você já tem uma solicitação pendente.');
    const id = 'SOL-' + Utilities.formatDate(new Date(), tz_(), 'yyyyMMddHHmmss') + '-' + Math.floor(Math.random() * 1000);
    upsertRow_(SH.SOLIC, null, {
      id: id, email: email, nome: nome, area: area, perfil_solicitado: perfil, justificativa: just,
      status: 'Pendente', criada_em: new Date(),
    });
    audit_(email, 'Solicitou acesso', perfil + ' · ' + area);
    bumpVersion_();
  });

  try {
    const admins = adminEmails_();
    if (admins.length) {
      MailApp.sendEmail({
        to: admins.join(','), name: APP.SHORT,
        subject: '[' + APP.SHORT + '] Nova solicitação de acesso — ' + nome,
        htmlBody: '<p>' + escHtml_(nome) + ' (' + escHtml_(email) + ') solicitou acesso como <b>' + escHtml_(perfil) +
          '</b>.</p><p>Área: ' + escHtml_(area) + '<br>Justificativa: ' + escHtml_(just) + '</p>' +
          '<p><a href="' + escHtml_(ScriptApp.getService().getUrl()) + '">Abrir o portal para aprovar</a></p>',
      });
    }
  } catch (e) { console.warn('Falha ao notificar admins: ' + e); }
  return getBootstrap();
}

function escHtml_(s) {
  return str_(s).replace(/[&<>"']/g, function (c) { return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]; });
}

/* =========================================================================
 * API pública — administração de acessos
 * ========================================================================= */

function getAdminData() {
  requireUser_(['Administrador']);
  const reqs = readTable_(SH.SOLIC).map(function (r) {
    return {
      id: str_(r.id), email: str_(r.email).toLowerCase(), nome: str_(r.nome), area: str_(r.area),
      perfil_solicitado: str_(r.perfil_solicitado), justificativa: str_(r.justificativa), status: str_(r.status),
      criada_em: fmtDT_(r.criada_em), decidido_por: str_(r.decidido_por), decidido_em: fmtDT_(r.decidido_em), motivo: str_(r.motivo),
    };
  });
  const audit = readTable_(SH.AUDIT).slice(-80).reverse().map(function (r) {
    return { quando: fmtDT_(r.quando), usuario: str_(r.usuario), acao: str_(r.acao), detalhe: str_(r.detalhe) };
  });
  return {
    usuarios: activeUsers_().map(stripRow_),
    pendentes: reqs.filter(function (r) { return r.status === 'Pendente'; }),
    decididas: reqs.filter(function (r) { return r.status !== 'Pendente'; }).slice(-20).reverse(),
    auditoria: audit,
    contratosRef: allContracts_().map(function (c) { return { id: c.id, label: c.ferramenta + ' · ' + c.parceiro }; }),
    me: currentEmail_(),
  };
}

function normEscopo_(s) {
  s = str_(s);
  if (!s || s === '*') return '*';
  const ids = uniq_(s.toUpperCase().split(/[,; ]+/).filter(function (x) { return /^[A-Z0-9_-]{1,20}$/.test(x); }));
  return ids.length ? ids.join(',') : '*';
}

function decideRequest(id, decisao, perfil, escopo, motivo) {
  const admin = requireUser_(['Administrador']);
  if (['aprovar', 'recusar'].indexOf(decisao) < 0) throw new Error('Decisão inválida.');
  if (decisao === 'aprovar' && APP.PERFIS.indexOf(perfil) < 0) throw new Error('Perfil inválido.');
  let req;
  withLock_(function () {
    req = readTable_(SH.SOLIC).filter(function (r) { return str_(r.id) === str_(id); })[0];
    if (!req) throw new Error('Solicitação não encontrada.');
    if (str_(req.status) !== 'Pendente') throw new Error('Esta solicitação já foi decidida.');
    const email = str_(req.email).toLowerCase();
    if (decisao === 'aprovar') {
      const u = findUser_(email);
      const data = { email: email, nome: safeText_(req.nome, 120), perfil: perfil, escopo: normEscopo_(escopo), status: 'Ativo', aprovado_por: admin.email };
      if (!u) { data.criado_em = new Date(); data.tema = 'system'; }
      upsertRow_(SH.USUARIOS, u ? u._row : null, data);
      forgetUserCache_(email);
    }
    upsertRow_(SH.SOLIC, req._row, {
      status: decisao === 'aprovar' ? 'Aprovada' : 'Recusada', decidido_por: admin.email,
      decidido_em: new Date(), motivo: safeText_(motivo, 500),
    });
    audit_(admin.email, decisao === 'aprovar' ? 'Aprovou acesso' : 'Recusou acesso',
      email + (decisao === 'aprovar' ? ' como ' + perfil : '') + (motivo ? ' · ' + motivo : ''));
    bumpVersion_();
  });

  try {
    const ok = decisao === 'aprovar';
    MailApp.sendEmail({
      to: str_(req.email), name: APP.SHORT,
      subject: '[' + APP.SHORT + '] Solicitação de acesso ' + (ok ? 'aprovada' : 'recusada'),
      htmlBody: ok
        ? '<p>Seu acesso ao <b>' + APP.NAME + '</b> foi aprovado com o perfil <b>' + escHtml_(perfil) + '</b>.</p><p><a href="' + escHtml_(ScriptApp.getService().getUrl()) + '">Acessar o portal</a></p>'
        : '<p>Sua solicitação de acesso ao <b>' + APP.NAME + '</b> foi recusada.</p>' + (motivo ? '<p>Motivo: ' + escHtml_(motivo) + '</p>' : ''),
    });
  } catch (e) { console.warn('Falha ao notificar solicitante: ' + e); }
  return getAdminData();
}

function updateUser(email, perfil, escopo) {
  const admin = requireUser_(['Administrador']);
  email = str_(email).toLowerCase();
  if (APP.PERFIS.indexOf(perfil) < 0) throw new Error('Perfil inválido.');
  withLock_(function () {
    const u = findUser_(email);
    if (!u || u.status !== 'Ativo') throw new Error('Usuário não encontrado.');
    if (u.perfil === 'Administrador' && perfil !== 'Administrador' && adminEmails_().length <= 1)
      throw new Error('Não é possível rebaixar o último administrador.');
    upsertRow_(SH.USUARIOS, u._row, { perfil: perfil, escopo: normEscopo_(escopo) });
    forgetUserCache_(email);
    audit_(admin.email, 'Alterou usuário', email + ' → ' + perfil + ' · escopo ' + normEscopo_(escopo));
    bumpVersion_();
  });
  return getAdminData();
}

function removeUser(email) {
  const admin = requireUser_(['Administrador']);
  email = str_(email).toLowerCase();
  if (email === admin.email) throw new Error('Você não pode remover o seu próprio acesso.');
  withLock_(function () {
    const u = findUser_(email);
    if (!u) throw new Error('Usuário não encontrado.');
    if (u.perfil === 'Administrador' && adminEmails_().length <= 1) throw new Error('Não é possível remover o último administrador.');
    sheet_(SH.USUARIOS).deleteRow(u._row);
    forgetUserCache_(email);
    audit_(admin.email, 'Removeu usuário', email + ' (' + u.perfil + ')');
    bumpVersion_();
  });
  return getAdminData();
}

/* =========================================================================
 * API pública — manutenção de contratos (Administrador)
 * ========================================================================= */

function contractFromForm_(f) {
  const o = {
    parceiro: safeText_(f.parceiro, 120), tipo_servico: safeText_(f.tipo_servico, 120), ferramenta: safeText_(f.ferramenta, 120),
    finalidade: safeText_(f.finalidade, 300), indice_reajuste: safeText_(f.indice_reajuste, 40), mes_reajuste: safeText_(f.mes_reajuste, 40),
    sla: safeText_(f.sla, 200), renovacao_tipo: safeText_(f.renovacao_tipo, 120), valor_suporte: safeText_(f.valor_suporte, 300),
    observacoes: safeText_(f.observacoes, 2000), gestor_nome: safeText_(f.gestor_nome, 120),
    gestor_email: str_(f.gestor_email).toLowerCase(), status_renovacao: str_(f.status_renovacao) || 'Em vigor',
  };
  if (!o.parceiro || !o.ferramenta) throw new Error('Parceiro e ferramenta são obrigatórios.');
  if (o.gestor_email && !isEmail_(o.gestor_email)) throw new Error('E-mail do gestor inválido.');
  if (APP.STATUS_CONTRATO.indexOf(o.status_renovacao) < 0) throw new Error('Status inválido.');
  ['vigencia_inicio', 'vigencia_fim'].forEach(function (k) {
    const v = str_(f[k]);
    if (v && !validIso_(v)) throw new Error('Data inválida em ' + k + '.');
    o[k] = v ? toSheetDate_(v) : '';
  });
  if (o.vigencia_inicio && o.vigencia_fim && o.vigencia_fim < o.vigencia_inicio) throw new Error('O fim da vigência é anterior ao início.');
  ['valor_orcado', 'valor_pago_anterior'].forEach(function (k) {
    const v = str_(f[k]);
    if (!v) { o[k] = ''; return; }
    const n = num_(v);
    if (n === null || n < 0) throw new Error('Valor inválido em ' + k + '.');
    o[k] = n;
  });
  return o;
}

function saveContract(f) {
  const admin = requireUser_(['Administrador']);
  f = f || {};
  const data = contractFromForm_(f);
  data.atualizado_em = new Date();
  data.atualizado_por = admin.email;
  withLock_(function () {
    const rows = readTable_(SH.CONTRATOS);
    const id = str_(f.id);
    if (id) {
      const r = rows.filter(function (x) { return str_(x.id) === id; })[0];
      if (!r) throw new Error('Contrato não encontrado.');
      upsertRow_(SH.CONTRATOS, r._row, data);
      audit_(admin.email, 'Editou contrato', id + ' · ' + data.ferramenta);
    } else {
      const max = rows.reduce(function (m, x) { const n = Number(str_(x.id).replace(/\D/g, '')); return n > m ? n : m; }, 0);
      data.id = 'CT-' + ('00' + (max + 1)).slice(-3);
      upsertRow_(SH.CONTRATOS, null, data);
      audit_(admin.email, 'Criou contrato', data.id + ' · ' + data.ferramenta);
    }
    bumpVersion_();
  });
  return getData();
}

function registerRenewal(id, inicio, fim, valorOrcado, obs) {
  const admin = requireUser_(['Administrador']);
  if (!validIso_(str_(inicio)) || !validIso_(str_(fim))) throw new Error('Informe início e fim válidos.');
  if (fim <= inicio) throw new Error('O fim deve ser posterior ao início.');
  withLock_(function () {
    const r = readTable_(SH.CONTRATOS).filter(function (x) { return str_(x.id) === str_(id); })[0];
    if (!r) throw new Error('Contrato não encontrado.');
    const upd = { vigencia_inicio: toSheetDate_(inicio), vigencia_fim: toSheetDate_(fim), status_renovacao: 'Em vigor', atualizado_em: new Date(), atualizado_por: admin.email };
    const n = num_(valorOrcado);
    if (str_(valorOrcado) && (n === null || n < 0)) throw new Error('Valor orçado inválido.');
    if (n !== null) upd.valor_orcado = n;
    if (str_(obs)) upd.observacoes = safeText_(str_(r.observacoes) + '\n[' + brDate_(today_()) + ' renovação] ' + str_(obs), 2000);
    upsertRow_(SH.CONTRATOS, r._row, upd);
    audit_(admin.email, 'Registrou renovação', id + ' · ' + brDate_(inicio) + ' → ' + brDate_(fim));
    bumpVersion_();
  });
  return getData();
}

function deleteContract(id) {
  const admin = requireUser_(['Administrador']);
  withLock_(function () {
    const r = readTable_(SH.CONTRATOS).filter(function (x) { return str_(x.id) === str_(id); })[0];
    if (!r) throw new Error('Contrato não encontrado.');
    sheet_(SH.CONTRATOS).deleteRow(r._row);
    audit_(admin.email, 'Excluiu contrato', id + ' · ' + str_(r.ferramenta) + ' (' + str_(r.parceiro) + ')');
    bumpVersion_();
  });
  return getData();
}

/* =========================================================================
 * API pública — regras de alerta (Administrador)
 * ========================================================================= */

function saveAlertConfig(r) {
  const admin = requireUser_(['Administrador']);
  r = r || {};
  const jan = uniq_((r.janelas || []).map(Number).filter(function (n) { return n > 0 && n <= 365; }));
  if (!jan.length) throw new Error('Mantenha ao menos uma janela de alerta ativa.');
  const cc = str_(r.ccFixo) ? splitEmails_(r.ccFixo) : [];
  if (str_(r.ccFixo) && !cc.length) throw new Error('Lista de cópia fixa inválida.');
  const hora = Number(r.horaGatilho), dia = Number(r.diaReenvio), auto = Number(r.autoRefreshSeg);
  if (!(hora >= 0 && hora <= 23)) throw new Error('Hora do gatilho inválida.');
  if (!(dia >= 1 && dia <= 7)) throw new Error('Dia de reenvio inválido.');
  if (!(auto >= 15 && auto <= 3600)) throw new Error('Intervalo de auto-atualização deve estar entre 15 e 3600 s.');

  withLock_(function () {
    const horaAnterior = Number(cfg_().HORA_GATILHO);
    setConfig_('JANELAS', jan.sort(function (a, b) { return b - a; }).join(','));
    setConfig_('NOTIFICAR_ADMINS', r.notificarAdmins ? 'SIM' : 'NAO');
    setConfig_('CC_FIXO', cc.join(','));
    setConfig_('REENVIO_VENCIDOS', r.reenvioVencidos ? 'SIM' : 'NAO');
    setConfig_('DIA_REENVIO', String(dia));
    setConfig_('HORA_GATILHO', String(hora));
    setConfig_('AUTO_REFRESH_SEGUNDOS', String(auto));
    if (horaAnterior !== hora) installDailyTrigger_();
    audit_(admin.email, 'Alterou regras de alerta', JSON.stringify(regrasPublicas_()));
    bumpVersion_();
  });
  return getData();
}

function sendTestAlert(id) {
  const admin = requireUser_(['Administrador']);
  const c = allContracts_().filter(function (x) { return x.id === str_(id); })[0];
  if (!c) throw new Error('Contrato não encontrado.');
  if (c.dias === null) throw new Error('Contrato sem data de fim de vigência.');
  const jan = janelas_();
  let janela = c.dias !== null && c.dias < 0 ? 'vencido' : null;
  if (!janela) {
    const asc = jan.slice().sort(function (a, b) { return a - b; });
    janela = asc.filter(function (w) { return c.dias !== null && c.dias <= w; })[0] || asc[asc.length - 1];
  }
  sendAlertEmail_(c, janela, [admin.email], [], true);
  return 'E-mail de teste enviado para ' + admin.email + '.';
}

function runAlertsNow() {
  requireUser_(['Administrador']);
  return withLock_(function () { return processAlerts_(currentEmail_()); });
}

/* =========================================================================
 * Motor de alertas 90/60/30 + vencidos
 * ========================================================================= */

/**
 * Regra: para cada contrato ativo com fim definido, identifica a MENOR janela que contém os dias
 * restantes. Envia uma única vez por chave (contrato|fim|janela). Se a vigência for renovada
 * (novo fim), as chaves mudam e o ciclo recomeça. Falhas não são gravadas como "Enviado" e são
 * tentadas de novo no próximo ciclo. Vencidos: reenvio semanal no DIA_REENVIO, se habilitado.
 */
function processAlerts_(executor) {
  const c = cfg_(), jan = janelas_(), today = today_();
  const asc = jan.slice().sort(function (a, b) { return a - b; });
  const sent = {};
  readTable_(SH.LOG).forEach(function (r) { if (str_(r.status) === 'Enviado') sent[str_(r.chave)] = true; });
  const admins = c.NOTIFICAR_ADMINS === 'SIM' ? adminEmails_() : [];
  const cc = splitEmails_(c.CC_FIXO);
  const weekday = Number(Utilities.formatDate(new Date(), tz_(), 'u')); // 1=segunda … 7=domingo
  const res = { avaliados: 0, enviados: 0, erros: 0, ignorados: 0 };

  allContracts_().forEach(function (k) {
    if (!k.vigencia_fim || k.status_renovacao === 'Encerrado') { res.ignorados++; return; }
    res.avaliados++;
    let janela = null, chave = null;
    if (k.dias >= 0) {
      for (let i = 0; i < asc.length; i++) if (k.dias <= asc[i]) { janela = asc[i]; break; }
      if (janela === null) return;
      chave = [k.id, k.vigencia_fim, janela].join('|');
    } else {
      if (c.REENVIO_VENCIDOS !== 'SIM' || weekday !== (Number(c.DIA_REENVIO) || 1)) return;
      janela = 'vencido';
      chave = [k.id, k.vigencia_fim, 'vencido', today].join('|');
    }
    if (sent[chave]) return;

    const to = k.gestor_email ? [k.gestor_email] : [];
    const ccList = uniq_(admins.concat(cc)).filter(function (e) { return to.indexOf(e) < 0; });
    if (!to.length && ccList.length) to.push(ccList.shift());
    if (!to.length) { res.erros++; logAlert_(chave, k, janela, '', 'Erro', 'Sem destinatários (gestor/admins/CC)'); return; }

    try {
      if (MailApp.getRemainingDailyQuota() < to.length + ccList.length) throw new Error('Cota diária de e-mail esgotada');
      sendAlertEmail_(k, janela, to, ccList, false);
      res.enviados++;
      logAlert_(chave, k, janela, to.concat(ccList).join(', '), 'Enviado', '');
    } catch (e) {
      res.erros++;
      logAlert_(chave, k, janela, to.concat(ccList).join(', '), 'Erro', String(e && e.message || e));
    }
  });

  res.em = Utilities.formatDate(new Date(), tz_(), 'dd/MM/yyyy HH:mm');
  res.por = executor || 'gatilho diário';
  PropertiesService.getScriptProperties().setProperty('LAST_RUN', JSON.stringify(res));
  bumpVersion_();
  return res;
}

function logAlert_(chave, k, janela, dest, status, erro) {
  sheet_(SH.LOG).appendRow([chave, k.id, safeText_(k.parceiro), safeText_(k.ferramenta), String(janela),
    k.vigencia_fim ? toSheetDate_(k.vigencia_fim) : '', new Date(), dest, status, safeText_(erro, 500)]);
}

function proximoAviso_(k, janela) {
  if (janela === 'vencido') return 'Reenvio semanal enquanto o contrato estiver vencido e sem renovação registrada.';
  const menores = janelas_().filter(function (w) { return w < janela; }).sort(function (a, b) { return b - a; });
  if (!menores.length) return 'Último aviso antes do vencimento.';
  return menores[0] + ' dias · ' + brDate_(addDays_(k.vigencia_fim, -menores[0]));
}

function sendAlertEmail_(k, janela, to, cc, teste) {
  const c = cfg_();
  const t = HtmlService.createTemplateFromFile('Email');
  const vencido = janela === 'vencido';
  t.d = {
    appName: APP.NAME, teste: !!teste, vencido: vencido, janela: vencido ? '' : String(janela),
    dias: k.dias, diasAbs: Math.abs(k.dias), fim: brDate_(k.vigencia_fim), inicio: brDate_(k.vigencia_inicio),
    parceiro: k.parceiro, ferramenta: k.ferramenta, finalidade: k.finalidade || '—', tipo: k.tipo_servico || '—',
    sla: k.sla || '—', reajuste: [k.indice_reajuste, k.mes_reajuste].filter(String).join(' · ') || '—',
    valorSuporte: k.valor_suporte || '—', orcado: brl_(k.valor_orcado), pago: brl_(k.valor_pago_anterior),
    rotOrcado: c.ROTULO_ORCADO, rotPago: c.ROTULO_PAGO, gestor: k.gestor_nome || '—', id: k.id,
    proximo: proximoAviso_(k, janela), url: ScriptApp.getService().getUrl() || '', temLogo: false,
  };
  const logo = logoBlob_();
  if (logo) t.d.temLogo = true;
  const html = t.evaluate().getContent();
  const subject = (teste ? '[TESTE] ' : '') + '[' + APP.SHORT + '] ' + (vencido
    ? 'VENCIDO há ' + Math.abs(k.dias) + ' dias · ' + k.ferramenta + ' (' + k.parceiro + ')'
    : 'Janela de ' + janela + ' dias · ' + k.ferramenta + ' (' + k.parceiro + ') vence em ' + brDate_(k.vigencia_fim));
  const plain = subject + '\n\nContrato ' + k.id + ' · ' + k.ferramenta + ' (' + k.parceiro + ')\nFim da vigência: ' +
    brDate_(k.vigencia_fim) + '\nPortal: ' + t.d.url;
  const opts = { to: to.join(','), subject: subject, body: plain, htmlBody: html, name: APP.SHORT };
  if (cc && cc.length) opts.cc = cc.join(',');
  if (logo) opts.inlineImages = { logo: logo };
  MailApp.sendEmail(opts);
}

/* =========================================================================
 * Logo: padrão embutido (Logo.gs) ou arquivo do Drive em Config.LOGO_FILE_ID
 * ========================================================================= */

function logoBlob_() {
  const id = str_(cfg_().LOGO_FILE_ID);
  if (id) {
    try { return DriveApp.getFileById(id).getBlob().setName('logo'); }
    catch (e) { console.warn('Logo do Drive inacessível, usando o padrão: ' + e); }
  }
  if (!LOGO_PADRAO || !LOGO_PADRAO.b64) return null;
  return Utilities.newBlob(Utilities.base64Decode(LOGO_PADRAO.b64), LOGO_PADRAO.mime, 'logo.png');
}

function logoDataUri_() {
  const id = str_(cfg_().LOGO_FILE_ID);
  if (!id) return LOGO_PADRAO && LOGO_PADRAO.b64 ? 'data:' + LOGO_PADRAO.mime + ';base64,' + LOGO_PADRAO.b64 : null;
  const cache = CacheService.getScriptCache();
  const key = 'logo:' + id;
  const hit = cache.get(key);
  if (hit) return hit;
  const b = logoBlob_();
  if (!b) return null;
  const uri = 'data:' + b.getContentType() + ';base64,' + Utilities.base64Encode(b.getBytes());
  if (uri.length < 95000) cache.put(key, uri, 21600);
  return uri;
}

/* =========================================================================
 * Setup e gatilhos
 * ========================================================================= */

/** Execute UMA vez pelo editor (com a planilha aberta como contêiner do script). */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('Abra o editor a partir da planilha (Extensões > Apps Script) e rode setup() novamente.');
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());
  _ss = ss; _cfg = null;

  Object.keys(HEADERS).forEach(function (name) { ensureSheet_(name, HEADERS[name]); });
  const existing = {};
  readTable_(SH.CONFIG).forEach(function (r) { existing[str_(r.chave)] = true; });
  DEFAULT_CONFIG.forEach(function (d) { if (!existing[d[0]]) sheet_(SH.CONFIG).appendRow(d); });
  _cfg = null;

  const me = str_(Session.getEffectiveUser().getEmail()).toLowerCase();
  if (!adminEmails_().length && me) {
    upsertRow_(SH.USUARIOS, null, { email: me, nome: me.split('@')[0], perfil: 'Administrador', escopo: '*', status: 'Ativo', criado_em: new Date(), aprovado_por: 'setup', tema: 'system' });
    audit_(me, 'Setup', 'Administrador inicial criado');
  }
  installTriggers_();
  bumpVersion_();
  console.log('Setup concluído. Admin: ' + me + ' · Gatilhos: ' + ScriptApp.getProjectTriggers().length);
}

function ensureSheet_(name, headers) {
  let sh = ss_().getSheetByName(name);
  if (!sh) sh = ss_().insertSheet(name);
  if (sh.getLastColumn() === 0 || sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].join('') === '') {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
    return;
  }
  const cur = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(function (h) { return String(h).trim(); });
  headers.forEach(function (h) {
    if (cur.indexOf(h) < 0) { sh.getRange(1, cur.length + 1).setValue(h).setFontWeight('bold'); cur.push(h); }
  });
}

function installTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (['onDataChange', 'runDailyAlerts'].indexOf(t.getHandlerFunction()) >= 0) ScriptApp.deleteTrigger(t);
  });
  installDailyTrigger_();
  ScriptApp.newTrigger('onDataChange').forSpreadsheet(ss_()).onEdit().create();
  ScriptApp.newTrigger('onDataChange').forSpreadsheet(ss_()).onChange().create();
}

function installDailyTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'runDailyAlerts') ScriptApp.deleteTrigger(t);
  });
  const hora = Math.min(23, Math.max(0, Number(cfg_().HORA_GATILHO) || 7));
  ScriptApp.newTrigger('runDailyAlerts').timeBased().everyDays(1).atHour(hora).inTimezone(APP.TZ).create();
}

/** Garante que a chamada veio de um gatilho real do projeto (e não de google.script.run). */
function isRealTrigger_(e) {
  if (!e || !e.triggerUid) return false;
  const uid = String(e.triggerUid);
  return ScriptApp.getProjectTriggers().some(function (t) { return t.getUniqueId() === uid; });
}

/** Gatilho diário de alertas. */
function runDailyAlerts(e) {
  if (!isRealTrigger_(e)) requireUser_(['Administrador']);
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(60000)) { console.warn('runDailyAlerts: lock ocupado'); return null; }
  try { return processAlerts_('gatilho diário'); } finally { lock.releaseLock(); }
}

/** Gatilhos onEdit/onChange da planilha: sinaliza ao portal que os dados mudaram. */
function onDataChange(e) {
  if (!isRealTrigger_(e)) return;
  bumpVersion_();
}
