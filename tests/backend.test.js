// Testes do backend (src/Code.gs) com mocks dos serviços do Apps Script.  node tests/backend.test.js
process.env.TZ = 'America/Sao_Paulo';
const fs = require('fs'), vm = require('vm'), assert = require('assert'), path = require('path');
const OUT = path.join(__dirname, '..', 'src') + '/';
const ART = path.join(__dirname, 'out'); fs.mkdirSync(ART, { recursive: true });
let NOW = Date.parse('2026-10-16T10:00:00-03:00');
const RealDate = Date;
class FixedDate extends RealDate { constructor(...a) { if (a.length) super(...a); else super(NOW); } static now() { return NOW; } }

const ctx = { console, atob: (s) => Buffer.from(s, 'base64').toString('binary'), btoa: (s) => Buffer.from(s, 'binary').toString('base64'), Date: FixedDate, JSON, Math };
ctx.window = ctx; ctx.globalThis = ctx;
vm.createContext(ctx);
const load = (f) => vm.runInContext(fs.readFileSync(f, 'utf8'), ctx, { filename: f });
load(path.join(__dirname, 'fixtures', 'seed.js'));
ctx.__FILES = { Email: fs.readFileSync(OUT + 'Email.html', 'utf8'), Styles: '', Scripts: '' };
load(path.join(__dirname, 'gas-mock.js'));
load(OUT + 'Logo.gs');
load(OUT + 'Code.gs');

const as = (u) => { ctx.__USER = u; };
const call = (fn, ...a) => {
  vm.runInContext('_cfg=null;_ss=null;_tz=null;', ctx); // cada chamada = execução nova no GAS
  const r = ctx[fn](...a);
  const hasDate = (o) => o instanceof RealDate || o instanceof FixedDate || (o && typeof o === 'object' && Object.values(o).some(hasDate));
  assert(!hasDate(r), fn + ' retornou Date (ilegal em google.script.run)');
  return r === undefined ? r : JSON.parse(JSON.stringify(r));
};
const throws = (fn, re, ...a) => { try { call(fn, ...a); } catch (e) { assert(re.test(e.message), fn + ': erro inesperado ' + e.message); return; } assert.fail(fn + ' deveria falhar'); };
let n = 0; const ok = (m) => console.log('  ✓', m, ++n && '');

const ADM = 'admin@example.com', GI = 'gestor.infra@example.com', EXT = 'fulano@gmail.com';

// setup
as(ADM); call('setup');
assert.equal(ctx.__TRIGGERS.length, 3); ok('setup cria 3 gatilhos (diário + onEdit + onChange)');
assert.equal(ctx.__SHEETS.Usuarios.data.length, 2); ok('setup cria administrador inicial');
call('setup'); assert.equal(ctx.__TRIGGERS.length, 3); assert.equal(ctx.__SHEETS.Usuarios.data.length, 2); ok('setup é idempotente');

// bootstrap/admin
let b = call('getBootstrap'); assert.equal(b.user.status, 'ativo'); assert.equal(b.user.perfil, 'Administrador'); ok('admin entra como ativo');
let d = call('getData');
assert.equal(d.contratos.length, 16); assert.equal(d.today, '2026-10-16');
const fx = (id) => d.contratos.find((c) => c.id === id);
assert.equal(fx('CT-013').dias, 55); assert.equal(fx('CT-013').faixa, 'j60'); ok('DLP: 55 dias → faixa j60');
assert.equal(d.contratos.filter((c) => c.faixa === 'vencido').length, 5); ok('5 vencidos');
assert.equal(fx('CT-015').faixa, 'sem_data'); ok('Proteção de API sem data');
assert.equal(d.proximos[0].id, 'CT-013'); assert.equal(d.proximos[0].janela, 30); assert.equal(d.proximos[0].data, '2026-11-10'); ok('próximo disparo: DLP 30d em 10/11/2026');

// sem acesso
as(GI); b = call('getBootstrap'); assert.equal(b.user.status, 'sem_acesso'); ok('usuário sem cadastro → sem_acesso');
throws('getData', /ACESSO_NEGADO/); throws('getVersion', /ACESSO_NEGADO/); throws('getAdminData', /ACESSO_NEGADO/); ok('sem acesso não lê dados');
throws('submitAccessRequest', /justificativa/, { nome: 'Gestor Infra', area: 'Infra', perfil: 'Gestor', justificativa: 'curta' });
b = call('submitAccessRequest', { nome: '=HYPERLINK("x")', area: 'Infra', perfil: 'Gestor', justificativa: 'Responsável pelos contratos de DLP' });
assert.equal(b.user.status, 'pendente'); ok('solicitação → pendente');
assert(ctx.__SHEETS.Solicitacoes.data[1][2].startsWith("'=")); ok('injeção de fórmula neutralizada');
assert(ctx.__MAILS.some((m) => m.to === ADM && /Nova solicitação/.test(m.subject))); ok('admins notificados');
throws('submitAccessRequest', /pendente/, { nome: 'Gestor Infra', area: 'Infra', perfil: 'Gestor', justificativa: 'Responsável pelos contratos de DLP' }); ok('não duplica pendente');
as(EXT); throws('submitAccessRequest', /example\.com/, { nome: 'Fulano', area: 'X', perfil: 'Leitura', justificativa: 'quero acesso ao portal' }); ok('domínio externo bloqueado');

// privado/trigger
as(GI); throws('runDailyAlerts', /ACESSO_NEGADO/, { triggerUid: 'FAKE' }); ok('runDailyAlerts com triggerUid forjado exige admin');
assert(typeof ctx.processAlerts_ === 'function'); // privado no GAS por convenção "_"

// aprovação
as(ADM); let a = call('getAdminData'); assert.equal(a.pendentes.length, 1);
a = call('decideRequest', a.pendentes[0].id, 'aprovar', 'Gestor', 'CT-013,ct-012, x;y', '');
assert.equal(a.usuarios.length, 2); const gi = a.usuarios.find((u) => u.email === GI); assert.equal(gi.escopo, 'CT-013,CT-012,X,Y'); ok('aprovação cria usuário com escopo normalizado');
throws('decideRequest', /já foi decidida/, a.decididas[0].id, 'recusar', '', '', ''); ok('não decide duas vezes');
as(GI); d = call('getData'); assert.deepEqual(d.contratos.map((c) => c.id).sort(), ['CT-012', 'CT-013']); ok('gestor vê só o escopo');
throws('saveContract', /Permissão/, { parceiro: 'X', ferramenta: 'Y' }); throws('removeUser', /Permissão/, ADM); ok('gestor não edita nem remove');

// remoção de usuário
as(ADM); throws('removeUser', /próprio acesso/, ADM); ok('admin não remove a si mesmo');
throws('updateUser', /último administrador/, ADM, 'Leitura', '*'); ok('não rebaixa o último admin');
a = call('removeUser', GI); assert.equal(a.usuarios.length, 1); ok('remove usuário');
assert(a.auditoria[0].acao === 'Removeu usuário'); ok('remoção auditada');
as(GI); throws('getData', /ACESSO_NEGADO/); throws('getVersion', /ACESSO_NEGADO/); b = call('getBootstrap'); assert.equal(b.user.status, 'sem_acesso'); ok('removido perde acesso imediatamente (inclusive no polling)');

// versão / auto-refresh
as(ADM); const v1 = call('getVersion'); NOW += 1000;
call('saveContract', { id: 'CT-013', parceiro: 'Muralha Dados', ferramenta: 'DLP', finalidade: 'Prevenção contra vazamento de dados', tipo_servico: 'Suporte 5x8', sla: 'Plano de Governança', indice_reajuste: 'IPCA', valor_suporte: 'R$ 25.000,00', valor_pago_anterior: '230000', renovacao_tipo: 'Renovação', vigencia_inicio: '2025-12-10', vigencia_fim: '2026-12-10', gestor_email: 'gestor.infra@example.com', gestor_nome: 'Gestor Infra', valor_orcado: '251234.50', status_renovacao: 'Em vigor' });
const v2 = call('getVersion'); assert.notEqual(v1, v2); ok('escrita muda a versão (auto-refresh detecta)');
NOW += 1000; ctx.onDataChange({ triggerUid: ctx.__TRIGGERS.find((t) => t.kind === 'edit').id }); assert.notEqual(call('getVersion'), v2); ok('edição direta na planilha (onEdit) muda a versão');
const v3 = call('getVersion'); ctx.onDataChange({ triggerUid: 'FAKE' }); assert.equal(call('getVersion'), v3); ok('onDataChange ignora chamada forjada');
d = call('getData'); assert.equal(d.contratos.find((c) => c.id === 'CT-013').valor_orcado, 251234.5); ok('valor numérico com ponto decimal');
throws('saveContract', /anterior ao início/, { parceiro: 'X', ferramenta: 'Y', vigencia_inicio: '2026-01-02', vigencia_fim: '2026-01-01' });
throws('saveContract', /Data inválida/, { parceiro: 'X', ferramenta: 'Y', vigencia_fim: '2026-02-31' }); ok('validação de datas');

// alertas (sexta-feira 16/10/2026 — não é dia de reenvio de vencidos)
ctx.__MAILS.length = 0;
let r = call('runDailyAlerts', { triggerUid: ctx.__TRIGGERS.find((t) => t.fn === 'runDailyAlerts').id });
assert.equal(r.enviados, 1); const m = ctx.__MAILS[0];
assert.equal(m.to, 'gestor.infra@example.com'); assert.equal(m.cc, ADM); assert(/Janela de 60 dias · DLP/.test(m.subject)); ok('envia 60d do DLP ao gestor com cópia aos admins');
assert(/55/.test(m.htmlBody) && /10\/11\/2026/.test(m.htmlBody)); ok('e-mail mostra dias reais e próximo aviso');
assert(!m.inlineImages); ok('sem logo configurado: e-mail sem imagem inline');
assert.equal(call('getBootstrap').app.logo, null); ok('sem logo configurado: portal usa o ícone padrão');
fs.writeFileSync(path.join(ART, 'email-sample.html'), m.htmlBody);
r = call('runDailyAlerts', { triggerUid: ctx.__TRIGGERS.find((t) => t.fn === 'runDailyAlerts').id }); assert.equal(r.enviados, 0); ok('idempotente: não reenvia a mesma janela');
// avança para segunda 19/10 → reenvio de vencidos
NOW = Date.parse('2026-10-19T07:00:00-03:00'); ctx.__MAILS.length = 0;
r = call('runAlertsNow'); assert.equal(r.enviados, 5); assert(ctx.__MAILS.every((x) => /VENCIDO/.test(x.subject))); ok('segunda-feira: reenvio semanal dos 5 vencidos');
r = call('runAlertsNow'); assert.equal(r.enviados, 0); ok('reenvio de vencidos 1x por semana');
// avança para 10/11 → janela de 30 do DLP
NOW = Date.parse('2026-11-10T07:00:00-03:00'); ctx.__MAILS.length = 0;
r = call('runAlertsNow'); assert(ctx.__MAILS.some((x) => /Janela de 30 dias · DLP/.test(x.subject))); ok('10/11: janela de 30 dias do DLP');
// renovação reinicia ciclo
NOW = Date.parse('2026-11-11T07:00:00-03:00');
d = call('registerRenewal', 'CT-013', '2026-12-11', '2027-12-10', '', 'renovado 12 meses');
const dlp = d.contratos.find((c) => c.id === 'CT-013'); assert.equal(dlp.faixa, 'ok'); assert(/renovação\] renovado/.test(dlp.observacoes)); ok('renovação atualiza vigência e anexa observação');
// contrato novo já dentro de 20 dias → só janela de 30
call('saveContract', { parceiro: 'Teste', ferramenta: 'Nova', vigencia_inicio: '2026-01-01', vigencia_fim: '2026-11-22' });
ctx.__MAILS.length = 0; r = call('runAlertsNow');
const novo = ctx.__MAILS.filter((x) => /Nova/.test(x.subject)); assert.equal(novo.length, 1); assert(/Janela de 30 dias/.test(novo[0].subject)); ok('contrato que entra com 20 dias recebe só a janela de 30');
assert.equal(novo[0].to, ADM); ok('sem gestor → envia aos admins');

// config
d = call('saveAlertConfig', { janelas: [120, 90, 60, 30], notificarAdmins: false, ccFixo: 'dl-cyber@example.com', reenvioVencidos: true, diaReenvio: 2, horaGatilho: 8, autoRefreshSeg: 30 });
assert.deepEqual(d.janelas, [120, 90, 60, 30]); assert.equal(d.regras.autoRefreshSeg, 30);
assert.equal(ctx.__TRIGGERS.find((t) => t.fn === 'runDailyAlerts').hour, 8); ok('regras salvas e gatilho reinstalado no novo horário');
throws('saveAlertConfig', /ao menos uma janela/, { janelas: [], diaReenvio: 1, horaGatilho: 7, autoRefreshSeg: 60 }); ok('valida regras');
assert(/teste/.test(call('sendTestAlert', 'CT-002'))); assert(/\[TESTE\]/.test(ctx.__MAILS[ctx.__MAILS.length - 1].subject)); ok('e-mail de teste');
call('setTheme', 'dark'); assert.equal(call('getBootstrap').user.tema, 'dark'); ok('tema salvo por usuário');
d = call('deleteContract', 'CT-017'); assert.equal(d.contratos.length, 16); ok('exclusão de contrato');
console.log('\n' + n + ' verificações OK');
