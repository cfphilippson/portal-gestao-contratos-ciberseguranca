// Testes de interface no Chromium (Playwright).  npm run test:ui   (SHOTS=docs/screenshots para regenerar as imagens do README)
const { chromium } = require('playwright');
const path = require('path'), fs = require('fs');
const URL = 'file://' + path.join(__dirname, 'out', 'harness.html');
const SHOTS = process.env.SHOTS ? path.resolve(process.env.SHOTS) : path.join(__dirname, 'out', 'shots');
fs.mkdirSync(SHOTS, { recursive: true });
(async () => {
  const browser = await chromium.launch();
  const errors = [];
  const mkPage = async (opts = {}) => {
    const ctx = await browser.newContext({ viewport: { width: opts.w || 1440, height: 900 }, timezoneId: 'America/Sao_Paulo', colorScheme: opts.scheme || 'light' });
    const page = await ctx.newPage();
    await page.clock.install({ time: new Date('2026-10-16T10:00:00-03:00') });
    page.on('console', (m) => { if (m.type() === 'error') errors.push(m.text()); });
    page.on('pageerror', (e) => errors.push('pageerror: ' + e.message));
    return page;
  };
  const shot = (p, n, full = true) => p.screenshot({ path: path.join(SHOTS, n + '.png'), fullPage: full });
  const step = (m) => console.log('  ✓', m);

  // ---- Admin
  let p = await mkPage();
  await p.goto(URL);
  await p.waitForSelector('.kpi');
  await p.evaluate(() => { try { localStorage.clear(); } catch (e) {} });
  await shot(p, '01-dashboard-claro'); step('dashboard (claro)');
  const w0 = await p.$eval('#sidebar', (e) => e.getBoundingClientRect().width);
  await p.click('[data-act="toggle-sidebar"]'); await p.waitForTimeout(400);
  const w1 = await p.$eval('#sidebar', (e) => e.getBoundingClientRect().width);
  if (!(w0 > 240 && w1 < 80)) throw new Error('sidebar ' + w0 + '→' + w1);
  if (await p.getAttribute('[data-act="toggle-sidebar"]', 'aria-expanded') !== 'false') throw new Error('aria-expanded');
  await shot(p, '01b-sidebar-recolhida', false);
  await p.reload(); await p.waitForSelector('.kpi');
  if ((await p.$eval('#sidebar', (e) => e.getBoundingClientRect().width)) > 80) throw new Error('estado da sidebar não persistiu');
  await p.click('[data-act="toggle-sidebar"]'); await p.waitForTimeout(400);
  step('sidebar recolhe/expande e lembra o estado');
  const title = await p.textContent('.sb-name'); if (!/Portal de Gestão de Contratos de Serviços em Cibersegurança/.test(title)) throw new Error('nome do portal');
  await p.click('[data-act="theme"][data-v="dark"]');
  if (await p.getAttribute('html', 'data-theme') !== 'dark') throw new Error('tema escuro');
  await shot(p, '02-dashboard-escuro'); step('tema escuro');
  await p.click('[data-act="dash-faixa"][data-v="vencido"]');
  const tlRows = await p.$$eval('.tl-row .tl-track', (x) => x.length); if (tlRows !== 5) throw new Error('filtro timeline ' + tlRows); step('filtro por faixa na linha do tempo');
  await p.click('[data-act="nav"][data-v="contratos"]'); await p.waitForSelector('table.tbl');
  await p.fill('#f-q', 'dlp'); await p.waitForTimeout(400);
  const rows = await p.$$eval('table.tbl tbody tr', (x) => x.length); if (rows !== 1) throw new Error('busca ' + rows);
  await p.fill('#f-q', ''); await p.waitForTimeout(400);
  await p.click('button.linkish:has-text("DLP")');
  await shot(p, '03-contratos-escuro'); step('contratos + busca + detalhe');
  await p.click('[data-act="renew"]'); await p.waitForSelector('#f-renew'); await shot(p, '04-modal-renovacao', false);
  await p.click('#f-renew [data-act="modal-close"]');
  await p.click('[data-act="edit-contract"]'); await p.fill('#c-gestor_nome', 'Gestor de Infraestrutura'); await p.fill('#c-gestor_email', 'gestor.infra@example.com');
  await p.click('#f-contract [type="submit"]'); await p.waitForSelector('#f-contract', { state: 'detached' });
  if (!(await p.textContent('.side')).includes('gestor.infra@example.com')) throw new Error('edição'); step('edição de contrato');
  await p.click('[data-act="theme"][data-v="light"]');
  await p.click('[data-act="nav"][data-v="alertas"]'); await p.waitForSelector('#f-regras');
  await shot(p, '05-alertas-claro'); step('alertas');
  await p.click('[data-act="nav"][data-v="acessos"]'); await p.waitForSelector('.req');
  await shot(p, '06-acessos-claro'); step('acessos');
  const before = await p.$$eval('[data-act="remove-user"]', (x) => x.length);
  const selfDisabled = await p.$eval('[data-act="remove-user"][data-email="admin@example.com"]', (b) => b.disabled); if (!selfDisabled) throw new Error('self remove habilitado');
  await p.click('[data-act="remove-user"][data-email="gestor.endpoint@example.com"]'); await p.waitForSelector('[data-act="confirm-yes"]');
  await shot(p, '07-confirmar-remocao', false);
  await p.click('[data-act="confirm-yes"]'); await p.waitForFunction((n) => document.querySelectorAll('[data-act="remove-user"]').length === n - 1, before);
  step('remover usuário (com confirmação; próprio usuário bloqueado)');
  await p.click('[data-act="approve"] >> nth=0'); await p.waitForFunction(() => document.querySelectorAll('.req').length === 1); step('aprovar solicitação');

  // ---- Auto-refresh: alteração direta na planilha
  await p.click('[data-act="nav"][data-v="dashboard"]');
  await p.evaluate(() => { const sh = window.__SHEETS.Contratos; const r = sh.data.find((x) => x[0] === 'CT-002'); r[3] = 'EDR Pro'; window.onDataChange({ triggerUid: window.__TRIGGERS.find((t) => t.kind === 'edit').id }); });
  await p.clock.runFor(61000); await p.waitForTimeout(500);
  const toast = await p.textContent('#toast');
  const html = await p.content();
  if (!/automaticamente/.test(toast) || !html.includes('EDR Pro')) throw new Error('auto-refresh: ' + toast);
  step('auto-refresh detecta edição na planilha e re-renderiza');
  await p.click('[data-act="toggle-auto"]');
  await p.evaluate(() => { const r = window.__SHEETS.Contratos.data.find((x) => x[0] === 'CT-002'); r[3] = 'EDR OFF'; window.onDataChange({ triggerUid: window.__TRIGGERS.find((t) => t.kind === 'edit').id }); });
  await p.clock.runFor(120000); await p.waitForTimeout(300);
  if ((await p.content()).includes('EDR OFF')) throw new Error('auto off ainda atualiza');
  await p.click('[data-act="refresh"]'); await p.waitForTimeout(500);
  if (!(await p.content()).includes('EDR OFF')) throw new Error('refresh manual'); step('auto desligado não atualiza; botão manual atualiza');

  // ---- Tema "sistema" segue o SO
  const p2 = await mkPage({ scheme: 'dark' });
  await p2.goto(URL); await p2.waitForSelector('.kpi');
  if (await p2.getAttribute('html', 'data-theme') !== 'dark') throw new Error('sistema→escuro'); step('tema "sistema" segue preferência do SO');

  // ---- Usuário sem acesso
  const g = await mkPage();
  await g.goto(URL + '?user=novo.usuario@example.com'); await g.waitForSelector('#f-req');
  await shot(g, '08-sem-acesso', false);
  await g.fill('#rq-nome', 'Novo Usuário'); await g.fill('#rq-area', 'TI'); await g.fill('#rq-just', 'Preciso consultar os contratos da minha área.');
  await g.click('#f-req [type="submit"]'); await g.waitForSelector('.steps');
  await shot(g, '09-pendente', false); step('solicitação de acesso → pendente');
  await g.evaluate(() => { const own = 'admin@example.com', me = window.__USER; window.__USER = own; _cfg = null; const r = getAdminData().pendentes.find((x) => x.email === me); decideRequest(r.id, 'aprovar', 'Leitura', '*', ''); window.__USER = me; });
  await g.clock.runFor(61000); await g.waitForSelector('.kpi', { timeout: 5000 }); step('pendente entra automaticamente após aprovação');

  // ---- Mobile
  const m = await mkPage({ w: 390 });
  await m.goto(URL); await m.waitForSelector('.kpi');
  const overflow = await m.evaluate(() => document.documentElement.scrollWidth > window.innerWidth + 1);
  const off = await m.$eval('#sidebar', (e) => e.getBoundingClientRect().right);
  if (off > 1) throw new Error('gaveta visível no mobile');
  await m.click('[data-act="open-drawer"]'); await m.waitForTimeout(400);
  await shot(m, '10b-mobile-gaveta', false);
  await m.click('.sb-item[data-v="contratos"]'); await m.waitForSelector('table.tbl'); await m.waitForTimeout(400);
  if ((await m.$eval('#sidebar', (e) => e.getBoundingClientRect().right)) > 1) throw new Error('gaveta não fechou');
  await m.click('[data-act="nav"][data-v="dashboard"]').catch(() => {});
  step('mobile: menu vira gaveta e fecha ao navegar');
  await shot(m, '10-mobile');
  if (overflow) throw new Error('scroll horizontal no mobile'); step('mobile 390px sem scroll horizontal');

  // ---- e-mail
  const e = await mkPage({ w: 700 }); await e.goto('file://' + path.join(__dirname, 'out', 'email-sample.html')); await shot(e, '11-email');
  console.log(errors.length ? 'ERROS:\n' + errors.join('\n') : '\nSem erros de console.');
  await browser.close();
})().catch((e) => { console.error('FALHA:', e.message); process.exit(1); });
