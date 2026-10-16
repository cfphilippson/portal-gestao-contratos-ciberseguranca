(function () {
  const owner = 'admin@example.com';
  window.__USER = owner; setup();
  const as = (u, f) => { window.__USER = u; _cfg = null; f(); };
  as('gestor.infra@example.com', () => submitAccessRequest({ nome: 'Gestor de Infraestrutura', area: 'Infraestrutura de TI', perfil: 'Gestor', justificativa: 'Responsável pelos contratos de DLP e proxy web.' }));
  as('analista.compras@example.com', () => submitAccessRequest({ nome: 'Analista de Compras', area: 'Suprimentos', perfil: 'Leitura', justificativa: 'Acompanhar renovações para planejamento de compras.' }));
  as('gestor.endpoint@example.com', () => submitAccessRequest({ nome: 'Gestor de Endpoint', area: 'Cyber', perfil: 'Gestor', justificativa: 'Gestão dos contratos de endpoint.' }));
  as(owner, () => { const p = getAdminData().pendentes.find((r) => r.email.startsWith('gestor.endpoint')); decideRequest(p.id, 'aprovar', 'Gestor', 'CT-001,CT-002', ''); processAlerts_('gatilho diário'); });
  const u = new URLSearchParams(location.search).get('user');
  window.__USER = u || owner;
  const mk = (ok, fail) => new Proxy({}, { get(_, k) {
    if (k === 'withSuccessHandler') return (f) => mk(f, fail);
    if (k === 'withFailureHandler') return (f) => mk(ok, f);
    return (...args) => setTimeout(() => {
      try { if (String(k).endsWith('_')) throw new Error('privado'); _cfg = null; _ss = null; _tz = null;
        const r = window[k](...args); ok && ok(r === undefined ? r : JSON.parse(JSON.stringify(r)));
      } catch (e) { fail && fail(e); }
    }, 60);
  } });
  window.google = { script: { run: mk(null, null) } };
})();
