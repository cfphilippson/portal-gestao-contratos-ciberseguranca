/* Mock mínimo dos serviços do Google Apps Script — para testes locais (Node e Chromium). */
(function (G) {
  const p2 = (n) => String(n).padStart(2, '0');
  function formatDate(d, tz, pat) {
    const t = new Date(d.getTime() - 3 * 3600e3); // America/Sao_Paulo (UTC-3, sem horário de verão)
    const v = { yyyy: t.getUTCFullYear(), MM: p2(t.getUTCMonth() + 1), dd: p2(t.getUTCDate()), HH: p2(t.getUTCHours()), mm: p2(t.getUTCMinutes()), ss: p2(t.getUTCSeconds()), u: ((t.getUTCDay() + 6) % 7) + 1 };
    return pat.replace(/yyyy|MM|dd|HH|mm|ss|u/g, (k) => v[k]);
  }
  const hydrate = (v) => (v && v.__d ? new Date(v.__d[0], v.__d[1] - 1, v.__d[2]) : v);

  class Range {
    constructor(sh, r, c, nr, nc) { Object.assign(this, { sh, r, c, nr, nc }); }
    getValues() {
      const out = [];
      for (let i = 0; i < this.nr; i++) {
        const row = this.sh.data[this.r - 1 + i] || [];
        const o = [];
        for (let j = 0; j < this.nc; j++) { const v = row[this.c - 1 + j]; o.push(v === undefined || v === null ? '' : v); }
        out.push(o);
      }
      return out;
    }
    setValues(v) {
      v.forEach((row, i) => {
        const ri = this.r - 1 + i;
        while (this.sh.data.length <= ri) this.sh.data.push([]);
        row.forEach((x, j) => { this.sh.data[ri][this.c - 1 + j] = x; });
      });
      return this;
    }
    setValue(x) { return this.setValues([[x]]); }
    setFontWeight() { return this; }
  }
  class Sheet {
    constructor(name, data) { this.name = name; this.data = data; }
    getLastColumn() { return Math.max(0, ...this.data.map((r) => { let i = r.length; while (i > 0 && (r[i - 1] === '' || r[i - 1] == null)) i--; return i; })); }
    getDataRange() { return new Range(this, 1, 1, Math.max(1, this.data.length), Math.max(1, this.getLastColumn())); }
    getRange(r, c, nr, nc) { return new Range(this, r, c, nr || 1, nc || 1); }
    appendRow(row) { this.data.push(row.slice()); }
    deleteRow(n) { this.data.splice(n - 1, 1); }
    setFrozenRows() {}
  }
  const sheets = {};
  Object.keys(G.__SEED).forEach((k) => { sheets[k] = new Sheet(k, G.__SEED[k].map((r) => r.map(hydrate))); });
  const SS = {
    getSheetByName: (n) => sheets[n] || null,
    insertSheet: (n) => (sheets[n] = new Sheet(n, [])),
    getId: () => 'SHEET-ID', getUrl: () => 'https://docs.google.com/spreadsheets/d/SHEET-ID',
    getSpreadsheetTimeZone: () => 'America/Sao_Paulo',
  };
  G.__SHEETS = sheets;
  G.__MAILS = [];
  G.__USER = G.__USER || 'admin@example.com';
  G.__OWNER = 'admin@example.com';

  const props = {}, cache = {};
  G.SpreadsheetApp = { openById: () => SS, getActiveSpreadsheet: () => SS };
  G.PropertiesService = { getScriptProperties: () => ({ getProperty: (k) => (k in props ? props[k] : null), setProperty: (k, v) => { props[k] = String(v); } }) };
  G.CacheService = { getScriptCache: () => ({ get: (k) => (k in cache ? cache[k] : null), put: (k, v) => { cache[k] = v; }, remove: (k) => { delete cache[k]; } }) };
  G.LockService = { getScriptLock: () => ({ tryLock: () => true, waitLock() {}, releaseLock() {} }) };
  G.Session = { getActiveUser: () => ({ getEmail: () => G.__USER }), getEffectiveUser: () => ({ getEmail: () => G.__OWNER }) };
  G.Utilities = { formatDate, base64Encode: (b) => btoa(String.fromCharCode.apply(null, b)),
    base64Decode: (s) => Array.from(atob(s), (c) => c.charCodeAt(0)),
    newBlob: (bytes, type, name) => ({ getBytes: () => bytes, getContentType: () => type, getName: () => name, setName() { return this; } }) };
  G.MailApp = { sendEmail: (o) => { G.__MAILS.push(o); }, getRemainingDailyQuota: () => 1500 };
  const triggers = [];
  let uid = 1;
  const builder = (fn) => {
    const t = { fn, id: 'T' + uid++ };
    const chain = { timeBased: () => chain, everyDays: () => chain, atHour: (h) => { t.hour = h; return chain; }, inTimezone: () => chain, forSpreadsheet: () => chain, onEdit: () => { t.kind = 'edit'; return chain; }, onChange: () => { t.kind = 'change'; return chain; }, create: () => { triggers.push(t); return t; } };
    return chain;
  };
  G.__TRIGGERS = triggers;
  G.ScriptApp = {
    newTrigger: builder,
    getProjectTriggers: () => triggers.map((t) => ({ getHandlerFunction: () => t.fn, getUniqueId: () => t.id, _t: t })),
    deleteTrigger: (x) => { const i = triggers.indexOf(x._t); if (i >= 0) triggers.splice(i, 1); },
    getService: () => ({ getUrl: () => 'https://script.google.com/a/macros/example.com/s/DEPLOY/exec' }),
  };
  G.DriveApp = { getFileById: () => { throw new Error('sem logo no mock'); } };

  function compile(src) {
    let code = 'let __o="";\n';
    const re = /<\?(!?=)?([\s\S]*?)\?>/g;
    let last = 0, m;
    while ((m = re.exec(src))) {
      code += '__o+=' + JSON.stringify(src.slice(last, m.index)) + ';\n';
      if (m[1] === '=') code += '__o+=__esc(' + m[2] + ');\n';
      else if (m[1] === '!=') code += '__o+=(' + m[2] + ');\n';
      else code += m[2] + '\n';
      last = re.lastIndex;
    }
    return code + '__o+=' + JSON.stringify(src.slice(last)) + ';return __o;';
  }
  const esc = (s) => String(s == null ? '' : s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
  G.__FILES = G.__FILES || {};
  G.HtmlService = {
    XFrameOptionsMode: { DEFAULT: 'DEFAULT' },
    createHtmlOutputFromFile: (n) => ({ getContent: () => G.__FILES[n] }),
    createTemplateFromFile: (n) => {
      const t = {};
      t.evaluate = () => {
        const fn = new Function('__esc', '__ctx', 'with(__ctx){' + compile(G.__FILES[n]) + '}');
        const html = fn(esc, t);
        return { getContent: () => html, setTitle() { return this; }, addMetaTag() { return this; }, setXFrameOptionsMode() { return this; } };
      };
      return t;
    },
  };
})(typeof window !== 'undefined' ? window : globalThis);
