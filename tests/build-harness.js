// Monta tests/out/harness.html: o portal real (src/) rodando no navegador sobre os mocks do Apps Script.
const fs = require('fs'), path = require('path');
const SRC = path.join(__dirname, '..', 'src'), OUT = path.join(__dirname, 'out');
fs.mkdirSync(OUT, { recursive: true });
const read = (f) => fs.readFileSync(path.join(SRC, f), 'utf8');
const rel = (f) => path.relative(OUT, f).split(path.sep).join('/');
fs.copyFileSync(path.join(SRC, 'Code.gs'), path.join(OUT, 'Code.gs.js'));
fs.copyFileSync(path.join(SRC, 'Logo.gs'), path.join(OUT, 'Logo.gs.js'));
const pre = '<script src="' + rel(path.join(__dirname, 'fixtures', 'seed.js')) + '"></script>' +
  '<script>window.__FILES=' + JSON.stringify({ Email: read('Email.html') }).replace(/<\//g, '<\\/') + ';</script>' +
  '<script src="' + rel(path.join(__dirname, 'gas-mock.js')) + '"></script><script src="Logo.gs.js"></script><script src="Code.gs.js"></script>' +
  '<script src="' + rel(path.join(__dirname, 'shim.js')) + '"></script>';
const html = read('Index.html')
  .replace('<?= appName ?>', 'Portal de Gestão de Contratos de Serviços em Cibersegurança')
  .replace("<?!= include('Styles'); ?>", () => read('Styles.html'))
  .replace("<?!= include('Scripts'); ?>", () => pre + read('Scripts.html'));
fs.writeFileSync(path.join(OUT, 'harness.html'), html);
console.log('harness:', path.join(OUT, 'harness.html'));
