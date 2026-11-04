'use strict';
const fs = require('fs');
const path = require('path');
const { createRuntime } = require('./gas_mock');

const ROOT = path.resolve(__dirname, '..');
const DEMO_LOGINS = ['admin.demo', 'auditor.demo', 'operator.demo'];

const demoPassword = (login) => 'demo-' + login.toLowerCase();

function loadPortal(options = {}) {
  const properties = {};
  if (options.seedPasswords !== false) {
    DEMO_LOGINS.forEach((login) => { properties['PORTAL_QUALIDADE_SEED_' + login.toUpperCase()] = demoPassword(login); });
  }
  const rt = createRuntime({
    files: [path.join(ROOT, 'apps-script', 'Codigo.gs')],
    htmlFiles: { Index: path.join(ROOT, 'apps-script', 'Index.html') },
    now: options.now,
    properties,
  });
  rt.evaluate('instalar()');
  const workbook = rt.workbooks.get(rt.properties['PORTAL_QUALIDADE_PLANILHA_ID']);
  if (options.seedNews !== false) {
    const seed = JSON.parse(fs.readFileSync(path.join(__dirname, 'demo_seed.json'), 'utf8'));
    const sheet = workbook.getSheetByName('Novidades');
    seed.novidades.forEach((n) => sheet.appendRow([n.id, n.titulo, n.corpo, new Date(n.data + 'T12:00:00'), n.setorRelacionado, n.criadoPor]));
  }
  return Object.assign(rt, { workbook });
}

module.exports = { loadPortal, demoPassword, DEMO_LOGINS, ROOT };
