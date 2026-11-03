'use strict';
// Run with: node --test tests/portal.test.js     (Node 18+, no dependencies)
// Executes the real apps-script/Codigo.gs inside an in-memory mock of the Apps Script
// services (demo/gas_mock.js), and validates the structure of the demo content in Index.html.
const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { loadPortal, demoPassword } = require('../demo/runtime');

const NOW = '2026-01-15T10:00:00-03:00';
const DEVICE = 'device-token-0001';
const plain = (x) => JSON.parse(JSON.stringify(x));
const fresh = (o) => loadPortal(Object.assign({ now: NOW }, o));
function login(rt, user) {
  const r = rt.context.verificarLogin(user, demoPassword(user), DEVICE);
  assert.equal(r.ok, true, 'login ' + user);
  return r.token;
}

/* ------------------------------------------------------------------ install and accounts */
test('install creates the sheets and the three seeded accounts from Script Properties', () => {
  const rt = fresh();
  const names = rt.workbook.getSheets().map((s) => s.getName()).sort();
  assert.deepEqual(names, ['ConteudoEditado', 'LogAcessos', 'Novidades', 'TentativasLogin', 'Usuarios']);
  const users = rt.workbook.getSheetByName('Usuarios').data.slice(1).map((r) => [r[0], r[3]]);
  assert.deepEqual(users, [['admin.demo', 'Admin'], ['auditor.demo', 'Auditor'], ['operator.demo', 'Operador']]);
});

test('install is idempotent and never creates accounts without a seed password', () => {
  const rt = fresh({ seedPasswords: false, seedNews: false });
  assert.equal(rt.workbook.getSheetByName('Usuarios').data.length, 1, 'header only');
  const before = rt.workbooks.size;
  rt.evaluate('instalar()');
  assert.equal(rt.workbooks.size, before, 'second run does not create another spreadsheet');
});

test('password hash is salted: same password, different salts, different hashes', () => {
  const rt = fresh();
  const a = rt.context.hashSenha_('same-password', 'salt-1');
  const b = rt.context.hashSenha_('same-password', 'salt-2');
  assert.notEqual(a, b);
  assert.match(a, /^[0-9a-f]{64}$/);
  assert.equal(a, rt.context.hashSenha_('same-password', 'salt-1'));
});

test('login: generic error for wrong password and unknown user', () => {
  const rt = fresh();
  const wrongPwd = plain(rt.context.verificarLogin('admin.demo', 'wrong-pass', DEVICE));
  const unknown = plain(rt.context.verificarLogin('nobody.here', 'wrong-pass', DEVICE));
  assert.equal(wrongPwd.ok, false);
  assert.equal(wrongPwd.mensagem, unknown.mensagem);
  assert.equal(rt.context.verificarLogin('admin.demo', demoPassword('admin.demo'), DEVICE).patente, 'Admin');
});

test('input validation blocks emoji, markup characters, odd logins and short passwords', () => {
  const rt = fresh();
  const v = (l, s) => rt.context.entradaValida_(l, s);
  assert.equal(v('ok.login', 'secret1'), true);
  assert.equal(v('has space', 'secret1'), false);
  assert.equal(v('ok.login', 'x<y>z'), false);
  assert.equal(v('ok.login', 'pw\u{1F680}pw'), false);
  assert.equal(v('ok.login', 'ab'), false);
  assert.equal(v('ok.login', 'a'.repeat(101)), false);
});

test('lockout escalates by cumulative failures and releases with the clock', () => {
  const rt = fresh();
  const d = rt.context.duracaoBloqueioMinutos_;
  assert.deepEqual([4, 5, 10, 15, 20, 99].map(d), [0, 2, 60, 360, 2880, 2880]);
  for (let i = 0; i < 4; i++) rt.context.verificarLogin('auditor.demo', 'bad-pass', DEVICE);
  const locked = plain(rt.context.verificarLogin('auditor.demo', 'bad-pass', DEVICE));
  assert.equal(locked.bloqueado, true);
  assert.match(locked.mensagem, /Muitas tentativas/);
  // a different browser is not locked by the device counter, but the login counter still applies
  assert.equal(rt.context.verificarLogin('auditor.demo', demoPassword('auditor.demo'), 'other-device-0002').bloqueado, true);
  rt.setNow(Date.parse(NOW) + 3 * 60 * 1000);
  assert.equal(rt.context.verificarLogin('auditor.demo', demoPassword('auditor.demo'), DEVICE).ok, true);
});

test('device lockout follows the browser token across different logins', () => {
  const rt = fresh();
  const names = ['a.one', 'b.two', 'c.three', 'd.four', 'e.five'];
  names.forEach((n) => rt.context.verificarLogin(n, 'bad-pass', 'shared-device-01'));
  const r = plain(rt.context.verificarLogin('admin.demo', demoPassword('admin.demo'), 'shared-device-01'));
  assert.equal(r.bloqueado, true, 'five failures from one browser lock that browser');
});

test('session token cannot be forged and expires after 8 hours', () => {
  const rt = fresh();
  const token = login(rt, 'operator.demo');
  assert.equal(plain(rt.context.validarTokenSessao(token)).patente, 'Operador');
  const sig = token.split('.')[1];
  const forgedPayload = Buffer.from('operator.demo|Admin|99999999999999').toString('base64').replace(/\+/g, '-').replace(/\//g, '_');
  assert.equal(rt.context.validarTokenSessao(forgedPayload + '.' + sig).ok, false);
  assert.equal(rt.context.validarTokenSessao('').ok, false);
  rt.setNow(Date.parse(NOW) + 8 * 3600 * 1000 + 1000);
  assert.equal(plain(rt.context.validarTokenSessao(token)).expirado, true);
});

test('password reset invalidates the old password immediately (cache cleared) and rename keeps the hash', () => {
  const rt = fresh();
  assert.equal(rt.context.verificarLogin('operator.demo', demoPassword('operator.demo'), DEVICE).ok, true); // warms the users cache
  assert.equal(rt.context.redefinirSenha_('operator.demo', 'brand-new-pass'), true);
  assert.equal(rt.context.verificarLogin('operator.demo', demoPassword('operator.demo'), 'dev-aaaaaaaa').ok, false);
  assert.equal(rt.context.verificarLogin('operator.demo', 'brand-new-pass', 'dev-bbbbbbbb').ok, true);
  assert.equal(rt.context.renomearLogin_('operator.demo', 'operator.two'), true);
  assert.equal(rt.context.verificarLogin('operator.two', 'brand-new-pass', 'dev-cccccccc').ok, true);
  assert.equal(rt.context.redefinirSenha_('ghost', 'whatever'), false);
});

/* ------------------------------------------------------------------ content editing */
const CARD = 'script|0|bloco|1';
const payload = JSON.stringify({ titulo: 'Sondagem revisada', itens: [{ rotulo: 'Atendente', texto: 'Texto novo' }] });

test('only Admin can write content: Auditor and Operador are refused and nothing is stored', () => {
  const rt = fresh();
  const aud = login(rt, 'auditor.demo');
  const op = login(rt, 'operator.demo');
  [aud, op, 'garbage'].forEach((tok) => {
    assert.equal(rt.context.salvarConteudoEditado(CARD, 'bloco', payload, tok, 'Sondagem', 'vendas').ok, false);
    assert.equal(rt.context.adicionarConteudoNovo('novo-1', 'bloco', payload, tok, 'x', 'vendas').ok, false);
    assert.equal(rt.context.excluirConteudoEditado(CARD, 'bloco', tok, 'x', 'vendas').ok, false);
    assert.equal(rt.context.salvarOrdemCartoes('script|0|bloco', '["a"]', tok).ok, false);
  });
  assert.equal(rt.workbook.getSheetByName('ConteudoEditado').data.length, 1, 'header only');
});

test('admin edit is stored as a diff, overwrites on second save and publishes a news item', () => {
  const rt = fresh({ seedNews: false });
  const admin = login(rt, 'admin.demo');
  assert.equal(rt.context.salvarConteudoEditado(CARD, 'bloco', payload, admin, 'Sondagem', 'vendas').ok, true);
  assert.equal(rt.context.salvarConteudoEditado(CARD, 'bloco', payload.replace('novo', 'mais novo'), admin, 'Sondagem', 'vendas').ok, true);
  const data = plain(rt.context.carregarConteudoEditado());
  assert.equal(data.overrides.length, 1, 'same id is updated in place, not duplicated');
  assert.match(data.overrides[0].json, /mais novo/);
  assert.equal(data.overrides[0].excluido, false);
  assert.equal(data.novidades.length, 2);
  assert.match(data.novidades[0].corpo, /admin\.demo/);
});

test('delete is soft with undo: the news item is only published once the removal is confirmed', () => {
  const rt = fresh({ seedNews: false });
  const admin = login(rt, 'admin.demo');
  rt.context.salvarConteudoEditado(CARD, 'bloco', payload, admin, 'Sondagem', 'vendas');
  const newsAfterSave = plain(rt.context.carregarConteudoEditado()).novidades.length;

  assert.equal(rt.context.excluirConteudoEditado(CARD, 'bloco', admin, 'Sondagem', 'vendas').ok, true);
  assert.equal(plain(rt.context.carregarConteudoEditado()).overrides[0].excluido, true);
  assert.equal(plain(rt.context.carregarConteudoEditado()).novidades.length, newsAfterSave, 'no news yet');

  // undo within the window: restored and still no news
  assert.equal(rt.context.restaurarConteudoEditado(CARD, admin).ok, true);
  rt.context.confirmarExclusaoDefinitiva(CARD, 'Sondagem', 'vendas', admin);
  assert.equal(plain(rt.context.carregarConteudoEditado()).overrides[0].excluido, false);
  assert.equal(plain(rt.context.carregarConteudoEditado()).novidades.length, newsAfterSave, 'undo never published a removal');

  // delete again and let the window pass
  rt.context.excluirConteudoEditado(CARD, 'bloco', admin, 'Sondagem', 'vendas');
  rt.context.confirmarExclusaoDefinitiva(CARD, 'Sondagem', 'vendas', admin);
  const after = plain(rt.context.carregarConteudoEditado());
  assert.equal(after.novidades.length, newsAfterSave + 1);
  assert.match(after.novidades[after.novidades.length - 1].corpo, /removido/);
});

test('deleting a card that was never edited creates a tombstone row (original content is not touched)', () => {
  const rt = fresh();
  const admin = login(rt, 'admin.demo');
  rt.context.excluirConteudoEditado('checklist|0|grupo|2', 'grupo', admin, 'Registro', 'vendas');
  const row = plain(rt.context.carregarConteudoEditado()).overrides.find((o) => o.id === 'checklist|0|grupo|2');
  assert.equal(row.excluido, true);
  assert.equal(row.json, '');
});

test('display order is stored per destination and the news list is not polluted by reordering', () => {
  const rt = fresh({ seedNews: false });
  const admin = login(rt, 'admin.demo');
  rt.context.salvarOrdemCartoes('script|0|bloco', JSON.stringify(['script|0|bloco|2', 'script|0|bloco|0']), admin);
  rt.context.salvarOrdemCartoes('script|0|bloco', JSON.stringify(['script|0|bloco|0']), admin);
  const data = plain(rt.context.carregarConteudoEditado());
  const orders = data.overrides.filter((o) => o.id === 'ordem|script|0|bloco');
  assert.equal(orders.length, 1);
  assert.deepEqual(JSON.parse(orders[0].json), ['script|0|bloco|0']);
  assert.equal(data.novidades.length, 0);
});

test('duplicated cards are appended and announced', () => {
  const rt = fresh({ seedNews: false });
  const admin = login(rt, 'admin.demo');
  assert.equal(rt.context.adicionarConteudoNovo('novo-abc', 'bloco', payload, admin, 'Cópia', 'vendas').ok, true);
  const data = plain(rt.context.carregarConteudoEditado());
  assert.ok(data.overrides.some((o) => o.id === 'novo-abc'));
  assert.match(data.novidades[0].titulo, /Novo card/);
});

/* ------------------------------------------------------------------ demo content integrity */
function loadContent() {
  const html = fs.readFileSync(path.join(__dirname, '..', 'apps-script', 'Index.html'), 'utf8');
  const start = html.indexOf('var SETORES = [');
  const endMarker = html.indexOf('   4. Navegação / render');
  assert.ok(start > 0 && endMarker > start, 'content block markers');
  const code = html.slice(start, html.lastIndexOf('/* ====', endMarker));
  const sandbox = {};
  vm.runInNewContext(code + '\n;this.out = { SETORES, CHECKLISTS, SCRIPTS, REFERENCIAS, FILAS_TRANSFERENCIA, CODIGOS_BANCOS, NOTICIAS, SETORES_COM_FILAS, REFERENCIAS_GERAIS, CLASSIFICACOES_PADRAO };', sandbox);
  return plain(sandbox.out);
}

test('demo content: six sectors with unique ids and every document points at a real sector', () => {
  const c = loadContent();
  assert.equal(c.SETORES.length, 6);
  const ids = c.SETORES.map((s) => s.id);
  assert.equal(new Set(ids).size, 6);
  c.CHECKLISTS.forEach((x) => assert.ok(ids.includes(x.setor), 'checklist sector ' + x.setor));
  c.SCRIPTS.forEach((x) => assert.ok(ids.includes(x.setor), 'script sector ' + x.setor));
  c.REFERENCIAS.forEach((x) => (Array.isArray(x.setor) ? x.setor : [x.setor]).forEach((s) => assert.ok(ids.includes(s), 'reference sector ' + s)));
  c.SETORES_COM_FILAS.forEach((s) => assert.ok(ids.includes(s)));
  c.NOTICIAS.forEach((n) => assert.ok(!n.setorRelacionado || ids.includes(n.setorRelacionado)));
  ids.forEach((id) => assert.ok(c.CHECKLISTS.some((x) => x.setor === id), 'sector without a checklist: ' + id));
  assert.equal(c.SCRIPTS.filter((s, i, a) => a.findIndex((o) => o.setor === s.setor) === i).length, c.SCRIPTS.length,
    'the UI shows one script per sector');
});

test('demo content: checklist criteria are well formed and weights are positive numbers', () => {
  const c = loadContent();
  c.CHECKLISTS.forEach((chk) => {
    assert.ok(chk.variante && chk.grupos.length > 0 && Array.isArray(chk.graves));
    chk.grupos.forEach((g) => {
      assert.ok(g.categoria && g.itens.length > 0);
      g.itens.forEach((it) => {
        assert.equal(typeof it.c, 'string');
        assert.ok(it.c.length > 5);
        assert.ok(Number.isFinite(it.p) && it.p > 0, 'weight of "' + it.c.slice(0, 30) + '"');
      });
    });
    const total = chk.grupos.reduce((s, g) => s + g.itens.reduce((a, it) => a + it.p, 0), 0);
    assert.ok(total <= 100, chk.variante + ' sums to ' + total + ', should not exceed 100');
  });
  assert.equal(c.CLASSIFICACOES_PADRAO.length, 3);
});

test('demo content: blocks and queue tables have the shape the renderer expects', () => {
  const c = loadContent();
  [...c.SCRIPTS, ...c.REFERENCIAS].forEach((doc) => {
    assert.ok(doc.titulo && doc.blocos.length > 0);
    doc.blocos.forEach((b) => {
      assert.ok(b.titulo && b.itens.length > 0);
      b.itens.forEach((it) => { assert.ok(it.rotulo && it.texto); });
    });
  });
  c.FILAS_TRANSFERENCIA.forEach((g) => { assert.match(g.cor, /^cor-[a-c]$/); g.itens.forEach((i) => assert.ok(i.setor && i.voz && i.chat)); });
  assert.equal(new Set(c.NOTICIAS.map((n) => n.id)).size, c.NOTICIAS.length);
  assert.ok(c.CODIGOS_BANCOS.every((b) => /^9\d\d$/.test(b.codigo)), 'bank codes are all fictional (9xx)');
});
