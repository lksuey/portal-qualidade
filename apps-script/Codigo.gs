var PLANILHA_ID_PROP = 'PORTAL_QUALIDADE_PLANILHA_ID';
var CHAVE_SESSAO_PROP = 'PORTAL_QUALIDADE_CHAVE_SESSAO';
var ABA_USUARIOS = 'Usuarios';
var ABA_TENTATIVAS = 'TentativasLogin';
var ABA_LOG = 'LogAcessos';
var ABA_CONTEUDO = 'ConteudoEditado';
var ABA_NOVIDADES = 'Novidades';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Portal da Qualidade')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function instalar() {
  var props = PropertiesService.getScriptProperties();
  var idExistente = props.getProperty(PLANILHA_ID_PROP);
  if (idExistente) {
    try {
      var existente = SpreadsheetApp.openById(idExistente);
      Logger.log('Já estava instalado. Planilha: ' + existente.getUrl());
      return;
    } catch (e) {
    }
  }

  var ss = SpreadsheetApp.create('Portal da Qualidade — Base de Login');
  props.setProperty(PLANILHA_ID_PROP, ss.getId());

  var usuarios = ss.getSheets()[0];
  usuarios.setName(ABA_USUARIOS);
  usuarios.appendRow(['Login', 'SenhaHash', 'Salt', 'Patente', 'Ativo', 'CriadoEm']);
  usuarios.setFrozenRows(1);
  usuarios.hideColumns(2, 2);

  var seed = [
    { login: 'admin.demo', patente: 'Admin' },
    { login: 'auditor.demo', patente: 'Auditor' },
    { login: 'operator.demo', patente: 'Operador' }
  ];
  seed.forEach(function (u) {
    var senha = props.getProperty('PORTAL_QUALIDADE_SEED_' + u.login.toUpperCase());
    if (!senha) { Logger.log('Sem senha inicial para ' + u.login + ': conta pulada.'); return; }
    var salt = Utilities.getUuid();
    usuarios.appendRow([u.login, hashSenha_(senha, salt), salt, u.patente, true, new Date()]);
  });

  var tentativas = ss.insertSheet(ABA_TENTATIVAS);
  tentativas.appendRow(['Chave', 'Tentativas', 'UltimaTentativa', 'BloqueadoAte']);
  tentativas.setFrozenRows(1);

  var log = ss.insertSheet(ABA_LOG);
  log.appendRow(['DataHora', 'Login', 'Patente', 'Sucesso', 'TokenDispositivo']);
  log.setFrozenRows(1);

  criarAbasConteudo_(ss);

  Logger.log('Instalado com sucesso: ' + ss.getUrl());
}

function criarAbasConteudo_(ss) {
  if (!ss.getSheetByName(ABA_CONTEUDO)) {
    var conteudo = ss.insertSheet(ABA_CONTEUDO);
    conteudo.appendRow(['ID', 'Tipo', 'JSON', 'Excluido', 'AtualizadoPor', 'AtualizadoEm']);
    conteudo.setFrozenRows(1);
  }
  if (!ss.getSheetByName(ABA_NOVIDADES)) {
    var novidades = ss.insertSheet(ABA_NOVIDADES);
    novidades.appendRow(['ID', 'Titulo', 'Corpo', 'Data', 'SetorRelacionado', 'CriadoPor']);
    novidades.setFrozenRows(1);
  }
}

function instalarEdicaoConteudo_10_09() {
  criarAbasConteudo_(planilha_());
  Logger.log('Abas de edição de conteúdo prontas.');
}

function planilha_() {
  var id = PropertiesService.getScriptProperties().getProperty(PLANILHA_ID_PROP);
  if (!id) throw new Error('Banco de dados ainda não instalado — rode a função instalar() uma vez pelo editor do Apps Script.');
  return SpreadsheetApp.openById(id);
}

function redefinirSenha_(login, novaSenha) {
  var aba = planilha_().getSheetByName(ABA_USUARIOS);
  var dados = aba.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).toLowerCase() === login.toLowerCase()) {
      var salt = Utilities.getUuid();
      aba.getRange(i + 1, 2, 1, 2).setValues([[hashSenha_(novaSenha, salt), salt]]);
      try { CacheService.getScriptCache().remove('usuarios_v1'); } catch (e) {}
      return true;
    }
  }
  return false;
}

function renomearLogin_(loginAntigo, loginNovo) {
  var aba = planilha_().getSheetByName(ABA_USUARIOS);
  var dados = aba.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).toLowerCase() === loginAntigo.toLowerCase()) {
      aba.getRange(i + 1, 1).setValue(loginNovo);
      try { CacheService.getScriptCache().remove('usuarios_v1'); } catch (e) {}
      return true;
    }
  }
  return false;
}

function hashSenha_(senha, salt) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + senha);
  return bytes.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

var LIMITES_BLOQUEIO = [
  { minimo: 20, minutos: 48 * 60 },
  { minimo: 15, minutos: 6 * 60 },
  { minimo: 10, minutos: 60 },
  { minimo: 5, minutos: 2 }
];
function duracaoBloqueioMinutos_(tentativas) {
  for (var i = 0; i < LIMITES_BLOQUEIO.length; i++) {
    if (tentativas >= LIMITES_BLOQUEIO[i].minimo) return LIMITES_BLOQUEIO[i].minutos;
  }
  return 0;
}

function lerTentativas_() {
  var aba = planilha_().getSheetByName(ABA_TENTATIVAS);
  var dados = aba.getDataRange().getValues();
  var mapa = {};
  for (var i = 1; i < dados.length; i++) {
    mapa[dados[i][0]] = { linha: i + 1, tentativas: Number(dados[i][1] || 0), bloqueadoAte: dados[i][3] || null };
  }
  return { aba: aba, mapa: mapa };
}

function bloqueioDoRegistro_(reg) {
  if (!reg || !reg.bloqueadoAte) return null;
  var ate = new Date(reg.bloqueadoAte);
  return ate.getTime() > Date.now() ? ate : null;
}

function registrarFalhaEm_(estado, chave) {
  var reg = estado.mapa[chave];
  var tentativas = (reg ? reg.tentativas : 0) + 1;
  var agora = new Date();
  var minutos = duracaoBloqueioMinutos_(tentativas);
  var ate = minutos > 0 ? new Date(Date.now() + minutos * 60000) : '';
  if (reg) {
    estado.aba.getRange(reg.linha, 2, 1, 3).setValues([[tentativas, agora, ate]]);
  } else {
    estado.aba.appendRow([chave, tentativas, agora, ate]);
  }
  return ate ? ate : null;
}

function limparTentativasEm_(estado, chave) {
  var reg = estado.mapa[chave];
  if (!reg) return;
  estado.aba.getRange(reg.linha, 2, 1, 3).setValues([[0, new Date(), '']]);
}

function maisTarde_(a, b) {
  if (a && b) return a.getTime() > b.getTime() ? a : b;
  return a || b || null;
}

function mensagemBloqueio_(ate) {
  var min = Math.max(1, Math.ceil((ate.getTime() - Date.now()) / 60000));
  if (min < 60) return 'Muitas tentativas. Tente novamente em ' + min + ' minuto' + (min > 1 ? 's' : '') + '.';
  var horas = Math.ceil(min / 60);
  return 'Muitas tentativas. Tente novamente em ' + horas + ' hora' + (horas > 1 ? 's' : '') + '.';
}

var REGEX_LOGIN_VALIDO = /^[A-Za-z0-9._-]{3,40}$/;
var REGEX_EMOJI = /[\u{1F300}-\u{1FAFF}\u{2600}-\u{27BF}\u{1F1E6}-\u{1F1FF}\u{2190}-\u{21FF}\u{2B00}-\u{2BFF}]/u;

function entradaValida_(login, senha) {
  if (typeof login !== 'string' || typeof senha !== 'string') return false;
  if (!REGEX_LOGIN_VALIDO.test(login)) return false;
  if (senha.length < 3 || senha.length > 100) return false;
  if (REGEX_EMOJI.test(login) || REGEX_EMOJI.test(senha)) return false;
  if (/[<>`]/.test(login) || /[<>`]/.test(senha)) return false;
  return true;
}

var MENSAGEM_ERRO_PADRAO = 'Login ou Senha incorreto. Tente novamente.';

function lerUsuarios_() {
  var cache = CacheService.getScriptCache();
  var cacheado = cache.get('usuarios_v1');
  if (cacheado) return JSON.parse(cacheado);
  var dados = planilha_().getSheetByName(ABA_USUARIOS).getDataRange().getValues();
  try { cache.put('usuarios_v1', JSON.stringify(dados), 300); } catch (e) {  }
  return dados;
}

function verificarLogin(login, senha, deviceToken) {
  login = String(login == null ? '' : login).trim();
  senha = String(senha == null ? '' : senha);
  deviceToken = String(deviceToken == null ? '' : deviceToken).trim();
  if (!/^[A-Za-z0-9-]{8,64}$/.test(deviceToken)) deviceToken = 'sem-token';

  if (!entradaValida_(login, senha)) {
    return { ok: false, mensagem: MENSAGEM_ERRO_PADRAO };
  }

  var chaveLogin = 'login:' + login.toLowerCase();
  var chaveDispositivo = 'device:' + deviceToken;

  var estado = lerTentativas_();
  var bloqueioAtivo = maisTarde_(bloqueioDoRegistro_(estado.mapa[chaveLogin]), bloqueioDoRegistro_(estado.mapa[chaveDispositivo]));
  if (bloqueioAtivo) {
    return { ok: false, mensagem: mensagemBloqueio_(bloqueioAtivo), bloqueado: true, bloqueadoAte: bloqueioAtivo.getTime() };
  }

  var dados = lerUsuarios_();
  var usuario = null;
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).toLowerCase() === login.toLowerCase() && dados[i][4] === true) {
      usuario = { login: dados[i][0], senhaHash: dados[i][1], salt: dados[i][2], patente: dados[i][3] };
      break;
    }
  }

  var sucesso = !!usuario && usuario.senhaHash === hashSenha_(senha, usuario.salt);

  if (!sucesso) {
    var ateFalha = maisTarde_(registrarFalhaEm_(estado, chaveLogin), registrarFalhaEm_(estado, chaveDispositivo));
    if (ateFalha) return { ok: false, mensagem: mensagemBloqueio_(ateFalha), bloqueado: true, bloqueadoAte: ateFalha.getTime() };
    return { ok: false, mensagem: MENSAGEM_ERRO_PADRAO };
  }

  limparTentativasEm_(estado, chaveLogin);
  limparTentativasEm_(estado, chaveDispositivo);
  return { ok: true, patente: usuario.patente, token: gerarTokenSessao_(usuario.login, usuario.patente) };
}

function registrarAcessoAsync(login, patente, sucesso, deviceToken) {
  try {
    planilha_().getSheetByName(ABA_LOG).appendRow([new Date(), login, patente, sucesso, deviceToken]);
  } catch (e) {
  }
}

function chaveAssinatura_() {
  var props = PropertiesService.getScriptProperties();
  var k = props.getProperty(CHAVE_SESSAO_PROP);
  if (!k) { k = Utilities.getUuid(); props.setProperty(CHAVE_SESSAO_PROP, k); }
  return k;
}
function assinar_(payload) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, payload + chaveAssinatura_());
  return bytes.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}
function gerarTokenSessao_(login, patente) {
  var expira = Date.now() + 8 * 60 * 60 * 1000;
  var payload = login + '|' + patente + '|' + expira;
  return Utilities.base64EncodeWebSafe(payload) + '.' + assinar_(payload);
}

function validarTokenSessao(token) {
  try {
    var partes = String(token || '').split('.');
    if (partes.length !== 2) return { ok: false };
    var payload = Utilities.newBlob(Utilities.base64DecodeWebSafe(partes[0])).getDataAsString();
    if (assinar_(payload) !== partes[1]) return { ok: false };
    var campos = payload.split('|');
    var expira = Number(campos[2]);
    if (!expira || Date.now() > expira) return { ok: false, expirado: true };
    return { ok: true, login: campos[0], patente: campos[1] };
  } catch (e) {
    return { ok: false };
  }
}


function linhaConteudoPorId_(aba, id) {
  var dados = aba.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    if (dados[i][0] === id) return i + 1;
  }
  return -1;
}

function carregarConteudoEditado() {
  var ss = planilha_();
  var overrides = [];
  var abaC = ss.getSheetByName(ABA_CONTEUDO);
  if (abaC) {
    var dadosC = abaC.getDataRange().getValues();
    for (var i = 1; i < dadosC.length; i++) {
      if (!dadosC[i][0]) continue;
      overrides.push({ id: dadosC[i][0], tipo: dadosC[i][1], json: dadosC[i][2], excluido: dadosC[i][3] === true });
    }
  }
  var novidades = [];
  var abaN = ss.getSheetByName(ABA_NOVIDADES);
  if (abaN) {
    var dadosN = abaN.getDataRange().getValues();
    for (var j = 1; j < dadosN.length; j++) {
      if (!dadosN[j][0]) continue;
      novidades.push({
        id: dadosN[j][0], titulo: dadosN[j][1], corpo: dadosN[j][2],
        data: Utilities.formatDate(new Date(dadosN[j][3]), Session.getScriptTimeZone() || 'GMT', 'yyyy-MM-dd'),
        setorRelacionado: dadosN[j][4] || ''
      });
    }
  }
  return { overrides: overrides, novidades: novidades };
}

function registrarNovidade_(titulo, corpo, setorRelacionado, autor) {
  try {
    var aba = planilha_().getSheetByName(ABA_NOVIDADES);
    aba.appendRow([Utilities.getUuid(), titulo, corpo, new Date(), setorRelacionado || '', autor || '']);
  } catch (e) {
  }
}

function exigirAdmin_(tokenSessao) {
  var v = validarTokenSessao(tokenSessao);
  if (!v.ok || v.patente !== 'Admin') return null;
  return v;
}

function salvarConteudoEditado(id, tipo, dadosJson, tokenSessao, tituloResumo, setorResumo) {
  var v = exigirAdmin_(tokenSessao);
  if (!v) return { ok: false, mensagem: 'Sem permissão para editar conteúdo.' };

  var aba = planilha_().getSheetByName(ABA_CONTEUDO);
  var linha = linhaConteudoPorId_(aba, id);
  var agora = new Date();
  if (linha === -1) {
    aba.appendRow([id, tipo, dadosJson, false, v.login, agora]);
  } else {
    aba.getRange(linha, 2, 1, 5).setValues([[tipo, dadosJson, false, v.login, agora]]);
  }
  registrarNovidade_(
    'Conteúdo atualizado',
    'O card "' + (tituloResumo || id) + '" foi atualizado por ' + v.login + '.',
    setorResumo || '', v.login
  );
  return { ok: true };
}

function adicionarConteudoNovo(id, tipo, dadosJson, tokenSessao, tituloResumo, setorResumo) {
  var v = exigirAdmin_(tokenSessao);
  if (!v) return { ok: false, mensagem: 'Sem permissão para adicionar conteúdo.' };

  var aba = planilha_().getSheetByName(ABA_CONTEUDO);
  var linha = linhaConteudoPorId_(aba, id);
  var agora = new Date();
  if (linha === -1) {
    aba.appendRow([id, tipo, dadosJson, false, v.login, agora]);
  } else {
    aba.getRange(linha, 2, 1, 5).setValues([[tipo, dadosJson, false, v.login, agora]]);
  }
  registrarNovidade_(
    'Novo card adicionado',
    'Um novo card ("' + (tituloResumo || id) + '") foi adicionado por ' + v.login + '.',
    setorResumo || '', v.login
  );
  return { ok: true };
}

function salvarOrdemCartoes(destino, ordemJson, tokenSessao) {
  var v = exigirAdmin_(tokenSessao);
  if (!v) return { ok: false, mensagem: 'Sem permissão para reordenar conteúdo.' };
  var id = 'ordem|' + destino;
  var aba = planilha_().getSheetByName(ABA_CONTEUDO);
  var linha = linhaConteudoPorId_(aba, id);
  var agora = new Date();
  if (linha === -1) {
    aba.appendRow([id, 'ordem', ordemJson, false, v.login, agora]);
  } else {
    aba.getRange(linha, 2, 1, 5).setValues([['ordem', ordemJson, false, v.login, agora]]);
  }
  return { ok: true };
}

function excluirConteudoEditado(id, tipo, tokenSessao, tituloResumo, setorResumo) {
  var v = exigirAdmin_(tokenSessao);
  if (!v) return { ok: false, mensagem: 'Sem permissão para excluir conteúdo.' };

  var aba = planilha_().getSheetByName(ABA_CONTEUDO);
  var linha = linhaConteudoPorId_(aba, id);
  var agora = new Date();
  if (linha === -1) {
    aba.appendRow([id, tipo, '', true, v.login, agora]);
  } else {
    aba.getRange(linha, 4, 1, 3).setValues([[true, v.login, agora]]);
  }
  return { ok: true };
}

function confirmarExclusaoDefinitiva(id, tituloResumo, setorResumo, tokenSessao) {
  var v = exigirAdmin_(tokenSessao);
  if (!v) return { ok: false };
  var aba = planilha_().getSheetByName(ABA_CONTEUDO);
  var linha = linhaConteudoPorId_(aba, id);
  if (linha === -1) return { ok: false };
  var excluido = aba.getRange(linha, 4).getValue();
  if (excluido !== true) return { ok: true };
  registrarNovidade_(
    'Conteúdo removido',
    'O card "' + (tituloResumo || id) + '" foi removido por ' + v.login + '.',
    setorResumo || '', v.login
  );
  return { ok: true };
}

function restaurarConteudoEditado(id, tokenSessao) {
  var v = exigirAdmin_(tokenSessao);
  if (!v) return { ok: false };
  var aba = planilha_().getSheetByName(ABA_CONTEUDO);
  var linha = linhaConteudoPorId_(aba, id);
  if (linha === -1) return { ok: false };
  aba.getRange(linha, 4).setValue(false);
  return { ok: true };
}
