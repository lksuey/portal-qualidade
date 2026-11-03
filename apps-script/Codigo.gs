/**
 * PORTAL DA QUALIDADE — central de processos e novidades da equipe de
 * atendimento, agora com login (09/09).
 *
 * ATENÇÃO — leia antes de rodar:
 * O Apps Script (e o navegador) NÃO tem acesso ao IP nem ao endereço MAC de
 * quem acessa a página — isso não existe nem no HtmlService nem em nenhuma
 * API do Apps Script (não é uma limitação deste código, é do próprio Google
 * Apps Script). O pedido original era travar por "IP e MAC address"; o mais
 * próximo tecnicamente possível — e o que foi implementado aqui — é travar
 * por dois eixos que dá pra medir de verdade:
 *   1) o LOGIN que está sendo tentado (trava quem tenta adivinhar a senha
 *      de um usuário específico, não importa de onde vem a tentativa);
 *   2) um TOKEN DE DISPOSITIVO gerado no navegador na primeira visita e
 *      guardado no localStorage — funciona como um "RG do navegador",
 *      persiste entre sessões daquele aparelho/navegador, mas não é o MAC
 *      de verdade (se a pessoa limpar os dados do site ou trocar de
 *      navegador, ganha um token novo).
 * Toda tentativa errada soma nos dois contadores; o mais restritivo dos
 * dois é quem trava o login. Ver verificarLogin() mais abaixo.
 *
 * Primeira execução: em Project Settings > Script Properties, crie as
 * propriedades PORTAL_QUALIDADE_SEED_ADMIN.DEMO, PORTAL_QUALIDADE_SEED_AUDITOR.DEMO
 * e PORTAL_QUALIDADE_SEED_OPERATOR.DEMO com as senhas iniciais que você quiser
 * (nenhuma senha mora no código). Depois abra este arquivo no editor,
 * selecione a função "instalar" e clique em Executar: isso cria a planilha
 * "Portal da Qualidade — Base de Login" com as contas iniciais
 * (Admin/Auditor/Operador) e as abas de controle de tentativas e log de
 * acessos. Rodar de novo depois de já instalado não duplica nada.
 */

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

/* =========================================================================
   Instalação — roda uma vez, manualmente, pelo editor do Apps Script.
   ========================================================================= */
function instalar() {
  var props = PropertiesService.getScriptProperties();
  var idExistente = props.getProperty(PLANILHA_ID_PROP);
  if (idExistente) {
    try {
      var existente = SpreadsheetApp.openById(idExistente);
      Logger.log('Já estava instalado. Planilha: ' + existente.getUrl());
      return;
    } catch (e) {
      // O ID salvo não abre mais (planilha apagada) — reinstala do zero.
    }
  }

  var ss = SpreadsheetApp.create('Portal da Qualidade — Base de Login');
  props.setProperty(PLANILHA_ID_PROP, ss.getId());

  var usuarios = ss.getSheets()[0];
  usuarios.setName(ABA_USUARIOS);
  usuarios.appendRow(['Login', 'SenhaHash', 'Salt', 'Patente', 'Ativo', 'CriadoEm']);
  usuarios.setFrozenRows(1);
  usuarios.hideColumns(2, 2); // esconde SenhaHash e Salt na visualização padrão

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

/* Rode esta função UMA VEZ pelo editor do Apps Script se você já tinha
   instalado o portal ANTES de existir a edição de conteúdo (10/09) — ela
   só adiciona as abas ConteudoEditado e Novidades na planilha que já
   existe, sem mexer no resto. Rodar de novo não duplica nada. */
function instalarEdicaoConteudo_10_09() {
  criarAbasConteudo_(planilha_());
  Logger.log('Abas de edição de conteúdo prontas.');
}

function planilha_() {
  var id = PropertiesService.getScriptProperties().getProperty(PLANILHA_ID_PROP);
  if (!id) throw new Error('Banco de dados ainda não instalado — rode a função instalar() uma vez pelo editor do Apps Script.');
  return SpreadsheetApp.openById(id);
}

/* =========================================================================
   Troca de senha: instalar() só semeia as contas na primeira vez, então
   pra atualizar a senha de um usuário que já existe na planilha é preciso
   isso aqui (nunca editar a coluna SenhaHash na mão — ela é derivada da
   senha + salt, não dá pra simplesmente escrever a senha nova nela).
   ========================================================================= */
function redefinirSenha_(login, novaSenha) {
  var aba = planilha_().getSheetByName(ABA_USUARIOS);
  var dados = aba.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    if (String(dados[i][0]).toLowerCase() === login.toLowerCase()) {
      var salt = Utilities.getUuid();
      aba.getRange(i + 1, 2, 1, 2).setValues([[hashSenha_(novaSenha, salt), salt]]);
      try { CacheService.getScriptCache().remove('usuarios_v1'); } catch (e) {} // senão a senha antiga ainda "funciona" até o cache expirar
      return true;
    }
  }
  return false;
}

/* Renomeia o LOGIN de um usuário já existente (mantém senha, salt e
   patente — só troca a coluna Login). */
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

/* =========================================================================
   Hash de senha — SHA-256 com salt aleatório por usuário (Apps Script não
   tem bcrypt/scrypt nativos; SHA-256 + salt é o que dá pra fazer sem
   biblioteca externa, e já é muito melhor que guardar a senha em texto
   puro na planilha).
   ========================================================================= */
function hashSenha_(senha, salt) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + ':' + senha);
  return bytes.map(function (b) { return ('0' + (b & 0xFF).toString(16)).slice(-2); }).join('');
}

/* =========================================================================
   Regras de tentativas e bloqueio progressivo (pedido do usuário, 09/09):
   5 tentativas seguidas -> 2 min | 10 -> 1h | 15 -> 6h | 20 -> 48h.
   O contador é cumulativo desde o último acerto (só zera em login com
   sucesso) — não zera sozinho quando o bloqueio expira.
   ========================================================================= */
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

/* ---------- Acesso otimizado à aba de tentativas ----------
   A versão original fazia até 4 leituras completas da aba TentativasLogin
   por tentativa de login (uma pra cada chave, em cada função) — cada
   chamada ao SpreadsheetApp é uma viagem de rede, e isso é a causa mais
   provável da lentidão percebida (pedido do usuário, 09/09: "a
   responsividade da pagina... está demorando bastante"). Agora é 1 leitura
   só (lerTentativas_), guardada em memória, e as escritas só acontecem
   quando realmente há algo pra gravar — no caminho mais comum (senha
   certa de primeira, sem bloqueio ativo), não escreve nada nessa aba. */
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
  if (!reg) return; // nunca teve tentativa registrada — nada pra limpar, pula a escrita
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

/* =========================================================================
   Validação de entrada — anti-emoji e anti-"injeção" (não existe SQL nessa
   pilha — é Google Sheets via SpreadsheetApp, não banco relacional — então
   "SQL Injection" de verdade não se aplica; o que dá pra e deve ser feito é
   nunca usar login/senha para montar fórmula/consulta dinâmica, e nunca
   fazer nada além de comparação literal de string, o que este código já
   faz. A validação abaixo é uma camada extra de sanidade em cima disso.)
   ========================================================================= */
var REGEX_LOGIN_VALIDO = /^[A-Za-z0-9._-]{3,40}$/;
// Emojis e símbolos pictográficos comuns (não cobre 100% do Unicode, mas
// cobre a esmagadora maioria dos emojis usados em teclados normais).
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

/* =========================================================================
   Login — chamada pelo cliente via google.script.run.
   Resposta de erro é SEMPRE a mesma frase pra credencial errada, por
   pedido do usuário — só a mensagem de bloqueio por excesso de tentativas
   é diferente (não revela se a credencial em si estava certa ou errada).
   ========================================================================= */
/* Usuarios muda raramente (só quando alguém mexe na planilha na mão), então
   cacheia por 5 minutos com CacheService — que é memória, não tem o custo
   de rede do SpreadsheetApp. Isso soma com a otimização da aba de
   tentativas pra deixar o caminho mais comum (senha certa, sem bloqueio)
   em praticamente 1 chamada de rede em vez de 6+. */
function lerUsuarios_() {
  var cache = CacheService.getScriptCache();
  var cacheado = cache.get('usuarios_v1');
  if (cacheado) return JSON.parse(cacheado);
  var dados = planilha_().getSheetByName(ABA_USUARIOS).getDataRange().getValues();
  try { cache.put('usuarios_v1', JSON.stringify(dados), 300); } catch (e) { /* linha muito grande pro cache — segue sem cachear */ }
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

  // 1 leitura só da aba de tentativas, cobrindo as duas chaves.
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

  // Só escreve se havia contador pra zerar — no caminho mais comum (acertou
  // de primeira) não existe registro nenhum ainda, então isso não gera
  // escrita nenhuma na planilha.
  limparTentativasEm_(estado, chaveLogin);
  limparTentativasEm_(estado, chaveDispositivo);
  return { ok: true, patente: usuario.patente, token: gerarTokenSessao_(usuario.login, usuario.patente) };
}

/* Log de acesso — chamado pelo cliente separadamente, DEPOIS de já ter
   processado a resposta do login (sucesso ou erro), sem que o cliente
   espere essa chamada terminar. Pedido do usuário (09/09): "preciso que
   isso seja feito pós login ou que haja uma otimização melhor" — aqui
   ficaram as duas coisas: a escrita foi tirada do caminho crítico E o
   acesso à planilha foi otimizado (ver lerTentativas_/lerUsuarios_). */
function registrarAcessoAsync(login, patente, sucesso, deviceToken) {
  try {
    planilha_().getSheetByName(ABA_LOG).appendRow([new Date(), login, patente, sucesso, deviceToken]);
  } catch (e) {
    // Nunca deixa o log de acesso travar nada — é best-effort.
  }
}

/* =========================================================================
   Token de sessão — assinado com uma chave guardada nas Script Properties,
   pra o cliente não conseguir forjar a própria patente no localStorage.
   Não é um JWT completo, mas resolve o mesmo problema pra esse tamanho de
   projeto: validar, do lado do servidor, que o token realmente veio daqui.
   Expira em 8 horas (duração de um turno de trabalho).
   ========================================================================= */
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

/* Chamada pelo cliente ao carregar a página, pra saber se o token salvo no
   localStorage ainda é válido (evita pedir login de novo a cada acesso
   dentro da janela de 8h). */
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

/* =========================================================================
   Edição de conteúdo pelo Admin (10/09)

   O conteúdo "de fábrica" (checklists, scripts, referências) continua
   fixo dentro do Index.html — não foi todo migrado pra planilha, o que
   deixaria a leitura inicial da página muito mais lenta. Em vez disso,
   cada card/bloco tem um ID estável (calculado no cliente a partir da
   posição dele no array de origem) e as EDIÇÕES ficam guardadas à parte,
   na aba ConteudoEditado, como "diffs" — no carregamento da página o
   cliente busca esses diffs e aplica por cima do conteúdo padrão.

   Exclusão nunca remove nada de verdade do conteúdo original — só marca
   Excluido=true pro ID daquele card. Isso é o que permite "desfazer": a
   qualquer momento, restaurarConteudo() só volta o Excluido pra false.
   Também é o que mantém os IDs dos outros cards estáveis (apagar um bloco
   não desloca a posição dos que vêm depois).
   ========================================================================= */

function linhaConteudoPorId_(aba, id) {
  var dados = aba.getDataRange().getValues();
  for (var i = 1; i < dados.length; i++) {
    if (dados[i][0] === id) return i + 1;
  }
  return -1;
}

/* Chamada pelo cliente ao entrar no portal — traz de uma vez as edições
   já feitas (pra aplicar por cima do conteúdo padrão) e as novidades
   registradas por edições (pra somar às novidades de exemplo). */
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
    // Nunca deixa a notificação travar a edição em si.
  }
}

/* Toda escrita de conteúdo exige um token de sessão válido com patente
   Admin — nunca confia na patente que o cliente diz ter, só na que está
   assinada dentro do token (ver gerarTokenSessao_). */
function exigirAdmin_(tokenSessao) {
  var v = validarTokenSessao(tokenSessao);
  if (!v.ok || v.patente !== 'Admin') return null;
  return v;
}

/* Salva (cria ou substitui) o conteúdo editado de um card. `dadosJson` já
   vem serializado do cliente (string JSON) — o servidor não precisa
   entender o formato, só guardar e devolver depois. */
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

/* Cria um card novo por duplicação (pedido do usuário, 14/09: "precisamos
   criar uma forma de adicionar cards, o mais viavel que eu achei é o
   duplicação"). O ID já vem pronto do cliente ("novo-...") e nunca existia
   antes, então isso sempre é um appendRow — a única diferença real pra
   salvarConteudoEditado() é o texto da novidade publicada. */
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

/* Salva a ordem de exibição escolhida pelo Admin pra um "destino" (ex.:
   "checklist|3|grupo" = os grupos de critérios daquele checklist). Fica
   guardada na mesma aba, com um ID próprio ("ordem|<destino>") — o valor é
   só a lista de IDs de card na ordem desejada. Não gera novidade (o
   arrasto acontece a cada solto do mouse; virar uma notificação a cada vez
   ia poluir o painel de novidades). Pedido do usuário (14/09): "preciso
   que seja adicionado agora uma forma de mover os cards, reorganizar com o
   proprio cursor". */
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

/* Exclusão é sempre "soft" — só marca Excluido=true, nunca apaga a linha.
   Isso é o que possibilita o desfazer de 15s no cliente. */
/* Não registra a novidade aqui na hora — o cliente ainda oferece 15s de
   "desfazer", e se a pessoa desfizer, a exclusão nunca "aconteceu" de
   verdade pra quem só vê a tela. A novidade só é criada por
   confirmarExclusaoDefinitiva(), chamada pelo cliente quando os 15s
   passam SEM desfazer. Bug relatado pelo usuário (14/09): "a notificação
   de update aparece mesmo após desfazer o deletamento". */
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

/* Chamada pelo cliente só depois que a janela de 15s de "desfazer" passa
   sem a pessoa desfazer — é nesse momento que a remoção vira definitiva de
   verdade e a novidade é publicada. Confere de novo se continua excluído
   (se foi desfeito, não notifica nada). */
function confirmarExclusaoDefinitiva(id, tituloResumo, setorResumo, tokenSessao) {
  var v = exigirAdmin_(tokenSessao);
  if (!v) return { ok: false };
  var aba = planilha_().getSheetByName(ABA_CONTEUDO);
  var linha = linhaConteudoPorId_(aba, id);
  if (linha === -1) return { ok: false };
  var excluido = aba.getRange(linha, 4).getValue();
  if (excluido !== true) return { ok: true }; // foi desfeito nesse meio-tempo — não notifica
  registrarNovidade_(
    'Conteúdo removido',
    'O card "' + (tituloResumo || id) + '" foi removido por ' + v.login + '.',
    setorResumo || '', v.login
  );
  return { ok: true };
}

/* Desfazer — só volta Excluido pra false, não gera novidade nova (a
   remoção nunca chegou a "acontecer" de fato pra quem só via a tela). */
function restaurarConteudoEditado(id, tokenSessao) {
  var v = exigirAdmin_(tokenSessao);
  if (!v) return { ok: false };
  var aba = planilha_().getSheetByName(ABA_CONTEUDO);
  var linha = linhaConteudoPorId_(aba, id);
  if (linha === -1) return { ok: false };
  aba.getRange(linha, 4).setValue(false);
  return { ok: true };
}
