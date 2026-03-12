// =======================
// CONFIGURAÇÕES GLOBAIS
// =======================

var BASE_CLARA_ID = "1_XW0IqbYjiCPpqtwdEi1xPxDlIP2MSkMrLGbeinLIeI"; // ID real da planilha que uso
var HIST_PEND_CLARA_RAW = "HIST_PEND_CLARA_RAW";

function normalizarLojaNumero_(valor) {
  var digits = String(valor || "").replace(/\D/g, "");
  if (!digits) return null;
  var n = Number(digits);
  return isFinite(n) ? n : null; // ignora zeros à esquerda
}

/**
 * BaseClara:
 * - Coluna R = 18 = "Grupos" (Time)
 * - Coluna V = 22 = "LojaNum"
 */
function construirMapaLojaParaTime_() {
  var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
  var sh = ss.getSheetByName("BaseClara");
  if (!sh) throw new Error("Aba BaseClara não encontrada.");

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return {};

  // Lê R:V (18..22) => 5 colunas: [Grupos, S, T, U, LojaNum]
  var values = sh.getRange(2, 18, lastRow - 1, 5).getValues();

  var map = {};      // lojaNum(number) -> time(string)
  var freq = {};     // para escolher o mais frequente por loja (caso exista mais de 1)

  values.forEach(function (r) {
    var time = String(r[0] || "").trim();   // R
    var lojaNum = normalizarLojaNumero_(r[4]); // V
    if (!lojaNum || !time) return;

    var key = String(lojaNum);
    if (!freq[key]) freq[key] = {};
    freq[key][time] = (freq[key][time] || 0) + 1;
  });

  Object.keys(freq).forEach(function (k) {
    var best = null;
    var bestN = -1;
    Object.keys(freq[k]).forEach(function (t) {
      if (freq[k][t] > bestN) { bestN = freq[k][t]; best = t; }
    });
    map[Number(k)] = best;
  });

  return map;
}

var PROP_LAST_SNAPSHOT_SIG = "VEKTOR_HISTPEND_LAST_SIG";

function LIMPAR_ANTI_SPAM() {
var props = PropertiesService.getScriptProperties();
  var cicloKey = getCicloKey06a05_();
  props.deleteProperty("VEKTOR_ALERTS_SENT_" + cicloKey);
  Logger.log("Limpou anti-spam do ciclo: " + cicloKey);
}

function isAdminEmail(email) {
  try {
    var e = String(email || "").trim().toLowerCase();
    if (!e) return false;

    // RBAC: mapa de e-mails e roles
    var map = vektorLoadEmailsRoleMap_(); // esperado: { byEmail: { "a@b": { role:"Administrador", ativo:true } } }
    var rec = map && map.byEmail ? map.byEmail[e] : null;

    if (!rec) return false;
    if (rec.ativo === false) return false;

    var role = String(rec.role || "").trim().toLowerCase();
    return role === "administrador";
  } catch (err) {
    return false;
  }
}

// =======================
// VEKTOR - CONTROLE DE ACESSO (WHITELIST via planilha VEKTOR_EMAILS)
// =======================

// ✅ A whitelist agora é a aba VEKTOR_EMAILS (EMAIL/ROLE/ATIVO)
// ATIVO=SIM => habilitado; ATIVO=NÃO => bloqueado
function isWhitelistedEmail_(email) {
  try {
    var e = String(email || "").trim().toLowerCase();
    if (!e) return false;

    var map = vektorLoadEmailsRoleMap_(); // { byEmail: { "a@b": { role, ativo } } }
    var rec = map && map.byEmail ? map.byEmail[e] : null;
    if (!rec) return false;

    // ATIVO precisa ser true
    return rec.ativo === true;
  } catch (err) {
    return false;
  }
}

// (recomendado) Use este "porteiro" no começo das funções expostas via google.script.run
function vektorAssertWhitelisted_() {
  var sess = (Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  if (!sess) throw new Error("Não foi possível identificar seu e-mail Google.");

  // ✅ fonte de verdade: VEKTOR_EMAILS (ATIVO)
  if (!isWhitelistedEmail_(sess)) {
    throw new Error("Acesso negado: usuário não habilitado no Vektor.");
  }
  return sess;
}

/**
 * Valida o login digitado no modal:
 * - deve bater com Session.getActiveUser().getEmail()
 * - e deve estar na whitelist
 */
function validarLoginVektor(emailInformado) {
  var sess = (Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  var inf  = (emailInformado || "").toString().trim().toLowerCase();

  if (!sess) {
    return { ok: false, error: "Não foi possível identificar seu e-mail Google (sessão vazia)." };
  }

  if (!inf) {
    return { ok: false, error: "Informe seu e-mail corporativo." };
  }

  if (inf !== sess) {
    return { ok: false, error: "O e-mail informado não confere com o seu login Google." };
  }

  if (!isWhitelistedEmail_(sess)) {
    return { ok: false, error: "Acesso negado: seu e-mail não está habilitado no Vektor." };
  }

  // RBAC: precisa estar ATIVO na VEKTOR_EMAILS
try {
  vektorGetUserRole_(); // valida ATIVO e retorna role
} catch (e) {
  return {
    ok: false,
    error: "Acesso não disponível. Solicite a liberação junto ao administrador do sistema."
  };
}

  var token = vektorCreateSessionToken_(sess);
    // ✅ registra usuário como "ativo hoje" (inclui o próprio acesso)
  try {
    vektorTrackActiveUserToday_(sess);
  } catch (eTrack) {
    // não pode quebrar login por falha de log
  }
  return { ok: true, email: sess, token: token, ttlSeconds: VEKTOR_SESSION_TTL_SECONDS };

}

// =======================
// VEKTOR - RBAC POR ROLE (VEKTOR_EMAILS + VEKTOR_ACESSOS)
// Mantém WHITELIST como porteiro 0
// =======================
var VEKTOR_EMAILS_SHEET = "VEKTOR_EMAILS";
var VEKTOR_ACESSOS_SHEET = "VEKTOR_ACESSOS";

// Usa a mesma planilha do Clara (BaseClara / Info_limites etc.)
var VEKTOR_ACL_SPREADSHEET_ID = "1_XW0IqbYjiCPpqtwdEi1xPxDlIP2MSkMrLGbeinLIeI";

var VEKTOR_ACL_CACHE_EMAILS = "VEKTOR_ACL_EMAILS_V1";
var VEKTOR_ACL_CACHE_ACESSOS = "VEKTOR_ACL_ACESSOS_V1";
var VEKTOR_ACL_CACHE_TTL = 120; // 2 min

function vektorNorm_(s) {
  return String(s || "").trim();
}
function vektorNormEmail_(s) {
  return vektorNorm_(s).toLowerCase();
}
function vektorIsAtivo_(v) {
  var x = String(v || "").trim().toUpperCase();
  return v === true || x === "TRUE" || x === "SIM" || x === "S" || x === "YES";
}

function vektorLoadEmailsRoleMap_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(VEKTOR_ACL_CACHE_EMAILS);
  if (cached) {
    try { return JSON.parse(cached); } catch (_) {}
  }

  var ss = SpreadsheetApp.openById(VEKTOR_ACL_SPREADSHEET_ID);
  var sh = ss.getSheetByName(VEKTOR_EMAILS_SHEET);
  if (!sh) throw new Error('Aba "' + VEKTOR_EMAILS_SHEET + '" não encontrada.');

  var values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return { byEmail: {} };

  var header = values[0].map(function (h) { return vektorNorm_(h); });
  var iEmail = header.indexOf("EMAIL");
  var iRole  = header.indexOf("ROLE");
  var iAtivo = header.indexOf("ATIVO");
  if (iEmail < 0 || iRole < 0 || iAtivo < 0) {
    throw new Error('Cabeçalho inválido em "' + VEKTOR_EMAILS_SHEET + '". Esperado: EMAIL, ROLE, ATIVO.');
  }

  var byEmail = {};
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var email = vektorNormEmail_(row[iEmail]);
    if (!email) continue;

    byEmail[email] = {
      role: vektorNorm_(row[iRole]) || "Acesso padrão",
      ativo: vektorIsAtivo_(row[iAtivo])
    };
  }

  var out = { byEmail: byEmail };
  cache.put(VEKTOR_ACL_CACHE_EMAILS, JSON.stringify(out), VEKTOR_ACL_CACHE_TTL);
  return out;
}

/**
 * Retorna lista única de "ROLE" da VEKTOR_EMAILS (para filtro Admin em Meus Alertas)
 */
function getRolesParaFiltroAlertasVektor() {
  try {
    if (typeof vektorAssertFunctionAllowed_ === "function") {
      vektorAssertFunctionAllowed_("getRolesParaFiltroAlertasVektor");
    }

    var map = vektorLoadEmailsRoleMap_();
    var set = {};

    if (map && map.byEmail) {
      Object.keys(map.byEmail).forEach(function (email) {
        var role = map.byEmail[email] && map.byEmail[email].role ? String(map.byEmail[email].role).trim() : "";
        if (!role) return;
        set[role] = true;
      });
    }

    return { ok: true, roles: Object.keys(set).sort() };
  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

function vektorLoadRoleAllowedFunctions_() {
  var cache = CacheService.getScriptCache();
  var cached = cache.get(VEKTOR_ACL_CACHE_ACESSOS);
  if (cached) {
    try { return JSON.parse(cached); } catch (_) {}
  }

  var ss = SpreadsheetApp.openById(VEKTOR_ACL_SPREADSHEET_ID);
  var sh = ss.getSheetByName(VEKTOR_ACESSOS_SHEET);
  if (!sh) throw new Error('Aba "' + VEKTOR_ACESSOS_SHEET + '" não encontrada.');

  var values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return { byRole: {} };

  var header = values[0].map(function (h) { return vektorNorm_(h); });
  var iRoles = header.indexOf("ROLES");
  var iFunc  = header.indexOf("FUNCTION_ALLOW");
  if (iRoles < 0 || iFunc < 0) {
    throw new Error('Cabeçalho inválido em "' + VEKTOR_ACESSOS_SHEET + '". Esperado: ROLES, FUNCTION_ALLOW, DESCRIPTION.');
  }

  var byRole = {}; // role -> { all:boolean, funcs:{name:true} }
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var role = vektorNorm_(row[iRoles]);
    var fn   = vektorNorm_(row[iFunc]);
    if (!role || !fn) continue;

    if (!byRole[role]) byRole[role] = { all: false, funcs: {} };

    if (fn === "*") {
      byRole[role].all = true;
    } else {
      byRole[role].funcs[fn] = true;
    }
  }

  var out = { byRole: byRole };
  cache.put(VEKTOR_ACL_CACHE_ACESSOS, JSON.stringify(out), VEKTOR_ACL_CACHE_TTL);
  return out;
}

function vektorGetUserRole_() {
  // ✅ email do usuário logado no domínio
  var sess = (Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  if (!sess) throw new Error("Não foi possível identificar seu e-mail Google.");

  // ✅ fonte de verdade: VEKTOR_EMAILS (ATIVO + ROLE)
  var emails = vektorLoadEmailsRoleMap_();
  var rec = emails && emails.byEmail ? emails.byEmail[sess] : null;

  if (!rec || rec.ativo !== true) {
  throw new Error("Acesso não disponível. Solicite a liberação junto ao administrador do sistema.");
    }
  return { email: sess, role: rec.role };
}

function vektorAssertFunctionAllowed_(fnName) {
  var ctx = vektorGetUserRole_();
  var acessos = vektorLoadRoleAllowedFunctions_();
  var rule = acessos.byRole[ctx.role];

  // Se o role não existe na VEKTOR_ACESSOS, então não tem acesso a nada.
  if (!rule) throw new Error("Não disponível para o seu perfil.");

  if (rule.all === true) return ctx;
  if (rule.funcs && rule.funcs[String(fnName || "").trim()] === true) return ctx;

  throw new Error("Não disponível para o seu perfil.");
}

// =======================
// VEKTOR - SESSAO, TEMPO DE LOGIN
// =======================
var VEKTOR_SESSION_TTL_SECONDS = 3 * 60 * 60; // 3 horas ou 5 minutos

function vektorCreateSessionToken_(email) {
  // token aleatório + carimbo
  var token = Utilities.getUuid() + "-" + new Date().getTime();
  var cache = CacheService.getScriptCache();

  // Armazena no cache: token -> email
  cache.put("VEKTOR_SESSION_" + token, String(email || ""), VEKTOR_SESSION_TTL_SECONDS);
  return token;
}

// =======================
// VEKTOR - USUÁRIOS ATIVOS HOJE (por dia)
// =======================

function vektorActiveUsersKey_(tz) {
  var z = tz || Session.getScriptTimeZone() || "America/Sao_Paulo";
  var now = new Date();
  return Utilities.formatDate(now, z, "yyyy-MM-dd");
}

// Armazena em Script Properties um JSON com e-mails únicos do dia
function vektorTrackActiveUserToday_(email) {
  var em = String(email || "").trim().toLowerCase();
  if (!em) return;

  var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
  var dayKey = vektorActiveUsersKey_(tz);
  var propKey = "VEKTOR_ACTIVE_USERS_" + dayKey;

  var lock = LockService.getScriptLock();
  lock.waitLock(15000);

  try {
    var props = PropertiesService.getScriptProperties();
    var raw = props.getProperty(propKey) || "[]";
    var arr;
    try { arr = JSON.parse(raw); } catch (_) { arr = []; }
    if (!Array.isArray(arr)) arr = [];

    if (arr.indexOf(em) < 0) {
      arr.push(em);
      props.setProperty(propKey, JSON.stringify(arr));
    }
  } finally {
    lock.releaseLock();
  }
}

function vektorGetActiveUsersTodayCount_(tz) {
  var z = tz || Session.getScriptTimeZone() || "America/Sao_Paulo";
  var dayKey = Utilities.formatDate(new Date(), z, "yyyy-MM-dd");
  var propKey = "VEKTOR_ACTIVE_USERS_" + dayKey;

  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty(propKey) || "[]";

  try {
    var arr = JSON.parse(raw);
    if (!Array.isArray(arr)) return 0;
    return arr.length;
  } catch (e) {
    return 0;
  }
}

function vektorValidateSessionToken_(token) {
  var t = String(token || "").trim();
  if (!t) return { ok: false, error: "Token vazio." };

  var emailSessao = (Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  if (!emailSessao) return { ok: false, error: "Não foi possível identificar seu e-mail Google." };

  // ✅ VEKTOR_EMAILS (ATIVO) é a fonte de verdade
  if (!isWhitelistedEmail_(emailSessao)) {
    return { ok: false, error: "Acesso negado: usuário não habilitado no Vektor." };
  }

  try {
  vektorGetUserRole_(); // garante ATIVO
} catch (e) {
  return { ok: false, error: "Não disponível para o seu perfil." };
}

  var cache = CacheService.getScriptCache();
  var emailDoToken = (cache.get("VEKTOR_SESSION_" + t) || "").trim().toLowerCase();

  if (!emailDoToken) return { ok: false, error: "Sessão expirada ou inválida. Faça login novamente." };
  if (emailDoToken !== emailSessao) return { ok: false, error: "Sessão não corresponde ao usuário logado." };

    // ✅ marca o usuário como "ativo hoje" sempre que a sessão for validada
    try {
      vektorTrackActiveUserToday_(emailSessao);
    } catch (eTrack) {
      // não quebra validação por falha de log
    }

  return { ok: true, email: emailSessao };
}

function validarSessaoVektor(token) {
  return vektorValidateSessionToken_(token);
}

function encerrarSessaoVektor(token) {
  try {
    var t = String(token || "").trim();
    if (!t) return { ok: true };

    var cache = CacheService.getScriptCache();
    cache.remove("VEKTOR_SESSION_" + t);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

/**
 * Serve o HTML do chat (index.html)
 */
function doGet(e) {
  // pega o e-mail do usuário logado no domínio
  var email = Session.getActiveUser().getEmail() || "";

  var nome = "";
  if (email) {
    var antesArroba = email.split("@")[0];           // ex: rodrigo.lisboa
    var partes = antesArroba.split(/[.\s_]+/);       // ["rodrigo","lisboa"]

    var nomeFormatado = "";
    if (partes.length > 0) {
      nomeFormatado =
        partes[0].charAt(0).toUpperCase() + partes[0].slice(1);
    }

    nome = nomeFormatado;
  }

 var role = "Sem acesso";
try {
  role = vektorGetUserRole_().role;
} catch (e) {
  role = "Sem acesso";
}

  var template = HtmlService
    .createTemplateFromFile('index');

  // passa o nome para o HTML
  template.userName  = nome;
  // 👇 passa também o e-mail bruto
  template.userEmail = email;
  template.userRole  = role;

  return template
  .evaluate()
  .setTitle('Grupo SBF | Vektor')
  .setFaviconUrl('https://raw.githubusercontent.com/rslisboa/cartoesclara/ce030011860b128fc826cd763582f60c0d68890c/Logo_Vektor_0503_2.png')
  .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);

  }

// =====================================================
// ✅ VERTEX AI — Assistente Política Clara
// =====================================================

const VEKTOR_VERTEX_PROJECT_ID = "genai4sap-data-lake";
const VEKTOR_VERTEX_LOCATION   = "us-central1";
const VEKTOR_VERTEX_MODEL      = "gemini-2.5-flash"; // alternativa: "gemini-2.5-pro"

// -----------------------------
// Entrada principal chamada pelo front
// -----------------------------
function vektorPolicyAssistantAsk(question, history) {
  vektorAssertFunctionAllowed_("vektorPolicyAssistantAsk");

  try {
    question = String(question || "").trim();
    if (!question) return { ok: false, error: "Pergunta vazia." };

    // ✅ small talk: não chama RAG/Vertex
    var qLow = question.toLowerCase();
    if (
      qLow === "obrigado" || qLow === "obrigada" ||
      qLow === "valeu" || qLow === "vlw" ||
      qLow === "ok" || qLow === "beleza" ||
      qLow === "show" || qLow === "perfeito" ||
      qLow === "obg" || qLow === "brigado" || qLow === "brigada"
    ) {
      return { ok: true, answer: "De nada! Se quiser, me diga sua dúvida sobre a política." };
    }

    var policyText = vektorPolicyLoadText_();
    if (!policyText) return { ok: false, error: "Não consegui ler o documento da política." };

    // ✅ chunking por seção (vale pra política toda)
    var chunks = vektorPolicyChunkText_(policyText, 1200);

    // ------------------------------
    // ✅ roteamento genérico por intenção
    // ------------------------------
    function routeSectionIds_(q) {
      var s = String(q || "").toLowerCase();

      var wantsPermission = /(pode|posso|permitid|proibid|autorizad|restriç|restric)/.test(s);
      var wantsAccountability = /(nota|comprovante|cupom|recibo|etiqueta|descri|prestação|prestacao|48\s*h|bloque)/.test(s);
      var wantsLimit = /(limite|aumento de limite|aumentar limite|alterar limite|mudan[cç]a de limite|consultar limite|ver o limite|limite dispon[ií]vel|dia\s*06|fatura|reestabelec)/.test(s);
      var wantsFraud = /(fraude|auditoria|monitoramento|canal|denúnc|denunc)/.test(s);
      var wantsRoles = /(responsabil|portador|financeiro|contas a receber|gerente|líder|lider|supervisor)/.test(s);
      var wantsDefs = /(defini|sigla|o que é|o que e|conceito)/.test(s);
      var wantsServicenow = /(servicenow|chamado|abertura de chamado|fluxo|solicita)/.test(s);
      var wantsStoreChange = /(trocar de loja|troca de loja|troca de gerente|mudan[cç]a entre lojas|mudar de loja|transfer[êe]ncia|transferir|loja anterior|loja nova)/.test(s);
      var wantsLabels = /(etiqueta|etiquetas|rotul|r[oó]tulo|classific|sap|quadro de etiquetas|anexo i)/.test(s);
      var wantsTermo = /(termo de responsabilidade|termo clara|aceite|formaliza[cç][aã]o do aceite|formulario|google forms|pdf assinado|reenviado ao departamento financeiro|termo assinado)/.test(s);

      var sec = [];
      // Regras principais por tema
      if (wantsPermission) sec.push("9");                  // Restrições de uso
      if (wantsAccountability) sec.push("8", "8.1");       // Prestação / Bloqueio
      if (wantsLimit) sec.push("7", "Anexo II");           // Limite dos cartões
      if (wantsFraud) sec.push("10", "10.1");              // Monitoramento / Canal
      if (wantsRoles) sec.push("5");                       // Responsabilidades
      if (wantsDefs) sec.push("4");                        // Definições
      if (wantsServicenow) sec.push("Anexo II");           // Solicitações no ServiceNow
      if (wantsStoreChange) sec.push("Anexo II", "5");
      if (wantsLabels) sec.push("Anexo I", "8");
      if (wantsTermo) sec.push("5.1");

      // fallback: se nada casou, não força seção (deixa ranker escolher)
      return sec;
    }

    var wanted = routeSectionIds_(question);

    // ------------------------------
    // ✅ índice simples por seção (a partir do prefixo "§ <id> | <título>")
    // ------------------------------
    var byId = {};
    (chunks || []).forEach(function(c){
      var m = String(c || "").match(/^§\s*([^\|]+)\|\s*(.+?)\s*\n/);
      if (!m) return;
      var id = String(m[1] || "").trim();
      if (!id) return;
      if (!byId[id]) byId[id] = c;
    });

    // ------------------------------
    // ✅ seleciona seeds por seção roteada
    // ------------------------------
    var seed = [];
    (wanted || []).forEach(function(id){
      if (byId[id]) seed.push(byId[id]);
    });

    // ------------------------------
    // ✅ ranker para completar com os mais relevantes
    // ------------------------------
    var topK = 10; // limite final que você já usa

    // ============================================
// ✅ SEED FORÇADO POR TEMA (anti “base insuficiente”)
// Cole antes do ranker: vektorPolicyPickTopChunks_(...)
// ============================================
function vektorNormNoAcc_(s){
  s = String(s || "").toLowerCase();
  try { s = s.normalize("NFD").replace(/[\u0300-\u036f]/g, ""); } catch(_){}
  return s;
}

function vektorSeedWindow_(chunks, anchor){
  var a = vektorNormNoAcc_(anchor);
  if (!a) return null;

  for (var i = 0; i < (chunks || []).length; i++) {
    var c = String(chunks[i] || "");
    if (!c) continue;

    if (vektorNormNoAcc_(c).indexOf(a) >= 0) {
      return c; // retorna o chunk completo, com prefixo "§ ..."
    }
  }

  return null;
}

function vektorSeedPushUnique_(arr, chunk){
  if (!chunk) return;
  if (arr.indexOf(chunk) < 0) arr.push(chunk);
}

// --- Detecta tema da pergunta
var qn = vektorNormNoAcc_(question);

// janela padrão (equilíbrio)
var FORCE_WINDOW = 1600;

// Map de temas -> regex -> anchors
var TOPIC_SEEDS = [
  // 5.2 / 4.12 (férias, afastamento, demissão)
  {
    re: /\b(ferias|afastamento|licenca|ausencia|portador temporario|temporario|substituto|substituicao|demissao|desligamento)\b/,
    anchors: [
      "5.2.", "procedimento em casos de afastamento", "demissao do gerente",
      "4.12.", "portador temporario", "nos afastamentos temporarios"
    ]
  },

  // 7 (limite mensal/diário, ciclo 06)
  {
    re: /\b(limite|mensal|diario|por dia|por mes|ciclo|06|faturamento|aumento de limite)\b/,
    anchors: [
      "7.", "7. limite dos cartoes",
      "importante destacar que o limite disponivel e mensal", // frase assinatura
      "o limite e reestabelecido", "todo dia 06",
      "solicitacoes de aumento de limite", "servicenow"
    ]
  },

  // 8 (prestação de contas / 48h / comprovante / etiqueta / descrição)
  {
    re: /\b(prestacao de contas|prestar contas|comprovante|nota fiscal|recibo|etiqueta|descricao|48 horas|prazo)\b/,
    anchors: [
      "8.", "8. prestacao de contas",
      "48 horas", "8.1.", "bloqueio preventivo do cartao",
      "inserir etiqueta", "anexar o comprovante fiscal", "preencher o campo \"descricao\""
    ]
  },

  // 8.1 bloqueio + desbloqueio via ServiceNow (Anexo II)
  {
    re: /\b(bloqueio|desbloqueio|cartao bloqueado|bloqueado preventivamente|regularizacao|servicenow)\b/,
    anchors: [
      "8.1.", "bloqueio preventivo do cartao",
      "o desbloqueio devera ser solicitado", "servicenow",
      "anexo ii", "solicitacao de desbloqueio de cartao"
    ]
  },

  // contestação (prazo 2 dias úteis) – aparece na seção de prestação/contestação
  {
    re: /\b(contestacao|contestar|chargeback|disputa|compra irregular|suporte da clara)\b/,
    anchors: [
      "contestacao", "suporte da clara", "2 dias uteis", "contato@clara.com"
    ]
  },

  // 9 – restrições de uso: saques, cashback, milhas, familiares etc.
  {
    re: /\b(restricao|proibido|vedado|nao pode|saque|dinheiro|cashback|milhas|fidelidade|cpf|familiares|conflito de interesses)\b/,
    anchors: [
      "9.", "9. restricoes de uso",
      "nao e permitido realizar adiantamentos", "saques",
      "cashback", "milhas", "cpf",
      "conflito de interesses"
    ]
  },

  // 9.1 – despesas pessoais / ressarcimento (prazo 2 dias úteis)
  {
    re: /\b(despesa pessoal|uso indevido|ressarcimento|reembolsar|devolver|pix|transferencia)\b/,
    anchors: [
      "9.1.", "utilizacao indevida do cartao", "despesas pessoais",
      "ressarcimento", "2 dias uteis",
      "despesa pessoal - uso indevido"
    ]
  },

  // 9.2 – patrimonial / itens de valor elevado (obrigatório compras)
  {
    re: /\b(patrimonial|infraestrutura|valor elevado|eletrodomestico|eletronico|notebook|celular|impressora|moveis|mobiliario|geladeira|micro-ondas|compras)\b/,
    anchors: [
      "9.2.", "aquisição de itens patrimoniais", "valor elevado",
      "obrigatoria a solicitacao via area de compras"
    ]
  },

  // 10 – monitoramento / auditoria / fraudes (típico)
  {
    re: /\b(auditoria|monitoramento|fraude|fiscalizacao|compliance|medidas disciplinares|medidas corretivas)\b/,
    anchors: [
      "10.", "monitoramento", "auditoria", "fraudes",
      "medidas disciplinares", "medidas corretivas"
    ]
  },

  // Anexos: etiquetas / ServiceNow
  {
    re: /\b(etiquetas|codigo sap|anexo i|anexo ii|servicenow|chamado)\b/,
    anchors: [
      "anexo i", "quadro de etiquetas", "codigo sap",
      "anexo ii", "solicitacoes no servicenow"
    ]
  },

  // 5.1 Sobre o termo de responsabilidade

  {
  re: /\b(termo de responsabilidade|termo clara|aceite|google forms|formulario|formulário|pdf assinado|termo assinado)\b/,
  anchors: [
    "termo de responsabilidade",
    "formalização do aceite",
    "formalizacao do aceite",
    "a liberação para uso do cartão está condicionada à formalização do aceite",
    "o gerente receberá um link de um formulário",
    "google forms",
    "será gerado o documento do termo em formato pdf",
    "deverá ser assinado e reenviado ao departamento financeiro"
  ]
}
];

// Aplica: se match no tema, injeta janelas ao redor das âncoras no seed
for (var t = 0; t < TOPIC_SEEDS.length; t++){
  var topic = TOPIC_SEEDS[t];
  if (!topic.re.test(qn)) continue;

  var anchors = topic.anchors || [];
  for (var a = 0; a < anchors.length; a++){
    var winTxt = vektorSeedWindow_(chunks, anchors[a]);
    vektorSeedPushUnique_(seed, winTxt);
  }
}

// (opcional) limite de seed para não “estourar” tokens por acidente
if (seed.length > 6) seed = seed.slice(0, 6);

    var ranked = vektorPolicyPickTopChunks_(question, chunks, topK);

    // monta lista final sem duplicar: seeds + ranked
    var finalChunks = [];
    seed.forEach(function(c){
      if (finalChunks.indexOf(c) < 0) finalChunks.push(c);
    });
    ranked.forEach(function(c){
      if (finalChunks.indexOf(c) < 0) finalChunks.push(c);
    });

    // corta no máximo 5 para não explodir tokens
    finalChunks = finalChunks.slice(0, topK);

    // ✅ cap por chunk (reduz custo sem matar cobertura)
      var MAX_CHARS_PER_CHUNK = 1600;
      finalChunks = finalChunks.map(function(c){
        c = String(c || "");
        if (c.length <= MAX_CHARS_PER_CHUNK) return c;
        return c.slice(0, MAX_CHARS_PER_CHUNK) + "…";
      });

    // fallback defensivo
    if (!finalChunks.length) finalChunks = [policyText.substring(0, 3500)];

    // gera resposta
      var answer = vektorVertexGeneratePolicyAnswer_(question, finalChunks, history || []);

      // prefixa assinatura
      var sig = "gemini-2.5-flash";
      var finalAnswer = sig + "\n" + String(answer || "");

      // ✅ LOG na VEKTOR_POLICY_HIST (antes do return)
      try {
        var userEmail = "";
        try { userEmail = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase(); } catch(_) {}
        var sectionsCsv = (wanted || []).map(String).join(",");
        var assunto = vektorPolicyAssuntoFromSections_(wanted);

        vektorPolicyHistAppend_({
          userEmail: userEmail,
          assunto: assunto,
          sectionsCsv: sectionsCsv,
          question: question,
          answer: finalAnswer,
          model: sig
        });
      } catch (eLog) {
        // não quebra resposta por erro de log
      }

      // retorna
      return { ok: true, answer: finalAnswer };

  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

// =======================================
// POLICY HIST (Sheets): grava + consulta
// =======================================
var VEKTOR_POLICY_HIST_SS_ID = "1_XW0IqbYjiCPpqtwdEi1xPxDlIP2MSkMrLGbeinLIeI";
var VEKTOR_POLICY_HIST_TAB   = "VEKTOR_POLICY_HIST";

function vektorPolicyHistGetSheet_(){
  var ss = SpreadsheetApp.openById(VEKTOR_POLICY_HIST_SS_ID);
  var sh = ss.getSheetByName(VEKTOR_POLICY_HIST_TAB);
  if (!sh) throw new Error("Aba '" + VEKTOR_POLICY_HIST_TAB + "' não encontrada na planilha do histórico.");
  return sh;
}

function vektorPolicyHistEnsureHeader_(){
  var sh = vektorPolicyHistGetSheet_();
  var lastCol = Math.max(sh.getLastColumn(), 1);
  var hdr = sh.getRange(1, 1, 1, lastCol).getValues()[0] || [];
  var hdrTxt = hdr.map(function(x){ return String(x||"").trim(); });

  // header esperado
  var exp = ["createdAt","userEmail","assunto","sectionsCsv","question","answer","model"];
  var ok = exp.every(function(name){ return hdrTxt.indexOf(name) >= 0; });

  if (!ok) {
    // Se a planilha está vazia ou header diferente, sobrescreve a linha 1 com o padrão.
    sh.getRange(1,1,1,exp.length).setValues([exp]);
    sh.getRange(1,1,1,exp.length).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function vektorPolicyAssuntoFromSections_(wanted){
  // wanted vem do routeSectionIds_(question) (ex: ["Anexo I","8","Anexo II","5"])
  var s = (wanted || []).map(String);

  // heurística bem “seca” (sem inventar):
  if (s.indexOf("8") >= 0) return "Etiquetas / Regras de uso";
  if (s.indexOf("5") >= 0) return "Solicitações / ServiceNow";
  if (s.indexOf("Anexo II") >= 0) return "Solicitações";
  if (s.indexOf("Anexo I") >= 0) return "Política / Regras";
  return "Geral";
}

function vektorPolicyHistAppend_(payload){
  payload = payload || {};
  var sh = vektorPolicyHistEnsureHeader_();
  var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
  var ts = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

  var userEmail = String(payload.userEmail || "").trim().toLowerCase();
  var assunto = String(payload.assunto || "").trim();
  var sectionsCsv = String(payload.sectionsCsv || "").trim();
  var question = String(payload.question || "").trim();
  var answer = String(payload.answer || "").trim();
  var model = String(payload.model || "").trim();

  // grava
  sh.appendRow([ts, userEmail, assunto, sectionsCsv, question, answer, model]);
}

/**
 * Front chama: vektorPolicyHistGet(email, limit)
 * Retorna últimos registros do usuário (mais recentes primeiro).
 */
function vektorPolicyHistGet(email, limit){
  vektorAssertFunctionAllowed_("vektorPolicyHistGet");

  try {
    limit = Number(limit) || 200;
    if (limit < 1) limit = 1;
    if (limit > 500) limit = 500;

    email = String(email || "").trim().toLowerCase();
    if (!email) {
      // fallback: tenta pegar do contexto do Apps Script
      try { email = String(Session.getActiveUser().getEmail() || "").trim().toLowerCase(); } catch(_) {}
    }
    if (!email) return { ok:false, error:"Email do usuário não identificado." };

    var sh = vektorPolicyHistEnsureHeader_();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok:true, rows: [] };

    var values = sh.getRange(2, 1, lastRow - 1, 7).getValues(); // 7 cols fixas do header esperado
    // [createdAt,userEmail,assunto,sectionsCsv,question,answer,model]

    var out = [];
    for (var i = values.length - 1; i >= 0; i--) { // mais recentes primeiro
      var r = values[i];
      var rEmail = String(r[1] || "").trim().toLowerCase();
      if (rEmail !== email) continue;

      out.push({
        createdAt: String(r[0] || ""),
        userEmail: String(r[1] || ""),
        assunto: String(r[2] || ""),
        sectionsCsv: String(r[3] || ""),
        question: String(r[4] || ""),
        answer: String(r[5] || ""),
        model: String(r[6] || "")
      });

      if (out.length >= limit) break;
    }

    return { ok:true, rows: out };

  } catch (e) {
    return { ok:false, error: (e && e.message) ? e.message : String(e) };
  }
}

// -----------------------------
// Lê o Google Doc convertido da política
// -----------------------------
function vektorPolicyLoadText_() {
  var html = HtmlService.createHtmlOutputFromFile("policy_clara_source").getContent();

  var m = html.match(/<textarea[^>]*id=["']policy-clara-text["'][^>]*>([\s\S]*?)<\/textarea>/i);
  if (!m || !m[1]) return "";

  var text = String(m[1] || "");

  text = text
    .replace(/&nbsp;/g, " ")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, '"')
    .replace(/&#39;/g, "'");

  return text.trim();
}

// -----------------------------
// Quebra texto em blocos
// -----------------------------
function vektorPolicyChunkText_(text, chunkSize) {
  // chunkSize agora vira apenas "limite de segurança", mas chunking é por SEÇÃO
  text = String(text || "").replace(/\r/g, "").trim();
  var maxLen = Math.max(2500, Number(chunkSize || 1200) * 3); // limite por chunk (segurança)
  if (!text) return [];

  var lines = text.split("\n");

  function isHeaderLine_(ln){
    var s = String(ln || "").trim();

    // Seções principais: "1.\tObjetivo", "9.\tRestrições de Uso", etc
    if (/^\d+\.\s+/.test(s) || /^\d+\.\t/.test(s)) return true;

    // Subitens: "8.1.\t Bloqueio...", "9.3.\t Utilização..."
    if (/^\d+\.\d+\.\s+/.test(s) || /^\d+\.\d+\.\t/.test(s)) return true;

    // Anexos: "Anexo I – ..." / "Anexo II – ..."
    if (/^Anexo\s+[IVXLC]+\s*[\-–—]/i.test(s) || /^Anexo\s+[IVXLC]+\b/i.test(s)) return true;

    return false;
  }

  function parseHeader_(ln){
    var s = String(ln || "").trim();

    // Anexo I/II etc
    var ma = s.match(/^Anexo\s+([IVXLC]+)\s*[\-–—]?\s*(.+)?$/i);
    if (ma) {
      var roman = String(ma[1] || "").trim().toUpperCase();
      var titleA = String(ma[2] || "").trim();
      var idA = "Anexo " + roman;
      return { id: idA, title: (titleA ? (idA + " – " + titleA) : idA) };
    }

    // Subitem 9.1., 8.4.1., etc
    var ms = s.match(/^(\d+\.\d+(?:\.\d+)*)\.\s*(.+)?$/);
    if (ms) {
      return { id: String(ms[1]), title: String(ms[2] || "").trim() };
    }

    // Seção principal 9.
    var mm = s.match(/^(\d+)\.\s*(.+)?$/);
    if (mm) {
      return { id: String(mm[1]), title: String(mm[2] || "").trim() };
    }

    return null;
  }

  // 1) encontra todos os headers
  var headers = [];
  for (var i = 0; i < lines.length; i++) {
    if (!isHeaderLine_(lines[i])) continue;
    var h = parseHeader_(lines[i]);
    if (!h || !h.id) continue;
    headers.push({ idx: i, id: h.id, title: h.title || "" });
  }

  // se não achou headers, cai no chunking antigo por tamanho (para não quebrar)
  if (!headers.length) {
    var outFallback = [];
    for (var k = 0; k < text.length; k += Math.max(500, Number(chunkSize || 1200))) {
      var sub = text.substring(k, k + Math.max(500, Number(chunkSize || 1200))).trim();
      if (sub) outFallback.push(sub);
    }
    return outFallback;
  }

  // 2) cria chunks por seção/subseção/anexo
  var out = [];
  for (var j = 0; j < headers.length; j++) {
    var start = headers[j].idx;
    var end = (j + 1 < headers.length) ? headers[j + 1].idx : lines.length;

    var id = headers[j].id;
    var title = headers[j].title || "";

    var body = lines.slice(start, end).join("\n").trim();
    if (!body) continue;

    // prefixo para permitir roteamento por seção
    var prefix = "§ " + id + " | " + (title || "") + "\n";
    var chunk = prefix + body;

    // 3) se ficar enorme, quebra internamente por blocos em branco, mas mantém o prefixo
    if (chunk.length <= maxLen) {
      out.push(chunk);
    } else {
      var blocks = body.split(/\n\s*\n+/).map(function(s){ return String(s||"").trim(); }).filter(Boolean);

      var acc = "";
      for (var b = 0; b < blocks.length; b++) {
        var blk = blocks[b];
        if (!acc) {
          acc = blk;
        } else if ((acc.length + 2 + blk.length) <= (maxLen - prefix.length)) {
          acc += "\n\n" + blk;
        } else {
          out.push(prefix + acc.trim());
          acc = blk;
        }
      }
      if (acc) out.push(prefix + acc.trim());
    }
  }

  return out;
}

// -----------------------------
// Ranking simples por relevância
// -----------------------------
function vektorPolicyPickTopChunks_(question, chunks, topK) {
  topK = Math.max(3, Number(topK || 5));

  var stop = {
    "de":1,"da":1,"do":1,"das":1,"dos":1,"a":1,"o":1,"e":1,"ou":1,"em":1,"no":1,"na":1,"nos":1,"nas":1,
    "para":1,"por":1,"com":1,"sem":1,"uma":1,"um":1,"as":1,"os":1,"que":1,"como":1,"qual":1,"quais":1,
    "é":1,"ser":1,"são":1,"sao":1,"tem":1,"ter":1,"vai":1,"pode":1,"posso":1
  };

  function norm_(s){
    return String(s || "")
      .toLowerCase()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "");
  }

  var qNorm = norm_(question);

  var tokens = qNorm
    .split(/[^a-z0-9]+/g)
    .map(function(t){ return String(t||"").trim(); })
    .filter(function(t){ return t && t.length >= 3 && !stop[t]; });

  // bigramas simples para melhorar match (ex.: "prestacao contas", "bloqueio preventivo")
  var bigrams = [];
  for (var i = 0; i < tokens.length - 1; i++) {
    bigrams.push(tokens[i] + " " + tokens[i+1]);
  }

  // pergunta de permissão? aumenta peso de linguagem normativa
  var isPermissionQ = /(pode|posso|permitid|proibid|autorizad|restric)/.test(qNorm);

  var scored = (chunks || []).map(function(chunk, idx){
    var base = norm_(chunk);

    var score = 0;

    // 1) overlap por token
    tokens.forEach(function(tok){
      if (base.indexOf(tok) >= 0) score += 1;
    });

    // 2) overlap por bigram (mais forte)
    bigrams.forEach(function(bg){
      if (base.indexOf(bg) >= 0) score += 3;
    });

    // 3) boost normativo (genérico para toda política)
    // (fica mais forte quando a pergunta é "pode/não pode")
    function addIf_(needle, pts){
      if (base.indexOf(needle) >= 0) score += pts;
    }

    addIf_("restricoes de uso", isPermissionQ ? 35 : 15);
    addIf_("e proibido",       isPermissionQ ? 28 : 12);
    addIf_("proibido",         isPermissionQ ? 18 : 8);
    addIf_("nao autorizado",   isPermissionQ ? 18 : 8);
    addIf_("nao e permitido",  isPermissionQ ? 18 : 8);
    addIf_("deve",             4);
    addIf_("obrigatorio",      8);
    addIf_("prazo",            6);
    addIf_("48 horas",         10);
    addIf_("bloqueio preventivo", 14);
    addIf_("prestacao de contas", 14);
    addIf_("servicenow",       10);
    addIf_("troca de gerente", 18);
    addIf_("mudanca entre lojas", 18);
    addIf_("mudar de loja", 14);
    addIf_("trocar de loja", 14);
    addIf_("loja anterior", 12);
    addIf_("loja nova", 12);
    addIf_("cartao fisico deve ser levado", 18);
    addIf_("limite", 10);
    addIf_("aumento de limite", 18);
    addIf_("solicitacao de aumento de limite", 22);
    addIf_("servicenow", 12);
    addIf_("limite disponivel", 12);
    addIf_("dia 06", 10);
    addIf_("ciclo de faturamento", 12);
    addIf_("anexo i", 18);
    addIf_("quadro de etiquetas", 24);
    addIf_("etiqueta", 16);
    addIf_("codigo sap", 14);
    addIf_("agua potavel", 22);
    addIf_("agua potavel", 22); // (se você usa normalize sem acento, basta 1)
    addIf_("agua potavel", 22);

    // 4) leve boost se o chunk é uma seção “alta” (tem prefixo §)
    if (/^§\s*/.test(String(chunk || ""))) score += 2;

    return { idx: idx, chunk: chunk, score: score };
  });

  scored.sort(function(a,b){
    return (b.score - a.score) || (a.idx - b.idx);
  });

  return scored.slice(0, topK).map(function(x){ return x.chunk; });
}

// =====================================================
// ✅ Vertex usage tracking (tokens + custo estimado)
// Cole ACIMA de vektorVertexGeneratePolicyAnswer_
// =====================================================

function vektorVertexGetMonthKey_() {
  var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
  return Utilities.formatDate(new Date(), tz, "yyyyMM");
}

function vektorVertexGetUsagePropKey_() {
  return "VEKTOR_VERTEX_USAGE_" + vektorVertexGetMonthKey_();
}

function vektorVertexEstimateUsd_(modelName, promptTokens, outputTokens) {
  modelName = String(modelName || "").toLowerCase();
  promptTokens = Number(promptTokens || 0);
  outputTokens = Number(outputTokens || 0);

  // Preços oficiais (estimativa por token)
  // gemini-2.5-flash: input US$ 0.30 / 1M | output US$ 2.50 / 1M
  // gemini-2.5-pro:   input US$ 1.25 / 1M | output US$ 10.00 / 1M
  var inputPerM = 0.30;
  var outputPerM = 2.50;

  if (modelName.indexOf("2.5-pro") >= 0) {
    inputPerM = 1.25;
    outputPerM = 10.00;
  }

  var inUsd = (promptTokens / 1000000) * inputPerM;
  var outUsd = (outputTokens / 1000000) * outputPerM;

  return {
    inputUsd: inUsd,
    outputUsd: outUsd,
    totalUsd: inUsd + outUsd
  };
}

function vektorVertexTrackUsage_(payload) {
  try {
    payload = payload || {};

    var props = PropertiesService.getScriptProperties();
    var key = vektorVertexGetUsagePropKey_();

    var raw = props.getProperty(key);
    var acc = raw ? JSON.parse(raw) : {
      monthKey: vektorVertexGetMonthKey_(),
      calls: 0,
      promptTokens: 0,
      outputTokens: 0,
      totalTokens: 0,
      estimatedUsd: 0,
      lastModel: "",
      lastModelVersion: "",
      lastPromptTokens: 0,
      lastOutputTokens: 0,
      lastTotalTokens: 0,
      lastEstimatedUsd: 0,
      lastAt: "",
      lastUserEmail: ""
    };

    var promptTokens = Number(payload.promptTokens || 0);
    var outputTokens = Number(payload.outputTokens || 0);
    var totalTokens = Number(payload.totalTokens || (promptTokens + outputTokens));
    var model = String(payload.model || "");
    var modelVersion = String(payload.modelVersion || "");
    var userEmail = String(payload.userEmail || "");

    var usd = vektorVertexEstimateUsd_(modelVersion || model, promptTokens, outputTokens);

    acc.calls += 1;
    acc.promptTokens += promptTokens;
    acc.outputTokens += outputTokens;
    acc.totalTokens += totalTokens;
    acc.estimatedUsd += usd.totalUsd;

    acc.lastModel = model;
    acc.lastModelVersion = modelVersion;
    acc.lastPromptTokens = promptTokens;
    acc.lastOutputTokens = outputTokens;
    acc.lastTotalTokens = totalTokens;
    acc.lastEstimatedUsd = usd.totalUsd;
    acc.lastAt = new Date().toISOString();
    acc.lastUserEmail = userEmail;

    props.setProperty(key, JSON.stringify(acc));
        try {
      var fx = vektorFxGetUsdBrl_();
      var brl = fx ? (usd.totalUsd * fx) : 0;

      var sh = getOrCreateVertexCostSheet_();
      sh.appendRow([
        new Date(),
        vektorVertexGetMonthKey_(),
        userEmail,
        VEKTOR_VERTEX_PROJECT_ID,
        model,
        modelVersion,
        promptTokens,
        outputTokens,
        totalTokens,
        usd.totalUsd,
        brl
      ]);
    } catch (logErr) {
      Logger.log("Falha ao gravar log detalhado Vertex: " + (logErr && logErr.message ? logErr.message : String(logErr)));
    }
  } catch (e) {
    // não quebra o chat por falha de métrica
    Logger.log("Falha ao registrar uso Vertex: " + (e && e.message ? e.message : String(e)));
  }
}

function vektorVertexGetUsageSummary_() {
  var props = PropertiesService.getScriptProperties();
  var key = vektorVertexGetUsagePropKey_();
  var raw = props.getProperty(key);

  if (!raw) {
    return {
      monthKey: vektorVertexGetMonthKey_(),
      calls: 0,
      promptTokens: 0,
      outputTokens: 0,
      totalTokens: 0,
      estimatedUsd: 0,
      lastModel: VEKTOR_VERTEX_MODEL,
      lastModelVersion: VEKTOR_VERTEX_MODEL,
      lastPromptTokens: 0,
      lastOutputTokens: 0,
      lastTotalTokens: 0,
      lastEstimatedUsd: 0,
      lastAt: "",
      lastUserEmail: ""
    };
  }

  try {
    return JSON.parse(raw);
  } catch (e) {
    return {
      monthKey: vektorVertexGetMonthKey_(),
      calls: 0,
      promptTokens: 0,
      outputTokens: 0,
      totalTokens: 0,
      estimatedUsd: 0,
      lastModel: VEKTOR_VERTEX_MODEL,
      lastModelVersion: VEKTOR_VERTEX_MODEL,
      lastPromptTokens: 0,
      lastOutputTokens: 0,
      lastTotalTokens: 0,
      lastEstimatedUsd: 0,
      lastAt: "",
      lastUserEmail: ""
    };
  }
}

function vektorFmtBrlFromUsd_(usd){
  usd = Number(usd || 0);

  var fx = vektorFxGetUsdBrl_();
  if (!fx) return "R$ —";

  var brl = usd * fx;

  var s = (brl < 0.01) ? brl.toFixed(4) : brl.toFixed(2);
  s = s.replace(/0+$/,"").replace(/\.$/,"");
  s = s.replace(".", ",");

  return "R$ " + s;
}

function vektorFmtUsd_(value) {
  var v = Number(value || 0);
  if (v === 0) return "US$ 0";
  // < 1 centavo: mostra até 4 casas sem trailing zeros
  var s = (v < 0.01) ? v.toFixed(4) : v.toFixed(2);
  s = s.replace(/0+$/,"").replace(/\.$/,"");
  return "US$ " + s;
}

// ===============================
// FX USD->BRL (automático + cache diário)
// ===============================

function vektorFxKeyToday_(){
  var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
  return Utilities.formatDate(new Date(), tz, "yyyyMMdd");
}

function vektorFxFetchUsdBrlFromBCB_(dateObj){
  // PTAX - Banco Central (OData)
  // A API pode não ter cotação no fim de semana/feriado (por isso faremos fallback de dias).
  var mm = ("0" + (dateObj.getMonth() + 1)).slice(-2);
  var dd = ("0" + dateObj.getDate()).slice(-2);
  var yyyy = String(dateObj.getFullYear());
  var dateParam = mm + "-" + dd + "-" + yyyy; // MM-DD-YYYY

  var url =
    "https://olinda.bcb.gov.br/olinda/servico/PTAX/versao/v1/odata/" +
    "CotacaoDolarDia(dataCotacao=@dataCotacao)?@dataCotacao='" + dateParam + "'&$format=json";

  var resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  if (resp.getResponseCode() < 200 || resp.getResponseCode() >= 300) return null;

  var json = JSON.parse(resp.getContentText() || "{}");
  var arr = (json && json.value) ? json.value : [];
  if (!arr || !arr.length) return null;

  // Usa cotacaoVenda (mais conservador pra “custo”)
  var venda = Number(arr[0].cotacaoVenda || 0);
  if (!isFinite(venda) || venda <= 0) return null;

  return venda;
}

function vektorFxGetUsdBrl_(){
  var props = PropertiesService.getScriptProperties();
  var key = "VEKTOR_FX_USD_BRL_" + vektorFxKeyToday_();

  // 1) cache do dia
  var cached = props.getProperty(key);
  if (cached) {
    var v = Number(cached);
    if (isFinite(v) && v > 0) return v;
  }

  // 2) tenta hoje e volta até 7 dias (fim de semana/feriado)
  var d = new Date();
  d.setHours(0,0,0,0);

  var fx = null;
  for (var i = 0; i < 7; i++) {
    fx = vektorFxFetchUsdBrlFromBCB_(d);
    if (fx) break;
    d.setDate(d.getDate() - 1);
  }

  if (fx) {
    props.setProperty(key, String(fx));
    return fx;
  }

  // 3) fallback: procura a última cotação gravada (até 30 dias)
  for (var j = 1; j <= 30; j++) {
    var dj = new Date();
    dj.setDate(dj.getDate() - j);
    dj.setHours(0,0,0,0);

    var k2 = "VEKTOR_FX_USD_BRL_" + Utilities.formatDate(dj, Session.getScriptTimeZone() || "America/Sao_Paulo", "yyyyMMdd");
    var c2 = props.getProperty(k2);
    if (c2) {
      var v2 = Number(c2);
      if (isFinite(v2) && v2 > 0) return v2;
    }
  }

  return null; // sem cotação disponível
}

function vektorFmtUsdWithBrl_(usdValue){
  var usd = Number(usdValue || 0);
  var fx = vektorFxGetUsdBrl_();
  if (!fx) return vektorFmtUsd_(usd) + " | R$ —";

  var brl = usd * fx;
  // 4 casas pra ficar coerente com seu US$
  return vektorFmtUsd_(usd) + " | R$ " + brl.toFixed(4);
}

// -----------------------------
// Chamada ao Vertex Gemini
// -----------------------------

function vektorVertexGeneratePolicyAnswer_(question, topChunks, history) {
  var systemText =
  "Você é o Assistente da Política de Cartões Clara do Grupo SBF.\n" +
  "Responda com base apenas nos trechos fornecidos da política.\n" +
  "Não invente regras, exceções, prazos, valores, permissões ou interpretações que não estejam sustentados nesses trechos.\n" +
  "Responda em português do Brasil, de forma natural, clara, objetiva e profissional.\n" +
  "\n" +
  "Estilo e fluidez (sem perder rigor):\n" +
  "• Responda como um assistente humano, com frases curtas e conectivos naturais quando fizer sentido.\n" +
  "• Evite jargão desnecessário, repetição mecânica e tom robótico.\n" +
  "• Se a regra estiver explícita, comece com a conclusão e depois explique de forma simples.\n" +
  "• Se a pergunta for continuação da anterior, considere o contexto recente da conversa.\n" +
  "• Não use linguagem dura ou punitiva; seja objetivo e cordial.\n" +
  "\n" +
  "Quando não houver base suficiente:\n" +
  "• Se os trechos não trouxerem base clara para responder com segurança, diga isso de forma natural.\n" +
  "• Nesses casos, finalize com: \"Base insuficiente nos trechos fornecidos\".\n" +
  "\n" +
  "IMPORTANTE SOBRE OS TRECHOS:\n" +
  "• O <TÍTULO> é o texto após a barra \"|\" no cabeçalho do trecho.\n" +
  "• Cada trecho começa com um cabeçalho no formato: \"§ <SEÇÃO> | <TÍTULO>\".\n" +
  "• Use esse <SEÇÃO> como referência (ex.: § 9, § 8.1, § Anexo II).\n" +
  "\n" +
  "Regras de exatidão:\n" +
  "• Use \"deve\", \"não deve\", \"pode\" e \"não pode\" apenas quando isso estiver explícito nos trechos.\n" +
  "• Sempre que afirmar uma regra, cite a base no final usando a seção e o título: \"Base: § <SEÇÃO> — <TÍTULO>\".\n" +
  "• Se houver conflito entre trechos, sinalize o conflito e oriente validação com o time de Compliance.\n" +
  "\n" +
  "Formato da resposta:\n" +
  "• Comece com uma resposta direta.\n" +
  "• Depois, explique em 1 ou 2 parágrafos curtos.\n" +
  "• Se existir exceção, condição ou ação necessária, apresente como: \"Condição:\", \"Exceção:\" ou \"Ação:\".\n" +
  "• Final obrigatório: \"Base: § <SEÇÃO> — <TÍTULO>\" ou \"Base insuficiente nos trechos fornecidos\".\n" +
  "\n" +
  "Regras para consultas de ETIQUETAS (Anexo I):\n" +
  "• Se o usuário pedir a lista completa de etiquetas, não cole a tabela inteira.\n" +
  "• Responda com um resumo curto, com alguns exemplos, e informe que a lista completa está no Anexo I.\n" +
  "• Se o usuário perguntar por um item específico, indique diretamente a etiqueta correspondente, se ela estiver coberta pelos trechos.\n" +
  "\n" +
  "Pergunta de esclarecimento (somente se necessário):\n" +
  "• Se faltar um dado essencial para aplicar a regra, faça apenas 1 pergunta objetiva.\n" +
  "• Não gere hipóteses e não responda com suposições.\n";

  // ============================
  // ✅ Histórico curto (para resolver “isso/isso aí”)
  // ============================
  function vektorPolicyFmtHistory_(hist) {
    if (!Array.isArray(hist) || !hist.length) return "";
    // usa só os últimos 2 itens (1 turno) para economizar tokens
    var last = hist.slice(-4);
    var out = last.map(function(h){
      var role = String((h && h.role) || "").toLowerCase();
      var text = String((h && h.text) || "");
      // corta para não inflar tokens (ajuste se quiser)
      if (text.length > 600) text = text.substring(0, 600) + "…";
      if (role === "assistant") return "ASSISTENTE (anterior): " + text;
      return "USUÁRIO (anterior): " + text;
    }).join("\n");
    return out ? ("CONTEXTO DA CONVERSA (último turno):\n" + out + "\n\n") : "";
  }

  var histText = vektorPolicyFmtHistory_(history);

  var userText =
  histText +
  "PERGUNTA DO USUÁRIO (atual):\n" + String(question || "").trim() + "\n\n" +
  "TRECHOS DA POLÍTICA:\n\n" +
  (topChunks || []).map(function (t, i) {
    var s = String(t || "");
    // Cabeçalho esperado no início do chunk: "§ <SEÇÃO> | <TÍTULO>\n"
    var m = s.match(/^§\s*([^\|]+)\|\s*([^\n]+)\n/);
    var sec = m ? String(m[1] || "").trim() : "";
    var ttl = m ? String(m[2] || "").trim() : "";

    // Exibe o nome da seção/título para facilitar citação na "Base"
    var head = (sec || ttl) ? ("§ " + sec + " — " + ttl) : ("Trecho " + (i + 1));

    return "Trecho " + (i + 1) + " (" + head + "):\n" + s;
  }).join("\n\n");

  var url =
    "https://" + VEKTOR_VERTEX_LOCATION + "-aiplatform.googleapis.com/v1/" +
    "projects/" + VEKTOR_VERTEX_PROJECT_ID +
    "/locations/" + VEKTOR_VERTEX_LOCATION +
    "/publishers/google/models/" + VEKTOR_VERTEX_MODEL +
    ":generateContent";

  var payload = {
    contents: [
      {
        role: "user",
        parts: [
          { text: systemText + "\n\n" + userText }
        ]
      }
    ],
    generationConfig: {
      temperature: 0.35,
      topP: 0.9,
      maxOutputTokens: 3072,
      candidateCount: 1
    }
  };

  var resp = UrlFetchApp.fetch(url, {
    method: "post",
    contentType: "application/json",
    headers: {
      Authorization: "Bearer " + ScriptApp.getOAuthToken()
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });

  var code = resp.getResponseCode();
  var body = resp.getContentText() || "";

  if (code < 200 || code >= 300) {
    throw new Error("Vertex erro HTTP " + code + ": " + body);
  }

  var json = JSON.parse(body);

  // Junta TODAS as parts de texto
  var parts =
    json &&
    json.candidates &&
    json.candidates[0] &&
    json.candidates[0].content &&
    json.candidates[0].content.parts;

  var answer = "";
  if (Array.isArray(parts) && parts.length) {
    answer = parts.map(function (p) {
      return String((p && p.text) || "");
    }).join("\n").trim();
  }

  if (!answer) {
    throw new Error("Vertex retornou resposta vazia.");
  }

  // usageMetadata oficial do Vertex
  var usage = (json && json.usageMetadata) ? json.usageMetadata : {};
  var promptTokens = Number(usage.promptTokenCount || 0);
  var outputTokens = Number(usage.candidatesTokenCount || 0);
  var totalTokens = (promptTokens + outputTokens);
  var modelVersion = String((json && json.modelVersion) || VEKTOR_VERTEX_MODEL);

  // usuário atual (se disponível)
  var email = "";
  try {
    var ctx = vektorGetUserRole_();
    email = String((ctx && ctx.email) || "");
  } catch (_) {}

  // registra consumo do mês
  vektorVertexTrackUsage_({
    model: VEKTOR_VERTEX_MODEL,
    modelVersion: modelVersion,
    promptTokens: promptTokens,
    outputTokens: outputTokens,
    totalTokens: totalTokens,
    userEmail: email
  });

  return answer;
}

// ✅ ID da planilha de métricas do Vektor
// (a planilha que você mandou)
const VEKTOR_METRICAS_SHEET_ID = '18yAuYoAR33JOagqapxgwHh86F1WeD0mZcj9AIJym07k';

// ✅ Nome da aba onde os logs serão gravados
const VEKTOR_METRICAS_TAB_NAME = 'Vektor_Metricas';
const VEKTOR_VERTEX_COST_TAB_NAME = 'Vektor_Vertex_Cost';

// ✅ Pasta onde serão salvos os Termos de Responsabilidade
// (ID da pasta que você mandou no link)
const VEKTOR_PASTA_TERMOS_ID = '1Qj1oXoBxKnkGUA9hKoaF6Ak_9m7bb4wD';

// =======================
// LOG DE ALERTAS ENVIADOS
// =======================
const VEKTOR_ALERTAS_LOG_TAB = "Vektor_Alertas_Log"; 


// 🌐 BigQuery – validação de loja
const PROJECT_ID = 'cnto-data-lake';
const BQ_TABLE_LOJAS = '`cnto-data-lake.refined.cnt_ref_gld_dim_estrutura_loja`';

function getOrCreateAlertasLogSheet_() {
  var ss = SpreadsheetApp.openById(VEKTOR_METRICAS_SHEET_ID);
  var sh = ss.getSheetByName(VEKTOR_ALERTAS_LOG_TAB);
  if (!sh) {
    sh = ss.insertSheet(VEKTOR_ALERTAS_LOG_TAB);
    sh.appendRow(["timestamp", "tipo", "loja", "time", "detalhe", "destinatarios", "origem"]);
    sh.getRange(1, 1, 1, 7).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

/**
 * Registra um alerta enviado (linha simples, rastreável).
 * tipo: "LIMITE", "PENDENCIAS", "USO_IRREGULAR", "OFENSORAS", etc.
 */
function registrarAlertaEnviado_(tipo, loja, time, detalhe, destinatarios, origem) {
  try {
    var sh = getOrCreateAlertasLogSheet_();
    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    var ts = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

    sh.appendRow([
      ts,
      String(tipo || "").trim(),
      String(loja || "").trim(),
      String(time || "").trim(),
      String(detalhe || "").trim(),
      String(destinatarios || "").trim(),
      String(origem || "").trim()
    ]);
  } catch (e) {
    Logger.log("Falha ao registrar alerta enviado: " + (e && e.message ? e.message : e));
  }
}

/**
 * Retorna alertas recentes para o modal.
 * dias: janela (ex 14)
 * limit: limite de linhas (ex 200)
 */
function getAlertasRecentes(dias, limit) {
  vektorAssertFunctionAllowed_("getAlertasRecentes");
  try {
    dias = Number(dias) || 14;
    limit = Number(limit) || 200;

    var sh = getOrCreateAlertasLogSheet_();
    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, rows: [] };

    var values = sh.getRange(2, 1, lastRow - 1, 7).getValues(); // sem cabeçalho
    // values: [ts,tipo,loja,time,detalhe,destinatarios,origem]

    // filtra por janela
    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    var agora = new Date();
    var ini = new Date(agora);
    ini.setDate(agora.getDate() - (dias - 1));
    ini.setHours(0, 0, 0, 0);

    function parseTs_(s) {
  // Se vier como Date (Sheets converte), usa direto
  if (s instanceof Date) return s;

  // Espera "yyyy-MM-dd HH:mm:ss"
  if (!s) return null;

  var m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})\s(\d{2}):(\d{2}):(\d{2})$/);
  if (m) {
    return new Date(
      Number(m[1]),
      Number(m[2]) - 1,
      Number(m[3]),
      Number(m[4]),
      Number(m[5]),
      Number(m[6])
    );
  }

  // Fallback: tenta parse nativo (caso venha em outro formato)
  var d2 = new Date(String(s));
  return isNaN(d2.getTime()) ? null : d2;
}

    var out = [];
    for (var i = values.length - 1; i >= 0; i--) { // mais recentes primeiro
      var r = values[i];
      var d = parseTs_(r[0]);
      if (!d) continue;
      if (d < ini) break; // como está em ordem cronológica, pode parar

      var tsTxt = "";
        try {
          tsTxt = Utilities.formatDate(d, tz, "dd/MM/yyyy HH:mm:ss");
        } catch (e) {
          tsTxt = String(r[0] || "");
        }

        out.push({
          timestamp: tsTxt,                 // ✅ string serializável no WebApp
          tipo: String(r[1] || ""),
          loja: String(r[2] || ""),
          time: String(r[3] || ""),
          detalhe: String(r[4] || ""),
          destinatarios: String(r[5] || ""),
          origem: String(r[6] || "")
        });


      if (out.length >= limit) break;
    }

    return { ok: true, rows: out, meta: { dias: dias, limit: limit } };

  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }

}

/**
 * Retorna informações do "Estado Operacional" para o modal do HTML.
 * Inclui:
 * - BaseClara: referência simples (última linha/data, quando possível)
 * - Jobs: se houver propriedade registrada (fallback N/D)
 * - Serviços Google: quota de e-mail + status de execução
 * - BigQuery: healthcheck simples (SELECT 1)
 * - Alertas: última linha do log (se existir)
 */
function getStatusOperacionalVektor() {
  try {
    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";

    // -------------------------
    // 1) BaseClara (sinal simples)
    // -------------------------
    var baseClaraTxt = "—";
    try {
      var ssBase = SpreadsheetApp.openById(BASE_CLARA_ID);
      var shBase = ssBase.getSheetByName("BaseClara");
      if (shBase) {
        var lr = shBase.getLastRow();
        if (lr >= 2) {
          // tenta capturar uma data “de referência” na última linha (coluna A)
          var vA = shBase.getRange(lr, 1).getValue();
          if (vA instanceof Date) {
            baseClaraTxt = "Linha " + lr + " | " + Utilities.formatDate(vA, tz, "dd/MM/yyyy HH:mm");
          } else {
            baseClaraTxt = "Linha " + lr + " | " + Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm");
          }
        } else {
          baseClaraTxt = "BaseClara sem dados suficientes.";
        }
      } else {
        baseClaraTxt = "Aba BaseClara não encontrada.";
      }
    } catch (eBase) {
      baseClaraTxt = "Falha ao ler BaseClara: " + (eBase && eBase.message ? eBase.message : String(eBase));
    }

    // -------------------------
    // 2) Jobs (se você tiver alguma property de controle)
    // -------------------------
    var jobsTxt = "—";
    try {
      var props = PropertiesService.getScriptProperties();
      // se você já grava algo como VEKTOR_LAST_JOBS_RUN, vai aparecer; senão, N/D
      var lastJobs = props.getProperty("VEKTOR_LAST_JOBS_RUN") || "";
      jobsTxt = lastJobs ? lastJobs : "N/D (não registrado)";
    } catch (eJobs) {
      jobsTxt = "Falha ao ler status de jobs.";
    }

    // -------------------------
    // 3) Serviços Google / E-mail (quota)
    // -------------------------
    var googleTxt = "—";
    try {
      var quota = MailApp.getRemainingDailyQuota(); // pode lançar exceção se serviço estiver com problema
      googleTxt = "OK | Quota e-mail restante hoje: " + quota;
    } catch (eMail) {
      googleTxt = "Falha no MailApp/quota: " + (eMail && eMail.message ? eMail.message : String(eMail));
    }

    // -------------------------
    // 4) BigQuery (healthcheck SELECT 1)
    // -------------------------
    var bqTxt = "—";
    try {
      var req = { query: "SELECT 1 AS ok", useLegacySql: false };
      var r = BigQuery.Jobs.query(req, PROJECT_ID);
      bqTxt = (r && r.jobComplete === true) ? "OK" : "Indisponível (job não completou)";
    } catch (eBQ) {
      bqTxt = "Falha BigQuery: " + (eBQ && eBQ.message ? eBQ.message : String(eBQ));
    }

    // -------------------------
    // 5) Alertas (último envio registrado)
    // -------------------------
    var alertasTxt = "—";
    try {
      var sh = getOrCreateAlertasLogSheet_(); // você já tem essa função no projeto
      var lastRow = sh.getLastRow();
      if (lastRow >= 2) {
        var ts = sh.getRange(lastRow, 1).getValue(); // timestamp
        var tipo = sh.getRange(lastRow, 2).getValue(); // tipo
        var tsFmt = (ts instanceof Date) ? Utilities.formatDate(ts, tz, "dd/MM/yyyy HH:mm:ss") : String(ts || "");
        alertasTxt = "Último: " + tsFmt + " | " + String(tipo || "");
      } else {
        alertasTxt = "Sem registros recentes.";
      }
    } catch (eAl) {
      alertasTxt = "Falha ao ler log de alertas.";
    }

    // -------------------------
    // 6) Status geral (regra simples)
    // -------------------------
    var geralTxt = "Operando";
    var temFalhaGoogle = String(googleTxt).toLowerCase().indexOf("falha") !== -1;
    var temFalhaBQ = String(bqTxt).toLowerCase().indexOf("falha") !== -1 || String(bqTxt).toLowerCase().indexOf("indispon") !== -1;

    if (temFalhaGoogle && temFalhaBQ) geralTxt = "Instável (Google + BigQuery)";
    else if (temFalhaGoogle) geralTxt = "Instável (Serviços Google/E-mail)";
    else if (temFalhaBQ) geralTxt = "Instável (BigQuery)";

    return {
      ok: true,
      baseclara: baseClaraTxt,
      jobs: jobsTxt,
      google: googleTxt,
      bigquery: bqTxt,
      alertas: alertasTxt,
      geral: geralTxt
    };

  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

// =====================================================
// MEUS ALERTAS (Transações por Etiqueta) - Back-end
// =====================================================

function getOrCreateUserAlertsSheet_() {
  // Reaproveita o mesmo spreadsheet do log de alertas
  var logSh = getOrCreateAlertasLogSheet_();
  var ss = logSh.getParent();

  var name = "VEKTOR_USER_ALERTS";
  var sh = ss.getSheetByName(name);
  if (!sh) {
  sh = ss.insertSheet(name);
  sh.appendRow([
    "alertId", "ownerEmail", "ownerRole",
    "createdAt", "isActive",
    "freq", "windowDays",
    "time", "sendAt", "lojasCsv", "etiqueta",
    "lastRunAt", "lastRowCount", "alertType"
  ]);
    } else {
      // MIGRAÇÃO: adiciona sendAt se não existir
      var head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
      if (head.indexOf("sendAt") < 0) {
        // insere a coluna sendAt logo após "time"
        var iTime = head.indexOf("time");
        var insertPos = (iTime >= 0 ? iTime + 2 : head.length + 1); // 1-based
        sh.insertColumnAfter(insertPos - 1);
        sh.getRange(1, insertPos).setValue("sendAt");
      }
      // MIGRAÇÃO: adiciona alertType se não existir (default: TRANSACOES)
      head = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
      if (head.indexOf("alertType") < 0) {
        var lastCol = sh.getLastColumn();
        sh.insertColumnAfter(lastCol);
        sh.getRange(1, lastCol + 1).setValue("alertType");

        // Preenche as linhas existentes com default TRANSACOES
        var lr = sh.getLastRow();
        if (lr >= 2) {
          var fill = [];
          for (var r = 2; r <= lr; r++) fill.push(["TRANSACOES"]);
          sh.getRange(2, lastCol + 1, lr - 1, 1).setValues(fill);
        }
      }
    }
    return sh;}

function getOrCreateUserAlertsRunsSheet_() {
  var logSh = getOrCreateAlertasLogSheet_();
  var ss = logSh.getParent();

  var name = "VEKTOR_USER_ALERT_RUNS";
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.appendRow([
      "runId", "alertId", "ownerEmail",
      "runAt", "periodoIni", "periodoFim",
      "rowCount", "rowsJsonPreview"
    ]);
  }
  return sh;
}

function getTimesELojasParaAlertasVektor() {
  vektorAssertFunctionAllowed_("getTimesELojasParaAlertasVektor");

  try {
    // ✅ ACL por Emails SOMENTE para Gerentes_Reg
    var ctx = vektorGetUserRole_(); // { email, role }
    var role = String((ctx && ctx.role) || "").trim();
    var email = String((ctx && ctx.email) || "").trim().toLowerCase();

    var allowedSet = null; // { "0092": true, "92": true, ... }
    if (role === "Gerentes_Reg") {
      var allowed = vektorGetAllowedLojasFromEmails_(email); // array ou null
      if (Array.isArray(allowed)) {
        allowedSet = {};
        allowed.forEach(function(x){
          x = String(x || "").trim();
          if (!x) return;
          allowedSet[x] = true;
          allowedSet[x.padStart(4, "0")] = true;
          // também guarda versão sem zeros (algumas bases usam "92")
          var dig = x.replace(/\D/g, "");
          if (dig) allowedSet[String(Number(dig) || "").trim()] = true;
        });
      }
    }

    var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
    var sh = ss.getSheetByName("BaseClara");
    if (!sh) return { ok:false, error:"Aba BaseClara não encontrada." };

    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok:true, times:[], lojasPorTime:{} };

    // ✅ Coluna R = Grupos (Time) | Coluna V = LojaNum
    // R..V => 18..22 (1-based), 5 colunas: [Grupos, S, T, U, LojaNum]
    var values = sh.getRange(2, 18, lastRow - 1, 5).getValues();

    var map = {}; // time -> { loja4:true }
    var timesSet = {};

    for (var i = 0; i < values.length; i++) {
      var r = values[i];

      var time = String(r[0] || "").trim();         // R
      var lojaNum = normalizarLojaNumero_(r[4]);    // V
      if (!time || !lojaNum) continue;

      var lojaStr = String(lojaNum);               // ex: "92"
      var loja4 = lojaStr.padStart(4, "0");        // ex: "0092"

      // ✅ aplica ACL somente para Gerentes_Reg (quando allowedSet não é null)
      if (allowedSet) {
        if (!allowedSet[lojaStr] && !allowedSet[loja4]) continue;
      }

      timesSet[time] = true;
      if (!map[time]) map[time] = {};
      map[time][loja4] = true; // sempre devolve loja no formato 4 dígitos
    }

    var times = Object.keys(timesSet).sort();
    var lojasPorTime = {};
    times.forEach(function(t){
      lojasPorTime[t] = Object.keys(map[t] || {}).sort();
    });

    return { ok:true, times: times, lojasPorTime: lojasPorTime };

  } catch (e) {
    return { ok:false, error: (e && e.message) ? e.message : String(e) };
  }
}

function criarAlertaEtiquetaVektor(payload) {
  try {
    payload = payload || {};

    var email = String(payload.email || "").trim().toLowerCase();
    var role = String(payload.role || "").trim();
    var freq = String(payload.freq || "DAILY").trim();
    var windowDays = Number(payload.windowDays || 30) || 30;
    var time = String(payload.time || "").trim();

    // ✅ FIX: ler sendAt do payload
    var sendAt = String(payload.sendAt || "").trim();

    var allowedTimes = { "11:30": true, "16:00": true };
    var alertType = String(payload.alertType || "TRANSACOES").trim();
    if (alertType !== "TRANSACOES" && alertType !== "PENDENCIAS") alertType = "TRANSACOES";

    var roleNorm = String(role || "").trim().toLowerCase();
    var canUsePendencias =
      roleNorm === "administrador" ||
      roleNorm === "gerentes_reg";

    if (alertType === "PENDENCIAS" && !canUsePendencias) {
      return { ok:false, error:"O tipo de alerta Pendências está disponível apenas para Administrador e Gerentes_Reg." };
    }

    // ✅ novo: pode vir array (multi) OU string (legado)
    var etiquetasArr = Array.isArray(payload.etiquetas)
      ? payload.etiquetas.map(function(x){ return String(x || "").trim(); })
      : [];
    etiquetasArr = etiquetasArr.filter(function(x){ return x && x !== "__ALL__"; });

    var etiquetaLegacy = String(payload.etiqueta || "").trim();
    var etiquetaFinalCsv = "";

    // Se veio array, usa array; se não veio, usa legado
    if (etiquetasArr.length) {
      etiquetaFinalCsv = etiquetasArr.join(" | ");
    } else if (etiquetaLegacy && etiquetaLegacy !== "__ALL__") {
      etiquetaFinalCsv = etiquetaLegacy;
    } else {
      // ✅ “todas”: deixa vazio no armazenamento (sem filtro)
      etiquetaFinalCsv = "";
    }

    // ✅ Pendências não usa etiqueta
    if (alertType === "PENDENCIAS") etiquetaFinalCsv = "";

    var lojas = Array.isArray(payload.lojas) ? payload.lojas.map(String) : [];
    lojas = lojas.map(function(s){ return String(s || "").trim(); }).filter(Boolean);

    if (!email) return { ok:false, error:"E-mail obrigatório." };
    if (!time) return { ok:false, error:"Time obrigatório." };

    // ✅ Se "Todos os times", lojas pode ser vazio (significa todas)
    if (time !== "__ALL__" && !lojas.length) {
      return { ok:false, error:"Selecione ao menos 1 loja." };
    }

    if (windowDays < 1 || windowDays > 365) return { ok:false, error:"Janela inválida (1..365)." };
    if (["DAILY","3D","WEEKLY","MONTHLY"].indexOf(freq) < 0) return { ok:false, error:"Frequência inválida." };

    // ✅ validação de horário obrigatório e restrito
    if (!sendAt) return { ok:false, error:"Horário obrigatório (11:30 ou 16:00)." };
    if (!allowedTimes[sendAt]) return { ok:false, error:"Horário inválido. Use somente 11:30 ou 16:00." };

    var sh = getOrCreateUserAlertsSheet_();
    var alertId = "AL" + Utilities.getUuid().replace(/-/g,"").slice(0, 12).toUpperCase();
    var now = new Date();

    sh.appendRow([
      alertId,
      email,
      role,
      now,
      true,
      freq,
      windowDays,
      time,
      sendAt,
      lojas.join(","),
      etiquetaFinalCsv, // ✅ vazio (todas) ou "A | B | C"
      "",               // lastRunAt
      "",               // lastRowCount
      alertType         // ✅ NOVO
    ]);

    // ✅ FIX: força sendAt como TEXTO (evita virar Date 30/12/1899)
    try {
      var row = sh.getLastRow(); // linha recém inserida
      var colSendAt = 9;         // sendAt = 9ª coluna
      var cell = sh.getRange(row, colSendAt);
      cell.setNumberFormat("@");
      cell.setValue(String(sendAt || "").trim());
    } catch (eFmt) {
      // não quebra criação se formatação falhar
    }

    return { ok:true, alertId: alertId };

  } catch (e) {
    return { ok:false, error: (e && e.message) ? e.message : String(e) };
  }
}

/**
 * Retorna Times e Lojas (por Time) para a tela "Meus Alertas".
 * SEMPRE por índice (BaseClara não muda de posição).
 *
 * BaseClara (A..W = 23 colunas):
 * - Loja (Alias do Cartão) = H = 8 => idx 7
 * - Time (Grupos)          = R = 18 => idx 17
 */
function getTimesELojasParaAlertasVektor() {
  vektorAssertFunctionAllowed_("getTimesELojasParaAlertasVektor");

  try {
    // ✅ ACL por Emails SOMENTE para Gerentes_Reg
    var ctx = vektorGetUserRole_(); // { email, role }
    var role = String((ctx && ctx.role) || "").trim();
    var email = String((ctx && ctx.email) || "").trim().toLowerCase();

    var allowedSet = null; // se null => sem restrição
    if (role === "Gerentes_Reg") {
      var allowed = vektorGetAllowedLojasFromEmails_(email); // array ou null
      if (Array.isArray(allowed)) {
        allowedSet = {};
        allowed.forEach(function(x){
          x = String(x || "").trim();
          if (!x) return;
          var dig = x.replace(/\D/g, "");
          if (!dig) return;

          var n = String(Number(dig) || "").trim();      // "92"
          var n4 = n.padStart(4, "0");                   // "0092"
          allowedSet[n] = true;
          allowedSet[n4] = true;
        });
      } else {
        // se não houver nenhuma loja liberada para o e-mail, retorna vazio
        allowedSet = {};
      }
    }

    var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
    var sh = ss.getSheetByName("BaseClara");
    if (!sh) return { ok:false, error:"Aba BaseClara não encontrada." };

    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok:true, times:[], lojasPorTime:{} };

    // ✅ BaseClara: R=Grupos(Time), V=LojaNum
    // R..V => col 18..22 (1-based), 5 colunas
    var values = sh.getRange(2, 18, lastRow - 1, 5).getValues();

    var map = {};      // time -> { loja4:true }
    var timesSet = {}; // time -> true

    for (var i = 0; i < values.length; i++) {
      var r = values[i];

      var time = String(r[0] || "").trim();        // R
      var lojaNum = normalizarLojaNumero_(r[4]);   // V
      if (!time || !lojaNum) continue;

      var lojaStr = String(lojaNum).replace(/\D/g, "");
      if (!lojaStr) continue;

      var lojaN = String(Number(lojaStr) || "").trim(); // "92"
      if (!lojaN) continue;

      var loja4 = lojaN.padStart(4, "0");               // "0092"

      // ✅ aplica ACL somente para Gerentes_Reg
      if (allowedSet) {
        if (!allowedSet[lojaN] && !allowedSet[loja4]) continue;
      }

      timesSet[time] = true;
      if (!map[time]) map[time] = {};
      map[time][loja4] = true;
    }

    var times = Object.keys(timesSet).sort(function(a,b){ return a.localeCompare(b, "pt-BR"); });

    var lojasPorTime = {};
    times.forEach(function(t){
      lojasPorTime[t] = Object.keys(map[t] || {}).sort(function(a,b){ return a.localeCompare(b); });
    });

    return { ok:true, times: times, lojasPorTime: lojasPorTime };

  } catch (e) {
    return { ok:false, error: (e && e.message) ? e.message : String(e) };
  }
}

/**
 * Lista todas as etiquetas existentes na BaseClara (coluna T).
 * Suporta células com múltiplas tags (separadas por vírgula, ponto e vírgula ou barra).
 */
function getEtiquetasDisponiveisVektor() {
  try {
    if (typeof vektorAssertFunctionAllowed_ === "function") {
      vektorAssertFunctionAllowed_("getEtiquetasDisponiveisVektor");
    }

    var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
    var sh = ss.getSheetByName("BaseClara");
    if (!sh) return { ok: false, error: "Aba BaseClara não encontrada." };

    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, etiquetas: [] };

    // lê só a coluna T (20) de 2..lastRow
    var colT = sh.getRange(2, 20, lastRow - 1, 1).getValues(); // T

    var set = {};

    // ✅ exclusões
    var EXCLUIR = {
      "AR": true,
      "POSTAGEM": true,
      "USO INDEVIDO - EXCLUSIVA FINANCEIRO": true,
      "NULL": true
    };

    colT.forEach(function (r) {
      var cell = String(r[0] || "").trim();
      if (!cell) return;

      cell
        .split(/[;,\/\|]+/g)
        .map(function (s) { return String(s || "").trim(); })
        .filter(Boolean)
        .forEach(function (t) {
          var tag = String(t || "").trim();
          if (!tag) return;

          var up = tag.toUpperCase().trim();
          if (EXCLUIR[up] === true) return;
          if (tag.toLowerCase().trim() === "null") return;

          set[tag] = true;
        });
    });

    return { ok: true, etiquetas: Object.keys(set).sort() };

  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

function listarAlertasUsuarioVektor(req) {
  try {
    req = req || {};
    var email = String(req.email || "").trim().toLowerCase();
    var role = String(req.role || "").trim();

    var adminFilterEmail = String(req.adminFilterEmail || "").trim().toLowerCase();
    var adminFilterRole  = String(req.adminFilterRole  || "").trim();

    var sh = getOrCreateUserAlertsSheet_();
    var values = sh.getDataRange().getValues();
    if (values.length < 2) return { ok: true, alertas: [] };

    var head = values[0].map(String);
    var idx = function (name) { return head.indexOf(name); };

    var iAlertId   = idx("alertId");
    var iOwner     = idx("ownerEmail");
    var iOwnerRole = idx("ownerRole"); // pode existir ou não
    var iActive    = idx("isActive");
    var iFreq      = idx("freq");
    var iWin       = idx("windowDays");
    var iTime      = idx("time");
    var iSendAt    = idx("sendAt");
    var iLojas     = idx("lojasCsv");
    var iEtq       = idx("etiqueta");
    var iLastRun   = idx("lastRunAt");
    var iType      = idx("alertType");


    var isAdmin = (role === "Administrador");

    // Normalização (evita falha por maiúsculas/minúsculas/espacos)
    var norm_ = function (s) { return String(s || "").trim().toLowerCase(); };

    // Carrega ACL 1 vez, só se precisar
    var acl = null;
    var getRoleFromAcl_ = function (emailLower) {
      try {
        if (!emailLower) return "";
        if (!acl) acl = vektorLoadEmailsRoleMap_(); // => { byEmail: { "a@b": {role:"X", ativo:true} } }
        if (!acl || !acl.byEmail || !acl.byEmail[emailLower]) return "";
        return String(acl.byEmail[emailLower].role || "").trim();
      } catch (e) {
        return "";
      }
    };

    var out = [];
    for (var r = 1; r < values.length; r++) {
      var row = values[r];

      var owner = String(row[iOwner] || "").trim().toLowerCase();

      // 1) ownerRole pela planilha
      var ownerRole = "";
      if (iOwnerRole >= 0) ownerRole = String(row[iOwnerRole] || "").trim();

      // 2) fallback: coluna "role" (se existir)
      if (!ownerRole) {
        var iRoleFallback = idx("role");
        ownerRole = (iRoleFallback >= 0) ? String(row[iRoleFallback] || "").trim() : "";
      }

      // 3) fallback final: ACL (fonte de verdade)
      if (!ownerRole) {
        ownerRole = getRoleFromAcl_(owner);
      }

      // Regra de visibilidade + filtros
      if (!isAdmin) {
        if (owner !== email) continue;
      } else {
        if (adminFilterEmail && owner.indexOf(adminFilterEmail) === -1) continue;

        // ✅ filtro por área/role (NORMALIZADO + com fallback ACL)
        if (adminFilterRole) {
          if (norm_(ownerRole) !== norm_(adminFilterRole)) continue;
        }
      }

      var lojasCsv = String(row[iLojas] || "");
      var lojasCount = lojasCsv ? lojasCsv.split(",").filter(Boolean).length : 0;

      var f = String(row[iFreq] || "").trim();
      var freqLabel =
        (f === "DAILY")   ? "Diário"  :
        (f === "3D")      ? "3 dias"  :
        (f === "WEEKLY")  ? "Semanal" :
        (f === "MONTHLY") ? "Mensal"  :
        (f || "—");

      var etiquetaRaw = String(row[iEtq] || "").trim();
      var etqs = etiquetaRaw
        ? etiquetaRaw.split("|").map(function (s) { return String(s || "").trim(); }).filter(Boolean)
        : [];
      var etiquetaCount = etqs.length;

      var etiquetaLabel = "";
      if (!etiquetaRaw) etiquetaLabel = "Todas";
      else if (etiquetaCount === 1) etiquetaLabel = etqs[0];
      else etiquetaLabel = etiquetaCount + " etiquetas";

      var tz = Session.getScriptTimeZone();

      function formatSendAt_(v) {
        if (!v) return "";
        // Se vier como Date (Sheets time), formata HH:mm
        if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
          return Utilities.formatDate(v, tz, "HH:mm");
        }
        // Se vier como texto, tenta normalizar
        var s = String(v).trim();
        // Se já estiver HH:mm, ok
        if (/^\d{2}:\d{2}$/.test(s)) return s;
        // Se vier com segundos HH:mm:ss, corta
        if (/^\d{2}:\d{2}:\d{2}$/.test(s)) return s.slice(0,5);
        return s; // fallback
      }

      out.push({
        alertId: String(row[iAlertId] || ""),
        ownerEmail: owner,
        ownerRole: ownerRole,
        isActive: row[iActive] === true,
        freq: f,
        freqLabel: freqLabel,
        windowDays: Number(row[iWin] || 30),
        time: String(row[iTime] || ""),
        sendAt: (iSendAt >= 0 ? formatSendAt_(row[iSendAt]) : ""),
        lojasCount: lojasCount,

        // etiquetas continuam existindo para Transações
        etiqueta: etiquetaRaw,
        etiquetaCount: etiquetaCount,
        etiquetaLabel: etiquetaLabel,

        // ✅ NOVO: tipo
        alertType: (iType >= 0 ? String(row[iType] || "TRANSACOES").trim() : "TRANSACOES"),
        alertTypeLabel: ((iType >= 0 ? String(row[iType] || "").trim() : "") === "PENDENCIAS" ? "Pendências" : "Transações"),

        lastRunAt: (row[iLastRun] instanceof Date)
          ? Utilities.formatDate(row[iLastRun], Session.getScriptTimeZone() || "America/Sao_Paulo", "dd/MM/yyyy HH:mm")
          : (row[iLastRun] ? String(row[iLastRun]) : "—")
      });
    }

    out.reverse();
    return { ok: true, alertas: out };

  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

function toggleAlertaEtiquetaVektor(req) {
  try {
    req = req || {};
    var alertId = String(req.alertId || "").trim();
    if (!alertId) return { ok:false, error:"alertId obrigatório." };

    var sh = getOrCreateUserAlertsSheet_();
    var values = sh.getDataRange().getValues();
    if (values.length < 2) return { ok:false, error:"Sem alertas." };

    var head = values[0].map(String);
    var iAlertId = head.indexOf("alertId");
    var iActive = head.indexOf("isActive");

    for (var r=1; r<values.length; r++) {
      if (String(values[r][iAlertId] || "") === alertId) {
        var cur = (values[r][iActive] === true);
        sh.getRange(r+1, iActive+1).setValue(!cur);
        return { ok:true, isActive: !cur };
      }
    }
    return { ok:false, error:"Alerta não encontrado." };

  } catch (e) {
    return { ok:false, error: (e && e.message) ? e.message : String(e) };
  }
}

function atualizarAlertaEtiquetaVektor(payload) {
  try {
    payload = payload || {};
    var alertId = String(payload.alertId || "").trim();
    if (!alertId) return { ok: false, error: "alertId obrigatório." };

    var email = String(payload.email || "").trim().toLowerCase();
    var role  = String(payload.role || "").trim();

    var freq = String(payload.freq || "DAILY").trim();
    var windowDays = Number(payload.windowDays || 30) || 30;
    var time = String(payload.time || "").trim();
    var allowedTimes = { "11:30": true, "16:00": true };
    var sendAt = String(payload.sendAt || "").trim();

    var alertType = String(payload.alertType || "TRANSACOES").trim();
    if (alertType !== "TRANSACOES" && alertType !== "PENDENCIAS") alertType = "TRANSACOES";

    var roleNorm = String(role || "").trim().toLowerCase();
    var canUsePendencias =
      roleNorm === "administrador" ||
      roleNorm === "gerentes_reg";

    if (alertType === "PENDENCIAS" && !canUsePendencias) {
      return { ok:false, error:"O tipo de alerta Pendências está disponível apenas para Administrador e Gerentes_Reg." };
    }

    // Multi etiquetas (mesma regra do criar)
    var etiquetasArr = Array.isArray(payload.etiquetas)
      ? payload.etiquetas.map(function(x){ return String(x || "").trim(); }).filter(Boolean)
      : [];
    etiquetasArr = etiquetasArr.filter(function(x){ return x && x !== "__ALL__"; });

    var etiquetaLegacy = String(payload.etiqueta || "").trim();
    var etiquetaFinalCsv = "";
    if (etiquetasArr.length) etiquetaFinalCsv = etiquetasArr.join(" | ");
    else if (etiquetaLegacy && etiquetaLegacy !== "__ALL__") etiquetaFinalCsv = etiquetaLegacy;
    else etiquetaFinalCsv = ""; // todas

    // ✅ Pendências não usa etiqueta: sempre vazio
    if (alertType === "PENDENCIAS") etiquetaFinalCsv = "";

    var lojas = Array.isArray(payload.lojas) ? payload.lojas.map(String) : [];
    lojas = lojas.map(function(s){ return String(s||"").trim(); }).filter(Boolean);

    // validações mínimas (iguais ao criar)
    if (!email) return { ok:false, error:"E-mail obrigatório." };
    if (!time) return { ok:false, error:"Time obrigatório." };
    if (time !== "__ALL__" && !lojas.length) return { ok:false, error:"Selecione ao menos 1 loja." };
    if (windowDays < 1 || windowDays > 365) return { ok:false, error:"Janela inválida (1..365)." };
    if (["DAILY","3D","WEEKLY","MONTHLY"].indexOf(freq) < 0) return { ok:false, error:"Frequência inválida." };

    if (!sendAt) return { ok:false, error:"Horário obrigatório (11:30 ou 16:00)." };
    if (!allowedTimes[sendAt]) return { ok:false, error:"Horário inválido. Use somente 11:30 ou 16:00." };

    var sh = getOrCreateUserAlertsSheet_();
    var values = sh.getDataRange().getValues();
    if (!values || values.length < 2) return { ok:false, error:"Sem alertas cadastrados." };

    var head = values[0].map(String);
    var idx = function (name) { return head.indexOf(name); };

    var iAlertId = idx("alertId");
    var iOwner   = idx("ownerEmail");
    var iFreq    = idx("freq");
    var iWin     = idx("windowDays");
    var iTime    = idx("time");
    var iSendAt  = idx("sendAt");
    var iLojas   = idx("lojasCsv");
    var iEtq     = idx("etiqueta");
    var iType    = idx("alertType");
    var iLastRun = idx("lastRunAt");
    var iLastCnt = idx("lastRowCount");

    if (iAlertId < 0) return { ok:false, error:"Cabeçalho inválido: alertId não encontrado." };

    // encontra linha
    var rowIndex = -1;
    for (var r = 1; r < values.length; r++) {
      if (String(values[r][iAlertId] || "").trim() === alertId) { rowIndex = r; break; }
    }
    if (rowIndex < 0) return { ok:false, error:"Alerta não encontrado." };

    // permissão: usuário comum só edita o próprio alerta
    var ownerEmail = (iOwner >= 0) ? String(values[rowIndex][iOwner] || "").trim().toLowerCase() : "";
    var isAdmin = (role === "Administrador");
    if (!isAdmin && ownerEmail && ownerEmail !== email) {
      return { ok:false, error:"Você não tem permissão para editar este alerta." };
    }

    // ✅ Se mudou freq e/ou sendAt, zera lastRunAt para permitir novo disparo no mesmo dia
    try {
      var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";

      var oldFreq = (iFreq >= 0) ? String(values[rowIndex][iFreq] || "").trim() : "";

      // Normaliza sendAt antigo (pode vir Date/number/string)
      var oldSendAtRaw = (iSendAt >= 0) ? values[rowIndex][iSendAt] : "";
      var oldSendAt = "";

      if (Object.prototype.toString.call(oldSendAtRaw) === "[object Date]" && !isNaN(oldSendAtRaw.getTime())) {
        oldSendAt = Utilities.formatDate(oldSendAtRaw, tz, "HH:mm");
      } else if (typeof oldSendAtRaw === "number" && isFinite(oldSendAtRaw)) {
        var totalMinutes = Math.round(oldSendAtRaw * 24 * 60);
        var hh = Math.floor(totalMinutes / 60) % 24;
        var mm = totalMinutes % 60;
        oldSendAt = (String(hh).padStart(2, "0") + ":" + String(mm).padStart(2, "0"));
      } else {
        oldSendAt = String(oldSendAtRaw || "").trim();
        if (/^\d{2}:\d{2}:\d{2}$/.test(oldSendAt)) oldSendAt = oldSendAt.slice(0, 5);
      }

      var newSendAt = String(sendAt || "").trim();
      if (/^\d{2}:\d{2}:\d{2}$/.test(newSendAt)) newSendAt = newSendAt.slice(0, 5);

      var changedSchedule = (oldFreq !== freq) || (oldSendAt !== newSendAt);

      if (changedSchedule) {
        if (iLastRun >= 0) sh.getRange(rowIndex + 1, iLastRun + 1).setValue("");
        if (iLastCnt >= 0) sh.getRange(rowIndex + 1, iLastCnt + 1).setValue("");
      }
    } catch (_) {}

    // atualiza campos
    if (iFreq   >= 0) sh.getRange(rowIndex + 1, iFreq   + 1).setValue(freq);
    if (iWin    >= 0) sh.getRange(rowIndex + 1, iWin    + 1).setValue(windowDays);
    if (iTime   >= 0) sh.getRange(rowIndex + 1, iTime   + 1).setValue(time);
    if (iSendAt >= 0) {
      var cell = sh.getRange(rowIndex + 1, iSendAt + 1);
      cell.setNumberFormat("@"); // texto
      cell.setValue(String(sendAt || "").trim());
    }
    if (iLojas  >= 0) sh.getRange(rowIndex + 1, iLojas  + 1).setValue(lojas.join(","));
    if (iEtq    >= 0) sh.getRange(rowIndex + 1, iEtq    + 1).setValue(etiquetaFinalCsv);

    // ✅ NOVO: salva alertType
    if (iType >= 0) sh.getRange(rowIndex + 1, iType + 1).setValue(alertType);

    return { ok:true, alertId: alertId };

  } catch (e) {
    return { ok:false, error: (e && e.message) ? e.message : String(e) };
  }
}

function excluirAlertaVektor(req) {
  try {
    req = req || {};
    var alertId = String(req.alertId || "").trim();
    if (!alertId) return { ok:false, error:"alertId obrigatório." };

    var sh = getOrCreateUserAlertsSheet_();
    var values = sh.getDataRange().getValues();
    if (values.length < 2) return { ok:false, error:"Sem alertas." };

    var head = values[0].map(String);
    var iAlertId = head.indexOf("alertId");

    for (var r = 1; r < values.length; r++) {
      if (String(values[r][iAlertId] || "") === alertId) {
        sh.deleteRow(r + 1);
        return { ok:true };
      }
    }

    return { ok:false, error:"Alerta não encontrado." };
  } catch (e) {
    return { ok:false, error: (e && e.message) ? e.message : String(e) };
  }
}

function buildAlertXlsxAttachment_(alertId, periodo, rows) {
  // 1. Cria planilha temporária
  var ss = SpreadsheetApp.create("Vektor_Alerta_" + alertId);
  var sh = ss.getSheets()[0];
  sh.setName("Alerta");

    // Detecta se é "Pendências" pela estrutura das linhas (sem depender do layout do front)
  var isPendencias = !!((rows && rows.length) && (rows[0] && (rows[0].pendencias != null || rows[0].titular != null)));

  var header = isPendencias
    ? ["Loja","Time","Data","Valor","Estabelecimento","Titular","Pendências"]
    : ["Loja","Time","Data","Estabelecimento","Valor","Etiqueta","Descrição"];

  sh.getRange(1, 1, 1, header.length).setValues([header]);

  if (rows && rows.length) {
    var values = rows.map(function (r) {
      if (isPendencias) {
        return [
          r.loja || "",
          r.time || "",
          r.data || "",
          r.valor || "",
          r.estabelecimento || "",
          r.titular || "",
          r.pendencias || ""
        ];
      }
      return [
        r.loja || "",
        r.time || "",
        r.data || "",
        r.estabelecimento || "",
        r.valor || "",
        r.etiqueta || "",
        r.descricao || ""
      ];
    });
    sh.getRange(2, 1, values.length, header.length).setValues(values);
  }

  sh.setFrozenRows(1);
  try { sh.autoResizeColumns(1, header.length); } catch (_) {}

  var fileId = ss.getId();

  // 2. EXPORTAÇÃO REAL PARA XLSX (isso é o que faltava)
  var url = "https://www.googleapis.com/drive/v3/files/" + fileId + "/export" +
            "?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

  var token = ScriptApp.getOAuthToken();
  var resp = UrlFetchApp.fetch(url, {
    headers: {
      Authorization: "Bearer " + token
    },
    muteHttpExceptions: true
  });

  // 3. Blob XLSX válido
  var blob = resp.getBlob().setName(
    "Vektor - Alerta " + alertId +
    " (" + periodo.inicio + " a " + periodo.fim + ").xlsx"
  );

  // 4. Limpa a planilha temporária
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (_) {}

  return blob;
}

function buildAlertAttachmentSmart_(alertId, periodo, rows) {
  // Heurística simples: se for muita linha, já vai de CSV
  // Ajuste esses limites conforme seu volume real.
  var CSV_ROW_LIMIT = 8000; // acima disso, XLSX tende a ficar grande/lento
  var CSV_SIZE_LIMIT_BYTES = 18 * 1024 * 1024; // ~18MB (buffer abaixo do limite do Gmail)

  // Se muito grande em linhas, já retorna CSV
  if (rows && rows.length > CSV_ROW_LIMIT) {
    return buildAlertCsvAttachment_(alertId, periodo, rows);
  }

  // Tenta XLSX e, se passar do tamanho, cai pra CSV
  var xlsx = buildAlertXlsxAttachment_(alertId, periodo, rows);
  if (xlsx && xlsx.getBytes && xlsx.getBytes().length > CSV_SIZE_LIMIT_BYTES) {
    return buildAlertCsvAttachment_(alertId, periodo, rows);
  }

  return xlsx;
}

function buildAlertCsvAttachment_(alertId, periodo, rows) {
  var esc = function (s) {
    // CSV com aspas e escape de aspas duplas
    var t = String(s == null ? "" : s);
    t = t.replace(/"/g, '""');
    return '"' + t + '"';
  };

    var isPendencias = !!((rows && rows.length) && (rows[0] && (rows[0].pendencias != null || rows[0].titular != null)));

  var header = isPendencias
    ? ["Loja","Time","Data","Valor","Estabelecimento","Titular","Pendências"]
    : ["Loja","Time","Data","Estabelecimento","Valor","Etiqueta","Descrição"];

  var lines = [];
  lines.push(header.map(esc).join(";")); // separador ; (pt-BR)

  (rows || []).forEach(function (r) {
    if (isPendencias) {
      lines.push([
        r.loja || "",
        r.time || "",
        r.data || "",
        r.valor || "",
        r.estabelecimento || "",
        r.titular || "",
        r.pendencias || ""
      ].map(esc).join(";"));
      return;
    }

    lines.push([
      r.loja || "",
      r.time || "",
      r.data || "",
      r.estabelecimento || "",
      r.valor || "",
      r.etiqueta || "",
      r.descricao || ""
    ].map(esc).join(";"));
  });

  var csv = lines.join("\n");

  return Utilities.newBlob(csv, "text/csv", 
    "Vektor - Alerta " + alertId + " (" + periodo.inicio + " a " + periodo.fim + ").csv"
  );
}

// Executa 1 alerta (preview=true: para uso do front; preview=false: para execução agendada)
function executarAlertaEtiquetaVektor(req) {
  try {
    req = req || {};
    var alertId = String(req.alertId || "").trim();
    var preview = !!req.preview;

    var sh = getOrCreateUserAlertsSheet_();
    var values = sh.getDataRange().getValues();
    if (!values || values.length < 2) return { ok: false, error: "Base de alertas vazia." };

    var head = values[0].map(String);
    var idx = function (name) { return head.indexOf(name); };

    var iAlertId  = idx("alertId");
    var iOwner    = idx("ownerEmail");
    var iOwnerRole= idx("ownerRole");        // ✅ novo (se existir na planilha)
    var iActive   = idx("isActive");
    var iFreq     = idx("freq");
    var iWin      = idx("windowDays");
    var iTime     = idx("time");
    var iLojas    = idx("lojasCsv");
    var iEtq      = idx("etiqueta");
    var iLastRun  = idx("lastRunAt");
    var iLastCnt  = idx("lastRowCount");
    var iType     = idx("alertType");

    var rowIdx = -1;
    for (var r = 1; r < values.length; r++) {
      if (String(values[r][iAlertId] || "") === alertId) { rowIdx = r; break; }
    }
    if (rowIdx < 0) return { ok: false, error: "Alerta não encontrado." };

    var row = values[rowIdx];
    var isActive = (row[iActive] === true);
    if (!isActive && !preview) return { ok: true, skipped: true, reason: "inativo" };

    var ownerEmail = String(row[iOwner] || "").trim();
    var ownerRole  = (iOwnerRole >= 0) ? String(row[iOwnerRole] || "").trim() : "";

    // ✅ fallback se não houver coluna ownerRole (opcional)
    // Se você NÃO tiver essa função, pode remover este bloco.
    if (!ownerRole && typeof vektorGetUserRoleByEmail_ === "function" && ownerEmail) {
      try {
        var rr = vektorGetUserRoleByEmail_(ownerEmail); // deve retornar algo como { role: "...", email:"..." }
        ownerRole = rr && rr.role ? String(rr.role).trim() : "";
      } catch(_) {}
    }

    var freq = String(row[iFreq] || "DAILY").trim();
    var windowDays = Number(row[iWin] || 30) || 30;

    var alertType = (iType >= 0 ? String(row[iType] || "TRANSACOES").trim() : "TRANSACOES");
    if (alertType !== "TRANSACOES" && alertType !== "PENDENCIAS") alertType = "TRANSACOES";

    // ATENÇÃO: "time" aqui é o Time/Grupo (não é horário)
    var time = String(row[iTime] || "").trim();

    var lojas = String(row[iLojas] || "")
      .split(",").map(function (s) { return String(s || "").trim(); }).filter(Boolean);

    // Etiquetas: "" => todas | "A | B | C" => múltiplas
    var etiquetaCsv = String(row[iEtq] || "").trim();
    var etiquetas = etiquetaCsv
      ? etiquetaCsv.split("|").map(function (s) { return String(s || "").trim(); }).filter(Boolean)
      : [];

    // ✅ ACL por Emails APENAS para Gerentes_Reg
    // Interseção: lojas do alerta X lojas permitidas pelo Emails
    if (ownerEmail && String(ownerRole) === "Gerentes_Reg") {
      var allowedLojasOwner = vektorGetAllowedLojasFromEmails_(String(ownerEmail).trim().toLowerCase()); // array ou null

      if (Array.isArray(allowedLojasOwner)) {
        // normaliza allowed para comparação
        var allowedSet = {};
        allowedLojasOwner.forEach(function(x){
          x = String(x || "").trim();
          if (!x) return;
          allowedSet[x] = true;
          allowedSet[x.padStart(4, "0")] = true;
        });

        // Se alerta estava como "todas" (lojasCsv vazio) => vira "todas permitidas"
        if (!lojas || !lojas.length) {
          lojas = Object.keys(allowedSet).filter(function(k){ return /^\d{4}$/.test(k); });
        } else {
          lojas = lojas.filter(function(x){
            x = String(x || "").trim();
            if (!x) return false;
            var x4 = x.padStart(4, "0");
            return !!allowedSet[x] || !!allowedSet[x4];
          });
        }

        if (!lojas || !lojas.length) {
          // ✅ não dispara nada se não sobrou loja permitida
          return { ok: true, alertId: alertId, skipped: true, rows: 0, reason: "Nenhuma loja permitida pelo Emails para este usuário." };
        }
      }
    }

    // Período: últimos N dias
    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    var now = new Date();
    var ini = new Date(now.getTime() - (windowDays * 24 * 60 * 60 * 1000));

    var periodo = {
      inicio: Utilities.formatDate(ini, tz, "dd/MM/yyyy"),
      fim: Utilities.formatDate(now, tz, "dd/MM/yyyy")
    };

    // Busca dados (Transações OU Pendências)
    var rows = [];

    if (alertType === "PENDENCIAS") {
      // Pendências não usam etiqueta
      rows = queryPendenciasBaseClaraAlert_(ini, now, time, lojas);
    } else {
      // Transações por etiqueta
      if (!etiquetas.length) {
        rows = queryTransacoesBaseClaraPorEtiqueta_(ini, now, time, lojas, "");
      } else {
        var acc = [];
        etiquetas.forEach(function (et) {
          acc = acc.concat(queryTransacoesBaseClaraPorEtiqueta_(ini, now, time, lojas, et));
        });
        rows = acc;
      }
    }

    // ✅ preview NÃO pode marcar execução
    if (!preview) {
      if (iLastRun >= 0) sh.getRange(rowIdx + 1, iLastRun + 1).setValue(now);
      if (iLastCnt >= 0) sh.getRange(rowIdx + 1, iLastCnt + 1).setValue(rows.length);
    }

    // Log de execução (guarda preview de no máx. 60 linhas)
    var runsSh = getOrCreateUserAlertsRunsSheet_();
    var previewRows = rows.slice(0, 60);

    runsSh.appendRow([
      "RUN" + Utilities.getUuid().replace(/-/g, "").slice(0, 10).toUpperCase(),
      alertId,
      ownerEmail,
      now,
      periodo.inicio,
      periodo.fim,
      rows.length,
      JSON.stringify(previewRows)
    ]);

    // ✅ Envio de e-mail SOMENTE quando NÃO for preview
    if (!preview) {
      try {
        var assunto = (String(alertType || "TRANSACOES") === "PENDENCIAS")
          ? "[Vektor - Grupo SBF] Alerta de Pendencias Clara — " + periodo.inicio + " a " + periodo.fim
          : "[Vektor - Grupo SBF] Alerta de Transações Clara — " + periodo.inicio + " a " + periodo.fim;

        var htmlBody = (String(alertType || "TRANSACOES") === "PENDENCIAS")
          ? montarEmailUserAlertPendencias_(alertId, time, periodo, rows)
          : montarEmailUserAlert_(alertId, time, etiquetaCsv, periodo, rows);

        var attachment = buildAlertAttachmentSmart_(alertId, periodo, rows);

        if (ownerEmail) {
          GmailApp.sendEmail(ownerEmail, assunto, " ", {
            from: "vektor@gruposbf.com.br",
            name: "Vektor - Grupo SBF",
            htmlBody: htmlBody,
            cc: "contasareceber@gruposbf.com.br",
            attachments: [attachment]
          });
        } else {
          Logger.log("Alerta " + alertId + " sem ownerEmail: não enviou e-mail.");
        }
      } catch (mailErr) {
        Logger.log(
          "Falha ao enviar e-mail do alerta " + alertId + ": " +
          (mailErr && mailErr.message ? mailErr.message : mailErr)
        );
      }
    }

    return {
      ok: true,
      alertId: alertId,
      ownerEmail: ownerEmail,
      ownerRole: ownerRole,
      freq: freq,
      time: time,
      alertType: alertType,
      etiqueta: (alertType === "PENDENCIAS" ? "" : (etiquetaCsv || "")),
      etiquetaCount: (alertType === "PENDENCIAS" ? 0 : etiquetas.length),
      periodo: periodo,
      rows: previewRows
    };

  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

function queryPendenciasBaseClaraAlert_(ini, fim, timeFiltro, lojasFiltro) {
  var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
  var sh = ss.getSheetByName("BaseClara");
  if (!sh) throw new Error("Aba BaseClara não encontrada.");

  var lr = sh.getLastRow();
  if (lr < 2) return [];

  // Lê A..W (23 colunas)
  var values = sh.getRange(2, 1, lr - 1, 23).getValues();

  // Índices zero-based (A..W)
  var IDX_DATA  = 0;   // A
  var IDX_ESTAB = 2;   // C  ✅ você pediu explicitamente
  var IDX_VALOR = 5;   // F
  var IDX_RECIBO = 14; // O
  var IDX_TITULAR = 16;// Q
  var IDX_GRUPO = 17;  // R (time/grupo)
  var IDX_ETIQUETA = 19; // T
  var IDX_DESC = 20;     // U
  var IDX_LOJA_NUM = 21; // V

  // Normaliza filtros
  var timeSel = String(timeFiltro || "").trim();
    var lojasSet = {};
  (lojasFiltro || []).forEach(function(l){
    // normaliza "CE0147" / "0147" / 147 => "147"
    var n = normalizarLojaNumero_(l);
    if (n !== null) lojasSet[String(n)] = true;
  });

    function isVazio_(v) {
    if (v === null || v === undefined) return true;

    // IMPORTANTÍSSIMO: no Google Sheets, checkbox pode virar boolean
    if (v === false) return true;

    var s = String(v).trim().toLowerCase();

    // placeholders comuns
    if (!s) return true;
    if (s === "-" || s === "—" || s === "n/a" || s === "na") return true;

    // casos que na BaseClara significam "sem recibo" (pendência)
    if (s === "não" || s === "nao") return true;
    if (s === "false" || s === "0") return true;

    // textos livres comuns
    if (s.indexOf("sem recibo") >= 0) return true;
    if (s.indexOf("sem nota") >= 0) return true;

    return false;
  }

    function isBlank_(v) {
    return String(v == null ? "" : v).trim() === "";
  }

  function isReciboPendente_(v) {
    // regra: pendência de NF quando coluna O for "Não"
    var s = String(v == null ? "" : v).trim().toLowerCase();
    return (s === "não" || s === "nao");
  }

  // Converte para datas comparáveis
    // ✅ evita drift quando ini/fim chegam como "YYYY-MM-DD"
  var iniD = (ini instanceof Date) ? ini : (vektorParseDateAny_(ini) || new Date(ini));
  var fimD = (fim instanceof Date) ? fim : (vektorParseDateAny_(fim) || new Date(fim));

  if (!(iniD instanceof Date) || isNaN(iniD.getTime()) || !(fimD instanceof Date) || isNaN(fimD.getTime())) {
    return [];
  }

  var iniMs = iniD.getTime();
  var fimMs = fimD.getTime();

  var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";

  var out = [];

  for (var i = 0; i < values.length; i++) {
    var row = values[i];

    var dt = row[IDX_DATA];
    var dt2 = (dt instanceof Date) ? dt : (vektorParseDateAny_(dt) || new Date(dt));
      if (!(dt2 instanceof Date) || isNaN(dt2.getTime())) continue;

    var tms = dt2.getTime();
    if (tms < iniMs || tms > fimMs) continue;

        // Loja (BaseClara coluna V = LojaNum)
    var lojaNum = normalizarLojaNumero_(row[IDX_LOJA_NUM]);
    if (lojaNum === null) continue;

    // para filtrar, usa chave numérica (sem "CE" e sem zeros à esquerda)
    var lojaKey = String(lojaNum);

    // para exibir, você pode padronizar como "CE####"
    var loja = "CE" + String(lojaNum).padStart(4, "0");

    // Filtro lojas (quando time != __ALL__)
    if (timeSel !== "__ALL__" && Object.keys(lojasSet).length) {
      if (!lojasSet[lojaKey]) continue;
    }

    // Time
    var grp = String(row[IDX_GRUPO] || "").trim();

    // Filtro por time
    if (timeSel && timeSel !== "__ALL__") {
      if (grp !== timeSel) continue;
    }

    var estab = String(row[IDX_ESTAB] || "").trim();
    var titular = String(row[IDX_TITULAR] || "").trim();
    var valor = Number(row[IDX_VALOR]) || 0;

    var etiquetas = String(row[IDX_ETIQUETA] || "").trim();
    var recibo = String(row[IDX_RECIBO] || "").trim();
    var desc = String(row[IDX_DESC] || "").trim();

    var pendEtiqueta = isBlank_(etiquetas);        // T vazia
    var pendNF       = isReciboPendente_(recibo); // O = "Não"
    var pendDesc     = isBlank_(desc);            // U vazia

    if (!pendEtiqueta && !pendNF && !pendDesc) continue;

    var pendList = [];
    if (pendEtiqueta) pendList.push("Etiqueta");
    if (pendDesc) pendList.push("Descrição");
    if (pendNF) pendList.push("Nota fiscal/Recibo");

    out.push({
      loja: loja,
      time: grp || "—",
      data: Utilities.formatDate(dt2, tz, "dd/MM/yyyy"),
      valor: valor,
      estabelecimento: estab,
      titular: titular,
      pendencias: pendList.join(", ")
    });
  }

  return out;
}

// =====================================================
// VALORES CONTABILIZADOS (BaseClara) - Back-end
// =====================================================
function getValoresContabilizadosEtiquetas(req) {
  vektorAssertFunctionAllowed_("getValoresContabilizadosEtiquetas");

  var ctxAcl = vektorGetUserRole_(); // {email, role}
  var emailAcl = String((ctxAcl && ctxAcl.email) || "").trim().toLowerCase();
  var roleAcl  = String((ctxAcl && ctxAcl.role)  || "").trim();

  var allowedLojas = null;
  if (roleAcl === "Gerentes_Reg") {
    allowedLojas = vektorGetAllowedLojasFromEmails_(emailAcl); // array ou null
  }

  try {
    req = req || {};

    var timeSel = String(req.time || "").trim();     // "" = todos
    var lojaSel = String(req.loja || "").trim();     // "" = todas
    var contaSel = String(req.conta || "").trim();   // "" = todas
    var iniIso  = String(req.dataInicioIso || "").trim();
    var fimIso  = String(req.dataFimIso || "").trim();

    // ✅ NORMALIZA "__ALL__" (front manda isso)
    if (timeSel === "__ALL__") timeSel = "";
    if (lojaSel === "__ALL__") lojaSel = "";
    if (contaSel === "__ALL__") contaSel = "";

    var ini = iniIso ? vektorParseIsoDateSafe_(iniIso) : null;
    var fim = fimIso ? vektorParseIsoDateSafe_(fimIso) : null;

    // período inclusivo (fim 23:59:59)
    if (ini) ini = new Date(ini.getFullYear(), ini.getMonth(), ini.getDate(), 0, 0, 0);
    if (fim) fim = new Date(fim.getFullYear(), fim.getMonth(), fim.getDate(), 23, 59, 59);

    // mapa loja->time
    var mapLojaTime = {};
    try { mapLojaTime = construirMapaLojaParaTime_() || {}; } catch (_) { mapLojaTime = {}; }

    var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
    var sh = ss.getSheetByName("BaseClara");
    if (!sh) throw new Error("Aba BaseClara não encontrada.");

    var lr = sh.getLastRow();
    if (lr < 2) return { ok:true, total:0, rows:[], categorias:[] };

    // A..W (23 colunas)
    var values = sh.getRange(2, 1, lr - 1, 23).getValues();

    var IDX_DATA      = 0;   // A
    var IDX_VALOR     = 5;   // F
    var IDX_GRUPO     = 17;  // R
    var IDX_ETIQUETA  = 19;  // T
    var IDX_LOJA_NUM  = 21;  // V

    function normLoja4_(x) {
      var s = String(x || "").trim();
      if (!s) return "";
      var m = s.match(/(\d{1,4})/);
      if (!m) return "";
      return String(Number(m[1])).padStart(4, "0");
    }

    // ✅ normaliza etiqueta para evitar duplicadas “iguais” (NBSP, hífens, espaços)
    function normEtq_(s){
      s = String(s || "");
      s = s.replace(/\u00A0/g, " ");      // NBSP -> espaço normal
      s = s.replace(/[–—]/g, "-");        // hífens “diferentes” -> "-"
      s = s.replace(/\s+/g, " ").trim();  // colapsa espaços
      return s;
    }

    function parseConta_(etq) {
      etq = normEtq_(etq);
      if (!etq) return { conta:"", etiqueta:"" };

      var m = etq.match(/^(\d{2,})\s*[-–]?\s*(.*)$/);
      if (m) {
        var num = String(m[1] || "").trim();
        return { conta: num, etiqueta: etq };
      }
      return { conta: etq, etiqueta: etq };
    }

    // categorias fixas
    var CATS = [
      { cat:"🏛️ Administrativo e Geral", keys:[
        "MATERIAL DE ESCRITÓRIO",
        "TAXAS E EMOLUMENTOS",
        "CORREIOS_SEDEX/AR/POSTAGEM",
        "SERVIÇOS GRÁFICOS E DE COPIADORAS"
      ]},
      { cat:"💰 Financeiro e Operações de Venda", keys:[
        "BOBINA ECF",
        "TRANSPORTE SERVICOS EMERGENCIAS"
      ]},
      { cat:"📢 Comercial e Marketing", keys:[
        "MARKETING_PUBLICIDADE E PROPAGANDA",
        "SERVIÇOS GRÁFICOS OPERAÇÕES"
      ]},
      { cat:"🛠️ Manutenção e Conservação", keys:[
        "MANUTENÇÃO CIVIL",
        "MANUTENÇÃO ELETRICO",
        "MANUTENÇÃO AR-CONDICIONADO",
        "MANUTENÇÃO EQUIPAMENTOS",
        "MANUTENÇÃO MAQ ESTAMPAR",
        "MATERIAL DE INFORMÁTICA",
        "CHAVEIRO EMERGENCIAL"
      ]},
      { cat:"🧼 Limpeza e Higiene", keys:[
        "MATERIAL DE LIMPEZA",
        "MATERIAL DE LIMPEZA OPERAÇÕES",
        "SERVIÇOS DE LIMPEZA"
      ]},
      { cat:"☕ Copa, Cozinha e Bem-estar", keys:[
        "MATERIAL DE COPA E COZINHA",
        "MATERIAL DE COPA E COZINHA OPERAÇÕES",
        "AGUA POTÁVEL",
        "LANCHES DE REFEIÇÕES",
        "ACAO BF"
      ]}
    ];

    function categoriaDaEtiqueta_(etq) {
      var s = normEtq_(etq).toUpperCase();

      // remove acentos (NFD)
      try {
        s = s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");
      } catch (_) {}

      for (var i=0; i<CATS.length; i++) {
        for (var j=0; j<CATS[i].keys.length; j++) {
          var k = String(CATS[i].keys[j] || "");
          k = normEtq_(k).toUpperCase();
          try { k = k.normalize("NFD").replace(/[\u0300-\u036f]/g, ""); } catch (_) {}
          if (k && s.indexOf(k) >= 0) return CATS[i].cat;
        }
      }
      return "Outros";
    }

    var sumByEtq = {}; // { etiquetaKey: { etiqueta, conta, categoria, valor } }
    var sumByCat = {}; // { categoria: valor }
    var total = 0;

    for (var i=0; i<values.length; i++) {
      var row = values[i];

      var dt = row[IDX_DATA];
      if (!(dt instanceof Date) || isNaN(dt.getTime())) continue;

      if (ini && dt < ini) continue;
      if (fim && dt > fim) continue;

      var loja4 = normLoja4_(row[IDX_LOJA_NUM]);
      var lojaNum = String(Number(loja4) || "").trim(); // ✅ FIX (evita "lojaNum is not defined")

      if (lojaSel) {
        var lojaSel4 = String(lojaSel).padStart(4,"0");
        if (loja4 !== lojaSel4) continue;
      }

      // ✅ ACL por email (somente Gerentes_Reg)
      if (Array.isArray(allowedLojas)) {
        if (allowedLojas.indexOf(lojaNum) < 0 && allowedLojas.indexOf(loja4) < 0) continue;
      }

      var timeRow = String(row[IDX_GRUPO] || "").trim();
      var timeFinal = timeRow || (mapLojaTime[loja4] ? String(mapLojaTime[loja4]).trim() : "N/D");
      if (timeSel && timeFinal !== timeSel) continue;

      var etqRaw = normEtq_(row[IDX_ETIQUETA]);
      if (!etqRaw) continue;

      var valor = Number(row[IDX_VALOR] || 0) || 0;
      if (!valor) continue;

      // ✅ total real do período (não duplicar por multi-etiqueta)
      total += valor;

      // ✅ quebra múltiplas etiquetas (com ou sem espaços ao redor do "|")
      var parts = etqRaw.split(/\s*\|\s*/).map(normEtq_).filter(function(s){ return !!s; });
      if (!parts.length) parts = [etqRaw];

      // ✅ divide para o total/% fechar 100%
      var valorPorEtiqueta = valor / parts.length;

      // ✅ processa cada etiqueta separadamente
      for (var p = 0; p < parts.length; p++) {
        var etqPart = parts[p];
        if (!etqPart) continue;

        var contaObj = parseConta_(etqPart); // {conta, etiqueta}
        var conta = contaObj.conta || "";
        var etiquetaFinal = normEtq_(contaObj.etiqueta || etqPart);

        // filtro conta (se houver)
        if (contaSel && conta !== contaSel) continue;

        // ✅ chave ÚNICA: somente pela etiqueta normalizada (1 linha por etiqueta)
        var etqKey = etiquetaFinal;

        var catPart = categoriaDaEtiqueta_(etqPart);

        if (!sumByEtq[etqKey]) {
          sumByEtq[etqKey] = {
            etiqueta: etiquetaFinal,
            conta: conta,
            categoria: catPart,
            valor: 0
          };
        }

        sumByEtq[etqKey].valor += valorPorEtiqueta;

        if (!sumByCat[catPart]) sumByCat[catPart] = 0;
        sumByCat[catPart] += valorPorEtiqueta;
      }
    }

    var rowsOut = Object.keys(sumByEtq).map(function(k){
      var r = sumByEtq[k];
      return {
        etiqueta: r.etiqueta,
        conta: r.conta,
        categoria: r.categoria,
        valor: r.valor,
        pct: (total > 0 ? (r.valor / total) : 0)
      };
    });

    rowsOut.sort(function(a,b){ return (b.valor||0) - (a.valor||0); });

    var catsOut = Object.keys(sumByCat).map(function(k){
      var v = sumByCat[k];
      return { categoria: k, valor: v, pct: (total > 0 ? (v / total) : 0) };
    });
    catsOut.sort(function(a,b){ return (b.valor||0) - (a.valor||0); });

    return { ok:true, total: total, rows: rowsOut, categorias: catsOut };

  } catch (e) {
    return { ok:false, error: (e && e.message) ? e.message : String(e) };
  }
}

// =====================================================
// VALORES CONTABILIZADOS (BaseClara) - Série 12 meses
// =====================================================
function getValoresContabilizadosSerie12m(req) {
  vektorAssertFunctionAllowed_("getValoresContabilizadosSerie12m");

  var ctxAcl = vektorGetUserRole_(); // {email, role}
  var emailAcl = String((ctxAcl && ctxAcl.email) || "").trim().toLowerCase();
  var roleAcl  = String((ctxAcl && ctxAcl.role)  || "").trim();

  var allowedLojas = null;
  if (roleAcl === "Gerentes_Reg") {
    allowedLojas = vektorGetAllowedLojasFromEmails_(emailAcl); // array ou null
  }

  try {
    req = req || {};

    var timeSel  = String(req.time || "").trim();
    var lojaSel  = String(req.loja || "").trim();
    var contaSel = String(req.conta || "").trim();
    var catSel   = String(req.categoria || "").trim();

    if (timeSel === "__ALL__") timeSel = "";
    if (lojaSel === "__ALL__") lojaSel = "";
    if (contaSel === "__ALL__") contaSel = "";
    if (catSel === "__ALL__") catSel = "";

    // ✅ SEMPRE: últimos 12 meses a partir de hoje
    var now = new Date();
    var endMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    var startMonth = new Date(endMonth.getFullYear(), endMonth.getMonth() - 11, 1);

    function mmYYYY_(d) {
      var mm = String(d.getMonth() + 1).padStart(2, "0");
      return mm + "/" + String(d.getFullYear());
    }
    function ymKey_(d) {
      var mm = String(d.getMonth() + 1).padStart(2, "0");
      return String(d.getFullYear()) + "-" + mm;
    }

    var labels = [];
    var monthKeys = [];
    var cur = new Date(startMonth.getFullYear(), startMonth.getMonth(), 1);
    var guard = 0;
    while (cur <= endMonth && guard < 36) {
      labels.push(mmYYYY_(cur));
      monthKeys.push(ymKey_(cur));
      cur = new Date(cur.getFullYear(), cur.getMonth() + 1, 1);
      guard++;
    }

    // mapa loja->time
    var mapLojaTime = {};
    try { mapLojaTime = construirMapaLojaParaTime_() || {}; } catch (_) { mapLojaTime = {}; }

    var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
    var sh = ss.getSheetByName("BaseClara");
    if (!sh) throw new Error("Aba BaseClara não encontrada.");

    var lr = sh.getLastRow();
    if (lr < 2) {
      return { ok:true, labels: labels, totais: labels.map(function(){return 0;}), variacoesPct: labels.map(function(){return 0;}) };
    }

    // A..W (23 colunas)
    var values = sh.getRange(2, 1, lr - 1, 23).getValues();

    var IDX_DATA      = 0;   // A
    var IDX_VALOR     = 5;   // F
    var IDX_ETIQUETA  = 19;  // T
    var IDX_LOJA_NUM  = 21;  // V

    function normLoja4_(x) {
      var s = String(x || "").trim();
      if (!s) return "";
      var m = s.match(/(\d{1,4})/);
      if (!m) return "";
      return String(Number(m[1])).padStart(4, "0");
    }

    function normEtq_(s){
      s = String(s || "");
      s = s.replace(/\u00A0/g, " ");
      s = s.replace(/[–—]/g, "-");
      s = s.replace(/\s+/g, " ").trim();
      return s;
    }

    function parseConta_(etq) {
      etq = normEtq_(etq);
      if (!etq) return { conta:"", etiqueta:"" };
      var m = etq.match(/^(\d{2,})\s*[-–]?\s*(.*)$/);
      if (m) {
        var num = String(m[1] || "").trim();
        return { conta: num, etiqueta: etq };
      }
      return { conta: etq, etiqueta: etq };
    }

    // ✅ mesma categorização usada na tabela de valores contabilizados
    var CATS = [
      { cat:"🏛️ Administrativo e Geral", keys:[
        "MATERIAL DE ESCRITÓRIO",
        "TAXAS E EMOLUMENTOS",
        "CORREIOS_SEDEX/AR/POSTAGEM",
        "SERVIÇOS GRÁFICOS E DE COPIADORAS"
      ]},
      { cat:"💰 Financeiro e Operações de Venda", keys:[
        "BOBINA ECF",
        "TRANSPORTE SERVICOS EMERGENCIAS"
      ]},
      { cat:"📢 Comercial e Marketing", keys:[
        "MARKETING_PUBLICIDADE E PROPAGANDA",
        "SERVIÇOS GRÁFICOS OPERAÇÕES"
      ]},
      { cat:"🛠️ Manutenção e Conservação", keys:[
        "MANUTENÇÃO CIVIL",
        "MANUTENÇÃO ELETRICO",
        "MANUTENÇÃO AR-CONDICIONADO",
        "MANUTENÇÃO EQUIPAMENTOS",
        "MANUTENÇÃO MAQ ESTAMPAR",
        "MATERIAL DE INFORMÁTICA",
        "CHAVEIRO EMERGENCIAL"
      ]},
      { cat:"🧼 Limpeza e Higiene", keys:[
        "MATERIAL DE LIMPEZA",
        "MATERIAL DE LIMPEZA OPERAÇÕES",
        "SERVIÇOS DE LIMPEZA"
      ]},
      { cat:"☕ Copa, Cozinha e Bem-estar", keys:[
        "MATERIAL DE COPA E COZINHA",
        "MATERIAL DE COPA E COZINHA OPERAÇÕES",
        "AGUA POTÁVEL",
        "LANCHES DE REFEIÇÕES",
        "ACAO BF"
      ]}
    ];

    function categoriaDaEtiqueta_(etq) {
      var s = normEtq_(etq).toUpperCase();
      try { s = s.normalize("NFD").replace(/[\u0300-\u036f]/g, ""); } catch (_) {}

      for (var i=0; i<CATS.length; i++) {
        for (var j=0; j<CATS[i].keys.length; j++) {
          var k = String(CATS[i].keys[j] || "");
          k = normEtq_(k).toUpperCase();
          try { k = k.normalize("NFD").replace(/[\u0300-\u036f]/g, ""); } catch (_) {}
          if (k && s.indexOf(k) >= 0) return CATS[i].cat;
        }
      }
      return "Outros";
    }

    // acumula por mês
    var sumByMonth = {};
    monthKeys.forEach(function(k){ sumByMonth[k] = 0; });

    for (var i = 0; i < values.length; i++) {
      var row = values[i];

      // data
      var dt = row[IDX_DATA];
      dt = (dt instanceof Date) ? dt : new Date(dt);
      if (!(dt instanceof Date) || isNaN(dt.getTime())) continue;

      // ✅ recorta apenas pela janela 12m (startMonth..agora)
      if (dt < startMonth || dt > now) continue;

      // loja
      var loja4 = normLoja4_(row[IDX_LOJA_NUM]);
      var lojaNum = String(Number(loja4) || "").trim(); // ✅ FIX (evita "lojaNum is not defined")

      if (lojaSel) {
        var lojaSel4 = normLoja4_(lojaSel);
        if (lojaSel4 && loja4 !== lojaSel4) continue;
      }

      // ✅ ACL por email (somente Gerentes_Reg)
      if (Array.isArray(allowedLojas)) {
        if (allowedLojas.indexOf(lojaNum) < 0 && allowedLojas.indexOf(loja4) < 0) continue;
      }

      // time (via map loja->time)
      var timeFinal = (loja4 && mapLojaTime[loja4] != null ? String(mapLojaTime[loja4]).trim() : "N/D");
      if (timeSel && timeFinal !== timeSel) continue;

      // etiqueta
      var etqRaw = normEtq_(row[IDX_ETIQUETA]);
      if (!etqRaw) continue;

      // valor
      var valor = Number(row[IDX_VALOR] || 0) || 0;
      if (!valor) continue;

      // quebra múltiplas etiquetas e aloca
      var parts = etqRaw.split(/\s*\|\s*/).map(normEtq_).filter(function(s){ return !!s; });
      if (!parts.length) parts = [etqRaw];

      var valorPorEtiqueta = valor / parts.length;

      // mês
      var ym = String(dt.getFullYear()) + "-" + String(dt.getMonth() + 1).padStart(2, "0");
      if (sumByMonth[ym] == null) continue;

      for (var p = 0; p < parts.length; p++) {
        var etqPart = parts[p];
        if (!etqPart) continue;

        // filtro conta
        var contaObj = parseConta_(etqPart);
        var conta = contaObj.conta || "";
        if (contaSel && conta !== contaSel) continue;

        // filtro categoria (cluster)
        var cat = categoriaDaEtiqueta_(etqPart) || "";
        if (catSel && cat !== catSel) continue;

        sumByMonth[ym] += valorPorEtiqueta;
      }
    }

    var totais = monthKeys.map(function(k){ return Number(sumByMonth[k] || 0) || 0; });

    // ✅ variação mensal (MoM) em %
    var variacoesPct = [];
    for (var j = 0; j < totais.length; j++) {
      if (j === 0) {
        variacoesPct.push(0);
      } else {
        var prev = Number(totais[j - 1] || 0) || 0;
        var curV = Number(totais[j] || 0) || 0;
        if (!prev) variacoesPct.push(0);
        else variacoesPct.push(((curV - prev) / prev) * 100);
      }
    }

    return { ok:true, labels: labels, totais: totais, variacoesPct: variacoesPct };

  } catch (e) {
    return { ok:false, error: (e && e.message) ? e.message : String(e) };
  }
}

function montarEmailUserAlert_(alertId, time, etiquetaCsv, periodo, rows) {
  var esc = function (s) {
    return String(s == null ? "" : s)
      .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
      .replace(/"/g,"&quot;").replace(/'/g,"&#39;");
  };

  var etiquetaTxt = etiquetaCsv ? esc(etiquetaCsv) : "Todas";
  var max = 120; // evita e-mails gigantes
  var view = (rows || []).slice(0, max);

  // Soma do valor total no período
  var totalValor = 0;
  (rows || []).forEach(function (r) {
    // remove R$, pontos e troca vírgula por ponto
    var v = String(r.valor || "")
      .replace(/[R$\s]/g, "")
      .replace(/\./g, "")
      .replace(",", ".");
    var n = Number(v);
    if (!isNaN(n)) totalValor += n;
  });

  // Formatação BRL
  var totalValorFmt = totalValor.toLocaleString("pt-BR", {
    style: "currency",
    currency: "BRL"
  });

  var h = "";
  h += "<div style='font-family:Arial,sans-serif;font-size:13px;color:#0f172a'>";
  h += "<h2 style='margin:0 0 8px 0'>Plataforma de Governança Financeira do Cartão Clara</h2>";
  h += "<div style='margin:0 0 10px 0'>";
  h += "<b>ID:</b> " + esc(alertId) + "<br/>";
  h += "<b>Time:</b> " + esc(time) + "<br/>";
  h += "<b>Etiqueta:</b> " + etiquetaTxt + "<br/>";
  h += "<b>Período:</b> " + esc(periodo.inicio) + " a " + esc(periodo.fim) + "<br/>";
  h += "<b>Total de linhas:</b> " + esc((rows || []).length) + (rows.length > max ? " (mostrando " + max + ")" : "") + "<br/>";
  h += "<b>Valor total no período analisado:</b> " + esc(totalValorFmt);
  h += "</div>";

  if (!view.length) {
    h += "<div style='padding:10px;border:1px solid #e2e8f0;border-radius:10px;background:#f8fafc'>";
    h += "Nenhuma transação encontrada para este alerta no período.";
    h += "</div></div>";
    return h;
  }

  h += "<div style='overflow:auto;border:1px solid #e2e8f0;border-radius:10px'>";
  h += "<table style='border-collapse:collapse;width:100%'>";
  h += "<thead><tr style='background:#0b1220;color:#fff'>";
  ["Loja","Time","Data","Estabelecimento","Valor","Etiqueta","Descrição"].forEach(function(c){
    h += "<th style='text-align:left;padding:8px;border:1px solid #111827;font-size:12px;white-space:nowrap'>" + c + "</th>";
  });
  h += "</tr></thead><tbody>";

  view.forEach(function(r){
    h += "<tr>";
    h += "<td style='padding:6px;border:1px solid #e2e8f0;white-space:nowrap'>" + esc(r.loja) + "</td>";
    h += "<td style='padding:6px;border:1px solid #e2e8f0;white-space:nowrap'>" + esc(r.time) + "</td>";
    h += "<td style='padding:6px;border:1px solid #e2e8f0;white-space:nowrap'>" + esc(r.data) + "</td>";
    h += "<td style='padding:6px;border:1px solid #e2e8f0'>" + esc(r.estabelecimento) + "</td>";
    h += "<td style='padding:6px;border:1px solid #e2e8f0;white-space:nowrap;text-align:right'>" + esc(r.valor) + "</td>";
    h += "<td style='padding:6px;border:1px solid #e2e8f0'>" + esc(r.etiqueta) + "</td>";
    h += "<td style='padding:6px;border:1px solid #e2e8f0'>" + esc(r.descricao) + "</td>";
    h += "</tr>";
  });

  h += "</tbody></table></div></div>";
  return h;
}

function montarEmailUserAlertPendencias_(alertId, time, periodo, rows) {
  var esc = function (s) {
    return String(s == null ? "" : s)
      .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
      .replace(/"/g,"&quot;").replace(/'/g,"&#39;");
  };

  var timeLabel = (String(time || "").trim() === "__ALL__") ? "Todos" : String(time || "").trim();

  var max = 120; // evita e-mail gigante
  var view = (rows || []).slice(0, max);

  // ---- Normaliza e conta tipos de pendência
  // Esperado em r.pendencias: "Etiqueta, Descrição, Nota fiscal/Recibo" (variações possíveis)
  var catCount = { "Etiqueta": 0, "Descrição": 0, "Nota fiscal/Recibo": 0, "Outros": 0 };

  function normalizarPendTipo_(p) {
    var x = String(p || "").trim().toLowerCase();
    if (!x) return "";
    if (x.indexOf("etiqueta") >= 0) return "Etiqueta";
    if (x.indexOf("descr") >= 0) return "Descrição";
    if (x.indexOf("nota") >= 0 || x.indexOf("recibo") >= 0) return "Nota fiscal/Recibo";
    return "Outros";
  }

  var totalPendencias = 0;

    // ---- Maior loja ofensora (por quantidade de transações pendentes)
  var lojaTxCount = {};   // loja -> quantidade de linhas (transações) com alguma pendência
  var lojaPendCount = {}; // loja -> total de pendências (tipos) (usado apenas para estatísticas, se quiser)

  // ---- Soma do valor total (pendente) no período
  var totalValor = 0;

  (rows || []).forEach(function (r) {
    // soma valor (aceita número ou string "R$ 1.234,56")
    var vNum = 0;
    if (typeof r.valor === "number") {
      vNum = r.valor;
    } else {
      var v = String(r.valor || "")
        .replace(/[R$\s]/g, "")
        .replace(/\./g, "")
        .replace(",", ".");
      var n = Number(v);
      vNum = isNaN(n) ? 0 : n;
    }
    totalValor += vNum;

    // conta pendências
    var pendRaw = String(r.pendencias || "").trim();
    if (!pendRaw) return;

    // separa por vírgula
    var parts = pendRaw.split(",").map(function(s){ return String(s||"").trim(); }).filter(Boolean);

        var lojaKey = String(r.loja || "").trim() || "—";

    // conta transação pendente (1 por linha)
    if (!lojaTxCount[lojaKey]) lojaTxCount[lojaKey] = 0;
    lojaTxCount[lojaKey] += 1;

    // (opcional) conta tipos de pendência (para total e maior ofensor)
    if (!lojaPendCount[lojaKey]) lojaPendCount[lojaKey] = 0;

    parts.forEach(function (p) {
      var cat = normalizarPendTipo_(p) || "Outros";
      if (!catCount[cat]) catCount[cat] = 0;
      catCount[cat] += 1;
      totalPendencias += 1;

      // total de tipos por loja (não usado para "maior loja", mas pode manter)
      lojaPendCount[lojaKey] += 1;
    });
  });

    // maior ofensor (tipo) — só mostra se houver 2+ lojas e vencedor único (sem empate)
  var maiorOfensor = "";
  var maiorOfensorQtd = -1;
  var empateOfensor = false;

  var lojasDistintas = Object.keys(lojaTxCount).length;

  if (lojasDistintas >= 2) {
    Object.keys(catCount).forEach(function (k) {
      var v = catCount[k] || 0;
      if (v > maiorOfensorQtd) {
        maiorOfensorQtd = v;
        maiorOfensor = k;
        empateOfensor = false;
      } else if (v === maiorOfensorQtd && v > 0) {
        empateOfensor = true;
      }
    });

    if (maiorOfensorQtd <= 0 || empateOfensor) {
      maiorOfensor = "";
      maiorOfensorQtd = 0;
    }
  } else {
    maiorOfensor = "";
    maiorOfensorQtd = 0;
  }

    // maior loja ofensora (por transações pendentes) — só mostra se houver vencedor único
  var maiorLoja = "";
  var maiorLojaQtd = -1;
  var empateLoja = false;

  Object.keys(lojaTxCount).forEach(function (lk) {
    var v = lojaTxCount[lk] || 0;
    if (v > maiorLojaQtd) {
      maiorLojaQtd = v;
      maiorLoja = lk;
      empateLoja = false;
    } else if (v === maiorLojaQtd && v > 0) {
      empateLoja = true;
    }
  });

  // se não tem vencedor único (empate) ou não tem dados, fica vazio
  if (maiorLojaQtd <= 0 || empateLoja) {
    maiorLoja = "";
    maiorLojaQtd = 0;
  }

  // Formatação BRL
  var totalValorFmt = totalValor.toLocaleString("pt-BR", {
    style: "currency",
    currency: "BRL"
  });

  // ---- Corpo do e-mail (com o “template” que você pediu)
  var h = "";
  h += "<div style='font-family:Arial,sans-serif;font-size:13px;color:#0f172a'>";
  h += "<h2 style='margin:0 0 8px 0'>Plataforma de Governança Financeira do Cartão Clara</h2>";
  h += "<div style='margin:0 0 10px 0'>";
  h += "<b>ID:</b> " + esc(alertId) + "<br/>";
  h += "<b>Time:</b> " + esc(timeLabel) + "<br/>";
  h += "<b>Período:</b> " + esc(periodo.inicio) + " a " + esc(periodo.fim) + "<br/>";
  h += "<b>Total de linhas:</b> " + esc((rows || []).length) + (rows.length > max ? " (mostrando apenas " + max + ")" : "") + "<br/>";
  h += "<b>Quantidade Pendencias:</b> " + esc(totalPendencias) + "<br/>";
  h += "<b>Maior Ofensor:</b> " + (maiorOfensor ? (esc(maiorOfensor) + " (" + esc(maiorOfensorQtd) + ")") : "") + "<br/>";
  h += "<b>Maior Loja ofensora:</b> " + (maiorLoja ? (esc(maiorLoja) + " (" + esc(maiorLojaQtd) + ")") : "") + "<br/>";
  h += "<b>Valor total pendente no período analisado:</b> " + esc(totalValorFmt) + "<br/>";
  h += "</div>";

  // ---- Tabela (preview)
  if (!view.length) {
    h += "<p>Nenhuma pendência encontrada para o filtro configurado.</p>";
    h += "<p style='margin-top:14px'><b>Vektor - Grupo SBF</b></p>";
    h += "</div>";
    return h;
  }

  var th = "background:#0b1220;color:#fff;border:1px solid #111827;padding:8px;font-size:12px;white-space:nowrap;";
  var td = "border:1px solid #e2e8f0;padding:6px;font-size:12px;vertical-align:top;";

  h += "<div style='overflow:auto;border:1px solid #e2e8f0;border-radius:10px'>";
  h += "<table style='border-collapse:collapse;width:100%'>";
  h += "<thead><tr>";
  ["Loja","Time","Data","Valor","Estabelecimento","Titular","Pendências"].forEach(function(c){
    h += "<th style='" + th + "'>" + esc(c) + "</th>";
  });
  h += "</tr></thead><tbody>";

  view.forEach(function(r){
    h += "<tr>";
    h += "<td style='" + td + "white-space:nowrap'>" + esc(r.loja) + "</td>";
    h += "<td style='" + td + "white-space:nowrap'>" + esc(r.time) + "</td>";
    h += "<td style='" + td + "white-space:nowrap'>" + esc(r.data) + "</td>";
    h += "<td style='" + td + "white-space:nowrap;text-align:right'>" + esc(r.valor) + "</td>";
    h += "<td style='" + td + "'>" + esc(r.estabelecimento) + "</td>";
    h += "<td style='" + td + "'>" + esc(r.titular) + "</td>";
    h += "<td style='" + td + "'>" + esc(r.pendencias) + "</td>";
    h += "</tr>";
  });

  h += "</tbody></table></div>";
  h += "<p style='margin-top:14px'><b>Vektor - Grupo SBF</b></p>";
  h += "</div>";
  return h;
}

// Gatilho diário (você cria no editor de Apps Script como time-driven 1x por dia)
function RUN_USER_ALERTS_SCHEDULER() {
  var sh = getOrCreateUserAlertsSheet_();
  var values = sh.getDataRange().getValues();
  if (values.length < 2) return;

  var head = values[0].map(String);
  var idx = function(n){ return head.indexOf(n); };

  var iAlertId = idx("alertId");
  var iActive  = idx("isActive");
  var iFreq    = idx("freq");
  var iLastRun = idx("lastRunAt");
  var iSendAt  = idx("sendAt");

  var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
  var now = new Date();

  // ✅ horários permitidos
  var allowedTimes = { "11:30": true, "16:00": true };

  for (var r=1; r<values.length; r++) {
    var row = values[r];
    if (row[iActive] !== true) continue;

    var alertId = String(row[iAlertId] || "").trim();
    if (!alertId) continue;

    var freq = String(row[iFreq] || "DAILY").trim();
    var lastRun = (row[iLastRun] instanceof Date) ? row[iLastRun] : null;

    // ✅ sendAt obrigatório e restrito
    var sendAtRaw = (iSendAt >= 0) ? String(row[iSendAt] || "").trim() : "";
    if (!sendAtRaw) continue;
    if (!allowedTimes[sendAtRaw]) continue;

    if (!isDueBySchedule_(freq, lastRun, sendAtRaw, now, tz)) continue;

    executarAlertaEtiquetaVektor({ alertId: alertId, preview: false });
  }
}

// ==============================
// DIAS ÚTEIS / FERIADOS / EMENDAS
// ==============================

// Calendário público de feriados do Brasil (Google)
var VEKTOR_BR_HOLIDAY_CAL_ID = "pt.brazilian#holiday@group.v.calendar.google.com";

// Cache simples (evita bater no Calendar toda hora)
var __vektorHolidayCache = { year: null, map: null };

// Converte Date -> "yyyy-MM-dd" no TZ do script
function vektorDateKey_(d, tz) {
  return Utilities.formatDate(d, tz || Session.getScriptTimeZone() || "America/Sao_Paulo", "yyyy-MM-dd");
}

// Retorna um Set/mapa { "yyyy-MM-dd": true } com feriados do ano
function vektorLoadBrHolidaysMap_(year, tz) {
  try {
    if (__vektorHolidayCache.year === year && __vektorHolidayCache.map) {
      return __vektorHolidayCache.map;
    }

    var out = {};
    var cal = CalendarApp.getCalendarById(VEKTOR_BR_HOLIDAY_CAL_ID);

    // Range do ano inteiro
    var ini = new Date(year, 0, 1);  ini.setHours(0,0,0,0);
    var fim = new Date(year, 11, 31); fim.setHours(23,59,59,999);

    // eventos de dia inteiro (feriados)
    var events = cal.getEvents(ini, fim);
    events.forEach(function(ev) {
      var allDay = ev.isAllDayEvent && ev.isAllDayEvent();
      // Mesmo se não for all-day, ainda marca o dia do start
      var d0 = ev.getStartTime();
      d0.setHours(0,0,0,0);
      out[vektorDateKey_(d0, tz)] = true;
    });

    // ✅ Extras manuais (para exceções corporativas, se quiser):
    // Script Properties: VEKTOR_EXTRA_OFF_DAYS="2026-02-16,2026-02-17"
    try {
      var extraRaw = PropertiesService.getScriptProperties().getProperty("VEKTOR_EXTRA_OFF_DAYS") || "";
      extraRaw.split(",").map(function(s){ return String(s||"").trim(); }).filter(Boolean).forEach(function(k){
        out[k] = true;
      });
    } catch (_) {}

    __vektorHolidayCache.year = year;
    __vektorHolidayCache.map = out;
    return out;

  } catch (e) {
    // Fallback mínimo: sem calendário, considera só fim de semana
    return {};
  }
}

// Emenda nacional (regra simples e explícita):
// - Se feriado cai na TERÇA => SEGUNDA vira off
// - Se feriado cai na QUINTA => SEXTA vira off
function vektorIsBridgeDay_(dateObj, holidaysMap, tz) {
  var d = new Date(dateObj); d.setHours(0,0,0,0);

  // se amanhã é feriado e amanhã é terça => hoje (segunda) é emenda
  var tomorrow = new Date(d); tomorrow.setDate(d.getDate() + 1);
  var keyTomorrow = vektorDateKey_(tomorrow, tz);
  if (holidaysMap[keyTomorrow] && tomorrow.getDay() === 2) return true; // 2 = terça

  // se ontem foi feriado e ontem foi quinta => hoje (sexta) é emenda
  var yesterday = new Date(d); yesterday.setDate(d.getDate() - 1);
  var keyYesterday = vektorDateKey_(yesterday, tz);
  if (holidaysMap[keyYesterday] && yesterday.getDay() === 4) return true; // 4 = quinta

  return false;
}

// Dia útil = não sábado/domingo, não feriado e não emenda (regra acima)
function vektorIsBusinessDay_(dateObj, tz) {
  var d = new Date(dateObj);
  d.setHours(0,0,0,0);

  var day = d.getDay();
  if (day === 0 || day === 6) return false; // domingo/sábado

  var year = d.getFullYear();
  var holidaysMap = vektorLoadBrHolidaysMap_(year, tz);
  var key = vektorDateKey_(d, tz);

  if (holidaysMap[key]) return false;
  if (vektorIsBridgeDay_(d, holidaysMap, tz)) return false;

  return true;
}

function parseSendAt_(v, tz) {
  // Retorna {hh, mm} ou null
  if (v == null || v === "") return null;

  // Caso 1: string "HH:mm"
  if (typeof v === "string") {
    var s = v.trim();
    var m = s.match(/^(\d{1,2}):(\d{2})$/);
    if (m) {
      var hh = Number(m[1]), mm = Number(m[2]);
      if (hh >= 0 && hh <= 23 && mm >= 0 && mm <= 59) return { hh: hh, mm: mm };
    }
    return null;
  }

  // Caso 2: Date (Sheets pode guardar horário como Date)
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
    // Extrai hora/minuto no timezone do script
    var hh2 = Number(Utilities.formatDate(v, tz, "H"));
    var mm2 = Number(Utilities.formatDate(v, tz, "m"));
    return { hh: hh2, mm: mm2 };
  }

  // Caso 3: número (fração do dia): 0.583333 = 14:00
  if (typeof v === "number" && isFinite(v)) {
    var totalMinutes = Math.round(v * 24 * 60);
    var hh3 = Math.floor(totalMinutes / 60) % 24;
    var mm3 = totalMinutes % 60;
    return { hh: hh3, mm: mm3 };
  }

  return null;
}

function isDueBySchedule_(freq, lastRun, sendAtRaw, now, tz) {
  // sendAtRaw pode ser string, number, Date
  var t = parseSendAt_(sendAtRaw, tz);

  // Se tem horário, só permite disparo depois daquele horário no "dia atual"
  if (t) {
    var todayAt = new Date(now);
    todayAt.setHours(t.hh, t.mm, 0, 0);

    // Compara em "relógio local" do script
    if (now.getTime() < todayAt.getTime()) return false;
  }

  // ✅ NOVO: só permite disparo em dia útil
  if (!vektorIsBusinessDay_(now, tz)) return false;

  // Nunca rodou → pode rodar (desde que passou do horário, se houver)
  if (!(lastRun instanceof Date) || isNaN(lastRun.getTime())) return true;

  var diffMs = now.getTime() - lastRun.getTime();
  var diffDays = Math.floor(diffMs / (24 * 60 * 60 * 1000));

  if (freq === "DAILY") return diffDays >= 1;
  if (freq === "3D") return diffDays >= 3;
  if (freq === "WEEKLY") return diffDays >= 7;

  if (freq === "MONTHLY") {
    return now.getFullYear() !== lastRun.getFullYear() || now.getMonth() !== lastRun.getMonth();
  }

  return false;
}

// Busca robusta por header (evita quebrar quando mudarem posições)
function queryTransacoesBaseClaraPorEtiqueta_(ini, fim, time, lojas, etiqueta) {
  var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
  var sh = ss.getSheetByName("BaseClara");
  if (!sh) return [];

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  // =========================
  // ÍNDICES FIXOS (0-based)
  // BaseClara: A..W
  // =========================
  var I_DATA = 0;        // A - Data da Transação
  var I_ESTAB = 2;       // C - Transação (Estabelecimento)
  var I_VALOR_RS = 5;    // F - Valor em R$
  var I_LOJA = 7;        // H - Alias Do Cartão (Loja)
  var I_TIME = 17;       // R - Grupos (Time)
  var I_ETIQUETAS = 19;  // T - Etiquetas
  var I_DESC = 20;       // U - Descrição (Item comprado)

  // Lê somente A..W (23 colunas)
  var NUM_COLS = 23;
  var values = sh.getRange(2, 1, lastRow - 1, NUM_COLS).getValues();

  // Normalização
  var norm = function (s) {
    return String(s == null ? "" : s).trim().toLowerCase();
  };

  // Set de lojas permitidas (comparação por string exata normalizada)
  var lojasSet = {};
  (lojas || []).forEach(function (l) {
    var x = String(l == null ? "" : l).trim();
    if (x) lojasSet[norm(x)] = true;
  });
  var hasLojasFilter = Object.keys(lojasSet).length > 0;

  // Etiqueta alvo (normalizada)
  var etqTarget = norm(etiqueta);

  // Se Etiquetas vier com múltiplos valores (ex.: "Alimentação, Viagem"),
  // vamos considerar "match" se uma das etiquetas for exatamente igual ao alvo
  var etiquetaMatch = function (cellValue) {
    if (!etqTarget) return true;
    var raw = String(cellValue == null ? "" : cellValue);
    if (!raw.trim()) return false;

    // separadores comuns
    var parts = raw.split(/[,;|]+/).map(function (p) { return norm(p); }).filter(Boolean);
    if (!parts.length) return (norm(raw) === etqTarget);

    for (var i = 0; i < parts.length; i++) {
      if (parts[i] === etqTarget) return true;
    }
    return false;
  };

  // Garantir que ini/fim são Date
  var iniD = (ini instanceof Date) ? ini : new Date(ini);
  var fimD = (fim instanceof Date) ? fim : new Date(fim);

  if (!(iniD instanceof Date) || isNaN(iniD.getTime())) return [];
  if (!(fimD instanceof Date) || isNaN(fimD.getTime())) return [];

  var out = [];
  var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";

  for (var i = 0; i < values.length; i++) {
    var r = values[i];

    var loja = String(r[I_LOJA] == null ? "" : r[I_LOJA]).trim();   // Alias do Cartão
    var t = String(r[I_TIME] == null ? "" : r[I_TIME]).trim();     // Grupos
    var etqCell = r[I_ETIQUETAS];

    // Filtros: time e lojas
    // ✅ "__ALL__" significa não filtrar time
    if (time && time !== "__ALL__" && t !== time) continue;

    if (hasLojasFilter) {
      if (!loja) continue;
      if (!lojasSet[norm(loja)]) continue;
    }

    // Filtro: etiqueta
    if (!etiquetaMatch(etqCell)) continue;

    // Data
    var d = r[I_DATA];
    var dObj = (d instanceof Date) ? d : new Date(d);
    if (!(dObj instanceof Date) || isNaN(dObj.getTime())) continue;
    if (dObj < iniD || dObj > fimD) continue;

    // Valor (F pode vir como número ou string)
    var v = r[I_VALOR_RS];
    var valor = (typeof v === "number") ? v : Number(String(v || "").replace(/\./g, "").replace(",", "."));

    out.push({
      loja: loja,
      time: t,
      data: Utilities.formatDate(dObj, tz, "dd/MM/yyyy"),
      estabelecimento: String(r[I_ESTAB] == null ? "" : r[I_ESTAB]),
      valor: isNaN(valor) ? 0 : valor,
      etiqueta: String(etqCell == null ? "" : etqCell),
      descricao: String(r[I_DESC] == null ? "" : r[I_DESC])
    });
  }

  return out;
}

/**
 * Busca demissões de gerentes (Senior/RH via BigQuery) a partir de uma data (inclusive).
 * Retorna colunas: matricula, des_email_corporativo, des_titulo_cargo,
 * nom_apelido_filial, nom_afastamento, dat_afastamento (dd/MM/yyyy)
 *
 * @param {string} desdeIso - "YYYY-MM-DD" (ex: "2025-12-01")
 * @return {object} { ok: true, rows: [...] } ou { ok: false, error: "..." }
 */
function normalizarDataParaISO_(input) {
  var s = (input || "").toString().trim();

  // já está ISO (YYYY-MM-DD)
  if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;

  // está DD/MM/YYYY -> converte para YYYY-MM-DD
  var m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
  if (m) return m[3] + "-" + m[2] + "-" + m[1];

  // fallback seguro
  return "2025-12-01";
}

function getUltimasDemissoesGerentes() {
    vektorAssertFunctionAllowed_("getUltimasDemissoesGerentes");
  try {
    var query = `
      WITH base AS (
        SELECT
          cod_matricula_colaborador AS matricula,
          des_email_corporativo,
          des_titulo_cargo,
          nom_apelido_filial,
          nom_afastamento,
          COALESCE(
            SAFE_CAST(dat_afastamento AS DATE),
            DATE(SAFE_CAST(dat_afastamento AS TIMESTAMP)),
            DATE(SAFE_CAST(dat_afastamento AS DATETIME)),
            SAFE.PARSE_DATE('%d/%m/%Y', CAST(dat_afastamento AS STRING))
          ) AS dat_afastamento_date
        FROM \`cnto-data-lake.refined.cnt_ref_gld_dim_snr_funcionarios\`
        WHERE dat_chave >= "2023-01-01"
          AND des_titulo_cargo LIKE '%GERENTE%'
          AND nom_afastamento = "Demitido"
      )
      SELECT
        matricula,
        des_email_corporativo,
        des_titulo_cargo,
        nom_apelido_filial,
        nom_afastamento,
        FORMAT_DATE('%d/%m/%Y', dat_afastamento_date) AS dat_afastamento
      FROM base
      WHERE dat_afastamento_date IS NOT NULL
        AND dat_afastamento_date >= DATE("2025-12-01")
      ORDER BY dat_afastamento_date DESC
    `;

    var request = {
      query: query,
      useLegacySql: false
    };

    var result = BigQuery.Jobs.query(request, PROJECT_ID);

    if (!result || result.jobComplete !== true) {
      throw new Error("Falha ao executar consulta no BigQuery (demissões).");
    }

    var rows = [];
    if (result.rows && result.rows.length) {
      rows = result.rows.map(function (r) {
        var f = r.f || [];
        return {
          matricula:             (f[0] && f[0].v) ? f[0].v : "",
          des_email_corporativo: (f[1] && f[1].v) ? f[1].v : "",
          des_titulo_cargo:      (f[2] && f[2].v) ? f[2].v : "",
          nom_apelido_filial:    (f[3] && f[3].v) ? f[3].v : "",
          nom_afastamento:       (f[4] && f[4].v) ? f[4].v : "",
          dat_afastamento:       (f[5] && f[5].v) ? f[5].v : ""
        };
      });
    }

    return { ok: true, rows: rows };

  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}

/**
 * Normaliza o código da loja (ex: "297" -> "0297")
 * e verifica se existe na tabela BigQuery
 * `cnto-data-lake.refined.cnt_ref_gld_dim_estrutura_loja` (coluna cod_loja).
 *
 * @param {string|number} lojaInformada
 * @return {string|null} código 4 dígitos se existir, senão null
 */

function normalizarLojaSeExistir(lojaInformada) {
  // nada informado
  if (lojaInformada === null || lojaInformada === undefined || lojaInformada === "") {
    return null;
  }

  // mantém só dígitos
  var apenasDigitos = String(lojaInformada).replace(/\D/g, '');
  if (!apenasDigitos) return null;

  // força 4 dígitos (297 -> 0297)
  var codigo4 = ('0000' + apenasDigitos).slice(-4);

  // 🔎 monta a query no BigQuery
  // OBS: assumindo que cod_loja pode ser comparado como STRING
  var query = ''
    + 'SELECT cod_loja '
    + 'FROM ' + BQ_TABLE_LOJAS + ' '
    + 'WHERE CAST(cod_loja AS STRING) = "' + codigo4 + '" '
    + 'LIMIT 1';

  var request = {
    query: query,
    useLegacySql: false
  };

  // Executa a query no BigQuery (serviço avançado)
  var queryResults = BigQuery.Jobs.query(request, PROJECT_ID);

  if (!queryResults || queryResults.jobComplete !== true) {
    throw new Error('Falha ao executar consulta no BigQuery para validar loja.');
  }

  var rows = queryResults.rows;
  if (rows && rows.length > 0) {
    // Existe ao menos um registro de cod_loja = codigo4
    return codigo4;
  }

  // Não achou a loja
  return null;
}



/**
 * Retorna o nome da loja (coluna nom_shopping)
 * a partir do código informado (cod_loja).
 *
 * @param {string|number} lojaCodigo
 * @return {object} { ok: true, nome: "CATUAÍ CASCAVEL" } ou { ok: false }
 */

function obterNomeLojaBigQuery(lojaCodigo) {
  try {
    if (!lojaCodigo) return { ok: false, error: "Código não informado." };

    var apenasDigitos = String(lojaCodigo).replace(/\D/g, '');
    if (!apenasDigitos) return { ok: false, error: "Código inválido." };

    var codigo4 = ('0000' + apenasDigitos).slice(-4);

    var query = `
      SELECT nom_shopping
      FROM \`cnto-data-lake.refined.cnt_ref_gld_dim_estrutura_loja\`
      WHERE CAST(cod_loja AS STRING) = "${codigo4}"
      LIMIT 1
    `;

    var request = {
      query: query,
      useLegacySql: false
    };

    var result = BigQuery.Jobs.query(request, PROJECT_ID);

    if (!result || !result.jobComplete) {
      throw new Error("Falha ao consultar BigQuery.");
    }

    if (!result.rows || result.rows.length === 0) {
      return { ok: false, error: "Loja não encontrada." };
    }

    var nome = result.rows[0].f[0].v || "";
    return { ok: true, nome: nome };

  } catch (err) {
    return { ok: false, error: err.message || err };
  }
}


/**
 * Função interna que lê CLARA_PEND e devolve:
 * - última data de cobrança da loja
 * - apenas linhas dessa data
 * - apenas linhas com alguma pendência K:N = "SIM"
 * Formato:
 * {
 *   ok: true,
 *   loja: "171",
 *   ultimaData: "29/10/2025",
 *   header: [...],
 *   rows: [ [B..G + textoPendencias], ... ]
 * }
 */

function _obterPendenciasLoja(lojaCodigo) {
  var lojaParam = (lojaCodigo || "").toString().trim().replace(/\D/g, "");
  var lojaNumero = lojaParam.replace(/^0+/, ""); // "0171" -> "171"

  if (!lojaNumero) {
    throw new Error("Código de loja inválido.");
  }

  var aba = getClaraPendSheet_();

  var values = aba.getDataRange().getValues();
  if (!values || values.length <= 5) {
    throw new Error("Aba 'CLARA_PEND' sem dados suficientes.");
  }

  var headerRowIndex = 4; // linha 5
  var header = values[headerRowIndex];
  var dados  = values.slice(headerRowIndex + 1); // a partir da linha 6

  // Índices zero-based das colunas usadas
  var COL_DATA_COBR  = 1;  // B
  var COL_DATA_TRANS = 2;  // C
  var COL_TRANSACAO  = 3;  // D
  var COL_VALOR      = 4;  // E
  var COL_CARTAO     = 5;  // F
  var COL_LOJA       = 6;  // G
  var COL_ETIQUETA   = 10; // K
  var COL_COMENT     = 11; // L
  var COL_NF         = 12; // M
  var COL_VALOR_D    = 13; // N

  var linhasLoja = [];
  var datasCob   = [];

  dados.forEach(function (linha) {
    var colLoja = (linha[COL_LOJA] || "").toString();
    var lojaDigits = colLoja.replace(/\D/g, "").replace(/^0+/, "");

    if (lojaDigits === lojaNumero) {
      var dataBruta = linha[COL_DATA_COBR];
      var dataObj   = null;

      if (dataBruta instanceof Date) {
        dataObj = dataBruta;
      } else if (typeof dataBruta === "string" && dataBruta.trim() !== "") {
        var partes = dataBruta.split("/");
        if (partes.length === 3) {
          dataObj = new Date(partes[2], partes[1] - 1, partes[0]);
        }
      }

      if (dataObj) {
        datasCob.push(dataObj);
      }
      linhasLoja.push(linha);
    }
  });

  if (linhasLoja.length === 0) {
    return {
      ok: true,
      loja: lojaNumero,
      ultimaData: "",
      header: [],
      rows: []
    };
  }

  if (datasCob.length === 0) {
    throw new Error("Não foi possível identificar a última data de cobrança.");
  }

  var ultimaData = new Date(Math.max.apply(null, datasCob));
  var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
  var dataFormatada = Utilities.formatDate(ultimaData, tz, "dd/MM/yyyy");

  var linhasFiltradas = [];

  linhasLoja.forEach(function (linha) {
    var dataLinha = linha[COL_DATA_COBR];
    var dataLinhaObj = null;

    if (dataLinha instanceof Date) {
      dataLinhaObj = dataLinha;
    } else if (typeof dataLinha === "string" && dataLinha.trim() !== "") {
      var partes = dataLinha.split("/");
      if (partes.length === 3) {
        dataLinhaObj = new Date(partes[2], partes[1] - 1, partes[0]);
      }
    }

    if (!dataLinhaObj) return;

    var mesmaData =
      dataLinhaObj.getFullYear() === ultimaData.getFullYear() &&
      dataLinhaObj.getMonth() === ultimaData.getMonth() &&
      dataLinhaObj.getDate() === ultimaData.getDate();

    if (!mesmaData) return;

    // monta texto de pendências K:N (só se tiver SIM)
    var pendencias = [];

    if ((linha[COL_ETIQUETA] || "").toString().toUpperCase() === "SIM") {
      pendencias.push("Etiqueta pendente");
    }
    if ((linha[COL_COMENT] || "").toString().toUpperCase() === "SIM") {
      pendencias.push("Comentário pendente");
    }
    if ((linha[COL_NF] || "").toString().toUpperCase() === "SIM") {
      pendencias.push("NF/Recibo pendente");
    }
    if ((linha[COL_VALOR_D] || "").toString().toUpperCase() === "SIM") {
      pendencias.push("Valor NF divergente");
    }

    if (pendencias.length === 0) return;

    var dataCobrFormat = (dataLinhaObj instanceof Date)
      ? Utilities.formatDate(dataLinhaObj, tz, "dd/MM/yyyy")
      : (linha[COL_DATA_COBR] || "");

        var dataTransFormat = "";
    var dataTransBruta = linha[COL_DATA_TRANS];

    if (dataTransBruta instanceof Date) {
      // ✅ evita “voltar 1 dia” quando o Date veio como UTC midnight (date-only)
      var isUtcMidnight =
        dataTransBruta.getUTCHours() === 0 &&
        dataTransBruta.getUTCMinutes() === 0 &&
        dataTransBruta.getUTCSeconds() === 0;

      dataTransFormat = Utilities.formatDate(
        dataTransBruta,
        isUtcMidnight ? "GMT" : tz,
        "dd/MM/yyyy"
      );

    } else {
      var s = String(dataTransBruta || "").trim();

      // ✅ ISO yyyy-mm-dd (não passa por new Date)
      var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (m) dataTransFormat = m[3] + "/" + m[2] + "/" + m[1];
      else dataTransFormat = s;
    }

    linhasFiltradas.push([
      dataCobrFormat,
      dataTransFormat,
      linha[COL_TRANSACAO],
      linha[COL_VALOR],
      linha[COL_CARTAO],
      linha[COL_LOJA],
      pendencias.join(", ")
    ]);
  });

  return {
    ok: true,
    loja: lojaNumero,
    ultimaData: dataFormatada,
    header: [
      "Data Cobrança",
      "Data da Transação",
      "Transação",
      "Valor original",
      "Cartão",
      "Loja",
      "Pendências"
    ],
    rows: linhasFiltradas
  };
}

/**
 * Usado pelo front (chat) para mostrar tabela de pendências no chat.
 */

function getPendenciasPorLoja(lojaCodigo) {
  vektorAssertFunctionAllowed_("getPendenciasPorLoja");
  try {
    // 🆕 normaliza + valida na BASE
    const lojaNormalizada = normalizarLojaSeExistir(lojaCodigo);

    if (!lojaNormalizada) {
      // Loja NÃO existe na planilha BASE
      return {
        ok: true,
        lojaInvalida: true
      };
    }

    // Usa a loja normalizada (ex.: "0297") no fluxo de pendências
    return _obterPendenciasLoja(lojaNormalizada);

  } catch (err) {
    return {
      ok: false,
      error: err.toString()
    };
  }
}

/**
 * Envia e-mail com pendências (usado depois do usuário informar o e-mail no chat).
 */

function enviarPendenciasPorEmail(lojaCodigo, emailDestino) {
  vektorAssertFunctionAllowed_("enviarPendenciasPorEmail");
  try {
    if (!emailDestino) {
      return { ok: false, error: "E-mail não informado." };
    }

    var emailUsuario = Session.getActiveUser().getEmail();
if (!emailUsuario) {
  return { ok: false, error: "Usuário sem e-mail ativo." };
}

// ✅ usa o parâmetro recebido do front; fallback para o e-mail do usuário logado
emailDestino = String(emailDestino || emailUsuario).trim();

// 🔒 trava domínio
var emailRegex = /^[^\s@]+@((gruposbf|centauro|fisia)\.com\.br)$/i;
if (!emailRegex.test(emailDestino)) {
  return {
    ok: false,
    error: "Informe um e-mail válido dos domínios do Grupo SBF."
  };
}

// CC somente se o destinatário for diferente do usuário
var ccEmail = "";
if (emailUsuario.toLowerCase() !== emailDestino.toLowerCase()) {
  ccEmail = emailUsuario;
}

    var dados = _obterPendenciasLoja(lojaCodigo);
    if (!dados.ok) {
      return dados;
    }

    if (!dados.rows || dados.rows.length === 0) {
      return {
        ok: false,
        error: "Não há pendências com 'SIM' na última data de cobrança."
      };
    }

    var lojaNumero    = dados.loja;
    var dataFormatada = dados.ultimaData;
    var tz            = Session.getScriptTimeZone() || "America/Sao_Paulo";

    var tabelaHtml = '<table style="border-collapse:collapse;width:100%;font-family:Arial, sans-serif;font-size:12px;">';
    tabelaHtml += '<thead><tr style="background-color:#003366;color:#ffffff;">';
    dados.header.forEach(function (h) {
      tabelaHtml += '<th style="border:1px solid #cccccc;padding:6px;">' + h + '</th>';
    });
    tabelaHtml += '</tr></thead><tbody>';

    dados.rows.forEach(function (linha) {
      tabelaHtml += '<tr>';
      linha.forEach(function (col) {
        tabelaHtml += '<td style="border:1px solid #cccccc;padding:6px;">' +
          (col !== undefined && col !== null ? col : '') +
          '</td>';
      });
      tabelaHtml += '</tr>';
    });

    tabelaHtml += '</tbody></table>';

    var agora = new Date();
    var hora  = parseInt(Utilities.formatDate(agora, tz, "HH"), 10);
    var saudacao = "Boa tarde!";
    if (hora < 12) {
      saudacao = "Bom dia!";
    } else if (hora >= 18) {
      saudacao = "Boa noite!";
    }

    var assunto = "Apontamento de Pendências | Loja " + lojaNumero;

    var corpoHtml = ""
      + "<p>" + saudacao + "</p>"
      + "<p>Segue o resumo das pendências Clara da loja <b>" + lojaNumero + "</b> "
      + "(data de cobrança <b>" + dataFormatada + "</b>), conforme falamos via chat. "
      + "Essa é a última data de cobrança, sempre verifique no app da Clara se não há mais transações além das apontadas:</p>"
      + tabelaHtml
      + "<br/><br/>"
      + "<p><b>Agente Vektor - Contas a Receber</b></p>";

    GmailApp.sendEmail(emailDestino, assunto, " ", {
      from: "vektor@gruposbf.com.br",
      cc: "rodrigo.lisboa@gruposbf.com.br",
      replyTo: "contasareceber@gruposbf.com.br",
      htmlBody: corpoHtml,
      name: "Vektor Grupo SBF"
    });

    registrarAlertaEnviado_(
  "PENDENCIAS_LOJA",
  lojaNumero,
  "",
  "Pendências enviadas por e-mail (data cobrança " + dataFormatada + "). Itens=" + ((dados.rows || []).length),
  emailDestino,
  "enviarPendenciasPorEmail"
);

    return {
      ok: true,
      loja: lojaNumero,
      data: dataFormatada
    };

  } catch (err) {
    return {
      ok: false,
      error: "Erro interno ao enviar e-mail: " + err
    };
  }
}

// Pendências para bloqueio: usa mesma aba CLARA_PEND, mas pega as 2 últimas datas de cobrança

function getPendenciasParaBloqueio(lojaCodigo) {
  vektorAssertFunctionAllowed_("getPendenciasParaBloqueio");
  try {
    // 🆕 normaliza + valida na BASE
    const lojaNormalizada = normalizarLojaSeExistir(lojaCodigo);

    if (!lojaNormalizada) {
      // Loja NÃO existe na planilha BASE
      return {
        ok: true,
        lojaInvalida: true
      };
    }

    // remove zeros à esquerda para comparar com a coluna de loja da CLARA_PEND
    var lojaNumero = lojaNormalizada.replace(/^0+/, ""); // "0171" -> "171"

    // Mesma planilha / aba usada no fluxo normal de pendências
      var aba;
      try {
        aba = getClaraPendSheet_();
      } catch (e) {
        return { ok: false, error: e.toString() };
      }


    var values = aba.getDataRange().getValues();
    if (!values || values.length <= 5) {
      return { ok: false, error: "Aba 'CLARA_PEND' sem dados suficientes." };
    }

    var headerRowIndex = 4; // linha 5
    var dados  = values.slice(headerRowIndex + 1); // a partir da linha 6

    // Índices zero-based das colunas usadas (mesmos da _obterPendenciasLoja)

    var COL_DATA_COBR  = 1;  // B
    var COL_DATA_TRANS = 2;  // C
    var COL_TRANSACAO  = 3;  // D
    var COL_VALOR      = 4;  // E
    var COL_CARTAO     = 5;  // F
    var COL_LOJA       = 6;  // G
    var COL_ETIQUETA   = 10; // K
    var COL_COMENT     = 11; // L
    var COL_NF         = 12; // M
    var COL_VALOR_D    = 13; // N

    var linhasLoja = [];
    var datasChave = [];

    function chaveData(d) {
      var ano = d.getFullYear();
      var mes = "" + (d.getMonth() + 1);
      var dia = "" + d.getDate();
      if (mes.length < 2) mes = "0" + mes;
      if (dia.length < 2) dia = "0" + dia;
      return ano + "-" + mes + "-" + dia; // yyyy-mm-dd
    }

    // Filtra linhas da loja e coleta datas de cobrança
    dados.forEach(function (linha) {
      var colLoja = (linha[COL_LOJA] || "").toString();
      var lojaDigits = colLoja.replace(/\D/g, "").replace(/^0+/, "");

      if (lojaDigits === lojaNumero) {
        var dataBruta = linha[COL_DATA_COBR];
        var dataObj   = null;

        if (dataBruta instanceof Date) {
          dataObj = dataBruta;
        } else if (typeof dataBruta === "string" && dataBruta.trim() !== "") {
          var partes = dataBruta.split("/");
          if (partes.length === 3) {
            dataObj = new Date(partes[2], partes[1] - 1, partes[0]);
          }
        }

        if (dataObj) {
          var chave = chaveData(dataObj);
          datasChave.push(chave);
        }
        linhasLoja.push(linha);
      }
    });

    // Loja existe na BASE, mas não tem pendências na CLARA_PEND
    if (linhasLoja.length === 0) {
      return {
        ok: true,
        loja: lojaNumero,
        html: '<p class="text-sm text-slate-700">Não encontrei pendências para esta loja.</p>'
      };
    }

    if (datasChave.length === 0) {
      return { ok: false, error: "Não foi possível identificar datas de cobrança para esta loja." };
    }

    // Remove duplicadas e ordena datas (mais recente primeiro)
    var datasUnicas = [];
    datasChave.forEach(function (c) {
      if (datasUnicas.indexOf(c) === -1) {
        datasUnicas.push(c);
      }
    });
    datasUnicas.sort(function (a, b) {
      // yyyy-mm-dd em string mantém ordem cronológica
      if (a < b) return 1;
      if (a > b) return -1;
      return 0;
    });

    // Pega as 2 últimas datas de cobrança
    var datasSelecionadas = datasUnicas.slice(0, 2);

    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    var linhasFiltradas = [];

    // Agora filtra as linhas da loja só pelas datas selecionadas
    dados.forEach(function (linha) {
      var colLoja = (linha[COL_LOJA] || "").toString();
      var lojaDigits = colLoja.replace(/\D/g, "").replace(/^0+/, "");
      if (lojaDigits !== lojaNumero) return;

      var dataLinha = linha[COL_DATA_COBR];
      var dataLinhaObj = null;

      if (dataLinha instanceof Date) {
        dataLinhaObj = dataLinha;
      } else if (typeof dataLinha === "string" && dataLinha.trim() !== "") {
        var partes = dataLinha.split("/");
        if (partes.length === 3) {
          dataLinhaObj = new Date(partes[2], partes[1] - 1, partes[0]);
        }
      }

      if (!dataLinhaObj) return;

      var chaveLinha = chaveData(dataLinhaObj);
      if (datasSelecionadas.indexOf(chaveLinha) === -1) {
        return; // não está entre as 2 últimas datas de cobrança
      }

      // monta texto de pendências K:N (só se tiver SIM)
      var pendencias = [];

      if ((linha[COL_ETIQUETA] || "").toString().toUpperCase() === "SIM") {
        pendencias.push("Etiqueta pendente");
      }
      if ((linha[COL_COMENT] || "").toString().toUpperCase() === "SIM") {
        pendencias.push("Comentário pendente");
      }
      if ((linha[COL_NF] || "").toString().toUpperCase() === "SIM") {
        pendencias.push("NF/Recibo pendente");
      }
      if ((linha[COL_VALOR_D] || "").toString().toUpperCase() === "SIM") {
        pendencias.push("Valor NF divergente");
      }

      if (pendencias.length === 0) return;

      var dataCobrFormat = (dataLinhaObj instanceof Date)
        ? Utilities.formatDate(dataLinhaObj, tz, "dd/MM/yyyy")
        : (linha[COL_DATA_COBR] || "");

          var dataTransFormat = "";
    var dataTransBruta = linha[COL_DATA_TRANS];

    if (dataTransBruta instanceof Date) {
      // ✅ evita “voltar 1 dia” quando o Date veio como UTC midnight (date-only)
      var isUtcMidnight =
        dataTransBruta.getUTCHours() === 0 &&
        dataTransBruta.getUTCMinutes() === 0 &&
        dataTransBruta.getUTCSeconds() === 0;

      dataTransFormat = Utilities.formatDate(
        dataTransBruta,
        isUtcMidnight ? "GMT" : tz,
        "dd/MM/yyyy"
      );

    } else {
      var s = String(dataTransBruta || "").trim();

      // ✅ ISO yyyy-mm-dd (não passa por new Date)
      var m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (m) dataTransFormat = m[3] + "/" + m[2] + "/" + m[1];
      else dataTransFormat = s;
    }

      linhasFiltradas.push([
        dataCobrFormat,
        dataTransFormat,
        linha[COL_TRANSACAO],
        linha[COL_VALOR],
        linha[COL_CARTAO],
        linha[COL_LOJA],
        pendencias.join(", ")
      ]);
    });

    if (linhasFiltradas.length === 0) {
      return {
        ok: true,
        loja: lojaNumero,
        html: '<p class="text-sm text-slate-700">Não encontrei pendências recentes para esta loja.</p>'
      };
    }

    // Monta HTML da tabela (mesmas colunas do fluxo normal de pendências)
    var headers = [
      "Data Cobrança",
      "Data da Transação",
      "Transação",
      "Valor original",
      "Cartão",
      "Loja",
      "Pendências"
    ];

    var html = ""
      + '<div class="text-sm text-slate-700">'
      + '<p>Encontrei abaixo as últimas pendências relacionadas ao cartão da loja <b>' + lojaNumero + '</b>.<br/>'
      + 'Essas pendências podem ter ocasionado o bloqueio do cartão.<br/><br/>'
      + '</p>'
      + '<div class="mt-2 overflow-x-auto">'
      + '<table class="min-w-full text-xs border border-slate-200">'
      + '<thead class="bg-slate-100"><tr>';

    headers.forEach(function (h) {
      html += '<th class="border px-2 py-1 text-left">' + h + '</th>';
    });

    html += '</tr></thead><tbody>';

    linhasFiltradas.forEach(function (linha) {
      html += '<tr>';
      for (var i = 0; i < linha.length; i++) {
        var col = linha[i];
        html += '<td class="border px-2 py-1">'
          + (col !== undefined && col !== null ? col : "")
          + '</td>';
      }
      html += '</tr>';
    });

    html += '</tbody></table></div></div>';

    return {
      ok: true,
      loja: lojaNumero,
      html: html
    };

  } catch (err) {
    return { ok: false, error: err.message || err.toString() };
  }
}

/**
 * Normaliza texto para comparação:
 * - transforma em string
 * - trim
 * - remove acentos
 * - deixa tudo minúsculo
 */
function normalizarTexto_(str) {
  if (!str) return "";
  return str
    .toString()
    .trim()
    .normalize("NFD")                // separa letras dos acentos
    .replace(/[\u0300-\u036f]/g, "") // remove os acentos
    .toLowerCase();
}

// ========= RELATÓRIO CLARA (PLANILHA Captura_Clara / aba BaseClara) ========= //

var SPREADSHEET_ID_CLARA = '1_XW0IqbYjiCPpqtwdEi1xPxDlIP2MSkMrLGbeinLIeI'; // Captura_Clara
var SHEET_NOME_BASE_CLARA = 'BaseClara';
var SHEET_NOME_INFO_LIMITES = 'Info_limites';

// ========= PLANILHA ANTIGA (onde fica a aba CLARA_PEND) ========= //
var SPREADSHEET_ID_CLARA_PEND = "1jcNdVTxdDYqwHwsOkT7gb_2BdZke9qIb39RiwgTKxUQ"; // planilha antiga
var SHEET_NOME_CLARA_PEND = "CLARA_PEND";

// Abre a aba CLARA_PEND na planilha antiga
function getClaraPendSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID_CLARA_PEND);
  var aba = ss.getSheetByName(SHEET_NOME_CLARA_PEND);
  if (!aba) throw new Error("Aba '" + SHEET_NOME_CLARA_PEND + "' não encontrada na planilha antiga.");
  return aba;
}


// Abre a aba BaseClara
function getBaseClaraSheet_() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID_CLARA);
  return ss.getSheetByName(SHEET_NOME_BASE_CLARA);
}

function parseDateClara_(value) {
  if (value === null || value === undefined || value === "") return null;

  // Date direto
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return isNaN(value.getTime()) ? null : value;
  }

  // Número (data serial do Sheets/Excel)
  if (typeof value === "number") {
    // ✅ usa base LOCAL (evita UTC drift)
    var d0 = new Date(1899, 11, 30, 12, 0, 0); // meio-dia
    var d = new Date(d0.getTime() + value * 24 * 60 * 60 * 1000);
    return isNaN(d.getTime()) ? null : d;
  }

  var s = String(value).trim();
  if (!s) return null;

  // dd/MM/yyyy ou dd/MM/yyyy HH:mm(:ss)
  var m1 = s.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?$/);
  if (m1) {
    var dd = Number(m1[1]), mm = Number(m1[2]) - 1, yy = Number(m1[3]);
    var hh = Number(m1[4] || 12), mi = Number(m1[5] || 0), ss = Number(m1[6] || 0);
    var d1 = new Date(yy, mm, dd, hh, mi, ss);
    return isNaN(d1.getTime()) ? null : d1;
  }

  // ✅ ISO yyyy-MM-dd (com ou sem hora/Z) — NÃO usar new Date("yyyy-mm-dd") puro
  var mIso = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T\s].*)?$/);
  if (mIso) {
    var y = Number(mIso[1]), mo = Number(mIso[2]) - 1, d0 = Number(mIso[3]);
    var dIso = new Date(y, mo, d0, 12, 0, 0); // meio-dia local
    return isNaN(dIso.getTime()) ? null : dIso;
  }

  // fallback
  var d2 = new Date(s);
  return isNaN(d2.getTime()) ? null : d2;
}

// Lê todas as linhas da BaseClara (ignorando cabeçalho)
function carregarLinhasBaseClara_() {
  var sh = getBaseClaraSheet_();
  if (!sh) {
    return { header: [], linhas: [], error: "Aba 'BaseClara' não encontrada na planilha Captura_Clara." };
  }
  var values = sh.getDataRange().getValues();
  if (!values || values.length < 2) {
    return { header: values && values[0] ? values[0] : [], linhas: [], error: null };
  }
  var header = values[0];
  var linhas = values.slice(1);
  return { header: header, linhas: linhas, error: null };
}

// =======================
// NOVO: ENVIO DE PENDÊNCIAS CLARA (RECUSADAS) VIA BASECLARA
// =======================

var SHEET_NOME_EMAILS_LOJAS = "Emails"; // aba Emails na mesma planilha da BaseClara
var VEKTOR_SLACK_GRUPO_CONTAS_A_RECEBER = "contas_a_receber-aaaaiglscd4gbv3eod7ao65qsy@gruposbf.slack.com";
var VEKTOR_CC_CONTAS_A_RECEBER = "contasareceber@gruposbf.com.br";

// =======================
// LOG de envios (envio único por transação)
// =======================
var VEKTOR_ENV_PEND_LOG_TAB = "HIST_ENVIO_PEND_RECUSADAS";

function vektorSha256Hex_(s) {
  var bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(s || ""), Utilities.Charset.UTF_8);
  var out = [];
  for (var i = 0; i < bytes.length; i++) {
    var v = (bytes[i] + 256) % 256;
    var h = v.toString(16);
    if (h.length === 1) h = "0" + h;
    out.push(h);
  }
  return out.join("");
}

function vektorTxKey_(r) {
  // hash de: lojaKey + dataTrans + valor + cartao + estabelecimento
  var base = [
    String(r.lojaKey || "").trim(),
    String(r.dataTransBR || "").trim(),
    String(r.valorOriginalTxt || "").trim(),
    String(r.cartao || "").trim(),
    String(r.estabelecimento || "").trim(),
    String(r.codigoAutorizacao || "").trim()
  ].join("||");
  return vektorSha256Hex_(base);
}

function vektorGetOrCreateEnvPendLogSheet_() {
  var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
  var sh = ss.getSheetByName(VEKTOR_ENV_PEND_LOG_TAB);
  if (!sh) {
    sh = ss.insertSheet(VEKTOR_ENV_PEND_LOG_TAB);
    sh.appendRow(["sentAt", "txKey", "lojaKey", "dataTransBR", "valorOriginalTxt", "cartao", "codigoAutorizacao", "estabelecimento", "pendenciasTxt", "to", "cc", "status", "error"]);
    sh.getRange(1, 1, 1, 13).setFontWeight("bold");
    sh.setFrozenRows(1);
  }
  return sh;
}

function vektorCarregarTxKeysJaEnviadas_() {
  var sh = vektorGetOrCreateEnvPendLogSheet_();
  var values = sh.getDataRange().getValues();
  var map = {};

  if (!values || values.length < 2) return map;

  // Descobre índices pelo header (linha 0)
  var hdr = values[0] || [];
  function idx_(name) {
    var n = String(name || "").trim().toLowerCase();
    for (var i = 0; i < hdr.length; i++) {
      if (String(hdr[i] || "").trim().toLowerCase() === n) return i;
    }
    return -1;
  }

  var iTx = idx_("txkey");     // coluna "txKey"
  var iSt = idx_("status");    // coluna "status"

  // fallback caso header não bata
  if (iTx < 0) iTx = 1;
  if (iSt < 0) iSt = hdr.length - 2; // normalmente penúltima

  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    if (!row || !row.length) continue;

    var tx = String(row[iTx] || "").trim();
    var st = String(row[iSt] || "").trim().toUpperCase();

    if (tx && st === "SENT") map[tx] = true;
  }

  return map;
}

function vektorLogEnvioPendencia_(payload) {
  try {
    var sh = vektorGetOrCreateEnvPendLogSheet_();
    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    var ts = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

    sh.appendRow([
      ts,
      payload.txKey || "",
      payload.lojaKey || "",
      payload.dataTransBR || "",
      payload.valorOriginalTxt || "",
      payload.cartao || "",
      payload.codigoAutorizacao || "", // ✅ NOVO
      payload.estabelecimento || "",
      payload.pendenciasTxt || "",
      payload.to || "",
      payload.cc || "",
      payload.status || "",
      payload.error || ""
]);
  } catch (e) {
    Logger.log("Falha ao logar envio pendência: " + (e && e.message ? e.message : e));
  }
}

function vektorGetHistoricoEnviosPendenciasResumo() {
  vektorAssertFunctionAllowed_("vektorGetHistoricoEnviosPendenciasResumo");

  var sh = vektorGetOrCreateEnvPendLogSheet_();
  var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
  var lastRow = sh.getLastRow();

  function parseMoneyPtBr_(v) {
    if (v === null || v === undefined || v === "") return 0;
    if (typeof v === "number") return isNaN(v) ? 0 : v;

    var s = String(v).trim();
    s = s.replace(/\s/g, "");
    s = s.replace(/R\$/gi, "");
    s = s.replace(/\./g, "");
    s = s.replace(/,/g, ".");
    var n = Number(s);
    return isNaN(n) ? 0 : n;
  }

  function parseSentAtSafe_(v) {
    if (v instanceof Date) return isNaN(v.getTime()) ? null : v;

    var s = String(v || "").trim();
    if (!s) return null;

    var m = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2})(?::(\d{2}))?)?$/);
    if (m) {
      var d1 = new Date(
        Number(m[1]),
        Number(m[2]) - 1,
        Number(m[3]),
        Number(m[4] || 0),
        Number(m[5] || 0),
        Number(m[6] || 0)
      );
      return isNaN(d1.getTime()) ? null : d1;
    }

    var d2 = new Date(s);
    return isNaN(d2.getTime()) ? null : d2;
  }

  function fmtSentAtRaw_(v) {
    if (!v) return "";

    if (v instanceof Date) {
      if (isNaN(v.getTime())) return "";
      return Utilities.formatDate(v, tz, "yyyy-MM-dd HH:mm:ss");
    }

    var s = String(v).trim();
    if (!s) return "";

    var m = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?$/);
    if (m) {
      return m[1] + "-" + m[2] + "-" + m[3] + " " + m[4] + ":" + m[5] + ":" + (m[6] || "00");
    }

    var d = new Date(s);
    if (!isNaN(d.getTime())) {
      return Utilities.formatDate(d, tz, "yyyy-MM-dd HH:mm:ss");
    }

    return s;
  }

  function fmtLastSentBr_(raw) {
    var s = String(raw || "").trim();
    if (!s) return "";

    var m = s.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})(?::(\d{2}))?$/);
    if (!m) return s;

    return m[3] + "/" + m[2] + "/" + m[1] + " - " + m[4] + ":" + m[5] + ":" + (m[6] || "00");
  }

  function monthKey_(d) {
    return Utilities.formatDate(d, tz, "yyyy-MM");
  }

  function dayKey_(d) {
    return Utilities.formatDate(d, tz, "yyyy-MM-dd");
  }

  var now = new Date();
  var monthSeeds = [];
  var monthMap = {};

  for (var k = 5; k >= 0; k--) {
    var d = new Date(now.getFullYear(), now.getMonth() - k, 1);
    var ym = Utilities.formatDate(d, tz, "yyyy-MM");
    var rec = {
      ym: ym,
      label: Utilities.formatDate(d, tz, "MM/yyyy"),
      qtd: 0,
      daysMap: {}
    };
    monthSeeds.push(rec);
    monthMap[ym] = rec;
  }

  if (lastRow < 2) {
    return {
      ok: true,
      lastSentAt: "",
      lastSentAtBr: "",
      months: monthSeeds.map(function(m){
        return { ym: m.ym, label: m.label, qtd: 0, days: [] };
      })
    };
  }

  var numRows = lastRow - 1;

  // A = sentAt
  // C = lojaKey
  // E = valorOriginalTxt
  // L = status
  // ✅ sentAt deve vir como TEXTO exibido na planilha para não sofrer deslocamento de fuso
  var sentAtCol = sh.getRange(2, 1, numRows, 1).getDisplayValues();
  var lojaCol   = sh.getRange(2, 3, numRows, 1).getValues();
  var valorCol  = sh.getRange(2, 5, numRows, 1).getValues();
  var statusCol = sh.getRange(2, 12, numRows, 1).getValues();

  var cutoff = new Date(now.getFullYear(), now.getMonth() - 5, 1);
  cutoff.setHours(0, 0, 0, 0);

  var lastSent = null;
  var lastSentRaw = "";
  var achouDentroJanela = false;

  for (var i = numRows - 1; i >= 0; i--) {
    var status = String((statusCol[i] && statusCol[i][0]) || "").trim().toUpperCase();
    if (status !== "SENT") continue;

        var sentCell = String((sentAtCol[i] && sentAtCol[i][0]) || "").trim();
    var dt = parseSentAtSafe_(sentCell);
    if (!dt) continue;

    if (!lastSent) {
      lastSent = dt;
      lastSentRaw = fmtSentAtRaw_(sentCell);
    }

    if (dt < cutoff) {
      if (achouDentroJanela) break;
      continue;
    }

    achouDentroJanela = true;

    var ym = monthKey_(dt);
    var bucket = monthMap[ym];
    if (!bucket) continue;

    bucket.qtd++;

    var dk = dayKey_(dt);
    if (!bucket.daysMap[dk]) {
      bucket.daysMap[dk] = {
        dia: dk,
        diaBr: Utilities.formatDate(dt, tz, "dd/MM"),
        qtd: 0,
        valorTotal: 0,
        lojasMap: {}
      };
    }

    var day = bucket.daysMap[dk];
    day.qtd++;
    day.valorTotal += parseMoneyPtBr_((valorCol[i] && valorCol[i][0]) || 0);

    var loja = String((lojaCol[i] && lojaCol[i][0]) || "").trim();
    if (loja) day.lojasMap[loja] = true;
  }

  var months = monthSeeds.map(function(m){
    var days = Object.keys(m.daysMap)
      .sort()
      .map(function(k){
        var d = m.daysMap[k];
        return {
          dia: d.dia,
          diaBr: d.diaBr,
          qtd: d.qtd,
          valorTotal: Number(d.valorTotal || 0),
          lojas: Object.keys(d.lojasMap).sort()
        };
      });

    return {
      ym: m.ym,
      label: m.label,
      qtd: m.qtd,
      days: days
    };
  });

  return {
    ok: true,
    lastSentAt: lastSentRaw || "",
    lastSentAtBr: fmtLastSentBr_(lastSentRaw),
    months: months
  };
}

function vektorPingHistoricoEnvios() {
  vektorAssertFunctionAllowed_("vektorPingHistoricoEnvios");

  return {
    ok: true,
    lastSentAt: "2026-03-11 10:30:00",
    lastSentAtBr: "11/03/2026 - 10:30:00",
    months: [
      { ym: "2025-10", label: "10/2025", qtd: 1, days: [{ dia:"2025-10-10", diaBr:"10/10", qtd:1, valorTotal:100, lojas:["CE0001"] }] },
      { ym: "2025-11", label: "11/2025", qtd: 2, days: [{ dia:"2025-11-15", diaBr:"15/11", qtd:2, valorTotal:300, lojas:["CE0002","CE0003"] }] },
      { ym: "2025-12", label: "12/2025", qtd: 0, days: [] },
      { ym: "2026-01", label: "01/2026", qtd: 3, days: [{ dia:"2026-01-08", diaBr:"08/01", qtd:3, valorTotal:450, lojas:["CE0004"] }] },
      { ym: "2026-02", label: "02/2026", qtd: 1, days: [{ dia:"2026-02-19", diaBr:"19/02", qtd:1, valorTotal:99, lojas:["CE0005"] }] },
      { ym: "2026-03", label: "03/2026", qtd: 4, days: [{ dia:"2026-03-02", diaBr:"02/03", qtd:4, valorTotal:1200, lojas:["CE0006","CE0007"] }] }
    ]
  };
}

function vektorNormLojaKey_(v) {
  // Aceita "CE0062", "0062", 62 etc -> retorna "CE0062"
  var s = String(v || "").trim().toUpperCase();
  if (!s) return "";
  var m = s.match(/CE\s*(\d{1,6})/i);
  var digits = "";
  if (m && m[1]) digits = m[1];
  else digits = s.replace(/\D/g, "");
  if (!digits) return "";
  var cod4 = ("0000" + digits).slice(-4);
  return "CE" + cod4;
}

function vektorCarregarMapaEmailsLojas_() {
  var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
  var sh = ss.getSheetByName(SHEET_NOME_EMAILS_LOJAS);
  if (!sh) throw new Error("Aba '" + SHEET_NOME_EMAILS_LOJAS + "' não encontrada na planilha BaseClara.");

  var lr = sh.getLastRow();
  if (lr < 2) return {};

  // A..G
  var values = sh.getRange(1, 1, lr, 7).getValues();
  // Cabeçalho esperado (mas vamos por posição, como você definiu):
  // A Loja, B LojaNorm, C Shopping, D Time, E Email Gerente, F Nome Gerente Regional, G Email Gerente Regional

  var map = {}; // "CE0001" -> { time, emailGerente, emailRegional }
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var lojaKey = vektorNormLojaKey_(row[0]); // A
    if (!lojaKey) continue;

    map[lojaKey] = {
      lojaKey: lojaKey,
      time: String(row[3] || "").trim(),              // D
      emailGerente: String(row[4] || "").trim(),      // E
      emailRegional: String(row[6] || "").trim()      // G
    };
  }
  return map;
}

function vektorIsBlank_(v) {
  if (v === null || v === undefined) return true;
  if (v === false) return true;
  var s = String(v).trim();
  if (!s) return true;
  var low = s.toLowerCase();
  return (low === "null" || low === "-" || low === "n/a");
}

function vektorIsReciboPendente_(v) {
  // Você pediu: pendência tudo que estiver como "Não" na coluna O
  var s = String(v || "").trim().toLowerCase();
  return s === "não" || s === "nao";
}

function vektorSaudacaoPorHora_() {
  var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
  var agora = new Date();
  var hora = parseInt(Utilities.formatDate(agora, tz, "HH"), 10);
  if (hora < 12) return "Bom dia!";
  if (hora >= 18) return "Boa noite!";
  return "Boa tarde!";
}

function vektorFormatDateBR_(d) {
  try {
    // Se vier string, tenta converter (blindagem)
    if (!(d instanceof Date)) d = parseDateClara_(d);

    if (!(d instanceof Date) || isNaN(d.getTime())) return String(d || "").trim();

    // ✅ Como o Vektor exibe só DATA (sem hora), formate sempre em GMT e elimina -1 dia
    return Utilities.formatDate(d, "GMT", "dd/MM/yyyy");
  } catch (e) {
    return String(d || "").trim();
  }
}

function vektorQueryPendenciasRecusadas_(ini, fim) {
  var base = carregarLinhasBaseClara_();
  if (base.error) throw new Error(base.error);

  var header = base.header || [];
  var linhas = base.linhas || [];

  // Use header lookup para reduzir risco de coluna trocada.
  var iDataTrans = encontrarIndiceColuna_(header, ["Data da Transação"]);
  var iEstab     = encontrarIndiceColuna_(header, ["Transação"]); 
  var iValor     = encontrarIndiceColuna_(header, ["Valor original"]);
  var iCartao    = encontrarIndiceColuna_(header, ["Cartão"]);
  var iAlias     = encontrarIndiceColuna_(header, ["Alias Do Cartão"]);
  var iLojaNum   = encontrarIndiceColuna_(header, ["LojaNum"]);
  var iStatusAp  = encontrarIndiceColuna_(header, ["Status de aprovação"]);
  var iRecibo    = encontrarIndiceColuna_(header, ["Recibo"]);
  var iEtiqueta  = encontrarIndiceColuna_(header, ["Etiquetas"]);
  var iDesc      = encontrarIndiceColuna_(header, ["Descrição"]);
  var iNotaAprov = encontrarIndiceColuna_(header, ["Nota do aprovador"]);  // ✅ NOVO
  var iCodAut    = encontrarIndiceColuna_(header, ["Código de autorização"]);

  // Falhas críticas
    var req = [
      ["Data da Transação", iDataTrans],
      ["Estabelecimento", iEstab],
      ["Valor original", iValor],
      ["Cartão", iCartao],
      ["Alias Do Cartão", iAlias],
      ["Status de aprovação", iStatusAp],
      ["Recibo", iRecibo],
      ["Etiquetas", iEtiqueta],
      ["Descrição", iDesc],
      ["Nota do aprovador", iNotaAprov],     // ✅ vírgula aqui
      ["Código de autorização", iCodAut]     // ✅ sem vírgula no último (opcional)
    ];

    req.forEach(function (p) {
      if (!p || p.length < 2) throw new Error("Erro interno: item inválido em req (pendências).");
      if (p[1] < 0) throw new Error("Não encontrei a coluna '" + p[0] + "' no cabeçalho da BaseClara.");
    });

  var emailsMap = vektorCarregarMapaEmailsLojas_();

  var out = []; // registros “linha a linha”
  for (var i = 0; i < linhas.length; i++) {
    var row = linhas[i];

    var dt = parseDateClara_(row[iDataTrans]);
    if (!dt) continue;

    // filtro período (inclusivo)
    if (ini && dt < ini) continue;
    if (fim) {
      var fim23 = new Date(fim.getFullYear(), fim.getMonth(), fim.getDate(), 23, 59, 59);
      if (dt > fim23) continue;
    }

    var status = String(row[iStatusAp] || "").trim().toLowerCase();
    if (status !== "recusada") continue;

    var pendRecibo = vektorIsReciboPendente_(row[iRecibo]);
    var pendEtq    = vektorIsBlank_(row[iEtiqueta]);
    var pendDesc   = vektorIsBlank_(row[iDesc]);

    // ✅ NOVO: divergência NF/recibo (coluna L - Nota do aprovador)
    var notaAprov = String(row[iNotaAprov] || "");
    var pendDivergNF = /nf\/?recibo\s+divergente/i.test(notaAprov);

    // Se não tiver nenhuma pendência, ignora
    if (!pendRecibo && !pendEtq && !pendDesc && !pendDivergNF) continue;

    var pendList = [];
    if (pendRecibo)   pendList.push("Nota fiscal/Recibo");
    if (pendEtq)      pendList.push("Etiqueta");
    if (pendDesc)     pendList.push("Descrição");
    if (pendDivergNF) pendList.push("NF/Recibo divergente"); // ✅ NOVO

    var lojaKey = vektorNormLojaKey_(row[iAlias] || row[iLojaNum]);
    if (!lojaKey) continue;

    var contato = emailsMap[lojaKey] || { emailGerente: "", emailRegional: "", time: "" };

    var obj = {
      lojaKey: lojaKey,

      // ⛔ NÃO retorne Date pro front
      // dataTrans: dt,

      // ✅ retorne string serializável
      dataTransISO: (dt instanceof Date) ? dt.toISOString() : "",
      dataTransBR: vektorFormatDateBR_(dt),

      estabelecimento: String(row[iEstab] || "").trim(),
      valorOriginal: row[iValor],
      valorOriginalTxt: String(row[iValor] || "").trim(),
      cartao: String(row[iCartao] || "").trim(),
      statusAprovacao: String(row[iStatusAp] || "").trim(),
      codigoAutorizacao: String(row[iCodAut] || "").trim(),   // se já existe aí no teu obj
      pendenciasTxt: pendList.join(", "),
      pendRecibo: pendRecibo,
      pendEtq: pendEtq,
      pendDesc: pendDesc,
      pendDivergNF: pendDivergNF,
      emailGerente: String(contato.emailGerente || "").trim(),
      emailRegional: String(contato.emailRegional || "").trim()
    };

    obj.txKey = vektorTxKey_(obj);
    out.push(obj);   
      }

  return out;
}

function vektorMontarTabelaPendenciasEmail_(rows) {
  function esc_(x){
    return String(x===null||x===undefined?"":x)
      .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
      .replace(/"/g,"&quot;").replace(/'/g,"&#039;");
  }

  var thBase = "background:#0b1f3a;color:#fff;border:1px solid #0f172a;padding:7px;font-size:12px;white-space:nowrap;text-align:left;";
  var thPend = "background:#ef4444;color:#fff;border:1px solid #0f172a;padding:7px;font-size:12px;white-space:nowrap;text-align:left;"; // vermelho claro
  var tdBase = "border:1px solid #0f172a;padding:7px;font-size:12px;vertical-align:top;white-space:nowrap;color:#0f172a;";

  var html = "";
  html += "<table style='border-collapse:collapse;width:100%;font-family:Arial,sans-serif;'>";
  html += "<thead><tr>";
  html += "<th style='" + thBase + "'>Data da Transação</th>";
  html += "<th style='" + thBase + "'>Estabelecimento</th>";
  html += "<th style='" + thBase + "'>Valor original</th>";
  html += "<th style='" + thBase + "'>Cartão</th>";
  html += "<th style='" + thPend + "'>Pendências</th>";
  html += "</tr></thead><tbody>";

  (rows || []).forEach(function(r){
    html += "<tr>";
    html += "<td style='" + tdBase + "'>" + esc_(r.dataTransBR || "") + "</td>";
    html += "<td style='" + tdBase + "'>" + esc_(r.estabelecimento || "") + "</td>";
    html += "<td style='" + tdBase + "'>" + esc_(r.valorOriginalTxt || "") + "</td>";
    html += "<td style='" + tdBase + "'>" + esc_(r.cartao || "") + "</td>";
    html += "<td style='" + tdBase + "'><b>" + esc_(r.pendenciasTxt || "") + "</b></td>";
    html += "</tr>";
  });

  html += "</tbody></table>";
  return html;
}

function vektorTiposPendenciasDoGrupo_(rows) {
  var set = {};
  (rows || []).forEach(function (r) {
    if (r.pendRecibo) set["Nota fiscal/Recibo"] = true;
    if (r.pendEtq) set["Etiqueta"] = true;
    if (r.pendDesc) set["Descrição"] = true;
    if (r.pendDivergNF) set["NF/Recibo divergente"] = true;
  });
  return Object.keys(set);
}

function vektorMontarCorpoEmailPendenciasClara_(saudacao, tabelaHtml, tiposPendencias) {
  var tipos = (tiposPendencias && tiposPendencias.length) ? tiposPendencias : [];
  var tiposTxt = tipos.length ? tipos.join(", ") : "justificativas";

  var html = "";
  html += "<div style='font-family:Arial,sans-serif;font-size:13px;color:#0f172a;line-height:1.45'>";
  html += "<p>Pessoal, " + saudacao + "</p>";

  // ✅ Texto variável conforme tipos
  html += "<p>Seguem abaixo transações pendentes de <b>" + tiposTxt + "</b> dentro do prazo de 48 horas após a compra, precisamos que sejam corrigidas o mais rápido possível. Assim que as pendências forem regularizadas, solicitamos a gentileza de responder a este e-mail confirmando a correção.</p>";

  html += "<p>O bloqueio do cartão já foi efetuado preventivamente, para que possamos seguir com o desbloqueio, encaminhe um chamado via Servicenow, caminho: Contas a Receber &gt; Cartão Clara &gt; Solicitação de Desbloqueio de Cartão.</p>";

  // ❌ REMOVIDO: bloco "Lembrando que para todas as transações..."
  html += tabelaHtml;

  html += "<br/><br/>";
  html += "<p><i>Caso tenha dúvidas ou precise de mais informações, entre em contato conosco.</i></p>";
  html += "<br/><br/>";
  html += "<p>Atenciosamente,<br/>Contas a Receber<br/>Grupo SBF<br/>contasareceber@gruposbf.com.br</p>";
  html += "</div>";
  return html;
}

// ✅ FIX DEFINITIVO: parser ISO robusto (evita TypeError m[1])
function vektorParseIsoDateSafe_(iso) {
  if (!iso) return null;

  // aceita "2026-01-26T00:00:00.000Z" ou "2026-01-26"
  var s = String(iso).trim();
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;

  var y = Number(m[1]);
  var mo = Number(m[2]) - 1;
  var d = Number(m[3]);

  var dt = new Date(y, mo, d);
  dt.setHours(0, 0, 0, 0);
  return isNaN(dt.getTime()) ? null : dt;
}

// ✅ parser de valor robusto (aceita number OU string pt-BR)
function vektorParseValorBR_(x) {
  if (x === null || x === undefined) return NaN;

  // se já é número, não destrói decimal
  if (typeof x === "number") return isNaN(x) ? NaN : x;

  var s = String(x).trim();
  if (!s) return NaN;

  // remove "R$" e espaços
  s = s.replace(/[R$\s]/g, "");

  // Se tem vírgula, assume pt-BR: "." milhar e "," decimal
  if (s.indexOf(",") >= 0) {
    s = s.replace(/\./g, "").replace(",", ".");
    return Number(s);
  }

  // Se não tem vírgula:
  // - se tem ponto, assume que ponto é decimal (vindo como 722.46)
  // - se não tem ponto, é inteiro
  return Number(s);
}

// 1) PREVIEW: devolve resumo pro chat
function previewEnvioPendenciasClaraRecusadas(dataInicioIso, dataFimIso) {
  vektorAssertFunctionAllowed_("previewEnvioPendenciasClaraRecusadas");

  var ini = vektorParseIsoDateSafe_(dataInicioIso);
  var fim = vektorParseIsoDateSafe_(dataFimIso);
  if (!ini || !fim) return { ok: false, error: "Informe data inicial e final válidas." };

  var rows = vektorQueryPendenciasRecusadas_(ini, fim);

  var total = rows.length;
  var cRec = 0, cEtq = 0, cDesc = 0, cDiv = 0; // ✅ NOVO
  var totalValorRecusadas = 0;
  var mapaValorPorLoja = {}; // { "0062": 1234.56, ... }

  // ✅ mapa Loja -> Time (vem da aba Emails/Lojas)
      var lojaInfoMap = {};
      try {
        lojaInfoMap = vektorCarregarMapaEmailsLojas_(); // { "CE0062": { time, ... }, ... }
      } catch (eMap) {
        lojaInfoMap = {};
      }

  rows.forEach(function (r) {
    // ✅ GARANTIA: só conta RECUSADA (protege contra qualquer mudança no query)
    var st = String(r.statusAprovacao || r.statusAprovacaoTxt || r.status || "").toUpperCase().trim();
    if (st && st !== "RECUSADA" && st !== "RECUSADO") return;

    if (r.pendRecibo) cRec++;
    if (r.pendEtq) cEtq++;
    if (r.pendDesc) cDesc++;
    if (r.pendDivergNF) cDiv++; // ✅ NOVO

    // soma “valor original” de forma tolerante

    var v = String(r.valorOriginal || r.valorOriginalTxt || "")
    .replace(/[R$\s]/g, "")
    .replace(/\./g, "")
    .replace(",", ".");
      var n = Number(v);

      if (!isNaN(n)) totalValorRecusadas += n;
    });

  var lojasSet = {};
  rows.forEach(function (r) { lojasSet[r.lojaKey] = true; });

  var lojasValorArr = Object.keys(mapaValorPorLoja).map(function(k){

  var lk = vektorNormLojaKey_(k); // garante "CE0000"
  var t = (lojaInfoMap[lk] && lojaInfoMap[lk].time) ? String(lojaInfoMap[lk].time) : "—";
  return { lojaKey: lk, valor: mapaValorPorLoja[k] || 0, time: t };

    });

    // maior valor primeiro
    lojasValorArr.sort(function(a,b){
      return (b.valor || 0) - (a.valor || 0);
    });

  return {
    ok: true,
    periodo: { inicio: vektorFormatDateBR_(ini), fim: vektorFormatDateBR_(fim) },
    totalTransacoes: total,
    totalLojas: Object.keys(lojasSet).length,

    pendRecibo: cRec,
    pendEtiqueta: cEtq,
    pendDescricao: cDesc,
    pendDivergNF: cDiv, // ✅ NOVO

    pctRecibo: total ? (cRec / total) : 0,
    pctEtiqueta: total ? (cEtq / total) : 0,
    pctDescricao: total ? (cDesc / total) : 0,
    pctDivergNF: total ? (cDiv / total) : 0, // ✅ NOVO

    totalValor: totalValorRecusadas,
    lojasValor: lojasValorArr,
  };
}

function vektorCalcularStatusMixPeriodo_(ini, fim) {
  var base = carregarLinhasBaseClara_();
  if (base.error) throw new Error(base.error);

  var header = base.header || [];
  var linhas = base.linhas || [];

  var iDataTrans = encontrarIndiceColuna_(header, ["Data da Transação"]);
  var iStatusAp  = encontrarIndiceColuna_(header, ["Status de aprovação"]);

  if (iDataTrans < 0) throw new Error("Não encontrei a coluna 'Data da Transação' no cabeçalho da BaseClara.");
  if (iStatusAp < 0) throw new Error("Não encontrei a coluna 'Status de aprovação' no cabeçalho da BaseClara.");

  var total = 0;
  var aprovada = 0;
  var recusada = 0;
  var necessita = 0;

  for (var i = 0; i < linhas.length; i++) {
    var row = linhas[i];

    var dt = parseDateClara_(row[iDataTrans]);
    if (!dt) continue;

    // filtro período (inclusivo)
    if (ini && dt < ini) continue;
    if (fim) {
      var fim23 = new Date(fim.getFullYear(), fim.getMonth(), fim.getDate(), 23, 59, 59);
      if (dt > fim23) continue;
    }

    var status = String(row[iStatusAp] || "").trim().toLowerCase();
    if (!status) continue; // ignora em branco (estorno)

    // normaliza variações
    if (status === "aprovada" || status === "aprovado") {
      total++;
      aprovada++;
      continue;
    }

    if (status === "recusada" || status === "recusado") {
      total++;
      recusada++;
      continue;
    }

    if (status === "necessita aprovação" || status === "necessita aprovacao") {
      total++;
      necessita++;
      continue;
    }

    // se aparecer outro status, ignore (ou inclua em "outros" se quiser)
  }

  var pct = function(x) { return total ? (x / total) : 0; };

  return {
    total: total,
    aprovada: aprovada,
    recusada: recusada,
    necessita: necessita,
    pctAprovada: pct(aprovada),
    pctRecusada: pct(recusada),
    pctNecessita: pct(necessita)
  };
}

function vektorCalcularValorRecusadasPorLojaPeriodo_(ini, fim) {
  var base = carregarLinhasBaseClara_();
  if (base.error) throw new Error(base.error);

  var header = base.header || [];
  var linhas = base.linhas || [];

  var iDataTrans = encontrarIndiceColuna_(header, ["Data da Transação"]);
  var iStatusAp  = encontrarIndiceColuna_(header, ["Status de aprovação"]);
  var iValor     = encontrarIndiceColuna_(header, ["Valor original"]);
  var iAlias     = encontrarIndiceColuna_(header, ["Alias Do Cartão"]);
  var iLojaNum   = encontrarIndiceColuna_(header, ["LojaNum"]);

  if (iDataTrans < 0) throw new Error("Não encontrei 'Data da Transação' na BaseClara.");
  if (iStatusAp  < 0) throw new Error("Não encontrei 'Status de aprovação' na BaseClara.");
  if (iValor     < 0) throw new Error("Não encontrei 'Valor original' na BaseClara.");
  if (iAlias < 0 && iLojaNum < 0) throw new Error("Não encontrei 'Alias Do Cartão' nem 'LojaNum' na BaseClara.");

  var totalValor = 0;
  var mapa = {}; // lojaKey -> soma

  for (var i = 0; i < linhas.length; i++) {
    var row = linhas[i];

    var dt = parseDateClara_(row[iDataTrans]);
    if (!dt) continue;

    // período inclusivo
    if (ini && dt < ini) continue;
    if (fim) {
      var fim23 = new Date(fim.getFullYear(), fim.getMonth(), fim.getDate(), 23, 59, 59);
      if (dt > fim23) continue;
    }

    var st = String(row[iStatusAp] || "").trim().toLowerCase();
    if (st !== "recusada" && st !== "recusado") continue;

    var n = vektorParseValorBR_(row[iValor]);
    if (!isFinite(n)) n = 0;

    var lojaKey = vektorNormLojaKey_(row[iAlias] || row[iLojaNum]);
    if (!lojaKey) continue;

    totalValor += n;
    mapa[lojaKey] = (mapa[lojaKey] || 0) + n;
  }

  var lojasValorArr = Object.keys(mapa).map(function(k){
    return { lojaKey: k, valor: mapa[k] || 0 };
  }).sort(function(a,b){
    return (b.valor || 0) - (a.valor || 0);
  });

  return {
    totalValor: totalValor,
    mapaValorPorLoja: mapa,
    lojasValorArr: lojasValorArr
  };
}

function previewEnvioPendenciasClaraRecusadasDetalhado(dataInicioIso, dataFimIso) {
  vektorAssertFunctionAllowed_("previewEnvioPendenciasClaraRecusadasDetalhado");

  try {
    // ✅ 1) Parse robusto (não confia só no vektorParseIsoDateSafe_)
    // Aceita "YYYY-MM-DD" e também ISO com hora.
    var ini0 = vektorParseIsoDateSafe_(dataInicioIso) || vektorParseDateAny_(dataInicioIso) || parseDateClara_(dataInicioIso);
    var fim0 = vektorParseIsoDateSafe_(dataFimIso)   || vektorParseDateAny_(dataFimIso)   || parseDateClara_(dataFimIso);

    // Re-normaliza via parseDateClara_ (corrige “date-only contaminado”)
    var ini = parseDateClara_(ini0);
    var fim = parseDateClara_(fim0);

    if (!ini || !fim) return { ok: false, error: "Informe data inicial e final válidas." };

    // ✅ 2) Janela inclusiva (dia inteiro)
    ini = new Date(ini.getFullYear(), ini.getMonth(), ini.getDate(), 0, 0, 0, 0);
    fim = new Date(fim.getFullYear(), fim.getMonth(), fim.getDate(), 23, 59, 59, 999);

    var rows = vektorQueryPendenciasRecusadas_(ini, fim);

    // ✅ NOVO: Status mix do período (Aprovada / Recusada / Necessita aprovação)
    var statusMix = vektorCalcularStatusMixPeriodo_(ini, fim); // { total, aprovada, recusada, necessita, pct... }

    var sentMap = vektorCarregarTxKeysJaEnviadas_() || {};
    if (typeof sentMap !== "object") sentMap = {};

    var total = rows.length;
    var cRec = 0, cEtq = 0, cDesc = 0, cDiv = 0;
    var totalValorRecusadas = 0;
    var mapaValorPorLoja = {}; // lojaKey -> valor
    var lojasSet = {};
    var jaEnviadas = 0;

    // ✅ série por data (para gráfico)
    var porData = {}; // { 'dd/MM/yyyy': qtd }

    function vektorParseValorBRL_(valorOriginal, valorTxt) {
      if (typeof valorOriginal === "number" && isFinite(valorOriginal)) {
        return valorOriginal;
      }
      var s = String(valorTxt || valorOriginal || "").trim();
      if (!s) return 0;
      s = s.replace(/[R$\s]/g, "");
      if (s.indexOf(",") >= 0) {
        s = s.replace(/\./g, "").replace(",", ".");
        var n1 = Number(s);
        return isNaN(n1) ? 0 : n1;
      }
      var n2 = Number(s);
      return isNaN(n2) ? 0 : n2;
    }

    // ✅ pendências por loja (para tooltip do donut)
    var pendPorLoja = {}; // lojaKey -> {rec, etq, desc, div, totalFlags}

    rows.forEach(function (r) {
      lojasSet[r.lojaKey] = true;

      if (r.pendRecibo) cRec++;
      if (r.pendEtq) cEtq++;
      if (r.pendDesc) cDesc++;
      if (r.pendDivergNF) cDiv++;

      // ✅ agrega por loja + tipo (stacked)
      var lk = String(r.lojaKey || "").trim().toUpperCase();
      if (lk) {
        var hasAny = !!(r.pendRecibo || r.pendEtq || r.pendDesc || r.pendDivergNF);
        if (hasAny) {
          if (!pendPorLoja[lk]) pendPorLoja[lk] = { rec: 0, etq: 0, desc: 0, div: 0, totalFlags: 0 };

          if (r.pendRecibo) { pendPorLoja[lk].rec++;  pendPorLoja[lk].totalFlags++; }
          if (r.pendEtq)    { pendPorLoja[lk].etq++;  pendPorLoja[lk].totalFlags++; }
          if (r.pendDesc)   { pendPorLoja[lk].desc++; pendPorLoja[lk].totalFlags++; }
          if (r.pendDivergNF){pendPorLoja[lk].div++;  pendPorLoja[lk].totalFlags++; }
        }
      }

      var tx = String(r.txKey || "").trim();
      r.jaEnviado = !!(tx && sentMap[tx]);
      if (r.jaEnviado) jaEnviadas++;

      // ===== VALOR TOTAL (RECUSADAS COM PENDÊNCIA) =====
      var n = vektorParseValorBRL_(r.valorOriginal, r.valorOriginalTxt);
      var st = String(r.statusAprovacao || "").toLowerCase().trim();
      if (
        (st === "recusada" || st === "recusado") &&
        (r.pendRecibo || r.pendEtq || r.pendDesc || r.pendDivergNF)
      ) {
        totalValorRecusadas += n;
        var lk2 = String(r.lojaKey || "").trim().toUpperCase();
        if (lk2) {
          mapaValorPorLoja[lk2] = (mapaValorPorLoja[lk2] || 0) + n;
        }
      }

      // ✅ GARANTIA: nada de Date crua no payload
      if (r.dataTrans instanceof Date) r.dataTransISO = r.dataTrans.toISOString();
      delete r.dataTrans;

      // ✅ agrega por data (1 por transação)
      var d = String(r.dataTransBR || "").trim();
      if (d) {
        if (!porData[d]) porData[d] = 0;
        porData[d]++;
      }
    });

    var lojasValorArr = Object.keys(mapaValorPorLoja).map(function(k){
      return { lojaKey: k, valor: mapaValorPorLoja[k] || 0 };
    }).sort(function(a,b){
      return (b.valor || 0) - (a.valor || 0);
    });

    // ===== ENRIQUECE lojasValorArr com TIME (para filtro no popup) =====
    var mapEmailsLojas = {};
    try {
      mapEmailsLojas = vektorCarregarMapaEmailsLojas_() || {};
    } catch (eMap) {
      mapEmailsLojas = {};
    }

    lojasValorArr = (lojasValorArr || []).map(function (it) {
      var lk3 = String(it && it.lojaKey ? it.lojaKey : "").trim().toUpperCase();
      var time = "";
      if (lk3 && mapEmailsLojas[lk3] && mapEmailsLojas[lk3].time) {
        time = String(mapEmailsLojas[lk3].time || "").trim();
      }
      return {
        lojaKey: lk3,
        valor: Number(it && it.valor ? it.valor : 0) || 0,
        time: time || "—"
      };
    });

    // ordena datas (dd/MM/yyyy -> yyyy-MM-dd)
    var seriePorData = Object.keys(porData)
      .sort(function (a, b) {
        var pa = a.split("/").reverse().join("-");
        var pb = b.split("/").reverse().join("-");
        return pa.localeCompare(pb);
      })
      .map(function (d) {
        return { data: d, total: porData[d] };
      });

    // ✅ array ordenado para o tooltip (maior volume de pendências primeiro)
    var lojasPendStack = Object.keys(pendPorLoja).map(function(k){
      var o = pendPorLoja[k];
      var den = o.totalFlags || 1;
      return {
        lojaKey: k,
        totalFlags: o.totalFlags || 0,
        pctRecibo:  (o.rec  || 0) / den,
        pctEtiqueta:(o.etq  || 0) / den,
        pctDescricao:(o.desc|| 0) / den,
        pctDivergNF:(o.div  || 0) / den
      };
    }).sort(function(a,b){
      return (b.totalFlags||0) - (a.totalFlags||0);
    });

    return {
      ok: true,
      resumo: {
        periodo: { inicio: vektorFormatDateBR_(ini), fim: vektorFormatDateBR_(fim) },
        totalTransacoes: total,
        totalLojas: Object.keys(lojasSet).length,
        pendRecibo: cRec,
        pendEtiqueta: cEtq,
        pendDescricao: cDesc,
        pendDivergNF: cDiv,
        lojasPendStack: lojasPendStack,
        pctRecibo: total ? (cRec / total) : 0,
        pctEtiqueta: total ? (cEtq / total) : 0,
        pctDescricao: total ? (cDesc / total) : 0,
        pctDivergNF: total ? (cDiv / total) : 0,
        totalValor: totalValorRecusadas,
        lojasValor: lojasValorArr,
        totalJaEnviadas: jaEnviadas,
        statusMix: statusMix
      },
      seriePorData: seriePorData,
      rows: rows
    };

  } catch (e) {
    var msg = (e && e.message) ? e.message : String(e);
    var st = (e && e.stack) ? String(e.stack) : "";
    return { ok: false, error: msg + (st ? ("\n" + st) : "") };
  }
}

function dispararEnvioPendenciasClaraRecusadasSelecionadas(dataInicioIso, dataFimIso, txKeys) {
  vektorAssertFunctionAllowed_("dispararEnvioPendenciasClaraRecusadasSelecionadas");

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    return { ok: false, error: "Já existe um envio em andamento. Aguarde ~30s e tente novamente." };
  }

  try {
    var ini = vektorParseIsoDateSafe_(dataInicioIso);
    var fim = vektorParseIsoDateSafe_(dataFimIso);
    if (!ini || !fim) return { ok: false, error: "Informe data inicial e final válidas." };

    txKeys = Array.isArray(txKeys) ? txKeys : [];
    var want = {};
    txKeys.forEach(function (k) {
      k = String(k || "").trim();
      if (k) want[k] = true;
    });
    if (!Object.keys(want).length) return { ok: false, error: "Nenhuma transação selecionada." };

    var rowsAll = vektorQueryPendenciasRecusadas_(ini, fim);

    // ✅ envio único (carrega APÓS pegar o lock, para evitar concorrência)
    var sentMap = vektorCarregarTxKeysJaEnviadas_();

    // filtra selecionadas e ainda não enviadas
    var rows = [];
    rowsAll.forEach(function (r) {
      var tx = String(r.txKey || "").trim();
      if (!tx || !want[tx]) return;
      if (sentMap[tx]) return; // já enviada uma vez -> nunca reenviar
      rows.push(r);
    });

    if (!rows.length) {
      return { ok: false, error: "Todas as selecionadas já foram enviadas anteriormente (envio único)." };
    }

    // agrupa por loja (assunto por data de cobrança = hoje)
    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    var dataCobrancaBR = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy");

    var grupos = {}; // lojaKey -> rows[]
    rows.forEach(function (r) {
      if (!grupos[r.lojaKey]) grupos[r.lojaKey] = [];
      grupos[r.lojaKey].push(r);
    });

    var saudacao = vektorSaudacaoPorHora_();
    var sucessoPorLoja = [];
    var falhasPorLoja = [];

    var emailsEnviados = 0;
    var txRegistradas = 0;

    Object.keys(grupos).forEach(function (lojaKey) {
      var itens = grupos[lojaKey];

      // destinatários
      var toSet = {};
      function addTo_(em) {
        em = String(em || "").trim();
        if (!em) return;
        toSet[em.toLowerCase()] = em;
      }
      addTo_(itens[0].emailGerente);
      addTo_(itens[0].emailRegional);
      addTo_(VEKTOR_SLACK_GRUPO_CONTAS_A_RECEBER);

      var toList = Object.keys(toSet).map(function (k) { return toSet[k]; }).join(",");
      if (!toList) {
        falhasPorLoja.push({ lojaKey: lojaKey, error: "Sem destinatários (gerente/regional/slack) na aba Emails." });
        return;
      }

      var assunto = "CLARA | JUSTIFICATIVAS PENDENTES | " + lojaKey + " - " + dataCobrancaBR;

      var tabela = vektorMontarTabelaPendenciasEmail_(itens);
      var tipos = vektorTiposPendenciasDoGrupo_(itens);
      var corpo = vektorMontarCorpoEmailPendenciasClara_(saudacao, tabela, tipos);

      try {
        GmailApp.sendEmail(toList, assunto, " ", {
          from: "vektor@gruposbf.com.br",
          name: "Vektor - Grupo SBF",
          cc: VEKTOR_CC_CONTAS_A_RECEBER,
          replyTo: VEKTOR_CC_CONTAS_A_RECEBER,
          htmlBody: corpo
        });

        emailsEnviados++;

        // log por transação (SENT)
        itens.forEach(function (r) {
          vektorLogEnvioPendencia_({
            txKey: r.txKey,
            lojaKey: lojaKey,
            dataTransBR: r.dataTransBR,
            valorOriginalTxt: r.valorOriginalTxt,
            cartao: r.cartao,
            codigoAutorizacao: r.codigoAutorizacao,
            estabelecimento: r.estabelecimento,
            pendenciasTxt: r.pendenciasTxt,
            to: toList,
            cc: VEKTOR_CC_CONTAS_A_RECEBER,
            status: "SENT",
            error: ""
          });
          txRegistradas++;
        });

        sucessoPorLoja.push({ lojaKey: lojaKey, qtdTx: itens.length });

      } catch (e) {
        var msg = (e && e.message) ? e.message : String(e);

        // log por transação (FAIL)
        itens.forEach(function (r) {
          vektorLogEnvioPendencia_({
            txKey: r.txKey,
            lojaKey: lojaKey,
            dataTransBR: r.dataTransBR,
            valorOriginalTxt: r.valorOriginalTxt,
            cartao: r.cartao,
            codigoAutorizacao: r.codigoAutorizacao,
            estabelecimento: r.estabelecimento,
            pendenciasTxt: r.pendenciasTxt,
            to: toList,
            cc: VEKTOR_CC_CONTAS_A_RECEBER,
            status: "FAIL",
            error: msg
          });
          txRegistradas++; // opcional: conta como "registrada no log" também
        });

        falhasPorLoja.push({ lojaKey: lojaKey, error: msg });
      }
    });

    return {
      ok: true,
      emailsEnviados: emailsEnviados,
      txRegistradas: txRegistradas,
      sucessoPorLoja: sucessoPorLoja,
      falhasPorLoja: falhasPorLoja
    };

  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

function dispararNotificacaoItensIrregularesSelecionados(rowsSelecionadas) {
  vektorAssertFunctionAllowed_("dispararNotificacaoItensIrregularesSelecionados");

  rowsSelecionadas = Array.isArray(rowsSelecionadas) ? rowsSelecionadas : [];
  if (!rowsSelecionadas.length) return { ok: false, error: "Nenhuma linha selecionada." };

  var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
  var hoje = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy");

  var REPLY_TO = "contasareceber@gruposbf.com.br";
  var CC_FIXO = "contasareceber@gruposbf.com.br";
  var SENDER_NAME = "Vektor - Grupo SBF";

  var mapEmails = vektorCarregarMapaEmailsLojas_();

  function normLojaKey_(lojaRaw) {
    var s = String(lojaRaw || "").trim().toUpperCase();
    if (!s) return "";

    var m1 = s.match(/^CE\s*0*(\d{1,4})/);
    if (m1 && m1[1]) return "CE" + ("0000" + m1[1]).slice(-4);

    var m2 = s.match(/(\d{1,4})/);
    if (m2 && m2[1]) return "CE" + ("0000" + m2[1]).slice(-4);

    return "";
  }

  var grupos = {};
  rowsSelecionadas.forEach(function (r) {
    var lk = normLojaKey_(r && r.loja);
    if (!lk) return;
    if (!grupos[lk]) grupos[lk] = [];
    grupos[lk].push(r || {});
  });

  var lojaKeys = Object.keys(grupos);
  if (!lojaKeys.length) {
    return { ok: false, error: "Não foi possível identificar lojas (lojaKey) nas linhas selecionadas." };
  }

  function esc_(s) {
    return String(s || "")
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/'/g, "&#39;");
  }

  function fmtBRL_(n) {
    try {
      return (Number(n || 0)).toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
    } catch (e) {
      return String(n || 0);
    }
  }

  function badgeHtml_(c) {
    var x = String(c || "").toUpperCase();
    var bg = "rgba(255,255,255,0.06)", bd = "rgba(148,163,184,0.25)", tx = "rgba(226,232,240,0.95)";
    if (x === "OK") { bg = "rgba(34,197,94,0.18)"; bd = "rgba(34,197,94,0.35)"; tx = "#14532d"; }
    if (x === "REVISAR") { bg = "rgba(245,158,11,0.20)"; bd = "rgba(245,158,11,0.40)"; tx = "#713f12"; }
    if (x === "ALERTA") { bg = "rgba(248,113,113,0.20)"; bd = "rgba(248,113,113,0.40)"; tx = "#7f1d1d"; }

    return '<span style="display:inline-flex; align-items:center; height:22px; padding:0 10px; border-radius:999px;'
      + 'border:1px solid ' + bd + '; background:' + bg + '; color:' + tx + '; font-weight:1000; font-size:11px;">'
      + esc_(x || "—") + '</span>';
  }

  function montarTabelaHtml_(rows, lojaKey) {
    function normTxt_(s) {
      return String(s || "")
        .toLowerCase()
        .normalize("NFD")
        .replace(/[\u0300-\u036f]/g, "");
    }

    function isViagemHosp_(r) {
      var txt = normTxt_((r.item || "") + " " + (r.estabelecimento || "") + " " + (r.motivo || ""));
      return (
        txt.indexOf("viagem") !== -1 ||
        txt.indexOf("hosped") !== -1 ||
        txt.indexOf("hotel") !== -1 ||
        txt.indexOf("passagem") !== -1 ||
        txt.indexOf("aereo") !== -1 ||
        txt.indexOf("aerea") !== -1
      );
    }

    var h = '';
    h += '<div style="font-family:Inter,Arial,sans-serif; color:#0f172a;">';
    h += '<div style="font-size:14px; font-weight:900; margin-bottom:10px;">Itens Irregulares</div>';

    var qtd = rows.length;
    var total = 0;
    for (var t = 0; t < rows.length; t++) total += Number(rows[t].valor || 0) || 0;

    h += '<div style="font-size:12px; line-height:1.45; margin-bottom:12px;">'
      + '<div><b>Loja:</b> ' + esc_(lojaKey || "") + '</div>'
      + '<div><b>Quantidade de itens:</b> ' + esc_(qtd) + '</div>'
      + '<div><b>Valor total:</b> ' + esc_(fmtBRL_(total)) + '</div>'
      + '</div>';

    var saudacao = vektorSaudacaoPorHora_();

    h += '<div style="font-size:12px; line-height:1.35; margin-bottom:12px;">'
      + 'Olá, ' + esc_(String(saudacao || "").toLowerCase())
      + '</div>';

    var temViagemHosp = false;
    var temOutrosIrreg = false;
    var temCartaoBloqueado = false;

    for (var j = 0; j < rows.length; j++) {
      if (isViagemHosp_(rows[j] || {})) temViagemHosp = true;
      else temOutrosIrreg = true;

      if (rows[j] && rows[j].cartaoBloqueado === true) temCartaoBloqueado = true;
    }

    h += '<div style="font-size:12px; line-height:1.35; margin-bottom:12px;">'
      + 'Identificamos que os itens abaixo, comprados com o cartão da Clara, não estão em conformidade com nossa Política de Uso dos Cartões. Solicitamos que nos informem o motivo da compra:'
      + '</div>';

    if (temCartaoBloqueado) {
      h += '<div style="font-size:12px; line-height:1.35; margin-bottom:12px;">'
        + 'Paralelo a isso, por medida de segurança, o cartão está previamente bloqueado, onde o desbloqueio deverá ser solicitado através de chamado no ServiceNow: '
        + 'Contas a Receber &gt; Cartão Clara &gt; Solicitação de desbloqueio de cartão.'
        + '</div>';
    }

    if (temViagemHosp && !temOutrosIrreg) {
      h += '<div style="font-size:12px; line-height:1.35; margin-bottom:12px;">'
        + 'Para os casos de passagens/hospedagem, precisamos que entrem no site onde efetuaram a compra e realizem o cancelamento.'
        + '</div>';
    } else if (temViagemHosp && temOutrosIrreg) {
      h += '<div style="font-size:12px; line-height:1.35; margin-bottom:12px;">'
        + 'Para os casos de passagens/hospedagem, também precisamos que realizem o cancelamento no site da operadora.'
        + '</div>';
    }

    h += '<div style="border:1px solid #e2e8f0; border-radius:12px; overflow:hidden;">';
    h += '<table style="width:100%; border-collapse:collapse;">';
    h += '<thead><tr style="background:#0b1220; color:#fff;">';
    h += '<th style="text-align:left; padding:10px; font-size:12px;">Data</th>';
    h += '<th style="text-align:right; padding:10px; font-size:12px;">Valor (R$)</th>';
    h += '<th style="text-align:left; padding:10px; font-size:12px;">Loja</th>';
    h += '<th style="text-align:left; padding:10px; font-size:12px;">Time</th>';
    h += '<th style="text-align:left; padding:10px; font-size:12px;">Item Comprado</th>';
    h += '<th style="text-align:left; padding:10px; font-size:12px;">Estabelecimento</th>';
    h += '<th style="text-align:left; padding:10px; font-size:12px;">Conformidade</th>';
    h += '<th style="text-align:left; padding:10px; font-size:12px;">Motivo</th>';
    h += '</tr></thead><tbody>';

    for (var i = 0; i < rows.length; i++) {
      var r = rows[i] || {};
      h += '<tr style="border-top:1px solid #e2e8f0;">';
      h += '<td style="padding:10px; font-size:12px;">' + esc_(r.dataTxt || "") + '</td>';
      h += '<td style="padding:10px; font-size:12px; text-align:right; font-weight:800;">' + esc_(fmtBRL_(r.valor || 0)) + '</td>';
      h += '<td style="padding:10px; font-size:12px;">' + esc_(r.loja || "") + '</td>';
      h += '<td style="padding:10px; font-size:12px;">' + esc_(r.time || "") + '</td>';
      h += '<td style="padding:10px; font-size:12px;">' + esc_(r.item || "") + '</td>';
      h += '<td style="padding:10px; font-size:12px;">' + esc_(r.estabelecimento || "") + '</td>';
      h += '<td style="padding:10px; font-size:12px;">' + badgeHtml_(r.conformidade || "ALERTA") + '</td>';
      h += '<td style="padding:10px; font-size:12px;">' + esc_(r.motivo || "") + '</td>';
      h += '</tr>';
    }

    h += '</tbody></table></div>';

    h += '<div style="height:16px;"></div>';
    h += '<div style="height:16px;"></div>';

    h += '<div style="font-size:12px; line-height:1.5; color:#0f172a;">'
      + 'Atenciosamente,<br>'
      + 'Contas a Receber<br>'
      + 'Grupo SBF<br>'
      + 'contasareceber@gruposbf.com.br'
      + '</div>';

    h += '</div>';
    return h;
  }

  var enviados = 0;
  var skipped = [];
  var erros = [];

  lojaKeys.forEach(function (lojaKey) {
    var pack = [];
    var to = "";
    var cc = [];
    var subject = "";

    try {
      pack = Array.isArray(grupos[lojaKey]) ? grupos[lojaKey] : [];
      if (!pack.length) return;

      var info = mapEmails ? mapEmails[lojaKey] : null;

      to = info && info.emailGerente ? String(info.emailGerente).trim() : "";
      cc = [];

      if (info && info.emailRegional) cc.push(String(info.emailRegional).trim());
      cc.push(CC_FIXO);

      if (!to && info && info.emailRegional) to = String(info.emailRegional).trim();
      if (!to) to = CC_FIXO;

      cc = cc.filter(function (x) { return x && x.indexOf("@") > 0; });

      var ccUniq = {};
      cc = cc.filter(function (x) {
        var k = String(x || "").toLowerCase();
        if (!k) return false;
        if (ccUniq[k]) return false;
        ccUniq[k] = true;
        return true;
      });

      cc = cc.filter(function (x) {
        return String(x || "").toLowerCase() !== String(to || "").toLowerCase();
      });

      var qtdItens = pack.length;

      subject = "[ALERTA CLARA | ITENS IRREGULARES] "
        + lojaKey + " | "
        + qtdItens + " " + (qtdItens === 1 ? "item" : "itens")
        + " | " + hoje;

      var htmlBody = montarTabelaHtml_(pack, lojaKey);

      GmailApp.sendEmail(to, subject, " ", {
        htmlBody: htmlBody,
        cc: (cc && cc.length ? cc.join(",") : undefined),
        from: "vektor@gruposbf.com.br",
        name: SENDER_NAME,
        replyTo: REPLY_TO
      });

      var timeResumo = "";
      for (var tt = 0; tt < pack.length; tt++) {
        if (String(pack[tt].time || "").trim()) {
          timeResumo = String(pack[tt].time || "").trim();
          break;
        }
      }

      var totalLoja2 = pack.reduce(function (acc, r) {
        return acc + (Number((r || {}).valor || 0) || 0);
      }, 0);

      vektorLogEnvioItensIrreg_({
        lojaKey: lojaKey,
        lojaRaw: (pack[0] && pack[0].loja) ? pack[0].loja : "",
        time: timeResumo,
        qtdItens: pack.length,
        valorTotal: totalLoja2,
        to: to,
        cc: cc.join(","),
        temViagemHosp: pack.some(function (r) {
          var txt = String(((r || {}).item || "") + " " + ((r || {}).estabelecimento || ""))
            .toLowerCase();
          return /viagem|hosped|hotel|passagem|aere/.test(txt);
        }),
        temCartaoBloqueado: pack.some(function (r) {
          return r && r.cartaoBloqueado === true;
        }),
        assunto: subject,
        status: "SENT",
        error: ""
      });

      enviados++;

    } catch (e) {
      var msgErro = String(e && e.message ? e.message : e);
      erros.push({ lojaKey: lojaKey, error: msgErro });

      try {
        var packSafe = Array.isArray(pack) ? pack : [];
        vektorLogEnvioItensIrreg_({
          lojaKey: lojaKey,
          lojaRaw: (packSafe[0] && packSafe[0].loja) ? packSafe[0].loja : "",
          time: (packSafe[0] && packSafe[0].time) ? packSafe[0].time : "",
          qtdItens: packSafe.length,
          valorTotal: packSafe.reduce(function (acc, r) {
            return acc + (Number((r || {}).valor || 0) || 0);
          }, 0),
          to: to || "",
          cc: (cc && cc.join) ? cc.join(",") : "",
          temViagemHosp: false,
          temCartaoBloqueado: packSafe.some(function (r) {
            return r && r.cartaoBloqueado === true;
          }),
          assunto: subject || "",
          status: "FAIL",
          error: msgErro
        });
      } catch (_) {}
    }
  });

  if (erros.length) {
    return { ok: true, enviados: enviados, lojas: lojaKeys.length, erros: erros, skipped: skipped };
  }

  return { ok: true, enviados: enviados, lojas: lojaKeys.length, skipped: skipped };
}

var VEKTOR_ITENS_IRREG_LOG_TAB = "Itens Irreg.";

function vektorGetOrCreateItensIrregLogSheet_() {
  var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
  var sh = ss.getSheetByName(VEKTOR_ITENS_IRREG_LOG_TAB);

  if (!sh) {
    sh = ss.insertSheet(VEKTOR_ITENS_IRREG_LOG_TAB);
    sh.appendRow([
      "sentAt","dataEnvioBR","lojaKey","lojaRaw","time","qtdItens","valorTotal",
      "to","cc","temViagemHosp","temCartaoBloqueado","assunto","status","error"
    ]);
    sh.getRange(1,1,1,14).setFontWeight("bold");
    sh.setFrozenRows(1);
  } else {
    var headers = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(function(x){
      return String(x || "").trim();
    });

    if (headers.indexOf("temCartaoBloqueado") === -1) {
      sh.insertColumnAfter(10);
      sh.getRange(1, 11).setValue("temCartaoBloqueado");
      sh.getRange(1,1,1,sh.getLastColumn()).setFontWeight("bold");
    }
  }
  return sh;
}

function vektorLogEnvioItensIrreg_(payload) {
  try {
    var sh = vektorGetOrCreateItensIrregLogSheet_();
    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    var now = new Date();

    sh.appendRow([
      Utilities.formatDate(now, tz, "yyyy-MM-dd HH:mm:ss"),
      Utilities.formatDate(now, tz, "dd/MM/yyyy"),
      payload.lojaKey || "",
      payload.lojaRaw || "",
      payload.time || "",
      Number(payload.qtdItens || 0),
      Number(payload.valorTotal || 0),
      payload.to || "",
      payload.cc || "",
      payload.temViagemHosp ? "SIM" : "NAO",
      payload.temCartaoBloqueado ? "SIM" : "NAO",
      payload.assunto || "",
      payload.status || "",
      payload.error || ""
    ]);
  } catch (e) {
    Logger.log("Falha ao logar envio de itens irregulares: " + (e && e.message ? e.message : e));
  }
}

// Procura o índice de uma coluna no cabeçalho da BaseClara
// usando uma lista de possíveis nomes (variações de texto).
function encontrarIndiceColuna_(header, nomesPossiveis) {
  // header: array de strings
  // nomesPossiveis: string OU array de strings

  if (!header || !header.length) return -1;

  // aceita string direta também
  var arr = Array.isArray(nomesPossiveis) ? nomesPossiveis : [nomesPossiveis];

  // normalizador defensivo (não depende de outras funcs)
  function norm_(s) {
    return String(s || "")
      .trim()
      .toLowerCase()
      .normalize("NFD")
      .replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ");
  }

  var headerNorm = header.map(norm_);

  // 1) match EXATO (melhor)
  for (var a = 0; a < arr.length; a++) {
    var alvo = norm_(arr[a]);
    if (!alvo) continue;
    for (var i = 0; i < headerNorm.length; i++) {
      if (headerNorm[i] === alvo) return i;
    }
  }

  // 2) match "contém" (fallback controlado)
  for (var b = 0; b < arr.length; b++) {
    var alvo2 = norm_(arr[b]);
    if (!alvo2) continue;
    for (var j = 0; j < headerNorm.length; j++) {
      if (headerNorm[j].indexOf(alvo2) !== -1) return j;
    }
  }

  return -1;
}

// Filtra linhas pelo período [dataInicioStr, dataFimStr].
// Se vier vazio, considera últimos 7 dias.
function filtrarLinhasPorPeriodo_(linhas, idxData, dataInicioStr, dataFimStr) {
  var hoje = new Date();
  var ini, fim;

  if (dataInicioStr) {
    ini = new Date(dataInicioStr);
  } else {
    ini = new Date(hoje);
    ini.setDate(hoje.getDate() - 30);
  }

  if (dataFimStr) {
    fim = new Date(dataFimStr);
  } else {
    fim = hoje;
  }

  var resultado = [];
  for (var i = 0; i < linhas.length; i++) {
    var row = linhas[i];
    var d = parseDateClara_(row[idxData]);
    if (!d) continue;
    if (d >= ini && d <= fim) {
      resultado.push(row);
    }
  }
  return resultado;
}

/**
 * Retorna os limites de ciclo (06 -> 05) para um "offset" de meses.
 * offsetMeses = 0 => ciclo atual
 * offsetMeses = 1 => ciclo anterior
 * offsetMeses = 2 => 2 ciclos atrás, etc.
 */
function getPeriodoCicloOffset_(offsetMeses) {
  var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
  var hoje = new Date();
  var y = hoje.getFullYear();
  var m = hoje.getMonth(); // 0..11

  // Se hoje ainda não chegou no dia 06, ciclo atual começou no mês anterior
  var cicloStartMonth = (hoje.getDate() >= 6) ? m : (m - 1);

  // Aplica offset (volta ciclos)
  cicloStartMonth = cicloStartMonth - (offsetMeses || 0);

  // Ajusta ano/mês
  var start = new Date(y, cicloStartMonth, 6, 0, 0, 0, 0);
  var end = new Date(y, cicloStartMonth + 1, 5, 23, 59, 59, 999);

  return { inicio: start, fim: end, tz: tz };
}

function getPendenciasResumoCicloAtual() {
  try {
    // ✅ Restrito a Administrador (não depende de VEKTOR_ACESSOS)
    var sess = (Session.getActiveUser().getEmail() || "").trim().toLowerCase();
    if (!isAdminEmail(sess)) {
      return { ok: false, restrito: true, error: "Não disponível para o seu perfil." };
    }

    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";

    // Período do ciclo atual (06 -> agora)
    var per = getPeriodoCicloOffset_(0); // {inicio,fim,tz}
    var ini = per && per.inicio ? per.inicio : null;
    if (!ini) return { ok: false, error: "Não consegui identificar o início do ciclo atual." };

    var fim = new Date();

    // BaseClara
    var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
    var sh = ss.getSheetByName("BaseClara");
    if (!sh) return { ok: false, error: "Aba BaseClara não encontrada." };

    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2) {
      return {
        ok: true,
        periodo: {
          inicio: Utilities.formatDate(ini, tz, "dd/MM/yyyy"),
          fim: Utilities.formatDate(fim, tz, "dd/MM/yyyy")
        },
        totais: { totalPendTrans: 0, pendEtiqueta: 0, pendDescricao: 0, pendRecibo: 0 },
        lojas: { total: 0, comPendencia: 0 },
        lojasComPendenciaLista: [],
        topLojas: []
      };
    }

    // Lê tudo (header + rows)
    var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
    var header = values[0].map(function (h) { return String(h || "").trim(); });
    var rows = values.slice(1);

    // helper: aliases exatos (igual você já vinha usando)
    function idxOf(possiveis) {
      for (var i = 0; i < possiveis.length; i++) {
        var p = possiveis[i];
        var ix = header.indexOf(p);
        if (ix >= 0) return ix;
      }
      return -1;
    }

    // helper local: match EXATO (evita "Recibo" bater em coluna errada)
    function findHeaderExactLocal_(headerArr, label) {
      var alvo = normalizarTexto_(label || "");
      for (var i = 0; i < headerArr.length; i++) {
        var h = normalizarTexto_(String(headerArr[i] || ""));
        if (h === alvo) return i;
      }
      return -1;
    }

    // ✅ índices principais
    var idxDataTrans  = idxOf(["Data da Transação", "Data Transação", "Data"]);
    var idxValorBRL   = idxOf(["Valor em R$", "Valor (R$)", "Valor"]);
    var idxLojaNum    = idxOf(["LojaNum", "Loja", "Código Loja", "cod_estbl", "cod_loja"]);

    if (idxDataTrans < 0) throw new Error("Não encontrei a coluna 'Data da Transação' na BaseClara.");
    if (idxValorBRL  < 0) throw new Error("Não encontrei a coluna 'Valor em R$' na BaseClara.");
    if (idxLojaNum   < 0) throw new Error("Não encontrei a coluna 'LojaNum' na BaseClara.");

    // ✅ índices de pendência (EXATO primeiro, depois fallback fixo)
    var idxRecibo = findHeaderExactLocal_(header, "Recibo");
    if (idxRecibo < 0) idxRecibo = encontrarIndiceColuna_(header, ["Recibo", "NF / Recibo", "NF/Recibo"]);
    if (idxRecibo < 0) idxRecibo = 14; // O (0-based)

    var idxEtiquetas = findHeaderExactLocal_(header, "Etiquetas");
    if (idxEtiquetas < 0) idxEtiquetas = findHeaderExactLocal_(header, "Etiqueta");
    if (idxEtiquetas < 0) idxEtiquetas = encontrarIndiceColuna_(header, ["Etiquetas", "Etiqueta"]);
    if (idxEtiquetas < 0) idxEtiquetas = 19; // T (0-based)

    var idxDescricao = findHeaderExactLocal_(header, "Descrição");
    if (idxDescricao < 0) idxDescricao = findHeaderExactLocal_(header, "Descricao");
    if (idxDescricao < 0) idxDescricao = encontrarIndiceColuna_(header, ["Descrição", "Descricao"]);
    if (idxDescricao < 0) idxDescricao = 20; // U (0-based)

    // (2) Mapa Loja->Time usando a sua regra oficial (col R=Time, V=LojaNum)
    var mapLojaTime = construirMapaLojaParaTime_();

    function parseNumberSafe_(v) {
      if (v === null || v === undefined || v === "") return 0;
      if (typeof v === "number") return v;
      var s = String(v).trim().replace(/\./g, "").replace(",", ".");
      var n = Number(s);
      return isFinite(n) ? n : 0;
    }

    function isVazio_(v) {
      if (v === null || v === undefined) return true;
      if (typeof v === "boolean") return (v === false); // checkbox
      var s = String(v).trim().toLowerCase();
      if (!s) return true;
      if (s === "-" || s === "—" || s === "n/a" || s === "na") return true;
      if (s === "false" || s === "0") return true;
      if (s === "não" || s === "nao") return true;
      if (s.indexOf("sem recibo") >= 0) return true;
      if (s.indexOf("sem etiqueta") >= 0) return true;
      return false;
    }

    // (3) Agregação
    var totPendTrans = 0;
    var totPEtiq = 0, totPDesc = 0, totPRec = 0;

    var mapaLojas = {}; // loja(4d) -> {loja,time,totalPendencias,valorPendente}

    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];

      // Data (filtra pelo ciclo atual: ini -> agora)
      var dt = row[idxDataTrans];
      var dtx = (dt instanceof Date) ? dt : new Date(dt);
      if (!(dtx instanceof Date) || isNaN(dtx.getTime())) continue;
      if (dtx < ini || dtx > fim) continue;

      // Loja
      var lojaNum = normalizarLojaNumero_(row[idxLojaNum]);
      if (!lojaNum) continue;
      var loja4 = String(lojaNum).padStart(4, "0");

      // Valor
      var valor = parseNumberSafe_(row[idxValorBRL]);

      // ✅ Campos de pendência (agora usando os índices corretos)
      var etiquetas = row[idxEtiquetas];
      var recibo = row[idxRecibo];
      var desc = row[idxDescricao];

      var temPendEtiqueta = isVazio_(etiquetas);
      var temPendRecibo   = isVazio_(recibo);
      var temPendDesc     = isVazio_(desc);

      var temPend = temPendEtiqueta || temPendRecibo || temPendDesc;
      if (!temPend) continue;

      totPendTrans++;

      if (temPendEtiqueta) totPEtiq++;
      if (temPendDesc)     totPDesc++;
      if (temPendRecibo)   totPRec++;

      if (!mapaLojas[loja4]) {
        mapaLojas[loja4] = {
          loja: loja4,
          time: mapLojaTime[Number(lojaNum)] || "—",
          totalPendencias: 0,
          valorPendente: 0
        };
      }

      mapaLojas[loja4].totalPendencias++;
      mapaLojas[loja4].valorPendente += valor;
    }

    // lista explícita (para export)
    var lojasComPendenciaLista = Object.keys(mapaLojas || {}).sort();

    // Métrica de abrangência por loja
    var totalLojasAtivas = Object.keys(mapLojaTime || {}).length;
    var lojasComPendencia = lojasComPendenciaLista.length;

    // Top lojas
    var topLojas = Object.keys(mapaLojas).map(function (k) { return mapaLojas[k]; });

    // ✅ ordena: MAIOR VOLUME de pendências, depois MAIOR valor pendente
      topLojas.sort(function(a,b){
        if (b.totalPendencias !== a.totalPendencias) return b.totalPendencias - a.totalPendencias;
        return b.valorPendente - a.valorPendente;
      });

    return {
      ok: true,
      periodo: {
        inicio: Utilities.formatDate(ini, tz, "dd/MM/yyyy"),
        fim: Utilities.formatDate(fim, tz, "dd/MM/yyyy")
      },
      totais: {
        totalPendTrans: totPendTrans,
        pendEtiqueta: totPEtiq,
        pendDescricao: totPDesc,
        pendRecibo: totPRec
      },
      lojas: {
        total: totalLojasAtivas,
        comPendencia: lojasComPendencia
      },
      lojasComPendenciaLista: lojasComPendenciaLista,
      topLojas: topLojas.slice(0, 10)
    };

  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

/**
 * Projeção de gasto por loja para o ciclo atual (06->05), usando sazonalidade:
 * - Base: média dos últimos 6 ciclos completos
 * - Sazonal (Nov/Dez): usa o MAIOR ciclo dos últimos 6 (conservador para evitar estouro)
 * - Fallback: se não tiver ciclos suficientes, usa os disponíveis; em último caso últimos 30 dias (projetado para um ciclo)
 *
 * Retorna:
 *  {
 *    proj: { "0287": 5400.25, ... },
 *    meta: { "0287": { fonte:"media6|max6|mediaN|ult30", nCiclos:6 }, ... }
 *  }
 */
function calcularProjecaoPorLojaUltimosCiclos_(linhas, idxData, idxValor, idxLoja) {

  function somaPorPeriodo(inicio, fim) {
    var soma = {};
    for (var i = 0; i < linhas.length; i++) {
      var row = linhas[i];
      if (!row) continue;

      var d = parseDateClara_(row[idxData]);
      if (!d || d < inicio || d > fim) continue;

      var lojaRaw = (row[idxLoja] || "").toString().trim();
      var dig = lojaRaw.replace(/\D/g, "");
      if (!dig) continue;

      var loja = String(Number(dig)).padStart(4, "0");
      var v = parseNumberSafe_(row[idxValor]);
      if (!v) continue;

      soma[loja] = (soma[loja] || 0) + v;
    }
    return soma;
  }

  function getCicloLenDias_(inicio, fim) {
    return Math.max(1, Math.round((fim.getTime() - inicio.getTime()) / (1000 * 60 * 60 * 24)) + 1);
  }

  // --- 6 ciclos completos anteriores (1..6) ---
  var ciclos = []; // [{ini,fim,somaPorLoja}]
  for (var c = 1; c <= 6; c++) {
    var per = getPeriodoCicloOffset_(c);
    ciclos.push({
      ini: per.inicio,
      fim: per.fim,
      soma: somaPorPeriodo(per.inicio, per.fim)
    });
  }

  // --- Últimos 30 dias corridos (fallback final) ---
  var hoje = new Date();
  var ini30 = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate() - 30, 0, 0, 0, 0);
  var soma30 = somaPorPeriodo(ini30, hoje);

  // ciclo atual para “projetar” 30 dias -> ciclo
  var perAtual = getPeriodoCicloOffset_(0);
  var diasCiclo = getCicloLenDias_(perAtual.inicio, new Date(perAtual.inicio.getFullYear(), perAtual.inicio.getMonth() + 1, 5, 23, 59, 59, 999));
  var fator30 = diasCiclo / 30;

  // lojas universo
  var lojas = {};
  ciclos.forEach(function(cy) {
    Object.keys(cy.soma).forEach(function(loja) { lojas[loja] = true; });
  });
  Object.keys(soma30).forEach(function(loja) { lojas[loja] = true; });

  var proj = {};
  var meta = {};

  // sazonalidade: se ciclo atual cai em novembro/dezembro (mês do início do ciclo)
  var mesInicio = perAtual.inicio.getMonth() + 1; // 1..12
  var sazonal = (mesInicio === 11 || mesInicio === 12);

  Object.keys(lojas).forEach(function(loja) {
    // coleta totais dos ciclos em que a loja apareceu
    var vals = [];
    for (var i = 0; i < ciclos.length; i++) {
      var v = ciclos[i].soma[loja];
      if (v != null) vals.push(v);
    }

    if (vals.length >= 6) {
      var soma = vals.reduce(function(a,b){return a+b;}, 0);
      var media6 = soma / 6;
      var max6 = Math.max.apply(null, vals);

      proj[loja] = sazonal ? max6 : media6;
      meta[loja] = { fonte: sazonal ? "max6" : "media6", nCiclos: 6 };

    } else if (vals.length >= 1) {
      // se tem 1..5 ciclos: usa média do que tiver
      var somaN = vals.reduce(function(a,b){return a+b;}, 0);
      var mediaN = somaN / vals.length;

      // em sazonal, ainda pode usar o máximo do que tiver (evita estouro)
      var maxN = Math.max.apply(null, vals);
      proj[loja] = sazonal ? maxN : mediaN;
      meta[loja] = { fonte: sazonal ? "maxN" : "mediaN", nCiclos: vals.length };

    } else {
      // fallback final: últimos 30 dias projetados para um ciclo
      var v30 = soma30[loja] || 0;
      proj[loja] = v30 * fator30;
      meta[loja] = { fonte: "ult30", nCiclos: 0 };
    }
  });

  return { proj: proj, meta: meta };
}

/**
 * Retorna, para um determinado time/grupo (ou geral se grupo vazio), um resumo de transações por loja:
 * - total de transações
 * - valor total em R$
 *
 * criterio:
 *   "quantidade" -> ordena pelo número de transações
 *   "valor"      -> ordena pelo valor total em R$
 *
 * É chamado pelo front via google.script.run.getResumoTransacoesPorGrupo(...)
 */
function getResumoTransacoesPorGrupo(grupo, dataInicioStr, dataFimStr, criterio) {
  vektorAssertFunctionAllowed_("getResumoTransacoesPorGrupo");
  var info = carregarLinhasBaseClara_();
  if (info.error) {
    return { ok: false, error: info.error };
  }

  // guarda o nome original (com acento/maiúsculas) pra exibir no chat
  var grupoOriginal = (grupo || "").toString().trim();
  // versão normalizada (sem acento, minúscula) para filtrar
  var grupoNorm = normalizarTexto_(grupoOriginal);

  // normaliza critério
  criterio = (criterio || "").toString().toLowerCase();
  if (criterio !== "valor" && criterio !== "quantidade") {
    // se vier vazio ou algo diferente, usa "quantidade" por padrão
    criterio = "quantidade";
  }

  var linhas = info.linhas;

  // Índices das colunas na BaseClara (começando em 0)
  // A: Data da Transação
  // F: Valor em R$
  // R: Grupos
  // V: LojaNum
  var IDX_DATA  = 0;   // "Data da Transação"
  var IDX_VALOR = 5;   // "Valor em R$"
  var IDX_GRUPO = 17;  // "Grupos"
  var IDX_LOJA  = 21;  // "LojaNum"

  var filtradas = filtrarLinhasPorPeriodo_(linhas, IDX_DATA, dataInicioStr, dataFimStr);

  var mapa = {};
  for (var i = 0; i < filtradas.length; i++) {
    var row = filtradas[i];

        // 🔹 FALTOU ESTA LINHA:
    var loja = (row[IDX_LOJA] || "").toString().trim();

    // valor de grupo na linha da planilha
    var grupoLinhaOriginal = (row[IDX_GRUPO] || "").toString();
    var grupoLinhaNorm = normalizarTexto_(grupoLinhaOriginal);

    // se o usuário informou um grupo/time no chat, aplica filtro
    if (grupoNorm) {
      // regra flexível:
      // - se a linha contiver o grupo completo (ex: "aguias do cerrado")
      //   OU
      // - se o grupo informado contiver o valor da linha (ex: "lobos sp" contém "lobos")
      var casaGrupo =
        grupoLinhaNorm.indexOf(grupoNorm) !== -1 ||
        grupoNorm.indexOf(grupoLinhaNorm) !== -1;

      if (!casaGrupo) {
        continue;
      }
    }

    if (!loja) continue;

    if (!mapa[loja]) {
      mapa[loja] = { loja: loja, total: 0, valorTotal: 0 };
    }
    mapa[loja].total++;
    var valor = Number(row[IDX_VALOR]) || 0;
    mapa[loja].valorTotal += valor;
  }

  var lojasArr = [];
  for (var key in mapa) {
    if (Object.prototype.hasOwnProperty.call(mapa, key)) {
      lojasArr.push(mapa[key]);
    }
  }

  // ordenação conforme critério
  if (criterio === "valor") {
    lojasArr.sort(function (a, b) {
      if (b.valorTotal !== a.valorTotal) {
        return b.valorTotal - a.valorTotal;
      }
      return b.total - a.total; // desempate por quantidade
    });
  } else {
    // "quantidade"
    lojasArr.sort(function (a, b) {
      if (b.total !== a.total) {
        return b.total - a.total;
      }
      return b.valorTotal - a.valorTotal; // desempate por valor
    });
  }

  var top = lojasArr.length ? lojasArr[0] : null;

  return {
    ok: true,
    grupoOriginal: grupoOriginal,
    grupo: grupo,
    criterio: criterio,
    lojas: lojasArr,
    top: top
  };
}

/**
 * Frequência de uso Clara (por TIME ou por LOJA), com período configurável.
 *
 * @param {"time"|"loja"} tipoFiltro
 * @param {string} valorFiltro
 * @param {number} mesesBack  // ex.: 1 = mês corrente, 3 = últimos 3 meses, 6 = último semestre
 */
function getFrequenciaUsoClara(tipoFiltro, valorFiltro, mesesBack) {
  vektorAssertFunctionAllowed_("getFrequenciaUsoClara");
  try {
    tipoFiltro = (tipoFiltro || "").toString().toLowerCase().trim();
    valorFiltro = (valorFiltro || "").toString().trim();
    mesesBack = Number(mesesBack) || 1;

    var info = carregarLinhasBaseClara_();
    if (info.error) return { ok: false, error: info.error };

    var header = info.header;
    var linhas = info.linhas;

    // Índices
    var idxData  = encontrarIndiceColuna_(header, ["Data da Transação", "Data Transação", "Data"]);
    var idxLoja  = encontrarIndiceColuna_(header, ["LojaNum", "Loja Num", "Loja", "Loja Número", "Loja Numero"]);
    var idxGrupo = encontrarIndiceColuna_(header, ["Grupos", "Grupo", "Time"]);
    var idxValor = encontrarIndiceColuna_(header, ["Valor em R$", "Valor (R$)", "Valor"]);

    if (idxData < 0 || idxLoja < 0 || idxGrupo < 0 || idxValor < 0) {
      return { ok: false, error: "Não encontrei colunas necessárias (Data / Loja / Grupo / Valor) na BaseClara." };
    }

    // ---------- Período analisado ----------
    // Regra:
    // - Se mesesBack >= 2: mantém lógica atual (meses calendário, incluindo mês corrente)
    // - Se mesesBack = 1 (padrão): usa últimos 30 dias (janela móvel)
    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    var hoje = new Date();

    var inicioPeriodo, fimPeriodo;

    // fim sempre é "agora" (fim do dia de hoje)
    fimPeriodo = new Date(hoje);
    fimPeriodo.setHours(23, 59, 59, 999);

    if (mesesBack >= 2) {
      // ✅ mantém a lógica atual por meses calendário
      inicioPeriodo = new Date(hoje.getFullYear(), hoje.getMonth() - (mesesBack - 1), 1);
      inicioPeriodo.setHours(0, 0, 0, 0);

      // fim do mês corrente (como era antes)
      fimPeriodo = new Date(hoje.getFullYear(), hoje.getMonth() + 1, 0);
      fimPeriodo.setHours(23, 59, 59, 999);
    } else {
      // ✅ default: últimos 30 dias (janela móvel)
      inicioPeriodo = new Date(hoje);
      inicioPeriodo.setDate(inicioPeriodo.getDate() - 29);
      inicioPeriodo.setHours(0, 0, 0, 0);
    }

    // ---------- Semana corrente (Seg–Dom) ----------
    var dow = hoje.getDay(); // 0=Dom
    var diffToMonday = (dow === 0) ? -6 : (1 - dow);
    var inicioSemana = new Date(hoje);
    inicioSemana.setDate(hoje.getDate() + diffToMonday);
    inicioSemana.setHours(0, 0, 0, 0);

    var fimSemana = new Date(inicioSemana);
    fimSemana.setDate(inicioSemana.getDate() + 6);
    fimSemana.setHours(23, 59, 59, 999);

    // ---------- Consistência (últimos 6 meses, incluindo o mês corrente) ----------
    var inicioConsistencia = new Date(hoje.getFullYear(), hoje.getMonth() - 5, 1);
    inicioConsistencia.setHours(0, 0, 0, 0);

    // ---------- Helpers ----------
    function toDateKey(d) { return Utilities.formatDate(d, tz, "yyyy-MM-dd"); }
    function toMonthKey(d) { return Utilities.formatDate(d, tz, "yyyy-MM"); }
    function countSet(obj) { return Object.keys(obj || {}).length; }

    function grupoMatchTime_(grupoLinhaRaw, filtroTimeNorm) {
      if (!grupoLinhaRaw) return false;

      // normaliza texto do campo "Grupos"
      var g = normalizarTexto_(grupoLinhaRaw.toString());
      if (!g) return false;

      // separadores comuns em "Grupos": | ; , / -
      var partes = g.split(/[\|\;\,\/\-]+/)
        .map(function(s){ return s.trim(); })
        .filter(function(s){ return !!s; });

      // 1) tenta match exato por parte
      for (var i = 0; i < partes.length; i++) {
        if (partes[i] === filtroTimeNorm) return true;
      }

      // 2) fallback: contém (apenas um lado, mais previsível)
      return g.indexOf(filtroTimeNorm) !== -1;
    }

    function mediaIntervaloDias(datas) {
      if (!datas || datas.length < 2) return null;
      datas.sort(function(a,b){ return a - b; });

      var gaps = [];
      for (var j = 1; j < datas.length; j++) {
        var diffDias = (datas[j] - datas[j - 1]) / 86400000;
        if (diffDias >= 0) gaps.push(diffDias);
      }
      if (!gaps.length) return null;

      var soma = gaps.reduce(function(acc,v){ return acc + v; }, 0);
      return soma / gaps.length;
    }

    function classificarPadrao(diasNoPeriodo) {
      if (diasNoPeriodo >= 20) return "Uso diário";
      if (diasNoPeriodo >= 10) return "Uso frequente";
      if (diasNoPeriodo >= 4)  return "Uso moderado";
      if (diasNoPeriodo >= 1)  return "Uso esporádico";
      return "Sem uso";
    }

    function rotuloCadencia(intervaloMedio) {
      if (intervaloMedio === null) return "—";
      if (intervaloMedio <= 1.2) return "Diariamente";
      if (intervaloMedio <= 2.2) return "De 2 em 2 dias";
      if (intervaloMedio <= 3.2) return "De 3 em 3 dias";
      if (intervaloMedio <= 7.5) return "Semanalmente";
      if (intervaloMedio <= 15)  return "Quinzenal";
      return "Mensal / esporádico";
    }

    function calcConsistencia(mesesObj) {
      var meses = Object.keys(mesesObj || {}).sort();
      if (meses.length < 2) return "Sem histórico";

      var ult = meses.slice(-6);
      var serie = ult.map(function(mk){ return countSet(mesesObj[mk]); });

      var n = serie.length;
      var mean = serie.reduce(function(a,b){ return a + b; }, 0) / n;

      var varSum = 0;
      for (var i = 0; i < n; i++) varSum += Math.pow(serie[i] - mean, 2);
      var sd = Math.sqrt(varSum / n);

      var delta = serie[serie.length - 1] - serie[0];

      if (sd <= 2 && Math.abs(delta) <= 2) return "Estável";
      if (delta >= 3) return "Crescendo";
      if (delta <= -3) return "Caindo";
      return "Oscilando";
    }

    function fmtMoedaBR_(v) {
      var n = Number(v) || 0;
      try {
        return n.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
      } catch (e) {
        // fallback seguro
        var s = (Math.round(n * 100) / 100).toFixed(2).replace(".", ",");
        return "R$ " + s;
      }
    }

    // ---------- Time atual por loja (janela fixa: últimos 30 dias) ----------
    var janelaDiasTimeAtual = 30;

    // ✅ time atual baseado no FIM do período analisado (não em "hoje")
    var inicioTimeAtual = new Date(fimPeriodo);
    inicioTimeAtual.setDate(inicioTimeAtual.getDate() - (janelaDiasTimeAtual - 1));
    inicioTimeAtual.setHours(0, 0, 0, 0);

    // lojaNorm -> { timeRaw, timeNorm, dataMaisRecente }
    var timeAtualPorLoja = {};

    // Só faz sentido para filtro por TIME
    if (filtroTimeNorm) {
      for (var t = 0; t < linhas.length; t++) {
        var rowTA = linhas[t];
        if (!rowTA) continue;

        var dTA = parseDateClara_(rowTA[idxData]);
        if (!dTA || isNaN(dTA.getTime())) continue;

        // considera apenas últimos 30 dias
        if (dTA < inicioTimeAtual || dTA > fimPeriodo) continue;

        var lojaRawTA = (rowTA[idxLoja] || "").toString();
        var lojaDigitsTA = lojaRawTA.replace(/\D/g, "");
        if (!lojaDigitsTA) continue;

        var lojaNormTA = ("0000" + lojaDigitsTA).slice(-4);

        var grupoRawTA = (rowTA[idxGrupo] || "").toString();
        if (!grupoRawTA) continue;

        var grupoNormTA = normalizarTexto_(grupoRawTA);
        if (!grupoNormTA) continue;

        // guarda o grupo mais recente nessa janela
        if (!timeAtualPorLoja[lojaNormTA] || dTA > timeAtualPorLoja[lojaNormTA].dataMaisRecente) {
          timeAtualPorLoja[lojaNormTA] = {
            timeRaw: grupoRawTA,
            timeNorm: grupoNormTA,
            dataMaisRecente: new Date(dTA)
          };
        }
      }
    }

    // ---------- Varredura / agregação ----------
    var mapa = {}; // lojaNorm -> stats

    for (var i = 0; i < linhas.length; i++) {
      var row = linhas[i];
      if (!row) continue;

      // Data
      var d = parseDateClara_(row[idxData]);
      if (!d || isNaN(d.getTime())) continue;

      // Loja
      var lojaRaw = (row[idxLoja] || "").toString();
      var lojaDigits = lojaRaw.replace(/\D/g, "");
      if (!lojaDigits) continue;

      var lojaNum = Number(lojaDigits);
      if (!isFinite(lojaNum)) continue;

      var lojaNorm = ("0000" + lojaDigits).slice(-4);

      // Time (para filtro) — usa TIME ATUAL da loja (últimos 30 dias)
      if (filtroTimeNorm) {
        var ta = timeAtualPorLoja[lojaNorm];

        if (!ta || !ta.timeRaw) {
          // ✅ OPÇÃO A: exclui do relatório por time (mais correto para sua regra)
          continue;
        }

        // compara o time atual da loja com o time solicitado
        if (!grupoMatchTime_(ta.timeRaw, filtroTimeNorm)) continue;
      }

      // filtro por loja
      if (filtroLojaNum !== null && lojaNum !== filtroLojaNum) continue;

      var valor = Number(row[idxValor]) || 0;
      var keyDia = toDateKey(d);

      if (!mapa[lojaNorm]) {
        mapa[lojaNorm] = {
          loja: lojaNorm,
          diasSemanaSet: {},
          diasPeriodoSet: {},
          diasTotalSet: {},
          datasTotal: [],
          topValorPeriodo: 0,
          ultimaDataPeriodo: null,
          picoPeriodoPorDia: {},
          meses: {}
        };
      }

      var obj = mapa[lojaNorm];

      // Total
      obj.diasTotalSet[keyDia] = true;
      obj.datasTotal.push(new Date(d));

      // Semana
      if (d >= inicioSemana && d <= fimSemana) {
        obj.diasSemanaSet[keyDia] = true;
      }

      // Período analisado
      if (d >= inicioPeriodo && d <= fimPeriodo) {
        obj.diasPeriodoSet[keyDia] = true;

        if (valor > obj.topValorPeriodo) obj.topValorPeriodo = valor;
        if (!obj.ultimaDataPeriodo || d > obj.ultimaDataPeriodo) obj.ultimaDataPeriodo = new Date(d);

        if (!obj.picoPeriodoPorDia[keyDia]) obj.picoPeriodoPorDia[keyDia] = { qtd: 0, valor: 0 };
        obj.picoPeriodoPorDia[keyDia].qtd += 1;
        obj.picoPeriodoPorDia[keyDia].valor += valor;
      }

      // Consistência
      if (d >= inicioConsistencia && d <= fimPeriodo) {
        var mk = toMonthKey(d);
        if (!obj.meses[mk]) obj.meses[mk] = {};
        obj.meses[mk][keyDia] = true;
      }
    }

    // ---------- Monta linhas ----------
    var rows = [];

    Object.keys(mapa).forEach(function(loja){
      var it = mapa[loja];

      var diasSem = countSet(it.diasSemanaSet);
      var diasPer = countSet(it.diasPeriodoSet);
      var diasTot = countSet(it.diasTotalSet);

      var intervaloMedio = mediaIntervaloDias(it.datasTotal);
      var padrao = classificarPadrao(diasPer);

      var freqUsoSem = rotuloCadencia(diasSem >= 2 ? (6 / Math.max(1, diasSem - 1)) : null);
      var freqUsoPer = rotuloCadencia(diasPer >= 2 ? ((mesesBack * 30) / Math.max(1, diasPer - 1)) : null);
      var freqUsoTot = rotuloCadencia(intervaloMedio);

      var ultimaDataPeriodoFmt = it.ultimaDataPeriodo
        ? Utilities.formatDate(it.ultimaDataPeriodo, tz, "dd/MM/yyyy")
        : "—";

      rows.push({
        loja: it.loja,
        freqUsoSem: freqUsoSem,
        freqUsoMes: freqUsoPer,
        freqUsoTotal: freqUsoTot,
        freqDias: diasPer,
        intervaloMedio: (intervaloMedio === null ? null : intervaloMedio),
        padrao: padrao,
        freqXValor: (padrao.indexOf("Uso") >= 0
          ? (diasPer >= 10 ? "Alta freq" : "Baixa freq") + " + " + (it.topValorPeriodo >= 1000 ? "alto valor" : "baixo valor")
          : "Sem uso"),
        topValor: it.topValorPeriodo || 0,
        ultimaDataTrans: ultimaDataPeriodoFmt,
        consistencia: calcConsistencia(it.meses),

        __diasPer: diasPer,
        __topValor: it.topValorPeriodo || 0,
        __picoDia: it.picoPeriodoPorDia
      });
    });

    rows.sort(function(a,b){
      if ((b.__diasPer||0) !== (a.__diasPer||0)) return (b.__diasPer||0) - (a.__diasPer||0);
      return (b.__topValor||0) - (a.__topValor||0);
    });

    // ---------- Insight principal ----------
    var insight = "";
    if (rows.length) {
      var top = rows[0];

      var picoData = "";
      var picoQtd = 0;
      var picoValor = 0;

      var m = top.__picoDia || {};
      Object.keys(m).forEach(function(dk){
        var x = m[dk] || {qtd:0, valor:0};
        if (x.qtd > picoQtd || (x.qtd === picoQtd && x.valor > picoValor)) {
          picoQtd = x.qtd;
          picoValor = x.valor;
          picoData = dk;
        }
      });

      var dataFmt = picoData ? (picoData.split("-")[2] + "/" + picoData.split("-")[1] + "/" + picoData.split("-")[0]) : "—";

      insight =
        "Maior impacto no período: loja <b>" + top.loja + "</b> (" + top.__diasPer + " dias distintos com uso). " +
        (picoData ? ("Pico de uso em <b>" + dataFmt + "</b> (" + picoQtd + " transações no dia).") : "");
    }

    // ---------- Novos insights: Limite / Atenção ----------
    function scoreAumentar(r) {
      var s = 0;
      if ((r.freqDias || 0) >= 12) s += 4;
      else if ((r.freqDias || 0) >= 8) s += 3;
      else if ((r.freqDias || 0) >= 4) s += 1;

      if (r.intervaloMedio !== null && r.intervaloMedio !== undefined) {
        if (r.intervaloMedio <= 2) s += 3;
        else if (r.intervaloMedio <= 4) s += 2;
        else if (r.intervaloMedio <= 7) s += 1;
      }

      if (r.padrao === "Uso diário") s += 3;
      else if (r.padrao === "Uso frequente") s += 2;
      else if (r.padrao === "Uso moderado") s += 1;

      if (r.consistencia === "Crescendo") s += 2;
      else if (r.consistencia === "Estável") s += 1;

      if ((r.freqXValor || "").indexOf("Alta freq") >= 0) s += 2;
      if ((r.freqXValor || "").indexOf("alto valor") >= 0) s += 1;

      if ((r.topValor || 0) >= 2000) s += 2;
      else if ((r.topValor || 0) >= 1000) s += 1;

      return s;
    }

    function scoreReduzir(r) {
      var s = 0;
      if ((r.freqDias || 0) <= 1) s += 4;
      else if ((r.freqDias || 0) <= 2) s += 3;
      else if ((r.freqDias || 0) <= 3) s += 1;

      if (r.intervaloMedio !== null && r.intervaloMedio !== undefined) {
        if (r.intervaloMedio >= 15) s += 3;
        else if (r.intervaloMedio >= 10) s += 2;
        else if (r.intervaloMedio >= 7) s += 1;
      }

      if (r.padrao === "Sem uso") s += 3;
      else if (r.padrao === "Uso esporádico") s += 2;

      if (r.consistencia === "Caindo") s += 2;
      else if (r.consistencia === "Oscilando") s += 1;

      if ((r.freqXValor || "").indexOf("Baixa freq") >= 0 && (r.freqXValor || "").indexOf("alto valor") >= 0) {
        s -= 1; // não reduzir automaticamente em caso de alto valor com baixa freq (vira atenção)
      }

      return s;
    }

    function isAtencao(r) {
      var fxv = (r.freqXValor || "");
      var topv = (r.topValor || 0);

      if (r.consistencia === "Oscilando" && topv >= 1000) return true;
      if (fxv.indexOf("Baixa freq") >= 0 && fxv.indexOf("alto valor") >= 0) return true;
      if ((r.freqDias || 0) >= 8 && r.consistencia === "Caindo") return true;
      return false;
    }

    var bestInc = null, bestIncScore = -999;
    var bestDec = null, bestDecScore = -999;
    var listaAtencao = [];

    rows.forEach(function(r){
      var si = scoreAumentar(r);
      var sd = scoreReduzir(r);

      if (si > bestIncScore) { bestIncScore = si; bestInc = r; }
      if (sd > bestDecScore) { bestDecScore = sd; bestDec = r; }

      if (isAtencao(r)) listaAtencao.push(r);
    });

    listaAtencao.sort(function(a,b){
      if ((b.topValor||0) !== (a.topValor||0)) return (b.topValor||0) - (a.topValor||0);
      return (b.freqDias||0) - (a.freqDias||0);
    });
    listaAtencao = listaAtencao.slice(0, 3);

    var insightLimite = [];
    if (bestInc && bestIncScore >= 6) {
      insightLimite.push(
        "Sugestão de <b>aumento de limite</b>: loja <b>" + bestInc.loja + "</b> — uso recorrente (<b>" + (bestInc.freqDias||0) + " dias</b>), padrão <b>" + (bestInc.padrao||"—") + "</b>, consistência <b>" + (bestInc.consistencia||"—") + "</b>."
      );
    }
    if (bestDec && bestDecScore >= 6) {
      var im = (bestDec.intervaloMedio === null || bestDec.intervaloMedio === undefined) ? "—" : (Math.round(bestDec.intervaloMedio * 10) / 10).toString().replace(".", ",");
      insightLimite.push(
        "Sugestão de <b>redução de limite</b>: loja <b>" + bestDec.loja + "</b> — baixa recorrência (<b>" + (bestDec.freqDias||0) + " dias</b>), padrão <b>" + (bestDec.padrao||"—") + "</b>, intervalo médio <b>" + im + "</b> dias."
      );
    }

    // ✅ AQUI ESTÁ A CORREÇÃO: 1 linha por loja, com motivo claro
    var insightAtencao = [];
    if (listaAtencao.length) {
      listaAtencao.forEach(function(r){
        var tvFmt = fmtMoedaBR_(r.topValor || 0);
        var motivo = "";

        if ((r.freqXValor || "").indexOf("Baixa freq") >= 0 && (r.freqXValor || "").indexOf("alto valor") >= 0) {
          motivo = "Alto valor pontual com baixa frequência (risco de compra fora do padrão)";
        } else if (r.consistencia === "Oscilando" && (r.topValor || 0) >= 1000) {
          motivo = "Tendência de uso <b>oscilante</b> nos últimos meses, com transação de alto valor no período, validar se é sazonalidade ou mudança operacional”.";
        } else if ((r.freqDias || 0) >= 8 && r.consistencia === "Caindo") {
          motivo = "Queda recente de uso (pode indicar mudança operacional ou desnecessidade de limite atual)";
        } else {
          motivo = "Padrão de uso que merece acompanhamento";
        }

        insightAtencao.push(
          "• Loja <b>" + r.loja + "</b>: " + motivo +
          ". Padrão <b>" + (r.padrao||"—") + "</b>, consistência <b>" + (r.consistencia||"—") +
          "</b>, Top Valor <b>" + tvFmt + "</b>."
        );
      });
    }

    // limpa internos
    rows.forEach(function(r){
      delete r.__diasPer;
      delete r.__topValor;
      delete r.__picoDia;
    });

    return {
      ok: true,
      tipoFiltro: tipoFiltro,
      valorFiltro: valorFiltro,
      mesesBack: mesesBack,
      periodo: {
        inicio: Utilities.formatDate(inicioPeriodo, tz, "dd/MM/yyyy"),
        fim: Utilities.formatDate(fimPeriodo, tz, "dd/MM/yyyy")
      },
      insight: insight,
      insightLimite: insightLimite,
      insightAtencao: insightAtencao,
      rows: rows
    };

  } catch (e) {
    return { ok: false, error: e && e.message ? e.message : String(e) };
  }
}

/**
 * Lista de itens comprados na Clara (coluna "Descrição"), com data/valor/loja e análise de conformidade.
 *
 * @param {"time"|"loja"} tipoFiltro
 * @param {string} valorFiltro
 * @param {string|Date} dataIni  (dd/MM/yyyy ou Date)
 * @param {string|Date} dataFim  (dd/MM/yyyy ou Date)
 * @param {number} janelaDiasReincidencia (default 15)
 */
function getListaItensCompradosClara(tipoFiltro, valorFiltro, dataIni, dataFim, janelaDiasReincidencia) {
  vektorAssertFunctionAllowed_("getListaItensCompradosClara");
  try {
    tipoFiltro = (tipoFiltro || "").toString().toLowerCase().trim(); // "time" | "loja"
    valorFiltro = (valorFiltro || "").toString().trim();
    janelaDiasReincidencia = Number(janelaDiasReincidencia) || 15;

    var info = carregarLinhasBaseClara_();
    if (info.error) return { ok: false, error: info.error };

    var header = info.header;
    var linhas = info.linhas;

    // Índices (seguindo padrão do projeto)
    var idxData = encontrarIndiceColuna_(header, ["Data da Transação", "Data Transação", "Data"]);
    var idxValor = encontrarIndiceColuna_(header, ["Valor em R$", "Valor (R$)", "Valor"]);
    var idxLojaNum = encontrarIndiceColuna_(header, ["LojaNum", "Loja Num", "Loja", "Loja Número", "Loja Numero"]);
    var idxGrupo = encontrarIndiceColuna_(header, ["Grupos", "Grupo", "Time"]);
    // ✅ Alias do Cartão (H) — match estrito para não cair em "Cartão" (G)
    var idxAlias = -1;
    for (var iA = 0; iA < header.length; iA++) {
      var hn = normalizarTexto_((header[iA] || "").toString());
      if (hn === "alias do cartao" || hn === "alias do cartão") { idxAlias = iA; break; }
    }
    if (idxAlias < 0) {
      for (var jA = 0; jA < header.length; jA++) {
        var hn2 = normalizarTexto_((header[jA] || "").toString());
        if (hn2.indexOf("alias") !== -1 && hn2.indexOf("cartao") !== -1) { idxAlias = jA; break; }
      }
    }
    var idxDescricao = encontrarIndiceColuna_(header, ["Descrição", "Descricao", "Item", "Histórico", "Historico"]);
    var idxTransacao = 2;

    if (idxData < 0 || idxValor < 0 || idxDescricao < 0) {
      return { ok: false, error: "Não encontrei colunas necessárias (Data / Valor / Descrição) na BaseClara." };
    }

    // Se não tiver alias, a gente ainda consegue entregar com LojaNum
    // Mas se não tiver lojaNum nem alias, não dá para “por loja”
    if (tipoFiltro === "loja" && idxLojaNum < 0 && idxAlias < 0) {
      return { ok: false, error: "Não encontrei colunas de Loja (LojaNum/Alias) na BaseClara para filtrar por loja." };
    }

    // Se for por time e não tiver grupo/time, não dá
    if (tipoFiltro === "time" && idxGrupo < 0) {
      return { ok: false, error: "Não encontrei coluna de Time/Grupo (Grupos/Grupo/Time) na BaseClara para filtrar por time." };
    }

    // Período
    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    var dtIni = parseDateClara_(dataIni);
    var dtFim = parseDateClara_(dataFim);

    // fallback: se vier ISO string do front, tenta new Date
    if (!dtIni && typeof dataIni === "string") {
      var tmp = new Date(dataIni);
      if (!isNaN(tmp.getTime())) dtIni = tmp;
    }
    if (!dtFim && typeof dataFim === "string") {
      var tmp2 = new Date(dataFim);
      if (!isNaN(tmp2.getTime())) dtFim = tmp2;
    }

    if (!dtIni || !dtFim) {
      return { ok: false, error: "Período inválido. Informe data inicial e final (dd/MM/aaaa)." };
    }

    // Normalizadores
    function grupoMatchTime_(grupoLinhaRaw, filtroTimeNorm) {
      if (!grupoLinhaRaw) return false;
      var g = normalizarTexto_(grupoLinhaRaw.toString());
      if (!g) return false;

      var partes = g.split(/[\|\;\,\/\-]+/)
        .map(function(s){ return s.trim(); })
        .filter(function(s){ return !!s; });

      for (var i = 0; i < partes.length; i++) {
        if (partes[i] === filtroTimeNorm) return true;
      }
      return g.indexOf(filtroTimeNorm) !== -1;
    }

    function normalizarItem_(txt) {
      var n = normalizarTexto_(txt || "");
      // remove pontuação básica
      n = n.replace(/[^\p{L}\p{N}\s]/gu, " ");
      n = n.replace(/\s+/g, " ").trim();
      // remove stopwords comuns (mantém termos relevantes)
      n = (" " + n + " ")
        .replace(/ (de|da|do|das|dos|para|pra|com|sem|um|uma|uns|umas|ao|aos|na|no|nas|nos|e) /g, " ")
        .replace(/\s+/g, " ")
        .trim();
      return n;
    }

    // Classificação por política (heurística conservadora)
    
    function classificarPolitica_(descricaoNorm) {
  var d = (descricaoNorm || "").trim();

  // Palavras-chave “permitidas prováveis” (operacionais recorrentes)
  // Obs.: como d já está normalizado (sem acento), use sempre sem acento.
  var permitidosProv = [
    // Comunicação / gráfica / sinalização
    "impressao", "imprimir", "grafica", "plotagem", "encadernacao",
    "banner", "placa", "adesivo", "folder", "panfleto",
    "comunicacao", "comunicacao loja", "sinalizacao", "placas", "cartaz", "cartazes",
    "papel couche", "couche", "laminacao", "recorte", "vinil", "bobina", "impressão",
    // Água / consumo básico
    "agua", "agua potavel", "potavel", "agua mineral", "galao", "garrafa",

    // Lanches / apoio operacional
    "lanche", "lanches", "coffee", "cafe", "cafezinho", "snack", "moral",

    // Materiais de escritório (comuns)
    "caneta", "lapis", "borracha", "apontador", "marcador", "pilot", "pincel",
    "papel a4", "papel sulfite", "sulfite", "pasta", "arquivo", "etiqueta", "etiquetas",
    "grampo", "grampeador", "clipes", "cola", "fita adesiva", "tesoura", "grampos", "grampeadores", "bobina", "bobinas", "cabo", "regua", "régua",

    // Materiais de limpeza (comuns)
    "detergente", "sabao", "desinfetante", "alcool", "agua sanitaria",
    "papel toalha", "papel higienico", "limpeza", "pano", "esponja", "vassoura", "rodo",

    // Copa/cozinha (comuns)
    "copo", "copos", "guardanapo", "prato", "talher", "talheres", "mexedor",

    // Postagens/correios (comuns)
    "correios", "sedex", "postagem", "ar"
  ];

  // Palavras-chave “alerta” (potencialmente proibido/patrimonial/restrito)
  var alerta = [
    "notebook", "computador", "pc", "tablet", "celular", "smartphone", "iphone",
    "impressora", "scanner", "monitor", "tv", "televisao", "camera", "fone", "headset",
    "geladeira", "microondas", "ar condicionado", "ventilador",
    "movel", "moveis", "cadeira", "mesa", "compressor", "microondas", "steamer", "capa",
    "combustivel", "gasolina", "etanol", "diesel", "posto",
    "uber", "taxi", "corrida", "hospedagem", "hotel", "passagem", "viagem",
    "assinatura", "mensalidade", "streaming",
    "bebida", "alcool", "cerveja", "vinho", "whisky",
    "presente", "gift", "estante", "estantes", "steamer", "stemer", "stemar", "carrinho", "prateleiras", "prateleira", "carregadores", "carregador", "plenária", "plenaria", "cafeteira",
  ];

  // =========================
  // 1) ALERTA sempre primeiro
  // =========================
  for (var i = 0; i < alerta.length; i++) {
    if (d.indexOf(alerta[i]) !== -1) {
      return { status: "ALERTA", motivo: "Possível item restrito/patrimonial (revisar política e comprovante)." };
    }
  }

  // ============================================
  // 2) Regras combinadas (menos falso positivo)
  // ============================================
  // Lanche + equipe / treinamento (quando explicitado)
  if (d.indexOf("lanche") !== -1 && (d.indexOf("equipe") !== -1 || d.indexOf("trein") !== -1)) {
    return { status: "OK", motivo: "Despesa operacional (lanche para equipe/treinamento) conforme descrição." };
  }

  // Água (bem objetivo)
  if (d.indexOf("agua") !== -1 || d.indexOf("potavel") !== -1) {
    return { status: "OK", motivo: "Despesa operacional (água) conforme descrição." };
  }

  // Comunicação (bem objetivo)
  if (d.indexOf("comunicacao") !== -1) {
    return { status: "OK", motivo: "Despesa operacional (comunicação) conforme descrição." };
  }

  // ===================================
  // 3) OK por palavras-chave permitidas
  // ===================================
  for (var j = 0; j < permitidosProv.length; j++) {
    if (d.indexOf(permitidosProv[j]) !== -1) {
      return { status: "OK", motivo: "Compatível com despesa operacional provável, conforme descrição." };
    }
  }

  // ====================
  // 4) Genéricos → revisar
  // ====================
  if (d.length < 6 || d === "material" || d === "impressao" || d === "compra" || d === "servico") {
    return { status: "REVISAR", motivo: "Descrição genérica. Necessário validar comprovante e detalhamento." };
  }

  return { status: "REVISAR", motivo: "Não foi possível confirmar apenas pela descrição. Revisar comprovante." };
}

var filtroNorm = normalizarTexto_(valorFiltro);

// garante fim do dia no período (evita “vazar” datas)
if (dtIni instanceof Date) dtIni.setHours(0,0,0,0);
if (dtFim instanceof Date) dtFim.setHours(23,59,59,999);

// extrai código da loja do texto do autocomplete (ex.: "0046 - ...")
var soDigitos = (valorFiltro || "").replace(/\D/g, "");
var dig4 = "";
var m4 = soDigitos.match(/\d{4}/);
if (m4) dig4 = m4[0];
else if (soDigitos.length >= 4) dig4 = soDigitos.slice(-4);

// 1) primeiro filtra as linhas (matriz crua)
var linhasFiltradas = [];
for (var r = 0; r < linhas.length; r++) {
  var row = linhas[r];

  var d = parseDateClara_(row[idxData]);
  if (!d || isNaN(d.getTime())) continue;

  // dentro do período
  if (d < dtIni || d > dtFim) continue;

  // ===============================
  // filtro por loja / time / geral
  // ===============================
  if (tipoFiltro === "geral") {
    // não filtra
  } else if (tipoFiltro === "time") {
    var grupoLinha = row[idxGrupo];
    if (!grupoMatchTime_(grupoLinha, filtroNorm)) continue;

  } else if (tipoFiltro === "loja") {
    var lojaNum = (idxLojaNum >= 0 ? (row[idxLojaNum] || "").toString().trim() : "");
    var alias   = (idxAlias   >= 0 ? (row[idxAlias]   || "").toString().trim() : "");

    var aliasNorm   = normalizarTexto_(alias);
    var lojaNumNorm = lojaNum.replace(/\D/g, "");
    var bateu = false;

    // match por número (o mais confiável)
    if (dig4) {
      if (lojaNumNorm === dig4) bateu = true;
      // base costuma ter "CE0xxx" no Alias, então basta conter o 4 dígitos
      if (aliasNorm && aliasNorm.indexOf(dig4) !== -1) bateu = true;
    }

    // fallback textual (quando não veio número)
    if (!bateu && filtroNorm) {
      if (aliasNorm && aliasNorm.indexOf(filtroNorm) !== -1) bateu = true;
    }

    if (!bateu) continue;
  } else {
    continue;
  }

  linhasFiltradas.push(row);
}

// 2) agora monta a saída (objetos), SEM misturar com a matriz crua
var rows = [];
for (var i = 0; i < linhasFiltradas.length; i++) {
  var row = linhasFiltradas[i];

  var valor = Number(row[idxValor]) || 0;

  // data BR
  var dataCel = row[idxData];
  var dataBr = "";
  if (dataCel instanceof Date) {
    dataBr = Utilities.formatDate(dataCel, tz, "dd/MM/yyyy");
  } else {
    dataBr = (dataCel || "").toString();
  }

  var lojaOut = (idxAlias >= 0 ? String(row[idxAlias] || "") : "");
  var timeOut = (idxGrupo >= 0 ? String(row[idxGrupo] || "") : "");

  var itemRaw = (idxDescricao >= 0 ? String(row[idxDescricao] || "") : "");
  var itemNorm = normalizarItem_(itemRaw);

  var cls = classificarPolitica_(itemNorm);

  var estabOut = (idxTransacao >= 0 ? String(row[idxTransacao] || "") : "");

  rows.push({
    data: dataBr,
    valor: valor,
    loja: lojaOut,
    time: timeOut,
    item: itemRaw,
    transacao: estabOut,
    estabelecimento: estabOut,
    status: cls.status,
    conformidade: cls.status,
    motivo: cls.motivo,
    reincidencia: "",
    itemNorm: itemNorm
  });
}

    // Ordena por data desc e valor desc (para facilitar auditoria)
    rows.sort(function(a,b){
      if (a.dataISO < b.dataISO) return 1;
      if (a.dataISO > b.dataISO) return -1;
      return (b.valor || 0) - (a.valor || 0);
    });

    // Reincidência por loja + itemNorm (janela curta)
    var porChave = {};
    for (var i2 = 0; i2 < rows.length; i2++) {
      var rr = rows[i2];
      var chave = normalizarTexto_(rr.loja) + "||" + rr.itemNorm;
      if (!porChave[chave]) porChave[chave] = [];
      porChave[chave].push(rr);
    }

    // marca reincidência analisando datas (ordem asc dentro do grupo)
    Object.keys(porChave).forEach(function(ch){
      var arr = porChave[ch].slice().sort(function(a,b){
        return a.dataISO < b.dataISO ? -1 : (a.dataISO > b.dataISO ? 1 : 0);
      });

      var ultima = null;
      var countJanela = 0;
      for (var k = 0; k < arr.length; k++) {
        var cur = arr[k];
        var curD = new Date(cur.dataISO + "T00:00:00");
        if (ultima) {
          var diff = (curD - ultima) / 86400000;
          if (diff <= janelaDiasReincidencia) {
            countJanela++;
            cur.reincidencia = "Sim (" + (countJanela + 1) + "x em " + Math.round(diff) + " dias)";
          } else {
            countJanela = 0;
            cur.reincidencia = "";
          }
        } else {
          cur.reincidencia = "";
        }
        ultima = curD;
      }
    });

    // Insights rápidos
    var total = rows.length;
    var alertas = rows.filter(function(x){ return x.status === "ALERTA"; }).length;
    var revisar = rows.filter(function(x){ return x.status === "REVISAR"; }).length;

    return {
      ok: true,
      tipoFiltro: tipoFiltro,
      valorFiltro: valorFiltro,
      periodo: {
        inicio: Utilities.formatDate(dtIni, tz, "dd/MM/yyyy"),
        fim: Utilities.formatDate(dtFim, tz, "dd/MM/yyyy")
      },
      janelaDiasReincidencia: janelaDiasReincidencia,
      resumo: { total: total, alertas: alertas, revisar: revisar },
      rows: rows
    };

  } catch (e) {
    return { ok: false, error: "Erro ao listar itens comprados: " + (e && e.message ? e.message : e) };
  }
}

function vektorParseDateAny_(s) {
  if (s instanceof Date && !isNaN(s.getTime())) return s;

  s = String(s || "").trim();
  if (!s) return null;

  // yyyy-mm-dd (ou yyyy-mm-ddTHH...)
  var mIso = s.match(/^(\d{4})-(\d{2})-(\d{2})(?:[T\s].*)?$/);
  if (mIso) return new Date(Number(mIso[1]), Number(mIso[2]) - 1, Number(mIso[3]));

  // yyyy/MM/dd (ou yyyy/MM/dd HH:mm:ss)
  var mIso2 = s.match(/^(\d{4})\/(\d{2})\/(\d{2})(?:\s.*)?$/);
  if (mIso2) return new Date(Number(mIso2[1]), Number(mIso2[2]) - 1, Number(mIso2[3]));

  // dd/MM/yyyy (ou dd/MM/yyyy HH:mm:ss)
  var mBr = s.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:\s.*)?$/);
  if (mBr) return new Date(Number(mBr[3]), Number(mBr[2]) - 1, Number(mBr[1]));

  return null;
}

function vektorFmtBR_(d) {
  var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
  return Utilities.formatDate(d, tz, "dd/MM/yyyy");
}

/**
 * Radar - Visão Itens Irregulares (usa o MESMO motor do chat: getListaItensCompradosClara)
 * params: { dtIniIso, dtFimIso, itemContains, conformidade, groupBy, limit }
 */
function getItensIrregularesRadarVisao(params) {
  vektorAssertFunctionAllowed_("getItensIrregularesRadarVisao");
  try {
    // ============================
    // ✅ BLINDAGEM: aceita objeto (novo) e posicional (legado)
    // ============================
    var p = params;

    // se vier posicional: getItensIrregularesRadarVisao(dtIniIso, dtFimIso, itemContains, conformidade, groupBy, limit)
    if (p && typeof p !== "object") {
      p = {
        dtIniIso: arguments[0],
        dtFimIso: arguments[1],
        itemContains: arguments[2],
        conformidade: arguments[3],
        groupBy: arguments[4],
        limit: arguments[5]
      };
    } else {
      p = p || {};
    }

    // ============================
    // ✅ Leitura tolerante de chaves
    // ============================
    var dtIniIso = String(p.dtIniIso || p.dataIniIso || p.dtIni || p.ini || "").trim();
    var dtFimIso = String(p.dtFimIso || p.dataFimIso || p.dtFim || p.fim || "").trim();
    var itemContains = String(p.itemContains || p.qItem || p.itemQ || "").trim();
    var conformidade = String(p.conformidade || p.conf || "").trim().toUpperCase();
    var groupBy = String(p.groupBy || "loja").trim().toLowerCase();
    var limit = Number(p.limit) || 2500;

    // Debug de entrada (pra não ficar cego depois)
    var debugIn = {
      paramsType: (params === null ? "null" : typeof params),
      keys: (p && typeof p === "object") ? Object.keys(p) : [],
      dtIniIso: dtIniIso,
      dtFimIso: dtFimIso,
      itemContains: itemContains,
      conformidade: conformidade,
      groupBy: groupBy,
      limit: limit
    };

    // ============================
    // Datas
    // ============================
    var dIni = vektorParseDateAny_(dtIniIso);
    var dFim = vektorParseDateAny_(dtFimIso);
    if (!dIni || !dFim) {
      return {
        ok: false,
        error: "Período inválido (use o calendário para selecionar as datas).",
        debug: { in: debugIn }
      };
    }

    // normaliza intervalo
    dIni.setHours(0, 0, 0, 0);
    dFim.setHours(23, 59, 59, 999);

    // normaliza BR pro motor do chat
    var dtIniBR = vektorFmtBR_(dIni);
    var dtFimBR = vektorFmtBR_(dFim);

    // ============================
    // Motor (geral)
    // ⚠️ Se o seu motor NÃO exigir "geral", ele ignora; se exigir, isso evita retorno 0.
    // ============================
    var res = getListaItensCompradosClara("geral", "", dtIniBR, dtFimBR, limit);
    if (!res || !res.ok) {
      return {
        ok: false,
        error: (res && res.error) ? res.error : "Falha ao ler itens.",
        debug: { in: debugIn, dtIniBR: dtIniBR, dtFimBR: dtFimBR }
      };
    }

    var rows = Array.isArray(res.rows) ? res.rows : [];

    // ============================
    // filtros adicionais
    // ============================
    var itemNormQ = normalizarTexto_(itemContains || "");
    if (itemNormQ) {
      rows = rows.filter(function (r) {
        var it = normalizarTexto_(r.item || r.descricao || r.itemComprado || "");
        return it.indexOf(itemNormQ) !== -1;
      });
    }

    if (conformidade) {
      rows = rows.filter(function (r) {
        return String(r.conformidade || r.status || "").toUpperCase() === conformidade;
      });
    }

    if (rows.length > limit) rows = rows.slice(0, limit);

    // ============================
    // ✅ NORMALIZAÇÃO FORTE de LOJA/TIME + DATA
    // ============================
    function normLoja_(x) {
      var s = String(x || "").trim();
      if (!s) return "";
      var m = s.match(/(\d{1,4})/);
      if (!m) return s;
      return String(Number(m[1])).padStart(4, "0");
    }
    function normTime_(x) {
      var s = String(x || "").trim();
      return s;
    }
    function normDataTxt_(x) {
      if (x instanceof Date && !isNaN(x.getTime())) return vektorFmtBR_(x);
      var s = String(x || "").trim();
      if (!s) return "";
      // se vier ISO, tenta converter
      var d = vektorParseDateAny_(s);
      if (d) return vektorFmtBR_(d);
      return s;
    }

    // transforma pro front
    var outRows = rows.map(function (r) {
      var lojaRaw =
        r.lojaKey || r.lojaNum || r.lojaNumero || r.codLoja || r.cod_estbl ||
        r.loja || r.estabelecimento || "";

      var timeRaw =
        r.time || r.grupo || r.grupos || r.gruposRaw || r.area || "";

      return {
        dataTxt: normDataTxt_(r.dataTxt || r.data || r.dt || r.dataTransacao || ""),
        valor: Number(r.valor || r.vlr || 0) || 0,
        loja: normLoja_(lojaRaw),
        time: normTime_(timeRaw),
        item: String(r.item || r.descricao || r.itemComprado || ""),
        estabelecimento: String(r.estabelecimento || r.transacao || r.nomeEstabelecimento || r.merchant || ""),
        conformidade: String(r.conformidade || r.status || ""),
        motivo: String(r.motivo || r.justificativa || "")
      };
    });

    // ============================
    // KPIs no schema que o FRONT espera
    // ============================
    var totalItens = outRows.length;
    var totalValor = 0;
    var alertaQtd = 0;
    var alertaValor = 0;

    for (var i = 0; i < outRows.length; i++) {
      var v = Number(outRows[i].valor || 0);
      totalValor += v;
      if (String(outRows[i].conformidade || "").toUpperCase() === "ALERTA") {
        alertaQtd++;
        alertaValor += v;
      }
    }

    var alertaPctValor = (totalValor > 0) ? (alertaValor / totalValor) : 0;
    var alertaPctValorTxt = (alertaPctValor * 100).toFixed(1).replace(".", ",") + "%";

    function aggByKey_(keyName) {
      var map = {};
      for (var j = 0; j < outRows.length; j++) {
        var k = String(outRows[j][keyName] || "").trim() || "—";
        if (!map[k]) map[k] = { key: k, valor: 0, qtd: 0 };
        map[k].valor += Number(outRows[j].valor || 0);
        map[k].qtd += 1;
      }
      var arr = Object.keys(map).map(function (k) { return map[k]; });
      arr.sort(function (a, b) { return (b.valor || 0) - (a.valor || 0); });

      for (var z = 0; z < arr.length; z++) {
        var pct = (totalValor > 0) ? (arr[z].valor / totalValor) : 0;
        arr[z].pctTxt = (pct * 100).toFixed(1).replace(".", ",") + "%";
      }
      return arr;
    }

    function alertaSeries_(keyName) {
      var map = {};
      for (var k2 = 0; k2 < outRows.length; k2++) {
        if (String(outRows[k2].conformidade || "").toUpperCase() !== "ALERTA") continue;
        var kk = String(outRows[k2][keyName] || "").trim() || "—";
        if (!map[kk]) map[kk] = { key: kk, qtd: 0, valor: 0 };
        map[kk].qtd += 1;
        map[kk].valor += Number(outRows[k2].valor || 0);
      }
      var arr2 = Object.keys(map).map(function (x) { return map[x]; });
      arr2.sort(function (a, b) { return (b.valor || 0) - (a.valor || 0); });
      return arr2;
    }

    return {
      ok: true,
      debug: {
        in: debugIn,
        dtIniRaw: dtIniIso,
        dtFimRaw: dtFimIso,
        dtIniBR: dtIniBR,
        dtFimBR: dtFimBR,
        totalMotor: rows.length,
        totalAposFiltros: outRows.length
      },
      kpis: {
        totalItens: totalItens,
        totalValor: totalValor,
        alertaQtd: alertaQtd,
        alertaValor: alertaValor,
        alertaPctValorTxt: alertaPctValorTxt
      },
      rows: outRows,
      aggLoja: aggByKey_("loja"),
      aggTime: aggByKey_("time"),
      alertaByLoja: alertaSeries_("loja"),
      alertaByTime: alertaSeries_("time")
    };

  } catch (e) {
    return { ok: false, error: "Erro em getItensIrregularesRadarVisao: " + (e && e.message ? e.message : e) };
  }
}

// =====================================================
// ✅ RELAÇÃO DE SALDOS (ADM) — ciclo 06 -> hoje (volátil)
// Aceita filtro: geral | loja | time
// - Geral: agrega por Cartão+Loja+Time (Grupos)
// - Time:  agrega por Cartão+Loja (sem coluna Time na tabela)
// - Loja:  agrega por Cartão
// =====================================================

function getRelacaoSaldosClara(tipoFiltro, valorFiltro) {
  vektorAssertFunctionAllowed_("getRelacaoSaldosClara");
  try {
    // 🔒 Apenas Administrador
    var email = Session.getActiveUser().getEmail();
    if (!isAdminEmail(email)) {
      return { ok: false, error: "Acesso restrito: apenas Administrador pode consultar a relação de saldos." };
    }

    tipoFiltro = (tipoFiltro || "geral").toString().toLowerCase().trim();
    valorFiltro = (valorFiltro || "").toString().trim();

    var LIMITE_TETO = 3500; // teto hard do recomendado/ações

    // ✅ Loja desabilitada (você pediu só time e geral)
    if (tipoFiltro === "loja") {
      return { ok: false, error: "Consulta por loja não está habilitada. Use 'Relação de saldos geral' ou 'Relação de saldos do time X'." };
    }

    // --- 1) Período volátil (06 -> hoje; se dia 01–05, começa em 06 do mês anterior) ---
    var periodo = getPeriodoCicloClara_();
    var inicio = periodo.inicio;
    var fim = periodo.fim;

    // --- 2) Lê Info_limites ---
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID_CLARA);
    var shLim = ss.getSheetByName(SHEET_NOME_INFO_LIMITES);
    if (!shLim) {
      return { ok: false, error: "Aba '" + SHEET_NOME_INFO_LIMITES + "' não encontrada na planilha Captura_Clara." };
    }

    var limValues = shLim.getDataRange().getValues();
    if (!limValues || limValues.length < 2) {
      return { ok: false, error: "Aba Info_limites sem dados." };
    }

    // Mapa: chaveCartao -> {nome, tipo, titular, limite}
    var limites = {};
    for (var i = 1; i < limValues.length; i++) {
      var r = limValues[i];
      var nome = (r[0] || "").toString().trim();   // A
      if (!nome) continue;

      var tipo = (r[1] || "").toString().trim();   // B
      var titular = (r[3] || "").toString().trim();// D
      var limite = Number(r[4]) || 0;              // E

      var k = cartaoKeyCE_(nome);
      if (!k) continue;

      limites[k] = { nomeCartao: nome, tipo: tipo, titular: titular, limite: limite };
    }

    // --- 3) Lê BaseClara ---
    var info = carregarLinhasBaseClara_();
    if (info.error) return { ok: false, error: info.error };

    var header = info.header || [];
    var linhas = info.linhas || [];
    if (!linhas.length) return { ok: true, rows: [], periodo: formatPeriodoBR_(inicio, fim) };

    // Índices fixos (conforme você definiu)
    var idxAlias  = 7;   // H
    var idxGrupos = 17;  // R

    // Dinâmicos
    var idxValor = encontrarIndiceColuna_(header, ["Valor em R$", "Valor (R$)", "Valor", "Total"]);
    var idxData  = encontrarIndiceColuna_(header, ["Data da Transação", "Data Transação", "Data"]);
    var idxLoja  = encontrarIndiceColuna_(header, ["LojaNum", "Loja Num", "Loja Número", "Loja Numero", "Loja"]);

    if (idxValor < 0) return { ok: false, error: "Não encontrei a coluna de Valor na BaseClara." };
    if (idxData  < 0) return { ok: false, error: "Não encontrei a coluna de Data na BaseClara." };
    if (idxLoja  < 0) return { ok: false, error: "Não encontrei a coluna de Loja na BaseClara." };

    // Projeção por loja (6 ciclos + sazonalidade) — assume que você já substituiu a função para retornar {proj,meta}
    var projInfo = calcularProjecaoPorLojaUltimosCiclos_(linhas, idxData, idxValor, idxLoja);
    var projPorLoja = (projInfo && projInfo.proj) ? projInfo.proj : {};

    // --- 4) Agregação ---
    // ✅ INCLUIR NOVO: gastosPorDia para média móvel (por chave agregada)
    var agg = {}; // key -> { cartaoKey, nomeCartao, loja, time, usado, gastosPorDia: { 'YYYY-MM-DD': valor } }

    // --- 4.1) Mapa de vínculo (histórico completo): última loja/time por cartão ---
    var vinculoPorCartao = {}; // cartaoKey -> { loja, time, nomeCartao, dt }

    for (var h = 0; h < linhas.length; h++) {
      var r0 = linhas[h];

      var alias0 = (r0[idxAlias] || "").toString().trim();
      if (!alias0) continue;

      var dt0 = r0[idxData];
      var data0 = (dt0 instanceof Date) ? dt0 : new Date(dt0);
      if (!(data0 instanceof Date) || isNaN(data0.getTime())) continue;

      // Loja (mesma lógica que você já usa)
      var lojaRaw0 = (r0[idxLoja] || "").toString().trim();
      var lojaDigits0 = lojaRaw0.replace(/\D/g, "");
      var lojaKey0 = lojaDigits0 ? String(Number(lojaDigits0)).padStart(4, "0") : "";

      // Time (Grupos)
      var gruposRaw0 = (r0[idxGrupos] || "").toString().trim();

      // Cartão (chave padronizada)
      var cartaoKey0 = cartaoKeyCE_(alias0);
      if (!cartaoKey0) continue;

      // Regra do Rodrigo: sem vínculo => não registrar (fica oculto até ter 1ª transação com vínculo)
      // Se você considera que "time vazio" OU "loja vazia" é "sem vínculo", mantenha assim:
      if (!lojaKey0 || !gruposRaw0) continue;

      var cur = vinculoPorCartao[cartaoKey0];
      if (!cur || data0 > cur.dt) {
        vinculoPorCartao[cartaoKey0] = { loja: lojaKey0, time: gruposRaw0, nomeCartao: alias0, dt: data0 };
      }
    }

    // Filtro por time (somente quando tipoFiltro="time")
    var filtroTimeNorm = "";
    if (tipoFiltro === "time") {
      filtroTimeNorm = normalizarTexto_(valorFiltro);
      if (!filtroTimeNorm) return { ok: true, rows: [], aviso: "Time inválido." };
    }

    // ✅ Helper local: chave de data para gastosPorDia
    function vektorDateKey_(d) {
      // normaliza para meia-noite
      var dd = new Date(d.getFullYear(), d.getMonth(), d.getDate());
      var y = dd.getFullYear();
      var m = String(dd.getMonth() + 1).padStart(2, "0");
      var day = String(dd.getDate()).padStart(2, "0");
      return y + "-" + m + "-" + day;
    }

    for (var j = 0; j < linhas.length; j++) {
      var row = linhas[j];

      var alias = (row[idxAlias] || "").toString().trim();
      if (!alias) continue;

      var dt = row[idxData];
      var data = (dt instanceof Date) ? dt : new Date(dt);
      if (!(data instanceof Date) || isNaN(data.getTime())) continue;
      if (data < inicio || data > fim) continue;

      // Loja
      var lojaRaw = (row[idxLoja] || "").toString().trim();
      var lojaDigits = lojaRaw.replace(/\D/g, "");
      var lojaKey = lojaDigits ? String(Number(lojaDigits)).padStart(4, "0") : "";

      // Time (Grupos)
      var gruposRaw = (row[idxGrupos] || "").toString().trim();
      var gruposNorm = normalizarTexto_(gruposRaw);

      if (tipoFiltro === "time") {
        if (!gruposNorm || gruposNorm.indexOf(filtroTimeNorm) === -1) continue;
      }

      // Valor
      var v = parseNumberSafe_(row[idxValor]);
      if (!v) continue;

      // Cartão (chave padronizada)
      var cartaoKey = cartaoKeyCE_(alias);
      if (!cartaoKey) continue;

      // chave de agregação
      var key;
      if (tipoFiltro === "geral") {
        key = cartaoKey + "||" + lojaKey + "||" + normalizarTexto_(gruposRaw);
      } else {
        key = cartaoKey + "||" + lojaKey;
      }

      if (!agg[key]) {
        agg[key] = {
          cartaoKey: cartaoKey,
          nomeCartao: alias, // fallback
          loja: lojaKey,
          time: gruposRaw,
          usado: 0,
          gastosPorDia: {} // ✅ NOVO
        };
      }

      agg[key].usado += v;

      // ✅ NOVO: acumula gasto por dia (para média móvel)
      var dk = vektorDateKey_(data);
      if (!agg[key].gastosPorDia[dk]) agg[key].gastosPorDia[dk] = 0;
      agg[key].gastosPorDia[dk] += v;
    }

    // --- 4.2) Se não houve transação no ciclo, ainda assim queremos mostrar saldos (usado=0)
    // para cartões que já têm vínculo loja/time (histórico).
    Object.keys(vinculoPorCartao).forEach(function(cartaoKey) {
      var v = vinculoPorCartao[cartaoKey];
      if (!v) return;

      // Se for filtro por time, respeita
      if (tipoFiltro === "time") {
        var filtroTimeNorm2 = normalizarTexto_(valorFiltro);
        var vNorm = normalizarTexto_(v.time);
        if (!filtroTimeNorm2 || !vNorm || vNorm.indexOf(filtroTimeNorm2) === -1) return;
      }

      var key;
      if (tipoFiltro === "geral") {
        key = cartaoKey + "||" + v.loja + "||" + normalizarTexto_(v.time);
      } else {
        key = cartaoKey + "||" + v.loja;
      }

      if (!agg[key]) {
        agg[key] = {
          cartaoKey: cartaoKey,
          nomeCartao: v.nomeCartao || "",
          loja: v.loja || "",
          time: v.time || "",
          usado: 0,
          gastosPorDia: {} // ✅ NOVO (vazio)
        };
      }
    });

    // ✅ Helper local: média simples
    function mediaSimples_(arr) {
      if (!arr || !arr.length) return 0;
      var s = 0;
      for (var i = 0; i < arr.length; i++) s += (Number(arr[i]) || 0);
      return arr.length ? (s / arr.length) : 0;
    }

    // ✅ Helper local: média aparada (remove 1 maior e 1 menor) somente se houver sinal suficiente
    function mediaAparada_(arrPositivos) {
      var a = (arrPositivos || []).slice().filter(function(x){ return (Number(x) || 0) > 0; });
      if (a.length < 5) return mediaSimples_(a); // regra: <5 dias com uso => não aparar
      a.sort(function(x, y){ return x - y; });
      // remove 1 menor e 1 maior
      a.shift();
      a.pop();
      return mediaSimples_(a);
    }

    // ✅ Helper local: obter últimos N dias (calendário) do ciclo, até hoje0 (inclusive)
    function obterJanelaUltimosDias_(gastosPorDia, inicioCiclo, hoje0, janelaDias) {
      var out = [];
      var msDia = 24 * 60 * 60 * 1000;

      // começa no máximo N-1 dias atrás, mas nunca antes do início do ciclo
      var start = new Date(hoje0.getFullYear(), hoje0.getMonth(), hoje0.getDate());
      start = new Date(start.getTime() - (Math.max(0, (janelaDias - 1)) * msDia));

      var ini0 = new Date(inicioCiclo.getFullYear(), inicioCiclo.getMonth(), inicioCiclo.getDate());
      if (start < ini0) start = ini0;

      var d = new Date(start.getFullYear(), start.getMonth(), start.getDate());
      var end = new Date(hoje0.getFullYear(), hoje0.getMonth(), hoje0.getDate());

      while (d.getTime() <= end.getTime()) {
        var dk = vektorDateKey_(d);
        var val = (gastosPorDia && gastosPorDia[dk]) ? Number(gastosPorDia[dk]) : 0;
        out.push(val);
        d = new Date(d.getTime() + msDia);
      }
      return out;
    }

    // --- 5) Monta rows com limites + recomendação ---
    var rows = [];

    Object.keys(agg).forEach(function(k) {
      var a = agg[k];
      var lim = limites[a.cartaoKey];

      var limite = lim ? (Number(lim.limite) || 0) : 0;

      // ✅ NOVO: não exibir cartões sem limite (zerado/inativo)
      if (limite <= 0) return;   // <-- ESSA LINHA

      var tipo = lim ? (lim.tipo || "") : "";
      var titular = lim ? (lim.titular || "") : "";

      var saldo = limite - (a.usado || 0);

      var tipo = lim ? (lim.tipo || "") : "";
      var titular = lim ? (lim.titular || "") : "";

      var saldo = limite - (a.usado || 0);

      // --- dias restantes até o fechamento do ciclo (dia 05) ---
      var hoje = new Date();
      var hoje0 = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());
      var dHoje = hoje0.getDate();

      // fim do ciclo: se hoje >= 6, fecha dia 05 do próximo mês; senão, dia 05 do mês atual
      var fimCiclo = (dHoje >= 6)
        ? new Date(hoje0.getFullYear(), hoje0.getMonth() + 1, 5)
        : new Date(hoje0.getFullYear(), hoje0.getMonth(), 5);

      var msDia = 24 * 60 * 60 * 1000;
      var diasRestantes = Math.max(0, Math.ceil((fimCiclo.getTime() - hoje0.getTime()) / msDia));

      // Projeção por loja
      var projLoja = (projPorLoja && a.loja && projPorLoja[a.loja]) ? Number(projPorLoja[a.loja]) : 0;

      // --- Tempo do ciclo + ritmo atual (run-rate) ---
      var pc = getPeriodoCicloClara_();
      var ini = pc.inicio;

      // normaliza datas para evitar erro por horário
      var ini0  = new Date(ini.getFullYear(), ini.getMonth(), ini.getDate());

      // fim do ciclo é sempre dia 05 (mês corrente se hoje<=05; senão próximo mês)
      var dHoje = hoje0.getDate();
      var fimCiclo = (dHoje >= 6)
        ? new Date(hoje0.getFullYear(), hoje0.getMonth() + 1, 5)
        : new Date(hoje0.getFullYear(), hoje0.getMonth(), 5);

      var msDia = 24*60*60*1000;
      var diasDecorridos = Math.max(1, Math.floor((hoje0.getTime() - ini0.getTime()) / msDia) + 1);
      var diasRestantes  = Math.max(0, Math.ceil((fimCiclo.getTime() - hoje0.getTime()) / msDia));

      // ✅ GARANTIA: diasTotal (evita variável não definida em trechos abaixo)
      var diasTotal = Math.max(1, diasDecorridos + diasRestantes);

      // Projeção por ritmo do ciclo atual
      var usado = (a.usado || 0);
      var mediaDiaAtual = usado / diasDecorridos;
      var projRunRate = usado + (mediaDiaAtual * diasRestantes);

      // margem (mais conservador em Nov/Dez)
      var mesInicio = getPeriodoCicloOffset_(0).inicio.getMonth() + 1;
      var margem = (mesInicio === 11 || mesInicio === 12) ? 0.25 : 0.20;

      // --- Projeção final: histórico vs ritmo atual (controlada por Ritmo) ---

      // suaviza run-rate no início do ciclo
      var fatorRunRate = (diasDecorridos <= 7) ? 0.85 : 1.0;

      var projBase = projLoja || 0;                 // histórico (média ciclos)
      var rr = projRunRate * fatorRunRate;          // run-rate suavizado

      // Classificação de ritmo (mesma lógica que você já usa para a coluna "Ritmo de consumo")
      var pctCiclo = (diasTotal > 0) ? (diasDecorridos / diasTotal) : 0;

      // (evita "undefined" aqui; limiteAtual ainda não foi redeclarado neste trecho)
      var limiteAtualTmp = (limite || 0);
      var pctUsoLimTmp = (limiteAtualTmp > 0) ? (usado / limiteAtualTmp) : null;

      var ritmoRatio = (pctUsoLimTmp !== null && pctCiclo > 0) ? (pctUsoLimTmp / pctCiclo) : null;

      var ritmo = "—";
      if (ritmoRatio !== null && isFinite(ritmoRatio)) {
        if (ritmoRatio > 1.20) ritmo = "Alto";
        else if (ritmoRatio < 0.85) ritmo = "Baixo";
        else ritmo = "Médio";
      }

      // Política de projeção:
      // - Ritmo Alto: proteger operação => usa o maior (histórico vs run-rate)
      // - Ritmo Médio/Baixo: não deixa histórico dominar => usa run-rate com teto no histórico
      var projFinal;
      if (ritmo === "Alto") {
        // Aqui faz sentido ser conservador
        projFinal = Math.max(projBase, rr);
      } else {
        // Aqui o histórico alto não deve inflar recomendação quando o ciclo está calmo
        // teto do histórico: no máximo +20% sobre o run-rate (evita “CE0234” inflando)
        var tetoHistorico = rr * 1.20;

        // também evita cair demais se run-rate estiver muito baixo por poucos dias
        // piso: pelo menos 60% do histórico (ajuste fino)
        var pisoHistorico = projBase > 0 ? (projBase * 0.60) : rr;

        projFinal = Math.max(pisoHistorico, Math.min(projBase, tetoHistorico));
      }

      // quanto ainda tende a gastar no ciclo
      var restante = Math.max(projFinal - usado, 0);

      // buffer mínimo
      var bufferMin = Math.max(200, projFinal * 0.05);

      // limite recomendado: utilizado + folga para o restante
      var limiteRec = usado + Math.max(restante * (1 + margem), bufferMin);

      // ✅ TETO: não recomendar acima de 3.500
        limiteRec = Math.min(limiteRec, LIMITE_TETO);

        // segurança: nunca recomendar abaixo do já utilizado
        limiteRec = Math.max(limiteRec, usado);

      // --- trava de redução por tempo do ciclo ---
      var hojeT = new Date();
      var pcT = getPeriodoCicloClara_();
      var iniT = pcT.inicio, fimT = pcT.fim;
      var msDiaT = 24 * 60 * 60 * 1000;

      var diasTot = Math.max(1, Math.round((fimT.getTime() - iniT.getTime()) / msDiaT) + 1);
      var diasDec = Math.max(1, Math.floor((hojeT.getTime() - iniT.getTime()) / msDiaT) + 1);
      var passouMetade = diasDec >= Math.ceil(diasTot / 2);

      var pctProj = (projLoja > 0) ? ((a.usado || 0) / projLoja) : null;

      // Se passou metade e já consumiu >50% da projeção, não reduzir
      var travaReducaoTempo = (passouMetade && pctProj !== null && pctProj > 0.50);

      // Ação
      var acao = "Manter";
      var delta = limiteRec - (limite || 0);

      // --- gatilho de aumento por risco (override) ---
      var limiteAtual = (limite || 0);
      var utilizado = (a.usado || 0);
      var pctUsoLim = (limiteAtual > 0) ? (utilizado / limiteAtual) : null;

      var saldoBaixo = (saldo <= 500);
      var faltamMuitosDias = (diasRestantes >= 10);
      var jaPassouMetadeLimite = (pctUsoLim !== null && pctUsoLim >= 0.50);

      // Se apertado + tempo suficiente: sugerir aumento mesmo com delta pequeno
      var forcarAumentoPorRisco = (saldoBaixo && faltamMuitosDias && jaPassouMetadeLimite);

      // ==========================================================
      // ✅ NOVO OVERRIDE (SEM MUDAR O RESTO): dias restantes + média móvel
      // ==========================================================
      // Regra:
      // - considera últimos 7 dias do ciclo (até hoje)
      // - usa média simples dos dias com uso e, se dias com uso >= 5, usa média aparada (remove maior e menor)
      // - mediaRef = max(média simples, média aparada) para não "maquiar" escalada real
      // - se saldo_restante / dias_restantes < mediaRef * fatorSeg => forçar aumento
      var janela = 7;
      var fatorSeg = 1.15;

      var janelaVals = obterJanelaUltimosDias_(a.gastosPorDia || {}, inicio, hoje0, janela);
      var valsPos = (janelaVals || []).filter(function(x){ return (Number(x) || 0) > 0; });

      var diasComUso = valsPos.length; // dentro da janela
      var media7Simples = mediaSimples_(valsPos);
      var media7Aparada = mediaAparada_(valsPos); // já respeita <5 => simples
      var mediaRef = Math.max(media7Simples, media7Aparada);

      var saldoRestante = (limiteAtual || 0) - (utilizado || 0);
      var saldoPorDiaRestante = (Math.max(diasRestantes, 1) > 0) ? (saldoRestante / Math.max(diasRestantes, 1)) : saldoRestante;

      // Travas (evitar ruído e decisões ruins):
      // - não aplicar quando faltam pouquíssimos dias (<=2) porque qualquer variação explode
      // - precisa ter algum sinal de uso (>=3 dias com uso na janela)
      // - precisa ter limite e dias restantes > 0
      var forcarAumentoPorMediaDias =
        (diasRestantes > 2) &&
        (limiteAtual > 0) &&
        (diasComUso >= 3) &&
        (mediaRef > 0) &&
        (saldoPorDiaRestante < (mediaRef * fatorSeg));

      var tol = 0.05;
      var minDelta = 200;
      var limiteAtual = (limite || 0);

      var nomeNorm = normalizarTexto_(lim ? lim.nomeCartao : a.nomeCartao);
      var bloqueiaReducao = nomeNorm.indexOf("temporario") !== -1
                         || nomeNorm.indexOf("virtual") !== -1
                         || nomeNorm.indexOf("virual") !== -1;

      // TRAVAS PARA "REDUZIR" (coerência com projeção e risco operacional)

      // % da projeção (se projeção existir)
      var pctProj = (projLoja > 0) ? ((a.usado || 0) / projLoja) : null;

      // 1) Se já bateu/ultrapassou a projeção, NUNCA reduzir
      var travaReducaoPorProj = (pctProj !== null && pctProj >= 1.0);

      // 2) Se o saldo já está "apertado", evitar redução (não piorar risco)
      var saldoAtual = (limite || 0) - (a.usado || 0);
      var travaReducaoPorSaldoApertado = saldoAtual <= 500; // alinhado com saldo crítico atual

      // Trava final: se qualquer uma for verdadeira, não permitir "Reduzir"
      var travaReducao = travaReducaoPorProj || travaReducaoPorSaldoApertado;

      if (limiteAtual <= 0) {
        acao = "Definir " + Utilities.formatString("R$ %.0f", limiteRec);

      } else if (forcarAumentoPorRisco) {
        // Override: saldo baixo + muitos dias restantes + já consumiu metade do limite
        // alvo mínimo: pelo menos +200 ou até o limiteRec (o que for maior)
        var alvo = Math.max(limiteRec, limiteAtual + 200);

        // arredonda para múltiplos de 100
        alvo = Math.ceil(alvo / 100) * 100;

        alvo = Math.min(alvo, LIMITE_TETO);

        var deltaRisco = alvo - limiteAtual;
        if (deltaRisco > 0) {
          acao = "Aumentar +" + moneyBR_(deltaRisco);

          // opcional (mas recomendado): alinhar o limiteRec com o alvo para consistência
          limiteRec = alvo;
          delta = limiteRec - limiteAtual;
        }

      } else if (forcarAumentoPorMediaDias) {
        // ✅ NOVO: Override por capacidade diária vs ritmo real (média móvel)
        // Objetivo: garantir saldo suficiente para sustentar o ritmo até o fechamento
        // alvoNecessario = utilizado + (mediaRef*fatorSeg*diasRestantes)
        var alvoNecessario = (utilizado || 0) + (mediaRef * fatorSeg * Math.max(diasRestantes, 1));

        // respeita pelo menos o limiteRec (não "briga" com o cálculo atual)
        var alvo2 = Math.max(limiteRec, alvoNecessario);

        // garante aumento mínimo coerente (>= +200) quando forçar
        alvo2 = Math.max(alvo2, limiteAtual + 200);

        // arredonda para múltiplos de 100
        alvo2 = Math.ceil(alvo2 / 100) * 100;

        alvo2 = Math.min(alvo2, LIMITE_TETO);

        var deltaMedia = alvo2 - limiteAtual;
        if (deltaMedia > 0) {
          acao = "Aumentar +" + moneyBR_(deltaMedia);

          // mantém consistência com o que será exibido como "limite recomendado"
          limiteRec = Math.max(limiteRec, alvo2);
          delta = limiteRec - limiteAtual;
        }

      } else if (limiteAtual < (limiteRec * (1 - tol)) && delta >= minDelta) {
        acao = "Aumentar +" + moneyBR_(delta);

      } else if (!bloqueiaReducao && !travaReducaoTempo && limiteAtual > (limiteRec * (1 + tol)) && (-delta) >= minDelta) {
        acao = "Reduzir -" + moneyBR_(-delta);
      }

      // 🔕 Exclusão pontual: CE0234 - VIRTUAL MARKETING (somente este alias)
      var nomeCartaoFinal = (lim ? lim.nomeCartao : a.nomeCartao) || "";
      var nomeNorm = normalizarTexto_(nomeCartaoFinal);

      // Regra: CE0234 + VIRTUAL + MARKETING
      var ehCE0234 = nomeNorm.indexOf("ce0234") === 0;
      var ehVirtual = nomeNorm.indexOf("virtual") !== -1 || nomeNorm.indexOf("virual") !== -1;
      var ehMarketing = nomeNorm.indexOf("marketing") !== -1;

      if (ehCE0234 && ehVirtual && ehMarketing) {
        return; // pula APENAS este cartão
      }

      // ------------------------------
      // Ritmo de consumo no ciclo (06→05)
      // ------------------------------
      var usado = (a.usado || 0);
      var limiteAtual = (limite || 0);

      var hoje = new Date();
      var hoje0 = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());

      // "inicio" já existe na função (periodo.inicio). Normaliza:
      var ini0 = new Date(inicio.getFullYear(), inicio.getMonth(), inicio.getDate());

      // Fim do ciclo: dia 05 do mês correto
      var dHoje = hoje0.getDate();
      var fimCiclo = (dHoje >= 6)
        ? new Date(hoje0.getFullYear(), hoje0.getMonth() + 1, 5)
        : new Date(hoje0.getFullYear(), hoje0.getMonth(), 5);

      var msDia = 24 * 60 * 60 * 1000;
      var diasDecorridos = Math.max(1, Math.floor((hoje0.getTime() - ini0.getTime()) / msDia) + 1);
      var diasTotal = Math.max(diasDecorridos, Math.round((fimCiclo.getTime() - ini0.getTime()) / msDia) + 1);

      var pctCiclo = diasTotal > 0 ? (diasDecorridos / diasTotal) : 0;
      var pctUsoLim = (limiteAtual > 0) ? (usado / limiteAtual) : null;

      var ritmoRatio = (pctUsoLim !== null && pctCiclo > 0) ? (pctUsoLim / pctCiclo) : null;

      var ritmo = "—";
      if (ritmoRatio !== null && isFinite(ritmoRatio)) {
        if (ritmoRatio > 1.20) ritmo = "Alto";
        else if (ritmoRatio < 0.85) ritmo = "Baixo";
        else ritmo = "Médio";
      }

      rows.push({
        nomeCartao: nomeCartaoFinal,
        loja: a.loja,
        time: a.time,
        tipo: tipo,
        titular: titular,
        limite: limite,
        utilizado: a.usado || 0,
        projecao: projLoja,
        limiteRecomendado: limiteRec,
        acao: acao,
        ritmo: ritmo,              // ✅ NOVO
        //ritmoRatio: ritmoRatio,  // opcional (não exibir na tabela)
        saldo: saldo
      });
    }); // ✅ FECHA O forEach corretamente aqui

    // Ordena por menor saldo
    rows.sort(function(x, y) { return (x.saldo || 0) - (y.saldo || 0); });

    var minRow = rows.length ? rows[0] : null;
    var maxRow = rows.length ? rows[rows.length - 1] : null;

    // Se você ainda usa esse campo em insights, mantém; se não, pode remover depois
    var proj = projeçãoCiclo_(inicio, fim, 0);

    return {
      ok: true,
      tipoFiltro: tipoFiltro,
      valorFiltro: valorFiltro,
      periodo: formatPeriodoBR_(inicio, fim),
      highlights: {
        menorSaldo: minRow ? { nomeCartao: minRow.nomeCartao, saldo: minRow.saldo, loja: minRow.loja, time: minRow.time } : null,
        maiorSaldo: maxRow ? { nomeCartao: maxRow.nomeCartao, saldo: maxRow.saldo, loja: maxRow.loja, time: maxRow.time } : null,
        predisposicao: proj
      },
      rows: rows
    };

  } catch (e) {
    return { ok: false, error: "Falha ao calcular relação de saldos: " + (e && e.message ? e.message : e) };
  }
}

/**
 * Retorna transações da loja (Alias do Cartão / nomeCartao) no período aberto do ciclo:
 * início = dia 06 do ciclo atual
 * fim = hoje (fim do dia)
 *
 * Colunas solicitadas (BaseClara):
 * A (0)  = Data da Transação
 * F (5)  = Valor em R$
 * H (7)  = Alias do Cartão (Loja no seu contexto)
 * U (20) = Descrição (Item comprado)
 */
function getTransacoesLojaPeriodoAberto(aliasLoja) {
  vektorAssertFunctionAllowed_("getTransacoesLojaPeriodoAberto");
  try {
    var email = Session.getActiveUser().getEmail();
    // (opcional) se quiser restringir só ADM:
    // if (!isAdminEmail(email)) return { ok:false, error:"Acesso restrito." };

    aliasLoja = (aliasLoja || "").toString().trim();
    if (!aliasLoja) return { ok: true, rows: [], meta: { inicio:"", fim:"" } };

    var info = carregarLinhasBaseClara_();
    if (info.error) return { ok: false, error: info.error };

    var linhas = info.linhas || [];
    if (!linhas.length) return { ok: true, rows: [], meta: { inicio:"", fim:"" } };

    // Período aberto do ciclo (06 -> hoje)
    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    var pc = getPeriodoCicloClara_(); // você já usa isso no projeto
    var ini = pc && pc.inicio ? new Date(pc.inicio) : null;
    if (!ini || isNaN(ini.getTime())) return { ok:false, error:"Não consegui determinar o início do ciclo (06)." };

    ini.setHours(0,0,0,0);

    var hoje = new Date();
    var fim = new Date(hoje);
    fim.setHours(23,59,59,999);

    // Índices fixos
    var IDX_DATA  = 0;   // A
    var IDX_VALOR = 5;   // F
    var IDX_ALIAS = 7;   // H
    var IDX_ITEM  = 20;  // U

    function fmtBR(d) {
      return Utilities.formatDate(d, tz, "dd/MM/yyyy");
    }

    var out = [];
    for (var i = 0; i < linhas.length; i++) {
      var r = linhas[i];
      if (!r) continue;

      var alias = (r[IDX_ALIAS] || "").toString().trim();
      if (alias !== aliasLoja) continue;

      var d = parseDateClara_(r[IDX_DATA]);
      if (!d || isNaN(d.getTime())) continue;

      if (d < ini || d > fim) continue;

      var v = Number(r[IDX_VALOR]) || 0;

      out.push({
        data: fmtBR(d),
        valor: v,
        loja: alias,                 // coluna H (Alias)
        item: (r[IDX_ITEM] || "").toString()
      });
    }

    // ordena por data desc (opcional)
    out.sort(function(a,b){
      // como data vem dd/MM/yyyy, ordena por Date real:
      function toDt(x){
        var p = (x||"").split("/");
        if (p.length !== 3) return new Date(0);
        return new Date(Number(p[2]), Number(p[1])-1, Number(p[0]));
      }
      return toDt(b.data) - toDt(a.data);
    });

    return {
      ok: true,
      rows: out,
      meta: {
        alias: aliasLoja,
        inicio: fmtBR(ini),
        fim: fmtBR(hoje) // “até a data atual” como você pediu
      }
    };

  } catch (e) {
    return { ok:false, error: (e && e.message) ? e.message : String(e) };
  }
}

function getTabelaEstornosClara(tipoFiltro, valorFiltro) {
  vektorAssertFunctionAllowed_("getTabelaEstornosClara");
  try {
    var info = carregarLinhasBaseClara_();
    if (!info) return { ok:false, error:"BaseClara não carregada (retorno vazio)." };
    if (info.error) return { ok:false, error: info.error };

    // seu loader retorna { header, linhas }
    var headers = info.header || [];
    var rowsAll = info.linhas || [];

    if (!rowsAll || !rowsAll.length) {
      return { ok:false, error:"BaseClara sem dados (linhas vazias)." };
    }

    // ====== ÍNDICES FIXOS (conforme sua regra)
    // Coluna C = Transação (Estabelecimento) | Coluna D = Valor original
    var IDX_ESTAB = 2;   // C (0-based)
    var IDX_VALOR = 3;   // D (0-based)
    var IDX_FATURA = 1;  // B (Extrato da conta) => Período da fatura


    function idxByNames(possiveis) {
      return encontrarIndiceColuna_(headers, possiveis);
    }

    // Esses podem variar de posição, então usamos fallback por nome
    // (mas sem deixar quebrar o estabelecimento/valor original que agora é por índice)
    var idxLojaNum = idxByNames(["LojaNum"]);
    var idxAlias   = idxByNames(["Alias Do Cartão", "Alias do Cartão", "Alias"]);
    var idxTime    = idxByNames(["Grupos"]);
    var idxData    = idxByNames(["Data da Transação"]);
    var idxTit     = idxByNames(["Titular"]);
    var idxCat     = idxByNames(["Categoria da Compra"]);

    // Fallback comum: se não achar Data por nome, tenta coluna A (index 0)
    if (idxData < 0) idxData = 0;

    var tipo = (tipoFiltro || "geral").toString().toLowerCase();
    var vf = (valorFiltro || "").toString().trim();

    function norm(s) { return normalizarTexto_(s || ""); }

    var filtroTimeNorm = (tipo === "time") ? norm(vf) : "";

    function parseMoneyBR(v) {
      if (v === null || v === undefined || v === "") return 0;
      if (typeof v === "number") return v;
      var s = String(v).trim();
      s = s.replace(/[R$\s]/g, "");
      s = s.replace(/\./g, "").replace(",", ".");
      var n = Number(s);
      return isFinite(n) ? n : 0;
    }

    function toDateMaybe(v) {
      if (!v) return null;
      if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) return v;

      var s = String(v).trim();

      var m1 = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
      if (m1) return new Date(Number(m1[3]), Number(m1[2]) - 1, Number(m1[1]));

      var m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (m2) return new Date(Number(m2[1]), Number(m2[2]) - 1, Number(m2[3]));

      var d = new Date(s);
      return isNaN(d.getTime()) ? null : d;
    }

    function fmtDate(d) {
      if (!d) return "";
      if (Object.prototype.toString.call(d) === "[object Date]" && !isNaN(d.getTime())) {
        var dd = ("0" + d.getDate()).slice(-2);
        var mm = ("0" + (d.getMonth()+1)).slice(-2);
        var yy = d.getFullYear();
        return dd + "/" + mm + "/" + yy;
      }
      return String(d);
    }

    function parseLojaNumFromAlias(alias) {
      var s = String(alias || "").trim();     // ex: CE0046
      var m = s.match(/(\d{1,4})/);
      if (!m) return "";
      return String(Number(m[1]));           // "0046" -> "46"
    }

    function parseLojaNum(v, alias) {
      var s = String(v || "").trim();
      var digits = s.replace(/\D/g, "");
      if (digits) return String(Number(digits));   // "0101" -> "101"
      if (alias) return parseLojaNumFromAlias(alias);
      return "";
    }

    var estornos = [];
    var minD = null, maxD = null;

    for (var i = 0; i < rowsAll.length; i++) {
      var r = rowsAll[i];

      // segurança de tamanho mínimo da linha
      if (!r || r.length < 4) continue;

      var valorO = parseMoneyBR(r[IDX_VALOR]);
      if (!(valorO < 0)) continue; // estorno = negativo

      var alias = (idxAlias >= 0 ? r[idxAlias] : "");
      var loja = parseLojaNum(idxLojaNum >= 0 ? r[idxLojaNum] : "", alias);

      var time = (idxTime >= 0 ? (r[idxTime] || "").toString().trim() : "");

      if (tipo === "time") {
        if (!filtroTimeNorm) continue;
        if (norm(time) !== filtroTimeNorm) continue;
      }

      var dt = toDateMaybe(r[idxData]);
      if (dt) {
        if (!minD || dt.getTime() < minD.getTime()) minD = dt;
        if (!maxD || dt.getTime() > maxD.getTime()) maxD = dt;
      }

      estornos.push({
        loja: loja,
        time: time,
        dataTransacao: fmtDate(dt || r[idxData]),
        periodoFatura: (r[IDX_FATURA] || "").toString(),
        valorEstorno: valorO,
        // ✅ estabelecimento fixo pela coluna C
        estabelecimento: (r[IDX_ESTAB] || "").toString(),
        titular: (idxTit >= 0 ? (r[idxTit] || "").toString() : ""),
        categoria: (idxCat >= 0 ? (r[idxCat] || "").toString() : "")
      });
    }

    // ordena por data desc
    estornos.sort(function(a,b){
      var da = toDateMaybe(a.dataTransacao);
      var db = toDateMaybe(b.dataTransacao);
      if (da && db) return db.getTime() - da.getTime();
      return 0;
    });

    var periodo = {
      inicio: minD ? fmtDate(minD) : "",
      fim: maxD ? fmtDate(maxD) : ""
    };

    // highlights
    var total = estornos.length;

    var countPorLoja = {};
    for (var j = 0; j < estornos.length; j++) {
      var lj = estornos[j].loja || "";
      countPorLoja[lj] = (countPorLoja[lj] || 0) + 1;
    }

    var topLoja = null;
    Object.keys(countPorLoja).forEach(function(lj){
      if (!topLoja || countPorLoja[lj] > topLoja.qtd) {
        topLoja = { loja: lj, qtd: countPorLoja[lj] };
      }
    });

    var lojaMaiorPct = null;
    if (topLoja && total > 0) {
      var pct = (topLoja.qtd / total) * 100;
      lojaMaiorPct = {
        loja: topLoja.loja,
        qtd: topLoja.qtd,
        pct: pct.toFixed(1).replace(".", ",") + "%"
      };
    }

    var maiorEstorno = null;
    for (var k = 0; k < estornos.length; k++) {
      var e = estornos[k];
      var mag = Math.abs(Number(e.valorEstorno) || 0);
      if (!maiorEstorno || mag > Math.abs(Number(maiorEstorno.valor) || 0)) {
        maiorEstorno = {
          valor: e.valorEstorno,
          data: e.dataTransacao,
          loja: e.loja,
          time: e.time
        };
      }
    }

    return {
      ok: true,
      tipoFiltro: tipoFiltro || "geral",
      valorFiltro: valorFiltro || "",
      periodo: periodo,
      highlights: {
        lojaMaiorPct: lojaMaiorPct,
        maiorEstorno: maiorEstorno
      },
      rows: estornos
    };

  } catch (e) {
    return { ok:false, error:"Erro em getTabelaEstornosClara: " + e };
  }
}

// ============================
// ALERTAS LIMITES (E-MAIL ADM)
// ============================

const VEKTOR_ALERT_SALDO_CRITICO = 500;      // R$ 500,00 (definido internamente)
const VEKTOR_ALERT_REDUCAO_MIN = 500;        // Redução relevante (R$ 500)
const VEKTOR_ALERT_MAX_RISCO = 15;           // Limite de itens no e-mail (risco)
const VEKTOR_ALERT_MAX_EFICIENCIA = 10;      // Limite de itens no e-mail (eficiência)
const VEKTOR_ALERT_MAX_ADMIN = 20;           // Limite de itens no e-mail (admin)
const VEKTOR_ALERT_TOL_PCT = 0.0000001;

// Disparo principal (use no gatilho diário)
function enviarAlertasLimitesClaraDiario() {

  if (typeof periodoStr !== "string") {
  try {
    if (periodoStr && (periodoStr.inicio || periodoStr.fim)) {
      periodoStr = (periodoStr.inicio || "06") + " a " + (periodoStr.fim || "05");
    } else {
      periodoStr = "06→05";
    }
  } catch (e) {
    periodoStr = "06→05";
  }
}

  // Segurança: só roda para Admin
  var email = Session.getEffectiveUser().getEmail();
  if (!isAdminEmail(email)) {
    return { ok: false, error: "Acesso restrito: apenas Administrador pode disparar alertas." };
  }

  // Pega base já calculada (mesma da tabela)
  var res = getRelacaoSaldosClara("geral", "");
  if (!res || !res.ok) {
    return { ok: false, error: (res && res.error) ? res.error : "Falha ao obter relação de saldos." };
  }

  var periodo = "";
if (typeof res.periodo === "string") {
  periodo = res.periodo;
} else if (res.periodo && (res.periodo.inicio || res.periodo.fim)) {
  periodo = (res.periodo.inicio || "06") + " a " + (res.periodo.fim || "05");
} else {
  periodo = "06→05";
}
  var rows = res.rows || [];
  if (!rows.length) return { ok: true, sent: false, msg: "Sem dados para alertar." };

  // Classifica
  var packs = classificarAlertasLimites_(rows);

  // Anti-spam por ciclo (06->05)
  var cicloKey = getCicloKey06a05_(); // ex: "2025-12-06_2026-01-05"
  var filtrados = aplicarAntiSpamCiclo_(cicloKey, packs);

  var risco = filtrados.risco;
  var monitoramento = filtrados.monitoramento;
  var eficiencia = filtrados.eficiencia;
  var admin = filtrados.admin;

  Logger.log("ALERT COUNTS: risco=%s monitoramento=%s eficiencia=%s admin=%s",
  risco.length, monitoramento.length, eficiencia.length, admin.length);

  if (!risco.length && !monitoramento.length && !eficiencia.length && !admin.length) {
    return { ok: true, sent: false, msg: "Sem alertas novos (anti-spam por ciclo)." };
  }


  // Monta e-mail
  var assunto = risco.length
    ? "⚠️ [ALERTA CLARA | LIMITE] Risco de estouro"
    : "⚠️ [ALERTA] Ajustes de limite recomendados – Vektor";

  var html = montarEmailAlertasLimites_(periodo, risco, monitoramento, eficiencia, admin);

  // Envia somente para ADM’s
  var destinatarios = getAdminEmails_();
  if (!destinatarios.length) return { ok: false, error: "Lista de admins vazia." };

  GmailApp.sendEmail(destinatarios.join(","), assunto, " ", {
  from: "vektor@gruposbf.com.br",
  htmlBody: html,
  name: "Vektor – Grupo SBF"
  });

  // Após MailApp.sendEmail(...)
  registrarAlertaEnviado_(
  "LIMITE",
  "",
  "",
  "Envio consolidado de alertas de limite. Risco=" + (risco.length) +
    ", Monitoramento=" + (monitoramento.length) + ", Eficiência=" + (eficiencia.length) + ", Admin=" + (admin.length),
  destinatarios.join(","),
  "enviarAlertasLimitesClaraDiario"
);

  // Registra enviados no ciclo para anti-spam
  registrarEnviadosCiclo_(cicloKey, filtrados._enviadosKeys);

  return {
    ok: true,
    sent: true,
    counts: { risco: risco.length, eficiencia: eficiencia.length, admin: admin.length },
    ciclo: cicloKey
  };
}

// ======================================
// RELATÓRIO OFENSORAS - PENDÊNCIAS CLARA
// ======================================
function gerarRelatorioOfensorasPendencias_(diasJanela) {
  diasJanela = Number(diasJanela) || 60;

  var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
  var hist = ss.getSheetByName(HIST_PEND_CLARA_RAW);
  if (!hist) throw new Error("Aba " + HIST_PEND_CLARA_RAW + " não encontrada.");

  var lr = hist.getLastRow();
  if (lr < 2) return { ok: true, rows: [], msg: "Histórico vazio.", janelaDias: diasJanela };

  // Colunas: A Data_snapshot, B Loja, C Data_transacao, D Valor, E Pendencia_etiqueta, F Pendencia_nf, G Pendencia_descricao, H Qtde Total
  var data = hist.getRange(2, 1, lr - 1, 8).getValues();

  // Mapa LojaNum -> Time (BaseClara!V -> BaseClara!R)
  var mapaTime = construirMapaLojaParaTime_();

  var hoje = new Date();
  var inicio = new Date(hoje.getTime() - diasJanela * 24 * 60 * 60 * 1000);

  // agregação por loja
  var m = {}; // lojaKey -> stats
  function getLojaKey(loja){ return String(loja || "").trim() || "(N/D)"; }

  data.forEach(function(r){
    var dtSnap = (r[0] instanceof Date) ? r[0] : (vektorParseDateAny_(r[0]) || new Date(r[0]));
    if (!(dtSnap instanceof Date) || isNaN(dtSnap.getTime())) return;
    if (dtSnap < inicio) return;

    var lojaKey = getLojaKey(r[1]);
    var lojaNum = normalizarLojaNumero_(lojaKey);

    var dtTx = (r[2] instanceof Date) ? r[2] : (vektorParseDateAny_(r[2]) || new Date(r[2]));


    var valor = Number(r[3]) || 0;
    var pe = Number(r[4]) || 0;
    var pn = Number(r[5]) || 0;
    var pd = Number(r[6]) || 0;
    var qt = Number(r[7]) || (pe + pn + pd);

    if (!m[lojaKey]) {
      m[lojaKey] = {
        loja: lojaKey,
        lojaNum: lojaNum,
        time: "N/D",
        qtde: 0,
        valor: 0,
        txCount: 0,
        pendEtiqueta: 0,
        pendNF: 0,
        pendDesc: 0,
        snaps: {},     // dias distintos de snapshot
        diasTx: {},    // dias distintos de transação (pendência)
        maxSnap: null  // último snapshot observado
      };
    }

    // resolve time (se existir no mapa)
    if (m[lojaKey].lojaNum && mapaTime[m[lojaKey].lojaNum]) {
      m[lojaKey].time = mapaTime[m[lojaKey].lojaNum];
    }

    m[lojaKey].qtde += qt;
    m[lojaKey].valor += valor;
    m[lojaKey].pendEtiqueta += pe;
    m[lojaKey].pendNF += pn;
    m[lojaKey].pendDesc += pd;
    m[lojaKey].txCount += 1;

    // snapshot day key
    var snapKey = Utilities.formatDate(dtSnap, "America/Sao_Paulo", "yyyy-MM-dd");
    m[lojaKey].snaps[snapKey] = true;

    // transação day key (se data válida)
    if (dtTx instanceof Date && !isNaN(dtTx.getTime())) {
      var txKey = Utilities.formatDate(dtTx, "America/Sao_Paulo", "yyyy-MM-dd");
      m[lojaKey].diasTx[txKey] = true;
    }

    // max snapshot
    if (!m[lojaKey].maxSnap || dtSnap > m[lojaKey].maxSnap) {
      m[lojaKey].maxSnap = dtSnap;
    }
  });

  var rows = Object.keys(m).map(function(k){
    var s = m[k];

    var diasPend = Object.keys(s.diasTx).length;
    var qtdSnaps = Object.keys(s.snaps).length;

    // aceleração recente (mantém sua lógica)
    var r14 = calcularTrend14dias_(data, s.loja);

    // score composto (ajuste se quiser outra ponderação)
    var score = (
      (s.qtde || 0) +
      (s.pendEtiqueta || 0) * 2 +
      (s.pendNF || 0) * 2 +
      (s.pendDesc || 0) * 1 +
      diasPend * 1
    );

    var classificacao = "Baixa";
    if (score >= 200) classificacao = "Crítica";
    else if (score >= 80) classificacao = "Alta";
    else if (score >= 30) classificacao = "Média";

    return {
      loja: s.loja,
      time: s.time || "N/D",

      qtde: s.qtde,
      valor: s.valor,
      diasComPendencia: diasPend,
      pendEtiqueta: s.pendEtiqueta,
      pendNF: s.pendNF,
      pendDesc: s.pendDesc,

      qtdeSnapshots: qtdSnaps,

      trend14: r14, // {ult14, ant14, delta}
      score: score,
      classificacao: classificacao,
      txCount: (s.txCount || 0),
    };
  });

  // ranking por qtde e por valor
  rows.sort(function(a,b){
    if (b.qtde !== a.qtde) return b.qtde - a.qtde;
    return b.valor - a.valor;
  });

  return { ok: true, rows: rows, janelaDias: diasJanela };
}

function calcularTrend14dias_(histData, lojaKey) {
  var hoje = new Date();
  var d0 = new Date(hoje.getTime() - 14 * 24*60*60*1000);
  var d1 = new Date(hoje.getTime() - 28 * 24*60*60*1000);

  var ult14 = 0;
  var ant14 = 0;

  histData.forEach(function(r){
    var dtSnap = r[0] instanceof Date ? r[0] : new Date(r[0]);
    if (!(dtSnap instanceof Date) || isNaN(dtSnap.getTime())) return;

    var loja = String(r[1] || "").trim() || "(N/D)";
    if (loja !== lojaKey) return;

    var qt = Number(r[7]) || 0;

    if (dtSnap >= d0) ult14 += qt;
    else if (dtSnap >= d1 && dtSnap < d0) ant14 += qt;
  });

  var deltaAbs = ult14 - ant14;

  // Percentual: só faz sentido se ant14 > 0
  // Se ant14 == 0 e ult14 > 0, é “novo” (sem base comparativa)
  var deltaPct = null;
  if (ant14 > 0) {
    deltaPct = (deltaAbs / ant14) * 100; // ex.: +241.6 (%)
  }

  // ✅ Mantém "delta" por compatibilidade com o resto do código
  return {
    ult14: ult14,
    ant14: ant14,
    delta: deltaAbs,       // compatibilidade
    deltaAbs: deltaAbs,    // novo (opção C)
    deltaPct: deltaPct     // novo (opção C) -> número ou null
  };
}

function montarEmailOfensorasPendencias_(rel) {
  var rows = (rel && rel.rows) ? rel.rows.slice() : [];
  var periodo = (rel && rel.periodo) ? rel.periodo : {};
  var meta = (rel && rel.meta) ? rel.meta : {};

  function esc(s){
    return String(s || "").replace(/[&<>"']/g, function(m){
      return ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[m]);
    });
  }
  function fmtMoney(v){
    try {
      return (Number(v)||0).toLocaleString("pt-BR", { style:"currency", currency:"BRL" });
    } catch(e) {
      return "R$ " + (Math.round((Number(v)||0)*100)/100).toString().replace(".", ",");
    }
  }
  function fmtNum(v){ return (Number(v)||0).toLocaleString("pt-BR"); }

  // ========= TOP (por quantidade) =========
  var top = rows.slice().sort(function(a,b){
    if ((b.qtde||0) !== (a.qtde||0)) return (b.qtde||0) - (a.qtde||0);
    return (b.valor||0) - (a.valor||0);
  }).slice(0, 15);

  // Set para impedir duplicidade nas futuras
  var topSet = {};
  top.forEach(function(r){ topSet[String(r.loja||"")] = true; });

  // ========= FUTURAS (por aceleração 14d), EXCLUINDO TOP =========
  var futuras = rows
    .filter(function(r){
      var loja = String(r.loja||"");
      if (topSet[loja]) return false;               // ✅ nunca pode repetir
      var d = Number(r.delta14)||0;
      if (d <= 0) return false;                     // só aceleração positiva
      // corte anti-ruído: evita “futuras” com volume muito baixo
      if ((Number(r.qtde)||0) < 5) return false;
      return true;
    })
    .sort(function(a,b){
      if ((b.delta14||0) !== (a.delta14||0)) return (b.delta14||0) - (a.delta14||0);
      return (b.score||0) - (a.score||0);
    })
    .slice(0, 10);

  // ========= Texto analítico: “por quê” =========
  function principalFalha_(r){
    var e = Number(r.pendEtiqueta)||0;
    var n = Number(r.pendNF)||0;
    var d = Number(r.pendDesc)||0;
    var total = e+n+d;
    if (!total) return "sem detalhamento por tipo";

    var arr = [
      {k:"NF/Recibo", v:n},
      {k:"Etiqueta", v:e},
      {k:"Descrição", v:d}
    ].sort(function(a,b){ return b.v-a.v; });

    var pct = Math.round((arr[0].v/total)*100);
    return arr[0].k + " (" + pct + "%)";
  }

  function linhaAnalitica_(r){
    var snaps = (r.qtdeSnapshots == null ? "—" : fmtNum(r.qtdeSnapshots));
    var score = (r.score == null ? "—" : Number(r.score).toFixed(1));
    return "<li><b>" + esc(r.loja) + "</b> (" + esc(r.time||"N/D") + ", " + esc(r.classificacao||"—") + "): " +
      "Qtde " + fmtNum(r.qtde) +
      ", principal falha: <b>" + esc(principalFalha_(r)) + "</b>" +
      ", #Snapshots " + esc(snaps) +
      ", Δ14d " + fmtNum(r.delta14||0) +
      ", Score " + esc(score) + ".</li>";
  }

  // ========= Tabela com as MESMAS colunas do chat =========
  function tabela_(titulo, lista, headerBg) {
    var th = "border:1px solid #cbd5e1;padding:6px;font-size:12px;white-space:nowrap;color:#fff;background:" + headerBg + ";";
    var td = "border:1px solid #cbd5e1;padding:6px;font-size:12px;white-space:nowrap;";
    var thMetric = "border:1px solid #cbd5e1;padding:6px;font-size:12px;white-space:nowrap;color:#0d0c0c;background:#e6e605;";

    var html = "";
    html += "<h3 style='margin:16px 0 8px 0'>" + esc(titulo) + "</h3>";
    html += "<table style='border-collapse:collapse;width:100%'>";
    html += "<tr>" +
      "<th style='" + th + "text-align:left'>Loja</th>" +
      "<th style='" + th + "text-align:left'>Time</th>" +
      "<th style='" + th + "text-align:center'>Qtde</th>" +
      "<th style='" + th + "text-align:right'>Valor (R$)</th>" +
      "<th style='" + th + "text-align:center'>Dias c/ pendência</th>" +
      "<th style='" + th + "text-align:center'>Etiqueta</th>" +
      "<th style='" + th + "text-align:center'>NF/Recibo</th>" +
      "<th style='" + th + "text-align:center'>Descrição</th>" +
      "<th style='" + thMetric + "text-align:center'># Snapshots</th>" +
      "<th style='" + thMetric + "text-align:center'>Variação - Δ 14d</th>" +
      "<th style='" + thMetric + "text-align:center'>% Var Δ 14d</th>" +
      "<th style='" + thMetric + "text-align:center'>Score</th>" +
      "<th style='" + thMetric + "text-align:left'>Classificação</th>" +
    "</tr>";

          (lista || []).forEach(function(r){
        var dAbs = Number(r.delta14 || 0);
        var dPct = (r.delta14Pct != null ? Number(r.delta14Pct) : null);

        var dAbsTxt = (dAbs > 0 ? "+" : "") + fmtNum(dAbs);
        var dPctTxt = (dPct == null) ? (dAbs > 0 ? "novo" : "—") : ((dPct > 0 ? "+" : "") + dPct.toFixed(0) + "%");
        var variacaoTxt = dAbsTxt + " (" + dPctTxt + ")";
        var pctColTxt = dPctTxt;

        html += "<tr>" +
          "<td style='" + td + "'>" + esc(r.loja) + "</td>" +
          "<td style='" + td + "'>" + esc(r.time || "N/D") + "</td>" +
          "<td style='" + td + "text-align:center'>" + fmtNum(r.qtde) + "</td>" +
          "<td style='" + td + "text-align:right'>" + fmtMoney(r.valor) + "</td>" +
          "<td style='" + td + "text-align:center'>" + fmtNum(r.diasComPendencia) + "</td>" +
          "<td style='" + td + "text-align:center'>" + fmtNum(r.pendEtiqueta) + "</td>" +
          "<td style='" + td + "text-align:center'>" + fmtNum(r.pendNF) + "</td>" +
          "<td style='" + td + "text-align:center'>" + fmtNum(r.pendDesc) + "</td>" +
          "<td style='" + td + "text-align:center'>" + (r.qtdeSnapshots == null ? "—" : fmtNum(r.qtdeSnapshots)) + "</td>" +

          // ✅ agora sim opção C aparece no email
          "<td style='" + td + "text-align:center'>" + esc(variacaoTxt) + "</td>" +
          "<td style='" + td + "text-align:center'>" + esc(pctColTxt) + "</td>" +
          "<td style='" + td + "text-align:center'>" + (r.score == null ? "—" : Number(r.score).toFixed(1)) + "</td>" +
          "<td style='" + td + "font-weight:700'>" + esc(r.classificacao || "—") + "</td>" +
        "</tr>";
      });


    html += "</table>";
    return html;
  }

  // ========= Montagem do email =========
  var html = "";
  html += "<div style='font-family:Arial,sans-serif;font-size:13px;color:#0f172a'>";
  html += "<h2 style='margin:0 0 6px 0'>Lojas ofensoras (pendências de justificativas)</h2>";

  html += "<p style='margin:0 0 10px 0'>";
  if (periodo.inicio || periodo.fim) {
    html += "<b>Período:</b> " + esc(periodo.inicio||"") + " a " + esc(periodo.fim||"") + " | ";
  }
  html += "<b>Janela:</b> Ciclo atual";
  html += "</p>";

  html += "<p style='margin:0 0 12px 0;color:#334155'>";
  html += "Top ofensoras = maior volume de pendências. Futuras ofensoras = aceleração recente (Δ14d) ";
  html += "com exclusão automática das lojas que já estão no Top (para evitar duplicidade).";
  html += "</p>";

  html += "<h3 style='margin:16px 0 8px 0'>Principais lojas ofensoras e por quê:</h3>";

  html += "<p style='margin:0 0 10px 0;font-size:13px;color:#334155;line-height:1.35;'>" +
  "<b>Como ler os indicadores:</b> " +
  "<b>Qtde</b> = total de pendências no período; " +
  "<b>Principal falha</b> = o tipo de pendência mais frequente na loja (Etiqueta, NF/Recibo ou Descrição) e o percentual indica a participação desse tipo no total de pendências da loja (ex.: <b>Descrição (81%)</b> significa que 81% das pendências são por falta/erro de descrição); " +
  "<b>#Snapshots</b> = Em quantas coletas diferentes a loja apareceu com pendência (proxy de recorrência ao longo do período); " +
  "<b>Δ14d</b> = variação do total de pendências nos últimos 14 dias versus os 14 dias anteriores; " +
  "<b>Score</b> = índice composto usado para priorização (combina volume, tipo de falha e recorrência); " +
  "<b>Classificação</b> = faixa do Score (Baixa/Média/Alta/Crítica)." +
"</p>";

  html += "<ul style='margin:0 0 12px 18px'>";
  top.slice(0, 5).forEach(function(r){ html += linhaAnalitica_(r); });
  html += "</ul>";

  html += tabela_("Top ofensoras (por quantidade)", top, "#0b2a57");

  html += "<h3 style='margin:16px 0 8px 0'>Prováveis futuras ofensoras (aceleração Δ 14d)</h3>";
  if (!futuras.length) {
    html += "<p style='margin:0'>Sem destaques de aceleração no período (ou todas já estão no Top).</p>";
  } else {
    html += tabela_("Futuras ofensoras (sem duplicidade)", futuras, "#8a6b00");
  }

  html += "<p style='margin:16px 0 0 0;color:#475569'>Base: Histórico consolidado de transações da Clara.</p>";
  html += "</div>";
  return html;
}

function enviarEmailOfensorasPendenciasClara(diasJanela) {
  try {
    // Segurança: só Admin (mesmo padrão do e-mail de limites)
    var email = Session.getActiveUser().getEmail();
    if (!isAdminEmail(email)) {
      return { ok: false, error: "Acesso restrito: apenas Administrador pode disparar esse relatório." };
    }

    var rel = getLojasOfensorasParaChat(diasJanela || 60);
    if (!rel || !rel.ok) return { ok: false, error: (rel && rel.error) ? rel.error : "Falha no relatório." };

    var destinatarios = getAdminEmails_();
    if (!destinatarios.length) return { ok: false, error: "Lista de admins vazia." };

    // --- ASSUNTO: trocar "60d" por "Ciclo atual" ---
    var assunto = "📌 [ALERTA CLARA | JUSTIFICATIVAS] Lojas ofensoras (Ciclo atual)";

    if (rel && rel.periodo && rel.periodo.inicio && rel.periodo.fim) {
      assunto += " | " + rel.periodo.inicio + " a " + rel.periodo.fim;
    }

var html = montarEmailOfensorasPendencias_(rel);


    GmailApp.sendEmail(destinatarios.join(","), assunto, " ", {
      from: "vektor@gruposbf.com.br",
      htmlBody: html,
      name: "Vektor - Grupo SBF"
    });

    registrarAlertaEnviado_(
  "PENDENCIAS",
  "",
  "",
  "Envio do relatório de lojas ofensoras (janela " + ((diasJanela || 60)) + "d). Total lojas=" + ((rel.rows || []).length),
  destinatarios.join(","),
  "enviarEmailOfensorasPendenciasClara"
);

    return { ok: true, sent: true, msg: "E-mail enviado para admins.", totalLojas: (rel.rows||[]).length };

  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

/**
 * Porteiro: só dispara envio se a aba BaseClara mudou de fato.
 * Coloque o trigger nesta função (não direto no enviarAlertasLimitesClaraDiario).
 */

function ENVIAR_EMAIL_LIMITE_CLARA() {
  var props = PropertiesService.getScriptProperties();

  // 1) calcula assinatura atual da BaseClara
  var sigAtual = calcularAssinaturaBaseClara_();
  if (!sigAtual || sigAtual.error) {
    Logger.log("Falha ao calcular assinatura BaseClara: " + (sigAtual && sigAtual.error ? sigAtual.error : sigAtual));
    return;
  }

  var keySig = "VEKTOR_SIG_BASECLARA_PROCESSADA";
  var sigAnterior = props.getProperty(keySig) || "";

  // 2) se não mudou, não envia
  if (sigAtual.sig === sigAnterior) {
    Logger.log("BaseClara não mudou desde a última verificação. Não envia alertas.");
    return;
  }

  // 3) janela de segurança opcional (para evitar disparar enquanto carga ainda está em andamento)
  // Se você não quiser atraso, pode remover este bloco inteiro.
  var AGUARDAR_MIN = 10; // ajuste aqui (10–20 costuma ser bom)
  var agora = new Date();
  var diffMin = (agora.getTime() - sigAtual.maxDataMs) / 60000;

  // Só aplica a janela se maxData veio preenchida (quando a coluna Data é válida)
  if (sigAtual.maxDataMs > 0 && diffMin >= 0 && diffMin < AGUARDAR_MIN) {
    Logger.log("BaseClara mudou, mas ainda dentro da janela de segurança (" + diffMin.toFixed(1) + " min).");
    return;
  }

  // 4) marca assinatura como processada e dispara envio
  props.setProperty(keySig, sigAtual.sig);

  // ✅ NOVO: snapshot só quando BaseClara mudou
try {
  var snap = REGISTRAR_SNAPSHOT();
  if (snap && snap.ok) {
    Logger.log("Snapshot pendências gravado. Linhas: " + (snap.gravados || 0));
  } else {
    Logger.log("Snapshot pendências falhou: " + (snap && snap.error ? snap.error : snap));
  }
} catch (e) {
  Logger.log("Snapshot pendências - erro: " + (e && e.message ? e.message : e));
}

  Logger.log("BaseClara mudou (sig anterior ≠ atual). Enviando alertas...");
  enviarAlertasLimitesClaraDiario();
}

// ==============================
// USO IRREGULAR (CONSERVADOR) - BASECLARA
// ==============================

function getPossivelUsoIrregularParaChat(modo) {
  // 🔒 RBAC por VEKTOR_ACESSOS (ROLE -> FUNCTION_ALLOW)
  vektorAssertFunctionAllowed_("getPossivelUsoIrregularParaChat");
  try {
    modo = (modo || "7d").toString().toLowerCase().trim(); // default 7 dias

    var rel = detectarUsoIrregularBaseClara_({ modo: modo });
    return rel;

  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

function getRadarIrregularidadeParaChat(modo) {
  vektorAssertFunctionAllowed_("getRadarIrregularidadeParaChat");
  try {

    modo = (modo || "7d").toString().toLowerCase().trim();

    // Reaproveita exatamente a mesma base do fluxo atual
    var rel = detectarUsoIrregularBaseClara_({ modo: modo });
    if (!rel || !rel.ok) return rel;

    var rows = Array.isArray(rel.rows) ? rel.rows : [];

    // ✅ Detalhe por loja: 1 linha “mais crítica” por loja (para tooltip no ranking)
var detalhesPorLoja = {};
function toNumberMoney_(v) {
  if (v === null || v === undefined) return 0;
  if (typeof v === "number") return v;
  var s = String(v);
  s = s.replace(/[^\d,.-]/g, "").replace(/\./g, "").replace(",", ".");
  var n = Number(s);
  return isFinite(n) ? n : 0;
}

rows.forEach(function (r) {
  var loja = String(r.loja || "").trim();
  if (!loja) return;

  var cur = detalhesPorLoja[loja];
  if (!cur) { detalhesPorLoja[loja] = r; return; }

  var sA = Number(r.score) || 0;
  var sB = Number(cur.score) || 0;
  if (sA !== sB) { if (sA > sB) detalhesPorLoja[loja] = r; return; }

  var vA = toNumberMoney_(r.valor);
  var vB = toNumberMoney_(cur.valor);
  if (vA !== vB) { if (vA > vB) detalhesPorLoja[loja] = r; return; }

  var sdA = toNumberMoney_(r.somaDia);
  var sdB = toNumberMoney_(cur.somaDia);
  if (sdA !== sdB) { if (sdA > sdB) detalhesPorLoja[loja] = r; return; }

  var qA = Number(r.qtdDia) || 0;
  var qB = Number(cur.qtdDia) || 0;
  if (qA > qB) detalhesPorLoja[loja] = r;
});

    var mapa = {}; // loja -> agg

    function toNumberMoney_(s) {
      // "R$ 1.980,00" -> 1980.00
      if (s == null) return 0;
      var t = String(s).replace(/[^\d,.-]/g, "");
      // troca separador milhar/ponto
      // casos comuns: "1.980,00" -> remove "." e troca "," por "."
      t = t.replace(/\./g, "").replace(",", ".");
      var n = Number(t);
      return isNaN(n) ? 0 : n;
    }

    function pendCount_(txt) {
      // "Sim (2)" -> 2; "Não" -> 0
      var s = String(txt || "").toLowerCase();
      var m = s.match(/\((\d+)\)/);
      if (m) return Number(m[1]) || 0;
      return s.indexOf("sim") >= 0 ? 1 : 0;
    }

    rows.forEach(function (r) {
      var loja = String(r.loja || "").trim();
      if (!loja) return;

      if (!mapa[loja]) {
        mapa[loja] = {
          loja: loja,
          time: String(r.time || "").trim(),
          casos: 0,
          scoreSum: 0,
          scoreMax: 0,
          qtdDiaSum: 0,
          valorMax: 0,
          pendEventos: 0,
          pendCountSum: 0
        };
      }

      var a = mapa[loja];

      var score = Number(r.score) || 0;
      var qtdDia = Number(r.qtdDia) || 0;
      var valor = toNumberMoney_(r.valor); // no seu retorno, "valor" já vem como money_()
      var pendCount = pendCount_(r.pendenciasTxt);

      a.casos += 1;
      a.scoreSum += score;
      if (score > a.scoreMax) a.scoreMax = score;

      a.qtdDiaSum += qtdDia;
      if (valor > a.valorMax) a.valorMax = valor;

      if (pendCount > 0) a.pendEventos += 1;
      a.pendCountSum += pendCount;

      // mantém último time não vazio
      if (!a.time && r.time) a.time = String(r.time || "").trim();
    });

    var lojas = Object.keys(mapa).map(function (k) {
      var a = mapa[k];
      var avgScore = a.casos ? (a.scoreSum / a.casos) : 0;
      var avgQtdDia = a.casos ? (a.qtdDiaSum / a.casos) : 0;
      var pendRate = a.casos ? (a.pendEventos / a.casos) : 0;

      return {
        loja: a.loja,
        time: a.time || "N/D",
        casos: a.casos,
        avgScore: Number(avgScore.toFixed(2)),
        maxScore: a.scoreMax,
        avgQtdDia: Number(avgQtdDia.toFixed(2)),
        maxValor: Number(a.valorMax.toFixed(2)),
        pendRate: Number(pendRate.toFixed(3)),
        pendCountSum: a.pendCountSum
      };
    });

    // ordena por “proximidade” simples (maxScore e casos)
    lojas.sort(function (x, y) {
      if (y.maxScore !== x.maxScore) return y.maxScore - x.maxScore;
      return y.casos - x.casos;
    });

    return {
      ok: true,
      meta: rel.meta || {},
      lojas: lojas,
      detalhesPorLoja: detalhesPorLoja
    };
  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

/**
 * Porteiro (igual ao limite): só roda se BaseClara mudou.
 * Coloque o gatilho nesta função.
 */
function ENVIAR_EMAIL_USO_IRREGULAR_CLARA() {
  var props = PropertiesService.getScriptProperties();

  // 1) assinatura da BaseClara (reaproveita o mesmo método do limite)
  var sigAtual = calcularAssinaturaBaseClara_();
  if (!sigAtual || sigAtual.error) {
    Logger.log("Falha ao calcular assinatura BaseClara: " + (sigAtual && sigAtual.error ? sigAtual.error : sigAtual));
    return;
  }

  var keySig = "VEKTOR_SIG_BASECLARA_IRREGULAR";
  var sigAnterior = props.getProperty(keySig) || "";

  if (sigAtual.sig === sigAnterior) {
    Logger.log("BaseClara não mudou desde a última verificação (uso irregular). Não envia.");
    return;
  }

  // 2) processa e envia
  var rel = detectarUsoIrregularBaseClara_();
  if (!rel || !rel.ok) {
    Logger.log("Relatório uso irregular falhou: " + (rel && rel.error ? rel.error : rel));
    return;
  }

  // 3) anti-spam por ciclo (igual seus padrões)
  var cicloKey = getCicloKey06a05_(); // já existe no seu arquivo :contentReference[oaicite:3]{index=3}
  var sentKey = "VEKTOR_IRREGULAR_SENT_" + cicloKey;

  // Se não tem alertas, atualiza assinatura e sai
  if (!rel.rows || rel.rows.length === 0) {
    props.setProperty(keySig, sigAtual.sig);
    props.deleteProperty(sentKey);
    Logger.log("Sem casos de uso irregular no ciclo. OK.");
    return;
  }

  // Se já enviou neste ciclo, não manda de novo
  if (props.getProperty(sentKey) === "1") {
    props.setProperty(keySig, sigAtual.sig);
    Logger.log("Uso irregular já enviado neste ciclo. Não reenvia.");
    return;
  }

  var envio = enviarEmailUsoIrregularClara_(rel);
  if (envio && envio.ok) {
    props.setProperty(sentKey, "1");
    props.setProperty(keySig, sigAtual.sig);
  }
}

function enviarEmailUsoIrregularClara_(rel) {
  // 🔒 RBAC por ROLE (VEKTOR_ACESSOS)
  vektorAssertFunctionAllowed_("enviarEmailUsoIrregularClara_");
  try {
    var destinatarios = getAdminEmails_(); // já existe no seu arquivo :contentReference[oaicite:4]{index=4}
    if (!destinatarios || !destinatarios.length) {
      return { ok: false, error: "Lista de admins vazia." };
    }

    var assunto = "📌 [ALERTA CLARA | POSSÍVEL USO IRREGULAR] " +
      (rel.meta && rel.meta.periodo ? ("| " + rel.meta.periodo) : "");

    // Tabela HTML (resumo)
    var html = montarEmailUsoIrregular_(rel);

    GmailApp.sendEmail(destinatarios.join(","), assunto, " ", {
      from: "vektor@gruposbf.com.br",
      htmlBody: html,
      name: "Vektor - Grupo SBF"
    });

    registrarAlertaEnviado_(
  "USO_IRREGULAR",
  "", // não é alerta por loja única (é consolidado)
  "",
  "Possível uso irregular (modelo conservador). Casos=" + ((rel.rows || []).length) +
    (rel.meta && rel.meta.periodo ? (" | " + rel.meta.periodo) : ""),
  destinatarios.join(","),
  "enviarEmailUsoIrregularClara_"
);

    return { ok: true, sent: true, total: (rel.rows || []).length };

  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

function montarEmailUsoIrregular_(rel) {
  function esc_(x){
    return String(x===null||x===undefined?"":x)
      .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
      .replace(/"/g,"&quot;").replace(/'/g,"&#039;");
  }
  function th_(t){
    return "<th style='background:#B91C1C;color:#fff;border:1px solid #0f172a;padding:6px;font-size:12px;white-space:nowrap;'>" + esc_(t) + "</th>";
  }
  function td_(t){
    return "<td style='border:1px solid #0f172a;padding:6px;font-size:12px;white-space:nowrap;vertical-align:top;'>" + esc_(t) + "</td>";
  }

  var rows = rel.rows || [];
  var top = rows.slice(0, 60); // evita e-mail gigante

  var html = "";
  html += "<p>Identificamos <b>padrões atípicos</b> que requerem validação (modelo conservador; 2+ critérios).</p>";
  html += "<p style='font-size:12px;color:#475569;'>" +
        "Critérios fortes: fracionamento, pendência + valor alto, recorrência anormal por estabelecimento/cartão. " +
        "Critérios auxiliares (quando já houver 2+ fortes): etiqueta rara por loja e novo estabelecimento." +
        "</p>";

  html += "<table style='border-collapse:collapse;width:100%;font-family:Arial,sans-serif;'>";
  html += "<thead><tr>";
  ["Loja","Time","Data","Cartão","Estabelecimento","Qtd (dia)","Soma (dia)","Valor (R$)","Pendências","Score","Regras"].forEach(function(h){
    html += th_(h);
  });
  html += "</tr></thead><tbody>";

  top.forEach(function(r){
    html += "<tr>";
    html += td_(r.loja);
    html += td_(r.time);
    html += td_(r.data);
    html += td_(r.cartao);
    html += td_(r.estabelecimento);
    html += td_(r.qtdDia);
    html += td_(r.somaDia);
    html += td_(r.valor);
    html += td_(r.pendenciasTxt);
    html += td_(r.score);
    html += td_(r.regrasTxt);
    html += "</tr>";
  });

  html += "</tbody></table>";
  html += "<br/><p><b>Vektor - Contas a Receber</b></p>";
  return html;
}

function detectarUsoIrregularBaseClara_(opts) {
  var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
  var sh = ss.getSheetByName("BaseClara");
  if (!sh) return { ok: false, error: "Aba BaseClara não encontrada." };

  var lastRow = sh.getLastRow();
  if (lastRow < 2) return { ok: true, rows: [], meta: { periodo: "" } };

  var values = sh.getRange(2, 1, lastRow - 1, 23).getValues(); // A..W = 23 cols

  // Índices zero-based (A..W)
  var IDX_DATA   = 0;   // A
  var IDX_TRANS  = 2;   // C (estabelecimento)
  var IDX_VALOR  = 5;   // F (R$)
  var IDX_CARTAO = 6;   // G (4 dígitos)
  var IDX_AUT    = 12;  // M (cód. autorização)
  var IDX_RECIBO = 14;  // O
  var IDX_TITULAR = 16;
  var IDX_GRUPO  = 17;  // R
  var IDX_ETIQ   = 19;  // T
  var IDX_DESC   = 20;  // U
  var IDX_LOJA   = 21;  // V

  // ------------------------------
  // ✅ Janela de análise por "modo"
  // ------------------------------
  opts = opts || {};
  var modo = String(opts.modo || "ciclo").toLowerCase().trim(); // default: ciclo (compatível)

  var tz = "America/Sao_Paulo";
  var ini = null;
  var fim = null;
  var periodoLabel = "";

  if (modo === "7d") {
    // últimos 7 dias (inclui hoje)
    var hoje = new Date();
    var hoje0 = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());
    ini = new Date(hoje0.getFullYear(), hoje0.getMonth(), hoje0.getDate() - 6);
    fim = hoje0;
    periodoLabel = Utilities.formatDate(ini, tz, "dd/MM/yyyy") + " a " + Utilities.formatDate(fim, tz, "dd/MM/yyyy");
  } else if (modo === "full") {
    // base toda
    ini = null;
    fim = null;
    periodoLabel = "Base toda (investigação)";
  } else {
    // ciclo 06–05 (padrão atual)
    var pc = getPeriodoCicloClara_();
    ini = pc.inicio;
    fim = pc.fim;
    periodoLabel = Utilities.formatDate(ini, tz, "dd/MM/yyyy") + " a " + Utilities.formatDate(fim, tz, "dd/MM/yyyy");
  }

  function norm_(s){
    return normalizarTexto_ ? normalizarTexto_(s) : String(s||"").toLowerCase();
  }
  function money_(n){
    var v = Number(n) || 0;
    return v.toLocaleString("pt-BR", { style:"currency", currency:"BRL" });
  }
  function isPendencia_(row){
    // conservador: pendência se vazio OU “sim”
    function p(v){
      var s = String(v||"").trim().toUpperCase();
      return (!s || s === "SIM");
    }
    return p(row[IDX_RECIBO]) || p(row[IDX_ETIQ]) || p(row[IDX_DESC]);
  }
  function pendTxt_(row){
    var p = [];
    function chk(v, nome){
      var s = String(v||"").trim().toUpperCase();
      if (!s || s === "SIM") p.push(nome);
    }
    chk(row[IDX_RECIBO], "Recibo");
    chk(row[IDX_ETIQ],   "Etiqueta");
    chk(row[IDX_DESC],   "Descrição");
    return p.join(", ");
  }

  // ========== 1) Indexação por (data|cartão|estab) para fracionamento ==========
  var gruposDia = {}; // key -> { loja,time, dataKey, cartao, estab, qtd, soma, maxValor, pendCount }

  // ========== 2) Stats auxiliares ==========
  var valoresJanela = [];     // percentil 95 dentro da janela (7d/ciclo/full)
  var byCartaoEstab = {};     // cartao||estab -> count + pend

    // ------------------------------
  // NOVO: estatísticas por loja (janela atual)
  // - para "Etiqueta rara por loja"
  // - para "Novo estabelecimento por loja"
  // ------------------------------
  var byLojaTotal = {};     // loja -> total trans na janela
  var byLojaEtiq  = {};     // loja -> { etiquetaNorm: count }
  var byLojaEstab = {};     // loja -> { estabNorm: count }

  function normKey_(s) {
    // normalização conservadora (evita estourar falso-positivo por variações pequenas)
    return String(s || "")
      .toUpperCase()
      .normalize("NFD").replace(/[\u0300-\u036f]/g, "")
      .replace(/\s+/g, " ")
      .trim();
  }

  function splitEtiqs_(raw) {
    var s = String(raw || "").trim();
    if (!s) return [];
    // separadores comuns: vírgula, ponto e vírgula, barra vertical
    var parts = s.split(/[;,|]/g).map(function(x){ return normKey_(x); }).filter(Boolean);

    // fallback: se vier uma etiqueta única sem separador
    if (!parts.length) {
      var one = normKey_(s);
      return one ? [one] : [];
    }
    return parts;
  }

  for (var i=0;i<values.length;i++){
    var r = values[i];
    if (!r) continue;

    var d = parseDateClara_(r[IDX_DATA]);
    if (!d || isNaN(d.getTime())) continue;

    // aplica janela se existir
    if (ini && d < ini) continue;
    if (fim && d > fim) continue;

    var lojaDigits = String(r[IDX_LOJA]||"").replace(/\D/g,"");
    if (!lojaDigits) continue;
    var loja = ("0000"+lojaDigits).slice(-4);

    var time = String(r[IDX_GRUPO]||"").trim();
    var cartao = String(r[IDX_CARTAO]||"").trim();
    var estab = String(r[IDX_TRANS]||"").trim();

        // NOVO: captura etiqueta(s) e atualiza estatísticas por loja
    var etiqRaw = String(r[IDX_ETIQ] || "").trim();
    var etiqs = splitEtiqs_(etiqRaw);

    var lojaK = loja;
    if (!byLojaTotal[lojaK]) { byLojaTotal[lojaK] = 0; byLojaEtiq[lojaK] = {}; byLojaEstab[lojaK] = {}; }
    byLojaTotal[lojaK]++;

    var estabK2 = normKey_(estab);
    if (estabK2) byLojaEstab[lojaK][estabK2] = (byLojaEstab[lojaK][estabK2] || 0) + 1;

    etiqs.forEach(function(t){
      if (!t) return;
      byLojaEtiq[lojaK][t] = (byLojaEtiq[lojaK][t] || 0) + 1;
    });

    var v = parseNumberSafe_(r[IDX_VALOR]);
    if (!isFinite(v) || v <= 0) continue;

    valoresJanela.push(v);

    // chave dia (dd/MM/yyyy) + cartao + estab
    var dataKey = Utilities.formatDate(d, tz, "dd/MM/yyyy");
    var kDia = dataKey + "||" + cartao + "||" + norm_(estab);

    if (!gruposDia[kDia]) {
      gruposDia[kDia] = {
        loja: loja, time: time, dataKey: dataKey, cartao: cartao, estab: estab,
        qtd: 0, soma: 0, maxValor: 0, pendCount: 0,
        etiqSet: {}
      };
    }

        // NOVO: agrega etiquetas por agrupamento
    if (etiqs && etiqs.length) {
      var gTmp = gruposDia[kDia];
      etiqs.forEach(function(t){
        if (t) gTmp.etiqSet[t] = true;
      });
    }

    gruposDia[kDia].qtd++;
    gruposDia[kDia].soma += v;
    if (v > gruposDia[kDia].maxValor) gruposDia[kDia].maxValor = v;
    if (isPendencia_(r)) gruposDia[kDia].pendCount++;

    var kCE = cartao + "||" + norm_(estab);
    if (!byCartaoEstab[kCE]) byCartaoEstab[kCE] = { count: 0, pend: 0 };
    byCartaoEstab[kCE].count++;
    if (isPendencia_(r)) byCartaoEstab[kCE].pend++;
  }

  // percentil 95 (conservador) dentro da janela escolhida
  valoresJanela.sort(function(a,b){return a-b;});
  var p95 = valoresJanela.length ? valoresJanela[Math.floor(valoresJanela.length*0.95)] : 999999;

  // ========== 3) Gerar alertas com regra 2+ critérios ==========
  var rows = [];

  Object.keys(gruposDia).forEach(function(k){
    var g = gruposDia[k];

    var regras = [];
    var score = 0;

    // Critério A: fracionamento
    if (g.qtd >= 3 && g.soma >= 800) {
      regras.push("Fracionamento (>=3 no dia)");
      score += 40;
    }

    // Critério B: pendência + valor alto (p95 da janela ou >=1500)
    if (g.pendCount > 0 && (g.maxValor >= 1500 || g.maxValor >= p95)) {
      regras.push("Pendência + valor alto");
      score += 25;
    }

    // Critério C: recorrência por cartão+estab na janela (ciclo/7d/full)
    var ce = byCartaoEstab[g.cartao + "||" + norm_(g.estab)];
    if (ce && ce.count >= 8 && ce.pend >= 2) {
      regras.push("Recorrência cartão/estab");
      score += 15;
    }

        // ------------------------------
    // ✅ NOVOS CRITÉRIOS (AUXILIARES)
    // Só entram se JÁ houver 2+ critérios fortes (A/B/C)
    // ------------------------------
    var criteriosFortes = regras.length;

    if (criteriosFortes >= 2) {
      // D) Etiqueta rara por loja (na janela)
      // Regras conservadoras:
      // - exige histórico mínimo na janela (>= 30 trans) para não “inventar rareza”
      // - etiqueta precisa ser MUITO rara: count <= 1 e share <= 2%
      try {
        var totalLoja = byLojaTotal[loja] || 0;
        if (totalLoja >= 30) {
          var etiqCounts = byLojaEtiq[loja] || {};
          var tagsGrupo = g.etiqSet ? Object.keys(g.etiqSet) : [];
          var raras = [];

          tagsGrupo.forEach(function(tag){
            var c = etiqCounts[tag] || 0;
            var share = totalLoja ? (c / totalLoja) : 0;
            if (c > 0 && c <= 1 && share <= 0.02) {
              raras.push(tag);
            }
          });

          if (raras.length) {
            // limita para não poluir regrasTxt
            var show = raras.slice(0, 2).join(", ");
            regras.push("Etiqueta rara (" + show + ")");
            score += 5;
          }
        }
      } catch (_) {}

      // E) Novo estabelecimento por loja (na janela)
      // Regras conservadoras:
      // - exige histórico mínimo na janela (>= 30 trans)
      // - estabelecimento do grupo só aparece 1 vez no período para essa loja
      try {
        var totalLoja2 = byLojaTotal[loja] || 0;
        if (totalLoja2 >= 30) {
          var estabNorm = normKey_(estabelecimento);
          var cEst = (byLojaEstab[loja] && estabNorm) ? (byLojaEstab[loja][estabNorm] || 0) : 0;
          if (cEst === 1) {
            regras.push("Novo estabelecimento");
            score += 5;
          }
        }
      } catch (_) {}
    }

    // Conservador: exige 2 critérios
    if (regras.length < 2) return;

    // Threshold final
    if (score < 50) return;

    rows.push({
      loja: g.loja,
      time: g.time,
      data: g.dataKey,
      cartao: g.cartao,
      estabelecimento: g.estab,
      qtdDia: g.qtd,
      somaDia: money_(g.soma),
      valor: money_(g.maxValor),
      pendenciasTxt: (g.pendCount > 0 ? "Sim (" + g.pendCount + ")" : "Não"),
      score: score,
      regrasTxt: regras.join(" + ")
    });
  });

  // ordenação por score e soma
  rows.sort(function(a,b){
    if (b.score !== a.score) return b.score - a.score;
    // fallback simples (não perfeito, mas mantém compatível com seu retorno atual)
    return (String(b.somaDia).length - String(a.somaDia).length);
  });

  return {
    ok: true,
    rows: rows,
    meta: {
      periodo: periodoLabel,
      p95: p95,
      modo: modo
    }
  };
}

function calcularAssinaturaBaseClara_() {
  try {
    var info = carregarLinhasBaseClara_(); // seu helper
    if (info.error) return { error: info.error };

    var header = info.header || [];
    var linhas = info.linhas || [];
    var lastRow = linhas.length;
    if (lastRow <= 0) return { sig: "EMPTY", maxDataMs: 0, lastRow: 0 };

    // Índices
    var idxAlias = 7; // Alias do Cartão (fixo)
    var idxValor = encontrarIndiceColuna_(header, ["Valor em R$", "Valor (R$)", "Valor", "Total"]);
    var idxData  = encontrarIndiceColuna_(header, ["Data da Transação", "Data Transação", "Data"]);

    if (idxValor < 0) return { error: "Não encontrei a coluna de Valor na BaseClara para assinatura." };
    if (idxData < 0)  return { error: "Não encontrei a coluna de Data na BaseClara para assinatura." };

    // -------------------------------
    // (A) Agregados robustos (varre tudo 1x)
    // -------------------------------
    var maxDataMs = 0;
    var minDataMs = 0;
    var sumCents = 0;         // soma do valor em centavos (inteiro)
    var cntValid = 0;         // qtde de linhas com data válida (para robustez)
    var cntAlias = 0;         // qtde de linhas com alias preenchido (para robustez)

    // Sample determinístico ao longo de toda a base
    // Queremos ~400 pontos no máximo (barato e cobre o dataset)
    var maxSamples = 400;
    var step = Math.max(1, Math.floor(lastRow / maxSamples));
    var samples = [];

    function normDateKey_(dt) {
      var d = (dt instanceof Date) ? dt : new Date(dt);
      if (!(d instanceof Date) || isNaN(d.getTime())) return "";
      // yyyy-MM-dd (estável)
      return d.toISOString().slice(0, 10);
    }

    function normMoneyCents_(v) {
      var n = (typeof v === "number") ? v : Number(String(v || "").replace(",", "."));
      if (!isFinite(n)) return 0;
      return Math.round(n * 100);
    }

    for (var i = 0; i < lastRow; i++) {
      var r = linhas[i];

      var dt = r[idxData];
      var d = (dt instanceof Date) ? dt : new Date(dt);
      if (d instanceof Date && !isNaN(d.getTime())) {
        var ms = d.getTime();
        if (ms > maxDataMs) maxDataMs = ms;
        if (minDataMs === 0 || ms < minDataMs) minDataMs = ms;
        cntValid++;
      }

      var alias = (r[idxAlias] || "").toString().trim();
      if (alias) cntAlias++;

      var cents = normMoneyCents_(r[idxValor]);
      sumCents += cents;

      // amostragem espalhada: pega linha 0, step, 2*step...
      if (i % step === 0) {
        var dKey = normDateKey_(r[idxData]);
        // valor como cents é mais estável do que toFixed / string
        samples.push(alias + "|" + dKey + "|" + cents);
      }
    }

    // -------------------------------
    // (B) Tail (últimas N linhas)
    // -------------------------------
    var N = 250; // 200–500 ok
    var start = Math.max(0, lastRow - N);

    var tailParts = [];
    for (var j = start; j < lastRow; j++) {
      var rr = linhas[j];

      var alias2 = (rr[idxAlias] || "").toString().trim();
      var d2s = normDateKey_(rr[idxData]);
      var cents2 = normMoneyCents_(rr[idxValor]);

      tailParts.push(alias2 + "|" + d2s + "|" + cents2);
    }

    // -------------------------------
    // (C) Payload final e MD5
    // -------------------------------
    // samples cobre o dataset; tail cobre o “fim”; agregados cobrem tendências gerais
    var payload =
      "LR=" + lastRow +
      ";MAX=" + maxDataMs +
      ";MIN=" + minDataMs +
      ";SUM=" + sumCents +
      ";CNTD=" + cntValid +
      ";CNTA=" + cntAlias +
      ";STEP=" + step +
      ";SAMPLE=" + samples.join("||") +
      ";TAIL=" + tailParts.join("||");

    var digest = Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5,
      payload,
      Utilities.Charset.UTF_8
    );

    var sig = digest.map(function (b) {
      var v = (b < 0 ? b + 256 : b).toString(16);
      return v.length === 1 ? "0" + v : v;
    }).join("");

    return { sig: sig, maxDataMs: maxDataMs, lastRow: lastRow };

  } catch (e) {
    return { error: "Falha ao calcular assinatura BaseClara: " + (e && e.message ? e.message : e) };
  }
}

// ================================
// SNAPSHOT PENDÊNCIAS - HISTÓRICO
// ================================
var HIST_PEND_CLARA_RAW = "HIST_PEND_CLARA_RAW";
var PROP_LAST_SNAPSHOT_SIG = "VEKTOR_HISTPEND_LAST_SIG";
var PROP_HISTPEND_CICLO_KEY = "VEKTOR_HISTPEND_CICLO_KEY";

/**
 * Faz snapshot das pendências atuais da BaseClara e grava em HIST_PEND_CLARA_RAW.
 * Recomendado: chamar apenas quando BaseClara foi atualizada (pelo seu gatilho já existente).
 */
function REGISTRAR_SNAPSHOT() {
  try {

    // (1) Lê BaseClara (reaproveite sua forma atual de abrir a planilha BaseClara)
    // Se você já tem BASE_CLARA_ID e nome da aba BaseClara em constantes, use-as.

    var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
    var sh = ss.getSheetByName("BaseClara");
    if (!sh) throw new Error("Aba BaseClara não encontrada.");


    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2) {
      Logger.log(">>> BaseClara sem linhas (lastRow=" + lastRow + ")");
      return { ok: true, msg: "BaseClara sem linhas." };
    }


    var values = sh.getRange(1, 1, lastRow, lastCol).getValues();
    var header = values[0].map(function(h){ return String(h || "").trim(); });
    var rows = values.slice(1);

    Logger.log("Rows lidas da BaseClara: " + rows.length);
    Logger.log("Headers BaseClara: " + JSON.stringify(header));


    // (2) Anti-duplicação por assinatura (evita gravar o mesmo snapshot repetidamente)
    // Reaproveita a mesma lógica: se você já tem uma função calcularAssinaturaBaseClara_(), use-a.
    var sigObj = calcularAssinaturaBaseClara_(); // se sua função exigir args, ajuste
    if (sigObj && sigObj.error) throw new Error(sigObj.error);

    var props = PropertiesService.getScriptProperties();
    var lastSig = props.getProperty(PROP_LAST_SNAPSHOT_SIG) || "";
    //if (sigObj && sigObj.sig && sigObj.sig === lastSig) {
      //return { ok: true, msg: "Snapshot ignorado (assinatura igual à última)." };
    //}

    // (3) Índices por nome de coluna (tolerante a variação)
    function idxOf(possiveis) {
      for (var i = 0; i < possiveis.length; i++) {
        var p = possiveis[i];
        var ix = header.indexOf(p);
        if (ix >= 0) return ix;
      }
      return -1;
    }

    var idxDataTrans  = idxOf(["Data da Transação", "Data Transação", "Data"]);
    var idxValorBRL   = idxOf(["Valor em R$", "Valor (R$)", "Valor"]);
    var idxLojaNum    = idxOf(["LojaNum", "Loja", "Código Loja", "cod_estbl", "cod_loja"]);
    var idxEtiquetas  = idxOf(["Etiquetas"]);
    var idxRecibo     = idxOf(["Recibo"]);
    var idxDescricao  = idxOf(["Descrição", "Descricao"]);

    if (idxDataTrans < 0) throw new Error("Não encontrei a coluna 'Data da Transação' na BaseClara.");
    if (idxValorBRL  < 0) throw new Error("Não encontrei a coluna 'Valor em R$' na BaseClara.");
    if (idxLojaNum   < 0) throw new Error("Não encontrei a coluna 'LojaNum' na BaseClara.");
    if (idxEtiquetas < 0) throw new Error("Não encontrei a coluna 'Etiquetas' na BaseClara.");
    if (idxRecibo    < 0) throw new Error("Não encontrei a coluna 'Recibo' na BaseClara.");
    if (idxDescricao < 0) throw new Error("Não encontrei a coluna 'Descrição' na BaseClara.");

    // (4) Monta linhas pendentes
    // Regra objetiva (do jeito que você descreveu):
    // - Pendencia_etiqueta = 1 se Etiquetas vazia
    // - Pendencia_nf       = 1 se Recibo vazio
    // - Pendencia_descricao= 1 se Descrição vazia
    // - Qtde Total = soma das 3
    var snapshotDate = new Date();
    var out = [];

    function isVazio_(v) {
  if (v === null || v === undefined) return true;
  if (v === false) return true; // IMPORTANTÍSSIMO: checkbox/boolean
  var s = String(v).trim().toLowerCase();

  // placeholders comuns
  if (!s) return true;
  if (s === "-" || s === "—" || s === "n/a" || s === "na") return true;
  if (s === "false" || s === "0") return true;
  if (s === "não" || s === "nao") return true;
  if (s.indexOf("sem recibo") >= 0) return true;
  if (s.indexOf("sem etiqueta") >= 0) return true;

  return false;
}

// ==============================
// RESET AUTOMÁTICO POR CICLO (06→05)
// ==============================
var props = PropertiesService.getScriptProperties();

var cicloKey = getCicloKey06a05_();  // já existe no projeto :contentReference[oaicite:10]{index=10}
var cicloLast = props.getProperty(PROP_HISTPEND_CICLO_KEY) || "";

// se mudou o ciclo, limpa a HIST_PEND_CLARA_RAW (mantém header)
if (cicloKey !== cicloLast) {
  var ssHist = SpreadsheetApp.openById(BASE_CLARA_ID);
  var histSh = ssHist.getSheetByName(HIST_PEND_CLARA_RAW);
  if (!histSh) throw new Error("Aba " + HIST_PEND_CLARA_RAW + " não encontrada.");

  var lr = histSh.getLastRow();
  if (lr >= 2) {
    histSh.getRange(2, 1, lr - 1, histSh.getLastColumn()).clearContent();
  }

  // reseta travas do snapshot
  props.deleteProperty(PROP_LAST_SNAPSHOT_SIG);

  // grava novo ciclo ativo
  props.setProperty(PROP_HISTPEND_CICLO_KEY, cicloKey);

  Logger.log("HIST_PEND_CLARA_RAW resetada para novo ciclo: " + cicloKey);
}

    var ciclo = getPeriodoCicloClaraCompleto_();
    var cicloIni = ciclo.inicio;
    var cicloFim = ciclo.fim;


    for (var r = 0; r < rows.length; r++) {
      var row = rows[r];

      var dt = row[idxDataTrans];
      var loja = String(row[idxLojaNum] || "").trim();
      var valor = Number(row[idxValorBRL]) || 0;

      var etiquetas = String(row[idxEtiquetas] || "").trim();
      var recibo = String(row[idxRecibo] || "").trim();
      var desc = String(row[idxDescricao] || "").trim();

      var pendEtiqueta = isVazio_(etiquetas) ? 1 : 0;
      var pendNF       = isVazio_(recibo)   ? 1 : 0;
      var pendDesc     = isVazio_(desc)     ? 1 : 0;


      var qtde = pendEtiqueta + pendNF + pendDesc;
      if (qtde <= 0) continue; // só grava se houver pendência

      // Guarda data transação como Date se vier string
      var dt2 = (dt instanceof Date) ? dt : (vektorParseDateAny_(dt) || new Date(dt));

      // ✅ filtro do ciclo: só grava transações dentro do 06→05
      if (!(dt2 instanceof Date) || isNaN(dt2.getTime())) continue;
      if (dt2 < cicloIni || dt2 > cicloFim) continue;

      out.push([
        snapshotDate,
        loja,
        dt2,
        valor,
        pendEtiqueta,
        pendNF,
        pendDesc,
        qtde
      ]);
    }

    // (5) Grava na HIST_PEND_CLARA_RAW
    var hist = ss.getSheetByName(HIST_PEND_CLARA_RAW);
    if (!hist) throw new Error("Aba " + HIST_PEND_CLARA_RAW + " não encontrada.");

    if (out.length) {
      hist.getRange(hist.getLastRow() + 1, 1, out.length, out[0].length).setValues(out);
    }

    // (6) Atualiza assinatura salva
    if (sigObj && sigObj.sig) props.setProperty(PROP_LAST_SNAPSHOT_SIG, sigObj.sig);

    Logger.log("Snapshot pendências - linhas geradas: " + out.length);

    return { ok: true, gravados: out.length, msg: "Snapshot gravado com sucesso." };

  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

/**
 * Remove do HIST_PEND_CLARA_RAW todas as linhas cujo Data_snapshot seja a data alvo.
 * A comparação é por "yyyy-MM-dd" no timezone America/Sao_Paulo (ignora hora).
 *
 * @param {Date} dataAlvo
 * @return {object} { ok:true, removidos:n } ou { ok:false, error:"..." }
 */
function REMOVER_SNAPSHOT_POR_DATA_(dataAlvo) {
  try {
    if (!(dataAlvo instanceof Date) || isNaN(dataAlvo.getTime())) {
      throw new Error("Data alvo inválida.");
    }

    var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
    var hist = ss.getSheetByName(HIST_PEND_CLARA_RAW);
    if (!hist) throw new Error("Aba " + HIST_PEND_CLARA_RAW + " não encontrada.");

    var lr = hist.getLastRow();
    if (lr < 2) return { ok: true, removidos: 0, msg: "Histórico vazio." };

    // No seu projeto o histórico tem 8 colunas (A:H)
    var numCols = 8;

    var tz = "America/Sao_Paulo";
    var alvoKey = Utilities.formatDate(dataAlvo, tz, "yyyy-MM-dd");

    var data = hist.getRange(2, 1, lr - 1, numCols).getValues();

    var mantidos = [];
    var removidos = 0;

    for (var i = 0; i < data.length; i++) {
      var r = data[i];
      var dtSnap = (r[0] instanceof Date) ? r[0] : new Date(r[0]);

      // Se não conseguir ler data, mantém (não assume que é do dia alvo)
      if (!(dtSnap instanceof Date) || isNaN(dtSnap.getTime())) {
        mantidos.push(r);
        continue;
      }

      var k = Utilities.formatDate(dtSnap, tz, "yyyy-MM-dd");
      if (k === alvoKey) {
        removidos++;
      } else {
        mantidos.push(r);
      }
    }

    // Reescreve abaixo do cabeçalho (não mexe na linha 1)
    if (lr > 2) {
      hist.getRange(2, 1, lr - 1, numCols).clearContent();
    }
    if (mantidos.length) {
      hist.getRange(2, 1, mantidos.length, numCols).setValues(mantidos);
    }

    return { ok: true, removidos: removidos, msg: "Snapshot(s) removidos do dia " + alvoKey };
  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

/**
 * Apaga os snaps de ONTEM do histórico e regrava rodando REGISTRAR_SNAPSHOT() novamente.
 * Use isso depois que você corrigiu a BaseClara (coluna Loja preenchida).
 */
function REPROCESSAR_SNAPSHOT_ONTEM() {
  var tz = "America/Sao_Paulo";
  var agora = new Date();

  // "ontem" no seu timezone (zerando horário para evitar bordas)
  var ontem = new Date(agora.getFullYear(), agora.getMonth(), agora.getDate() - 1);

  // 1) remove do histórico
  var r1 = REMOVER_SNAPSHOT_POR_DATA_(ontem);
  if (!r1.ok) throw new Error("Falha ao remover snapshot de ontem: " + r1.error);

  // 2) limpa assinatura (trava de anti-duplicação do snapshot)
  // Property usada pelo REGISTRAR_SNAPSHOT: VEKTOR_HISTPEND_LAST_SIG
  var props = PropertiesService.getScriptProperties();
  props.deleteProperty(PROP_LAST_SNAPSHOT_SIG);

  // 3) regrava snapshot (agora com loja correta)
  var r2 = REGISTRAR_SNAPSHOT();
  if (!r2 || !r2.ok) throw new Error("Falha ao regravar snapshot: " + (r2 && r2.error ? r2.error : r2));

  return {
    ok: true,
    removidosOntem: r1.removidos || 0,
    gravadosAgora: r2.gravados || 0,
    msg: "OK: removi ontem e regravei o snapshot."
  };
}

// -------------------------
// Classificação de alertas
// -------------------------
function classificarAlertasLimites_(rows) {
  var risco = [];
  var monitoramento = [];
  var eficiencia = [];
  var admin = [];

  var hoje = new Date();
  var infoCiclo = getPeriodoCicloClara_(); // você já usa no projeto
  var inicio = infoCiclo.inicio;
  var diaDoMes = hoje.getDate();

  // ------------------------------
// Regra de risco combinada por fase do ciclo
// 1ª quinzena: saldo<=500 e %uso>=50%
// depois:      saldo<=500 e %uso>=70%
// ------------------------------
var msDia = 24 * 60 * 60 * 1000;
var diasDesdeInicio = Math.floor((hoje.getTime() - inicio.getTime()) / msDia) + 1;
var limiarPctUsoRisco = (diasDesdeInicio <= 14) ? 0.50 : 0.70;

// saldo crítico fixo (valor), mas combinado com %uso por fase
var saldoCriticoValor = 500;


  // metade do ciclo: regra simples pedida (até dia 15)
  var antesMetade = (diaDoMes <= 15);

  rows.forEach(function(r) {
    var limite = Number(r.limite) || 0;
    var utilizado = Number(r.utilizado) || 0;
    var saldo = Number(r.saldo);
    if (!isFinite(saldo)) saldo = limite - utilizado;

    var proj = Number(r.projecao) || 0;
    var pctProj = (proj > 0) ? (utilizado / proj) : null;
    var pctLim = (limite > 0) ? (utilizado / limite) : null;

    var acao = (r.acao || "").toString().trim();
    var acaoLower = acao.toLowerCase();

    // 1) ADMIN: definir
    if (acaoLower.indexOf("definir") === 0) {
      admin.push(enriquecerRowAlerta_(r, { motivo: "Limite não cadastrado/zerado (ação: Definir)." }));
      return;
    }

    // 2) RISCO (alto)
    var ehAumentar = (acaoLower.indexOf("aumentar") === 0);

    // gatilho de projeção vira MONITORAMENTO (não crítico)
    var monitorPorProj = ehAumentar && (pctProj !== null) && (pctProj + VEKTOR_ALERT_TOL_PCT >= 0.90);

    // risco crítico só por sinais “duros”
    var riscoPorSaldoUso = (saldo <= saldoCriticoValor) && (pctLim !== null) && ((pctLim + VEKTOR_ALERT_TOL_PCT) >= limiarPctUsoRisco);
    var riscoPorAcelerado = false; // desativado (regra nova já cobre)

    if (riscoPorSaldoUso || riscoPorAcelerado) {
    var motivos = [];
    if (riscoPorSaldoUso) motivos.push("Saldo ≤ R$ " + saldoCriticoValor.toFixed(2) + " e %uso ≥ " + Math.round(limiarPctUsoRisco * 100) + "%");
    if (riscoPorAcelerado) motivos.push("Uso acelerado");
    risco.push(enriquecerRowAlerta_(r, { motivo: motivos.join(" | "), pctProj: pctProj, pctLim: pctLim }));
    return;
  }


    if (monitorPorProj) {
    monitoramento.push(enriquecerRowAlerta_(r, {
      motivo: "Ação=Aumentar e %Projeção ≥ 90% (monitoramento, não crítico)",
      pctProj: pctProj,
      pctLim: pctLim
    }));
    return;
  }

    // 3) EFICIÊNCIA (médio)
    var ehReduzir = (acaoLower.indexOf("reduzir") === 0);
    var deltaReducao = extrairDeltaReducao_(acao); // retorna número positivo se "Reduzir -R$ X"
    var eficienciaPorReducao = ehReduzir && deltaReducao >= VEKTOR_ALERT_REDUCAO_MIN;

    // regra de "≤50% por ciclos repetidos" é melhor baseada em histórico, mas aqui deixo sinal simples:
    // se %Projeção existe e está muito baixa no ciclo atual, marca candidato (não “repetido” ainda).
    // Para “repetido”, você pode ligar depois usando soma por ciclos. (Eu deixo preparado no e-mail como "observação".)
    var eficienciaPorPctProj = (pctProj !== null) && (pctProj + VEKTOR_ALERT_TOL_PCT <= 0.50);

    if (eficienciaPorReducao || eficienciaPorPctProj) {
      var motivosEf = [];
      if (eficienciaPorReducao) motivosEf.push("Ação=Reduzir e redução sugerida ≥ R$ " + VEKTOR_ALERT_REDUCAO_MIN.toFixed(2));
      if (eficienciaPorPctProj) motivosEf.push("%Projeção ≤ 50% (avaliar recorrência nos ciclos)");

      eficiencia.push(enriquecerRowAlerta_(r, { motivo: motivosEf.join(" | "), pctProj: pctProj, pctLim: pctLim }));
      return;
    }

    // caso contrário: não alerta
  });

  // Ordenações úteis
  risco.sort(function(a,b){ return (a.saldo||0) - (b.saldo||0); }); // menor saldo primeiro
  eficiencia.sort(function(a,b){ return (b.deltaReducao||0) - (a.deltaReducao||0); }); // maior redução primeiro
  admin.sort(function(a,b){ return (a.nomeCartao||"").localeCompare(b.nomeCartao||""); });

  // corta para evitar e-mails gigantes
  risco = risco.slice(0, VEKTOR_ALERT_MAX_RISCO);
  eficiencia = eficiencia.slice(0, VEKTOR_ALERT_MAX_EFICIENCIA);
  admin = admin.slice(0, VEKTOR_ALERT_MAX_ADMIN);

  monitoramento.sort(function(a,b){ return (b.pctProj||0) - (a.pctProj||0); });
  monitoramento = monitoramento.slice(0, 15);

  return { risco: risco, monitoramento: monitoramento, eficiencia: eficiencia, admin: admin };

}

// -------------------------
// Anti-spam por ciclo (06→05)
// -------------------------
function aplicarAntiSpamCiclo_(cicloKey, packs) {
  var props = PropertiesService.getScriptProperties();
  var raw = props.getProperty("VEKTOR_ALERTS_SENT_" + cicloKey) || "[]";
  var sentKeys = {};
  try {
    JSON.parse(raw).forEach(function(k){ sentKeys[k] = true; });
  } catch(e) {}

  function rowKey(r) {
    // chave estável: cartaoKey + loja + time + tipoAlerta
    var loja = (r.loja || "").toString().trim();
    var time = (r.time || "").toString().trim();
    var cartao = (r.nomeCartao || "").toString().trim();
    return cartao + "||" + loja + "||" + time;
  }

  var enviadosKeys = [];

  function filtrar(lista, tipo) {
    var out = [];
    lista.forEach(function(r){
      var k = rowKey(r) + "||" + tipo;
      if (sentKeys[k]) return;
      out.push(r);
      enviadosKeys.push(k);
    });
    return out;
  }

  var risco = filtrar(packs.risco || [], "risco");
  var monitoramento = filtrar(packs.monitoramento || [], "monitoramento");
  var eficiencia = filtrar(packs.eficiencia || [], "eficiencia");
  var admin = filtrar(packs.admin || [], "admin");

  return { risco: risco, monitoramento: monitoramento, eficiencia: eficiencia, admin: admin, _enviadosKeys: enviadosKeys };

}

function registrarEnviadosCiclo_(cicloKey, keys) {
  if (!keys || !keys.length) return;
  var props = PropertiesService.getScriptProperties();
  var propName = "VEKTOR_ALERTS_SENT_" + cicloKey;

  var raw = props.getProperty(propName) || "[]";
  var arr = [];
  try { arr = JSON.parse(raw) || []; } catch(e) { arr = []; }

  // evita crescer infinito
  keys.forEach(function(k){ arr.push(k); });
  arr = arr.slice(-1000);

  props.setProperty(propName, JSON.stringify(arr));
}

// -------------------------
// Montagem do e-mail
// -------------------------
function montarEmailAlertasLimites_(periodoStr, risco, monitoramento, eficiencia, admin) {
  function money(n){ return (Number(n)||0).toLocaleString("pt-BR",{style:"currency",currency:"BRL"}); }
  function pct(p){ return (p===null || p===undefined) ? "—" : (p*100).toFixed(1)+"%"; }

    var html = "";
    html += "<div style='font-family:Arial,sans-serif;font-size:13px;color:#0f172a;'>";
    html += "<h2 style='margin:0 0 8px 0;'>Alertas de Limites (Clara)</h2>";

    // Dias restantes para o fim do ciclo (06→05)
    // Regra: se hoje é dia 06+ => fecha dia 05 do próximo mês
    //        se hoje é dia 01–05 => fecha dia 05 do mês corrente
    var hoje = new Date();

    // "hoje" normalizado para início do dia (evita erro por horário)
    var hoje0 = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());

    // fim do ciclo: dia 05 no mês correto, às 23:59:59
    var y = hoje0.getFullYear();
    var m = hoje0.getMonth(); // 0-based
    var d = hoje0.getDate();

    var fimCiclo;
    if (d >= 6) {
      // próximo mês
      fimCiclo = new Date(y, m + 1, 5, 23, 59, 59);
    } else {
      // mês atual
      fimCiclo = new Date(y, m, 5, 23, 59, 59);
    }

    var msDia = 24 * 60 * 60 * 1000;
    var diasRestantes = Math.max(0, Math.ceil((fimCiclo.getTime() - hoje0.getTime()) / msDia));

    html += "<p style='margin:0 0 10px 0;'><b>Período do ciclo:</b> " 
    + (periodoStr || "06→05") 
    + " | <b>Dias restantes:</b> " + diasRestantes 
    + "</p>";

    html += "<p style='margin:0 0 14px 0;color:#334155;'>Saldo crítico configurado: <b>" + money(VEKTOR_ALERT_SALDO_CRITICO) + "</b></p>";

  // RISCO
  if (risco && risco.length) {
    html += "<h3 style='margin:16px 0 6px 0;color:#b91c1c;'>🔴 Risco operacional (prioridade alta)</h3>";
    html += "<p style='margin:0 0 8px 0;color:#334155;'><b>Interpretação:</b> Risco elevado de impacto no uso do cartão.<br/><b>Ação recomendada:</b> Se a coluna <b>Ação</b> indicar aumento, priorizar ajuste de limite. Se indicar <b>Manter</b>, tratar como alerta operacional (monitorar consumo/saldo e evitar problemas na utilização).</p>";

    html += tabelaAlertas_(risco, money, pct);
  }

  // MONITORAMENTO

  if (monitoramento && monitoramento.length) {
  html += "<h3 style='margin:16px 0 6px 0;color:#a16207;'>🟡 Monitoramento (não crítico)</h3>";
  html += "<p style='margin:0 0 8px 0;color:#334155;'><b>Interpretação:</b> tendência de consumo próxima do esperado para o ciclo, porém ainda sem sinais críticos.<br/><b>Ação recomendada:</b> acompanhar e antecipar ajuste se necessário.</p>";
  html += tabelaAlertas_(monitoramento, money, pct);
}

  // EFICIÊNCIA
  if (eficiencia && eficiencia.length) {
    html += "<h3 style='margin:16px 0 6px 0;color:#b45309;'>🟠 Eficiência (prioridade média)</h3>";
    html += "<p style='margin:0 0 8px 0;color:#334155;'><b>Interpretação:</b> Limite acima do padrão esperado.<br/><b>Ação recomendada:</b> Avaliar redução para otimização de capital, sem impacto operacional.</p>";
    html += tabelaAlertas_(eficiencia, money, pct);
    html += "<p style='margin:8px 0 0 0;color:#64748b;'><i>Observação:</i> casos com %Projeção baixa devem ser confirmados como recorrentes em 2–3 ciclos antes de redução estrutural.</p>";
  }

  // ADMIN
  if (admin && admin.length) {
    html += "<h3 style='margin:16px 0 6px 0;color:#2563eb;'>🔵 Pendências administrativas</h3>";
    html += "<p style='margin:0 0 8px 0;color:#334155;'><b>Interpretação:</b> cartão com consumo sem limite cadastrado/zerado.<br/><b>Ação recomendada:</b> definir limite na aba Info_limites.</p>";
    html += tabelaAlertas_(admin, money, pct);
  }

  // Rodapé metodológico
  html += "<hr style='margin:16px 0;border:none;border-top:1px solid #e2e8f0;'/>";
  html += "<p style='margin:0;color:#475569;'><b>Metodologia (resumo):</b> Projeção baseada nos últimos 6 ciclos (06→05). Em sazonalidade (Nov/Dez), considera-se cenário conservador para evitar subestimação. Recomendações são heurísticas e devem ser validadas pelo time ADM.</p>";
  html += "</div>";
  return html;
}

function tabelaAlertas_(lista, moneyFn, pctFn) {
  var html = "";
  html += "<table cellpadding='0' cellspacing='0' style='border-collapse:collapse;width:100%;margin-top:6px;'>";
  html += "<tr>";
  html += th_("Loja") + th_("Time") + th_("Cartão") + th_("Limite") + th_("Utilizado") + th_("Saldo") + th_("Projeção") + th_("% Projeção") + th_("Ação") + th_("Motivo");
  html += "</tr>";

  lista.forEach(function(r){
    html += "<tr>";
    html += td_(r.loja || "N/D");
    html += td_(r.time || "N/D");
    html += td_(r.nomeCartao || "N/D");
    html += td_(moneyFn(r.limite));
    html += td_(moneyFn(r.utilizado));
    html += td_(moneyFn(r.saldo));
    html += td_(moneyFn(r.projecao));
    html += td_(pctFn(r.pctProj));
    html += td_((r.acao || "—"));
    html += td_((r.motivo || "—"));
    html += "</tr>";
  });

  html += "</table>";
  return html;

  function th_(t){
    return "<th style='border:1px solid #0f172a;background:#0b2a57;color:#fff;padding:6px;text-align:left;font-size:12px;white-space:nowrap;'>" + esc_(t) + "</th>";
  }
  function td_(t){
    return "<td style='border:1px solid #0f172a;padding:6px;font-size:12px;vertical-align:top;white-space:nowrap;'>" + esc_(t) + "</td>";
  }
  function esc_(x){
    return String(x===null||x===undefined?"":x)
      .replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;")
      .replace(/"/g,"&quot;").replace(/'/g,"&#039;");
  }
}

// -------------------------
// Helpers
// -------------------------
function enriquecerRowAlerta_(r, extra) {
  var out = Object.assign({}, r);
  if (extra) Object.keys(extra).forEach(function(k){ out[k] = extra[k]; });

  // delta de redução útil para ordenação
  out.deltaReducao = extrairDeltaReducao_(out.acao || "");
  return out;
}

function extrairDeltaReducao_(acaoStr) {
  // Espera algo como: "Reduzir -R$ 500,00"
  var s = (acaoStr || "").toString();
  if (s.toLowerCase().indexOf("reduzir") !== 0) return 0;

  // captura números após "-"
  var m = s.match(/-\s*R\$\s*([\d\.\,]+)/i);
  if (!m) return 0;

  var num = m[1].replace(/\./g,"").replace(",",".");
  var v = Number(num);
  return isFinite(v) ? v : 0;
}

function getCicloKey06a05_() {
  // Usa sua regra: se dia 01–05, ciclo começou dia 06 do mês anterior
  var p = getPeriodoCicloClara_();
  var ini = p.inicio;
  var fim = p.fim;
  return Utilities.formatDate(ini, "America/Sao_Paulo", "yyyy-MM-dd") + "_" +
         Utilities.formatDate(fim, "America/Sao_Paulo", "yyyy-MM-dd");
}

function getAdminEmails_() {
  // Reaproveita sua própria lista central via isAdminEmail
  // Se você tiver a lista em outro lugar, adapte aqui.
  // Estratégia: varrer lista conhecida — se você já tem array interno em isAdminEmail, replique.
  var admins = [
    "rodrigo.lisboa@gruposbf.com.br"
    // adicione aqui os outros admins que já existem no isAdminEmail
  ];

  // limpa duplicados
  var seen = {};
  var out = [];
  admins.forEach(function(e){
    var k = (e||"").toLowerCase().trim();
    if (!k || seen[k]) return;
    seen[k] = true;
    out.push(k);
  });
  return out;
}

function chaveCartaoClara_(raw) {
  var s = (raw || "").toString().trim();
  if (!s) return "";

  var norm = normalizarTexto_(s); // seu normalizador atual
  var isVirtual = norm.indexOf("virtual") !== -1;

  // Extrai dígitos
  var dig = s.replace(/\D/g, "");
  if (!dig) return "";

  // Pad para 4 dígitos (se vier 223 vira 0223)
  dig = String(Number(dig)).padStart(4, "0");

  // Chave padrão: ce#### + marcador virtual
  return "ce" + dig + (isVirtual ? "|virtual" : "");
}

// --- Helpers locais (não conflitam com seu projeto) ---

function getPeriodoCicloClara_() {
  var hoje = new Date();
  var y = hoje.getFullYear();
  var m = hoje.getMonth();
  var d = hoje.getDate();

  var inicio;
  if (d >= 6) inicio = new Date(y, m, 6, 0, 0, 0);
  else inicio = new Date(y, m - 1, 6, 0, 0, 0);

  var fim = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate(), 23, 59, 59);
  return { inicio: inicio, fim: fim };
}

function getPeriodoCicloClaraCompleto_() {
  var hoje = new Date();
  var y = hoje.getFullYear();
  var m = hoje.getMonth();
  var d = hoje.getDate();

  // início do ciclo: dia 06 (mês atual se hoje>=6, senão mês anterior)
  var inicio = (d >= 6) ? new Date(y, m, 6, 0, 0, 0) : new Date(y, m - 1, 6, 0, 0, 0);

  // fim do ciclo: dia 05 (mês corrente se hoje<=5, senão próximo mês) às 23:59:59
  var fim = (d >= 6) ? new Date(y, m + 1, 5, 23, 59, 59) : new Date(y, m, 5, 23, 59, 59);

  return { inicio: inicio, fim: fim };
}

function parseNumberSafe_(v) {
  if (v === null || v === undefined || v === "") return 0;
  if (typeof v === "number") return v;
  var s = String(v).trim();
  // aceita "1234,56" e "1.234,56"
  s = s.replace(/\./g, "").replace(",", ".");
  var n = Number(s);
  return isFinite(n) ? n : 0;
}

function formatPeriodoBR_(ini, fim) {
  return {
    inicio: Utilities.formatDate(ini, "America/Sao_Paulo", "dd/MM/yyyy"),
    fim: Utilities.formatDate(fim, "America/Sao_Paulo", "dd/MM/yyyy")
  };
}

function projeçãoCiclo_(ini, fim, totalUsado) {
  try {
    var hoje = new Date();
    var diasDecorridos = Math.max(1, Math.floor((hoje.getTime() - ini.getTime()) / (1000 * 60 * 60 * 24)) + 1);

    // ciclo 06->05 tem ~30/31 dias; projetar até o próximo dia 05
    var prox05 = new Date(ini.getFullYear(), ini.getMonth() + 1, 5, 23, 59, 59);
    var diasCiclo = Math.max(1, Math.floor((prox05.getTime() - ini.getTime()) / (1000 * 60 * 60 * 24)) + 1);

    var mediaDia = (Number(totalUsado) || 0) / diasDecorridos;
    var projFinal = mediaDia * diasCiclo;

    return {
      diasDecorridos: diasDecorridos,
      diasCiclo: diasCiclo,
      mediaDia: mediaDia,
      projFinal: projFinal
    };
  } catch (e) {
    return null;
  }
}

function cartaoKeyCE_(raw) {
  var s = (raw || "").toString();
  var norm = normalizarTexto_(s); // já remove acentos etc.
  if (!norm) return "";

  // captura CE + 4 dígitos em qualquer lugar do texto
  var m = norm.match(/\bce\s*0*(\d{1,4})\b/);
  if (!m) return "";

  var dig = String(Number(m[1] || "")).padStart(4, "0");

  // virtual: aceita "virtual" e também o typo "virual"
  var isVirtual = (norm.indexOf("virtual") !== -1) || (norm.indexOf("virual") !== -1);

  return "ce" + dig + (isVirtual ? "|virtual" : "");
}

function moneyBR_(n) {
  var v = Number(n) || 0;
  // retorna ex: "R$ 1.200"
  return v.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
}

/**
 * Para um determinado grupo/time (opcional) e período,
 * devolve as transações por LOJA com:
 *  - pendências de justificativa (Etiqueta / Descrição vazias ou Recibo = "Não")
 *  - justificativas OK      (Etiqueta e Descrição preenchidas e Recibo = "Sim")
 *
 * É chamada pelo front via google.script.run.getPendenciasEJustificativasPorLojas(...)
 *
 * @param {string} grupo           Nome do time (pode ser vazio)
 * @param {string} dataInicioStr   Data início em ISO (pode ser vazio)
 * @param {string} dataFimStr      Data fim em ISO (pode ser vazio)
 * @param {Array}  lojasFiltro     Lista de códigos de loja (strings). Se vazio, considera todas.
 */
function getPendenciasEJustificativasPorLojas(
  grupo,
  dataInicioStr,
  dataFimStr,
  lojasFiltro
) {
  vektorAssertFunctionAllowed_("getPendenciasEJustificativasPorLojas");
  try {
    var info = carregarLinhasBaseClara_();
    if (info.error) {
      return { ok: false, error: info.error };
    }

    var header = info.header;
    var linhas = info.linhas;

    // Índices fixos já usados no seu getResumoTransacoesPorGrupo
    var IDX_DATA  = 0;   // "Data da Transação"
    var IDX_VALOR = 5;   // "Valor em R$"
    var IDX_GRUPO = 17;  // "Grupos"
    var IDX_LOJA  = 21;  // "LojaNum"

    // Índices dinâmicos para as colunas de justificativa
    var idxRecibo = encontrarIndiceColuna_(header, [
      "Recibo",
      "NF / Recibo",
      "NF/Recibo"
    ]);

    var idxEtiqueta = encontrarIndiceColuna_(header, [
      "Etiquetas",
      "Etiqueta"
    ]);

    var idxDescricao = encontrarIndiceColuna_(header, [
      "Descrição",
      "Descricao",
      "Comentário"
    ]);

    if (idxRecibo < 0 || idxEtiqueta < 0 || idxDescricao < 0) {
      return {
        ok: false,
        error: "Não encontrei as colunas de Recibo/Etiquetas/Descrição na BaseClara."
      };
    }

    // Normaliza grupo (time) informado
    var grupoOriginal = (grupo || "").toString().trim();
    var grupoNorm = normalizarTexto_(grupoOriginal);

    // Normaliza lista de lojas (filtro é opcional)
    var lojasSet = {};
    if (Array.isArray(lojasFiltro)) {
      lojasFiltro.forEach(function (cod) {
        if (!cod) return;
        var c = cod.toString().trim();
        if (c) lojasSet[c] = true;
      });
    }

    // Aplica filtro de período usando função já existente
    var filtradas = filtrarLinhasPorPeriodo_(
      linhas,
      IDX_DATA,
      dataInicioStr,
      dataFimStr
    );

    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";

    var pendencias   = []; // [loja, data, valor, etiqueta, descricao, recibo]
    var justificadas = []; // idem

    filtradas.forEach(function (row) {
      var loja = (row[IDX_LOJA] || "").toString().trim();
      if (!loja) return;

      // Se recebeu filtro de lista de lojas, respeita
      if (Object.keys(lojasSet).length > 0 && !lojasSet[loja]) {
        return;
      }

      // Filtro por grupo/time, se informado
      if (grupoNorm) {
        var grupoLinhaNorm = normalizarTexto_(row[IDX_GRUPO] || "");
        var casaGrupo =
          grupoLinhaNorm.indexOf(grupoNorm) !== -1 ||
          grupoNorm.indexOf(grupoLinhaNorm) !== -1;

        if (!casaGrupo) return;
      }

      // Data da transação formatada
      var d = parseDateClara_(row[IDX_DATA]);
      var dataStr = d
        ? Utilities.formatDate(d, tz, "dd/MM/yyyy")
        : (row[IDX_DATA] || "");

      var valor = Number(row[IDX_VALOR]) || 0;

      var etiqueta  = (row[idxEtiqueta] || "").toString().trim();
      var descricao = (row[idxDescricao] || "").toString().trim();
      var recibo    = (row[idxRecibo] || "").toString().trim();

      var reciboNorm = normalizarTexto_(recibo);

      // Regras:
      // Pendência  -> etiqueta vazia OU descricao vazia OU recibo = "não"
      // Justificada-> etiqueta preenchida E descricao preenchida E recibo = "sim"
      var temPendencia =
        (!etiqueta) ||
        (!descricao) ||
        (reciboNorm === "nao" || reciboNorm === "não");

      var temJustificativa =
        (!!etiqueta) &&
        (!!descricao) &&
        (reciboNorm === "sim");

      var linhaArr = [
        loja,
        dataStr,
        valor,
        etiqueta,
        descricao,
        recibo
      ];

      if (temPendencia) {
        pendencias.push(linhaArr);
      }

      if (temJustificativa) {
        justificadas.push(linhaArr);
      }
    });

    return {
      ok: true,
      grupoOriginal: grupoOriginal,
      colunas: [
        "Loja",
        "Data da Transação",
        "Valor (R$)",
        "Etiqueta",
        "Descrição",
        "Recibo"
      ],
      pendencias: pendencias,
      justificadas: justificadas
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : e
    };
  }
}

/**
 * Resumo de pendências POR LOJA, dentro de um grupo/time (opcional).
 *
 * Saída:
 * {
 *   ok: true,
 *   grupoOriginal: "...",
 *   linhas: [
 *     {
 *       loja: "123",
 *       totalTransacoes: 10,
 *       valorTransacionado: 2000.00,
 *       totalPendencias: 3,
 *       valorPendente: 500.00,
 *       percPendente: 25.0,
 *       pendEtiqueta: 2,
 *       pendDescricao: 1,
 *       pendRecibo: 3
 *     },
 *     ...
 *   ],
 *   totais: {
 *     totalTransacoes: ...,
 *     valorTransacionado: ...,
 *     totalPendencias: ...,
 *     valorPendente: ...,
 *     pendEtiqueta: ...,
 *     pendDescricao: ...,
 *     pendRecibo: ...
 *   }
 * }
 */

function getResumoPendenciasPorLoja(grupo, dataInicioStr, dataFimStr) {
  vektorAssertFunctionAllowed_("getResumoPendenciasPorLoja");
  try {
    var info = carregarLinhasBaseClara_();
    if (info.error) {
      return { ok: false, error: info.error };
    }

    var header = info.header;
    var linhas = info.linhas;

    var IDX_DATA  = 0;   // "Data da Transação"
    var IDX_VALOR = 5;   // "Valor em R$"
    var IDX_GRUPO = 17;  // "Grupos"
    var IDX_LOJA  = 21;  // "LojaNum"

    // helper local: match EXATO (evita "Recibo" bater em "Nome dos Recibos")
    function findHeaderExactLocal_(headerArr, label) {
      var alvo = normalizarTexto_(label || "");
      for (var i = 0; i < headerArr.length; i++) {
        var h = normalizarTexto_(String(headerArr[i] || ""));
        if (h === alvo) return i;
      }
      return -1;
    }

    // === índices EXATOS primeiro, depois fallback ===
    var idxRecibo = findHeaderExactLocal_(header, "Recibo");
    if (idxRecibo < 0) idxRecibo = encontrarIndiceColuna_(header, ["Recibo", "NF / Recibo", "NF/Recibo"]);
    if (idxRecibo < 0) idxRecibo = 14; // O (fallback fixo)

    var idxEtiqueta = findHeaderExactLocal_(header, "Etiquetas");
    if (idxEtiqueta < 0) idxEtiqueta = findHeaderExactLocal_(header, "Etiqueta");
    if (idxEtiqueta < 0) idxEtiqueta = encontrarIndiceColuna_(header, ["Etiquetas", "Etiqueta"]);
    if (idxEtiqueta < 0) idxEtiqueta = 19; // T

    var idxDescricao = findHeaderExactLocal_(header, "Descrição");
    if (idxDescricao < 0) idxDescricao = findHeaderExactLocal_(header, "Descricao");
    if (idxDescricao < 0) idxDescricao = encontrarIndiceColuna_(header, ["Descrição", "Descricao"]);
    if (idxDescricao < 0) idxDescricao = 20; // U

    var idxRecibo = encontrarIndiceColuna_(header, ["Recibo", "NF / Recibo", "NF/Recibo"]);
    var idxEtiqueta = encontrarIndiceColuna_(header, ["Etiquetas", "Etiqueta"]);
    var idxDescricao = encontrarIndiceColuna_(header, ["Descrição", "Descricao", "Comentário"]);

    if (idxRecibo < 0 || idxEtiqueta < 0 || idxDescricao < 0) {
      return {
        ok: false,
        error: "Não encontrei as colunas de Recibo/Etiquetas/Descrição na BaseClara."
      };
    }

    var grupoOriginal = (grupo || "").toString().trim();
    var grupoNorm = normalizarTexto_(grupoOriginal);

    var filtradas = filtrarLinhasPorPeriodo_(
      linhas,
      IDX_DATA,
      dataInicioStr,
      dataFimStr
    );

    var mapa = {}; // key = loja

    filtradas.forEach(function (row) {
      var loja = (row[IDX_LOJA] || "").toString().trim();
      if (!loja) return;

      var grupoLinhaOriginal = (row[IDX_GRUPO] || "").toString();
      var grupoLinhaNorm = normalizarTexto_(grupoLinhaOriginal);

      if (grupoNorm) {
        var casaGrupo =
          grupoLinhaNorm.indexOf(grupoNorm) !== -1 ||
          grupoNorm.indexOf(grupoLinhaNorm) !== -1;
        if (!casaGrupo) return;
      }

      var valor = Number(row[IDX_VALOR]) || 0;

      var etiqueta  = (row[idxEtiqueta]  || "").toString().trim();
      var descricao = (row[idxDescricao] || "").toString().trim();
      var recibo    = (row[idxRecibo]    || "").toString().trim();

      var reciboNorm = normalizarTexto_(recibo);

      var temPendenciaEtiqueta  = !etiqueta;
      var temPendenciaDescricao = !descricao;
      var temPendenciaRecibo =
        !recibo || reciboNorm === "nao" || reciboNorm === "não";

      var temPendencia =
        temPendenciaEtiqueta || temPendenciaDescricao || temPendenciaRecibo;

      if (!mapa[loja]) {
        mapa[loja] = {
          loja: loja,
          totalTransacoes: 0,
          valorTransacionado: 0,
          totalPendencias: 0,
          valorPendente: 0,
          pendEtiqueta: 0,
          pendDescricao: 0,
          pendRecibo: 0
        };
      }

      var item = mapa[loja];

      // Todas as transações entram no volume total
      item.totalTransacoes++;
      item.valorTransacionado += valor;

      if (temPendencia) {
        // 1 transação pendente
        item.totalPendencias++;
        item.valorPendente += valor;

        // Cada tipo é contado separado. Uma transação pode somar em mais de uma coluna.
        if (temPendenciaEtiqueta) {
          item.pendEtiqueta++;
        }
        if (temPendenciaDescricao) {
          item.pendDescricao++;
        }
        if (temPendenciaRecibo) {
          item.pendRecibo++;
        }
      }
    });

    var linhasSaida = [];
    var totTrans = 0;
    var totValTrans = 0;
    var totPend = 0;
    var totValPend = 0;
    var totPEtiq = 0;
    var totPDesc = 0;
    var totPRec = 0;

    Object.keys(mapa).forEach(function (loja) {
      var it = mapa[loja];

      totTrans    += it.totalTransacoes;
      totValTrans += it.valorTransacionado;
      totPend     += it.totalPendencias;
      totValPend  += it.valorPendente;
      totPEtiq    += it.pendEtiqueta;
      totPDesc    += it.pendDescricao;
      totPRec     += it.pendRecibo;

      var perc = 0;
      if (it.valorTransacionado > 0 && it.valorPendente > 0) {
        perc = (it.valorPendente / it.valorTransacionado) * 100;
      }

      linhasSaida.push({
        loja: it.loja,
        totalTransacoes: it.totalTransacoes,
        valorTransacionado: it.valorTransacionado,
        totalPendencias: it.totalPendencias,
        valorPendente: it.valorPendente,
        percPendente: perc,
        pendEtiqueta: it.pendEtiqueta,
        pendDescricao: it.pendDescricao,
        pendRecibo: it.pendRecibo
      });
    });

    linhasSaida.sort(function (a, b) {
      if (b.valorPendente !== a.valorPendente) {
        return b.valorPendente - a.valorPendente;
      }
      return b.totalPendencias - a.totalPendencias;
    });

    return {
      ok: true,
      grupoOriginal: grupoOriginal,
      linhas: linhasSaida,
      totais: {
        totalTransacoes: totTrans,
        valorTransacionado: totValTrans,
        totalPendencias: totPend,
        valorPendente: totValPend,
        pendEtiqueta: totPEtiq,
        pendDescricao: totPDesc,
        pendRecibo: totPRec
      }
    };

  } catch (e) {
    return { ok: false, error: String((e && e.message) ? e.message : e) };
  }
}

/**
 * Resumo de pendências POR TIME.
 *
 * Se grupoFiltro vier preenchido, filtra só aquele grupo.
 */

function getResumoPendenciasPorTime(dataInicioStr, dataFimStr, grupoFiltro) {
  vektorAssertFunctionAllowed_("getResumoPendenciasPorTime");
  try {
    var info = carregarLinhasBaseClara_();
    if (info.error) {
      return { ok: false, error: info.error };
    }

    var header = info.header;
    var linhas = info.linhas;

    var IDX_DATA  = 0;
    var IDX_VALOR = 5;
    var IDX_GRUPO = 17;

    // helper local: match EXATO (evita "Recibo" bater em "Nome dos Recibos")
    function findHeaderExactLocal_(headerArr, label) {
      var alvo = normalizarTexto_(label || "");
      for (var i = 0; i < headerArr.length; i++) {
        var h = normalizarTexto_(String(headerArr[i] || ""));
        if (h === alvo) return i;
      }
      return -1;
    }

    // === índices EXATOS primeiro, depois fallback ===
    var idxRecibo = findHeaderExactLocal_(header, "Recibo");
    if (idxRecibo < 0) idxRecibo = encontrarIndiceColuna_(header, ["Recibo", "NF / Recibo", "NF/Recibo"]);
    if (idxRecibo < 0) idxRecibo = 14; // O (fallback fixo)

    var idxEtiqueta = findHeaderExactLocal_(header, "Etiquetas");
    if (idxEtiqueta < 0) idxEtiqueta = findHeaderExactLocal_(header, "Etiqueta");
    if (idxEtiqueta < 0) idxEtiqueta = encontrarIndiceColuna_(header, ["Etiquetas", "Etiqueta"]);
    if (idxEtiqueta < 0) idxEtiqueta = 19; // T

    var idxDescricao = findHeaderExactLocal_(header, "Descrição");
    if (idxDescricao < 0) idxDescricao = findHeaderExactLocal_(header, "Descricao");
    if (idxDescricao < 0) idxDescricao = encontrarIndiceColuna_(header, ["Descrição", "Descricao"]);
    if (idxDescricao < 0) idxDescricao = 20; // U

    var idxRecibo = encontrarIndiceColuna_(header, ["Recibo", "NF / Recibo", "NF/Recibo"]);
    var idxEtiqueta = encontrarIndiceColuna_(header, ["Etiquetas", "Etiqueta"]);
    var idxDescricao = encontrarIndiceColuna_(header, ["Descrição", "Descricao", "Comentário"]);

    if (idxRecibo < 0 || idxEtiqueta < 0 || idxDescricao < 0) {
      return {
        ok: false,
        error: "Não encontrei as colunas de Recibo/Etiquetas/Descrição na BaseClara."
      };
    }

    var filtradas = filtrarLinhasPorPeriodo_(
      linhas,
      IDX_DATA,
      dataInicioStr,
      dataFimStr
    );

    var grupoFiltroOriginal = (grupoFiltro || "").toString().trim();
    var grupoFiltroNorm = normalizarTexto_(grupoFiltroOriginal);

    var mapa = {}; // key = nome do time

    filtradas.forEach(function (row) {
      var grupoLinhaOriginal = (row[IDX_GRUPO] || "").toString().trim();
      if (!grupoLinhaOriginal) return;

      var grupoLinhaNorm = normalizarTexto_(grupoLinhaOriginal);

      if (grupoFiltroNorm) {
        var casaGrupo =
          grupoLinhaNorm.indexOf(grupoFiltroNorm) !== -1 ||
          grupoFiltroNorm.indexOf(grupoLinhaNorm) !== -1;
        if (!casaGrupo) return;
      }

      var valor = Number(row[IDX_VALOR]) || 0;

      var etiqueta  = (row[idxEtiqueta]  || "").toString().trim();
      var descricao = (row[idxDescricao] || "").toString().trim();
      var recibo    = (row[idxRecibo]    || "").toString().trim();

      var reciboNorm = normalizarTexto_(recibo);

      var temPendenciaEtiqueta  = !etiqueta;
      var temPendenciaDescricao = !descricao;
      var temPendenciaRecibo =
        !recibo || reciboNorm === "nao" || reciboNorm === "não";

      var temPendencia =
        temPendenciaEtiqueta || temPendenciaDescricao || temPendenciaRecibo;

      if (!mapa[grupoLinhaOriginal]) {
        mapa[grupoLinhaOriginal] = {
          time: grupoLinhaOriginal,
          totalTransacoes: 0,
          valorTransacionado: 0,
          totalPendencias: 0,
          valorPendente: 0,
          pendEtiqueta: 0,
          pendDescricao: 0,
          pendRecibo: 0
        };
      }

      var item = mapa[grupoLinhaOriginal];

      item.totalTransacoes++;
      item.valorTransacionado += valor;

      if (temPendencia) {
        item.totalPendencias++;
        item.valorPendente += valor;

        if (temPendenciaEtiqueta) {
          item.pendEtiqueta++;
        }
        if (temPendenciaDescricao) {
          item.pendDescricao++;
        }
        if (temPendenciaRecibo) {
          item.pendRecibo++;
        }
      }
    });

    var linhasSaida = [];
    var totTrans = 0;
    var totValTrans = 0;
    var totPend = 0;
    var totValPend = 0;
    var totPEtiq = 0;
    var totPDesc = 0;
    var totPRec = 0;

    Object.keys(mapa).forEach(function (key) {
      var it = mapa[key];

      totTrans    += it.totalTransacoes;
      totValTrans += it.valorTransacionado;
      totPend     += it.totalPendencias;
      totValPend  += it.valorPendente;
      totPEtiq    += it.pendEtiqueta;
      totPDesc    += it.pendDescricao;
      totPRec     += it.pendRecibo;

      var perc = 0;
      if (it.valorTransacionado > 0 && it.valorPendente > 0) {
        perc = (it.valorPendente / it.valorTransacionado) * 100;
      }

      linhasSaida.push({
        time: it.time,
        totalTransacoes: it.totalTransacoes,
        valorTransacionado: it.valorTransacionado,
        totalPendencias: it.totalPendencias,
        valorPendente: it.valorPendente,
        percPendente: perc,
        pendEtiqueta: it.pendEtiqueta,
        pendDescricao: it.pendDescricao,
        pendRecibo: it.pendRecibo
      });
    });

    linhasSaida.sort(function (a, b) {
      if (b.valorPendente !== a.valorPendente) {
        return b.valorPendente - a.valorPendente;
      }
      return b.totalPendencias - a.totalPendencias;
    });

    return {
      ok: true,
      linhas: linhasSaida,
      totais: {
        totalTransacoes: totTrans,
        valorTransacionado: totValTrans,
        totalPendencias: totPend,
        valorPendente: totValPend,
        pendEtiqueta: totPEtiq,
        pendDescricao: totPDesc,
        pendRecibo: totPRec
      }
    };

  } catch (e) {
    return { ok: false, error: String((e && e.message) ? e.message : e) };
  }
}

function enviarResumoPorEmail(grupo) {
  try {
    const emailDestino = Session.getActiveUser().getEmail();
    if (!emailDestino) return { ok: false, error: "Usuário sem e-mail ativo" };

    const resumo = getResumoTransacoesPorGrupo(grupo, "", "");
    if (!resumo.ok || !resumo.top) return { ok: false, error: "Sem dados" };

    let corpo = `
  <p>Segue resumo de transações para o time <b>${resumo.grupo}</b>:</p>
  <table border="1" cellspacing="0" cellpadding="6"
         style="border-collapse:collapse;font-family:Arial,sans-serif;font-size:12px;text-align:center">
    <tr style="background:#06167d;color:#fff">
      <th style="text-align:center">Loja</th>
      <th style="text-align:center">Qtd Transações</th>
      <th style="text-align:center">Valor (R$)</th>
    </tr>
    ${resumo.lojas.slice(0,10).map(l => `
      <tr>
        <td style="text-align:center">${l.loja}</td>
        <td style="text-align:center">${l.total}</td>
        <td style="text-align:center">
          ${l.valorTotal.toLocaleString("pt-BR",{style:"currency",currency:"BRL"})}
        </td>
      </tr>
    `).join("")}
  </table>
  <br/>
  <p><i>Gerado automaticamente pelo Assistente Vektor</i></p>`;

    GmailApp.sendEmail(emailDestino, `Resumo de transações | ${resumo.grupo}`, " ", {
      from: "vektor@gruposbf.com.br",
      htmlBody: corpo,
      name: "Vektor - Grupo SBF"
    });

    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

/**
 * Ranking POR TIME (grupo), por quantidade ou por valor.
 * @param {string} dataInicioStr ISO ou vazio
 * @param {string} dataFimStr ISO ou vazio
 * @param {string} criterio "quantidade" | "valor"
 */
function getResumoTransacoesPorTime(dataInicioStr, dataFimStr, criterio) {
  vektorAssertFunctionAllowed_("getResumoTransacoesPorTime");
  try {
    var info = carregarLinhasBaseClara_();
    if (info.error) return { ok:false, error: info.error };

    criterio = (criterio || "").toString().toLowerCase();
    if (criterio !== "valor" && criterio !== "quantidade") criterio = "quantidade";

    var linhas = info.linhas;

    // Índices fixos conforme sua base
    var IDX_DATA  = 0;   // "Data da Transação"
    var IDX_VALOR = 5;   // "Valor em R$"
    var IDX_GRUPO = 17;  // "Grupos"

    var filtradas = filtrarLinhasPorPeriodo_(linhas, IDX_DATA, dataInicioStr, dataFimStr);

    var mapa = {}; // chave = grupo normalizado; valor = { time: nomeOriginal, total, valorTotal }
    for (var i = 0; i < filtradas.length; i++) {
      var row = filtradas[i];

      var grupoOriginal = (row[IDX_GRUPO] || "").toString().trim();
      if (!grupoOriginal) continue;

      var grupoNorm = normalizarTexto_(grupoOriginal);
      if (!grupoNorm) continue;

      if (!mapa[grupoNorm]) {
        mapa[grupoNorm] = { time: grupoOriginal, total: 0, valorTotal: 0 };
      }

      mapa[grupoNorm].total++;
      var valor = Number(row[IDX_VALOR]) || 0;
      mapa[grupoNorm].valorTotal += valor;
    }

    var timesArr = [];
    for (var k in mapa) {
      if (Object.prototype.hasOwnProperty.call(mapa, k)) {
        timesArr.push(mapa[k]);
      }
    }

    // ordena conforme critério
    if (criterio === "valor") {
      timesArr.sort(function (a, b) {
        if (b.valorTotal !== a.valorTotal) return b.valorTotal - a.valorTotal;
        return b.total - a.total; // empate por quantidade
      });
    } else {
      timesArr.sort(function (a, b) {
        if (b.total !== a.total) return b.total - a.total;
        return b.valorTotal - a.valorTotal; // empate por valor
      });
    }

    var top = timesArr.length ? timesArr[0] : null;

    return {
      ok: true,
      criterio: criterio,
      times: timesArr,  // [{time, total, valorTotal}, ...]
      top: top
    };
  } catch (e) {
    return { ok:false, error: "Falha em getResumoTransacoesPorTime: " + e };
  }
}

/**
 * Resumo de transações por TIME, filtrando pela coluna
 * "Extrato da conta" (coluna B da BaseClara).
 *
 * @param {string} extratoConta  Texto exato do extrato (ex.: "06 Nov 2025 - 05 Dec 2025")
 * @param {string} criterio      "valor" ou "quantidade" (mantém a mesma lógica do resumo por time)
 *
 * Retorna objeto compatível com renderResumoTransacoesPorTime:
 * {
 *   ok: true,
 *   criterio: "valor",
 *   extratoOriginal: "06 Nov 2025 - 05 Dec 2025",
 *   times: [
 *     { time: "Águias do Cerrado", total: 10, valorTotal: 1234.56 },
 *     ...
 *   ],
 *   top: { ... }
 * }
 */
function getResumoTransacoesPorTimeExtrato(extratoConta, criterio) {
  vektorAssertFunctionAllowed_("getResumoTransacoesPorTimeExtrato");
  try {
    var info = carregarLinhasBaseClara_();
    if (info.error) {
      return { ok: false, error: info.error };
    }

    var header = info.header;
    var linhas = info.linhas;

    // Índice da coluna "Extrato da conta"
    var idxExtrato = encontrarIndiceColuna_(header, [
      "Extrato da conta",
      "Extrato conta",
      "Extrato"
    ]);

    // Índice da coluna de valor
    var idxValor = encontrarIndiceColuna_(header, [
      "Valor em R$",
      "Valor (R$)",
      "Valor"
    ]);

    // Índice da coluna de GRUPO / TIME
    var idxGrupo = encontrarIndiceColuna_(header, [
      "Grupos",
      "Grupo",
      "Time"
    ]);

    if (idxExtrato < 0 || idxValor < 0 || idxGrupo < 0) {
      return {
        ok: false,
        error: "Não encontrei as colunas 'Extrato da conta', 'Valor' e 'Grupo/Time' na BaseClara."
      };
    }

    // Normaliza critério
    if (!criterio) {
      criterio = "valor";
    }
    criterio = String(criterio).toLowerCase();
    if (criterio !== "valor" && criterio !== "quantidade") {
      criterio = "valor";
    }

    // Normaliza o texto do extrato informado
    var alvoOriginal = (extratoConta || "").toString().trim();
    var alvoNorm = normalizarTexto_(alvoOriginal);
    if (!alvoNorm) {
      return { ok: false, error: "Extrato da conta não informado." };
    }

    // Agrupa por time, considerando somente as linhas desse extrato
    var mapa = {}; // chave = time normalizado
    for (var i = 0; i < linhas.length; i++) {
      var row = linhas[i];

      var extratoLinha = (row[idxExtrato] || "").toString();
      var extratoNorm = normalizarTexto_(extratoLinha);
      if (!extratoNorm || extratoNorm !== alvoNorm) {
        continue; // ignora linhas de outros ciclos
      }

      var grupoOriginal = (row[idxGrupo] || "").toString().trim();
      if (!grupoOriginal) continue;

      var grupoNorm = normalizarTexto_(grupoOriginal);
      if (!grupoNorm) continue;

      if (!mapa[grupoNorm]) {
        mapa[grupoNorm] = {
          time: grupoOriginal,
          total: 0,
          valorTotal: 0
        };
      }

      mapa[grupoNorm].total++;

      var valor = Number(row[idxValor]) || 0;
      mapa[grupoNorm].valorTotal += valor;
    }

    // Converte mapa em array
    var arr = [];
    for (var chave in mapa) {
      if (!Object.prototype.hasOwnProperty.call(mapa, chave)) continue;
      arr.push(mapa[chave]);
    }

    // Ordena conforme critério (mesma lógica do getResumoTransacoesPorTime)
    arr.sort(function (a, b) {
      if (criterio === "quantidade") {
        if (b.total !== a.total) {
          return b.total - a.total;
        }
        return b.valorTotal - a.valorTotal;
      }
      // padrão: valor
      if (b.valorTotal !== a.valorTotal) {
        return b.valorTotal - a.valorTotal;
      }
      return b.total - a.total;
    });

    var top = arr.length ? arr[0] : null;

    return {
      ok: true,
      criterio: criterio,
      extratoOriginal: alvoOriginal,
      times: arr,
      top: top
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : e
    };
  }
}

function exportarTransacoesFaturaXlsx(extratoConta) {
  vektorAssertFunctionAllowed_("exportarTransacoesFaturaXlsx");
  try {
    var info = carregarLinhasBaseClara_();
    if (info.error) {
      return { ok: false, error: info.error };
    }

    var header = info.header || [];
    var linhas = info.linhas || [];

    // Coluna B = "Extrato da conta"
    var idxExtrato = encontrarIndiceColuna_(header, ["Extrato da conta", "Extrato"]);
    if (idxExtrato < 0) {
      return { ok: false, error: "Coluna 'Extrato da conta' não encontrada." };
    }

    var alvo = String(extratoConta || "").trim();
    var alvoNorm = normalizarTexto_(alvo);
    if (!alvo) {
      return { ok: false, error: "Extrato não informado." };
    }

    // A até W = 23 colunas
    var COLS = 23;
    var dados = [];
    dados.push(header.slice(0, COLS));

    for (var i = 0; i < linhas.length; i++) {
      var row = linhas[i];
      if (!row) continue;

        var extratoLinha = String(row[idxExtrato] || "").trim();
        var extratoNorm = normalizarTexto_(extratoLinha);
        if (!extratoNorm || extratoNorm !== alvoNorm) continue;

      dados.push(row.slice(0, COLS));
    }

    if (dados.length <= 1) {
      return { ok: false, error: "Nenhuma transação encontrada para essa fatura." };
    }

    // Cria planilha temporária
var ss = SpreadsheetApp.create("TMP_EXPORT_FATURA");
var sh = ss.getActiveSheet();
sh.getRange(1, 1, dados.length, dados[0].length).setValues(dados);

// Garante que os dados foram gravados
SpreadsheetApp.flush();

// URL oficial de exportação do Drive (XLSX)
var fileId = ss.getId();
var url =
  "https://www.googleapis.com/drive/v3/files/" +
  fileId +
  "/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";

// Token OAuth do script
var token = ScriptApp.getOAuthToken();

// Faz download do XLSX
var resp = UrlFetchApp.fetch(url, {
  headers: {
    Authorization: "Bearer " + token
  },
  muteHttpExceptions: true
});

// Apaga planilha temporária
DriveApp.getFileById(fileId).setTrashed(true);

// Valida resposta
if (resp.getResponseCode() !== 200) {
  return {
    ok: false,
    error: "Falha ao exportar XLSX (HTTP " + resp.getResponseCode() + ")"
  };
}

// Converte para Base64
var bytes = resp.getBlob().getBytes();
var base64 = Utilities.base64Encode(bytes);

// Nome do arquivo
var filename =
  "transacoes_fatura_" +
  alvo.replace(/[^\w]+/g, "_") +
  ".xlsx";

return {
  ok: true,
  filename: filename,
  xlsxBase64: base64
};

  } catch (e) {
    return { ok: false, error: e.message || String(e) };
  }
}

function exportarLojasComPendenciasCicloAtualXlsx() {
  // ✅ RBAC (admin-only): cadastre esta função no VEKTOR_ACESSOS para Administrador
  vektorAssertFunctionAllowed_("exportarLojasComPendenciasCicloAtualXlsx");

  try {
    // 1) Pega o “universo” de lojas com pendência no ciclo atual
    var resumo = getPendenciasResumoCicloAtual();
    if (!resumo || !resumo.ok) {
      return { ok: false, error: (resumo && resumo.error) ? resumo.error : "Falha ao obter resumo de pendências do ciclo atual." };
    }

    // ✅ Fonte correta: lista explícita (prioriza ROOT; fallback para formato antigo dentro de "lojas")
    var lojas =
      (resumo && Array.isArray(resumo.lojasComPendenciaLista))
        ? resumo.lojasComPendenciaLista
        : ((resumo && resumo.lojas && Array.isArray(resumo.lojas.lojasComPendenciaLista))
            ? resumo.lojas.lojasComPendenciaLista
            : []);

    // Fallback: tenta extrair das topLojas (caso a lista não exista por alguma razão)
    if (!lojas.length) {
      var top = (resumo && Array.isArray(resumo.topLojas)) ? resumo.topLojas : [];
      lojas = top.map(function (x) { return String(x && x.loja ? x.loja : "").trim(); }).filter(Boolean);
    }

    if (!Array.isArray(lojas) || !lojas.length) {
      return { ok: false, error: "Não encontrei lojas com pendências no ciclo atual para exportar." };
    }

    // normaliza para set rápido
    var set = {};
    lojas.forEach(function (l) {
      var dig = String(l || "").replace(/\D/g, "");
      if (!dig) return;
      var cod = String(Number(dig)).padStart(4, "0");
      set[cod] = true;
    });

    // 2) Lê BaseClara inteira (header + linhas)
    var info = carregarLinhasBaseClara_();
    if (info.error) return { ok: false, error: info.error };

    var header = info.header || [];
    var linhas = info.linhas || [];

    // 3) Descobre colunas necessárias para filtrar (Loja + Data do ciclo)
    var idxLoja = encontrarIndiceColuna_(header, ["LojaNum", "Loja", "Código da loja", "Codigo da loja", "cod_loja"]);
    if (idxLoja < 0) return { ok: false, error: "Não encontrei coluna de Loja na BaseClara (LojaNum/Loja/cod_loja)." };

    var idxData = encontrarIndiceColuna_(header, ["Data da Transação", "Data", "Data_transacao", "Data transacao"]);
    if (idxData < 0) return { ok: false, error: "Não encontrei coluna de Data na BaseClara (Data da Transação/Data)." };

    // ✅ índices para aplicar filtro de pendência (mesma regra do resumo)
    var idxRecibo = encontrarIndiceColuna_(header, ["Recibo", "NF / Recibo", "NF/Recibo"]);
    if (idxRecibo < 0) return { ok: false, error: "Não encontrei coluna 'Recibo' na BaseClara." };

    var idxEtiqueta = encontrarIndiceColuna_(header, ["Etiquetas", "Etiqueta"]);
    if (idxEtiqueta < 0) return { ok: false, error: "Não encontrei coluna 'Etiquetas' na BaseClara." };

    var idxDescricao = encontrarIndiceColuna_(header, ["Descrição", "Descricao"]);
    if (idxDescricao < 0) return { ok: false, error: "Não encontrei coluna 'Descrição' na BaseClara." };

    function isVazioPend_(v) {
      if (v === null || v === undefined) return true;
      if (typeof v === "boolean") return (v === false);

      var s = String(v).trim().toLowerCase();
      if (!s) return true;

      if (s === "-" || s === "—" || s === "n/a") return true;
      if (s === "não" || s === "nao") return true;
      if (s === "false" || s === "0") return true;

      return false;
    }

    // 4) Calcula período do ciclo (06->hoje)
    var pc = getPeriodoCicloClara_();
    var ini = pc && pc.inicio ? pc.inicio : null;
    var fim = new Date();

    if (!ini) return { ok: false, error: "Não consegui identificar o início do ciclo atual." };

    // 5) Filtra linhas: (loja ∈ set) e (data ∈ [ini..fim]) e (tem pendência)
    var filtradas = [];
    for (var i = 0; i < linhas.length; i++) {
      var row = linhas[i];

      var lojaRaw = String(row[idxLoja] || "").trim();
      var dig = lojaRaw.replace(/\D/g, "");
      if (!dig) continue;
      var loja4 = String(Number(dig)).padStart(4, "0");
      if (!set[loja4]) continue;

      var d = parseDateClara_(row[idxData]);
      if (!d) continue;
      if (d < ini || d > fim) continue;

      var recibo = row[idxRecibo];
      var etiqueta = row[idxEtiqueta];
      var desc = row[idxDescricao];

      var temPendRecibo = isVazioPend_(recibo);
      var temPendEtiqueta = isVazioPend_(etiqueta);
      var temPendDescricao = isVazioPend_(desc);

      if (!(temPendRecibo || temPendEtiqueta || temPendDescricao)) continue;

      filtradas.push(row);
    }

    if (!filtradas.length) {
      return { ok: false, error: "Nenhuma transação com pendência encontrada para essas lojas no ciclo atual." };
    }

    // 6) Gera XLSX temporário e devolve base64
    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    var nome = "Vektor - Lojas com pendências (ciclo atual) - " + Utilities.formatDate(new Date(), tz, "yyyyMMdd_HHmm") + ".xlsx";

    var xlsxBlob = buildXlsxFromTable_(header, filtradas, "BaseClara_filtrada");
    var b64 = Utilities.base64Encode(xlsxBlob.getBytes());

    return {
      ok: true,
      filename: nome,
      xlsxBase64: b64,
      meta: { totalRows: filtradas.length, totalLojas: Object.keys(set).length }
    };

  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

function enviarLojasComPendenciasCicloAtualEmail() {
  // ✅ RBAC (admin-only): cadastre esta função no VEKTOR_ACESSOS para Administrador
  vektorAssertFunctionAllowed_("enviarLojasComPendenciasCicloAtualEmail");

  try {
    var to = (Session.getActiveUser().getEmail() || "").trim().toLowerCase();
    if (!to) {
      return { ok: false, error: "Não foi possível identificar seu e-mail Google (sessão vazia)." };
    }

    // Reaproveita o mesmo export (fonte única da verdade)
    var exp = exportarLojasComPendenciasCicloAtualXlsx();
    if (!exp || !exp.ok || !exp.xlsxBase64) {
      return { ok: false, error: (exp && exp.error) ? exp.error : "Falha ao gerar o anexo para envio." };
    }

    var bytes = Utilities.base64Decode(exp.xlsxBase64);
    var filename = exp.filename || "Vektor - Lojas com pendências (ciclo atual).xlsx";

    var blob = Utilities.newBlob(
      bytes,
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      filename
    );

    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    var hoje = Utilities.formatDate(new Date(), tz, "dd/MM/yyyy HH:mm");

    var subject = "Vektor | Lojas com pendências (ciclo atual)";
    var htmlBody =
      "<div style='font-family:Arial,sans-serif; font-size:13px; color:#0f172a;'>" +
        "<p>Segue em anexo a listagem de <b>transações com pendências</b> do ciclo atual (extração: " + hoje + ").</p>" +
        "<p style='color:#64748b; font-size:12px;'>Enviado automaticamente pelo Vektor.</p>" +
      "</div>";

    GmailApp.sendEmail(to, subject, " ", {
      from: "vektor@gruposbf.com.br",
      htmlBody: htmlBody,
      attachments: [blob],
      name: "Vektor - Grupo SBF"
    });

    return { ok: true, to: to, filename: filename, meta: exp.meta || {} };

  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

/**
 * Helper: cria um XLSX a partir de (header + rows2d).
 * Usa planilha temporária e exporta para XLSX.
 */
function buildXlsxFromTable_(header, rows2d, sheetName) {
  sheetName = sheetName || "Export";

  // cria planilha temp
  var temp = SpreadsheetApp.create("TEMP_VEKTOR_EXPORT_" + new Date().getTime());
  var fileId = temp.getId();

  var sh = temp.getSheets()[0];
  sh.setName(sheetName);

  // escreve header + dados
  var all = [header].concat(rows2d);
  sh.getRange(1, 1, all.length, header.length).setValues(all);
  sh.setFrozenRows(1);

  // exporta XLSX via endpoint do Drive
  var url = "https://www.googleapis.com/drive/v3/files/" + fileId + "/export?mimeType=application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
  var token = ScriptApp.getOAuthToken();

  var resp = UrlFetchApp.fetch(url, {
    headers: { Authorization: "Bearer " + token },
    muteHttpExceptions: true
  });

  if (resp.getResponseCode() !== 200) {
    // tenta limpar
    try { DriveApp.getFileById(fileId).setTrashed(true); } catch (_) {}
    throw new Error("Falha ao exportar XLSX (HTTP " + resp.getResponseCode() + "): " + resp.getContentText());
  }

  var blob = resp.getBlob().setName("export.xlsx");

  // limpa temp
  try { DriveApp.getFileById(fileId).setTrashed(true); } catch (_) {}

  return blob;
}

/**
 * Resumo de transações por CATEGORIA DA COMPRA (BaseClara).
 * 
 * - dataInicioStr / dataFimStr: datas em ISO (como já usamos nas outras funções). 
 *   Se vierem vazias, usa o comportamento padrão da filtrarLinhasPorPeriodo_ (últimos dias).
 * - criterio: "valor" ou "quantidade" (qual critério será usado para ordenar).
 *
 * Retorna:
 * {
 *   ok: true,
 *   criterio: "valor" ou "quantidade",
 *   categorias: [
 *     { categoria: "Alimentação", total: 10, valorTotal: 1234.56 },
 *     ...
 *   ],
 *   top: { ... } // primeira posição da lista (maior valor ou maior quantidade)
 * }
 */
function getResumoTransacoesPorCategoria(dataInicioStr, dataFimStr, criterio) {
  vektorAssertFunctionAllowed_("getResumoTransacoesPorCategoria");
  try {
    var info = carregarLinhasBaseClara_();
    if (info.error) {
      return { ok: false, error: info.error };
    }

    var header = info.header;
    var linhas = info.linhas;

    // Descobre os índices das colunas dinamicamente pelo cabeçalho
    var idxData = encontrarIndiceColuna_(header, [
      "Data da Transação",
      "Data Transação",
      "Data"
    ]);

    var idxValor = encontrarIndiceColuna_(header, [
      "Valor em R$",
      "Valor (R$)",
      "Valor"
    ]);

    var idxCategoria = encontrarIndiceColuna_(header, [
      "Categoria da Compra",
      "Categoria da compra",
      "Categoria",
      "Categoria Compra"
    ]);

    if (idxData < 0 || idxValor < 0 || idxCategoria < 0) {
      return {
        ok: false,
        error: "Não encontrei as colunas necessárias em BaseClara (Data / Valor / Categoria)."
      };
    }

    // normaliza critério
    criterio = (criterio || "").toString().toLowerCase();
    if (criterio !== "valor" && criterio !== "quantidade") {
      criterio = "quantidade";
    }

    // filtra por período (usa mesma função que já existe)
    var filtradas = filtrarLinhasPorPeriodo_(linhas, idxData, dataInicioStr, dataFimStr);

    var mapa = {}; // chave = nome da categoria
    for (var i = 0; i < filtradas.length; i++) {
      var row = filtradas[i];

      var cat = (row[idxCategoria] || "Sem categoria").toString().trim();
      if (!cat) cat = "Sem categoria";

      if (!mapa[cat]) {
        mapa[cat] = {
          categoria: cat,
          total: 0,
          valorTotal: 0
        };
      }

      mapa[cat].total++;
      var valor = Number(row[idxValor]) || 0;
      mapa[cat].valorTotal += valor;
    }

    // transforma o mapa em array
    var arr = [];
    for (var chave in mapa) {
      if (Object.prototype.hasOwnProperty.call(mapa, chave)) {
        arr.push(mapa[chave]);
      }
    }

    // ordena conforme o critério
    if (criterio === "valor") {
      arr.sort(function (a, b) {
        if (b.valorTotal !== a.valorTotal) return b.valorTotal - a.valorTotal;
        return b.total - a.total; // desempate pela quantidade
      });
    } else {
      // "quantidade"
      arr.sort(function (a, b) {
        if (b.total !== a.total) return b.total - a.total;
        return b.valorTotal - a.valorTotal; // desempate pelo valor
      });
    }

    var top = arr.length ? arr[0] : null;

    return {
      ok: true,
      criterio: criterio,
      categorias: arr,
      top: top,
      // 🔹 novo: devolve o período usado
      dataInicioIso: dataInicioStr || "",
      dataFimIso:    dataFimStr    || ""
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : e
    };
  }
}

function getTransacoesPorCategoria(dataInicioStr, dataFimStr, categoriaNome) {
  vektorAssertFunctionAllowed_("getTransacoesPorCategoria");
  try {
    var info = carregarLinhasBaseClara_();
    if (info.error) {
      return { ok: false, error: info.error };
    }

    var header = info.header;
    var linhas = info.linhas;

    // Índices principais
    var idxData = encontrarIndiceColuna_(header, [
      "Data da Transação", "Data Transação", "Data"
    ]);

    var idxValor = encontrarIndiceColuna_(header, [
      "Valor em R$", "Valor (R$)", "Valor"
    ]);

    var idxCategoria = encontrarIndiceColuna_(header, [
      "Categoria da Compra", "Categoria"
    ]);

    var idxLoja = encontrarIndiceColuna_(header, [
      "LojaNum", "Loja Num", "Loja Número", "Loja Numero", "Loja"
    ]);

    // Coluna C = "Transação" (nome do estabelecimento / texto da transação)
    var idxTransacao = 2;

    // Novos índices (tenta pelo header; se não achar, cai no índice fixo por letra)
    var idxRecibo = encontrarIndiceColuna_(header, ["Recibo"]);
    if (idxRecibo < 0) idxRecibo = 14; // O

    var idxTitular = encontrarIndiceColuna_(header, ["Titular"]);
    if (idxTitular < 0) idxTitular = 16; // Q

    var idxGrupos = encontrarIndiceColuna_(header, ["Grupos", "Grupo", "Time"]);
    if (idxGrupos < 0) idxGrupos = 17; // R

    var idxEtiquetas = encontrarIndiceColuna_(header, ["Etiquetas", "Etiqueta"]);
    if (idxEtiquetas < 0) idxEtiquetas = 19; // T

    var idxDescricao = encontrarIndiceColuna_(header, ["Descrição", "Descricao"]);
    if (idxDescricao < 0) idxDescricao = 20; // U

    if (idxData < 0 || idxValor < 0 || idxCategoria < 0) {
      return {
        ok: false,
        error: "Não encontrei as colunas necessárias em BaseClara (Data / Valor / Categoria)."
      };
    }

    // Filtra período
    var filtradas = filtrarLinhasPorPeriodo_(
      linhas,
      idxData,
      dataInicioStr,
      dataFimStr
    ) || [];

    var categoriaAlvoNorm = normalizarTexto_(categoriaNome || "");

    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    var lista = [];

    filtradas.forEach(function (row) {
      if (!row) return;

      var catOriginal = (row[idxCategoria] || "").toString().trim();
      if (!catOriginal) catOriginal = "Sem categoria";
      var catNorm = normalizarTexto_(catOriginal);

      if (categoriaAlvoNorm && catNorm !== categoriaAlvoNorm) {
        return;
      }

      var valor = Number(row[idxValor]) || 0;
      if (!valor) return;

      var dataCel = row[idxData];
      var dataBr;
      if (dataCel instanceof Date) {
        dataBr = Utilities.formatDate(dataCel, tz, "dd/MM/yyyy");
      } else {
        dataBr = dataCel || "";
      }

      var loja = row[idxLoja] != null ? String(row[idxLoja]) : "";

      lista.push({
        loja: loja,
        transacao: String(row[idxTransacao] || ""),
        data: dataBr,
        valor: valor,
        // novos
      titular: String(row[idxTitular] || ""),
      grupos: String(row[idxGrupos] || ""),
      recibo: String(row[idxRecibo] || ""),
      etiquetas: String(row[idxEtiquetas] || ""),
      descricao: String(row[idxDescricao] || "")
      });
    });

    if (!lista.length) {
      return {
        ok: true,
        categoria: categoriaNome || "",
        linhas: [],
        dataInicioIso: dataInicioStr || "",
        dataFimIso: dataFimStr || "",
        total: 0
      };
    }

    // Ordena do maior para o menor valor
    lista.sort(function (a, b) {
      return b.valor - a.valor;
    });

    var total = 0;
    for (var i = 0; i < lista.length; i++) {
      total += Number(lista[i].valor) || 0;
    }

    return {
      ok: true,
      categoria: categoriaNome || "",
      linhas: lista,
      dataInicioIso: dataInicioStr || "",
      dataFimIso: dataFimStr || "",
      total: total
    };

  } catch (e) {
    return {
      ok: false,
      error: "Erro em getTransacoesPorCategoria: " + (e && e.message ? e.message : e)
    };
  }
}



// Remove zeros à esquerda de um código de loja, para comparar "0035" com "35"
function removerZerosEsquerda_(codigo) {
  if (codigo == null) return "";
  var s = String(codigo).trim();
  s = s.replace(/^0+/, "");
  return s || "0";
}

// Gera um texto curto de período: "Últimos 30 dias" ou "de 01/12/2025 a 10/12/2025"
function montarDescricaoPeriodoSimples_(iniDate, fimDate) {
  try {
    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    var iniBr = Utilities.formatDate(iniDate, tz, "dd/MM/yyyy");
    var fimBr = Utilities.formatDate(fimDate, tz, "dd/MM/yyyy");
    return "de " + iniBr + " a " + fimBr;
  } catch (e) {
    return "";
  }
}

/**
 * Retorna as maiores transações individuais da BaseClara,
 * filtrando por time (grupo) e/ou loja, em um período.
 *
 * @param {string} grupoNome      - nome do time (coluna "Grupos")
 * @param {string} lojaCodigo     - código da loja (coluna "LojaNum")
 * @param {string} dataInicioStr  - ISO string do início (pode vir vazio)
 * @param {string} dataFimStr     - ISO string do fim (pode vir vazio)
 * @param {number} topN           - quantidade de linhas desejadas (Top N)
 */
function getMaioresTransacoesIndividuais(grupoNome, lojaCodigo, dataInicioStr, dataFimStr, topN) {
  vektorAssertFunctionAllowed_("getMaioresTransacoesIndividuais");
  try {
    // Flag para saber se o período veio do usuário (frase) ou se é o default (últimos 30 dias)
    var periodoFoiInformadoPeloUsuario = !!(dataInicioStr && dataFimStr);

    var info = carregarLinhasBaseClara_();
    if (!info || info.error) {
      return { ok: false, error: info && info.error ? info.error : "Não foi possível ler a BaseClara." };
    }

    var header = info.header || [];
    var linhas = info.linhas || [];

    // Índices das colunas
    var idxData   = encontrarIndiceColuna_(header, "Data da Transação");
    var idxValor  = encontrarIndiceColuna_(header, "Valor em R$");
    var idxGrupo  = encontrarIndiceColuna_(header, "Grupos");
    var idxLoja   = encontrarIndiceColuna_(header, "LojaNum");
    var idxStatus    = encontrarIndiceColuna_(header, "Status");
    var idxCategoria = encontrarIndiceColuna_(header, "Categoria da Compra");
    var idxTitular   = encontrarIndiceColuna_(header, "Titular");

    // Validação das novas colunas
    if (idxStatus < 0 || idxCategoria < 0 || idxTitular < 0) {
      return {
        ok: false,
        error: "Não encontrei Status, Categoria da Compra ou Titular na BaseClara."
      };
    }

    // ATENÇÃO: aqui queremos a coluna C = "Transação" (nome do estabelecimento).
    // Não podemos usar encontrarIndiceColuna_ de forma vaga,
    // senão ele pega "Data da Transação".
    var idxDescricaoEst = -1;
    for (var i = 0; i < header.length; i++) {
      var hNorm = normalizarTexto_(header[i] || "");
      if (hNorm === "transacao") { // igualdade exata após normalização
        idxDescricaoEst = i;
        break;
      }
    }

    if (idxData < 0 || idxValor < 0 || idxLoja < 0 || idxDescricaoEst < 0) {
      return {
        ok: false,
        error: "Não encontrei alguma das colunas obrigatórias na BaseClara (Data da Transação, Valor em R$, LojaNum, Transação)."
      };
    }

    // Se não vier período, usamos últimos 30 dias
    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    var hoje = new Date();
    var iniDate, fimDate;

    if (dataInicioStr && dataFimStr) {
      iniDate = new Date(dataInicioStr);
      fimDate = new Date(dataFimStr);
    } else {
      fimDate = hoje;
      iniDate = new Date();
      iniDate.setDate(hoje.getDate() - 30);
    }

    var dataInicioIso = iniDate.toISOString();
    var dataFimIso    = fimDate.toISOString();

    // Texto do período
    var periodoDescricao;
    if (periodoFoiInformadoPeloUsuario) {
      periodoDescricao = montarDescricaoPeriodoSimples_(iniDate, fimDate);
    } else {
      periodoDescricao = "Últimos 30 dias";
    }

    // Filtra por período
    var filtradas = filtrarLinhasPorPeriodo_(linhas, idxData, dataInicioIso, dataFimIso) || [];

    // Normalizações para filtros
    var grupoNorm  = grupoNome ? normalizarTexto_(grupoNome) : "";
    var lojaFiltro = lojaCodigo ? String(lojaCodigo).trim() : "";

    if (lojaFiltro) {
      // compara sempre sem zeros à esquerda
      lojaFiltro = removerZerosEsquerda_(lojaFiltro);
    }

    var lista = [];

    filtradas.forEach(function (linha) {
      if (!linha) return;

      var valor = Number(linha[idxValor]) || 0;
      if (valor <= 0) return;

      // Valores de loja da linha
      var lojaLinha     = (linha[idxLoja] || "").toString().trim();
      var lojaLinhaNorm = removerZerosEsquerda_(lojaLinha);

      // Regra de filtro:
      // 1) Se veio loja, ela manda: ignora grupo (filtra só por loja).
      // 2) Só usa grupo quando NÃO houver lojaFiltro.
      if (lojaFiltro) {
        if (lojaLinhaNorm !== lojaFiltro) return;
      } else if (grupoNorm) {
        var grupoLinhaNorm = normalizarTexto_(linha[idxGrupo] || "");
        if (grupoLinhaNorm !== grupoNorm) return;
      }

      // Data em dd/MM/yyyy
      var dataCel = linha[idxData];
      var dataBr;
      if (dataCel instanceof Date) {
        dataBr = Utilities.formatDate(dataCel, tz, "dd/MM/yyyy");
      } else {
        dataBr = dataCel || "";
      }

      lista.push({
        loja: String(linha[idxLoja] || ""),
        estabelecimento: String(linha[idxDescricaoEst] || ""),
        data: dataBr,
        valor: valor,
        status: String(linha[idxStatus] || ""),
        categoria: String(linha[idxCategoria] || ""),
        titular: String(linha[idxTitular] || "")
      });
    }); // ← FALTAVA FECHAR O forEach AQUI

    // Se não houve nenhuma linha após filtros
    if (!lista.length) {
      return {
        ok: true,
        linhas: [],
        dataInicioIso: dataInicioIso,
        dataFimIso: dataFimIso,
        periodoDescricao: periodoDescricao,
        grupo: grupoNome || "",
        loja: lojaCodigo || "",
        topN: topN || null,
        totalSelecionadas: 0
      };
    }

    // Ordena do maior valor para o menor
    lista.sort(function (a, b) {
      return b.valor - a.valor;
    });

    var limite = (typeof topN === "number" && topN > 0) ? topN : 10;
    var selecionadas = lista.slice(0, limite);

    // Soma dos valores das linhas exibidas
    var totalSelecionadas = 0;
    for (var j = 0; j < selecionadas.length; j++) {
      totalSelecionadas += Number(selecionadas[j].valor) || 0;
    }

    return {
      ok: true,
      linhas: selecionadas,
      dataInicioIso: dataInicioIso,
      dataFimIso: dataFimIso,
      periodoDescricao: periodoDescricao,
      grupo: grupoNome || "",
      loja: lojaCodigo || "",
      topN: limite,
      totalSelecionadas: totalSelecionadas
    };

  } catch (e) {
    return {
      ok: false,
      error: "Erro em getMaioresTransacoesIndividuais: " + (e && e.message ? e.message : e)
    };
  }
}

/**
 * Resumo de categorias filtrando por TIME (grupo).
 *
 * @param {string} dataInicioStr ISO ou vazio
 * @param {string} dataFimStr ISO ou vazio
 * @param {string} criterio "valor" | "quantidade"
 * @param {string} grupo Nome do time/grupo
 */
function getResumoTransacoesPorCategoriaTime(dataInicioStr, dataFimStr, criterio, grupo) {
  vektorAssertFunctionAllowed_("getResumoTransacoesPorCategoriaTime");
  try {
    var info = carregarLinhasBaseClara_();
    if (info.error) {
      return { ok: false, error: info.error };
    }

    var header = info.header;
    var linhas = info.linhas;

    var idxData = encontrarIndiceColuna_(header, [
      "Data da Transação", "Data Transação", "Data"
    ]);

    var idxValor = encontrarIndiceColuna_(header, [
      "Valor em R$", "Valor (R$)", "Valor"
    ]);

    var idxCategoria = encontrarIndiceColuna_(header, [
      "Categoria da Compra", "Categoria da compra", "Categoria", "Categoria Compra"
    ]);

    var idxGrupo = encontrarIndiceColuna_(header, [
      "Grupos", "Grupo", "Time"
    ]);

    if (idxData < 0 || idxValor < 0 || idxCategoria < 0 || idxGrupo < 0) {
      return {
        ok: false,
        error: "Não encontrei as colunas necessárias em BaseClara (Data / Valor / Categoria / Grupo)."
      };
    }

    criterio = (criterio || "").toString().toLowerCase();
    if (criterio !== "valor" && criterio !== "quantidade") {
      criterio = "quantidade";
    }

    var grupoOriginal = (grupo || "").toString().trim();
    var grupoNorm = normalizarTexto_(grupoOriginal);

    var filtradas = filtrarLinhasPorPeriodo_(linhas, idxData, dataInicioStr, dataFimStr);

    var mapa = {}; // chave = categoria
    for (var i = 0; i < filtradas.length; i++) {
      var row = filtradas[i];

      // filtro por grupo/time
      var grupoLinhaOriginal = (row[idxGrupo] || "").toString();
      var grupoLinhaNorm = normalizarTexto_(grupoLinhaOriginal);
      if (grupoNorm && (!grupoLinhaNorm ||
           (grupoLinhaNorm.indexOf(grupoNorm) === -1 &&
            grupoNorm.indexOf(grupoLinhaNorm) === -1))) {
        continue;
      }

      var cat = (row[idxCategoria] || "Sem categoria").toString().trim();
      if (!cat) cat = "Sem categoria";

      if (!mapa[cat]) {
        mapa[cat] = { categoria: cat, total: 0, valorTotal: 0 };
      }
      mapa[cat].total++;
      var valor = Number(row[idxValor]) || 0;
      mapa[cat].valorTotal += valor;
    }

    var arr = [];
    for (var k in mapa) {
      if (Object.prototype.hasOwnProperty.call(mapa, k)) {
        arr.push(mapa[k]);
      }
    }

    if (criterio === "valor") {
      arr.sort(function (a, b) {
        if (b.valorTotal !== a.valorTotal) return b.valorTotal - a.valorTotal;
        return b.total - a.total;
      });
    } else {
      arr.sort(function (a, b) {
        if (b.total !== a.total) return b.total - a.total;
        return b.valorTotal - a.valorTotal;
      });
    }

    var top = arr.length ? arr[0] : null;

    return {
      ok: true,
      criterio: criterio,
      grupoOriginal: grupoOriginal,
      categorias: arr,
      top: top
    };

  } catch (e) {
    return { ok: false, error: String((e && e.message) ? e.message : e) };
  }
}

/**
 * Resumo de categorias filtrando por LOJA (LojaNum).
 *
 * @param {string} dataInicioStr ISO ou vazio
 * @param {string} dataFimStr ISO ou vazio
 * @param {string} criterio "valor" | "quantidade"
 * @param {string} lojaCodigo Código da loja (com ou sem zeros à esquerda)
 */
function getResumoTransacoesPorCategoriaLoja(dataInicioStr, dataFimStr, criterio, lojaCodigo) {
  vektorAssertFunctionAllowed_("getResumoTransacoesPorCategoriaLoja");
  try {
    var info = carregarLinhasBaseClara_();
    if (info.error) {
      return { ok: false, error: info.error };
    }

    var header = info.header;
    var linhas = info.linhas;

    var idxData = encontrarIndiceColuna_(header, [
      "Data da Transação", "Data Transação", "Data"
    ]);

    var idxValor = encontrarIndiceColuna_(header, [
      "Valor em R$", "Valor (R$)", "Valor"
    ]);

    var idxCategoria = encontrarIndiceColuna_(header, [
      "Categoria da Compra", "Categoria da compra", "Categoria", "Categoria Compra"
    ]);

    var idxLoja = encontrarIndiceColuna_(header, [
      "LojaNum", "Loja Num", "Loja Número", "Loja Numero", "Loja"
    ]);

    if (idxData < 0 || idxValor < 0 || idxCategoria < 0 || idxLoja < 0) {
      return {
        ok: false,
        error: "Não encontrei as colunas necessárias em BaseClara (Data / Valor / Categoria / Loja)."
      };
    }

    criterio = (criterio || "").toString().toLowerCase();
    if (criterio !== "valor" && criterio !== "quantidade") {
      criterio = "quantidade";
    }

    var lojaOriginal = (lojaCodigo || "").toString().trim();
    var lojaDigits = lojaOriginal.replace(/\D/g, "");
    var lojaNormalizada = lojaDigits ? ("0000" + lojaDigits).slice(-4) : "";

    var filtradas = filtrarLinhasPorPeriodo_(linhas, idxData, dataInicioStr, dataFimStr);

    var mapa = {};
    for (var i = 0; i < filtradas.length; i++) {
      var row = filtradas[i];

      // filtro por loja
      if (lojaNormalizada) {
        var lojaLinha = (row[idxLoja] || "").toString();
        var digitsLinha = lojaLinha.replace(/\D/g, "");
        var cod4 = digitsLinha ? ("0000" + digitsLinha).slice(-4) : "";
        if (cod4 !== lojaNormalizada) continue;
      }

      var cat = (row[idxCategoria] || "Sem categoria").toString().trim();
      if (!cat) cat = "Sem categoria";

      if (!mapa[cat]) {
        mapa[cat] = { categoria: cat, total: 0, valorTotal: 0 };
      }
      mapa[cat].total++;
      var valor = Number(row[idxValor]) || 0;
      mapa[cat].valorTotal += valor;
    }

    var arr = [];
    for (var k in mapa) {
      if (Object.prototype.hasOwnProperty.call(mapa, k)) {
        arr.push(mapa[k]);
      }
    }

    if (criterio === "valor") {
      arr.sort(function (a, b) {
        if (b.valorTotal !== a.valorTotal) return b.valorTotal - a.valorTotal;
        return b.total - a.total;
      });
    } else {
      arr.sort(function (a, b) {
        if (b.total !== a.total) return b.total - a.total;
        return b.valorTotal - a.valorTotal;
      });
    }

    var top = arr.length ? arr[0] : null;

    return {
      ok: true,
      criterio: criterio,
      lojaOriginal: lojaOriginal,
      categorias: arr,
      top: top
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : e
    };
  }
}

/**
 * Resumo de transações por CATEGORIA, filtrando por LOJA específica.
 *
 * @param {string} dataInicioStr ISO ou vazio
 * @param {string} dataFimStr    ISO ou vazio
 * @param {string} criterio      "valor" | "quantidade"
 * @param {string} loja          código da loja (ex.: "0297" ou "297")
 */
function getResumoTransacoesPorCategoriaLoja(dataInicioStr, dataFimStr, criterio, loja) {
  vektorAssertFunctionAllowed_("getResumoTransacoesPorCategoriaLoja");
  try {
    var info = carregarLinhasBaseClara_();
    if (info.error) {
      return { ok: false, error: info.error };
    }

    var header = info.header;
    var linhas = info.linhas;

    // Índices das colunas
    var idxData = encontrarIndiceColuna_(header, [
      "Data da Transação",
      "Data Transação",
      "Data"
    ]);

    var idxValor = encontrarIndiceColuna_(header, [
      "Valor em R$",
      "Valor (R$)",
      "Valor"
    ]);

    var idxCategoria = encontrarIndiceColuna_(header, [
      "Categoria da Compra",
      "Categoria da compra",
      "Categoria",
      "Categoria Compra"
    ]);

    var idxLoja = encontrarIndiceColuna_(header, [
      "LojaNum",
      "Loja Num",
      "Loja Número",
      "Loja Numero",
      "Loja"
    ]);

    if (idxData < 0 || idxValor < 0 || idxCategoria < 0 || idxLoja < 0) {
      return {
        ok: false,
        error: "Não encontrei as colunas necessárias em BaseClara (Data / Valor / Categoria / Loja)."
      };
    }

    // normaliza critério
    criterio = (criterio || "").toString().toLowerCase();
    if (criterio !== "valor" && criterio !== "quantidade") {
      criterio = "quantidade";
    }

    // normaliza loja informada
    var lojaOriginal = (loja || "").toString().trim();
    var lojaDigits = lojaOriginal.replace(/\D/g, "");
    var lojaNormalizada = lojaDigits ? ("0000" + lojaDigits).slice(-4) : "";

    // filtra por período
    var filtradas = filtrarLinhasPorPeriodo_(linhas, idxData, dataInicioStr, dataFimStr);

    var mapa = {}; // chave = categoria
    for (var i = 0; i < filtradas.length; i++) {
      var row = filtradas[i];

      // filtro por loja (se veio parâmetro)
      if (lojaNormalizada) {
        var lojaLinha = (row[idxLoja] || "").toString();
        var digitsLinha = lojaLinha.replace(/\D/g, "");
        var cod4 = digitsLinha ? ("0000" + digitsLinha).slice(-4) : "";
        if (cod4 !== lojaNormalizada) continue;
      }

      var cat = (row[idxCategoria] || "Sem categoria").toString().trim();
      if (!cat) cat = "Sem categoria";

      if (!mapa[cat]) {
        mapa[cat] = {
          categoria: cat,
          total: 0,
          valorTotal: 0
        };
      }

      mapa[cat].total++;
      var valor = Number(row[idxValor]) || 0;
      mapa[cat].valorTotal += valor;
    }

    var arr = [];
    for (var chave in mapa) {
      if (Object.prototype.hasOwnProperty.call(mapa, chave)) {
        arr.push(mapa[chave]);
      }
    }

    // ordena conforme o critério
    if (criterio === "valor") {
      arr.sort(function (a, b) {
        if (b.valorTotal !== a.valorTotal) return b.valorTotal - a.valorTotal;
        return b.total - a.total;
      });
    } else {
      arr.sort(function (a, b) {
        if (b.total !== a.total) return b.total - a.total;
        return b.valorTotal - a.valorTotal;
      });
    }

    var top = arr.length ? arr[0] : null;

    return {
      ok: true,
      criterio: criterio,
      lojaOriginal: lojaOriginal,
      categorias: arr,
      top: top
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : e
    };
  }
}

function getTransacoesIndividuaisPorEstabelecimento(dataInicioStr, dataFimStr, estabelecimento) {
  vektorAssertFunctionAllowed_("getTransacoesIndividuaisPorEstabelecimento");
  try {
    var info = carregarLinhasBaseClara_();
    if (!info || info.error) {
      return { ok: false, error: info && info.error ? info.error : "Não foi possível ler a BaseClara." };
    }

    var header = info.header || [];
    var linhas = info.linhas || [];

    // Índices
    var idxData  = encontrarIndiceColuna_(header, ["Data da Transação", "Data Transação", "Data"]);
    var idxValor = encontrarIndiceColuna_(header, ["Valor em R$", "Valor (R$)", "Valor"]);
    var idxLoja  = encontrarIndiceColuna_(header, ["LojaNum", "Loja Num", "Loja Número", "Loja Numero", "Loja"]);

    // Coluna C fixa (Transação / nome do estabelecimento) = índice 2
    var idxTransacao = 2;

    var idxRecibo = encontrarIndiceColuna_(header, ["Recibo"]);
    if (idxRecibo < 0) idxRecibo = 14; // O

    var idxTitular = encontrarIndiceColuna_(header, ["Titular"]);
    if (idxTitular < 0) idxTitular = 16; // Q

    var idxGrupos = encontrarIndiceColuna_(header, ["Grupos", "Grupo", "Time"]);
    if (idxGrupos < 0) idxGrupos = 17; // R

    var idxEtiquetas = encontrarIndiceColuna_(header, ["Etiquetas", "Etiqueta"]);
    if (idxEtiquetas < 0) idxEtiquetas = 19; // T

    var idxDescricao = encontrarIndiceColuna_(header, ["Descrição", "Descricao"]);
    if (idxDescricao < 0) idxDescricao = 20; // U

    if (idxData < 0 || idxValor < 0 || idxLoja < 0 || idxTransacao < 0) {
      return { ok: false, error: "Não encontrei colunas necessárias (Data/Valor/LojaNum/Transação) na BaseClara." };
    }

    // Período default (últimos 30 dias) se vier vazio
    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    var hoje = new Date();
    var iniDate, fimDate;

    if (dataInicioStr && dataFimStr) {
      iniDate = new Date(dataInicioStr);
      fimDate = new Date(dataFimStr);
    } else {
      fimDate = hoje;
      iniDate = new Date();
      iniDate.setDate(hoje.getDate() - 30);
    }

    var dataInicioIso = iniDate.toISOString();
    var dataFimIso    = fimDate.toISOString();

    // Filtra período
    var filtradas = filtrarLinhasPorPeriodo_(linhas, idxData, dataInicioIso, dataFimIso) || [];
    if (!filtradas.length) {
      return {
        ok: true,
        linhas: [],
        estabelecimento: estabelecimento || "",
        dataInicioIso: dataInicioIso,
        dataFimIso: dataFimIso
      };
    }

    // Normaliza estabelecimento (como o clique vem da própria tabela, normalmente bate exato)
    var estabNorm = normalizarTexto_((estabelecimento || "").toString().trim());

    var lista = [];
    var contPorLoja = {};

    filtradas.forEach(function(row) {
      if (!row) return;

      var estabLinha = (row[idxTransacao] || "").toString().trim();
      if (!estabLinha) return;

      // match por normalização (igualdade)
      var estabLinhaNorm = normalizarTexto_(estabLinha);
      if (estabNorm && estabLinhaNorm !== estabNorm) return;

      var valor = Number(row[idxValor]) || 0;

      var loja = (row[idxLoja] || "").toString().trim();

      // data BR
      var dataCel = row[idxData];
      var dataBr = "";
      if (dataCel instanceof Date) {
        dataBr = Utilities.formatDate(dataCel, tz, "dd/MM/yyyy");
      } else {
        dataBr = (dataCel || "").toString();
      }

      // conta por loja (para ordenação por qtd)
      var lojaKey = loja || "—";
      contPorLoja[lojaKey] = (contPorLoja[lojaKey] || 0) + 1;

      lista.push({
        loja: loja,
        estabelecimento: estabLinha,
        data: dataBr,
        valor: valor,
         // novos
        titular: String(row[idxTitular] || ""),
        grupos: String(row[idxGrupos] || ""),
        recibo: String(row[idxRecibo] || ""),
        etiquetas: String(row[idxEtiquetas] || ""),
        descricao: String(row[idxDescricao] || "")
      });
    });

    if (!lista.length) {
      return {
        ok: true,
        linhas: [],
        estabelecimento: estabelecimento || "",
        dataInicioIso: dataInicioIso,
        dataFimIso: dataFimIso
      };
    }

    // Ordena por: (1) loja com mais transações, (2) valor desc, (3) data
    lista.sort(function(a, b) {
      var ca = contPorLoja[a.loja || "—"] || 0;
      var cb = contPorLoja[b.loja || "—"] || 0;
      if (cb !== ca) return cb - ca;

      var va = Number(a.valor) || 0;
      var vb = Number(b.valor) || 0;
      if (vb !== va) return vb - va;

      return (a.data || "").localeCompare(b.data || "");
    });

    return {
      ok: true,
      linhas: lista,
      estabelecimento: estabelecimento || "",
      dataInicioIso: dataInicioIso,
      dataFimIso: dataFimIso
    };

  } catch (e) {
    return { ok: false, error: String((e && e.message) ? e.message : e) };
  }
}

/**
 * Resumo de transações por ESTABELECIMENTO (coluna "Transação" da BaseClara).
 *
 * - dataInicioStr / dataFimStr: datas em ISO (como nas outras funções)
 * - criterio: "valor" ou "quantidade"
 * - grupo: nome do time/grupo (opcional) para filtrar
 * - loja: código da loja (opcional) para filtrar
 *
 * Retorna:
 * {
 *   ok: true,
 *   criterio: "valor" ou "quantidade",
 *   grupoOriginal: "...",   // se informado
 *   lojaOriginal:  "...",   // se informado
 *   estabelecimentos: [
 *     { estabelecimento: "IFood", total: 15, valorTotal: 2000.50 },
 *     ...
 *   ],
 *   top: { ... } // estabelecimento campeão
 * }
 */
function getResumoTransacoesPorEstabelecimento(
  dataInicioStr,
  dataFimStr,
  criterio,
  grupo,
  loja
) {
  vektorAssertFunctionAllowed_("getResumoTransacoesPorEstabelecimento");
  try {
    var info = carregarLinhasBaseClara_();
    if (info.error) {
      return { ok: false, error: info.error };
    }

    var header = info.header;
    var linhas = info.linhas;

    // Descobre índices dinamicamente
    var idxData = encontrarIndiceColuna_(header, [
      "Data da Transação",
      "Data Transação",
      "Data"
    ]);

    var idxValor = encontrarIndiceColuna_(header, [
      "Valor em R$",
      "Valor (R$)",
      "Valor"
    ]);

    // ========== COLUNA TRANSACAO ==========

// A coluna C (Transação) é SEMPRE índice 2 → solução definitiva
var idxTransacao = 2;


// DEBUG PARA VERIFICAR
Logger.log("IDX TRANSACAO = " + idxTransacao);
Logger.log("VALOR TRANSACAO PRIMEIRA LINHA = " + linhas[0][idxTransacao]);


    // Grupo e Loja são opcionais (só se quiser filtrar)
    var idxGrupo = encontrarIndiceColuna_(header, [
      "Grupos",
      "Grupo",
      "Time"
    ]);

    var idxLoja = encontrarIndiceColuna_(header, [
      "LojaNum",
      "Loja Num",
      "Loja Número",
      "Loja Numero",
      "Loja"
    ]);

    if (idxData < 0 || idxValor < 0 || idxTransacao < 0) {
      return {
        ok: false,
        error: "Não encontrei as colunas necessárias em BaseClara (Data / Valor / Transação)."
      };
    }

    criterio = (criterio || "").toString().toLowerCase();
    if (criterio !== "valor" && criterio !== "quantidade") {
      criterio = "quantidade";
    }

    // Normaliza filtros de grupo e loja
    var grupoOriginal = (grupo || "").toString().trim();
    var grupoNorm = normalizarTexto_(grupoOriginal);

    var lojaOriginal = (loja || "").toString().trim();
    var lojaDigits = lojaOriginal.replace(/\D/g, "");
    var lojaNormalizada = lojaDigits ? ("0000" + lojaDigits).slice(-4) : "";

    // Filtra período
    var filtradas = filtrarLinhasPorPeriodo_(linhas, idxData, dataInicioStr, dataFimStr);

    var mapa = {}; // chave = nome do estabelecimento
    for (var i = 0; i < filtradas.length; i++) {
      var row = filtradas[i];

      // filtro por grupo/time (se informado e se a coluna existir)
      if (grupoNorm && idxGrupo >= 0) {
        var grupoLinhaOriginal = (row[idxGrupo] || "").toString();
        var grupoLinhaNorm = normalizarTexto_(grupoLinhaOriginal);

        if (!grupoLinhaNorm) continue;

        var casaGrupo =
          grupoLinhaNorm.indexOf(grupoNorm) !== -1 ||
          grupoNorm.indexOf(grupoLinhaNorm) !== -1;

        if (!casaGrupo) continue;
      }

      // filtro por loja (se informado e se a coluna existir)
      if (lojaNormalizada && idxLoja >= 0) {
        var lojaLinha = (row[idxLoja] || "").toString();
        var digitsLinha = lojaLinha.replace(/\D/g, "");
        var cod4 = digitsLinha ? ("0000" + digitsLinha).slice(-4) : "";
        if (cod4 !== lojaNormalizada) continue;
      }

      var estab = (row[idxTransacao] || "Sem nome").toString().trim();
      if (!estab) estab = "Sem nome";

      if (!mapa[estab]) {
        mapa[estab] = {
          estabelecimento: estab,
          total: 0,
          valorTotal: 0
        };
      }

      mapa[estab].total++;
      var valor = Number(row[idxValor]) || 0;
      mapa[estab].valorTotal += valor;
    }

    // transforma o mapa em array
    var arr = [];
    for (var k in mapa) {
      if (Object.prototype.hasOwnProperty.call(mapa, k)) {
        arr.push(mapa[k]);
      }
    }

    // ordena pelo critério escolhido
    if (criterio === "valor") {
      arr.sort(function (a, b) {
        if (b.valorTotal !== a.valorTotal) return b.valorTotal - a.valorTotal;
        return b.total - a.total;
      });
    } else {
      arr.sort(function (a, b) {
        if (b.total !== a.total) return b.total - a.total;
        return b.valorTotal - a.valorTotal;
      });
    }

    var top = arr.length ? arr[0] : null;

    return {
      ok: true,
      criterio: criterio,
      grupoOriginal: grupoOriginal,
      lojaOriginal: lojaOriginal,
      estabelecimentos: arr,
      top: top,
      // 🔹 período usado no cálculo (vai para o front)
      dataInicioIso: dataInicioStr || "",
      dataFimIso:    dataFimStr    || ""
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : e
    };
  }
}

function getResumoLojasPorEstabelecimento(estabelecimento, dataIni, dataFim) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID_CLARA);
  const sh = ss.getSheetByName("BaseClara");
  const valores = sh.getDataRange().getValues();
  const header  = valores.shift();

  const idxLoja = 0;
  const idxEst  = 2;
  const idxData = 3;
  const idxVal  = 10;

  const out = {};

  valores.forEach(l => {
    const est = l[idxEst];
    if (!est || est.toString().trim() !== estabelecimento.toString().trim()) return;

    const data = new Date(l[idxData]);
    if (dataIni && data < dataIni) return;
    if (dataFim && data > dataFim) return;

    const loja = l[idxLoja];

    if (!out[loja]) out[loja] = { loja, qtd: 0, valor: 0 };
    out[loja].qtd++;
    out[loja].valor += Number(l[idxVal]);
  });

  return Object.values(out).sort((a,b)=>b.valor-a.valor);
}

function getResumoLojasPorCategoria(categoria, dataIni, dataFim) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID_CLARA);
  const sh = ss.getSheetByName("BaseClara");
  const valores = sh.getDataRange().getValues();
  const header  = valores.shift();

  const idxLoja = 0;
  const idxCat  = 8;
  const idxData = 3;
  const idxVal  = 10;

  const out = {};

  valores.forEach(l => {
    const cat = l[idxCat];
    if (!cat || cat.toString().trim() !== categoria.toString().trim()) return;

    const data = new Date(l[idxData]);
    if (dataIni && data < dataIni) return;
    if (dataFim && data > dataFim) return;

    const loja = l[idxLoja];

    if (!out[loja]) out[loja] = { loja, qtd: 0, valor: 0 };
    out[loja].qtd++;
    out[loja].valor += Number(l[idxVal]);
  });

  return Object.values(out).sort((a,b)=>b.valor-a.valor);
}

/**
 * Retorna lista de lojas para autocomplete:
 * [{ codigo: "0297", nome: "CATUAÍ CASCAVEL" }, ...]
 */

function getListaLojas() {
  vektorAssertFunctionAllowed_("getListaLojas");
  try {
    var info = carregarLinhasBaseClara_();
    if (info.error) return [];

    var header = info.header;
    var linhas = info.linhas;

    // ✅ ACL por Emails somente para Gerentes_Reg
    var ctx = vektorGetUserRole_(); // deve retornar {email, role}
    var role = String((ctx && ctx.role) || "").trim();
    var email = String((ctx && ctx.email) || "").trim().toLowerCase();

    var allowedLojas = null;
    if (role === "Gerentes_Reg") {
      allowedLojas = vektorGetAllowedLojasFromEmails_(email); // array ou null
    }

    var allowedSet = null;
    if (Array.isArray(allowedLojas)) {
      allowedSet = {};
      allowedLojas.forEach(function(x){
        x = String(x || "").trim();
        if (!x) return;
        allowedSet[x] = true;
        allowedSet[x.padStart(4, "0")] = true;
      });
    }

    // 1) Índice da coluna do código da loja (continua dinâmico)
    var idxLoja = encontrarIndiceColuna_(header, [
      "LojaNum", "Loja Num", "Loja Número", "Loja Numero", "Loja", "cod_loja", "codLoja"
    ]);

    if (idxLoja < 0) return [];

    // 2) Índice da coluna "Descrição Loja"
    var idxNome = header.indexOf("Descrição Loja");
    if (idxNome < 0) idxNome = header.indexOf("Descricao Loja");

    var temNome = idxNome >= 0;

    var mapa = {};

    linhas.forEach(function (row) {
      var codRaw = (row[idxLoja] || "").toString().trim();
      if (!codRaw) return;

      var digits = codRaw.replace(/\D/g, "");
      if (!digits) return;

      var cod4 = ("0000" + digits).slice(-4);

      // ✅ aplica ACL (só Gerentes_Reg)
      if (allowedSet && !allowedSet[cod4] && !allowedSet[String(Number(cod4) || "").trim()]) return;

      var nome = "";
      if (temNome) {
        nome = (row[idxNome] || "").toString().trim();
      }

      mapa[cod4] = nome;
    });

    var out = [];
    for (var c in mapa) {
      if (Object.prototype.hasOwnProperty.call(mapa, c)) {
        out.push({ codigo: c, nome: mapa[c] });
      }
    }

    out.sort(function (a, b) { return a.codigo.localeCompare(b.codigo); });
    return out;

  } catch (e) {
    return [];
  }
}

function getListaEtiquetasClara() {
  try {
    var info = carregarLinhasBaseClara_();
    if (info.error) return [];

    var header = info.header;
    var linhas = info.linhas;

    // Procurar EXATAMENTE a coluna "Etiquetas" (coluna T da BaseClara)
    var idxEtiqueta = header.indexOf("Etiquetas");
    if (idxEtiqueta < 0) {
      // fallback simples, se algum dia mudar para "Etiqueta"
      idxEtiqueta = header.indexOf("Etiqueta");
    }

    if (idxEtiqueta < 0) {
      // não achou a coluna de etiquetas
      return [];
    }

    // mapa para garantir apenas UMA ocorrência de cada valor de célula,
    // sem alterar o texto
    var mapa = {};

    linhas.forEach(function (row) {
      var valor = row[idxEtiqueta];
      if (valor === null || valor === undefined) return;

      // mantém exatamente como está na planilha
      valor = valor.toString();

      // se quiser ignorar células que sejam só espaços, descomente a linha abaixo:
      // if (valor.trim() === "") return;

      if (!mapa.hasOwnProperty(valor)) {
        mapa[valor] = true;
      }
    });

    // converte as chaves do mapa em array de etiquetas "cruas"
    var out = Object.keys(mapa);

    // ordena alfabeticamente (sem mexer no conteúdo)
    out.sort(function (a, b) {
      return a.localeCompare(b, "pt-BR");
    });

    return out;

  } catch (e) {
    return [];
  }
}

// =====================================================
// GASTOS POR ETIQUETAS (BaseClara) - por indices fixos
// =====================================================

// Indices fixos conforme solicitado (A=0, C=2, D=3, H=7, R=17, T=19, U=20)
var _ETQ_IDX_DATA_  = 0;   // A
var _ETQ_IDX_ESTAB_ = 2;   // C
var _ETQ_IDX_VALOR_ = 3;   // D
var _ETQ_IDX_LOJA_  = 7;   // H
var _ETQ_IDX_TIME_  = 17;  // R
var _ETQ_IDX_TAGS_  = 19;  // T
var _ETQ_IDX_DESC_  = 20;  // U

function _parseDataToDate_(v) {
  if (!v) return null;
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) return v;

  var s = String(v).trim();
  if (!s) return null;

  // dd/MM/yyyy
  var m1 = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
  if (m1) return new Date(Number(m1[3]), Number(m1[2]) - 1, Number(m1[1]));

  // yyyy-MM-dd
  var m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (m2) return new Date(Number(m2[1]), Number(m2[2]) - 1, Number(m2[3]));

  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}

function _fmtMes_(dt, tz) {
  try {
    return Utilities.formatDate(dt, tz, "MM/yyyy");
  } catch (e) {
    return "";
  }
}

function _toNumberValor_(v) {
  if (v === null || v === undefined) return 0;
  if (typeof v === "number") return isFinite(v) ? v : 0;
  var s = String(v).trim();
  if (!s) return 0;
  // remove separador milhar e troca vírgula por ponto
  s = s.replace(/\./g, "").replace(",", ".");
  var n = parseFloat(s);
  return isFinite(n) ? n : 0;
}

function _splitTags_(cell) {
  var txt = (cell === null || cell === undefined) ? "" : String(cell);
  txt = txt.trim();
  if (!txt) return [];
  return txt.split("|").map(function (x) { return String(x || "").trim(); }).filter(Boolean);
}

function _rowHasTag_(row, tag) {
  if (!tag) return true;
  var tags = _splitTags_(row[_ETQ_IDX_TAGS_]);
  for (var i = 0; i < tags.length; i++) {
    if (tags[i] === tag) return true;
  }
  return false;
}

function _keysSorted_(obj) {
  return Object.keys(obj || {}).sort(function (a, b) { return a.localeCompare(b, "pt-BR"); });
}

function _keysSortedMes_(obj) {
  function parseMes_(s) {
    // aceita "MM/yyyy" e tenta tolerar "M/yyyy"
    var m = String(s || "").match(/^(\d{1,2})\/(\d{4})$/);
    if (!m) return { y: 0, mo: 0 };
    return { mo: Number(m[1]) || 0, y: Number(m[2]) || 0 };
  }

  return Object.keys(obj || {}).sort(function(a, b) {
    var pa = parseMes_(a);
    var pb = parseMes_(b);
    if (pa.y !== pb.y) return pa.y - pb.y;
    return pa.mo - pb.mo;
  });
}

/**
 * Retorna:
 * - itens: [{ etiqueta, valorTotal, percentual }]
 * - somaValores: soma total de valores (coluna D) das transações filtradas (base)
 * - totalGeral: soma dos valores alocados por etiqueta (atenção: pode "duplicar" se houver múltiplas etiquetas por transação)
 * - filtros dependentes: meses/times/lojas/etiquetas
 */
function getGastosPorEtiquetasClara(filtro) {
  vektorAssertFunctionAllowed_("getGastosPorEtiquetasClara");
  try {
    filtro = filtro && typeof filtro === "object" ? filtro : {};
    var fMes = String(filtro.mes || "").trim();       // "MM/yyyy"
    var fTime = String(filtro.time || "").trim();
    var fLoja = String(filtro.loja || "").trim();
    var fTag = String(filtro.etiqueta || "").trim();

    var fIni = String(filtro.periodoIni || "").trim(); // "YYYY-MM-DD"
    var fFim = String(filtro.periodoFim || "").trim(); // "YYYY-MM-DD"

    function _parseISODate_(s) {
      if (!s) return null;
      var m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (!m) return null;
      return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 0, 0, 0, 0);
    }

    var dtIni = _parseISODate_(fIni);
    var dtFim = _parseISODate_(fFim);
    if (dtFim) dtFim = new Date(dtFim.getFullYear(), dtFim.getMonth(), dtFim.getDate(), 23, 59, 59, 999);

    // Reaproveita seu loader existente
    var info = carregarLinhasBaseClara_();
    if (info && info.error) return { ok: false, error: info.error };

    var linhas = (info && info.linhas) ? info.linhas : [];
    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";

    // sets dependentes
    var setMes = {}, setTime = {}, setLoja = {}, setTag = {};

    // agregação
    var mapa = {};       // etiqueta -> valorTotal
    var somaValores = 0; // soma base (coluna D) das transações filtradas (sem duplicar por etiqueta)
    var totalGeral = 0;  // soma alocada por etiqueta (pode duplicar quando multi-etiqueta)

    for (var i = 0; i < linhas.length; i++) {
      var row = linhas[i];
      if (!row) continue;

      var dt = _parseDataToDate_(row[_ETQ_IDX_DATA_]);
      if (!dt) continue;

      var mes = _fmtMes_(dt, tz);
      if (!mes) continue;

      if (dtIni && dt < dtIni) continue;
      if (dtFim && dt > dtFim) continue;

      var time = String(row[_ETQ_IDX_TIME_] || "").trim();
      var loja = String(row[_ETQ_IDX_LOJA_] || "").trim();
      loja = loja ? loja.replace(/\D/g, "") : "";
      if (loja) loja = ("0000" + loja).slice(-4);

      var tags = _splitTags_(row[_ETQ_IDX_TAGS_]);
      var temTagSelecionada = _rowHasTag_(row, fTag);

      // -------------------------
      // Filtros dependentes:
      // cada conjunto ignora o próprio filtro e respeita os demais
      // -------------------------

      // meses disponíveis (respeita time/loja/etiqueta)
      if ((!fTime || time === fTime) && (!fLoja || loja === fLoja) && temTagSelecionada) {
        setMes[mes] = true;
      }

      // times disponíveis (respeita mes/loja/etiqueta)
      if ((!fMes || mes === fMes) && (!fLoja || loja === fLoja) && temTagSelecionada) {
        if (time) setTime[time] = true;
      }

      // lojas disponíveis (respeita mes/time/etiqueta)
      if ((!fMes || mes === fMes) && (!fTime || time === fTime) && temTagSelecionada) {
        if (loja) setLoja[loja] = true;
      }

      // etiquetas disponíveis (respeita mes/time/loja; ignora filtro etiqueta)
      if ((!fMes || mes === fMes) && (!fTime || time === fTime) && (!fLoja || loja === fLoja)) {
        for (var t = 0; t < tags.length; t++) setTag[tags[t]] = true;
      }

      // -------------------------
      // Aplicação dos filtros para o resumo (respeita TODOS)
      // -------------------------
      if (fMes && mes !== fMes) continue;
      if (fTime && time !== fTime) continue;
      if (fLoja && loja !== fLoja) continue;
      if (fTag && !temTagSelecionada) continue;

      var valorNum = _toNumberValor_(row[_ETQ_IDX_VALOR_]);
      if (!isFinite(valorNum)) valorNum = 0;

      somaValores += valorNum;

      // aloca por etiqueta (se não filtrou etiqueta, soma para todas as tags da transação)
      for (var k = 0; k < tags.length; k++) {
        var et = tags[k];
        if (fTag && et !== fTag) continue;

        if (!mapa[et]) mapa[et] = 0;
        mapa[et] += valorNum;
        totalGeral += valorNum;
      }
    }

    var itens = [];
    Object.keys(mapa).forEach(function (et) {
      itens.push({ etiqueta: et, valorTotal: Number(mapa[et] || 0) });
    });
    itens.sort(function (a, b) { return (b.valorTotal || 0) - (a.valorTotal || 0); });

    itens.forEach(function (it) {
      var p = totalGeral > 0 ? (it.valorTotal / totalGeral) * 100 : 0;
      it.percentual = p;
    });

    return {
      ok: true,
      itens: itens,
      somaValores: somaValores,
      totalGeral: totalGeral,
      filtros: {
        meses: _keysSortedMes_(setMes),
        times: _keysSorted_(setTime),
        lojas: _keysSorted_(setLoja),
        etiquetas: _keysSorted_(setTag)
      }
    };
  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

/**
 * Tabela detalhada por etiqueta (respeita filtros).
 * Retorna colunas conforme solicitado:
 * Loja(H), Time(R), Data(A), Estabelecimento(C), Valor(D), Etiqueta(T), Descrição(U)
 */
function getTransacoesPorEtiquetaClara(payload) {
  try {
    payload = payload && typeof payload === "object" ? payload : {};
    var fMes = String(payload.mes || "").trim();
    var fTime = String(payload.time || "").trim();
    var fLoja = String(payload.loja || "").trim();
    var tagSel = String(payload.etiqueta || "").trim();
    var fIni = String(payload.periodoIni || "").trim();
    var fFim = String(payload.periodoFim || "").trim();

      function _parseISODate_(s) {
        if (!s) return null;
        var m = String(s).match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (!m) return null;
        return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]), 0, 0, 0, 0);
      }

      var dtIni = _parseISODate_(fIni);
      var dtFim = _parseISODate_(fFim);
      if (dtFim) dtFim = new Date(dtFim.getFullYear(), dtFim.getMonth(), dtFim.getDate(), 23, 59, 59, 999);

    if (!tagSel) return { ok: true, rows: [] };

    var info = carregarLinhasBaseClara_();
    if (info && info.error) return { ok: false, error: info.error };

    var linhas = (info && info.linhas) ? info.linhas : [];
    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";

    var out = [];
    for (var i = 0; i < linhas.length; i++) {
      var row = linhas[i];
      if (!row) continue;

      var dt = _parseDataToDate_(row[_ETQ_IDX_DATA_]);
      if (!dt) continue;

      if (dtIni && dt < dtIni) continue;
      if (dtFim && dt > dtFim) continue;

      var mes = _fmtMes_(dt, tz);
      if (fMes && mes !== fMes) continue;

      var time = String(row[_ETQ_IDX_TIME_] || "").trim();
      if (fTime && time !== fTime) continue;

      var loja = String(row[_ETQ_IDX_LOJA_] || "").trim();
      loja = loja ? loja.replace(/\D/g, "") : "";
      if (loja) loja = ("0000" + loja).slice(-4);
      if (fLoja && loja !== fLoja) continue;

      if (!_rowHasTag_(row, tagSel)) continue;

      var estab = String(row[_ETQ_IDX_ESTAB_] || "").trim();
      var valorNum = _toNumberValor_(row[_ETQ_IDX_VALOR_]);
      var etiquetaCell = String(row[_ETQ_IDX_TAGS_] || "").trim();
      var desc = String(row[_ETQ_IDX_DESC_] || "").trim();

      out.push({
        loja: loja,
        time: time,
        data: Utilities.formatDate(dt, tz, "dd/MM/yyyy"),
        estabelecimento: estab,
        valor: valorNum,
        etiqueta: etiquetaCell,
        descricao: desc
      });

      if (out.length >= 1500) break;
    }

    return { ok: true, etiqueta: tagSel, rows: out };
  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

/**
 * Envia a tabela detalhada por e-mail
 * Assunto: Base de Gastos por Etiquetas
 * Remetente: Vektor - Grupo SBF
 */
function enviarEmailGastosPorEtiquetasClara(payload) {
  vektorAssertFunctionAllowed_("enviarEmailGastosPorEtiquetasClara");
  try {
    payload = payload && typeof payload === "object" ? payload : {};

    var emailUsuario = Session.getActiveUser().getEmail();
    if (!emailUsuario) return { ok: false, error: "Usuário sem e-mail ativo." };

    // destinatário vindo do front (modal). fallback: se não vier, manda para o próprio usuário
    var emailDestino = String(payload.emailDestino ? payload.emailDestino : emailUsuario).trim();

    // 🔒 trava domínio: apenas @gruposbf.com.br ou @centauro.com.br
    var emailRegex = /^[^\s@]+@((gruposbf|centauro|fisia)\.com\.br)$/i;
    if (!emailRegex.test(emailDestino)) {
      return { ok: false, error: "E-mail inválido. Use apenas @gruposbf.com.br, @centauro.com.br ou @fisia.com.br." };
    }

    // CC: por padrão o usuário logado, exceto quando ele é o próprio destinatário
    var ccEmail = "";
    if (emailDestino.toLowerCase() !== emailUsuario.toLowerCase()) {
      ccEmail = emailUsuario;
    }

    var det = getTransacoesPorEtiquetaClara(payload);
    if (!det || !det.ok) return { ok: false, error: (det && det.error) ? det.error : "Falha ao montar base." };

    var rows = det.rows || [];
    if (!rows.length) return { ok: false, error: "Sem transações para enviar com os filtros atuais." };

    function esc_(x) {
      return String(x === null || x === undefined ? "" : x)
        .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;").replace(/'/g, "&#039;");
    }

    var assunto = "Base de Gastos por Etiquetas";

    var cab = "<p style='margin:0'>Segue a base detalhada de <b>Gastos por Etiquetas</b>.</p>";
    cab += "<p style='margin:0;margin-top:6px;font-size:12px;color:#475569'>Etiqueta selecionada: <b>" + esc_(det.etiqueta || "") + "</b></p>";

    var th = "background:#003366;color:#fff;border:1px solid #0f172a;padding:6px;white-space:nowrap;text-align:left;";
    var td = "border:1px solid #0f172a;padding:6px;vertical-align:top;";

    var t = "";
    t += "<div style='margin-top:10px'>";
    t += "<table style='border-collapse:collapse;width:100%;font-family:Arial,sans-serif;font-size:12px;'>";
    t += "<thead><tr>";
    t += "<th style='" + th + "'>Loja</th>";
    t += "<th style='" + th + "'>Time</th>";
    t += "<th style='" + th + "'>Data</th>";
    t += "<th style='" + th + "text-align:left;'>Estabelecimento</th>";
    t += "<th style='" + th + "text-align:right;'>Valor</th>";
    t += "<th style='" + th + "'>Etiqueta inserida</th>";
    t += "<th style='" + th + "'>Item comprado (descrição)</th>";
    t += "</tr></thead><tbody>";

    rows.slice(0, 1500).forEach(function (r) {
      t += "<tr>";
      t += "<td style='" + td + "white-space:nowrap;'>" + esc_(r.loja) + "</td>";
      t += "<td style='" + td + "white-space:nowrap;'>" + esc_(r.time) + "</td>";
      t += "<td style='" + td + "white-space:nowrap;'>" + esc_(r.data) + "</td>";
      t += "<td style='" + td + "'>" + esc_(r.estabelecimento) + "</td>";
      t += "<td style='" + td + "white-space:nowrap;text-align:right;'>" +
           esc_(Number(r.valor || 0).toLocaleString("pt-BR", { style: "currency", currency: "BRL" })) + "</td>";
      t += "<td style='" + td + "'>" + esc_(r.etiqueta) + "</td>";
      t += "<td style='" + td + "'>" + esc_(r.descricao) + "</td>";
      t += "</tr>";
    });

    t += "</tbody></table></div>";

    var mailObj = {
      to: emailDestino,
      subject: assunto,
      htmlBody: cab + t,
      name: "Vektor - Grupo SBF"
    };
    if (ccEmail) mailObj.cc = ccEmail;

    MailApp.sendEmail(mailObj);

    return { ok: true, to: emailDestino, cc: ccEmail || "" };
  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

function getResumoEtiquetasClara() {
  try {
    var info = carregarLinhasBaseClara_();
    if (info.error) return { totalGeral: 0, itens: [] };

    var header = info.header;
    var linhas = info.linhas;

    // Índice da coluna de etiquetas (coluna T: "Etiquetas")
    var idxEtiqueta = header.indexOf("Etiquetas");
    if (idxEtiqueta < 0) {
      idxEtiqueta = header.indexOf("Etiqueta");
    }
    if (idxEtiqueta < 0) {
      return { totalGeral: 0, itens: [] };
    }

    // Índice da coluna de VALOR da transação
    // Ajuste essa lista se o nome do cabeçalho for diferente
    var idxValor = encontrarIndiceColuna_(header, [
      "Valor original"
    ]);

    if (idxValor < 0) {
      // sem coluna de valor, não faz sentido calcular percentuais
      return { totalGeral: 0, itens: [] };
    }

    var mapa = {};       // etiqueta -> { valor: number, count: number }
    var totalGeral = 0;  // soma de todos os valores

    linhas.forEach(function (row) {
      var celEtiqueta = row[idxEtiqueta];
      if (celEtiqueta === null || celEtiqueta === undefined) return;

      var textoCelula = celEtiqueta.toString();
      if (!textoCelula) return;

      var rawValor = row[idxValor];
      var valorNum = 0;

      if (typeof rawValor === "number") {
        valorNum = rawValor;
      } else if (rawValor !== null && rawValor !== undefined) {
        var s = rawValor.toString().trim();
        if (s) {
          // tenta tratar "1.234,56" e "1234.56"
          s = s.replace(/\./g, "").replace(",", ".");
          var parsed = parseFloat(s);
          if (!isNaN(parsed)) valorNum = parsed;
        }
      }

      // Se ainda não conseguiu número, ignora esse valor
      if (isNaN(valorNum)) valorNum = 0;

      // Divide a célula em múltiplas etiquetas, separadas por "|"
      textoCelula.split("|").forEach(function (parte) {
        var et = parte.trim();
        if (!et) return;

        if (!mapa[et]) {
          mapa[et] = { valor: 0, count: 0 };
        }
        mapa[et].valor += valorNum;
        mapa[et].count += 1;
      });

      totalGeral += valorNum;
    });

    var itens = [];
    for (var etiqueta in mapa) {
      if (!Object.prototype.hasOwnProperty.call(mapa, etiqueta)) continue;
      var dado = mapa[etiqueta];
      var valorEti = dado.valor || 0;
      var perc = 0;

      if (totalGeral > 0) {
        perc = (valorEti / totalGeral) * 100;
      }

      itens.push({
        etiqueta: etiqueta,
        valorTotal: valorEti,
        percentual: perc,
        quantidade: dado.count
      });
    }

    // ordena por valor total decrescente
    itens.sort(function (a, b) {
      return b.valorTotal - a.valorTotal;
    });

    return {
      totalGeral: totalGeral,
      itens: itens
    };

  } catch (e) {
    return { totalGeral: 0, itens: [] };
  }
}

/**
 * Converte o texto da coluna "Extrato da conta"
 * (ex.: "06 Nov 2025 - 05 Dec 2025")
 * em datas de início/fim.
 */
function parseExtratoContaPeriodo_(texto) {
  if (!texto) return null;

  var m = texto.match(
    /(\d{1,2})\s+([A-Za-zÀ-ÿ]{3,})\s+(\d{4})\s*-\s*(\d{1,2})\s+([A-Za-zÀ-ÿ]{3,})\s+(\d{4})/
  );
  if (!m) return null;

  var dia1 = Number(m[1]);
  var mes1Str = m[2];
  var ano1 = Number(m[3]);

  var dia2 = Number(m[4]);
  var mes2Str = m[5];
  var ano2 = Number(m[6]);

  function mesFromStr(str) {
    var s = str.toLowerCase();
    s = s.normalize("NFD").replace(/[\u0300-\u036f]/g, "");

    var mapa = {
      jan: 0, janeiro: 0,
      feb: 1, fev: 1, fevereiro: 1,
      mar: 2, marco: 2, marcoo: 2,
      apr: 3, abr: 3, abril: 3,
      may: 4, mai: 4, maio: 4,
      jun: 5, junho: 5,
      jul: 6, julho: 6,
      aug: 7, ago: 7, agosto: 7,
      sep: 8, set: 8, setembro: 8,
      oct: 9, out: 9, outubro: 9,
      nov: 10, novembro: 10,
      dec: 11, dez: 11, dezembro: 11
    };

    // normaliza para 3 letras pra funcionar com "Nov", "Dec", "Dez"
    var chave3 = s.slice(0, 3);
    if (mapa.hasOwnProperty(chave3)) return mapa[chave3];

    if (mapa.hasOwnProperty(s)) return mapa[s];

    return null;
  }

  var mes1 = mesFromStr(mes1Str);
  var mes2 = mesFromStr(mes2Str);

  if (mes1 === null || mes2 === null) return null;

  return {
    inicio: new Date(ano1, mes1, dia1),
    fim:    new Date(ano2, mes2, dia2)
  };
}

/**
 * Lê a BaseClara, agrupa por "Extrato da conta" (coluna B)
 * e devolve a soma de valor por fatura.
 *
 * Retorno:
 * {
 *   ok: true,
 *   faturas: [
 *     {
 *       extrato: "06 Nov 2025 - 05 Dec 2025",
 *       valorTotal: 12345.67,
 *       dataInicioIso: "2025-11-06T03:00:00.000Z",
 *       dataFimIso:    "2025-12-05T03:00:00.000Z"
 *     },
 *     ...
 *   ]
 * }
 */
function getResumoFaturasClara() {
  vektorAssertFunctionAllowed_("getResumoFaturasClara");
  try {
    var info = carregarLinhasBaseClara_();
    if (info.error) {
      return { ok: false, error: info.error };
    }

    var header = info.header;
    var linhas = info.linhas;

    // Índice da coluna "Extrato da conta"
    var idxExtrato = encontrarIndiceColuna_(header, [
      "Extrato da conta",
      "Extrato conta",
      "Extrato"
    ]);

    // Índice da coluna de valor (mesmo critério que você já usa)
    var idxValor = encontrarIndiceColuna_(header, [
      "Valor em R$",
      "Valor (R$)",
      "Valor"
    ]);

    if (idxExtrato < 0 || idxValor < 0) {
      return {
        ok: false,
        error: "Não encontrei as colunas 'Extrato da conta' e 'Valor' na BaseClara."
      };
    }

    var mapa = {}; // chave = texto do extrato
    linhas.forEach(function (row) {
      var extrato = (row[idxExtrato] || "").toString().trim();
      if (!extrato) return;

      var valor = Number(row[idxValor]) || 0;
      if (!mapa[extrato]) {
        var periodo = parseExtratoContaPeriodo_(extrato);
        mapa[extrato] = {
          extrato: extrato,
          totalValor: 0,
          dataInicio: periodo ? periodo.inicio : null,
          dataFim:    periodo ? periodo.fim    : null
        };
      }
      mapa[extrato].totalValor += valor;
    });

    var faturas = [];
    for (var k in mapa) {
      if (!Object.prototype.hasOwnProperty.call(mapa, k)) continue;
      var f = mapa[k];
      faturas.push({
        extrato: k,
        valorTotal: f.totalValor,
        dataInicioIso: f.dataInicio ? f.dataInicio.toISOString() : "",
        dataFimIso:    f.dataFim    ? f.dataFim.toISOString()    : ""
      });
    }

    // Ordena por data de início (ou fim) crescente
    faturas.sort(function (a, b) {
      var da = a.dataInicioIso || a.dataFimIso || "";
      var db = b.dataInicioIso || b.dataFimIso || "";
      if (da && db) {
        if (da < db) return -1;
        if (da > db) return 1;
        return 0;
      }
      return a.extrato.localeCompare(b.extrato);
    });

    return {
      ok: true,
      faturas: faturas
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : e
    };
  }
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function getVektorMetricSheet_() {
  const ss = SpreadsheetApp.openById(VEKTOR_METRICAS_SHEET_ID);
  let sheet = ss.getSheetByName(VEKTOR_METRICAS_TAB_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(VEKTOR_METRICAS_TAB_NAME);
    // Cabeçalho padrão
    sheet.appendRow([
      'Timestamp',
      'UsuarioNome',
      'UsuarioEmail',
      'Loja',
      'Intencao',
      'Topico',
      'MensagemOriginal',
      'Norm',
      'Resultado',
      'FuncaoOrigem'
    ]);
  }

  return sheet;
}

function getOrCreateVertexCostSheet_() {
  const ss = SpreadsheetApp.openById(VEKTOR_METRICAS_SHEET_ID);
  let sh = ss.getSheetByName(VEKTOR_VERTEX_COST_TAB_NAME);

  if (!sh) {
    sh = ss.insertSheet(VEKTOR_VERTEX_COST_TAB_NAME);
    sh.appendRow([
      "Timestamp",
      "MonthKey",
      "UsuarioEmail",
      "ProjetoApi",
      "Model",
      "ModelVersion",
      "PromptTokens",
      "OutputTokens",
      "TotalTokens",
      "EstimatedUsd",
      "EstimatedBrl"
    ]);
    sh.getRange(1, 1, 1, 11).setFontWeight("bold");
    sh.setFrozenRows(1);
  }

  return sh;
}

function getVertexCostDashboard() {
  vektorAssertFunctionAllowed_("getVertexCostDashboard");

  var ctx = vektorGetUserRole_();
  var role = String((ctx && ctx.role) || "").trim();
  if (role !== "Administrador") {
    throw new Error("Não disponível para o seu perfil.");
  }

  var sh = getOrCreateVertexCostSheet_();
  var lr = sh.getLastRow();

  if (lr < 2) {
    return {
      ok: true,
      kpis: {
        users: 0,
        totalTokens: 0,
        totalUsd: 0,
        totalUsdFmt: "US$ 0",
        totalBrlFmt: "R$ 0"
      },
      rows: [],
      monthly: []
    };
  }

  var values = sh.getRange(2, 1, lr - 1, 11).getValues();

  var byUser = {};
  var byMonth = {};
  var totalTokensAll = 0;
  var totalUsdAll = 0;
  var totalBrlAll = 0;

  for (var i = 0; i < values.length; i++) {
    var r = values[i];

    var monthKey = String(r[1] || "").trim();
    var email = String(r[2] || "").trim().toLowerCase();
    var prompt = Number(r[6] || 0) || 0;
    var output = Number(r[7] || 0) || 0;
    var total = Number(r[8] || 0) || 0;
    var usd = Number(r[9] || 0) || 0;
    var brl = Number(r[10] || 0) || 0;

    if (!email) email = "não identificado";

    if (!byUser[email]) {
      byUser[email] = {
        userEmail: email,
        calls: 0,
        promptTokens: 0,
        outputTokens: 0,
        totalTokens: 0,
        totalUsd: 0,
        totalBrl: 0
      };
    }

    byUser[email].calls += 1;
    byUser[email].promptTokens += prompt;
    byUser[email].outputTokens += output;
    byUser[email].totalTokens += total;
    byUser[email].totalUsd += usd;
    byUser[email].totalBrl += brl;

    if (monthKey) {
      if (!byMonth[monthKey]) byMonth[monthKey] = { totalUsd: 0, totalBrl: 0 };
      byMonth[monthKey].totalUsd += usd;
      byMonth[monthKey].totalBrl += brl;
    }

    totalTokensAll += total;
    totalUsdAll += usd;
    totalBrlAll += brl;
  }

    var rows = Object.keys(byUser).map(function(email){
    var x = byUser[email];
    x.usdFmt = vektorFmtUsd_(x.totalUsd);
    x.brlFmt = Number(x.totalBrl || 0).toLocaleString("pt-BR", {
      style: "currency",
      currency: "BRL"
    });
    return x;
  });

  rows.sort(function(a,b){
    return (Number(b.totalUsd || 0) - Number(a.totalUsd || 0))
        || (Number(b.totalTokens || 0) - Number(a.totalTokens || 0))
        || String(a.userEmail || "").localeCompare(String(b.userEmail || ""), "pt-BR");
  });

  var monthly = Object.keys(byMonth).sort().map(function(k){
    var y = String(k).slice(0,4);
    var m = String(k).slice(4,6);
    return {
      monthKey: k,
      monthLabel: m + "/" + y,
      totalUsd: Number(byMonth[k].totalUsd || 0) || 0,
      totalBrl: Number(byMonth[k].totalBrl || 0) || 0
    };
  });

  return {
    ok: true,
    projectId: VEKTOR_VERTEX_PROJECT_ID,
    kpis: {
      users: rows.length,
      totalTokens: totalTokensAll,
      totalUsd: totalUsdAll,
      totalUsdFmt: vektorFmtUsd_(totalUsdAll),
      totalBrlFmt: vektorFmtBrlFromUsd_(totalUsdAll)
    },
    rows: rows,
    monthly: monthly
  };
}

/**
 * Recebe o Termo de Responsabilidade (arquivo em base64 + dados do usuário),
 * faz validações básicas e salva no formato original na pasta configurada.
 *
 * Nome final: "Aceite – Política e Termo de Responsabilidade Clara - NOME COMPLETO.ext"
 *
 * Após salvar, envia um e-mail para o Rodrigo com o arquivo em anexo
 * para conferência.
 *
 * @param {Object} payload
 * @return {Object} { ok: true, fileUrl: "..."} ou { ok: false, error: "..." }
 */
function salvarTermoResponsabilidade(payload) {
  vektorAssertFunctionAllowed_("salvarTermoResponsabilidade");
  try {
    if (!payload || !payload.base64) {
      throw new Error("Arquivo não recebido.");
    }

    // --- Validação de tipo MIME (robusta) ---
    var mimeType = (payload.mimeType || "").toLowerCase();

    var isPdf  = mimeType === "application/pdf";
    var isPng  = mimeType === "image/png";
    var isHeic = mimeType.indexOf("heic") !== -1 || mimeType.indexOf("heif") !== -1;
    var isJpeg = mimeType.indexOf("jpeg") !== -1 ||
                 mimeType.indexOf("jpg")  !== -1 ||
                 mimeType.indexOf("pjpeg")!== -1 ||
                 mimeType.indexOf("jfif") !== -1;

    if (!(isPdf || isPng || isHeic || isJpeg)) {
      throw new Error("Tipo de arquivo não permitido. Envie somente PDF, JPG, JPEG, PNG ou HEIC.");
    }

    // --- Verificação mínima se "parece" ser o Termo (pelo nome do arquivo) ---
    var fileNameOriginal = payload.fileNameOriginal || "arquivo_sem_nome";
    var nomeLower = fileNameOriginal.toLowerCase();

    if (!(nomeLower.indexOf("termo") !== -1 && nomeLower.indexOf("responsa") !== -1)) {
      throw new Error(
        "O arquivo não parece ser o Termo de Responsabilidade. " +
        "Renomeie o arquivo incluindo as palavras 'termo' e 'responsabilidade' e envie novamente."
      );
    }

    // --- Nome completo do usuário (já veio do chat) ---
    var nomeCompleto = payload.usuarioNome || "";
    if (!nomeCompleto) {
      throw new Error("Nome completo do usuário não informado.");
    }

    // Sanitiza o nome para não quebrar o nome do arquivo
    var nomeSanitizado = nomeCompleto.replace(/[\\/:*?\"<>|]/g, " ").trim();
    if (!nomeSanitizado) {
      nomeSanitizado = "Nome_indefinido";
    }

    // --- Define extensão de acordo com o tipo original ---
    var ext = "bin";
    if (isPdf)       ext = "pdf";
    else if (isPng)  ext = "png";
    else if (isHeic) ext = "heic";
    else if (isJpeg) ext = "jpg";

    var nomeFinal = "Aceite – Política e Termo de Responsabilidade Clara - " +
                    nomeSanitizado + "." + ext;

    // --- Decodifica base64 e monta o blob final NO FORMATO ORIGINAL ---
    var bytes     = Utilities.base64Decode(payload.base64);
    var blobFinal = Utilities.newBlob(
      bytes,
      payload.mimeType || "application/octet-stream",
      nomeFinal
    );

    // --- Salva na pasta do Drive configurada ---
    var pasta = DriveApp.getFolderById(VEKTOR_PASTA_TERMOS_ID);
    var file  = pasta.createFile(blobFinal);

    // --- Tenta enviar e-mail para conferência ---
    try {
      var assunto = "Validar - Termo enviado via Agent Vektor";

      var corpo =
        "Um novo Termo de Responsabilidade foi enviado via Agent Vektor.\n\n" +
        "Nome completo: " + nomeCompleto + "\n" +
        "E-mail do usuário: " + (payload.usuarioEmail || "") + "\n" +
        "Loja: " + (payload.loja || "") + "\n" +
        "Nome do arquivo salvo: " + nomeFinal + "\n\n" +
        "Link no Drive: " + file.getUrl() + "\n\n" +
        "Por favor, valide o conteúdo e o aceite desse termo.";

      GmailApp.sendEmail("contasareceber@gruposbf.com.br", assunto, corpo, {
        from: "vektor@gruposbf.com.br",
        name: "Vektor Grupo SBF",
        attachments: [file.getBlob()]
      });

    } catch (eMail) {
      // Não quebra o fluxo do usuário se o e-mail falhar; apenas loga
      console.error("Erro ao enviar e-mail de validação do Termo: " + eMail);
    }

    return {
      ok: true,
      fileId: file.getId(),
      fileUrl: file.getUrl()
    };

  } catch (e) {
    return {
      ok: false,
      error: e && e.message ? e.message : String(e)
    };
  }
}

function registrarMetricaVektor(payload) {
  // ✅ só exige que o usuário exista e esteja ATIVO (VEKTOR_EMAILS), além da whitelist
  vektorGetUserRole_(); // valida whitelist + VEKTOR_EMAILS (ATIVO)

  try {
    const sheet = getVektorMetricSheet_();
    const now = new Date();

    const linha = [
      now,
      payload.usuarioNome   || '',
      payload.usuarioEmail  || '',
      payload.loja          || '',
      payload.intencao      || '',
      payload.topico        || '',
      payload.mensagemOriginal || '',
      payload.norm          || '',
      payload.resultado     || '',
      payload.funcaoOrigem || ""   // coluna extra
    ];

    sheet.appendRow(linha);
  } catch (e) {
    console.error('Erro ao registrar métrica do Vektor: ' + e);
  }
}

function getLojasOfensorasParaChat(diasJanela) {
  vektorAssertFunctionAllowed_("getLojasOfensorasParaChat");

  // ✅ Janela agora é o ciclo atual (06 → hoje). "diasJanela" vira apenas o tamanho do ciclo até hoje.
  const tz = "America/Sao_Paulo";
  const ciclo = getPeriodoCicloClara_(); // já existe no projeto (inicio=dia 06 do ciclo, fim=hoje 23:59:59)
  const hoje = new Date();
  const inicioCiclo = ciclo.inicio;

  diasJanela = Math.max(
    1,
    Math.ceil((hoje.getTime() - inicioCiclo.getTime()) / (24 * 60 * 60 * 1000))
  );

  const rel = gerarRelatorioOfensorasPendencias_(diasJanela);
  if (!rel || !rel.ok) {
    return { ok: false, error: "Falha ao gerar relatório." };
  }

  // período (ciclo atual: 06 → hoje)
  var periodo = {
    inicio: Utilities.formatDate(inicioCiclo, tz, "dd/MM/yyyy"),
    fim: Utilities.formatDate(hoje, tz, "dd/MM/yyyy")
  };

  return {
    ok: true,
    periodo: periodo,
    meta: {
      janela: "Ciclo atual",
      diasJanela: diasJanela,
      totalLojas: (rel.rows || []).length
    },
    rows: (rel.rows || []).map(r => {
      const t14 = r.trend14 || {};

      // ✅ delta absoluto (compatível com versões antigas)
      const deltaAbs = (t14.deltaAbs != null) ? t14.deltaAbs : (t14.delta != null ? t14.delta : 0);

      return {
        loja: r.loja,
        time: r.time || "N/D",

        qtde: r.qtde,
        valor: r.valor,
        txCount: (r.txCount != null ? r.txCount : 0),
        diasComPendencia: r.diasComPendencia,
        pendEtiqueta: r.pendEtiqueta,
        pendNF: r.pendNF,
        pendDesc: r.pendDesc,

        // ✅ não force 0: se não existe, deixa null para o front/email mostrar "—"
        qtdeSnapshots: (r.qtdeSnapshots != null ? r.qtdeSnapshots : null),

        ult14: t14.ult14 || 0,
        ant14: t14.ant14 || 0,

        // ✅ opção C
        delta14: deltaAbs,
        delta14Pct: (t14.deltaPct != null ? t14.deltaPct : null),

        score: (r.score != null ? r.score : null),
        classificacao: r.classificacao || "—"
      };
    })
  };
}

// ===============================
// PENDENCIAS DO CICLO (BACKEND) — BASECLARA ONLY
// Fonte única: aba BaseClara
// Retorna transações pendentes detalhadas + agregações (cards/tabela/análise)
// ===============================
function getResumoCicloPendencias() {
  vektorAssertFunctionAllowed_("getResumoCicloPendencias");

  try {
    // ✅ ACL por Emails (mesmo padrão da Análise de Gastos)
    var ctx = vektorGetUserRole_();
    var email = (ctx && ctx.email)
      ? String(ctx.email).trim().toLowerCase()
      : String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();

    if (!email) throw new Error("Não foi possível identificar seu e-mail Google.");

    var allowedLojas = vektorGetAllowedLojasFromEmails_(email); // null => admin, array => lojas permitidas

    // ✅ cache curto (2 min) — evita “demorar pra carregar” em cliques/refresh
    var cache = CacheService.getScriptCache();
    var cacheSuffix = (allowedLojas === null) ? "ALL" : email; // admin = ALL, demais = email
    var cacheKey = "RC_BASECLARA_" + getCicloKey06a05_() + "_" + cacheSuffix;
    var cached = cache.get(cacheKey);
    if (cached) return JSON.parse(cached);

    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";

    // ✅ ciclo completo 06→05 (pra limitar o date-picker no front)
    var pc = getPeriodoCicloClaraCompleto_();
    var iniCiclo = pc.inicio;
    var fimCiclo = pc.fim;

    // mapa Loja -> Time
    var mapaTime = construirMapaLojaParaTime_();

    var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
    var sh = ss.getSheetByName("BaseClara");
    if (!sh) throw new Error("Aba BaseClara não encontrada.");

    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2) {
      var vazio = {
        ok: true,
        ciclo: formatPeriodoBR_(iniCiclo, fimCiclo),
        periodo: formatPeriodoBR_(iniCiclo, new Date()),
        meta: {},
        tx: [],
        aggs: {}
      };
      cache.put(cacheKey, JSON.stringify(vazio), 120);
      return vazio;
    }

    var all = sh.getRange(1, 1, lastRow, lastCol).getValues();
    var header = all[0].map(function(h){ return String(h || "").trim(); });
    var linhas = all.slice(1);

    function idxOf(possiveis) {
      for (var i = 0; i < possiveis.length; i++) {
        var ix = header.indexOf(possiveis[i]);
        if (ix >= 0) return ix;
      }
      return -1;
    }

    // ===== mapeamento colunas =====
    var idxDataTrans  = idxOf(["Data da Transação"]);
    var idxValorBRL   = idxOf(["Valor em R$", "Valor (R$)", "Valor"]);
    var idxLojaNum    = idxOf(["LojaNum", "Loja", "Código Loja", "cod_estbl", "cod_loja"]);
    var idxEstab      = idxOf(["Transação"]); // estabelecimento
    var idxEtiquetas  = idxOf(["Etiquetas"]);
    var idxRecibo     = idxOf(["Recibo"]);
    var idxDescricao  = idxOf(["Descrição"]);
    var idxCategoria  = idxOf(["Categoria da Compra"]);

    if (idxDataTrans < 0) throw new Error("Não encontrei a coluna de Data na BaseClara.");
    if (idxValorBRL  < 0) throw new Error("Não encontrei a coluna de Valor na BaseClara.");
    if (idxLojaNum   < 0) throw new Error("Não encontrei a coluna de Loja na BaseClara.");

    // ✅ Etiqueta/Descrição: regra atual (vazio/—/n/a/não/etc)
    function isVazioPend_(v) {
      if (v === null || v === undefined) return true;
      if (typeof v === "boolean") return (v === false);
      var s = String(v).trim().toLowerCase();
      if (!s) return true;
      if (s === "-" || s === "—" || s === "n/a") return true;
      if (s === "não" || s === "nao") return true;
      if (s === "false" || s === "0") return true;
      return false;
    }

    // ✅ Recibo: SOMENTE "Não" deve ser pendência
    function isPendRecibo_(v) {
      var s = String(v || "").trim().toLowerCase();
      return (s === "não" || s === "nao");
    }

    // ===== coletar transações pendentes =====
    var tx = [];

    for (var i = 0; i < linhas.length; i++) {
      var r = linhas[i];

      var d = parseDateClara_(r[idxDataTrans]);
      if (!d) continue;

      // limita dentro do ciclo 06→05 (completo)
      if (d < iniCiclo || d > fimCiclo) continue;

      var lojaRaw = String(r[idxLojaNum] || "").trim();
      var dig = lojaRaw.replace(/\D/g, "");
      if (!dig) continue;

      var lojaNum = String(Number(dig));
      var loja4 = lojaNum.padStart(4, "0");

      // ✅ aplica ACL por loja (Emails)
      if (Array.isArray(allowedLojas)) {
        if (allowedLojas.indexOf(lojaNum) < 0 && allowedLojas.indexOf(loja4) < 0) continue;
      }

      var timeDaLoja = (mapaTime && (mapaTime[loja4] || mapaTime[lojaNum])) ? (mapaTime[loja4] || mapaTime[lojaNum]) : "N/D";
      var valor = parseNumberSafe_(r[idxValorBRL]);

      var recibo = (idxRecibo >= 0 ? r[idxRecibo] : "");
      var etiqueta = (idxEtiquetas >= 0 ? r[idxEtiquetas] : "");
      var desc = (idxDescricao >= 0 ? r[idxDescricao] : "");

      var temPendRecibo = (idxRecibo >= 0) ? isPendRecibo_(recibo) : false;
      var temPendEtiqueta = isVazioPend_(etiqueta);
      var temPendDescricao = isVazioPend_(desc);

      if (!(temPendRecibo || temPendEtiqueta || temPendDescricao)) continue;

      var pendTxt = [
        temPendRecibo ? "NF/Recibo" : null,
        temPendEtiqueta ? "Etiqueta" : null,
        temPendDescricao ? "Descrição" : null
      ].filter(Boolean).join(" • ");

      tx.push({
        loja: lojaNum,
        loja4: loja4,
        time: String(timeDaLoja || "N/D"),
        dataIso: Utilities.formatDate(d, tz, "yyyy-MM-dd"),
        dataTxt: Utilities.formatDate(d, tz, "dd/MM/yyyy"),
        estabelecimento: (idxEstab >= 0 ? String(r[idxEstab] || "—") : "—"),
        valor: valor,
        valorTxt: valor.toLocaleString("pt-BR", { style: "currency", currency: "BRL" }),
        categoria: (idxCategoria >= 0 ? String(r[idxCategoria] || "—") : "—"),
        pendencias: pendTxt,
        pendNF: temPendRecibo ? 1 : 0,
        pendEtiqueta: temPendEtiqueta ? 1 : 0,
        pendDesc: temPendDescricao ? 1 : 0
      });
    }

    // ===== agregações =====
    var aggTime = {};
    var aggLoja = {};
    var aggCat  = {};

    var totalPend = 0;
    var totalValor = 0;
    var totalTx = tx.length;
    var pendNF = 0, pendDesc = 0, pendEtiqueta = 0;

    tx.forEach(function (t) {
      var qtdPendTx = (t.pendNF + t.pendEtiqueta + t.pendDesc);
      totalPend += qtdPendTx;
      totalValor += Number(t.valor || 0);

      pendNF += t.pendNF;
      pendDesc += t.pendDesc;
      pendEtiqueta += t.pendEtiqueta;

      var kT = t.time || "N/D";
      if (!aggTime[kT]) aggTime[kT] = { time: kT, pend: 0, valor: 0, tx: 0 };
      aggTime[kT].pend += qtdPendTx;
      aggTime[kT].valor += Number(t.valor || 0);
      aggTime[kT].tx += 1;

      var kL = t.loja4 || t.loja || "";
      if (!aggLoja[kL]) aggLoja[kL] = { loja: (t.loja || ""), loja4: (t.loja4 || ""), time: (t.time || "N/D"), pend: 0, valor: 0, tx: 0 };
      aggLoja[kL].pend += qtdPendTx;
      aggLoja[kL].valor += Number(t.valor || 0);
      aggLoja[kL].tx += 1;

      var kC = t.categoria || "—";
      if (!aggCat[kC]) aggCat[kC] = { categoria: kC, pend: 0, valor: 0, tx: 0 };
      aggCat[kC].pend += qtdPendTx;
      aggCat[kC].valor += Number(t.valor || 0);
      aggCat[kC].tx += 1;
    });

    var timesArr = Object.keys(aggTime).map(function(k){ return aggTime[k]; }).sort(function(a,b){ return (b.pend - a.pend) || (b.valor - a.valor); });
    var lojasArr = Object.keys(aggLoja).map(function(k){ return aggLoja[k]; }).sort(function(a,b){ return (b.pend - a.pend) || (b.valor - a.valor); });
    var catsArr  = Object.keys(aggCat).map(function(k){ return aggCat[k]; }).sort(function(a,b){ return (b.valor - a.valor) || (b.pend - a.pend); });

    var topTime = timesArr[0] || null;
    var topLoja = lojasArr[0] || null;

    var resp = {
      ok: true,
      ciclo: formatPeriodoBR_(iniCiclo, fimCiclo),
      periodo: formatPeriodoBR_(iniCiclo, new Date()),
      meta: {
        totalPendencias: totalPend,
        totalValor: totalValor,
        totalTransacoes: totalTx,
        pendNF: pendNF,
        pendDesc: pendDesc,
        pendEtiqueta: pendEtiqueta,
        topTime: topTime ? { time: topTime.time, pend: topTime.pend, valor: topTime.valor, tx: topTime.tx } : null,
        topLoja: topLoja ? { loja: (topLoja.loja || ""), loja4: (topLoja.loja4 || ""), time: (topLoja.time || "N/D"), pend: topLoja.pend, valor: topLoja.valor, tx: topLoja.tx } : null
      },
      tx: tx,
      aggs: {
        porTime: timesArr,
        porLoja: lojasArr,
        porCategoria: catsArr
      }
    };

    cache.put(cacheKey, JSON.stringify(resp), 120);
    return resp;

  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

// =====================================================
// PENDENCIAS DO CICLO — Detalhe por LOJA (modo "resumo")
// Regra: aceita Pendências (texto) OU Recibo/Etiqueta/Descrição
// =====================================================
function getResumoCicloPendenciasDetalheLojaResumo(loja) {
  vektorAssertFunctionAllowed_("getResumoCicloPendenciasDetalheLojaResumo");

  try {
    var lojaIn = String(loja || "").trim();
    if (!lojaIn) return { ok: false, error: "Loja inválida." };

    // normaliza loja (remove não-dígitos)
    var lojaDig = lojaIn.replace(/\D/g, "");
    if (!lojaDig) return { ok: false, error: "Loja inválida." };
    var lojaNum = String(Number(lojaDig));
    var loja4 = lojaNum.padStart(4, "0");

    // abre BaseClara
    var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
    var sh = ss.getSheetByName("BaseClara");
    if (!sh) return { ok: false, error: "Aba BaseClara não encontrada." };

    var lastRow = sh.getLastRow();
    if (lastRow < 2) return { ok: true, itens: [] };

    // lê header + dados
    var all = sh.getRange(1, 1, lastRow, sh.getLastColumn()).getValues();
    var header = all[0].map(function (h) { return String(h || "").trim(); });
    var linhas = all.slice(1);

    function idxOf(possiveis) {
      for (var i = 0; i < possiveis.length; i++) {
        var ix = header.indexOf(possiveis[i]);
        if (ix >= 0) return ix;
      }
      return -1;
    }

    // índices
    var idxDataTrans = idxOf(["Data da Transação", "Data Transação", "Data"]);
    var idxValorBRL  = idxOf(["Valor em R$", "Valor (R$)", "Valor"]);
    var idxLojaNum   = idxOf(["LojaNum", "Loja", "Código Loja", "cod_estbl", "cod_loja"]);

    var idxRecibo    = idxOf(["Recibo", "NF/Recibo"]);
    var idxEtiquetas = idxOf(["Etiquetas", "Etiqueta"]);
    var idxDescricao = idxOf(["Descrição", "Descricao"]);
    var idxCategoria = idxOf(["Categoria da Compra", "Categoria", "Categoria Compra", "Categoria da Compra (Clara)"]);

    // texto agregado de pendências (quando existir)
    var idxPendTxt   = idxOf(["Pendências", "Pendencias", "Pendência", "Pendencia"]);

    if (idxDataTrans < 0) throw new Error("Não encontrei a coluna de Data na BaseClara.");
    if (idxValorBRL  < 0) throw new Error("Não encontrei a coluna de Valor na BaseClara.");
    if (idxLojaNum   < 0) throw new Error("Não encontrei a coluna de Loja na BaseClara.");

    // período do ciclo atual
    var pc = getPeriodoCicloClara_();
    var ini = pc && pc.inicio ? pc.inicio : null;
    var fim = new Date();
    if (!ini) throw new Error("Não consegui identificar o início do ciclo atual.");

    function parseDateClara_(v) {
      if (v instanceof Date) return v;
      var s = String(v || "").trim();
      if (!s) return null;
      // tenta dd/MM/yyyy
      var m = s.match(/^(\d{2})\/(\d{2})\/(\d{4})/);
      if (m) return new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
      var d = new Date(s);
      return isNaN(d.getTime()) ? null : d;
    }

    // ✅ mantém para Etiquetas/Descrição (pendência costuma ser vazio/"Não"/etc)
    function isVazioPend_(v) {
      if (v === null || v === undefined) return true;
      if (typeof v === "boolean") return (v === false);
      var s = String(v).trim().toLowerCase();
      if (!s) return true;
      if (s === "-" || s === "—" || s === "n/a") return true;
      if (s === "não" || s === "nao") return true;
      if (s === "false" || s === "0") return true;
      return false;
    }

    // ✅ NOVO: Recibo só é pendência quando for explicitamente "Não"
    function isPendRecibo_(v) {
      var s = String(v || "").trim().toLowerCase();
      return (s === "não" || s === "nao");
    }

    // monta itens
    var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
    var itens = [];

    for (var i = 0; i < linhas.length; i++) {
      var r = linhas[i];

      // loja da linha
      var lojaRaw = String(r[idxLojaNum] || "").trim();
      var dig = lojaRaw.replace(/\D/g, "");
      if (!dig) continue;
      var ln = String(Number(dig));
      var ln4 = ln.padStart(4, "0");

      // filtra loja alvo
      if (!(ln === lojaNum || ln4 === loja4)) continue;

      // data no ciclo
      var d = parseDateClara_(r[idxDataTrans]);
      if (!d) continue;
      if (d < ini || d > fim) continue;

      // regras de pendência
      // ✅ Recibo: SOMENTE "Não"
      var temPendRecibo = (idxRecibo >= 0) ? isPendRecibo_(r[idxRecibo]) : false;

      // ✅ Etiquetas/Descrição: mantém tua heurística atual
      var temPendEtiqueta  = isVazioPend_(idxEtiquetas >= 0 ? r[idxEtiquetas] : "");
      var temPendDescricao = isVazioPend_(idxDescricao >= 0 ? r[idxDescricao] : "");

      // pendência por texto agregado (se coluna existir)
      var pendTxt = (idxPendTxt >= 0) ? String(r[idxPendTxt] || "").trim() : "";
      var temPendTxt = !!pendTxt;

      // vale se tiver texto OU uma das 3 pendências clássicas
      if (!(temPendTxt || temPendRecibo || temPendEtiqueta || temPendDescricao)) continue;

      var v = Number(r[idxValorBRL] || 0);
      var categoria = (idxCategoria >= 0 ? (r[idxCategoria] || "—") : "—");

      var pendenciasTxt = temPendTxt
        ? pendTxt
        : [
            temPendRecibo ? "NF/Recibo" : null,
            temPendEtiqueta ? "Etiqueta" : null,
            temPendDescricao ? "Descrição" : null
          ].filter(Boolean).join(" • ");

      itens.push({
        loja: ln,
        data: Utilities.formatDate(d, tz, "dd/MM/yyyy"),
        valor: v,
        valorFmt: v.toLocaleString("pt-BR", { style: "currency", currency: "BRL" }),
        categoria: String(categoria || "—"),
        pendenciasTxt: pendenciasTxt || "—"
      });
    }

    return { ok: true, itens: itens };

  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

function getComparativoFaturasClaraCore_(extratoAtualOpt, extratoAnteriorOpt) {
  try {
    var info = carregarLinhasBaseClara_();
    if (info.error) return { ok: false, error: info.error };

    var header = info.header || [];
    var linhas = info.linhas || [];

    function findHeaderExact_(headerArr, label) {
      var alvo = String(label || "").trim().toLowerCase();
      for (var i = 0; i < headerArr.length; i++) {
        var h = String(headerArr[i] || "").trim().toLowerCase();
        if (h === alvo) return i;
      }
      return -1;
    }

    var idxExtrato   = encontrarIndiceColuna_(header, ["Extrato da conta", "Extrato conta", "Extrato"]);
    var idxValor     = encontrarIndiceColuna_(header, ["Valor em R$", "Valor (R$)", "Valor"]);
    var idxLojaNum   = encontrarIndiceColuna_(header, ["LojaNum", "Loja Num", "Loja", "Cod Loja", "Código Loja"]);
    var idxTime      = encontrarIndiceColuna_(header, ["Grupos", "Grupo", "Time"]);
    var idxCategoria = encontrarIndiceColuna_(header, ["Categoria", "Etiqueta", "Tipo de gasto", "Tag"]);

    var idxEstab = findHeaderExact_(header, "Transação");
    var idxData  = findHeaderExact_(header, "Data da transação");

    if (idxExtrato < 0 || idxValor < 0 || idxLojaNum < 0) {
      return {
        ok: false,
        error: "Não encontrei as colunas mínimas ('Extrato da conta', 'Valor', 'LojaNum') na BaseClara."
      };
    }

    // ===== Descobre todos os extratos existentes
    var mapaExtratos = {};
    for (var i = 0; i < linhas.length; i++) {
      var ex = (linhas[i][idxExtrato] || "").toString().trim();
      if (!ex) continue;

      if (!mapaExtratos[ex]) {
        var p = parseExtratoContaPeriodo_(ex);
        mapaExtratos[ex] = {
          extrato: ex,
          inicio: p ? p.inicio : null,
          fim: p ? p.fim : null
        };
      }
    }

    var extratos = Object.keys(mapaExtratos)
      .map(function(k){ return mapaExtratos[k]; })
      .filter(function(x){ return x && x.extrato; });

    extratos.sort(function(a, b){
      var da = a.inicio ? a.inicio.getTime() : (a.fim ? a.fim.getTime() : 0);
      var db = b.inicio ? b.inicio.getTime() : (b.fim ? b.fim.getTime() : 0);
      return da - db;
    });

    if (extratos.length < 2) {
      return {
        ok: false,
        error: "Não há faturas suficientes para comparação (preciso de pelo menos 2 extratos)."
      };
    }

    var extratoAtual = String(extratoAtualOpt || "").trim();
    var extratoAnterior = String(extratoAnteriorOpt || "").trim();

    var fatAtual = null;
    var fatAnterior = null;

    if (!extratoAtual && !extratoAnterior) {
      fatAnterior = extratos[extratos.length - 2];
      fatAtual    = extratos[extratos.length - 1];
      extratoAnterior = fatAnterior.extrato;
      extratoAtual    = fatAtual.extrato;
    } else {
      if (!extratoAtual || !extratoAnterior) {
        return {
          ok: false,
          error: "Informe as duas faturas para a análise temporal."
        };
      }

      if (extratoAtual === extratoAnterior) {
        return {
          ok: false,
          error: "A fatura base e a fatura de comparação não podem ser iguais."
        };
      }

      fatAtual = mapaExtratos[extratoAtual] || null;
      fatAnterior = mapaExtratos[extratoAnterior] || null;

      if (!fatAtual || !fatAnterior) {
        return {
          ok: false,
          error: "Não encontrei uma ou ambas as faturas selecionadas na BaseClara."
        };
      }
    }

    var tz = "America/Sao_Paulo";

    var periodo = {
      anterior: {
        extrato: extratoAnterior,
        inicio: fatAnterior.inicio ? Utilities.formatDate(fatAnterior.inicio, tz, "dd/MM/yyyy") : "",
        fim:    fatAnterior.fim    ? Utilities.formatDate(fatAnterior.fim, tz, "dd/MM/yyyy") : ""
      },
      atual: {
        extrato: extratoAtual,
        inicio: fatAtual.inicio ? Utilities.formatDate(fatAtual.inicio, tz, "dd/MM/yyyy") : "",
        fim:    fatAtual.fim    ? Utilities.formatDate(fatAtual.fim, tz, "dd/MM/yyyy") : ""
      }
    };

    var hoje = new Date();
    hoje = new Date(Utilities.formatDate(hoje, tz, "yyyy/MM/dd") + " 00:00:00");

    var inicioAtual    = fatAtual.inicio;
    var fimAtual       = fatAtual.fim;
    var inicioAnterior = fatAnterior.inicio;
    var fimAnterior    = fatAnterior.fim;

    var usarRecorte = !!(inicioAtual && fimAtual && inicioAnterior && fimAnterior && idxData >= 0);

    var fimRecorteAtual = null;
    var fimRecorteAnterior = null;
    var eventosSazonais = [];

    if (usarRecorte) {
      fimRecorteAtual = (hoje.getTime() < fimAtual.getTime()) ? hoje : fimAtual;

      function addDays_(d, n) {
        return new Date(d.getTime() + n * 24 * 60 * 60 * 1000);
      }

      function lastFridayOfNovember_(year) {
        var d = new Date(year, 10, 30, 0, 0, 0, 0);
        while (d.getDay() !== 5) d = new Date(d.getTime() - 24*60*60*1000);
        return d;
      }

      function secondSunday_(year, monthIndex0) {
        var d = new Date(year, monthIndex0, 1, 0, 0, 0, 0);
        while (d.getDay() !== 0) d = new Date(d.getTime() + 24*60*60*1000);
        return new Date(d.getTime() + 7*24*60*60*1000);
      }

      function detectRetailEvents_(startDate, endDate) {
        var events = [];
        var y = startDate.getFullYear();
        var y2 = endDate.getFullYear();

        for (var year = y; year <= y2; year++) {
          var bf = lastFridayOfNovember_(year);
          events.push({ nome: "Black Friday", start: addDays_(bf, -3), end: addDays_(bf, 3) });
          events.push({ nome: "Natal", start: new Date(year, 11, 20), end: new Date(year, 11, 26) });
          events.push({ nome: "Ano Novo", start: new Date(year, 11, 28), end: new Date(year + 1, 0, 2) });

          var maes = secondSunday_(year, 4);
          events.push({ nome: "Dia das Mães", start: addDays_(maes, -3), end: addDays_(maes, 3) });

          var pais = secondSunday_(year, 7);
          events.push({ nome: "Dia dos Pais", start: addDays_(pais, -3), end: addDays_(pais, 3) });

          events.push({ nome: "Dia das Crianças", start: new Date(year, 9, 9), end: new Date(year, 9, 15) });
          events.push({ nome: "Dia dos Namorados", start: new Date(year, 5, 9), end: new Date(year, 5, 15) });
        }

        var hit = [];
        for (var i = 0; i < events.length; i++) {
          var e = events[i];
          var intersects = !(e.end.getTime() < startDate.getTime() || e.start.getTime() > endDate.getTime());
          if (intersects) hit.push(e.nome);
        }

        var seen = {};
        return hit.filter(function(n){
          if (seen[n]) return false;
          seen[n] = true;
          return true;
        });
      }

      var recorteAtualInicio = inicioAtual;
      var recorteAtualFim = fimRecorteAtual;
      if (recorteAtualInicio && recorteAtualFim) {
        eventosSazonais = detectRetailEvents_(recorteAtualInicio, recorteAtualFim);
      }

      var msDia = 24 * 60 * 60 * 1000;
      var diasDecorridos = Math.floor((fimRecorteAtual.getTime() - inicioAtual.getTime()) / msDia);

      fimRecorteAnterior = new Date(inicioAnterior.getTime() + diasDecorridos * msDia);
      if (fimRecorteAnterior.getTime() > fimAnterior.getTime()) fimRecorteAnterior = fimAnterior;

      periodo.atual.fim = Utilities.formatDate(fimRecorteAtual, tz, "dd/MM/yyyy");
      periodo.anterior.fim = Utilities.formatDate(fimRecorteAnterior, tz, "dd/MM/yyyy");
    }

    var mapaTime = construirMapaLojaParaTime_();

    var stats = {};
    var dayPrevGeral = {};
    var dayCurGeral  = {};
    var ultimaDataConsiderada = null;

    function lojaKey(v) {
      var n = normalizarLojaNumero_(v);
      return n ? String(n) : String(v || "").trim() || "(N/D)";
    }
    function valNum(v){ return Number(v) || 0; }
    function str(v){ return (v == null ? "" : String(v)).trim(); }
    function dayKey(dt){
      if (!(dt instanceof Date) || isNaN(dt.getTime())) return "";
      return Utilities.formatDate(dt, tz, "dd/MM/yyyy");
    }

    for (var r = 0; r < linhas.length; r++) {
      var row = linhas[r];
      var ex2 = (row[idxExtrato] || "").toString().trim();
      if (ex2 !== extratoAtual && ex2 !== extratoAnterior) continue;

      var dtRow = null;
      if (usarRecorte) {
        dtRow = row[idxData] instanceof Date ? row[idxData] : new Date(row[idxData]);
        if (!(dtRow instanceof Date) || isNaN(dtRow.getTime())) continue;

        if (ex2 === extratoAtual) {
          if (!ultimaDataConsiderada || dtRow.getTime() > ultimaDataConsiderada.getTime()) {
            ultimaDataConsiderada = dtRow;
          }
        }

        dtRow = new Date(Utilities.formatDate(dtRow, tz, "yyyy/MM/dd") + " 00:00:00");

        if (ex2 === extratoAtual) {
          if (dtRow.getTime() < inicioAtual.getTime() || dtRow.getTime() > fimRecorteAtual.getTime()) continue;
        } else {
          if (dtRow.getTime() < inicioAnterior.getTime() || dtRow.getTime() > fimRecorteAnterior.getTime()) continue;
        }
      }

      var loja = lojaKey(row[idxLojaNum]);
      if (!stats[loja]) {
        stats[loja] = {
          loja: loja,
          time: "",
          prev: 0,
          cur: 0,
          catPrev: {},
          catCur: {},
          estabPrev: {},
          estabCur: {},
          dayPrev: {},
          dayCur: {}
        };
      }

      var st = stats[loja];

      var timeLinha = (idxTime >= 0) ? str(row[idxTime]) : "";
      if (!st.time) st.time = timeLinha || (mapaTime[loja] || "N/D");

      var valor = valNum(row[idxValor]);
      var cat = (idxCategoria >= 0 ? str(row[idxCategoria]) : "") || "Sem categoria";

      var estabRaw = (idxEstab >= 0 ? row[idxEstab] : "");
      var estab = "Sem estabelecimento";
      if (idxEstab >= 0) {
        if (estabRaw instanceof Date && !isNaN(estabRaw.getTime())) {
          estab = Utilities.formatDate(estabRaw, tz, "dd/MM/yyyy");
        } else {
          estab = str(estabRaw) || "Sem estabelecimento";
        }
      }

      var dk = "";
      if (dtRow && dtRow instanceof Date && !isNaN(dtRow.getTime())) {
        dk = dayKey(dtRow);
      } else if (idxData >= 0) {
        var dtTry = row[idxData] instanceof Date ? row[idxData] : new Date(row[idxData]);
        dk = dayKey(dtTry);
      }

      if (ex2 === extratoAnterior) {
        st.prev += valor;
        st.catPrev[cat] = (st.catPrev[cat] || 0) + valor;
        st.estabPrev[estab] = (st.estabPrev[estab] || 0) + valor;

        if (dk) {
          st.dayPrev[dk] = (st.dayPrev[dk] || 0) + valor;
          dayPrevGeral[dk] = (dayPrevGeral[dk] || 0) + valor;
        }
      } else {
        st.cur += valor;
        st.catCur[cat] = (st.catCur[cat] || 0) + valor;
        st.estabCur[estab] = (st.estabCur[estab] || 0) + valor;

        if (dk) {
          st.dayCur[dk] = (st.dayCur[dk] || 0) + valor;
          dayCurGeral[dk] = (dayCurGeral[dk] || 0) + valor;
        }
      }
    }

    var rows = [];
    var lojas = Object.keys(stats);

    function pickDriverCategory(st, delta){
      var cats = {};
      Object.keys(st.catPrev || {}).forEach(function(c){ cats[c] = (cats[c] || 0) - st.catPrev[c]; });
      Object.keys(st.catCur  || {}).forEach(function(c){ cats[c] = (cats[c] || 0) + st.catCur[c]; });

      var bestCat = null;
      var bestDelta = (delta >= 0 ? -1e18 : 1e18);

      Object.keys(cats).forEach(function(c){
        var d = cats[c] || 0;
        if (delta >= 0) {
          if (d > bestDelta) { bestDelta = d; bestCat = c; }
        } else {
          if (d < bestDelta) { bestDelta = d; bestCat = c; }
        }
      });

      if (!bestCat) return { cat: "Sem categoria", deltaCat: 0 };
      return { cat: bestCat, deltaCat: bestDelta };
    }

    function pickPeakDay(st){
      var bestDay = "";
      var bestVal = 0;
      Object.keys(st.dayCur || {}).forEach(function(d){
        var v = st.dayCur[d] || 0;
        if (v > bestVal) { bestVal = v; bestDay = d; }
      });
      return { day: bestDay || "—", value: bestVal || 0 };
    }

    function pickEstabCondicional(st, deltaAbs){
      var deltas = {};
      Object.keys(st.estabPrev || {}).forEach(function(e){ deltas[e] = (deltas[e] || 0) - st.estabPrev[e]; });
      Object.keys(st.estabCur  || {}).forEach(function(e){ deltas[e] = (deltas[e] || 0) + st.estabCur[e]; });

      var best = null;
      var bestD = (deltaAbs >= 0 ? -1e18 : 1e18);

      Object.keys(deltas).forEach(function(e){
        if (!(st.estabPrev[e] > 0)) return;

        var d = deltas[e] || 0;
        if (deltaAbs >= 0) {
          if (d > bestD) { bestD = d; best = e; }
        } else {
          if (d < bestD) { bestD = d; best = e; }
        }
      });

      if (!best) return null;

      var share = (Math.abs(deltaAbs) > 0) ? (Math.abs(bestD) / Math.abs(deltaAbs)) : 0;
      if (share < 0.30) return null;

      return { estab: best, deltaEstab: bestD, share: share };
    }

    lojas.forEach(function(k){
      var st = stats[k];
      var prev = st.prev || 0;
      var cur  = st.cur || 0;
      var delta = cur - prev;

      var varPct = null;
      var varPctTxt = "";
      if (prev > 0) {
        varPct = (delta / prev) * 100;
        varPctTxt = (varPct > 0 ? "+" : "") + varPct.toFixed(1) + "%";
      } else {
        varPctTxt = (cur > 0 ? "Início no período" : "—");
      }

      var dCat = pickDriverCategory(st, delta);
      var pico = pickPeakDay(st);
      var estabInfo = pickEstabCondicional(st, delta);

      var fator = "";
      if (cur === 0 && prev === 0) {
        fator = "Sem gastos nos dois períodos.";
      } else if (prev === 0 && cur > 0) {
        fator = "Início de gasto no período atual; Categoria: " + dCat.cat +
                " (Δ R$ " + (delta >= 0 ? "+" : "") + delta.toFixed(2) + "). Pico em " + pico.day + ".";
      } else if (delta === 0) {
        fator = "Sem variação relevante entre os períodos.";
      } else {
        var catPart = "Categoria: " + dCat.cat +
                      " (Δ R$ " + (dCat.deltaCat >= 0 ? "+" : "") + dCat.deltaCat.toFixed(2) + ")";
        var picoPart = "Pico em " + pico.day;
        var estabPart = "";
        if (estabInfo) {
          estabPart = "; Estab: " + estabInfo.estab +
                      " (Δ R$ " + (estabInfo.deltaEstab >= 0 ? "+" : "") + estabInfo.deltaEstab.toFixed(2) +
                      ", " + Math.round(estabInfo.share * 100) + "% do Δ)";
        }
        fator = (delta > 0 ? "Aumento puxado por " : "Queda puxada por ") + catPart + estabPart + ". " + picoPart + ".";
      }

      rows.push({
        loja: st.loja,
        time: st.time || "N/D",
        valorAnterior: prev,
        valorAtual: cur,
        deltaValor: delta,
        variacaoPctTxt: varPctTxt,
        variacaoPctNum: varPct,
        categoriaDriver: dCat.cat,
        picoDia: pico.day,
        picoValor: pico.value,
        fatorVariacao: fator
      });
    });

    rows.sort(function(a, b){
      return (b.deltaValor || 0) - (a.deltaValor || 0);
    });

    var totalPrev = 0, totalCur = 0;
    rows.forEach(function(r){
      totalPrev += Number(r.valorAnterior) || 0;
      totalCur  += Number(r.valorAtual) || 0;
    });

    var totalDelta = totalCur - totalPrev;
    var totalVarPctTxt = (totalPrev > 0)
      ? (((totalDelta / totalPrev * 100) > 0 ? "+" : "") + (totalDelta / totalPrev * 100).toFixed(1) + "%")
      : (totalCur > 0 ? "Início no período" : "—");

    var deltaCatGeral = {};
    Object.keys(stats).forEach(function(loja){
      var st = stats[loja];
      Object.keys(st.catPrev || {}).forEach(function(c){ deltaCatGeral[c] = (deltaCatGeral[c] || 0) - st.catPrev[c]; });
      Object.keys(st.catCur  || {}).forEach(function(c){ deltaCatGeral[c] = (deltaCatGeral[c] || 0) + st.catCur[c]; });
    });

    var topCats = Object.keys(deltaCatGeral).map(function(c){
      return { categoria: c, delta: deltaCatGeral[c] || 0 };
    }).sort(function(a, b){
      return Math.abs(b.delta) - Math.abs(a.delta);
    }).slice(0, 3);

    var topLojas = rows.slice().sort(function(a, b){
      return Math.abs(b.deltaValor || 0) - Math.abs(a.deltaValor || 0);
    }).slice(0, 5);

    var deltaTimeGeral = {};
    rows.forEach(function(r){
      var t = String(r.time || "N/D").trim() || "N/D";
      deltaTimeGeral[t] = (deltaTimeGeral[t] || 0) + (Number(r.deltaValor) || 0);
    });

    var topTimes = Object.keys(deltaTimeGeral).map(function(t){
      return { time: t, delta: deltaTimeGeral[t] || 0 };
    }).sort(function(a, b){
      return Math.abs(b.delta) - Math.abs(a.delta);
    }).slice(0, 5);

    var diasSet = {};
    Object.keys(dayPrevGeral).forEach(function(d){ diasSet[d] = true; });
    Object.keys(dayCurGeral).forEach(function(d){ diasSet[d] = true; });

    var topDias = Object.keys(diasSet).map(function(d){
      var vPrev = dayPrevGeral[d] || 0;
      var vCur  = dayCurGeral[d] || 0;
      return { dia: d, prev: vPrev, cur: vCur, delta: (vCur - vPrev) };
    }).sort(function(a, b){
      return Math.abs(b.delta) - Math.abs(a.delta);
    }).slice(0, 5);

    var insights = rows.filter(function(r){
      return (r.deltaValor || 0) > 0;
    }).slice(0, 5).map(function(r){
      return {
        loja: r.loja,
        time: r.time,
        deltaValor: r.deltaValor,
        categoriaDriver: r.categoriaDriver,
        picoDia: r.picoDia,
        fatorVariacao: r.fatorVariacao
      };
    });

    var ultimaDataConsideradaTxt = ultimaDataConsiderada
      ? Utilities.formatDate(ultimaDataConsiderada, tz, "dd/MM/yyyy")
      : (usarRecorte && fimRecorteAtual ? Utilities.formatDate(fimRecorteAtual, tz, "dd/MM/yyyy") : "");

    return {
      ok: true,
      periodo: periodo,
      meta: {
        extratoAtual: extratoAtual,
        extratoAnterior: extratoAnterior,
        totalLojas: rows.length,
        ultimaDataConsiderada: ultimaDataConsideradaTxt
      },
      insights: insights,
      summary: {
        totalPrev: totalPrev,
        totalCur: totalCur,
        totalDelta: totalDelta,
        totalVarPctTxt: totalVarPctTxt,
        topCats: topCats,
        topLojas: topLojas,
        topTimes: topTimes,
        topDias: topDias,
        eventosSazonais: eventosSazonais,
        sazonalidadeTexto: (eventosSazonais && eventosSazonais.length)
          ? ("Observação sazonal: o recorte atual coincide com " + eventosSazonais.join(", ") + ", o que pode explicar parte da variação em relação ao período anterior.")
          : ""
      },
      rows: rows
    };

  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

function getComparativoFaturasClaraParaChat() {
  vektorAssertFunctionAllowed_("getComparativoFaturasClaraParaChat");
  return getComparativoFaturasClaraCore_("", "");
}

function getAnaliseTemporalFaturasVektor(extratoAtual, extratoAnterior) {
  vektorAssertFunctionAllowed_("getAnaliseTemporalFaturasVektor");
  return getComparativoFaturasClaraCore_(extratoAtual, extratoAnterior);
}

function DISPARAR_EMAIL_OFENSORAS_SEMANA() {
  var props = PropertiesService.getScriptProperties();

  // assinatura atual da BaseClara
  var sigAtual = calcularAssinaturaBaseClara_();
  if (!sigAtual || sigAtual.error) {
    Logger.log("Falha ao calcular assinatura BaseClara");
    return;
  }

  var KEY_ULT_ENVIO = "VEKTOR_OFENSORAS_SIG_ULT_ENVIO";
  var sigUltEnvio = props.getProperty(KEY_ULT_ENVIO) || "";

  // Se BaseClara NÃO mudou desde o último envio semanal → não envia
  if (sigAtual.sig === sigUltEnvio) {
    Logger.log("BaseClara não mudou desde o último e-mail semanal. Não envia.");
    return;
  }

  // Envia o e-mail (admins)
  var res = enviarEmailOfensorasPendenciasClara(0);
  if (!res || !res.ok) {
    Logger.log("Falha ao enviar e-mail de ofensoras");
    return;
  }

  // Marca assinatura como já enviada
  props.setProperty(KEY_ULT_ENVIO, sigAtual.sig);

  Logger.log("E-mail semanal de lojas ofensoras enviado com sucesso.");
}

function DISPARAR_EMAIL_ITENS_IRREGULARES_SEMANA() {
  var props = PropertiesService.getScriptProperties();

  // assinatura atual da BaseClara
  var sigAtual = calcularAssinaturaBaseClara_();
  if (!sigAtual || sigAtual.error) {
    Logger.log("Falha ao calcular assinatura BaseClara (itens irregulares)");
    return;
  }

  var KEY_ULT_ENVIO = "VEKTOR_ITENS_IRREG_SIG_ULT_ENVIO";
  var sigUltEnvio = props.getProperty(KEY_ULT_ENVIO) || "";

  // Se BaseClara NÃO mudou desde o último envio semanal → não envia
  if (sigAtual.sig === sigUltEnvio) {
    Logger.log("BaseClara não mudou desde o último e-mail semanal (itens irregulares). Não envia.");
    return;
  }

  var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
  var hoje = new Date();
  var fim = new Date(hoje); fim.setHours(23,59,59,999);
  var ini = new Date(hoje); ini.setDate(ini.getDate() - 6); ini.setHours(0,0,0,0);

  // usa o mesmo motor de conformidade do chat e filtra ALERTA
  var res = getListaItensCompradosClara("", "", ini, fim, 2500);
  if (!res || !res.ok) {
    Logger.log("Falha ao listar itens comprados para e-mail semanal: " + (res && res.error ? res.error : ""));
    return;
  }

  var rows = Array.isArray(res.rows) ? res.rows : [];
  rows = rows.filter(function(r){
    return String(r.conformidade || r.status || "").toUpperCase() === "ALERTA";
  });

  // limita para não explodir email/quota
  if (rows.length > 500) rows = rows.slice(0, 500);

  // se não tem alerta, ainda registra o gate para evitar spam repetido
  var admins = vektorGetAdminEmails_();
  var to = (admins && admins.join) ? admins.join(",") : "";
  if (!to) {
    Logger.log("Sem e-mails de admin para envio (itens irregulares).");
    return;
  }

  var periodoTxt = Utilities.formatDate(ini, tz, "dd/MM/yyyy") + " → " + Utilities.formatDate(fim, tz, "dd/MM/yyyy");
  var html = buildEmailItensIrregulares_(rows, periodoTxt);

  GmailApp.sendEmail(to, "Vektor — Possíveis itens irregulares (ALERTA) — " + periodoTxt, " ", {
      from: "vektor@gruposbf.com.br",
      htmlBody: html
    });

  // ✅ registra no log para aparecer no modal “Disparo de Ocorrências”
  try {
    registrarAlertaEnviado_(
      "ITENS_IRREGULARES",
      "", // loja (agregado)
      "", // time (agregado)
      "Semanal | ALERTA | período " + periodoTxt + " | linhas=" + rows.length,
      to,
      "AUTO_SEMANAL"
    );
  } catch (eLog) {}

  // Marca assinatura como já enviada
  props.setProperty(KEY_ULT_ENVIO, sigAtual.sig);

  Logger.log("E-mail semanal de itens irregulares enviado com sucesso.");
}

function RESETAR_GATE_EMAIL_ITENS_IRREGULARES_SEMANA() {
  PropertiesService.getScriptProperties().deleteProperty("VEKTOR_ITENS_IRREG_SIG_ULT_ENVIO");
  Logger.log("Gate resetado: VEKTOR_ITENS_IRREG_SIG_ULT_ENVIO removida. Próximo disparo enviará novamente.");
}

function buildEmailItensIrregulares_(rows, periodoTxt) {
  rows = Array.isArray(rows) ? rows : [];
  var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";

  function badgeHtml_(c) {
  var x = String(c || "").toUpperCase();
  var bg = "rgba(255,255,255,0.06)", bd = "rgba(148,163,184,0.25)", tx = "rgba(226,232,240,0.95)";
  if (x === "OK") { bg = "rgba(34,197,94,0.18)"; bd = "rgba(34,197,94,0.35)"; tx = "#14532d"; }
  if (x === "REVISAR") { bg = "rgba(245,158,11,0.20)"; bd = "rgba(245,158,11,0.40)"; tx = "#713f12"; }
  if (x === "ALERTA") { bg = "rgba(248,113,113,0.20)"; bd = "rgba(248,113,113,0.40)"; tx = "#7f1d1d"; }
  return '<span style="display:inline-flex; align-items:center; height:22px; padding:0 10px; border-radius:999px;'
    + 'border:1px solid ' + bd + '; background:' + bg + '; color:' + tx + '; font-weight:1000; font-size:11px;">'
    + esc_(x || "—") + '</span>';
}

  function esc_(s){
    return String(s || "")
      .replace(/&/g,"&amp;")
      .replace(/</g,"&lt;")
      .replace(/>/g,"&gt;")
      .replace(/"/g,"&quot;");
  }

  function fmtBRL_(v){
    var n = Number(v || 0);
    return n.toLocaleString("pt-BR", { style:"currency", currency:"BRL" });
  }

  var h = "";
  h += '<div style="font-family:Arial,Helvetica,sans-serif; color:#0f172a;">';
  h += '<div style="font-size:16px; font-weight:900;">Possíveis itens irregulares (Conformidade: <span style="color:#b91c1c;">ALERTA</span>)</div>';
  h += '<div style="margin-top:6px; font-size:13px; color:#334155;">Período analisado: <b>' + esc_(periodoTxt) + '</b></div>';
  h += '<div style="margin-top:10px; font-size:13px; color:#334155; line-height:1.4;">';
  h += 'Este e-mail é um <b>relatório de triagem</b> baseado em regras do Vektor (mesma lógica usada no chat). ';
  h += 'Recomendação: revisar os itens marcados como <b>ALERTA</b> e validar aderência à política.';
  h += '</div>';

  if (!rows.length) {
    h += '<div style="margin-top:14px; padding:12px; border:1px solid #e2e8f0; border-radius:10px; background:#f8fafc;">';
    h += 'Nenhum item em <b>ALERTA</b> encontrado no período.';
    h += '</div></div>';
    return h;
  }

  h += '<div style="margin-top:14px; font-size:13px; color:#334155;"><b>Total de linhas:</b> ' + rows.length + '</div>';

  h += '<div style="margin-top:10px; overflow:auto; border:1px solid #e2e8f0; border-radius:12px;">';
  h += '<table style="width:100%; border-collapse:collapse; min-width:980px;">';
  h += '<thead><tr style="background:#0b1220; color:#fff;">';
  h += '<th style="text-align:left; padding:10px; font-size:12px;">Data</th>';
  h += '<th style="text-align:right; padding:10px; font-size:12px;">Valor (R$)</th>';
  h += '<th style="text-align:left; padding:10px; font-size:12px;">Loja</th>';
  h += '<th style="text-align:left; padding:10px; font-size:12px;">Time</th>';
  h += '<th style="text-align:left; padding:10px; font-size:12px;">Item Comprado</th>';
  h += '<th style="text-align:left; padding:10px; font-size:12px;">Conformidade</th>';
  h += '<th style="text-align:left; padding:10px; font-size:12px;">Motivo</th>';
  h += '</tr></thead><tbody>';

  for (var i = 0; i < rows.length; i++) {
    var r = rows[i] || {};
    h += '<tr style="border-top:1px solid #e2e8f0;">';
    h += '<td style="padding:10px; font-size:12px;">' + esc_(r.dataTxt || r.data || "") + '</td>';
    h += '<td style="padding:10px; font-size:12px; text-align:right; font-weight:800;">' + esc_(fmtBRL_(r.valor || 0)) + '</td>';
    h += '<td style="padding:10px; font-size:12px;">' + esc_(r.loja || r.alias || "") + '</td>';
    h += '<td style="padding:10px; font-size:12px;">' + esc_(r.time || "") + '</td>';
    h += '<td style="padding:10px; font-size:12px;">' + esc_(r.item || r.descricao || "") + '</td>';
    h += '<td style="padding:10px; font-size:12px;">' + badgeHtml_(r.conformidade || r.status || "ALERTA") + '</td>';
    h += '<td style="padding:10px; font-size:12px;">' + esc_(r.motivo || "") + '</td>';
    h += '</tr>';
  }

  h += '</tbody></table></div></div>';
  return h;
}

function vektorGetHistoricoEnviosItensIrregularesResumo_() {
  vektorAssertFunctionAllowed_("vektorGetHistoricoEnviosItensIrregularesResumo_");

  try {
    var sh = vektorGetOrCreateItensIrregLogSheet_();
    var range = sh.getDataRange();
    var values = range.getValues();               // números / datas / etc.
    var displayValues = range.getDisplayValues(); // texto exatamente como aparece na planilha

    if (!values || values.length < 2) return { ok: true, rows: [] };

    var hdr = values[0];

    function idx_(name) {
      var n = String(name || "").toLowerCase().trim();
      for (var i = 0; i < hdr.length; i++) {
        if (String(hdr[i] || "").toLowerCase().trim() === n) return i;
      }
      return -1;
    }

    var iLojaKey = idx_("lojakey");
    var iTime = idx_("time");
    var iQtdItens = idx_("qtditens");
    var iValorTotal = idx_("valortotal");
    var iDataEnvioBR = idx_("dataenviobr");
    var iStatus = idx_("status");

    var agg = {}; // lojaKey -> resumo

    function fmtDataBR_(v) {
      if (!v) return "";
      try {
        // 1) Se veio Date real da planilha
        if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v.getTime())) {
          return Utilities.formatDate(v, Session.getScriptTimeZone() || "America/Sao_Paulo", "dd/MM/yyyy");
        }

        var s = String(v).trim();
        if (!s) return "";

        // 2) Se já está em dd/MM/yyyy, mantém
        if (/^\d{2}\/\d{2}\/\d{4}$/.test(s)) return s;

        // 3) Se vier ISO date-only (YYYY-MM-DD), NÃO usar new Date(s)
        //    (evita voltar 1 dia por fuso)
        var mIso = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
        if (mIso) {
          return mIso[3] + "/" + mIso[2] + "/" + mIso[1];
        }

        // 4) Se vier ISO com hora, tenta parsear
        //    Ex.: 2026-02-25T11:59:00-03:00
        if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
          var dIso = new Date(s);
          if (!isNaN(dIso.getTime())) {
            return Utilities.formatDate(dIso, Session.getScriptTimeZone() || "America/Sao_Paulo", "dd/MM/yyyy");
          }
        }

        // 5) Fallback final
        return s;
      } catch (e) {
        return String(v || "");
      }
}

    for (var r = 1; r < values.length; r++) {
      var row = values[r] || [];
      var st = String(row[iStatus] || "").toUpperCase().trim();
      if (st && st !== "SENT") continue; // considera só enviados com sucesso

      var lojaKey = String(row[iLojaKey] || "").trim();
      if (!lojaKey) continue;

      if (!agg[lojaKey]) {
        agg[lojaKey] = {
          lojaKey: lojaKey,
          time: String(row[iTime] || "").trim(),
          qtdEnvios: 0,
          qtdItens: 0,
          valorTotal: 0,
          ultimoEnvio: ""
        };
      }

      agg[lojaKey].qtdEnvios += 1;
      agg[lojaKey].qtdItens += Number(row[iQtdItens] || 0) || 0;
      agg[lojaKey].valorTotal += Number(row[iValorTotal] || 0) || 0;

      var rowDisp = displayValues[r] || [];
      var dt = String(rowDisp[iDataEnvioBR] || "").trim(); // usa TEXTO exibido na planilha
      if (dt) agg[lojaKey].ultimoEnvio = dt; // último lido (append)
      if (!agg[lojaKey].time && row[iTime]) agg[lojaKey].time = String(row[iTime]).trim();
    }

    var out = Object.keys(agg).map(function(k){ return agg[k]; });

    out.sort(function(a,b){
      return (Number(b.qtdEnvios||0) - Number(a.qtdEnvios||0))
          || (Number(b.valorTotal||0) - Number(a.valorTotal||0));
    });

    return { ok: true, rows: out };
  } catch (e) {
    return { ok: false, error: String(e && e.message ? e.message : e) };
  }
}

function vektorGetHistoricoEnviosItensIrregularesResumo() {
  return vektorGetHistoricoEnviosItensIrregularesResumo_();
}

function RESETAR_GATE_EMAIL_OFENSORAS_SEMANA() {
  PropertiesService.getScriptProperties().deleteProperty("VEKTOR_OFENSORAS_SIG_ULT_ENVIO");
  Logger.log("Gate resetado: VEKTOR_OFENSORAS_SIG_ULT_ENVIO removida. Próximo disparo enviará novamente.");
}

function LIMPAR_ALERTA_LIMITE() {
  var props = PropertiesService.getScriptProperties();

  // 1) Gate do envio de limite (porteiro ENVIAR_EMAIL_LIMITE_CLARA)
  props.deleteProperty("VEKTOR_SIG_BASECLARA_PROCESSADA");

  // 2) Anti-spam do ciclo (se existir no seu fluxo)
  try {
    var cicloKey = getCicloKey06a05_();
    props.deleteProperty("VEKTOR_ALERTS_SENT_" + cicloKey);
  } catch (e) {}

  // 3) Mantém sua limpeza antiga (se ainda for usada em outra parte)
  props.deleteProperty("VEKTOR_HISTPEND_LAST_SIG");

  Logger.log("Gate do alerta de LIMITE limpo com sucesso.");
}

function vektorStatusSistema() {
  // Gate por função (deve existir na VEKTOR_ACESSOS para o ROLE)
  vektorAssertFunctionAllowed_("vektorStatusSistema");

  // Admin agora vem do RBAC (VEKTOR_EMAILS)
  var ctx = vektorGetUserRole_(); // { email, role }
    // ✅ garante que quem abriu o modal entre como "ativo hoje"
  try {
    vektorTrackActiveUserToday_(ctx.email);
  } catch (eTrack) {}
  var isAdmin = String(ctx.role || "").toLowerCase() === "administrador";

  const file = DriveApp.getFileById(BASE_CLARA_ID);
  const ultimaAtualizacao = file.getLastUpdated();

  // Sempre retorna Base e Status Geral
  const baseClaraTxt = Utilities.formatDate(
    ultimaAtualizacao,
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm"
  );

  // Não-admin: retorna só o necessário (segurança)
  if (!isAdmin) {
    return {
      baseClara: baseClaraTxt,
      geral: "Em operação"
    };
  }
  // Serviços Google (Apps Script / E-mail): quota diária restante
let googleTxt = "OK";
try {
  const quota = MailApp.getRemainingDailyQuota();
  googleTxt = "OK | Quota e-mail restante hoje: " + quota;
} catch (e) {
  // Se falhar, devolve a falha (pra você enxergar no modal em vez de mascarar)
  googleTxt = "Falha ao ler quota de e-mail: " + (e && e.message ? e.message : String(e));
}

  // ===== BigQuery: healthcheck real (Job + métricas) =====
let bqTxt = "Indisponível";
try {
  const t0 = Date.now();

  const req = {
    query: "SELECT 1 AS ok",
    useLegacySql: false,
    timeoutMs: 10000 // evita travar o modal
  };

  const r = BigQuery.Jobs.query(req, PROJECT_ID);

  const ms = Date.now() - t0;
  const jobId = r && r.jobReference ? r.jobReference.jobId : "";
  const loc = r && r.jobReference ? r.jobReference.location : "";
  const complete = r && r.jobComplete === true;

  // totalBytesProcessed costuma vir como string
  const bytes = r && r.statistics ? r.statistics.totalBytesProcessed : "";
  const cacheHit = r && r.statistics ? r.statistics.cacheHit : "";

  if (!complete) {
    bqTxt = `Instável | job não completou | ${ms}ms` + (jobId ? ` | jobId ${jobId}` : "");
  } else {
    bqTxt =
      `OK | ${ms}ms` +
      (bytes ? ` | bytes ${bytes}` : "") +
      (cacheHit !== "" ? ` | cacheHit ${cacheHit}` : "") +
      (loc ? ` | ${loc}` : "") +
      (jobId ? ` | jobId ${jobId}` : "");
  }
} catch (eBQ) {
  bqTxt = "Falha BigQuery: " + (eBQ && eBQ.message ? eBQ.message : String(eBQ));
}

  // ===== Vertex AI: resumo de uso do mês =====
  var vertex = vektorVertexGetUsageSummary_();
  var vertexTxt = vertex.calls > 0
    ? ("OK | " + vertex.calls + " chamadas no mês")
    : "Sem uso registrado no mês";

  // Admin: retorna completo
  return {
  baseClara: baseClaraTxt,
  jobs: "Executados com sucesso",
  google: googleTxt,
  bigquery: bqTxt,
  alertas: "Ativos",
  usuariosAtivosHoje: vektorGetActiveUsersTodayCount_(Session.getScriptTimeZone()),
  geral: "Em operação",

  // Vertex AI
  vertexStatus: vertexTxt,
  vertexCallsMes: vertex.calls || 0,
  vertexPromptTokensMes: vertex.promptTokens || 0,
  vertexOutputTokensMes: vertex.outputTokens || 0,
  vertexTotalTokensMes: vertex.totalTokens || 0,
  vertexEstimatedUsdMes: vektorFmtUsdWithBrl_(vertex.estimatedUsd || 0),
  vertexLastModel: vertex.lastModelVersion || vertex.lastModel || VEKTOR_VERTEX_MODEL,
  vertexLastTokens: vertex.lastTotalTokens || 0,
  vertexLastEstimatedUsd: vektorFmtUsdWithBrl_(vertex.lastEstimatedUsd || 0),
  vertexLastAt: vertex.lastAt || "—"
  };
}

// ===============================
// FIX DEFINITIVO: parser ISO seguro (evita TypeError m[1])
// ===============================
function vektorParseIsoDate_(iso) {
  if (!iso) return null;

  // já é Date
  if (Object.prototype.toString.call(iso) === "[object Date]") {
    var d0 = iso;
    return isNaN(d0.getTime()) ? null : d0;
  }

  var s = String(iso || "").trim();

  // aceita: "2026-01-26T00:00:00.000Z" ou "2026-01-26"
  var m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
  if (!m) return null;

  var y = Number(m[1]);
  var mo = Number(m[2]) - 1;
  var d = Number(m[3]);

  var dt = new Date(y, mo, d);
  if (isNaN(dt.getTime())) return null;

  // padroniza para 00:00:00
  dt.setHours(0, 0, 0, 0);
  return dt;
}

// ===============================
// ANALISE DE GASTOS (BACKEND) — BaseClara
// - Admin: tudo
// - Não-admin: lojas permitidas via aba "Emails"
//   - Coluna B: LojaNorm
//   - Coluna "E-mail Gerente Regional"
//   - Coluna H: "Usuários adicionais"
// ===============================

function vektorNormLower_(s){ return String(s || "").trim().toLowerCase(); }

function vektorSplitUsers_(s){
  var raw = String(s || "").trim();
  if (!raw) return [];
  return raw
    .split(/[;,|\n\r]+/g)
    .map(function(x){ return String(x||"").trim(); })
    .filter(function(x){ return !!x; });
}

function vektorHeaderIndex_(header, names){
  header = (header || []).map(function(h){ return String(h||"").trim().toLowerCase(); });
  for (var i=0; i<names.length; i++){
    var target = String(names[i]||"").trim().toLowerCase();
    for (var c=0; c<header.length; c++){
      if (header[c] === target) return c;
    }
  }
  return -1;
}

/**
 * Retorna lojas permitidas do usuário (não-admin) via aba "Emails".
 * Admin => null (significa "todas").
 */
function vektorGetAllowedLojasFromEmails_(userEmail){
  var em = vektorNormLower_(userEmail);
  if (!em) return [];

  // Admin e Analista Pro vê tudo
  var ctx = vektorGetUserRole_();
  var role = String(ctx && ctx.role ? ctx.role : "").trim().toLowerCase();
  if (role === "administrador" || role === "analista pro" || role === "marketing" || role === "analista" ) return null;

  var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
  var sh = ss.getSheetByName("Emails");
  if (!sh) throw new Error('Aba "Emails" não encontrada na planilha BaseClara.');

  var values = sh.getDataRange().getValues();
  if (!values || values.length < 2) return [];

  var header = values[0];

  var iLoja = 1; // coluna B = LojaNorm (fixo como você disse)
  var iGer = vektorHeaderIndex_(header, ["e-mail gerente regional", "email gerente regional", "e-mail ger regional", "email ger regional"]);
  var iAdd = 7; // coluna H (0-based 7)

  if (iGer < 0) throw new Error('Não encontrei a coluna "E-mail Gerente Regional" na aba Emails.');

  var allowed = {};
  for (var r=1; r<values.length; r++){
    var row = values[r];
    var lojaNorm = normalizarLojaNumero_(row[iLoja]);
    if (!lojaNorm) continue;

    var ger = vektorNormLower_(row[iGer]);
    var addList = vektorSplitUsers_(row[iAdd]).map(vektorNormLower_);

    var ok = false;
    if (ger && ger === em) ok = true;
    if (!ok && addList && addList.length){
      for (var k=0; k<addList.length; k++){
        if (addList[k] && addList[k] === em) { ok = true; break; }
      }
    }

    if (ok) allowed[String(lojaNorm)] = true;
  }

  return Object.keys(allowed).sort(function(a,b){ return Number(a)-Number(b); });
}

/**
 * Meta da página: times/lojas/categorias permitidos
 * - Admin: times/lojas/categorias de toda BaseClara
 * - Não-admin: lojas filtradas pela aba Emails, e times/categorias derivadas dessas lojas
 */
function getAnaliseGastosMeta(){
  vektorAssertFunctionAllowed_("getAnaliseGastosMeta");

  try {
    var ctx = vektorGetUserRole_();
    var email = (ctx && ctx.email)
      ? String(ctx.email).trim().toLowerCase()
      : String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();

    if (!email) throw new Error("Não foi possível identificar seu e-mail Google.");

    // null => admin (vê tudo); array => lojas permitidas
    var allowedLojas = vektorGetAllowedLojasFromEmails_(email);

    var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
    var sh = ss.getSheetByName("BaseClara");
    if (!sh) throw new Error('Aba "BaseClara" não encontrada.');

    var lastRow = sh.getLastRow();
    if (lastRow < 2) {
      return {
        ok: true,
        lojas: [],
        times: [],
        categorias: [],
        combos: [],
        anos: [],
        mesesPorAno: {}
      };
    }

    // Lê só A:V (22 colunas), que já contém tudo que os filtros precisam
    var values = sh.getRange(2, 1, lastRow - 1, 22).getValues();

    // índices locais dentro de A:V (0-based)
    var iData      = 0;   // A
    var iLojaAlias = 7;   // H
    var iCategoria = 13;  // N
    var iTime      = 17;  // R
    var iLojaNum   = 21;  // V

    var lojasSet = {};
    var timesSet = {};
    var catsSet  = {};
    var comboMap = {};
    var anosSet = {};
    var mesesPorAnoMap = {};

    var allowedSet = null;
    if (Array.isArray(allowedLojas)) {
      allowedSet = {};
      allowedLojas.forEach(function(x){
        var n = normalizarLojaNumero_(x);
        if (n != null) {
          allowedSet[String(n)] = true;
          allowedSet[String(n).padStart(4, "0")] = true;
        }
      });
    }

    function toISODateMeta_(v){
      if (v instanceof Date) {
        var y = v.getFullYear();
        var m = String(v.getMonth() + 1).padStart(2, "0");
        var d = String(v.getDate()).padStart(2, "0");
        return y + "-" + m + "-" + d;
      }

      var s = String(v || "").trim();
      if (!s) return "";

      var m1 = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
      if (m1) return m1[3] + "-" + m1[2] + "-" + m1[1];

      var m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (m2) return s;

      return "";
    }

    function lojaLabelMeta_(lojaNum, lojaAlias){
      var num = String(lojaNum || "").replace(/\D/g, "");
      if (!num) return "";
      var num4 = num.padStart(4, "0");
      var alias = String(lojaAlias || "").trim();
      return alias ? (num4 + " - " + alias) : num4;
    }

    for (var i = 0; i < values.length; i++) {
      var row = values[i];

      var lojaNum = normalizarLojaNumero_(row[iLojaNum]);
      if (lojaNum == null) continue;

      if (allowedSet && !allowedSet[String(lojaNum)] && !allowedSet[String(lojaNum).padStart(4, "0")]) {
        continue;
      }

      var lojaAlias = String(row[iLojaAlias] || "").trim();
      var categoria = String(row[iCategoria] || "").trim();
      var time = String(row[iTime] || "").trim();
      var dataIso = toISODateMeta_(row[iData]);

      var lojaLabel = lojaLabelMeta_(lojaNum, lojaAlias);

      if (lojaLabel) lojasSet[lojaLabel] = true;
      if (time) timesSet[time] = true;
      if (categoria) catsSet[categoria] = true;

      var comboKey = [time || "", lojaLabel || "", categoria || ""].join("¦");
      comboMap[comboKey] = {
        time: time || "",
        loja: lojaLabel || "",
        categoria: categoria || ""
      };

      if (dataIso) {
        var ano = dataIso.slice(0, 4);
        var mes = dataIso.slice(5, 7);

        if (ano && /^\d{4}$/.test(ano)) {
          anosSet[ano] = true;
          if (!mesesPorAnoMap[ano]) mesesPorAnoMap[ano] = {};
          if (mes && /^\d{2}$/.test(mes)) mesesPorAnoMap[ano][mes] = true;
        }
      }
    }

    var lojas = Object.keys(lojasSet).sort(function(a, b){
      var na = Number(String(a).split(" - ")[0].replace(/\D/g, "")) || 0;
      var nb = Number(String(b).split(" - ")[0].replace(/\D/g, "")) || 0;
      return na - nb;
    });

    var times = Object.keys(timesSet).sort();
    var categorias = Object.keys(catsSet).sort();

    var combos = Object.keys(comboMap).map(function(k){
      return comboMap[k];
    }).sort(function(a, b){
      var ta = String(a.time || "");
      var tb = String(b.time || "");
      if (ta !== tb) return ta.localeCompare(tb, "pt-BR");

      var la = String(a.loja || "");
      var lb = String(b.loja || "");
      if (la !== lb) return la.localeCompare(lb, "pt-BR");

      return String(a.categoria || "").localeCompare(String(b.categoria || ""), "pt-BR");
    });

    var anos = Object.keys(anosSet).sort(function(a,b){
      return Number(a) - Number(b);
    });

    var mesesPorAno = {};
    Object.keys(mesesPorAnoMap).forEach(function(ano){
      mesesPorAno[ano] = Object.keys(mesesPorAnoMap[ano]).sort(function(a,b){
        return Number(a) - Number(b);
      });
    });

    return {
      ok: true,
      lojas: lojas,
      times: times,
      categorias: categorias,
      combos: combos,
      anos: anos,
      mesesPorAno: mesesPorAno
    };

  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

/**
 * Dataset da página: retorna tx já filtradas por ACL e pelo recorte solicitado.
 * req: {dtIni, dtFim, time, loja, categoria}
 */
function getAnaliseGastosDataset(req){
  vektorAssertFunctionAllowed_("getAnaliseGastosDataset");

  try {
    req = req || {};
    var dtIni = String(req.dtIni || "").trim();   // yyyy-mm-dd
    var dtFim = String(req.dtFim || "").trim();   // yyyy-mm-dd
    
    var fTimeArr = Array.isArray(req.time) ? req.time.map(String) : (req.time ? [String(req.time)] : []);
    var fLojaArr = Array.isArray(req.loja) ? req.loja.map(String) : (req.loja ? [String(req.loja)] : []);
    var fCatArr  = Array.isArray(req.categoria) ? req.categoria.map(String) : (req.categoria ? [String(req.categoria)] : []);

    fTimeArr = fTimeArr.map(function(s){ return String(s||"").trim(); }).filter(Boolean);
    fLojaArr = fLojaArr.map(function(s){ return String(s||"").trim(); }).filter(Boolean);
    fCatArr  = fCatArr.map(function(s){ return String(s||"").trim(); }).filter(Boolean);

    if (!dtIni || !dtFim) throw new Error("Informe dtIni e dtFim.");
    if (dtIni > dtFim) throw new Error("Período inválido: dtIni > dtFim.");

    var ctx = vektorGetUserRole_();
    var email = (ctx && ctx.email)
      ? String(ctx.email).trim().toLowerCase()
      : String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();

    if (!email) throw new Error("Não foi possível identificar seu e-mail Google.");

    // null => admin; array => lojas permitidas (LojaNorm da aba Emails)
    var allowedLojas = vektorGetAllowedLojasFromEmails_(email);

    var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
    var sh = ss.getSheetByName("BaseClara");
    if (!sh) throw new Error('Aba "BaseClara" não encontrada.');

    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();

    if (lastRow < 2) {
      return { ok: true, meta: { periodoTxt: dtIni + " a " + dtFim, totalTx: 0 }, tx: [] };
    }

    var values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // =========================
    // ÍNDICES FIXOS (0-based) — BaseClara
    // =========================
    var iData      = 0;   // A - Data da Transação
    var iEstab     = 2;   // C - Transação (Estabelecimento)
    var iValor     = 5;   // F - Valor em R$
    var iLojaAlias = 7;   // H - Alias do Cartão (opcional exibição)
    var iCategoria = 13;  // N - Categoria da Compra ✅
    var iTime      = 17;  // R - Grupos (Time)
    var iDesc      = 20;  // U - Descrição (Item comprado)
    var iLojaNum   = 21;  // V - LojaNum ✅ ACL/Filtro

    function toISODate_(v){
      if (v instanceof Date) {
        var y = v.getFullYear();
        var m = String(v.getMonth() + 1).padStart(2, "0");
        var d = String(v.getDate()).padStart(2, "0");
        return y + "-" + m + "-" + d;
      }

      var s = String(v || "").trim();
      if (!s) return "";

      // dd/mm/yyyy
      var m1 = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
      if (m1) return m1[3] + "-" + m1[2] + "-" + m1[1];

      // yyyy-mm-dd
      var m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (m2) return s;

      return "";
    }

    function toBR_(iso){
      if (!iso) return "";
      var p = String(iso).split("-");
      if (p.length === 3) return p[2] + "/" + p[1] + "/" + p[0];
      return iso;
    }

    var tx = [];

    values.forEach(function(row){
      var lojaNum = normalizarLojaNumero_(row[iLojaNum]);
      if (!lojaNum) return;
      lojaNum = String(lojaNum);

      // ACL (comparando LojaNorm da aba Emails com BaseClara!V = LojaNum)
      if (Array.isArray(allowedLojas)) {
        if (allowedLojas.indexOf(lojaNum) < 0) return;
      }

      var dataISO = toISODate_(row[iData]);
      if (!dataISO) return;

      // período
      if (dataISO < dtIni || dataISO > dtFim) return;

      var time = String(row[iTime] || "").trim();
      if (!time) time = "N/D";

      var categoria = String(row[iCategoria] || "").trim();
      if (!categoria) categoria = "Sem categoria";

      var estabelecimento = String(row[iEstab] || "").trim();
      if (!estabelecimento) estabelecimento = "—";

      var descricao = String(row[iDesc] || "").trim();
      if (!descricao) descricao = "—";

      var valor = Number(row[iValor]) || 0;

      // filtro Loja (agora é LojaNum)
      if (fTimeArr.length && fTimeArr.indexOf(time) < 0) return;
      if (fLojaArr.length && fLojaArr.indexOf(lojaNum) < 0) return;
      if (fCatArr.length  && fCatArr.indexOf(categoria) < 0) return;

      // Exibição: se quiser mostrar alias junto, você pode concatenar
      // var lojaAlias = String(row[iLojaAlias] || "").trim();
      // var lojaExib = lojaAlias ? (lojaNum + " - " + lojaAlias) : lojaNum;

      tx.push({
        dataISO: dataISO,
        dataBR: toBR_(dataISO),
        loja: lojaNum, // ✅ mantendo filtro/coluna Loja como número da loja
        time: time,
        categoria: categoria,
        estabelecimento: estabelecimento,
        valor: valor,
        descricao: descricao
      });
    });

    return {
      ok: true,
      meta: {
        periodoTxt: dtIni + " a " + dtFim,
        totalTx: tx.length
      },
      tx: tx
    };

  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

/**
 * Medianas por Loja x Categoria
 * req: {dtIni, dtFim, time:[], loja:[], categoria:[]}
 * Retorna:
 * - rows: [{loja,categoria,mediana,max,min}]
 * - metrics: {median, mean, normalization, min, max, countTx}
 */
function getAnaliseGastosMedianas(req){
  vektorAssertFunctionAllowed_("getAnaliseGastosMedianas");

  try {
    req = req || {};
    var dtIni = String(req.dtIni || "").trim();
    var dtFim = String(req.dtFim || "").trim();

    if (!dtIni || !dtFim) throw new Error("Informe dtIni e dtFim.");
    if (dtIni > dtFim) throw new Error("Período inválido: dtIni > dtFim.");

    var fTimeArr = Array.isArray(req.time) ? req.time.map(String) : (req.time ? [String(req.time)] : []);
    var fLojaArr = Array.isArray(req.loja) ? req.loja.map(String) : (req.loja ? [String(req.loja)] : []);
    var fCatArr  = Array.isArray(req.categoria) ? req.categoria.map(String) : (req.categoria ? [String(req.categoria)] : []);

    fTimeArr = fTimeArr.map(function(s){ return String(s||"").trim(); }).filter(Boolean);
    fLojaArr = fLojaArr.map(function(s){ return String(s||"").trim(); }).filter(Boolean);
    fCatArr  = fCatArr.map(function(s){ return String(s||"").trim(); }).filter(Boolean);

    var ctx = vektorGetUserRole_();
    var email = (ctx && ctx.email)
      ? String(ctx.email).trim().toLowerCase()
      : String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();

    if (!email) throw new Error("Não foi possível identificar seu e-mail Google.");

    // null => admin; array => lojas permitidas
    var allowedLojas = vektorGetAllowedLojasFromEmails_(email);

    var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
    var sh = ss.getSheetByName("BaseClara");
    if (!sh) throw new Error('Aba "BaseClara" não encontrada.');

    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2) {
      return { ok:true, rows:[], metrics:{ median:0, mean:0, normalization:0, min:0, max:0, countTx:0 } };
    }

    var values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // índices BaseClara (0-based)
    var iData      = 0;   // A
    var iValor     = 5;   // F
    var iCategoria = 13;  // N
    var iTime      = 17;  // R
    var iLojaNum   = 21;  // V

    function toISODate_(v){
      if (v instanceof Date) {
        var y = v.getFullYear();
        var m = String(v.getMonth() + 1).padStart(2, "0");
        var d = String(v.getDate()).padStart(2, "0");
        return y + "-" + m + "-" + d;
      }
      var s = String(v || "").trim();
      if (!s) return "";
      var m1 = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
      if (m1) return m1[3] + "-" + m1[2] + "-" + m1[1];
      var m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (m2) return s;
      return "";
    }

    function median_(arr){
      var a = (arr || []).slice().map(function(x){ return Number(x)||0; }).sort(function(x,y){ return x-y; });
      if (!a.length) return 0;
      var mid = Math.floor(a.length / 2);
      if (a.length % 2) return a[mid];
      return (a[mid-1] + a[mid]) / 2;
    }

    var groups = {}; // key loja||cat -> {loja,categoria,vals:[],min,max}
    var allVals = [];

    values.forEach(function(row){
      var lojaNum = normalizarLojaNumero_(row[iLojaNum]);
      if (!lojaNum) return;
      lojaNum = String(lojaNum);

      // ACL
      if (Array.isArray(allowedLojas)) {
        if (allowedLojas.indexOf(lojaNum) < 0) return;
      }

      var dataISO = toISODate_(row[iData]);
      if (!dataISO) return;
      if (dataISO < dtIni || dataISO > dtFim) return;

      var time = String(row[iTime] || "").trim();
      if (!time) time = "N/D";

      var categoria = String(row[iCategoria] || "").trim();
      if (!categoria) categoria = "Sem categoria";

      var valor = Number(row[iValor]) || 0;

      // filtros multi
      if (fTimeArr.length && fTimeArr.indexOf(time) < 0) return;
      if (fLojaArr.length && fLojaArr.indexOf(lojaNum) < 0) return;
      if (fCatArr.length  && fCatArr.indexOf(categoria) < 0) return;

      var key = lojaNum + "||" + categoria;
      if (!groups[key]) {
        groups[key] = { loja: lojaNum, categoria: categoria, vals: [], min: null, max: null };
      }

      groups[key].vals.push(valor);
      allVals.push(valor);

      if (groups[key].min === null || valor < groups[key].min) groups[key].min = valor;
      if (groups[key].max === null || valor > groups[key].max) groups[key].max = valor;
    });

    var rows = Object.keys(groups).map(function(k){
      var g = groups[k];
      return {
        loja: g.loja,
        categoria: g.categoria,
        mediana: median_(g.vals),
        max: Number(g.max || 0),
        min: Number(g.min || 0)
      };
    });

    // ordena: maior mediana primeiro
    rows.sort(function(a,b){ return (Number(b.mediana||0) - Number(a.mediana||0)); });

    // métricas gerais
    var globalMedian = median_(allVals);
    var globalMean = 0;
    if (allVals.length) {
      var sum = allVals.reduce(function(s,x){ return s + (Number(x)||0); }, 0);
      globalMean = sum / allVals.length;
    }
    var globalMin = allVals.length ? Math.min.apply(null, allVals) : 0;
    var globalMax = allVals.length ? Math.max.apply(null, allVals) : 0;

    var norm = 0;
    var denom = (globalMax - globalMin);
    if (denom > 0) norm = (globalMedian - globalMin) / denom;

    return {
      ok: true,
      rows: rows,
      metrics: {
        median: globalMedian,
        mean: globalMean,
        normalization: norm,
        min: globalMin,
        max: globalMax,
        countTx: allVals.length
      }
    };

  } catch (e) {
    return { ok:false, error:(e && e.message) ? e.message : String(e) };
  }
}

/**
 * Macro Visões (modal independente da tela principal)
 * Retorna 2 séries mensais:
 * - serieLojas: somatória mensal (opcional filtro de loja)
 * - serieCategorias: somatória mensal (opcional filtro de categoria)
 * Ambas respeitam ACL/permissão do usuário.
 *
 * req: { loja, categoria }
 */
function getAnaliseGastosMacroVisoesDataset(req){
  vektorAssertFunctionAllowed_("getAnaliseGastosMacroVisoesDataset");

  try {
    req = req || {};
    var fLoja = String(req.loja || "").trim();        // LojaNum (string)
    var fCat  = String(req.categoria || "").trim();   // categoria
    var fAno  = String(req.ano || "").trim();         // "2025", "2026"...
    if (fAno && !/^\d{4}$/.test(fAno)) fAno = "";

    var ctx = vektorGetUserRole_();
    var email = (ctx && ctx.email)
      ? String(ctx.email).trim().toLowerCase()
      : String(Session.getActiveUser().getEmail() || "").trim().toLowerCase();

    if (!email) throw new Error("Não foi possível identificar seu e-mail Google.");

    // null => admin; array => lojas permitidas
    var allowedLojas = vektorGetAllowedLojasFromEmails_(email);

    var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
    var sh = ss.getSheetByName("BaseClara");
    if (!sh) throw new Error('Aba "BaseClara" não encontrada.');

    var lastRow = sh.getLastRow();
    var lastCol = sh.getLastColumn();
    if (lastRow < 2) {
      return {
        ok: true,
        filtros: { loja: fLoja, categoria: fCat, ano: fAno },
        serieLojas: { labels: [], totais: [], variacoesPct: [] },
        serieCategorias: { labels: [], totais: [], variacoesPct: [] },
        top10Lojas: { labels: [], totais: [] },
        top10Categorias: { labels: [], totais: [] }
      };
    }

    var values = sh.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // BaseClara (0-based)
    var iData      = 0;   // A
    var iValor     = 5;   // F
    var iCategoria = 13;  // N
    var iLojaNum   = 21;  // V

    function toISODate_(v){
      if (v instanceof Date) {
        var y = v.getFullYear();
        var m = String(v.getMonth() + 1).padStart(2, "0");
        var d = String(v.getDate()).padStart(2, "0");
        return y + "-" + m + "-" + d;
      }
      var s = String(v || "").trim();
      if (!s) return "";

      var m1 = s.match(/^(\d{2})\/(\d{2})\/(\d{4})$/);
      if (m1) return m1[3] + "-" + m1[2] + "-" + m1[1];

      var m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
      if (m2) return s;

      return "";
    }

    function buildSerie_(monthlyMap, anoBase){
      var keys = [];
      var labels = [];
      var totais = [];
      var variacoesPct = [];

      function monthKeyToLabel_(ym){
        var y = String(ym || "").slice(0,4);
        var m = String(ym || "").slice(5,7);
        var map = {
          "01":"jan","02":"fev","03":"mar","04":"abr","05":"mai","06":"jun",
          "07":"jul","08":"ago","09":"set","10":"out","11":"nov","12":"dez"
        };
        return (map[m] || m) + "/" + y;
      }

      if (anoBase && /^\d{4}$/.test(String(anoBase))) {
        // força jan..dez do ano selecionado
        for (var m = 1; m <= 12; m++) {
          keys.push(String(anoBase) + "-" + String(m).padStart(2, "0"));
        }
      } else {
        // sem ano -> usa últimos 12 meses disponíveis
        var all = Object.keys(monthlyMap || {}).sort(); // yyyy-mm
        keys = all.slice(-12);
      }

      var prev = null;
      for (var i = 0; i < keys.length; i++){
        var ym = keys[i];
        var total = Number(monthlyMap[ym] || 0) || 0;

        labels.push(monthKeyToLabel_(ym));
        totais.push(total);

        if (prev === null || prev === 0) {
          variacoesPct.push(0);
        } else {
          variacoesPct.push(((total - prev) / prev) * 100);
        }

        prev = total;
      }

      return { labels: labels, totais: totais, variacoesPct: variacoesPct };
    }

    function buildTopN_(mapObj, topN){
      topN = Number(topN) || 10;
      var arr = Object.keys(mapObj || {}).map(function(k){
        return {
          nome: String(k || ""),
          total: Number(mapObj[k] || 0) || 0
        };
      });

      arr = arr.filter(function(x){ return x.total !== 0; });

      arr.sort(function(a,b){
        if (b.total !== a.total) return b.total - a.total;
        return String(a.nome).localeCompare(String(b.nome));
      });

      arr = arr.slice(0, topN);

      return {
        labels: arr.map(function(x){ return x.nome; }),
        totais: arr.map(function(x){ return x.total; })
      };
    }

    var monthlyLojas = {};      // série mensal para visão de lojas (opcional filtro de loja)
    var monthlyCats  = {};      // série mensal para visão de categorias (opcional filtro de categoria)
    var totalByLoja = {};       // top 10 lojas (somatória no período/filtros)
    var totalByCategoria = {};  // top 10 categorias (somatória no período/filtros)

    values.forEach(function(row){
      var lojaNum = normalizarLojaNumero_(row[iLojaNum]);
      if (!lojaNum) return;
      lojaNum = String(lojaNum);

      // ACL
      if (Array.isArray(allowedLojas)) {
        if (allowedLojas.indexOf(lojaNum) < 0) return;
      }

      var dataISO = toISODate_(row[iData]);
      if (!dataISO) return;

      var ym = String(dataISO).slice(0, 7); // yyyy-mm
      if (!/^\d{4}-\d{2}$/.test(ym)) return;

      var anoRow = String(dataISO).slice(0, 4);

      // Se usuário selecionou ano, usa somente aquele ano
      if (fAno && anoRow !== fAno) return;

      var categoria = String(row[iCategoria] || "").trim();
      if (!categoria) categoria = "Sem categoria";

      var valor = Number(row[iValor]) || 0;

      // ✅ Top 10 Lojas e Top 10 Categorias respeitam TODOS os filtros do modal
      // (ano + loja + categoria)
      var passaFiltroLoja = (!fLoja || lojaNum === fLoja);
      var passaFiltroCat  = (!fCat || categoria === fCat);

      if (passaFiltroLoja && passaFiltroCat) {
        totalByLoja[lojaNum] = (Number(totalByLoja[lojaNum]) || 0) + valor;
        totalByCategoria[categoria] = (Number(totalByCategoria[categoria]) || 0) + valor;
      }

      // Série Lojas (deve respeitar loja E categoria do modal)
      if ((!fLoja || lojaNum === fLoja) && (!fCat || categoria === fCat)) {
        monthlyLojas[ym] = (Number(monthlyLojas[ym]) || 0) + valor;
      }

      // Série Categorias (deve respeitar categoria E loja do modal)
      if ((!fCat || categoria === fCat) && (!fLoja || lojaNum === fLoja)) {
        monthlyCats[ym] = (Number(monthlyCats[ym]) || 0) + valor;
      }
    });

    return {
      ok: true,
      filtros: { loja: fLoja, categoria: fCat, ano: fAno },
      serieLojas: buildSerie_(monthlyLojas, fAno),
      serieCategorias: buildSerie_(monthlyCats, fAno),
      top10Lojas: buildTopN_(totalByLoja, 10),
      top10Categorias: buildTopN_(totalByCategoria, 10)
    };

  } catch (e) {
    return { ok: false, error: (e && e.message) ? e.message : String(e) };
  }
}

// =====================================================
// FLUXO NUMERÁRIO SAP — BACKEND
// =====================================================

var VEKTOR_SAP_PROJECT_ID = "gruposbf-data-lake";
var VEKTOR_SAP_SANGRIA_LOG_SHEET = "VEKTOR_SAP_SANGRIA_LOG";

function vektorSapFmtDateBr_(v) {
  try {
    if (!v) return "—";
    if (v instanceof Date) {
      return Utilities.formatDate(v, Session.getScriptTimeZone() || "America/Sao_Paulo", "dd/MM/yyyy");
    }

    var s = String(v).trim();
    if (!s) return "—";

    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) {
      var p = s.split("-");
      return p[2] + "/" + p[1] + "/" + p[0];
    }

    if (/^\d{4}-\d{2}-\d{2}T/.test(s)) {
      var d = new Date(s);
      if (!isNaN(d.getTime())) {
        return Utilities.formatDate(d, Session.getScriptTimeZone() || "America/Sao_Paulo", "dd/MM/yyyy");
      }
    }

    return s;
  } catch (e) {
    return String(v || "—");
  }
}

function vektorSapParseDateIso_(v) {
  if (!v) return null;
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone() || "America/Sao_Paulo", "yyyy-MM-dd");
  }

  var s = String(v).trim();
  if (!s) return null;

  var m1 = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m1) return s;

  var m2 = s.match(/^(\d{4})-(\d{2})-(\d{2})T/);
  if (m2) return m2[1] + "-" + m2[2] + "-" + m2[3];

  return null;
}

function vektorSapFmtMoneyBr_(n) {
  n = Number(n || 0) || 0;
  return n.toLocaleString("pt-BR", { style: "currency", currency: "BRL" });
}

function vektorSapNormDocKey_(row) {
  var loja = String(row.loja || "").trim().toUpperCase();
  var numdoc = String(row.numdoc || "").trim();
  var dt = String(row.dataLancIso || "").trim();
  var val = String(Number(row.valor || 0).toFixed(2));
  return [loja, numdoc, dt, val].join("|");
}

function vektorSapGetOrCreateLogSheet_() {
  var ss = SpreadsheetApp.openById(BASE_CLARA_ID);
  var sh = ss.getSheetByName(VEKTOR_SAP_SANGRIA_LOG_SHEET);

  if (!sh) {
    sh = ss.insertSheet(VEKTOR_SAP_SANGRIA_LOG_SHEET);
    sh.appendRow([
      "createdAt",
      "userEmail",
      "docKey",
      "lojaKey",
      "time",
      "numdoc",
      "datalanc",
      "valor",
      "to",
      "cc",
      "status",
      "error"
    ]);
    sh.getRange(1,1,1,12).setFontWeight("bold");
    sh.setFrozenRows(1);
  }

  return sh;
}

function vektorSapJaFoiNotificado_(docKey) {
  var sh = vektorSapGetOrCreateLogSheet_();
  var lr = sh.getLastRow();
  if (lr < 2) return false;

  var values = sh.getRange(2, 1, lr - 1, 12).getValues();
  for (var i = values.length - 1; i >= 0; i--) {
    var row = values[i];
    var dk = String(row[2] || "").trim();
    var st = String(row[10] || "").trim().toUpperCase();
    if (dk === docKey && st === "SENT") return true;
  }
  return false;
}

function vektorSapLogNotificacao_(payload) {
  payload = payload || {};
  var sh = vektorSapGetOrCreateLogSheet_();
  var tz = Session.getScriptTimeZone() || "America/Sao_Paulo";
  var ts = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd HH:mm:ss");

  sh.appendRow([
    ts,
    String(payload.userEmail || "").trim().toLowerCase(),
    String(payload.docKey || "").trim(),
    String(payload.lojaKey || "").trim(),
    String(payload.time || "").trim(),
    String(payload.numdoc || "").trim(),
    String(payload.datalanc || "").trim(),
    Number(payload.valor || 0) || 0,
    String(payload.to || "").trim(),
    String(payload.cc || "").trim(),
    String(payload.status || "").trim(),
    String(payload.error || "").trim()
  ]);
}

function vektorSapRunQuery_(sql) {
  var req = {
    query: sql,
    useLegacySql: false,
    timeoutMs: 120000
  };

  var res = BigQuery.Jobs.query(req, VEKTOR_SAP_PROJECT_ID);
  var jobId = res.jobReference && res.jobReference.jobId ? res.jobReference.jobId : "";

  while (!res.jobComplete) {
    Utilities.sleep(1200);
    res = BigQuery.Jobs.getQueryResults(VEKTOR_SAP_PROJECT_ID, jobId, { timeoutMs: 120000 });
  }

  return res;
}

function vektorSapMapRows_(res) {
  var fields = (((res || {}).schema || {}).fields || []);
  var rows = (res && res.rows) ? res.rows : [];
  if (!fields.length || !rows.length) return [];

  var names = fields.map(function(f){ return String(f.name || ""); });

  return rows.map(function(r){
    var obj = {};
    var vals = (r && r.f) ? r.f : [];
    for (var i = 0; i < names.length; i++) {
      obj[names[i]] = vals[i] ? vals[i].v : null;
    }
    return obj;
  });
}

function getFluxoNumerarioSapMeta() {
  vektorAssertFunctionAllowed_("getFluxoNumerarioSapMeta");

  try {
    var ctx = vektorGetUserRole_();
    var email = String((ctx && ctx.email) || "").trim().toLowerCase();
    if (!email) throw new Error("Não foi possível identificar seu e-mail Google.");

    var allowedLojas = vektorGetAllowedLojasFromEmails_(email); // null admin | array restrito
    var mapLojaTime = construirMapaLojaParaTime_() || {};

    var lojasSet = {};
    var timesSet = {};

    Object.keys(mapLojaTime).forEach(function(lojaNum){
      var loja4 = String(lojaNum).padStart(4, "0");
      var lojaKey = "CE" + loja4;

      if (Array.isArray(allowedLojas) && allowedLojas.indexOf(String(Number(lojaNum))) < 0 && allowedLojas.indexOf(loja4) < 0) {
        return;
      }

      lojasSet[lojaKey] = true;
      var t = String(mapLojaTime[lojaNum] || "").trim();
      if (t) timesSet[t] = true;
    });

    return {
      ok: true,
      lojas: Object.keys(lojasSet).sort(),
      times: Object.keys(timesSet).sort()
    };
  } catch (e) {
    return { ok:false, error: (e && e.message) ? e.message : String(e) };
  }
}

function getFluxoNumerarioSapData(req) {
  vektorAssertFunctionAllowed_("getFluxoNumerarioSapData");

  try {
    req = req || {};

    var ctx = vektorGetUserRole_();
    var email = String((ctx && ctx.email) || "").trim().toLowerCase();
    if (!email) throw new Error("Não foi possível identificar seu e-mail Google.");

    var allowedLojas = vektorGetAllowedLojasFromEmails_(email); // null => admin
    var allowedSet = null;
    if (Array.isArray(allowedLojas)) {
      allowedSet = {};
      allowedLojas.forEach(function(x){
        var n = normalizarLojaNumero_(x);
        if (n == null) return;
        allowedSet[String(n)] = true;
        allowedSet[String(n).padStart(4, "0")] = true;
        allowedSet["CE" + String(n).padStart(4, "0")] = true;
      });
    }

    var dtIni = String(req.dtIni || "").trim();
    var dtFim = String(req.dtFim || "").trim();
    var timeSel = String(req.time || "").trim();
    var lojaSel = String(req.loja || "").trim().toUpperCase();

    if (!dtIni) {
  var hoje = new Date();
  var ini = new Date(hoje);
  ini.setMonth(ini.getMonth() - 3);

  function fmtIso_(d) {
    return Utilities.formatDate(d, Session.getScriptTimeZone() || "America/Sao_Paulo", "yyyy-MM-dd");
  }

  dtIni = fmtIso_(ini);
}

    if (dtFim && dtIni && dtFim < dtIni) {
      throw new Error("Período inválido: a data final não pode ser menor que a inicial.");
    }

    var wherePeriodo = 'bseg.h_budat >= DATE("2025-01-01")';
    if (dtIni && dtFim) {
      wherePeriodo = 'bseg.h_budat BETWEEN DATE("' + dtIni + '") AND DATE("' + dtFim + '")';
    } else if (dtIni) {
      wherePeriodo = 'bseg.h_budat >= DATE("' + dtIni + '")';
    } else if (dtFim) {
      wherePeriodo = 'bseg.h_budat BETWEEN DATE("2025-01-01") AND DATE("' + dtFim + '")';
    }

    var sql = `
    WITH base_filtrada AS (
      SELECT 
        bseg.bukrs   AS Empresa,
        bseg.zuonr   AS Atribuicao,
        bseg.belnr   AS Numdoc,
        bseg.h_blart AS Tipodoc,
        bseg.h_budat AS Datalanc,
        bseg.h_bldat AS Datadoc,
        bseg.bschl   AS CL,
        ABS(
          CASE 
            WHEN bseg.bschl = '50' THEN -bseg.dmbtr 
            ELSE bseg.dmbtr 
          END
        ) AS MontanteValor,
        bkpf.xblnr   AS Referencia,
        bseg.sgtxt   AS Texto,
        bseg.gkont   AS ContaContraPartida,
        bseg.gjahr   AS Exercicio, 
      FROM \`gruposbf-data-lake.trusted.sbf_trd_sap_0000_sap_bseg\` AS bseg
      INNER JOIN \`gruposbf-data-lake.trusted.sbf_trd_sap_0000_sap_bkpf\` AS bkpf
        ON bseg.belnr = bkpf.belnr
      WHERE ${wherePeriodo}
        AND bseg.bukrs = "7010"
        AND bseg.hkont = "1101005003"
        AND bkpf.stblg IS NULL
        AND EXTRACT(YEAR FROM bkpf.cpudt) = EXTRACT(YEAR FROM bseg.h_budat)
        AND bseg.h_blart IN ("DX", "RV", "SG")
    )
    SELECT 
      bf.*,
      CURRENT_DATE("America/Sao_Paulo") AS DataAtualizacao,
      FORMAT_TIMESTAMP('%H:%M:%S', CURRENT_TIMESTAMP(), 'America/Sao_Paulo') AS HoraAtualizacao,
      CONCAT('CE', LPAD(bf.Atribuicao, 4, '0')) AS LocalNegCorreto,
      FORMAT('%.2f', bf.MontanteValor) AS MontanteFmt
    FROM base_filtrada bf
    WHERE bf.Tipodoc = "SG"
    ORDER BY bf.Datalanc DESC
    `;

    var raw = vektorSapMapRows_(vektorSapRunQuery_(sql));
    var mapLojaTime = construirMapaLojaParaTime_() || {};
    var emailMap = {};
    try { emailMap = vektorCarregarMapaEmailsLojas_() || {}; } catch (_) { emailMap = {}; }

    var rows = raw.map(function(r){
      var lojaKey = String(r.LocalNegCorreto || "").trim().toUpperCase();
      var lojaNum = normalizarLojaNumero_(lojaKey);
      var loja4 = lojaNum != null ? String(lojaNum).padStart(4, "0") : "";
      var time = lojaNum != null ? String(mapLojaTime[lojaNum] || "").trim() : "";

      var valor = Math.abs(Number(r.MontanteValor || 0) || 0);
      var dataLancIso = vektorSapParseDateIso_(r.Datalanc);

      return {
        selected: false,
        loja: lojaKey,
        lojaNum: loja4,
        time: time || (emailMap[lojaKey] && emailMap[lojaKey].time ? emailMap[lojaKey].time : ""),
        valor: valor,
        valorFmt: vektorSapFmtMoneyBr_(valor),
        dataLanc: vektorSapFmtDateBr_(r.Datalanc),
        dataLancIso: dataLancIso || "",
        texto: String(r.Texto || "").trim(),
        numdoc: String(r.Numdoc || "").trim(),
        tipodoc: String(r.Tipodoc || "").trim(),
        emailGerente: String((emailMap[lojaKey] && emailMap[lojaKey].emailGerente) || "").trim(),
        emailRegional: String((emailMap[lojaKey] && emailMap[lojaKey].emailRegional) || "").trim()
      };
    });

    rows = rows.filter(function(r){
      if (!r.loja) return false;

      if (allowedSet && !allowedSet[r.loja] && !allowedSet[r.lojaNum] && !allowedSet[String(Number(r.lojaNum) || "")]) {
        return false;
      }

      if (timeSel && r.time !== timeSel) return false;
      if (lojaSel && r.loja !== lojaSel) return false;

      return true;
    });

    rows.sort(function(a,b){
      var da = Date.parse(a.dataLancIso || "") || 0;
      var db = Date.parse(b.dataLancIso || "") || 0;
      return db - da;
    });

    var topLoja = null;
    var maxValor = -Infinity;
    rows.forEach(function(r){
      if (Number(r.valor || 0) > maxValor) {
        maxValor = Number(r.valor || 0);
        topLoja = r;
      }
    });

    var ultima = rows.length ? rows[0] : null;

    var byLoja = {};
    rows.forEach(function(r){
      var k = r.loja || "—";
      byLoja[k] = (byLoja[k] || 0) + (Number(r.valor || 0) || 0);
    });

    var chartRows = Object.keys(byLoja).map(function(loja){
      return { loja: loja, valor: Number(byLoja[loja] || 0) || 0 };
    }).sort(function(a,b){
      return b.valor - a.valor;
    }).slice(0, 20);

    return {
      ok: true,
      rows: rows,
      cards: {
        totalOcorrencias: rows.length,
        lojaMaiorSangria: topLoja ? topLoja.loja : "—",
        lojaMaiorSangriaValor: topLoja ? topLoja.valorFmt : "—",
        ultimaLoja: ultima ? ultima.loja : "—",
        ultimaData: ultima ? ultima.dataLanc : "—",
        ultimaValor: ultima ? ultima.valorFmt : "—"
      },
      chart: chartRows
    };

  } catch (e) {
    return { ok:false, error: (e && e.message) ? e.message : String(e) };
  }
}

function vektorSapMontarTabelaEmail_(rows) {
  var html = "";
  html += "<table style='width:100%; border-collapse:collapse; font-size:12px;'>";
  html += "<thead><tr style='background:#0b1220;color:#fff;'>";
  html += "<th style='padding:8px;border:1px solid #cbd5e1;'>Loja</th>";
  html += "<th style='padding:8px;border:1px solid #cbd5e1;'>Time</th>";
  html += "<th style='padding:8px;border:1px solid #cbd5e1;'>Valor</th>";
  html += "<th style='padding:8px;border:1px solid #cbd5e1;'>Data lançamento</th>";
  html += "<th style='padding:8px;border:1px solid #cbd5e1;'>Texto</th>";
  html += "<th style='padding:8px;border:1px solid #cbd5e1;'>Número documento</th>";
  html += "</tr></thead><tbody>";

  rows.forEach(function(r){
    html += "<tr>";
    html += "<td style='padding:8px;border:1px solid #e2e8f0;'>" + String(r.loja || "—") + "</td>";
    html += "<td style='padding:8px;border:1px solid #e2e8f0;'>" + String(r.time || "—") + "</td>";
    html += "<td style='padding:8px;border:1px solid #e2e8f0; text-align:right;'>" + String(r.valorFmt || "—") + "</td>";
    html += "<td style='padding:8px;border:1px solid #e2e8f0;'>" + String(r.dataLanc || "—") + "</td>";
    html += "<td style='padding:8px;border:1px solid #e2e8f0;'>" + String(r.texto || "—").replace(/</g,"&lt;") + "</td>";
    html += "<td style='padding:8px;border:1px solid #e2e8f0;'>" + String(r.numdoc || "—") + "</td>";
    html += "</tr>";
  });

  html += "</tbody></table>";
  return html;
}

function sendFluxoNumerarioSapNotificacao(rows) {
  vektorAssertFunctionAllowed_("sendFluxoNumerarioSapNotificacao");

  try {
    rows = Array.isArray(rows) ? rows : [];
    if (!rows.length) return { ok:false, error:"Nenhuma linha selecionada." };

    var ctx = vektorGetUserRole_();
    var userEmail = String((ctx && ctx.email) || "").trim().toLowerCase();

    var grupos = {};
    rows.forEach(function(r){
      var lojaKey = String(r.loja || "").trim().toUpperCase();
      if (!lojaKey) return;
      if (!grupos[lojaKey]) grupos[lojaKey] = [];
      grupos[lojaKey].push(r);
    });

    var enviados = [];
    var falhas = [];
    var saudacao = vektorSaudacaoPorHora_();

    Object.keys(grupos).forEach(function(lojaKey){
      var itens = grupos[lojaKey];
      if (!itens.length) return;

      var docKey = vektorSapNormDocKey_(itens[0]);
      if (vektorSapJaFoiNotificado_(docKey)) {
        falhas.push({ loja: lojaKey, error: "Essa ocorrência já foi notificada anteriormente." });
        return;
      }

      var toSet = {};
      function addTo_(em){
        em = String(em || "").trim();
        if (!em) return;
        toSet[em.toLowerCase()] = em;
      }

      addTo_(itens[0].emailGerente);
      addTo_(itens[0].emailRegional);

      var toList = Object.keys(toSet).map(function(k){ return toSet[k]; }).join(",");
      if (!toList) {
        falhas.push({ loja: lojaKey, error: "Sem e-mails de gerente/regional para a loja na aba Emails." });
        return;
      }

      var dataRef = String(itens[0].dataLanc || "").trim() || "—";
      var assunto = "[ALERTA CLARA | SANGRIA] Verificação de Uso - Loja: " + lojaKey + " - " + dataRef;

      var tabela = vektorSapMontarTabelaEmail_(itens);

      var corpo = "";
      corpo += "<div style='font-family:Arial,sans-serif;color:#0f172a;'>";
      corpo += "<p>" + saudacao + "</p>";
      corpo += "<p>Identificamos lançamento(s) de <b>sangria</b> para a loja <b>" + lojaKey + "</b>, mesmo com operação já contemplada pelo cartão Clara.</p>";
      corpo += "<p>Pedimos, por gentileza, a validação e o esclarecimento do motivo do uso desse procedimento.</p>";
      corpo += "<div style='margin:14px 0;'>" + tabela + "</div>";
      corpo += "<p>Atenciosamente,</p>";
      corpo += "<p><b>Vektor - Contas a Receber</b></p>";
      corpo += "</div>";

      try {
        GmailApp.sendEmail(toList, assunto, " ", {
          from: "vektor@gruposbf.com.br",
          name: "Vektor - Grupo SBF",
          cc: VEKTOR_CC_CONTAS_A_RECEBER,
          replyTo: VEKTOR_CC_CONTAS_A_RECEBER,
          htmlBody: corpo
        });

        vektorSapLogNotificacao_({
          userEmail: userEmail,
          docKey: docKey,
          lojaKey: lojaKey,
          time: String(itens[0].time || "").trim(),
          numdoc: String(itens[0].numdoc || "").trim(),
          datalanc: String(itens[0].dataLanc || "").trim(),
          valor: Number(itens[0].valor || 0) || 0,
          to: toList,
          cc: VEKTOR_CC_CONTAS_A_RECEBER,
          status: "SENT",
          error: ""
        });

        enviados.push(lojaKey);

      } catch (eEnv) {
        vektorSapLogNotificacao_({
          userEmail: userEmail,
          docKey: docKey,
          lojaKey: lojaKey,
          time: String(itens[0].time || "").trim(),
          numdoc: String(itens[0].numdoc || "").trim(),
          datalanc: String(itens[0].dataLanc || "").trim(),
          valor: Number(itens[0].valor || 0) || 0,
          to: toList,
          cc: VEKTOR_CC_CONTAS_A_RECEBER,
          status: "ERROR",
          error: (eEnv && eEnv.message) ? eEnv.message : String(eEnv)
        });

        falhas.push({
          loja: lojaKey,
          error: (eEnv && eEnv.message) ? eEnv.message : String(eEnv)
        });
      }
    });

    return {
      ok: enviados.length > 0,
      enviados: enviados,
      falhas: falhas
    };

  } catch (e) {
    return { ok:false, error: (e && e.message) ? e.message : String(e) };
  }
}

 // ==========TESTES===============//

function TESTAR_POLITICA() {
  var txt = vektorPolicyLoadText_();
  Logger.log("Tamanho: " + txt.length);
  Logger.log(txt.substring(0, 1000));
}

function TESTE_MIME_TYPE() {
  var file = DriveApp.getFileById("1Lj4i5he1kKDSBbXJSwyw51SszCYu8KOB");
  Logger.log("Nome: " + file.getName());
  Logger.log("MimeType: " + file.getMimeType());
}

function debug_sendAs() {
  Logger.log(Session.getActiveUser().getEmail()); // quem está executando
  Logger.log(GmailApp.getAliases());              // aliases disponíveis p/ envio nessa conta

  GmailApp.sendEmail("rodrigo.lisboa@gruposbf.com.br", "Teste sendAs", " ", {
    htmlBody: "<b>teste</b>",
    from: "vektor@gruposbf.com.br",
    name: "Vektor"
  });
}
