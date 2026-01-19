// =======================
// CONFIGURAÇÕES GLOBAIS
// =======================

var BASE_CLARA_ID = "1_XW0IqbYjiCPpqtwdEi1xPxDlIP2MSkMrLGbeinLIeI"; // ID real da planilha BaseClara
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

//function isAdminEmail(email) {
  //if (!email) return false;
  //email = email.toLowerCase();

  //var ADM_EMAILS = [
   // "rodrigo.lisboa@gruposbf.com.br",
    //"tainara.nascimento@gruposbf.com.br"
  //];

  //return ADM_EMAILS.indexOf(email) !== -1;
//}

// =======================
// VEKTOR - CONTROLE DE ACESSO (WHITELIST) -- LIBERAR ACESSO AQUI!!!!!!!!!!!!!
// =======================

// ✅ Lista de e-mails autorizados a usar o Vektor (whitelist)
var VEKTOR_WHITELIST_EMAILS = [
  "rodrigo.lisboa@gruposbf.com.br",
  "tainara.nascimento@gruposbf.com.br",
  "durval.neto@centauro.com.br",
  "gabriela.crochiquia@centauro.com.br"
  // adicione outros aqui
];

function isWhitelistedEmail_(email) {
  if (!email) return false;
  var e = String(email).trim().toLowerCase();
  return VEKTOR_WHITELIST_EMAILS.map(function(x){ return String(x).trim().toLowerCase(); }).indexOf(e) !== -1;
}

// (recomendado) Use este "porteiro" no começo das funções expostas via google.script.run
function vektorAssertWhitelisted_() {
  var sess = (Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  if (!sess) throw new Error("Não foi possível identificar seu e-mail Google.");
  if (!isWhitelistedEmail_(sess)) throw new Error("Acesso negado: usuário não habilitado no Vektor.");
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
  var ctx = vektorGetUserRole_();
} catch (e) {
  var sessDebug = (Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  var emailsMap = null;
  var recDebug = null;

  try {
    emailsMap = vektorLoadEmailsRoleMap_();
    recDebug = (emailsMap && emailsMap.byEmail) ? emailsMap.byEmail[sessDebug] : null;
  } catch (e2) {
    return { ok: false, error: "RBAC DEBUG: sess=" + sessDebug + " | loadEmailsRoleMap falhou: " + (e2 && e2.message ? e2.message : e2) };
  }

  return {
    ok: false,
    error:
      "RBAC DEBUG: sess=" + sessDebug +
      " | rec=" + JSON.stringify(recDebug) +
      " | err=" + (e && e.message ? e.message : e)
  };
}

  var token = vektorCreateSessionToken_(sess);
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
var VEKTOR_ACL_CACHE_TTL = 300; // 5 min

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
  var sess = vektorAssertWhitelisted_(); // mantém whitelist intacta :contentReference[oaicite:4]{index=4}
  var emails = vektorLoadEmailsRoleMap_();
  var rec = emails.byEmail[vektorNormEmail_(sess)];

  if (!rec || !rec.ativo) {
    throw new Error("Não disponível para o seu perfil.");
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
var VEKTOR_SESSION_TTL_SECONDS = 5 * 60; // 3 horas ou 5 minutos

function vektorCreateSessionToken_(email) {
  // token aleatório + carimbo
  var token = Utilities.getUuid() + "-" + new Date().getTime();
  var cache = CacheService.getScriptCache();

  // Armazena no cache: token -> email
  cache.put("VEKTOR_SESSION_" + token, String(email || ""), VEKTOR_SESSION_TTL_SECONDS);
  return token;
}

function vektorValidateSessionToken_(token) {
  var t = String(token || "").trim();
  if (!t) return { ok: false, error: "Token vazio." };

  var emailSessao = (Session.getActiveUser().getEmail() || "").trim().toLowerCase();
  if (!emailSessao) return { ok: false, error: "Não foi possível identificar seu e-mail Google." };

  // whitelist continua sendo a fonte de verdade
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

  return { ok: true, email: emailSessao };
}

function validarSessaoVektor(token) {
  return vektorValidateSessionToken_(token);
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
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ✅ ID da planilha de métricas do Vektor
// (a planilha que você mandou)
const VEKTOR_METRICAS_SHEET_ID = '18yAuYoAR33JOagqapxgwHh86F1WeD0mZcj9AIJym07k';

// ✅ Nome da aba onde os logs serão gravados
const VEKTOR_METRICAS_TAB_NAME = 'Vektor_Metricas';

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
        AND dat_afastamento_date >= DATE_TRUNC(CURRENT_DATE("America/Sao_Paulo"), MONTH)
        AND dat_afastamento_date < DATE_ADD(DATE_TRUNC(CURRENT_DATE("America/Sao_Paulo"), MONTH), INTERVAL 1 MONTH)
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
      dataTransFormat = Utilities.formatDate(dataTransBruta, tz, "dd/MM/yyyy");
    } else {
      dataTransFormat = dataTransBruta;
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

var emailDestino = String(
  (payload && payload.emailDestino) ? payload.emailDestino : emailUsuario
).trim();

// 🔒 trava domínio
var emailRegex = /^[^\s@]+@((gruposbf|centauro)\.com\.br)$/i;
if (!emailRegex.test(emailDestino)) {
  return {
    ok: false,
    error: "Informe um e-mail válido dos domínios @gruposbf.com.br ou @centauro.com.br."
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

    MailApp.sendEmail({
      to: emailDestino,
      cc: "rodrigo.lisboa@gruposbf.com.br",//"contasareceber@gruposbf.com.br",
      subject: assunto,
      replyto: "contasareceber@gruposbf.com.br",
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
        dataTransFormat = Utilities.formatDate(dataTransBruta, tz, "dd/MM/yyyy");
      } else {
        dataTransFormat = dataTransBruta;
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

// Converte o valor da coluna "Data da Transação" para Date
function parseDateClara_(value) {
  if (!value) return null;
  if (Object.prototype.toString.call(value) === "[object Date]") {
    return value;
  }
  if (typeof value === "string") {
    // tenta formato dd/MM/yyyy
    var parts = value.split("/");
    if (parts.length === 3) {
      var d = parseInt(parts[0], 10);
      var m = parseInt(parts[1], 10) - 1;
      var y = parseInt(parts[2], 10);
      return new Date(y, m, d);
    }
  }
  return null;
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

// Procura o índice de uma coluna no cabeçalho da BaseClara
// usando uma lista de possíveis nomes (variações de texto).
function encontrarIndiceColuna_(header, possiveisNomes) {
  if (!header || !header.length) return -1;

  if (!Array.isArray(possiveisNomes)) {
    possiveisNomes = [possiveisNomes];
  }

  // normaliza os nomes que queremos achar
  var nomesNorm = possiveisNomes.map(function (nome) {
    return normalizarTexto_(nome);
  });

  for (var i = 0; i < header.length; i++) {
    var hNorm = normalizarTexto_(header[i]);
    if (!hNorm) continue;

    for (var j = 0; j < nomesNorm.length; j++) {
      var alvo = nomesNorm[j];
      if (!alvo) continue;

      // bate se for igual ou se um contém o outro
      if (hNorm === alvo ||
          hNorm.indexOf(alvo) !== -1 ||
          alvo.indexOf(hNorm) !== -1) {
        return i;
      }
    }
  }

  return -1; // não encontrou
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
    "papel couche", "couche", "laminacao", "recorte", "vinil", "bobina",

    // Água / consumo básico
    "agua", "agua potavel", "potavel", "agua mineral", "galao", "garrafa",

    // Lanches / apoio operacional
    "lanche", "lanches", "coffee", "cafe", "cafezinho", "snack",

    // Materiais de escritório (comuns)
    "caneta", "lapis", "borracha", "apontador", "marcador", "pilot", "pincel",
    "papel a4", "papel sulfite", "sulfite", "pasta", "arquivo", "etiqueta", "etiquetas",
    "grampo", "grampeador", "clipes", "cola", "fita adesiva", "tesoura",

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
    "presente", "gift"
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


  rows.push({
    data: dataBr,
    valor: valor,
    loja: lojaOut,
    time: timeOut,
    item: itemRaw,
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
    var agg = {}; // key -> { cartaoKey, nomeCartao, loja, time, usado }

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
          usado: 0
        };
      }
      agg[key].usado += v;
    }

    // --- 4.2) Se não houve transação no ciclo, ainda assim queremos mostrar saldos (usado=0)
// para cartões que já têm vínculo loja/time (histórico).
Object.keys(vinculoPorCartao).forEach(function(cartaoKey) {
  var v = vinculoPorCartao[cartaoKey];
  if (!v) return;

  // Se for filtro por time, respeita
  if (tipoFiltro === "time") {
    var filtroTimeNorm = normalizarTexto_(valorFiltro);
    var vNorm = normalizarTexto_(v.time);
    if (!filtroTimeNorm || !vNorm || vNorm.indexOf(filtroTimeNorm) === -1) return;
  }

  // mesma “chave” da sua agregação:
  // - geral inclui time no key (para não colapsar times diferentes)
  // - time (ou loja, se existir) usa key menor
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
      usado: 0
    };
  }
});

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
      var hoje = new Date();

      // normaliza datas para evitar erro por horário
      var hoje0 = new Date(hoje.getFullYear(), hoje.getMonth(), hoje.getDate());
      var ini0  = new Date(ini.getFullYear(), ini.getMonth(), ini.getDate());

      // fim do ciclo é sempre dia 05 (mês corrente se hoje<=05; senão próximo mês)
      var dHoje = hoje0.getDate();
      var fimCiclo = (dHoje >= 6)
        ? new Date(hoje0.getFullYear(), hoje0.getMonth() + 1, 5)
        : new Date(hoje0.getFullYear(), hoje0.getMonth(), 5);

      var msDia = 24*60*60*1000;
      var diasDecorridos = Math.max(1, Math.floor((hoje0.getTime() - ini0.getTime()) / msDia) + 1);
      var diasRestantes  = Math.max(0, Math.ceil((fimCiclo.getTime() - hoje0.getTime()) / msDia));

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
      var pctUsoLim = (limiteAtual > 0) ? (usado / limiteAtual) : null;
      var ritmoRatio = (pctUsoLim !== null && pctCiclo > 0) ? (pctUsoLim / pctCiclo) : null;

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



      // --- trava de redução por tempo do ciclo ---
      var hoje = new Date();
      var pc = getPeriodoCicloClara_();
      var ini = pc.inicio, fim = pc.fim;
      var msDia = 24 * 60 * 60 * 1000;

      var diasTot = Math.max(1, Math.round((fim.getTime() - ini.getTime()) / msDia) + 1);
      var diasDec = Math.max(1, Math.floor((hoje.getTime() - ini.getTime()) / msDia) + 1);
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

      var deltaRisco = alvo - limiteAtual;
      if (deltaRisco > 0) {
        acao = "Aumentar +" + moneyBR_(deltaRisco);

        // opcional (mas recomendado): alinhar o limiteRec com o alvo para consistência
        limiteRec = alvo;
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
  var email = Session.getActiveUser().getEmail();
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

  MailApp.sendEmail({
    to: destinatarios.join(","),
    subject: assunto,
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
    var dtSnap = (r[0] instanceof Date) ? r[0] : new Date(r[0]);
    if (!(dtSnap instanceof Date) || isNaN(dtSnap.getTime())) return;
    if (dtSnap < inicio) return;

    var lojaKey = getLojaKey(r[1]);
    var lojaNum = normalizarLojaNumero_(lojaKey);

    var dtTx = (r[2] instanceof Date) ? r[2] : new Date(r[2]);

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
      classificacao: classificacao
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
  html += "<b>Janela:</b> últimos " + esc(meta.diasJanela || "") + " dias";
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

    var assunto = "📌 [ALERTA CLARA | JUSTIFICATIVAS] Lojas ofensoras (" +
  ((rel && rel.meta && rel.meta.diasJanela) ? rel.meta.diasJanela : (diasJanela||60)) + "d)";

    if (rel && rel.periodo && rel.periodo.inicio && rel.periodo.fim) {
      assunto += " | " + rel.periodo.inicio + " a " + rel.periodo.fim;
    }

var html = montarEmailOfensorasPendencias_(rel);


    MailApp.sendEmail({
      to: destinatarios.join(","),
      subject: assunto,
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
  var AGUARDAR_MIN = 12; // ajuste aqui (10–20 costuma ser bom)
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
  vektorAssertFunctionAllowed_("getPossivelUsoIrregularParaChat");
  try {
    var email = Session.getActiveUser().getEmail();
    if (!isAdminEmail(email)) {
      return { ok: false, restrito: true, error: "Acesso restrito: apenas Administrador." };
    }

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
    var email = Session.getActiveUser().getEmail();
    if (!isAdminEmail(email)) {
      return { ok: false, restrito: true, error: "Acesso restrito: apenas Administrador." };
    }

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
  try {
    // Segurança: apenas Admin pode disparar manualmente também
    var email = Session.getActiveUser().getEmail();
    if (email && !isAdminEmail(email)) {
      return { ok: false, error: "Acesso restrito: apenas Administrador." };
    }

    var destinatarios = getAdminEmails_(); // já existe no seu arquivo :contentReference[oaicite:4]{index=4}
    if (!destinatarios || !destinatarios.length) {
      return { ok: false, error: "Lista de admins vazia." };
    }

    var assunto = "📌 [ALERTA CLARA | POSSÍVEL USO IRREGULAR] " +
      (rel.meta && rel.meta.periodo ? ("| " + rel.meta.periodo) : "");

    // Tabela HTML (resumo)
    var html = montarEmailUsoIrregular_(rel);

    MailApp.sendEmail({
      to: destinatarios.join(","),
      subject: assunto,
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
          "Critérios podem incluir fracionamento, pendência + valor alto, recorrência anormal por estabelecimento/cartão." +
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

    var v = parseNumberSafe_(r[IDX_VALOR]);
    if (!isFinite(v) || v <= 0) continue;

    valoresJanela.push(v);

    // chave dia (dd/MM/yyyy) + cartao + estab
    var dataKey = Utilities.formatDate(d, tz, "dd/MM/yyyy");
    var kDia = dataKey + "||" + cartao + "||" + norm_(estab);

    if (!gruposDia[kDia]) {
      gruposDia[kDia] = {
        loja: loja, time: time, dataKey: dataKey, cartao: cartao, estab: estab,
        qtd: 0, soma: 0, maxValor: 0, pendCount: 0
      };
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

/**
 * Calcula uma assinatura leve da BaseClara, suficiente para detectar atualização real da aba.
 * Estratégia:
 * - usa header para localizar colunas Data/Valor (robusto)
 * - usa Alias fixo (col H) se você quiser; aqui mantive robusto por nome também
 * - lê só as últimas N linhas das 3 colunas (Alias, Data, Valor) e faz hash
 * - inclui maxData e lastRow para reforçar detecção
 */
function calcularAssinaturaBaseClara_() {
  try {
    var info = carregarLinhasBaseClara_(); // seu helper
    if (info.error) return { error: info.error };

    var header = info.header || [];
    var linhas = info.linhas || [];
    var lastRow = linhas.length;
    if (lastRow <= 0) return { sig: "EMPTY", maxDataMs: 0, lastRow: 0 };

    // Índices: Alias fixo e Data/Valor por nome (para não confundir com "Cartão")
    // Alias Do Cartão = coluna H => índice 7 (0-based) (você já validou que isso é fixo)
    var idxAlias = 7;

    // Data/Valor: mantém por nome para suportar variações de header
    var idxValor = encontrarIndiceColuna_(header, ["Valor em R$", "Valor (R$)", "Valor", "Total"]);
    var idxData  = encontrarIndiceColuna_(header, ["Data da Transação", "Data Transação", "Data"]);

    if (idxValor < 0) return { error: "Não encontrei a coluna de Valor na BaseClara para assinatura." };
    if (idxData < 0)  return { error: "Não encontrei a coluna de Data na BaseClara para assinatura." };

    // lê só as últimas N linhas para a assinatura (evita custo alto)
    var N = 250; // ajuste se quiser (200–500 geralmente ok)
    var start = Math.max(0, lastRow - N);

    // calcula maxData do conjunto total (não só das últimas N)
    // (se preferir leve, pode calcular só em N; mas maxData total é mais “forte”)
    var maxDataMs = 0;
    for (var i = 0; i < lastRow; i++) {
      var dt = linhas[i][idxData];
      var d = (dt instanceof Date) ? dt : new Date(dt);
      if (d instanceof Date && !isNaN(d.getTime())) {
        var ms = d.getTime();
        if (ms > maxDataMs) maxDataMs = ms;
      }
    }

    // monta um payload determinístico das últimas N linhas usando só Alias/Data/Valor
    var parts = [];
    for (var j = start; j < lastRow; j++) {
      var r = linhas[j];

      var alias = (r[idxAlias] || "").toString().trim();
      var v = r[idxValor];
      var dt2 = r[idxData];

      var d2 = (dt2 instanceof Date) ? dt2 : new Date(dt2);
      var d2s = (d2 instanceof Date && !isNaN(d2.getTime())) ? d2.toISOString().slice(0, 10) : "";

      // valor como string estável
      var vs = (typeof v === "number") ? v.toFixed(2) : (v || "").toString().trim();

      parts.push(alias + "|" + d2s + "|" + vs);
    }

    var payload = "LR=" + lastRow + ";MAX=" + maxDataMs + ";DATA=" + parts.join("||");
    var digest = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, payload, Utilities.Charset.UTF_8);
    var sig = digest.map(function(b) {
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
      var dt2 = (dt instanceof Date) ? dt : new Date(dt);

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
    return { ok: false, error: e && e.message ? e.message : e };
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
    return { ok: false, error: e && e.message ? e.message : e };
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

    MailApp.sendEmail({
      to: emailDestino,
      subject: `Resumo de transações | ${resumo.grupo}`,
      htmlBody: corpo,
      name: "Assistente Vektor"
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
    return { ok: false, error: e && e.message ? e.message : e };
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
    return { ok: false, error: e && e.message ? e.message : e };
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

    // 1) Índice da coluna do código da loja (continua dinâmico)
    var idxLoja = encontrarIndiceColuna_(header, [
      "LojaNum", "Loja Num", "Loja Número", "Loja Numero", "Loja", "cod_loja", "codLoja"
    ]);

    if (idxLoja < 0) {
      return [];
    }

    // 2) Índice da coluna de "Descrição Loja" — coluna W
    // aqui vamos ser bem específicos para NÃO pegar "Descrição" da coluna U
    var idxNome = header.indexOf("Descrição Loja");
    if (idxNome < 0) {
      idxNome = header.indexOf("Descricao Loja"); // fallback sem acento, se for o caso
    }

    // Se mesmo assim não achar, melhor retornar só código
    var temNome = idxNome >= 0;

    var mapa = {};

    linhas.forEach(function (row) {
      var codRaw = (row[idxLoja] || "").toString().trim();
      if (!codRaw) return;

      var digits = codRaw.replace(/\D/g, "");
      if (!digits) return;

      var cod4 = ("0000" + digits).slice(-4);

      var nome = "";
      if (temNome) {
        nome = (row[idxNome] || "").toString().trim();
      }

      mapa[cod4] = nome;
    });

    var out = [];
    for (var c in mapa) {
      if (Object.prototype.hasOwnProperty.call(mapa, c)) {
        out.push({
          codigo: c,
          nome: mapa[c]   // <- agora vem especificamente da coluna W
        });
      }
    }

    out.sort(function (a, b) {
      return a.codigo.localeCompare(b.codigo);
    });

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
    var emailRegex = /^[^\s@]+@((gruposbf|centauro)\.com\.br)$/i;
    if (!emailRegex.test(emailDestino)) {
      return { ok: false, error: "E-mail inválido. Use apenas @gruposbf.com.br ou @centauro.com.br." };
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

      MailApp.sendEmail({
        to: "rodrigo.lisboa@gruposbf.com.br",
        subject: assunto,
        body: corpo,
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
  diasJanela = Number(diasJanela) || 60;

  const rel = gerarRelatorioOfensorasPendencias_(diasJanela);
  if (!rel || !rel.ok) {
    return { ok: false, error: "Falha ao gerar relatório." };
  }

  // período
  var hoje = new Date();
  var inicio = new Date(hoje.getTime() - diasJanela * 24 * 60 * 60 * 1000);
  var tz = "America/Sao_Paulo";

  var periodo = {
    inicio: Utilities.formatDate(inicio, tz, "dd/MM/yyyy"),
    fim: Utilities.formatDate(hoje, tz, "dd/MM/yyyy")
  };

  return {
    ok: true,
    periodo: periodo,
    meta: { diasJanela: diasJanela, totalLojas: (rel.rows || []).length },
    rows: (rel.rows || []).map(r => {
      const t14 = r.trend14 || {};

      // ✅ delta absoluto (compatível com versões antigas)
      const deltaAbs = (t14.deltaAbs != null) ? t14.deltaAbs : (t14.delta != null ? t14.delta : 0);

      return {
        loja: r.loja,
        time: r.time || "N/D",

        qtde: r.qtde,
        valor: r.valor,
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

function getComparativoFaturasClaraParaChat() {
  vektorAssertFunctionAllowed_("getComparativoFaturasClaraParaChat");
  try {
    var info = carregarLinhasBaseClara_();
    if (info.error) return { ok: false, error: info.error };

    var header = info.header || [];
    var linhas = info.linhas || [];

    // ===== helper: header exato (evita confundir "Transação" com "Data da transação")
    function findHeaderExact_(headerArr, label) {
      var alvo = String(label || "").trim().toLowerCase();
      for (var i = 0; i < headerArr.length; i++) {
        var h = String(headerArr[i] || "").trim().toLowerCase();
        if (h === alvo) return i;
      }
      return -1;
    }

    // ===== Índices (robusto por nome)
    var idxExtrato  = encontrarIndiceColuna_(header, ["Extrato da conta", "Extrato conta", "Extrato"]);
    var idxValor    = encontrarIndiceColuna_(header, ["Valor em R$", "Valor (R$)", "Valor"]);
    var idxLojaNum  = encontrarIndiceColuna_(header, ["LojaNum", "Loja Num", "Loja", "Cod Loja", "Código Loja"]);
    var idxTime     = encontrarIndiceColuna_(header, ["Grupos", "Grupo", "Time"]);
    var idxCategoria= encontrarIndiceColuna_(header, ["Categoria", "Etiqueta", "Tipo de gasto", "Tag"]);

    var idxEstab = findHeaderExact_(header, "Transação");          // EXATO
    var idxData  = findHeaderExact_(header, "Data da transação");  // EXATO

    if (idxExtrato < 0 || idxValor < 0 || idxLojaNum < 0) {
      return { ok: false, error: "Não encontrei as colunas mínimas ('Extrato da conta', 'Valor', 'LojaNum') na BaseClara." };
    }

    // ===== 1) Descobrir as 2 últimas faturas (por Extrato)
    var mapaExtratos = {}; // extrato -> {extrato, inicio, fim}
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

    extratos.sort(function(a,b){
      var da = a.inicio ? a.inicio.getTime() : (a.fim ? a.fim.getTime() : 0);
      var db = b.inicio ? b.inicio.getTime() : (b.fim ? b.fim.getTime() : 0);
      return da - db;
    });

    if (extratos.length < 2) {
      return { ok: false, error: "Não há faturas suficientes para comparação (preciso de pelo menos 2 extratos)." };
    }

    var fatAnterior = extratos[extratos.length - 2];
    var fatAtual    = extratos[extratos.length - 1];

    var extratoAnterior = fatAnterior.extrato;
    var extratoAtual    = fatAtual.extrato;

    var tz = "America/Sao_Paulo";

    // ===== Período base (vai ser ajustado pelo recorte)
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

    // ===== Recorte: comparar o mesmo intervalo do ciclo (até "hoje" no ciclo atual)
    var hoje = new Date();
    hoje = new Date(Utilities.formatDate(hoje, tz, "yyyy/MM/dd") + " 00:00:00"); // zera hora

    var inicioAtual = fatAtual.inicio;
    var fimAtual    = fatAtual.fim;
    var inicioAnterior = fatAnterior.inicio;
    var fimAnterior    = fatAnterior.fim;

    var usarRecorte = !!(inicioAtual && fimAtual && inicioAnterior && fimAnterior && idxData >= 0);

    var fimRecorteAtual = null;
    var fimRecorteAnterior = null;

    if (usarRecorte) {
      // fim do recorte do atual = min(hoje, fimAtual)
      fimRecorteAtual = (hoje.getTime() < fimAtual.getTime()) ? hoje : fimAtual;

      // ===== Sazonalidade (varejo): detecta eventos próximos ao recorte atual
        function parseBRDate_(s) {
          // "dd/MM/yyyy" -> Date (00:00 no tz)
          var parts = String(s || "").split("/");
          if (parts.length !== 3) return null;
          var dd = Number(parts[0]), mm = Number(parts[1]), yy = Number(parts[2]);
          if (!dd || !mm || !yy) return null;
          return new Date(yy, mm - 1, dd, 0, 0, 0, 0);
        }

        function lastFridayOfNovember_(year) {
          // Black Friday: última sexta-feira de novembro (regra prática)
          var d = new Date(year, 10, 30, 0, 0, 0, 0); // 30/11
          while (d.getDay() !== 5) d = new Date(d.getTime() - 24*60*60*1000); // 5=sexta
          return d;
        }

        function secondSunday_(year, monthIndex0) {
          // 2º domingo de um mês (monthIndex0: 0=jan)
          var d = new Date(year, monthIndex0, 1, 0, 0, 0, 0);
          while (d.getDay() !== 0) d = new Date(d.getTime() + 24*60*60*1000); // 0=domingo
          // primeiro domingo encontrado; soma 7 dias -> segundo
          return new Date(d.getTime() + 7*24*60*60*1000);
        }

        function withinWindow_(date, start, end) {
          if (!date || !start || !end) return false;
          return date.getTime() >= start.getTime() && date.getTime() <= end.getTime();
        }

        function addDays_(d, n) {
          return new Date(d.getTime() + n * 24*60*60*1000);
        }

        function detectRetailEvents_(startDate, endDate) {
          // Retorna eventos relevantes dentro do intervalo [startDate, endDate]
          var events = [];
          var y = startDate.getFullYear();
          var y2 = endDate.getFullYear();

          // Para intervalos que cruzam ano, checa os 2 anos
          for (var year = y; year <= y2; year++) {
            // Black Friday (janela: semana do evento)
            var bf = lastFridayOfNovember_(year);
            var bfStart = addDays_(bf, -3); // terça
            var bfEnd   = addDays_(bf, +3); // segunda
            events.push({ nome: "Black Friday", start: bfStart, end: bfEnd });

            // Natal (janela: 20/12 a 26/12)
            events.push({ nome: "Natal", start: new Date(year, 11, 20), end: new Date(year, 11, 26) });

            // Ano Novo (janela: 28/12 a 02/01)
            events.push({ nome: "Ano Novo", start: new Date(year, 11, 28), end: new Date(year + 1, 0, 2) });

            // Dia das Mães (2º domingo de maio) — janela: semana do evento
            var maes = secondSunday_(year, 4); // maio
            events.push({ nome: "Dia das Mães", start: addDays_(maes, -3), end: addDays_(maes, +3) });

            // Dia dos Pais (2º domingo de agosto) — janela: semana do evento
            var pais = secondSunday_(year, 7); // agosto
            events.push({ nome: "Dia dos Pais", start: addDays_(pais, -3), end: addDays_(pais, +3) });

            // Dia das Crianças (12/10) — janela: semana
            events.push({ nome: "Dia das Crianças", start: new Date(year, 9, 9), end: new Date(year, 9, 15) });

            // Dia dos Namorados (12/06 no BR) — janela: semana
            events.push({ nome: "Dia dos Namorados", start: new Date(year, 5, 9), end: new Date(year, 5, 15) });
          }

          // Filtra só os que interceptam o intervalo
          var hit = [];
          for (var i = 0; i < events.length; i++) {
            var e = events[i];
            var intersects = !(e.end.getTime() < startDate.getTime() || e.start.getTime() > endDate.getTime());
            if (intersects) hit.push(e.nome);
          }

          // remove duplicados
          var seen = {};
          return hit.filter(function(n){
            if (seen[n]) return false;
            seen[n] = true;
            return true;
          });
        }

        // Intervalo real do recorte atual (datas Date)
        var recorteAtualInicio = inicioAtual;
        var recorteAtualFim = usarRecorte ? fimRecorteAtual : fimAtual;

        var eventosSazonais = [];
        if (recorteAtualInicio && recorteAtualFim) {
          eventosSazonais = detectRetailEvents_(recorteAtualInicio, recorteAtualFim);
        }

      // dias decorridos desde o início do ciclo atual (0 = mesmo dia do início)
      var msDia = 24 * 60 * 60 * 1000;
      var diasDecorridos = Math.floor((fimRecorteAtual.getTime() - inicioAtual.getTime()) / msDia);

      // fim do recorte anterior = inicioAnterior + mesmos dias decorridos
      fimRecorteAnterior = new Date(inicioAnterior.getTime() + diasDecorridos * msDia);

      // trava: não ultrapassar o fim real do ciclo anterior
      if (fimRecorteAnterior.getTime() > fimAnterior.getTime()) fimRecorteAnterior = fimAnterior;

      // atualiza período mostrado no chat
      periodo.atual.fim = Utilities.formatDate(fimRecorteAtual, tz, "dd/MM/yyyy");
      periodo.anterior.fim = Utilities.formatDate(fimRecorteAnterior, tz, "dd/MM/yyyy");
    }

    // ===== Mapa fallback Loja -> Time (se vier vazio na linha)
    var mapaTime = construirMapaLojaParaTime_();

    // ===== 2) Agregação por loja
    var stats = {}; // loja -> objeto stats

    // agregado geral por dia (para topDias no resumo)
    var dayPrevGeral = {}; // "dd/MM/yyyy" -> valor
    var dayCurGeral  = {}; // "dd/MM/yyyy" -> valor

    var ultimaDataConsiderada = null; // Date (maior data que entrou no recorte atual)

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

      // ===== Aplica recorte por data (mesmo intervalo do ciclo)
      var dtRow = null;
      if (usarRecorte) {
        dtRow = row[idxData] instanceof Date ? row[idxData] : new Date(row[idxData]);
        if (!(dtRow instanceof Date) || isNaN(dtRow.getTime())) continue;

        if (usarRecorte && ex2 === extratoAtual && dtRow && dtRow instanceof Date && !isNaN(dtRow.getTime())) {
        if (!ultimaDataConsiderada || dtRow.getTime() > ultimaDataConsiderada.getTime()) {
          ultimaDataConsiderada = dtRow;
        }
      }

        // zera hora para comparar por dia
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

      // Time
      var timeLinha = (idxTime >= 0) ? str(row[idxTime]) : "";
      if (!st.time) st.time = timeLinha || (mapaTime[loja] || "N/D");

      var valor = valNum(row[idxValor]);
      var cat = (idxCategoria >= 0 ? str(row[idxCategoria]) : "") || "Sem categoria";

      // Estabelecimento: se for Date por algum motivo, converte para string (evita aparecer GMT)
      var estabRaw = (idxEstab >= 0 ? row[idxEstab] : "");
      var estab = "Sem estabelecimento";
      if (idxEstab >= 0) {
        if (estabRaw instanceof Date && !isNaN(estabRaw.getTime())) {
          estab = Utilities.formatDate(estabRaw, tz, "dd/MM/yyyy");
        } else {
          estab = str(estabRaw) || "Sem estabelecimento";
        }
      }

      // dia (se tiver dtRow do recorte, usa; senão tenta derivar)
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
    } // fim loop linhas

    // ===== 3) Montar rows + fator variação
    var rows = [];
    var lojas = Object.keys(stats);

    function pickDriverCategory(st, delta){
      // delta por categoria = cur - prev
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
      // Escolhe o estab com maior delta na direção do deltaAbs, mas só se existia antes e explica >= 30% do delta
      var deltas = {};
      Object.keys(st.estabPrev || {}).forEach(function(e){ deltas[e] = (deltas[e] || 0) - st.estabPrev[e]; });
      Object.keys(st.estabCur  || {}).forEach(function(e){ deltas[e] = (deltas[e] || 0) + st.estabCur[e]; });

      var best = null;
      var bestD = (deltaAbs >= 0 ? -1e18 : 1e18);

      Object.keys(deltas).forEach(function(e){
        if (!(st.estabPrev[e] > 0)) return; // só considera se existia no período anterior

        var d = deltas[e] || 0;
        if (deltaAbs >= 0) {
          if (d > bestD) { bestD = d; best = e; }
        } else {
          if (d < bestD) { bestD = d; best = e; }
        }
      });

      if (!best) return null;

      var share = (Math.abs(deltaAbs) > 0) ? (Math.abs(bestD) / Math.abs(deltaAbs)) : 0;
      if (share < 0.30) return null; // condicional: não explica o suficiente

      return { estab: best, deltaEstab: bestD, share: share };
    }

    lojas.forEach(function(k){
      var st = stats[k];
      var prev = st.prev || 0;
      var cur  = st.cur || 0;
      var delta = cur - prev;

      // var%
      var varPct = null;
      var varPctTxt = "";
      if (prev > 0) {
        varPct = (delta / prev) * 100;
        varPctTxt = (varPct > 0 ? "+" : "") + varPct.toFixed(1) + "%";
      } else {
        varPctTxt = (cur > 0 ? "Início no período" : "—");
      }

      // driver categoria
      var dCat = pickDriverCategory(st, delta);

      // pico
      var pico = pickPeakDay(st);

      // estab condicional
      var estabInfo = pickEstabCondicional(st, delta);

      // fator variação (texto)
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
        variacaoPctNum: varPct, // pode ser null
        categoriaDriver: dCat.cat,
        picoDia: pico.day,
        picoValor: pico.value,
        fatorVariacao: fator
      });
    });

    // Ordena: maiores aumentos em R$ primeiro
    rows.sort(function(a,b){
      return (b.deltaValor || 0) - (a.deltaValor || 0);
    });

    // ===== Totais geral
    var totalPrev = 0, totalCur = 0;
    rows.forEach(function(r){
      totalPrev += Number(r.valorAnterior) || 0;
      totalCur  += Number(r.valorAtual) || 0;
    });
    var totalDelta = totalCur - totalPrev;
    var totalVarPctTxt = (totalPrev > 0)
      ? ((totalDelta/totalPrev*100 > 0 ? "+" : "") + (totalDelta/totalPrev*100).toFixed(1) + "%")
      : (totalCur > 0 ? "Início no período" : "—");

    // ===== Top categorias (delta geral)
    var deltaCatGeral = {}; // cat -> delta
    Object.keys(stats).forEach(function(loja){
      var st = stats[loja];
      Object.keys(st.catPrev || {}).forEach(function(c){ deltaCatGeral[c] = (deltaCatGeral[c] || 0) - st.catPrev[c]; });
      Object.keys(st.catCur  || {}).forEach(function(c){ deltaCatGeral[c] = (deltaCatGeral[c] || 0) + st.catCur[c]; });
    });

    var topCats = Object.keys(deltaCatGeral).map(function(c){
      return { categoria: c, delta: deltaCatGeral[c] || 0 };
    }).sort(function(a,b){
      return Math.abs(b.delta) - Math.abs(a.delta);
    }).slice(0,3);

    // ===== Top lojas contribuição (delta)
    var topLojas = rows.slice().sort(function(a,b){
      return Math.abs(b.deltaValor||0) - Math.abs(a.deltaValor||0);
    }).slice(0,5);

    // ===== Top dias (delta diário = cur - prev) — para "Leitura do período"
    var diasSet = {};
    Object.keys(dayPrevGeral).forEach(function(d){ diasSet[d] = true; });
    Object.keys(dayCurGeral ).forEach(function(d){ diasSet[d] = true; });

    var topDias = Object.keys(diasSet).map(function(d){
      var vPrev = dayPrevGeral[d] || 0;
      var vCur  = dayCurGeral[d]  || 0;
      return { dia: d, prev: vPrev, cur: vCur, delta: (vCur - vPrev) };
    }).sort(function(a,b){
      return Math.abs(b.delta) - Math.abs(a.delta);
    }).slice(0,5);

    // Insights (se você ainda usa em algum lugar)
    var top = rows.filter(function(r){ return (r.deltaValor || 0) > 0; }).slice(0,5);
    var insights = top.map(function(r){
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
      meta: { extratoAtual: extratoAtual, extratoAnterior: extratoAnterior, totalLojas: rows.length, ultimaDataConsiderada: ultimaDataConsideradaTxt },
      insights: insights,
      summary: {
        totalPrev: totalPrev,
        totalCur: totalCur,
        totalDelta: totalDelta,
        totalVarPctTxt: totalVarPctTxt,
        topCats: topCats,
        topLojas: topLojas,
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
  var res = enviarEmailOfensorasPendenciasClara(60);
  if (!res || !res.ok) {
    Logger.log("Falha ao enviar e-mail de ofensoras");
    return;
  }

  // Marca assinatura como já enviada
  props.setProperty(KEY_ULT_ENVIO, sigAtual.sig);

  Logger.log("E-mail semanal de lojas ofensoras enviado com sucesso.");
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

  // Admin: retorna completo
  return {
    baseClara: baseClaraTxt,
    jobs: "Executados com sucesso",

    // ajuste conforme o que você já implementou
    google: googleTxt,
    bigquery: bqTxt,
    alertas: "Ativos",
    geral: "Em operação"
  };
}
