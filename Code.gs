function isAdminEmail(email) {
  if (!email) return false;
  email = email.toLowerCase();

  var ADM_EMAILS = [
    "rodrigo.lisboa@gruposbf.com.br",
    "tainara.nascimento@gruposbf.com.br"
  ];

  return ADM_EMAILS.indexOf(email) !== -1;
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

 var role = "Acesso padrão";
if (isAdminEmail(email)) {
  role = "Administrador";
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


// 🌐 BigQuery – validação de loja
const PROJECT_ID = 'cnto-data-lake';
const BQ_TABLE_LOJAS = '`cnto-data-lake.refined.cnt_ref_gld_dim_estrutura_loja`';

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
  try {
    if (!emailDestino) {
      return { ok: false, error: "E-mail não informado." };
    }

    var emailRegex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
    if (!emailRegex.test(emailDestino)) {
      return { ok: false, error: "E-mail inválido." };
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
    "papel couche", "couche", "laminacao", "recorte", "vinil",

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
