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

  var template = HtmlService
    .createTemplateFromFile('index');

  // passa o nome para o HTML
  template.userName  = nome;
  // 👇 passa também o e-mail bruto
  template.userEmail = email;


  return template
    .evaluate()
    .setTitle('Grupo SBF | Vektor')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


// 🌐 BigQuery – validação de loja
const PROJECT_ID = 'cnto-data-lake';
const BQ_TABLE_LOJAS = '`cnto-data-lake.refined.cnt_ref_gld_dim_estrutura_loja`';

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

  var ss  = SpreadsheetApp.getActiveSpreadsheet();
  var aba = ss.getSheetByName("CLARA_PEND");
  if (!aba) {
    throw new Error("Aba 'CLARA_PEND' não encontrada.");
  }

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
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName("CLARA_PEND");
    if (!aba) {
      return { ok: false, error: "Aba 'CLARA_PEND' não encontrada." };
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
      top: top
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
