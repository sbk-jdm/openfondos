// ================================================================
// OPENFONDOS — Informe Diario Automatico
// Google Apps Script  |  Version 1.5
// ================================================================
// Cambios v1.5:
//   - fetchMarketNews(): descarga titulares RSS de Reuters,
//     Yahoo Finance y Expansion, filtrando por relevancia
//     para la cartera (ultimas 48h, keywords sectoriales)
//   - callAnthropicAPI(): nuevo parametro 'news', calcula
//     top movers (bajadas/subidas en €) antes del prompt
//   - Prompt reescrito: pide analisis causal y cuantificado,
//     correlacionando noticias RSS con movimientos concretos
//     de fondos en euros
// ================================================================

// ── CONFIGURACION ────────────────────────────────────────────────
function getConfig() {
  const props = PropertiesService.getScriptProperties();
  return {
    spreadsheetId : props.getProperty('SPREADSHEET_ID'),
    sheetName     : props.getProperty('SHEET_NAME') || 'Portfolio',
    emailTo       : [props.getProperty('EMAIL_TO'), props.getProperty('EMAIL_TO_2')]
                    .filter(Boolean).join(','),
    apiKey        : props.getProperty('ANTHROPIC_API_KEY'),
    model         : 'claude-sonnet-4-6',
  };
}

// ================================================================
// FUNCION PRINCIPAL
// ================================================================
function sendDailyReport() {
  try {
    const cfg = getConfig();

    Logger.log('[ 1/5 ] Leyendo cartera...');
    const portfolio = readPortfolioData(cfg);

    Logger.log('[ 2/5 ] Datos de mercado...');
    const market = fetchMarketData();

    Logger.log('[ 2b  ] Noticias RSS...');
    const news = fetchMarketNews();

    Logger.log('[ 3/5 ] Analisis IA...');
    const analysis = callAnthropicAPI(cfg, portfolio, market, news);

    Logger.log('[ 4/5 ] Construyendo email...');
    const html = buildEmailHTML(portfolio, market, analysis);

    Logger.log('[ 5/5 ] Enviando a ' + cfg.emailTo + '...');
    const today = Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM/yyyy');
    GmailApp.sendEmail(
      cfg.emailTo,
      'OpenFondos | Informe Diario ' + today,
      'Su cliente de correo no soporta HTML.',
      { htmlBody: html, name: 'OpenFondos Informe Diario' }
    );
    Logger.log('OK.');

  } catch (err) {
    Logger.log('ERROR: ' + err.message);
    try {
      GmailApp.sendEmail(
        getConfig().emailTo,
        'OpenFondos | ERROR informe diario',
        'Error: ' + err.message + '\n\nRevise los logs en Apps Script > Ejecuciones.'
      );
    } catch(e) {}
  }
}

// ================================================================
// LECTURA DE LA HOJA
// ================================================================
function readPortfolioData(cfg) {
  const ss    = SpreadsheetApp.openById(cfg.spreadsheetId);
  const sheet = ss.getSheetByName(cfg.sheetName);
  const data  = sheet.getDataRange().getValues();

  // ── Metricas globales ────────────────────────────────────────
  let capitalInvertido = 0, rentGlobal = 0, rentFiscalGlobal = 0, tirGlobal = 0;
  for (let r = 0; r < Math.min(10, data.length); r++) {
    for (let c = 0; c < data[r].length - 1; c++) {
      const lbl = String(data[r][c]).trim().toLowerCase();
      const val = data[r][c + 1];
      if (lbl.includes('capital real desembolsado'))  capitalInvertido = parseCurrency(val);
      if (lbl === 'rentabilidad global %:')            rentGlobal       = parsePct(val);
      if (lbl === 'rentabilidad global fiscal %:')     rentFiscalGlobal = parsePct(val);
      if (lbl === 'tir global:')                       tirGlobal        = parsePct(val);
    }
  }

  // ── Cabecera de fondos ───────────────────────────────────────
  let headerRow = -1;
  for (let r = 0; r < data.length; r++) {
    if (String(data[r][0]).trim().toUpperCase() === 'ISIN') { headerRow = r; break; }
  }
  if (headerRow === -1)
    throw new Error('Cabecera ISIN no encontrada en hoja ' + cfg.sheetName);

  const norm = s => String(s).trim().toLowerCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g, '');
  const headers = data[headerRow].map(norm);
  Logger.log('Headers: ' + headers.join(' | '));

  const fc = kws => {
    for (const kw of kws) {
      const i = headers.findIndex(h => h.includes(kw));
      if (i >= 0) return i;
    }
    return -1;
  };

  const col = {
    isin         : fc(['isin']),
    fondo        : fc(['fondo']),
    nav          : fc(['nav']),
    currency     : fc(['currency']),
    returnytd    : fc(['returnytd']),
    return1d     : fc(['return1d']),
    return1m     : fc(['return1m']),
    return1y     : fc(['return1y']),
    date         : fc(['date']),
    category     : fc(['category']),
    participac   : fc(['participaciones']),
    valorCompra  : fc(['valor compra']),
    valorActual  : fc(['valor actual']),
    rentBrutaEur : fc(['rentabilidad actual bruta', 'rentabilidad latente bruta',
                       'plusvalia latente bruta']),
    rentFiscalEur: fc(['rentabilidad latente fiscal €', 'rentabilidad latente fiscal e',
                       'plusvalia latente eur', 'plusvalia latente €']),
    rentActualPct: fc(['rentabilidad actual %', 'rentabilidad latente %',
                       'plusvalia latente %']),
    rentFiscalPct: fc(['rentabilidad latente fiscal %', 'rentabilidad fiscal %',
                       'porcentaje de ganancia']),
    tir          : fc(['tir %', 'tir']),
  };

  Logger.log('Cols: rentBrutaEur=' + col.rentBrutaEur +
             ' rentFiscalEur='     + col.rentFiscalEur +
             ' rentActualPct='     + col.rentActualPct +
             ' rentFiscalPct='     + col.rentFiscalPct);

  // ── Detectar ISINs con traspaso ──────────────────────────────
  const traspasoISINs = new Set();
  for (let r = headerRow + 1; r < data.length; r++) {
    const tipo = String(data[r][11] || '').trim().toUpperCase();
    if (tipo.includes('TRASPASO')) {
      const isin = String(data[r][1] || '').trim();
      if (isin.length >= 10) traspasoISINs.add(isin);
    }
  }

  // ── Filas de fondos ──────────────────────────────────────────
  const funds = [];
  let totalActual = 0, totalAnterior = 0;

  for (let r = headerRow + 1; r < data.length; r++) {
    const row      = data[r];
    const isin     = String(row[col.isin] || '').trim();
    if (!isin || isin.length < 10) continue;
    const participac = parseFloat(row[col.participac]) || 0;
    if (participac === 0) continue;

    const valorAct   = col.valorActual   >= 0 ? parseCurrency(row[col.valorActual])   : 0;
    const ret1d      = col.return1d      >= 0 ? parsePct(row[col.return1d])           : null;
    const rentBrEur  = col.rentBrutaEur  >= 0 ? parseCurrency(row[col.rentBrutaEur])  : 0;
    const rentFisEur = col.rentFiscalEur >= 0 ? parseCurrency(row[col.rentFiscalEur]) : rentBrEur;
    const rentActPct = col.rentActualPct >= 0 ? parsePct(row[col.rentActualPct])      : null;
    const rentFisPct = col.rentFiscalPct >= 0 ? parsePct(row[col.rentFiscalPct])      : null;

    const valorCompEfectivo = valorAct - rentBrEur;
    const valorAyer         = (ret1d !== null && ret1d !== 0 && valorAct > 0)
      ? valorAct / (1 + ret1d / 100) : valorAct;
    const varDiaEur         = valorAct - valorAyer;

    totalActual   += valorAct;
    totalAnterior += valorAyer;

    funds.push({
      isin          : isin,
      nombre        : String(row[col.fondo]    || '').trim(),
      nav           : parseFloat(row[col.nav]) || 0,
      moneda        : String(row[col.currency] || 'EUR').trim(),
      category      : String(row[col.category] || '').trim(),
      returnYtd     : col.returnytd >= 0 ? parsePct(row[col.returnytd]) : null,
      return1d      : ret1d,
      return1m      : col.return1m >= 0 ? parsePct(row[col.return1m]) : null,
      return1y      : col.return1y >= 0 ? parsePct(row[col.return1y]) : null,
      valorCompra   : valorCompEfectivo,
      importeActual : valorAct,
      varDiaEur     : varDiaEur,
      rentBrutaEur  : rentBrEur,
      rentFiscalEur : rentFisEur,
      rentActualPct : rentActPct,
      rentFiscalPct : rentFisPct,
      esTraspaso    : traspasoISINs.has(isin),
      tir           : col.tir >= 0 ? parsePct(row[col.tir]) : null,
      fechaNAV      : String(row[col.date] || '').trim(),
    });
  }

  const variacionEur = totalActual - totalAnterior;
  const variacionPct = totalAnterior > 0 ? (variacionEur / totalAnterior * 100) : 0;
  const varRentPp    = capitalInvertido > 0 ? (variacionEur / capitalInvertido * 100) : 0;

  if (!capitalInvertido) capitalInvertido = 218059.29;
  if (!rentGlobal)       rentGlobal = capitalInvertido > 0
    ? (totalActual - capitalInvertido) / capitalInvertido * 100 : 0;
  if (!rentFiscalGlobal) rentFiscalGlobal = rentGlobal;

  return {
    totalActual, totalAnterior, variacionEur, variacionPct, varRentPp,
    capitalInvertido,
    plusvaliaTotal   : totalActual - capitalInvertido,
    rentGlobal, rentFiscalGlobal, tirGlobal,
    fechaInforme     : Utilities.formatDate(new Date(), 'Europe/Madrid', 'dd/MM/yyyy'),
    diaSemana        : getDiaSemana(),
    funds,
  };
}

// ── Parseo de moneda ─────────────────────────────────────────────
function parseCurrency(val) {
  if (val === null || val === undefined || val === '') return 0;
  if (typeof val === 'number') return val;
  const s = String(val).replace(/[€$\s]/g, '').replace(/\./g, '').replace(',', '.');
  return parseFloat(s) || 0;
}

// ── Parseo de porcentaje ──────────────────────────────────────────
// Sheets guarda % como decimal (0.0127 = 1.27%); si viene como numero → *100
function parsePct(val) {
  if (val === null || val === undefined || val === '') return null;
  if (typeof val === 'number') {
    if (isNaN(val)) return null;
    return val * 100;
  }
  const s = String(val).replace('%', '').replace(',', '.').trim();
  if (!s || s === '#N/A' || s === '#ERROR!' || s === 'N/D' || s === '-') return null;
  const n = parseFloat(s);
  return isNaN(n) ? null : n;
}

function getDiaSemana() {
  return ['Domingo','Lunes','Martes','Miercoles','Jueves','Viernes','Sabado'][new Date().getDay()];
}

// ================================================================
// DATOS DE MERCADO — Yahoo Finance
// ================================================================
function fetchMarketData() {
  const tickers = {
    ibex   : '%5EIBEX',
    sp500  : '%5EGSPC',
    nasdaq : '%5EIXIC',
    dax    : '%5EGDAXI',
    brent  : 'BZ%3DF',
    oro    : 'GC%3DF',
    eurusd : 'EURUSD%3DX',
  };
  const res = {};
  for (const [key, sym] of Object.entries(tickers)) {
    try {
      const r = UrlFetchApp.fetch(
        'https://query1.finance.yahoo.com/v8/finance/chart/' + sym + '?interval=1d&range=2d',
        { muteHttpExceptions: true }
      );
      if (r.getResponseCode() !== 200) { res[key] = null; continue; }
      const meta  = JSON.parse(r.getContentText())?.chart?.result?.[0]?.meta;
      if (!meta)  { res[key] = null; continue; }
      const precio   = meta.regularMarketPrice || 0;
      const anterior = meta.chartPreviousClose || meta.previousClose || precio;
      res[key] = { precio, cambio: anterior > 0 ? (precio - anterior) / anterior * 100 : 0 };
    } catch(e) { res[key] = null; }
    Utilities.sleep(300);
  }
  res.euribor = getEuriborFromSheet();
  return res;
}

function getEuriborFromSheet() {
  try {
    const cfg  = getConfig();
    const data = SpreadsheetApp.openById(cfg.spreadsheetId)
      .getSheetByName(cfg.sheetName).getDataRange().getValues();
    for (const row of data) {
      for (let c = 0; c < row.length - 1; c++) {
        if (String(row[c]).toLowerCase().includes('euribor')) {
          const v = parseFloat(String(row[c + 1]).replace(',', '.'));
          if (!isNaN(v)) return { precio: v, cambio: 0 };
        }
      }
    }
  } catch(e) {}
  return { precio: 2.18, cambio: 0 };
}

// ================================================================
// RSS DE NOTICIAS FINANCIERAS
// ================================================================
function fetchMarketNews() {
  const feeds = [
    { url: 'https://feeds.reuters.com/reuters/technologyNews',  label: 'Reuters Tech'      },
    { url: 'https://feeds.reuters.com/reuters/businessNews',    label: 'Reuters Business'   },
    { url: 'https://finance.yahoo.com/rss/topfinstories',       label: 'Yahoo Finance'      },
    { url: 'https://e00-expansion.uecdn.es/rss/mercados.xml',   label: 'Expansion Mercados' },
  ];

  // Keywords relevantes para los sectores de la cartera
  const keywords = [
    'nvidia','microsoft','alphabet','amazon','meta','apple','broadcom','amd','intel',
    'semiconductor','chip','artificial intelligence','ai','machine learning',
    'nasdaq','s&p','tech','tecnolog',
    'fed','federal reserve','reserva federal','interest rate','tipo','inflation','inflacion',
    'nonfarm payroll','nominas','empleo','employment',
    'oil','brent','opec','petroleo','energy','energia',
    'gold','oro','mining','mineria',
    'defence','defense','defensa','military','nato','otan',
    'cybersecurity','ciberseguridad','crowdstrike',
    'space','espacio','spacex','satellite',
    'ecb','bce','bank','earnings','results','beneficios','resultados',
    'ibex','dax','eurostoxx','europe','europa','bolsa','mercado',
  ];

  const headlines = [];
  const cutoffMs  = Date.now() - 48 * 60 * 60 * 1000; // ultimas 48h

  for (const feed of feeds) {
    try {
      const resp = UrlFetchApp.fetch(feed.url, {
        muteHttpExceptions : true,
        headers            : { 'User-Agent': 'Mozilla/5.0 (compatible; OpenFondos/1.5)' },
      });

      if (resp.getResponseCode() !== 200) {
        Logger.log('RSS ' + feed.label + ' → HTTP ' + resp.getResponseCode());
        continue;
      }

      const doc  = XmlService.parse(resp.getContentText());
      const root = doc.getRootElement();

      // Detectar RSS 2.0 o Atom
      let items = [];
      const channel = root.getChild('channel');
      if (channel) {
        items = channel.getChildren('item');           // RSS 2.0
      } else {
        const atomNs = XmlService.getNamespace('http://www.w3.org/2005/Atom');
        items = root.getChildren('entry', atomNs);
        if (!items || items.length === 0) items = root.getChildren('entry');
      }

      Logger.log('RSS ' + feed.label + ' → ' + items.length + ' items');

      for (const item of items) {
        try {
          const title   = (item.getChildText('title')   || '').replace(/<[^>]+>/g, '').trim();
          const pubDate =  item.getChildText('pubDate')  ||
                           item.getChildText('published') || '';
          const desc    = (item.getChildText('description') ||
                           item.getChildText('summary')     || '')
                          .replace(/<[^>]+>/g, '').trim().substring(0, 300);

          // Filtrar por fecha
          if (pubDate) {
            const d = new Date(pubDate);
            if (!isNaN(d.getTime()) && d.getTime() < cutoffMs) continue;
          }

          // Filtrar por relevancia
          const text = (title + ' ' + desc).toLowerCase();
          if (!keywords.some(kw => text.includes(kw))) continue;

          if (title) headlines.push('[' + feed.label + '] ' + title);

        } catch(e) { /* item malformado, ignorar */ }
      }

    } catch(e) {
      Logger.log('RSS error ' + feed.label + ': ' + e.message);
    }
    Utilities.sleep(200);
  }

  // Eliminar duplicados exactos y limitar a 15 titulares
  const unique = [...new Set(headlines)].slice(0, 15);
  Logger.log('Noticias relevantes encontradas: ' + unique.length);
  return unique;
}

// ================================================================
// API ANTHROPIC — v1.5: incluye news, top movers y prompt causal
// ================================================================
function callAnthropicAPI(cfg, portfolio, market, news) {

  // ── Top movers del dia (para el prompt) ─────────────────────
  const conVar     = portfolio.funds.filter(f => f.varDiaEur !== 0);
  const topBajadas = [...conVar].sort((a, b) => a.varDiaEur - b.varDiaEur).slice(0, 5);
  const topSubidas = [...conVar].sort((a, b) => b.varDiaEur - a.varDiaEur).slice(0, 5);

  const fmtMover = f =>
    `${f.nombre} | Actual: ${fmtEur(f.importeActual)} | Var.dia: ${fmtEur(f.varDiaEur)} ` +
    `(${f.return1d !== null ? (f.return1d >= 0 ? '+' : '') + fmtNum(f.return1d, 2) + '%' : '—'}) ` +
    `| R.Total: ${fmtPct(f.rentActualPct)} | Sector: ${f.category}`;

  const resumenBajadas = topBajadas.length > 0
    ? topBajadas.map(fmtMover).join('\n')
    : 'Sin bajadas significativas.';
  const resumenSubidas = topSubidas.length > 0
    ? topSubidas.map(fmtMover).join('\n')
    : 'Sin subidas significativas.';

  // ── Resumen de todos los fondos ──────────────────────────────
  const fondosResumen = [...portfolio.funds]
    .sort((a, b) => b.importeActual - a.importeActual)
    .map(f =>
      `${f.nombre} | Actual: ${fmtEur(f.importeActual)} | Dia: ${fmtPct(f.return1d)} | ` +
      `R.Total: ${fmtEur(f.rentBrutaEur)} (${fmtPct(f.rentActualPct)}) | ` +
      `R.Fiscal: ${fmtEur(f.rentFiscalEur)} (${fmtPct(f.rentFiscalPct)})`
    ).join('\n');

  // ── Noticias RSS ─────────────────────────────────────────────
  const newsSection = (news && news.length > 0)
    ? news.join('\n')
    : 'No se han podido obtener noticias externas en este momento.';

  // ── Datos de mercado ─────────────────────────────────────────
  const mktStr = (key, dec, suf) => market[key]
    ? fmtNum(market[key].precio, dec) + suf + ' (' + fmtPct(market[key].cambio) + ')'
    : 'N/D';

  // ── Prompt ───────────────────────────────────────────────────
  const prompt = `Eres un asesor de inversiones experto analizando una cartera de fondos de inversion española.

Genera DOS bloques en español, SIN markdown ni asteriscos, lenguaje claro y accesible:

BLOQUE 1 — ANALISIS DEL DIA (5-6 parrafos, este es el mas importante):
Sigue este esquema obligatoriamente:
- Parrafo 1: resumen ejecutivo del dia en la cartera: cuanto ha subido o bajado en euros y %, y la causa principal en una frase directa.
- Parrafos 2 y 3: para cada fondo que haya bajado significativamente, identifica el evento concreto de las noticias (resultados empresariales, dato macro, decision del banco central, noticia geopolitica) que lo ha provocado. Explica por que ese evento afecta especificamente a ese fondo usando la cadena de causalidad: evento → efecto en el mercado → efecto en el fondo → impacto en euros en la cartera. Se especifico: no digas "el sector tecnologico bajo", di "Broadcom publico una guia de ingresos por chips de IA por debajo de expectativas, lo que arrastra a fondos de semiconductores como X que perdio Y euros".
- Parrafo 4: fondos que han subido o aguantado bien y por que esos sectores se comportan diferente en este entorno. Menciona el importe ganado.
- Parrafo 5: contexto y perspectiva: si la caida/subida es una correccion tecnica, un cambio de tendencia o ruido puntual. Incluye cuanto lleva acumulado la cartera y que significa en el contexto de la rentabilidad historica.
- Parrafo 6 (solo si hay algo relevante): implicacion fiscal, ventaja tactica o punto de atencion especial.

REGLAS OBLIGATORIAS DEL ANALISIS:
- Menciona siempre los fondos por su nombre abreviado.
- Incluye siempre el impacto en euros de los movimientos principales.
- Usa las noticias RSS proporcionadas como base para identificar los eventos del dia.
- Explica la cadena de causalidad de forma sencilla.
- No uses terminos tecnicos sin explicarlos. No uses vinetas ni listas, solo prosa.

BLOQUE 2 — ESCENARIOS PROXIMO MES (exactamente 3):
Para Optimista, Central y Adverso: subtitulo breve, probabilidad (suman 100%), impacto especifico y cuantificado en esta cartera en 2-3 frases. Menciona fondos concretos.

Devuelve SOLO JSON valido sin ningun texto ni markdown alrededor:
{
  "analisis": "parrafo1\\n\\nparrafo2\\n\\nparrafo3\\n\\nparrafo4\\n\\nparrafo5",
  "escenarios": [
    {"nombre":"Optimista","subtitulo":"...","probabilidad":40,"descripcion":"..."},
    {"nombre":"Central","subtitulo":"...","probabilidad":35,"descripcion":"..."},
    {"nombre":"Adverso","subtitulo":"...","probabilidad":25,"descripcion":"..."}
  ]
}

DATOS DE LA CARTERA (${portfolio.fechaInforme}):
Valor total: ${fmtEur(portfolio.totalActual)}
Variacion dia: ${fmtEur(portfolio.variacionEur)} (${fmtPct(portfolio.variacionPct)})
Capital invertido: ${fmtEur(portfolio.capitalInvertido)}
Plusvalia acumulada: ${fmtEur(portfolio.plusvaliaTotal)}
Rent. s/capital: ${fmtPct(portfolio.rentGlobal)} | Rent. fiscal: ${fmtPct(portfolio.rentFiscalGlobal)} | TIR: ${fmtPct(portfolio.tirGlobal)}

FONDOS QUE MAS HAN BAJADO HOY:
${resumenBajadas}

FONDOS QUE MAS HAN SUBIDO HOY:
${resumenSubidas}

TODOS LOS FONDOS (por valor):
${fondosResumen}

MERCADOS HOY:
IBEX 35: ${mktStr('ibex', 0, ' pts')} | S&P 500: ${mktStr('sp500', 0, '')} | Nasdaq: ${mktStr('nasdaq', 0, '')} | DAX: ${mktStr('dax', 0, '')}
Brent: ${mktStr('brent', 1, '$/b')} | Oro: ${mktStr('oro', 0, '$/oz')} | EUR/USD: ${mktStr('eurusd', 4, '')} | Euribor 12M: ${market.euribor ? fmtNum(market.euribor.precio, 2) + '%' : 'N/D'}

NOTICIAS RELEVANTES DE LAS ULTIMAS 48 HORAS (usa estas noticias para explicar los movimientos de los fondos):
${newsSection}`;

  // ── Llamada a la API ─────────────────────────────────────────
  const resp = UrlFetchApp.fetch('https://api.anthropic.com/v1/messages', {
    method      : 'post',
    contentType : 'application/json',
    headers     : { 'x-api-key': cfg.apiKey, 'anthropic-version': '2023-06-01' },
    payload     : JSON.stringify({
      model      : cfg.model,
      max_tokens : 3000,
      messages   : [{ role: 'user', content: prompt }],
    }),
    muteHttpExceptions: true,
  });

  if (resp.getResponseCode() !== 200)
    throw new Error('API Anthropic ' + resp.getResponseCode() + ': ' +
                    resp.getContentText().substring(0, 200));

  const text  = JSON.parse(resp.getContentText())?.content?.[0]?.text || '{}';
  const clean = text
    .replace(/^```json[\r\n]*/i, '')
    .replace(/^```[\r\n]*/i, '')
    .replace(/[\r\n]*```$/i, '')
    .trim();

  try {
    return JSON.parse(clean);
  } catch(e) {
    Logger.log('JSON parse error. Texto recibido: ' + text.substring(0, 400));
    return {
      analisis: 'El analisis no pudo generarse hoy. Revise los logs.',
      escenarios: [
        { nombre: 'Optimista', subtitulo: 'N/D', probabilidad: 40, descripcion: '—' },
        { nombre: 'Central',   subtitulo: 'N/D', probabilidad: 35, descripcion: '—' },
        { nombre: 'Adverso',   subtitulo: 'N/D', probabilidad: 25, descripcion: '—' },
      ],
    };
  }
}

// ================================================================
// EMAIL HTML — v1.5
// ================================================================
function buildEmailHTML(portfolio, market, analysis) {
  const p      = portfolio;
  const varPos = p.variacionEur >= 0;
  const vCol   = varPos ? '#3f9e6a' : '#c94040';
  const vSign  = varPos ? '+' : '';

  // ── Fondos ordenados por valor descendente ───────────────────
  const fondos = [...p.funds].sort((a, b) => b.importeActual - a.importeActual);

  // ── Estilos (declarados ANTES del map para evitar TDZ) ───────
  const tdBase           = 'padding:10px 10px;vertical-align:top;border-bottom:1px solid #f0ede7;';
  const tdL              = tdBase;
  const tdR              = tdBase + 'text-align:right;';
  const numBadgeStyle    = 'display:inline-block;width:18px;height:18px;background:#0d1820;color:#c9a84c;' +
    'font-size:9px;font-weight:700;border-radius:2px;text-align:center;line-height:18px;flex-shrink:0;margin-top:1px;';
  const catTagStyle      = 'display:inline-block;font-size:9.5px;background:#ebe7df;color:#6a6258;' +
    'padding:1px 6px;border-radius:2px;margin-top:3px;';
  const tagTraspasoStyle = 'display:inline-block;font-size:9px;background:#0d1820;color:#c9a84c;' +
    'padding:1px 5px;border-radius:2px;margin-top:4px;font-weight:600;letter-spacing:0.5px;';
  const tagVentajaStyle  = 'display:inline-block;font-size:9px;background:#eaf5ee;color:#1a6038;' +
    'border:1px solid #c2e8cc;padding:1px 5px;border-radius:2px;margin-top:4px;';

  // ── Filas de la tabla ────────────────────────────────────────
  const filas = fondos.map((f, idx) => {
    const vd    = f.varDiaEur;
    const vdPos = vd >= 0;
    const vdCol = vdPos ? '#2a7a4f' : '#b83030';
    const vdSgn = vdPos ? '+' : '';

    const colInv = `
      <td style="${tdR}">
        <div style="font-size:10.5px;color:#9a9288;margin-bottom:2px;">Compra: ${fmtEur(f.valorCompra)}</div>
        <div style="font-size:16px;font-weight:700;color:#0d1820;font-family:Georgia,serif;">${fmtEur(f.importeActual)}</div>
        <div style="font-size:10.5px;font-weight:500;color:${vdCol};margin-top:4px;">${vdSgn}${fmtEur(vd)} (${f.return1d !== null ? (f.return1d >= 0 ? '+' : '') + fmtNum(f.return1d, 2) + '%' : '—'})</div>
      </td>`;

    const rbPos = f.rentBrutaEur >= 0;
    const rbCol = rbPos ? '#2a7a4f' : '#b83030';
    const rbSgn = rbPos ? '+' : '';
    const tagT  = f.esTraspaso ? `<span style="${tagTraspasoStyle}">T</span>` : '';
    const colRT = `
      <td style="${tdR}">
        <div style="font-size:16px;font-weight:700;color:${rbCol};font-family:Georgia,serif;">${rbSgn}${fmtEur(f.rentBrutaEur)}</div>
        <div style="font-size:11.5px;font-weight:500;color:${rbCol};margin-top:2px;">${rbSgn}${f.rentActualPct !== null ? fmtNum(f.rentActualPct, 2) + '%' : '—'}</div>
        <div style="font-size:10px;color:#aaa;margin-top:4px;padding-top:4px;border-top:1px dashed #e8e4dc;">&Delta; ayer ${vdSgn}${fmtEur(vd)}</div>
        ${tagT}
      </td>`;

    const rfPos   = f.rentFiscalEur >= 0;
    const rfCol   = rfPos ? '#2a7a4f' : '#b83030';
    const rfSgn   = rfPos ? '+' : '';
    const tagVent = (!rfPos && f.rentFiscalEur < -5)
      ? `<span style="${tagVentajaStyle}">&#9651; Ventaja fiscal</span>` : '';
    const colRF = `
      <td style="${tdR}">
        <div style="font-size:16px;font-weight:700;color:${rfCol};font-family:Georgia,serif;">${rfSgn}${fmtEur(f.rentFiscalEur)}</div>
        <div style="font-size:11.5px;font-weight:500;color:${rfCol};margin-top:2px;">${rfSgn}${f.rentFiscalPct !== null ? fmtNum(f.rentFiscalPct, 2) + '%' : '—'}</div>
        ${tagVent}
      </td>`;

    return `<tr style="${idx % 2 === 1 ? 'background:#faf8f5' : ''}">
      <td style="${tdL}">
        <div style="display:flex;align-items:flex-start;gap:6px;">
          <span style="${numBadgeStyle}">${idx + 1}</span>
          <div>
            <div style="font-weight:600;color:#0d1820;font-size:12.5px;">${esc(f.nombre)}</div>
            <div style="font-size:10px;color:#9a9288;font-family:monospace;margin-top:2px;">${esc(f.isin)}</div>
            <span style="${catTagStyle}">${esc(catCorta(f.category))}</span>
          </div>
        </div>
      </td>
      ${colInv}${colRT}${colRF}
    </tr>`;
  }).join('');

  // ── Sub-linea variacion cabecera ─────────────────────────────
  function varSub(eur, pct) {
    const c = eur >= 0 ? '#3f9e6a' : '#c94040';
    const s = eur >= 0 ? '+' : '';
    return `<div style="font-size:10px;color:${c};margin-top:4px;">${s}${fmtEur(eur)} (${s}${fmtNum(pct, 2)}%) vs ayer</div>`;
  }

  // ── Escenarios ───────────────────────────────────────────────
  const escCols = { Optimista: '#2a7a4f', Central: '#c9a84c', Adverso: '#b83030' };
  const escBgs  = { Optimista: '#eaf5ee', Central: '#fdf8ec', Adverso: '#faeaea' };
  const escHtml = (analysis.escenarios || []).map(e => {
    const col = escCols[e.nombre] || '#555';
    const bg  = escBgs[e.nombre]  || '#f9f9f9';
    return `<div style="border:1px solid #ebe7df;margin-bottom:12px;overflow:hidden;">
      <div style="background:${bg};border-left:4px solid ${col};padding:12px 16px;display:flex;justify-content:space-between;align-items:center;">
        <div style="font-weight:600;font-size:14px;color:${col};">${esc(e.nombre)} &mdash; ${esc(e.subtitulo || '')}</div>
        <div style="font-size:24px;font-weight:700;color:${col};font-family:Georgia,serif;">${e.probabilidad}%</div>
      </div>
      <div style="padding:12px 16px;font-size:14px;line-height:1.7;color:#262018;">${esc(e.descripcion || '')}</div>
    </div>`;
  }).join('');

  // ── Narrativo ────────────────────────────────────────────────
  const narr = (analysis.analisis || '').split('\n\n').filter(t => t.trim())
    .map(t => `<p style="margin-bottom:14px;">${esc(t.trim())}</p>`).join('');

  // ── Celda indice ─────────────────────────────────────────────
  function idxCell(lbl, mkt, dec, suf) {
    const base = 'background:#f9f7f3;border:1px solid #ebe7df;border-top:3px solid #0d1820;padding:12px 14px;';
    if (!mkt) return `<td width="25%" style="padding:0 5px 10px;"><div style="${base}">
      <div style="font-size:10px;letter-spacing:1px;text-transform:uppercase;color:#9a9288;margin-bottom:6px;">${lbl}</div>
      <div style="font-size:18px;font-family:Georgia,serif;color:#0d1820;">N/D</div>
      <div style="font-size:12px;color:#888;">—</div></div></td>`;
    const c  = mkt.cambio >= 0 ? '#2a7a4f' : '#b83030';
    const sg = mkt.cambio >= 0 ? '+' : '';
    return `<td width="25%" style="padding:0 5px 10px;"><div style="${base}">
      <div style="font-size:10px;letter-spacing:1px;text-transform:uppercase;color:#9a9288;margin-bottom:6px;">${lbl}</div>
      <div style="font-size:18px;font-family:Georgia,serif;color:#0d1820;margin-bottom:3px;">${fmtNum(mkt.precio, dec)}${suf}</div>
      <div style="font-size:12px;font-weight:600;color:${c};">${sg}${fmtNum(mkt.cambio, 2)}%</div></div></td>`;
  }

  // ── Barra de contexto ────────────────────────────────────────
  function mktItem(key, label, dec, suf) {
    if (!market[key]) return '';
    const sg = market[key].cambio >= 0 ? '+' : '';
    return `${label}: ${fmtNum(market[key].precio, dec)}${suf} (${sg}${fmtNum(market[key].cambio, 2)}%)`;
  }
  const mktBar = [
    mktItem('ibex',   'IBEX 35',    0, ' pts'),
    mktItem('sp500',  'S&amp;P 500', 0, ''),
    mktItem('brent',  'Brent',      1, '$/b'),
    market.eurusd  ? 'EUR/USD: '    + fmtNum(market.eurusd.precio,  4)       : '',
    market.euribor ? 'Euribor: '    + fmtNum(market.euribor.precio, 2) + '%' : '',
  ].filter(Boolean).join(' &nbsp;&bull;&nbsp; ');

  const hoy = p.fechaInforme;
  const dia = p.diaSemana;

  // ================================================================
  // HTML COMPLETO
  // ================================================================
  return `<!DOCTYPE html>
<html lang="es">
<head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1.0">
<title>OpenFondos | ${hoy}</title></head>
<body style="margin:0;padding:28px 12px;background:#eceae4;font-family:Georgia,serif;font-size:15px;">
<div style="max-width:900px;margin:0 auto;background:#fff;border:1px solid #dedad2;">

  <!-- CABECERA -->
  <div style="background:#0d1820;padding:28px 40px 0;">
    <div style="display:flex;justify-content:space-between;align-items:flex-start;margin-bottom:20px;">
      <div>
        <div style="font-size:23px;color:#c9a84c;font-family:Georgia,serif;">OpenFondos</div>
        <div style="font-size:11px;color:#617588;letter-spacing:2px;text-transform:uppercase;margin-top:3px;">Informe Diario de Cartera</div>
      </div>
      <div style="text-align:right;">
        <div style="font-size:28px;color:#fff;font-family:Georgia,serif;line-height:1;">${hoy.substring(0, 2)}</div>
        <div style="font-size:11px;color:#617588;letter-spacing:1.5px;text-transform:uppercase;margin-top:4px;">${dia} &nbsp;|&nbsp; ${hoy.substring(3)}</div>
      </div>
    </div>
    <table width="100%" cellpadding="0" cellspacing="1" bgcolor="#162330">
      <tr>
        <td width="34%" bgcolor="#0d1820" style="padding:15px 14px;text-align:center;">
          <div style="font-size:10px;color:#617588;letter-spacing:1.3px;text-transform:uppercase;margin-bottom:6px;">Valor de Mercado</div>
          <div style="font-size:20px;color:#c9a84c;font-family:Georgia,serif;">${fmtEur(p.totalActual)}</div>
        </td>
        <td width="33%" bgcolor="#0d1820" style="padding:15px 14px;text-align:center;">
          <div style="font-size:10px;color:#617588;letter-spacing:1.3px;text-transform:uppercase;margin-bottom:6px;">Variacion Dia (EUR)</div>
          <div style="font-size:20px;color:${vCol};font-family:Georgia,serif;">${vSign}${fmtEur(p.variacionEur)}</div>
        </td>
        <td width="33%" bgcolor="#0d1820" style="padding:15px 14px;text-align:center;">
          <div style="font-size:10px;color:#617588;letter-spacing:1.3px;text-transform:uppercase;margin-bottom:6px;">Variacion Dia (%)</div>
          <div style="font-size:20px;color:${vCol};font-family:Georgia,serif;">${vSign}${fmtNum(p.variacionPct, 2)}%</div>
        </td>
      </tr>
    </table>
    <table width="100%" cellpadding="0" cellspacing="1" bgcolor="#1e3245" style="border-top:1px solid #1e3245;">
      <tr>
        <td width="25%" bgcolor="#0d1820" style="padding:13px 12px;text-align:center;">
          <div style="font-size:10px;color:#617588;letter-spacing:1.3px;text-transform:uppercase;margin-bottom:5px;">Plusvalia Latente</div>
          <div style="font-size:16px;color:#3f9e6a;font-family:Georgia,serif;">+${fmtEur(p.plusvaliaTotal)}</div>
          ${varSub(p.variacionEur, p.variacionPct)}
        </td>
        <td width="25%" bgcolor="#0d1820" style="padding:13px 12px;text-align:center;">
          <div style="font-size:10px;color:#617588;letter-spacing:1.3px;text-transform:uppercase;margin-bottom:5px;">Rent. s/Capital</div>
          <div style="font-size:16px;color:#3f9e6a;font-family:Georgia,serif;">+${fmtNum(p.rentGlobal, 2)}%</div>
          ${varSub(p.varRentPp, p.variacionPct)}
        </td>
        <td width="25%" bgcolor="#0d1820" style="padding:13px 12px;text-align:center;">
          <div style="font-size:10px;color:#617588;letter-spacing:1.3px;text-transform:uppercase;margin-bottom:5px;">Rent. s/Coste Fiscal</div>
          <div style="font-size:16px;color:#3f9e6a;font-family:Georgia,serif;">+${fmtNum(p.rentFiscalGlobal, 2)}%</div>
          ${varSub(p.varRentPp, p.variacionPct)}
        </td>
        <td width="25%" bgcolor="#0d1820" style="padding:13px 12px;text-align:center;">
          <div style="font-size:10px;color:#617588;letter-spacing:1.3px;text-transform:uppercase;margin-bottom:5px;">TIR Global Anual</div>
          <div style="font-size:16px;color:#c9a84c;font-family:Georgia,serif;">${fmtNum(p.tirGlobal, 2)}%</div>
          <div style="font-size:10px;color:#617588;margin-top:4px;">Capital: ${fmtEur(p.capitalInvertido)}</div>
        </td>
      </tr>
    </table>
  </div>

  <!-- ALERTA -->
  <div style="background:#fdf8ec;border-left:4px solid #c9a84c;padding:15px 26px;">
    <div style="font-size:10px;letter-spacing:2px;text-transform:uppercase;color:#8a7030;font-weight:600;margin-bottom:5px;">Contexto &nbsp;|&nbsp; ${hoy}</div>
    <div style="font-size:14px;line-height:1.65;color:#3a2f14;">${mktBar}</div>
  </div>

  <!-- ANALISIS -->
  <div style="padding:26px 40px;border-bottom:1px solid #ebe7df;">
    <div style="font-size:10px;letter-spacing:2.5px;text-transform:uppercase;color:#9a9288;font-weight:600;margin-bottom:18px;padding-bottom:10px;border-bottom:1px solid #ebe7df;">Analisis del Dia &mdash; Elaborado por IA</div>
    <div style="font-size:15px;line-height:1.85;color:#262018;">${narr}</div>
  </div>

  <!-- ESCENARIOS -->
  <div style="padding:26px 40px;border-bottom:1px solid #ebe7df;">
    <div style="font-size:10px;letter-spacing:2.5px;text-transform:uppercase;color:#9a9288;font-weight:600;margin-bottom:18px;padding-bottom:10px;border-bottom:1px solid #ebe7df;">Prediccion de Escenarios &mdash; Proximo Mes</div>
    ${escHtml}
  </div>

  <!-- INDICES -->
  <div style="padding:26px 40px;border-bottom:1px solid #ebe7df;">
    <div style="font-size:10px;letter-spacing:2.5px;text-transform:uppercase;color:#9a9288;font-weight:600;margin-bottom:18px;padding-bottom:10px;border-bottom:1px solid #ebe7df;">Indices y Referentes</div>
    <table width="100%" cellpadding="0" cellspacing="0">
      <tr>${idxCell('IBEX 35',     market.ibex,   0, '')}${idxCell('S&amp;P 500', market.sp500, 0, '')}${idxCell('Nasdaq',    market.nasdaq, 0, '')}${idxCell('DAX 40',     market.dax,    0, '')}</tr>
      <tr>${idxCell('Brent ($/b)', market.brent,  1, '')}${idxCell('Oro ($/oz)',  market.oro,   0, '')}${idxCell('EUR / USD', market.eurusd, 4, '')}${idxCell('Euribor 12M',market.euribor,2, '%')}</tr>
    </table>
  </div>

  <!-- TABLA FONDOS -->
  <div style="padding:26px 40px;border-bottom:1px solid #ebe7df;">
    <div style="font-size:10px;letter-spacing:2.5px;text-transform:uppercase;color:#9a9288;font-weight:600;margin-bottom:18px;padding-bottom:10px;border-bottom:1px solid #ebe7df;">Estado de sus Fondos &mdash; Por Valor Descendente</div>
    <table width="100%" cellpadding="0" cellspacing="0" style="border-collapse:collapse;">
      <thead>
        <tr style="background:#f4f1ec;">
          <th style="font-size:9.5px;letter-spacing:1px;text-transform:uppercase;color:#9a9288;font-weight:600;padding:8px 10px;text-align:left;border-bottom:2px solid #dedad2;width:32%;">Fondo</th>
          <th style="font-size:9.5px;letter-spacing:1px;text-transform:uppercase;color:#9a9288;font-weight:600;padding:8px 10px;text-align:right;border-bottom:2px solid #dedad2;width:18%;">Inversion &nbsp;/&nbsp; Actual &nbsp;/&nbsp; Var.Dia</th>
          <th style="font-size:9.5px;letter-spacing:1px;text-transform:uppercase;color:#9a9288;font-weight:600;padding:8px 10px;text-align:right;border-bottom:2px solid #dedad2;width:25%;">R. Total <span style="font-weight:400;text-transform:none;letter-spacing:0;font-size:8.5px;">(s/coste compra)</span></th>
          <th style="font-size:9.5px;letter-spacing:1px;text-transform:uppercase;color:#9a9288;font-weight:600;padding:8px 10px;text-align:right;border-bottom:2px solid #dedad2;width:25%;">R. Fiscal <span style="font-weight:400;text-transform:none;letter-spacing:0;font-size:8.5px;">(s/coste fiscal)</span></th>
        </tr>
      </thead>
      <tbody>${filas}</tbody>
    </table>
    <p style="font-size:10.5px;color:#9a9288;margin-top:12px;line-height:1.6;">
      <strong>R. Total</strong>: ganancia sobre el coste de compra de la operacion actual. &nbsp;
      <strong>R. Fiscal</strong>: ganancia sobre el coste fiscal (en traspasos hereda el coste del fondo origen). &nbsp;
      <strong>T</strong>: fondo implicado en un traspaso. &nbsp;
      <strong>&#9651; Ventaja fiscal</strong>: minusvalia latente aprovechable para compensar ganancias en el ejercicio del rescate. &nbsp;
      <strong>&Delta; ayer</strong>: variacion en euros respecto al ultimo NAV anterior.
    </p>
  </div>

  <!-- FOOTER -->
  <div style="background:#0d1820;padding:18px 40px;display:flex;justify-content:space-between;align-items:center;">
    <div style="font-family:Georgia,serif;font-size:14px;color:#c9a84c;">OpenFondos</div>
    <div style="font-size:10px;color:#3e5060;text-align:right;line-height:1.6;">
      ${hoy} &nbsp;|&nbsp; Google Sheets Openfondos &nbsp;|&nbsp; No constituye asesoramiento de inversion &nbsp;|&nbsp; v1.5
    </div>
  </div>

</div>
</body></html>`;
}

// ================================================================
// HELPERS
// ================================================================
function fmtEur(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  return new Intl.NumberFormat('es-ES', {
    style: 'currency', currency: 'EUR', maximumFractionDigits: 0
  }).format(n);
}
function fmtNum(n, d) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  return n.toFixed(d !== undefined ? d : 2).replace('.', ',');
}
function fmtPct(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  return (n >= 0 ? '+' : '') + fmtNum(n, 2) + '%';
}
function esc(s) {
  if (!s) return '';
  return String(s)
    .replace(/&/g, '&amp;').replace(/</g, '&lt;')
    .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}
function catCorta(cat) {
  if (!cat) return '';
  const m = {
    'sector equity technology'                  : 'Tecnologia',
    'sector equity financial services'          : 'Financiero',
    'sector equity energy'                      : 'Energia',
    'sector equity precious metals'             : 'Met. Preciosos',
    'sector equity natural resources'           : 'Mat. Primas',
    'sector equity alternative energy'          : 'Energia Alt.',
    'sector equity industrial materials'        : 'Defensa',
    'sector equity healthcare'                  : 'Salud',
    'us large-cap blend equity'                 : 'EEUU Large Cap',
    'us large-cap growth equity'                : 'EEUU Growth',
    'europe large-cap blend equity'             : 'Europa',
    'europe flex-cap equity'                    : 'Europa',
    'global flex-cap equity'                    : 'Global Flex',
    'property - indirect global'                : 'Inmobiliario',
    'other bond'                                : 'Renta Fija',
    'other equity'                              : 'Renta Variable',
    'global inflation-linked bond - eur hedged' : 'RF Inflacion',
  };
  return m[cat.trim().toLowerCase()] || cat.split(' ').slice(0, 3).join(' ');
}

// ================================================================
// TEST
// ================================================================
function testScript() {
  Logger.log('=== TEST OpenFondos v1.5 ===');

  const cfg = getConfig();
  Logger.log('Email: ' + cfg.emailTo);

  const p = readPortfolioData(cfg);
  Logger.log('Fondos: '       + p.funds.length);
  Logger.log('Valor: '        + fmtEur(p.totalActual));
  Logger.log('Variacion: '    + fmtEur(p.variacionEur) + ' (' + fmtPct(p.variacionPct) + ')');
  Logger.log('Rent.Capital: ' + fmtNum(p.rentGlobal, 2) + '%');
  Logger.log('Rent.Fiscal: '  + fmtNum(p.rentFiscalGlobal, 2) + '%');
  Logger.log('TIR: '          + fmtNum(p.tirGlobal, 2) + '%');

  Logger.log('--- Top 3 por valor ---');
  [...p.funds].sort((a, b) => b.importeActual - a.importeActual).slice(0, 3)
    .forEach((f, i) => {
      Logger.log((i + 1) + '. ' + f.nombre);
      Logger.log('   Actual:   ' + fmtEur(f.importeActual));
      Logger.log('   Var.dia:  ' + fmtEur(f.varDiaEur));
      Logger.log('   R.Total:  ' + fmtEur(f.rentBrutaEur) + ' (' + fmtPct(f.rentActualPct) + ')');
      Logger.log('   R.Fiscal: ' + fmtEur(f.rentFiscalEur) + ' (' + fmtPct(f.rentFiscalPct) + ')');
    });

  Logger.log('--- Mercado ---');
  const mkt = fetchMarketData();
  Logger.log('IBEX:  ' + (mkt.ibex   ? fmtNum(mkt.ibex.precio,   0) : 'N/D'));
  Logger.log('SP500: ' + (mkt.sp500  ? fmtNum(mkt.sp500.precio,  0) : 'N/D'));
  Logger.log('Brent: ' + (mkt.brent  ? fmtNum(mkt.brent.precio,  1) : 'N/D'));

  Logger.log('--- RSS ---');
  const news = fetchMarketNews();
  Logger.log('Noticias encontradas: ' + news.length);
  news.forEach((n, i) => Logger.log((i + 1) + '. ' + n));

  Logger.log('--- API Anthropic ---');
  const an = callAnthropicAPI(cfg, p, mkt, news);
  Logger.log('Analisis: ' + an.analisis.substring(0, 120) + '...');
  Logger.log('Escenarios: ' + an.escenarios.map(e => e.nombre + ' ' + e.probabilidad + '%').join(' | '));

  Logger.log('=== TEST OK v1.5 ===');
}
