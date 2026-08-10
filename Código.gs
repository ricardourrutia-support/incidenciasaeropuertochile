// ============================================================
//  CONFIGURACIÓN
// ============================================================
const CONFIG = {
  TABLEAU_HOST:             "https://tableau.cabify-data.com",
  TABLEAU_API_VERSION:      "3.21",
  TABLEAU_SITE:             "cabify",
  TABLEAU_PAT_NAME:         "B2B Support",
  TABLEAU_PAT_SECRET:       "KI/Fpf9RTVaA8F2oH/vqKQ==:kL07mNjXKGkJcnPJT4v4BjB0jXxmOFuI",
  TABLEAU_VIEW_ID:          "e155f9ad-e4cc-4ace-a1b1-f766edcfc6eb",
  TABLEAU_VIEW_CONTENT_URL: "CLB2BSupportCrossTab/sheets/Sheet1",

  CSAT_COLUMN:      "% CSAT",
  NPS_COLUMN:       "NPS Score",
  AGENT_COLUMN:     "Assignee FullName",
  TICKET_COLUMN:    "Ticket Number",
  WEEK_COLUMN:      "Week of Solved At Utc Dt",
  SOLVED_AT_COLUMN: "Solved At Utc Dt",
};

// ============================================================
//  WEB APP
// ============================================================
function doGet() {
  return HtmlService
    .createHtmlOutputFromFile("dashboard")
    .setTitle("CSAT & NPS Tracker B2B — Cabify")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
//  EXCLUSIÓN DE TICKETS
//  Tickets que cayeron en la bandeja B2B por error y no deben
//  contribuir a las métricas. Se persisten en Script Properties
//  (aplican para todos los usuarios y sobreviven al refresh).
// ============================================================
const EXCLUDED_PROP_KEY = "EXCLUDED_TICKETS_V1";

function getExcludedMap_() {
  const raw = PropertiesService.getScriptProperties().getProperty(EXCLUDED_PROP_KEY);
  if (!raw) return {};
  try { return JSON.parse(raw); } catch (e) { return {}; }
}

function saveExcludedMap_(map) {
  PropertiesService.getScriptProperties().setProperty(EXCLUDED_PROP_KEY, JSON.stringify(map));
}

/**
 * Excluye un ticket de las métricas. Guarda quién y cuándo.
 * Llamado desde el frontend tras la confirmación del usuario.
 */
function excludeTicket(ticketId) {
  const id = String(ticketId || "").trim();
  if (!id) throw new Error("Ticket inválido.");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const map = getExcludedMap_();
    map[id] = {
      by: (Session.getActiveUser() && Session.getActiveUser().getEmail()) || "desconocido",
      at: new Date().toISOString()
    };
    saveExcludedMap_(map);
    return { ok: true, ticket: id, by: map[id].by, at: map[id].at, excludedCount: Object.keys(map).length };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Restaura un ticket previamente excluido (vuelve a contar en métricas).
 */
function restoreTicket(ticketId) {
  const id = String(ticketId || "").trim();
  if (!id) throw new Error("Ticket inválido.");

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const map = getExcludedMap_();
    delete map[id];
    saveExcludedMap_(map);
    return { ok: true, ticket: id, excludedCount: Object.keys(map).length };
  } finally {
    lock.releaseLock();
  }
}

// ============================================================
//  API principal — devuelve todos los datos al frontend
// ============================================================
function getAllData() {
  const auth = getTableauToken();
  if (!auth) throw new Error("No se pudo autenticar con Tableau.");

  try {
    const rows = fetchCrosstab(auth);
    const excludedMap = getExcludedMap_();

    // ── Consolidar por ticket ──────────────────────────────
    // El crosstab exporta múltiples filas por ticket (una por medida).
    // FIX: al consolidar CSAT, el valor 1 (satisfecho) gana sobre 0.
    // Así una fila con CSAT vacío o 0 no bloquea un 1 que llega en
    // otra fila del mismo ticket.
    const ticketMap = {};

    rows.forEach(row => {
      const id = row[CONFIG.TICKET_COLUMN] || "";
      if (!ticketMap[id]) {
        ticketMap[id] = {
          ticket:     id,
          agent:      row[CONFIG.AGENT_COLUMN]           || "",
          week:       row[CONFIG.WEEK_COLUMN]            || "",
          solvedAt:   row[CONFIG.SOLVED_AT_COLUMN]       || "",
          tag1:       row["ES Output Tags 1st Level v2"] || "",
          tag2:       row["ES Output Tags 2nd Level v2"] || "",
          tag3:       row["ES Output Tags 3rd Level v2"] || "",
          autoAnswer: row["Abi/Auto Answer"]             || "",
          csat:       null,
          nps:        null,
          excluded:   !!excludedMap[String(id)],
          excludedBy: excludedMap[String(id)] ? excludedMap[String(id)].by : null,
          excludedAt: excludedMap[String(id)] ? excludedMap[String(id)].at : null,
        };
      }

      const t       = ticketMap[id];
      const rowCsat = parseCsat(row[CONFIG.CSAT_COLUMN]);
      const rowNps  = parseNps(row[CONFIG.NPS_COLUMN]);

      // CSAT: priorizar 1 (satisfecho). Si ya tenemos 1, no pisamos.
      // Si teníamos null o 0, y llega un 1 → actualizamos.
      if (rowCsat !== null) {
        if (t.csat === null) {
          t.csat = rowCsat;
        } else if (rowCsat === 1 && t.csat !== 1) {
          t.csat = 1;  // el 1 gana siempre
        }
      }

      // NPS: tomar el primer valor no nulo
      if (t.nps === null && rowNps !== null) {
        t.nps = rowNps;
      }

      // Completar campos de texto si aún están vacíos
      if (!t.week     && row[CONFIG.WEEK_COLUMN])       t.week     = row[CONFIG.WEEK_COLUMN];
      if (!t.solvedAt && row[CONFIG.SOLVED_AT_COLUMN])  t.solvedAt = row[CONFIG.SOLVED_AT_COLUMN];
      if (!t.agent    && row[CONFIG.AGENT_COLUMN])      t.agent    = row[CONFIG.AGENT_COLUMN];
      if (!t.tag1     && row["ES Output Tags 1st Level v2"]) t.tag1 = row["ES Output Tags 1st Level v2"];
      if (!t.tag2     && row["ES Output Tags 2nd Level v2"]) t.tag2 = row["ES Output Tags 2nd Level v2"];
      if (!t.tag3     && row["ES Output Tags 3rd Level v2"]) t.tag3 = row["ES Output Tags 3rd Level v2"];
    });

    const tickets = Object.values(ticketMap);

    // Solo los tickets NO excluidos contribuyen a las métricas.
    // Los excluidos igual se devuelven al frontend para poder
    // revisarlos y restaurarlos desde la pestaña Detalle.
    const activeTickets = tickets.filter(t => !t.excluded);

    // ── Agrupar por semana ──────────────────────────────────
    const byWeek = {};
    activeTickets.forEach(t => {
      if (!t.week) return;
      if (!byWeek[t.week]) byWeek[t.week] = { csatScores: [], npsScores: [], count: 0 };
      byWeek[t.week].count++;
      if (t.csat !== null) byWeek[t.week].csatScores.push(t.csat);
      if (t.nps  !== null) byWeek[t.week].npsScores.push(t.nps);
    });

    const weeks = Object.entries(byWeek)
      .map(([week, d]) => {
        const parsed = parseSpanishDate(week);
        return {
          week,
          weekNum:  parsed ? isoWeekNumber(parsed) : 0,
          weekYear: parsed ? parsed.getFullYear()  : 0,
          sortKey:  parsed ? parsed.getTime()       : 0,
          count:    d.count,
          avgCsat:  avg(d.csatScores),
          avgNps:   avg(d.npsScores),
        };
      })
      .sort((a, b) => a.sortKey - b.sortKey);

    // ── Agrupar por agente × semana ─────────────────────────
    const byAgent = {};
    activeTickets.forEach(t => {
      const name = t.agent || "Sin asignar";
      if (!byAgent[name]) byAgent[name] = { weeks: {}, totalCsat: [], totalNps: [], ticketCount: 0 };
      byAgent[name].ticketCount++;
      if (t.csat !== null) byAgent[name].totalCsat.push(t.csat);
      if (t.nps  !== null) byAgent[name].totalNps.push(t.nps);
      if (t.week) {
        if (!byAgent[name].weeks[t.week]) byAgent[name].weeks[t.week] = { csatScores: [], npsScores: [], count: 0 };
        byAgent[name].weeks[t.week].count++;
        if (t.csat !== null) byAgent[name].weeks[t.week].csatScores.push(t.csat);
        if (t.nps  !== null) byAgent[name].weeks[t.week].npsScores.push(t.nps);
      }
    });

    const weekLabels = weeks.map(w => w.week);
    const agents = Object.entries(byAgent)
      .map(([name, d]) => ({
        name,
        ticketCount: d.ticketCount,
        avgCsat:     avg(d.totalCsat),
        avgNps:      avg(d.totalNps),
        byWeek: weekLabels.map(w => {
          const wd = d.weeks[w];
          return wd
            ? { week: w, count: wd.count, avgCsat: avg(wd.csatScores), avgNps: avg(wd.npsScores) }
            : { week: w, count: 0, avgCsat: null, avgNps: null };
        }),
      }))
      .sort((a, b) => b.ticketCount - a.ticketCount);

    const withWeek  = activeTickets.filter(t => t.week);
    const sinSemana = activeTickets.filter(t => !t.week).length;

    return {
      updatedAt:  new Date().toLocaleString("es-MX"),
      tickets,                                            // TODOS (incluidos los excluidos, marcados con .excluded)
      excludedCount: tickets.length - activeTickets.length,
      weeks,
      weekLabels,
      agents,
      sinSemana,
      totals: {
        tickets: activeTickets.length,
        avgCsat: avg(withWeek.filter(t => t.csat !== null).map(t => t.csat)),
        avgNps:  avg(withWeek.filter(t => t.nps  !== null).map(t => t.nps)),
        agents:  new Set(activeTickets.map(t => t.agent).filter(Boolean)).size,
      }
    };
  } finally {
    signOutTableau(auth);
  }
}

// ============================================================
//  PARSERS
// ============================================================
function parseCsat(raw) {
  if (raw === null || raw === undefined || raw === "") return null;
  const n = parseFloat(String(raw).replace("%", "").trim());
  if (isNaN(n)) return null;
  // Tableau entrega 0 o 1 directamente (binario por ticket).
  // Si por alguna razón llegara en escala 0-100, normalizar.
  return n > 1 ? n / 100 : n;
}

function parseNps(raw) {
  if (raw === null || raw === undefined || raw === "") return null;
  const n = parseFloat(String(raw).replace(",", ".").trim());
  return isNaN(n) ? null : n;
}

function avg(arr) {
  if (!arr || !arr.length) return null;
  return arr.reduce((a, b) => a + b, 0) / arr.length;
}

// ============================================================
//  TABLEAU — Auth, Fetch, Parse, Sign Out
// ============================================================
function getTableauToken() {
  const url = `${CONFIG.TABLEAU_HOST}/api/${CONFIG.TABLEAU_API_VERSION}/auth/signin`;
  const res = UrlFetchApp.fetch(url, {
    method:      "post",
    contentType: "application/json",
    headers:     { "Accept": "application/json" },
    payload:     JSON.stringify({
      credentials: {
        personalAccessTokenName:   CONFIG.TABLEAU_PAT_NAME,
        personalAccessTokenSecret: CONFIG.TABLEAU_PAT_SECRET,
        site: { contentUrl: CONFIG.TABLEAU_SITE }
      }
    }),
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200) return null;
  const json = JSON.parse(res.getContentText());
  return { token: json.credentials.token, siteId: json.credentials.site.id };
}

function fetchCrosstab({ token, siteId }) {
  const url = `${CONFIG.TABLEAU_HOST}/api/${CONFIG.TABLEAU_API_VERSION}/sites/${siteId}/views/${CONFIG.TABLEAU_VIEW_ID}/data`;
  const res = UrlFetchApp.fetch(url, {
    method:  "get",
    headers: { "x-tableau-auth": token },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() !== 200)
    throw new Error(`Error ${res.getResponseCode()}: ${res.getContentText()}`);
  return parseCSV(res.getContentText());
}

function parseCSV(csvText) {
  const rows = Utilities.parseCsv(csvText);
  if (rows.length < 2) return [];
  const headers = rows[0];
  return rows.slice(1).map(row => {
    const obj = {};
    headers.forEach((h, i) => { obj[h.trim()] = (row[i] || "").trim(); });
    return obj;
  });
}

function signOutTableau({ token }) {
  UrlFetchApp.fetch(`${CONFIG.TABLEAU_HOST}/api/${CONFIG.TABLEAU_API_VERSION}/auth/signout`, {
    method:  "post",
    headers: { "x-tableau-auth": token },
    muteHttpExceptions: true
  });
}

// ============================================================
//  UTILIDADES DE FECHA
// ============================================================
const MESES_ES = {
  'enero':1,'febrero':2,'marzo':3,'abril':4,'mayo':5,'junio':6,
  'julio':7,'agosto':8,'septiembre':9,'octubre':10,'noviembre':11,'diciembre':12
};

function parseSpanishDate(str) {
  if (!str) return null;
  const m = str.toLowerCase().match(/(\d{1,2})\s+de\s+(\w+)\s+de\s+(\d{4})/);
  if (!m) return null;
  const mes = MESES_ES[m[2]];
  if (!mes) return null;
  return new Date(parseInt(m[3]), mes - 1, parseInt(m[1]));
}

function isoWeekNumber(date) {
  const d = new Date(Date.UTC(date.getFullYear(), date.getMonth(), date.getDate()));
  d.setUTCDate(d.getUTCDate() + 4 - (d.getUTCDay() || 7));
  const yearStart = new Date(Date.UTC(d.getUTCFullYear(), 0, 1));
  return Math.ceil((((d - yearStart) / 86400000) + 1) / 7);
}

// ============================================================
//  DEBUG — ejecuta una vez para ver columnas disponibles
// ============================================================
function debugColumns() {
  const auth = getTableauToken();
  if (!auth) return;
  try {
    const data = fetchCrosstab(auth);
    if (!data.length) { Logger.log("Sin datos"); return; }
    Logger.log("Columnas disponibles:");
    Object.keys(data[0]).forEach(col => Logger.log(`  "${col}" → "${data[0][col]}"`));
  } finally {
    signOutTableau(auth);
  }
}
