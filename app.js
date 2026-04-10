const STORAGE_KEY = "mlti-dashboard-upload-v1";
const quotesInput = document.getElementById("quotesFile");
const closuresInput = document.getElementById("closuresFile");
const quotesStatus = document.getElementById("quotesStatus");
const closuresStatus = document.getElementById("closuresStatus");
const feedback = document.getElementById("uploadFeedback");
const processBtn = document.getElementById("processBtn");
const clearBtn = document.getElementById("clearBtn");

const QUOTE_REQUIRED = ["date", "reference", "client", "service", "user", "status"];
const CLOSURE_REQUIRED = ["date", "client", "service"];

quotesInput?.addEventListener("change", () => updateFileBadge(quotesInput, quotesStatus, false));
closuresInput?.addEventListener("change", () => updateFileBadge(closuresInput, closuresStatus, true));
processBtn?.addEventListener("click", processFiles);
clearBtn?.addEventListener("click", clearSavedUpload);

function updateFileBadge(input, badge, optional) {
  if (!input.files || !input.files[0]) {
    badge.textContent = optional ? "Opcional" : "Sin archivo";
    badge.className = optional ? "status-pill warn" : "status-pill";
    return;
  }
  badge.textContent = input.files[0].name;
  badge.className = "status-pill ready";
}

function setFeedback(message, tone = "") {
  feedback.textContent = message;
  feedback.className = "upload-feedback" + (tone ? " " + tone : "");
}

async function processFiles() {
  const quoteFile = quotesInput.files && quotesInput.files[0];
  const closureFile = closuresInput.files && closuresInput.files[0];

  if (!quoteFile) {
    setFeedback("Necesito al menos el archivo de cotizaciones para recalcular el dashboard.", "error");
    return;
  }

  try {
    processBtn.disabled = true;
    setFeedback("Procesando archivos y recalculando métricas...");

    const quoteWorkbook = await readWorkbook(quoteFile);
    const quoteSheet = pickSheet(quoteWorkbook, "quotes");
    const quoteRows = XLSX.utils.sheet_to_json(quoteWorkbook.Sheets[quoteSheet], { header: 1, defval: "" });
    const quoteMap = detectColumns(quoteRows[0] || [], "quotes");
    validateColumns(quoteMap, QUOTE_REQUIRED, "cotizaciones");
    const quotes = normalizeQuotes(quoteRows.slice(1), quoteMap);

    let closures = [];
    if (closureFile) {
      const closureWorkbook = await readWorkbook(closureFile);
      const closureSheet = pickSheet(closureWorkbook, "closures");
      const closureRows = XLSX.utils.sheet_to_json(closureWorkbook.Sheets[closureSheet], { header: 1, defval: "" });
      const closureMap = detectColumns(closureRows[0] || [], "closures");
      validateColumns(closureMap, CLOSURE_REQUIRED, "cierres");
      closures = normalizeClosures(closureRows.slice(1), closureMap);
    }

    const summary = buildSummary(quotes, closures, quoteFile.name, closureFile ? closureFile.name : "");
    localStorage.setItem(STORAGE_KEY, JSON.stringify(summary));
    renderSummary(summary);
    setFeedback("Dashboard actualizado y guardado en este navegador.", "success");
  } catch (error) {
    console.error(error);
    setFeedback(error.message || "No pude procesar los archivos.", "error");
  } finally {
    processBtn.disabled = false;
  }
}

async function readWorkbook(file) {
  const buffer = await file.arrayBuffer();
  return XLSX.read(buffer, { type: "array", cellDates: true });
}

function pickSheet(workbook, type) {
  const names = workbook.SheetNames;
  const found = names.find((name) => {
    const key = normalize(name);
    return type === "quotes" ? /cotiz|quote|pricing/.test(key) : /cier|venta|close|embarque/.test(key);
  });
  return found || names[0];
}

function detectColumns(headers, type) {
  const rules = {
    date: ["dia", "fecha", "date"],
    reference: type === "quotes" ? ["referencia", "folio", "ref", "quote"] : ["referencia", "folio", "ref", "embarque", "shipment", "codigo"],
    client: ["cliente", "customer", "agent"],
    service: ["servicio", "service", "modo"],
    user: ["usuario", "ejecutivo", "responsable", "comercial", "owner", "agent"],
    status: ["estatus", "status", "estado"],
    network: ["network", "red"]
  };

  const map = {};
  headers.forEach((header, index) => {
    const key = normalize(header);
    Object.entries(rules).forEach(([field, words]) => {
      if (map[field] !== undefined) return;
      if (words.some((word) => key.includes(word))) map[field] = index;
    });
  });
  return map;
}

function validateColumns(map, required, label) {
  const missing = required.filter((field) => map[field] === undefined);
  if (missing.length) {
    throw new Error("Al archivo de " + label + " le faltan columnas detectables para: " + missing.join(", ") + ".");
  }
}

function normalizeQuotes(rows, map) {
  return rows
    .filter((row) => row.some((cell) => String(cell || "").trim() !== ""))
    .map((row, index) => ({
      reference: cell(row, map.reference) || "Q-" + index,
      client: label(row, map.client, "Sin cliente"),
      network: label(row, map.network, "Sin red"),
      services: splitServices(cell(row, map.service)),
      users: splitList(cell(row, map.user), "Sin usuario"),
      status: normalizeStatus(cell(row, map.status)),
      monthKey: monthKey(parseDate(row[map.date])),
      rawDate: row[map.date]
    }));
}

function normalizeClosures(rows, map) {
  return rows
    .filter((row) => row.some((cell) => String(cell || "").trim() !== ""))
    .map((row, index) => ({
      reference: cell(row, map.reference) || "C-" + index,
      client: label(row, map.client, "Sin cliente"),
      services: splitServices(cell(row, map.service)),
      users: splitList(cell(row, map.user), "Sin usuario"),
      monthKey: monthKey(parseDate(row[map.date]))
    }));
}

function buildSummary(quotes, closures, quoteFileName, closureFileName) {
  const monthKeys = [...new Set([...quotes, ...closures].map((item) => item.monthKey).filter(Boolean))].sort();
  const months = monthKeys.slice(-2);
  const [monthA, monthB] = months;
  const activeStatuses = new Set(["Cotizado Agentes", "Cotizado Pricing"]);
  const lossStatuses = new Set(["No Cotizada", "No Cotizado por Pricing", "Cotizado Pricing Fuera de Tiempo", "Cotizacion Cancelada"]);
  const internalStatuses = new Set(["No Cotizada", "No Cotizado por Pricing", "Cotizado Pricing Fuera de Tiempo"]);

  const quoteRefs = new Set(quotes.map((item) => item.reference));
  const totalQuotes = quoteRefs.size;
  const totalClosures = closures.length;
  const monthAQuotes = countUnique(quotes.filter((item) => item.monthKey === monthA), "reference");
  const monthBQuotes = countUnique(quotes.filter((item) => item.monthKey === monthB), "reference");
  const monthAClosures = closures.filter((item) => item.monthKey === monthA).length;
  const monthBClosures = closures.filter((item) => item.monthKey === monthB).length;
  const pipelineActive = countUnique(quotes.filter((item) => activeStatuses.has(item.status)), "reference");
  const operationalLosses = countUnique(quotes.filter((item) => lossStatuses.has(item.status)), "reference");
  const internalLosses = countUnique(quotes.filter((item) => internalStatuses.has(item.status)), "reference");

  const modeRows = buildServiceRows(quotes, closures, modeOf);
  const subtypeRows = buildServiceRows(quotes, closures, (value) => value);
  const ownershipRows = buildOwnershipRows(quotes);
  const clientRows = buildClientRows(quotes, closures);
  const teamRows = buildTeamRows(quotes, closures, monthA, monthB);
  const losses = buildLossRows(quotes, monthA, monthB);
  const trendServices = buildTrendServices(quotes, closures, monthA, monthB);
  const repeatedClients = buildRepeatedClients(closures, monthA, monthB);
  const clientDeltas = buildClientDeltas(closures, monthA, monthB);

  return {
    meta: {
      quoteFileName,
      closureFileName,
      monthA,
      monthB,
      monthALabel: monthLabel(monthA),
      monthBLabel: monthLabel(monthB),
      updatedAt: new Date().toISOString()
    },
    totals: {
      totalQuotes,
      totalClosures,
      monthAQuotes,
      monthBQuotes,
      monthAClosures,
      monthBClosures,
      globalHitRate: totalClosures / Math.max(totalQuotes, 1),
      monthAHitRate: monthAClosures / Math.max(monthAQuotes, 1),
      monthBHitRate: monthBClosures / Math.max(monthBQuotes, 1),
      pipelineActive,
      operationalLosses,
      internalLosses
    },
    modeRows,
    subtypeRows,
    ownershipRows,
    clientRows,
    teamRows,
    losses,
    trendServices,
    repeatedClients,
    clientDeltas
  };
}

function buildServiceRows(quotes, closures, mapper) {
  const q = new Map();
  const c = new Map();
  quotes.forEach((item) => item.services.forEach((service) => addSet(q, mapper(service), item.reference + "|" + service)));
  closures.forEach((item) => item.services.forEach((service) => addSet(c, mapper(service), item.reference + "|" + service)));
  return [...new Set([...q.keys(), ...c.keys()])]
    .map((name) => {
      const quotesCount = q.get(name) ? q.get(name).size : 0;
      const closuresCount = c.get(name) ? c.get(name).size : 0;
      return { name, quotes: quotesCount, closures: closuresCount, hitRate: closuresCount / Math.max(quotesCount, 1) };
    })
    .filter((row) => row.quotes || row.closures)
    .sort((a, b) => b.quotes - a.quotes || b.closures - a.closures);
}

function buildOwnershipRows(quotes) {
  const map = new Map();
  quotes.forEach((item) => {
    if (!["Cotizado Agentes", "Cotizado Pricing"].includes(item.status)) return;
    item.services.forEach((service) => {
      if (!map.has(service)) map.set(service, { agents: new Set(), pricing: new Set() });
      const bucket = item.status === "Cotizado Agentes" ? "agents" : "pricing";
      map.get(service)[bucket].add(item.reference);
    });
  });
  return [...map.entries()].map(([service, buckets]) => {
    const agents = buckets.agents.size;
    const pricing = buckets.pricing.size;
    const base = Math.max(agents + pricing, 1);
    return { service, agentsShare: agents / base, pricingShare: pricing / base };
  }).sort((a, b) => a.service.localeCompare(b.service, "es"));
}

function buildClientRows(quotes, closures) {
  const q = new Map();
  const c = new Map();
  quotes.forEach((item) => addSet(q, item.client, item.reference));
  closures.forEach((item) => addSet(c, item.client, item.reference));
  return [...q.keys()].map((name) => {
    const quotesCount = q.get(name).size;
    const closuresCount = c.get(name) ? c.get(name).size : 0;
    return { name, quotes: quotesCount, closures: closuresCount, hitRate: closuresCount / Math.max(quotesCount, 1), flag: clientFlag(quotesCount, closuresCount) };
  }).sort((a, b) => b.quotes - a.quotes || b.closures - a.closures);
}

function buildTeamRows(quotes, closures, monthA, monthB) {
  const users = new Set([...quotes.flatMap((item) => item.users), ...closures.flatMap((item) => item.users)]);
  return [...users].map((name) => {
    const quoteRefs = new Set();
    const noQuoted = new Set();
    const noPricing = new Set();
    quotes.forEach((item) => {
      if (!item.users.includes(name)) return;
      quoteRefs.add(item.reference);
      if (item.status === "No Cotizada") noQuoted.add(item.reference);
      if (item.status === "No Cotizado por Pricing") noPricing.add(item.reference);
    });
    const closuresCount = closures.filter((item) => item.users.includes(name)).length;
    return {
      name,
      quotes: quoteRefs.size,
      closures: closuresCount,
      hitRate: closuresCount / Math.max(quoteRefs.size, 1),
      noQuoted: noQuoted.size,
      noPricing: noPricing.size,
      monthAQuotes: countUnique(quotes.filter((item) => item.monthKey === monthA && item.users.includes(name)), "reference"),
      monthBQuotes: countUnique(quotes.filter((item) => item.monthKey === monthB && item.users.includes(name)), "reference")
    };
  }).filter((row) => row.quotes || row.closures).sort((a, b) => b.quotes - a.quotes || b.closures - a.closures);
}

function buildLossRows(quotes, monthA, monthB) {
  const labels = [
    ["No Cotizada", "Equipo"],
    ["No Cotizado por Pricing", "Pricing"],
    ["Cotizacion Cancelada", "Cliente o info"],
    ["Cotizado Pricing Fuera de Tiempo", "Pricing"]
  ];
  return labels.map(([cause, owner]) => ({
    cause,
    owner,
    monthA: countUnique(quotes.filter((item) => item.status === cause && item.monthKey === monthA), "reference"),
    monthB: countUnique(quotes.filter((item) => item.status === cause && item.monthKey === monthB), "reference")
  })).map((row) => ({ ...row, total: row.monthA + row.monthB })).filter((row) => row.total > 0);
}

function buildTrendServices(quotes, closures, monthA, monthB) {
  const names = [...new Set([...quotes.flatMap((item) => item.services), ...closures.flatMap((item) => item.services)])];
  return names.map((service) => ({
    service,
    monthAQuotes: countUnique(quotes.filter((item) => item.monthKey === monthA && item.services.includes(service)), "reference"),
    monthBQuotes: countUnique(quotes.filter((item) => item.monthKey === monthB && item.services.includes(service)), "reference"),
    monthBClosures: closures.filter((item) => item.monthKey === monthB && item.services.includes(service)).length
  })).filter((row) => row.monthAQuotes || row.monthBQuotes || row.monthBClosures)
    .map((row) => ({ ...row, delta: row.monthBQuotes - row.monthAQuotes }))
    .sort((a, b) => (b.monthAQuotes + b.monthBQuotes) - (a.monthAQuotes + a.monthBQuotes))
    .slice(0, 12);
}

function buildRepeatedClients(closures, monthA, monthB) {
  const names = [...new Set(closures.map((item) => item.client))];
  return names.map((name) => ({
    name,
    monthA: closures.filter((item) => item.client === name && item.monthKey === monthA).length,
    monthB: closures.filter((item) => item.client === name && item.monthKey === monthB).length
  })).filter((row) => row.monthA > 0 && row.monthB > 0)
    .sort((a, b) => (b.monthA + b.monthB) - (a.monthA + a.monthB))
    .slice(0, 10);
}

function buildClientDeltas(closures, monthA, monthB) {
  const names = [...new Set(closures.map((item) => item.client))];
  return names.map((name) => ({
    name,
    monthA: closures.filter((item) => item.client === name && item.monthKey === monthA).length,
    monthB: closures.filter((item) => item.client === name && item.monthKey === monthB).length
  })).filter((row) => row.monthA || row.monthB)
    .sort((a, b) => (b.monthB - b.monthA) - (a.monthB - a.monthA))
    .slice(0, 8);
}

function renderSummary(data) {
  renderOverview(data);
  renderServices(data);
  renderClients(data);
  renderTeam(data);
  renderLosses(data);
  renderTrend(data);
}

function renderOverview(data) {
  const m = data.meta;
  const t = data.totals;
  document.getElementById("resumen").innerHTML = `
    <div class="cards">
      ${card("Referencias válidas", int(t.totalQuotes), m.quoteFileName)}
      ${card(m.monthALabel, int(t.monthAQuotes), "Cotizaciones", "orange")}
      ${card(m.monthBLabel, int(t.monthBQuotes), "Cotizaciones", "orange")}
      ${card("Cierres reales", int(t.totalClosures), m.closureFileName || "Sin archivo de cierres", "green")}
      ${card("Hit Rate global", pct(t.globalHitRate), int(t.totalClosures) + " cierres sobre " + int(t.totalQuotes), "blue")}
      ${card("Pipeline activo", int(t.pipelineActive), "Estatus cotizados", "yellow")}
      ${card("Pérdidas operativas", int(t.operationalLosses), int(t.internalLosses) + " por falla interna", "red")}
    </div>
    <div class="insight"><strong>Actualización dinámica:</strong> este dashboard ya se recalcula desde Excel y queda guardado en este navegador.</div>
  `;
}

function renderServices(data) {
  document.getElementById("servicios").innerHTML = `
    <div class="grid2">
      <div class="panel">
        <div class="panel-title"><div class="dot dot-blue"></div>Hit Rate por Modo</div>
        ${table(["Modo","Cotizaciones","Cierres","Hit Rate"], data.modeRows.map((row) => [row.name, int(row.quotes), int(row.closures), pct(row.hitRate)]))}
      </div>
      <div class="panel">
        <div class="panel-title"><div class="dot dot-green"></div>Desglose por Subtipo</div>
        ${table(["Subtipo","Cotizaciones","Cierres","Hit Rate"], data.subtypeRows.map((row) => [row.name, int(row.quotes), int(row.closures), pct(row.hitRate)]))}
      </div>
    </div>
    <div class="panel panel-full">
      <div class="panel-title"><div class="dot dot-orange"></div>Quién está cotizando los servicios activos</div>
      ${table(["Servicio","Agentes","Pricing"], data.ownershipRows.map((row) => [row.service, pct(row.agentsShare), pct(row.pricingShare)]), 3)}
    </div>
  `;
}

function renderClients(data) {
  const top = data.clientRows.slice(0, 12);
  const risk = data.clientRows.filter((row) => row.flag !== "Alto").slice(0, 8);
  document.getElementById("clientes").innerHTML = `
    <div class="grid2">
      <div class="panel">
        <div class="panel-title"><div class="dot dot-green"></div>Clientes Top</div>
        ${table(["Cliente","Cot.","Cierres","Hit Rate","Semáforo"], top.map((row) => [row.name, int(row.quotes), int(row.closures), pct(row.hitRate), row.flag]), 4)}
      </div>
      <div class="panel">
        <div class="panel-title"><div class="dot dot-red"></div>Clientes en Observación</div>
        ${table(["Cliente","Cot.","Cierres","Hit Rate","Semáforo"], risk.map((row) => [row.name, int(row.quotes), int(row.closures), pct(row.hitRate), row.flag]), 4)}
      </div>
    </div>
  `;
}

function renderTeam(data) {
  document.getElementById("equipo").innerHTML = `
    <div class="grid2">
      <div class="panel">
        <div class="panel-title"><div class="dot dot-blue"></div>Eficiencia por Ejecutivo</div>
        ${table(["Ejecutivo","Cot.","Cierres","Hit Rate","No Cot.","No Pricing"], data.teamRows.map((row) => [row.name, int(row.quotes), int(row.closures), pct(row.hitRate), int(row.noQuoted), int(row.noPricing)]), 5)}
      </div>
      <div class="panel">
        <div class="panel-title"><div class="dot dot-orange"></div>Volumen por Mes</div>
        ${table(["Ejecutivo",data.meta.monthALabel,data.meta.monthBLabel,"Cambio"], data.teamRows.map((row) => [row.name, int(row.monthAQuotes), int(row.monthBQuotes), signed(row.monthBQuotes - row.monthAQuotes)]), 3)}
      </div>
    </div>
  `;
}

function renderLosses(data) {
  document.getElementById("perdidas").innerHTML = `
    <div class="grid2">
      <div class="panel">
        <div class="panel-title"><div class="dot dot-red"></div>Pérdidas Operativas</div>
        ${table(["Causa",data.meta.monthALabel,data.meta.monthBLabel,"Total","Quién falla"], data.losses.map((row) => [row.cause, int(row.monthA), int(row.monthB), int(row.total), row.owner]), 3)}
      </div>
      <div class="panel">
        <div class="panel-title"><div class="dot dot-orange"></div>Impacto Estimado</div>
        <div class="stat-row"><div class="stat-key">Pérdidas totales</div><div class="stat-val">${int(data.totals.operationalLosses)}</div></div>
        <div class="stat-row"><div class="stat-key">Falla interna</div><div class="stat-val">${int(data.totals.internalLosses)}</div></div>
        <div class="stat-row"><div class="stat-key">Hit rate actual</div><div class="stat-val">${pct(data.totals.globalHitRate)}</div></div>
        <div class="stat-row"><div class="stat-key">Cierres potenciales</div><div class="stat-val">~${int(Math.round(data.totals.internalLosses * data.totals.globalHitRate))}</div></div>
      </div>
    </div>
  `;
}

function renderTrend(data) {
  document.getElementById("tendencia").innerHTML = `
    <div class="cards">
      ${card("Cot. " + data.meta.monthALabel, int(data.totals.monthAQuotes), "Base", "orange")}
      ${card("Cot. " + data.meta.monthBLabel, int(data.totals.monthBQuotes), signed(data.totals.monthBQuotes - data.totals.monthAQuotes), "orange")}
      ${card("HR " + data.meta.monthALabel, pct(data.totals.monthAHitRate), "Conversión", "blue")}
      ${card("HR " + data.meta.monthBLabel, pct(data.totals.monthBHitRate), signedPct(data.totals.monthBHitRate - data.totals.monthAHitRate), "green")}
    </div>
    <div class="grid2">
      <div class="panel">
        <div class="panel-title"><div class="dot dot-orange"></div>Mix de Servicios</div>
        ${table(["Servicio",data.meta.monthALabel,data.meta.monthBLabel,"Δ Cot.","Cierres " + data.meta.monthBLabel], data.trendServices.map((row) => [row.service, int(row.monthAQuotes), int(row.monthBQuotes), signed(row.delta), int(row.monthBClosures)]), 4)}
      </div>
      <div class="panel">
        <div class="panel-title"><div class="dot dot-green"></div>Clientes con Cierres en Ambos Meses</div>
        ${table(["Cliente",data.meta.monthALabel,data.meta.monthBLabel,"Tendencia"], data.repeatedClients.map((row) => [row.name, int(row.monthA), int(row.monthB), row.monthB > row.monthA ? "Subió" : row.monthB < row.monthA ? "Bajó" : "Estable"]))}
      </div>
    </div>
    <div class="panel panel-full">
      <div class="panel-title"><div class="dot dot-purple"></div>Clientes que Más Cambiaron</div>
      ${table(["Cliente",data.meta.monthALabel,data.meta.monthBLabel,"Δ Cierres"], data.clientDeltas.map((row) => [row.name, int(row.monthA), int(row.monthB), signed(row.monthB - row.monthA)]), 3)}
    </div>
  `;
}

function card(label, value, sub, color) {
  const tone = color ? " " + color : "";
  return `<div class="card${tone}"><div class="card-label">${escapeHtml(label)}</div><div class="card-value${tone}">${escapeHtml(value)}</div><div class="card-sub">${escapeHtml(sub)}</div></div>`;
}

function table(headers, rows, numericFrom = 1) {
  if (!rows.length) return `<div class="tbl-scroll"><table class="tbl"><tr><td>Sin datos suficientes para esta vista.</td></tr></table></div>`;
  return `<div class="tbl-scroll"><table class="tbl"><tr>${headers.map((header, index) => `<th${index >= numericFrom ? ` class="num"` : ""}>${escapeHtml(header)}</th>`).join("")}</tr>${rows.map((row) => `<tr>${row.map((cell, index) => `<td${index >= numericFrom ? ` class="num"` : ""}>${badgeIfNeeded(cell)}</td>`).join("")}</tr>`).join("")}</table></div>`;
}

function badgeIfNeeded(value) {
  if (["Alto","Medio","Bajo","Crítico","Critico","Subió","Bajó","Estable"].includes(value)) {
    const cls = value === "Alto" || value === "Subió" ? "badge-green" : value === "Medio" || value === "Estable" ? "badge-yellow" : "badge-red";
    return `<span class="badge ${cls}">${escapeHtml(value)}</span>`;
  }
  return escapeHtml(value);
}

function clientFlag(quotesCount, closuresCount) {
  const hr = closuresCount / Math.max(quotesCount, 1);
  if (quotesCount >= 15 && hr <= 0.05) return "Crítico";
  if (hr >= 0.5) return "Alto";
  if (hr >= 0.2) return "Medio";
  return "Bajo";
}

function normalize(value) {
  return String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9]+/g, " ").trim();
}

function parseDate(value) {
  if (value instanceof Date) return value;
  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
  }
  const text = String(value || "").trim();
  const parts = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
  if (parts) {
    const year = Number(parts[3]) < 100 ? 2000 + Number(parts[3]) : Number(parts[3]);
    return new Date(year, Number(parts[2]) - 1, Number(parts[1]));
  }
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? null : parsed;
}

function monthKey(date) {
  if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
  return date.getFullYear() + "-" + String(date.getMonth() + 1).padStart(2, "0");
}

function monthLabel(key) {
  if (!key) return "Sin mes";
  const [year, month] = key.split("-");
  return new Date(Number(year), Number(month) - 1, 1).toLocaleDateString("es-MX", { month: "long", year: "numeric" });
}

function cell(row, index) {
  return index === undefined ? "" : String(row[index] == null ? "" : row[index]).trim();
}

function label(row, index, fallback) {
  return cell(row, index) || fallback;
}

function splitList(value, fallback) {
  const text = String(value || "").trim();
  if (!text) return [fallback];
  return [...new Set(text.replace(/\s+y\s+/gi, "|").replace(/[;,/\\]+/g, "|").split("|").map((item) => item.trim()).filter(Boolean))];
}

function splitServices(value) {
  const text = String(value || "").trim();
  if (!text) return ["Sin clasificar"];
  const pieces = text.replace(/[;,]+/g, "|").replace(/\s+\/\s+/g, "|").split("|").map((item) => item.trim()).filter(Boolean);
  const services = [];
  pieces.forEach((piece) => normalizeService(piece).forEach((item) => services.push(item)));
  return [...new Set(services.length ? services : ["Sin clasificar"])];
}

function normalizeService(value) {
  const key = normalize(value);
  const found = [];
  if (key.includes("free hand")) found.push("Free Hand");
  if (/d2d|otr|terrestre/.test(key)) found.push("D2D OTR");
  if (/dap/.test(key) && /marit|ocean|sea/.test(key)) found.push("DAP Maritimo");
  if (/exw/.test(key) && /marit|ocean|sea/.test(key)) found.push("EXW Maritimo");
  if (/dap/.test(key) && /aer|air/.test(key)) found.push("DAP Aereo");
  if (/exw/.test(key) && /aer|air/.test(key)) found.push("EXW Aereo");
  if (!found.length && /marit|ocean|sea/.test(key)) found.push("Maritimo");
  if (!found.length && /aer|air/.test(key)) found.push("Aereo");
  return found.length ? found : [String(value || "").trim() || "Sin clasificar"];
}

function normalizeStatus(value) {
  const key = normalize(value);
  if (key.includes("fuera de tiempo") && key.includes("pricing")) return "Cotizado Pricing Fuera de Tiempo";
  if ((key.includes("no cotizado") || key.includes("no cotizada")) && key.includes("pricing")) return "No Cotizado por Pricing";
  if (key.includes("no cotizado") || key.includes("no cotizada")) return "No Cotizada";
  if (key.includes("cancel")) return "Cotizacion Cancelada";
  if (key.includes("cerrado") || key.includes("ganado") || key.includes("closed")) return "Cerrado";
  if (key.includes("cotizado") && key.includes("agente")) return "Cotizado Agentes";
  if (key.includes("cotizado") && key.includes("pricing")) return "Cotizado Pricing";
  return String(value || "").trim() || "Otro";
}

function modeOf(service) {
  const key = normalize(service);
  if (key.includes("free hand")) return "Free Hand";
  if (key.includes("otr") || key.includes("d2d")) return "OTR";
  if (key.includes("marit")) return "Maritimo";
  if (key.includes("aereo")) return "Aereo";
  return "Sin clasificar";
}

function addSet(map, key, value) {
  if (!map.has(key)) map.set(key, new Set());
  map.get(key).add(value);
}

function countUnique(items, field) {
  return new Set(items.map((item) => item[field])).size;
}

function int(value) { return Number(value || 0).toLocaleString("es-MX"); }
function pct(value) { return (Number(value || 0) * 100).toFixed(1) + "%"; }
function signed(value) { const n = Number(value || 0); return (n > 0 ? "+" : "") + n.toLocaleString("es-MX"); }
function signedPct(value) { const n = (Number(value || 0) * 100).toFixed(1); return (value >= 0 ? "+" : "") + n + " pp"; }
function escapeHtml(value) { return String(value == null ? "" : value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;"); }

function clearSavedUpload() {
  localStorage.removeItem(STORAGE_KEY);
  setFeedback("Se borró la carga guardada de este navegador.", "success");
}

try {
  const saved = localStorage.getItem(STORAGE_KEY);
  if (saved) {
    renderSummary(JSON.parse(saved));
    setFeedback("Se cargó la última versión guardada en este navegador.", "success");
  }
} catch (error) {
  console.error(error);
}
