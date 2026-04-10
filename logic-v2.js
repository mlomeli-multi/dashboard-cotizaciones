// Dashboard logic v2: real sales closures vs client-month quote matching.
(() => {
  const STORAGE_KEY_V2 = "mlti-dashboard-upload-v2";
  const SALES_OWNERS = new Set([
    "Rodrigo Alanis",
    "Luz Adriana Calatrava",
    "Brenda Rodriguez",
    "Agentes Internacionales",
    "Joselyn Valdez"
  ]);
  const SALES_VALID_STATUSES = new Set(["embarque finalizado", "entregado"]);
  const ACTIVE_SHIPMENT_MODES = {
    IA: "Aereo",
    EA: "Aereo",
    IH: "Aereo",
    EH: "Aereo",
    DH: "Aereo",
    DA: "Aereo",
    IC: "Aereo",
    EC: "Aereo",
    DC: "Aereo",
    AA: "Aereo",
    EM: "Maritimo",
    IM: "Maritimo",
    AM: "Maritimo",
    DT: "OTR",
    ET: "OTR",
    IT: "OTR",
    WH: "OTR",
    DPA: "Aereo",
    EPA: "Aereo",
    IPA: "Aereo",
    APM: "Maritimo",
    EPM: "Maritimo",
    IPM: "Maritimo",
    DPT: "OTR",
    EPT: "OTR",
    IPT: "OTR"
  };
  const INACTIVE_SHIPMENT_PREFIXES = new Set(["DM", "DP", "EP", "IP"]);
  const CLIENT_ALIASES = [
    { canonical: "BARSAN", patterns: [/barsan/] },
    { canonical: "BSI", patterns: [/^bsi$/, /best services international/] },
    { canonical: "QUICK", patterns: [/^quick$/, /quick\s*\(?k\s*&?\s*n/] },
    { canonical: "WORLD CARGO", patterns: [/world cargo/] },
    { canonical: "LOXSON", patterns: [/loxson/] },
    { canonical: "VERSANT", patterns: [/versant/] },
    { canonical: "TIME MATTERS", patterns: [/time matters/] },
    { canonical: "PRIORITY FREIGHT", patterns: [/priority freight/] },
    { canonical: "ANCHOR", patterns: [/anchor/] },
    { canonical: "EAS", patterns: [/^eas$/, /eas\s/] },
    { canonical: "SHANGHAI YIHENG", patterns: [/shanghai yiheng/] },
    { canonical: "ZENCARGO", patterns: [/zencargo/, /worldwide freight logistics limited/] },
    { canonical: "PANAMA COLD AGENCY", patterns: [/panama cold agency/] },
    { canonical: "GLOBAL LINER AGENCIES", patterns: [/global line?r agencies/] },
    { canonical: "UNIEXPRESS", patterns: [/uniexpress/] },
    { canonical: "SHARE LOGISTICS", patterns: [/share logistics/] },
    { canonical: "SUNCARGO", patterns: [/suncargo/] }
  ];
  const LEGAL_TOKENS = new Set([
    "sa", "de", "cv", "s", "rl", "srl", "ltd", "limited", "co", "company", "inc", "corp",
    "corporation", "llc", "logistics", "logistic", "freight", "forwarding", "global",
    "mexico", "services", "service", "international", "internacional", "head", "office", "t", "a"
  ]);
  const QUOTE_REQUIRED = ["date", "reference", "client", "service", "user", "status"];
  const SALES_REQUIRED = ["client", "owner"];

  const refs = rebindDom();
  const { quotesInput, closuresInput, quotesStatus, closuresStatus, feedback, processBtn, clearBtn } = refs;

  quotesInput.addEventListener("change", () => updateFileBadge(quotesInput, quotesStatus, false));
  closuresInput.addEventListener("change", () => updateFileBadge(closuresInput, closuresStatus, true));
  processBtn.addEventListener("click", processFiles);
  clearBtn.addEventListener("click", clearSavedUpload);

  tryLoadSaved();

  function rebindDom() {
    const ids = ["quotesFile", "closuresFile", "processBtn", "clearBtn"];
    ids.forEach((id) => {
      const original = document.getElementById(id);
      const clone = original.cloneNode(true);
      original.replaceWith(clone);
    });
    return {
      quotesInput: document.getElementById("quotesFile"),
      closuresInput: document.getElementById("closuresFile"),
      quotesStatus: document.getElementById("quotesStatus"),
      closuresStatus: document.getElementById("closuresStatus"),
      feedback: document.getElementById("uploadFeedback"),
      processBtn: document.getElementById("processBtn"),
      clearBtn: document.getElementById("clearBtn")
    };
  }

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
    const salesFile = closuresInput.files && closuresInput.files[0];
    if (!quoteFile) {
      setFeedback("Necesito al menos el archivo de cotizaciones para recalcular el dashboard.", "error");
      return;
    }
    try {
      processBtn.disabled = true;
      setFeedback("Procesando archivos y aplicando el cruce cliente-mes...");
      const quoteWorkbook = await readWorkbook(quoteFile);
      const quotes = normalizeQuoteWorkbook(quoteWorkbook);
      let salesRows = [];
      if (salesFile) {
        const salesWorkbook = await readWorkbook(salesFile);
        salesRows = normalizeSalesWorkbook(salesWorkbook);
      }
      const summary = buildSummary(quotes, salesRows, quoteFile.name, salesFile ? salesFile.name : "");
      localStorage.setItem(STORAGE_KEY_V2, JSON.stringify(summary));
      renderSummaryV2(summary);
      setFeedback(`Archivos leídos. Cotizaciones: ${int(summary.totals.totalQuotes)} | Cierres reales: ${int(summary.totals.totalRealClosures)} | Hit rate: ${pct(summary.totals.globalHitRate)}`, "success");
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

  function normalizeQuoteWorkbook(workbook) {
    const quotes = [];
    workbook.SheetNames.forEach((sheetName) => {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
      if (!rows.length) return;
      const map = detectColumns(rows[0], "quotes");
      if (!hasRequiredColumns(map, QUOTE_REQUIRED)) return;
      quotes.push(...normalizeQuotes(rows.slice(1), map, sheetName));
    });
    if (!quotes.length) {
      throw new Error("No encontré hojas válidas de cotizaciones. Revisa que existan DIA, Referencia, Cliente, Servicio, Usuario y Estatus.");
    }
    return quotes;
  }

  function normalizeSalesWorkbook(workbook) {
    for (const sheetName of workbook.SheetNames) {
      const rows = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: "" });
      if (!rows.length) continue;
      const map = detectColumns(rows[0], "sales");
      if (!hasRequiredColumns(map, SALES_REQUIRED)) continue;
      const sales = normalizeSalesRows(rows.slice(1), map);
      if (sales.length) return sales;
    }
    throw new Error("No pude encontrar en el reporte de ventas las columnas de Cliente y Dueño Cuenta.");
  }

  function detectColumns(headers, type) {
    const rules = {
      quotes: {
        date: ["dia", "fecha"],
        reference: ["referencia", "ref", "folio"],
        client: ["cliente"],
        network: ["network", "red"],
        service: ["servicio", "service"],
        user: ["usuario", "ejecutivo", "responsable"],
        status: ["estatus", "estado"]
      },
      sales: {
        client: ["cliente"],
        owner: ["dueno cuenta", "dueño cuenta", "owner"],
        shipment: ["embarque"],
        quoteCode: ["cotizacion"],
        status: ["estatus de embarque", "status embarque"],
        mode: ["modalidad"],
        service: ["tipo de servicio", "servicio"],
        finalDate: ["fecha embarque finalizado"],
        invoiceDate: ["fecha embarque facturado"],
        createdDate: ["fecha creacion embarque", "fecha creación embarque"]
      }
    };
    const map = {};
    headers.forEach((header, index) => {
      const normalizedHeader = normalize(header);
      Object.entries(rules[type]).forEach(([field, words]) => {
        if (map[field] !== undefined) return;
        if (words.some((word) => normalizedHeader.includes(word))) map[field] = index;
      });
    });
    return map;
  }

  function hasRequiredColumns(map, required) {
    return required.every((field) => map[field] !== undefined);
  }

  function normalizeQuotes(rows, map, sheetName) {
    return rows
      .filter((row) => [map.date, map.client, map.service, map.user, map.status].every((index) => cell(row, index) !== ""))
      .map((row, index) => {
        const date = parseDate(row[map.date]);
        const services = splitServices(cell(row, map.service));
        return {
          id: `${sheetName}:${cell(row, map.reference) || index + 1}`,
          sheetName,
          reference: cell(row, map.reference) || `${sheetName}-${index + 1}`,
          client: label(row, map.client, "Sin cliente"),
          canonicalClient: canonicalClient(label(row, map.client, "Sin cliente")),
          network: label(row, map.network, "Sin red"),
          services,
          modes: unique(services.map(modeOf)),
          users: splitList(cell(row, map.user), "Sin usuario"),
          status: normalizeStatus(cell(row, map.status)),
          monthKey: monthKey(date),
          quoteDate: date
        };
      })
      .filter((row) => row.monthKey);
  }

  function normalizeSalesRows(rows, map) {
    return rows
      .filter((row) => row.some((value) => String(value || "").trim() !== ""))
      .map((row, index) => {
        const closureDate = parseDate(row[map.finalDate]) || parseDate(row[map.invoiceDate]) || parseDate(row[map.createdDate]);
        const shipment = cell(row, map.shipment);
        const shipmentMode = shipmentModeFromCode(shipment);
        return {
          id: cell(row, map.shipment) || cell(row, map.quoteCode) || `sale-${index + 1}`,
          owner: label(row, map.owner, ""),
          client: label(row, map.client, "Sin cliente"),
          canonicalClient: canonicalClient(label(row, map.client, "Sin cliente")),
          shipment,
          quoteCode: cell(row, map.quoteCode),
          mode: shipmentMode || normalizeSalesMode(cell(row, map.mode), cell(row, map.service)),
          statusText: normalize(cell(row, map.status)),
          monthKey: monthKey(closureDate),
          closureDate
        };
      })
      .filter((row) => row.monthKey && SALES_OWNERS.has(row.owner) && SALES_VALID_STATUSES.has(row.statusText))
      .filter((row) => {
        const prefix = shipmentPrefix(row.shipment);
        return !prefix || !INACTIVE_SHIPMENT_PREFIXES.has(prefix);
      });
  }

  function buildSummary(quotes, salesRows, quoteFileName, salesFileName) {
    const quoteRefs = dedupeQuotes(quotes);
    const assignments = assignConvertedQuotes(quoteRefs, salesRows);
    const convertedIds = new Set(assignments.map((item) => item.quote.id));
    const matchedSalesIds = new Set(assignments.map((item) => item.sale.id));
    const convertedQuotes = quoteRefs.filter((quote) => convertedIds.has(quote.id));
    const unmatchedSales = salesRows.filter((sale) => !matchedSalesIds.has(sale.id));
    const months = buildMonthList(quoteRefs, salesRows);
    const latest = months.slice(-2);
    const monthA = latest[0] || null;
    const monthB = latest[1] || latest[0] || null;
    const activeStatuses = new Set(["Cotizado Agentes", "Cotizado Pricing"]);
    const lossStatuses = new Set(["No Cotizada", "No Cotizado por Pricing", "Cotizado Pricing Fuera de Tiempo", "Cotizacion Cancelada"]);
    const internalStatuses = new Set(["No Cotizada", "No Cotizado por Pricing", "Cotizado Pricing Fuera de Tiempo"]);

    const monthRows = months.map((month) => {
      const monthQuotes = quoteRefs.filter((quote) => quote.monthKey === month.key);
      const monthSales = salesRows.filter((sale) => sale.monthKey === month.key);
      const monthConverted = convertedQuotes.filter((quote) => quote.monthKey === month.key);
      const monthSalesOnly = unmatchedSales.filter((sale) => sale.monthKey === month.key);
      return {
        key: month.key,
        label: month.label,
        quotes: monthQuotes.length,
        realClosures: monthSales.length,
        matchedClosures: monthConverted.length,
        salesOnlyClosures: monthSalesOnly.length,
        hitRate: monthConverted.length / Math.max(monthQuotes.length, 1)
      };
    });

    return {
      meta: { quoteFileName, salesFileName, updatedAt: new Date().toISOString(), months, monthA, monthB },
      totals: {
        totalQuotes: quoteRefs.length,
        totalRealClosures: salesRows.length,
        totalMatchedClosures: convertedQuotes.length,
        totalSalesOnlyClosures: unmatchedSales.length,
        globalHitRate: convertedQuotes.length / Math.max(quoteRefs.length, 1),
        pipelineActive: quoteRefs.filter((quote) => activeStatuses.has(quote.status)).length,
        operationalLosses: quoteRefs.filter((quote) => lossStatuses.has(quote.status)).length,
        internalLosses: quoteRefs.filter((quote) => internalStatuses.has(quote.status)).length
      },
      monthRows,
      modeRows: buildServiceRows(quoteRefs, convertedQuotes, (service) => modeOf(service)),
      subtypeRows: buildServiceRows(quoteRefs, convertedQuotes, (service) => service),
      ownershipRows: buildOwnershipRows(quoteRefs),
      clientRows: buildClientRows(quoteRefs, salesRows, convertedQuotes),
      teamRows: buildTeamRows(quoteRefs, convertedQuotes, monthA && monthA.key, monthB && monthB.key),
      losses: buildLossRows(quoteRefs, months),
      trendServices: buildTrendServiceRows(quoteRefs, convertedQuotes, monthA && monthA.key, monthB && monthB.key),
      repeatedClients: buildRepeatedClients(salesRows, monthA && monthA.key, monthB && monthB.key),
      clientDeltas: buildClientDeltas(salesRows, monthA && monthA.key, monthB && monthB.key),
      salesOnlyByMonth: buildSalesOnlyClients(unmatchedSales)
    };
  }

  function buildMonthList(quotes, salesRows) {
    return unique([...quotes, ...salesRows].map((item) => item.monthKey).filter(Boolean)).sort().map((key) => ({ key, label: monthLabel(key) }));
  }

  function dedupeQuotes(quotes) {
    const map = new Map();
    quotes.forEach((quote) => {
      if (!map.has(quote.reference)) {
        map.set(quote.reference, { ...quote, services: [...quote.services], modes: [...quote.modes], users: [...quote.users] });
        return;
      }
      const current = map.get(quote.reference);
      current.services = unique([...current.services, ...quote.services]);
      current.modes = unique([...current.modes, ...quote.modes]);
      current.users = unique([...current.users, ...quote.users]);
      current.status = prioritizeStatus(current.status, quote.status);
      if (current.network === "Sin red" && quote.network !== "Sin red") current.network = quote.network;
    });
    return Array.from(map.values());
  }

  function prioritizeStatus(a, b) {
    const order = ["Cerrado", "Cotizado Agentes", "Cotizado Pricing", "Cotizado Pricing Fuera de Tiempo", "No Cotizado por Pricing", "No Cotizada", "Cotizacion Cancelada"];
    const ia = order.indexOf(a);
    const ib = order.indexOf(b);
    if (ia === -1) return b;
    if (ib === -1) return a;
    return ia <= ib ? a : b;
  }

  function assignConvertedQuotes(quotes, salesRows) {
    const quotesByKey = new Map();
    const salesByKey = new Map();
    quotes.forEach((quote) => addToGroup(quotesByKey, `${quote.monthKey}|${quote.canonicalClient}`, quote));
    salesRows.forEach((sale) => addToGroup(salesByKey, `${sale.monthKey}|${sale.canonicalClient}`, sale));
    const assignments = [];
    salesByKey.forEach((salesGroup, key) => {
      const quoteGroup = quotesByKey.get(key);
      if (!quoteGroup || !quoteGroup.length) return;
      const assignable = Math.min(salesGroup.length, quoteGroup.length);
      const sortedQuotes = [...quoteGroup].sort(compareQuotesForConversion);
      for (let index = 0; index < assignable; index += 1) {
        assignments.push({ sale: salesGroup[index], quote: sortedQuotes[index] });
      }
    });
    return assignments;
  }

  function compareQuotesForConversion(a, b) {
    const score = (quote) => quote.status === "Cerrado" ? 0 : quote.status === "Cotizado Agentes" ? 1 : quote.status === "Cotizado Pricing" ? 2 : 3;
    return score(a) - score(b) || a.reference.localeCompare(b.reference, "es");
  }

  function buildServiceRows(quotes, convertedQuotes, mapper) {
    const quoteMap = new Map();
    const convertedMap = new Map();
    quotes.forEach((quote) => quote.services.forEach((service) => addSet(quoteMap, mapper(service), quote.reference)));
    convertedQuotes.forEach((quote) => quote.services.forEach((service) => addSet(convertedMap, mapper(service), quote.reference)));
    return unique([...quoteMap.keys(), ...convertedMap.keys()]).map((name) => {
      const quotesCount = quoteMap.get(name) ? quoteMap.get(name).size : 0;
      const closuresCount = convertedMap.get(name) ? convertedMap.get(name).size : 0;
      return { name, quotes: quotesCount, closures: closuresCount, hitRate: closuresCount / Math.max(quotesCount, 1) };
    }).filter((row) => row.quotes || row.closures).sort((a, b) => b.quotes - a.quotes || b.closures - a.closures);
  }

  function buildOwnershipRows(quotes) {
    const map = new Map();
    quotes.forEach((quote) => {
      if (!["Cotizado Agentes", "Cotizado Pricing"].includes(quote.status)) return;
      quote.services.forEach((service) => {
        if (!map.has(service)) map.set(service, { agents: new Set(), pricing: new Set() });
        map.get(service)[quote.status === "Cotizado Agentes" ? "agents" : "pricing"].add(quote.reference);
      });
    });
    return Array.from(map.entries()).map(([service, bucket]) => {
      const agents = bucket.agents.size;
      const pricing = bucket.pricing.size;
      const base = Math.max(agents + pricing, 1);
      return { service, agentsShare: agents / base, pricingShare: pricing / base };
    }).sort((a, b) => a.service.localeCompare(b.service, "es"));
  }

  function buildClientRows(quotes, salesRows, convertedQuotes) {
    const quoteMap = new Map();
    const convertedMap = new Map();
    const labels = new Map();
    quotes.forEach((quote) => {
      addSet(quoteMap, quote.canonicalClient, quote.reference);
      labels.set(quote.canonicalClient, shortestLabel(labels.get(quote.canonicalClient), quote.client));
    });
    convertedQuotes.forEach((quote) => addSet(convertedMap, quote.canonicalClient, quote.reference));
    return Array.from(quoteMap.keys()).map((key) => {
      const quotesCount = quoteMap.get(key).size;
      const closuresCount = convertedMap.get(key) ? convertedMap.get(key).size : 0;
      return { name: labels.get(key) || key, quotes: quotesCount, closures: closuresCount, hitRate: closuresCount / Math.max(quotesCount, 1), flag: clientFlag(quotesCount, closuresCount) };
    }).sort((a, b) => b.quotes - a.quotes || b.closures - a.closures);
  }

  function buildTeamRows(quotes, convertedQuotes, monthAKey, monthBKey) {
    return unique([...quotes.flatMap((quote) => quote.users), ...convertedQuotes.flatMap((quote) => quote.users)]).map((user) => {
      const userQuotes = quotes.filter((quote) => quote.users.includes(user));
      const userConverted = convertedQuotes.filter((quote) => quote.users.includes(user));
      return {
        name: user,
        quotes: userQuotes.length,
        closures: userConverted.length,
        hitRate: userConverted.length / Math.max(userQuotes.length, 1),
        noQuoted: userQuotes.filter((quote) => quote.status === "No Cotizada").length,
        noPricing: userQuotes.filter((quote) => quote.status === "No Cotizado por Pricing").length,
        monthAQuotes: userQuotes.filter((quote) => quote.monthKey === monthAKey).length,
        monthBQuotes: userQuotes.filter((quote) => quote.monthKey === monthBKey).length
      };
    }).filter((row) => row.quotes || row.closures).sort((a, b) => b.quotes - a.quotes || b.closures - a.closures);
  }

  function buildLossRows(quotes, months) {
    const labels = [["No Cotizada", "Equipo"], ["No Cotizado por Pricing", "Pricing"], ["Cotizacion Cancelada", "Cliente o info"], ["Cotizado Pricing Fuera de Tiempo", "Pricing"]];
    return labels.map(([cause, owner]) => {
      const monthValues = months.map((month) => quotes.filter((quote) => quote.monthKey === month.key && quote.status === cause).length);
      return { cause, owner, monthValues, total: monthValues.reduce((sum, value) => sum + value, 0) };
    }).filter((row) => row.total > 0);
  }

  function buildTrendServiceRows(quotes, convertedQuotes, monthAKey, monthBKey) {
    return unique([...quotes.flatMap((quote) => quote.services), ...convertedQuotes.flatMap((quote) => quote.services)]).map((service) => {
      const monthAQuotes = quotes.filter((quote) => quote.monthKey === monthAKey && quote.services.includes(service)).length;
      const monthBQuotes = quotes.filter((quote) => quote.monthKey === monthBKey && quote.services.includes(service)).length;
      const monthBClosures = convertedQuotes.filter((quote) => quote.monthKey === monthBKey && quote.services.includes(service)).length;
      return { service, monthAQuotes, monthBQuotes, monthBClosures, delta: monthBQuotes - monthAQuotes };
    }).filter((row) => row.monthAQuotes || row.monthBQuotes || row.monthBClosures).sort((a, b) => (b.monthAQuotes + b.monthBQuotes) - (a.monthAQuotes + a.monthBQuotes)).slice(0, 12);
  }

  function buildRepeatedClients(salesRows, monthAKey, monthBKey) {
    return unique(salesRows.map((sale) => sale.canonicalClient)).map((client) => ({
      name: client,
      monthA: salesRows.filter((sale) => sale.canonicalClient === client && sale.monthKey === monthAKey).length,
      monthB: salesRows.filter((sale) => sale.canonicalClient === client && sale.monthKey === monthBKey).length
    })).filter((row) => row.monthA > 0 && row.monthB > 0).sort((a, b) => (b.monthA + b.monthB) - (a.monthA + a.monthB)).slice(0, 10);
  }

  function buildClientDeltas(salesRows, monthAKey, monthBKey) {
    return unique(salesRows.map((sale) => sale.canonicalClient)).map((client) => ({
      name: client,
      monthA: salesRows.filter((sale) => sale.canonicalClient === client && sale.monthKey === monthAKey).length,
      monthB: salesRows.filter((sale) => sale.canonicalClient === client && sale.monthKey === monthBKey).length
    })).filter((row) => row.monthA || row.monthB).sort((a, b) => (b.monthB - b.monthA) - (a.monthB - a.monthA)).slice(0, 8);
  }

  function buildSalesOnlyClients(unmatchedSales) {
    const grouped = new Map();
    unmatchedSales.forEach((sale) => addToGroup(grouped, sale.monthKey, sale.canonicalClient));
    return Array.from(grouped.entries()).map(([monthKey, clients]) => ({
      monthKey,
      monthLabel: monthLabel(monthKey),
      clients: unique(clients).sort((a, b) => a.localeCompare(b, "es"))
    })).sort((a, b) => a.monthKey.localeCompare(b.monthKey));
  }

  function renderSummaryV2(data) {
    window.__mltiLatestSummary = data;
  }

  function renderOverview(data) {
    const t = data.totals;
    document.getElementById("resumen").innerHTML = `
      <div class="cards">
        ${card("Referencias válidas", int(t.totalQuotes), data.meta.quoteFileName)}
        ${card("Cierres reales", int(t.totalRealClosures), data.meta.salesFileName || "Sin reporte de ventas", "green")}
        ${card("Cierres que sí cruzan", int(t.totalMatchedClosures), "Los que sí entran al hit rate", "blue")}
        ${card("Hit Rate global", pct(t.globalHitRate), int(t.totalMatchedClosures) + " cierres sobre " + int(t.totalQuotes), "blue")}
        ${card("Pipeline activo", int(t.pipelineActive), "Cotizado Agentes + Cotizado Pricing", "yellow")}
        ${card("Pérdidas operativas", int(t.operationalLosses), int(t.internalLosses) + " por falla interna", "red")}
        ${card("Cierres sin cotización", int(t.totalSalesOnlyClosures), "Ventas que no pasaron por el cotizador", "orange")}
      </div>
      <div class="insight"><strong>Nueva lógica:</strong> el hit rate solo usa cierres del reporte de ventas que encuentran cotizaciones del mismo cliente en el mismo mes.</div>
      <div class="panel panel-full">
        <div class="panel-title"><div class="dot dot-blue"></div>Resumen por Mes</div>
        ${table(["Mes", "Cotizaciones", "Cierres reales", "Cierres cruzados", "Hit Rate", "Sin cotización"], data.monthRows.map((row) => [row.label, int(row.quotes), int(row.realClosures), int(row.matchedClosures), pct(row.hitRate), int(row.salesOnlyClosures)]), 1)}
      </div>
    `;
  }

  function renderServices(data) {
    document.getElementById("servicios").innerHTML = `
      <div class="grid2">
        <div class="panel">
          <div class="panel-title"><div class="dot dot-blue"></div>Hit Rate por Modo</div>
          ${table(["Modo", "Cotizaciones", "Cierres", "Hit Rate"], data.modeRows.map((row) => [row.name, int(row.quotes), int(row.closures), pct(row.hitRate)]), 1)}
        </div>
        <div class="panel">
          <div class="panel-title"><div class="dot dot-green"></div>Desglose por Subtipo</div>
          ${table(["Subtipo", "Cotizaciones", "Cierres", "Hit Rate"], data.subtypeRows.map((row) => [row.name, int(row.quotes), int(row.closures), pct(row.hitRate)]), 1)}
        </div>
      </div>
      <div class="panel panel-full">
        <div class="panel-title"><div class="dot dot-orange"></div>Quién cotiza los servicios activos</div>
        ${table(["Servicio", "Agentes", "Pricing"], data.ownershipRows.map((row) => [row.service, pct(row.agentsShare), pct(row.pricingShare)]), 1)}
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
          ${table(["Cliente", "Cot.", "Cierres", "Hit Rate", "Semáforo"], top.map((row) => [row.name, int(row.quotes), int(row.closures), pct(row.hitRate), row.flag]), 1)}
        </div>
        <div class="panel">
          <div class="panel-title"><div class="dot dot-red"></div>Clientes en Observación</div>
          ${table(["Cliente", "Cot.", "Cierres", "Hit Rate", "Semáforo"], risk.map((row) => [row.name, int(row.quotes), int(row.closures), pct(row.hitRate), row.flag]), 1)}
        </div>
      </div>
    `;
  }

  function renderTeam(data) {
    const monthALabel = data.meta.monthA ? data.meta.monthA.label : "Mes A";
    const monthBLabel = data.meta.monthB ? data.meta.monthB.label : "Mes B";
    document.getElementById("equipo").innerHTML = `
      <div class="grid2">
        <div class="panel">
          <div class="panel-title"><div class="dot dot-blue"></div>Eficiencia por Ejecutivo</div>
          ${table(["Ejecutivo", "Cot.", "Cierres", "Hit Rate", "No Cot.", "No Pricing"], data.teamRows.map((row) => [row.name, int(row.quotes), int(row.closures), pct(row.hitRate), int(row.noQuoted), int(row.noPricing)]), 1)}
        </div>
        <div class="panel">
          <div class="panel-title"><div class="dot dot-orange"></div>Volumen por Mes</div>
          ${table(["Ejecutivo", monthALabel, monthBLabel, "Cambio"], data.teamRows.map((row) => [row.name, int(row.monthAQuotes), int(row.monthBQuotes), signed(row.monthBQuotes - row.monthAQuotes)]), 1)}
        </div>
      </div>
    `;
  }

  function renderLosses(data) {
    const headers = ["Causa", ...data.meta.months.map((month) => month.label), "Total", "Quién falla"];
    document.getElementById("perdidas").innerHTML = `
      <div class="grid2">
        <div class="panel">
          <div class="panel-title"><div class="dot dot-red"></div>Pérdidas Operativas</div>
          ${table(headers, data.losses.map((row) => [row.cause, ...row.monthValues.map((value) => int(value)), int(row.total), row.owner]), 1)}
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
    const monthALabel = data.meta.monthA ? data.meta.monthA.label : "Mes A";
    const monthBLabel = data.meta.monthB ? data.meta.monthB.label : "Mes B";
    const monthAStats = data.monthRows.find((row) => row.key === (data.meta.monthA && data.meta.monthA.key));
    const monthBStats = data.monthRows.find((row) => row.key === (data.meta.monthB && data.meta.monthB.key));
    document.getElementById("tendencia").innerHTML = `
      <div class="cards">
        ${card("Cot. " + monthALabel, int(monthAStats ? monthAStats.quotes : 0), "Base", "orange")}
        ${card("Cot. " + monthBLabel, int(monthBStats ? monthBStats.quotes : 0), signed((monthBStats ? monthBStats.quotes : 0) - (monthAStats ? monthAStats.quotes : 0)), "orange")}
        ${card("HR " + monthALabel, pct(monthAStats ? monthAStats.hitRate : 0), "Conversión", "blue")}
        ${card("HR " + monthBLabel, pct(monthBStats ? monthBStats.hitRate : 0), signedPct((monthBStats ? monthBStats.hitRate : 0) - (monthAStats ? monthAStats.hitRate : 0)), "green")}
      </div>
      <div class="grid2">
        <div class="panel">
          <div class="panel-title"><div class="dot dot-orange"></div>Mix de Servicios</div>
          ${table(["Servicio", monthALabel, monthBLabel, "Δ Cot.", "Cierres " + monthBLabel], data.trendServices.map((row) => [row.service, int(row.monthAQuotes), int(row.monthBQuotes), signed(row.delta), int(row.monthBClosures)]), 1)}
        </div>
        <div class="panel">
          <div class="panel-title"><div class="dot dot-green"></div>Clientes con Cierres en Ambos Meses</div>
          ${table(["Cliente", monthALabel, monthBLabel, "Tendencia"], data.repeatedClients.map((row) => [row.name, int(row.monthA), int(row.monthB), row.monthB > row.monthA ? "Subió" : row.monthB < row.monthA ? "Bajó" : "Estable"]), 1)}
        </div>
      </div>
      <div class="panel panel-full">
        <div class="panel-title"><div class="dot dot-purple"></div>Clientes que Más Cambiaron</div>
        ${table(["Cliente", monthALabel, monthBLabel, "Δ Cierres"], data.clientDeltas.map((row) => [row.name, int(row.monthA), int(row.monthB), signed(row.monthB - row.monthA)]), 1)}
        ${renderSalesOnlyNotes(data.salesOnlyByMonth)}
      </div>
    `;
  }

  function renderSalesOnlyNotes(items) {
    if (!items.length) return "";
    return `<div class="insight" style="margin-top:16px;"><strong>Cierres sin cotización registrada:</strong> ${items.map((item) => `${item.monthLabel}: ${item.clients.join(", ")}`).join(" · ")}</div>`;
  }

  function card(label, value, sub, color = "") {
    const tone = color ? " " + color : "";
    return `<div class="card${tone}"><div class="card-label">${escapeHtml(label)}</div><div class="card-value${tone}">${escapeHtml(value)}</div><div class="card-sub">${escapeHtml(sub)}</div></div>`;
  }

  function table(headers, rows, numericFrom = 1) {
    if (!rows.length) return `<div class="tbl-scroll"><table class="tbl"><tr><td>Sin datos suficientes para esta vista.</td></tr></table></div>`;
    return `<div class="tbl-scroll"><table class="tbl"><tr>${headers.map((header, index) => `<th${index >= numericFrom ? ` class="num"` : ""}>${escapeHtml(header)}</th>`).join("")}</tr>${rows.map((row) => `<tr>${row.map((value, index) => `<td${index >= numericFrom ? ` class="num"` : ""}>${badgeIfNeeded(value)}</td>`).join("")}</tr>`).join("")}</table></div>`;
  }

  function badgeIfNeeded(value) {
    if (["Alto", "Medio", "Bajo", "Crítico", "Critico", "Subió", "Bajó", "Estable"].includes(value)) {
      const cls = value === "Alto" || value === "Subió" ? "badge-green" : value === "Medio" || value === "Estable" ? "badge-yellow" : "badge-red";
      return `<span class="badge ${cls}">${escapeHtml(value)}</span>`;
    }
    return escapeHtml(value);
  }

  function clientFlag(quotesCount, closuresCount) {
    const rate = closuresCount / Math.max(quotesCount, 1);
    if (quotesCount >= 15 && rate <= 0.05) return "Crítico";
    if (rate >= 0.5) return "Alto";
    if (rate >= 0.2) return "Medio";
    return "Bajo";
  }

  function normalize(value) {
    return String(value || "").normalize("NFD").replace(/[\u0300-\u036f]/g, "").toLowerCase().replace(/[^a-z0-9&]+/g, " ").replace(/\s+/g, " ").trim();
  }

  function parseDate(value) {
    if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
    if (typeof value === "number") {
      const parsed = XLSX.SSF.parse_date_code(value);
      if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d);
    }
    const text = String(value || "").trim();
    if (!text) return null;
    const match = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{2,4})$/);
    if (match) {
      const year = Number(match[3]) < 100 ? 2000 + Number(match[3]) : Number(match[3]);
      return new Date(year, Number(match[2]) - 1, Number(match[1]));
    }
    const parsed = new Date(text);
    return Number.isNaN(parsed.getTime()) ? null : parsed;
  }

  function monthKey(date) {
    if (!(date instanceof Date) || Number.isNaN(date.getTime())) return "";
    return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, "0")}`;
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
    return unique(text.replace(/\s+y\s+/gi, "|").replace(/[;,/\\]+/g, "|").split("|").map((item) => item.trim()).filter(Boolean));
  }

  function splitServices(value) {
    const text = String(value || "").trim();
    if (!text) return ["Sin clasificar"];
    const pieces = text.replace(/[;]+/g, ",").split(",").map((item) => item.trim()).filter(Boolean);
    const services = [];
    pieces.forEach((piece) => normalizeService(piece).forEach((service) => services.push(service)));
    return unique(services.length ? services : ["Sin clasificar"]);
  }

  function normalizeService(value) {
    const key = normalize(value);
    const services = [];
    if (key.includes("free hand")) services.push("Free Hand");
    if (/d2d|otr/.test(key)) services.push("D2D OTR");
    if (/dap/.test(key) && /marit|ocean|sea/.test(key)) services.push("DAP Maritimo");
    if (/exw/.test(key) && /marit|ocean|sea/.test(key)) services.push("EXW Maritimo");
    if (/dap/.test(key) && /aer|air/.test(key)) services.push("DAP Aereo");
    if (/exw/.test(key) && /aer|air/.test(key)) services.push("EXW Aereo");
    return services.length ? services : [String(value || "").trim() || "Sin clasificar"];
  }

  function normalizeSalesMode(modalityValue, serviceValue) {
    const modality = normalize(modalityValue);
    const service = normalize(serviceValue);
    if (modality.includes("hand carrier")) return "Aereo";
    if (service.includes("free hand")) return "Free Hand";
    if (modality.includes("marit")) return "Maritimo";
    if (modality.includes("aer")) return "Aereo";
    if (modality.includes("terrestre") || service.includes("otr") || service.includes("ltl") || service.includes("ftl")) return "OTR";
    return String(modalityValue || serviceValue || "Sin clasificar").trim();
  }

  function shipmentPrefix(shipment) {
    const text = String(shipment || "").trim().toUpperCase();
    const match = text.match(/^([A-Z]{2,3})-/);
    return match ? match[1] : "";
  }

  function shipmentModeFromCode(shipment) {
    const prefix = shipmentPrefix(shipment);
    return ACTIVE_SHIPMENT_MODES[prefix] || "";
  }

  function normalizeStatus(value) {
    const key = normalize(value);
    if (!key) return "Otro";
    if (key.includes("fuera de tiempo") && key.includes("pricing")) return "Cotizado Pricing Fuera de Tiempo";
    if ((key.includes("no cotizado") || key.includes("no cotizada")) && key.includes("pricing")) return "No Cotizado por Pricing";
    if (key.includes("no cotizado") || key.includes("no cotizada")) return "No Cotizada";
    if (key.includes("cancel")) return "Cotizacion Cancelada";
    if (key.includes("cerrado")) return "Cerrado";
    if (key.includes("cotizado") && key.includes("agente")) return "Cotizado Agentes";
    if (key.includes("pricing")) return "Cotizado Pricing";
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

  function canonicalClient(value) {
    const raw = String(value || "").trim();
    const normalized = normalize(raw);
    if (!normalized) return "SIN CLIENTE";
    for (const alias of CLIENT_ALIASES) {
      if (alias.patterns.some((pattern) => pattern.test(normalized))) return alias.canonical;
    }
    const tokens = normalized.split(" ").filter((token) => token && !LEGAL_TOKENS.has(token));
    return (tokens.length ? tokens.slice(0, 3).join(" ") : raw).toUpperCase();
  }

  function shortestLabel(current, candidate) {
    if (!candidate) return current || "";
    if (!current) return candidate;
    return candidate.length < current.length ? candidate : current;
  }

  function addSet(map, key, value) {
    if (!map.has(key)) map.set(key, new Set());
    map.get(key).add(value);
  }

  function addToGroup(map, key, value) {
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(value);
  }

  function unique(items) {
    return Array.from(new Set(items));
  }

  function int(value) { return Number(value || 0).toLocaleString("es-MX"); }
  function pct(value) { return `${(Number(value || 0) * 100).toFixed(1)}%`; }
  function signed(value) { const number = Number(value || 0); return `${number > 0 ? "+" : ""}${number.toLocaleString("es-MX")}`; }
  function signedPct(value) { const number = (Number(value || 0) * 100).toFixed(1); return `${value >= 0 ? "+" : ""}${number} pp`; }
  function escapeHtml(value) { return String(value == null ? "" : value).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#39;"); }

  function clearSavedUpload() {
    localStorage.removeItem(STORAGE_KEY_V2);
    setFeedback("Se borró la carga guardada de este navegador.", "success");
  }

  function tryLoadSaved() {
    try {
      const saved = localStorage.getItem(STORAGE_KEY_V2);
      if (saved) {
        renderSummaryV2(JSON.parse(saved));
        setFeedback("Se cargó la última versión guardada con la lógica nueva.", "success");
      }
    } catch (error) {
      console.error(error);
    }
  }
})();
