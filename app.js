import {
  PublicClientApplication,
  InteractionRequiredAuthError
} from "https://cdn.jsdelivr.net/npm/@azure/msal-browser@4.29.0/+esm";

import { CONFIG } from "./config.js?v=8";

/* ================= FORMATOS ================= */

const money = new Intl.NumberFormat("es-PA", {
  style: "currency",
  currency: "USD"
});

const num = new Intl.NumberFormat("es-PA", {
  maximumFractionDigits: 2
});

const pct = new Intl.NumberFormat("es-PA", {
  style: "percent",
  maximumFractionDigits: 1
});

const dateFmt = new Intl.DateTimeFormat("es-PA", {
  day: "2-digit",
  month: "short",
  year: "numeric"
});

const colors = [
  "#2563eb", "#0891b2", "#059669", "#d97706",
  "#7c3aed", "#dc2626", "#64748b", "#db2777"
];

const charts = {};
let account = null;

function el(id) {
  const node = document.getElementById(id);
  if (!node) {
    throw new Error(
      `Falta "${id}". Reemplaza también index.html por la versión 8.`
    );
  }
  return node;
}

function status(message, type = "") {
  el("status").textContent = message;
  el("status").className = `status ${type}`.trim();
}

function busy(value) {
  ["loginButton", "refreshButton", "logoutButton"].forEach(id => {
    el(id).disabled = value;
  });
  el("refreshButton").textContent =
    value ? "Actualizando..." : "Actualizar";
}

function session(value) {
  el("loginButton").hidden = value;
  el("refreshButton").hidden = !value;
  el("logoutButton").hidden = !value;
}

function numeric(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }
  const result = Number(
    String(value ?? "").replace(/[$,]/g, "").trim()
  );
  return Number.isFinite(result) ? result : 0;
}

function median(values) {
  const a = values.filter(Number.isFinite).sort((x, y) => x - y);
  if (!a.length) return 0;
  const i = Math.floor(a.length / 2);
  return a.length % 2 ? a[i] : (a[i - 1] + a[i]) / 2;
}

function parseDate(value) {
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : value;
  }

  if (typeof value === "number") {
    const d = XLSX.SSF.parse_date_code(value);
    return d ? new Date(d.y, d.m - 1, d.d, 12) : null;
  }

  const text = String(value ?? "").trim();
  if (!text) return null;

  const match = text.match(
    /^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/
  );

  if (match) {
    const day = Number(match[1]);
    const month = Number(match[2]);
    const year = Number(match[3]);
    const d = new Date(year, month - 1, day, 12);

    return d.getFullYear() === year &&
      d.getMonth() === month - 1 &&
      d.getDate() === day ? d : null;
  }

  const d = new Date(text);
  return Number.isNaN(d.getTime()) ? null : d;
}

function monthKey(d) {
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, "0")}`;
}

function dayKey(d) {
  return `${monthKey(d)}-${String(d.getDate()).padStart(2, "0")}`;
}

function html(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function groupKey(category, subcategory) {
  return JSON.stringify([category, subcategory]);
}

function createChart(name, id, config) {
  if (charts[name]) charts[name].destroy();
  charts[name] = new Chart(el(id), config);
}

function resizeCharts() {
  Object.values(charts).forEach(chart => chart.resize());
}

/* ================= ETIQUETAS DE AÑO ================= */

// Dibuja cada año centrado bajo sus meses.
const yearBand = {
  id: "yearBand",

  afterDraw(chart) {
    const months = chart.options.plugins.yearBand?.months;
    if (!months?.length || !chart.scales.x) return;

    const groups = new Map();

    months.forEach((month, i) => {
      const year = month.slice(0, 4);
      if (!groups.has(year)) groups.set(year, []);
      groups.get(year).push(i);
    });

    const { ctx, scales } = chart;

    ctx.save();
    ctx.textAlign = "center";
    ctx.fillStyle = "#667085";
    ctx.font = "12px system-ui";

    for (const [year, indexes] of groups) {
      const left = scales.x.getPixelForTick(indexes[0]);
      const right = scales.x.getPixelForTick(indexes.at(-1));
      ctx.fillText(year, (left + right) / 2, scales.x.bottom + 18);
    }

    ctx.restore();
  }
};

/* ================= AUTENTICACIÓN ================= */

const msal = new PublicClientApplication({
  auth: {
    clientId: CONFIG.clientId,
    authority: CONFIG.authority,
    redirectUri: CONFIG.redirectUri,
    postLogoutRedirectUri: CONFIG.redirectUri
  },
  cache: {
    cacheLocation: "localStorage"
  }
});

async function token() {
  const request = { account, scopes: CONFIG.scopes };

  try {
    return (await msal.acquireTokenSilent(request)).accessToken;
  } catch (error) {
    if (!(error instanceof InteractionRequiredAuthError)) throw error;
    return (await msal.acquireTokenPopup(request)).accessToken;
  }
}

async function graph(url, accessToken) {
  const response = await fetch(url, {
    headers: { Authorization: `Bearer ${accessToken}` },
    cache: "no-store"
  });

  if (!response.ok) {
    throw new Error(
      `Microsoft Graph ${response.status}: ${await response.text()}`
    );
  }

  return response;
}

async function getWorkbook(accessToken) {
  const path = CONFIG.workbookPath
    .split("/")
    .map(encodeURIComponent)
    .join("/");

  const response = await graph(
    "https://graph.microsoft.com/v1.0/me/drive/" +
    `special/documents:/${path}:` +
    "?$select=id,name,lastModifiedDateTime",
    accessToken
  );

  const item = await response.json();

  const content = await graph(
    "https://graph.microsoft.com/v1.0/me/drive/items/" +
    `${encodeURIComponent(item.id)}/content`,
    accessToken
  );

  return { item, buffer: await content.arrayBuffer() };
}

/* ================= LECTURA DEL EXCEL ================= */

function readRows(buffer) {
  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: true
  });

  const sheet = workbook.Sheets[CONFIG.worksheetName];

  if (!sheet) {
    throw new Error(`No existe la hoja "${CONFIG.worksheetName}".`);
  }

  const matrix = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    raw: true,
    defval: null
  });

  const headers = (matrix[0] || []).map(x => String(x ?? "").trim());

  const required = [
    "ID factura", "Fecha Factura", "Descripción", "Cantidad",
    "Valor Unitario", "Valor Item", "Total Factura",
    "Producto normalizado", "Categoría", "Subcategoría"
  ];

  const missing = required.filter(x => !headers.includes(x));

  if (missing.length) {
    throw new Error(
      `Faltan columnas: ${missing.join(", ")}. ` +
      "Guarda y sincroniza el Excel actualizado."
    );
  }

  const index = Object.fromEntries(
    required.map(x => [x, headers.indexOf(x)])
  );

  const rows = [];
  let skipped = 0;

  for (const source of matrix.slice(1)) {
    if (!source.some(x => x !== null && x !== "")) continue;

    const get = name => source[index[name]];
    const id = String(get("ID factura") ?? "").trim();
    const date = parseDate(get("Fecha Factura"));
    const description = String(get("Descripción") ?? "").trim();
    const product = String(
      get("Producto normalizado") || description
    ).trim();

    if (!id || !date || !product) {
      skipped++;
      continue;
    }

    rows.push({
      id,
      date,
      description,
      product,
      category: String(get("Categoría") || "Sin categoría").trim(),
      subcategory: String(get("Subcategoría") || "Por clasificar").trim(),
      quantity: numeric(get("Cantidad")),
      price: numeric(get("Valor Unitario")),
      spend: numeric(get("Valor Item")),
      invoiceTotal: numeric(get("Total Factura"))
    });
  }

  return { rows, skipped };
}

/* ================= CÁLCULOS ================= */

function aggregate(rows) {
  const invoices = new Map();
  const groups = new Map();
  const products = new Map();
  const monthly = new Map();

  for (const row of rows) {
    if (!invoices.has(row.id)) {
      invoices.set(row.id, {
        date: row.date,
        total: row.invoiceTotal
      });
    }

    const key = groupKey(row.category, row.subcategory);

    if (!groups.has(key)) {
      groups.set(key, {
        key,
        category: row.category,
        subcategory: row.subcategory,
        spend: 0,
        invoices: new Set()
      });
    }

    const group = groups.get(key);
    group.spend += row.spend;
    group.invoices.add(row.id);

    // Una sola entrada por producto normalizado.
    if (!products.has(row.product)) {
      products.set(row.product, {
        product: row.product,
        groups: new Set(),
        invoices: new Set(),
        quantities: new Map(),
        prices: [],
        pricesByMonth: new Map(),
        pricesByDay: new Map()
      });
    }

    const p = products.get(row.product);
    p.groups.add(key);
    p.invoices.add(row.id);
    p.quantities.set(
      row.id,
      (p.quantities.get(row.id) || 0) + row.quantity
    );

    if (row.price > 0 && row.quantity > 0) {
      p.prices.push(row.price);

      const month = monthKey(row.date);
      if (!p.pricesByMonth.has(month)) p.pricesByMonth.set(month, []);
      p.pricesByMonth.get(month).push(row.price);

      const day = dayKey(row.date);
      if (!p.pricesByDay.has(day)) p.pricesByDay.set(day, []);
      p.pricesByDay.get(day).push(row.price);
    }
  }

  for (const invoice of invoices.values()) {
    const month = monthKey(invoice.date);
    monthly.set(month, (monthly.get(month) || 0) + invoice.total);
  }

  const count = invoices.size;

  for (const p of products.values()) {
    p.recurrence = p.invoices.size / count;
    p.medianPrice = median(p.prices);
    p.medianQuantity = median([...p.quantities.values()]);
    p.basketCost = p.medianPrice * p.medianQuantity;

    const days = [...p.pricesByDay.keys()].sort();
    p.comparable = days.length >= 2;

    p.change = p.comparable
      ? median(p.pricesByDay.get(days.at(-1))) /
        median(p.pricesByDay.get(days[0])) - 1
      : null;
  }

  const dates = [...invoices.values()]
    .map(x => x.date)
    .sort((a, b) => a - b);

  return {
    rows,
    invoices,
    count,
    groups: [...groups.values()].sort((a, b) => b.spend - a.spend),
    products: [...products.values()],
    monthly: [...monthly.entries()].sort(([a], [b]) => a.localeCompare(b)),
    total: [...invoices.values()].reduce((sum, x) => sum + x.total, 0),
    dates
  };
}

/* ================= GRÁFICOS GENERALES ================= */

function monthlyConfig(monthly) {
  const months = monthly.map(([month]) => month);

  return {
    type: "bar",
    plugins: [yearBand],
    data: {
      labels: months.map(x => String(Number(x.slice(5)))),
      datasets: [{
        label: "Gasto",
        data: monthly.map(([, value]) => value),
        backgroundColor: "#2563eb",
        borderRadius: 5,
        maxBarThickness: 48
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { bottom: 24 } },
      plugins: {
        yearBand: { months },
        legend: { display: false },
        tooltip: {
          callbacks: {
            title: items => months[items[0].dataIndex],
            label: context => money.format(context.raw)
          }
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: { autoSkip: false, maxRotation: 0 }
        },
        y: {
          beginAtZero: true,
          title: { display: true, text: "Miles" },
          ticks: { callback: value => num.format(value / 1000) }
        }
      }
    }
  };
}

function barConfig(labels, values, percentages = false) {
  return {
    type: "bar",
    data: {
      labels,
      datasets: [{
        label: percentages ? "Cambio mediano" : "Gasto",
        data: values,
        backgroundColor: values.map((value, i) =>
          percentages
            ? value >= 0 ? "#dc2626" : "#059669"
            : colors[i % colors.length]
        ),
        borderRadius: 4
      }]
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            label: context => percentages
              ? `${num.format(context.raw)}%`
              : money.format(context.raw)
          }
        }
      },
      scales: {
        x: {
          beginAtZero: true,
          ticks: {
            callback: value => percentages
              ? `${num.format(value)}%`
              : `${num.format(value / 1000)}k`
          }
        },
        y: {
          ticks: { autoSkip: false, font: { size: 11 } }
        }
      }
    }
  };
}

/* ================= RESUMEN ================= */

function renderSummary(data) {
  el("totalGasto").textContent = money.format(data.total);
  el("totalFacturas").textContent = num.format(data.count);
  el("ticketPromedio").textContent = money.format(data.total / data.count);
  el("totalSubcategorias").textContent = num.format(data.groups.length);

  createChart(
    "summaryMonthly",
    "summaryMonthlyChart",
    monthlyConfig(data.monthly.slice(-8))
  );

  const top = data.groups.slice(0, 10).map(x => ({
    label: x.subcategory,
    value: x.spend
  }));

  if (data.groups.length > 10) {
    top.push({
      label: "Resto",
      value: data.groups.slice(10).reduce((sum, x) => sum + x.spend, 0)
    });
  }

  createChart(
    "summarySubcategory",
    "summarySubcategoryChart",
    barConfig(top.map(x => x.label), top.map(x => x.value))
  );

  createChart("monthly", "monthlyChart", monthlyConfig(data.monthly));
}

/* ================= SUBCATEGORÍAS ================= */

function renderGroups(data) {
  el("subcategoryChartBox").style.height =
    `${Math.max(440, data.groups.length * 29 + 70)}px`;

  createChart(
    "subcategories",
    "subcategoryChart",
    barConfig(
      data.groups.map(x => `${x.subcategory} · ${x.category}`),
      data.groups.map(x => x.spend)
    )
  );

  const categories = new Map();

  for (const group of data.groups) {
    if (!categories.has(group.category)) {
      categories.set(group.category, {
        name: group.category,
        total: 0,
        groups: []
      });
    }
    const c = categories.get(group.category);
    c.total += group.spend;
    c.groups.push(group);
  }

  el("categoryDetails").innerHTML = [...categories.values()]
    .sort((a, b) => b.total - a.total)
    .map(category => `
      <details open>
        <summary>
          ${html(category.name)}
          <span>${money.format(category.total)}</span>
        </summary>
        <div class="table-wrap">
          <table>
            <thead>
              <tr>
                <th>Subcategoría</th>
                <th class="num">Gasto</th>
                <th class="num">% categoría</th>
                <th class="num">Facturas</th>
                <th class="num">Recurrencia</th>
              </tr>
            </thead>
            <tbody>
              ${category.groups.map(group => `
                <tr>
                  <td>${html(group.subcategory)}</td>
                  <td class="num">${money.format(group.spend)}</td>
                  <td class="num">${category.total
                    ? pct.format(group.spend / category.total)
                    : "—"}</td>
                  <td class="num">${group.invoices.size}</td>
                  <td class="num">${pct.format(
                    group.invoices.size / data.count
                  )}</td>
                </tr>
              `).join("")}
            </tbody>
          </table>
        </div>
      </details>
    `).join("");
}

/* ================= PRECIOS RESUMIDOS ================= */

function renderPrices(data) {
  const map = new Map();

  for (const p of data.products) {
    // Excluye productos con clasificación contradictoria.
    if (!p.comparable || p.groups.size !== 1) continue;

    const key = [...p.groups][0];
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(p.change);
  }

  const summary = data.groups.filter(g => map.has(g.key))
    .map(g => ({
      ...g,
      comparable: map.get(g.key).length,
      change: median(map.get(g.key))
    }))
    .sort((a, b) => Math.abs(b.change) - Math.abs(a.change));

  const changed = summary.filter(x => Math.abs(x.change) >= 0.05);

  el("priceInfo").textContent =
    `${summary.length} subcategorías con productos comparables. ` +
    `${changed.length} con cambio mediano de ±5% o más. ` +
    "El gráfico muestra estas últimas; la tabla incluye todas.";

  el("priceChartBox").hidden = !changed.length;
  el("priceChartBox").style.height =
    `${Math.max(340, changed.length * 30 + 70)}px`;

  if (charts.prices) {
    charts.prices.destroy();
    delete charts.prices;
  }

  if (changed.length) {
    createChart(
      "prices",
      "priceChart",
      barConfig(
        changed.map(x => `${x.subcategory} · ${x.category}`),
        changed.map(x => x.change * 100),
        true
      )
    );
  }

  el("priceTable").innerHTML = summary.length
    ? summary.map(x => `
        <tr>
          <td>${html(x.subcategory)}</td>
          <td>${html(x.category)}</td>
          <td class="num">${x.comparable}</td>
          <td class="num ${x.change > 0
            ? "positive"
            : x.change < 0 ? "negative" : ""}">
            ${pct.format(x.change)}
          </td>
        </tr>
      `).join("")
    : `<tr><td colspan="4">
         No hay productos con precios positivos en dos fechas distintas.
       </td></tr>`;
}

/* ================= CANASTA RESUMIDA ================= */

function renderBasket(data) {
  const habitual = data.products.filter(p =>
    p.recurrence >= 0.4 &&
    p.prices.length > 0 &&
    p.medianQuantity > 0 &&
    p.groups.size === 1
  );

  const groups = new Map();

  for (const p of habitual) {
    const key = [...p.groups][0];
    const source = data.groups.find(g => g.key === key);

    if (!groups.has(key)) {
      groups.set(key, {
        ...source,
        value: 0,
        componentInvoices: new Set()
      });
    }

    const group = groups.get(key);
    group.value += p.basketCost;

    for (const id of p.invoices) group.componentInvoices.add(id);
  }

  const baseline = habitual.reduce((sum, p) => sum + p.basketCost, 0);
  const months = data.monthly.map(([month]) => month);

  const monthly = months.map(month => {
    let imputed = 0;

    const value = habitual.reduce((sum, p) => {
      const prices = p.pricesByMonth.get(month);
      if (!prices?.length) imputed++;

      const price = prices?.length ? median(prices) : p.medianPrice;
      return sum + price * p.medianQuantity;
    }, 0);

    return {
      value,
      imputed,
      deviation: baseline ? (value - baseline) / baseline : 0
    };
  });

  const last = monthly.at(-1) || {
    value: 0,
    imputed: 0,
    deviation: 0
  };

  el("basketBaseline").textContent = money.format(baseline);
  el("basketLatest").textContent = money.format(last.value);
  el("basketDeviation").textContent = pct.format(last.deviation);
  el("basketDeviation").className =
    last.deviation > 0 ? "positive" : last.deviation < 0 ? "negative" : "";
  el("basketGroups").textContent = num.format(groups.size);

  el("basketInfo").textContent =
    `${habitual.length} componentes internos agrupados en ` +
    `${groups.size} subcategorías. Último mes: ${last.imputed} ` +
    "componentes sin precio observado; se usa su mediana general. " +
    "La base usa cantidad mediana por factura donde se compró cada " +
    "componente × precio mediano de todo el periodo.";

  createChart("basket", "basketChart", {
    plugins: [yearBand],
    data: {
      labels: months.map(x => String(Number(x.slice(5)))),
      datasets: [
        {
          type: "line",
          label: "Valor estimado",
          data: monthly.map(x => x.value),
          borderColor: "#2563eb",
          backgroundColor: "rgba(37,99,235,.08)",
          fill: true,
          tension: 0.2,
          yAxisID: "y"
        },
        {
          type: "bar",
          label: "Desviación",
          data: monthly.map(x => x.deviation * 100),
          backgroundColor: monthly.map(x => x.deviation >= 0
            ? "rgba(220,38,38,.3)"
            : "rgba(5,150,105,.3)"),
          yAxisID: "y1"
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      layout: { padding: { bottom: 24 } },
      interaction: { mode: "index", intersect: false },
      plugins: {
        yearBand: { months },
        tooltip: {
          callbacks: {
            title: items => months[items[0].dataIndex],
            label: context =>
              `${context.dataset.label}: ` +
              (context.dataset.yAxisID === "y1"
                ? `${num.format(context.raw)}%`
                : money.format(context.raw))
          }
        }
      },
      scales: {
        x: {
          ticks: { autoSkip: false, maxRotation: 0 }
        },
        y: {
          position: "left",
          title: { display: true, text: "Valor" },
          ticks: { callback: value => num.format(value) }
        },
        y1: {
          position: "right",
          beginAtZero: true,
          grid: { drawOnChartArea: false },
          ticks: { callback: value => `${num.format(value)}%` }
        }
      }
    }
  });

  el("basketTable").innerHTML = groups.size
    ? [...groups.values()].sort((a, b) => b.value - a.value)
      .map(g => `
        <tr>
          <td>${html(g.subcategory)}</td>
          <td>${html(g.category)}</td>
          <td class="num">${pct.format(
            g.componentInvoices.size / data.count
          )}</td>
          <td class="num">${money.format(g.value)}</td>
        </tr>
      `).join("")
    : `<tr><td colspan="4">
         No hay componentes válidos con recurrencia ≥40%.
       </td></tr>`;
}

/* ================= CARGA Y EVENTOS ================= */

async function load() {
  busy(true);
  status("Descargando el Excel guardado en OneDrive…");

  try {
    if (!window.XLSX || !window.Chart) {
      throw new Error("No se cargaron las librerías. Recarga la página.");
    }

    const accessToken = await token();

    const profileResponse = await graph(
      "https://graph.microsoft.com/v1.0/me?$select=displayName",
      accessToken
    );

    const profile = await profileResponse.json();
    el("usuario").textContent = profile.displayName || "";

    const { item, buffer } = await getWorkbook(accessToken);
    const result = readRows(buffer);

    if (!result.rows.length) {
      throw new Error("No hay líneas de factura válidas.");
    }

    const data = aggregate(result.rows);

    el("dashboard").hidden = false;
    el("tabs").hidden = false;

    el("periodo").textContent =
      `${dateFmt.format(data.dates[0])} — ` +
      dateFmt.format(data.dates.at(-1));

    el("sourceInfo").textContent =
      `${item.name} · ${result.rows.length} líneas · ` +
      `archivo modificado ${dateFmt.format(
        new Date(item.lastModifiedDateTime)
      )}`;

    renderSummary(data);
    renderGroups(data);
    renderPrices(data);
    renderBasket(data);

    const conflicts = data.products.filter(p => p.groups.size > 1).length;
    const pendingLabels = [
      "por revisar", "por clasificar", "sin categoría", "revisar"
    ];

    const pending = new Set(data.rows.filter(row =>
      pendingLabels.includes(row.category.toLowerCase()) ||
      pendingLabels.includes(row.subcategory.toLowerCase())
    ).map(row => row.description)).size;

    const warnings = [];
    if (result.skipped) warnings.push(`${result.skipped} líneas omitidas`);
    if (conflicts) {
      warnings.push(
        `${conflicts} productos con clasificación contradictoria; ` +
        "excluidos de precios y canasta, incluidos en gasto"
      );
    }
    if (pending) warnings.push(`${pending} descripciones por clasificar`);

    status(
      `Versión 8: ${data.count} facturas y ` +
      `${data.groups.length} subcategorías. ` +
      (warnings.length ? warnings.join(". ") + "." : "Datos cargados."),
      warnings.length ? "warning" : "success"
    );

    requestAnimationFrame(resizeCharts);
  } catch (error) {
    console.error(error);
    el("dashboard").hidden = true;
    el("tabs").hidden = true;
    status(error.message || "Error al cargar los datos.", "error");
  } finally {
    busy(false);
  }
}

async function login() {
  busy(true);
  try {
    const result = await msal.loginPopup({
      scopes: CONFIG.scopes,
      prompt: "select_account"
    });

    account = result.account;
    msal.setActiveAccount(account);
    session(true);
    await load();
  } catch (error) {
    status(error.message || "Error al iniciar sesión.", "error");
  } finally {
    busy(false);
  }
}

async function logout() {
  busy(true);
  try {
    await msal.logoutPopup({
      account,
      postLogoutRedirectUri: CONFIG.redirectUri
    });

    account = null;
    msal.setActiveAccount(null);
    Object.values(charts).forEach(chart => chart.destroy());
    Object.keys(charts).forEach(key => delete charts[key]);

    el("dashboard").hidden = true;
    el("tabs").hidden = true;
    el("usuario").textContent = "";
    el("periodo").textContent = "Inicia sesión para cargar los datos.";
    session(false);
    status("Sesión cerrada.");
  } catch (error) {
    status(error.message || "Error al cerrar sesión.", "error");
  } finally {
    busy(false);
  }
}

document.querySelectorAll(".tab").forEach(button => {
  button.addEventListener("click", () => {
    document.querySelectorAll(".tab").forEach(x =>
      x.classList.remove("active")
    );
    document.querySelectorAll(".tab-panel").forEach(x =>
      x.classList.remove("active")
    );
    button.classList.add("active");
    el(button.dataset.tab).classList.add("active");
    requestAnimationFrame(resizeCharts);
  });
});

async function start() {
  busy(true);
  try {
    await msal.initialize();

    el("loginButton").addEventListener("click", login);
    el("refreshButton").addEventListener("click", load);
    el("logoutButton").addEventListener("click", logout);

    account = msal.getActiveAccount() ||
      msal.getAllAccounts()[0] || null;

    session(Boolean(account));

    if (account) {
      msal.setActiveAccount(account);
      await load();
    }
  } catch (error) {
    status(error.message || "No se pudo iniciar la aplicación.", "error");
  } finally {
    busy(false);
  }
}

start();
