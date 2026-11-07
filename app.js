import {
  PublicClientApplication,
  InteractionRequiredAuthError
} from "https://cdn.jsdelivr.net/npm/@azure/msal-browser@4.29.0/+esm";

import { CONFIG } from "./config.js";

/* =========================================================
   FORMATOS Y ESTADO
========================================================= */

const currency = new Intl.NumberFormat("es-PA", {
  style: "currency",
  currency: "USD"
});

const number = new Intl.NumberFormat("es-PA", {
  maximumFractionDigits: 2
});

const integer = new Intl.NumberFormat("es-PA", {
  maximumFractionDigits: 0
});

const percent = new Intl.NumberFormat("es-PA", {
  style: "percent",
  maximumFractionDigits: 1
});

const dateFormat = new Intl.DateTimeFormat("es-PA", {
  day: "2-digit",
  month: "short",
  year: "numeric"
});

const colors = [
  "#2563eb", "#0891b2", "#059669", "#d97706",
  "#7c3aed", "#dc2626", "#64748b", "#db2777",
  "#65a30d", "#4f46e5", "#ea580c"
];

let account = null;
let products = [];
const charts = {};

/* =========================================================
   ELEMENTOS Y UTILIDADES
========================================================= */

function element(id) {
  const node = document.getElementById(id);

  if (!node) {
    throw new Error(
      `Falta el elemento "${id}" en index.html. ` +
      "Utiliza el HTML compatible con las cinco pestañas."
    );
  }

  return node;
}

const loginButton = element("loginButton");
const logoutButton = element("logoutButton");
const refreshButton = element("refreshButton");
const statusBox = element("status");
const dashboard = element("dashboard");
const tabs = element("tabs");

function setStatus(message, type = "") {
  statusBox.textContent = message;
  statusBox.className = `status ${type}`.trim();
}

function setLoading(loading) {
  loginButton.disabled = loading;
  refreshButton.disabled = loading;
  logoutButton.disabled = loading;

  loginButton.textContent = loading
    ? "Conectando..."
    : "Iniciar sesión";

  refreshButton.textContent = loading
    ? "Actualizando..."
    : "Actualizar";
}

function showSession(signedIn) {
  loginButton.hidden = signedIn;
  logoutButton.hidden = !signedIn;
  refreshButton.hidden = !signedIn;
}

function asNumber(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  const parsed = Number(
    String(value ?? "")
      .replace(/\$/g, "")
      .replace(/,/g, "")
      .trim()
  );

  return Number.isFinite(parsed) ? parsed : 0;
}

function asDate(value) {
  if (value instanceof Date) {
    return Number.isNaN(value.getTime()) ? null : value;
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);

    return parsed
      ? new Date(parsed.y, parsed.m - 1, parsed.d, 12)
      : null;
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
    const date = new Date(year, month - 1, day, 12);

    return (
      date.getFullYear() === year &&
      date.getMonth() === month - 1 &&
      date.getDate() === day
    ) ? date : null;
  }

  const date = new Date(text);
  return Number.isNaN(date.getTime()) ? null : date;
}

function monthKey(date) {
  return (
    `${date.getFullYear()}-` +
    String(date.getMonth() + 1).padStart(2, "0")
  );
}

function dayKey(date) {
  return (
    `${monthKey(date)}-` +
    String(date.getDate()).padStart(2, "0")
  );
}

function median(values) {
  const sorted = values
    .filter(value =>
      typeof value === "number" &&
      Number.isFinite(value)
    )
    .sort((a, b) => a - b);

  if (!sorted.length) return 0;

  const middle = Math.floor(sorted.length / 2);

  return sorted.length % 2
    ? sorted[middle]
    : (sorted[middle - 1] + sorted[middle]) / 2;
}

function escapeHtml(value) {
  return String(value ?? "")
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function palette(index) {
  return colors[index % colors.length];
}

function destroyChart(name) {
  if (charts[name]) {
    charts[name].destroy();
    delete charts[name];
  }
}

function makeChart(name, canvasId, config) {
  destroyChart(name);
  charts[name] = new Chart(element(canvasId), config);
}

function resizeCharts() {
  Object.values(charts).forEach(chart => chart.resize());
}

function valueTooltip(context) {
  return currency.format(context.raw);
}

/* =========================================================
   MESES Y AÑOS AGRUPADOS
========================================================= */

// Meses numéricos. Cada año aparece una sola vez.
function groupedMonthLabels(months) {
  const groups = new Map();

  months.forEach((month, index) => {
    const year = month.slice(0, 4);
    if (!groups.has(year)) groups.set(year, []);
    groups.get(year).push(index);
  });

  const centers = new Map();

  for (const [year, indexes] of groups) {
    const center = Math.floor(
      (indexes[0] + indexes[indexes.length - 1]) / 2
    );
    centers.set(center, year);
  }

  return months.map((month, index) => [
    String(Number(month.slice(5, 7))),
    centers.get(index) || ""
  ]);
}

/* =========================================================
   AUTENTICACIÓN Y ONEDRIVE
========================================================= */

const msalInstance = new PublicClientApplication({
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

async function getAccessToken() {
  const request = {
    account,
    scopes: CONFIG.scopes
  };

  try {
    const response =
      await msalInstance.acquireTokenSilent(request);

    return response.accessToken;
  } catch (error) {
    if (!(error instanceof InteractionRequiredAuthError)) {
      throw error;
    }

    const response =
      await msalInstance.acquireTokenPopup(request);

    return response.accessToken;
  }
}

async function graphFetch(url, token) {
  const response = await fetch(url, {
    headers: {
      Authorization: `Bearer ${token}`
    },
    cache: "no-store"
  });

  if (!response.ok) {
    const detail = await response.text();

    throw new Error(
      `Microsoft Graph ${response.status}: ${detail}`
    );
  }

  return response;
}

async function findWorkbook(token) {
  if (!CONFIG.workbookPath) {
    throw new Error("Falta workbookPath en config.js.");
  }

  const path = CONFIG.workbookPath
    .split("/")
    .map(segment => encodeURIComponent(segment))
    .join("/");

  const url =
    "https://graph.microsoft.com/v1.0/me/drive/" +
    `special/documents:/${path}:` +
    "?$select=id,name,lastModifiedDateTime";

  const response = await graphFetch(url, token);
  return response.json();
}

async function downloadWorkbook(token, itemId) {
  const url =
    "https://graph.microsoft.com/v1.0/me/drive/items/" +
    `${encodeURIComponent(itemId)}/content`;

  const response = await graphFetch(url, token);
  return response.arrayBuffer();
}

/* =========================================================
   LEER LAS LÍNEAS DE FACTURAS
========================================================= */

function readInvoiceRows(buffer) {
  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: true
  });

  const sheet = workbook.Sheets[CONFIG.worksheetName];

  if (!sheet) {
    throw new Error(
      `No existe la hoja "${CONFIG.worksheetName}".`
    );
  }

  const matrix = XLSX.utils.sheet_to_json(sheet, {
    header: 1,
    defval: null,
    raw: true
  });

  const headers = (matrix[0] || [])
    .map(value => String(value ?? "").trim());

  const required = [
    "ID factura",
    "Fecha Factura",
    "Descripción",
    "Cantidad",
    "Valor Unitario",
    "Valor Item",
    "Total Factura",
    "Producto normalizado",
    "Categoría",
    "Subcategoría"
  ];

  const missing = required.filter(
    name => !headers.includes(name)
  );

  if (missing.length) {
    throw new Error(
      `Faltan columnas: ${missing.join(", ")}. ` +
      "Guarda y sincroniza el Excel con la nueva Subcategoría."
    );
  }

  const indexes = Object.fromEntries(
    required.map(name => [name, headers.indexOf(name)])
  );

  const sourceRows = matrix.slice(1).filter(row =>
    row.some(value => value !== null && value !== "")
  );

  const rows = [];
  let skipped = 0;

  for (const source of sourceRows) {
    const get = name => source[indexes[name]];

    const invoiceId = String(get("ID factura") ?? "").trim();
    const date = asDate(get("Fecha Factura"));
    const description = String(
      get("Descripción") ?? ""
    ).trim();

    const product = String(
      get("Producto normalizado") || description
    ).trim();

    if (!invoiceId || !date || !product) {
      skipped++;
      continue;
    }

    rows.push({
      invoiceId,
      date,
      description,
      product,
      category: String(
        get("Categoría") || "Sin categoría"
      ).trim(),
      subcategory: String(
        get("Subcategoría") || "Por clasificar"
      ).trim(),
      quantity: asNumber(get("Cantidad")),
      unitPrice: asNumber(get("Valor Unitario")),
      itemValue: asNumber(get("Valor Item")),
      invoiceTotal: asNumber(get("Total Factura"))
    });
  }

  return {
    rows,
    skipped,
    sourceCount: sourceRows.length
  };
}

/* =========================================================
   AGRUPACIÓN DEL 80%
========================================================= */

function groupToEightyPercent(entries, remainderLabel) {
  const sorted = [...entries]
    .sort((a, b) => b.value - a.value);

  if (!sorted.length) return [];

  const total = sorted.reduce(
    (sum, item) => sum + item.value,
    0
  );

  // Si hay importes negativos, no aplicamos un Pareto engañoso.
  if (total <= 0 || sorted.some(item => item.value < 0)) {
    return sorted;
  }

  const visible = [];
  let cumulative = 0;
  let position = 0;

  while (
    position < sorted.length &&
    (cumulative < total * 0.8 || visible.length === 0)
  ) {
    visible.push(sorted[position]);
    cumulative += sorted[position].value;
    position++;
  }

  const remainder = sorted.slice(position);
  const remainderValue = remainder.reduce(
    (sum, item) => sum + item.value,
    0
  );

  if (remainder.length && remainderValue > 0) {
    visible.push({
      label:
        `${remainderLabel} (` +
        `${percent.format(remainderValue / total)})`,
      value: remainderValue,
      isRemainder: true
    });
  }

  return visible;
}

/* =========================================================
   AGREGACIÓN: UNA ENTRADA POR PRODUCTO
========================================================= */

function aggregateData(rows) {
  const invoices = new Map();
  const productMap = new Map();
  const categoryMap = new Map();
  const monthlyMap = new Map();

  for (const row of rows) {
    if (!invoices.has(row.invoiceId)) {
      invoices.set(row.invoiceId, {
        date: row.date,
        total: row.invoiceTotal
      });
    }

    categoryMap.set(
      row.category,
      (categoryMap.get(row.category) || 0) + row.itemValue
    );

    if (!productMap.has(row.product)) {
      productMap.set(row.product, {
        product: row.product,
        categories: new Set(),
        subcategories: new Set(),
        invoices: new Set(),
        quantity: 0,
        spend: 0,
        prices: [],
        quantitiesByInvoice: new Map(),
        pricesByMonth: new Map(),
        pricesByDay: new Map()
      });
    }

    const product = productMap.get(row.product);

    product.categories.add(row.category);
    product.subcategories.add(row.subcategory);
    product.invoices.add(row.invoiceId);
    product.quantity += row.quantity;
    product.spend += row.itemValue;

    product.quantitiesByInvoice.set(
      row.invoiceId,
      (product.quantitiesByInvoice.get(row.invoiceId) || 0) +
      row.quantity
    );

    if (row.unitPrice > 0 && row.quantity > 0) {
      product.prices.push(row.unitPrice);

      const month = monthKey(row.date);
      if (!product.pricesByMonth.has(month)) {
        product.pricesByMonth.set(month, []);
      }
      product.pricesByMonth.get(month).push(row.unitPrice);

      const day = dayKey(row.date);
      if (!product.pricesByDay.has(day)) {
        product.pricesByDay.set(day, {
          date: row.date,
          prices: []
        });
      }
      product.pricesByDay.get(day).prices.push(row.unitPrice);
    }
  }

  for (const invoice of invoices.values()) {
    const month = monthKey(invoice.date);

    monthlyMap.set(
      month,
      (monthlyMap.get(month) || 0) + invoice.total
    );
  }

  const invoiceCount = invoices.size;

  products = [...productMap.values()].map(product => {
    const events = product.invoices.size;
    const recurrence = invoiceCount
      ? events / invoiceCount
      : 0;

    const medianPrice = median(product.prices);
    const medianQuantity = median(
      [...product.quantitiesByInvoice.values()]
    );

    // Si hay varios precios el mismo día, usamos su mediana.
    const observations = [...product.pricesByDay.entries()]
      .sort(([a], [b]) => a.localeCompare(b))
      .map(([, observation]) => ({
        date: observation.date,
        price: median(observation.prices)
      }));

    const firstObservation = observations[0] || null;
    const lastObservation = observations.at(-1) || null;

    const priceChange =
      firstObservation && lastObservation
        ? (
            lastObservation.price - firstObservation.price
          ) / firstObservation.price
        : 0;

    return {
      ...product,
      category: [...product.categories].join(" / "),
      subcategory: [...product.subcategories].join(" / "),
      events,
      recurrence,
      medianPrice,
      medianQuantity,
      basketCost: medianPrice * medianQuantity,
      habitual: recurrence >= 0.4,
      observationDays: observations.length,
      firstObservation,
      lastObservation,
      priceChange
    };
  }).sort((a, b) => b.spend - a.spend);

  const conflictingProducts = products.filter(product =>
    product.categories.size > 1 ||
    product.subcategories.size > 1
  );

  const dates = [...invoices.values()]
    .map(invoice => invoice.date)
    .sort((a, b) => a - b);

  return {
    rows,
    invoices,
    invoiceCount,
    conflictingProducts,
    totalSpend: [...invoices.values()].reduce(
      (sum, invoice) => sum + invoice.total,
      0
    ),
    dates,
    monthly: [...monthlyMap.entries()]
      .sort(([a], [b]) => a.localeCompare(b)),
    categories: [...categoryMap.entries()]
      .map(([label, value]) => ({ label, value }))
      .sort((a, b) => b.value - a.value)
  };
}

/* =========================================================
   RESUMEN Y GASTO MENSUAL
========================================================= */

function monthlyChartConfig(monthly, compact = false) {
  const months = monthly.map(([month]) => month);

  return {
    type: "bar",
    data: {
      labels: groupedMonthLabels(months),
      datasets: [{
        label: "Gasto mensual",
        data: monthly.map(([, value]) => value),
        backgroundColor: "#2563eb",
        borderRadius: 5,
        maxBarThickness: 55
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: compact,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: {
            title: items => months[items[0].dataIndex],
            label: valueTooltip
          }
        }
      },
      scales: {
        x: {
          grid: { display: false },
          ticks: { maxRotation: 0, autoSkip: false }
        },
        y: {
          beginAtZero: true,
          title: {
            display: !compact,
            text: "Miles"
          },
          ticks: {
            callback: value => compact
              ? `${number.format(value / 1000)}k`
              : number.format(value / 1000)
          }
        }
      }
    }
  };
}

function renderSummary(data) {
  element("totalGasto").textContent =
    currency.format(data.totalSpend);

  element("totalFacturas").textContent =
    integer.format(data.invoiceCount);

  element("ticketPromedio").textContent =
    currency.format(data.totalSpend / data.invoiceCount);

  element("totalProductos").textContent =
    integer.format(products.length);

  makeChart(
    "summaryMonthly",
    "summaryMonthlyChart",
    monthlyChartConfig(data.monthly.slice(-8), true)
  );

  const categories = groupToEightyPercent(
    data.categories,
    "Resto"
  );

  makeChart("summaryCategory", "summaryCategoryChart", {
    type: "doughnut",
    data: {
      labels: categories.map(item => item.label),
      datasets: [{
        data: categories.map(item => item.value),
        backgroundColor: categories.map(
          (_, index) => palette(index)
        )
      }]
    },
    options: {
      responsive: true,
      maintainAspectRatio: true,
      plugins: {
        legend: { position: "bottom" },
        tooltip: {
          callbacks: {
            label: context =>
              `${context.label}: ${currency.format(context.raw)}`
          }
        }
      }
    }
  });
}

function renderMonthly(data) {
  makeChart(
    "monthly",
    "monthlyChart",
    monthlyChartConfig(data.monthly)
  );
}

/* =========================================================
   CATEGORÍA → SUBCATEGORÍA → PRODUCTO
========================================================= */

function renderCategories(data) {
  const grouped = groupToEightyPercent(
    data.categories,
    "Resto de categorías"
  );

  makeChart("category", "categoryChart", {
    type: "bar",
    data: {
      labels: grouped.map(item => item.label),
      datasets: [{
        label: "Gasto",
        data: grouped.map(item => item.value),
        backgroundColor: grouped.map(
          (_, index) => palette(index)
        ),
        borderRadius: 5
      }]
    },
    options: {
      indexAxis: "y",
      responsive: true,
      maintainAspectRatio: false,
      plugins: {
        legend: { display: false },
        tooltip: {
          callbacks: { label: valueTooltip }
        }
      },
      scales: {
        x: {
          beginAtZero: true,
          ticks: {
            callback: value =>
              `${number.format(value / 1000)}k`
          }
        }
      }
    }
  });

  const categoryMap = new Map();

  // No utiliza las métricas del diccionario.
  for (const row of data.rows) {
    if (!categoryMap.has(row.category)) {
      categoryMap.set(row.category, {
        name: row.category,
        spend: 0,
        subcategories: new Map()
      });
    }

    const category = categoryMap.get(row.category);
    category.spend += row.itemValue;

    if (!category.subcategories.has(row.subcategory)) {
      category.subcategories.set(row.subcategory, {
        name: row.subcategory,
        spend: 0,
        invoices: new Set(),
        products: new Map()
      });
    }

    const sub = category.subcategories.get(row.subcategory);
    sub.spend += row.itemValue;
    sub.invoices.add(row.invoiceId);

    if (!sub.products.has(row.product)) {
      sub.products.set(row.product, {
        label: row.product,
        value: 0,
        invoices: new Set()
      });
    }

    const product = sub.products.get(row.product);
    product.value += row.itemValue;
    product.invoices.add(row.invoiceId);
  }

  const categories = [...categoryMap.values()]
    .sort((a, b) => b.spend - a.spend);

  element("categoryDetails").innerHTML =
    categories.map(category => {
      const subcategories = [
        ...category.subcategories.values()
      ].sort((a, b) => b.spend - a.spend);

      const detail = subcategories.map(sub => {
        const groupedProducts = groupToEightyPercent(
          [...sub.products.values()],
          "Resto de productos"
        );

        const rows = groupedProducts.map(product => {
          const events = product.isRemainder
            ? null
            : product.invoices.size;

          return `
            <tr>
              <td>${escapeHtml(product.label)}</td>
              <td style="text-align:right">
                ${currency.format(product.value)}
              </td>
              <td>${events === null
                ? "—"
                : integer.format(events)}</td>
              <td>${events === null
                ? "—"
                : percent.format(
                    events / data.invoiceCount
                  )}</td>
            </tr>
          `;
        }).join("");

        return `
          <details style="margin:10px 12px">
            <summary style="flex-wrap:wrap">
              ${escapeHtml(sub.name)}
              <span>
                ${currency.format(sub.spend)}
                · ${integer.format(sub.invoices.size)} facturas
                · ${percent.format(
                  sub.invoices.size / data.invoiceCount
                )} recurrencia
              </span>
            </summary>

            <div class="table-wrap">
              <table>
                <thead>
                  <tr>
                    <th>Producto</th>
                    <th style="text-align:right">Gasto</th>
                    <th>Facturas</th>
                    <th>Recurrencia global</th>
                  </tr>
                </thead>
                <tbody>${rows}</tbody>
              </table>
            </div>
          </details>
        `;
      }).join("");

      return `
        <details>
          <summary style="flex-wrap:wrap">
            ${escapeHtml(category.name)}
            <span>${currency.format(category.spend)}</span>
          </summary>
          ${detail}
        </details>
      `;
    }).join("");
}

/* =========================================================
   VARIACIÓN DE PRECIOS
========================================================= */

function renderPrices() {
  const changed = products.filter(product =>
    product.observationDays >= 2 &&
    Math.abs(product.priceChange) >= 0.05
  ).sort((a, b) =>
    Math.abs(b.priceChange) - Math.abs(a.priceChange)
  );

  const selected = changed.slice(0, 20);
  const canvas = element("priceChangeChart");
  const panel = canvas.closest(".panel");

  if (panel) {
    panel.style.height =
      `${Math.max(560, selected.length * 28 + 130)}px`;

    const subtitle = panel.querySelector(".panel-title span");
    if (subtitle) {
      subtitle.textContent =
        "Cambios de ±5% o más; gráfico con hasta 20 productos. " +
        "La tabla muestra todos. Precios antes de descuentos.";
    }
  }

  makeChart("priceChange", "priceChangeChart", {
    type: "bar",
    data: {
      labels: selected.map(product => product.product),
      datasets: [{
        label: "Variación de precio",
        data: selected.map(product => product.priceChange * 100),
        backgroundColor: selected.map(product =>
          product.priceChange >= 0 ? "#dc2626" : "#059669"
        ),
        borderRadius: 5
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
            label: context =>
              `${number.format(context.raw)}%`
          }
        }
      },
      scales: {
        x: {
          beginAtZero: true,
          ticks: {
            callback: value => `${number.format(value)}%`
          }
        },
        y: {
          ticks: { autoSkip: false }
        }
      }
    }
  });

  element("priceTable").innerHTML = changed.length
    ? changed.map(product => `
        <tr>
          <td>${escapeHtml(product.product)}</td>
          <td>${escapeHtml(
            `${product.category} · ${product.subcategory}`
          )}</td>
          <td>${dateFormat.format(
            product.firstObservation.date
          )}</td>
          <td>${currency.format(
            product.firstObservation.price
          )}</td>
          <td>${dateFormat.format(
            product.lastObservation.date
          )}</td>
          <td>${currency.format(
            product.lastObservation.price
          )}</td>
          <td class="${
            product.priceChange >= 0 ? "positive" : "negative"
          }">${percent.format(product.priceChange)}</td>
        </tr>
      `).join("")
    : `<tr><td colspan="7" style="text-align:left">
         No hay productos con dos fechas de precio y
         variación de ±5% o más.
       </td></tr>`;
}

/* =========================================================
   CANASTA HABITUAL
========================================================= */

function renderBasket(data) {
  // products contiene una sola entrada por nombre normalizado.
  const habitual = products.filter(product => product.habitual)
    .sort((a, b) => b.basketCost - a.basketCost);

  const baseline = habitual.reduce(
    (sum, product) => sum + product.basketCost,
    0
  );

  const months = data.monthly.map(([month]) => month);

  const monthly = months.map(month => {
    let imputed = 0;

    const value = habitual.reduce((sum, product) => {
      const prices = product.pricesByMonth.get(month);

      // Si no hay precio observado ese mes,
      // se utiliza la mediana general del producto.
      if (!prices?.length) imputed++;

      const price = prices?.length
        ? median(prices)
        : product.medianPrice;

      return sum + price * product.medianQuantity;
    }, 0);

    return {
      month,
      value,
      imputed,
      deviation: baseline ? (value - baseline) / baseline : 0
    };
  });

  const latest = monthly.at(-1) || {
    value: 0,
    deviation: 0,
    imputed: 0
  };

  element("basketBaseline").textContent =
    currency.format(baseline);

  element("basketLatest").textContent =
    currency.format(latest.value);

  element("basketDeviation").textContent =
    percent.format(latest.deviation);

  element("basketDeviation").className =
    latest.deviation > 0
      ? "positive"
      : latest.deviation < 0 ? "negative" : "";

  element("habitualProducts").textContent =
    integer.format(habitual.length);

  const panel = element("basketChart").closest(".panel");
  const subtitle = panel?.querySelector(".panel-title span");

  if (subtitle) {
    subtitle.textContent =
      "Recurrencia ≥40%; cantidades fijas y precios medianos. " +
      "Valor estimado, no gasto real mensual. " +
      `Último mes: ${latest.imputed} de ${habitual.length} ` +
      "productos sin precio mensual; se usa su mediana general.";
  }

  makeChart("basket", "basketChart", {
    data: {
      labels: groupedMonthLabels(months),
      datasets: [
        {
          type: "line",
          label: "Valor estimado",
          data: monthly.map(item => item.value),
          borderColor: "#2563eb",
          backgroundColor: "rgba(37,99,235,.10)",
          fill: true,
          tension: 0.3,
          yAxisID: "y"
        },
        {
          type: "bar",
          label: "Desviación frente a base",
          data: monthly.map(item => item.deviation * 100),
          backgroundColor: monthly.map(item =>
            item.deviation >= 0
              ? "rgba(220,38,38,.35)"
              : "rgba(5,150,105,.35)"
          ),
          yAxisID: "y1"
        }
      ]
    },
    options: {
      responsive: true,
      maintainAspectRatio: false,
      interaction: { mode: "index", intersect: false },
      plugins: {
        tooltip: {
          callbacks: {
            title: items => months[items[0].dataIndex],
            label: context =>
              `${context.dataset.label}: ` +
              (context.dataset.yAxisID === "y1"
                ? `${number.format(context.raw)}%`
                : currency.format(context.raw)),
            afterBody: items => {
              const item = monthly[items[0].dataIndex];
              return (
                `${item.imputed} productos valorados ` +
                "con su mediana general."
              );
            }
          }
        }
      },
      scales: {
        x: {
          ticks: { maxRotation: 0, autoSkip: false }
        },
        y: {
          position: "left",
          beginAtZero: false,
          ticks: {
            callback: value => number.format(value)
          }
        },
        y1: {
          position: "right",
          beginAtZero: true,
          grid: { drawOnChartArea: false },
          ticks: {
            callback: value => `${number.format(value)}%`
          }
        }
      }
    }
  });

  element("basketTable").innerHTML = habitual.length
    ? habitual.map(product => `
        <tr>
          <td>${escapeHtml(product.product)}</td>
          <td>${escapeHtml(
            `${product.category} · ${product.subcategory}`
          )}</td>
          <td>${percent.format(product.recurrence)}</td>
          <td>${number.format(product.medianQuantity)}</td>
          <td>${currency.format(product.medianPrice)}</td>
          <td>${currency.format(product.basketCost)}</td>
        </tr>
      `).join("")
    : `<tr><td colspan="6" style="text-align:left">
         No hay productos con recurrencia de al menos 40%.
       </td></tr>`;
}

/* =========================================================
   PRESENTACIÓN Y PESTAÑAS
========================================================= */

function renderDashboard(data, workbookItem, readResult) {
  element("periodo").textContent =
    `${dateFormat.format(data.dates[0])} — ` +
    `${dateFormat.format(data.dates.at(-1))}`;

  element("sourceInfo").textContent =
    `${workbookItem.name} · ` +
    `${integer.format(data.rows.length)} líneas válidas de ` +
    `${integer.format(readResult.sourceCount)} · ` +
    `archivo modificado ${
      dateFormat.format(new Date(workbookItem.lastModifiedDateTime))
    }`;

  // Mostrar antes de crear gráficos permite medir el resumen.
  dashboard.hidden = false;
  tabs.hidden = false;

  renderSummary(data);
  renderMonthly(data);
  renderCategories(data);
  renderPrices();
  renderBasket(data);

  requestAnimationFrame(resizeCharts);
}

document.querySelectorAll(".tab").forEach(button => {
  button.addEventListener("click", () => {
    document.querySelectorAll(".tab").forEach(tab =>
      tab.classList.remove("active")
    );

    document.querySelectorAll(".tab-panel").forEach(panel =>
      panel.classList.remove("active")
    );

    button.classList.add("active");
    element(button.dataset.tab).classList.add("active");

    requestAnimationFrame(resizeCharts);
  });
});

/* =========================================================
   CARGAR EL DASHBOARD
========================================================= */

async function loadDashboard() {
  setLoading(true);
  setStatus("Localizando el Excel en OneDrive…");

  try {
    if (!window.XLSX || !window.Chart) {
      throw new Error(
        "No se cargaron las librerías de Excel o gráficos. " +
        "Recarga la página y comprueba tu conexión."
      );
    }

    const token = await getAccessToken();

    const profileResponse = await graphFetch(
      "https://graph.microsoft.com/v1.0/me" +
      "?$select=displayName,userPrincipalName",
      token
    );

    const profile = await profileResponse.json();

    element("usuario").textContent =
      profile.displayName || profile.userPrincipalName || "";

    const workbookItem = await findWorkbook(token);

    setStatus("Descargando la última versión guardada del Excel…");

    const buffer = await downloadWorkbook(token, workbookItem.id);
    const readResult = readInvoiceRows(buffer);
    const data = aggregateData(readResult.rows);

    if (!data.invoiceCount) {
      throw new Error("El Excel no contiene facturas válidas.");
    }

    setStatus("Preparando categorías, subcategorías y productos…");

    renderDashboard(data, workbookItem, readResult);

    const warnings = [];

    if (readResult.skipped) {
      warnings.push(
        `${readResult.skipped} líneas omitidas por falta de ` +
        "ID, fecha válida o producto"
      );
    }

    if (data.conflictingProducts.length) {
      warnings.push(
        `${data.conflictingProducts.length} productos tienen ` +
        "más de una categoría o subcategoría; revisa el diccionario"
      );
    }

    const pending = new Set(
      data.rows.filter(row =>
        ["por revisar", "por clasificar", "sin categoría", "revisar"]
          .includes(row.subcategory.toLowerCase()) ||
        ["por revisar", "por clasificar", "sin categoría", "revisar"]
          .includes(row.category.toLowerCase())
      ).map(row => row.description)
    ).size;

    if (pending) {
      warnings.push(
        `${pending} descripciones con clasificación pendiente`
      );
    }

    setStatus(
      `Dashboard actualizado: ${integer.format(data.invoiceCount)} ` +
      `facturas, ${integer.format(data.rows.length)} líneas y ` +
      `${integer.format(products.length)} productos.` +
      (warnings.length ? ` Avisos: ${warnings.join("; ")}.` : ""),
      warnings.length ? "" : "success"
    );
  } catch (error) {
    console.error(error);
    dashboard.hidden = true;
    tabs.hidden = true;

    setStatus(
      error.message || "No fue posible cargar el dashboard.",
      "error"
    );
  } finally {
    setLoading(false);
  }
}

/* =========================================================
   INICIAR Y CERRAR SESIÓN
========================================================= */

async function login() {
  setLoading(true);

  try {
    const response = await msalInstance.loginPopup({
      scopes: CONFIG.scopes,
      prompt: "select_account"
    });

    account = response.account;
    msalInstance.setActiveAccount(account);
    showSession(true);

    await loadDashboard();
  } catch (error) {
    console.error(error);
    setStatus(
      error.message || "No fue posible iniciar sesión.",
      "error"
    );
  } finally {
    setLoading(false);
  }
}

async function logout() {
  setLoading(true);

  try {
    await msalInstance.logoutPopup({
      account,
      postLogoutRedirectUri: CONFIG.redirectUri
    });

    account = null;
    products = [];

    Object.keys(charts).forEach(destroyChart);

    dashboard.hidden = true;
    tabs.hidden = true;
    showSession(false);

    element("usuario").textContent = "";
    element("periodo").textContent =
      "Inicia sesión para cargar los datos.";

    setStatus("Sesión cerrada.");
  } catch (error) {
    console.error(error);
    setStatus(
      error.message || "No fue posible cerrar sesión.",
      "error"
    );
  } finally {
    setLoading(false);
  }
}

/* =========================================================
   ARRANQUE
========================================================= */

async function start() {
  setLoading(true);

  try {
    await msalInstance.initialize();

    loginButton.addEventListener("click", login);
    refreshButton.addEventListener("click", loadDashboard);
    logoutButton.addEventListener("click", logout);

    account =
      msalInstance.getActiveAccount() ||
      msalInstance.getAllAccounts()[0] ||
      null;

    showSession(Boolean(account));

    if (account) {
      msalInstance.setActiveAccount(account);
      await loadDashboard();
    }
  } catch (error) {
    console.error(error);
    setStatus(
      error.message || "No se pudo inicializar la aplicación.",
      "error"
    );
  } finally {
    setLoading(false);
  }
}

start();
