import {
  PublicClientApplication,
  InteractionRequiredAuthError
} from "https://cdn.jsdelivr.net/npm/@azure/msal-browser@4.29.0/+esm";

import { CONFIG } from "./config.js";

/* =========================================================
   FORMATOS
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

const monthNames = [
  "Ene", "Feb", "Mar", "Abr", "May", "Jun",
  "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"
];

/* =========================================================
   AUTENTICACIÓN
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

await msalInstance.initialize();

/* =========================================================
   ELEMENTOS HTML
========================================================= */

const loginButton = document.getElementById("loginButton");
const logoutButton = document.getElementById("logoutButton");
const refreshButton = document.getElementById("refreshButton");
const statusBox = document.getElementById("status");
const dashboard = document.getElementById("dashboard");
const tabs = document.getElementById("tabs");

let account = null;
let products = [];
let charts = {};

/* =========================================================
   FUNCIONES GENERALES
========================================================= */

function element(id) {
  const result = document.getElementById(id);

  if (!result) {
    throw new Error(
      `Falta el elemento "${id}" en index.html. ` +
      "Comprueba que index.html y app.js sean las versiones nuevas."
    );
  }

  return result;
}

function setStatus(message, type = "") {
  statusBox.textContent = message;
  statusBox.className = `status ${type}`.trim();
}

function setLoading(loading) {
  loginButton.disabled = loading;
  refreshButton.disabled = loading;

  loginButton.textContent = loading
    ? "Conectando..."
    : "Iniciar sesión";

  refreshButton.textContent = loading
    ? "Actualizando..."
    : "Actualizar";
}

function asNumber(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : 0;
  }

  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return 0;
  }

  const result = Number(
    String(value)
      .replace(/\$/g, "")
      .replace(/,/g, "")
      .trim()
  );

  return Number.isFinite(result) ? result : 0;
}

function asDate(value) {
  if (
    value instanceof Date &&
    !Number.isNaN(value.getTime())
  ) {
    return value;
  }

  if (typeof value === "number") {
    const parsed = XLSX.SSF.parse_date_code(value);

    if (parsed) {
      return new Date(
        parsed.y,
        parsed.m - 1,
        parsed.d,
        12
      );
    }
  }

  const text = String(value ?? "").trim();

  const match = text.match(
    /^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/
  );

  if (match) {
    return new Date(
      Number(match[3]),
      Number(match[2]) - 1,
      Number(match[1]),
      12
    );
  }

  const parsed = new Date(text);

  return Number.isNaN(parsed.getTime())
    ? null
    : parsed;
}

function dateKey(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");
  const day = String(date.getDate()).padStart(2, "0");

  return `${year}-${month}-${day}`;
}

function monthKey(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, "0");

  return `${year}-${month}`;
}

function median(values) {
  const clean = values
    .map(asNumber)
    .filter(Number.isFinite)
    .sort((a, b) => a - b);

  if (!clean.length) return 0;

  const middle = Math.floor(clean.length / 2);

  return clean.length % 2
    ? clean[middle]
    : (clean[middle - 1] + clean[middle]) / 2;
}

function averageDays(dates) {
  const uniqueDates = [
    ...new Set(
      dates.filter(Boolean).map(dateKey)
    )
  ]
    .sort()
    .map(value => {
      const [year, month, day] = value
        .split("-")
        .map(Number);

      return new Date(year, month - 1, day, 12);
    });

  if (uniqueDates.length < 2) return null;

  const differences = [];

  for (let i = 1; i < uniqueDates.length; i++) {
    differences.push(
      (uniqueDates[i] - uniqueDates[i - 1]) /
      86400000
    );
  }

  return differences.reduce(
    (sum, value) => sum + value,
    0
  ) / differences.length;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function destroyChart(name) {
  if (charts[name]) {
    charts[name].destroy();
    charts[name] = null;
  }
}

function palette(index) {
  const colors = [
    "#2563eb",
    "#0891b2",
    "#059669",
    "#d97706",
    "#7c3aed",
    "#dc2626",
    "#64748b",
    "#db2777",
    "#65a30d",
    "#4f46e5",
    "#ea580c"
  ];

  return colors[index % colors.length];
}

/* =========================================================
   ETIQUETAS DE MES Y AÑO
========================================================= */

function groupedMonthLabels(months) {
  const yearGroups = new Map();

  months.forEach((month, index) => {
    const year = month.slice(0, 4);

    if (!yearGroups.has(year)) {
      yearGroups.set(year, []);
    }

    yearGroups.get(year).push(index);
  });

  const yearCenters = new Map();

  for (const [year, indexes] of yearGroups.entries()) {
    yearCenters.set(
      Math.floor(
        (indexes[0] + indexes[indexes.length - 1]) / 2
      ),
      year
    );
  }

  return months.map((month, index) => {
    const monthNumber = Number(month.slice(5, 7));
    const year = yearCenters.get(index) || "";

    return [
      monthNames[monthNumber - 1],
      year
    ];
  });
}

/* =========================================================
   AUTENTICACIÓN Y GRAPH
========================================================= */

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
    if (error instanceof InteractionRequiredAuthError) {
      const response =
        await msalInstance.acquireTokenPopup(request);

      return response.accessToken;
    }

    throw error;
  }
}

async function graphFetch(url, token, options = {}) {
  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      ...(options.headers || {})
    }
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
    throw new Error(
      "Falta workbookPath en config.js."
    );
  }

  const encodedPath = CONFIG.workbookPath
    .split("/")
    .map(segment => encodeURIComponent(segment))
    .join("/");

  const url =
    "https://graph.microsoft.com/v1.0/me/drive/" +
    `special/documents:/${encodedPath}:` +
    "?$select=id,name,size,lastModifiedDateTime," +
    "parentReference,webUrl";

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
   LEER EXCEL
========================================================= */

function readInvoiceRows(buffer) {
  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: true
  });

  const sheet = workbook.Sheets[CONFIG.worksheetName];

  if (!sheet) {
    throw new Error(
      `No existe la pestaña "${CONFIG.worksheetName}".`
    );
  }

  const rows = XLSX.utils.sheet_to_json(sheet, {
    defval: null,
    raw: true
  });

  const required = [
    "ID factura",
    "Fecha Factura",
    "Descripción",
    "Cantidad",
    "Valor Unitario",
    "Valor Item",
    "Total Factura",
    "Producto normalizado",
    "Categoría"
  ];

  const headers = rows.length
    ? Object.keys(rows[0])
    : [];

  const missing = required.filter(
    header => !headers.includes(header)
  );

  if (missing.length) {
    throw new Error(
      `Faltan columnas: ${missing.join(", ")}`
    );
  }

  return rows
    .map(row => ({
      invoiceId: String(
        row["ID factura"] ?? ""
      ).trim(),

      date: asDate(row["Fecha Factura"]),

      description: String(
        row["Descripción"] ?? ""
      ).trim(),

      quantity: asNumber(row["Cantidad"]),

      unitPrice: asNumber(
        row["Valor Unitario"]
      ),

      itemValue: asNumber(
        row["Valor Item"]
      ),

      invoiceTotal: asNumber(
        row["Total Factura"]
      ),

      product: String(
        row["Producto normalizado"] ||
        row["Descripción"] ||
        "Sin producto"
      ).trim(),

      category: String(
        row["Categoría"] ||
        "Sin categoría"
      ).trim()
    }))
    .filter(row =>
      row.invoiceId &&
      row.date &&
      row.product
    );
}

/* =========================================================
   AGRUPAR HASTA 80%
========================================================= */

function groupToEightyPercent(entries, remainderLabel) {
  if (!entries.length) return [];

  const sorted = [...entries]
    .sort((a, b) => b.value - a.value);

  const total = sorted.reduce(
    (sum, entry) => sum + entry.value,
    0
  );

  let cumulative = 0;
  const visible = [];
  const remainder = [];

  for (const entry of sorted) {
    if (
      cumulative / total < 0.8 ||
      visible.length === 0
    ) {
      visible.push(entry);
      cumulative += entry.value;
    } else {
      remainder.push(entry);
    }
  }

  const remainderValue = remainder.reduce(
    (sum, entry) => sum + entry.value,
    0
  );

  if (remainderValue > 0) {
    visible.push({
      label: remainderLabel,
      value: remainderValue
    });
  }

  return visible;
}

/* =========================================================
   CÁLCULOS PRINCIPALES
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
      (categoryMap.get(row.category) || 0) +
      row.itemValue
    );

    if (!productMap.has(row.product)) {
      productMap.set(row.product, {
        product: row.product,
        category: row.category,
        invoices: new Set(),
        dates: [],
        quantity: 0,
        spend: 0,
        prices: [],
        quantitiesByInvoice: new Map(),
        pricesByMonth: new Map(),
        priceObservations: []
      });
    }

    const product = productMap.get(row.product);

    product.invoices.add(row.invoiceId);
    product.dates.push(row.date);
    product.quantity += row.quantity;
    product.spend += row.itemValue;

    product.quantitiesByInvoice.set(
      row.invoiceId,
      (
        product.quantitiesByInvoice.get(
          row.invoiceId
        ) || 0
      ) + row.quantity
    );

    if (row.unitPrice > 0) {
      product.prices.push(row.unitPrice);

      product.priceObservations.push({
        date: row.date,
        price: row.unitPrice
      });

      const month = monthKey(row.date);

      if (!product.pricesByMonth.has(month)) {
        product.pricesByMonth.set(month, []);
      }

      product.pricesByMonth
        .get(month)
        .push(row.unitPrice);
    }
  }

  for (const invoice of invoices.values()) {
    const month = monthKey(invoice.date);

    monthlyMap.set(
      month,
      (monthlyMap.get(month) || 0) +
      invoice.total
    );
  }

  const invoiceCount = invoices.size;

  products = [...productMap.values()]
    .map(product => {
      const events = product.invoices.size;
      const recurrence = invoiceCount
        ? events / invoiceCount
        : 0;

      const medianPrice = median(product.prices);

      const medianQuantity = median([
        ...product.quantitiesByInvoice.values()
      ]);

      const observations = product.priceObservations
        .sort((a, b) => a.date - b.date);

      const firstObservation = observations[0] || null;
      const lastObservation =
        observations[observations.length - 1] || null;

      const priceChange =
        firstObservation &&
        lastObservation &&
        firstObservation.price > 0
          ? (
              lastObservation.price -
              firstObservation.price
            ) / firstObservation.price
          : 0;

      return {
        ...product,
        events,
        recurrence,
        medianPrice,
        medianQuantity,
        basketCost: medianPrice * medianQuantity,
        averageDays: averageDays(product.dates),
        habitual: recurrence >= 0.4,
        firstObservation,
        lastObservation,
        priceChange
      };
    })
    .sort((a, b) => b.spend - a.spend);

  const totalSpend = [...invoices.values()]
    .reduce(
      (sum, invoice) => sum + invoice.total,
      0
    );

  const dates = [...invoices.values()]
    .map(invoice => invoice.date)
    .sort((a, b) => a - b);

  const monthly = [...monthlyMap.entries()]
    .sort(([a], [b]) => a.localeCompare(b));

  const categories = [...categoryMap.entries()]
    .map(([label, value]) => ({
      label,
      value
    }))
    .sort((a, b) => b.value - a.value);

  return {
    rows,
    invoices,
    invoiceCount,
    totalSpend,
    dates,
    monthly,
    categories
  };
}

/* =========================================================
   RESUMEN
========================================================= */

function renderSummary(data) {
  element("totalGasto").textContent =
    currency.format(data.totalSpend);

  element("totalFacturas").textContent =
    integer.format(data.invoiceCount);

  element("ticketPromedio").textContent =
    currency.format(
      data.totalSpend / data.invoiceCount
    );

  element("totalProductos").textContent =
    integer.format(products.length);

  destroyChart("summaryMonthly");

  const recentMonths = data.monthly.slice(-8);

  charts.summaryMonthly = new Chart(
    element("summaryMonthlyChart"),
    {
      type: "bar",
      data: {
        labels: groupedMonthLabels(
          recentMonths.map(([month]) => month)
        ),
        datasets: [{
          label: "Gasto",
          data: recentMonths.map(
            ([, value]) => value
          ),
          backgroundColor: "#2563eb",
          borderRadius: 5
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: true,
        plugins: {
          legend: {
            display: false
          },
          tooltip: {
            callbacks: {
              label: context =>
                currency.format(context.raw)
            }
          }
        },
        scales: {
          y: {
            beginAtZero: true,
            ticks: {
              callback: value =>
                `${number.format(value / 1000)}k`
            }
          }
        }
      }
    }
  );

  destroyChart("summaryCategory");

  const summaryCategories =
    groupToEightyPercent(
      data.categories,
      "Resto"
    );

  charts.summaryCategory = new Chart(
    element("summaryCategoryChart"),
    {
      type: "doughnut",
      data: {
        labels: summaryCategories.map(
          item => item.label
        ),
        datasets: [{
          data: summaryCategories.map(
            item => item.value
          ),
          backgroundColor:
            summaryCategories.map(
              (_, index) => palette(index)
            )
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: true,
        plugins: {
          legend: {
            position: "bottom"
          },
          tooltip: {
            callbacks: {
              label: context =>
                `${context.label}: ` +
                currency.format(context.raw)
            }
          }
        }
      }
    }
  );
}

/* =========================================================
   GASTO MENSUAL
========================================================= */

function renderMonthly(data) {
  destroyChart("monthly");

  const months = data.monthly.map(
    ([month]) => month
  );

  charts.monthly = new Chart(
    element("monthlyChart"),
    {
      type: "bar",
      data: {
        labels: groupedMonthLabels(months),
        datasets: [{
          label: "Gasto mensual",
          data: data.monthly.map(
            ([, value]) => value
          ),
          backgroundColor: "#2563eb",
          borderRadius: 6,
          maxBarThickness: 55
        }]
      },
      options: {
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            display: false
          },
          tooltip: {
            callbacks: {
              label: context =>
                currency.format(context.raw)
            }
          }
        },
        scales: {
          x: {
            grid: {
              display: false
            },
            ticks: {
              maxRotation: 0,
              autoSkip: false
            }
          },
          y: {
            beginAtZero: true,
            title: {
              display: true,
              text: "Miles"
            },
            ticks: {
              callback: value =>
                number.format(value / 1000)
            }
          }
        }
      }
    }
  );
}

/* =========================================================
   CATEGORÍAS Y PRODUCTOS
========================================================= */

function renderCategories(data) {
  const groupedCategories =
    groupToEightyPercent(
      data.categories,
      "20% restante"
    );

  destroyChart("category");

  charts.category = new Chart(
    element("categoryChart"),
    {
      type: "bar",
      data: {
        labels: groupedCategories.map(
          item => item.label
        ),
        datasets: [{
          label: "Gasto",
          data: groupedCategories.map(
            item => item.value
          ),
          backgroundColor:
            groupedCategories.map(
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
          legend: {
            display: false
          },
          tooltip: {
            callbacks: {
              label: context =>
                currency.format(context.raw)
            }
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
    }
  );

  const categories = [...new Set(
    products.map(product => product.category)
  )]
    .map(category => ({
      category,
      spend: products
        .filter(
          product => product.category === category
        )
        .reduce(
          (sum, product) => sum + product.spend,
          0
        )
    }))
    .sort((a, b) => b.spend - a.spend);

  element("categoryDetails").innerHTML =
    categories.map(categoryData => {
      const categoryProducts = products
        .filter(
          product =>
            product.category === categoryData.category
        )
        .map(product => ({
          label: product.product,
          value: product.spend,
          events: product.events,
          recurrence: product.recurrence
        }))
        .sort((a, b) => b.value - a.value);

      const groupedProducts =
        groupToEightyPercent(
          categoryProducts,
          "20% restante"
        );

      const rows = groupedProducts.map(item => {
        const original = categoryProducts.find(
          product => product.label === item.label
        );

        return `
          <tr>
            <td>${escapeHtml(item.label)}</td>
            <td>${currency.format(item.value)}</td>
            <td>${
              original
                ? integer.format(original.events)
                : "—"
            }</td>
            <td>${
              original
                ? percent.format(original.recurrence)
                : "—"
            }</td>
          </tr>
        `;
      }).join("");

      return `
        <details>
          <summary>
            ${escapeHtml(categoryData.category)}
            <span>${currency.format(categoryData.spend)}</span>
          </summary>

          <div class="table-wrap">
            <table>
              <thead>
                <tr>
                  <th>Producto</th>
                  <th>Gasto</th>
                  <th>Eventos</th>
                  <th>Recurrencia</th>
                </tr>
              </thead>
              <tbody>${rows}</tbody>
            </table>
          </div>
        </details>
      `;
    }).join("");
}

/* =========================================================
   CAMBIOS DE PRECIO
========================================================= */

function renderPrices() {
  const changedProducts = products
    .filter(product =>
      product.firstObservation &&
      product.lastObservation &&
      Math.abs(product.priceChange) >= 0.05
    )
    .sort(
      (a, b) =>
        Math.abs(b.priceChange) -
        Math.abs(a.priceChange)
    );

  const chartProducts =
    changedProducts.slice(0, 20);

  destroyChart("priceChange");

  charts.priceChange = new Chart(
    element("priceChangeChart"),
    {
      type: "bar",
      data: {
        labels: chartProducts.map(
          product => product.product
        ),
        datasets: [{
          label: "Variación",
          data: chartProducts.map(
            product => product.priceChange * 100
          ),
          backgroundColor: chartProducts.map(
            product =>
              product.priceChange >= 0
                ? "#dc2626"
                : "#059669"
          ),
          borderRadius: 5
        }]
      },
      options: {
        indexAxis: "y",
        responsive: true,
        maintainAspectRatio: false,
        plugins: {
          legend: {
            display: false
          },
          tooltip: {
            callbacks: {
              label: context =>
                `${number.format(context.raw)}%`
            }
          }
        },
        scales: {
          x: {
            ticks: {
              callback: value => `${value}%`
            }
          }
        }
      }
    }
  );

  element("priceTable").innerHTML =
    changedProducts.map(product => `
      <tr>
        <td>${escapeHtml(product.product)}</td>
        <td>${escapeHtml(product.category)}</td>
        <td>${
          dateFormat.format(
            product.firstObservation.date
          )
        }</td>
        <td>${
          currency.format(
            product.firstObservation.price
          )
        }</td>
        <td>${
          dateFormat.format(
            product.lastObservation.date
          )
        }</td>
        <td>${
          currency.format(
            product.lastObservation.price
          )
        }</td>
        <td class="${
          product.priceChange >= 0
            ? "positive"
            : "negative"
        }">
          ${percent.format(product.priceChange)}
        </td>
      </tr>
    `).join("");
}

/* =========================================================
   CANASTA HABITUAL
========================================================= */

function renderBasket(data) {
  const habitualProducts = products
    .filter(product => product.habitual)
    .sort(
      (a, b) =>
        b.basketCost - a.basketCost
    );

  const baseline = habitualProducts.reduce(
    (sum, product) =>
      sum + product.basketCost,
    0
  );

  const months = data.monthly.map(
    ([month]) => month
  );

  /*
    Cuando un producto habitual no aparece en un mes,
    se usa su precio mediano general. Esto permite comparar
    una canasta completa y estable todos los meses.
  */
  const monthlyBasket = months.map(month => {
    const value = habitualProducts.reduce(
      (sum, product) => {
        const observedPrices =
          product.pricesByMonth.get(month);

        const monthlyPrice =
          observedPrices?.length
            ? median(observedPrices)
            : product.medianPrice;

        return sum +
          monthlyPrice *
          product.medianQuantity;
      },
      0
    );

    return {
      month,
      value,
      deviation:
        baseline > 0
          ? (value - baseline) / baseline
          : 0
    };
  });

  const latest =
    monthlyBasket[monthlyBasket.length - 1] || {
      value: 0,
      deviation: 0
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
      : latest.deviation < 0
        ? "negative"
        : "";

  element("habitualProducts").textContent =
    integer.format(habitualProducts.length);

  destroyChart("basket");

  charts.basket = new Chart(
    element("basketChart"),
    {
      data: {
        labels: groupedMonthLabels(months),
        datasets: [
          {
            type: "line",
            label: "Valor de la canasta",
            data: monthlyBasket.map(
              item => item.value
            ),
            borderColor: "#2563eb",
            backgroundColor:
              "rgba(37,99,235,.10)",
            fill: true,
            tension: 0.3,
            yAxisID: "y"
          },
          {
            type: "bar",
            label: "Desviación",
            data: monthlyBasket.map(
              item => item.deviation * 100
            ),
            backgroundColor:
              monthlyBasket.map(item =>
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
        interaction: {
          mode: "index",
          intersect: false
        },
        plugins: {
          tooltip: {
            callbacks: {
              label: context => {
                if (context.dataset.yAxisID === "y1") {
                  return (
                    `${context.dataset.label}: ` +
                    `${number.format(context.raw)}%`
                  );
                }

                return (
                  `${context.dataset.label}: ` +
                  currency.format(context.raw)
                );
              }
            }
          }
        },
        scales: {
          x: {
            ticks: {
              maxRotation: 0,
              autoSkip: false
            }
          },
          y: {
            beginAtZero: false,
            position: "left",
            ticks: {
              callback: value =>
                currency.format(value)
            }
          },
          y1: {
            position: "right",
            grid: {
              drawOnChartArea: false
            },
            ticks: {
              callback: value => `${value}%`
            }
          }
        }
      }
    }
  );

  element("basketTable").innerHTML =
    habitualProducts.map(product => `
      <tr>
        <td>${escapeHtml(product.product)}</td>
        <td>${escapeHtml(product.category)}</td>
        <td>${percent.format(product.recurrence)}</td>
        <td>${number.format(product.medianQuantity)}</td>
        <td>${currency.format(product.medianPrice)}</td>
        <td>${currency.format(product.basketCost)}</td>
      </tr>
    `).join("");
}

/* =========================================================
   PRESENTAR TODO
========================================================= */

function renderDashboard(data, workbookItem) {
  const firstDate = data.dates[0];
  const lastDate =
    data.dates[data.dates.length - 1];

  element("periodo").textContent =
    `${dateFormat.format(firstDate)} — ` +
    `${dateFormat.format(lastDate)}`;

  element("sourceInfo").textContent =
    `${workbookItem.name} · ` +
    `${data.rows.length.toLocaleString("es-PA")} líneas · ` +
    `actualizado ${
      dateFormat.format(
        new Date(workbookItem.lastModifiedDateTime)
      )
    }`;

  renderSummary(data);
  renderMonthly(data);
  renderCategories(data);
  renderPrices();
  renderBasket(data);

  dashboard.hidden = false;
  tabs.hidden = false;
}

/* =========================================================
   PESTAÑAS
========================================================= */

function configureTabs() {
  document.querySelectorAll(".tab").forEach(button => {
    button.addEventListener("click", () => {
      document.querySelectorAll(".tab")
        .forEach(tab => tab.classList.remove("active"));

      document.querySelectorAll(".tab-panel")
        .forEach(panel => panel.classList.remove("active"));

      button.classList.add("active");

      const panel = document.getElementById(
        button.dataset.tab
      );

      if (panel) {
        panel.classList.add("active");
      }

      setTimeout(() => {
        Object.values(charts).forEach(chart => {
          if (chart) chart.resize();
        });
      }, 50);
    });
  });
}

configureTabs();

/* =========================================================
   CARGAR DASHBOARD
========================================================= */

async function loadDashboard() {
  setLoading(true);
  setStatus(
    "Localizando el Excel en OneDrive…"
  );

  try {
    const token = await getAccessToken();

    const profileResponse = await graphFetch(
      "https://graph.microsoft.com/v1.0/me" +
      "?$select=displayName,userPrincipalName",
      token
    );

    const profile = await profileResponse.json();

    element("usuario").textContent =
      profile.displayName ||
      profile.userPrincipalName;

    const workbookItem =
      await findWorkbook(token);

    setStatus(
      "Descargando la última versión del Excel…"
    );

    const buffer = await downloadWorkbook(
      token,
      workbookItem.id
    );

    setStatus(
      "Procesando facturas y preparando los análisis…"
    );

    const rows = readInvoiceRows(buffer);
    const data = aggregateData(rows);

    if (!data.invoiceCount) {
      throw new Error(
        "El Excel no contiene facturas válidas."
      );
    }

    renderDashboard(data, workbookItem);

    setStatus(
      `Dashboard actualizado: ` +
      `${integer.format(data.invoiceCount)} facturas, ` +
      `${integer.format(rows.length)} líneas y ` +
      `${integer.format(products.length)} productos.`,
      "success"
    );
  } catch (error) {
    console.error(error);

    setStatus(
      error.message ||
      "No fue posible cargar el dashboard.",
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
  try {
    const response =
      await msalInstance.loginPopup({
        scopes: CONFIG.scopes,
        prompt: "select_account"
      });

    account = response.account;
    msalInstance.setActiveAccount(account);

    loginButton.hidden = true;
    logoutButton.hidden = false;
    refreshButton.hidden = false;

    await loadDashboard();
  } catch (error) {
    console.error(error);

    setStatus(
      error.message ||
      "No fue posible iniciar sesión.",
      "error"
    );

    setLoading(false);
  }
}

async function logout() {
  await msalInstance.logoutPopup({
    account,
    postLogoutRedirectUri:
      CONFIG.redirectUri
  });

  account = null;
  dashboard.hidden = true;
  tabs.hidden = true;

  loginButton.hidden = false;
  logoutButton.hidden = true;
  refreshButton.hidden = true;

  element("usuario").textContent = "";

  setStatus(
    "Sesión cerrada. Conecta tu cuenta Microsoft para cargar los datos."
  );
}

/* =========================================================
   EVENTOS
========================================================= */

loginButton.addEventListener("click", login);
refreshButton.addEventListener("click", loadDashboard);
logoutButton.addEventListener("click", logout);

/* =========================================================
   RECUPERAR SESIÓN EXISTENTE
========================================================= */

const existingAccounts =
  msalInstance.getAllAccounts();

if (existingAccounts.length) {
  account = existingAccounts[0];
  msalInstance.setActiveAccount(account);

  loginButton.hidden = true;
  logoutButton.hidden = false;
  refreshButton.hidden = false;

  loadDashboard();
}
