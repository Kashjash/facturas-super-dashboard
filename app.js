import {
  PublicClientApplication,
  InteractionRequiredAuthError
} from "https://cdn.jsdelivr.net/npm/@azure/msal-browser@4.29.0/+esm";

import { CONFIG } from "./config.js";

// FORMATO
const currency = new Intl.NumberFormat("es-PA", {
  style: "currency",
  currency: "USD"
});

const number = new Intl.NumberFormat("es-PA", {
  maximumFractionDigits: 2
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

// AUTENTICACIÓN
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

// ELEMENTOS HTML
const loginButton = document.getElementById("loginButton");
const logoutButton = document.getElementById("logoutButton");
const refreshButton = document.getElementById("refreshButton");
const statusBox = document.getElementById("status");
const dashboard = document.getElementById("dashboard");
const productSearch = document.getElementById("productSearch");

let account = null;
let products = [];
let charts = {};

// ESTADO
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

// CONVERSIONES
function asNumber(value) {
  if (typeof value === "number") return value;

  if (
    value === null ||
    value === undefined ||
    value === ""
  ) {
    return 0;
  }

  const cleaned = String(value)
    .replace(/\$/g, "")
    .replace(/,/g, "")
    .trim();

  const result = Number(cleaned);

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
      dates
        .filter(Boolean)
        .map(dateKey)
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

// TOKEN
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

// MICROSOFT GRAPH
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

// LOCALIZAR EL EXCEL POR SU RUTA EXACTA
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
  const workbookItem = await response.json();

  if (
    workbookItem.name?.toLowerCase() !==
    CONFIG.workbookName.toLowerCase()
  ) {
    throw new Error(
      `La ruta encontrada no corresponde a "${CONFIG.workbookName}".`
    );
  }

  return workbookItem;
}

// DESCARGAR EL EXCEL
async function downloadWorkbook(token, itemId) {
  const url =
    "https://graph.microsoft.com/v1.0/me/drive/items/" +
    `${encodeURIComponent(itemId)}/content`;

  const response = await graphFetch(url, token);

  return response.arrayBuffer();
}

// LEER LA PESTAÑA DE FACTURAS
function readInvoiceRows(buffer) {
  const workbook = XLSX.read(buffer, {
    type: "array",
    cellDates: true
  });

  const sheet = workbook.Sheets[CONFIG.worksheetName];

  if (!sheet) {
    throw new Error(
      `No existe la pestaña "${CONFIG.worksheetName}" en el Excel.`
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
      `Faltan columnas en el Excel: ${missing.join(", ")}`
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
      row.date
    );
}

// CALCULAR MÉTRICAS
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
        quantityByInvoice: new Map(),
        pricesByMonth: new Map()
      });
    }

    const product = productMap.get(row.product);

    product.invoices.add(row.invoiceId);
    product.dates.push(row.date);
    product.quantity += row.quantity;
    product.spend += row.itemValue;

    if (row.unitPrice > 0) {
      product.prices.push(row.unitPrice);
    }

    product.quantityByInvoice.set(
      row.invoiceId,
      (
        product.quantityByInvoice.get(
          row.invoiceId
        ) || 0
      ) + row.quantity
    );

    const month = monthKey(row.date);

    if (!product.pricesByMonth.has(month)) {
      product.pricesByMonth.set(month, []);
    }

    if (row.unitPrice > 0) {
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
        ...product.quantityByInvoice.values()
      ]);

      return {
        ...product,
        events,
        recurrence,
        medianPrice,
        averageDays: averageDays(product.dates),
        usual: recurrence >= 0.4,
        medianQuantity,
        basketCost:
          medianPrice * medianQuantity
      };
    })
    .sort((a, b) => b.spend - a.spend);

  const dates = [...invoices.values()]
    .map(invoice => invoice.date)
    .sort((a, b) => a - b);

  const totalSpend = [...invoices.values()]
    .reduce(
      (sum, invoice) =>
        sum + invoice.total,
      0
    );

  return {
    rows,
    invoiceCount,
    totalSpend,
    dates,

    monthly: [...monthlyMap.entries()]
      .sort(
        ([a], [b]) => a.localeCompare(b)
      ),

    categories: [...categoryMap.entries()]
      .sort(
        (a, b) => b[1] - a[1]
      )
  };
}

// GRÁFICOS
function destroyChart(name) {
  if (charts[name]) {
    charts[name].destroy();
  }
}

function renderCharts(data) {
  destroyChart("monthly");
  destroyChart("category");
  destroyChart("price");

  charts.monthly = new Chart(
    document.getElementById("monthlyChart"),
    {
      type: "line",
      data: {
        labels: data.monthly.map(
          ([month]) => month
        ),
        datasets: [{
          label: "Gasto mensual",
          data: data.monthly.map(
            ([, value]) => value
          ),
          borderColor: "#2563eb",
          backgroundColor:
            "rgba(37,99,235,.12)",
          fill: true,
          tension: 0.3
        }]
      },
      options: {
        responsive: true,
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
                currency.format(value)
            }
          }
        }
      }
    }
  );

  charts.category = new Chart(
    document.getElementById("categoryChart"),
    {
      type: "doughnut",
      data: {
        labels: data.categories.map(
          ([category]) => category
        ),
        datasets: [{
          data: data.categories.map(
            ([, value]) => value
          ),
          backgroundColor: [
            "#2563eb",
            "#22c1c3",
            "#f59e0b",
            "#8b5cf6",
            "#10b981",
            "#ef4444",
            "#64748b",
            "#ec4899",
            "#06b6d4",
            "#84cc16",
            "#f97316",
            "#6366f1"
          ]
        }]
      },
      options: {
        responsive: true,
        plugins: {
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

  const priceProducts = [...products]
    .filter(product => product.events >= 3)
    .sort((a, b) => b.events - a.events)
    .slice(0, 5);

  const months = [
    ...new Set(
      priceProducts.flatMap(
        product => [
          ...product.pricesByMonth.keys()
        ]
      )
    )
  ].sort();

  const colors = [
    "#2563eb",
    "#22c1c3",
    "#f59e0b",
    "#8b5cf6",
    "#ef4444"
  ];

  charts.price = new Chart(
    document.getElementById("priceChart"),
    {
      type: "line",
      data: {
        labels: months,
        datasets: priceProducts.map(
          (product, index) => ({
            label: product.product,

            data: months.map(month => {
              const values =
                product.pricesByMonth.get(month);

              return values?.length
                ? median(values)
                : null;
            }),

            borderColor: colors[index],
            tension: 0.25,
            spanGaps: true
          })
        )
      },
      options: {
        responsive: true,
        interaction: {
          mode: "nearest",
          intersect: false
        },
        scales: {
          y: {
            beginAtZero: true,
            ticks: {
              callback: value =>
                currency.format(value)
            }
          }
        }
      }
    }
  );
}

// TABLA DE PRODUCTOS
function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

function renderProductTable(filter = "") {
  const query = filter.trim().toLowerCase();

  const filtered = products.filter(
    product =>
      product.product
        .toLowerCase()
        .includes(query) ||
      product.category
        .toLowerCase()
        .includes(query)
  );

  document.getElementById(
    "productTable"
  ).innerHTML = filtered
    .map(product => `
      <tr>
        <td>${escapeHtml(product.product)}</td>
        <td>${escapeHtml(product.category)}</td>
        <td>${number.format(product.events)}</td>
        <td>${percent.format(product.recurrence)}</td>
        <td>${number.format(product.quantity)}</td>
        <td>${currency.format(product.spend)}</td>
        <td>${currency.format(product.medianPrice)}</td>
        <td>${
          product.averageDays === null
            ? "—"
            : number.format(product.averageDays)
        }</td>
        <td>${
          product.usual
            ? '<span class="badge">Sí</span>'
            : "No"
        }</td>
        <td>${number.format(product.medianQuantity)}</td>
        <td>${currency.format(product.basketCost)}</td>
      </tr>
    `)
    .join("");
}

// MOSTRAR DASHBOARD
function renderDashboard(data, workbookItem) {
  document.getElementById(
    "totalGasto"
  ).textContent =
    currency.format(data.totalSpend);

  document.getElementById(
    "totalFacturas"
  ).textContent =
    number.format(data.invoiceCount);

  document.getElementById(
    "ticketPromedio"
  ).textContent =
    currency.format(
      data.totalSpend / data.invoiceCount
    );

  document.getElementById(
    "totalProductos"
  ).textContent =
    number.format(products.length);

  const firstDate = data.dates[0];
  const lastDate =
    data.dates[data.dates.length - 1];

  document.getElementById(
    "periodo"
  ).textContent =
    `${dateFormat.format(firstDate)} — ` +
    `${dateFormat.format(lastDate)}`;

  document.getElementById(
    "sourceInfo"
  ).textContent =
    `${workbookItem.name} · ` +
    `${data.rows.length.toLocaleString("es-PA")} líneas · ` +
    "actualizado " +
    dateFormat.format(
      new Date(workbookItem.lastModifiedDateTime)
    );

  renderCharts(data);
  renderProductTable();

  dashboard.hidden = false;
}

// CARGAR DASHBOARD
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

    document.getElementById(
      "usuario"
    ).textContent =
      profile.displayName ||
      profile.userPrincipalName;

    const workbookItem =
      await findWorkbook(token);

    setStatus(
      "Descargando la última versión guardada del Excel…"
    );

    const buffer = await downloadWorkbook(
      token,
      workbookItem.id
    );

    setStatus(
      "Procesando facturas y calculando métricas…"
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
      `${number.format(data.invoiceCount)} facturas y ` +
      `${number.format(rows.length)} líneas analizadas.`,
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

// INICIAR SESIÓN
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

// CERRAR SESIÓN
async function logout() {
  await msalInstance.logoutPopup({
    account,
    postLogoutRedirectUri:
      CONFIG.redirectUri
  });

  account = null;
  dashboard.hidden = true;

  loginButton.hidden = false;
  logoutButton.hidden = true;
  refreshButton.hidden = true;

  document.getElementById(
    "usuario"
  ).textContent = "";

  setStatus(
    "Sesión cerrada. Conecta tu cuenta Microsoft para cargar los datos."
  );
}

// EVENTOS
loginButton.addEventListener(
  "click",
  login
);

refreshButton.addEventListener(
  "click",
  loadDashboard
);

logoutButton.addEventListener(
  "click",
  logout
);

productSearch.addEventListener(
  "input",
  event => {
    renderProductTable(event.target.value);
  }
);

// RECUPERAR SESIÓN
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
