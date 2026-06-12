const views = {
  dashboard: {
    title: "Dashboard Entry",
    text: "把常用功能整理成入口首頁，使用者一進來先看到最常進入的頁面、目前狀態與可執行動作。",
    columns: [
      {
        label: "Quick Access",
        items: ["Shipment Overview", "Weekly Summary", "Alert Queue"]
      },
      {
        label: "Today",
        items: ["2 items need review", "1 high priority", "3 shortcuts pinned"]
      }
    ]
  },
  tasks: {
    title: "Task Snapshot",
    text: "入口頁不只是放連結，也能先看到今天最需要注意的項目，幫助使用者少切幾層畫面。",
    columns: [
      {
        label: "Priority",
        items: ["Review pending files", "Confirm weekly output", "Check exception list"]
      },
      {
        label: "Signals",
        items: ["1 overdue warning", "2 fresh updates", "Next sync at 15:30"]
      }
    ]
  }
};

const shipmentSample = {
  fileName: "ASUS & HP AIO shipment summary_2026_April_20260504_Ver.01.xlsx",
  monthLabel: "2026-04",
  overallLatest: 291563,
  overallChange: -158554,
  overallChangePct: -35.2,
  dominantBrand: "ASUS",
  brandShare: [
    { brand: "HP-AIO", value: 53385, sharePct: 18.3 },
    { brand: "ASUS", value: 238178, sharePct: 81.7 }
  ],
  emailLines: [
    "Shipment summary for 2026-04",
    "Overall shipment: 291,563 units (-158,554 units, -35.2% MoM)",
    "HP-AIO: 53,385 units in 2026-04 (-27,008 units, -33.6% MoM)",
    "Top models: BPS | Jumanji L10 14,555, BPS | Boggle5 L10 12,336, BPS | Kazaam 24 L6 11,832"
  ],
  brands: [
    {
      brand: "HP-AIO",
      latestMonthLabel: "2026-04",
      latestTotal: 53385,
      momChange: -27008,
      momChangePct: -33.6,
      trendLabels: ["2025-05", "2025-06", "2025-07", "2025-08", "2025-09", "2025-10", "2025-11", "2025-12", "2026-01", "2026-02", "2026-03", "2026-04"],
      trendValues: [43199, 74840, 55932, 52333, 53720, 35891, 56896, 71098, 49249, 38989, 80393, 53385],
      topModels: [
        { model: "BPS | Jumanji L10", latestMonth: 14555, delta: -20929, grandTotal: 153770 },
        { model: "BPS | Boggle5 L10", latestMonth: 12336, delta: 4543, grandTotal: 347243 },
        { model: "BPS | Kazaam 24 L6", latestMonth: 11832, delta: 3960, grandTotal: 48183 },
        { model: "BPS | Kazaam 27 L6", latestMonth: 4536, delta: 1944, grandTotal: 23472 },
        { model: "BPS | Kazaam 24 L10", latestMonth: 3747, delta: -5147, grandTotal: 28287 }
      ]
    },
    {
      brand: "ASUS",
      latestMonthLabel: "2026-04",
      latestTotal: 238178,
      momChange: -131546,
      momChangePct: -35.6,
      trendLabels: ["2025-05", "2025-06", "2025-07", "2025-08", "2025-09", "2025-10", "2025-11", "2025-12", "2026-01", "2026-02", "2026-03", "2026-04"],
      trendValues: [421630, 511154, 324659, 388741, 443716, 259865, 182831, 223372, 219787, 168702, 369724, 238178],
      topModels: [
        { model: "FX608R", latestMonth: 46653, delta: 26335, grandTotal: 66971 },
        { model: "FX608", latestMonth: 34972, delta: -35452, grandTotal: 1102611 },
        { model: "FX607VJ", latestMonth: 9886, delta: -8845, grandTotal: 77778 },
        { model: "S5406SA", latestMonth: 9337, delta: 4110, grandTotal: 115044 },
        { model: "UX5406", latestMonth: 7629, delta: -6798, grandTotal: 151429 }
      ]
    }
  ],
  highlights: [
    "Jan~Mar 2026 vs Jan~Mar 2025: +137,348 units, +17.2%",
    "ASUS Mar 2026 vs Mar 2025: +1,441 units",
    "ASUS Mar 2026 territory mix: CHINA 20.74% / TAIWAN 1.48% / APAC 11.98% / EEMEA 20.33%"
  ],
  territoryShare: [
    { label: "WEMEA", ratioPct: 24.62, total: 91041 },
    { label: "CHINA", ratioPct: 20.74, total: 76669 },
    { label: "EEMEA", ratioPct: 20.33, total: 75159 },
    { label: "APAC", ratioPct: 11.98, total: 44301 },
    { label: "LATAM", ratioPct: 8.08, total: 29862 },
    { label: "TAIWAN", ratioPct: 1.48, total: 5457 }
  ],
  customerShipment: [
    { label: "HP_AIO (BPS)", values: [36901, 74581, 51757], total: 163239 },
    { label: "HP_AIO (CPS)", values: [4697, 10545, 1871], total: 17113 },
    { label: "ASUS", values: [168702, 369724, 238178], total: 776604 },
    { label: "Total", values: [210300, 454850, 291806], total: 956956 }
  ],
  periodLabels: ["2026-02", "2026-03", "2026-04"]
};

const allowedEmployeeIds = new Set(["IEC950458"]);
const loginStorageKey = "sites_prototype_employee_id";

const detailTitle = document.querySelector("#detail-title");
const detailText = document.querySelector("#detail-text");
const previewContent = document.querySelector("#preview-content");
const loginShell = document.querySelector("#login-shell");
const loginForm = document.querySelector("#login-form");
const loginError = document.querySelector("#login-error");
const employeeIdInput = document.querySelector("#employee-id");
const activeEmployeeId = document.querySelector("#active-employee-id");
const logoutButton = document.querySelector("#logout-btn");

function renderPreview(viewKey) {
  const view = views[viewKey];
  if (!view) return;

  detailTitle.textContent = view.title;
  detailText.textContent = view.text;
  previewContent.innerHTML = "";

  view.columns.forEach((column, index) => {
    const wrapper = document.createElement("div");
    wrapper.className = `preview-column${index === 1 ? " accent" : ""}`;

    const label = document.createElement("p");
    label.className = "preview-label";
    label.textContent = column.label;
    wrapper.appendChild(label);

    column.items.forEach((item) => {
      const line = document.createElement(index === 0 ? "strong" : "span");
      line.textContent = item;
      wrapper.appendChild(line);
    });

    previewContent.appendChild(wrapper);
  });
}

function formatNumber(value) {
  return new Intl.NumberFormat("en-US").format(value);
}

function formatSigned(value) {
  const sign = value > 0 ? "+" : "";
  return `${sign}${formatNumber(value)}`;
}

function formatPercent(value) {
  if (value === null || value === undefined) return "N/A";
  const sign = value > 0 ? "+" : "";
  return `${sign}${value.toFixed(1)}%`;
}

function buildLineChartMarkup(values, labels, options = {}) {
  const width = options.width ?? 520;
  const height = options.height ?? 180;
  const paddingX = 18;
  const paddingY = 18;
  const colorClass = options.colorClass ?? "";
  const minValue = Math.min(...values);
  const maxValue = Math.max(...values);
  const range = Math.max(maxValue - minValue, 1);
  const stepX = values.length > 1 ? (width - paddingX * 2) / (values.length - 1) : 0;
  const points = values.map((value, index) => {
    const x = paddingX + stepX * index;
    const normalized = (value - minValue) / range;
    const y = height - paddingY - normalized * (height - paddingY * 2);
    return { x, y, value, label: labels[index], index };
  });
  const path = points.map((point, index) => `${index === 0 ? "M" : "L"} ${point.x.toFixed(1)} ${point.y.toFixed(1)}`).join(" ");
  const latestIndex = values.length - 1;
  const peakIndex = values.indexOf(maxValue);
  const lowIndex = values.indexOf(minValue);
  const gridYs = [paddingY, height / 2, height - paddingY];

  return `
    <div class="line-chart-card">
      <svg class="line-chart-svg" viewBox="0 0 ${width} ${height}" preserveAspectRatio="none" aria-hidden="true">
        ${gridYs.map((y) => `<line class="line-chart-grid" x1="${paddingX}" y1="${y}" x2="${width - paddingX}" y2="${y}"></line>`).join("")}
        <path class="line-chart-path ${colorClass}" d="${path}"></path>
        ${points.map((point, index) => `
          <circle
            class="line-chart-dot ${colorClass}${index === latestIndex ? " is-latest" : ""}"
            cx="${point.x.toFixed(1)}"
            cy="${point.y.toFixed(1)}"
            r="${index === latestIndex ? 5 : 4}"
          ></circle>
        `).join("")}
      </svg>
      <div class="line-chart-axis">
        ${labels.map((label) => `<span>${label}</span>`).join("")}
      </div>
      <div class="line-chart-summary">
        <div>
          <strong>Peak ${labels[peakIndex]}</strong>
          <span>${formatNumber(maxValue)}</span>
        </div>
        <div>
          <strong>Low ${labels[lowIndex]}</strong>
          <span>${formatNumber(minValue)}</span>
        </div>
        <div>
          <strong>Latest ${labels[latestIndex]}</strong>
          <span>${formatNumber(values[latestIndex])}</span>
        </div>
      </div>
    </div>
  `;
}

function renderShipmentPilot() {
  const sourceFile = document.querySelector("#shipment-source-file");
  const period = document.querySelector("#shipment-period");
  const overallTrend = document.querySelector("#shipment-overall-trend");
  const dominantBrand = document.querySelector("#shipment-dominant-brand");
  const overallTotal = document.querySelector("#shipment-overall-total");
  const summaryMetrics = document.querySelector("#summary-metrics");
  const mailSummaryLines = document.querySelector("#mail-summary-lines");
  const brandSplitList = document.querySelector("#brand-split-list");
  const brandWatchGrid = document.querySelector("#brand-watch-grid");
  const insightList = document.querySelector("#insight-list");
  const customerShipmentBody = document.querySelector("#customer-shipment-body");
  const overallChartCaption = document.querySelector("#overall-chart-caption");
  const overallBarChart = document.querySelector("#overall-bar-chart");
  const brandCompareStack = document.querySelector("#brand-compare-stack");
  const trendDeckGrid = document.querySelector("#trend-deck-grid");
  const territoryHero = document.querySelector("#territory-hero");
  const territoryStack = document.querySelector("#territory-stack");
  const customerVisualStack = document.querySelector("#customer-visual-stack");

  if (!sourceFile) return;

  sourceFile.textContent = shipmentSample.fileName;
  period.textContent = shipmentSample.monthLabel;
  overallTrend.textContent = `${formatSigned(shipmentSample.overallChange)} units / ${formatPercent(shipmentSample.overallChangePct)} MoM`;
  dominantBrand.textContent = shipmentSample.dominantBrand;
  overallTotal.textContent = `${formatNumber(shipmentSample.overallLatest)} units overall`;

  const metricCards = [
    {
      label: "Overall shipment",
      value: `${formatNumber(shipmentSample.overallLatest)} units`,
      note: `${formatSigned(shipmentSample.overallChange)} units vs previous month`
    },
    {
      label: "MoM change",
      value: formatPercent(shipmentSample.overallChangePct),
      note: "Directly pulled from the sample workbook analysis"
    },
    {
      label: "Dominant brand",
      value: shipmentSample.dominantBrand,
      note: "Largest share in the current period"
    },
    {
      label: "Connected mode",
      value: "Prototype + real summary",
      note: "Existing local UI stays untouched"
    }
  ];

  summaryMetrics.innerHTML = metricCards.map((card) => `
    <article class="summary-card">
      <span>${card.label}</span>
      <strong>${card.value}</strong>
      <p>${card.note}</p>
    </article>
  `).join("");

  mailSummaryLines.innerHTML = shipmentSample.emailLines.map((line) => `<p>${line}</p>`).join("");

  brandSplitList.innerHTML = shipmentSample.brandShare.map((item) => `
    <article class="brand-share-card" style="--share-size: ${item.sharePct}%;">
      <span>${item.brand}</span>
      <strong>${item.sharePct.toFixed(1)}%</strong>
      <p>${formatNumber(item.value)} units in ${shipmentSample.monthLabel}</p>
    </article>
  `).join("");

  brandWatchGrid.innerHTML = shipmentSample.brands.map((brand) => `
    <article class="brand-watch-card">
      <span>${brand.brand}</span>
      <strong>${formatNumber(brand.latestTotal)} units</strong>
      <p>${brand.latestMonthLabel} / ${formatSigned(brand.momChange)} / ${formatPercent(brand.momChangePct)} MoM</p>
      <ul>
        ${brand.topModels.slice(0, 3).map((item) => `<li>${item.model} ${formatNumber(item.latestMonth)}</li>`).join("")}
      </ul>
    </article>
  `).join("");

  insightList.innerHTML = shipmentSample.highlights.map((item, index) => `
    <article class="insight-item">
      <strong>Insight 0${index + 1}</strong>
      <p>${item}</p>
    </article>
  `).join("");

  customerShipmentBody.innerHTML = shipmentSample.customerShipment.map((row) => `
    <tr>
      <td>${row.label}</td>
      <td>${formatNumber(row.values[0])}</td>
      <td>${formatNumber(row.values[1])}</td>
      <td>${formatNumber(row.values[2])}</td>
      <td>${formatNumber(row.total)}</td>
    </tr>
  `).join("");

  const totalSeries = shipmentSample.customerShipment.find((row) => row.label === "Total")?.values ?? [];
  overallChartCaption.textContent = `${shipmentSample.periodLabels[0]} to ${shipmentSample.periodLabels[2]}`;
  overallBarChart.innerHTML = buildLineChartMarkup(totalSeries, shipmentSample.periodLabels.map((label) => label.slice(5)), {
    colorClass: "teal",
    height: 180
  });

  const hpSeries = shipmentSample.customerShipment
    .filter((row) => row.label.startsWith("HP_AIO"))
    .reduce((accumulator, row) => accumulator.map((value, index) => value + row.values[index]), [0, 0, 0]);
  const asusSeries = shipmentSample.customerShipment.find((row) => row.label === "ASUS")?.values ?? [0, 0, 0];

  const compareRows = [
    {
      label: "HP-AIO current weight",
      latest: hpSeries[2],
      peak: Math.max(...hpSeries, 1),
      caption: `${formatSigned(hpSeries[2] - hpSeries[1])} vs previous month`,
      theme: "coral"
    },
    {
      label: "ASUS current weight",
      latest: asusSeries[2],
      peak: Math.max(...asusSeries, 1),
      caption: `${formatSigned(asusSeries[2] - asusSeries[1])} vs previous month`,
      theme: "teal"
    }
  ];

  brandCompareStack.innerHTML = compareRows.map((row) => `
    <article class="compare-card">
      <strong>${row.label}</strong>
      <p>${row.caption}</p>
      <div class="compare-bar-shell">
        <div class="compare-bar-labels">
          <span>${formatNumber(row.latest)} units</span>
          <span>peak ${formatNumber(row.peak)}</span>
        </div>
        <div class="compare-bar-track">
          <div class="compare-bar-fill${row.theme === "coral" ? " coral" : ""}" style="--fill-width: ${(row.latest / row.peak) * 100}%"></div>
        </div>
      </div>
    </article>
  `).join("");

  trendDeckGrid.innerHTML = shipmentSample.brands.map((brand) => {
    const trendMax = Math.max(...brand.trendValues, 1);
    const trendMin = Math.min(...brand.trendValues);
    const trendPeakIndex = brand.trendValues.indexOf(trendMax);
    const trendLowIndex = brand.trendValues.indexOf(trendMin);
    const toneClass = brand.brand === "ASUS" ? "teal" : "";
    const pillClass = brand.momChange < 0 ? "trend-metric-pill down" : "trend-metric-pill";
    return `
      <article class="trend-brand-card">
        <div class="trend-brand-head">
          <div>
            <strong>${brand.brand}</strong>
            <p>${brand.latestMonthLabel} latest shipment ${formatNumber(brand.latestTotal)} units</p>
            <div class="trend-meta">
              <span class="trend-meta-chip">Peak ${brand.trendLabels[trendPeakIndex]} / ${formatNumber(trendMax)}</span>
              <span class="trend-meta-chip">Low ${brand.trendLabels[trendLowIndex]} / ${formatNumber(trendMin)}</span>
            </div>
          </div>
          <span class="${pillClass}">${formatSigned(brand.momChange)} / ${formatPercent(brand.momChangePct)} MoM</span>
        </div>
        <div class="trend-bars-mini">${buildLineChartMarkup(brand.trendValues, brand.trendLabels.map((label) => label.slice(5)), {
          colorClass: toneClass,
          width: 560,
          height: 190
        })}</div>
        <table class="top-model-table">
          <thead>
            <tr>
              <th>Top model</th>
              <th>Latest</th>
              <th>MoM</th>
              <th>Total</th>
            </tr>
          </thead>
          <tbody>
            ${brand.topModels.map((item) => `
              <tr>
                <td>${item.model}</td>
                <td>${formatNumber(item.latestMonth)}</td>
                <td class="${item.delta >= 0 ? "trend-up" : "trend-down"}">${formatSigned(item.delta)}</td>
                <td>${formatNumber(item.grandTotal)}</td>
              </tr>
            `).join("")}
          </tbody>
        </table>
      </article>
    `;
  }).join("");

  const sortedTerritories = [...shipmentSample.territoryShare].sort((a, b) => b.ratioPct - a.ratioPct);
  territoryHero.innerHTML = [
    {
      label: "Top region",
      value: sortedTerritories[0]?.label ?? "N/A",
      note: `${sortedTerritories[0]?.ratioPct?.toFixed(2) ?? "0.00"}% share`
    },
    {
      label: "Second region",
      value: sortedTerritories[1]?.label ?? "N/A",
      note: `${sortedTerritories[1]?.ratioPct?.toFixed(2) ?? "0.00"}% share`
    },
    {
      label: "Coverage shown",
      value: `${sortedTerritories.length} regions`,
      note: "Based on current summary extraction"
    }
  ].map((item) => `
    <article>
      <span>${item.label}</span>
      <strong>${item.value}</strong>
      <p>${item.note}</p>
    </article>
  `).join("");

  territoryStack.innerHTML = sortedTerritories.map((item) => `
    <article class="territory-item">
      <div class="territory-row-head">
        <strong>${item.label}</strong>
        <span>${item.ratioPct.toFixed(2)}%</span>
      </div>
      <div class="territory-track">
        <div class="territory-fill" style="--territory-width: ${item.ratioPct}%;"></div>
      </div>
      <p>3-month total ${formatNumber(item.total)} units</p>
    </article>
  `).join("");

  const customerRows = shipmentSample.customerShipment.filter((row) => row.label !== "Total");
  customerVisualStack.innerHTML = customerRows.map((row, index) => {
    const change = row.values[2] - row.values[1];
    const toneClass = index === customerRows.length - 1 ? "teal" : "";
    return `
      <article class="customer-visual-card">
        <div class="customer-visual-head">
          <strong>${row.label}</strong>
          <span class="${change >= 0 ? "trend-up" : "trend-down"}">${formatSigned(change)} vs M-1</span>
        </div>
        <p>${formatNumber(row.total)} units across 3 months</p>
        <div class="spark-line">${buildLineChartMarkup(row.values, shipmentSample.periodLabels.map((label) => label.slice(5)), {
          colorClass: toneClass,
          width: 420,
          height: 150
        })}</div>
        <div class="spark-caption">${formatNumber(row.values[0])} -> ${formatNumber(row.values[1])} -> ${formatNumber(row.values[2])}</div>
      </article>
    `;
  }).join("");
}

function normalizeEmployeeId(value) {
  return value.trim().toUpperCase();
}

function openLoginGate() {
  if (employeeIdInput) {
    employeeIdInput.value = "";
  }
  if (loginError) {
    loginError.hidden = true;
  }
  if (loginShell) {
    loginShell.hidden = false;
  }
  document.body.classList.add("is-locked");
}

function grantAccess(employeeId) {
  const normalized = normalizeEmployeeId(employeeId);
  localStorage.setItem(loginStorageKey, normalized);
  if (activeEmployeeId) {
    activeEmployeeId.textContent = normalized;
  }
  if (loginShell) {
    loginShell.hidden = true;
  }
  if (loginError) {
    loginError.hidden = true;
  }
  document.body.classList.remove("is-locked");
}

function initializeAccessGate() {
  const rememberedId = localStorage.getItem(loginStorageKey);
  if (rememberedId && allowedEmployeeIds.has(rememberedId)) {
    grantAccess(rememberedId);
  } else {
    openLoginGate();
  }

  loginForm?.addEventListener("submit", (event) => {
    event.preventDefault();
    const submitted = normalizeEmployeeId(employeeIdInput?.value ?? "");
    if (allowedEmployeeIds.has(submitted)) {
      grantAccess(submitted);
      return;
    }
    if (loginError) {
      loginError.hidden = false;
    }
  });

  logoutButton?.addEventListener("click", () => {
    localStorage.removeItem(loginStorageKey);
    openLoginGate();
    employeeIdInput?.focus();
  });
}

document.querySelectorAll("[data-scroll-target]").forEach((button) => {
  button.addEventListener("click", () => {
    const target = document.querySelector(button.dataset.scrollTarget);
    target?.scrollIntoView({ behavior: "smooth", block: "start" });
  });
});

document.querySelectorAll("[data-highlight]").forEach((button) => {
  button.addEventListener("click", () => {
    renderPreview(button.dataset.highlight);
    document.querySelector("#detail-panel")?.scrollIntoView({ behavior: "smooth", block: "center" });
  });
});

renderPreview("dashboard");
renderShipmentPilot();
initializeAccessGate();
