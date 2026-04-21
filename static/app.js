const STORAGE_KEYS = {
  records: "eluruVillageIssues.records",
  reportDate: "eluruVillageIssues.reportDate",
};

const CHART_PALETTE = {
  caste: "#4F81BD",
  political: "#C0504D",
  communal: "#9BBB59",
  general: "#8064A2",
  categoryA: "#C0504D",
  categoryB: "#F2A65A",
  categoryC: "#9BBB59",
};

const state = {
  metadata: null,
  sampleRecords: [],
  records: [],
  reportDate: "",
  dashboard: null,
  editingId: null,
  searchQuery: "",
};

const els = {};

document.addEventListener("DOMContentLoaded", () => {
  cacheElements();
  bindEvents();
  initializeApp().catch((error) => {
    setStatus(error.message || "Unable to load the dashboard builder.", true);
  });
});

function cacheElements() {
  els.issueForm = document.getElementById("issue-form");
  els.recordId = document.getElementById("record-id");
  els.issueType = document.getElementById("issue-type");
  els.policeStation = document.getElementById("police-station");
  els.subDivision = document.getElementById("sub-division");
  els.category = document.getElementById("category");
  els.village = document.getElementById("village");
  els.partyCombination = document.getElementById("party-combination");
  els.partyCombinationList = document.getElementById("party-combination-list");
  els.issueSummary = document.getElementById("issue-summary");
  els.issueDetails = document.getElementById("issue-details");
  els.actionTaken = document.getElementById("action-taken");
  els.presentStatus = document.getElementById("present-status");
  els.remarks = document.getElementById("remarks");
  els.reportDate = document.getElementById("report-date");
  els.recordsTableBody = document.getElementById("records-table-body");
  els.recordSearch = document.getElementById("record-search");
  els.formModeBadge = document.getElementById("form-mode-badge");
  els.recordCountBadge = document.getElementById("record-count-badge");
  els.workbookMeta = document.getElementById("workbook-meta");
  els.statusMessage = document.getElementById("status-message");
  els.heroTitle = document.getElementById("hero-title");
  els.heroSubtitle = document.getElementById("hero-subtitle");
  els.heroDateChip = document.getElementById("hero-date-chip");
  els.heroTotalChip = document.getElementById("hero-total-chip");
  els.kpiTotalIssues = document.getElementById("kpi-total-issues");
  els.kpiPoliticalIssues = document.getElementById("kpi-political-issues");
  els.kpiGeneralIssues = document.getElementById("kpi-general-issues");
  els.kpiCasteCommunalIssues = document.getElementById("kpi-caste-communal-issues");
  els.kpiPoliticalNote = document.getElementById("kpi-political-note");
  els.kpiGeneralNote = document.getElementById("kpi-general-note");
  els.kpiCasteCommunalNote = document.getElementById("kpi-caste-communal-note");
  els.executiveSummaryBody = document.getElementById("executive-summary-body");
  els.executiveSummaryFoot = document.getElementById("executive-summary-foot");
  els.psWiseBody = document.getElementById("pswise-body");
  els.psWiseFoot = document.getElementById("pswise-foot");
  els.categoryBody = document.getElementById("category-body");
  els.categoryFoot = document.getElementById("category-foot");
  els.politicalBody = document.getElementById("political-body");
  els.actionBody = document.getElementById("action-body");
  els.issueDistributionChart = document.getElementById("issue-distribution-chart");
  els.subdivisionBreakdownChart = document.getElementById("subdivision-breakdown-chart");
  els.topStationsChart = document.getElementById("top-stations-chart");
  els.categoryDistributionChart = document.getElementById("category-distribution-chart");
  els.categorySplitChart = document.getElementById("category-split-chart");
  els.politicalChart = document.getElementById("political-chart");
  els.tabButtons = Array.from(document.querySelectorAll(".tab-button"));
  els.tabPanels = Array.from(document.querySelectorAll(".tab-panel"));
  els.loadSampleButton = document.getElementById("load-sample-button");
  els.clearDataButton = document.getElementById("clear-data-button");
  els.downloadButton = document.getElementById("download-button");
  els.resetFormButton = document.getElementById("reset-form-button");
}

function bindEvents() {
  els.issueForm.addEventListener("submit", onFormSubmit);
  els.resetFormButton.addEventListener("click", resetForm);
  els.policeStation.addEventListener("change", syncSubDivisionField);
  els.issueType.addEventListener("change", syncIssueTypeMode);
  els.reportDate.addEventListener("change", onReportDateChange);
  els.recordSearch.addEventListener("input", onSearchChange);
  els.recordsTableBody.addEventListener("click", onRecordTableAction);
  els.loadSampleButton.addEventListener("click", loadSampleRecordsIntoState);
  els.clearDataButton.addEventListener("click", clearAllRecords);
  els.downloadButton.addEventListener("click", downloadDashboardWorkbook);
  els.tabButtons.forEach((button) => {
    button.addEventListener("click", () => activateTab(button.dataset.tab));
  });
  window.addEventListener("resize", debounce(() => {
    if (state.dashboard) {
      renderCharts(state.dashboard);
    }
  }, 150));
}

async function initializeApp() {
  setStatus("Loading workbook data...");

  const metadata = await fetchJson("/api/bootstrap");
  state.metadata = metadata;
  state.sampleRecords = clone(metadata.sampleRecords || []);

  populateSelectOptions(metadata);
  renderWorkbookMeta(metadata);

  const storedRecords = readStorage(STORAGE_KEYS.records, []);
  const storedReportDate = readStorage(STORAGE_KEYS.reportDate, metadata.defaultReportDate);

  state.records = storedRecords.length > 0 ? storedRecords : clone(state.sampleRecords);
  state.reportDate = storedReportDate || metadata.defaultReportDate;
  els.reportDate.value = state.reportDate;

  syncIssueTypeMode();
  renderRecordsTable();
  await refreshDashboard();

  setStatus(
    `Loaded ${metadata.templateWorkbookName} as the Excel template and ${metadata.sourceWorkbookName} as the sample record source.`
  );
}

function populateSelectOptions(metadata) {
  const issueTypeOptions = metadata.issueTypes || [];
  issueTypeOptions.forEach((issueType) => {
    els.issueType.appendChild(createOption(issueType, issueType));
  });

  (metadata.policeStations || []).forEach((station) => {
    els.policeStation.appendChild(createOption(station.name, `${station.name} (${station.subDivision})`));
  });

  (metadata.categories || []).forEach((category) => {
    els.category.appendChild(createOption(category.value, `${category.label} (${category.alert})`));
  });

  (metadata.partyCombinations || []).forEach((combo) => {
    const option = document.createElement("option");
    option.value = combo;
    els.partyCombinationList.appendChild(option);
  });
}

function createOption(value, label) {
  const option = document.createElement("option");
  option.value = value;
  option.textContent = label;
  return option;
}

function renderWorkbookMeta(metadata) {
  els.workbookMeta.innerHTML = `
    <p><strong>Template:</strong> ${escapeHtml(metadata.templateWorkbookName || "")}</p>
    <p><strong>Sample Source:</strong> ${escapeHtml(metadata.sourceWorkbookName || "")}</p>
  `;
}

async function onFormSubmit(event) {
  event.preventDefault();

  const issueType = els.issueType.value;
  const policeStation = els.policeStation.value;
  const category = els.category.value;
  const village = els.village.value.trim();

  if (!issueType || !policeStation || !category || !village) {
    setStatus("Issue type, police station, category, and village are required.", true);
    return;
  }

  const record = {
    id: els.recordId.value || createRecordId(),
    issueType,
    policeStation,
    subDivision: els.subDivision.value.trim(),
    category,
    village,
    partyCombination: issueType === "Political Issues" ? els.partyCombination.value.trim() : "",
    issueSummary: els.issueSummary.value.trim(),
    issueDetails: els.issueDetails.value.trim(),
    actionTaken: els.actionTaken.value.trim(),
    presentStatus: els.presentStatus.value.trim(),
    remarks: els.remarks.value.trim(),
  };

  const existingIndex = state.records.findIndex((item) => item.id === record.id);
  if (existingIndex >= 0) {
    state.records[existingIndex] = record;
    setStatus("Record updated. Refreshing dashboard preview...");
  } else {
    state.records.unshift(record);
    setStatus("Record added. Refreshing dashboard preview...");
  }

  resetForm();
  renderRecordsTable();
  await refreshDashboard();
}

function resetForm() {
  els.issueForm.reset();
  els.recordId.value = "";
  state.editingId = null;
  els.formModeBadge.textContent = "New Record";
  els.formModeBadge.classList.remove("badge-warning");
  els.subDivision.value = "";
  syncIssueTypeMode();
}

function syncSubDivisionField() {
  const policeStation = els.policeStation.value;
  const station = (state.metadata?.policeStations || []).find((item) => item.name === policeStation);
  els.subDivision.value = station ? station.subDivision : "";
}

function syncIssueTypeMode() {
  const isPolitical = els.issueType.value === "Political Issues";
  els.partyCombination.disabled = !isPolitical;
  els.partyCombination.closest(".field").classList.toggle("field-disabled", !isPolitical);
  if (!isPolitical) {
    els.partyCombination.value = "";
  }
  syncSubDivisionField();
}

async function onReportDateChange() {
  state.reportDate = els.reportDate.value || state.metadata?.defaultReportDate || "";
  await refreshDashboard();
}

function onSearchChange() {
  state.searchQuery = els.recordSearch.value.trim().toLowerCase();
  renderRecordsTable();
}

function onRecordTableAction(event) {
  const actionButton = event.target.closest("button[data-action]");
  if (!actionButton) {
    return;
  }

  const recordId = actionButton.dataset.id;
  const action = actionButton.dataset.action;
  const record = state.records.find((item) => item.id === recordId);

  if (!record) {
    return;
  }

  if (action === "edit") {
    populateFormForEdit(record);
  }

  if (action === "delete") {
    state.records = state.records.filter((item) => item.id !== recordId);
    if (state.editingId === recordId) {
      resetForm();
    }
    renderRecordsTable();
    refreshDashboard().catch((error) => setStatus(error.message, true));
  }
}

function populateFormForEdit(record) {
  state.editingId = record.id;
  els.recordId.value = record.id;
  els.issueType.value = record.issueType || "";
  els.policeStation.value = record.policeStation || "";
  syncSubDivisionField();
  els.category.value = record.category || "";
  els.village.value = record.village || "";
  els.partyCombination.value = record.partyCombination || "";
  els.issueSummary.value = record.issueSummary || "";
  els.issueDetails.value = record.issueDetails || "";
  els.actionTaken.value = record.actionTaken || "";
  els.presentStatus.value = record.presentStatus || "";
  els.remarks.value = record.remarks || "";

  syncIssueTypeMode();
  els.formModeBadge.textContent = "Editing";
  els.formModeBadge.classList.add("badge-warning");
  els.village.focus();
}

async function loadSampleRecordsIntoState() {
  state.records = clone(state.sampleRecords);
  resetForm();
  renderRecordsTable();
  await refreshDashboard();
  setStatus("Sample workbook data loaded into the page.");
}

async function clearAllRecords() {
  state.records = [];
  resetForm();
  renderRecordsTable();
  await refreshDashboard();
  setStatus("All input records cleared.");
}

async function refreshDashboard() {
  writeStorage(STORAGE_KEYS.records, state.records);
  writeStorage(STORAGE_KEYS.reportDate, state.reportDate);
  renderRecordCount();

  const dashboard = await fetchJson("/api/dashboard", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      records: state.records,
      reportDate: state.reportDate,
    }),
  });

  state.dashboard = dashboard;
  renderDashboard(dashboard);
}

function renderDashboard(dashboard) {
  const titleParts = (dashboard.titles.executive || "").split("\n");
  els.heroTitle.textContent = titleParts[0] || "Dashboard Preview";
  els.heroSubtitle.textContent = titleParts[1] || "Live workbook-style dashboard preview";
  els.heroDateChip.textContent = `As on ${dashboard.reportDateLabel}`;
  els.heroTotalChip.textContent = `${dashboard.kpis.totalIssues} Total Issues`;

  els.kpiTotalIssues.textContent = dashboard.kpis.totalIssues;
  els.kpiPoliticalIssues.textContent = dashboard.kpis.politicalIssues;
  els.kpiGeneralIssues.textContent = dashboard.kpis.generalIssues;
  els.kpiCasteCommunalIssues.textContent = dashboard.kpis.casteCommunalIssues;
  els.kpiPoliticalNote.textContent = dashboard.kpis.politicalInsight;
  els.kpiGeneralNote.textContent = dashboard.kpis.generalInsight;
  els.kpiCasteCommunalNote.textContent = dashboard.kpis.casteCommunalInsight;

  renderExecutiveTable(dashboard);
  renderPsWiseTable(dashboard);
  renderCategoryTable(dashboard);
  renderPoliticalTable(dashboard);
  renderActionTable(dashboard);
  renderCharts(dashboard);
}

function renderExecutiveTable(dashboard) {
  const rows = dashboard.executive.subDivisionRows || [];
  els.executiveSummaryBody.innerHTML = rows.map((row) => `
    <tr>
      <td class="align-left strong">${escapeHtml(row.name)}</td>
      <td>${row.casteConflicts}</td>
      <td>${row.politicalIssues}</td>
      <td>${row.communalIssues}</td>
      <td>${row.generalIssues}</td>
      <td>${row.totalIssues}</td>
      <td>${escapeHtml(row.share)}</td>
    </tr>
  `).join("");

  const total = dashboard.executive.subDivisionTotal;
  els.executiveSummaryFoot.innerHTML = `
    <tr>
      <th class="align-left">GRAND TOTAL</th>
      <th>${total.casteConflicts}</th>
      <th>${total.politicalIssues}</th>
      <th>${total.communalIssues}</th>
      <th>${total.generalIssues}</th>
      <th>${total.totalIssues}</th>
      <th>${escapeHtml(total.share)}</th>
    </tr>
  `;
}

function renderPsWiseTable(dashboard) {
  const rows = dashboard.psWise.rows || [];
  els.psWiseBody.innerHTML = rows.map((row) => `
    <tr>
      <td>${row.serial}</td>
      <td class="align-left">${escapeHtml(row.policeStation)}</td>
      <td>${escapeHtml(row.subDivision)}</td>
      <td>${row.casteConflicts}</td>
      <td>${row.politicalIssues}</td>
      <td>${row.communalIssues}</td>
      <td>${row.generalIssues}</td>
      <td>${row.totalIssues}</td>
      <td><span class="severity severity-${row.severity.toLowerCase()}">${escapeHtml(row.severity)}</span></td>
    </tr>
  `).join("");

  const total = dashboard.psWise.total;
  els.psWiseFoot.innerHTML = `
    <tr>
      <th colspan="3" class="align-left">GRAND TOTAL</th>
      <th>${total.casteConflicts}</th>
      <th>${total.politicalIssues}</th>
      <th>${total.communalIssues}</th>
      <th>${total.generalIssues}</th>
      <th>${total.totalIssues}</th>
      <th></th>
    </tr>
  `;
}

function renderCategoryTable(dashboard) {
  const rows = dashboard.categoryAnalysis.rows || [];
  els.categoryBody.innerHTML = rows.map((row) => `
    <tr>
      <td class="align-left">${escapeHtml(row.issueType)}</td>
      <td>${row.categoryA}</td>
      <td>${row.categoryB}</td>
      <td>${row.categoryC}</td>
      <td>${row.total}</td>
      <td>${escapeHtml(row.percent)}</td>
    </tr>
  `).join("");

  const total = dashboard.categoryAnalysis.totals;
  els.categoryFoot.innerHTML = `
    <tr>
      <th class="align-left">TOTAL</th>
      <th>${total.categoryA}</th>
      <th>${total.categoryB}</th>
      <th>${total.categoryC}</th>
      <th>${total.total}</th>
      <th>${escapeHtml(total.percent)}</th>
    </tr>
  `;
}

function renderPoliticalTable(dashboard) {
  const rows = dashboard.politicalAnalysis.rows || [];
  els.politicalBody.innerHTML = rows.length > 0
    ? rows.map((row) => `
      <tr>
        <td class="align-left">${escapeHtml(row.partyCombination)}</td>
        <td>${escapeHtml(row.subDivision)}</td>
        <td>${row.policeStationsAffected}</td>
        <td>${row.issueCount}</td>
        <td>${escapeHtml(row.categories)}</td>
        <td class="align-left">${escapeHtml(row.presentStatus)}</td>
      </tr>
    `).join("")
    : `<tr><td colspan="6" class="empty-cell">No political issue combinations available.</td></tr>`;
}

function renderActionTable(dashboard) {
  const rows = dashboard.actionTracker.rows || [];
  els.actionBody.innerHTML = rows.length > 0
    ? rows.map((row) => `
      <tr>
        <td>${row.serial}</td>
        <td class="align-left">${escapeHtml(row.policeStation)}</td>
        <td class="align-left">${escapeHtml(row.village)}</td>
        <td class="align-left">${escapeHtml(row.issueSummary)}</td>
        <td>${escapeHtml(row.category)}</td>
        <td>${escapeHtml(row.alertLevel)}</td>
        <td class="align-left">${escapeHtml(row.actionTaken)}</td>
        <td class="align-left">${escapeHtml(row.presentStatus)}</td>
      </tr>
    `).join("")
    : `<tr><td colspan="8" class="empty-cell">No records available for the action tracker.</td></tr>`;
}

function renderCharts(dashboard) {
  renderDonutChart(els.issueDistributionChart, dashboard.executive.issueDistribution, {
    centerLabel: dashboard.kpis.totalIssues,
    centerSubLabel: "Total Issues",
  });

  renderGroupedBarChart(
    els.subdivisionBreakdownChart,
    dashboard.executive.subDivisionRows.map((row) => row.shortName),
    dashboard.executive.subDivisionChartSeries,
    {
      maxValueHint: Math.max(
        1,
        ...dashboard.executive.subDivisionChartSeries.flatMap((series) => series.values)
      ),
    }
  );

  renderHorizontalBarChart(els.topStationsChart, dashboard.psWise.topStations, {
    labelKey: "policeStation",
    valueKey: "totalIssues",
    color: CHART_PALETTE.caste,
    emptyLabel: "No station totals available.",
  });

  renderGroupedBarChart(
    els.categoryDistributionChart,
    dashboard.categoryAnalysis.rows.map((row) => row.issueType),
    dashboard.categoryAnalysis.categoryChartSeries,
    {
      maxValueHint: Math.max(
        1,
        ...dashboard.categoryAnalysis.categoryChartSeries.flatMap((series) => series.values)
      ),
    }
  );

  renderDonutChart(els.categorySplitChart, dashboard.categoryAnalysis.categorySplit, {
    centerLabel: dashboard.categoryAnalysis.totals.total,
    centerSubLabel: "Categories",
  });

  renderHorizontalBarChart(els.politicalChart, dashboard.politicalAnalysis.chartRows, {
    labelKey: "partyCombination",
    valueKey: "issueCount",
    color: CHART_PALETTE.political,
    emptyLabel: "No political combinations available.",
  });
}

function renderRecordsTable() {
  const rows = filterRecords(state.records, state.searchQuery);

  if (rows.length === 0) {
    els.recordsTableBody.innerHTML = `
      <tr>
        <td colspan="5" class="empty-cell">No records found.</td>
      </tr>
    `;
    return;
  }

  els.recordsTableBody.innerHTML = rows.map((record) => `
    <tr>
      <td>${escapeHtml(record.issueType || "")}</td>
      <td>${escapeHtml(record.policeStation || "")}</td>
      <td>${escapeHtml(record.village || "")}</td>
      <td>${escapeHtml(record.category || "")}</td>
      <td class="action-cell">
        <button type="button" class="mini-button" data-action="edit" data-id="${escapeHtml(record.id)}">Edit</button>
        <button type="button" class="mini-button mini-button-danger" data-action="delete" data-id="${escapeHtml(record.id)}">Delete</button>
      </td>
    </tr>
  `).join("");
}

function filterRecords(records, searchQuery) {
  if (!searchQuery) {
    return records;
  }

  return records.filter((record) => {
    const text = [
      record.issueType,
      record.policeStation,
      record.village,
      record.issueSummary,
      record.category,
    ].join(" ").toLowerCase();

    return text.includes(searchQuery);
  });
}

function renderRecordCount() {
  els.recordCountBadge.textContent = `${state.records.length} Records`;
}

async function downloadDashboardWorkbook() {
  try {
    setStatus("Generating Excel dashboard...");
    els.downloadButton.disabled = true;

    const response = await fetch("/api/export", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        records: state.records,
        reportDate: state.reportDate,
      }),
    });

    if (!response.ok) {
      throw new Error("Unable to generate the Excel dashboard.");
    }

    const blob = await response.blob();
    const downloadUrl = URL.createObjectURL(blob);
    const filename = getFilenameFromHeader(response.headers.get("Content-Disposition"))
      || `Eluru_Village_Issues_Dashboard_${state.reportDate || "export"}.xlsx`;

    const anchor = document.createElement("a");
    anchor.href = downloadUrl;
    anchor.download = filename;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(downloadUrl);

    setStatus("Excel dashboard downloaded successfully.");
  } catch (error) {
    setStatus(error.message || "Excel export failed.", true);
  } finally {
    els.downloadButton.disabled = false;
  }
}

function getFilenameFromHeader(contentDisposition) {
  if (!contentDisposition) {
    return "";
  }
  const match = contentDisposition.match(/filename="?([^"]+)"?/i);
  return match ? match[1] : "";
}

function activateTab(tabName) {
  els.tabButtons.forEach((button) => {
    button.classList.toggle("is-active", button.dataset.tab === tabName);
  });
  els.tabPanels.forEach((panel) => {
    panel.classList.toggle("is-active", panel.dataset.tabPanel === tabName);
  });
}

function renderDonutChart(container, items, options = {}) {
  const dataset = (items || []).filter((item) => Number(item.value) > 0);
  const total = dataset.reduce((sum, item) => sum + Number(item.value || 0), 0);

  if (total <= 0) {
    container.innerHTML = `<div class="chart-empty">${escapeHtml(options.emptyLabel || "No chart data available.")}</div>`;
    return;
  }

  const width = 360;
  const height = 250;
  const centerX = 128;
  const centerY = 126;
  const radius = 82;
  const strokeWidth = 34;

  let currentAngle = -90;
  const arcs = dataset.map((item) => {
    const sliceAngle = (Number(item.value) / total) * 360;
    const startAngle = currentAngle;
    const endAngle = currentAngle + sliceAngle;
    currentAngle = endAngle;
    return {
      ...item,
      path: describeArc(centerX, centerY, radius, startAngle, endAngle),
      percent: ((Number(item.value) / total) * 100).toFixed(1),
    };
  });

  container.innerHTML = `
    <div class="chart-flex chart-flex-donut">
      <svg viewBox="0 0 ${width} ${height}" class="chart-svg" role="img" aria-label="Donut chart">
        <circle cx="${centerX}" cy="${centerY}" r="${radius}" fill="none" stroke="#E6ECF5" stroke-width="${strokeWidth}"></circle>
        ${arcs.map((arc) => `
          <path d="${arc.path}" fill="none" stroke="${arc.color}" stroke-width="${strokeWidth}" stroke-linecap="butt"></path>
        `).join("")}
        <text x="${centerX}" y="${centerY - 4}" text-anchor="middle" class="chart-center-value">${escapeHtml(String(options.centerLabel || total))}</text>
        <text x="${centerX}" y="${centerY + 18}" text-anchor="middle" class="chart-center-label">${escapeHtml(options.centerSubLabel || "Total")}</text>
      </svg>
      <div class="chart-legend">
        ${arcs.map((arc) => `
          <div class="legend-row">
            <span class="legend-swatch" style="background:${arc.color}"></span>
            <span class="legend-label">${escapeHtml(arc.label)}</span>
            <span class="legend-value">${arc.value} (${arc.percent}%)</span>
          </div>
        `).join("")}
      </div>
    </div>
  `;
}

function renderGroupedBarChart(container, categories, seriesList, options = {}) {
  const categoriesSafe = categories || [];
  const seriesSafe = seriesList || [];
  const maxValue = Math.max(
    1,
    options.maxValueHint || 0,
    ...seriesSafe.flatMap((series) => (series.values || []).map((value) => Number(value || 0)))
  );

  if (categoriesSafe.length === 0 || seriesSafe.length === 0) {
    container.innerHTML = `<div class="chart-empty">No chart data available.</div>`;
    return;
  }

  const width = Math.max(540, categoriesSafe.length * 110);
  const height = 300;
  const margin = { top: 18, right: 18, bottom: 56, left: 54 };
  const plotWidth = width - margin.left - margin.right;
  const plotHeight = height - margin.top - margin.bottom;
  const tickCount = 5;
  const groupWidth = plotWidth / categoriesSafe.length;
  const barWidth = (groupWidth * 0.72) / Math.max(1, seriesSafe.length);
  const roundedMax = roundUpChartMax(maxValue);

  const yTicks = Array.from({ length: tickCount + 1 }, (_, index) => {
    const value = (roundedMax / tickCount) * index;
    const y = margin.top + plotHeight - (value / roundedMax) * plotHeight;
    return { value, y };
  });

  container.innerHTML = `
    <div class="chart-flex chart-flex-column">
      <svg viewBox="0 0 ${width} ${height}" class="chart-svg" role="img" aria-label="Grouped bar chart">
        <line x1="${margin.left}" y1="${margin.top}" x2="${margin.left}" y2="${margin.top + plotHeight}" class="axis-line"></line>
        <line x1="${margin.left}" y1="${margin.top + plotHeight}" x2="${margin.left + plotWidth}" y2="${margin.top + plotHeight}" class="axis-line"></line>
        ${yTicks.map((tick) => `
          <g>
            <line x1="${margin.left}" y1="${tick.y}" x2="${margin.left + plotWidth}" y2="${tick.y}" class="grid-line"></line>
            <text x="${margin.left - 10}" y="${tick.y + 4}" text-anchor="end" class="axis-text">${formatTick(tick.value)}</text>
          </g>
        `).join("")}
        ${categoriesSafe.map((category, categoryIndex) => {
          const groupStart = margin.left + categoryIndex * groupWidth + (groupWidth * 0.14);
          const labelX = margin.left + categoryIndex * groupWidth + (groupWidth / 2);
          const bars = seriesSafe.map((series, seriesIndex) => {
            const value = Number(series.values?.[categoryIndex] || 0);
            const barHeight = (value / roundedMax) * plotHeight;
            const x = groupStart + seriesIndex * barWidth;
            const y = margin.top + plotHeight - barHeight;
            return `
              <rect x="${x}" y="${y}" width="${Math.max(barWidth - 4, 8)}" height="${barHeight}" rx="4" fill="${series.color}"></rect>
              <text x="${x + (Math.max(barWidth - 4, 8) / 2)}" y="${Math.max(y - 6, margin.top + 12)}" text-anchor="middle" class="bar-value">${value}</text>
            `;
          }).join("");
          return `
            <g>
              ${bars}
              <text x="${labelX}" y="${height - 20}" text-anchor="middle" class="axis-text axis-text-wrap">${escapeHtml(shortLabel(category, 16))}</text>
            </g>
          `;
        }).join("")}
      </svg>
      <div class="chart-legend">
        ${seriesSafe.map((series) => `
          <div class="legend-row">
            <span class="legend-swatch" style="background:${series.color}"></span>
            <span class="legend-label">${escapeHtml(series.label)}</span>
          </div>
        `).join("")}
      </div>
    </div>
  `;
}

function renderHorizontalBarChart(container, rows, options = {}) {
  const dataset = (rows || []).filter((item) => Number(item[options.valueKey] || 0) > 0);
  if (dataset.length === 0) {
    container.innerHTML = `<div class="chart-empty">${escapeHtml(options.emptyLabel || "No chart data available.")}</div>`;
    return;
  }

  const width = 560;
  const rowHeight = 34;
  const height = Math.max(220, 52 + dataset.length * rowHeight);
  const margin = { top: 18, right: 24, bottom: 24, left: 170 };
  const plotWidth = width - margin.left - margin.right;
  const maxValue = roundUpChartMax(Math.max(1, ...dataset.map((item) => Number(item[options.valueKey] || 0))));

  container.innerHTML = `
    <svg viewBox="0 0 ${width} ${height}" class="chart-svg" role="img" aria-label="Horizontal bar chart">
      <line x1="${margin.left}" y1="${margin.top}" x2="${margin.left}" y2="${height - margin.bottom}" class="axis-line"></line>
      <line x1="${margin.left}" y1="${height - margin.bottom}" x2="${width - margin.right}" y2="${height - margin.bottom}" class="axis-line"></line>
      ${Array.from({ length: 5 }, (_, index) => {
        const value = (maxValue / 4) * index;
        const x = margin.left + (value / maxValue) * plotWidth;
        return `
          <g>
            <line x1="${x}" y1="${margin.top}" x2="${x}" y2="${height - margin.bottom}" class="grid-line vertical-grid"></line>
            <text x="${x}" y="${height - 6}" text-anchor="middle" class="axis-text">${formatTick(value)}</text>
          </g>
        `;
      }).join("")}
      ${dataset.map((item, index) => {
        const value = Number(item[options.valueKey] || 0);
        const barWidth = (value / maxValue) * plotWidth;
        const y = margin.top + index * rowHeight + 6;
        return `
          <g>
            <text x="${margin.left - 10}" y="${y + 14}" text-anchor="end" class="axis-text axis-label-left">${escapeHtml(shortLabel(item[options.labelKey], 26))}</text>
            <rect x="${margin.left}" y="${y}" width="${barWidth}" height="18" rx="6" fill="${options.color || CHART_PALETTE.caste}"></rect>
            <text x="${margin.left + barWidth + 8}" y="${y + 14}" class="bar-value">${value}</text>
          </g>
        `;
      }).join("")}
    </svg>
  `;
}

function describeArc(centerX, centerY, radius, startAngle, endAngle) {
  const start = polarToCartesian(centerX, centerY, radius, endAngle);
  const end = polarToCartesian(centerX, centerY, radius, startAngle);
  const largeArcFlag = endAngle - startAngle <= 180 ? "0" : "1";
  return [
    "M", start.x, start.y,
    "A", radius, radius, 0, largeArcFlag, 0, end.x, end.y,
  ].join(" ");
}

function polarToCartesian(centerX, centerY, radius, angleInDegrees) {
  const angleInRadians = ((angleInDegrees - 90) * Math.PI) / 180.0;
  return {
    x: centerX + (radius * Math.cos(angleInRadians)),
    y: centerY + (radius * Math.sin(angleInRadians)),
  };
}

function shortLabel(value, limit) {
  const text = String(value || "");
  return text.length <= limit ? text : `${text.slice(0, limit - 3)}...`;
}

function formatTick(value) {
  return Number.isInteger(value) ? value : value.toFixed(1);
}

function roundUpChartMax(value) {
  if (value <= 5) {
    return 5;
  }
  if (value <= 10) {
    return 10;
  }
  const magnitude = 10 ** Math.floor(Math.log10(value));
  return Math.ceil(value / magnitude) * magnitude;
}

async function fetchJson(url, options = {}) {
  const response = await fetch(url, options);
  if (!response.ok) {
    throw new Error(`Request failed with status ${response.status}.`);
  }
  return response.json();
}

function setStatus(message, isError = false) {
  els.statusMessage.textContent = message;
  els.statusMessage.classList.toggle("is-error", isError);
}

function readStorage(key, fallbackValue) {
  try {
    const rawValue = localStorage.getItem(key);
    return rawValue ? JSON.parse(rawValue) : fallbackValue;
  } catch {
    return fallbackValue;
  }
}

function writeStorage(key, value) {
  try {
    localStorage.setItem(key, JSON.stringify(value));
  } catch {
    // Ignore storage quota issues and keep the page working in-memory.
  }
}

function createRecordId() {
  if (window.crypto?.randomUUID) {
    return window.crypto.randomUUID();
  }
  return `record-${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function debounce(fn, wait) {
  let timeoutId = null;
  return (...args) => {
    window.clearTimeout(timeoutId);
    timeoutId = window.setTimeout(() => fn(...args), wait);
  };
}

function escapeHtml(value) {
  return String(value ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function clone(value) {
  return JSON.parse(JSON.stringify(value));
}
