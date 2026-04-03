const state = {
  rows: [],
  sampleRows: [],
  columns: [],
  numericColumns: [],
  dateColumns: [],
  categoryColumns: [],
  idColumns: [],
  otherColumns: [],
  activeColumnFilter: "all",
  columnSearchTerm: "",
  fileName: "",
  isProcessing: false,
};

const ANALYSIS_SAMPLE_LIMIT = 1000;
const PREVIEW_ROW_LIMIT = 8;
const DELIMITER_CANDIDATES = [",", "\t", ";", "|"];
const MAX_GROUPS = 20;

const fileInput = document.getElementById("file-input");
const fileStatus = document.getElementById("file-status");
const chartOptions = document.getElementById("chart-options");
const xAxisSelect = document.getElementById("x-axis");
const yAxisSelect = document.getElementById("y-axis");
const groupAxisSelect = document.getElementById("group-axis");
const generateBtn = document.getElementById("generate-btn");
const smartDashboardBtn = document.getElementById("smart-dashboard-btn");
const clearPlotsBtn = document.getElementById("clear-plots-btn");
const chartClearInlineBtn = document.getElementById("chart-clear-inline-btn");
const clearBtn = document.getElementById("clear-btn");
const rowCount = document.getElementById("row-count");
const columnCount = document.getElementById("column-count");
const numericCount = document.getElementById("numeric-count");
const columnMeta = document.getElementById("column-meta");
const previewHead = document.querySelector("#preview-table thead");
const previewBody = document.querySelector("#preview-table tbody");
const chartGrid = document.getElementById("chart-grid");
const chartMessage = document.getElementById("chart-message");
const chartPanel = document.querySelector(".chart-panel");
const uploadZone = document.querySelector(".upload-zone");
const recommendations = document.getElementById("recommendations");
const chartModal = document.getElementById("chart-modal");
const chartModalSurface = document.getElementById("chart-modal-surface");
const chartModalClose = document.getElementById("chart-modal-close");
const chartModalTitle = document.getElementById("chart-modal-title");
const columnSearch = document.getElementById("column-search");
const columnFilterTabs = document.getElementById("column-filter-tabs");
const selectorGroups = document.getElementById("selector-groups");
const processingHint = document.getElementById("processing-hint");
let chartCounter = 0;

// State helpers
function setProcessing(isProcessing, message) {
  state.isProcessing = isProcessing;
  fileStatus.textContent = message;
  fileStatus.classList.toggle("is-loading", isProcessing);
  processingHint.hidden = !isProcessing;
}

fileInput.addEventListener("change", (event) => {
  const [file] = event.target.files;
  if (file) {
    loadFile(file);
  }
});

["dragenter", "dragover"].forEach((eventName) => {
  uploadZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    uploadZone.classList.add("dragover");
  });
});

["dragleave", "drop"].forEach((eventName) => {
  uploadZone.addEventListener(eventName, (event) => {
    event.preventDefault();
    uploadZone.classList.remove("dragover");
  });
});

uploadZone.addEventListener("drop", (event) => {
  const [file] = event.dataTransfer.files;
  if (file) {
    fileInput.files = event.dataTransfer.files;
    loadFile(file);
  }
});

generateBtn.addEventListener("click", renderCharts);
smartDashboardBtn.addEventListener("click", renderSmartDashboard);
clearPlotsBtn.addEventListener("click", clearCharts);
chartClearInlineBtn.addEventListener("click", clearCharts);
clearBtn.addEventListener("click", resetApp);
chartModalClose.addEventListener("click", closeChartModal);
chartModal.addEventListener("click", (event) => {
  if (event.target.dataset.closeModal === "true") {
    closeChartModal();
  }
});
document.querySelector(".chart-modal-dialog").addEventListener("click", (event) => {
  event.stopPropagation();
});
document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !chartModal.hidden) {
    closeChartModal();
  }
});
columnSearch.addEventListener("input", (event) => {
  state.columnSearchTerm = event.target.value.trim().toLowerCase();
  renderColumnMeta();
});
columnFilterTabs.addEventListener("click", (event) => {
  const filterButton = event.target.closest("[data-filter]");
  if (!filterButton) {
    return;
  }
  state.activeColumnFilter = filterButton.dataset.filter;
  updateFilterChips();
  renderColumnMeta();
});

function loadFile(file) {
  const extension = file.name.split(".").pop().toLowerCase();
  setProcessing(true, `Loading ${file.name}...`);

  if (!file.size) {
    handleFileError(file.name, "This file is empty. Upload a CSV, TXT, XLS, or XLSX file with tabular data.");
    return;
  }

  if (["xls", "xlsx"].includes(extension)) {
    const reader = new FileReader();
    reader.onload = (event) => {
      try {
        const data = new Uint8Array(event.target.result);
        const workbook = XLSX.read(data, { type: "array" });
        if (!workbook.SheetNames.length) {
          handleFileError(file.name, "No worksheets were found in this spreadsheet.");
          return;
        }
        const firstSheet = workbook.SheetNames[0];
        const sheet = workbook.Sheets[firstSheet];
        const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
        hydrateDataset(rows, file.name);
      } catch (error) {
        handleFileError(file.name, "This spreadsheet could not be read. Make sure the first sheet contains a clear header row and plain table data.");
      }
    };
    reader.onerror = () => {
      handleFileError(file.name, "The spreadsheet could not be opened in the browser.");
    };
    reader.readAsArrayBuffer(file);
    return;
  }

  const reader = new FileReader();
  reader.onload = (event) => {
    try {
      const text = event.target.result;
      const rows = parseDelimitedText(text);
      hydrateDataset(rows, file.name);
    } catch (error) {
      handleFileError(file.name, "This text file could not be parsed. Check that the file uses one delimiter consistently and includes a header row.");
    }
  };
  reader.onerror = () => {
    handleFileError(file.name, "The selected file could not be read.");
  };
  reader.readAsText(file);
}

// Parsing
function parseDelimitedText(text) {
  if (typeof text !== "string" || !text.trim()) {
    return [];
  }

  if (looksLikeBinaryText(text)) {
    return [];
  }

  const lines = splitTextIntoRows(text)
    .map((line) => line.trim())
    .filter(Boolean);

  if (lines.length < 2) {
    return [];
  }

  const delimiter = detectDelimiter(lines.slice(0, 12));
  if (!delimiter) {
    return [];
  }

  const headers = sanitizeHeaders(splitDelimitedLine(lines[0], delimiter));
  const validHeaders = headers.filter((header) => header.trim() !== "");

  if (!validHeaders.length || validHeaders.length < 2) {
    return [];
  }

  if (!hasConsistentDelimitedShape(lines, delimiter, headers.length)) {
    return [];
  }

  return lines.slice(1).map((line) => {
    const values = splitDelimitedLine(line, delimiter);
    const normalizedValues = normalizeDelimitedValues(values, headers.length);
    return headers.reduce((record, header, index) => {
      record[header] = normalizedValues[index] ?? "";
      return record;
    }, {});
  }).filter((row) => Object.values(row).some((value) => String(value).trim() !== ""));
}

function splitTextIntoRows(text) {
  const rows = [];
  let current = "";
  let inQuotes = false;

  for (let index = 0; index < text.length; index += 1) {
    const character = text[index];
    const nextCharacter = text[index + 1];

    if (character === '"') {
      current += character;
      if (inQuotes && nextCharacter === '"') {
        current += nextCharacter;
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if ((character === "\n" || character === "\r") && !inQuotes) {
      if (character === "\r" && nextCharacter === "\n") {
        index += 1;
      }
      rows.push(current);
      current = "";
      continue;
    }

    current += character;
  }

  if (current.trim()) {
    rows.push(current);
  }

  return rows;
}

function detectDelimiter(lines) {
  let best = null;
  let bestCount = 1;

  DELIMITER_CANDIDATES.forEach((delimiter) => {
    const counts = lines.map((line) => splitDelimitedLine(line, delimiter).length);
    const averageCount = counts.reduce((sum, count) => sum + count, 0) / counts.length;
    const consistentCount = counts.filter((count) => Math.abs(count - counts[0]) <= 1).length;
    const score = averageCount + consistentCount * 0.5;
    if (score > bestCount && averageCount >= 2) {
      best = delimiter;
      bestCount = score;
    }
  });

  return best;
}

function looksLikeBinaryText(text) {
  const sample = text.slice(0, 4000);
  if (!sample) {
    return false;
  }

  const nullByteCount = (sample.match(/\u0000/g) || []).length;
  if (nullByteCount > 0) {
    return true;
  }

  let suspiciousCount = 0;
  for (let index = 0; index < sample.length; index += 1) {
    const code = sample.charCodeAt(index);
    const isAllowedControl = code === 9 || code === 10 || code === 13;
    const isSuspiciousControl = code < 32 && !isAllowedControl;
    const isReplacementChar = code === 65533;
    if (isSuspiciousControl || isReplacementChar) {
      suspiciousCount += 1;
    }
  }

  return suspiciousCount / sample.length > 0.02;
}

function hasConsistentDelimitedShape(lines, delimiter, headerCount) {
  const sampleLines = lines.slice(1, Math.min(lines.length, 20));
  if (!sampleLines.length) {
    return false;
  }

  const matchingRows = sampleLines.filter((line) => {
    const values = splitDelimitedLine(line, delimiter);
    return Math.abs(values.length - headerCount) <= 1;
  }).length;

  return matchingRows >= Math.max(1, Math.ceil(sampleLines.length * 0.6));
}

function splitDelimitedLine(line, delimiter) {
  const values = [];
  let current = "";
  let inQuotes = false;

  for (let index = 0; index < line.length; index += 1) {
    const character = line[index];
    const nextCharacter = line[index + 1];

    if (character === '"') {
      if (inQuotes && nextCharacter === '"') {
        current += '"';
        index += 1;
      } else {
        inQuotes = !inQuotes;
      }
      continue;
    }

    if (character === delimiter && !inQuotes) {
      values.push(current.trim());
      current = "";
      continue;
    }

    current += character;
  }

  values.push(current.trim());
  return values;
}

function sanitizeHeaders(headers) {
  const seen = new Map();
  return headers.map((header, index) => {
    const base = String(header || "").trim() || `Column ${index + 1}`;
    const count = seen.get(base) || 0;
    seen.set(base, count + 1);
    return count ? `${base} (${count + 1})` : base;
  });
}

function normalizeDelimitedValues(values, targetLength) {
  if (values.length === targetLength) {
    return values;
  }

  if (values.length < targetLength) {
    return [...values, ...Array.from({ length: targetLength - values.length }, () => "")];
  }

  return [...values.slice(0, targetLength - 1), values.slice(targetLength - 1).join(" ")];
}

// Dataset analysis
function hydrateDataset(rows, fileName) {
  if (!rows.length) {
    resetApp();
    fileStatus.textContent = `No processable tabular rows were found in ${fileName}. Add a header row and keep delimiters consistent across lines.`;
    return;
  }

  const columns = collectColumns(rows);
  if (!columns.length || columns.every((column) => String(column).trim() === "")) {
    handleFileError(fileName, "The file does not contain usable column headers.");
    return;
  }

  const normalizedRows = rows.map((row) => normalizeRow(row, columns));
  const populatedRows = normalizedRows.filter((row) => columns.some((column) => String(row[column]).trim() !== ""));
  if (!populatedRows.length) {
    handleFileError(fileName, "The file contains headers but no processable row values.");
    return;
  }

  const sampleRows = populatedRows.slice(0, ANALYSIS_SAMPLE_LIMIT);
  const numericColumns = columns.filter((column) => isNumericColumn(sampleRows, column));
  const dateColumns = columns.filter((column) => isDateColumn(sampleRows, column));
  const idColumns = columns.filter((column) => isIdColumn(sampleRows, column, numericColumns, dateColumns));
  const categoryColumns = columns.filter((column) => isCategoryColumn(sampleRows, column, numericColumns, dateColumns, idColumns));
  const otherColumns = columns.filter((column) => !numericColumns.includes(column) && !dateColumns.includes(column) && !idColumns.includes(column) && !categoryColumns.includes(column));

  state.rows = populatedRows;
  state.sampleRows = sampleRows;
  state.columns = columns;
  state.numericColumns = numericColumns;
  state.dateColumns = dateColumns;
  state.categoryColumns = categoryColumns;
  state.idColumns = idColumns;
  state.otherColumns = otherColumns;
  state.activeColumnFilter = "all";
  state.columnSearchTerm = "";
  state.fileName = fileName;

  setProcessing(false, `${fileName} loaded. ${populatedRows.length} rows detected.`);

  updateSelectors();
  updateSummary();
  updateFilterChips();
  renderColumnMeta();
  renderPreview();
  renderRecommendations();
  renderSelectorGroups();
  clearCharts();
  columnSearch.value = "";
}

function handleFileError(fileName, message) {
  resetApp();
  fileStatus.textContent = `${fileName}: ${message} Try exporting again as CSV/XLSX with one header row and plain values.`;
}

function collectColumns(rows) {
  const uniqueColumns = new Set();
  rows.slice(0, ANALYSIS_SAMPLE_LIMIT).forEach((row) => {
    Object.keys(row).forEach((column) => uniqueColumns.add(column));
  });
  return [...uniqueColumns];
}

function normalizeRow(row, columns) {
  return columns.reduce((record, column) => {
    record[column] = row[column] ?? "";
    return record;
  }, {});
}

function isNumericColumn(rows, column) {
  const values = collectColumnValues(rows, column);
  if (!values.length) {
    return false;
  }
  return values.every((value) => Number.isFinite(Number(value)));
}

function isDateColumn(rows, column) {
  const values = collectColumnValues(rows, column).slice(0, 50);

  if (!values.length) {
    return false;
  }

  const matches = values.filter((value) => {
    const timestamp = Date.parse(value);
    return Number.isFinite(timestamp) && !/^\d+$/.test(value);
  }).length;

  return matches / values.length >= 0.7;
}

function isCategoryColumn(rows, column, numericColumns, dateColumns, idColumns) {
  if (numericColumns.includes(column) || dateColumns.includes(column) || idColumns.includes(column)) {
    return false;
  }

  const values = collectColumnValues(rows, column);

  if (!values.length) {
    return false;
  }

  const uniqueCount = new Set(values).size;
  return uniqueCount <= Math.max(12, Math.round(values.length * 0.35));
}

function isIdColumn(rows, column, numericColumns, dateColumns) {
  if (dateColumns.includes(column)) {
    return false;
  }

  const values = collectColumnValues(rows, column);
  if (!values.length) {
    return false;
  }

  const uniqueRatio = new Set(values).size / values.length;
  const shortValues = values.filter((value) => value.length <= 24).length / values.length;
  const nameLooksLikeId = /(id|uuid|code|key|serial|number|no)$/i.test(column);
  const wholeNumberLike = numericColumns.includes(column) && values.every((value) => /^-?\d+$/.test(value));

  return uniqueRatio >= 0.9 && shortValues >= 0.8 && (nameLooksLikeId || wholeNumberLike);
}

function collectColumnValues(rows, column) {
  const values = [];
  for (let index = 0; index < rows.length; index += 1) {
    const value = String(rows[index][column]).trim();
    if (value) {
      values.push(value);
    }
  }
  return values;
}

function updateSelectors() {
  populateSelect(xAxisSelect, buildSelectorOptions("x"));
  populateSelect(groupAxisSelect, buildSelectorOptions("group", true));
  populateSelect(yAxisSelect, buildSelectorOptions("y"));

  xAxisSelect.value = state.columns[0] || "";
  yAxisSelect.value = state.numericColumns[0] || state.columns[1] || state.columns[0] || "";
  groupAxisSelect.value = "None";
}

function populateSelect(select, options) {
  select.innerHTML = "";
  options.forEach((item) => {
    if (item.type === "group") {
      const group = document.createElement("optgroup");
      group.label = item.label;
      item.options.forEach((option) => {
        const element = document.createElement("option");
        element.value = option.value;
        element.textContent = option.label;
        group.appendChild(element);
      });
      select.appendChild(group);
      return;
    }

    const element = document.createElement("option");
    element.value = item.value;
    element.textContent = item.label;
    select.appendChild(element);
  });
}

function updateSummary() {
  rowCount.textContent = String(state.rows.length);
  columnCount.textContent = String(state.columns.length);
  numericCount.textContent = String(state.numericColumns.length);
}

function renderColumnMeta() {
  columnMeta.innerHTML = "";
  const visibleColumns = getFilteredColumns();
  if (!visibleColumns.length) {
    columnMeta.textContent = "No columns match the current filters.";
    return;
  }

  visibleColumns.forEach((column) => {
    const chip = document.createElement("div");
    chip.className = "meta-chip";
    const type = inferColumnRole(column);
    chip.textContent = `${column} : ${type} : ${describeColumn(column)}`;
    columnMeta.appendChild(chip);
  });
}

function renderPreview() {
  const sampleRows = state.rows.slice(0, PREVIEW_ROW_LIMIT);
  previewHead.innerHTML = "";
  previewBody.innerHTML = "";

  const headerRow = document.createElement("tr");
  state.columns.forEach((column) => {
    const th = document.createElement("th");
    th.textContent = column;
    headerRow.appendChild(th);
  });
  previewHead.appendChild(headerRow);

  sampleRows.forEach((row) => {
    const tr = document.createElement("tr");
    state.columns.forEach((column) => {
      const td = document.createElement("td");
      td.textContent = row[column];
      tr.appendChild(td);
    });
    previewBody.appendChild(tr);
  });
}

function clearCharts() {
  chartGrid.innerHTML = "";
  chartCounter = 0;
  chartMessage.textContent = "Choose charts, add them to the dashboard, or build a smart dashboard. Tip: click any chart card to expand it.";
  chartMessage.hidden = false;
}

function renderRecommendations() {
  recommendations.innerHTML = "";

  const suggestions = buildRecommendations();
  if (!suggestions.length) {
    recommendations.textContent = "No smart recommendations yet. Add at least one numeric measure column for stronger suggestions.";
    return;
  }

  suggestions.forEach((suggestion) => {
    const button = document.createElement("button");
    button.type = "button";
    button.className = "recommendation-card";
    button.addEventListener("click", () => applyRecommendation(suggestion));

    const title = document.createElement("strong");
    title.textContent = `${labelForChart(suggestion.chartType)}: ${suggestion.x} vs ${suggestion.y}`;

    const meta = document.createElement("span");
    const parts = [suggestion.reason];
    if (suggestion.group) {
      parts.push(`Group by ${suggestion.group}`);
    }
    meta.textContent = parts.join(" | ");

    button.appendChild(title);
    button.appendChild(meta);
    recommendations.appendChild(button);
  });
}

function buildRecommendations() {
  const suggestions = [];
  const rankedCategories = [...state.categoryColumns].sort((left, right) => getCategoryScore(right) - getCategoryScore(left));
  const rankedMeasures = [...state.numericColumns].sort((left, right) => getNumericScore(right) - getNumericScore(left));
  const dateColumn = state.dateColumns[0] || null;
  const category = rankedCategories[0] || null;
  const categoryAlt = rankedCategories[1] || null;
  const measure = rankedMeasures[0] || null;
  const measureAlt = rankedMeasures[1] || null;

  if (dateColumn && measure) {
    suggestions.push({
      chartType: "line",
      x: dateColumn,
      y: measure,
      group: category,
      score: 100,
      reason: "Best for trends over time",
    });
  }

  if (category && measure) {
    suggestions.push({
      chartType: "bar",
      x: category,
      y: measure,
      group: categoryAlt,
      score: 95,
      reason: "Best for comparing categories",
    });

    suggestions.push({
      chartType: "pie",
      x: category,
      y: measure,
      group: null,
      score: 72,
      reason: "Works well for part-to-whole views",
    });

    suggestions.push({
      chartType: "box",
      x: category,
      y: measure,
      group: category,
      score: 88,
      reason: "Best for spread across categories",
    });
  }

  if (measure && measureAlt) {
    suggestions.push({
      chartType: "scatter",
      x: measure,
      y: measureAlt,
      group: category,
      score: 90,
      reason: "Best for relationships between measures",
    });
  }

  if (measure) {
    suggestions.push({
      chartType: "histogram",
      x: measure,
      y: measure,
      group: category,
      score: 80,
      reason: "Best for distribution of one measure",
    });
  }

  if (!dateColumn && measure && state.columns[0] && state.columns[0] !== measure) {
    const fallbackX = rankedCategories[0] || state.columns[0];
    suggestions.push({
      chartType: "line",
      x: fallbackX,
      y: measure,
      group: category,
      score: fallbackX === state.columns[0] ? 55 : 68,
      reason: fallbackX === state.columns[0] ? "Useful if the first column has natural ordering" : "Useful for ordered category comparisons",
    });
  }

  return dedupeRecommendations(suggestions)
    .filter((suggestion) => canRenderChart(suggestion.chartType, suggestion.x, suggestion.y, state.numericColumns.includes(suggestion.y)))
    .sort((left, right) => right.score - left.score)
    .slice(0, 5);
}

function getNumericScore(column) {
  const values = collectColumnValues(state.sampleRows, column).map(Number).filter((value) => Number.isFinite(value));
  if (!values.length) {
    return 0;
  }
  const uniqueRatio = new Set(values).size / values.length;
  const max = Math.max(...values);
  const min = Math.min(...values);
  const spread = Math.abs(max - min);
  return uniqueRatio * 70 + (spread > 0 ? 30 : 0);
}

function getCategoryScore(column) {
  const values = collectColumnValues(state.sampleRows, column);
  if (!values.length) {
    return 0;
  }
  const uniqueCount = new Set(values).size;
  const ratio = uniqueCount / values.length;
  return uniqueCount <= MAX_GROUPS ? (1 - ratio) * 60 + uniqueCount : 0;
}

function dedupeRecommendations(suggestions) {
  return suggestions.filter((suggestion, index, array) => index === array.findIndex((item) =>
    item.chartType === suggestion.chartType &&
    item.x === suggestion.x &&
    item.y === suggestion.y &&
    item.group === suggestion.group
  ));
}

function applyRecommendation(suggestion) {
  xAxisSelect.value = suggestion.x;
  yAxisSelect.value = suggestion.y;
  groupAxisSelect.value = suggestion.group || "None";

  [...chartOptions.querySelectorAll("input")].forEach((input) => {
    input.checked = input.value === suggestion.chartType;
  });

  addChartCard(suggestion.chartType, suggestion.x, suggestion.y, suggestion.group);
}

function renderCharts() {
  if (!state.rows.length) {
    chartMessage.textContent = "Please upload a dataset first.";
    chartMessage.hidden = false;
    return;
  }

  const selectedCharts = [...chartOptions.querySelectorAll("input:checked")].map((input) => input.value);
  if (!selectedCharts.length) {
    chartMessage.textContent = "Pick at least one chart type.";
    chartMessage.hidden = false;
    return;
  }

  const xColumn = xAxisSelect.value;
  const yColumn = yAxisSelect.value;
  const groupColumn = groupAxisSelect.value === "None" ? null : groupAxisSelect.value;
  const numericY = state.numericColumns.includes(yColumn);
  showChartLoading("Building charts...");

  selectedCharts.forEach((chartType, index) => {
    if (!canRenderChart(chartType, xColumn, yColumn, numericY)) {
      return;
    }
    addChartCard(chartType, xColumn, yColumn, groupColumn, index);
  });

  if (!chartGrid.children.length) {
    chartMessage.textContent = "The chosen columns do not fit the selected chart types.";
    chartMessage.hidden = false;
    return;
  }
  chartMessage.hidden = true;
  focusChartPanel();
}

function renderSmartDashboard() {
  if (!state.rows.length) {
    chartMessage.textContent = "Please upload a dataset first.";
    chartMessage.hidden = false;
    return;
  }

  const suggestions = buildRecommendations();
  if (!suggestions.length) {
    chartMessage.textContent = "No smart dashboard could be built from this dataset yet.";
    chartMessage.hidden = false;
    return;
  }

  showChartLoading("Building a smart dashboard...");
  suggestions.slice(0, 4).forEach((suggestion, index) => {
    addChartCard(suggestion.chartType, suggestion.x, suggestion.y, suggestion.group, index);
  });
  chartMessage.hidden = true;
  focusChartPanel();
}

// Chart rendering
function addChartCard(chartType, xColumn, yColumn, groupColumn, index = 0) {
  const numericY = state.numericColumns.includes(yColumn);
  if (!canRenderChart(chartType, xColumn, yColumn, numericY)) {
    return false;
  }

  chartMessage.hidden = true;

  const card = document.createElement("article");
  card.className = "chart-card";

  const header = document.createElement("div");
  header.className = "chart-card-header";

  const title = document.createElement("h3");
  title.textContent = chartTitle(chartType, xColumn, yColumn);

  const removeButton = document.createElement("button");
  removeButton.type = "button";
  removeButton.className = "chart-remove-btn";
  removeButton.textContent = "Remove";
  removeButton.addEventListener("click", (event) => {
    event.stopPropagation();
    card.remove();
    if (!chartGrid.children.length) {
      chartMessage.textContent = "Choose charts, add them to the dashboard, or build a smart dashboard. Tip: click any chart card to expand it.";
      chartMessage.hidden = false;
    }
  });

  const actions = document.createElement("div");
  actions.className = "chart-card-actions";

  const exportButton = document.createElement("button");
  exportButton.type = "button";
  exportButton.className = "chart-remove-btn";
  exportButton.textContent = "Export PNG";
  exportButton.title = "Download this chart as an image";
  exportButton.addEventListener("click", async (event) => {
    event.stopPropagation();
    await Plotly.downloadImage(surface, {
      format: "png",
      filename: buildExportName(chartType, xColumn, yColumn),
      width: 1280,
      height: 720,
    });
  });

  actions.appendChild(exportButton);
  actions.appendChild(removeButton);
  header.appendChild(title);
  header.appendChild(actions);

  const surface = document.createElement("div");
  surface.className = "plot-surface";
  surface.id = `plot-${chartType}-${chartCounter}-${index}`;
  surface.addEventListener("click", () => {
    openChartModal(chartType, xColumn, yColumn, groupColumn);
  });

  card.appendChild(header);
  card.appendChild(surface);
  chartGrid.appendChild(card);

  const config = buildPlotConfig(chartType, xColumn, yColumn, groupColumn);
  Plotly.react(surface, config.data, config.layout, {
    responsive: true,
    displaylogo: false,
  });
  chartCounter += 1;
  return true;
}

function openChartModal(chartType, xColumn, yColumn, groupColumn) {
  const config = buildPlotConfig(chartType, xColumn, yColumn, groupColumn, true);
  chartModalTitle.textContent = chartTitle(chartType, xColumn, yColumn);
  chartModal.hidden = false;
  document.body.classList.add("modal-open");
  Plotly.react(chartModalSurface, config.data, config.layout, {
    responsive: true,
    displaylogo: false,
  });
}

function closeChartModal() {
  chartModal.hidden = true;
  document.body.classList.remove("modal-open");
  Plotly.purge(chartModalSurface);
}

function canRenderChart(chartType, xColumn, yColumn, numericY) {
  if (["scatter", "line", "bar", "box", "histogram", "pie"].includes(chartType)) {
    return Boolean(xColumn) && Boolean(yColumn) && numericY;
  }
  return false;
}

function buildPlotConfig(chartType, xColumn, yColumn, groupColumn, isExpanded = false) {
  const palette = ["#ef6c33", "#2563eb", "#10b981", "#f59e0b", "#7c3aed", "#ef4444"];
  const chartRows = getChartRows(chartType, xColumn, yColumn, groupColumn);
  const grouped = groupColumn ? groupRowsByColumn(chartRows, groupColumn) : { All: chartRows };
  const traces = Object.entries(grouped).map(([groupName, rows], index) => {
    const xValues = rows.map((row) => row[xColumn]);
    const yValues = rows.map((row) => Number(row[yColumn]));
    const color = palette[index % palette.length];

    switch (chartType) {
      case "scatter":
        return {
          type: "scatter",
          mode: "markers",
          name: groupName,
          x: xValues,
          y: yValues,
          marker: { color, size: 10, opacity: 0.8 },
        };
      case "line":
        return {
          type: "scatter",
          mode: "lines+markers",
          name: groupName,
          x: xValues,
          y: yValues,
          line: { color, width: 3 },
        };
      case "bar":
        return {
          type: "bar",
          name: groupName,
          x: xValues,
          y: yValues,
          marker: { color },
        };
      case "histogram":
        return {
          type: "histogram",
          name: groupName,
          x: yValues,
          marker: { color, opacity: 0.8 },
        };
      case "box":
        return {
          type: "box",
          name: groupName,
          y: yValues,
          marker: { color },
          boxpoints: "outliers",
        };
      case "pie":
        return {
          type: "pie",
          name: groupName,
          labels: xValues,
          values: yValues,
          textinfo: "label+percent",
        };
      default:
        return null;
    }
  }).filter(Boolean);

  return {
    data: traces,
    layout: {
      autosize: true,
      height: isExpanded ? 640 : 300,
      paper_bgcolor: "rgba(0,0,0,0)",
      plot_bgcolor: "rgba(255,255,255,0.72)",
      font: {
        family: "Space Grotesk, sans-serif",
        color: "#1f2937",
      },
      margin: { l: 52, r: 20, t: 12, b: 48 },
      xaxis: chartType === "histogram" || chartType === "box" ? undefined : { title: xColumn },
      yaxis: chartType === "pie" ? undefined : { title: yColumn },
      legend: { orientation: "h" },
      barmode: groupColumn && chartType === "bar" ? "group" : undefined,
      hovermode: "closest",
    },
  };
}

function getChartRows(chartType, xColumn, yColumn, groupColumn) {
  const baseRows = state.rows.length > ANALYSIS_SAMPLE_LIMIT * 4 ? state.rows.slice(0, ANALYSIS_SAMPLE_LIMIT * 2) : state.rows;

  if (chartType === "bar" || chartType === "pie") {
    return summarizeRows(baseRows, xColumn, yColumn, groupColumn);
  }

  return baseRows;
}

function summarizeRows(rows, xColumn, yColumn, groupColumn) {
  const groups = new Map();

  rows.forEach((row) => {
    const categoryKey = String(row[xColumn] || "Unspecified");
    const groupKey = groupColumn ? String(row[groupColumn] || "Unspecified") : "All";
    const key = `${groupKey}__${categoryKey}`;
    const current = groups.get(key) || { [xColumn]: categoryKey, [yColumn]: 0, [groupColumn || "__group"]: groupKey };
    current[yColumn] += Number(row[yColumn]) || 0;
    groups.set(key, current);
  });

  return [...groups.values()];
}

function groupRowsByColumn(rows, column) {
  return rows.reduce((groups, row) => {
    const key = row[column] || "Unspecified";
    if (!groups[key]) {
      groups[key] = [];
    }
    groups[key].push(row);
    return groups;
  }, {});
}

function buildExportName(chartType, xColumn, yColumn) {
  return `${chartType}-${String(xColumn).toLowerCase().replace(/\s+/g, "-")}-${String(yColumn).toLowerCase().replace(/\s+/g, "-")}`;
}

function chartTitle(chartType, xColumn, yColumn) {
  const titles = {
    scatter: `Scatter Plot: ${xColumn} vs ${yColumn}`,
    line: `Line Plot: ${xColumn} vs ${yColumn}`,
    bar: `Bar Chart: ${xColumn} vs ${yColumn}`,
    histogram: `Histogram: ${yColumn}`,
    box: `Box Plot: ${yColumn}`,
    pie: `Pie Chart: ${xColumn} by ${yColumn}`,
  };
  return titles[chartType] || "Chart";
}

function labelForChart(chartType) {
  const labels = {
    scatter: "Scatter",
    line: "Line",
    bar: "Bar",
    histogram: "Histogram",
    box: "Box",
    pie: "Pie",
  };
  return labels[chartType] || "Chart";
}

function buildSelectorOptions(target, includeNone = false) {
  const options = [];
  if (includeNone) {
    options.push({ type: "option", value: "None", label: "None" });
  }

  const groups = [
    { key: "numeric", label: "Numeric Columns", columns: state.numericColumns },
    { key: "category", label: "Category Columns", columns: state.categoryColumns },
    { key: "date", label: "Date Columns", columns: state.dateColumns },
    { key: "id", label: "ID Columns", columns: state.idColumns },
    { key: "other", label: "Other Columns", columns: state.otherColumns },
  ];

  groups.forEach((group) => {
    let columns = group.columns;
    if (target === "y" && group.key !== "numeric") {
      columns = [];
    }
    if (!columns.length) {
      return;
    }
    options.push({
      type: "group",
      label: group.label,
      options: columns.map((column) => ({ value: column, label: column })),
    });
  });

  if (target === "y" && !state.numericColumns.length && state.columns.length) {
    options.push({
      type: "group",
      label: "All Columns",
      options: state.columns.map((column) => ({ value: column, label: column })),
    });
  }

  return options;
}

function renderSelectorGroups() {
  if (!state.columns.length) {
    selectorGroups.textContent = "Load a file to browse grouped column options.";
    return;
  }

  selectorGroups.innerHTML = "";
  [
    { title: "Numeric", columns: state.numericColumns },
    { title: "Category", columns: state.categoryColumns },
    { title: "Date", columns: state.dateColumns },
    { title: "ID", columns: state.idColumns },
    { title: "Other", columns: state.otherColumns },
  ].forEach((section) => {
    const card = document.createElement("div");
    card.className = "selector-group-card";

    const title = document.createElement("strong");
    title.textContent = `${section.title} (${section.columns.length})`;

    const body = document.createElement("span");
    body.textContent = section.columns.length ? section.columns.join(", ") : "None detected";

    card.appendChild(title);
    card.appendChild(body);
    selectorGroups.appendChild(card);
  });
}

function inferColumnRole(column) {
  if (state.idColumns.includes(column)) {
    return "id";
  }
  if (state.dateColumns.includes(column)) {
    return "date";
  }
  if (state.numericColumns.includes(column)) {
    return "numeric";
  }
  if (state.categoryColumns.includes(column)) {
    return "category";
  }
  return "other";
}

function getFilteredColumns() {
  return state.columns.filter((column) => {
    const matchesSearch = !state.columnSearchTerm || column.toLowerCase().includes(state.columnSearchTerm);
    const matchesFilter = state.activeColumnFilter === "all" || inferColumnRole(column) === state.activeColumnFilter;
    return matchesSearch && matchesFilter;
  });
}

function updateFilterChips() {
  [...columnFilterTabs.querySelectorAll("[data-filter]")].forEach((button) => {
    button.classList.toggle("active", button.dataset.filter === state.activeColumnFilter);
  });
}

function describeColumn(column) {
  const values = collectColumnValues(state.sampleRows, column);
  if (!values.length) {
    return "empty";
  }

  if (state.numericColumns.includes(column)) {
    const numbers = values.map(Number);
    return `min ${formatNumber(Math.min(...numbers))}, max ${formatNumber(Math.max(...numbers))}`;
  }

  return `${new Set(values).size} unique`;
}

function formatNumber(value) {
  if (!Number.isFinite(value)) {
    return "-";
  }
  return Number.isInteger(value) ? String(value) : value.toFixed(2);
}

function resetApp() {
  state.rows = [];
  state.sampleRows = [];
  state.columns = [];
  state.numericColumns = [];
  state.dateColumns = [];
  state.categoryColumns = [];
  state.idColumns = [];
  state.otherColumns = [];
  state.activeColumnFilter = "all";
  state.columnSearchTerm = "";
  state.fileName = "";
  state.isProcessing = false;

  fileInput.value = "";
  fileStatus.textContent = "No file loaded yet.";
  fileStatus.classList.remove("is-loading");
  rowCount.textContent = "0";
  columnCount.textContent = "0";
  numericCount.textContent = "0";
  xAxisSelect.innerHTML = "";
  yAxisSelect.innerHTML = "";
  groupAxisSelect.innerHTML = "";
  columnMeta.innerHTML = "Load a file to inspect column types.";
  selectorGroups.innerHTML = "Load a file to browse grouped column options.";
  recommendations.innerHTML = "Load a file to see recommended plot combinations.";
  columnSearch.value = "";
  processingHint.hidden = true;
  updateFilterChips();
  previewHead.innerHTML = "";
  previewBody.innerHTML = "";
  clearCharts();
}

// UI helpers
function showChartLoading(message) {
  chartMessage.textContent = message;
  chartMessage.hidden = false;
}

function focusChartPanel() {
  chartPanel?.scrollIntoView({ behavior: "smooth", block: "start" });
}
