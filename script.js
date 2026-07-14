const state = {
  rawRows: [],
  rows: [],
  sampleRows: [],
  normalizedDateRows: [],
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
  aiDashboardLoading: false,
};

const ANALYSIS_SAMPLE_LIMIT = 1000;
const PREVIEW_ROW_LIMIT = 8;
const DELIMITER_CANDIDATES = [",", "\t", ";", "|"];
const MAX_GROUPS = 20;
const EMPTY_LIKE_VALUES = new Set(["", "null", "n/a", "na", "-"]);
const LOCAL_EDA_ENDPOINT = "http://127.0.0.1:5000/eda";
const ALLOWED_AI_CHART_TYPES = new Set(["bar", "line", "scatter", "histogram", "box", "pie"]);
const SMART_DASHBOARD_TIMEOUT_MS = 30000;
const SUMMARY_STRING_LIMIT = 30;

const fileInput = document.getElementById("file-input");
const fileStatus = document.getElementById("file-status");
const chartOptions = document.getElementById("chart-options");
const xAxisSelect = document.getElementById("x-axis");
const yAxisSelect = document.getElementById("y-axis");
const groupAxisSelect = document.getElementById("group-axis");
const aggregationSelect = document.getElementById("aggregation");
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
const selectorWarning = document.getElementById("selector-warning");
let chartCounter = 0;

// State helpers
function setProcessing(isProcessing, message) {
  state.isProcessing = isProcessing;
  fileStatus.textContent = message;
  fileStatus.classList.toggle("is-loading", isProcessing);
  processingHint.hidden = !isProcessing;
}

function setSmartDashboardLoading(isLoading) {
  state.aiDashboardLoading = isLoading;
  smartDashboardBtn.disabled = isLoading;
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
chartOptions.addEventListener("change", updateSelectorWarning);
[xAxisSelect, yAxisSelect, groupAxisSelect, aggregationSelect].forEach((select) => {
  select.addEventListener("change", () => {
    syncSelectorOptions();
    updateSelectorWarning();
  });
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
  const populatedRawRows = normalizedRows.filter((row) => columns.some((column) => String(row[column]).trim() !== ""));
  if (!populatedRawRows.length) {
    handleFileError(fileName, "The file contains headers but no processable row values.");
    return;
  }

  const cleanedRows = populatedRawRows.map((row) => cleanAndNormalizeRow(row, columns));
  const populatedRows = cleanedRows.filter((row) => columns.some((column) => row[column] !== null));
  if (!populatedRows.length) {
    handleFileError(fileName, "The file contains only empty or invalid row values after cleaning.");
    return;
  }

  const sampleRows = populatedRows.slice(0, ANALYSIS_SAMPLE_LIMIT);
  const normalizedDateRows = populatedRows.map((row) => buildNormalizedDateRow(row, columns));
  const numericColumns = columns.filter((column) => isNumericColumn(sampleRows, column));
  const dateColumns = columns.filter((column) => isDateColumn(sampleRows, normalizedDateRows, column));
  const idColumns = columns.filter((column) => isIdColumn(sampleRows, column, numericColumns, dateColumns));
  const categoryColumns = columns.filter((column) => isCategoryColumn(sampleRows, column, numericColumns, dateColumns, idColumns));
  const otherColumns = columns.filter((column) => !numericColumns.includes(column) && !dateColumns.includes(column) && !idColumns.includes(column) && !categoryColumns.includes(column));

  state.rawRows = populatedRawRows;
  state.rows = populatedRows;
  state.sampleRows = sampleRows;
  state.normalizedDateRows = normalizedDateRows;
  state.columns = columns;
  state.numericColumns = numericColumns;
  state.dateColumns = dateColumns;
  state.categoryColumns = categoryColumns;
  state.idColumns = idColumns;
  state.otherColumns = otherColumns;
  state.activeColumnFilter = "all";
  state.columnSearchTerm = "";
  state.fileName = fileName;

  setProcessing(false, `${fileName} loaded. ${populatedRows.length} cleaned rows detected.`);

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

function cleanAndNormalizeRow(row, columns) {
  return columns.reduce((record, column) => {
    record[column] = normalizeCellValue(row[column]);
    return record;
  }, {});
}

function normalizeCellValue(value) {
  if (value === null || value === undefined) {
    return null;
  }

  if (typeof value === "number") {
    return Number.isFinite(value) ? value : null;
  }

  if (value instanceof Date) {
    return Number.isFinite(value.getTime()) ? value.toISOString() : null;
  }

  const normalized = String(value).trim();
  if (EMPTY_LIKE_VALUES.has(normalized.toLowerCase())) {
    return null;
  }

  return normalized;
}

function isNumericColumn(rows, column) {
  const values = collectColumnValues(rows, column);
  if (!values.length) {
    return false;
  }
  const numericMatches = values.filter((value) => Number.isFinite(parseNumericValue(value)));
  return numericMatches.length / values.length >= 0.9;
}

function isDateColumn(rows, normalizedDateRows, column) {
  const originalValues = collectColumnValues(rows, column).slice(0, 50);
  if (!originalValues.length) {
    return false;
  }

  const numericLikeCount = originalValues.filter((value) => Number.isFinite(parseNumericValue(value))).length;
  if (numericLikeCount / originalValues.length >= 0.8) {
    return false;
  }

  const values = rows
    .map((_, index) => normalizedDateRows[index]?.[column])
    .filter((value) => value !== null)
    .slice(0, 50);

  if (!values.length) {
    return false;
  }

  return values.length / Math.min(rows.length, 50) >= 0.7;
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
    const value = rows[index][column];
    if (value !== null && value !== undefined && String(value).trim() !== "") {
      values.push(value);
    }
  }
  return values;
}

function parseNumericValue(value) {
  if (typeof value === "number") {
    return Number.isFinite(value) ? value : NaN;
  }

  if (value === null || value === undefined) {
    return NaN;
  }

  const normalized = String(value)
    .trim()
    .replace(/[$€£¥₹,\s]/g, "")
    .replace(/\((.+)\)/, "-$1");

  if (!normalized) {
    return NaN;
  }

  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : NaN;
}

function parseDateValue(value) {
  if (value === null || value === undefined) {
    return null;
  }

  if (value instanceof Date) {
    const timestamp = value.getTime();
    return Number.isFinite(timestamp) ? timestamp : null;
  }

  const normalized = String(value).trim();
  if (!normalized || /^\d+(?:\.\d+)?$/.test(normalized)) {
    return null;
  }

  const looksLikeDate = /[\/-]/.test(normalized) || /[a-z]/i.test(normalized);
  if (!looksLikeDate) {
    return null;
  }

  const timestamp = Date.parse(normalized);
  return Number.isFinite(timestamp) ? timestamp : null;
}

function buildNormalizedDateRow(row, columns) {
  return columns.reduce((record, column) => {
    record[column] = parseDateValue(row[column]);
    return record;
  }, {});
}

function getMultiSelectValues(select) {
  return Array.from(select.selectedOptions).map(opt => opt.value).filter(val => val !== "None");
}

function restoreMultiSelect(select, values) {
  if (!values.length) {
    if (select.options.length) select.options[0].selected = true;
    return;
  }
  Array.from(select.options).forEach(opt => {
    opt.selected = values.includes(opt.value);
  });
}

function updateSelectors() {
  syncSelectorOptions();
  restoreMultiSelect(xAxisSelect, []);
  restoreMultiSelect(yAxisSelect, []);
  groupAxisSelect.value = "None";
  aggregationSelect.value = "sum";
  updateSelectorWarning();
}

function syncSelectorOptions() {
  const currentSelection = {
    x: getMultiSelectValues(xAxisSelect),
    y: getMultiSelectValues(yAxisSelect),
    group: groupAxisSelect.value || "None",
  };

  const xExcludes = [...currentSelection.y, currentSelection.group];
  const yExcludes = [...currentSelection.x, currentSelection.group];
  const groupExcludes = [...currentSelection.x, ...currentSelection.y];

  populateSelect(xAxisSelect, buildSelectorOptions("x", true, { exclude: xExcludes }));
  populateSelect(yAxisSelect, buildSelectorOptions("y", true, { exclude: yExcludes }));
  populateSelect(groupAxisSelect, buildSelectorOptions("group", true, { exclude: groupExcludes }));

  restoreMultiSelect(xAxisSelect, currentSelection.x);
  restoreMultiSelect(yAxisSelect, currentSelection.y);
  groupAxisSelect.value = optionExists(groupAxisSelect, currentSelection.group) ? currentSelection.group : "None";
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
    if (item.disabled) {
      element.disabled = true;
    }
    select.appendChild(element);
  });
}

function optionExists(select, value) {
  return [...select.options].some((option) => option.value === value);
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
    chip.textContent = `${column} (${type})`;
    chip.title = describeColumn(column);
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
  setChartAreaMessage("No charts yet. Select columns and click 'Add to Dashboard' to visualize your data.");
}

function buildDatasetSummary() {
  return {
    columns: state.columns.map((column) => ({
      name: column,
      type: inferSummaryColumnType(column),
    })),
    sample_rows: state.rows.slice(0, 5).map((row) => {
      return state.columns.reduce((record, column) => {
        record[column] = serializeSummaryValue(row[column]);
        return record;
      }, {});
    }),
  };
}

function inferSummaryColumnType(column) {
  if (state.numericColumns.includes(column)) {
    return "numeric";
  }
  if (state.dateColumns.includes(column)) {
    return "date";
  }
  if (state.categoryColumns.includes(column) || state.idColumns.includes(column)) {
    return "category";
  }
  return "other";
}

function serializeSummaryValue(value) {
  if (value === null || value === undefined) {
    return null;
  }
  if (typeof value === "number") {
    return value;
  }
  if (value instanceof Date) {
    return value.toISOString().slice(0, SUMMARY_STRING_LIMIT);
  }
  const textValue = String(value).trim();
  return textValue.length > SUMMARY_STRING_LIMIT ? `${textValue.slice(0, SUMMARY_STRING_LIMIT)}...` : textValue;
}

async function requestSmartDashboardSuggestions(summary) {
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), SMART_DASHBOARD_TIMEOUT_MS);

  let response;
  try {
    response = await fetch(LOCAL_EDA_ENDPOINT, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({ summary }),
      signal: controller.signal,
    });
  } finally {
    clearTimeout(timeoutId);
  }

  const payload = await response.json().catch(() => ({}));
  if (!response.ok) {
    throw new Error(payload.error || "Local EDA request failed.");
  }
  return normalizeAiDashboard(payload);
}

function normalizeAiDashboard(payload) {
  const insights = Array.isArray(payload?.insights)
    ? payload.insights
      .map((item) => String(item || "").trim())
      .filter(Boolean)
    : [];

  const charts = Array.isArray(payload?.charts)
    ? payload.charts
      .map(normalizeAiChart)
      .filter(Boolean)
    : [];

  return { insights, charts };
}

function normalizeAiChart(chart) {
  if (!chart || typeof chart !== "object") {
    return null;
  }

  const chartType = String(chart.type || "").trim().toLowerCase();
  if (!ALLOWED_AI_CHART_TYPES.has(chartType)) {
    return null;
  }

  let x = normalizeAiColumn(chart.x);
  const y = normalizeAiColumn(chart.y);
  let group = normalizeAiOptionalColumn(chart.group);
  const reason = String(chart.reason || "").trim();

  if (!y || !state.numericColumns.includes(y)) {
    return null;
  }

  if (chartType === "histogram") {
    x = x || y;
  }

  if (chartType === "box" && !group && x && (state.categoryColumns.includes(x) || state.dateColumns.includes(x))) {
    group = x;
  }

  if (!x || !state.columns.includes(x)) {
    return null;
  }

  const validation = canRenderChart(chartType, x ? [x] : [], y ? [y] : [], group);
  if (!validation.valid) {
    return null;
  }

  return {
    chartType,
    x,
    y,
    group,
    reason,
  };
}

function normalizeAiColumn(value) {
  if (typeof value !== "string") {
    return null;
  }
  const normalized = value.trim();
  return state.columns.includes(normalized) ? normalized : null;
}

function normalizeAiOptionalColumn(value) {
  if (value === null || value === undefined || value === "" || value === "null") {
    return null;
  }
  return normalizeAiColumn(value);
}

function renderAiInsights(insights) {
  recommendations.innerHTML = "";

  if (!insights.length) {
    recommendations.textContent = "AI insights will appear here after building the smart dashboard.";
    return;
  }

  insights.forEach((insight) => {
    const card = document.createElement("div");
    card.className = "recommendation-card";

    const title = document.createElement("strong");
    title.textContent = "Insight";

    const meta = document.createElement("span");
    meta.textContent = insight;

    card.appendChild(title);
    card.appendChild(meta);
    recommendations.appendChild(card);
  });
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
  const rankedDates = [...state.dateColumns].sort((left, right) => getDateScore(right) - getDateScore(left));
  const dateColumn = rankedDates[0] || null;
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
    const fallbackX = rankedCategories[0] || state.dateColumns[0] || state.columns.find((column) => column !== measure) || state.columns[0];
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
    .filter((suggestion) => canRenderChart(suggestion.chartType, suggestion.x ? [suggestion.x] : [], suggestion.y ? [suggestion.y] : [], suggestion.group).valid)
    .sort((left, right) => right.score - left.score)
    .slice(0, 5);
}

function getNumericScore(column) {
  const values = collectColumnValues(state.sampleRows, column).map(parseNumericValue).filter((value) => Number.isFinite(value));
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

function getDateScore(column) {
  const values = collectColumnValues(state.sampleRows, column);
  if (!values.length) {
    return 0;
  }
  const uniqueCount = new Set(values).size;
  return uniqueCount >= 2 ? uniqueCount : 0;
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
  restoreMultiSelect(xAxisSelect, [suggestion.x]);
  restoreMultiSelect(yAxisSelect, [suggestion.y]);
  groupAxisSelect.value = suggestion.group || "None";
  syncSelectorOptions();

  [...chartOptions.querySelectorAll("input")].forEach((input) => {
    input.checked = input.value === suggestion.chartType;
  });

  updateSelectorWarning();
  addChartCard(suggestion.chartType, [suggestion.x], [suggestion.y], suggestion.group);
}

async function renderCharts() {
  if (!state.rows.length) {
    setChartAreaMessage("Please upload a dataset first.");
    return;
  }

  const selectedCharts = [...chartOptions.querySelectorAll("input:checked")].map((input) => input.value);
  if (!selectedCharts.length) {
    setChartAreaMessage("Pick at least one chart type.");
    return;
  }

  const xColumns = getMultiSelectValues(xAxisSelect);
  const yColumns = getMultiSelectValues(yAxisSelect);
  const groupColumn = groupAxisSelect.value === "None" ? null : groupAxisSelect.value;
  
  const inlineFeedback = getInlineSelectionFeedback(selectedCharts, xColumns, yColumns, groupColumn);
  if (inlineFeedback) {
    showSelectorWarning(inlineFeedback);
  }

  showChartLoading("Building charts...");
  await waitForNextPaint();

  let renderedCount = 0;

  selectedCharts.forEach((chartType, index) => {
    const result = addChartCard(chartType, xColumns, yColumns, groupColumn, index);
    if (result) {
      renderedCount += 1;
    }
  });

  if (!renderedCount) {
    setChartAreaMessage("This chart cannot be generated with the selected columns. Please choose different columns.");
    return;
  }
  chartMessage.hidden = true;
  chartMessage.classList.remove("is-loading");
  focusChartPanel();
}

async function renderSmartDashboard() {
  if (!state.rows.length) {
    setChartAreaMessage("Please upload a dataset first.");
    return;
  }

  if (state.aiDashboardLoading) {
    return;
  }

  setSmartDashboardLoading(true);
  showChartLoading("Building a smart dashboard...");
  await waitForNextPaint();

  try {
    const summary = buildDatasetSummary();
    const result = await requestSmartDashboardSuggestions(summary);

    renderAiInsights(result.insights);

    if (!result.charts.length) {
      setChartAreaMessage("The AI did not return any valid charts for this dataset.");
      return;
    }

    clearCharts();

    let renderedCount = 0;
    result.charts.forEach((suggestion, index) => {
      const rendered = addChartCard(suggestion.chartType, [suggestion.x], [suggestion.y], suggestion.group, index);
      if (rendered) {
        renderedCount += 1;
      }
    });

    if (!renderedCount) {
      setChartAreaMessage("The AI returned chart suggestions, but none matched the current dataset columns.");
      return;
    }

    chartMessage.hidden = true;
    chartMessage.classList.remove("is-loading");
    focusChartPanel();
  } catch (error) {
    console.error(error);
    const message = error?.name === "AbortError"
      ? "AI request timed out. Try a smaller dataset or try again in a moment."
      : "AI dashboard generation failed. Make sure the backend is running and try again.";
    setChartAreaMessage(message);
  } finally {
    setSmartDashboardLoading(false);
  }
}

// Chart rendering
function addChartCard(chartType, xColumn, yColumn, groupColumn, index = 0) {
  const validation = canRenderChart(chartType, xColumn, yColumn, groupColumn);
  if (!validation.valid) {
    addChartFeedbackCard(chartType, validation.message);
    return false;
  }

  chartMessage.hidden = true;
  chartMessage.classList.remove("is-loading");

  const card = document.createElement("article");
  card.className = "chart-card";

  const header = document.createElement("div");
  header.className = "chart-card-header";

  const title = document.createElement("h3");
  title.textContent = chartTitle(chartType, xColumns, yColumns);

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
      filename: buildExportName(chartType, xColumns, yColumns),
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
    openChartModal(chartType, xColumns, yColumns, groupColumn);
  });

  card.appendChild(header);
  card.appendChild(surface);
  chartGrid.appendChild(card);

  const config = buildPlotConfig(chartType, xColumns, yColumns, groupColumn);
  if (!config.data.length) {
    card.remove();
    addChartFeedbackCard(chartType, "This chart cannot be generated with the selected columns. Please choose different columns.");
    return false;
  }
  Plotly.react(surface, config.data, config.layout, {
    responsive: true,
    displaylogo: false,
  });
  chartCounter += 1;
  return true;
}

function openChartModal(chartType, xColumns, yColumns, groupColumn) {
  const config = buildPlotConfig(chartType, xColumns, yColumns, groupColumn, true);
  chartModalTitle.textContent = chartTitle(chartType, xColumns, yColumns);
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

function canRenderChart(chartType, xColumns, yColumns, groupColumn) {
  const validation = validateChartSelection(chartType, xColumns, yColumns, groupColumn, state.sampleRows);
  return {
    valid: validation.valid,
    message: validation.message,
  };
}

function validateChartSelection(chartType, xColumns, yColumns, groupColumn, rows) {
  if (!rows.length) {
    return { valid: false, message: "Please upload a dataset first." };
  }

  if (!chartType) {
    return { valid: false, message: "Select at least one chart type." };
  }

  const requiresX = chartType !== "histogram" && chartType !== "box" && chartType !== "pie";
  if (requiresX && (!xColumns.length || !xColumns.every(col => state.columns.includes(col)))) {
    return { valid: false, message: "Choose at least one valid X axis or category column." };
  }

  if (chartType === "pie" && (!xColumns.length || !xColumns.every(col => state.columns.includes(col)))) {
    return { valid: false, message: "Choose at least one valid X axis (category) for a pie chart." };
  }

  if (!yColumns.length || !yColumns.every(col => state.columns.includes(col))) {
    return { valid: false, message: "Choose at least one valid Y axis or value column." };
  }

  if (groupColumn && !state.columns.includes(groupColumn)) {
    return { valid: false, message: "Choose a valid group column or set Group / Color to None." };
  }

  if (!yColumns.every(col => state.numericColumns.includes(col))) {
    return { valid: false, message: `All Y axis columns must contain numeric values for this chart.` };
  }

  if (groupColumn) {
    const groupValues = collectColumnValues(rows, groupColumn);
    const uniqueGroups = new Set(groupValues).size;
    if (!groupValues.length || uniqueGroups < 1) {
      return { valid: false, message: "The selected group column has no usable values." };
    }
    if (uniqueGroups > MAX_GROUPS) {
      return { valid: false, message: "The selected group column has too many unique values for a readable grouped chart." };
    }
  }

  const validRows = getRenderableRows(rows, xColumns, yColumns, groupColumn, chartType);
  if (validRows.length < 2) {
    return { valid: false, message: "At least 2 valid data points are required to generate this chart." };
  }

  if (chartType === "pie") {
    const pieRows = summarizeRows(validRows, xColumns, yColumns, groupColumn, "sum").filter((row) => yColumns.some(yCol => parseNumericValue(row[yCol]) > 0));
    if (pieRows.length < 2) {
      return { valid: false, message: "Pie charts need at least 2 categories with positive numeric values." };
    }
  }

  if (chartType === "histogram") {
    const numericValues = validRows.flatMap((row) => yColumns.map(yCol => parseNumericValue(row[yCol]))).filter((value) => Number.isFinite(value));
    if (new Set(numericValues).size < 2) {
      return { valid: false, message: "Histogram charts need varied numeric values." };
    }
  }

  return { valid: true, message: "" };
}

function buildPlotConfig(chartType, xColumns, yColumns, groupColumn, isExpanded = false) {
  const palette = ["#38bdf8", "#8b5cf6", "#ec4899", "#10b981", "#f59e0b", "#f43f5e"];
  const aggregation = aggregationSelect.value;
  const chartRows = getChartRows(chartType, xColumns, yColumns, groupColumn, aggregation);
  const grouped = groupColumn ? groupRowsByColumn(chartRows, groupColumn) : { All: chartRows };
  
  const traces = [];
  let traceIndex = 0;
  
  Object.entries(grouped).forEach(([groupName, rows]) => {
    const xValues = rows.map((row) => xColumns.map(col => getChartDisplayValue(row, col)).join(" - "));
      
    yColumns.forEach(yCol => {
      const yValues = rows.map((row) => parseNumericValue(row[yCol]));
      const color = palette[traceIndex % palette.length];
      const name = yColumns.length > 1 ? (groupColumn ? `${groupName} - ${yCol}` : yCol) : groupName;
      
      switch (chartType) {
        case "scatter":
          traces.push({
            type: "scatter", mode: "markers", name,
            x: xValues, y: yValues, marker: { color, size: 10, opacity: 0.8 },
          }); break;
        case "line":
          traces.push({
            type: "scatter", mode: "lines+markers", name,
            x: xValues, y: yValues, line: { color, width: 3 },
          }); break;
        case "bar":
          traces.push({
            type: "bar", name,
            x: xValues, y: yValues, marker: { color },
          }); break;
        case "histogram":
          traces.push({
            type: "histogram", name,
            x: yValues, marker: { color, opacity: 0.8 },
          }); break;
        case "box":
          traces.push({
            type: "box", name,
            x: xColumns.length > 0 && state.columns.includes(xColumns[0]) ? xValues : undefined,
            y: yValues, marker: { color }, boxpoints: "outliers",
          }); break;
        case "pie":
          traces.push({
            type: "pie", name,
            labels: xValues, values: yValues, textinfo: "label+percent",
          }); break;
      }
      traceIndex += 1;
    });
  });

  return {
    data: traces,
    layout: {
      autosize: true,
      height: isExpanded ? 640 : 300,
      paper_bgcolor: "rgba(0,0,0,0)",
      plot_bgcolor: "rgba(0,0,0,0)",
      font: {
        family: "Space Grotesk, sans-serif",
        color: "#cbd5e1",
      },
      margin: { l: 52, r: 20, t: 12, b: 48 },
      xaxis: { title: chartType === "histogram" || chartType === "box" ? undefined : xColumns.join(", "), gridcolor: "rgba(255,255,255,0.05)", zerolinecolor: "rgba(255,255,255,0.1)" },
      yaxis: { title: chartType === "pie" ? undefined : (yColumns.length > 1 ? "Values" : yColumns[0]), gridcolor: "rgba(255,255,255,0.05)", zerolinecolor: "rgba(255,255,255,0.1)" },
      legend: { orientation: "h", font: { color: "#9ca3af" } },
      barmode: groupColumn && chartType === "bar" ? "group" : undefined,
      hovermode: "closest",
    },
  };
}

function getChartRows(chartType, xColumns, yColumns, groupColumn, aggregation) {
  const baseRows = getRenderableRows(state.rows, xColumns, yColumns, groupColumn, chartType);

  if (chartType === "bar" || chartType === "pie") {
    return summarizeRows(baseRows, xColumns, yColumns, groupColumn, aggregation);
  }

  return baseRows;
}

function getRenderableRows(rows, xColumns, yColumns, groupColumn, chartType) {
  const requiresX = chartType !== "histogram" && chartType !== "box";
  return rows.filter((row) => {
    const hasValidX = !requiresX || xColumns.every(col => {
      const xValue = row[col];
      return xValue !== null && xValue !== undefined && String(xValue).trim() !== "";
    });
    const hasValidY = yColumns.every(col => Number.isFinite(parseNumericValue(row[col])));
    const hasValidGroup = !groupColumn || (row[groupColumn] !== null && row[groupColumn] !== undefined && String(row[groupColumn]).trim() !== "");
    return hasValidX && hasValidY && hasValidGroup;
  });
}

function getChartDisplayValue(row, column) {
  if (state.dateColumns.includes(column)) {
    const timestamp = parseDateValue(row[column]);
    if (timestamp !== null) {
      return new Date(timestamp);
    }
  }

  return row[column];
}

function summarizeRows(rows, xColumns, yColumns, groupColumn, aggregation) {
  const groups = new Map();

  rows.forEach((row) => {
    const categoryKeys = xColumns.map(col => String(getChartDisplayValue(row, col) || "Unspecified"));
    const groupKey = groupColumn ? String(row[groupColumn] || "Unspecified") : "All";
    const key = `${groupKey}__${categoryKeys.join("__")}`;
    
    if (!groups.has(key)) {
      const initial = { __group: groupKey };
      xColumns.forEach((col, idx) => initial[col] = categoryKeys[idx]);
      yColumns.forEach(col => { initial[col] = 0; initial[`${col}__values`] = []; });
      groups.set(key, initial);
    }
    
    const current = groups.get(key);
    yColumns.forEach(col => {
      const val = parseNumericValue(row[col]);
      if (Number.isFinite(val)) {
        current[col] += val;
        current[`${col}__values`].push(val);
      }
    });
  });

  const result = [...groups.values()];
  
  result.forEach(row => {
    yColumns.forEach(col => {
      const vals = row[`${col}__values`];
      if (vals.length === 0) {
        row[col] = 0;
      } else if (aggregation === "avg") {
        row[col] = vals.reduce((a, b) => a + b, 0) / vals.length;
      } else if (aggregation === "count") {
        row[col] = vals.length;
      } else if (aggregation === "min") {
        row[col] = Math.min(...vals);
      } else if (aggregation === "max") {
        row[col] = Math.max(...vals);
      }
    });
  });

  return result;
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

function buildExportName(chartType, xColumns, yColumns) {
  const xName = xColumns.join("-").toLowerCase().replace(/\s+/g, "-");
  const yName = yColumns.join("-").toLowerCase().replace(/\s+/g, "-");
  return `${chartType}-${xName}-${yName}`;
}

function chartTitle(chartType, xColumns, yColumns) {
  const xName = xColumns.join(", ");
  const yName = yColumns.join(", ");
  const titles = {
    scatter: `Scatter Plot: ${xName} vs ${yName}`,
    line: `Line Plot: ${xName} vs ${yName}`,
    bar: `Bar Chart: ${xName} vs ${yName}`,
    histogram: `Histogram: ${yName}`,
    box: `Box Plot: ${yName}`,
    pie: `Pie Chart: ${xName} by ${yName}`,
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

function buildSelectorOptions(target, includeNone = false, config = {}) {
  const options = [];
  if (includeNone) {
    const placeholderLabels = {
      x: "Select X Axis",
      y: "Select Y Axis",
      group: "Select Group / Color",
    };
    options.push({ type: "option", value: "None", label: placeholderLabels[target] || "None" });
  }

  const excludedColumns = new Set((config.exclude || []).filter((value) => value && value !== "None"));

  const groups = [
    { key: "numeric", label: "Numeric Columns", columns: state.numericColumns },
    { key: "category", label: "Category Columns", columns: state.categoryColumns },
    { key: "date", label: "Date Columns", columns: state.dateColumns },
    { key: "id", label: "ID Columns", columns: state.idColumns },
    { key: "other", label: "Other Columns", columns: state.otherColumns },
  ];

  groups.forEach((group) => {
    let columns = group.columns.filter((column) => !excludedColumns.has(column));
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
      options: state.columns
        .filter((column) => !excludedColumns.has(column))
        .map((column) => ({ value: column, label: column })),
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
  if (state.numericColumns.includes(column)) {
    return "numeric";
  }
  if (state.dateColumns.includes(column)) {
    return "date";
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
    const numbers = values.map(parseNumericValue).filter((value) => Number.isFinite(value));
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
  state.rawRows = [];
  state.rows = [];
  state.sampleRows = [];
  state.normalizedDateRows = [];
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
  state.aiDashboardLoading = false;

  fileInput.value = "";
  fileStatus.textContent = "No file loaded yet.";
  fileStatus.classList.remove("is-loading");
  smartDashboardBtn.disabled = false;
  rowCount.textContent = "0";
  columnCount.textContent = "0";
  numericCount.textContent = "0";
  xAxisSelect.innerHTML = "";
  yAxisSelect.innerHTML = "";
  groupAxisSelect.innerHTML = "";
  columnMeta.innerHTML = "Load a file to inspect column types.";
  selectorGroups.innerHTML = "Load a file to browse grouped column options.";
  recommendations.innerHTML = "Load a file to see recommended plot combinations.";
  selectorWarning.hidden = true;
  selectorWarning.textContent = "";
  columnSearch.value = "";
  processingHint.hidden = true;
  updateFilterChips();
  previewHead.innerHTML = "";
  previewBody.innerHTML = "";
  clearCharts();
}

// UI helpers
function showChartLoading(message) {
  setChartAreaMessage(message, { loading: true });
}

function setChartAreaMessage(message, options = {}) {
  chartMessage.textContent = message;
  chartMessage.hidden = false;
  chartMessage.classList.toggle("is-loading", Boolean(options.loading));
}

function focusChartPanel() {
  if (window.innerWidth <= 920) {
    chartPanel?.scrollIntoView({ behavior: "smooth", block: "start" });
  }
}

function updateSelectorWarning() {
  const selectedCharts = [...chartOptions.querySelectorAll("input:checked")].map((input) => input.value);
  const feedback = getInlineSelectionFeedback(
    selectedCharts,
    getMultiSelectValues(xAxisSelect),
    getMultiSelectValues(yAxisSelect),
    groupAxisSelect.value === "None" ? null : groupAxisSelect.value
  );

  if (!feedback) {
    selectorWarning.hidden = true;
    selectorWarning.textContent = "";
    return;
  }

  showSelectorWarning(feedback);
}

function getInlineSelectionFeedback(selectedCharts, xColumns, yColumns, groupColumn) {
  if (!state.sampleRows.length || !selectedCharts.length) {
    return "";
  }

  const failedCharts = selectedCharts
    .map((chartType) => ({ chartType, ...canRenderChart(chartType, xColumns, yColumns, groupColumn) }))
    .filter((result) => !result.valid);

  if (!failedCharts.length) {
    return "";
  }

  if (failedCharts.length === 1) {
    return `${labelForChart(failedCharts[0].chartType)} chart: ${failedCharts[0].message}`;
  }

  return "Some selected chart types are not valid with the current column combination.";
}

function showSelectorWarning(message) {
  selectorWarning.textContent = message;
  selectorWarning.hidden = false;
}

function addChartFeedbackCard(chartType, message) {
  chartMessage.hidden = true;
  chartMessage.classList.remove("is-loading");

  const card = document.createElement("article");
  card.className = "chart-card is-feedback";

  const header = document.createElement("div");
  header.className = "chart-card-header";

  const title = document.createElement("h3");
  title.textContent = labelForChart(chartType);

  header.appendChild(title);

  const feedback = document.createElement("div");
  feedback.className = "chart-feedback";
  feedback.textContent = message || "This chart cannot be generated with the selected columns. Please choose different columns.";

  card.appendChild(header);
  card.appendChild(feedback);
  chartGrid.appendChild(card);
}

function waitForNextPaint() {
  return new Promise((resolve) => {
    requestAnimationFrame(() => resolve());
  });
}
