const SAMPLE_FILE = "Исходные данные/Мониторинг NPS за 2025 год.xlsx";
const MONTHS = [
  "январь", "февраль", "март", "апрель", "май", "июнь",
  "июль", "август", "сентябрь", "октябрь", "ноябрь", "декабрь",
];
const SHORT_MONTHS = ["янв", "фев", "мар", "апр", "май", "июн", "июл", "авг", "сен", "окт", "ноя", "дек"];
const COLORS = ["#4468f2", "#4cc6a4", "#f0a45d", "#e87998", "#8d73dd", "#58a7d9", "#9bb2cf", "#d08cce"];
const HOVER_COLOR = "#c13b70";
const TRAINER_COLORS = new Map();

const state = {
  rows: [],
  filteredRows: [],
  missingRows: [],
  filteredMissingRows: [],
  summaryTrends: null,
  sourceName: "",
};

const $ = (id) => document.getElementById(id);
const filters = {
  month: $("monthFilter"),
  trainer: $("trainerFilter"),
  training: $("trainingFilter"),
  company: $("companyFilter"),
};

document.addEventListener("DOMContentLoaded", () => {
  const shell = document.querySelector(".app-shell");
  const sidebar = document.querySelector(".sidebar");
  $("fileInput").addEventListener("change", handleFileUpload);
  $("loadSampleButton").addEventListener("click", loadSample);
  $("resetFiltersButton").addEventListener("click", resetFilters);
  $("topResetFiltersButton").addEventListener("click", resetFilters);
  $("toggleMissingReportsButton").addEventListener("click", toggleMissingReports);
  document.querySelectorAll("[data-kpi-filter]").forEach((button) => {
    button.addEventListener("click", () => applyKpiFilter(button.dataset.kpiFilter, button.dataset.inputId));
  });
  document.querySelectorAll(".kpi-filter-card").forEach((card) => {
    card.addEventListener("click", (event) => {
      if (!event.target.closest(".kpi-popover") && !event.target.closest(".kpi-selection-clear")) openKpiCard(card);
    });
  });
  document.addEventListener("click", (event) => {
    if (!event.target.closest(".kpi-filter-card")) closeKpiCards();
  });
  document.querySelectorAll(".kpi-popover input").forEach((input) => {
    input.addEventListener("input", () => renderKpiSuggestions(input));
    input.addEventListener("focus", () => renderKpiSuggestions(input));
    input.addEventListener("keydown", (event) => {
      if (event.key === "Enter") applyKpiFilter(input.closest(".kpi-popover").querySelector("button").dataset.kpiFilter, input.id);
    });
  });
  $("exportButton").addEventListener("click", exportDashboard);
  $("sidebarToggle").addEventListener("click", () => shell.classList.remove("sidebar-collapsed"));
  $("sidebarToggle").addEventListener("mouseenter", () => shell.classList.remove("sidebar-collapsed"));
  sidebar.addEventListener("mouseleave", () => shell.classList.add("sidebar-collapsed"));
  Object.values(filters).forEach((select) => select.addEventListener("change", render));
  document.querySelectorAll("[data-filter-sync]").forEach((select) => {
    select.addEventListener("change", () => {
      filters[select.dataset.filterSync].value = select.value;
      render();
    });
  });
  window.addEventListener("resize", debounce(renderCharts, 120));
  loadSample();
});

async function loadSample() {
  setStatus("Подключаю образец...");
  try {
    const response = await fetch(SAMPLE_FILE);
    if (!response.ok) throw new Error("Файл образца не найден");
    await loadWorkbook(await response.arrayBuffer(), "Мониторинг NPS за 2025 год.xlsx");
  } catch (error) {
    setStatus("Не удалось открыть образец. Выберите Excel-файл вручную.", true);
  }
}

async function handleFileUpload(event) {
   event.preventDefault();
   $("fileInput").value = "";
   alert("Данная функция в данный момент недоступна");
 }

async function loadWorkbook(buffer, sourceName) {
  if (!window.XLSX) throw new Error("Библиотека XLSX не загружена");
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const { rows, missingRows } = parseWorkbook(workbook);
  if (!rows.length) throw new Error("В файле не найдено записей NPS");
  state.rows = rows;
  state.missingRows = missingRows;
  state.summaryTrends = parseSummaryTrends(workbook);
  state.sourceName = sourceName;
  populateFilters();
  render();
  const trainers = unique(rows.map((row) => row.trainer)).length;
  setStatus(`Загружено: ${rows.length} отчётов NPS · ${missingRows.length} без отчёта · ${trainers} тренеров`);
  $("footerSource").textContent = `Источник: ${sourceName}`;
  $("updatedAt").textContent = new Intl.DateTimeFormat("ru-RU", {
    day: "2-digit", month: "2-digit", year: "numeric",
  }).format(new Date());
}

function parseWorkbook(workbook) {
  const preferredSheet = findSheet(workbook, "статистика по тренерам");
  const preferred = preferredSheet ? parseNormalizedSheet(preferredSheet) : { rows: [], missingRows: [] };
  const monthly = parseMonthlySheets(workbook);
  return monthly.rows.length >= preferred.rows.length ? monthly : preferred;
}

function parseSummaryTrends(workbook) {
  const sheet = findSheet(workbook, "графика по среднему нпс");
  if (!sheet) return null;
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
  const trainings = matrix.find((row) => row.some((cell) => normalize(cell) === "тренинги"));
  const factories = matrix.find((row) => row.some((cell) => normalize(cell) === "фабрики процессов"));
  if (!trainings || !factories) return null;
  const trainingsStart = trainings.findIndex((cell) => normalize(cell) === "тренинги") + 1;
  const factoriesStart = factories.findIndex((cell) => normalize(cell) === "фабрики процессов") + 1;
  return {
    trainings: trainings.slice(trainingsStart, trainingsStart + 12).map(toTrendValue),
    factories: factories.slice(factoriesStart, factoriesStart + 12).map(toTrendValue),
  };
}

function toTrendValue(value) {
  const number = Number(value);
  return Number.isFinite(number) ? number : null;
}

function findSheet(workbook, needle) {
  const name = workbook.SheetNames.find((item) => normalize(item).includes(needle));
  return name ? workbook.Sheets[name] : null;
}

function parseNormalizedSheet(sheet) {
  const matrix = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, raw: true });
  const headerIndex = matrix.findIndex((row) => {
    const keys = row.map(normalize);
    return keys.includes("месяц") && keys.includes("фио тренера") && keys.includes("тренинг")
      && keys.includes("предприятие") && keys.includes("nps");
  });
  if (headerIndex < 0) return { rows: [], missingRows: [] };
  const headers = matrix[headerIndex].map(normalize);
  const index = {
    month: headers.indexOf("месяц"),
    day: headers.indexOf("день"),
    trainer: headers.indexOf("фио тренера"),
    training: headers.indexOf("тренинг"),
    company: headers.indexOf("предприятие"),
    nps: headers.indexOf("nps"),
  };
  const rows = [];
  const missingRows = [];
  matrix.slice(headerIndex + 1).forEach((row) => {
    const args = [row[index.month], row[index.day], row[index.trainer], row[index.training], row[index.company], row[index.nps]];
    const record = toRecord(...args);
    if (record) rows.push(record);
    else {
      const missing = toMissingRecord(...args);
      if (missing) missingRows.push(missing);
    }
  });
  return { rows, missingRows };
}

function parseMonthlySheets(workbook) {
  const records = [];
  const missingRows = [];
  workbook.SheetNames.forEach((sheetName) => {
    const month = MONTHS.indexOf(normalize(sheetName)) + 1;
    if (!month) return;
    const matrix = XLSX.utils.sheet_to_json(workbook.Sheets[sheetName], { header: 1, defval: null, raw: true });
    matrix.forEach((row) => {
      const textCells = row.filter((cell) => typeof cell === "string");
      if (textCells.length < 3) return;
      const trainerIndex = row.findIndex((cell) => typeof cell === "string" && /\s/.test(cell) && !/^(ооо|ао)\b/i.test(cell));
      if (trainerIndex < 0) return;
      if (normalize(row[trainerIndex]) === "фио тренера") return;
      const day = Number(row[trainerIndex - 1]);
      const nps = row.slice(trainerIndex + 3).reverse().find((cell) => typeof cell === "number" && cell >= 0 && cell <= 1);
      const args = [month, day, row[trainerIndex], row[trainerIndex + 1], row[trainerIndex + 2], nps];
      const record = toRecord(...args);
      if (record) records.push(record);
      else {
        const missing = toMissingRecord(...args);
        if (missing) missingRows.push(missing);
      }
    });
  });
  return { rows: records, missingRows };
}

function toRecord(month, day, trainer, training, company, nps) {
  if (nps === null || nps === undefined || nps === "") return null;
  const monthNumber = Number(month);
  const dayNumber = Number(day);
  let npsNumber = Number(nps);
  if (!Number.isFinite(monthNumber) || monthNumber < 1 || monthNumber > 12) return null;
  if (!trainer || !training || !company || !Number.isFinite(npsNumber)) return null;
  if (npsNumber <= 1.5) npsNumber *= 100;
  return {
    month: monthNumber,
    day: Number.isFinite(dayNumber) ? dayNumber : 1,
    trainer: String(trainer).trim(),
    training: normalizeTrainingName(training),
    company: normalizeCompanyName(company),
    nps: npsNumber,
  };
}

function toMissingRecord(month, day, trainer, training, company, nps) {
  if (nps !== null && nps !== undefined && nps !== "") return null;
  const monthNumber = Number(month);
  const dayNumber = Number(day);
  if (!Number.isFinite(monthNumber) || monthNumber < 1 || monthNumber > 12) return null;
  if (!trainer || !training || !company) return null;
  return {
    month: monthNumber,
    day: Number.isFinite(dayNumber) ? dayNumber : 1,
    trainer: String(trainer).trim(),
    training: normalizeTrainingName(training),
    company: normalizeCompanyName(company),
  };
}

function normalizeTrainingName(value) {
  const training = String(value).trim();
  return normalize(training) === "мрп" ? "Мет.реш.проблем" : training;
}

function normalizeCompanyName(value) {
  const company = String(value).trim();
  return normalize(company).replace(/\s+/g, "") === 'ооо"дизайнинвест"' ? 'ООО "ДизайнИнвест"' : company;
}

function populateFilters() {
  const values = {
    month: unique(state.rows.map((row) => row.month)).sort((a, b) => a - b),
    trainer: unique(state.rows.map((row) => row.trainer)).sort(localeSort),
    training: unique(state.rows.map((row) => row.training)).sort(localeSort),
    company: unique(state.rows.map((row) => row.company)).sort(localeSort),
  };
  setOptions(filters.month, values.month, (value) => MONTHS[value - 1]);
  setOptions(filters.trainer, values.trainer);
  setOptions(filters.training, values.training);
  setOptions(filters.company, values.company);
  syncPanelFilters();
}

function setOptions(select, values, label = (value) => value) {
  select.innerHTML = '<option value="">Все</option>';
  values.forEach((value) => {
    const option = document.createElement("option");
    option.value = value;
    option.textContent = label(value);
    select.append(option);
  });
}

function resetFilters() {
  Object.values(filters).forEach((select) => {
    select.value = "";
  });
  closeKpiCards();
  render();
}

function toggleFilter(type, value) {
  const select = filters[type];
  if (!select) return;
  select.value = select.value === value ? "" : value;
  render();
}

function syncPanelFilters() {
  document.querySelectorAll("[data-filter-sync]").forEach((select) => {
    const source = filters[select.dataset.filterSync];
    select.innerHTML = source.innerHTML;
    select.value = source.value;
  });
  $("kpiTrainingSearch").value = filters.training.value;
  $("kpiTrainerSearch").value = filters.trainer.value;
  $("kpiCompanySearch").value = filters.company.value;
  renderKpiSelection("training", "kpiTrainingSelection");
  renderKpiSelection("trainer", "kpiTrainerSelection");
  renderKpiSelection("company", "kpiCompanySelection");
}

function openKpiCard(card) {
  document.querySelectorAll(".kpi-filter-card.is-open").forEach((item) => {
    if (item !== card) item.classList.remove("is-open");
  });
  card.classList.add("is-open");
  document.querySelector(".app-shell").classList.add("kpi-focus-mode");
  const input = card.querySelector(".kpi-popover input");
  renderKpiSuggestions(input);
  requestAnimationFrame(() => input.focus());
}

function closeKpiCards() {
  document.querySelectorAll(".kpi-filter-card.is-open").forEach((card) => card.classList.remove("is-open"));
  document.querySelector(".app-shell").classList.remove("kpi-focus-mode");
}

function renderKpiSuggestions(input) {
  const button = input.closest(".kpi-popover").querySelector("[data-kpi-filter]");
  const type = button.dataset.kpiFilter;
  const query = normalize(input.value);
  const values = [...filters[type].options]
    .map((option) => option.value)
    .filter((value) => value && (!query || normalize(value).includes(query)))
    .slice(0, 12);
  input.nextElementSibling.innerHTML = values.length
    ? values.map((value) => `<button type="button" class="kpi-suggestion" data-value="${escapeHtml(value)}">${escapeHtml(value)}</button>`).join("")
    : '<span class="kpi-suggestion-empty">Совпадений нет</span>';
  input.nextElementSibling.querySelectorAll(".kpi-suggestion").forEach((item) => {
    item.addEventListener("click", () => {
      input.value = item.dataset.value;
      applyKpiFilter(type, input.id);
    });
  });
}

function renderKpiSelection(type, containerId) {
  const container = $(containerId);
  const value = filters[type].value;
  container.innerHTML = value
    ? `<span title="${escapeHtml(value)}">${escapeHtml(value)}</span><button class="kpi-selection-clear" type="button" aria-label="Сбросить фильтр">×</button>`
    : "";
  container.querySelector(".kpi-selection-clear")?.addEventListener("click", (event) => {
    event.stopPropagation();
    filters[type].value = "";
    render();
  });
}

function applyKpiFilter(type, inputId) {
  const query = $(inputId).value.trim();
  if (!query) {
    filters[type].value = "";
    render();
    return;
  }
  const options = [...filters[type].options].map((option) => option.value).filter(Boolean);
  const exact = options.find((value) => normalize(value) === normalize(query));
  const partial = options.find((value) => normalize(value).includes(normalize(query)));
  filters[type].value = exact || partial || "";
  $(inputId).value = filters[type].value || query;
  closeKpiCards();
  render();
}

function render() {
  state.filteredRows = state.rows.filter(matchesFilters);
  state.filteredMissingRows = state.missingRows.filter(matchesFilters);
  const rows = state.filteredRows;
  const activityRows = getFilteredActivityRows();
  $("avgTrainingNps").textContent = formatTrendKpi("trainings");
  $("avgFactoryNps").textContent = formatTrendKpi("factories");
  $("eventsCount").textContent = activityRows.length.toLocaleString("ru-RU");
  $("trainersCount").textContent = unique(activityRows.map((row) => row.trainer)).length;
  $("companiesCount").textContent = unique(activityRows.map((row) => row.company)).length;
  $("reportSubtitle").textContent = buildSubtitle(rows);
  syncPanelFilters();
  renderCharts();
  renderMissingReports();
}

function getFilteredActivityRows() {
  return [...state.filteredRows, ...state.filteredMissingRows];
}

function matchesFilters(row, ignoredType = "") {
  return (
    (ignoredType === "month" || !filters.month.value || row.month === Number(filters.month.value))
    && (ignoredType === "trainer" || !filters.trainer.value || row.trainer === filters.trainer.value)
    && (ignoredType === "training" || !filters.training.value || row.training === filters.training.value)
    && (ignoredType === "company" || !filters.company.value || row.company === filters.company.value)
  );
}

function formatTrendKpi(type) {
  const values = state.summaryTrends?.[type] || [];
  if (!values.length) return "—";
  const month = Number(filters.month.value);
  const value = month ? values[month - 1] : average(values.filter((item) => item !== null));
  return Number.isFinite(value) ? `${Math.round(value)}%` : "—";
}

function renderCharts() {
  if (!state.rows.length) return;
  renderTrendChart();
  renderTrainerChart();
  renderTrainerCountChart();
  renderTrainingChart();
  renderTrainerTrendChart();
  renderTrainingNpsChart();
}

function renderMissingReports() {
  const rows = state.filteredMissingRows;
  $("missingReportsCount").textContent = `${rows.length} пустых записей`;
  const data = [...groupBy(rows, (row) => `${row.trainer}|||${row.training}|||${row.company}`).entries()]
    .map(([name, items]) => ({ name, value: items.length }))
    .sort((a, b) => b.value - a.value || localeSort(a.name, b.name));
  $("missingPairs").innerHTML = data.length
    ? data.map((item) => {
      const [trainer, training, company] = item.name.split("|||");
      return `<div class="missing-row">
        ${filterButton("trainer", trainer)}
        ${filterButton("training", training)}
        ${filterButton("company", company)}
        <strong>${item.value}</strong>
      </div>`;
    }).join("")
    : '<span class="missing-empty">Нет пропусков</span>';
  $("missingPairs").querySelectorAll("[data-filter-type]").forEach((button) => {
    button.addEventListener("click", () => toggleFilter(button.dataset.filterType, button.dataset.filterValue));
  });
}

function toggleMissingReports() {
  const pairs = $("missingPairs");
  const button = $("toggleMissingReportsButton");
  const isHidden = pairs.classList.toggle("is-hidden");
  button.textContent = isHidden ? "Показать результаты" : "Скрыть результаты";
  button.setAttribute("aria-expanded", String(!isHidden));
}

function filterButton(type, value) {
  const selected = filters[type].value === value ? " is-active" : "";
  return `<button class="missing-filter missing-filter-${type}${selected}" type="button" data-filter-type="${type}" data-filter-value="${escapeHtml(value)}">${escapeHtml(value)}</button>`;
}

function renderTrendChart() {
  const groups = groupBy(state.filteredRows, (row) => row.month);
  const filteredTrainings = MONTHS.map((_, index) => {
    const rows = groups.get(index + 1) || [];
    return rows.length ? average(rows.map((row) => row.nps)) : null;
  });
  const hasDetailFilter = filters.trainer.value || filters.training.value || filters.company.value;
  const trainings = maskSelectedMonth(state.summaryTrends && !hasDetailFilter ? state.summaryTrends.trainings : filteredTrainings);
  const factories = maskSelectedMonth(state.summaryTrends ? state.summaryTrends.factories : []);
  drawComparisonBarChart($("trendChart"), [
    { name: "NPS тренингов", values: trainings, color: "#4468f2" },
    { name: "NPS фабрик процессов", values: factories, color: "#4cc6a4" },
  ], 75, { filterType: "month" });
}

function maskSelectedMonth(values) {
  const month = Number(filters.month.value);
  return month ? values.map((value, index) => index === month - 1 ? value : null) : values;
}

function renderTrainerChart() {
  const groups = groupBy(state.filteredRows, (row) => row.trainer);
  const data = [...groups.entries()]
    .map(([name, rows]) => ({ name, value: average(rows.map((row) => row.nps)), count: rows.length }))
    .sort((a, b) => b.value - a.value || b.count - a.count);
  ensureTrainerColors(data.map((item) => item.name));
  drawBarChart($("trainerChart"), data, { filterType: "trainer", colorFn: (item) => getTrainerColor(item.name) });
}

function renderTrainerCountChart() {
  const groups = groupBy(getFilteredActivityRows(), (row) => row.trainer);
  const data = [...groups.entries()]
    .map(([name, rows]) => ({ name, value: rows.length }))
    .sort((a, b) => b.value - a.value || localeSort(a.name, b.name));
  ensureTrainerColors(data.map((item) => item.name));
  drawCountBarChart($("trainerCountChart"), data, {
    filterType: "trainer",
    colorFn: (item) => getTrainerColor(item.name),
    labelFn: shortenName,
    pad: { top: 4, right: 28, bottom: 2, left: 84 },
    barGapFn: (length) => length > 8 ? 3 : length > 6 ? 6 : 11,
    maxBarHeight: 26,
  });
}

function renderTrainingChart() {
  const groups = groupBy(getFilteredActivityRows(), (row) => row.training);
  const data = [...groups.entries()]
    .map(([name, rows]) => ({ name, value: rows.length }))
    .sort((a, b) => b.value - a.value);
  drawCountBarChart($("trainingChart"), data, { filterType: "training" });
}

function renderTrainingNpsChart() {
  const groups = groupBy(state.filteredRows, (row) => row.training);
  const data = [...groups.entries()]
    .map(([name, rows]) => ({ name, value: average(rows.map((row) => row.nps)), count: rows.length }))
    .sort((a, b) => b.value - a.value || b.count - a.count);
  drawBarChart($("trainingNpsChart"), data, { filterType: "training" });
}

function renderTrainerTrendChart() {
  const groups = groupBy(state.filteredRows, (row) => row.trainer);
  const data = [...groups.entries()]
    .map(([name, rows]) => {
      const byMonth = groupBy(rows, (row) => row.month);
      return {
        name,
        count: rows.length,
        values: MONTHS.map((_, index) => {
          const monthRows = byMonth.get(index + 1) || [];
          return monthRows.length ? average(monthRows.map((row) => row.nps)) : null;
        }),
      };
    })
    .sort((a, b) => b.count - a.count || localeSort(a.name, b.name));
  ensureTrainerColors(data.map((item) => item.name));
  drawGroupedBarChart($("trainerTrendChart"), data, 75, { filterType: "trainer", legendId: "trainerTrendLegend" });
  $("trainerTrendLegend").innerHTML = data.map((item, index) =>
    `<span class="clickable-legend" data-filter-value="${escapeHtml(item.name)}"><i style="background:${getTrainerColor(item.name)}"></i>${escapeHtml(shortenName(item.name))}</span>`
  ).join("");
  bindLegendFilters($("trainerTrendLegend"), "trainer");
}

function ensureTrainerColors(names) {
  names.forEach((name) => {
    if (!TRAINER_COLORS.has(name)) TRAINER_COLORS.set(name, COLORS[TRAINER_COLORS.size % COLORS.length]);
  });
}

function getTrainerColor(name) {
  return TRAINER_COLORS.get(name) || COLORS[0];
}

function bindLegendFilters(container, type) {
  container.querySelectorAll("[data-filter-value]").forEach((item) => {
    item.addEventListener("click", () => toggleFilter(type, item.dataset.filterValue));
  });
}

function drawComparisonBarChart(canvas, series, target, options = {}) {
  const { ctx, width, height } = prepareCanvas(canvas);
  const pad = { top: 22, right: 15, bottom: 24, left: 34 };
  const chartW = width - pad.left - pad.right;
  const chartH = height - pad.top - pad.bottom;
  ctx.font = "10px Inter, sans-serif";
  ctx.textAlign = "right";
  ctx.textBaseline = "middle";
  [0, 25, 50, 75, 100].forEach((tick) => {
    const y = pad.top + chartH - (tick / 100) * chartH;
    ctx.strokeStyle = "#edf0f5";
    ctx.beginPath();
    ctx.moveTo(pad.left, y);
    ctx.lineTo(width - pad.right, y);
    ctx.stroke();
    ctx.fillStyle = "#9aa5b6";
    ctx.fillText(tick, pad.left - 8, y);
  });
  ctx.textAlign = "center";
  SHORT_MONTHS.forEach((month, index) => {
    const x = pad.left + (index + 0.5) * (chartW / 12);
    ctx.fillStyle = "#9aa5b6";
    ctx.fillText(month, x, height - 9);
  });
  const groupW = chartW / 12;
  const innerW = groupW * 0.66;
  const barW = innerW / series.length;
  const regions = [];
  const draw = (hovered = null) => {
    ctx.clearRect(0, 0, width, height);
    ctx.font = "10px Inter, sans-serif";
    ctx.textAlign = "right";
    ctx.textBaseline = "middle";
    [0, 25, 50, 75, 100].forEach((tick) => {
      const y = pad.top + chartH - (tick / 100) * chartH;
      ctx.strokeStyle = "#edf0f5";
      ctx.beginPath();
      ctx.moveTo(pad.left, y);
      ctx.lineTo(width - pad.right, y);
      ctx.stroke();
      ctx.fillStyle = "#9aa5b6";
      ctx.fillText(tick, pad.left - 8, y);
    });
    ctx.textAlign = "center";
    SHORT_MONTHS.forEach((month, index) => {
      const x = pad.left + (index + 0.5) * groupW;
      ctx.fillStyle = "#9aa5b6";
      ctx.fillText(month, x, height - 9);
    });
    regions.length = 0;
    series.forEach((item, seriesIndex) => {
      item.values.forEach((value, monthIndex) => {
        if (value === null) return;
        const groupX = pad.left + monthIndex * groupW + (groupW - innerW) / 2;
        const x = groupX + seriesIndex * barW;
        const y = pad.top + chartH - (value / 100) * chartH;
        const width = Math.max(2, barW - 2);
        const isHovered = hovered?.seriesIndex === seriesIndex && hovered?.monthIndex === monthIndex;
        ctx.fillStyle = isHovered ? HOVER_COLOR : item.color;
        ctx.beginPath();
        ctx.rect(x, y, width, pad.top + chartH - y);
        ctx.fill();
        ctx.fillStyle = isHovered ? HOVER_COLOR : item.color;
        ctx.font = "700 9px Inter, sans-serif";
        ctx.textAlign = "center";
        ctx.textBaseline = "bottom";
        ctx.fillText(`${Math.round(value)}`, x + width / 2, y - 3);
        regions.push({ x, y, width, height: pad.top + chartH - y, monthIndex, seriesIndex });
      });
    });
    drawTargetLine(ctx, width, pad, chartH, target);
  };
  const getRegion = (event) => {
    const rect = canvas.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    return regions.find((region) => x >= region.x - 3 && x <= region.x + region.width + 3 && y >= region.y - 5 && y <= region.y + region.height);
  };
  draw();
  canvas.onclick = options.filterType ? (event) => {
    const region = getRegion(event);
    if (region) toggleFilter(options.filterType, String(region.monthIndex + 1));
  } : null;
  canvas.onmousemove = options.filterType ? (event) => {
    const region = getRegion(event);
    canvas.style.cursor = region ? "pointer" : "default";
    draw(region);
  } : null;
  canvas.onmouseleave = options.filterType ? () => draw() : null;
}

function drawGroupedBarChart(canvas, series, target, options = {}) {
  const { ctx, width, height } = prepareCanvas(canvas);
  const pad = { top: 14, right: 15, bottom: 28, left: 34 };
  const chartW = width - pad.left - pad.right;
  const chartH = height - pad.top - pad.bottom;
  const groupW = chartW / 12;
  const innerW = groupW * 0.76;
  const barW = Math.max(2, innerW / Math.max(series.length, 1));
  const regions = [];
  const legend = options.legendId ? $(options.legendId) : null;
  const highlightLegend = (seriesIndex = -1) => {
    if (!legend) return;
    legend.querySelectorAll("[data-filter-value]").forEach((item) => {
      item.classList.toggle("is-highlighted", item.dataset.filterValue === series[seriesIndex]?.name);
    });
  };
  const draw = (hoveredSeries = -1) => {
    ctx.clearRect(0, 0, width, height);
    ctx.font = "10px Inter, sans-serif";
    ctx.textAlign = "right";
    ctx.textBaseline = "middle";
    [0, 25, 50, 75, 100].forEach((tick) => {
      const y = pad.top + chartH - (tick / 100) * chartH;
      ctx.strokeStyle = "#edf0f5";
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(pad.left, y);
      ctx.lineTo(width - pad.right, y);
      ctx.stroke();
      ctx.fillStyle = "#9aa5b6";
      ctx.fillText(tick, pad.left - 8, y);
    });
    ctx.textAlign = "center";
    SHORT_MONTHS.forEach((month, index) => {
      const x = pad.left + (index + 0.5) * groupW;
      ctx.fillStyle = "#9aa5b6";
      ctx.fillText(month, x, height - 9);
    });
    regions.length = 0;
    series.forEach((item, seriesIndex) => {
      const color = hoveredSeries === seriesIndex ? HOVER_COLOR : getTrainerColor(item.name);
      item.values.forEach((value, monthIndex) => {
        if (value === null) return;
        const groupX = pad.left + monthIndex * groupW + (groupW - innerW) / 2;
        const x = groupX + seriesIndex * barW;
        const y = pad.top + chartH - (value / 100) * chartH;
        const width = Math.max(1.5, barW - 1);
        ctx.fillStyle = color;
        ctx.beginPath();
        ctx.rect(x, y, width, pad.top + chartH - y);
        ctx.fill();
        ctx.fillStyle = color;
        ctx.font = "700 7px Inter, sans-serif";
        ctx.textAlign = "center";
        ctx.textBaseline = "bottom";
        ctx.fillText(`${Math.round(value)}`, x + width / 2, y - 2 - (seriesIndex % 3) * 7);
        regions.push({ x, y, width, height: pad.top + chartH - y, seriesIndex });
      });
    });
    drawTargetLine(ctx, width, pad, chartH, target);
  };
  const getSeriesIndex = (event) => {
    const rect = canvas.getBoundingClientRect();
    const x = event.clientX - rect.left;
    const y = event.clientY - rect.top;
    const matches = regions
      .map((region) => ({
        ...region,
        distance: Math.abs(x - (region.x + region.width / 2)),
      }))
      .filter((region) => (
        x >= region.x - 4
        && x <= region.x + region.width + 4
        && y >= region.y - 6
        && y <= region.y + region.height
      ))
      .sort((a, b) => a.distance - b.distance);
    return matches[0]?.seriesIndex ?? -1;
  };
  draw();
  canvas.onclick = options.filterType ? (event) => {
    const index = getSeriesIndex(event);
    if (index >= 0) toggleFilter(options.filterType, series[index].name);
  } : null;
  canvas.onmousemove = options.filterType ? (event) => {
    const index = getSeriesIndex(event);
    canvas.style.cursor = index >= 0 ? "pointer" : "default";
    highlightLegend(index);
    draw(index);
  } : null;
  canvas.onmouseleave = options.filterType ? () => {
    highlightLegend();
    draw();
  } : null;
}

function drawTargetLine(ctx, width, pad, chartH, target) {
  const targetY = pad.top + chartH - (target / 100) * chartH;
  ctx.save();
  ctx.setLineDash([6, 4]);
  ctx.strokeStyle = "#f0a45d";
  ctx.lineWidth = 2;
  ctx.beginPath();
  ctx.moveTo(pad.left, targetY);
  ctx.lineTo(width - pad.right, targetY);
  ctx.stroke();
  ctx.restore();
}

function drawBarChart(canvas, data, options = {}) {
  const { ctx, width, height } = prepareCanvas(canvas);
  if (!data.length) return drawEmpty(ctx, width, height);
  const pad = { top: 4, right: 28, bottom: 2, left: 84 };
  const barGap = data.length > 8 ? 3 : data.length > 6 ? 6 : 11;
  const barH = Math.min(26, Math.max(6, (height - pad.top - pad.bottom - barGap * (data.length - 1)) / data.length));
  const totalHeight = barH * data.length + barGap * (data.length - 1);
  const startY = pad.top + Math.max(0, (height - pad.top - pad.bottom - totalHeight) / 2);
  const draw = (hoveredIndex = -1) => {
    ctx.clearRect(0, 0, width, height);
    ctx.font = `${data.length > 8 ? 7 : 9}px Inter, sans-serif`;
    ctx.textBaseline = "middle";
    data.forEach((item, index) => {
    const y = startY + index * (barH + barGap);
    const valueW = ((width - pad.left - pad.right) * item.value) / 100;
    const isHovered = index === hoveredIndex;
    if (isHovered) {
      ctx.fillStyle = "#fff0f5";
      roundRect(ctx, 0, y - 4, width, barH + 8, 8);
      ctx.fill();
    }
    ctx.fillStyle = "#f0f3f8";
    roundRect(ctx, pad.left, y, width - pad.left - pad.right, barH, 6);
    ctx.fill();
    ctx.fillStyle = isHovered ? HOVER_COLOR : options.colorFn?.(item, index) || COLORS[index % COLORS.length];
    roundRect(ctx, pad.left, y, valueW, barH, 6);
    ctx.fill();
    ctx.fillStyle = isHovered ? HOVER_COLOR : "#67758b";
    ctx.textAlign = "right";
    ctx.fillText(shortenName(item.name), pad.left - 8, y + barH / 2);
    if (isHovered) {
      const labelWidth = ctx.measureText(shortenName(item.name)).width;
      ctx.strokeStyle = HOVER_COLOR;
      ctx.lineWidth = 1;
      ctx.beginPath();
      ctx.moveTo(pad.left - 8 - labelWidth, y + barH / 2 + 6);
      ctx.lineTo(pad.left - 8, y + barH / 2 + 6);
      ctx.stroke();
    }
    ctx.fillStyle = isHovered ? HOVER_COLOR : "#344157";
    ctx.textAlign = "left";
    ctx.fillText(`${Math.round(item.value)}%`, pad.left + valueW + 7, y + barH / 2);
    });
  };
  const getHoveredIndex = (event) => {
    const rect = canvas.getBoundingClientRect();
    const y = event.clientY - rect.top;
    return data.findIndex((_, itemIndex) => {
      const rowY = startY + itemIndex * (barH + barGap);
      return y >= rowY - 4 && y <= rowY + barH + 4;
    });
  };
  draw();
  canvas.onclick = options.filterType ? (event) => {
    const index = getHoveredIndex(event);
    if (index < 0) return;
    toggleFilter(options.filterType, data[index].name);
  } : null;
  canvas.onmousemove = options.filterType ? (event) => {
    const index = getHoveredIndex(event);
    canvas.style.cursor = index >= 0 ? "pointer" : "default";
    draw(index);
  } : null;
  canvas.onmouseleave = options.filterType ? () => draw() : null;
  canvas.style.cursor = options.filterType ? "pointer" : "default";
}

function drawCountBarChart(canvas, data, options = {}) {
  const { ctx, width, height } = prepareCanvas(canvas);
  if (!data.length) return drawEmpty(ctx, width, height);
  const pad = options.pad || { top: 4, right: 24, bottom: 2, left: 90 };
  const maxValue = Math.max(...data.map((item) => item.value), 1);
  const barGap = options.barGapFn?.(data.length) ?? (data.length > 8 ? 4 : 7);
  const barH = Math.min(options.maxBarHeight || 22, Math.max(7, (height - pad.top - pad.bottom - barGap * (data.length - 1)) / data.length));
  const totalHeight = barH * data.length + barGap * (data.length - 1);
  const startY = pad.top + Math.max(0, (height - pad.top - pad.bottom - totalHeight) / 2);
  const draw = (hoveredIndex = -1) => {
    ctx.clearRect(0, 0, width, height);
    ctx.font = `${data.length > 8 ? 7 : 9}px Inter, sans-serif`;
    ctx.textBaseline = "middle";
    data.forEach((item, index) => {
      const y = startY + index * (barH + barGap);
      const valueW = ((width - pad.left - pad.right) * item.value) / maxValue;
      const isHovered = index === hoveredIndex;
      if (isHovered) {
        ctx.fillStyle = "#fff0f5";
        roundRect(ctx, 0, y - 4, width, barH + 8, 8);
        ctx.fill();
      }
      ctx.fillStyle = "#f0f3f8";
      roundRect(ctx, pad.left, y, width - pad.left - pad.right, barH, 6);
      ctx.fill();
      ctx.fillStyle = isHovered ? HOVER_COLOR : options.colorFn?.(item, index) || COLORS[index % COLORS.length];
      roundRect(ctx, pad.left, y, valueW, barH, 6);
      ctx.fill();
      ctx.fillStyle = isHovered ? HOVER_COLOR : "#67758b";
      ctx.textAlign = "right";
      const rawLabel = options.labelFn?.(item.name) || item.name;
      const label = rawLabel.length > 18 ? `${rawLabel.slice(0, 16)}...` : rawLabel;
      ctx.fillText(label, pad.left - 8, y + barH / 2);
      if (isHovered) {
        const labelWidth = ctx.measureText(label).width;
        ctx.strokeStyle = HOVER_COLOR;
        ctx.beginPath();
        ctx.moveTo(pad.left - 8 - labelWidth, y + barH / 2 + 6);
        ctx.lineTo(pad.left - 8, y + barH / 2 + 6);
        ctx.stroke();
      }
      ctx.fillStyle = isHovered ? HOVER_COLOR : "#344157";
      ctx.textAlign = "left";
      ctx.fillText(item.value, pad.left + valueW + 7, y + barH / 2);
    });
  };
  const getHoveredIndex = (event) => {
    const rect = canvas.getBoundingClientRect();
    const y = event.clientY - rect.top;
    return data.findIndex((_, index) => {
      const rowY = startY + index * (barH + barGap);
      return y >= rowY - 4 && y <= rowY + barH + 4;
    });
  };
  draw();
  canvas.onclick = options.filterType ? (event) => {
    const index = getHoveredIndex(event);
    if (index >= 0) toggleFilter(options.filterType, data[index].name);
  } : null;
  canvas.onmousemove = options.filterType ? (event) => {
    const index = getHoveredIndex(event);
    canvas.style.cursor = index >= 0 ? "pointer" : "default";
    draw(index);
  } : null;
  canvas.onmouseleave = options.filterType ? () => draw() : null;
}

function drawDonutChart(canvas, data, options = {}) {
  const { ctx, width, height } = prepareCanvas(canvas);
  if (!data.length) return drawEmpty(ctx, width, height);
  const total = data.reduce((sum, item) => sum + item.value, 0);
  const x = width / 2;
  const y = height / 2;
  const radius = Math.min(width, height) * 0.42;
  const sectors = [];
  let angle = -Math.PI / 2;
  data.forEach((item) => {
    const next = angle + (item.value / total) * Math.PI * 2;
    sectors.push({ start: angle, end: next });
    angle = next;
  });
  const draw = (hoveredIndex = -1) => {
    ctx.clearRect(0, 0, width, height);
    sectors.forEach((sector, index) => {
      ctx.lineWidth = index === hoveredIndex ? 22 : 17;
      ctx.strokeStyle = COLORS[index % COLORS.length];
      const gap = Math.min(0.018, (sector.end - sector.start) * 0.25);
      ctx.beginPath();
      ctx.arc(x, y, radius, sector.start + gap, sector.end - gap);
      ctx.stroke();
      const middle = sector.start + (sector.end - sector.start) / 2;
      const labelX = x + Math.cos(middle) * radius;
      const labelY = y + Math.sin(middle) * radius;
      ctx.fillStyle = "#fff";
      ctx.font = "800 10px Inter, sans-serif";
      ctx.textAlign = "center";
      ctx.textBaseline = "middle";
      ctx.fillText(data[index].value, labelX, labelY);
    });
  };
  const getSectorIndex = (event) => {
    const rect = canvas.getBoundingClientRect();
    const dx = event.clientX - rect.left - x;
    const dy = event.clientY - rect.top - y;
    const distance = Math.sqrt(dx * dx + dy * dy);
    if (distance < radius - 16 || distance > radius + 16) return -1;
    let angle = Math.atan2(dy, dx);
    if (angle < -Math.PI / 2) angle += Math.PI * 2;
    return sectors.findIndex((sector) => angle >= sector.start && angle <= sector.end);
  };
  draw();
  canvas.onclick = options.filterType ? (event) => {
    const index = getSectorIndex(event);
    if (index >= 0) toggleFilter(options.filterType, data[index].name);
  } : null;
  canvas.onmousemove = options.filterType ? (event) => {
    const index = getSectorIndex(event);
    canvas.style.cursor = index >= 0 ? "pointer" : "default";
    draw(index);
  } : null;
  canvas.onmouseleave = options.filterType ? () => draw() : null;
}

function prepareCanvas(canvas) {
  const ratio = window.devicePixelRatio || 1;
  const rect = canvas.getBoundingClientRect();
  canvas.width = Math.round(rect.width * ratio);
  canvas.height = Math.round(rect.height * ratio);
  const ctx = canvas.getContext("2d");
  ctx.scale(ratio, ratio);
  ctx.clearRect(0, 0, rect.width, rect.height);
  return { ctx, width: rect.width, height: rect.height };
}

function drawEmpty(ctx, width, height) {
  ctx.fillStyle = "#9aa5b6";
  ctx.font = "12px Inter, sans-serif";
  ctx.textAlign = "center";
  ctx.fillText("Нет данных для отображения", width / 2, height / 2);
}

async function exportDashboard() {
  if (!window.html2canvas) {
    setStatus("Модуль экспорта не загружен.", true);
    return;
  }
  const button = $("exportButton");
  button.disabled = true;
  button.querySelector("span").textContent = "Готовлю PNG...";
  try {
    const canvas = await html2canvas($("dashboardView"), {
      scale: 2,
      backgroundColor: "#f7f9fc",
      useCORS: true,
      ignoreElements: (element) => element.id === "topResetFiltersButton",
    });
    const link = document.createElement("a");
    link.download = `NPS-dashboard-${new Date().toISOString().slice(0, 10)}.png`;
    link.href = canvas.toDataURL("image/png");
    link.click();
    setStatus("PNG сохранён. Его можно пересылать.");
  } catch (error) {
    console.error(error);
    setStatus("Не удалось сохранить PNG.", true);
  } finally {
    button.disabled = false;
    button.querySelector("span").textContent = "Сохранить как PNG";
  }
}

function buildSubtitle(rows) {
  const selected = [];
  if (filters.month.value) selected.push(`${MONTHS[Number(filters.month.value) - 1]} 2025 года`);
  if (filters.trainer.value) selected.push(filters.trainer.value);
  if (filters.training.value) selected.push(filters.training.value);
  if (filters.company.value) selected.push(filters.company.value);
  const missing = state.filteredMissingRows.length;
  return `${selected.length ? selected.join(" · ") : "Все данные за 2025 год"} · ${rows.length} отчётов NPS${missing ? ` · ${missing} без отчёта` : ""}`;
}

function groupBy(rows, keyFn) {
  return rows.reduce((map, row) => {
    const key = keyFn(row);
    if (!map.has(key)) map.set(key, []);
    map.get(key).push(row);
    return map;
  }, new Map());
}

function average(values) {
  return values.length ? values.reduce((sum, value) => sum + value, 0) / values.length : 0;
}

function unique(values) {
  return [...new Set(values.filter(Boolean))];
}

function normalize(value) {
  return String(value ?? "").trim().toLowerCase();
}

function localeSort(a, b) {
  return String(a).localeCompare(String(b), "ru");
}

function shortenName(name) {
  const parts = name.split(/\s+/);
  return parts.length > 1 ? `${parts[0]} ${parts[1][0]}.` : name;
}

function pad(value) {
  return String(value).padStart(2, "0");
}

function roundRect(ctx, x, y, width, height, radius) {
  const r = Math.min(radius, width / 2, height / 2);
  ctx.beginPath();
  ctx.roundRect(x, y, width, height, r);
}

function escapeHtml(value) {
  return String(value).replace(/[&<>"']/g, (char) => ({
    "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#039;",
  }[char]));
}

function setStatus(message, isError = false) {
  $("fileStatus").textContent = message;
  $("fileStatus").style.color = isError ? "#ffb7b7" : "";
}

function debounce(fn, delay) {
  let timer;
  return (...args) => {
    clearTimeout(timer);
    timer = setTimeout(() => fn(...args), delay);
  };
}
