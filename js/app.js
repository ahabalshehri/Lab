const TARGETS = {
  stage1: 10,
  stage2: 10,
  stage3: 35,
  total: 60,
};

const INTERVALS = [
  {
    key: "stage1",
    label: "Doctor Order to Sample Collection",
    shortLabel: "Collection delay",
    target: TARGETS.stage1,
  },
  {
    key: "stage2",
    label: "Sample Collection to Lab Received",
    shortLabel: "Transport / receiving delay",
    target: TARGETS.stage2,
  },
  {
    key: "stage3",
    label: "Lab Received to Verified Result",
    shortLabel: "Lab processing / verification delay",
    target: TARGETS.stage3,
  },
];

let records = [];

const elements = {
  fileInput: document.getElementById("fileInput"),
  exportButton: document.getElementById("exportButton"),
  sampleButton: document.getElementById("sampleButton"),
  printButton: document.getElementById("printButton"),
  fromDate: document.getElementById("fromDate"),
  toDate: document.getElementById("toDate"),
  testFilter: document.getElementById("testFilter"),
  statusFilter: document.getElementById("statusFilter"),
  resetFiltersButton: document.getElementById("resetFiltersButton"),
  uploadNotice: document.getElementById("uploadNotice"),
  totalSamples: document.getElementById("totalSamples"),
  executiveTitle: document.getElementById("executiveTitle"),
  executiveSummary: document.getElementById("executiveSummary"),
  overallCompliance: document.getElementById("overallCompliance"),
  avgTat: document.getElementById("avgTat"),
  lateSamples: document.getElementById("lateSamples"),
  medianTat: document.getElementById("medianTat"),
  p90Tat: document.getElementById("p90Tat"),
  missingSamples: document.getElementById("missingSamples"),
  worstInterval: document.getElementById("worstInterval"),
  stage1Compliance: document.getElementById("stage1Compliance"),
  stage2Compliance: document.getElementById("stage2Compliance"),
  stage3Compliance: document.getElementById("stage3Compliance"),
  stage1Avg: document.getElementById("stage1Avg"),
  stage2Avg: document.getElementById("stage2Avg"),
  stage3Avg: document.getElementById("stage3Avg"),
  insightList: document.getElementById("insightList"),
  intervalCount: document.getElementById("intervalCount"),
  intervalRows: document.getElementById("intervalRows"),
  hourCount: document.getElementById("hourCount"),
  hourRows: document.getElementById("hourRows"),
  shiftCount: document.getElementById("shiftCount"),
  shiftRows: document.getElementById("shiftRows"),
  testCount: document.getElementById("testCount"),
  testRows: document.getElementById("testRows"),
  workflowCount: document.getElementById("workflowCount"),
  workflowRows: document.getElementById("workflowRows"),
};

elements.fileInput.addEventListener("change", handleFileUpload);
elements.exportButton.addEventListener("click", exportDelayedCsv);
elements.sampleButton.addEventListener("click", loadSampleData);
elements.printButton.addEventListener("click", () => window.print());
elements.fromDate.addEventListener("change", render);
elements.toDate.addEventListener("change", render);
elements.testFilter.addEventListener("change", render);
elements.statusFilter.addEventListener("change", render);
elements.resetFiltersButton.addEventListener("click", resetFilters);

render();

async function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) return;

  if (!window.XLSX) {
    alert("Excel reader library did not load. Check the internet connection and retry.");
    return;
  }

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  records = rows.map(normalizeRow).filter(Boolean);
  elements.uploadNotice.innerHTML = `<strong>Loaded ${records.length.toLocaleString("en")} records.</strong><span>${escapeHtml(file.name)}</span>`;
  updateTestOptions();
  setDefaultDates();
  render();
}

function normalizeRow(row, index) {
  const get = (...names) => {
    for (const name of names) {
      const match = Object.keys(row).find((key) => normalizeKey(key) === normalizeKey(name));
      if (match && row[match] !== "") return row[match];
    }
    return "";
  };

  const item = {
    id: String(
      get(
        "order_id",
        "orderid",
        "sample_id",
        "sampleid",
        "accession",
        "accession_number",
        "specimen_number",
        "patient_id",
        "mrn",
      ) || `ER-${index + 1}`,
    ).trim(),
    testName: String(get("test_name", "test", "testname", "profile", "assay", "investigation") || "Unknown").trim(),
    department: String(get("department", "section", "lab_section") || "ER Lab").trim(),
    priority: String(get("priority", "urgency", "class") || "ER").trim(),
    doctorOrderTime: parseDate(
      get(
        "doctor_order_time",
        "doctororder",
        "order_time",
        "ordered_time",
        "request_time",
        "requestdate",
        "order_date",
        "ordered_date",
      ),
    ),
    collectionTime: parseDate(
      get(
        "sample_collection_time",
        "collection_time",
        "collected_time",
        "draw_time",
        "sampledrawntime",
        "specimen_collection_time",
      ),
    ),
    labReceivedTime: parseDate(
      get(
        "lab_received_time",
        "received_time",
        "receive_time",
        "received_in_lab",
        "specimen_received_time",
        "labreceive",
      ),
    ),
    verifiedTime: parseDate(
      get(
        "verified_time",
        "verify_time",
        "verified_result_time",
        "released_time",
        "release_time",
        "result_verified_time",
        "authorization_time",
        "authorized_time",
        "result_time",
      ),
    ),
  };

  return withMetrics(item);
}

function withMetrics(item) {
  const stage1 = diffMinutes(item.doctorOrderTime, item.collectionTime);
  const stage2 = diffMinutes(item.collectionTime, item.labReceivedTime);
  const stage3 = diffMinutes(item.labReceivedTime, item.verifiedTime);
  const total = diffMinutes(item.doctorOrderTime, item.verifiedTime);
  const missing = [item.doctorOrderTime, item.collectionTime, item.labReceivedTime, item.verifiedTime].some((value) => !value);
  const status = missing ? "incomplete" : total <= TARGETS.total ? "ok" : "late";

  return {
    ...item,
    stage1,
    stage2,
    stage3,
    total,
    status,
    currentStage: currentStage(item),
    weakness: missing ? "Missing timestamps" : mainWeakness({ stage1, stage2, stage3 }),
  };
}

function currentStage(item) {
  if (!item.doctorOrderTime) return "Missing doctor order";
  if (!item.collectionTime) return "Waiting sample collection";
  if (!item.labReceivedTime) return "Waiting lab receiving";
  if (!item.verifiedTime) return "Waiting verification/release";
  return "Completed";
}

function mainWeakness(values) {
  const scored = INTERVALS.map((interval) => {
    const value = values[interval.key];
    const overTarget = Number.isFinite(value) ? value - interval.target : -Infinity;
    return { ...interval, value, overTarget };
  }).sort((a, b) => b.overTarget - a.overTarget);

  if (!Number.isFinite(scored[0].value)) return "Missing timestamps";
  if (scored[0].overTarget <= 0) {
    const largest = scored.slice().sort((a, b) => b.value - a.value)[0];
    return largest.shortLabel;
  }
  return scored[0].shortLabel;
}

function render() {
  const data = filteredRecords();
  renderKpis(data);
  renderExecutiveSummary(data);
  renderStages(data);
  renderInsights(data);
  renderIntervals(data);
  renderHours(data);
  renderShifts(data);
  renderTests(data);
  renderWorkflow(data);
}

function filteredRecords() {
  const from = elements.fromDate.value ? new Date(`${elements.fromDate.value}T00:00:00`) : null;
  const to = elements.toDate.value ? new Date(`${elements.toDate.value}T23:59:59`) : null;
  const test = elements.testFilter.value;
  const status = elements.statusFilter.value;

  return records.filter((record) => {
    const anchorDate = record.doctorOrderTime || record.collectionTime || record.labReceivedTime || record.verifiedTime;
    if (from && anchorDate && anchorDate < from) return false;
    if (to && anchorDate && anchorDate > to) return false;
    if (test && record.testName !== test) return false;
    if (status && record.status !== status) return false;
    return true;
  });
}

function renderKpis(data) {
  const complete = data.filter((record) => record.status !== "incomplete");
  const totals = complete.map((record) => record.total).filter(Number.isFinite);
  const worst = worstInterval(complete);
  elements.totalSamples.textContent = data.length.toLocaleString("en");
  elements.overallCompliance.textContent = percent(count(complete, (record) => record.status === "ok"), complete.length);
  elements.avgTat.textContent = formatMinutes(avg(totals));
  elements.lateSamples.textContent = count(data, (record) => record.status === "late").toLocaleString("en");
  elements.medianTat.textContent = formatMinutes(median(totals));
  elements.p90Tat.textContent = formatMinutes(percentile(totals, 90));
  elements.missingSamples.textContent = count(data, (record) => record.status === "incomplete").toLocaleString("en");
  elements.worstInterval.textContent = worst ? worst.shortLabel.replace(" / ", " /") : "--";
}

function renderExecutiveSummary(data) {
  const complete = data.filter((record) => record.status !== "incomplete");
  const delayed = complete.filter((record) => record.status === "late");
  if (!data.length) {
    elements.executiveTitle.textContent = "Awaiting ER data";
    elements.executiveSummary.textContent = "Upload the ER Excel file to identify whether the main delay is before collection, during transport/receiving, or inside laboratory processing and verification.";
    return;
  }

  const compliance = complianceValue(data);
  const totalAvg = avg(complete.map((record) => record.total));
  const weakness = topWeakness(delayed) || "No dominant delay";
  const worst = worstInterval(complete);
  const missing = count(data, (record) => record.status === "incomplete");
  const grade = compliance >= 90 ? "Strong ER TAT control" : compliance >= 75 ? "Moderate ER TAT control" : "ER TAT requires focused intervention";

  elements.executiveTitle.textContent = `${grade}: ${compliance.toFixed(1)}% within ${TARGETS.total} minutes`;
  elements.executiveSummary.textContent = [
    `${complete.length.toLocaleString("en")} complete records were analyzed with an average total TAT of ${formatMinutes(totalAvg)}.`,
    `${delayed.length.toLocaleString("en")} records exceeded the ${TARGETS.total}-minute target; the leading weakness is ${weakness}.`,
    worst ? `The largest average interval is ${worst.label} at ${formatMinutes(worst.avg)}.` : "",
    missing ? `${missing.toLocaleString("en")} records have missing timestamps and should be cleaned before final reporting.` : "Timestamp completeness is acceptable for the filtered data.",
  ].filter(Boolean).join(" ");
}

function renderStages(data) {
  renderStage("stage1", data, TARGETS.stage1, elements.stage1Compliance, elements.stage1Avg);
  renderStage("stage2", data, TARGETS.stage2, elements.stage2Compliance, elements.stage2Avg);
  renderStage("stage3", data, TARGETS.stage3, elements.stage3Compliance, elements.stage3Avg);
}

function renderStage(key, data, target, complianceElement, avgElement) {
  const values = data.map((record) => record[key]).filter(Number.isFinite);
  complianceElement.textContent = percent(values.filter((value) => value <= target).length, values.length);
  avgElement.textContent = `Average ${formatMinutes(avg(values))}`;
}

function renderInsights(data) {
  const complete = data.filter((record) => record.status !== "incomplete");
  const delayed = complete.filter((record) => record.status === "late");
  const strengths = INTERVALS.map((interval) => ({
    label: interval.shortLabel,
    avg: avg(complete.map((record) => record[interval.key])),
  })).filter((item) => Number.isFinite(item.avg)).sort((a, b) => a.avg - b.avg);
  const weakness = topWeakness(delayed);
  const bestHour = bestGroup(complete, (record) => hourLabel(record.doctorOrderTime));
  const worstHour = worstGroup(complete, (record) => hourLabel(record.doctorOrderTime));

  const rows = [
    insight("Main weakness", weakness || "--", `Most common main delay among records over ${TARGETS.total} minutes`),
    insight("Strongest interval", strengths[0] ? `${strengths[0].label} (${formatMinutes(strengths[0].avg)})` : "--", "Lowest average interval"),
    insight("Best order hour", bestHour || "--", `Highest ${TARGETS.total}-minute compliance`),
    insight("Weakest order hour", worstHour || "--", `Lowest ${TARGETS.total}-minute compliance`),
  ];

  elements.insightList.innerHTML = rows.join("");
}

function renderIntervals(data) {
  const delayed = data.filter((record) => record.status === "late");
  const rows = INTERVALS.map((interval) => {
    const values = data.map((record) => record[interval.key]).filter(Number.isFinite);
    const mainCauseCount = count(delayed, (record) => record.weakness === interval.shortLabel);
    return `<tr>
      <td>${escapeHtml(interval.label)}</td>
      <td class="mono">${formatMinutes(avg(values))}</td>
      <td class="mono">${formatMinutes(median(values))}</td>
      <td class="mono">${formatMinutes(max(values))}</td>
      <td class="mono">${mainCauseCount.toLocaleString("en")}</td>
      <td>${bar(percentValue(values.filter((value) => value <= interval.target).length, values.length))}</td>
    </tr>`;
  });

  elements.intervalCount.textContent = `${INTERVALS.length} intervals`;
  elements.intervalRows.innerHTML = rows.join("");
}

function renderHours(data) {
  const groups = groupBy(data.filter((record) => record.doctorOrderTime), (record) => hourLabel(record.doctorOrderTime));
  const rows = Object.entries(groups)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([hour, items]) => summaryRow(hour, items));

  elements.hourCount.textContent = `${Object.keys(groups).length.toLocaleString("en")} hours`;
  elements.hourRows.innerHTML = rows.join("") || emptyRow(5, "No hourly data");
}

function renderShifts(data) {
  const groups = groupBy(data.filter((record) => record.doctorOrderTime), (record) => shiftLabel(record.doctorOrderTime));
  const order = ["Night 00:00-07:59", "Morning 08:00-15:59", "Evening 16:00-23:59"];
  const rows = Object.entries(groups)
    .sort(([a], [b]) => order.indexOf(a) - order.indexOf(b))
    .map(([shift, items]) => summaryRow(shift, items));

  elements.shiftCount.textContent = `${Object.keys(groups).length.toLocaleString("en")} shifts`;
  elements.shiftRows.innerHTML = rows.join("") || emptyRow(5, "No shift data");
}

function renderTests(data) {
  const groups = groupBy(data, (record) => record.testName);
  const rows = Object.entries(groups)
    .sort(([, a], [, b]) => b.length - a.length)
    .slice(0, 20)
    .map(([test, items]) => summaryRow(test, items));

  elements.testCount.textContent = `${Object.keys(groups).length.toLocaleString("en")} tests`;
  elements.testRows.innerHTML = rows.join("") || emptyRow(5, "No test data");
}

function summaryRow(label, items) {
  const complete = items.filter((record) => record.status !== "incomplete");
  return `<tr>
    <td>${escapeHtml(label)}</td>
    <td class="mono">${items.length.toLocaleString("en")}</td>
    <td>${bar(complianceValue(items))}</td>
    <td class="mono">${formatMinutes(avg(complete.map((record) => record.total)))}</td>
    <td>${escapeHtml(topWeakness(complete.filter((record) => record.status === "late")) || "--")}</td>
  </tr>`;
}

function renderWorkflow(data) {
  const delayedFirst = data
    .slice()
    .sort((a, b) => {
      if (a.status === "late" && b.status !== "late") return -1;
      if (a.status !== "late" && b.status === "late") return 1;
      return (b.total || 0) - (a.total || 0);
    })
    .slice(0, 250);

  const rows = delayedFirst.map((record) => `<tr>
    <td class="mono">${escapeHtml(record.id)}</td>
    <td>${escapeHtml(record.testName)}</td>
    <td class="mono">${formatDateTime(record.doctorOrderTime)}</td>
    <td class="mono">${formatDateTime(record.collectionTime)}</td>
    <td class="mono">${formatDateTime(record.labReceivedTime)}</td>
    <td class="mono">${formatDateTime(record.verifiedTime)}</td>
    <td class="mono">${formatMinutes(record.total)}</td>
    <td>${escapeHtml(record.weakness)}</td>
    <td>${statusPill(record.status)}</td>
  </tr>`);

  elements.workflowCount.textContent = `${data.length.toLocaleString("en")} records`;
  elements.workflowRows.innerHTML = rows.join("") || emptyRow(9, "No ER records");
}

function updateTestOptions() {
  const current = elements.testFilter.value;
  const tests = [...new Set(records.map((record) => record.testName))].sort();
  elements.testFilter.innerHTML = `<option value="">All tests</option>${tests
    .map((test) => `<option value="${escapeHtml(test)}">${escapeHtml(test)}</option>`)
    .join("")}`;
  elements.testFilter.value = tests.includes(current) ? current : "";
}

function setDefaultDates() {
  const dates = records.map((record) => record.doctorOrderTime).filter(Boolean).sort((a, b) => a - b);
  if (!dates.length) return;
  elements.fromDate.value = toInputDate(dates[0]);
  elements.toDate.value = toInputDate(dates[dates.length - 1]);
}

function resetFilters() {
  elements.fromDate.value = "";
  elements.toDate.value = "";
  elements.testFilter.value = "";
  elements.statusFilter.value = "";
  render();
}

function loadSampleData() {
  const now = new Date();
  const tests = ["CBC", "Troponin", "Chemistry Panel", "Coagulation", "Blood Gas", "Lactate"];
  records = Array.from({ length: 180 }, (_, index) => {
    const base = new Date(now);
    base.setDate(now.getDate() - Math.floor(index / 26));
    base.setHours(index % 24, (index * 11) % 60, 0, 0);
    const pattern = index % 6;
    const stage1 = pattern === 0 ? 20 + (index % 10) : 4 + (index % 9);
    const stage2 = pattern === 1 ? 18 + (index % 8) : 4 + ((index * 2) % 8);
    const stage3 = pattern === 2 ? 48 + (index % 28) : 22 + ((index * 5) % 24);
    const doctorOrderTime = base;
    const collectionTime = addMinutes(doctorOrderTime, stage1);
    const labReceivedTime = addMinutes(collectionTime, stage2);
    const verifiedTime = addMinutes(labReceivedTime, stage3);
    return withMetrics({
      id: `ER-${String(260000 + index).padStart(6, "0")}`,
      testName: tests[index % tests.length],
      department: index % 2 === 0 ? "Chemistry" : "Hematology",
      priority: "ER",
      doctorOrderTime,
      collectionTime,
      labReceivedTime,
      verifiedTime,
    });
  });
  elements.uploadNotice.innerHTML = `<strong>Sample ER data loaded.</strong><span>Upload Excel to replace it with real data.</span>`;
  updateTestOptions();
  setDefaultDates();
  render();
}

function exportDelayedCsv() {
  const delayed = filteredRecords().filter((record) => record.status === "late");
  if (!delayed.length) {
    alert("No delayed records to export for the current filters.");
    return;
  }

  const headers = [
    "id",
    "test_name",
    "doctor_order_time",
    "collection_time",
    "lab_received_time",
    "verified_time",
    "stage1_minutes",
    "stage2_minutes",
    "stage3_minutes",
    "total_tat_minutes",
    "main_weakness",
  ];
  const rows = delayed.map((record) => [
    record.id,
    record.testName,
    isoValue(record.doctorOrderTime),
    isoValue(record.collectionTime),
    isoValue(record.labReceivedTime),
    isoValue(record.verifiedTime),
    numericValue(record.stage1),
    numericValue(record.stage2),
    numericValue(record.stage3),
    numericValue(record.total),
    record.weakness,
  ]);

  const csv = [headers, ...rows].map((row) => row.map(csvCell).join(",")).join("\n");
  const blob = new Blob([csv], { type: "text/csv;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url;
  link.download = `er-delayed-records-${toInputDate(new Date())}.csv`;
  link.click();
  URL.revokeObjectURL(url);
}

function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number" && window.XLSX) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) return new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H, parsed.M, parsed.S);
  }
  const normalized = String(value).trim();
  const date = new Date(normalized);
  return Number.isNaN(date.getTime()) ? null : date;
}

function normalizeKey(value) {
  return String(value).toLowerCase().replace(/[\s_\-./():]+/g, "");
}

function diffMinutes(start, end) {
  if (!start || !end) return null;
  const value = (end.getTime() - start.getTime()) / 60000;
  return value >= 0 ? value : null;
}

function addMinutes(date, minutes) {
  return new Date(date.getTime() + minutes * 60000);
}

function avg(values) {
  const nums = values.filter(Number.isFinite);
  if (!nums.length) return null;
  return nums.reduce((sum, value) => sum + value, 0) / nums.length;
}

function median(values) {
  const nums = values.filter(Number.isFinite).sort((a, b) => a - b);
  if (!nums.length) return null;
  const mid = Math.floor(nums.length / 2);
  return nums.length % 2 ? nums[mid] : (nums[mid - 1] + nums[mid]) / 2;
}

function percentile(values, percentileValue) {
  const nums = values.filter(Number.isFinite).sort((a, b) => a - b);
  if (!nums.length) return null;
  const index = (percentileValue / 100) * (nums.length - 1);
  const lower = Math.floor(index);
  const upper = Math.ceil(index);
  if (lower === upper) return nums[lower];
  return nums[lower] + (nums[upper] - nums[lower]) * (index - lower);
}

function max(values) {
  const nums = values.filter(Number.isFinite);
  return nums.length ? Math.max(...nums) : null;
}

function count(items, predicate) {
  return items.filter(predicate).length;
}

function percent(numerator, denominator) {
  if (!denominator) return "--%";
  return `${percentValue(numerator, denominator).toFixed(1)}%`;
}

function percentValue(numerator, denominator) {
  if (!denominator) return 0;
  return (numerator / denominator) * 100;
}

function complianceValue(items) {
  const complete = items.filter((record) => record.status !== "incomplete");
  return percentValue(count(complete, (record) => record.status === "ok"), complete.length);
}

function complianceColor(value) {
  if (value >= 90) return "var(--green)";
  if (value >= 75) return "var(--amber)";
  return "var(--red)";
}

function worstInterval(items) {
  const ranked = INTERVALS.map((interval) => ({
    ...interval,
    avg: avg(items.map((record) => record[interval.key])),
  })).filter((interval) => Number.isFinite(interval.avg)).sort((a, b) => b.avg - a.avg);
  return ranked[0] || null;
}

function formatMinutes(value) {
  if (!Number.isFinite(value)) return "--";
  return `${value.toFixed(1)} min`;
}

function formatDateTime(value) {
  if (!value) return "--";
  return value.toLocaleString("en-GB", {
    year: "2-digit",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function statusPill(status) {
  const labels = {
    ok: "Within target",
    late: `Over ${TARGETS.total}`,
    incomplete: "Missing data",
  };
  return `<span class="status ${status}">${labels[status] || status}</span>`;
}

function groupBy(items, keyFn) {
  return items.reduce((groups, item) => {
    const key = keyFn(item) || "Unknown";
    groups[key] ||= [];
    groups[key].push(item);
    return groups;
  }, {});
}

function topWeakness(items) {
  const groups = groupBy(items, (record) => record.weakness);
  const sorted = Object.entries(groups).sort(([, a], [, b]) => b.length - a.length);
  return sorted[0] ? `${sorted[0][0]} (${sorted[0][1].toLocaleString("en")})` : "";
}

function bestGroup(items, keyFn) {
  return rankedGroup(items, keyFn, "best");
}

function worstGroup(items, keyFn) {
  return rankedGroup(items, keyFn, "worst");
}

function rankedGroup(items, keyFn, mode) {
  const groups = Object.entries(groupBy(items, keyFn)).filter(([, rows]) => rows.length >= 3);
  if (!groups.length) return "";
  groups.sort(([, a], [, b]) => {
    const diff = complianceValue(a) - complianceValue(b);
    return mode === "best" ? -diff : diff;
  });
  const [label, rows] = groups[0];
  return `${label} (${complianceValue(rows).toFixed(1)}%)`;
}

function hourLabel(date) {
  if (!date) return "Unknown";
  return `${String(date.getHours()).padStart(2, "0")}:00`;
}

function shiftLabel(date) {
  if (!date) return "Unknown";
  const hour = date.getHours();
  if (hour < 8) return "Night 00:00-07:59";
  if (hour < 16) return "Morning 08:00-15:59";
  return "Evening 16:00-23:59";
}

function isoValue(value) {
  return value ? value.toISOString() : "";
}

function numericValue(value) {
  return Number.isFinite(value) ? value.toFixed(1) : "";
}

function csvCell(value) {
  const text = String(value ?? "");
  return `"${text.replaceAll('"', '""')}"`;
}

function toInputDate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function bar(value) {
  const clamped = Math.max(0, Math.min(100, value));
  return `<div class="bar">
    <span>${clamped.toFixed(1)}%</span>
    <div class="bar-track"><div class="bar-fill" style="width:${clamped}%;background:${complianceColor(clamped)}"></div></div>
  </div>`;
}

function insight(title, value, detail) {
  return `<div class="insight-item">
    <span>${escapeHtml(title)}</span>
    <strong>${escapeHtml(value)}</strong>
    <small>${escapeHtml(detail)}</small>
  </div>`;
}

function emptyRow(colspan, text) {
  return `<tr><td colspan="${colspan}" style="text-align:center;color:var(--muted);padding:28px">${escapeHtml(text)}</td></tr>`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
