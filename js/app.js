const TARGETS = {
  stage1: 10,
  stage2: 10,
  stage3: 40,
  total: 60,
};

let records = [];

const elements = {
  fileInput: document.getElementById("fileInput"),
  sampleButton: document.getElementById("sampleButton"),
  printButton: document.getElementById("printButton"),
  fromDate: document.getElementById("fromDate"),
  toDate: document.getElementById("toDate"),
  hospitalFilter: document.getElementById("hospitalFilter"),
  statusFilter: document.getElementById("statusFilter"),
  resetFiltersButton: document.getElementById("resetFiltersButton"),
  uploadNotice: document.getElementById("uploadNotice"),
  totalSamples: document.getElementById("totalSamples"),
  overallCompliance: document.getElementById("overallCompliance"),
  avgTat: document.getElementById("avgTat"),
  lateSamples: document.getElementById("lateSamples"),
  stage1Compliance: document.getElementById("stage1Compliance"),
  stage2Compliance: document.getElementById("stage2Compliance"),
  stage3Compliance: document.getElementById("stage3Compliance"),
  stage1Avg: document.getElementById("stage1Avg"),
  stage2Avg: document.getElementById("stage2Avg"),
  stage3Avg: document.getElementById("stage3Avg"),
  weeklyList: document.getElementById("weeklyList"),
  hospitalCount: document.getElementById("hospitalCount"),
  comparisonRows: document.getElementById("comparisonRows"),
  workflowCount: document.getElementById("workflowCount"),
  workflowRows: document.getElementById("workflowRows"),
};

elements.fileInput.addEventListener("change", handleFileUpload);
elements.sampleButton.addEventListener("click", loadSampleData);
elements.printButton.addEventListener("click", () => window.print());
elements.fromDate.addEventListener("change", render);
elements.toDate.addEventListener("change", render);
elements.hospitalFilter.addEventListener("change", render);
elements.statusFilter.addEventListener("change", render);
elements.resetFiltersButton.addEventListener("click", resetFilters);

render();

async function handleFileUpload(event) {
  const file = event.target.files[0];
  if (!file) {
    return;
  }

  if (!window.XLSX) {
    alert("تعذر تحميل مكتبة قراءة ملفات Excel. تأكد من الاتصال بالإنترنت ثم أعد المحاولة.");
    return;
  }

  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array", cellDates: true });
  const sheet = workbook.Sheets[workbook.SheetNames[0]];
  const rows = XLSX.utils.sheet_to_json(sheet, { defval: "" });
  records = rows.map(normalizeRow).filter(Boolean);
  elements.uploadNotice.innerHTML = `<strong>تم تحميل ${records.length.toLocaleString("en")} سجل.</strong><span>${escapeHtml(file.name)}</span>`;
  updateHospitalOptions();
  setDefaultDates();
  render();
}

function normalizeRow(row, index) {
  const get = (...names) => {
    for (const name of names) {
      const match = Object.keys(row).find((key) => normalizeKey(key) === normalizeKey(name));
      if (match && row[match] !== "") {
        return row[match];
      }
    }
    return "";
  };

  const hospital = String(get("hospital", "facility", "المستشفى", "Hospital Name") || "غير محدد").trim();
  const sampleId = String(get("sample_id", "sample", "order_id", "accession", "رقم العينة") || `SAMPLE-${index + 1}`).trim();
  const orderTime = parseDate(get("order_time", "order", "وقت الطلب", "Order Date"));
  const collectionTime = parseDate(get("collection_time", "collection", "sample_collection", "وقت السحب"));
  const labReceivedTime = parseDate(get("lab_received_time", "received_time", "lab_received", "وقت استقبال المختبر"));
  const resultTime = parseDate(get("result_time", "result", "result_date", "وقت النتيجة"));

  const item = {
    hospital,
    sampleId,
    testName: String(get("test_name", "test", "الفحص") || "غير محدد").trim(),
    department: String(get("department", "section", "القسم") || "المختبر").trim(),
    priority: String(get("priority", "الأولوية") || "Routine").trim(),
    orderTime,
    collectionTime,
    labReceivedTime,
    resultTime,
  };

  return withMetrics(item);
}

function withMetrics(item) {
  const stage1 = diffMinutes(item.orderTime, item.collectionTime);
  const stage2 = diffMinutes(item.collectionTime, item.labReceivedTime);
  const stage3 = diffMinutes(item.labReceivedTime, item.resultTime);
  const total = diffMinutes(item.orderTime, item.resultTime);
  const incomplete = [item.orderTime, item.collectionTime, item.labReceivedTime, item.resultTime].some((value) => !value);
  const status = incomplete ? "incomplete" : total <= TARGETS.total ? "ok" : "late";

  return {
    ...item,
    stage1,
    stage2,
    stage3,
    total,
    status,
    currentStage: currentStage(item),
  };
}

function currentStage(item) {
  if (!item.orderTime) return "لم يبدأ";
  if (!item.collectionTime) return "بانتظار سحب العينة";
  if (!item.labReceivedTime) return "بانتظار استقبال المختبر";
  if (!item.resultTime) return "بانتظار ظهور النتيجة";
  return "مكتمل";
}

function render() {
  const data = filteredRecords();
  renderKpis(data);
  renderStages(data);
  renderWeekly(data);
  renderComparison(data);
  renderWorkflow(data);
}

function filteredRecords() {
  const from = elements.fromDate.value ? new Date(`${elements.fromDate.value}T00:00:00`) : null;
  const to = elements.toDate.value ? new Date(`${elements.toDate.value}T23:59:59`) : null;
  const hospital = elements.hospitalFilter.value;
  const status = elements.statusFilter.value;

  return records.filter((record) => {
    const anchorDate = record.orderTime || record.collectionTime || record.labReceivedTime || record.resultTime;
    if (from && anchorDate && anchorDate < from) return false;
    if (to && anchorDate && anchorDate > to) return false;
    if (hospital && record.hospital !== hospital) return false;
    if (status && record.status !== status) return false;
    return true;
  });
}

function renderKpis(data) {
  const completed = data.filter((record) => record.status !== "incomplete");
  elements.totalSamples.textContent = data.length.toLocaleString("en");
  elements.overallCompliance.textContent = percent(count(completed, (record) => record.status === "ok"), completed.length);
  elements.avgTat.textContent = formatMinutes(avg(completed.map((record) => record.total)));
  elements.lateSamples.textContent = count(data, (record) => record.status === "late").toLocaleString("en");
}

function renderStages(data) {
  renderStage("stage1", data, TARGETS.stage1, elements.stage1Compliance, elements.stage1Avg);
  renderStage("stage2", data, TARGETS.stage2, elements.stage2Compliance, elements.stage2Avg);
  renderStage("stage3", data, TARGETS.stage3, elements.stage3Compliance, elements.stage3Avg);
}

function renderStage(key, data, target, complianceElement, avgElement) {
  const values = data.map((record) => record[key]).filter((value) => Number.isFinite(value));
  const compliant = values.filter((value) => value <= target).length;
  complianceElement.textContent = percent(compliant, values.length);
  avgElement.textContent = `متوسط ${formatMinutes(avg(values))}`;
}

function renderWeekly(data) {
  const groups = groupBy(data, (record) => weekKey(record.orderTime || record.collectionTime || record.resultTime));
  const rows = Object.entries(groups)
    .sort(([a], [b]) => a.localeCompare(b))
    .map(([week, items]) => {
      const completed = items.filter((record) => record.status !== "incomplete");
      return `<div class="weekly-item">
        <div>
          <strong>${escapeHtml(week)}</strong>
          <span>${items.length.toLocaleString("en")} عينة</span>
        </div>
        <strong>${percent(count(completed, (record) => record.status === "ok"), completed.length)}</strong>
      </div>`;
    });

  elements.weeklyList.innerHTML = rows.join("") || `<div class="weekly-item"><span>لا توجد بيانات أسبوعية</span><strong>--</strong></div>`;
}

function renderComparison(data) {
  const groups = groupBy(data, (record) => record.hospital);
  const rows = Object.entries(groups)
    .sort(([, a], [, b]) => complianceValue(b) - complianceValue(a))
    .map(([hospital, items]) => {
      const completed = items.filter((record) => record.status !== "incomplete");
      const compliance = complianceValue(items);
      return `<tr>
        <td>${escapeHtml(hospital)}</td>
        <td class="mono">${items.length.toLocaleString("en")}</td>
        <td>
          <div class="bar">
            <span>${percent(count(completed, (record) => record.status === "ok"), completed.length)}</span>
            <div class="bar-track"><div class="bar-fill" style="width:${Math.max(0, compliance)}%;background:${complianceColor(compliance)}"></div></div>
          </div>
        </td>
        <td class="mono">${formatMinutes(avg(completed.map((record) => record.total)))}</td>
        <td class="mono">${count(items, (record) => record.status === "late")}</td>
      </tr>`;
    });

  elements.hospitalCount.textContent = `${Object.keys(groups).length.toLocaleString("en")} مستشفى`;
  elements.comparisonRows.innerHTML = rows.join("") || emptyRow(5, "لا توجد بيانات مقارنة");
}

function renderWorkflow(data) {
  const rows = data
    .slice()
    .sort((a, b) => (b.orderTime || 0) - (a.orderTime || 0))
    .slice(0, 200)
    .map((record) => `<tr>
      <td class="mono">${escapeHtml(record.sampleId)}</td>
      <td>${escapeHtml(record.hospital)}</td>
      <td>${escapeHtml(record.testName)}</td>
      <td>${escapeHtml(record.department)}</td>
      <td>${escapeHtml(record.priority)}</td>
      <td>${escapeHtml(record.currentStage)}</td>
      <td class="mono">${formatMinutes(record.total)}</td>
      <td>${statusPill(record.status)}</td>
    </tr>`);

  elements.workflowCount.textContent = `${data.length.toLocaleString("en")} عينة`;
  elements.workflowRows.innerHTML = rows.join("") || emptyRow(8, "لا توجد بيانات سير عمل");
}

function updateHospitalOptions() {
  const current = elements.hospitalFilter.value;
  const hospitals = [...new Set(records.map((record) => record.hospital))].sort();
  elements.hospitalFilter.innerHTML = `<option value="">كل المستشفيات</option>${hospitals
    .map((hospital) => `<option value="${escapeHtml(hospital)}">${escapeHtml(hospital)}</option>`)
    .join("")}`;
  elements.hospitalFilter.value = hospitals.includes(current) ? current : "";
}

function setDefaultDates() {
  const dates = records.map((record) => record.orderTime).filter(Boolean).sort((a, b) => a - b);
  if (!dates.length) {
    return;
  }
  elements.fromDate.value = toInputDate(dates[0]);
  elements.toDate.value = toInputDate(dates[dates.length - 1]);
}

function resetFilters() {
  elements.fromDate.value = "";
  elements.toDate.value = "";
  elements.hospitalFilter.value = "";
  elements.statusFilter.value = "";
  render();
}

function loadSampleData() {
  const now = new Date();
  const hospitals = ["مستشفى الملك فهد", "مستشفى أحد", "مستشفى الميقات", "مستشفى ينبع العام"];
  const tests = ["CBC", "Chemistry Panel", "Troponin", "Coagulation", "Blood Gas"];
  records = Array.from({ length: 120 }, (_, index) => {
    const base = new Date(now);
    base.setDate(now.getDate() - Math.floor(index / 18));
    base.setHours(index % 24, (index * 7) % 60, 0, 0);
    const stage1 = 4 + (index % 15);
    const stage2 = 5 + ((index * 2) % 13);
    const stage3 = 24 + ((index * 5) % 48);
    const orderTime = base;
    const collectionTime = addMinutes(orderTime, stage1);
    const labReceivedTime = addMinutes(collectionTime, stage2);
    const resultTime = addMinutes(labReceivedTime, stage3);
    return withMetrics({
      hospital: hospitals[index % hospitals.length],
      sampleId: `ER-${String(240000 + index).padStart(6, "0")}`,
      testName: tests[index % tests.length],
      department: index % 3 === 0 ? "Hematology" : index % 3 === 1 ? "Chemistry" : "Blood Bank",
      priority: index % 5 === 0 ? "STAT" : "Routine",
      orderTime,
      collectionTime,
      labReceivedTime,
      resultTime,
    });
  });
  elements.uploadNotice.innerHTML = `<strong>تم تحميل بيانات تجريبية.</strong><span>يمكنك استبدالها برفع ملف Excel.</span>`;
  updateHospitalOptions();
  setDefaultDates();
  render();
}

function parseDate(value) {
  if (!value) return null;
  if (value instanceof Date && !Number.isNaN(value.getTime())) return value;
  if (typeof value === "number" && window.XLSX) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      return new Date(parsed.y, parsed.m - 1, parsed.d, parsed.H, parsed.M, parsed.S);
    }
  }
  const date = new Date(String(value).trim());
  return Number.isNaN(date.getTime()) ? null : date;
}

function normalizeKey(value) {
  return String(value).toLowerCase().replace(/[\s_\-./]+/g, "");
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
  const nums = values.filter((value) => Number.isFinite(value));
  if (!nums.length) return null;
  return nums.reduce((sum, value) => sum + value, 0) / nums.length;
}

function count(items, predicate) {
  return items.filter(predicate).length;
}

function percent(numerator, denominator) {
  if (!denominator) return "--%";
  return `${((numerator / denominator) * 100).toFixed(1)}%`;
}

function complianceValue(items) {
  const completed = items.filter((record) => record.status !== "incomplete");
  if (!completed.length) return -1;
  return (count(completed, (record) => record.status === "ok") / completed.length) * 100;
}

function complianceColor(value) {
  if (value >= 90) return "var(--green)";
  if (value >= 75) return "var(--amber)";
  return "var(--red)";
}

function formatMinutes(value) {
  if (!Number.isFinite(value)) return "--";
  return `${value.toFixed(1)} د`;
}

function statusPill(status) {
  const labels = {
    ok: "ضمن المستهدف",
    late: "متأخر",
    incomplete: "بيانات ناقصة",
  };
  return `<span class="status ${status}">${labels[status] || status}</span>`;
}

function groupBy(items, keyFn) {
  return items.reduce((groups, item) => {
    const key = keyFn(item) || "غير محدد";
    groups[key] ||= [];
    groups[key].push(item);
    return groups;
  }, {});
}

function weekKey(date) {
  if (!date) return "غير محدد";
  const first = new Date(date);
  first.setHours(0, 0, 0, 0);
  first.setDate(first.getDate() - first.getDay());
  const last = new Date(first);
  last.setDate(first.getDate() + 6);
  return `${toInputDate(first)} إلى ${toInputDate(last)}`;
}

function toInputDate(date) {
  const y = date.getFullYear();
  const m = String(date.getMonth() + 1).padStart(2, "0");
  const d = String(date.getDate()).padStart(2, "0");
  return `${y}-${m}-${d}`;
}

function emptyRow(colspan, text) {
  return `<tr><td colspan="${colspan}" style="text-align:center;color:var(--muted);padding:28px">${text}</td></tr>`;
}

function escapeHtml(value) {
  return String(value)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}
