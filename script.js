const STORAGE_KEYS = {
  templates: "waybill-import-templates",
  activeTemplateId: "waybill-import-active-template-id",
  mappingRules: "waybill-import-mapping-rules",
  previewRows: "waybill-import-preview-rows",
  records: "waybill-import-records",
};

const TEMPERATURE_OPTIONS = ["常温", "冷藏", "冷冻"];

const DEFAULT_FIELDS = [
  { key: "externalCode", label: "外部编码", required: false, type: "text", aliases: ["外部编码", "外部系统订单唯一编号", "客户单号", "外部订单号", "Ref Code"] },
  { key: "senderName", label: "发件人姓名", required: true, type: "text", aliases: ["发件人姓名", "寄件人姓名", "发货人", "发件人", "Sender"] },
  { key: "senderPhone", label: "发件人电话", required: true, type: "phone", aliases: ["发件人电话", "寄件人联系方式", "发货电话", "发件电话", "Sender Tel"] },
  { key: "senderAddress", label: "发件人地址", required: true, type: "text", aliases: ["发件人地址", "寄件人完整地址", "发货地址", "发件地址", "Sender Address"] },
  { key: "receiverName", label: "收件人姓名", required: true, type: "text", aliases: ["收件人姓名", "收货人姓名", "收货人", "收件人", "Receiver"] },
  { key: "receiverPhone", label: "收件人电话", required: true, type: "phone", aliases: ["收件人电话", "收货人联系方式", "收货电话", "收件电话", "Receiver Tel"] },
  { key: "receiverAddress", label: "收件人地址", required: true, type: "text", aliases: ["收件人地址", "收货人完整地址", "收货地址", "收件地址", "Receiver Address"] },
  { key: "weight", label: "重量 (kg)", required: true, type: "number", aliases: ["重量 (kg)", "重量(kg)", "重量(KG)", "货物重量", "Weight(kg)"] },
  { key: "quantity", label: "件数", required: true, type: "integer", aliases: ["件数", "包裹数量", "数量", "Qty"] },
  { key: "temperature", label: "温层", required: true, type: "enum", aliases: ["温层", "温度要求", "Temp Zone"] },
  { key: "remark", label: "备注", required: false, type: "text", aliases: ["备注", "附加说明", "附言", "Note"] },
];

const BUILT_IN_TEMPLATES = [
  {
    id: "standard-template",
    name: "标准模板",
    description: "表头在首行，字段顺序标准",
    builtIn: true,
    fields: DEFAULT_FIELDS,
  },
  {
    id: "ecommerce-template",
    name: "电商模板",
    description: "支持前置说明行和电商列名",
    builtIn: true,
    fields: DEFAULT_FIELDS.map((field) => ({ ...field, aliases: Array.from(new Set(field.aliases)) })),
  },
  {
    id: "english-template",
    name: "英文模板",
    description: "支持英文列名和空行分隔",
    builtIn: true,
    fields: DEFAULT_FIELDS.map((field) => ({ ...field, aliases: Array.from(new Set(field.aliases)) })),
  },
];

function loadJson(key, fallback) {
  try {
    const raw = localStorage.getItem(key);
    return raw ? JSON.parse(raw) : fallback;
  } catch {
    return fallback;
  }
}

function loadTemplates() {
  const saved = loadJson(STORAGE_KEYS.templates, []);
  const merged = [...BUILT_IN_TEMPLATES];
  saved.forEach((tpl) => {
    if (!merged.some((item) => item.id === tpl.id)) merged.push(tpl);
  });
  return merged;
}

const state = {
  templates: loadTemplates(),
  activeTemplateId: localStorage.getItem(STORAGE_KEYS.activeTemplateId) || "standard-template",
  previewRows: loadJson(STORAGE_KEYS.previewRows, []),
  records: loadJson(STORAGE_KEYS.records, []),
  mappingRules: loadJson(STORAGE_KEYS.mappingRules, []),
  selectedSheetName: "",
  currentFile: null,
  workbook: null,
  workbookSheets: [],
  currentSheetRows: [],
  currentHeaderRowIndex: -1,
  currentHeaderCells: [],
  currentMapping: {},
  importErrors: [],
  validationErrors: [],
  duplicateMeta: {},
  submitResult: null,
  recordSearch: "",
  page: 1,
  pageSize: 20,
};

const els = {
  templateList: document.getElementById("templateList"),
  templateName: document.getElementById("templateName"),
  templateDescription: document.getElementById("templateDescription"),
  fieldEditor: document.getElementById("fieldEditor"),
  templateSelect: document.getElementById("templateSelect"),
  sheetSelect: document.getElementById("sheetSelect"),
  fileInput: document.getElementById("fileInput"),
  dropzone: document.getElementById("dropzone"),
  downloadTemplateBtn: document.getElementById("downloadTemplateBtn"),
  newTemplateBtn: document.getElementById("newTemplateBtn"),
  duplicateTemplateBtn: document.getElementById("duplicateTemplateBtn"),
  saveTemplateBtn: document.getElementById("saveTemplateBtn"),
  addFieldBtn: document.getElementById("addFieldBtn"),
  saveMappingBtn: document.getElementById("saveMappingBtn"),
  resetImportBtn: document.getElementById("resetImportBtn"),
  mappingPanel: document.getElementById("mappingPanel"),
  analysisText: document.getElementById("analysisText"),
  importProgressLabel: document.getElementById("importProgressLabel"),
  importProgressCount: document.getElementById("importProgressCount"),
  importProgressBar: document.getElementById("importProgressBar"),
  previewHead: document.getElementById("previewHead"),
  previewBody: document.getElementById("previewBody"),
  addRowBtn: document.getElementById("addRowBtn"),
  exportBtn: document.getElementById("exportBtn"),
  submitBtn: document.getElementById("submitBtn"),
  errorList: document.getElementById("errorList"),
  validationSummary: document.getElementById("validationSummary"),
  submitProgressLabel: document.getElementById("submitProgressLabel"),
  submitProgressCount: document.getElementById("submitProgressCount"),
  submitProgressBar: document.getElementById("submitProgressBar"),
  submitResult: document.getElementById("submitResult"),
  templateCount: document.getElementById("templateCount"),
  previewCount: document.getElementById("previewCount"),
  errorCount: document.getElementById("errorCount"),
  recordCount: document.getElementById("recordCount"),
  activeTemplateChip: document.getElementById("activeTemplateChip"),
  currentSheetChip: document.getElementById("currentSheetChip"),
  recordSearch: document.getElementById("recordSearch"),
  pageSize: document.getElementById("pageSize"),
  recordBody: document.getElementById("recordBody"),
  pageInfo: document.getElementById("pageInfo"),
  prevPageBtn: document.getElementById("prevPageBtn"),
  nextPageBtn: document.getElementById("nextPageBtn"),
  toast: document.getElementById("toast"),
};

function persistTemplates() {
  const customTemplates = state.templates.filter((tpl) => !tpl.builtIn);
  localStorage.setItem(STORAGE_KEYS.templates, JSON.stringify(customTemplates));
}

function persistPreview() {
  localStorage.setItem(STORAGE_KEYS.previewRows, JSON.stringify(state.previewRows));
}

function persistRecords() {
  localStorage.setItem(STORAGE_KEYS.records, JSON.stringify(state.records));
}

function persistMappings() {
  localStorage.setItem(STORAGE_KEYS.mappingRules, JSON.stringify(state.mappingRules));
}

function persistActiveTemplate() {
  localStorage.setItem(STORAGE_KEYS.activeTemplateId, state.activeTemplateId);
}

function normalizeText(value) {
  return String(value ?? "").trim();
}

function slugify(value) {
  return normalizeText(value)
    .toLowerCase()
    .replace(/[^a-z0-9\u4e00-\u9fa5]+/g, "-")
    .replace(/^-+|-+$/g, "");
}

function sanitizeHeader(value) {
  return normalizeText(value)
    .replace(/\s+/g, "")
    .replace(/[（]/g, "(")
    .replace(/[）]/g, ")")
    .toLowerCase();
}

function createId(prefix = "tpl") {
  return `${prefix}-${Math.random().toString(36).slice(2, 8)}-${Date.now().toString(36)}`;
}

function cloneTemplate(template) {
  return JSON.parse(JSON.stringify(template));
}

function getActiveTemplate() {
  return state.templates.find((tpl) => tpl.id === state.activeTemplateId) || state.templates[0];
}

function ensureActiveTemplate() {
  if (!state.templates.some((tpl) => tpl.id === state.activeTemplateId)) {
    state.activeTemplateId = state.templates[0]?.id || "";
  }
}

function escapeHtml(value) {
  return normalizeText(value)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;");
}

function showToast(message) {
  els.toast.textContent = message;
  els.toast.classList.add("show");
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => els.toast.classList.remove("show"), 1800);
}

function waitFrame() {
  return new Promise((resolve) => requestAnimationFrame(() => resolve()));
}

function updateStats() {
  els.templateCount.textContent = String(state.templates.length);
  els.previewCount.textContent = String(state.previewRows.length);
  els.errorCount.textContent = String(state.validationErrors.length + state.importErrors.length);
  els.recordCount.textContent = String(state.records.length);
  const activeTemplate = getActiveTemplate();
  els.activeTemplateChip.textContent = `当前模板：${activeTemplate?.name || "-"}`;
  els.currentSheetChip.textContent = `当前 Sheet：${state.selectedSheetName || "-"}`;
}

function renderTemplateList() {
  ensureActiveTemplate();
  els.templateList.innerHTML = state.templates.map((tpl) => {
    const activeClass = tpl.id === state.activeTemplateId ? "active" : "";
    const tags = [tpl.builtIn ? "内置" : "自定义", `${tpl.fields.length} 字段`];
    return `
      <article class="template-item ${activeClass}" data-template-id="${tpl.id}">
        <h3>${escapeHtml(tpl.name)}</h3>
        <div class="muted">${escapeHtml(tpl.description || "暂无说明")}</div>
        <div class="template-tags">${tags.map((tag) => `<span class="tag">${escapeHtml(tag)}</span>`).join("")}</div>
      </article>
    `;
  }).join("");
  renderTemplateSelect();
}

function renderTemplateSelect() {
  els.templateSelect.innerHTML = state.templates.map((tpl) => `
    <option value="${tpl.id}" ${tpl.id === state.activeTemplateId ? "selected" : ""}>${escapeHtml(tpl.name)}</option>
  `).join("");
}

function renderTemplateEditor() {
  const template = getActiveTemplate();
  if (!template) return;
  els.templateName.value = template.name;
  els.templateDescription.value = template.description || "";
  els.fieldEditor.innerHTML = template.fields.map((field, index) => `
    <div class="field-row" data-index="${index}">
      <input data-field="label" value="${escapeHtml(field.label)}" placeholder="字段名">
      <input data-field="key" value="${escapeHtml(field.key)}" placeholder="字段 key" ${template.builtIn ? "disabled" : ""}>
      <input data-field="aliases" value="${escapeHtml((field.aliases || []).join(","))}" placeholder="别名，英文逗号分隔">
      <select data-field="type">
        ${["text", "phone", "number", "integer", "enum"].map((type) => `<option value="${type}" ${field.type === type ? "selected" : ""}>${type}</option>`).join("")}
      </select>
      <select data-field="required">
        <option value="true" ${field.required ? "selected" : ""}>必填</option>
        <option value="false" ${!field.required ? "selected" : ""}>选填</option>
      </select>
      <button class="row-delete" data-action="remove-field" ${template.builtIn ? "disabled" : ""}>×</button>
    </div>
  `).join("");
}

function renderSheetSelect() {
  if (!state.workbookSheets.length) {
    els.sheetSelect.innerHTML = `<option value="">请选择 Sheet</option>`;
    return;
  }
  els.sheetSelect.innerHTML = state.workbookSheets.map((sheet) => `
    <option value="${escapeHtml(sheet.name)}" ${sheet.name === state.selectedSheetName ? "selected" : ""}>${escapeHtml(sheet.name)}</option>
  `).join("");
}

function renderMappingPanel() {
  const template = getActiveTemplate();
  if (!template || !state.currentHeaderCells.length) {
    els.mappingPanel.innerHTML = `<div class="mapping-tip">上传文件后可查看并调整列映射关系。</div>`;
    return;
  }
  const options = state.currentHeaderCells.map((header) => `<option value="${escapeHtml(header)}">${escapeHtml(header)}</option>`).join("");
  els.mappingPanel.innerHTML = `
    <div class="mapping-tip">系统已根据列名和历史学习规则自动匹配，可手动调整后保存为学习规则。</div>
    <div class="mapping-grid">
      ${template.fields.map((field) => `
        <div class="mapping-row" data-key="${field.key}">
          <div>
            <label>
              <span>${escapeHtml(field.label)}${field.required ? " *" : ""}</span>
              <input value="${escapeHtml(field.type)}" disabled>
            </label>
          </div>
          <div>
            <label>
              <span>映射列</span>
              <select data-field-key="${field.key}">
                <option value="">未映射</option>
                ${options}
              </select>
            </label>
          </div>
          <div class="muted">${escapeHtml((field.aliases || []).join(" / "))}</div>
        </div>
      `).join("")}
    </div>
  `;
  template.fields.forEach((field) => {
    const select = els.mappingPanel.querySelector(`select[data-field-key="${field.key}"]`);
    if (select) select.value = state.currentMapping[field.key] || "";
  });
}

function renderPreviewTable() {
  const template = getActiveTemplate();
  if (!template) return;
  const headColumns = template.fields.map((field) => `<th>${escapeHtml(field.label)}${field.required ? " *" : ""}</th>`).join("");
  els.previewHead.innerHTML = `<tr><th style="min-width:72px;">行号</th>${headColumns}<th style="min-width:96px;">操作</th></tr>`;
  if (!state.previewRows.length) {
    els.previewBody.innerHTML = `<tr class="empty-row"><td colspan="${template.fields.length + 2}">暂无预览数据，请先上传文件。</td></tr>`;
    return;
  }

  els.previewBody.innerHTML = state.previewRows.map((row, rowIndex) => {
    const rowErrors = state.validationErrors.filter((error) => error.rowIndex === rowIndex);
    const rowErrorMap = rowErrors.reduce((acc, error) => {
      acc[error.fieldKey] = acc[error.fieldKey] || [];
      acc[error.fieldKey].push(error.message);
      return acc;
    }, {});
    const duplicate = state.duplicateMeta[rowIndex];
    const rowClass = rowErrors.length ? "row-error" : "";
    const cells = template.fields.map((field) => {
      const messages = rowErrorMap[field.key] || [];
      const invalid = messages.length ? "invalid" : "";
      const duplicateText = field.key === "externalCode" && duplicate ? `<div class="duplicate-badge">${escapeHtml(duplicate)}</div>` : "";
      return `
        <td>
          <div class="cell-wrap">
            <input class="cell-input ${invalid}" data-row="${rowIndex}" data-key="${field.key}" value="${escapeHtml(row[field.key] ?? "")}" title="${escapeHtml(messages.join("；"))}">
            ${messages.length ? `<div class="cell-error">${escapeHtml(messages.join("；"))}</div>` : ""}
            ${duplicateText}
          </div>
        </td>
      `;
    }).join("");
    return `
      <tr class="${rowClass}">
        <td>${rowIndex + 1}</td>
        ${cells}
        <td><button class="btn btn-soft" data-action="delete-row" data-row="${rowIndex}">删除</button></td>
      </tr>
    `;
  }).join("");
}

function renderErrors() {
  const errors = [...state.importErrors, ...state.validationErrors];
  if (!errors.length) {
    els.validationSummary.textContent = "暂无错误。";
    els.errorList.innerHTML = "";
    return;
  }
  els.validationSummary.textContent = `共 ${errors.length} 个问题，已一次性列出全部错误。`;
  els.errorList.innerHTML = errors.map((error) => `<li>${escapeHtml(error.message)}</li>`).join("");
}

function renderRecords() {
  const keyword = normalizeText(state.recordSearch).toLowerCase();
  const filtered = state.records.filter((record) => {
    if (!keyword) return true;
    return [record.externalCode, record.receiverName, record.submittedAt]
      .some((value) => normalizeText(value).toLowerCase().includes(keyword));
  });
  const totalPages = Math.max(1, Math.ceil(filtered.length / state.pageSize));
  if (state.page > totalPages) state.page = totalPages;
  const start = (state.page - 1) * state.pageSize;
  const pageRows = filtered.slice(start, start + state.pageSize);

  if (!pageRows.length) {
    els.recordBody.innerHTML = `<tr class="empty-row"><td colspan="5">暂无历史记录。</td></tr>`;
  } else {
    els.recordBody.innerHTML = pageRows.map((record) => `
      <tr>
        <td>${escapeHtml(record.externalCode || "-")}</td>
        <td>${escapeHtml(record.receiverName)}</td>
        <td>${escapeHtml(record.receiverPhone)}</td>
        <td>${escapeHtml(record.templateName || "-")}</td>
        <td>${escapeHtml(record.submittedAt)}</td>
      </tr>
    `).join("");
  }

  els.pageInfo.textContent = `第 ${state.page} / ${totalPages} 页`;
  els.prevPageBtn.disabled = state.page <= 1;
  els.nextPageBtn.disabled = state.page >= totalPages;
}

function syncTemplateFormToState() {
  const template = getActiveTemplate();
  if (!template) return;
  template.name = normalizeText(els.templateName.value) || template.name;
  template.description = normalizeText(els.templateDescription.value);
}

function updateImportProgress(current, total, label = "导入进度") {
  const percent = total ? Math.round((current / total) * 100) : 0;
  els.importProgressLabel.textContent = `${label} ${percent}%`;
  els.importProgressCount.textContent = `${current} / ${total}`;
  els.importProgressBar.style.width = `${percent}%`;
}

function updateSubmitProgress(current, total) {
  const percent = total ? Math.round((current / total) * 100) : 0;
  els.submitProgressLabel.textContent = `提交进度 ${percent}%`;
  els.submitProgressCount.textContent = `${current} / ${total}`;
  els.submitProgressBar.style.width = `${percent}%`;
}

function detectHeaderRow(rows, template) {
  let best = { rowIndex: -1, headers: [], score: -1 };
  rows.slice(0, 8).forEach((row, rowIndex) => {
    const score = row.reduce((sum, cell) => {
      const header = sanitizeHeader(cell);
      if (!header) return sum;
      const matched = template.fields.some((field) => (field.aliases || [field.label]).some((alias) => sanitizeHeader(alias) === header));
      return sum + (matched ? 1 : 0);
    }, 0);
    if (score > best.score) best = { rowIndex, headers: row.map((item) => normalizeText(item)), score };
  });
  return best;
}

function findBestTemplate(rows) {
  return state.templates.map((template) => {
    const detected = detectHeaderRow(rows, template);
    return {
      templateId: template.id,
      templateName: template.name,
      rowIndex: detected.rowIndex,
      headers: detected.headers,
      score: detected.score,
    };
  }).sort((a, b) => b.score - a.score)[0];
}

function fingerprintHeaders(headers) {
  return headers.map((item) => sanitizeHeader(item)).filter(Boolean).join("|");
}

function findMappingRule(headers, templateId) {
  const fingerprint = fingerprintHeaders(headers);
  const exact = state.mappingRules.find((rule) => rule.templateId === templateId && rule.fingerprint === fingerprint);
  if (exact) return exact;
  const headerSet = new Set(headers.map((item) => sanitizeHeader(item)));
  let bestRule = null;
  let bestScore = 0;
  state.mappingRules.forEach((rule) => {
    if (rule.templateId !== templateId) return;
    const matched = rule.headers.map((item) => sanitizeHeader(item)).filter((item) => headerSet.has(item)).length;
    const score = rule.headers.length ? matched / rule.headers.length : 0;
    if (score > bestScore && score >= 0.6) {
      bestScore = score;
      bestRule = rule;
    }
  });
  return bestRule;
}

function buildMapping(headers, template) {
  const mapping = {};
  const headerNormMap = headers.reduce((acc, header) => {
    acc[sanitizeHeader(header)] = header;
    return acc;
  }, {});
  const learned = findMappingRule(headers, template.id);
  if (learned) Object.assign(mapping, learned.mapping);
  template.fields.forEach((field) => {
    if (mapping[field.key] && headers.includes(mapping[field.key])) return;
    const aliases = field.aliases?.length ? field.aliases : [field.label];
    const matchedAlias = aliases.find((alias) => headerNormMap[sanitizeHeader(alias)]);
    if (matchedAlias) mapping[field.key] = headerNormMap[sanitizeHeader(matchedAlias)];
  });
  return mapping;
}

function getSheetData(sheetName) {
  if (!state.workbook || !sheetName) return [];
  const worksheet = state.workbook.Sheets[sheetName];
  if (!worksheet) return [];
  return XLSX.utils.sheet_to_json(worksheet, {
    header: 1,
    defval: "",
    raw: false,
    blankrows: false,
  });
}

function mapRowsByTemplate(sheetRows, template, headerRowIndex, mapping) {
  return sheetRows.slice(headerRowIndex + 1).reduce((list, row, index) => {
    const item = { __sourceRowNumber: headerRowIndex + 2 + index };
    template.fields.forEach((field) => {
      const header = mapping[field.key];
      const headerIndex = state.currentHeaderCells.indexOf(header);
      item[field.key] = headerIndex >= 0 ? normalizeText(row[headerIndex]) : "";
    });
    const hasValue = template.fields.some((field) => normalizeText(item[field.key]));
    if (hasValue) list.push(item);
    return list;
  }, []);
}

function phoneValid(value) {
  return /^1\d{10}$/.test(normalizeText(value));
}

function validateRows() {
  const template = getActiveTemplate();
  const errors = [];
  const duplicateMeta = {};
  const batchMap = new Map();
  state.previewRows.forEach((row, rowIndex) => {
    const ext = normalizeText(row.externalCode);
    if (!ext) return;
    if (!batchMap.has(ext)) batchMap.set(ext, []);
    batchMap.get(ext).push(rowIndex);
  });

  const existingMap = new Map();
  state.records.forEach((record) => {
    const ext = normalizeText(record.externalCode);
    if (!ext) return;
    if (!existingMap.has(ext)) existingMap.set(ext, []);
    existingMap.get(ext).push(record.submittedAt);
  });

  state.previewRows.forEach((row, rowIndex) => {
    template.fields.forEach((field) => {
      const value = normalizeText(row[field.key]);
      if (field.required && !value) {
        errors.push({ rowIndex, fieldKey: field.key, message: `第 ${rowIndex + 1} 行，${field.label}：必填字段缺失` });
        return;
      }
      if (!value) return;
      if (field.type === "phone" && !phoneValid(value)) {
        errors.push({ rowIndex, fieldKey: field.key, message: `第 ${rowIndex + 1} 行，${field.label}：格式错误` });
      }
      if (field.key === "weight" && !(Number(value) > 0)) {
        errors.push({ rowIndex, fieldKey: field.key, message: `第 ${rowIndex + 1} 行，${field.label}：必须为正数` });
      }
      if (field.key === "quantity") {
        const num = Number(value);
        if (!Number.isInteger(num) || num <= 0) {
          errors.push({ rowIndex, fieldKey: field.key, message: `第 ${rowIndex + 1} 行，${field.label}：必须为正整数` });
        }
      }
      if (field.key === "temperature" && !TEMPERATURE_OPTIONS.includes(value)) {
        errors.push({ rowIndex, fieldKey: field.key, message: `第 ${rowIndex + 1} 行，${field.label}：仅支持 ${TEMPERATURE_OPTIONS.join(" / ")}` });
      }
    });
    const ext = normalizeText(row.externalCode);
    if (ext && batchMap.get(ext)?.length > 1) {
      const siblings = batchMap.get(ext).filter((index) => index !== rowIndex).map((index) => index + 1);
      duplicateMeta[rowIndex] = `同批次重复，重复行：${siblings.join("、")}`;
      errors.push({ rowIndex, fieldKey: "externalCode", message: `第 ${rowIndex + 1} 行，外部编码：与第 ${siblings.join("、")} 行重复` });
    } else if (ext && existingMap.has(ext)) {
      duplicateMeta[rowIndex] = `与历史记录重复，提交时间：${existingMap.get(ext)[0]}`;
      errors.push({ rowIndex, fieldKey: "externalCode", message: `第 ${rowIndex + 1} 行，外部编码：与已存在数据重复` });
    }
  });

  state.validationErrors = errors;
  state.duplicateMeta = duplicateMeta;
  updateStats();
}

function recalcAndRender() {
  validateRows();
  renderPreviewTable();
  renderErrors();
  renderRecords();
  updateStats();
  persistPreview();
}

function downloadWorkbook(rows, template, fileName) {
  const headers = template.fields.map((field) => field.label);
  const jsonRows = rows.map((row) => {
    const obj = {};
    template.fields.forEach((field) => {
      obj[field.label] = row[field.key] ?? "";
    });
    return obj;
  });
  const sheet = XLSX.utils.json_to_sheet(jsonRows, { header: headers });
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, sheet, "订单数据");
  XLSX.writeFile(wb, fileName);
}

function downloadTemplate() {
  const template = getActiveTemplate();
  if (!template) return;
  const example = {
    externalCode: "ORD-2026-001",
    senderName: "张三",
    senderPhone: "13800138000",
    senderAddress: "上海市浦东新区世纪大道100号",
    receiverName: "李四",
    receiverPhone: "13900139000",
    receiverAddress: "北京市朝阳区建国路88号",
    weight: "5.5",
    quantity: "2",
    temperature: "冷藏",
    remark: "样例备注",
  };
  downloadWorkbook([example], template, `${template.name}-导入模板.xlsx`);
  showToast("模板文件已下载");
}

async function fetchRecordsFromApi() {
  try {
    const response = await fetch("/api/records");
    if (!response.ok) return false;
    const payload = await response.json();
    if (!payload || payload.mode === "local_fallback" || !Array.isArray(payload.data)) return false;
    state.records = payload.data;
    persistRecords();
    renderRecords();
    updateStats();
    return true;
  } catch {
    return false;
  }
}

async function saveRecordsToApi(records) {
  try {
    const response = await fetch("/api/submit", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ rows: records }),
    });
    if (!response.ok) return null;
    return await response.json();
  } catch {
    return null;
  }
}

async function readFile(file) {
  if (!file) return;
  const ext = file.name.split(".").pop()?.toLowerCase();
  if (!["xlsx", "xls", "csv", "txt"].includes(ext)) {
    state.importErrors = [{ message: "文件格式错误：仅支持 .xlsx / .xls / .csv / .txt" }];
    renderErrors();
    updateStats();
    return;
  }

  state.currentFile = file;
  state.importErrors = [];
  updateImportProgress(0, 0, "读取文件");

  try {
    const buffer = await file.arrayBuffer();
    if (!buffer.byteLength) throw new Error("文件为空，无法导入");
    if (!window.XLSX) throw new Error("Excel 解析库加载失败，请检查网络或稍后重试");

    state.workbook = XLSX.read(buffer, {
      type: "array",
      cellText: true,
      dense: false,
    });
    state.workbookSheets = state.workbook.SheetNames.map((name) => ({ name }));
    if (!state.workbookSheets.length) throw new Error("Sheet 不存在，未读取到任何工作表");

    state.selectedSheetName = state.workbookSheets.find((item) => !/说明|instruction/i.test(item.name))?.name || state.workbookSheets[0].name;
    renderSheetSelect();
    updateImportProgress(0, 0, "解析工作簿");
    await importCurrentSheet();
  } catch (error) {
    state.importErrors = [{ message: error.message || "导入失败，请检查文件内容或编码" }];
    state.previewRows = [];
    renderMappingPanel();
    recalcAndRender();
    updateImportProgress(0, 0, "导入失败");
  }
}

async function importCurrentSheet() {
  const template = getActiveTemplate();
  if (!template) return;
  if (!state.selectedSheetName) {
    state.importErrors = [{ message: "Sheet 不存在，请手动选择有效工作表" }];
    renderErrors();
    return;
  }

  const rows = getSheetData(state.selectedSheetName);
  state.currentSheetRows = rows;
  if (!rows.length) {
    state.importErrors = [{ message: "文件为空，当前 Sheet 没有可导入数据" }];
    state.previewRows = [];
    renderMappingPanel();
    recalcAndRender();
    return;
  }

  updateImportProgress(0, 0, "识别模板");
  const bestTemplate = findBestTemplate(rows);
  if (bestTemplate.score < 3) {
    state.importErrors = [{ message: "无法自动识别模板，请手动选择模板并调整列映射" }];
  } else {
    state.importErrors = [];
    state.activeTemplateId = bestTemplate.templateId;
    persistActiveTemplate();
  }

  const activeTemplate = getActiveTemplate();
  const headerDetection = detectHeaderRow(rows, activeTemplate);
  state.currentHeaderRowIndex = headerDetection.rowIndex;
  state.currentHeaderCells = headerDetection.headers;

  if (headerDetection.rowIndex < 0 || !headerDetection.headers.length) {
    state.importErrors = [{ message: "未找到表头行，请确认文件格式、编码或 Sheet 内容" }];
    state.previewRows = [];
    renderMappingPanel();
    recalcAndRender();
    updateImportProgress(0, 0, "未识别到表头");
    return;
  }

  state.currentMapping = buildMapping(headerDetection.headers, activeTemplate);
  renderTemplateList();
  renderTemplateEditor();
  renderMappingPanel();

  const missingRequired = activeTemplate.fields.filter((field) => field.required && !state.currentMapping[field.key]).map((field) => field.label);
  if (missingRequired.length) {
    state.importErrors = [{ message: `缺少必要列映射：${missingRequired.join("、")}。请手动补齐映射关系。` }];
  }

  updateImportProgress(0, rows.length ? rows.length - headerDetection.rowIndex - 1 : 0, "转换数据");
  const mapped = [];
  const sheetData = rows.slice(headerDetection.rowIndex + 1);
  const total = sheetData.length || 1;
  for (let i = 0; i < sheetData.length; i += 50) {
    const chunk = sheetData.slice(i, i + 50);
    chunk.forEach((row, chunkIndex) => {
      const item = { __sourceRowNumber: headerDetection.rowIndex + 2 + i + chunkIndex };
      activeTemplate.fields.forEach((field) => {
        const header = state.currentMapping[field.key];
        const headerIndex = headerDetection.headers.indexOf(header);
        item[field.key] = headerIndex >= 0 ? normalizeText(row[headerIndex]) : "";
      });
      const hasValue = activeTemplate.fields.some((field) => normalizeText(item[field.key]));
      if (hasValue) mapped.push(item);
    });
    updateImportProgress(Math.min(i + chunk.length, total), total, "处理数据");
    await waitFrame();
  }

  if (!mapped.length) {
    state.importErrors = [{ message: "文件为空，未读取到有效数据行" }];
  }

  state.previewRows = mapped;
  state.submitResult = null;
  persistPreview();
  recalcAndRender();
  els.analysisText.textContent = `已识别模板：${activeTemplate.name}；表头位于第 ${headerDetection.rowIndex + 1} 行；数据 ${mapped.length} 条；${state.importErrors.length ? "存在待处理问题" : "自动映射完成"}`;
  updateImportProgress(total, total, "导入完成");
}

function saveMappingRule() {
  const template = getActiveTemplate();
  if (!template || !state.currentHeaderCells.length) return;
  const fingerprint = fingerprintHeaders(state.currentHeaderCells);
  const rule = {
    id: createId("map"),
    templateId: template.id,
    templateName: template.name,
    headers: [...state.currentHeaderCells],
    fingerprint,
    mapping: { ...state.currentMapping },
    savedAt: new Date().toLocaleString("zh-CN"),
  };
  state.mappingRules = state.mappingRules.filter((item) => !(item.templateId === template.id && item.fingerprint === fingerprint));
  state.mappingRules.unshift(rule);
  persistMappings();
  showToast("映射规则已保存，下次遇到相似模板会自动应用");
}

function saveTemplate() {
  syncTemplateFormToState();
  const template = getActiveTemplate();
  if (!template) return;
  if (!normalizeText(template.name)) {
    showToast("模板名称不能为空");
    return;
  }
  template.fields = template.fields.map((field) => ({
    ...field,
    key: normalizeText(field.key) || slugify(field.label),
    aliases: (field.aliases || []).map((item) => normalizeText(item)).filter(Boolean),
    required: Boolean(field.required),
  }));
  persistTemplates();
  renderTemplateList();
  renderTemplateEditor();
  updateStats();
  showToast("模板已保存");
}

function createCustomTemplate() {
  const template = {
    id: createId("template"),
    name: "新模板",
    description: "自定义模板说明",
    builtIn: false,
    fields: cloneTemplate(DEFAULT_FIELDS),
  };
  state.templates.push(template);
  state.activeTemplateId = template.id;
  persistActiveTemplate();
  persistTemplates();
  renderAll();
  showToast("已创建新模板");
}

function duplicateTemplate() {
  const template = cloneTemplate(getActiveTemplate());
  template.id = createId("template");
  template.name = `${template.name}-副本`;
  template.builtIn = false;
  state.templates.push(template);
  state.activeTemplateId = template.id;
  persistActiveTemplate();
  persistTemplates();
  renderAll();
  showToast("模板已复制");
}

function addField() {
  const template = getActiveTemplate();
  if (!template || template.builtIn) {
    showToast("内置模板不可直接扩展字段，请先复制为自定义模板");
    return;
  }
  template.fields.push({
    key: createId("field"),
    label: "新字段",
    required: false,
    type: "text",
    aliases: [],
  });
  renderTemplateEditor();
}

function addEmptyRow() {
  const template = getActiveTemplate();
  if (!template) return;
  const row = template.fields.reduce((acc, field) => {
    acc[field.key] = "";
    return acc;
  }, {});
  state.previewRows.push(row);
  recalcAndRender();
}

function clearPreview() {
  state.previewRows = [];
  state.importErrors = [];
  state.validationErrors = [];
  state.duplicateMeta = {};
  state.submitResult = null;
  persistPreview();
  renderAll();
  updateImportProgress(0, 0);
  updateSubmitProgress(0, 0);
  els.analysisText.textContent = "等待上传文件。";
  showToast("预览数据已清空");
}

async function submitOrders() {
  validateRows();
  if (state.validationErrors.length || state.importErrors.length) {
    renderErrors();
    showToast("存在错误数据，请先修正后再提交");
    return;
  }
  if (!state.previewRows.length) {
    showToast("没有可提交的数据");
    return;
  }
  const total = state.previewRows.length;
  let success = 0;
  let failed = 0;
  const results = [];
  updateSubmitProgress(0, total);
  els.submitResult.textContent = "";

  for (let i = 0; i < state.previewRows.length; i += 50) {
    const chunk = state.previewRows.slice(i, i + 50);
    chunk.forEach((row) => {
      success += 1;
      results.push({
        ...row,
        templateName: getActiveTemplate().name,
        submittedAt: new Date().toLocaleString("zh-CN"),
      });
    });
    updateSubmitProgress(Math.min(i + chunk.length, total), total);
    await waitFrame();
  }

  const apiResult = await saveRecordsToApi(results);
  if (apiResult && Array.isArray(apiResult.data)) {
    state.records = apiResult.data;
    persistRecords();
    success = apiResult.successCount ?? success;
    failed = apiResult.failedCount ?? failed;
  } else {
    state.records = [...results, ...state.records];
    persistRecords();
  }
  state.submitResult = { success, failed };
  els.submitResult.textContent = `提交完成：成功 ${success} 条，失败 ${failed} 条`;
  renderRecords();
  updateStats();
  showToast(`提交完成：成功 ${success} 条`);
}

function exportPreview() {
  const template = getActiveTemplate();
  if (!state.previewRows.length) {
    showToast("暂无可导出的预览数据");
    return;
  }
  downloadWorkbook(state.previewRows, template, `预览数据-${Date.now()}.xlsx`);
  showToast("预览数据已导出为 Excel");
}

function updateMappingFromPanel() {
  els.mappingPanel.querySelectorAll("select[data-field-key]").forEach((select) => {
    state.currentMapping[select.dataset.fieldKey] = select.value;
  });
}

function reapplyMappingAndImport() {
  if (!state.currentSheetRows.length || state.currentHeaderRowIndex < 0) return;
  updateMappingFromPanel();
  const template = getActiveTemplate();
  const missingRequired = template.fields.filter((field) => field.required && !state.currentMapping[field.key]).map((field) => field.label);
  if (missingRequired.length) {
    state.importErrors = [{ message: `缺少必要列映射：${missingRequired.join("、")}` }];
    recalcAndRender();
    return;
  }
  state.importErrors = [];
  state.previewRows = mapRowsByTemplate(state.currentSheetRows, template, state.currentHeaderRowIndex, state.currentMapping);
  saveMappingRule();
  recalcAndRender();
  els.analysisText.textContent = `已按手动映射重新生成预览，共 ${state.previewRows.length} 条。`;
  showToast("已应用新的映射关系");
}

function handleTemplateFieldEditor(event) {
  const template = getActiveTemplate();
  if (!template) return;
  const row = event.target.closest(".field-row");
  if (!row) return;
  const index = Number(row.dataset.index);
  const field = template.fields[index];
  if (!field) return;
  if (event.target.matches("[data-field='label']")) field.label = event.target.value;
  if (event.target.matches("[data-field='key']")) field.key = slugify(event.target.value || field.label);
  if (event.target.matches("[data-field='aliases']")) field.aliases = event.target.value.split(",").map((item) => normalizeText(item)).filter(Boolean);
  if (event.target.matches("[data-field='type']")) field.type = event.target.value;
  if (event.target.matches("[data-field='required']")) field.required = event.target.value === "true";
}

function handlePreviewEdit(event) {
  const input = event.target.closest(".cell-input");
  if (!input) return;
  const rowIndex = Number(input.dataset.row);
  const key = input.dataset.key;
  if (!state.previewRows[rowIndex]) return;
  state.previewRows[rowIndex][key] = input.value;
  recalcAndRender();
}

function handleDeleteRow(event) {
  const button = event.target.closest("button[data-action='delete-row']");
  if (!button) return;
  state.previewRows.splice(Number(button.dataset.row), 1);
  recalcAndRender();
}

function renderAll() {
  renderTemplateList();
  renderTemplateEditor();
  renderSheetSelect();
  renderMappingPanel();
  renderPreviewTable();
  renderErrors();
  renderRecords();
  updateStats();
  els.submitResult.textContent = state.submitResult ? `提交完成：成功 ${state.submitResult.success} 条，失败 ${state.submitResult.failed} 条` : "";
}

function bindEvents() {
  els.templateList.addEventListener("click", (event) => {
    const item = event.target.closest("[data-template-id]");
    if (!item) return;
    state.activeTemplateId = item.dataset.templateId;
    persistActiveTemplate();
    renderAll();
  });
  els.templateSelect.addEventListener("change", () => {
    state.activeTemplateId = els.templateSelect.value;
    persistActiveTemplate();
    renderAll();
    if (state.currentHeaderCells.length) {
      state.currentMapping = buildMapping(state.currentHeaderCells, getActiveTemplate());
      renderMappingPanel();
    }
  });
  els.templateName.addEventListener("input", syncTemplateFormToState);
  els.templateDescription.addEventListener("input", syncTemplateFormToState);
  els.newTemplateBtn.addEventListener("click", createCustomTemplate);
  els.duplicateTemplateBtn.addEventListener("click", duplicateTemplate);
  els.saveTemplateBtn.addEventListener("click", saveTemplate);
  els.addFieldBtn.addEventListener("click", addField);
  els.downloadTemplateBtn.addEventListener("click", downloadTemplate);
  els.resetImportBtn.addEventListener("click", clearPreview);
  els.saveMappingBtn.addEventListener("click", reapplyMappingAndImport);
  els.addRowBtn.addEventListener("click", addEmptyRow);
  els.exportBtn.addEventListener("click", exportPreview);
  els.submitBtn.addEventListener("click", submitOrders);
  els.sheetSelect.addEventListener("change", async () => {
    state.selectedSheetName = els.sheetSelect.value;
    updateStats();
    await importCurrentSheet();
  });
  els.fieldEditor.addEventListener("input", handleTemplateFieldEditor);
  els.fieldEditor.addEventListener("change", handleTemplateFieldEditor);
  els.fieldEditor.addEventListener("click", (event) => {
    const button = event.target.closest("[data-action='remove-field']");
    if (!button) return;
    const template = getActiveTemplate();
    if (!template || template.builtIn) return;
    const row = event.target.closest(".field-row");
    template.fields.splice(Number(row.dataset.index), 1);
    renderTemplateEditor();
  });
  els.mappingPanel.addEventListener("change", (event) => {
    if (event.target.matches("select[data-field-key]")) state.currentMapping[event.target.dataset.fieldKey] = event.target.value;
  });
  els.previewBody.addEventListener("input", handlePreviewEdit);
  els.previewBody.addEventListener("click", handleDeleteRow);
  els.fileInput.addEventListener("change", () => {
    const file = els.fileInput.files?.[0];
    if (file) readFile(file);
    els.fileInput.value = "";
  });
  ["dragenter", "dragover"].forEach((eventName) => {
    els.dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      els.dropzone.classList.add("dragover");
    });
  });
  ["dragleave", "drop"].forEach((eventName) => {
    els.dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      els.dropzone.classList.remove("dragover");
    });
  });
  els.dropzone.addEventListener("drop", (event) => {
    const file = event.dataTransfer?.files?.[0];
    if (file) readFile(file);
  });
  els.recordSearch.addEventListener("input", () => {
    state.recordSearch = els.recordSearch.value;
    state.page = 1;
    renderRecords();
  });
  els.pageSize.addEventListener("change", () => {
    state.pageSize = Number(els.pageSize.value) || 20;
    state.page = 1;
    renderRecords();
  });
  els.prevPageBtn.addEventListener("click", () => {
    if (state.page > 1) state.page -= 1;
    renderRecords();
  });
  els.nextPageBtn.addEventListener("click", () => {
    state.page += 1;
    renderRecords();
  });
}

function boot() {
  ensureActiveTemplate();
  state.pageSize = Number(els.pageSize.value) || 20;
  validateRows();
  renderAll();
  bindEvents();
  fetchRecordsFromApi();
}

boot();
