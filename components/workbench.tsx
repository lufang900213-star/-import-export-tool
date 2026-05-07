"use client";

import { useEffect, useMemo, useRef, useState } from "react";
import * as XLSX from "xlsx";
import { fetchRecords, submitRecords } from "../lib/api";
import { TEMPLATE_PROFILES } from "../lib/constants";
import {
  loadMappingRules,
  loadPreviewRows,
  loadTemplateId,
  persistMappingRules,
  persistPreviewRows,
  persistTemplateId,
} from "../lib/storage";
import {
  buildMapping,
  buildSubmitPayload,
  buildTemplateWorkbook,
  createEmptyRow,
  exportRowsToWorkbook,
  getSheetRows,
  guessTemplate,
  mapRowsWithProgress,
  preparePreferredSheet,
  readWorkbookFile,
  saveMappingRule,
  validateRows,
} from "../lib/workbook";
import type {
  FieldKey,
  MappingRule,
  PersistedRecord,
  PreviewRow,
  ProgressState,
  ValidationError,
} from "../lib/types";

const EMPTY_PROGRESS: ProgressState = {
  label: "等待开始",
  current: 0,
  total: 0,
  percent: 0,
};

function getTemplate(templateId: string) {
  return TEMPLATE_PROFILES.find((item) => item.id === templateId) || TEMPLATE_PROFILES[0];
}

export function Workbench() {
  const [templateId, setTemplateId] = useState(TEMPLATE_PROFILES[0].id);
  const [records, setRecords] = useState<PersistedRecord[]>([]);
  const [previewRows, setPreviewRows] = useState<PreviewRow[]>([]);
  const [mappingRules, setMappingRules] = useState<MappingRule[]>([]);
  const [validationErrors, setValidationErrors] = useState<ValidationError[]>([]);
  const [duplicateMeta, setDuplicateMeta] = useState<Record<string, string>>({});
  const [importIssues, setImportIssues] = useState<string[]>([]);
  const [analysisText, setAnalysisText] = useState("等待上传 Excel 文件");
  const [workbook, setWorkbook] = useState<XLSX.WorkBook | null>(null);
  const [sheetNames, setSheetNames] = useState<string[]>([]);
  const [selectedSheet, setSelectedSheet] = useState("");
  const [headers, setHeaders] = useState<string[]>([]);
  const [headerRowIndex, setHeaderRowIndex] = useState(-1);
  const [rawRows, setRawRows] = useState<string[][]>([]);
  const [mapping, setMapping] = useState<Partial<Record<FieldKey, string>>>({});
  const [importProgress, setImportProgress] = useState<ProgressState>(EMPTY_PROGRESS);
  const [submitProgress, setSubmitProgress] = useState<ProgressState>(EMPTY_PROGRESS);
  const [submitSummary, setSubmitSummary] = useState<{ success: number; failed: number } | null>(null);
  const [toast, setToast] = useState("");
  const [recordSearch, setRecordSearch] = useState("");
  const [page, setPage] = useState(1);
  const [pageSize, setPageSize] = useState(10);
  const fileInputRef = useRef<HTMLInputElement | null>(null);

  const activeTemplate = useMemo(() => getTemplate(templateId), [templateId]);

  useEffect(() => {
    setTemplateId(loadTemplateId());
    setPreviewRows(loadPreviewRows());
    setMappingRules(loadMappingRules());
  }, []);

  useEffect(() => {
    persistTemplateId(templateId);
  }, [templateId]);

  useEffect(() => {
    persistPreviewRows(previewRows);
  }, [previewRows]);

  useEffect(() => {
    persistMappingRules(mappingRules);
  }, [mappingRules]);

  useEffect(() => {
    void fetchRecords()
      .then((data) => setRecords(data))
      .catch(() => setToast("读取历史运单失败"));
  }, []);

  useEffect(() => {
    const result = validateRows(previewRows, records);
    setValidationErrors(result.validationErrors);
    setDuplicateMeta(result.duplicateMeta);
  }, [previewRows, records]);

  useEffect(() => {
    if (!toast) return;
    const timer = window.setTimeout(() => setToast(""), 1800);
    return () => window.clearTimeout(timer);
  }, [toast]);

  const filteredRecords = useMemo(() => {
    const keyword = recordSearch.trim().toLowerCase();
    if (!keyword) return records;
    return records.filter((record) =>
      [record.externalCode, record.receiverName, record.submittedAt].some((value) =>
        String(value ?? "").toLowerCase().includes(keyword),
      ),
    );
  }, [records, recordSearch]);

  const totalPages = Math.max(1, Math.ceil(filteredRecords.length / pageSize));
  const safePage = Math.min(page, totalPages);
  const pagedRecords = filteredRecords.slice((safePage - 1) * pageSize, safePage * pageSize);

  async function importSheetRows(nextWorkbook: XLSX.WorkBook, sheetName: string, forcedTemplateId?: string) {
    const rows = getSheetRows(nextWorkbook, sheetName);
    setRawRows(rows);

    if (!rows.length) {
      setImportIssues(["当前 Sheet 没有可导入的数据"]);
      setPreviewRows([]);
      setAnalysisText("当前 Sheet 为空");
      return;
    }

    const guessed = guessTemplate(rows);
    const resolvedTemplateId = forcedTemplateId || (guessed.score >= 4 ? guessed.templateId : templateId);
    setTemplateId(resolvedTemplateId);

    const resolvedTemplate = getTemplate(resolvedTemplateId);
    const resolvedHeaderRowIndex = guessed.rowIndex;
    const resolvedHeaders = rows[resolvedHeaderRowIndex] || [];
    setHeaderRowIndex(resolvedHeaderRowIndex);
    setHeaders(resolvedHeaders);

    if (resolvedHeaderRowIndex < 0 || !resolvedHeaders.length) {
      setImportIssues(["未找到表头行，请检查文件内容"]);
      setPreviewRows([]);
      setAnalysisText("未识别到表头");
      return;
    }

    const nextMapping = buildMapping(resolvedHeaders, resolvedTemplate, mappingRules);
    setMapping(nextMapping);

    const missingRequired = resolvedTemplate.fields
      .filter((field) => field.required && !nextMapping[field.key])
      .map((field) => field.label);
    setImportIssues(missingRequired.length ? [`缺少必要列映射：${missingRequired.join("、")}`] : []);

    const mappedRows = await mapRowsWithProgress(
      rows,
      resolvedHeaderRowIndex,
      resolvedHeaders,
      resolvedTemplate,
      nextMapping,
      setImportProgress,
    );
    setPreviewRows(mappedRows);
    setAnalysisText(
      `已识别模板：${resolvedTemplate.name}；表头位于第 ${resolvedHeaderRowIndex + 1} 行；共导入 ${mappedRows.length} 条数据`,
    );
    setImportProgress({
      label: "导入完成",
      current: mappedRows.length,
      total: mappedRows.length,
      percent: mappedRows.length ? 100 : 0,
    });
  }

  async function handleFile(file: File) {
    try {
      setImportIssues([]);
      setImportProgress({ label: "读取文件", current: 0, total: 0, percent: 0 });
      const nextWorkbook = await readWorkbookFile(file);
      const nextSheetNames = nextWorkbook.SheetNames;
      const preferredSheet = preparePreferredSheet(nextSheetNames);
      setWorkbook(nextWorkbook);
      setSheetNames(nextSheetNames);
      setSelectedSheet(preferredSheet);
      await importSheetRows(nextWorkbook, preferredSheet);
    } catch (error) {
      setImportIssues([error instanceof Error ? error.message : "导入失败"]);
      setPreviewRows([]);
      setAnalysisText("导入失败");
      setToast("导入失败");
    }
  }

  async function handleSheetChange(nextSheet: string) {
    setSelectedSheet(nextSheet);
    if (!workbook) return;
    await importSheetRows(workbook, nextSheet);
  }

  async function handleTemplateChange(nextTemplateId: string) {
    setTemplateId(nextTemplateId);
    if (workbook && selectedSheet) {
      await importSheetRows(workbook, selectedSheet, nextTemplateId);
    }
  }

  function applyManualMapping() {
    if (!rawRows.length || headerRowIndex < 0 || !headers.length) return;
    const nextRules = saveMappingRule(activeTemplate, headers, mapping, mappingRules);
    setMappingRules(nextRules);
    void mapRowsWithProgress(rawRows, headerRowIndex, headers, activeTemplate, mapping, setImportProgress).then((rows) => {
      setPreviewRows(rows);
      setAnalysisText(`已按手动映射重新生成预览，共 ${rows.length} 条`);
      setToast("映射规则已保存");
    });
  }

  function updateCell(rowId: string, fieldKey: FieldKey, value: string) {
    setPreviewRows((current) =>
      current.map((row) => (row.id === rowId ? { ...row, values: { ...row.values, [fieldKey]: value } } : row)),
    );
  }

  function addEmptyRow() {
    setPreviewRows((current) => [...current, createEmptyRow()]);
  }

  function removeRow(rowId: string) {
    setPreviewRows((current) => current.filter((row) => row.id !== rowId));
  }

  function clearPreview() {
    setPreviewRows([]);
    setImportIssues([]);
    setValidationErrors([]);
    setDuplicateMeta({});
    setAnalysisText("等待上传 Excel 文件");
    setImportProgress(EMPTY_PROGRESS);
    setSubmitProgress(EMPTY_PROGRESS);
    setSubmitSummary(null);
  }

  async function handleSubmit() {
    const result = validateRows(previewRows, records);
    setValidationErrors(result.validationErrors);
    setDuplicateMeta(result.duplicateMeta);
    if (result.validationErrors.length || importIssues.length) {
      setToast("存在错误，请先修正");
      return;
    }
    if (!previewRows.length) {
      setToast("没有可提交的数据");
      return;
    }

    const payload = buildSubmitPayload(previewRows, activeTemplate);
    setSubmitProgress({ label: "提交进度", current: 0, total: payload.length, percent: 0 });
    for (let index = 0; index < payload.length; index += 50) {
      const current = Math.min(index + 50, payload.length);
      setSubmitProgress({
        label: "提交进度",
        current,
        total: payload.length,
        percent: payload.length ? Math.round((current / payload.length) * 100) : 0,
      });
      await new Promise((resolve) => setTimeout(resolve, 0));
    }

    try {
      const response = await submitRecords(payload);
      setRecords(response.data);
      setSubmitSummary({ success: response.successCount, failed: response.failedCount });
      setToast(`提交成功 ${response.successCount} 条`);
    } catch (error) {
      setToast(error instanceof Error ? error.message : "提交失败");
    }
  }

  const errorMap = useMemo(() => {
    return validationErrors.reduce<Record<string, Record<string, string[]>>>((acc, error) => {
      acc[error.rowId] ??= {};
      acc[error.rowId][error.fieldKey] ??= [];
      acc[error.rowId][error.fieldKey].push(error.message);
      return acc;
    }, {});
  }, [validationErrors]);

  return (
    <div className="page-shell">
      <div className="page-glow page-glow-a" />
      <div className="page-glow page-glow-b" />
      <main className="page">
        <section className="hero card">
          <div>
            <div className="eyebrow">Waybill Import Workbench</div>
            <h1>多模板识别、在线校验编辑、批量提交与历史运单查看</h1>
            <p>面向物流批量下单场景，支持不同列名、不同列序、相似模板复用映射规则，并将成功提交数据持久化到数据库。</p>
          </div>
          <div className="hero-meta">
            <span>当前模板：{activeTemplate.name}</span>
            <span>当前 Sheet：{selectedSheet || "-"}</span>
          </div>
        </section>

        <section className="stats-grid">
          <StatCard label="模板数量" value={String(TEMPLATE_PROFILES.length)} />
          <StatCard label="预览行数" value={String(previewRows.length)} />
          <StatCard label="错误数量" value={String(validationErrors.length + importIssues.length)} />
          <StatCard label="历史运单" value={String(records.length)} />
        </section>

        <section className="module-grid">
          <div className="card">
            <div className="section-head">
              <div>
                <div className="subhead">模块一</div>
                <h2>模板管理与文件导入</h2>
              </div>
              <button className="ghost-button" onClick={() => buildTemplateWorkbook(activeTemplate)}>
                下载模板
              </button>
            </div>

            <div className="template-row">
              <label className="field">
                <span>模板选择</span>
                <select value={templateId} onChange={(event) => void handleTemplateChange(event.target.value)}>
                  {TEMPLATE_PROFILES.map((template) => (
                    <option key={template.id} value={template.id}>
                      {template.name}
                    </option>
                  ))}
                </select>
              </label>
              <label className="field">
                <span>Sheet 选择</span>
                <select value={selectedSheet} onChange={(event) => void handleSheetChange(event.target.value)}>
                  {sheetNames.length ? (
                    sheetNames.map((name) => (
                      <option key={name} value={name}>
                        {name}
                      </option>
                    ))
                  ) : (
                    <option value="">请先上传文件</option>
                  )}
                </select>
              </label>
            </div>

            <button className="upload-zone" onClick={() => fileInputRef.current?.click()}>
              <strong>拖拽 Excel 到此处，或点击上传</strong>
              <span>仅支持 .xlsx / .xls，系统将自动识别模板、Sheet 和列映射</span>
            </button>
            <input
              ref={fileInputRef}
              type="file"
              accept=".xlsx,.xls"
              hidden
              onChange={(event) => {
                const file = event.target.files?.[0];
                if (file) void handleFile(file);
                event.target.value = "";
              }}
            />

            <ProgressCard progress={importProgress} />

            <div className="mapping-card">
              <div className="section-head compact">
                <div>
                  <div className="subhead">手动映射</div>
                  <p>自动识别失败时，可手动选择 Excel 列与系统字段的映射关系。</p>
                </div>
                <button className="primary-button" onClick={applyManualMapping}>
                  应用并记忆
                </button>
              </div>
              <div className="mapping-grid">
                {activeTemplate.fields.map((field) => (
                  <label className="mapping-item" key={field.key}>
                    <span>
                      {field.label}
                      {field.required ? " *" : ""}
                    </span>
                    <select
                      value={mapping[field.key] || ""}
                      onChange={(event) => setMapping((current) => ({ ...current, [field.key]: event.target.value }))}
                    >
                      <option value="">未映射</option>
                      {headers.map((header) => (
                        <option key={header} value={header}>
                          {header}
                        </option>
                      ))}
                    </select>
                    <small>{field.aliases.join(" / ")}</small>
                  </label>
                ))}
              </div>
            </div>

            <div className="analysis-box">
              <strong>识别结果</strong>
              <p>{analysisText}</p>
              {submitSummary ? <p>最近一次提交：成功 {submitSummary.success} 条，失败 {submitSummary.failed} 条</p> : null}
            </div>
          </div>

          <div className="card">
            <div className="section-head">
              <div>
                <div className="subhead">模块二</div>
                <h2>数据预览与在线编辑</h2>
              </div>
              <div className="inline-actions">
                <button className="ghost-button" onClick={addEmptyRow}>
                  新增空行
                </button>
                <button className="ghost-button" onClick={() => exportRowsToWorkbook(previewRows, activeTemplate)}>
                  导出 Excel
                </button>
                <button className="ghost-button" onClick={clearPreview}>
                  清空预览
                </button>
                <button className="primary-button" onClick={() => void handleSubmit()}>
                  提交下单
                </button>
              </div>
            </div>

            <div className="error-panel">
              <strong>全部错误</strong>
              {importIssues.length || validationErrors.length ? (
                <ul>
                  {importIssues.map((issue) => (
                    <li key={issue}>{issue}</li>
                  ))}
                  {validationErrors.map((error) => (
                    <li key={`${error.rowId}-${error.fieldKey}-${error.message}`}>{error.message}</li>
                  ))}
                </ul>
              ) : (
                <p>当前没有错误</p>
              )}
            </div>

            <div className="table-wrap">
              <table className="excel-table">
                <thead>
                  <tr>
                    <th>行号</th>
                    {activeTemplate.fields.map((field) => (
                      <th key={field.key}>
                        {field.label}
                        {field.required ? " *" : ""}
                      </th>
                    ))}
                    <th>操作</th>
                  </tr>
                </thead>
                <tbody>
                  {previewRows.length ? (
                    previewRows.map((row, rowIndex) => (
                      <tr key={row.id}>
                        <td>{rowIndex + 1}</td>
                        {activeTemplate.fields.map((field) => {
                          const messages = errorMap[row.id]?.[field.key] || [];
                          return (
                            <td key={field.key}>
                              <div className="cell-stack">
                                <input
                                  className={messages.length ? "cell-input invalid" : "cell-input"}
                                  value={row.values[field.key]}
                                  onChange={(event) => updateCell(row.id, field.key, event.target.value)}
                                />
                                {messages.length ? <small className="error-text">{messages.join("；")}</small> : null}
                                {field.key === "externalCode" && duplicateMeta[row.id] ? (
                                  <small className="duplicate-text">{duplicateMeta[row.id]}</small>
                                ) : null}
                              </div>
                            </td>
                          );
                        })}
                        <td>
                          <button className="danger-button" onClick={() => removeRow(row.id)}>
                            删除
                          </button>
                        </td>
                      </tr>
                    ))
                  ) : (
                    <tr>
                      <td colSpan={activeTemplate.fields.length + 2} className="empty-cell">
                        暂无预览数据
                      </td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>

            <ProgressCard progress={submitProgress} />
          </div>
        </section>

        <section className="card">
          <div className="section-head">
            <div>
              <div className="subhead">模块四</div>
              <h2>已导入运单列表</h2>
            </div>
            <div className="inline-actions">
              <input
                className="search-input"
                placeholder="搜索外部编码 / 收件人 / 提交时间"
                value={recordSearch}
                onChange={(event) => {
                  setRecordSearch(event.target.value);
                  setPage(1);
                }}
              />
              <select
                value={pageSize}
                onChange={(event) => {
                  setPageSize(Number(event.target.value));
                  setPage(1);
                }}
              >
                <option value={10}>10 / 页</option>
                <option value={20}>20 / 页</option>
                <option value={50}>50 / 页</option>
              </select>
            </div>
          </div>

          <div className="table-wrap">
            <table className="records-table">
              <thead>
                <tr>
                  <th>外部编码</th>
                  <th>收件人姓名</th>
                  <th>收件人电话</th>
                  <th>模板</th>
                  <th>提交时间</th>
                </tr>
              </thead>
              <tbody>
                {pagedRecords.length ? (
                  pagedRecords.map((record) => (
                    <tr key={record.id}>
                      <td>{record.externalCode || "-"}</td>
                      <td>{record.receiverName || "-"}</td>
                      <td>{record.receiverPhone || "-"}</td>
                      <td>{record.templateName || "-"}</td>
                      <td>{record.submittedAt || "-"}</td>
                    </tr>
                  ))
                ) : (
                  <tr>
                    <td colSpan={5} className="empty-cell">
                      暂无历史记录
                    </td>
                  </tr>
                )}
              </tbody>
            </table>
          </div>

          <div className="pagination">
            <button className="ghost-button" disabled={safePage <= 1} onClick={() => setPage((current) => Math.max(1, current - 1))}>
              上一页
            </button>
            <span>
              第 {safePage} / {totalPages} 页
            </span>
            <button
              className="ghost-button"
              disabled={safePage >= totalPages}
              onClick={() => setPage((current) => Math.min(totalPages, current + 1))}
            >
              下一页
            </button>
          </div>
        </section>
      </main>

      {toast ? <div className="toast">{toast}</div> : null}
    </div>
  );
}

function StatCard({ label, value }: { label: string; value: string }) {
  return (
    <article className="stat-card card">
      <span>{label}</span>
      <strong>{value}</strong>
    </article>
  );
}

function ProgressCard({ progress }: { progress: ProgressState }) {
  return (
    <div className="progress-card">
      <div className="progress-head">
        <span>
          {progress.label} {progress.percent}%
        </span>
        <span>
          {progress.current} / {progress.total}
        </span>
      </div>
      <div className="progress-track">
        <i style={{ width: `${progress.percent}%` }} />
      </div>
    </div>
  );
}
