import * as XLSX from "xlsx";
import { SYSTEM_FIELDS, TEMPERATURE_OPTIONS, TEMPLATE_PROFILES } from "./constants";
import type {
  FieldKey,
  MappingRule,
  PersistedRecord,
  PreviewRow,
  ProgressState,
  TemplateField,
  TemplateProfile,
  ValidationError,
} from "./types";
import { createId, fingerprintHeaders, formatDateTime, normalizeHeader, normalizeText, sleepFrame } from "./utils";

export function createEmptyRow(): PreviewRow {
  const values = SYSTEM_FIELDS.reduce<Record<FieldKey, string>>((acc, field) => {
    acc[field.key] = "";
    return acc;
  }, {} as Record<FieldKey, string>);

  return {
    id: createId("row"),
    sourceRowNumber: 0,
    values,
  };
}

export async function readWorkbookFile(file: File): Promise<XLSX.WorkBook> {
  const ext = file.name.split(".").pop()?.toLowerCase();
  if (!ext || !["xlsx", "xls"].includes(ext)) {
    throw new Error("文件格式错误，仅支持 .xlsx / .xls");
  }

  const buffer = await file.arrayBuffer();
  if (!buffer.byteLength) {
    throw new Error("文件为空，无法导入");
  }

  try {
    const workbook = XLSX.read(buffer, { type: "array", cellDates: false, dense: false });
    if (!workbook.SheetNames.length) {
      throw new Error("工作簿中没有可用 Sheet");
    }
    return workbook;
  } catch (error) {
    throw new Error(error instanceof Error ? `Excel 解析失败：${error.message}` : "Excel 解析失败");
  }
}

export function getSheetRows(workbook: XLSX.WorkBook, sheetName: string): string[][] {
  const sheet = workbook.Sheets[sheetName];
  if (!sheet) return [];

  const rows = XLSX.utils.sheet_to_json<unknown[]>(sheet, {
    header: 1,
    blankrows: false,
    defval: "",
    raw: false,
  });

  return rows
    .map((row) => row.map((cell) => normalizeText(cell)))
    .filter((row) => row.some((cell) => normalizeText(cell)));
}

function matchFieldScore(cell: string, field: TemplateField): number {
  const normalized = normalizeHeader(cell);
  if (!normalized) return 0;
  if (normalizeHeader(field.label) === normalized) return 4;
  if (field.aliases.some((alias) => normalizeHeader(alias) === normalized)) return 3;
  if (normalized.includes(normalizeHeader(field.label))) return 2;
  return 0;
}

export function detectHeaderRow(
  rows: string[][],
  template: TemplateProfile,
): { rowIndex: number; headers: string[]; score: number } {
  let best = { rowIndex: -1, headers: [] as string[], score: -1 };

  rows.slice(0, 10).forEach((row, rowIndex) => {
    const score = row.reduce((sum, cell) => {
      const cellScore = template.fields.reduce((fieldBest, field) => Math.max(fieldBest, matchFieldScore(cell, field)), 0);
      return sum + cellScore;
    }, 0);
    if (score > best.score) {
      best = { rowIndex, headers: row, score };
    }
  });

  return best;
}

export function guessTemplate(rows: string[][]): {
  templateId: string;
  score: number;
  rowIndex: number;
  headers: string[];
} {
  const guesses = TEMPLATE_PROFILES.map((template) => {
    const result = detectHeaderRow(rows, template);
    return {
      templateId: template.id,
      score: result.score,
      rowIndex: result.rowIndex,
      headers: result.headers,
    };
  });

  return guesses.sort((a, b) => b.score - a.score)[0] || {
    templateId: TEMPLATE_PROFILES[0].id,
    score: 0,
    rowIndex: -1,
    headers: [],
  };
}

function resolveMappingFromRule(headers: string[], rule: MappingRule): Partial<Record<FieldKey, string>> {
  const headerMap = new Map(headers.map((header) => [normalizeHeader(header), header]));
  return Object.entries(rule.mapping).reduce<Partial<Record<FieldKey, string>>>((acc, [fieldKey, header]) => {
    const matched = headerMap.get(normalizeHeader(header));
    if (matched) {
      acc[fieldKey as FieldKey] = matched;
    }
    return acc;
  }, {});
}

export function buildMapping(
  headers: string[],
  template: TemplateProfile,
  rules: MappingRule[],
): Partial<Record<FieldKey, string>> {
  const fingerprint = fingerprintHeaders(headers);
  const exactRule = rules.find((rule) => rule.templateId === template.id && rule.fingerprint === fingerprint);
  if (exactRule) {
    return resolveMappingFromRule(headers, exactRule);
  }

  const headerMap = headers.map((header) => ({ raw: header, normalized: normalizeHeader(header) }));
  return template.fields.reduce<Partial<Record<FieldKey, string>>>((acc, field) => {
    const hit = headerMap.find((item) => {
      const aliases = [field.label, ...field.aliases].map((alias) => normalizeHeader(alias));
      return aliases.includes(item.normalized);
    });
    if (hit) {
      acc[field.key] = hit.raw;
    }
    return acc;
  }, {});
}

export async function mapRowsWithProgress(
  rows: string[][],
  headerRowIndex: number,
  headers: string[],
  template: TemplateProfile,
  mapping: Partial<Record<FieldKey, string>>,
  onProgress?: (progress: ProgressState) => void,
): Promise<PreviewRow[]> {
  const dataRows = rows.slice(headerRowIndex + 1);
  const headerIndexMap = new Map(headers.map((header, index) => [header, index]));
  const result: PreviewRow[] = [];
  const total = dataRows.length;

  for (let start = 0; start < dataRows.length; start += 50) {
    const chunk = dataRows.slice(start, start + 50);
    chunk.forEach((row, offset) => {
      const preview = createEmptyRow();
      preview.sourceRowNumber = headerRowIndex + 2 + start + offset;
      template.fields.forEach((field) => {
        const header = mapping[field.key];
        const index = header ? headerIndexMap.get(header) : -1;
        preview.values[field.key] = index == null || index < 0 ? "" : normalizeText(row[index]);
      });
      if (template.fields.some((field) => normalizeText(preview.values[field.key]))) {
        result.push(preview);
      }
    });

    const current = Math.min(start + chunk.length, total);
    onProgress?.({
      label: "导入进度",
      current,
      total,
      percent: total ? Math.round((current / total) * 100) : 0,
    });
    await sleepFrame();
  }

  return result;
}

export function validateRows(
  rows: PreviewRow[],
  records: PersistedRecord[],
): {
  validationErrors: ValidationError[];
  duplicateMeta: Record<string, string>;
} {
  const validationErrors: ValidationError[] = [];
  const duplicateMeta: Record<string, string> = {};
  const recordMap = new Map<string, PersistedRecord>();
  const seenExternalCodes = new Map<string, PreviewRow>();

  records.forEach((record) => {
    if (record.externalCode) {
      recordMap.set(record.externalCode, record);
    }
  });

  const pushError = (row: PreviewRow, fieldKey: FieldKey, message: string) => {
    validationErrors.push({
      rowId: row.id,
      rowNumber: row.sourceRowNumber || rows.findIndex((item) => item.id === row.id) + 1,
      fieldKey,
      message,
    });
  };

  rows.forEach((row, rowIndex) => {
    SYSTEM_FIELDS.forEach((field) => {
      const value = normalizeText(row.values[field.key]);
      const rowLabel = `第 ${rowIndex + 1} 行`;

      if (field.required && !value) {
        pushError(row, field.key, `${rowLabel}，${field.label}：不能为空`);
        return;
      }

      if (!value) return;

      if (field.type === "phone" && !/^[0-9+\-\s()]{6,20}$/.test(value)) {
        pushError(row, field.key, `${rowLabel}，${field.label}：格式错误`);
      }

      if (field.key === "weight") {
        const num = Number(value);
        if (!Number.isFinite(num) || num <= 0) {
          pushError(row, field.key, `${rowLabel}，${field.label}：必须为正数`);
        }
      }

      if (field.key === "quantity") {
        const num = Number(value);
        if (!Number.isInteger(num) || num <= 0) {
          pushError(row, field.key, `${rowLabel}，${field.label}：必须为正整数`);
        }
      }

      if (field.key === "temperature" && !TEMPERATURE_OPTIONS.includes(value as (typeof TEMPERATURE_OPTIONS)[number])) {
        pushError(row, field.key, `${rowLabel}，${field.label}：必须为 ${TEMPERATURE_OPTIONS.join(" / ")}`);
      }
    });

    const externalCode = normalizeText(row.values.externalCode);
    if (!externalCode) return;

    const existing = recordMap.get(externalCode);
    if (existing) {
      duplicateMeta[row.id] = `与历史记录重复（提交时间：${existing.submittedAt}）`;
      pushError(row, "externalCode", `第 ${rowIndex + 1} 行，外部编码：与历史记录重复`);
    }

    const seenRow = seenExternalCodes.get(externalCode);
    if (!seenRow) {
      seenExternalCodes.set(externalCode, row);
      return;
    }

    const firstIndex = rows.findIndex((item) => item.id === seenRow.id) + 1;
    duplicateMeta[seenRow.id] = `与第 ${rowIndex + 1} 行重复`;
    duplicateMeta[row.id] = `与第 ${firstIndex} 行重复`;
    pushError(seenRow, "externalCode", `第 ${firstIndex} 行，外部编码：与第 ${rowIndex + 1} 行重复`);
    pushError(row, "externalCode", `第 ${rowIndex + 1} 行，外部编码：与第 ${firstIndex} 行重复`);
  });

  return { validationErrors, duplicateMeta };
}

export function saveMappingRule(
  template: TemplateProfile,
  headers: string[],
  mapping: Partial<Record<FieldKey, string>>,
  rules: MappingRule[],
): MappingRule[] {
  const fingerprint = fingerprintHeaders(headers);
  const nextRule: MappingRule = {
    id: createId("map"),
    templateId: template.id,
    templateName: template.name,
    headers: [...headers],
    fingerprint,
    mapping,
    savedAt: formatDateTime(),
  };

  return [nextRule, ...rules.filter((rule) => !(rule.templateId === template.id && rule.fingerprint === fingerprint))];
}

export function exportRowsToWorkbook(rows: PreviewRow[], template: TemplateProfile): void {
  const data = rows.map((row) =>
    template.fields.reduce<Record<string, string>>((acc, field) => {
      acc[field.label] = row.values[field.key];
      return acc;
    }, {}),
  );

  const sheet = XLSX.utils.json_to_sheet(data, { header: template.fields.map((field) => field.label) });
  const workbook = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(workbook, sheet, "预览数据");
  XLSX.writeFile(workbook, `preview-${Date.now()}.xlsx`);
}

export function buildTemplateWorkbook(template: TemplateProfile): void {
  const row = createEmptyRow();
  row.values.externalCode = "ORD-2026-0001";
  row.values.senderName = "张三";
  row.values.senderPhone = "13800138000";
  row.values.senderAddress = "上海市浦东新区世纪大道100号";
  row.values.receiverName = "李四";
  row.values.receiverPhone = "13900139000";
  row.values.receiverAddress = "北京市朝阳区建国路88号";
  row.values.weight = "5.5";
  row.values.quantity = "2";
  row.values.temperature = "冷藏";
  row.values.remark = "示例备注";
  exportRowsToWorkbook([row], template);
}

export function preparePreferredSheet(sheetNames: string[]): string {
  return sheetNames.find((name) => !/说明|readme|instruction/i.test(name)) || sheetNames[0] || "";
}

export function buildSubmitPayload(rows: PreviewRow[], template: TemplateProfile): PersistedRecord[] {
  return rows.map((row) => ({
    id: createId("record"),
    templateId: template.id,
    templateName: template.name,
    submittedAt: formatDateTime(),
    externalCode: normalizeText(row.values.externalCode),
    receiverName: normalizeText(row.values.receiverName),
    receiverPhone: normalizeText(row.values.receiverPhone),
    senderName: normalizeText(row.values.senderName),
    senderPhone: normalizeText(row.values.senderPhone),
    senderAddress: normalizeText(row.values.senderAddress),
    receiverAddress: normalizeText(row.values.receiverAddress),
    weight: normalizeText(row.values.weight),
    quantity: normalizeText(row.values.quantity),
    temperature: normalizeText(row.values.temperature),
    remark: normalizeText(row.values.remark),
  }));
}
