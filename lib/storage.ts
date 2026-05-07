import { STORAGE_KEYS, TEMPLATE_PROFILES } from "./constants";
import type { MappingRule, PreviewRow } from "./types";

function loadJson<T>(key: string, fallback: T): T {
  if (typeof window === "undefined") return fallback;
  try {
    const raw = window.localStorage.getItem(key);
    return raw ? (JSON.parse(raw) as T) : fallback;
  } catch {
    return fallback;
  }
}

export function loadTemplateId(): string {
  if (typeof window === "undefined") return TEMPLATE_PROFILES[0].id;
  return window.localStorage.getItem(STORAGE_KEYS.templateId) || TEMPLATE_PROFILES[0].id;
}

export function persistTemplateId(templateId: string): void {
  if (typeof window === "undefined") return;
  window.localStorage.setItem(STORAGE_KEYS.templateId, templateId);
}

export function loadPreviewRows(): PreviewRow[] {
  return loadJson<PreviewRow[]>(STORAGE_KEYS.previewRows, []);
}

export function persistPreviewRows(rows: PreviewRow[]): void {
  if (typeof window === "undefined") return;
  window.localStorage.setItem(STORAGE_KEYS.previewRows, JSON.stringify(rows));
}

export function loadMappingRules(): MappingRule[] {
  return loadJson<MappingRule[]>(STORAGE_KEYS.mappingRules, []);
}

export function persistMappingRules(rules: MappingRule[]): void {
  if (typeof window === "undefined") return;
  window.localStorage.setItem(STORAGE_KEYS.mappingRules, JSON.stringify(rules));
}
