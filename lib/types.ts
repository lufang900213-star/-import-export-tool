export type FieldType = "text" | "phone" | "number" | "integer" | "enum";

export type FieldKey =
  | "externalCode"
  | "senderName"
  | "senderPhone"
  | "senderAddress"
  | "receiverName"
  | "receiverPhone"
  | "receiverAddress"
  | "weight"
  | "quantity"
  | "temperature"
  | "remark";

export interface TemplateField {
  key: FieldKey;
  label: string;
  description: string;
  required: boolean;
  type: FieldType;
  aliases: string[];
}

export interface TemplateProfile {
  id: string;
  name: string;
  description: string;
  builtIn: boolean;
  fields: TemplateField[];
}

export interface MappingRule {
  id: string;
  templateId: string;
  templateName: string;
  headers: string[];
  fingerprint: string;
  mapping: Partial<Record<FieldKey, string>>;
  savedAt: string;
}

export interface PreviewRow {
  id: string;
  sourceRowNumber: number;
  values: Record<FieldKey, string>;
}

export interface ValidationError {
  rowId: string;
  rowNumber: number;
  fieldKey: FieldKey;
  message: string;
}

export interface ImportIssue {
  message: string;
}

export interface ProgressState {
  label: string;
  current: number;
  total: number;
  percent: number;
}

export interface PersistedRecord extends Record<string, string> {
  id: string;
  templateId: string;
  templateName: string;
  submittedAt: string;
  externalCode: string;
  receiverName: string;
  receiverPhone: string;
}
