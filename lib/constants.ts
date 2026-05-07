import type { TemplateField, TemplateProfile } from "./types";

export const STORAGE_KEYS = {
  templateId: "waybill-workbench-template-id",
  previewRows: "waybill-workbench-preview-rows",
  mappingRules: "waybill-workbench-mapping-rules",
} as const;

export const TEMPERATURE_OPTIONS = ["常温", "冷藏", "冷冻"] as const;

function field(
  key: TemplateField["key"],
  label: string,
  description: string,
  required: boolean,
  type: TemplateField["type"],
  aliases: string[],
): TemplateField {
  return { key, label, description, required, type, aliases };
}

export const SYSTEM_FIELDS: TemplateField[] = [
  field("externalCode", "外部编码", "外部系统订单唯一编号，用于去重", false, "text", [
    "外部编码",
    "外部订单号",
    "客户单号",
    "客户订单号",
    "订单编号",
    "Ref Code",
    "Ref",
  ]),
  field("senderName", "发件人姓名", "寄件人姓名", true, "text", ["发件人姓名", "发件人", "寄件人姓名", "发货人", "Sender"]),
  field("senderPhone", "发件人电话", "寄件人联系方式", true, "phone", [
    "发件人电话",
    "发件电话",
    "寄件人电话",
    "发货电话",
    "Sender Tel",
    "Sender Phone",
  ]),
  field("senderAddress", "发件人地址", "寄件人完整地址", true, "text", [
    "发件人地址",
    "发件地址",
    "寄件人地址",
    "发货地址",
    "Sender Address",
  ]),
  field("receiverName", "收件人姓名", "收货人姓名", true, "text", ["收件人姓名", "收货人姓名", "收方", "Receiver", "收件人"]),
  field("receiverPhone", "收件人电话", "收货人联系方式", true, "phone", [
    "收件人电话",
    "收件电话",
    "收货人电话",
    "收方电话",
    "Receiver Tel",
    "Receiver Phone",
  ]),
  field("receiverAddress", "收件人地址", "收货人完整地址", true, "text", [
    "收件人地址",
    "收件地址",
    "收货人地址",
    "收方地址",
    "Receiver Address",
  ]),
  field("weight", "重量(kg)", "货物重量，必须为正数", true, "number", ["重量", "重量(kg)", "Weight", "Weight(kg)"]),
  field("quantity", "件数", "包裹数量，必须为正整数", true, "integer", ["件数", "数量", "Qty", "Package Count"]),
  field("temperature", "温层", "常温 / 冷藏 / 冷冻（可选值之一）", true, "enum", ["温层", "温度要求", "温控", "Temp Zone"]),
  field("remark", "备注", "附加说明", false, "text", ["备注", "附言", "Note"]),
];

export const TEMPLATE_PROFILES: TemplateProfile[] = [
  {
    id: "standard",
    name: "标准模板",
    description: "第一行即表头，列名为常见中文字段。",
    builtIn: true,
    fields: SYSTEM_FIELDS,
  },
  {
    id: "ecommerce",
    name: "电商模板",
    description: "支持说明行、不同列序及中文/英文混排表头。",
    builtIn: true,
    fields: SYSTEM_FIELDS,
  },
  {
    id: "english",
    name: "英文模板",
    description: "支持 Receiver / Sender / Weight 等英文列名。",
    builtIn: true,
    fields: SYSTEM_FIELDS,
  },
  {
    id: "grouped",
    name: "分组模板",
    description: "支持发件方 / 收件方 / 货物信息的分组表头。",
    builtIn: true,
    fields: SYSTEM_FIELDS,
  },
  {
    id: "multi-sheet",
    name: "多 Sheet 模板",
    description: "支持说明页 + 数据页结构。",
    builtIn: true,
    fields: SYSTEM_FIELDS,
  },
];
