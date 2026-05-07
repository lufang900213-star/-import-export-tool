create table if not exists public.waybill_records (
  id text primary key,
  templateId text not null,
  templateName text not null,
  submittedAt text not null,
  externalCode text not null default '',
  receiverName text not null default '',
  receiverPhone text not null default '',
  senderName text not null default '',
  senderPhone text not null default '',
  senderAddress text not null default '',
  receiverAddress text not null default '',
  weight text not null default '',
  quantity text not null default '',
  temperature text not null default '',
  remark text not null default ''
);

alter table public.waybill_records
  add column if not exists templateId text not null default '',
  add column if not exists templateName text not null default '',
  add column if not exists submittedAt text not null default '',
  add column if not exists externalCode text not null default '',
  add column if not exists receiverName text not null default '',
  add column if not exists receiverPhone text not null default '',
  add column if not exists senderName text not null default '',
  add column if not exists senderPhone text not null default '',
  add column if not exists senderAddress text not null default '',
  add column if not exists receiverAddress text not null default '',
  add column if not exists weight text not null default '',
  add column if not exists quantity text not null default '',
  add column if not exists temperature text not null default '',
  add column if not exists remark text not null default '';

create index if not exists idx_waybill_records_external_code
on public.waybill_records (externalCode);

create index if not exists idx_waybill_records_receiver_name
on public.waybill_records (receiverName);

create index if not exists idx_waybill_records_submitted_at
on public.waybill_records (submittedAt desc);
