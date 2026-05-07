# 物流批量下单导入系统

基于 `Next.js App Router + TypeScript` 实现的物流批量下单 Web 应用，支持多种 Excel 模板自动识别、手动列映射、映射规则记忆、在线预览编辑、全量错误校验、批量提交下单，以及历史运单数据库查询。

## 功能概览

- 支持上传 `.xlsx / .xls`
- 支持多模板自动识别
- 支持不同列名、不同列序、说明行、多 Sheet、分组表头
- 支持手动列映射
- 支持模板映射规则记忆
- 支持 1000+ 行分块导入进度
- 支持预览表格在线编辑
- 支持必填、格式、范围、重复值全量校验
- 支持导出当前预览数据为 Excel
- 支持批量提交并写入数据库
- 支持历史运单搜索、筛选、分页查看

## 技术栈

- Next.js 15
- React 19
- TypeScript
- XLSX
- PostgreSQL
- Vercel

## 本地开发

### 1. 安装依赖

```bash
npm install
```

### 2. 配置环境变量

在项目根目录创建 `.env.local`：

```env
DATABASE_URL=你的 PostgreSQL 连接串
DB_TABLE=waybill_records
```

`DB_TABLE` 不填时默认使用 `waybill_records`。

如果你使用的是 Vercel 环境变量，请在项目设置里添加同名变量，然后重新部署。

### 3. 初始化数据库

在数据库中执行：

[supabase/schema.sql](/C:/Users/ztocc/import-export-tool/supabase/schema.sql)

或直接执行以下 SQL：

```sql
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

create index if not exists idx_waybill_records_external_code
on public.waybill_records (externalCode);

create index if not exists idx_waybill_records_receiver_name
on public.waybill_records (receiverName);

create index if not exists idx_waybill_records_submitted_at
on public.waybill_records (submittedAt desc);
```

### 4. 启动开发环境

```bash
npm run dev
```

默认访问：

```text
http://localhost:3000
```

### 5. 生产构建验证

```bash
npm run build
```

## Vercel 部署

### 1. 推送代码到 Git 仓库

支持：

- GitHub
- GitLab
- Gitee

### 2. 在 Vercel 导入项目

Vercel 中选择仓库后，框架会自动识别为 `Next.js`。

### 3. 配置环境变量

在 Vercel 项目设置中添加：

```env
SUPABASE_URL
SUPABASE_ANON_KEY
SUPABASE_TABLE
```

如果你使用 Vercel Marketplace 集成数据库，也需要确认对应环境变量已经实际注入到项目。

### 4. 点击部署

部署成功后可获得在线访问地址，例如：

```text
https://your-project.vercel.app
```

## 目录结构

```text
app/
  api/
    records/route.ts
    submit/route.ts
  globals.css
  layout.tsx
  page.tsx
components/
  workbench.tsx
lib/
  api.ts
  constants.ts
  db.ts
  storage.ts
  types.ts
  utils.ts
  workbook.ts
supabase/
  schema.sql
```

## 关键说明

### Excel 模板兼容

当前已针对以下类型做兼容：

- 标准中文表头
- 说明行后接表头
- 英文表头
- 分组表头
- 多 Sheet 模板

已覆盖的常见别名包括：

- `外部编码 / 外部订单号 / 客户单号 / Ref Code`
- `发件人姓名 / 发件人 / 发货人 / Sender`
- `发件人电话 / 发件电话 / 发货电话 / Sender Tel`
- `收件人姓名 / 收件人 / 收货人姓名 / Receiver`
- `收件人电话 / 收件电话 / 收货人电话 / Receiver Tel`

### 数据库要求

系统当前要求必须配置数据库环境变量后才能成功提交。

如果未配置：

- 历史运单接口会报错
- 提交接口会返回数据库未配置错误

这符合“运单数据写入数据库而非本地存储”的要求。

## 提交要求

- 在线地址：Vercel 部署后的可访问 URL
- 源码仓库：Git 仓库地址

## 已完成验证

- `Next.js App Router + TypeScript` 已落地
- 本地 `npm run build` 已通过
- 前后端 API 已接通
- Supabase 建表 SQL 已提供
