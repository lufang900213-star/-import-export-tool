import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "物流批量导入工作台",
  description: "支持多模板 Excel 自动识别、在线预览编辑、批量提交与数据库查看。",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh-CN">
      <body>{children}</body>
    </html>
  );
}
