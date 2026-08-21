import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "中国工商银行外部数据产品管理平台",
  description: "商业银行外部数据产品标准化建设管理概念原型",
  icons: {
    icon: "/favicon.svg",
    shortcut: "/favicon.svg",
  },
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
