import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "万能导入 —— 多模板自动导入下单系统",
  description: "AI编程考试题目实现",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="zh">
      <body className="font-sans">{children}</body>
    </html>
  );
}
