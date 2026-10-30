import type { Metadata } from "next";
import localFont from "next/font/local";
import "./globals.css";

const workspaceFont = localFont({
  src: [
    { path: "./fonts/NotoSansThai-Regular.ttf", weight: "400", style: "normal" },
    { path: "./fonts/NotoSansThai-Bold.ttf", weight: "700", style: "normal" },
  ],
  variable: "--font-workspace",
  display: "swap",
});

export const metadata: Metadata = {
  title: "Price Checker · ตรวจและปรับราคาสินค้า",
  description: "จัดการสินค้า เปรียบเทียบราคา และตรวจสอบข้อมูลของ SRC, KKL และ SSS",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="th">
      <body
        className={`${workspaceFont.variable} antialiased`}
      >
        {children}
      </body>
    </html>
  );
}
