import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "파일 업로드",
  description: "엑셀 파일 업로드 시스템",
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="ko">
      <body className="bg-gray-50 min-h-screen">
        <main className="max-w-3xl mx-auto px-4 py-10">{children}</main>
      </body>
    </html>
  );
}
