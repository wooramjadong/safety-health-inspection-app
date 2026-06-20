import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "안전보건 점검관리",
  description: "정기안전보건평가 및 중처법준수평가 반자동화 웹앱"
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="ko">
      <body>{children}</body>
    </html>
  );
}
