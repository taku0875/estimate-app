import type { Metadata, Viewport } from "next";
import "./globals.css";

export const metadata: Metadata = { title: "TSUMIKI", description: "見積依頼書(Excel)を作るツール" };
export const viewport: Viewport = { width: "device-width", initialScale: 1, viewportFit: "cover" };

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ja">
      <head>
        <link rel="preconnect" href="https://fonts.googleapis.com" />
        <link rel="preconnect" href="https://fonts.gstatic.com" crossOrigin="" />
        <link rel="stylesheet" href="https://fonts.googleapis.com/css2?family=Noto+Sans+JP:wght@400;500;700;900&family=Orbitron:wght@700;800&display=swap" />
      </head>
      <body>{children}</body>
    </html>
  );
}
