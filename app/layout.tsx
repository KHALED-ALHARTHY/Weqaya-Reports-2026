import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "تقارير برامج وقاية",
  description: "إدخال نتائج برامج مكافحة العدوى ومتابعتها وإعداد التقارير حسب الفترة.",
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
    <html lang="ar" dir="rtl">
      <body className="antialiased">{children}</body>
    </html>
  );
}
