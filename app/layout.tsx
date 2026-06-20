import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Actualización",
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="es">
      <body className="min-h-screen bg-[#fafaf8] text-[#2C2C2A]">{children}</body>
    </html>
  );
}
