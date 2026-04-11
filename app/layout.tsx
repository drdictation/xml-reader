import type { Metadata } from "next";
import "./globals.css";

export const metadata: Metadata = {
  title: "Genie XML Reader",
  description: "Read-only browser viewer for Genie XML patient exports.",
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
