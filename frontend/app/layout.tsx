import type { Metadata, Viewport } from "next";
import { Geist, Geist_Mono } from "next/font/google";
import "./globals.css";
import PwaRegister from "./pwa-register";

const basePath = process.env.NEXT_PUBLIC_BASE_PATH ?? "";

const geistSans = Geist({
  variable: "--font-geist-sans",
  subsets: ["latin"],
});

const geistMono = Geist_Mono({
  variable: "--font-geist-mono",
  subsets: ["latin"],
});

export const metadata: Metadata = {
  title: "Fichas de Mantenimiento CNM",
  description: "Automatización de fichas técnicas de mantenimiento.",
  applicationName: "Fichas CNM",
  manifest: `${basePath}/manifest.webmanifest`,
  icons: { apple: `${basePath}/icon-v2-192.png` },
  appleWebApp: {
    capable: true,
    statusBarStyle: "default",
    title: "Fichas CNM",
  },
};

export const viewport: Viewport = {
  themeColor: "#1A4A5A",
};

export default function RootLayout({ children }: LayoutProps<"/">) {
  return (
    <html
      lang="es"
      suppressHydrationWarning // <--- AQUÍ ESTÁ LA MAGIA QUE EVITA EL ERROR
      className={`${geistSans.variable} ${geistMono.variable} h-full antialiased`}
    >
      <body className="min-h-full flex flex-col">
        <PwaRegister />
        {children}
      </body>
    </html>
  );
}
