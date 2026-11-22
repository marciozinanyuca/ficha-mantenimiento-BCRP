import type { NextConfig } from "next";

const isGitHubPages = process.env.GITHUB_PAGES === "true";

const nextConfig: NextConfig = {
  // Pages necesita export estático; el resto conserva la salida actual usada
  // por Vercel y Azure.
  distDir: process.env.BUILD_DIR || ".next",
  ...(isGitHubPages
    ? {
        output: "export" as const,
        trailingSlash: true,
        basePath: "/ficha-mantenimiento-BCRP",
      }
    : { output: "standalone" as const }),
  allowedDevOrigins: ["127.0.0.1", "localhost", ".app.github.dev"],
};

export default nextConfig;
