import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // Le decimos a Next.js que confíe en la red de Codespaces
  allowedDevOrigins: ['127.0.0.1', 'localhost', '.app.github.dev'],
};

export default nextConfig;
