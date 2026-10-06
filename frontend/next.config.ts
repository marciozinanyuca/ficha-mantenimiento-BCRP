import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // TRUCO DE DEVOPS: Si existe la variable BUILD_DIR (Azure), úsala. Si no (Vercel), usa '.next'.
  distDir: process.env.BUILD_DIR || '.next',
  
  output: "standalone",
  
  allowedDevOrigins: ['127.0.0.1', 'localhost', '.app.github.dev'],
};

export default nextConfig;