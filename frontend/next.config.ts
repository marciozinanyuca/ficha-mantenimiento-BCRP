import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // Habilita la compilación standalone optimizada para contenedores Docker
  output: "standalone",
  
  // Mantenemos tu configuración para Codespaces (útil para desarrollo)
  allowedDevOrigins: ['127.0.0.1', 'localhost', '.app.github.dev'],
};

export default nextConfig;
