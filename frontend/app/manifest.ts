import type { MetadataRoute } from "next";

export const dynamic = "force-static";

export default function manifest(): MetadataRoute.Manifest {
  const basePath = process.env.NEXT_PUBLIC_BASE_PATH ?? "";

  return {
    id: `${basePath}/`,
    name: "Fichas de Mantenimiento CNM",
    short_name: "Fichas CNM",
    description: "Automatización de fichas técnicas de mantenimiento.",
    start_url: `${basePath}/`,
    scope: `${basePath}/`,
    display: "standalone",
    background_color: "#D5EAD8",
    theme_color: "#D5EAD8",
    icons: [
      { src: `${basePath}/icon-v3-192.png`, sizes: "192x192", type: "image/png", purpose: "any" },
      { src: `${basePath}/icon-v3-512.png`, sizes: "512x512", type: "image/png", purpose: "maskable" },
      { src: `${basePath}/icon.svg`, sizes: "any", type: "image/svg+xml", purpose: "any" },
    ],
  };
}
