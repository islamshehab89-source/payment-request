import type { NextConfig } from "next";

// Static export for GitHub Pages. On Pages the site lives under a sub-path
// (e.g. /payment-request), supplied at build time via NEXT_PUBLIC_BASE_PATH so
// local `next dev` stays at the root.
const basePath = process.env.NEXT_PUBLIC_BASE_PATH || "";

const nextConfig: NextConfig = {
  output: "export",
  trailingSlash: true,
  basePath: basePath || undefined,
  images: { unoptimized: true },
};

export default nextConfig;
