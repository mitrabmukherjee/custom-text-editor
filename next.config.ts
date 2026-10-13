import type { NextConfig } from "next";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { parseAllowedOrigins } from "./lib/bridge/origins";

const projectRoot = path.dirname(fileURLToPath(import.meta.url));

const allowedParentOrigins = parseAllowedOrigins(
  process.env.NEXT_PUBLIC_ALLOWED_PARENT_ORIGINS,
);

const nextConfig: NextConfig = {
  // C:\Users\Steora has its own package-lock.json; without this Next infers
  // the home directory as the workspace root.
  turbopack: {
    root: projectRoot,
  },
  async headers() {
    return [
      {
        source: "/embed",
        headers: [
          {
            key: "Content-Security-Policy",
            value: `frame-ancestors 'self' ${allowedParentOrigins.join(" ")}`,
          },
        ],
      },
    ];
  },
};

export default nextConfig;
