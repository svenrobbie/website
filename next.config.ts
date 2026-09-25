import type { NextConfig } from 'next';

const nextConfig: NextConfig = {
  agentRules: false,
  output: 'export',
  distDir: 'dist',
  images: { unoptimized: true },
  trailingSlash: true,
  experimental: { useTypeScriptCli: false },
};

export default nextConfig;
