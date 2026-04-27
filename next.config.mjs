import { execSync } from "node:child_process";

function getBuildNumber() {
  try {
    return execSync("git rev-list --count HEAD", { encoding: "utf8" }).trim();
  } catch {
    return "0";
  }
}

/** @type {import('next').NextConfig} */
const nextConfig = {
  reactStrictMode: true,
  env: {
    NEXT_PUBLIC_BUILD_NUMBER: getBuildNumber()
  }
};

export default nextConfig;
