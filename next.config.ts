import type { NextConfig } from "next";

const nextConfig: NextConfig = {
  // Google Drive 이미지 도메인 허용
  images: {
    remotePatterns: [
      { protocol: "https", hostname: "drive.google.com" },
      { protocol: "https", hostname: "lh3.googleusercontent.com" },
    ],
  },
  // 서버사이드 전용 모듈 (번들러 겁데기 방지)
  serverExternalPackages: ["googleapis"],
};

export default nextConfig;
