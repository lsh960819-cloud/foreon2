/** @type {import('tailwindcss').Config} */
// 바이비(BYB) 관리자 화면 느낌: 파란 포인트색 · 각진 모서리 · 그림자 없음 · Noto Sans KR
const blue = {
  50: "#eef5fd", 100: "#d9e9fb", 200: "#b6d4f6", 300: "#86b8f0", 400: "#5a9ee9",
  500: "#3d8be4", 600: "#2f80e0", 700: "#2468bd", 800: "#1f5599", 900: "#1b467c",
};
export default {
  content: ["./index.html", "./src/**/*.{js,jsx}"],
  theme: {
    extend: {
      colors: { emerald: blue, byb: { bg: "#f0f2f5", line: "#e3e6ea", text: "#333", sub: "#888" } },
      fontFamily: { sans: ['"Noto Sans KR"', '"Malgun Gothic"', "맑은 고딕", "sans-serif"] },
      borderRadius: { md: "2px", lg: "3px", xl: "3px", "2xl": "4px" },
      boxShadow: { sm: "none", DEFAULT: "0 1px 2px rgba(0,0,0,.04)" },
    },
  },
  plugins: [],
};
