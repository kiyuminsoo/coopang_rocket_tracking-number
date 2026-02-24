import "./globals.css";

export const metadata = {
  title: "밀크런 운송장 파서",
  description: "밀크런 운송장 PDF에서 FC와 MRB 운송장번호를 추출합니다."
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ko">
      <body>{children}</body>
    </html>
  );
}
