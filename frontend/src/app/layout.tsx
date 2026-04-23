import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'CRM 매출분석 보고서 생성기 | 열정의시간',
  description: '열정의시간 마케팅팀 — CRM 엑셀 파일을 업로드하면 자동으로 PPT 보고서를 생성합니다.',
  icons: {
    icon: '/logo_passion.png',
    apple: '/logo_passion.png',
  },
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="ko">
      <body>{children}</body>
    </html>
  );
}
