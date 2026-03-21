import type { Metadata } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'SonaCore Web',
  description: 'ブラウザで動作する録音・文字起こしアプリ',
};

export default function RootLayout({
  children,
}: {
  children: React.ReactNode;
}) {
  return (
    <html lang="ja">
      <body className="bg-gray-950 text-white antialiased">{children}</body>
    </html>
  );
}
