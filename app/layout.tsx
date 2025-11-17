import './globals.css';
import type { Metadata } from 'next';

export const metadata: Metadata = {
  title: 'Excel AI Agent',
  description: 'Update Excel files using natural language (Hindi/English)'
};

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="en">
      <body>{children}</body>
    </html>
  );
}
