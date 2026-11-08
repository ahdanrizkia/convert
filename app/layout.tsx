import type { Metadata, Viewport } from 'next';
import './globals.css';

export const metadata: Metadata = {
  title: 'TMPCONV // Stateless Ephemeral File Converter',
  description:
    'Zero-database, zero-persistence, ultra-lightweight web file converter with immediate post-download server deletion.',
};

export const viewport: Viewport = {
  width: 'device-width',
  initialScale: 1,
  themeColor: '#0a0a0a',
};

export default function RootLayout({
  children,
}: Readonly<{
  children: React.ReactNode;
}>) {
  return (
    <html lang="en">
      <body className="min-h-screen bg-neutral-50 font-sans text-neutral-950">
        {children}
      </body>
    </html>
  );
}
