import type { Metadata } from 'next';
import '@fontsource-variable/inter';
import '@fontsource-variable/jetbrains-mono';
import { Header } from '@/components/Header';
import { Footer } from '@/components/Footer';
import { PageProgress } from '@/components/PageProgress';
import './globals.css';

export const metadata: Metadata = {
  title: { default: 'Sven van de Lagemaat | Cybersecurity and DevOps', template: '%s | Sven van de Lagemaat' },
  description: 'Portfolio of Sven van de Lagemaat, a cybersecurity student and programmer focused on DevOps, secure coding, and self-hosted infrastructure.',
};

export default function RootLayout({ children }: Readonly<{ children: React.ReactNode }>) {
  return (
    <html lang="en" data-scroll-behavior="smooth">
      <body>
        <PageProgress />
        <Header />
        <main>{children}</main>
        <Footer />
      </body>
    </html>
  );
}
