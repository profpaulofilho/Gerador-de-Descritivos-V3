import './globals.css'
import type { Metadata } from 'next'

export const metadata: Metadata = {
  title: 'SENAI Bahia — Gerador de Descritivos V3',
  description: 'Gerador de descritivos e fichas de produto com IA',
  icons: { icon: '/favicon.svg' },
}

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="pt-BR" data-theme="dark">
      <head>
        <link rel="icon" type="image/svg+xml" href="/favicon.svg" />
      </head>
      <body>{children}</body>
    </html>
  )
}
