import './globals.css'
import type { Metadata } from 'next'

export const metadata: Metadata = {
  title: 'SENAI Bahia — Gerador de Descritivos V3',
  description: 'Gerador de descritivos e fichas de produto com IA',
}

export default function RootLayout({ children }: { children: React.ReactNode }) {
  return (
    <html lang="pt-BR">
      <body>{children}</body>
    </html>
  )
}
