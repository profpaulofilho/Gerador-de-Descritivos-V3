# Gerador de Descritivos V3 — SENAI Bahia

Aplicação Next.js para Vercel com:

- Login por Firebase Authentication (e-mail/senha)
- API segura server-side para Anthropic/Claude
- Gerador de Descritivos de Curso
- Gerador de Ficha de Produto
- Upload de documentos `.docx`, `.pdf` e `.txt`
- EJA Profissionalizante como atribuição simples, com regras internas no prompt
- Conteúdo atualizado de Aprender a Empreender
- Layout moderno com identidade SENAI

## Variáveis de ambiente

Configure na Vercel em **Project Settings > Environment Variables**:

```env
ANTHROPIC_API_KEY=sk-ant-...
ANTHROPIC_MODEL=claude-sonnet-4-20250514

NEXT_PUBLIC_FIREBASE_API_KEY=
NEXT_PUBLIC_FIREBASE_AUTH_DOMAIN=
NEXT_PUBLIC_FIREBASE_PROJECT_ID=
NEXT_PUBLIC_FIREBASE_STORAGE_BUCKET=
NEXT_PUBLIC_FIREBASE_MESSAGING_SENDER_ID=
NEXT_PUBLIC_FIREBASE_APP_ID=
```

## Rodar localmente

```bash
npm install
cp .env.example .env.local
npm run dev
```

## Deploy

1. Suba estes arquivos no GitHub.
2. Importe o repositório na Vercel.
3. Configure as variáveis de ambiente.
4. Faça Deploy.

## Segurança

A chave da Anthropic não fica no navegador. A chamada ao Claude acontece em `app/api/generate/route.ts`.
