# Gerador de Descritivos V3 - SENAI Bahia

Aplicação Next.js para Vercel com layout moderno, login Firebase e API Anthropic protegida.

## Recursos

- Login com Firebase Authentication por e-mail/senha
- Geração de Descritivo de Curso
- Geração de Ficha de Produto
- Upload de referência em DOCX, PDF e TXT
- EJA Profissionalizante como botão simples
- Aprender a Empreender atualizado
- Geração de DOCX no navegador
- Chave Anthropic protegida em rota server-side `/api/generate`

## Variáveis de ambiente na Vercel

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

## Publicação na Vercel

1. Suba o conteúdo desta pasta para a raiz do repositório GitHub.
2. Na Vercel, use Framework Preset: Next.js.
3. Build Command: `npm run build`.
4. Install Command: `npm install`.
5. Output Directory: deixe vazio.
6. Configure as variáveis de ambiente.
7. Faça Redeploy.

## Firebase

1. Abra Firebase Console.
2. Crie ou use um projeto existente.
3. Ative Authentication > Sign-in method > Email/Password.
4. Crie um usuário.
5. Copie o firebaseConfig do app Web para as variáveis `NEXT_PUBLIC_FIREBASE_*`.

## Segurança

Não coloque a chave da Anthropic no frontend. Use apenas `ANTHROPIC_API_KEY` nas variáveis de ambiente da Vercel.
