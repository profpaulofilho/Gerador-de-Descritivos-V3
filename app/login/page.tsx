'use client'

import { useEffect, useState } from 'react'
import { useRouter } from 'next/navigation'
import { onAuthStateChanged, signInWithEmailAndPassword } from 'firebase/auth'
import { auth } from '../../lib/firebase'

export default function LoginPage() {
  const router = useRouter()
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (user) => {
      if (user) router.replace('/app')
    })
    return () => unsub()
  }, [router])

  async function handleLogin(event: React.FormEvent<HTMLFormElement>) {
    event.preventDefault()
    setLoading(true)
    setError('')

    try {
      await signInWithEmailAndPassword(auth, email, password)
      router.replace('/app')
    } catch {
      setError('E-mail ou senha inválidos. Confira o usuário criado no Firebase Authentication.')
    } finally {
      setLoading(false)
    }
  }

  return (
    <main className="login-page">
      <div className="login-shell">
        <section className="login-hero">
          <div className="brand-card">
            <img src="/senai-logo.svg" alt="SENAI" />
            <div>
              <strong>SENAI BAHIA</strong>
              <span>Gerador de Descritivos V3</span>
            </div>
          </div>

          <div className="eyebrow">IA institucional segura</div>
         <h1>Descritivos e fichas de produto <span>em padrão SENAI</span></h1>
          <p>
            Uma interface moderna para gerar documentos pedagógicos com
            upload de referências, EJA Profissionalizante e conteúdo atualizado de
            Aprender a Empreender.
          </p>

          <div className="feature-grid compact">
            <div><b>🔐</b><span>Login Firebase</span></div>
            <div><b>🧠</b><span>Claude via API segura</span></div>
            <div><b>📎</b><span>DOCX, PDF e TXT</span></div>
            <div><b>🗂️</b><span>Ficha de Produto</span></div>
          </div>
        </section>

        <section className="login-card glass-panel">
          <div className="card-kicker">Acesso restrito</div>
          <h2>Entrar no sistema</h2>
          <p className="muted">Use o e-mail e senha cadastrados no Firebase.</p>

          <form onSubmit={handleLogin} className="form-stack">
            <label>
              <span>E-mail</span>
              <input
                type="email"
                required
                placeholder="nome@senai.br"
                value={email}
                onChange={(event) => setEmail(event.target.value)}
              />
            </label>

            <label>
              <span>Senha</span>
              <input
                type="password"
                required
                placeholder="••••••••"
                value={password}
                onChange={(event) => setPassword(event.target.value)}
              />
            </label>

            {error && <div className="alert error">{error}</div>}

            <button className="primary-btn" disabled={loading} type="submit">
              {loading ? 'Entrando...' : 'Entrar no Hub →'}
            </button>
          </form>

          <div className="security-note">
            <span>✓</span> Ambiente protegido.
          </div>
        </section>
      </div>
    </main>
  )
}
