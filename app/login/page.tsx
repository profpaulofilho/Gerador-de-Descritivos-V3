'use client'

import { useEffect, useState } from 'react'
import { useRouter } from 'next/navigation'
import {
  onAuthStateChanged,
  signInWithEmailAndPassword,
  updatePassword,
} from 'firebase/auth'
import { auth } from '../../lib/firebase'

const SENHA_PADRAO = 'Senai@2026'

export default function LoginPage() {
  const router = useRouter()
  const [email, setEmail] = useState('')
  const [password, setPassword] = useState('')
  const [newPassword, setNewPassword] = useState('')
  const [confirmPassword, setConfirmPassword] = useState('')
  const [loading, setLoading] = useState(false)
  const [error, setError] = useState('')
  const [step, setStep] = useState<'login' | 'change'>('login')

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (user) => {
      if (user && step === 'login') router.replace('/app')
    })
    return () => unsub()
  }, [router, step])

  async function handleLogin(e: React.FormEvent) {
    e.preventDefault()
    setLoading(true)
    setError('')
    try {
      await signInWithEmailAndPassword(auth, email, password)
      if (password === SENHA_PADRAO) {
        setStep('change')
        setLoading(false)
      } else {
        router.replace('/app')
      }
    } catch {
      setError('E-mail ou senha inválidos.')
      setLoading(false)
    }
  }

  async function handleChangePassword(e: React.FormEvent) {
    e.preventDefault()
    setError('')
    if (newPassword.length < 8) {
      setError('A nova senha deve ter pelo menos 8 caracteres.')
      return
    }
    if (newPassword !== confirmPassword) {
      setError('As senhas não coincidem.')
      return
    }
    if (newPassword === SENHA_PADRAO) {
      setError('A nova senha não pode ser igual à senha padrão.')
      return
    }
    setLoading(true)
    try {
      const user = auth.currentUser
      if (!user) throw new Error('Sessão expirada.')
      await updatePassword(user, newPassword)
      router.replace('/app')
    } catch {
      setError('Erro ao alterar senha. Tente fazer login novamente.')
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
          <div className="eyebrow" style={{ marginTop: 36 }}>IA institucional segura</div>
          <h1>Descritivos e fichas de produto em padrão SENAI</h1>
          <p>
            Gere documentos pedagógicos com upload de referências,
            EJA Profissionalizante e conteúdo atualizado de Aprender a Empreender.
          </p>
        </section>

        <section className="login-card glass-panel">
          {step === 'login' ? (
            <>
              <div className="login-card-eyebrow">Acesso restrito</div>
              <h2>Entrar no sistema</h2>
              <p className="login-card-sub">Use o e-mail e senha do seu cadastro.</p>
              <form onSubmit={handleLogin}>
                <div className="field">
                  <label>E-mail</label>
                  <input
                    type="email"
                    required
                    placeholder="nome@senai.br"
                    value={email}
                    onChange={e => setEmail(e.target.value)}
                  />
                </div>
                <div className="field">
                  <label>Senha</label>
                  <input
                    type="password"
                    required
                    placeholder="••••••••"
                    value={password}
                    onChange={e => setPassword(e.target.value)}
                  />
                </div>
                {error && <div className="error-msg">{error}</div>}
                <button
                  className="primary-btn"
                  disabled={loading}
                  type="submit"
                  style={{ width: '100%', marginTop: 8 }}
                >
                  {loading ? 'Entrando...' : 'Entrar →'}
                </button>
              </form>
            </>
          ) : (
            <>
              <div className="login-card-eyebrow">Primeiro acesso</div>
              <h2>Crie sua nova senha</h2>
              <p className="login-card-sub">
                Por segurança, defina uma senha pessoal para continuar.
              </p>
              <form onSubmit={handleChangePassword}>
                <div className="field">
                  <label>Nova senha</label>
                  <input
                    type="password"
                    required
                    placeholder="Mínimo 8 caracteres"
                    value={newPassword}
                    onChange={e => setNewPassword(e.target.value)}
                  />
                </div>
                <div className="field">
                  <label>Confirmar nova senha</label>
                  <input
                    type="password"
                    required
                    placeholder="Repita a senha"
                    value={confirmPassword}
                    onChange={e => setConfirmPassword(e.target.value)}
                  />
                </div>
                {error && <div className="error-msg">{error}</div>}
                <button
                  className="primary-btn"
                  disabled={loading}
                  type="submit"
                  style={{ width: '100%', marginTop: 8 }}
                >
                  {loading ? 'Salvando...' : 'Salvar e entrar →'}
                </button>
              </form>
            </>
          )}
        </section>

      </div>
    </main>
  )
}
