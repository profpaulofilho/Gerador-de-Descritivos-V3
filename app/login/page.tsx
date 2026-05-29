'use client'

import { useEffect, useState } from 'react'
import { useRouter } from 'next/navigation'
import { signInWithEmailAndPassword, onAuthStateChanged } from 'firebase/auth'
import { auth } from '../../lib/firebase'

export default function LoginPage(){
  const router = useRouter()
  const [email,setEmail]=useState('')
  const [password,setPassword]=useState('')
  const [loading,setLoading]=useState(false)
  const [error,setError]=useState('')

  useEffect(()=>{
    const unsub = onAuthStateChanged(auth, user => { if(user) router.replace('/app') })
    return () => unsub()
  },[router])

  async function login(e: React.FormEvent){
    e.preventDefault(); setLoading(true); setError('')
    try{
      await signInWithEmailAndPassword(auth,email,password)
      router.replace('/app')
    }catch(err:any){
      setError('E-mail ou senha inválidos. Verifique o usuário criado no Firebase.')
    }finally{ setLoading(false) }
  }

  return <main className="page-bg" style={{display:'grid',placeItems:'center',padding:24}}>
    <div style={{width:'100%',maxWidth:1100,display:'grid',gridTemplateColumns:'1.1fr .9fr',gap:28,alignItems:'center'}}>
      <section>
        <div className="brand" style={{marginBottom:34}}>
          <img src="/senai-logo.svg" alt="SENAI" />
          <div><div className="brand-title">SENAI BAHIA</div><div className="brand-sub">Gerador de Descritivos V3</div></div>
        </div>
        <h1 className="hero-title">Documentos pedagógicos <span>com IA segura</span></h1>
        <p className="muted" style={{fontSize:18,lineHeight:1.7,maxWidth:620}}>Gere descritivos de curso e fichas de produto no padrão SENAI, com upload de documentos de referência, EJA Profissionalizante e conteúdo atualizado de Aprender a Empreender.</p>
        <div style={{display:'flex',gap:10,flexWrap:'wrap',marginTop:26}}>
          <span className="tag">🔐 Firebase Login</span><span className="tag">🧠 Claude protegido</span><span className="tag">📎 Upload DOCX/PDF/TXT</span><span className="tag">🗂️ Ficha de Produto</span>
        </div>
      </section>
      <section className="glass card">
        <div style={{marginBottom:24}}>
          <div className="tag" style={{marginBottom:14}}>Acesso restrito</div>
          <h2 style={{margin:'0 0 6px',fontSize:30}}>Entrar no sistema</h2>
          <p className="muted" style={{margin:0}}>Use o usuário criado no Firebase Authentication.</p>
        </div>
        <form onSubmit={login}>
          <div className="field"><label>E-mail</label><input type="email" required value={email} onChange={e=>setEmail(e.target.value)} placeholder="nome@senai.br" /></div>
          <div className="field"><label>Senha</label><input type="password" required value={password} onChange={e=>setPassword(e.target.value)} placeholder="••••••••" /></div>
          {error && <div style={{padding:12,borderRadius:12,background:'rgba(248,113,113,.12)',border:'1px solid rgba(248,113,113,.25)',color:'#fecdd3',fontSize:13,marginBottom:14}}>{error}</div>}
          <button className="btn btn-primary" disabled={loading} style={{width:'100%'}}>{loading?'Entrando...':'Entrar →'}</button>
        </form>
        <div style={{marginTop:20,paddingTop:18,borderTop:'1px solid var(--line)',color:'var(--muted2)',fontSize:12,lineHeight:1.7}}>A chave da Anthropic não fica no navegador. A geração usa API segura na Vercel.</div>
      </section>
    </div>
  </main>
}
