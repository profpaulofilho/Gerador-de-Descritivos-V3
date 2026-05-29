'use client'

import { useEffect, useMemo, useState } from 'react'
import { useRouter } from 'next/navigation'
import { onAuthStateChanged, signOut, User } from 'firebase/auth'
import { auth } from '../../lib/firebase'
import { Document, Packer, Paragraph, TextRun, HeadingLevel, AlignmentType, Table, TableRow, TableCell, WidthType, BorderStyle, ShadingType } from 'docx'
import * as mammoth from 'mammoth'

type Mode = 'descritivo' | 'ficha'
type LogItem = { text:string, type?: 'ok'|'err'|'info' }

const defaultDesc = {
  nomeCurso:'', cargaHoraria:'', cbo:'', modalidade:'Qualificação Profissional', eixo:'Informação e Comunicação', escolaridade:'Ensino Médio incompleto', idadeMinima:'16', cliente:'', requisitos:'', moduloBasico:false, aprenderEmpreender:true, eja:false, numUC:'', ucs:[] as {nome:string,ch:string}[]
}
const defaultFicha = { nome:'', ch:'', modalidade:'Aperfeiçoamento Profissional', eixo:'Informação e Comunicação', publico:'', certificado:'', frequencia:'75%' }

async function extractText(file: File){
  const name = file.name.toLowerCase()
  if(name.endsWith('.txt')) return await file.text()
  if(name.endsWith('.docx')){
    const arrayBuffer = await file.arrayBuffer()
    const res = await mammoth.extractRawText({ arrayBuffer })
    return res.value || ''
  }
  if(name.endsWith('.pdf')){
    const pdfjsLib = await import('pdfjs-dist') as any
    pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.mjs`
    const arrayBuffer = await file.arrayBuffer()
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise
    let text = ''
    for(let i=1;i<=Math.min(pdf.numPages,20);i++){
      const page = await pdf.getPage(i)
      const content = await page.getTextContent()
      text += content.items.map((it:any)=>it.str).join(' ') + '\n'
    }
    return text
  }
  return ''
}

function run(text:string,bold=false,size=22,color='111827'){ return new TextRun({ text, bold, size, color, font:'Arial' }) }
function para(text:string,bold=false){ return new Paragraph({ spacing:{before:80,after:80}, alignment:AlignmentType.JUSTIFIED, children:[run(text,bold)] }) }
function title(text:string){ return new Paragraph({ heading:HeadingLevel.HEADING_1, spacing:{before:220,after:120}, children:[run(text,true,26,'1F3864')] }) }
function bullet(text:string){ return new Paragraph({ bullet:{level:0}, spacing:{before:40,after:40}, children:[run(text,false,20)] }) }
function tableRows(rows: [string,string][]) {
  return new Table({ width:{size:100,type:WidthType.PERCENTAGE}, rows: rows.map(([k,v])=>new TableRow({ children:[
    new TableCell({ width:{size:32,type:WidthType.PERCENTAGE}, shading:{fill:'E8EEF8',type:ShadingType.CLEAR}, borders:{top:{style:BorderStyle.SINGLE,size:1,color:'1F3864'},bottom:{style:BorderStyle.SINGLE,size:1,color:'1F3864'},left:{style:BorderStyle.SINGLE,size:1,color:'1F3864'},right:{style:BorderStyle.SINGLE,size:1,color:'1F3864'}}, children:[new Paragraph({children:[run(k,true,20,'1F3864')]})] }),
    new TableCell({ width:{size:68,type:WidthType.PERCENTAGE}, borders:{top:{style:BorderStyle.SINGLE,size:1,color:'1F3864'},bottom:{style:BorderStyle.SINGLE,size:1,color:'1F3864'},left:{style:BorderStyle.SINGLE,size:1,color:'1F3864'},right:{style:BorderStyle.SINGLE,size:1,color:'1F3864'}}, children:[new Paragraph({children:[run(v||'—',false,20)]})] })
  ]})) })
}

async function buildDescritivoDocx(ai:any, form:any){
  const children:any[] = [
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{after:120}, children:[run('SENAI BAHIA',true,38,'1F3864')] }),
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{after:260}, children:[run('DESCRITIVO DE CURSO',true,30,'1F3864')] }),
    title('1. IDENTIFICAÇÃO'), tableRows([['Título do curso',form.nomeCurso],['Ocupação/CBO',form.cbo],['Modalidade',form.modalidade],['Eixo tecnológico',form.eixo],['Cliente',form.cliente],['Carga horária',form.cargaHoraria]]),
    title('2. JUSTIFICATIVA'), para(ai.justificativa||''), title('3. OBJETIVO'), para(ai.objetivo||''), title('4. DESCRIÇÃO DO CURSO'), para(ai.descricao||''), title('5. PÚBLICO-ALVO'), para(ai.publicoAlvo||''),
    title('6. REQUISITOS DE ACESSO'), para(`Escolaridade mínima: ${form.escolaridade}. Idade mínima: ${form.eja ? '18' : form.idadeMinima} anos. ${form.requisitos||''}`),
    title('7. PERFIL PROFISSIONAL DE CONCLUSÃO'), ...(ai.perfilSaida||[]).map((x:string)=>bullet(x)),
    title('8. ORGANIZAÇÃO CURRICULAR')
  ]
  ;(ai.modulos||[]).forEach((m:any)=>{
    children.push(new Paragraph({ spacing:{before:160,after:80}, children:[run(m.nome + (m.chModulo?` — ${m.chModulo}`:''),true,24,'1F3864')] }))
    ;(m.ucs||[]).forEach((u:any)=>{
      children.push(new Paragraph({ spacing:{before:100,after:60}, children:[run(`${u.nome} — ${u.ch||''}`,true,22)] }))
      children.push(para(u.objetivo||''))
      children.push(new Paragraph({ children:[run('Capacidades:',true,20,'1F3864')] }))
      ;(u.capacidades||[]).forEach((c:string)=>children.push(bullet(c)))
      children.push(new Paragraph({ children:[run('Conhecimentos:',true,20,'1F3864')] }))
      ;(u.conhecimentos||[]).forEach((c:string)=>children.push(bullet(c)))
    })
  })
  children.push(title('9. METODOLOGIA'), para(ai.metodologia||''), title('10. CRITÉRIOS DE AVALIAÇÃO'), para(ai.criteriosAvaliacao||''), title('11. CRITÉRIOS DE CERTIFICAÇÃO'), para(ai.criteriosCertificacao||''), title('12. PERFIL DOCENTE'), para(ai.perfilDocente||''), title('13. INFRAESTRUTURA MÍNIMA'))
  const infra = ai.infraestrutura || {}
  ;[...(infra.ambiente||[]), ...(infra.softwares||[]), ...(infra.epis||[]), ...(infra.conectividade||[])].forEach((i:string)=>children.push(bullet(i)))
  if(infra.materiais) children.push(para(infra.materiais))
  if(ai.eja?.aplicar){ children.push(title('14. DIRETRIZES ESPECÍFICAS DO EJA PROFISSIONALIZANTE'), para(ai.eja.diretrizes||'')); (ai.eja.atividadesAssincronas||[]).forEach((a:string)=>children.push(bullet(a))); (ai.eja.permanenciaExito||[]).forEach((a:string)=>children.push(bullet(a))) }
  children.push(title('REFERÊNCIAS')); (ai.referencias||[]).forEach((r:string)=>children.push(bullet(r)))
  const doc = new Document({ sections:[{ properties:{ page:{ margin:{top:1134,right:1440,bottom:1134,left:1440} } }, children }] })
  return await Packer.toBlob(doc)
}

async function buildFichaDocx(ai:any, form:any){
  const children:any[] = [
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{after:80}, children:[run('SENAI BAHIA',true,34,'1F3864')] }),
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{after:200}, children:[run('FICHA DE PRODUTO',true,28,'1F3864')] }),
    new Paragraph({ alignment:AlignmentType.CENTER, spacing:{after:180}, children:[run(ai.nome || form.nome,true,30,'1F3864')] }),
    tableRows([['Carga horária',ai.cargaHoraria||form.ch],['Modalidade',form.modalidade],['Eixo tecnológico',form.eixo],['Participantes',ai.participantes||form.publico],['Frequência mínima',ai.frequencia||form.frequencia],['Certificado',ai.certificado||form.certificado]]),
    title('Descrição do Curso'), para(ai.descricao||''), title('Programação')
  ]
  ;(ai.programacao||[]).forEach((m:any)=>{ children.push(new Paragraph({ spacing:{before:120,after:60}, children:[run(m.modulo,true,23,'1F3864')] })); (m.topicos||[]).forEach((t:string)=>children.push(bullet(t))) })
  const doc = new Document({ sections:[{ properties:{ page:{ margin:{top:1134,right:1440,bottom:1134,left:1440} } }, children }] })
  return await Packer.toBlob(doc)
}

export default function AppPage(){
  const router=useRouter(); const [user,setUser]=useState<User|null>(null); const [checking,setChecking]=useState(true)
  const [mode,setMode]=useState<Mode>('descritivo'); const [desc,setDesc]=useState(defaultDesc); const [ficha,setFicha]=useState(defaultFicha)
  const [documentText,setDocumentText]=useState(''); const [documentName,setDocumentName]=useState(''); const [logs,setLogs]=useState<LogItem[]>([]); const [busy,setBusy]=useState(false); const [downloadUrl,setDownloadUrl]=useState(''); const [downloadName,setDownloadName]=useState('')
  const valid = mode==='descritivo' ? !!desc.nomeCurso && !!desc.cargaHoraria : !!ficha.nome && !!ficha.ch
  const currentTitle = mode==='descritivo' ? 'Gerador de Descritivos' : 'Ficha de Produto'

  useEffect(()=>{ const unsub=onAuthStateChanged(auth,u=>{ if(!u) router.replace('/login'); else setUser(u); setChecking(false) }); return()=>unsub() },[router])
  function addLog(text:string,type?:LogItem['type']){ setLogs(prev=>[...prev,{text,type}]) }
  async function handleFile(file?:File){ if(!file) return; setDocumentName(file.name); addLog(`Extraindo texto de ${file.name}...`,'info'); const txt=await extractText(file); setDocumentText(txt); addLog(`Documento carregado com ${txt.length} caracteres.`,'ok') }
  async function generate(){
    if(!valid) return; setBusy(true); setLogs([]); setDownloadUrl(''); addLog('Preparando dados...','info')
    try{
      const form = mode==='descritivo' ? desc : ficha
      addLog('Chamando API segura da Vercel...','info')
      const res = await fetch('/api/generate',{ method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({ mode, form, documentText }) })
      const json = await res.json(); if(!res.ok) throw new Error(json.error || 'Erro na geração')
      addLog('Resposta recebida. Construindo DOCX...','ok')
      const blob = mode==='descritivo' ? await buildDescritivoDocx(json.result, desc) : await buildFichaDocx(json.result, ficha)
      const url = URL.createObjectURL(blob); const base = (mode==='descritivo' ? desc.nomeCurso : ficha.nome).replace(/[^a-zA-Z0-9\s]/g,'').replace(/\s+/g,'-')
      setDownloadUrl(url); setDownloadName(`${mode==='descritivo'?'Descritivo':'Ficha-Produto'}-${base}.docx`); addLog('Documento gerado com sucesso!','ok')
    }catch(e:any){ addLog(e.message || 'Erro ao gerar.','err') } finally{ setBusy(false) }
  }

  if(checking) return <main className="page-bg" style={{display:'grid',placeItems:'center'}}><div className="glass card">Carregando...</div></main>
  return <main className="page-bg">
    <header className="topbar"><div className="brand"><img src="/senai-logo.svg" alt="SENAI"/><div><div className="brand-title">SENAI BAHIA</div><div className="brand-sub">Gerador de Descritivos & Fichas V3</div></div></div><div style={{display:'flex',gap:10,alignItems:'center'}}><span className="tag">{user?.email}</span><button className="btn btn-ghost" onClick={()=>signOut(auth)}>Sair</button></div></header>
    <div className="app-wrap"><div className="app-grid">
      <aside className="sidebar glass card"><div className="section-title">Módulos</div><div className={`navitem ${mode==='descritivo'?'active':''}`} onClick={()=>setMode('descritivo')}>📄 Descritivo</div><div className={`navitem ${mode==='ficha'?'active':''}`} onClick={()=>setMode('ficha')}>🗂️ Ficha de Produto</div><div style={{height:1,background:'var(--line)',margin:'18px 0'}}/><div className="tag">🔐 Claude via API segura</div><p className="muted" style={{fontSize:13,lineHeight:1.6}}>A chave Anthropic fica protegida nas variáveis da Vercel.</p></aside>
      <section>
        <div style={{display:'flex',justifyContent:'space-between',alignItems:'end',gap:12,marginBottom:18,flexWrap:'wrap'}}><div><div className="tag">{mode==='descritivo'?'Descritivo de Curso':'Ficha de Produto'}</div><h1 style={{fontSize:34,letterSpacing:'-.04em',margin:'12px 0 4px'}}>{currentTitle}</h1><p className="muted" style={{margin:0}}>Base V2 migrada para Vercel, com login Firebase e layout moderno.</p></div><div className="tabs"><button className={`tab ${mode==='descritivo'?'active':''}`} onClick={()=>setMode('descritivo')}>Descritivo</button><button className={`tab ${mode==='ficha'?'active':''}`} onClick={()=>setMode('ficha')}>Ficha</button></div></div>
        <div className="glass card">
          <div className="section-title">Dados principais</div>
          {mode==='descritivo' ? <>
            <div className="field"><label>Nome do curso *</label><input value={desc.nomeCurso} onChange={e=>setDesc({...desc,nomeCurso:e.target.value})} placeholder="Ex: Assistente em Tecnologias da Indústria 4.0"/></div>
            <div className="grid2"><div className="field"><label>Carga horária *</label><input value={desc.cargaHoraria} onChange={e=>setDesc({...desc,cargaHoraria:e.target.value})} placeholder="Ex: 200h"/></div><div className="field"><label>CBO/Ocupação</label><input value={desc.cbo} onChange={e=>setDesc({...desc,cbo:e.target.value})} placeholder="Opcional"/></div></div>
            <div className="grid2"><div className="field"><label>Modalidade</label><select value={desc.modalidade} onChange={e=>setDesc({...desc,modalidade:e.target.value})}><option>Qualificação Profissional</option><option>Aprendizagem Industrial</option><option>Aperfeiçoamento Profissional</option><option>Habilitação Profissional</option></select></div><div className="field"><label>Eixo tecnológico</label><select value={desc.eixo} onChange={e=>setDesc({...desc,eixo:e.target.value})}>{['Ambiente e Saúde','Controle e Processos Industriais','Gestão e Negócios','Informação e Comunicação','Infraestrutura','Produção Alimentícia','Produção Cultural e Design','Produção Industrial','Recursos Naturais','Segurança','Turismo, Hospitalidade e Lazer'].map(x=><option key={x}>{x}</option>)}</select></div></div>
            <div className="grid3"><div className="field"><label>Escolaridade mínima</label><input value={desc.escolaridade} onChange={e=>setDesc({...desc,escolaridade:e.target.value})}/></div><div className="field"><label>Idade mínima</label><input value={desc.idadeMinima} onChange={e=>setDesc({...desc,idadeMinima:e.target.value})}/></div><div className="field"><label>Cliente</label><input value={desc.cliente} onChange={e=>setDesc({...desc,cliente:e.target.value})}/></div></div>
            <div className="field"><label>Outros requisitos</label><input value={desc.requisitos} onChange={e=>setDesc({...desc,requisitos:e.target.value})} /></div>
            <div className="section-title" style={{marginTop:8}}>Atribuições e módulos</div>
            <Toggle title="Módulo Básico" desc="Sustentabilidade, SST, Qualidade e Indústria 4.0 (+40h)" checked={desc.moduloBasico} onChange={v=>setDesc({...desc,moduloBasico:v})}/>
            <Toggle title="Aprender a Empreender" desc="Conteúdo atualizado com empreendedorismo digital, mobile, marketing, finanças e direitos trabalhistas." checked={desc.aprenderEmpreender} onChange={v=>setDesc({...desc,aprenderEmpreender:v})}/>
            <Toggle title="EJA Profissionalizante" desc="Apenas atribuir modalidade EJA. Os parâmetros ficam internos no prompt: 18 anos, 80% presencial, 20% assíncrono, MRS e permanência." checked={desc.eja} onChange={v=>setDesc({...desc,eja:v,idadeMinima:v?'18':desc.idadeMinima})}/>
            <div className="grid2"><div className="field"><label>Quantidade de UCs técnicas (opcional)</label><input type="number" min="1" max="12" value={desc.numUC} onChange={e=>setDesc({...desc,numUC:e.target.value})} placeholder="Ex: 3"/></div><div className="field"><label>Documento de referência</label><input type="file" accept=".docx,.pdf,.txt" onChange={e=>handleFile(e.target.files?.[0])}/></div></div>
          </> : <>
            <div className="field"><label>Nome do curso/produto *</label><input value={ficha.nome} onChange={e=>setFicha({...ficha,nome:e.target.value})} placeholder="Ex: Criação de Prompts e Agentes de IA Especialistas"/></div>
            <div className="grid2"><div className="field"><label>Carga horária *</label><input value={ficha.ch} onChange={e=>setFicha({...ficha,ch:e.target.value})} placeholder="Ex: 80 horas"/></div><div className="field"><label>Modalidade</label><select value={ficha.modalidade} onChange={e=>setFicha({...ficha,modalidade:e.target.value})}><option>Aperfeiçoamento Profissional</option><option>Qualificação Profissional</option><option>Aprendizagem Industrial</option><option>Habilitação Profissional</option></select></div></div>
            <div className="grid2"><div className="field"><label>Eixo tecnológico</label><select value={ficha.eixo} onChange={e=>setFicha({...ficha,eixo:e.target.value})}>{['Ambiente e Saúde','Controle e Processos Industriais','Gestão e Negócios','Informação e Comunicação','Infraestrutura','Produção Alimentícia','Produção Cultural e Design','Produção Industrial','Recursos Naturais','Segurança','Turismo, Hospitalidade e Lazer'].map(x=><option key={x}>{x}</option>)}</select></div><div className="field"><label>Frequência mínima</label><input value={ficha.frequencia} onChange={e=>setFicha({...ficha,frequencia:e.target.value})}/></div></div>
            <div className="field"><label>Público-alvo</label><input value={ficha.publico} onChange={e=>setFicha({...ficha,publico:e.target.value})}/></div><div className="field"><label>Certificado emitido</label><input value={ficha.certificado} onChange={e=>setFicha({...ficha,certificado:e.target.value})}/></div>
            <div className="field"><label>Adicionar documento de referência</label><input type="file" accept=".docx,.pdf,.txt" onChange={e=>handleFile(e.target.files?.[0])}/></div>
          </>}
          {documentName && <div style={{marginTop:10}} className="tag">📎 {documentName} · {documentText.length} caracteres extraídos</div>}
          <div style={{display:'flex',justifyContent:'space-between',gap:12,marginTop:24,flexWrap:'wrap'}}><button className="btn btn-ghost" onClick={()=>{setDocumentText('');setDocumentName('');setLogs([]);setDownloadUrl('')}}>Limpar</button><button className="btn btn-primary" disabled={!valid||busy} onClick={generate}>{busy?'Gerando...':'⚡ Gerar documento'}</button></div>
        </div>
        {(logs.length>0 || downloadUrl) && <div className="glass card" style={{marginTop:18}}><div className="section-title">Resultado</div><div className="log">{logs.map((l,i)=><div key={i} className={l.type}>{'> '+l.text}</div>)}</div>{downloadUrl && <a className="btn btn-primary" style={{marginTop:16}} href={downloadUrl} download={downloadName}>📥 Baixar {downloadName}</a>}</div>}
      </section>
    </div></div>
  </main>
}

function Toggle({title,desc,checked,onChange}:{title:string,desc:string,checked:boolean,onChange:(v:boolean)=>void}){ return <div className="switch-row"><div><div className="switch-title">{title}</div><div className="switch-desc">{desc}</div></div><label className="switch"><input type="checkbox" checked={checked} onChange={e=>onChange(e.target.checked)}/><span className="knob"/></label></div> }
