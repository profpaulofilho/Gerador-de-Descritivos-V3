'use client'

import { useEffect, useMemo, useState } from 'react'
import { useRouter } from 'next/navigation'
import { onAuthStateChanged, signOut, User } from 'firebase/auth'
import { auth } from '../../lib/firebase'
import {
  AlignmentType,
  BorderStyle,
  Document,
  HeadingLevel,
  Packer,
  Paragraph,
  ShadingType,
  Table,
  TableCell,
  TableRow,
  TextRun,
  WidthType,
} from 'docx'
import * as mammoth from 'mammoth'

type Mode = 'descritivo' | 'ficha'
type LogKind = 'ok' | 'err' | 'info'
type LogItem = { text: string; type?: LogKind }

type DescritivoForm = {
  nomeCurso: string
  cargaHoraria: string
  cbo: string
  modalidade: string
  eixo: string
  escolaridade: string
  idadeMinima: string
  cliente: string
  requisitos: string
  moduloBasico: boolean
  aprenderEmpreender: boolean
  aprenderEmpreenderCh: string
  eja: boolean
  numUC: string
  usarNomesUC: boolean
  ucs: { nome: string; ch: string }[]
}

type FichaForm = {
  nome: string
  ch: string
  modalidade: string
  eixo: string
  publico: string
  certificado: string
  frequencia: string
}

const defaultDesc: DescritivoForm = {
  nomeCurso: '',
  cargaHoraria: '',
  cbo: '',
  modalidade: 'Qualificação Profissional',
  eixo: 'Informação e Comunicação',
  escolaridade: 'Ensino Médio incompleto',
  idadeMinima: '16',
  cliente: '',
  requisitos: '',
  moduloBasico: false,
  aprenderEmpreender: true,
  aprenderEmpreenderCh: '24',
  eja: false,
  numUC: '',
  usarNomesUC: false,
  ucs: [],
}

const defaultFicha: FichaForm = {
  nome: '',
  ch: '',
  modalidade: 'Aperfeiçoamento Profissional',
  eixo: 'Informação e Comunicação',
  publico: '',
  certificado: '',
  frequencia: '75%',
}

const modalidades = [
  'Qualificação Profissional',
  'Aperfeiçoamento Profissional',
  'Aprendizagem Industrial',
  'Habilitação Profissional',
]

const eixos = [
  'Ambiente e Saúde',
  'Controle e Processos Industriais',
  'Gestão e Negócios',
  'Informação e Comunicação',
  'Infraestrutura',
  'Produção Alimentícia',
  'Produção Cultural e Design',
  'Produção Industrial',
  'Recursos Naturais',
  'Segurança',
  'Turismo, Hospitalidade e Lazer',
]

async function extractText(file: File) {
  const name = file.name.toLowerCase()

  if (name.endsWith('.txt')) return await file.text()

  if (name.endsWith('.docx')) {
    const arrayBuffer = await file.arrayBuffer()
    const result = await mammoth.extractRawText({ arrayBuffer })
    return result.value || ''
  }

  if (name.endsWith('.pdf')) {
    const pdfjsLib = (await import('pdfjs-dist')) as any
    pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.mjs`
    const arrayBuffer = await file.arrayBuffer()
    const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise
    let text = ''
    for (let i = 1; i <= Math.min(pdf.numPages, 20); i++) {
      const page = await pdf.getPage(i)
      const content = await page.getTextContent()
      text += content.items.map((item: any) => item.str).join(' ') + '\n'
    }
    return text
  }

  return ''
}

function run(text: string, bold = false, size = 22, color = '111827') {
  return new TextRun({ text, bold, size, color, font: 'Arial' })
}
function para(text: string, bold = false) {
  return new Paragraph({
    spacing: { before: 80, after: 80 },
    alignment: AlignmentType.JUSTIFIED,
    children: [run(text || '—', bold)],
  })
}
function title(text: string) {
  return new Paragraph({
    heading: HeadingLevel.HEADING_1,
    spacing: { before: 220, after: 120 },
    children: [run(text, true, 26, '1F3864')],
  })
}
function bullet(text: string) {
  return new Paragraph({
    bullet: { level: 0 },
    spacing: { before: 40, after: 40 },
    children: [run(text, false, 20)],
  })
}
function infoTable(rows: [string, string][]) {
  const border = { style: BorderStyle.SINGLE, size: 1, color: '1F3864' }
  return new Table({
    width: { size: 100, type: WidthType.PERCENTAGE },
    rows: rows.map(
      ([key, value]) =>
        new TableRow({
          children: [
            new TableCell({
              width: { size: 32, type: WidthType.PERCENTAGE },
              shading: { fill: 'E8EEF8', type: ShadingType.CLEAR },
              borders: { top: border, bottom: border, left: border, right: border },
              children: [new Paragraph({ children: [run(key, true, 20, '1F3864')] })],
            }),
            new TableCell({
              width: { size: 68, type: WidthType.PERCENTAGE },
              borders: { top: border, bottom: border, left: border, right: border },
              children: [new Paragraph({ children: [run(value || '—', false, 20)] })],
            }),
          ],
        }),
    ),
  })
}

async function buildDescritivoDocx(ai: any, form: DescritivoForm) {
  const children: any[] = [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 120 }, children: [run('SENAI BAHIA', true, 38, '1F3864')] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 260 }, children: [run('DESCRITIVO DE CURSO', true, 30, '1F3864')] }),
    title('1. IDENTIFICAÇÃO'),
    infoTable([
      ['Título do curso', form.nomeCurso],
      ['Ocupação/CBO', form.cbo],
      ['Modalidade', form.modalidade],
      ['Eixo tecnológico', form.eixo],
      ['Cliente', form.cliente],
      ['Carga horária', form.cargaHoraria],
    ]),
    title('2. JUSTIFICATIVA'),
    para(ai.justificativa),
    title('3. OBJETIVO'),
    para(ai.objetivo),
    title('4. DESCRIÇÃO DO CURSO'),
    para(ai.descricao),
    title('5. PÚBLICO-ALVO'),
    para(ai.publicoAlvo),
    title('6. REQUISITOS DE ACESSO'),
    para(`Escolaridade mínima: ${form.escolaridade}. Idade mínima: ${form.eja ? '18' : form.idadeMinima} anos. ${form.requisitos || ''}`),
    title('7. PERFIL PROFISSIONAL DE CONCLUSÃO'),
    ...(ai.perfilSaida || []).map((item: string) => bullet(item)),
    title('8. ORGANIZAÇÃO CURRICULAR'),
  ]

  ;(ai.modulos || []).forEach((modulo: any) => {
    children.push(new Paragraph({ spacing: { before: 160, after: 80 }, children: [run(`${modulo.nome}${modulo.chModulo ? ` — ${modulo.chModulo}` : ''}`, true, 24, '1F3864')] }))
    ;(modulo.ucs || []).forEach((uc: any) => {
      children.push(new Paragraph({ spacing: { before: 100, after: 60 }, children: [run(`${uc.nome} — ${uc.ch || ''}`, true, 22)] }))
      children.push(para(uc.objetivo))
      children.push(new Paragraph({ children: [run('Capacidades:', true, 20, '1F3864')] }))
      ;(uc.capacidades || []).forEach((item: string) => children.push(bullet(item)))
      children.push(new Paragraph({ children: [run('Conhecimentos:', true, 20, '1F3864')] }))
      ;(uc.conhecimentos || []).forEach((item: string) => children.push(bullet(item)))
    })
  })

  children.push(
    title('9. METODOLOGIA'),
    para(ai.metodologia),
    title('10. CRITÉRIOS DE AVALIAÇÃO'),
    para(ai.criteriosAvaliacao),
    title('11. CRITÉRIOS DE CERTIFICAÇÃO'),
    para(ai.criteriosCertificacao),
    title('12. PERFIL DOCENTE'),
    para(ai.perfilDocente),
    title('13. INFRAESTRUTURA MÍNIMA'),
  )

  const infra = ai.infraestrutura || {}
  ;[...(infra.ambiente || []), ...(infra.softwares || []), ...(infra.epis || []), ...(infra.conectividade || [])].forEach((item: string) => children.push(bullet(item)))
  if (infra.materiais) children.push(para(infra.materiais))

  if (ai.eja?.aplicar) {
    children.push(title('14. DIRETRIZES ESPECÍFICAS DO EJA PROFISSIONALIZANTE'), para(ai.eja.diretrizes))
    ;(ai.eja.atividadesAssincronas || []).forEach((item: string) => children.push(bullet(item)))
    ;(ai.eja.permanenciaExito || []).forEach((item: string) => children.push(bullet(item)))
  }

  children.push(title('REFERÊNCIAS'))
  ;(ai.referencias || []).forEach((item: string) => children.push(bullet(item)))

  const doc = new Document({ sections: [{ properties: { page: { margin: { top: 1134, right: 1440, bottom: 1134, left: 1440 } } }, children }] })
  return await Packer.toBlob(doc)
}

async function buildFichaDocx(ai: any, form: FichaForm) {
  const children: any[] = [
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 80 }, children: [run('SENAI BAHIA', true, 34, '1F3864')] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 220 }, children: [run(ai.nome || form.nome, true, 30, '1F3864')] }),
    title('DESCRIÇÃO DO CURSO'),
    para(ai.descricao),
    infoTable([
      ['Carga horária', ai.cargaHoraria || form.ch],
      ['Participantes', ai.participantes || form.publico],
      ['Certificado', ai.certificado || form.certificado],
      ['Frequência mínima', ai.frequencia || form.frequencia],
    ]),
    title('PROGRAMAÇÃO'),
  ]

  ;(ai.programacao || []).forEach((modulo: any) => {
    children.push(new Paragraph({ spacing: { before: 140, after: 60 }, children: [run(modulo.modulo, true, 22, '1F3864')] }))
    ;(modulo.topicos || []).forEach((item: string) => children.push(bullet(item)))
  })

  const doc = new Document({ sections: [{ properties: { page: { margin: { top: 1134, right: 1440, bottom: 1134, left: 1440 } } }, children }] })
  return await Packer.toBlob(doc)
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = filename
  a.click()
  URL.revokeObjectURL(url)
}

export default function AppPage() {
  const router = useRouter()
  const [user, setUser] = useState<User | null>(null)
  const [checking, setChecking] = useState(true)
  const [mode, setMode] = useState<Mode>('descritivo')
  const [desc, setDesc] = useState<DescritivoForm>(defaultDesc)
  const [ficha, setFicha] = useState<FichaForm>(defaultFicha)
  const [docText, setDocText] = useState('')
  const [docName, setDocName] = useState('')
  const [fichaDocText, setFichaDocText] = useState('')
  const [fichaDocName, setFichaDocName] = useState('')
  const [loading, setLoading] = useState(false)
  const [logs, setLogs] = useState<LogItem[]>([])
  const [resultLabel, setResultLabel] = useState('')

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (currentUser) => {
      setUser(currentUser)
      setChecking(false)
      if (!currentUser) router.replace('/login')
    })
    return () => unsub()
  }, [router])

  const ready = useMemo(() => {
    if (mode === 'descritivo') return Boolean(desc.nomeCurso.trim() && desc.cargaHoraria.trim())
    return Boolean(ficha.nome.trim() && ficha.ch.trim())
  }, [mode, desc.nomeCurso, desc.cargaHoraria, ficha.nome, ficha.ch])

  function log(text: string, type: LogKind = 'info') {
    setLogs((items) => [...items, { text, type }])
  }

  async function handleFile(file: File | undefined, target: 'descritivo' | 'ficha') {
    if (!file) return
    const text = await extractText(file)
    if (target === 'descritivo') {
      setDocName(file.name)
      setDocText(text)
    } else {
      setFichaDocName(file.name)
      setFichaDocText(text)
    }
  }

  async function generate() {
    if (!ready) return
    setLoading(true)
    setLogs([])
    setResultLabel('')

    try {
      log('Preparando dados do formulário...', 'info')
      log(mode === 'descritivo' ? `Curso: ${desc.nomeCurso}` : `Ficha: ${ficha.nome}`, 'info')
      log((mode === 'descritivo' ? docText : fichaDocText) ? 'Documento de referência carregado.' : 'Sem documento de referência.', 'info')

      const payload = mode === 'descritivo'
        ? { mode, form: desc, documentText: docText }
        : { mode, form: ficha, documentText: fichaDocText }

      log('Chamando API segura da Vercel...', 'info')
      const response = await fetch('/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      })
      const data = await response.json()
      if (!response.ok) throw new Error(data.error || 'Erro ao gerar conteúdo.')

      log('Resposta recebida. Montando DOCX...', 'ok')
      const blob = mode === 'descritivo'
        ? await buildDescritivoDocx(data.result, desc)
        : await buildFichaDocx(data.result, ficha)

      const baseName = (mode === 'descritivo' ? desc.nomeCurso : ficha.nome).replace(/[^a-zA-Z0-9\s-]/g, '').trim().replace(/\s+/g, '-')
      downloadBlob(blob, `${mode === 'descritivo' ? 'Descritivo' : 'Ficha-Produto'}-${baseName || 'SENAI'}.docx`)
      log('Documento gerado e download iniciado.', 'ok')
      setResultLabel('Documento pronto!')
    } catch (error: any) {
      log(error?.message || 'Erro inesperado.', 'err')
      setResultLabel('Não foi possível concluir.')
    } finally {
      setLoading(false)
    }
  }

  if (checking) {
    return <main className="app-loading">Carregando...</main>
  }

  return (
    <main className="app-page">
      <header className="topbar">
        <div className="brand-card slim">
          <img src="/senai-logo.svg" alt="SENAI" />
          <div>
            <strong>SENAI BAHIA</strong>
            <span>Gerador de Descritivos V3</span>
          </div>
        </div>

        <div className="top-actions">
          <span className="user-pill">{user?.email}</span>
          <button className="ghost-btn" onClick={() => signOut(auth)}>Sair</button>
        </div>
      </header>

      <section className="app-hero glass-panel">
        <div>
          <div className="eyebrow">Ambiente protegido</div>
          <h1>Geração documental com Claude</h1>
          <p>
            Produza descritivos de curso e fichas de produto com padrão SENAI,
            mantendo a chave da Anthropic protegida na Vercel.
          </p>
        </div>
        <div className="stats-strip">
          <span>📄 Descritivo</span>
          <span>🗂️ Ficha</span>
          <span>📎 Upload</span>
          <span>🎓 EJA</span>
        </div>
      </section>

      <div className="workspace">
        <aside className="sidebar glass-panel">
          <button className={`navitem ${mode === 'descritivo' ? 'active' : ''}`} onClick={() => setMode('descritivo')}>📄 Descritivo de Curso</button>
          <button className={`navitem ${mode === 'ficha' ? 'active' : ''}`} onClick={() => setMode('ficha')}>🗂️ Ficha de Produto</button>
          <div className="sidebar-note">
            <b>Base v2 migrada</b>
            <span>EJA, Aprender a Empreender, upload de documentos e geração DOCX.</span>
          </div>
        </aside>

        <section className="content-stack">
          {mode === 'descritivo' ? (
            <>
              <div className="glass-panel form-card">
                <div className="section-head"><span>1</span><div><h2>Dados do curso</h2><p>Informe os dados principais do descritivo.</p></div></div>
                <div className="field"><label>Nome do curso *</label><input value={desc.nomeCurso} onChange={(e) => setDesc({ ...desc, nomeCurso: e.target.value })} placeholder="Ex: Assistente em Tecnologias da Indústria 4.0" /></div>
                <div className="grid2">
                  <div className="field"><label>Carga horária *</label><input value={desc.cargaHoraria} onChange={(e) => setDesc({ ...desc, cargaHoraria: e.target.value })} placeholder="Ex: 200h" /></div>
                  <div className="field"><label>CBO/Ocupação</label><input value={desc.cbo} onChange={(e) => setDesc({ ...desc, cbo: e.target.value })} placeholder="Opcional" /></div>
                </div>
                <div className="grid2">
                  <div className="field"><label>Modalidade</label><select value={desc.modalidade} onChange={(e) => setDesc({ ...desc, modalidade: e.target.value })}>{modalidades.map((item) => <option key={item}>{item}</option>)}</select></div>
                  <div className="field"><label>Eixo tecnológico</label><select value={desc.eixo} onChange={(e) => setDesc({ ...desc, eixo: e.target.value })}>{eixos.map((item) => <option key={item}>{item}</option>)}</select></div>
                </div>
                <div className="grid3">
                  <div className="field"><label>Escolaridade</label><input value={desc.escolaridade} onChange={(e) => setDesc({ ...desc, escolaridade: e.target.value })} /></div>
                  <div className="field"><label>Idade mínima</label><input value={desc.idadeMinima} onChange={(e) => setDesc({ ...desc, idadeMinima: e.target.value })} /></div>
                  <div className="field"><label>Cliente</label><input value={desc.cliente} onChange={(e) => setDesc({ ...desc, cliente: e.target.value })} placeholder="Opcional" /></div>
                </div>
                <div className="field"><label>Outros requisitos</label><input value={desc.requisitos} onChange={(e) => setDesc({ ...desc, requisitos: e.target.value })} placeholder="Opcional" /></div>
              </div>

              <div className="glass-panel form-card">
                <div className="section-head"><span>2</span><div><h2>Configurações pedagógicas</h2><p>Ative somente o que fará parte da proposta.</p></div></div>
                <SwitchRow title="Módulo Básico" desc="Sustentabilidade, SST, qualidade e Indústria 4.0 (+40h)." checked={desc.moduloBasico} onChange={(checked) => setDesc({ ...desc, moduloBasico: checked })} />
                <div className="switch-row">
                  <div><b>Aprender a Empreender</b><p>Conteúdo atualizado com empreendedorismo digital e mobile.</p></div>
                  <div className="switch-controls">
                    <select value={desc.aprenderEmpreenderCh} onChange={(e) => setDesc({ ...desc, aprenderEmpreenderCh: e.target.value })}><option>16</option><option>24</option><option>32</option><option>40</option></select>
                    <label className="switch"><input type="checkbox" checked={desc.aprenderEmpreender} onChange={(e) => setDesc({ ...desc, aprenderEmpreender: e.target.checked })} /><span /></label>
                  </div>
                </div>
                <SwitchRow title="EJA Profissionalizante" desc="Botão simples: aplica 18 anos, MRS, 80% presencial e 20% assíncrono no documento." checked={desc.eja} onChange={(checked) => setDesc({ ...desc, eja: checked })} />
              </div>

              <div className="glass-panel form-card">
                <div className="section-head"><span>3</span><div><h2>Unidades curriculares</h2><p>Opcional: defina quantidade ou nomes específicos.</p></div></div>
                <div className="grid2">
                  <div className="field"><label>Quantidade de UCs técnicas</label><input value={desc.numUC} onChange={(e) => setDesc({ ...desc, numUC: e.target.value })} placeholder="Ex: 3" /></div>
                  <SwitchRow title="Informar nomes das UCs" desc="Use quando já houver matriz definida." checked={desc.usarNomesUC} onChange={(checked) => setDesc({ ...desc, usarNomesUC: checked })} compact />
                </div>
                {desc.usarNomesUC && <UcEditor items={desc.ucs} onChange={(ucs) => setDesc({ ...desc, ucs })} />}
              </div>

              <UploadCard fileName={docName} text={docText} onFile={(file) => handleFile(file, 'descritivo')} onClear={() => { setDocName(''); setDocText('') }} />
            </>
          ) : (
            <>
              <div className="glass-panel form-card">
                <div className="section-head"><span>1</span><div><h2>Ficha de produto</h2><p>Preencha os dados mínimos para gerar a ficha.</p></div></div>
                <div className="field"><label>Nome do curso/produto *</label><input value={ficha.nome} onChange={(e) => setFicha({ ...ficha, nome: e.target.value })} placeholder="Ex: Criação de Prompts e Agentes de IA Especialistas" /></div>
                <div className="grid2">
                  <div className="field"><label>Carga horária *</label><input value={ficha.ch} onChange={(e) => setFicha({ ...ficha, ch: e.target.value })} placeholder="Ex: 80 horas" /></div>
                  <div className="field"><label>Modalidade</label><select value={ficha.modalidade} onChange={(e) => setFicha({ ...ficha, modalidade: e.target.value })}>{modalidades.map((item) => <option key={item}>{item}</option>)}</select></div>
                </div>
                <div className="grid2">
                  <div className="field"><label>Eixo tecnológico</label><select value={ficha.eixo} onChange={(e) => setFicha({ ...ficha, eixo: e.target.value })}>{eixos.map((item) => <option key={item}>{item}</option>)}</select></div>
                  <div className="field"><label>Frequência mínima</label><input value={ficha.frequencia} onChange={(e) => setFicha({ ...ficha, frequencia: e.target.value })} /></div>
                </div>
                <div className="field"><label>Público-alvo</label><input value={ficha.publico} onChange={(e) => setFicha({ ...ficha, publico: e.target.value })} placeholder="Profissionais e estudantes..." /></div>
                <div className="field"><label>Certificado emitido</label><input value={ficha.certificado} onChange={(e) => setFicha({ ...ficha, certificado: e.target.value })} placeholder="Opcional" /></div>
              </div>
              <UploadCard fileName={fichaDocName} text={fichaDocText} onFile={(file) => handleFile(file, 'ficha')} onClear={() => { setFichaDocName(''); setFichaDocText('') }} />
            </>
          )}

          <div className="glass-panel generate-panel">
            <div>
              <h2>Gerar documento</h2>
              <p>{mode === 'descritivo' ? 'Será criado um descritivo completo em .docx.' : 'Será criada uma ficha de produto em .docx.'}</p>
            </div>
            <button className="primary-btn" disabled={!ready || loading} onClick={generate}>{loading ? 'Gerando...' : '⚡ Gerar e baixar'}</button>
          </div>

          {(logs.length > 0 || resultLabel) && (
            <div className="glass-panel log-panel">
              <div className="section-head small"><span>✓</span><div><h2>{resultLabel || 'Processamento'}</h2><p>Acompanhe as etapas da geração.</p></div></div>
              <div className="log-box">{logs.map((item, index) => <div key={index} className={item.type}>› {item.text}</div>)}</div>
            </div>
          )}
        </section>
      </div>
    </main>
  )
}

function SwitchRow({ title, desc, checked, onChange, compact = false }: { title: string; desc: string; checked: boolean; onChange: (value: boolean) => void; compact?: boolean }) {
  return (
    <div className={`switch-row ${compact ? 'compact-switch' : ''}`}>
      <div><b>{title}</b><p>{desc}</p></div>
      <label className="switch"><input type="checkbox" checked={checked} onChange={(e) => onChange(e.target.checked)} /><span /></label>
    </div>
  )
}

function UcEditor({ items, onChange }: { items: { nome: string; ch: string }[]; onChange: (items: { nome: string; ch: string }[]) => void }) {
  function update(index: number, patch: Partial<{ nome: string; ch: string }>) {
    onChange(items.map((item, current) => (current === index ? { ...item, ...patch } : item)))
  }
  return (
    <div className="uc-list">
      {items.map((item, index) => (
        <div className="uc-row" key={index}>
          <input value={item.nome} onChange={(e) => update(index, { nome: e.target.value })} placeholder="Nome da UC" />
          <input value={item.ch} onChange={(e) => update(index, { ch: e.target.value })} placeholder="40h" />
          <button onClick={() => onChange(items.filter((_, current) => current !== index))}>×</button>
        </div>
      ))}
      <button className="ghost-btn" onClick={() => onChange([...items, { nome: '', ch: '' }])}>+ Adicionar UC</button>
    </div>
  )
}

function UploadCard({ fileName, text, onFile, onClear }: { fileName: string; text: string; onFile: (file: File | undefined) => void; onClear: () => void }) {
  return (
    <div className="glass-panel form-card">
      <div className="section-head"><span>4</span><div><h2>Documento de referência</h2><p>Opcional: use itinerários, fichas anteriores, catálogos ou arquivos técnicos.</p></div></div>
      <label className="upload-box">
        <input type="file" accept=".docx,.pdf,.txt" onChange={(event) => onFile(event.target.files?.[0])} />
        <b>📎 Adicionar documento</b>
        <span>Arraste ou clique para enviar DOCX, PDF ou TXT</span>
      </label>
      {fileName && (
        <div className="doc-preview">
          <div><b>✅ {fileName}</b><button onClick={onClear}>Remover</button></div>
          <p>{text.slice(0, 600)}{text.length > 600 ? '...' : ''}</p>
        </div>
      )}
    </div>
  )
}
