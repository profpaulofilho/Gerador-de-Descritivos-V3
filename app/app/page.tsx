// v3.1.0 - build fix
'use client'

import { useEffect, useMemo, useState } from 'react'
import { useRouter } from 'next/navigation'
import { onAuthStateChanged, signOut, User } from 'firebase/auth'
import { collection, addDoc, getDocs, query, orderBy, serverTimestamp } from 'firebase/firestore'
import { auth, db } from '../../lib/firebase'
import {
  AlignmentType, BorderStyle, Document, LevelFormat,
  Packer, Paragraph, ShadingType, Table, TableCell,
  TableRow, TextRun, WidthType,
} from 'docx'
import * as mammoth from 'mammoth'

/* ── TYPES ──────────────────────────────────────────────── */
type Mode = 'descritivo' | 'ficha'
type LogKind = 'ok' | 'err' | 'info'
type LogItem = { text: string; type?: LogKind }
type HistoryItem = { id: string; tipo: string; nome: string; criadoEm: any }

type DescritivoForm = {
  nomeCurso: string; cargaHoraria: string; cbo: string
  modalidade: string; eixo: string; escolaridade: string
  idadeMinima: string; cliente: string; requisitos: string
  moduloBasico: boolean; aprenderEmpreender: boolean; aprenderEmpreenderCh: string
  eja: boolean; numUC: string; usarNomesUC: boolean
  ucs: { nome: string; ch: string }[]
  chTotal?: string // calculado automaticamente
}
type FichaForm = {
  nome: string; ch: string; modalidade: string; eixo: string
  publico: string; certificado: string; frequencia: string
}

const defaultDesc: DescritivoForm = {
  nomeCurso: '', cargaHoraria: '', cbo: '',
  modalidade: 'Qualificação Profissional', eixo: 'Informação e Comunicação',
  escolaridade: 'Ensino Médio incompleto', idadeMinima: '16',
  cliente: '', requisitos: '', moduloBasico: false,
  aprenderEmpreender: true, aprenderEmpreenderCh: '24',
  eja: false, numUC: '', usarNomesUC: false, ucs: [],
}
const defaultFicha: FichaForm = {
  nome: '', ch: '', modalidade: 'Aperfeiçoamento Profissional',
  eixo: 'Informação e Comunicação', publico: '', certificado: '', frequencia: '75%',
}

const modalidades = ['Qualificação Profissional','Aperfeiçoamento Profissional','Aprendizagem Industrial','Habilitação Profissional']
const eixos = ['Ambiente e Saúde','Controle e Processos Industriais','Gestão e Negócios','Informação e Comunicação','Infraestrutura','Produção Alimentícia','Produção Cultural e Design','Produção Industrial','Recursos Naturais','Segurança','Turismo, Hospitalidade e Lazer']

/* ── FILE EXTRACTION ─────────────────────────────────────── */
async function extractText(file: File) {
  const name = file.name.toLowerCase()
  if (name.endsWith('.txt')) return await file.text()
  if (name.endsWith('.docx')) {
    const ab = await file.arrayBuffer()
    const r = await mammoth.extractRawText({ arrayBuffer: ab })
    return r.value || ''
  }
  if (name.endsWith('.pdf')) {
    const pdfjsLib = (await import('pdfjs-dist')) as any
    pdfjsLib.GlobalWorkerOptions.workerSrc = `https://cdnjs.cloudflare.com/ajax/libs/pdf.js/${pdfjsLib.version}/pdf.worker.min.mjs`
    const ab = await file.arrayBuffer()
    const pdf = await pdfjsLib.getDocument({ data: ab }).promise
    let txt = ''
    for (let i = 1; i <= Math.min(pdf.numPages, 20); i++) {
      const pg = await pdf.getPage(i)
      const ct = await pg.getTextContent()
      txt += ct.items.map((it: any) => it.str).join(' ') + '\n'
    }
    return txt
  }
  return ''
}

/* ── DOCX HELPERS ────────────────────────────────────────── */
const BLUE = '154194'; const WHITE = 'FFFFFF'; const GRAY = 'D9D9D9'; const BLACK = '000000'; const BLACK = '000000'
const TW = 9026 // content width A4 with 1.27cm margins (DXA)

function run(text: string, bold = false, size = 20, color = '000000') {
  return new TextRun({ text, bold, size, color, font: 'Arial' })
}
function para(text: string, bold = false, justify = true) {
  return new Paragraph({
    spacing: { before: 60, after: 60 },
    alignment: justify ? AlignmentType.JUSTIFIED : AlignmentType.LEFT,
    children: [run(text || '—', bold)],
  })
}
function sectionTitle(num: string, text: string) {
  return new Paragraph({
    spacing: { before: 240, after: 120 },
    children: [run(`${num}. ${text.toUpperCase()}`, true, 24, BLUE)],
  })
}
function subTitle(text: string) {
  return new Paragraph({
    spacing: { before: 120, after: 60 },
    children: [run(text, true, 22, BLUE)],
  })
}
function bullet(text: string) {
  return new Paragraph({
    numbering: { reference: 'bullets', level: 0 },
    spacing: { before: 30, after: 30 },
    children: [run(text, false, 20)],
  })
}
function bS(c = 'AAAAAA') { return { style: BorderStyle.SINGLE, size: 4, color: c } }
function bAll(c = 'AAAAAA') { return { top: bS(c), bottom: bS(c), left: bS(c), right: bS(c) } }

/* tabela de identificação 2 colunas */
function idTable(rows: [string, string][]) {
  return new Table({
    width: { size: TW, type: WidthType.DXA },
    columnWidths: [3200, TW - 3200],
    rows: rows.map(([k, v]) => new TableRow({ children: [
      new TableCell({
        borders: bAll(BLUE), width: { size: 3200, type: WidthType.DXA },
        shading: { fill: 'FFFFFF', type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [run(k, true, 20, BLACK)] })],
      }),
      new TableCell({
        borders: bAll(BLUE), width: { size: TW - 3200, type: WidthType.DXA },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ children: [run(v || '', false, 20, BLACK)] })],
      }),
    ]})),
  })
}

/* tabela UC com capacidades e conhecimentos */
function ucTable(moduloNome: string, ucNome: string, ch: string, objetivo: string, capacidades: string[], conhecimentos: string[]) {
  const half = Math.floor(TW / 2)
  return [
    // header módulo
    new Table({
      width: { size: TW, type: WidthType.DXA }, columnWidths: [TW],
      rows: [new TableRow({ children: [new TableCell({
        borders: bAll('AAAAAA'), width: { size: TW, type: WidthType.DXA },
        shading: { fill: 'D9D9D9', type: ShadingType.CLEAR },
        margins: { top: 80, bottom: 80, left: 120, right: 120 },
        children: [new Paragraph({ alignment: AlignmentType.CENTER, children: [run(moduloNome, true, 22, BLACK)] })],
      })]})],
    }),
    // UC name / ch / objetivo
    new Table({
      width: { size: TW, type: WidthType.DXA }, columnWidths: [TW],
      rows: [
        new TableRow({ children: [new TableCell({ borders: bAll(), width: { size: TW, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [run('Unidade Curricular: ', true, 20), run(ucNome, false, 20)] })] })]}),
        new TableRow({ children: [new TableCell({ borders: bAll(), width: { size: TW, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [run('Carga Horária: ', true, 20), run(ch, false, 20)] })] })]}),
        new TableRow({ children: [new TableCell({ borders: bAll(), width: { size: TW, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [run('Objetivo Geral: ', true, 20), run(objetivo, false, 20)] })] })]}),
      ],
    }),
    // header capacidades / conhecimentos
    new Table({
      width: { size: TW, type: WidthType.DXA }, columnWidths: [half, TW - half],
      rows: [
        new TableRow({ children: [
          new TableCell({ borders: bAll(GRAY), width: { size: half, type: WidthType.DXA }, shading: { fill: GRAY, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [run('CAPACIDADES BÁSICAS', true, 20)] })] }),
          new TableCell({ borders: bAll(GRAY), width: { size: TW - half, type: WidthType.DXA }, shading: { fill: GRAY, type: ShadingType.CLEAR }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: [new Paragraph({ children: [run('CONHECIMENTOS', true, 20)] })] }),
        ]}),
        new TableRow({ children: [
          new TableCell({ borders: bAll(), width: { size: half, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: capacidades.map(c => new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { before: 20, after: 20 }, children: [run(c, false, 18)] })) }),
          new TableCell({ borders: bAll(), width: { size: TW - half, type: WidthType.DXA }, margins: { top: 80, bottom: 80, left: 120, right: 120 }, children: conhecimentos.map(c => new Paragraph({ numbering: { reference: 'bullets', level: 0 }, spacing: { before: 20, after: 20 }, children: [run(c, false, 18)] })) }),
        ]}),
      ],
    }),
    new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }),
  ]
}

/* tabela equipamentos */
function equipTable(rows: [string, string][]) {
  return new Table({
    width: { size: TW, type: WidthType.DXA }, columnWidths: [6200, TW - 6200],
    rows: [
      new TableRow({ children: [
        new TableCell({ borders: bAll(GRAY), shading: { fill: GRAY, type: ShadingType.CLEAR }, width: { size: 6200, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 100, right: 100 }, children: [new Paragraph({ children: [run('Equipamento', true, 18)] })] }),
        new TableCell({ borders: bAll(GRAY), shading: { fill: GRAY, type: ShadingType.CLEAR }, width: { size: TW - 6200, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 100, right: 100 }, children: [new Paragraph({ children: [run('Qtd. Mínima', true, 18)] })] }),
      ]}),
      ...rows.map(([eq, qt]) => new TableRow({ children: [
        new TableCell({ borders: bAll(), width: { size: 6200, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 100, right: 100 }, children: [new Paragraph({ children: [run(eq, false, 18)] })] }),
        new TableCell({ borders: bAll(), width: { size: TW - 6200, type: WidthType.DXA }, margins: { top: 60, bottom: 60, left: 100, right: 100 }, children: [new Paragraph({ children: [run(qt, false, 18)] })] }),
      ]})),
    ],
  })
}

function docxDoc(children: any[]) {
  return new Document({
    numbering: { config: [{ reference: 'bullets', levels: [{ level: 0, format: LevelFormat.BULLET, text: '\u2022', alignment: AlignmentType.LEFT, style: { paragraph: { indent: { left: 540, hanging: 260 } } } }] }] },
    styles: { default: { document: { run: { font: 'Arial', size: 20, color: '000000' } } } },
    sections: [{ properties: { page: { size: { width: 11906, height: 16838 }, margin: { top: 1134, right: 1800, bottom: 1134, left: 1800 } } }, children }],
  })
}

/* ── BUILD DESCRITIVO DOCX ───────────────────────────────── */
async function buildDescritivoDocx(ai: any, form: DescritivoForm) {
  const children: any[] = [
    // Cabeçalho
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [run('SENAI BAHIA', true, 36, BLUE)] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 40 }, children: [run(form.eixo.toUpperCase(), true, 22, BLUE)] }),
    new Paragraph({ alignment: AlignmentType.CENTER, spacing: { after: 280 }, children: [run('DESCRITIVO DE CURSO', true, 30, BLUE)] }),

    sectionTitle('1', 'IDENTIFICAÇÃO'),
    new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }),
    idTable([
      ['TÍTULO DO CURSO:', form.nomeCurso],
      ['OCUPAÇÃO/CBO:', form.cbo || ''],
      ['MODALIDADE:', form.modalidade],
      ['EIXO TECNOLÓGICO:', form.eixo],
      ['CLIENTE:', form.cliente || ''],
      ['CARGA HORÁRIA:', form.chTotal || form.cargaHoraria],
    ]),

    sectionTitle('2', 'JUSTIFICATIVA'),
    para(ai.justificativa),
    subTitle('2.1 OBJETIVO:'),
    para(ai.objetivo),
    subTitle('2.2 DESCRIÇÃO DO CURSO'),
    para(ai.descricao),
    subTitle('2.3 PÚBLICO-ALVO'),
    para(ai.publicoAlvo),

    sectionTitle('3', 'REQUISITOS DE ACESSO OBRIGATÓRIO'),
    para(`Escolaridade mínima: ${form.escolaridade}.`),
    para(`Idade mínima: ${form.eja ? '18' : form.idadeMinima} anos.`),
    ...(form.eja ? [para('Não ter concluído o Ensino Médio.'), para('Estar apto à matrícula na EJA Profissionalizante SESI/SENAI.')] : []),
    ...(form.requisitos ? [para(form.requisitos)] : []),

    sectionTitle('4', 'PERFIL PROFISSIONAL DE CONCLUSÃO'),
    para('O egresso será capaz de:'),
    ...(ai.perfilSaida || []).map((item: string) => bullet(item)),

    sectionTitle('5', 'ORGANIZAÇÃO CURRICULAR'),
    subTitle('5.1 UNIDADES CURRICULARES, CARGAS HORÁRIAS E CONTEÚDOS FORMATIVOS:'),
    new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }),
  ]

  // Módulo Básico — ANTES do Aprender a Empreender
  if (form.moduloBasico) {
    const basicUcs = [
      { nome: 'Sustentabilidade nos Processos Industriais', ch: '8h', objetivo: 'Desenvolver a capacidade do estudante de aplicar práticas de sustentabilidade em processos industriais, identificando impactos ambientais e propondo soluções sustentáveis.', cap: ['Identificar impactos ambientais de processos produtivos', 'Aplicar práticas de economia circular no contexto industrial', 'Propor soluções de redução de resíduos e consumo', 'Implementar princípios de gestão ambiental'], con: ['Conceitos de sustentabilidade e desenvolvimento sustentável', 'Legislação ambiental brasileira', 'ISO 14001 — Gestão ambiental', 'Economia circular e logística reversa', 'Poluição industrial e tratamento de resíduos', 'Boas práticas ambientais na indústria'] },
      { nome: 'Saúde e Segurança no Trabalho', ch: '12h', objetivo: 'Capacitar o estudante a identificar e aplicar normas de saúde e segurança do trabalho, prevenindo riscos e promovendo ambientes laborais seguros.', cap: ['Identificar riscos e perigos no ambiente de trabalho', 'Aplicar normas regulamentadoras pertinentes à área', 'Utilizar corretamente EPIs e EPCs', 'Adotar posturas ergonômicas e práticas seguras'], con: ['Normas Regulamentadoras NR-01 a NR-06', 'CIPA — Comissão Interna de Prevenção de Acidentes', 'EPIs e EPCs: tipos e uso correto', 'Ergonomia e doenças ocupacionais', 'SIPAT e programas de segurança', 'Legislação trabalhista básica — CLT'] },
      { nome: 'Fundamentos da Qualidade e Produtividade', ch: '8h', objetivo: 'Desenvolver compreensão sobre qualidade, produtividade e melhoria contínua nos processos industriais.', cap: ['Identificar conceitos e ferramentas de qualidade', 'Aplicar técnicas de melhoria contínua', 'Analisar indicadores de produtividade', 'Reconhecer padrões de qualidade aplicados à área'], con: ['Conceitos de qualidade total (TQM)', 'Ferramentas da qualidade: 5S, PDCA, Ishikawa', 'Indicadores de produtividade', 'ISO 9001 — Sistemas de gestão da qualidade', 'Padronização de processos', 'Kaizen e melhoria contínua'] },
      { nome: 'Fundamentos da Indústria 4.0', ch: '12h', objetivo: 'Apresentar os princípios e tecnologias da Indústria 4.0 e sua aplicação nos processos produtivos modernos.', cap: ['Identificar as principais tecnologias da Indústria 4.0', 'Relacionar transformação digital com o setor produtivo', 'Reconhecer aplicações de IoT e automação', 'Compreender o impacto das tecnologias digitais no trabalho'], con: ['Conceito e pilares da Indústria 4.0', 'Internet das Coisas (IoT) aplicada à indústria', 'Big Data e Analytics', 'Inteligência Artificial e Machine Learning', 'Computação em nuvem', 'Cibersegurança industrial', 'Impressão 3D e manufatura aditiva'] },
    ]
    basicUcs.forEach(uc => children.push(...ucTable('MÓDULO BÁSICO', uc.nome, uc.ch, uc.objetivo, uc.cap, uc.con)))
  }

  // Módulo Aprender a Empreender — SEMPRE PRIMEIRO, conteúdo completo do documento atualizado
  if (form.aprenderEmpreender) {
    const chEmp = form.aprenderEmpreenderCh + 'h'
    children.push(...ucTable(
      'MÓDULO APRENDER A EMPREENDER',
      'Aprender a Empreender', chEmp,
      'Desenvolver atitudes empreendedoras e aplicar conhecimentos básicos de marketing, finanças e inovação, com uso de tecnologias digitais acessíveis, para planejar e divulgar pequenos negócios ou soluções criativas voltadas à realidade socioprofissional.',
      [
        'Identificar características e atitudes empreendedoras',
        'Reconhecer o próprio potencial para empreender com base em experiências pessoais',
        'Utilizar ferramentas digitais (mobile) para organização de ideias e metas (ex.: mapas mentais, quadros de tarefas)',
        'Relacionar valores pessoais com o propósito de um negócio ou iniciativa',
        'Aplicar ferramentas de ideação (ex.: Canvas simplificado) na elaboração de ideias de negócio',
        'Realizar pesquisa de mercado utilizando canais digitais acessíveis',
        'Identificar oportunidades e necessidades no entorno social para empreender',
        'Utilizar plataformas digitais para captar e interpretar dados simples de mercado (ex.: Google Forms, enquetes no Instagram)',
        'Aplicar estratégias de marketing digital e atendimento utilizando aplicativos móveis',
        'Produzir conteúdo visual (fotos, vídeos, textos) de forma criativa para divulgar produtos ou serviços',
        'Utilizar ferramentas digitais para edição e divulgação de forma simples e eficaz',
        'Identificar boas práticas de comunicação com o cliente nos canais digitais (WhatsApp, Instagram, Facebook)',
        'Organizar o orçamento pessoal e do pequeno negócio com base em aplicativos de finanças',
        'Realizar registro básico de receitas e despesas',
        'Identificar custos, calcular preço de venda e analisar viabilidade do produto/serviço',
        'Aplicar princípios básicos de educação financeira na vida pessoal e profissional',
        'Compreender e aplicar direitos trabalhistas básicos, incluindo saúde e segurança do trabalho, igualdade de gênero, e combate ao trabalho escravo e infantil',
        'Relacionar práticas empreendedoras com responsabilidade social e legal, incorporando noções da CLT no cotidiano profissional',
      ],
      [
        'Atitude Empreendedora e Autoconhecimento (5h)',
        '• Características e atitudes empreendedoras',
        '• Autoempreendedorismo: missão pessoal e valores',
        '• Direitos fundamentais do trabalhador (CLT): saúde e segurança do trabalho, igualdade de gênero, combate ao trabalho escravo e infantil',
        '• Introdução ao uso de ferramentas digitais (Canva e Trello mobile) para registro de ideias e organização de projetos',
        'Empreendendo com o que tenho (10h)',
        '• Ideação e modelagem simples de negócios (Canvas simplificado)',
        '• Pesquisa de mercado com uso de redes sociais e aplicativos gratuitos (Google Forms, Instagram)',
        '• Identificação de oportunidades locais',
        '• Exemplos reais de pequenos empreendedores',
        'Divulgação, Atendimento e Direitos na Prática Profissional (5h)',
        '• Atendimento ao cliente com foco em canais digitais (WhatsApp Business, Instagram, delivery)',
        '• Produção de conteúdo e divulgação mobile-first (fotos, vídeos e descrições)',
        '• Aplicativos gratuitos para edição e divulgação (CapCut, Canva, InShot)',
        '• Relação entre prática de atendimento, respeito aos direitos dos consumidores e práticas trabalhistas éticas e legais',
        'Finanças Pessoais e do Negócio (4h)',
        '• Educação financeira pessoal e familiar',
        '• Controle financeiro com apps gratuitos (Mobilis, Minhas Economias)',
        '• Noções básicas de precificação e viabilidade do negócio',
      ]
    ))
  }

  // Módulos específicos da IA — filtrar qualquer módulo de Empreender gerado pela IA
  ;(ai.modulos || []).forEach((modulo: any) => {
    const nomeUpper = (modulo.nome || '').toUpperCase()
    // Ignorar módulos de Empreender e Básico gerados pela IA (já inseridos manualmente acima)
    if (
      nomeUpper.includes('EMPREEND') ||
      nomeUpper.includes('APRENDER') ||
      nomeUpper === 'MÓDULO BÁSICO' ||
      nomeUpper === 'MODULO BASICO'
    ) return
    ;(modulo.ucs || []).forEach((uc: any) => {
      children.push(...ucTable(
        modulo.nome + (modulo.chModulo ? ` — ${modulo.chModulo}` : ''),
        uc.nome, uc.ch || '', uc.objetivo,
        uc.capacidades || [], uc.conhecimentos || []
      ))
    })
  })

  // Seções seguintes
  children.push(
    sectionTitle('6', 'METODOLOGIA E ESTRATÉGIAS PEDAGÓGICAS'), para(ai.metodologia),
    sectionTitle('7', 'BIBLIOTECA, INSTALAÇÕES E EQUIPAMENTOS'),
    para('A Biblioteca deverá ser adequada, quando necessária, para o referido curso. Será disponibilizada a sistemática de biblioteca (itinerante) e do acervo bibliográfico. Os alunos receberão material didático para cada Unidade Curricular do curso.'),
    sectionTitle('8', 'CRITÉRIOS DE AVALIAÇÃO'),
    para(form.eja
      ? 'A avaliação será processual e contínua, com instrumentos diversificados: situações de aprendizagem, projetos integradores, portfólios, estudos de caso, avaliações práticas e autoavaliação por rubricas. Será considerado aprovado o aluno que alcançar nota igual ou superior a 5,0 (cinco) e frequência mínima de 75% (setenta e cinco por cento), conforme diretrizes do EJA Profissionalizante SESI/SENAI.'
      : 'A avaliação será processual e contínua, envolverá os aspectos qualitativos e quantitativos. Avaliação formativa por meio de exercícios práticos, desenvolvimento de projetos, participação nas discussões e apresentação de portfólio. Será considerado aprovado por média final, o aluno que alcançar nota igual ou superior a 7,0 (sete) e frequência mínima de 75% (setenta e cinco por cento) no curso, conforme diretrizes do Regimento Comum das Escolas Técnicas do SENAI DR-BA.'),
    sectionTitle('9', 'PERFIL DOCENTE'), para(ai.perfilDocente),
    sectionTitle('10', 'CRITÉRIOS DE CERTIFICAÇÃO'),
    para(`O certificado de ${form.modalidade} em ${form.nomeCurso} será conferido ao aluno que obtiver média de aproveitamento e índice de frequência estabelecido nos critérios de avaliação.`),
    subTitle('10.1 CONTEÚDO DO CERTIFICADO:'),
    new Paragraph({ spacing: { before: 60, after: 40 }, children: [run(`${form.modalidade} em ${form.nomeCurso} — ${form.chTotal || form.cargaHoraria}`, true, 22)] }),
  )

  if (form.moduloBasico) {
    children.push(new Paragraph({ spacing: { before: 60, after: 30 }, children: [run('MÓDULO BÁSICO — 40h', true, 20)] }))
    ;['Sustentabilidade nos Processos Industriais — 8h','Saúde e Segurança no Trabalho — 12h','Fundamentos da Qualidade e Produtividade — 8h','Fundamentos da Indústria 4.0 — 12h'].forEach(t => children.push(para(t)))
  }
  if (form.aprenderEmpreender) {
    children.push(new Paragraph({ spacing: { before: 60, after: 30 }, children: [run(`MÓDULO APRENDER A EMPREENDER — ${form.aprenderEmpreenderCh}h`, true, 20)] }))
    children.push(para(`Aprender a Empreender — ${form.aprenderEmpreenderCh}h`))
  }
  ;(ai.modulos || []).forEach((m: any) => {
    children.push(new Paragraph({ spacing: { before: 60, after: 30 }, children: [run(`${m.nome} — ${m.chModulo || ''}`, true, 20)] }))
    ;(m.ucs || []).forEach((u: any) => children.push(para(`${u.nome} — ${u.ch || ''}`)))
  })

  // Infraestrutura
  const infra = ai.infraestrutura || {}
  children.push(
    sectionTitle('11', 'INFRAESTRUTURA MÍNIMA RECOMENDADA'),
    subTitle('11.1 Ambiente Físico:'),
    ...(infra.ambiente || []).map((i: string) => bullet(i)),
    subTitle('11.2 Equipamentos Tecnológicos:'),
    new Paragraph({ spacing: { before: 80, after: 80 }, children: [] }),
    equipTable(infra.equipamentos || []),
    subTitle('11.3 Softwares:'),
    ...(infra.softwares || []).map((s: string) => bullet(s)),
    subTitle('11.4 Materiais Didáticos:'),
    para(infra.materiais || ''),
    subTitle('11.5 EPIs e Segurança:'),
    ...(infra.epis || []).map((e: string) => bullet(e)),
    subTitle('11.6 Conectividade:'),
    ...(infra.conectividade || []).map((c: string) => bullet(c)),
    sectionTitle('12', 'REFERÊNCIAS: https://estantedelivros.senai.br/'),
    ...(ai.referencias || []).map((r: string) => para(r)),
  )

  // EJA
  if (form.eja && ai.eja?.aplicar) {
    children.push(
      sectionTitle('13', 'DIRETRIZES ESPECÍFICAS DO EJA PROFISSIONALIZANTE'),
      para(ai.eja.diretrizes || ''),
      ...(ai.eja.atividadesAssincronas || []).map((i: string) => bullet(i)),
      ...(ai.eja.permanenciaExito || []).map((i: string) => bullet(i)),
    )
  }

  return await Packer.toBlob(docxDoc(children))
}

/* ── BUILD FICHA DOCX ────────────────────────────────────── */
async function buildFichaDocx(ai: any, form: FichaForm) {
  // Normalizar campos — a API pode retornar nomes ligeiramente diferentes
  const nomeCurso    = ai.nome       || form.nome        || ''
  const descricao    = ai.descricao  || ''
  const cargaHoraria = ai.cargaHoraria || form.ch        || ''
  const participantes= ai.participantes || ai.publicoAlvo || form.publico || ''
  const certificado  = ai.certificado  || form.certificado || `Certificado de ${form.modalidade} em ${form.nome}`
  const frequencia   = ai.frequencia   || form.frequencia  || '75%'
  const programacao  = Array.isArray(ai.programacao) ? ai.programacao : []

  const children: any[] = [
    // Cabeçalho azul
    new Table({
      width: { size: TW, type: WidthType.DXA }, columnWidths: [TW],
      rows: [new TableRow({ children: [new TableCell({
        borders: bAll(BLUE), width: { size: TW, type: WidthType.DXA },
        shading: { fill: BLUE, type: ShadingType.CLEAR },
        margins: { top: 160, bottom: 160, left: 200, right: 200 },
        children: [
          new Paragraph({ alignment: AlignmentType.CENTER, children: [run('SENAI BAHIA', true, 24, WHITE)] }),
          new Paragraph({ alignment: AlignmentType.CENTER, spacing: { before: 60 }, children: [run(nomeCurso, true, 32, WHITE)] }),
        ],
      })]})],
    }),
    new Paragraph({ spacing: { before: 160, after: 80 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE } }, children: [run('DESCRIÇÃO DO CURSO', true, 22, BLUE)] }),
    para(descricao),
    new Paragraph({ spacing: { before: 60, after: 80 }, children: [run('Carga horária: ', true, 20), run(cargaHoraria + ';', false, 20)] }),
    new Paragraph({ spacing: { before: 0, after: 60 }, children: [run('Participantes: ', true, 20), run(participantes, false, 20)] }),
    new Paragraph({ spacing: { before: 0, after: 120 }, children: [
      run(`Ao final do curso, os participantes que concluírem com frequência mínima de ${frequencia} e aproveitamento satisfatório receberão o `, false, 20),
      run(certificado, true, 20),
      run(', emitido pelo SENAI.', false, 20),
    ]}),
    new Paragraph({ spacing: { before: 160, after: 80 }, border: { bottom: { style: BorderStyle.SINGLE, size: 6, color: BLUE } }, children: [run('PROGRAMAÇÃO', true, 22, BLUE)] }),
  ]

  programacao.forEach((modulo: any) => {
    const nomeModulo = modulo.modulo || modulo.nome || ''
    const topicos = Array.isArray(modulo.topicos) ? modulo.topicos : []
    children.push(new Paragraph({ spacing: { before: 120, after: 60 }, children: [run('◆ ' + nomeModulo, true, 22, BLUE)] }))
    topicos.forEach((t: string) => children.push(bullet(t)))
  })

  return await Packer.toBlob(docxDoc(children))
}

function downloadBlob(blob: Blob, filename: string) {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url; a.download = filename; a.click()
  URL.revokeObjectURL(url)
}

/* ── MAIN COMPONENT ──────────────────────────────────────── */
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
  const [history, setHistory] = useState<HistoryItem[]>([])
  const [theme, setTheme] = useState<'dark' | 'light'>('dark')

  useEffect(() => {
    const unsub = onAuthStateChanged(auth, (currentUser) => {
      setUser(currentUser)
      setChecking(false)
      if (!currentUser) router.replace('/login')
      else loadHistory(currentUser.uid)
    })
    return () => unsub()
  }, [router])

  useEffect(() => {
    document.documentElement.setAttribute('data-theme', theme)
  }, [theme])

  async function loadHistory(uid: string) {
    try {
      const q = query(collection(db, 'documentos', uid, 'gerados'), orderBy('criadoEm', 'desc'))
      const snap = await getDocs(q)
      setHistory(snap.docs.map(d => ({ id: d.id, ...d.data() } as HistoryItem)))
    } catch { /* firestore pode não estar configurado ainda */ }
  }

  async function saveToFirestore(tipo: string, nome: string, resultado: any) {
    if (!user) return
    try {
      await addDoc(collection(db, 'documentos', user.uid, 'gerados'), {
        tipo, nome, resultado: JSON.stringify(resultado).slice(0, 50000),
        criadoEm: serverTimestamp(),
      })
      await loadHistory(user.uid)
    } catch { /* silencioso */ }
  }

  const ready = useMemo(() => {
    if (mode === 'descritivo') return Boolean(desc.nomeCurso.trim() && desc.cargaHoraria.trim())
    return Boolean(ficha.nome.trim() && ficha.ch.trim())
  }, [mode, desc.nomeCurso, desc.cargaHoraria, ficha.nome, ficha.ch])

  // Cálculo automático da carga horária total
  const chTotal = useMemo(() => {
    const base = parseInt(desc.cargaHoraria.replace(/[^0-9]/g, '')) || 0
    if (base === 0) return ''
    const basico = desc.moduloBasico ? 40 : 0
    const empreender = desc.aprenderEmpreender ? (parseInt(desc.aprenderEmpreenderCh) || 0) : 0
    return (base + basico + empreender) + 'h'
  }, [desc.cargaHoraria, desc.moduloBasico, desc.aprenderEmpreender, desc.aprenderEmpreenderCh])

  function log(text: string, type: LogKind = 'info') {
    setLogs(items => [...items, { text, type }])
  }

  async function handleFile(file: File | undefined, target: 'descritivo' | 'ficha') {
    if (!file) return
    const text = await extractText(file)
    if (target === 'descritivo') { setDocName(file.name); setDocText(text) }
    else { setFichaDocName(file.name); setFichaDocText(text) }
  }

  async function deleteHistoryItem(id: string) {
    if (!user) return
    try {
      const { deleteDoc, doc } = await import('firebase/firestore')
      await deleteDoc(doc(db, 'documentos', user.uid, 'gerados', id))
      setHistory(prev => prev.filter(h => h.id !== id))
    } catch { /* silencioso */ }
  }

  async function deleteAllHistory() {
    if (!user) return
    if (!confirm('Apagar todo o histórico? Esta ação não pode ser desfeita.')) return
    try {
      const { deleteDoc, doc } = await import('firebase/firestore')
      await Promise.all(history.map(h => deleteDoc(doc(db, 'documentos', user.uid, 'gerados', h.id))))
      setHistory([])
    } catch { /* silencioso */ }
  }

  async function generate() {
    setLoading(true); setLogs([]); setResultLabel('')
    try {
      log('Preparando dados...', 'ok')
      log(mode === 'descritivo' ? `Curso: ${desc.nomeCurso}` : `Ficha: ${ficha.nome}`, 'ok')
      log((mode === 'descritivo' ? docText : fichaDocText) ? 'Documento de referência carregado.' : 'Sem documento de referência.', 'ok')

      const payload = mode === 'descritivo'
        ? { mode, form: { ...desc, chTotal }, documentText: docText }
        : { mode, form: ficha, documentText: fichaDocText }

      log('Chamando API...', 'info')
      const response = await fetch('/api/generate', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(payload),
      })

      const data = await response.json()

      if (!response.ok) {
        const errMsg = typeof data?.error === 'string' ? data.error : JSON.stringify(data?.error || data)
        throw new Error(errMsg)
      }

      if (!data.result) throw new Error('A API não retornou resultado válido.')

      log('Resposta recebida. ✓', 'ok')
      log('Montando DOCX...', 'info')

      const blob = mode === 'descritivo'
        ? await buildDescritivoDocx(data.result, { ...desc, chTotal })
        : await buildFichaDocx(data.result, ficha)

      const baseName = (mode === 'descritivo' ? desc.nomeCurso : ficha.nome)
        .replace(/[^a-zA-Z0-9\s-]/g, '').trim().replace(/\s+/g, '-')
      downloadBlob(blob, `${mode === 'descritivo' ? 'Descritivo' : 'Ficha-Produto'}-${baseName || 'SENAI'}.docx`)

      log('Documento gerado. Download iniciado. ✓', 'ok')
      setResultLabel('Documento pronto!')
      await saveToFirestore(mode, mode === 'descritivo' ? desc.nomeCurso : ficha.nome, data.result)
      log('Salvo no histórico Firebase. ✓', 'ok')
    } catch (error: any) {
      const msg = typeof error?.message === 'string' ? error.message : JSON.stringify(error)
      log('Erro: ' + msg, 'err')
      setResultLabel('Não foi possível concluir.')
    } finally {
      setLoading(false)
    }
 }
}
  if (checking) return <main className="app-loading">Carregando...</main>

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
          <button className="theme-btn" onClick={() => setTheme(t => t === 'dark' ? 'light' : 'dark')}>
            {theme === 'dark' ? '☀️ Modo claro' : '🌙 Modo escuro'}
          </button>
          <span className="user-pill">{user?.email}</span>
          <button className="ghost-btn" onClick={() => signOut(auth)}>Sair</button>
        </div>
      </header>

      <section className="app-hero glass-panel">
        <div>
          <div className="eyebrow">Ambiente protegido</div>
          <h1>Geração documental SENAI</h1>
          <p>Produza descritivos de curso e fichas de produto no padrão SENAI com IA.</p>
        </div>
      </section>

      <div className="workspace">
        <aside className="sidebar glass-panel">
          <button className={`navitem ${mode === 'descritivo' ? 'active' : ''}`} onClick={() => setMode('descritivo')}>📄 Descritivo de Curso</button>
          <button className={`navitem ${mode === 'ficha' ? 'active' : ''}`} onClick={() => setMode('ficha')}>🗂️ Ficha de Produto</button>
        </aside>

        <section className="content-stack">
          {mode === 'descritivo' ? (
            <>
              <div className="glass-panel form-card">
                <div className="section-head"><span>1</span><div><h2>Dados do curso</h2><p>Informe os dados principais do descritivo.</p></div></div>
                <div className="field"><label>Nome do curso *</label><input value={desc.nomeCurso} onChange={e => setDesc({ ...desc, nomeCurso: e.target.value })} placeholder="Ex: Assistente em Tecnologias da Indústria 4.0" /></div>
                <div className="grid3">
                  <div className="field">
                    <label>Carga horária específica *</label>
                    <input
                      value={desc.cargaHoraria}
                      onChange={e => setDesc({ ...desc, cargaHoraria: e.target.value })}
                      placeholder="Ex: 176h"
                    />
                  </div>
                  <div className="field">
                    <label>Carga horária total</label>
                    <input
                      readOnly
                      value={chTotal || '—'}
                      style={{ background: 'var(--input-bg)', opacity: .8, cursor: 'default', fontWeight: 700, color: chTotal ? 'var(--senai-red)' : 'var(--muted)' }}
                    />
                  </div>
                  <div className="field"><label>CBO/Ocupação</label><input value={desc.cbo} onChange={e => setDesc({ ...desc, cbo: e.target.value })} placeholder="Opcional" /></div>
                </div>
                <div className="grid2">
                  <div className="field"><label>Modalidade</label><select value={desc.modalidade} onChange={e => setDesc({ ...desc, modalidade: e.target.value })}>{modalidades.map(i => <option key={i}>{i}</option>)}</select></div>
                  <div className="field"><label>Eixo tecnológico</label><select value={desc.eixo} onChange={e => setDesc({ ...desc, eixo: e.target.value })}>{eixos.map(i => <option key={i}>{i}</option>)}</select></div>
                </div>
                <div className="grid3">
                  <div className="field"><label>Escolaridade</label><input value={desc.escolaridade} onChange={e => setDesc({ ...desc, escolaridade: e.target.value })} /></div>
                  <div className="field"><label>Idade mínima</label><input value={desc.idadeMinima} onChange={e => setDesc({ ...desc, idadeMinima: e.target.value })} /></div>
                  <div className="field"><label>Cliente</label><input value={desc.cliente} onChange={e => setDesc({ ...desc, cliente: e.target.value })} placeholder="Opcional" /></div>
                </div>
                <div className="field"><label>Outros requisitos</label><input value={desc.requisitos} onChange={e => setDesc({ ...desc, requisitos: e.target.value })} placeholder="Opcional" /></div>
              </div>

              <div className="glass-panel form-card">
                <div className="section-head"><span>2</span><div><h2>Configurações pedagógicas</h2><p>Ative somente o que fará parte da proposta.</p></div></div>
                <SwitchRow title="Módulo Básico" desc="Sustentabilidade, SST, Qualidade e Indústria 4.0 (+40h)." checked={desc.moduloBasico} onChange={v => setDesc({ ...desc, moduloBasico: v })} />
                <div className="switch-row">
                  <div><b>Aprender a Empreender</b><p>Conteúdo atualizado com empreendedorismo digital e mobile.</p></div>
                  <div className="switch-controls">
                    <select value={desc.aprenderEmpreenderCh} onChange={e => setDesc({ ...desc, aprenderEmpreenderCh: e.target.value })}><option>16</option><option>24</option><option>32</option><option>40</option></select>
                    <label className="switch"><input type="checkbox" checked={desc.aprenderEmpreender} onChange={e => setDesc({ ...desc, aprenderEmpreender: e.target.checked })} /><span /></label>
                  </div>
                </div>
                <SwitchRow title="EJA Profissionalizante" desc="Aplica 18 anos, MRS, 80% presencial e 20% assíncrono." checked={desc.eja} onChange={v => setDesc({ ...desc, eja: v })} />
              </div>

              <div className="glass-panel form-card">
                <div className="section-head"><span>3</span><div><h2>Unidades curriculares</h2><p>Opcional: defina quantidade ou nomes específicos.</p></div></div>
                <div className="grid2">
                  <div className="field"><label>Qtd. de UCs técnicas</label><input value={desc.numUC} onChange={e => setDesc({ ...desc, numUC: e.target.value })} placeholder="Ex: 3" /></div>
                  <SwitchRow title="Informar nomes das UCs" desc="Use quando já houver matriz definida." checked={desc.usarNomesUC} onChange={v => setDesc({ ...desc, usarNomesUC: v })} compact />
                </div>
                {desc.usarNomesUC && <UcEditor items={desc.ucs} onChange={ucs => setDesc({ ...desc, ucs })} />}
              </div>

              <UploadCard fileName={docName} text={docText} onFile={f => handleFile(f, 'descritivo')} onClear={() => { setDocName(''); setDocText('') }} />
            </>
          ) : (
            <>
              <div className="glass-panel form-card">
                <div className="section-head"><span>1</span><div><h2>Ficha de produto</h2><p>Preencha os dados mínimos para gerar a ficha.</p></div></div>
                <div className="field"><label>Nome do curso/produto *</label><input value={ficha.nome} onChange={e => setFicha({ ...ficha, nome: e.target.value })} placeholder="Ex: Criação de Prompts e Agentes de IA" /></div>
                <div className="grid2">
                  <div className="field"><label>Carga horária *</label><input value={ficha.ch} onChange={e => setFicha({ ...ficha, ch: e.target.value })} placeholder="Ex: 80 horas" /></div>
                  <div className="field"><label>Modalidade</label><select value={ficha.modalidade} onChange={e => setFicha({ ...ficha, modalidade: e.target.value })}>{modalidades.map(i => <option key={i}>{i}</option>)}</select></div>
                </div>
                <div className="grid2">
                  <div className="field"><label>Eixo tecnológico</label><select value={ficha.eixo} onChange={e => setFicha({ ...ficha, eixo: e.target.value })}>{eixos.map(i => <option key={i}>{i}</option>)}</select></div>
                  <div className="field"><label>Frequência mínima</label><input value={ficha.frequencia} onChange={e => setFicha({ ...ficha, frequencia: e.target.value })} /></div>
                </div>
                <div className="field"><label>Público-alvo</label><input value={ficha.publico} onChange={e => setFicha({ ...ficha, publico: e.target.value })} placeholder="Profissionais e estudantes..." /></div>
                <div className="field"><label>Certificado emitido</label><input value={ficha.certificado} onChange={e => setFicha({ ...ficha, certificado: e.target.value })} placeholder="Opcional" /></div>
              </div>
              <UploadCard fileName={fichaDocName} text={fichaDocText} onFile={f => handleFile(f, 'ficha')} onClear={() => { setFichaDocName(''); setFichaDocText('') }} />
            </>
          )}

          <div className="glass-panel generate-panel">
            <div>
              <h2>Gerar documento</h2>
              <p>{mode === 'descritivo' ? 'Descritivo completo em .docx no padrão SENAI.' : 'Ficha de produto em .docx no padrão SENAI.'}</p>
            </div>
            <button className="primary-btn" disabled={!ready || loading} onClick={generate}>
              {loading ? 'Gerando...' : '⚡ Gerar e baixar'}
            </button>
          </div>

          {(logs.length > 0 || resultLabel) && (
            <div className="glass-panel log-panel">
              <div className="section-head small"><span>✓</span><div><h2>{resultLabel || 'Processamento'}</h2><p>Acompanhe as etapas da geração.</p></div></div>
              <div className="log-box">{logs.map((item, i) => <div key={i} className={item.type}>› {item.text}</div>)}</div>
            </div>
          )}

          {history.length > 0 && (
            <div className="glass-panel history-panel">
              <div className="section-head">
                <span>📋</span>
                <div><h2>Histórico</h2><p>Documentos gerados e salvos no Firebase.</p></div>
                <button
                  onClick={deleteAllHistory}
                  style={{ marginLeft: 'auto', background: 'none', border: '1px solid rgba(200,16,46,.3)', borderRadius: 8, padding: '4px 12px', color: 'var(--senai-red)', fontSize: 12, cursor: 'pointer', flexShrink: 0 }}
                >
                  🗑️ Limpar tudo
                </button>
              </div>
              <div className="history-list">
                {history.slice(0, 10).map(item => (
                  <div className="history-item" key={item.id}>
                    <div className="h-info">
                      <div className="h-name">{item.nome}</div>
                      <div className="h-meta">{item.criadoEm?.toDate?.()?.toLocaleString('pt-BR') || '—'}</div>
                    </div>
                    <span className="h-type">{item.tipo}</span>
                    <button
                      onClick={() => deleteHistoryItem(item.id)}
                      title="Apagar este registro"
                      style={{ background: 'none', border: 'none', color: 'var(--muted2)', cursor: 'pointer', fontSize: 16, padding: '0 4px', flexShrink: 0 }}
                    >✕</button>
                  </div>
                ))}
              </div>
            </div>
          )}
        </section>
      </div>

      <footer className="app-footer">
        Criado por <strong>Paulo da Silva Filho</strong> — Especialista em TI da GEP · SENAI Bahia · 2026
      </footer>
    </main>
  )
}

/* ── SUB COMPONENTS ──────────────────────────────────────── */
function SwitchRow({ title, desc, checked, onChange, compact = false }: { title: string; desc: string; checked: boolean; onChange: (v: boolean) => void; compact?: boolean }) {
  return (
    <div className={`switch-row ${compact ? 'compact-switch' : ''}`}>
      <div><b>{title}</b><p>{desc}</p></div>
      <label className="switch"><input type="checkbox" checked={checked} onChange={e => onChange(e.target.checked)} /><span /></label>
    </div>
  )
}

function UcEditor({ items, onChange }: { items: { nome: string; ch: string }[]; onChange: (items: { nome: string; ch: string }[]) => void }) {
  return (
    <div className="uc-list">
      {items.map((item, i) => (
        <div className="uc-row" key={i}>
          <input value={item.nome} onChange={e => onChange(items.map((it, j) => j === i ? { ...it, nome: e.target.value } : it))} placeholder="Nome da UC" />
          <input value={item.ch} onChange={e => onChange(items.map((it, j) => j === i ? { ...it, ch: e.target.value } : it))} placeholder="40h" />
          <button onClick={() => onChange(items.filter((_, j) => j !== i))}>×</button>
        </div>
      ))}
      <button className="ghost-btn" onClick={() => onChange([...items, { nome: '', ch: '' }])}>+ Adicionar UC</button>
    </div>
  )
}

function UploadCard({ fileName, text, onFile, onClear }: { fileName: string; text: string; onFile: (f: File | undefined) => void; onClear: () => void }) {
  return (
    <div className="glass-panel form-card">
      <div className="section-head"><span>4</span><div><h2>Documento de referência</h2><p>Opcional: use itinerários, fichas anteriores, catálogos ou arquivos técnicos.</p></div></div>
      <label className="upload-box">
        <input type="file" accept=".docx,.pdf,.txt" onChange={e => onFile(e.target.files?.[0])} />
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
