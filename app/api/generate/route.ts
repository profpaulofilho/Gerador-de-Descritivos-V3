import Anthropic from '@anthropic-ai/sdk'
import { NextResponse } from 'next/server'

const APRENDER_EMPREENDEDOR = `
Parte 1: Aprender a Empreender — Carga Horária sugerida: 24h.
Objetivo geral: Desenvolver atitudes empreendedoras e aplicar conhecimentos básicos de marketing, finanças e inovação, com uso de tecnologias digitais acessíveis, para planejar e divulgar pequenos negócios ou soluções criativas voltadas à realidade socioprofissional.
Capacidades: Identificar características e atitudes empreendedoras; reconhecer potencial empreendedor; utilizar ferramentas digitais mobile para organizar ideias; relacionar valores pessoais com propósito de negócio; aplicar Canvas simplificado; realizar pesquisa de mercado por canais digitais; identificar oportunidades locais; usar Google Forms e enquetes; aplicar marketing digital e atendimento por apps; produzir conteúdo visual; utilizar CapCut, Canva, InShot; comunicar-se com clientes por WhatsApp Business, Instagram e Facebook; organizar orçamento pessoal e do negócio; registrar receitas/despesas; identificar custos, precificar e analisar viabilidade; aplicar educação financeira; compreender direitos trabalhistas básicos, SST, igualdade de gênero e combate ao trabalho escravo e infantil; relacionar empreendedorismo com responsabilidade social e legal.
Conhecimentos: Atitude Empreendedora e Autoconhecimento (5h); Empreendendo com o que tenho (10h); Divulgação, Atendimento e Direitos na Prática Profissional (5h); Finanças Pessoais e do Negócio (4h).
`

const EJA_ORIENTACOES = `
EJA Profissionalizante: oferta vinculada a cursos de Qualificação Profissional. Deve contemplar integração entre Ensino Médio e Qualificação Profissional; organização curricular articulada; Metodologia de Reconhecimento de Saberes (MRS); distribuição de 80% presencial e 20% assíncrona; contextualização com jovens e adultos trabalhadores; estratégias pedagógicas que valorizem experiências prévias; idade mínima de 18 anos; não ter concluído o Ensino Médio; requisitos da qualificação vinculada quando aplicável. Incluir avaliação por competências com situações de aprendizagem, projetos integradores, portfólios, estudos de caso, avaliações práticas, autoavaliação e rubricas; média mínima de referência 5,0 quando aplicável; frequência mínima 75% ou normativa vigente; recuperação; permanência e êxito; atividades assíncronas como videoaulas, fóruns, pesquisas, portfólios, desafios práticos e questionários; certificação; acompanhamento pedagógico; competências socioemocionais.
`

function buildDescritivoPrompt(payload: any) {
  const f = payload.form || {}
  const docText = (payload.documentText || '').slice(0, 4000)
  const ucInstruction = f.ucs?.length
    ? `Use EXATAMENTE estas UCs técnicas: ${f.ucs.map((u: any, i: number) => `UC ${i + 1}: ${u.nome} (${u.ch || 'carga a definir'})`).join('; ')}`
    : f.numUC
      ? `Gere exatamente ${f.numUC} UCs técnicas, distribuindo a carga horária técnica.`
      : 'Defina a melhor organização de UCs técnicas conforme a natureza do curso.'

  return `Você é especialista SENAI Bahia em MSEP, desenvolvimento de cursos e documentação pedagógica.

Crie um DESCRITIVO DE CURSO no padrão SENAI Bahia, robusto e tecnicamente coerente.

ATENÇÃO CRÍTICA: Retorne SOMENTE o JSON puro, sem nenhum texto antes ou depois, sem markdown, sem explicações. O JSON deve estar 100% completo e válido.

DADOS DO CURSO:
Nome: ${f.nomeCurso}
Carga horária técnica/base: ${f.cargaHoraria}
Modalidade: ${f.modalidade}
Eixo tecnológico: ${f.eixo}
CBO/Ocupação: ${f.cbo || 'não informado'}
Escolaridade mínima: ${f.escolaridade || 'não informado'}
Idade mínima: ${f.idadeMinima || 'não informado'}
Cliente: ${f.cliente || 'não informado'}
Outros requisitos: ${f.requisitos || 'não informado'}
Módulo Básico: ${f.moduloBasico ? 'sim, incluir 40h' : 'não'}
Aprender a Empreender: ${f.aprenderEmpreender ? 'sim, incluir conteúdo atualizado' : 'não'}
EJA Profissionalizante: ${f.eja ? 'sim' : 'não'}
${ucInstruction}

CONTEÚDO ATUALIZADO DE APRENDER A EMPREENDER:
${f.aprenderEmpreender ? APRENDER_EMPREENDEDOR : 'Não incluir este módulo.'}

ORIENTAÇÕES EJA:
${f.eja ? EJA_ORIENTACOES : 'Não aplicar EJA.'}

DOCUMENTO NORTEADOR ENVIADO PELO USUÁRIO:
---
${docText || 'Não foi enviado documento norteador. Gere com base nos dados informados e melhores práticas SENAI.'}
---

Retorne SOMENTE JSON válido, sem markdown, no formato:
{
  "tipo":"descritivo",
  "justificativa":"2 a 4 parágrafos",
  "objetivo":"objetivo geral do curso",
  "descricao":"descrição do curso",
  "publicoAlvo":"público-alvo",
  "perfilSaida":["competência 1","competência 2","competência 3","competência 4","competência 5","competência 6"],
  "metodologia":"metodologia MSEP SENAI com situações de aprendizagem, práticas e projeto integrador quando adequado",
  "perfilDocente":"formação e experiência recomendada",
  "criteriosAvaliacao":"critérios de avaliação completos",
  "criteriosCertificacao":"critérios de certificação",
  "modulos":[{"nome":"MÓDULO ESPECÍFICO I","chModulo":"Xh","ucs":[{"nome":"UC técnica","ch":"Xh","objetivo":"objetivo","capacidades":["capacidade"],"conhecimentos":["conhecimento"]}]}],
  "infraestrutura":{"ambiente":["item"],"equipamentos":[["equipamento","quantidade"]],"softwares":["item"],"materiais":"texto","epis":["item"],"conectividade":["item"]},
  "referencias":["referência"],
  "eja":{"aplicar":${f.eja ? 'true' : 'false'},"diretrizes":"texto","atividadesAssincronas":["atividade"],"permanenciaExito":["ação"]}
}

Regras: não invente dados legais específicos; use o documento enviado como base principal quando existir; mantenha linguagem formal, técnica e aplicável ao SENAI Bahia. Seja conciso para garantir JSON completo dentro do limite de tokens.`
}

function buildFichaPrompt(payload: any) {
  const f = payload.form || {}
  const docText = (payload.documentText || '').slice(0, 4000)
  return `Você é especialista SENAI Bahia. Crie uma FICHA DE PRODUTO de curso no padrão SENAI.

ATENÇÃO CRÍTICA: Retorne SOMENTE o JSON puro, sem nenhum texto antes ou depois, sem markdown, sem explicações.

DADOS:
Nome: ${f.nome}
Carga horária: ${f.ch}
Modalidade: ${f.modalidade}
Eixo tecnológico: ${f.eixo}
Público-alvo: ${f.publico || 'não informado'}
Certificado: ${f.certificado || 'gerar conforme modalidade e nome'}
Frequência mínima: ${f.frequencia || '75%'}

DOCUMENTO DE REFERÊNCIA:
---
${docText || 'Não enviado. Criar ficha coerente com o tema.'}
---

Retorne SOMENTE JSON válido, sem markdown:
{
  "tipo":"ficha",
  "nome":"nome do curso",
  "descricao":"parágrafo de descrição do curso no estilo ficha de produto SENAI",
  "cargaHoraria":"carga horária",
  "participantes":"público-alvo completo",
  "certificado":"certificado emitido",
  "frequencia":"frequência mínima",
  "programacao":[{"modulo":"Nome do módulo (Xh)","topicos":["tópico 1","tópico 2"]}]
}

Seja conciso para garantir JSON completo.`
}

function extractJson(text: string) {
  const cleaned = text.trim()
    .replace(/^```json\s*/i, '')
    .replace(/^```\s*/i, '')
    .replace(/```$/i, '')
    .trim()

  // Tenta parse direto
  try { return JSON.parse(cleaned) } catch {}

  // Tenta extrair objeto JSON da resposta
  const start = cleaned.indexOf('{')
  const end = cleaned.lastIndexOf('}')
  if (start >= 0 && end > start) {
    try { return JSON.parse(cleaned.slice(start, end + 1)) } catch {}
  }

  // Tenta reparar JSON truncado
  if (start >= 0) {
    try {
      let partial = cleaned.slice(start)
      const lastNewline = partial.lastIndexOf('\n')
      if (lastNewline > 0) partial = partial.slice(0, lastNewline)
      const opens = (partial.match(/\[/g) || []).length - (partial.match(/\]/g) || []).length
      const objOpens = (partial.match(/\{/g) || []).length - (partial.match(/\}/g) || []).length
      partial += ']'.repeat(Math.max(0, opens))
      partial += '}'.repeat(Math.max(0, objOpens))
      return JSON.parse(partial)
    } catch {}
  }

  // ✅ NOVO: loga os primeiros e últimos 500 chars para diagnóstico no Vercel
  console.error('JSON inválido — início:', cleaned.slice(0, 500))
  console.error('JSON inválido — fim:', cleaned.slice(-500))
  throw new Error('A resposta da IA não veio em JSON válido.')
}

export async function POST(req: Request) {
  try {
    if (!process.env.ANTHROPIC_API_KEY) {
      return NextResponse.json({ error: 'ANTHROPIC_API_KEY não configurada na Vercel.' }, { status: 500 })
    }
    const payload = await req.json()
    const prompt = payload.mode === 'ficha' ? buildFichaPrompt(payload) : buildDescritivoPrompt(payload)
    const anthropic = new Anthropic({ apiKey: process.env.ANTHROPIC_API_KEY })
    const message = await anthropic.messages.create({
      model: process.env.ANTHROPIC_MODEL || 'claude-haiku-4-5',
      max_tokens: payload.mode === 'ficha' ? 2000 : 6000,
      messages: [{ role: 'user', content: prompt }],
    })

    // ✅ NOVO: loga stop_reason para diagnóstico
    console.log('stop_reason:', message.stop_reason)
    console.log('uso de tokens:', JSON.stringify(message.usage))

    const text = message.content.map((c: any) => c.type === 'text' ? c.text : '').join('\n')

    // ✅ NOVO: retorna erro descritivo com stop_reason se não for JSON
    if (message.stop_reason === 'max_tokens') {
      console.error('Resposta cortada por max_tokens. Tokens usados:', message.usage)
      return NextResponse.json({
        error: 'O descritivo gerado foi muito longo. Tente reduzir o número de UCs ou a carga horária.'
      }, { status: 500 })
    }

    const json = extractJson(text)
    return NextResponse.json({ result: json })
  } catch (error: any) {
    return NextResponse.json({ error: error?.message || 'Erro ao gerar conteúdo.' }, { status: 500 })
  }
}
