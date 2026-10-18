// Fonction Supabase « lire-documents » (V9, 18/10/2026 ; V10 : nouveaux essais si Claude ne répond pas, avancement visible dans l'appli ;
// V10b : documents envoyés d'abord à l'API Files d'Anthropic, la demande de lecture ne transporte plus les PDF scannés lourds).
// Lit les documents de début de chantier déposés dans l'appli (prévisite, liste du matériel, bons de commande…)
// et remplit les rubriques VIDES de la fiche mission. Ce qui a été saisi à la main n'est jamais remplacé,
// la variante (avec ou sans démolition) n'est jamais changée : un écart part dans « À vérifier ».
// Appelée par l'appli (AFER ou DOMetVIE) après un dépôt de documents, ou par le bouton « Remplir la mission depuis les documents ».
// Secret à créer dans Supabase > Edge Functions > Secrets : ANTHROPIC_API_KEY.
// La fonction vérifie elle-même l'appelant (connexion AFER ou lien DOMetVIE) : la vérification JWT de la plateforme peut rester coupée.
import { createClient } from 'npm:@supabase/supabase-js@2.117.2'
import Anthropic, { toFile } from 'npm:@anthropic-ai/sdk@0.135.0'
import { Buffer } from 'node:buffer'

const MODELE = 'claude-opus-5-5'
// Supabase arrête la fonction après 150 s (offre gratuite) : la lecture et ses nouveaux essais tiennent dans ce délai
const DELAI_MS = 140_000
const MAX_OCTETS = 20 * 1024 * 1024 // sous la limite de 32 Mo par requête, encodage base64 compris
const IMAGES = ['image/jpeg', 'image/png', 'image/gif', 'image/webp']
const CHAMPS = ['ville', 'adresse', 'etage', 'acces', 'client', 'tel', 'dispo', 'conducteur', 'retraits', 'materiel', 'travaux', 'verifs']
const NOMS: Record<string, string> = {
  ville: 'ville', adresse: 'adresse', etage: 'étage', acces: 'accès', client: 'client', tel: 'téléphone du client',
  dispo: 'disponibilité', conducteur: 'conduite de travaux', retraits: 'retraits du matériel', materiel: 'matériel',
  travaux: 'travaux', verifs: 'points à vérifier',
}

const txt = (description: string) => ({ type: 'string', description })
const liste = (description: string) => ({ type: 'array', items: { type: 'string' }, description })
const SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: [...CHAMPS, 'variante'],
  properties: {
    ville: txt('Commune du chantier, sans code postal'),
    adresse: txt('Adresse complète du chantier : numéro, rue, code postal, commune'),
    etage: txt("Étage et précisions pour trouver le logement (bâtiment, escalier, ascenseur)"),
    acces: txt("Code d'entrée, digicode ou interphone"),
    client: txt('Nom du bénéficiaire, chez qui ont lieu les travaux'),
    tel: txt('Téléphone du bénéficiaire'),
    dispo: txt('Disponibilités du bénéficiaire ou date prévue des travaux, telles qu’écrites'),
    conducteur: txt('Téléphone, et nom si indiqué, du conducteur de travaux ou de l’interlocuteur DOMetVIE'),
    variante: {
      type: 'string', enum: ['avec', 'sans', 'inconnu'],
      description: "« avec » s'il faut déposer ou démolir l'existant (baignoire, douche, faïence…), « sans » si les documents disent qu'il n'y a rien à démolir, sinon « inconnu »",
    },
    retraits: {
      type: 'array',
      description: 'Lieux où récupérer le matériel (fournisseur, agence, dépôt), un par lieu',
      items: {
        type: 'object', additionalProperties: false, required: ['lieu', 'adresse', 'code', 'articles'],
        properties: {
          lieu: txt('Nom du fournisseur ou du lieu, ex. CEDEO'),
          adresse: txt('Adresse du lieu de retrait'),
          code: txt('Numéro ou code de retrait ou de commande, recopié à l’identique'),
          articles: txt('Résumé court de ce qui est à récupérer à cet endroit, ex. « 11 articles : receveur, colonne, paroi… »'),
        },
      },
    },
    materiel: liste('Matériel à poser, un article par élément, avec référence, dimensions et quantité quand elles sont écrites'),
    travaux: liste('Travaux à réaliser, une action courte par élément, dans l’ordre du chantier'),
    verifs: liste('Vraies incohérences entre documents (dimensions, références, quantités) ou informations indispensables qui manquent, une phrase courte par point'),
  },
}
const CONSIGNE = `Ces documents sont transmis par DOMetVIE, donneur d'ordre, à AFER, entreprise de travaux, pour un chantier de rénovation de salle de bain adaptée aux personnes à mobilité réduite : compte rendu de prévisite technique, liste des produits à récupérer, bons de commande, plans ou photos.
Remplis la fiche mission du technicien à partir de ces documents uniquement.
- N'invente rien : une information absente des documents reste vide (chaîne vide ou liste vide).
- Recopie à l'identique les adresses, numéros de téléphone, codes d'accès et numéros ou codes de retrait.
- Matériel : un article par élément, avec référence, dimensions et quantité quand elles sont écrites.
- Travaux : des actions courtes, compréhensibles par le technicien, dans l'ordre du chantier.
- Points à vérifier : seulement les vraies incohérences entre documents ou les informations indispensables qui manquent.`

// deno-lint-ignore no-explicit-any
type Admin = any
type Doc = { id: string; name?: string; cat?: string; at?: string }
type Fiche = Record<string, unknown> & { auto?: { docs?: string[]; champs?: string[] }; variante?: string; verifs?: unknown }
type Lu = Record<string, unknown> & { variante?: string }
type Bloc = Record<string, unknown>
export type Piece = { nom: string; type: string; octets: Uint8Array<ArrayBuffer> }

export const vide = (v: unknown) => v == null || (typeof v === 'string' && !v.trim()) || (Array.isArray(v) && v.length === 0)
const court = (e: unknown) => String((e as Error)?.message || e).slice(0, 160)
// chemin du fichier dans le stockage, comme dans l'appli (fpath)
const chemin = (ch: string, id: string) => `${ch}/doc/${('doc:' + id).replace(/[^A-Za-z0-9._-]/g, '_')}`
function typeDe(blob: Blob, nom = '') {
  const t = (blob.type || '').toLowerCase()
  if (t === 'application/pdf' || IMAGES.includes(t)) return t
  const ext = nom.toLowerCase().split('.').pop() || ''
  return ({ pdf: 'application/pdf', jpg: 'image/jpeg', jpeg: 'image/jpeg', png: 'image/png', gif: 'image/gif', webp: 'image/webp' } as Record<string, string>)[ext] || ''
}

// Ajoute à la fiche ce qui y manque ; renvoie la nouvelle fiche et les rubriques remplies.
export function fusion(fiche: Fiche, lu: Lu, vus: string[], ignores: string[]) {
  const f: Fiche = { ...fiche }, remplis: string[] = []
  for (const k of CHAMPS) {
    let v = lu[k]
    if (k === 'retraits') {
      v = (Array.isArray(v) ? v : []).map((r: Record<string, unknown>) => ({
        lieu: String(r?.lieu ?? '').trim(), adresse: String(r?.adresse ?? '').trim(), code: String(r?.code ?? '').trim(), articles: String(r?.articles ?? '').trim(),
      })).filter((r) => r.lieu || r.adresse)
    } else if (Array.isArray(v)) v = v.map((x) => String(x).trim()).filter(Boolean)
    else v = typeof v === 'string' ? v.trim() : ''
    if (vide(f[k]) && !vide(v)) { f[k] = v; remplis.push(k) }
  }
  const lv = lu.variante, fv = f.variante || 'avec'
  if ((lv === 'avec' || lv === 'sans') && lv !== fv) {
    const note = `Variante : les documents indiquent « ${lv} démolition », la fiche dit « ${fv} démolition ».`
    const vf = Array.isArray(f.verifs) ? f.verifs as string[] : []
    if (!vf.includes(note)) { f.verifs = [...vf, note]; if (!remplis.includes('verifs')) remplis.push('verifs') }
  }
  const a = fiche.auto || {}
  f.auto = {
    le: new Date().toISOString(),
    docs: [...new Set([...(a.docs || []), ...vus])],
    champs: [...new Set([...(a.champs || []), ...remplis])],
    ignores,
  } as Fiche['auto']
  return { fiche: f, remplis }
}

// Message lisible pour l'appli (au lieu de la page d'erreur brute renvoyée par le service)
// Détail renvoyé par l'API (type et message), pour savoir exactement ce qui bloque
export function detailApi(e: unknown): string {
  const st = (e as { status?: number })?.status
  const x = (e as { error?: { error?: { type?: string; message?: string } } })?.error?.error
  if (!st) return court(e)
  return `${st}${x?.type ? ' ' + x.type : ''}${x?.message ? ' : ' + String(x.message).slice(0, 160) : ''}`
}
export function messageErreur(e: unknown): string {
  const essais = (e as { essais?: string[] })?.essais
  const base = messageBase(e)
  return essais && essais.length ? `${base} [réponse de Claude : ${essais.join(' | ')}]` : base
}
function messageBase(e: unknown): string {
  const st = (e as { status?: number })?.status
  if (e instanceof Anthropic.APIConnectionTimeoutError) return 'lecture trop longue, relance avec « Remplir la mission depuis les documents »'
  if (e instanceof Anthropic.APIConnectionError) return 'Claude injoignable pour le moment, réessaie dans quelques minutes'
  if (st === 401) return 'clé API Anthropic refusée (vérifier le secret ANTHROPIC_API_KEY dans Supabase)'
  if (st === 403) return 'accès refusé par Anthropic pour cette demande (droits de la clé ou du modèle)'
  if (st === 429) return 'trop de demandes envoyées à Claude, réessaie dans une minute'
  if (st && st >= 500) return `service Claude momentanément indisponible (erreur ${st}), réessaie dans quelques minutes`
  if (st === 400 || st === 413) {
    const m = (e as { error?: { error?: { message?: string } } })?.error?.error?.message
    return 'documents refusés par Claude' + (m ? ' : ' + String(m).slice(0, 120) : '')
  }
  return court(e)
}
const repassable = (e: unknown) => {
  const st = (e as { status?: number })?.status
  return (e instanceof Anthropic.APIConnectionError && !(e instanceof Anthropic.APIConnectionTimeoutError)) || st === 429 || (!!st && st >= 500)
}
const pause = (ms: number) => new Promise((r) => setTimeout(r, ms))

type Essai = (n: number, total: number) => Promise<void> | void
// Trois essais au plus dans le délai : si le service répond par une erreur passagère (502, 529…), on relance ;
// le dernier essai passe sans l'option de repli (beta), au cas où ce serait elle qui bloque.
// Les documents passent d'abord par l'API Files (envoi séparé, plus sûr pour les gros scans), puis sont supprimés.
// Si l'envoi séparé échoue, ils repartent dans la demande elle-même (base64), comme en V9.
async function contenuDe(client: Anthropic, pieces: Piece[], ids: string[]): Promise<Bloc[]> {
  const contenu: Bloc[] = []
  for (const [i, p] of pieces.entries()) {
    let source: Bloc
    try {
      const f = await client.files.upload({ file: await toFile(p.octets, p.nom, { type: p.type }) })
      ids.push(f.id); source = { type: 'file', file_id: f.id }
    } catch (e) {
      console.error('lire-documents : envoi séparé impossible, document joint à la demande', p.nom, court(e))
      source = { type: 'base64', media_type: p.type, data: Buffer.from(p.octets).toString('base64') }
    }
    contenu.push({ type: 'text', text: `Document ${i + 1} : ${p.nom}` })
    contenu.push({ type: p.type === 'application/pdf' ? 'document' : 'image', source })
  }
  contenu.push({ type: 'text', text: CONSIGNE })
  return contenu
}

export async function appelClaude(pieces: Piece[], surEssai?: Essai, fin = Date.now() + DELAI_MS): Promise<Lu> {
  const apiKey = Deno.env.get('ANTHROPIC_API_KEY')
  if (!apiKey) throw new Error('clé ANTHROPIC_API_KEY absente des secrets de la fonction')
  const ids: string[] = []
  const envoi = new Anthropic({ apiKey, timeout: 60_000, maxRetries: 2 })
  try {
    const contenu = await contenuDe(envoi, pieces, ids)
    return await lecture(apiKey, contenu, surEssai, fin)
  } finally {
    for (const id of ids) await envoi.files.delete(id).catch(() => {})
  }
}

async function lecture(apiKey: string, contenu: Bloc[], surEssai: Essai | undefined, fin: number): Promise<Lu> {
  const ESSAIS = 3
  let derniere: unknown
  const essais: string[] = []
  for (let i = 1; i <= ESSAIS; i++) {
    const reste = fin - Date.now()
    if (reste < 25_000) break
    if (i > 1) await surEssai?.(i, ESSAIS)
    const client = new Anthropic({ apiKey, timeout: reste - 5_000, maxRetries: 0 })
    const demande = {
      model: MODELE,
      max_tokens: 16000,
      output_config: { effort: 'low', format: { type: 'json_schema', schema: SCHEMA } }, // extraction : effort bas = lecture plus rapide
      // deno-lint-ignore no-explicit-any
      messages: [{ role: 'user', content: contenu as any }],
    }
    try {
      const rep = i < ESSAIS
        // si le modèle décline, l'API relance la demande sur le modèle de repli recommandé
        // deno-lint-ignore no-explicit-any
        ? await client.beta.messages.stream({ ...demande, betas: ['server-side-fallback-2026-07-01'], fallbacks: 'default' } as any).finalMessage()
        // deno-lint-ignore no-explicit-any
        : await client.messages.stream(demande as any).finalMessage()
      return lireReponse(rep as unknown as { stop_reason: string; content: { type: string; text?: string }[] })
    } catch (e) {
      derniere = e
      essais.push(`essai ${i} : ${detailApi(e)}`)
      console.error('lire-documents : essai', i, (e as { requestID?: string })?.requestID ?? '', detailApi(e))
      if (!repassable(e)) { Object.assign(e as object, { essais }); throw e }
      await pause(Math.min(3_000 * i, Math.max(0, fin - Date.now() - 25_000)))
    }
  }
  if (derniere) Object.assign(derniere as object, { essais })
  throw derniere ?? new Error('lecture trop longue')
}

function lireReponse(rep: { stop_reason: string; content: { type: string; text?: string }[] }): Lu {
  if (rep.stop_reason === 'refusal') throw new Error('lecture refusée par le modèle')
  if (rep.stop_reason === 'max_tokens') throw new Error('réponse trop longue')
  const bloc = rep.content.find((b) => b.type === 'text') as { text: string } | undefined
  if (!bloc) throw new Error('réponse vide')
  return JSON.parse(bloc.text)
}

// Avancement visible dans l'appli (fiche.lecture) : étape, pourcentage, erreur. Lu par AFER et DOMetVIE.
type Etat = { etat: 'en_cours' | 'fini' | 'erreur'; etape?: string; pct?: number; msg?: string; debut?: string }
export async function avancement(admin: Admin, ch: string, e: Etat) {
  const { data: c } = await admin.from('chantiers').select('fiche').eq('id', ch).maybeSingle()
  if (!c) return
  const f = (c.fiche || {}) as Fiche, avant = (f.lecture || {}) as Etat
  await admin.from('chantiers').update({ fiche: { ...f, lecture: { ...avant, ...e, maj: new Date().toISOString() } } }).eq('id', ch)
}

async function alerte(admin: Admin, ch: string, pour: string[], texte: string, type: string) {
  await admin.from('evenements').insert({ chantier_id: ch, pour, texte: texte.slice(0, 500), type, par_role: 'afer' })
}

// Lecture complète : documents -> Claude -> fiche du chantier -> alerte dans l'appli
export async function lire(admin: Admin, ch: string, docs: Doc[], appel = appelClaude) {
  const pas = (e: Partial<Etat>) => avancement(admin, ch, { etat: 'en_cours', ...e }).catch(() => {})
  await pas({ etape: 'Récupération des documents', pct: 5, msg: '', debut: new Date().toISOString() })
  const pieces: Piece[] = [], ignores: string[] = [], vus: string[] = []
  let total = 0, n = 0
  for (const d of [...docs].sort((a, b) => String(a.at).localeCompare(String(b.at)))) {
    vus.push(d.id)
    const nom = d.name || 'document sans nom'
    const { data: blob, error } = await admin.storage.from('fichiers').download(chemin(ch, d.id))
    if (error || !blob) { ignores.push(nom); continue }
    const type = typeDe(blob, d.name)
    if (!type || total + blob.size > MAX_OCTETS) { ignores.push(nom); continue }
    total += blob.size; n++
    pieces.push({ nom, type, octets: new Uint8Array(await blob.arrayBuffer()) })
  }
  if (!n) throw new Error('aucun document lisible (formats lus : PDF, JPEG, PNG, WebP)')
  await pas({ etape: `Lecture de ${n} document${n > 1 ? 's' : ''} par Claude`, pct: 25 })
  const lu = await appel(pieces, (i, total) => pas({ etape: `Claude ne répond pas, nouvel essai (${i}/${total})`, pct: 25 }))
  await pas({ etape: 'Remplissage de la mission', pct: 95 })

  // la fiche a pu changer pendant la lecture : on repart de la version en base
  const { data: c, error } = await admin.from('chantiers').select('fiche').eq('id', ch).maybeSingle()
  if (error) throw error
  if (!c) return { remplis: [] } // chantier supprimé entre-temps
  const { fiche, remplis } = fusion((c.fiche || {}) as Fiche, lu, vus, ignores)
  fiche.lecture = { ...(fiche.lecture as Etat || {}), etat: 'fini', etape: 'Mission remplie', pct: 100, nb: remplis.length, maj: new Date().toISOString() }
  const maj: Record<string, unknown> = { fiche, maj_le: new Date().toISOString() }
  if (remplis.includes('ville')) maj.nom = 'Chantier ' + fiche.ville
  const r = await admin.from('chantiers').update(maj).eq('id', ch)
  if (r.error) throw r.error
  const non = ignores.length ? ` (non lu${ignores.length > 1 ? 's' : ''} : ${ignores.join(', ')})` : ''
  await alerte(admin, ch, remplis.length ? ['afer', 'tech'] : ['afer'],
    remplis.length ? 'Mission remplie depuis les documents : ' + remplis.map((k) => NOMS[k]).join(', ') + non : 'Documents lus : rien de nouveau pour la mission' + non,
    'mission')
  return { remplis }
}

const CORS = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type', 'Access-Control-Allow-Methods': 'POST, OPTIONS' }
const reponse = (corps: unknown, status = 200) => new Response(JSON.stringify(corps), { status, headers: { ...CORS, 'Content-Type': 'application/json' } })
// clés fournies par Supabase à la fonction : nouveau format d'abord, ancien en secours
function cle(...noms: string[]) {
  for (const n of noms) {
    const v = Deno.env.get(n)
    if (!v) continue
    if (v.trim().startsWith('{')) { try { const k = JSON.parse(v); if (k.default) return String(k.default) } catch { /* format inattendu */ } } else return v
  }
  return ''
}
const SANS_SESSION = { auth: { persistSession: false, autoRefreshToken: false, detectSessionInUrl: false } }

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') return new Response('ok', { headers: CORS })
  if (req.method !== 'POST') return reponse({ message: 'Méthode non prise en charge' }, 405)
  const auth = req.headers.get('Authorization') || '', jwt = auth.replace(/^Bearer\s+/i, '')
  const url = Deno.env.get('SUPABASE_URL') || ''
  if (!jwt || jwt === auth) return reponse({ message: 'Connexion requise' }, 401)
  // l'appelant, avec ses droits (connexion AFER ou appareil ouvert par un lien)
  const moi = createClient(url, cle('SUPABASE_PUBLISHABLE_KEYS', 'SUPABASE_PUBLISHABLE_KEY', 'SUPABASE_ANON_KEY'), { ...SANS_SESSION, global: { headers: { Authorization: auth } } })
  const { data: u, error: eu } = await moi.auth.getUser(jwt)
  if (eu || !u?.user) return reponse({ message: 'Connexion requise' }, 401)

  let corps: { chantier_id?: string; force?: boolean } = {}
  try { corps = await req.json() } catch { /* corps vide */ }
  const ch = String(corps.chantier_id || '')
  if (!ch) return reponse({ message: 'Chantier manquant' }, 400)

  // mêmes droits que pour déposer les documents : AFER, ou DOMetVIE sur ce chantier
  const { data: afer } = await moi.rpc('est_afer')
  let autorise = afer === true
  if (!autorise) {
    const { data: acces } = await moi.rpc('mes_acces')
    autorise = ((acces || []) as { chantier_id: string | null; role: string }[])
      .some((a) => a.role === 'domevie' && (a.chantier_id === ch || a.chantier_id === null))
  }
  if (!autorise) return reponse({ message: 'Accès refusé' }, 403)

  const admin = createClient(url, cle('SUPABASE_SECRET_KEYS', 'SUPABASE_SECRET_KEY', 'SUPABASE_SERVICE_ROLE_KEY'), SANS_SESSION)
  const { data: c } = await admin.from('chantiers').select('fiche').eq('id', ch).maybeSingle()
  if (!c) return reponse({ message: 'Chantier introuvable' }, 404)
  const { data: els, error } = await admin.from('elements').select('id,data').eq('chantier_id', ch).eq('type', 'doc')
  if (error) return reponse({ message: 'Lecture des documents impossible pour l’instant' }, 500)
  const docs: Doc[] = ((els || []) as { id: string; data: Doc }[]).map((e) => ({ ...e.data, id: e.id })).filter((d) => (d.cat || 'dmv') === 'dmv')
  if (!docs.length) return reponse({ message: 'Aucun document de début de chantier à lire' })
  const fiche = (c.fiche || {}) as Fiche, deja = new Set(fiche.auto?.docs || [])
  const lc = fiche.lecture as Etat & { maj?: string } | undefined
  if (lc?.etat === 'en_cours' && Date.now() - Date.parse(lc.maj || '') < 160_000) return reponse({ message: 'Lecture des documents déjà en cours' })
  if (!corps.force && docs.every((d) => deja.has(d.id))) return reponse({ message: 'Documents déjà lus' })
  if (!CHAMPS.some((k) => vide(fiche[k]))) return reponse({ message: 'La mission est déjà remplie : rien à compléter' })

  // la lecture prend du temps : on répond tout de suite, le résultat arrive dans l'appli (fiche + alerte)
  const travail = lire(admin, ch, docs).catch(async (e) => {
    console.error('lire-documents', e)
    const m = messageErreur(e)
    await avancement(admin, ch, { etat: 'erreur', etape: 'Lecture impossible', msg: m }).catch(() => {})
    await alerte(admin, ch, ['afer'], 'Lecture automatique des documents impossible : ' + m, 'docs').catch(() => {})
  })
  const rt = (globalThis as { EdgeRuntime?: { waitUntil(p: Promise<unknown>): void } }).EdgeRuntime
  if (rt?.waitUntil) {
    rt.waitUntil(travail)
    return reponse({ lance: true, message: 'Lecture des documents en cours : la mission se remplit toute seule' }, 202)
  }
  await travail
  return reponse({ message: 'Lecture des documents terminée' })
})
