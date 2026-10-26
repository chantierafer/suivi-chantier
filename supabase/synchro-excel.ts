// Fonction Supabase « synchro-excel » (26/10/2026) : relie les deux Excel du dossier SharePoint « Gestion AFER »
// (personnel et absences ; véhicules, entretiens, carburant et matériel) à la table « gestion » de l'appli.
// Les Excel restent la référence : la fonction en garde une copie dans la table (lignes personne, absence, vehicule,
// materiel, entretien, carburant) et y écrit les modifications faites dans l'appli (file « excel_ecrire »).
// Contrat complet : gestion/CONTRAT-SYNCHRO-EXCEL.md. Mise en place : supabase/MISE-EN-PLACE-V12.md.
// Lancée toutes les 3 minutes et à chaque modification dans l'appli (complement-v12-gestion.sql), clé d'appel
// « cle_appel » de la table synchro_excel dans l'en-tête x-synchro-cle.
// Accès à SharePoint : même application Microsoft Entra que la copie SharePoint (Sites.ReadWrite.All),
// secrets MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET. L'API Excel de Microsoft Graph n'accepte pas ce type
// d'accès : la fonction lit et modifie donc le fichier .xlsx elle-même (format Office Open XML), puis le renvoie
// en vérifiant qu'il n'a pas changé entre-temps (If-Match). Aucune donnée personnelle n'est écrite dans les journaux.
import { createClient } from 'npm:@supabase/supabase-js@2.117.2'
import { unzipSync, zipSync } from 'npm:fflate@0.8.2'

// deno-lint-ignore no-explicit-any
type Any = any
const G = globalThis as Any
const env = (k: string): string => (G.Deno ? G.Deno.env.get(k) : G.process?.env?.[k]) || ''

const DRIVE = env('SP_DRIVE_ID') || 'b!BDWYjvp0M0WVzP8uET8TlMvgPxAZXl1PsD1PQ9hb_kLzQHFSNZ0LSJj4RiSepD3f' // data - Documents
const DOSSIER = env('SP_DOSSIER_GESTION') || 'Gestion AFER'
const LOGIN = env('MS_LOGIN_URL') || 'https://login.microsoftonline.com'
const GRAPH = env('MS_GRAPH_URL') || 'https://graph.microsoft.com/v1.0'
const DELAI_MS = 110_000 // Supabase arrête la fonction à 150 s
const ATTENTE_VERROU_MS = Number(env('SYNCHRO_ATTENTE_MS') || 45_000) // un passage déjà en cours : on attend son tour
const ESSAIS_MAX = 5
const XLSX = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'

// Fichiers et tableaux reliés (les ID d'élément SharePoint ne changent pas si le fichier est renommé ou déplacé ;
// s'il a été recréé, la fonction le retrouve par son nom dans le dossier Gestion AFER)
export type Fichier = { cle: string; itemId: string; nom: string; tables: Record<string, string> }
export const FICHIERS: Fichier[] = [
  { cle: 'personnel', itemId: '01TF36A6C74PYKZQIOMJCZI2UMZK4FTG5K', nom: 'AFER - Personnel et absences.xlsx', tables: { personne: 'T_Personnel', absence: 'T_Absences' } },
  { cle: 'vehicules', itemId: '01TF36A6EKHQA2OQUP7ZBYK7TV3XTEKTWD', nom: 'AFER - Véhicules et matériel.xlsx', tables: { vehicule: 'T_Vehicules', materiel: 'T_Materiel', entretien: 'T_Entretiens', carburant: 'T_Carburant' } },
]
export const ECRIVABLES = new Set(['personne', 'absence', 'vehicule', 'materiel'])
const FORMAT_ID: Record<string, [string, number]> = { personne: ['P-', 3], absence: ['A-', 3], vehicule: ['V-', 2], materiel: ['M-', 2] }
// colonnes numériques encore vides dans les fichiers (le type ne se devine pas sans valeur)
const NOMBRES = new Set(['Coefficient', 'Jours travaillés par mois', 'Kilométrage (dernier relevé)', 'Dernière vidange : km', 'Prochaine vidange : km', 'Km', 'Litres', 'Km au compteur'])

// =====================================================================================================
// Lecture et modification d'un classeur .xlsx (sans dépendance autre que la décompression)
// =====================================================================================================
const dec = new TextDecoder(), enc = new TextEncoder()
const ENT: Record<string, string> = { amp: '&', lt: '<', gt: '>', quot: '"', apos: "'" }
export const deXml = (s: string) => s.replace(/&(#x[0-9a-fA-F]+|#\d+|amp|lt|gt|quot|apos);/g, (_m, e: string) =>
  e[0] === '#' ? String.fromCodePoint(e[1] === 'x' || e[1] === 'X' ? parseInt(e.slice(2), 16) : parseInt(e.slice(1), 10)) : ENT[e])
const deX = (s: string) => s.replace(/_x([0-9A-Fa-f]{4})_/g, (_m, h: string) => String.fromCharCode(parseInt(h, 16))) // échappements Excel (_x000D_)
export const versXml = (s: string) => s.replace(/[\x00-\x08\x0B\x0C\x0E-\x1F]/g, '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
const versAttr = (s: string) => versXml(s).replace(/"/g, '&quot;')
function lireAttrs(s: string) {
  const o: Record<string, string> = {}
  for (const m of s.matchAll(/([\w:.-]+)\s*=\s*(?:"([^"]*)"|'([^']*)')/g)) o[m[1]] = deXml(m[2] ?? m[3] ?? '')
  return o
}
const attr = (a: Record<string, string>, local: string) => { for (const k in a) if (k === local || k.endsWith(':' + local)) return a[k]; return undefined }
function texteRiche(x: string) {
  return deX(deXml([...x.replace(/<rPh\b[\s\S]*?<\/rPh>/g, '').matchAll(/<t(?:\s[^>]*)?>([\s\S]*?)<\/t>/g)].map((m) => m[1]).join('')))
}
export const colNum = (l: string) => [...l.toUpperCase()].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0)
export const colLettre = (n: number) => { let s = ''; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26) } return s }
export function plage(ref: string) {
  const m = /^\$?([A-Za-z]{1,3})\$?(\d+)(?::\$?([A-Za-z]{1,3})\$?(\d+))?$/.exec(ref.trim())
  if (!m) return null
  const c1 = colNum(m[1]), r1 = +m[2], c2 = m[3] ? colNum(m[3]) : c1, r2 = m[4] ? +m[4] : r1
  return { c1: Math.min(c1, c2), r1: Math.min(r1, r2), c2: Math.max(c1, c2), r2: Math.max(r1, r2) }
}
const refDe = (c: number, r: number) => colLettre(c) + r
// dates Excel (calendrier 1900, ou 1904 si le classeur le demande)
export function serieVersIso(n: number, d1904 = false) {
  const ms = Date.UTC(d1904 ? 1904 : 1899, d1904 ? 0 : 11, d1904 ? 1 : 30) + Math.floor(n + 1e-9) * 864e5
  return new Date(ms).toISOString().slice(0, 10)
}
export function isoVersSerie(s: string, d1904 = false) {
  const [y, m, d] = s.split('-').map(Number)
  return Math.round((Date.UTC(y, m - 1, d) - Date.UTC(d1904 ? 1904 : 1899, d1904 ? 0 : 11, d1904 ? 1 : 30)) / 864e5)
}
function formatDate(id: number, code?: string) {
  if (code == null) return (id >= 14 && id <= 17) || id === 22 || (id >= 27 && id <= 36) || (id >= 50 && id <= 58)
  const c = code.replace(/"[^"]*"/g, '').replace(/\[[^\]]*\]/g, '').replace(/\\./g, '')
  return /[dy]/i.test(c)
}
// décalage des références relatives d'une formule (recopie d'une ligne vers la suivante)
export function decaler(f: string, dr: number) {
  let out = '', i = 0
  while (i < f.length) {
    const ch = f[i]
    if (ch === '"' || ch === "'") { let j = i + 1; while (j < f.length) { if (f[j] === ch) { if (f[j + 1] === ch) { j += 2; continue } break } j++ } out += f.slice(i, j + 1); i = j + 1; continue }
    if (ch === '[') { const j = f.indexOf(']', i); const k = j < 0 ? f.length - 1 : j; out += f.slice(i, k + 1); i = k + 1; continue }
    const m = /^(\$?)([A-Z]{1,3})(\$?)(\d+)(?![\w(.])/.exec(f.slice(i))
    if (m && (i === 0 || !/[\w.]/.test(f[i - 1]))) { out += m[1] + m[2] + m[3] + (m[3] ? m[4] : String(+m[4] + dr)); i += m[0].length; continue }
    out += ch; i++
  }
  return out
}

type Rel = { type: string; target: string; mode?: string }
const dossierDe = (p: string) => p.includes('/') ? p.slice(0, p.lastIndexOf('/') + 1) : ''
const relsDe = (p: string) => dossierDe(p) + '_rels/' + p.slice(p.lastIndexOf('/') + 1) + '.rels'
function resoudre(base: string, cible: string) {
  if (cible.startsWith('/')) return cible.slice(1)
  const parts = (base + cible).split('/'), out: string[] = []
  for (const x of parts) { if (x === '..') out.pop(); else if (x !== '.' && x !== '') out.push(x) }
  return out.join('/')
}
function lireRels(xml: string | null, base: string) {
  const o: Record<string, Rel> = {}
  if (!xml) return o
  for (const m of xml.matchAll(/<Relationship\b([^>]*?)\/?>/g)) {
    const a = lireAttrs(m[1])
    o[a.Id] = { type: a.Type || '', target: a.TargetMode === 'External' ? a.Target : resoudre(base, a.Target || ''), mode: a.TargetMode }
  }
  return o
}

export type Cellule = { col: number; a: Record<string, string>; inner: string; raw: string; modif: boolean }
export type Ligne = { r: number; a: string; cells: Cellule[]; raw: string; modif: boolean }
export type FeuilleM = { avant: string; apres: string; lignes: Ligne[]; modif: boolean }
export type Tableau = { nom: string; chemin: string; xml: string; feuille: string; cheminFeuille: string; c1: number; r1: number; c2: number; r2: number; colonnes: string[]; totaux: number; modif: boolean }
export type Classeur = {
  z: Record<string, Uint8Array>; ordre: string[]; wbChemin: string; d1904: boolean
  feuilles: { nom: string; chemin: string; rels: Record<string, Rel> }[]
  partages: string[]; dates: Set<number>; euros: Set<number>; noms: Record<string, string>
  tableaux: Record<string, Tableau>; modeles: Record<string, FeuilleM>; change: boolean
}

export function ouvrir(octets: Uint8Array): Classeur {
  const z = unzipSync(octets), ordre = Object.keys(z)
  const txt = (p: string) => z[p] ? dec.decode(z[p]) : null
  const racine = lireRels(txt('_rels/.rels'), '')
  const wbChemin = Object.values(racine).find((r) => /\/officeDocument$/.test(r.type))?.target || 'xl/workbook.xml'
  const wb = txt(wbChemin)
  if (!wb) throw new Error('classeur illisible (workbook.xml absent)')
  const wbRels = lireRels(txt(relsDe(wbChemin)), dossierDe(wbChemin))
  const cl: Classeur = { z, ordre, wbChemin, d1904: /<workbookPr\b[^>]*\bdate1904="(?:1|true)"/.test(wb), feuilles: [], partages: [], dates: new Set(), euros: new Set(), noms: {}, tableaux: {}, modeles: {}, change: false }
  for (const m of wb.matchAll(/<sheet\b([^>]*?)\/?>/g)) {
    const a = lireAttrs(m[1]), rid = attr(a, 'id'), rel = rid ? wbRels[rid] : undefined
    if (rel) cl.feuilles.push({ nom: a.name, chemin: rel.target, rels: lireRels(txt(relsDe(rel.target)), dossierDe(rel.target)) })
  }
  for (const m of wb.matchAll(/<definedName\b([^>]*)>([\s\S]*?)<\/definedName>/g)) { const a = lireAttrs(m[1]); if (a.localSheetId == null) cl.noms[a.name] = deXml(m[2]) }
  const sst = Object.values(wbRels).find((r) => /\/sharedStrings$/.test(r.type))
  const sx = sst ? txt(sst.target) : null
  if (sx) for (const m of sx.matchAll(/<si\b[^>]*>([\s\S]*?)<\/si>|<si\s*\/>/g)) cl.partages.push(m[1] ? texteRiche(m[1]) : '')
  const st = Object.values(wbRels).find((r) => /\/styles$/.test(r.type))
  const sty = st ? txt(st.target) : null
  if (sty) {
    const fmts: Record<number, string> = {}
    for (const m of sty.matchAll(/<numFmt\b([^>]*?)\/?>/g)) { const a = lireAttrs(m[1]); fmts[+a.numFmtId] = a.formatCode || '' }
    const bloc = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(sty)
    let i = 0
    if (bloc) for (const m of bloc[1].matchAll(/<xf\b([^>]*?)(?:\/>|>)/g)) {
      const id = +(lireAttrs(m[1]).numFmtId || 0), code = fmts[id]
      if (formatDate(id, code)) cl.dates.add(i)
      if (code && /€|\[\$EUR|EUR/.test(code)) cl.euros.add(i)
      i++
    }
  }
  for (const f of cl.feuilles) for (const rel of Object.values(f.rels)) {
    if (!/\/table$/.test(rel.type)) continue
    const x = txt(rel.target); if (!x) continue
    const ta = lireAttrs((/<table\b([^>]*)>/.exec(x) || [])[1] || ''), p = plage(ta.ref || '')
    if (!p) continue
    const colonnes = [...x.matchAll(/<tableColumn\b([^>]*?)(?:\/>|>)/g)].map((m) => deX(lireAttrs(m[1]).name || ''))
    const nom = ta.displayName || ta.name
    if (ta.headerRowCount === '0') continue // tableau sans ligne de titres : pas géré
    cl.tableaux[nom] = { nom, chemin: rel.target, xml: x, feuille: f.nom, cheminFeuille: f.chemin, ...p, colonnes, totaux: +(ta.totalsRowCount || 0), modif: false }
  }
  return cl
}

export function modele(cl: Classeur, chemin: string): FeuilleM {
  if (cl.modeles[chemin]) return cl.modeles[chemin]
  const xml = dec.decode(cl.z[chemin])
  const m = /<sheetData\s*\/>|<sheetData\b[^>]*>([\s\S]*?)<\/sheetData>/.exec(xml)
  if (!m) throw new Error('feuille illisible (' + chemin + ')')
  const f: FeuilleM = { avant: xml.slice(0, m.index) + '<sheetData>', apres: '</sheetData>' + xml.slice(m.index + m[0].length), lignes: [], modif: false }
  let rAuto = 0
  for (const rm of (m[1] || '').matchAll(/<row\b([^>]*?)(?:\/>|>([\s\S]*?)<\/row>)/g)) {
    const ra = lireAttrs(rm[1]), r = ra.r ? +ra.r : rAuto + 1
    rAuto = r
    const cells: Cellule[] = []
    let cAuto = 0
    for (const cm of (rm[2] || '').matchAll(/<c\b([^>]*?)(?:\/>|>([\s\S]*?)<\/c>)/g)) {
      const a = lireAttrs(cm[1]), p = a.r ? plage(a.r) : null, col = p ? p.c1 : cAuto + 1
      cAuto = col
      cells.push({ col, a, inner: cm[2] || '', raw: cm[0], modif: false })
    }
    f.lignes.push({ r, a: rm[1], cells, raw: rm[0], modif: false })
  }
  cl.modeles[chemin] = f
  return f
}

type Val = { v: unknown; f: boolean }
function valeur(cl: Classeur, c: Cellule | undefined): Val {
  if (!c) return { v: undefined, f: false }
  const f = /<f\b/.test(c.inner), t = c.a.t
  const vm = /<v>([\s\S]*?)<\/v>/.exec(c.inner), v = vm ? deXml(vm[1]) : null
  if (t === 's') return { v: v == null || v === '' ? undefined : cl.partages[+v], f }
  if (t === 'inlineStr') { const is = /<is>([\s\S]*?)<\/is>/.exec(c.inner); return { v: is ? texteRiche(is[1]) : undefined, f } }
  if (t === 'str') return { v: v == null ? undefined : deX(v), f }
  if (t === 'b') return { v: v == null ? undefined : v === '1' || v === 'true', f }
  if (t === 'e') return { v: undefined, f }
  if (v == null || v === '') return { v: undefined, f }
  const n = Number(v)
  if (!isFinite(n)) return { v, f }
  if (cl.dates.has(+(c.a.s || 0))) return { v: serieVersIso(n, cl.d1904), f }
  return { v: n, f }
}
const vide = (v: unknown) => v === undefined || v === null || (typeof v === 'string' && v.trim() === '')
const cellDe = (l: Ligne | undefined, col: number) => l?.cells.find((c) => c.col === col)
const ligneDe = (f: FeuilleM, r: number) => f.lignes.find((l) => l.r === r)

// valeurs d'une plage (listes déroulantes) : « Listes!$A$2:$A$4 », « 'Listes'!A2:A4 » ou nom défini
function valeursPlage(cl: Classeur, ref: string, prof = 0): string[] {
  let s = ref.trim().replace(/^=/, '')
  if (/^"[\s\S]*"$/.test(s)) return s.slice(1, -1).split(',').map((x) => x.trim()).filter(Boolean)
  if (cl.noms[s] != null && prof < 3) return valeursPlage(cl, cl.noms[s], prof + 1)
  const m = /^(?:'((?:[^']|'')+)'|([^!]+))!(.+)$/.exec(s)
  if (!m) return []
  const nomF = (m[1] || m[2]).replace(/''/g, "'"), p = plage(m[3]), fe = cl.feuilles.find((x) => x.nom === nomF)
  if (!p || !fe) return []
  const f = modele(cl, fe.chemin), out: string[] = []
  for (const l of f.lignes) {
    if (l.r < p.r1 || l.r > p.r2) continue
    for (let c = p.c1; c <= p.c2; c++) { const v = valeur(cl, cellDe(l, c)).v; if (!vide(v)) out.push(String(v).trim()) }
  }
  return [...new Set(out)]
}
function validations(cl: Classeur, chemin: string) {
  const xml = dec.decode(cl.z[chemin]), out: { sqref: string; formule: string }[] = []
  for (const m of xml.matchAll(/<dataValidation\b([^>]*)>([\s\S]*?)<\/dataValidation>/g)) {
    const a = lireAttrs(m[1]); if (a.type !== 'list') continue
    const fo = /<formula1>([\s\S]*?)<\/formula1>/.exec(m[2]); if (fo) out.push({ sqref: a.sqref || '', formule: deXml(fo[1]) })
  }
  for (const m of xml.matchAll(/<x14:dataValidation\b([^>]*)>([\s\S]*?)<\/x14:dataValidation>/g)) {
    const a = lireAttrs(m[1]); if (a.type !== 'list') continue
    const fo = /<x14:formula1>\s*<xm:f>([\s\S]*?)<\/xm:f>/.exec(m[2]), sq = /<xm:sqref>([\s\S]*?)<\/xm:sqref>/.exec(m[2])
    if (fo && sq) out.push({ sqref: sq[1], formule: deXml(fo[1]) })
  }
  return out
}
function liens(cl: Classeur, fe: { chemin: string; rels: Record<string, Rel> }) {
  const xml = dec.decode(cl.z[fe.chemin]), o: Record<string, string> = {}
  for (const m of xml.matchAll(/<hyperlink\b([^>]*?)\/?>/g)) {
    const a = lireAttrs(m[1]), rid = attr(a, 'id'), url = rid && fe.rels[rid] ? fe.rels[rid].target : a.location ? '#' + a.location : ''
    if (a.ref && url) o[a.ref.split(':')[0]] = url
  }
  return o
}

export type ColMeta = { nom: string; type: 'texte' | 'nombre' | 'euro' | 'date' | 'formule'; liste?: string[] }
export type LigneTab = { r: number; id: string; data: Record<string, unknown>; vide: boolean }
// contenu d'un tableau : colonnes (type, liste) et lignes (vides comprises, marquées)
export function lireTableau(cl: Classeur, nom: string) {
  const tb = cl.tableaux[nom]
  if (!tb) throw new Error(`tableau ${nom} introuvable`)
  const fe = cl.feuilles.find((x) => x.chemin === tb.cheminFeuille)!
  const f = modele(cl, tb.cheminFeuille), dv = validations(cl, tb.cheminFeuille), lk = liens(cl, fe)
  const idCol = tb.colonnes.indexOf('ID'), fin = tb.r2 - tb.totaux
  const colonnes: ColMeta[] = tb.colonnes.map((n, j) => {
    const c = tb.c1 + j, styles: Record<string, number> = {}
    let formule = false, nombre = false, texte = false
    for (const l of f.lignes) {
      if (l.r <= tb.r1 || l.r > fin) continue
      const ce = cellDe(l, c); if (!ce) continue
      if (ce.a.s != null) styles[ce.a.s] = (styles[ce.a.s] || 0) + 1
      const v = valeur(cl, ce)
      if (v.f) formule = true
      else if (typeof v.v === 'number') nombre = true
      else if (!vide(v.v)) texte = true
    }
    const s = Object.entries(styles).sort((a, b) => b[1] - a[1])[0]?.[0]
    const type: ColMeta['type'] = formule ? 'formule' : s != null && cl.dates.has(+s) ? 'date' : s != null && cl.euros.has(+s) ? 'euro' : (nombre && !texte) || NOMBRES.has(n) ? 'nombre' : 'texte'
    const o: ColMeta = { nom: n, type }
    const v = dv.find((d) => d.sqref.split(/\s+/).some((x) => { const p = plage(x); return p && c >= p.c1 && c <= p.c2 && tb.r1 + 1 >= p.r1 && tb.r1 + 1 <= p.r2 }))
    if (v && type !== 'formule') { const l = valeursPlage(cl, v.formule); if (l.length) o.liste = l }
    return o
  })
  const lignes: LigneTab[] = []
  for (let r = tb.r1 + 1; r <= fin; r++) {
    const l = ligneDe(f, r), data: Record<string, unknown> = {}
    let contenu = false
    tb.colonnes.forEach((n, j) => {
      const ce = cellDe(l, tb.c1 + j), v = valeur(cl, ce)
      if (vide(v.v)) return
      data[n] = typeof v.v === 'string' && j === idCol ? v.v.trim() : v.v
      if (!v.f && j !== idCol) contenu = true
      const u = lk[refDe(tb.c1 + j, r)]
      if (u) { const L = (data._liens || {}) as Record<string, string>; L[n] = u; data._liens = L }
    })
    const id = idCol >= 0 ? String(data.ID ?? '').trim() : 'L' + r
    data._ligne = r
    lignes.push({ r, id, data, vide: !contenu })
  }
  return { tb, colonnes, lignes, idCol }
}

function marquer(cl: Classeur, chemin: string, l: Ligne) { l.modif = true; cl.modeles[chemin].modif = true; cl.change = true }
function celluleOuNouvelle(l: Ligne, col: number, s?: string) {
  let c = cellDe(l, col)
  if (c) return c
  c = { col, a: { r: '' }, inner: '', raw: '', modif: true }
  if (s != null) c.a.s = s
  const i = l.cells.findIndex((x) => x.col > col)
  if (i < 0) l.cells.push(c); else l.cells.splice(i, 0, c)
  return c
}
// style d'une colonne pour une cellule créée : celui de la ligne la plus proche qui en a un
function styleColonne(f: FeuilleM, col: number, r: number) {
  let best: string | undefined, d = Infinity
  for (const l of f.lignes) { const c = cellDe(l, col); if (c?.a.s != null && Math.abs(l.r - r) < d) { d = Math.abs(l.r - r); best = c.a.s } }
  return best
}
function poser(cl: Classeur, c: Cellule, v: unknown, type: ColMeta['type'] | undefined) {
  if (/<f\b/.test(c.inner)) return false // jamais une formule
  delete c.a.t
  let inner = ''
  if (vide(v)) inner = ''
  else if (typeof v === 'boolean') { c.a.t = 'b'; inner = `<v>${v ? 1 : 0}</v>` }
  else if (typeof v === 'number' && isFinite(v)) inner = `<v>${v}</v>`
  else {
    const s = String(v), sn = s.replace(/[\s €]/g, '').replace(',', '.')
    if ((type === 'date' || cl.dates.has(+(c.a.s || 0))) && /^\d{4}-\d{2}-\d{2}$/.test(s.trim())) inner = `<v>${isoVersSerie(s.trim(), cl.d1904)}</v>`
    else if ((type === 'nombre' || type === 'euro') && /^-?\d+(\.\d+)?$/.test(sn)) inner = `<v>${Number(sn)}</v>`
    else { c.a.t = 'inlineStr'; inner = `<is><t xml:space="preserve">${versXml(s)}</t></is>` }
  }
  c.inner = inner; c.modif = true
  return true
}
// formule partagée (t="shared") : retrouve le texte pour la cellule donnée
function texteFormule(f: FeuilleM, c: Cellule, r: number): string | null {
  const m = /<f\b([^>]*?)(?:\/>|>([\s\S]*?)<\/f>)/.exec(c.inner)
  if (!m) return null
  const a = lireAttrs(m[1])
  if (a.t !== 'shared' || m[2]) return m[2] ? deXml(m[2]) : null
  for (const l of f.lignes) for (const x of l.cells) {
    const mm = /<f\b([^>]*?)>([\s\S]*?)<\/f>/.exec(x.inner)
    if (mm) { const b = lireAttrs(mm[1]); if (b.t === 'shared' && b.si === a.si) return decaler(deXml(mm[2]), r - l.r) }
  }
  return null
}
// ligne « vide » d'un tableau : aucune valeur hors formules et hors ID (ex. la ligne modèle « A-0001 »)
function estVide(cl: Classeur, tb: Tableau, l: Ligne | undefined, idCol: number) {
  if (!l) return true
  for (let j = 0; j < tb.colonnes.length; j++) {
    if (j === idCol) continue
    const v = valeur(cl, cellDe(l, tb.c1 + j))
    if (!v.f && !vide(v.v)) return false
  }
  return true
}
function inserer(f: FeuilleM, l: Ligne) { const i = f.lignes.findIndex((x) => x.r > l.r); if (i < 0) f.lignes.push(l); else f.lignes.splice(i, 0, l) }
// nouvelle ligne du tableau : la première ligne vide, sinon à la suite (le tableau s'agrandit)
function ajouterLigne(cl: Classeur, tb: Tableau, idCol: number): Ligne {
  const f = modele(cl, tb.cheminFeuille), fin = tb.r2 - tb.totaux
  for (let r = tb.r1 + 1; r <= fin; r++) {
    const l = ligneDe(f, r)
    if (estVide(cl, tb, l, idCol)) {
      if (l) return l
      const n: Ligne = { r, a: ` r="${r}"`, cells: [], raw: '', modif: true }
      inserer(f, n); return n
    }
  }
  if (tb.totaux) throw new Error(`tableau ${tb.nom} avec ligne de total : ajout impossible`)
  const nr = tb.r2 + 1, prec = ligneDe(f, tb.r2)
  let l = ligneDe(f, nr)
  if (l) {
    for (let c = tb.c1; c <= tb.c2; c++) { const v = valeur(cl, cellDe(l, c)); if (v.f || !vide(v.v)) throw new Error(`ligne ${nr} sous le tableau ${tb.nom} déjà occupée`) }
  } else { l = { r: nr, a: (prec?.a || '').replace(/\s(?:r|spans)="[^"]*"/g, '') + ` r="${nr}"`, cells: [], raw: '', modif: true }; inserer(f, l) }
  for (let c = tb.c1; c <= tb.c2; c++) {
    const p = cellDe(prec, c), ce = celluleOuNouvelle(l, c, p?.a.s ?? styleColonne(f, c, nr))
    const tf = p ? texteFormule(f, p, tb.r2) : null
    if (tf) { delete ce.a.t; ce.inner = `<f>${versXml(decaler(tf, 1))}</f>`; ce.modif = true }
  }
  tb.r2 = nr; tb.modif = true
  const ref = `${refDe(tb.c1, tb.r1)}:${refDe(tb.c2, tb.r2)}`
  tb.xml = tb.xml.replace(/(<table\b[^>]*?\sref=")[^"]*(")/, `$1${ref}$2`).replace(/(<autoFilter\b[^>]*?\sref=")[^"]*(")/, `$1${ref}$2`)
  f.avant = f.avant.replace(/(<dimension\b[^>]*?\sref=")([^"]*)(")/, (_m, a: string, d: string, b: string) => {
    const p = plage(d.includes(':') ? d : d + ':' + d)
    return p && p.r2 < nr ? a + `${refDe(p.c1, p.r1)}:${refDe(Math.max(p.c2, tb.c2), nr)}` + b : a + d + b
  })
  return l
}
// les formules d'une ligne modifiée seront recalculées par Excel à l'ouverture : on retire l'ancien résultat
function oublierResultats(l: Ligne) {
  for (const c of l.cells) if (/<f\b/.test(c.inner) && /<v>/.test(c.inner)) { c.inner = c.inner.replace(/<v>[\s\S]*?<\/v>/, ''); delete c.a.t; c.modif = true }
}

export type Op = { cle: string; table: string; id: string; valeurs: Record<string, unknown>; nouveau?: boolean; supprimer?: boolean }
export type Resultat = { cle: string; ok: boolean; erreur?: string; definitif?: boolean }
// applique les écritures de l'appli à un classeur (en mémoire)
export function appliquer(cl: Classeur, fichier: Fichier, ops: Op[]): Resultat[] {
  const out: Resultat[] = []
  for (const op of ops) {
    try {
      const nomT = fichier.tables[op.table]
      if (!nomT || !ECRIVABLES.has(op.table)) { out.push({ cle: op.cle, ok: false, erreur: `table « ${op.table} » en lecture seule ou inconnue`, definitif: true }); continue }
      const { tb, colonnes, lignes, idCol } = lireTableau(cl, nomT)
      if (idCol < 0) throw new Error(`colonne ID absente du tableau ${nomT}`)
      const id = String(op.id || '').trim()
      if (!id) { out.push({ cle: op.cle, ok: false, erreur: 'ID manquant', definitif: true }); continue }
      const f = modele(cl, tb.cheminFeuille), trouve = lignes.find((x) => !x.vide && x.id === id)
      if (op.supprimer) {
        if (trouve) {
          const l = ligneDe(f, trouve.r)!
          for (let j = 0; j < tb.colonnes.length; j++) { const c = cellDe(l, tb.c1 + j); if (c && !/<f\b/.test(c.inner)) poser(cl, c, null, undefined) }
          f.apres = f.apres.replace(new RegExp(`<hyperlink\\b[^>]*\\sref="[A-Z]+${trouve.r}"[^>]*/>`, 'g'), '').replace(/<hyperlinks>\s*<\/hyperlinks>/, '')
          oublierResultats(l); marquer(cl, tb.cheminFeuille, l)
        }
        out.push({ cle: op.cle, ok: true }); continue
      }
      let l: Ligne
      if (trouve) l = ligneDe(f, trouve.r)!
      else if (op.nouveau) {
        l = ajouterLigne(cl, tb, idCol)
        poser(cl, celluleOuNouvelle(l, tb.c1 + idCol, styleColonne(f, tb.c1 + idCol, l.r)), id, 'texte')
      } else { out.push({ cle: op.cle, ok: false, erreur: `fiche ${id} introuvable dans l'Excel` }); continue }
      for (const [k, v] of Object.entries(op.valeurs || {})) {
        const j = tb.colonnes.indexOf(k)
        if (j < 0 || j === idCol || k.startsWith('_') || colonnes[j].type === 'formule') continue
        poser(cl, celluleOuNouvelle(l, tb.c1 + j, styleColonne(f, tb.c1 + j, l.r)), v, colonnes[j].type)
      }
      oublierResultats(l); marquer(cl, tb.cheminFeuille, l)
      out.push({ cle: op.cle, ok: true })
    } catch (e) { out.push({ cle: op.cle, ok: false, erreur: String((e as Error)?.message || e).slice(0, 200) }) }
  }
  return out
}
export function prochainId(type: string, ids: string[]) {
  let [pre, larg] = FORMAT_ID[type] || ['X-', 3], max = 0
  const pres: Record<string, number> = {}
  for (const x of ids) { const m = /^([A-Za-z]+-?)(\d+)$/.exec(x); if (!m) continue; pres[m[1]] = (pres[m[1]] || 0) + 1; max = Math.max(max, +m[2]); larg = Math.max(larg, m[2].length) }
  const p = Object.entries(pres).sort((a, b) => b[1] - a[1])[0]; if (p) pre = p[0]
  return pre + String(max + 1).padStart(larg, '0')
}
// lignes remplies dans l'Excel sans ID : la fonction leur en donne un
export function remplirIds(cl: Classeur, fichier: Fichier) {
  const faits: { table: string; id: string }[] = []
  for (const [type, nomT] of Object.entries(fichier.tables)) {
    if (!ECRIVABLES.has(type) || !cl.tableaux[nomT]) continue
    const { tb, lignes, idCol } = lireTableau(cl, nomT)
    if (idCol < 0) continue
    const ids = lignes.map((x) => x.id).filter(Boolean), f = modele(cl, tb.cheminFeuille)
    for (const x of lignes) {
      if (x.vide || x.id) continue
      const id = prochainId(type, ids); ids.push(id)
      const l = ligneDe(f, x.r)!
      poser(cl, celluleOuNouvelle(l, tb.c1 + idCol, styleColonne(f, tb.c1 + idCol, x.r)), id, 'texte')
      marquer(cl, tb.cheminFeuille, l); faits.push({ table: type, id })
    }
  }
  return faits
}
function serCellule(c: Cellule, r: number) {
  if (!c.modif) return c.raw
  const a: Record<string, string> = { ...c.a, r: refDe(c.col, r) }
  const ordre = ['r', 's', 't', ...Object.keys(a).filter((k) => !['r', 's', 't'].includes(k))]
  const at = ordre.filter((k) => a[k] != null).map((k) => ` ${k}="${versAttr(a[k])}"`).join('')
  return c.inner ? `<c${at}>${c.inner}</c>` : `<c${at}/>`
}
function serLigne(l: Ligne) {
  if (!l.modif) return l.raw
  const a = l.a.replace(/\sspans="[^"]*"/, '')
  const at = /\sr="/.test(a) ? a.replace(/\sr="[^"]*"/, ` r="${l.r}"`) : ` r="${l.r}"` + a
  return `<row${at}>${l.cells.map((c) => serCellule(c, l.r)).join('')}</row>`
}
// classeur modifié → octets (calcChain retiré, recalcul complet à l'ouverture dans Excel)
export function versOctets(cl: Classeur): Uint8Array {
  const z = { ...cl.z }
  for (const [chemin, f] of Object.entries(cl.modeles)) if (f.modif) z[chemin] = enc.encode(f.avant + f.lignes.map(serLigne).join('') + f.apres)
  for (const tb of Object.values(cl.tableaux)) if (tb.modif) z[tb.chemin] = enc.encode(tb.xml)
  const cc = Object.keys(z).find((k) => /(^|\/)calcChain\.xml$/.test(k))
  if (cc) {
    delete z[cc]
    const rp = relsDe(cl.wbChemin)
    if (z[rp]) z[rp] = enc.encode(dec.decode(z[rp]).replace(/<Relationship\b[^>]*calcChain[^>]*\/>/g, ''))
    if (z['[Content_Types].xml']) z['[Content_Types].xml'] = enc.encode(dec.decode(z['[Content_Types].xml']).replace(/<Override\b[^>]*calcChain[^>]*\/>/g, ''))
  }
  let wb = dec.decode(z[cl.wbChemin])
  if (/<calcPr\b/.test(wb)) wb = wb.replace(/<calcPr\b([^>]*?)(\/?)>/, (_m, a: string, s: string) => `<calcPr${a.replace(/\sfullCalcOnLoad="[^"]*"/, '')} fullCalcOnLoad="1"${s}>`)
  else {
    const m = /<(?:\w+:)?(?:oleSize|customWorkbookViews|pivotCaches|smartTagPr|smartTagTypes|webPublishing|fileRecoveryPr|webPublishObjects|extLst)\b|<\/workbook>/.exec(wb)
    if (m) wb = wb.slice(0, m.index) + '<calcPr fullCalcOnLoad="1"/>' + wb.slice(m.index)
  }
  z[cl.wbChemin] = enc.encode(wb)
  const out: Record<string, Uint8Array> = {}
  for (const k of cl.ordre) if (z[k]) out[k] = z[k]
  for (const k of Object.keys(z)) if (!out[k]) out[k] = z[k]
  return zipSync(out, { level: 6 })
}

// =====================================================================================================
// Microsoft Graph (accès d'application)
// =====================================================================================================
let jeton = '', expire = 0
async function graphJeton() {
  if (jeton && Date.now() < expire) return jeton
  const t = env('MS_TENANT_ID'), id = env('MS_CLIENT_ID'), sec = env('MS_CLIENT_SECRET')
  if (!t || !id || !sec) throw new Error('secrets MS_TENANT_ID, MS_CLIENT_ID ou MS_CLIENT_SECRET absents')
  const r = await fetch(`${LOGIN}/${t}/oauth2/v2.0/token`, { method: 'POST', body: new URLSearchParams({ grant_type: 'client_credentials', client_id: id, client_secret: sec, scope: 'https://graph.microsoft.com/.default' }) })
  const j = await r.json().catch(() => ({}))
  if (!r.ok || !j.access_token) throw new Error('connexion Microsoft refusée : ' + (j.error_description || j.error || r.status).toString().slice(0, 200))
  jeton = j.access_token; expire = Date.now() + (Number(j.expires_in || 3600) - 120) * 1000
  return jeton
}
async function graph(url: string, init: RequestInit = {}) {
  for (let i = 0; ; i++) {
    const r = await fetch(url, { ...init, headers: { ...(init.headers || {}), Authorization: 'Bearer ' + await graphJeton() } })
    if ((r.status === 429 || r.status >= 500) && i < 2) { await r.body?.cancel(); await dormir(Number(r.headers.get('retry-after') || 2) * 1000); continue }
    return r
  }
}
type Meta = { id: string; name: string; eTag: string; cTag?: string; webUrl?: string; lastModifiedDateTime?: string }
const SEL = '?$select=id,name,eTag,cTag,webUrl,lastModifiedDateTime,size'
const encChemin = (p: string) => p.split('/').map(encodeURIComponent).join('/')
async function metaFichier(f: Fichier): Promise<Meta> {
  let r = await graph(`${GRAPH}/drives/${DRIVE}/items/${f.itemId}${SEL}`)
  if (r.status === 404) { await r.body?.cancel(); r = await graph(`${GRAPH}/drives/${DRIVE}/root:/${encChemin(DOSSIER + '/' + f.nom)}${SEL}`) }
  if (!r.ok) throw new Error(`SharePoint ${r.status} sur ${f.nom} : ${(await r.text()).slice(0, 150)}`)
  return await r.json()
}
async function telecharger(m: Meta) {
  const r = await graph(`${GRAPH}/drives/${DRIVE}/items/${m.id}/content`)
  if (!r.ok) throw new Error(`SharePoint ${r.status} au téléchargement de ${m.name}`)
  return new Uint8Array(await r.arrayBuffer())
}
async function envoyer(m: Meta, octets: Uint8Array) {
  const r = await graph(`${GRAPH}/drives/${DRIVE}/items/${m.id}/content`, { method: 'PUT', headers: { 'Content-Type': XLSX, 'If-Match': m.eTag }, body: octets as unknown as BodyInit })
  const corps = await r.text()
  let meta: Meta | null = null
  try { meta = r.ok ? JSON.parse(corps) : null } catch { /* réponse sans détail */ }
  return { status: r.status, ok: r.ok, meta, message: corps.slice(0, 150) }
}
const dormir = (ms: number) => new Promise((s) => setTimeout(s, ms))

// =====================================================================================================
// Un passage : écritures en attente, puis lecture des Excel et mise à jour de la table « gestion »
// =====================================================================================================
const stable = (o: unknown) => JSON.stringify(o, (_k, v) => v && typeof v === 'object' && !Array.isArray(v) ? Object.fromEntries(Object.keys(v).sort().map((k) => [k, v[k]])) : v)
const signature = (d: Any) => stable({ v: d?.valeurs || {}, n: !!d?.nouveau, s: !!d?.supprimer })
type FileLigne = { type: string; id: string; data: Any; maj_le: string }
async function tout(q: () => Any): Promise<Any[]> {
  const out: Any[] = []
  for (let i = 0; ; i += 1000) { const { data, error } = await q().range(i, i + 999); if (error) throw new Error(error.message); out.push(...data); if (data.length < 1000) return out }
}
const lireFile = (admin: Any): Promise<FileLigne[]> => tout(() => admin.from('gestion').select('type,id,data,maj_le').eq('type', 'excel_ecrire').order('id'))
const tableDe = (q: FileLigne) => String(q.data?.table || q.id.split(':')[0])
const idDe = (q: FileLigne) => String(q.data?.id ?? q.id.slice(q.id.indexOf(':') + 1))
async function noterErreur(admin: Any, q: FileLigne, erreur: string, compte: boolean, definitif = false) {
  const essais = definitif ? ESSAIS_MAX : (q.data?.essais || 0) + (compte ? 1 : 0)
  if (q.data?.erreur === erreur && (q.data?.essais || 0) === essais) return
  await admin.from('gestion').update({ data: { ...q.data, erreur, essais } }).eq('type', 'excel_ecrire').eq('id', q.id).eq('maj_le', q.maj_le)
}

type Etat = { at: string; ok: boolean; message?: string; fichiers: Record<string, Any>; ecrits: number; erreurs: { id: string; message: string }[] }
export async function passage(admin: Any, o: { force?: boolean; tentes?: Map<string, string> } = {}) {
  const tentes = o.tentes || new Map<string, string>()
  const at = new Date().toISOString()
  const prec = ((await admin.from('gestion').select('data').eq('type', 'excel_etat').eq('id', 'dernier').maybeSingle()).data?.data || {}) as Partial<Etat>
  const metas = new Set((await tout(() => admin.from('gestion').select('id').eq('type', 'excel_meta'))).map((r: Any) => r.id))
  const etat: Etat = { at, ok: true, fichiers: { ...(prec.fichiers || {}) }, ecrits: 0, erreurs: [] }
  let file = await lireFile(admin)
  // écritures impossibles (table inconnue ou en lecture seule)
  for (const q of file) if ((q.data?.essais || 0) < ESSAIS_MAX && !FICHIERS.some((f) => f.tables[tableDe(q)] && ECRIVABLES.has(tableDe(q)))) {
    await noterErreur(admin, q, `table « ${tableDe(q)} » en lecture seule ou inconnue`, true, true); tentes.set(q.id, signature(q.data))
  }
  for (const f of FICHIERS) {
    const aFaire = file.filter((q) => (q.data?.essais || 0) < ESSAIS_MAX && f.tables[tableDe(q)] && ECRIVABLES.has(tableDe(q)))
    const ops: Op[] = aFaire.map((q) => ({ cle: q.id, table: tableDe(q), id: idDe(q), valeurs: q.data?.valeurs || {}, nouveau: !!q.data?.nouveau, supprimer: !!q.data?.supprimer }))
    try {
      let cl: Classeur | null = null, meta: Meta | null = null, res: Resultat[] = [], provisoire = '', idsRemplis = 0
      for (let essai = 0; essai < 3; essai++) {
        meta = await metaFichier(f)
        const signe = meta.cTag || meta.eTag, connu = etat.fichiers[f.cle]
        if (!ops.length && !o.force && connu?.ctag === signe && Object.keys(f.tables).every((t) => metas.has(t))) { cl = null; break }
        const octets = await telecharger(meta)
        cl = ouvrir(octets)
        res = ops.length ? appliquer(cl, f, ops) : []
        idsRemplis = remplirIds(cl, f).length
        if (!cl.change) break
        const dep = await envoyer(meta, versOctets(cl))
        if (dep.ok && dep.meta) { meta = dep.meta; etat.ecrits += res.filter((x) => x.ok).length; break }
        cl = ouvrir(octets) // l'écriture n'est pas passée : la copie reflète le fichier tel qu'il est
        if (dep.status === 412) { provisoire = 'fichier modifié pendant l\'écriture, nouvel essai au prochain passage'; if (essai < 2) { await dormir(500); continue } }
        else if (dep.status === 423) provisoire = 'fichier ouvert et verrouillé dans Excel, nouvel essai au prochain passage'
        else throw new Error(`SharePoint ${dep.status} à l'enregistrement de ${f.nom} : ${dep.message}`)
        res = res.map((x) => x.ok ? { cle: x.cle, ok: false, erreur: provisoire, provisoire: true } as Resultat & { provisoire: boolean } : x)
        break
      }
      // file d'attente : fait → retiré (s'il n'a pas changé entre-temps), sinon erreur notée
      for (const x of res) {
        const q = aFaire.find((y) => y.id === x.cle)!
        tentes.set(q.id, signature(q.data))
        if (x.ok) await admin.from('gestion').delete().eq('type', 'excel_ecrire').eq('id', q.id).eq('maj_le', q.maj_le)
        else await noterErreur(admin, q, x.erreur || 'erreur', !(x as Any).provisoire && !x.definitif, !!x.definitif)
      }
      if (cl && meta) {
        etat.fichiers[f.cle] = { nom: meta.name, webUrl: meta.webUrl, etag: meta.eTag, ctag: meta.cTag || meta.eTag, modifie_le: meta.lastModifiedDateTime, lu_le: new Date().toISOString(), ...(idsRemplis ? { ids_remplis: idsRemplis } : {}) }
        file = await lireFile(admin)
        await miroir(admin, f, cl, new Set(file.map((q) => `${tableDe(q)}:${idDe(q)}`)), etat)
      } else if (meta) etat.fichiers[f.cle] = { ...(etat.fichiers[f.cle] || {}), nom: meta.name, webUrl: meta.webUrl, etag: meta.eTag, modifie_le: meta.lastModifiedDateTime }
    } catch (e) {
      const m = String((e as Error)?.message || e).slice(0, 250)
      etat.erreurs.push({ id: f.nom, message: m })
      for (const q of aFaire) { tentes.set(q.id, signature(q.data)); await noterErreur(admin, q, m, true) }
    }
  }
  // erreurs restantes de la file (visibles dans l'appli)
  for (const q of await lireFile(admin)) if (q.data?.erreur) etat.erreurs.push({ id: q.id, message: String(q.data.erreur) })
  etat.ok = !etat.erreurs.length
  if (!etat.ok) etat.message = `${etat.erreurs.length} erreur(s)`
  await admin.from('gestion').upsert({ type: 'excel_etat', id: 'dernier', data: etat })
  return etat
}

// copie d'un fichier dans la table : fiches modifiées seulement, fiches disparues retirées, sauf écriture en attente
async function miroir(admin: Any, f: Fichier, cl: Classeur, pendants: Set<string>, etat: Etat) {
  const types = Object.keys(f.tables), lus: string[] = [], voulu = new Map<string, Any>(), metas: Any[] = []
  for (const type of types) {
    let t
    try { t = lireTableau(cl, f.tables[type]) } catch (e) { etat.erreurs.push({ id: f.nom, message: String((e as Error).message) }); continue }
    lus.push(type)
    const fe = cl.feuilles.find((x) => x.chemin === t.tb.cheminFeuille)
    metas.push({ type: 'excel_meta', id: type, data: { fichier: f.nom, itemId: f.itemId, feuille: fe?.nom, tableau: t.tb.nom, colonnes: t.colonnes } })
    for (const l of t.lignes) {
      if (l.vide || !l.id) continue
      const k = `${type}:${l.id}`
      if (voulu.has(k)) { etat.erreurs.push({ id: k, message: `ID ${l.id} en double dans l'Excel (ligne ${l.r})` }); continue }
      voulu.set(k, { type, id: l.id, data: l.data })
    }
  }
  const avant = await tout(() => admin.from('gestion').select('type,id,data').in('type', [...types, 'excel_meta']).order('type').order('id'))
  const ex = new Map<string, Any>(avant.map((r: Any) => [`${r.type}:${r.id}`, r]))
  const maintenant = new Date().toISOString()
  const ecrire: Any[] = [], retirer: Any[] = []
  for (const [k, r] of voulu) { if (pendants.has(k)) continue; const e = ex.get(k); if (!e || stable(e.data) !== stable(r.data)) ecrire.push(r) }
  for (const m of metas) {
    const e = ex.get(`excel_meta:${m.id}`), { lu_le: _x, webUrl: _w, ...ancien } = e?.data || {}
    const webUrl = etat.fichiers[f.cle]?.webUrl
    if (!e || stable(ancien) !== stable(m.data) || (webUrl && e.data.webUrl !== webUrl)) ecrire.push({ ...m, data: { ...m.data, webUrl, lu_le: maintenant } })
  }
  // un tableau illisible (renommé ?) ne vide pas sa copie : seules les fiches des tableaux lus sont retirées
  for (const r of avant) { const k = `${r.type}:${r.id}`; if (lus.includes(r.type) && !voulu.has(k) && !pendants.has(k)) retirer.push(r) }
  for (let i = 0; i < ecrire.length; i += 200) { const { error } = await admin.from('gestion').upsert(ecrire.slice(i, i + 200)); if (error) throw new Error(error.message) }
  for (const type of types) {
    const ids = retirer.filter((r) => r.type === type).map((r) => r.id)
    for (let i = 0; i < ids.length; i += 100) { const { error } = await admin.from('gestion').delete().eq('type', type).in('id', ids.slice(i, i + 100)); if (error) throw new Error(error.message) }
  }
}

// verrou : un seul passage à la fois ; un appel qui arrive pendant un passage attend son tour
async function prendreVerrou(admin: Any, moi: string, limite: number) {
  for (;;) {
    const { error } = await admin.from('synchro_excel').insert({ cle: 'verrou', valeur: moi })
    if (!error) return true
    if (error.code !== '23505' && !/duplicate|doublon/i.test(error.message || '')) throw new Error(error.message)
    const { data: v } = await admin.from('synchro_excel').select('valeur').eq('cle', 'verrou').maybeSingle()
    if (v && Date.now() - Date.parse(String(v.valeur).split('|')[0]) > 150_000) { await admin.from('synchro_excel').delete().eq('cle', 'verrou').eq('valeur', v.valeur); continue }
    if (Date.now() > limite) return false
    await dormir(1500)
  }
}
export async function executer(admin: Any, o: { force?: boolean; fin?: number } = {}) {
  const fin = o.fin || Date.now() + DELAI_MS
  const moi = new Date().toISOString() + '|' + crypto.randomUUID()
  if (!await prendreVerrou(admin, moi, Math.min(Date.now() + ATTENTE_VERROU_MS, fin - 40_000))) return { message: 'Passage déjà en cours' }
  try {
    const tentes = new Map<string, string>()
    let force = !!o.force, n = 0, etat: Etat
    const demandeLe = async () => (await admin.from('gestion').select('maj_le').eq('type', 'excel_demande').eq('id', 'maintenant').maybeSingle()).data?.maj_le || ''
    for (;;) {
      const d0 = await demandeLe()
      etat = await passage(admin, { force, tentes }); n++
      // arrivé pendant le passage : nouvelle écriture dans l'appli, ou « Synchroniser maintenant »
      const nouveau = (await lireFile(admin)).some((q) => (q.data?.essais || 0) < ESSAIS_MAX && tentes.get(q.id) !== signature(q.data))
      const demande = (await demandeLe()) !== d0
      if ((!nouveau && !demande) || n >= 4 || Date.now() > fin - 25_000) break
      force = demande
    }
    return { passages: n, ok: etat.ok, ecrits: etat.ecrits, erreurs: etat.erreurs.length }
  } finally {
    await admin.from('synchro_excel').delete().eq('cle', 'verrou').eq('valeur', moi)
  }
}

const reponse = (corps: unknown, status = 200) => new Response(JSON.stringify(corps), { status, headers: { 'Content-Type': 'application/json' } })
function cle(...noms: string[]) {
  for (const n of noms) { const v = env(n); if (!v) continue; if (v.trim().startsWith('{')) { try { const k = JSON.parse(v); if (k.default) return String(k.default) } catch { /* format inattendu */ } } else return v }
  return ''
}
export async function handler(req: Request): Promise<Response> {
  const admin = createClient(env('SUPABASE_URL'), cle('SUPABASE_SECRET_KEYS', 'SUPABASE_SECRET_KEY', 'SUPABASE_SERVICE_ROLE_KEY'), { auth: { persistSession: false, autoRefreshToken: false } })
  const { data: k } = await admin.from('synchro_excel').select('valeur').eq('cle', 'cle_appel').maybeSingle()
  if (!k?.valeur || req.headers.get('x-synchro-cle') !== k.valeur) return reponse({ message: 'Accès refusé' }, 401)
  const corps = await req.json().catch(() => ({}))
  const travail = executer(admin, { force: corps?.type === 'excel_demande' }).catch(async (e) => {
    const message = String((e as Error)?.message || e).slice(0, 300)
    console.error('synchro-excel :', message)
    await admin.from('gestion').upsert({ type: 'excel_etat', id: 'dernier', data: { at: new Date().toISOString(), ok: false, message, fichiers: {}, ecrits: 0, erreurs: [{ id: 'synchro-excel', message }] } })
    return { ok: false, message }
  })
  const rt = G.EdgeRuntime as { waitUntil?(p: Promise<unknown>): void } | undefined
  if (rt?.waitUntil) { rt.waitUntil(travail); return reponse({ lance: true }, 202) }
  return reponse(await travail)
}
if (G.Deno?.serve) G.Deno.serve(handler)
