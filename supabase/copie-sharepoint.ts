// Fonction Supabase « copie-sharepoint » (18/10/2026) : copie le suivi de chantier dans SharePoint, sans PC allumé.
// Lancée toutes les 3 minutes par Supabase (pg_cron, voir ../../complement-v10-copie-sharepoint.sql), elle reprend
// l'organisation du programme PC (synchro-sharepoint/synchro.ps1, abandonné) : un dossier par chantier dans
// data - Documents > Suivi Chantier SDB, rien n'est jamais supprimé dans SharePoint.
// Accès à SharePoint : application inscrite dans Microsoft Entra (Microsoft Graph, permission d'application
// Sites.ReadWrite.All). Secrets de la fonction : MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET.
// L'appel est réservé à la tâche programmée : elle envoie la clé « cle_appel » de la table copie_sharepoint.
import { createClient } from 'npm:@supabase/supabase-js@2.117.2'

const DRIVE = Deno.env.get('SP_DRIVE_ID') || 'b!BDWYjvp0M0WVzP8uET8TlMvgPxAZXl1PsD1PQ9hb_kLzQHFSNZ0LSJj4RiSepD3f' // data - Documents
const RACINE = Deno.env.get('SP_DOSSIER') || 'Suivi Chantier SDB'
const RACINE_GESTION = Deno.env.get('SP_DOSSIER_GESTION') || 'Gestion AFER' // PDF des devis, factures et avoirs
const LOGIN = Deno.env.get('MS_LOGIN_URL') || 'https://login.microsoftonline.com'
const GRAPH = Deno.env.get('MS_GRAPH_URL') || 'https://graph.microsoft.com/v1.0'
const DELAI_MS = 110_000 // Supabase arrête la fonction à 150 s ; le reste part au passage suivant

// deno-lint-ignore no-explicit-any
type Any = any
type El = { chantier_id: string; type: string; id: string; data: Any }

// ---------- noms (mêmes règles que le programme PC) ----------
export function nom(s: unknown, max = 70) {
  let t = String(s ?? '').replace(/[<>:"/\\|?*\x00-\x1F#%]/g, ' ').replace(/\s+/g, ' ').trim().replace(/\.+$/, '')
  if (t.length > max) t = t.slice(0, max).trim()
  return t || 'sans nom'
}
const paris = new Intl.DateTimeFormat('fr-FR', { timeZone: 'Europe/Paris', year: 'numeric', month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit', hourCycle: 'h23' })
export function quand(v: unknown) {
  const d = v ? new Date(String(v)) : null
  if (!d || isNaN(+d)) return 'sans date'
  const p = Object.fromEntries(paris.formatToParts(d).map((x) => [x.type, x.value]))
  return `${p.year}-${p.month}-${p.day} ${p.hour}h${p.minute}`
}
const jour = (v: unknown) => quand(v).slice(0, 10)
const EXT: Record<string, string> = { 'image/jpeg': '.jpg', 'image/png': '.png', 'image/webp': '.webp', 'image/heic': '.heic', 'application/pdf': '.pdf' }
const ETAPES: [string, string][] = [['constat', 'Constat avant travaux'], ['protection', 'Protection et installation'], ['depose', 'Dépose et démolition'],
  ['plomberie', 'Plomberie'], ['receveur', 'Receveur extra-plat'], ['etancheite', 'Étanchéité'], ['dalles', 'Revêtement mural'],
  ['colonne', 'Robinetterie et colonne'], ['paroi', 'Paroi de douche'], ['pmr', 'Équipements PMR'], ['finitions', 'Finitions et nettoyage'], ['reception', 'Réception']]
const RUBRIQUE: Record<string, string> = { pb: 'Problèmes signalés', req: 'Travaux en plus/Demandes du technicien', bl: 'Bons de livraison', frais: 'Factures et frais', sav: 'SAV' }

// chemin dans le dossier du chantier, sans extension
export function destination(el: El, devis: Record<string, Any>) {
  const d = el.data || {}
  if (el.type === 'photo') {
    const slot = String(d.slot || ''), q = quand(d.at)
    if (slot === 'materiel') return `1 Réception du matériel/${q}`
    const i = ETAPES.findIndex(([k]) => k === slot)
    if (i >= 0) return `2 Étapes/${String(i + 1).padStart(2, '0')} ${ETAPES[i][1]}/${q}`
    const m = slot.match(/^x:(.+):(av|ap)$/)
    if (m) return `3 Travaux en plus/${nom(devis[m[1]]?.designation || 'Ligne ' + m[1])}/${m[2] === 'av' ? 'avant' : 'après'} ${q}`
    return `2 Étapes/Autres/${nom(slot, 30)} ${q}`
  }
  if (el.type === 'doc') {
    const base = String(d.name || '').replace(/\.[^.]*$/, '')
    return `4 Documents/${d.cat === 'autre' ? 'Autres documents' : 'Documents de début de chantier'}/${jour(d.at)} ${nom(base)}`
  }
  const lib = d.type || d.lib || d.cat || d.motif
  return `5 ${RUBRIQUE[el.type]}/${quand(d.at)}${lib ? ' ' + nom(lib, 50) : ''}`
}
const cleFichier = (el: El) => el.type === 'photo' ? el.id : el.type === 'doc' ? 'doc:' + el.id : el.data?.photo ? String(el.data.photo) : ''
const stockage = (ch: string, t: string, k: string) => `${ch}/${t}/${k.replace(/[^A-Za-z0-9._-]/g, '_')}`

export function texteFiche(c: Any) {
  const f = c.fiche || {}, L: string[] = ['Chantier ' + (f.ville || ''), 'Copie automatique depuis l\'appli de suivi de chantier. Ne pas modifier ici.', '']
  for (const [t, k] of [['Adresse', 'adresse'], ['Étage', 'etage'], ['Accès', 'acces'], ['Client', 'client'], ['Téléphone client', 'tel'], ['Disponibilité', 'dispo'], ['Conduite de travaux DOMetVIE', 'conducteur']]) L.push(`${t} : ${f[k] ?? ''}`)
  L.push('Variante : ' + (f.variante === 'sans' ? 'sans démolition' : 'avec démolition'))
  ;(f.retraits || []).filter(Boolean).forEach((r: Any, i: number) => { L.push('', `Retrait du matériel ${i + 1} : ${r.lieu || ''}  ${r.adresse || ''}  code ${r.code || ''}`); if (r.articles) L.push('  Contenu : ' + r.articles) })
  for (const [t, k] of [['Matériel du chantier', 'materiel'], ['Travaux à réaliser', 'travaux'], ['À vérifier avant de partir', 'verifs']]) {
    const x = (f[k] || []).filter(Boolean); if (x.length) L.push('', t + ' :', ...x.map((v: string) => '  - ' + v))
  }
  if (c.clos) L.push('', 'Chantier clôturé.')
  return L.join('\r\n')
}
export function textePV(v: Any) {
  const choix: Record<string, string> = { '1': 'Réception sans réserve', '2': 'Réception avec réserves', '3': 'Réception refusée' }
  const L = ['PV de réception signé dans l\'appli de suivi de chantier', '', 'Signé le : ' + quand(v.at), 'Date portée sur le PV : ' + (v.dateTxt || ''), 'Case cochée par le client : ' + (choix[String(v.choice)] || v.choice || ''), '']
  for (const [k, x] of Object.entries(v.fields || {})) L.push(`${k} : ${x ?? ''}`)
  return L.join('\r\n')
}
async function empreinte(o: unknown) {
  const h = await crypto.subtle.digest('SHA-1', new TextEncoder().encode(JSON.stringify(o)))
  return [...new Uint8Array(h)].map((b) => b.toString(16).padStart(2, '0')).join('')
}

// ---------- Microsoft Graph ----------
let jeton = '', expire = 0
async function graphJeton() {
  if (jeton && Date.now() < expire) return jeton
  const t = Deno.env.get('MS_TENANT_ID'), id = Deno.env.get('MS_CLIENT_ID'), sec = Deno.env.get('MS_CLIENT_SECRET')
  if (!t || !id || !sec) throw new Error('secrets MS_TENANT_ID, MS_CLIENT_ID ou MS_CLIENT_SECRET absents')
  const r = await fetch(`${LOGIN}/${t}/oauth2/v2.0/token`, { method: 'POST', body: new URLSearchParams({ grant_type: 'client_credentials', client_id: id, client_secret: sec, scope: 'https://graph.microsoft.com/.default' }) })
  const j = await r.json().catch(() => ({}))
  if (!r.ok || !j.access_token) throw new Error('connexion Microsoft refusée : ' + (j.error_description || j.error || r.status).toString().slice(0, 200))
  jeton = j.access_token; expire = Date.now() + (Number(j.expires_in || 3600) - 120) * 1000
  return jeton
}
const enc = (p: string) => p.split('/').map(encodeURIComponent).join('/')
// dépose un fichier (dossiers créés au passage) ; « remplacer » pour les textes, sinon nom unique (« photo 1.jpg »)
export async function deposer(chemin: string, corps: Uint8Array<ArrayBuffer> | string, type: string, remplacer: boolean, racine = RACINE) {
  const url = `${GRAPH}/drives/${DRIVE}/root:/${enc(racine + '/' + chemin)}:/content?@microsoft.graph.conflictBehavior=${remplacer ? 'replace' : 'rename'}`
  for (let i = 0; ; i++) {
    const r = await fetch(url, { method: 'PUT', headers: { Authorization: 'Bearer ' + await graphJeton(), 'Content-Type': type }, body: corps })
    if (r.ok) return (await r.json()).name as string
    if ((r.status === 429 || r.status >= 500) && i < 2) { await new Promise((s) => setTimeout(s, Number(r.headers.get('retry-after') || 2) * 1000)); continue }
    throw new Error(`SharePoint ${r.status} sur ${chemin} : ${(await r.text()).slice(0, 200)}`)
  }
}

// chemin d'un PDF de la fenêtre « Devis et factures » : <année>/<Devis|Factures|Avoirs>/<numéro> - <client>.pdf
const DOSSIER_DOC: Record<string, string> = { devis: 'Devis', facture: 'Factures', avoir: 'Avoirs' }
export function cheminDoc(d: Any) {
  const an = /^\d{4}/.test(String(d.date || '')) ? String(d.date).slice(0, 4) : String(d.numero).match(/-(\d{4})/)?.[1] || 'sans date'
  const num = nom(String(d.numero).replace(/\//g, '-'), 40), cl = d.client?.nom ? ' - ' + nom(d.client.nom, 40) : ''
  return `Devis et factures/${an}/${DOSSIER_DOC[d.kind] || 'Autres'}/${num}${cl}.pdf`
}

// ---------- un passage ----------
export async function passage(admin: Any, fin = Date.now() + DELAI_MS) {
  const lire = async (q: () => Any) => { const out: Any[] = []; for (let i = 0; ; i += 1000) { const { data, error } = await q().range(i, i + 999); if (error) throw error; out.push(...data); if (data.length < 1000) return out } }
  const { data: st, error: es } = await admin.from('copie_sharepoint').select('cle,valeur')
  if (es) throw es
  const E = new Map<string, string>((st || []).map((r: Any) => [r.cle, r.valeur]))
  const noter = async (cle: string, valeur: string) => { E.set(cle, valeur); const { error } = await admin.from('copie_sharepoint').upsert({ cle, valeur, maj: new Date().toISOString() }); if (error) throw error }
  const chantiers = await lire(() => admin.from('chantiers').select('id,nom,fiche,clos,created_at').order('id'))
  const els: El[] = await lire(() => admin.from('elements').select('chantier_id,type,id,data').order('chantier_id').order('type').order('id'))
  const pris = new Set([...E.entries()].filter(([k]) => k.startsWith('dossier|')).map(([, v]) => v))
  let copies = 0, erreurs = 0, reste = false
  for (const c of chantiers) {
    if (Date.now() > fin) { reste = true; break }
    // dossier « Chantier <ville> - <date de création> », figé à la première copie ; sans ville : on attend
    let dossier = E.get('dossier|' + c.id)
    if (!dossier) {
      const ville = String(c.fiche?.ville || '').trim(); if (!ville) continue
      const base = nom(`Chantier ${ville} - ${jour(c.fiche?.created || c.created_at)}`, 90)
      dossier = base; for (let n = 2; pris.has(dossier); n++) dossier = `${base} (${n})`
      pris.add(dossier); await noter('dossier|' + c.id, dossier)
    }
    const mine = els.filter((e) => e.chantier_id === c.id), devis: Record<string, Any> = {}
    mine.filter((e) => e.type === 'devis').forEach((e) => devis[e.id] = e.data)
    try {
      let h = await empreinte([c.nom, c.fiche, c.clos])
      if (E.get('fiche|' + c.id) !== h) { await deposer(`${dossier}/Fiche chantier.txt`, '﻿' + texteFiche(c), 'text/plain; charset=utf-8', true); await noter('fiche|' + c.id, h); copies++ }
      const pv = mine.find((e) => e.type === 'pv')?.data?.v
      if (pv) {
        h = await empreinte(pv)
        if (E.get('pv|' + c.id) !== h) {
          await deposer(`${dossier}/6 PV de réception/PV de réception.txt`, '﻿' + textePV(pv), 'text/plain; charset=utf-8', true)
          const m = String(pv.sig || '').match(/^data:image\/png;base64,(.+)$/)
          if (m) await deposer(`${dossier}/6 PV de réception/Signature du client.png`, Uint8Array.from(atob(m[1]), (x) => x.charCodeAt(0)), 'image/png', true)
          await noter('pv|' + c.id, h); copies++
        }
      }
    } catch (e) { erreurs++; console.error('copie-sharepoint', c.id, String(e)) }
    for (const el of mine) {
      if (!['photo', 'doc', 'pb', 'req', 'bl', 'frais', 'sav'].includes(el.type)) continue
      if (Date.now() > fin) { reste = true; break }
      const k = cleFichier(el); if (!k) continue
      const src = stockage(c.id, el.type, k)
      if (E.has('fichier|' + src)) continue
      const { data: blob, error } = await admin.storage.from('fichiers').download(src)
      if (error || !blob) continue // pas encore envoyé par le téléphone : on réessaiera
      let ext = EXT[(blob.type || '').split(';')[0].trim()] || ''
      if (el.type === 'doc') ext = (String(el.data?.name || '').match(/\.[^.]+$/) || [''])[0]
      try {
        const nomDepose = await deposer(`${dossier}/${destination(el, devis)}${ext || '.bin'}`, new Uint8Array(await blob.arrayBuffer()), blob.type || 'application/octet-stream', false)
        await noter('fichier|' + src, nomDepose); copies++
      } catch (e) { erreurs++; console.error('copie-sharepoint', src, String(e)) }
    }
  }
  // PDF figés de la fenêtre « Devis et factures » (table gestion, type « doc ») ; table absente ou vide : on passe
  try {
    const docs = await lire(() => admin.from('gestion').select('id,data').eq('type', 'doc').order('id'))
    for (const g of docs) {
      const d = g.data || {}
      if (!d.pdf || !d.numero) continue // brouillon : pas de numéro ni de PDF
      const cle = `gdoc|${g.id}|${d.pdf}`
      if (E.has(cle)) continue
      if (Date.now() > fin) { reste = true; break }
      try {
        const { data: blob, error } = await admin.storage.from('fichiers').download(String(d.pdf))
        if (error || !blob) continue // PDF pas encore envoyé : on réessaiera
        const nomDepose = await deposer(cheminDoc(d), new Uint8Array(await blob.arrayBuffer()), 'application/pdf', true, RACINE_GESTION)
        await noter(cle, nomDepose); copies++
      } catch (e) { erreurs++; console.error('copie-sharepoint', 'doc', g.id, String(e)) }
    }
  } catch (e) { console.error('copie-sharepoint', 'gestion', String(e?.message || e)) }
  return { copies, erreurs, reste }
}

const reponse = (corps: unknown, status = 200) => new Response(JSON.stringify(corps), { status, headers: { 'Content-Type': 'application/json' } })
function cle(...noms: string[]) {
  for (const n of noms) { const v = Deno.env.get(n); if (!v) continue; if (v.trim().startsWith('{')) { try { const k = JSON.parse(v); if (k.default) return String(k.default) } catch { /* format inattendu */ } } else return v }
  return ''
}

Deno.serve(async (req) => {
  const admin = createClient(Deno.env.get('SUPABASE_URL') || '', cle('SUPABASE_SECRET_KEYS', 'SUPABASE_SECRET_KEY', 'SUPABASE_SERVICE_ROLE_KEY'), { auth: { persistSession: false, autoRefreshToken: false } })
  // appel réservé à la tâche programmée de Supabase (en-tête x-copie-cle = ligne « cle_appel » de la table)
  const { data: k } = await admin.from('copie_sharepoint').select('valeur').eq('cle', 'cle_appel').maybeSingle()
  if (!k?.valeur || req.headers.get('x-copie-cle') !== k.valeur) return reponse({ message: 'Accès refusé' }, 403)
  // un seul passage à la fois
  const { data: v } = await admin.from('copie_sharepoint').select('valeur').eq('cle', 'verrou').maybeSingle()
  if (v && Date.now() - Date.parse(v.valeur) < 150_000) return reponse({ message: 'Passage déjà en cours' })
  await admin.from('copie_sharepoint').upsert({ cle: 'verrou', valeur: new Date().toISOString(), maj: new Date().toISOString() })
  const travail = passage(admin).then(async (r) => {
    await admin.from('copie_sharepoint').upsert({ cle: 'dernier', valeur: JSON.stringify({ le: new Date().toISOString(), ...r }), maj: new Date().toISOString() })
    return r
  }).catch(async (e) => {
    console.error('copie-sharepoint', e)
    await admin.from('copie_sharepoint').upsert({ cle: 'dernier', valeur: JSON.stringify({ le: new Date().toISOString(), erreur: String(e?.message || e).slice(0, 300) }), maj: new Date().toISOString() })
  }).finally(() => admin.from('copie_sharepoint').delete().eq('cle', 'verrou'))
  const rt = (globalThis as { EdgeRuntime?: { waitUntil(p: Promise<unknown>): void } }).EdgeRuntime
  if (rt?.waitUntil) { rt.waitUntil(travail); return reponse({ lance: true }, 202) }
  return reponse(await travail)
})
