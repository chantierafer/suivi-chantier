// Fonction Supabase « notif-push » (20/10/2026) : notification sur le téléphone de DOMetVIE à chaque avancement.
// Appelée par la base à chaque événement pour DOMetVIE (étape validée avec le %, matériel reçu, problème, PV signé,
// clôture ; voir ../../complement-v11-notifications.sql), elle envoie une notification Web Push à chaque téléphone
// DOMetVIE qui a touché « Activer les notifications » et qui a accès au chantier.
// Aucun secret à saisir : les clés d'envoi (VAPID) sont créées au premier appel et gardées dans la table push_reglages.
// Appels :
//   { action: 'cle' }                      → clé publique, pour l'abonnement du téléphone (appli)
//   { action: 'test' } + session du tel    → notification de test sur les téléphones de cet appareil (appli)
//   { evenement: id } + en-tête x-push-cle → envoi pour un événement (base de données seulement)
import { createClient } from 'npm:@supabase/supabase-js@2.117.2'

// deno-lint-ignore no-explicit-any
type Any = any
const SUJET = 'https://chantierafer.github.io/suivi-chantier/' // contact VAPID exigé par Apple et Google
const CORS = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type', 'Content-Type': 'application/json' }
const rep = (o: Any, status = 200) => new Response(JSON.stringify(o), { status, headers: CORS })

// ---------- base64url ----------
export const b64u = (b: Uint8Array) => btoa(String.fromCharCode(...b)).replace(/\+/g, '-').replace(/\//g, '_').replace(/=+$/, '')
export const deb64u = (s: string) => Uint8Array.from(atob(s.replace(/-/g, '+').replace(/_/g, '/') + '==='.slice((s.length + 3) % 4)), (c) => c.charCodeAt(0))
const cat = (...L: Uint8Array[]) => { const o = new Uint8Array(L.reduce((a, x) => a + x.length, 0)); let i = 0; for (const x of L) { o.set(x, i); i += x.length } return o }
const txt = (s: string) => new TextEncoder().encode(s)

// ---------- clés VAPID ----------
export async function nouvellesCles() {
  const k = await crypto.subtle.generateKey({ name: 'ECDSA', namedCurve: 'P-256' }, true, ['sign', 'verify'])
  const pub = new Uint8Array(await crypto.subtle.exportKey('raw', k.publicKey))
  return { publique: b64u(pub), prive: JSON.stringify(await crypto.subtle.exportKey('jwk', k.privateKey)) }
}
async function jwtVapid(endpoint: string, prive: string) {
  const key = await crypto.subtle.importKey('jwk', JSON.parse(prive), { name: 'ECDSA', namedCurve: 'P-256' }, false, ['sign'])
  const h = b64u(txt(JSON.stringify({ typ: 'JWT', alg: 'ES256' })))
  const c = b64u(txt(JSON.stringify({ aud: new URL(endpoint).origin, exp: Math.floor(Date.now() / 1000) + 12 * 3600, sub: SUJET })))
  const sig = new Uint8Array(await crypto.subtle.sign({ name: 'ECDSA', hash: 'SHA-256' }, key, txt(h + '.' + c)))
  return h + '.' + c + '.' + b64u(sig)
}

// ---------- chiffrement du message (RFC 8291, aes128gcm) ----------
async function hkdf(salt: Uint8Array, ikm: Uint8Array, info: Uint8Array, bits: number) {
  const k = await crypto.subtle.importKey('raw', ikm, 'HKDF', false, ['deriveBits'])
  return new Uint8Array(await crypto.subtle.deriveBits({ name: 'HKDF', hash: 'SHA-256', salt, info }, k, bits))
}
export async function chiffrer(sub: { keys: { p256dh: string; auth: string } }, message: string) {
  const ua = deb64u(sub.keys.p256dh), auth = deb64u(sub.keys.auth)
  const eph = await crypto.subtle.generateKey({ name: 'ECDH', namedCurve: 'P-256' }, true, ['deriveBits'])
  const as = new Uint8Array(await crypto.subtle.exportKey('raw', eph.publicKey))
  const uaKey = await crypto.subtle.importKey('raw', ua, { name: 'ECDH', namedCurve: 'P-256' }, false, [])
  const partage = new Uint8Array(await crypto.subtle.deriveBits({ name: 'ECDH', public: uaKey }, eph.privateKey, 256))
  const ikm = await hkdf(auth, partage, cat(txt('WebPush: info\0'), ua, as), 256)
  const sel = crypto.getRandomValues(new Uint8Array(16))
  const cek = await hkdf(sel, ikm, txt('Content-Encoding: aes128gcm\0'), 128)
  const nonce = await hkdf(sel, ikm, txt('Content-Encoding: nonce\0'), 96)
  const k = await crypto.subtle.importKey('raw', cek, 'AES-GCM', false, ['encrypt'])
  const chiffre = new Uint8Array(await crypto.subtle.encrypt({ name: 'AES-GCM', iv: nonce }, k, cat(txt(message), new Uint8Array([2]))))
  const rs = new Uint8Array([0, 0, 16, 0]) // 4096
  return cat(sel, rs, new Uint8Array([as.length]), as, chiffre)
}

// Envoi à un téléphone. Renvoie le code HTTP du service de notification (Apple, Google, Mozilla…).
export async function envoyer(sub: Any, message: Any, cles: { publique: string; prive: string }) {
  const corps = await chiffrer(sub, JSON.stringify(message))
  const r = await fetch(sub.endpoint, {
    method: 'POST',
    headers: { TTL: '86400', Urgency: 'high', 'Content-Encoding': 'aes128gcm', 'Content-Type': 'application/octet-stream',
      Authorization: `vapid t=${await jwtVapid(sub.endpoint, cles.prive)}, k=${cles.publique}` },
    body: corps,
  })
  await r.body?.cancel()
  return r.status
}

// ---------- texte de la notification ----------
// « Étape 7/12 validée : Revêtement mural · 58 % » → titre « Chantier Saint-Denis · 58 % », texte « Étape 7/12 validée : Revêtement mural »
export function message(ev: Any, ch: Any, url: string) {
  const f = ch?.fiche || {}
  const nom = f.ville ? 'Chantier ' + f.ville : 'Nouveau chantier'
  const m = String(ev.texte || '').match(/^(.*?)\s*·\s*(\d+(?:,\d+)?)\s*%\s*$/)
  const lien = url ? url.replace(/[?&]ch=[^&]*/, '') + (url.includes('?') ? '&' : '?') + 'ch=' + encodeURIComponent(ev.chantier_id) : ''
  return { title: m ? `${nom} · ${m[2]} %` : nom, body: m ? m[1] : String(ev.texte || ''), url: lien, tag: 'ev' + ev.id }
}

// Clés fournies par Supabase (nouveau format JSON { default: … } ou ancien format texte)
function cle(...noms: string[]) {
  for (const n of noms) { const v = Deno.env.get(n); if (!v) continue; if (v.trim().startsWith('{')) { try { const k = JSON.parse(v); if (k.default) return String(k.default) } catch { /* format inattendu */ } } else return v }
  return ''
}
const SANS_SESSION = { auth: { persistSession: false, autoRefreshToken: false } }

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') return new Response('ok', { headers: CORS })
  const url = Deno.env.get('SUPABASE_URL') || ''
  const db = createClient(url, cle('SUPABASE_SECRET_KEYS', 'SUPABASE_SECRET_KEY', 'SUPABASE_SERVICE_ROLE_KEY'), SANS_SESSION)
  const reglage = async (cle: string) => (await db.from('push_reglages').select('valeur').eq('cle', cle).maybeSingle()).data?.valeur as string | undefined
  async function cles() {
    let publique = await reglage('vapid_publique'), prive = await reglage('vapid_prive')
    if (!publique || !prive) {
      const n = await nouvellesCles()
      await db.from('push_reglages').upsert([{ cle: 'vapid_publique', valeur: n.publique }, { cle: 'vapid_prive', valeur: n.prive }], { onConflict: 'cle', ignoreDuplicates: true })
      publique = await reglage('vapid_publique'); prive = await reglage('vapid_prive')
    }
    if (!publique || !prive) throw new Error('Table push_reglages absente : coller complement-v11-notifications.sql')
    return { publique, prive }
  }
  // Envoie et retire les abonnements morts (téléphone désinstallé, notifications coupées)
  async function diffuser(cibles: Any[], msg: (c: Any) => Any) {
    const k = await cles()
    const res = await Promise.all(cibles.map(async (c) => {
      try {
        const st = await envoyer(c.sub, msg(c), k)
        if (st === 404 || st === 410) await db.from('push_abonnements').delete().eq('endpoint', c.endpoint)
        return st
      } catch (e) { console.error('envoi', e); return 0 }
    }))
    return { envoyes: res.filter((s) => s >= 200 && s < 300).length, total: res.length, codes: res }
  }
  try {
    const body: Any = await req.json().catch(() => ({}))
    if (body.action === 'cle') return rep({ cle: (await cles()).publique })

    if (body.action === 'test') {
      const auth = req.headers.get('Authorization') || '', jwt = auth.replace(/^Bearer\s+/i, '')
      const moi = createClient(url, cle('SUPABASE_PUBLISHABLE_KEYS', 'SUPABASE_PUBLISHABLE_KEY', 'SUPABASE_ANON_KEY'), { ...SANS_SESSION, global: { headers: { Authorization: auth } } })
      const { data: u } = jwt && jwt !== auth ? await moi.auth.getUser(jwt) : { data: null }
      if (!u?.user) return rep({ erreur: 'Connexion requise' }, 401)
      const { data: subs } = await db.from('push_abonnements').select('endpoint,sub,url').eq('user_id', u.user.id)
      if (!subs?.length) return rep({ envoyes: 0, total: 0 })
      return rep(await diffuser(subs, (c) => ({ title: 'Suivi chantier AFER', body: 'Notifications activées : vous serez prévenu à chaque étape validée.', url: c.url || '', tag: 'test' })))
    }

    if (body.evenement) {
      const cle = await reglage('cle_appel')
      if (!cle || req.headers.get('x-push-cle') !== cle) return rep({ erreur: 'Appel refusé' }, 403)
      const { data: ev } = await db.from('evenements').select('*').eq('id', body.evenement).maybeSingle()
      if (!ev) return rep({ erreur: 'Événement introuvable' }, 404)
      const { data: ch } = await db.from('chantiers').select('id,fiche').eq('id', ev.chantier_id).maybeSingle()
      const { data: cibles, error } = await db.rpc('push_cibles', { p_ch: ev.chantier_id })
      if (error) throw error
      return rep(await diffuser(cibles || [], (c) => message(ev, ch, c.url || '')))
    }
    return rep({ erreur: 'Demande inconnue' }, 400)
  } catch (e) {
    console.error(e)
    return rep({ erreur: String((e as Error)?.message || e) }, 500)
  }
})
