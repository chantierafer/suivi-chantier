// Fonction Supabase « lire-ticket » (V17, 09/11/2026) : lit la photo d'un ticket (carburant, péage, parking, lavage)
// envoyée depuis l'espace ouvrier, et renvoie de quoi pré-remplir le formulaire. Rien n'est enregistré ici :
// l'ouvrier vérifie les valeurs, puis l'espace ouvrier les enregistre avec ouvrier_depense (complement-v17-depenses-vehicules.sql).
// Appelée sans compte : la fonction vérifie elle-même le lien de l'ouvrier (jeton actif) ; la vérification JWT peut rester coupée.
// Secret déjà en place pour lire-documents : ANTHROPIC_API_KEY. Modèle : secret MODELE_TICKET, sinon Haiku 5.5 (peu coûteux).
import { createClient } from 'npm:@supabase/supabase-js@2.117.2'
import Anthropic from 'npm:@anthropic-ai/sdk@0.135.0'

const MODELE = Deno.env.get('MODELE_TICKET') || 'claude-haiku-5-5'
const MAX_CARACTERES = 900_000 // photo réduite par l'espace ouvrier (environ 1280 px, JPEG)
const CORS = { 'Access-Control-Allow-Origin': '*', 'Access-Control-Allow-Headers': 'authorization, x-client-info, apikey, content-type', 'Access-Control-Allow-Methods': 'POST, OPTIONS' }
const reponse = (corps: unknown, status = 200) => new Response(JSON.stringify(corps), { status, headers: { ...CORS, 'Content-Type': 'application/json' } })

const nombre = (description: string) => ({ type: ['number', 'null'], description })
const SCHEMA = {
  type: 'object',
  additionalProperties: false,
  required: ['type', 'montant', 'litres', 'date', 'station'],
  properties: {
    type: { type: 'string', enum: ['plein', 'peage', 'parking', 'lavage', 'autre', 'illisible'], description: 'Nature du ticket ; « illisible » si ce n’est pas un ticket ou s’il ne se lit pas' },
    montant: nombre('Montant total payé TTC en euros, tel qu’imprimé ; null s’il ne se lit pas'),
    litres: nombre('Volume de carburant en litres pour un plein ; null sinon'),
    date: { type: ['string', 'null'], description: 'Date du ticket au format AAAA-MM-JJ ; null si absente' },
    station: { type: 'string', description: 'Nom et ville de la station, du péage ou du parking, tels qu’écrits ; vide si absent' },
  },
}

Deno.serve(async (req) => {
  if (req.method === 'OPTIONS') return new Response('ok', { headers: CORS })
  if (req.method !== 'POST') return reponse({ erreur: 'POST attendu' }, 405)
  let corps: { jeton?: string; image?: string }
  try { corps = await req.json() } catch { return reponse({ erreur: 'Demande illisible' }, 400) }
  const jeton = String(corps.jeton || ''), image = String(corps.image || '')
  const m = image.match(/^data:(image\/(?:jpeg|png|webp));base64,(.+)$/)
  if (!m || image.length > MAX_CARACTERES) return reponse({ erreur: 'Photo absente ou trop lourde' }, 400)

  // le lien de l'ouvrier doit exister et être actif
  const admin = createClient(Deno.env.get('SUPABASE_URL')!, Deno.env.get('SUPABASE_SERVICE_ROLE_KEY')!, { auth: { persistSession: false } })
  if (jeton.length < 32) return reponse({ erreur: 'Lien invalide ou désactivé' }, 403)
  const { data: lien } = await admin.from('gestion').select('id,data').eq('type', 'lien_ouvrier').eq('data->>jeton', jeton).maybeSingle()
  if (!lien || lien.data?.actif !== true) return reponse({ erreur: 'Lien invalide ou désactivé' }, 403)

  const cle = Deno.env.get('ANTHROPIC_API_KEY')
  if (!cle) return reponse({ erreur: 'Lecture non configurée (secret ANTHROPIC_API_KEY)' }, 503)
  const client = new Anthropic({ apiKey: cle, baseURL: Deno.env.get('ANTHROPIC_BASE_URL') || undefined, timeout: 60_000, maxRetries: 2 })
  try {
    const rep = await client.messages.create({
      model: MODELE,
      max_tokens: 1000,
      // deno-lint-ignore no-explicit-any
      output_config: { format: { type: 'json_schema', schema: SCHEMA } } as any,
      messages: [{
        role: 'user',
        content: [
          // deno-lint-ignore no-explicit-any
          { type: 'image', source: { type: 'base64', media_type: m[1] as any, data: m[2] } },
          { type: 'text', text: 'Ticket de caisse photographié par un ouvrier du bâtiment pour une dépense de véhicule. Recopie exactement ce qui est imprimé, sans rien deviner : si une valeur ne se lit pas, mets null.' },
        ],
      }],
    // deno-lint-ignore no-explicit-any
    } as any)
    // deno-lint-ignore no-explicit-any
    const bloc = (rep.content as any[]).find((b) => b.type === 'text')
    if (rep.stop_reason === 'refusal' || !bloc) return reponse({ erreur: 'Ticket non lu' }, 422)
    const lu = JSON.parse(bloc.text)
    return reponse({ lu })
  } catch (e) {
    console.error('lire-ticket :', (e as Error)?.message)
    return reponse({ erreur: 'Lecture impossible pour le moment' }, 502)
  }
})
