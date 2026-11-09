-- V17 (09/11/2026) : dépenses des véhicules saisies par les ouvriers (plein, péage, parking, lavage, autre).
-- À coller une fois dans Supabase > SQL Editor, puis Run. Rejouable sans risque. Prérequis : complement-v16-ouvriers.sql.
--
-- L'ouvrier saisit la dépense depuis son espace (ouvrier.html?o=…), sans compte : photo du ticket, montant, litres, kilométrage,
-- véhicule et chantier. Rien n'est lisible sans le lien : il ne voit que ses propres dépenses, jamais celles des autres.
-- Rangement dans la table gestion (lisible par AFER seul) :
--   type 'depveh'       id 'D-AAAAMMJJ-HHMISS-xxxx'  data {type, montant, litres, km, veh, ch, d, station, par, at, photo, lu}
--   type 'depveh_photo' même id                         data {img}  (photo réduite, lue seulement quand AFER l'ouvre)

-- 1. Le lien de l'ouvrier -> son identifiant 'p:<ID>' (erreur si le lien est invalide ou coupé)
create or replace function public.ouvrier_rid(p_jeton text) returns text
language plpgsql security definer set search_path = public as $$
declare rid text;
begin
  if p_jeton is null or length(p_jeton) < 32 then raise exception 'Lien invalide ou désactivé'; end if;
  select id into rid from public.gestion
   where type = 'lien_ouvrier' and data->>'jeton' = p_jeton and coalesce((data->>'actif')::boolean, false);
  if rid is null then raise exception 'Lien invalide ou désactivé'; end if;
  return rid;
end $$;
revoke execute on function public.ouvrier_rid(text) from public, anon, authenticated;

-- 2. Ce qu'il faut pour le formulaire : véhicules en service, véhicules placés aujourd'hui, ses 15 dernières dépenses
create or replace function public.ouvrier_vehicules(p_jeton text) returns json
language plpgsql security definer set search_path = public as $$
declare
  rid text := public.ouvrier_rid(p_jeton);
  auj text := to_char((now() at time zone 'Europe/Paris')::date, 'YYYY-MM-DD');
  veh json; jour json; deps json;
begin
  select coalesce(json_agg(json_build_object(
           'id', v.id,
           'nom', coalesce(nullif(v.data->>'Marque / modèle', ''), v.id),
           'immat', coalesce(v.data->>'Immatriculation', ''),
           'habituel', coalesce(v.data->>'Conducteur habituel (ID personne)', '') = substr(rid, 3))
         order by v.id), '[]'::json) into veh
    from public.gestion v
   where v.type = 'vehicule' and coalesce(v.data->>'Statut', '') not in ('Vendu', 'Réformé', 'Sorti');
  select coalesce(json_agg(json_build_object('v', substr(a.data->>'r', 3), 'ch', a.data->>'ch')), '[]'::json) into jour
    from public.gestion a
   where a.type = 'aff' and a.data->>'d' = auj and a.data->>'r' like 'v:%';
  select coalesce(json_agg(s.j order by s.at desc), '[]'::json) into deps from (
    select (d.data - 'par') || jsonb_build_object('id', d.id) as j, d.data->>'at' as at
      from public.gestion d
     where d.type = 'depveh' and d.data->>'par' = rid
     order by d.data->>'at' desc limit 15) s;
  return json_build_object('auj', auj, 'vehicules', veh, 'jour', jour, 'depenses', deps);
end $$;
revoke execute on function public.ouvrier_vehicules(text) from public;
grant execute on function public.ouvrier_vehicules(text) to anon, authenticated;

-- 3. Enregistrer une dépense (contrôles côté base : type connu, montant entre 0 et 5 000 €, véhicule existant)
create or replace function public.ouvrier_depense(p_jeton text, p_dep jsonb) returns text
language plpgsql security definer set search_path = public as $$
declare
  rid text := public.ouvrier_rid(p_jeton);
  t text := coalesce(p_dep->>'type', '');
  m numeric; l numeric; k numeric;
  v text := coalesce(p_dep->>'veh', '');
  img text := p_dep->>'photo';
  nid text := 'D-' || to_char(now() at time zone 'Europe/Paris', 'YYYYMMDD-HH24MISS') || '-' || substr(md5(random()::text), 1, 4);
  d text := coalesce(nullif(p_dep->>'d', ''), to_char((now() at time zone 'Europe/Paris')::date, 'YYYY-MM-DD'));
begin
  if t not in ('plein', 'peage', 'parking', 'lavage', 'autre') then raise exception 'Type de dépense inconnu'; end if;
  m := nullif(p_dep->>'montant', '')::numeric;
  if m is null or m <= 0 or m > 5000 then raise exception 'Montant à vérifier'; end if;
  l := nullif(p_dep->>'litres', '')::numeric;
  if l is not null and (l <= 0 or l > 400) then raise exception 'Litres à vérifier'; end if;
  k := nullif(p_dep->>'km', '')::numeric;
  if k is not null and (k < 0 or k > 2000000) then raise exception 'Kilométrage à vérifier'; end if;
  if not exists (select 1 from public.gestion where type = 'vehicule' and id = v) then raise exception 'Véhicule inconnu'; end if;
  if d !~ '^\d{4}-\d{2}-\d{2}$' then raise exception 'Date à vérifier'; end if;
  if img is not null and (length(img) > 900000 or img !~ '^data:image/(jpeg|png|webp);base64,') then img := null; end if;
  insert into public.gestion (type, id, data, maj_le, maj_par) values ('depveh', nid, jsonb_build_object(
      'type', t, 'montant', m, 'litres', l, 'km', k, 'veh', v,
      'ch', nullif(p_dep->>'ch', ''), 'd', d, 'station', left(coalesce(p_dep->>'station', ''), 120),
      'lu', coalesce((p_dep->>'lu')::boolean, false), 'photo', img is not null,
      'par', rid, 'at', to_char(now() at time zone 'UTC', 'YYYY-MM-DD"T"HH24:MI:SS"Z"')), now(), null);
  if img is not null then
    insert into public.gestion (type, id, data, maj_le, maj_par) values ('depveh_photo', nid, jsonb_build_object('img', img), now(), null);
  end if;
  return nid;
end $$;
revoke execute on function public.ouvrier_depense(text, jsonb) from public;
grant execute on function public.ouvrier_depense(text, jsonb) to anon, authenticated;
