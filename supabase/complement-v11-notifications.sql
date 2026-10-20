-- V11 (20/10/2026) : notification sur le téléphone de DOMetVIE à chaque avancement.
-- À coller une fois dans Supabase > SQL Editor, puis Run. Rejouable sans risque.
-- Prérequis : la fonction « notif-push » (supabase/fonctions/notif-push/index.ts) déployée, Verify JWT désactivé.

-- 1. Téléphones abonnés (un par appareil) : invisibles depuis l'appli, gérés par les fonctions ci-dessous
create table if not exists public.push_abonnements (
  endpoint   text primary key,
  user_id    uuid not null references auth.users (id) on delete cascade,
  sub        jsonb not null,
  url        text,
  created_at timestamptz not null default now()
);
alter table public.push_abonnements enable row level security; -- aucune règle : fonctions seulement

-- 2. Réglages de l'envoi : clé d'appel (base → fonction) et clés VAPID (créées par la fonction au premier appel)
create table if not exists public.push_reglages (
  cle    text primary key,
  valeur text not null
);
alter table public.push_reglages enable row level security; -- aucune règle : invisible depuis l'appli
insert into public.push_reglages (cle, valeur) values ('cle_appel', gen_random_uuid()::text) on conflict (cle) do nothing;

-- 3. Abonnement / désabonnement du téléphone (bouton « Activer les notifications » de l'espace DOMetVIE)
create or replace function public.push_abonner(p_sub jsonb, p_url text) returns void
language plpgsql security definer set search_path = public as $$
begin
  if auth.uid() is null then raise exception 'Connexion requise'; end if;
  if not public.est_afer() and not exists (select 1 from public.mes_acces() where role = 'domevie') then
    raise exception 'Réservé à DOMetVIE';
  end if;
  if coalesce(p_sub->>'endpoint', '') !~ '^https://' then raise exception 'Abonnement invalide'; end if;
  insert into public.push_abonnements (endpoint, user_id, sub, url)
    values (p_sub->>'endpoint', auth.uid(), p_sub, left(p_url, 500))
    on conflict (endpoint) do update set user_id = excluded.user_id, sub = excluded.sub, url = excluded.url;
end $$;

create or replace function public.push_desabonner(p_endpoint text) returns void
language sql security definer set search_path = public as $$
  delete from public.push_abonnements where endpoint = p_endpoint and user_id = auth.uid();
$$;

-- Téléphones DOMetVIE à prévenir pour un chantier (lien actif sur ce chantier ou sur tous les chantiers)
create or replace function public.push_cibles(p_ch text)
returns table (endpoint text, sub jsonb, url text)
language sql stable security definer set search_path = public as $$
  select distinct p.endpoint, p.sub, p.url
  from public.push_abonnements p
  join public.membres m on m.user_id = p.user_id
  join public.acces a on a.token = m.token
  where a.actif and a.role = 'domevie' and (a.chantier_id is null or a.chantier_id = p_ch);
$$;

revoke execute on function public.push_abonner(jsonb, text) from public, anon;
revoke execute on function public.push_desabonner(text) from public, anon;
revoke execute on function public.push_cibles(text) from public, anon, authenticated;
grant execute on function public.push_abonner(jsonb, text) to authenticated;
grant execute on function public.push_desabonner(text) to authenticated;
grant execute on function public.push_cibles(text) to service_role;

-- 4. Chaque alerte de suivi pour DOMetVIE (étape validée, matériel reçu, problème, PV signé, clôture) appelle la fonction.
--    L'appel part après l'enregistrement et ne bloque jamais l'appli, même si la fonction est absente.
create extension if not exists pg_net;
create or replace function public.push_evenement() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  if 'domevie' = any (new.pour) and coalesce(new.type, '') in ('suivi', 'pv') then
    perform net.http_post(
      url := 'https://xgliespwrxlxiaqaxtzj.supabase.co/functions/v1/notif-push',
      headers := jsonb_build_object('Content-Type', 'application/json',
        'x-push-cle', (select valeur from public.push_reglages where cle = 'cle_appel')),
      body := jsonb_build_object('evenement', new.id),
      timeout_milliseconds := 10000);
  end if;
  return null;
exception when others then
  return null;
end $$;
drop trigger if exists evenements_push on public.evenements;
create trigger evenements_push after insert on public.evenements
  for each row execute function public.push_evenement();

-- Contrôle : select count(*) from public.push_abonnements;  (nombre de téléphones abonnés)
-- Derniers envois : select status_code, content from net._http_response order by id desc limit 5;
