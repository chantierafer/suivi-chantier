-- V12 (26/10/2026) : gestion AFER (planning, personnel et matériel, clients, rentabilité) reliée aux deux Excel
-- du dossier SharePoint « Gestion AFER ». À coller une fois dans Supabase > SQL Editor, puis Run. Rejouable sans risque.
-- Prérequis : la fonction « synchro-excel » (supabase/fonctions/synchro-excel/index.ts) déployée, Verify JWT désactivé.
-- Les Excel restent la référence : la table « gestion » en garde une copie pour l'appli et reçoit ses modifications,
-- la fonction fait l'aller-retour avec les fichiers (voir gestion/CONTRAT-SYNCHRO-EXCEL.md).

-- 1. Données de la gestion AFER : une ligne par fiche (type + id), contenu libre dans data. AFER seul.
create table if not exists public.gestion (
  type   text not null,
  id     text not null,
  data   jsonb not null default '{}'::jsonb,
  maj_le timestamptz not null default now(),
  maj_par uuid default auth.uid(),
  primary key (type, id)
);
alter table public.gestion replica identity full; -- l'appli voit aussi passer les suppressions en direct
alter table public.gestion enable row level security;
drop policy if exists gestion_afer on public.gestion;
create policy gestion_afer on public.gestion for all to authenticated
  using (public.est_afer()) with check (public.est_afer());

-- Horodatage et auteur de chaque modification
create or replace function public.gestion_maj() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  new.maj_le := now();
  new.maj_par := auth.uid();
  return new;
end $$;
drop trigger if exists gestion_maj on public.gestion;
create trigger gestion_maj before insert or update on public.gestion
  for each row execute function public.gestion_maj();

-- 2. Clé d'appel de la fonction (tirée au hasard ici) et verrou d'un passage : invisibles depuis l'appli
create table if not exists public.synchro_excel (
  cle text primary key,
  valeur text not null,
  maj timestamptz not null default now()
);
alter table public.synchro_excel enable row level security; -- aucune règle : fonction seulement
insert into public.synchro_excel (cle, valeur) values ('cle_appel', gen_random_uuid()::text) on conflict (cle) do nothing;

-- 3. Une modification faite dans l'appli (ou « Synchroniser maintenant ») appelle la fonction tout de suite.
--    L'appel part après l'enregistrement et ne bloque jamais l'appli, même si la fonction est absente.
--    Les écritures de la fonction elle-même (clé service) ne la relancent pas.
create extension if not exists pg_net;
create or replace function public.gestion_vers_excel() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  if coalesce(nullif(current_setting('request.jwt.claims', true), '')::jsonb ->> 'role', '') = 'service_role' then
    return null;
  end if;
  perform net.http_post(
    url := 'https://xgliespwrxlxiaqaxtzj.supabase.co/functions/v1/synchro-excel',
    headers := jsonb_build_object('Content-Type', 'application/json',
      'x-synchro-cle', (select valeur from public.synchro_excel where cle = 'cle_appel')),
    body := jsonb_build_object('type', new.type, 'id', new.id),
    timeout_milliseconds := 10000);
  return null;
exception when others then
  return null;
end $$;
drop trigger if exists gestion_vers_excel on public.gestion;
create trigger gestion_vers_excel after insert or update on public.gestion
  for each row when (new.type in ('excel_ecrire', 'excel_demande'))
  execute function public.gestion_vers_excel();

-- 4. Lecture des Excel toutes les 3 minutes (ce que Nariman ou toi modifiez dans Excel arrive dans l'appli)
create extension if not exists pg_cron;
select cron.unschedule('synchro-excel') where exists (select 1 from cron.job where jobname = 'synchro-excel');
select cron.schedule('synchro-excel', '*/3 * * * *', $$
  select net.http_post(
    url := 'https://xgliespwrxlxiaqaxtzj.supabase.co/functions/v1/synchro-excel',
    headers := jsonb_build_object('Content-Type', 'application/json',
      'x-synchro-cle', (select valeur from public.synchro_excel where cle = 'cle_appel')),
    body := '{}'::jsonb,
    timeout_milliseconds := 10000
  )
$$);

-- 5. Direct (Realtime) : l'appli se met à jour dès qu'une fiche change
do $$
begin
  if not exists (select 1 from pg_publication_tables
                 where pubname = 'supabase_realtime' and schemaname = 'public' and tablename = 'gestion') then
    alter publication supabase_realtime add table public.gestion;
  end if;
end $$;

-- Contrôle : select data from public.gestion where type = 'excel_etat';  (dernier passage, fichiers lus, erreurs)
-- Fiches lues : select type, count(*) from public.gestion group by type order by type;
-- Derniers appels : select status_code, content from net._http_response order by id desc limit 5;
-- Fin. Le message attendu est « Success. No rows returned ».
