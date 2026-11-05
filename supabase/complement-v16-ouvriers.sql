-- V16 (05/11/2026) : espace ouvrier. Un lien par personne (page ouvrier.html?o=…), valable pour tous ses chantiers.
-- À coller une fois dans Supabase > SQL Editor, puis Run. Rejouable sans risque. Prérequis : complement-v12-gestion.sql.
--
-- Le lien est créé par AFER dans Gestion > Personnel et matériel > fiche de la personne > « Créer son lien ouvrier » :
-- table gestion, type 'lien_ouvrier', id = 'p:<ID Excel>', data = {jeton, actif}.
-- L'ouvrier voit ses affectations du planning (7 jours passés, 60 jours à venir), sans aucun prix.
-- Sur une salle de bain (chantier de l'appli SDB), il reçoit un accès technicien à ce chantier tant qu'il y est placé :
-- une ligne de public.acces (rôle tech, libellé « Espace ouvrier · <nom> [p:…] ») créée, réactivée ou coupée
-- automatiquement à chaque changement du planning ou du lien.

-- 1. Accès technicien des chantiers SDB où la personne est placée
create or replace function public.ouvrier_sync(p_r text) returns void
language plpgsql security definer set search_path = public as $$
declare
  lien jsonb;
  tag  text := '[' || p_r || ']';
  nom  text;
  ch   text;
begin
  if p_r is null or p_r = '' then return; end if;
  select data into lien from public.gestion where type = 'lien_ouvrier' and id = p_r;
  if lien is null or coalesce((lien->>'actif')::boolean, false) = false then
    update public.acces set actif = false
      where role = 'tech' and actif and right(coalesce(libelle, ''), length(tag)) = tag;
    return;
  end if;
  select coalesce(nullif(data->>'Nom affiché dans l''appli', ''), nullif(data->>'Prénom', ''), p_r) into nom
    from public.gestion where type = 'personne' and id = substr(p_r, 3);
  for ch in
    select distinct g.data->>'ch' from public.gestion g
      join public.chantiers c on c.id = g.data->>'ch'
     where g.type = 'aff' and g.data->>'r' = p_r
  loop
    if exists (select 1 from public.acces where chantier_id = ch and role = 'tech' and right(coalesce(libelle, ''), length(tag)) = tag) then
      update public.acces set actif = true
        where chantier_id = ch and role = 'tech' and not actif and right(coalesce(libelle, ''), length(tag)) = tag;
    else
      insert into public.acces (chantier_id, role, libelle)
        values (ch, 'tech', 'Espace ouvrier · ' || coalesce(nom, p_r) || ' ' || tag);
    end if;
  end loop;
  update public.acces set actif = false
   where role = 'tech' and actif and right(coalesce(libelle, ''), length(tag)) = tag
     and chantier_id not in (
       select distinct data->>'ch' from public.gestion
        where type = 'aff' and data->>'r' = p_r and data->>'ch' is not null);
end $$;
revoke execute on function public.ouvrier_sync(text) from public, anon, authenticated;

-- 2. Recalcul à chaque affectation ajoutée, modifiée ou retirée, et à chaque lien créé ou désactivé
create or replace function public.gestion_ouvrier() returns trigger
language plpgsql security definer set search_path = public as $$
begin
  if tg_op in ('UPDATE', 'DELETE') and old.type in ('aff', 'lien_ouvrier') then
    perform public.ouvrier_sync(case when old.type = 'aff' then old.data->>'r' else old.id end);
  end if;
  if tg_op in ('INSERT', 'UPDATE') and new.type in ('aff', 'lien_ouvrier') then
    perform public.ouvrier_sync(case when new.type = 'aff' then new.data->>'r' else new.id end);
  end if;
  return null;
end $$;
drop trigger if exists gestion_ouvrier on public.gestion;
create trigger gestion_ouvrier after insert or update or delete on public.gestion
  for each row execute function public.gestion_ouvrier();

-- 3. Ce que voit l'ouvrier avec son lien (appelée par ouvrier.html, sans compte)
create or replace function public.espace_ouvrier(p_jeton text) returns json
language plpgsql security definer set search_path = public as $$
declare
  rid text;
  nom text;
  tag text;
  res json;
begin
  if p_jeton is null or length(p_jeton) < 32 then raise exception 'Lien invalide ou désactivé'; end if;
  select id into rid from public.gestion
   where type = 'lien_ouvrier' and data->>'jeton' = p_jeton and coalesce((data->>'actif')::boolean, false);
  if rid is null then raise exception 'Lien invalide ou désactivé'; end if;
  tag := '[' || rid || ']';
  select coalesce(nullif(data->>'Nom affiché dans l''appli', ''), nullif(data->>'Prénom', ''), rid) into nom
    from public.gestion where type = 'personne' and id = substr(rid, 3);
  select coalesce(json_agg(x order by x.d, x.c nulls first, x.nom), '[]'::json) into res from (
    select a.data->>'d' as d, a.data->>'c' as c, a.data->>'m' as m,
           case when sc.id is not null then coalesce('Chantier ' || nullif(sc.fiche->>'ville', ''), sc.nom)
                else coalesce(gc.data->>'nom', 'Chantier') end as nom,
           coalesce(sc.fiche->>'adresse', '') as adresse,
           coalesce(nullif(sc.fiche->>'ville', ''), gc.data->>'ville', '') as ville,
           (sc.id is not null) as sdb,
           (select ac.token from public.acces ac
             where ac.chantier_id = sc.id and ac.role = 'tech' and ac.actif
               and right(coalesce(ac.libelle, ''), length(tag)) = tag limit 1) as k
      from public.gestion a
      left join public.chantiers sc on sc.id = a.data->>'ch'
      left join public.gestion gc on gc.type = 'chantier' and gc.id = a.data->>'ch'
     where a.type = 'aff' and a.data->>'r' = rid
       and a.data->>'d' >= to_char(current_date - 7, 'YYYY-MM-DD')
       and a.data->>'d' <= to_char(current_date + 60, 'YYYY-MM-DD')) x;
  return json_build_object('nom', nom, 'missions', res);
end $$;
revoke execute on function public.espace_ouvrier(text) from public;
grant execute on function public.espace_ouvrier(text) to anon, authenticated;
