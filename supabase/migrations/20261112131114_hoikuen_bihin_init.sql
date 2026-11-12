-- 保育園 物品管理台帳: multi-device shared storage
-- Server is the single source of truth. Each 園 (org) is isolated by RLS.
-- Records are JSON documents keyed by the client id; writes are partial (field-level) patches
-- applied through public.hb_apply so concurrent edits of different fields never overwrite each other.
-- Deletes are tombstones (deleted = true) so that nothing silently disappears and other devices learn about them.

create schema if not exists hb_private;
revoke all on schema hb_private from public, anon, authenticated;
grant usage on schema hb_private to authenticated, service_role;

-- ---------------------------------------------------------------- orgs / memberships
create table public.orgs (
  id uuid primary key default gen_random_uuid(),
  name text not null check (length(name) between 1 and 100),
  created_at timestamptz not null default now()
);

create table public.memberships (
  org_id uuid not null references public.orgs (id) on delete cascade,
  user_id uuid not null references auth.users (id) on delete cascade,
  role text not null check (role in ('admin', 'member')),
  created_at timestamptz not null default now(),
  primary key (org_id, user_id)
);
create index memberships_user_idx on public.memberships (user_id);

create function hb_private.is_member(p_org uuid) returns boolean
language sql stable security definer set search_path = '' as $$
  select exists (select 1 from public.memberships m where m.org_id = p_org and m.user_id = (select auth.uid()));
$$;
create function hb_private.is_admin(p_org uuid) returns boolean
language sql stable security definer set search_path = '' as $$
  select exists (select 1 from public.memberships m where m.org_id = p_org and m.user_id = (select auth.uid()) and m.role = 'admin');
$$;
-- storage paths start with the org id; anything that is not a uuid is never a member
create function hb_private.is_member_path(p_seg text) returns boolean
language sql stable security definer set search_path = '' as $$
  select case when p_seg ~* '^[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}$'
    then hb_private.is_member(p_seg::uuid) else false end;
$$;
revoke all on function hb_private.is_member(uuid), hb_private.is_admin(uuid), hb_private.is_member_path(text) from public, anon;
grant execute on function hb_private.is_member(uuid), hb_private.is_admin(uuid), hb_private.is_member_path(text) to authenticated;

alter table public.orgs enable row level security;
alter table public.memberships enable row level security;
revoke all on public.orgs, public.memberships from anon, authenticated;
grant select on public.orgs, public.memberships to authenticated;
create policy orgs_select on public.orgs for select to authenticated using (hb_private.is_member(id));
-- own membership rows, or every row of an org where the caller is admin. No client writes at all.
create policy memberships_select on public.memberships for select to authenticated
  using (user_id = (select auth.uid()) or hb_private.is_admin(org_id));

-- ---------------------------------------------------------------- data tables
create function hb_private.touch() returns trigger
language plpgsql set search_path = '' as $$
begin
  new.updated_at := clock_timestamp();
  new.updated_by := (select auth.uid());
  if tg_op = 'UPDATE' then
    new.org_id := old.org_id;  -- a row can never move to another 園
    new.id := old.id;
    new.created_at := old.created_at;
  end if;
  return new;
end $$;

do $$
declare t text;
begin
  foreach t in array array['items', 'assets', 'logs', 'stocktakes', 'settings', 'photos'] loop
    execute format($f$
      create table public.%1$I (
        org_id uuid not null references public.orgs (id) on delete cascade,
        id text not null check (length(id) between 1 and 64),
        data jsonb not null default '{}'::jsonb check (jsonb_typeof(data) = 'object' and pg_column_size(data) < 262144),
        deleted boolean not null default false,
        created_at timestamptz not null default now(),
        updated_at timestamptz not null default clock_timestamp(),
        updated_by uuid,
        primary key (org_id, id)
      );
      create index %1$s_sync_idx on public.%1$I (org_id, updated_at);
      create trigger %1$s_touch before insert or update on public.%1$I for each row execute function hb_private.touch();
      alter table public.%1$I enable row level security;
      revoke all on public.%1$I from anon, authenticated;
      grant select, insert, update on public.%1$I to authenticated;
      create policy %1$s_select on public.%1$I for select to authenticated using (hb_private.is_member(org_id));
      create policy %1$s_insert on public.%1$I for insert to authenticated with check (hb_private.is_member(org_id));
      create policy %1$s_update on public.%1$I for update to authenticated using (hb_private.is_member(org_id)) with check (hb_private.is_member(org_id));
    $f$, t);
  end loop;
end $$;

-- exactly-once application of client batches (retries after a lost response must not double-count stock)
create table public.applied_batches (
  org_id uuid not null references public.orgs (id) on delete cascade,
  batch_id text not null check (length(batch_id) between 1 and 100),
  user_id uuid default auth.uid(),
  applied_at timestamptz not null default now(),
  primary key (org_id, batch_id)
);
alter table public.applied_batches enable row level security;
revoke all on public.applied_batches from anon, authenticated;
grant select, insert on public.applied_batches to authenticated;
create policy applied_select on public.applied_batches for select to authenticated using (hb_private.is_member(org_id));
create policy applied_insert on public.applied_batches for insert to authenticated with check (hb_private.is_member(org_id));

-- ---------------------------------------------------------------- merge rules
-- set: shallow field replace (fields not in the patch keep their server value)
--      except 'checks' / 'photoLabels', which merge one level deeper (union of stocktake checks)
-- inc: numeric delta (consumable stock), so two devices' in/out never cancel each other
-- arr: {add:[], rem:[]} on list fields (photoIds, master lists): server list minus removed, plus added
create function hb_private.merge(p_cur jsonb, p_set jsonb, p_inc jsonb, p_arr jsonb) returns jsonb
language plpgsql immutable set search_path = '' as $$
declare nd jsonb := coalesce(p_cur, '{}'::jsonb); k text; v jsonb; base numeric; cur_arr jsonb; res jsonb; e jsonb;
begin
  if jsonb_typeof(p_set) = 'object' then
    for k, v in select * from jsonb_each(p_set) loop
      if k in ('checks', 'photoLabels') and jsonb_typeof(v) = 'object' then
        nd := nd || jsonb_build_object(k, (case when jsonb_typeof(nd -> k) = 'object' then nd -> k else '{}'::jsonb end) || v);
      else
        nd := nd || jsonb_build_object(k, v);
      end if;
    end loop;
  end if;
  if jsonb_typeof(p_inc) = 'object' then
    for k, v in select * from jsonb_each(p_inc) loop
      if jsonb_typeof(v) <> 'number' then continue; end if;
      base := case
        when jsonb_typeof(nd -> k) = 'number' then (nd ->> k)::numeric
        when jsonb_typeof(nd -> k) = 'string' and (nd ->> k) ~ '^-?[0-9]+(\.[0-9]+)?$' then (nd ->> k)::numeric
        else 0 end;
      nd := nd || jsonb_build_object(k, base + (v #>> '{}')::numeric);
    end loop;
  end if;
  if jsonb_typeof(p_arr) = 'object' then
    for k, v in select * from jsonb_each(p_arr) loop
      if jsonb_typeof(v) <> 'object' then continue; end if;
      cur_arr := case when jsonb_typeof(nd -> k) = 'array' then nd -> k else '[]'::jsonb end;
      select coalesce(jsonb_agg(x.e order by x.ord), '[]'::jsonb) into res
        from jsonb_array_elements(cur_arr) with ordinality as x(e, ord)
        where not (case when jsonb_typeof(v -> 'rem') = 'array' then v -> 'rem' else '[]'::jsonb end) @> jsonb_build_array(x.e);
      if jsonb_typeof(v -> 'add') = 'array' then
        for e in select * from jsonb_array_elements(v -> 'add') loop
          if not res @> jsonb_build_array(e) then res := res || jsonb_build_array(e); end if;
        end loop;
      end if;
      nd := nd || jsonb_build_object(k, res);
    end loop;
  end if;
  return nd;
end $$;
revoke all on function hb_private.merge(jsonb, jsonb, jsonb, jsonb) from public, anon;
grant execute on function hb_private.merge(jsonb, jsonb, jsonb, jsonb) to authenticated;

-- ---------------------------------------------------------------- write API (security invoker: RLS applies)
create function public.hb_apply(p_org uuid, p_batch text, p_ops jsonb) returns jsonb
language plpgsql security invoker set search_path = '' as $$
declare
  op jsonb; t text; v_id text; v_kind text; cur jsonb; nd jsonb; n int := 0; v_found boolean;
begin
  if (select auth.uid()) is null or not hb_private.is_member(p_org) then
    raise exception 'この園のメンバーではありません' using errcode = '42501';
  end if;
  if p_batch is null or length(p_batch) not between 1 and 100 then raise exception 'bad batch id' using errcode = '22023'; end if;
  if jsonb_typeof(p_ops) <> 'array' or jsonb_array_length(p_ops) > 1000 then raise exception 'bad ops' using errcode = '22023'; end if;

  insert into public.applied_batches (org_id, batch_id) values (p_org, p_batch) on conflict do nothing;
  get diagnostics n = row_count;
  if n = 0 then return jsonb_build_object('ok', true, 'dup', true); end if;
  n := 0;

  for op in select * from jsonb_array_elements(p_ops) loop
    t := op ->> 't'; v_id := op ->> 'id'; v_kind := coalesce(op ->> 'k', 'patch');
    if t is null or t not in ('items', 'assets', 'logs', 'stocktakes', 'settings', 'photos') then
      raise exception 'bad table %', t using errcode = '22023';
    end if;
    if v_id is null or length(v_id) not between 1 and 64 then raise exception 'bad id' using errcode = '22023'; end if;
    if v_kind not in ('patch', 'delete') then raise exception 'bad kind %', v_kind using errcode = '22023'; end if;

    execute format('select data from public.%I where org_id = $1 and id = $2 for update', t) into cur using p_org, v_id;
    v_found := cur is not null;

    if v_kind = 'delete' then
      -- an item that another device has just given a live unit stays (its name would otherwise vanish)
      if v_found and not (t = 'items' and exists (
          select 1 from public.assets a where a.org_id = p_org and not a.deleted and a.data ->> 'itemId' = v_id)) then
        execute format('update public.%I set deleted = true where org_id = $1 and id = $2', t) using p_org, v_id;
      end if;
    else
      nd := hb_private.merge(cur, op -> 'set', op -> 'inc', op -> 'arr') - 'id';
      if v_found then
        -- a patch revives a tombstone: an edit made on another device is never silently dropped
        execute format('update public.%I set data = $3, deleted = false where org_id = $1 and id = $2', t) using p_org, v_id, nd;
      else
        execute format('insert into public.%I (org_id, id, data) values ($1, $2, $3)', t) using p_org, v_id, nd;
      end if;
      -- keep parents alive when a child is written after the parent was deleted elsewhere
      if t = 'assets' and nd ? 'itemId' then
        update public.items set deleted = false where org_id = p_org and id = nd ->> 'itemId' and deleted;
      elsif t = 'logs' then
        if nd ? 'assetId' then update public.assets set deleted = false where org_id = p_org and id = nd ->> 'assetId' and deleted; end if;
        if nd ? 'itemId' then update public.items set deleted = false where org_id = p_org and id = nd ->> 'itemId' and deleted; end if;
      end if;
    end if;
    n := n + 1;
  end loop;
  return jsonb_build_object('ok', true, 'n', n);
end $$;
revoke all on function public.hb_apply(uuid, text, jsonb) from public, anon;
grant execute on function public.hb_apply(uuid, text, jsonb) to authenticated;

-- capability declaration (multi-device-sync 原則7)
create function public.hb_version() returns jsonb language sql immutable set search_path = '' as $$
  select jsonb_build_object('app', 'hoikuen-bihin', 'proto', 1, 'merge', jsonb_build_array('set', 'inc', 'arr', 'checks'));
$$;
grant execute on function public.hb_version() to anon, authenticated;

create function public.hb_my_orgs() returns table (org_id uuid, name text, role text)
language sql stable security invoker set search_path = '' as $$
  select o.id, o.name, m.role from public.memberships m join public.orgs o on o.id = m.org_id
  where m.user_id = (select auth.uid()) order by o.name;
$$;
revoke all on function public.hb_my_orgs() from public, anon;
grant execute on function public.hb_my_orgs() to authenticated;

-- ---------------------------------------------------------------- invites (account issuance / password reset)
-- A 園 admin issues a one-time code for an email address. The code itself is never stored, only its sha256.
-- kind 'join'  : new account (via the hb-invite edge function) or an existing account joining another 園 (hb_join)
-- kind 'reset' : the member sets a new password (via the hb-invite edge function). No email delivery is needed.
create table hb_private.invites (
  id uuid primary key default gen_random_uuid(),
  org_id uuid not null references public.orgs (id) on delete cascade,
  email text not null check (length(email) between 3 and 254),
  role text not null check (role in ('admin', 'member')),
  kind text not null check (kind in ('join', 'reset')),
  code_hash text not null unique,
  created_by uuid,
  created_at timestamptz not null default now(),
  expires_at timestamptz not null,
  used_at timestamptz,
  used_by uuid,
  attempts int not null default 0
);
alter table hb_private.invites enable row level security;
revoke all on hb_private.invites from public, anon, authenticated;

create function hb_private.code_hash(p_code text) returns text
language sql immutable set search_path = '' as $$
  select encode(extensions.digest(upper(regexp_replace(coalesce(p_code, ''), '[^0-9A-Za-z]', '', 'g')), 'sha256'), 'hex');
$$;

create function hb_private.new_invite(p_org uuid, p_email text, p_role text, p_kind text, p_by uuid) returns text
language plpgsql security definer set search_path = '' as $$
declare alphabet text := 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; b bytea := extensions.gen_random_bytes(12); c text := ''; i int;
begin
  for i in 0..11 loop c := c || substr(alphabet, (get_byte(b, i) % 32) + 1, 1); end loop;
  -- an older unused code for the same person and purpose stops working
  update hb_private.invites set expires_at = now()
    where org_id = p_org and lower(email) = lower(p_email) and kind = p_kind and used_at is null and expires_at > now();
  insert into hb_private.invites (org_id, email, role, kind, code_hash, created_by, expires_at)
    values (p_org, lower(trim(p_email)), p_role, p_kind, hb_private.code_hash(c), p_by,
            now() + case when p_kind = 'reset' then interval '1 day' else interval '7 days' end);
  return substr(c, 1, 4) || '-' || substr(c, 5, 4) || '-' || substr(c, 9, 4);
end $$;
revoke all on function hb_private.new_invite(uuid, text, text, text, uuid) from public, anon, authenticated;

create function public.hb_issue_invite(p_org uuid, p_email text, p_role text default 'member', p_kind text default 'join')
returns jsonb language plpgsql security definer set search_path = '' as $$
declare v_uid uuid; v_target uuid; v_code text;
begin
  v_uid := (select auth.uid());
  if v_uid is null or not hb_private.is_admin(p_org) then raise exception '管理者だけが発行できます' using errcode = '42501'; end if;
  if p_email is null or p_email !~ '^[^@\s]+@[^@\s]+\.[^@\s]+$' then raise exception 'メールアドレスの形が正しくありません' using errcode = '22023'; end if;
  if p_role not in ('admin', 'member') or p_kind not in ('join', 'reset') then raise exception 'bad role/kind' using errcode = '22023'; end if;
  if p_kind = 'reset' then
    select u.id into v_target from auth.users u where lower(u.email) = lower(trim(p_email));
    if v_target is null or not exists (select 1 from public.memberships m where m.org_id = p_org and m.user_id = v_target) then
      raise exception 'この園のメンバーではないため、パスワードの再設定コードは出せません' using errcode = '22023';
    end if;
    -- an admin may only reset accounts that belong exclusively to 園 they administer
    if exists (select 1 from public.memberships m where m.user_id = v_target
               and not exists (select 1 from public.memberships a where a.org_id = m.org_id and a.user_id = v_uid and a.role = 'admin')) then
      raise exception 'ほかの園にも所属している人のため、ここでは再設定できません' using errcode = '42501';
    end if;
  end if;
  v_code := hb_private.new_invite(p_org, p_email, p_role, p_kind, v_uid);
  return jsonb_build_object('code', v_code, 'kind', p_kind, 'email', lower(trim(p_email)),
    'expires_at', now() + case when p_kind = 'reset' then interval '1 day' else interval '7 days' end);
end $$;
revoke all on function public.hb_issue_invite(uuid, text, text, text) from public, anon;
grant execute on function public.hb_issue_invite(uuid, text, text, text) to authenticated;

-- service_role only: validate a code without consuming it (the edge function creates the account, then completes)
create function public.hb_invite_peek(p_email text, p_code text) returns jsonb
language plpgsql security definer set search_path = '' as $$
declare inv hb_private.invites; v_user uuid;
begin
  select * into inv from hb_private.invites
    where code_hash = hb_private.code_hash(p_code) and lower(email) = lower(trim(p_email))
      and used_at is null and expires_at > now() and attempts < 10;
  if inv.id is null then
    update hb_private.invites set attempts = attempts + 1
      where lower(email) = lower(trim(p_email)) and used_at is null and expires_at > now();
    return jsonb_build_object('ok', false);
  end if;
  select u.id into v_user from auth.users u where lower(u.email) = lower(trim(p_email));
  return jsonb_build_object('ok', true, 'invite_id', inv.id, 'org_id', inv.org_id, 'role', inv.role, 'kind', inv.kind, 'user_id', v_user);
end $$;

create function public.hb_invite_complete(p_invite uuid, p_user uuid) returns jsonb
language plpgsql security definer set search_path = '' as $$
declare inv hb_private.invites;
begin
  update hb_private.invites set used_at = now(), used_by = p_user
    where id = p_invite and used_at is null and expires_at > now() returning * into inv;
  if inv.id is null then return jsonb_build_object('ok', false); end if;
  if inv.kind = 'join' then
    insert into public.memberships (org_id, user_id, role) values (inv.org_id, p_user, inv.role)
      on conflict (org_id, user_id) do nothing;
  end if;
  return jsonb_build_object('ok', true, 'org_id', inv.org_id);
end $$;
revoke all on function public.hb_invite_peek(text, text), public.hb_invite_complete(uuid, uuid) from public, anon, authenticated;
grant execute on function public.hb_invite_peek(text, text), public.hb_invite_complete(uuid, uuid) to service_role;

-- a signed-in user joins another 園 with a join code issued to their own email
create function public.hb_join(p_code text) returns jsonb
language plpgsql security definer set search_path = '' as $$
declare v_uid uuid := (select auth.uid()); v_email text; inv hb_private.invites;
begin
  if v_uid is null then raise exception 'ログインしてください' using errcode = '42501'; end if;
  select u.email into v_email from auth.users u where u.id = v_uid;
  update hb_private.invites set used_at = now(), used_by = v_uid
    where code_hash = hb_private.code_hash(p_code) and lower(email) = lower(v_email) and kind = 'join'
      and used_at is null and expires_at > now() and attempts < 10
    returning * into inv;
  if inv.id is null then
    update hb_private.invites set attempts = attempts + 1 where lower(email) = lower(v_email) and used_at is null and expires_at > now();
    raise exception 'コードが正しくないか、期限が切れています' using errcode = '22023';
  end if;
  insert into public.memberships (org_id, user_id, role) values (inv.org_id, v_uid, inv.role) on conflict (org_id, user_id) do nothing;
  return jsonb_build_object('ok', true, 'org_id', inv.org_id);
end $$;
revoke all on function public.hb_join(text) from public, anon;
grant execute on function public.hb_join(text) to authenticated;

create function public.hb_list_members(p_org uuid) returns table (user_id uuid, email text, role text, created_at timestamptz)
language plpgsql stable security definer set search_path = '' as $$
begin
  if not hb_private.is_admin(p_org) then raise exception '管理者だけが見られます' using errcode = '42501'; end if;
  return query select m.user_id, u.email::text, m.role, m.created_at
    from public.memberships m join auth.users u on u.id = m.user_id where m.org_id = p_org order by m.created_at;
end $$;
revoke all on function public.hb_list_members(uuid) from public, anon;
grant execute on function public.hb_list_members(uuid) to authenticated;

create function public.hb_set_member(p_org uuid, p_user uuid, p_role text) returns jsonb
language plpgsql security definer set search_path = '' as $$
begin
  if not hb_private.is_admin(p_org) then raise exception '管理者だけが変更できます' using errcode = '42501'; end if;
  if p_role is not null and p_role not in ('admin', 'member') then raise exception 'bad role' using errcode = '22023'; end if;
  -- never leave a 園 without an admin
  if (p_role is null or p_role = 'member')
     and exists (select 1 from public.memberships where org_id = p_org and user_id = p_user and role = 'admin')
     and (select count(*) from public.memberships where org_id = p_org and role = 'admin') <= 1 then
    raise exception '管理者が1人もいなくなるため変更できません' using errcode = '22023';
  end if;
  if p_role is null then
    delete from public.memberships where org_id = p_org and user_id = p_user;
  else
    update public.memberships set role = p_role where org_id = p_org and user_id = p_user;
  end if;
  return jsonb_build_object('ok', true);
end $$;
revoke all on function public.hb_set_member(uuid, uuid, text) from public, anon;
grant execute on function public.hb_set_member(uuid, uuid, text) to authenticated;

-- operator only (run as postgres from SQL): create a 園 and the first admin's join code
create function hb_private.create_org(p_name text, p_admin_email text) returns jsonb
language plpgsql security definer set search_path = '' as $$
declare v_org uuid;
begin
  insert into public.orgs (name) values (p_name) returning id into v_org;
  return jsonb_build_object('org_id', v_org, 'code', hb_private.new_invite(v_org, p_admin_email, 'admin', 'join', null));
end $$;
revoke all on function hb_private.create_org(text, text) from public, anon, authenticated;

-- ---------------------------------------------------------------- storage (private, one folder per 園)
insert into storage.buckets (id, name, public, file_size_limit, allowed_mime_types)
values ('hb-photos', 'hb-photos', false, 15728640, array['image/jpeg', 'image/png', 'image/webp', 'application/pdf'])
on conflict (id) do nothing;

create policy hb_photos_select on storage.objects for select to authenticated
  using (bucket_id = 'hb-photos' and hb_private.is_member_path((storage.foldername(name))[1]));
create policy hb_photos_insert on storage.objects for insert to authenticated
  with check (bucket_id = 'hb-photos' and hb_private.is_member_path((storage.foldername(name))[1]));
create policy hb_photos_update on storage.objects for update to authenticated
  using (bucket_id = 'hb-photos' and hb_private.is_member_path((storage.foldername(name))[1]))
  with check (bucket_id = 'hb-photos' and hb_private.is_member_path((storage.foldername(name))[1]));
