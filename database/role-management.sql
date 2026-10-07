-- Canonical shared-role migration. Keep identical to Associates supabase/role_management_schema.sql.
-- Export existing rows first. Access and editing remain disabled until review and deployment complete.
create table if not exists public.app_role_assignments (
  email text primary key check (email = lower(trim(email))),
  roles text[] not null default '{}',
  updated_by text not null,
  updated_at timestamptz not null default now()
);
alter table public.app_role_assignments add column if not exists roles text[];
do $$
begin
  if exists (select 1 from information_schema.columns where table_schema='public' and table_name='app_role_assignments' and column_name='role') then
    -- Retain the legacy column as migration evidence; new writers need not supply it.
    execute 'alter table public.app_role_assignments alter column role drop not null';
    execute $conversion$
      update public.app_role_assignments set roles = case role
        when 'superadmin' then array['superadmin']
        when 'admin' then array['manager']
        when 'finance' then array['financeManager']
        when 'operations' then array['careCoordinator']
        when 'care_manager' then array['careCoordinator']
        when 'consultant' then array['consultant']
        when 'photo_layout' then array['consultant']
        when 'enquiries_only' then array['consultant']
        when 'marketing' then array['marketing']
        when 'hr_only' then array['hr']
        when 'logged_in' then array['care']
        else '{}'::text[] end
      where roles is null
    $conversion$;
    -- Unmapped legacy roles are retained in role. Review them before activating access.
  end if;
end $$;
update public.app_role_assignments set roles = '{}' where roles is null;
alter table public.app_role_assignments alter column roles set default '{}';
alter table public.app_role_assignments alter column roles set not null;
do $$
declare item record;
begin
  for item in select conname from pg_constraint where conrelid='public.app_role_assignments'::regclass
    and contype='c' and pg_get_constraintdef(oid) ~ '\mroles\M'
  loop execute format('alter table public.app_role_assignments drop constraint %I', item.conname); end loop;
end $$;
alter table public.app_role_assignments add constraint app_role_assignments_shared_roles_check
  check (roles <@ array['superadmin','manager','financeManager','consultant','careCoordinator','care','marketing','hr','careAdmin']::text[]
    and array_position(roles, null) is null);

create table if not exists public.app_role_management_state (
  id text primary key check (id='shared'),
  access_enabled boolean not null default false,
  writes_enabled boolean not null default false,
  check (not writes_enabled or access_enabled)
);
create table if not exists public.app_role_aliases (
  alias_email text primary key check (alias_email = lower(trim(alias_email))),
  canonical_email text not null references public.app_role_assignments(email),
  check (alias_email <> canonical_email)
);
insert into public.app_role_management_state(id) values ('shared') on conflict do nothing;
alter table public.app_role_assignments enable row level security;
alter table public.app_role_management_state enable row level security;
alter table public.app_role_aliases enable row level security;
revoke all on public.app_role_assignments, public.app_role_management_state, public.app_role_aliases from anon, authenticated;
grant select, insert, update on public.app_role_assignments to service_role;
grant select, update on public.app_role_management_state to service_role;
grant select, insert, update on public.app_role_aliases to service_role;
