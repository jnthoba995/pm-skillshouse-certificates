create table if not exists public.admin_users (
  id uuid primary key references auth.users(id) on delete cascade,
  email text not null unique,
  full_name text,
  role text not null check (role in ('super_admin', 'staff')),
  is_active boolean not null default true,
  invited_by uuid references auth.users(id),
  created_at timestamptz not null default now()
);

comment on table public.admin_users is 'Closed-access admin/staff users for PM Skillshouse certificate system.';
comment on column public.admin_users.role is 'super_admin = owner/admin, staff = invited internal user.';
comment on column public.admin_users.is_active is 'Allows admin access to be disabled without deleting auth user.';

alter table public.admin_users enable row level security;

drop policy if exists "deny_all_admin_users" on public.admin_users;
create policy "deny_all_admin_users"
on public.admin_users
for all
to public
using (false)
with check (false);
