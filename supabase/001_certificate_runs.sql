create extension if not exists pgcrypto;

create table if not exists public.certificate_runs (
  id uuid primary key default gen_random_uuid(),
  created_at timestamptz not null default now(),
  client text not null check (client in ('liberty', 'stanlib', 'pm', 'alexforbes', 'sanlam')),
  certificate_count integer not null check (certificate_count >= 0),
  delivery_types text[] not null default '{}',
  source text not null default 'web',
  notes text
);

comment on table public.certificate_runs is 'Privacy-safe certificate generation metadata only. No participant names, ID numbers, emails, raw Excel data, or PDFs.';
comment on column public.certificate_runs.client is 'Selected client/template key.';
comment on column public.certificate_runs.certificate_count is 'Number of certificates generated in the run.';
comment on column public.certificate_runs.delivery_types is 'Selected delivery outputs such as merged_pdf, zip, email_reserved.';
comment on column public.certificate_runs.source is 'Origin of the event, e.g. web.';
comment on column public.certificate_runs.notes is 'Optional system note only. Do not store personal data here.';

alter table public.certificate_runs enable row level security;

drop policy if exists "deny_all_certificate_runs" on public.certificate_runs;
create policy "deny_all_certificate_runs"
on public.certificate_runs
for all
to public
using (false)
with check (false);
