create extension if not exists "pgcrypto";

create table if not exists public.profiles (
  id uuid primary key references auth.users (id) on delete cascade,
  name text not null,
  role text not null default 'Administrador',
  created_at timestamptz not null default now()
);

create table if not exists public.leads (
  id uuid primary key default gen_random_uuid(),
  name text not null,
  phone text not null,
  service text not null,
  address text not null,
  source text not null default 'Site',
  priority text not null default 'Media',
  status text not null default 'novo',
  notes text not null default '',
  created_at timestamptz not null default now(),
  last_contact text not null default 'Agora'
);

create table if not exists public.appointments (
  id uuid primary key default gen_random_uuid(),
  customer text not null,
  phone text not null default '',
  service text not null,
  address text not null,
  date date not null,
  time text not null,
  notes text not null default '',
  created_at timestamptz not null default now()
);

create table if not exists public.orders (
  id uuid primary key default gen_random_uuid(),
  code text unique not null default 'OS-' || floor(1000 + random() * 9000)::text,
  customer text not null,
  phone text not null default '',
  service text not null,
  address text not null,
  date date not null,
  time text not null,
  amount numeric(10, 2) not null default 0,
  status text not null default 'Agendado',
  payment_status text not null default 'Pendente',
  notes text not null default '',
  created_at timestamptz not null default now()
);

create table if not exists public.attachments (
  id uuid primary key default gen_random_uuid(),
  order_id uuid not null references public.orders (id) on delete cascade,
  original_name text not null,
  storage_path text not null,
  mime_type text not null,
  file_size integer not null,
  created_at timestamptz not null default now()
);

create table if not exists public.activities (
  id uuid primary key default gen_random_uuid(),
  title text not null,
  description text not null,
  time_label text not null default 'Agora',
  created_at timestamptz not null default now()
);

alter table public.profiles enable row level security;
alter table public.leads enable row level security;
alter table public.appointments enable row level security;
alter table public.orders enable row level security;
alter table public.attachments enable row level security;
alter table public.activities enable row level security;

drop policy if exists "authenticated profiles" on public.profiles;
create policy "authenticated profiles" on public.profiles
for all to authenticated
using (true)
with check (true);

drop policy if exists "authenticated leads" on public.leads;
create policy "authenticated leads" on public.leads
for all to authenticated
using (true)
with check (true);

drop policy if exists "public site lead insert" on public.leads;
create policy "public site lead insert" on public.leads
for insert to anon
with check (
  source = 'Site'
  and status = 'novo'
  and priority in ('Alta', 'Media', 'Baixa')
);

drop policy if exists "authenticated appointments" on public.appointments;
create policy "authenticated appointments" on public.appointments
for all to authenticated
using (true)
with check (true);

drop policy if exists "authenticated orders" on public.orders;
create policy "authenticated orders" on public.orders
for all to authenticated
using (true)
with check (true);

drop policy if exists "authenticated attachments" on public.attachments;
create policy "authenticated attachments" on public.attachments
for all to authenticated
using (true)
with check (true);

drop policy if exists "authenticated activities" on public.activities;
create policy "authenticated activities" on public.activities
for all to authenticated
using (true)
with check (true);

insert into storage.buckets (id, name, public)
values ('crm-anexos', 'crm-anexos', false)
on conflict (id) do nothing;

drop policy if exists "authenticated upload attachments" on storage.objects;
create policy "authenticated upload attachments" on storage.objects
for insert to authenticated
with check (bucket_id = 'crm-anexos');

drop policy if exists "authenticated read attachments" on storage.objects;
create policy "authenticated read attachments" on storage.objects
for select to authenticated
using (bucket_id = 'crm-anexos');

drop policy if exists "authenticated update attachments" on storage.objects;
create policy "authenticated update attachments" on storage.objects
for update to authenticated
using (bucket_id = 'crm-anexos')
with check (bucket_id = 'crm-anexos');

drop policy if exists "authenticated delete attachments" on storage.objects;
create policy "authenticated delete attachments" on storage.objects
for delete to authenticated
using (bucket_id = 'crm-anexos');

create or replace function public.handle_new_user()
returns trigger
language plpgsql
security definer
as $$
begin
  insert into public.profiles (id, name, role)
  values (
    new.id,
    coalesce(new.raw_user_meta_data->>'name', split_part(new.email, '@', 1)),
    'Administrador'
  )
  on conflict (id) do nothing;

  return new;
end;
$$;

drop trigger if exists on_auth_user_created on auth.users;
create trigger on_auth_user_created
after insert on auth.users
for each row execute procedure public.handle_new_user();
