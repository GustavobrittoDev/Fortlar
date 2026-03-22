create table if not exists public.financial_entries (
  id uuid primary key default gen_random_uuid(),
  entry_type text not null check (entry_type in ('Receita', 'Despesa')),
  category text not null,
  description text not null,
  entry_date date not null,
  amount numeric(10, 2) not null default 0,
  status text not null default 'Pendente' check (status in ('Pendente', 'Pago')),
  payment_method text not null default 'Pix',
  reference text not null default '',
  created_at timestamptz not null default now()
);

alter table public.financial_entries enable row level security;

drop policy if exists "authenticated financial entries" on public.financial_entries;
create policy "authenticated financial entries" on public.financial_entries
for all to authenticated
using (true)
with check (true);
