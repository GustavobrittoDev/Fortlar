drop policy if exists "public site lead insert" on public.leads;

create policy "public site lead insert" on public.leads
for insert to anon
with check (
  source = 'Site'
  and status = 'novo'
  and priority in ('Alta', 'Media', 'Baixa')
);
