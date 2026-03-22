alter table public.orders
add column if not exists amount_paid numeric(10, 2) not null default 0;

update public.orders
set amount_paid = amount
where payment_status = 'Pago'
  and coalesce(amount_paid, 0) = 0;

update public.orders
set amount_paid = 0
where payment_status = 'Pendente'
  and coalesce(amount_paid, 0) <> 0;
