insert into public.leads (name, phone, service, address, source, priority, status, notes, last_contact)
values
  ('Mariana Souza', '(12) 99111-2233', 'Hidraulica', 'Vila Branca, Jacarei', 'Site', 'Alta', 'novo', 'Vazamento na pia da cozinha e torneira pingando.', 'Hoje, 09:10'),
  ('Rogerio Lima', '(12) 99671-9090', 'Montagem de moveis', 'Jardim California, Jacarei', 'Indicacao', 'Media', 'orcamento', 'Montagem de painel com TV e fixacao de prateleira.', 'Ontem, 16:45'),
  ('Patricia Alves', '(12) 99881-4545', 'Eletrica', 'Centro, Jacarei', 'Instagram', 'Alta', 'negociacao', 'Instalacao de spots e revisao de tomada que nao funciona.', 'Hoje, 08:20'),
  ('Carlos Mota', '(12) 99700-1188', 'Reparos em geral', 'Parque Meia Lua, Jacarei', 'WhatsApp', 'Baixa', 'fechado', 'Ajuste de porta, prateleira e suporte de cortina.', '21/03, 13:00');

insert into public.appointments (customer, phone, service, address, date, time, notes)
values
  ('Mariana Souza', '(12) 99111-2233', 'Hidraulica', 'Vila Branca, Jacarei', '2026-03-22', '10:30', ''),
  ('Patricia Alves', '(12) 99881-4545', 'Eletrica', 'Centro, Jacarei', '2026-03-22', '14:00', ''),
  ('Rogerio Lima', '(12) 99671-9090', 'Montagem de moveis', 'Jardim California, Jacarei', '2026-03-23', '09:00', ''),
  ('Bianca Nunes', '(12) 99711-7788', 'Instalacoes', 'Cidade Jardim, Jacarei', '2026-03-24', '16:30', '');

insert into public.orders (code, customer, phone, service, address, date, time, amount, status, payment_status, notes)
values
  ('OS-301', 'Rogerio Lima', '(12) 99671-9090', 'Montagem de moveis', 'Jardim California, Jacarei', '2026-03-22', '14:00', 580, 'Agendado', 'Pendente', 'Montagem de painel e fixacao de prateleira.'),
  ('OS-302', 'Patricia Alves', '(12) 99881-4545', 'Eletrica', 'Centro, Jacarei', '2026-03-22', '16:00', 1250, 'Em andamento', 'Parcial', 'Instalacao de spots e revisao eletrica.'),
  ('OS-303', 'Carlos Mota', '(12) 99700-1188', 'Reparos em geral', 'Parque Meia Lua, Jacarei', '2026-03-21', '11:00', 390, 'Concluido', 'Pago', 'Ajustes finais em porta e suportes.');

insert into public.activities (title, description, time_label)
values
  ('Lead recebido pelo site', 'Mariana Souza solicitou avaliacao hidraulica.', 'Hoje, 09:10'),
  ('Orcamento enviado', 'Rogerio Lima recebeu proposta de montagem.', 'Ontem, 16:45'),
  ('OS em andamento', 'Patricia Alves esta com eletrica em execucao.', 'Hoje, 08:20');
