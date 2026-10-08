create schema if not exists pre;
create schema if not exists pro;

create table if not exists pro.personas (like public.personas including all);
create table if not exists pro.servicios (like public.servicios including all);
create table if not exists pro.asignaciones (like public.asignaciones including all);

insert into pro.personas select * from public.personas on conflict (id) do nothing;
insert into pro.servicios select * from public.servicios on conflict (id) do nothing;
insert into pro.asignaciones select * from public.asignaciones on conflict (id) do nothing;

create table if not exists pre.personas (like pro.personas including all);
create table if not exists pre.servicios (like pro.servicios including all);
create table if not exists pre.asignaciones (like pro.asignaciones including all);
insert into pre.personas select * from pro.personas on conflict (id) do nothing;
insert into pre.servicios select * from pro.servicios on conflict (id) do nothing;
insert into pre.asignaciones select * from pro.asignaciones on conflict (id) do nothing;

alter table pre.personas enable row level security;
alter table pre.servicios enable row level security;
alter table pre.asignaciones enable row level security;
alter table pro.personas enable row level security;
alter table pro.servicios enable row level security;
alter table pro.asignaciones enable row level security;

create policy pre_personas_public on pre.personas for all to anon using (true) with check (true);
create policy pre_servicios_public on pre.servicios for all to anon using (true) with check (true);
create policy pre_asignaciones_public on pre.asignaciones for all to anon using (true) with check (true);
create policy pro_personas_public on pro.personas for all to anon using (true) with check (true);
create policy pro_servicios_public on pro.servicios for all to anon using (true) with check (true);
create policy pro_asignaciones_public on pro.asignaciones for all to anon using (true) with check (true);

grant usage on schema pre, pro to anon, authenticated;
grant all on all tables in schema pre, pro to anon, authenticated;
