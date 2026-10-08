create extension if not exists pgcrypto;

create table if not exists public.personas (
  id text primary key,
  nombre text not null,
  apellidos text not null,
  email text,
  tipo text not null check (tipo in ('SEGITTUR','INECO','Externo')),
  area text check (area is null or area in ('RRII','Desarrollo de Negocio','IDI','Comunicación')),
  estado text not null default 'Activo' check (estado in ('Activo','Inactivo')),
  notas text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
create table if not exists public.servicios (
  id text primary key,
  codigo_servicio text not null unique,
  nombre text not null,
  descripcion text,
  url_pre text,
  url_pro text,
  estado text not null default 'Activo' check (estado in ('Activo','Inactivo')),
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
create table if not exists public.asignaciones (
  id text primary key,
  persona_id text not null references public.personas(id) on delete cascade,
  servicio_id text not null references public.servicios(id) on delete cascade,
  rol text not null check (rol in ('Principal','Corresponsable')),
  estado text not null default 'Validado' check (estado in ('Validado','No Validado')),
  notas text,
  created_at timestamptz not null default now(),
  updated_at timestamptz not null default now()
);
create index if not exists asignaciones_persona_idx on public.asignaciones(persona_id);
create index if not exists asignaciones_servicio_idx on public.asignaciones(servicio_id);
alter table public.personas enable row level security;
alter table public.servicios enable row level security;
alter table public.asignaciones enable row level security;
drop policy if exists personas_public_read_write on public.personas;
drop policy if exists servicios_public_read_write on public.servicios;
drop policy if exists asignaciones_public_read_write on public.asignaciones;
create policy personas_public_read_write on public.personas for all to anon using (true) with check (true);
create policy servicios_public_read_write on public.servicios for all to anon using (true) with check (true);
create policy asignaciones_public_read_write on public.asignaciones for all to anon using (true) with check (true);
