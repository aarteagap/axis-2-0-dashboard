-- COMPASS — Control de Presupuesto de Transporte
-- Fase 2: esquema de tarifas vigentes + historial de versiones.
-- Ejecutar completo en: Supabase → SQL Editor → New query → Run.

create table if not exists public.tarifa_overrides (
  route_key        text primary key,          -- "PROVEEDOR|||RUTA"
  proveedor        text not null,
  ruta             text not null,
  tarifa_actual    numeric not null,
  pct_adj_nueva    numeric not null,
  fecha_inicio     date not null,
  fecha_vencimiento date not null,
  updated_at       timestamptz not null default now()
);

create table if not exists public.tarifa_overrides_historial (
  id               bigint generated always as identity primary key,
  route_key        text not null,
  proveedor        text not null,
  ruta             text not null,
  tarifa_nueva     numeric not null,
  pct_adj_nueva    numeric not null,
  fecha_inicio     date not null,
  fecha_vencimiento date not null,
  creado_en        timestamptz not null default now()
);

create index if not exists idx_historial_route_key
  on public.tarifa_overrides_historial (route_key, creado_en);

-- RLS: plataforma interna sin login (solo un par de usuarios del área).
-- Se deja lectura/escritura abierta a la anon key, igual que el resto
-- del prototipo. Si más adelante quieres restringir quién escribe,
-- aquí es donde se agregaría una policy más estricta.
alter table public.tarifa_overrides enable row level security;
alter table public.tarifa_overrides_historial enable row level security;

create policy "anon select overrides" on public.tarifa_overrides
  for select using (true);
create policy "anon upsert overrides" on public.tarifa_overrides
  for insert with check (true);
create policy "anon update overrides" on public.tarifa_overrides
  for update using (true);

create policy "anon select historial" on public.tarifa_overrides_historial
  for select using (true);
create policy "anon insert historial" on public.tarifa_overrides_historial
  for insert with check (true);

-- Habilita Realtime para que el Dashboard reciba cambios de tarifa
-- al instante sin recargar (opcional, ya usamos postMessage como
-- respaldo dentro de la misma sesión del navegador).
alter publication supabase_realtime add table public.tarifa_overrides;
