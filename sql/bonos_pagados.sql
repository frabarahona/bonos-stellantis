-- Tabla de bonos pagados por la marca (cartas de aporte de Stellantis, una fila por VIN)
create table if not exists public.bonos_pagados (
  id bigint generated always as identity primary key,
  folio text not null,
  marca text,
  titulo text,
  fecha date,
  semana integer,
  archivo text,
  vin text not null,
  modelo text,
  version text,
  mes_inmat text,
  anio_inmat text,
  canal text,
  aporte numeric,
  diferencia numeric default 0,
  estado text,
  carga_id bigint,
  cargado_por text,
  cargado_at timestamptz default now(),
  unique (folio, vin)
);
create index if not exists bonos_pagados_vin_idx on public.bonos_pagados (vin);
alter table public.bonos_pagados enable row level security;
drop policy if exists bonos_pagados_all on public.bonos_pagados;
create policy bonos_pagados_all on public.bonos_pagados for all to authenticated
  using (public.es_usuario_permitido()) with check (public.es_usuario_permitido());
revoke all on public.bonos_pagados from anon;
