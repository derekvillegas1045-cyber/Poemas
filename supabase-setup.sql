-- ============================================================
-- Configuración de Supabase para la Colección de Poemas
-- Pega y ejecuta TODO este script en:
-- Supabase Dashboard → SQL Editor → New query → Run
-- ============================================================

-- 1. Tabla de poemas
create table if not exists public.poems (
    id          uuid primary key default gen_random_uuid(),
    created_at  timestamptz not null default now(),
    title       text not null check (char_length(title) between 1 and 80),
    dedicatory  text check (dedicatory is null or char_length(dedicatory) <= 120),
    author      text not null default 'Anónimo' check (char_length(author) <= 60),
    theme       text not null default 'rosa'
                check (theme in ('rosa', 'medianoche', 'atardecer', 'esmeralda', 'sepia')),
    lines       jsonb not null default '[]'::jsonb,
    badge       text check (badge is null or char_length(badge) <= 4),
    columns     boolean not null default false,
    user_id     uuid references auth.users (id) on delete set null
);

-- 2. Seguridad a nivel de fila (RLS)
alter table public.poems enable row level security;

-- Cualquier visitante (anon incluido) puede LEER los poemas
drop policy if exists "Poemas visibles para todos" on public.poems;
create policy "Poemas visibles para todos"
    on public.poems
    for select
    using (true);

-- Solo usuarios autenticados pueden PUBLICAR (y el user_id se valida solo)
drop policy if exists "Usuarios autenticados pueden publicar" on public.poems;
create policy "Usuarios autenticados pueden publicar"
    on public.poems
    for insert
    to authenticated
    with check (user_id = auth.uid());

-- Solo el autor puede EDITAR sus poemas
drop policy if exists "Autor puede editar sus poemas" on public.poems;
create policy "Autor puede editar sus poemas"
    on public.poems
    for update
    to authenticated
    using (user_id = auth.uid())
    with check (user_id = auth.uid());

-- Solo el autor puede BORRAR sus poemas
drop policy if exists "Autor puede borrar sus poemas" on public.poems;
create policy "Autor puede borrar sus poemas"
    on public.poems
    for delete
    to authenticated
    using (user_id = auth.uid());

-- 3. Índice para listar por fecha
create index if not exists poems_created_at_idx on public.poems (created_at asc);

-- Listo. Ahora copia tu Project URL y anon key en index.html
-- (constantes SUPABASE_URL y SUPABASE_ANON_KEY).
