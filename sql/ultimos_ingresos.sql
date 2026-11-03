-- Últimos inicios de sesión del tablero (desde el registro de Supabase Auth).
-- Correr una vez en Supabase → SQL Editor.
create or replace function public.ultimos_ingresos(n integer default 5)
returns table (email text, at timestamptz)
language plpgsql stable security definer
set search_path = public, auth
as $$
begin
  if not public.es_usuario_permitido() then return; end if;
  if exists (select 1 from auth.audit_log_entries where payload ->> 'action' = 'login') then
    return query
      select (e.payload ->> 'actor_username')::text, e.created_at
      from auth.audit_log_entries e
      where e.payload ->> 'action' = 'login'
      order by e.created_at desc
      limit least(coalesce(n, 5), 50);
  else
    return query
      select u.email::text, u.last_sign_in_at
      from auth.users u
      where u.last_sign_in_at is not null
      order by u.last_sign_in_at desc
      limit least(coalesce(n, 5), 50);
  end if;
end $$;
revoke all on function public.ultimos_ingresos(integer) from public, anon;
grant execute on function public.ultimos_ingresos(integer) to authenticated;
