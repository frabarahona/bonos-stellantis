-- Da acceso al tablero al usuario admin (Franco). Antes, crear el usuario en
-- Supabase → Authentication → Users → Add user → Create new user (marcar "Auto Confirm User").
do $$
declare col text;
begin
  select column_name into col from information_schema.columns
   where table_schema = 'public' and table_name = 'usuarios' and column_name in ('email', 'correo') limit 1;
  if col is null then raise exception 'No encontré la columna de correo en public.usuarios'; end if;
  execute format('insert into public.usuarios (%I) select $1 where not exists (select 1 from public.usuarios where lower(%I) = lower($1))', col, col)
  using 'fra.barahona@gmail.com';
end $$;
