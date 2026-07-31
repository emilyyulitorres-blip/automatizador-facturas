# Automatizador de Facturas - Solar Team

Cambios de esta versión:
- El proyecto es opcional.
- El usuario puede escribir un proyecto nuevo directamente.
- Los proyectos nuevos se crean automáticamente en Supabase.
- Las facturas también pueden guardarse sin proyecto.
- El logo se muestra con ancho fijo para evitar recortes o deformaciones.
- Se mantiene el aprendizaje de Proveedor, Categoría y Consumo por RUC.

Antes de usar la creación automática de proyectos, ejecuta en Supabase:

```sql
grant insert on table public.proyectos to anon;

drop policy if exists "anon crea proyectos"
on public.proyectos;

create policy "anon crea proyectos"
on public.proyectos
for insert
to anon
with check (true);
```
