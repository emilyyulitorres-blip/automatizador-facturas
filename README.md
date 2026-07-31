# Automatizador de Facturas - Solar Team

Versión con:
- Extracción de PDF e imágenes.
- Excel profesional.
- Proyectos cargados desde Supabase.
- Guardado compartido de facturas.
- Detección de duplicados por RUC + número de factura.
- Memoria de Proveedor, Categoría y Consumo por RUC.

Antes de publicar:
1. Agrega `supabase` a requirements.txt.
2. Configura `SUPABASE_URL` y `SUPABASE_KEY` en Streamlit Secrets.
3. Permite SELECT en `proyectos` y `correcciones`, e INSERT en `facturas` y `correcciones`.
