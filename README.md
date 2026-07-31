# Automatizador de Facturas v4

Versión final modular.

## Archivos que debes subir a GitHub

- `app.py`
- carpeta `modules`
- `logo_corregido.png` o `logo.png`
- `requirements.txt`
- `packages.txt`

## Streamlit Secrets

```toml
SUPABASE_URL = "..."
SUPABASE_KEY = "..."
```

## Importante

El encabezado usa componentes nativos de Streamlit (`st.columns`, `st.image`,
`st.title` y `st.caption`). Por eso el HTML ya no puede aparecer como texto.
