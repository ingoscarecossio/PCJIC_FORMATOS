# Hotfix Streamlit Cloud 7.0.1

Este ajuste evita que la aplicación se quede indefinidamente en **Your app is in the oven** cuando la base de datos externa está mal configurada.

## Cambios

- `connect_timeout` explícito para PostgreSQL/Supabase.
- Detección de `DATABASE_URL` con valores de ejemplo como `USUARIO`, `CLAVE`, `HOST` o `BASE`.
- Pantalla de error controlado si la base de datos falla al iniciar.
- Mensaje de diagnóstico para corregir Secrets en Streamlit Cloud.

## Secrets mínimos recomendados

```toml
APP_ENV = "production"
DATABASE_URL = "postgresql://USUARIO_REAL:CLAVE_REAL@HOST_REAL:5432/BASE_REAL"
DB_CONNECT_TIMEOUT = 8

INITIAL_ADMIN_USER = "admin"
INITIAL_ADMIN_PASSWORD = "Cambiar_Esta_Clave_123*"
INITIAL_ADMIN_NAME = "Administrador del sistema"
INITIAL_ADMIN_EMAIL = "admin@institucion.edu.co"
MAX_EVIDENCE_MB = 15
```

Si está usando Supabase, normalmente conviene usar la cadena del **Session pooler** o **Transaction pooler** entregada por Supabase, no inventarla manualmente.
