# Changelog

## 4.0.0-cloud-production

- Optimización para despliegue en Streamlit Community Cloud.
- Soporte productivo para PostgreSQL/Supabase mediante `DATABASE_URL`.
- Modo local con SQLite como fallback de desarrollo.
- Credenciales iniciales parametrizadas por `st.secrets` o variables de entorno.
- Política de contraseñas reforzada en `APP_ENV=production`.
- Bloqueo temporal por intentos fallidos de login.
- Evidencias guardadas también en base de datos para no depender del sistema de archivos efímero.
- Respaldo ZIP mejorado con tablas CSV, evidencias y manifiesto.
- Diagnóstico productivo dentro de la app.
- `.gitignore`, `runtime.txt`, `secrets.example.toml` y `scripts/init_postgres.sql`.
- Dockerfile con healthcheck y docker-compose con PostgreSQL.

## 3.0.0-enterprise

- Expediente académico por curso.
- Centro de control académico.
- Planeador superior.
- Evidencias y soportes.
- Validador institucional.
- Copias de seguridad.
