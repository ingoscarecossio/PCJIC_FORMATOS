# Despliegue en Streamlit Community Cloud

## 1. Preparar repositorio

Suba el contenido del proyecto a GitHub. No suba:

- `.streamlit/secrets.toml`
- `app_data/`
- `salida_demo/`
- bases `.sqlite3`

El `.gitignore` ya bloquea esos elementos.

## 2. Crear base de datos

Use PostgreSQL administrado. Supabase funciona bien para este caso.

Guarde la cadena de conexión en `DATABASE_URL`.

## 3. Configurar secretos en Streamlit

En la app de Streamlit Cloud:

```toml
APP_ENV = "production"
DATABASE_URL = "postgresql://USUARIO:CLAVE@HOST:5432/BASE"
INITIAL_ADMIN_USER = "admin"
INITIAL_ADMIN_PASSWORD = "Cambiar_Esta_Clave_123*"
INITIAL_ADMIN_NAME = "Administrador del sistema"
INITIAL_ADMIN_EMAIL = "admin@institucion.edu.co"
MAX_EVIDENCE_MB = 15
```

## 4. Primer ingreso

Al primer arranque, la app crea las tablas y el usuario administrador inicial. Después del ingreso:

1. Cambie la contraseña.
2. Cree usuarios nominales.
3. No use el administrador inicial para operación diaria.

## 5. Evidencias

Las evidencias se guardan en base de datos en `contenido_b64` y también se cachean localmente si el servidor lo permite. Para Streamlit Cloud, la fuente confiable será la base de datos.

## 6. Respaldo

Desde **Copias y restauración**, descargue un ZIP con:

- CSV de tablas principales.
- Evidencias reconstruidas desde base de datos.
- Manifiesto técnico.

## 7. Checklist antes de publicar

- `APP_ENV = "production"`.
- `DATABASE_URL` configurado.
- Contraseña inicial fuerte.
- Prueba de login exitosa.
- Generación de FD-GC71 correcta.
- Generación de plantilla Excel correcta.
- Carga y descarga de evidencia probada.
- Descarga de respaldo probada.
