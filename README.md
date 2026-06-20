# Gestor Académico FD-GC71 / FD-GC72 — Cloud Production

Aplicación en **Streamlit** para planear cursos, generar la **Guía Didáctica FD-GC71**, construir plantillas de evaluación, alimentar el **Informe Académico FD-GC72**, controlar usuarios/perfiles, registrar evidencias y mantener auditoría.

## Qué trae esta versión productiva

- Compatible con **Streamlit Community Cloud**.
- Modo local con **SQLite** para desarrollo.
- Modo productivo con **PostgreSQL/Supabase** mediante `DATABASE_URL`.
- Login con perfiles: Administrador, Coordinador, Docente y Consulta.
- Contraseñas con PBKDF2-HMAC-SHA256, salt individual y política reforzada en producción.
- Bloqueo temporal por intentos fallidos.
- Auditoría de ingresos y acciones sensibles.
- Evidencias persistidas en base de datos como respaldo, con caché local opcional.
- Descarga de respaldo completo en ZIP: tablas, evidencias y manifiesto.
- `.gitignore`, `runtime.txt`, `secrets.example.toml` y script SQL opcional para PostgreSQL.

## Estructura recomendada del repositorio

```text
.
├── app.py
├── requirements.txt
├── runtime.txt
├── .gitignore
├── .streamlit/
│   ├── config.toml
│   └── secrets.example.toml
├── plantilla/
│   ├── FD-GC71.docx
│   ├── FD-GC72-Informe_Academico.docx
│   ├── logo_poli.png
│   └── logo_icontec.png
├── scripts/
│   ├── init_postgres.sql
│   ├── smoke_test.py
│   └── run_windows.ps1
└── docs/
```

## Ejecución local

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

En local, si no configura `DATABASE_URL`, se usa SQLite en `app_data/fdgc_app.sqlite3`.

## Despliegue en Streamlit Cloud

1. Suba este proyecto a GitHub.
2. Cree una base PostgreSQL, por ejemplo en Supabase.
3. En Streamlit Cloud, cree una app apuntando a `app.py`.
4. Configure los secretos desde **App settings > Secrets** usando el contenido base de `.streamlit/secrets.example.toml`.
5. Cambie `INITIAL_ADMIN_PASSWORD` por una contraseña fuerte.
6. Inicie la app, entre con el usuario inicial y cambie la contraseña.
7. Cree usuarios nominales para docentes/coordinadores.

## Secretos mínimos para producción

```toml
APP_ENV = "production"
DATABASE_URL = "postgresql://USUARIO:CLAVE@HOST:5432/BASE"
INITIAL_ADMIN_USER = "admin"
INITIAL_ADMIN_PASSWORD = "Cambiar_Esta_Clave_123*"
INITIAL_ADMIN_NAME = "Administrador del sistema"
INITIAL_ADMIN_EMAIL = "admin@institucion.edu.co"
MAX_EVIDENCE_MB = 15
```

## Verificación rápida

```powershell
python scripts\smoke_test.py
```

El smoke test genera FD-GC71, Excel de evaluación, calendario ICS y paquete ZIP de prueba.

## Recomendación institucional

Para operación real, use PostgreSQL externo. SQLite queda como modo local o demostración. En Streamlit Cloud el sistema de archivos no debe tratarse como archivo institucional permanente; por eso esta versión guarda la evidencia también en base de datos.
