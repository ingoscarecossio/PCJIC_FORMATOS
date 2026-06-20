# Despliegue

## Local

Use SQLite automáticamente si no configura `DATABASE_URL`.

```powershell
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Cloud

Use PostgreSQL externo y configure secretos desde el panel de Streamlit Cloud. Ver `DESPLIEGUE_STREAMLIT_CLOUD.md`.

## Docker institucional

```bash
docker compose up --build -d
```

El `docker-compose.yml` incluido levanta PostgreSQL y la aplicación.
