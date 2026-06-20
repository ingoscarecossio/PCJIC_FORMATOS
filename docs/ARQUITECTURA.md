# Arquitectura

## Componentes

- `app.py`: aplicación Streamlit y lógica de negocio.
- `plantilla/`: formatos base FD-GC71 y FD-GC72, logos institucionales.
- `app_data/fdgc_app.sqlite3`: base local SQLite.
- `app_data/evidencias/`: repositorio de soportes por curso.

## Capas funcionales

1. Seguridad: usuarios, perfiles, contraseña con PBKDF2-HMAC-SHA256 y auditoría.
2. Expediente: cursos, estados, avance y propietario.
3. Planeación: unidades, horarios, sesiones, evaluación y estudiantes.
4. Generación documental: Word FD-GC71, Word FD-GC72, Excel de evaluación, ICS y ZIP.
5. Evidencias: carga, almacenamiento y descarga de soportes.
6. Gobierno: validación, auditoría y backups.

## Base de datos

Tablas principales:
- `usuarios`
- `auditoria`
- `cursos`
- `evidencias`
- `artefactos`

Los datos detallados de planeación se guardan en `cursos.payload_json` para permitir evolución del modelo sin migraciones pesadas.

## Seguridad

La autenticación es local. Para producción institucional puede integrarse después con LDAP, Microsoft Entra ID o Google Workspace.
