# 7.0.1 - Hotfix Streamlit Cloud

- Se agregó timeout explícito a conexiones PostgreSQL/Supabase.
- Se detectan valores de ejemplo en `DATABASE_URL` para no bloquear el arranque.
- Se agregó pantalla de diagnóstico cuando `init_db()` falla.
- Se documentó el problema `Your app is in the oven`.

# CHANGELOG

## 6.0.0-suite-inteligente

- Nuevo cargador inteligente de Excel con detección de encabezados y mapeo asistido.
- Nuevo comparador de cortes: listado inicial vs parcial vs final.
- Nuevo semáforo del expediente por curso.
- Nueva exportación masiva institucional con ZIP maestro e índice CSV.
- Nuevo módulo de hash SHA-256 y QR documental.
- Nuevo asistente académico de redacción editable para FD-GC71 y FD-GC72.
- Menú, permisos y pantalla de inicio ajustados para flujo inteligente.

# Changelog

## 3.0 Enterprise
- Agrega expediente académico por curso.
- Agrega centro de control con métricas.
- Agrega planeador superior con validación, DOCX, Excel, ICS, CSV, JSON y ZIP.
- Agrega evidencias y soportes por curso.
- Agrega validador institucional.
- Agrega respaldos completos.
- Amplía perfiles y permisos.

## 2.0 Login y perfiles
- Login local con SQLite.
- Usuarios, perfiles, auditoría y cambio de contraseña.

## 1.0 Base
- FD-GC71 y FD-GC72 con Streamlit.


## 5.0.0-ux-premium

### Diseño y experiencia de usuario
- Nuevo sistema visual premium para Streamlit: fondo institucional, tarjetas, botones, métricas, pestañas y contenedores con mejor jerarquía visual.
- Nueva pantalla de login con mensaje institucional y orientación de primera instalación.
- Sidebar rediseñado con búsqueda, iconos, descripciones, progreso operativo y cierre de sesión claro.
- Panel de inicio convertido en centro de mando con accesos rápidos, alertas y estado técnico.
- Hero contextual por módulo y ruta visual del expediente académico.
- Centro de control rediseñado con filtros, gráficos y acciones sugeridas.
- Reportes ejecutivos con experiencia de filtrado y descarga más clara.

### Compatibilidad
- Mantiene generación de FD-GC71, FD-GC72, plantillas Excel, calendario ICS, ZIP de expediente, auditoría, aprobaciones, versionamiento y despliegue Streamlit Cloud.


## 7.0.0-institucional

- Agregado banco institucional de asignaturas.
- Agregado motor de coherencia académica con score, hallazgos, checklist y matriz RA-contenido-evaluación.
- Agregado flujo de aprobación bloqueante con bloqueo/desbloqueo auditado.
- Agregado informe ejecutivo institucional descargable en Excel y Word.
- Agregada exportación institucional estructurada en ZIP maestro.
- Agregada auditoría por expediente.
- Ampliados permisos por rol y navegación UX.
