# Gestor académico FD-GC71 / FD-GC72 con login y perfiles

Aplicación en Streamlit para gestionar el flujo completo de planeación, evaluación e informe académico:

- **FD-GC71**: guía didáctica, concertación de evaluación, módulos/unidades, intensidad horaria, calendario automático de sesiones y plantilla Excel de evaluación.
- **FD-GC72**: informe académico alimentado desde listado tradicional de clase y plantillas de evaluación de mitad/final del curso.
- **Login local** con base SQLite.
- **Perfiles de usuario**: Administrador, Coordinador, Docente y Consulta.
- **Auditoría básica** de ingresos y cambios de seguridad.

## Instalación en Windows

```powershell
cd programa_fd_gc71_gc72_streamlit
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
streamlit run app.py
```

## Credencial inicial

La primera vez que corre la aplicación se crea automáticamente un administrador local:

- Usuario: `admin`
- Contraseña: `Admin123*`

Al ingresar, cambie la contraseña desde **Mi cuenta**. Luego cree los usuarios reales desde **Usuarios y perfiles**.

## Perfiles

| Perfil | Permisos principales |
|---|---|
| Administrador | Acceso total, usuarios, perfiles, auditoría, FD-GC71 y FD-GC72 |
| Coordinador | Planeación, informe académico, revisión y auditoría |
| Docente | Planeación, informe académico y cuenta propia |
| Consulta | Inicio, ayuda y cuenta propia |

## Flujo recomendado

1. Ingresar como administrador y crear docentes/coordinadores.
2. El docente diligencia FD-GC71: identificación, textos académicos, unidades, intensidad, horarios y evaluación.
3. La app genera automáticamente el calendario de sesiones según los días y horarios de clase.
4. Descargar FD-GC71 en Word y plantilla Excel de evaluación.
5. A mitad/final del curso, cargar listado tradicional y plantilla de calificaciones.
6. Generar FD-GC72 con métricas calculadas y análisis descriptivo por curso.

## Archivos de datos

La base local se crea en:

```text
app_data/fdgc_app.sqlite3
```

Esa carpeta se puede respaldar para conservar usuarios, auditoría y datos locales.

## Notas técnicas

- El login es local, suficiente para operación interna en equipo o red controlada.
- Para despliegue institucional con varios usuarios, ubique el proyecto en un servidor interno y proteja la carpeta `app_data`.
- Para autenticación corporativa futura se puede conectar contra LDAP, Microsoft Entra ID o Google Workspace.


## Versión 5.0.0 Suite Inteligente

Esta versión mejora sustancialmente la experiencia de usuario para operación institucional en Streamlit:

- Login rediseñado con pantalla institucional y mensajes de seguridad.
- Tema visual premium con tarjetas, métricas, contenedores, pestañas y botones consistentes.
- Navegación lateral con búsqueda de módulos, iconografía, descripción contextual y barra de progreso operativo.
- Panel de inicio tipo centro de mando con accesos rápidos, alertas, estado técnico y ruta operativa.
- Hero institucional por módulo con ambiente, base de datos, versión y perfil activo.
- Ruta visual del expediente: crear → planear → evidenciar → revisar → informar → cerrar.
- Centro de control rediseñado con filtros, pestañas, mapa de riesgo, gráficos y acciones sugeridas.
- Reportes ejecutivos con filtros visibles, métricas de corte y descargas más claras.
- Microcopy operativo para orientar al docente, coordinador y administrador sin saturar la pantalla.

Archivo principal para Streamlit Cloud: `app.py`.


## Versión 6.0.0 - Suite Inteligente

Incluye cargador inteligente de Excel, comparador inicial/parcial/final, semáforo de expediente, exportación masiva, hash/QR documental y asistente académico editable. Esta versión está pensada para operar el flujo completo: planeación, concertación, ejecución, seguimiento, informe, cierre y auditoría.


## Versión 7.0.1 Hotfix Streamlit Cloud

Esta versión convierte la suite en una plataforma académica institucional de punta a punta:

- Banco institucional de asignaturas con unidades, resultados, metodología, bibliografía y evaluación base.
- Motor de coherencia académica: cruza resultados de aprendizaje, contenidos, evaluación, horas, evidencias y observaciones.
- Aprobación bloqueante: los expedientes aprobados o cerrados quedan bloqueados con hash y evento de workflow.
- Informe ejecutivo institucional en Excel y Word para coordinación/comité.
- Exportación institucional estructurada por período, programa y curso.
- Auditoría por expediente con versiones, observaciones, workflow, bloqueo y registro general.

Archivo principal para Streamlit Cloud: `app.py`.


## Hotfix 7.0.1 Streamlit Cloud

Incluye timeout de conexión a PostgreSQL/Supabase, detección de `DATABASE_URL` de ejemplo y pantalla de diagnóstico cuando la base de datos falla al iniciar. Esto evita que la app se quede cargando indefinidamente en `Your app is in the oven`.
