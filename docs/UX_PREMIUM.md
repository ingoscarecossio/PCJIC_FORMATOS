# Guía UX Premium

## Objetivo

La versión 5.0.0 convierte la aplicación en una experiencia más clara para docentes, coordinadores y administradores. El diseño prioriza tres cosas: orientación, trazabilidad y rapidez operativa.

## Principios de diseño

1. **Menos fricción:** los módulos se buscan desde la barra lateral y cada página explica brevemente para qué sirve.
2. **Ruta visible:** el usuario ve el flujo del expediente académico: crear, planear, evidenciar, revisar, informar y cerrar.
3. **Decisión rápida:** el inicio y el centro de control muestran métricas, alertas y acciones sugeridas.
4. **Operación institucional:** los mensajes diferencian modo local/demo de modo productivo con PostgreSQL/Supabase.
5. **Microcopy útil:** se agregan ayudas breves donde normalmente el usuario se pierde.

## Cambios visibles

- Login institucional.
- Sidebar con búsqueda, iconos y descripción del módulo.
- Hero superior por módulo.
- Cards de acceso rápido.
- Centro de control por pestañas.
- Reportes ejecutivos con filtros agrupados.
- Alertas y estado técnico en el inicio.

## Recomendación de uso

Para operación real, el administrador debe configurar secretos de Streamlit Cloud y usar PostgreSQL/Supabase mediante `DATABASE_URL`. SQLite debe quedar solo para desarrollo, demos o pruebas locales.
