# Checklist de producción

## Seguridad

- [ ] `APP_ENV=production` configurado.
- [ ] `DATABASE_URL` apunta a PostgreSQL externo.
- [ ] No existe `.streamlit/secrets.toml` en GitHub.
- [ ] Contraseña inicial cambiada tras el primer ingreso.
- [ ] Usuarios nominales creados por rol.
- [ ] No se comparte el usuario administrador.

## Datos

- [ ] PostgreSQL crea tablas correctamente.
- [ ] Se registran cursos.
- [ ] Se registran auditorías.
- [ ] Se cargan y descargan evidencias.
- [ ] Se genera respaldo ZIP.

## Formatos

- [ ] FD-GC71 se genera con identificación, unidades, sesiones, evaluación y socialización.
- [ ] Plantilla Excel de evaluación se genera desde contenidos/evaluaciones.
- [ ] FD-GC72 se alimenta desde listado y calificaciones.
- [ ] El informe final conserva tabla, análisis descriptivo y métricas.

## Operación

- [ ] Hay responsable de backups.
- [ ] Hay responsable de usuarios.
- [ ] Se definió límite de tamaño para evidencias.
- [ ] Se definió política para cierre de semestre.
