# Diccionario de datos resumido

## usuarios
- `usuario`: identificador de ingreso.
- `rol`: Administrador, Coordinador, Docente o Consulta.
- `salt`, `password_hash`: credenciales protegidas.
- `debe_cambiar_clave`: obliga cambio de clave inicial.

## cursos
- `codigo`, `grupo`, `asignatura`, `programa`, `periodo`: identificación académica.
- `profesor`, `email_profesor`: responsable.
- `htp`, `hti`: dedicación horaria.
- `estado`: Planeación, En ejecución, Corte parcial, Cierre final o Archivado.
- `avance_contenido`, `avance_evaluado`: indicadores para FD-GC72.
- `payload_json`: planeación extendida.

## evidencias
- `curso_id`: vínculo con curso.
- `tipo`: clasificación del soporte.
- `nombre_original`: nombre del archivo cargado.
- `nombre_archivo`: nombre interno seguro.
- `subido_por`, `subido_en`: trazabilidad.

## auditoria
- `fecha`, `usuario`, `rol`, `accion`, `detalle`: bitácora operativa.
