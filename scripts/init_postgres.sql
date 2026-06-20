-- Script opcional. La aplicación crea las tablas automáticamente al iniciar.
-- Úselo si desea preparar la base manualmente antes del primer despliegue.

CREATE TABLE IF NOT EXISTS usuarios (
    id SERIAL PRIMARY KEY,
    usuario TEXT UNIQUE NOT NULL,
    nombre_completo TEXT NOT NULL,
    email TEXT,
    rol TEXT NOT NULL,
    salt TEXT NOT NULL,
    password_hash TEXT NOT NULL,
    activo INTEGER NOT NULL DEFAULT 1,
    debe_cambiar_clave INTEGER NOT NULL DEFAULT 1,
    creado_en TEXT NOT NULL,
    actualizado_en TEXT,
    ultimo_login TEXT,
    intentos_fallidos INTEGER DEFAULT 0,
    bloqueado_hasta TEXT
);

CREATE TABLE IF NOT EXISTS auditoria (
    id SERIAL PRIMARY KEY,
    fecha TEXT NOT NULL,
    usuario TEXT,
    rol TEXT,
    accion TEXT NOT NULL,
    detalle TEXT
);

CREATE TABLE IF NOT EXISTS cursos (
    id SERIAL PRIMARY KEY,
    codigo TEXT,
    grupo TEXT,
    asignatura TEXT NOT NULL,
    programa TEXT,
    periodo TEXT,
    profesor TEXT,
    email_profesor TEXT,
    creditos TEXT,
    htp DOUBLE PRECISION DEFAULT 0,
    hti DOUBLE PRECISION DEFAULT 0,
    fecha_inicio TEXT,
    fecha_fin TEXT,
    estado TEXT DEFAULT 'Planeación',
    avance_contenido DOUBLE PRECISION DEFAULT 0,
    avance_evaluado DOUBLE PRECISION DEFAULT 0,
    propietario_usuario TEXT,
    creado_por TEXT,
    creado_en TEXT NOT NULL,
    actualizado_en TEXT,
    payload_json TEXT DEFAULT '{}'
);

CREATE TABLE IF NOT EXISTS evidencias (
    id SERIAL PRIMARY KEY,
    curso_id INTEGER REFERENCES cursos(id) ON DELETE SET NULL,
    tipo TEXT,
    nombre_original TEXT NOT NULL,
    nombre_archivo TEXT NOT NULL,
    mime TEXT,
    tamano INTEGER DEFAULT 0,
    descripcion TEXT,
    subido_por TEXT,
    subido_en TEXT NOT NULL,
    contenido_b64 TEXT
);

CREATE TABLE IF NOT EXISTS artefactos (
    id SERIAL PRIMARY KEY,
    curso_id INTEGER REFERENCES cursos(id) ON DELETE SET NULL,
    tipo TEXT NOT NULL,
    nombre_archivo TEXT NOT NULL,
    descripcion TEXT,
    generado_por TEXT,
    generado_en TEXT NOT NULL
);

CREATE INDEX IF NOT EXISTS idx_cursos_estado ON cursos(estado);
CREATE INDEX IF NOT EXISTS idx_cursos_owner ON cursos(propietario_usuario);
CREATE INDEX IF NOT EXISTS idx_evidencias_curso ON evidencias(curso_id);


-- V7 Institucional
CREATE TABLE IF NOT EXISTS asignaturas_base (
    id SERIAL PRIMARY KEY,
    codigo TEXT,
    nombre TEXT NOT NULL,
    programa TEXT,
    area_formacion TEXT,
    creditos TEXT,
    htp DOUBLE PRECISION DEFAULT 0,
    hti DOUBLE PRECISION DEFAULT 0,
    tipo_asignatura TEXT,
    justificacion TEXT,
    competencias TEXT,
    resultados TEXT,
    objetivos TEXT,
    metodologia TEXT,
    ambientes TEXT,
    medios TEXT,
    bibliografia TEXT,
    unidades_json TEXT DEFAULT '[]',
    evaluaciones_json TEXT DEFAULT '[]',
    activo INTEGER DEFAULT 1,
    creado_por TEXT,
    creado_en TEXT,
    actualizado_en TEXT
);

CREATE TABLE IF NOT EXISTS workflow_eventos (
    id SERIAL PRIMARY KEY,
    curso_id INTEGER,
    evento TEXT NOT NULL,
    estado_anterior TEXT,
    estado_nuevo TEXT,
    resultado TEXT,
    detalle TEXT,
    hash_expediente TEXT,
    usuario TEXT,
    rol TEXT,
    creado_en TEXT NOT NULL
);

CREATE TABLE IF NOT EXISTS curso_bloqueos (
    curso_id INTEGER PRIMARY KEY,
    bloqueado INTEGER DEFAULT 0,
    motivo TEXT,
    hash_bloqueo TEXT,
    bloqueado_por TEXT,
    bloqueado_en TEXT
);
