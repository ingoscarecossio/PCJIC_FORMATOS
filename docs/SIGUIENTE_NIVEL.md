# Módulos Next Level

## 1. Aprobaciones y versionamiento

Este módulo convierte el curso en un expediente académico controlado.

Estados disponibles:

```text
Borrador → En revisión → Observado → Ajustado → Aprobado → Cerrado
```

Cada cambio genera:

- usuario;
- fecha;
- acción;
- nota;
- versión consecutiva;
- hash documental SHA-256.

## 2. Observaciones

El coordinador o administrador puede registrar observaciones por expediente, clasificadas por:

- prioridad: Alta, Media, Baja;
- categoría: Académica, Evaluación, Cronograma, Evidencia, Forma u Otra;
- estado: Abierta o Resuelta.

## 3. Motor académico

Calcula un score de calidad del expediente sobre 100 puntos, considerando:

- evaluación concertada;
- cronograma;
- sesiones sin programar;
- evidencias;
- observaciones abiertas;
- estado del expediente;
- coherencia mínima de datos académicos.

También genera textos sugeridos para:

- justificación;
- metodología;
- plan de mejora para FD-GC72.

## 4. Reportes ejecutivos

Permite filtrar y exportar cursos por:

- periodo;
- estado;
- riesgo;
- score de calidad;
- evidencias;
- programa;
- docente.

## 5. Parámetros institucionales

Los textos base y umbrales se pueden ajustar desde interfaz sin modificar código.

Parámetros iniciales:

- `umbral_riesgo_reprobacion`
- `dias_alerta_cierre`
- `texto_metodologia_base`
- `texto_analisis_positivo_base`
- `texto_plan_mejora_base`

## 6. Recomendación productiva

Para operación institucional real:

- usar PostgreSQL/Supabase;
- configurar secretos en Streamlit Cloud;
- no usar SQLite para producción;
- hacer backup periódico;
- crear usuarios nominales;
- no operar con el usuario `admin` para tareas diarias.
