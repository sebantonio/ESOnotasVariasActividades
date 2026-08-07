# Diseño — Adaptar app a plantilla con instrumentos por criterio

**Fecha:** 2026-08-07
**Proyecto:** ESOnotasVariasActividades (continuación de ESOplantillaNotas)
**Estado:** aprobado para pasar a plan de implementación

## Contexto

`ESOnotasVariasActividades` parte como copia exacta de `ESOplantillaNotas`
(mismo `src-tauri`, mismo frontend). El Excel de trabajo cambia: se comparó
`Plantilla_Notas_ESO_ANTIGUO.xlsx` (estructura actual, la que soporta la app)
contra `Plantilla_Notas_ESO.xlsx` (plantilla nueva) con `openpyxl`. La única
diferencia estructural real está en las 15 hojas de unidad (`U1`..`U15`): cada
criterio (`CR1.1`, `CR1.2`...) pasa de evaluarse con una nota directa a
evaluarse con hasta 4 notas de instrumento que se combinan en una media
ponderada. Las hojas `DATOS`, `PESOS` y las hojas de evaluación (`1ª EVA`,
`2ª EVA`, `3ª EVA`, `FINAL`, variantes `-solo`) no cambian de estructura.

## Diferencia de layout en hoja de unidad (`Uxx`)

**Antes (2 columnas por criterio):**

| Fila (Excel) | Contenido |
| --- | --- |
| 3 | Cabecera: `CR1.1 \| Rec \| CR1.2 \| Rec \| ...` |
| 4 | Fórmula de ponderación del criterio (`INDEX/MATCH` sobre `PESOS`), solo en la columna del CR |
| 5+ | Datos de alumno: nota directa (manual) + Rec (manual) |

**Ahora (6 columnas por criterio):**

| Fila (Excel) | Contenido |
| --- | --- |
| 2 | Fila 2 col B..F: `nota final unidad \| i1 \| i2 \| i3 \| i4` — etiquetas de los 4 slots de instrumento, **compartidas por toda la hoja** (no por criterio) |
| 3 | Sub-cabecera por bloque de criterio: `=i1 \| =i2 \| =i3 \| =i4 \| FINAL` |
| 4 | `peso del instrumento`: pesos de los 4 slots (ej. 20/40/20/20%), **globales para toda la hoja**, editable |
| 5 | Código de criterio (`CR1.1`) en la 1ª col del bloque + `Rec` en la última |
| 6 | Fórmula de ponderación del criterio (`INDEX/MATCH` sobre `PESOS`), en la 1ª col del bloque — mismo mecanismo que antes, solo desplazada |
| 7+ | Datos de alumno: `i1, i2, i3, i4` (manual), `FINAL` (fórmula `SUMPRODUCT`, **solo lectura**), `Rec` (manual) |

Fórmula de `FINAL` (columna L del primer bloque, ejemplo real de `U1!L7`):

```
=IFERROR(SUMPRODUCT((H7:K7<>"")*C$4:F$4*H7:K7)/SUMPRODUCT((H7:K7<>"")*C$4:F$4),"")
```

Media ponderada de los 4 instrumentos, ignorando slots vacíos tanto en el
numerador como en el denominador.

Las hojas de evaluación referencian ahora la nueva posición de `FINAL`/`Rec`
dentro de `Uxx` (ej. `N('U1'!L7)*'U1'!H$6` en vez de `'U1'!B5*'U1'!B$4`), pero
su propio layout (fila 17 cabecera `NOTA CE`/`CRx.y`, 2 columnas por criterio)
**no cambia**.

No existe lógica en `main.rs` que clone hojas desde `Plantilla Unidades`
(grep sin resultados) — las 15 hojas `Uxx` ya existen fijas en la plantilla.
Añadir una 16ª unidad sería una operación manual de Excel fuera del alcance
de la app; no se implementa gestión de "nueva hoja de unidad".

## Decisiones (confirmadas con el usuario)

1. **Pesos de los 4 instrumentos son editables por unidad** desde la app
   (no fijos). Deben sumar 100%.
2. **Los 4 slots (i1-i4) están ligados al catálogo de Instrumentos**
   (`gestor-instrumentos.html`, hoja `DATOS!N5:O13`, máx. 10 instrumentos
   definidos, pero cada unidad solo puede usar 4 simultáneos — límite
   estructural de columnas del template). El profesor elige, por unidad, qué
   4 instrumentos del catálogo ocupan los slots i1-i4.

## Cambios en backend (`src-tauri/src/main.rs`)

### `load_notas_unidad` (línea ~1236)

- Detección de cabecera de criterio: escanear fila 5 (0-idx 4) en vez de
  fila 3 (0-idx 2). Ancho de bloque pasa de 2 a 6 columnas.
- `first_row` (inicio de alumnos) pasa de `4` (fila 5) a `6` (fila 7).
- Leer, una vez por unidad (no por criterio):
  - Fila 2 (0-idx 1), columnas C:F (0-idx 2:6) → abreviaturas de instrumento
    asignadas a los 4 slots.
  - Fila 4 (0-idx 3), columnas C:F → pesos de cada slot.
- Por criterio, `crNotas` pasa de `{codigo, colIdx, nota, display, recDisplay}`
  a incluir los 4 valores de instrumento + `final` calculado:
  ```
  { codigo, colIdx, i1ColIdx: colIdx, i2ColIdx: colIdx+1, i3ColIdx: colIdx+2,
    i4ColIdx: colIdx+3, finalColIdx: colIdx+4, recColIdx: colIdx+5,
    i1, i2, i3, i4, final, recDisplay }
  ```
  `final` se recalcula en Rust con la misma fórmula que la celda Excel
  (media ponderada de i1..i4 ignorando vacíos, usando los pesos leídos de
  fila 4) — no se lee de la celda `FINAL` porque es una fórmula y
  `calamine`/lectura XML puede no tener el caché actualizado si el usuario
  edita fuera de Excel.
- Respuesta añade `instrumentosUnidad: [{slot: "i1", abrev, peso}, ...]`.

### `excel_save_notas_unidad_impl` (línea ~2240)

- Payload `crNotas[codigo]` pasa de `{nota, rec}` a `{i1, i2, i3, i4, rec}`
  (cualquier campo ausente no se toca, igual que hoy con `nota`/`rec`
  opcionales).
- Escribe i1..i4 en `colIdx..colIdx+3` de la hoja `Uxx`, `rec` en
  `colIdx+5`. **Nunca** escribe en `colIdx+4` (`FINAL`, fórmula).
- Antes de llamar a `build_eval_sheet_edits`, calcula `final` en Rust (misma
  fórmula) a partir de los i1..i4 resultantes (los nuevos si vienen en el
  payload, si no los existentes en la hoja) y lo pasa como el valor `nota`
  que ya consumía esa función — **`build_eval_sheet_edits` no cambia**, sigue
  operando por `(alumno, código)` sobre las hojas de evaluación, cuyo layout
  es idéntico al de antes.

### Nuevo comando `excel_save_unidad_instrumentos`

- Payload: `{ unidad, slots: [{abrev, peso}, ...] }` (1 a 4 slots).
- Valida que la suma de pesos sea 100% (o 0 si no hay ninguno configurado
  aún) y que cada `abrev` exista en el catálogo de instrumentos
  (`DATOS!N5:O13`).
- Escribe abreviaturas en fila 2 col C:F y pesos (como fracción, ej. 0.2)
  en fila 4 col C:F de la hoja `Uxx`.
- Registrar en `tauri::generate_handler!` junto a los demás comandos de
  unidad (línea ~3102).

## Cambios en frontend

### `gestor-notas.html`

- Cada criterio pasa de 1 celda editable a 5 celdas: `i1, i2, i3, i4`
  (inputs) + `FINAL` (solo lectura, con las mismas clases de color
  `nota-alta/media/baja` aplicadas sobre el valor `final`, no sobre un input
  crudo).
- Cabecera de columna de cada instrumento muestra la abreviatura real
  (`cr.instrumentosUnidad[i].abrev`), no `i1`/`i2` literal.
- Construcción de payload (`crNotasMap`) añade `i1..i4` en vez de `nota`.
- Autosave/validación por celda: mismo mecanismo, pero por slot de
  instrumento en vez de por criterio.

### `gestor-unidades.html`

- Nueva sección/modal por unidad: selector de hasta 4 instrumentos del
  catálogo + input de peso (%) por slot, con validación de suma 100% antes
  de guardar (llama a `excel_save_unidad_instrumentos`).

### `gestor-recuperaciones.html`

- `recomputeAlumno()`: donde hoy usa la nota cruda del criterio para
  agregarla a la Nota CE del grupo, usa `final` (ya ponderado). El campo
  `Rec` sigue siendo un override manual igual que hoy, solo cambia su
  posición de columna dentro del bloque (ya resuelta por backend, no
  requiere cambio de lógica en JS más allá de leer `recColIdx` del nuevo
  payload).

### `visor-unidades.html`

- Vista de solo lectura: mostrar i1..i4 + `FINAL` en vez de una sola
  columna de nota por criterio.

## Fuera de alcance

- Hojas de evaluación (`1ª EVA`, `2ª EVA`, `3ª EVA`, `FINAL`, variantes
  `-solo`): sin cambios de estructura ni de lógica de agregación CE.
- Hoja `PESOS`: sin cambios (ponderación de criterio dentro de CE sigue
  igual).
- `informes.html`, `diario.html`: sin cambios, no dependen de la hoja de
  unidad.
- Gestión de alta de nuevas hojas `Uxx` (más allá de las 15 ya existentes en
  la plantilla): no la usa la app hoy, no se añade.

## Testing

- No hay tests de integración existentes sobre `load_notas_unidad`/
  `excel_save_notas_unidad` (pendiente conocido del proyecto, ver
  `CLAUDE.md`). Para este cambio, verificación manual mínima:
  1. Abrir la app con `Plantilla_Notas_ESO.xlsx` (nueva) seleccionada.
  2. `gestor-notas.html`: introducir i1..i4 en un par de criterios de `U1`,
     comprobar que `FINAL` se recalcula visualmente igual que en Excel
     abierto en paralelo (abrir el `.xlsx` tras guardar y comparar el valor
     de la celda `FINAL`, que recalcula sola al abrir).
  3. `gestor-unidades.html`: cambiar pesos de instrumento de una unidad,
     guardar, reabrir y comprobar que persisten.
  4. `gestor-recuperaciones.html`: comprobar que la Nota CE por grupo se
     recalcula con el `FINAL` correcto tras editar Rec.
  5. `visor-unidades.html`: comprobar que se listan i1..i4 + FINAL sin
     errores para una unidad con datos.
- Riesgo principal: fórmula `SUMPRODUCT` de Excel y el cálculo replicado en
  Rust deben coincidir exactamente (mismo criterio de "ignorar vacíos") para
  que el caché que la app escribe en las hojas de evaluación no diverja del
  valor que Excel recalcula al abrir el archivo.
