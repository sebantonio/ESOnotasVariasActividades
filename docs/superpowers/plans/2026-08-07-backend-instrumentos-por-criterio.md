# Backend: instrumentos por criterio en hojas de unidad — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Adaptar `src-tauri/src/main.rs` para que `load_notas_unidad` y
`excel_save_notas_unidad` lean/escriban el nuevo layout de las hojas `Uxx`
(4 notas de instrumento + `FINAL` calculado + `Rec`, en vez de 1 nota + `Rec`
directos) y añadir un comando para editar los pesos/instrumentos de cada
unidad. No toca frontend (plan aparte, después de este).

**Architecture:** Todo el cambio vive en `src-tauri/src/main.rs` (el proyecto
mantiene un único archivo de backend — no se reestructura). Se extraen 3
funciones puras y testeables sin I/O (`compute_final_weighted`,
`cr_cols_from_cells`, `parse_unidad_instrument_config`) siguiendo el patrón ya
usado en el archivo (`find_evaluation_layout_indices`, con sus tests en
`mod eval_layout_tests`). El resto de cambios son ediciones quirúrgicas a
`load_notas_unidad` y `excel_save_notas_unidad_impl` para usar los nuevos
offsets de fila/columna, más un comando nuevo `excel_save_unidad_instrumentos`.

**Tech Stack:** Rust, `calamine` (lectura xlsx), `zip`+regex (escritura XML
directa del xlsx), `serde_json`, Tauri v2 `#[tauri::command]`.

## Global Constraints

- No añadir dependencias nuevas a `Cargo.toml` (no hay `tempfile` — los tests
  de integración copian el fixture a `std::env::temp_dir()` a mano).
- Nunca escribir en la columna `FINAL` de una hoja `Uxx` (es una fórmula
  `SUMPRODUCT`, se recalcula sola al abrir el Excel en Excel — CLAUDE.md: "CR
  y Rec son FÓRMULAS... NUNCA escribir aquí directamente").
- `build_eval_sheet_edits` (línea ~2151) **no se modifica** — sigue operando
  por `(alumno, código)` sobre las hojas de evaluación, cuyo layout no cambió.
- El fixture real `Plantilla_Notas_ESO.xlsx` (committeado en la raíz del
  repo) es la fuente de verdad para los tests de integración — no se crean
  xlsx sintéticos para eso, son demasiado frágiles de construir a mano dado
  el formato OOXML.
- Números reales verificados contra ese fixture (hoja `U1`): 100 criterios
  (`CR1.1`..`CR10.10`), primer `colIdx` = 7 (columna H), último `colIdx` =
  601, pesos de instrumento en fila 4 (Excel) columnas C:F = `[0.2, 0.4, 0.2,
  0.2]`, etiquetas en fila 2 columnas C:F = `["i1","i2","i3","i4"]`, primera
  fila de alumno = fila 7 (Excel) = índice 6 (0-idx). Catálogo de
  instrumentos en `DATOS!N5:O9` = `PE, TD, TI, TG, L`.

---

### Task 1: `compute_final_weighted` — media ponderada ignorando vacíos

**Files:**
- Modify: `src-tauri/src/main.rs` (añadir función cerca de `normalize_grade`,
  línea ~2057, antes de `excel_save_notas_actividad`)
- Test: mismo archivo, nuevo módulo `#[cfg(test)] mod final_weighted_tests`
  (añadir justo después del nuevo `fn compute_final_weighted`, siguiendo el
  patrón de `mod eval_layout_tests` en la línea 3032)

**Interfaces:**
- Produces: `fn compute_final_weighted(values: [Option<f64>; 4], weights: [f64; 4]) -> Option<f64>`
  — usado por Task 3 (lectura) y Task 5 (sincronización a hojas de evaluación).

- [ ] **Step 1: Escribir los tests (deben fallar: la función no existe aún)**

Añadir en `src-tauri/src/main.rs`, después de la línea 3030 (cierre de
`mod eval_layout_tests`) — módulo nuevo, no lo metas dentro del existente:

```rust
#[cfg(test)]
mod final_weighted_tests {
    use super::*;

    // Replica la fórmula real de la plantilla (U1!L7):
    // =IFERROR(SUMPRODUCT((H7:K7<>"")*C$4:F$4*H7:K7)/SUMPRODUCT((H7:K7<>"")*C$4:F$4),"")
    const PESOS_REALES: [f64; 4] = [0.2, 0.4, 0.2, 0.2];

    #[test]
    fn media_ponderada_con_los_4_instrumentos() {
        let values = [Some(8.0), Some(7.0), Some(9.0), Some(6.0)];
        let result = compute_final_weighted(values, PESOS_REALES);
        // (0.2*8 + 0.4*7 + 0.2*9 + 0.2*6) / (0.2+0.4+0.2+0.2) = 7.4 / 1.0
        assert_eq!(result, Some(7.4));
    }

    #[test]
    fn ignora_instrumentos_vacios_en_numerador_y_denominador() {
        let values = [Some(8.0), None, Some(9.0), Some(6.0)];
        let result = compute_final_weighted(values, PESOS_REALES);
        // (0.2*8 + 0.2*9 + 0.2*6) / (0.2+0.2+0.2) = 4.6 / 0.6 = 7.666...
        assert!((result.unwrap() - 7.6666666666666).abs() < 1e-9);
    }

    #[test]
    fn valor_cero_explicito_cuenta_distinto_de_vacio() {
        // Un 0 tecleado SI cuenta (no es "" en Excel); solo None (celda vacia) se ignora.
        let values = [Some(0.0), None, None, None];
        let result = compute_final_weighted(values, PESOS_REALES);
        assert_eq!(result, Some(0.0));
    }

    #[test]
    fn todos_vacios_devuelve_none() {
        let result = compute_final_weighted([None, None, None, None], PESOS_REALES);
        assert_eq!(result, None);
    }

    #[test]
    fn peso_cero_en_slot_presente_no_distorsiona_el_resto() {
        let values = [Some(10.0), Some(4.0), None, None];
        let weights = [0.0, 0.5, 0.3, 0.2];
        let result = compute_final_weighted(values, weights);
        // (0*10 + 0.5*4) / (0 + 0.5) = 2.0 / 0.5 = 4.0
        assert_eq!(result, Some(4.0));
    }
}
```

- [ ] **Step 2: Ejecutar los tests para verificar que fallan**

Run: `cargo test --manifest-path src-tauri/Cargo.toml final_weighted_tests`
Expected: FAIL — `cannot find function 'compute_final_weighted' in this scope`

- [ ] **Step 3: Implementar la función**

Añadir en `src-tauri/src/main.rs`, justo antes de `fn normalize_grade` (línea
2051):

```rust
// Replica SUMPRODUCT((vals<>"")*pesos*vals)/SUMPRODUCT((vals<>"")*pesos) de la
// plantilla: media ponderada de hasta 4 instrumentos, ignorando los que estan
// vacios (None) tanto en el numerador como en el denominador. Un valor 0
// explicito SI cuenta (distinto de vacio, igual que en Excel "0" <> "").
fn compute_final_weighted(values: [Option<f64>; 4], weights: [f64; 4]) -> Option<f64> {
    let mut num = 0.0;
    let mut den = 0.0;
    for i in 0..4 {
        if let Some(v) = values[i] {
            num += weights[i] * v;
            den += weights[i];
        }
    }
    if den == 0.0 { None } else { Some(num / den) }
}
```

- [ ] **Step 4: Ejecutar los tests para verificar que pasan**

Run: `cargo test --manifest-path src-tauri/Cargo.toml final_weighted_tests`
Expected: PASS (5 tests)

- [ ] **Step 5: Commit**

```bash
git add src-tauri/src/main.rs
git commit -m "feat(backend): compute_final_weighted - media ponderada de instrumentos por criterio"
```

---

### Task 2: `cr_cols_from_cells` — detección de criterios generalizada

**Files:**
- Modify: `src-tauri/src/main.rs:1253-1278` (bloque de detección de `cr_cols`
  dentro de `load_notas_unidad`)
- Test: nuevo módulo `#[cfg(test)] mod cr_cols_tests`

**Interfaces:**
- Consumes: `is_cr_code(s: &str) -> bool` (línea 375, sin cambios)
- Produces: `fn cr_cols_from_cells(cells: impl Iterator<Item = (usize, String)>) -> Vec<(String, usize)>`
  — usado por Task 3 para reemplazar el escaneo inline de `cr_cols`.

**Contexto:** hoy el código busca la cabecera de criterios en la fila 3 y,
como fallback, en la fila 4 (Excel, 1-indexed) de la hoja `Uxx` — layout
antiguo. En la plantilla nueva la cabecera de códigos `CRx.y` está en la fila
5, con fallback en la fila 6. La lógica de "de una fila de celdas, sacar los
`(código, columna)` que sean CRx.y" es idéntica en ambos casos — se extrae a
una función pura para poder testearla sin abrir un xlsx, replicando el patrón
que ya usa `find_evaluation_layout_indices`/`mod eval_layout_tests`.

- [ ] **Step 1: Escribir los tests (deben fallar: la función no existe aún)**

Añadir en `src-tauri/src/main.rs`, después del nuevo `mod final_weighted_tests`
del Task 1:

```rust
#[cfg(test)]
mod cr_cols_tests {
    use super::*;

    #[test]
    fn extrae_codigos_cr_ordenados_por_columna() {
        let cells = vec![
            (13usize, "CR1.2".to_string()),
            (7usize, "CR1.1".to_string()),
            (2usize, "Alumno".to_string()),
        ];
        let result = cr_cols_from_cells(cells.into_iter());
        assert_eq!(result, vec![
            ("CR1.1".to_string(), 7),
            ("CR1.2".to_string(), 13),
        ]);
    }

    #[test]
    fn normaliza_a_mayusculas() {
        let cells = vec![(7usize, "cr1.1".to_string())];
        let result = cr_cols_from_cells(cells.into_iter());
        assert_eq!(result, vec![("CR1.1".to_string(), 7)]);
    }

    #[test]
    fn ignora_celdas_que_no_son_codigo_cr() {
        let cells = vec![
            (0usize, "Alumno".to_string()),
            (7usize, "CR1.1".to_string()),
            (12usize, "Rec".to_string()),
            (13usize, "FINAL".to_string()),
        ];
        let result = cr_cols_from_cells(cells.into_iter());
        assert_eq!(result, vec![("CR1.1".to_string(), 7)]);
    }

    #[test]
    fn fila_sin_codigos_devuelve_vacio() {
        let cells = vec![(0usize, "nota final unidad".to_string())];
        assert_eq!(cr_cols_from_cells(cells.into_iter()), Vec::<(String, usize)>::new());
    }
}
```

- [ ] **Step 2: Ejecutar los tests para verificar que fallan**

Run: `cargo test --manifest-path src-tauri/Cargo.toml cr_cols_tests`
Expected: FAIL — `cannot find function 'cr_cols_from_cells'`

- [ ] **Step 3: Implementar la función y usarla en `load_notas_unidad`**

Añadir la función nueva justo antes de `fn load_notas_unidad` (línea 1236):

```rust
// Dado un conjunto de celdas (columna, texto) de una fila de cabecera, extrae
// los pares (codigo CR en mayusculas, columna) ordenados por columna. Pura y
// testeable sin abrir ningun xlsx — reutilizada tanto para la lectura via XML
// directo como para el fallback via calamine.
fn cr_cols_from_cells(cells: impl Iterator<Item = (usize, String)>) -> Vec<(String, usize)> {
    let mut sorted: Vec<(usize, String)> = cells.collect();
    sorted.sort_by_key(|(ci, _)| *ci);
    sorted.into_iter()
        .filter(|(_, s)| is_cr_code(s))
        .map(|(ci, s)| (s.to_uppercase(), ci))
        .collect()
}
```

Reemplazar el bloque de detección de `cr_cols` dentro de `load_notas_unidad`
(líneas 1253-1278 del archivo actual) por:

```rust
    // Detectar CRs en hoja Ux usando XML directo (calamine trunca columnas lejanas
    // en hojas anchas — la plantilla nueva llega hasta la columna ~601).
    // Fila 5 en Excel = row_1=5 (cabecera de criterios en el layout con
    // instrumentos por criterio); fallback fila 6 por si el layout varia.
    let mut cr_cols: Vec<(String, usize)> = Vec::new();
    for check_row_1 in [5usize, 6usize] {
        let xml_row = read_row_from_xml(path, unidad, check_row_1);
        if !xml_row.is_empty() {
            cr_cols = cr_cols_from_cells(xml_row.into_iter());
        }
        if !cr_cols.is_empty() { break; }
    }
    // Fallback a calamine si XML falla (sin tope de columna: la plantilla nueva
    // tiene criterios mas alla de la columna 200 que usaba el tope antiguo).
    if cr_cols.is_empty() {
        for check_ri in 4..=5usize {
            if let Some(row) = rows.get(check_ri) {
                let cells = row.iter().enumerate().map(|(ci, v)| (ci, cell_val_str(v)));
                cr_cols = cr_cols_from_cells(cells);
            }
            if !cr_cols.is_empty() { break; }
        }
    }
```

- [ ] **Step 4: Ejecutar los tests para verificar que pasan**

Run: `cargo test --manifest-path src-tauri/Cargo.toml cr_cols_tests`
Expected: PASS (4 tests)

- [ ] **Step 5: Commit**

```bash
git add src-tauri/src/main.rs
git commit -m "feat(backend): cr_cols_from_cells - deteccion de criterios en fila 5 (layout con instrumentos)"
```

---

### Task 3: `load_notas_unidad` — leer instrumentos por criterio

**Files:**
- Modify: `src-tauri/src/main.rs:1236-1358` (`load_notas_unidad`, resto de la
  función tras el cambio del Task 2)
- Test: nuevo módulo `#[cfg(test)] mod load_notas_unidad_tests` (integración,
  usa el fixture real committeado)

**Interfaces:**
- Consumes: `compute_final_weighted` (Task 1), `cr_cols_from_cells` (Task 2),
  `cell_f64`, `cell_str`, `read_sheet_rows`, `list_unit_sheets`, `load_alumnos`
  — todas sin cambios de firma.
- Produces: `load_notas_unidad(path: &str, unidad: &str) -> Result<Value, String>`
  cuyo JSON de salida cambia de forma (ver Step 3) — Task 4/5 y el futuro plan
  de frontend consumen esta forma nueva.

**Contexto:** hoy cada criterio en la respuesta es
`{codigo, colIdx, recColIdx, nota, display, recDisplay}` y los alumnos
empiezan en la fila 5 (Excel, índice 4). En el layout nuevo cada bloque de
criterio ocupa 6 columnas (`i1,i2,i3,i4,FINAL,Rec`) y los alumnos empiezan en
la fila 7 (índice 6). Además hay que leer, una vez por hoja, las etiquetas
(fila 2, columnas C:F) y pesos (fila 4, columnas C:F) de instrumento.

- [ ] **Step 1: Escribir el test de integración (debe fallar: la forma vieja no tiene estos campos)**

Añadir al final de `src-tauri/src/main.rs`, después de `mod cr_cols_tests` del
Task 2:

```rust
#[cfg(test)]
mod load_notas_unidad_tests {
    use super::*;

    // Fixture real committeado en la raiz del repo (mismo archivo que usa la app
    // como plantilla vacia). Verificado a mano con openpyxl: hoja U1, 100
    // criterios (CE1..CE10 x 10 CR), primer colIdx=7 (columna H), ultimo
    // colIdx=601, pesos fila4 C:F=[0.2,0.4,0.2,0.2], etiquetas fila2 C:F=
    // ["i1","i2","i3","i4"].
    fn fixture_path() -> String {
        concat!(env!("CARGO_MANIFEST_DIR"), "/../Plantilla_Notas_ESO.xlsx").to_string()
    }

    #[test]
    fn lee_100_criterios_de_u1_con_offsets_correctos() {
        let data = load_notas_unidad(&fixture_path(), "U1").expect("debe cargar U1");
        let criterios = data["criterios"].as_array().expect("criterios debe ser array");
        assert_eq!(criterios.len(), 100, "U1 tiene 100 criterios (CE1..CE10 x 10): {criterios:?}");

        let primero = &criterios[0];
        assert_eq!(primero["codigo"], "CR1.1");
        assert_eq!(primero["colIdx"], 7);

        let ultimo = criterios.last().unwrap();
        assert_eq!(ultimo["codigo"], "CR10.10");
        assert_eq!(ultimo["colIdx"], 601);
    }

    #[test]
    fn lee_instrumentos_y_pesos_de_la_unidad() {
        let data = load_notas_unidad(&fixture_path(), "U1").expect("debe cargar U1");
        let instrumentos = data["instrumentosUnidad"].as_array().expect("debe existir instrumentosUnidad");
        assert_eq!(instrumentos.len(), 4);
        let pesos: Vec<f64> = instrumentos.iter().map(|i| i["peso"].as_f64().unwrap()).collect();
        assert_eq!(pesos, vec![0.2, 0.4, 0.2, 0.2]);
        let abrevs: Vec<String> = instrumentos.iter().map(|i| i["abrev"].as_str().unwrap().to_string()).collect();
        assert_eq!(abrevs, vec!["i1", "i2", "i3", "i4"]);
    }

    #[test]
    fn alumnos_empiezan_en_fila_indice_6() {
        let data = load_notas_unidad(&fixture_path(), "U1").expect("debe cargar U1");
        let alumnos = data["alumnos"].as_array().expect("alumnos debe ser array");
        assert!(!alumnos.is_empty(), "la plantilla trae alumnos de ejemplo");
        assert_eq!(alumnos[0]["rowIdx"], 6, "primer alumno en fila Excel 7 = indice 6");
    }

    #[test]
    fn cada_criterio_de_alumno_trae_i1_i4_final_y_rec() {
        let data = load_notas_unidad(&fixture_path(), "U1").expect("debe cargar U1");
        let alumno = &data["alumnos"][0];
        let cr = alumno["crNotas"].as_array().unwrap().iter()
            .find(|c| c["codigo"] == "CR1.1")
            .expect("CR1.1 debe estar presente");
        for campo in ["i1", "i2", "i3", "i4", "final", "recDisplay"] {
            assert!(cr.get(campo).is_some(), "falta el campo {campo} en {cr:?}");
        }
    }
}
```

- [ ] **Step 2: Ejecutar los tests para verificar que fallan**

Run: `cargo test --manifest-path src-tauri/Cargo.toml load_notas_unidad_tests`
Expected: FAIL — `criterios.len()` no será 100 (sigue leyendo layout viejo) y
`instrumentosUnidad` no existirá en el JSON.

- [ ] **Step 3: Reescribir el resto de `load_notas_unidad`**

`first_row` (línea 1241) pasa de `4` a `6`:

```rust
    let first_row: usize = 6; // fila 7 en Excel (0-indexed) — layout con instrumentos por criterio
```

Justo después del bloque de detección de `cr_cols` (ya reescrito en el
Task 2), añadir la lectura de instrumentos de la unidad:

```rust
    // Instrumentos de la unidad: fila 2 (etiquetas) y fila 4 (pesos), columnas
    // C:F (0-idx 2..6) — globales para toda la hoja, compartidos por todos los
    // criterios (no hay una fila 2/4 por bloque, todos los bloques referencian
    // estas mismas 4 celdas via formula '=$C$2' etc.).
    let instrument_weights: [f64; 4] = [
        cell_f64(&rows, 3, 2).unwrap_or(0.0),
        cell_f64(&rows, 3, 3).unwrap_or(0.0),
        cell_f64(&rows, 3, 4).unwrap_or(0.0),
        cell_f64(&rows, 3, 5).unwrap_or(0.0),
    ];
    let instrumentos_unidad: Vec<Value> = (0..4usize).map(|slot| {
        json!({
            "slot": slot,
            "abrev": cell_str(&rows, 1, 2 + slot),
            "peso": instrument_weights[slot],
        })
    }).collect();
```

Reemplazar el bloque que construye `cr_notas` por alumno (líneas 1345-1350
del archivo actual) por:

```rust
        let cr_notas: Vec<Value> = cr_cols.iter().map(|(code, ci)| {
            let i1 = cell_f64(&rows, ri, *ci);
            let i2 = cell_f64(&rows, ri, *ci + 1);
            let i3 = cell_f64(&rows, ri, *ci + 2);
            let i4 = cell_f64(&rows, ri, *ci + 3);
            let final_val = compute_final_weighted([i1, i2, i3, i4], instrument_weights);
            let rec_display = cell_str(&rows, ri, *ci + 5);
            json!({
                "codigo": code, "colIdx": ci,
                "i1": i1, "i2": i2, "i3": i3, "i4": i4,
                "final": final_val,
                "recDisplay": rec_display
            })
        }).collect();
```

Y en el `criterios_json` (líneas 1330-1335), `recColIdx` pasa de `ci + 1` a
`ci + 5`:

```rust
    let criterios_json: Vec<Value> = cr_cols.iter()
        .map(|(code, ci)| {
            let ponderacion = ponderaciones_por_cr.get(code).copied().unwrap_or(0.0);
            json!({ "codigo": code, "colIdx": ci, "recColIdx": ci + 5, "ponderacion": ponderacion })
        })
        .collect();
```

Y en el `json!` de retorno final de la función (línea 1357), añadir
`instrumentosUnidad`:

```rust
    Ok(json!({
        "filePath": path, "fileName": file_name, "unidad": unidad, "titulo": titulo,
        "unidades": unidades, "criterios": criterios_json, "instrumentosUnidad": instrumentos_unidad,
        "alumnos": alumnos
    }))
```

- [ ] **Step 4: Ejecutar los tests para verificar que pasan**

Run: `cargo test --manifest-path src-tauri/Cargo.toml load_notas_unidad_tests`
Expected: PASS (4 tests). Si `lee_100_criterios_de_u1_con_offsets_correctos`
falla porque `criterios.len()` da un número mucho menor que 100 (p.ej. se
corta en la columna ~200), es que calamine está truncando la hoja pese al XML
directo del header — en ese caso, extender también la construcción de
`cr_notas` (Step 3, arriba) para leer i1..i4 vía `read_row_from_xml` en vez de
`cell_f64(&rows, ...)`, igual que ya se hace para el header. No lo escribas
preventivamente si el test pasa: la hoja del fixture real declara su rango
completo (`<dimension ref="A1:WI43"/>`) y calamine debería respetarlo.

- [ ] **Step 5: `cargo build` completo para detectar errores de tipos en el resto del archivo**

Run: `cargo build --manifest-path src-tauri/Cargo.toml`
Expected: compila sin errores (otros comandos que llaman a `load_notas_unidad`
— `excel_get_notas_unidad`, `excel_resync_unidad_eval_impl` — siguen
compilando porque solo leen campos por nombre de un `Value` genérico).

- [ ] **Step 6: Commit**

```bash
git add src-tauri/src/main.rs
git commit -m "feat(backend): load_notas_unidad lee instrumentos i1-i4 + FINAL calculado del layout nuevo"
```

---

### Task 4: `excel_save_notas_unidad_impl` — escribir i1-i4 y Rec en los offsets nuevos

**Files:**
- Modify: `src-tauri/src/main.rs:2251-2276` (`unit_edit_fn` dentro de
  `excel_save_notas_unidad_impl`)
- Test: nuevo módulo `#[cfg(test)] mod save_notas_unidad_tests` (integración,
  usa una COPIA del fixture en `std::env::temp_dir()` — nunca escribe sobre el
  fixture committeado)

**Interfaces:**
- Consumes: `set_xml_cell`, `normalize_grade`, `edit_workbook_sheets_xml`
  (sin cambios de firma), `load_notas_unidad` (Task 3, para verificar en el
  test que lo escrito se relee bien).
- Produces: el payload que acepta `excel_save_notas_unidad` cambia de
  `crNotas[codigo] = {colIdx, nota, rec}` a
  `crNotas[codigo] = {colIdx, i1, i2, i3, i4, rec}` — lo consume el futuro
  plan de frontend.

- [ ] **Step 1: Escribir el test de integración (debe fallar: hoy escribe en offsets viejos)**

Añadir al final de `src-tauri/src/main.rs`:

```rust
#[cfg(test)]
mod save_notas_unidad_tests {
    use super::*;
    use std::sync::atomic::{AtomicU32, Ordering};

    static COUNTER: AtomicU32 = AtomicU32::new(0);

    // Copia el fixture real a un archivo temporal unico por test, para que cada
    // test pueda escribir sin pisar al fixture committeado ni a otros tests que
    // corran en paralelo (cargo test corre tests en threads distintos por
    // defecto).
    fn copia_fixture_temporal() -> String {
        let src = concat!(env!("CARGO_MANIFEST_DIR"), "/../Plantilla_Notas_ESO.xlsx");
        let n = COUNTER.fetch_add(1, Ordering::SeqCst);
        let dst = std::env::temp_dir().join(format!("test_save_notas_unidad_{n}_{}.xlsx", std::process::id()));
        std::fs::copy(src, &dst).expect("debe poder copiar el fixture");
        dst.to_string_lossy().to_string()
    }

    #[test]
    fn guarda_i1_i4_y_rec_en_las_columnas_correctas_y_no_toca_final() {
        let path = copia_fixture_temporal();

        // CR1.1 en U1 tiene colIdx=7 (verificado en load_notas_unidad_tests).
        // Bloque: 7=i1, 8=i2, 9=i3, 10=i4, 11=FINAL (formula, no se toca), 12=Rec.
        let payload = json!({
            "unidad": "U1",
            "syncEval": false,
            "notas": [
                { "rowIdx": 6, "crNotas": {
                    "CR1.1": { "colIdx": 7, "i1": 8.0, "i2": 7.0, "i3": 9.0, "i4": 6.0, "rec": 5.0 }
                }}
            ]
        });
        excel_save_notas_unidad_impl_with_path(&path, payload).expect("debe guardar");

        let data = load_notas_unidad(&path, "U1").expect("debe releer U1");
        let alumno = &data["alumnos"][0];
        let cr = alumno["crNotas"].as_array().unwrap().iter()
            .find(|c| c["codigo"] == "CR1.1").unwrap();
        assert_eq!(cr["i1"], 8.0);
        assert_eq!(cr["i2"], 7.0);
        assert_eq!(cr["i3"], 9.0);
        assert_eq!(cr["i4"], 6.0);
        // FINAL se recalcula en Rust con los mismos pesos que Excel (0.2/0.4/0.2/0.2):
        // (0.2*8 + 0.4*7 + 0.2*9 + 0.2*6) / 1.0 = 7.4
        assert!((cr["final"].as_f64().unwrap() - 7.4).abs() < 1e-9);
        assert_eq!(cr["recDisplay"], "5");

        std::fs::remove_file(&path).ok();
    }
}
```

Este test llama a una función `excel_save_notas_unidad_impl_with_path` que
todavía no existe — la implementación actual (`excel_save_notas_unidad_impl`)
lee la ruta del Excel de `require_selected_path()` (estado global de la app),
lo cual no es testeable en aislamiento. Extraer la variante parametrizada por
`path` es parte de este mismo step de implementación (Step 3).

- [ ] **Step 2: Ejecutar el test para verificar que falla**

Run: `cargo test --manifest-path src-tauri/Cargo.toml save_notas_unidad_tests`
Expected: FAIL — `cannot find function 'excel_save_notas_unidad_impl_with_path'`

- [ ] **Step 3: Extraer la variante parametrizada y reescribir `unit_edit_fn`**

Reemplazar la función `excel_save_notas_unidad_impl` completa (líneas
2240-2301 del archivo actual) por:

```rust
fn excel_save_notas_unidad_impl(payload: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    excel_save_notas_unidad_impl_with_path(&path, payload)
}

fn excel_save_notas_unidad_impl_with_path(path: &str, payload: Value) -> Result<Value, String> {
    let unidad = payload["unidad"].as_str().ok_or("Falta unidad")?.to_string();
    let notas = payload["notas"].as_array().ok_or("Falta notas")?.clone();
    let sync_eval = payload["syncEval"].as_bool().unwrap_or(true);
    let notas_for_unit = notas.clone();

    let unit_edit_fn: Box<dyn Fn(&str) -> Result<String, String>> = Box::new(move |xml: &str| {
        let mut s = xml.to_string();
        for nota_item in &notas_for_unit {
            if let Some(ri) = nota_item["rowIdx"].as_u64().map(|n| n as usize) {
                if let Some(cr_notas) = nota_item["crNotas"].as_object() {
                    for (_code, val_obj) in cr_notas {
                        let Some(ci) = val_obj.get("colIdx").and_then(|v| v.as_u64()).map(|n| n as usize) else { continue };
                        // Bloque de 6 columnas: ci=i1, ci+1=i2, ci+2=i3, ci+3=i4,
                        // ci+4=FINAL (formula, NUNCA se escribe), ci+5=Rec.
                        for (offset, key) in [(0usize, "i1"), (1, "i2"), (2, "i3"), (3, "i4")] {
                            if let Some(val) = val_obj.get(key) {
                                match normalize_grade(val) {
                                    Some(n) => { s = set_xml_cell(&s, ri, ci + offset, Some(&json!(n)), "number")?; }
                                    None    => { s = set_xml_cell(&s, ri, ci + offset, None, "number")?; }
                                }
                            }
                        }
                        if let Some(val) = val_obj.get("rec") {
                            match normalize_grade(val) {
                                Some(n) => { s = set_xml_cell(&s, ri, ci + 5, Some(&json!(n)), "number")?; }
                                None    => { s = set_xml_cell(&s, ri, ci + 5, None, "number")?; }
                            }
                        }
                    }
                }
            }
        }
        Ok(s)
    });

    edit_workbook_sheets_xml(path, vec![(unidad.as_str(), unit_edit_fn)])?;

    if !sync_eval {
        // Guardado ligero (autosave por celda): no releer/recalcular la unidad
        // entera para no bloquear la escritura mientras el usuario sigue
        // tecleando — mismo comportamiento que la version anterior.
        return Ok(Value::Null);
    }

    let notas_for_eval = build_notas_for_eval_sync(path, &unidad, &notas)?;
    let eval_edits = build_eval_sheet_edits(path, &unidad, &notas_for_eval)?;
    if !eval_edits.is_empty() {
        let (eval_names, eval_fns): (Vec<String>, Vec<Box<dyn Fn(&str) -> Result<String, String>>>) = eval_edits.into_iter().unzip();
        let all_eval_edits: Vec<(&str, Box<dyn Fn(&str) -> Result<String, String>>)> = eval_names.iter()
            .zip(eval_fns.into_iter())
            .map(|(name, f)| (name.as_str(), f))
            .collect();
        edit_workbook_sheets_xml(path, all_eval_edits)?;
    }

    load_notas_unidad(path, &unidad)
}
```

Nota: esta reescritura ya deja preparada la llamada a `build_notas_for_eval_sync`
(Task 5) — hasta completar el Task 5 el `cargo build` de este Task 4 fallará
por función inexistente. Para poder cerrar el Step 4 de este Task de forma
aislada, añade temporalmente un stub justo antes de esta función:

```rust
fn build_notas_for_eval_sync(_path: &str, _unidad: &str, notas: &[Value]) -> Result<Vec<Value>, String> {
    Ok(notas.to_vec()) // stub temporal — Task 5 lo reemplaza por el calculo real
}
```

También ejecuta el test del **guardado directo con `syncEval: false`**
primero (Step 4 abajo) — ese camino no llama a `build_notas_for_eval_sync` en
absoluto, así que el test de este Task pasa igual con el stub presente.

- [ ] **Step 4: Ejecutar el test para verificar que pasa**

Run: `cargo test --manifest-path src-tauri/Cargo.toml save_notas_unidad_tests`
Expected: PASS (1 test)

- [ ] **Step 5: Commit**

```bash
git add src-tauri/src/main.rs
git commit -m "feat(backend): excel_save_notas_unidad_impl escribe i1-i4/Rec en offsets del layout nuevo"
```

---

### Task 5: `build_notas_for_eval_sync` — propagar el FINAL calculado a las hojas de evaluación

**Files:**
- Modify: `src-tauri/src/main.rs` (reemplazar el stub del Task 4)
- Test: extender `mod save_notas_unidad_tests` del Task 4

**Interfaces:**
- Consumes: `compute_final_weighted` (Task 1), `read_sheet_rows`, `cell_f64`
  (sin cambios), `build_eval_sheet_edits` (línea ~2151, **sin modificar** —
  sigue esperando `notas: &[Value]` con la misma forma
  `{rowIdx, crNotas: {codigo: {nota, rec}}}` que ya consumía).
- Produces: `fn build_notas_for_eval_sync(path: &str, unidad: &str, notas: &[Value]) -> Result<Vec<Value>, String>`
  ya queda cableada en `excel_save_notas_unidad_impl_with_path` (Task 4).

**Contexto:** `build_eval_sheet_edits` no sabe nada de instrumentos — solo
necesita, por `(alumno, código)`, un `nota` (el valor a cachear en la celda
CR de la hoja de evaluación) y opcionalmente un `rec`. Antes ese `nota` era
la nota directa que el profesor tecleaba. Ahora es el `FINAL` calculado a
partir de i1..i4. Como `edit_workbook_sheets_xml` para la hoja de unidad ya
se ejecutó (Task 4, antes de esta llamada), releer la hoja aquí ve los
valores i1..i4 recién guardados.

**Importante — semántica de "borrar":** si el payload trae una entrada para
un criterio (aunque sea para vaciar sus 4 instrumentos), el `FINAL`
recalculado se debe propagar **siempre** (incluso si da `None`, que debe
serializarse como JSON `null` para que `build_eval_sheet_edits` limpie la
celda cacheada en la hoja de evaluación en vez de dejar un valor viejo
obsoleto). `rec` se pasa tal cual venía en el payload original (no se
recalcula, es un valor manual).

- [ ] **Step 1: Escribir el test (debe fallar: el stub del Task 4 no calcula nada)**

Añadir dentro de `mod save_notas_unidad_tests` (mismo módulo del Task 4), dos
tests nuevos:

```rust
    #[test]
    fn sincroniza_el_final_calculado_hacia_las_hojas_de_evaluacion() {
        let path = copia_fixture_temporal();

        let payload = json!({
            "unidad": "U1",
            "syncEval": true,
            "notas": [
                { "rowIdx": 6, "crNotas": {
                    "CR1.1": { "colIdx": 7, "i1": 8.0, "i2": 7.0, "i3": 9.0, "i4": 6.0 }
                }}
            ]
        });
        let result = excel_save_notas_unidad_impl_with_path(&path, payload).expect("debe guardar y sincronizar");

        // load_notas_unidad tras el guardado confirma que el FINAL persistido es 7.4
        let alumno = &result["alumnos"][0];
        let cr = alumno["crNotas"].as_array().unwrap().iter()
            .find(|c| c["codigo"] == "CR1.1").unwrap();
        assert!((cr["final"].as_f64().unwrap() - 7.4).abs() < 1e-9);

        std::fs::remove_file(&path).ok();
    }

    #[test]
    fn build_notas_for_eval_sync_pasa_null_al_vaciar_todos_los_instrumentos() {
        let path = copia_fixture_temporal();

        // Primero rellena CR1.1 para tener algo que vaciar despues.
        let relleno = json!({ "unidad": "U1", "syncEval": false, "notas": [
            { "rowIdx": 6, "crNotas": { "CR1.1": { "colIdx": 7, "i1": 8.0, "i2": 7.0, "i3": 9.0, "i4": 6.0 } } }
        ]});
        excel_save_notas_unidad_impl_with_path(&path, relleno).expect("relleno inicial");

        // Vacia los 4 instrumentos (equivalente a borrar el contenido de las celdas).
        let vacio = json!({ "unidad": "U1", "syncEval": false, "notas": [
            { "rowIdx": 6, "crNotas": { "CR1.1": { "colIdx": 7, "i1": null, "i2": null, "i3": null, "i4": null } } }
        ]});
        excel_save_notas_unidad_impl_with_path(&path, vacio).expect("vaciado");

        let notas_eval = build_notas_for_eval_sync(&path, "U1", &json!([
            { "rowIdx": 6, "crNotas": { "CR1.1": { "colIdx": 7 } } }
        ]).as_array().unwrap().to_vec()).expect("debe calcular sync");

        let entry = &notas_eval[0]["crNotas"]["CR1.1"];
        assert!(entry["nota"].is_null(), "debe propagar null para limpiar el cache de la hoja de evaluacion: {entry:?}");

        std::fs::remove_file(&path).ok();
    }
```

- [ ] **Step 2: Ejecutar los tests para verificar que fallan**

Run: `cargo test --manifest-path src-tauri/Cargo.toml save_notas_unidad_tests`
Expected: FAIL en los 2 tests nuevos (el stub devuelve las notas sin
recalcular `final`, así que `nota` en el payload sincronizado sigue sin
existir o no vale lo esperado).

- [ ] **Step 3: Reemplazar el stub por la implementación real**

Sustituir el stub temporal del Task 4 por:

```rust
// Tras guardar la hoja de unidad, relee los valores i1..i4 ya persistidos y
// calcula el FINAL de cada criterio tocado en este guardado, para
// propagarlo al cache de las hojas de evaluacion via build_eval_sheet_edits
// (que no sabe nada de instrumentos, solo espera {nota, rec} por criterio).
fn build_notas_for_eval_sync(path: &str, unidad: &str, notas: &[Value]) -> Result<Vec<Value>, String> {
    let rows = read_sheet_rows(path, unidad).map_err(|_| format!("No se encontró la hoja \"{unidad}\"."))?;
    let weights: [f64; 4] = [
        cell_f64(&rows, 3, 2).unwrap_or(0.0),
        cell_f64(&rows, 3, 3).unwrap_or(0.0),
        cell_f64(&rows, 3, 4).unwrap_or(0.0),
        cell_f64(&rows, 3, 5).unwrap_or(0.0),
    ];

    let out: Vec<Value> = notas.iter().filter_map(|nota_item| {
        let ri = nota_item["rowIdx"].as_u64()? as usize;
        let cr_notas = nota_item["crNotas"].as_object()?;
        let mut cr_obj = serde_json::Map::new();
        for (code, val_obj) in cr_notas {
            let ci = val_obj.get("colIdx")?.as_u64()? as usize;
            let values = [
                cell_f64(&rows, ri, ci),
                cell_f64(&rows, ri, ci + 1),
                cell_f64(&rows, ri, ci + 2),
                cell_f64(&rows, ri, ci + 3),
            ];
            let final_val = compute_final_weighted(values, weights);
            let mut entry = serde_json::Map::new();
            entry.insert("colIdx".to_string(), json!(ci));
            entry.insert("nota".to_string(), json!(final_val));
            if let Some(rec_val) = val_obj.get("rec") {
                entry.insert("rec".to_string(), rec_val.clone());
            }
            cr_obj.insert(code.clone(), Value::Object(entry));
        }
        Some(json!({ "rowIdx": ri, "crNotas": Value::Object(cr_obj) }))
    }).collect();
    Ok(out)
}
```

- [ ] **Step 4: Ejecutar todos los tests del módulo para verificar que pasan**

Run: `cargo test --manifest-path src-tauri/Cargo.toml save_notas_unidad_tests`
Expected: PASS (3 tests: el del Task 4 + los 2 nuevos)

- [ ] **Step 5: `cargo build` completo**

Run: `cargo build --manifest-path src-tauri/Cargo.toml`
Expected: compila sin errores ni warnings nuevos.

- [ ] **Step 6: Commit**

```bash
git add src-tauri/src/main.rs
git commit -m "feat(backend): build_notas_for_eval_sync propaga el FINAL calculado a las hojas de evaluacion"
```

---

### Task 6: comando `excel_save_unidad_instrumentos`

**Files:**
- Modify: `src-tauri/src/main.rs` (nueva función + comando, cerca de
  `excel_save_notas_unidad`, y registro en `generate_handler!` ~línea 3102)
- Modify: `app-bridge.js:34` (añadir wrapper junto a `saveNotasUnidad`)
- Test: nuevo módulo `#[cfg(test)] mod unidad_instrumentos_tests`

**Interfaces:**
- Consumes: `load_instrumentos(path) -> Result<Value, String>` (línea 2881,
  sin cambios — catálogo `DATOS!N5:O13`), `set_xml_cell`,
  `edit_workbook_sheets_xml`, `require_selected_path`.
- Produces: comando Tauri `excel_save_unidad_instrumentos`, payload
  `{ unidad: String, slots: [{abrev: String, peso: f64}] }` (1 a 4 slots),
  usado por el futuro plan de frontend (`gestor-unidades.html`).

- [ ] **Step 1: Escribir los tests de validación (deben fallar: la función no existe aún)**

Añadir al final de `src-tauri/src/main.rs`:

```rust
#[cfg(test)]
mod unidad_instrumentos_tests {
    use super::*;

    #[test]
    fn rechaza_si_los_pesos_no_suman_100_por_ciento() {
        let slots = vec![
            json!({"abrev": "PE", "peso": 0.5}),
            json!({"abrev": "TD", "peso": 0.3}),
        ];
        let catalogo = vec!["PE".to_string(), "TD".to_string(), "TI".to_string()];
        let result = validate_unidad_instrumentos(&slots, &catalogo);
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("100"), "el mensaje debe explicar que falta sumar 100%");
    }

    #[test]
    fn acepta_pesos_que_suman_100_por_ciento() {
        let slots = vec![
            json!({"abrev": "PE", "peso": 0.2}),
            json!({"abrev": "TD", "peso": 0.4}),
            json!({"abrev": "TI", "peso": 0.2}),
            json!({"abrev": "TG", "peso": 0.2}),
        ];
        let catalogo = vec!["PE".to_string(), "TD".to_string(), "TI".to_string(), "TG".to_string()];
        assert!(validate_unidad_instrumentos(&slots, &catalogo).is_ok());
    }

    #[test]
    fn rechaza_abreviatura_que_no_esta_en_el_catalogo() {
        let slots = vec![json!({"abrev": "ZZ", "peso": 1.0})];
        let catalogo = vec!["PE".to_string(), "TD".to_string()];
        let result = validate_unidad_instrumentos(&slots, &catalogo);
        assert!(result.is_err());
        assert!(result.unwrap_err().contains("ZZ"));
    }

    #[test]
    fn rechaza_mas_de_4_slots() {
        let slots: Vec<Value> = (0..5).map(|_| json!({"abrev": "PE", "peso": 0.2})).collect();
        let catalogo = vec!["PE".to_string()];
        assert!(validate_unidad_instrumentos(&slots, &catalogo).is_err());
    }
}
```

- [ ] **Step 2: Ejecutar los tests para verificar que fallan**

Run: `cargo test --manifest-path src-tauri/Cargo.toml unidad_instrumentos_tests`
Expected: FAIL — `cannot find function 'validate_unidad_instrumentos'`

- [ ] **Step 3: Implementar la validación, el comando y el registro**

Añadir en `src-tauri/src/main.rs`, después de `build_notas_for_eval_sync`
(Task 5):

```rust
// Valida el payload de excel_save_unidad_instrumentos: maximo 4 slots (limite
// estructural de columnas del template), cada abreviatura debe existir en el
// catalogo de Instrumentos (DATOS!N5:O13), y los pesos deben sumar 100%
// (tolerancia 0.001 por redondeos de coma flotante).
fn validate_unidad_instrumentos(slots: &[Value], catalogo_abrevs: &[String]) -> Result<(), String> {
    if slots.len() > 4 {
        return Err(format!("Máximo 4 instrumentos por unidad, se han recibido {}.", slots.len()));
    }
    let mut suma = 0.0;
    for slot in slots {
        let abrev = slot["abrev"].as_str().unwrap_or("");
        if !catalogo_abrevs.iter().any(|a| a == abrev) {
            return Err(format!("El instrumento \"{abrev}\" no existe en el catálogo de Instrumentos."));
        }
        suma += slot["peso"].as_f64().unwrap_or(0.0);
    }
    if (suma - 1.0).abs() > 0.001 {
        return Err(format!("Los pesos deben sumar 100% (suman {:.1}%).", suma * 100.0));
    }
    Ok(())
}

fn excel_save_unidad_instrumentos_impl(payload: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let unidad = payload["unidad"].as_str().ok_or("Falta unidad")?.to_string();
    let slots = payload["slots"].as_array().ok_or("Falta slots")?.clone();

    let catalogo = load_instrumentos(&path)?;
    let catalogo_abrevs: Vec<String> = catalogo["instrumentos"].as_array().unwrap_or(&vec![])
        .iter().filter_map(|i| i["codigo"].as_str().map(|s| s.to_string())).collect();
    validate_unidad_instrumentos(&slots, &catalogo_abrevs)?;

    let slots_owned = slots.clone();
    edit_workbook_sheets_xml(&path, vec![(unidad.as_str(), Box::new(move |xml: &str| {
        let mut s = xml.to_string();
        for i in 0..4usize {
            let ci = 2 + i; // columnas C:F (0-idx 2..6)
            match slots_owned.get(i) {
                Some(slot) => {
                    let abrev = slot["abrev"].as_str().unwrap_or("");
                    let peso = slot["peso"].as_f64().unwrap_or(0.0);
                    s = set_xml_cell(&s, 1, ci, Some(&json!(abrev)), "text")?; // fila 2 (0-idx 1)
                    s = set_xml_cell(&s, 3, ci, Some(&json!(peso)), "number")?; // fila 4 (0-idx 3)
                }
                None => {
                    s = set_xml_cell(&s, 1, ci, None, "text")?;
                    s = set_xml_cell(&s, 3, ci, None, "number")?;
                }
            }
        }
        Ok(s)
    }) as Box<dyn Fn(&str) -> Result<String, String>>)])?;

    load_notas_unidad(&path, &unidad)
}

#[tauri::command]
async fn excel_save_unidad_instrumentos(payload: Value) -> Result<Value, String> {
    tauri::async_runtime::spawn_blocking(move || excel_save_unidad_instrumentos_impl(payload))
        .await
        .map_err(|e| e.to_string())?
}
```

Registrar el comando en `generate_handler!` (línea 3102 del archivo actual):
buscar la línea

```rust
            excel_get_notas_unidad, excel_save_notas_unidad, excel_resync_unidad_eval, excel_get_alumnos_informes, app_open_external,
```

y añadir `excel_save_unidad_instrumentos` al final de esa lista:

```rust
            excel_get_notas_unidad, excel_save_notas_unidad, excel_save_unidad_instrumentos, excel_resync_unidad_eval, excel_get_alumnos_informes, app_open_external,
```

En `app-bridge.js`, añadir junto a `saveNotasUnidad` (línea 34):

```javascript
    saveUnidadInstrumentos: (payload) => invoke("excel_save_unidad_instrumentos", { payload }),
```

- [ ] **Step 4: Ejecutar los tests para verificar que pasan**

Run: `cargo test --manifest-path src-tauri/Cargo.toml unidad_instrumentos_tests`
Expected: PASS (4 tests)

- [ ] **Step 5: Commit**

```bash
git add src-tauri/src/main.rs app-bridge.js
git commit -m "feat(backend): comando excel_save_unidad_instrumentos para editar instrumentos/pesos por unidad"
```

---

### Task 7: verificación final completa

**Files:** ninguno nuevo — solo verificación.

- [ ] **Step 1: `cargo test` completo**

Run: `cargo test --manifest-path src-tauri/Cargo.toml`
Expected: todos los tests pasan (los preexistentes de `formula_cache_tests` y
`eval_layout_tests` + los añadidos en Tasks 1-6).

- [ ] **Step 2: `cargo build --release` para confirmar que el binario final compila**

Run: `cargo build --release --manifest-path src-tauri/Cargo.toml`
Expected: build limpio, sin errores.

- [ ] **Step 3: Confirmar que no quedan archivos temporales de test sin limpiar**

Run (PowerShell): `Get-ChildItem $env:TEMP -Filter "test_save_notas_unidad_*"`
Expected: sin resultados (cada test borra su copia al terminar con
`std::fs::remove_file`). Si aparecen restos de una ejecución anterior
interrumpida, bórralos a mano — no forman parte del repo.

- [ ] **Step 4: Commit final de cierre (si algo quedó sin commitear)**

```bash
git status --short
# si hay cambios pendientes:
git add -A
git commit -m "chore(backend): cierre plan instrumentos por criterio"
```

---

## Qué queda fuera de este plan

- Frontend (`gestor-notas.html`, `gestor-unidades.html`,
  `gestor-recuperaciones.html`, `visor-unidades.html`) — plan aparte, a
  escribir después de completar este, ya con el contrato JSON de
  `load_notas_unidad`/`excel_save_notas_unidad`/`excel_save_unidad_instrumentos`
  estable y probado.
- README.md, CLAUDE.md y `memory/` del proyecto — se actualizan en una tarea
  de documentación separada antes del commit/push final a GitHub (pedido
  explícito del usuario), no dentro de este plan de implementación.
