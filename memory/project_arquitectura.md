---
name: Arquitectura técnica ESOplantillaNotas
description: Stack Tauri v2 + Rust, estructura Excel ESO, comandos IPC
type: project
---
**Stack:** Tauri v2 + Rust backend puro + HTML/CSS/Vanilla JS frontend.

**Crates Rust:**
- `calamine 0.26` (features: dates) — leer hojas XLSX como Vec<Vec<Value>>
- `zip 2` (features: deflate) — abrir/reescribir el ZIP interno del XLSX
- `regex 1` — manipulación XML de hojas
- `once_cell 1` + `chrono 0.4` — estado global y fechas seriales Excel

**Comandos Tauri registrados en main.rs:**
excel_select_file, excel_get_selected_file, excel_set_selected_file, excel_verify_file_exists,
excel_get_alumnos, excel_save_alumnos,
excel_get_unidades, excel_save_unidades,
excel_get_rraa_criterios, excel_save_rraa_criterios,
excel_get_instrumentos, excel_save_instrumentos,
excel_get_notas_actividad, excel_get_notas_actividades_tipo, excel_save_notas_actividad,
excel_add_actividad, excel_save_ce_notas,
excel_get_notas_evaluacion, excel_get_notas_evaluacion_alumno,
excel_get_alumnos_informes, excel_get_notas_unidad,
excel_get_diario, excel_save_diario_entrada, excel_delete_diario_entrada,
app_open_external

**Estructura Excel ESO — hoja DATOS (rangos fijos):**
- Alumnos: A4:B41 — fila 4=header "Alumnado", datos filas 5-41 (0-idx:4-40), max 37; Rust busca header dinámicamente
- Unidades: I5:K20 → col 8=código, col 9=nombre, col 10=evaluación (filas 4-19 en 0-indexed)
- Instrumentos: N4:O13 → col 13=abreviatura, col 14=nombre (filas 3-12 en 0-indexed, max 10)

**Hoja PESOS:** CE (criterios de evaluación) y sus ponderaciones por unidad. Códigos tipo CR1.1, CR2.3...

**Escritura XML:** Se abre el XLSX como ZIP, se edita el XML de cada hoja con regex (`set_xml_cell`), se elimina calcChain, se fuerza forceFullCalc=1, se reescribe el ZIP.

**Patrón IPC:** HTML → app-bridge.js → `window.__TAURI__.core.invoke("comando", {args})` → Rust fn → Result<Value, String>

**How to apply:** main.rs es la única fuente de verdad para lógica Excel. Los rangos son FIJOS — nunca buscar por contenido de celda para alumnos/unidades/instrumentos.
