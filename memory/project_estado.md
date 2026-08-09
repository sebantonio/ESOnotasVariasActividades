---
name: Estado del proyecto ESOnotasVariasActividades
description: Estado actualizado 2026-08-09 — nuevo layout con instrumentos por criterio, 7 tasks backend completadas + fix post-review
type: project
---

**Estado (2026-08-09, sesión 7 - Proyecto ESOnotasVariasActividades):**
- **Proyecto bifurcado**: ESOnotasVariasActividades es una continuación de ESOplantillaNotas que adapta la app para una nueva estructura de Excel donde cada criterio se evalúa a través de hasta 4 instrumentos (i1-i4) con pesos configurables, en lugar de una nota directa.
- **7 tasks backend completadas**:
  - Task 1: `compute_final_weighted(i1..i4, pesos) → final` — media ponderada, ignora vacíos, con 5 tests
  - Task 2: `cr_cols_from_cells()` — generaliza detección de criterios para nuevas filas (5→6, offset de 2→6 cols/criterio)
  - Task 3: `load_notas_unidad()` refactorizado — lee i1-i4 + final + rec, instrumentosUnidad (abrev/pesos), filas comenzando en 6 (no 4)
  - Task 4: `excel_save_notas_unidad_impl()` escriba i1-i4 en cols ci..ci+3, rec en ci+5, **nunca** en ci+4 (FINAL es fórmula)
  - Task 5: `build_notas_for_eval_sync()` — recalcula FINAL en Rust tras guardar i1-i4, propaga a hojas de evaluación
  - Task 6: Nuevo comando Tauri `excel_save_unidad_instrumentos(unidad, slots)` — configura instrumentos/pesos por unidad
  - Task 7: Verificación final — 30 tests all passing (fue 29, +1 test post-review), `cargo build --release` limpio
- **Fix post-review (Important issue)**:
  - `excel_resync_unidad_eval_impl` leía campo obsoleto `cr["nota"]` que cambió a `cr["final"]` en Task 3
  - Fix: línea 2502 cambiada a `cr["final"]`, test de regresión `resync_unidad_eval_propaga_el_final_leyendo_el_campo_final_no_nota` agregado
  - Resultado: 30/30 tests passing, todo pusheado a origin/master (commit 71cf4df)
- **README actualizado** con nueva sección "Evaluación por instrumentos" explicando flujo i1-i4 → FINAL → evaluaciones
- **Versión**: 0.1.204 (bumped automáticamente durante build)

**Cambios en estructura Excel (hojas Uxx):**
- Antes: fila 3=header CR+Rec (2 cols/criterio), fila 4=fórmulas, fila 5+=datos, students empiezan en row 5 (0-idx 4)
- Ahora: fila 2=labels (i1/i2/i3/i4 compartidos), fila 3=subheader, fila 4=pesos, fila 5=códigos, fila 6=fórmulas, fila 7+=datos, students comienzan en row 7 (0-idx 6), 6 cols/criterio (ci..ci+5)

**How to apply:** 
- Plantilla Excel debe usar `Plantilla_Notas_ESO.xlsx` (100 criterios en U1 de prueba, estructura new layout)
- Frontend aún no reescrito (gestor-notas.html, visor-unidades.html, gestor-unidades.html para UI instrumentos/pesos siguen con layout anterior) — pendiente
- Backend 100% listo, tests verifican correctitud con fixture real committed

**Estado (2026-08-06, sesión 6):**
- **Bug crítico corregido**: guardar Rec en recuperaciones vaciaba por completo las hojas 2ª/3ª EVA. Causa raíz: `set_xml_formula_cache_number` no reconocía celdas con `<v/>` self-closing (cache vacío típico de `IFERROR(...,"")`) y añadía un `<v>` duplicado → XML inválido → Excel reparaba vaciando la hoja. Fix + 9 tests de regresión (`formula_cache_tests`, `eval_layout_tests`). **Why:** confirmado con el Excel real del usuario (`CCGG PLANTILLA - CE AMPLIADOSv3.xlsx`), tenía backup y no hubo pérdida de datos.
- **Bug de congelado corregido**: comandos Tauri síncronos (`excel_save_notas_unidad` y otros 6 de guardado) bloqueaban la ventana en guardados pesados (~4MB xlsx). Pasados a `async fn` + `tauri::async_runtime::spawn_blocking`.
- **Backup automático**: `edit_workbook_sheets_xml`/`ensure_diario_sheet` escriben un `.bak` best-effort justo antes de sobrescribir el xlsx.
- **Limpieza**: borrados `main.js`/`preload.js`/`tauri-node-backend.js` (Electron legacy, confirmado sin uso — ver pendiente #6, ya resuelto) y limpiado `package.json` (sin scripts/deps de Electron). CI básica añadida en `.github/workflows/ci.yml` (`cargo build --release` + `cargo test`).
- **Rediseño visual completo (Hallmark)**: `design.md` en la raíz del repo es ahora la fuente de verdad del sistema de diseño (paleta OKLCH sobre el indigo/teal ya existente, tipografía solo de sistema — offline, sin CDN de fuentes —, espaciado 4pt, motion con easings nombrados). `ux-common.css` tiene los tokens nuevos (`--color-*`, `--font-*`, etc.); `--ux-*` quedan como alias. Las 14 páginas que comparten `ux-common.css` se unificaron (antes cada una usaba un tono de indigo distinto suelto, y `incluir-actividad.html` tenía un gradiente morado propio). **How to apply:** cualquier cambio visual futuro debe leer `design.md` primero y usar los tokens `--color-*`/`--font-*`, no hex sueltos. `asteroides.html` queda fuera de este sistema (no usa `ux-common.css`).
- Versión progresó 0.1.191 → 0.1.196 en esta sesión.
- Repo tiene además `.agents/`, `.claude/skills/`, `skills-lock.json` sin trackear, de origen ajeno a esta sesión (no tocados).

**Estado (2026-08-03, sesión 5):**
- Añadido branding propio: `notasESOicon.png` (raíz repo) sustituye el icono heredado del proyecto FP original. Regenerados todos los iconos de `src-tauri/icons/` con `npx tauri icon notasESOicon.png`, referenciados en `tauri.conf.json` → `bundle.icon`. Logo de cabecera de `index.html` (`.brand-mark`) ahora usa la imagen en vez de texto "ESO".
- Nueva función: botón "Descargar plantilla Excel" en `index.html` → comando Rust `excel_download_template` que copia `Plantilla_Notas_ESO.xlsx` (embebida en el binario vía `include_bytes!`) a la ruta que elija el usuario. Requirió excepción en `.gitignore` (`!Plantilla_Notas_ESO.xlsx`) porque el patrón `*.xlsx` lo ignoraba.
- `scripts/prepare-tauri-web.js` ahora copia también archivos binarios (PNG) a `tauri-web/`, antes solo trataba HTML/JS/CSS como texto UTF-8.
- Toolchain Rust no estaba instalado en esta máquina; instalado con `rustup-init.exe -y` (el instalador de winget se colgaba sin `-y`, exit code 3221225786 = STATUS_CONTROL_C_EXIT). Disco C: tenía solo ~130MB libres, insuficiente — se reformateó D: (exFAT vacío → NTFS, exFAT no soporta hardlinks que rustup necesita para sus proxies). Toolchain queda en `RUSTUP_HOME=D:\rust\rustup`, `CARGO_HOME=D:\rust\cargo`. **Why:** sin esto no compila nada con cargo en esta máquina — variables de entorno persistidas a nivel de usuario (`setx`/`[Environment]::SetEnvironmentVariable(...,"User")`), deberían sobrevivir a reinicio de sesión.

**How to apply:** en sesiones futuras en esta máquina, verificar `cargo --version` antes de asumir que hace falta reinstalar Rust — ya debería estar en D:\rust\cargo\bin. Si `*.xlsx` se sigue tocando en `.gitignore`, recordar que `Plantilla_Notas_ESO.xlsx` tiene excepción explícita y es necesaria para compilar (`include_bytes!` en main.rs).
Proyecto creado el 2026-05-10 replicando la estructura completa de plantillaNotas para el Excel ESO `CCGG PLANTILLA - RECUv45.xlsx`. La app compila y funciona.

**Why:** El usuario quiere una app Tauri para gestionar notas ESO sobre su Excel propio.

**Estado actual (2026-05-10):**
- App compila y genera EXE en `C:\cargo-target\plantillaNotas\release\bundle\nsis\ESO Notas Local_0.1.0_x64-setup.exe`
- Excel analizado: hoja DATOS, tabla de alumnos en A4:B30 (A=nº, B=nombre), unidades en I5:K20 (I=código, J=nombre, K=evaluación), instrumentos en N4:O13 (N=abrev, O=nombre)
- Todos los gestores funcionan: alumnos, unidades, instrumentos, CE+criterios
- index.html adaptado a ESO: branding, botones, modal de importación CSV

**Cambios clave aplicados a main.rs:**
- `load_alumnos`: lee A4:B30 (cols 0/1, filas 3-29)
- `load_unidades`: lee I5:K20 FIJO (cols 8/9/10, filas 4-19) — sin buscar headers
- `save_unidades_to_file`: escribe I5:K20 fijo (cols 8/9/10)
- `load_instrumentos`: lee N4:O13 (cols 13/14, filas 3-12, max 10)
- `save_instrumentos_to_file`: escribe N4:O13 fijo
- Eliminado: toda lógica de "buscar UNIDADES" por header — la tabla es fija

**Cambios en HTML:**
- index.html: branding ESO, modal importar (CSV → Excel activo), botón Instrumentos de Evaluación
- gestor-unidades.html: sin columna Horas, lee I5:K20 directo
- gestor-instrumentos.html: max 10, reescrito en ASCII puro (sin emojis ni tildes)
- gestor-rraa-criterios.html: RA→CE en toda la UI, eliminadas columnas Instituto/Empresa
- app-bridge.js: añadidos getInstrumentos/saveInstrumentos
- scripts/prepare-tauri-web.js: añadido gestor-instrumentos.html a la lista de copia

**How to apply:** La estructura del Excel ESO es diferente al FP. Siempre usar rangos fijos (I5:K20, N4:O13, A4:B30) — no buscar por contenido de celdas.
