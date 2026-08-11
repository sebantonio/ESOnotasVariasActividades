# SDD ledger — plan: docs/superpowers/plans/2026-08-11-ponderaciones-editables.md

## Tasks

- [x] Task 1: Sección "Configurar Ponderaciones" en gestor-unidades.html
- [x] Task 2: Cargar unidades en dropdown + validación en tiempo real
- [x] Task 3: Backend - Comandos IPC para cargar/guardar pesos
- [x] Task 4: Guardar ponderaciones - click handler
- [x] Task 5: Integrar pesos en cálculo FINAL de gestor-notas.html
- [ ] Task 6: Testing Manual

## Progress

### Task 6: Testing Manual
- Status: BLOCKED
- Issue: Critical integration bug found
- Root cause: Missing IPC bridge method definitions
- Impact: All interactive tests blocked (0/6 test groups executable)
- Evidence: HTML calls `excel_get_unidad_pesos()` and `excel_save_unidad_pesos()` but these methods are not exposed in app-bridge.js
- Recommendation: Add missing methods to app-bridge.js and retry testing

### Detailed Findings
- Rust backend: ✓ Complete and correct
- HTML frontend: ✓ Structure and logic correct
- JavaScript bridge: ✗ Missing 2 method definitions
- Pre-test verification: Uncovered integration gap between Tasks 3-4 and frontend

