# Task 6: Testing Manual - REPORT

**Date**: 2026-08-11  
**Status**: BLOCKED - Critical Integration Bug Found  
**Tester**: Claude Code Agent

---

## Executive Summary

Testing of the weighting interface (ponderaciones) is **BLOCKED** due to a critical integration bug discovered during pre-test verification. The HTML frontend is calling methods that do not exist in the app-bridge.js IPC bridge, preventing any UI interaction from reaching the backend.

**Result**: ALL TESTS BLOCKED (0/6 test groups can execute)

---

## Critical Issue Found: Missing IPC Bridge Methods

### Problem Description

The `tauri-web/gestor-unidades.html` file contains code that attempts to call methods via `window.electronExcel` that are not exposed in `app-bridge.js`:

**Methods Called by HTML but NOT in Bridge:**
1. `window.electronExcel.excel_get_unidades()` (line 1325)
2. `window.electronExcel.excel_get_unidad_pesos(...)` (line 1343)  
3. `window.electronExcel.excel_save_unidad_pesos(...)` (line 1407)

### Root Cause

**File**: `app-bridge.js` (lines 14-48)

The bridge exposes methods using camelCase naming (e.g., `getUnidades`, `getInstrumentos`), but the ponderaciones section HTML uses snake_case with `excel_` prefix (e.g., `excel_get_unidades`).

**Current app-bridge.js methods** (sample):
```javascript
window.electronExcel = {
  getUnidades: () => invoke("excel_get_unidades"),
  getInstrumentos: () => invoke("excel_get_instrumentos"),
  getNotasUnidad: (payload) => invoke("excel_get_notas_unidad", { payload }),
  ...
  // Missing: excel_get_unidad_pesos, excel_save_unidad_pesos
}
```

**HTML calling pattern** (tauri-web/gestor-unidades.html):
```javascript
// Line 1325 - Uses wrong method name (should be getUnidades)
const data = await window.electronExcel.excel_get_unidades();

// Line 1343 - Method doesn't exist in bridge
const data = await window.electronExcel.excel_get_unidad_pesos({unidad});

// Line 1407 - Method doesn't exist in bridge  
await window.electronExcel.excel_save_unidad_pesos({ unidad, pesos });
```

### Impact

When any of the following user actions occur, JavaScript will throw an error:
1. **UI Toggle Test**: Section display should work (no IPC call)
2. **Dropdown Test**: FAILS on line 1343 - `excel_get_unidad_pesos is not a function`
3. **Validation Test**: Can partially proceed (validation is client-side only)
4. **Save Test**: FAILS on line 1407 - `excel_save_unidad_pesos is not a function`
5. **Integration Test**: FAILS - cannot save weights
6. **All tests dependent on backend IPC**: BLOCKED

### Code Evidence

**File locations**:
- Bridge definition: `app-bridge.js:14-48`
- HTML calls to missing methods: `tauri-web/gestor-unidades.html:1325,1343,1407`

**Console error that will occur**:
```
Uncaught TypeError: window.electronExcel.excel_get_unidad_pesos is not a function
  at HTMLSelectElement.<anonymous> (gestor-unidades.html:1343)
```

---

## Test Plan Execution Status

### Test 1: UI Toggle Test ✗ PARTIAL (no backend call)
**Status**: Can verify visually, but dependent test 2 will fail

**Expected**: Click "Configurar Ponderaciones" button → section shows/hides  
**Result**: Button and section exist in HTML ✓  
**Blocker**: No JavaScript error expected on toggle (section display is pure DOM manipulation)

**Code verified**:
- Button element exists: `togglePonderacionesBtn` (line 549)
- Section element exists: `ponderacionesSection` (line 590)
- Toggle listener registered: lines 711-717
  ```javascript
  togglePonderacionesBtn.addEventListener('click', () => {
    const isVisible = ponderacionesSection.style.display !== 'none';
    ponderacionesSection.style.display = isVisible ? 'none' : 'block';
    if (!isVisible) loadUnidadesForPonderaciones();  // CALLS MISSING METHOD
  });
  ```
**Critical Issue**: Line 716 calls `loadUnidadesForPonderaciones()` which calls the missing method

**Verdict**: ✗ BLOCKED - will fail when section is opened

### Test 2: Dropdown Test ✗ BLOCKED
**Status**: Cannot execute - requires missing `excel_get_unidad_pesos` method

**Expected**: Click "Configurar Ponderaciones" → dropdown populates with units → select a unit → weights table shows

**What will happen**:
1. Click button → section displays ✓
2. Call `loadUnidadesForPonderaciones()` → calls `window.electronExcel.excel_get_unidades()`
3. **ERROR**: TypeError - method doesn't exist

**Code location**: Line 1325
```javascript
const data = await window.electronExcel.excel_get_unidades();  // NOT DEFINED
```

**Verdict**: ✗ BLOCKED - JavaScript error prevents execution

### Test 3: Validation Test ✗ PARTIAL (client-side only)
**Status**: Can verify validation logic in isolation, but cannot test full flow

**What can be tested**:
- Input fields exist: `peso_i1`, `peso_i2`, `peso_i3`, `peso_i4` ✓
- Validation function exists: `actualizarValidacion()` ✓
- Validation logic:
  ```javascript
  const suma = i1 + i2 + i3 + i4;
  if (Math.abs(suma - 1.0) <= 0.001) {
    // Valid - shows green message, enables button
  } else {
    // Invalid - shows red message, disables button  
  }
  ```
- Message display logic verified ✓
- Button enable/disable logic verified ✓

**What cannot be tested**:
- Loading initial weights from backend (blocked - missing method)
- Persistence after save (blocked - missing method)
- Full user flow from load → edit → validate → save

**Test cases verified via code inspection**:
| Weights | Sum | Expected | Code Correct? |
|---------|-----|----------|---------------|
| 0.5, 0.3, 0.1, 0.1 | 1.0 | Valid ✓ | Yes ✓ |
| 0.4, 0.3, 0.1, 0.1 | 0.9 | Invalid ✗ | Yes ✓ |
| 0.6, 0.3, 0.1, 0.1 | 1.1 | Invalid ✗ | Yes ✓ |

**Verdict**: ✓ PARTIAL - Validation logic is correct, but UI flow blocked by missing methods

### Test 4: Save Test ✗ BLOCKED
**Status**: Cannot execute - requires missing `excel_save_unidad_pesos` method

**Expected**: Valid weights → click save → success message → section closes → weights persist

**What will happen**:
1. Validation passes (client-side) ✓
2. Click save button
3. Handler executes line 1407:
   ```javascript
   await window.electronExcel.excel_save_unidad_pesos({ unidad, pesos });
   ```
4. **ERROR**: TypeError - method doesn't exist

**Verdict**: ✗ BLOCKED - JavaScript error prevents save

### Test 5: Integration Test ✗ BLOCKED
**Status**: Cannot execute - requires prior save (blocked) and backend calculation

**Expected**: Save weights → open gestor-notas → enter instrument grades → verify FINAL uses new weights

**Dependencies**:
- Test 4 (Save) - BLOCKED
- Backend must use weights in calculation

**Verdict**: ✗ BLOCKED - Upstream test (Test 4) is blocked

### Test 6: Persistence Test ✗ BLOCKED  
**Status**: Cannot execute - requires successful save (blocked)

**Expected**: Save weights → close/reopen page → weights still there

**Verdict**: ✗ BLOCKED - Upstream save is blocked

---

## Backend Implementation Verification

### Rust Commands Status: ✓ COMPLETE
The Rust backend handlers **ARE correctly implemented**:

**File**: `src-tauri/src/main.rs:3149-3214`

✓ `excel_get_unidad_pesos` (lines 3149-3170)
- Reads weights from unit sheet row 5, columns C:F
- Defaults to [0.25, 0.25, 0.25, 0.25]
- Returns: `{ "pesos": [f64; 4] }`

✓ `excel_save_unidad_pesos` (lines 3173-3214)
- Validates: exactly 4 weights, all numbers, sum = 1.0 ± 0.001
- Writes to row 5, columns C:F via XML editing
- Returns: `{ "pesos": [f64; 4] }` or error

✓ Both registered in `tauri::generate_handler!` macro (line 3652)

**Issue**: Handlers exist in Rust, but are **not exposed** in the JavaScript bridge

---

## Frontend Implementation Status

### HTML Structure: ✓ COMPLETE
- UI elements all present
- Section styling correct
- Input fields properly configured
- Buttons with correct IDs

### JavaScript Logic: ✓ COMPLETE (but unreachable)
- Validation logic correct (lines 1356-1384)
- Event listeners attached (lines 1387-1389)
- Save/cancel handlers defined (lines 1392-1417)
- But: handlers call non-existent methods

### Integration: ✗ BROKEN
- Methods exist in Rust ✓
- Methods defined in HTML ✓
- Methods exposed in bridge ✗ **MISSING**

---

## What Would Be Needed to Proceed

**To fix and test successfully**, one of the following is required:

**Option A**: Add methods to `app-bridge.js` (5 lines):
```javascript
window.electronExcel = {
  // ... existing methods ...
  excel_get_unidad_pesos: (payload) => invoke("excel_get_unidad_pesos", payload),
  excel_save_unidad_pesos: (payload) => invoke("excel_save_unidad_pesos", payload),
}
```

**Option B**: Fix HTML method names to match bridge conventions:
- Change `excel_get_unidades()` → `getUnidades()` (line 1325)
- Change `excel_get_unidad_pesos()` → new method in bridge
- Change `excel_save_unidad_pesos()` → new method in bridge

**Option C**: Fix method naming in HTML to use correct bridge names for pesos calls

---

## Summary of Findings

### Critical Issues
1. **Missing bridge methods**: 2 methods called by HTML but not defined in app-bridge.js
2. **Inconsistent naming**: HTML mixes camelCase and snake_case conventions
3. **Blocked execution**: All interactive tests fail at first backend IPC call

### What Works
- ✓ Rust backend implementations are complete and correct
- ✓ HTML structure and validation logic are correct
- ✓ UI layout and styling are correct
- ✓ Individual function implementations are sound

### What's Broken
- ✗ JavaScript bridge doesn't expose required methods
- ✗ HTML cannot call backend IPC due to missing bridge methods
- ✗ All tests requiring backend interaction are blocked

---

## Test Execution Report

| Test Group | Status | Issue | Evidence |
|-----------|--------|-------|----------|
| 1. UI Toggle | ✗ BLOCKED | Missing method in toggle handler | gestor-unidades.html:716 calls missing `loadUnidadesForPonderaciones()` → line 1325 missing method |
| 2. Dropdown | ✗ BLOCKED | Missing `excel_get_unidad_pesos` | Line 1343: `window.electronExcel.excel_get_unidad_pesos` undefined |
| 3. Validation | ✓ PARTIAL | Logic OK, but UI flow blocked | Client-side validation correct, but cannot load/save data |
| 4. Save | ✗ BLOCKED | Missing `excel_save_unidad_pesos` | Line 1407: `window.electronExcel.excel_save_unidad_pesos` undefined |
| 5. Integration | ✗ BLOCKED | Depends on Test 4 (blocked) | Cannot verify FINAL calculation uses new weights |
| 6. Persistence | ✗ BLOCKED | Depends on Test 4 (blocked) | Cannot verify weights persist after reload |

**Overall Result**: 0/6 test groups can execute. All blocked by integration bug.

---

## Recommendations

### Priority 1: Critical Fix Required
Add the missing methods to `app-bridge.js` to expose the backend commands:
```javascript
excel_get_unidad_pesos: (payload) => invoke("excel_get_unidad_pesos", payload),
excel_save_unidad_pesos: (payload) => invoke("excel_save_unidad_pesos", payload),
```

### Priority 2: Naming Consistency
Audit all method calls in HTML to ensure they match bridge exports:
- Current: mix of `camelCase` and `excel_snake_case`
- Should be: consistent with existing bridge pattern

### Priority 3: Pre-Test Validation
Implement a pre-test validation that checks:
- All `window.electronExcel.X()` calls have corresponding bridge methods
- All bridge methods have corresponding Rust handlers
- All Rust handlers are registered in the `tauri::generate_handler!` macro

---

## Conclusion

**Cannot proceed with manual testing** due to critical integration bug in the JavaScript bridge.  
**Root cause**: Missing IPC bridge method definitions for the ponderaciones feature.  
**Impact**: Feature is not functional in its current state.  
**Action required**: Fix bridge integration before testing can proceed.

The backend implementation (Rust) is complete and correct, but the frontend cannot communicate with it due to missing method definitions in the IPC bridge.

---

**Tested by**: Claude Code Agent  
**Test Date**: 2026-08-11  
**Test Duration**: Code review + integration verification (interactive testing not possible)  
**Environment**: Windows 11, Tauri v2 dev environment
