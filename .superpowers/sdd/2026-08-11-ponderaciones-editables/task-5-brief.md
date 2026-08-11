# Task 5: Integrar pesos en cálculo FINAL de gestor-notas.html

**Files:**
- Modify: `gestor-notas.html` - usar pesos de Excel en calculateFinal()

**Dependencies:**
- Task 3 must be complete (backend command excel_get_unidad_pesos available)

**Steps:**

### 5.1: Actualizar calculateFinal()

Buscar función `function calculateFinal(crNota, instrumentos)` y reemplazar:

```javascript
function calculateFinal(crNota, instrumentos, unidadPesos) {
    if (!crNota) return null;
    
    const values = [crNota.i1, crNota.i2, crNota.i3, crNota.i4];
    const weights = unidadPesos || [0.25, 0.25, 0.25, 0.25];  // Default si no se pasan pesos
    
    let numerator = 0, denominator = 0;
    for (let i = 0; i < 4; i++) {
        if (values[i] !== undefined && values[i] !== null) {
            numerator += weights[i] * values[i];
            denominator += weights[i];
        }
    }
    return denominator > 0 ? numerator / denominator : null;
}
```

### 5.2: Cargar pesos de la unidad al entrar a gestor-notas

En `loadNotasConGuardado()`, después de cargar unidad:

```javascript
// Cargar pesos de esta unidad
const pesosData = await window.electronExcel.excel_get_unidad_pesos({unidad: currentState.unidad});
currentState.unidadPesos = pesosData.pesos;
```

### 5.3: Usar pesos en updateFinalDisplay()

Buscar `updateFinalDisplay()` y cambiar:

```javascript
function updateFinalDisplay(studentIdx, crCodigo) {
    if (!currentNotes[studentIdx]) return;
    const crNota = currentNotes[studentIdx].crNotas?.find(n => n.codigo === crCodigo);
    if (!crNota) return;
    const final = calculateFinal(crNota, null, currentState.unidadPesos);  // ← pasar pesos
    crNota.final = final;
    // ... resto igual
}
```

**Notes:**
- Default weights [0.25, 0.25, 0.25, 0.25] provide fallback if pesos not loaded
- The weighted average formula: handles missing values by only summing weights of non-null instruments
- currentState.unidadPesos is initialized in loadNotasConGuardado() and used in all calls to calculateFinal()

**Global Constraints:**
- Pesos deben sumar 100% (tolerancia 0.001)
- Valores 0-1 (decimales: 0.2 = 20%)
