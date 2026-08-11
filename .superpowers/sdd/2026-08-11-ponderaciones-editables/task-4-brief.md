# Task 4: Guardar ponderaciones - click handler

**Files:**
- Modify: `gestor-unidades.html` - agregar handler

**Dependencies:**
- Task 3 must be complete (backend commands excel_save_unidad_pesos available)

**Steps:**

### 4.1: Click handler guardar

```javascript
document.getElementById('guardarPonderacionesBtn').addEventListener('click', async () => {
    const unidad = document.getElementById('unidadSelector').value;
    if (!unidad) {
        showMessage('Selecciona una unidad', 'error');
        return;
    }
    
    const pesos = [
        parseFloat(document.getElementById('peso_i1').value),
        parseFloat(document.getElementById('peso_i2').value),
        parseFloat(document.getElementById('peso_i3').value),
        parseFloat(document.getElementById('peso_i4').value)
    ];
    
    try {
        await window.electronExcel.excel_save_unidad_pesos({ unidad, pesos });
        showMessage(`Ponderaciones guardadas para ${unidad}`, 'success');
        document.getElementById('ponderacionesSection').style.display = 'none';
    } catch (error) {
        showMessage(error.message || 'Error al guardar', 'error');
    }
});

document.getElementById('cancelarPonderacionesBtn').addEventListener('click', () => {
    document.getElementById('ponderacionesSection').style.display = 'none';
});
```

**Notes:**
- Uses existing showMessage() helper (already in gestor-unidades.html)
- Assumes validation by Task 2 ensures suma = 100%
- Closes section after successful save
- Both handlers use the showMessage pattern consistent with existing code

**Global Constraints:**
- Pesos deben sumar 100% (tolerancia 0.001) — validation done in Task 2, backend re-validates in Task 3
