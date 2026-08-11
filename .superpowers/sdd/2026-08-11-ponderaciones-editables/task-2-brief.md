# Task 2: Cargar unidades en dropdown + validación en tiempo real

**Files:**
- Modify: `gestor-unidades.html` - agregar script

**Steps:**

### 2.1: Función para cargar unidades

```javascript
async function loadUnidadesForPonderaciones() {
    const selector = document.getElementById('unidadSelector');
    selector.innerHTML = '<option value="">-- Elige una unidad --</option>';
    
    // Cargar lista de unidades (mismo que en main)
    const data = await window.electronExcel.excel_get_unidades();
    (data.unidades || []).forEach(u => {
        const option = document.createElement('option');
        option.value = u.codigo;
        option.textContent = `${u.codigo} - ${u.nombre}`;
        selector.appendChild(option);
    });
}
```

### 2.2: Event listener para dropdown

```javascript
document.getElementById('unidadSelector').addEventListener('change', async (e) => {
    const unidad = e.target.value;
    if (!unidad) {
        document.getElementById('ponderacionesTable').style.display = 'none';
        return;
    }
    
    // Cargar pesos actuales de la unidad
    const data = await window.electronExcel.excel_get_unidad_pesos({unidad});
    const pesos = data.pesos || [0.25, 0.25, 0.25, 0.25];
    
    document.getElementById('peso_i1').value = pesos[0];
    document.getElementById('peso_i2').value = pesos[1];
    document.getElementById('peso_i3').value = pesos[2];
    document.getElementById('peso_i4').value = pesos[3];
    
    document.getElementById('ponderacionesTable').style.display = 'table';
    actualizarValidacion();
});
```

### 2.3: Función validación en tiempo real

```javascript
function actualizarValidacion() {
    const i1 = parseFloat(document.getElementById('peso_i1').value) || 0;
    const i2 = parseFloat(document.getElementById('peso_i2').value) || 0;
    const i3 = parseFloat(document.getElementById('peso_i3').value) || 0;
    const i4 = parseFloat(document.getElementById('peso_i4').value) || 0;
    
    // Actualizar porcentajes
    document.getElementById('pct_i1').textContent = Math.round(i1 * 100);
    document.getElementById('pct_i2').textContent = Math.round(i2 * 100);
    document.getElementById('pct_i3').textContent = Math.round(i3 * 100);
    document.getElementById('pct_i4').textContent = Math.round(i4 * 100);
    
    const suma = i1 + i2 + i3 + i4;
    const validationMsg = document.getElementById('validationMessage');
    
    if (Math.abs(suma - 1.0) <= 0.001) {
        validationMsg.style.background = '#dcfce7';
        validationMsg.style.color = '#166534';
        validationMsg.textContent = '✓ Suma = 100% (válido)';
        validationMsg.style.display = 'block';
        document.getElementById('guardarPonderacionesBtn').disabled = false;
    } else {
        validationMsg.style.background = '#fee2e2';
        validationMsg.style.color = '#991b1b';
        validationMsg.textContent = `✗ Suma = ${Math.round(suma * 100)}% (debe ser 100%)`;
        validationMsg.style.display = 'block';
        document.getElementById('guardarPonderacionesBtn').disabled = true;
    }
}

// Escuchar cambios en inputs
['peso_i1', 'peso_i2', 'peso_i3', 'peso_i4'].forEach(id => {
    document.getElementById(id).addEventListener('input', actualizarValidacion);
});
```

**Global Constraints:**
- Pesos deben sumar 100% (tolerancia 0.001)
- Valores 0-1 (decimales: 0.2 = 20%)
- Mostrar en porcentaje (20%, 40%, etc.)
