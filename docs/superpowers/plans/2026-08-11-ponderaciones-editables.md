# Interfaz Editable de Ponderaciones - Instrumentos por Unidad

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans

**Goal:** Profesor puede ver y editar pesos (i1-i4) por unidad, guardar en Excel fila 5 (C5:F5), usar en cálculo FINAL

**Architecture:** 
- gestor-unidades.html: Sección "Configurar Ponderaciones" con lista de unidades y tabla editable
- Modal de edición: Inputs para i1, i2, i3, i4 con validación suma=100%
- Backend: Comandos IPC para guardar/cargar pesos desde Excel

**Tech Stack:** Tauri IPC, HTML forms, JavaScript validation

## Global Constraints
- Pesos deben sumar 100% (tolerancia 0.001)
- Valores 0-1 (decimales: 0.2 = 20%)
- Mostrar en porcentaje (20%, 40%, etc.)
- Fila 5 de cada hoja de unidad (C5:F5)

---

## Task 1: Sección "Configurar Ponderaciones" en gestor-unidades.html

**Files:**
- Modify: `gestor-unidades.html` - agregar sección + modal

**Interfaces:**
- Consumes: currentState.unidad (string), lista de unidades disponibles
- Produces: UI que permite ver/editar pesos por unidad

**Steps:**

- [ ] **1.1: Agregar HTML para sección**

Agregar después de </main> en gestor-unidades.html:

```html
<section id="ponderacionesSection" style="display:none; padding: 20px; background: #f9fafb; margin-top: 20px; border-radius: 8px;">
    <h2>Configurar Ponderaciones por Unidad</h2>
    <p style="color: #6b7280; font-size: 0.9rem;">Ajusta los pesos de los instrumentos (i1, i2, i3, i4) para cada unidad. La suma debe ser 100%.</p>
    
    <label for="unidadSelector" style="display: block; margin-bottom: 10px; font-weight: 600;">Seleccionar Unidad:</label>
    <select id="unidadSelector" style="width: 100%; max-width: 300px; padding: 8px; margin-bottom: 20px; border: 1px solid #d1d5db; border-radius: 6px;">
        <option value="">-- Elige una unidad --</option>
    </select>

    <table id="ponderacionesTable" style="width: 100%; border-collapse: collapse; margin-bottom: 20px; display: none;">
        <thead>
            <tr style="background: #f3f4f6; border-bottom: 2px solid #d1d5db;">
                <th style="padding: 10px; text-align: left; border: 1px solid #e5e7eb;">Instrumento</th>
                <th style="padding: 10px; text-align: center; border: 1px solid #e5e7eb;">Peso (0-1)</th>
                <th style="padding: 10px; text-align: center; border: 1px solid #e5e7eb;">Porcentaje</th>
            </tr>
        </thead>
        <tbody>
            <tr style="border-bottom: 1px solid #e5e7eb;">
                <td style="padding: 10px; border: 1px solid #e5e7eb; font-weight: 600;">i1</td>
                <td style="padding: 10px; border: 1px solid #e5e7eb; text-align: center;">
                    <input type="number" id="peso_i1" min="0" max="1" step="0.01" style="width: 80px; padding: 6px; border: 1px solid #d1d5db; border-radius: 4px;">
                </td>
                <td style="padding: 10px; border: 1px solid #e5e7eb; text-align: center;">
                    <span id="pct_i1" style="font-weight: 600; color: #374151;">-</span>%
                </td>
            </tr>
            <tr style="border-bottom: 1px solid #e5e7eb;">
                <td style="padding: 10px; border: 1px solid #e5e7eb; font-weight: 600;">i2</td>
                <td style="padding: 10px; border: 1px solid #e5e7eb; text-align: center;">
                    <input type="number" id="peso_i2" min="0" max="1" step="0.01" style="width: 80px; padding: 6px; border: 1px solid #d1d5db; border-radius: 4px;">
                </td>
                <td style="padding: 10px; border: 1px solid #e5e7eb; text-align: center;">
                    <span id="pct_i2" style="font-weight: 600; color: #374151;">-</span>%
                </td>
            </tr>
            <tr style="border-bottom: 1px solid #e5e7eb;">
                <td style="padding: 10px; border: 1px solid #e5e7eb; font-weight: 600;">i3</td>
                <td style="padding: 10px; border: 1px solid #e5e7eb; text-align: center;">
                    <input type="number" id="peso_i3" min="0" max="1" step="0.01" style="width: 80px; padding: 6px; border: 1px solid #d1d5db; border-radius: 4px;">
                </td>
                <td style="padding: 10px; border: 1px solid #e5e7eb; text-align: center;">
                    <span id="pct_i3" style="font-weight: 600; color: #374151;">-</span>%
                </td>
            </tr>
            <tr style="background: #f9fafb; border-bottom: 1px solid #e5e7eb;">
                <td style="padding: 10px; border: 1px solid #e5e7eb; font-weight: 600;">i4</td>
                <td style="padding: 10px; border: 1px solid #e5e7eb; text-align: center;">
                    <input type="number" id="peso_i4" min="0" max="1" step="0.01" style="width: 80px; padding: 6px; border: 1px solid #d1d5db; border-radius: 4px;">
                </td>
                <td style="padding: 10px; border: 1px solid #e5e7eb; text-align: center;">
                    <span id="pct_i4" style="font-weight: 600; color: #374151;">-</span>%
                </td>
            </tr>
        </tbody>
    </table>

    <div id="validationMessage" style="margin-bottom: 15px; padding: 10px; border-radius: 6px; display: none;">
    </div>

    <div style="display: flex; gap: 10px;">
        <button id="guardarPonderacionesBtn" type="button" style="padding: 10px 20px; background: #16a34a; color: white; border: none; border-radius: 6px; font-weight: 600; cursor: pointer;">
            Guardar Ponderaciones
        </button>
        <button id="cancelarPonderacionesBtn" type="button" style="padding: 10px 20px; background: #6b7280; color: white; border: none; border-radius: 6px; font-weight: 600; cursor: pointer;">
            Cancelar
        </button>
    </div>
</section>
```

- [ ] **1.2: Agregar toggle para mostrar sección**

En gestor-unidades.html, agregar botón en toolbar (después de "Agregar Unidad"):

```html
<button id="togglePonderacionesBtn" type="button" style="padding: 8px 16px; background: #3b82f6; color: white; border: none; border-radius: 6px; font-weight: 600; cursor: pointer;">
    ⚙️ Configurar Ponderaciones
</button>
```

- [ ] **1.3: Script para toggle**

En DOMContentLoaded:

```javascript
const togglePonderacionesBtn = document.getElementById('togglePonderacionesBtn');
const ponderacionesSection = document.getElementById('ponderacionesSection');
togglePonderacionesBtn.addEventListener('click', () => {
    const isVisible = ponderacionesSection.style.display !== 'none';
    ponderacionesSection.style.display = isVisible ? 'none' : 'block';
    if (!isVisible) loadUnidadesForPonderaciones();
});
```

---

## Task 2: Cargar unidades en dropdown + validación en tiempo real

**Files:**
- Modify: `gestor-unidades.html` - agregar script

**Steps:**

- [ ] **2.1: Función para cargar unidades**

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

- [ ] **2.2: Event listener para dropdown**

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

- [ ] **2.3: Función validación en tiempo real**

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

---

## Task 3: Backend - Comandos IPC para cargar/guardar pesos

**Files:**
- Modify: `src-tauri/src/main.rs` - agregar comandos

**Steps:**

- [ ] **3.1: Comando excel_get_unidad_pesos**

```rust
#[tauri::command]
async fn excel_get_unidad_pesos(payload: Value) -> Result<Value, String> {
    tauri::async_runtime::spawn_blocking(move || excel_get_unidad_pesos_impl(payload))
        .await
        .map_err(|e| e.to_string())?
}

fn excel_get_unidad_pesos_impl(payload: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let unidad = payload["unidad"].as_str().ok_or("Falta unidad")?.to_string();
    
    let rows = read_sheet_rows(&path, &unidad)?;
    
    // Fila 5 (0-idx 4): C5:F5 = columnas 2:5
    let pesos = [
        cell_f64(&rows, 4, 2).unwrap_or(0.25),  // C5 = i1
        cell_f64(&rows, 4, 3).unwrap_or(0.25),  // D5 = i2
        cell_f64(&rows, 4, 4).unwrap_or(0.25),  // E5 = i3
        cell_f64(&rows, 4, 5).unwrap_or(0.25),  // F5 = i4
    ];
    
    Ok(json!({ "pesos": pesos }))
}
```

- [ ] **3.2: Comando excel_save_unidad_pesos**

```rust
#[tauri::command]
async fn excel_save_unidad_pesos(payload: Value) -> Result<Value, String> {
    tauri::async_runtime::spawn_blocking(move || excel_save_unidad_pesos_impl(payload))
        .await
        .map_err(|e| e.to_string())?
}

fn excel_save_unidad_pesos_impl(payload: Value) -> Result<Value, String> {
    let path = require_selected_path()?;
    let unidad = payload["unidad"].as_str().ok_or("Falta unidad")?.to_string();
    let pesos = payload["pesos"].as_array().ok_or("Falta pesos")?;
    
    if pesos.len() != 4 {
        return Err("Se requieren exactamente 4 pesos".to_string());
    }
    
    let pesos_f64: Vec<f64> = pesos.iter()
        .filter_map(|p| p.as_f64())
        .collect();
    
    if pesos_f64.len() != 4 {
        return Err("Todos los pesos deben ser números".to_string());
    }
    
    let suma: f64 = pesos_f64.iter().sum();
    if (suma - 1.0).abs() > 0.001 {
        return Err(format!("Suma debe ser 1.0 (100%), es {:.2}", suma));
    }
    
    let pesos_owned = pesos_f64.clone();
    let edit_fn: Box<dyn Fn(&str) -> Result<String, String>> = Box::new(move |xml: &str| {
        let mut s = xml.to_string();
        // Fila 5 (0-idx 4), columnas C:F (0-idx 2:5)
        for (i, &peso) in pesos_owned.iter().enumerate() {
            s = set_xml_cell(&s, 4, 2 + i, Some(&json!(peso)), "number")?;
        }
        Ok(s)
    });
    
    edit_workbook_sheets_xml(&path, vec![(unidad.as_str(), edit_fn)])?;
    
    Ok(json!({ "pesos": pesos_f64 }))
}
```

---

## Task 4: Guardar ponderaciones - click handler

**Files:**
- Modify: `gestor-unidades.html` - agregar handler

**Steps:**

- [ ] **4.1: Click handler guardar**

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

---

## Task 5: Integrar pesos en cálculo FINAL de gestor-notas.html

**Files:**
- Modify: `gestor-notas.html` - usar pesos de Excel en calculateFinal()

**Steps:**

- [ ] **5.1: Actualizar calculateFinal()**

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

- [ ] **5.2: Cargar pesos de la unidad al entrar a gestor-notas**

En `loadNotasConGuardado()`, después de cargar unidad:

```javascript
// Cargar pesos de esta unidad
const pesosData = await window.electronExcel.excel_get_unidad_pesos({unidad: currentState.unidad});
currentState.unidadPesos = pesosData.pesos;
```

- [ ] **5.3: Usar pesos en updateFinalDisplay()**

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

---

## Task 6: Testing Manual

- [ ] Abrir gestor-unidades.html
- [ ] Click "Configurar Ponderaciones"
- [ ] Seleccionar una unidad
- [ ] Editar pesos (ejemplo: 0.5, 0.3, 0.1, 0.1)
- [ ] Verificar suma = 50+30+10+10 = 100% ✓
- [ ] Click "Guardar"
- [ ] Abrir gestor-notas con esa unidad
- [ ] Entrar nota y verificar que FINAL se calcula con los nuevos pesos

---

**Total: 6 tareas**

Cobertura: UI + validación + backend IPC + integración en cálculo
