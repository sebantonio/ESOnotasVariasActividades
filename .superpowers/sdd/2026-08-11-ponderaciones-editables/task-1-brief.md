# Task 1: Sección "Configurar Ponderaciones" en gestor-unidades.html

**Files:**
- Modify: `gestor-unidades.html` - agregar sección + modal

**Interfaces:**
- Consumes: currentState.unidad (string), lista de unidades disponibles
- Produces: UI que permite ver/editar pesos por unidad

**Steps:**

### 1.1: Agregar HTML para sección

Agregar después de `</main>` en gestor-unidades.html:

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

### 1.2: Agregar toggle para mostrar sección

En gestor-unidades.html, agregar botón en toolbar (después de "Agregar Unidad"):

```html
<button id="togglePonderacionesBtn" type="button" style="padding: 8px 16px; background: #3b82f6; color: white; border: none; border-radius: 6px; font-weight: 600; cursor: pointer;">
    ⚙️ Configurar Ponderaciones
</button>
```

### 1.3: Script para toggle

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

**Global Constraints (from plan):**
- Pesos deben sumar 100% (tolerancia 0.001)
- Valores 0-1 (decimales: 0.2 = 20%)
- Mostrar en porcentaje (20%, 40%, etc.)
- Fila 5 de cada hoja de unidad (C5:F5)
