# Frontend: instrumentos por criterio en UI — Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Actualizar las páginas HTML (gestor-notas, gestor-unidades, gestor-recuperaciones, visor-unidades) para que muestren y editen 4 campos de instrumento (i1, i2, i3, i4) + FINAL calculado + Rec por criterio, en lugar del layout antiguo de una nota directa por criterio.

**Architecture:** Los cambios viven en JS vanilla (sin framework). Cada página que gestiona/visualiza notas reutiliza el mismo patrón: (1) cargar datos via `app.excel.getNotasUnidad()` que ahora devuelve la nueva forma con `i1-i4-final`, (2) renderizar tabla con 4 inputs por CR + 1 readonly FINAL + 1 Rec editable, (3) salvar con payload `{crNotas[codigo] = {colIdx, i1, i2, i3, i4, rec}}` para las tres primeras, (4) solo lectura para visor-unidades. `gestor-unidades.html` obtiene un nuevo flujo: selector de 4 instrumentos del catálogo + inputs de peso, validar que sumen 100%, salvar via el comando new `excel_save_unidad_instrumentos`.

**Tech Stack:** HTML5 + Vanilla JS, Tauri IPC via `app-bridge.js`, serde_json payloads, Excel XLSX via backend.

## Global Constraints

- No dependencies nuevas en `package.json` (vanilla JS puro).
- Los pesos de instrumento son valores `0.0`—`1.0` (NOT porcentajes × 100 — el backend calcula con estos valores directos).
- NUNCA editar ni guardar la columna FINAL (es una fórmula Excel `SUMPRODUCT`, se recalcula sola).
- Validar en JS que pesos sumen 100% (tolerancia ±0.1%) antes de enviar al backend.
- Mantener el mismo CSS/UX (sticky columna alumno, paginación 15/página, colores CE, modales inline-style `display`).
- No quebrar visor-notas.html ni visor-recuperaciones.html (no son páginas de evaluación, no tocarlas; aunque vea referencias aquí son contexto nada más).

---

### Task 1: `app-bridge.js` — wrappers para nuevos comandos

**Files:**
- Modify: `app-bridge.js:34` (añadir junto a `saveNotasUnidad`)

**Interfaces:**
- Produces: `app.excel.saveUnidadInstrumentos(payload)` que llama a Tauri `excel_save_unidad_instrumentos`

- [ ] **Step 1: Leer el archivo actual**

Lee `app-bridge.js` líneas 1-50 para ver estructura de wrappers.

- [ ] **Step 2: Añadir el wrapper**

Después de la línea que define `saveNotasUnidad`, añadir:

```javascript
    saveUnidadInstrumentos: (payload) => invoke("excel_save_unidad_instrumentos", { payload }),
```

(Nota: el backend ya tiene registrado este comando — Task 6 del plan backend)

- [ ] **Step 3: Commit**

```bash
git add app-bridge.js
git commit -m "feat(frontend): wrapper para excel_save_unidad_instrumentos en app-bridge.js"
```

---

### Task 2: `gestor-notas.html` — tabla con 4 instrumentos por criterio + FINAL

**Files:**
- Modify: `gestor-notas.html` (función `renderTable()` + payload en `saveNotasUnidad()`)

**Interfaces:**
- Consumes: `app.excel.getNotasUnidad(unidad)` que devuelve nuevo JSON con:
  - `criterios[].{codigo, colIdx, ponderacion, recColIdx}` (sin cambios)
  - `instrumentosUnidad[].{slot, abrev, peso}` (nuevo)
  - `alumnos[].crNotas[].{codigo, colIdx, i1, i2, i3, i4, final, recDisplay}` (cambio: `i1-i4-final` en vez de `nota`)
- Produces: `saveNotasUnidad()` envia payload con forma nueva: `crNotas[codigo] = {colIdx, i1, i2, i3, i4, rec}`

- [ ] **Step 1: Leer estructura actual de la tabla**

Lee `gestor-notas.html` línea de `renderTable()` (busca `function renderTable`) y línea de `saveNotasUnidad()` (busca `function saveNotasUnidad`). Nota cuál es el patrón de inputs por alumno/criterio.

- [ ] **Step 2: Actualizar cabecera de tabla en `renderTable()`**

Busca el bloque que construye el `<tr>` de cabecera (`<thead>`). 

**Cambio esperado:** donde antes había 1 columna de Nota (+ Rec), ahora:
- 4 columnas: i1 | i2 | i3 | i4
- 1 columna: FINAL (readonly)
- 1 columna: Rec

Por ejemplo, reemplazar:

```html
<th>Nota</th><th>Rec</th>
```

por:

```html
<th>i1</th><th>i2</th><th>i3</th><th>i4</th><th>FINAL</th><th>Rec</th>
```

- [ ] **Step 3: Actualizar fila de datos en `renderTable()`**

Busca el bloque que construye cada celda de criterio (`<td>` con inputs). 

**Cambio:** reemplazar el único `<input>` de `nota` por 4 inputs (i1-i4) + 1 celda readonly FINAL + 1 input Rec.

Código exacto a reemplazar (aprox línea 300-320, busca `data-criteria-row`):

```javascript
const cellHTML = `
    <input type="number" min="0" max="10" step="0.1" 
        class="criteria-input" 
        data-student-idx="${studentIdx}" 
        data-criteria-code="${criteria.codigo}"
        data-field="nota"
        value="${criteria.nota || ''}" />
    <input type="number" min="0" max="10" step="0.1" 
        class="criteria-input rec-input" 
        data-student-idx="${studentIdx}" 
        data-criteria-code="${criteria.codigo}"
        data-field="rec"
        value="${criteria.recDisplay || ''}" />
`;
```

Reemplazar por:

```javascript
const i1Val = crNote.i1 || '';
const i2Val = crNote.i2 || '';
const i3Val = crNote.i3 || '';
const i4Val = crNote.i4 || '';
const finalVal = crNote.final !== null ? crNote.final.toFixed(1) : '';

const cellHTML = `
    <input type="number" min="0" max="10" step="0.1" 
        class="criteria-input" 
        data-student-idx="${studentIdx}" 
        data-criteria-code="${criteria.codigo}"
        data-field="i1"
        value="${i1Val}" />
    <input type="number" min="0" max="10" step="0.1" 
        class="criteria-input" 
        data-student-idx="${studentIdx}" 
        data-criteria-code="${criteria.codigo}"
        data-field="i2"
        value="${i2Val}" />
    <input type="number" min="0" max="10" step="0.1" 
        class="criteria-input" 
        data-student-idx="${studentIdx}" 
        data-criteria-code="${criteria.codigo}"
        data-field="i3"
        value="${i3Val}" />
    <input type="number" min="0" max="10" step="0.1" 
        class="criteria-input" 
        data-student-idx="${studentIdx}" 
        data-criteria-code="${criteria.codigo}"
        data-field="i4"
        value="${i4Val}" />
    <span class="final-display">${finalVal}</span>
    <input type="number" min="0" max="10" step="0.1" 
        class="criteria-input rec-input" 
        data-student-idx="${studentIdx}" 
        data-criteria-code="${criteria.codigo}"
        data-field="rec"
        value="${crNote.recDisplay || ''}" />
`;
```

(Nota: `.final` puede ser `null` si todos los instrumentos están vacíos — comprobar con `!== null` antes de `.toFixed()`)

- [ ] **Step 4: Actualizar `saveNotasUnidad()` — payload**

Busca la función `saveNotasUnidad()`. Localiza donde construye el objeto `notas[]` con `crNotas[codigo]`.

**Cambio:** el payload viejo era:

```javascript
{
  unidad: "U1",
  syncEval: true,
  notas: [
    { rowIdx: 6, crNotas: { "CR1.1": { colIdx: 7, nota: 8.5, rec: 5.0 } } },
    ...
  ]
}
```

**Nuevo payload:**

```javascript
{
  unidad: "U1",
  syncEval: true,
  notas: [
    { rowIdx: 6, crNotas: { 
        "CR1.1": { colIdx: 7, i1: 8.0, i2: 7.0, i3: 9.0, i4: 6.0, rec: 5.0 }
      } 
    },
    ...
  ]
}
```

Reemplaza el bloque de lectura de inputs (búsqueda de todos los `data-student-idx`/`data-criteria-code`):

```javascript
// Viejo (reemplazar):
const nota = parseFloat(input.value) || null;
crNotas[code] = { colIdx, nota, rec };

// Nuevo:
const field = input.dataset.field;
const val = input.value ? parseFloat(input.value) : null;

if (!crNotas[code]) {
    crNotas[code] = { colIdx };
}
crNotas[code][field] = val;
```

- [ ] **Step 5: Añadir CSS para la columna FINAL (readonly)**

Busca la sección `<style>` o `style.css` incluido.

Añadir (o actualizar si existe `.final-display`):

```css
.final-display {
    display: inline-block;
    min-width: 3rem;
    text-align: right;
    font-weight: 600;
    color: #666;
}
```

- [ ] **Step 6: Test manual**

1. Abre `gestor-notas.html`, selecciona una unidad.
2. Verifica que se ven 4 inputs por criterio + FINAL readonly + Rec.
3. Teclea valores en i1-i4 → FINAL debe cambiar en tiempo real (calculado en el backend al guardar).
4. Teclea un valor en Rec.
5. Guarda → confirmar que el payload contiene `{i1, i2, i3, i4, rec}`.
6. Recarga → verifica que los valores se releen correctamente.

- [ ] **Step 7: Commit**

```bash
git add gestor-notas.html
git commit -m "feat(frontend): gestor-notas render 4 instrumentos + FINAL + Rec por criterio"
```

---

### Task 3: `gestor-recuperaciones.html` — tabla Rec por unidad (similar a Task 2)

**Files:**
- Modify: `gestor-recuperaciones.html` (función `renderTable()` + `saveAllRec()`)

**Interfaces:**
- Consumes: `app.excel.getNotasUnidad(unidad)` (misma forma que Task 2)
- Produces: payload con `crNotas[codigo] = {colIdx, rec}` (solo Rec, i1-i4 no se tocan aquí)

**Contexto:** gestor-recuperaciones es específico para editar Rec por unidad (no por evaluación). Nota CE por grupo se recalcula en JS.

- [ ] **Step 1: Entender la lógica actual**

Lee `gestor-recuperaciones.html` líneas de `renderTable()` y `saveAllRec()`. Nota que:
- Usa `excel_get_notas_unidad` (igual que gestor-notas)
- Renderiza paginación 15/página
- Agrupa criterios por `getCeNum()` (extrae el número del prefijo CR)
- Recalcula Nota CE en JS sumando Rec de cada grupo

- [ ] **Step 2: Actualizar tabla — desaparecer columna de Nota**

Busca dónde renderiza las columnas de criterio. **Cambio:** antes mostraba "Nota | Rec", ahora solo "Rec" (sin Nota, porque Rec es lo único editable en esta página).

Reemplaza la cabecera de criterio:

```javascript
// Viejo:
<th>Nota</th><th>Rec</th>

// Nuevo:
<th>Rec</th>
```

Y en la fila de datos:

```javascript
// Viejo (reemplazar):
const cellHTML = `
    <span class="nota-display">${nota || '-'}</span>
    <input type="number" min="0" max="10" step="0.1" ... data-field="rec" />
`;

// Nuevo:
const cellHTML = `
    <input type="number" min="0" max="10" step="0.1" 
        class="criteria-input rec-input" 
        data-student-idx="${studentIdx}" 
        data-criteria-code="${criteria.codigo}"
        data-field="rec"
        value="${crNote.recDisplay || ''}" />
`;
```

- [ ] **Step 3: Actualizar `saveAllRec()` — payload**

El payload es similar a Task 2 pero **solo incluye `rec`**, no `i1-i4`:

```javascript
{
  unidad: "U1",
  syncEval: false,  // autosave silencioso, no propagar a eval (Rec no se cachea en eval, solo en unidad)
  notas: [
    { rowIdx: 6, crNotas: { "CR1.1": { colIdx: 7, rec: 5.0 } } },
    ...
  ]
}
```

Busca el bloque de recolección de inputs y reemplaza:

```javascript
// Viejo:
const nota = ...; 
const rec = ...;
crNotas[code] = { colIdx, nota, rec };

// Nuevo:
const rec = input.value ? parseFloat(input.value) : null;
crNotas[code] = { colIdx, rec };
```

- [ ] **Step 4: Verificar que la recalculación de Nota CE sigue funcionando**

Busca `recomputeAlumno()` (debería existir). Verifica que:
- Lee `criterios[].ponderacion` de la respuesta de `load_notas_unidad`
- Agrupa CRs por `getCeNum()` (extrae número de "CR1.2" → 1)
- Suma Rec ponderadas por CE

Este cálculo NO debe cambiar — solo necesita acceder a `crNote.recDisplay` igual que antes.

- [ ] **Step 5: Test manual**

1. Abre `gestor-recuperaciones.html`, selecciona unidad.
2. Verifica que ve columna Rec pero NO Nota.
3. Edita varios Rec.
4. Guarda → confirma payload contiene solo `{colIdx, rec}`.
5. Recarga → valores persisten.

- [ ] **Step 6: Commit**

```bash
git add gestor-recuperaciones.html
git commit -m "feat(frontend): gestor-recuperaciones solo Rec por criterio (sin notas i1-i4)"
```

---

### Task 4: `gestor-unidades.html` — configurar instrumentos/pesos por unidad

**Files:**
- Modify: `gestor-unidades.html` (nueva sección UI + `saveUnidadInstrumentos()`)

**Interfaces:**
- Consumes: `app.excel.loadInstrumentos()` que devuelve catálogo de todos los instrumentos (via backend `load_instrumentos`)
- Produces: `app.excel.saveUnidadInstrumentos(payload)` donde payload es `{unidad, slots: [{abrev, peso}, ...]}`
- Validates: pesos suman 100% (tolerancia ±0.1%)

**Contexto:** nueva funcionalidad — permite elegir qué 4 instrumentos usa cada unidad y asignar pesos. No existía antes.

- [ ] **Step 1: Leer la estructura de gestor-unidades.html**

Lee el archivo completo. Nota:
- Cómo se carga la lista de unidades (probablemente tabla).
- Dónde se podría insertar un formulario de instrumentos.

- [ ] **Step 2: Diseñar el modal/formulario de instrumentos**

Crear una nueva sección (puede ser un modal o un expandible dentro de la tabla). **Flujo:**
1. Usuario hace clic en "Configurar instrumentos" para una unidad.
2. Modal muestra: 4 slots (dropdown de catálogo + input de peso) + botón Guardar.
3. Carga los instrumentos actuales de esa unidad via `load_notas_unidad(unidad).instrumentosUnidad`.
4. Valida en JS que pesos sumen a 1.0 (±0.001).
5. Envía payload al backend.

Añadir al HTML (dentro del `<body>`):

```html
<div id="modalInstrumentos" style="display:none; position:fixed; top:0; left:0; width:100%; height:100%; background:rgba(0,0,0,0.5); z-index:1000;">
  <div style="background:white; margin:50px auto; padding:20px; max-width:500px; border-radius:8px;">
    <h2>Configurar instrumentos para <span id="modalUnidadName"></span></h2>
    
    <div id="slotsContainer">
      <!-- Se genera con JS -->
    </div>
    
    <div style="margin-top:20px; text-align:right;">
      <button onclick="guardarUnidadInstrumentos()">Guardar</button>
      <button onclick="cerrarModalInstrumentos()">Cancelar</button>
    </div>
    <div id="mensajeValidacion" style="color:red; margin-top:10px; font-size:0.9rem;"></div>
  </div>
</div>
```

- [ ] **Step 3: Implementar `abrirModalInstrumentos(unidad)` en JS**

```javascript
async function abrirModalInstrumentos(unidad) {
    document.getElementById('modalUnidadName').textContent = unidad;
    
    // Cargar instrumentos actuales
    const data = await app.excel.getNotasUnidad(unidad);
    const instrumentosActuales = data.instrumentosUnidad || [];
    
    // Cargar catálogo disponible
    const catalogo = await app.excel.loadInstrumentos();
    const opcionesCatalogo = catalogo.instrumentos.map(i => i.codigo);
    
    // Renderizar 4 slots
    const container = document.getElementById('slotsContainer');
    container.innerHTML = '';
    
    for (let i = 0; i < 4; i++) {
        const actual = instrumentosActuales[i] || {};
        const slot = document.createElement('div');
        slot.style.marginBottom = '15px';
        
        const selectHTML = `
            <label>Instrumento ${i + 1}:</label>
            <select id="slot${i}_abrev">
                <option value="">-- Vacío --</option>
                ${opcionesCatalogo.map(op => `<option value="${op}" ${actual.abrev === op ? 'selected' : ''}>${op}</option>`).join('')}
            </select>
            <label style="margin-left:10px;">Peso:</label>
            <input type="number" min="0" max="1" step="0.01" id="slot${i}_peso" 
                value="${actual.peso || 0}" style="width:60px;" />
        `;
        slot.innerHTML = selectHTML;
        container.appendChild(slot);
    }
    
    document.getElementById('modalInstrumentos').style.display = 'flex';
    document.getElementById('mensajeValidacion').textContent = '';
}

function cerrarModalInstrumentos() {
    document.getElementById('modalInstrumentos').style.display = 'none';
}
```

- [ ] **Step 4: Implementar validación `validarPesos()` en JS**

```javascript
function validarPesos() {
    let suma = 0.0;
    const slots = [];
    
    for (let i = 0; i < 4; i++) {
        const abrev = document.getElementById(`slot${i}_abrev`).value;
        const peso = parseFloat(document.getElementById(`slot${i}_peso`).value) || 0;
        
        if (abrev) {
            slots.push({ abrev, peso });
            suma += peso;
        }
    }
    
    // Validar suma (tolerancia ±0.1% = 0.001)
    if (slots.length === 0) {
        return { valid: false, msg: "Debe elegir al menos un instrumento." };
    }
    if (Math.abs(suma - 1.0) > 0.001) {
        return { valid: false, msg: `Los pesos deben sumar 100% (suman ${(suma * 100).toFixed(1)}%).` };
    }
    
    return { valid: true, slots };
}
```

- [ ] **Step 5: Implementar `guardarUnidadInstrumentos()` en JS**

```javascript
async function guardarUnidadInstrumentos() {
    const validacion = validarPesos();
    if (!validacion.valid) {
        document.getElementById('mensajeValidacion').textContent = validacion.msg;
        return;
    }
    
    const unidad = document.getElementById('modalUnidadName').textContent;
    const payload = {
        unidad,
        slots: validacion.slots
    };
    
    try {
        await app.excel.saveUnidadInstrumentos(payload);
        alert(`Instrumentos de ${unidad} guardados correctamente.`);
        cerrarModalInstrumentos();
        // Opcional: recargar la tabla si existe
        if (typeof renderTable === 'function') {
            renderTable();
        }
    } catch (err) {
        document.getElementById('mensajeValidacion').textContent = `Error: ${err}`;
    }
}
```

- [ ] **Step 6: Añadir botón "Configurar" en la tabla de unidades**

Busca dónde renderiza la tabla de unidades. Añade una columna/botón:

```html
<button onclick="abrirModalInstrumentos('${unidad}')">Configurar instrumentos</button>
```

- [ ] **Step 7: Necesidad de `loadInstrumentos` en app-bridge.js**

Verifica que el backend tiene `load_instrumentos` y está expuesto. Si no existe, necesita añadirse (pero el backend plan Task 6 dice que existe en línea 2881). En `app-bridge.js`, añadir wrapper si falta:

```javascript
loadInstrumentos: () => invoke("excel_load_instrumentos", {}),
```

(Verificar que el backend tiene registrado este comando en `generate_handler!`)

- [ ] **Step 8: Test manual**

1. Abre `gestor-unidades.html`.
2. Haz clic en "Configurar instrumentos" para una unidad.
3. Elige 4 instrumentos y ajusta pesos para sumar 100%.
4. Guarda → confirma sin error.
5. Recarga y vuelve a abrir → verifica que se leen los valores guardados.
6. Intenta guardar con pesos que NO sumen 100% → debe rechazar con mensaje.

- [ ] **Step 9: Commit**

```bash
git add gestor-unidades.html app-bridge.js
git commit -m "feat(frontend): gestor-unidades configurar instrumentos/pesos por unidad"
```

---

### Task 5: `visor-unidades.html` — mostrar i1-i4-FINAL por criterio (read-only)

**Files:**
- Modify: `visor-unidades.html` (función `renderTable()`)

**Interfaces:**
- Consumes: `app.excel.getNotasUnidad(unidad)` (misma forma que Task 2)
- Produces: solo lectura — sin cambios de interfaz

**Contexto:** página de visualización — no hay inputs, solo mostrar valores.

- [ ] **Step 1: Leer estructura actual**

Lee `visor-unidades.html`. Nota cómo renderiza las notas actualmente.

- [ ] **Step 2: Actualizar cabecera de criterio**

Reemplazar:

```html
<th>Nota</th>
```

por:

```html
<th>i1</th><th>i2</th><th>i3</th><th>i4</th><th>FINAL</th>
```

- [ ] **Step 3: Actualizar fila de datos**

Reemplazar:

```javascript
const cellHTML = `<span class="nota-display">${criteria.nota || '-'}</span>`;
```

por:

```javascript
const i1 = crNote.i1 !== undefined ? crNote.i1.toFixed(1) : '-';
const i2 = crNote.i2 !== undefined ? crNote.i2.toFixed(1) : '-';
const i3 = crNote.i3 !== undefined ? crNote.i3.toFixed(1) : '-';
const i4 = crNote.i4 !== undefined ? crNote.i4.toFixed(1) : '-';
const final = crNote.final !== undefined && crNote.final !== null ? crNote.final.toFixed(1) : '-';

const cellHTML = `
    <span class="nota-display">${i1}</span>
    <span class="nota-display">${i2}</span>
    <span class="nota-display">${i3}</span>
    <span class="nota-display">${i4}</span>
    <span class="nota-display" style="font-weight:bold;">${final}</span>
`;
```

- [ ] **Step 4: Test manual**

1. Abre `visor-unidades.html`, selecciona una unidad.
2. Verifica que ve i1, i2, i3, i4, FINAL (todos read-only).
3. Los valores deben coincidir con lo guardado en gestor-notas.

- [ ] **Step 5: Commit**

```bash
git add visor-unidades.html
git commit -m "feat(frontend): visor-unidades muestra i1-i4-FINAL por criterio (read-only)"
```

---

### Task 6: Verificación final + build

**Files:** ninguno nuevo.

- [ ] **Step 1: `npm run tauri:dev` — test manual completo**

1. Inicia dev: `node scripts/prepare-tauri-web.js && npm run tauri:dev`.
2. Abre index.html → botón "Introducir notas" → gestor-notas.
3. Selecciona un Excel + unidad.
4. **Flujo A (gestor-notas):** 
   - Teclea valores en i1, i2, i3, i4 de varios criterios.
   - Verifica que FINAL se calcula (si lo calcula el backend al guardar, recarga para ver el valor).
   - Teclea Rec.
   - Guarda → sin errores.
   - Recarga → valores persisten.
5. **Flujo B (gestor-recuperaciones):**
   - Selecciona una unidad.
   - Verifica que solo ve columna Rec (sin i1-i4).
   - Edita varios Rec → guarda.
   - Recarga.
6. **Flujo C (gestor-unidades):**
   - Abre "Configurar instrumentos" para una unidad.
   - Elige 4 instrumentos, asigna pesos.
   - Guarda.
   - Recarga gestor-notas → verifica que la fila de cabecera muestra las etiquetas nuevas (si se cargan dinámicamente).
7. **Flujo D (visor-unidades):**
   - Abre visor-unidades.
   - Selecciona unidad.
   - Verifica que ve i1-i4-FINAL (read-only) con valores correctos.

- [ ] **Step 2: `npm run tauri:build` — build release**

Run: `npm run tauri:build`
Expected: compila sin errores, genera exe en `exe/ESO_notas_Actividades_*.exe`.

- [ ] **Step 3: Test en el exe**

1. Instala el exe.
2. Ejecuta la aplicación.
3. Repite flujos A-D (gestor-notas, gestor-recuperaciones, gestor-unidades, visor-unidades).

- [ ] **Step 4: Commit final**

Si quedan cambios sueltos:

```bash
git status --short
git add -A
git commit -m "chore(frontend): cierre plan instrumentos por criterio"
```

- [ ] **Step 5: Resumen de cambios**

Verificar que se han tocado exactamente estas páginas:
- `app-bridge.js` — wrapper nuevo
- `gestor-notas.html` — 4 inputs por CR + FINAL
- `gestor-recuperaciones.html` — solo Rec por CR
- `gestor-unidades.html` — modal de instrumentos/pesos
- `visor-unidades.html` — display i1-i4-FINAL

Ninguna otra página toca notas (visor-notas.html, informes.html, diario.html no se modifican).

---

## Qué queda fuera de este plan

- Autosave por celda en gestor-notas (puede implementarse después — igual que hoy).
- Tests unitarios en JS (vanilla, sin framework — los tests son via manual o e2e).
- Actualización de CLAUDE.md/README — tarea de documentación separada.
