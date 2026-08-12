# Spec: Panel de Ponderaciones de Instrumentos en Gestor-Notas

**Fecha**: 2026-08-11  
**Componente**: gestor-notas.html  
**Estado**: Diseño aprobado

## Objetivo

Hacer visible y accesible la configuración de ponderaciones de instrumentos directamente en la página de introducción de notas por unidad, eliminando la necesidad de ir a gestor-unidades.html para este task.

## Diseño UI

### Ubicación
Panel collapsible (expandido por defecto) ubicado entre `.controls` y `.table-wrapper` en gestor-notas.html.

### Estructura del Panel

```
┌─ Ponderaciones de Instrumentos ▼ ─────────────────────────────┐
│                                                                  │
│  Instrumento 1: [select] ────── Peso: [input] %                │
│  Instrumento 2: [select] ────── Peso: [input] %                │
│  Instrumento 3: [select] ────── Peso: [input] %                │
│  Instrumento 4: [select] ────── Peso: [input] %                │
│                                                                  │
│  Suma total: 100% ✓                                             │
│                                                  [Guardar] [Cancelar] │
└────────────────────────────────────────────────────────────────┘
```

### Estados Visuales
- **Expandido por defecto**: Panel visible al cargar la unidad
- **Colapsable**: Click en encabezado toglea visibilidad (icono chevron)
- **Suma válida (100%)**: Texto verde con checkmark
- **Suma inválida**: Texto naranja/rojo, aviso "⚠ Total debe ser 100%"
- **Cambios sin guardar**: Indicador visual en encabezado (ej: asterisco o badge)

## Datos & Flujo

### Carga
1. Usuario selecciona unidad en `.controls` → trigger `loadNotasUnidad()`
2. Respuesta contiene `data.instrumentosUnidad` (array de {instrumento, peso})
3. Panel se rellena con esos valores (igual lógica que gestor-unidades.html modal)

### Edición
1. Usuario cambia select (instrumento) o input (peso)
2. Validación en tiempo real: suma pesos, actualiza indicador color/aviso
3. Cambios se marcan en el panel (indicador "hay cambios sin guardar")

### Guardar
- Click botón "Guardar" → llama `excel_save_instrumentos()` con estructura:
  ```json
  {
    "unidad": "U1",
    "instrumentosUnidad": [
      { "instrumento": "EX", "peso": 40 },
      { "instrumento": "TR", "peso": 30 },
      { "instrumento": "OB", "peso": 30 },
      { "instrumento": "", "peso": 0 }
    ]
  }
  ```
- Respuesta exitosa → borra indicador "cambios sin guardar", muestra éxito
- Nota: valores de peso son **números directos** (40 = 40%), NO se multiplican por 100

### Cancelar
- Click "Cancelar" → recarga valores originales desde `currentState.instrumentosUnidad`
- Limpia indicador "cambios sin guardar"
- No toca Excel

## Técnica

### HTML
- Agregar sección `.ponderaciones-panel` con estructura:
  ```html
  <div class="ponderaciones-panel">
    <div class="ponderaciones-header" onclick="togglePonderaciones()">
      <span>Ponderaciones de Instrumentos</span>
      <span class="chevron">▼</span>
      <span class="unsaved-indicator" id="unsavedIndicator"></span>
    </div>
    <div class="ponderaciones-content" id="ponderacionesContent">
      <div class="slots-grid">
        <!-- 4 slots: cada uno select + input -->
      </div>
      <div class="suma-row">
        <span id="sumaPesos">0%</span>
        <span id="sumaPesosAviso"></span>
      </div>
      <div class="button-row">
        <button onclick="guardarPonderaciones()">Guardar</button>
        <button onclick="cancelarPonderaciones()">Cancelar</button>
      </div>
    </div>
  </div>
  ```

### CSS
- `.ponderaciones-panel`: margenes, borde, fondo similar a `.controls`
- `.ponderaciones-header`: flex, clickable, chevron rotable (transform)
- `.slots-grid`: grid 2 columnas (instrumento + peso) o similar gestor-unidades
- `.suma-row`: flex, texto color según validación (verde=100%, naranja=!=100%)
- `.unsaved-indicator`: display:none por defecto, mostrar si hay cambios

### JavaScript
Agregar al contexto global del HTML (dentro del `<script>` existente):

**Estado**:
```js
let ponderacionesModales = {};  // copia de trabajo
let ponderacionesOriginal = {}; // snapshots para cancel
```

**Funciones**:
- `togglePonderaciones()` — toggle `.ponderaciones-content` display + rotar chevron
- `cargarPonderacionesPanel(data)` — rellena slots desde `data.instrumentosUnidad`
- `actualizarSumaPesos()` — suma pesos, actualiza color/aviso
- `guardarPonderaciones()` — llamar IPC `excel_save_instrumentos`, mostrar resultado
- `cancelarPonderaciones()` — recargar valores originales, limpiar indicador
- `marcarPonderacionesCambiadas()` — mostrar indicador "cambios sin guardar"
- `limpiarPonderacionesCambiadas()` — ocultar indicador

**Hooks**:
- Al cargar unidad (línea ~1250 donde se asigna `currentState = data`):  
  `cargarPonderacionesPanel(data);`
- En cada `onchange` de select/input:  
  `marcarPonderacionesCambiadas(); actualizarSumaPesos();`

## Validación

- **Suma de pesos**: debe ser 100% para poder guardar (opcional: bloquear botón si ≠ 100%, o avisar)
- **Instrumentos duplicados**: permitir (el Excel/PESOS lo maneja)
- **Campos vacíos**: permitir (instrumento="", peso=0 es válido en slot 4)

## Testing

Casos de prueba (Playwright):
1. Cargar unidad → panel visible, expandido, relleno con ponderaciones actuales
2. Cambiar instrumento en slot 1 → actualizar suma, marcar cambios
3. Cambiar peso en slot 2 → actualizar suma, validación color
4. Click Guardar → persiste en Excel, indicador desaparece
5. Click Cancelar sin guardar → restaura valores, no toca Excel
6. Suma ≠ 100% → aviso visual, (si requiere) botón Guardar deshabilitado
7. Collapse/expand panel → estado visual correcto
8. Cambiar de unidad sin guardar → avisar/descartar cambios (opcional)

## Dependencias

- `excel_get_notas_unidad()` → ya trae `instrumentosUnidad`
- `excel_save_instrumentos()` → ya existe en Rust/IPC
- `escaparHtml()` / `escapeHtml()` — usar lo que ya existe
- Catálogo de instrumentos — cargar igual que gestor-unidades (IPC `excel_get_instrumentos()`)

## Consideraciones

- **No afecta cálculo de FINAL**: la función `calculateFinal()` ya usa `currentState.instrumentosUnidad`. Al guardar, se recarga la unidad (optional) o se actualiza en memoria.
- **Orden de slots**: mantener igual a gestor-unidades (slots 1-4, posiciones fijas)
- **Sin cambios a Rust/IPC**: reutiliza handlers existentes
