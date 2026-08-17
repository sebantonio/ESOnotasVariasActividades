# Diseño: Filtrado de Instrumentos por Ponderación y Alineación UI

**Fecha**: 2026-08-13  
**Alcance**: visor-unidades.html, gestor-notas.html, gestor-recuperaciones.html  
**Objetivo**: Filtrar columnas de instrumentos sin ponderación, mejorar UX, sincronizar UI con patrón de referencia (ESOplantillaNotas)

---

## 1. Problema

- **visor-unidades.html**: No muestra nota final del criterio ni de la unidad (intenta acceder a `crNota['FINAL']` que no existe)
- **gestor-notas.html** y **gestor-recuperaciones.html**: Muestran 5 columnas por criterio (i1/i2/i3/i4/FINAL) aunque algunos instrumentos tengan peso=0, contaminando la UI
- **Inconsistencia**: Patrón UI diferente al proyecto padre (ESOplantillaNotas), que muestra solo CR + Rec

## 2. Solución: Filtrado por Ponderación

### 2.1 Data Structure
```javascript
// Cada criterio tiene asociados pesos de instrumentos:
data.instrumentosUnidad = [
  { abrev: "i1", nombre: "...", peso: 25 },    // i1: 25%
  { abrev: "i2", nombre: "...", peso: 0 },     // i2: sin peso (ocultar)
  { abrev: "i3", nombre: "...", peso: 35 },    // i3: 35%
  { abrev: "i4", nombre: "...", peso: 40 }     // i4: 40%
]
```

### 2.2 Regla de Filtrado
- **Visible**: solo instrumentos donde `peso > 0`
- **Cálculo FINAL**: media ponderada de instrumentos visibles únicamente
- **Redistribución**: NO. Los pesos no se redistribuyen; se usan tal cual
- **Ejemplo**: Si i2=0, FINAL = (i1*25 + i3*35 + i4*40) / 100

## 3. Cambios por Página

### 3.1 visor-unidades.html
**Cambio de renderización: i1/i2/i3/i4/FINAL → CR + Rec**

#### Encabezado (headRow2)
- Por cada criterio, mostrar 2 columnas en lugar de 5:
  - Col 1: CR (código + ponderación)
  - Col 2: Rec (recuperación)
- Omitir criterios donde no haya instrumento con peso > 0 (en la vista, pero los datos siguen llegando)

#### Cuerpo (rows)
- Para cada alumno/criterio:
  - Col 1: `crNota.display` (nota del CR, precalculada)
  - Col 2: `crNota.recDisplay` (recuperación, precalculada)
- Nota Final de la unidad: seguir usando `computeAlumnoNotas(alumno, criterios).final`

**Resultado visual**: Tabla compacta, consistente con ESOplantillaNotas

---

### 3.2 gestor-notas.html
**Cambio: mostrar solo instrumentos activos (peso > 0)**

#### Encabezado (headRow2)
- Construir dinámicamente según `data.instrumentosUnidad` filtrado
- Solo mostrar columnas para i1, i3, i4 (si i2=0)
- Cada instrumento activo: `[i_label (peso%)]`
- Mostrar una sola columna FINAL al final

#### Cuerpo (rows)
- Input fields solo para instrumentos con peso > 0
- Campos vacíos/deshabilitados para instrumentos con peso=0 (ocultarlos)
- Cálculo de FINAL: sumar solo los valores de instrumentos activos, ponderados

#### Ejemplo Layout
```
CR1.1 (30%)  | CR1.2 (20%)     | CR1.3 (50%)  | FINAL
i1   | i3 i4| i1 i2   i3 i4  | 
25%  | 35% 40%| 20% 10% 20% 50%|
[__] | [__][__]| [__][__][__][__] | [auto]
```

---

### 3.3 gestor-recuperaciones.html
**Idem a 3.2**: mostrar solo instrumentos con peso > 0

---

## 4. Cambios de Lógica

### 4.1 Función Helper: `getActiveInstruments(data)`
```javascript
function getActiveInstruments(data) {
  return (data.instrumentosUnidad || []).filter(inst => inst.peso > 0);
}
```
**Usada en**: visor-unidades, gestor-notas, gestor-recuperaciones (renderización + validación)

### 4.2 Cálculo FINAL (actualizar si no lo hace correctamente)
```javascript
function calculateFinal(crNota, instrumentos) {
  const activeInsts = instrumentos.filter(i => i.peso > 0);
  if (!activeInsts.length) return null;
  
  let num = 0, denom = 0;
  activeInsts.forEach(inst => {
    const val = crNota[inst.abrev];
    if (val !== null && val !== undefined && val > 0) {
      num += inst.peso * val;
      denom += inst.peso;
    }
  });
  return denom > 0 ? num / denom : null;
}
```

---

## 5. Puntos Críticos de Implementación

### 5.1 visor-unidades.html
- Línea ~386: cambiar `colspan = g.collapsed ? 1 : g.count * 2 + 1` (antes era `* 5 + 1`)
- Línea ~407: cambiar encabezados a `[CR] [Rec]` (antes era `[i1] [i2] [i3] [i4] [F]`)
- Línea ~432: cambiar lectura de datos a `crNota.display` y `crNota.recDisplay`

### 5.2 gestor-notas.html
- Línea ~403-404: cambiar construcción de `instLabels` a leer desde `getActiveInstruments()`
- Línea ~406-407: mostrar ponderación del instrumento (%, del peso)
- Línea ~1850: placeholder de input debe cambiar dinámicamente según instrumento activo
- Cálculo FINAL: usar función `calculateFinal()` con solo instrumentos activos

### 5.3 gestor-recuperaciones.html
- Idem a 5.2

---

## 6. Testing

- [ ] visor-unidades: cargar unidad con i2=0, verificar que no aparezca columna i2
- [ ] visor-unidades: nota final de unidad se calcula e muestra
- [ ] gestor-notas: introducir notas solo en i1, i3, i4; FINAL calcula sin i2
- [ ] gestor-notas: introducir recuperación, verificar que se guarda
- [ ] gestor-recuperaciones: idem
- [ ] Compatibilidad: si todos los instrumentos tienen peso, comportamiento es idéntico al actual

---

## 7. Archivos a Modificar

1. `tauri-web/visor-unidades.html` (renderTable, headRow1/2, bodyHtml)
2. `tauri-web/gestor-notas.html` (renderTable, headRow2, fieldInput, calculateFinal)
3. `tauri-web/gestor-recuperaciones.html` (renderTable, headRow2, fieldInput, calculateFinal)

---

## 8. No en Scope

- Refactorización de `main.rs` (Rust backend está fuera)
- Cambios en data schema de Excel
- Migración de archivos Excel existentes
- UI de configuración de ponderaciones (ya existe en gestor-instrumentos.html)
