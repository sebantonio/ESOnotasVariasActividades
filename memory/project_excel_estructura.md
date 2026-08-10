---
name: Estructura del Excel ESO
description: Rangos exactos de cada tabla en CCGG PLANTILLA - RECUv45.xlsx
type: project
---
Archivo: `CCGG PLANTILLA - RECUv45.xlsx`

**Hoja DATOS — rangos fijos:**
| Tabla | Rango Excel | 0-indexed (filas, cols) |
|-------|-------------|--------------------------|
| Alumnos | A4:B41 | fila 4=header "Alumnado" (0-idx:3); datos filas 5-41 (0-idx:4-40); max 37 |
| Unidades | I5:K20 | filas 4-19, cols 8(I) 9(J) 10(K) |
| Instrumentos evaluación | N4:O13 | filas 3-12, cols 13(N) 14(O) |

- Col I = código unidad (U1, U2...)
- Col J = nombre unidad
- Col K = evaluación (1ª / 2ª / 3ª)
- Col N = abreviatura instrumento (PE, TD...)
- Col O = nombre instrumento

**Hoja DATOS — CE y CR (cols R-X):**
- R(17)=Nº CE, S(18)=texto CE — header en fila 4 (idx 3), datos desde fila 5 (idx 4)
- V(21)=nº CE (celda combinada), W(22)=código CR (CR1.1...), X(23)=texto CR

**Hoja PESOS:** ponderaciones por CR y unidad (si existe). Los CE/CR NO se leen de PESOS.

**Why:** El Excel ESO tiene una estructura diferente al Excel FP (plantilla313_dual). Los rangos son fijos y no hay que buscarlos por contenido.

**How to apply:** Siempre usar rangos fijos. No buscar headers ni "UNIDADES" por texto. Si el usuario añade columnas nuevas en el Excel, revisar estos rangos.

---

## Nuevo layout: Hojas de unidad (U1..U15) con instrumentos por criterio

*Introducido en ESOnotasVariasActividades (agosto 2026). Archivo: `Plantilla_Notas_ESO.xlsx`*

**Estructura por filas (0-indexed):**

| Fila (Excel) | Fila (0-idx) | Contenido |
|---|---|---|
| 2 | 1 | Etiquetas de slots: `i1`, `i2`, `i3`, `i4` en cols C:F (compartidas por toda la hoja) |
| 3 | 2 | Subcabecera por bloque de criterio: `=i1`, `=i2`, `=i3`, `=i4`, `FINAL` (6 cols/criterio) |
| 4 | 3 | Pesos de los 4 slots: fracción 0-1 en cols C:F (ej. 0.2, 0.4, 0.2, 0.2) |
| 5 | 4 | Código criterio (`CR1.1`) en 1ª col del bloque; `Rec` en col +5 |
| 6 | 5 | Fórmula ponderación criterio (INDEX/MATCH sobre hoja `PESOS`); misma lógica que antes, solo offset |
| 7+ | 6+ | Datos de alumnos: i1-i4 (manual, cols ci..ci+3), FINAL (fórmula SUMPRODUCT, col ci+4, **solo lectura**), Rec (manual, col ci+5) |

**Fórmula FINAL (ejemplo U1!L7 para CR1.1 en bloque 1):**
```
=IFERROR(SUMPRODUCT((H7:K7<>"")*C$4:F$4*H7:K7)/SUMPRODUCT((H7:K7<>"")*C$4:F$4),"")
```
- Media ponderada de i1:i4 (H:K) usando pesos (C:F, fila 4)
- Ignora slots vacíos tanto numerador como denominador
- Devuelve "" si todos los slots están vacíos

**How to apply (backend):**
- Row index de códigos criterio (CR1.1...): fila 5 (0-idx 4)
- Row index de inicio datos alumnos: fila 7 (0-idx 6)
- Ancho de bloque por criterio: 6 columnas (ci, ci+1, ci+2, ci+3, ci+4, ci+5)
- Nunca escribir en col ci+4 (FINAL) — es fórmula, se propaga desde unit a eval sheets
- Leer/escribir i1-i4 en ci..ci+3, Rec en ci+5
- Pesos y etiquetas instrumentos en filas 2-4, cols C:F (compartidos, no por criterio)

**Relación con hojas de evaluación:**
- Las hojas `1ª EVA`, `2ª EVA`, `3ª EVA`, `FINAL` mantienen layout 2-cols/criterio (CR + Rec)
- Formulas en eval sheets referencian el FINAL recién calculado de cada criterio en su hoja de unidad
- Ej. antiguo: `='U1'!B5*'U1'!B$4` → nuevo: `=IF('U1'!L7="","",('U1'!L7*'U1'!H$6))` (con ajustes para nueva posición)
