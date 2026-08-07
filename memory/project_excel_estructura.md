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
