# Menú VBA en Excel — Plan

Réplica del programa exe (Tauri) como formulario VBA embebido en `CCGG PLANTILLA - RECUv45.xlsx`, que se abre solo al abrir el libro.

## Pasos

1. Guardar el Excel como `.xlsm` (Excel habilitado para macros).
2. Archivo → Opciones → Personalizar cinta → activar **Programador**.
3. Programador → Seguridad de macros → marcar "Confiar en el acceso al modelo de objetos de proyectos VBA".
4. Alt+F11 → editor VBA.
5. Insertar → UserForm (menú principal) + Insertar → Módulo (código común).
6. En `ThisWorkbook`, evento `Workbook_Open` → `UserForm1.Show`.
7. Pegar ahí el código generado por Claude en Excel.
8. Guardar, cerrar y reabrir el `.xlsm` para probar.

## Prompt para Claude en Excel

```
Quiero un menú de formulario VBA (UserForm) que se abra automáticamente al abrir
este libro Excel (evento Workbook_Open), para gestionar notas de ESO sin salir
de Excel. Estructura real del libro:

HOJA "DATOS":
- Alumnos: rango A4:B41 (fila 4 = cabecera, datos filas 5-41, máx 37 alumnos)
- Unidades: rango I5:K20 → I=código, J=nombre, K=evaluación (1ª/2ª/3ª)
- Instrumentos: rango N4:O13 → N=abreviatura, O=nombre (máx 10)

HOJA "PESOS":
- Fila 4 = cabecera con códigos CR (CR1.1, CR1.2, CR2.1...) por columna
- Filas 5-20 = ponderaciones (%) por unidad, columna A = nombre unidad

HOJAS DE EVALUACIÓN ("1ª EVA","2ª EVA","3ª EVA","FINAL","2ª EVA-solo","3ª EVA-solo"):
- Fila 17 = cabecera: NOTA CE | CR1.1 | Rec | CR1.2 | Rec | ... | NOTA FINAL
- Fila 18 = subetiquetas "Rec"
- Fila 19+ = datos de alumnos
- Las columnas CR y Rec son FÓRMULAS que agregan desde las hojas de unidad
  (ej: =IF('U1'!$A$4="1ª",'U1'!C5,...)) — NUNCA escribir valores directos ahí,
  se recalculan solas.

HOJAS DE UNIDAD (U1, U2, U3...):
- Cada unidad tiene su propia hoja con columnas CR1.1, CR1.2... empezando en
  columna A, y ahí SÍ se escriben notas directamente.
- Junto a cada CR hay columna "Rec" (recuperación) editable.

FUNCIONALIDAD QUE NECESITO EN EL FORMULARIO (menú con botones, uno por módulo):

1. Gestor de alumnos: alta/baja/editar en A4:B41 de DATOS.
2. Gestor de unidades: alta/baja/editar en I5:K20 de DATOS.
3. Gestor de instrumentos: alta/baja/editar en N4:O13 de DATOS (máx 10).
4. Gestor de CE y criterios: editar cabeceras/ponderaciones en hoja PESOS.
5. Introducir notas por unidad: formulario tipo tabla, alumno x criterio CR,
   escribe en la hoja de unidad correspondiente (Un), respetando que Rec va
   en la columna adyacente al CR.
6. Introducir recuperaciones: igual que el punto 5 pero centrado en columnas Rec.
7. Visor de notas por evaluación: solo lectura, lee de la hoja de evaluación
   correspondiente (usar el layout de fila 17/18 arriba descrito).
8. Informe final por alumno: resumen de notas CE/CR/Final de todas las unidades.

Quiero que generes:
- Un UserForm "MenuPrincipal" con botones a cada módulo.
- Un UserForm por módulo (o subrutinas que muestren/oculten controles).
- Módulo estándar con funciones auxiliares reutilizables (buscar fila alumno,
  buscar columna CR, validar notas 0-10, etc.)
- Código para Workbook_Open que abra MenuPrincipal.

Empieza SOLO por el punto 1 (gestor de alumnos) y el menú principal con los
botones (aunque los demás aún no funcionen), para poder probarlo antes de
seguir con el resto módulo a módulo.
```

## Nota

Pedir por fases: alumnos → unidades → instrumentos → notas → recuperaciones → informes. Todo de golpe genera código roto o incompleto.
