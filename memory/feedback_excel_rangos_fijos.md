---
name: Usar rangos fijos para el Excel ESO
description: No buscar headers — los rangos de tablas son fijos y conocidos
type: feedback
---
Usar siempre rangos fijos para leer/escribir el Excel ESO. No usar lógica de búsqueda de headers ni funciones como `find_unidades_start`.

**Why:** El usuario lo indicó explícitamente: "SI eres capaz de leer los instrumentos, los alumnos... Solo tienes que leer la tabla que hay en I5:K20". La búsqueda de headers causó que solo se mostrara la primera unidad durante varias iteraciones.

**How to apply:** En main.rs, leer/escribir directamente por índice de fila/columna. En gestor-unidades.html (modo browser con XLSX.js), también leer `rows[4..19]` cols 8/9/10 directamente.
