# Design — Gestor de Notas ESO

Sistema de diseño bloqueado para esta app. Cada página redisecada lee este
archivo antes de tocar su CSS. No se regenera por página — se amplía este
archivo cuando el sistema necesita crecer.

**Contexto del rediseño:** app de escritorio Tauri (offline, Windows), 14
páginas HTML que ya comparten `ux-common.css`. No es un sitio de marketing —
son pantallas funcionales (tablas de notas, formularios, informes) que usa un
profesor de ESO a diario. El objetivo no es "impresionar", es que se sienta
cuidada, coherente y rápida de leer bajo presión (introduciendo notas entre
clases).

## Género
modern-minimal — herramienta profesional de datos. Sin hero de marketing, sin
CTA de venta, sin testimonios. La "portada" (`index.html`) es la única pantalla
con algo de personalidad de bienvenida; el resto prioriza función.

## Adaptación de "macrostructure" (app, no landing page)

Esta app no tiene páginas de marketing, así que los 21 macrostructures del
catálogo Hallmark (pensados para hero/CTA/testimonios) no aplican literalmente.
En su lugar, el "macrostructure" de cada tipo de página es:

- **Portada** (`index.html`): tratamiento tipo Marquee ligero — título grande
  del centro/instituto, un bloque de bienvenida, grid de accesos (antes
  "botones de menú", ahora tarjetas con jerarquía visual real).
- **Páginas de gestión/edición** (`gestor-*.html`, `incluir-actividad.html`,
  `diario.html`): patrón "Workbench" adaptado — cabecera fija con navegación,
  barra de herramientas, tabla o formulario como contenido principal, barra de
  guardado flotante donde aplica.
- **Páginas de solo lectura** (`visor-*.html`, `informes.html`): mismo patrón
  Workbench sin barra de guardado; énfasis en legibilidad de tabla (columna
  alumno sticky, cabeceras con contraste alto).
- **Utilidades** (`utilidades.html`): patrón Index-First — es literalmente una
  lista de accesos a herramientas, no necesita tabla ni formulario.

## Paleta (OKLCH, ancla en el indigo/teal que la app ya usa)

No se sustituye la marca — se refina. `--ux-primary #4f46e5` y
`--ux-accent #0f766e` ya eran la identidad real de la app; aquí se llevan a
OKLCH y se hacen consistentes (antes cada página tenía su propio hex suelto,
ej. `#6366f1` en gestor-recuperaciones.html vs `#4f46e5` en ux-common.css).

```
--color-paper        oklch(97% 0.01 260)     /* fondo app, existente #f3f6fb */
--color-paper-2       oklch(98% 0.007 275)    /* superficie muted, #f8f9ff */
--color-surface       oklch(100% 0 0)         /* tarjetas/tablas, blanco puro */
--color-ink           oklch(28% 0.03 262)     /* texto principal, #1f2937 */
--color-ink-2         oklch(55% 0.03 258)     /* texto secundario, #64748b */
--color-rule          oklch(89% 0.014 253)    /* bordes, #dbe3ef */
--color-rule-strong   oklch(84% 0.06 274)     /* bordes marcados, #c7d2fe */
--color-accent        oklch(51% 0.22 277)     /* indigo primario, #4f46e5 */
--color-accent-strong oklch(36% 0.15 278)     /* indigo hover/strong, #3730a3 */
--color-accent-soft   oklch(96% 0.02 279)     /* fondo indigo suave, #eef2ff */
--color-teal          oklch(48% 0.09 180)     /* acento secundario (CE/evaluación), #0f766e */
--color-focus         oklch(60% 0.19 277)     /* anillo de foco, mas claro que accent */
--color-danger        oklch(44% 0.18 27)      /* #b91c1c */
--color-warning       oklch(47% 0.14 46)      /* #b45309 */
--color-success       oklch(39% 0.10 152)     /* #166534 */
```

Estos tokens **sustituyen** los `--ux-*` de `ux-common.css`; los nombres
`--ux-*` se mantienen como alias (`--ux-primary: var(--color-accent)`) para no
romper ninguna página que aún no se haya migrado en una pasada posterior.

## Tipografía — solo fuentes de sistema (offline, Windows)

La app corre sin conexión; no se cargan fuentes de Google Fonts ni se
empaquetan archivos de fuente nuevos (evita romper el arranque offline y no
añade peso al instalador). Se usa el stack que Windows 10/11 ya trae:

- **Display** (h1/h2, títulos de página): `"Segoe UI Semibold", "Segoe UI", system-ui, sans-serif` — mismo grotesco humanista, peso mayor en vez de una familia distinta. Es una pareja legítima "por peso", no una que evite el trabajo: en un stack solo-sistema, cambiar de familia significaría arriesgar que la fuente no esté instalada.
- **Body**: `"Segoe UI", system-ui, sans-serif`, 400/600.
- **Tabular/mono** (notas, columnas de números): `"Cascadia Mono", Consolas, "Courier New", monospace` — alinea dígitos en las tablas de notas, detalle real de calidad para una app de datos.

```
--font-display: "Segoe UI Semibold", "Segoe UI", system-ui, sans-serif;
--font-body:    "Segoe UI", system-ui, sans-serif;
--font-mono:    "Cascadia Mono", Consolas, "Courier New", monospace;
```

## Espaciado (escala 4pt)

```
--space-3xs: 0.25rem;  --space-2xs: 0.5rem;  --space-xs: 0.75rem;
--space-sm:  1rem;     --space-md:  1.5rem;  --space-lg: 2rem;
--space-xl:  3rem;     --space-2xl: 4.5rem;  --space-3xl: 7rem;
```

## Escala tipográfica

```
--text-xs: 0.75rem;   --text-sm: 0.875rem;  --text-md: 1rem;
--text-lg: 1.125rem;  --text-xl: 1.375rem;  --text-2xl: 1.75rem;
--text-display: 2.25rem;   /* h1 de portada, unico sitio que la usa a este tamano */
```

## Movimiento

```
--ease-out:    cubic-bezier(0.16, 1, 0.3, 1);
--ease-in:     cubic-bezier(0.7, 0, 0.84, 0);
--ease-in-out: cubic-bezier(0.65, 0, 0.35, 1);
--dur-short:   150ms;
--dur-med:     220ms;
```

- Reveal: ninguno por defecto (es una herramienta de trabajo, no una landing).
  Únicamente transiciones de estado (hover/focus/guardado) y el toggle de
  autoguardado ya existente (`.saved-rec` pulse).
- `prefers-reduced-motion: reduce` → todas las transiciones de opacidad/estado
  caen a ≤150ms, sin transform.

## Postura de microinteracciones

- Éxito silencioso: el toast `.app-ux-status` existente (no modales de "guardado
  correcto").
- Guardado optimista + estado visual (`.saved-rec`, `.invalid-rec`) — ya
  implementado, se refina visualmente, no se cambia el mecanismo.
- Foco visible instantáneo (`:focus-visible`, sin transición de aparición).
- Hover en botones: cambio de fondo + `translateY(-1px)`, sin sombra
  añadida agresiva. Active: `translateY(0)`.

## Voz de CTA

- Primario: relleno `--color-accent`, texto blanco, `--radius-input`, peso 700.
- Secundario: borde `--color-rule-strong`, texto `--color-accent`, fondo
  `--color-surface`.
- Peligro (borrar, descartar): borde/fondo sutil en `--color-danger`, nunca
  relleno sólido salvo confirmación explícita en modal.

## Radios

```
--radius-card:  10px;
--radius-pill:  999px;
--radius-input: 7px;
```

## Qué deben compartir todas las páginas
- Los tokens de color/tipografía/espaciado/movimiento de arriba.
- La cabecera de navegación (`.ux-page-nav`): mismo alto, mismo tratamiento.
- La voz de botón primario/secundario.
- Las badges de nota (`.ux-grade` ok/warn/bad/empty) con los mismos 4 colores
  semánticos en toda la app.
- El foco visible y el toast de estado.

## Qué puede variar por página
- Composición del contenido principal (tabla ancha con columna sticky vs.
  formulario vs. grid de tarjetas en `index.html`/`utilidades.html`).
- Densidad (gestor-notas.html ya tiene modo compacto — se conserva).
- `index.html` es la única página con `--text-display` y algo de calidez
  extra en la bienvenida; el resto se queda en `--text-2xl` como techo.

## Qué NO cambia (fuera de alcance de este rediseño)
- Lógica de la app, IPC, cálculo de notas, guardado en Excel.
- Estructura de rutas/archivos HTML.
- El mecanismo de modo oscuro (`filter: invert()` en `ux-common.css`) — se
  mantiene funcionalmente, solo se ajustan los colores base que invierte.
- `asteroides.html` — no usa `ux-common.css`, queda fuera de este sistema.
