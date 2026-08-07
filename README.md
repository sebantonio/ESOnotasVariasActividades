# ESO Notas Local

Aplicacion de escritorio para gestionar calificaciones de ESO sobre una plantilla Excel local. Permite mantener alumnos, unidades, competencias especificas, criterios, instrumentos de evaluacion, actividades, recuperaciones, informes y diario de clase desde una interfaz HTML empaquetada con Tauri.

El proyecto esta pensado para trabajar sin servidor externo: los datos se leen y escriben directamente en un archivo `.xlsx` seleccionado por el usuario.

## Funcionalidades

- Seleccion de plantilla Excel de trabajo.
- Gestion de alumnos.
- Gestion de unidades didacticas por evaluacion.
- Gestion de instrumentos de evaluacion.
- Gestion de competencias especificas y criterios.
- Introduccion de notas por unidad y actividad.
- Introduccion de notas de recuperacion por evaluacion.
- Visualizacion de notas por unidad.
- Visualizacion de notas por evaluacion.
- Informes por alumno.
- Diario de clase.
- Importacion de datos desde CSV o Excel para alumnos, criterios, instrumentos y unidades.
- Generacion de instalador Windows mediante Tauri.

## Stack tecnico

- Frontend: HTML, CSS y JavaScript vanilla.
- Escritorio: Tauri v2.
- Backend local: Rust.
- Lectura Excel: `calamine`.
- Escritura Excel: manipulacion directa del ZIP/XML interno del `.xlsx`.
- Utilidades JS: `xlsx`, `jszip`.
- Empaquetado alternativo legado: Electron y `electron-builder`.

## Requisitos

- Windows.
- Node.js y npm.
- Rust estable.
- Dependencias de Tauri v2 para Windows.

Para desarrollo con Tauri debe estar disponible el CLI de Tauri instalado por npm:

```powershell
npm install
```

> Nota: en este repositorio `package-lock.json` esta ignorado actualmente, por lo que la instalacion puede resolver versiones compatibles dentro de los rangos definidos en `package.json`.

## Uso

1. Abre la aplicacion.
2. Selecciona el archivo Excel de la plantilla.
3. Usa las secciones principales:
   - Alumnos.
   - CE y criterios.
   - Unidades.
   - Instrumentos.
   - Introducir notas.
   - Recuperaciones.
   - Visualizar notas.
   - Informes.
   - Diario.

La aplicacion guarda los cambios en el Excel seleccionado. Es recomendable conservar una copia de seguridad de la plantilla antes de hacer cambios masivos.

## Comandos

Preparar los archivos web que usa Tauri:

```powershell
node scripts/prepare-tauri-web.js
```

Ejecutar en modo desarrollo:

```powershell
npm run tauri:dev
```

Compilar instalador Windows:

```powershell
npm run tauri:build
```

El comando de build:

1. Incrementa automaticamente la version en `package.json`, `src-tauri/tauri.conf.json` y `src-tauri/Cargo.toml`.
2. Prepara la carpeta `tauri-web/`.
3. Ejecuta `tauri build`.
4. Copia el instalador generado a `exe/`.

Tambien existen comandos Electron:

```powershell
npm start
npm run pack
npm run dist
```

La ruta principal recomendada para esta version del proyecto es Tauri.

## Estructura del proyecto

```text
.
+-- index.html                         Pantalla principal
+-- gestor-alumnos.html                Gestion de alumnos
+-- gestor-rraa-criterios.html         Gestion de CE y criterios
+-- gestor-unidades.html               Gestion de unidades
+-- gestor-instrumentos.html           Gestion de instrumentos
+-- gestor-notas.html                  Introduccion de notas por unidad
+-- gestor-recuperaciones.html         Recuperaciones
+-- visor-notas.html                   Vista de notas por evaluacion
+-- visor-unidades.html                Vista de notas por unidad
+-- visor-actividades.html             Vista de actividades
+-- incluir-actividad.html             Alta de actividades
+-- informes.html                      Informes por alumno
+-- diario.html                        Diario de clase
+-- app-bridge.js                      Puente JS hacia comandos Tauri
+-- main.js                            Entrada Electron legado
+-- preload.js                         Preload Electron legado
+-- tauri-node-backend.js              Backend Node legado
+-- scripts/
|   +-- prepare-tauri-web.js           Copia HTML/JS a tauri-web y actualiza version visible
|   +-- bump-version.js                Incrementa version antes del build
|   +-- copy-exe.js                    Copia instalador a exe/
+-- src-tauri/
|   +-- src/main.rs                    Backend Rust y comandos Tauri
|   +-- Cargo.toml                     Dependencias Rust
|   +-- tauri.conf.json                Configuracion Tauri
+-- tauri-web/                         Salida web generada para Tauri
+-- test-data/                         CSV de prueba
+-- memory/                            Notas internas del proyecto
```

## Plantilla Excel

La aplicacion trabaja sobre una plantilla Excel con estructura fija. La hoja principal es `DATOS`.

Rangos relevantes:

| Seccion | Rango Excel | Notas |
| --- | --- | --- |
| Alumnos | `A4:B41` | Maximo 37 alumnos |
| Unidades | `I5:K20` | Codigo, nombre y evaluacion |
| Instrumentos | `N4:O13` | Maximo 10 instrumentos |
| CE y criterios | Hoja `PESOS` | Ponderaciones por unidad |

Columnas de unidades:

| Columna | Contenido |
| --- | --- |
| `I` | Codigo |
| `J` | Nombre |
| `K` | Evaluacion |

Columnas de instrumentos:

| Columna | Contenido |
| --- | --- |
| `N` | Abreviatura |
| `O` | Nombre |

Hojas de evaluacion esperadas:

- `1a EVA`
- `2a EVA`
- `3a EVA`
- `FINAL`
- Variantes como `2a EVA-solo` o `3a EVA-solo`, si existen en la plantilla.

En las hojas de evaluacion:

| Fila/columna | Uso |
| --- | --- |
| Fila 17 | Cabecera de criterios y notas |
| Fila 18 | Subetiquetas de recuperacion |
| Fila 19+ | Datos de alumnos |
| Columna `CB` | Nota final |

## Importacion de datos

Desde la pantalla principal se pueden importar datos externos:

- Alumnos desde CSV o Excel.
- CE y criterios desde CSV o Excel.
- Instrumentos desde CSV o Excel.
- Unidades desde CSV o Excel.

La importacion esta orientada a copiar datos al libro activo. Conviene revisar el Excel seleccionado antes de importar.

## Datos y persistencia

- El archivo Excel seleccionado se mantiene en memoria durante la sesion de la aplicacion.
- Algunas pantallas usan `localStorage` para recordar archivos recientes.
- Si se reinicia la aplicacion y no se encuentra un archivo por defecto junto al ejecutable, puede ser necesario seleccionar de nuevo el Excel.
- No hay base de datos externa.

## Build y distribucion

El instalador se genera con:

```powershell
npm run tauri:build
```

La configuracion Tauri esta en:

```text
src-tauri/tauri.conf.json
```

El instalador copiado queda en:

```text
exe/
```

La carpeta `exe/` esta ignorada por Git.

## Archivos ignorados

El `.gitignore` actual excluye:

```text
node_modules/
src-tauri/target/
package-lock.json
*.xlsx
exe/
```

Esto evita subir dependencias, compilaciones, instaladores y archivos Excel reales.

## Consideraciones de mantenimiento

- `src-tauri/src/main.rs` concentra la mayor parte de la logica de lectura y escritura Excel.
- `tauri-web/` se genera desde los HTML raiz con `scripts/prepare-tauri-web.js`.
- La version visible de la app se reemplaza en los HTML durante la preparacion de `tauri-web/`.
- El comando `npm run tauri:build` incrementa la version automaticamente.
- Si se cambia la estructura de la plantilla Excel, hay que revisar los rangos fijos usados por el backend.

## Recomendaciones tecnicas pendientes

- Versionar `package-lock.json` para builds npm reproducibles.
- Revisar la politica CSP de Tauri antes de distribuir ampliamente.
- Evitar dependencias cargadas desde CDN en la app de escritorio.
- Validar de forma estricta las rutas y extensiones de archivos recibidas desde el frontend.
- Limitar los enlaces permitidos en la funcion que abre URLs externas.
- Decidir si `tauri-web/` debe seguir versionado o tratarse como artefacto generado.

## Licencia

No se ha definido una licencia en este repositorio. Si el proyecto se va a publicar o compartir, conviene anadir un archivo `LICENSE`.
