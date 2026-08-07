const fs = require('fs');
const path = require('path');

const root = path.join(__dirname, '..');
const outDir = path.join(root, 'tauri-web');
const pkg = JSON.parse(fs.readFileSync(path.join(root, 'package.json'), 'utf8'));
const version = pkg.version || '0.0.0';

const files = [
  'index.html',
  'gestor-alumnos.html',
  'gestor-rraa-criterios.html',
  'gestor-unidades.html',
  'gestor-notas.html',
  'gestor-recuperaciones.html',
  'incluir-actividad.html',
  'visor-notas.html',
  'visor-actividades.html',
  'visor-unidades.html',
  'informes.html',
  'diario.html',
  'gestor-instrumentos.html',
  'utilidades.html',
  'asteroides.html',
  'app-bridge.js',
  'ux-common.css',
  'ux-common.js'
];

fs.rmSync(outDir, { recursive: true, force: true });
fs.mkdirSync(outDir, { recursive: true });

files.forEach((file) => {
  const src = path.join(root, file);
  let content = fs.readFileSync(src, 'utf8');
  content = content.replace(/v0\.\d+\.\d+/g, `v${version}`);
  fs.writeFileSync(path.join(outDir, file), content, 'utf8');
});

const binaryFiles = ['notasESOicon.png'];
binaryFiles.forEach((file) => {
  fs.copyFileSync(path.join(root, file), path.join(outDir, file));
});

console.log(`Preparados ${files.length + binaryFiles.length} archivos para Tauri en ${outDir} (v${version})`);
