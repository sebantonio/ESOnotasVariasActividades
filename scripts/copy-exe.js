const fs = require('fs');
const path = require('path');

const pkg = JSON.parse(fs.readFileSync('package.json', 'utf8'));
const version = pkg.version;

const bundleDir = 'C:\\cargo-target\\plantillaNotas\\release\\bundle\\nsis';
const exeDir = path.join(__dirname, '..', 'exe');

fs.mkdirSync(exeDir, { recursive: true });

const src = path.join(bundleDir, `ESO_notas_Actividades_${version}_x64-setup.exe`);
const dst = path.join(exeDir, `ESO_notas_Actividades_${version}_x64-setup.exe`);

fs.copyFileSync(src, dst);
console.log(`Copiado: exe/ESO_notas_Actividades_${version}_x64-setup.exe`);
