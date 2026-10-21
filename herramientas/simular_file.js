// Simula abrir index.html con DOBLE CLIC (file://): el fetch de los .xlsx está
// bloqueado y la página usa el respaldo embebido en datos.js.
// Uso:  node herramientas/simular_file.js
const fs = require('fs');
const vm = require('vm');
const path = require('path');
const BASE = path.resolve(__dirname, '..');
const XLSX = require(path.join(BASE, 'xlsx.full.min.js'));

function fakeEl(id) {
  return {
    id, innerHTML: '', textContent: '', value: '', dataset: {}, style: {},
    classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } },
    addEventListener() {}, appendChild() {}, querySelectorAll() { return []; },
    querySelector() { return null; }, focus() {}, scrollIntoView() {},
  };
}
const els = {};
const document = {
  getElementById(id) { return els[id] || (els[id] = fakeEl(id)); },
  querySelector() { return { dataset: { filtro: 'todo', zona: '', panel: 'trabajos', segfiltro: '' }, classList: { add() {}, remove() {}, toggle() {}, contains() { return false; } } }; },
  querySelectorAll() { return []; },
  createElement() { return fakeEl('nuevo'); },
  addEventListener() {},
};
const sandbox = {
  document,
  location: { protocol: 'file:', hash: '' },
  window: { scrollTo() {}, addEventListener() {} },
  history: { scrollRestoration: 'manual', replaceState() {} },
  XLSX, console,
  setTimeout, clearTimeout, setInterval, clearInterval,
  Math, Date, JSON, Object, Array, String, Number, Boolean, Promise, RegExp, Error, isNaN, parseInt, parseFloat,
  fetch() { return Promise.resolve({ ok: false, arrayBuffer: async () => null }); },
};
sandbox.globalThis = sandbox;
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(BASE, 'datos.js'), 'utf8'), sandbox, { filename: 'datos.js' });
vm.runInContext(fs.readFileSync(path.join(BASE, 'script.js'), 'utf8'), sandbox, { filename: 'script.js' });

setTimeout(() => {
  const lista = els['lista'] ? els['lista'].innerHTML : '';
  const cursos = els['listaCursos'] ? els['listaCursos'].innerHTML : '';
  console.log('MODO file:// (con respaldo datos.js)');
  console.log('banner:', els['bannerResumen'] && els['bannerResumen'].textContent);
  console.log('tarjetas en #lista:', (lista.match(/<article class="tarjeta/g) || []).length);
  console.log('destacadas:', els['destacado'] ? (els['destacado'].innerHTML.match(/<article/g) || []).length : 0);
  console.log('cursos:', (cursos.match(/<article class="tarjeta/g) || []).length);
  (lista.match(/<h3 class="titulo">[^<]*/g) || []).slice(0, 5).forEach(t => console.log('  ·', t.replace(/<[^>]*>/g, '')));
}, 2500);
