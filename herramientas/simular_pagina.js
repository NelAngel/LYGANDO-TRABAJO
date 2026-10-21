// Simula la página REAL (script.js + index.html) con un DOM mínimo y los Excels
// locales, e imprime cuántas tarjetas se pintan. Sirve para verificar los
// cambios sin abrir el navegador.
// Uso:  node herramientas/simular_pagina.js
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
  location: { protocol: 'http:', hash: '' },
  window: { scrollTo() {}, addEventListener() {} },
  history: { scrollRestoration: 'manual', replaceState() {} },
  XLSX, console,
  setTimeout, clearTimeout, setInterval, clearInterval,
  Math, Date, JSON, Object, Array, String, Number, Boolean, Promise, RegExp, Error, isNaN, parseInt, parseFloat,
  fetch(url) {
    const p = path.join(BASE, url);
    if (!fs.existsSync(p)) return Promise.resolve({ ok: false, arrayBuffer: async () => null });
    return Promise.resolve({
      ok: true,
      arrayBuffer: async () => {
        const b = fs.readFileSync(p);
        return b.buffer.slice(b.byteOffset, b.byteOffset + b.byteLength);
      },
    });
  },
};
sandbox.globalThis = sandbox;
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(BASE, 'script.js'), 'utf8'), sandbox, { filename: 'script.js' });

setTimeout(() => {
  const lista = els['lista'] ? els['lista'].innerHTML : '';
  const cursos = els['listaCursos'] ? els['listaCursos'].innerHTML : '';
  console.log('MODO http (lee los Excel con fetch)');
  console.log('banner:', els['bannerResumen'] && els['bannerResumen'].textContent);
  console.log('stTotal:', els['stTotal'] && els['stTotal'].textContent,
              '| enlace:', els['stEnlace'] && els['stEnlace'].textContent,
              '| whatsapp:', els['stWhatsapp'] && els['stWhatsapp'].textContent,
              '| llamada:', els['stLlamada'] && els['stLlamada'].textContent,
              '| zonas:', els['stZonas'] && els['stZonas'].textContent);
  console.log('tarjetas en #lista:', (lista.match(/<article class="tarjeta/g) || []).length);
  console.log('destacadas:', els['destacado'] ? (els['destacado'].innerHTML.match(/<article/g) || []).length : 0);
  console.log('cursos:', (cursos.match(/<article class="tarjeta/g) || []).length);
  (lista.match(/<h3 class="titulo">[^<]*/g) || []).slice(0, 5).forEach(t => console.log('  ·', t.replace(/<[^>]*>/g, '')));
}, 2500);
