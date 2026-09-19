// LYGANDO TRABAJOS: lee dos Excels (la BD) y muestra los anuncios.
//  - ofertas.xlsx        : anuncios de FOTOS (manual)
//  - convocatorias.xlsx  : convocatorias por SCRAPING (scraper_onpe.py)
const FUENTES = [
  { archivo: 'ofertas.xlsx', _origen: 'Foto' },
  { archivo: 'convocatorias.xlsx', _origen: 'Portaltrabajo' },
  // Cursos Capacita-T (MTPE): se muestran en su propia sección (#cursos),
  // NO aplican la regla de 10 días (los cursos no caducan).
  { archivo: 'cursos.xlsx', _origen: 'Capacita' },
];

Promise.all(FUENTES.map(f =>
  fetch(f.archivo, { cache: 'no-store' })
    .then(res => (res.ok ? res.arrayBuffer() : null))
    .then(buf => {
      if (!buf) return [];
      const libro = XLSX.read(buf, { type: 'array' });
      const hoja = libro.Sheets[libro.SheetNames[0]];
      return (XLSX.utils.sheet_to_json(hoja, { defval: '' }))
        .filter(esOfertaVisible)
        .filter(o => esCurso(o) || esReciente(o))   // solo las ofertas/convocatorias caducan
        .map(o => Object.assign(o, { _origen: f._origen }));
    })
    .catch(() => [])
))
  .then(arrays => {
    const todos = [].concat(...arrays);
    const cursos  = todos.filter(o => o._origen === 'Capacita')
      .sort((a, b) => new Date(serieFecha(b.fecha)) - new Date(serieFecha(a.fecha)));
    const ofertas = dedupeOfertas(todos.filter(o => o._origen !== 'Capacita')
      .sort((a, b) => new Date(serieFecha(b.fecha)) - new Date(serieFecha(a.fecha))));
    renderizar(ofertas);
    renderCursos(cursos);
  })
  .catch(err => {
    document.getElementById('lista').innerHTML =
      '<p class="vacio">No se pudo leer los Excel.<br>' +
      'Abre la página desde tu Live Server de Visual Studio Code (no con doble clic del archivo).</p>';
    console.error('LygandoTrabajos:', err);
  });

function esOfertaVisible(o) {
  if (!o.titulo) return false;
  return !['no', '0', 'false', 'NO'].includes(String(o.visible || '').trim().toLowerCase());
}

// Un registro es CURSO si viene de cursos.xlsx (Capacita-T MTPE). Solo los que
// traen las columnas propias de curso (empezar_curso / ver_ruta) los trato así:
// los cursos NO caducan, así que no pasan por esReciente() en la carga.
function esCurso(o) {
  return !!(o.empezar_curso || o.ver_ruta);
}

// Los enteros de fecha de Excel (45358...) se convierten a formato ISO.
function serieFecha(v) {
  if (!v) return '';
  if (typeof v === 'number' && v > 20000) {
    const d = new Date(Math.round((v - 25569) * 86400 * 1000));
    return d.toISOString().slice(0, 10);
  }
  return String(v).trim();
}

// Una oferta "Con enlace" es la que mantiene algún botón web (PDF o postulación).
function esEnlazada(o) {
  return !!(o.funciones || o.postular);
}

// Días calendario entre la fecha de subida y HOY (medianoche local en ambas).
// Así una oferta subida HOY a las 11 pm sigue siendo "hoy" (0 días), nunca "ayer".
function diasDesdeHoy(iso) {
  const f = new Date(iso + 'T00:00:00');
  if (isNaN(f)) return null;
  const hoy = new Date();
  hoy.setHours(0, 0, 0, 0);
  return Math.round((hoy - f) / 86400000);
}

// Anuncios con más de 10 días de antigüedad no se publican (regla del usuario).
const DIAS_MAX = 10;
function esReciente(o) {
  const iso = serieFecha(o.fecha);
  if (!iso) return false;
  const dias = diasDesdeHoy(iso);
  return dias !== null && dias >= 0 && dias <= DIAS_MAX;
}

// El contacto por llamada (columna `contacto=llamada`) abre el teléfono;
// cualquier otro número se considera WhatsApp.
function esLlamada(o) {
  return String(o.contacto || '').trim().toLowerCase() === 'llamada';
}

// ---------- Anti-duplicados ----------
// Regla (2026-09-17): si TODOS los datos se repiten (puesto + número + lugar +
// empresa) -> es la misma oferta, se publica una sola vez (la más reciente).
// SI ALGO CAMBIA (lugar, puesto o número) no se oculta nada: son ofertas
// parecidas pero distintas, y la página las muestra. El aviso de "cambió algo"
// lo da verificar_duplicados.py al registrar cada foto nueva.
function dedupeOfertas(lista) {
  const vistos = new Set();
  return lista.filter(o => {
    const clave = claveOferta(o);
    if (!clave || vistos.has(clave)) return false;
    vistos.add(clave);
    return true;
  });
}

// Normaliza el teléfono para comparar: se quitan el código de país 51 y un 0
// inicial. Así '997...', '+51 997...' y '0997...' cuentan como el MISMO número.
function normalizarNumero(v) {
  let n = String(v || '').replace(/\D/g, '');
  for (let i = 0; i < 2; i++) {
    if (n.startsWith('0051')) n = n.slice(4);
    if (n.startsWith('51') && n.length >= 11) n = n.slice(2);
    if (n.startsWith('0') && n.length >= 10) n = n.slice(1);
  }
  return n;
}

// Clave de "duplicado COMPLETO": puesto + número + lugar + empresa.
// Si cualquiera de esos datos cambia, la clave cambia y no se oculta la tarjeta.
function claveOferta(o) {
  const tit = normalizar(o.titulo);
  if (!tit) return '';
  return [tit, normalizarNumero(o.whatsapp), normalizar(o.ubicacion), normalizar(o.empresa)].join('|');
}

// ---------- Regiones (departamentos) del Perú ----------
// La página muestra SOLO regiones reales (los 24 departamentos + Callao),
// nunca ciudades ("Trujillo" es La Libertad) ni direcciones de calle.
// Dato de seguridad: puede existir una "Av. Lima" que NO sea la región Lima
// (sería una avenida, p.ej. dentro de Espinar). Por eso jamás se deduce una
// región de un nombre de calle: primero se QUITAN los prefijos de vía y solo
// se busca nombre de departamento o ciudad/provincia conocida.
// Si la hoja trae una columna `region` (departamento explícito), esa manda.
const REGIONES_PERU = [
  'madre de dios', 'la libertad', 'san martin',
  'amazonas', 'ancash', 'apurimac', 'arequipa', 'ayacucho',
  'cajamarca', 'callao', 'cuzco', 'cusco', 'huancavelica', 'huanuco',
  'ica', 'junin', 'lambayeque', 'lima', 'loreto', 'moquegua', 'pasco',
  'piura', 'puno', 'tacna', 'tumbes', 'ucayali',
];
// Ciudades/provincias cuyo nombre NO es el del departamento -> su región.
const CIUDAD_A_REGION = {
  trujillo: 'La Libertad', chiclayo: 'Lambayeque',
  chimbote: 'Ancash', huaraz: 'Ancash',
  huancayo: 'Junin', tarma: 'Junin', jauja: 'Junin',
  pucallpa: 'Ucayali', iquitos: 'Loreto', tarapoto: 'San Martin', moyobamba: 'San Martin',
  abancay: 'Apurimac', andahuaylas: 'Apurimac', huamanga: 'Ayacucho',
  juliaca: 'Puno', chincha: 'Ica', pisco: 'Ica', nazca: 'Ica',
  'cerro de pasco': 'Pasco',
  // Lugares propios del proyecto (minas/zonas cercanas a Espinar).
  espinar: 'Cusco', condesuyos: 'Arequipa', colquemarca: 'Cusco',
  chilloroya: 'Cusco', ccapmarca: 'Cusco', antapaccay: 'Cusco', tintaya: 'Cusco',
};
// Prefijos de vía: quitar "Av. Lima", "Jr. Cusco", "calle Ica", etc. para que
// un nombre de calle jamás se convierta en una región.
const PREFIJOS_CALLE = /(^|[\s,.;\-])(av\.?|avda\.?|avenida|jir\.?|jiron|jr\.?|calle|cal\.?|cda\.?|cuadra|pje\.?|pasaje|psj\.?|mza\.?|manzana|nro\.?|numero|sector|urb\.?|urbanizacion|prolongacion)\s*\.?\s+\w{2,}/gi;

// Región real (departamento) de una oferta. Columna `region` si existe;
// si no, se deduce de `ubicacion` sin inventar nada.
function regionDe(o) {
  const r = String(o.region || '').trim();
  if (r) return r;
  return regionDeUbicacion(o.ubicacion);
}

function regionDeUbicacion(ubicacion) {
  const u = normalizar(ubicacion || '')
    .replace(PREFIJOS_CALLE, '')
    .replace(/[^a-z0-9 ]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (!u) return '';
  // 1) Aparece un nombre de departamento real -> ese es.
  for (const r of REGIONES_PERU) {
    if (new RegExp('(^| )' + r + '( |$)').test(u)) return tituloRegion(r);
  }
  // 2) Ciudad/provincia conocida -> la región a la que pertenece.
  for (const ciudad of Object.keys(CIUDAD_A_REGION)) {
    if (new RegExp('(^| )' + ciudad + '( |$)').test(u)) return CIUDAD_A_REGION[ciudad];
  }
  return '';
}

function tituloRegion(r) {
  return r.split(' ').map(w => (w === 'de' ? w : w[0].toUpperCase() + w.slice(1))).join(' ');
}

function renderizar(ofertas) {
  const zonas = [...new Set(ofertas.map(regionDe).filter(Boolean))];
  document.getElementById('bannerResumen').textContent =
    ofertas.length ? `+${ofertas.length} ofertas en ${zonas.length} regiones del Perú` : 'No hay ofertas por ahora';
  document.getElementById('stTotal').textContent = ofertas.length;
  document.getElementById('stEnlace').textContent = ofertas.filter(esEnlazada).length;
  document.getElementById('stWhatsapp').textContent = ofertas.filter(o => o.whatsapp && !esLlamada(o)).length;
  document.getElementById('stLlamada').textContent = ofertas.filter(o => o.whatsapp && esLlamada(o)).length;
  document.getElementById('stZonas').textContent = zonas.length;

  renderDestacado(ofertas);
  renderChipsZonas(ofertas);
  renderFooterZonas(zonas);
  renderFiltroOrigen(ofertas);
  renderGrid(ofertas);
}

// ---------- Filtro superior por ORIGEN (Foto / Portaltrabajo) ----------
// Equivalente al filtro de instituciones de los cursos: botones de color ARRIBA
// de la lista de trabajos para que el visitante elija su fuente sin bajar.
function renderFiltroOrigen(ofertas) {
  const cont = document.getElementById('filtroOrigen');
  if (!cont) return;
  const foto = ofertas.filter(o => o._origen === 'Foto').length;
  const portal = ofertas.filter(o => o._origen === 'Portaltrabajo').length;
  const boton = (orig, nombre, n, cls, color, activa) =>
    `<button class="segfiltro ${cls}${activa ? ' activa' : ''}" data-segfiltro="${orig}">` +
    `<span class="punto" style="background:${color}"></span>${nombre}<span class="n">${n}</span></button>`;
  cont.innerHTML =
    boton('', 'Todos', ofertas.length, 'segf-todos',
      'linear-gradient(120deg,var(--naranja),var(--aubergine))', true) +
    boton('Foto', 'Fotos', foto, 'segf-foto', 'var(--naranja)', false) +
    boton('Portaltrabajo', 'Portaltrabajo (ONPE)', portal, 'segf-portal', '#7c3aed', false);
  cont.querySelectorAll('.segfiltro').forEach(b => b.addEventListener('click', () => {
    cont.querySelectorAll('.segfiltro').forEach(x => x.classList.remove('activa'));
    b.classList.add('activa');
    filtrar();
  }));
}

// Ofertas destacadas = los 3 trabajos MÁS RECIENTES (últimos subidos), resaltados
// arriba de la lista. Si empatan en fecha, manda el de id mayor (el último en
// registrarse ese día).
function renderDestacado(ofertas) {
  const cont = document.getElementById('destacado');
  if (!ofertas.length) return;
  const recientes = ofertas.slice().sort((a, b) => {
    const d = new Date(serieFecha(b.fecha)) - new Date(serieFecha(a.fecha));
    return d || (Number(b.id) || 0) - (Number(a.id) || 0);
  }).slice(0, 3);
cont.innerHTML = `<div class="destacado-grid">${recientes.map((o, i) => crearTarjetaDestacada(o, i)).join('')}</div>`;
}

function crearTarjetaDestacada(o, i) {
  const tipos = [];
  if (o.whatsapp) tipos.push(esLlamada(o)
    ? '<span class="tipo tipo-llamada"><span class="pt"></span>Llamada</span>'
    : '<span class="tipo tipo-whatsapp"><span class="pt"></span>WhatsApp</span>');
  const chipFecha = fechaChip(o, true);
  const meta = [o.empresa, o.ubicacion].filter(Boolean).map(e).join(' · ');
  return `
    <article class="destacado-tarjeta">
      <div class="dest-top">
        <span class="desta-rango">${i + 1}<i>º</i></span>
        <div class="categoria">${tipos.join('')}${chipFecha}</div>
        <span class="dest-pin">★ DESTACADA</span>
      </div>
      <h3>${e(o.titulo)}</h3>
      ${meta ? `<p class="meta">${meta}</p>` : ''}
      ${o.descripcion ? `<p class="entradilla">${e(o.descripcion)}</p>` : ''}
      ${o.sueldo ? `<span class="sueldo">${e(o.sueldo)}</span>` : ''}
      <div class="botones">${botonFunciones(o)}${botonPostular(o)}${botonWhatsapp(o)}</div>
    </article>`;
}

function renderChipsZonas(ofertas) {
  const conteo = {};
  ofertas.forEach(o => { const z = regionDe(o); if (z) conteo[z] = (conteo[z] || 0) + 1; });
  const top = Object.entries(conteo).sort((a, b) => b[1] - a[1]).slice(0, 12);
  const fila = document.getElementById('chipUbicaciones');
  fila.innerHTML = `<button class="chip activa" data-zona="">Todas las regiones</button>` +
    top.map(([z, n]) => `<button class="chip" data-zona="${e(z)}">${e(z)} (${n})</button>`).join('') +
    (Object.keys(conteo).length > top.length ? `<button class="chip">+${Object.keys(conteo).length - top.length} más</button>` : '');
  fila.querySelectorAll('.chip').forEach(chip => {
    chip.addEventListener('click', () => {
      fila.querySelectorAll('.chip').forEach(c => c.classList.remove('activa'));
      chip.classList.add('activa');
      filtrar();
    });
  });
}

function renderFooterZonas(zonas) {
  document.getElementById('footerZonas').innerHTML =
    zonas.map(z => `<li><a href="#ofertas">${e(z)}</a></li>`).join('') || '<li>Próximamente</li>';
}

function renderGrid(ofertas) {
  const lista = document.getElementById('lista');
  if (!ofertas.length) {
    lista.innerHTML = '<p class="vacio">Aún no hay ofertas publicadas.</p>';
    return;
  }
  lista.innerHTML = ofertas.map(crearTarjeta).join('') +
    '<p class="vacio" id="sinResultados" style="display:none;">No se encontraron ofertas con esos criterios.</p>';
  filtrar();
}

function crearTarjeta(o) {
  const tipos = [];
  if (o.enlace) tipos.push('<span class="tipo tipo-enlace">Enlace</span>');
  if (o.whatsapp) tipos.push(esLlamada(o)
    ? '<span class="tipo tipo-llamada"><span class="pt"></span>Llamada</span>'
    : '<span class="tipo tipo-whatsapp"><span class="pt"></span>WhatsApp</span>');
  const etiquetaFecha = fechaChip(o);
  const meta = [o.empresa, o.ubicacion].filter(Boolean).map(e).join(' · ');
  const texto = normalizar(o.titulo + ' ' + o.empresa + ' ' + o.ubicacion + ' ' + o.descripcion);
  return `
    <article class="tarjeta" data-tipo="${tipoDe(o)}" data-zona="${e(regionDe(o))}" data-origen="${e(o._origen || '')}" data-texto="${e(texto)}">
      ${tipos.join('') || etiquetaFecha ? `<div class="categoria">${tipos.join('')}${etiquetaFecha}</div>` : ''}
      <h3 class="titulo">${e(o.titulo)}</h3>
      ${meta ? `<p class="meta">${meta}</p>` : ''}
      ${o.descripcion ? `<p class="descripcion">${e(o.descripcion)}</p>` : ''}
      ${o.sueldo ? `<span class="sueldo">${e(o.sueldo)}</span>` : ''}
      <div class="botones">${botonFunciones(o)}${botonPostular(o)}${botonWhatsapp(o)}</div>
    </article>`;
}

// ---------- Cursos gratuitos (Capacita-T MTPE) ----------
// Los cursos NO caducan: se muestran siempre (a diferencia de ofertas empeñadas
// en caducar a los 10 días). Tarjeta con duración / certificado y botones propios.
// Los cursos se agrupan por SEGMENTO = entidad que los emite (columna `segmento`:
// MTPE, Cisco, Huawei, Fundación Romero, ABC del BCP). Cada segmento va con su
// cabecera y contador.
const SEGMENTOS_ORDEN = ['MTPE', 'Cisco', 'Huawei', 'Fundación Romero', 'ABC del BCP'];
const SEGMENTO_CLASE = {
  'MTPE': 'seg-mtpe',
  'Cisco': 'seg-cisco',
  'Huawei': 'seg-huawei',
  'Fundación Romero': 'seg-romero',
  'ABC del BCP': 'seg-bcp',
  'Otro': 'seg-otro',
};
// Colores de marca de cada institución (botones del filtro lateral).
const SEGMENTO_COLORES = {
  'MTPE': '#16a34a',
  'Cisco': '#2563eb',
  'Huawei': '#dc2626',
  'Fundación Romero': '#ea580c',
  'ABC del BCP': '#0f766e',
  'Otro': '#6b7280',
};
const SEGMENTO_BTN = {
  'MTPE': 'segf-mtpe',
  'Cisco': 'segf-cisco',
  'Huawei': 'segf-huawei',
  'Fundación Romero': 'segf-romero',
  'ABC del BCP': 'segf-bcp',
  'Otro': 'segf-otro',
};

function renderCursos(cursos) {
  const lista = document.getElementById('listaCursos');
  if (!lista) return;
  if (!cursos.length) {
    lista.innerHTML = '<p class="vacio">Aún no hay cursos gratuitos publicados.</p>';
    return;
  }
  const porSegmento = {};
  cursos.forEach(o => {
    const s = o.segmento || 'Otro';
    (porSegmento[s] = porSegmento[s] || []).push(o);
  });
  const otros = Object.keys(porSegmento).filter(s => SEGMENTOS_ORDEN.indexOf(s) === -1).sort();
  const orden = SEGMENTOS_ORDEN.filter(s => porSegmento[s]).concat(otros);
  lista.innerHTML = orden.map(seg => {
    const listaSeg = porSegmento[seg]
      .slice()
      .sort((a, b) => normalizar(a.titulo).localeCompare(normalizar(b.titulo)));
    const cls = SEGMENTO_CLASE[seg] || 'seg-otro';
    return `
      <div class="segmento" data-segmento="${e(seg)}">
        <h3 class="segmento-titulo ${cls}">${e(seg)} <span class="segmento-n">${listaSeg.length} cursos</span></h3>
        <div class="grid">${listaSeg.map(crearTarjetaCurso).join('')}</div>
      </div>`;
  }).join('') +
    '<p class="vacio" id="sinResultadosCurso" style="display:none;">No se encontraron cursos con esos criterios.</p>';
  renderSegmentosFiltros(porSegmento, orden);
  filtrarCursos();
}

// Botones de colores del panel lateral: "Todos" + cada institución emisora.
// Al pulsar uno, la lista muestra solo ese segmento (sin bajar).
function renderSegmentosFiltros(porSegmento, orden) {
  const cont = document.getElementById('segmentosFiltros');
  if (!cont) return;
  const total = Object.values(porSegmento).reduce((a, l) => a + l.length, 0);
  const boton = (seg, n, activa) => {
    const nombre = seg || 'Todos';
    const cls = seg ? (SEGMENTO_BTN[seg] || 'segf-otro') : 'segf-todos';
    const color = seg ? (SEGMENTO_COLORES[seg] || '#6b7280') : '';
    const punto = color ? `<span class="punto" style="background:${color}"></span>`
      : '<span class="punto" style="background:linear-gradient(120deg,var(--naranja),var(--aubergine))"></span>';
    return `<button class="segfiltro ${cls}${activa ? ' activa' : ''}" data-segfiltro="${e(seg)}">${punto}${e(nombre)}<span class="n">${n}</span></button>`;
  };
  cont.innerHTML = boton('', total, true) +
    orden.map(s => boton(s, porSegmento[s].length, false)).join('');
  cont.querySelectorAll('.segfiltro').forEach(b => {
    b.addEventListener('click', () => {
      cont.querySelectorAll('.segfiltro').forEach(x => x.classList.remove('activa'));
      b.classList.add('activa');
      filtrarCursos();
    });
  });
}

// Filtro de búsqueda de la pestaña CURSOS (misma lógica que las ofertas:
// normaliza tildes, todas las palabras deben coincidir, admite prefijos).
// Respeta el segmento elegido en el panel lateral y oculta/muestra las
// cabeceras de segmento según los cursos visibles.
function filtrarCursos() {
  const busca = document.getElementById('buscadorCursos').value;
  const seg = ((document.querySelector('#segmentosFiltros .segfiltro.activa') || {}).dataset || {}).segfiltro || '';
  let visibles = 0;
  document.querySelectorAll('#listaCursos .segmento').forEach(segEl => {
    const segOk = seg === '' || segEl.dataset.segmento === seg;
    let enSeg = 0;
    segEl.querySelectorAll('.tarjeta').forEach(t => {
      const muestra = segOk && coincide(t.dataset.texto || '', busca);
      t.classList.toggle('oculto', !muestra);
      if (muestra) enSeg++;
    });
    segEl.classList.toggle('oculto', !enSeg);
    visibles += enSeg;
  });
  const aviso = document.getElementById('sinResultadosCurso');
  if (aviso) aviso.style.display = visibles ? 'none' : '';
}

function crearTarjetaCurso(o) {
  const badge = o.certifica && String(o.certifica).toLowerCase() !== 'no'
    ? '<span class="tipo tipo-certificado">Certificado</span>' : '';
  const meta = [o.empresa, o.ubicacion, o.duracion ? 'Duración: ' + o.duracion : '']
    .filter(Boolean).map(e).join(' · ');
  const texto = normalizar(o.titulo + ' ' + o.segmento + ' ' + o.empresa + ' ' + o.ubicacion + ' ' + o.descripcion);
  return `
    <article class="tarjeta tarjeta-curso" data-tipo="curso" data-zona="${e(o.ubicacion)}" data-texto="${e(texto)}" data-segmento="${e(o.segmento || '')}">
      <div class="categoria">${badge || '<span class="tipo">Curso</span>'}</div>
      <h3 class="titulo">${e(o.titulo)}</h3>
      ${meta ? `<p class="meta">${meta}</p>` : ''}
      ${o.descripcion ? `<p class="descripcion">${e(o.descripcion)}</p>` : ''}
      <div class="botones">${botonEmpezar(o)}${botonRuta(o)}</div>
    </article>`;
}

// [ EMPEZAR CURSO ] -> plataforma LMS de Capacita-T (donde se accede al curso).
function botonEmpezar(o) {
  return o.empezar_curso
    ? `<a class="boton boton-postular" href="${e(o.empezar_curso)}" target="_blank" rel="noopener">EMPEZAR CURSO</a>`
    : '';
}

// [ VER RUTA ] -> ruta formativa del curso (opcional).
function botonRuta(o) {
  return o.ver_ruta
    ? `<a class="boton boton-funciones" href="${e(o.ver_ruta)}" target="_blank" rel="noopener">VER RUTA</a>`
    : '';
}

function tipoDe(o) {
  const t = [];
  if (esEnlazada(o)) t.push('enlace');
  if (o.whatsapp) t.push(esLlamada(o) ? 'llamada' : 'whatsapp');
  return t.join('-') || 'todo';
}

function botonFunciones(o) {
  return o.funciones
    ? `<a class="boton boton-funciones" href="${e(o.funciones)}" target="_blank" rel="noopener">FUNCIONES</a>`
    : '';
}

// Enlace directo al sistema de postulación ("[ POSTULAR ]").
function botonPostular(o) {
  return o.postular
    ? `<a class="boton boton-postular" href="${e(o.postular)}" target="_blank" rel="noopener">POSTULAR</a>`
    : '';
}

function botonWhatsapp(o) {
  if (!o.whatsapp) return '';
  const n = String(o.whatsapp).replace(/\D/g, '');
  const m = encodeURIComponent('Hola, me interesa el puesto de ' + o.titulo);
  return `<a class="boton boton-llamada" href="tel:+51${n}" rel="noopener">LLAMAR</a>` +
    `<a class="boton boton-whatsapp" href="https://wa.me/51${n}?text=${m}" target="_blank" rel="noopener">WHATSAPP</a>`;
}

// Chip de fecha de SUBIDA: referencia visible de cuándo se publicó la oferta.
// La fecha es la del día en que se subió la foto/anuncio (columna `fecha`).
// Al apuntar con el mouse muestra la fecha exacta.
function fechaChip(o, compacto) {
  const iso = serieFecha(o.fecha);
  if (!iso) return '';
  const f = new Date(iso);
  if (isNaN(f)) return '';
  const exacta = f.toLocaleDateString('es-PE', { day: 'numeric', month: 'long', year: 'numeric' });
  const texto = compacto ? fechaSubidaCorta(iso) : fechaSubida(iso);
  return `<span class="tipo tipo-fecha" title="Subido el ${e(exacta)}"><span class="pt"></span>${e(texto)}</span>`;
}

function fechaSubida(iso) {
  const dias = diasDesdeHoy(iso);
  if (dias === null) return '';
  const f = new Date(iso + 'T00:00:00');
  if (dias <= 0) return 'Subido hoy';
  if (dias === 1) return 'Subido ayer';
  if (dias < 7) return `Subido hace ${dias} días`;
  return 'Subido el ' + f.toLocaleDateString('es-PE', { day: 'numeric', month: 'short' });
}

// Versión corta para las tarjetas destacadas (cabe en la fila superior):
// "hoy" / "ayer" / "hace N días" / "5 sep" sin la palabra "Subido".
function fechaSubidaCorta(iso) {
  const dias = diasDesdeHoy(iso);
  if (dias === null) return '';
  const f = new Date(iso + 'T00:00:00');
  if (dias <= 0) return 'hoy';
  if (dias === 1) return 'ayer';
  if (dias < 7) return `hace ${dias} días`;
  return f.toLocaleDateString('es-PE', { day: 'numeric', month: 'short' });
}

function e(s) {
  return String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

// ---------- Análisis de búsqueda ----------
// Normaliza el texto: minúsculas y quita tildes (café -> cafe).
function normalizar(s) {
  return String(s || '')
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase();
}

function tokensDe(s) {
  return normalizar(s).split(/[^a-z0-9]+/).filter(Boolean);
}

// Comprueba que TODAS las palabras de la consulta aparezcan en el texto,
// permitiendo prefijos ("reparti" encuentra "repartidor").
function coincide(textoNorm, consulta) {
  const palabras = tokensDe(consulta);
  if (!palabras.length) return true;
  const palabrasTexto = textoNorm.split(' ');
  return palabras.every(p => {
    if (textoNorm.indexOf(p) !== -1) return true;
    return palabrasTexto.some(w => w.length > 0 && w.startsWith(p));
  });
}

// ---------- Filtros sobre las tarjetas ya renderizadas ----------
function filtrar() {
  const tarjetas = document.querySelectorAll('#lista .tarjeta');
  const busca = document.getElementById('buscador').value;
  const tipo = (document.querySelector('.filtro.activo') || {}).dataset.filtro || 'todo';
  const zona = (document.querySelector('.chip.activa') || { dataset: { zona: '' } }).dataset.zona || '';
  const origen = ((document.querySelector('#filtroOrigen .segfiltro.activa') || {}).dataset || {}).segfiltro || '';

  let visibles = 0;
  tarjetas.forEach(t => {
    const d = t.dataset;
    const tipoOk = tipo === 'todo' || (d.tipo || '').indexOf(tipo) !== -1;
    const zonaOk = zona === '' || d.zona === zona;
    const origenOk = origen === '' || d.origen === origen;
    const textoOk = coincide(d.texto || '', busca);
    const muestra = tipoOk && zonaOk && origenOk && textoOk;
    t.classList.toggle('oculto', !muestra);
    if (muestra) visibles++;
  });

  const aviso = document.getElementById('sinResultados');
  if (aviso) aviso.style.display = visibles ? 'none' : '';
}

document.getElementById('buscador').addEventListener('input', filtrar);
document.getElementById('buscadorCursos').addEventListener('input', filtrarCursos);
document.querySelectorAll('.filtro').forEach(b => {
  b.addEventListener('click', () => {
    document.querySelectorAll('.filtro').forEach(x => x.classList.remove('activo'));
    b.classList.add('activo');
    filtrar();
  });
});

// ---------- Pestañas TRABAJOS / CURSOS ----------
// Cada pestaña muestra su panel y lleva el scroll hasta el campo
// correspondiente (convocatorias/ofertas o cursos Capacita-T).
function activarPestana(panel) {
  document.querySelectorAll('#panelTabs .pestana').forEach(b =>
    b.classList.toggle('activa', b.dataset.panel === panel));
  ['trabajos', 'cursos'].forEach(p => {
    const pane = document.getElementById('pane-' + p);
    if (pane) pane.classList.toggle('oculto', p !== panel);
  });
}

document.getElementById('panelTabs').addEventListener('click', e => {
  const boton = e.target.closest('.pestana');
  if (!boton) return;
  const panel = boton.dataset.panel;
  activarPestana(panel);
  const destino = document.getElementById('pane-' + panel);
  if (destino) destino.scrollIntoView({ behavior: 'smooth', block: 'start' });
});

// Los enlaces ancla de la barra superior / footer también cambian de pestaña.
document.querySelectorAll('a[href^="#"]').forEach(a => {
  a.addEventListener('click', () => {
    const h = a.getAttribute('href');
    activarPestana(h.indexOf('#cursos') === 0 || h.indexOf('pane-cursos') !== -1 ? 'cursos' : 'trabajos');
  });
});

// Si la URL llega con #cursos, abrir directamente la pestaña de cursos.
if (location.hash && location.hash.indexOf('#cursos') === 0) activarPestana('cursos');