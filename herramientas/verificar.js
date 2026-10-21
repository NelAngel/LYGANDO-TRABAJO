// Verifica la hoja "Ofertas" con la MISMA lógica de la página:
//   - vigencia (10 días / 14 con WhatsApp), regiones, dedupe y campos obligatorios.
// Uso:  node herramientas/verificar.js
const XLSX = require('../xlsx.full.min.js');
const fs = require('fs');
const path = require('path');
const BASE = path.resolve(__dirname, '..');

const REGIONES_PERU = ['madre de dios','la libertad','san martin','amazonas','ancash','apurimac','arequipa','ayacucho','cajamarca','callao','cuzco','cusco','huancavelica','huanuco','ica','junin','lambayeque','lima','loreto','moquegua','pasco','piura','puno','tacna','tumbes','ucayali'];
const CIUDAD_A_REGION = {trujillo:'La Libertad',chiclayo:'Lambayeque',chimbote:'Ancash',huaraz:'Ancash',huancayo:'Junin',tarma:'Junin',jauja:'Junin',pucallpa:'Ucayali',iquitos:'Loreto',tarapoto:'San Martin',moyobamba:'San Martin',abancay:'Apurimac',andahuaylas:'Apurimac',huamanga:'Ayacucho',juliaca:'Puno',chincha:'Ica',pisco:'Ica',nazca:'Ica','cerro de pasco':'Pasco',espinar:'Cusco',condesuyos:'Arequipa',colquemarca:'Cusco',chiloroya:'Cusco',ccapmarca:'Cusco',antapaccay:'Cusco',tintaya:'Cusco'};
const PREFIJOS_CALLE = /(^|[\s,.;\-])(av\.?|avda\.?|avenida|jir\.?|jiron|jr\.?|calle|cal\.?|cda\.?|cuadra|pje\.?|pasaje|psj\.?|mza\.?|manzana|nro\.?|numero|sector|urb\.?|urbanizacion|prolongacion)\s*\.?\s+\w{2,}/gi;
const normalizar = s => String(s||'').normalize('NFD').replace(/[\u0300-\u036f]/g,'').toLowerCase().trim();
function regionDeUbicacion(ubicacion){
  const u = normalizar(ubicacion).replace(PREFIJOS_CALLE,'').replace(/[^a-z0-9 ]/g,' ').replace(/\s+/g,' ').trim();
  if(!u) return '';
  for(const r of REGIONES_PERU) if(new RegExp('(^| )'+r+'( |$)').test(u)) return r;
  for(const c in CIUDAD_A_REGION) if(new RegExp('(^| )'+c+'( |$)').test(u)) return CIUDAD_A_REGION[c];
  return '';
}
const normalizarNumero = v => { let n=String(v||'').replace(/\D/g,''); for(let i=0;i<2;i++){ if(n.startsWith('0051'))n=n.slice(4); if(n.startsWith('51')&&n.length>=11)n=n.slice(2); if(n.startsWith('0')&&n.length>=10)n=n.slice(1);} return n; };
const claveOferta = o => { const t=normalizar(o.titulo); if(!t) return ''; return [t,normalizarNumero(o.whatsapp),normalizar(o.ubicacion),normalizar(o.empresa)].join('|'); };

const hoy = new Date(); hoy.setHours(0,0,0,0);
const diasDesde = iso => { const f=new Date(iso+'T00:00:00'); if(isNaN(f)) return null; return Math.round((hoy-f)/86400000); };

const wb = XLSX.read(fs.readFileSync(path.join(BASE,'ofertas.xlsx')), {type:'buffer'});
const rows = XLSX.utils.sheet_to_json(wb.Sheets['Ofertas'], {defval: ''});

const vistos = new Map(), duplicados = [], nuevas = [];
for(const o of rows){
  const id = Number(o.id);
  if(String(o.visible||'').trim().toLowerCase() !== 'si') continue;
  const d = diasDesde(String(o.fecha||'').slice(0,10));
  if(d===null || d<0 || d > (o.whatsapp ? 14 : 10)) continue;
  const k = claveOferta(o);
  if(vistos.has(k)){ duplicados.push([id, vistos.get(k), o.titulo]); continue; }
  vistos.set(k, id);
  if(id>=104) nuevas.push({id, titulo:o.titulo, ubic:o.ubicacion, region:regionDeUbicacion(o.ubicacion), wa:o.whatsapp, cont:o.contacto, dias:d, empresa:o.empresa});
}
console.log('tarjetas publicables (Ofertas):', vistos.size);
console.log('duplicados exactos ocultos:', duplicados.length, JSON.stringify(duplicados));
console.log('nuevas (id>=104) visibles:', nuevas.length);
console.log('sin region:', nuevas.filter(n=>!n.region).map(n=>n.id+' '+n.ubic));
const porRegion = {};
nuevas.forEach(n=>{ const r=n.region||'NINGUNA'; porRegion[r]=(porRegion[r]||0)+1; });
console.log('por region:', porRegion);
console.log('contacto whatsapp:', nuevas.filter(n=>n.cont==='whatsapp').length, '| llamada:', nuevas.filter(n=>n.cont==='llamada').length);
const incompletas = nuevas.filter(n=>!n.titulo||!n.ubic||!n.wa||!n.cont);
console.log('incompletas (REGLA DE ORO):', incompletas.map(n=>n.id));
const numCount = {};
rows.forEach(o=>{ const n=normalizarNumero(o.whatsapp); if(n) numCount[n]=(numCount[n]||0)+1; });
console.log('numeros repetidos en la hoja:', Object.entries(numCount).filter(([k,v])=>v>1));
