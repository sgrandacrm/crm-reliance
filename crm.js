// ══════════════════════════════════════════════════════
//  DATA & STATE 
// ══════════════════════════════════════════════════════
const USERS = [
  {id:'admin',name:'Rubi González',email:'rubi@reliance.ec',pass:'admin2026',rol:'admin',color:'#c84b1a',initials:'RG'},
  {id:'belen',name:'Belén Zapata',email:'belen@reliance.ec',pass:'belen2026',rol:'ejecutivo',color:'#1a4c84',initials:'BZ'},
  {id:'juan',name:'Juan Pérez',email:'juan@reliance.ec',pass:'juan2026',rol:'ejecutivo',color:'#2d6a4f',initials:'JP'},
  {id:'maria',name:'María López',email:'maria@reliance.ec',pass:'maria2026',rol:'ejecutivo',color:'#b8860b',initials:'ML'},
  {id:'carlos',name:'Carlos Mora',email:'carlos@reliance.ec',pass:'carlos2026',rol:'ejecutivo',color:'#6f42c1',initials:'CM'},
  {id:'ana',name:'Ana Vega',email:'ana@reliance.ec',pass:'ana2026',rol:'ejecutivo',color:'#fd7e14',initials:'AV'},
  {id:'pedro',name:'Pedro Torres',email:'pedro@reliance.ec',pass:'pedro2026',rol:'ejecutivo',color:'#dc3545',initials:'PT'},
  {id:'lucia',name:'Lucia Salas',email:'lucia@reliance.ec',pass:'lucia2026',rol:'ejecutivo',color:'#17a2b8',initials:'LS'},
  {id:'diego',name:'Diego Ruiz',email:'diego@reliance.ec',pass:'diego2026',rol:'ejecutivo',color:'#28a745',initials:'DR'},
];

// ── Tasas de comisión por aseguradora (editables desde Admin) ────────────────
const COMISIONES_DEFAULT = {
  SWEADEN:15, MAPFRE:12, GENERALI:12, ZURICH:14,
  LATINA:15, ALIANZA:15, 'ASEG. DEL SUR':13, 'AXA ASSISTANCE':10,
};
function _getComisiones(){
  try{ return Object.assign({...COMISIONES_DEFAULT}, JSON.parse(localStorage.getItem('reliance_comisiones')||'{}')); }
  catch(e){ return {...COMISIONES_DEFAULT}; }
}
function _saveComisiones(obj){
  localStorage.setItem('reliance_comisiones', JSON.stringify(obj));
  _flushComisiones(); // sincronizar a SP
}

// Sube comisiones Y tasas a CRM_Comisiones (una fila por aseguradora)
async function _flushComisiones(){
  if(!_spReady) return;
  try{
    const v2    = _getTasasV2();
    const comis = _getComisiones();
    const cache = Array.isArray(_cache.comisiones) ? [..._cache.comisiones] : [];

    // ── 1. Filas V2: una por aseguradora+región+tipo ──────────────────────────
    for(const [crmId, row] of Object.entries(v2)){
      const existing  = cache.find(x => x.crm_id === crmId);
      const comPct    = row.comisionPct !== undefined ? row.comisionPct
                      : (comis[row.aseg] !== undefined ? comis[row.aseg] : (COMISIONES_DEFAULT[row.aseg]||0));
      const lbl = [row.aseg, row.region, row.tipo].filter(Boolean).join(' ');
      const data = {
        Title:      lbl,
        region:     row.region   || '',
        comisionPct: comPct,
        tasa_r1:    row.tasas[0] || 0,
        tasa_r2:    row.tasas[1] || 0,
        tasa_r3:    row.tasas[2] || 0,
        tasa_r4:    row.tasas[3] || 0,
        tasa_r5:    row.tasas[4] || 0,
        limite_r1:  row.limites[0] || 0,
        limite_r2:  row.limites[1] || 0,
        limite_r3:  row.limites[2] || 0,
        crm_id:     crmId,
      };
      try{
        if(existing && existing._spId){
          await spUpdate('comisiones', existing._spId, data);
          const idx = cache.findIndex(x => x.crm_id === crmId);
          if(idx >= 0) cache[idx] = {...cache[idx], ...data};
        } else {
          const id = await spCreate('comisiones', data);
          if(id){ const ni={...data,_spId:id}; const idx=cache.findIndex(x=>x.crm_id===crmId); if(idx>=0) cache[idx]=ni; else cache.push(ni); }
        }
      }catch(eItem){ console.warn('[comisiones] Error sync V2', crmId, eItem.message); }
    }

    // ── 2. Aseguradoras sin fila V2 (sólo comisión, sin tasas) ───────────────
    const asegConV2 = new Set(Object.values(v2).map(r => r.aseg));
    const sinV2 = Object.keys(COMISIONES_DEFAULT).filter(a => !asegConV2.has(a));
    for(const aseg of sinV2){
      const existing = cache.find(x => x.crm_id === aseg || x.Title === aseg);
      const data = {
        Title: aseg, region: '',
        comisionPct: comis[aseg] !== undefined ? comis[aseg] : (COMISIONES_DEFAULT[aseg]||0),
        tasa_r1:0, tasa_r2:0, tasa_r3:0, tasa_r4:0, tasa_r5:0,
        limite_r1:0, limite_r2:0, limite_r3:0, crm_id: aseg,
      };
      try{
        if(existing && existing._spId){
          await spUpdate('comisiones', existing._spId, data);
          const idx=cache.findIndex(x=>x.crm_id===aseg||x.Title===aseg); if(idx>=0) cache[idx]={...cache[idx],...data};
        } else {
          const id=await spCreate('comisiones',data); if(id){ cache.push({...data,_spId:id}); }
        }
      }catch(eItem){ console.warn('[comisiones] Error sync legacy', aseg, eItem.message); }
    }
    _cache.comisiones = cache;
  }catch(e){ console.warn('[comisiones] _flushComisiones error:', e.message); }
}

// Elimina filas SP de formato antiguo (crm_id no reconocido en V2 ni en sinV2)
async function _cleanupComisionesLegacy(){
  if(!_spReady) return;
  try{
    const v2 = _getTasasV2();
    const asegConV2  = new Set(Object.values(v2).map(r => r.aseg));
    const sinV2Names = new Set(Object.keys(COMISIONES_DEFAULT).filter(a => !asegConV2.has(a)));
    const validIds   = new Set([...Object.keys(TASAS_V2_DEFAULT), ...sinV2Names]);
    const cache      = Array.isArray(_cache.comisiones) ? [..._cache.comisiones] : [];
    let deleted = 0;
    for(const row of cache){
      if(!row._spId) continue;
      const cid = row.crm_id || '';
      if(!cid || !validIds.has(cid)){
        try{ await spDelete('comisiones', row._spId); deleted++; }
        catch(e2){ console.warn('[comisiones] cleanup delete error:', cid, e2.message); }
      }
    }
    if(deleted > 0){
      _cache.comisiones = cache.filter(r => r.crm_id && validIds.has(r.crm_id));
      console.log('[comisiones] Limpieza SP: eliminadas', deleted, 'filas obsoletas');
    }
  }catch(e){ console.warn('[comisiones] _cleanupComisionesLegacy error:', e.message); }
}

// ── Tasas por aseguradora con rangos de suma asegurada ───────────────────────
// Breakpoints: límite SUPERIOR de cada rango (Infinity = sin límite)
const RANGOS_VA    = [10000, 20000, 30000, 50000, Infinity];
const RANGOS_LABEL = ['Hasta $10k', '$10k – $20k', '$20k – $30k', '$30k – $50k', 'Más de $50k'];
// Defaults: [r1, r2, r3, r4, r5] en decimal por aseguradora
// (igual valor en todos los rangos = mismo comportamiento que antes hasta que el Admin los diferencie)
const TASAS_RANGOS_DEFAULT = {
  ZURICH:          [0.043, 0.043, 0.043, 0.043, 0.043],
  LATINA:          [0.038, 0.038, 0.038, 0.038, 0.038],
  GENERALI:        [0.035, 0.035, 0.035, 0.035, 0.035],
  ADS:             [0.045, 0.045, 0.045, 0.045, 0.045],
  SWEADEN:         [0.035, 0.035, 0.035, 0.035, 0.035],
  MAPFRE:          [0.037, 0.037, 0.037, 0.037, 0.037],
  ALIANZA:         [0.032, 0.032, 0.032, 0.032, 0.032],
  'ASEG. DEL SUR': [0.034, 0.034, 0.034, 0.034, 0.034],
  EQUINOCCIAL:     [0.034, 0.034, 0.034, 0.034, 0.034],
  ATLANTIDA:       [0.036, 0.036, 0.036, 0.036, 0.036],
  AIG:             [0.038, 0.038, 0.038, 0.038, 0.038],
};
// ── Tasas V2: por Aseguradora + Región + Tipo + Rangos Dinámicos ─────────────
// Fuente: archivo "TASAS SEG VH MASIVOS". crm_id = clave única en SP.
// tasas[]: decimal (0.035 = 3.5%)  |  limites[]: USD techo de cada rango
// Para N tasas → N-1 límites: va<=limites[0]→tasas[0], …, último→tasas[N-1]
const TASAS_V2_DEFAULT = {
  // ── SWEADEN: diferencia Sierra vs Costa ─────────────────────────────────────
  'SWEADEN_Sierra':   {aseg:'SWEADEN',        region:'SIERRA', tipo:'',           tasas:[0.035,0.028,0.025],       limites:[20000,30000],       comisionPct:15},
  'SWEADEN_Costa':    {aseg:'SWEADEN',        region:'COSTA',  tipo:'',           tasas:[0.035,0.030,0.028],       limites:[20000,30000],       comisionPct:15},
  // ── MAPFRE: diferencia Renovación vs Nuevo ───────────────────────────────────
  'MAPFRE_Renov':     {aseg:'MAPFRE',         region:'',       tipo:'RENOVACION', tasas:[0.035,0.028,0.025],       limites:[20000,30000],       comisionPct:12},
  'MAPFRE_Nuevo':     {aseg:'MAPFRE',         region:'',       tipo:'NUEVO',      tasas:[0.049,0.032,0.025],       limites:[20000,30000],       comisionPct:12},
  // ── Nacionales (sin diferencia regional) ─────────────────────────────────────
  'ALIANZA':          {aseg:'ALIANZA',        region:'',       tipo:'',           tasas:[0.030,0.025,0.022],       limites:[20000,30000],       comisionPct:15},
  'ZURICH':           {aseg:'ZURICH',         region:'',       tipo:'',           tasas:[0.043,0.025,0.022],       limites:[30000,40000],       comisionPct:14},
  'LATINA':           {aseg:'LATINA',         region:'',       tipo:'',           tasas:[0.035,0.028,0.025],       limites:[20000,30000],       comisionPct:15},
  'GENERALI':         {aseg:'GENERALI',       region:'',       tipo:'',           tasas:[0.035,0.035,0.035],       limites:[20000,30000],       comisionPct:12},
  // ── ASEG. DEL SUR (clave "ASEG. DEL SUR" en cotizador — estándar) ────────────
  'ADS_UIO':          {aseg:'ASEG. DEL SUR',  region:'SIERRA', tipo:'',           tasas:[0.045,0.031,0.022,0.020], limites:[20000,30000,40000], comisionPct:13},
  'ADS_GYE':          {aseg:'ASEG. DEL SUR',  region:'COSTA',  tipo:'',           tasas:[0.055,0.034,0.022,0.020], limites:[20000,30000,40000], comisionPct:13},
  // ── ADS (clave "ADS" en cotizador — masivos, cargo adicional $80) ────────────
  'ADS_MAS_UIO':      {aseg:'ADS',            region:'SIERRA', tipo:'',           tasas:[0.045,0.031,0.022,0.020], limites:[20000,30000,40000], comisionPct:13},
  'ADS_MAS_GYE':      {aseg:'ADS',            region:'COSTA',  tipo:'',           tasas:[0.055,0.034,0.022,0.020], limites:[20000,30000,40000], comisionPct:13},
};
// ── Planes de Vida / Asistencia Médica por aseguradora ───────────────────────
// Costo en decimal y coberturas para mostrar en PDF (sin precio)
const PLANES_VIDA = {
  LATINA: {
    'Plan Único': {
      costo: 50,
      coberturas: [
        { concepto:'Vida / Muerte Accidental',        valor:'$20,000' },
        { concepto:'Renta Hospitalización',           valor:'$50/día · máx. 10 días' },
        { concepto:'Telemedicina',                    valor:'E-DOCTOR (Med. General, Psicología, Nutrición)' },
        { concepto:'Beneficio Dental',                valor:'Prevención y Cirugía · 70–100% cobertura' },
        { concepto:'Asegurados',                      valor:'Titular y Cónyuge' },
      ],
    },
  },
  MAPFRE: {
    'Plan Único': {
      costo: 63.19,
      coberturas: [
        { concepto:'Vida / Muerte Accidental',        valor:'$5,000' },
        { concepto:'Enfermedades Graves',             valor:'Anticipo 50% cobertura principal ($2,500)' },
        { concepto:'Renta Hospitalización (acc.)',    valor:'$20/día · máx. 30 días · hasta $600' },
        { concepto:'Gastos de Sepelio',               valor:'$500' },
        { concepto:'Telemedicina / E-Doctor',         valor:'Sin límite' },
        { concepto:'Médico a Domicilio',              valor:'Copago $10/evento · sin límite' },
        { concepto:'Asegurados',                      valor:'Titular y Cónyuge' },
      ],
    },
  },
  ALIANZA: {
    'Plan Único': {
      costo: 55.74,
      coberturas: [
        { concepto:'Vida / Muerte Accidental',        valor:'$5,000' },
        { concepto:'Enfermedades Graves',             valor:'$2,500' },
        { concepto:'Renta Hospitalización',           valor:'$20/día · máx. 25 días · hasta $500' },
        { concepto:'Gastos de Sepelio',               valor:'$500' },
        { concepto:'Telemedicina',                    valor:'Orientación médica telefónica sin límite + E-Doctor' },
        { concepto:'Médico a Domicilio',              valor:'Copago $10/evento · Titular y Cónyuge' },
        { concepto:'Asegurados',                      valor:'Titular y Cónyuge' },
      ],
    },
  },
  SWEADEN: {
    'Plan 1': {
      costo: 59.99,
      coberturas: [
        { concepto:'Vida / Muerte Accidental',        valor:'$5,000' },
        { concepto:'Desmembración / Incapacidad tot.', valor:'$5,000' },
        { concepto:'Enfermedades Graves',             valor:'$2,500' },
        { concepto:'Gastos Médicos por Accidente',    valor:'$250' },
        { concepto:'Ambulancia por Accidente',        valor:'$150' },
        { concepto:'Renta Hospitalización',           valor:'$20/día · máx. 30 días' },
        { concepto:'Gastos de Sepelio',               valor:'$500' },
        { concepto:'Médico a Domicilio + Telemedicina',valor:'Copago $10/evento' },
        { concepto:'Asegurados',                      valor:'Titular y Cónyuge' },
      ],
    },
    'Plan 2': {
      costo: 99.99,
      coberturas: [
        { concepto:'Vida / Muerte Accidental',        valor:'$10,000' },
        { concepto:'Desmembración / Incapacidad tot.', valor:'$10,000' },
        { concepto:'Enfermedades Graves',             valor:'$5,000' },
        { concepto:'Gastos Médicos por Accidente',    valor:'$500' },
        { concepto:'Ambulancia por Accidente',        valor:'$150' },
        { concepto:'Renta Hospitalización',           valor:'$25/día · máx. 30 días' },
        { concepto:'Gastos de Sepelio',               valor:'$800' },
        { concepto:'Médico a Domicilio + Telemedicina',valor:'Copago $10/evento' },
        { concepto:'Asegurados',                      valor:'Grupo familiar (Titular, Cónyuge y 3 hijos ≤18 años)' },
      ],
    },
  },
};
// Devuelve el nombre del plan a partir del costo cobrado
function _getPlanVidaNombre(aseg, costo){
  if(!costo || !PLANES_VIDA[aseg]) return '';
  for(const [nombre, plan] of Object.entries(PLANES_VIDA[aseg])){
    if(Math.abs(plan.costo - costo) < 0.01) return nombre;
  }
  return '';
}

function _getTasasRangos(){
  try{ return Object.assign({...TASAS_RANGOS_DEFAULT}, JSON.parse(localStorage.getItem('_reliance_tasas_rangos')||'{}')); }
  catch(e){ return {...TASAS_RANGOS_DEFAULT}; }
}
function _saveTasasRangos(obj){
  localStorage.setItem('_reliance_tasas_rangos', JSON.stringify(obj));
  _flushComisiones(); // tasas y comisiones viven en la misma lista SP
}
// ── API V2: tasas con región + tipo + límites dinámicos ──────────────────────
function _getTasasV2(){
  try{ return Object.assign({...TASAS_V2_DEFAULT}, JSON.parse(localStorage.getItem('_reliance_tasas_v2')||'{}')); }
  catch(e){ return {...TASAS_V2_DEFAULT}; }
}
function _saveTasasV2(obj){
  localStorage.setItem('_reliance_tasas_v2', JSON.stringify(obj));
  _flushComisiones();
}
// Devuelve la tasa correspondiente al valor asegurado para una aseguradora (legacy)
function _getTasaRango(name, va){
  const rangos = _getTasasRangos()[name];
  if(!rangos || !rangos.length) return ASEGURADORAS[name]?.tasa || 0.035;
  for(let i = 0; i < RANGOS_VA.length; i++){
    if(va <= RANGOS_VA[i]) return rangos[i];
  }
  return rangos[rangos.length - 1];
}
// Devuelve la tasa usando la estructura V2 (región + tipo + límites dinámicos)
// region: 'SIERRA'|'COSTA'  —  tipo: 'NUEVO'|'RENOVACION'
// Sistema de puntuación: región exacta +2, tipo exacto +1, distinto -10
function _getTasaRangoV2(asegName, va, region, tipo){
  const all    = _getTasasV2();
  region       = (region||'').toUpperCase();
  tipo         = (tipo||'').toUpperCase();
  const filas  = Object.values(all).filter(r => r.aseg === asegName);
  if(!filas.length) return _getTasaRango(asegName, va); // fallback legacy
  const score  = r => {
    let s = 0;
    if(region){ if(r.region === region) s += 2; else if(r.region) s -= 10; }
    if(tipo){   if(r.tipo   === tipo)   s += 1; else if(r.tipo)   s -= 10; }
    return s;
  };
  const fila   = [...filas].sort((a,b) => score(b)-score(a))[0];
  const {tasas, limites} = fila;
  for(let i = 0; i < limites.length; i++){
    if(va <= limites[i]) return tasas[i];
  }
  return tasas[tasas.length - 1];
}
// Sanitiza el nombre de aseguradora para usarlo como sufijo de ID en el DOM
function _safeName(n){ return n.replace(/[^a-zA-Z0-9]/g,'_'); }
// Deduce SIERRA/COSTA a partir del nombre de la ciudad (para clientes sin región guardada)
function _ciudadToRegion(ciudad){
  if(!ciudad) return '';
  const c = ciudad.toUpperCase().trim();
  const COSTA  = ['GUAYAQUIL','GYE','SALINAS','PORTOVIEJO','MANTA','MACHALA','ESMERALDAS','QUEVEDO','SANTO DOMINGO','DAULE','MILAGRO','BABAHOYO','LA LIBERTAD','LIBERTAD'];
  const SIERRA = ['QUITO','UIO','CUENCA','AMBATO','RIOBAMBA','LOJA','IBARRA','LATACUNGA','TULCAN','AZOGUES','GUARANDA','SANGOLQUI','CUMBAYA','TUMBACO'];
  if(COSTA.some(x  => c===x || c.startsWith(x) || x.startsWith(c.slice(0,5)))) return 'COSTA';
  if(SIERRA.some(x => c===x || c.startsWith(x) || x.startsWith(c.slice(0,5)))) return 'SIERRA';
  return '';
}
// Asigna marca en #cot-marca; si no está en la lista la inserta dinámicamente
// para que cualquier marca del Excel quede visible (no se pierda en 'OTRO')
function _setMarcaSelect(el, marca){
  if(!el || !marca) return;
  const mu = marca.toUpperCase().trim();
  // 1. Coincidencia exacta (sin importar mayúsculas)
  let opt = [...el.options].find(o => o.value.toUpperCase() === mu);
  // 2. Coincidencia parcial por primeras letras (ej: "VW"→VOLKSWAGEN, "GREAT-WALL"→GREAT WALL)
  if(!opt && mu.length >= 3)
    opt = [...el.options].find(o => {
      const ov = o.value.toUpperCase();
      return ov !== 'OTRO' && (ov.startsWith(mu.slice(0,5)) || mu.startsWith(ov.slice(0,5)));
    });
  if(opt){ el.value = opt.value; return; }
  // 3. Marca no reconocida → agregar como opción dinámica (antes de OTRO)
  if(![...el.options].find(o => o.value === mu)){
    const newOpt = document.createElement('option');
    newOpt.value = mu; newOpt.textContent = mu;
    const otroOpt = [...el.options].find(o => o.value === 'OTRO');
    if(otroOpt) el.insertBefore(newOpt, otroOpt); else el.appendChild(newOpt);
  }
  el.value = mu;
}

// Lee la tasa del input de la tarjeta (si ejecutivo la modificó);
// si no, usa tasas V2 considerando región y tipo de póliza del cotizador
function _getTasaFromCard(name){
  const s = _safeName(name);
  const input = document.getElementById('aseg-tasa-input-'+s);
  if(input && !input.dataset.computed){ const v=parseFloat(input.value); if(!isNaN(v)&&v>0) return v/100; }
  const va     = parseFloat(document.getElementById('cot-va')?.value)||0;
  const ext    = parseFloat(document.getElementById('cot-extras')?.value)||0;
  const region = document.getElementById('cot-region')?.value || '';
  const tipo   = document.getElementById('cot-tipo')?.value   || '';
  return _getTasaRangoV2(name, va + ext, region, tipo);
}

let currentUser = null;
let currentSegIdx = null;
let currentSegEstado = 'PENDIENTE';
let importedRows = [];
let calYear, calMonth;

// ── SP: Carga y guardado de datos ───────────────────────
let DB = [];

function loadDB(){ return _cache.clientes || []; }

async function loadDBAsync(){
  const data = await spGetAll('clientes');
  if(data.length > 0){ DB = data; } else if(!_spReady){ DB = getDefaultClientes(); } else { DB = []; }
  _cache.clientes = DB;
}

function saveDB(){
  localStorage.setItem('reliance_clientes', JSON.stringify(DB)); // fallback
  _cache.clientes = [...DB];
  if(_spReady){
    // Solo marcar dirty los que aún no tienen _spId (nuevos sin sync)
    // Los que ya tienen _spId solo se sincronizan si explícitamente se marcaron _dirty
    DB.forEach(c => { if(!c._spId) c._dirty = true; });
    _flushDB();
  }
}

function _countDirty(){
  let n = 0;
  DB.forEach(c=>{ if(c._dirty) n++; });
  (_getCotizaciones()||[]).forEach(c=>{ if(c._dirty) n++; });
  (_getCierres()||[]).forEach(c=>{ if(c._dirty) n++; });
  (_getTareas()||[]).forEach(t=>{ if(t._dirty) n++; });
  (_getGestionCobranza()||[]).forEach(g=>{ if(g._dirty) n++; });
  return n;
}
function _forceSync(){
  if(!_spReady){ showToast('SharePoint no disponible','error'); return; }
  showToast('⟳ Sincronizando con SharePoint…','info');
  saveDB();
  const allC=_getCotizaciones(); if(allC.some(x=>x._dirty)) _flushCotizaciones(allC);
  const allCi=_getCierres();     if(allCi.some(x=>x._dirty)) _flushCierres(allCi);
  const allG=_getGestionCobranza(); if(allG.some(x=>x._dirty)) _flushGestionCobranza();
}

async function _flushDB(){
  let changed = false;
  for(const cliente of DB){
    if(cliente._dirty){
      if(cliente._spId) await spUpdate('clientes', cliente._spId, cliente);
      else { const id=await spCreate('clientes', cliente); if(id) cliente._spId=id; }
      delete cliente._dirty;
      changed = true;
    }
  }
  if(changed){
    // localStorage ya fue guardado en saveDB(); solo re-renderizar
    const activePage = document.querySelector('.page.active');
    if(activePage && (activePage.id==='page-clientes')) renderClientes();
    renderDashboard();
    updateSpStatus('online','● SharePoint');
  }
}

// Helpers para cotizaciones y cierres (cache + SP + localStorage fallback)
function _getCotizaciones(){
  return _cache.cotizaciones || [];
}
function _saveCotizaciones(all){
  _cache.cotizaciones = all;
  localStorage.setItem('reliance_cotizaciones', JSON.stringify(all));
  if(_spReady) _flushCotizaciones(all);
}
async function _flushCotizaciones(all){
  let changed = false;
  for(const cot of all){
    if(cot._dirty){
      if(cot._spId) await spUpdate('cotizaciones', cot._spId, cot);
      else { const id=await spCreate('cotizaciones', cot); if(id){cot._spId=id;} }
      delete cot._dirty;
      changed = true;
    }
  }
  if(changed){
    // localStorage ya fue guardado en _saveCotizaciones(); solo re-renderizar
    const activePage = document.querySelector('.page.active');
    if(activePage && activePage.id==='page-cotizaciones') renderCotizaciones();
    updateSpStatus('online','● SharePoint');
  }
}
function _getCierres(){
  return _cache.cierres || [];
}
function _saveCierres(all){
  _cache.cierres = all;
  localStorage.setItem('reliance_cierres', JSON.stringify(all));
  if(_spReady) _flushCierres(all);
}
async function _flushCierres(all){
  let changed = false;
  for(const c of all){
    if(c._dirty){
      if(c._spId) await spUpdate('cierres', c._spId, c);
      else { const id=await spCreate('cierres', c); if(id){c._spId=id;} }
      delete c._dirty;
      changed = true;
    }
  }
  if(changed){
    // localStorage ya fue guardado en _saveCierres(); solo re-renderizar
    const activePage = document.querySelector('.page.active');
    if(activePage && activePage.id==='page-cierres') renderCierres();
    renderDashboard();
    updateSpStatus('online','● SharePoint');
  }
}

function loadUsers(){ /* usuarios cargados en initApp desde SP */ }

function saveUsers(){
  localStorage.setItem('reliance_users', JSON.stringify(USERS));
  if(_spReady){
    (async()=>{
      for(const u of USERS){
        try{
          // Buscar en cache de SP por userId o crm_id
          const spRec = (_cache.usuarios||[]).find(r=>
            String(r.userId||r.crm_id||r.id) === String(u.id)
          );
          const spId = u._spId || spRec?._spId;
          if(spId){
            await spUpdate('usuarios', spId, {...u, userId:u.id});
            u._spId = spId; // guardar para futuras actualizaciones
          } else {
            const id = await spCreate('usuarios', {...u, userId:u.id});
            if(id) u._spId = id;
          }
        }catch(e){ console.error('Error guardando usuario:', u.id, e); }
      }
    })();
  }
}

function getDefaultClientes(){
  const today = new Date();
  function addDays(d,n){const r=new Date(d);r.setDate(r.getDate()+n);return r.toISOString().split('T')[0]}
  function subDays(d,n){return addDays(d,-n)}
  return [
    {id:1,ejecutivo:'belen',nombre:'VERA PAZMIÑO LUIS ALEJANDRO',ci:'0930136445',tipo:'RENOVACION',region:'COSTA',ciudad:'GUAYAQUIL',obs:'ENDOSO',celular:'593987304569',aseguradora:'GENERALI ECUADOR',poliza:'ENDOSO',desde:subDays(today.toISOString().split('T')[0],365),hasta:addDays(today.toISOString().split('T')[0],20),va:23530,dep:21200,tasa:null,pn:0,marca:'CHERY',modelo:'TIGGO 7 PRO COMFORT AC 1.5 5P 4X2 TA HYBRID',anio:2024,motor:'SQRE4T15CDBPD00358',chasis:'LVVDB21B6RD008056',color:'NEGRO',placa:'T03072750',polizaAnterior:'ENDOSO',aseguradoraAnterior:'GENERALI ECUADOR',correo:'luis_vera_5@hotmail.com',estado:'PENDIENTE',nota:'',ultimoContacto:''},
    {id:2,ejecutivo:'belen',nombre:'VELEZ TELLO YULISSA PAOLA',ci:'0927086793',tipo:'RENOVACION',region:'COSTA',ciudad:'SALINAS',obs:'RENOVACION',celular:'593980622768',aseguradora:'GENERALI ECUADOR',poliza:'329541-BI204101',desde:subDays(today.toISOString().split('T')[0],340),hasta:addDays(today.toISOString().split('T')[0],25),va:19200,dep:17300,tasa:3.5,pn:605.5,marca:'KIA',modelo:'SPORTAGE GT AC 2.0 5P 4X2 TA',anio:2020,motor:'G4NALH303427',chasis:'U5YPK81ABLL905071',color:'BLANCO',placa:'T02586551',polizaAnterior:'329541-BI204101',aseguradoraAnterior:'GENERALI ECUADOR',correo:'yuli-velez@hotmail.com',estado:'PENDIENTE',nota:'Llamada realizada, espera respuesta',ultimoContacto:subDays(today.toISOString().split('T')[0],2)},
    {id:3,ejecutivo:'belen',nombre:'ESPINOSA SARZOSA FABIAN ENRIQUE',ci:'1706865639',tipo:'RENOVACION',region:'SIERRA',ciudad:'QUITO',obs:'RENOVACION+VD+AXA',celular:'593987428687',aseguradora:'SWEADEN SEGUROS',poliza:'0364883-000954',desde:subDays(today.toISOString().split('T')[0],330),hasta:addDays(today.toISOString().split('T')[0],35),va:21000,dep:18900,tasa:2.84,pn:536.76,marca:'VOLKSWAGEN',modelo:'T-CROSS TRENDLINE AC 1.6 5P 4X2 TM',anio:2023,motor:'CWS156319',chasis:'9BWBL6BFXP4024562',color:'VINO',placa:'T02995681',polizaAnterior:'0364883-000954',aseguradoraAnterior:'SWEADEN SEGUROS',correo:'fabianespinosas@icloud.com',estado:'PENDIENTE',nota:'',ultimoContacto:''},
    {id:4,ejecutivo:'belen',nombre:'MEJIA CHICA ELIANA SOFIA',ci:'1312455577',tipo:'RENOVACION',region:'COSTA',ciudad:'PORTOVIEJO',obs:'ENDOSO',celular:'593992012429',aseguradora:'SWEADEN SEGUROS',poliza:'ENDOSO',desde:subDays(today.toISOString().split('T')[0],360),hasta:addDays(today.toISOString().split('T')[0],5),va:16669,dep:15000,tasa:null,pn:0,marca:'KIA',modelo:'RIO LX AC 1.4 4P 4X2 TM',anio:2023,placa:'T02965444',correo:'sofimejia_sv@hotmail.com',estado:'PENDIENTE',nota:'Interesada en renovar con Alianza',ultimoContacto:subDays(today.toISOString().split('T')[0],1)},
    {id:5,ejecutivo:'belen',nombre:'SUQUILLO VILCA MANUEL RODRIGO',ci:'1706269972',tipo:'RENOVACION',region:'SIERRA',ciudad:'QUITO',obs:'RENOVACION',celular:'593999845214',aseguradora:'LIBERTY SEGUROS',poliza:'94306217',desde:subDays(today.toISOString().split('T')[0],300),hasta:addDays(today.toISOString().split('T')[0],65),va:16800,dep:15100,tasa:2.19,pn:330.69,marca:'FIAT',modelo:'FIAT PULSE AC 1.3 5P 4X2 TM',anio:2023,placa:'E02901786',correo:'patty_lu_eli@hotmail.com',estado:'PENDIENTE',nota:'',ultimoContacto:''},
    {id:6,ejecutivo:'belen',nombre:'AGUIRRE SEVILLA CAYETANO',ci:'1713719662',tipo:'NUEVO',region:'SIERRA',ciudad:'QUITO',obs:'NUEVO',celular:'593984218772',aseguradora:'GENERALI ECUADOR',poliza:'384605-BI191701',desde:subDays(today.toISOString().split('T')[0],200),hasta:addDays(today.toISOString().split('T')[0],165),va:18990,dep:17100,tasa:3.5,pn:598.5,marca:'VOLKSWAGEN',modelo:'POLO TRACK AC 1.6 5P 4X2 TM',anio:2025,placa:'T03240281',correo:'aguirresevilla@gmail.com',estado:'RENOVADO',nota:'Renovado con Generali',ultimoContacto:subDays(today.toISOString().split('T')[0],10)},
    {id:7,ejecutivo:'belen',nombre:'RAMIREZ ANGULO YULEIMA MARIA',ci:'1761758935',tipo:'NUEVO',region:'SIERRA',ciudad:'QUITO',obs:'NUEVO',celular:'593988470289',aseguradora:'MAPFRE SEGUROS',poliza:'8004125000869',desde:subDays(today.toISOString().split('T')[0],250),hasta:addDays(today.toISOString().split('T')[0],115),va:15699,dep:14100,tasa:3.5,pn:493.5,marca:'KIA',modelo:'SOLUTO LX AC 1.4 4P 4X2 TM',anio:2025,placa:'T03286112',correo:'yuleimaramirez04011982@gmail.com',estado:'PENDIENTE',nota:'',ultimoContacto:''},
    {id:8,ejecutivo:'belen',nombre:'DONOSO GARZON HIPATIA',ci:'0502006182',tipo:'RENOVACION',region:'SIERRA',ciudad:'QUITO',obs:'RENOVACION',celular:'593995415318',aseguradora:'GENERALI ECUADOR',poliza:'337468-BI305101',desde:subDays(today.toISOString().split('T')[0],320),hasta:addDays(today.toISOString().split('T')[0],45),va:20200,dep:18200,tasa:3.5,pn:637,marca:'KIA',modelo:'SONET AC 1.5 5P 4X2 TM',anio:2023,placa:'AYM2202377',correo:'hipady@hotmail.com',estado:'PENDIENTE',nota:'',ultimoContacto:''},
    {id:9,ejecutivo:'belen',nombre:'LIGÑA CONZA JEFFERSON HERNAN',ci:'1726513185',tipo:'RENOVACION',region:'SIERRA',ciudad:'QUITO',obs:'RENOVACION+VD',celular:'593998892757',aseguradora:'SWEADEN SEGUROS',poliza:'0364883-000966',desde:subDays(today.toISOString().split('T')[0],310),hasta:addDays(today.toISOString().split('T')[0],55),va:15000,dep:13500,tasa:3.4,pn:459,marca:'CHERY',modelo:'ARRIZO 5 PRO COMFORT AC 1.5 4P 4X2 TM',anio:2024,placa:'G03030468',correo:'jhligna@gmail.com',estado:'PENDIENTE',nota:'',ultimoContacto:''},
    {id:10,ejecutivo:'belen',nombre:'RECALDE AGUAS DAYSI ALEXANDRA',ci:'1719372193',tipo:'RENOVACION',region:'SIERRA',ciudad:'QUITO',obs:'RENOVACION',celular:'593983456789',aseguradora:'ALIANZA SEGUROS',poliza:'PENDIENTE',desde:subDays(today.toISOString().split('T')[0],365),hasta:addDays(today.toISOString().split('T')[0],15),va:8700,dep:7800,tasa:3.0,pn:261,marca:'GREAT WALL',modelo:'VOLEEX C30 CONFORT AC 1.5 4P',anio:2018,placa:'',correo:'',estado:'PENDIENTE',nota:'Ver comparativo Generali/Sweaden/Alianza',ultimoContacto:''},
    {id:11,ejecutivo:'belen',nombre:'BRIONES CEDEÑO ESTHER MARINA',ci:'0930863055',tipo:'RENOVACION',region:'COSTA',ciudad:'GUAYAQUIL',obs:'POLIZA ANULADA',celular:'593983230471',aseguradora:'SEGUROS COLONIAL',poliza:'0366146-001167',desde:subDays(today.toISOString().split('T')[0],365),hasta:addDays(today.toISOString().split('T')[0],-10),va:15800,dep:14200,tasa:null,pn:0,marca:'CHERY',modelo:'TIGGO 2 PRO AC 1.5 5P 4X2 TM',anio:2024,placa:'GTR6865',correo:'sthrbriones@gmail.com',estado:'PÓLIZA ANULADA',nota:'Póliza anulada',ultimoContacto:''},
    {id:12,ejecutivo:'juan',nombre:'VARGAS CARPIO EDIN GREGORIO',ci:'1203241375',tipo:'NUEVO',region:'COSTA',ciudad:'GUAYAQUIL',obs:'NUEVO',celular:'593993474243',aseguradora:'MAPFRE SEGUROS',poliza:'8004125000899',desde:subDays(today.toISOString().split('T')[0],280),hasta:addDays(today.toISOString().split('T')[0],85),va:17539,dep:15800,tasa:3.5,pn:553,marca:'KIA',modelo:'SOLUTO LX AC 1.4 4P 4X2 TM',anio:2025,placa:'T03269319',correo:'edin16261525@gmail.com',estado:'PENDIENTE',nota:'',ultimoContacto:''},
    {id:13,ejecutivo:'juan',nombre:'ORTIZ MARMOL VICTOR MANUEL',ci:'1720193828',tipo:'NUEVO',region:'SIERRA',ciudad:'QUITO',obs:'NUEVO',celular:'593960543586',aseguradora:'MAPFRE SEGUROS',poliza:'8004125000908',desde:subDays(today.toISOString().split('T')[0],260),hasta:addDays(today.toISOString().split('T')[0],105),va:26399,dep:23800,tasa:2.8,pn:666.4,marca:'KIA',modelo:'K3 CROSS LX AC 1.4 5P 4X2 TM',anio:2025,placa:'T03242305',correo:'vmcompare@hotmail.com',estado:'PENDIENTE',nota:'',ultimoContacto:''},
    {id:14,ejecutivo:'maria',nombre:'TORRES DIAZ JOSELYN LISSETH',ci:'1206413856',tipo:'NUEVO',region:'SIERRA',ciudad:'QUITO',obs:'NUEVO',celular:'593988291950',aseguradora:'EQUISUIZA',poliza:'314653',desde:subDays(today.toISOString().split('T')[0],200),hasta:addDays(today.toISOString().split('T')[0],165),va:42990,dep:38700,tasa:2.5,pn:967.5,marca:'NISSAN',modelo:'X-TRAIL EPOWER ADVANCE AC 5P 4X4 TA EV',anio:2024,placa:'T03176173',correo:'JOSELYNDULCE1992@ICLOUD.COM',estado:'PENDIENTE',nota:'',ultimoContacto:''},
    {id:15,ejecutivo:'maria',nombre:'ANGOS GUERRA MIGUEL HERIBERTO',ci:'1712649795',tipo:'NUEVO',region:'SIERRA',ciudad:'QUITO',obs:'NUEVO',celular:'593985967174',aseguradora:'MAPFRE SEGUROS',poliza:'8004125000992',desde:subDays(today.toISOString().split('T')[0],240),hasta:addDays(today.toISOString().split('T')[0],125),va:19990,dep:18000,tasa:3.5,pn:630,marca:'SHINERAY',modelo:'SWM G01 F AC 1.5 5P 4X2 TM',anio:2025,placa:'CIA2402258',correo:'miguelangos1974@hotmail.com',estado:'PENDIENTE',nota:'',ultimoContacto:''},
  ];
}

// ══════════════════════════════════════════════════════
//  AUTH
// ══════════════════════════════════════════════════════
// ── Timeout de sesión por inactividad ──────────────────────────────────────
const SESSION_INACTIVITY_MS = 20 * 60 * 1000; // 20 min sin actividad → advertencia
const SESSION_WARN_SECS     = 60;              // segundos de cuenta regresiva antes de logout

let _sessionTimer    = null;
let _sessionWarnInt  = null;
let _sessionWarnSecs = SESSION_WARN_SECS;

function _resetSessionTimer(){
  clearTimeout(_sessionTimer);
  clearInterval(_sessionWarnInt);
  _hideSessionWarning();
  if(!currentUser) return;
  _sessionTimer = setTimeout(_showSessionWarning, SESSION_INACTIVITY_MS);
}

function _showSessionWarning(){
  const modal = document.getElementById('modal-session-timeout');
  if(!modal || !currentUser) return;
  _sessionWarnSecs = SESSION_WARN_SECS;
  const cdEl = document.getElementById('session-countdown');
  if(cdEl) cdEl.textContent = _sessionWarnSecs;
  modal.style.display = 'flex';
  clearInterval(_sessionWarnInt);
  _sessionWarnInt = setInterval(() => {
    _sessionWarnSecs--;
    if(cdEl) cdEl.textContent = _sessionWarnSecs;
    if(_sessionWarnSecs <= 0){
      clearInterval(_sessionWarnInt);
      _hideSessionWarning();
      doLogout();
      setTimeout(()=>{ const err=document.getElementById('login-err'); if(err) err.textContent='Sesión cerrada por inactividad.'; }, 50);
    }
  }, 1000);
}

function _hideSessionWarning(){
  const modal = document.getElementById('modal-session-timeout');
  if(modal) modal.style.display = 'none';
  clearInterval(_sessionWarnInt);
}

function _keepSession(){
  _hideSessionWarning();
  _resetSessionTimer();
}

const _SESSION_EVENTS = ['mousemove','mousedown','keydown','scroll','touchstart','click'];

function _startSessionTracking(){
  _SESSION_EVENTS.forEach(ev => document.addEventListener(ev, _resetSessionTimer, { passive: true }));
  _resetSessionTimer();
}

function _stopSessionTracking(){
  clearTimeout(_sessionTimer);
  clearInterval(_sessionWarnInt);
  _hideSessionWarning();
  _SESSION_EVENTS.forEach(ev => document.removeEventListener(ev, _resetSessionTimer));
  _sessionTimer = null;
}

async function doLogin(){
  const u = document.getElementById('login-user').value.trim().toLowerCase();
  const p = document.getElementById('login-pass').value;
  document.getElementById('login-err').textContent = '';

  // Fusionar usuarios de SharePoint con USERS locales
  if(_spReady && _cache.usuarios && _cache.usuarios.length){
    _cache.usuarios.forEach(spU => {
      const spId = spU.userId || spU.crm_id || spU.id;
      const local = USERS.find(x => String(x.id)===String(spId) || x.email===spU.email);
      if(local){
        if(spU.rol)      local.rol      = spU.rol;
        if(spU.color)    local.color    = spU.color;
        if(spU.initials) local.initials = spU.initials;
        if(spU.email)    local.email    = spU.email;
      } else if(spId){
        USERS.push({
          id:       spId,
          name:     spU.nombre || spU.Title || spU.email || spId,
          email:    spU.email   || '',
          pass:     spId,
          rol:      spU.rol     || 'ejecutivo',
          color:    spU.color   || '#1a4c84',
          initials: spU.initials|| (spU.email||spId)[0].toUpperCase(),
        });
      }
    });
  }

  const user = USERS.find(x => (x.email.toLowerCase()===u || String(x.id)===String(u)) && x.pass===p);
  if(!user){ document.getElementById('login-err').textContent='Usuario o contraseña incorrectos'; return; }
  currentUser = user;
  document.getElementById('login-screen').style.display = 'none';
  document.getElementById('sidebar-avatar').textContent = user.initials;
  document.getElementById('sidebar-avatar').style.background = `linear-gradient(135deg,${user.color},${user.color}99)`;
  document.getElementById('sidebar-name').textContent = user.name;
  document.getElementById('sidebar-role').textContent = user.rol==='admin'?'Administrador':'Ejecutivo Comercial';
  document.querySelectorAll('.admin-only').forEach(el=>el.style.display=user.rol==='admin'?'':'none');
  _startSessionTracking();
  await initApp();
}
function doLogout(){
  _stopSessionTracking();
  currentUser=null;
  // Limpiar intervals para que no sigan corriendo sin sesión activa
  if(typeof _syncInterval  !== 'undefined' && _syncInterval)  { clearInterval(_syncInterval);  _syncInterval  = null; }
  if(typeof _notifInterval !== 'undefined' && _notifInterval) { clearInterval(_notifInterval); _notifInterval = null; }
  if(typeof _visibilityHandler !== 'undefined' && _visibilityHandler){
    document.removeEventListener('visibilitychange', _visibilityHandler);
    _visibilityHandler = null;
  }
  document.getElementById('login-screen').style.display = 'flex';
  document.getElementById('login-user').value='';
  document.getElementById('login-pass').value='';
  document.getElementById('login-err').textContent='';
}

// ══════════════════════════════════════════════════════
//  UTILS
// ══════════════════════════════════════════════════════
function fmt(n){return '$'+Number(n||0).toLocaleString('es-EC',{minimumFractionDigits:2,maximumFractionDigits:2})}
function myClientes(){
  if(!currentUser) return [];
  return currentUser.rol==='admin' ? DB : DB.filter(c=>String(c.ejecutivo)===String(currentUser.id));
}
function daysUntil(dateStr){
  if(!dateStr) return 9999;
  const today=new Date(); today.setHours(0,0,0,0);
  const d=new Date(dateStr); d.setHours(0,0,0,0);
  return Math.round((d-today)/(1000*60*60*24));
}
function vencClass(days){
  if(days<0) return 'venc-vencida';
  if(days<=30) return 'venc-30';
  if(days<=60) return 'venc-60';
  return 'venc-90';
}
function estadoBadge(e){
  const cfg = ESTADOS_RELIANCE[e];
  if(!cfg) return `<span class="badge badge-gray">${e||'PENDIENTE'}</span>`;
  return `<span class="badge ${cfg.badge}" title="${cfg.def}">${cfg.icon} ${cfg.label}</span>`;
}
function showToast(msg,type='success'){
  const t=document.getElementById('toast-container');
  const d=document.createElement('div');
  d.className=`toast ${type}`;
  d.innerHTML=(type==='success'?'✓':type==='error'?'✕':'ℹ')+' '+msg;
  t.appendChild(d);
  setTimeout(()=>d.remove(),3000);
}
function closeModal(id){
  document.getElementById(id).classList.remove('open');
  if(id==='modal-cierre-venta' && cierreVentaData && cierreVentaData.editandoCierreId){
    cierreVentaData.editandoCierreId=null;
    const btnG=document.querySelector('#modal-cierre-venta .btn-green[onclick="guardarCierreVenta()"]');
    if(btnG){ btnG.textContent='✓ Registrar Cierre'; btnG.style.background=''; }
  }
}
function openModal(id){document.getElementById(id).classList.add('open')}

// ══════════════════════════════════════════════════════
//  NAVIGATION
// ══════════════════════════════════════════════════════
const pageTitles={dashboard:'Dashboard',cierres:'Cierres de Venta',clientes:'Cartera de Clientes',vencimientos:'Vencimientos de Pólizas',calendario:'Calendario de Vencimientos',seguimiento:'Seguimiento de Clientes',cotizador:'Cotizador de Primas',comparativo:'Comparativo de Coberturas',tasas:'Tabla de Tasas',admin:'Panel de Administración','nuevo-cliente':'Registrar Cliente',cobranza:'Módulo de Cobranza',cola:'Cola de Envío'};
function showPage(id){
  // Cerrar cualquier modal abierto al cambiar de módulo
  document.querySelectorAll('.modal-overlay.open').forEach(m=>m.classList.remove('open'));
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  const navItems = document.querySelectorAll('.nav-item');
  navItems.forEach(n=>n.classList.remove('active'));
  document.getElementById('page-'+id).classList.add('active');
  document.getElementById('page-title').textContent=pageTitles[id]||id;
  navItems.forEach(n=>{
    if(n.getAttribute('onclick')&&n.getAttribute('onclick').includes("'"+id+"'")) n.classList.add('active');
  });
  const renders={clientes:renderClientes,vencimientos:()=>showPage('seguimiento'),calendario:()=>{renderCalendario();renderTareasCalendario();},seguimiento:renderSeguimiento,dashboard:renderDashboard,admin:()=>{renderAdmin();showAdminTab('importar',document.querySelector('#admin-tabs .pill'));},comparativo:renderComparativo,cierres:renderCierres,reportes:renderReportes,cotizaciones:renderCotizaciones,cobranza:()=>renderCobranza(_currentFiltroCobranza||'mes'),cotizador:()=>setTimeout(calcCotizacion,100),tasas:renderTasas,cola:initCola};
  if(renders[id]) renders[id]();
}

// ══════════════════════════════════════════════════════
//  DASHBOARD
// ══════════════════════════════════════════════════════
function renderDashboard(){
  const mine = myClientes();
  const venc30 = mine.filter(c=>{ const d=daysUntil(c.hasta); return d>=0&&d<=30; }).length;
  document.getElementById('stat-total').textContent = mine.length;
  document.getElementById('stat-renov').textContent = mine.filter(c=>c.tipo==='RENOVACION').length;
  document.getElementById('stat-nuevo').textContent = mine.filter(c=>c.tipo==='NUEVO').length;
  document.getElementById('stat-venc30').textContent = venc30;
  document.getElementById('stat-mes').textContent = new Date().toLocaleDateString('es-EC',{month:'long',year:'numeric'}).replace(/^\w/,c=>c.toUpperCase());
  document.getElementById('badge-clientes').textContent = mine.length;

  // Vencimientos próximos en dashboard
  const proximos = mine.filter(c=>{ const d=daysUntil(c.hasta); return d<=30&&d>=-10; })
    .sort((a,b)=>new Date(a.hasta)-new Date(b.hasta)).slice(0,5);
  document.getElementById('dash-vencimientos').innerHTML = proximos.length
    ? proximos.map(c=>{ const d=daysUntil(c.hasta); return `
      <div class="venc-alert ${vencClass(d)}" style="padding:10px 12px;margin-bottom:6px">
        <div class="venc-days" style="font-size:16px;min-width:36px">${d<0?'Vencida':d+'d'}</div>
        <div class="venc-info">
          <div class="venc-name" style="font-size:12px">${c.nombre.split(' ').slice(0,2).join(' ')}</div>
          <div class="venc-meta">${c.aseguradora} · ${c.hasta}</div>
        </div>
        ${estadoBadge(c.estado)}
      </div>` }).join('')
    : '<div class="text-muted" style="font-size:12px;padding:10px">No hay vencimientos críticos.</div>';

  // Seguimiento stats
  const estados = Object.keys(ESTADOS_RELIANCE);
  const colors = ['var(--muted)','var(--accent2)','var(--green)','var(--red)'];
  const emojis = ['⚪','🔵','🟢','🔴'];
  document.getElementById('dash-seguimiento').innerHTML = estados.map((e,i)=>{
    const cnt = mine.filter(c=>(c.estado||'PENDIENTE')===e).length;
    const pct = mine.length ? (cnt/mine.length*100).toFixed(0) : 0;
    return `<div style="margin-bottom:12px">
      <div style="display:flex;justify-content:space-between;margin-bottom:4px">
        <span style="font-size:12px">${emojis[i]} ${e}</span>
        <span class="mono" style="font-size:12px;color:var(--muted)">${cnt} (${pct}%)</span>
      </div>
      <div class="progress-wrap"><div class="progress-bar" style="width:${pct}%;background:${colors[i]}"></div></div>
    </div>`;
  }).join('');

  // Aseg dist
  const dist={};
  mine.forEach(c=>{ dist[c.aseguradora]=(dist[c.aseguradora]||0)+1; });
  const sorted=Object.entries(dist).sort((a,b)=>b[1]-a[1]);
  const clrs=['var(--accent)','var(--accent2)','var(--green)','var(--gold)','#6b5b95','#4a4a4a','var(--muted)'];
  document.getElementById('aseg-dist-table').innerHTML = sorted.map(([aseg,cnt],i)=>{
    const pct=(cnt/mine.length*100).toFixed(0);
    return `<tr>
      <td style="padding:7px 10px;font-size:12px">${aseg}</td>
      <td style="padding:7px 10px;font-size:12px;text-align:right;font-family:'DM Mono',monospace;font-weight:600">${cnt}</td>
      <td style="padding:7px 10px"><div style="display:flex;align-items:center;gap:6px">
        <div class="progress-wrap" style="flex:1"><div class="progress-bar" style="width:${pct}%;background:${clrs[i%clrs.length]}"></div></div>
        <span style="font-size:10px;color:var(--muted);width:28px;font-family:'DM Mono',monospace">${pct}%</span>
      </div></td>
    </tr>`;
  }).join('');

  // Activity
  const recent = mine.slice(0,6);
  document.getElementById('activity-feed').innerHTML = recent.map(c=>`
    <div class="tl-item"><div class="tl-dot"></div>
    <div class="tl-date">${c.desde||'—'}</div>
    <div class="tl-text"><b>${c.nombre.split(' ').slice(0,2).join(' ')}</b> — ${c.obs} <span class="badge ${c.tipo==='NUEVO'?'badge-blue':'badge-gold'}" style="font-size:10px">${c.tipo}</span></div>
    </div>`).join('');

  // ── KPI por ejecutiva (solo admin) ──────────────────────────────────────
  const execPanel = document.getElementById('dash-exec-panel');
  if(execPanel){
    if(currentUser?.rol === 'admin'){
      const ejecutivas = USERS.filter(u => u.rol === 'ejecutivo');
      if(ejecutivas.length){
        const rows = ejecutivas.map(u => {
          const cli = DB.filter(c => String(c.ejecutivo) === String(u.id));
          const renov = cli.filter(c => c.tipo === 'RENOVACION').length;
          const nuevos = cli.filter(c => c.tipo === 'NUEVO').length;
          const v30 = cli.filter(c => { const d = daysUntil(c.hasta); return d >= 0 && d <= 30; }).length;
          const renovados = cli.filter(c => ['RENOVADO','EMITIDO','EMISIÓN','PÓLIZA VIGENTE'].includes(c.estado)).length;
          const pct = cli.length ? Math.round(renovados / cli.length * 100) : 0;
          const alertColor = v30 > 10 ? 'var(--accent)' : v30 > 5 ? 'var(--gold)' : 'var(--green)';
          return `
            <div style="background:#fff;border:1px solid var(--border);border-radius:12px;padding:16px 18px;cursor:pointer;transition:box-shadow .15s"
                 onmouseover="this.style.boxShadow='0 4px 16px rgba(0,0,0,.1)'" onmouseout="this.style.boxShadow=''"
                 onclick="showPage('clientes')">
              <div style="display:flex;align-items:center;gap:10px;margin-bottom:12px">
                <div style="width:36px;height:36px;border-radius:50%;background:linear-gradient(135deg,${u.color},${u.color}99);display:flex;align-items:center;justify-content:center;color:#fff;font-weight:700;font-size:13px;flex-shrink:0">${u.initials}</div>
                <div style="min-width:0">
                  <div style="font-weight:700;font-size:13px;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${u.name}</div>
                  <div style="font-size:10px;color:var(--muted)">${cli.length} cliente${cli.length!==1?'s':''}</div>
                </div>
              </div>
              <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;margin-bottom:12px">
                <div style="background:#f8f9fa;border-radius:7px;padding:7px 10px;text-align:center">
                  <div style="font-size:18px;font-weight:800;color:#1a4c84">${renov}</div>
                  <div style="font-size:10px;color:var(--muted)">Renovaciones</div>
                </div>
                <div style="background:#f8f9fa;border-radius:7px;padding:7px 10px;text-align:center">
                  <div style="font-size:18px;font-weight:800;color:var(--accent2)">${nuevos}</div>
                  <div style="font-size:10px;color:var(--muted)">Nuevos</div>
                </div>
              </div>
              <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:6px">
                <span style="font-size:11px;color:var(--muted)">Progreso renovaciones</span>
                <span style="font-size:11px;font-weight:700;color:${u.color}">${pct}%</span>
              </div>
              <div style="background:#f0f0f0;border-radius:4px;height:6px;overflow:hidden;margin-bottom:10px">
                <div style="height:100%;width:${pct}%;background:linear-gradient(90deg,${u.color},${u.color}99);border-radius:4px;transition:width .4s"></div>
              </div>
              <div style="display:flex;align-items:center;justify-content:space-between">
                <span style="font-size:11px;color:var(--muted)">Vencen en 30d</span>
                <span style="font-size:13px;font-weight:800;color:${alertColor}">${v30}</span>
              </div>
            </div>`;
        }).join('');
        execPanel.innerHTML = `
          <div style="display:flex;align-items:center;justify-content:space-between;margin-bottom:12px">
            <div style="font-weight:700;font-size:14px;color:#1a4c84">👥 Panel de Ejecutivas</div>
            <button class="btn btn-ghost btn-sm" onclick="showPage('reportes')">Ver reporte completo →</button>
          </div>
          <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(210px,1fr));gap:12px">${rows}</div>`;
      } else {
        execPanel.innerHTML = '';
      }
    } else {
      execPanel.innerHTML = '';
    }
  }

  // Adaptar label "Mi Cartera" con nombre de la ejecutiva
  const labelCartera = document.querySelector('.stat-card.s1 .stat-label');
  if(labelCartera && currentUser){
    labelCartera.textContent = currentUser.rol === 'admin' ? 'Total Cartera' : 'Mi Cartera';
  }

  // Actualizar badge de urgencia en sidebar Seguimiento
  const vencBadge = mine.filter(c=>{ const d=daysUntil(c.hasta); return d<0||d<=30; }).length;
  const badgeSeg=document.getElementById('badge-seg-urgente');
  if(badgeSeg){ badgeSeg.textContent=vencBadge||'0'; badgeSeg.style.display=vencBadge>0?'':'none'; }

  // ── Panel de Espera — clientes en EMISIÓN aguardando póliza de aseguradora ──
  const enEmision = mine.filter(c=>c.estado==='EMISIÓN');
  const panelWrap  = document.getElementById('panel-espera-wrap');
  const panelLista = document.getElementById('panel-espera-lista');
  const panelCount = document.getElementById('panel-espera-count');
  if(panelWrap && panelLista){
    if(enEmision.length){
      panelWrap.style.display='';
      if(panelCount) panelCount.textContent=enEmision.length+' pendiente'+(enEmision.length>1?'s':'');
      panelLista.innerHTML=enEmision.map(c=>{
        // Días de espera desde ultimoContacto (fecha en que cambió a EMISIÓN)
        const diasEspera=c.ultimoContacto?Math.floor((Date.now()-new Date(c.ultimoContacto))/(86400000)):null;
        const urgente=typeof diasEspera==='number'&&diasEspera>=7;
        const diasTxt=diasEspera!==null?(diasEspera===0?'hoy':diasEspera+'d'):'-';
        return `<div style="display:flex;align-items:center;gap:10px;padding:9px 0;border-bottom:1px solid var(--border)">
          <div style="flex:1;min-width:0">
            <div style="font-size:13px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${c.nombre.split(' ').slice(0,3).join(' ')}</div>
            <div style="font-size:11px;color:var(--muted)">${c.aseguradora||'—'}</div>
          </div>
          <div style="font-size:11px;font-weight:600;color:${urgente?'var(--red)':'var(--muted)'};white-space:nowrap;min-width:40px;text-align:right">
            📅 ${diasTxt}${urgente?' ⚠️':''}
          </div>
          <button class="btn btn-green btn-xs" style="white-space:nowrap" onclick="abrirCierreDesdeCliente('${c.id}')">📋 Registrar</button>
        </div>`;
      }).join('');
    } else {
      panelWrap.style.display='none';
    }
  }
}

// ══════════════════════════════════════════════════════
//  CLIENTES
// ══════════════════════════════════════════════════════
let clientesFiltrados = [];

// ── Sort state para cartera ──
let _sortCarteraCol = 'dias';
let _sortCarteraDir = 1; // 1=asc, -1=desc

const _CARTERA_COLS = [
  { key:'nombre',    label:'Cliente',     sortFn:(a,b)=>(a.nombre||'').localeCompare(b.nombre||'') },
  { key:'ci',        label:'CI',          sortFn:(a,b)=>(a.ci||'').localeCompare(b.ci||'') },
  { key:'tipo',      label:'Tipo',        sortFn:(a,b)=>(a.tipo||'').localeCompare(b.tipo||'') },
  { key:'aseg',      label:'Aseguradora', sortFn:(a,b)=>(a.aseguradora||'').localeCompare(b.aseguradora||'') },
  { key:'vehiculo',  label:'Vehículo',    sortFn:(a,b)=>(`${a.marca} ${a.modelo}`).localeCompare(`${b.marca} ${b.modelo}`) },
  { key:'placa',     label:'Placa',       sortFn:(a,b)=>(a.placa||'').localeCompare(b.placa||'') },
  { key:'vence',     label:'Vence',       sortFn:(a,b)=>(a.hasta||'').localeCompare(b.hasta||'') },
  { key:'dias',      label:'Días',        sortFn:(a,b)=>daysUntil(a.hasta)-daysUntil(b.hasta) },
  { key:'estado',    label:'Estado',      sortFn:(a,b)=>(a.estado||'').localeCompare(b.estado||'') },
  { key:'_acciones', label:'Acciones',    sortFn:null },
];

function sortCartera(col){
  if(_sortCarteraCol === col){ _sortCarteraDir *= -1; }
  else { _sortCarteraCol = col; _sortCarteraDir = 1; }
  filterClientes();
}

function _renderCarteraThead(){
  const tr = document.getElementById('cartera-thead-row');
  if(!tr) return;
  tr.innerHTML = _CARTERA_COLS.map(col => {
    if(!col.sortFn) return `<th>${col.label}</th>`;
    const activo = _sortCarteraCol === col.key;
    const flecha = activo ? (_sortCarteraDir === 1 ? ' ↑' : ' ↓') : '';
    const style = activo ? 'cursor:pointer;color:var(--primary);user-select:none' : 'cursor:pointer;user-select:none';
    return `<th style="${style}" onclick="sortCartera('${col.key}')" title="Ordenar por ${col.label}">${col.label}<span style="opacity:${activo?1:0.3}">${flecha||' ↕'}</span></th>`;
  }).join('');
}

function renderClientes(){
  initFilters();
  filterClientes();
}
function filterClientes(){
  const _sc=document.getElementById('search-clientes'); if(!_sc) return;
  const q=(_sc.value||'').toLowerCase();
  const tipoEl=document.getElementById('filter-tipo');   const tipo=tipoEl?tipoEl.value:'';
  const asegEl=document.getElementById('filter-aseg');   const aseg=asegEl?asegEl.value:'';
  const regionEl=document.getElementById('filter-region');const region=regionEl?regionEl.value:'';
  const estadoEl=document.getElementById('filter-estado');const estado=estadoEl?estadoEl.value:'';
  const colDef = _CARTERA_COLS.find(c=>c.key===_sortCarteraCol) || _CARTERA_COLS.find(c=>c.key==='dias');
  clientesFiltrados = myClientes().filter(c=>{
    const mq=!q||(c.nombre||'').toLowerCase().includes(q)||(c.ci||'').includes(q)||(c.placa||'').toLowerCase().includes(q)||(c.aseguradora||'').toLowerCase().includes(q);
    const mt=!tipo||c.tipo===tipo;
    const ma=!aseg||c.aseguradora===aseg;
    const mr=!region||c.region===region;
    const me=!estado||(c.estado||'PENDIENTE')===estado;
    return mq&&mt&&ma&&mr&&me;
  }).sort((a,b)=>colDef.sortFn(a,b)*_sortCarteraDir);
  _renderCarteraThead();
  document.getElementById('clientes-count').textContent=clientesFiltrados.length+' clientes';
  const obsColors={RENOVACION:'badge-gold',ENDOSO:'badge-blue',NUEVO:'badge-green','POLIZA ANULADA':'badge-red','NO REGISTRADO':'badge-gray','RENOVACION+VD':'badge-gold','RENOVACION+AXA':'badge-gold','RENOVACION+VD+AXA':'badge-blue'};
  document.getElementById('clientes-tbody').innerHTML = clientesFiltrados.map(c=>{
    const days=daysUntil(c.hasta);
    const daysCls=days<0?'text-accent font-bold':days<=30?'text-accent':days<=60?'':'text-muted';
    const daysText=days<0?`Venc. hace ${Math.abs(days)}d`:days===9999?'—':days+'d';
    return `<tr>
      <td><span style="font-weight:500;font-size:12px">${c.nombre}</span></td>
      <td><span class="mono" style="font-size:11px">${c.ci}</span></td>
      <td><span class="badge ${c.tipo==='NUEVO'?'badge-blue':'badge-gold'}">${c.tipo}</span></td>
      <td style="font-size:12px">${c.aseguradora}</td>
      <td style="font-size:11px">${c.marca} ${(c.modelo||'').substring(0,20)}${(c.modelo||'').length>20?'…':''}</td>
      <td><span class="mono" style="font-size:11px">${c.placa||'—'}</span></td>
      <td><span class="mono" style="font-size:11px">${c.hasta||'—'}</span></td>
      <td class="${daysCls}"><span class="mono" style="font-size:11px;font-weight:600">${daysText}</span></td>
      <td>${estadoBadge(c.estado||'PENDIENTE')}</td>
      <td><div style="display:flex;gap:4px">
        <button class="btn btn-ghost btn-xs" onclick="showClienteModal('${c.id}')">👁</button>
        <button class="btn btn-ghost btn-xs" onclick="openEditar('${c.id}')">✏</button>
        <button class="btn btn-ghost btn-xs" onclick="openSeguimiento('${c.id}')">📞</button>
        <button class="btn btn-ghost btn-xs" onclick="prefillCotizador_show('${c.id}')" title="Cotizar">🧮</button>
        <button class="btn btn-xs" style="background:#25D366;color:#fff" onclick="openWhatsApp('${c.id}','vencimiento')" title="WhatsApp">💬</button>
        <button class="btn btn-xs" style="background:#0078d4;color:#fff" onclick="openEmail('${c.id}','vencimiento')" title="Email">✉️</button>
        <button class="btn btn-ghost btn-xs" onclick="nuevaTareaDesdeCliente('${c.id}')" title="Nueva tarea">📌</button>
        ${(['EMITIDO','EMISIÓN'].includes(c.estado)&&!c.factura)?`<button class="btn btn-green btn-xs" onclick="abrirCierreDesdeCliente('${c.id}')" title="Registrar cierre de venta">📋</button>`:''}
        ${c.factura?`<span title="Cierre registrado: ${c.factura}" style="font-size:14px;cursor:default">✅</span>`:''}
      </div></td>
    </tr>`;
  }).join('') || `<tr><td colspan="10"><div class="empty-state"><div class="empty-icon">🔍</div><p>No se encontraron clientes</p></div></td></tr>`;
}
function initFilters(){
  const s=document.getElementById('filter-aseg');
  if(s.options.length>1) return;
  [...new Set(myClientes().map(c=>c.aseguradora))].sort().forEach(a=>{
    const o=document.createElement('option'); o.value=a; o.textContent=a; s.appendChild(o);
  });
}

// ══════════════════════════════════════════════════════
//  DETALLE MODAL
// ══════════════════════════════════════════════════════
function showClienteModal(id){
  const c=DB.find(x=>String(x.id)===String(id)); if(!c) return;
  document.getElementById('modal-cliente-title').textContent=c.nombre;
  const exec=USERS.find(u=>u.id===c.ejecutivo);
  document.getElementById('modal-cliente-body').innerHTML=`
    <div class="detail-grid">
      <div class="card"><div class="card-body">
        <div class="detail-section">
          <div class="detail-section-title">Datos Personales</div>
          <div class="detail-row"><span class="detail-key">CI</span><span class="detail-val mono">${c.ci||'—'}</span></div>
          <div class="detail-row"><span class="detail-key">Celular</span><span class="detail-val mono">${c.celular||'—'}</span></div>
          <div class="detail-row"><span class="detail-key">Correo</span><span class="detail-val" style="font-size:11px">${c.correo||'—'}</span></div>
          <div class="detail-row"><span class="detail-key">Ciudad</span><span class="detail-val">${c.ciudad||'—'}</span></div>
          ${c.direccionOfi?`<div class="detail-row"><span class="detail-key">Dir. Oficina</span><span class="detail-val" style="font-size:11px">${c.direccionOfi}</span></div>`:''}
          <div class="detail-row"><span class="detail-key">Ejecutivo</span><span class="detail-val">${exec?exec.name:'—'}</span></div>
        </div>
        <div class="detail-section">
          <div class="detail-section-title">Vehículo</div>
          <div class="detail-row"><span class="detail-key">Marca/Año</span><span class="detail-val">${c.marca} ${c.anio}</span></div>
          <div class="detail-row"><span class="detail-key">Modelo</span><span class="detail-val" style="font-size:11px">${c.modelo||'—'}</span></div>
          <div class="detail-row"><span class="detail-key">Placa</span><span class="detail-val mono">${c.placa||'—'}</span></div>
          <div class="detail-row"><span class="detail-key">Color</span><span class="detail-val">${c.color||'—'}</span></div>
          <div class="detail-row"><span class="detail-key">N° Motor</span><span class="detail-val mono" style="font-size:10px">${c.motor||'—'}</span></div>
          <div class="detail-row"><span class="detail-key">N° Chasis</span><span class="detail-val mono" style="font-size:10px">${c.chasis||'—'}</span></div>
          <div class="detail-row"><span class="detail-key">Color</span><span class="detail-val" style="font-size:11px">${c.color||'—'}</span></div>
          <div class="detail-row"><span class="detail-key">Póliza anterior</span><span class="detail-val mono" style="font-size:10px">${c.polizaAnterior||c.poliza||'—'}</span></div>
          <div class="detail-row"><span class="detail-key">Aseg. anterior</span><span class="detail-val" style="font-size:11px">${c.aseguradoraAnterior||'—'}</span></div>
        </div>
          <div class="detail-row"><span class="detail-key">Val. Asegurado</span><span class="detail-val text-accent font-bold">${fmt(c.va)}</span></div>
        </div>
        <div class="detail-section">
          <div class="detail-section-title">Seguimiento</div>
          <div style="margin-bottom:8px">${estadoBadge(c.estado||'PENDIENTE')}</div>
          ${c.nota?`<div style="font-size:12px;color:var(--muted);padding:8px;background:var(--warm);border-radius:6px">💬 ${c.nota}</div>`:''}
          ${c.ultimoContacto?`<div class="detail-row" style="margin-top:6px"><span class="detail-key">Últ. contacto</span><span class="detail-val mono">${c.ultimoContacto}</span></div>`:''}
        </div>
      </div></div>
      <div class="card"><div class="card-body">
        <div class="detail-section">
          <div class="detail-section-title">Póliza</div>
          <div class="detail-row"><span class="detail-key">Tipo</span><span class="detail-val"><span class="badge ${c.tipo==='NUEVO'?'badge-blue':'badge-gold'}">${c.tipo}</span></span></div>
          ${c.ramo?`<div class="detail-row"><span class="detail-key">Ramo</span><span class="detail-val">${c.ramo}</span></div>`:''}
          <div class="detail-row"><span class="detail-key">Aseguradora</span><span class="detail-val font-bold">${c.aseguradora||'—'}</span></div>
          <div class="detail-row"><span class="detail-key">Nº Póliza</span><span class="detail-val mono">${c.poliza||'—'}</span></div>
          <div class="detail-row"><span class="detail-key">OBS</span><span class="detail-val">${c.obs||'—'}</span></div>
          <div class="detail-row"><span class="detail-key">Vigencia Desde</span><span class="detail-val mono">${c.desde||'—'}</span></div>
          <div class="detail-row"><span class="detail-key">Vigencia Hasta</span><span class="detail-val mono">${c.hasta||'—'}</span></div>
          ${c.tasa?`<div class="detail-row"><span class="detail-key">Tasa</span><span class="detail-val mono">${c.tasa}%</span></div>`:''}
          ${c.tasaAnterior?`<div class="detail-row"><span class="detail-key">Tasa ant.</span><span class="detail-val mono">${c.tasaAnterior}%</span></div>`:''}
          ${c.tasaRenov?`<div class="detail-row"><span class="detail-key">Tasa renov. actual</span><span class="detail-val mono">${c.tasaRenov}%</span></div>`:''}
          ${c.pn?`<div class="detail-row"><span class="detail-key">Prima Neta</span><span class="detail-val mono">${fmt(c.pn)}</span></div>`:''}
        </div>
        <div class="detail-section">
          <div class="detail-section-title">Días para Vencimiento</div>
          ${(()=>{ const d=daysUntil(c.hasta);
            const cls=d<0?'badge-red':d<=30?'badge-red':d<=60?'badge-orange':'badge-green';
            const txt=d<0?`Venció hace ${Math.abs(d)} días`:d===9999?'Sin fecha':`${d} días restantes`;
            return `<span class="badge ${cls}" style="font-size:14px;padding:6px 14px">${txt}</span>`;
          })()}
        </div>
        ${c.comentario?`<div style="padding:10px 14px;background:var(--warm);border-radius:8px;font-size:12px;color:var(--muted)">💬 ${c.comentario}</div>`:''}
        ${(c.garantia||c.cuentaBanc||c.prestamo||c.saldo||c.monto||c.estadoCredito||c.fechaDesembolso||c.fechaVtoCred)?`
        <div class="detail-section" style="margin-top:10px">
          <div class="detail-section-title">Crédito Produbanco</div>
          ${c.garantia?`<div class="detail-row"><span class="detail-key">Garantía</span><span class="detail-val mono" style="font-size:10px">${c.garantia}</span></div>`:''}
          ${c.cuentaBanc?`<div class="detail-row"><span class="detail-key">Cuenta</span><span class="detail-val mono">${c.cuentaBanc}</span></div>`:''}
          ${c.prestamo?`<div class="detail-row"><span class="detail-key">N° Préstamo</span><span class="detail-val mono" style="font-size:10px">${c.prestamo}</span></div>`:''}
          ${c.monto?`<div class="detail-row"><span class="detail-key">Monto</span><span class="detail-val">${fmt(c.monto)}</span></div>`:''}
          ${c.saldo?`<div class="detail-row"><span class="detail-key">Saldo</span><span class="detail-val font-bold">${fmt(c.saldo)}</span></div>`:''}
          ${c.estadoCredito?`<div class="detail-row"><span class="detail-key">Estado</span><span class="detail-val">${c.estadoCredito}</span></div>`:''}
          ${c.fechaDesembolso?`<div class="detail-row"><span class="detail-key">Desembolso</span><span class="detail-val mono">${c.fechaDesembolso}</span></div>`:''}
          ${c.fechaVtoCred?`<div class="detail-row"><span class="detail-key">Vto. Crédito</span><span class="detail-val mono">${c.fechaVtoCred}</span></div>`:''}
        </div>`:''}
      </div></div>
    </div>
    <div class="card" style="margin-top:12px">
      <div class="card-header">
        <div class="card-title">🕐 Historial de Actividad</div>
        <span style="font-size:11px;color:var(--muted)" id="modal-timeline-count"></span>
      </div>
      <div class="card-body" style="padding:12px 16px;max-height:420px;overflow-y:auto" id="modal-cliente-timeline"></div>
    </div>
    `;
  document.getElementById('modal-btn-cotizar').style.display='';
  document.getElementById('modal-btn-editar').style.display='';
  document.getElementById('modal-btn-eliminar').style.display='';
  document.getElementById('modal-btn-cotizar').onclick=()=>{ closeModal('modal-cliente'); prefillCotizador(c); showPage('cotizador'); setTimeout(calcCotizacion,200); };
  document.getElementById('modal-btn-editar').onclick=()=>{ closeModal('modal-cliente'); openEditar(id); };
  document.getElementById('modal-btn-eliminar').onclick=()=>{ if(confirm(`¿Eliminar a ${c.nombre}?`)){ const cliToDel=DB.find(x=>String(x.id)===String(id)); if(cliToDel?._spId && _spReady) spDelete('clientes', cliToDel._spId); DB=DB.filter(x=>String(x.id)!==String(id)); saveDB(); closeModal('modal-cliente'); renderClientes(); renderDashboard(); showToast('Cliente eliminado','error'); }};

  // Timeline unificado
  const tlEl = document.getElementById('modal-cliente-timeline');
  if(tlEl){
    tlEl.innerHTML = _renderClienteTimeline(c);
    const tlEntries = _buildClienteTimeline(c);
    const cntEl = document.getElementById('modal-timeline-count');
    if(cntEl) cntEl.textContent = tlEntries.length ? `${tlEntries.length} evento${tlEntries.length!==1?'s':''}` : '';
  }

  openModal('modal-cliente');
}

// ══════════════════════════════════════════════════════
//  EDITAR CLIENTE
// ══════════════════════════════════════════════════════
function openEditar(id){
  const c=DB.find(x=>String(x.id)===String(id)); if(!c) return;
  document.getElementById('modal-editar-body').innerHTML=`
    <div class="form-grid form-grid-2" style="gap:12px">
      <div class="form-group full">
        <label class="form-label">Tipo de Cliente</label>
        <select class="form-select" id="ed-tipo-cliente">
          <option value=""${!c.tipoCliente?' selected':''}>— Sin especificar —</option>
          <option value="PRODUBANCO"${c.tipoCliente==='PRODUBANCO'?' selected':''}>🏦 Produbanco</option>
          <option value="PARTICULAR"${c.tipoCliente==='PARTICULAR'?' selected':''}>👤 Particular</option>
          <option value="NUEVO"${c.tipoCliente==='NUEVO'?' selected':''}>✨ Nuevo Seguro</option>
        </select>
      </div>
      <div class="form-group full"><label class="form-label">Nombre</label><input class="form-input" id="ed-nombre" value="${c.nombre||''}"></div>
      <div class="form-group"><label class="form-label">CI</label><input class="form-input" id="ed-ci" value="${c.ci||''}"></div>
      <div class="form-group"><label class="form-label">Celular Principal</label><input class="form-input" id="ed-cel" value="${c.celular||''}"></div>
      <div class="form-group"><label class="form-label">Celular 2</label><input class="form-input" id="ed-cel2" value="${c.celular2||''}"></div>
      <div class="form-group"><label class="form-label">Teléfono Fijo</label><input class="form-input" id="ed-tel-fijo" value="${c.telFijo||''}"></div>
      <div class="form-group"><label class="form-label">Correo</label><input class="form-input" id="ed-email" value="${c.correo||''}"></div>
      <div class="form-group"><label class="form-label">Ciudad</label><input class="form-input" id="ed-ciudad" value="${c.ciudad||''}"></div>
      <div class="form-group"><label class="form-label">Nacimiento</label><input class="form-input" type="date" id="ed-nacimiento" value="${c.fechaNac||''}"></div>
      <div class="form-group"><label class="form-label">Género</label><select class="form-select" id="ed-genero"><option value=""${!c.genero?' selected':''}>—</option><option${c.genero==='MASCULINO'?' selected':''}>MASCULINO</option><option${c.genero==='FEMENINO'?' selected':''}>FEMENINO</option></select></div>
      <div class="form-group"><label class="form-label">Estado Civil</label><select class="form-select" id="ed-civil"><option value=""${!c.estadoCivil?' selected':''}>—</option><option${c.estadoCivil==='SOLTERO'?' selected':''}>SOLTERO</option><option${c.estadoCivil==='CASADO'?' selected':''}>CASADO</option><option${c.estadoCivil==='DIVORCIADO'?' selected':''}>DIVORCIADO</option><option${c.estadoCivil==='VIUDO'?' selected':''}>VIUDO</option></select></div>
      <div class="form-group"><label class="form-label">Profesión</label><input class="form-input" id="ed-profesion" value="${c.profesion||''}"></div>
      <div class="form-group"><label class="form-label">Aseguradora</label><input class="form-input" id="ed-aseg" value="${c.aseguradora||''}"></div>
      <div class="form-group"><label class="form-label">Póliza</label><input class="form-input" id="ed-poliza" value="${c.poliza||''}"></div>
      <div class="form-group"><label class="form-label">Vigencia Hasta</label><input class="form-input" type="date" id="ed-hasta" value="${c.hasta||''}"></div>
      <div class="form-group"><label class="form-label">Val. Asegurado ($)</label><input class="form-input" type="number" id="ed-va" value="${c.va||''}"></div>
      <div class="form-group"><label class="form-label">Tasa (%)</label><input class="form-input" type="number" step="0.01" id="ed-tasa" value="${c.tasa||''}"></div>
      <div class="form-group"><label class="form-label">Marca</label><input class="form-input" id="ed-marca" value="${c.marca||''}"></div>
      <div class="form-group"><label class="form-label">Placa</label><input class="form-input" id="ed-placa" value="${c.placa||''}"></div>
      <div class="form-group"><label class="form-label">Cuenta Produbanco</label><input class="form-input" id="ed-cuenta" value="${c.cuentaBanc||c.cuenta||''}"></div>
      <div class="form-group"><label class="form-label">N° Préstamo</label><input class="form-input" id="ed-prestamo" value="${c.prestamo||''}"></div>
      <div class="form-group"><label class="form-label">Saldo Crédito ($)</label><input class="form-input" type="number" id="ed-saldo" value="${c.saldo||''}"></div>
      <div class="form-group"><label class="form-label">Vto. Crédito</label><input class="form-input" type="date" id="ed-vto-cred" value="${c.fechaVtoCred||''}"></div>
      ${currentUser&&currentUser.rol==='admin'?`<div class="form-group"><label class="form-label">Asignar a Ejecutivo</label><select class="form-select" id="ed-ejecutivo">${USERS.filter(u=>u.rol==='ejecutivo').map(u=>`<option value="${u.id}"${c.ejecutivo===u.id?' selected':''}>${u.name}</option>`).join('')}</select></div>`:''}
      <div class="form-group full"><label class="form-label">Comentarios</label><textarea class="form-textarea" id="ed-comentario">${c.comentario||''}</textarea></div>
    </div>`;
  document.getElementById('modal-btn-guardar-editar').onclick=async()=>{
    c.nombre=document.getElementById('ed-nombre').value;
    c.ci=document.getElementById('ed-ci').value;
    c.celular=document.getElementById('ed-cel').value;
    c.celular2=document.getElementById('ed-cel2')?.value||c.celular2||'';
    c.telFijo=document.getElementById('ed-tel-fijo')?.value||c.telFijo||'';
    c.correo=document.getElementById('ed-email').value;
    c.ciudad=document.getElementById('ed-ciudad').value;
    c.tipoCliente=document.getElementById('ed-tipo-cliente')?.value||c.tipoCliente||'';
    c.fechaNac=document.getElementById('ed-nacimiento')?.value||c.fechaNac||'';
    c.genero=document.getElementById('ed-genero')?.value||c.genero||'';
    c.estadoCivil=document.getElementById('ed-civil')?.value||c.estadoCivil||'';
    c.profesion=document.getElementById('ed-profesion')?.value||c.profesion||'';
    c.aseguradora=document.getElementById('ed-aseg').value;
    c.poliza=document.getElementById('ed-poliza').value;
    c.hasta=document.getElementById('ed-hasta').value;
    c.va=parseFloat(document.getElementById('ed-va').value)||0;
    c.tasa=parseFloat(document.getElementById('ed-tasa').value)||null;
    c.marca=document.getElementById('ed-marca').value;
    c.placa=document.getElementById('ed-placa').value;
    c.cuentaBanc=document.getElementById('ed-cuenta')?.value||c.cuentaBanc||'';
    c.prestamo=document.getElementById('ed-prestamo')?.value||c.prestamo||'';
    c.saldo=parseFloat(document.getElementById('ed-saldo')?.value)||c.saldo||0;
    c.fechaVtoCred=document.getElementById('ed-vto-cred')?.value||c.fechaVtoCred||'';
    c.comentario=document.getElementById('ed-comentario').value;
    if(document.getElementById('ed-ejecutivo')) c.ejecutivo=document.getElementById('ed-ejecutivo').value;
    await sincronizarCotizPorCliente(c.id, c.nombre, c.ci, c.estado);
    c._dirty = true;
    saveDB(); closeModal('modal-editar'); renderClientes(); renderDashboard(); showToast('Cliente actualizado');
  };
  openModal('modal-editar');
}

// ══════════════════════════════════════════════════════
//  VENCIMIENTOS
// ══════════════════════════════════════════════════════
let vencFilter='all';
// renderVencimientos: ahora delega en _updateVencStats (las stat-cards viven en Seguimiento)
function renderVencimientos(){ _updateVencStats(); }
function filterVencimientos(f,el){
  vencFilter=f;
  document.querySelectorAll('#venc-pills .pill').forEach(p=>p.classList.remove('active'));
  el.classList.add('active');
  renderVencimientos();
}

// ══════════════════════════════════════════════════════
//  CALENDARIO
// ══════════════════════════════════════════════════════
function renderCalendario(){
  const today=new Date();
  if(!calYear) calYear=today.getFullYear();
  if(!calMonth) calMonth=today.getMonth();
  const monthNames=['Enero','Febrero','Marzo','Abril','Mayo','Junio','Julio','Agosto','Septiembre','Octubre','Noviembre','Diciembre'];
  document.getElementById('cal-month-title').textContent=`${monthNames[calMonth]} ${calYear}`;
  const days=['Dom','Lun','Mar','Mié','Jue','Vie','Sáb'];
  document.getElementById('cal-headers').innerHTML=days.map(d=>`<div class="cal-day-header">${d}</div>`).join('');
  const first=new Date(calYear,calMonth,1);
  const last=new Date(calYear,calMonth+1,0);
  const mine=myClientes();
  // Build event map
  const evMap={};
  mine.forEach(c=>{
    if(!c.hasta) return;
    const d=new Date(c.hasta); const key=d.toISOString().split('T')[0];
    if(!evMap[key]) evMap[key]=[];
    evMap[key].push(c);
  });
  // Mapa de tareas por fecha
  const tareaMap={};
  myTareas().filter(t=>t.estado==='pendiente').forEach(t=>{
    if(!t.fechaVence) return;
    if(!tareaMap[t.fechaVence]) tareaMap[t.fechaVence]=[];
    tareaMap[t.fechaVence].push(t);
  });

  let cells='';
  const startDay=first.getDay();
  for(let i=0;i<startDay;i++){
    const d=new Date(calYear,calMonth,1-startDay+i);
    cells+=`<div class="cal-day other-month"><div class="cal-day-num">${d.getDate()}</div></div>`;
  }
  for(let d=1;d<=last.getDate();d++){
    const dateStr=`${calYear}-${String(calMonth+1).padStart(2,'0')}-${String(d).padStart(2,'0')}`;
    const isToday=d===today.getDate()&&calMonth===today.getMonth()&&calYear===today.getFullYear();
    const events=evMap[dateStr]||[];
    const tareas=tareaMap[dateStr]||[];
    const evHtml=events.slice(0,2).map(c=>{ const days=daysUntil(c.hasta); const cls=days<0?'evencida':days<=30?'e30':days<=60?'e60':'e90'; return `<div class="cal-event ${cls}">${c.nombre.split(' ')[0]}</div>`; }).join('');
    const tareaHtml=tareas.slice(0,2).map(t=>`<div class="cal-event" style="background:#e8f0fb;border-left:2px solid var(--accent2);color:var(--accent2)">📌 ${t.titulo.substring(0,12)}</div>`).join('');
    const more=(events.length+tareas.length)>4?`<div style="font-size:9px;color:var(--muted)">+${events.length+tareas.length-4} más</div>`:'';
    const hasAny=events.length||tareas.length;
    cells+=`<div class="cal-day${isToday?' today':''}${hasAny?' has-events':''}"
      onclick="${hasAny?`showCalDia('${dateStr}')`:''}"
      title="${events.length?events.length+' vencimiento(s)':''}${tareas.length?' · '+tareas.length+' tarea(s)':''}">
      <div class="cal-day-num">${d}</div>${evHtml}${tareaHtml}${more}
    </div>`;
  }
  const remaining=(7-((startDay+last.getDate())%7))%7;
  for(let i=1;i<=remaining;i++) cells+=`<div class="cal-day other-month"><div class="cal-day-num">${i}</div></div>`;
  document.getElementById('cal-body').innerHTML=cells;
}
function calNav(dir){calMonth+=dir;if(calMonth>11){calMonth=0;calYear++;}if(calMonth<0){calMonth=11;calYear--;}renderCalendario();}
function showCalEvents(dateStr){ showCalDia(dateStr); }
function showCalDia(dateStr){
  const events = myClientes().filter(c=>c.hasta===dateStr);
  const tareas  = myTareas().filter(t=>t.fechaVence===dateStr && t.estado==='pendiente');
  const titulo  = new Date(dateStr+'T00:00:00').toLocaleDateString('es-EC',{weekday:'long',day:'numeric',month:'long'});
  document.getElementById('modal-cal-title').textContent = titulo;
  const TIPO_ICON = { llamada:'📞', email:'✉️', reunion:'🤝', seguimiento:'📋', otro:'📌' };
  let html = '';
  if(tareas.length){
    html += `<div style="font-size:11px;font-weight:700;text-transform:uppercase;color:var(--muted);margin-bottom:8px">📌 Tareas</div>`;
    html += tareas.map(t=>`
      <div style="display:flex;align-items:center;gap:8px;padding:8px;background:var(--warm);border-radius:6px;margin-bottom:6px;cursor:pointer"
        onclick="closeModal('modal-cal');abrirDetalleTarea('${t.id}')">
        <span>${TIPO_ICON[t.tipo]||'📌'}</span>
        <div style="flex:1">
          <div style="font-size:12px;font-weight:600">${t.titulo}</div>
          ${t.horaVence?`<div style="font-size:10px;color:var(--muted)">${t.horaVence}</div>`:''}
        </div>
        <button class="btn btn-ghost btn-xs" onclick="event.stopPropagation();completarTarea('${t.id}')">✅</button>
      </div>`).join('');
  }
  if(events.length){
    html += `<div style="font-size:11px;font-weight:700;text-transform:uppercase;color:var(--muted);margin:${tareas.length?'12px':0} 0 8px">⏰ Vencimientos</div>`;
    html += events.map(c=>{
      const days=daysUntil(c.hasta);
      return `<div class="venc-alert ${vencClass(days)}" style="margin-bottom:8px;cursor:pointer"
        onclick="closeModal('modal-cal');openSeguimiento('${c.id}')">
        <div class="venc-days" style="font-size:16px;min-width:36px">${days<0?'VENC':days+'d'}</div>
        <div class="venc-info">
          <div class="venc-name">${c.nombre}</div>
          <div class="venc-meta">${c.aseguradora} · ${fmt(c.va)}</div>
        </div>
        ${estadoBadge(c.estado||'PENDIENTE')}
      </div>`;
    }).join('');
  }
  if(!html) html='<div style="color:var(--muted);font-size:12px;text-align:center;padding:16px">Sin eventos este día</div>';
  document.getElementById('modal-cal-body').innerHTML = html;
  openModal('modal-cal');
}

// ══════════════════════════════════════════════════════
//  SEGUIMIENTO
// ══════════════════════════════════════════════════════
let segFilterEstado='';
let segVencFilter='all'; // 'all'|'vencida'|'30'|'60'|'90'

// Actualiza las 4 stat-cards de urgencia y el badge del sidebar
function _updateVencStats(){
  const mine=myClientes();
  const venc0 =mine.filter(c=>daysUntil(c.hasta)<0).length;
  const venc30=mine.filter(c=>{const d=daysUntil(c.hasta);return d>=0&&d<=30}).length;
  const venc60=mine.filter(c=>{const d=daysUntil(c.hasta);return d>30&&d<=60}).length;
  const venc90=mine.filter(c=>{const d=daysUntil(c.hasta);return d>60&&d<=90}).length;
  const vals=[venc0,venc30,venc60,venc90];
  ['vstat-0','vstat-30','vstat-60','vstat-90'].forEach((id,i)=>{
    const el=document.getElementById(id); if(el) el.textContent=vals[i];
  });
  // Resaltar tarjeta activa
  ['vencida','30','60','90'].forEach(f=>{
    const card=document.getElementById('svcard-'+f);
    if(card) card.style.outline=segVencFilter===f?'2px solid var(--accent)':'none';
  });
  const urgent=venc0+venc30;
  const badge=document.getElementById('badge-seg-urgente');
  if(badge){ badge.textContent=urgent; badge.style.display=urgent>0?'':'none'; }
}

function renderSeguimiento(){
  _updateVencStats();
  filterSeguimiento();
}

// Toggle filtro de urgencia por vencimiento (click de nuevo para limpiar)
function filterSegVenc(f){
  segVencFilter=(segVencFilter===f)?'all':f;
  _updateVencStats();
  filterSeguimiento();
}

function filterSeguimiento(){
  const _ss=document.getElementById('search-seg'); if(!_ss) return;
  const q=(_ss.value||'').toLowerCase();
  let data=myClientes().filter(c=>{
    const mq=!q||(c.nombre||'').toLowerCase().includes(q);
    const me=!segFilterEstado||(c.estado||'PENDIENTE')===segFilterEstado;
    const days=daysUntil(c.hasta);
    let mv=true;
    if(segVencFilter==='vencida') mv=days<0;
    else if(segVencFilter==='30')  mv=days>=0&&days<=30;
    else if(segVencFilter==='60')  mv=days>30&&days<=60;
    else if(segVencFilter==='90')  mv=days>60&&days<=90;
    return mq&&me&&mv;
  }).sort((a,b)=>daysUntil(a.hasta)-daysUntil(b.hasta));
  document.getElementById('seg-count').textContent=data.length+' clientes';
  document.getElementById('seguimiento-tbody').innerHTML=data.map(c=>{
    const days=daysUntil(c.hasta);
    const daysCls=days<0?'text-accent':days<=30?'text-accent':'text-muted';
    return `<tr>
      <td><span style="font-weight:500;font-size:12px">${c.nombre}</span><br><span class="mono text-muted" style="font-size:10px">${c.ci}</span></td>
      <td style="font-size:12px">${c.aseguradora}</td>
      <td><span class="mono" style="font-size:11px">${c.hasta||'—'}</span></td>
      <td class="${daysCls}"><span class="mono font-bold" style="font-size:11px">${days<0?'Vencida':days+'d'}</span></td>
      <td>${estadoBadge(c.estado||'PENDIENTE')}</td>
      <td style="font-size:11px;max-width:150px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap" title="${c.nota||''}">${c.nota||'—'}</td>
      <td><span class="mono" style="font-size:11px">${c.ultimoContacto||'—'}</span></td>
      <td><div style="display:flex;gap:4px;flex-wrap:wrap">
        <button class="btn btn-blue btn-xs" onclick="openSeguimiento('${c.id}')">📞 Actualizar</button>
        <button class="btn btn-ghost btn-xs" onclick="prefillCotizador_show('${c.id}')" title="Ir al cotizador con datos de este cliente">🧮 Cotizar</button>
        <button class="btn btn-xs" style="background:#25D366;color:#fff" onclick="openWhatsApp('${c.id}','vencimiento')">💬 WA</button>
        <button class="btn btn-xs" style="background:#0078d4;color:#fff" onclick="openEmail('${c.id}','vencimiento')">✉️ Mail</button>
        <button class="btn btn-ghost btn-xs" onclick="nuevaTareaDesdeCliente('${c.id}')" title="Nueva tarea">📌</button>
        ${(['EMITIDO','EMISIÓN'].includes(c.estado)&&!c.factura)?`<button class="btn btn-green btn-xs" onclick="abrirCierreDesdeCliente('${c.id}')">📋 Cierre</button>`:''}
        ${c.factura?`<span class="badge badge-green" style="font-size:10px">✅ Cerrado</span>`:''}
      </div></td>
    </tr>`;
  }).join('')||`<tr><td colspan="8"><div class="empty-state"><div class="empty-icon">📋</div><p>Sin resultados</p></div></td></tr>`;
}
function filterSegEstado(e,el){
  segFilterEstado=e;
  document.querySelectorAll('#page-seguimiento .filter-pills .pill').forEach(p=>p.classList.remove('active'));
  el.classList.add('active');
  filterSeguimiento();
}
function openSeguimiento(id){
  const c=DB.find(x=>String(x.id)===String(id)); if(!c) return;
  currentSegIdx=id; currentSegEstado=c.estado||'PENDIENTE';
  document.getElementById('modal-seg-nombre').textContent=c.nombre;
  document.getElementById('seg-nota').value=''; // limpiar para nueva entrada
  document.querySelectorAll('.estado-btn').forEach(b=>{ b.classList.remove('active'); if(b.classList.contains(currentSegEstado)) b.classList.add('active'); });
  // Mostrar banner cierre si ya está EMITIDO o EMISIÓN (reutiliza la misma lógica de setEstado)
  setEstado(currentSegEstado);
  // Renderizar bitácora existente dentro del modal
  _renderBitacoraModal(c.bitacora||[]);
  openModal('modal-seguimiento');
}
function setEstado(e){
  currentSegEstado=e;
  document.querySelectorAll('.estado-btn').forEach(b=>{
    b.classList.remove('active');
    if(b.getAttribute('onclick')&&b.getAttribute('onclick').includes("'"+e+"'")) b.classList.add('active');
  });
  // Mostrar banner "Registrar Cierre" al seleccionar EMITIDO o EMISIÓN
  const banner  = document.getElementById('seg-cierre-banner');
  const titulo  = document.getElementById('seg-banner-titulo');
  const desc    = document.getElementById('seg-banner-desc');
  if(banner){
    if(e==='EMITIDO'){
      banner.style.background='#d4edda'; banner.style.borderColor='#2d6a4f';
      if(titulo){ titulo.textContent='✓ Póliza emitida — ¿Registrar cierre de venta?'; titulo.style.color='var(--green)'; }
      if(desc)   desc.textContent='Guarda el estado y abre el formulario de cierre (factura, póliza, forma de pago).';
      banner.style.display='block';
    } else if(e==='EMISIÓN'){
      banner.style.background='#fff3e0'; banner.style.borderColor='#e65100';
      if(titulo){ titulo.textContent='📬 ¿Ya recibiste la póliza de la aseguradora?'; titulo.style.color='#e65100'; }
      if(desc)   desc.textContent='Registra el cierre directamente. El sistema marcará EMITIDO y RENOVADO de forma automática.';
      banner.style.display='block';
    } else {
      banner.style.display='none';
    }
  }
  // Solo actualiza la UI — el estado se persiste al presionar "Guardar" o "Registrar Cierre"
}

// ── Bitácora de gestión ──────────────────────────────────────
// Crea una entrada nueva en c.bitacora[]
function _bitacoraAdd(cliente, nota, tipo='manual'){
  if(!cliente.bitacora) cliente.bitacora = [];
  const exec = USERS.find(u=>String(u.id)===String(currentUser?.id));
  const entrada = {
    fecha:     new Date().toISOString().split('T')[0],
    hora:      new Date().toLocaleTimeString('es-EC',{hour:'2-digit',minute:'2-digit'}),
    ejecutivo: exec?.name || currentUser?.name || 'Sistema',
    estado:    cliente.estado || 'PENDIENTE',
    nota:      nota || '',
    tipo,      // 'manual' | 'sistema' | 'cotizacion' | 'cierre'
  };
  cliente.bitacora.unshift(entrada); // más reciente primero
  // Mantener nota principal como la más reciente con texto
  if(nota) cliente.nota = nota;
  // Marcar para sync con SharePoint
  cliente._dirty = true;
}

// ── Timeline unificado del cliente ────────────────────────────────────────
const _TL_CFG = {
  manual:     { icon:'💬', color:'#1a4c84', label:'Nota'        },
  sistema:    { icon:'⚙️', color:'#888',    label:'Sistema'     },
  cotizacion: { icon:'📋', color:'#2196f3', label:'Cotización'  },
  cierre:     { icon:'✅', color:'#28a745', label:'Cierre'      },
  wa:         { icon:'💚', color:'#25d366', label:'WhatsApp'    },
  poliza:     { icon:'📄', color:'#28a745', label:'Póliza'      },
  tarea:      { icon:'📌', color:'#6f42c1', label:'Tarea'       },
};

function _buildClienteTimeline(c){
  const entries = [];

  // 1. Bitácora (fuente principal)
  (c.bitacora||[]).forEach(e=>{
    const cfg = _TL_CFG[e.tipo] || _TL_CFG.manual;
    entries.push({ fecha:e.fecha||'0000-00-00', hora:e.hora||'', tipo:e.tipo||'manual',
      icono:cfg.icon, color:cfg.color, label:cfg.label,
      titulo:e.nota||'(sin nota)', estado:e.estado||'', ejecutivo:e.ejecutivo||'' });
  });

  // 2. historialWa — incluir los que no estén ya cubiertos en bitácora ese día
  (c.historialWa||[]).forEach(h=>{
    const yaEnBit = (c.bitacora||[]).some(e=>e.fecha===h.fecha &&
      (e.nota||'').toLowerCase().includes('whatsapp'));
    if(!yaEnBit){
      entries.push({ fecha:h.fecha||'0000-00-00', hora:'', tipo:'wa',
        icono:'💚', color:'#25d366', label:'WhatsApp',
        titulo:h.resumen||'WhatsApp enviado', estado:'', ejecutivo:h.ejecutivo||'' });
    }
  });

  // 3. Cierres — como entradas enriquecidas si no hay ya entry tipo=cierre ese día
  _getCierres().filter(x=>
    (x._clienteId && String(x._clienteId)===String(c.id)) ||
    (x.clienteNombre||'').toUpperCase().trim()===(c.nombre||'').toUpperCase().trim()
  ).forEach(h=>{
    const yaEnBit = (c.bitacora||[]).some(e=>e.fecha===h.fechaRegistro && e.tipo==='cierre');
    if(!yaEnBit){
      entries.push({ fecha:h.fechaRegistro||'0000-00-00', hora:'', tipo:'poliza',
        icono:'📄', color:'#28a745', label:'Póliza',
        titulo:`Póliza registrada — ${h.aseguradora||'—'} · ${h.polizaNueva||'—'} · ${fmt(h.primaTotal||0)}`,
        estado:'', ejecutivo:h.ejecutivo||'' });
    }
  });

  // Ordenar: más reciente primero
  entries.sort((a,b)=>{
    const d = b.fecha.localeCompare(a.fecha);
    return d!==0 ? d : (b.hora||'').localeCompare(a.hora||'');
  });
  return entries;
}

function _renderClienteTimeline(c){
  const entries = _buildClienteTimeline(c);
  if(!entries.length) return '<div style="color:var(--muted);font-size:12px;padding:16px;text-align:center">Sin historial de actividad aún</div>';

  const today = new Date().toISOString().split('T')[0];
  const ayer  = new Date(Date.now()-86400000).toISOString().split('T')[0];
  const fmtFecha = f => {
    if(f===today) return 'Hoy';
    if(f===ayer)  return 'Ayer';
    if(!f||f==='0000-00-00') return '—';
    return new Date(f+'T12:00:00').toLocaleDateString('es-EC',{day:'numeric',month:'short',year:'numeric'});
  };

  // Agrupar por fecha
  const grupos = {};
  entries.forEach(e=>{ if(!grupos[e.fecha]) grupos[e.fecha]=[]; grupos[e.fecha].push(e); });

  return Object.keys(grupos).sort((a,b)=>b.localeCompare(a)).map(fecha=>`
    <div style="margin-bottom:14px">
      <div style="font-size:10px;font-weight:700;text-transform:uppercase;color:var(--muted);
        letter-spacing:.8px;padding:3px 0 7px;border-bottom:1px solid var(--border);margin-bottom:6px">${fmtFecha(fecha)}</div>
      ${grupos[fecha].map(e=>`
        <div style="display:flex;gap:10px;padding:7px 0;border-bottom:1px solid var(--warm)">
          <div style="flex-shrink:0;width:28px;height:28px;border-radius:50%;
            background:${e.color}18;border:1.5px solid ${e.color}55;
            display:flex;align-items:center;justify-content:center;font-size:12px">${e.icono}</div>
          <div style="flex:1;min-width:0">
            <div style="display:flex;align-items:center;gap:5px;flex-wrap:wrap;margin-bottom:2px">
              <span style="font-size:10px;font-weight:700;color:${e.color};text-transform:uppercase;letter-spacing:.4px">${e.label}</span>
              ${e.hora?`<span style="font-size:10px;color:var(--muted)">${e.hora}</span>`:''}
              ${e.estado?`<span style="font-size:9px;padding:1px 5px;border-radius:3px;background:var(--warm);color:var(--muted)">${e.estado}</span>`:''}
              ${e.ejecutivo?`<span style="font-size:10px;color:var(--muted)">· ${e.ejecutivo}</span>`:''}
            </div>
            <div style="font-size:12px;color:var(--ink);line-height:1.45">${e.titulo}</div>
          </div>
        </div>`).join('')}
    </div>`).join('');
}

// Renderiza el historial de bitácora dentro del modal de seguimiento
function _renderBitacoraModal(bitacora){
  const el = document.getElementById('bitacora-lista');
  if(!el) return;
  const lista = Array.isArray(bitacora) ? bitacora : [];
  const cntEl = document.getElementById('modal-bitacora-count');
  if(cntEl) cntEl.textContent = lista.length ? `(${lista.length} entrada${lista.length!==1?'s':''})` : '';
  if(!lista.length){
    el.innerHTML = '<div style="color:var(--muted);font-size:12px;text-align:center;padding:12px">Sin historial de gestión aún</div>';
    return;
  }
  const TIPO_ICON = { manual:'💬', sistema:'⚙️', cotizacion:'📋', cierre:'✅' };
  el.innerHTML = lista.map((e,i)=>{
    const cfg  = ESTADOS_RELIANCE[e.estado] || {};
    const icon = TIPO_ICON[e.tipo] || '💬';
    const isFirst = i===0;
    return `<div style="display:flex;gap:10px;padding:10px 0;${i>0?'border-top:1px solid var(--border)':''}">
      <div style="flex-shrink:0;width:32px;height:32px;border-radius:50%;
        background:${isFirst?'var(--accent2)':'var(--warm)'};
        display:flex;align-items:center;justify-content:center;font-size:14px">
        ${icon}
      </div>
      <div style="flex:1;min-width:0">
        <div style="display:flex;align-items:center;gap:6px;flex-wrap:wrap;margin-bottom:3px">
          <span style="font-weight:600;font-size:12px">${e.ejecutivo}</span>
          <span style="font-size:10px;color:var(--muted)">${e.fecha} ${e.hora||''}</span>
          ${e.estado?`<span class="badge ${cfg.badge||'badge-gray'}" style="font-size:9px;padding:1px 5px">${cfg.icon||''} ${e.estado}</span>`:''}
        </div>
        ${e.nota?`<div style="font-size:12px;color:var(--ink);line-height:1.4">${e.nota}</div>`:''}
      </div>
    </div>`;
  }).join('');
}

function guardarSeguimiento(){
  const c=DB.find(x=>String(x.id)===String(currentSegIdx)); if(!c) return;
  const nota = document.getElementById('seg-nota').value.trim();
  const estadoAnterior = c.estado;
  c._dirty = true;
  c.estado = currentSegEstado;
  c.ultimoContacto = new Date().toISOString().split('T')[0];
  // Agregar entrada a bitácora
  _bitacoraAdd(c, nota, 'manual');
  // Si cambió el estado sin nota, registrar el cambio automáticamente
  if(!nota && estadoAnterior !== currentSegEstado){
    c.bitacora[0].nota = `Estado cambiado: ${estadoAnterior} → ${currentSegEstado}`;
    c.bitacora[0].tipo = 'sistema';
  }
  saveDB();
  sincronizarCotizPorCliente(c.id, c.nombre, c.ci, currentSegEstado);
  closeModal('modal-seguimiento'); renderSeguimiento(); renderDashboard();
  showToast('Seguimiento actualizado');
}
function guardarSeguimientoYCierre(){
  const c=DB.find(x=>String(x.id)===String(currentSegIdx)); if(!c) return;
  const notaYC = document.getElementById('seg-nota').value.trim();
  // Persiste estado actual en DB y abre el cierre directamente.
  // Desde EMISIÓN: el estado se queda como EMISIÓN y guardarCierreVenta lo lleva a RENOVADO (con EMITIDO en bitácora).
  // Desde EMITIDO: guarda EMITIDO y guardarCierreVenta lo lleva a RENOVADO.
  const estadoAnterior = c.estado;
  c._dirty = true;
  c.estado = currentSegEstado;
  c.ultimoContacto = new Date().toISOString().split('T')[0];
  const notaFinal = notaYC || (estadoAnterior !== currentSegEstado ? `Estado: ${estadoAnterior} → ${currentSegEstado}` : '');
  if(notaFinal) _bitacoraAdd(c, notaFinal, notaYC ? 'cierre' : 'sistema');
  saveDB();
  closeModal('modal-seguimiento');
  abrirCierreDesdeCliente(c.id, true);
}

// ══════════════════════════════════════════════════════
//  COTIZADOR
// ══════════════════════════════════════════════════════
const ESTADOS_RELIANCE = {"PENDIENTE": {"cod": 1, "label": "Pendiente", "color": "#9e9e9e", "badge": "badge-gray", "icon": "⚪", "grupo": "gestion", "def": "Cliente en gestión de renovación. Sin confirmación aún."}, "INSPECCIÓN": {"cod": 2, "label": "Inspección", "color": "#ff9800", "badge": "badge-orange", "icon": "🔍", "grupo": "gestion", "def": "Vehículo pendiente de inspección requerida por aseguradora."}, "EMISIÓN": {"cod": 3, "label": "Emisión", "color": "#2196f3", "badge": "badge-blue", "icon": "📝", "grupo": "gestion", "def": "Póliza en trámite de emisión. Aún no activa."}, "EMITIDO": {"cod": 4, "label": "Emitido", "color": "#1976d2", "badge": "badge-blue2", "icon": "📄", "grupo": "gestion", "def": "Póliza emitida formalmente y activa."}, "RENOVADO": {"cod": 5, "label": "Renovado", "color": "#2d6a4f", "badge": "badge-green", "icon": "✅", "grupo": "positivo", "def": "Renovación completada exitosamente."}, "PÓLIZA VIGENTE": {"cod": 6, "label": "Póliza Vigente", "color": "#388e3c", "badge": "badge-green2", "icon": "🛡", "grupo": "positivo", "def": "Póliza activa dentro del período de cobertura."}, "CRÉDITO VENCIDO": {"cod": 7, "label": "Crédito Vencido", "color": "#e65100", "badge": "badge-orange2", "icon": "⏰", "grupo": "riesgo", "def": "Cliente mantiene cuotas vencidas con el banco que financia el vehículo."}, "CRÉDITO CANCELADO": {"cod": 8, "label": "Crédito Cancelado", "color": "#5d4037", "badge": "badge-brown", "icon": "💳", "grupo": "riesgo", "def": "Cliente canceló totalmente su crédito con el banco."}, "SINIESTRO": {"cod": 9, "label": "Siniestro", "color": "#c62828", "badge": "badge-red2", "icon": "🚨", "grupo": "riesgo", "def": "Cliente con evento reportado en trámite con la aseguradora."}, "PÉRDIDA TOTAL": {"cod": 10, "label": "Pérdida Total", "color": "#b71c1c", "badge": "badge-red", "icon": "🚗", "grupo": "riesgo", "def": "Vehículo declarado pérdida total por la aseguradora."}, "PÓLIZA ANULADA": {"cod": 11, "label": "Póliza Anulada", "color": "#c84b1a", "badge": "badge-orange3", "icon": "❌", "grupo": "cierre", "def": "Póliza cancelada antes del vencimiento."}, "ENDOSO": {"cod": 12, "label": "Endoso", "color": "#6b5b95", "badge": "badge-purple", "icon": "📋", "grupo": "cierre", "def": "Cliente realizó el seguro directamente sin intermediación de Reliance."}, "NO RENOVADO": {"cod": 13, "label": "No Renovado", "color": "#795548", "badge": "badge-brown2", "icon": "🚫", "grupo": "cierre", "def": "Cliente decidió no renovar la póliza al vencimiento."}, "INUBICABLE": {"cod": 14, "label": "Inubicable", "color": "#607d8b", "badge": "badge-slate", "icon": "📵", "grupo": "cierre", "def": "No se logra contacto con el cliente tras varios intentos."}, "NO ASEGURABLE": {"cod": 15, "label": "No Asegurable", "color": "#37474f", "badge": "badge-dark", "icon": "⛔", "grupo": "cierre", "def": "Riesgo rechazado por aseguradora."}, "AUTO VENDIDO": {"cod": 16, "label": "Auto Vendido", "color": "#4caf50", "badge": "badge-teal", "icon": "🔑", "grupo": "cierre", "def": "Cliente ya no posee el vehículo asegurado."}, "CLIENTE FALLECIDO": {"cod": 17, "label": "Cliente Fallecido", "color": "#546e7a", "badge": "badge-slate2", "icon": "🕊", "grupo": "cierre", "def": "Titular fallecido. Se cierra gestión o se coordina con herederos."}};

const GRUPOS_ESTADOS = {"gestion": {"titulo": "📋 En Gestión", "estados": ["PENDIENTE", "INSPECCIÓN", "EMISIÓN", "EMITIDO"]}, "positivo": {"titulo": "✅ Positivos", "estados": ["RENOVADO", "PÓLIZA VIGENTE"]}, "riesgo": {"titulo": "⚠️ Riesgo", "estados": ["CRÉDITO VENCIDO", "CRÉDITO CANCELADO", "SINIESTRO", "PÉRDIDA TOTAL"]}, "cierre": {"titulo": "🔒 Cierre/Baja", "estados": ["PÓLIZA ANULADA", "ENDOSO", "NO RENOVADO", "INUBICABLE", "NO ASEGURABLE", "AUTO VENDIDO", "CLIENTE FALLECIDO"]}};

// ──────────────────────────────────────────────────
//  FORMAS DE PAGO — Catálogo exacto de la aseguradora
// ──────────────────────────────────────────────────
const FORMAS_PAGO = {
  CONTADO: {
    icon:'💵', label:'Contado',
    opciones:[
      {val:'CONTADO TRANSFERENCIA', label:'Transferencia'},
      {val:'CONTADO DEPOSITO',      label:'Depósito'},
    ]
  },
  DEBITO_PRODUBANCO: {
    icon:'🏦', label:'Débito Produbanco',
    opciones:[1,2,3,4,5,6,7,8,9,10,11,12].map(n=>({
      val:`${n} DEBITO${n>1?'S':''} PRODUBANCO`,
      label:`${n} débito${n>1?'s':''}`
    }))
  },
  TC: {
    icon:'💳', label:'Tarjeta Crédito',
    opciones:[3,6,9,12].map(n=>({val:`TC ${n} MESES`, label:`TC ${n} meses`}))
  },
  DEBITO_OTROS: {
    icon:'🏧', label:'Otros Bancos',
    opciones:[8,10,12].map(n=>({
      val:`${n} DEBITOS OTROS BANCOS`, label:`${n} débitos`
    }))
  },
  PAGOS_DIRECTOS: {
    icon:'📋', label:'Pagos Directos',
    opciones:[3,8,9,10,12].map(n=>({
      val:`${n} PAGOS DIRECTOS`, label:`${n} pagos`
    }))
  },
  CHEQUES: {
    icon:'📝', label:'Cheques',
    opciones:[{val:'CHEQUES POSFECHADOS', label:'Cheques Posfechados'}]
  },
  DEBITO_REC_TC: {
    icon:'🔄', label:'Rec. TC',
    opciones:[2,3,4,5,6,7,8,9,10,11,12].map(n=>({
      val:`${n} DEBITOS RECURRENTES TC`, label:`${n} cuotas`
    }))
  },
  MIXTO: {
    icon:'🔀', label:'Pago Mixto',
    opciones:[{val:'CUOTA INICIAL CONTADO- DIFERENCIA DEBITOS', label:'Entrada + Débitos'}]
  }
};

/** Clasifica forma de pago exacta en su grupo */
function _fpTipo(val){
  if(!val) return null;
  const v=val.toUpperCase().trim();
  if(v.startsWith('CONTADO')) return 'CONTADO';
  if(v.includes('PRODUBANCO')) return 'DEBITO_PRODUBANCO';
  if(v.startsWith('TC ')) return 'TC';
  if(v.includes('OTROS BANCOS')) return 'DEBITO_OTROS';
  if(v.includes('PAGOS DIRECTOS')) return 'PAGOS_DIRECTOS';
  if(v.includes('CHEQUES')) return 'CHEQUES';
  if(v.includes('RECURRENTES TC')) return 'DEBITO_REC_TC';
  if(v.includes('CUOTA INICIAL')||v==='MIXTO') return 'MIXTO';
  // Compatibilidad con valores abstractos legados
  if(v==='DEBITO_BANCARIO') return 'DEBITO_PRODUBANCO';
  if(v==='TARJETA_CREDITO') return 'TC';
  if(v==='DEBITO_RECURRENTE_TC') return 'DEBITO_REC_TC';
  return null;
}

/** Extrae número de cuotas de la forma de pago exacta */
function _fpCuotas(val){
  if(!val) return 1;
  const v=val.toUpperCase().trim();
  if(v.startsWith('TC ')){ const m=v.match(/TC\s+(\d+)/); return m?parseInt(m[1]):1; }
  const m=v.match(/^(\d+)\s/);
  return m?parseInt(m[1]):1;
}

const ASEGURADORAS={
  // ── Las 7 aseguradoras del cotizador Excel (PRODUCTOS PRODUBANCO) ──────────
  ZURICH:{
    color:'#4a4a4a', pnMin:550, tcMax:12, debMax:10,
    tasa:0.043,                           // 4.3% fijo (fila 9, col B)
    axaDisponible:false, vidaDefault:0,
    pisoTC:0, pisoDeb:0,                  // sin piso de cuota mínima
    extraFijo:0,
    resp_civil:30000, muerte_ocupante:10000, muerte_titular:10000, gastos_medicos:2000,
    amparo:'COMPLETO',
    auto_sust:'10 días · siniestro >$1,000',
    legal:'SÍ', exequial:'SÍ',
    vida:'N/A', enf_graves:'N/A', renta_hosp:'N/A', sepelio:'N/A',
    telemedicina:'N/A', dental:'N/A', medico_dom:'N/A',
    ded_parcial:'10%VS / 1%VA / mín.$350',
    ded_daño:'15%VA', ded_robo_sin:'30%VA', ded_robo_con:'20%VA',
  },
  LATINA:{
    color:'#6b5b95', pnMin:350, tcMax:12, debMax:10,
    tasa:0.038,                           // 3.8% fijo (fila 9, col C)
    axaDisponible:false, vidaDefault:50,  // prima vida $50 (G8 Excel)
    pisoTC:0, pisoDeb:0,
    extraFijo:0,
    resp_civil:30000, muerte_ocupante:10000, muerte_titular:10000, gastos_medicos:2500,
    amparo:'COMPLETO',
    auto_sust:'10 días · siniestro >$1,150',
    legal:'SÍ', exequial:'SÍ',
    vida:'$10,000–$20,000 (plan)', enf_graves:'N/A',
    renta_hosp:'$50/día · máx.10 días', sepelio:'N/A',
    telemedicina:'E-DOCTOR (med.general, psicología, nutrición)',
    dental:'Prevención y cirugía 70–100%', medico_dom:'N/A',
    ded_parcial:'10%VS / 1%VA / mín.$350 (VA>$15k) · $250 (VA≤$15k)',
    ded_daño:'25%VA (VA≤$15k) / 20%VA (VA>$15k)',
    ded_robo_sin:'20%VA', ded_robo_con:'20%VA',
  },
  GENERALI:{
    color:'#c84b1a', pnMin:400, tcMax:12, debMax:10,
    tasa:0.035,                           // 3.5% fijo (fila 9, col D)
    axaDisponible:false, vidaDefault:0,
    pisoTC:35, pisoDeb:35,                // cuota mínima $35
    extraFijo:0,
    resp_civil:35000, muerte_ocupante:8000, muerte_titular:8000, gastos_medicos:3000,
    amparo:'CON COSTO adicional',
    auto_sust:'10 días · siniestro >$1,600+IVA',
    legal:'SÍ', exequial:'SÍ',
    vida:'N/A', enf_graves:'N/A', renta_hosp:'N/A', sepelio:'N/A',
    telemedicina:'N/A', dental:'N/A', medico_dom:'N/A',
    ded_parcial:'AVEO/SPARK/KIA: 15%VS/2%VA/mín.$500 · Otros: 10%VS/1%VA/mín.$250',
    ded_daño:'por tabla de modelo', ded_robo_sin:'por tabla de modelo', ded_robo_con:'por tabla de modelo',
  },
  ADS:{
    color:'#e63946', pnMin:350, tcMax:12, debMax:10,
    tasa:0.045,                           // 4.5% fijo (fila 9, col E)
    axaDisponible:false, vidaDefault:0,
    pisoTC:0, pisoDeb:0,
    extraFijo:80,                         // +$80 fijo ADS (fila 20, col E)
    resp_civil:30000, muerte_ocupante:10000, muerte_titular:10000, gastos_medicos:3000,
    amparo:'COMPLETO',
    auto_sust:'15 días · siniestro >$1,500',
    legal:'SÍ', exequial:'NO',
    vida:'N/A', enf_graves:'N/A', renta_hosp:'N/A', sepelio:'N/A',
    telemedicina:'N/A', dental:'N/A', medico_dom:'N/A',
    ded_parcial:'10%VS / 1%VA / mín.$250',
    ded_daño:'15%VA', ded_robo_sin:'20%VA', ded_robo_con:'20%VA',
  },
  SWEADEN:{
    color:'#1a4c84', pnMin:350, tcMax:9, debMax:10,
    tasa:0.035,                           // 3.5% fijo (fila 9, col F)
    axaDisponible:true,  vidaDefault:59.99, // AXA disponible; prima vida $59.99 (F8)
    pisoTC:50, pisoDeb:50,                // cuota mínima $50
    extraFijo:0,
    resp_civil:30000, muerte_ocupante:5000, muerte_titular:10000, gastos_medicos:2000,
    amparo:'COMPLETO',
    auto_sust:'c/AXA: 30 días sin mínimo · s/AXA: >$1,500 (VA>$20k) 7 días',
    legal:'SÍ', exequial:'NO',
    vida:'$5,000', enf_graves:'$2,500',
    renta_hosp:'$25/día · máx.30 días', sepelio:'$500',
    telemedicina:'6 consultas', dental:'N/A', medico_dom:'Copago $10/evento',
    ded_parcial:'10%VS / 1%VA / mín.$250 (por tabla)',
    ded_daño:'20%VA', ded_robo_sin:'20%VA', ded_robo_con:'20%VA',
  },
  MAPFRE:{
    color:'#b8860b', pnMin:400, tcMax:9, debMax:10,
    tasa:0.037,                           // 3.7% fijo (fila 9, col G)
    axaDisponible:false, vidaDefault:63.19, // prima vida $63.19 (G6)
    pisoTC:0, pisoDeb:0,
    extraFijo:0,
    resp_civil:30000, muerte_ocupante:6000, muerte_titular:null, gastos_medicos:3000,
    amparo:'COMPLETO',
    auto_sust:'10 días · siniestro >$1,250',
    legal:'SÍ', exequial:'NO',
    vida:'$5,000', enf_graves:'Anticipo 50% cob. principal',
    renta_hosp:'$20/día acc. · máx.30 días · hasta $600', sepelio:'$500',
    telemedicina:'SÍ', dental:'N/A', medico_dom:'SÍ',
    ded_parcial:'10%VS / 1.5%VA / mín.$350',
    ded_daño:'15%VA', ded_robo_sin:'30%VA', ded_robo_con:'15%VA',
  },
  ALIANZA:{
    color:'#2d6a4f', pnMin:350, tcMax:9, debMax:10,
    tasa:0.032,                           // 3.2% fijo (fila 9, col H)
    axaDisponible:false, vidaDefault:55.74, // prima vida $55.74 (H8)
    pisoTC:0, pisoDeb:50,                 // débito: cuota mínima $50
    extraFijo:0,
    resp_civil:30000, muerte_ocupante:5000, muerte_titular:null, gastos_medicos:2500,
    amparo:'COMPLETO',
    auto_sust:'Sedán/SP: >$500 · SUV ≤$40k: >$1,000 · SUV >$60k: >$60,001',
    legal:'SÍ', exequial:'SÍ',
    vida:'$5,000', enf_graves:'$2,500',
    renta_hosp:'$20/día · máx.25 días · hasta $500', sepelio:'$500',
    telemedicina:'Orientación médica telefónica sin límite',
    dental:'N/A', medico_dom:'Copago $10 · titular y cónyuge',
    ded_parcial:'10%VS / 1%VA / mín.$250',
    ded_daño:'15%VA', ded_robo_sin:'15%VA', ded_robo_con:'15%VA',
  },
  // ── Aseguradoras adicionales en cartera (sin datos Excel detallados) ─────────
  'ASEG. DEL SUR':{
    color:'#e63946', pnMin:350, tcMax:12, debMax:10,
    tasa:0.034, axaDisponible:false, vidaDefault:0,
    pisoTC:0, pisoDeb:0, extraFijo:0,
    resp_civil:30000, muerte_ocupante:10000, muerte_titular:10000, gastos_medicos:3000,
    amparo:'COMPLETO', auto_sust:'15 días · siniestro >$1,500',
    legal:'SÍ', exequial:'NO',
    vida:'N/A', enf_graves:'N/A', renta_hosp:'N/A', sepelio:'N/A',
    telemedicina:'N/A', dental:'N/A', medico_dom:'N/A',
    ded_parcial:'10%VS / 1%VA / mín.$250',
    ded_daño:'15%VA', ded_robo_sin:'20%VA', ded_robo_con:'20%VA',
  },
  EQUINOCCIAL:{
    color:'#0077b6', pnMin:350, tcMax:12, debMax:10,
    tasa:0.034, axaDisponible:false, vidaDefault:0,
    pisoTC:0, pisoDeb:0, extraFijo:0,
    resp_civil:30000, muerte_ocupante:8000, muerte_titular:8000, gastos_medicos:2500,
    amparo:'COMPLETO', auto_sust:'10 días · siniestro >$1,000',
    legal:'SÍ', exequial:'SÍ',
    vida:'N/A', enf_graves:'N/A', renta_hosp:'N/A', sepelio:'N/A',
    telemedicina:'N/A', dental:'N/A', medico_dom:'N/A',
    ded_parcial:'10%VS / 1%VA / mín.$250',
    ded_daño:'20%VA', ded_robo_sin:'20%VA', ded_robo_con:'20%VA',
  },
  ATLANTIDA:{
    color:'#457b9d', pnMin:350, tcMax:12, debMax:10,
    tasa:0.036, axaDisponible:false, vidaDefault:0,
    pisoTC:0, pisoDeb:0, extraFijo:0,
    resp_civil:30000, muerte_ocupante:5000, muerte_titular:5000, gastos_medicos:2000,
    amparo:'COMPLETO', auto_sust:'10 días · siniestro >$1,200',
    legal:'SÍ', exequial:'NO',
    vida:'N/A', enf_graves:'N/A', renta_hosp:'N/A', sepelio:'N/A',
    telemedicina:'N/A', dental:'N/A', medico_dom:'N/A',
    ded_parcial:'10%VS / 1%VA / mín.$250',
    ded_daño:'20%VA', ded_robo_sin:'20%VA', ded_robo_con:'20%VA',
  },
  AIG:{
    color:'#2a9d8f', pnMin:400, tcMax:12, debMax:10,
    tasa:0.038, axaDisponible:false, vidaDefault:0,
    pisoTC:0, pisoDeb:0, extraFijo:0,
    resp_civil:40000, muerte_ocupante:10000, muerte_titular:10000, gastos_medicos:5000,
    amparo:'COMPLETO', auto_sust:'15 días · siniestro >$800',
    legal:'SÍ', exequial:'SÍ',
    vida:'N/A', enf_graves:'N/A', renta_hosp:'N/A', sepelio:'N/A',
    telemedicina:'N/A', dental:'N/A', medico_dom:'N/A',
    ded_parcial:'10%VS / 1%VA / mín.$250',
    ded_daño:'15%VA', ded_robo_sin:'15%VA', ded_robo_con:'15%VA',
  },
};

// ── Helpers de cálculo (equivalentes exactos del Excel PRODUCTOS PRODUBANCO) ──

// Derechos de Emisión — escala tiered (fila 15 del Excel)
function _calcDerechosEmision(pn){
  if(pn > 4000) return 9;
  if(pn > 2000) return 7;
  if(pn > 1000) return 5;
  if(pn > 500)  return 3;
  if(pn > 250)  return 1;
  return 0.50;
}

// Cálculo completo de prima para una aseguradora
// Incluye derechos, campesino, SuperBancos, AXA, IVA y vida (post-IVA)
function calcPrima(va, tasa, pnMin=0, axaIncluido=false, vidaPrima=0, extraFijo=0){
  const pnCalc   = va * tasa;
  const pn       = Math.max(pnCalc, pnMin>0 ? pnMin : 0);
  const aplicaMin= pnMin > 0 && pnCalc < pnMin;

  const der  = _calcDerechosEmision(pn);
  const camp = Math.round(pn * 0.005 * 100) / 100;
  const sb   = Math.round(pn * 0.035 * 100) / 100;
  const axa  = axaIncluido ? Math.round((60/1.15) * 100) / 100 : 0; // $52.17 neto

  // Subtotal pre-IVA: pn + cargos + AXA + extra fijo (ADS: +$80)
  const sub = Math.round((pn + der + camp + sb + axa + extraFijo) * 100) / 100;
  const iva = Math.round(sub * 0.15 * 100) / 100;

  // Vida se agrega DESPUÉS del IVA — no tributa (fila 22 del Excel)
  const total = Math.round((sub + iva + vidaPrima) * 100) / 100;

  return { pn, der, camp, sb, axa, extraFijo, sub, iva, vida:vidaPrima, total,
           vaEfectivo:va, vaOriginal:va, ajustado:aplicaMin, pnCalc, tasa };
}

// Cuotas TC — respeta tcMax y cuota mínima por aseguradora
function calcCuotasTc(total, tcMax, nCuotas, piso=0){
  let n = Math.min(nCuotas, tcMax);
  if(piso > 0){ while(n > 1 && (total/n) < piso) n--; }
  return { n, cuota: Math.round(total/n*100)/100 };
}

// Cuotas débito — respeta cuota mínima por aseguradora
function calcCuotasDeb(total, nCuotas, piso=0){
  let n = nCuotas;
  if(piso > 0){ while(n > 1 && (total/n) < piso) n--; }
  return { n, cuota: Math.round(total/n*100)/100 };
}

// Calcula y muestra la fecha "Vigencia Hasta" (desde + 1 año - 1 día) en el cotizador
function cotActualizarHasta(){
  const desdeVal = document.getElementById('cot-desde')?.value;
  const hastaEl  = document.getElementById('cot-hasta-display');
  if(!hastaEl) return;
  if(!desdeVal){ hastaEl.style.display='none'; return; }
  const h = new Date(desdeVal + 'T00:00:00');
  h.setFullYear(h.getFullYear() + 1);
  h.setDate(h.getDate() - 1);
  const hastaISO = h.toISOString().split('T')[0];
  const [y,m,d] = hastaISO.split('-');
  hastaEl.textContent = `Vence: ${d}/${m}/${y}`;
  hastaEl.style.display = 'block';
}

// ID del cliente seleccionado en el cotizador (para vincular clienteId al guardar)
let _cotizClienteId = '';

function prefillCotizador(c){
  _cotizClienteId = c.id ? String(c.id) : '';
  // — Datos del cliente —
  document.getElementById('cot-nombre').value=c.nombre||'';
  document.getElementById('cot-ci').value=c.ci||'';
  document.getElementById('cot-cel').value=c.celular||'';
  document.getElementById('cot-email').value=c.correo||'';
  // Ciudad: match exacto primero; si no coincide, búsqueda parcial (ej: "GYE" → GUAYAQUIL)
  const ciudadEl=document.getElementById('cot-ciudad');
  if(ciudadEl && c.ciudad){
    const cu = c.ciudad.toUpperCase().trim();
    let opt = [...ciudadEl.options].find(o=>o.value.toUpperCase()===cu);
    if(!opt) opt = [...ciudadEl.options].find(o=>cu.startsWith(o.value.slice(0,4).toUpperCase()) || o.value.toUpperCase().startsWith(cu.slice(0,4)));
    if(opt) ciudadEl.value=opt.value;
  }
  // Región: del cliente si es válida; si no → deducir desde ciudad del cliente o ciudad seleccionada
  const regionEl=document.getElementById('cot-region');
  if(regionEl){
    let asignado=false;
    if(c.region){
      const opt=[...regionEl.options].find(o=>o.value.toUpperCase()===c.region.toUpperCase());
      if(opt){ regionEl.value=opt.value; asignado=true; }
    }
    if(!asignado){
      // Deducir región desde el nombre de ciudad guardado (más confiable que el select)
      const regionDeducida = _ciudadToRegion(c.ciudad) || _ciudadToRegion(ciudadEl?.value||'');
      if(regionDeducida) regionEl.value=regionDeducida;
    }
  }
  // Tipo de póliza: si el cliente tiene póliza anterior → RENOVACION, si no → NUEVO
  const tipoEl=document.getElementById('cot-tipo');
  if(tipoEl) tipoEl.value=(c.polizaNueva||c.poliza||c.polizaAnterior)?'RENOVACION':'NUEVO';

  // — Datos del vehículo —
  _setMarcaSelect(document.getElementById('cot-marca'), c.marca);
  document.getElementById('cot-anio').value=c.anio||(new Date().getFullYear());
  document.getElementById('cot-modelo').value=c.modelo||'';
  document.getElementById('cot-placa').value=c.placa||'';
  // Usar valor depreciado (AI del Excel) como VA sugerido si está disponible
  document.getElementById('cot-va').value=c.dep||c.va||20000;
  const depHintEl=document.getElementById('cot-dep-hint');
  if(depHintEl){
    if(c.dep && c.dep !== c.va){
      depHintEl.textContent=`💡 Valor depreciado Excel: $${Number(c.dep).toLocaleString('es-EC')} · VA anterior: $${Number(c.va||0).toLocaleString('es-EC')}`;
      depHintEl.style.display='block';
    } else {
      depHintEl.style.display='none';
    }
  }
  if(document.getElementById('cot-color'))  document.getElementById('cot-color').value=c.color||'';
  if(document.getElementById('cot-motor'))  document.getElementById('cot-motor').value=c.motor||'';
  if(document.getElementById('cot-chasis')) document.getElementById('cot-chasis').value=c.chasis||'';

  // — Póliza anterior (para renovación) —
  if(document.getElementById('cot-poliza-anterior')) document.getElementById('cot-poliza-anterior').value=c.polizaAnterior||c.polizaNueva||c.poliza||'';
  if(document.getElementById('cot-aseg-anterior'))   document.getElementById('cot-aseg-anterior').value=c.aseguradoraAnterior||c.aseguradora||'';

  // — Vigencia nueva: pre-llenar desde = c.hasta (nueva póliza arranca el mismo día que vence la anterior) —
  const desdeEl = document.getElementById('cot-desde');
  if(desdeEl && c.hasta){
    desdeEl.value = c.hasta;
    cotActualizarHasta(); // actualiza el display de "Vence:"
  }

  // — Hint de vencimiento póliza anterior —
  const venceHintEl = document.getElementById('cot-vence-hint');
  if(venceHintEl){
    if(c.hasta){
      const [y,m,d] = c.hasta.split('-');
      const desdeRef = c.desde ? (()=>{ const [dy,dm,dd]=c.desde.split('-'); return `${dd}/${dm}/${dy} → `; })() : '';
      venceHintEl.textContent = `⚠️ Póliza anterior vence: ${desdeRef}${d}/${m}/${y}`;
      venceHintEl.style.display = 'block';
    } else {
      venceHintEl.style.display = 'none';
    }
  }

  // Hint de tasas de referencia para el ejecutivo
  const tasaHintEl=document.getElementById('cot-tasa-hint');
  if(tasaHintEl){
    if(c.tasaAnterior||c.tasaRenov){
      const parts=[];
      if(c.tasaAnterior) parts.push(`Tasa ant.: ${c.tasaAnterior}%`);
      if(c.tasaRenov)    parts.push(`Tasa renov. actual: ${c.tasaRenov}%`);
      tasaHintEl.textContent=`📊 Referencia: ${parts.join(' · ')}`;
      tasaHintEl.style.display='block';
    } else {
      tasaHintEl.style.display='none';
    }
  }
}
function prefillCotizador_show(id){
  const c=DB.find(x=>String(x.id)===String(id)); if(!c) return;
  prefillCotizador(c);
  // Mostrar nombre en el buscador para feedback visual
  const buscarEl=document.getElementById('cot-buscar-cliente');
  if(buscarEl) buscarEl.value=c.nombre;
  showPage('cotizador'); setTimeout(calcCotizacion,200);
}

// ─── Búsqueda de cliente dentro del cotizador ────────────────────────────────
function buscarClienteCotizador(q){
  const box=document.getElementById('cot-sugerencias'); if(!box) return;
  if(!q||q.length<2){ box.style.display='none'; return; }
  const ql=q.toLowerCase();
  const matches=myClientes().filter(c=>
    (c.nombre||'').toLowerCase().includes(ql)||
    String(c.ci||'').includes(ql)||
    (c.placa||'').toLowerCase().includes(ql)
  ).slice(0,8);
  if(!matches.length){ box.style.display='none'; return; }
  box.innerHTML=matches.map(c=>`
    <div onclick="seleccionarClienteCotizador('${c.id}')"
      style="padding:9px 14px;cursor:pointer;border-bottom:1px solid var(--border);font-size:13px"
      onmouseover="this.style.background='var(--warm)'"
      onmouseout="this.style.background=''">
      <div style="font-weight:600">${c.nombre}</div>
      <div style="font-size:11px;color:var(--muted)">${c.ci||'—'} &nbsp;·&nbsp; ${c.marca||''} ${c.modelo||''} &nbsp;·&nbsp; ${c.placa||'—'}</div>
    </div>`).join('');
  box.style.display='block';
}

function seleccionarClienteCotizador(id){
  const c=DB.find(x=>String(x.id)===String(id)); if(!c) return;
  prefillCotizador(c);
  const buscarEl=document.getElementById('cot-buscar-cliente');
  if(buscarEl) buscarEl.value=c.nombre;
  const box=document.getElementById('cot-sugerencias');
  if(box) box.style.display='none';
  setTimeout(calcCotizacion,200);
  showToast(`✓ ${c.nombre.split(' ')[0]} cargado — datos de vehículo pre-llenados`,'success');
}

function limpiarCotizador(){
  ['cot-nombre','cot-ci','cot-cel','cot-email','cot-modelo','cot-placa',
   'cot-color','cot-motor','cot-chasis','cot-poliza-anterior','cot-aseg-anterior'].forEach(id=>{
    const el=document.getElementById(id); if(el) el.value='';
  });
  const va=document.getElementById('cot-va'); if(va) va.value=20000;
  const anio=document.getElementById('cot-anio'); if(anio) anio.value=new Date().getFullYear();
  const buscarEl=document.getElementById('cot-buscar-cliente'); if(buscarEl) buscarEl.value='';
  const box=document.getElementById('cot-sugerencias'); if(box) box.style.display='none';
  _cotizClienteId = '';
  // Cancelar edición pendiente para no marcar REEMPLAZADA si el ejecutivo abandona sin guardar
  window._editandoCotizId = null; window._editandoCotizCodigo = null; window._editandoCotizVersion = null;
}

function calcCotizacion(){
  cotActualizarHasta(); // mantener fecha Vence actualizada
  const va  = parseFloat(document.getElementById('cot-va')?.value)||0;
  const ext = parseFloat(document.getElementById('cot-extras')?.value)||0;
  const vaT = va + ext;
  if(vaT < 500){ showToast('Ingrese un valor asegurado válido','error'); return; }

  const cuotasTcReq  = parseInt(document.getElementById('cot-cuotas-tc')?.value)||12;
  const cuotasDebReq = parseInt(document.getElementById('cot-cuotas-deb')?.value)||10;

  // Leer toggle AXA (SWEADEN)
  const axaActivo = document.getElementById('cot-axa')?.checked || false;

  // Leer primas de vida por aseguradora
  const vidaInputs = {
    LATINA:  parseFloat(document.getElementById('cot-vida-latina')?.value)||0,
    SWEADEN: parseFloat(document.getElementById('cot-vida-sweaden')?.value)||0,
    MAPFRE:  parseFloat(document.getElementById('cot-vida-mapfre')?.value)||0,
    ALIANZA: parseFloat(document.getElementById('cot-vida-alianza')?.value)||0,
  };

  const selectedAseg = getSelectedAseg();
  if(selectedAseg.length === 0){ showToast('Selecciona al menos una aseguradora','error'); return; }

  const results = Object.entries(ASEGURADORAS)
    .filter(([name]) => selectedAseg.includes(name))
    .map(([name, cfg]) => {
      // Tasa: lee el input de la tarjeta si fue editado; si no, usa el rango de SA configurado en Admin
      const tasa = typeof cfg.tasa === 'function' ? cfg.tasa(vaT) : _getTasaFromCard(name);
      const axaInc   = name === 'SWEADEN' ? axaActivo : false;
      const vida     = vidaInputs[name] !== undefined ? vidaInputs[name] : 0;
      const p = calcPrima(vaT, tasa, cfg.pnMin, axaInc, vida, cfg.extraFijo||0);
      const tc  = calcCuotasTc(p.total, cfg.tcMax, cuotasTcReq,  cfg.pisoTC||0);
      const deb = calcCuotasDeb(p.total, Math.min(cuotasDebReq, cfg.debMax||cuotasDebReq), cfg.pisoDeb||0);
      const planVida = _getPlanVidaNombre(name, vida);
      return { name, cfg, ...p, tc, deb, planVida };
    });

  const minTotal = Math.min(...results.map(r=>r.total));

  // Inicializar almacén de datos por tarjeta (para botón WA y recálculo)
  _cardData = {};
  results.forEach(r => { _cardData[r.name] = { total: r.total, tc: r.tc, deb: r.deb }; });

  document.getElementById('aseg-cards-result').innerHTML = results.map(r => {
    const warnings = [];
    if(r.ajustado) warnings.push(`⚠ Prima mínima aplicada: $${r.pn.toFixed(2)}`);
    if(r.tc.n < Math.min(cuotasTcReq, r.cfg.tcMax)) warnings.push(`⚠ TC ajustado a ${r.tc.n} cuotas`);
    if(r.tc.n < cuotasTcReq && r.cfg.tcMax < cuotasTcReq) warnings.push(`⚠ TC máx. permitido: ${r.cfg.tcMax} cuotas`);
    if(r.deb.n < cuotasDebReq) warnings.push(`⚠ Débito ajustado a ${r.deb.n} cuotas`);

    const axaBadge = r.axa > 0
      ? `<div style="background:#e8f0fb;border:1px solid #1a4c84;border-radius:6px;padding:4px 8px;margin-bottom:6px;font-size:11px;color:#1a4c84;display:flex;justify-content:space-between">
           <span>🚗 Auto sustituto AXA (neto)</span><span style="font-weight:700;font-family:'DM Mono',monospace">${fmt(r.axa)}</span></div>` : '';
    const vidaBadge = r.vida > 0
      ? `<div style="background:#e8f5e9;border:1px solid #2d6a4f;border-radius:6px;padding:4px 8px;margin-bottom:6px;font-size:11px;color:#2d6a4f;display:flex;justify-content:space-between">
           <span>❤ ${r.planVida||'Plan de Vida'}</span><span style="font-weight:700;font-family:'DM Mono',monospace">${fmt(r.vida)}</span></div>` : '';
    const extraBadge = r.extraFijo > 0
      ? `<div class="aseg-row" style="color:#e63946"><span class="aseg-key">Cargo adicional</span><span class="aseg-val">${fmt(r.extraFijo)}</span></div>` : '';

    const s = _safeName(r.name);
    return `<div class="aseg-card${r.total===minTotal?' mejor':''}" id="aseg-card-${s}">
      <div class="aseg-name" style="color:${r.cfg.color}">${r.name}</div>
      ${warnings.length?`<div style="background:#fff8e1;border:1px solid #f0c040;border-radius:6px;padding:5px 8px;margin-bottom:8px;font-size:10px;color:#7a5c00;line-height:1.6">${warnings.join('<br>')}</div>`:''}
      ${axaBadge}${vidaBadge}
      <div class="aseg-row">
        <span class="aseg-key">Tasa</span>
        <span class="aseg-val" style="display:flex;align-items:center;gap:3px">
          <input type="number" id="aseg-tasa-input-${s}"
                 value="${(r.tasa*100).toFixed(2)}"
                 data-computed="1"
                 min="0.01" max="20" step="0.01"
                 title="Editar tasa para esta cotización"
                 style="width:58px;border:1px solid #ccc;border-radius:4px;padding:2px 4px;font-size:12px;font-family:inherit;text-align:right;background:var(--bg,#fff);color:inherit"
                 onchange="this.removeAttribute('data-computed');recalcCard('${r.name}',this.value)"> %
        </span>
      </div>
      <div class="aseg-row"><span class="aseg-key">Prima Neta</span><span class="aseg-val" id="aseg-pn-${s}">${fmt(r.pn)}</span></div>
      <div class="aseg-row"><span class="aseg-key">Der. Emisión</span><span class="aseg-val">${fmt(r.der)}</span></div>
      <div class="aseg-row"><span class="aseg-key">Seg. Campesino</span><span class="aseg-val">${fmt(r.camp)}</span></div>
      <div class="aseg-row"><span class="aseg-key">Super Bancos</span><span class="aseg-val">${fmt(r.sb)}</span></div>
      ${extraBadge}
      <div class="aseg-row"><span class="aseg-key">Subtotal</span><span class="aseg-val" id="aseg-sub-${s}">${fmt(r.sub)}</span></div>
      <div class="aseg-row"><span class="aseg-key">IVA 15%</span><span class="aseg-val" id="aseg-iva-${s}">${fmt(r.iva)}</span></div>
      ${r.vida>0?`<div class="aseg-row" style="color:#2d6a4f"><span class="aseg-key">Vida (post-IVA)</span><span class="aseg-val">${fmt(r.vida)}</span></div>`:''}
      <div class="aseg-total"><span class="aseg-total-key">COSTO TOTAL</span><span class="aseg-total-val" id="aseg-total-${s}">${fmt(r.total)}</span></div>
      <div class="aseg-cuota" id="aseg-tc-${s}">💳 TC ${r.tc.n} cuotas: <b>${fmt(r.tc.cuota)}/mes</b></div>
      <div class="aseg-cuota" id="aseg-deb-${s}">🏦 Débito ${r.deb.n} cuotas: <b>${fmt(r.deb.cuota)}/mes</b></div>
      <div style="margin-top:10px">
        <button class="btn btn-blue btn-xs w-full" onclick="enviarWhatsAppCotiz('${r.name}',_cardData['${r.name}'].total,_cardData['${r.name}'].tc.cuota,_cardData['${r.name}'].tc.n)">📱 Enviar por WhatsApp</button>
      </div>
    </div>`;
  }).join('');

  // Tabla coberturas comparadas
  document.getElementById('coberturas-table-wrap').innerHTML=`<table class="comp-table">
    <thead><tr><th class="col-cob" style="text-align:left">Cobertura / Deducible</th>${results.map(r=>`<th style="color:${r.cfg.color}">${r.name}</th>`).join('')}</tr></thead>
    <tbody>
      <tr class="section-row"><td colspan="${results.length+1}">Coberturas Básicas</td></tr>
      <tr><td class="col-cob">Todo Riesgo</td>${results.map(()=>`<td class="col-val yes">SÍ</td>`).join('')}</tr>
      <tr><td class="col-cob">Pérdida parcial</td>${results.map(()=>`<td class="col-val yes">SÍ</td>`).join('')}</tr>
      <tr><td class="col-cob">Pérdida total robo/daño</td>${results.map(()=>`<td class="col-val yes">SÍ</td>`).join('')}</tr>
      <tr><td class="col-cob">Cobertura Airbags</td>${results.map(()=>`<td class="col-val yes">SÍ</td>`).join('')}</tr>
      <tr><td class="col-cob">Extraterritorial</td>${results.map(()=>`<td class="col-val yes">SÍ</td>`).join('')}</tr>
      <tr><td class="col-cob">Gastos Wincha</td>${results.map(()=>`<td class="col-val yes">SÍ</td>`).join('')}</tr>
      <tr class="section-row"><td colspan="${results.length+1}">Amparos Adicionales</td></tr>
      <tr><td class="col-cob">Responsabilidad Civil</td>${results.map(r=>`<td class="col-val">${fmt(r.cfg.resp_civil)}</td>`).join('')}</tr>
      <tr><td class="col-cob">Muerte acc. ocupante</td>${results.map(r=>`<td class="col-val">${fmt(r.cfg.muerte_ocupante)}</td>`).join('')}</tr>
      <tr><td class="col-cob">Muerte acc. titular</td>${results.map(r=>`<td class="col-val">${r.cfg.muerte_titular?fmt(r.cfg.muerte_titular):'<span class="no">N/A</span>'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Gastos Médicos</td>${results.map(r=>`<td class="col-val">${fmt(r.cfg.gastos_medicos)}</td>`).join('')}</tr>
      <tr><td class="col-cob">Amparo Patrimonial</td>${results.map(r=>`<td class="col-val" style="font-size:10px">${r.cfg.amparo}</td>`).join('')}</tr>
      <tr class="section-row"><td colspan="${results.length+1}">Beneficios y Servicios</td></tr>
      <tr><td class="col-cob">Auto sustituto</td>${results.map(r=>`<td class="col-val" style="font-size:10px">${r.cfg.auto_sust}</td>`).join('')}</tr>
      <tr><td class="col-cob">Asist. Legal en situ</td>${results.map(r=>`<td class="col-val ${r.cfg.legal==='SÍ'?'yes':'no'}">${r.cfg.legal}</td>`).join('')}</tr>
      <tr><td class="col-cob">Asistencia Exequial</td>${results.map(r=>`<td class="col-val ${r.cfg.exequial==='SÍ'?'yes':'no'}">${r.cfg.exequial}</td>`).join('')}</tr>
      <tr class="section-row"><td colspan="${results.length+1}">Plan de Vida / Asistencia Médica</td></tr>
      <tr><td class="col-cob">Vida / Muerte Accidental</td>${results.map(r=>`<td class="col-val" style="font-size:10px">${r.cfg.vida||'N/A'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Enfermedades Graves</td>${results.map(r=>`<td class="col-val" style="font-size:10px">${r.cfg.enf_graves||'N/A'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Renta Hospitalización</td>${results.map(r=>`<td class="col-val" style="font-size:10px">${r.cfg.renta_hosp||'N/A'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Gastos de Sepelio</td>${results.map(r=>`<td class="col-val" style="font-size:10px">${r.cfg.sepelio||'N/A'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Telemedicina</td>${results.map(r=>`<td class="col-val" style="font-size:10px">${r.cfg.telemedicina||'N/A'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Beneficio Dental</td>${results.map(r=>`<td class="col-val" style="font-size:10px">${r.cfg.dental||'N/A'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Médico a Domicilio</td>${results.map(r=>`<td class="col-val" style="font-size:10px">${r.cfg.medico_dom||'N/A'}</td>`).join('')}</tr>
      <tr class="section-row"><td colspan="${results.length+1}">Deducibles</td></tr>
      <tr><td class="col-cob">Pérdida parcial</td>${results.map(r=>`<td class="col-val" style="font-size:10px">${r.cfg.ded_parcial}</td>`).join('')}</tr>
      <tr><td class="col-cob">Pérd. total daños</td>${results.map(r=>`<td class="col-val">${r.cfg.ded_daño}</td>`).join('')}</tr>
      <tr><td class="col-cob">Pérd. total robo s/disp.</td>${results.map(r=>`<td class="col-val">${r.cfg.ded_robo_sin}</td>`).join('')}</tr>
      <tr><td class="col-cob">Pérd. total robo c/disp.</td>${results.map(r=>`<td class="col-val">${r.cfg.ded_robo_con}</td>`).join('')}</tr>
      <tr class="section-row"><td colspan="${results.length+1}">Condiciones Comerciales</td></tr>
      <tr><td class="col-cob">Prima Neta Mínima</td>${results.map(r=>`<td class="col-val mono">${r.cfg.pnMin>0?fmt(r.cfg.pnMin):'—'}</td>`).join('')}</tr>
      <tr><td class="col-cob">TC — Máx. cuotas</td>${results.map(r=>`<td class="col-val mono">${r.cfg.tcMax} cuotas</td>`).join('')}</tr>
      <tr><td class="col-cob">Débito — Máx. cuotas</td>${results.map(r=>`<td class="col-val mono">${r.cfg.debMax||10} cuotas</td>`).join('')}</tr>
    </tbody></table>`;

  document.getElementById('cotizacion-resultado').style.display='block';
}

// ── Recálculo por tarjeta cuando el ejecutivo cambia la tasa ──────────────────
let _cardData = {};  // { 'ZURICH': { total, tc, deb }, ... }

function recalcCard(name, tasaPct){
  const cfg = ASEGURADORAS[name]; if(!cfg) return;
  const s   = _safeName(name);
  const tasa = parseFloat(tasaPct)/100;
  if(isNaN(tasa)||tasa<=0) return;
  const va  = parseFloat(document.getElementById('cot-va')?.value)||0;
  const ext = parseFloat(document.getElementById('cot-extras')?.value)||0;
  const vaT = va+ext;
  const cuotasTcReq  = parseInt(document.getElementById('cot-cuotas-tc')?.value)||12;
  const cuotasDebReq = parseInt(document.getElementById('cot-cuotas-deb')?.value)||10;
  const axaInc = name==='SWEADEN' ? (document.getElementById('cot-axa')?.checked||false) : false;
  const vidaIds = {LATINA:'cot-vida-latina',SWEADEN:'cot-vida-sweaden',MAPFRE:'cot-vida-mapfre',ALIANZA:'cot-vida-alianza'};
  const vida = parseFloat(document.getElementById(vidaIds[name])?.value)||0;
  const p   = calcPrima(vaT, tasa, cfg.pnMin, axaInc, vida, cfg.extraFijo||0);
  const tc  = calcCuotasTc(p.total, cfg.tcMax, cuotasTcReq, cfg.pisoTC||0);
  const deb = calcCuotasDeb(p.total, Math.min(cuotasDebReq, cfg.debMax||cuotasDebReq), cfg.pisoDeb||0);
  // Actualizar valores en el DOM
  const set=(id,v)=>{ const el=document.getElementById(id); if(el) el.textContent=v; };
  set('aseg-pn-'+s,    fmt(p.pn));
  set('aseg-sub-'+s,   fmt(p.sub));
  set('aseg-iva-'+s,   fmt(p.iva));
  set('aseg-total-'+s, fmt(p.total));
  const tcEl=document.getElementById('aseg-tc-'+s);
  if(tcEl) tcEl.innerHTML=`💳 TC ${tc.n} cuotas: <b>${fmt(tc.cuota)}/mes</b>`;
  const debEl=document.getElementById('aseg-deb-'+s);
  if(debEl) debEl.innerHTML=`🏦 Débito ${deb.n} cuotas: <b>${fmt(deb.cuota)}/mes</b>`;
  // Actualizar almacén y refrescar badge "mejor"
  _cardData[name]={total:p.total, tc, deb};
  _actualizarMejorBadge();
}

function _actualizarMejorBadge(){
  const entries=Object.entries(_cardData);
  if(!entries.length) return;
  const min=Math.min(...entries.map(([,d])=>d.total));
  entries.forEach(([name,d])=>{
    const el=document.getElementById('aseg-card-'+_safeName(name));
    if(!el) return;
    el.classList.toggle('mejor', d.total===min);
  });
}

// ── CIERRE DE VENTA ──────────────────────────────────
let cierreVentaData={};
// ══════════════════════════════════════════════════════
//  TIPO CLIENTE — NUEVO CLIENTE FORM
// ══════════════════════════════════════════════════════
function selTipoCliente(tipo, btn){
  document.querySelectorAll('#page-nuevo-cliente .pill').forEach(b=>b.classList.remove('active'));
  if(btn) btn.classList.add('active');
  document.getElementById('nc-tipo-cliente').value=tipo;
  const sec=document.getElementById('nc-produbanco-section');
  if(sec) sec.style.display=tipo==='PRODUBANCO'?'':'none';
}

let _pendingCierreClienteId=null;

// Normaliza nombre para comparación: mayúsculas, sin acentos, sin espacios dobles
function _normNombre(s){
  return (s||'').trim().toUpperCase()
    .normalize('NFD').replace(/[\u0300-\u036f]/g,'')
    .replace(/\s+/g,' ');
}

// Busca la cotización más reciente activa para un cliente.
// Retorna {cotiz, vencida} si activa; {cotiz:null, emitida} si solo hay EMITIDA.
function _cotizParaCierre(clienteId, ci, nombre){
  const all=_getCotizaciones();
  const _match=cot=>
    (clienteId && String(cot.clienteId)===String(clienteId)) ||
    (ci && ci.length>3 && cot.clienteCI===ci) ||
    _normNombre(cot.clienteNombre)===_normNombre(nombre);

  const activas=all.filter(cot=>{
    if(['REEMPLAZADA','EMITIDA'].includes(cot.estado)) return false;
    return _match(cot);
  }).sort((a,b)=>(b.fecha||'').localeCompare(a.fecha||''));

  if(activas.length){
    const cotiz=activas[0];
    const dias=Math.floor((Date.now()-new Date(cotiz.fecha||''))/(86400000));
    return {cotiz, vencida:dias>30};
  }
  // Sin activas — detectar si existe EMITIDA (para mensaje diferenciado en Caso 4)
  const emitida=all.find(cot=>cot.estado==='EMITIDA' && _match(cot))||null;
  return {cotiz:null, vencida:false, emitida};
}

// Limpia TODOS los campos del modal de cierre antes de cada apertura nueva.
// Evita que valores de un cierre anterior queden visibles en el siguiente.
function _resetCierreModal(){
  // Campos estáticos que pueden quedar sucios entre cierres
  ['cv-factura','cv-poliza','cv-observacion','cv-tipo-pago',
   'cv-vida-prima','cv-fecha-cobro-inicial']
    .forEach(id=>{ const el=document.getElementById(id); if(el) el.value=''; });
  // Campos dinámicos de Vida/AP (generados por renderCvExtras — su presencia en DOM
  // hace que renderCvExtras los "preserve" si no se limpian aquí primero)
  ['cv-poliza-vida','cv-factura-vida','cv-total-vida']
    .forEach(id=>{ const el=document.getElementById(id); if(el) el.value=''; });
  // Secciones dinámicas: limpiar innerHTML para que siempre partan de cero
  const extras=document.getElementById('cv-extras-section');
  if(extras) extras.innerHTML='';
  const panel=document.getElementById('cv-desglose-panel');
  if(panel) panel.innerHTML='';
  // Campos hidden del desglose
  ['cv-der-emision','cv-seg-camp','cv-sup-bancos','cv-axa-prima-val','cv-iva-val']
    .forEach(id=>{ const el=document.getElementById(id); if(el) el.value='0'; });
  // Ocultar fila de vida-prima (recalcDesglose la mostrará si corresponde)
  const vidaWrap=document.getElementById('cv-vida-prima-wrap');
  if(vidaWrap) vidaWrap.style.display='none';
  // Limpiar mensaje de error de factura
  const factErr=document.getElementById('cv-factura-error');
  if(factErr){ factErr.style.display='none'; factErr.textContent=''; }
}

// Pre-rellena y abre el formulario de cierre usando datos del cliente (sin cotización)
function _abrirCierreDirecto(id){
  _resetCierreModal();
  const c=DB.find(x=>String(x.id)===String(id)); if(!c) return;
  currentSegIdx=id;
  const aseg=c.aseguradora||'';
  const va=c.va||0;
  const cfg=Object.entries(ASEGURADORAS).find(([k])=>aseg.toUpperCase().includes(k));
  let total=0,pn=0,tcN=12,debN=10,tcCuota=0,debCuota=0;
  if(cfg&&va>0){
    const [asegKey,cfgObj]=cfg;
    const tasa=typeof cfgObj.tasa==='function'?cfgObj.tasa(va):_getTasaRango(asegKey,va);
    const p=calcPrima(va,tasa,cfgObj.pnMin||0);
    const tc=calcCuotasTc(p.total,cfgObj.tcMax||12,12,cfgObj.pisoTC||0);
    const deb=calcCuotasDeb(p.total,10,cfgObj.pisoDeb||0);
    total=p.total; pn=p.pn; tcN=tc.n; debN=deb.n; tcCuota=tc.cuota; debCuota=deb.cuota;
  }
  cierreVentaData={asegNombre:aseg,total,pn,cuotaTc:tcCuota,cuotaDeb:debCuota,nTc:tcN,nDeb:debN,clienteId:id,clienteNombre:c.nombre};
  document.getElementById('cv-aseg').textContent=aseg;
  document.getElementById('cv-total').textContent=total>0?`${fmt(total)} total`:'Ingrese prima manualmente';
  document.getElementById('cv-cliente').value=c.nombre;
  document.getElementById('cv-nueva-aseg').value=aseg;
  const hoyPlus1=new Date(); hoyPlus1.setFullYear(hoyPlus1.getFullYear()+1);
  document.getElementById('cv-desde').value=c.hasta||new Date().toISOString().split('T')[0];
  document.getElementById('cv-hasta').value=hoyPlus1.toISOString().split('T')[0];
  document.getElementById('cv-pn').value=pn>0?pn.toFixed(2):'';
  document.getElementById('cv-total-val').value=total>0?total.toFixed(2):'';
  if(document.getElementById('cv-cuenta')) document.getElementById('cv-cuenta').value=c.cuentaBanc||c.cuenta||'';
  if(document.getElementById('cv-axavd')) document.getElementById('cv-axavd').value=c.obs&&c.obs.includes('AXA')&&c.obs.includes('VD')?'AXA+VD':c.obs&&c.obs.includes('AXA')?'AXA':c.obs&&c.obs.includes('VD')?'VD':'';
  ['cv-factura','cv-poliza','cv-observacion'].forEach(fid=>{ const el=document.getElementById(fid); if(el) el.value=''; });
  // Pre-fill póliza anterior desde el cliente (AC del Excel)
  const polAntDirEl=document.getElementById('cv-poliza-anterior');
  if(polAntDirEl) polAntDirEl.value=c.polizaAnterior||c.polizaNueva||c.poliza||'';
  const asegAntDirEl=document.getElementById('cv-aseg-anterior');
  if(asegAntDirEl) asegAntDirEl.value=c.aseguradoraAnterior||c.aseguradora||'';
  // Pre-fill Valor Asegurado desde el cliente
  const vaDirectEl=document.getElementById('cv-va-cierre');
  const vaDirectDisplay=document.getElementById('cv-va-display');
  if(vaDirectEl) vaDirectEl.value=c.va||'';
  if(vaDirectDisplay) vaDirectDisplay.textContent=(c.va||0)>0?`VA: ${fmt(c.va)}`:'';
  document.getElementById('cv-forma-pago').value='DEBITO_BANCARIO';
  recalcDesglose();
  renderCvExtras();
  renderCvFormaPago();
  openModal('modal-cierre-venta');
}

// Callback del botón "Continuar de todos modos" en el modal de confirmación
function _continuarCierreSinCotiz(){
  closeModal('modal-confirm-cotiz');
  if(_pendingCierreClienteId) _abrirCierreDirecto(_pendingCierreClienteId);
  _pendingCierreClienteId=null;
}

// Valida cotización activa antes de abrir el cierre; enruta según el caso
function abrirCierreDesdeCliente(id, skipEstadoCheck=false){
  const c=DB.find(x=>String(x.id)===String(id)); if(!c) return;
  if(!skipEstadoCheck && !['RENOVADO','EMITIDO','EMISIÓN'].includes(c.estado)){showToast('El estado debe ser EMISIÓN, EMITIDO o RENOVADO para registrar un cierre','error');return;}
  currentSegIdx=id;

  const {cotiz, vencida, emitida}=_cotizParaCierre(c.id, c.ci, c.nombre);

  // Caso 1 — cotización activa con asegElegida y vigente → flujo ideal desde cotización
  if(cotiz && cotiz.asegElegida && !vencida){
    irAEmision(cotiz.id);
    return;
  }
  // Caso 2 — cotización activa pero sin aseguradora elegida → redirigir a Cotizaciones
  if(cotiz && !cotiz.asegElegida){
    showToast(`${c.nombre} tiene la cotización ${cotiz.codigo||''} activa. Selecciona la aseguradora en Cotizaciones para continuar.`,'info');
    showPage('cotizaciones');
    return;
  }
  // Caso 3 — cotización con asegElegida pero vencida (>30 días)
  if(cotiz && cotiz.asegElegida && vencida){
    const dias=Math.floor((Date.now()-new Date(cotiz.fecha||''))/(86400000));
    _pendingCierreClienteId=id;
    document.getElementById('confirm-cotiz-titulo').textContent='⚠️ Cotización vencida';
    document.getElementById('confirm-cotiz-msg').innerHTML=
      `La cotización <b>${cotiz.codigo||''}</b> tiene <b>${dias} días</b> (válida por 30 días).<br>
       Los precios de prima podrían haber cambiado. Se recomienda generar una nueva cotización.`;
    openModal('modal-confirm-cotiz');
    return;
  }
  // Caso 4 — sin cotización activa
  _pendingCierreClienteId=id;
  if(emitida){
    // Hay cotización pero ya fue marcada como EMITIDA (cierre previo registrado)
    document.getElementById('confirm-cotiz-titulo').textContent='⚠️ Cotización ya emitida';
    document.getElementById('confirm-cotiz-msg').innerHTML=
      `La cotización <b>${emitida.codigo||''}</b> de <b>${c.nombre}</b> ya fue marcada como EMITIDA.<br>
       Si necesitas registrar un nuevo cierre, genera una nueva cotización primero.<br>
       O puedes continuar e ingresar los datos manualmente.`;
  } else {
    document.getElementById('confirm-cotiz-titulo').textContent='⚠️ Sin cotización registrada';
    document.getElementById('confirm-cotiz-msg').innerHTML=
      `<b>${c.nombre}</b> no tiene cotización activa registrada.<br>
       Los montos de prima deberán ingresarse manualmente sin respaldo comercial.`;
  }
  openModal('modal-confirm-cotiz');
}

function abrirCierreVenta(asegNombre, total, pn, cuotaTc, cuotaDeb, nTc, nDeb){
  _resetCierreModal();
  const clienteNombre=document.getElementById('cot-nombre').value||'';
  // Buscar cliente en DB para obtener ID y cuenta
  const cMatch=DB.find(x=>x.nombre.trim().toUpperCase()===clienteNombre.trim().toUpperCase());
  cierreVentaData={asegNombre,total,pn,cuotaTc,cuotaDeb,nTc,nDeb,clienteNombre,clienteId:cMatch?String(cMatch.id):null};
  const desde=document.getElementById('cot-desde').value||'';
  // Calcular fecha de hasta (1 año desde)
  let hastaVal='';
  if(desde){const h=new Date(desde);h.setFullYear(h.getFullYear()+1);hastaVal=h.toISOString().split('T')[0];}
  document.getElementById('cv-aseg').textContent=asegNombre;
  document.getElementById('cv-total').textContent=`${fmt(total)} total`;
  document.getElementById('cv-cliente').value=clienteNombre;
  document.getElementById('cv-nueva-aseg').value=asegNombre;
  document.getElementById('cv-desde').value=desde;
  document.getElementById('cv-hasta').value=hastaVal;
  document.getElementById('cv-pn').value=pn.toFixed(2);
  document.getElementById('cv-total-val').value=total.toFixed(2);
  if(cMatch&&document.getElementById('cv-cuenta')) document.getElementById('cv-cuenta').value=cMatch.cuentaBanc||cMatch.cuenta||'';
  if(document.getElementById('cv-axavd')) document.getElementById('cv-axavd').value='';
  ['cv-factura','cv-poliza','cv-fecha-cobro-inicial','cv-observacion'].forEach(id=>{ const el=document.getElementById(id); if(el) el.value=''; });
  document.getElementById('cv-forma-pago').value='DEBITO_BANCARIO';
  renderCvFormaPago();
  renderCvExtras();
  openModal('modal-cierre-venta');
}
function getTotal(){ return parseFloat(document.getElementById('cv-total-val')?.value)||cierreVentaData.total||0; }
function recalcPrimaTotal(){
  recalcDesglose(); // backward compat alias
}

// Helper: map aseg display name → ASEGURADORAS key
function _cvGetAsegKey(){
  const asegVal=(document.getElementById('cv-nueva-aseg')?.value||'').toUpperCase();
  const knownKeys=['ZURICH','LATINA','GENERALI','ADS','SWEADEN','MAPFRE','ALIANZA'];
  for(const k of knownKeys){
    if(asegVal.includes(k)) return k;
    if(k==='ADS' && asegVal.includes('SUR')) return 'ADS';
    if(k==='GENERALI' && asegVal.includes('GENERALI')) return 'GENERALI';
  }
  return null;
}

function recalcDesglose(){
  const pn=parseFloat(document.getElementById('cv-pn')?.value)||0;
  const axavd=document.getElementById('cv-axavd')?.value||'';
  const vidaWrap=document.getElementById('cv-vida-prima-wrap');
  if(vidaWrap) vidaWrap.style.display=(axavd.includes('VD'))?'':'none';

  if(pn<=0){
    const panel=document.getElementById('cv-desglose-panel'); if(panel) panel.innerHTML='';
    const tv=document.getElementById('cv-total-val'); if(tv) tv.value='';
    return;
  }

  const asegKey=_cvGetAsegKey();
  const cfg=(asegKey&&ASEGURADORAS[asegKey])||{};
  const extraFijo=cfg.extraFijo||0;
  const axaIncluido=axavd.includes('AXA');
  const vida=parseFloat(document.getElementById('cv-vida-prima')?.value)||0;

  const der=_calcDerechosEmision(pn);
  const camp=Math.round(pn*0.005*100)/100;
  const sb=Math.round(pn*0.035*100)/100;
  const axa=axaIncluido?Math.round((60/1.15)*100)/100:0;
  const sub=Math.round((pn+der+camp+sb+axa+extraFijo)*100)/100;
  const iva=Math.round(sub*0.15*100)/100;
  const total=Math.round((sub+iva+vida)*100)/100;

  // Store in hidden fields
  const _s=(id,v)=>{const el=document.getElementById(id);if(el)el.value=v;};
  _s('cv-der-emision',der);
  _s('cv-seg-camp',camp);
  _s('cv-sup-bancos',sb);
  _s('cv-axa-prima-val',axa);
  _s('cv-iva-val',iva);
  _s('cv-total-val',total.toFixed(2));

  // Render desglose panel
  const panel=document.getElementById('cv-desglose-panel');
  if(panel){
    const row=(label,val,bold=false,muted=false)=>
      `<div style="display:flex;justify-content:space-between;padding:3px 0;font-size:12px;${bold?'font-weight:700;':''}${muted?'color:var(--muted);':''}">
        <span>${label}</span><span style="font-family:'DM Mono',monospace">${typeof val==='number'?val.toFixed(2):val}</span>
      </div>`;
    panel.innerHTML=`<div style="background:var(--warm);border:1px solid var(--border);border-radius:8px;padding:10px 14px;font-size:12px">
      ${row('Prima Neta',pn)}
      ${row('Derechos Emisión',der,false,true)}
      ${row('Seguro Campesino (0.5%)',camp,false,true)}
      ${row('Superintendencia de Bancos (3.5%)',sb,false,true)}
      ${axa>0?row('AXA Asistencia Vial',axa,false,true):''}
      ${extraFijo>0?row(`Cargo fijo ${asegKey}`,extraFijo,false,true):''}
      <div style="border-top:1px solid var(--border);margin:4px 0"></div>
      ${row('Subtotal',sub)}
      ${row('IVA 15%',iva,false,true)}
      ${vida>0?row('Prima Vida Desgravamen',vida,false,true):''}
      <div style="border-top:1px solid var(--border);margin:4px 0"></div>
      ${row('TOTAL',total,true)}
    </div>`;
  }

  cierreVentaData.total=total;
  cierreVentaData.pn=pn;
  renderCvFormaPago();
}

// Muestra sección de extras según aseguradora + axavd seleccionados:
//   SWEADEN + AXA/VD → resumen de productos incluidos (solo lectura, para que el ejecutivo vea qué vendió)
//   MAPFRE/LATINA/ALIANZA + VD → campos póliza/factura/total Vida-AP (documento separado, obligatorio)
function renderCvExtras(){
  const wrap = document.getElementById('cv-extras-section');
  if(!wrap) return;
  const aseg  = (document.getElementById('cv-nueva-aseg')?.value||'').toUpperCase();
  const axavd = document.getElementById('cv-axavd')?.value||'';
  const pnAxa  = parseFloat(document.getElementById('cv-axa-prima-val')?.value)||0;
  const pnVida = parseFloat(document.getElementById('cv-vida-prima')?.value)||0;

  const esSweaden    = aseg.includes('SWEADEN');
  const esConDocVida = aseg.includes('MAPFRE') || aseg.includes('LATINA') || aseg.includes('ALIANZA');
  const tieneAxa     = axavd.includes('AXA');
  const tieneVida    = axavd.includes('VD');

  if(esSweaden && (tieneAxa || tieneVida)){
    // Sección informativa — sin documentos separados
    const items = [];
    if(tieneAxa)  items.push(`
      <div style="display:flex;justify-content:space-between;align-items:center;
                  padding:8px 12px;background:var(--warm);border-radius:6px;font-size:13px">
        <span>🚗 <b>Auto Sustituto AXA</b>
          <span style="font-size:11px;color:var(--muted);margin-left:6px">incluido en póliza del vehículo</span>
        </span>
        <span style="font-weight:700;color:var(--green)">$${pnAxa.toFixed(2)}</span>
      </div>`);
    if(tieneVida){
      const planNombre = _getPlanVidaNombre('SWEADEN', pnVida);
      const vidaLabel  = planNombre ? `Vida ${planNombre}` : 'Vida';
      items.push(`
      <div style="display:flex;justify-content:space-between;align-items:center;
                  padding:8px 12px;background:var(--warm);border-radius:6px;font-size:13px">
        <span><b>${vidaLabel}</b>
          <span style="font-size:11px;color:var(--muted);margin-left:6px">incluido en póliza del vehículo</span>
        </span>
        <span style="font-weight:700;color:var(--green)">$${pnVida.toFixed(2)}</span>
      </div>`);
    }
    wrap.innerHTML = `
      <div style="border:2px solid var(--accent2);border-radius:8px;padding:14px;background:#f0f4ff">
        <div style="font-size:11px;font-weight:700;text-transform:uppercase;
                    letter-spacing:.6px;color:var(--accent2);margin-bottom:10px">
          📦 Productos incluidos en esta póliza
        </div>
        <div style="display:flex;flex-direction:column;gap:6px">${items.join('')}</div>
        <div style="font-size:10px;color:var(--muted);margin-top:8px">
          Costos registrados para reportería. No generan póliza ni factura adicional.
        </div>
      </div>`;
  } else if(esConDocVida && tieneVida){
    // Sección con campos obligatorios — documento separado
    const polizaVal  = document.getElementById('cv-poliza-vida')?.value  || '';
    const facturaVal = document.getElementById('cv-factura-vida')?.value || '';
    const totalVida  = document.getElementById('cv-total-vida')?.value   || '';
    wrap.innerHTML = `
      <div style="border:2px solid #e67e22;border-radius:8px;padding:14px">
        <div style="font-size:11px;font-weight:700;text-transform:uppercase;
                    letter-spacing:.6px;color:#e67e22;margin-bottom:6px">
          📋 Póliza Vida/AP — Documento Separado
        </div>
        <div style="font-size:12px;color:var(--muted);margin-bottom:12px">
          Prima referencia (cotización): <b>$${pnVida.toFixed(2)}</b>
          · La vigencia es la misma que el vehículo
        </div>
        <div class="form-grid form-grid-2" style="gap:12px">
          <div class="form-group">
            <label class="form-label">N° Póliza Vida/AP <span style="color:var(--red)">*</span></label>
            <input class="form-input" id="cv-poliza-vida" value="${polizaVal}"
                   placeholder="000000-000000" style="font-family:'DM Mono',monospace">
          </div>
          <div class="form-group">
            <label class="form-label">N° Factura Vida/AP <span style="color:var(--red)">*</span></label>
            <input class="form-input" id="cv-factura-vida" value="${facturaVal}"
                   placeholder="001-001-000000000" maxlength="17"
                   oninput="maskFactura(this)" onkeydown="return allowFacturaKey(event)"
                   style="font-family:'DM Mono',monospace;letter-spacing:1px">
          </div>
          <div class="form-group">
            <label class="form-label">Total Vida/AP <span style="color:var(--red)">*</span></label>
            <input class="form-input" id="cv-total-vida" value="${totalVida}"
                   type="number" step="0.01" placeholder="0.00">
          </div>
        </div>
      </div>`;
  } else {
    wrap.innerHTML = '';
  }
}

// Sync tipoPago dropdown → forma de pago pill
function syncTipoPagoToFormaPago(){
  const tp=(document.getElementById('cv-tipo-pago')?.value||'').toUpperCase();
  let fp='DEBITO_BANCARIO';
  if(tp.startsWith('CONTADO')) fp='CONTADO';
  else if(tp.startsWith('TC ')) fp='TARJETA_CREDITO';
  else if(tp.includes('RECURRENTES TC')) fp='DEBITO_RECURRENTE_TC';
  else if(tp.includes('CUOTA INICIAL')) fp='MIXTO';
  else if(tp.includes('CHEQUES')) fp='CONTADO';
  else if(tp.includes('DIRECTO')) fp='CONTADO';
  const fpEl=document.getElementById('cv-forma-pago');
  if(fpEl) fpEl.value=fp;
  // Update pill active state
  document.querySelectorAll('#modal-cierre-venta .pill').forEach(b=>{
    const oc=b.getAttribute('onclick')||'';
    b.classList.toggle('active',oc.includes(`'${fp}'`));
  });
  renderCvFormaPago();
}
function syncCuotasTotal(){
  const total=parseFloat(document.getElementById('cv-total-val')?.value)||0;
  cierreVentaData.total=total;
  renderCvFormaPago();
}
function renderCvFormaPago(){
  const fp=document.getElementById('cv-forma-pago').value;
  const total=getTotal();
  const {nTc,nDeb}=cierreVentaData;
  const asegNombre=(cierreVentaData.asegNombre||'').toUpperCase();
  const tcMaxAseg=(asegNombre.includes('SWEADEN')||asegNombre.includes('ALIANZA'))?9:(nTc||12);
  const wrap=document.getElementById('cv-cuotas-wrap');
  if(fp==='CONTADO'){
    wrap.innerHTML=`
    <div class="highlight-card" style="background:#d4edda;border-color:#2d6a4f">
      <div style="font-family:'DM Serif Display',serif;font-size:16px;color:var(--green)">Pago de Contado / Transferencia</div>
      <div style="font-size:13px;margin-top:4px">Monto total: <b>${fmt(total)}</b></div>
    </div>
    <div class="form-grid form-grid-2" style="gap:12px;margin-top:12px">
      <div class="form-group"><label class="form-label">Fecha de cobro <span style="color:var(--red)">*</span></label>
        <input class="form-input" type="date" id="cv-fecha-cobro-total" required>
      </div>
      <div class="form-group"><label class="form-label">Referencia / N° transferencia</label>
        <input class="form-input" id="cv-ref-transferencia" placeholder="Ref. o N° comprobante">
      </div>
    </div>`;
  } else if(fp==='DEBITO_BANCARIO'){
    // cuotas válidas: cuota >= $50
    const cuotasValidas=[1,2,3,4,5,6,7,8,9,10,11,12].filter(n=>total/n>=50||n===1);
    const cuotasOpts=cuotasValidas.map(n=>`<option value="${n}"${n===(nDeb||10)?' selected':''}>${n} cuota${n>1?'s':''} — ${fmt(total/n)}/mes</option>`).join('');
    const cuenta=document.getElementById('cv-cuenta')?.value||'';
    wrap.innerHTML=`
    <div class="form-grid form-grid-2" style="gap:12px">
      <div class="form-group"><label class="form-label">Banco <span style="color:var(--red)">*</span></label>
        <select class="form-select" id="cv-banco-deb">
          <option value="Produbanco" selected>Produbanco</option>
          <option value="Pichincha">Pichincha</option><option value="Guayaquil">Guayaquil</option>
          <option value="Pacífico">Pacífico</option><option value="Internacional">Internacional</option>
          <option value="Bolivariano">Bolivariano</option><option value="Austro">Austro</option>
          <option value="Otro">Otro</option>
        </select>
      </div>
      <div class="form-group"><label class="form-label">N° Cuenta <span style="color:var(--red)">*</span></label>
        <input class="form-input" id="cv-cuenta-deb" value="${cuenta}" placeholder="Nº de cuenta" style="font-family:'DM Mono',monospace">
      </div>
      <div class="form-group"><label class="form-label">Número de cuotas <span style="color:var(--red)">*</span></label>
        <select class="form-select" id="cv-n-cuotas" onchange="renderCvDebCalendar()">${cuotasOpts}</select>
      </div>
      <div class="form-group"><label class="form-label">Fecha de 1ª cuota <span style="color:var(--red)">*</span></label>
        <input class="form-input" type="date" id="cv-fecha-primera" onchange="renderCvDebCalendar()" required>
      </div>
    </div>
    <div id="cv-deb-calendar" style="margin-top:12px"></div>`;
    setTimeout(()=>renderCvDebCalendar(),100);
  } else if(fp==='TARJETA_CREDITO'){
    const cuotasValidas=[1,2,3,4,5,6,7,8,9,10,11,12].filter(n=>n<=tcMaxAseg&&(total/n>=50||n===1));
    const cuotasOpts=cuotasValidas.map(n=>`<option value="${n}"${n===(nTc||12)?' selected':''}>${n} cuota${n>1?'s':''} — ${fmt(total/n)}/mes</option>`).join('');
    const advertencia=(asegNombre.includes('SWEADEN')||asegNombre.includes('ALIANZA'))
      ?`<div class="highlight-card" style="margin-top:10px;background:#e8f0fb;border-color:#1a4c84"><div style="font-size:12px;color:var(--accent2)">⚠ ${cierreVentaData.asegNombre}: máximo <b>${tcMaxAseg} cuotas</b> con TC</div></div>`:'';
    wrap.innerHTML=`
    <div class="form-grid form-grid-2" style="gap:12px">
      <div class="form-group"><label class="form-label">Número de cuotas TC <span style="color:var(--red)">*</span></label>
        <select class="form-select" id="cv-n-cuotas-tc">${cuotasOpts}</select>
      </div>
      <div class="form-group"><label class="form-label">Fecha de contacto para cobro <span style="color:var(--red)">*</span></label>
        <input class="form-input" type="date" id="cv-fecha-contacto-tc" required>
      </div>
      <div class="form-group"><label class="form-label">Banco / Emisor <span style="color:var(--red)">*</span></label>
        <select class="form-select" id="cv-banco-tc">
          <option>Produbanco</option><option>Pichincha</option><option>Guayaquil</option>
          <option>Pacífico</option><option>Internacional</option><option>Bolivariano</option>
          <option>Austro</option><option>Diners</option><option>Otro</option>
        </select>
      </div>
      <div class="form-group"><label class="form-label">Últimos 4 dígitos (opcional)</label>
        <input class="form-input" id="cv-tc-digits" placeholder="XXXX" maxlength="4" style="font-family:'DM Mono',monospace">
      </div>
    </div>
    ${advertencia}`;
  } else if(fp==='DEBITO_RECURRENTE_TC'){
    // Débito recurrente mensual en TC — genera calendario igual que débito bancario
    const tpVal=document.getElementById('cv-tipo-pago')?.value||'';
    const nRec=parseInt(tpVal.split(' ')[0])||2;
    const cuotasValidas=[2,3,4,5,6,7,8,9,10,11,12].filter(n=>total/n>=50||n<=2);
    const cuotasOpts=cuotasValidas.map(n=>`<option value="${n}"${n===nRec?' selected':''}>${n} cuota${n>1?'s':''} — ${fmt(total/n)}/mes</option>`).join('');
    wrap.innerHTML=`
    <div class="highlight-card" style="background:#e8f0fb;border-color:#1a4c84;margin-bottom:10px">
      <div style="font-size:12px;color:var(--accent2)">💳 Débitos Recurrentes TC — cargo mensual automático en tarjeta de crédito</div>
    </div>
    <div class="form-grid form-grid-2" style="gap:12px">
      <div class="form-group"><label class="form-label">Banco / Emisor TC <span style="color:var(--red)">*</span></label>
        <select class="form-select" id="cv-banco-tc">
          <option>Produbanco</option><option>Pichincha</option><option>Guayaquil</option>
          <option>Pacífico</option><option>Internacional</option><option>Bolivariano</option>
          <option>Austro</option><option>Diners</option><option>Otro</option>
        </select>
      </div>
      <div class="form-group"><label class="form-label">Últimos 4 dígitos (opcional)</label>
        <input class="form-input" id="cv-tc-digits" placeholder="XXXX" maxlength="4" style="font-family:'DM Mono',monospace">
      </div>
      <div class="form-group"><label class="form-label">Número de cuotas <span style="color:var(--red)">*</span></label>
        <select class="form-select" id="cv-n-cuotas" onchange="renderCvDebCalendar()">${cuotasOpts}</select>
      </div>
      <div class="form-group"><label class="form-label">Fecha de 1ª cuota <span style="color:var(--red)">*</span></label>
        <input class="form-input" type="date" id="cv-fecha-primera" onchange="renderCvDebCalendar()" required>
      </div>
    </div>
    <div id="cv-deb-calendar" style="margin-top:12px"></div>`;
    setTimeout(()=>renderCvDebCalendar(),100);
  } else if(fp==='MIXTO'){
    // Para MIXTO: si el resto va a TC, respetar máximo de la aseguradora
    wrap.innerHTML=`
    <div class="form-grid form-grid-2" style="gap:12px">
      <div class="form-group"><label class="form-label">Cuota inicial (efectivo/transf.) <span style="color:var(--red)">*</span></label>
        <input class="form-input" type="number" step="0.01" id="cv-monto-inicial" placeholder="0.00" oninput="calcMixtoResto()">
      </div>
      <div class="form-group"><label class="form-label">Fecha cobro cuota inicial <span style="color:var(--red)">*</span></label>
        <input class="form-input" type="date" id="cv-fecha-mixto-inicial" required>
      </div>
      <div class="form-group"><label class="form-label">Resto a financiar</label>
        <input class="form-input" id="cv-mixto-resto" readonly style="background:var(--warm);font-family:'DM Mono',monospace">
      </div>
      <div class="form-group"><label class="form-label">Método del resto <span style="color:var(--red)">*</span></label>
        <select class="form-select" id="cv-mixto-metodo" onchange="calcMixtoResto()">
          <option value="DEBITO">Débito bancario</option>
          <option value="TC">Tarjeta de crédito</option>
        </select>
      </div>
      <div class="form-group"><label class="form-label">Cuotas del resto <span style="color:var(--red)">*</span></label>
        <select class="form-select" id="cv-mixto-n-cuotas">
          ${[1,2,3,4,5,6,7,8,9,10,11,12].map(n=>`<option value="${n}">${n}</option>`).join('')}
        </select>
      </div>
      <div class="form-group"><label class="form-label">Fecha 1ª cuota del resto <span style="color:var(--red)">*</span></label>
        <input class="form-input" type="date" id="cv-mixto-fecha-cuota" required>
      </div>
      <div class="form-group" id="cv-mixto-banco-wrap"><label class="form-label">Banco <span style="color:var(--red)">*</span></label>
        <select class="form-select" id="cv-mixto-banco">
          <option>Produbanco</option><option>Pichincha</option><option>Guayaquil</option>
          <option>Pacífico</option><option>Internacional</option><option>Bolivariano</option><option>Otro</option>
        </select>
      </div>
      <div class="form-group"><label class="form-label">N° Cuenta / Tarjeta</label>
        <input class="form-input" id="cv-mixto-cuenta" placeholder="Nº cuenta o últimos 4 dígitos" style="font-family:'DM Mono',monospace">
      </div>
    </div>
    <div id="cv-mixto-advertencia" style="margin-top:8px"></div>`;
    setTimeout(()=>{
      const ini=document.getElementById('cv-monto-inicial');
      if(ini&&total) ini.value=(total*0.3).toFixed(2);
      calcMixtoResto();
    },100);
  }
}
function calcMixtoResto(){
  const total=getTotal();
  let ini=parseFloat(document.getElementById('cv-monto-inicial')?.value)||0;
  // Validar no supere el total
  if(ini>total){ ini=total; const el=document.getElementById('cv-monto-inicial'); if(el) el.value=total.toFixed(2); showToast('La cuota inicial no puede superar la prima total','error'); }
  const resto=Math.max(0,total-ini);
  const elResto=document.getElementById('cv-mixto-resto');
  if(elResto) elResto.value=fmt(resto);
  // Advertencia TC para SWEADEN/ALIANZA en mixto
  const metodo=document.getElementById('cv-mixto-metodo')?.value;
  const asegNombre=(cierreVentaData.asegNombre||'').toUpperCase();
  const tcMax=(asegNombre.includes('SWEADEN')||asegNombre.includes('ALIANZA'))?9:12;
  const adv=document.getElementById('cv-mixto-advertencia');
  if(adv&&metodo==='TC'&&(asegNombre.includes('SWEADEN')||asegNombre.includes('ALIANZA'))){
    adv.innerHTML=`<div class="highlight-card" style="background:#e8f0fb;border-color:#1a4c84"><div style="font-size:12px;color:var(--accent2)">⚠ ${cierreVentaData.asegNombre}: resto a TC máximo <b>${tcMax} cuotas</b></div></div>`;
    // Limitar select de cuotas
    const sel=document.getElementById('cv-mixto-n-cuotas');
    if(sel) Array.from(sel.options).forEach(o=>{ o.disabled=parseInt(o.value)>tcMax; });
  } else if(adv){ adv.innerHTML=''; }
  // Cuota mínima $50 check
  const nCuotas=parseInt(document.getElementById('cv-mixto-n-cuotas')?.value)||1;
  if(nCuotas>1&&resto/nCuotas<50){
    const adv2=document.getElementById('cv-mixto-advertencia');
    if(adv2) adv2.innerHTML+='<div style="margin-top:4px;font-size:11px;color:var(--accent)">⚠ Cuota del resto menor a $50 — reduzca el número de cuotas</div>';
  }
}
function renderCvDebCalendar(){
  const n=parseInt(document.getElementById('cv-n-cuotas')?.value||0);
  const fechaStr=document.getElementById('cv-fecha-primera')?.value;
  const wrap=document.getElementById('cv-deb-calendar'); if(!wrap) return;
  if(!fechaStr||!n){wrap.innerHTML='';return;}
  const total=getTotal();
  const cuota=total/n;
  const fp=document.getElementById('cv-forma-pago')?.value||'';
  const esTC=fp==='DEBITO_RECURRENTE_TC';
  const banco=esTC?(document.getElementById('cv-banco-tc')?.value||'—'):(document.getElementById('cv-banco-deb')?.value||'—');
  const cuenta=esTC?('TC ****'+(document.getElementById('cv-tc-digits')?.value||'')):(document.getElementById('cv-cuenta-deb')?.value||'—');
  const labelCal=esTC?'Calendario TC recurrente':'Calendario de débitos';
  // Validar cuota mínima $50
  const aviso=cuota<50?`<div style="margin-bottom:8px;padding:6px 10px;background:#fde8e0;border-radius:6px;font-size:11px;color:var(--accent)">⚠ Cuota menor a $50 — considere reducir el número de cuotas</div>`:'';
  let html=`${aviso}<div class="section-divider">${labelCal} — ${banco} · ${cuenta}</div>
  <div style="font-size:12px;margin-bottom:8px;color:var(--muted)">${n} cuotas de <b style="color:var(--ink)">${fmt(cuota)}</b>/mes · Total: <b style="color:var(--accent)">${fmt(total)}</b></div>
  <table style="width:100%;font-size:12px"><thead><tr>
    <th style="padding:6px 10px;text-align:left;color:var(--muted);font-size:10px;background:var(--paper)">Cuota</th>
    <th style="padding:6px 10px;text-align:left;color:var(--muted);font-size:10px;background:var(--paper)">Fecha de Débito</th>
    <th style="padding:6px 10px;text-align:right;color:var(--muted);font-size:10px;background:var(--paper)">Monto</th>
    <th style="padding:6px 10px;text-align:center;color:var(--muted);font-size:10px;background:var(--paper)">Estado</th>
  </tr></thead><tbody>`;
  const base=new Date(fechaStr);
  const today=new Date(); today.setHours(0,0,0,0);
  for(let i=0;i<n;i++){
    const d=new Date(base); d.setMonth(d.getMonth()+i);
    const isPast=d<today;
    const isToday=d.toDateString()===today.toDateString();
    const rowStyle=isPast?'background:#f8f9fa;':isToday?'background:#e8f4fd;':'';
    const estado=isPast?'<span class="badge badge-gray" style="font-size:9px">Pendiente</span>':isToday?'<span class="badge badge-blue" style="font-size:9px">Hoy</span>':'<span style="font-size:10px;color:var(--muted)">—</span>';
    html+=`<tr style="border-bottom:1px solid var(--warm);${rowStyle}">
      <td style="padding:7px 10px;font-weight:600">${i+1}</td>
      <td style="padding:7px 10px;font-family:'DM Mono',monospace">${d.toISOString().split('T')[0]}</td>
      <td style="padding:7px 10px;text-align:right;font-family:'DM Mono',monospace;font-weight:600">${fmt(cuota)}</td>
      <td style="padding:7px 10px;text-align:center">${estado}</td>
    </tr>`;
  }
  html+=`</tbody></table>`;
  wrap.innerHTML=html;
}
function guardarCierreVenta(){
  // Validaciones obligatorias
  const factura=document.getElementById('cv-factura').value.trim();
  const poliza=document.getElementById('cv-poliza').value.trim();
  const desde=document.getElementById('cv-desde').value;
  const hasta=document.getElementById('cv-hasta').value;
  const aseg=document.getElementById('cv-nueva-aseg').value.trim();
  const fp=document.getElementById('cv-forma-pago').value;
  const _cvClienteRaw=(document.getElementById('cv-cliente').value.trim())||cierreVentaData.clienteNombre||'';
  const obs=document.getElementById('cv-observacion').value.trim();
  const errors=[];
  const totalVal=parseFloat(document.getElementById('cv-total-val')?.value)||0;
  if(!factura) errors.push('N° de Factura');
  else if(!validarFactura(factura)) errors.push('N° de Factura (formato inválido — debe ser 001-001-000000000)');
  if(!poliza) errors.push('N° de Póliza');
  if(!desde) errors.push('Vigencia Desde');
  if(!hasta) errors.push('Vigencia Hasta');
  if(!aseg) errors.push('Aseguradora');
  if(!fp) errors.push('Forma de pago');
  if(totalVal<=0) errors.push('Prima Total (debe ser mayor a 0)');
  // Validar póliza Vida/AP adicional (MAPFRE/LATINA/ALIANZA con VD)
  const _asegUpper = aseg.toUpperCase();
  const _axavd = document.getElementById('cv-axavd')?.value||'';
  const _esConDocVida = (_asegUpper.includes('MAPFRE')||_asegUpper.includes('LATINA')||_asegUpper.includes('ALIANZA')) && _axavd.includes('VD');
  if(_esConDocVida){
    if(!(document.getElementById('cv-poliza-vida')?.value||'').trim()) errors.push('N° Póliza Vida/AP');
    const _factVida=(document.getElementById('cv-factura-vida')?.value||'').trim();
    if(!_factVida) errors.push('N° Factura Vida/AP');
    else if(!validarFactura(_factVida)) errors.push('N° Factura Vida/AP (formato inválido — debe ser 001-001-000000000)');
    if(!(parseFloat(document.getElementById('cv-total-vida')?.value)||0)) errors.push('Total Vida/AP');
  }
  // Validar campos de forma de pago
  if(fp==='CONTADO'){
    if(!document.getElementById('cv-fecha-cobro-total')?.value) errors.push('Fecha de cobro de contado');
  } else if(fp==='DEBITO_BANCARIO'){
    if(!document.getElementById('cv-n-cuotas')?.value) errors.push('N° de cuotas');
    if(!document.getElementById('cv-fecha-primera')?.value) errors.push('Fecha 1ª cuota');
  } else if(fp==='TARJETA_CREDITO'){
    if(!document.getElementById('cv-n-cuotas-tc')?.value) errors.push('N° cuotas TC');
    if(!document.getElementById('cv-fecha-contacto-tc')?.value) errors.push('Fecha contacto TC');
  } else if(fp==='DEBITO_RECURRENTE_TC'){
    if(!document.getElementById('cv-n-cuotas')?.value) errors.push('N° de cuotas TC recurrentes');
    if(!document.getElementById('cv-fecha-primera')?.value) errors.push('Fecha 1ª cuota TC recurrente');
  } else if(fp==='MIXTO'){
    if(!document.getElementById('cv-monto-inicial')?.value) errors.push('Monto cuota inicial');
    if(!document.getElementById('cv-fecha-mixto-inicial')?.value) errors.push('Fecha cuota inicial');
    if(!document.getElementById('cv-mixto-fecha-cuota')?.value) errors.push('Fecha 1ª cuota resto');
  }
  if(errors.length){showToast('Campos obligatorios: '+errors.join(', '),'error');return;}

  // Armar registro de cierre
  const total=getTotal();
  const pago={forma:fp};
  if(fp==='CONTADO'){
    pago.fechaCobro=document.getElementById('cv-fecha-cobro-total').value;
    pago.referencia=document.getElementById('cv-ref-transferencia')?.value||'';
  } else if(fp==='DEBITO_BANCARIO'){
    const n=parseInt(document.getElementById('cv-n-cuotas').value);
    const fecha=document.getElementById('cv-fecha-primera').value;
    pago.banco=document.getElementById('cv-banco-deb')?.value||'';
    pago.cuenta=document.getElementById('cv-cuenta-deb')?.value||document.getElementById('cv-cuenta')?.value||'';
    pago.nCuotas=n; pago.fechaPrimera=fecha;
    pago.cuotaMonto=(total/n).toFixed(2);
    pago.calendario=Array.from({length:n},(_,i)=>{
      const d=new Date(fecha); d.setMonth(d.getMonth()+i);
      return d.toISOString().split('T')[0];
    });
  } else if(fp==='TARJETA_CREDITO'){
    pago.nCuotas=parseInt(document.getElementById('cv-n-cuotas-tc').value);
    pago.fechaContacto=document.getElementById('cv-fecha-contacto-tc').value;
    pago.banco=document.getElementById('cv-banco-tc').value;
    pago.digitos=document.getElementById('cv-tc-digits').value;
    pago.cuotaMonto=(total/pago.nCuotas).toFixed(2);
  } else if(fp==='DEBITO_RECURRENTE_TC'){
    const n=parseInt(document.getElementById('cv-n-cuotas').value);
    const fecha=document.getElementById('cv-fecha-primera').value;
    pago.banco=document.getElementById('cv-banco-tc')?.value||'';
    pago.digitos=document.getElementById('cv-tc-digits')?.value||'';
    pago.nCuotas=n; pago.fechaPrimera=fecha;
    pago.cuotaMonto=(total/n).toFixed(2);
    pago.calendario=Array.from({length:n},(_,i)=>{
      const d=new Date(fecha); d.setMonth(d.getMonth()+i);
      return d.toISOString().split('T')[0];
    });
  } else if(fp==='MIXTO'){
    pago.montoInicial=document.getElementById('cv-monto-inicial').value;
    pago.fechaInicial=document.getElementById('cv-fecha-mixto-inicial').value;
    pago.metodoResto=document.getElementById('cv-mixto-metodo').value;
    pago.nCuotasResto=document.getElementById('cv-mixto-n-cuotas').value;
    pago.fechaCuotaResto=document.getElementById('cv-mixto-fecha-cuota').value;
    pago.banco=document.getElementById('cv-mixto-banco')?.value||'';
    pago.cuenta=document.getElementById('cv-mixto-cuenta')?.value||'';
  }
  const axavd=document.getElementById('cv-axavd')?.value||'';

  // Buscar cliente en DB
  const c=cierreVentaData.clienteId ? DB.find(x=>String(x.id)===String(cierreVentaData.clienteId)) : DB.find(x=>x.nombre.trim().toUpperCase()===_cvClienteRaw.trim().toUpperCase());
  // Resolver clienteNombre final con fallback a DB
  const clienteNombre = _cvClienteRaw || (c ? c.nombre : '');

  // Armar registro de cierre
  // _clienteId y _placa usan prefijo _ para ser solo locales (no se envían a SharePoint)
  const _g=(id,def=0)=>parseFloat(document.getElementById(id)?.value)||def;
  const _gs=(id,def='')=>document.getElementById(id)?.value||def;
  const cierre={
    id: Date.now(),
    fechaRegistro: new Date().toISOString().split('T')[0],
    clienteNombre, aseguradora:aseg, polizaNueva:poliza,
    facturaAseg:factura, primaTotal:total,
    primaNeta:_g('cv-pn')||cierreVentaData.pn||0,
    vigDesde:desde, vigHasta:hasta,
    formaPago:pago, observacion:obs,
    axavd, cuenta:_gs('cv-cuenta'),
    ejecutivo:currentUser?currentUser.id:'',
    clienteId: c ? String(c.id) : '',
    cotizacionId: cierreVentaData.fromCotizacion||'',
    // Desglose de prima (Fase 2)
    derechosEmision: _g('cv-der-emision'),
    segCampesino:    _g('cv-seg-camp'),
    supBancos:       _g('cv-sup-bancos'),
    iva:             _g('cv-iva-val'),
    vidaPrima:       _g('cv-vida-prima'),
    axaPrima:        _g('cv-axa-prima-val'),
    // Extras Vida/AP
    poliza_vida:     _gs('cv-poliza-vida'),
    factura_vida:    _gs('cv-factura-vida'),
    total_vida:      _g('cv-total-vida'),
    // Pago detallado
    tipoPago:        _gs('cv-tipo-pago'),
    polizaAnterior:  _gs('cv-poliza-anterior'),
    asegAnterior:    _gs('cv-aseg-anterior'),
    tasaAplicada:    cierreVentaData.tasa||0,
    valorAsegurado:  parseFloat(document.getElementById('cv-va-cierre')?.value)||0,
    _clienteId: c ? String(c.id) : '',
    _placa: c?.placa||'',
  };

  const allCierres=_getCierres();

  if(cierreVentaData.editandoCierreId){
    // MODO EDICIÓN — actualizar cierre existente preservando id y fecha original
    const idx=allCierres.findIndex(x=>String(x.id)===String(cierreVentaData.editandoCierreId));
    if(idx>=0){
      allCierres[idx]={...allCierres[idx], ...cierre,
        id:cierreVentaData.editandoCierreId,
        fechaRegistro:allCierres[idx].fechaRegistro // preservar fecha original
      };
    }
    _saveCierres(allCierres);
    showToast(`✓ Cierre actualizado — ${aseg} — Póliza ${poliza}`,'success');
  } else {
    // MODO NUEVO — validar duplicados antes de guardar

    // 1. Duplicado por póliza (una póliza no puede registrarse dos veces)
    const dupPoliza = allCierres.find(x => x.polizaNueva && x.polizaNueva.trim().toLowerCase() === poliza.trim().toLowerCase());
    if(dupPoliza){
      showToast(`Ya existe un cierre con la póliza ${poliza} registrado el ${dupPoliza.fechaRegistro} para ${dupPoliza.clienteNombre}`, 'error');
      return;
    }

    // Calcular comisión estimada para este cierre
    const _comisiones = _getComisiones();
    cierre.comisionPct = _comisiones[aseg] || 0;
    cierre.comision    = Math.round((cierre.primaNeta||0) * (cierre.comisionPct/100) * 100) / 100;

    cierre._dirty = true;
    allCierres.push(cierre);
    _saveCierres(allCierres);
    showToast(`✓ Venta cerrada — ${aseg} — Póliza ${poliza}`,'success');

    // Actualizar cliente en DB solo al registrar un cierre nuevo (no al editar)
    if(c){
      c._dirty = true;
      // Preservar póliza vigente como anterior (AC del Excel) antes de sobreescribir con la nueva
      if(c.polizaNueva) c.polizaAnterior=c.polizaNueva;
      c.polizaNueva=poliza; c.factura=factura; c.aseguradora=aseg;
      c.desde=desde; c.hasta=hasta; c.formaPago=pago;
      c.primaTotal=total; c.axavd=axavd;
      // Si venía de EMISIÓN, registrar EMITIDO en bitácora como paso intermedio (trazabilidad completa)
      if(c.estado==='EMISIÓN') _bitacoraAdd(c,'Póliza recibida de aseguradora — emitida','sistema');
      c.estado='RENOVADO'; _bitacoraAdd(c, `Cierre registrado${obs?' — '+obs:''}. Aseg: ${cierre?.aseguradora||''}`, 'cierre');
      c.ultimoContacto=new Date().toISOString().split('T')[0];
      saveDB();
      sincronizarCotizPorCliente(c.id, c.nombre, c.ci, 'RENOVADO');
    }

    // Marcar cotización de origen como EMITIDA (independiente del estado actual)
    if(cierreVentaData.fromCotizacion){
      const allCotiz=_getCotizaciones();
      const ci=allCotiz.findIndex(x=>String(x.id)===String(cierreVentaData.fromCotizacion));
      if(ci>=0&&!['EMITIDA','REEMPLAZADA'].includes(allCotiz[ci].estado)){
        allCotiz[ci].estado='EMITIDA';
        allCotiz[ci].fechaAcept=allCotiz[ci].fechaAcept||new Date().toISOString().split('T')[0];
        allCotiz[ci]._dirty=true;
        _saveCotizaciones(allCotiz);
      }
    }
  }

  // Reset modo edición y botón
  cierreVentaData.editandoCierreId=null;
  const btnG=document.querySelector('#modal-cierre-venta .btn-green[onclick="guardarCierreVenta()"]');
  if(btnG){ btnG.textContent='✓ Registrar Cierre'; btnG.style.background=''; }

  closeModal('modal-cierre-venta');
  renderDashboard();
  renderCierres();
  renderCotizaciones();
  actualizarBadgeCotizaciones();
  actualizarBadgeCobranza();
}
// ── WhatsApp directo desde resultado de cotización ──────────────────
function enviarWhatsAppCotiz(aseg, total, cuota, nCuotas){
  const nombre  = document.getElementById('cot-nombre')?.value || '';
  const celular = (document.getElementById('cot-cel')?.value||'').replace(/\D/g,'');
  const marca   = document.getElementById('cot-marca')?.value||'';
  const modelo  = document.getElementById('cot-modelo')?.value||'';
  const anio    = document.getElementById('cot-anio')?.value||'';
  if(!celular){ showToast('Ingrese el celular del cliente primero','error'); return; }
  const phone = celular.startsWith('593') ? celular : `593${celular.replace(/^0/,'')}`;
  const msg = encodeURIComponent(
    `Estimado/a ${nombre}, adjunto cotización para el seguro de su vehículo ${marca} ${modelo} ${anio}.\n\n` +
    `✅ *${aseg}* — Total: *$${total.toFixed(2)}*\n` +
    `💳 ${nCuotas} cuotas de $${cuota.toFixed(2)}/mes\n\n` +
    `Cobertura total, asistencia 24/7. ¿Desea proceder?\n— Reliance Broker de Seguros`
  );
  window.open(`https://web.whatsapp.com/send?phone=${phone}&text=${msg}`, '_blank');
}

function printOneAseg(name,total,pn,cuotaTc,cuotaDeb,nTc,nDeb){
  const cfg=ASEGURADORAS[name];
  const nombre=document.getElementById('cot-nombre').value||'—';
  const ci=document.getElementById('cot-ci').value||'—';
  const marca=document.getElementById('cot-marca').value;
  const modelo=document.getElementById('cot-modelo').value;
  const anio=document.getElementById('cot-anio').value;
  const placa=document.getElementById('cot-placa').value||'—';
  const va=document.getElementById('cot-va').value;
  const desde=document.getElementById('cot-desde').value||'—';
  document.getElementById('print-ejecutivo').textContent=currentUser?currentUser.name:'—';
  document.getElementById('print-fecha').textContent='Fecha: '+new Date().toLocaleDateString('es-EC');
  document.getElementById('print-content').innerHTML=`
    <h2 style="font-size:20px;margin-bottom:4px;font-family:'DM Serif Display',serif">COTIZACIÓN DE SEGURO VEHICULAR</h2>
    <div style="font-size:12px;color:#7a7060;margin-bottom:20px">Válido por 30 días · Sujeto a aceptación de la aseguradora</div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:24px;margin-bottom:24px">
      <div style="padding:16px;border:1px solid #d4cbb8;border-radius:8px">
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#7a7060;margin-bottom:10px">Datos del Cliente</div>
        <div style="font-size:13px;margin-bottom:4px"><b>${nombre}</b></div>
        <div style="font-size:12px;color:#7a7060">CI: ${ci}</div>
      </div>
      <div style="padding:16px;border:1px solid #d4cbb8;border-radius:8px">
        <div style="font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:1px;color:#7a7060;margin-bottom:10px">Datos del Vehículo</div>
        <div style="font-size:13px;margin-bottom:4px"><b>${marca} ${modelo} ${anio}</b></div>
        <div style="font-size:12px;color:#7a7060">Placa: ${placa} · VA: $${Number(va).toLocaleString('es-EC')}</div>
      </div>
    </div>
    <div style="padding:20px;border:2px solid ${cfg.color};border-radius:10px;margin-bottom:20px">
      <div style="display:flex;justify-content:space-between;align-items:center;margin-bottom:16px;padding-bottom:12px;border-bottom:1px solid #d4cbb8">
        <div style="font-size:22px;font-weight:700;color:${cfg.color}">${name}</div>
        <div style="text-align:right">
          <div style="font-size:11px;color:#7a7060">COSTO TOTAL</div>
          <div style="font-size:28px;font-weight:700;color:${cfg.color}">$${Number(total).toLocaleString('es-EC',{minimumFractionDigits:2})}</div>
        </div>
      </div>
      <div style="display:grid;grid-template-columns:1fr 1fr;gap:12px;font-size:12px">
        <div><b>Prima Neta:</b> $${Number(pn).toLocaleString('es-EC',{minimumFractionDigits:2})}</div>
        <div><b>Vigencia desde:</b> ${desde}</div>
        <div><b>TC ${nTc} cuotas:</b> $${Number(cuotaTc).toLocaleString('es-EC',{minimumFractionDigits:2})}/mes</div>
        <div><b>Débito ${nDeb} cuotas:</b> $${Number(cuotaDeb).toLocaleString('es-EC',{minimumFractionDigits:2})}/mes</div>
      </div>
    </div>
    <div style="font-size:11px;margin-bottom:8px;font-weight:700;text-transform:uppercase;letter-spacing:.8px;color:#7a7060">Coberturas y Amparos Adicionales</div>
    <table style="width:100%;border-collapse:collapse;font-size:12px">
      <tr style="background:#f5f0e8"><td style="padding:8px;border:1px solid #d4cbb8">Responsabilidad Civil</td><td style="padding:8px;border:1px solid #d4cbb8;font-weight:600">$${Number(cfg.resp_civil).toLocaleString('es-EC')}</td></tr>
      <tr><td style="padding:8px;border:1px solid #d4cbb8">Muerte acc. por ocupante</td><td style="padding:8px;border:1px solid #d4cbb8;font-weight:600">$${Number(cfg.muerte_ocupante).toLocaleString('es-EC')}</td></tr>
      <tr style="background:#f5f0e8"><td style="padding:8px;border:1px solid #d4cbb8">Gastos Médicos/ocupante</td><td style="padding:8px;border:1px solid #d4cbb8;font-weight:600">$${Number(cfg.gastos_medicos).toLocaleString('es-EC')}</td></tr>
      <tr><td style="padding:8px;border:1px solid #d4cbb8">Amparo Patrimonial</td><td style="padding:8px;border:1px solid #d4cbb8;font-weight:600">${cfg.amparo}</td></tr>
      <tr style="background:#f5f0e8"><td style="padding:8px;border:1px solid #d4cbb8">Auto Sustituto</td><td style="padding:8px;border:1px solid #d4cbb8;font-weight:600">${cfg.auto_sust}</td></tr>
      <tr><td style="padding:8px;border:1px solid #d4cbb8">Ded. Pérdida Parcial</td><td style="padding:8px;border:1px solid #d4cbb8;font-weight:600">${cfg.ded_parcial}</td></tr>
      <tr style="background:#f5f0e8"><td style="padding:8px;border:1px solid #d4cbb8">Ded. Pérdida Total Daños</td><td style="padding:8px;border:1px solid #d4cbb8;font-weight:600">${cfg.ded_daño}</td></tr>
    </table>`;
  document.getElementById('print-area').style.display='block';
  window.print();
  document.getElementById('print-area').style.display='none';
}
function printCotizacion(){
  // Leer datos del cotizador
  const nombre   = document.getElementById('cot-nombre')?.value || '—';
  const ci       = document.getElementById('cot-ci')?.value     || '—';
  const marca    = document.getElementById('cot-marca')?.value  || '';
  const modelo   = document.getElementById('cot-modelo')?.value || '';
  const anio     = document.getElementById('cot-anio')?.value   || '';
  const placa    = document.getElementById('cot-placa')?.value  || '—';
  const va       = parseFloat(document.getElementById('cot-va')?.value)||0;
  const extras   = parseFloat(document.getElementById('cot-extras')?.value)||0;
  const vaT      = va + extras;
  const desde    = document.getElementById('cot-desde')?.value  || '—';
  const cuotasTcReq  = parseInt(document.getElementById('cot-cuotas-tc')?.value)||12;
  const cuotasDebReq = parseInt(document.getElementById('cot-cuotas-deb')?.value)||10;
  // Leer mismos toggles que calcCotizacion para garantizar precios idénticos al PDF
  const axaActivo = document.getElementById('cot-axa')?.checked || false;
  const vidaInputs = {
    LATINA:  parseFloat(document.getElementById('cot-vida-latina')?.value)||0,
    SWEADEN: parseFloat(document.getElementById('cot-vida-sweaden')?.value)||0,
    MAPFRE:  parseFloat(document.getElementById('cot-vida-mapfre')?.value)||0,
    ALIANZA: parseFloat(document.getElementById('cot-vida-alianza')?.value)||0,
  };
  const exec     = currentUser ? currentUser.name : '—';
  const fecha    = new Date().toLocaleDateString('es-EC',{day:'2-digit',month:'long',year:'numeric'});

  // Calcular resultados solo de las aseguradoras seleccionadas
  const selected = getSelectedAseg();
  if(!selected.length){showToast('Selecciona al menos una aseguradora','error');return;}

  const results = selected.map(name=>{
    const cfg = ASEGURADORAS[name];
    // Tasa: lee del input editable en la tarjeta (si fue modificado); sino usa default guardado
    const tasa = typeof cfg.tasa === 'function' ? cfg.tasa(vaT) : _getTasaFromCard(name);
    const axaInc = name === 'SWEADEN' ? axaActivo : false;
    const vida   = vidaInputs[name] || 0;
    const p = calcPrima(vaT, tasa, cfg.pnMin, axaInc, vida, cfg.extraFijo||0);
    const tc  = calcCuotasTc(p.total, cfg.tcMax, cuotasTcReq, cfg.pisoTC||0);
    const deb = calcCuotasDeb(p.total, Math.min(cuotasDebReq, cfg.debMax||cuotasDebReq), cfg.pisoDeb||0);
    const planVida = _getPlanVidaNombre(name, vida);
    return {name,cfg,tasa,...p,tc,deb,planVida};
  });
  const minTotal = Math.min(...results.map(r=>r.total));

  // ── Construir HTML del PDF ──
  const colW = Math.min(220, Math.floor(680/results.length));

  const html = `
  <style>
    @page{ size:A4; margin:18mm 15mm 18mm 15mm; }
    *{ box-sizing:border-box; font-family:'DM Sans',Arial,sans-serif; }
    body{ color:#1a1a1a; font-size:11px; }
    .header{ display:flex; justify-content:space-between; align-items:center;
             border-bottom:3px solid #c84b1a; padding-bottom:10px; margin-bottom:14px; }
    .logo-reliance{ font-family:serif; font-size:22px; font-weight:700; color:#c84b1a; letter-spacing:1px; }
    .header-right{ text-align:right; font-size:10px; color:#666; line-height:1.7; }
    .title{ font-size:15px; font-weight:700; color:#1a1a1a; margin-bottom:2px; }
    .subtitle{ font-size:10px; color:#888; margin-bottom:14px; }
    .datos-grid{ display:grid; grid-template-columns:1fr 1fr; gap:10px; margin-bottom:14px; }
    .datos-box{ border:1px solid #ddd; border-radius:6px; padding:10px 12px; }
    .datos-box-title{ font-size:9px; font-weight:700; text-transform:uppercase;
                      letter-spacing:.8px; color:#888; margin-bottom:6px; }
    .datos-box-main{ font-size:12px; font-weight:700; margin-bottom:2px; }
    .datos-box-sub{ font-size:10px; color:#666; }
    /* Tabla comparativa */
    table.comp{ width:100%; border-collapse:collapse; font-size:10px; margin-bottom:14px; }
    table.comp th{ background:#1a1a1a; color:#fff; padding:6px 8px; text-align:center; font-size:10px; }
    table.comp th.col-cob{ text-align:left; width:130px; }
    table.comp th.mejor-th{ background:#2d6a4f; }
    table.comp td{ padding:5px 8px; border-bottom:1px solid #eee; text-align:center; }
    table.comp td.col-cob{ text-align:left; color:#555; font-size:10px; }
    table.comp tr.section td{ background:#f5f5f5; font-weight:700; font-size:9px;
                               text-transform:uppercase; letter-spacing:.6px; color:#888;
                               padding:4px 8px; }
    table.comp tr:nth-child(even) td:not(.col-cob){ background:#fafafa; }
    .total-row td{ font-weight:700; font-size:12px; padding:8px; }
    .cuotas-row td{ font-size:10px; color:#444; padding:3px 8px; }
    .mejor-col{ color:#2d6a4f; font-weight:700; }
    .mejor-badge{ background:#2d6a4f; color:#fff; font-size:8px; border-radius:3px;
                  padding:1px 5px; display:inline-block; margin-bottom:2px; }
    .footer{ margin-top:14px; padding-top:10px; border-top:1px solid #ddd;
             font-size:9px; color:#aaa; display:flex; justify-content:space-between; }
    .no-print-note{ font-size:9px; color:#888; font-style:italic; margin-bottom:10px; }
    .aseg-color-bar{ height:3px; border-radius:0; }
    .sweaden-extra{ font-size:9px; color:#1a4c84; font-weight:600; }
  </style>

  <div class="header">
    <div>
      <div class="logo-reliance">RELIANCE</div>
      <div style="font-size:9px;color:#888;margin-top:2px">Asesores Productores de Seguros</div>
    </div>
    <div class="header-right">
      <div>Ejecutivo: <b>${exec}</b></div>
      <div>Fecha: <b>${fecha}</b></div>
      <div>Válida por <b>30 días</b></div>
    </div>
  </div>

  <div class="title">COTIZACIÓN DE SEGURO VEHICULAR</div>
  <div class="subtitle">Comparativo ${results.length} aseguradoras · Sujeto a aceptación</div>

  <div class="datos-grid">
    <div class="datos-box">
      <div class="datos-box-title">Datos del Cliente</div>
      <div class="datos-box-main">${nombre}</div>
      <div class="datos-box-sub">CI/RUC: ${ci}</div>
    </div>
    <div class="datos-box">
      <div class="datos-box-title">Datos del Vehículo</div>
      <div class="datos-box-main">${marca} ${modelo} ${anio}</div>
      <div class="datos-box-sub">Placa: ${placa} &nbsp;·&nbsp; Valor Asegurado: <b>$${vaT.toLocaleString('es-EC',{minimumFractionDigits:2})}</b></div>
    </div>
  </div>

  <table class="comp">
    <thead>
      <tr>
        <th class="col-cob">Cobertura</th>
        ${results.map(r=>`
          <th class="${r.total===minTotal?'mejor-th':''}" style="color:${r.total===minTotal?'#fff':'#fff'}">
            ${r.total===minTotal?'<div class="mejor-badge">✓ MEJOR</div><br>':''}
            ${r.name}
          </th>`).join('')}
      </tr>
    </thead>
    <tbody>
      <tr class="section"><td class="col-cob" colspan="${results.length+1}">Prima y Costos</td></tr>
      <tr>
        <td class="col-cob">Tasa</td>
        ${results.map(r=>`<td>${(r.tasa*100).toFixed(2)}%</td>`).join('')}
      </tr>
      <tr>
        <td class="col-cob">Prima Neta</td>
        ${results.map(r=>`<td>$${r.pn.toFixed(2)}</td>`).join('')}
      </tr>
      <tr>
        <td class="col-cob">Der. Emisión + Imp.</td>
        ${results.map(r=>`<td>$${(r.der+r.camp+r.sb).toFixed(2)}</td>`).join('')}
      </tr>
      <tr>
        <td class="col-cob">IVA 15%</td>
        ${results.map(r=>`<td>$${r.iva.toFixed(2)}</td>`).join('')}
      </tr>
      ${results.some(r=>r.axa>0)?`
      <tr>
        <td class="col-cob">🚗 Auto sustituto AXA</td>
        ${results.map(r=>`<td class="sweaden-extra">${r.axa>0?'$'+r.axa.toFixed(2):'—'}</td>`).join('')}
      </tr>`:''}
      ${results.some(r=>r.vida>0)?`
      <tr style="background:#e8f5e9">
        <td class="col-cob" style="color:#2d6a4f">❤ Plan de Vida</td>
        ${results.map(r=>`<td style="color:#2d6a4f;font-size:9px">${r.planVida||'—'}</td>`).join('')}
      </tr>`:''}
      <tr class="total-row">
        <td class="col-cob" style="font-weight:700">TOTAL</td>
        ${results.map(r=>`<td class="${r.total===minTotal?'mejor-col':''}" style="font-size:13px;color:${r.cfg.color}">$${r.total.toFixed(2)}</td>`).join('')}
      </tr>
      <tr class="cuotas-row">
        <td class="col-cob">💳 TC ${cuotasTcReq} cuotas</td>
        ${results.map(r=>`<td>${r.tc.n} × <b>$${r.tc.cuota.toFixed(2)}</b>/mes</td>`).join('')}
      </tr>
      <tr class="cuotas-row">
        <td class="col-cob">🏦 Débito ${cuotasDebReq} cuotas</td>
        ${results.map(r=>`<td>${r.deb.n} × <b>$${r.deb.cuota.toFixed(2)}</b>/mes</td>`).join('')}
      </tr>

      <tr class="section"><td class="col-cob" colspan="${results.length+1}">Amparos Adicionales</td></tr>
      <tr><td class="col-cob">Resp. Civil</td>${results.map(r=>`<td>$${r.cfg.resp_civil.toLocaleString()}</td>`).join('')}</tr>
      <tr><td class="col-cob">Muerte acc./ocupante</td>${results.map(r=>`<td>$${r.cfg.muerte_ocupante.toLocaleString()}</td>`).join('')}</tr>
      <tr><td class="col-cob">Muerte acc./titular</td>${results.map(r=>`<td>${r.cfg.muerte_titular?'$'+r.cfg.muerte_titular.toLocaleString():'N/A'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Gastos Médicos</td>${results.map(r=>`<td>$${r.cfg.gastos_medicos.toLocaleString()}</td>`).join('')}</tr>
      <tr><td class="col-cob">Amparo Patrimonial</td>${results.map(r=>`<td style="font-size:9px">${r.cfg.amparo}</td>`).join('')}</tr>

      <tr class="section"><td class="col-cob" colspan="${results.length+1}">Beneficios</td></tr>
      <tr><td class="col-cob">Auto Sustituto</td>${results.map(r=>`<td style="font-size:8px">${r.cfg.auto_sust}</td>`).join('')}</tr>
      <tr><td class="col-cob">Asist. Legal</td>${results.map(r=>`<td style="color:${(r.cfg.legal||'SÍ')==='SÍ'?'#2d6a4f':'#c84b1a'};font-weight:600">${r.cfg.legal||'SÍ'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Asist. Exequial</td>${results.map(r=>`<td style="color:${r.cfg.exequial==='SÍ'?'#2d6a4f':'#c84b1a'};font-weight:600">${r.cfg.exequial}</td>`).join('')}</tr>

      <tr class="section"><td class="col-cob" colspan="${results.length+1}">Deducibles</td></tr>
      <tr><td class="col-cob">Pérd. Parcial</td>${results.map(r=>`<td style="font-size:9px">${r.cfg.ded_parcial}</td>`).join('')}</tr>
      <tr><td class="col-cob">Pérd. Total Daños</td>${results.map(r=>`<td>${r.cfg.ded_daño}</td>`).join('')}</tr>
      <tr><td class="col-cob">Pérd. Total Robo</td>${results.map(r=>`<td>${r.cfg.ded_robo_sin}</td>`).join('')}</tr>
    </tbody>
  </table>

  ${results.some(r=>r.vida>0)?`
  <div style="margin-top:16px;border:2px solid #2d6a4f;border-radius:6px;overflow:hidden;page-break-inside:avoid">
    <div style="background:#2d6a4f;color:#fff;padding:7px 12px;font-size:10px;font-weight:700;text-transform:uppercase;letter-spacing:.6px">
      ❤ Beneficios incluidos en esta cotización
    </div>
    <div style="padding:10px 12px;display:flex;flex-wrap:wrap;gap:14px;background:#f0faf4">
      ${results.filter(r=>r.vida>0&&r.planVida).map(r=>{
        const plan = PLANES_VIDA[r.name]?.[r.planVida];
        if(!plan) return '';
        return `<div style="flex:1;min-width:180px">
          <div style="font-weight:700;font-size:11px;color:${r.cfg.color};margin-bottom:5px;border-bottom:1px solid #c8e6c9;padding-bottom:3px">
            ${r.name} — ${r.planVida}
          </div>
          <table style="font-size:9px;width:100%;border-collapse:collapse">
            ${plan.coberturas.map((c,i)=>`
              <tr style="${i%2===0?'background:#e8f5e9':''}">
                <td style="padding:2px 4px;color:#3a5a40">${c.concepto}</td>
                <td style="padding:2px 4px;font-weight:600;color:#1a3a24;text-align:right;white-space:nowrap">${c.valor}</td>
              </tr>`).join('')}
          </table>
        </div>`;
      }).join('')}
    </div>
    <div style="padding:5px 12px;background:#e8f5e9;font-size:9px;color:#5a7a5a;font-style:italic;border-top:1px solid #c8e6c9">
      * Los planes de vida/asistencia médica son opcionales y están sujetos a condiciones de la póliza. SWEADEN: incluido en la póliza del vehículo. LATINA/MAPFRE/ALIANZA: póliza y factura independientes.
    </div>
  </div>`:''}

  <div class="footer">
    <span>Elaborado por ${exec} · Reliance Asesores Productores de Seguros · Quito, Ecuador</span>
    <span>Cotización válida por 30 días desde ${fecha}</span>
  </div>`;

  // Abrir ventana de impresión limpia
  const win = window.open('','_blank','width=900,height=700');
  win.document.write('<!DOCTYPE html><html><head><meta charset="UTF-8"><title>Cotización '+nombre+'</title></head><body>'+html+'</body></html>');
  win.document.close();
  win.focus();
  setTimeout(()=>{ win.print(); }, 400);
}
function guardarCotizacion(){
  const nombre  = (document.getElementById('cot-nombre')?.value||'').trim();
  const ci      = document.getElementById('cot-ci')?.value||'';
  const marca   = document.getElementById('cot-marca')?.value||'';
  const modelo  = document.getElementById('cot-modelo')?.value||'';
  const anio    = document.getElementById('cot-anio')?.value||'';
  const placa   = (document.getElementById('cot-placa')?.value||'').toUpperCase().trim();
  const ciudad  = document.getElementById('cot-ciudad')?.value||'';
  const region  = document.getElementById('cot-region')?.value||'';
  const tipo    = document.getElementById('cot-tipo')?.value||'NUEVO';
  const celular = document.getElementById('cot-cel')?.value||'';
  const correo  = document.getElementById('cot-email')?.value||'';
  const color   = (document.getElementById('cot-color')?.value||'').toUpperCase().trim();
  const motor   = (document.getElementById('cot-motor')?.value||'').trim();
  const chasis  = (document.getElementById('cot-chasis')?.value||'').trim();
  const asegAnterior  = document.getElementById('cot-aseg-anterior')?.value||'';
  const polizaAnterior = document.getElementById('cot-poliza-anterior')?.value||'';
  const va      = parseFloat(document.getElementById('cot-va')?.value)||0;
  const extras  = parseFloat(document.getElementById('cot-extras')?.value)||0;
  const vaT     = va + extras;
  const desde   = document.getElementById('cot-desde')?.value||'';
  // Calcular vigencia hasta = desde + 1 año
  const hasta   = desde ? (()=>{ const d=new Date(desde+'T00:00:00'); d.setFullYear(d.getFullYear()+1); d.setDate(d.getDate()-1); return d.toISOString().split('T')[0]; })() : '';
  const cuotasTcReq  = parseInt(document.getElementById('cot-cuotas-tc')?.value)||12;
  const cuotasDebReq = parseInt(document.getElementById('cot-cuotas-deb')?.value)||10;
  const axaIncluido = document.getElementById('cot-axa')?.checked||false;
  const vidaLatina  = parseFloat(document.getElementById('cot-vida-latina')?.value)||0;
  const vidaSweaden = parseFloat(document.getElementById('cot-vida-sweaden')?.value)||0;
  const vidaMapfre  = parseFloat(document.getElementById('cot-vida-mapfre')?.value)||0;
  const vidaAlianza = parseFloat(document.getElementById('cot-vida-alianza')?.value)||0;
  const vidaInputsSave = { LATINA:vidaLatina, SWEADEN:vidaSweaden, MAPFRE:vidaMapfre, ALIANZA:vidaAlianza };

  if(!nombre){ showToast('Ingrese el nombre del cliente','error'); return; }
  if(vaT<500){ showToast('Ingrese un valor asegurado válido','error'); return; }

  const selected = getSelectedAseg();
  if(!selected.length){ showToast('Selecciona al menos una aseguradora','error'); return; }

  // Calcular resultados de las aseguradoras seleccionadas
  const resultados = selected.map(name=>{
    const cfg = ASEGURADORAS[name];
    // Usa la tasa que el ejecutivo ingresó en la tarjeta (si la cambió); sino usa el default
    const tasa = typeof cfg.tasa === 'function' ? cfg.tasa(vaT) : _getTasaFromCard(name);
    const axaInc = name==='SWEADEN' ? axaIncluido : false;
    const vida = vidaInputsSave[name]||0;
    const p = calcPrima(vaT, tasa, cfg.pnMin, axaInc, vida, cfg.extraFijo||0);
    const debN = Math.min(cuotasDebReq, cfg.debMax||cuotasDebReq);
    const tc  = calcCuotasTc(p.total, cfg.tcMax, cuotasTcReq, cfg.pisoTC||0);
    const deb = calcCuotasDeb(p.total, debN, cfg.pisoDeb||0);
    const planVida = _getPlanVidaNombre(name, vida);
    return { name, tasa, pn:p.pn, total:p.total, axa:p.axa, vida:p.vida, planVida, extraFijo:p.extraFijo,
             tcN:tc.n, tcCuota:tc.cuota, debN:deb.n, debCuota:deb.cuota };
  });

  // Buscar cliente en DB por CI o nombre
  const clienteMatch = DB.find(c=>
    (ci && c.ci===ci) || c.nombre.trim().toUpperCase()===nombre.toUpperCase()
  );

  const allCotiz = _getCotizaciones();
  const esEdicion  = !!window._editandoCotizId;
  const originalId = window._editandoCotizId; // capturar antes de limpiar
  const codigo  = esEdicion ? window._editandoCotizCodigo  : generarCodigoCotiz();
  const version = esEdicion ? window._editandoCotizVersion : 1;
  if(esEdicion){ window._editandoCotizId=null; window._editandoCotizCodigo=null; window._editandoCotizVersion=null; }

  const cotiz = {
    id: 'CQ' + Date.now(),
    codigo, version, reemplazadaPor: null,
    fecha: new Date().toISOString().split('T')[0],
    ejecutivo: currentUser?.id||'',
    // Datos cliente
    clienteNombre: nombre, clienteCI: ci,
    clienteId: _cotizClienteId || (clienteMatch ? String(clienteMatch.id) : ''),
    celular, correo, ciudad, region,
    // Datos vehículo
    vehiculo: `${marca} ${modelo} ${anio}`.trim(),
    marca, modelo, anio, placa, color, motor, chasis,
    // Datos póliza
    tipo, va: vaT, desde, hasta,
    asegAnterior, polizaAnterior,
    cuotasTc: cuotasTcReq, cuotasDeb: cuotasDebReq,
    extras: parseFloat(document.getElementById('cot-extras')?.value)||0,
    axaIncluido: axaIncluido ? 'SI' : 'NO',
    vidaLatina, vidaSweaden, vidaMapfre, vidaAlianza,
    aseguradoras: selected, resultados,
    estado: 'ENVIADA', asegElegida: null, obsAcept: '', fechaAcept: null,
  };

  // Si es edición, marcar la original como REEMPLAZADA AHORA (solo si el ejecutivo llegó a guardar)
  if(esEdicion && originalId){
    const origIdx = allCotiz.findIndex(c=>String(c.id)===String(originalId));
    if(origIdx>=0){ allCotiz[origIdx].estado='REEMPLAZADA'; allCotiz[origIdx]._dirty=true; }
  }

  cotiz._dirty = true;
  allCotiz.push(cotiz);
  _saveCotizaciones(allCotiz);
  actualizarBadgeCotizaciones();
  _cotizClienteId = ''; // limpiar vínculo de cliente tras guardar
  renderCotizaciones();
  showToast(`✅ ${codigo} guardada — ${nombre.split(' ')[0]} · ${selected.length} aseguradoras`, 'success');
}

// ══════════════════════════════════════════════════════
//  MÓDULO DE COBRANZA — FASE 3
// ══════════════════════════════════════════════════════
let _currentFiltroCobranza = 'mes';

// Extrae todas las cuotas de un cierre con su estado actual
// Normaliza estados legacy → PAGADO / IMPAGO
function _normEstado(s){ return s==='COBRADO'||s==='PAGADO'?'PAGADO':'IMPAGO'; }

function _getCuotasFromCierre(cierre){
  const fp = cierre.formaPago || {};
  const raw = cierre.cuotasEstado;
  const estadosRaw = Array.isArray(raw) ? raw
    : (typeof raw === 'string' && raw.startsWith('[')) ? (()=>{try{return JSON.parse(raw);}catch(e){return [];}})()
    : [];
  const estados = estadosRaw.map(_normEstado);

  const cuotas = [];
  const base = {
    cierreId:      cierre.id,
    clienteNombre: cierre.clienteNombre||'',
    aseguradora:   cierre.aseguradora||'',
    polizaNueva:   cierre.polizaNueva||'',
    clienteId:     cierre.clienteId||cierre._clienteId||'',
    primaTotal:    cierre.primaTotal||0,
  };

  if(fp.forma === 'DEBITO_BANCARIO' && Array.isArray(fp.calendario) && fp.calendario.length){
    fp.calendario.forEach((fecha, i) => {
      cuotas.push({...base,
        idx: i, fecha,
        monto: parseFloat(fp.cuotaMonto)||Math.round((cierre.primaTotal||0)/(fp.calendario.length||1)*100)/100,
        estado: estados[i] || 'IMPAGO',
        nCuota: i+1, totalCuotas: fp.calendario.length,
        tipo: 'DÉBITO', banco: fp.banco||'Produbanco', cuenta: fp.cuenta||cierre.cuenta||'',
      });
    });
  } else if(fp.forma === 'TARJETA_CREDITO'){
    // TC: el financiamiento a cuotas lo maneja el emisor de la tarjeta.
    // Reliance registra 1 solo cobro por el total en la fecha de contacto.
    cuotas.push({...base,
      idx: 0,
      fecha: fp.fechaContacto || cierre.vigDesde || '',
      monto: cierre.primaTotal || 0,
      estado: estados[0] || 'IMPAGO',
      nCuota: 1, totalCuotas: 1,
      tipo: 'TC', banco: fp.banco || '',
      nota: `TC ${fp.nCuotas||''} cuotas (financiado por emisor)`,
    });
  } else if(fp.forma === 'CONTADO'){
    cuotas.push({...base,
      idx: 0, fecha: fp.fechaCobro||cierre.vigDesde||'',
      monto: cierre.primaTotal||0,
      estado: estados[0] || 'IMPAGO',
      nCuota: 1, totalCuotas: 1, tipo: 'CONTADO',
    });
  } else if(fp.forma === 'MIXTO'){
    cuotas.push({...base,
      idx: 0, fecha: fp.fechaInicial||cierre.vigDesde||'',
      monto: parseFloat(fp.montoInicial)||0,
      estado: estados[0] || 'IMPAGO',
      nCuota: 1, totalCuotas: 1+(parseInt(fp.nCuotasResto)||0), tipo: 'MIXTO-INICIAL',
    });
    const nR = parseInt(fp.nCuotasResto||0);
    const montoR = Math.round(((cierre.primaTotal||0)-(parseFloat(fp.montoInicial)||0))/(nR||1)*100)/100;
    for(let i=0; i<nR; i++){
      let fecha = fp.fechaCuotaResto||'';
      if(fecha && i>0){ const d=new Date(fecha); d.setMonth(d.getMonth()+i); fecha=d.toISOString().split('T')[0]; }
      cuotas.push({...base,
        idx: i+1, fecha,
        monto: montoR,
        estado: estados[i+1]||'IMPAGO',
        nCuota: i+2, totalCuotas: 1+nR, tipo: 'MIXTO',
      });
    }
  }
  return cuotas;
}

// Filtrar cuotas según el período seleccionado
function _filtrarCuotas(cuotas, filtro){
  const hoy = new Date(); hoy.setHours(0,0,0,0);
  const iniMes = new Date(hoy.getFullYear(), hoy.getMonth(), 1);
  const finMes = new Date(hoy.getFullYear(), hoy.getMonth()+1, 0);
  const finSem = new Date(hoy); finSem.setDate(hoy.getDate()+7);
  switch(filtro){
    case 'mes':    return cuotas.filter(c=>{ const d=new Date(c.fecha); return d>=iniMes&&d<=finMes; });
    case 'semana': return cuotas.filter(c=>{ const d=new Date(c.fecha); return d>=hoy&&d<=finSem; });
    case 'vencidas': return cuotas.filter(c=>{ const d=new Date(c.fecha); return d<hoy&&c.estado==='IMPAGO'; });
    case 'impago': return cuotas.filter(c=>c.estado==='IMPAGO');
    default: return cuotas;
  }
}

function filtrarCobranza(filtro, btn){
  _currentFiltroCobranza = filtro;
  document.querySelectorAll('#page-cobranza .pill').forEach(b=>b.classList.remove('active'));
  if(btn) btn.classList.add('active');
  renderCobranza(filtro);
}

function renderCobranza(filtro='mes'){
  _currentFiltroCobranza = filtro;
  const allCierres = _getCierres();
  const busq = (document.getElementById('cobranza-search')?.value||'').toLowerCase();

  // Recopilar todas las cuotas
  let todasCuotas = [];
  allCierres.forEach(c => { _getCuotasFromCierre(c).forEach(q=>todasCuotas.push(q)); });

  // Aplicar filtro de período
  let cuotas = _filtrarCuotas(todasCuotas, filtro);

  // Aplicar búsqueda
  if(busq) cuotas = cuotas.filter(c=>
    (c.clienteNombre||'').toLowerCase().includes(busq) ||
    (c.polizaNueva||'').toLowerCase().includes(busq) ||
    (c.aseguradora||'').toLowerCase().includes(busq)
  );

  // Stats
  const hoy = new Date(); hoy.setHours(0,0,0,0);
  const pagadas  = cuotas.filter(c=>c.estado==='PAGADO');
  const impagas  = cuotas.filter(c=>c.estado==='IMPAGO');
  const vencidas = impagas.filter(c=>new Date(c.fecha)<hoy);
  const montoPag = pagadas.reduce((s,c)=>s+c.monto, 0);
  const montoImp = impagas.reduce((s,c)=>s+c.monto, 0);

  const statsEl = document.getElementById('cobranza-stats');
  if(statsEl) statsEl.innerHTML = `
    <div class="stat-card"><div class="stat-value">${cuotas.length}</div><div class="stat-label">Total cuotas</div></div>
    <div class="stat-card" style="border-color:var(--red)"><div class="stat-value" style="color:var(--red)">${vencidas.length}</div><div class="stat-label">Vencidas / Urgentes</div></div>
    <div class="stat-card" style="border-color:var(--green)"><div class="stat-value" style="color:var(--green)">${fmt(montoPag)}</div><div class="stat-label">Monto pagado</div></div>
    <div class="stat-card" style="border-color:var(--accent)"><div class="stat-value" style="color:var(--accent)">${fmt(montoImp)}</div><div class="stat-label">Monto impago</div></div>
  `;

  document.getElementById('cobranza-count').textContent = `${cuotas.length} cuota${cuotas.length!==1?'s':''}`;

  // Ordenar: vencidas primero, luego por fecha asc
  cuotas.sort((a,b)=>{
    const av=new Date(a.fecha)<hoy&&a.estado==='IMPAGO';
    const bv=new Date(b.fecha)<hoy&&b.estado==='IMPAGO';
    if(av!==bv) return av?-1:1;
    return (a.fecha||'').localeCompare(b.fecha||'');
  });

  const wrap = document.getElementById('cobranza-tabla');
  if(!cuotas.length){
    wrap.innerHTML=`<div style="padding:40px;text-align:center;color:var(--muted)">
      <div style="font-size:32px;margin-bottom:8px">✅</div>
      <div>No hay cuotas para este período</div>
    </div>`;
    return;
  }

  const estadoBadgeCob = (e, fecha) => {
    const venc = new Date(fecha) < hoy && e==='IMPAGO';
    if(e==='PAGADO') return `<span class="badge badge-green">✅ Pagado</span>`;
    if(venc)         return `<span class="badge badge-red">⚠️ Vencida</span>`;
    return `<span class="badge badge-orange">⏳ Impago</span>`;
  };
  const tipoBadge = t => {
    const cfg = {DÉBITO:'#1a4c84','TC':'#7c4dff','CONTADO':'#2d6a4f','MIXTO':'#e65100','MIXTO-INICIAL':'#e65100'};
    return `<span style="display:inline-block;padding:2px 7px;border-radius:4px;font-size:10px;font-weight:700;background:${cfg[t]||'#555'}22;color:${cfg[t]||'#555'}">${t}</span>`;
  };

  // Contar gestiones por cuota para el badge
  const allGest = _getGestionCobranza();
  const gestCount = {};
  allGest.forEach(g=>{ const k=`${g.cierreId}_${g.cuotaIdx}`; gestCount[k]=(gestCount[k]||0)+1; });

  const rows = cuotas.map(c=>{
    const gestKey = `${c.cierreId}_${c.idx}`;
    const nGest = gestCount[gestKey]||0;
    const venc = new Date(c.fecha)<hoy && c.estado==='IMPAGO';
    return `
    <tr style="border-bottom:1px solid var(--border)${venc?';background:#fff3e0':''}">
      <td style="padding:8px 12px;font-size:12px">
        <div style="font-weight:600">${c.clienteNombre||'—'}</div>
        <div style="color:var(--muted);font-size:11px">${c.aseguradora||'—'}</div>
      </td>
      <td style="padding:8px 12px;font-size:11px;font-family:'DM Mono',monospace">${c.polizaNueva||'—'}</td>
      <td style="padding:8px 12px;text-align:center">${tipoBadge(c.tipo)}</td>
      <td style="padding:8px 12px;text-align:center;font-size:12px;font-family:'DM Mono',monospace">
        ${c.nCuota} / ${c.totalCuotas}
      </td>
      <td style="padding:8px 12px;font-size:12px;font-family:'DM Mono',monospace;white-space:nowrap">${c.fecha||'—'}</td>
      <td style="padding:8px 12px;text-align:right;font-weight:700;color:var(--green)">${fmt(c.monto)}</td>
      <td style="padding:8px 12px">${estadoBadgeCob(c.estado, c.fecha)}</td>
      <td style="padding:8px 12px">
        <div style="display:flex;gap:4px;flex-wrap:wrap">
          ${c.estado!=='PAGADO'?`<button class="btn btn-sm btn-green" title="Marcar pagada" onclick="marcarCuota('${c.cierreId}',${c.idx},'PAGADO')" style="padding:3px 8px;font-size:11px">✅ Pagado</button>`:''}
          ${c.estado!=='IMPAGO'?`<button class="btn btn-sm btn-ghost" title="Marcar impago" onclick="marcarCuota('${c.cierreId}',${c.idx},'IMPAGO')" style="padding:3px 8px;font-size:11px">⏳ Impago</button>`:''}
          <button class="btn btn-sm btn-ghost" title="Gestión / Historial" onclick="abrirGestionCobranza('${c.cierreId}',${c.idx},'${(c.clienteNombre||'').replace(/'/g,"\\'")}')" style="padding:3px 8px;font-size:11px">📝${nGest>0?` <span style="background:var(--accent);color:#fff;border-radius:8px;padding:0 5px;font-size:10px">${nGest}</span>`:''}</button>
          <button class="btn btn-sm btn-ghost" title="Enviar recordatorio WhatsApp" onclick="enviarRecordatorioCobranza('${c.cierreId}',${c.idx})" style="padding:3px 8px;font-size:11px">📲</button>
        </div>
      </td>
    </tr>`;
  }).join('');

  wrap.innerHTML = `<div class="tbl-wrap"><table>
    <thead><tr>
      <th>Cliente / Aseguradora</th><th>Póliza</th><th>Tipo</th>
      <th style="text-align:center">Cuota</th><th>Fecha</th>
      <th style="text-align:right">Monto</th><th>Estado</th><th>Acciones</th>
    </tr></thead>
    <tbody>${rows}</tbody>
  </table></div>`;
}

function marcarCuota(cierreId, cuotaIdx, estado){
  const allCierres = _getCierres();
  const cierre = allCierres.find(c=>String(c.id)===String(cierreId));
  if(!cierre) return;
  const fp = cierre.formaPago||{};
  const nCuotas = fp.calendario?.length || fp.nCuotas || 1;
  if(!Array.isArray(cierre.cuotasEstado) || cierre.cuotasEstado.length < nCuotas){
    cierre.cuotasEstado = Array(nCuotas).fill('IMPAGO');
  }
  cierre.cuotasEstado[cuotaIdx] = estado;
  cierre._dirty = true;
  _saveCierres(allCierres);
  actualizarBadgeCobranza();
  renderCobranza(_currentFiltroCobranza||'mes');
  const msgs = {PAGADO:'✅ Cuota marcada como PAGADA',IMPAGO:'⏳ Cuota marcada como IMPAGO'};
  showToast(msgs[estado]||estado, estado==='PAGADO'?'success':'info');
}

function enviarRecordatorioCobranza(cierreId, cuotaIdx){
  const cierre = _getCierres().find(c=>String(c.id)===String(cierreId));
  if(!cierre) return;
  const cuotas = _getCuotasFromCierre(cierre);
  const cuota = cuotas[cuotaIdx];
  if(!cuota) return;
  const cliente = cuota.clienteId ? DB.find(x=>String(x.id)===String(cuota.clienteId)) : null;
  const celular = (cliente?.celular||'').replace(/\D/g,'');
  if(!celular){ showToast('Sin número de celular registrado','error'); return; }
  const phone = celular.startsWith('593') ? celular : `593${celular.replace(/^0/,'')}`;
  const nombre = (cierre.clienteNombre||'').split(' ').slice(-2).join(' ');
  const fechaFmt = cuota.fecha ? new Date(cuota.fecha+'T12:00:00').toLocaleDateString('es-EC',{day:'numeric',month:'long',year:'numeric'}) : cuota.fecha;
  const msg = encodeURIComponent(
    `Estimado/a *${nombre}*, le recordamos que la cuota *${cuota.nCuota} de ${cuota.totalCuotas}* de su seguro vehicular con *${cierre.aseguradora}*`+
    (cierre.polizaNueva ? ` — Póliza ${cierre.polizaNueva}` : '')+
    ` por *$${cuota.monto.toFixed(2)}* se procesará el *${fechaFmt}*.\n\nReliance Broker de Seguros`
  );
  window.open(`https://web.whatsapp.com/send?phone=${phone}&text=${msg}`, '_blank');
}

function actualizarBadgeCobranza(){
  const allCierres = _getCierres();
  const hoy = new Date(); hoy.setHours(0,0,0,0);
  let impagasVencidas = 0;
  allCierres.forEach(cierre=>{
    _getCuotasFromCierre(cierre).forEach(c=>{
      if(c.estado==='IMPAGO' && new Date(c.fecha)<hoy) impagasVencidas++;
    });
  });
  const el = document.getElementById('badge-cobranza');
  if(el){ el.textContent=impagasVencidas; el.style.display=impagasVencidas>0?'':'none'; }
}

function exportCobranzaExcel(){
  const allCierres = _getCierres();
  const rows = [];
  allCierres.forEach(cierre=>{
    _getCuotasFromCierre(cierre).forEach(c=>{
      rows.push({
        'Cliente': c.clienteNombre,
        'Aseguradora': c.aseguradora,
        'Póliza': c.polizaNueva,
        'Tipo Pago': c.tipo,
        'Cuota N°': c.nCuota,
        'Total Cuotas': c.totalCuotas,
        'Fecha': c.fecha,
        'Monto': c.monto,
        'Estado': c.estado,
        'Banco': c.banco||'',
        'Cuenta': c.cuenta||'',
      });
    });
  });
  if(!rows.length){ showToast('Sin datos para exportar','error'); return; }
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Cobranza');
  XLSX.writeFile(wb, `Cobranza_Reliance_${new Date().toISOString().split('T')[0]}.xlsx`);
}

// ══════════════════════════════════════════════════════
//  GESTIÓN DE COBRANZA
// ══════════════════════════════════════════════════════
const _GEST_COBR_KEY = '_reliance_gest_cobr';
function _getGestionCobranza(){ try{ return JSON.parse(localStorage.getItem(_GEST_COBR_KEY)||'[]'); }catch(e){return[];} }
function _saveGestionCobranza(arr){ localStorage.setItem(_GEST_COBR_KEY, JSON.stringify(arr)); }

let _gestionCobranzaActiva = null; // {cierreId, cuotaIdx, clienteNombre}

function abrirGestionCobranza(cierreId, cuotaIdx, clienteNombre){
  _gestionCobranzaActiva = {cierreId, cuotaIdx, clienteNombre};
  document.getElementById('gcobr-titulo').textContent = `📝 Gestión — ${clienteNombre}`;
  document.getElementById('gcobr-tipo').value = 'LLAMADA';
  document.getElementById('gcobr-nota').value = '';
  document.getElementById('gcobr-resultado').value = '';
  document.getElementById('gcobr-seguimiento').value = '';
  _renderGestionCobranzaHistorial(cierreId, cuotaIdx);
  openModal('modal-gestion-cobranza');
}

function guardarGestionCobranza(){
  if(!_gestionCobranzaActiva) return;
  const tipo = document.getElementById('gcobr-tipo').value;
  const nota = document.getElementById('gcobr-nota').value.trim();
  const resultado = document.getElementById('gcobr-resultado').value;
  const seguimiento = document.getElementById('gcobr-seguimiento').value;
  if(!nota){ showToast('Ingresa una nota de gestión','error'); return; }
  const all = _getGestionCobranza();
  const entrada = {
    id: Date.now(),
    cierreId: _gestionCobranzaActiva.cierreId,
    cuotaIdx: _gestionCobranzaActiva.cuotaIdx,
    fecha: new Date().toISOString().split('T')[0],
    hora: new Date().toLocaleTimeString('es-EC',{hour:'2-digit',minute:'2-digit'}),
    tipo, nota, resultado, seguimiento,
    ejecutivo: currentUser ? (currentUser.name||currentUser.id) : 'Sistema',
    _dirty: true,
  };
  all.unshift(entrada);
  _saveGestionCobranza(all);
  // Sync SP — guardar _spId igual que el resto de entidades para permitir reintentos
  if(_spReady && typeof spCreate === 'function'){
    spCreate('cobranzas', entrada).then(spId=>{
      if(spId){
        const stored = _getGestionCobranza();
        const idx = stored.findIndex(x=>String(x.id)===String(entrada.id));
        if(idx>=0){ delete stored[idx]._dirty; stored[idx]._spId = spId; _saveGestionCobranza(stored); }
      }
    }).catch(e=>console.warn('SP gest cobr:', e));
  }
  // Si se indicó cambio de estado, aplicarlo
  if(resultado==='PAGADO'||resultado==='IMPAGO'){
    marcarCuota(_gestionCobranzaActiva.cierreId, _gestionCobranzaActiva.cuotaIdx, resultado);
  }
  showToast('✓ Gestión registrada','success');
  _renderGestionCobranzaHistorial(_gestionCobranzaActiva.cierreId, _gestionCobranzaActiva.cuotaIdx);
  document.getElementById('gcobr-nota').value = '';
  document.getElementById('gcobr-resultado').value = '';
  document.getElementById('gcobr-seguimiento').value = '';
}

// Flush pendientes de cobranza — reintenta las gestiones con _dirty=true que no llegaron a SP
async function _flushGestionCobranza(){
  if(!_spReady) return;
  const all = _getGestionCobranza();
  let changed = false;
  for(const g of all){
    if(!g._dirty) continue;
    try{
      if(g._spId){
        await spUpdate('cobranzas', g._spId, g);
      } else {
        const spId = await spCreate('cobranzas', g);
        if(spId) g._spId = spId;
      }
      delete g._dirty;
      changed = true;
    }catch(e){ /* mantener _dirty para próximo intento */ }
  }
  if(changed) _saveGestionCobranza(all);
}

function _renderGestionCobranzaHistorial(cierreId, cuotaIdx){
  const all = _getGestionCobranza();
  const entries = all.filter(g=>String(g.cierreId)===String(cierreId)&&g.cuotaIdx===cuotaIdx);
  const wrap = document.getElementById('gcobr-historial');
  if(!wrap) return;
  if(!entries.length){
    wrap.innerHTML='<div style="padding:16px;text-align:center;color:var(--muted);font-size:12px">Sin gestiones registradas aún</div>';
    return;
  }
  const tipoIcon = {LLAMADA:'📞',WHATSAPP:'💬',EMAIL:'📧',VISITA:'🚗',NOTA:'📝'};
  const resClr = {PAGADO:'var(--green)',IMPAGO:'var(--accent)'};
  wrap.innerHTML = entries.map(g=>`
    <div style="padding:10px 14px;border-bottom:1px solid var(--border);font-size:12px">
      <div style="display:flex;justify-content:space-between;align-items:flex-start;gap:8px">
        <div>
          <span style="font-size:13px">${tipoIcon[g.tipo]||'•'}</span>
          <strong style="margin-left:4px">${g.tipo}</strong>
          ${g.resultado?`<span style="margin-left:8px;padding:1px 8px;border-radius:10px;font-size:10px;font-weight:700;background:${resClr[g.resultado]||'#888'}22;color:${resClr[g.resultado]||'#888'}">${g.resultado}</span>`:''}
        </div>
        <div style="color:var(--muted);font-size:11px;white-space:nowrap">${g.fecha} ${g.hora||''}</div>
      </div>
      <div style="margin-top:4px;color:var(--text)">${g.nota}</div>
      ${g.seguimiento?`<div style="margin-top:3px;font-size:11px;color:var(--accent)">📅 Seguimiento: ${g.seguimiento}</div>`:''}
      <div style="margin-top:3px;font-size:10px;color:var(--muted)">👤 ${g.ejecutivo||'—'}</div>
    </div>`).join('');
}

// ══════════════════════════════════════════════════════
//  COLA DE ENVÍO — Cooldown + Envío Masivo + Agenda
// ══════════════════════════════════════════════════════

// Cooldown por urgencia (días entre contactos)
const COOLDOWN_DIAS = { CRITICO: 2, URGENTE: 4, NORMAL: 7, PROXIMO: 14 };

// Calcula la urgencia de un cliente según días hasta vencimiento
function _urgenciaCliente(c){
  const d = daysUntil(c.hasta);
  if(d < 0 || d > 60) return 'CRITICO';   // ya venció o más de 60 días (lejos)
  if(d <= 15) return 'URGENTE';
  if(d <= 60) return 'NORMAL';
  return 'PROXIMO';
}

// Calcula días desde último contacto
function _diasDesdeContacto(c){
  if(!c.ultimoContacto) return null; // nunca contactado
  const hoy = new Date(); hoy.setHours(0,0,0,0);
  const uc = new Date(c.ultimoContacto); uc.setHours(0,0,0,0);
  return Math.floor((hoy - uc) / 86400000);
}

// Verifica si un cliente está en cooldown
function _enCooldown(c){
  const dias = _diasDesdeContacto(c);
  if(dias === null) return { enCooldown: false, diasRestantes: 0 }; // nunca contactado = listo
  const urgencia = _urgenciaCliente(c);
  const limite = COOLDOWN_DIAS[urgencia] || 7;
  if(dias >= limite) return { enCooldown: false, diasRestantes: 0 };
  return { enCooldown: true, diasRestantes: limite - dias };
}

// Clasifica todos los clientes del ejecutivo para la cola
function _clasificarCola(){
  const clientes = myClientes();
  const listos = [], sinContacto = [], enCooldown = [];

  clientes.forEach(c => {
    // Excluir estados finales
    if(['EMITIDA','CERRADO','ARCHIVADO'].includes(c.estado)) return;

    const urgencia = _urgenciaCliente(c);
    const diasUC = _diasDesdeContacto(c);
    const cd = _enCooldown(c);
    const tieneEmail = !!(c.email && c.email.includes('@'));
    const tieneCelular = !!(c.celular && c.celular.replace(/\D/g,'').length >= 7);

    const item = {
      ...c, _urgencia: urgencia, _diasDesdeContacto: diasUC,
      _tieneEmail: tieneEmail, _tieneCelular: tieneCelular,
      _cooldown: cd
    };

    if(diasUC === null){
      sinContacto.push(item);
    } else if(cd.enCooldown){
      item._diasRestantes = cd.diasRestantes;
      enCooldown.push(item);
    } else {
      listos.push(item);
    }
  });

  // Ordenar: críticos primero
  const urgOrder = { CRITICO: 0, URGENTE: 1, NORMAL: 2, PROXIMO: 3 };
  const sortFn = (a, b) => (urgOrder[a._urgencia]||9) - (urgOrder[b._urgencia]||9);
  listos.sort(sortFn);
  sinContacto.sort(sortFn);
  enCooldown.sort((a, b) => a._diasRestantes - b._diasRestantes);

  return { listos, sinContacto, enCooldown };
}

// Genera agenda de liberación: cuántos clientes se liberan cada día (30 días)
function _generarAgendaCooldown(){
  const clientes = myClientes();
  const hoy = new Date(); hoy.setHours(0,0,0,0);
  const agenda = {}; // fecha -> count

  // Inicializar 30 días
  for(let i = 0; i < 30; i++){
    const d = new Date(hoy); d.setDate(hoy.getDate() + i);
    agenda[d.toISOString().split('T')[0]] = 0;
  }

  clientes.forEach(c => {
    if(['EMITIDA','CERRADO','ARCHIVADO'].includes(c.estado)) return;
    const cd = _enCooldown(c);
    if(!cd.enCooldown || cd.diasRestantes <= 0) return;
    if(cd.diasRestantes > 30) return;
    const fechaLibre = new Date(hoy);
    fechaLibre.setDate(hoy.getDate() + cd.diasRestantes);
    const key = fechaLibre.toISOString().split('T')[0];
    if(agenda[key] !== undefined) agenda[key]++;
  });

  return agenda;
}

let _colaData = null; // cache
const TAMANO_LOTE = 30;

function _calcularLimiteDiario(){
  if(!_colaData) return null;
  const hoy = new Date(); hoy.setHours(0,0,0,0);

  // Clientes que se liberan hoy (diasRestantes <= 1)
  const liberadosHoy = _colaData.enCooldown.filter(c => (c._diasRestantes || 0) <= 1).length;
  const clientesListos = _colaData.listos.length + _colaData.sinContacto.length + liberadosHoy;

  if(!clientesListos) return { limiteDiario:0, diasHabilesRestantes:0, clientesListos:0, lotesSugeridos:0, tamanoLote:TAMANO_LOTE };

  // Determinar fecha límite del período: percentil 80 de fechas de vencimiento activas
  const todos = [..._colaData.listos, ..._colaData.sinContacto, ..._colaData.enCooldown];
  const fechas = todos.filter(c => c.hasta).map(c => new Date(c.hasta)).sort((a,b) => a-b);

  let fechaHasta;
  if(fechas.length){
    const idx = Math.min(Math.floor(fechas.length * 0.8), fechas.length - 1);
    fechaHasta = fechas[idx];
    // Mínimo 7 días calendario en el futuro para no generar límites absurdos
    const minFecha = new Date(hoy); minFecha.setDate(hoy.getDate() + 7);
    if(fechaHasta < minFecha) fechaHasta = minFecha;
  } else {
    fechaHasta = new Date(hoy); fechaHasta.setDate(hoy.getDate() + 30);
  }

  // Contar días hábiles desde hoy hasta fechaHasta (inclusive)
  let diasHabiles = 0;
  const cur = new Date(hoy);
  while(cur <= fechaHasta){
    const dow = cur.getDay();
    if(dow !== 0 && dow !== 6) diasHabiles++;
    cur.setDate(cur.getDate() + 1);
  }
  diasHabiles = Math.max(diasHabiles, 1);

  const limiteDiario = Math.ceil(clientesListos / diasHabiles);
  const lotesSugeridos = Math.max(1, Math.ceil(limiteDiario / TAMANO_LOTE));

  return { limiteDiario, diasHabilesRestantes: diasHabiles, clientesListos, lotesSugeridos, tamanoLote: TAMANO_LOTE, fechaHasta };
}

function _renderColaBanner(){
  const el = document.getElementById('cola-banner-lotes');
  if(!el) return;
  const calc = _calcularLimiteDiario();
  if(!calc || !calc.clientesListos){ el.innerHTML = ''; return; }

  const { limiteDiario, diasHabilesRestantes, clientesListos, lotesSugeridos, tamanoLote } = calc;
  const tamLote = Math.ceil(limiteDiario / lotesSugeridos);
  const HORARIOS = ['9am','11am','1pm','3pm','5pm'];
  const horarios = HORARIOS.slice(0, Math.min(lotesSugeridos, HORARIOS.length)).join(' · ');

  el.innerHTML = `
    <div style="background:linear-gradient(135deg,#f0f7ff,#e8f0fb);border:1px solid #1a4c84;border-radius:10px;padding:12px 16px;display:flex;align-items:center;justify-content:space-between;gap:12px;flex-wrap:wrap">
      <div style="display:flex;align-items:center;gap:10px">
        <span style="font-size:22px">📊</span>
        <div>
          <div style="font-weight:600;font-size:13px;color:#1a4c84">Hoy: enviar <strong>~${limiteDiario} clientes</strong> en <strong>${lotesSugeridos} lote${lotesSugeridos!==1?'s':''}</strong> de ~${tamLote} cada 2h</div>
          <div style="font-size:11px;color:var(--muted);margin-top:2px">${diasHabilesRestantes} días hábiles restantes · ${clientesListos} clientes listos · Horarios: ${horarios}</div>
        </div>
      </div>
    </div>`;
}

function seleccionarLote(){
  // Cambiar filtro a "listos" para priorizar urgentes, luego sin_contacto
  const filtroEl = document.getElementById('cola-filtro-tipo');
  if(filtroEl && filtroEl.value === 'en_cooldown') filtroEl.value = 'listos';
  renderCola();

  const checks = document.querySelectorAll('.cola-check');
  let count = 0;
  checks.forEach(cb => {
    cb.checked = count < TAMANO_LOTE;
    if(count < TAMANO_LOTE) count++;
  });
  _updateColaSelInfo();
  if(count > 0) showToast(`${count} clientes seleccionados para este lote`, 'success');
  else showToast('No hay clientes disponibles en la vista actual', 'error');
}

function initCola(){
  _colaData = _clasificarCola();
  _renderColaStats(_colaData);
  _renderColaBanner();
  _renderColaAgenda();
  renderCola();
  actualizarBadgeCola();
}

function _renderColaStats(data){
  const el = document.getElementById('cola-stats');
  if(!el) return;
  const total = data.listos.length + data.sinContacto.length + data.enCooldown.length;
  const conEmail = [...data.listos, ...data.sinContacto].filter(c => c._tieneEmail).length;
  const conWA = [...data.listos, ...data.sinContacto].filter(c => c._tieneCelular).length;
  el.innerHTML = `
    <div class="stat-card" style="border-color:var(--green)"><div class="stat-value" style="color:var(--green)">${data.listos.length}</div><div class="stat-label">Listos para contactar</div></div>
    <div class="stat-card" style="border-color:#1a4c84"><div class="stat-value" style="color:#1a4c84">${data.sinContacto.length}</div><div class="stat-label">Sin contacto previo</div></div>
    <div class="stat-card" style="border-color:var(--accent)"><div class="stat-value" style="color:var(--accent)">${data.enCooldown.length}</div><div class="stat-label">En cooldown</div></div>
    <div class="stat-card"><div class="stat-value">${conEmail} 📧 / ${conWA} 📲</div><div class="stat-label">Con datos de contacto</div></div>
  `;
}

function _renderColaAgenda(){
  const agenda = _generarAgendaCooldown();
  const el = document.getElementById('cola-agenda');
  if(!el) return;
  const hoy = new Date().toISOString().split('T')[0];
  const dias = Object.entries(agenda).sort((a,b) => a[0].localeCompare(b[0]));
  const maxVal = Math.max(...dias.map(d => d[1]), 1);

  el.innerHTML = dias.map(([fecha, count]) => {
    const d = new Date(fecha + 'T12:00:00');
    const diaSem = d.toLocaleDateString('es-EC', { weekday: 'short' }).substring(0, 2);
    const diaNum = d.getDate();
    const esHoy = fecha === hoy;
    const intensity = count > 0 ? Math.max(0.15, count / maxVal) : 0;
    const bg = count > 0 ? `rgba(26,76,132,${intensity})` : (esHoy ? '#fff3e0' : '#f8f9fa');
    const color = intensity > 0.5 ? '#fff' : '#333';
    return `<div style="width:42px;text-align:center;padding:4px 2px;border-radius:6px;background:${bg};color:${color};font-size:10px;border:${esHoy?'2px solid var(--accent)':'1px solid var(--border)'}">
      <div style="font-weight:600;text-transform:uppercase">${diaSem}</div>
      <div style="font-size:14px;font-weight:700">${diaNum}</div>
      ${count > 0 ? `<div style="font-size:9px;font-weight:700;margin-top:2px">+${count}</div>` : ''}
    </div>`;
  }).join('');
}

function toggleColaAgenda(){
  const wrap = document.getElementById('cola-agenda-wrap');
  const toggle = document.getElementById('cola-agenda-toggle');
  if(!wrap) return;
  const visible = wrap.style.display !== 'none';
  wrap.style.display = visible ? 'none' : 'block';
  if(toggle) toggle.textContent = visible ? '▼ Expandir' : '▲ Colapsar';
}

function renderCola(){
  if(!_colaData) _colaData = _clasificarCola();
  const filtroTipo = document.getElementById('cola-filtro-tipo')?.value || 'listos';
  const filtroUrg = document.getElementById('cola-filtro-urgencia')?.value || '';
  const busq = (document.getElementById('cola-search')?.value || '').toLowerCase();

  let items = [];
  switch(filtroTipo){
    case 'listos': items = _colaData.listos; break;
    case 'sin_contacto': items = _colaData.sinContacto; break;
    case 'en_cooldown': items = _colaData.enCooldown; break;
    default: items = [..._colaData.listos, ..._colaData.sinContacto, ..._colaData.enCooldown]; break;
  }

  if(filtroUrg) items = items.filter(c => c._urgencia === filtroUrg);
  if(busq) items = items.filter(c =>
    (c.nombre||'').toLowerCase().includes(busq) ||
    (c.aseguradora||'').toLowerCase().includes(busq) ||
    (c.email||'').toLowerCase().includes(busq)
  );

  const countEl = document.getElementById('cola-count');
  if(countEl) countEl.textContent = `${items.length} cliente${items.length!==1?'s':''}`;

  const wrap = document.getElementById('cola-tabla');
  if(!wrap) return;

  if(!items.length){
    wrap.innerHTML = `<div style="padding:40px;text-align:center;color:var(--muted)">
      <div style="font-size:32px;margin-bottom:8px">✅</div>
      <div>No hay clientes en esta categoría</div>
    </div>`;
    _updateColaSelInfo();
    return;
  }

  const urgBadge = u => {
    const cfg = { CRITICO: { color:'#b71c1c', bg:'#fde8e0', icon:'🔴' },
                  URGENTE: { color:'#e65100', bg:'#fff3e0', icon:'🟠' },
                  NORMAL:  { color:'#f9a825', bg:'#fffde7', icon:'🟡' },
                  PROXIMO: { color:'#2d6a4f', bg:'#d4edda', icon:'🟢' }};
    const c = cfg[u] || cfg.NORMAL;
    return `<span style="background:${c.bg};color:${c.color};border:1px solid ${c.color}33;border-radius:10px;padding:2px 8px;font-size:10px;font-weight:600">${c.icon} ${u}</span>`;
  };

  const rows = items.map(c => {
    const diasVenc = daysUntil(c.hasta);
    const contactable = filtroTipo !== 'en_cooldown';
    const statusTxt = c._cooldown.enCooldown
      ? `<span style="color:var(--accent);font-size:11px">⏳ ${c._diasRestantes||c._cooldown.diasRestantes}d restantes</span>`
      : (c._diasDesdeContacto === null
        ? '<span style="color:#1a4c84;font-size:11px">🆕 Nunca contactado</span>'
        : `<span style="color:var(--green);font-size:11px">✅ Listo (${c._diasDesdeContacto}d desde últ.)</span>`);

    return `<tr style="border-bottom:1px solid var(--border)">
      <td style="padding:8px 6px;text-align:center">
        ${contactable ? `<input type="checkbox" class="cola-check" data-id="${c.id}" onchange="_updateColaSelInfo()">` : ''}
      </td>
      <td style="padding:8px 12px">${urgBadge(c._urgencia)}</td>
      <td style="padding:8px 12px;font-size:12px">
        <div style="font-weight:600">${c.nombre||'—'}</div>
        <div style="color:var(--muted);font-size:11px">${c.aseguradora||'—'}</div>
      </td>
      <td style="padding:8px 12px;font-size:11px;font-family:'DM Mono',monospace">${c.cedula||c.ruc||'—'}</td>
      <td style="padding:8px 12px;font-size:11px">
        ${c._tieneEmail ? `📧 <span style="color:var(--green)">✓</span>` : `📧 <span style="color:var(--muted)">✗</span>`}
        ${c._tieneCelular ? ` 📲 <span style="color:var(--green)">✓</span>` : ` 📲 <span style="color:var(--muted)">✗</span>`}
      </td>
      <td style="padding:8px 12px;font-size:11px;font-family:'DM Mono',monospace">${c.hasta||'—'}</td>
      <td style="padding:8px 12px;font-size:11px">${c.ultimoContacto||'—'}</td>
      <td style="padding:8px 12px">${statusTxt}</td>
      <td style="padding:8px 6px">
        <div style="display:flex;gap:4px">
          ${contactable && c._tieneEmail ? `<button class="btn btn-sm btn-ghost" title="Email individual" onclick="abrirEmailDesdeCola('${c.id}')" style="padding:3px 6px;font-size:11px">📧</button>` : ''}
          ${contactable && c._tieneCelular ? `<button class="btn btn-sm btn-ghost" title="WhatsApp individual" onclick="abrirWaDesdeCola('${c.id}')" style="padding:3px 6px;font-size:11px">📲</button>` : ''}
        </div>
      </td>
    </tr>`;
  }).join('');

  wrap.innerHTML = `<div class="tbl-wrap"><table>
    <thead><tr>
      <th style="width:30px"></th><th>Urgencia</th><th>Cliente</th><th>Identificación</th>
      <th>Contacto</th><th>Vencimiento</th><th>Últ. Contacto</th><th>Estado</th><th>Acciones</th>
    </tr></thead>
    <tbody>${rows}</tbody>
  </table></div>`;

  _updateColaSelInfo();
}

function toggleAllCola(checked){
  document.querySelectorAll('.cola-check').forEach(cb => cb.checked = checked);
  _updateColaSelInfo();
}

function _updateColaSelInfo(){
  const checks = document.querySelectorAll('.cola-check:checked');
  const n = checks.length;
  const infoEl = document.getElementById('cola-sel-info');
  if(infoEl){
    if(n > TAMANO_LOTE){
      infoEl.innerHTML = `<span style="color:var(--accent);font-weight:600">⚠ ${n} seleccionados (supera el lote de ${TAMANO_LOTE} — considera dividir)</span>`;
    } else {
      infoEl.textContent = `${n} seleccionado${n!==1?'s':''}`;
    }
  }
  const btnEmail = document.getElementById('cola-btn-email');
  const btnWa = document.getElementById('cola-btn-wa');
  if(btnEmail) btnEmail.disabled = n === 0;
  if(btnWa) btnWa.disabled = n === 0;
}

// Abrir email para un cliente individual desde la cola
function abrirEmailDesdeCola(clienteId){
  const c = DB.find(x => String(x.id) === String(clienteId));
  if(!c) return;
  // Registrar contacto
  c._dirty = true;
  _bitacoraAdd(c, 'Email enviado desde Cola de Envío', 'manual');
  c.ultimoContacto = new Date().toISOString().split('T')[0];
  saveDB();
  // Abrir modal de email existente
  emailClienteId = clienteId;
  document.getElementById('email-destino').value = c.email || '';
  selEmailPlantilla('vencimiento', document.querySelector('#email-tipo-pills .pill'));
  openModal('modal-email');
}

function abrirWaDesdeCola(clienteId){
  const c = DB.find(x => String(x.id) === String(clienteId));
  if(!c) return;
  // Registrar contacto
  c._dirty = true;
  _bitacoraAdd(c, 'WhatsApp enviado desde Cola de Envío', 'manual');
  c.ultimoContacto = new Date().toISOString().split('T')[0];
  saveDB();
  // Abrir modal de WhatsApp
  waClienteId = clienteId;
  const celular = (c.celular || '').replace(/\D/g, '');
  const phone = celular.startsWith('593') ? celular : `593${celular.replace(/^0/, '')}`;
  document.getElementById('wa-numero').value = phone;
  selWaPlantilla('vencimiento', document.querySelector('#wa-tipo-pills .pill'));
  openModal('modal-whatsapp');
}

// Envío masivo de email (abre mailto para cada seleccionado + registra contacto)
function envioMasivoEmail(){
  const checks = document.querySelectorAll('.cola-check:checked');
  if(!checks.length){ showToast('Selecciona al menos un cliente','error'); return; }
  if(checks.length > TAMANO_LOTE && !confirm(`Tienes ${checks.length} clientes seleccionados (el lote recomendado es ${TAMANO_LOTE}).\n\n¿Deseas enviar a todos de una vez o prefieres dividir en lotes de ~${TAMANO_LOTE} cada 2h?\n\nPresiona Aceptar para continuar con ${checks.length}, o Cancelar para ajustar.`)) return;

  const ids = Array.from(checks).map(cb => cb.dataset.id);
  const clientes = ids.map(id => DB.find(x => String(x.id) === String(id))).filter(Boolean);
  const conEmail = clientes.filter(c => c.email && c.email.includes('@'));
  const sinEmail = clientes.length - conEmail.length;

  if(!conEmail.length){ showToast('Ninguno de los seleccionados tiene email','error'); return; }

  // Validar cooldown
  const omitidos = [];
  const enviables = [];
  conEmail.forEach(c => {
    const cd = _enCooldown(c);
    if(cd.enCooldown){
      omitidos.push({ nombre: c.nombre, motivo: `En cooldown (${cd.diasRestantes}d)` });
    } else {
      enviables.push(c);
    }
  });

  if(!enviables.length){
    showToast('Todos los seleccionados están en cooldown','error');
    return;
  }

  // Generar lista de emails y abrir mailto masivo
  const emails = enviables.map(c => c.email).join(',');
  const exec = currentUser ? currentUser.name : 'Ejecutivo RELIANCE';
  const asunto = encodeURIComponent('Renovación de Póliza — RELIANCE Broker de Seguros');
  const cuerpo = encodeURIComponent(
    `Estimado/a cliente,\n\nLe informamos que su póliza de seguro está próxima a vencer.\n` +
    `Para renovar y mantener su protección, contáctenos a la brevedad.\n\n` +
    `Saludos cordiales,\n${exec}\nRELIANCE — Asesores de Seguros`
  );

  // Registrar contacto en cada cliente
  const hoy = new Date().toISOString().split('T')[0];
  enviables.forEach(c => {
    c._dirty = true;
    _bitacoraAdd(c, `Email masivo enviado desde Cola de Envío (${enviables.length} clientes)`, 'sistema');
    c.ultimoContacto = hoy;
  });
  saveDB();

  // Abrir mailto
  window.location.href = `mailto:?bcc=${emails}&subject=${asunto}&body=${cuerpo}`;

  // Mostrar reporte
  _mostrarReporteCola(enviables, omitidos, sinEmail, 'email');

  // Refresh
  setTimeout(() => { _colaData = null; initCola(); }, 500);
}

function envioMasivoWA(){
  const checks = document.querySelectorAll('.cola-check:checked');
  if(!checks.length){ showToast('Selecciona al menos un cliente','error'); return; }
  if(checks.length > TAMANO_LOTE && !confirm(`Tienes ${checks.length} clientes seleccionados (el lote recomendado es ${TAMANO_LOTE}).\n\n¿Deseas enviar a todos de una vez o prefieres dividir en lotes de ~${TAMANO_LOTE} cada 2h?\n\nPresiona Aceptar para continuar con ${checks.length}, o Cancelar para ajustar.`)) return;

  const ids = Array.from(checks).map(cb => cb.dataset.id);
  const clientes = ids.map(id => DB.find(x => String(x.id) === String(id))).filter(Boolean);
  const conWA = clientes.filter(c => c.celular && c.celular.replace(/\D/g,'').length >= 7);

  if(!conWA.length){ showToast('Ninguno de los seleccionados tiene celular','error'); return; }

  const omitidos = [];
  const enviables = [];
  conWA.forEach(c => {
    const cd = _enCooldown(c);
    if(cd.enCooldown){
      omitidos.push({ nombre: c.nombre, motivo: `En cooldown (${cd.diasRestantes}d)` });
    } else {
      enviables.push(c);
    }
  });

  if(!enviables.length){
    showToast('Todos los seleccionados están en cooldown','error');
    return;
  }

  // Registrar contacto
  const hoy = new Date().toISOString().split('T')[0];
  enviables.forEach(c => {
    c._dirty = true;
    _bitacoraAdd(c, `WhatsApp masivo desde Cola de Envío (${enviables.length} clientes)`, 'sistema');
    c.ultimoContacto = hoy;
  });
  saveDB();

  // Abrir WhatsApp Web para el primer cliente, mostrar lista del resto
  const first = enviables[0];
  const cel = (first.celular||'').replace(/\D/g,'');
  const phone = cel.startsWith('593') ? cel : `593${cel.replace(/^0/,'')}`;
  const exec = currentUser ? currentUser.name : 'Ejecutivo RELIANCE';
  const msg = encodeURIComponent(
    `Estimado/a *${primerNombre(first.nombre)}*, le recordamos que su póliza de seguro vehicular con *${first.aseguradora||'su aseguradora'}* está próxima a vencer.\n\n` +
    `Para renovar y mantener su protección, contáctenos a la brevedad.\n\n${exec}\nRELIANCE — Asesores de Seguros`
  );
  window.open(`https://web.whatsapp.com/send?phone=${phone}&text=${msg}`, '_blank');

  // Mostrar reporte con lista de números pendientes
  _mostrarReporteCola(enviables, omitidos, clientes.length - conWA.length, 'whatsapp');

  setTimeout(() => { _colaData = null; initCola(); }, 500);
}

function _mostrarReporteCola(enviados, omitidos, sinDatos, canal){
  const repEl = document.getElementById('cola-reporte');
  if(!repEl) return;
  repEl.style.display = 'block';

  const canalIcon = canal === 'email' ? '📧' : '📲';
  const canalTxt = canal === 'email' ? 'Email' : 'WhatsApp';

  let html = `<div class="card" style="border-color:var(--green)">
    <div class="card-header"><div class="card-title">${canalIcon} Reporte de Envío Masivo — ${canalTxt}</div></div>
    <div class="card-body" style="padding:16px">
      <div class="grid-3" style="gap:12px;margin-bottom:16px">
        <div style="background:#d4edda;padding:12px;border-radius:8px;text-align:center">
          <div style="font-size:20px;font-weight:700;color:var(--green)">${enviados.length}</div>
          <div style="font-size:11px;color:var(--green)">Contactados</div>
        </div>
        <div style="background:#fff3e0;padding:12px;border-radius:8px;text-align:center">
          <div style="font-size:20px;font-weight:700;color:var(--accent)">${omitidos.length}</div>
          <div style="font-size:11px;color:var(--accent)">Omitidos (cooldown)</div>
        </div>
        <div style="background:#f8f9fa;padding:12px;border-radius:8px;text-align:center">
          <div style="font-size:20px;font-weight:700;color:var(--muted)">${sinDatos}</div>
          <div style="font-size:11px;color:var(--muted)">Sin datos de ${canalTxt.toLowerCase()}</div>
        </div>
      </div>`;

  if(canal === 'whatsapp' && enviados.length > 1){
    html += `<div style="margin-top:12px"><strong style="font-size:12px">📲 Números pendientes de enviar manualmente:</strong>
      <div style="margin-top:8px;max-height:200px;overflow-y:auto">
        ${enviados.slice(1).map(c => {
          const cel = (c.celular||'').replace(/\D/g,'');
          const ph = cel.startsWith('593') ? cel : `593${cel.replace(/^0/,'')}`;
          return `<div style="display:flex;justify-content:space-between;align-items:center;padding:4px 8px;border-bottom:1px solid var(--border);font-size:12px">
            <span>${c.nombre||'—'}</span>
            <a href="https://web.whatsapp.com/send?phone=${ph}" target="_blank" style="color:var(--green);font-weight:600">📲 ${ph}</a>
          </div>`;
        }).join('')}
      </div>
    </div>`;
  }

  if(omitidos.length){
    html += `<details style="margin-top:12px"><summary style="font-size:12px;cursor:pointer;color:var(--accent)">Ver ${omitidos.length} omitidos por cooldown</summary>
      <div style="margin-top:6px;font-size:11px">
        ${omitidos.map(o => `<div style="padding:3px 8px;border-bottom:1px solid var(--border)">${o.nombre} — ${o.motivo}</div>`).join('')}
      </div>
    </details>`;
  }

  html += `<button class="btn btn-sm" style="margin-top:12px" onclick="document.getElementById('cola-reporte').style.display='none'">✕ Cerrar reporte</button>
    </div></div>`;

  repEl.innerHTML = html;
}

function exportColaExcel(){
  if(!_colaData) _colaData = _clasificarCola();
  const all = [..._colaData.listos, ..._colaData.sinContacto, ..._colaData.enCooldown];
  if(!all.length){ showToast('Sin datos para exportar','error'); return; }
  const rows = all.map(c => ({
    'Nombre': c.nombre||'',
    'Identificación': c.cedula||c.ruc||'',
    'Aseguradora': c.aseguradora||'',
    'Email': c.email||'',
    'Celular': c.celular||'',
    'Vencimiento': c.hasta||'',
    'Últ. Contacto': c.ultimoContacto||'',
    'Urgencia': c._urgencia||'',
    'Estado Cooldown': c._cooldown.enCooldown ? `En cooldown (${c._cooldown.diasRestantes}d)` : 'Listo',
  }));
  const ws = XLSX.utils.json_to_sheet(rows);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Cola de Envío');
  XLSX.writeFile(wb, `Cola_Envio_${new Date().toISOString().split('T')[0]}.xlsx`);
}

function actualizarBadgeCola(){
  const data = _colaData || _clasificarCola();
  const n = data.listos.length + data.sinContacto.length;
  const el = document.getElementById('badge-cola');
  if(el){ el.textContent = n; el.style.display = n > 0 ? '' : 'none'; }
}


function actualizarBadgeCotizaciones(){
  const all = _getCotizaciones();
  const pendientes = all.filter(c=>['ENVIADA','VISTA'].includes(c.estado)).length;
  const el = document.getElementById('badge-cotizaciones');
  if(el){ el.textContent = pendientes||''; el.style.display=pendientes?'':'none'; }
}


// ══════════════════════════════════════════════════════
//  MÓDULO COTIZACIONES
// ══════════════════════════════════════════════════════
const CQ_ESTADOS = {
  'ENVIADA':    { color:'#1a4c84', bg:'#e8f0fb', icon:'📤' },
  'VISTA':      { color:'#7c4dff', bg:'#ede7f6', icon:'👁'  },
  'ACEPTADA':   { color:'#2d6a4f', bg:'#d4edda', icon:'✅' },
  'EN EMISIÓN': { color:'#e65100', bg:'#fff3e0', icon:'📝' },
  'EMITIDA':    { color:'#1b5e20', bg:'#c8e6c9', icon:'🛡' },
  'RECHAZADA':  { color:'#b71c1c', bg:'#fde8e0', icon:'❌' },
  'VENCIDA':    { color:'#78909c', bg:'#eceff1', icon:'⌛' },
};

function cqBadge(estado){
  const e = CQ_ESTADOS[estado]||CQ_ESTADOS['ENVIADA'];
  return `<span style="background:${e.bg};color:${e.color};border:1px solid ${e.color}44;
    border-radius:12px;padding:2px 9px;font-size:10px;font-weight:600;white-space:nowrap">
    ${e.icon} ${estado}</span>`;
}

// ── Exportar cotizaciones aceptadas / en emisión a Excel ─
function exportarCotizAceptadas(){
  const all = _getCotizaciones();
  const esAdmin = currentUser?.rol==='admin';

  const lista = all.filter(c=>{
    if(!esAdmin && String(c.ejecutivo) !== String(currentUser?.id||"")) return false;
    return ['ACEPTADA','EN EMISIÓN','EMITIDA'].includes(c.estado);
  }).sort((a,b)=>(b.fecha||'').localeCompare(a.fecha||''));

  if(!lista.length){ showToast('No hay cotizaciones aceptadas para exportar','error'); return; }

  const exec_map = {};
  USERS.forEach(u=>exec_map[u.id]=u.name);

  // Filas del Excel — columnas completas para emisión de póliza
  const filas = lista.map(cot=>{
    const r = (cot.resultados||[]).find(x=>x.name===cot.asegElegida)||{};
    // Compatibilidad: si no tiene campos directos, extraer de vehiculo string
    const partes = (cot.vehiculo||'').split(' ');
    const anioV  = cot.anio || (partes.length>0 && /^\d{4}$/.test(partes[partes.length-1]) ? partes[partes.length-1] : '');
    const marcaV = cot.marca || partes[0] || '';
    const modeloV= cot.modelo || (anioV ? partes.slice(0,-1).slice(1).join(' ') : partes.slice(1).join(' '));
    // Vigencia hasta: usar guardado o calcular
    const hastaV = cot.hasta || (cot.desde ? (()=>{ const d=new Date(cot.desde+'T00:00:00'); d.setFullYear(d.getFullYear()+1); d.setDate(d.getDate()-1); return d.toISOString().split('T')[0]; })() : '');
    return {
      'Código':            cot.codigo||cot.id,
      'Fecha Cotización':  cot.fecha,
      'Estado':            cot.estado,
      'Ejecutivo':         exec_map[cot.ejecutivo]||cot.ejecutivo||'—',
      // Datos cliente
      'Nombre Cliente':    cot.clienteNombre,
      'CI / RUC':          cot.clienteCI||'',
      'Celular':           cot.celular||'',
      'Correo':            cot.correo||'',
      'Ciudad':            cot.ciudad||'',
      'Región':            cot.region||'',
      // Datos vehículo
      'Marca':             marcaV,
      'Modelo':            modeloV,
      'Año':               anioV,
      'Placa':             cot.placa||'',
      'Color':             cot.color||'',
      'N° Motor':          cot.motor||'',
      'N° Chasis':         cot.chasis||'',
      // Datos póliza
      'Tipo':              cot.tipo==='NUEVO'?'NUEVA':'RENOVACIÓN',
      'Aseg. Anterior':    cot.asegAnterior||'',
      'Póliza Anterior':   cot.polizaAnterior||'',
      'Valor Asegurado':   Number(cot.va||0),
      'Vigencia Desde':    cot.desde||'',
      'Vigencia Hasta':    hastaV,
      // Datos prima
      'Aseguradora Elegida': cot.asegElegida||'',
      'Prima Neta':        r.pn    ? Number(r.pn.toFixed(2))    : '',
      'Prima Total':       r.total ? Number(r.total.toFixed(2)) : '',
      'Cuotas TC':         r.tcN||'',
      'Cuota TC $':        r.tcCuota  ? Number(r.tcCuota.toFixed(2))  : '',
      'Cuotas Débito':     r.debN||'',
      'Cuota Débito $':    r.debCuota ? Number(r.debCuota.toFixed(2)) : '',
      'Observación':       cot.obsAcept||'',
    };
  });

  if(typeof XLSX==='undefined'){ showToast('Librería Excel no disponible','error'); return; }

  const ws = XLSX.utils.json_to_sheet(filas);

  // Ancho de columnas
  ws['!cols'] = [
    {wch:14},{wch:14},{wch:12},{wch:18},  // Código, Fecha, Estado, Ejecutivo
    {wch:28},{wch:13},{wch:14},{wch:26},{wch:12},{wch:10},  // Cliente: nombre, CI, celular, correo, ciudad, región
    {wch:12},{wch:22},{wch:6},{wch:10},{wch:10},{wch:18},{wch:22},  // Vehículo: marca, modelo, año, placa, color, motor, chasis
    {wch:12},{wch:18},{wch:18},{wch:14},{wch:14},{wch:14},  // Póliza: tipo, aseg ant, poliza ant, VA, desde, hasta
    {wch:18},{wch:11},{wch:11},{wch:10},{wch:10},{wch:13},{wch:13},{wch:22}  // Prima
  ];

  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Cotizaciones Aceptadas');

  // Segunda hoja: resumen por aseguradora
  const porAseg = {};
  lista.forEach(cot=>{
    const aseg = cot.asegElegida||'Sin aseguradora';
    const r = (cot.resultados||[]).find(x=>x.name===cot.asegElegida)||{};
    if(!porAseg[aseg]) porAseg[aseg]={ aseg, cantidad:0, totalPN:0, totalPT:0 };
    porAseg[aseg].cantidad++;
    porAseg[aseg].totalPN  += r.pn    || 0;
    porAseg[aseg].totalPT  += r.total || 0;
  });
  const resumen = Object.values(porAseg).map(x=>({
    'Aseguradora':    x.aseg,
    'Cotizaciones':   x.cantidad,
    'Total Prima Neta':  Number(x.totalPN.toFixed(2)),
    'Total Prima Final': Number(x.totalPT.toFixed(2)),
  }));
  const ws2 = XLSX.utils.json_to_sheet(resumen);
  ws2['!cols']=[{wch:18},{wch:13},{wch:18},{wch:18}];
  XLSX.utils.book_append_sheet(wb, ws2, 'Resumen por Aseguradora');

  const fecha = new Date().toISOString().split('T')[0];
  XLSX.writeFile(wb, `Cotizaciones_Aceptadas_${fecha}.xlsx`);
  showToast(`✅ Excel exportado — ${lista.length} cotización${lista.length!==1?'es':''}`, 'success');
}

// ── Nueva versión de cotización ─────────────────────────
function nuevaVersionCotiz(id){
  const all = _getCotizaciones();
  const idx = all.findIndex(c=>String(c.id)===String(id));
  if(idx<0) return;
  const original = all[idx];
  all[idx].estado = 'REEMPLAZADA';
  all[idx]._dirty = true;
  _saveCotizaciones(all);
  actualizarBadgeCotizaciones();
  showPage('cotizador');
  setTimeout(()=>{
    const n=document.getElementById('cot-nombre'); if(n) n.value=original.clienteNombre||'';
    const ci=document.getElementById('cot-ci');    if(ci) ci.value=original.clienteCI||'';
    const pla=document.getElementById('cot-placa');if(pla) pla.value=original.placa||'';
    const ciu=document.getElementById('cot-ciudad');if(ciu) ciu.value=original.ciudad||'';
    const des=document.getElementById('cot-desde');if(des) des.value=original.desde||'';
    // Vehículo — usar campos directos si existen, si no parsear de vehiculo string
    const anioV = original.anio || (()=>{ const p=(original.vehiculo||'').split(' '); return p.length>0&&/^\d{4}$/.test(p[p.length-1])?p[p.length-1]:''; })();
    const marcaV = original.marca || (original.vehiculo||'').split(' ')[0] || '';
    const modeloV = original.modelo || (()=>{ const p=(original.vehiculo||'').split(' ').filter(x=>!/^\d{4}$/.test(x)); return p.slice(1).join(' '); })();
    _setMarcaSelect(document.getElementById('cot-marca'), marcaV);
    const moEl=document.getElementById('cot-modelo'); if(moEl) moEl.value=modeloV;
    const aEl=document.getElementById('cot-anio');    if(aEl) aEl.value=anioV;
    const vaEl=document.getElementById('cot-va');     if(vaEl) vaEl.value=original.va||'';
    // Campos adicionales vehículo y cliente
    const colEl=document.getElementById('cot-color');    if(colEl) colEl.value=original.color||'';
    const motEl=document.getElementById('cot-motor');    if(motEl) motEl.value=original.motor||'';
    const chaEl=document.getElementById('cot-chasis');   if(chaEl) chaEl.value=original.chasis||'';
    const regEl=document.getElementById('cot-region');   if(regEl) regEl.value=original.region||'SIERRA';
    const tipEl=document.getElementById('cot-tipo');     if(tipEl) tipEl.value=original.tipo||'NUEVO';
    const celEl=document.getElementById('cot-cel');      if(celEl) celEl.value=original.celular||'';
    const emEl=document.getElementById('cot-email');     if(emEl) emEl.value=original.correo||'';
    const aaPEl=document.getElementById('cot-aseg-anterior');   if(aaPEl) aaPEl.value=original.asegAnterior||'';
    const polPEl=document.getElementById('cot-poliza-anterior');if(polPEl) polPEl.value=original.polizaAnterior||'';
    document.querySelectorAll('.aseg-check').forEach(chk=>{
      chk.checked=(original.aseguradoras||[]).includes(chk.dataset.aseg||chk.value);
    });
    showToast(`↺ Nueva versión de ${original.codigo||id} — ajusta y guarda`,'info');
  },200);
}

// ── Editar cotización (corrección de datos) ─────────────
function editarCotizacion(id){
  const all = _getCotizaciones();
  const idx = all.findIndex(c=>String(c.id)===String(id));
  if(idx<0) return;
  const original = all[idx];
  if(!['ENVIADA'].includes(original.estado)){
    showToast('Solo se pueden editar cotizaciones ENVIADAS','error'); return;
  }
  window._editandoCotizId      = id;
  window._editandoCotizCodigo  = original.codigo;
  window._editandoCotizVersion = (original.version||1)+1;
  // NO marcar REEMPLAZADA aquí — se hace en guardarCotizacion() solo si el ejecutivo guarda la nueva versión
  actualizarBadgeCotizaciones();
  showPage('cotizador');
  setTimeout(()=>{
    const n=document.getElementById('cot-nombre'); if(n) n.value=original.clienteNombre||'';
    const ci=document.getElementById('cot-ci');    if(ci) ci.value=original.clienteCI||'';
    const pla=document.getElementById('cot-placa');if(pla) pla.value=original.placa||'';
    const ciu=document.getElementById('cot-ciudad');if(ciu) ciu.value=original.ciudad||'';
    const des=document.getElementById('cot-desde');if(des) des.value=original.desde||'';
    // Vehículo — usar campos directos si existen, si no parsear de vehiculo string
    const anioV = original.anio || (()=>{ const p=(original.vehiculo||'').split(' '); return p.length>0&&/^\d{4}$/.test(p[p.length-1])?p[p.length-1]:''; })();
    const marcaV = original.marca || (original.vehiculo||'').split(' ')[0] || '';
    const modeloV = original.modelo || (()=>{ const p=(original.vehiculo||'').split(' ').filter(x=>!/^\d{4}$/.test(x)); return p.slice(1).join(' '); })();
    _setMarcaSelect(document.getElementById('cot-marca'), marcaV);
    const moEl=document.getElementById('cot-modelo'); if(moEl) moEl.value=modeloV;
    const aEl=document.getElementById('cot-anio');    if(aEl) aEl.value=anioV;
    const vaEl=document.getElementById('cot-va');     if(vaEl) vaEl.value=original.va||'';
    // Campos adicionales vehículo y cliente
    const colEl=document.getElementById('cot-color');    if(colEl) colEl.value=original.color||'';
    const motEl=document.getElementById('cot-motor');    if(motEl) motEl.value=original.motor||'';
    const chaEl=document.getElementById('cot-chasis');   if(chaEl) chaEl.value=original.chasis||'';
    const regEl=document.getElementById('cot-region');   if(regEl) regEl.value=original.region||'SIERRA';
    const tipEl=document.getElementById('cot-tipo');     if(tipEl) tipEl.value=original.tipo||'NUEVO';
    const celEl=document.getElementById('cot-cel');      if(celEl) celEl.value=original.celular||'';
    const emEl=document.getElementById('cot-email');     if(emEl) emEl.value=original.correo||'';
    const aaPEl=document.getElementById('cot-aseg-anterior');   if(aaPEl) aaPEl.value=original.asegAnterior||'';
    const polPEl=document.getElementById('cot-poliza-anterior');if(polPEl) polPEl.value=original.polizaAnterior||'';
    document.querySelectorAll('.aseg-check').forEach(chk=>{
      chk.checked=(original.aseguradoras||[]).includes(chk.dataset.aseg||chk.value);
    });
    showToast(`✏ Editando ${original.codigo} v${window._editandoCotizVersion} — corrige y guarda`,'info');
  },200);
}


// Normaliza resultados de cotización a array (puede venir como JSON string desde SP)
function _parseResultados(r){
  if(!r) return [];
  if(Array.isArray(r)) return r;
  if(typeof r==='string'){ try{ const p=JSON.parse(r); return Array.isArray(p)?p:[]; }catch(e){return [];} }
  return [];
}
function renderCotizaciones(){
  const all = _getCotizaciones();

  // Poblar filtro aseguradoras
  const selAseg = document.getElementById('cq-filter-aseg');
  if(selAseg && selAseg.options.length===1){
    Object.keys(ASEGURADORAS).forEach(a=>{
      const o=document.createElement('option'); o.value=a; o.textContent=a;
      selAseg.appendChild(o);
    });
  }

  // Stats
  const activas = all.filter(c=>c.estado!=='REEMPLAZADA');
  const pend  = activas.filter(c=>['ENVIADA','VISTA'].includes(c.estado)).length;
  const acept = activas.filter(c=>['ACEPTADA','EN EMISIÓN','EMITIDA'].includes(c.estado)).length;
  const tasa  = activas.length ? Math.round(acept/activas.length*100) : 0;
  document.getElementById('cq-stat-total').textContent = activas.length;
  document.getElementById('cq-stat-pend').textContent  = pend;
  document.getElementById('cq-stat-acept').textContent = acept;
  document.getElementById('cq-stat-tasa').textContent  = tasa+'%';

  // Filtros
  const q      = (document.getElementById('cq-search')?.value||'').toLowerCase();
  const fEstado= document.getElementById('cq-filter-estado')?.value||'';
  const fAseg  = document.getElementById('cq-filter-aseg')?.value||'';
  const esAdmin= currentUser?.rol==='admin';

  const mostrarReemp = fEstado==='REEMPLAZADA';
  let lista = all.filter(c=>{
    if(!esAdmin && String(c.ejecutivo) !== String(currentUser?.id||"")) return false;
    if(c.estado==='REEMPLAZADA' && !mostrarReemp) return false;
    if(q && !(c.clienteNombre||'').toLowerCase().includes(q) && !c.clienteCI?.includes(q) && !(c.codigo||'').toLowerCase().includes(q)) return false;
    if(fEstado && c.estado!==fEstado) return false;
    if(fAseg && !(Array.isArray(c.aseguradoras)?c.aseguradoras:(typeof c.aseguradoras==='string'?JSON.parse(c.aseguradoras||'[]'):[]) ).includes(fAseg)) return false;
    return true;
  }).sort((a,b)=>(b.fecha||'').localeCompare(a.fecha||'')||(String(b.id)||'').localeCompare(String(a.id)||''));// más recientes primero

  document.getElementById('cq-count').textContent = `${lista.length} cotización${lista.length!==1?'es':''}`;

  const exec_map = {};
  USERS.forEach(u=>exec_map[u.id]=u.name);

  const tbody = document.getElementById('cq-tbody');
  if(!tbody) return;

  if(!lista.length){
    tbody.innerHTML=`<tr><td colspan="9"><div class="empty-state">
      <div class="empty-icon">📋</div>
      <p>No hay cotizaciones${fEstado||q?' con esos filtros':' registradas aún'}</p>
      <button class="btn btn-primary btn-sm" onclick="showPage('cotizador')">+ Crear primera cotización</button>
    </div></td></tr>`;
    return;
  }

  tbody.innerHTML = lista.map(c=>{
    const exec = exec_map[c.ejecutivo]||c.ejecutivo||'—';
    // Días desde cotización (desde fecha real)
    const msDesde = Date.now() - new Date(c.fecha+'T00:00:00').getTime();
    const dias = Math.max(0, Math.floor(msDesde/(1000*60*60*24)));
    const vencida = dias>30 && ['ENVIADA','VISTA'].includes(c.estado);
    const diasLabel = dias===0?'hoy':dias===1?'ayer':`hace ${dias}d`;

    // Mini chips aseguradoras con sus totales
    const asegChips = _parseResultados(c.resultados).map(r=>{
      const cfg = ASEGURADORAS[r.name];
      const esElegida = c.asegElegida===r.name;
      return `<span style="background:${esElegida?(cfg?.color||'#333')+'22':'var(--warm)'};
        color:${esElegida?(cfg?.color||'#333'):'var(--ink)'};
        border:1px solid ${esElegida?(cfg?.color||'#333')+'66':'var(--border)'};
        border-radius:4px;padding:2px 6px;font-size:10px;font-weight:${esElegida?700:400};
        white-space:nowrap">
        ${esElegida?'✓ ':''}${r.name} $${r.total?.toFixed(0)||'—'}
      </span>`;
    }).join(' ');

    // Acciones según estado
    let acciones = '';
    if(['ENVIADA','VISTA'].includes(c.estado)){
      acciones = `
        <button class="btn btn-xs" style="background:#e8f0fb;color:var(--accent2)"
          onclick="cambiarEstadoCotiz('${c.id}','VISTA')" title="Marcar como vista">👁</button>
        <button class="btn btn-green btn-xs"
          onclick="abrirAceptarCotiz('${c.id}')" title="Cliente acepta">✅ Acepta</button>
        <button class="btn btn-xs" style="background:#fde8e0;color:#b71c1c"
          onclick="cambiarEstadoCotiz('${c.id}','RECHAZADA')" title="Rechaza">❌</button>
        <button class="btn btn-xs" style="background:#fff3e0;color:#e65100"
          onclick="editarCotizacion('${c.id}')" title="Editar datos">✏</button>
        <button class="btn btn-xs" style="background:#f3e5f5;color:#7b1fa2"
          onclick="nuevaVersionCotiz('${c.id}')" title="Nueva versión con cambios">↺ v.nueva</button>`;
    } else if(c.estado==='RECHAZADA'){
      acciones = `
        <button class="btn btn-xs" style="background:#f3e5f5;color:#7b1fa2"
          onclick="nuevaVersionCotiz('${c.id}')" title="Nueva versión">↺ Nueva versión</button>`;
    } else if(c.estado==='ACEPTADA'){
      acciones = `
        <button class="btn btn-primary btn-xs"
          onclick="irAEmision('${c.id}')" title="Registrar cierre / emisión">📝 Emitir</button>`;
    } else if(c.estado==='EN EMISIÓN'){
      // Botón "Emitida" eliminado — la cotización pasa a EMITIDA automáticamente al registrar el cierre
    }
    acciones += `
      <button class="btn btn-ghost btn-xs" onclick="verDetalleCotiz('${c.id}')" title="Ver detalle">🔍</button>
      <button class="btn btn-ghost btn-xs" onclick="reimprimirCotiz('${c.id}')" title="Re-imprimir PDF">🖨</button>`;

    const esReemp = c.estado==='REEMPLAZADA';
    return `<tr style="${esReemp?'opacity:.4;background:#f8f8f6':vencida?'opacity:.7':''}">
      <td>
        <div class="mono" style="font-weight:700;font-size:12px;color:var(--accent2);letter-spacing:.3px">
          ${c.codigo||('CQ-'+String(c.id).replace('CQ',''))}
          ${(c.version||1)>1?`<span style="font-size:9px;background:var(--warm);border-radius:3px;padding:1px 4px;color:var(--muted);margin-left:3px">v${c.version}</span>`:''}
        </div>
        <div class="mono" style="font-size:10px;color:var(--muted)">${c.fecha} · ${diasLabel}</div>
      </td>
      <td>
        <div style="font-weight:600;font-size:12px">${c.clienteNombre}</div>
        ${c.clienteCI?`<div class="mono" style="font-size:10px;color:var(--muted)">${c.clienteCI}</div>`:''}
      </td>
      <td style="font-size:11px">
        <div>${c.vehiculo||'—'}</div>
        ${c.placa?`<div class="mono" style="font-size:10px;color:var(--muted)">${c.placa}</div>`:''}
      </td>
      <td style="text-align:right;font-family:'DM Mono',monospace;font-size:12px;font-weight:600">
        $${Number(c.va||0).toLocaleString('es-EC')}
      </td>
      <td style="max-width:240px">
        <div style="display:flex;flex-wrap:wrap;gap:3px">${asegChips}</div>
      </td>
      <td>${cqBadge(vencida?'VENCIDA':c.estado)}</td>
      <td style="font-size:11px;color:var(--accent2);font-weight:600">
        ${c.asegElegida||'<span style="color:var(--muted)">—</span>'}
      </td>
      <td style="font-size:11px">${exec}</td>
      <td>
        <div style="display:flex;gap:4px;flex-wrap:wrap;justify-content:center">${acciones}</div>
      </td>
    </tr>`;
  }).join('');
}

// ── Cambiar estado simple ────────────────────────────────
function cambiarEstadoCotiz(id, nuevoEstado){
  const all = _getCotizaciones();
  const idx = all.findIndex(c=>String(c.id)===String(id));
  if(idx<0) return;
  all[idx].estado = nuevoEstado;
  all[idx]._dirty = true;
  _saveCotizaciones(all);
  renderCotizaciones();
  actualizarBadgeCotizaciones();
  showToast(`Estado actualizado: ${nuevoEstado}`);
}

// ── Modal confirmar aceptación ───────────────────────────
let cotizAceptarId = null;
let cotizAsegSeleccionada = null;
let cotizFormaPagoSeleccionada  = null; // 'TARJETA_CREDITO' | 'DEBITO_BANCARIO' | 'CONTADO' | 'MIXTO'
let cotizCuotasElegidasAcept    = null; // cuotas reales que eligió el cliente al aceptar

// Genera código único COT-AAMM-NNN (secuencial global continuo)
function generarCodigoCotiz(){
  const all = _getCotizaciones();
  let maxN = 0;
  all.forEach(cot=>{
    if(cot.codigo){
      const m = cot.codigo.match(/COT-\d{4}-(\d+)/);
      if(m) maxN = Math.max(maxN, parseInt(m[1]));
    }
  });
  const now = new Date();
  const aamm = String(now.getFullYear()).slice(2) + String(now.getMonth()+1).padStart(2,'0');
  return 'COT-' + aamm + '-' + String(maxN+1).padStart(3,'0');
}

// Sincroniza el estado de la cotización cuando cambia el estado del cliente en cartera.
// Busca la cotización EN EMISIÓN vinculada al cliente y la actualiza.
function sincronizarCotizPorCliente(clienteId, clienteNombre, clienteCI, nuevoEstadoCliente){
  try{
    const all = _getCotizaciones();
    let changed = false;

    // Buscar cotizaciones EN EMISIÓN o ACEPTADA vinculadas a este cliente
    all.forEach(cot=>{
      if(!['EN EMISIÓN','ACEPTADA'].includes(cot.estado)) return;
      const vinculada =
        (clienteId  && String(cot.clienteId) === String(clienteId)) ||
        (clienteCI  && clienteCI.length>3 && cot.clienteCI === clienteCI) ||
        (clienteNombre && cot.clienteNombre.trim().toUpperCase() === clienteNombre.trim().toUpperCase());
      if(!vinculada) return;

      // Mapeo estado cliente → estado cotización
      if(['RENOVADO','PÓLIZA VIGENTE'].includes(nuevoEstadoCliente)){
        cot.estado = 'EMITIDA';
        cot.fechaAcept = cot.fechaAcept || new Date().toISOString().split('T')[0];
        changed = true;
      }
    });

    if(changed){
      _saveCotizaciones(all);
      actualizarBadgeCotizaciones();
    }
  }catch(e){ console.error('sincronizarCotizPorCliente:', e); }
}

// Migrar IDs numéricos viejos a string con prefijo CQ (compatibilidad)
function migrarCotizacionesIds(){
  try{
    const all = _getCotizaciones();
    let changed = false;
    all.forEach(c=>{
      if(typeof c.id === 'number'){ c.id = 'CQ'+c.id; changed=true; }
    });
    if(changed){ _cache.cotizaciones=all; localStorage.setItem('reliance_cotizaciones',JSON.stringify(all)); }
  }catch(e){}
}

function abrirAceptarCotiz(id){
  const all = _getCotizaciones();
  const c = all.find(x=>String(x.id)===String(id)); if(!c) return;
  cotizAceptarId = String(id); // normalizar siempre a string
  cotizAsegSeleccionada      = null;
  cotizFormaPagoSeleccionada = null;
  cotizCuotasElegidasAcept   = null;

  // Info cliente
  document.getElementById('modal-cotiz-cliente-info').innerHTML=`
    <div style="font-weight:700;font-size:14px;margin-bottom:4px">${c.clienteNombre}</div>
    <div style="color:var(--muted);font-size:12px">
      ${c.vehiculo} · VA: $${Number(c.va).toLocaleString('es-EC')} · Vigencia desde: ${c.desde||'—'}
    </div>`;

  // Guardar resultados globalmente — evita JSON.stringify en onclick (rompe HTML)
  window._cotizResultados = c.resultados || [];
  document.getElementById('cotiz-aseg-options').innerHTML=_parseResultados(c.resultados).map(r=>{
    const cfg = ASEGURADORAS[r.name]||{color:'#333'};
    return `<button class="cotiz-aseg-btn" data-name="${r.name}"
      onclick="seleccionarAsegAcept('${r.name}')"
      style="border:2px solid ${cfg.color}44;border-radius:8px;padding:10px 14px;
        background:var(--paper);cursor:pointer;text-align:left;transition:all .15s;min-width:140px">
      <div style="font-weight:700;color:${cfg.color};font-size:13px;margin-bottom:4px">${r.name}</div>
      <div style="font-size:18px;font-weight:700;color:var(--ink)">$${r.total?.toFixed(2)}</div>
      <div style="font-size:10px;color:var(--muted);margin-top:2px">
        💳 ${r.tcN}×$${r.tcCuota?.toFixed(2)} · 🏦 ${r.debN}×$${r.debCuota?.toFixed(2)}
      </div>
    </button>`;
  }).join('');

  document.getElementById('cotiz-resumen-elegida').style.display='none';
  document.getElementById('cotiz-obs-acept').value='';
  document.getElementById('btn-confirmar-acept').disabled=true;
  openModal('modal-cotiz-aceptar');
}

function seleccionarAsegAcept(name){
  const resultado = (window._cotizResultados||[]).find(r=>r.name===name)||{};
  cotizAsegSeleccionada      = name;
  cotizFormaPagoSeleccionada = null;
  cotizCuotasElegidasAcept   = null;
  window._cotizResActual     = resultado;
  const cfg = ASEGURADORAS[name]||{color:'#333'};

  // Resaltar botón seleccionado
  document.querySelectorAll('.cotiz-aseg-btn').forEach(btn=>{
    const sel = btn.dataset.name===name;
    btn.style.border=`2px solid ${sel?cfg.color:cfg.color+'44'}`;
    btn.style.background = sel ? cfg.color+'15' : 'var(--paper)';
    btn.style.transform = sel ? 'scale(1.02)' : '';
  });

  // Resumen de montos
  const wrap = document.getElementById('cotiz-resumen-elegida');
  wrap.style.display='block';
  wrap.style.borderColor=cfg.color;
  document.getElementById('cotiz-resumen-body').innerHTML=`
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:12px">
      <div><span style="color:var(--muted)">Aseguradora:</span> <b style="color:${cfg.color}">${name}</b></div>
      <div><span style="color:var(--muted)">Prima Neta:</span> <b>$${resultado.pn?.toFixed(2)||'—'}</b></div>
      <div><span style="color:var(--muted)">Total:</span> <b style="font-size:15px">$${resultado.total?.toFixed(2)||'—'}</b></div>
      <div><span style="color:var(--muted)">TC hasta ${resultado.tcN||'—'} cuotas:</span> <b>$${resultado.tcCuota?.toFixed(2)||'—'}/mes</b></div>
      <div><span style="color:var(--muted)">Débito hasta ${resultado.debN||'—'} cuotas:</span> <b>$${resultado.debCuota?.toFixed(2)||'—'}/mes</b></div>
    </div>`;

  // Opciones de forma de pago — 8 grupos
  const pagoSection = document.getElementById('cotiz-pago-section');
  const pagoOpts    = document.getElementById('cotiz-pago-options');
  const cuotasPanel = document.getElementById('cotiz-cuotas-panel');
  if(cuotasPanel){ cuotasPanel.innerHTML=''; cuotasPanel.style.display='none'; }
  if(pagoSection && pagoOpts){
    // Mostrar siempre los 8 grupos — los chips internos filtran por tcMax/debMax
    const grupos = Object.entries(FORMAS_PAGO);
    pagoOpts.innerHTML = `
      <div style="display:grid;grid-template-columns:repeat(4,1fr);gap:6px">
        ${grupos.map(([tipo,g])=>`
          <button class="cotiz-pago-btn" data-grupo="${tipo}"
            onclick="selGrupoAcept('${tipo}')"
            style="border:2px solid var(--border);border-radius:8px;padding:8px 6px;
                   background:var(--paper);cursor:pointer;text-align:center;
                   transition:border .15s,background .15s;width:100%">
            <div style="font-size:18px">${g.icon}</div>
            <div style="font-size:10px;font-weight:700;margin-top:2px;line-height:1.2">${g.label}</div>
          </button>`).join('')}
      </div>`;
    pagoSection.style.display = 'block';
  }

  // Confirmar deshabilitado hasta elegir forma de pago
  document.getElementById('btn-confirmar-acept').disabled = true;
}

function selGrupoAcept(tipo){
  window._cotizGrupoAcept    = tipo;
  cotizFormaPagoSeleccionada = null;
  cotizCuotasElegidasAcept   = null;

  // Resaltar grupo seleccionado
  document.querySelectorAll('.cotiz-pago-btn').forEach(btn=>{
    const sel = btn.dataset.grupo===tipo;
    btn.style.border     = sel ? '2px solid var(--green)' : '2px solid var(--border)';
    btn.style.background = sel ? '#d4edda' : 'var(--paper)';
    btn.style.fontWeight = sel ? '700' : '';
  });

  const grupo = FORMAS_PAGO[tipo];
  if(!grupo) return;

  const resultado    = window._cotizResActual||{};
  const total        = resultado.total||0;
  const cuotasPanel  = document.getElementById('cotiz-cuotas-panel');
  document.getElementById('btn-confirmar-acept').disabled = true;

  // Opción única — auto-seleccionar
  if(grupo.opciones.length===1){
    if(cuotasPanel){ cuotasPanel.innerHTML=''; cuotasPanel.style.display='none'; }
    selFormaAcept(grupo.opciones[0].val);
    return;
  }

  // Múltiples opciones — mostrar chips
  const cfg   = ASEGURADORAS[cotizAsegSeleccionada]||{};
  let opciones = grupo.opciones;
  // Filtrar TC por tcMax de aseguradora
  if(tipo==='TC'){
    const tcMax = cfg.tcMax || resultado.tcN || 12;
    opciones = opciones.filter(o=> _fpCuotas(o.val) <= tcMax);
  }
  // Filtrar Débito Produbanco por debMax
  if(tipo==='DEBITO_PRODUBANCO'){
    const debMax = cfg.debMax || resultado.debN || 12;
    opciones = opciones.filter(o=> _fpCuotas(o.val) <= debMax);
  }

  cuotasPanel.innerHTML=`
    <div style="margin-top:10px;padding:10px 12px;background:#f8f9fa;border-radius:8px;border:1px solid var(--border)">
      <div style="font-size:11px;font-weight:700;text-transform:uppercase;letter-spacing:.6px;color:var(--muted);margin-bottom:8px">
        Selecciona la forma exacta <span style="color:var(--red)">*</span>
      </div>
      <div style="display:flex;flex-wrap:wrap;gap:6px">
        ${opciones.map(o=>{
          const n=_fpCuotas(o.val);
          const cuotaStr=n>1&&total>0?` — $${(total/n).toFixed(2)}/mes`:(total>0?` — $${total.toFixed(2)}`:``);
          return `<button class="cotiz-cuota-chip" data-forma="${o.val}"
            onclick="selFormaAcept('${o.val.replace(/\\/g,'\\\\').replace(/'/g,"\\'")}') "
            style="border:2px solid var(--border);border-radius:6px;padding:6px 10px;
                   cursor:pointer;background:var(--paper);font-size:12px;font-weight:400;
                   transition:border .15s,background .15s;white-space:nowrap">
            ${o.label}${cuotaStr}
          </button>`;
        }).join('')}
      </div>
      <div id="cotiz-cuota-label" style="margin-top:8px;font-size:13px;font-weight:600;color:var(--green);display:none"></div>
    </div>`;
  cuotasPanel.style.display='';
}

function selFormaAcept(val){
  cotizFormaPagoSeleccionada = val;
  cotizCuotasElegidasAcept   = _fpCuotas(val);

  // Resaltar chip seleccionado
  document.querySelectorAll('.cotiz-cuota-chip').forEach(chip=>{
    const sel = chip.dataset.forma===val;
    chip.style.border      = sel ? '2px solid var(--green)' : '2px solid var(--border)';
    chip.style.background  = sel ? '#d4edda' : 'var(--paper)';
    chip.style.fontWeight  = sel ? '700' : '400';
  });

  // Actualizar etiqueta confirmación
  const total = (window._cotizResActual||{}).total||0;
  const n     = cotizCuotasElegidasAcept;
  const label = document.getElementById('cotiz-cuota-label');
  if(label){
    if(n>1&&total>0) label.textContent=`✓ ${val} × $${(total/n).toFixed(2)}/mes`;
    else if(total>0) label.textContent=`✓ ${val} — $${total.toFixed(2)}`;
    else             label.textContent=`✓ ${val}`;
    label.style.display='';
  }

  document.getElementById('btn-confirmar-acept').disabled = false;
}

function actualizarCuotasAcept(n){ /* reemplazado por selFormaAcept */ }

function confirmarAceptacion(){
  if(!cotizAceptarId || !cotizAsegSeleccionada){
    showToast('Selecciona una aseguradora primero','error'); return;
  }
  if(!cotizFormaPagoSeleccionada){
    showToast('Selecciona la forma de pago del cliente','error'); return;
  }
  const obs = (document.getElementById('cotiz-obs-acept').value||'').trim();
  const idStr = String(cotizAceptarId);

  // 1. Leer y actualizar cotización
  const all = _getCotizaciones();
  const idx = all.findIndex(c=>String(c.id)===idStr);
  if(idx < 0){
    showToast('Error: no se encontró la cotización ID=' + idStr,'error');
    console.error('IDs disponibles:', all.map(c=>c.id));
    return;
  }
  all[idx].asegElegida        = cotizAsegSeleccionada;
  all[idx].formaPagoElegida   = cotizFormaPagoSeleccionada;
  all[idx].nCuotasElegidas = cotizCuotasElegidasAcept || _fpCuotas(cotizFormaPagoSeleccionada) || 1;
  all[idx].obsAcept    = obs;
  all[idx].fechaAcept  = new Date().toISOString().split('T')[0];
  all[idx].estado      = 'EN EMISIÓN';
  all[idx]._dirty      = true;
  _saveCotizaciones(all);

  // 2. Actualizar cliente en cartera → EMISIÓN
  const cotiz = all[idx];
  const hoy = new Date().toISOString().split('T')[0];
  let clienteDB = null;
  if(cotiz.clienteId)  clienteDB = DB.find(x=>String(x.id)===String(cotiz.clienteId));
  if(!clienteDB && cotiz.clienteCI && cotiz.clienteCI.length>3)
    clienteDB = DB.find(x=>(x.ci||'')===cotiz.clienteCI);
  if(!clienteDB)
    clienteDB = DB.find(x=>x.nombre.trim().toUpperCase()===cotiz.clienteNombre.trim().toUpperCase());
  if(clienteDB){
    // Backfill: si la cotización no tenía clienteId vinculado, enlazarla ahora
    if(!all[idx].clienteId){
      all[idx].clienteId = String(clienteDB.id);
      all[idx].clienteCI = all[idx].clienteCI || clienteDB.ci || '';
      _saveCotizaciones(all);
    }
    clienteDB._dirty         = true;
    clienteDB.estado         = 'EMISIÓN';
    clienteDB.aseguradora    = cotizAsegSeleccionada.includes('SEGUROS') ? cotizAsegSeleccionada : cotizAsegSeleccionada + ' SEGUROS';
    clienteDB.ultimoContacto = hoy;
    const notaAcept = `Cotización aceptada — ${cotizAsegSeleccionada} — ${cotizFormaPagoSeleccionada||''}${obs?' · '+obs:''}`;
    _bitacoraAdd(clienteDB, notaAcept, 'cotizacion');
    saveDB();
  }

  // 3. Cerrar modal y actualizar UI
  closeModal('modal-cotiz-aceptar');
  actualizarBadgeCotizaciones();
  renderCotizaciones();
  renderDashboard();
  renderVencimientos();
  renderSeguimiento();

  const clienteMsg = clienteDB
    ? ' · ' + clienteDB.nombre.split(' ')[0] + ' → EMISIÓN'
    : ' (cliente no vinculado a cartera)';
  showToast('✅ ' + cotizAsegSeleccionada + ' confirmada' + clienteMsg, 'success');
}

// ── Ir a emisión (abre cierre de venta pre-llenado) ──────
function irAEmision(id){
  _resetCierreModal();
  const idStr = String(id);
  const all = _getCotizaciones();
  const cotiz = all.find(c=>String(c.id)===idStr);
  if(!cotiz){ showToast('Cotización no encontrada','error'); return; }
  if(!cotiz.asegElegida){ showToast('Sin aseguradora elegida','error'); return; }

  const r = (cotiz.resultados||[]).find(x=>x.name===cotiz.asegElegida)||{};

  // Buscar cliente
  let cliente = cotiz.clienteId ? DB.find(x=>String(x.id)===String(cotiz.clienteId)) : null;
  if(!cliente && cotiz.clienteCI && cotiz.clienteCI.length>3)
    cliente = DB.find(x=>(x.ci||'')===cotiz.clienteCI);
  if(!cliente)
    cliente = DB.find(x=>x.nombre.trim().toUpperCase()===cotiz.clienteNombre.trim().toUpperCase());

  // Resolver aseguradora en select
  const selAseg = document.getElementById('cv-nueva-aseg');
  let asegVal = 'OTRO';
  if(selAseg){
    const primer = cotiz.asegElegida.split(' ')[0].toUpperCase();
    const match  = Array.from(selAseg.options).find(o=>o.value.toUpperCase().includes(primer));
    asegVal = match ? match.value : 'OTRO';
    selAseg.value = asegVal;
  }

  // Setear cierreVentaData completo
  cierreVentaData = {
    clienteId:      cliente ? cliente.id : null,
    clienteNombre:  cotiz.clienteNombre,
    asegNombre:     asegVal,
    nTc:            r.tcN  || cotiz.cuotasTc  || 12,
    nDeb:           r.debN || cotiz.cuotasDeb || 10,
    cuotaTc:        r.tcCuota  || 0,
    cuotaDeb:       r.debCuota || 0,
    total:          r.total    || 0,
    pn:             r.pn       || 0,
    fromCotizacion: idStr,
  };

  // Llenar campos y calcular ANTES de abrir modal
  // Nombre del cliente (campo obligatorio — sin esto guardarCierreVenta guarda vacío)
  const clienteNombreResuelto = cotiz.clienteNombre || (cliente ? cliente.nombre : '');
  const clienteEl2 = document.getElementById('cv-cliente');
  if(clienteEl2) clienteEl2.value = clienteNombreResuelto;
  // Cabecera del modal
  const cvAsegEl = document.getElementById('cv-aseg');
  const cvTotalEl = document.getElementById('cv-total');
  if(cvAsegEl) cvAsegEl.textContent = cotiz.asegElegida;
  if(cvTotalEl) cvTotalEl.textContent = r.total > 0 ? `${fmt(r.total)} total` : 'Ingrese prima manualmente';
  // Limpiar campos que no deben heredar valores de usos anteriores del formulario
  ['cv-factura','cv-poliza','cv-observacion'].forEach(fid=>{ const el=document.getElementById(fid); if(el) el.value=''; });

  const pnEl    = document.getElementById('cv-pn');
  const desdeEl = document.getElementById('cv-desde');
  if(pnEl    && r.pn)        pnEl.value    = r.pn.toFixed(2);
  if(desdeEl && cotiz.desde) desdeEl.value = cotiz.desde;
  // Pre-fill AXA/VD from cotización
  const axavdEl=document.getElementById('cv-axavd');
  if(axavdEl){
    const hasAxa=cotiz.axaIncluido==='SI';
    const hasVida=(cotiz.vidaLatina||cotiz.vidaSweaden||cotiz.vidaMapfre||cotiz.vidaAlianza)>0;
    axavdEl.value=hasAxa&&hasVida?'AXA+VD':hasAxa?'AXA':hasVida?'VD':'';
  }
  // Pre-fill vida prima from cotización result
  const vidaVal=r.vida||0;
  const vidaEl=document.getElementById('cv-vida-prima');
  if(vidaEl&&vidaVal>0) vidaEl.value=vidaVal.toFixed(2);
  // Pre-fill previous policy from client
  const polAntEl=document.getElementById('cv-poliza-anterior');
  const asegAntEl=document.getElementById('cv-aseg-anterior');
  const clienteData=cotiz.clienteId?DB.find(x=>String(x.id)===String(cotiz.clienteId)):null;
  if(polAntEl&&clienteData) polAntEl.value=clienteData.polizaAnterior||clienteData.polizaNueva||clienteData.poliza||'';
  if(asegAntEl&&clienteData) asegAntEl.value=clienteData.aseguradoraAnterior||clienteData.aseguradora||'';
  // Pre-fill Valor Asegurado from cotización / cliente
  const vaVal=cotiz.va||(clienteData?clienteData.va:0)||0;
  const vaEl=document.getElementById('cv-va-cierre');
  const vaDisplayEl=document.getElementById('cv-va-display');
  if(vaEl) vaEl.value=vaVal||'';
  if(vaDisplayEl) vaDisplayEl.textContent=vaVal>0?`VA: ${fmt(vaVal)}`:'';
  // Pre-fill cuenta Produbanco desde el cliente
  const cuentaEl=document.getElementById('cv-cuenta');
  if(cuentaEl&&clienteData) cuentaEl.value=clienteData.cuentaBanc||clienteData.cuenta||'';
  // Store tasa for saving
  if(r.tasa) cierreVentaData.tasa=r.tasa;

  // Pre-cargar forma de pago que el cliente eligió al aceptar
  const fpElegida = cotiz.formaPagoElegida || '';
  const fpEl      = document.getElementById('cv-forma-pago');
  const tpEl      = document.getElementById('cv-tipo-pago');
  // Pre-fill dropdown exacto (cv-tipo-pago)
  if(fpElegida && tpEl){
    const matchOpt = Array.from(tpEl.options).find(o=>o.value===fpElegida);
    if(matchOpt) tpEl.value = fpElegida;
  }
  // Mapear valor exacto al tipo abstracto que usa cv-forma-pago y renderCvFormaPago
  if(fpEl){
    const fpTipo = _fpTipo(fpElegida);
    let fpAbstract = 'DEBITO_BANCARIO';
    if(fpTipo==='CONTADO'||fpTipo==='CHEQUES'||fpTipo==='PAGOS_DIRECTOS') fpAbstract='CONTADO';
    else if(fpTipo==='TC')            fpAbstract='TARJETA_CREDITO';
    else if(fpTipo==='DEBITO_REC_TC') fpAbstract='DEBITO_RECURRENTE_TC';
    else if(fpTipo==='MIXTO')         fpAbstract='MIXTO';
    // DEBITO_PRODUBANCO y DEBITO_OTROS → 'DEBITO_BANCARIO' (default)
    fpEl.value = fpAbstract;
    // Ajustar nTc/nDeb en cierreVentaData según lo elegido
    if(cotiz.nCuotasElegidas){
      if(fpTipo==='TC')                                          cierreVentaData.nTc  = cotiz.nCuotasElegidas;
      if(fpTipo==='DEBITO_PRODUBANCO'||fpTipo==='DEBITO_OTROS') cierreVentaData.nDeb = cotiz.nCuotasElegidas;
      if(fpTipo==='DEBITO_REC_TC')                              cierreVentaData.nDeb = cotiz.nCuotasElegidas;
    }
  }
  recalcDesglose();
  renderCvExtras();
  renderCvFormaPago();
  autocalcHasta();

  openModal('modal-cierre-venta');
  const fpLabel = fpElegida || '';
  showToast('📝 Cierre pre-llenado · ' + cotiz.asegElegida + (fpLabel?' · '+fpLabel:''), 'info');
}

// ── Ver detalle cotización ───────────────────────────────
function verDetalleCotiz(id){
  const all = _getCotizaciones();
  const c = all.find(x=>String(x.id)===String(id)); if(!c) return;
  const exec = USERS.find(u=>u.id===c.ejecutivo);
  let html = `<b>${c.clienteNombre}</b> · ${c.vehiculo} · VA $${Number(c.va).toLocaleString()}<br>`;
  html += `Fecha: ${c.fecha} · Vigencia: ${c.desde||'—'} → ${c.hasta||'—'}<br>`;
  html += `Ejecutivo: ${exec?.name||c.ejecutivo} · Estado: ${c.estado}${c.asegElegida?' · Elegida: <b>'+c.asegElegida+'</b>':''}<br>`;
  if(c.placa||c.color||c.motor||c.chasis){
    html += `<br><b>Vehículo:</b> Placa: ${c.placa||'—'} · Color: ${c.color||'—'}<br>`;
    if(c.motor)  html += `Motor: ${c.motor}<br>`;
    if(c.chasis) html += `Chasis: ${c.chasis}<br>`;
  }
  if(c.celular||c.correo) html += `<br>Contacto: ${c.celular||''} ${c.correo||''}<br>`;
  if(c.asegAnterior||c.polizaAnterior) html += `Póliza anterior: ${c.asegAnterior||''} ${c.polizaAnterior||''}<br>`;
  html += '<br><b>Opciones cotizadas:</b><br>';
  (c.resultados||[]).forEach(r=>{
    html+=`• ${r.name}: $${r.total?.toFixed(2)} | TC ${r.tcN}×$${r.tcCuota?.toFixed(2)} | Déb ${r.debN}×$${r.debCuota?.toFixed(2)}<br>`;
  });
  if(c.obsAcept) html+=`<br>Obs: ${c.obsAcept}`;
  alert(html);
}

// ── Re-imprimir PDF ──────────────────────────────────────
function reimprimirCotiz(id){
  const all = _getCotizaciones();
  const c = all.find(x=>String(x.id)===String(id)); if(!c) return;
  // Pre-llenar campos del cotizador y re-imprimir
  const parts = (c.vehiculo||'').split(' ');
  if(document.getElementById('cot-nombre')) document.getElementById('cot-nombre').value=c.clienteNombre;
  if(document.getElementById('cot-ci'))     document.getElementById('cot-ci').value=c.clienteCI||'';
  if(document.getElementById('cot-va'))     document.getElementById('cot-va').value=c.va||20000;
  if(document.getElementById('cot-desde'))  document.getElementById('cot-desde').value=c.desde||'';
  // Seleccionar aseguradoras de la cotización
  document.querySelectorAll('.aseg-check').forEach(cb=>{
    cb.checked = (c.aseguradoras||[]).includes(cb.dataset.aseg);
  });
  printCotizacion();
}

// ══════════════════════════════════════════════════════
//  COMPARATIVO
// ══════════════════════════════════════════════════════
function renderComparativo(){
  // Dinámico: usa datos reales de ASEGURADORAS
  const asegList = Object.entries(ASEGURADORAS);
  const headers = asegList.map(([name,cfg])=>`<th style="color:${cfg.color};font-size:11px;padding:6px 4px">${name}</th>`).join('');

  function row(label, fn, cls=''){
    const vals = asegList.map(([,cfg])=>{
      const v = fn(cfg);
      return `<td class="col-val ${cls}" style="font-size:10px">${v}</td>`;
    }).join('');
    return `<tr><td class="col-cob">${label}</td>${vals}</tr>`;
  }
  function yesNo(v){ return v==='SÍ'?`<span class="yes">SÍ</span>`:`<span class="no">${v}</span>`; }
  function money(v){ return v?`$${Number(v).toLocaleString('es-EC')}`:'-'; }

  // Actualizar headers de la tabla
  const thead = document.querySelector('#page-comparativo .comp-table thead tr');
  if(thead) thead.innerHTML=`<th class="col-cob" style="text-align:left">Cobertura</th>`+headers;

  document.getElementById('comp-tbody').innerHTML=[
    `<tr class="section-row"><td colspan="${asegList.length+1}">COBERTURAS BÁSICAS</td></tr>`,
    row('Todo Riesgo', ()=>`<span class="yes">SÍ</span>`),
    row('Pérdida parcial robo/daño', ()=>`<span class="yes">SÍ</span>`),
    row('Pérdida total robo/daño', ()=>`<span class="yes">SÍ</span>`),
    `<tr class="section-row"><td colspan="${asegList.length+1}">AMPAROS ADICIONALES</td></tr>`,
    row('Responsabilidad Civil', cfg=>money(cfg.resp_civil)),
    row('Muerte acc. ocupante', cfg=>money(cfg.muerte_ocupante)),
    row('Muerte acc. titular', cfg=>cfg.muerte_titular?money(cfg.muerte_titular):`<span class="no">N/A</span>`),
    row('Gastos Médicos/ocupante', cfg=>money(cfg.gastos_medicos)),
    row('Airbags / Extraterritorial / Wincha', ()=>'SÍ/SÍ/SÍ'),
    row('Amparo Patrimonial', cfg=>cfg.amparo),
    `<tr class="section-row"><td colspan="${asegList.length+1}">BENEFICIOS</td></tr>`,
    row('Auto sustituto', cfg=>cfg.auto_sust),
    row('Asist. Legal en situ', cfg=>yesNo(cfg.legal||'SÍ')),
    row('Asist. Exequial', cfg=>yesNo(cfg.exequial)),
    row('Vida/Muerte Accidental', cfg=>cfg.vida||'N/A'),
    row('Enfermedades Graves', cfg=>cfg.enf_graves||'N/A'),
    row('Renta Hospitalización', cfg=>cfg.renta_hosp||'N/A'),
    row('Gastos de Sepelio', cfg=>cfg.sepelio||'N/A'),
    row('Telemedicina', cfg=>cfg.telemedicina||'N/A'),
    row('Beneficio Dental', cfg=>cfg.dental||'N/A'),
    row('Médico a Domicilio', cfg=>cfg.medico_dom||'N/A'),
    `<tr class="section-row"><td colspan="${asegList.length+1}">PRIMA MÍNIMA / TC MAX / DÉB MAX</td></tr>`,
    row('Prima Neta Mínima', cfg=>cfg.pnMin>0?money(cfg.pnMin):'—'),
    row('TC Máx. cuotas', cfg=>cfg.tcMax+' cuotas'),
    row('Débito Máx. cuotas', cfg=>(cfg.debMax||10)+' cuotas'),
    `<tr class="section-row"><td colspan="${asegList.length+1}">DEDUCIBLES</td></tr>`,
    row('Pérd. parcial', cfg=>cfg.ded_parcial),
    row('Pérd. total daños', cfg=>cfg.ded_daño),
    row('Pérd. total robo s/disp.', cfg=>cfg.ded_robo_sin),
    row('Pérd. total robo c/disp.', cfg=>cfg.ded_robo_con),
  ].join('');
}

// ══════════════════════════════════════════════════════
//  ADMIN — EJECUTIVOS
// ══════════════════════════════════════════════════════
async function reconfigurarColumnasSP(){
  if(!confirm('¿Reconfigurar columnas en SharePoint?\n\nEsto creará columnas y listas que falten (no borra datos existentes).\nSe recargará la página al terminar.')) return;
  localStorage.removeItem('sp_cols_done');
  showToast('🔄 Reconfigurando columnas SP…', 'info');
  try{
    const logs = [];
    await spAsegurarColumnas(msg => { logs.push(msg); console.log('[SP cols]', msg); });
    localStorage.setItem('sp_cols_done','17');
    showToast('✅ Listas y columnas configuradas — recargando…', 'success');
    console.log('[SP setup]', logs.join('\n'));
    setTimeout(()=>location.reload(), 1500);
  }catch(e){
    showToast('⚠ Error al reconfigurar: ' + e.message, 'error');
    console.error('reconfigurarColumnasSP error:', e);
  }
}

function renderComisionesAdmin(){
  const el = document.getElementById('admin-comisiones-body'); if(!el) return;
  const comis = _getComisiones();
  const aseguradoras = Object.keys(COMISIONES_DEFAULT);
  el.innerHTML = `
    <div style="font-size:12px;color:var(--muted);margin-bottom:12px">
      Porcentaje de comisión sobre prima neta que recibe Reliance por cada aseguradora.
      Se usa para calcular la comisión estimada en cada cierre de venta.
    </div>
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;margin-bottom:14px">
      ${aseguradoras.map(a=>`
        <div class="form-group" style="margin:0">
          <label class="form-label" style="font-size:11px">${a}</label>
          <div style="display:flex;align-items:center;gap:6px">
            <input class="form-input" type="number" min="0" max="50" step="0.5"
              id="comis-${a.replace(/[^a-zA-Z0-9]/g,'_')}"
              value="${comis[a]||COMISIONES_DEFAULT[a]||0}"
              style="text-align:right;width:80px">
            <span style="font-size:13px;color:var(--muted)">%</span>
          </div>
        </div>`).join('')}
    </div>
    <button class="btn btn-primary w-full" onclick="guardarComisionesAdmin()">💾 Guardar comisiones</button>
  `;
}

function guardarComisionesAdmin(){
  const comis = {};
  Object.keys(COMISIONES_DEFAULT).forEach(a=>{
    const id = 'comis-'+a.replace(/[^a-zA-Z0-9]/g,'_');
    const el = document.getElementById(id);
    if(el) comis[a] = parseFloat(el.value)||0;
  });
  _saveComisiones(comis);
  showToast('✅ Comisiones actualizadas','success');
}

// ══════════════════════════════════════════════════════
//  ADMIN — TASAS POR ASEGURADORA (rangos de SA)
// ══════════════════════════════════════════════════════
function renderTasasAdmin(){
  const el = document.getElementById('admin-tasas-body'); if(!el) return;
  const v2   = _getTasasV2();
  const filas = Object.entries(v2).filter(([,row])=>row.aseg); // solo filas V2 válidas
  const th  = (txt,extra='') => `<th style="padding:6px 8px;border:1px solid var(--border,#ddd);white-space:nowrap;${extra}">${txt}</th>`;
  const inp = (val, id) =>
    `<input type="number" id="${id}" value="${val||''}" min="0" max="20" step="0.01"
       style="width:56px;border:1px solid #ccc;border-radius:4px;padding:2px 4px;font-size:12px;
              font-family:'DM Mono',monospace;text-align:right;font-weight:600">`;
  const inpL = (val, id) =>
    `<input type="number" id="${id}" value="${val||''}" placeholder="—" min="0" step="1000"
       style="width:68px;border:1px solid #ccc;border-radius:4px;padding:2px 4px;font-size:11px;
              font-family:'DM Mono',monospace;text-align:right;color:var(--muted)">`;
  el.innerHTML = `
    <p style="font-size:12px;color:var(--muted);margin-bottom:14px">
      Tasa por <b>aseguradora · región · tipo de póliza</b> con rangos de VA dinámicos.
      Cada fila define hasta 4 tasas y hasta 3 límites (en USD). El último rango no tiene techo.
    </p>
    <div style="overflow-x:auto;margin-bottom:16px">
      <table style="width:100%;border-collapse:collapse;font-size:12px;min-width:780px">
        <thead>
          <tr style="background:var(--bg2,#f5f5f5);text-align:center">
            ${th('Aseguradora','text-align:left')}
            ${th('Región','font-size:11px')}
            ${th('Tipo','font-size:11px')}
            ${th('Rango 1','background:#e8f0fe;font-size:11px')}
            ${th('Hasta $1','background:#e8f0fe;font-size:11px')}
            ${th('Rango 2','background:#e8f0fe;font-size:11px')}
            ${th('Hasta $2','background:#e8f0fe;font-size:11px')}
            ${th('Rango 3','background:#e8f0fe;font-size:11px')}
            ${th('Hasta $3','background:#e8f0fe;font-size:11px')}
            ${th('Rango 4','background:#e8f0fe;font-size:11px')}
          </tr>
          <tr style="background:var(--bg2,#f5f5f5);font-size:10px;color:var(--muted);text-align:center">
            <td colspan="3" style="border:1px solid var(--border,#ddd)"></td>
            <td style="padding:2px 8px;border:1px solid var(--border,#ddd);background:#e8f0fe">%</td>
            <td style="padding:2px 8px;border:1px solid var(--border,#ddd);background:#e8f0fe">USD techo</td>
            <td style="padding:2px 8px;border:1px solid var(--border,#ddd);background:#e8f0fe">%</td>
            <td style="padding:2px 8px;border:1px solid var(--border,#ddd);background:#e8f0fe">USD techo</td>
            <td style="padding:2px 8px;border:1px solid var(--border,#ddd);background:#e8f0fe">%</td>
            <td style="padding:2px 8px;border:1px solid var(--border,#ddd);background:#e8f0fe">USD techo</td>
            <td style="padding:2px 8px;border:1px solid var(--border,#ddd);background:#e8f0fe">%</td>
          </tr>
        </thead>
        <tbody>
          ${filas.map(([crmId, row])=>{
            const cfg = ASEGURADORAS[row.aseg]||{};
            const def = TASAS_V2_DEFAULT[crmId]||row;
            const t   = row.tasas  || [];
            const l   = row.limites|| [];
            const dt  = def.tasas  || [];
            const dl  = def.limites|| [];
            const s   = _safeName(crmId);
            const mod = t.some((v,i)=>Math.abs(v-(dt[i]||v))>0.00001) || l.some((v,i)=>Math.abs(v-(dl[i]||v))>0.00001);
            const td  = (content,bg='') => `<td style="padding:4px 6px;border:1px solid var(--border,#ddd);text-align:center${bg?';background:'+bg:''}">${content}</td>`;
            return `<tr${mod?' style="background:#fff8f0"':''}>
              <td style="padding:6px 10px;border:1px solid var(--border,#ddd);font-weight:700;color:${cfg.color||'var(--ink)'}">
                ${row.aseg}${mod?` <span style="font-size:9px;color:var(--accent,#c84b1a)">●</span>`:''}
              </td>
              <td style="padding:4px 6px;border:1px solid var(--border,#ddd);text-align:center;font-size:11px;color:var(--muted)">${row.region||'—'}</td>
              <td style="padding:4px 6px;border:1px solid var(--border,#ddd);text-align:center;font-size:11px;color:var(--muted)">${row.tipo||'—'}</td>
              ${td(`${inp(((t[0]||0)*100).toFixed(2), `adm-v2-${s}-t0`)} %`,'#e8f0fe')}
              ${td(inpL(l[0]||'', `adm-v2-${s}-l0`),'#e8f0fe')}
              ${td(`${inp(((t[1]||0)*100||''), `adm-v2-${s}-t1`)} %`,'#e8f0fe')}
              ${td(inpL(l[1]||'', `adm-v2-${s}-l1`),'#e8f0fe')}
              ${td(`${inp(((t[2]||0)*100||''), `adm-v2-${s}-t2`)} %`,'#e8f0fe')}
              ${td(inpL(l[2]||'', `adm-v2-${s}-l2`),'#e8f0fe')}
              ${td(`${inp(((t[3]||0)*100||''), `adm-v2-${s}-t3`)} %`,'#e8f0fe')}
            </tr>`;
          }).join('')}
        </tbody>
      </table>
    </div>
    <div style="display:flex;gap:8px;justify-content:flex-end">
      <button class="btn btn-ghost" onclick="restaurarTasasDefault()">↺ Restaurar defaults</button>
      <button class="btn btn-primary" onclick="guardarTasasAdmin()">💾 Guardar tasas</button>
    </div>`;
}

function guardarTasasAdmin(){
  const v2     = _getTasasV2();
  const nuevas = {};
  Object.entries(v2).filter(([,row])=>row.aseg).forEach(([crmId, row])=>{
    const s  = _safeName(crmId);
    const gT = i => { const e=document.getElementById(`adm-v2-${s}-t${i}`); return e?(parseFloat(e.value)||0)/100:0; };
    const gL = i => { const e=document.getElementById(`adm-v2-${s}-l${i}`); return e?(parseFloat(e.value)||0):0; };
    const limites = [gL(0),gL(1),gL(2)].filter(v=>v>0);
    const tasas   = [gT(0),gT(1),gT(2),gT(3)].slice(0, limites.length+1).filter(v=>v>0);
    nuevas[crmId] = {...row, tasas, limites};
  });
  _saveTasasV2(nuevas);
  showToast('✅ Tasas actualizadas','success');
  renderTasasAdmin();
  if(document.getElementById('page-tasas')?.classList.contains('active')) renderTasas();
}

function restaurarTasasDefault(){
  if(!confirm('¿Restaurar todas las tasas a los valores por defecto del sistema?')) return;
  localStorage.removeItem('_reliance_tasas_rangos');
  localStorage.removeItem('_reliance_tasas_v2');
  showToast('↺ Tasas restauradas a valores por defecto','info');
  renderTasasAdmin();
  if(document.getElementById('page-tasas')?.classList.contains('active')) renderTasas();
}

// ── Página "Tabla de Tasas" — muestra rangos V2 por aseguradora ──────────────
function renderTasas(){
  const el = document.getElementById('tasas-dinamicas'); if(!el) return;
  const v2 = _getTasasV2();
  // Agrupar filas por aseguradora
  const porAseg = {};
  Object.entries(v2).filter(([,r])=>r.aseg).forEach(([crmId,row])=>{
    if(!porAseg[row.aseg]) porAseg[row.aseg] = [];
    porAseg[row.aseg].push({crmId, ...row});
  });
  const principales = ['SWEADEN','MAPFRE','ALIANZA','ZURICH','LATINA','ASEG. DEL SUR','GENERALI'];
  el.innerHTML = principales.filter(n=>porAseg[n]).map(name=>{
    const cfg  = ASEGURADORAS[name]||{};
    const filas = porAseg[name];
    return `<div class="card">
      <div class="card-header" style="background:${cfg.color||'#1a4c84'}18">
        <div class="card-title" style="color:${cfg.color||'var(--ink)'}">${name}</div>
      </div>
      <div class="card-body" style="padding:10px 12px">
        ${filas.map(row=>{
          const label = [row.region, row.tipo].filter(Boolean).join(' · ') || 'Nacional';
          const rangos = row.tasas.map((t,i)=>{
            const desde = i===0 ? '$0' : `$${(row.limites[i-1]||0).toLocaleString()}`;
            const hasta = row.limites[i] ? `$${row.limites[i].toLocaleString()}` : '∞';
            return `<tr>
              <td style="padding:3px 0;color:var(--muted);font-size:11px">${desde} – ${hasta}</td>
              <td style="text-align:right;font-family:'DM Mono',monospace;font-weight:700;color:${cfg.color||'var(--ink)'}">
                ${(t*100).toFixed(2)}%
              </td>
            </tr>`;
          }).join('');
          return `<div style="margin-bottom:${filas.length>1?'10px':'0'}">
            ${filas.length>1?`<div style="font-size:10px;font-weight:600;color:var(--muted);margin-bottom:4px;text-transform:uppercase">${label}</div>`:''}
            <table style="width:100%;border-collapse:collapse;font-size:11px">${rangos}</table>
          </div>`;
        }).join('')}
        ${cfg.pnMin?`<div style="font-size:10px;color:var(--muted);margin-top:8px;border-top:1px solid var(--border,#eee);padding-top:6px">
          Prima mín: <b>$${cfg.pnMin}</b> · TC máx: <b>${cfg.tcMax} cuotas</b> · Débito: <b>${cfg.debMax||10} cuotas</b>
        </div>`:''}
      </div>
    </div>`;
  }).join('');
}

function renderAdmin(){
  const execs=USERS.filter(u=>u.rol==='ejecutivo');
  // Poblar select ejecutivos
  const execSelect=document.getElementById('import-assign-exec');
  if(execSelect) execSelect.innerHTML='<option value="">— Todos los ejecutivos —</option>'+execs.map(u=>`<option value="${u.id}">${u.name}</option>`).join('');
  // Mes actual por defecto
  const mesInput=document.getElementById('import-mes');
  if(mesInput&&!mesInput.value) mesInput.value=new Date().toISOString().substring(0,7);

  // Stats admin
  const statsEl=document.getElementById('admin-stats');
  if(statsEl){
    const totalClientes=DB.length;
    const venc30=DB.filter(c=>{const d=daysUntil(c.hasta);return d>=0&&d<=30}).length;
    const renovados=DB.filter(c=>c.estado==='RENOVADO').length;
    const cierres=_getCierres().length;
    statsEl.innerHTML=`
      <div class="stat-card"><div class="stat-label">Total clientes</div><div class="stat-value">${totalClientes}</div></div>
      <div class="stat-card"><div class="stat-label">Vencen en 30 días</div><div class="stat-value" style="color:var(--accent)">${venc30}</div></div>
      <div class="stat-card"><div class="stat-label">Renovados</div><div class="stat-value" style="color:var(--green)">${renovados}</div></div>
      <div class="stat-card"><div class="stat-label">Cierres registrados</div><div class="stat-value" style="color:var(--accent2)">${cierres}</div></div>`;
  }

  // Exec grid
  const execGrid=document.getElementById('exec-grid');
  if(execGrid) execGrid.innerHTML=execs.map(u=>{
    const mine=DB.filter(c=>c.ejecutivo===u.id);
    const venc30=mine.filter(c=>{const d=daysUntil(c.hasta);return d>=0&&d<=30}).length;
    const renovados=mine.filter(c=>(c.estado||'PENDIENTE')==='RENOVADO').length;
    return `<div class="exec-card">
      <div class="exec-avatar-lg" style="background:linear-gradient(135deg,${u.color},${u.color}88)">${u.initials}</div>
      <div style="font-weight:600;font-size:14px;margin-bottom:2px">${u.name}</div>
      <div style="font-size:11px;color:var(--muted);margin-bottom:14px">${u.email}</div>
      <div style="display:flex;gap:8px;margin-bottom:12px">
        <div class="exec-stat"><div class="exec-stat-val">${mine.length}</div><div class="exec-stat-key">Clientes</div></div>
        <div class="exec-stat"><div class="exec-stat-val text-accent">${venc30}</div><div class="exec-stat-key">Vencen 30d</div></div>
        <div class="exec-stat"><div class="exec-stat-val text-green">${renovados}</div><div class="exec-stat-key">Renovados</div></div>
      </div>
      <div style="display:flex;gap:6px">
        <button class="btn btn-ghost btn-sm" style="flex:1" onclick="showExecClientes('${u.id}')">Ver cartera →</button>
        <button class="btn btn-red btn-xs" onclick="confirmarEliminarEjecutivo('${u.id}')" title="Eliminar ejecutivo">✕</button>
      </div>
    </div>`;
  }).join('');

  // Resumen datos
  const dbStats=document.getElementById('admin-db-stats');
  if(dbStats) dbStats.innerHTML=`<div style="font-size:13px;line-height:2">
    📋 <b>${DB.length}</b> clientes en base de datos<br>
    👥 <b>${execs.length}</b> ejecutivos activos<br>
    ✅ <b>${_getCierres().length}</b> cierres registrados
  </div>`;

  // Resumen por ejecutivo (tab datos)
  const execResumen=document.getElementById('admin-exec-resumen');
  if(execResumen) execResumen.innerHTML=`<table style="width:100%;font-size:12px">
    <thead><tr>
      <th style="text-align:left;padding:6px;color:var(--muted)">Ejecutivo</th>
      <th style="padding:6px;color:var(--muted)">Clientes</th>
      <th style="padding:6px;color:var(--muted)">Pend.</th>
      <th style="padding:6px;color:var(--muted)">Renov.</th>
      <th style="padding:6px;color:var(--muted)">Venc. 30d</th>
    </tr></thead>
    <tbody>${execs.map(u=>{
      const mine=DB.filter(c=>c.ejecutivo===u.id);
      return `<tr style="border-bottom:1px solid var(--warm)">
        <td style="padding:6px;font-weight:500">${u.name}</td>
        <td style="padding:6px;text-align:center">${mine.length}</td>
        <td style="padding:6px;text-align:center">${mine.filter(c=>c.estado==='PENDIENTE').length}</td>
        <td style="padding:6px;text-align:center;color:var(--green)">${mine.filter(c=>c.estado==='RENOVADO').length}</td>
        <td style="padding:6px;text-align:center;color:var(--accent)">${mine.filter(c=>{const d=daysUntil(c.hasta);return d>=0&&d<=30}).length}</td>
      </tr>`;
    }).join('')}</tbody>
  </table>`;

  // Historial importaciones
  renderImportHistorial();
  renderComisionesAdmin();
  renderTasasAdmin();
}
function showAdminTab(tab, el){
  ['importar','historial','ejecutivos','datos','tasas'].forEach(t=>{
    const el2=document.getElementById('admin-tab-'+t);
    if(el2) el2.style.display=t===tab?'':'none';
  });
  document.querySelectorAll('#admin-tabs .pill').forEach(p=>p.classList.remove('active'));
  if(el) el.classList.add('active');
  if(tab==='datos' || tab==='tasas') renderAdmin(); // Refresh stats y tasas
}
function renderImportHistorial(){
  const historial=(_cache.historial||JSON.parse(localStorage.getItem('reliance_import_historial')||'[]'));
  const wrap=document.getElementById('import-historial-wrap'); if(!wrap) return;
  if(!historial.length){
    wrap.innerHTML='<div class="empty-state"><div class="empty-icon">📂</div><p>No hay cargas registradas aún</p></div>';
    return;
  }
  wrap.innerHTML=`<table style="width:100%;font-size:12px">
    <thead><tr>
      <th style="text-align:left;padding:6px 10px;color:var(--muted)">Fecha</th>
      <th style="text-align:left;padding:6px 10px;color:var(--muted)">Archivo</th>
      <th style="text-align:left;padding:6px 10px;color:var(--muted)">Mes cartera</th>
      <th style="text-align:left;padding:6px 10px;color:var(--muted)">Ejecutivo</th>
      <th style="text-align:left;padding:6px 10px;color:var(--muted)">Modo</th>
      <th style="padding:6px 10px;color:var(--muted)">Registros</th>
    </tr></thead>
    <tbody>${historial.slice().reverse().map(h=>`<tr style="border-bottom:1px solid var(--warm)">
      <td style="padding:6px 10px;font-family:'DM Mono',monospace;font-size:11px">${h.fecha}</td>
      <td style="padding:6px 10px;font-size:11px">${h.archivo}</td>
      <td style="padding:6px 10px">${h.mes||'—'}</td>
      <td style="padding:6px 10px">${h.ejecutivo||'Todos'}</td>
      <td style="padding:6px 10px"><span class="badge badge-blue" style="font-size:9px">${h.modo}</span></td>
      <td style="padding:6px 10px;text-align:center;font-weight:700;color:var(--green)">+${h.agregados}</td>
    </tr>`).join('')}</tbody>
  </table>`;
}
function confirmarEliminarEjecutivo(execId){
  const u=USERS.find(x=>String(x.id)===String(execId)); if(!u) return;
  const nClientes=DB.filter(c=>c.ejecutivo===execId).length;
  if(!confirm(`¿Eliminar al ejecutivo ${u.name}? Tiene ${nClientes} clientes asignados.`)) return;
  // Borrar en SharePoint si tiene _spId
  if(u._spId && _spReady) spDelete('usuarios', u._spId);
  const idx=USERS.findIndex(x=>String(x.id)===String(execId));
  if(idx>=0) USERS.splice(idx,1);
  saveUsers(); renderAdmin();
  showToast(`Ejecutivo ${u.name} eliminado`,'error');
}
function limpiarCarteraEjecutivo(){
  const execs=USERS.filter(u=>u.rol==='ejecutivo');
  const sel=prompt('¿Qué ejecutivo limpiar?\n'+execs.map((u,i)=>`${i+1}. ${u.name}`).join('\n')+'\n\nEscribe el nombre exacto:');
  const exec=execs.find(u=>u.name.toLowerCase()===sel?.toLowerCase());
  if(!exec){showToast('Ejecutivo no encontrado','error');return;}
  if(!confirm(`¿Eliminar TODOS los clientes de ${exec.name}? Esta acción no se puede deshacer.`)) return;
  const antes=DB.filter(c=>c.ejecutivo===exec.id).length;
  DB=DB.filter(c=>c.ejecutivo!==exec.id);
  saveDB(); renderAdmin();
  showToast(`${antes} clientes eliminados de ${exec.name}`,'error');
}
function exportarTodoExcel(){
  const esAdmin = currentUser?.rol==='admin';
  const fuente = esAdmin ? DB : myClientes();
  const label = esAdmin ? 'Clientes' : (currentUser?.name||'Mis_Clientes').replace(/\s+/g,'_');
  const data=fuente.map(c=>({
    'Ejecutivo':c.ejecutivo,'Nombre':c.nombre,'CI':c.ci,'Celular':c.celular,
    'Aseguradora':c.aseguradora,'Póliza':c.poliza,'Vigencia Desde':c.desde,'Vigencia Hasta':c.hasta,
    'VA':c.va,'Prima Neta':c.pn,'Estado':c.estado,'Marca':c.marca,'Modelo':c.modelo,
    'Año':c.anio,'Placa':c.placa,'Motor':c.motor,'Chasis':c.chasis,'Color':c.color,
    'Correo':c.correo,'Cuenta':c.cuenta,'OBS':c.obs,'Último Contacto':c.ultimoContacto
  }));
  const ws=XLSX.utils.json_to_sheet(data);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Clientes');
  XLSX.writeFile(wb,`Reliance_${label}_${new Date().toISOString().split('T')[0]}.xlsx`);
}
function exportarBackupJSON(){
  if(currentUser?.rol!=='admin'){ showToast('Solo el administrador puede generar backups','error'); return; }
  const backup={
    fecha:new Date().toISOString(),
    clientes:DB,
    cierres:_getCierres(),
    users:USERS
  };
  const blob=new Blob([JSON.stringify(backup,null,2)],{type:'application/json'});
  const a=document.createElement('a'); a.href=URL.createObjectURL(blob);
  a.download=`Reliance_Backup_${new Date().toISOString().split('T')[0]}.json`;
  a.click();
  showToast('Backup exportado');
}
function restaurarBackup(event){
  const file=event.target.files[0]; if(!file) return;
  const reader=new FileReader();
  reader.onload=function(e){
    try{
      const backup=JSON.parse(e.target.result);
      if(!confirm(`¿Restaurar backup del ${backup.fecha?.substring(0,10)}? Esto reemplazará los datos actuales.`)) return;
      if(backup.clientes) { DB=backup.clientes; saveDB(); }
      if(backup.cierres){ _saveCierres(backup.cierres); }
      renderDashboard(); renderAdmin();
      showToast('Backup restaurado exitosamente');
    }catch(err){ showToast('Archivo de backup inválido','error'); }
  };
  reader.readAsText(file);
}
function showExecClientes(execId){
  const u=USERS.find(x=>String(x.id)===String(execId));
  showToast(`Mostrando cartera de ${u?u.name:'—'}. Filtre en Clientes.`,'info');
  showPage('clientes');
}
function openModalNewExec(){openModal('modal-new-exec');}
function crearEjecutivo(){
  const nombre=document.getElementById('exec-nombre').value.trim();
  const email=document.getElementById('exec-user').value.trim();
  const pass=document.getElementById('exec-pass').value;
  const rol=document.getElementById('exec-rol').value;
  if(!nombre||!email||!pass){showToast('Completa todos los campos','error');return;}
  const id=email.split('@')[0].toLowerCase().replace(/[^a-z0-9]/g,'');
  const colors=['#c84b1a','#1a4c84','#2d6a4f','#b8860b','#6f42c1','#fd7e14','#dc3545','#17a2b8'];
  const initials=nombre.split(' ').map(w=>w[0]).join('').substring(0,2).toUpperCase();
  USERS.push({id,name:nombre,email,pass,rol,color:colors[USERS.length%colors.length],initials});
  saveUsers();
  closeModal('modal-new-exec');
  renderAdmin();
  showToast(`Ejecutivo ${nombre} creado`);
}

// ══════════════════════════════════════════════════════
//  EXCEL IMPORT
// ══════════════════════════════════════════════════════
function handleDragOver(e){e.preventDefault();document.getElementById('drop-zone').classList.add('drag-over');}
function handleDragLeave(e){document.getElementById('drop-zone').classList.remove('drag-over');}
function handleDrop(e){e.preventDefault();document.getElementById('drop-zone').classList.remove('drag-over');const f=e.dataTransfer.files[0];if(f) processExcel(f);}
function handleFileSelect(e){const f=e.target.files[0];if(f) processExcel(f);}
let importedFileName='';
function processExcel(file){
  importedFileName=file.name;
  const reader=new FileReader();
  reader.onload=function(e){
    const data=new Uint8Array(e.target.result);
    const wb=XLSX.read(data,{type:'array'});
    // Elegir hoja según selector
    const sheetPref=document.getElementById('import-sheet')?.value||'auto';
    let sheetName;
    if(sheetPref==='auto') sheetName=wb.SheetNames.find(s=>s.includes('VH')||s.includes('vh'))||wb.SheetNames[0];
    else sheetName=wb.SheetNames.find(s=>s===sheetPref)||wb.SheetNames[0];
    const ws=wb.Sheets[sheetName];
    const rawJson=XLSX.utils.sheet_to_json(ws,{defval:''});
    if(!rawJson.length){showToast('No se encontraron datos en el archivo','error');return;}
    // Normalizar nombres de columna (eliminar espacios extra en headers del Excel)
    const json=rawJson.map(row=>{
      const n={}; Object.entries(row).forEach(([k,v])=>{n[k.trim()]=v;}); return n;
    });
    // Filtrar filas con nombre
    importedRows=json.filter(r=>{
      const n=(r['Nombre Cliente']||r['NOMBRE']||r['nombre']||'').toString().trim();
      return n.length>2;
    });
    showImportPreview(importedRows, sheetName, file.name);
  };
  reader.readAsArrayBuffer(file);
}
function showImportPreview(rows, sheetName, fileName){
  const total = rows.length;
  const existingCIs = new Set(DB.map(c=>(c.ci||'').toString().trim()));

  // Mapear todas las filas con el nuevo helper
  const mapped = rows.map(r => _mapExcelRowToCliente(r));
  const nuevos = mapped.filter(m => m.ci && !existingCIs.has(m.ci)).length;
  const duplicados = total - nuevos;

  // Detectar campos ricos (Fase 2)
  const conFechaNac = mapped.filter(m=>m.fechaNac).length;
  const conPrestamo = mapped.filter(m=>m.prestamo||m.saldo).length;
  const conCelular2 = mapped.filter(m=>m.celular2).length;
  const produbanco  = mapped.filter(m=>m.tipoCliente==='PRODUBANCO').length;

  document.getElementById('import-preview-meta').textContent=`${total} registros encontrados en hoja "${sheetName}"`;

  const mapeoEl = document.getElementById('import-mapeo-resumen');
  if(mapeoEl) mapeoEl.innerHTML=`
    <div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(180px,1fr));gap:8px">
      <div style="padding:8px;background:#fff;border-radius:6px;border:1px solid var(--border)">
        📋 <b>${total}</b> filas totales
      </div>
      <div style="padding:8px;background:#d4edda;border-radius:6px;border:1px solid #2d6a4f">
        🆕 <b>${nuevos}</b> clientes nuevos
      </div>
      <div style="padding:8px;background:var(--warm);border-radius:6px;border:1px solid var(--border)">
        🔁 <b>${duplicados}</b> ya existen (por CI)
      </div>
      <div style="padding:8px;background:#e8f0fb;border-radius:6px;border:1px solid #1a4c84">
        🏦 <b>${produbanco}</b> tipo Produbanco
      </div>
      ${conFechaNac?`<div style="padding:8px;background:var(--warm);border-radius:6px;border:1px solid var(--border)">🎂 <b>${conFechaNac}</b> con fecha nacimiento</div>`:''}
      ${conPrestamo?`<div style="padding:8px;background:var(--warm);border-radius:6px;border:1px solid var(--border)">💳 <b>${conPrestamo}</b> con datos crédito</div>`:''}
      ${conCelular2?`<div style="padding:8px;background:var(--warm);border-radius:6px;border:1px solid var(--border)">📱 <b>${conCelular2}</b> con celular 2</div>`:''}
    </div>`;

  const modo = document.getElementById('import-modo')?.value||'agregar';
  const warnEl = document.getElementById('import-warning');
  if(warnEl){
    if(modo==='reemplazar') warnEl.innerHTML=`<span style="color:var(--red)">⚠ Modo REEMPLAZAR: se eliminarán todos los clientes del ejecutivo seleccionado</span>`;
    else if(modo==='actualizar') warnEl.innerHTML=`<span style="color:var(--accent)">ℹ Modo ACTUALIZAR: se sobreescribirán datos de clientes existentes</span>`;
    else warnEl.innerHTML='';
  }

  // Tabla preview con campos mapeados (no los headers crudos)
  const html=`<table><thead><tr>
    <th>Nombre</th><th>CI</th><th>Tipo</th><th>Celular</th>
    <th>Aseguradora</th><th>Placa</th><th>VA</th><th>Vig. Hasta</th>
    <th>Crédito</th><th>Estado</th>
  </tr></thead>
  <tbody>${mapped.slice(0,20).map(m=>{
    const esNuevo = m.ci && !existingCIs.has(m.ci);
    return `<tr style="${!esNuevo?'opacity:.7':''}">
      <td style="font-size:11px;max-width:140px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;font-weight:600">${m.nombre||'—'}</td>
      <td style="font-size:11px;font-family:'DM Mono',monospace">${m.ci||'—'}</td>
      <td style="font-size:10px"><span style="padding:2px 6px;border-radius:4px;background:${m.tipoCliente==='PRODUBANCO'?'#e8f0fb':'#f3f4f6'};color:${m.tipoCliente==='PRODUBANCO'?'#1a4c84':'#555'}">${m.tipoCliente}</span></td>
      <td style="font-size:11px">${m.celular||'—'}</td>
      <td style="font-size:11px;max-width:100px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${m.aseguradora||'—'}</td>
      <td style="font-size:11px;font-family:'DM Mono',monospace">${m.placa||'—'}</td>
      <td style="font-size:11px;text-align:right">${m.va?'$'+m.va.toLocaleString('es-EC'):'—'}</td>
      <td style="font-size:11px;font-family:'DM Mono',monospace">${m.hasta||'—'}</td>
      <td style="font-size:10px">${m.prestamo?'💳 '+m.prestamo:m.saldo?'$'+m.saldo:'—'}</td>
      <td>${esNuevo?'<span style="font-size:10px;color:var(--green);font-weight:700">✓ Nuevo</span>':'<span style="font-size:10px;color:var(--muted)">Existe</span>'}</td>
    </tr>`;
  }).join('')}${mapped.length>20?`<tr><td colspan="10" style="text-align:center;padding:8px;color:var(--muted);font-size:11px">... y ${mapped.length-20} registros más</td></tr>`:''}</tbody>
  </table>`;
  document.getElementById('import-table-wrap').innerHTML=html;
  document.getElementById('import-preview').style.display='block';
  showToast(`${total} registros detectados — ${nuevos} nuevos`,'info');
}
function confirmImport(){
  if(!importedRows.length) return;
  const execId = document.getElementById('import-assign-exec').value;
  const modo   = document.getElementById('import-modo')?.value||'agregar';
  const mes    = document.getElementById('import-mes')?.value||'';

  if(modo==='reemplazar' && execId){
    if(!confirm('¿Confirmar REEMPLAZAR toda la cartera del ejecutivo seleccionado?')) return;
    if(_spReady){
      const toDelSP = DB.filter(c=>String(c.ejecutivo)===String(execId)&&c._spId);
      toDelSP.forEach(c=>spDelete('clientes',c._spId));
    }
    DB = DB.filter(c=>String(c.ejecutivo)!==String(execId));
  }

  const existingCIs = new Set(DB.map(c=>(c.ci||'').toString().trim()));
  const maxId = DB.length ? Math.max(...DB.map(c=>c.id||0)) : 0;
  let added=0, updated=0, omitidos=0;

  importedRows.forEach((row, i) => {
    const m = _mapExcelRowToCliente(row);
    if(!m.nombre || m.nombre.length < 2) { omitidos++; return; }

    // MODO ACTUALIZAR — actualizar campos si el cliente ya existe por CI
    if(modo==='actualizar' && m.ci && existingCIs.has(m.ci)){
      const idx = DB.findIndex(c=>(c.ci||'').toString().trim()===m.ci);
      if(idx >= 0){
        DB[idx] = {
          ...DB[idx],
          // Datos de contacto
          celular:   m.celular  || DB[idx].celular,
          celular2:  m.celular2 || DB[idx].celular2,
          telFijo:   m.telFijo  || DB[idx].telFijo,
          correo:    m.correo   || DB[idx].correo,
          // Datos demográficos
          fechaNac:   m.fechaNac  || DB[idx].fechaNac,
          genero:     m.genero    || DB[idx].genero,
          estadoCivil:m.estadoCivil||DB[idx].estadoCivil,
          profesion:  m.profesion || DB[idx].profesion,
          direccionDom:m.direccionDom||DB[idx].direccionDom,
          // Crédito Produbanco
          cuentaBanc: m.cuentaBanc || DB[idx].cuentaBanc,
          prestamo:   m.prestamo   || DB[idx].prestamo,
          saldo:      m.saldo      || DB[idx].saldo,
          fechaVtoCred:m.fechaVtoCred||DB[idx].fechaVtoCred,
          // Póliza
          aseguradora:  m.aseguradora  || DB[idx].aseguradora,
          aseguradoraAnterior: m.aseguradoraAnterior || DB[idx].aseguradoraAnterior,
          polizaAnterior: m.polizaAnterior || DB[idx].polizaAnterior,
          va:   m.va   || DB[idx].va,
          dep:  m.dep  || DB[idx].dep,
          tasa: m.tasa || DB[idx].tasa,
          pn:   m.pn   || DB[idx].pn,
          desde:m.desde|| DB[idx].desde,
          hasta:m.hasta|| DB[idx].hasta,
          // Tipo cliente (solo si auto-detectado y el existente no tiene)
          tipoCliente: DB[idx].tipoCliente || m.tipoCliente,
          tasaAnterior: m.tasaAnterior || DB[idx].tasaAnterior,
          // Fase 3 — Datos adicionales Produbanco
          garantia:       m.garantia       || DB[idx].garantia,
          ramo:           m.ramo           || DB[idx].ramo,
          estadoCredito:  m.estadoCredito  || DB[idx].estadoCredito,
          fechaDesembolso:m.fechaDesembolso|| DB[idx].fechaDesembolso,
          monto:          m.monto          || DB[idx].monto,
          // Renovación
          tasaRenov:      m.tasaRenov      || DB[idx].tasaRenov,
          direccionOfi:   m.direccionOfi   || DB[idx].direccionOfi,
          _dirty: true,
        };
        updated++;
        return;
      }
    }

    // MODO AGREGAR — saltar si ya existe por CI
    if(modo==='agregar' && m.ci && existingCIs.has(m.ci)){ omitidos++; return; }

    // INSERTAR NUEVO CLIENTE
    const nuevo = {
      id: maxId + added + 1,
      ejecutivo: execId,
      tipo: 'RENOVACION',
      estado: 'PENDIENTE',
      nota: '', ultimoContacto: '', comentario: '', bitacora: [],
      // Campos mapeados del Excel
      nombre:   m.nombre,
      ci:       m.ci,
      tipoCliente: m.tipoCliente,
      celular:  m.celular,
      celular2: m.celular2,
      telFijo:  m.telFijo,
      correo:   m.correo,
      ciudad:   m.ciudad,
      region:   m.region,
      direccionDom: m.direccionDom,
      // Crédito
      cuentaBanc:   m.cuentaBanc,
      prestamo:     m.prestamo,
      saldo:        m.saldo,
      fechaVtoCred: m.fechaVtoCred,
      // Demografía
      fechaNac:   m.fechaNac,
      genero:     m.genero,
      estadoCivil:m.estadoCivil,
      profesion:  m.profesion,
      tasaAnterior: m.tasaAnterior,
      estadoGest:   m.estadoGest,
      // Vehículo
      marca: m.marca, modelo: m.modelo, anio: m.anio,
      motor: m.motor, chasis: m.chasis, color: m.color, placa: m.placa,
      // Póliza
      aseguradora:          m.aseguradora,
      aseguradoraAnterior:  m.aseguradoraAnterior,
      polizaAnterior:       m.polizaAnterior,
      poliza:               m.poliza,
      obs:                  m.obs,
      desde: m.desde, hasta: m.hasta,
      // Financiero
      va: m.va, dep: m.dep, tasa: m.tasa, pn: m.pn,
      // Fase 3 — Datos adicionales Produbanco
      garantia:       m.garantia,
      ramo:           m.ramo,
      estadoCredito:  m.estadoCredito,
      fechaDesembolso:m.fechaDesembolso,
      monto:          m.monto||null,
      // Renovación
      tasaRenov:      m.tasaRenov||null,
      direccionOfi:   m.direccionOfi,
      _dirty: true,
    };
    DB.push(nuevo);
    existingCIs.add(m.ci); // evitar duplicados dentro del mismo archivo
    _bitacoraAdd(nuevo, `Importado desde Excel — ${importedFileName||'archivo.xlsx'}`, 'sistema');
    added++;
  });

  saveDB();

  // Historial de importación
  const execUser = USERS.find(u=>u.id===execId);
  const historial = (_cache.historial||JSON.parse(localStorage.getItem('reliance_import_historial')||'[]'));
  historial.push({
    fecha: new Date().toISOString().split('T')[0],
    archivo: importedFileName||'archivo.xlsx',
    mes, ejecutivo: execUser?execUser.name:'Todos',
    modo, agregados: added, actualizados: updated,
  });
  _cache.historial = historial;
  localStorage.setItem('reliance_import_historial', JSON.stringify(historial));

  document.getElementById('import-preview').style.display='none';
  importedRows=[];
  renderDashboard(); renderAdmin();
  actualizarBadgeCobranza();

  const msg=`✓ ${added} importados${updated?' · '+updated+' actualizados':''}${omitidos?' · '+omitidos+' omitidos':''}${modo==='reemplazar'?' (cartera reemplazada)':''}`;
  showToast(msg, 'success');
}
function cancelImport(){document.getElementById('import-preview').style.display='none';importedRows=[];}

// ══════════════════════════════════════════════════════
//  IMPORT HELPERS — FASE 4
// ══════════════════════════════════════════════════════

// Convierte serial de fecha Excel o string a ISO yyyy-mm-dd
function _excelDateToISO(v){
  if(!v && v!==0) return '';
  const s = String(v).trim();
  if(!s) return '';
  // Ya es fecha ISO
  if(/^\d{4}-\d{2}-\d{2}/.test(s)) return s.slice(0,10);
  // Formato dd/mm/yyyy
  if(/^\d{1,2}\/\d{1,2}\/\d{4}/.test(s)){
    const [d,m,y]=s.split('/'); return `${y}-${m.padStart(2,'0')}-${d.padStart(2,'0')}`;
  }
  // Formato dd-mm-yyyy
  if(/^\d{1,2}-\d{1,2}-\d{4}/.test(s)){
    const [d,m,y]=s.split('-'); return `${y}-${m.padStart(2,'0')}-${d.padStart(2,'0')}`;
  }
  // Serial numérico de Excel (días desde 1900-01-01)
  const n = parseFloat(s);
  if(!isNaN(n) && n > 40000 && n < 60000){
    const d = new Date(Math.round((n - 25569) * 86400 * 1000));
    if(!isNaN(d)) return d.toISOString().split('T')[0];
  }
  return '';
}

// Lee un campo de una fila de Excel probando múltiples nombres de columna
function _excelCellStr(row, ...keys){
  for(const k of keys){
    const v = row[k];
    if(v !== undefined && v !== null && String(v).trim() !== '') return String(v).trim();
  }
  return '';
}

// Mapea una fila raw del Excel (PRODU VH) a un objeto cliente CRM
function _mapExcelRowToCliente(row){
  const str = (...k) => _excelCellStr(row, ...k);
  const num = (...k) => { const v=str(...k); return parseFloat(v.replace(',','.'))||0; };
  const date = (...k) => _excelDateToISO(str(...k));

  // Nombre — múltiples variantes de columna
  const nombre = str('Nombre Cliente','NOMBRE CLIENTE','Nombre','NOMBRE','nombre').toUpperCase();
  const ci     = str('CI','CEDULA','Cédula','C.I.','ci','cédula','IDENTIFICACION','RUC');
  // Celular: el archivo tiene "Celular 1" en col 9 (primario) y "Celular 1_1" (duplicado renombrado col 58)
  const celular= str('Celular 1','CELULAR 1','Celular','CELULAR','celular','Cel 1','TEL CELULAR','Celular 1_1');
  const celular2=str('Celular 2','CELULAR 2','cel2','celular2','Cel 2','CELULAR2','TELÉFONOS ADICIONALES');
  // Teléfono fijo: el archivo tiene "Teléfono fijo" (col 54) y "Teléfono fijo_1" (col 56, renombrado)
  const telFijo =str('Teléfono fijo','Tel Fijo','TEL FIJO','Teléfono Fijo','TELEFONO FIJO','telefono','telfijo','Teléfono fijo_1');
  const correo  =str('Correo','CORREO','Email','EMAIL','email','correo','E-MAIL');
  const ciudad  =str('Ciudad','CIUDAD','ciudad','CANTON');
  const region  =str('REGION','Region','región','region');
  // Dirección: el archivo usa "Direccion Dom"
  const dir     =str('Direccion Dom','Dirección Dom','DIRECCION DOM','Dirección','DIRECCION','Direccion','direccion','Domicilio','DOMICILIO');

  // Cuenta Produbanco — col J en PRODU VH, header varía
  const cuentaBanc = str('Cuenta Produbanco','CUENTA PRODUBANCO','Cuenta','CUENTA','cuenta','Nro Cuenta','NRO CUENTA','No Cuenta');

  // Crédito Produbanco — el archivo usa "Nº Préstamo" y "Fecha Vencimiento Crédito"
  const prestamo    = str('Nº Préstamo','Nº Prestamo','N° Préstamo','Préstamo','PRESTAMO','Prestamo','No Prestamo','NRO PRESTAMO','N Prestamo','prestamo','# Prestamo');
  const saldo       = num('Saldo','SALDO','Saldo Credito','SALDO CREDITO','saldo','Monto','MONTO');
  const fechaVtoCred= date('Fecha Vencimiento Crédito','Vencimiento Credito','Fecha Vto Credito','VTO CREDITO','FechaVtoCred','fechaVtoCred','Vto Credito');

  // Demografía — el archivo usa "FECHA_NACIMIENTO" y "ESTADO_CIVIL" con guiones bajos
  const fechaNac  = date('FECHA_NACIMIENTO','Fecha Nacimiento','FECHA NACIMIENTO','Nacimiento','F_NACIMIENTO','fechaNac','F. Nacimiento');
  const genero    = str('GENERO','Genero','Género','Sexo','SEXO','genero').toUpperCase();
  const estadoCivil=str('ESTADO_CIVIL','Estado Civil','ESTADO CIVIL','Civil','CIVIL','estadoCivil','E. Civil').toUpperCase();
  const profesion = str('PROFESION','Profesion','Profesión','Ocupacion','OCUPACION','profesion');
  const tasaAnterior=num('tasa VIG ANTERIOR','Tasa Anterior','TASA ANTERIOR','tasa presupuesto','tasaAnterior','Tasa Vig Ant','TASA VIG ANT');

  // Vehículo — el archivo tiene "Año " con espacio (ya normalizado al trim del key)
  const marca  = str('Marca','MARCA','marca');
  const modelo = str('Modelo','MODELO','modelo');
  const anio   = parseInt(str('Año','AÑO','anio','ANO','Modelo Año','AO'))||2024;
  const motor  = str('Motor','MOTOR','motor','N Motor','N. Motor');
  const chasis = str('No De Chasis','N De Chasis','CHASIS','Chasis','chasis','N° Chasis');
  const color  = str('Color','COLOR','color');
  const placa  = str('Placa','PLACA','placa');

  // Póliza — el archivo usa "POLIZA ACTUAL" y "ASEGURADORA ACTUAL"
  const polizaAnterior = str('POLIZA ACTUAL','POLIZA','Poliza','Póliza','poliza','Poliza Anterior','POLIZA ANTERIOR');
  const aseguradora    = str('ASEGURADORA ACTUAL','ASEGURADORA','Aseguradora','aseguradora','ASEG');
  const aseguradoraAnterior = aseguradora; // En PRODU VH la aseguradora es la anterior

  // Vigencia
  const desde = date('Fc_desde ultima vigencia','Vigencia Desde','VIGENCIA DESDE','desde','DESDE','Fc Desde');
  const hasta = date('Fc_hasta ultima vigencia','Vigencia Hasta','VIGENCIA HASTA','hasta','HASTA','Fc Hasta');

  // Valores financieros — el archivo tiene "Ultimo Val_Aseg." y "PRIMA NETA" (col AK, la real)
  const va  = num('Ultimo Val_Aseg.','Valor asegurado','Valor Asegurado','VALOR ASEGURADO','VA','va','Val Aseg');
  const dep = num('VALOR AUTO DEPRECIADO','v dep','Depreciacion','DEP','dep');
  const pn  = num('PRIMA NETA','Prima Neta','pn','PN');   // PRIMA NETA col AK tiene los valores reales
  const tasa= num('tasa presupuesto','Tasa','TASA','tasa')||tasaAnterior||null;
  const obs = str('OBS POLIZA','OBS','obs','Obs','observacion','OBSERVACION')||'RENOVACION';

  // Estado de gestión
  const estadoGest = str('Estado Gestion','ESTADO GESTION','Estado','ESTADO GESTION','estadoGest','K');

  // Datos adicionales Produbanco (Fase 3 — Excel: PRODU VH)
  const garantia        = str('Garantia','GARANTIA','garantia');                             // col C
  const ramo            = str('RAMO','Ramo','ramo');                                         // col AD
  const estadoCredito   = str('Estado','ESTADO','estadoCredito');                            // col AW
  const fechaDesembolso = date('Fecha Desembolso','FECHA DESEMBOLSO','fechaDesembolso');     // col AX
  const monto           = num('Monto','MONTO','monto');                                      // col AY
  // Renovación — tasas de referencia y dirección oficina
  const tasaRenov       = num('tasa renov aseg actual','TASA RENOV ASEG ACTUAL','tasaRenov'); // col AN
  const direccionOfi    = str('Direccion Ofi','Dirección Ofi','DIRECCION OFI','direccionOfi');// col BD

  // Auto-detectar tipoCliente
  let tipoCliente = str('Tipo Cliente','TIPO CLIENTE','tipoCliente');
  if(!tipoCliente){
    tipoCliente = (cuentaBanc||prestamo) ? 'PRODUBANCO' : 'PARTICULAR';
  }

  return {
    nombre, ci, celular, celular2, telFijo, correo,
    ciudad: ciudad||'QUITO', region: region || _ciudadToRegion(ciudad) || 'SIERRA',
    direccionDom: dir,
    tipoCliente, cuentaBanc, prestamo,
    saldo, fechaVtoCred,
    fechaNac, genero, estadoCivil, profesion,
    tasaAnterior: tasaAnterior||null,
    estadoGest,
    marca, modelo, anio, motor, chasis, color, placa,
    aseguradora, aseguradoraAnterior, polizaAnterior,
    poliza: str('POLIZA RENOVADA','poliza_nueva','Poliza Nueva','POLIZA NUEVA')||'',
    desde, hasta,
    va, dep, tasa: tasa||null, pn,
    obs,
    // Fase 3 — Datos adicionales Produbanco
    garantia, ramo, estadoCredito, fechaDesembolso,
    monto: monto||null,
    // Renovación
    tasaRenov: tasaRenov||null,
    direccionOfi,
  };
}

// ══════════════════════════════════════════════════════
//  NUEVO CLIENTE
// ══════════════════════════════════════════════════════
function guardarNuevoCliente(){
  const nombre=document.getElementById('nc-nombre').value.trim();
  const ci=document.getElementById('nc-ci').value.trim();
  if(!nombre||!ci){showToast('Nombre y CI son requeridos','error');return;}
  const maxId=DB.length?Math.max(...DB.map(c=>c.id||0)):0;
  DB.push({
    id:maxId+1,ejecutivo:currentUser?.id||'',
    nombre:nombre.toUpperCase(),ci,tipo:document.getElementById('nc-tipo').value,
    tipoCliente:document.getElementById('nc-tipo-cliente')?.value||'',
    region:document.getElementById('nc-region').value,ciudad:document.getElementById('nc-ciudad').value,
    obs:document.getElementById('nc-obs').value,celular:document.getElementById('nc-cel').value,
    celular2:document.getElementById('nc-cel2')?.value||'',
    telFijo:document.getElementById('nc-tel-fijo')?.value||'',
    correo:document.getElementById('nc-email').value,aseguradora:document.getElementById('nc-aseg').value,
    va:parseFloat(document.getElementById('nc-va').value)||0,
    tasa:parseFloat(document.getElementById('nc-tasa').value)||null,
    marca:document.getElementById('nc-marca').value,modelo:document.getElementById('nc-modelo').value,
    anio:parseInt(document.getElementById('nc-anio').value)||2025,placa:document.getElementById('nc-placa').value,
    chasis:document.getElementById('nc-chasis')?.value||'',
    color:document.getElementById('nc-color')?.value||'',
    desde:document.getElementById('nc-desde').value,hasta:document.getElementById('nc-hasta').value,
    comentario:document.getElementById('nc-comentario').value,
    fechaNac:document.getElementById('nc-nacimiento')?.value||'',
    genero:document.getElementById('nc-genero')?.value||'',
    estadoCivil:document.getElementById('nc-civil')?.value||'',
    profesion:document.getElementById('nc-profesion')?.value||'',
    direccionDom:document.getElementById('nc-dir')?.value||'',
    cuentaBanc:document.getElementById('nc-cuenta')?.value||'',
    prestamo:document.getElementById('nc-prestamo')?.value||'',
    saldo:parseFloat(document.getElementById('nc-saldo')?.value)||0,
    fechaVtoCred:document.getElementById('nc-vto-cred')?.value||'',
    dep:0,poliza:'',estado:'PENDIENTE',nota:'',ultimoContacto:''
  });
  const nuevoC = DB[DB.length-1];
  if(nuevoC) _bitacoraAdd(nuevoC, 'Cliente creado en RelianceDesk', 'sistema');
  saveDB();
  showToast(`Cliente ${nombre.split(' ')[0]} registrado`);
  setTimeout(()=>showPage('clientes'),500);
}

// ══════════════════════════════════════════════════════
//  EXPORT EXCEL
// ══════════════════════════════════════════════════════
function exportToExcel(){
  const data=clientesFiltrados.map(c=>({
    'Nombre':c.nombre,'CI':c.ci,'Tipo':c.tipo,'Aseguradora':c.aseguradora,
    'Marca':c.marca,'Modelo':c.modelo,'Año':c.anio,'Placa':c.placa,
    'Ciudad':c.ciudad,'Región':c.region,'Celular':c.celular,'Correo':c.correo,
    'Val.Asegurado':c.va,'Tasa':c.tasa,'Prima Neta':c.pn,
    'Vigencia Desde':c.desde,'Vigencia Hasta':c.hasta,
    'Estado':c.estado||'PENDIENTE','Nota':c.nota||'',
    'OBS':c.obs,'Póliza':c.poliza
  }));
  const ws=XLSX.utils.json_to_sheet(data);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Clientes');
  XLSX.writeFile(wb,`Reliance_Cartera_${new Date().toISOString().split('T')[0]}.xlsx`);
  showToast(`${data.length} clientes exportados`);
}

// ══════════════════════════════════════════════════════
//  CIERRES DE VENTA
// ══════════════════════════════════════════════════════
function renderCierres(){
  let cierres=_getCierres();
  if(currentUser && currentUser.rol!=='admin') cierres=cierres.filter(c=>String(c.ejecutivo)===String(currentUser.id));
  cierres.sort((a,b)=>new Date(b.fechaRegistro)-new Date(a.fechaRegistro));

  // Poblar filtro de meses
  const mesesSel=document.getElementById('cierres-filter-mes');
  if(mesesSel&&mesesSel.options.length<=1){
    const meses=[...new Set(cierres.map(c=>(c.fechaRegistro||'').substring(0,7)))].sort().reverse();
    meses.forEach(m=>{ const o=document.createElement('option'); o.value=m; o.textContent=m; mesesSel.appendChild(o); });
  }

  // Aplicar filtros
  const search=(document.getElementById('cierres-search')?.value||'').toLowerCase();
  const fAseg=(document.getElementById('cierres-filter-aseg')?.value||'').toUpperCase();
  const fFp=document.getElementById('cierres-filter-fp')?.value||'';
  const fMes=document.getElementById('cierres-filter-mes')?.value||'';
  let filtered=cierres.filter(c=>{
    if(search && !((c.clienteNombre||'').toLowerCase().includes(search)||(c.polizaNueva||'').toLowerCase().includes(search)||(c.facturaAseg||'').toLowerCase().includes(search))) return false;
    if(fAseg && !(c.aseguradora||'').toUpperCase().includes(fAseg)) return false;
    if(fFp && (c.formaPago?.forma||'')!==fFp) return false;
    if(fMes && !(c.fechaRegistro||'').startsWith(fMes)) return false;
    return true;
  });

  // Stats
  const totalPrimas=filtered.reduce((s,c)=>s+(c.primaTotal||0),0);
  const byFp={DEBITO_BANCARIO:0,TARJETA_CREDITO:0,CONTADO:0,MIXTO:0};
  filtered.forEach(c=>{ if(c.formaPago?.forma) byFp[c.formaPago.forma]=(byFp[c.formaPago.forma]||0)+1; });
  const statsEl=document.getElementById('cierres-stats');
  if(statsEl) statsEl.innerHTML=`
    <div class="stat-card"><div class="stat-label">Total cierres</div><div class="stat-value">${filtered.length}</div></div>
    <div class="stat-card"><div class="stat-label">Prima total recaudada</div><div class="stat-value" style="color:var(--green);font-size:18px">${fmt(totalPrimas)}</div></div>
    <div class="stat-card"><div class="stat-label">Débito / TC</div><div class="stat-value">${byFp.DEBITO_BANCARIO} / ${byFp.TARJETA_CREDITO}</div></div>
    <div class="stat-card"><div class="stat-label">Contado / Mixto</div><div class="stat-value">${byFp.CONTADO} / ${byFp.MIXTO}</div></div>`;

  document.getElementById('cierres-count').textContent=`${filtered.length} de ${cierres.length} cierres`;
  document.getElementById('badge-cierres').textContent=cierres.length;

  const fpLabel={DEBITO_BANCARIO:'🏦 Débito',TARJETA_CREDITO:'💳 TC',CONTADO:'💵 Contado',MIXTO:'🔀 Mixto'};
  document.getElementById('cierres-tbody').innerHTML=filtered.map((c,i)=>{
    const fp=c.formaPago||{};
    let fpDetail='',bancoCuenta='—';
    if(fp.forma==='DEBITO_BANCARIO'){
      fpDetail=`${fp.nCuotas} cuotas de ${fmt(fp.cuotaMonto)}<br><span style="color:var(--muted)">1ª: ${fp.fechaPrimera||'—'}</span>`;
      bancoCuenta=`<span style="font-size:11px">${fp.banco||'—'}</span><br><span class="mono" style="font-size:10px;color:var(--muted)">${fp.cuenta||c.cuenta||'—'}</span>`;
    } else if(fp.forma==='TARJETA_CREDITO'){
      fpDetail=`${fp.nCuotas} cuotas de ${fmt(fp.cuotaMonto)}<br><span style="color:var(--muted)">Contacto: ${fp.fechaContacto||'—'}</span>`;
      bancoCuenta=`<span style="font-size:11px">${fp.banco||'—'}</span>${fp.digitos?`<br><span class="mono" style="font-size:10px;color:var(--muted)">·····${fp.digitos}</span>`:''}`;
    } else if(fp.forma==='CONTADO'){
      fpDetail=`Cobro: ${fp.fechaCobro||'—'}`;
      if(fp.referencia) fpDetail+=`<br><span style="color:var(--muted);font-size:10px">Ref: ${fp.referencia}</span>`;
    } else if(fp.forma==='MIXTO'){
      fpDetail=`Inicial: ${fmt(fp.montoInicial)}<br><span style="color:var(--muted)">Resto ${fp.nCuotasResto}c. ${fp.metodoResto}</span>`;
      bancoCuenta=`<span style="font-size:11px">${fp.banco||'—'}</span>`;
    }
    const axavdBadge=c.axavd?`<span class="badge" style="background:#6b5b95;color:#fff;font-size:9px">${c.axavd}</span>`:'<span style="color:var(--muted);font-size:11px">—</span>';
    return `<tr>
      <td><span class="mono" style="font-size:11px">${c.fechaRegistro||'—'}</span></td>
      <td><span style="font-weight:500;font-size:12px">${c.clienteNombre||'—'}</span></td>
      <td><span class="mono" style="font-size:11px">${c._placa||'—'}</span></td>
      <td><span style="font-size:11px">${c.aseguradora||'—'}</span></td>
      <td><span class="mono" style="font-size:10px">${c.polizaNueva||'—'}</span></td>
      <td><span class="mono" style="font-size:10px">${c.facturaAseg||'—'}</span></td>
      <td class="mono" style="font-weight:700;color:var(--green)">${fmt(c.primaTotal)}</td>
      <td style="font-size:11px"><span class="badge badge-blue">${fpLabel[fp.forma]||fp.forma||'—'}</span><br><span style="font-size:10px;color:var(--muted)">${fpDetail}</span></td>
      <td style="font-size:11px">${bancoCuenta}</td>
      <td><span class="mono" style="font-size:10px">${c.vigDesde||'—'}</span><br><span class="mono" style="font-size:10px;color:var(--muted)">${c.vigHasta||'—'}</span></td>
      <td>${axavdBadge}</td>
      <td><div style="display:flex;gap:4px">
        <button class="btn btn-ghost btn-xs" onclick="verDetalleCierre(${i})">👁 Ver</button>
        <button class="btn btn-blue btn-xs" onclick="editarCierre('${c.id}')">✏ Editar</button>
        <button class="btn btn-red btn-xs" onclick="eliminarCierre(${i})">✕</button>
      </div></td>
    </tr>`;
  }).join('')||'<tr><td colspan="11"><div class="empty-state"><div class="empty-icon">📋</div><p>No hay cierres registrados aún</p></div></td></tr>';
}
function verDetalleCierre(idx){
  let cierres=_getCierres();
  if(currentUser&&currentUser.rol!=='admin') cierres=cierres.filter(c=>String(c.ejecutivo)===String(currentUser.id));
  const c=cierres.sort((a,b)=>new Date(b.fechaRegistro)-new Date(a.fechaRegistro))[idx];
  if(!c) return;
  const fp=c.formaPago||{};
  let fpHtml='';
  if(fp.forma==='DEBITO_BANCARIO'&&fp.calendario){
    fpHtml='<div class="section-divider">Calendario de débitos</div><table style="width:100%;font-size:11px"><thead><tr><th style="text-align:left;padding:4px;color:var(--muted)">Cuota</th><th style="text-align:left;padding:4px;color:var(--muted)">Fecha</th><th style="text-align:right;padding:4px;color:var(--muted)">Monto</th></tr></thead><tbody>'+
      fp.calendario.map((d,i)=>`<tr style="border-bottom:1px solid var(--warm)"><td style="padding:4px">${i+1}</td><td style="padding:4px;font-family:'DM Mono',monospace">${d}</td><td style="padding:4px;text-align:right;font-family:'DM Mono',monospace">${fmt(fp.cuotaMonto)}</td></tr>`).join('')+'</tbody></table>';
  } else if(fp.forma==='TARJETA_CREDITO'){
    fpHtml=`<div class="highlight-card" style="background:#e8f0fb;border-color:var(--accent2)">
      <div><b>TC — ${fp.banco||'—'} ${fp.digitos?'·····'+fp.digitos:''}</b></div>
      <div>${fp.nCuotas} cuotas de ${fmt(fp.cuotaMonto)} · Fecha contacto: <b>${fp.fechaContacto||'—'}</b></div></div>`;
  } else if(fp.forma==='CONTADO'){
    fpHtml=`<div class="highlight-card" style="background:#d4edda;border-color:var(--green)"><b>Contado — Cobro: ${fp.fechaCobro||'—'}</b></div>`;
  } else if(fp.forma==='MIXTO'){
    fpHtml=`<div class="detail-row"><span class="detail-key">Cuota inicial</span><span class="detail-val">${fmt(fp.montoInicial)} el ${fp.fechaInicial||'—'}</span></div>
    <div class="detail-row"><span class="detail-key">Resto</span><span class="detail-val">${fp.nCuotasResto} cuotas via ${fp.metodoResto} desde ${fp.fechaCuotaResto||'—'}</span></div>`;
  }
  const exec=USERS.find(u=>u.id===c.ejecutivo);
  document.getElementById('modal-cliente-title').textContent='Cierre: '+c.clienteNombre;
  document.getElementById('modal-cliente-body').innerHTML=`
    <div class="grid-2">
      <div class="card"><div class="card-body">
        <div class="detail-section"><div class="detail-section-title">Datos del Cierre</div>
          <div class="detail-row"><span class="detail-key">Fecha registro</span><span class="detail-val mono">${c.fechaRegistro}</span></div>
          <div class="detail-row"><span class="detail-key">Ejecutivo</span><span class="detail-val">${exec?exec.name:'—'}</span></div>
          <div class="detail-row"><span class="detail-key">Aseguradora</span><span class="detail-val font-bold">${c.aseguradora}</span></div>
          <div class="detail-row"><span class="detail-key">N° Póliza</span><span class="detail-val mono">${c.polizaNueva}</span></div>
          <div class="detail-row"><span class="detail-key">N° Factura</span><span class="detail-val mono">${c.facturaAseg}</span></div>
          <div class="detail-row"><span class="detail-key">Vigencia</span><span class="detail-val mono">${c.vigDesde} → ${c.vigHasta}</span></div>
        </div>
        <div class="detail-section"><div class="detail-section-title">Prima</div>
          <div class="detail-row"><span class="detail-key">Prima Neta</span><span class="detail-val mono">${fmt(c.primaNeta)}</span></div>
          <div class="detail-row"><span class="detail-key">Prima Total</span><span class="detail-val mono font-bold text-green">${fmt(c.primaTotal)}</span></div>
        </div>
        ${c.observacion?`<div style="padding:10px;background:var(--warm);border-radius:6px;font-size:12px">💬 ${c.observacion}</div>`:''}
      </div></div>
      <div class="card"><div class="card-body">
        <div class="detail-section"><div class="detail-section-title">Forma de Pago</div>${fpHtml}</div>
      </div></div>
    </div>`;
  document.getElementById('modal-btn-cotizar').style.display='none';
  document.getElementById('modal-btn-eliminar').style.display='none';
  // Botón editar reutilizado para cierre
  const btnEd=document.getElementById('modal-btn-editar');
  btnEd.style.display='';
  btnEd.textContent='✏ Editar Cierre';
  btnEd.onclick=()=>{ closeModal('modal-cliente'); editarCierre(c.id); };
  openModal('modal-cliente');
}
function editarCierre(cierreId){
  const allCierres=_getCierres();
  const c=allCierres.find(x=>String(x.id)===String(cierreId)); if(!c) return;
  const fp=c.formaPago||{};
  // Marcar modo edición
  cierreVentaData={
    asegNombre:c.aseguradora, total:c.primaTotal, pn:c.primaNeta,
    cuotaTc:fp.cuotaMonto||0, cuotaDeb:fp.cuotaMonto||0,
    nTc:fp.nCuotas||12, nDeb:fp.nCuotas||10,
    clienteId:null, editandoCierreId:cierreId
  };
  // Cabecera
  document.getElementById('cv-aseg').textContent='✏ EDITANDO CIERRE';
  document.getElementById('cv-total').textContent=`${fmt(c.primaTotal)} · ${c.clienteNombre}`;
  // Campos básicos
  document.getElementById('cv-cliente').value=c.clienteNombre;
  // Aseguradora select — normalizar nombre para match
  const selAseg=document.getElementById('cv-nueva-aseg');
  const asegOpts=Array.from(selAseg.options).map(o=>o.value);
  // Buscar match exacto primero, luego parcial
  const asegMatch=asegOpts.find(o=>o===c.aseguradora) ||
    asegOpts.find(o=>o!==''&&o!=='OTRO'&&c.aseguradora&&c.aseguradora.toUpperCase().includes(o.split(' ')[0]));
  selAseg.value=asegMatch||'OTRO';
  // Factura con máscara
  const factInput=document.getElementById('cv-factura');
  factInput.value=c.facturaAseg||'';
  document.getElementById('cv-poliza').value=c.polizaNueva||'';
  document.getElementById('cv-desde').value=c.vigDesde||'';
  document.getElementById('cv-hasta').value=c.vigHasta||'';
  document.getElementById('cv-pn').value=c.primaNeta||'';
  document.getElementById('cv-total-val').value=c.primaTotal||'';
  if(document.getElementById('cv-cuenta')) document.getElementById('cv-cuenta').value=c.cuenta||fp.cuenta||'';
  if(document.getElementById('cv-axavd')) document.getElementById('cv-axavd').value=c.axavd||'';
  document.getElementById('cv-observacion').value=c.observacion||'';
  // Pre-llenar datos Vida/AP si existen
  recalcDesglose();
  renderCvExtras();
  setTimeout(()=>{
    if(document.getElementById('cv-poliza-vida'))  document.getElementById('cv-poliza-vida').value  = c.poliza_vida  || '';
    if(document.getElementById('cv-factura-vida')) document.getElementById('cv-factura-vida').value = c.factura_vida || '';
    if(document.getElementById('cv-total-vida'))   document.getElementById('cv-total-vida').value   = c.total_vida   || '';
  }, 50);
  // Forma de pago
  document.getElementById('cv-forma-pago').value=fp.forma||'DEBITO_BANCARIO';
  // Marcar pill activa
  document.querySelectorAll('#modal-cierre-venta .pill').forEach(p=>{
    p.classList.remove('active');
    if(p.getAttribute('onclick')&&p.getAttribute('onclick').includes("'"+fp.forma+"'")) p.classList.add('active');
  });
  renderCvFormaPago();
  // Pre-rellenar campos de la forma de pago después del render
  setTimeout(()=>{
    if(fp.forma==='DEBITO_BANCARIO'){
      if(document.getElementById('cv-banco-deb')) document.getElementById('cv-banco-deb').value=fp.banco||'Produbanco';
      if(document.getElementById('cv-cuenta-deb')) document.getElementById('cv-cuenta-deb').value=fp.cuenta||c.cuenta||'';
      if(document.getElementById('cv-n-cuotas')) document.getElementById('cv-n-cuotas').value=fp.nCuotas||10;
      if(document.getElementById('cv-fecha-primera')) document.getElementById('cv-fecha-primera').value=fp.fechaPrimera||'';
      renderCvDebCalendar();
    } else if(fp.forma==='TARJETA_CREDITO'){
      if(document.getElementById('cv-n-cuotas-tc')) document.getElementById('cv-n-cuotas-tc').value=fp.nCuotas||12;
      if(document.getElementById('cv-fecha-contacto-tc')) document.getElementById('cv-fecha-contacto-tc').value=fp.fechaContacto||'';
      if(document.getElementById('cv-banco-tc')) document.getElementById('cv-banco-tc').value=fp.banco||'Produbanco';
      if(document.getElementById('cv-tc-digits')) document.getElementById('cv-tc-digits').value=fp.digitos||'';
    } else if(fp.forma==='CONTADO'){
      if(document.getElementById('cv-fecha-cobro-total')) document.getElementById('cv-fecha-cobro-total').value=fp.fechaCobro||'';
      if(document.getElementById('cv-ref-transferencia')) document.getElementById('cv-ref-transferencia').value=fp.referencia||'';
    } else if(fp.forma==='MIXTO'){
      if(document.getElementById('cv-monto-inicial')) document.getElementById('cv-monto-inicial').value=fp.montoInicial||'';
      if(document.getElementById('cv-fecha-mixto-inicial')) document.getElementById('cv-fecha-mixto-inicial').value=fp.fechaInicial||'';
      if(document.getElementById('cv-mixto-metodo')) document.getElementById('cv-mixto-metodo').value=fp.metodoResto||'DEBITO';
      if(document.getElementById('cv-mixto-n-cuotas')) document.getElementById('cv-mixto-n-cuotas').value=fp.nCuotasResto||1;
      if(document.getElementById('cv-mixto-fecha-cuota')) document.getElementById('cv-mixto-fecha-cuota').value=fp.fechaCuotaResto||'';
      if(document.getElementById('cv-mixto-banco')) document.getElementById('cv-mixto-banco').value=fp.banco||'Produbanco';
      if(document.getElementById('cv-mixto-cuenta')) document.getElementById('cv-mixto-cuenta').value=fp.cuenta||'';
      calcMixtoResto();
    }
  },150);
  // Cambiar botón guardar a "Actualizar"
  const btnGuardar=document.querySelector('#modal-cierre-venta .btn-green[onclick="guardarCierreVenta()"]');
  if(btnGuardar){ btnGuardar.textContent='💾 Actualizar Cierre'; btnGuardar.style.background='var(--accent2)'; }
  // Actualizar cabecera de modal para indicar modo edición
  document.getElementById('cv-aseg').textContent='✏ EDITANDO CIERRE — '+c.aseguradora;
  document.getElementById('cv-total').textContent=`${fmt(c.primaTotal)} · ${c.clienteNombre}`;
  openModal('modal-cierre-venta');
}

function eliminarCierre(idx){
  if(!confirm('¿Eliminar este cierre?')) return;
  let cierres=_getCierres();
  if(currentUser&&currentUser.rol!=='admin') cierres=cierres.filter(c=>String(c.ejecutivo)===String(currentUser.id));
  const sorted=cierres.sort((a,b)=>new Date(b.fechaRegistro)-new Date(a.fechaRegistro));
  const toDelete=sorted[idx];
  if(!toDelete) return;
  // Borrar en SharePoint si tiene _spId
  if(toDelete._spId && _spReady) spDelete('cierres', toDelete._spId);
  const allCierres=_getCierres();
  const filtered=allCierres.filter(c=>String(c.id)!==String(toDelete.id));
  _saveCierres(filtered);
  renderCierres();
  showToast('Cierre eliminado','error');
}
function exportCierresExcel(){
  const cierres=_getCierres();
  const data=cierres.map(c=>{
    const fp=c.formaPago||{};
    return{
      'Fecha':c.fechaRegistro,'Cliente':c.clienteNombre,'Aseguradora':c.aseguradora,
      'N° Póliza':c.polizaNueva,'N° Factura':c.facturaAseg,
      'Prima Neta':c.primaNeta,'Prima Total':c.primaTotal,
      'Forma Pago':fp.forma,'N° Cuotas':fp.nCuotas||'—',
      'Cuota Monto':fp.cuotaMonto||'—',
      'Fecha 1ª Cuota/Cobro':fp.fechaPrimera||fp.fechaCobro||fp.fechaContacto||fp.fechaInicial||'—',
      'Banco TC':fp.banco||'—','Fecha Contacto TC':fp.fechaContacto||'—',
      'Vigencia Desde':c.vigDesde,'Vigencia Hasta':c.vigHasta,
      'Observación':c.observacion||''
    };
  });
  const ws=XLSX.utils.json_to_sheet(data);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Cierres');
  XLSX.writeFile(wb,'Reliance_Cierres_'+new Date().toISOString().split('T')[0]+'.xlsx');
  showToast(data.length+' cierres exportados');
}
function selFormaPago(val, el){
  document.getElementById('cv-forma-pago').value=val;
  document.querySelectorAll('#modal-cierre-venta .pill').forEach(p=>p.classList.remove('active'));
  el.classList.add('active');
  renderCvFormaPago();
}
function autocalcHasta(){
  const desde=document.getElementById('cv-desde')?.value;
  if(!desde) return;
  const h=new Date(desde); h.setFullYear(h.getFullYear()+1);
  document.getElementById('cv-hasta').value=h.toISOString().split('T')[0];
}


// ══════════════════════════════════════════════════════
//  WHATSAPP
// ══════════════════════════════════════════════════════
let waClienteId = null;

const WA_PLANTILLAS = {
  vencimiento: (c, exec) => {
    const days = daysUntil(c.hasta);
    const daysText = days < 0 ? `venció hace ${Math.abs(days)} días` : `vence en *${days} días* (${c.hasta})`;
    return `Estimado/a *${primerNombre(c.nombre)}*,

Le informamos que su póliza de seguro vehicular:
🚗 *${c.marca} ${c.modelo} ${c.anio}*
📋 Placa: *${c.placa||'—'}*
🏢 Aseguradora: *${c.aseguradora}*

${daysText}.

Para renovar y mantener su vehículo protegido, contáctenos a la brevedad. Tenemos las mejores opciones del mercado para usted.

*RELIANCE — Asesores de Seguros*
_${exec}_`;
  },
  cotizacion: (c, exec) => {
    return `Estimado/a *${primerNombre(c.nombre)}*,

Le enviamos el comparativo de seguros para su vehículo:
🚗 *${c.marca} ${c.modelo} ${c.anio}*
📋 Placa: *${c.placa||'—'}*
💰 Valor asegurado: *${fmt(c.va)}*

Contamos con las mejores aseguradoras del mercado: GENERALI, SWEADEN, ALIANZA, MAPFRE, LATINA, ZURICH.

Por favor indíquenos si desea recibir el comparativo completo o si prefiere que le llamemos para explicarle las coberturas.

*RELIANCE — Asesores de Seguros*
_${exec}_`;
  },
  cierre: (c, exec) => {
    const poliza = c.polizaNueva || c.poliza || '—';
    return `Estimado/a *${primerNombre(c.nombre)}*,

¡Su póliza ha sido emitida exitosamente! 🎉

🔹 Aseguradora: *${c.aseguradora}*
🔹 N° Póliza: *${poliza}*
🔹 Vehículo: *${c.marca} ${c.modelo} ${c.anio}*
🔹 Vigencia: *${c.desde}* al *${c.hasta}*
${c.factura?`🔹 Factura: *${c.factura}*`:''}

Ante cualquier siniestro comuníquese con nosotros de inmediato. Estamos para servirle.

Gracias por confiar en *RELIANCE*.
_${exec}_`;
  },
  documentos: (c, exec) => {
    return `Estimado/a *${primerNombre(c.nombre)}*,

Para proceder con la ${c.tipo==='RENOVACION'?'renovación':'emisión'} de su póliza de seguro vehicular, necesitamos los siguientes documentos:

📄 *Cédula de identidad* vigente (ambos lados)
🚗 *Matrícula del vehículo* actualizada
🪪 *Licencia de conducir* vigente
${c.tipo==='RENOVACION'?'📋 *Póliza anterior* (si la tiene disponible)':''}

Por favor envíelos por este medio o escríbanos para coordinar.

*RELIANCE — Asesores de Seguros*
_${exec}_`;
  },
  libre: () => ''
};

function primerNombre(nombre){
  if(!nombre) return '—';
  const partes = nombre.trim().split(' ');
  // Formato: APELLIDO1 APELLIDO2 NOMBRE1 NOMBRE2 → devuelve NOMBRE1
  return partes.length >= 3 ? partes[2] : partes[partes.length-1];
}

function openWhatsApp(id, tipoInicial='vencimiento'){
  const c = DB.find(x=>String(x.id)===String(id)); if(!c) return;
  waClienteId = id;
  const exec = currentUser ? currentUser.name : 'Ejecutivo RELIANCE';
  // Número: limpiar y asegurar formato internacional
  const num = (c.celular||'').toString().replace(/\D/g,'');
  document.getElementById('wa-numero').value = num;
  document.getElementById('wa-cliente-meta').textContent =
    `${c.nombre} · ${c.aseguradora} · Vence: ${c.hasta||'—'}`;
  // Seleccionar plantilla inicial
  selWaPlantilla(tipoInicial, document.querySelector(`#wa-tipo-pills .pill`));
  // Marcar pill correcta
  document.querySelectorAll('#wa-tipo-pills .pill').forEach(p=>{
    p.classList.remove('active');
    if(p.getAttribute('onclick')&&p.getAttribute('onclick').includes("'"+tipoInicial+"'")) p.classList.add('active');
  });
  // Generar mensaje
  const fn = WA_PLANTILLAS[tipoInicial];
  const msg = fn ? fn(c, exec) : '';
  document.getElementById('wa-mensaje').value = msg;
  document.getElementById('wa-char-count').textContent = msg.length + ' caracteres';
  openModal('modal-whatsapp');
}

function selWaPlantilla(tipo, el){
  document.querySelectorAll('#wa-tipo-pills .pill').forEach(p=>p.classList.remove('active'));
  if(el) el.classList.add('active');
  const c = DB.find(x=>String(x.id)===String(waClienteId));
  if(!c) return;
  const exec = currentUser ? currentUser.name : 'Ejecutivo RELIANCE';
  const fn = WA_PLANTILLAS[tipo];
  const msg = fn ? fn(c, exec) : '';
  const ta = document.getElementById('wa-mensaje');
  ta.value = msg;
  document.getElementById('wa-char-count').textContent = msg.length + ' caracteres';
}

function enviarWhatsApp(){
  const num = (document.getElementById('wa-numero').value||'').replace(/\D/g,'');
  const msg = document.getElementById('wa-mensaje').value.trim();
  if(!num){ showToast('Ingresa el número de WhatsApp','error'); return; }
  if(!msg){ showToast('El mensaje no puede estar vacío','error'); return; }
  // Registrar en historial del cliente
  const c = DB.find(x=>String(x.id)===String(waClienteId));
  if(c){
    c._dirty = true;
    if(!c.historialWa) c.historialWa=[];
    c.historialWa.push({ fecha: new Date().toISOString().split('T')[0], tipo: 'WhatsApp', resumen: msg.substring(0,80)+'…', ejecutivo: currentUser?.id||'' });
    c.ultimoContacto = new Date().toISOString().split('T')[0];
    saveDB();
  }
  const url = `https://wa.me/${num}?text=${encodeURIComponent(msg)}`;
  window.open(url, '_blank');
  closeModal('modal-whatsapp');
  showToast('WhatsApp abierto — recuerda enviar el mensaje ✓');
}

function copiarWa(){
  const msg = document.getElementById('wa-mensaje').value;
  navigator.clipboard.writeText(msg).then(()=>showToast('Mensaje copiado al portapapeles'));
}

// ══════════════════════════════════════════════════════════════
//  EMAIL — Plantillas y funciones (mismo patrón que WhatsApp)
// ══════════════════════════════════════════════════════════════
let emailClienteId = null;

const EMAIL_PLANTILLAS = {
  vencimiento: (c, exec) => ({
    asunto: `Renovación de Póliza — ${c.marca||''} ${c.modelo||''} ${c.anio||''} · ${c.hasta||''}`.trim(),
    cuerpo: `Estimado/a ${primerNombre(c.nombre)},

Le informamos que su póliza de seguro vehicular está próxima a vencer:

  Vehículo:     ${c.marca||'—'} ${c.modelo||''} ${c.anio||''}
  Placa:        ${c.placa||'—'}
  Aseguradora:  ${c.aseguradora||'—'}
  Vencimiento:  ${c.hasta||'—'}

Para renovar y mantener su vehículo protegido, contáctenos a la brevedad. Tenemos las mejores opciones del mercado para usted.

Quedo a su disposición para cualquier consulta.

Saludos cordiales,
${exec}
RELIANCE — Asesores de Seguros`
  }),

  cotizacion: (c, exec) => ({
    asunto: `Cotización de Seguro Vehicular — ${c.marca||''} ${c.modelo||''} ${c.anio||''}`.trim(),
    cuerpo: `Estimado/a ${primerNombre(c.nombre)},

Le enviamos el comparativo de seguros para su vehículo:

  Vehículo:         ${c.marca||'—'} ${c.modelo||''} ${c.anio||''}
  Placa:            ${c.placa||'—'}
  Valor asegurado:  ${fmt(c.va||0)}

Contamos con las mejores aseguradoras del mercado: GENERALI, SWEADEN, ALIANZA, MAPFRE, LATINA, ZURICH.

Por favor indíquenos si desea recibir el comparativo completo o si prefiere que le llamemos para explicarle las coberturas.

Quedo a su disposición para cualquier consulta.

Saludos cordiales,
${exec}
RELIANCE — Asesores de Seguros`
  }),

  cierre: (c, exec) => ({
    asunto: `Póliza Emitida — ${c.aseguradora||''} · ${c.polizaNueva||c.poliza||''}`.trim(),
    cuerpo: `Estimado/a ${primerNombre(c.nombre)},

Con mucho gusto le informamos que su póliza ha sido emitida exitosamente:

  Aseguradora:  ${c.aseguradora||'—'}
  N° Póliza:    ${c.polizaNueva||c.poliza||'—'}
  Vehículo:     ${c.marca||'—'} ${c.modelo||''} ${c.anio||''}
  Vigencia:     ${c.desde||'—'} al ${c.hasta||'—'}
  ${c.factura ? '  Factura:     '+c.factura : ''}

Ante cualquier siniestro comuníquese con nosotros de inmediato. Estamos para servirle.

Gracias por confiar en RELIANCE.

Saludos cordiales,
${exec}
RELIANCE — Asesores de Seguros`
  }),

  documentos: (c, exec) => ({
    asunto: `Documentos requeridos — ${c.tipo==='RENOVACION'?'Renovación':'Emisión'} de Póliza`,
    cuerpo: `Estimado/a ${primerNombre(c.nombre)},

Para proceder con la ${c.tipo==='RENOVACION'?'renovación':'emisión'} de su póliza de seguro vehicular, necesitamos los siguientes documentos:

  - Cédula de identidad vigente (ambos lados)
  - Matrícula del vehículo actualizada
  - Licencia de conducir vigente
  ${c.tipo==='RENOVACION'?'- Póliza anterior (si la tiene disponible)':''}

Por favor envíelos como adjuntos a este correo o escríbanos para coordinar.

Quedo a su disposición.

Saludos cordiales,
${exec}
RELIANCE — Asesores de Seguros`
  }),

  libre: () => ({ asunto: '', cuerpo: '' })
};

function openEmail(id, tipoInicial='vencimiento'){
  const c = DB.find(x=>String(x.id)===String(id)); if(!c) return;
  emailClienteId = id;
  const exec = currentUser ? currentUser.name : 'Ejecutivo RELIANCE';
  document.getElementById('email-destino').value = c.correo||'';
  document.getElementById('email-cliente-meta').textContent =
    `${c.nombre} · ${c.aseguradora||'—'} · Vence: ${c.hasta||'—'}`;
  // Seleccionar plantilla inicial
  document.querySelectorAll('#email-tipo-pills .pill').forEach(p=>{
    p.classList.remove('active');
    if(p.getAttribute('onclick')&&p.getAttribute('onclick').includes("'"+tipoInicial+"'")) p.classList.add('active');
  });
  const pl = EMAIL_PLANTILLAS[tipoInicial];
  const data = pl ? pl(c, exec) : {asunto:'', cuerpo:''};
  document.getElementById('email-asunto').value = data.asunto;
  document.getElementById('email-cuerpo').value = data.cuerpo;
  document.getElementById('email-char-count').textContent = data.cuerpo.length + ' caracteres';
  openModal('modal-email');
}

function selEmailPlantilla(tipo, el){
  document.querySelectorAll('#email-tipo-pills .pill').forEach(p=>p.classList.remove('active'));
  if(el) el.classList.add('active');
  const c = DB.find(x=>String(x.id)===String(emailClienteId)); if(!c) return;
  const exec = currentUser ? currentUser.name : 'Ejecutivo RELIANCE';
  const pl = EMAIL_PLANTILLAS[tipo];
  const data = pl ? pl(c, exec) : {asunto:'', cuerpo:''};
  document.getElementById('email-asunto').value = data.asunto;
  const ta = document.getElementById('email-cuerpo');
  ta.value = data.cuerpo;
  document.getElementById('email-char-count').textContent = data.cuerpo.length + ' caracteres';
}

function enviarEmail(){
  const destino = (document.getElementById('email-destino').value||'').trim();
  const asunto  = (document.getElementById('email-asunto').value||'').trim();
  const cuerpo  = document.getElementById('email-cuerpo').value.trim();
  if(!destino){ showToast('Ingresa el correo del cliente','error'); return; }
  if(!cuerpo){  showToast('El mensaje no puede estar vacío','error'); return; }
  // Registrar en bitácora
  const c = DB.find(x=>String(x.id)===String(emailClienteId));
  if(c){
    c._dirty = true;
    _bitacoraAdd(c, `Email enviado: ${asunto||'(sin asunto)'}`, 'manual');
    c.ultimoContacto = new Date().toISOString().split('T')[0];
    saveDB();
  }
  // Abrir cliente de correo con mailto:
  const mailto = `mailto:${encodeURIComponent(destino)}?subject=${encodeURIComponent(asunto)}&body=${encodeURIComponent(cuerpo)}`;
  window.location.href = mailto;
  closeModal('modal-email');
  showToast('Outlook abierto — revisa y envía el correo ✓');
}

function copiarEmail(){
  const asunto = document.getElementById('email-asunto').value||'';
  const cuerpo = document.getElementById('email-cuerpo').value||'';
  const texto  = asunto ? `Asunto: ${asunto}\n\n${cuerpo}` : cuerpo;
  navigator.clipboard.writeText(texto).then(()=>{
    showToast('Correo copiado al portapapeles 📋');
  }).catch(()=>{
    const ta=document.getElementById('email-cuerpo');
    ta.select(); document.execCommand('copy');
    showToast('Correo copiado al portapapeles 📋');
  });
}

// ══════════════════════════════════════════════════════════════
//  TAREAS / AGENDA
// ══════════════════════════════════════════════════════════════
let _tareaEditId = null;
let _tareaDetalleId = null;

function _getTareas(){
  try{ return JSON.parse(localStorage.getItem('reliance_tareas')||'[]'); }
  catch(e){ return []; }
}
function _saveTareas(arr){
  localStorage.setItem('reliance_tareas', JSON.stringify(arr));
}

function myTareas(){
  const all = _getTareas();
  if(!currentUser) return [];
  if(currentUser.rol==='admin') return all;
  return all.filter(t => String(t.ejecutivo) === String(currentUser.id));
}

function nuevaTarea(clienteId=null){
  _tareaEditId = null;
  const c = clienteId ? DB.find(x=>String(x.id)===String(clienteId)) : null;
  document.getElementById('tarea-titulo').value = c ? `Llamar a ${primerNombre(c.nombre)}` : '';
  document.getElementById('tarea-tipo').value = 'llamada';
  document.getElementById('tarea-prioridad').value = 'media';
  document.getElementById('tarea-hora').value = '';
  document.getElementById('tarea-desc').value = '';
  // Fecha default = hoy
  document.getElementById('tarea-fecha').value = new Date().toISOString().split('T')[0];
  // Meta del cliente
  const meta = document.getElementById('tarea-cliente-meta');
  if(meta) meta.textContent = c ? `Cliente: ${c.nombre} · ${c.aseguradora||'—'}` : '';
  // Guardar clienteId para cuando se guarde
  document.getElementById('modal-tarea').dataset.clienteId = clienteId||'';
  openModal('modal-tarea');
}
function nuevaTareaDesdeCliente(clienteId){
  // Si se pasa clienteId directo (desde botón), usarlo; si no, usar currentSegIdx
  const id = clienteId || currentSegIdx;
  const c = DB.find(x=>String(x.id)===String(id));
  if(!c) return;
  nuevaTarea(c.id);
}

function guardarTarea(){
  const titulo = document.getElementById('tarea-titulo').value.trim();
  const fecha  = document.getElementById('tarea-fecha').value;
  if(!titulo){ showToast('El título es obligatorio','error'); return; }
  if(!fecha){  showToast('La fecha es obligatoria','error'); return; }

  const clienteId = document.getElementById('modal-tarea').dataset.clienteId || null;
  const c = clienteId ? DB.find(x=>String(x.id)===String(clienteId)) : null;

  const all = _getTareas();
  const tarea = {
    id:            _tareaEditId || ('T' + Date.now()),
    titulo,
    descripcion:   document.getElementById('tarea-desc').value.trim(),
    clienteId:     clienteId || null,
    clienteNombre: c ? c.nombre : '',
    fechaVence:    fecha,
    horaVence:     document.getElementById('tarea-hora').value || '',
    tipo:          document.getElementById('tarea-tipo').value,
    prioridad:     document.getElementById('tarea-prioridad').value,
    estado:        'pendiente',
    ejecutivo:     currentUser?.id || '',
    fechaCreacion: new Date().toISOString().split('T')[0],
    _dirty:        true,
  };

  if(_tareaEditId){
    const idx = all.findIndex(t=>t.id===_tareaEditId);
    if(idx>=0) all[idx] = {...all[idx], ...tarea};
  } else {
    all.push(tarea);
    // Registrar en bitácora del cliente si está vinculada
    if(c) _bitacoraAdd(c, `Tarea creada: ${titulo} para ${fecha}`, 'sistema');
  }

  _saveTareas(all);
  spCreate('tareas', tarea).catch(()=>{});
  closeModal('modal-tarea');
  actualizarBadgeTareas();
  renderDashTareas();
  renderCalendario();
  renderTareasCalendario();
  showToast(`📌 Tarea guardada — ${fecha}`,'success');
}

function completarTarea(id){
  const all = _getTareas();
  const t = all.find(x=>x.id===id); if(!t) return;
  t.estado = 'completada';
  t._dirty = true;
  _saveTareas(all);
  if(t._spId) spUpdate('tareas', t._spId, t).catch(()=>{});
  // Registrar en bitácora del cliente
  const c = t.clienteId ? DB.find(x=>String(x.id)===String(t.clienteId)) : null;
  if(c) _bitacoraAdd(c, `Tarea completada: ${t.titulo}`, 'sistema');
  actualizarBadgeTareas();
  renderDashTareas();
  renderCalendario();
  renderTareasCalendario();
  closeModal('modal-tarea-detalle');
  showToast('✅ Tarea completada');
}

function eliminarTarea(id){
  if(!confirm('¿Eliminar esta tarea?')) return;
  const all = _getTareas().filter(t=>t.id!==id);
  _saveTareas(all);
  actualizarBadgeTareas();
  renderDashTareas();
  renderCalendario();
  renderTareasCalendario();
  closeModal('modal-tarea-detalle');
  showToast('Tarea eliminada');
}

function completarTareaDetalle(){ if(_tareaDetalleId) completarTarea(_tareaDetalleId); }
function eliminarTareaDetalle(){  if(_tareaDetalleId) eliminarTarea(_tareaDetalleId);  }

function abrirDetalleTarea(id){
  const t = _getTareas().find(x=>x.id===id); if(!t) return;
  _tareaDetalleId = id;
  const TIPO_LABEL = { llamada:'📞 Llamada', email:'✉️ Email', reunion:'🤝 Reunión', seguimiento:'📋 Seguimiento', otro:'📌 Otro' };
  const PRIO_LABEL = { alta:'🔴 Alta', media:'🟡 Media', baja:'🟢 Baja' };
  document.getElementById('tarea-det-titulo').textContent = t.titulo;
  document.getElementById('tarea-det-body').innerHTML = `
    <div style="display:flex;flex-direction:column;gap:10px">
      <div style="display:flex;gap:8px;flex-wrap:wrap">
        <span class="badge badge-blue" style="font-size:11px">${TIPO_LABEL[t.tipo]||t.tipo}</span>
        <span class="badge" style="font-size:11px">${PRIO_LABEL[t.prioridad]||t.prioridad}</span>
        <span class="badge ${t.estado==='completada'?'badge-green':'badge-gray'}" style="font-size:11px">
          ${t.estado==='completada'?'✅ Completada':'⏳ Pendiente'}
        </span>
      </div>
      <div style="font-size:13px">📅 <b>${t.fechaVence}</b>${t.horaVence?' a las <b>'+t.horaVence+'</b>':''}</div>
      ${t.clienteNombre?`<div style="font-size:12px;color:var(--muted)">👤 Cliente: ${t.clienteNombre}</div>`:''}
      ${t.descripcion?`<div style="font-size:13px;padding:10px;background:var(--warm);border-radius:6px;margin-top:4px">${t.descripcion}</div>`:''}
    </div>`;
  // Ocultar botón completar si ya está completada
  const btnC = document.getElementById('tarea-det-btn-completar');
  if(btnC) btnC.style.display = t.estado==='completada' ? 'none' : '';
  openModal('modal-tarea-detalle');
}

// ── Renders ─────────────────────────────────────────────────

function actualizarBadgeTareas(){
  const hoy = new Date().toISOString().split('T')[0];
  const pendHoy = myTareas().filter(t=>t.estado==='pendiente' && t.fechaVence<=hoy).length;
  const el = document.getElementById('badge-tareas');
  if(!el) return;
  if(pendHoy>0){
    el.textContent = pendHoy;
    el.style.display = '';
  } else {
    el.style.display = 'none';
  }
}

function renderDashTareas(){
  const el = document.getElementById('dash-tareas'); if(!el) return;
  const hoy = new Date().toISOString().split('T')[0];
  const tareas = myTareas()
    .filter(t=>t.estado==='pendiente' && t.fechaVence===hoy)
    .sort((a,b)=>(a.horaVence||'99:99').localeCompare(b.horaVence||'99:99'));
  const manana = new Date(); manana.setDate(manana.getDate()+1);
  const mananaStr = manana.toISOString().split('T')[0];
  const proximas = myTareas()
    .filter(t=>t.estado==='pendiente' && t.fechaVence>hoy && t.fechaVence<=mananaStr);

  if(!tareas.length && !proximas.length){
    el.innerHTML = '<div style="color:var(--muted);font-size:12px;padding:8px">Sin tareas para hoy 🎉</div>';
    return;
  }
  const TIPO_ICON = { llamada:'📞', email:'✉️', reunion:'🤝', seguimiento:'📋', otro:'📌' };
  const PRIO_COLOR = { alta:'var(--red)', media:'var(--gold)', baja:'var(--green)' };
  el.innerHTML = tareas.map(t=>`
    <div style="display:flex;align-items:center;gap:8px;padding:8px 0;border-bottom:1px solid var(--border);cursor:pointer"
      onclick="abrirDetalleTarea('${t.id}')">
      <span style="font-size:16px">${TIPO_ICON[t.tipo]||'📌'}</span>
      <div style="flex:1;min-width:0">
        <div style="font-size:12px;font-weight:600;white-space:nowrap;overflow:hidden;text-overflow:ellipsis">${t.titulo}</div>
        ${t.horaVence?`<div style="font-size:10px;color:var(--muted)">${t.horaVence}${t.clienteNombre?' · '+primerNombre(t.clienteNombre):''}</div>`:''}
      </div>
      <div style="width:6px;height:6px;border-radius:50%;background:${PRIO_COLOR[t.prioridad]||'var(--muted)'};flex-shrink:0"></div>
      <button class="btn btn-ghost btn-xs" onclick="event.stopPropagation();completarTarea('${t.id}')">✅</button>
    </div>`).join('') +
  (proximas.length?`<div style="font-size:10px;color:var(--muted);padding:6px 0">Mañana: ${proximas.length} tarea${proximas.length>1?'s':''}</div>`:'');
}

function renderTareasCalendario(){
  const el = document.getElementById('cal-tareas-lista'); if(!el) return;
  const hoy = new Date().toISOString().split('T')[0];
  const pendientes = myTareas()
    .filter(t=>t.estado==='pendiente')
    .sort((a,b)=>a.fechaVence.localeCompare(b.fechaVence));
  if(!pendientes.length){
    el.innerHTML='<div style="color:var(--muted);font-size:12px;padding:8px">Sin tareas pendientes</div>';
    return;
  }
  const TIPO_ICON = { llamada:'📞', email:'✉️', reunion:'🤝', seguimiento:'📋', otro:'📌' };
  const PRIO_COLOR = { alta:'var(--red)', media:'var(--gold)', baja:'var(--green)' };
  el.innerHTML = pendientes.map(t=>{
    const dias = Math.round((new Date(t.fechaVence) - new Date(hoy)) / 86400000);
    const vencida = dias < 0;
    return `
    <div style="display:flex;align-items:center;gap:8px;padding:8px;border-radius:6px;
      background:${vencida?'#fde8e0':'var(--warm)'};margin-bottom:6px;cursor:pointer"
      onclick="abrirDetalleTarea('${t.id}')">
      <span style="font-size:16px">${TIPO_ICON[t.tipo]||'📌'}</span>
      <div style="flex:1;min-width:0">
        <div style="font-size:12px;font-weight:600">${t.titulo}</div>
        <div style="font-size:10px;color:var(--muted)">
          ${vencida?'<span style="color:var(--red)">Vencida</span>':dias===0?'<b>Hoy</b>':dias===1?'Mañana':'En '+dias+' días'}
          ${t.horaVence?' · '+t.horaVence:''} ${t.clienteNombre?' · '+primerNombre(t.clienteNombre):''}
        </div>
      </div>
      <div style="width:6px;height:6px;border-radius:50%;background:${PRIO_COLOR[t.prioridad]||'var(--muted)'};flex-shrink:0"></div>
      <button class="btn btn-ghost btn-xs" onclick="event.stopPropagation();completarTarea('${t.id}')" title="Completar">✅</button>
    </div>`; }).join('');
}

// ══════════════════════════════════════════════════════════════
//  NOTIFICACIONES DEL NAVEGADOR
// ══════════════════════════════════════════════════════════════
let _notifPermiso = false;
let _notifDisparadas = false;
let _notifInterval = null;

async function initNotificaciones(){
  if(!('Notification' in window)) return;
  if(Notification.permission === 'granted'){
    _notifPermiso = true;
  } else if(Notification.permission !== 'denied'){
    const perm = await Notification.requestPermission();
    _notifPermiso = perm === 'granted';
  }
  if(_notifPermiso && !_notifDisparadas){
    _notifDisparadas = true;
    setTimeout(dispararNotificaciones, 2000);
    // Repetir cada hora — guardar referencia para poder limpiarlo
    if(!_notifInterval) _notifInterval = setInterval(dispararNotificaciones, 3600000);
  }
}

function dispararNotificaciones(){
  if(!_notifPermiso || !currentUser) return;
  const hoy = new Date().toISOString().split('T')[0];
  const mine = myClientes();
  const excluir = ['EMITIDO','PÓLIZA VIGENTE'];

  // ── Deduplicación diaria por categoría ──────────────────────────────────────
  // Notificaciones de vencimiento lejano solo se disparan UNA VEZ por día
  const _logKey = `notif_diario_${currentUser.id}_${hoy}`;
  let _fired = {};
  try{ _fired = JSON.parse(localStorage.getItem(_logKey)||'{}'); }catch(e){}
  const _markFired = (cat) => {
    _fired[cat] = true;
    localStorage.setItem(_logKey, JSON.stringify(_fired));
  };
  // Limpiar logs antiguos (>2 días)
  const ayer = new Date(); ayer.setDate(ayer.getDate()-2);
  const ayerStr = ayer.toISOString().split('T')[0];
  Object.keys(localStorage).filter(k=>k.startsWith('notif_diario_')&&k<`notif_diario_${currentUser.id}_${ayerStr}`).forEach(k=>localStorage.removeItem(k));

  // ── Vence HOY — siempre notificar ──────────────────────────────────────────
  const hoyVenc = mine.filter(c=>c.hasta===hoy && !excluir.includes(c.estado));
  if(hoyVenc.length){
    _notif(`⚠️ ${hoyVenc.length} póliza${hoyVenc.length>1?'s':''} vence${hoyVenc.length>1?'n':''} HOY`,
      hoyVenc.slice(0,3).map(c=>primerNombre(c.nombre)).join(', ')+(hoyVenc.length>3?` y ${hoyVenc.length-3} más`:''));
  }

  // ── Vence en 7 días — siempre notificar ───────────────────────────────────
  const venc7 = mine.filter(c=>{ const d=daysUntil(c.hasta); return d>0&&d<=7&&!excluir.includes(c.estado); });
  if(venc7.length){
    _notif(`📅 ${venc7.length} cliente${venc7.length>1?'s':''} vence${venc7.length>1?'n':''} en 7 días`,
      venc7.slice(0,3).map(c=>`${primerNombre(c.nombre)} (${daysUntil(c.hasta)}d)`).join(', '));
  }

  // ── Vence en 8-15 días — una vez por día ──────────────────────────────────
  if(!_fired['15d']){
    const venc15 = mine.filter(c=>{ const d=daysUntil(c.hasta); return d>7&&d<=15&&!excluir.includes(c.estado); });
    if(venc15.length){
      _notif(`🔴 ${venc15.length} póliza${venc15.length>1?'s':''} vence${venc15.length>1?'n':''} en 15 días — urgente`,
        venc15.slice(0,3).map(c=>`${primerNombre(c.nombre)} (${daysUntil(c.hasta)}d)`).join(', '));
      _markFired('15d');
    }
  }

  // ── Vence en 16-30 días — una vez por día ─────────────────────────────────
  if(!_fired['30d']){
    const venc30 = mine.filter(c=>{ const d=daysUntil(c.hasta); return d>15&&d<=30&&!excluir.includes(c.estado); });
    if(venc30.length){
      _notif(`⚡ ${venc30.length} póliza${venc30.length>1?'s':''} vence${venc30.length>1?'n':''} en 30 días`,
        venc30.slice(0,3).map(c=>`${primerNombre(c.nombre)} (${daysUntil(c.hasta)}d)`).join(', '));
      _markFired('30d');
    }
  }

  // ── Vence en 31-60 días — una vez por día ─────────────────────────────────
  if(!_fired['60d']){
    const venc60 = mine.filter(c=>{ const d=daysUntil(c.hasta); return d>30&&d<=60&&!excluir.includes(c.estado); });
    if(venc60.length){
      _notif(`⏳ ${venc60.length} póliza${venc60.length>1?'s':''} vence${venc60.length>1?'n':''} en 60 días — agenda renovaciones`,
        venc60.slice(0,3).map(c=>`${primerNombre(c.nombre)} (${daysUntil(c.hasta)}d)`).join(', '));
      _markFired('60d');
    }
  }

  // ── Vence en 61-90 días — una vez por día ─────────────────────────────────
  if(!_fired['90d']){
    const venc90 = mine.filter(c=>{ const d=daysUntil(c.hasta); return d>60&&d<=90&&!excluir.includes(c.estado); });
    if(venc90.length){
      _notif(`🗓️ ${venc90.length} póliza${venc90.length>1?'s':''} vence${venc90.length>1?'n':''} en 90 días — inicia cotización`,
        venc90.slice(0,3).map(c=>`${primerNombre(c.nombre)} (${daysUntil(c.hasta)}d)`).join(', '));
      _markFired('90d');
    }
  }

  // ── Tareas pendientes de hoy o vencidas — siempre notificar ───────────────
  const tareasHoy = myTareas().filter(t=>t.estado==='pendiente' && t.fechaVence<=hoy);
  if(tareasHoy.length){
    _notif(`📌 ${tareasHoy.length} tarea${tareasHoy.length>1?'s':''} pendiente${tareasHoy.length>1?'s':''}`,
      tareasHoy.slice(0,3).map(t=>t.titulo).join(', '));
  }

  // ── Tareas con hora — notif cuando la hora actual coincide (±5 min) ────────
  const ahoraH = new Date().getHours().toString().padStart(2,'0');
  const ahoraM = new Date().getMinutes();
  myTareas().filter(t=>t.estado==='pendiente' && t.fechaVence===hoy && t.horaVence).forEach(t=>{
    const [h,m] = (t.horaVence||'').split(':').map(Number);
    if(!isNaN(h) && !isNaN(m)){
      const diffMin = Math.abs((h*60+m) - (parseInt(ahoraH)*60+ahoraM));
      if(diffMin <= 5){
        _notif(`⏰ ${t.titulo}`, `${t.horaVence}${t.clienteNombre?' · '+primerNombre(t.clienteNombre):''}`);
      }
    }
  });
}

function _notif(titulo, cuerpo){
  if(!_notifPermiso) return;
  try{
    const n = new Notification(titulo, {
      body:  cuerpo,
      icon:  'https://sgrandacrm.github.io/crm-reliance/favicon.ico',
      badge: 'https://sgrandacrm.github.io/crm-reliance/favicon.ico',
      tag:   titulo, // evita duplicados
    });
    n.onclick = () => { window.focus(); n.close(); };
    setTimeout(()=>n.close(), 8000);
  } catch(e){}
}



// ══════════════════════════════════════════════════════
//  REPORTES Y MÉTRICAS
// ══════════════════════════════════════════════════════
let repCharts = {};

function getPeriodoRange(){
  const p = document.getElementById('rep-periodo')?.value || 'mes';
  const now = new Date();
  let desde;
  if(p==='mes') desde = new Date(now.getFullYear(), now.getMonth(), 1);
  else if(p==='trimestre') desde = new Date(now.getFullYear(), Math.floor(now.getMonth()/3)*3, 1);
  else if(p==='anio') desde = new Date(now.getFullYear(), 0, 1);
  else desde = new Date('2000-01-01');
  return { desde, hasta: now };
}

function renderReportes(){
  const {desde, hasta} = getPeriodoRange();
  const desdeStr = desde.toISOString().split('T')[0];

  // Todos los cierres del período
  const allCierres = _getCierres();
  const cierresPeriodo = allCierres.filter(c => (c.fechaRegistro||'') >= desdeStr);

  // Clientes activos del período (vencen en el período o gestionados)
  const clientesPeriodo = DB.filter(c => (c.ultimoContacto||'') >= desdeStr || (c.hasta||'') >= desdeStr);

  // ── KPIs ──
  const totalPrimas = cierresPeriodo.reduce((s,c)=>s+(c.primaTotal||0),0);
  const totalCierres = cierresPeriodo.length;
  const tasaRenovacion = DB.length > 0 ? Math.round((DB.filter(c=>c.estado==='RENOVADO').length / DB.length)*100) : 0;
  const venc30 = DB.filter(c=>{ const d=daysUntil(c.hasta); return d>=0&&d<=30; }).length;

  const kpisEl = document.getElementById('rep-kpis');
  if(kpisEl) kpisEl.innerHTML = `
    <div class="stat-card">
      <div class="stat-label">Primas del período</div>
      <div class="stat-value" style="color:var(--green);font-size:20px">${fmt(totalPrimas)}</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Cierres registrados</div>
      <div class="stat-value" style="color:var(--accent2)">${totalCierres}</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Tasa de renovación</div>
      <div class="stat-value" style="color:${tasaRenovacion>=60?'var(--green)':'var(--accent)'}">${tasaRenovacion}%</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Vencen en 30 días</div>
      <div class="stat-value" style="color:var(--accent)">${venc30}</div>
    </div>
    <div class="stat-card">
      <div class="stat-label">Comisión estimada período</div>
      <div class="stat-value" style="color:var(--gold);font-size:20px" id="rep-kpi-comision">${fmt(cierresPeriodo.reduce((s,c)=>s+((c.comision)||(Math.round((c.primaNeta||0)*((_getComisiones()[c.aseguradora]||0)/100)*100)/100)),0))}</div>
    </div>`;

  // ── CHART 1: Estados ──
  // Contar todos los 17 estados dinámicamente
  const estadoCount = {};
  Object.keys(ESTADOS_RELIANCE).forEach(k=>estadoCount[k]=0);
  DB.forEach(c => { const e=c.estado||'PENDIENTE'; if(estadoCount[e]!==undefined) estadoCount[e]++; });
  // Para el donut mostrar solo los que tienen datos (top 6)
  const estadoTop = Object.entries(estadoCount)
    .filter(([,v])=>v>0)
    .sort((a,b)=>b[1]-a[1])
    .slice(0,6);
  renderDonut('rep-chart-estados',
    estadoTop.map(([k])=>ESTADOS_RELIANCE[k].icon+' '+ESTADOS_RELIANCE[k].label),
    estadoTop.map(([,v])=>v),
    estadoTop.map(([k])=>ESTADOS_RELIANCE[k].color));

  // ── CHART 2: Cierres por aseguradora ──
  const asegCount = {};
  cierresPeriodo.forEach(c => {
    const a = (c.aseguradora||'OTRO').split(' ')[0];
    asegCount[a] = (asegCount[a]||0) + 1;
  });
  const asegSorted = Object.entries(asegCount).sort((a,b)=>b[1]-a[1]);
  renderBarChart('rep-chart-aseg',
    asegSorted.map(x=>x[0]), asegSorted.map(x=>x[1]),
    '#1a4c84', 'Cierres');

  // ── CHART 3: Primas por forma de pago ──
  const fpPrimas = {Débito:0, TC:0, Contado:0, Mixto:0};
  cierresPeriodo.forEach(c => {
    const fp = c.formaPago?.forma||'';
    if(fp==='DEBITO_BANCARIO') fpPrimas['Débito']+=(c.primaTotal||0);
    else if(fp==='TARJETA_CREDITO') fpPrimas['TC']+=(c.primaTotal||0);
    else if(fp==='CONTADO') fpPrimas['Contado']+=(c.primaTotal||0);
    else if(fp==='MIXTO') fpPrimas['Mixto']+=(c.primaTotal||0);
  });
  renderDonut('rep-chart-fp',
    Object.keys(fpPrimas), Object.values(fpPrimas).map(v=>Math.round(v)),
    ['#1a4c84','#2d6a4f','#f0c040','#c84b1a']);

  // ── CHART 4: Vencimientos próximos 90 días por semana ──
  const semanas = ['1-15 días','16-30 días','31-45 días','46-60 días','61-90 días'];
  const vencSem = [0,0,0,0,0];
  DB.forEach(c => {
    const d = daysUntil(c.hasta);
    if(d>=1&&d<=15) vencSem[0]++;
    else if(d>=16&&d<=30) vencSem[1]++;
    else if(d>=31&&d<=45) vencSem[2]++;
    else if(d>=46&&d<=60) vencSem[3]++;
    else if(d>=61&&d<=90) vencSem[4]++;
  });
  renderBarChart('rep-chart-venc', semanas, vencSem, '#c84b1a', 'Clientes');

  // ── RANKING EJECUTIVOS ──
  const execs = USERS.filter(u=>u.rol==='ejecutivo');
  const rankingEl = document.getElementById('rep-ranking');
  if(rankingEl){
    const execStats = execs.map(u => {
      const mine = DB.filter(c=>c.ejecutivo===u.id);
      const myCierres = cierresPeriodo.filter(c=>c.ejecutivo===u.id);
      const prima = myCierres.reduce((s,c)=>s+(c.primaTotal||0),0);
      const comisionTotal = myCierres.reduce((s,c)=>{
        const com = c.comision || Math.round((c.primaNeta||0)*((_getComisiones()[c.aseguradora]||0)/100)*100)/100;
        return s+com;
      },0);
      const renovados = mine.filter(c=>c.estado==='RENOVADO').length;
      const tasa = mine.length > 0 ? Math.round(renovados/mine.length*100) : 0;
      return {u, mine:mine.length, cierres:myCierres.length, prima, comisionTotal, tasa};
    }).sort((a,b)=>b.prima-a.prima);

    rankingEl.innerHTML = `<table style="width:100%;font-size:13px">
      <thead><tr>
        <th style="padding:8px;text-align:left;color:var(--muted)">#</th>
        <th style="padding:8px;text-align:left;color:var(--muted)">Ejecutivo</th>
        <th style="padding:8px;color:var(--muted)">Clientes</th>
        <th style="padding:8px;color:var(--muted)">Cierres período</th>
        <th style="padding:8px;color:var(--muted)">Prima recaudada</th>
        <th style="padding:8px;color:var(--muted)">Comisión estimada</th>
        <th style="padding:8px;color:var(--muted)">Tasa renovación</th>
        <th style="padding:8px;color:var(--muted)">Progreso</th>
      </tr></thead>
      <tbody>${execStats.map((s,i)=>`
        <tr style="border-bottom:1px solid var(--warm)">
          <td style="padding:8px;font-weight:700;font-size:16px;color:${i===0?'#f0c040':i===1?'#9e9e9e':i===2?'#cd7f32':'var(--muted)'}">${i===0?'🥇':i===1?'🥈':i===2?'🥉':i+1}</td>
          <td style="padding:8px">
            <div style="display:flex;align-items:center;gap:8px">
              <div style="width:28px;height:28px;border-radius:50%;background:${s.u.color};color:#fff;display:flex;align-items:center;justify-content:center;font-size:11px;font-weight:700">${s.u.initials}</div>
              <span style="font-weight:500">${s.u.name}</span>
            </div>
          </td>
          <td style="padding:8px;text-align:center">${s.mine}</td>
          <td style="padding:8px;text-align:center;font-weight:600;color:var(--accent2)">${s.cierres}</td>
          <td style="padding:8px;text-align:center;font-weight:700;color:var(--green)">${fmt(s.prima)}</td>
          <td style="padding:8px;text-align:center;font-weight:600;color:var(--gold)">${fmt(s.comisionTotal)}</td>
          <td style="padding:8px;text-align:center">
            <span style="color:${s.tasa>=60?'var(--green)':'var(--accent)'};font-weight:600">${s.tasa}%</span>
          </td>
          <td style="padding:8px;min-width:120px">
            <div style="background:var(--warm);border-radius:99px;height:8px;overflow:hidden">
              <div style="height:100%;width:${Math.min(100,s.tasa)}%;background:${s.tasa>=60?'var(--green)':'var(--accent)'};border-radius:99px;transition:width .5s"></div>
            </div>
          </td>
        </tr>`).join('')}</tbody>
    </table>`;
  }

  // ── TABLA CIERRES ──
  const tbody = document.getElementById('rep-cierres-tbody');
  const countEl = document.getElementById('rep-cierres-count');
  if(countEl) countEl.textContent = cierresPeriodo.length+' cierres';
  const fpLabel={DEBITO_BANCARIO:'🏦 Débito',TARJETA_CREDITO:'💳 TC',CONTADO:'💵 Contado',MIXTO:'🔀 Mixto'};
  if(tbody) tbody.innerHTML = cierresPeriodo
    .sort((a,b)=>(b.fechaRegistro||'').localeCompare(a.fechaRegistro||''))
    .map(c=>{
      const exec = USERS.find(u=>u.id===c.ejecutivo);
      return `<tr>
        <td><span class="mono" style="font-size:11px">${c.fechaRegistro||'—'}</span></td>
        <td style="font-weight:500;font-size:12px">${c.clienteNombre||'—'}</td>
        <td style="font-size:11px">${c.aseguradora||'—'}</td>
        <td class="mono" style="font-weight:700;color:var(--green)">${fmt(c.primaTotal)}</td>
        <td><span class="badge badge-blue" style="font-size:10px">${fpLabel[c.formaPago?.forma]||'—'}</span></td>
        <td style="font-size:11px">${exec?exec.name:'—'}</td>
      </tr>`;
    }).join('') || '<tr><td colspan="6"><div class="empty-state"><div class="empty-icon">📊</div><p>Sin cierres en este período</p></div></td></tr>';

  // ── CHART 5: Producción mensual últimos 6 meses ──
  const meses6 = [], primasMes = [];
  for(let i=5; i>=0; i--){
    const d = new Date();
    d.setDate(1); d.setMonth(d.getMonth()-i);
    const y=d.getFullYear(), m=String(d.getMonth()+1).padStart(2,'0');
    meses6.push(d.toLocaleString('es-ES',{month:'short',year:'2-digit'}));
    primasMes.push(Math.round(
      allCierres.filter(c=>(c.fechaRegistro||'').startsWith(`${y}-${m}`))
               .reduce((s,c)=>s+(c.primaTotal||0),0)
    ));
  }
  renderBarChart('rep-chart-produccion', meses6, primasMes, '#1a4c84', 'Prima $');

  // ── CHART 6: Cobranza del mes actual ──
  const ahora = new Date();
  const mesIni = new Date(ahora.getFullYear(), ahora.getMonth(), 1).toISOString().split('T')[0];
  const mesFin = new Date(ahora.getFullYear(), ahora.getMonth()+1, 0).toISOString().split('T')[0];
  let cobMonto=0, pendMonto=0, fallMonto=0, cobN=0, pendN=0, fallN=0;
  allCierres.forEach(cierre=>{
    _getCuotasFromCierre(cierre).forEach(q=>{
      if((q.fecha||'')<mesIni||(q.fecha||'')>mesFin) return;
      if(q.estado==='COBRADO')      { cobMonto+=q.monto; cobN++; }
      else if(q.estado==='FALLIDO') { fallMonto+=q.monto; fallN++; }
      else                          { pendMonto+=q.monto; pendN++; }
    });
  });
  const cobEl = document.getElementById('rep-cobranza-stats');
  if(cobEl) cobEl.innerHTML = `
    <div class="stat-card" style="flex:1;min-width:110px;padding:10px">
      <div class="stat-label" style="font-size:11px">Cobrado</div>
      <div style="color:var(--green);font-weight:700;font-size:15px">${fmt(cobMonto)}</div>
      <div style="color:var(--muted);font-size:10px">${cobN} cuotas</div>
    </div>
    <div class="stat-card" style="flex:1;min-width:110px;padding:10px">
      <div class="stat-label" style="font-size:11px">Pendiente</div>
      <div style="color:var(--accent);font-weight:700;font-size:15px">${fmt(pendMonto)}</div>
      <div style="color:var(--muted);font-size:10px">${pendN} cuotas</div>
    </div>
    <div class="stat-card" style="flex:1;min-width:110px;padding:10px">
      <div class="stat-label" style="font-size:11px">Fallido</div>
      <div style="color:var(--red);font-weight:700;font-size:15px">${fmt(fallMonto)}</div>
      <div style="color:var(--muted);font-size:10px">${fallN} cuotas</div>
    </div>`;
  renderBarChart('rep-chart-cobranza',
    ['Cobrado','Pendiente','Fallido'],
    [Math.round(cobMonto), Math.round(pendMonto), Math.round(fallMonto)],
    ['#2d6a4f','#e6820a','#c84b1a'], '$');

  // ── PROYECCIÓN DE RENOVACIONES 90 días ──
  const proxVenc = DB.filter(c=>{ const d=daysUntil(c.hasta); return d>=0&&d<=90; })
    .sort((a,b)=>(a.hasta||'').localeCompare(b.hasta||''));
  const proyCount = document.getElementById('rep-proyeccion-count');
  if(proyCount) proyCount.textContent = `${proxVenc.length} clientes`;
  const primaEstTot = proxVenc.reduce((s,c)=>{
    const ult = allCierres.filter(x=>String(x.clienteId)===String(c.id))
      .sort((a,b)=>(b.fechaRegistro||'').localeCompare(a.fechaRegistro||''))[0];
    return s+(ult?.primaTotal||0);
  },0);
  const proyEl = document.getElementById('rep-proyeccion');
  if(proyEl){
    proyEl.innerHTML = `
      <div style="padding:12px 16px;background:var(--bg-alt);border-bottom:1px solid var(--warm);display:flex;gap:28px;flex-wrap:wrap">
        <div><span style="color:var(--muted);font-size:12px">Clientes por vencer: </span><b>${proxVenc.length}</b></div>
        <div><span style="color:var(--muted);font-size:12px">Prima estimada: </span><b style="color:var(--green)">${fmt(primaEstTot)}</b></div>
        <div><span style="color:var(--muted);font-size:12px">En 30 días: </span><b style="color:var(--accent)">${proxVenc.filter(c=>daysUntil(c.hasta)<=30).length}</b></div>
      </div>
      <div class="tbl-wrap"><table>
        <thead><tr>
          <th>Cliente</th><th>Vence</th><th>Días</th><th>Aseguradora</th>
          <th>Prima anterior</th><th>Estado</th><th>Ejecutivo</th>
        </tr></thead>
        <tbody>${proxVenc.slice(0,20).map(c=>{
          const dias = daysUntil(c.hasta);
          const ult = allCierres.filter(x=>String(x.clienteId)===String(c.id))
            .sort((a,b)=>(b.fechaRegistro||'').localeCompare(a.fechaRegistro||''))[0];
          const exec = USERS.find(u=>u.id===c.ejecutivo);
          const est = ESTADOS_RELIANCE[c.estado]||{icon:'•',label:c.estado||'—',color:'#999'};
          return `<tr>
            <td style="font-weight:500;font-size:12px">${c.nombre||'—'}</td>
            <td class="mono" style="font-size:11px">${c.hasta||'—'}</td>
            <td style="text-align:center">
              <span style="font-weight:700;color:${dias<=15?'var(--red)':dias<=30?'var(--accent)':'var(--text)'}">${dias}</span>
            </td>
            <td style="font-size:11px">${c.aseguradora||'—'}</td>
            <td class="mono" style="font-weight:700;color:var(--green);font-size:12px">${ult?fmt(ult.primaTotal):'—'}</td>
            <td><span class="badge" style="background:${est.color}20;color:${est.color};font-size:10px">${est.icon} ${est.label}</span></td>
            <td style="font-size:11px">${exec?exec.name:'—'}</td>
          </tr>`;
        }).join('')}${proxVenc.length>20?`<tr><td colspan="7" style="text-align:center;padding:8px;color:var(--muted);font-size:12px">... y ${proxVenc.length-20} más</td></tr>`:''}</tbody>
      </table></div>`;
  }

  // ── CHART 7: Cartera por tipo de cliente ──
  const tiposCount = {};
  DB.forEach(c=>{ const t=c.tipoCliente||'PARTICULAR'; tiposCount[t]=(tiposCount[t]||0)+1; });
  const tiposColors = {'PARTICULAR':'#1a4c84','PRODUBANCO':'#2d6a4f','FLOTA':'#c84b1a','OTRO':'#9c59b6'};
  renderDonut('rep-chart-tipos',
    Object.keys(tiposCount), Object.values(tiposCount),
    Object.keys(tiposCount).map(k=>tiposColors[k]||'#9e9e9e'));

  // ── CHART 8: Cartera por región ──
  const regCount = {};
  DB.forEach(c=>{ const r=(c.region||'Sin región').trim()||'Sin región'; regCount[r]=(regCount[r]||0)+1; });
  const regSorted = Object.entries(regCount).sort((a,b)=>b[1]-a[1]).slice(0,8);
  renderBarChart('rep-chart-regiones',
    regSorted.map(x=>x[0]), regSorted.map(x=>x[1]), '#6c3483', 'Clientes');
}

function renderDonut(canvasId, labels, data, colors){
  const canvas = document.getElementById(canvasId);
  if(!canvas) return;
  if(repCharts[canvasId]){ repCharts[canvasId].destroy(); delete repCharts[canvasId]; }
  // Dibujar donut manual en canvas 2D
  const ctx = canvas.getContext('2d');
  const W = canvas.offsetWidth||300, H = canvas.height||220;
  canvas.width = W; canvas.height = H;
  const total = data.reduce((s,v)=>s+v,0);
  if(total===0){ ctx.fillStyle='#eee'; ctx.font='14px sans-serif'; ctx.textAlign='center'; ctx.fillText('Sin datos',W/2,H/2); return; }
  const cx=W*0.38, cy=H/2, r=Math.min(cx,cy)-20, ri=r*0.55;
  let angle=-Math.PI/2;
  data.forEach((v,i)=>{
    if(v===0) return;
    const sweep = (v/total)*Math.PI*2;
    ctx.beginPath(); ctx.moveTo(cx,cy);
    ctx.arc(cx,cy,r,angle,angle+sweep);
    ctx.closePath(); ctx.fillStyle=colors[i%colors.length]; ctx.fill();
    // Borde blanco
    ctx.beginPath(); ctx.moveTo(cx,cy);
    ctx.arc(cx,cy,r,angle,angle+sweep);
    ctx.closePath(); ctx.strokeStyle='#fff'; ctx.lineWidth=2; ctx.stroke();
    angle+=sweep;
  });
  // Hueco central
  ctx.beginPath(); ctx.arc(cx,cy,ri,0,Math.PI*2);
  ctx.fillStyle=getComputedStyle(document.documentElement).getPropertyValue('--paper')||'#fff';
  ctx.fill();
  // Total en centro
  ctx.fillStyle='#333'; ctx.font='bold 14px sans-serif'; ctx.textAlign='center'; ctx.textBaseline='middle';
  ctx.fillText(total, cx, cy-8);
  ctx.font='11px sans-serif'; ctx.fillStyle='#888';
  ctx.fillText('total', cx, cy+10);
  // Leyenda derecha
  const lx=W*0.72, ly=H/2-(labels.length*22)/2;
  labels.forEach((l,i)=>{
    if(data[i]===0) return;
    const y=ly+i*24;
    ctx.fillStyle=colors[i%colors.length]; ctx.fillRect(lx,y,12,12);
    ctx.fillStyle='#555'; ctx.font='11px sans-serif'; ctx.textAlign='left';
    ctx.fillText(`${l}: ${data[i]}`,lx+18,y+10);
  });
}

function renderBarChart(canvasId, labels, data, color, label){
  const canvas = document.getElementById(canvasId);
  if(!canvas) return;
  const ctx = canvas.getContext('2d');
  const W = canvas.offsetWidth||300, H = canvas.height||220;
  canvas.width = W; canvas.height = H;
  const maxVal = Math.max(...data, 1);
  const padL=45, padR=20, padT=20, padB=45;
  const chartW=W-padL-padR, chartH=H-padT-padB;
  const barW = chartW/labels.length*0.6;
  const gap = chartW/labels.length;
  // Grid lines
  ctx.strokeStyle='#eee'; ctx.lineWidth=1;
  for(let i=0;i<=4;i++){
    const y=padT+chartH*(1-i/4);
    ctx.beginPath(); ctx.moveTo(padL,y); ctx.lineTo(W-padR,y); ctx.stroke();
    ctx.fillStyle='#aaa'; ctx.font='10px sans-serif'; ctx.textAlign='right';
    ctx.fillText(Math.round(maxVal*i/4), padL-6, y+4);
  }
  // Bars
  data.forEach((v,i)=>{
    const bh = (v/maxVal)*chartH;
    const bx = padL+i*gap+(gap-barW)/2;
    const by = padT+chartH-bh;
    // Gradient effect (color puede ser string o array)
    const c = Array.isArray(color) ? (color[i]||color[0]) : color;
    const grad=ctx.createLinearGradient(0,by,0,by+bh);
    grad.addColorStop(0,c); grad.addColorStop(1,c+'88');
    ctx.fillStyle=grad;
    ctx.beginPath();
    const rad=4;
    ctx.roundRect(bx,by,barW,bh,rad);
    ctx.fill();
    // Valor encima
    if(v>0){
      ctx.fillStyle='#555'; ctx.font='bold 11px sans-serif'; ctx.textAlign='center';
      ctx.fillText(v, bx+barW/2, by-5);
    }
    // Label abajo
    ctx.fillStyle='#666'; ctx.font='10px sans-serif'; ctx.textAlign='center';
    const lbl=labels[i].length>8?labels[i].substring(0,8)+'…':labels[i];
    ctx.fillText(lbl, bx+barW/2, padT+chartH+15);
  });
}

function exportarReporteExcel(){
  const {desde} = getPeriodoRange();
  const desdeStr = desde.toISOString().split('T')[0];
  const allCierres = _getCierres();
  const cierres = allCierres.filter(c=>(c.fechaRegistro||'')>=desdeStr);
  const data = cierres.map(c=>{
    const exec=USERS.find(u=>u.id===c.ejecutivo);
    return{
      'Fecha':c.fechaRegistro,'Cliente':c.clienteNombre,'Aseguradora':c.aseguradora,
      'Póliza':c.polizaNueva,'Factura':c.facturaAseg,
      'Prima Neta':c.primaNeta,'Prima Total':c.primaTotal,
      'Forma Pago':c.formaPago?.forma||'','Banco':c.formaPago?.banco||'',
      'Cuotas':c.formaPago?.nCuotas||'','Cuota Monto':c.formaPago?.cuotaMonto||'',
      'Vigencia Desde':c.vigDesde,'Vigencia Hasta':c.vigHasta,
      'AXA/VD':c.axavd||'','Ejecutivo':exec?exec.name:'',
      'Observación':c.observacion||''
    };
  });
  const ws=XLSX.utils.json_to_sheet(data);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Reporte Cierres');
  // Sheet resumen por ejecutivo
  const execs=USERS.filter(u=>u.rol==='ejecutivo');
  const resumen=execs.map(u=>{
    const mc=cierres.filter(c=>c.ejecutivo===u.id);
    return{'Ejecutivo':u.name,'Cierres':mc.length,'Prima Total':mc.reduce((s,c)=>s+(c.primaTotal||0),0).toFixed(2)};
  });
  const ws2=XLSX.utils.json_to_sheet(resumen);
  XLSX.utils.book_append_sheet(wb,ws2,'Resumen Ejecutivos');

  // Sheet: Proyección renovaciones 90 días
  const proxVencExp = DB.filter(c=>{ const d=daysUntil(c.hasta); return d>=0&&d<=90; })
    .sort((a,b)=>(a.hasta||'').localeCompare(b.hasta||''));
  const proyData = proxVencExp.map(c=>{
    const ult = allCierres.filter(x=>String(x.clienteId)===String(c.id))
      .sort((a,b)=>(b.fechaRegistro||'').localeCompare(a.fechaRegistro||''))[0];
    const exec = USERS.find(u=>u.id===c.ejecutivo);
    return {
      'Cliente':c.nombre||'','CI/RUC':c.ci||'','Teléfono':c.celular||'',
      'Aseguradora':c.aseguradora||'','Póliza':c.poliza||'',
      'Vence':c.hasta||'','Días restantes':daysUntil(c.hasta),
      'Prima anterior':ult?.primaTotal||0,'Estado':c.estado||'',
      'Ejecutivo':exec?exec.name:''
    };
  });
  const ws3=XLSX.utils.json_to_sheet(proyData);
  XLSX.utils.book_append_sheet(wb,ws3,'Proyección Renovaciones');

  // Sheet: Cobranza del mes actual
  const ahoraExp = new Date();
  const mesIniExp = new Date(ahoraExp.getFullYear(), ahoraExp.getMonth(), 1).toISOString().split('T')[0];
  const mesFinExp = new Date(ahoraExp.getFullYear(), ahoraExp.getMonth()+1, 0).toISOString().split('T')[0];
  const cobRows = [];
  allCierres.forEach(cierre=>{
    _getCuotasFromCierre(cierre).forEach(q=>{
      if((q.fecha||'')<mesIniExp||(q.fecha||'')>mesFinExp) return;
      cobRows.push({
        'Cliente':cierre.clienteNombre||'','Aseguradora':cierre.aseguradora||'',
        'Póliza':cierre.polizaNueva||'','Cuota N°':q.nCuota,'Total cuotas':q.totalCuotas,
        'Fecha':q.fecha,'Monto':q.monto,'Estado':q.estado,'Tipo':q.tipo||''
      });
    });
  });
  const ws4=XLSX.utils.json_to_sheet(cobRows);
  XLSX.utils.book_append_sheet(wb,ws4,'Cobranza Mes Actual');

  const periodo=document.getElementById('rep-periodo')?.value||'periodo';
  XLSX.writeFile(wb,`Reliance_Reporte_${periodo}_${new Date().toISOString().split('T')[0]}.xlsx`);
  showToast('Reporte exportado con 4 hojas');
}

// ══════════════════════════════════════════════════════
//  MÁSCARA FACTURA SRI ECUADOR  (001-001-000000000)
// ══════════════════════════════════════════════════════
function maskFactura(input){
  // Extraer solo dígitos
  let digits = input.value.replace(/\D/g,'');
  if(digits.length > 15) digits = digits.substring(0,15);

  // Aplicar máscara: XXX-XXX-XXXXXXXXX
  let masked = '';
  for(let i=0; i<digits.length; i++){
    if(i===3||i===6) masked += '-';
    masked += digits[i];
  }
  input.value = masked;

  // Validación en tiempo real
  const errEl = document.getElementById('cv-factura-error');
  if(!errEl) return;
  if(digits.length===0){
    errEl.style.display='none'; input.style.borderColor='';
  } else if(digits.length<15){
    errEl.textContent = `Faltan ${15-digits.length} dígito${15-digits.length!==1?'s':''}`;
    errEl.style.display='block'; input.style.borderColor='var(--accent)';
  } else {
    // Validar que los dos primeros bloques no sean 000
    const p1=digits.substring(0,3), p2=digits.substring(3,6);
    if(p1==='000'||p2==='000'){
      errEl.textContent='Los primeros bloques no pueden ser 000';
      errEl.style.display='block'; input.style.borderColor='var(--red)';
    } else {
      errEl.style.display='none'; input.style.borderColor='var(--green)';
    }
  }
}

function allowFacturaKey(e){
  // Permitir: backspace, delete, tab, flechas, ctrl+a/c/v/x
  if([8,9,46,37,38,39,40].includes(e.keyCode)) return true;
  if((e.ctrlKey||e.metaKey)&&[65,67,86,88].includes(e.keyCode)) return true;
  // Solo dígitos
  if(e.key>='0'&&e.key<='9') return true;
  e.preventDefault(); return false;
}

function validarFactura(valor){
  const digits=(valor||'').replace(/\D/g,'');
  if(digits.length!==15) return false;
  if(digits.substring(0,3)==='000') return false;
  if(digits.substring(3,6)==='000') return false;
  return true;
}

// ══════════════════════════════════════════════════════
//  INIT
// ══════════════════════════════════════════════════════
function initAsegSelector(){
  const wrap=document.getElementById('aseg-selector'); if(!wrap) return;
  // Solo mostrar las 7 del Excel primero, luego las adicionales
  const orden=['ZURICH','LATINA','GENERALI','ADS','SWEADEN','MAPFRE','ALIANZA'];
  const extras=Object.keys(ASEGURADORAS).filter(n=>!orden.includes(n));
  const todas=[...orden,...extras];
  wrap.innerHTML=todas.map(name=>{
    const cfg=ASEGURADORAS[name];
    const enOrden=orden.includes(name);
    return `<label style="display:flex;align-items:center;gap:6px;padding:5px 8px;border:1.5px solid ${cfg.color}33;border-radius:6px;cursor:pointer;font-size:11px;font-weight:600;color:${cfg.color};background:${enOrden?cfg.color+'22':cfg.color+'0d'};user-select:none;transition:background .15s" title="${name}">
      <input type="checkbox" class="aseg-check" data-aseg="${name}" ${enOrden?'checked':''}
        style="accent-color:${cfg.color};width:13px;height:13px;cursor:pointer"
        onchange="this.closest('label').style.background=this.checked?'${cfg.color}22':'${cfg.color}08';_actualizarVidaInputs()">
      ${name}
    </label>`;
  }).join('');
  _actualizarVidaInputs();
}

// Mostrar/ocultar inputs de vida y AXA según aseguradoras seleccionadas
function _actualizarVidaInputs(){
  const sel=getSelectedAseg();
  // AXA — solo SWEADEN
  const axaWrap=document.getElementById('cot-axa-wrap');
  if(axaWrap) axaWrap.style.display=sel.includes('SWEADEN')?'flex':'none';
  // Vida por aseguradora
  const vidaAseg=['LATINA','SWEADEN','MAPFRE','ALIANZA'];
  let anyVida=false;
  vidaAseg.forEach(n=>{
    const row=document.getElementById(`cot-vida-${n.toLowerCase()}-row`);
    const visible=sel.includes(n);
    if(row) row.style.display=visible?'flex':'none';
    if(visible) anyVida=true;
  });
  const noneEl=document.getElementById('cot-vida-none');
  if(noneEl) noneEl.style.display=anyVida?'none':'block';
}
function toggleAllAseg(state){
  document.querySelectorAll('.aseg-check').forEach(cb=>{
    cb.checked=state;
    const cfg=ASEGURADORAS[cb.dataset.aseg];
    if(cfg) cb.closest('label').style.background=state?cfg.color+'22':cfg.color+'08';
  });
}
function getSelectedAseg(){
  return Array.from(document.querySelectorAll('.aseg-check:checked')).map(cb=>cb.dataset.aseg);
}

// ── Sync automático SP→RelianceDesk (polling ligero) ────────
// Se ejecuta cada 60s para reflejar cambios de otros usuarios
let _syncInterval = null;
let _lastSyncHash = '';
let _visibilityHandler = null;

async function _syncFromSP(){
  if(!_spReady || !currentUser) return;
  try{
    // Capturar hashes previos ANTES de las llamadas a SP
    const prevCot = (_cache.cotizaciones||[]);
    const prevC   = (_cache.cierres||[]);
    const prevT   = (_cache.tareas||[]);
    const prevHashCot = prevCot.length + '|' + (prevCot[prevCot.length-1]?._spId||'');
    const prevHashC   = prevC.length   + '|' + (prevC[prevC.length-1]?._spId||'');
    const prevHashT   = prevT.length   + '|' + (prevT[prevT.length-1]?._spId||'');

    // Capturar IDs dirty ANTES del await (flush async puede borrar ._dirty durante la espera)
    const dirtyIdsCotiz   = new Set((_cache.cotizaciones||[]).filter(c=>c._dirty).map(c=>String(c.id)));
    const dirtyIdsCierres = new Set((_cache.cierres||[]).filter(c=>c._dirty).map(c=>String(c.id)));

    // Cargar las listas en paralelo (spGetAll actualiza _cache internamente)
    const [clientes, cotizaciones, cierres, tareas, gestCobr, comisiones] = await Promise.all([
      spGetAll('clientes'),
      spGetAll('cotizaciones'),
      spGetAll('cierres'),
      spGetAll('tareas'),
      spGetAll('cobranzas'),
      spGetAll('comisiones'),
    ]);

    // Una sola query al DOM para toda la función
    const pg = document.querySelector('.page.active')?.id;

    // Clientes
    const hash = clientes.length + '|' + (clientes[0]?._spId||'');
    if(hash !== _lastSyncHash){
      _lastSyncHash = hash;
      DB = clientes;
      localStorage.setItem('reliance_clientes', JSON.stringify(clientes));
      if(pg==='page-clientes'){ renderClientes(); renderVencimientos(); }
      renderDashboard();
    }
    // Cotizaciones: re-fusionar registros dirty locales que aún no llegaron a SP
    // (usamos dirtyIdsCotiz capturado antes del await para evitar race con _flushCotizaciones)
    const dirtyLocalCotiz = prevCot.filter(lc => dirtyIdsCotiz.has(String(lc.id)));
    if(dirtyLocalCotiz.length){
      dirtyLocalCotiz.forEach(lc => {
        const spIdx = cotizaciones.findIndex(sc => String(sc.id)===String(lc.id));
        if(spIdx >= 0) cotizaciones[spIdx] = lc; // local dirty tiene prioridad
        else cotizaciones.push(lc);               // registro aún no en SP
      });
      _cache.cotizaciones = cotizaciones;
    }
    const hashCot = cotizaciones.length + '|' + (cotizaciones[cotizaciones.length-1]?._spId||'');
    if(hashCot !== prevHashCot){
      localStorage.setItem('reliance_cotizaciones', JSON.stringify(cotizaciones));
      if(pg==='page-cotizaciones') renderCotizaciones();
      actualizarBadgeCotizaciones();
    }
    // Cierres: mismo patrón — preservar dirty locales (usando IDs capturados antes del await)
    const dirtyLocalCierres = prevC.filter(lc => dirtyIdsCierres.has(String(lc.id)));
    if(dirtyLocalCierres.length){
      dirtyLocalCierres.forEach(lc => {
        const spIdx = cierres.findIndex(sc => String(sc.id)===String(lc.id));
        if(spIdx >= 0) cierres[spIdx] = lc;
        else cierres.push(lc);
      });
      _cache.cierres = cierres;
    }
    const hashC = cierres.length + '|' + (cierres[cierres.length-1]?._spId||'');
    if(hashC !== prevHashC){
      localStorage.setItem('reliance_cierres', JSON.stringify(cierres));
      if(pg==='page-cierres') renderCierres();
    }
    // Tareas
    const hashT = tareas.length + '|' + (tareas[tareas.length-1]?._spId||'');
    if(hashT !== prevHashT){
      _saveTareas(tareas.map(t=>({...t, _spId:t._spId})));
      actualizarBadgeTareas();
      renderDashTareas();
      if(pg==='page-calendario') renderCalendario();
    }
    // Gestiones de cobranza (append-only: solo agregar las que no existen localmente)
    // Solo se procesan entradas con cierreId — evita contaminar con entradas huérfanas antiguas
    if(gestCobr && gestCobr.length){
      const localG = _getGestionCobranza();
      let changed = false;
      gestCobr.forEach(sg=>{
        if(!sg.cierreId) return; // ignorar entradas sin cierreId (no son gestiones de cobranza)
        const localId = String(sg.id||sg.crm_id||'');
        if(localId && !localG.find(lg=>String(lg.id)===localId)){
          // Marcar como sincronizado (tiene _spId desde SP)
          localG.unshift({...sg, _spId: sg._spId||sg._spId, _dirty: false});
          changed=true;
        }
      });
      if(changed){
        localG.sort((a,b)=>(b.fecha||'').localeCompare(a.fecha||''));
        _saveGestionCobranza(localG);
        if(pg==='page-cobranza') renderCobranza(_currentFiltroCobranza||'mes');
      }
    }
    // Comisiones y tasas — actualizar localStorage si SP cambió (admin modificó en otro equipo)
    if(comisiones && comisiones.length){
      const prevComisStr = localStorage.getItem('reliance_comisiones') || '{}';
      const newComisObj = {}, newTasasObj = {};
      comisiones.forEach(item=>{
        if(!item.Title) return;
        if(item.comisionPct !== undefined && item.comisionPct !== null)
          newComisObj[item.Title] = item.comisionPct;
        const t=[item.tasa_r1,item.tasa_r2,item.tasa_r3,item.tasa_r4,item.tasa_r5].map(v=>parseFloat(v)||0); if(t.some(v=>v>0)) newTasasObj[item.Title]=t;
      });
      const newComisStr = JSON.stringify(newComisObj);
      if(newComisStr !== prevComisStr){
        _cache.comisiones = comisiones;
        localStorage.setItem('reliance_comisiones', newComisStr);
        if(Object.keys(newTasasObj).length) localStorage.setItem('_reliance_tasas_rangos', JSON.stringify(newTasasObj));
        if(pg==='page-admin') { renderComisionesAdmin(); renderTasasAdmin(); }
      }
    }
  }catch(e){ /* sync silencioso — no mostrar error */ }
}

function startAutoSync(){
  if(_syncInterval) clearInterval(_syncInterval);
  _syncInterval = setInterval(_syncFromSP, 60000); // cada 60 segundos
  // También sync al volver a la pestaña — remover handler previo para evitar acumulación
  if(_visibilityHandler) document.removeEventListener('visibilitychange', _visibilityHandler);
  _visibilityHandler = ()=>{ if(document.visibilityState==='visible' && currentUser) _syncFromSP(); };
  document.addEventListener('visibilitychange', _visibilityHandler);
}

// ─── Auto-vencimiento de cotizaciones antiguas ────────────────────────────────
// Cotizaciones ENVIADA/VISTA con más de 30 días sin respuesta → VENCIDA
function _vencerCotizacionesAntiguas(){
  const cutoff = new Date(); cutoff.setDate(cutoff.getDate()-30);
  const cutoffStr = cutoff.toISOString().split('T')[0];
  const all = _getCotizaciones();
  let n = 0;
  all.forEach(c=>{
    if(['ENVIADA','VISTA'].includes(c.estado) && c.fecha && c.fecha < cutoffStr){
      c.estado  = 'VENCIDA';
      c._dirty  = true;
      n++;
    }
  });
  if(n){
    _saveCotizaciones(all);
    actualizarBadgeCotizaciones();
    console.info(`_vencerCotizacionesAntiguas: ${n} cotización(es) marcadas VENCIDA`);
  }
}

// ─── Auto-reset renovación anual ─────────────────────────────────────────────
// Detecta clientes en RENOVADO cuya póliza vence en ≤90 días y los reactiva
// como PENDIENTE para iniciar el ciclo de renovación, preservando datos de la
// póliza anterior en los campos polizaAnterior / aseguradoraAnterior.
function _resetearCicloRenovacion(){
  if(!currentUser) return;
  const d90 = new Date(); d90.setDate(d90.getDate()+90);
  const d90Str = d90.toISOString().split('T')[0];

  const paraReset = myClientes().filter(c=>
    c.estado==='RENOVADO' && c.hasta && c.hasta <= d90Str
  );
  if(!paraReset.length) return;

  paraReset.forEach(c=>{
    c.polizaAnterior      = c.polizaNueva || c.poliza || '';
    c.aseguradoraAnterior = c.aseguradora || '';
    c.estado              = 'PENDIENTE';
    c._dirty              = true;
    _bitacoraAdd(c,
      `Ciclo de renovación iniciado automáticamente — póliza anterior: ${c.polizaAnterior} (vigente hasta ${c.hasta})`,
      'sistema'
    );
  });

  saveDB();
  showToast(
    `🔄 ${paraReset.length} cliente${paraReset.length>1?'s':''} reactivado${paraReset.length>1?'s':''} para renovación`,
    'info'
  );
}

async function initApp(){
  // Cargar todos los datos desde SharePoint en paralelo
  const [cotiz, cierres, spUsers, spTareas, spGestCobr, spComisiones] = await Promise.all([
    spGetAll('cotizaciones'),
    spGetAll('cierres'),
    spGetAll('usuarios'),
    spGetAll('tareas'),
    spGetAll('cobranzas'),
    spGetAll('comisiones'),
  ]);
  await loadDBAsync();
  _cache.cotizaciones = cotiz;
  _cache.cierres      = cierres;
  // Fusionar gestiones de cobranza de SP con localStorage
  if(spGestCobr && spGestCobr.length){
    const localG = _getGestionCobranza();
    const merged = [...localG];
    spGestCobr.forEach(sg=>{
      if(!merged.find(lg=>String(lg.id)===String(sg.id||sg.crm_id))) merged.push(sg);
    });
    merged.sort((a,b)=>(b.fecha||'').localeCompare(a.fecha||''));
    _saveGestionCobranza(merged);
  }
  _cache.usuarios     = spUsers;
  // Fusionar tareas SP con localStorage (SP tiene prioridad por _spId)
  if(spTareas && spTareas.length){
    const localT = _getTareas();
    const merged = [...localT];
    spTareas.forEach(st=>{
      const idx = merged.findIndex(lt=>lt.id===st.id || lt._spId===st._spId);
      if(idx>=0) merged[idx] = {...merged[idx], ...st};
      else merged.push(st);
    });
    _saveTareas(merged);
  }

  // Comisiones y tasas desde SP → localStorage (SP es la fuente de verdad)
  if(spComisiones && spComisiones.length){
    _cache.comisiones = spComisiones;
    // Comisiones planas (por aseg) para _getComisiones()
    const comisObj = {};
    // Estructura V2 (comienza con los defaults y SP los sobreescribe)
    const tasasV2 = {...TASAS_V2_DEFAULT};
    spComisiones.forEach(item=>{
      if(!item.crm_id && !item.Title) return;
      const key = item.crm_id || item.Title;
      if(item.comisionPct !== undefined && item.comisionPct !== null)
        comisObj[key] = parseFloat(item.comisionPct)||0;
      // Reconstruir fila V2 sólo si tiene tasas configuradas
      const t   = [item.tasa_r1,item.tasa_r2,item.tasa_r3,item.tasa_r4].map(v=>parseFloat(v)||0);
      const lim = [item.limite_r1,item.limite_r2,item.limite_r3].map(v=>parseFloat(v)||0).filter(v=>v>0);
      const tasasActivas = lim.length ? t.slice(0, lim.length+1) : t.filter(v=>v>0);
      if(tasasActivas.some(v=>v>0) && item.crm_id){
        const base = TASAS_V2_DEFAULT[item.crm_id] || {};
        tasasV2[item.crm_id] = {
          ...base,
          region:     ((item.region||base.region)||'').toUpperCase(),
          tipo:       base.tipo || '',
          tasas:      tasasActivas,
          limites:    lim.length ? lim : (base.limites||[]),
          comisionPct: parseFloat(item.comisionPct)||base.comisionPct||0,
        };
      }
    });
    if(Object.keys(comisObj).length) localStorage.setItem('reliance_comisiones',  JSON.stringify(comisObj));
    if(Object.keys(tasasV2).length)  localStorage.setItem('_reliance_tasas_v2',   JSON.stringify(tasasV2));
    // Migración automática: si faltan filas V2 en SP (formato antiguo), crearlas en background
    const crmIdsEnSP = new Set(spComisiones.map(x=>x.crm_id).filter(Boolean));
    const faltanFilasV2 = Object.keys(TASAS_V2_DEFAULT).some(k => !crmIdsEnSP.has(k));
    if(faltanFilasV2) _flushComisiones();
    // Limpiar filas obsoletas (crm_id de formato antiguo que ya no corresponden a V2)
    _cleanupComisionesLegacy();
  } else {
    // SP vacío (lista recién creada) → subir los valores actuales como datos iniciales
    _flushComisiones();
  }

  // Mezclar usuarios SP con los hardcodeados (SP tiene prioridad)
  spUsers.forEach(u=>{
    const local=USERS.find(x=>String(x.id)===String(u.userId||u.id));
    if(local) Object.assign(local, u);
    else USERS.push({...u, id:u.userId||u.id});
  });

  // Inicializar UI
  const today=new Date();
  const el=document.getElementById('current-date');
  if(el) el.textContent=today.toLocaleDateString('es-EC',{weekday:'short',day:'numeric',month:'long',year:'numeric'});
  const todayStr=today.toISOString().split('T')[0];
  if(document.getElementById('cot-desde')) document.getElementById('cot-desde').value=todayStr;
  if(document.getElementById('nc-desde')){
    document.getElementById('nc-desde').value=todayStr;
    const ny=new Date(today); ny.setFullYear(ny.getFullYear()+1);
    document.getElementById('nc-hasta').value=ny.toISOString().split('T')[0];
  }
  initAsegSelector();
  migrarCotizacionesIds();
  _vencerCotizacionesAntiguas();
  _resetearCicloRenovacion();
  renderDashboard();
  actualizarBadgeCotizaciones();
  actualizarBadgeTareas();
  actualizarBadgeCobranza();
  actualizarBadgeCola();
  renderDashTareas();

  // Ocultar login, mostrar app
  hideLoader();
  document.getElementById('login-screen').style.display='none';
  const appEl=document.querySelector('.app');
  if(appEl) appEl.style.display='flex';
  startAutoSync(); // Iniciar sync bidireccional automático
  // Notificaciones del navegador (leve delay para que cargue todo)
  setTimeout(initNotificaciones, 3000);
  // Cerrar dropdown de sugerencias del cotizador al hacer click fuera
  document.addEventListener('click', e=>{
    const box=document.getElementById('cot-sugerencias');
    const inp=document.getElementById('cot-buscar-cliente');
    if(box && inp && !inp.contains(e.target) && !box.contains(e.target)){
      box.style.display='none';
    }
  });
}
