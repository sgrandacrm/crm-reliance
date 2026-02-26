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
    DB.forEach(c => { if(!c._spId || c._dirty) c._dirty = true; });
    _flushDB();
  }
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
  await initApp();
}
function doLogout(){
  currentUser=null;
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
const pageTitles={dashboard:'Dashboard',cierres:'Cierres de Venta',clientes:'Cartera de Clientes',vencimientos:'Vencimientos de Pólizas',calendario:'Calendario de Vencimientos',seguimiento:'Seguimiento de Clientes',cotizador:'Cotizador de Primas',comparativo:'Comparativo de Coberturas',tasas:'Tabla de Tasas',admin:'Panel de Administración','nuevo-cliente':'Registrar Cliente'};
function showPage(id){
  document.querySelectorAll('.page').forEach(p=>p.classList.remove('active'));
  const navItems = document.querySelectorAll('.nav-item');
  navItems.forEach(n=>n.classList.remove('active'));
  document.getElementById('page-'+id).classList.add('active');
  document.getElementById('page-title').textContent=pageTitles[id]||id;
  navItems.forEach(n=>{
    if(n.getAttribute('onclick')&&n.getAttribute('onclick').includes("'"+id+"'")) n.classList.add('active');
  });
  const renders={clientes:()=>{renderClientes();renderVencimientos();},vencimientos:()=>{showPage('clientes');switchClienteTab('vencimientos');},calendario:()=>{renderCalendario();renderTareasCalendario();},seguimiento:renderSeguimiento,dashboard:renderDashboard,admin:()=>{renderAdmin();showAdminTab('importar',document.querySelector('#admin-tabs .pill'));},comparativo:renderComparativo,cierres:renderCierres,reportes:renderReportes,cotizaciones:renderCotizaciones};
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
  document.getElementById('stat-mes').textContent = `Febrero 2026`;
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

  // Badge
  const vencBadge = mine.filter(c=>{ const d=daysUntil(c.hasta); return d>=0&&d<=30; }).length;
  document.getElementById('badge-venc').textContent = vencBadge||'0';
}

// ══════════════════════════════════════════════════════
//  CLIENTES
// ══════════════════════════════════════════════════════
let clientesFiltrados = [];
function switchClienteTab(tab){
  // tabs
  document.getElementById('tab-cartera').style.color      = tab==='cartera'     ?'var(--accent)':'var(--muted)';
  document.getElementById('tab-vencimientos').style.color = tab==='vencimientos'?'var(--accent)':'var(--muted)';
  document.getElementById('tab-cartera').style.borderBottomColor      = tab==='cartera'     ?'var(--accent)':'transparent';
  document.getElementById('tab-vencimientos').style.borderBottomColor = tab==='vencimientos'?'var(--accent)':'transparent';
  // subtabs
  document.getElementById('subtab-cartera').style.display      = tab==='cartera'     ?'':'none';
  document.getElementById('subtab-vencimientos').style.display = tab==='vencimientos'?'':'none';
  // render el subtab activo
  if(tab==='cartera')     renderClientes();
  if(tab==='vencimientos') renderVencimientos();
}
function renderClientes(){
  initFilters();
  filterClientes();
}
function filterClientes(){
  const q=(document.getElementById('search-clientes').value||'').toLowerCase();
  const tipo=document.getElementById('filter-tipo').value;
  const aseg=document.getElementById('filter-aseg').value;
  const region=document.getElementById('filter-region').value;
  const estado=document.getElementById('filter-estado').value;
  clientesFiltrados = myClientes().filter(c=>{
    const mq=!q||(c.nombre||'').toLowerCase().includes(q)||(c.ci||'').includes(q)||(c.placa||'').toLowerCase().includes(q)||(c.aseguradora||'').toLowerCase().includes(q);
    const mt=!tipo||c.tipo===tipo;
    const ma=!aseg||c.aseguradora===aseg;
    const mr=!region||c.region===region;
    const me=!estado||(c.estado||'PENDIENTE')===estado;
    return mq&&mt&&ma&&mr&&me;
  }).sort((a,b)=>daysUntil(a.hasta)-daysUntil(b.hasta));
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
        <button class="btn btn-xs" style="background:#25D366;color:#fff" onclick="openWhatsApp('${c.id}','vencimiento')" title="WhatsApp">💬</button>
        <button class="btn btn-xs" style="background:#0078d4;color:#fff" onclick="openEmail('${c.id}','vencimiento')" title="Email">✉️</button>
        <button class="btn btn-ghost btn-xs" onclick="nuevaTareaDesdeCliente('${c.id}')" title="Nueva tarea">📌</button>
        ${(['RENOVADO','EMITIDO'].includes(c.estado)&&!c.factura)?`<button class="btn btn-green btn-xs" onclick="abrirCierreDesdeCliente('${c.id}')" title="Registrar cierre de venta">📋</button>`:''}
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
        ${(c.historialWa&&c.historialWa.length)?`
        <div class="detail-section">
          <div class="detail-section-title" style="display:flex;align-items:center;gap:8px">
            <span style="font-size:16px">💬</span> Historial de WhatsApp
          </div>
          <table style="width:100%;font-size:11px">
            <thead><tr>
              <th style="padding:4px 8px;text-align:left;color:var(--muted);font-size:10px">Fecha</th>
              <th style="padding:4px 8px;text-align:left;color:var(--muted);font-size:10px">Resumen</th>
            </tr></thead>
            <tbody>
              ${c.historialWa.map(h=>`<tr style="border-bottom:1px solid var(--warm)">
                <td style="padding:5px 8px;font-family:'DM Mono',monospace;white-space:nowrap">${h.fecha}</td>
                <td style="padding:5px 8px;color:var(--muted)">${h.resumen}</td>
              </tr>`).join('')}
            </tbody>
          </table>
        </div>`:''}

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
          <div class="detail-row"><span class="detail-key">Aseguradora</span><span class="detail-val font-bold">${c.aseguradora||'—'}</span></div>
          <div class="detail-row"><span class="detail-key">Nº Póliza</span><span class="detail-val mono">${c.poliza||'—'}</span></div>
          <div class="detail-row"><span class="detail-key">OBS</span><span class="detail-val">${c.obs||'—'}</span></div>
          <div class="detail-row"><span class="detail-key">Vigencia Desde</span><span class="detail-val mono">${c.desde||'—'}</span></div>
          <div class="detail-row"><span class="detail-key">Vigencia Hasta</span><span class="detail-val mono">${c.hasta||'—'}</span></div>
          ${c.tasa?`<div class="detail-row"><span class="detail-key">Tasa</span><span class="detail-val mono">${c.tasa}%</span></div>`:''}
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
      </div></div>
    </div>`;
  document.getElementById('modal-btn-cotizar').style.display='';
  document.getElementById('modal-btn-editar').style.display='';
  document.getElementById('modal-btn-eliminar').style.display='';
  document.getElementById('modal-btn-cotizar').onclick=()=>{ closeModal('modal-cliente'); prefillCotizador(c); showPage('cotizador'); setTimeout(calcCotizacion,200); };
  document.getElementById('modal-btn-editar').onclick=()=>{ closeModal('modal-cliente'); openEditar(id); };
  document.getElementById('modal-btn-eliminar').onclick=()=>{ if(confirm(`¿Eliminar a ${c.nombre}?`)){ const cliToDel=DB.find(x=>String(x.id)===String(id)); if(cliToDel?._spId && _spReady) spDelete('clientes', cliToDel._spId); DB=DB.filter(x=>String(x.id)!==String(id)); saveDB(); closeModal('modal-cliente'); renderClientes(); renderDashboard(); showToast('Cliente eliminado','error'); }};
  openModal('modal-cliente');
}

// ══════════════════════════════════════════════════════
//  EDITAR CLIENTE
// ══════════════════════════════════════════════════════
function openEditar(id){
  const c=DB.find(x=>String(x.id)===String(id)); if(!c) return;
  document.getElementById('modal-editar-body').innerHTML=`
    <div class="form-grid form-grid-2" style="gap:12px">
      <div class="form-group full"><label class="form-label">Nombre</label><input class="form-input" id="ed-nombre" value="${c.nombre||''}"></div>
      <div class="form-group"><label class="form-label">CI</label><input class="form-input" id="ed-ci" value="${c.ci||''}"></div>
      <div class="form-group"><label class="form-label">Celular</label><input class="form-input" id="ed-cel" value="${c.celular||''}"></div>
      <div class="form-group"><label class="form-label">Correo</label><input class="form-input" id="ed-email" value="${c.correo||''}"></div>
      <div class="form-group"><label class="form-label">Ciudad</label><input class="form-input" id="ed-ciudad" value="${c.ciudad||''}"></div>
      <div class="form-group"><label class="form-label">Aseguradora</label><input class="form-input" id="ed-aseg" value="${c.aseguradora||''}"></div>
      <div class="form-group"><label class="form-label">Póliza</label><input class="form-input" id="ed-poliza" value="${c.poliza||''}"></div>
      <div class="form-group"><label class="form-label">Vigencia Hasta</label><input class="form-input" type="date" id="ed-hasta" value="${c.hasta||''}"></div>
      <div class="form-group"><label class="form-label">Val. Asegurado ($)</label><input class="form-input" type="number" id="ed-va" value="${c.va||''}"></div>
      <div class="form-group"><label class="form-label">Tasa (%)</label><input class="form-input" type="number" step="0.01" id="ed-tasa" value="${c.tasa||''}"></div>
      <div class="form-group"><label class="form-label">Marca</label><input class="form-input" id="ed-marca" value="${c.marca||''}"></div>
      <div class="form-group"><label class="form-label">Placa</label><input class="form-input" id="ed-placa" value="${c.placa||''}"></div>
      ${currentUser&&currentUser.rol==='admin'?`<div class="form-group"><label class="form-label">Asignar a Ejecutivo</label><select class="form-select" id="ed-ejecutivo">${USERS.filter(u=>u.rol==='ejecutivo').map(u=>`<option value="${u.id}"${c.ejecutivo===u.id?' selected':''}>${u.name}</option>`).join('')}</select></div>`:''}
      <div class="form-group full"><label class="form-label">Comentarios</label><textarea class="form-textarea" id="ed-comentario">${c.comentario||''}</textarea></div>
    </div>`;
  document.getElementById('modal-btn-guardar-editar').onclick=async()=>{
    c.nombre=document.getElementById('ed-nombre').value;
    c.ci=document.getElementById('ed-ci').value;
    c.celular=document.getElementById('ed-cel').value;
    c.correo=document.getElementById('ed-email').value;
    c.ciudad=document.getElementById('ed-ciudad').value;
    c.aseguradora=document.getElementById('ed-aseg').value;
    c.poliza=document.getElementById('ed-poliza').value;
    c.hasta=document.getElementById('ed-hasta').value;
    c.va=parseFloat(document.getElementById('ed-va').value)||0;
    c.tasa=parseFloat(document.getElementById('ed-tasa').value)||null;
    c.marca=document.getElementById('ed-marca').value;
    c.placa=document.getElementById('ed-placa').value;
    c.comentario=document.getElementById('ed-comentario').value;
    if(document.getElementById('ed-ejecutivo')) c.ejecutivo=document.getElementById('ed-ejecutivo').value;
    await sincronizarCotizPorCliente(c.id, c.nombre, c.ci, c.estado);
    saveDB(); closeModal('modal-editar'); renderClientes(); renderDashboard(); showToast('Cliente actualizado');
  };
  openModal('modal-editar');
}

// ══════════════════════════════════════════════════════
//  VENCIMIENTOS
// ══════════════════════════════════════════════════════
let vencFilter='all';
function renderVencimientos(){
  const mine=myClientes();
  const all=mine.filter(c=>c.hasta&&daysUntil(c.hasta)<=90).sort((a,b)=>new Date(a.hasta)-new Date(b.hasta));
  const venc0=mine.filter(c=>daysUntil(c.hasta)<0).length;
  const venc30=mine.filter(c=>{const d=daysUntil(c.hasta);return d>=0&&d<=30}).length;
  const venc60=mine.filter(c=>{const d=daysUntil(c.hasta);return d>30&&d<=60}).length;
  const venc90=mine.filter(c=>{const d=daysUntil(c.hasta);return d>60&&d<=90}).length;
  document.getElementById('vstat-0').textContent=venc0;
  document.getElementById('vstat-30').textContent=venc30;
  document.getElementById('vstat-60').textContent=venc60;
  document.getElementById('vstat-90').textContent=venc90;
  document.getElementById('badge-venc').textContent=venc30+venc0;
  let filtered=all;
  if(vencFilter==='vencida') filtered=all.filter(c=>daysUntil(c.hasta)<0);
  else if(vencFilter==='30') filtered=all.filter(c=>{const d=daysUntil(c.hasta);return d>=0&&d<=30});
  else if(vencFilter==='60') filtered=all.filter(c=>{const d=daysUntil(c.hasta);return d>30&&d<=60});
  else if(vencFilter==='90') filtered=all.filter(c=>{const d=daysUntil(c.hasta);return d>60&&d<=90});
  document.getElementById('vencimientos-list').innerHTML = filtered.length
    ? filtered.map(c=>{ const days=daysUntil(c.hasta); return `
      <div class="venc-alert ${vencClass(days)}">
        <div class="venc-days">${days<0?`<span style="font-size:11px">hace</span><br>${Math.abs(days)}d`:days+'d'}</div>
        <div class="venc-info">
          <div class="venc-name">${c.nombre}</div>
          <div class="venc-meta">${c.aseguradora} · ${c.marca} ${c.anio} · Placa: ${c.placa||'—'} · ${fmt(c.va)} · Vence: ${c.hasta}</div>
        </div>
        <div style="display:flex;flex-direction:column;gap:4px;align-items:flex-end">
          ${estadoBadge(c.estado||'PENDIENTE')}
          <button class="btn btn-ghost btn-xs" onclick="openSeguimiento('${c.id}')">📞 Gestionar</button>
          <button class="btn btn-ghost btn-xs" onclick="prefillCotizador_show('${c.id}')">🧮 Cotizar</button>
          <button class="btn btn-xs" style="background:#25D366;color:#fff" onclick="openWhatsApp('${c.id}','vencimiento')">💬 WhatsApp</button>
          <button class="btn btn-xs" style="background:#0078d4;color:#fff" onclick="openEmail('${c.id}','vencimiento')">✉️ Email</button>
          <button class="btn btn-ghost btn-xs" onclick="nuevaTareaDesdeCliente('${c.id}')">📌 Tarea</button>
        </div>
      </div>` }).join('')
    : '<div class="empty-state"><div class="empty-icon">✅</div><p>No hay vencimientos en este rango</p></div>';
}
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
function renderSeguimiento(){
  filterSeguimiento();
}
function filterSeguimiento(){
  const q=(document.getElementById('search-seg').value||'').toLowerCase();
  let data=myClientes().filter(c=>{
    const mq=!q||(c.nombre||'').toLowerCase().includes(q);
    const me=!segFilterEstado||(c.estado||'PENDIENTE')===segFilterEstado;
    return mq&&me;
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
        <button class="btn btn-xs" style="background:#25D366;color:#fff" onclick="openWhatsApp('${c.id}','vencimiento')">💬 WA</button>
        <button class="btn btn-xs" style="background:#0078d4;color:#fff" onclick="openEmail('${c.id}','vencimiento')">✉️ Mail</button>
        <button class="btn btn-ghost btn-xs" onclick="nuevaTareaDesdeCliente('${c.id}')" title="Nueva tarea">📌</button>
        ${(c.estado==='RENOVADO'&&!c.factura)?`<button class="btn btn-green btn-xs" onclick="abrirCierreDesdeCliente('${c.id}')">📋 Cierre</button>`:''}
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
  // Mostrar banner cierre si ya está RENOVADO
  const banner=document.getElementById('seg-cierre-banner');
  if(banner) banner.style.display=['RENOVADO','EMITIDO','PÓLIZA VIGENTE'].includes(currentSegEstado)?'block':'none';
  // Renderizar bitácora existente
  renderBitacora(c.bitacora||[]);
  openModal('modal-seguimiento');
}
function setEstado(e){
  currentSegEstado=e;
  document.querySelectorAll('.estado-btn').forEach(b=>{
    b.classList.remove('active');
    if(b.getAttribute('onclick')&&b.getAttribute('onclick').includes("'"+e+"'")) b.classList.add('active');
  });
  // Mostrar banner "Registrar Cierre" solo al seleccionar RENOVADO
  const banner=document.getElementById('seg-cierre-banner');
  if(banner) banner.style.display=(e==='RENOVADO')?'block':'none';
  // Solo actualiza la UI — el estado se persiste al presionar "Guardar" o "Guardar y Cierre"
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
}

// Renderiza el historial de bitácora dentro del modal
function renderBitacora(bitacora){
  const el = document.getElementById('bitacora-lista');
  if(!el) return;
  const lista = Array.isArray(bitacora) ? bitacora : [];
  const cntEl = document.getElementById('bitacora-count');
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
  closeModal('modal-seguimiento'); renderSeguimiento(); renderVencimientos(); renderDashboard();
  showToast('Seguimiento actualizado');
}
function guardarSeguimientoYCierre(){
  const c=DB.find(x=>String(x.id)===String(currentSegIdx)); if(!c) return;
  const notaYC = document.getElementById('seg-nota').value.trim();
  // Solo guarda la nota en bitácora; el estado RENOVADO lo fija guardarCierreVenta al completar el cierre
  if(notaYC){
    _bitacoraAdd(c, notaYC, 'cierre');
    saveDB();
  }
  closeModal('modal-seguimiento');
  // Abrir modal de cierre saltando la validación de estado (aún no es RENOVADO en DB)
  abrirCierreDesdeCliente(c.id, true);
}

// ══════════════════════════════════════════════════════
//  COTIZADOR
// ══════════════════════════════════════════════════════
const ESTADOS_RELIANCE = {"PENDIENTE": {"cod": 1, "label": "Pendiente", "color": "#9e9e9e", "badge": "badge-gray", "icon": "⚪", "grupo": "gestion", "def": "Cliente en gestión de renovación. Sin confirmación aún."}, "INSPECCIÓN": {"cod": 2, "label": "Inspección", "color": "#ff9800", "badge": "badge-orange", "icon": "🔍", "grupo": "gestion", "def": "Vehículo pendiente de inspección requerida por aseguradora."}, "EMISIÓN": {"cod": 3, "label": "Emisión", "color": "#2196f3", "badge": "badge-blue", "icon": "📝", "grupo": "gestion", "def": "Póliza en trámite de emisión. Aún no activa."}, "EMITIDO": {"cod": 4, "label": "Emitido", "color": "#1976d2", "badge": "badge-blue2", "icon": "📄", "grupo": "gestion", "def": "Póliza emitida formalmente y activa."}, "RENOVADO": {"cod": 5, "label": "Renovado", "color": "#2d6a4f", "badge": "badge-green", "icon": "✅", "grupo": "positivo", "def": "Renovación completada exitosamente."}, "PÓLIZA VIGENTE": {"cod": 6, "label": "Póliza Vigente", "color": "#388e3c", "badge": "badge-green2", "icon": "🛡", "grupo": "positivo", "def": "Póliza activa dentro del período de cobertura."}, "CRÉDITO VENCIDO": {"cod": 7, "label": "Crédito Vencido", "color": "#e65100", "badge": "badge-orange2", "icon": "⏰", "grupo": "riesgo", "def": "Cliente mantiene cuotas vencidas con el banco que financia el vehículo."}, "CRÉDITO CANCELADO": {"cod": 8, "label": "Crédito Cancelado", "color": "#5d4037", "badge": "badge-brown", "icon": "💳", "grupo": "riesgo", "def": "Cliente canceló totalmente su crédito con el banco."}, "SINIESTRO": {"cod": 9, "label": "Siniestro", "color": "#c62828", "badge": "badge-red2", "icon": "🚨", "grupo": "riesgo", "def": "Cliente con evento reportado en trámite con la aseguradora."}, "PÉRDIDA TOTAL": {"cod": 10, "label": "Pérdida Total", "color": "#b71c1c", "badge": "badge-red", "icon": "🚗", "grupo": "riesgo", "def": "Vehículo declarado pérdida total por la aseguradora."}, "PÓLIZA ANULADA": {"cod": 11, "label": "Póliza Anulada", "color": "#c84b1a", "badge": "badge-orange3", "icon": "❌", "grupo": "cierre", "def": "Póliza cancelada antes del vencimiento."}, "ENDOSO": {"cod": 12, "label": "Endoso", "color": "#6b5b95", "badge": "badge-purple", "icon": "📋", "grupo": "cierre", "def": "Cliente realizó el seguro directamente sin intermediación de Reliance."}, "NO RENOVADO": {"cod": 13, "label": "No Renovado", "color": "#795548", "badge": "badge-brown2", "icon": "🚫", "grupo": "cierre", "def": "Cliente decidió no renovar la póliza al vencimiento."}, "INUBICABLE": {"cod": 14, "label": "Inubicable", "color": "#607d8b", "badge": "badge-slate", "icon": "📵", "grupo": "cierre", "def": "No se logra contacto con el cliente tras varios intentos."}, "NO ASEGURABLE": {"cod": 15, "label": "No Asegurable", "color": "#37474f", "badge": "badge-dark", "icon": "⛔", "grupo": "cierre", "def": "Riesgo rechazado por aseguradora."}, "AUTO VENDIDO": {"cod": 16, "label": "Auto Vendido", "color": "#4caf50", "badge": "badge-teal", "icon": "🔑", "grupo": "cierre", "def": "Cliente ya no posee el vehículo asegurado."}, "CLIENTE FALLECIDO": {"cod": 17, "label": "Cliente Fallecido", "color": "#546e7a", "badge": "badge-slate2", "icon": "🕊", "grupo": "cierre", "def": "Titular fallecido. Se cierra gestión o se coordina con herederos."}};

const GRUPOS_ESTADOS = {"gestion": {"titulo": "📋 En Gestión", "estados": ["PENDIENTE", "INSPECCIÓN", "EMISIÓN", "EMITIDO"]}, "positivo": {"titulo": "✅ Positivos", "estados": ["RENOVADO", "PÓLIZA VIGENTE"]}, "riesgo": {"titulo": "⚠️ Riesgo", "estados": ["CRÉDITO VENCIDO", "CRÉDITO CANCELADO", "SINIESTRO", "PÉRDIDA TOTAL"]}, "cierre": {"titulo": "🔒 Cierre/Baja", "estados": ["PÓLIZA ANULADA", "ENDOSO", "NO RENOVADO", "INUBICABLE", "NO ASEGURABLE", "AUTO VENDIDO", "CLIENTE FALLECIDO"]}};

const ASEGURADORAS={
  ZURICH:{
    color:'#4a4a4a',pnMin:500,tcMax:12,debMax:10,
    tasa:va=>va<=29999?0.045:va<=39999?0.025:0.022,
    resp_civil:30000,muerte_ocupante:10000,muerte_titular:10000,gastos_medicos:2000,
    amparo:'Incluye completo',
    auto_sust:'10 días · siniestro >$1,000',
    legal:'SÍ',exequial:'SÍ',
    vida:'N/A',enf_graves:'N/A',renta_hosp:'N/A',sepelio:'N/A',
    telemedicina:'N/A',dental:'N/A',medico_dom:'N/A',
    ded_parcial:'10%VS / 1%VA / mín.$350',
    ded_daño:'15%VA',ded_robo_sin:'30%VA',ded_robo_con:'20%VA',
  },
  LATINA:{
    color:'#6b5b95',pnMin:350,tcMax:12,debMax:10,
    tasa:va=>va<=19999?0.039:va<=29999?0.028:0.025,
    resp_civil:30000,muerte_ocupante:10000,muerte_titular:10000,gastos_medicos:2500,
    amparo:'Incluye completo',
    auto_sust:'10 días · siniestro >$1,150',
    legal:'SÍ',exequial:'SÍ',
    vida:'$20,000',enf_graves:'N/A',
    renta_hosp:'$50/día · máx.10 días',
    sepelio:'N/A',
    telemedicina:'E-DOCTOR (medicina, psicología, nutrición)',
    dental:'Prevención y cirugía 70–100%',medico_dom:'N/A',
    ded_parcial:'10%VS / 1%VA / mín.$250',
    ded_daño:'20%VA',ded_robo_sin:'20%VA',ded_robo_con:'20%VA',
  },
  GENERALI:{
    color:'#c84b1a',pnMin:0,tcMax:12,debMax:10,
    tasa:va=>va<=30000?0.0475:0.039,
    resp_civil:35000,muerte_ocupante:8000,muerte_titular:8000,gastos_medicos:3000,
    amparo:'Con costo adicional',
    auto_sust:'10 días · siniestro >$1,600+IVA',
    legal:'SÍ',exequial:'SÍ',
    vida:'N/A',enf_graves:'N/A',renta_hosp:'N/A',sepelio:'N/A',
    telemedicina:'N/A',dental:'N/A',medico_dom:'N/A',
    ded_parcial:'Espec: 15%VS/2%VA/mín.$500 · Otros: 10%VS/1%VA/mín.$250',
    ded_daño:'25%VA',ded_robo_sin:'25%VA',ded_robo_con:'25%VA',
  },
  'ASEG. DEL SUR':{
    color:'#e63946',pnMin:350,tcMax:12,debMax:8,
    tasa:va=>va<=20000?0.034:va<=30000?0.027:0.024,
    resp_civil:30000,muerte_ocupante:10000,muerte_titular:10000,gastos_medicos:3000,
    amparo:'Incluye completo',
    auto_sust:'15 días · siniestro >$1,500',
    legal:'SÍ',exequial:'NO',
    vida:'N/A',enf_graves:'N/A',renta_hosp:'N/A',sepelio:'SÍ',
    telemedicina:'N/A',dental:'N/A',medico_dom:'N/A',
    ded_parcial:'10%VS / 1%VA / mín.$250',
    ded_daño:'15%VA',ded_robo_sin:'20%VA',ded_robo_con:'20%VA',
  },
  SWEADEN:{
    color:'#1a4c84',pnMin:350,tcMax:9,debMax:10,
    tasa:va=>va<=20000?0.035:va<=30000?0.028:0.025,
    resp_civil:30000,muerte_ocupante:5000,muerte_titular:10000,gastos_medicos:2000,
    amparo:'Incluye completo',
    auto_sust:'30 días · siniestro >$1,200',
    legal:'SÍ',exequial:'NO',
    vida:'$5,000',enf_graves:'$2,500',
    renta_hosp:'$25/día · máx.30 días',
    sepelio:'$100',
    telemedicina:'6 consultas',
    dental:'N/A',medico_dom:'Copago $10/evento',
    ded_parcial:'10%VS / 1%VA / mín.$250',
    ded_daño:'20%VA',ded_robo_sin:'20%VA',ded_robo_con:'20%VA',
  },
  MAPFRE:{
    color:'#b8860b',pnMin:400,tcMax:9,debMax:10,
    tasa:va=>va<=20000?0.035:va<=30000?0.028:0.025,
    resp_civil:30000,muerte_ocupante:6000,muerte_titular:null,gastos_medicos:2000,
    amparo:'Incluye completo',
    auto_sust:'10 días · siniestro >$1,250',
    legal:'SÍ',exequial:'NO',
    vida:'$5,000',enf_graves:'$2,500',
    renta_hosp:'$20/día accidente · máx.30 días',
    sepelio:'$100',
    telemedicina:'SÍ',
    dental:'N/A',medico_dom:'SÍ',
    ded_parcial:'10%VS / 1.5%VA / mín.$350',
    ded_daño:'15%VA',ded_robo_sin:'30%VA',ded_robo_con:'15%VA',
  },
  ALIANZA:{
    color:'#2d6a4f',pnMin:350,tcMax:9,debMax:10,
    tasa:va=>va<=20000?0.030:va<=30000?0.025:0.022,
    resp_civil:30000,muerte_ocupante:5000,muerte_titular:null,gastos_medicos:2100,
    amparo:'Incluye completo',
    auto_sust:'Sedan/SP: >$500 · SUV≤$40k: >$40,001 · SUV>$60k: >$60,001',
    legal:'SÍ',exequial:'SÍ',
    vida:'$5,000',enf_graves:'$2,500',
    renta_hosp:'$20/día · máx.25 días · hasta $500',
    sepelio:'$100',
    telemedicina:'SÍ',
    dental:'N/A',medico_dom:'Copago $10 · Titular y Cónyuge',
    ded_parcial:'10%VS / 1%VA / mín.$250 (>40,000 km: 1.5%VA)',
    ded_daño:'15%VA',ded_robo_sin:'15%VA',ded_robo_con:'15%VA',
  },
  EQUINOCCIAL:{
    color:'#0077b6',pnMin:350,tcMax:12,debMax:10,
    tasa:va=>va<=20000?0.034:va<=30000?0.027:0.024,
    resp_civil:30000,muerte_ocupante:8000,muerte_titular:8000,gastos_medicos:2500,
    amparo:'Incluye completo',
    auto_sust:'10 días · siniestro >$1,000',
    legal:'SÍ',exequial:'SÍ',
    vida:'N/A',enf_graves:'N/A',renta_hosp:'N/A',sepelio:'N/A',
    telemedicina:'N/A',dental:'N/A',medico_dom:'N/A',
    ded_parcial:'10%VS / 1%VA / mín.$250',
    ded_daño:'20%VA',ded_robo_sin:'20%VA',ded_robo_con:'20%VA',
  },
  ATLANTIDA:{
    color:'#457b9d',pnMin:350,tcMax:12,debMax:10,
    tasa:va=>va<=20000?0.036:va<=30000?0.029:0.025,
    resp_civil:30000,muerte_ocupante:5000,muerte_titular:5000,gastos_medicos:2000,
    amparo:'Incluye completo',
    auto_sust:'10 días · siniestro >$1,200',
    legal:'SÍ',exequial:'NO',
    vida:'N/A',enf_graves:'N/A',renta_hosp:'N/A',sepelio:'N/A',
    telemedicina:'N/A',dental:'N/A',medico_dom:'N/A',
    ded_parcial:'10%VS / 1%VA / mín.$250',
    ded_daño:'20%VA',ded_robo_sin:'20%VA',ded_robo_con:'20%VA',
  },
  AIG:{
    color:'#2a9d8f',pnMin:400,tcMax:12,debMax:10,
    tasa:va=>va<=20000?0.038:va<=30000?0.030:0.026,
    resp_civil:40000,muerte_ocupante:10000,muerte_titular:10000,gastos_medicos:5000,
    amparo:'Incluye completo',
    auto_sust:'15 días · siniestro >$800',
    legal:'SÍ',exequial:'SÍ',
    vida:'N/A',enf_graves:'N/A',renta_hosp:'N/A',sepelio:'N/A',
    telemedicina:'N/A',dental:'N/A',medico_dom:'N/A',
    ded_parcial:'10%VS / 1%VA / mín.$250',
    ded_daño:'15%VA',ded_robo_sin:'15%VA',ded_robo_con:'15%VA',
  },
};

// Prima mínima: si la prima calculada < pnMin se aplica pnMin directamente.
// El VA del cliente NO se ajusta — siempre se muestra el valor real asegurado.
function calcPrima(va, tasa, pnMin=0){
  const pnCalc = va * tasa;
  const pnAplicada = (pnMin > 0 && pnCalc < pnMin) ? pnMin : pnCalc;
  const aplicaMin  = pnMin > 0 && pnCalc < pnMin;
  const pn=pnAplicada, der=3, camp=pn*0.005, sb=pn*0.035;
  const sub=pn+der+camp+sb, iva=sub*0.15, total=sub+iva;
  // vaEfectivo = va siempre (no se altera el valor asegurado)
  return{pn,der,camp,sb,sub,iva,total,vaEfectivo:va,vaOriginal:va,ajustado:aplicaMin,pnCalc};
}

// Cuotas TC: respeta máximo por aseguradora, y cuota mínima $50
function calcCuotasTc(total, tcMax, nCuotas){
  // No puede superar el máximo de la aseguradora
  let n = Math.min(nCuotas, tcMax);
  // Cuota no puede ser < $50 → reducir número de cuotas hasta que cuota >= 50
  while(n > 1 && total/n < 50) n--;
  return {n, cuota: total/n};
}

// Cuotas débito: cuota mínima $50
function calcCuotasDeb(total, nCuotas){
  let n = nCuotas;
  while(n > 1 && total/n < 50) n--;
  return {n, cuota: total/n};
}

function prefillCotizador(c){
  document.getElementById('cot-nombre').value=c.nombre||'';
  document.getElementById('cot-ci').value=c.ci||'';
  document.getElementById('cot-cel').value=c.celular||'';
  document.getElementById('cot-email').value=c.correo||'';
  document.getElementById('cot-va').value=c.va||20000;
  document.getElementById('cot-marca').value=c.marca||'KIA';
  document.getElementById('cot-modelo').value=c.modelo||'';
  document.getElementById('cot-anio').value=c.anio||2024;
  document.getElementById('cot-placa').value=c.placa||'';
  // Datos vehículo adicionales
  if(document.getElementById('cot-motor')) document.getElementById('cot-motor').value=c.motor||'';
  if(document.getElementById('cot-chasis')) document.getElementById('cot-chasis').value=c.chasis||'';
  if(document.getElementById('cot-color')) document.getElementById('cot-color').value=c.color||'';
  // Póliza anterior
  if(document.getElementById('cot-poliza-anterior')) document.getElementById('cot-poliza-anterior').value=c.polizaAnterior||c.poliza||'';
  if(document.getElementById('cot-aseg-anterior')) document.getElementById('cot-aseg-anterior').value=c.aseguradoraAnterior||c.aseguradora||'';
}
function prefillCotizador_show(id){
  const c=DB.find(x=>String(x.id)===String(id)); if(!c) return;
  prefillCotizador(c); showPage('cotizador'); setTimeout(calcCotizacion,200);
}

function calcCotizacion(){
  const va=parseFloat(document.getElementById('cot-va').value)||0;
  const extras=parseFloat(document.getElementById('cot-extras').value)||0;
  const vaT=va+extras;
  if(vaT<1000){showToast('Ingrese un valor asegurado válido','error');return;}
  const cuotasTcReq=parseInt(document.getElementById('cot-cuotas-tc').value)||12;
  const cuotasDebReq=parseInt(document.getElementById('cot-cuotas-deb').value)||10;

  // ── Auto Sustituto SWEADEN ──
  const AUTOSUST_COSTO = 60;
  const autoSustActivo = document.getElementById('sweaden-autosust')?.checked || false;

  // Filtrar por aseguradoras seleccionadas
  const selectedAseg=getSelectedAseg();
  if(selectedAseg.length===0){showToast('Selecciona al menos una aseguradora','error');return;}

  const results=Object.entries(ASEGURADORAS)
    .filter(([name])=>selectedAseg.includes(name))
    .map(([name,cfg])=>{
      const tasa=cfg.tasa(vaT);
      const p=calcPrima(vaT,tasa,cfg.pnMin);
      const esSweaden = name==='SWEADEN';
      const extraAutoSust = (esSweaden && autoSustActivo) ? AUTOSUST_COSTO : 0;
      const totalConExtra = p.total + extraAutoSust;
      const tc=calcCuotasTc(totalConExtra,cfg.tcMax,cuotasTcReq);
      const debN=Math.min(cuotasDebReq,cfg.debMax||cuotasDebReq);
      const deb=calcCuotasDeb(totalConExtra,debN);
      return{name,cfg,tasa,...p,total:totalConExtra,extraAutoSust,tc,deb};
    });

  // Toggle auto sustituto solo visible si SWEADEN está seleccionada
  const wrapEl=document.getElementById('sweaden-autosust-wrap');
  if(wrapEl) wrapEl.style.display=selectedAseg.includes('SWEADEN')?'flex':'none';

  const minTotal=Math.min(...results.map(r=>r.total));

  document.getElementById('aseg-cards-result').innerHTML=results.map(r=>{
    const warnings=[];
    if(r.ajustado) warnings.push(`⚠ Prima mínima aplicada: $${r.pn.toFixed(2)} — VA $${fmt(r.vaEfectivo)} no alcanza prima mínima de ${fmt(r.cfg.pnMin)}`);
    if(r.tc.n < Math.min(cuotasTcReq,r.cfg.tcMax)) warnings.push(`⚠ TC ajustado a ${r.tc.n} cuotas (cuota mín. $50)`);
    if(r.tc.n < cuotasTcReq && r.cfg.tcMax < cuotasTcReq) warnings.push(`⚠ TC máximo permitido: ${r.cfg.tcMax} cuotas`);
    if(r.deb.n < cuotasDebReq) warnings.push(`⚠ Débito ajustado a ${r.deb.n} cuotas (cuota mín. $50)`);

    // Badge especial en tarjeta SWEADEN cuando auto sustituto está activo
    const autoSustBadge = (r.name==='SWEADEN' && r.extraAutoSust>0)
      ? `<div style="background:#e8f0fb;border:1px solid #1a4c84;border-radius:6px;padding:5px 8px;margin-bottom:8px;font-size:11px;color:#1a4c84;display:flex;justify-content:space-between;align-items:center">
           <span>🚗 Auto Sustituto incluido</span>
           <span style="font-weight:700;font-family:'DM Mono',monospace">+${fmt(r.extraAutoSust)}</span>
         </div>`
      : '';

    // Fila de desglose auto sustituto en el detalle de prima
    const autoSustRow = (r.name==='SWEADEN' && r.extraAutoSust>0)
      ? `<div class="aseg-row" style="color:#1a4c84;font-weight:500">
           <span class="aseg-key">🚗 Auto Sustituto</span>
           <span class="aseg-val">${fmt(r.extraAutoSust)}</span>
         </div>`
      : '';

    return `<div class="aseg-card${r.total===minTotal?' mejor':''}">
      <div class="aseg-name" style="color:${r.cfg.color}">${r.name}</div>
      ${warnings.length?`<div style="background:#fff8e1;border:1px solid #f0c040;border-radius:6px;padding:6px 8px;margin-bottom:8px;font-size:10px;color:#7a5c00;line-height:1.6">${warnings.join('<br>')}</div>`:''}
      ${autoSustBadge}
      <div class="aseg-row"><span class="aseg-key">Tasa</span><span class="aseg-val">${(r.tasa*100).toFixed(2)}%</span></div>
      <div class="aseg-row"><span class="aseg-key">Prima Neta</span><span class="aseg-val">${fmt(r.pn)}</span></div>
      <div class="aseg-row"><span class="aseg-key">Der. Emisión</span><span class="aseg-val">${fmt(r.der)}</span></div>
      <div class="aseg-row"><span class="aseg-key">Seg. Campesino</span><span class="aseg-val">${fmt(r.camp)}</span></div>
      <div class="aseg-row"><span class="aseg-key">Super Bancos</span><span class="aseg-val">${fmt(r.sb)}</span></div>
      <div class="aseg-row"><span class="aseg-key">IVA 15%</span><span class="aseg-val">${fmt(r.iva)}</span></div>
      ${autoSustRow}
      <div class="aseg-total"><span class="aseg-total-key">Total${r.extraAutoSust>0?' c/Auto Sust.':''}</span><span class="aseg-total-val">${fmt(r.total)}</span></div>
      <div class="aseg-cuota">💳 TC ${r.tc.n} cuotas: <b>${fmt(r.tc.cuota)}/mes</b></div>
      <div class="aseg-cuota">🏦 Débito ${r.deb.n} cuotas: <b>${fmt(r.deb.cuota)}/mes</b></div>
      <div style="margin-top:10px;display:flex;gap:4px">
        <button class="btn btn-ghost btn-xs w-full" onclick="printOneAseg('${r.name}',${r.total.toFixed(2)},${r.pn.toFixed(2)},${r.tc.cuota.toFixed(2)},${r.deb.cuota.toFixed(2)},${r.tc.n},${r.deb.n})">🖨 Imprimir / PDF</button>
      </div>
    </div>`;
  }).join('');

  // Coberturas
  document.getElementById('coberturas-table-wrap').innerHTML=`<table class="comp-table">
    <thead><tr><th class="col-cob" style="text-align:left">Cobertura</th>${results.map(r=>`<th style="color:${r.cfg.color}">${r.name}</th>`).join('')}</tr></thead>
    <tbody>
      <tr class="section-row"><td colspan="${results.length+1}">Coberturas Básicas</td></tr>
      <tr><td class="col-cob">Todo Riesgo</td>${results.map(()=>`<td class="col-val yes">SÍ</td>`).join('')}</tr>
      <tr><td class="col-cob">Pérdida parcial</td>${results.map(()=>`<td class="col-val yes">SÍ</td>`).join('')}</tr>
      <tr><td class="col-cob">Pérdida total</td>${results.map(()=>`<td class="col-val yes">SÍ</td>`).join('')}</tr>
      <tr class="section-row"><td colspan="${results.length+1}">Amparos Adicionales</td></tr>
      <tr><td class="col-cob">Responsabilidad Civil</td>${results.map(r=>`<td class="col-val">${fmt(r.cfg.resp_civil)}</td>`).join('')}</tr>
      <tr><td class="col-cob">Muerte acc./ocupante</td>${results.map(r=>`<td class="col-val">${fmt(r.cfg.muerte_ocupante)}</td>`).join('')}</tr>
      <tr><td class="col-cob">Muerte acc./titular</td>${results.map(r=>`<td class="col-val">${r.cfg.muerte_titular?fmt(r.cfg.muerte_titular):'<span class="no">N/A</span>'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Gastos Médicos</td>${results.map(r=>`<td class="col-val">${fmt(r.cfg.gastos_medicos)}</td>`).join('')}</tr>
      <tr><td class="col-cob">Amparo Patrimonial</td>${results.map(r=>`<td class="col-val" style="font-size:10px">${r.cfg.amparo}</td>`).join('')}</tr>
      <tr class="section-row"><td colspan="${results.length+1}">Beneficios</td></tr>
      <tr><td class="col-cob">Auto sustituto</td>${results.map(r=>{
        const incluyeExtra=r.name==='SWEADEN'&&autoSustActivo;
        return `<td class="col-val" style="font-size:10px${incluyeExtra?';color:#1a4c84;font-weight:600':''}">
          ${r.cfg.auto_sust}
          ${incluyeExtra?'<br><span style="background:#1a4c84;color:#fff;font-size:9px;border-radius:3px;padding:1px 5px">✓ Sin mínimo de siniestro — incluido $60</span>':''}
        </td>`;
      }).join('')}</tr>
      <tr><td class="col-cob">Asist. Legal en situ</td>${results.map(r=>`<td class="col-val ${(r.cfg.legal||'SÍ')==='SÍ'?'yes':'no'}">${r.cfg.legal||'SÍ'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Asistencia Exequial</td>${results.map(r=>`<td class="col-val ${r.cfg.exequial==='SÍ'?'yes':'no'}">${r.cfg.exequial}</td>`).join('')}</tr>
      <tr><td class="col-cob">Vida/Muerte Accidental</td>${results.map(r=>`<td class="col-val" style="font-size:10px">${r.cfg.vida||'N/A'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Enfermedades Graves</td>${results.map(r=>`<td class="col-val" style="font-size:10px">${r.cfg.enf_graves||'N/A'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Renta Hospitalización</td>${results.map(r=>`<td class="col-val" style="font-size:10px">${r.cfg.renta_hosp||'N/A'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Gastos de Sepelio</td>${results.map(r=>`<td class="col-val" style="font-size:10px">${r.cfg.sepelio||'N/A'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Telemedicina</td>${results.map(r=>`<td class="col-val" style="font-size:10px">${r.cfg.telemedicina||'N/A'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Beneficio Dental</td>${results.map(r=>`<td class="col-val" style="font-size:10px">${r.cfg.dental||'N/A'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Médico a Domicilio</td>${results.map(r=>`<td class="col-val" style="font-size:10px">${r.cfg.medico_dom||'N/A'}</td>`).join('')}</tr>
      <tr class="section-row"><td colspan="${results.length+1}">Prima Mínima / Cuotas Máx.</td></tr>
      <tr><td class="col-cob">Prima Neta Mínima</td>${results.map(r=>`<td class="col-val mono">${r.cfg.pnMin>0?fmt(r.cfg.pnMin):'—'}</td>`).join('')}</tr>
      <tr><td class="col-cob">TC Máx. cuotas</td>${results.map(r=>`<td class="col-val mono">${r.cfg.tcMax} cuotas</td>`).join('')}</tr>
      <tr><td class="col-cob">Débito Máx. cuotas</td>${results.map(r=>`<td class="col-val mono">${r.cfg.debMax||10} cuotas</td>`).join('')}</tr>
      <tr class="section-row"><td colspan="${results.length+1}">Deducibles</td></tr>
      <tr><td class="col-cob">Pérdida parcial</td>${results.map(r=>`<td class="col-val" style="font-size:10px">${r.cfg.ded_parcial}</td>`).join('')}</tr>
      <tr><td class="col-cob">Pérd. total daños</td>${results.map(r=>`<td class="col-val">${r.cfg.ded_daño}</td>`).join('')}</tr>
      <tr><td class="col-cob">Pérd. total robo s/d</td>${results.map(r=>`<td class="col-val">${r.cfg.ded_robo_sin}</td>`).join('')}</tr>
      <tr><td class="col-cob">Pérd. total robo c/d</td>${results.map(r=>`<td class="col-val">${r.cfg.ded_robo_con}</td>`).join('')}</tr>
    </tbody></table>`;
  document.getElementById('cotizacion-resultado').style.display='block';
}

// ── CIERRE DE VENTA ──────────────────────────────────
let cierreVentaData={};
function abrirCierreDesdeCliente(id, skipEstadoCheck=false){
  const c=DB.find(x=>String(x.id)===String(id)); if(!c) return;
  if(!skipEstadoCheck && !['RENOVADO','EMITIDO'].includes(c.estado)){showToast('El estado debe ser RENOVADO o EMITIDO para registrar un cierre','error');return;}
  currentSegIdx=id;
  const aseg=c.aseguradora||'';
  const va=c.va||0;
  const cfg=Object.entries(ASEGURADORAS).find(([k])=>aseg.toUpperCase().includes(k));
  let total=0,pn=0,tcN=12,debN=10,tcCuota=0,debCuota=0;
  if(cfg&&va>0){
    const [,cfgObj]=cfg;
    const tasa=cfgObj.tasa(va);
    const p=calcPrima(va,tasa,cfgObj.pnMin);
    const tc=calcCuotasTc(p.total,cfgObj.tcMax,12);
    const deb=calcCuotasDeb(p.total,10);
    total=p.total; pn=p.pn; tcN=tc.n; debN=deb.n; tcCuota=tc.cuota; debCuota=deb.cuota;
  }
  cierreVentaData={asegNombre:aseg,total,pn,cuotaTc:tcCuota,cuotaDeb:debCuota,nTc:tcN,nDeb:debN,clienteId:id};
  document.getElementById('cv-aseg').textContent=aseg;
  document.getElementById('cv-total').textContent=total>0?`${fmt(total)} total`:'Ingrese prima manualmente';
  document.getElementById('cv-cliente').value=c.nombre;
  document.getElementById('cv-nueva-aseg').value=aseg;
  const hoyPlus1=new Date(); hoyPlus1.setFullYear(hoyPlus1.getFullYear()+1);
  document.getElementById('cv-desde').value=c.hasta||new Date().toISOString().split('T')[0];
  document.getElementById('cv-hasta').value=hoyPlus1.toISOString().split('T')[0];
  document.getElementById('cv-pn').value=pn>0?pn.toFixed(2):'';
  document.getElementById('cv-total-val').value=total>0?total.toFixed(2):'';
  if(document.getElementById('cv-cuenta')) document.getElementById('cv-cuenta').value=c.cuenta||'';
  if(document.getElementById('cv-axavd')) document.getElementById('cv-axavd').value=c.obs&&c.obs.includes('AXA')&&c.obs.includes('VD')?'AXA+VD':c.obs&&c.obs.includes('AXA')?'AXA':c.obs&&c.obs.includes('VD')?'VD':'';
  ['cv-factura','cv-poliza','cv-observacion'].forEach(id=>{ const el=document.getElementById(id); if(el) el.value=''; });
  document.getElementById('cv-forma-pago').value='DEBITO_BANCARIO';
  renderCvFormaPago();
  openModal('modal-cierre-venta');
}

function abrirCierreVenta(asegNombre, total, pn, cuotaTc, cuotaDeb, nTc, nDeb){
  cierreVentaData={asegNombre,total,pn,cuotaTc,cuotaDeb,nTc,nDeb};
  const clienteNombre=document.getElementById('cot-nombre').value||'';
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
  // Buscar cuenta del cliente por nombre
  const cMatch=DB.find(x=>x.nombre.trim().toUpperCase()===clienteNombre.trim().toUpperCase());
  if(cMatch&&document.getElementById('cv-cuenta')) document.getElementById('cv-cuenta').value=cMatch.cuenta||'';
  if(document.getElementById('cv-axavd')) document.getElementById('cv-axavd').value='';
  ['cv-factura','cv-poliza','cv-fecha-cobro-inicial','cv-observacion'].forEach(id=>{ const el=document.getElementById(id); if(el) el.value=''; });
  document.getElementById('cv-forma-pago').value='DEBITO_BANCARIO';
  renderCvFormaPago();
  openModal('modal-cierre-venta');
}
function getTotal(){ return parseFloat(document.getElementById('cv-total-val')?.value)||cierreVentaData.total||0; }
function recalcPrimaTotal(){
  // cuando editan prima neta, recalcular total (pn + derechos + campesino + sb + iva)
  const pn=parseFloat(document.getElementById('cv-pn')?.value)||0;
  const der=3, camp=pn*0.005, sb=pn*0.035, sub=pn+der+camp+sb, iva=sub*0.15, total=sub+iva;
  const el=document.getElementById('cv-total-val'); if(el) el.value=total.toFixed(2);
  cierreVentaData.total=total; cierreVentaData.pn=pn;
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
  const banco=document.getElementById('cv-banco-deb')?.value||'—';
  const cuenta=document.getElementById('cv-cuenta-deb')?.value||'—';
  // Validar cuota mínima $50
  const aviso=cuota<50?`<div style="margin-bottom:8px;padding:6px 10px;background:#fde8e0;border-radius:6px;font-size:11px;color:var(--accent)">⚠ Cuota menor a $50 — considere reducir el número de cuotas</div>`:'';
  let html=`${aviso}<div class="section-divider">Calendario de débitos — ${banco} · Cta: ${cuenta}</div>
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
  const clienteNombre=document.getElementById('cv-cliente').value.trim();
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
  // Validar campos de forma de pago
  if(fp==='CONTADO'){
    if(!document.getElementById('cv-fecha-cobro-total')?.value) errors.push('Fecha de cobro de contado');
  } else if(fp==='DEBITO_BANCARIO'){
    if(!document.getElementById('cv-n-cuotas')?.value) errors.push('N° de cuotas');
    if(!document.getElementById('cv-fecha-primera')?.value) errors.push('Fecha 1ª cuota');
  } else if(fp==='TARJETA_CREDITO'){
    if(!document.getElementById('cv-n-cuotas-tc')?.value) errors.push('N° cuotas TC');
    if(!document.getElementById('cv-fecha-contacto-tc')?.value) errors.push('Fecha contacto TC');
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
  const c=cierreVentaData.clienteId ? DB.find(x=>String(x.id)===String(cierreVentaData.clienteId)) : DB.find(x=>x.nombre.trim().toUpperCase()===clienteNombre.trim().toUpperCase());

  // Armar registro de cierre
  // _clienteId y _placa usan prefijo _ para ser solo locales (no se envían a SharePoint)
  const cierre={
    id: Date.now(),
    fechaRegistro: new Date().toISOString().split('T')[0],
    clienteNombre, aseguradora:aseg, polizaNueva:poliza,
    facturaAseg:factura, primaTotal:total,
    primaNeta:parseFloat(document.getElementById('cv-pn')?.value)||cierreVentaData.pn||0,
    vigDesde:desde, vigHasta:hasta,
    formaPago:pago, observacion:obs,
    axavd, cuenta:document.getElementById('cv-cuenta')?.value||'',
    ejecutivo:currentUser?currentUser.id:'',
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

    // 2. Duplicado por cliente + placa (mismo vehículo)
    if(cierre._clienteId && cierre._placa){
      const dupVehiculo = allCierres.find(x =>
        x._clienteId && String(x._clienteId) === cierre._clienteId &&
        x._placa && x._placa.trim().toUpperCase() === cierre._placa.trim().toUpperCase()
      );
      if(dupVehiculo){
        showToast(`Ya existe un cierre para ${clienteNombre} con placa ${cierre._placa} (${dupVehiculo.fechaRegistro}). Si es un vehículo distinto, actualiza la placa del cliente primero.`, 'error');
        return;
      }
    }

    cierre._dirty = true;
    allCierres.push(cierre);
    _saveCierres(allCierres);
    showToast(`✓ Venta cerrada — ${aseg} — Póliza ${poliza}`,'success');
  }

  // Actualizar cliente en DB
  if(c){
    c.polizaNueva=poliza; c.factura=factura; c.aseguradora=aseg;
    c.desde=desde; c.hasta=hasta; c.formaPago=pago;
    c.primaTotal=total; c.axavd=axavd;
    c.estado='RENOVADO'; _bitacoraAdd(c, `Cierre registrado${obs?' — '+obs:''}. Aseg: ${cierre?.aseguradora||''}`, 'cierre');
    c.ultimoContacto=new Date().toISOString().split('T')[0];
    saveDB();
    sincronizarCotizPorCliente(c.id, c.nombre, c.ci, 'RENOVADO');
  }

  // Reset modo edición y botón
  cierreVentaData.editandoCierreId=null;
  const btnG=document.querySelector('#modal-cierre-venta .btn-green[onclick="guardarCierreVenta()"]');
  if(btnG){ btnG.textContent='✓ Registrar Cierre'; btnG.style.background=''; }

  closeModal('modal-cierre-venta');
  renderDashboard();
  renderCierres();
  actualizarBadgeCotizaciones();
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
  const autoSustActivo = document.getElementById('sweaden-autosust')?.checked||false;
  const exec     = currentUser ? currentUser.name : '—';
  const fecha    = new Date().toLocaleDateString('es-EC',{day:'2-digit',month:'long',year:'numeric'});

  // Calcular resultados solo de las aseguradoras seleccionadas
  const selected = getSelectedAseg();
  if(!selected.length){showToast('Selecciona al menos una aseguradora','error');return;}

  const results = selected.map(name=>{
    const cfg = ASEGURADORAS[name];
    const tasa = cfg.tasa(vaT);
    const p = calcPrima(vaT,tasa,cfg.pnMin);
    const extraAutoSust = (name==='SWEADEN'&&autoSustActivo) ? 60 : 0;
    const total = p.total + extraAutoSust;
    const tc  = calcCuotasTc(total, cfg.tcMax, cuotasTcReq);
    const deb = calcCuotasDeb(total, cuotasDebReq);
    return {name,cfg,tasa,...p,total,extraAutoSust,tc,deb};
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
      ${results.some(r=>r.extraAutoSust>0)?`
      <tr>
        <td class="col-cob">🚗 Auto Sustituto</td>
        ${results.map(r=>`<td class="sweaden-extra">${r.extraAutoSust>0?'$'+r.extraAutoSust.toFixed(2):'—'}</td>`).join('')}
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
      <tr><td class="col-cob">Vida/Muerte Acc.</td>${results.map(r=>`<td style="font-size:9px">${r.cfg.vida||'N/A'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Enf. Graves</td>${results.map(r=>`<td style="font-size:9px">${r.cfg.enf_graves||'N/A'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Renta Hosp.</td>${results.map(r=>`<td style="font-size:9px">${r.cfg.renta_hosp||'N/A'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Sepelio</td>${results.map(r=>`<td style="font-size:9px">${r.cfg.sepelio||'N/A'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Telemedicina</td>${results.map(r=>`<td style="font-size:9px">${r.cfg.telemedicina||'N/A'}</td>`).join('')}</tr>
      <tr><td class="col-cob">Médico Domicilio</td>${results.map(r=>`<td style="font-size:9px">${r.cfg.medico_dom||'N/A'}</td>`).join('')}</tr>

      <tr class="section"><td class="col-cob" colspan="${results.length+1}">Deducibles</td></tr>
      <tr><td class="col-cob">Pérd. Parcial</td>${results.map(r=>`<td style="font-size:9px">${r.cfg.ded_parcial}</td>`).join('')}</tr>
      <tr><td class="col-cob">Pérd. Total Daños</td>${results.map(r=>`<td>${r.cfg.ded_daño}</td>`).join('')}</tr>
      <tr><td class="col-cob">Pérd. Total Robo</td>${results.map(r=>`<td>${r.cfg.ded_robo_sin}</td>`).join('')}</tr>
    </tbody>
  </table>

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
  const autoSustActivo = document.getElementById('sweaden-autosust')?.checked||false;

  if(!nombre){ showToast('Ingrese el nombre del cliente','error'); return; }
  if(vaT<1000){ showToast('Ingrese un valor asegurado válido','error'); return; }

  const selected = getSelectedAseg();
  if(!selected.length){ showToast('Selecciona al menos una aseguradora','error'); return; }

  // Calcular resultados de las aseguradoras seleccionadas
  const resultados = selected.map(name=>{
    const cfg = ASEGURADORAS[name];
    const tasa = cfg.tasa(vaT);
    const p = calcPrima(vaT, tasa, cfg.pnMin);
    const extraAutoSust = (name==='SWEADEN' && autoSustActivo) ? 60 : 0;
    const total = p.total + extraAutoSust;
    const debN = Math.min(cuotasDebReq, cfg.debMax||cuotasDebReq);
    const tc  = calcCuotasTc(total, cfg.tcMax, cuotasTcReq);
    const deb = calcCuotasDeb(total, debN);
    return { name, tasa, pn:p.pn, total, extraAutoSust,
             tcN:tc.n, tcCuota:tc.cuota, debN:deb.n, debCuota:deb.cuota };
  });

  // Buscar cliente en DB por CI o nombre
  const clienteMatch = DB.find(c=>
    (ci && c.ci===ci) || c.nombre.trim().toUpperCase()===nombre.toUpperCase()
  );

  const allCotiz = _getCotizaciones();
  const esEdicion = !!window._editandoCotizId;
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
    clienteId: clienteMatch?.id||null,
    celular, correo, ciudad, region,
    // Datos vehículo
    vehiculo: `${marca} ${modelo} ${anio}`.trim(),
    marca, modelo, anio, placa, color, motor, chasis,
    // Datos póliza
    tipo, va: vaT, desde, hasta,
    asegAnterior, polizaAnterior,
    cuotasTc: cuotasTcReq, cuotasDeb: cuotasDebReq,
    autoSust: autoSustActivo,
    aseguradoras: selected, resultados,
    estado: 'ENVIADA', asegElegida: null, obsAcept: '', fechaAcept: null,
  };

  cotiz._dirty = true;
  allCotiz.push(cotiz);
  _saveCotizaciones(allCotiz);
  actualizarBadgeCotizaciones();
  renderCotizaciones();
  showToast(`✅ ${codigo} guardada — ${nombre.split(' ')[0]} · ${selected.length} aseguradoras`, 'success');
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
    const mEl=document.getElementById('cot-marca');   if(mEl) mEl.value=marcaV;
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
  all[idx].estado = 'REEMPLAZADA';
  all[idx]._dirty = true;
  _saveCotizaciones(all);
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
    const mEl=document.getElementById('cot-marca');   if(mEl) mEl.value=marcaV;
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
      acciones = `
        <button class="btn btn-green btn-xs"
          onclick="cambiarEstadoCotiz('${c.id}','EMITIDA')" title="Marcar como emitida">🛡 Emitida</button>`;
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

    // Buscar cotizaciones EN EMISIÓN vinculadas a este cliente
    all.forEach(cot=>{
      if(cot.estado !== 'EN EMISIÓN') return;
      const vinculada =
        (clienteId  && String(cot.clienteId) === String(clienteId)) ||
        (clienteCI  && clienteCI.length>3 && cot.clienteCI === clienteCI) ||
        (clienteNombre && cot.clienteNombre.trim().toUpperCase() === clienteNombre.trim().toUpperCase());
      if(!vinculada) return;

      // Mapeo estado cliente → estado cotización
      if(['RENOVADO','EMITIDO','PÓLIZA VIGENTE'].includes(nuevoEstadoCliente)){
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
  cotizAsegSeleccionada = null;

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
  cotizAsegSeleccionada = name;
  const cfg = ASEGURADORAS[name]||{color:'#333'};

  // Resaltar botón seleccionado
  document.querySelectorAll('.cotiz-aseg-btn').forEach(btn=>{
    const sel = btn.dataset.name===name;
    btn.style.border=`2px solid ${sel?cfg.color:cfg.color+'44'}`;
    btn.style.background = sel ? cfg.color+'15' : 'var(--paper)';
    btn.style.transform = sel ? 'scale(1.02)' : '';
  });

  // Resumen
  const wrap = document.getElementById('cotiz-resumen-elegida');
  wrap.style.display='block';
  wrap.style.borderColor=cfg.color;
  document.getElementById('cotiz-resumen-body').innerHTML=`
    <div style="display:grid;grid-template-columns:1fr 1fr;gap:8px;font-size:12px">
      <div><span style="color:var(--muted)">Aseguradora:</span> <b style="color:${cfg.color}">${name}</b></div>
      <div><span style="color:var(--muted)">Prima Neta:</span> <b>$${resultado.pn?.toFixed(2)}</b></div>
      <div><span style="color:var(--muted)">Total:</span> <b style="font-size:15px">$${resultado.total?.toFixed(2)}</b></div>
      <div><span style="color:var(--muted)">TC ${resultado.tcN} cuotas:</span> <b>$${resultado.tcCuota?.toFixed(2)}/mes</b></div>
      <div><span style="color:var(--muted)">Débito ${resultado.debN} cuotas:</span> <b>$${resultado.debCuota?.toFixed(2)}/mes</b></div>
    </div>`;

  document.getElementById('btn-confirmar-acept').disabled=false;
}

function confirmarAceptacion(){
  if(!cotizAceptarId || !cotizAsegSeleccionada){
    showToast('Selecciona una aseguradora primero','error'); return;
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
  all[idx].asegElegida = cotizAsegSeleccionada;
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
    clienteDB.estado         = 'EMISIÓN';
    clienteDB.aseguradora    = cotizAsegSeleccionada.includes('SEGUROS') ? cotizAsegSeleccionada : cotizAsegSeleccionada + ' SEGUROS';
    clienteDB.ultimoContacto = hoy;
    const notaAcept = `Cotización aceptada — ${cotizAsegSeleccionada}${obs?' · '+obs:''}`;
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
  const pnEl    = document.getElementById('cv-pn');
  const desdeEl = document.getElementById('cv-desde');
  if(pnEl    && r.pn)        pnEl.value    = r.pn.toFixed(2);
  if(desdeEl && cotiz.desde) desdeEl.value = cotiz.desde;
  recalcPrimaTotal();
  autocalcHasta();
  renderCvFormaPago();

  openModal('modal-cierre-venta');
  showToast('📝 Cierre pre-llenado · ' + cotiz.asegElegida, 'info');
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
}
function showAdminTab(tab, el){
  ['importar','historial','ejecutivos','datos'].forEach(t=>{
    const el2=document.getElementById('admin-tab-'+t);
    if(el2) el2.style.display=t===tab?'':'none';
  });
  document.querySelectorAll('#admin-tabs .pill').forEach(p=>p.classList.remove('active'));
  if(el) el.classList.add('active');
  if(tab==='datos') renderAdmin(); // Refresh stats
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
  const data=DB.map(c=>({
    'Ejecutivo':c.ejecutivo,'Nombre':c.nombre,'CI':c.ci,'Celular':c.celular,
    'Aseguradora':c.aseguradora,'Póliza':c.poliza,'Vigencia Desde':c.desde,'Vigencia Hasta':c.hasta,
    'VA':c.va,'Prima Neta':c.pn,'Estado':c.estado,'Marca':c.marca,'Modelo':c.modelo,
    'Año':c.anio,'Placa':c.placa,'Motor':c.motor,'Chasis':c.chasis,'Color':c.color,
    'Correo':c.correo,'Cuenta':c.cuenta,'OBS':c.obs,'Último Contacto':c.ultimoContacto
  }));
  const ws=XLSX.utils.json_to_sheet(data);
  const wb=XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb,ws,'Clientes');
  XLSX.writeFile(wb,`Reliance_Clientes_${new Date().toISOString().split('T')[0]}.xlsx`);
}
function exportarBackupJSON(){
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
    const json=XLSX.utils.sheet_to_json(ws,{defval:''});
    if(!json.length){showToast('No se encontraron datos en el archivo','error');return;}
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
  const total=rows.length;
  // Calcular nuevos vs existentes
  const existingCIs=new Set(DB.map(c=>(c.ci||'').toString().trim()));
  const nuevos=rows.filter(r=>{
    const ci=(r['CI']||r['CEDULA']||r['ci']||'').toString().trim();
    return ci&&!existingCIs.has(ci);
  }).length;
  const duplicados=total-nuevos;

  document.getElementById('import-preview-meta').textContent=`${total} registros encontrados en hoja "${sheetName}"`;
  
  // Resumen mapeo
  const mapeoEl=document.getElementById('import-mapeo-resumen');
  if(mapeoEl) mapeoEl.innerHTML=`
    <div style="display:flex;gap:20px;flex-wrap:wrap">
      <div>📋 <b>${total}</b> filas con datos</div>
      <div style="color:var(--green)">🆕 <b>${nuevos}</b> clientes nuevos</div>
      <div style="color:var(--muted)">🔁 <b>${duplicados}</b> ya existen (por CI)</div>
      <div>📅 Mes: <b>${document.getElementById('import-mes')?.value||'—'}</b></div>
    </div>`;

  // Advertencia modo
  const modo=document.getElementById('import-modo')?.value||'agregar';
  const warnEl=document.getElementById('import-warning');
  if(warnEl){
    if(modo==='reemplazar') warnEl.innerHTML=`<span style="color:var(--red)">⚠ Modo REEMPLAZAR: se eliminarán todos los clientes del ejecutivo seleccionado</span>`;
    else if(modo==='actualizar') warnEl.innerHTML=`<span style="color:var(--accent)">ℹ Modo ACTUALIZAR: se sobreescribirán datos de clientes existentes</span>`;
    else warnEl.innerHTML='';
  }

  // Tabla preview — columnas clave
  const previewCols=['Nombre Cliente','CI','ASEGURADORA','Fc_hasta ultima vigencia','Ultimo Val_Aseg.','Marca','Modelo','Placa'];
  const keys=previewCols.filter(k=>rows[0]&&k in rows[0]);
  const html=`<table><thead><tr>${keys.map(k=>`<th style="white-space:nowrap">${k}</th>`).join('')}<th>Estado</th></tr></thead>
    <tbody>${rows.slice(0,15).map(r=>{
      const ci=(r['CI']||r['CEDULA']||'').toString().trim();
      const esNuevo=ci&&!existingCIs.has(ci);
      return `<tr style="${esNuevo?'':'background:#f0f7f0'}">
        ${keys.map(k=>`<td style="font-size:11px;max-width:130px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap">${r[k]||''}</td>`).join('')}
        <td>${esNuevo?'<span style="font-size:10px;color:var(--accent)">Nuevo</span>':'<span style="font-size:10px;color:var(--muted)">Existe</span>'}</td>
      </tr>`;
    }).join('')}${rows.length>15?`<tr><td colspan="${keys.length+1}" style="text-align:center;padding:8px;color:var(--muted);font-size:11px">... y ${rows.length-15} registros más</td></tr>`:''}</tbody>
  </table>`;
  document.getElementById('import-table-wrap').innerHTML=html;
  document.getElementById('import-preview').style.display='block';
  showToast(`${total} registros detectados — ${nuevos} nuevos`,'info');
}
function confirmImport(){
  if(!importedRows.length){return;}
  const execId=document.getElementById('import-assign-exec').value;
  const modo=document.getElementById('import-modo')?.value||'agregar';
  const mes=document.getElementById('import-mes')?.value||'';
  // Modo reemplazar: eliminar cartera existente del ejecutivo
  if(modo==='reemplazar'&&execId){
    if(!confirm('¿Confirmar REEMPLAZAR toda la cartera del ejecutivo seleccionado?')) return;
    // Borrar en SP los clientes del ejecutivo antes de reemplazar
    if(_spReady){ const toDelSP=DB.filter(c=>String(c.ejecutivo)===String(execId)&&c._spId); toDelSP.forEach(c=>spDelete('clientes',c._spId)); }
    DB=DB.filter(c=>String(c.ejecutivo)!==String(execId));
  }
  const existingCIs=new Set(DB.map(c=>(c.ci||'').toString().trim()));
  const maxId=DB.length?Math.max(...DB.map(c=>c.id||0)):0;
  let added=0, updated=0;
  importedRows.forEach((row,i)=>{
    // Map common column names
    const nombre=(row['Nombre Cliente']||row['NOMBRE']||row['nombre']||'').toString().toUpperCase().trim();
    if(!nombre) return;
    const hasta=(row['Fc_hasta ultima vigencia']||row['VIGENCIA HASTA']||row['hasta']||'').toString().split(' ')[0];
    const desde=(row['Fc_desde ultima vigencia']||row['VIGENCIA DESDE']||row['desde']||'').toString().split(' ')[0];
    const ci=(row['CI']||row['CEDULA']||row['ci']||'').toString().trim();
    // Modo actualizar: si ya existe, actualizar datos
    if(modo==='actualizar'&&ci&&existingCIs.has(ci)){
      const idx=DB.findIndex(c=>(c.ci||'').toString().trim()===ci);
      if(idx>=0){
        DB[idx]={...DB[idx],aseguradora:(row['ASEGURADORA']||row['Aseguradora']||DB[idx].aseguradora||'').toString().trim(),
          va:parseFloat(row['Ultimo Val_Aseg.']||row['VA']||DB[idx].va)||DB[idx].va,
          hasta,desde};
        updated++; return;
      }
    }
    // Modo agregar: omitir si ya existe por CI
    if(modo==='agregar'&&ci&&existingCIs.has(ci)) return;
    DB.push({
      id:maxId+added+1,ejecutivo:execId,
      nombre,ci,
      tipo:(row['TIPO']||row['tipo']||'RENOVACION').toString().trim(),
      region:(row['REGION']||row['region']||'SIERRA').toString().trim(),
      ciudad:(row['Ciudad']||row['CIUDAD']||'QUITO').toString().trim(),
      obs:(row['OBS POLIZA']||row['obs']||'RENOVACION').toString().trim(),
      celular:(row['Celular 1']||row['CELULAR']||row['Celular']||'').toString().trim(),
      aseguradora:(row['ASEGURADORA']||row['aseguradora']||row['Aseguradora']||row['AE']||'').toString().trim(),
      // Col AC = POLIZA (póliza vigente anterior), Col AE = ASEGURADORA (aseg anterior)
      polizaAnterior:(row['POLIZA']||'').toString().trim(),
      aseguradoraAnterior:(row['ASEGURADORA']||row['Aseguradora']||'').toString().trim(),
      poliza:(row['POLIZA RENOVADA']||row['poliza_nueva']||'').toString().trim(),
      desde,hasta,
      va:parseFloat(row['Ultimo Val_Aseg.']||row['Valor asegurado']||row['VA']||0)||0,
      dep:parseFloat(row['v dep']||row['dep']||0)||0,
      tasa:parseFloat(row['tasa presupuesto']||row['Tasa']||row['tasa']||0)||null,
      pn:parseFloat(row['pn']||row['Prima Neta ']||row['Prima Neta']||0)||0,
      marca:(row['Marca']||row['marca']||row['AO']||'').toString().trim(),
      modelo:(row['Modelo']||row['modelo']||row['AP']||'').toString().trim(),
      anio:parseInt(row['Año ']||row['Año']||row['AQ']||row['anio']||2024)||2024,
      motor:(row['Motor']||row['AR']||'').toString().trim(),
      chasis:(row['No De Chasis']||row['AS']||row['chasis']||'').toString().trim(),
      color:(row['Color']||row['AT']||row['color']||'').toString().trim(),
      placa:(row['Placa']||row['AU']||row['placa']||'').toString().trim(),
      correo:(row['Correo']||row['Ac']||row['correo']||'').toString().trim(),
      cuenta:(row['Cuenta']||row['AV']||'').toString().trim(),
      estado:'PENDIENTE',nota:'',ultimoContacto:'',comentario:'',
      fechaDebito:(row['FECHA DEBITO']||row['BT']||'').toString().split(' ')[0],
      fechaCash:(row['FECHA DE CASH']||row['BM']||'').toString().split(' ')[0],
      bitacora:[],
    });
    // Registrar en bitácora del cliente recién creado
    const nuevoImport = DB[DB.length-1];
    if(nuevoImport) _bitacoraAdd(nuevoImport, `Cliente importado desde Excel — ${importedFileName||'archivo.xlsx'}`, 'sistema');
    added++;
  });
  saveDB();
  // Registrar en historial
  const execUser=USERS.find(u=>u.id===execId);
  const historial=(_cache.historial||JSON.parse(localStorage.getItem('reliance_import_historial')||'[]'));
  historial.push({
    fecha:new Date().toISOString().split('T')[0],
    archivo:importedFileName||'archivo.xlsx',
    mes, ejecutivo:execUser?execUser.name:'Todos',
    modo, agregados:added, actualizados:updated
  });
  { _cache.historial=historial; localStorage.setItem('reliance_import_historial',JSON.stringify(historial)); }
  document.getElementById('import-preview').style.display='none';
  importedRows=[];
  renderDashboard(); renderAdmin();
  showToast(`✓ ${added} clientes importados${updated?' · '+updated+' actualizados':''}${modo==='reemplazar'?' (cartera reemplazada)':''}`);
}
function cancelImport(){document.getElementById('import-preview').style.display='none';importedRows=[];}

// ══════════════════════════════════════════════════════
//  NUEVO CLIENTE
// ══════════════════════════════════════════════════════
function guardarNuevoCliente(){
  const nombre=document.getElementById('nc-nombre').value.trim();
  const ci=document.getElementById('nc-ci').value.trim();
  if(!nombre||!ci){showToast('Nombre y CI son requeridos','error');return;}
  const maxId=DB.length?Math.max(...DB.map(c=>c.id||0)):0;
  DB.push({
    id:maxId+1,ejecutivo:currentUser.id,
    nombre:nombre.toUpperCase(),ci,tipo:document.getElementById('nc-tipo').value,
    region:document.getElementById('nc-region').value,ciudad:document.getElementById('nc-ciudad').value,
    obs:document.getElementById('nc-obs').value,celular:document.getElementById('nc-cel').value,
    correo:document.getElementById('nc-email').value,aseguradora:document.getElementById('nc-aseg').value,
    va:parseFloat(document.getElementById('nc-va').value)||0,
    tasa:parseFloat(document.getElementById('nc-tasa').value)||null,
    marca:document.getElementById('nc-marca').value,modelo:document.getElementById('nc-modelo').value,
    anio:parseInt(document.getElementById('nc-anio').value)||2025,placa:document.getElementById('nc-placa').value,
    desde:document.getElementById('nc-desde').value,hasta:document.getElementById('nc-hasta').value,
    comentario:document.getElementById('nc-comentario').value,
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

  // Clientes que vencen HOY
  const hoyVenc = mine.filter(c=>c.hasta===hoy && !['RENOVADO','EMITIDO','PÓLIZA VIGENTE'].includes(c.estado));
  if(hoyVenc.length){
    _notif(`⚠️ ${hoyVenc.length} póliza${hoyVenc.length>1?'s':''} vence${hoyVenc.length>1?'n':''} HOY`,
      hoyVenc.slice(0,3).map(c=>primerNombre(c.nombre)).join(', ') + (hoyVenc.length>3?` y ${hoyVenc.length-3} más`:''));
  }

  // Clientes que vencen en 7 días
  const venc7 = mine.filter(c=>{ const d=daysUntil(c.hasta); return d>0&&d<=7&&!['RENOVADO','EMITIDO','PÓLIZA VIGENTE'].includes(c.estado); });
  if(venc7.length){
    _notif(`📅 ${venc7.length} cliente${venc7.length>1?'s':''} vence${venc7.length>1?'n':''} en 7 días`,
      venc7.slice(0,3).map(c=>`${primerNombre(c.nombre)} (${daysUntil(c.hasta)}d)`).join(', '));
  }

  // Tareas pendientes de hoy o vencidas
  const tareasHoy = myTareas().filter(t=>t.estado==='pendiente' && t.fechaVence<=hoy);
  if(tareasHoy.length){
    _notif(`📌 ${tareasHoy.length} tarea${tareasHoy.length>1?'s':''} pendiente${tareasHoy.length>1?'s':''}`,
      tareasHoy.slice(0,3).map(t=>t.titulo).join(', '));
  }

  // Tareas con hora — notif cuando la hora actual coincide (±5 min)
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
      const renovados = mine.filter(c=>c.estado==='RENOVADO').length;
      const tasa = mine.length > 0 ? Math.round(renovados/mine.length*100) : 0;
      return {u, mine:mine.length, cierres:myCierres.length, prima, tasa};
    }).sort((a,b)=>b.prima-a.prima);

    rankingEl.innerHTML = `<table style="width:100%;font-size:13px">
      <thead><tr>
        <th style="padding:8px;text-align:left;color:var(--muted)">#</th>
        <th style="padding:8px;text-align:left;color:var(--muted)">Ejecutivo</th>
        <th style="padding:8px;color:var(--muted)">Clientes</th>
        <th style="padding:8px;color:var(--muted)">Cierres período</th>
        <th style="padding:8px;color:var(--muted)">Prima recaudada</th>
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
    // Gradient effect
    const grad=ctx.createLinearGradient(0,by,0,by+bh);
    grad.addColorStop(0,color); grad.addColorStop(1,color+'88');
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
  const periodo=document.getElementById('rep-periodo')?.value||'periodo';
  XLSX.writeFile(wb,`Reliance_Reporte_${periodo}_${new Date().toISOString().split('T')[0]}.xlsx`);
  showToast('Reporte exportado');
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
  wrap.innerHTML=Object.entries(ASEGURADORAS).map(([name,cfg])=>`
    <label style="display:flex;align-items:center;gap:6px;padding:5px 8px;border:1.5px solid ${cfg.color}33;border-radius:6px;cursor:pointer;font-size:11px;font-weight:600;color:${cfg.color};background:${cfg.color}0d;user-select:none;transition:background .15s" title="${name}">
      <input type="checkbox" class="aseg-check" data-aseg="${name}" checked
        style="accent-color:${cfg.color};width:13px;height:13px;cursor:pointer"
        onchange="this.closest('label').style.background=this.checked?'${cfg.color}22':'${cfg.color}08'">
      ${name}
    </label>`).join('');
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

    // Cargar las 4 listas en paralelo (spGetAll actualiza _cache internamente)
    const [clientes, cotizaciones, cierres, tareas] = await Promise.all([
      spGetAll('clientes'),
      spGetAll('cotizaciones'),
      spGetAll('cierres'),
      spGetAll('tareas'),
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
    // Cotizaciones
    const hashCot = cotizaciones.length + '|' + (cotizaciones[cotizaciones.length-1]?._spId||'');
    if(hashCot !== prevHashCot){
      localStorage.setItem('reliance_cotizaciones', JSON.stringify(cotizaciones));
      if(pg==='page-cotizaciones') renderCotizaciones();
      actualizarBadgeCotizaciones();
    }
    // Cierres
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

async function initApp(){
  // Cargar todos los datos desde SharePoint en paralelo
  const [cotiz, cierres, spUsers, spTareas] = await Promise.all([
    spGetAll('cotizaciones'),
    spGetAll('cierres'),
    spGetAll('usuarios'),
    spGetAll('tareas'),
  ]);
  await loadDBAsync();
  _cache.cotizaciones = cotiz;
  _cache.cierres      = cierres;
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
  renderDashboard();
  actualizarBadgeCotizaciones();
  actualizarBadgeTareas();
  renderDashTareas();

  // Ocultar login, mostrar app
  hideLoader();
  document.getElementById('login-screen').style.display='none';
  const appEl=document.querySelector('.app');
  if(appEl) appEl.style.display='flex';
  startAutoSync(); // Iniciar sync bidireccional automático
  // Notificaciones del navegador (leve delay para que cargue todo)
  setTimeout(initNotificaciones, 3000);
}
