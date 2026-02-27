// ═══════════════════════════════════════════════════════════════
//  SHAREPOINT LAYER — RelianceDesk
//  Reemplaza localStorage con SharePoint Lists via Graph API
//  Autenticación: MSAL.js v2 (SPA flow)
// ═══════════════════════════════════════════════════════════════

const SP_CONFIG = {
  clientId:  'f90218c5-1b8a-4d04-bf77-64374b34bd3f',
  tenantId:  '1d5a476a-076d-4e1a-a605-e8bdcfcf429f',
  siteUrl:   'https://reliancesa.sharepoint.com/sites/crmreliance',
  siteId:    null, // se resuelve automáticamente al iniciar
  lists: {
    clientes:      'CRM_Clientes',
    tareas:        'CRM_Tareas',
    cotizaciones:  'CRM_Cotizaciones',
    cierres:       'CRM_Cierres',
    usuarios:      'CRM_Usuarios',
  }
};

// MSAL instance
let _msalApp = null;
let _account  = null;
let _token    = null;
let _siteId   = null;
let _listIds  = {};   // nombre → id de lista en Graph
let _spReady  = false;

// Cache en memoria para evitar llamadas repetidas en la misma sesión
const _cache = {
  clientes:     null,
  cotizaciones: null,
  cierres:      null,
  usuarios:     null,
  tareas:       null,
};

// ── Inicializar MSAL ─────────────────────────────────────────
async function spInit(){
  try{ updateSpStatus('syncing', '⟳ Conectando...'); }catch(e){}

  const msalConfig = {
    auth: {
      clientId:    SP_CONFIG.clientId,
      authority:   `https://login.microsoftonline.com/${SP_CONFIG.tenantId}`,
      redirectUri: 'https://sgrandacrm.github.io/crm-reliance/',
    },
    cache: { cacheLocation: 'sessionStorage' }
  };

  _msalApp = new msal.PublicClientApplication(msalConfig);
  await _msalApp.initialize();

  // Manejar redirect de vuelta del login
  const resp = await _msalApp.handleRedirectPromise();
  if(resp) _account = resp.account;

  // Buscar cuenta ya logueada
  if(!_account){
    const accounts = _msalApp.getAllAccounts();
    if(accounts.length > 0) _account = accounts[0];
  }

  if(!_account){
    // No hay sesión — mostrar pantalla de login
    showSpLogin();
    return false;
  }

  // Obtener token silencioso
  try{
    const tokenResp = await _msalApp.acquireTokenSilent({
      scopes: ['https://graph.microsoft.com/Sites.ReadWrite.All',
               'https://graph.microsoft.com/Sites.Manage.All',
               'https://graph.microsoft.com/Files.ReadWrite',
               'https://graph.microsoft.com/User.Read'],
      account: _account
    });
    _token = tokenResp.accessToken;
  }catch(e){
    // Token expirado — obtener con popup
    try{
      const tokenResp = await _msalApp.acquireTokenPopup({
        scopes: ['https://graph.microsoft.com/Sites.ReadWrite.All',
                 'https://graph.microsoft.com/Sites.Manage.All',
                 'https://graph.microsoft.com/Files.ReadWrite',
                 'https://graph.microsoft.com/User.Read'],
      });
      _token = tokenResp.accessToken;
      _account = tokenResp.account;
    }catch(e2){
      showSpError('Error de autenticación: ' + e2.message);
      return false;
    }
  }

  // Resolver Site ID
  try{
    const hostname = 'reliancesa.sharepoint.com';
    const sitePath = '/sites/crmreliance';
    const siteResp = await spGraph(`sites/${hostname}:${sitePath}`);
    _siteId = siteResp.id;
  }catch(e){
    showSpError('No se puede acceder al sitio SharePoint: ' + e.message);
    return false;
  }

  _spReady = true;
  updateSpStatus('online', '● SharePoint');
  return true;
}

// ── Llamada a Graph API ──────────────────────────────────────
async function spGraph(endpoint, method='GET', body=null){
  // Renovar token si expiró
  if(_msalApp && _account){
    try{
      const r = await _msalApp.acquireTokenSilent({
        scopes:['https://graph.microsoft.com/Sites.ReadWrite.All',
                'https://graph.microsoft.com/Files.ReadWrite'],
        account: _account
      });
      _token = r.accessToken;
    }catch(e){}
  }

  const opts = {
    method,
    headers: {
      'Authorization': `Bearer ${_token}`,
      'Content-Type':  'application/json',
      'Accept':        'application/json',
    }
  };
  if(body) opts.body = JSON.stringify(body);

  const url = endpoint.startsWith('http')
    ? endpoint
    : `https://graph.microsoft.com/v1.0/${endpoint}`;

  const resp = await fetch(url, opts);
  if(!resp.ok){
    const err = await resp.json().catch(()=>({error:{message:resp.statusText}}));
    throw new Error(err.error?.message || resp.statusText);
  }
  if(resp.status===204) return null;
  return resp.json();
}

// ── Resolver ID de una lista por nombre ─────────────────────
async function spGetListId(listName){
  if(_listIds[listName]) return _listIds[listName];
  try{
    const r = await spGraph(`sites/${_siteId}/lists/${listName}`);
    _listIds[listName] = r.id;
    return r.id;
  }catch(e){
    return null; // lista no existe aún
  }
}

// ── Obtener TODOS los ítems de una lista ────────────────────
async function spGetAll(listKey){
  if(!_spReady) return _cache[listKey] || [];

  const listName = SP_CONFIG.lists[listKey];
  updateSpStatus('syncing','⟳ Cargando...');

  try{
    let items = [];
    const listId = await spGetListId(listName);
    if(!listId) throw new Error('No se encontró la lista en SharePoint: ' + listName);
    let url = `sites/${_siteId}/lists/${listId}/items?$expand=fields&$top=999`;

    while(url){
      const r = await spGraph(url);
      items = items.concat(r.value || []);
      url = r['@odata.nextLink'] || null;
    }

    // Leer campos individuales de SharePoint
    const result = items.map(item => {
      const f = item.fields;
      const obj = { _spId: item.id };
      Object.keys(f).forEach(k => {
        if(k.startsWith('_') || k==='ContentType' || k==='Attachments' || k==='LinkTitle') return;
        let val = f[k];
        if(typeof val === 'string' && (val.startsWith('[') || val.startsWith('{'))){
          try{ val = JSON.parse(val); }catch(e){}
        }
        // Normalizar fechas ISO que SP devuelve con T00:00:00Z
        if(typeof val === 'string' && val.length>=10 && val[10]==='T' && /^\d{4}-\d{2}-\d{2}T/.test(val)){
          val = val.substring(0,10);
        }
        obj[k] = val;
      });
      // Restaurar id desde crm_id
      if(f.crm_id) obj.id = isNaN(f.crm_id) ? f.crm_id : Number(f.crm_id);
      // Restaurar nombre desde Title
      if(f.Title){
        if(listKey==='clientes'||listKey==='usuarios') obj.nombre = f.Title;
        if(listKey==='cotizaciones') obj.codigo = f.Title;
        if(listKey==='cierres') obj.polizaNueva = f.Title;
      }
      // Aplicar defaults según tipo de lista para evitar undefined en UI
      if(listKey==='cotizaciones'){
        obj.clienteNombre = obj.clienteNombre || obj.Title || '(sin nombre)';
        obj.codigo        = obj.codigo || obj.Title || '';
        obj.estado        = obj.estado || 'ENVIADA';
        obj.fecha         = obj.fecha  || '';
        obj.vehiculo      = obj.vehiculo || '';
        obj.va            = obj.va || 0;
        obj.resultados    = obj.resultados || [];
        obj.aseguradoras  = obj.aseguradoras || [];
        obj.version       = obj.version || 1;
      }
      if(listKey==='tareas'){
        obj.titulo      = obj.titulo      || obj.Title || '(sin título)';
        obj.estado      = obj.estado      || 'pendiente';
        obj.prioridad   = obj.prioridad   || 'media';
        obj.tipo        = obj.tipo        || 'llamada';
        obj.fechaVence  = obj.fechaVence  || '';
        obj.horaVence   = obj.horaVence   || '';
        obj.clienteNombre = obj.clienteNombre || '';
        obj.descripcion = obj.descripcion || '';
      }
      if(listKey==='clientes'){
        obj.nombre       = obj.nombre || obj.Title || '(sin nombre)';
        obj.estado       = obj.estado || 'PENDIENTE';
        obj.aseguradora  = obj.aseguradora || '';
        obj.hasta        = obj.hasta || '';
        obj.va           = obj.va || 0;
      }
      if(listKey==='cierres'){
        obj.clienteNombre = obj.clienteNombre || obj.Title || '(sin nombre)';
        obj.aseguradora   = obj.aseguradora || '';
        obj.primaTotal    = obj.primaTotal || 0;
        obj.fechaRegistro = obj.fechaRegistro || '';
      }
      return obj;
    });

    _cache[listKey] = result;
    updateSpStatus('online','● SharePoint');
    return result;
  }catch(e){
    updateSpStatus('error','⚠ Error SP');
    console.error('spGetAll error:', listKey, e);
    return _cache[listKey] || [];
  }
}

// ── Crear ítem en lista ──────────────────────────────────────
async function spCreate(listKey, data){
  if(!_spReady){ console.warn('SP not ready'); return null; }

  const listName = SP_CONFIG.lists[listKey];
  updateSpStatus('syncing','⟳ Guardando...');

  try{
    const listId = await spGetListId(listName);
    if(!listId) throw new Error('No se encontró la lista en SharePoint: ' + listName);
    const fields = spToFields(listKey, data);
    const r = await spGraph(
      `sites/${_siteId}/lists/${listId}/items`,
      'POST',
      { fields }
    );
    // Actualizar cache sin borrarlo — evita que la lista quede vacía
    if(_cache[listKey]) _cache[listKey] = _cache[listKey].filter(x=>String(x.id)!==String(data.id));
    if(_cache[listKey]) _cache[listKey].push({...data, _spId: r.id});
    updateSpStatus('online','● SharePoint');
    return r.id;
  }catch(e){
    // Si falla (campo no existe / tipo incorrecto), reintentar con campos mínimos
    if(e.message){
      try{
        const listId = await spGetListId(listName);
        const minFields = { Title: spToFields(listKey, data).Title || String(data.id||''), crm_id: String(data.id||'') };
        const r2 = await spGraph(`sites/${_siteId}/lists/${listId}/items`, 'POST', { fields: minFields });
        // Actualizar cache con datos completos (aunque SP solo tenga campos mínimos)
        if(_cache[listKey]) _cache[listKey] = _cache[listKey].filter(x=>String(x.id)!==String(data.id));
        if(_cache[listKey]) _cache[listKey].push({...data, _spId: r2.id});
        updateSpStatus('online','● SharePoint');
        console.warn('spCreate: guardado con campos mínimos (columnas SP incompletas):', listKey);
        // Intentar update completo tras creación mínima (columnas pueden estar creándose)
        setTimeout(async ()=>{
          try{
            const allFields = spToFields(listKey, data);
            await spGraph(`sites/${_siteId}/lists/${listId}/items/${r2.id}/fields`, 'PATCH', allFields);
          }catch(e2){ /* columnas aún no existen, se sincronizará luego */ }
        }, 3000);
        return r2.id;
      }catch(e2){
        updateSpStatus('error','⚠ Error al guardar');
        console.error('spCreate error:', listKey, e2);
        return null;
      }
    }
    updateSpStatus('error','⚠ Error al guardar');
    console.error('spCreate error:', listKey, e);
    return null;
  }
}

// ── Actualizar ítem existente ────────────────────────────────
async function spUpdate(listKey, spId, data){
  if(!_spReady){ console.warn('SP not ready'); return false; }

  const listName = SP_CONFIG.lists[listKey];
  updateSpStatus('syncing','⟳ Actualizando...');

  try{
    const listId = await spGetListId(listName);
    if(!listId) throw new Error('No se encontró la lista en SharePoint: ' + listName);
    const fields = spToFields(listKey, data);
    await spGraph(
      `sites/${_siteId}/lists/${listId}/items/${spId}/fields`,
      'PATCH',
      fields
    );
    // Actualizar item en cache sin borrar toda la lista
    if(_cache[listKey]){
      const ci = _cache[listKey].findIndex(x=>x._spId===spId||String(x.id)===String(data.id));
      if(ci>=0) _cache[listKey][ci] = {..._cache[listKey][ci], ...data};
    }
    updateSpStatus('online','● SharePoint');
    return true;
  }catch(e){
    if(e.message){
      try{
        const listId = await spGetListId(listName);
        const minFields = { Title: spToFields(listKey, data).Title || String(data.id||''), crm_id: String(data.id||'') };
        await spGraph(`sites/${_siteId}/lists/${listId}/items/${spId}/fields`, 'PATCH', minFields);
        // Mantener cache actualizado
        if(_cache[listKey]){
          const ci = _cache[listKey].findIndex(x=>x._spId===spId||String(x.id)===String(data.id));
          if(ci>=0) _cache[listKey][ci] = {..._cache[listKey][ci], ...data};
        }
        updateSpStatus('online','● SharePoint');
        console.warn('spUpdate: actualizado con campos mínimos:', listKey, '— error original:', e.message);
        return true;
      }catch(e2){
        updateSpStatus('error','⚠ Error al actualizar');
        console.error('spUpdate error:', listKey, e2);
        return false;
      }
    }
    updateSpStatus('error','⚠ Error al actualizar');
    console.error('spUpdate error:', listKey, e);
    return false;
  }
}

// ── Eliminar ítem ────────────────────────────────────────────
async function spDelete(listKey, spId){
  if(!_spReady) return false;
  const listName = SP_CONFIG.lists[listKey];
  try{
    await spGraph(`sites/${_siteId}/lists/${listName}/items/${spId}`, 'DELETE');
    // Remover solo el item borrado del cache
    if(_cache[listKey]) _cache[listKey] = _cache[listKey].filter(x=>x._spId!==spId);
    return true;
  }catch(e){
    console.error('spDelete error:', e);
    return false;
  }
}

// ── Convertir objeto JS → campos de lista SP ─────────────────
// Campos que SharePoint no acepta como nombre de columna
const SP_SKIP = new Set(['id','ID','version','Version']);

function spToFields(listKey, data){
  // Solo campos que realmente existen en cada lista SP
  const validos = {
    clientes:     new Set(['Title','ci','tipo','region','ciudad','aseguradora','ejecutivo','estado','placa','marca','modelo','anio','va','pn','primaTotal','desde','hasta','celular','correo','nota','ultimoContacto','factura','poliza','obs','color','motor','chasis','dep','tasa','axavd','formaPago','crm_id','polizaNueva','aseguradoraAnterior','historialWa','bitacora']),
    tareas: new Set(['Title','titulo','descripcion','clienteId','clienteNombre','fechaVence','horaVence','tipo','prioridad','estado','ejecutivo','fechaCreacion','crm_id']),
    cotizaciones: new Set(['Title','codigo','version','fecha','ejecutivo','clienteNombre','clienteCI','clienteId','celular','correo','ciudad','region','tipo','vehiculo','marca','modelo','anio','placa','color','motor','chasis','va','desde','hasta','asegAnterior','polizaAnterior','estado','asegElegida','resultados','aseguradoras','obsAcept','fechaAcept','reemplazadaPor','crm_id','cuotasTc','cuotasDeb','autoSust']),
    cierres:      new Set(['Title','clienteNombre','aseguradora','primaTotal','primaNeta','vigDesde','vigHasta','formaPago','facturaAseg','ejecutivo','fechaRegistro','observacion','axavd','crm_id','polizaNueva','cuenta','clienteId','cotizacionId','comision','comisionPct']),
    usuarios:     new Set(['Title','userId','rol','email','activo','color','initials','crm_id']),
  };
  const permitidos = validos[listKey] || new Set();
  const camposNum  = new Set(['anio','va','pn','primaTotal','dep','tasa','version','primaNeta','cuotasTc','cuotasDeb','comision','comisionPct']);
  const ignorar    = new Set(['id','nombre','name','pass','password','_spId','_dirty','_spEtag']);

  const fields = {};

  // Title = identificador principal
  let title = '';
  if(listKey==='clientes')     title = data.nombre || data.name || data.id || '';
  if(listKey==='tareas')       title = data.titulo || data.id || '';
  if(listKey==='cotizaciones') title = data.codigo || data.id || '';
  if(listKey==='cierres')      title = data.polizaNueva || data.id || '';
  if(listKey==='usuarios')     title = data.nombre || data.name || data.email || '';
  fields['Title'] = String(title||'').substring(0, 255);

  // id → crm_id
  if(data.id !== undefined && permitidos.has('crm_id')){
    fields['crm_id'] = String(data.id);
  }

  Object.keys(data).forEach(k => {
    if(ignorar.has(k) || k.startsWith('_')) return;
    if(!permitidos.has(k)) return;
    let val = data[k];
    if(val !== null && val !== undefined && typeof val === 'object') val = JSON.stringify(val);
    if(typeof val === 'boolean') val = String(val);
    if(camposNum.has(k) && val !== '' && val !== null && val !== undefined) val = parseFloat(val)||0;
    if(val === undefined || val === null) return;
    if(typeof val === 'string') val = val.substring(0, 3999);
    fields[k] = val;
  });

  return fields;
}

// ── SETUP AUTOMÁTICO DE LISTAS ───────────────────────────────
async function spSetupLists(onProgress){
  onProgress('Verificando listas en SharePoint...');

  const listDefs = {
    CRM_Clientes: [
      {name:'ci',type:'Text'},{name:'tipo',type:'Text'},{name:'region',type:'Text'},
      {name:'ciudad',type:'Text'},{name:'aseguradora',type:'Text'},{name:'ejecutivo',type:'Text'},
      {name:'estado',type:'Text'},{name:'placa',type:'Text'},{name:'marca',type:'Text'},
      {name:'modelo',type:'Text'},{name:'anio',type:'Number'},{name:'va',type:'Number'},
      {name:'pn',type:'Number'},{name:'primaTotal',type:'Number'},
      {name:'desde',type:'Text'},{name:'hasta',type:'Text'},
      {name:'celular',type:'Text'},{name:'correo',type:'Text'},
      {name:'nota',type:'Note'},{name:'ultimoContacto',type:'Text'},
      {name:'factura',type:'Text'},{name:'poliza',type:'Text'},
      {name:'polizaNueva',type:'Text'},{name:'aseguradoraAnterior',type:'Text'},
      {name:'obs',type:'Text'},{name:'color',type:'Text'},
      {name:'motor',type:'Text'},{name:'chasis',type:'Text'},
      {name:'dep',type:'Number'},{name:'tasa',type:'Number'},
      {name:'axavd',type:'Text'},{name:'formaPago',type:'Text'},
      {name:'historialWa',type:'note'},{name:'crmid',type:'text'},
    ],
    CRM_Tareas: [
      {name:'titulo',type:'Text'},{name:'descripcion',type:'Note'},
      {name:'clienteId',type:'Text'},{name:'clienteNombre',type:'Text'},
      {name:'fechaVence',type:'Text'},{name:'horaVence',type:'Text'},
      {name:'tipo',type:'Text'},{name:'prioridad',type:'Text'},
      {name:'estado',type:'Text'},{name:'ejecutivo',type:'Text'},
      {name:'fechaCreacion',type:'Text'},{name:'crm_id',type:'Text'},
    ],
    CRM_Cotizaciones: [
      {name:'codigo',type:'Text'},{name:'version',type:'Number'},
      {name:'fecha',type:'Text'},{name:'ejecutivo',type:'Text'},
      {name:'clienteNombre',type:'Text'},{name:'clienteCI',type:'Text'},
      {name:'clienteId',type:'Text'},{name:'celular',type:'Text'},
      {name:'correo',type:'Text'},{name:'ciudad',type:'Text'},
      {name:'region',type:'Text'},{name:'tipo',type:'Text'},
      {name:'vehiculo',type:'Text'},{name:'marca',type:'Text'},
      {name:'modelo',type:'Text'},{name:'anio',type:'Number'},
      {name:'placa',type:'Text'},{name:'color',type:'Text'},
      {name:'motor',type:'Text'},{name:'chasis',type:'Text'},
      {name:'va',type:'Number'},{name:'desde',type:'Text'},
      {name:'hasta',type:'Text'},{name:'asegAnterior',type:'Text'},
      {name:'polizaAnterior',type:'Text'},
      {name:'estado',type:'Text'},{name:'asegElegida',type:'Text'},
      {name:'resultados',type:'Note'},{name:'aseguradoras',type:'Note'},
      {name:'obsAcept',type:'Note'},{name:'fechaAcept',type:'Text'},
      {name:'reemplazadaPor',type:'Text'},{name:'autoSust',type:'Boolean'},
      {name:'cuotasTc',type:'Number'},{name:'cuotasDeb',type:'Number'},
      {name:'spLocalId',type:'Text'},
    ],
    CRM_Cierres: [
      {name:'clienteNombre',type:'Text'},{name:'aseguradora',type:'Text'},
      {name:'primaTotal',type:'Number'},{name:'primaNeta',type:'Number'},
      {name:'vigDesde',type:'Text'},{name:'vigHasta',type:'Text'},
      {name:'formaPago',type:'Text'},{name:'facturaAseg',type:'Text'},
      {name:'ejecutivo',type:'Text'},{name:'fechaRegistro',type:'Text'},
      {name:'observacion',type:'Note'},{name:'axavd',type:'Text'},
      {name:'cuenta',type:'Text'},{name:'spLocalId',type:'Text'},
    ],
    CRM_Usuarios: [
      {name:'userId',type:'Text'},{name:'rol',type:'Text'},
      {name:'email',type:'Text'},{name:'activo',type:'Boolean'},
    ],
  };

  for(const [listName, columns] of Object.entries(listDefs)){
    onProgress(`Configurando lista: ${listName}...`);
    const exists = await spGetListId(listName.replace('CRM_','').toLowerCase());

    // Crear lista si no existe
    let listId = _listIds[listName];
    if(!listId){
      try{
        const newList = await spGraph(`sites/${_siteId}/lists`, 'POST', {
          displayName: listName,
          list: { template: 'genericList' }
        });
        listId = newList.id;
        _listIds[listName] = listId;
        onProgress(`✓ Lista creada: ${listName}`);
      }catch(e){
        onProgress(`⚠ Error creando ${listName}: ${e.message}`);
        continue;
      }
    }

    // Agregar columnas que faltan
    for(const col of columns){
      try{
        const colDef = { name: col.name, hidden: false };
        if(col.type==='Text')     colDef.text = {};
        if(col.type==='Note')     colDef.text = { allowMultipleLines: true };
        if(col.type==='Number')   colDef.number = {};
        if(col.type==='DateTime') colDef.dateTime = { format: 'dateOnly' };
        if(col.type==='Boolean')  colDef.boolean = {};
        await spGraph(`sites/${_siteId}/lists/${listId}/columns`, 'POST', colDef);
      }catch(e){
        // Columna ya existe — ignorar
      }
    }
    onProgress(`✓ ${listName} lista`);
  }

  onProgress('✅ Configuración completada');
  localStorage.setItem('sp_setup_done', '1');
}

// ── MIGRACIÓN desde localStorage ────────────────────────────
async function spMigrateFromLocal(onProgress){
  onProgress('Migrando datos existentes...');

  const maps = [
    ['reliance_clientes',     'clientes'],
    ['reliance_cotizaciones', 'cotizaciones'],
    ['reliance_cierres',      'cierres'],
    ['reliance_users',        'usuarios'],
  ];

  for(const [lsKey, spKey] of maps){
    const raw = localStorage.getItem(lsKey);
    if(!raw) continue;
    try{
      const items = JSON.parse(raw);
      if(!items.length) continue;
      onProgress(`Migrando ${items.length} registros de ${spKey}...`);
      for(const item of items){
        await spCreate(spKey, item);
      }
      onProgress(`✓ ${spKey}: ${items.length} registros migrados`);
    }catch(e){
      onProgress(`⚠ Error migrando ${spKey}: ${e.message}`);
    }
  }
  onProgress('✅ Migración completada');
}

// ── STATUS INDICATOR ─────────────────────────────────────────
function updateSpStatus(state, text){
  const el = document.getElementById('sp-status');
  if(!el) return;
  el.className = 'sp-status ' + (state||'');
  // Count pending dirty records if available
  let badge = '';
  if(typeof _countDirty === 'function' && state !== 'syncing'){
    const n = _countDirty();
    if(n > 0){
      badge = ` <span style="background:#c84b1a;color:#fff;border-radius:99px;padding:1px 6px;font-size:10px;font-weight:700;margin-left:3px" title="${n} cambio(s) pendiente(s) de sincronizar">${n}</span>`;
    }
  }
  el.innerHTML = (text||'') + badge;
  el.title = state==='online' ? (badge?`${(el.querySelector('span')?el.querySelector('span').title:'')}`:'Sincronizado con SharePoint') : (text||'');
  el.style.cursor = (state==='online'||state==='error') ? 'pointer' : 'default';
  el.onclick = (state==='online'||state==='error') && typeof _forceSync==='function'
    ? ()=>_forceSync()
    : null;
}

// ── LOGIN UI ─────────────────────────────────────────────────
let _showSpLoginRetries = 0;
function showSpLogin(){
  // Ocultar loader con pequeño delay para que el DOM esté listo
  setTimeout(hideLoader, 100);
  const setupEl = document.getElementById('sp-setup');
  if(!setupEl){
    if(_showSpLoginRetries++ < 20) setTimeout(showSpLogin, 200);
    return;
  }
  _showSpLoginRetries = 0;
  setupEl.style.display = 'flex';
  setupEl.innerHTML = `
    <div class="setup-card">
      <div style="font-size:40px;margin-bottom:16px">🛡</div>
      <h2>RelianceDesk</h2>
      <p>Inicia sesión con tu cuenta Microsoft para acceder a RelianceDesk compartido en SharePoint.</p>
      <button class="btn btn-primary" style="width:100%;justify-content:center;font-size:15px;padding:14px"
        onclick="spLogin()">
        <img src="https://learn.microsoft.com/en-us/azure/active-directory/develop/media/howto-add-branding-in-apps/ms-symbollockup_mssymbol_19.svg"
          style="width:18px;height:18px;filter:brightness(10)">
        Iniciar sesión con Microsoft
      </button>
      <p style="margin-top:12px;font-size:11px">
        Solo para cuentas @reliance.ec
      </p>
    </div>`;
}

async function spLogin(){
  try{
    await _msalApp.loginRedirect({
      scopes: ['https://graph.microsoft.com/Sites.ReadWrite.All',
               'https://graph.microsoft.com/Sites.Manage.All',
               'https://graph.microsoft.com/Files.ReadWrite',
               'https://graph.microsoft.com/User.Read'],
    });
  }catch(e){
    showSpError('Error de login: ' + e.message);
  }
}

function showSpError(msg){
  hideLoader();
  const setupEl = document.getElementById('sp-setup');
  if(setupEl){
    setupEl.style.display='flex';
    setupEl.innerHTML=`<div class="setup-card">
      <div style="font-size:40px;margin-bottom:16px">⚠️</div>
      <h2>Error de conexión</h2>
      <p>${msg}</p>
      <button class="btn btn-secondary" style="width:100%;justify-content:center;margin-top:8px"
        onclick="location.reload()">Reintentar</button>
    </div>`;
  }
}

function hideLoader(){
  try{
    const l = document.getElementById('sp-loader');
    if(l) l.style.display='none';
  }catch(e){}
}

// ── SETUP UI ─────────────────────────────────────────────────
async function spRunSetup(){
  const setupEl = document.getElementById('sp-setup');
  if(!setupEl) return;
  setupEl.style.display='flex';
  setupEl.innerHTML=`<div class="setup-card">
    <div style="font-size:40px;margin-bottom:16px">⚙️</div>
    <h2>Configurando SharePoint</h2>
    <p>Creando las listas y estructura necesaria. Esto solo ocurre una vez.</p>
    <div class="setup-steps" id="setup-steps"></div>
    <div class="spinner" style="margin:0 auto"></div>
  </div>`;

  const steps = document.getElementById('setup-steps');
  const log = (msg) => {
    const div = document.createElement('div');
    const done = msg.startsWith('✓') || msg.startsWith('✅');
    const err  = msg.startsWith('⚠');
    div.className = `setup-step ${done?'done':err?'error':'active'}`;
    div.textContent = msg;
    if(steps) steps.appendChild(div);
  };

  await spSetupLists(log);

  // Migrar datos si hay en localStorage
  const hasLocal = ['reliance_clientes','reliance_cotizaciones','reliance_cierres']
    .some(k=>localStorage.getItem(k) && JSON.parse(localStorage.getItem(k)||'[]').length>0);

  if(hasLocal){
    log('Datos locales detectados. Migrando...');
    await spMigrateFromLocal(log);
  }

  setupEl.style.display='none';
  localStorage.setItem('sp_setup_done', '1');
  showToast('✅ SharePoint configurado correctamente','success');
  await initApp();
}

// ── Verificar que las listas existen en SP ──────────────────
async function spVerificarListas(){
  try{
    const r = await spGraph(`sites/${_siteId}/lists/CRM_Clientes`);
    return !!r.id;
  }catch(e){
    return false;
  }
}

// ── Crear columnas individuales en cada lista ────────────────
// Si la columna ya existe el error se ignora silenciosamente
async function spAsegurarColumnas(logCol){
  if(!logCol) logCol=()=>{};

  const colsDef = {
    CRM_Clientes: [
      {name:'ci',text:{}},{name:'tipo',text:{}},{name:'region',text:{}},
      {name:'ciudad',text:{}},{name:'aseguradora',text:{}},{name:'ejecutivo',text:{}},
      {name:'estado',text:{}},{name:'placa',text:{}},{name:'marca',text:{}},
      {name:'modelo',text:{}},{name:'anio',number:{}},{name:'va',number:{}},
      {name:'pn',number:{}},{name:'primaTotal',number:{}},
      {name:'desde',text:{}},{name:'hasta',text:{}},
      {name:'celular',text:{}},{name:'correo',text:{}},
      {name:'nota',text:{allowMultipleLines:true}},
      {name:'ultimoContacto',text:{}},
      {name:'factura',text:{}},{name:'poliza',text:{}},
      {name:'polizaNueva',text:{}},{name:'aseguradoraAnterior',text:{}},
      {name:'obs',text:{}},{name:'color',text:{}},
      {name:'motor',text:{}},{name:'chasis',text:{}},
      {name:'dep',number:{}},{name:'tasa',number:{}},
      {name:'axavd',text:{}},{name:'formaPago',text:{}},
      {name:'historialWa',text:{allowMultipleLines:true}},
      {name:'bitacora',text:{allowMultipleLines:true}},
      {name:'crm_id',text:{}},
    ],
    CRM_Tareas: [
      {name:'titulo',text:{}},{name:'descripcion',text:{allowMultipleLines:true}},
      {name:'clienteId',text:{}},{name:'clienteNombre',text:{}},
      {name:'fechaVence',text:{}},{name:'horaVence',text:{}},
      {name:'tipo',text:{}},{name:'prioridad',text:{}},
      {name:'estado',text:{}},{name:'ejecutivo',text:{}},
      {name:'fechaCreacion',text:{}},{name:'crm_id',text:{}},
    ],
    CRM_Cotizaciones: [
      {name:'codigo',text:{}},{name:'version',number:{}},
      {name:'fecha',text:{}},{name:'ejecutivo',text:{}},
      {name:'clienteNombre',text:{}},{name:'clienteCI',text:{}},
      {name:'clienteId',text:{}},{name:'celular',text:{}},
      {name:'correo',text:{}},{name:'ciudad',text:{}},
      {name:'region',text:{}},{name:'tipo',text:{}},
      {name:'vehiculo',text:{}},{name:'marca',text:{}},
      {name:'modelo',text:{}},{name:'anio',number:{}},
      {name:'placa',text:{}},{name:'color',text:{}},
      {name:'motor',text:{}},{name:'chasis',text:{}},
      {name:'va',number:{}},{name:'desde',text:{}},
      {name:'hasta',text:{}},{name:'asegAnterior',text:{}},
      {name:'polizaAnterior',text:{}},
      {name:'estado',text:{}},{name:'asegElegida',text:{}},
      {name:'resultados',text:{allowMultipleLines:true}},
      {name:'aseguradoras',text:{allowMultipleLines:true}},
      {name:'obsAcept',text:{allowMultipleLines:true}},
      {name:'fechaAcept',text:{}},{name:'reemplazadaPor',text:{}},
      {name:'crm_id',text:{}},
      {name:'cuotasTc',number:{}},{name:'cuotasDeb',number:{}},{name:'autoSust',text:{}},
    ],
    CRM_Cierres: [
      {name:'clienteNombre',text:{}},{name:'aseguradora',text:{}},
      {name:'primaTotal',number:{}},{name:'primaNeta',number:{}},
      {name:'vigDesde',text:{}},{name:'vigHasta',text:{}},
      {name:'formaPago',text:{}},{name:'facturaAseg',text:{}},
      {name:'ejecutivo',text:{}},{name:'fechaRegistro',text:{}},
      {name:'observacion',text:{allowMultipleLines:true}},
      {name:'axavd',text:{}},{name:'polizaNueva',text:{}},
      {name:'crm_id',text:{}},
      {name:'cuenta',text:{}},{name:'clienteId',text:{}},{name:'cotizacionId',text:{}},
      {name:'comision',number:{}},{name:'comisionPct',number:{}},
    ],
    CRM_Usuarios: [
      {name:'userId',text:{}},{name:'rol',text:{}},
      {name:'email',text:{}},{name:'activo',text:{}},
      {name:'color',text:{}},{name:'initials',text:{}},
      {name:'crm_id',text:{}},
    ],
  };

  for(const [listName, cols] of Object.entries(colsDef)){
    logCol(`Configurando ${listName}...`);
    // Use the full SP list name (e.g. 'CRM_Cotizaciones') directly — NOT the short key name
    const listId = await spGetListId(listName);
    if(!listId){ logCol(`⚠ No se encontró ${listName}`); continue; }
    let ok=0, skip=0;
    for(const col of cols){
      try{
        await spGraph(`sites/${_siteId}/lists/${listId}/columns`, 'POST', col);
        ok++;
      }catch(e){ skip++; } // Ya existe — ok
    }
    logCol(`✅ ${listName}: ${ok} columnas nuevas, ${skip} ya existían`);
  }
}

async function _spAsegurarColumnas_UNUSED(logCol){
  if(!logCol) logCol=()=>{};
  const esquema = {
    CRM_Clientes: [
      {name:'ci',type:'text'},{name:'tipo',type:'text'},{name:'region',type:'text'},
      {name:'ciudad',type:'text'},{name:'aseguradora',type:'text'},{name:'ejecutivo',type:'text'},
      {name:'estado',type:'text'},{name:'placa',type:'text'},{name:'marca',type:'text'},
      {name:'modelo',type:'text'},{name:'anio',type:'number'},{name:'va',type:'number'},
      {name:'pn',type:'number'},{name:'primaTotal',type:'number'},
      {name:'desde',type:'text'},{name:'hasta',type:'text'},
      {name:'celular',type:'text'},{name:'correo',type:'text'},
      {name:'nota',type:'note'},{name:'ultimoContacto',type:'text'},
      {name:'factura',type:'text'},{name:'poliza',type:'text'},
      {name:'obs',type:'text'},{name:'color',type:'text'},
      {name:'motor',type:'text'},{name:'chasis',type:'text'},
      {name:'dep',type:'number'},{name:'tasa',type:'number'},
      {name:'axavd',type:'text'},{name:'formaPago',type:'text'},
      {name:'crmid',type:'text'},{name:'polizaNueva',type:'text'},
      {name:'aseguradoraAnterior',type:'text'},{name:'historialWa',type:'note'},{name:'bitacora',type:'Note'},
    ],
    CRM_Cotizaciones: [
      {name:'codigo',type:'text'},
      {name:'fecha',type:'text'},{name:'ejecutivo',type:'text'},
      {name:'clienteNombre',type:'text'},{name:'clienteCI',type:'text'},
      {name:'clienteId',type:'text'},{name:'ciudad',type:'text'},
      {name:'vehiculo',type:'text'},{name:'placa',type:'text'},
      {name:'va',type:'number'},{name:'desde',type:'text'},
      {name:'estado',type:'text'},{name:'asegElegida',type:'text'},
      {name:'resultados',type:'note'},{name:'aseguradoras',type:'note'},
      {name:'obsAcept',type:'note'},{name:'fechaAcept',type:'text'},
      {name:'reemplazadaPor',type:'text'},{name:'crmid',type:'text'},
    ],
    CRM_Cierres: [
      {name:'clienteNombre',type:'text'},{name:'aseguradora',type:'text'},
      {name:'primaTotal',type:'number'},{name:'primaNeta',type:'number'},
      {name:'vigDesde',type:'text'},{name:'vigHasta',type:'text'},
      {name:'formaPago',type:'text'},{name:'facturaAseg',type:'text'},
      {name:'ejecutivo',type:'text'},{name:'fechaRegistro',type:'text'},
      {name:'observacion',type:'note'},{name:'axavd',type:'text'},
      {name:'crmid',type:'text'},{name:'polizaNueva',type:'text'},
    ],
    CRM_Usuarios: [
      {name:'userId',type:'text'},{name:'rol',type:'text'},
      {name:'email',type:'text'},{name:'activo',type:'text'},
      {name:'color',type:'text'},{name:'initials',type:'text'},
    ],
  };

  for(const [listName, cols] of Object.entries(esquema)){
    logCol(`Configurando ${listName}...`);
    let ok=0;
    for(const col of cols){
      try{
        const def = { name: col.name };
        if(col.type==='text')   def.text = {};
        if(col.type==='note')   def.text = { allowMultipleLines: true };
        if(col.type==='number') def.number = {};
        await spGraph(`sites/${_siteId}/lists/${listName}/columns`, 'POST', def);
        ok++;
      }catch(e){
        // Columna ya existe — ignorar
      }
    }
    logCol(`✅ ${listName}: ${ok} columnas`);
  }
}

// ── ARRANQUE PRINCIPAL ───────────────────────────────────────
async function bootApp(){
  // Timeout de seguridad: si en 30s no carga, mostrar error
  const safetyTimeout = setTimeout(()=>{
    hideLoader();
    showSpError('Tiempo de espera agotado. Verifica tu conexión y vuelve a intentarlo.');
  }, 120000);

  try{
    // 1. Inicializar MSAL y autenticar
    const ok = await spInit();
    clearTimeout(safetyTimeout);
    if(!ok) return;

    // 2. Verificar que las listas existen realmente en SharePoint
    const listasOk = await spVerificarListas();

    if(listasOk){
      // Verificar si ya se crearon columnas antes
      const colsDone = localStorage.getItem('sp_cols_done');
      if(!colsDone || !['3','4','5','6','7'].includes(colsDone)){
        hideLoader();
        const setupEl = document.getElementById('sp-setup');
        if(setupEl){
          setupEl.style.display='flex';
          setupEl.innerHTML=`<div class="setup-card">
            <div style="font-size:40px;margin-bottom:16px">⚙️</div>
            <h2>Configurando campos</h2>
            <p>Creando columnas en SharePoint. Solo ocurre una vez.</p>
            <div class="setup-steps" id="col-steps" style="max-height:180px;overflow-y:auto"></div>
            <div class="spinner" style="margin:16px auto 0"></div>
          </div>`;
        }
        const logCol=(msg)=>{
          const el=document.getElementById('col-steps');
          if(!el) return;
          const d=document.createElement('div');
          d.className='setup-step '+(msg.startsWith('✅')?'done':msg.startsWith('⚠')?'error':'active');
          d.textContent=msg; el.appendChild(d); el.scrollTop=el.scrollHeight;
        };
        await spAsegurarColumnas(logCol);
        logCol('✅ Columnas configuradas');
        localStorage.setItem('sp_cols_done','7');
        if(setupEl) setupEl.style.display='none';
      }
    }

    if(!listasOk){
      // Listas no existen — mostrar setup
      localStorage.removeItem('sp_setup_done'); // limpiar flag incorrecto
      hideLoader();
      const setupEl = document.getElementById('sp-setup');
      if(setupEl){
        setupEl.style.display='flex';
        setupEl.innerHTML=`<div class="setup-card">
          <div style="font-size:40px;margin-bottom:16px">🚀</div>
          <h2>Primera configuración</h2>
          <p>Las listas de datos no existen aún en SharePoint.<br>
          Se crearán automáticamente ahora.<br><br>
          <strong>Solo toma 1-2 minutos.</strong></p>
          <button class="btn btn-primary" style="width:100%;justify-content:center;padding:14px"
            onclick="spRunSetup()">⚙️ Crear listas en SharePoint</button>
          <p style="margin-top:12px;font-size:11px;color:var(--muted)">
            Requiere permisos de administrador del sitio
          </p>
        </div>`;
      }
      return;
    }

    // 3. Listas OK — pre-cargar usuarios SP y mostrar login RelianceDesk
    localStorage.setItem('sp_setup_done','1');
    try{
      const spUsers = await spGetAll('usuarios');
      _cache.usuarios = spUsers;
    }catch(e){ console.warn('Pre-carga usuarios SP falló:', e); }
    hideLoader();
    const loginEl = document.getElementById('login-screen');
    if(loginEl) loginEl.style.display = 'flex';

  }catch(err){
    clearTimeout(safetyTimeout);
    console.error('bootApp error:', err);
    showSpError('Error al iniciar: ' + err.message);
  }
}
