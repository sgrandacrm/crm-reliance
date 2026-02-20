// ═══════════════════════════════════════════════════════════════
//  SHAREPOINT LAYER — CRM Reliance
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
    cotizaciones:  'CRM_Cotizaciones',
    cierres:       'CRM_Cierres',
    usuarios:      'CRM_Usuarios',
    historial:     'CRM_ImportHistorial',
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
  historial:    null,
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
    let url = `sites/${_siteId}/lists/${listName}/items?$expand=fields&$top=999`;

    while(url){
      const r = await spGraph(url);
      items = items.concat(r.value || []);
      url = r['@odata.nextLink'] || null;
    }

    // Extraer campos individuales
    const result = items.map(item => {
      const f = item.fields;
      const obj = { _spId: item.id };
      Object.keys(f).forEach(k => {
        if(k.startsWith('_') || k==='ContentType' || k==='Attachments') return;
        let val = f[k];
        if(typeof val === 'string' && (val.startsWith('[') || val.startsWith('{'))){
          try{ val = JSON.parse(val); }catch(e){}
        }
        // Revertir prefijo crm_ a nombre original
        const realKey = k.startsWith('crm_') ? k.slice(4) : k;
        obj[realKey] = val;
      });
      if(f.Title && !obj.nombre) obj.nombre = f.Title;
      if(f.Title && listKey==='cotizaciones') obj.codigo = f.Title;
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
    const fields = spToFields(listKey, data);
    const r = await spGraph(
      `sites/${_siteId}/lists/${listName}/items`,
      'POST',
      { fields }
    );
    _cache[listKey] = null; // invalidar cache
    updateSpStatus('online','● SharePoint');
    return r.id;
  }catch(e){
    updateSpStatus('error','⚠ Error al guardar');
    console.error('spCreate error:', listKey, e);
    showToast('Error guardando en SharePoint: ' + e.message, 'error');
    return null;
  }
}

// ── Actualizar ítem existente ────────────────────────────────
async function spUpdate(listKey, spId, data){
  if(!_spReady){ console.warn('SP not ready'); return false; }

  const listName = SP_CONFIG.lists[listKey];
  updateSpStatus('syncing','⟳ Actualizando...');

  try{
    const fields = spToFields(listKey, data);
    await spGraph(
      `sites/${_siteId}/lists/${listName}/items/${spId}/fields`,
      'PATCH',
      fields
    );
    _cache[listKey] = null;
    updateSpStatus('online','● SharePoint');
    return true;
  }catch(e){
    updateSpStatus('error','⚠ Error al actualizar');
    console.error('spUpdate error:', listKey, e);
    showToast('Error actualizando SharePoint: ' + e.message, 'error');
    return false;
  }
}

// ── Eliminar ítem ────────────────────────────────────────────
async function spDelete(listKey, spId){
  if(!_spReady) return false;
  const listName = SP_CONFIG.lists[listKey];
  try{
    await spGraph(`sites/${_siteId}/lists/${listName}/items/${spId}`, 'DELETE');
    _cache[listKey] = null;
    return true;
  }catch(e){
    console.error('spDelete error:', e);
    return false;
  }
}

// ── Convertir objeto JS → campos de lista SP ─────────────────
// Campos reservados de SharePoint que no se pueden usar directamente
const SP_RESERVED = new Set(['id','ID','title','author','editor','created','modified',
  'ContentType','Attachments','_UIVersionString','owshiddenversion','FileDirRef',
  'FileLeafRef','FileRef','FSObjType','UniqueId','GUID','WorkflowVersion','MetaInfo']);

function spToFields(listKey, data){
  const fields = {};
  Object.keys(data).forEach(k => {
    if(k.startsWith('_')) return;
    // Renombrar campos reservados
    let fieldName = k;
    if(SP_RESERVED.has(k)) fieldName = 'crm_' + k;
    
    let val = data[k];
    if(val !== null && val !== undefined && typeof val === 'object'){
      val = JSON.stringify(val);
    }
    if(k==='nombre' && listKey==='clientes')     { fields['Title'] = String(val||'').substring(0,255); return; }
    if(k==='codigo' && listKey==='cotizaciones') { fields['Title'] = String(val||'').substring(0,255); return; }
    if(k==='nombre' && listKey==='usuarios')     { fields['Title'] = String(val||'').substring(0,255); return; }
    if(k==='polizaNueva' && listKey==='cierres') { fields['Title'] = String(val||'').substring(0,255); }
    if(typeof val === 'string') val = val.substring(0, 3999);
    fields[fieldName] = val;
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
      {name:'desde',type:'DateTime'},{name:'hasta',type:'DateTime'},
      {name:'celular',type:'Text'},{name:'correo',type:'Text'},
      {name:'nota',type:'Note'},{name:'ultimoContacto',type:'DateTime'},
      {name:'factura',type:'Text'},{name:'poliza',type:'Text'},
      {name:'polizaNueva',type:'Text'},{name:'aseguradoraAnterior',type:'Text'},
      {name:'obs',type:'Text'},{name:'color',type:'Text'},
      {name:'motor',type:'Text'},{name:'chasis',type:'Text'},
      {name:'dep',type:'Number'},{name:'tasa',type:'Number'},
      {name:'axavd',type:'Text'},{name:'formaPago',type:'Text'},
      {name:'historialWa',type:'note'},{name:'crm_id',type:'text'},
    ],
    CRM_Cotizaciones: [
      {name:'codigo',type:'Text'},{name:'version',type:'Number'},
      {name:'fecha',type:'DateTime'},{name:'ejecutivo',type:'Text'},
      {name:'clienteNombre',type:'Text'},{name:'clienteCI',type:'Text'},
      {name:'clienteId',type:'Text'},{name:'ciudad',type:'Text'},
      {name:'vehiculo',type:'Text'},{name:'placa',type:'Text'},
      {name:'va',type:'Number'},{name:'desde',type:'DateTime'},
      {name:'estado',type:'Text'},{name:'asegElegida',type:'Text'},
      {name:'resultados',type:'Note'},{name:'aseguradoras',type:'Note'},
      {name:'obsAcept',type:'Note'},{name:'fechaAcept',type:'DateTime'},
      {name:'reemplazadaPor',type:'Text'},{name:'autoSust',type:'Boolean'},
      {name:'cuotasTc',type:'Number'},{name:'cuotasDeb',type:'Number'},
      {name:'spLocalId',type:'Text'},
    ],
    CRM_Cierres: [
      {name:'clienteNombre',type:'Text'},{name:'aseguradora',type:'Text'},
      {name:'primaTotal',type:'Number'},{name:'primaNeta',type:'Number'},
      {name:'vigDesde',type:'DateTime'},{name:'vigHasta',type:'DateTime'},
      {name:'formaPago',type:'Text'},{name:'facturaAseg',type:'Text'},
      {name:'ejecutivo',type:'Text'},{name:'fechaRegistro',type:'DateTime'},
      {name:'observacion',type:'Note'},{name:'axavd',type:'Text'},
      {name:'cuenta',type:'Text'},{name:'spLocalId',type:'Text'},
    ],
    CRM_Usuarios: [
      {name:'userId',type:'Text'},{name:'rol',type:'Text'},
      {name:'email',type:'Text'},{name:'activo',type:'Boolean'},
    ],
    CRM_ImportHistorial: [
      {name:'fecha',type:'DateTime'},{name:'archivo',type:'Text'},
      {name:'registros',type:'Number'},{name:'usuario',type:'Text'},
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
    ['reliance_import_historial','historial'],
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
  try{
    const el = document.getElementById('sp-status');
    if(!el) return;
    el.className = `sp-status ${state}`;
    el.textContent = text;
  }catch(e){}
}

// ── LOGIN UI ─────────────────────────────────────────────────
function showSpLogin(){
  // Ocultar loader con pequeño delay para que el DOM esté listo
  setTimeout(hideLoader, 100);
  const setupEl = document.getElementById('sp-setup');
  if(!setupEl){ setTimeout(showSpLogin, 200); return; }
  setupEl.style.display = 'flex';
  setupEl.innerHTML = `
    <div class="setup-card">
      <div style="font-size:40px;margin-bottom:16px">🛡</div>
      <h2>CRM Reliance</h2>
      <p>Inicia sesión con tu cuenta Microsoft para acceder al CRM compartido en SharePoint.</p>
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
      {name:'crm_id',type:'text'},{name:'polizaNueva',type:'text'},
      {name:'aseguradoraAnterior',type:'text'},{name:'historialWa',type:'note'},
    ],
    CRM_Cotizaciones: [
      {name:'codigo',type:'text'},{name:'version',type:'number'},
      {name:'fecha',type:'text'},{name:'ejecutivo',type:'text'},
      {name:'clienteNombre',type:'text'},{name:'clienteCI',type:'text'},
      {name:'clienteId',type:'text'},{name:'ciudad',type:'text'},
      {name:'vehiculo',type:'text'},{name:'placa',type:'text'},
      {name:'va',type:'number'},{name:'desde',type:'text'},
      {name:'estado',type:'text'},{name:'asegElegida',type:'text'},
      {name:'resultados',type:'note'},{name:'aseguradoras',type:'note'},
      {name:'obsAcept',type:'note'},{name:'fechaAcept',type:'text'},
      {name:'reemplazadaPor',type:'text'},{name:'crm_id',type:'text'},
    ],
    CRM_Cierres: [
      {name:'clienteNombre',type:'text'},{name:'aseguradora',type:'text'},
      {name:'primaTotal',type:'number'},{name:'primaNeta',type:'number'},
      {name:'vigDesde',type:'text'},{name:'vigHasta',type:'text'},
      {name:'formaPago',type:'text'},{name:'facturaAseg',type:'text'},
      {name:'ejecutivo',type:'text'},{name:'fechaRegistro',type:'text'},
      {name:'observacion',type:'note'},{name:'axavd',type:'text'},
      {name:'crm_id',type:'text'},{name:'polizaNueva',type:'text'},
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
      if(!colsDone){
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
        localStorage.setItem('sp_cols_done','1');
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

    // 3. Listas OK — cargar app
    localStorage.setItem('sp_setup_done','1');
    hideLoader();
    await initApp();

  }catch(err){
    clearTimeout(safetyTimeout);
    console.error('bootApp error:', err);
    showSpError('Error al iniciar: ' + err.message);
  }
}
