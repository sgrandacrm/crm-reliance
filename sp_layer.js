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
  updateSpStatus('syncing', '⟳ Conectando...');

  const msalConfig = {
    auth: {
      clientId:    SP_CONFIG.clientId,
      authority:   `https://login.microsoftonline.com/${SP_CONFIG.tenantId}`,
      redirectUri: window.location.href.split('?')[0],
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

    // Extraer campos y parsear JSON donde corresponda
    const result = items.map(item => {
      const f = item.fields;
      const obj = { _spId: item.id, _spEtag: item['@odata.etag'] };
      Object.keys(f).forEach(k => {
        if(k.startsWith('_') || k==='ContentType' || k==='Attachments') return;
        let val = f[k];
        // Intentar parsear campos JSON (resultados de cotizaciones, etc)
        if(typeof val === 'string' && (val.startsWith('[') || val.startsWith('{'))){
          try{ val = JSON.parse(val); }catch(e){}
        }
        obj[k] = val;
      });
      // Mapear Title → nombre para clientes
      if(f.Title && !obj.nombre) obj.nombre = f.Title;
      if(f.Title && listKey==='cotizaciones') obj.codigo = f.Title;
      if(f.Title && listKey==='cierres') obj.polizaNueva = f.Title;
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
function spToFields(listKey, data){
  const fields = {};
  Object.keys(data).forEach(k => {
    if(k.startsWith('_')) return; // ignorar metadatos internos
    let val = data[k];
    // Serializar arrays/objetos como JSON string
    if(val !== null && typeof val === 'object'){
      val = JSON.stringify(val);
    }
    // Mapeo de campos especiales
    if(k==='nombre' && listKey==='clientes') { fields['Title'] = val; return; }
    if(k==='codigo' && listKey==='cotizaciones') { fields['Title'] = val; return; }
    if(k==='polizaNueva' && listKey==='cierres') { fields['Title'] = val; return; }
    if(k==='nombre' && listKey==='usuarios') { fields['Title'] = val; return; }
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
      {name:'desde',type:'DateTime'},{name:'hasta',type:'DateTime'},
      {name:'celular',type:'Text'},{name:'correo',type:'Text'},
      {name:'nota',type:'Note'},{name:'ultimoContacto',type:'DateTime'},
      {name:'factura',type:'Text'},{name:'poliza',type:'Text'},
      {name:'polizaNueva',type:'Text'},{name:'aseguradoraAnterior',type:'Text'},
      {name:'obs',type:'Text'},{name:'color',type:'Text'},
      {name:'motor',type:'Text'},{name:'chasis',type:'Text'},
      {name:'dep',type:'Number'},{name:'tasa',type:'Number'},
      {name:'axavd',type:'Text'},{name:'formaPago',type:'Text'},
      {name:'historialWa',type:'Note'},{name:'id',type:'Text'},
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
  const el = document.getElementById('sp-status');
  if(!el) return;
  el.className = `sp-status ${state}`;
  el.textContent = text;
}

// ── LOGIN UI ─────────────────────────────────────────────────
function showSpLogin(){
  hideLoader();
  const setupEl = document.getElementById('sp-setup');
  if(!setupEl) return;
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
  const l = document.getElementById('sp-loader');
  if(l) l.style.display='none';
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
  showToast('✅ SharePoint configurado correctamente','success');
  // Recargar app
  await initApp();
}

// ── ARRANQUE PRINCIPAL ───────────────────────────────────────
async function bootApp(){
  // 1. Inicializar MSAL y autenticar
  const ok = await spInit();
  if(!ok) return; // muestra login o error

  // 2. Ver si ya se hizo el setup
  const setupDone = localStorage.getItem('sp_setup_done');
  if(!setupDone){
    // Primera vez: mostrar botón de setup
    hideLoader();
    const setupEl = document.getElementById('sp-setup');
    if(setupEl){
      setupEl.style.display='flex';
      setupEl.innerHTML=`<div class="setup-card">
        <div style="font-size:40px;margin-bottom:16px">🚀</div>
        <h2>Primera configuración</h2>
        <p>Esta es la primera vez que se usa el CRM en este SharePoint.<br>
        Se crearán las listas necesarias automáticamente.<br><br>
        <strong>Solo toma 1-2 minutos.</strong></p>
        <button class="btn btn-primary" style="width:100%;justify-content:center;padding:14px"
          onclick="spRunSetup()">⚙️ Configurar SharePoint</button>
        <p style="margin-top:12px;font-size:11px;color:var(--muted)">
          Requiere permisos de administrador del sitio
        </p>
      </div>`;
    }
    return;
  }

  // 3. Setup ya hecho — cargar app normalmente
  hideLoader();
  await initApp();
}
