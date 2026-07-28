'use strict';

const GoogleSheetsSync = (() => {
  const cfg = window.CUSTOMER_GOOGLE_CONFIG || {};
  const scopes = 'https://www.googleapis.com/auth/spreadsheets';
  const stateKey = 'customerCaptureGoogleSheets_v1';
  const appDataSheet = '_App Data';
  const dataChunkSize = 40000;
  let tokenClient = null;
  let accessToken = '';
  let tokenExpiresAt = 0;
  let spreadsheetId = cfg.spreadsheetId || localStorage.getItem(`${stateKey}:spreadsheetId`) || '';
  let syncTimer = null;
  let syncing = false;
  let pending = false;
  let originalSaveStorage = null;

  const configured = () => Boolean(cfg.clientId && !cfg.clientId.startsWith('PASTE_'));
  const enc = encodeURIComponent;
  const api = path => `https://sheets.googleapis.com/v4/${path}`;

  function status(text, type='') {
    const el = document.getElementById('cloudStatus');
    if (el) {
      el.textContent = text;
      el.dataset.state = type;
    }
    if (typeof setStatus === 'function' && text) {
      setStatus(text, type === 'error' ? 'bad' : type === 'ok' ? 'good' : '');
    }
  }

  function saveState(extra={}) {
    const old = JSON.parse(localStorage.getItem(stateKey) || '{}');
    localStorage.setItem(stateKey, JSON.stringify({...old, ...extra, savedAt:new Date().toISOString()}));
  }

  function initTokenClient() {
    if (!configured()) {
      status('Google Sheets setup required', 'warning');
      return false;
    }
    if (!window.google?.accounts?.oauth2) {
      throw new Error('Google Identity Services did not load. Check the internet connection.');
    }
    if (!tokenClient) {
      tokenClient = google.accounts.oauth2.initTokenClient({
        client_id: cfg.clientId,
        scope: scopes,
        callback: () => {},
        error_callback: err => status(`Google sign-in failed: ${err?.message || err?.type || 'Unknown error'}`, 'error')
      });
    }
    return true;
  }

  function requestToken(prompt='') {
    return new Promise((resolve, reject) => {
      if (!initTokenClient()) return reject(new Error('Google Sheets has not been configured.'));
      tokenClient.callback = response => {
        if (response.error) return reject(new Error(response.error_description || response.error));
        accessToken = response.access_token;
        tokenExpiresAt = Date.now() + Math.max(60, Number(response.expires_in || 3600) - 60) * 1000;
        resolve(accessToken);
      };
      tokenClient.requestAccessToken({prompt});
    });
  }

  async function token(interactive=true) {
    if (accessToken && Date.now() < tokenExpiresAt) return accessToken;
    if (!interactive) throw new Error('Google authorisation expired. Reconnect Google Sheets.');
    return requestToken(accessToken ? '' : 'consent');
  }

  async function request(path, options={}, interactive=true) {
    const bearer = await token(interactive);
    const response = await fetch(api(path), {
      ...options,
      headers: {
        Authorization: `Bearer ${bearer}`,
        ...(options.body ? {'Content-Type':'application/json'} : {}),
        ...(options.headers || {})
      }
    });
    if (!response.ok) {
      let message = `${response.status} ${response.statusText}`;
      try {
        const body = await response.json();
        message = body?.error?.message || message;
      } catch {}
      const error = new Error(message);
      error.status = response.status;
      throw error;
    }
    if (response.status === 204) return null;
    return response.json();
  }

  async function connect() {
    try {
      status('Connecting to Google Sheets…', 'busy');
      await requestToken('consent');
      updateButtons();
      if (!spreadsheetId) {
        status('Connected. Create a new Google Sheet or enter its ID in google-config.js.', 'ok');
        return;
      }
      await validateSpreadsheet();
      await pull(false);
    } catch (e) {
      status(`Google Sheets connection failed: ${e.message}`, 'error');
    }
  }

  function disconnect() {
    if (accessToken && window.google?.accounts?.oauth2) google.accounts.oauth2.revoke(accessToken, () => {});
    accessToken = '';
    tokenExpiresAt = 0;
    updateButtons();
    status('Google Sheets disconnected', 'warning');
  }

  async function validateSpreadsheet() {
    if (!spreadsheetId) throw new Error('No Google Sheet has been selected.');
    return request(`spreadsheets/${enc(spreadsheetId)}?fields=spreadsheetId,properties.title,sheets.properties`, {}, true);
  }

  async function createSpreadsheet() {
    try {
      if (!accessToken) await requestToken('consent');
      status('Creating the live Google Sheet…', 'busy');
      const result = await request('spreadsheets', {
        method:'POST',
        body:JSON.stringify({properties:{title:cfg.spreadsheetName || 'Customer Sales Activity Manager LIVE'}})
      }, true);
      spreadsheetId = result.spreadsheetId;
      localStorage.setItem(`${stateKey}:spreadsheetId`, spreadsheetId);
      await ensureSheets();
      await syncNow('create');
      updateButtons();
      status('Google Sheet created and populated.', 'ok');
      window.open(`https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`, '_blank', 'noopener');
    } catch (e) {
      status(`Could not create Google Sheet: ${e.message}`, 'error');
    }
  }

  function mergeRecords(localRecords, remoteRecords) {
    const map = new Map();
    [...(remoteRecords || []), ...(localRecords || [])].forEach(r => {
      if (!r?.id) return;
      const existing = map.get(r.id);
      const a = new Date(existing?.updated || existing?.created || 0).getTime();
      const b = new Date(r.updated || r.created || 0).getTime();
      if (!existing || b >= a) map.set(r.id, r);
    });
    return [...map.values()];
  }

  function a1(sheet, range='') {
    return `'${String(sheet).replace(/'/g, "''")}'${range ? `!${range}` : ''}`;
  }

  async function ensureSheets() {
    const desired = [appDataSheet, 'Dashboard', 'Customers', 'Contacts', 'Purchases', 'Activities', 'Quotes', 'Follow-ups', 'Smart Planning', 'Calendar Events'];
    const meta = await validateSpreadsheet();
    const existing = new Set((meta.sheets || []).map(s => s.properties.title));
    const requests = desired.filter(name => !existing.has(name)).map(title => ({addSheet:{properties:{title}}}));
    if (requests.length) {
      await request(`spreadsheets/${enc(spreadsheetId)}:batchUpdate`, {
        method:'POST', body:JSON.stringify({requests})
      }, true);
    }
    const refreshed = await validateSpreadsheet();
    const appSheet = (refreshed.sheets || []).find(s => s.properties.title === appDataSheet);
    if (appSheet && !appSheet.properties.hidden) {
      await request(`spreadsheets/${enc(spreadsheetId)}:batchUpdate`, {
        method:'POST', body:JSON.stringify({requests:[{updateSheetProperties:{properties:{sheetId:appSheet.properties.sheetId,hidden:true},fields:'hidden'}}]})
      }, true);
    }
  }

  async function readRemoteRecords(allowMissing=true) {
    try {
      const result = await request(`spreadsheets/${enc(spreadsheetId)}/values/${enc(a1(appDataSheet,'A1:A'))}?majorDimension=ROWS`, {}, false);
      const rows = result.values || [];
      if (!rows.length || rows[0]?.[0] !== 'CUSTOMER_CAPTURE_JSON_V1') return [];
      const json = rows.slice(2).map(r => r[0] || '').join('');
      if (!json) return [];
      const payload = JSON.parse(json);
      return Array.isArray(payload) ? payload : (payload.records || []);
    } catch (e) {
      if (allowMissing && (e.status === 400 || e.status === 404)) return [];
      throw e;
    }
  }

  function lastFollowUp(record) {
    return (record.interactions || []).map(i => i.followUpDate).filter(Boolean).sort()[0] || '';
  }

  function sheetRows(dataRecords) {
    const ctx = typeof excelContext === 'function' ? excelContext(dataRecords) : dataRecords.map(r => ({record:r,sheetName:r.details?.customerName || 'Customer'}));
    const customerRows = [['Customer ID','Customer Name','Account Number','Lead Source','Customer Category','Project / Lead Group','Lead Date','Opening Horizon','Target Open Date','Account Opened Date','Days to Open','Lost Date','Lost Reason','Main Contact','Phone','Email','Additional Contacts','Assigned Rep','Status','Monthly Spend','Last Activity','Next Follow-up','Open Quotes','Last Updated']];
    ctx.forEach(x => {
      const r=x.record,d=r.details||{};
      customerRows.push([r.id,d.customerName||'',d.accountNumber||'',d.leadSource||'',d.customerCategory||'',d.categoryGroup||'',d.leadDate||'',d.accountHorizon||'',d.targetOpenDate||'',d.accountOpenedDate||'',conversionDays(r)??'',d.lostDate||'',d.lostReason||'',d.contactPerson||'',d.phone||'',d.email||'',(r.contacts||[]).length,d.assignedRep||'',d.customerStatus||'',recordSpend(r),lastActivityDate(r),lastFollowUp(r),(r.quotes||[]).filter(isOpenQuote).length,r.updated||'']);
    });

    const contactRows = [['Customer ID','Customer','Customer Category','Project / Lead Group','Contact Name','Position / Department','Decision Role','Phone','WhatsApp','Email','Preferred Method','Notes','Active']];
    const purchaseRows = [['Customer ID','Customer','Customer Category','Project / Lead Group','Product','Category','Brand','Current Supplier','Unit Price','Monthly Usage','Unit','Monthly Spend','Notes']];
    const activityRows = [['Customer ID','Customer','Customer Category','Project / Lead Group','Date','Time','Activity Type','Contact','Subject / Purpose','Summary','Outcome','Quote Required?','Related Quote No.','Follow-up Date','Follow-up Time','Owner','Status']];
    const quoteRows = [['Customer ID','Customer','Customer Category','Project / Lead Group','Quote No.','Quote Date','Contact','Description / Products','Amount','Status','Follow-up Date','Last Followed Up','Outcome / Reason','Owner','Account Status']];
    ctx.forEach(x => {
      const r=x.record,d=r.details||{},cat=d.customerCategory||'',grp=d.categoryGroup||'';
      (r.contacts||[]).forEach(c=>contactRows.push([r.id,d.customerName||'',cat,grp,c.name||'',c.position||'',c.role||'',c.phone||'',c.whatsapp||'',c.email||'',c.preferred||'',c.notes||'',c.active||'']));
      (r.purchases||[]).forEach(p=>purchaseRows.push([r.id,d.customerName||'',cat,grp,p.product||'',p.category||'',p.brand||'',p.supplier||'',Number(p.unitPrice)||0,Number(p.monthlyUsage)||0,p.unit||'',(Number(p.unitPrice)||0)*(Number(p.monthlyUsage)||0),p.notes||'']));
      (r.interactions||[]).forEach(i=>activityRows.push([r.id,d.customerName||'',cat,grp,i.activityDate||i.date||'',i.activityTime||'',i.type||'',i.contact||'',i.subject||'',i.summary||'',i.outcome||'',i.quoteRequired||'',i.quoteNumber||'',i.followUpDate||'',i.followUpTime||'',i.owner||'',i.status||'']));
      (r.quotes||[]).forEach(q=>quoteRows.push([r.id,d.customerName||'',cat,grp,q.quoteNumber||'',q.quoteDate||'',q.contact||'',q.description||'',Number(q.amount)||0,q.status||'',q.followUpDate||'',q.lastFollowUpDate||'',q.outcome||'',q.owner||'',d.customerStatus||'']));
    });

    const followRows = [['Customer ID','Customer','Customer Category','Project / Lead Group','Due Date','Time','Days to Due','Timing','Source','Contact','Subject / Quote','Owner','Status','Last Outcome']];
    getAllFollowUps().forEach(f=>{const r=dataRecords.find(x=>x.id===f.recordId),d=r?.details||{};followRows.push([f.recordId,f.customer||'',d.customerCategory||'',d.categoryGroup||'',f.dueDate||'',f.time||'',f.days??'',categoryLabel(f.category,f.days),f.source||'',f.contact||'',f.subject||'',f.owner||'',f.status||'',f.outcome||''])});

    const planRows = [['Customer ID','Priority','Customer','Customer Category','Project / Lead Group','Due','Type','Recommended Action','Reason','Owner','Quote No.']];
    getSmartPlan().forEach(p=>{const r=dataRecords.find(x=>x.id===p.recordId),d=r?.details||{};planRows.push([p.recordId,p.priority||'',p.customer||'',d.customerCategory||'',d.categoryGroup||'',p.due||'',p.type||'',p.action||'',p.reason||'',p.owner||'',p.quoteNumber||''])});

    const eventRows = [['Customer ID','Date','Time','Customer','Customer Category','Project / Lead Group','Event Type','Activity / Source','Subject / Detail','Owner','Status']];
    calendarEvents().forEach(e=>{const r=dataRecords.find(x=>x.id===e.recordId),d=r?.details||{};eventRows.push([e.recordId,e.date||'',e.time||'',d.customerName||'',d.customerCategory||'',d.categoryGroup||'',e.kind||'',e.title||'',e.detail||'',d.assignedRep||'',e.status||''])});

    const seen = new Set();
    dataRecords.forEach(r=>(r.interactions||[]).forEach(i=>{if(i.type==='In-person Visit')seen.add(r.id)}));
    const converted=dataRecords.filter(r=>r.details?.accountOpenedDate),lost=dataRecords.filter(r=>r.details?.lostDate||r.details?.customerStatus==='Lost'),decided=converted.length+lost.length;
    const days=converted.map(conversionDays).filter(x=>x!==null&&x>=0),allQuotes=dataRecords.flatMap(r=>r.quotes||[]),won=allQuotes.filter(q=>q.status==='Accepted').length,quoteLost=allQuotes.filter(q=>['Rejected','Expired'].includes(q.status)).length;
    const dashboardRows = [
      ['SALES & CUSTOMER DASHBOARD','Value'],
      ['Last updated',new Date().toISOString()],
      ['Total Customers',dataRecords.length],
      ['Customers Seen',seen.size],
      ['Converted',converted.length],
      ['Lost',lost.length],
      ['Conversion Rate',decided ? converted.length/decided : 0],
      ['Average Days to Open',days.length ? days.reduce((a,b)=>a+b,0)/days.length : 0],
      ['Open Quotes',allQuotes.filter(isOpenQuote).length],
      ['Quote Win Rate',(won+quoteLost)?won/(won+quoteLost):0],
      ['Overdue Follow-ups',getAllFollowUps().filter(f=>f.category==='overdue').length],
      [],['CUSTOMER CATEGORY','COUNT']
    ];
    const categoryCounts={};dataRecords.forEach(r=>{const k=r.details?.customerCategory||'Uncategorised';categoryCounts[k]=(categoryCounts[k]||0)+1});Object.entries(categoryCounts).sort((a,b)=>b[1]-a[1]).forEach(x=>dashboardRows.push(x));
    dashboardRows.push([],['PROJECT / LEAD GROUP','COUNT']);const groupCounts={};dataRecords.forEach(r=>{const k=r.details?.categoryGroup;if(k)groupCounts[k]=(groupCounts[k]||0)+1});Object.entries(groupCounts).sort((a,b)=>b[1]-a[1]).forEach(x=>dashboardRows.push(x));

    return {
      'Dashboard':dashboardRows,'Customers':customerRows,'Contacts':contactRows,'Purchases':purchaseRows,
      'Activities':activityRows,'Quotes':quoteRows,'Follow-ups':followRows,'Smart Planning':planRows,'Calendar Events':eventRows
    };
  }

  async function writeAllSheets(dataRecords) {
    await ensureSheets();
    const payload = {version:6,updated:new Date().toISOString(),records:dataRecords};
    const json = JSON.stringify(payload);
    const chunks=[]; for(let i=0;i<json.length;i+=dataChunkSize) chunks.push([json.slice(i,i+dataChunkSize)]);
    const appRows=[['CUSTOMER_CAPTURE_JSON_V1'],[new Date().toISOString()],...chunks];
    const datasets=sheetRows(dataRecords);
    datasets[appDataSheet]=appRows;
    const ranges=Object.keys(datasets).map(name=>a1(name,'A:AZ'));
    await request(`spreadsheets/${enc(spreadsheetId)}/values:batchClear`, {
      method:'POST', body:JSON.stringify({ranges})
    }, true);
    const data=Object.entries(datasets).map(([name,values])=>({range:a1(name,'A1'),majorDimension:'ROWS',values}));
    await request(`spreadsheets/${enc(spreadsheetId)}/values:batchUpdate`, {
      method:'POST', body:JSON.stringify({valueInputOption:'USER_ENTERED',data})
    }, true);
  }

  async function formatSheets() {
    const meta=await validateSpreadsheet();
    const requests=[];
    (meta.sheets||[]).forEach(s=>{
      if(s.properties.title===appDataSheet)return;
      const id=s.properties.sheetId;
      requests.push({updateSheetProperties:{properties:{sheetId:id,gridProperties:{frozenRowCount:s.properties.title==='Dashboard'?1:1}},fields:'gridProperties.frozenRowCount'}});
      requests.push({repeatCell:{range:{sheetId:id,startRowIndex:0,endRowIndex:1},cell:{userEnteredFormat:{backgroundColor:{red:0.031,green:0.169,blue:0.298},textFormat:{foregroundColor:{red:1,green:1,blue:1},bold:true},wrapStrategy:'WRAP'}},fields:'userEnteredFormat'}});
      requests.push({autoResizeDimensions:{dimensions:{sheetId:id,dimension:'COLUMNS',startIndex:0,endIndex:24}}});
    });
    if(requests.length) await request(`spreadsheets/${enc(spreadsheetId)}:batchUpdate`,{method:'POST',body:JSON.stringify({requests})},true);
  }

  async function pull(confirmMerge=true) {
    if (!spreadsheetId) return status('No Google Sheet selected.', 'warning');
    if (!accessToken) return connect();
    try {
      status('Loading live Google Sheet data…', 'busy');
      const remote=await readRemoteRecords(true);
      if(remote.length){
        const merged=mergeRecords(records,remote);
        if(JSON.stringify(merged)!==JSON.stringify(records) && (!confirmMerge || confirm(`Load and merge ${remote.length} Google Sheet customer record(s)?`))){
          records=merged;records.forEach(migrateRecord);localStorage.setItem(STORAGE_KEY,JSON.stringify(records));resetForm();renderDashboard();renderFollowUps();renderPlanning();renderCalendar();switchTab('dashboard');
        }
      }
      saveState({lastPull:new Date().toISOString(),count:records.length,spreadsheetId});
      status(`Google Sheets live · ${records.length} customer record(s)`, 'ok');
    } catch(e){status(`Could not load Google Sheets data: ${e.message}`, 'error')}
  }

  async function syncNow(reason='manual') {
    if (!accessToken) { if(reason==='manual') return connect(); return; }
    if (!spreadsheetId) { if(reason==='manual') return createSpreadsheet(); return; }
    if(syncing){pending=true;return} syncing=true;
    try{
      status('Updating the live Google Sheet…','busy');
      let remote=[];try{remote=await readRemoteRecords(true)}catch{}
      const merged=mergeRecords(records,remote);
      if(JSON.stringify(merged)!==JSON.stringify(records)){records=merged;records.forEach(migrateRecord);localStorage.setItem(STORAGE_KEY,JSON.stringify(records))}
      await writeAllSheets(records);
      if(reason==='create'||reason==='manual') await formatSheets();
      saveState({lastSync:new Date().toISOString(),count:records.length,spreadsheetId});
      renderList(document.getElementById('customerSearch')?.value||'');renderDashboard();renderFollowUps();renderPlanning();renderCalendar();
      status(`Google Sheet updated · ${new Date().toLocaleTimeString('en-ZA',{hour:'2-digit',minute:'2-digit'})}`,'ok');
    }catch(e){status(`Google Sheet save failed; local copy retained: ${e.message}`,'error')}
    finally{syncing=false;if(pending){pending=false;setTimeout(()=>syncNow('queued'),400)}}
  }

  function scheduleSync(reason='change') {
    if(!cfg.autoSync||!accessToken||!spreadsheetId)return;
    clearTimeout(syncTimer);syncTimer=setTimeout(()=>syncNow(reason),Number(cfg.syncDelayMs)||1000);
  }

  function openSheet() {
    if(!spreadsheetId)return status('Create or configure a Google Sheet first.','warning');
    window.open(`https://docs.google.com/spreadsheets/d/${spreadsheetId}/edit`,'_blank','noopener');
  }

  function updateButtons(){
    const connect=document.getElementById('cloudConnectBtn'),sync=document.getElementById('cloudSyncBtn'),create=document.getElementById('cloudCreateBtn'),open=document.getElementById('cloudOpenBtn');
    if(connect)connect.textContent=accessToken?'Disconnect Google':'Connect Google';
    if(sync)sync.disabled=!accessToken||!spreadsheetId;
    if(create)create.disabled=!accessToken;
    if(open)open.disabled=!spreadsheetId;
  }

  function injectUi(){
    const actions=document.querySelector('.top-actions');if(!actions||document.getElementById('cloudConnectBtn'))return;
    const st=document.createElement('span');st.id='cloudStatus';st.style.cssText='align-self:center;font-size:11px;max-width:250px;opacity:.9';st.textContent=configured()?'Google Sheets not connected':'Google Sheets setup required';
    const connect=document.createElement('button');connect.id='cloudConnectBtn';connect.className='btn btn-light';connect.textContent='Connect Google';connect.onclick=()=>accessToken?disconnect():connectGoogle();
    const create=document.createElement('button');create.id='cloudCreateBtn';create.className='btn btn-outline';create.textContent='Create Google Sheet';create.disabled=true;create.onclick=createSpreadsheet;
    const open=document.createElement('button');open.id='cloudOpenBtn';open.className='btn btn-outline';open.textContent='Open Sheet';open.disabled=!spreadsheetId;open.onclick=openSheet;
    const sync=document.createElement('button');sync.id='cloudSyncBtn';sync.className='btn btn-outline';sync.textContent='Sync Now';sync.disabled=true;sync.onclick=()=>syncNow('manual');
    actions.prepend(sync);actions.prepend(open);actions.prepend(create);actions.prepend(connect);actions.prepend(st);
  }

  async function connectGoogle(){return connect()}

  function initialise(){
    injectUi();
    originalSaveStorage=saveStorage;
    saveStorage=function googleAwareSaveStorage(){originalSaveStorage();scheduleSync('record-change')};
    updateButtons();
  }

  return {initialise,connect,disconnect,createSpreadsheet,pull,syncNow,scheduleSync,openSheet};
})();

document.addEventListener('DOMContentLoaded',()=>GoogleSheetsSync.initialise());
