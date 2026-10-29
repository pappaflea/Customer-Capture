'use strict';
(() => {
  const byId=id=>document.getElementById(id);
  const customerTabs=new Set(['details','contacts','purchases','quotes','interactions','timeline']);
  const activityTabs=new Set(['daily','quickvisit','calendar','followups']);
  const reportTabs=new Set(['planning','datatools']);

  function css(){
    const s=document.createElement('style');
    s.textContent=
      'body.simple-ui .tabs{display:none!important}'+
      '.simple-nav,.simple-subnav{display:flex;gap:7px;flex-wrap:wrap;margin:0 0 14px;padding:8px;background:#fff;border:1px solid var(--line);border-radius:10px;box-shadow:var(--shadow)}'+
      '.simple-subnav{display:none;background:#eef2f7;box-shadow:none;margin-top:-5px}.simple-subnav.show{display:flex}'+
      '.simple-nav button,.simple-subnav button{border:0;background:transparent;color:var(--navy);padding:9px 12px;border-radius:7px;font-weight:700;cursor:pointer}'+
      '.simple-nav button.active,.simple-subnav button.active{background:var(--navy);color:#fff}'+
      '.simple-subnav .danger{background:#fff1f0;color:var(--danger)}'+
      '.simple-actions{display:grid;grid-template-columns:repeat(4,minmax(130px,1fr));gap:10px;margin-bottom:14px}'+
      '.simple-actions button{padding:14px;border:1px solid var(--line);border-radius:9px;background:#fff;text-align:left;cursor:pointer}'+
      '.simple-actions strong{display:block;color:var(--navy);font-size:14px}.simple-actions span{display:block;color:var(--muted);font-size:11px;margin-top:3px}'+
      '.simple-advanced{display:none!important}body.simple-ui.show-more .simple-advanced,body.simple-ui .simple-force{display:block!important}'+
      '.simple-quick{background:var(--gold)!important;color:var(--navy)!important}'+
      '@media(max-width:760px){.simple-nav{display:grid;grid-template-columns:1fr 1fr}.simple-nav button:last-child{grid-column:1/-1}.simple-actions{grid-template-columns:1fr}}';
    document.head.appendChild(s);
  }

  function area(tab){
    if(tab==='dashboard')return'dashboard';
    if(tab==='customers'||customerTabs.has(tab))return'customers';
    if(tab==='pipeline')return'pipeline';
    if(activityTabs.has(tab))return'activity';
    if(reportTabs.has(tab))return'reports';
    return'dashboard';
  }

  function navButton(label,target,areaName){
    const b=document.createElement('button');
    b.type='button';b.textContent=label;b.dataset.area=areaName||'';b.dataset.target=target;
    b.onclick=()=>switchTab(target);
    return b;
  }

  function buildNav(){
    const tabs=document.querySelector('.tabs');
    if(!tabs||byId('simpleNav'))return;
    const nav=document.createElement('div');nav.id='simpleNav';nav.className='simple-nav';
    nav.append(
      navButton('Dashboard','dashboard','dashboard'),
      navButton('Customers','customers','customers'),
      navButton('Pipeline','pipeline','pipeline'),
      navButton('Activity','daily','activity'),
      navButton('Reports / Admin','planning','reports')
    );
    tabs.parentElement.insertBefore(nav,tabs);

    const c=document.createElement('div');c.id='simpleCustomerSub';c.className='simple-subnav';
    [['Overview','details'],['Contacts','contacts'],['Purchases','purchases'],['Quotes','quotes'],['Activity','interactions'],['Timeline','timeline']].forEach(x=>c.appendChild(navButton(x[0],x[1])));
    const out=document.createElement('button');out.type='button';out.textContent='Update Outcome';out.className='danger';out.onclick=updateOutcome;c.appendChild(out);

    const a=document.createElement('div');a.id='simpleActivitySub';a.className='simple-subnav';
    [['Today','daily'],['Quick Capture','quickvisit'],['Calendar','calendar'],['Follow-ups','followups']].forEach(x=>a.appendChild(navButton(x[0],x[1])));

    const r=document.createElement('div');r.id='simpleReportsSub';r.className='simple-subnav';
    [['Next Actions','planning'],['Import / Data Quality','datatools']].forEach(x=>r.appendChild(navButton(x[0],x[1])));
    [['PowerPoint','powerPointReportBtn'],['Excel','excelBtn'],['Backup JSON','backupBtn']].forEach(x=>{
      const b=document.createElement('button');b.type='button';b.textContent=x[0];b.onclick=()=>byId(x[1])?.click();r.appendChild(b);
    });

    tabs.parentElement.insertBefore(c,tabs);
    tabs.parentElement.insertBefore(a,tabs);
    tabs.parentElement.insertBefore(r,tabs);
  }

  function buildActivityActions(){
    const p=byId('panel-daily');if(!p||byId('simpleActions'))return;
    const d=document.createElement('div');d.id='simpleActions';d.className='simple-actions';
    [['Visit','In-person Visit','Record visit + next action'],['Call','Phone Call','Capture call + follow-up'],['WhatsApp / Email','WhatsApp','Log customer contact'],['Quote','Quote','Open customer quotation screen']].forEach(x=>{
      const b=document.createElement('button');b.type='button';b.innerHTML='<strong>'+x[0]+'</strong><span>'+x[2]+'</span>';b.onclick=()=>quickAction(x[1]);d.appendChild(b);
    });
    p.insertBefore(d,p.firstChild);
  }

  function quickAction(type){
    if(type==='Quote'){
      if(typeof currentId!=='undefined'&&currentId){openRecord(currentId,'quotes');return;}
      switchTab('customers');return;
    }
    if(typeof updateQuickCustomerOptions==='function')updateQuickCustomerOptions();
    if(typeof currentId!=='undefined'&&currentId&&byId('quickCustomer'))byId('quickCustomer').value=currentId;
    if(byId('quickType'))byId('quickType').value=type;
    switchTab('quickvisit');
  }

  function addQuickButton(){
    const actions=document.querySelector('.top-actions');if(!actions||byId('simpleQuickBtn'))return;
    const b=document.createElement('button');b.id='simpleQuickBtn';b.type='button';b.className='btn simple-quick';b.textContent='+ Quick Capture';b.onclick=()=>quickAction('In-person Visit');actions.insertBefore(b,actions.firstChild);
  }

  function simplifyDetails(){
    const hidden=['tradingName','registrationNumber','vatNumber','industry','accountOpenedDate','lostDate','lostReason','leadSourceDetail','customerCategory','pipelineStage','opportunityValue','winProbability','expectedCloseDate','accountHorizon','targetOpenDate','branch','website','address','contactRole','whatsapp','preferredContact'];
    hidden.forEach(id=>byId(id)?.closest('.field')?.classList.add('simple-advanced'));
    const head=byId('panel-details')?.querySelector('.card .section-title');
    if(head&&!byId('simpleMoreBtn')){
      const b=document.createElement('button');b.id='simpleMoreBtn';b.type='button';b.className='btn btn-secondary btn-small';b.textContent='Show More Details';
      b.onclick=()=>{const on=document.body.classList.toggle('show-more');b.textContent=on?'Hide Extra Details':'Show More Details';conditionalFields();};head.appendChild(b);
    }
    byId('customerStatus')?.addEventListener('change',conditionalFields);
    conditionalFields();
  }

  function conditionalFields(){
    document.querySelectorAll('.simple-force').forEach(x=>x.classList.remove('simple-force'));
    if(document.body.classList.contains('show-more'))return;
    const st=byId('customerStatus')?.value||'';
    if(st==='Lost')['lostDate','lostReason'].forEach(id=>byId(id)?.closest('.field')?.classList.add('simple-force'));
    if(st==='Active')byId('accountOpenedDate')?.closest('.field')?.classList.add('simple-force');
  }

  function updateOutcome(){
    if(typeof currentId==='undefined'||!currentId){alert('Open a customer first.');return;}
    const c=prompt('Customer outcome:\\n1 = Won / Active\\n2 = Lost\\n3 = On Hold\\n4 = Working / In Progress','1');
    if(!c)return;
    if(c==='1'){
      byId('customerStatus').value='Active';byId('pipelineStage').value='Won';
      if(!byId('accountOpenedDate').value)byId('accountOpenedDate').value=todayIso();
    }else if(c==='2'){
      const reason=prompt('Lost reason / competitor:','')||'';
      if(!reason&&!confirm('Mark Lost without a reason?'))return;
      byId('customerStatus').value='Lost';byId('pipelineStage').value='Lost';
      if(!byId('lostDate').value)byId('lostDate').value=todayIso();
      byId('lostReason').value=reason;
    }else if(c==='3'){
      byId('customerStatus').value='On Hold';
    }else if(c==='4'){
      byId('customerStatus').value='Working / In Progress';
      if(['New Lead','Contacted'].includes(byId('pipelineStage').value))byId('pipelineStage').value='Working / In Progress';
    }else return;
    conditionalFields();
    if(typeof saveCustomer==='function')saveCustomer();
  }

  function updateNav(tab){
    const ar=area(tab);
    document.querySelectorAll('#simpleNav [data-area]').forEach(b=>b.classList.toggle('active',b.dataset.area===ar));
    const c=byId('simpleCustomerSub'),a=byId('simpleActivitySub'),r=byId('simpleReportsSub');
    c?.classList.toggle('show',customerTabs.has(tab)&&typeof currentId!=='undefined'&&!!currentId);
    a?.classList.toggle('show',activityTabs.has(tab));
    r?.classList.toggle('show',reportTabs.has(tab));
    [c,a,r].forEach(n=>n?.querySelectorAll('[data-target]').forEach(b=>b.classList.toggle('active',b.dataset.target===tab)));
    if(ar==='reports'){byId('recordKpis')?.classList.add('hidden');if(byId('recordTitle'))byId('recordTitle').textContent='Reports / Admin';if(byId('recordSub'))byId('recordSub').textContent='Management reports, next actions, exports and data-quality tools.';}
    if(tab==='daily'){if(byId('recordTitle'))byId('recordTitle').textContent='Activity';if(byId('recordSub'))byId('recordSub').textContent='Capture customer contact and manage today’s next actions.';}
    conditionalFields();
  }

  function wrapSwitch(){
    if(window.__simpleUiWrapped)return;window.__simpleUiWrapped=true;
    const base=switchTab;
    switchTab=function(name){base(name);updateNav(name);};
  }

  function init(){
    if(window.__simpleUiLoaded)return;window.__simpleUiLoaded=true;
    document.body.classList.add('simple-ui');
    css();buildNav();buildActivityActions();addQuickButton();simplifyDetails();wrapSwitch();
    const active=document.querySelector('.panel.active')?.id?.replace('panel-','')||'dashboard';
    updateNav(active);
  }

  if(document.readyState==='loading')document.addEventListener('DOMContentLoaded',()=>setTimeout(init,0));
  else setTimeout(init,0);
})();