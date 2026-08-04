'use strict';

(() => {
  const NAVY = '041E42';
  const GOLD = 'EAAA00';
  const LIGHT = 'EEF2F7';
  const WHITE = 'FFFFFF';
  const TEXT = '142033';
  const MUTED = '637083';
  const RED = 'C62828';
  const GREEN = '16825D';
  const BLUE = '175CD3';

  const byId = id => document.getElementById(id);
  const safe = value => String(value ?? '').trim();
  const num = value => Number(value) || 0;
  const moneyText = value => new Intl.NumberFormat('en-ZA', {
    style: 'currency', currency: 'ZAR', maximumFractionDigits: 0
  }).format(num(value));
  const localDate = value => {
    if (!value) return '';
    const d = new Date(`${value}T00:00:00`);
    return Number.isNaN(d.getTime()) ? safe(value) : d.toLocaleDateString('en-ZA');
  };
  const dateIso = d => `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`;
  const todayIsoLocal = () => dateIso(new Date());

  function getPeriodStart(type) {
    if (typeof periodStart === 'function') return periodStart(type);
    const d = new Date();
    const y = d.getFullYear();
    const m = d.getMonth();
    if (type === 'all') return '';
    if (type === '30') { d.setDate(d.getDate() - 29); return dateIso(d); }
    if (type === 'quarter') return dateIso(new Date(y, Math.floor(m / 3) * 3, 1));
    if (type === 'year') return `${y}-01-01`;
    return dateIso(new Date(y, m, 1));
  }

  function inPeriod(value, start) {
    if (!start) return true;
    return Boolean(value && value >= start && value <= todayIsoLocal());
  }

  function spend(record) {
    if (typeof recordSpend === 'function') return recordSpend(record);
    return (record.purchases || []).reduce((sum, p) => sum + num(p.unitPrice) * num(p.monthlyUsage), 0);
  }

  function conversionTime(record) {
    if (typeof conversionDays === 'function') return conversionDays(record);
    const a = record.details?.leadDate;
    const b = record.details?.accountOpenedDate;
    if (!a || !b) return null;
    return Math.round((new Date(`${b}T00:00:00`) - new Date(`${a}T00:00:00`)) / 86400000);
  }

  function isQuoteOpen(quote) {
    if (typeof isOpenQuote === 'function') return isOpenQuote(quote);
    return !['Accepted', 'Rejected', 'Expired', 'Cancelled'].includes(quote.status);
  }

  function currentFilters() {
    const periodEl = byId('dashboardPeriod');
    const groupEl = byId('dashboardGroupFilter');
    const statusEl = byId('dashboardStatusFilter');
    return {
      period: periodEl?.value || 'month',
      periodLabel: periodEl?.selectedOptions?.[0]?.textContent || 'This Month',
      group: groupEl?.value || '',
      groupLabel: groupEl?.selectedOptions?.[0]?.textContent || 'All Project / Lead Groups',
      status: statusEl?.value || '',
      statusLabel: statusEl?.selectedOptions?.[0]?.textContent || 'All Customer Statuses'
    };
  }

  function sourceRecords(filters) {
    const all = (typeof records !== 'undefined' && Array.isArray(records)) ? records : [];
    return all.filter(r => {
      const d = r.details || {};
      return (!filters.group || d.categoryGroup === filters.group) &&
        (!filters.status || d.customerStatus === filters.status);
    });
  }

  function reportData() {
    const filters = currentFilters();
    const customers = sourceRecords(filters);
    const start = getPeriodStart(filters.period);
    const seenIds = new Set();
    customers.forEach(r => (r.interactions || []).forEach(i => {
      const date = i.activityDate || i.date || '';
      if (i.type === 'In-person Visit' && inPeriod(date, start)) seenIds.add(r.id);
    }));
    const converted = customers.filter(r => inPeriod(r.details?.accountOpenedDate, start));
    const lost = customers.filter(r => {
      const d = r.details || {};
      return (d.customerStatus === 'Lost' || d.lostDate) && inPeriod(d.lostDate || r.updated?.slice(0, 10), start);
    });
    const inProgress = customers.filter(r => ['Prospect', 'Working / In Progress', 'On Hold'].includes(r.details?.customerStatus));
    const quotes = customers.flatMap(r => (r.quotes || []).map(q => ({...q, customer:r.details?.customerName || '', group:r.details?.categoryGroup || '', customerStatus:r.details?.customerStatus || ''})));
    const periodQuotes = quotes.filter(q => inPeriod(q.quoteDate || q.followUpDate, start));
    const openQuotes = quotes.filter(isQuoteOpen);
    const wonQuotes = periodQuotes.filter(q => q.status === 'Accepted');
    const lostQuotes = periodQuotes.filter(q => ['Rejected', 'Expired'].includes(q.status));
    const decided = converted.length + lost.length;
    const days = converted.map(conversionTime).filter(v => v !== null && v >= 0);
    const planning = typeof getSmartPlan === 'function' ? getSmartPlan().filter(p => customers.some(r => r.id === p.recordId)) : [];
    return {
      filters, customers, start, seenIds, converted, lost, inProgress, quotes, periodQuotes,
      openQuotes, wonQuotes, lostQuotes, planning,
      conversionRate: decided ? converted.length / decided : 0,
      avgDays: days.length ? days.reduce((a, b) => a + b, 0) / days.length : 0,
      monthlySpend: customers.reduce((sum, r) => sum + spend(r), 0)
    };
  }

  function loadPptxGenJS() {
    if (window.PptxGenJS) return Promise.resolve(window.PptxGenJS);
    return new Promise((resolve, reject) => {
      const existing = document.querySelector('script[data-customer-pptx]');
      if (existing) {
        existing.addEventListener('load', () => resolve(window.PptxGenJS), {once:true});
        existing.addEventListener('error', () => reject(new Error('PowerPoint library could not be loaded.')), {once:true});
        return;
      }
      const script = document.createElement('script');
      script.dataset.customerPptx = 'true';
      script.src = 'https://cdn.jsdelivr.net/npm/pptxgenjs@3.12.0/dist/pptxgen.bundle.js';
      script.onload = () => window.PptxGenJS ? resolve(window.PptxGenJS) : reject(new Error('PowerPoint library did not initialise.'));
      script.onerror = () => reject(new Error('PowerPoint library could not be downloaded. Check the internet connection.'));
      document.head.appendChild(script);
    });
  }

  function addBrand(slide, title, subtitle = '') {
    slide.background = {color: WHITE};
    slide.addShape('rect', {x:0, y:0, w:13.333, h:0.55, fill:{color:NAVY}, line:{color:NAVY}});
    slide.addShape('rect', {x:0, y:0.55, w:13.333, h:0.06, fill:{color:GOLD}, line:{color:GOLD}});
    slide.addText(title, {x:0.55, y:0.75, w:8.8, h:0.45, fontFace:'Aptos Display', fontSize:24, bold:true, color:NAVY, margin:0});
    if (subtitle) slide.addText(subtitle, {x:0.57, y:1.23, w:11.9, h:0.3, fontFace:'Aptos', fontSize:10.5, color:MUTED, margin:0});
    slide.addText('Customer Sales & Activity Manager', {x:9.35, y:0.13, w:3.4, h:0.22, fontFace:'Aptos', fontSize:9, bold:true, color:WHITE, align:'right', margin:0});
    slide.addText(new Date().toLocaleDateString('en-ZA'), {x:11.4, y:7.17, w:1.35, h:0.18, fontFace:'Aptos', fontSize:8, color:MUTED, align:'right', margin:0});
  }

  function addMetric(slide, x, y, w, label, value, accent = BLUE) {
    slide.addShape('roundRect', {x, y, w, h:1.05, rectRadius:0.06, fill:{color:'F8FAFC'}, line:{color:'D8E0E9', width:1}});
    slide.addShape('rect', {x, y, w:0.08, h:1.05, fill:{color:accent}, line:{color:accent}});
    slide.addText(label.toUpperCase(), {x:x+0.2, y:y+0.16, w:w-0.3, h:0.2, fontFace:'Aptos', fontSize:9, bold:true, color:MUTED, margin:0});
    slide.addText(String(value), {x:x+0.2, y:y+0.43, w:w-0.3, h:0.38, fontFace:'Aptos Display', fontSize:22, bold:true, color:NAVY, margin:0});
  }

  function addTable(slide, rows, options = {}) {
    const data = rows.length ? rows : [[{text:'No records for the selected filters.', options:{color:MUTED, italic:true}}]];
    slide.addTable(data, {
      x:options.x ?? 0.55, y:options.y ?? 1.65, w:options.w ?? 12.2, h:options.h,
      border:{type:'solid', color:'D8E0E9', pt:0.7},
      fill:WHITE, color:TEXT, fontFace:'Aptos', fontSize:options.fontSize ?? 9,
      margin:0.06, rowH:0.28, breakLine:false,
      autoFit:false, valign:'mid',
      bold:false,
      ...options
    });
  }

  function topRows(items, mapper, max = 12) {
    return items.slice(0, max).map(mapper);
  }

  async function generatePowerPoint() {
    const button = byId('powerPointReportBtn');
    const original = button?.textContent;
    try {
      if (button) { button.disabled = true; button.textContent = 'Building PowerPoint…'; }
      const PptxGenJS = await loadPptxGenJS();
      const data = reportData();
      if (!data.customers.length && !confirm('No customers match the selected dashboard filters. Generate an empty report?')) return;
      const pptx = new PptxGenJS();
      pptx.layout = 'LAYOUT_WIDE';
      pptx.author = 'Customer Sales & Activity Manager';
      pptx.company = 'Unique Welding';
      pptx.subject = 'Customer pipeline report';
      pptx.title = 'Customer Pipeline Report';
      pptx.lang = 'en-ZA';
      pptx.theme = {headFontFace:'Aptos Display', bodyFontFace:'Aptos', lang:'en-ZA'};

      const filterText = `${data.filters.periodLabel} · ${data.filters.groupLabel} · ${data.filters.statusLabel}`;

      let slide = pptx.addSlide();
      slide.background = {color:NAVY};
      slide.addShape('rect', {x:0, y:0, w:13.333, h:7.5, fill:{color:NAVY}, line:{color:NAVY}});
      slide.addShape('rect', {x:0.65, y:1.05, w:0.15, h:4.6, fill:{color:GOLD}, line:{color:GOLD}});
      slide.addText('CUSTOMER PIPELINE REPORT', {x:1.15, y:1.25, w:10.8, h:0.7, fontFace:'Aptos Display', fontSize:30, bold:true, color:WHITE, margin:0});
      slide.addText(filterText, {x:1.17, y:2.15, w:10.6, h:0.38, fontFace:'Aptos', fontSize:15, color:'D5DEEB', margin:0});
      slide.addText(`${data.customers.length} customers · ${moneyText(data.monthlySpend)} estimated monthly spend`, {x:1.17, y:2.8, w:10.6, h:0.35, fontFace:'Aptos', fontSize:17, bold:true, color:GOLD, margin:0});
      slide.addText(`Generated ${new Date().toLocaleString('en-ZA')}`, {x:1.17, y:5.45, w:5.5, h:0.25, fontFace:'Aptos', fontSize:10, color:'B9C6D8', margin:0});

      slide = pptx.addSlide();
      addBrand(slide, 'Executive Pipeline Overview', filterText);
      addMetric(slide, 0.55, 1.7, 2.25, 'Customers Seen', data.seenIds.size, BLUE);
      addMetric(slide, 2.95, 1.7, 2.25, 'Converted', data.converted.length, GREEN);
      addMetric(slide, 5.35, 1.7, 2.25, 'Lost', data.lost.length, RED);
      addMetric(slide, 7.75, 1.7, 2.25, 'Conversion Rate', `${Math.round(data.conversionRate * 100)}%`, GOLD);
      addMetric(slide, 10.15, 1.7, 2.25, 'In Progress', data.inProgress.length, BLUE);
      addMetric(slide, 0.55, 3.0, 2.85, 'Average Days to Open', data.avgDays ? data.avgDays.toFixed(1) : '—', GOLD);
      addMetric(slide, 3.55, 3.0, 2.85, 'Open Quotes', data.openQuotes.length, GOLD);
      addMetric(slide, 6.55, 3.0, 2.85, 'Open Quote Value', moneyText(data.openQuotes.reduce((s,q)=>s+num(q.amount),0)), BLUE);
      addMetric(slide, 9.55, 3.0, 2.85, 'Monthly Purchase Value', moneyText(data.monthlySpend), GREEN);
      slide.addText('Management focus', {x:0.58, y:4.55, w:2.5, h:0.3, fontSize:15, bold:true, color:NAVY, margin:0});
      const focus = [];
      if (data.planning.filter(p => p.priority === 'High').length) focus.push(`${data.planning.filter(p=>p.priority==='High').length} high-priority next actions require attention.`);
      if (data.openQuotes.length) focus.push(`${data.openQuotes.length} open quotations require customer decisions or follow-up.`);
      if (data.inProgress.length) focus.push(`${data.inProgress.length} customers remain in the active development pipeline.`);
      if (!focus.length) focus.push('No urgent pipeline exceptions were identified for the selected filters.');
      slide.addText(focus.map(text => ({text, options:{bullet:{indent:14}, hanging:3, breakLine:true}})), {x:0.7, y:4.95, w:11.7, h:1.3, fontFace:'Aptos', fontSize:14, color:TEXT, margin:0.04, breakLine:true});

      slide = pptx.addSlide();
      addBrand(slide, 'Customer Pipeline by Status', filterText);
      const statusOrder = ['Prospect','Working / In Progress','Active','Dormant','On Hold','Lost'];
      const statusCounts = statusOrder.map(status => data.customers.filter(r => r.details?.customerStatus === status).length);
      slide.addChart(pptx.ChartType.bar, [{name:'Customers', labels:statusOrder, values:statusCounts}], {
        x:0.75, y:1.7, w:7.2, h:4.9, catAxisLabelFontSize:11, valAxisLabelFontSize:10,
        showLegend:false, showTitle:false, showValue:true, showCatName:false,
        chartColors:[NAVY], gridLine:{color:'D8E0E9', width:1},
        valAxisMinVal:0, showValAxisTitle:false, showCatAxisTitle:false, border:{color:'D8E0E9', pt:1}
      });
      const pipelineRows = [['Status','Customers','Share']].concat(statusOrder.map((s,i)=>[s,statusCounts[i],data.customers.length ? `${Math.round(statusCounts[i]/data.customers.length*100)}%` : '0%']));
      addTable(slide, pipelineRows, {x:8.3, y:1.7, w:4.25, fontSize:10, rowH:0.4, colW:[2.2,0.9,0.9]});

      slide = pptx.addSlide();
      addBrand(slide, 'Customers Seen', `${filterText} · In-person visits`);
      const seenCustomers = data.customers.filter(r => data.seenIds.has(r.id));
      const seenRows = [['Customer','Project / Lead Group','Status','Last Visit','Outcome','Owner']].concat(topRows(seenCustomers, r => {
        const visits = (r.interactions || []).filter(i => i.type === 'In-person Visit' && inPeriod(i.activityDate || i.date, data.start)).sort((a,b)=>safe(b.activityDate||b.date).localeCompare(safe(a.activityDate||a.date)));
        const v = visits[0] || {};
        return [r.details?.customerName || '', r.details?.categoryGroup || '', r.details?.customerStatus || '', localDate(v.activityDate || v.date), v.outcome || v.summary || '', v.owner || r.details?.assignedRep || ''];
      }, 16));
      addTable(slide, seenRows, {x:0.45, y:1.65, w:12.45, fontSize:8.5, colW:[2.2,2.2,1.25,1.1,4.0,1.45]});

      slide = pptx.addSlide();
      addBrand(slide, 'Customers In Progress', filterText);
      const progressSorted = [...data.inProgress].sort((a,b)=>spend(b)-spend(a));
      const progressRows = [['Customer','Project / Lead Group','Status','Target Open','Monthly Value','Open Quotes','Next Action']].concat(topRows(progressSorted, r => {
        const actions = data.planning.filter(p => p.recordId === r.id).sort((a,b)=>safe(a.due).localeCompare(safe(b.due)));
        return [r.details?.customerName || '', r.details?.categoryGroup || '', r.details?.customerStatus || '', localDate(r.details?.targetOpenDate), moneyText(spend(r)), (r.quotes||[]).filter(isQuoteOpen).length, actions[0]?.action || 'No next action recorded'];
      }, 16));
      addTable(slide, progressRows, {x:0.35, y:1.65, w:12.65, fontSize:8.1, colW:[2.05,2.1,1.35,1.05,1.25,0.9,3.65]});

      slide = pptx.addSlide();
      addBrand(slide, 'Lost Customers and Reasons', filterText);
      const lostRows = [['Customer','Project / Lead Group','Lost Date','Reason / Competitor','Last Owner']].concat(topRows(data.lost, r => [r.details?.customerName || '', r.details?.categoryGroup || '', localDate(r.details?.lostDate), r.details?.lostReason || 'Not recorded', r.details?.assignedRep || ''], 16));
      addTable(slide, lostRows, {x:0.45, y:1.65, w:12.4, fontSize:8.8, colW:[2.25,2.25,1.15,5.0,1.55]});

      slide = pptx.addSlide();
      addBrand(slide, 'Quotation Pipeline', filterText);
      addMetric(slide, 0.55, 1.65, 2.35, 'Open Quotes', data.openQuotes.length, GOLD);
      addMetric(slide, 3.05, 1.65, 2.35, 'Open Value', moneyText(data.openQuotes.reduce((s,q)=>s+num(q.amount),0)), BLUE);
      addMetric(slide, 5.55, 1.65, 2.35, 'Accepted', data.wonQuotes.length, GREEN);
      addMetric(slide, 8.05, 1.65, 2.35, 'Rejected / Expired', data.lostQuotes.length, RED);
      addMetric(slide, 10.55, 1.65, 2.25, 'Win Rate', `${(data.wonQuotes.length+data.lostQuotes.length)?Math.round(data.wonQuotes.length/(data.wonQuotes.length+data.lostQuotes.length)*100):0}%`, GOLD);
      const quoteRows = [['Customer','Quote No.','Description','Amount','Status','Follow-up','Outcome']].concat(topRows([...data.openQuotes].sort((a,b)=>num(b.amount)-num(a.amount)), q => [q.customer,q.quoteNumber||'',q.description||'',moneyText(q.amount),q.status||'',localDate(q.followUpDate),q.outcome||''], 12));
      addTable(slide, quoteRows, {x:0.4, y:3.05, w:12.55, fontSize:8.1, colW:[1.95,1.25,3.35,1.25,1.05,1.1,2.6]});

      slide = pptx.addSlide();
      addBrand(slide, 'Project / Lead Group Breakdown', filterText);
      const groupMap = {};
      data.customers.forEach(r => { const key = r.details?.categoryGroup || 'Unassigned'; groupMap[key] = (groupMap[key] || 0) + 1; });
      const groups = Object.entries(groupMap).sort((a,b)=>b[1]-a[1]).slice(0,12);
      slide.addChart(pptx.ChartType.bar, [{name:'Customers', labels:groups.map(x=>x[0]), values:groups.map(x=>x[1])}], {
        x:0.7, y:1.65, w:7.2, h:5.1, catAxisLabelFontSize:10, valAxisLabelFontSize:10,
        showLegend:false, showValue:true, chartColors:[GOLD], gridLine:{color:'D8E0E9', width:1}, border:{color:'D8E0E9', pt:1}
      });
      const groupRows = [['Project / Lead Group','Customers','In Progress','Converted','Lost']].concat(groups.map(([g,count])=>[
        g, count,
        data.inProgress.filter(r=>(r.details?.categoryGroup||'Unassigned')===g).length,
        data.converted.filter(r=>(r.details?.categoryGroup||'Unassigned')===g).length,
        data.lost.filter(r=>(r.details?.categoryGroup||'Unassigned')===g).length
      ]));
      addTable(slide, groupRows, {x:8.2, y:1.65, w:4.45, fontSize:8.3, colW:[2.2,0.6,0.7,0.7,0.6]});

      slide = pptx.addSlide();
      addBrand(slide, 'Priority Actions — Next 7 Days', filterText);
      const horizon = new Date(); horizon.setDate(horizon.getDate() + 7); const horizonIso = dateIso(horizon);
      const priorities = data.planning.filter(p => !p.due || p.due <= horizonIso).sort((a,b) => {
        const pr = {High:0,Medium:1,Low:2}; return (pr[a.priority]??9)-(pr[b.priority]??9) || safe(a.due).localeCompare(safe(b.due));
      });
      const actionRows = [['Priority','Customer','Due','Type','Recommended Action','Reason','Owner','Quote No.']].concat(topRows(priorities, p => [p.priority||'',p.customer||'',localDate(p.due),p.type||'',p.action||'',p.reason||'',p.owner||'',p.quoteNumber||''], 16));
      addTable(slide, actionRows, {x:0.25, y:1.6, w:12.85, fontSize:7.7, colW:[0.75,1.65,0.9,0.85,2.55,3.65,1.35,1.15]});

      const filename = `Customer_Pipeline_Report_${todayIsoLocal()}${data.filters.group ? `_${data.filters.group.replace(/[^a-z0-9]+/gi,'_').slice(0,35)}` : ''}.pptx`;
      await pptx.writeFile({fileName:filename});
      if (typeof setStatus === 'function') setStatus(`PowerPoint report created: ${filename}`, 'good');
    } catch (error) {
      console.error(error);
      if (typeof setStatus === 'function') setStatus(`PowerPoint report failed: ${error.message}`, 'bad');
      else alert(`PowerPoint report failed: ${error.message}`);
    } finally {
      if (button) { button.disabled = false; button.textContent = original || 'Generate PowerPoint'; }
    }
  }

  function injectButton() {
    if (byId('powerPointReportBtn')) return;
    const controls = byId('dashboardPeriod')?.closest('.section-controls');
    if (!controls) return;
    const button = document.createElement('button');
    button.id = 'powerPointReportBtn';
    button.type = 'button';
    button.className = 'btn btn-primary btn-small';
    button.textContent = 'Generate PowerPoint';
    button.title = 'Create an editable PowerPoint report using the current dashboard filters';
    button.addEventListener('click', generatePowerPoint);
    controls.appendChild(button);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', injectButton);
  else injectButton();
  window.CustomerPowerPointReport = {generate:generatePowerPoint};
})();
