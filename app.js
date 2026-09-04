(() => {
  'use strict';

  const API_URL = 'https://script.google.com/macros/s/AKfycbycm7-F4_Go9kNuFz34hTC1yWvDxDkQe1WkfwXAJgQXQaOi9fiWibtjqt9h63lRJf0A/exec';
  const FALLBACK = window.STEN_FALLBACK_DATA;
  const Engine = window.StenEngine;
  const STORAGE_KEY = 'sten-master-web-v38-state';
  const OLD_STORAGE_KEY = 'sten-master-web-v35-state';
  const MAX_CERT_BYTES = 8 * 1024 * 1024;

  let model = clone(FALLBACK);
  let sourceMode = 'fallback';
  let state = loadState(model.defaults);
  let results = [];
  let historyData = null;
  let historySelection = new Set();
  let admin = {
    connected: false,
    token: sessionStorage.getItem('sten_admin_token') || '',
    actor: sessionStorage.getItem('sten_admin_actor') || '',
    audit: [], control: [], strategic: [], view: 'control'
  };

  const $ = sel => document.querySelector(sel);
  const $$ = sel => [...document.querySelectorAll(sel)];
  const esc = v => String(v ?? '').replace(/[&<>'"]/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;',"'":'&#39;','"':'&quot;'}[ch]));
  const fmtN = (v, d=2) => new Intl.NumberFormat('ru-RU',{maximumFractionDigits:d,minimumFractionDigits:0}).format(Number(v)||0);
  const fmtRub = (v, d=0) => new Intl.NumberFormat('ru-RU',{style:'currency',currency:'RUB',maximumFractionDigits:d,minimumFractionDigits:d}).format(Number(v)||0);
  const fmtSigned = (v, unit='') => `${Number(v)>0?'+':''}${fmtN(v,2)}${unit}`;
  const finishOptions = ['Без отделки','Предчистовая','Чистовая'];
  const catMeta = {
    ext:{label:'Наружные стены',icon:'🧱',unit:'м³',priceUnit:'₽/м³'},
    inter:{label:'Межквартирные стены',icon:'🏢',unit:'м³',priceUnit:'₽/м³'},
    part:{label:'Перегородки',icon:'🚪',unit:'м²',priceUnit:'₽/м²'}
  };
  const chartPalette = ['#0070F3','#059669','#D97706','#7C3AED','#DC2626','#0891B2','#4F46E5','#65A30D','#DB2777','#475569','#EA580C','#0D9488'];

  function clone(x){ return JSON.parse(JSON.stringify(x)); }
  function defaultState(defaults){
    return {
      inputs:{city:defaults.city,saleArea:defaults.saleArea,totalArea:defaults.totalArea,perimExt:defaults.perimExt,perimInter:defaults.perimInter,perimPart:defaults.perimPart,floorH:defaults.floorH,floors:defaults.floors,salePrice:defaults.salePrice},
      project:clone(defaults.project), scenarios:clone(defaults.scenarios), activeTab:'overview'
    };
  }
  function loadState(defaults){
    try{
      const raw = localStorage.getItem(STORAGE_KEY) || localStorage.getItem(OLD_STORAGE_KEY) || 'null';
      const saved=JSON.parse(raw);
      return saved && saved.inputs && saved.project && Array.isArray(saved.scenarios) ? saved : defaultState(defaults);
    }catch{return defaultState(defaults);}
  }
  function saveState(){ localStorage.setItem(STORAGE_KEY,JSON.stringify(state)); }

  function validateState(){
    for(const cat of ['ext','inter','part']){
      const list=model.constructions?.[cat]||[];
      const names=list.map(x=>x.name);
      const fallbackProject = names.includes(model.defaults?.project?.[cat]?.material) ? clone(model.defaults.project[cat]) : {material:names[0]||'',finish:'Без отделки'};
      if(!state.project?.[cat] || !names.includes(state.project[cat].material)) state.project[cat]=fallbackProject;
      if(!finishOptions.includes(state.project[cat].finish)) state.project[cat].finish='Без отделки';
      state.scenarios = Array.isArray(state.scenarios) ? state.scenarios.slice(0,3) : [];
      while(state.scenarios.length<3) state.scenarios.push(clone(model.defaults?.scenarios?.[state.scenarios.length] || model.defaults?.project || {}));
      state.scenarios.forEach((sc,i)=>{
        if(!sc[cat] || !names.includes(sc[cat].material)) {
          const d=model.defaults?.scenarios?.[i]?.[cat];
          sc[cat]=d && names.includes(d.material)?clone(d):clone(fallbackProject);
        }
        if(!finishOptions.includes(sc[cat].finish)) sc[cat].finish='Без отделки';
      });
    }
  }

  function renderAll(){
    validateState();
    results=Engine.calcAll(state,model);
    renderInputs(); renderProject(); renderOverview();
    renderCategory('ext'); renderCategory('inter'); renderCategory('part');
    renderPricing(); renderHistoryPicker(); renderCertificates(); renderAdmin(); renderSystem();
    setActiveTab(state.activeTab||'overview',false);
  }

  function renderInputs(){
    const cities=Object.keys(model.delivery?.distances||{});
    const fields=[
      ['city','Город объекта','select',cities],['saleArea','Продаваемая площадь, м²','number'],['totalArea','Общая площадь секции, м²','number'],
      ['perimExt','Периметр наружных, м','number'],['perimInter','Периметр межквартирных, м','number'],['perimPart','Периметр перегородок, м','number'],
      ['floorH','Высота этажа, м','number'],['floors','Количество этажей','number'],['salePrice','Стоимость 1 м² продаваемой площади, ₽','number']
    ];
    $('#inputsGrid').innerHTML=fields.map(([key,label,type,opts])=>{
      if(type==='select') return `<div class="field"><label>${esc(label)}</label><select data-input="${key}">${opts.map(x=>`<option ${state.inputs[key]===x?'selected':''}>${esc(x)}</option>`).join('')}</select></div>`;
      const step=key==='floors'?'1':key==='salePrice'?'1000':'0.01';
      return `<div class="field"><label>${esc(label)}</label><input data-input="${key}" type="number" step="${step}" value="${esc(state.inputs[key])}"></div>`;
    }).join('');
    $$('[data-input]').forEach(el=>el.addEventListener(el.tagName==='SELECT'?'change':'input',e=>{
      const key=e.currentTarget.dataset.input;
      state.inputs[key]=e.currentTarget.tagName==='SELECT'?e.currentTarget.value:Number(e.currentTarget.value||0);
      saveState(); renderDynamicOnly();
    }));
  }

  function renderProject(){
    $('#projectEditor').innerHTML=['ext','inter','part'].map(cat=>{
      const p=state.project[cat];
      const mats=model.constructions[cat]||[];
      const mat=Engine.getConstruction(model,cat,p.material);
      const thick=Engine.getEffectiveThickness(mat,p.finish);
      return `<div class="project-row" data-project-row="${cat}">
        <div class="project-label">${catMeta[cat].icon} ${esc(catMeta[cat].label)}</div>
        <select data-project-material="${cat}">${mats.map(x=>`<option ${x.name===p.material?'selected':''}>${esc(x.name)}</option>`).join('')}</select>
        <select data-project-finish="${cat}">${finishOptions.map(x=>`<option ${x===p.finish?'selected':''}>${esc(x)}</option>`).join('')}</select>
        <div class="project-thick">${fmtN(thick,0)} мм</div>
      </div>`;
    }).join('');
    $$('[data-project-material]').forEach(el=>el.onchange=e=>{const cat=e.currentTarget.dataset.projectMaterial;state.project[cat].material=e.currentTarget.value;saveState();renderAll();});
    $$('[data-project-finish]').forEach(el=>el.onchange=e=>{const cat=e.currentTarget.dataset.projectFinish;state.project[cat].finish=e.currentTarget.value;saveState();renderAll();});
  }

  function renderDynamicOnly(){
    results=Engine.calcAll(state,model);
    renderProject(); renderOverview(); renderCategory('ext'); renderCategory('inter'); renderCategory('part'); renderSystemMeta();
  }

  function renderOverview(){
    const minFull=Math.min(...results.map(x=>x.fullTotal));
    $('#scenarioSummary').innerHTML=results.map((r,i)=>{
      const best=Math.abs(r.fullTotal-minFull)<1e-6;
      return `<article class="scenario-card ${best?'best':''}">${best?'<span class="best-tag">Мин. прямые затраты</span>':''}
        <h3 class="scenario-title">Сценарий ${i+1}</h3>
        <div class="category-mini">${['ext','inter','part'].map(cat=>`<div class="mini-row"><span>${catMeta[cat].icon}</span><span class="mini-name" title="${esc(state.scenarios[i][cat].material)}">${esc(state.scenarios[i][cat].material)}</span><span class="mini-cost">${fmtRub(r.categories[cat].total)}</span></div>`).join('')}</div>
        <div class="metric-list">
          <div class="metric"><span class="metric-label">ИТОГО выбранных решений</span><span class="metric-value primary">${fmtRub(r.changedTotal)}</span></div>
          <div class="metric"><span class="metric-label">Δ продаваемой площади объекта</span><span class="metric-value ${r.deltaSaleObj>0?'good':r.deltaSaleObj<0?'bad':''}">${fmtSigned(r.deltaSaleObj,' м²')}</span></div>
          <div class="metric"><span class="metric-label">Δ стоимости продаваемой площади</span><span class="metric-value ${r.deltaSaleValue>0?'good':r.deltaSaleValue<0?'bad':''}">${r.deltaSaleValue>0?'+':''}${fmtRub(r.deltaSaleValue)}</span></div>
          <div class="metric"><span class="metric-label">Справочно: доставка</span><span class="metric-value">${fmtRub(['ext','inter','part'].reduce((a,c)=>a+r.categories[c].deliveryCost,0))}</span></div>
        </div>
        <div class="scenario-total"><div class="metric"><span class="metric-label">Выбранные решения / м² продаваемой</span><span class="metric-value primary">${fmtRub(r.changedPerSaleArea)}</span></div></div>
      </article>`;
    }).join('');

    const rows=[
      ['Δ площади под стенами на этаж','deltaFootprint','м²'],['Δ продаваемой площади на этаж','deltaSaleFloor','м²'],['Δ продаваемой площади объекта','deltaSaleObj','м²'],['Δ стоимости продаваемой площади','deltaSaleValue','₽'],
      ['Наружные — стоимость выбранного решения','changedExt','₽'],['Межквартирные — стоимость выбранного решения','changedInter','₽'],['Перегородки — стоимость выбранного решения','changedPart','₽'],['ИТОГО ВЫБРАННЫХ РЕШЕНИЙ','changedTotal','₽'],['Цена выбранных / м² общей площади','changedPerTotalArea','₽/м²'],['Цена выбранных / м² продаваемой площади','changedPerSaleArea','₽/м²']
    ];
    $('#summaryTable').innerHTML=`<table class="data-table"><thead><tr><th>Показатель</th><th>Сценарий 1</th><th>Сценарий 2</th><th>Сценарий 3</th></tr></thead><tbody>${rows.map(([label,key,unit])=>`<tr><td>${esc(label)}</td>${results.map(r=>{
      const val=key==='changedExt'?r.changedByCategory.ext:key==='changedInter'?r.changedByCategory.inter:key==='changedPart'?r.changedByCategory.part:r[key];
      if(unit==='₽'||unit==='₽/м²') return `<td class="num">${fmtRub(val)}</td>`;
      return `<td class="num">${fmtSigned(val,' '+unit)}</td>`;
    }).join('')}</tr>`).join('')}</tbody></table>`;
  }

  function renderCategory(cat){
    const container=$(`#${cat}Variants`); if(!container) return;
    const totals=results.map(r=>r.categories[cat].total), min=Math.min(...totals);
    container.innerHTML=state.scenarios.map((sc,i)=>renderVariantCard(cat,i,results[i].categories[cat],Math.abs(totals[i]-min)<1e-6)).join('');
    $$(`[data-scenario-material^="${cat}:"]`).forEach(el=>el.onchange=e=>{const [,idx]=e.currentTarget.dataset.scenarioMaterial.split(':');state.scenarios[Number(idx)][cat].material=e.currentTarget.value;saveState();renderDynamicOnly();});
    $$(`[data-scenario-finish^="${cat}:"]`).forEach(el=>el.onchange=e=>{const [,idx]=e.currentTarget.dataset.scenarioFinish.split(':');state.scenarios[Number(idx)][cat].finish=e.currentTarget.value;saveState();renderDynamicOnly();});
    renderCompare(cat);
  }

  function renderVariantCard(cat,i,c,best){
    const selection=state.scenarios[i][cat], mats=model.constructions[cat]||[], mat=Engine.getConstruction(model,cat,selection.material);
    const strategic=(mat?.components||[]).filter(x=>x.purchaseGroup==='Стратегические материалы' && Number(x.sheetContribution)!==0).length;
    const certs=certificatesForConstruction(mat);
    const reference=cat==='ext'?c.insulation:c.noise;
    const badge=reference?`<span class="badge ${/не проходит/i.test(String(reference))?'fail':/нет данных|требует/i.test(String(reference))?'warn':'pass'}">${cat==='ext'?'утеплитель':'звук'}: ${esc(reference)}</span>`:'';
    return `<article class="variant-card ${best?'best':''}">${best?'<span class="best-tag">минимум</span>':''}
      <div class="material-glyph">${esc(initials(selection.material))}</div>
      <h3 class="variant-title">Сценарий ${i+1}</h3>
      <div class="variant-controls">
        <div class="field"><label>Материал</label><select class="variant-select" data-scenario-material="${cat}:${i}">${mats.map(x=>`<option ${x.name===selection.material?'selected':''}>${esc(x.name)}</option>`).join('')}</select></div>
        <div class="field"><label>Тип отделки</label><select class="variant-select" data-scenario-finish="${cat}:${i}">${finishOptions.map(x=>`<option ${x===selection.finish?'selected':''}>${esc(x)}</option>`).join('')}</select></div>
      </div>
      <div class="badges"><span class="badge">📍 ${esc(c.city||'—')}</span>${strategic?`<span class="badge strategic">★ стратегических: ${strategic}</span>`:''}${certs.length?`<span class="badge pass">📄 документов: ${certs.length}</span>`:''}${badge}</div>
      <div class="metric-list">
        <div class="metric"><span class="metric-label">Толщина кладки / с отделкой</span><span class="metric-value">${fmtN(c.thick,0)} / ${fmtN(c.thickFin,0)} мм</span></div>
        <div class="metric"><span class="metric-label">Цена конструкции</span><span class="metric-value primary">${fmtRub(c.unitPrice,2)} / ${catMeta[cat].unit}</span></div>
        <div class="metric"><span class="metric-label">В том числе работы</span><span class="metric-value">${fmtRub(c.workUnit,2)}</span></div>
        <div class="metric"><span class="metric-label">В том числе материалы</span><span class="metric-value">${fmtRub(c.materialUnit,2)}</span></div>
        <div class="metric"><span class="metric-label">${cat==='part'?'Площадь кладки/этаж':'Объём кладки/этаж'}</span><span class="metric-value">${fmtN(cat==='part'?c.wallAreaFloor:c.volumeFloor,2)} ${catMeta[cat].unit}</span></div>
        <div class="metric"><span class="metric-label">Масса / этаж</span><span class="metric-value">${fmtN(c.massFloor,2)} т</span></div>
        <div class="metric"><span class="metric-label">Справочно: доставка</span><span class="metric-value">${c.trucks} маш. · ${fmtRub(c.deliveryCost)}</span></div>
        <div class="metric"><span class="metric-label">Отделка</span><span class="metric-value">${fmtRub(c.finishCost)}</span></div>
        <div class="metric"><span class="metric-label">Работы + материалы на объект</span><span class="metric-value">${fmtRub(c.costObj)}</span></div>
        <div class="metric"><span class="metric-label">ИТОГО без доставки</span><span class="metric-value primary">${fmtRub(c.total)}</span></div>
      </div>
      ${certs.length?`<details class="details"><summary>📄 Сертификаты и протоколы · ${certs.length}</summary>${renderCertificateList(certs,false)}</details>`:''}
      ${renderComponents(mat)}
    </article>`;
  }

  function renderComponents(mat){
    const comps=(mat?.components||[]).filter(c=>Number(c.sheetContribution)!==0);
    if(!comps.length) return '';
    return `<details class="details"><summary>Состав конструкции · ${comps.length} поз.</summary><div class="components"><table class="data-table"><thead><tr><th>Код</th><th>Тип</th><th>Наименование</th><th>Вклад</th><th>Закупка</th></tr></thead><tbody>${comps.map(c=>`<tr><td>${esc(c.code)}</td><td>${esc(c.type)}</td><td>${esc(c.name)}</td><td class="num">${fmtRub(c.sheetContribution,2)}</td><td>${c.purchaseGroup==='Стратегические материалы'?'<span class="badge strategic">Стратегический</span>':esc(c.purchaseGroup||'')}</td></tr>`).join('')}</tbody></table></div></details>`;
  }

  function renderCompare(cat){
    const target=$(`#${cat}Compare`), items=[...(model.constructions[cat]||[])].sort((a,b)=>b.sheetPrice-a.sheetPrice);
    target.innerHTML=`<table class="data-table"><thead><tr><th>Материал</th><th>Город</th><th>Толщина</th><th>Работы</th><th>Материалы</th><th>Цена конструкции</th><th>${cat==='ext'?'Утеплитель':'Звукоизоляция'}</th><th>Документы</th></tr></thead><tbody>${items.map(x=>{
      const certs=certificatesForConstruction(x);
      return `<tr><td>${esc(x.name)}</td><td>${esc(x.city)}</td><td class="num">${fmtN(x.thick,0)} мм</td><td class="num">${fmtRub(x.workCostSheet,2)}</td><td class="num">${fmtRub(x.materialCostSheet,2)}</td><td class="num">${fmtRub(x.sheetPrice,2)}</td><td>${esc(x.reference??'')}</td><td>${certs.length?`<button class="link-btn" data-open-cert-tab="${esc(certs[0].code)}">📄 ${certs.length}</button>`:'—'}</td></tr>`;
    }).join('')}</tbody></table>`;
    $$('[data-open-cert-tab]').forEach(b=>b.onclick=()=>{setActiveTab('certificates');$('#certificateSearch').value=b.dataset.openCertTab;renderCertificates();});
  }

  function renderPricing(){
    const q=($('#priceSearch')?.value||'').trim().toLowerCase(), type=$('#priceType')?.value||'all', purchase=$('#pricePurchase')?.value||'all';
    let rows=[];
    if(type!=='materials') rows.push(...(model.works||[]).map(x=>({...x,_kind:'works',purchaseGroup:''})));
    if(type!=='works') rows.push(...(model.materials||[]).map(x=>({...x,_kind:'materials'})));
    rows=rows.filter(x=>(!q || `${x.code} ${x.name} ${x.technicalGroup}`.toLowerCase().includes(q)) && (purchase==='all'||x.purchaseGroup===purchase));
    const techGroups=[...new Set((model.materials||[]).map(x=>x.technicalGroup).filter(Boolean))].sort();
    const workGroups=[...new Set((model.works||[]).map(x=>x.technicalGroup).filter(Boolean))].sort();
    $('#pricingTable').innerHTML=`<table class="data-table"><thead><tr><th>Код</th><th>Наименование</th><th>Ед.</th><th>${type==='works'?'Цена':'Цена с НДС и доставкой'}</th><th>Группа</th><th>Закупка</th><th>Где используется</th>${admin.connected?'<th></th>':''}</tr></thead><tbody>${rows.map(x=>{
      const strategic=x.purchaseGroup==='Стратегические материалы';
      const groups=x._kind==='works'?workGroups:techGroups;
      return `<tr class="${strategic?'strategic-row':''}" data-price-row="${esc(x.code)}" data-kind="${x._kind}">
        <td>${esc(x.code)}</td><td>${esc(x.name)}</td><td>${esc(x.unit)}</td>
        <td>${admin.connected?`<input class="editable price-input" data-field="price" type="number" step="0.01" value="${esc(x.price)}">`:`<span class="num">${fmtRub(x.price,2)}</span>`}</td>
        <td>${admin.connected?`<select class="editable group-select" data-field="technicalGroup">${groups.map(g=>`<option ${g===x.technicalGroup?'selected':''}>${esc(g)}</option>`).join('')}</select>`:esc(x.technicalGroup)}</td>
        <td>${x._kind==='materials'?(admin.connected?`<select class="editable group-select" data-field="purchaseGroup"><option ${x.purchaseGroup==='Обычные материалы'?'selected':''}>Обычные материалы</option><option ${strategic?'selected':''}>Стратегические материалы</option></select>`:(strategic?'<span class="badge strategic">Стратегический</span>':'Обычный')):'—'}</td>
        <td title="${esc(x.usage||'')}">${esc(shortUsage(x.usage))}</td>
        ${admin.connected?`<td><button class="btn btn-outline save-row" data-save-ref="${esc(x.code)}">Сохранить</button></td>`:''}
      </tr>`;
    }).join('')}</tbody></table>`;
    $$('[data-save-ref]').forEach(btn=>btn.onclick=()=>saveReferenceRow(btn.dataset.saveRef));
  }

  async function saveReferenceRow(code){
    if(!admin.connected){goAdmin('Подключите администратора для изменения расценок.');return;}
    const tr=document.querySelector(`[data-price-row="${cssEscape(code)}"]`); if(!tr) return;
    const kind=tr.dataset.kind, sheet=kind==='works'?'Работы':'Материалы';
    const source=(kind==='works'?model.works:model.materials).find(x=>x.code===code); if(!source) return;
    const changes=[];
    tr.querySelectorAll('[data-field]').forEach(el=>{
      const f=el.dataset.field, val=f==='price'?Number(el.value):el.value;
      const src=f==='technicalGroup'?source.technicalGroup:f==='purchaseGroup'?source.purchaseGroup:source.price;
      if(String(val)!==String(src)) changes.push({field:f,value:val});
    });
    if(!changes.length){toast('Изменений нет.');return;}
    setBusy(true);
    try{
      for(const ch of changes){
        const resp=await apiPost({action:'updateReference',token:admin.token,actor:admin.actor,sheet,code,field:ch.field,value:ch.value});
        if(!resp.ok) throw new Error(resp.error||'Ошибка обновления');
        if(resp.data) applyRemoteModel(resp.data);
        if(resp.audit) admin.audit=resp.audit;
      }
      toast(`Сохранено: ${code}`,'success'); renderAll();
    }catch(err){toast(err.message,'error');}
    finally{setBusy(false);}
  }

  function renderHistoryPicker(){
    const picker=$('#historyMaterialPicker'); if(!picker) return;
    const q=($('#historySearch')?.value||'').trim().toLowerCase();
    const mats=(model.materials||[]).filter(x=>!q || `${x.code} ${x.name} ${x.technicalGroup}`.toLowerCase().includes(q)).sort((a,b)=>a.name.localeCompare(b.name,'ru'));
    if(!historySelection.size){
      (model.materials||[]).filter(x=>x.technicalGroup==='Стеновые материалы' && !/архив|не применяется/i.test(x.calculatorStatus||'')).slice(0,3).forEach(x=>historySelection.add(x.code));
    }
    picker.innerHTML=mats.map(x=>`<label class="material-choice"><input type="checkbox" data-history-code="${esc(x.code)}" ${historySelection.has(x.code)?'checked':''}><span><strong>${esc(x.code)} · ${esc(x.name)}</strong><span>${esc(x.technicalGroup||'—')} · ${esc(x.unit||'—')} · ${fmtRub(x.price,2)}</span></span></label>`).join('');
    $$('[data-history-code]').forEach(ch=>ch.onchange=()=>{if(ch.checked)historySelection.add(ch.dataset.historyCode);else historySelection.delete(ch.dataset.historyCode);});
  }

  async function loadPriceHistory(){
    const codes=[...historySelection].slice(0,30);
    if(!codes.length){toast('Выберите хотя бы один материал.','error');return;}
    setBusy(true);
    try{
      const resp=await apiPost({action:'priceHistory',codes});
      if(!resp.ok) throw new Error(resp.error||'История цен недоступна');
      historyData=resp.data; renderPriceCharts(); toast('История цен загружена.','success');
    }catch(err){
      historyData=buildFallbackHistory(codes); renderPriceCharts(); toast('API истории недоступен. Показаны данные из локального снимка.','error');
    }finally{setBusy(false);}
  }

  function buildFallbackHistory(codes){
    const mats=codes.map(c=>(model.materials||[]).find(x=>x.code===c)).filter(Boolean);
    const events=[];
    mats.forEach(m=>{
      if(m.previousPrice!=null && m.changedAt){events.push({date:m.changedAt,code:m.code,before:Number(m.previousPrice)||0,after:Number(m.price)||0,field:'Цена'});}
    });
    return {generatedAt:new Date().toISOString(),materials:mats.map(m=>({code:m.code,name:m.name,unit:m.unit,currentPrice:m.price,technicalGroup:m.technicalGroup,purchaseGroup:m.purchaseGroup})),events};
  }

  function renderPriceCharts(){
    const target=$('#priceCharts'); if(!target) return;
    if(!historyData || !historyData.materials?.length){target.innerHTML='<div class="empty-state">Нет данных для построения.</div>';return;}
    const grouping=$('#historyGrouping').value;
    const groups=new Map();
    historyData.materials.forEach(m=>{
      const key=grouping==='technical'?(m.technicalGroup||'Без технической группы'):grouping==='purchase'?(m.purchaseGroup||'Без группы закупки'):grouping==='unit'?(m.unit||'Без единицы'):'Выбранные материалы';
      if(!groups.has(key))groups.set(key,[]); groups.get(key).push(m);
    });
    target.innerHTML=[...groups.entries()].map(([key,mats],idx)=>`<section class="chart-card"><h3>${esc(key)}</h3><div class="chart-meta">${mats.length} материал(а/ов) · линии показаны в исходных единицах цены</div><div class="chart-canvas-wrap"><canvas class="chart-canvas" id="priceChart${idx}"></canvas></div><div class="chart-legend">${mats.map((m,j)=>`<span class="legend-item" style="color:${chartPalette[j%chartPalette.length]}"><span class="legend-swatch"></span>${esc(m.code)} · ${esc(m.unit)}</span>`).join('')}</div></section>`).join('');
    [...groups.entries()].forEach(([,mats],idx)=>drawPriceChart($(`#priceChart${idx}`),mats,historyData.events||[],historyData.generatedAt));
  }

  function drawPriceChart(canvas,mats,events,generatedAt){
    if(!canvas) return;
    const rect=canvas.parentElement.getBoundingClientRect(), dpr=Math.max(1,window.devicePixelRatio||1);
    const width=Math.max(500,rect.width),height=Math.max(260,rect.height);
    canvas.width=width*dpr;canvas.height=height*dpr;const ctx=canvas.getContext('2d');ctx.scale(dpr,dpr);
    const series=mats.map(m=>({m,points:priceSeries(m,events,generatedAt)}));
    const all=series.flatMap(s=>s.points), pad={l:76,r:18,t:18,b:42};
    let xmin=Math.min(...all.map(p=>p.t)),xmax=Math.max(...all.map(p=>p.t)),ymin=Math.min(...all.map(p=>p.v)),ymax=Math.max(...all.map(p=>p.v));
    if(!Number.isFinite(xmin)||!Number.isFinite(xmax)){xmin=Date.now()-86400000;xmax=Date.now();}
    if(xmin===xmax){xmin-=43200000;xmax+=43200000;}
    if(!Number.isFinite(ymin)||!Number.isFinite(ymax)){ymin=0;ymax=1;}
    if(ymin===ymax){ymin=Math.max(0,ymin*.9);ymax=ymax*1.1+1;}
    const yr=(ymax-ymin)||1;ymin=Math.max(0,ymin-yr*.08);ymax+=yr*.08;
    const x=t=>pad.l+(t-xmin)/(xmax-xmin)*(width-pad.l-pad.r), y=v=>pad.t+(ymax-v)/(ymax-ymin)*(height-pad.t-pad.b);
    ctx.clearRect(0,0,width,height);ctx.font='10px Golos Text, sans-serif';ctx.fillStyle='#64748B';ctx.strokeStyle='#E2E8F0';ctx.lineWidth=1;
    for(let i=0;i<=5;i++){const yy=pad.t+i*(height-pad.t-pad.b)/5;ctx.beginPath();ctx.moveTo(pad.l,yy);ctx.lineTo(width-pad.r,yy);ctx.stroke();const val=ymax-(ymax-ymin)*i/5;ctx.fillText(fmtN(val,0),8,yy+3);}
    for(let i=0;i<=4;i++){const xx=pad.l+i*(width-pad.l-pad.r)/4;const t=xmin+(xmax-xmin)*i/4;ctx.fillText(new Date(t).toLocaleDateString('ru-RU',{day:'2-digit',month:'2-digit',year:'2-digit'}),Math.min(xx-22,width-60),height-14);}
    series.forEach((s,idx)=>{ctx.strokeStyle=chartPalette[idx%chartPalette.length];ctx.fillStyle=ctx.strokeStyle;ctx.lineWidth=2;ctx.beginPath();s.points.forEach((p,j)=>{const xx=x(p.t),yy=y(p.v);if(j===0)ctx.moveTo(xx,yy);else ctx.lineTo(xx,yy);});ctx.stroke();s.points.forEach(p=>{ctx.beginPath();ctx.arc(x(p.t),y(p.v),3,0,Math.PI*2);ctx.fill();});});
  }

  function priceSeries(mat,events,generatedAt){
    const ev=events.filter(e=>e.code===mat.code).sort((a,b)=>new Date(a.date)-new Date(b.date));
    if(!ev.length)return [{t:new Date(generatedAt||Date.now()).getTime(),v:Number(mat.currentPrice)||0}];
    const pts=[];const firstT=new Date(ev[0].date).getTime()||Date.now();pts.push({t:firstT-1000,v:Number(ev[0].before)||0});
    ev.forEach(e=>pts.push({t:new Date(e.date).getTime()||Date.now(),v:Number(e.after)||0}));
    const current=Number(mat.currentPrice)||0,last=pts[pts.length-1];if(Math.abs(last.v-current)>1e-9)pts.push({t:new Date(generatedAt||Date.now()).getTime(),v:current});
    return pts;
  }

  function activeWallMaterials(){
    return (model.materials||[]).filter(x=>x.technicalGroup==='Стеновые материалы' && !/архив|не применяется/i.test(x.calculatorStatus||'')).sort((a,b)=>a.name.localeCompare(b.name,'ru'));
  }
  function wallMaterialCodesForConstruction(mat){
    const refs=new Map((model.materials||[]).map(x=>[x.code,x]));
    return [...new Set((mat?.components||[]).filter(c=>c.type==='Материал' && Number(c.sheetContribution)!==0 && refs.get(c.code)?.technicalGroup==='Стеновые материалы').map(c=>c.code))];
  }
  function certificatesForConstruction(mat){
    const codes=new Set(wallMaterialCodesForConstruction(mat));
    return (model.certificates||[]).filter(c=>codes.has(c.code));
  }

  function renderCertificates(){
    const target=$('#certificateTable'); if(!target) return;
    const q=($('#certificateSearch')?.value||'').trim().toLowerCase();
    const certs=model.certificates||[];
    const mats=activeWallMaterials().filter(m=>!q || `${m.code} ${m.name}`.toLowerCase().includes(q));
    target.innerHTML=`<table class="data-table"><thead><tr><th>Код</th><th>Стеновой материал</th><th>Ед.</th><th>Цена</th><th>Сертификаты / протоколы</th></tr></thead><tbody>${mats.map(m=>{const docs=certs.filter(c=>c.code===m.code);return `<tr><td>${esc(m.code)}</td><td>${esc(m.name)}</td><td>${esc(m.unit)}</td><td class="num">${fmtRub(m.price,2)}</td><td>${docs.length?renderCertificateList(docs,admin.connected):'<span class="cert-none">Нет загруженных документов</span>'}</td></tr>`;}).join('')}</tbody></table>`;
  }
  function renderCertificateList(docs,withDelete){
    return `<div class="cert-list">${docs.map(c=>`<div class="cert-link"><button type="button" data-cert-download="${esc(c.id)}">📄 ${esc(c.name||c.fileName||'Документ')}</button><span class="cert-meta">${fmtBytes(c.size)}${withDelete?` · <button type="button" class="delete-cert" data-cert-delete="${esc(c.id)}">удалить</button>`:''}</span></div>`).join('')}</div>`;
  }
  async function downloadCertificate(id){
    setBusy(true);
    try{const resp=await apiPost({action:'downloadCertificate',certificateId:id});if(!resp.ok)throw new Error(resp.error||'Файл недоступен');openBase64File(resp.certificate.base64,resp.certificate.mime,resp.certificate.fileName,true);}catch(err){toast(err.message,'error');}finally{setBusy(false);}
  }
  async function deleteCertificate(id){
    if(!admin.connected){goAdmin('Удаление сертификата требует администратора.');return;}
    if(!confirm('Удалить сертификат из общего хранилища?'))return;
    setBusy(true);try{const resp=await apiPost({action:'deleteCertificate',token:admin.token,actor:admin.actor,certificateId:id});if(!resp.ok)throw new Error(resp.error||'Ошибка удаления');if(resp.data)applyRemoteModel(resp.data);else if(resp.certificates)model.certificates=resp.certificates;renderCertificates();toast('Сертификат удалён.','success');}catch(err){toast(err.message,'error');}finally{setBusy(false);}
  }

  function renderAdmin(){
    const ws=$('#adminWorkspace'); if(!ws)return;
    ws.classList.toggle('hidden',!admin.connected);$('#adminLogoutBtn').classList.toggle('hidden',!admin.connected);$('#adminLoginBtn').classList.toggle('hidden',admin.connected);
    $('#adminStatusText').textContent=admin.connected?`подключён · ${admin.actor||'администратор'}`:'не подключён';
    $('#adminHeaderBtn').textContent=admin.connected?'⚙ Администратор подключён':'⚙ Администрирование';
    if($('#adminToken'))$('#adminToken').value=admin.token;if($('#adminActor'))$('#adminActor').value=admin.actor;
    const groups=[...new Set((model.materials||[]).map(x=>x.technicalGroup).filter(Boolean))].sort();$('#techGroupList').innerHTML=groups.map(g=>`<option value="${esc(g)}"></option>`).join('');
    $('#certificateMaterialSelect').innerHTML=activeWallMaterials().map(m=>`<option value="${esc(m.code)}">${esc(m.code)} · ${esc(m.name)}</option>`).join('');
  }

  async function adminLogin(){
    admin.token=$('#adminToken').value.trim();admin.actor=$('#adminActor').value.trim();
    if(!admin.token){toast('Введите Admin key.','error');return;}
    setBusy(true);
    try{
      const resp=await apiPost({action:'adminBootstrap',token:admin.token,actor:admin.actor,limit:100});
      if(!resp.ok) throw new Error(resp.error||'Не удалось подключить администратора');
      admin.connected=true;admin.audit=resp.audit||[];admin.control=resp.control||[];
      sessionStorage.setItem('sten_admin_token',admin.token);sessionStorage.setItem('sten_admin_actor',admin.actor);
      if(resp.data){applyRemoteModel(resp.data);sourceMode='api';}
      renderAll();toast('Администрирование подключено.','success');
    }catch(err){admin.connected=false;renderAdmin();toast(err.message,'error');}
    finally{setBusy(false);}
  }
  function adminLogout(){admin.connected=false;admin.token='';admin.actor='';admin.audit=[];admin.control=[];sessionStorage.removeItem('sten_admin_token');sessionStorage.removeItem('sten_admin_actor');renderAll();toast('Администратор отключён.');}
  function goAdmin(message=''){setActiveTab('admin');if(message)toast(message,'error');}

  async function addMaterial(e){
    e.preventDefault();if(!admin.connected){goAdmin();return;}
    const form=e.currentTarget,fd=new FormData(form),material={};for(const [k,v] of fd.entries())material[k]=v;
    material.price=Number(material.price||0);material.densitySpec=material.densitySpec===''?'':Number(material.densitySpec||0);
    setBusy(true);$('#newMaterialResult').textContent='Сохранение…';
    try{const resp=await apiPost({action:'addMaterial',token:admin.token,actor:admin.actor,material});if(!resp.ok)throw new Error(resp.error||'Материал не добавлен');if(resp.data)applyRemoteModel(resp.data);if(resp.audit)admin.audit=resp.audit;$('#newMaterialResult').textContent=`Добавлен ${resp.material.code}`;form.reset();renderAll();toast(`Материал ${resp.material.code} добавлен в Google Sheets.`,'success');}catch(err){$('#newMaterialResult').textContent=err.message;toast(err.message,'error');}finally{setBusy(false);}
  }

  async function uploadCertificate(e){
    e.preventDefault();if(!admin.connected){goAdmin();return;}
    const form=e.currentTarget,fd=new FormData(form),file=fd.get('certificateFile');if(!(file instanceof File)||!file.size){toast('Выберите файл.','error');return;}if(file.size>MAX_CERT_BYTES){toast('Файл больше 8 МБ.','error');return;}
    setBusy(true);$('#certificateUploadResult').textContent='Загрузка…';
    try{const base64=await fileToBase64(file);const resp=await apiPost({action:'uploadCertificate',token:admin.token,actor:admin.actor,code:fd.get('code'),documentName:fd.get('documentName'),fileName:file.name,mimeType:file.type||'application/pdf',base64});if(!resp.ok)throw new Error(resp.error||'Сертификат не загружен');if(resp.data)applyRemoteModel(resp.data);else if(resp.certificates)model.certificates=resp.certificates;$('#certificateUploadResult').textContent='Документ сохранён';form.reset();renderAll();toast('Сертификат сохранён централизованно.','success');}catch(err){$('#certificateUploadResult').textContent=err.message;toast(err.message,'error');}finally{setBusy(false);}
  }

  async function systemAction(action){
    if(!admin.connected){goAdmin('Системная операция требует Admin key.');return;}
    setBusy(true);
    try{const resp=await apiPost({action,token:admin.token,actor:admin.actor});if(!resp.ok)throw new Error(resp.error||'Ошибка операции');if(resp.data){applyRemoteModel(resp.data);sourceMode='api';}if(resp.control)admin.control=resp.control;if(resp.strategic)admin.strategic=resp.strategic;if(resp.backup)toast(`Backup создан: ${resp.backup.name}`,'success');else toast('Операция выполнена.','success');renderAll();}catch(err){toast(err.message,'error');}finally{setBusy(false);}
  }

  async function buildPdfReport(){
    const btn=$('#pdfBtn');btn.classList.add('pdf-pending');setBusy(true);
    try{
      const report={source:sourceMode==='api'?(model.source||'Google Sheets'):'Fallback snapshot v3.8',inputs:clone(state.inputs),project:clone(state.project),scenarios:clone(state.scenarios),results:results.map(r=>({categories:Object.fromEntries(['ext','inter','part'].map(c=>[c,{total:r.categories[c].total}])),changedTotal:r.changedTotal,fullTotal:r.fullTotal,deltaSaleObj:r.deltaSaleObj,deltaSaleValue:r.deltaSaleValue}))};
      const resp=await apiPost({action:'buildPdfReport',report});if(!resp.ok)throw new Error(resp.error||'PDF не сформирован');downloadBase64File(resp.base64,resp.mime||'application/pdf',resp.fileName||'STEN_MASTER_Расчёт.pdf');toast('PDF отчёт сформирован.','success');
    }catch(err){toast('Backend PDF пока недоступен. Открыта печатная версия, её можно сохранить как PDF.','error');setTimeout(()=>window.print(),100);}finally{btn.classList.remove('pdf-pending');setBusy(false);}
  }

  function renderSystem(){renderEngineCheck();renderSystemMeta();renderSystemOutput();}
  function renderEngineCheck(){const check=Engine.runGoldenTests(model);$('#engineCheck').className=`check-box ${check.pass?'pass':'fail'}`;$('#engineCheck').innerHTML=check.pass?`<div class="check-title">✓ Расчётный движок совпадает с Google Sheets</div><div>${check.passed} / ${check.total} контрольных значений совпали.</div>`:`<div class="check-title">✕ Найдены расхождения</div><div>${check.passed} / ${check.total}. Первое: ${esc(check.failed[0]?.name||'')}</div>`;}
  function renderSystemMeta(){if(!$('#systemMeta'))return;$('#systemMeta').innerHTML=[['Источник данных',sourceMode==='api'?'Google Sheets API':'Fallback snapshot'],['Версия модели',model.version||'3.8'],['Синхронизация',model.timestamp?new Date(model.timestamp).toLocaleString('ru-RU'):'снимок сборки'],['Работы',(model.works||[]).length],['Материалы',(model.materials||[]).length],['Конструкции',Object.values(model.constructions||{}).reduce((a,b)=>a+b.length,0)],['Сертификаты',(model.certificates||[]).length],['Администратор',admin.connected?'подключён':'не подключён']].map(([a,b])=>`<div class="meta-row"><span>${esc(a)}</span><span>${esc(b)}</span></div>`).join('');}
  function renderSystemOutput(){const out=$('#systemOutput');if(!out)return;if(admin.view==='audit'){out.innerHTML=admin.audit.length?`<table class="data-table"><thead><tr><th>Дата</th><th>Кто</th><th>Лист</th><th>Код</th><th>Что</th><th>Было</th><th>Стало</th></tr></thead><tbody>${admin.audit.map(x=>`<tr><td>${esc(x.date)}</td><td>${esc(x.actor)}</td><td>${esc(x.sheet)}</td><td>${esc(x.code)}</td><td>${esc(x.field)}</td><td>${esc(x.before)}</td><td>${esc(x.after)}</td></tr>`).join('')}</tbody></table>`:'<div class="empty-state">Подключите администратора, чтобы загрузить журнал с авторством.</div>';}else if(admin.view==='strategic'){const strategic=(model.materials||[]).filter(x=>x.purchaseGroup==='Стратегические материалы');out.innerHTML=`<table class="data-table"><thead><tr><th>Код</th><th>Материал</th><th>Группа</th><th>Цена</th><th>Использование</th></tr></thead><tbody>${strategic.map(x=>`<tr><td>${esc(x.code)}</td><td>${esc(x.name)}</td><td>${esc(x.technicalGroup)}</td><td class="num">${fmtRub(x.price,2)}</td><td>${esc(shortUsage(x.usage))}</td></tr>`).join('')}</tbody></table>`;}else{if(!admin.control.length){out.innerHTML='<div class="empty-state">Запустите «Проверить целостность» в разделе администрирования.</div>';return;}out.innerHTML=`<table class="data-table"><tbody>${admin.control.map(r=>`<tr>${r.map(x=>`<td>${esc(x)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;}}

  function setActiveTab(tab,save=true){state.activeTab=tab;if(save)saveState();$$('.tab').forEach(x=>x.classList.toggle('active',x.dataset.tab===tab));$$('.tab-panel').forEach(x=>x.classList.toggle('active',x.dataset.panel===tab));if(tab==='analytics'&&!historyData)renderHistoryPicker();if(tab==='certificates')renderCertificates();if(tab==='admin')renderAdmin();}

  async function syncFromSheets(){
    setSync('loading','Синхронизация…');
    try{const resp=await fetch(`${API_URL}?action=bootstrap&force=1&_=${Date.now()}`,{cache:'no-store'});if(!resp.ok)throw new Error(`HTTP ${resp.status}`);const data=await resp.json();if(!data.ok||!data.constructions||!data.works)throw new Error(data.error||'Неверный ответ API');if(!String(data.apiVersion||'').startsWith('3.8'))throw new Error('Backend Apps Script ещё не обновлён до v3.8');applyRemoteModel(data);sourceMode='api';renderAll();setSync('online','Google Sheets · '+new Date(data.timestamp).toLocaleTimeString('ru-RU',{hour:'2-digit',minute:'2-digit'}));toast('Данные синхронизированы с Google Sheets.','success');}catch(err){sourceMode='fallback';setSync('offline','Fallback v3.8 · Google Sheets v3.7');toast(`Синхронизация: ${err.message}`,'error');}
  }
  function applyRemoteModel(data){const baseline=data.baseline||model.baseline||FALLBACK.baseline;model={...data,baseline,certificates:data.certificates||[]};validateState();}

  async function apiPost(payload){const r=await fetch(API_URL,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(payload),redirect:'follow'});const text=await r.text();try{return JSON.parse(text);}catch{throw new Error('Web API вернул не-JSON ответ. Проверьте deployment Apps Script.');}}
  function setSync(mode,text){const el=$('#syncState');el.className='sync-state '+(mode==='online'?'':mode==='loading'?'offline':mode==='offline'?'offline':'error');$('#syncText').textContent=text;}
  function setBusy(v){$$('button').forEach(b=>{if(!['resetBtn','instructionCloseBtn'].includes(b.id))b.disabled=v;});}
  function toast(msg,type=''){const el=document.createElement('div');el.className='toast '+type;el.textContent=msg;$('#toastStack').appendChild(el);setTimeout(()=>el.remove(),4500);}
  function initials(name){return String(name||'').replace(/["'()]/g,' ').split(/\s+/).filter(Boolean).slice(0,2).map(x=>x[0]).join('').toUpperCase()||'СМ';}
  function shortUsage(v){const s=String(v||'');if(!s)return '—';const n=s.split(';').length;return n>2?`${n} мест использования`:s;}
  function cssEscape(v){return String(v).replace(/"/g,'\\"');}
  function fmtBytes(n){const v=Number(n)||0;return v>1024*1024?`${fmtN(v/1024/1024,1)} МБ`:v>1024?`${fmtN(v/1024,0)} КБ`:`${v} Б`;}
  function fileToBase64(file){return new Promise((resolve,reject)=>{const r=new FileReader();r.onload=()=>resolve(String(r.result).split(',')[1]||'');r.onerror=()=>reject(r.error||new Error('Ошибка чтения файла'));r.readAsDataURL(file);});}
  function base64ToBlob(b64,mime){const bin=atob(b64),bytes=new Uint8Array(bin.length);for(let i=0;i<bin.length;i++)bytes[i]=bin.charCodeAt(i);return new Blob([bytes],{type:mime||'application/octet-stream'});}
  function downloadBase64File(b64,mime,name){const url=URL.createObjectURL(base64ToBlob(b64,mime));const a=document.createElement('a');a.href=url;a.download=name||'file';document.body.appendChild(a);a.click();a.remove();setTimeout(()=>URL.revokeObjectURL(url),10000);}
  function openBase64File(b64,mime,name,newTab){const url=URL.createObjectURL(base64ToBlob(b64,mime));if(newTab){const w=window.open(url,'_blank','noopener');if(!w)downloadBase64File(b64,mime,name);}else downloadBase64File(b64,mime,name);setTimeout(()=>URL.revokeObjectURL(url),60000);}

  function openInstruction(){const d=$('#instructionDialog');if(typeof d.showModal==='function')d.showModal();else d.setAttribute('open','');}
  function closeInstruction(){const d=$('#instructionDialog');if(typeof d.close==='function')d.close();else d.removeAttribute('open');}

  function bindStaticEvents(){
    $('#tabs').addEventListener('click',e=>{const b=e.target.closest('[data-tab]');if(b)setActiveTab(b.dataset.tab);});
    $('#syncBtn').onclick=syncFromSheets;$('#pdfBtn').onclick=buildPdfReport;$('#adminHeaderBtn').onclick=()=>setActiveTab('admin');$('#pricingAdminBtn').onclick=()=>setActiveTab('admin');$('#certificateAdminBtn').onclick=()=>setActiveTab('admin');$('#systemAdminBtn').onclick=()=>setActiveTab('admin');
    $('#resetBtn').onclick=()=>{state=defaultState(model.defaults);saveState();renderAll();toast('Параметры возвращены к текущим значениям Google Sheets.');};
    $('#instructionBtn').onclick=openInstruction;$('#footerInstructionBtn').onclick=openInstruction;$('#instructionCloseBtn').onclick=closeInstruction;$('#instructionDialog').addEventListener('click',e=>{if(e.target===$('#instructionDialog'))closeInstruction();});
    $('#adminLoginBtn').onclick=adminLogin;$('#adminLogoutBtn').onclick=adminLogout;
    ['priceSearch','priceType','pricePurchase'].forEach(id=>document.getElementById(id).addEventListener(id==='priceSearch'?'input':'change',renderPricing));
    $('#historySearch').addEventListener('input',renderHistoryPicker);$('#historyGrouping').addEventListener('change',()=>{if(historyData)renderPriceCharts();});$('#loadHistoryBtn').onclick=loadPriceHistory;
    $('#certificateSearch').addEventListener('input',renderCertificates);
    $('#newMaterialForm').addEventListener('submit',addMaterial);$('#certificateUploadForm').addEventListener('submit',uploadCertificate);
    document.addEventListener('click',e=>{const d=e.target.closest('[data-cert-download]');if(d){e.preventDefault();downloadCertificate(d.dataset.certDownload);return;}const x=e.target.closest('[data-cert-delete]');if(x){e.preventDefault();deleteCertificate(x.dataset.certDelete);}});
    $$('.system-action').forEach(b=>b.onclick=()=>systemAction(b.dataset.action));
    $$('.seg').forEach(b=>b.onclick=()=>{admin.view=b.dataset.systemView;$$('.seg').forEach(x=>x.classList.toggle('active',x===b));renderSystemOutput();});
    window.addEventListener('resize',()=>{if(historyData)requestAnimationFrame(renderPriceCharts);});
  }

  bindStaticEvents();
  renderAll();
  setSync('offline','Fallback v3.8 · Google Sheets v3.7');
  syncFromSheets();
})();
