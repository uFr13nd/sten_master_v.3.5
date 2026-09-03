(() => {
  'use strict';

  const API_URL = 'https://script.google.com/macros/s/AKfycbw81LdRYaPOAiE7rHCaXJBdxb7J98B5Jt_uOoBDX49BYSJ-vZsfRu2I5y7fEk6n6K-U/exec';
  const FALLBACK = window.STEN_FALLBACK_DATA;
  const Engine = window.StenEngine;
  const STORAGE_KEY = 'sten-master-web-v35-state';

  let model = clone(FALLBACK);
  let sourceMode = 'fallback';
  let state = loadState(model.defaults);
  let results = [];
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
  const fmtSigned = (v, unit='') => `${v>0?'+':''}${fmtN(v,2)}${unit}`;
  const finishOptions = ['Без отделки','Предчистовая','Чистовая'];
  const catMeta = {
    ext:{label:'Наружные стены',icon:'🧱',unit:'м³',priceUnit:'₽/м³'},
    inter:{label:'Межквартирные',icon:'🏢',unit:'м³',priceUnit:'₽/м³'},
    part:{label:'Перегородки',icon:'🚪',unit:'м²',priceUnit:'₽/м²'}
  };

  function clone(x){ return JSON.parse(JSON.stringify(x)); }
  function defaultState(defaults){
    return {
      inputs:{city:defaults.city,saleArea:defaults.saleArea,totalArea:defaults.totalArea,perimExt:defaults.perimExt,perimInter:defaults.perimInter,perimPart:defaults.perimPart,floorH:defaults.floorH,floors:defaults.floors,salePrice:defaults.salePrice},
      project:clone(defaults.project), scenarios:clone(defaults.scenarios), activeTab:'overview'
    };
  }
  function loadState(defaults){
    try{
      const saved=JSON.parse(localStorage.getItem(STORAGE_KEY)||'null');
      return saved && saved.inputs && saved.project && saved.scenarios ? saved : defaultState(defaults);
    }catch{return defaultState(defaults);}
  }
  function saveState(){ localStorage.setItem(STORAGE_KEY,JSON.stringify(state)); }

  function validateState(){
    for(const cat of ['ext','inter','part']){
      const names=(model.constructions[cat]||[]).map(x=>x.name);
      if(!names.includes(state.project?.[cat]?.material)) state.project[cat]=clone(model.defaults.project[cat]);
      state.scenarios.forEach((sc,i)=>{
        if(!names.includes(sc?.[cat]?.material)) sc[cat]=clone(model.defaults.scenarios[i]?.[cat] || model.defaults.project[cat]);
        if(!finishOptions.includes(sc[cat].finish)) sc[cat].finish='Без отделки';
      });
      if(!finishOptions.includes(state.project[cat].finish)) state.project[cat].finish='Без отделки';
    }
  }

  function renderAll(){
    validateState();
    results=Engine.calcAll(state,model);
    renderInputs(); renderProject(); renderOverview();
    renderCategory('ext'); renderCategory('inter'); renderCategory('part');
    renderPricing(); renderSystem();
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
    $$('[data-input]').forEach(el=>el.addEventListener('input',e=>{
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
      const deltaCls=r.deltaSaleValue>0?'good':r.deltaSaleValue<0?'bad':'';
      return `<article class="scenario-card ${best?'best':''}">${best?'<span class="best-tag">Мин. прямые затраты</span>':''}
        <h3 class="scenario-title">Сценарий ${i+1}</h3>
        <div class="category-mini">${['ext','inter','part'].map(cat=>`<div class="mini-row"><span>${catMeta[cat].icon}</span><span class="mini-name" title="${esc(state.scenarios[i][cat].material)}">${esc(state.scenarios[i][cat].material)}</span><span class="mini-cost">${fmtRub(r.categories[cat].total)}</span></div>`).join('')}</div>
        <div class="metric-list">
          <div class="metric"><span class="metric-label">Полная стоимость выбранных решений</span><span class="metric-value primary">${fmtRub(r.fullTotal)}</span></div>
          <div class="metric"><span class="metric-label">Итого изменяемых решений</span><span class="metric-value">${fmtRub(r.changedTotal)}</span></div>
          <div class="metric"><span class="metric-label">Δ продаваемой площади объекта</span><span class="metric-value ${r.deltaSaleObj>0?'good':r.deltaSaleObj<0?'bad':''}">${fmtSigned(r.deltaSaleObj,' м²')}</span></div>
          <div class="metric"><span class="metric-label">Δ стоимости продаваемой площади</span><span class="metric-value ${deltaCls}">${r.deltaSaleValue>0?'+':''}${fmtRub(r.deltaSaleValue)}</span></div>
          <div class="metric"><span class="metric-label">Аналитический net-эффект*</span><span class="metric-value gold">${r.netEconomicEffect>0?'+':''}${fmtRub(r.netEconomicEffect)}</span></div>
        </div>
        <div class="scenario-total"><div class="metric"><span class="metric-label">Изменяемые решения / м² продаваемой</span><span class="metric-value primary">${fmtRub(r.changedPerSaleArea)}</span></div></div>
      </article>`;
    }).join('');

    const rows=[
      ['Δ площади под стенами на этаж','deltaFootprint','м²'],['Δ продаваемой площади на этаж','deltaSaleFloor','м²'],['Δ продаваемой площади объекта','deltaSaleObj','м²'],['Δ стоимости продаваемой площади','deltaSaleValue','₽'],
      ['Наружные — изменяемое решение','changedExt','₽'],['Межквартирные — изменяемое решение','changedInter','₽'],['Перегородки — изменяемое решение','changedPart','₽'],['ИТОГО ИЗМЕНЯЕМЫХ РЕШЕНИЙ','changedTotal','₽'],['Цена изменяемых / м² общей площади','changedPerTotalArea','₽/м²'],['Цена изменяемых / м² продаваемой площади','changedPerSaleArea','₽/м²']
    ];
    $('#summaryTable').innerHTML=`<table class="data-table"><thead><tr><th>Показатель</th><th>Сценарий 1</th><th>Сценарий 2</th><th>Сценарий 3</th></tr></thead><tbody>${rows.map(([label,key,unit])=>`<tr><td>${esc(label)}</td>${results.map(r=>{
      let val=key==='changedExt'?r.changedByCategory.ext:key==='changedInter'?r.changedByCategory.inter:key==='changedPart'?r.changedByCategory.part:r[key];
      if(unit==='₽'||unit==='₽/м²') return `<td class="num">${fmtRub(val)}</td>`;
      return `<td class="num">${fmtSigned(val,' '+unit)}</td>`;
    }).join('')}</tr>`).join('')}</tbody></table>`;
  }

  function renderCategory(cat){
    const container=$(`#${cat}Variants`);
    if(!container) return;
    const totals=results.map(r=>r.categories[cat].total), min=Math.min(...totals);
    container.innerHTML=state.scenarios.map((sc,i)=>renderVariantCard(cat,i,results[i].categories[cat],Math.abs(totals[i]-min)<1e-6)).join('');
    $$(`[data-scenario-material^="${cat}:"]`).forEach(el=>el.onchange=e=>{const [,idx]=e.currentTarget.dataset.scenarioMaterial.split(':');state.scenarios[Number(idx)][cat].material=e.currentTarget.value;saveState();renderDynamicOnly();});
    $$(`[data-scenario-finish^="${cat}:"]`).forEach(el=>el.onchange=e=>{const [,idx]=e.currentTarget.dataset.scenarioFinish.split(':');state.scenarios[Number(idx)][cat].finish=e.currentTarget.value;saveState();renderDynamicOnly();});
    renderCompare(cat);
  }

  function renderVariantCard(cat,i,c,best){
    const selection=state.scenarios[i][cat], mats=model.constructions[cat]||[], mat=Engine.getConstruction(model,cat,selection.material);
    const strategic=(mat?.components||[]).filter(x=>x.purchaseGroup==='Стратегические материалы').length;
    const reference=String(mat?.reference??'');
    let badge='';
    if(cat==='inter'||cat==='part') badge=reference.includes('Не проходит')?`<span class="badge fail">${esc(reference)}</span>`:reference.includes('Проходит')?`<span class="badge pass">${esc(reference)}</span>`:`<span class="badge warn">${esc(reference||'Нет данных')}</span>`;
    if(cat==='ext' && reference!=='') badge=`<span class="badge">Утеплитель ≥ ${fmtN(reference,0)} мм</span>`;
    return `<article class="variant-card ${best?'best':''}">${best?'<span class="best-tag">Мин. стоимость</span>':''}
      <div class="material-glyph">${esc(initials(selection.material))}</div>
      <h3 class="variant-title">Сценарий ${i+1}</h3>
      <div class="variant-controls">
        <div class="field"><label>Материал</label><select class="variant-select" data-scenario-material="${cat}:${i}">${mats.map(x=>`<option ${x.name===selection.material?'selected':''}>${esc(x.name)}</option>`).join('')}</select></div>
        <div class="field"><label>Тип отделки</label><select class="variant-select" data-scenario-finish="${cat}:${i}">${finishOptions.map(x=>`<option ${x===selection.finish?'selected':''}>${esc(x)}</option>`).join('')}</select></div>
      </div>
      <div class="badges"><span class="badge">📍 ${esc(c.city||'—')}</span>${strategic?`<span class="badge strategic">★ стратегических: ${strategic}</span>`:''}${badge}</div>
      <div class="metric-list">
        <div class="metric"><span class="metric-label">Толщина кладки / с отделкой</span><span class="metric-value">${fmtN(c.thick,0)} / ${fmtN(c.thickFin,0)} мм</span></div>
        <div class="metric"><span class="metric-label">Цена конструкции</span><span class="metric-value primary">${fmtRub(c.unitPrice,2)} / ${catMeta[cat].unit}</span></div>
        <div class="metric"><span class="metric-label">В том числе работы</span><span class="metric-value">${fmtRub(c.workUnit,2)}</span></div>
        <div class="metric"><span class="metric-label">В том числе материалы</span><span class="metric-value">${fmtRub(c.materialUnit,2)}</span></div>
        <div class="metric"><span class="metric-label">${cat==='part'?'Площадь кладки/этаж':'Объём кладки/этаж'}</span><span class="metric-value">${fmtN(cat==='part'?c.wallAreaFloor:c.volumeFloor,2)} ${catMeta[cat].unit}</span></div>
        <div class="metric"><span class="metric-label">Масса / этаж</span><span class="metric-value">${fmtN(c.massFloor,2)} т</span></div>
        <div class="metric"><span class="metric-label">Доставка</span><span class="metric-value">${c.trucks} маш. · ${fmtRub(c.deliveryCost)}</span></div>
        <div class="metric"><span class="metric-label">Отделка</span><span class="metric-value">${fmtRub(c.finishCost)}</span></div>
        <div class="metric"><span class="metric-label">Работы + материалы на объект</span><span class="metric-value">${fmtRub(c.costObj)}</span></div>
        <div class="metric"><span class="metric-label">ИТОГО на объект</span><span class="metric-value primary">${fmtRub(c.total)}</span></div>
      </div>
      ${renderComponents(mat)}
    </article>`;
  }

  function renderComponents(mat){
    const comps=mat?.components||[];
    if(!comps.length) return '';
    return `<details class="details"><summary>Состав конструкции · ${comps.length} поз.</summary><div class="components"><table class="data-table"><thead><tr><th>Код</th><th>Тип</th><th>Наименование</th><th>Вклад</th><th>Закупка</th></tr></thead><tbody>${comps.map(c=>`<tr><td>${esc(c.code)}</td><td>${esc(c.type)}</td><td>${esc(c.name)}</td><td class="num">${fmtRub(c.sheetContribution,2)}</td><td>${c.purchaseGroup==='Стратегические материалы'?'<span class="badge strategic">Стратегический</span>':esc(c.purchaseGroup||'')}</td></tr>`).join('')}</tbody></table></div></details>`;
  }

  function renderCompare(cat){
    const target=$(`#${cat}Compare`), items=[...(model.constructions[cat]||[])].sort((a,b)=>a.sheetPrice-b.sheetPrice);
    target.innerHTML=`<table class="data-table"><thead><tr><th>Материал</th><th>Город</th><th>Толщина</th><th>Работы</th><th>Материалы</th><th>Цена конструкции</th><th>${cat==='ext'?'Утеплитель':'Звукоизоляция'}</th></tr></thead><tbody>${items.map((x,i)=>`<tr><td>${i===0?'⭐ ':''}${esc(x.name)}</td><td>${esc(x.city)}</td><td class="num">${fmtN(x.thick,0)} мм</td><td class="num">${fmtRub(x.workCostSheet,2)}</td><td class="num">${fmtRub(x.materialCostSheet,2)}</td><td class="num">${fmtRub(x.sheetPrice,2)}</td><td>${esc(x.reference??'')}</td></tr>`).join('')}</tbody></table>`;
  }

  function renderPricing(){
    const q=($('#priceSearch')?.value||'').trim().toLowerCase(), type=$('#priceType')?.value||'all', purchase=$('#pricePurchase')?.value||'all';
    let rows=[];
    if(type!=='materials') rows.push(...(model.works||[]).map(x=>({...x,_kind:'works',purchaseGroup:''})));
    if(type!=='works') rows.push(...(model.materials||[]).map(x=>({...x,_kind:'materials'})));
    rows=rows.filter(x=>(!q || `${x.code} ${x.name} ${x.technicalGroup}`.toLowerCase().includes(q)) && (purchase==='all'||x.purchaseGroup===purchase));
    const techGroups=[...new Set((model.materials||[]).map(x=>x.technicalGroup).filter(Boolean))].sort();
    const workGroups=[...new Set((model.works||[]).map(x=>x.technicalGroup).filter(Boolean))].sort();
    $('#pricingTable').innerHTML=`<table class="data-table"><thead><tr><th>Код</th><th>Наименование</th><th>Ед.</th><th>Цена</th><th>Группа</th><th>Закупка</th><th>Где используется</th>${admin.connected?'<th></th>':''}</tr></thead><tbody>${rows.map(x=>{
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

  function renderSystem(){ renderEngineCheck(); renderSystemMeta(); renderSystemOutput(); }
  function renderEngineCheck(){
    const check=Engine.runGoldenTests(model);
    $('#engineCheck').className=`check-box ${check.pass?'pass':'fail'}`;
    $('#engineCheck').innerHTML=check.pass?`<div class="check-title">✓ Расчётный движок совпадает с Google Sheets</div><div>${check.passed} / ${check.total} контрольных значений совпали.</div>`:`<div class="check-title">✕ Найдены расхождения</div><div>${check.passed} / ${check.total}. Первое: ${esc(check.failed[0]?.name||'')}</div>`;
  }
  function renderSystemMeta(){
    if(!$('#systemMeta')) return;
    $('#systemMeta').innerHTML=[
      ['Источник данных',sourceMode==='api'?'Google Sheets API':'Fallback snapshot'],['Версия модели',model.version||'3.5'],['Синхронизация',model.timestamp?new Date(model.timestamp).toLocaleString('ru-RU'):'снимок сборки'],['Работы',(model.works||[]).length],['Материалы',(model.materials||[]).length],['Конструкции',Object.values(model.constructions||{}).reduce((a,b)=>a+b.length,0)],['Администратор',admin.connected?'подключён':'не подключён']
    ].map(([a,b])=>`<div class="meta-row"><span>${esc(a)}</span><span>${esc(b)}</span></div>`).join('');
  }
  function renderSystemOutput(){
    const out=$('#systemOutput'); if(!out) return;
    if(admin.view==='audit'){
      out.innerHTML=admin.audit.length?`<table class="data-table"><thead><tr><th>Дата</th><th>Кто</th><th>Лист</th><th>Код</th><th>Что</th><th>Было</th><th>Стало</th></tr></thead><tbody>${admin.audit.map(x=>`<tr><td>${esc(x.date)}</td><td>${esc(x.actor)}</td><td>${esc(x.sheet)}</td><td>${esc(x.code)}</td><td>${esc(x.field)}</td><td>${esc(x.before)}</td><td>${esc(x.after)}</td></tr>`).join('')}</tbody></table>`:'<div class="card">Подключите режим администратора, чтобы загрузить журнал.</div>';
    }else if(admin.view==='strategic'){
      const strategic=(model.materials||[]).filter(x=>x.purchaseGroup==='Стратегические материалы');
      out.innerHTML=`<table class="data-table"><thead><tr><th>Код</th><th>Материал</th><th>Группа</th><th>Цена</th><th>Использование</th></tr></thead><tbody>${strategic.map(x=>`<tr><td>${esc(x.code)}</td><td>${esc(x.name)}</td><td>${esc(x.technicalGroup)}</td><td class="num">${fmtRub(x.price,2)}</td><td>${esc(shortUsage(x.usage))}</td></tr>`).join('')}</tbody></table>`;
    }else{
      if(!admin.control.length){out.innerHTML='<div style="padding:20px" class="muted">Запустите «Проверить целостность» в режиме администратора.</div>';return;}
      out.innerHTML=`<table class="data-table"><tbody>${admin.control.map(r=>`<tr>${r.map(x=>`<td>${esc(x)}</td>`).join('')}</tr>`).join('')}</tbody></table>`;
    }
  }

  function setActiveTab(tab,save=true){
    state.activeTab=tab; if(save)saveState();
    $$('.tab').forEach(x=>x.classList.toggle('active',x.dataset.tab===tab));
    $$('.tab-panel').forEach(x=>x.classList.toggle('active',x.dataset.panel===tab));
  }

  async function syncFromSheets(){
    setSync('loading','Синхронизация…');
    try{
      const resp=await fetch(`${API_URL}?action=bootstrap&force=1&_=${Date.now()}`,{cache:'no-store'});
      if(!resp.ok) throw new Error(`HTTP ${resp.status}`);
      const data=await resp.json();
      if(!data.ok || !data.constructions || !data.works) throw new Error(data.error||'Неверный ответ API');
      applyRemoteModel(data); sourceMode='api'; renderAll(); setSync('online','Google Sheets · '+new Date(data.timestamp).toLocaleTimeString('ru-RU',{hour:'2-digit',minute:'2-digit'})); toast('Данные синхронизированы с Google Sheets.','success');
    }catch(err){sourceMode='fallback';setSync('offline','Fallback · автономный режим');toast(`Синхронизация недоступна: ${err.message}`,'error');}
  }

  function applyRemoteModel(data){
    const baseline=data.baseline || model.baseline || FALLBACK.baseline;
    model={...data,baseline};
    validateState();
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
      $('#adminPanel').classList.add('hidden');$('#adminToggleBtn').textContent='Администратор подключён';toast('Редактирование расценок подключено.','success');renderAll();
    }catch(err){admin.connected=false;toast(err.message,'error');}
    finally{setBusy(false);}
  }

  async function systemAction(action){
    if(!admin.connected){$('#adminPanel').classList.remove('hidden');setActiveTab('pricing');toast('Сначала подключите режим администратора.','error');return;}
    setBusy(true);
    try{
      const resp=await apiPost({action,token:admin.token,actor:admin.actor});
      if(!resp.ok) throw new Error(resp.error||'Ошибка операции');
      if(resp.data){applyRemoteModel(resp.data);sourceMode='api';}
      if(resp.control) admin.control=resp.control;
      if(resp.strategic) admin.strategic=resp.strategic;
      if(resp.backup) toast(`Backup создан: ${resp.backup.name}`,'success'); else toast('Операция выполнена.','success');
      renderAll();
    }catch(err){toast(err.message,'error');}
    finally{setBusy(false);}
  }

  async function apiPost(payload){
    const r=await fetch(API_URL,{method:'POST',headers:{'Content-Type':'text/plain;charset=utf-8'},body:JSON.stringify(payload),redirect:'follow'});
    const text=await r.text();
    try{return JSON.parse(text);}catch{throw new Error('Web API вернул не-JSON ответ. Проверьте deployment Apps Script.');}
  }

  function setSync(mode,text){const el=$('#syncState');el.className='sync-state '+(mode==='online'?'':mode==='loading'?'offline':mode==='offline'?'offline':'error');$('#syncText').textContent=text;}
  function setBusy(v){$$('button').forEach(b=>{if(b.id!=='resetBtn')b.disabled=v;});}
  function toast(msg,type=''){const el=document.createElement('div');el.className='toast '+type;el.textContent=msg;$('#toastStack').appendChild(el);setTimeout(()=>el.remove(),4500);}
  function initials(name){return String(name||'').replace(/["'()]/g,' ').split(/\s+/).filter(Boolean).slice(0,2).map(x=>x[0]).join('').toUpperCase()||'СМ';}
  function shortUsage(v){const s=String(v||'');if(!s)return '—';const n=s.split(';').length;return n>2?`${n} мест использования`:s;}
  function cssEscape(v){return String(v).replace(/"/g,'\\"');}

  function bindStaticEvents(){
    $('#tabs').addEventListener('click',e=>{const b=e.target.closest('[data-tab]');if(b)setActiveTab(b.dataset.tab);});
    $('#syncBtn').onclick=syncFromSheets;
    $('#resetBtn').onclick=()=>{state=defaultState(model.defaults);saveState();renderAll();toast('Параметры возвращены к текущим значениям Google Sheets.');};
    $('#adminToggleBtn').onclick=()=>{$('#adminPanel').classList.toggle('hidden');$('#adminToken').value=admin.token;$('#adminActor').value=admin.actor;};
    $('#adminLoginBtn').onclick=adminLogin;
    ['priceSearch','priceType','pricePurchase'].forEach(id=>document.getElementById(id).addEventListener(id==='priceSearch'?'input':'change',renderPricing));
    $$('.system-action').forEach(b=>b.onclick=()=>systemAction(b.dataset.action));
    $$('.seg').forEach(b=>b.onclick=()=>{admin.view=b.dataset.systemView;$$('.seg').forEach(x=>x.classList.toggle('active',x===b));renderSystemOutput();});
  }

  bindStaticEvents();
  renderAll();
  setSync('offline','Fallback · проверенная модель v3.5');
  // Try live data without blocking the UI. If the old deployment is still active, validation will keep fallback.
  syncFromSheets();
})();
