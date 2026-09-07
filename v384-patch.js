(() => {
  'use strict';

  const API_URL = window.STEN_CONFIG?.API_URL || 'https://script.google.com/macros/s/AKfycbwvCgR_FeVh3QkKc6u1UaxYkAFQbQkFJ6j39XxsRHeoOfTHBJHW02hTpgvRIoI-r5kD/exec';
  const CHANNEL = 'sten-master-v38';
  const TIMEOUT = 30000;
  const JSONP_TIMEOUT = 30000;
  const pending = new Map();
  let bootstrapData = null;
  let pricingObserver = null;

  const $ = (s, root = document) => root.querySelector(s);
  const $$ = (s, root = document) => [...root.querySelectorAll(s)];
  const esc = v => String(v ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));

  function requestId() {
    if (window.crypto && crypto.getRandomValues) {
      const a = new Uint32Array(4);
      crypto.getRandomValues(a);
      return Array.from(a, x => x.toString(16).padStart(8, '0')).join('');
    }
    return `${Date.now()}-${Math.random().toString(36).slice(2)}`;
  }

  window.addEventListener('message', event => {
    let trusted = false;
    try {
      const host = new URL(event.origin).hostname;
      trusted = event.origin === 'https://script.google.com' || host === 'script.google.com' || host.endsWith('.googleusercontent.com');
    } catch (_) {}
    if (!trusted) return;
    const data = event.data || {};
    if (data.channel !== CHANNEL || !data.id) return;
    const item = pending.get(String(data.id));
    if (!item) return;
    pending.delete(String(data.id));
    clearTimeout(item.timer);
    item.cleanup();
    item.resolve(data.response || {ok:false,error:'Пустой ответ Web API'});
  });

  function apiGetJsonp(action, params = {}) {
    const id = requestId().replace(/[^A-Za-z0-9_]/g, '');
    const cb = `__stenPatchJsonp_${id}`;
    return new Promise((resolve, reject) => {
      const script = document.createElement('script');
      const cleanup = () => { clearTimeout(timer); delete window[cb]; script.remove(); };
      const timer = setTimeout(() => { cleanup(); reject(new Error('GET bootstrap не ответил.')); }, JSONP_TIMEOUT);
      window[cb] = data => { cleanup(); resolve(data); };
      const q = new URLSearchParams({action, callback:cb, _:String(Date.now())});
      Object.entries(params || {}).forEach(([k,v]) => q.set(k, String(v)));
      script.src = `${API_URL}?${q.toString()}`;
      script.async = true;
      script.onerror = () => { cleanup(); reject(new Error('GET bootstrap недоступен.')); };
      document.head.appendChild(script);
    });
  }

  function apiCall(payload) {
    const id = requestId();
    const frameName = `sten_v384_${id}`;
    return new Promise((resolve, reject) => {
      const iframe = document.createElement('iframe');
      iframe.name = frameName;
      iframe.title = 'STEN MASTER API v3.8.4';
      iframe.style.cssText = 'position:fixed;width:1px;height:1px;opacity:0;pointer-events:none;border:0;left:-9999px;top:-9999px';
      const form = document.createElement('form');
      form.method = 'POST';
      form.action = API_URL;
      form.target = frameName;
      form.style.display = 'none';
      const fields = {
        bridge: '1',
        bridgeChannel: CHANNEL,
        requestId: id,
        origin: location.origin,
        payload: JSON.stringify(payload)
      };
      Object.entries(fields).forEach(([name, value]) => {
        const input = document.createElement('input');
        input.type = 'hidden';
        input.name = name;
        input.value = value;
        form.appendChild(input);
      });
      const cleanup = () => setTimeout(() => { form.remove(); iframe.remove(); }, 0);
      const timer = setTimeout(() => {
        pending.delete(id);
        cleanup();
        reject(new Error('Apps Script не вернул postMessage. Проверьте deployment v3.8.4.'));
      }, TIMEOUT);
      pending.set(id, {resolve, reject, timer, cleanup});
      document.body.appendChild(iframe);
      document.body.appendChild(form);
      try { form.submit(); }
      catch (err) {
        pending.delete(id);
        clearTimeout(timer);
        cleanup();
        reject(err);
      }
    });
  }

  function toast(message, type = '') {
    const stack = $('#toastStack');
    if (!stack) return;
    const el = document.createElement('div');
    el.className = `toast ${type}`;
    el.textContent = message;
    stack.appendChild(el);
    setTimeout(() => el.remove(), 5500);
  }

  function formatDate(value, withTime = false) {
    if (!value) return 'не зафиксировано';
    const d = new Date(value);
    if (Number.isNaN(d.getTime())) return String(value);
    return d.toLocaleString('ru-RU', withTime ? {day:'2-digit',month:'2-digit',year:'numeric',hour:'2-digit',minute:'2-digit'} : {day:'2-digit',month:'2-digit',year:'numeric'});
  }

  function daysOld(value) {
    if (!value) return null;
    const d = new Date(value);
    if (Number.isNaN(d.getTime())) return null;
    return Math.max(0, Math.floor((Date.now() - d.getTime()) / 86400000));
  }

  function installStyles() {
    if ($('#v384Styles')) return;
    const style = document.createElement('style');
    style.id = 'v384Styles';
    style.textContent = `
      .price-date-cell{white-space:nowrap;font-size:11px}.price-date{display:inline-flex;align-items:center;gap:5px;padding:4px 7px;border-radius:999px;background:#eef6ff;color:#0b5ba8;font-weight:700}.price-date.stale{background:#fff4d6;color:#9a5b00}.price-date.none{background:#f1f5f9;color:#64748b;font-weight:600}
      .v384-layout{display:grid;grid-template-columns:minmax(0,1.5fr) minmax(300px,.7fr);gap:18px}.v384-stack{display:flex;flex-direction:column;gap:12px}.v384-row{display:flex;gap:10px;align-items:flex-end;flex-wrap:wrap}.v384-row .field{min-width:220px;flex:1}.v384-status{padding:14px 16px;border:1px solid #dbe5ef;border-radius:12px;background:#f8fbff}.v384-status.good{border-color:#b7e1c5;background:#f3fbf6}.v384-status.blocked{border-color:#f3d28f;background:#fffaf0}.v384-status strong{display:block;margin-bottom:5px}
      .chat-list{display:flex;flex-direction:column;gap:10px;max-height:520px;overflow:auto;padding:4px}.chat-item{border:1px solid #e2e8f0;border-radius:12px;padding:10px 12px;background:#fff}.chat-meta{display:flex;gap:10px;align-items:center;justify-content:space-between;color:#64748b;font-size:11px;margin-bottom:5px}.chat-text{white-space:pre-wrap;word-break:break-word}.chat-context{margin-top:5px;color:#94a3b8;font-size:10px}.v384-note{font-size:11px;color:#64748b}.v384-empty{padding:28px;text-align:center;color:#64748b}
      @media(max-width:900px){.v384-layout{grid-template-columns:1fr}}
    `;
    document.head.appendChild(style);
  }

  function updateVersionLabels() {
    const pill = $('.version-pill');
    if (pill) pill.textContent = 'v3.8.4 WEB';
    const title = document.title;
    if (title.includes('v3.8.3')) document.title = title.replace('v3.8.3', 'v3.8.4');
    const footer = $('footer span');
    if (footer) footer.textContent = footer.textContent.replace('v3.8.3', 'v3.8.4');
  }

  function addTabsAndPanels() {
    const tabs = $('#tabs');
    if (!tabs) return;
    const adminTab = $('[data-tab="admin"]', tabs);
    if (!$('[data-tab="requests"]', tabs)) {
      const b = document.createElement('button');
      b.className = 'tab'; b.dataset.tab = 'requests'; b.textContent = '↻ Запрос цен';
      tabs.insertBefore(b, adminTab || null);
    }
    if (!$('[data-tab="chat"]', tabs)) {
      const b = document.createElement('button');
      b.className = 'tab'; b.dataset.tab = 'chat'; b.textContent = '💬 Чат';
      tabs.insertBefore(b, adminTab || null);
    }

    const main = $('main.page');
    const systemPanel = $('[data-panel="system"]');
    if (!main) return;
    if (!$('[data-panel="requests"]')) {
      const section = document.createElement('section');
      section.className = 'tab-panel'; section.dataset.panel = 'requests';
      section.innerHTML = `
        <article class="card">
          <div class="card-head"><div><span class="eyebrow">10</span><h2>Централизованный запрос актуализации цен</h2><p class="muted">Один общий запрос на актуализацию можно создавать не чаще одного раза в 30 дней. Ограничение проверяется на сервере, а не только кнопкой в браузере.</p></div></div>
          <div class="v384-layout">
            <div class="v384-stack">
              <div class="v384-row">
                <div class="field"><label for="priceReqActor">Инициатор</label><input id="priceReqActor" type="text" maxlength="200" placeholder="ФИО или email"></div>
                <div class="field"><label for="priceReqComment">Комментарий</label><input id="priceReqComment" type="text" maxlength="1000" placeholder="Что проверить в первую очередь"></div>
              </div>
              <div><button class="btn btn-primary" id="priceReqBtn" type="button">Запросить актуализацию цен</button></div>
              <div class="v384-note">Запрос записывается в лист «Запросы_цен». Повторный запрос раньше разрешённой даты backend отклонит, даже если попытаться обойти интерфейс.</div>
            </div>
            <div id="priceReqStatus" class="v384-status">Загрузка статуса…</div>
          </div>
        </article>`;
      main.insertBefore(section, systemPanel || null);
    }
    if (!$('[data-panel="chat"]')) {
      const section = document.createElement('section');
      section.className = 'tab-panel'; section.dataset.panel = 'chat';
      section.innerHTML = `
        <article class="card">
          <div class="card-head"><div><span class="eyebrow">11</span><h2>Чат проекта</h2><p class="muted">Общая история хранится централизованно в Google Sheets. Сообщения видят все пользователи веб-калькулятора.</p></div><button class="btn btn-outline" id="chatRefreshBtn" type="button">Обновить</button></div>
          <div class="v384-layout">
            <div><div id="chatList" class="chat-list"><div class="v384-empty">Загрузка истории…</div></div></div>
            <div class="v384-stack">
              <div class="field"><label for="chatActor">Имя / email</label><input id="chatActor" type="text" maxlength="200" placeholder="Кто пишет"></div>
              <div class="field"><label for="chatMessage">Сообщение</label><textarea id="chatMessage" rows="7" maxlength="2000" placeholder="Сообщение по проекту"></textarea></div>
              <button class="btn btn-primary" id="chatSendBtn" type="button">Отправить</button>
              <div class="v384-note">История не привязана к конкретному браузеру. Это общий рабочий канал проекта, а не ещё один локальный блокнот, который потом никто не найдёт.</div>
            </div>
          </div>
        </article>`;
      main.insertBefore(section, systemPanel || null);
    }

    const savedActor = localStorage.getItem('sten-v384-actor') || sessionStorage.getItem('sten_admin_actor') || '';
    if ($('#priceReqActor') && !$('#priceReqActor').value) $('#priceReqActor').value = savedActor;
    if ($('#chatActor') && !$('#chatActor').value) $('#chatActor').value = savedActor;

    const savedState = (() => { try { return JSON.parse(localStorage.getItem('sten-master-web-v38-state') || 'null'); } catch { return null; } })();
    if (savedState && ['requests','chat'].includes(savedState.activeTab)) {
      $$('.tab').forEach(x => x.classList.toggle('active', x.dataset.tab === savedState.activeTab));
      $$('.tab-panel').forEach(x => x.classList.toggle('active', x.dataset.panel === savedState.activeTab));
    }
  }

  async function loadBootstrap(silent = false) {
    try {
      const data = await apiGetJsonp('bootstrap', {force:1});
      if (!data || !data.ok) throw new Error(data?.error || 'Web API вернул ошибку');
      bootstrapData = data;
      patchPricingTable();
      if (!silent && data.apiVersion) toast(`Backend ${data.apiVersion} подключён.`, 'success');
      return true;
    } catch (err) {
      if (!silent) toast(`v3.8.4: ${err.message}`, 'error');
      return false;
    }
  }

  function patchPricingTable() {
    const table = $('#pricingTable table');
    if (!table || !bootstrapData) return;
    const header = $('thead tr', table);
    if (!header) return;
    if (!$('[data-v384-price-date-head]', header)) {
      const th = document.createElement('th');
      th.dataset.v384PriceDateHead = '1';
      th.textContent = 'Цена изменена';
      const ref = header.children[4] || null;
      header.insertBefore(th, ref);
    }
    const map = new Map((bootstrapData.materials || []).map(x => [x.code, x]));
    $$('tbody tr[data-price-row]', table).forEach(tr => {
      const code = tr.dataset.priceRow || '';
      if ($('[data-v384-price-date]', tr)) return;
      const td = document.createElement('td');
      td.dataset.v384PriceDate = '1';
      td.className = 'price-date-cell';
      const item = map.get(code);
      if (!code.startsWith('M-')) td.innerHTML = '<span class="price-date none">—</span>';
      else if (!item || !item.changedAt) td.innerHTML = '<span class="price-date none">не зафиксировано</span>';
      else {
        const age = daysOld(item.changedAt);
        const stale = age != null && age > 35;
        td.innerHTML = `<span class="price-date ${stale?'stale':''}" title="${esc(formatDate(item.changedAt, true))}">${esc(formatDate(item.changedAt))}${age!=null?` · ${age} дн.`:''}</span>`;
      }
      const ref = tr.children[4] || null;
      tr.insertBefore(td, ref);
    });
  }

  function observePricing() {
    const root = $('#pricingTable');
    if (!root || pricingObserver) return;
    pricingObserver = new MutationObserver(() => patchPricingTable());
    pricingObserver.observe(root, {childList:true, subtree:true});
    patchPricingTable();
  }

  async function refreshRequestStatus() {
    const box = $('#priceReqStatus');
    if (!box) return;
    box.className = 'v384-status'; box.textContent = 'Загрузка статуса…';
    try {
      const resp = await apiCall({action:'priceRequestStatus'});
      if (!resp.ok) throw new Error(resp.error || 'Статус недоступен');
      const s = resp.status || {};
      const btn = $('#priceReqBtn');
      if (btn) btn.disabled = !s.canRequest;
      box.className = `v384-status ${s.canRequest?'good':'blocked'}`;
      if (s.canRequest) {
        box.innerHTML = `<strong>Запрос разрешён</strong>${s.lastRequestAt?`Последний запрос: ${esc(formatDate(s.lastRequestAt, true))}.`: 'Ранее запросы не фиксировались.'}`;
      } else {
        box.innerHTML = `<strong>Повторный запрос заблокирован</strong>Последний запрос: ${esc(formatDate(s.lastRequestAt, true))}.<br>Следующая актуализация доступна: <b>${esc(formatDate(s.nextAllowedAt))}</b>.`;
      }
    } catch (err) {
      box.className = 'v384-status blocked';
      box.textContent = err.message;
    }
  }

  async function createPriceRequest() {
    const actor = ($('#priceReqActor')?.value || '').trim();
    const comment = ($('#priceReqComment')?.value || '').trim();
    if (!actor) { toast('Укажите инициатора запроса.', 'error'); return; }
    localStorage.setItem('sten-v384-actor', actor);
    const btn = $('#priceReqBtn'); if (btn) btn.disabled = true;
    try {
      const resp = await apiCall({action:'priceRequestCreate', actor, comment});
      if (!resp.ok) throw new Error(resp.error || 'Запрос не создан');
      toast(`Запрос ${resp.request?.id || ''} создан.`, 'success');
      if ($('#priceReqComment')) $('#priceReqComment').value = '';
      await refreshRequestStatus();
    } catch (err) {
      toast(err.message, 'error');
      await refreshRequestStatus();
    }
  }

  async function refreshChat() {
    const list = $('#chatList');
    if (!list) return;
    try {
      const resp = await apiCall({action:'chatRead', limit:150});
      if (!resp.ok) throw new Error(resp.error || 'Чат недоступен');
      const messages = resp.messages || [];
      if (!messages.length) { list.innerHTML = '<div class="v384-empty">Сообщений пока нет.</div>'; return; }
      list.innerHTML = messages.map(m => `<div class="chat-item"><div class="chat-meta"><strong>${esc(m.actor || 'Без имени')}</strong><span>${esc(formatDate(m.date, true))}</span></div><div class="chat-text">${esc(m.message || '')}</div>${m.context?`<div class="chat-context">${esc(m.context)}</div>`:''}</div>`).join('');
      list.scrollTop = list.scrollHeight;
    } catch (err) {
      list.innerHTML = `<div class="v384-empty">${esc(err.message)}</div>`;
    }
  }

  async function sendChat() {
    const actor = ($('#chatActor')?.value || '').trim();
    const message = ($('#chatMessage')?.value || '').trim();
    if (!actor || !message) { toast('Для отправки нужны имя и сообщение.', 'error'); return; }
    localStorage.setItem('sten-v384-actor', actor);
    const context = `WEB v3.8.4 · ${location.pathname}`;
    const btn = $('#chatSendBtn'); if (btn) btn.disabled = true;
    try {
      const resp = await apiCall({action:'chatWrite', actor, message, context});
      if (!resp.ok) throw new Error(resp.error || 'Сообщение не отправлено');
      if ($('#chatMessage')) $('#chatMessage').value = '';
      await refreshChat();
      toast('Сообщение добавлено в общую историю.', 'success');
    } catch (err) { toast(err.message, 'error'); }
    finally { if (btn) btn.disabled = false; }
  }

  function bind() {
    $('#priceReqBtn')?.addEventListener('click', createPriceRequest);
    $('#chatSendBtn')?.addEventListener('click', sendChat);
    $('#chatRefreshBtn')?.addEventListener('click', refreshChat);
    $('#priceReqActor')?.addEventListener('change', e => localStorage.setItem('sten-v384-actor', e.target.value.trim()));
    $('#chatActor')?.addEventListener('change', e => localStorage.setItem('sten-v384-actor', e.target.value.trim()));
    $('#syncBtn')?.addEventListener('click', () => setTimeout(() => loadBootstrap(true), 1200));
    $('#tabs')?.addEventListener('click', e => {
      const tab = e.target.closest('[data-tab]')?.dataset.tab;
      if (tab === 'requests') refreshRequestStatus();
      if (tab === 'chat') refreshChat();
    });
  }

  async function init() {
    installStyles();
    updateVersionLabels();
    addTabsAndPanels();
    bind();
    observePricing();
    const ok = await loadBootstrap(true);
    if (ok) {
      await Promise.all([refreshRequestStatus(), refreshChat()]);
    } else {
      const status = $('#priceReqStatus');
      if (status) { status.className = 'v384-status blocked'; status.textContent = 'Backend v3.8.4 ещё не развернут. Синхронизация останется в fallback.'; }
    }
    setInterval(() => {
      const active = $('.tab.active')?.dataset.tab;
      if (active === 'chat') refreshChat();
    }, 60000);
  }

  if (document.readyState === 'loading') document.addEventListener('DOMContentLoaded', init, {once:true});
  else init();
})();
