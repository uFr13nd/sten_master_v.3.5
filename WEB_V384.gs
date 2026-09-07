/**
 * СТЕН_МАСТЕР WEB v3.8.4 PATCH
 *
 * Требует существующий backend v3.8.x в этом же Apps Script проекте.
 * Перед добавлением файла переименуйте старые обработчики:
 *   doGet  -> doGetLegacy
 *   doPost -> doPostLegacy
 * Остальные функции v3.8.x оставьте без изменений.
 *
 * Исправляет hidden-form/postMessage bridge, добавляет POST bootstrap,
 * дату изменения цены в публичный bootstrap, централизованные запросы цен
 * с серверным лимитом 30 дней и общий чат с историей.
 */

const WEB_V384 = Object.freeze({
  apiVersion: '3.8.4-web',
  bridgeChannel: 'sten-master-v38',
  priceRequestSheet: 'Запросы_цен',
  chatSheet: 'Чат',
  requestCooldownDays: 30,
  maxChatMessage: 2000,
  maxActor: 200,
  maxComment: 1000
});

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || 'bootstrap');
    if (action === 'health') {
      return webJsonV384_({ok:true, version:WEB_V384.apiVersion, timestamp:new Date().toISOString()});
    }
    if (action === 'bootstrap') {
      const force = String((e && e.parameter && e.parameter.force) || '') === '1';
      return webJsonV384_(webBootstrapV384_(force));
    }
    if (typeof doGetLegacy === 'function') return doGetLegacy(e);
    throw new Error('Неизвестная команда GET: ' + action);
  } catch (err) {
    return webJsonV384_({ok:false, error:String(err && err.message || err)});
  }
}

function doPost(e) {
  const bridge = webBridgeMetaV384_(e);
  let response;
  try {
    const body = webParseBodyV384_(e);
    const action = String(body.action || '');
    if (!action) throw new Error('Не указано действие.');
    const ss = SpreadsheetApp.openById(WEB_V38.spreadsheetId);

    if (action === 'bootstrap') {
      response = webBootstrapV384_(String(body.force || '') === '1' || body.force === true);
    } else if (action === 'priceRequestStatus') {
      response = {ok:true, status:webPriceRequestStatusV384_(ss)};
    } else if (action === 'priceRequestCreate') {
      response = webCreatePriceRequestV384_(ss, body);
    } else if (action === 'chatRead') {
      response = {ok:true, messages:webReadChatV384_(ss, Number(body.limit || 100))};
    } else if (action === 'chatWrite') {
      response = webWriteChatV384_(ss, body);
    } else if (action === 'updateReferencesBatch') {
      response = webUpdateReferencesBatchV384_(ss, body);
    } else {
      response = webLegacyPostV384_(body);
      response = webPostProcessV384_(ss, response);
    }
  } catch (err) {
    response = {ok:false, error:String(err && err.message || err)};
  }

  return bridge.enabled ? webBridgeResponseV384_(bridge, response) : webJsonV384_(response);
}

function webParseBodyV384_(e) {
  const p = (e && e.parameter) || {};
  if (p.payload) {
    try { return JSON.parse(String(p.payload)); }
    catch (err) { throw new Error('Некорректный payload Web API.'); }
  }
  const raw = e && e.postData && e.postData.contents ? e.postData.contents : '';
  if (raw) {
    try { return JSON.parse(raw); } catch (ignore) {}
  }
  return p;
}

function webBridgeMetaV384_(e) {
  const p = (e && e.parameter) || {};
  return {
    enabled: String(p.bridge || '') === '1',
    id: String(p.requestId || ''),
    origin: webSafeOriginV384_(String(p.origin || '*')),
    channel: String(p.bridgeChannel || WEB_V384.bridgeChannel)
  };
}

function webSafeOriginV384_(origin) {
  if (origin === '*') return '*';
  try {
    const u = new URL(origin);
    if (u.protocol !== 'https:' && u.protocol !== 'http:') return '*';
    return u.origin;
  } catch (err) {
    return '*';
  }
}

function webBridgeResponseV384_(meta, obj) {
  const message = {
    channel: meta.channel || WEB_V384.bridgeChannel,
    id: meta.id,
    response: obj
  };
  const json = JSON.stringify(message)
    .replace(/</g, '\\u003c')
    .replace(/>/g, '\\u003e')
    .replace(/&/g, '\\u0026');
  const target = JSON.stringify(meta.origin || '*');
  const html = '<!doctype html><html><head><meta charset="utf-8"></head><body>' +
    '<script>(function(){var msg=' + json + ';window.parent.postMessage(msg,' + target + ');})();<\/script>' +
    '</body></html>';
  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function webJsonV384_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function webLegacyPostV384_(body) {
  if (typeof doPostLegacy !== 'function') {
    throw new Error('Не найден doPostLegacy. Переименуйте старый doPost в doPostLegacy перед установкой WEB_V384.gs.');
  }
  const legacyEvent = {
    postData: {contents: JSON.stringify(body), type:'application/json'},
    parameter: {}
  };
  const out = doPostLegacy(legacyEvent);
  if (!out || typeof out.getContent !== 'function') {
    throw new Error('Legacy backend вернул неподдерживаемый ответ.');
  }
  const text = out.getContent();
  try { return JSON.parse(text); }
  catch (err) { throw new Error('Legacy backend вернул не JSON: ' + String(text).slice(0,300)); }
}

function webPostProcessV384_(ss, response) {
  const out = response && typeof response === 'object' ? response : {ok:false,error:'Пустой ответ legacy backend'};
  if (out.data && out.data.materials && out.data.works) out.data = webEnrichBootstrapV384_(ss, out.data);
  return out;
}

function webBootstrapV384_(force) {
  const ss = SpreadsheetApp.openById(WEB_V38.spreadsheetId);
  const base = webBootstrap_(!!force);
  return webEnrichBootstrapV384_(ss, base);
}

function webEnrichBootstrapV384_(ss, base) {
  const data = JSON.parse(JSON.stringify(base || {}));
  data.ok = data.ok !== false;
  data.apiVersion = WEB_V384.apiVersion;
  data.priceRequestStatus = webPriceRequestStatusV384_(ss);
  const meta = webReferencePriceMetaV384_(ss);
  (data.works || []).forEach(item => {
    const x = meta.works[item.code];
    if (!x) return;
    item.previousPrice = x.previousPrice;
    item.changedAt = x.changedAt;
  });
  (data.materials || []).forEach(item => {
    const x = meta.materials[item.code];
    if (!x) return;
    item.previousPrice = x.previousPrice;
    item.changedAt = x.changedAt;
  });
  return data;
}

function webReferencePriceMetaV384_(ss) {
  return {
    works: webReadPriceMetaSheetV384_(ss.getSheetByName(STEN_V35.workSheet)),
    materials: webReadPriceMetaSheetV384_(ss.getSheetByName(STEN_V35.materialSheet))
  };
}

function webReadPriceMetaSheetV384_(sheet) {
  const out = {};
  if (!sheet || sheet.getLastRow() < 2) return out;
  const rows = sheet.getRange(2, 1, sheet.getLastRow() - 1, 7).getValues();
  rows.forEach(r => {
    const code = String(r[0] || '').trim();
    if (!code) return;
    out[code] = {
      previousPrice: r[4] === '' || r[4] == null ? null : r[4],
      changedAt: webIsoDateV384_(r[5])
    };
  });
  return out;
}

function webIsoDateV384_(value) {
  if (!value) return '';
  if (value instanceof Date && !isNaN(value.getTime())) return value.toISOString();
  const d = new Date(value);
  return isNaN(d.getTime()) ? String(value) : d.toISOString();
}

function webEnsurePriceRequestSheetV384_(ss) {
  let sh = ss.getSheetByName(WEB_V384.priceRequestSheet);
  if (!sh) sh = ss.insertSheet(WEB_V384.priceRequestSheet);
  const headers = ['ID','Дата запроса','Пользователь','Статус','Комментарий','Дата выполнения','Следующая актуализация'];
  sh.getRange(1,1,1,headers.length).setValues([headers]);
  sh.setFrozenRows(1);
  return sh;
}

function webPriceRequestStatusV384_(ss) {
  const sh = webEnsurePriceRequestSheetV384_(ss);
  const lastRow = sh.getLastRow();
  let last = null;
  if (lastRow >= 2) {
    const rows = sh.getRange(2,1,lastRow-1,7).getValues();
    for (let i = rows.length - 1; i >= 0; i--) {
      const r = rows[i];
      const status = String(r[3] || '').trim().toLowerCase();
      if (status === 'отменён' || status === 'отменен') continue;
      const date = r[1] instanceof Date ? r[1] : new Date(r[1]);
      if (date && !isNaN(date.getTime())) {
        last = {id:String(r[0]||''), date, actor:String(r[2]||''), status:String(r[3]||''), comment:String(r[4]||'')};
        break;
      }
    }
  }
  const now = new Date();
  if (!last) return {canRequest:true,lastRequestAt:'',nextAllowedAt:''};
  const next = new Date(last.date.getTime() + WEB_V384.requestCooldownDays * 86400000);
  return {
    canRequest: now.getTime() >= next.getTime(),
    lastRequestId: last.id,
    lastRequestAt: last.date.toISOString(),
    lastActor: last.actor,
    lastStatus: last.status,
    nextAllowedAt: next.toISOString(),
    cooldownDays: WEB_V384.requestCooldownDays
  };
}

function webCreatePriceRequestV384_(ss, body) {
  const actor = webRequiredV384_(body.actor, 'Инициатор', WEB_V384.maxActor);
  const comment = String(body.comment || '').trim().slice(0, WEB_V384.maxComment);
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const status = webPriceRequestStatusV384_(ss);
    if (!status.canRequest) {
      const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
      const next = Utilities.formatDate(new Date(status.nextAllowedAt), tz, 'dd.MM.yyyy');
      throw new Error('Актуализацию уже запрашивали. Следующий запрос разрешён ' + next + '.');
    }
    const sh = webEnsurePriceRequestSheetV384_(ss);
    const now = new Date();
    const next = new Date(now.getTime() + WEB_V384.requestCooldownDays * 86400000);
    const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
    const id = 'PR-' + Utilities.formatDate(now, tz, 'yyyyMMdd-HHmmss') + '-' + Math.floor(Math.random()*900+100);
    sh.appendRow([id, now, actor, 'Новый', comment, '', next]);
    const row = sh.getLastRow();
    sh.getRange(row,2).setNumberFormat('dd.MM.yyyy HH:mm:ss');
    sh.getRange(row,7).setNumberFormat('dd.MM.yyyy');
    return {
      ok:true,
      request:{id:id,date:now.toISOString(),actor:actor,status:'Новый',comment:comment,nextAllowedAt:next.toISOString()},
      status:webPriceRequestStatusV384_(ss)
    };
  } finally {
    lock.releaseLock();
  }
}

function webEnsureChatSheetV384_(ss) {
  let sh = ss.getSheetByName(WEB_V384.chatSheet);
  if (!sh) sh = ss.insertSheet(WEB_V384.chatSheet);
  sh.getRange(1,1,1,4).setValues([['Дата/время','Пользователь','Сообщение','Контекст']]);
  sh.setFrozenRows(1);
  return sh;
}

function webReadChatV384_(ss, limit) {
  const sh = webEnsureChatSheetV384_(ss);
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  const safeLimit = Math.max(1, Math.min(Number(limit) || 100, 300));
  const start = Math.max(2, lastRow - safeLimit + 1);
  const rows = sh.getRange(start,1,lastRow-start+1,4).getValues();
  return rows.map(r => ({
    date:webIsoDateV384_(r[0]),
    actor:String(r[1] || ''),
    message:String(r[2] || ''),
    context:String(r[3] || '')
  })).filter(x => x.message);
}

function webWriteChatV384_(ss, body) {
  const actor = webRequiredV384_(body.actor, 'Имя / email', WEB_V384.maxActor);
  const message = webRequiredV384_(body.message, 'Сообщение', WEB_V384.maxChatMessage);
  const context = String(body.context || '').trim().slice(0,500);
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const sh = webEnsureChatSheetV384_(ss);
    const now = new Date();
    sh.appendRow([now, actor, message, context]);
    sh.getRange(sh.getLastRow(),1).setNumberFormat('dd.MM.yyyy HH:mm:ss');
    return {ok:true, message:{date:now.toISOString(),actor:actor,message:message,context:context}};
  } finally {
    lock.releaseLock();
  }
}

function webUpdateReferencesBatchV384_(ss, body) {
  webAssertAdmin_(body.token);
  const actor = webActor_(body.actor);
  const changes = Array.isArray(body.changes) ? body.changes.slice(0,500) : [];
  if (!changes.length) return {ok:true,updatedCount:0,data:webBootstrapV384_(true),audit:webReadAudit_(ss,50)};

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const bySheet = {};
    [STEN_V35.workSheet, STEN_V35.materialSheet].forEach(name => {
      const sh = ss.getSheetByName(name);
      if (!sh || sh.getLastRow() < 2) return;
      const codes = sh.getRange(2,1,sh.getLastRow()-1,1).getDisplayValues();
      const map = {};
      codes.forEach((r,i) => { const code=String(r[0]||'').trim(); if(code) map[code]=i+2; });
      bySheet[name] = {sheet:sh, rows:map};
    });

    const auditRows = [];
    const now = new Date();
    let updated = 0;
    changes.forEach(ch => {
      const sheetName = ch.sheet === STEN_V35.workSheet ? STEN_V35.workSheet : ch.sheet === STEN_V35.materialSheet ? STEN_V35.materialSheet : '';
      const holder = bySheet[sheetName];
      const code = String(ch.code || '').trim();
      if (!holder || !holder.rows[code]) return;
      const row = holder.rows[code];
      const field = String(ch.field || '');
      let col = 0, fieldLabel = '';
      if (field === 'price') { col=4; fieldLabel='Цена'; }
      else if (field === 'technicalGroup') { col=8; fieldLabel=sheetName===STEN_V35.materialSheet?'Техническая группа':'Группа'; }
      else if (field === 'purchaseGroup' && sheetName===STEN_V35.materialSheet) { col=11; fieldLabel='Группа закупки'; }
      else return;

      const cell = holder.sheet.getRange(row,col);
      const oldValue = cell.getValue();
      const newValue = field === 'price' ? webNum_(ch.value) : String(ch.value == null ? '' : ch.value).trim();
      if (valuesEqual_(oldValue,newValue)) return;
      const name = String(holder.sheet.getRange(row,2).getValue() || '');
      cell.setValue(newValue);
      if (field === 'price') {
        holder.sheet.getRange(row,5).setValue(oldValue);
        holder.sheet.getRange(row,6).setValue(now).setNumberFormat('dd.MM.yyyy HH:mm:ss');
        holder.sheet.getRange(row,7).setValue(actor);
      }
      auditRows.push([now,actor,sheetName,code,name,oldValue,newValue,cell.getA1Notation(),fieldLabel]);
      updated++;
    });

    if (auditRows.length) {
      const log = ensureLogSheet_(ss);
      const start = log.getLastRow()+1;
      log.getRange(start,1,auditRows.length,9).setValues(auditRows);
      log.getRange(start,1,auditRows.length,1).setNumberFormat('dd.MM.yyyy HH:mm:ss');
    }

    if (updated) {
      synchronizeBasesCore_(ss,false);
      if (typeof buildStrategicReportCore_ === 'function') buildStrategicReportCore_(ss,false);
      if (typeof rebuildAuditSnapshot_ === 'function') rebuildAuditSnapshot_(ss, ensureSnapshotSheet_(ss));
      SpreadsheetApp.flush();
      webClearCache_();
    }
    return {ok:true,updatedCount:updated,data:webBootstrapV384_(true),audit:webReadAudit_(ss,50)};
  } finally {
    lock.releaseLock();
  }
}

function webRequiredV384_(value, label, max) {
  const text = String(value == null ? '' : value).trim();
  if (!text) throw new Error('Заполните поле «' + label + '».');
  return text.slice(0, max || 500);
}
