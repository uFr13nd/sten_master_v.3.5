/**
 * СТЕН_МАСТЕР v3.5 — единый пакет контроля справочников и баз.
 *
 * Возможности:
 *  1) аудит изменений цены / технической группы / группы закупки;
 *  2) синхронизация кодов, типов позиций, групп закупки и «Где используется»;
 *  3) проверка целостности формул, кодов, дублей и классификации;
 *  4) отчёт по стратегическим материалам;
 *  5) присвоение кодов новым строкам;
 *  6) резервная копия книги из меню.
 *
 * После полной замены старого кода этим файлом один раз запустите setupSystemV35().
 */

const STEN_V35 = Object.freeze({
  version: '3.5',
  workSheet: 'Работы',
  materialSheet: 'Материалы',
  logSheet: 'Журнал_цен',
  snapshotSheet: '_AUDIT_V35_SNAPSHOT',
  controlSheet: 'Контроль_системы',
  strategicSheet: 'Отчёт_Стратегические',
  systemSheet: 'Система_v3.5',
  baseSheets: ['База_Наружные', 'База_Межквартирные', 'База_Перегородки', 'База_Отделка'],
  firstDataRow: 2,
  codeColumn: 1,
  nameColumn: 2,
  priceColumn: 4,
  previousPriceColumn: 5,
  changedAtColumn: 6,
  changedByColumn: 7,
  workGroupColumn: 8,
  materialTechGroupColumn: 8,
  usageColumn: 9,
  commentColumn: 10,
  materialPurchaseGroupColumn: 11,
  baseCodeColumn: 15,       // O
  baseTypeColumn: 16,       // P
  basePurchaseGroupColumn: 17, // Q
  maxAuditIssues: 1000,
  tracked: {
    'Работы': {
      4: 'Цена',
      8: 'Группа'
    },
    'Материалы': {
      4: 'Цена',
      8: 'Техническая группа',
      11: 'Группа закупки'
    }
  }
});

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('СТЕН_МАСТЕР')
    .addItem('🔄 Синхронизировать базы', 'synchronizeBases')
    .addItem('✅ Проверить целостность', 'checkSystemIntegrity')
    .addItem('⭐ Отчёт по стратегическим материалам', 'buildStrategicReport')
    .addItem('🔎 Найти дубли и ошибки', 'findDuplicatesAndErrors')
    .addSeparator()
    .addItem('🆔 Присвоить коды новым строкам', 'assignCodesToNewRows')
    .addItem('📋 Открыть журнал изменений', 'openAuditLog')
    .addItem('🧪 Открыть контроль системы', 'openControlSheet')
    .addSeparator()
    .addItem('💾 Создать резервную копию', 'createSpreadsheetBackup')
    .addItem('⚙ Настроить / обновить систему v3.5', 'setupSystemV35')
    .addItem('🌐 Настроить Web API v3.8', 'setupWebApi')
    .addToUi();
}

/** Запустить один раз вручную после полной замены старого кода. */
function setupSystemV35() {
  const ss = SpreadsheetApp.getActive();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    ensureLogSheet_(ss);
    const snapshot = ensureSnapshotSheet_(ss);
    ensureOutputSheet_(ss, STEN_V35.controlSheet);
    ensureOutputSheet_(ss, STEN_V35.strategicSheet);

    rebuildAuditSnapshot_(ss, snapshot);
    snapshot.hideSheet();

    // Удаляем только триггеры нашего старого и нового аудита.
    ScriptApp.getProjectTriggers().forEach(trigger => {
      const handler = trigger.getHandlerFunction();
      if (trigger.getEventType() === ScriptApp.EventType.ON_EDIT &&
          ['priceAuditOnEdit', 'auditOnEditV35'].includes(handler)) {
        ScriptApp.deleteTrigger(trigger);
      }
    });

    ScriptApp.newTrigger('auditOnEditV35')
      .forSpreadsheet(ss)
      .onEdit()
      .create();

    synchronizeBasesCore_(ss, false);
    buildStrategicReportCore_(ss, false);
    checkSystemIntegrityCore_(ss, false);

    ss.toast('Система v3.5 настроена: аудит, синхронизация, контроль и стратегический отчёт активированы.', 'СТЕН_МАСТЕР', 8);
  } finally {
    lock.releaseLock();
  }
}

/** Installable onEdit. Не запускать вручную. */
function auditOnEditV35(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const trackedColumns = STEN_V35.tracked[sheetName];
  if (!trackedColumns) return;

  const firstRow = Math.max(e.range.getRow(), STEN_V35.firstDataRow);
  const lastRow = e.range.getLastRow();
  if (lastRow < STEN_V35.firstDataRow) return;

  const editedFirstCol = e.range.getColumn();
  const editedLastCol = e.range.getLastColumn();
  const relevantColumns = Object.keys(trackedColumns)
    .map(Number)
    .filter(col => col >= editedFirstCol && col <= editedLastCol);

  if (!relevantColumns.length) return;

  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const ss = e.source || SpreadsheetApp.getActive();
    const log = ensureLogSheet_(ss);
    const snapshot = ensureSnapshotSheet_(ss);
    const snapshotData = readAuditSnapshot_(snapshot);
    const actor = getActor_(e);
    const now = new Date();
    const logRows = [];
    const snapshotUpdates = [];
    const purchaseGroupChanges = [];

    const rowCount = lastRow - firstRow + 1;
    const codes = sheet.getRange(firstRow, STEN_V35.codeColumn, rowCount, 1).getDisplayValues();
    const names = sheet.getRange(firstRow, STEN_V35.nameColumn, rowCount, 1).getDisplayValues();

    relevantColumns.forEach(col => {
      const fieldName = trackedColumns[col];
      const values = sheet.getRange(firstRow, col, rowCount, 1).getValues();

      for (let i = 0; i < rowCount; i++) {
        const row = firstRow + i;
        const code = String(codes[i][0] || '').trim();
        const name = String(names[i][0] || '').trim();
        if (!code) continue;

        const newValue = values[i][0];
        const key = makeSnapshotKey_(sheetName, code, fieldName);
        let oldValue = snapshotData.values.has(key) ? snapshotData.values.get(key) : '';

        // Для первого одиночного изменения можно использовать oldValue события.
        if (!snapshotData.values.has(key) &&
            rowCount === 1 && relevantColumns.length === 1 &&
            e.oldValue !== undefined) {
          oldValue = e.oldValue;
        }

        if (valuesEqual_(oldValue, newValue)) continue;

        const a1 = sheet.getRange(row, col).getA1Notation();
        logRows.push([now, actor, sheetName, code, name, oldValue, newValue, a1, fieldName]);
        snapshotUpdates.push({ key: key, value: newValue });

        if (fieldName === 'Цена') {
          sheet.getRange(row, STEN_V35.previousPriceColumn).setValue(oldValue);
          sheet.getRange(row, STEN_V35.changedAtColumn).setValue(now).setNumberFormat('dd.MM.yyyy HH:mm:ss');
          sheet.getRange(row, STEN_V35.changedByColumn).setValue(actor);
        }

        if (sheetName === STEN_V35.materialSheet && fieldName === 'Группа закупки') {
          purchaseGroupChanges.push({ code: code, group: String(newValue || '').trim() });
        }
      }
    });

    if (logRows.length) {
      const start = log.getLastRow() + 1;
      log.getRange(start, 1, logRows.length, 9).setValues(logRows);
      log.getRange(start, 1, logRows.length, 1).setNumberFormat('dd.MM.yyyy HH:mm:ss');
    }

    applySnapshotUpdates_(snapshot, snapshotData, snapshotUpdates);

    purchaseGroupChanges.forEach(change => {
      updateMaterialGroupInBases_(ss, change.code, change.group);
    });
  } finally {
    lock.releaseLock();
  }
}

/** Публичная команда меню. */
function synchronizeBases() {
  const ss = SpreadsheetApp.getActive();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    synchronizeBasesCore_(ss, true);
  } finally {
    lock.releaseLock();
  }
}

function synchronizeBasesCore_(ss, showToast) {
  const refs = readReferenceMaps_(ss);
  const usage = new Map();
  refs.allCodes.forEach(code => usage.set(code, []));

  STEN_V35.baseSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 1) return;

    const scanCols = Math.min(14, Math.max(1, sheet.getLastColumn())); // A:N
    const formulas = sheet.getRange(1, 1, lastRow, scanCols).getFormulas();
    const displays = sheet.getRange(1, 1, lastRow, scanCols).getDisplayValues();
    const currentClass = sheet.getRange(1, STEN_V35.baseCodeColumn, lastRow, 3).getDisplayValues();
    const out = [];

    for (let r = 0; r < lastRow; r++) {
      const rowNumber = r + 1;
      if (r === 0) {
        out.push(['Код справочника', 'Тип позиции', 'Группа закупки']);
        continue;
      }
      const foundCodes = [];

      for (let c = 0; c < scanCols; c++) {
        const formula = String(formulas[r][c] || '');
        if (!formula) continue;
        const codes = extractCodes_(formula);
        codes.forEach(code => {
          if (!foundCodes.includes(code)) foundCodes.push(code);
          if (!usage.has(code)) usage.set(code, []);
          usage.get(code).push(`${sheetName}!${columnToLetter_(c + 1)}${rowNumber}`);
        });
      }

      let code = foundCodes.length ? foundCodes[0] : '';
      let type = currentClass[r] ? String(currentClass[r][1] || '') : '';
      let purchase = '';

      if (code) {
        type = code.indexOf('M-') === 0 ? 'Материал' : 'Работа';
        if (type === 'Материал') purchase = refs.materialPurchase.get(code) || '';
      } else {
        type = inferBaseRowType_(displays[r], type, sheetName);
      }

      out.push([code, type, purchase]);
    }

    sheet.getRange(1, STEN_V35.baseCodeColumn, lastRow, 3).setValues(out);
  });

  writeUsageMap_(ss, refs, usage);

  if (showToast) {
    ss.toast('Коды, типы позиций, группы закупки и «Где используется» синхронизированы.', 'СТЕН_МАСТЕР', 6);
  }
}

function checkSystemIntegrity() {
  const ss = SpreadsheetApp.getActive();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    checkSystemIntegrityCore_(ss, true);
  } finally {
    lock.releaseLock();
  }
}

function findDuplicatesAndErrors() {
  checkSystemIntegrity();
  openControlSheet();
}

function checkSystemIntegrityCore_(ss, showToast) {
  const issues = [];
  const refs = readReferenceMaps_(ss);
  const actualUsage = new Map();
  refs.allCodes.forEach(code => actualUsage.set(code, []));

  const required = [
    STEN_V35.workSheet,
    STEN_V35.materialSheet,
    STEN_V35.logSheet,
    STEN_V35.snapshotSheet
  ].concat(STEN_V35.baseSheets);

  required.forEach(name => {
    if (!ss.getSheetByName(name)) {
      addIssue_(issues, 'ОШИБКА', 'Структура', name, '', 'Отсутствует обязательный лист.');
    }
  });

  validateReferenceSheet_(ss, STEN_V35.workSheet, 'W-', issues);
  validateReferenceSheet_(ss, STEN_V35.materialSheet, 'M-', issues);
  validateDuplicateNames_(ss, STEN_V35.workSheet, issues);
  validateDuplicateNames_(ss, STEN_V35.materialSheet, issues);

  STEN_V35.baseSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    const lastCol = Math.min(Math.max(sheet.getLastColumn(), 17), 30);
    if (lastRow < 1) return;

    const range = sheet.getRange(1, 1, lastRow, lastCol);
    const formulas = range.getFormulas();
    const display = range.getDisplayValues();

    for (let r = 0; r < lastRow; r++) {
      const rowCodes = [];
      for (let c = 0; c < Math.min(14, lastCol); c++) {
        extractCodes_(formulas[r][c]).forEach(code => {
          if (!rowCodes.includes(code)) rowCodes.push(code);
        });
      }
      if (r > 0 && lastCol >= STEN_V35.basePurchaseGroupColumn) {
        const classifiedCode = String(display[r][STEN_V35.baseCodeColumn - 1] || '').trim();
        const classifiedType = String(display[r][STEN_V35.baseTypeColumn - 1] || '').trim();
        const actualCode = rowCodes.length ? rowCodes[0] : '';
        if (actualCode && classifiedCode !== actualCode) {
          addIssue_(issues, 'ОШИБКА', 'Классификация', sheetName, `O${r + 1}`, `Код классификации «${classifiedCode}» не совпадает с кодом в формуле «${actualCode}». Выполните синхронизацию.`);
        }
        if (!actualCode && ['Материал', 'Работа'].includes(classifiedType)) {
          addIssue_(issues, 'ОШИБКА', 'Связь со справочником', sheetName, `P${r + 1}`, `Строка помечена как «${classifiedType}», но в формулах A:N нет кода M-/W-. Возможна ручная цена или потерянная ссылка на справочник.`);
        }
      }
      for (let c = 0; c < lastCol; c++) {
        const value = String(display[r][c] || '');
        if (/^#(REF!|N\/A|DIV\/0!|VALUE!|NAME\?|ERROR!|NUM!)/i.test(value)) {
          addIssue_(issues, 'ОШИБКА', 'Формула', sheetName, `${columnToLetter_(c + 1)}${r + 1}`, `Ошибка формулы: ${value}`);
        }

        const formula = String(formulas[r][c] || '');
        if (!formula) continue;
        extractCodes_(formula).forEach(code => {
          if (!actualUsage.has(code)) actualUsage.set(code, []);
          actualUsage.get(code).push(`${sheetName}!${columnToLetter_(c + 1)}${r + 1}`);

          if (code.indexOf('M-') === 0 && !refs.materialCodes.has(code)) {
            addIssue_(issues, 'ОШИБКА', 'Код', sheetName, `${columnToLetter_(c + 1)}${r + 1}`, `Формула ссылается на отсутствующий материал ${code}.`);
          }
          if (code.indexOf('W-') === 0 && !refs.workCodes.has(code)) {
            addIssue_(issues, 'ОШИБКА', 'Код', sheetName, `${columnToLetter_(c + 1)}${r + 1}`, `Формула ссылается на отсутствующую работу ${code}.`);
          }
        });
      }
    }

    // Проверяем O:P:Q.
    if (lastCol >= STEN_V35.basePurchaseGroupColumn) {
      const classData = sheet.getRange(2, STEN_V35.baseCodeColumn, Math.max(lastRow - 1, 1), 3).getDisplayValues();
      classData.forEach((row, idx) => {
        const code = String(row[0] || '').trim();
        if (!code) return;
        const type = String(row[1] || '').trim();
        const purchase = String(row[2] || '').trim();
        const sheetRow = idx + 2;

        if (code.indexOf('M-') === 0) {
          if (type !== 'Материал') {
            addIssue_(issues, 'ОШИБКА', 'Классификация', sheetName, `P${sheetRow}`, `${code}: ожидался тип «Материал», сейчас «${type}».`);
          }
          const expected = refs.materialPurchase.get(code) || '';
          if (purchase !== expected) {
            addIssue_(issues, 'ОШИБКА', 'Классификация', sheetName, `Q${sheetRow}`, `${code}: группа закупки «${purchase}», в справочнике «${expected}».`);
          }
        } else if (code.indexOf('W-') === 0 && type !== 'Работа') {
          addIssue_(issues, 'ОШИБКА', 'Классификация', sheetName, `P${sheetRow}`, `${code}: ожидался тип «Работа», сейчас «${type}».`);
        }
      });
    }
  });

  // Неиспользуемые позиции.
  refs.allCodes.forEach(code => {
    const refsForCode = actualUsage.get(code) || [];
    if (!refsForCode.length) {
      const sheetName = code.indexOf('M-') === 0 ? STEN_V35.materialSheet : STEN_V35.workSheet;
      addIssue_(issues, 'ПРЕДУПРЕЖДЕНИЕ', 'Использование', sheetName, code, 'Код не используется ни одной формулой в рабочих базах.');
    }
  });

  // Сравниваем «Где используется» с фактом.
  validateUsageColumn_(ss, STEN_V35.workSheet, refs.workCodes, actualUsage, issues);
  validateUsageColumn_(ss, STEN_V35.materialSheet, refs.materialCodes, actualUsage, issues);

  if (!issues.some(row => row[0] === 'ОШИБКА')) {
    addIssue_(issues, 'OK', 'Итог', '', '', 'Критических ошибок целостности не обнаружено.');
  }

  writeControlReport_(ss, issues);

  if (showToast) {
    const errors = issues.filter(row => row[0] === 'ОШИБКА').length;
    const warnings = issues.filter(row => row[0] === 'ПРЕДУПРЕЖДЕНИЕ').length;
    ss.toast(`Проверка завершена: ошибок ${errors}, предупреждений ${warnings}.`, 'СТЕН_МАСТЕР', 7);
  }
}

function buildStrategicReport() {
  const ss = SpreadsheetApp.getActive();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    buildStrategicReportCore_(ss, true);
    ss.setActiveSheet(ss.getSheetByName(STEN_V35.strategicSheet));
  } finally {
    lock.releaseLock();
  }
}

function buildStrategicReportCore_(ss, showToast) {
  const src = ss.getSheetByName(STEN_V35.materialSheet);
  if (!src) throw new Error('Не найден лист «Материалы».');

  const sheet = ensureOutputSheet_(ss, STEN_V35.strategicSheet);
  const lastRow = src.getLastRow();
  const rows = lastRow >= 2 ? src.getRange(2, 1, lastRow - 1, 11).getValues() : [];

  const output = [];
  rows.forEach(row => {
    const code = String(row[0] || '').trim();
    const purchase = String(row[10] || '').trim();
    if (!code || purchase !== 'Стратегические материалы') return;

    const usage = String(row[8] || '').trim();
    const usageCount = usage ? usage.split(';').map(s => s.trim()).filter(Boolean).length : 0;
    output.push([
      code,
      String(row[7] || ''),
      String(row[1] || ''),
      String(row[2] || ''),
      row[3],
      usageCount,
      usage,
      row[4],
      row[5],
      row[6],
      row[9]
    ]);
  });

  output.sort((a, b) => {
    const ga = String(a[1] || '').toLowerCase();
    const gb = String(b[1] || '').toLowerCase();
    if (ga < gb) return -1;
    if (ga > gb) return 1;
    return String(a[2] || '').localeCompare(String(b[2] || ''), 'ru');
  });

  sheet.clear();
  sheet.getRange('A1:K1').merge();
  sheet.getRange('A1').setValue('СТРАТЕГИЧЕСКИЕ МАТЕРИАЛЫ — СВОДНЫЙ ОТЧЁТ');
  sheet.getRange('A2').setValue('Обновлено');
  sheet.getRange('B2').setValue(new Date()).setNumberFormat('dd.MM.yyyy HH:mm:ss');
  sheet.getRange('D2').setValue('Количество');
  sheet.getRange('E2').setValue(output.length);

  const headers = ['Код', 'Техническая группа', 'Наименование', 'Ед. изм.', 'Цена, ₽/ед.', 'Использований', 'Где используется', 'Предыдущая цена', 'Дата изменения', 'Изменил', 'Комментарий'];
  sheet.getRange(4, 1, 1, headers.length).setValues([headers]);
  if (output.length) {
    sheet.getRange(5, 1, output.length, headers.length).setValues(output);
    sheet.getRange(5, 5, output.length, 1).setNumberFormat('#,##0.00 ₽');
    sheet.getRange(5, 8, output.length, 1).setNumberFormat('#,##0.00 ₽');
    sheet.getRange(5, 9, output.length, 1).setNumberFormat('dd.MM.yyyy HH:mm:ss');
  }

  formatStrategicSheet_(sheet, Math.max(output.length + 4, 5));

  if (showToast) {
    ss.toast(`Отчёт обновлён: ${output.length} стратегических материалов.`, 'СТЕН_МАСТЕР', 6);
  }
}

function assignCodesToNewRows() {
  const ss = SpreadsheetApp.getActive();
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const changes = [];
    changes.push(assignCodesOnSheet_(ss, STEN_V35.workSheet, 'W-'));
    changes.push(assignCodesOnSheet_(ss, STEN_V35.materialSheet, 'M-'));

    const snapshot = ensureSnapshotSheet_(ss);
    rebuildAuditSnapshot_(ss, snapshot);
    synchronizeBasesCore_(ss, false);

    const total = changes.reduce((sum, n) => sum + n, 0);
    ss.toast(`Присвоено новых кодов: ${total}.`, 'СТЕН_МАСТЕР', 6);
  } finally {
    lock.releaseLock();
  }
}

function createSpreadsheetBackup() {
  const ss = SpreadsheetApp.getActive();
  const file = DriveApp.getFileById(ss.getId());
  const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'GMT';
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd_HH-mm-ss');
  const name = `${ss.getName()}_BACKUP_${stamp}`;

  const parents = file.getParents();
  if (parents.hasNext()) {
    file.makeCopy(name, parents.next());
  } else {
    file.makeCopy(name);
  }
  ss.toast(`Создана резервная копия: ${name}`, 'СТЕН_МАСТЕР', 7);
}

function openAuditLog() {
  activateSheet_(STEN_V35.logSheet);
}

function openControlSheet() {
  activateSheet_(STEN_V35.controlSheet);
}

function openStrategicReport() {
  activateSheet_(STEN_V35.strategicSheet);
}

// =========================
// ВНУТРЕННИЕ ФУНКЦИИ
// =========================

function readReferenceMaps_(ss) {
  const workSheet = ss.getSheetByName(STEN_V35.workSheet);
  const materialSheet = ss.getSheetByName(STEN_V35.materialSheet);
  if (!workSheet || !materialSheet) throw new Error('Не найдены листы «Работы» и/или «Материалы».');

  const workCodes = new Map();
  const materialCodes = new Map();
  const materialPurchase = new Map();
  const allCodes = new Set();

  if (workSheet.getLastRow() >= 2) {
    const data = workSheet.getRange(2, 1, workSheet.getLastRow() - 1, 10).getValues();
    data.forEach((row, i) => {
      const code = String(row[0] || '').trim();
      if (!code) return;
      workCodes.set(code, { row: i + 2, name: row[1], price: row[3], usage: row[8] });
      allCodes.add(code);
    });
  }

  if (materialSheet.getLastRow() >= 2) {
    const data = materialSheet.getRange(2, 1, materialSheet.getLastRow() - 1, 11).getValues();
    data.forEach((row, i) => {
      const code = String(row[0] || '').trim();
      if (!code) return;
      materialCodes.set(code, { row: i + 2, name: row[1], price: row[3], usage: row[8], techGroup: row[7], purchaseGroup: row[10] });
      materialPurchase.set(code, String(row[10] || '').trim());
      allCodes.add(code);
    });
  }

  return { workCodes, materialCodes, materialPurchase, allCodes };
}

function writeUsageMap_(ss, refs, usage) {
  const targets = [
    { sheetName: STEN_V35.workSheet, map: refs.workCodes },
    { sheetName: STEN_V35.materialSheet, map: refs.materialCodes }
  ];

  targets.forEach(target => {
    const sheet = ss.getSheetByName(target.sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;
    const count = sheet.getLastRow() - 1;
    const codes = sheet.getRange(2, 1, count, 1).getDisplayValues();
    const out = codes.map(row => {
      const code = String(row[0] || '').trim();
      const refsForCode = usage.get(code) || [];
      const unique = Array.from(new Set(refsForCode));
      return [unique.join('; ')];
    });
    sheet.getRange(2, STEN_V35.usageColumn, count, 1).setValues(out).setWrap(true);
  });
}

function updateMaterialGroupInBases_(ss, code, group) {
  STEN_V35.baseSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < 2) return;
    const count = sheet.getLastRow() - 1;
    const codes = sheet.getRange(2, STEN_V35.baseCodeColumn, count, 1).getDisplayValues();
    const q = sheet.getRange(2, STEN_V35.basePurchaseGroupColumn, count, 1).getValues();
    let changed = false;
    for (let i = 0; i < count; i++) {
      if (String(codes[i][0] || '').trim() === code) {
        q[i][0] = group;
        changed = true;
      }
    }
    if (changed) sheet.getRange(2, STEN_V35.basePurchaseGroupColumn, count, 1).setValues(q);
  });
}

function inferBaseRowType_(rowValues, currentType, sheetName) {
  const joined = rowValues.map(v => String(v || '')).join(' ').replace(/\s+/g, ' ').trim();
  if (!joined) return currentType || '';

  if (/ИТОГО|Итого:/i.test(joined)) return 'Итог';
  if (/МАТЕРИАЛЫ\s*[—-]\s*ОСНОВНЫЕ|Основные материалы/i.test(joined)) return 'Группа материалов — основные';
  if (/МАТЕРИАЛЫ\s*[—-]\s*СОПУТСТВУЮЩИЕ|Сопутствующие материалы/i.test(joined)) return 'Группа материалов — сопутствующие';
  if (sheetName === STEN_V35.baseSheets[3] && /Работ/i.test(joined) && !/материал/i.test(joined)) return 'Группа работ';
  if (sheetName === STEN_V35.baseSheets[3] && /Материал/i.test(joined)) return 'Группа материалов — отделка';

  return currentType || '';
}

function extractCodes_(formula) {
  const result = [];
  const re = /\b([MW]-\d{3})\b/gi;
  let match;
  while ((match = re.exec(String(formula || ''))) !== null) {
    const code = String(match[1]).toUpperCase();
    if (!result.includes(code)) result.push(code);
  }
  return result;
}

function validateReferenceSheet_(ss, sheetName, prefix, issues) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  if (sheet.getLastRow() < 2) return;

  const width = sheetName === STEN_V35.materialSheet ? 11 : 10;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, width).getValues();
  const seen = new Map();

  data.forEach((row, idx) => {
    const rowNumber = idx + 2;
    const code = String(row[0] || '').trim();
    const name = String(row[1] || '').trim();
    const price = row[3];

    if (!code && !name) return;
    if (!code) {
      addIssue_(issues, 'ОШИБКА', 'Справочник', sheetName, `A${rowNumber}`, 'Есть наименование без кода.');
      return;
    }
    if (code.indexOf(prefix) !== 0) {
      addIssue_(issues, 'ОШИБКА', 'Справочник', sheetName, `A${rowNumber}`, `Неверный префикс кода ${code}. Ожидался ${prefix}.`);
    }
    if (seen.has(code)) {
      addIssue_(issues, 'ОШИБКА', 'Справочник', sheetName, `A${rowNumber}`, `Дублирующийся код ${code}; первое вхождение строка ${seen.get(code)}.`);
    } else {
      seen.set(code, rowNumber);
    }
    if (!name) addIssue_(issues, 'ОШИБКА', 'Справочник', sheetName, `B${rowNumber}`, `${code}: пустое наименование.`);

    if (price === '' || price === null || typeof price !== 'number' || !Number.isFinite(price) || price < 0) {
      addIssue_(issues, 'ОШИБКА', 'Цена', sheetName, `D${rowNumber}`, `${code}: некорректная цена «${price}».`);
    }

    if (sheetName === STEN_V35.workSheet && !String(row[7] || '').trim()) {
      addIssue_(issues, 'ПРЕДУПРЕЖДЕНИЕ', 'Группа', sheetName, `H${rowNumber}`, `${code}: не заполнена группа работы.`);
    }

    if (sheetName === STEN_V35.materialSheet) {
      if (!String(row[7] || '').trim()) {
        addIssue_(issues, 'ПРЕДУПРЕЖДЕНИЕ', 'Группа', sheetName, `H${rowNumber}`, `${code}: не заполнена техническая группа.`);
      }
      const purchase = String(row[10] || '').trim();
      if (!['Стратегические материалы', 'Обычные материалы'].includes(purchase)) {
        addIssue_(issues, 'ОШИБКА', 'Группа закупки', sheetName, `K${rowNumber}`, `${code}: допустимо только «Стратегические материалы» или «Обычные материалы».`);
      }
    }
  });
}

function validateDuplicateNames_(ss, sheetName, issues) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  const groups = new Map();

  data.forEach((row, idx) => {
    const code = String(row[0] || '').trim();
    const name = normalizeName_(row[1]);
    const unit = normalizeName_(row[2]);
    const price = row[3];
    if (!code || !name) return;
    const key = `${name}|${unit}`;
    if (!groups.has(key)) groups.set(key, []);
    groups.get(key).push({ code: code, price: price, row: idx + 2 });
  });

  groups.forEach(items => {
    if (items.length < 2) return;
    const distinct = Array.from(new Set(items.map(item => String(item.price))));
    if (distinct.length > 1) {
      const details = items.map(item => `${item.code}=${item.price}`).join('; ');
      addIssue_(issues, 'ОШИБКА', 'Дубли', sheetName, items.map(i => i.code).join(', '), `Одинаковое наименование имеет разные цены: ${details}. По принятому правилу цены должны быть унифицированы по максимуму.`);
    }
  });
}

function validateUsageColumn_(ss, sheetName, codeMap, actualUsage, issues) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet || sheet.getLastRow() < 2) return;
  const count = sheet.getLastRow() - 1;
  const data = sheet.getRange(2, 1, count, 9).getDisplayValues();

  data.forEach((row, idx) => {
    const code = String(row[0] || '').trim();
    if (!code || !codeMap.has(code)) return;
    const current = normalizeUsage_(row[8]);
    const expected = normalizeUsage_((actualUsage.get(code) || []).join('; '));
    if (current !== expected) {
      addIssue_(issues, 'ПРЕДУПРЕЖДЕНИЕ', 'Где используется', sheetName, `I${idx + 2}`, `${code}: список мест использования не соответствует фактическим формулам. Выполните «Синхронизировать базы».`);
    }
  });
}

function writeControlReport_(ss, issues) {
  const sheet = ensureOutputSheet_(ss, STEN_V35.controlSheet);
  sheet.clear();

  const errors = issues.filter(row => row[0] === 'ОШИБКА').length;
  const warnings = issues.filter(row => row[0] === 'ПРЕДУПРЕЖДЕНИЕ').length;

  sheet.getRange('A1:F1').merge();
  sheet.getRange('A1').setValue('КОНТРОЛЬ ЦЕЛОСТНОСТИ СТЕН_МАСТЕР');
  sheet.getRange('A2').setValue('Проверено');
  sheet.getRange('B2').setValue(new Date()).setNumberFormat('dd.MM.yyyy HH:mm:ss');
  sheet.getRange('D2').setValue('Ошибок');
  sheet.getRange('E2').setValue(errors);
  sheet.getRange('D3').setValue('Предупреждений');
  sheet.getRange('E3').setValue(warnings);

  const headers = ['Статус', 'Категория', 'Лист', 'Ячейка / код', 'Описание', 'Версия'];
  sheet.getRange(5, 1, 1, headers.length).setValues([headers]);
  const rows = issues.slice(0, STEN_V35.maxAuditIssues).map(row => row.concat([STEN_V35.version]));
  if (rows.length) sheet.getRange(6, 1, rows.length, headers.length).setValues(rows);

  formatControlSheet_(sheet, Math.max(rows.length + 5, 6));
}

function addIssue_(issues, status, category, sheet, cell, description) {
  if (issues.length >= STEN_V35.maxAuditIssues) return;
  issues.push([status, category, sheet, cell, description]);
}

function assignCodesOnSheet_(ss, sheetName, prefix) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  const data = sheet.getRange(2, 1, lastRow - 1, Math.max(sheetName === STEN_V35.materialSheet ? 11 : 10, 2)).getValues();
  let maxNumber = 0;
  data.forEach(row => {
    const code = String(row[0] || '').trim();
    const m = code.match(new RegExp(`^${prefix.replace('-', '\\-')}(\\d+)$`));
    if (m) maxNumber = Math.max(maxNumber, Number(m[1]));
  });

  let assigned = 0;
  data.forEach((row, idx) => {
    const code = String(row[0] || '').trim();
    const name = String(row[1] || '').trim();
    if (code || !name) return;
    maxNumber++;
    const newCode = `${prefix}${String(maxNumber).padStart(3, '0')}`;
    sheet.getRange(idx + 2, 1).setValue(newCode);
    if (sheetName === STEN_V35.materialSheet && !String(row[10] || '').trim()) {
      sheet.getRange(idx + 2, STEN_V35.materialPurchaseGroupColumn).setValue('Обычные материалы');
    }
    assigned++;
  });
  return assigned;
}

function ensureLogSheet_(ss) {
  let sheet = ss.getSheetByName(STEN_V35.logSheet);
  if (!sheet) sheet = ss.insertSheet(STEN_V35.logSheet);
  const headers = ['Дата/время', 'Пользователь', 'Лист', 'Код', 'Наименование', 'Было', 'Стало', 'Ячейка', 'Что изменено'];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (sheet.getLastRow() > 1) {
    const count = sheet.getLastRow() - 1;
    const dates = sheet.getRange(2, 1, count, 1).getValues();
    const fields = sheet.getRange(2, 9, count, 1).getValues();
    let changed = false;
    for (let i = 0; i < count; i++) {
      if (dates[i][0] && !String(fields[i][0] || '').trim()) {
        fields[i][0] = 'Цена';
        changed = true;
      }
    }
    if (changed) sheet.getRange(2, 9, count, 1).setValues(fields);
  }
  sheet.setFrozenRows(1);
  return sheet;
}

function ensureSnapshotSheet_(ss) {
  let sheet = ss.getSheetByName(STEN_V35.snapshotSheet);
  if (!sheet) sheet = ss.insertSheet(STEN_V35.snapshotSheet);
  if (sheet.getRange(1, 1).getValue() !== 'Ключ') {
    sheet.clear();
    sheet.getRange(1, 1, 1, 2).setValues([['Ключ', 'Значение']]);
  }
  return sheet;
}

function ensureOutputSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function rebuildAuditSnapshot_(ss, snapshot) {
  const rows = [['Ключ', 'Значение']];
  Object.keys(STEN_V35.tracked).forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet || sheet.getLastRow() < STEN_V35.firstDataRow) return;

    const count = sheet.getLastRow() - STEN_V35.firstDataRow + 1;
    const codes = sheet.getRange(STEN_V35.firstDataRow, STEN_V35.codeColumn, count, 1).getDisplayValues();

    Object.keys(STEN_V35.tracked[sheetName]).map(Number).forEach(col => {
      const fieldName = STEN_V35.tracked[sheetName][col];
      const values = sheet.getRange(STEN_V35.firstDataRow, col, count, 1).getValues();
      for (let i = 0; i < count; i++) {
        const code = String(codes[i][0] || '').trim();
        if (!code) continue;
        rows.push([makeSnapshotKey_(sheetName, code, fieldName), values[i][0]]);
      }
    });
  });

  snapshot.clearContents();
  snapshot.getRange(1, 1, rows.length, 2).setValues(rows);
}

function readAuditSnapshot_(snapshot) {
  const result = { values: new Map(), rows: new Map() };
  const lastRow = snapshot.getLastRow();
  if (lastRow < 2) return result;
  const data = snapshot.getRange(2, 1, lastRow - 1, 2).getValues();
  data.forEach((row, idx) => {
    const key = String(row[0] || '').trim();
    if (!key) return;
    result.values.set(key, row[1]);
    result.rows.set(key, idx + 2);
  });
  return result;
}

function applySnapshotUpdates_(snapshot, snapshotData, updates) {
  updates.forEach(update => {
    if (snapshotData.rows.has(update.key)) {
      snapshot.getRange(snapshotData.rows.get(update.key), 2).setValue(update.value);
      snapshotData.values.set(update.key, update.value);
    } else {
      const row = snapshot.getLastRow() + 1;
      snapshot.getRange(row, 1, 1, 2).setValues([[update.key, update.value]]);
      snapshotData.rows.set(update.key, row);
      snapshotData.values.set(update.key, update.value);
    }
  });
}

function makeSnapshotKey_(sheetName, code, fieldName) {
  return `${sheetName}|${code}|${fieldName}`;
}

function getActor_(e) {
  try {
    if (e && e.user && typeof e.user.getEmail === 'function') {
      const email = e.user.getEmail();
      if (email) return email;
    }
  } catch (err) {}

  try {
    const activeEmail = Session.getActiveUser().getEmail();
    if (activeEmail) return activeEmail;
  } catch (err) {}

  return 'Пользователь не раскрыт Google';
}

function valuesEqual_(a, b) {
  if ((a === '' || a === null) && (b === '' || b === null)) return true;
  if (a instanceof Date && b instanceof Date) return a.getTime() === b.getTime();

  const na = typeof a === 'number' ? a : Number(String(a).replace(',', '.'));
  const nb = typeof b === 'number' ? b : Number(String(b).replace(',', '.'));
  if (!Number.isNaN(na) && !Number.isNaN(nb)) return Math.abs(na - nb) < 1e-9;
  return String(a) === String(b);
}

function normalizeName_(value) {
  return String(value || '').toLowerCase().replace(/ё/g, 'е').replace(/\s+/g, ' ').trim();
}

function normalizeUsage_(value) {
  const items = String(value || '')
    .split(';')
    .map(s => s.trim())
    .filter(Boolean);
  return Array.from(new Set(items)).sort().join('; ');
}

function columnToLetter_(column) {
  let n = column;
  let result = '';
  while (n > 0) {
    const rem = (n - 1) % 26;
    result = String.fromCharCode(65 + rem) + result;
    n = Math.floor((n - 1) / 26);
  }
  return result;
}

function activateSheet_(name) {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(name);
  if (!sheet) {
    ss.toast(`Лист «${name}» не найден.`, 'СТЕН_МАСТЕР', 5);
    return;
  }
  ss.setActiveSheet(sheet);
}

function formatControlSheet_(sheet, lastRow) {
  sheet.setFrozenRows(5);
  sheet.setHiddenGridlines(true);
  sheet.getRange('A1:F1').setBackground('#2f2b27').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('A5:F5').setBackground('#1f4e78').setFontColor('#ffffff').setFontWeight('bold');
  if (lastRow >= 6) sheet.getRange(6, 1, lastRow - 5, 6).setWrap(true).setVerticalAlignment('top');
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 190);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 650);
  sheet.setColumnWidth(6, 80);
}

function formatStrategicSheet_(sheet, lastRow) {
  sheet.setFrozenRows(4);
  sheet.setHiddenGridlines(true);
  sheet.getRange('A1:K1').setBackground('#2f2b27').setFontColor('#ffffff').setFontWeight('bold').setHorizontalAlignment('center');
  sheet.getRange('A4:K4').setBackground('#1f4e78').setFontColor('#ffffff').setFontWeight('bold').setWrap(true);
  if (lastRow >= 5) sheet.getRange(5, 1, lastRow - 4, 11).setWrap(true).setVerticalAlignment('top');
  const widths = [90, 190, 360, 90, 120, 100, 520, 120, 150, 220, 360];
  widths.forEach((w, i) => sheet.setColumnWidth(i + 1, w));
}
/* === WEB API v3.8 FOR GITHUB PAGES === */
const WEB_V38 = Object.freeze({
  version: '3.8-web',
  modelVersion: '3.8',
  spreadsheetId: '1kpJgn6SKs9fQ2l-Bf9w_jfvHQ87UT1A_wfHIfUjKm0M',
  adminKeyProperty: 'STEN_MASTER_WEB_ADMIN_KEY',
  cacheKey: 'STEN_MASTER_WEB_BOOTSTRAP_V38',
  cacheSeconds: 60,
  systemSheet: 'Система_v3.5',
  certificateSheet: 'Сертификаты',
  certificateFolderProperty: 'STEN_MASTER_CERTIFICATE_FOLDER_ID',
  certificateFolderName: 'СТЕН_МАСТЕР_Сертификаты',
  maxCertificateBytes: 8 * 1024 * 1024
});

/**
 * One-time setup for WEB v3.8. Keeps the existing Admin key if it already exists,
 * creates the certificate registry/folder and refreshes Web API metadata.
 */
function setupWebApi() {
  const ss = SpreadsheetApp.openById(WEB_V38.spreadsheetId);
  const key = webEnsureAdminKey_();
  webEnsureCertificateSheet_(ss);
  webEnsureCertificateFolder_(ss);
  let sheet = ss.getSheetByName(WEB_V38.systemSheet);
  if (!sheet) sheet = ss.insertSheet(WEB_V38.systemSheet);
  sheet.getRange('A29:B36').setValues([
    ['WEB API v3.8', 'GitHub Pages + администрирование + PDF + история цен + сертификаты'],
    ['Admin key', key],
    ['Важно', 'Не публикуйте Admin key в GitHub. Вводите его только в панели администратора веб-калькулятора.'],
    ['Spreadsheet ID', WEB_V38.spreadsheetId],
    ['API version', WEB_V38.version],
    ['Сертификаты', WEB_V38.certificateSheet],
    ['Макс. размер сертификата', Math.round(WEB_V38.maxCertificateBytes / 1024 / 1024) + ' МБ'],
    ['Обновлено', new Date()]
  ]);
  sheet.getRange('A29:A36').setFontWeight('bold');
  webClearCache_();
  return key;
}

function setupWebV38() { return setupWebApi(); }

function doGet(e) {
  try {
    const action = String((e && e.parameter && e.parameter.action) || 'bootstrap');
    if (action === 'health') {
      return webJson_({ok:true, version:WEB_V38.version, timestamp:new Date().toISOString()});
    }
    if (action !== 'bootstrap') throw new Error('Неизвестная команда GET: ' + action);
    const force = String((e && e.parameter && e.parameter.force) || '') === '1';
    return webJson_(webBootstrap_(force));
  } catch (err) {
    return webJson_({ok:false, error:String(err && err.message || err)});
  }
}

function doPost(e) {
  try {
    const body = webParseBody_(e);
    const action = String(body.action || '');
    if (!action) throw new Error('Не указано действие.');
    const ss = SpreadsheetApp.openById(WEB_V38.spreadsheetId);

    // Public read/report actions. They never expose editor identity or private Drive file IDs.
    if (action === 'priceHistory') {
      return webJson_({ok:true, data:webReadPriceHistory_(ss, body.codes)});
    }
    if (action === 'buildPdfReport') {
      return webJson_(webBuildPdfReport_(ss, body.report || {}));
    }
    if (action === 'downloadCertificate') {
      return webJson_(webDownloadCertificate_(ss, String(body.certificateId || '')));
    }

    webAssertAdmin_(body.token);

    if (action === 'adminBootstrap') {
      return webJson_({
        ok:true,
        data:webBootstrap_(true),
        audit:webReadAudit_(ss, Number(body.limit || 100)),
        control:webReadControl_(ss),
        certificates:webReadCertificates_(ss)
      });
    }
    if (action === 'updateReference') {
      const result = webUpdateReference_(ss, body);
      return webJson_({ok:true, result, data:webBootstrap_(true), audit:webReadAudit_(ss, 50)});
    }
    if (action === 'addMaterial') {
      const material = webAddMaterial_(ss, body);
      return webJson_({ok:true, material, data:webBootstrap_(true), audit:webReadAudit_(ss, 50)});
    }
    if (action === 'uploadCertificate') {
      const certificate = webUploadCertificate_(ss, body);
      webClearCache_();
      return webJson_({ok:true, certificate, certificates:webReadCertificates_(ss), data:webBootstrap_(true)});
    }
    if (action === 'deleteCertificate') {
      webDeleteCertificate_(ss, String(body.certificateId || ''), webActor_(body.actor));
      webClearCache_();
      return webJson_({ok:true, certificates:webReadCertificates_(ss), data:webBootstrap_(true)});
    }
    if (action === 'synchronize') {
      synchronizeBasesCore_(ss, false);
      SpreadsheetApp.flush();
      webClearCache_();
      return webJson_({ok:true, data:webBootstrap_(true)});
    }
    if (action === 'checkIntegrity') {
      checkSystemIntegrityCore_(ss, false);
      return webJson_({ok:true, control:webReadControl_(ss)});
    }
    if (action === 'strategicReport') {
      buildStrategicReportCore_(ss, false);
      return webJson_({ok:true, strategic:webReadStrategic_(ss)});
    }
    if (action === 'assignCodes') {
      const works = assignCodesOnSheet_(ss, STEN_V35.workSheet, 'W-');
      const mats = assignCodesOnSheet_(ss, STEN_V35.materialSheet, 'M-');
      synchronizeBasesCore_(ss, false);
      rebuildAuditSnapshot_(ss, ensureSnapshotSheet_(ss));
      SpreadsheetApp.flush();
      webClearCache_();
      return webJson_({ok:true, assigned:{works,materials:mats}, data:webBootstrap_(true)});
    }
    if (action === 'backup') {
      const tz = ss.getSpreadsheetTimeZone() || Session.getScriptTimeZone();
      const stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd_HH-mm-ss');
      const file = DriveApp.getFileById(WEB_V38.spreadsheetId);
      const copy = file.makeCopy(file.getName() + '_BACKUP_' + stamp);
      return webJson_({ok:true, backup:{id:copy.getId(), name:copy.getName(), url:copy.getUrl()}});
    }
    throw new Error('Неизвестная команда POST: ' + action);
  } catch (err) {
    return webJson_({ok:false, error:String(err && err.message || err)});
  }
}

function webBootstrap_(force) {
  const cache = CacheService.getScriptCache();
  if (!force) {
    const cached = cache.get(WEB_V38.cacheKey);
    if (cached) return JSON.parse(cached);
  }

  const ss = SpreadsheetApp.openById(WEB_V38.spreadsheetId);
  const refs = webReadPublicReferences_(ss);
  const priceMap = {};
  refs.works.concat(refs.materials).forEach(x => priceMap[x.code] = x.price);
  const payload = {
    ok: true,
    version: WEB_V38.modelVersion,
    apiVersion: WEB_V38.version,
    timestamp: new Date().toISOString(),
    source: ss.getName(),
    defaults: webReadDefaults_(ss),
    baseline: webReadBaseline_(ss),
    works: refs.works,
    materials: refs.materials,
    constructions: {
      ext: webParseConstructionSheet_(ss.getSheetByName('База_Наружные'), 'ext', priceMap),
      inter: webParseConstructionSheet_(ss.getSheetByName('База_Межквартирные'), 'inter', priceMap),
      part: webParseConstructionSheet_(ss.getSheetByName('База_Перегородки'), 'part', priceMap)
    },
    finish: webParseFinish_(ss, priceMap),
    delivery: webParseDelivery_(ss),
    certificates: webReadCertificates_(ss),
    constants: {extOpeningPct:0.15, interOpeningPct:0.07, partOpeningPct:0.10, unitWasteFactor:1.07}
  };

  const txt = JSON.stringify(payload);
  if (txt.length < 95000) cache.put(WEB_V38.cacheKey, txt, WEB_V38.cacheSeconds);
  return payload;
}

function webReadBaseline_(ss) {
  const sh=ss.getSheetByName('ГЛАВНАЯ СТРАНИЦА');
  const get=(a1)=>webNum_(sh.getRange(a1).getValue());
  const out={ext:[],inter:[],part:[],summary:[]};
  ['D','G','J'].forEach(col=>{
    out.ext.push({
      thick:get(col+'21'),thickFin:get(col+'22'),footprint:get(col+'23'),density:get(col+'24'),
      quantity:get(col+'26'),massFloor:get(col+'27'),unitPrice:get(col+'28'),costFloor:get(col+'29'),
      costObj:get(col+'30'),trucks:get(col+'31'),delivery:get(col+'32'),finishAreaFloor:get(col+'33'),
      finishCost:get(col+'34'),total:get(col+'35'),workUnit:get(col+'37'),materialUnit:get(col+'38')
    });
    out.inter.push({
      thick:get(col+'45'),thickFin:get(col+'46'),footprint:get(col+'47'),density:get(col+'48'),
      quantity:get(col+'50'),massFloor:get(col+'51'),workUnit:get(col+'52'),materialUnit:get(col+'53'),
      unitPrice:get(col+'54'),costFloor:get(col+'55'),costObj:get(col+'56'),trucks:get(col+'57'),
      delivery:get(col+'58'),finishAreaFloor:get(col+'59'),finishRate:get(col+'60'),finishCost:get(col+'61'),total:get(col+'62')
    });
    out.part.push({
      thick:get(col+'71'),thickFin:get(col+'72'),footprint:get(col+'73'),density:get(col+'74'),areaFloor:get(col+'75'),
      quantity:get(col+'76'),massFloor:get(col+'77'),unitPrice:get(col+'78'),costFloor:get(col+'79'),costObj:get(col+'80'),
      trucks:get(col+'81'),deliveryVolFloor:get(col+'82'),trucksFloor:get(col+'83'),delivery:get(col+'84'),
      finishAreaFloor:get(col+'85'),finishRate:get(col+'86'),finishCost:get(col+'87'),total:get(col+'88'),
      workUnit:get(col+'90'),materialUnit:get(col+'91')
    });
    out.summary.push({deltaFootprint:get(col+'95'),deltaSaleFloor:get(col+'96'),deltaSaleObj:get(col+'97'),deltaSaleValue:get(col+'98')});
  });
  ['E','H','K'].forEach((col,i)=>{
    out.summary[i].changedExt=get(col+'100');
    out.summary[i].changedInter=get(col+'101');
    out.summary[i].changedPart=get(col+'102');
    out.summary[i].changedTotal=get(col+'103');
    out.summary[i].changedPerTotalArea=get(col+'104');
    out.summary[i].changedPerSaleArea=get(col+'105');
  });
  return out;
}

function webReadPublicReferences_(ss) {
  return {
    works: webReadReferenceSheet_(ss.getSheetByName(STEN_V35.workSheet), false, false),
    materials: webReadReferenceSheet_(ss.getSheetByName(STEN_V35.materialSheet), true, false)
  };
}

function webReadReferenceSheet_(sheet, isMaterial, includePrivate) {
  if (!sheet || sheet.getLastRow() < 2) return [];
  const cols = isMaterial ? 19 : 10;
  const values = sheet.getRange(2, 1, sheet.getLastRow() - 1, cols).getValues();
  return values.filter(r => String(r[0] || '').trim() && String(r[1] || '').trim()).map(r => {
    const item = {
      code: String(r[0]).trim(), name: String(r[1]).trim(), unit: String(r[2] || '').trim(),
      price: webNum_(r[3]), technicalGroup: String(r[7] || '').trim(),
      usage: String(r[8] || '').trim(), comment: String(r[9] || '').trim()
    };
    if (isMaterial) {
      item.purchaseGroup = String(r[10] || '').trim() || 'Обычные материалы';
      item.dimensions = String(r[11] || '').trim();
      item.densitySpec = r[12] === '' || r[12] == null ? null : webNum_(r[12]);
      item.strength = String(r[13] || '').trim();
      item.characteristicsSource = String(r[14] || '').trim();
      item.verificationStatus = String(r[15] || '').trim();
      item.unitComment = String(r[16] || '').trim();
      item.soundInsulation = String(r[17] || '').trim();
      item.calculatorStatus = String(r[18] || '').trim();
    }
    if (includePrivate) {
      item.previousPrice = r[4];
      item.changedAt = r[5] instanceof Date ? r[5].toISOString() : String(r[5] || '');
      item.changedBy = String(r[6] || '');
    }
    return item;
  });
}

function webReadDefaults_(ss) {
  const sh = ss.getSheetByName('ГЛАВНАЯ СТРАНИЦА');
  return {
    city:String(sh.getRange('D5').getValue() || '').trim(), saleArea:webNum_(sh.getRange('D6').getValue()),
    totalArea:webNum_(sh.getRange('D7').getValue()), perimExt:webNum_(sh.getRange('D8').getValue()),
    perimInter:webNum_(sh.getRange('D9').getValue()), perimPart:webNum_(sh.getRange('D10').getValue()),
    floorH:webNum_(sh.getRange('D11').getValue()), floors:webNum_(sh.getRange('D12').getValue()),
    salePrice:webNum_(sh.getRange('D13').getValue()),
    project:{
      ext:{material:String(sh.getRange('G8').getValue()||''),finish:String(sh.getRange('I8').getValue()||'Без отделки')},
      inter:{material:String(sh.getRange('G9').getValue()||''),finish:String(sh.getRange('I9').getValue()||'Без отделки')},
      part:{material:String(sh.getRange('G10').getValue()||''),finish:String(sh.getRange('I10').getValue()||'Без отделки')}
    },
    scenarios:[
      {ext:{material:String(sh.getRange('D17').getValue()||''),finish:String(sh.getRange('D18').getValue()||'Без отделки')},inter:{material:String(sh.getRange('D41').getValue()||''),finish:String(sh.getRange('D42').getValue()||'Без отделки')},part:{material:String(sh.getRange('D67').getValue()||''),finish:String(sh.getRange('D68').getValue()||'Без отделки')}},
      {ext:{material:String(sh.getRange('G17').getValue()||''),finish:String(sh.getRange('G18').getValue()||'Без отделки')},inter:{material:String(sh.getRange('G41').getValue()||''),finish:String(sh.getRange('G42').getValue()||'Без отделки')},part:{material:String(sh.getRange('G67').getValue()||''),finish:String(sh.getRange('G68').getValue()||'Без отделки')}},
      {ext:{material:String(sh.getRange('J17').getValue()||''),finish:String(sh.getRange('J18').getValue()||'Без отделки')},inter:{material:String(sh.getRange('J41').getValue()||''),finish:String(sh.getRange('J42').getValue()||'Без отделки')},part:{material:String(sh.getRange('J67').getValue()||''),finish:String(sh.getRange('J68').getValue()||'Без отделки')}}
    ]
  };
}

function webParseConstructionSheet_(sheet, category, priceMap) {
  if (!sheet) return [];
  const lastRow = Math.min(sheet.getLastRow(), 300);
  const values = sheet.getRange(1, 1, lastRow, 17).getValues();
  const starts = [];
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][15] || '').trim() === 'Конструкция' && String(values[i][0] || '').trim()) starts.push(i);
  }
  const out = [];
  starts.forEach((start, idx) => {
    const end = idx + 1 < starts.length ? starts[idx + 1] - 1 : Math.min(values.length - 1, start + 40);
    const top = values[start];
    const name = String(top[0]||'').trim();
    if (/полистирол/i.test(name)) return;
    if (category === 'part' && /^не проходит/i.test(String(top[9] || '').trim())) return;
    const item = {
      name:name, city:String(top[1]||'').trim(), thick:webNum_(top[2]), thickFin:webNum_(top[3]),
      density:webNum_(top[4]), L:webNum_(top[5]), W:webNum_(top[6]), H:webNum_(top[7]),
      sheetPrice:webNum_(top[8]), reference:top[9] == null ? '' : top[9],
      workCostSheet:webNum_(top[12]), materialCostSheet:webNum_(top[13]), sourceRow:start+1, components:[]
    };
    for (let r = start + 1; r <= end; r++) {
      const row = values[r];
      const type = String(row[15] || '').trim();
      const code = String(row[14] || '').trim();
      if (!code || (type !== 'Работа' && type !== 'Материал')) continue;
      item.components.push({
        code, type, name:String(row[2] || '').trim(), price:webNum_(priceMap[code]),
        sheetContribution:webNum_(row[7]), purchaseGroup:String(row[16] || '').trim(),
        note:String(row[9] || '').trim(), sourceRow:r+1,
        normCells:{D:row[3], E:row[4], F:row[5]}
      });
    }
    out.push(item);
  });
  return out;
}

function webParseFinish_(ss, priceMap) {
  const sh = ss.getSheetByName('База_Отделка');
  const values = sh.getRange(1,1,Math.min(sh.getLastRow(),100),17).getValues();
  const defs = [
    ['Предчистовая','Кирпич/блок', [8,9,10,11,12,15,16,17,18,19,20,21,22], 'W-010', 6],
    ['Предчистовая','ГВЛВ', [26,27,28,29,30,31,34,35,36,37,38,39,40,41], 'W-009', 24],
    ['Чистовая','Кирпич/блок', webRange_(46,61), 'W-010', 44],
    ['Чистовая','ГВЛВ', webRange_(65,81), 'W-009', 63]
  ];
  const out = [
    {type:'Без отделки',subtype:'Кирпич/блок',sheetRate:0,components:[]},
    {type:'Без отделки',subtype:'ГВЛВ',sheetRate:0,components:[]}
  ];
  defs.forEach(def => {
    const type=def[0], subtype=def[1], rows=def[2], extra=def[3], rateRow=def[4];
    const components=[];
    rows.forEach(oneBased => {
      const row=values[oneBased-1];
      const code=String(row[14]||'').trim();
      const ctype=String(row[15]||'').trim();
      if(!code || (ctype!=='Работа'&&ctype!=='Материал')) return;
      components.push({code,type:ctype,name:String(row[0]||'').trim(),price:webNum_(priceMap[code]),sheetContribution:webNum_(row[3]),sourceRow:oneBased,purchaseGroup:String(row[16]||'').trim()});
    });
    components.push({code:extra,type:'Работа',name:'Электрика',price:webNum_(priceMap[extra]),sheetContribution:webNum_(priceMap[extra]),sourceRow:null,purchaseGroup:''});
    out.push({type,subtype,sheetRate:webNum_(values[rateRow-1][1]),components});
  });
  return out;
}

function webParseDelivery_(ss) {
  const sh=ss.getSheetByName('База_Доставка');
  const values=sh.getRange(1,1,15,7).getValues();
  const productionCities=values[0].slice(1).map(x=>String(x||'').trim());
  const distances={};
  for(let r=1;r<=9;r++) distances[String(values[r][0]||'').trim()]=values[r].slice(1,7).map(webNum_);
  return {productionCities,distances,rateKm:webNum_(values[12][2]),truckVol:webNum_(values[13][2]),truckKg:webNum_(values[14][2])};
}

function webUpdateReference_(ss, body) {
  const sheetName=String(body.sheet||'');
  if (![STEN_V35.workSheet,STEN_V35.materialSheet].includes(sheetName)) throw new Error('Недопустимый справочник.');
  const code=String(body.code||'').trim();
  if(!code) throw new Error('Не указан код позиции.');
  const field=String(body.field||'');
  const map = sheetName===STEN_V35.workSheet
    ? {price:{col:4,label:'Цена'},technicalGroup:{col:8,label:'Группа'}}
    : {price:{col:4,label:'Цена с НДС и доставкой'},technicalGroup:{col:8,label:'Техническая группа'},purchaseGroup:{col:11,label:'Группа закупки'}};
  if(!map[field]) throw new Error('Недопустимое поле: '+field);
  const sheet=ss.getSheetByName(sheetName);
  const data=sheet.getRange(2,1,Math.max(sheet.getLastRow()-1,1),Math.max(sheet.getLastColumn(),19)).getValues();
  const idx=data.findIndex(r=>String(r[0]||'').trim()===code);
  if(idx<0) throw new Error('Код не найден: '+code);
  const row=idx+2, meta=map[field], cell=sheet.getRange(row,meta.col), oldValue=cell.getValue();
  let newValue=body.value;
  if(field==='price') {
    newValue=webNum_(newValue);
    if(newValue<0) throw new Error('Цена не может быть отрицательной.');
  } else {
    newValue=String(newValue||'').trim();
    if(!newValue) throw new Error('Значение не может быть пустым.');
  }
  const actor=webActor_(body.actor);
  const lock=LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    cell.setValue(newValue);
    webWriteAudit_(ss,sheet,row,code,meta.label,oldValue,newValue,actor,cell.getA1Notation());
    if(field==='price') webApplyDuplicateMaximum_(ss,sheetName,row,actor);
    synchronizeBasesCore_(ss,false);
    SpreadsheetApp.flush();
    rebuildAuditSnapshot_(ss, ensureSnapshotSheet_(ss));
    webClearCache_();
  } finally { lock.releaseLock(); }
  return {sheet:sheetName,code,field,oldValue,newValue};
}

function webAddMaterial_(ss, body) {
  const m = body.material || {};
  const name = webRequiredText_(m.name, 'Наименование', 500);
  const unit = webRequiredText_(m.unit, 'Единица измерения', 40);
  const price = webNum_(m.price);
  if (price < 0) throw new Error('Цена не может быть отрицательной.');
  const technicalGroup = webRequiredText_(m.technicalGroup || 'Прочие материалы', 'Техническая группа', 120);
  const purchaseGroup = webRequiredText_(m.purchaseGroup || 'Обычные материалы', 'Группа закупки', 80);
  const actor = webActor_(body.actor);
  const sheet = ss.getSheetByName(STEN_V35.materialSheet);
  if (!sheet) throw new Error('Не найден лист «Материалы».');
  const lock=LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const last = Math.max(sheet.getLastRow(), 1);
    const existing = last >= 2 ? sheet.getRange(2,1,last-1,19).getValues() : [];
    const dup = existing.find(r => webNormalize_(r[1]) === webNormalize_(name) && webNormalize_(r[2]) === webNormalize_(unit));
    if (dup) throw new Error('Материал с таким наименованием и единицей уже существует: ' + String(dup[0] || ''));
    const code = webNextCode_(existing.map(r=>String(r[0]||'')), 'M-');
    const row = last + 1;
    if (row > sheet.getMaxRows()) sheet.insertRowsAfter(sheet.getMaxRows(), Math.max(10, row-sheet.getMaxRows()));
    if (last >= 2) {
      sheet.getRange(last,1,1,19).copyTo(sheet.getRange(row,1,1,19), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      sheet.getRange(last,1,1,19).copyTo(sheet.getRange(row,1,1,19), SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION, false);
    }
    const now = new Date();
    const values = [[
      code, name, unit, price, '', now, actor, technicalGroup, '', String(m.comment||'').trim(), purchaseGroup,
      String(m.dimensions||'').trim(), m.densitySpec===''||m.densitySpec==null?'':webNum_(m.densitySpec), String(m.strength||'').trim(),
      String(m.characteristicsSource||'').trim(), String(m.verificationStatus||'').trim(), String(m.unitComment||'').trim(),
      String(m.soundInsulation||'').trim(), String(m.calculatorStatus||'Новый материал').trim()
    ]];
    sheet.getRange(row,1,1,19).setValues(values);
    const log=ensureLogSheet_(ss);
    log.appendRow([now,actor,STEN_V35.materialSheet,code,name,'',price,'A'+row,'Добавлен материал']);
    synchronizeBasesCore_(ss,false);
    rebuildAuditSnapshot_(ss, ensureSnapshotSheet_(ss));
    SpreadsheetApp.flush();
    webClearCache_();
    return {code,name,unit,price,technicalGroup,purchaseGroup,row};
  } finally { lock.releaseLock(); }
}

function webApplyDuplicateMaximum_(ss, sheetName, editedRow, actor) {
  const sheet=ss.getSheetByName(sheetName);
  const last=sheet.getLastRow();
  if(last<2) return;
  const values=sheet.getRange(2,1,last-1,Math.min(Math.max(sheet.getLastColumn(),11),19)).getValues();
  const target=values[editedRow-2];
  const targetName=webNormalize_(target[1]), targetUnit=webNormalize_(target[2]);
  const matches=[];
  values.forEach((r,i)=>{
    if(webNormalize_(r[1])===targetName && webNormalize_(r[2])===targetUnit) matches.push({i,row:i+2,code:String(r[0]||'').trim(),price:webNum_(r[3])});
  });
  if(matches.length<2) return;
  const maxPrice=Math.max.apply(null,matches.map(x=>x.price));
  matches.forEach(x=>{
    if(Math.abs(x.price-maxPrice)<1e-9) return;
    const cell=sheet.getRange(x.row,4);
    const old=x.price;
    cell.setValue(maxPrice);
    webWriteAudit_(ss,sheet,x.row,x.code,'Цена (правило максимума)',old,maxPrice,actor,cell.getA1Notation());
  });
}

function webWriteAudit_(ss, sheet, row, code, fieldLabel, oldValue, newValue, actor, a1) {
  const now=new Date();
  const name=String(sheet.getRange(row,2).getValue()||'').trim();
  if(fieldLabel.indexOf('Цена')===0) {
    sheet.getRange(row,5).setValue(oldValue);
    sheet.getRange(row,6).setValue(now);
    sheet.getRange(row,7).setValue(actor);
  }
  const log=ensureLogSheet_(ss);
  log.appendRow([now,actor,sheet.getName(),code,name,oldValue,newValue,a1,fieldLabel]);
}

function webReadAudit_(ss, limit) {
  const sh=ensureLogSheet_(ss);
  const last=sh.getLastRow();
  if(last<2) return [];
  const count=Math.min(Math.max(limit||50,1),200,last-1);
  const values=sh.getRange(last-count+1,1,count,9).getDisplayValues();
  return values.reverse().map(r=>({date:r[0],actor:r[1],sheet:r[2],code:r[3],name:r[4],before:r[5],after:r[6],cell:r[7],field:r[8]}));
}

function webReadPriceHistory_(ss, requestedCodes) {
  const refs=webReadReferenceSheet_(ss.getSheetByName(STEN_V35.materialSheet),true,false);
  const byCode={}; refs.forEach(x=>byCode[x.code]=x);
  let codes=Array.isArray(requestedCodes)?requestedCodes.map(String):[];
  codes=codes.filter(c=>byCode[c]).slice(0,30);
  if(!codes.length) codes=refs.filter(x=>x.technicalGroup==='Стеновые материалы' && !/архив|не применяется/i.test(x.calculatorStatus||'')).slice(0,12).map(x=>x.code);
  const allowed=new Set(codes);
  const sh=ensureLogSheet_(ss), events=[];
  const last=sh.getLastRow();
  if(last>=2){
    const rows=sh.getRange(2,1,last-1,9).getValues();
    rows.forEach(r=>{
      const code=String(r[3]||'').trim(), sheet=String(r[2]||'').trim(), field=String(r[8]||'').trim();
      if(sheet!==STEN_V35.materialSheet || !allowed.has(code) || field.indexOf('Цена')!==0) return;
      const d=r[0] instanceof Date?r[0].toISOString():String(r[0]||'');
      events.push({date:d,code,before:webNum_(r[5]),after:webNum_(r[6]),field});
    });
  }
  events.sort((a,b)=>String(a.date).localeCompare(String(b.date)));
  return {
    generatedAt:new Date().toISOString(),
    materials:codes.map(c=>({code:c,name:byCode[c].name,unit:byCode[c].unit,currentPrice:byCode[c].price,technicalGroup:byCode[c].technicalGroup,purchaseGroup:byCode[c].purchaseGroup})),
    events
  };
}

function webEnsureCertificateSheet_(ss) {
  let sh=ss.getSheetByName(WEB_V38.certificateSheet);
  if(!sh) sh=ss.insertSheet(WEB_V38.certificateSheet);
  const headers=['ID','Код материала','Материал','Документ','Drive file ID','Имя файла','MIME','Размер, байт','Дата загрузки','Загрузил','Статус'];
  sh.getRange(1,1,1,headers.length).setValues([headers]).setFontWeight('bold').setBackground('#1f4e78').setFontColor('#ffffff');
  sh.setFrozenRows(1); sh.setHiddenGridlines(true);
  const widths=[230,100,420,260,220,300,160,120,160,220,100]; widths.forEach((w,i)=>sh.setColumnWidth(i+1,w));
  return sh;
}

function webEnsureCertificateFolder_(ss) {
  const props=PropertiesService.getScriptProperties();
  const saved=props.getProperty(WEB_V38.certificateFolderProperty);
  if(saved){ try { return DriveApp.getFolderById(saved); } catch(err) {} }
  const spreadsheetFile=DriveApp.getFileById(WEB_V38.spreadsheetId);
  const parents=spreadsheetFile.getParents();
  const parent=parents.hasNext()?parents.next():DriveApp.getRootFolder();
  const existing=parent.getFoldersByName(WEB_V38.certificateFolderName);
  const folder=existing.hasNext()?existing.next():parent.createFolder(WEB_V38.certificateFolderName);
  props.setProperty(WEB_V38.certificateFolderProperty,folder.getId());
  return folder;
}

function webReadCertificates_(ss) {
  const sh=webEnsureCertificateSheet_(ss);
  const last=sh.getLastRow();
  if(last<2) return [];
  return sh.getRange(2,1,last-1,11).getValues().filter(r=>String(r[0]||'').trim() && String(r[10]||'Активен')!=='Удалён').map(r=>({
    id:String(r[0]||''),code:String(r[1]||''),material:String(r[2]||''),name:String(r[3]||''),fileName:String(r[5]||''),
    mime:String(r[6]||''),size:webNum_(r[7]),uploadedAt:r[8] instanceof Date?r[8].toISOString():String(r[8]||''),status:String(r[10]||'Активен')
  }));
}

function webUploadCertificate_(ss, body) {
  const code=webRequiredText_(body.code,'Код материала',30);
  const ref=webReadReferenceSheet_(ss.getSheetByName(STEN_V35.materialSheet),true,false).find(x=>x.code===code);
  if(!ref) throw new Error('Материал не найден: '+code);
  if(ref.technicalGroup!=='Стеновые материалы') throw new Error('Сертификаты в этом разделе привязываются только к стеновым материалам.');
  const fileName=webRequiredText_(body.fileName,'Имя файла',240);
  const mime=String(body.mimeType||'application/pdf').trim().toLowerCase();
  if(!['application/pdf','image/jpeg','image/png'].includes(mime)) throw new Error('Разрешены PDF, JPG и PNG.');
  const b64=String(body.base64||'').replace(/^data:[^,]+,/, '');
  if(!b64) throw new Error('Файл не передан.');
  const bytes=Utilities.base64Decode(b64);
  if(bytes.length>WEB_V38.maxCertificateBytes) throw new Error('Файл больше '+Math.round(WEB_V38.maxCertificateBytes/1024/1024)+' МБ.');
  const docName=webRequiredText_(body.documentName||fileName,'Название документа',240);
  const actor=webActor_(body.actor);
  const folder=webEnsureCertificateFolder_(ss);
  const blob=Utilities.newBlob(bytes,mime,fileName);
  const file=folder.createFile(blob);
  const id=Utilities.getUuid();
  const sh=webEnsureCertificateSheet_(ss);
  sh.appendRow([id,code,ref.name,docName,file.getId(),fileName,mime,bytes.length,new Date(),actor,'Активен']);
  return {id,code,material:ref.name,name:docName,fileName,mime,size:bytes.length,uploadedAt:new Date().toISOString(),status:'Активен'};
}

function webDownloadCertificate_(ss, certificateId) {
  if(!certificateId) throw new Error('Не указан сертификат.');
  const sh=webEnsureCertificateSheet_(ss), last=sh.getLastRow();
  if(last<2) throw new Error('Сертификат не найден.');
  const rows=sh.getRange(2,1,last-1,11).getValues();
  const row=rows.find(r=>String(r[0]||'')===certificateId && String(r[10]||'Активен')!=='Удалён');
  if(!row) throw new Error('Сертификат не найден или удалён.');
  const file=DriveApp.getFileById(String(row[4]||''));
  const blob=file.getBlob();
  const bytes=blob.getBytes();
  if(bytes.length>WEB_V38.maxCertificateBytes) throw new Error('Сертификат слишком большой для веб-просмотра.');
  return {ok:true,certificate:{id:certificateId,name:String(row[3]||''),fileName:String(row[5]||file.getName()),mime:String(row[6]||blob.getContentType()),base64:Utilities.base64Encode(bytes)}};
}

function webDeleteCertificate_(ss, certificateId, actor) {
  if(!certificateId) throw new Error('Не указан сертификат.');
  const sh=webEnsureCertificateSheet_(ss), last=sh.getLastRow();
  const rows=last>=2?sh.getRange(2,1,last-1,11).getValues():[];
  const idx=rows.findIndex(r=>String(r[0]||'')===certificateId && String(r[10]||'Активен')!=='Удалён');
  if(idx<0) throw new Error('Сертификат не найден.');
  const rowNum=idx+2, fileId=String(rows[idx][4]||'');
  try { if(fileId) DriveApp.getFileById(fileId).setTrashed(true); } catch(err) {}
  sh.getRange(rowNum,11).setValue('Удалён');
  sh.getRange(rowNum,10).setValue(actor);
}

function webBuildPdfReport_(ss, report) {
  if (JSON.stringify(report || {}).length > 60000) throw new Error('Слишком большой пакет данных для PDF отчёта.');
  const title='СТЕН_МАСТЕР — отчёт по расчёту';
  const now=new Date();
  const tz=ss.getSpreadsheetTimeZone()||Session.getScriptTimeZone();
  const stamp=Utilities.formatDate(now,tz,'yyyy-MM-dd_HH-mm');
  const doc=DocumentApp.create('STEN_MASTER_REPORT_'+stamp+'_'+Utilities.getUuid().slice(0,8));
  try {
    const body=doc.getBody();
    body.appendParagraph(title).setHeading(DocumentApp.ParagraphHeading.TITLE);
    body.appendParagraph('Сформировано: '+Utilities.formatDate(now,tz,'dd.MM.yyyy HH:mm'));
    body.appendParagraph('Источник модели: '+String(report.source||ss.getName())+' · WEB v3.8');
    const note=body.appendParagraph('Доставка приведена справочно и не входит в итоговые затраты. Цены материалов принимаются с НДС и доставкой.');
    note.editAsText().setItalic(true);

    body.appendParagraph('Исходные данные').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    const inp=report.inputs||{};
    body.appendTable([
      ['Параметр','Значение'],
      ['Город объекта',webPdfText_(inp.city)],
      ['Продаваемая площадь, м²',webPdfNum_(inp.saleArea)],
      ['Общая площадь, м²',webPdfNum_(inp.totalArea)],
      ['Периметр наружных стен, м',webPdfNum_(inp.perimExt)],
      ['Периметр межквартирных стен, м',webPdfNum_(inp.perimInter)],
      ['Периметр перегородок, м',webPdfNum_(inp.perimPart)],
      ['Высота этажа, м',webPdfNum_(inp.floorH)],
      ['Этажей',webPdfNum_(inp.floors)],
      ['Цена 1 м² продаваемой площади, ₽',webPdfNum_(inp.salePrice)]
    ]);

    body.appendParagraph('Проектное решение').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendTable(webPdfSelectionTable_(report.project||{}));

    body.appendParagraph('Сценарии').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    const scenarios=Array.isArray(report.scenarios)?report.scenarios.slice(0,3):[];
    const results=Array.isArray(report.results)?report.results.slice(0,3):[];
    scenarios.forEach((sc,i)=>{
      body.appendParagraph('Сценарий '+(i+1)).setHeading(DocumentApp.ParagraphHeading.HEADING2);
      const r=results[i]||{};
      const rows=[['Категория','Материал','Отделка','Стоимость, ₽']];
      ['ext','inter','part'].forEach(cat=>{
        const sel=(sc&&sc[cat])||{};
        const cost=r.categories&&r.categories[cat]?r.categories[cat].total:'';
        rows.push([webPdfCat_(cat),webPdfText_(sel.material),webPdfText_(sel.finish),webPdfNum_(cost)]);
      });
      rows.push(['ИТОГО','','',webPdfNum_(r.changedTotal!=null?r.changedTotal:r.fullTotal)]);
      body.appendTable(rows);
      body.appendParagraph('Δ продаваемой площади объекта: '+webPdfNum_(r.deltaSaleObj)+' м²; денежный эффект площади: '+webPdfNum_(r.deltaSaleValue)+' ₽.');
    });

    body.appendParagraph('Примечание').setHeading(DocumentApp.ParagraphHeading.HEADING1);
    body.appendParagraph('Отчёт сформирован веб-калькулятором на основании текущих параметров пользователя. Для договорных и закупочных решений необходимо сверять исходные цены и технические документы с актуальной Google-таблицей.');
    doc.saveAndClose();
    const file=DriveApp.getFileById(doc.getId());
    const pdf=file.getBlob().getAs('application/pdf');
    const bytes=pdf.getBytes();
    const name='STEN_MASTER_Расчёт_'+stamp+'.pdf';
    file.setTrashed(true);
    return {ok:true,fileName:name,mime:'application/pdf',base64:Utilities.base64Encode(bytes)};
  } catch(err) {
    try { DriveApp.getFileById(doc.getId()).setTrashed(true); } catch(ignore) {}
    throw err;
  }
}

function webPdfSelectionTable_(project) {
  const rows=[['Категория','Материал','Отделка']];
  ['ext','inter','part'].forEach(cat=>{
    const p=project[cat]||{};
    rows.push([webPdfCat_(cat),webPdfText_(p.material),webPdfText_(p.finish)]);
  });
  return rows;
}
function webPdfCat_(cat){ return cat==='ext'?'Наружные стены':cat==='inter'?'Межквартирные стены':'Перегородки'; }
function webPdfText_(v){ return String(v==null?'':v).replace(/[\u0000-\u001f]/g,' ').slice(0,800); }
function webPdfNum_(v){ const x=webNum_(v); return Utilities.formatString('%.2f',x).replace('.',','); }

function webReadControl_(ss) {
  const sh=ss.getSheetByName(STEN_V35.controlSheet);
  if(!sh || sh.getLastRow()<1) return [];
  return sh.getRange(1,1,Math.min(sh.getLastRow(),500),Math.min(sh.getLastColumn(),6)).getDisplayValues();
}

function webReadStrategic_(ss) {
  const sh=ss.getSheetByName(STEN_V35.strategicSheet);
  if(!sh || sh.getLastRow()<1) return [];
  return sh.getRange(1,1,Math.min(sh.getLastRow(),500),Math.min(sh.getLastColumn(),11)).getDisplayValues();
}

function webEnsureAdminKey_() {
  const props=PropertiesService.getScriptProperties();
  let key=props.getProperty(WEB_V38.adminKeyProperty);
  if(!key) {
    key=Utilities.getUuid().replace(/-/g,'') + Utilities.getUuid().replace(/-/g,'').slice(0,16);
    props.setProperty(WEB_V38.adminKeyProperty,key);
  }
  return key;
}

function webAssertAdmin_(token) {
  const expected=PropertiesService.getScriptProperties().getProperty(WEB_V38.adminKeyProperty);
  if(!expected) throw new Error('Web API не настроен. Один раз запустите setupWebV38().');
  if(!token || String(token)!==String(expected)) throw new Error('Неверный Admin key.');
}

function webActor_(provided) {
  try {
    const email=Session.getActiveUser().getEmail();
    if(email) return email;
  } catch(err) {}
  return String(provided||'web-admin').trim() || 'web-admin';
}

function webParseBody_(e) {
  const raw=e && e.postData && e.postData.contents ? e.postData.contents : '';
  if(raw) {
    try { return JSON.parse(raw); } catch(err) {}
  }
  const p=(e&&e.parameter)||{};
  return p;
}

function webJson_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function webNum_(v) {
  if(typeof v==='number') return isFinite(v)?v:0;
  const n=parseFloat(String(v==null?'':v).replace(/\s/g,'').replace(',','.').replace(/[^0-9.\-]/g,''));
  return isNaN(n)?0:n;
}

function webRange_(a,b) { const out=[]; for(let i=a;i<=b;i++) out.push(i); return out; }
function webNormalize_(v) { return String(v==null?'':v).trim().toLowerCase().replace(/\s+/g,' '); }
function webClearCache_() { CacheService.getScriptCache().remove(WEB_V38.cacheKey); }
function webRequiredText_(v,label,max){ const x=String(v==null?'':v).trim(); if(!x) throw new Error('Заполните поле «'+label+'».'); return x.slice(0,max||500); }
function webNextCode_(codes,prefix){ let max=0; (codes||[]).forEach(c=>{ const m=String(c||'').match(new RegExp('^'+prefix.replace('-','\\-')+'(\\d+)$')); if(m) max=Math.max(max,Number(m[1])||0); }); return prefix+String(max+1).padStart(3,'0'); }
