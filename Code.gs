/************** CONFIG **************/
const TEAMS = ['Team 1','Team 2','Team 3','Team 4','Team 5'];
const DATA_START_ROW = 3; // данные начинаются с этой строки
// какие оттенки считаем «красный ready»
const REDS = new Set(['#ff0000','#ff5b5b','#ff6666','#f44336','#ea4335','#d32f2f','#e06666','#ea9999']);
// старт первого недельного блока и шаг до следующего (AK..AU = 11 кол, потом AV (дырка), старт AW => шаг 12)
const FIRST_BLOCK_START = 'AK';
const BLOCK_STEP = 12;   // смещение по колонкам к следующей неделе
const DAYS_IN_WEEK = 7;  // AK..AQ (7 дней) в каждом блоке

/************** MENU **************/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Dispatch')
    .addItem('READY (FAST, dynamic week block)', 'buildReadyDynamicFast')
    .addToUi();
}

/************** MAIN **************/
function buildReadyDynamicFast() {
  const ss = SpreadsheetApp.getActive();
  const tz = ss.getSpreadsheetTimeZone();
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const outName = `READY ${stamp}`;

  // создать/очистить выходной лист
  let out = ss.getSheetByName(outName);
  if (!out) out = ss.insertSheet(outName);
  out.clear({contentsOnly:true});

  // ---------- ШАПКА: берём с первого листа, где найдём активный блок ----------
  let headerSet = false;
  for (const t of TEAMS) {
    const sh = ss.getSheetByName(t);
    if (!sh) continue;
    const block = findActiveWeekBlock_(sh); // {startCol, weekCol, notesCol, hdrRow}
    if (!block) continue;
    try {
      // Заголовки “Dispatcher / Drivers” (E1:F1) -> A1:B1
      sh.getRange(1, colLetterToIndex_('E'), 1, 2)
        .copyTo(out.getRange(1, 1, 1, 2), {contentsOnly:false});
      // Даты над днями: из активного блока (hdrRow, 7 дней) -> C1:I1
      sh.getRange(block.hdrRow, block.startCol, 1, DAYS_IN_WEEK)
        .copyTo(out.getRange(1, 3, 1, DAYS_IN_WEEK), {contentsOnly:false});
      // WEEK заголовок: из строки 1 колонки weekCol -> J1
      sh.getRange(1, block.weekCol, 1, 1)
        .copyTo(out.getRange(1, 10, 1, 1), {contentsOnly:false});
      // NOTES заголовок: из строки 1 колонки notesCol -> K1
      sh.getRange(1, block.notesCol, 1, 1)
        .copyTo(out.getRange(1, 11, 1, 1), {contentsOnly:false});
      headerSet = true;
      break;
    } catch (e) {
      // если вдруг не получилось — пробуем следующую команду
    }
  }
  if (!headerSet) {
    out.getRange('A1').setValue('No active week block found.');
    return;
  }

  // ---------- 1) Собираем кандидатов по всем командам (быстро, bulk) ----------
  const candidates = []; // {sheet,row,count,maxRun,score,block}
  let orderSeq = 0;

  for (const name of TEAMS) {
    const sh = ss.getSheetByName(name);
    if (!sh) continue;

    const block = findActiveWeekBlock_(sh); // динамически на каждом листе
    if (!block) continue;

    const lastRow = sh.getLastRow();
    if (lastRow < DATA_START_ROW) continue;

    // bulk чтение значений и цветов только по активному блоку (7 дней)
    const numRows = lastRow - DATA_START_ROW + 1;
    const daysVals  = sh.getRange(DATA_START_ROW, block.startCol, numRows, DAYS_IN_WEEK).getValues();
    const daysBgs   = sh.getRange(DATA_START_ROW, block.startCol, numRows, DAYS_IN_WEEK).getBackgrounds();

    for (let i = 0; i < numRows; i++) {
      const vals = daysVals[i];
      const bgs  = daysBgs[i];

      let hasAny = false, count = 0, maxRun = 0, run = 0;
      for (let d = 0; d < DAYS_IN_WEEK; d++) {
        const v  = String(vals[d] ?? '').trim().toLowerCase();
        const bg = String(bgs[d]  ?? '').toLowerCase();
        const isRedReady = (v === 'ready') && (REDS.has(bg) || isRedLoose_(bg));
        if (isRedReady) {
          hasAny = true; count++; run++; if (run > maxRun) maxRun = run;
        } else {
          run = 0;
        }
      }

      if (!hasAny) continue;

      const row = DATA_START_ROW + i;
      const score = (count >= 3 ? 2 : 0) + (maxRun >= 2 ? 1 : 0);
      candidates.push({sheet: sh, row, count, maxRun, score, order: orderSeq++, block});
    }
  }

  if (candidates.length === 0) {
    out.getRange('A2').setValue('No RED READY found in active week block.');
    autosize_(out, 11);
    return;
  }

  // ---------- 2) Сортировка по приоритетам ----------
  candidates.sort((a,b) => b.score - a.score || b.count - a.count || a.order - b.order);

  // ---------- 3) Чистый Ctrl+C/V по отфильтрованным строкам ----------
  let rOut = 2;
  for (const c of candidates) {
    const sh = c.sheet;
    const b  = c.block;
    // E:F -> A:B
    safeCopy_(sh, c.row, colLetterToIndex_('E'), 1, 2, out, rOut, 1);
    // 7 дней (активный блок) -> C:I
    safeCopy_(sh, c.row, b.startCol, 1, DAYS_IN_WEEK, out, rOut, 3);
    // WEEK (ячейка текущей строки в weekCol) -> J
    safeCopy_(sh, c.row, b.weekCol, 1, 1, out, rOut, 10);
    // NOTES (ячейка текущей строки в notesCol) -> K
    safeCopy_(sh, c.row, b.notesCol, 1, 1, out, rOut, 11);
    rOut++;
  }

  // заморозим WEEK/NOTES как значения, чтобы отчёт не "переехал" на след. неделе
  if (rOut > 2) {
    const rngWeek  = out.getRange(2, 10, rOut - 2, 1); // J
    const rngNotes = out.getRange(2, 11, rOut - 2, 1); // K
    rngWeek.setValues(rngWeek.getDisplayValues());
    rngNotes.setValues(rngNotes.getDisplayValues());
  }

  autosize_(out, 11);
}

/************** HELPERS **************/
// Находит активный недельный блок: правыйmost, где в шапке над 7 днями вижу даты.
// Возвращает {startCol, weekCol, notesCol, hdrRow} или null.
function findActiveWeekBlock_(sh) {
  const maxCol = sh.getMaxColumns();
  let start = colLetterToIndex_(FIRST_BLOCK_START);
  let found = null;

  while (start + 10 <= maxCol) {
    const hdrRow = detectDateHeaderRowInBlock_(sh, start);
    if (hdrRow) {
      // правый самый блок перезапишет found — именно он нам и нужен
      found = {
        startCol: start,          // S
        weekCol:  start + 8,      // S+8 (после 7 дней + 1 сеп)
        notesCol: start + 10,     // S+10 (после WEEK + 1 сеп)
        hdrRow
      };
    }
    start += BLOCK_STEP; // следующий блок (AW, BI, ...)
  }
  return found;
}

// Детект строки заголовков дат для конкретного блока (startCol..startCol+6)
function detectDateHeaderRowInBlock_(sh, startCol) {
  const candidates = [1,2,3]; // обычно даты в одной из первых трёх строк
  for (const r of candidates) {
    const vals = sh.getRange(r, startCol, 1, DAYS_IN_WEEK).getDisplayValues()[0];
    const ok = vals.every(v => {
      const s = String(v || '').trim().toLowerCase();
      if (!s) return false;
      if (s === 'ready' || s === 'enroute' || s === 'break') return false;
      return /[-/]/.test(s) || /(jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)/i.test(s);
    });
    if (ok) return r;
  }
  return null;
}

// Безопасный copyTo (как Ctrl+C/V) по индексам (row/col)
function safeCopy_(srcSheet, srcRow, srcCol, numRows, numCols, dstSheet, dstRow, dstCol) {
  try {
    srcSheet.getRange(srcRow, srcCol, numRows, numCols)
            .copyTo(dstSheet.getRange(dstRow, dstCol, numRows, numCols), {contentsOnly:false});
  } catch (_) {}
}

// Мягкая проверка «красного», если точного кода нет в REDS
function isRedLoose_(hex) {
  if (!/^#[0-9a-fA-F]{6}$/.test(hex)) return false;
  const r = parseInt(hex.slice(1,3),16),
        g = parseInt(hex.slice(3,5),16),
        b = parseInt(hex.slice(5,7),16);
  return r >= 200 && g <= 90 && b <= 90;
}

// Колонки: A1->1, B1->2 ...
function colLetterToIndex_(letter) {
  let n = 0;
  for (let i=0;i<letter.length;i++) n = n*26 + (letter.charCodeAt(i)-64);
  return n;
}

// Авто-ширина первых N колонок
function autosize_(sheet, nCols) {
  for (let c = 1; c <= nCols; c++) sheet.autoResizeColumn(c);
}
