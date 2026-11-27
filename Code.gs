const SHEET_ID      = '1D1Ohn82qQ3ZYKDkG_RdMVPRc8loUmkRnpZzICGf8sGU';
const REMINDER_DAYS = 2;
const ATTACH_FOLDER = 'DashboardTracker_Attachments';
const TIMEZONE      = 'Asia/Kolkata';

/* =========================================================
   HELPERS
   ========================================================= */

function ss() { return SpreadsheetApp.openById(SHEET_ID); }
function tz() { return TIMEZONE; }

/* ---------- Date-safe formatters (Option B) ----------
   Read the local year/month/day from the Date object WITHOUT
   TZ conversion, so the formatted value always matches the
   calendar day the sheet displays. This is immune to the
   script runtime TZ and to the 05:30:00 IST offset baked into
   stored Date values.
------------------------------------------------------- */

function pad2(n){ return String(n).padStart(2,'0'); }

function ymd(d){
  if (!d) return '';
  if (typeof d === 'string') {
    // Already a yyyy-MM-dd[ ...] string? return the date part.
    const m = d.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return `${m[1]}-${m[2]}-${m[3]}`;
    // Fall back to a Date parse
    const dt = new Date(d);
    if (isNaN(dt.getTime())) return '';
    d = dt;
  }
  if (!(d instanceof Date)) return '';
  return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())}`;
}

function ym(d){
  if (!d) return '';
  if (typeof d === 'string') {
    const m = d.match(/^(\d{4})-(\d{2})/);
    if (m) return `${m[1]}-${m[2]}`;
    const dt = new Date(d);
    if (isNaN(dt.getTime())) return '';
    d = dt;
  }
  if (!(d instanceof Date)) return '';
  return `${d.getFullYear()}-${pad2(d.getMonth()+1)}`;
}

function ymdhm(d){
  if (!d) return '';
  if (typeof d === 'string') {
    const m = d.match(/^(\d{4})-(\d{2})-(\d{2})[ T](\d{2}):(\d{2})/);
    if (m) return `${m[1]}-${m[2]}-${m[3]} ${m[4]}:${m[5]}`;
    const dt = new Date(d);
    if (isNaN(dt.getTime())) return d; // leave unrecognized strings as-is
    d = dt;
  }
  if (!(d instanceof Date)) return '';
  return `${d.getFullYear()}-${pad2(d.getMonth()+1)}-${pad2(d.getDate())} ${pad2(d.getHours())}:${pad2(d.getMinutes())}`;
}

/* ---------- Cell value normalizer ----------
   Converts any Date to string; leaves primitives as strings;
   null/undefined become ''. Used everywhere we return values
   to the client.
-------------------------------------------- */
function normVal(v, kind){
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    if (kind === 'month')  return ym(v);
    if (kind === 'ymdhm')  return ymdhm(v);
    return ymd(v); // default: date only
  }
  if (kind === 'number') return Number(v) || 0;
  return String(v);
}

/* ---------- Parsing helpers ---------- */
function parseDate(v){
  if (!v) return null;
  if (v instanceof Date) return v;
  if (typeof v === 'string') {
    // yyyy-MM-dd or yyyy-MM-dd HH:mm
    const m = v.match(/^(\d{4})-(\d{2})-(\d{2})(?:[ T](\d{2}):(\d{2}))?/);
    if (m) {
      return new Date(Number(m[1]), Number(m[2])-1, Number(m[3]),
                      Number(m[4]||0), Number(m[5]||0));
    }
    const d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
  }
  return null;
}

function daysBetween(a,b){
  if (!a || !b) return '';
  const A = parseDate(a), B = parseDate(b);
  if (!A || !B) return '';
  const da = new Date(A.getFullYear(), A.getMonth(), A.getDate());
  const db = new Date(B.getFullYear(), B.getMonth(), B.getDate());
  return Math.floor((db - da) / 86400000);
}

/* =========================================================
   SETUP — auto-creates all sheets + settings
   ========================================================= */

function setupSpreadsheet() {
  const book = ss();

  ensureSheet('Clients', [
    'ClientID','ClientName','OnboardedDate','PolicyStart','PolicyEnd',
    'RelationManager','TeamLeader','Insurer','TPA',
    'CycleStart','CycleEnd','TotalDashboardsDue','DashboardsSent','DashboardsRemaining',
    'Status','CycleNumber','ParentClientID','Notes'
  ]);
  ensureSheet('Log', [
    'ClientID','ClientName','Month','MonthStartDate','CycleNo',
    'EmailsSent','Status','LastEmailDate','DaysSinceEmail','NeedsReminder','Notes','EmailLog'
  ]);
  ensureSheet('Attachments', [
    'AttachmentID','ClientID','FileName','FileURL','FileID','UploadedAt','UploadedBy','Remarks'
  ]);
  ensureSheet('EmailAudit', [
    'AuditID','ClientID','ClientName','Month','Timestamp','EmailsAfter','Notes','RecordedAt'
  ]);
  ensureSheet('Settings', ['Key','Value']);

  const set = book.getSheetByName('Settings');
  if (set.getLastRow() < 2) {
    const defaults = [
      ['DigestEmail',     Session.getActiveUser().getEmail() || ''],
      ['Timezone',        TIMEZONE],
      ['DailyDigest',     'ON'],
      ['WeeklyDigest',    'ON'],
      ['MonthlyDigest',   'ON'],
      ['QuarterlyDigest', 'ON'],
      ['YearlyDigest',    'ON'],
      ['ReminderDays',    String(REMINDER_DAYS)],
      ['AppURL',          '']
    ];
    set.getRange(2, 1, defaults.length, 2).setValues(defaults);
    set.getRange(1,1,1,2).setBackground('#1f2a44').setFontColor('#ffffff').setFontWeight('bold');
    set.setFrozenRows(1);
    set.autoResizeColumns(1, 2);
  }
}

function ensureSheet(name, headers) {
  const book = ss();
  let sh = book.getSheetByName(name);
  if (!sh) sh = book.insertSheet(name);

  const firstRow = sh.getRange(1,1,1,headers.length).getValues()[0];
  const isEmpty  = firstRow.every(v => v === '');
  if (isEmpty) {
    sh.getRange(1,1,1,headers.length).setValues([headers]);
    sh.getRange(1,1,1,headers.length)
      .setBackground('#1f2a44').setFontColor('#ffffff')
      .setFontWeight('bold').setVerticalAlignment('middle');
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, headers.length);
  }
  if (name === 'Log') {
    const lastCol = sh.getLastColumn();
    const hdr = sh.getRange(1,1,1,lastCol).getValues()[0];
    if (hdr.indexOf('EmailLog') === -1) {
      sh.getRange(1, lastCol + 1).setValue('EmailLog').setBackground('#1f2a44').setFontColor('#ffffff').setFontWeight('bold');
    }
  }
}

function getSetting(key) {
  const sh = ss().getSheetByName('Settings');
  if (!sh) return '';
  const data = sh.getDataRange().getValues().slice(1);
  for (const r of data) if (r[0] === key) return r[1];
  return '';
}

/* =========================================================
   WEB APP ENTRY
   ========================================================= */

function doGet() {
  setupSpreadsheet();
  const url = ScriptApp.getService().getUrl();
  const set = ss().getSheetByName('Settings');
  const rows = set.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) if (rows[i][0]==='AppURL'){ set.getRange(i+1,2).setValue(url); break; }

  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Dashboard Tracker')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/* =========================================================
   CLIENTS
   ========================================================= */

function getClients() {
  try {
    setupSpreadsheet();
    const sh = ss().getSheetByName('Clients');
    if (!sh) return [];

    const rows = sh.getDataRange().getValues();
    if (rows.length < 2) return [];

    const headers = rows.shift();
    const dateOnly = {
      OnboardedDate: true, PolicyStart: true, PolicyEnd: true,
      CycleStart: true, CycleEnd: true
    };
    const numeric = {
      TotalDashboardsDue: true, DashboardsSent: true,
      DashboardsRemaining: true, CycleNumber: true
    };

    return rows
      .filter(r => r[0] && String(r[0]).trim() !== '')
      .map(r => {
        const obj = {};
        headers.forEach((h, i) => {
          const v = r[i];
          if (v instanceof Date) obj[h] = dateOnly[h] ? ymd(v) : ymd(v);
          else if (v === null || v === undefined) obj[h] = '';
          else if (numeric[h]) obj[h] = Number(v) || 0;
          else if (typeof v === 'object') obj[h] = String(v);
          else obj[h] = v;
        });
        return obj;
      });
  } catch (e) {
    Logger.log('EXCEPTION in getClients: ' + e.toString());
    return [];
  }
}

function addClient(data) {
  setupSpreadsheet();
  const sh  = ss().getSheetByName('Clients');

  const existing = sh.getDataRange().getValues().slice(1);
  if (existing.some(r => String(r[0]) === String(data.ClientID))) {
    return { ok:false, error:'ClientID already exists: ' + data.ClientID };
  }

  const info = computeCycle(data.OnboardedDate, data.PolicyStart, data.PolicyEnd);

  sh.appendRow([
    data.ClientID, data.ClientName,
    parseDate(data.OnboardedDate), parseDate(data.PolicyStart), parseDate(data.PolicyEnd),
    data.RelationManager||'', data.TeamLeader||'', data.Insurer||'', data.TPA||'',
    info.cycleStart, info.cycleEnd, info.totalDue, 0, info.totalDue,
    'Active', data.CycleNumber||1, data.ParentClientID||data.ClientID,
    data.Notes||''
  ]);

  seedLogRows(data.ClientID, data.ClientName, info);
  recalcClients();
  return { ok:true, totalDue: info.totalDue };
}

function updateClient(data) {
  const sh = ss().getSheetByName('Clients');
  const log = ss().getSheetByName('Log');
  const rows = sh.getDataRange().getValues();
  let rowIdx = -1;
  for (let i=1;i<rows.length;i++) if (rows[i][0]===data.OriginalClientID){ rowIdx=i+1; break; }
  if (rowIdx===-1) return { ok:false, error:'Client not found' };

  const info = computeCycle(data.OnboardedDate, data.PolicyStart, data.PolicyEnd);
  const oldTotal = Number(rows[rowIdx-1][11])||0;

  sh.getRange(rowIdx,1,1,18).setValues([[
    data.ClientID, data.ClientName,
    parseDate(data.OnboardedDate), parseDate(data.PolicyStart), parseDate(data.PolicyEnd),
    data.RelationManager||'', data.TeamLeader||'', data.Insurer||'', data.TPA||'',
    info.cycleStart, info.cycleEnd, info.totalDue,
    rows[rowIdx-1][12],
    Math.max(0, info.totalDue - rows[rowIdx-1][12]),
    rows[rowIdx-1][14], rows[rowIdx-1][15], rows[rowIdx-1][16],
    data.Notes || rows[rowIdx-1][17] || ''
  ]]);

  if (oldTotal !== info.totalDue) {
    rebuildLogRows(data.OriginalClientID, data.ClientID, data.ClientName, info);
  } else if (data.ClientID !== data.OriginalClientID) {
    const logData = log.getDataRange().getValues();
    for (let i=1;i<logData.length;i++)
      if (logData[i][0]===data.OriginalClientID)
        log.getRange(i+1,1,1,2).setValues([[data.ClientID, data.ClientName]]);
  }
  return { ok:true };
}

function updateClientNotes(clientId, notes) {
  const sh = ss().getSheetByName('Clients');
  const rows = sh.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) if (rows[i][0]===clientId){ sh.getRange(i+1,18).setValue(notes); break; }
  return { ok:true };
}

function deleteClient(clientId) {
  const sh  = ss().getSheetByName('Clients');
  const log = ss().getSheetByName('Log');
  const att = ss().getSheetByName('Attachments');

  const data = sh.getDataRange().getValues();
  for (let i=data.length-1;i>=1;i--) if (data[i][0]===clientId){ sh.deleteRow(i+1); break; }

  const logData = log.getDataRange().getValues();
  for (let i=logData.length-1;i>=1;i--) if (logData[i][0]===clientId) log.deleteRow(i+1);

  const aData = att.getDataRange().getValues();
  for (let i=aData.length-1;i>=1;i--) {
    if (aData[i][1]===clientId) {
      try { DriveApp.getFileById(aData[i][4]).setTrashed(true); } catch(e){}
      att.deleteRow(i+1);
    }
  }
  return { ok:true };
}

function renewClient(originalClientId, newOnboarded, newPolicyStart, newPolicyEnd) {
  const sh = ss().getSheetByName('Clients');
  const rows = sh.getDataRange().getValues();
  let orig = null;
  for (let i=1;i<rows.length;i++) if (rows[i][0]===originalClientId){ orig=rows[i]; break; }
  if (!orig) return { ok:false, error:'Original not found' };

  const newCycleNum = (Number(orig[15])||1) + 1;
  const newId = `${orig[16]}-C${newCycleNum}`;

  return addClient({
    ClientID:newId, ClientName:orig[1],
    OnboardedDate:newOnboarded, PolicyStart:newPolicyStart, PolicyEnd:newPolicyEnd,
    RelationManager:orig[5], TeamLeader:orig[6], Insurer:orig[7], TPA:orig[8],
    CycleNumber:newCycleNum, ParentClientID:orig[16], Notes:orig[17]||''
  });
}

function computeCycle(onbStr, psStr, peStr) {
  const onb = parseDate(onbStr), ps = parseDate(psStr), pe = parseDate(peStr);
  let cycleStart = new Date(onb.getFullYear(), onb.getMonth(), 1);
  if (onb.getDate() > 20) cycleStart = new Date(onb.getFullYear(), onb.getMonth()+1, 1);
  const psMonth = new Date(ps.getFullYear(), ps.getMonth(), 1);
  if (cycleStart < psMonth) cycleStart = psMonth;
  const cycleEnd = new Date(pe.getFullYear(), pe.getMonth()+1, 0);
  const totalDue = Math.max(1, Math.min(12, monthsBetween(cycleStart, cycleEnd) + 1));
  return { cycleStart, cycleEnd, totalDue };
}
function monthsBetween(a,b){ return (b.getFullYear()-a.getFullYear())*12 + (b.getMonth()-a.getMonth()); }

function seedLogRows(clientId, clientName, info) {
  const log = ss().getSheetByName('Log');
  const rows = [];
  for (let i=0;i<info.totalDue;i++) {
    const d = new Date(info.cycleStart.getFullYear(), info.cycleStart.getMonth()+i, 1);
    // Write Month as a string 'yyyy-MM' so future reads don't need Date parsing.
    rows.push([clientId, clientName, ym(d), d, i+1, 0, 'W', '', '', 'NO', '', '']);
  }
  if (rows.length) log.getRange(log.getLastRow()+1, 1, rows.length, 12).setValues(rows);
}

function rebuildLogRows(clientId, newId, newName, info) {
  const log = ss().getSheetByName('Log');
  const data = log.getDataRange().getValues();
  for (let i=data.length-1;i>=1;i--) if (data[i][0]===clientId) log.deleteRow(i+1);
  seedLogRows(newId, newName, info);
}

/* =========================================================
   EMAIL TIMESTAMPS
   ========================================================= */

function appendEmailTimestamp(clientId, month, timestampStr, notes) {
  const log = ss().getSheetByName('Log');
  const audit = ss().getSheetByName('EmailAudit');
  const data = log.getDataRange().getValues();

  for (let i=1;i<data.length;i++) {
    const rowMonth = normVal(data[i][2], 'month');
    if (data[i][0]===clientId && rowMonth===month) {
      const prev = data[i][11] ? String(data[i][11]).split(' | ').filter(Boolean) : [];
      prev.push(timestampStr);
      prev.sort();
      log.getRange(i+1, 12).setValue(prev.join(' | '));

      const latest = prev[prev.length-1];
      log.getRange(i+1, 8).setValue(latest);
      log.getRange(i+1, 9).setValue(daysBetween(latest, new Date()));

      const id = 'AUD' + new Date().getTime() + Math.floor(Math.random()*1000);
      audit.appendRow([id, clientId, String(data[i][1]), month, timestampStr, prev.length, notes||'', new Date()]);
      return { ok:true, count: prev.length };
    }
  }
  return { ok:false };
}

function trimEmailTimestamps(clientId, month, k) {
  const log = ss().getSheetByName('Log');
  const audit = ss().getSheetByName('EmailAudit');
  const data = log.getDataRange().getValues();

  for (let i=1;i<data.length;i++) {
    const rowMonth = normVal(data[i][2], 'month');
    if (data[i][0]===clientId && rowMonth===month) {
      let list = data[i][11] ? String(data[i][11]).split(' | ').filter(Boolean) : [];
      const removed = list.splice(Math.max(0, list.length - k), k);
      log.getRange(i+1, 12).setValue(list.join(' | '));

      const latest = list.length ? list[list.length-1] : '';
      log.getRange(i+1, 8).setValue(latest);
      log.getRange(i+1, 9).setValue(latest ? daysBetween(latest, new Date()) : '');

      removed.forEach(ts=>{
        const id = 'AUD' + new Date().getTime() + Math.floor(Math.random()*1000);
        audit.appendRow([id, clientId, String(data[i][1]), month, ts, list.length, 'trimmed', new Date()]);
      });
      return { ok:true, remaining: list.length };
    }
  }
  return { ok:false };
}

function api_addEmailTimestamp(clientId, month, timestampStr, notes) {
  return appendEmailTimestamp(clientId, month, timestampStr, notes);
}

function api_getEmailLog(clientId, month) {
  const log = ss().getSheetByName('Log');
  const data = log.getDataRange().getValues().slice(1);
  for (const r of data) {
    const rowMonth = normVal(r[2], 'month');
    if (r[0]===clientId && rowMonth===month) {
      const list = r[11] ? String(r[11]).split(' | ').filter(Boolean) : [];
      return { list };
    }
  }
  return { list: [] };
}

/* =========================================================
   MONTHLY UPDATE
   ========================================================= */

function getMonthView(monthLabel) {
  setupSpreadsheet();
  const log = ss().getSheetByName('Log');
  const data = log.getDataRange().getValues().slice(1);
  return data
    .filter(r => normVal(r[2], 'month') === monthLabel)
    .map(r => {
      const lastEmail = r[7] ? normVal(r[7], 'ymdhm') : '';
      return {
        ClientID:   String(r[0]),
        ClientName: String(r[1]),
        Month:      normVal(r[2], 'month'),
        CycleNo:    Number(r[4])||0,
        EmailsSent: Number(r[5])||0,
        Status:     String(r[6]||''),
        LastEmailDate: lastEmail ? ymd(lastEmail) : '',
        DaysSinceEmail: lastEmail ? daysBetween(lastEmail, new Date()) : '',
        NeedsReminder: String(r[9]||'NO'),
        Notes:      String(r[10]||''),
        EmailLog:   String(r[11]||'')
      };
    });
}

function api_prepareMonthUpdate(rows) {
  const log = ss().getSheetByName('Log');
  const data = log.getDataRange().getValues();
  const needsTimestamps = [];

  rows.forEach(upd=>{
    for (let i=1;i<data.length;i++) {
      const rowMonth = normVal(data[i][2], 'month');
      if (data[i][0]===upd.ClientID && rowMonth===upd.Month) {
        const prev = Number(data[i][5])||0;
        const next = Number(upd.EmailsSent)||0;
        const diff = next - prev;
        if (diff > 0) {
          needsTimestamps.push({
            ClientID: upd.ClientID,
            ClientName: String(data[i][1]),
            Month: upd.Month,
            Count: diff,
            Existing: data[i][11] ? String(data[i][11]).split(' | ').filter(Boolean) : []
          });
        }
        break;
      }
    }
  });
  return { needsTimestamps };
}

function updateMonth(rows, timestampsMap) {
  setupSpreadsheet();
  const log = ss().getSheetByName('Log');
  const data = log.getDataRange().getValues();
  const today = new Date();
  timestampsMap = timestampsMap || {};

  Object.keys(timestampsMap).forEach(k=>{
    timestampsMap[k] = (timestampsMap[k]||[]).filter(Boolean).sort();
  });

  rows.forEach(upd=>{
    for (let i=1;i<data.length;i++) {
      const rowMonth = normVal(data[i][2], 'month');
      if (data[i][0]===upd.ClientID && rowMonth===upd.Month) {
        const prev = Number(data[i][5])||0;
        const next = Number(upd.EmailsSent)||0;
        const diff = next - prev;

        log.getRange(i+1, 6).setValue(next);
        log.getRange(i+1, 7).setValue(upd.Status);
        log.getRange(i+1, 11).setValue(upd.Notes || data[i][10] || '');

        const key = upd.ClientID + '|' + upd.Month;

        if (diff > 0) {
          const list = data[i][11] ? String(data[i][11]).split(' | ').filter(Boolean) : [];
          const toAdd = (timestampsMap[key] && timestampsMap[key].length === diff)
            ? timestampsMap[key]
            : Array(diff).fill(ymdhm(today));
          toAdd.forEach(ts=> list.push(ts));
          list.sort();
          log.getRange(i+1, 12).setValue(list.join(' | '));
          const latest = list[list.length-1];
          log.getRange(i+1, 8).setValue(latest);
          log.getRange(i+1, 9).setValue(daysBetween(latest, today));
          toAdd.forEach(ts=>{
            const id='AUD'+new Date().getTime()+Math.floor(Math.random()*1000);
            ss().getSheetByName('EmailAudit').appendRow(
              [id, upd.ClientID, String(data[i][1]), upd.Month, ts, list.length, '', new Date()]);
          });
        }
        else if (diff < 0) {
          trimEmailTimestamps(upd.ClientID, upd.Month, -diff);
        }
        break;
      }
    }
  });

  recalcClients();
  return { ok:true };
}

/* =========================================================
   RECALC + REMINDERS
   ========================================================= */

function recalcClients() {
  setupSpreadsheet();
  const clientsSh = ss().getSheetByName('Clients');
  const logData   = ss().getSheetByName('Log').getDataRange().getValues().slice(1);
  const cdata     = clientsSh.getDataRange().getValues();
  const today     = new Date();

  for (let i=1;i<cdata.length;i++) {
    const id = cdata[i][0];
    const sent = logData.filter(r=>r[0]===id && r[6]==='✓').length;
    const remaining = Math.max(0, Number(cdata[i][11]) - sent);
    const policyEnd = parseDate(cdata[i][4]);
    const status = (policyEnd && today > policyEnd) ? 'Expired' : 'Active';
    clientsSh.getRange(i+1, 13, 1, 3).setValues([[sent, remaining, status]]);
  }
  updateReminderFlags();
}

function updateReminderFlags() {
  const log = ss().getSheetByName('Log');
  const data = log.getDataRange().getValues();
  const today = new Date();
  const days = Number(getSetting('ReminderDays')) || REMINDER_DAYS;
  const col = [];
  for (let i=1;i<data.length;i++) {
    const status = data[i][6];
    const list = data[i][11] ? String(data[i][11]).split(' | ').filter(Boolean) : [];
    const lastStr = list.length ? list[list.length-1] : '';
    const last = parseDate(lastStr);
    let flag = 'NO';
    if ((status==='P'||status==='R') && last && daysBetween(last, today) >= days) flag = 'YES';
    col.push([flag]);
  }
  if (col.length) log.getRange(2, 10, col.length, 1).setValues(col);
}

function getReminders() {
  setupSpreadsheet();
  updateReminderFlags();
  const log = ss().getSheetByName('Log');
  const data = log.getDataRange().getValues().slice(1);
  return data
    .filter(r => r[9]==='YES')
    .map(r => {
      const list = r[11] ? String(r[11]).split(' | ').filter(Boolean) : [];
      const last = list.length ? list[list.length-1] : '';
      return {
        ClientID:   String(r[0]),
        ClientName: String(r[1]),
        Month:      normVal(r[2], 'month'),
        CycleNo:    Number(r[4])||0,
        EmailsSent: Number(r[5])||0,
        Status:     String(r[6]||''),
        LastEmailDate: last ? ymd(last) : '',
        DaysSinceEmail: last ? daysBetween(last, new Date()) : ''
      };
    });
}

/* =========================================================
   TIMELINE
   ========================================================= */

function getTimeline(clientId) {
  setupSpreadsheet();
  const log = ss().getSheetByName('Log');
  const data = log.getDataRange().getValues().slice(1);
  return data
    .filter(r => r[0]===clientId)
    .map(r => {
      const lastEmail = r[7] ? normVal(r[7], 'ymdhm') : '';
      return {
        Month:      normVal(r[2], 'month'),
        CycleNo:    Number(r[4])||0,
        EmailsSent: Number(r[5])||0,
        Status:     String(r[6]||''),
        LastEmailDate: lastEmail ? ymd(lastEmail) : '',
        EmailLog:   String(r[11]||''),
        Notes:      String(r[10]||'')
      };
    });
}

/* =========================================================
   ATTACHMENTS
   ========================================================= */

function getAttachmentFolder() {
  const it = DriveApp.getFoldersByName(ATTACH_FOLDER);
  return it.hasNext() ? it.next() : DriveApp.createFolder(ATTACH_FOLDER);
}

function uploadAttachment(clientId, fileName, mimeType, base64Data, remarks) {
  const folder = getAttachmentFolder();
  const blob = Utilities.newBlob(Utilities.base64Decode(base64Data), mimeType, fileName);
  const file = folder.createFile(blob);

  const sh = ss().getSheetByName('Attachments');
  const id = 'ATT' + new Date().getTime();
  sh.appendRow([id, clientId, fileName, file.getUrl(), file.getId(), new Date(), Session.getActiveUser().getEmail(), remarks||'']);
  return { ok:true, id, url:file.getUrl() };
}

function getAttachments(clientId) {
  setupSpreadsheet();
  const sh = ss().getSheetByName('Attachments');
  const rows = sh.getDataRange().getValues().slice(1);
  return rows
    .filter(r => r[1]===clientId)
    .map(r => ({
      AttachmentID: String(r[0]),
      ClientID:     String(r[1]),
      FileName:     String(r[2]),
      FileURL:      String(r[3]),
      FileID:       String(r[4]),
      UploadedAt:   r[5] ? ymd(r[5]) : '',
      UploadedBy:   String(r[6]||''),
      Remarks:      String(r[7]||'')
    }));
}

function deleteAttachment(attachmentId) {
  const sh = ss().getSheetByName('Attachments');
  const rows = sh.getDataRange().getValues();
  for (let i=1;i<rows.length;i++) {
    if (rows[i][0]===attachmentId) {
      try { DriveApp.getFileById(rows[i][4]).setTrashed(true); } catch(e){}
      sh.deleteRow(i+1); break;
    }
  }
  return { ok:true };
}

/* =========================================================
   REPORTS + CHARTS
   ========================================================= */

function getReports() {
  setupSpreadsheet();
  const log = ss().getSheetByName('Log').getDataRange().getValues().slice(1);
  const now = new Date();
  const month = ym(now);
  const year  = now.getFullYear();
  const q = Math.floor(now.getMonth()/3)+1;
  const qStart = new Date(year,(q-1)*3,1);
  const qEnd   = new Date(year,q*3,0,23,59,59);
  const weekAgo = new Date(now); weekAgo.setDate(now.getDate()-7);

  const inRange = (d,a,b)=> { const D = parseDate(d); return D && D>=a && D<=b; };
  return {
    weekly:{
      emails: log.filter(r=>inRange(r[3], weekAgo, now)).reduce((s,r)=>s+(Number(r[5])||0),0),
      sent:   log.filter(r=>inRange(r[3], weekAgo, now) && r[6]==='✓').length
    },
    monthly:{
      emails: log.filter(r=>normVal(r[2],'month')===month).reduce((s,r)=>s+(Number(r[5])||0),0),
      sent:   log.filter(r=>normVal(r[2],'month')===month && r[6]==='✓').length,
      pending:log.filter(r=>normVal(r[2],'month')===month && (r[6]==='P'||r[6]==='R')).length
    },
    quarterly:{
      emails: log.filter(r=>inRange(r[3], qStart, qEnd)).reduce((s,r)=>s+(Number(r[5])||0),0),
      sent:   log.filter(r=>inRange(r[3], qStart, qEnd) && r[6]==='✓').length
    },
    yearly:{
      emails: log.filter(r=>{ const D=parseDate(r[3]); return D && D.getFullYear()===year; }).reduce((s,r)=>s+(Number(r[5])||0),0),
      sent:   log.filter(r=>{ const D=parseDate(r[3]); return D && D.getFullYear()===year && r[6]==='✓'; }).length
    }
  };
}

function getChartData() {
  setupSpreadsheet();
  const log = ss().getSheetByName('Log').getDataRange().getValues().slice(1);
  const clients = getClients();
  const monthsMap = {}, statusMap = {W:0,P:0,R:0,'✓':0,'✗':0}, rmMap = {};

  log.forEach(r=>{
    const m = normVal(r[2], 'month');
    if (!m) return;
    monthsMap[m] = monthsMap[m] || {month:m, sent:0, emails:0, pending:0};
    if (r[6]==='✓') monthsMap[m].sent++;
    if (r[6]==='P'||r[6]==='R') monthsMap[m].pending++;
    monthsMap[m].emails += Number(r[5])||0;
    if (statusMap[r[6]]!==undefined) statusMap[r[6]]++;
  });

  clients.forEach(c=>{
    const rm = c.RelationManager || 'Unassigned';
    rmMap[rm] = rmMap[rm] || {rm, sent:0, pending:0};
    rmMap[rm].sent    += Number(c.DashboardsSent)||0;
    rmMap[rm].pending += Number(c.DashboardsRemaining)||0;
  });

  const monthly = Object.values(monthsMap).sort((a,b)=>a.month.localeCompare(b.month)).slice(-12);
  return {
    monthlyTrend: monthly,
    statusBreakdown: [
      ['Status','Count'],
      ['Sent',statusMap['✓']],
      ['Pending',statusMap['P']],
      ['Review',statusMap['R']],
      ['Waiting',statusMap['W']],
      ['N/A',statusMap['✗']]
    ],
    byRM: [['RM','Sent','Pending'], ...Object.values(rmMap).map(o=>[o.rm,o.sent,o.pending])]
  };
}

/* =========================================================
   DIGESTS
   ========================================================= */

function buildDigestHtml(sections) {
  const url = getSetting('AppURL') || '#';
  return `
  <div style="font-family:system-ui,Arial,sans-serif;max-width:640px;margin:auto;background:#f4f6fb;padding:0">
    <div style="background:#1f2a44;color:#fff;padding:18px 22px;border-radius:12px 12px 0 0">
      <h2 style="margin:0">📊 Dashboard Tracker Digest</h2>
      <div style="opacity:.8;font-size:13px;margin-top:4px">${Utilities.formatDate(new Date(), tz(), 'EEE, dd MMM yyyy HH:mm')} IST</div>
    </div>
    <div style="background:#fff;padding:20px 22px">
      ${sections}
    </div>
    <div style="text-align:center;padding:18px">
      <a href="${url}" style="background:#2c7be5;color:#fff;padding:10px 22px;border-radius:8px;text-decoration:none;font-weight:600">Open Tracker →</a>
    </div>
    <div style="text-align:center;font-size:12px;color:#999;padding-bottom:20px">
      Sent by your Dashboard Tracker • Toggle digests in the Settings sheet
    </div>
  </div>`;
}

function kpi(label, value, color) {
  return `<div style="display:inline-block;min-width:130px;background:#f4f6fb;border-radius:10px;padding:12px 16px;margin:6px 6px 6px 0">
    <div style="font-size:11px;color:#777;text-transform:uppercase;letter-spacing:.04em">${label}</div>
    <div style="font-size:22px;font-weight:700;color:${color||'#1f2a44'}">${value}</div>
  </div>`;
}
function section(title, body) {
  return `<h3 style="margin:22px 0 8px;color:#1f2a44;font-size:15px">${title}</h3>${body}`;
}

function sendDigest(kind) {
  const email = getSetting('DigestEmail') || Session.getActiveUser().getEmail();
  const log = ss().getSheetByName('Log').getDataRange().getValues().slice(1);
  const clients = getClients();
  const now = new Date();

  let rangeStart, title, subject;
  if (kind==='daily')   { const d=new Date(now); d.setDate(now.getDate()-1); rangeStart=d; title='Last 24 Hours'; subject='Daily Digest'; }
  if (kind==='weekly')  { const d=new Date(now); d.setDate(now.getDate()-7); rangeStart=d; title='Last 7 Days'; subject='Weekly Digest'; }
  if (kind==='monthly') { rangeStart=new Date(now.getFullYear(), now.getMonth()-1, 1); title='Last Month'; subject='Monthly Digest'; }
  if (kind==='quarterly'){ const q=Math.floor(now.getMonth()/3); rangeStart=new Date(now.getFullYear(), (q-1)*3, 1); title='Last Quarter'; subject='Quarterly Digest'; }
  if (kind==='yearly')  { rangeStart=new Date(now.getFullYear()-1, 0, 1); title='Last Year'; subject='Yearly Digest'; }

  const latestTs = r => { const l = r[11] ? String(r[11]).split(' | ').filter(Boolean) : []; return l.length ? l[l.length-1] : null; };
  const inRange = r => { const ts = latestTs(r); const T = parseDate(ts); return T && T >= rangeStart; };

  const scoped = log.filter(inRange);
  const emails = scoped.reduce((s,r)=> s + (Number(r[5])||0), 0);
  const sent   = scoped.filter(r=>r[6]==='✓').length;
  const pending= scoped.filter(r=>r[6]==='P'||r[6]==='R').length;

  const days = Number(getSetting('ReminderDays')||REMINDER_DAYS);
  const reminders = log.filter(r=>{
    if (!(r[6]==='P'||r[6]==='R')) return false;
    const ts = latestTs(r);
    if (!ts) return false;
    return daysBetween(ts, now) >= days;
  }).slice(0,10);

  const newClients = clients.filter(c=>{
    const onb = parseDate(c.OnboardedDate);
    return onb && onb >= rangeStart;
  });

  const expiring = clients.filter(c=>{
    const pe = parseDate(c.PolicyEnd);
    if (!pe) return false;
    const diff = daysBetween(now, pe);
    return diff>=0 && diff<=30;
  });

  let html = '';
  html += section('Summary — '+title,
    kpi('Emails', emails, '#2c7be5') + kpi('Dashboards Sent', sent, '#27ae60') +
    kpi('Pending/Review', pending, '#f39c12')
  );

  if (reminders.length) {
    html += section('⚠️ Needs Follow-up',
      '<table style="width:100%;border-collapse:collapse;font-size:13px">'+
      '<tr style="background:#fafbfd"><th style="text-align:left;padding:6px">Client</th>'+
      '<th style="text-align:left;padding:6px">Month</th><th style="text-align:left;padding:6px">Status</th>'+
      '<th style="text-align:left;padding:6px">Days</th></tr>'+
      reminders.map(r=>{
        const ts = latestTs(r);
        return `<tr><td style="padding:6px;border-top:1px solid #eef0f5">${r[1]}</td>
        <td style="padding:6px;border-top:1px solid #eef0f5">${normVal(r[2],'month')}</td>
        <td style="padding:6px;border-top:1px solid #eef0f5">${r[6]}</td>
        <td style="padding:6px;border-top:1px solid #eef0f5"><b>${daysBetween(ts, now)}d</b></td></tr>`;
      }).join('')+
      '</table>');
  }

  if (newClients.length) {
    html += section('🆕 Newly Onboarded',
      '<ul style="margin:0;padding-left:18px;font-size:13px">'+
      newClients.map(c=>`<li>${c.ClientName} — onboarded ${c.OnboardedDate}</li>`).join('')+
      '</ul>');
  }

  if (expiring.length) {
    html += section('🔔 Expiring in 30 Days',
      '<ul style="margin:0;padding-left:18px;font-size:13px">'+
      expiring.map(c=>`<li>${c.ClientName} — policy ends ${c.PolicyEnd}</li>`).join('')+
      '</ul>');
  }

  MailApp.sendEmail({
    to: email,
    subject: `📊 ${subject} — ${Utilities.formatDate(now, tz(), 'dd MMM yyyy')}`,
    htmlBody: buildDigestHtml(html)
  });
}

function runDailyDigest()     { if(getSetting('DailyDigest')==='ON')     sendDigest('daily'); }
function runWeeklyDigest()    { if(getSetting('WeeklyDigest')==='ON')    sendDigest('weekly'); }
function runMonthlyDigest()   { if(getSetting('MonthlyDigest')==='ON')   sendDigest('monthly'); }
function runQuarterlyDigest() {
  if (getSetting('QuarterlyDigest')!=='ON') return;
  const m = new Date().getMonth() + 1;
  if (![1,4,7,10].includes(m)) return;
  sendDigest('quarterly');
}
function runYearlyDigest() {
  if (getSetting('YearlyDigest')!=='ON') return;
  if (new Date().getMonth() !== 0) return;
  sendDigest('yearly');
}

/* =========================================================
   TRIGGERS — run once manually
   ========================================================= */

function installTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('runDailyDigest').timeBased()
    .atHour(8).everyDays(1).inTimezone(TIMEZONE).create();

  ScriptApp.newTrigger('runWeeklyDigest').timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(8).inTimezone(TIMEZONE).create();

  ScriptApp.newTrigger('runMonthlyDigest').timeBased()
    .onMonthDay(1).atHour(8).inTimezone(TIMEZONE).create();

  ScriptApp.newTrigger('runQuarterlyDigest').timeBased()
    .onMonthDay(1).atHour(8).inTimezone(TIMEZONE).create();

  ScriptApp.newTrigger('runYearlyDigest').timeBased()
    .onMonthDay(1).atHour(8).inTimezone(TIMEZONE).create();
}

/* =========================================================
   DIAGNOSTICS
   ========================================================= */

function diagnose() {
  const book = ss();
  Logger.log('Tabs: ' + book.getSheets().map(s=>s.getName()).join(', '));

  const sh = book.getSheetByName('Clients');
  if (!sh) { Logger.log('❌ No Clients tab'); return; }
  Logger.log('Clients last row/col: ' + sh.getLastRow() + '/' + sh.getLastColumn());

  const log = book.getSheetByName('Log');
  if (!log) { Logger.log('❌ No Log tab'); return; }
  Logger.log('Log last row/col: ' + log.getLastRow() + '/' + log.getLastColumn());

  // Sample the first two Log rows to see Month type
  if (log.getLastRow() >= 2) {
    const sample = log.getRange(2, 1, Math.min(3, log.getLastRow()-1), 12).getValues();
    sample.forEach((r,i)=>{
      Logger.log('Log row '+(i+2)+': Month value = '+JSON.stringify(r[2])+
                 ' typeof='+typeof r[2]+
                 ' → normVal = '+normVal(r[2],'month'));
    });
  }

  const clients = getClients();
  Logger.log('getClients() → '+clients.length+' clients');

  const mv = getMonthView(ym(new Date()));
  Logger.log('getMonthView('+ym(new Date())+') → '+mv.length+' rows');
}

function fixLogMonthColumn() {
  // One-time repair: convert any Date-typed Month cells in Log → text 'yyyy-MM'.
  // Safe to run multiple times.
  const log = ss().getSheetByName('Log');
  if (!log || log.getLastRow() < 2) return 'No Log rows';
  const range = log.getRange(2, 3, log.getLastRow()-1, 1);
  const vals = range.getValues();
  let changed = 0;
  const out = vals.map(row=>{
    const v = row[0];
    if (v instanceof Date) { changed++; return [ym(v)]; }
    if (typeof v === 'string' && /^\d{4}-\d{2}$/.test(v)) return [v];
    if (typeof v === 'string' && v.length) { const D = parseDate(v); return [D ? ym(D) : v]; }
    return [v];
  });
  range.setValues(out);
  Logger.log('Log Month cells converted to text: '+changed);
  return 'Converted '+changed+' cells';
}

function testWebAppCall2() {
  const result = getClients();
  const serialized = JSON.stringify(result);
  Logger.log('Clients serialized OK, length: '+serialized.length);
  const m = getMonthView(ym(new Date()));
  Logger.log('MonthView serialized OK, length: '+JSON.stringify(m).length);
}

function whereAmI() {
  Logger.log('Web app URL: ' + ScriptApp.getService().getUrl());
  Logger.log('Clients: ' + getClients().length);
  Logger.log('MonthView('+ym(new Date())+'): ' + getMonthView(ym(new Date())).length);
}

function helloWorld() { return 'hi from backend'; }
