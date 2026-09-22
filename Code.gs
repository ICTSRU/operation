/**
 * الخطة التشغيلية 2026/2027 — الإدارة التنفيذية للاتصالات وتقنية المعلومات
 * Google Apps Script Web App backend  —  Version v1.1
 *
 * التركيب:
 *  1. أنشئ Google Sheet جديد باسم: "الخطة التشغيلية 2026-2027 - ICTD"
 *  2. Extensions ▸ Apps Script ▸ الصق هذا الملف
 *  3. Deploy ▸ New deployment ▸ Web app
 *       Execute as:  Me
 *       Who has access:  Anyone
 *  4. انسخ رابط /exec وضعه في المتغير SHEET_URL داخل ملف HTML
 */

var SHEET_NAME = 'الخطة التشغيلية';
var META_SHEET = 'بيانات الإدارة';

var FIELDS = ['strategicGoal','operationalGoal','initiative','priority','status','startDate','endDate',
              'activities','activityDates','kpi','targetValue','budget','resources','risks','mitigation'];

var HEADERS = ['معرف السجل',' الهدف الإستراتيجي','الهدف التشغيلي','المبادرة الإستراتيجية','الأولوية','حالة التنفيذ',
               'تاريخ البداية','تاريخ النهاية','الأنشطة/المشاريع التفصيلية للمبادرة',
               'موعد إتمام الأنشطة / المشاريع (July 2026 - Jun 2027)','مؤشر الأداء (تشغيلي لقياس إنجاز المبادرة)',
               'القيمة المستهدفة (لمؤشر الأداء)','الميزانية','الموارد المطلوبة','المخاطر المحتملة','خطة المعالجة',
               'آخر تحديث'];

/* ------------------------------------------------------------------ */

function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.appendRow(HEADERS);
    var hr = sh.getRange(1, 1, 1, HEADERS.length);
    hr.setBackground('#00b0f0').setFontWeight('bold').setFontSize(11)
      .setHorizontalAlignment('center').setVerticalAlignment('middle').setWrap(true);
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  }
  return sh;
}

function getMetaSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(META_SHEET);
  if (!sh) {
    sh = ss.insertSheet(META_SHEET);
    sh.setRightToLeft(true);
    sh.getRange('A1:B1').setValues([['البيان', 'القيمة']])
      .setBackground('#7030a0').setFontColor('#ffffff').setFontWeight('bold');
    sh.getRange('A2:A4').setValues([['العام الأكاديمي'], ['الإدارة / القسم'], ['تاريخ إعداد الخطة']]);
    sh.setColumnWidth(1, 200); sh.setColumnWidth(2, 420);
  }
  return sh;
}

function rowToArray_(r) {
  var a = [r.id || ('r' + new Date().getTime())];
  FIELDS.forEach(function (f) { a.push(r[f] == null ? '' : String(r[f])); });
  a.push(Utilities.formatDate(new Date(), 'Asia/Riyadh', 'yyyy-MM-dd HH:mm'));
  return a;
}

function saveMeta_(meta) {
  if (!meta) return;
  var sh = getMetaSheet_();
  sh.getRange('B2:B4').setValues([[meta.academicYear || ''], [meta.department || ''], [meta.planDate || '']]);
}

function styleData_(sh) {
  var last = sh.getLastRow();
  if (last < 2) return;
  sh.getRange(2, 1, last - 1, HEADERS.length)
    .setVerticalAlignment('middle').setWrap(true).setFontSize(10);
  sh.setColumnWidth(4, 260);   // المبادرة
  sh.setColumnWidth(9, 420);   // الأنشطة
  sh.setColumnWidth(12, 320);  // القيمة المستهدفة
  sh.setColumnWidth(16, 300);  // خطة المعالجة
}

/* ------------------------------------------------------------------ */

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || 'list';
    if (action !== 'list') return json_({ ok: false, error: 'إجراء غير معروف: ' + action });

    var sh = getSheet_();
    var values = sh.getDataRange().getValues();
    var out = [];
    for (var i = 1; i < values.length; i++) {
      var v = values[i];
      if (!v[0] && !v[3]) continue;
      var o = { id: String(v[0]) };
      FIELDS.forEach(function (f, k) { o[f] = v[k + 1] instanceof Date
        ? Utilities.formatDate(v[k + 1], 'Asia/Riyadh', 'yyyy-MM-dd') : String(v[k + 1] || ''); });
      out.push(o);
    }
    var ms = getMetaSheet_().getRange('B2:B4').getValues();
    return json_({ ok: true, rows: out, meta: {
      academicYear: String(ms[0][0] || ''), department: String(ms[1][0] || ''),
      planDate: ms[2][0] instanceof Date ? Utilities.formatDate(ms[2][0], 'Asia/Riyadh', 'yyyy-MM-dd') : String(ms[2][0] || '')
    }});
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var body = JSON.parse(e.postData.contents);
    var sh = getSheet_();
    saveMeta_(body.meta);

    if (body.action === 'replaceAll') {
      var list = body.rows || [];
      if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, HEADERS.length).clearContent();
      if (list.length) {
        sh.getRange(2, 1, list.length, HEADERS.length).setValues(list.map(rowToArray_));
      }
      styleData_(sh);
      return json_({ ok: true, count: list.length });
    }

    if (body.action === 'upsert') {
      var r = body.row || {};
      var arr = rowToArray_(r);
      var ids = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues() : [];
      var target = -1;
      for (var i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(arr[0])) { target = i + 2; break; }
      if (target === -1) target = sh.getLastRow() + 1;
      sh.getRange(target, 1, 1, HEADERS.length).setValues([arr]);
      styleData_(sh);
      return json_({ ok: true, row: target });
    }

    if (body.action === 'delete') {
      var id = String(body.id || '');
      var ids2 = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 1).getValues() : [];
      for (var j = 0; j < ids2.length; j++) if (String(ids2[j][0]) === id) { sh.deleteRow(j + 2); break; }
      return json_({ ok: true });
    }

    return json_({ ok: false, error: 'إجراء غير معروف' });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
