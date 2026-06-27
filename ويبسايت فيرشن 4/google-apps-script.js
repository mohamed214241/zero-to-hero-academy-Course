/**
 * ═══════════════════════════════════════════════════
 *  Zero To Hero — Google Apps Script (FIXED)
 * ═══════════════════════════════════════════════════
 *  تعليمات النشر:
 *  1. Extensions → Apps Script → امسح الكود القديم والصق ده
 *  2. Deploy → New Deployment (مش Manage — لازم New)
 *  3. Type: Web App
 *  4. Execute as: Me
 *  5. Who has access: Anyone  ← مهم جداً
 *  6. انسخ الـ URL الجديد
 *  7. ضعه في admin.html و register.html في SHEETS_URL
 * ═══════════════════════════════════════════════════
 */

const SHEET_NAME = 'Registrations';

// ── GET — للقراءة من الأدمن + استقبال تسجيل جديد عبر JSONP ──
function doGet(e) {
  const callback = e && e.parameter && e.parameter.callback;
  const action   = e && e.parameter && e.parameter.action;

  // ping — اختبار الاتصال
  if (action === 'ping') {
    const out = JSON.stringify({ ok: true });
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + out + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
  }

  // add — استقبال تسجيل جديد عبر JSONP GET (backup لما POST يفشل)
  if (action === 'add') {
    var p = e.parameter;
    var result = addRegistration({
      name:         p.name        || '',
      first:        p.first       || '',
      last:         p.last        || '',
      email:        p.email       || '',
      phone:        p.phone       || '',
      payMethod:    p.payMethod   || '',
      finalPrice:   p.finalPrice  || 2500,
      discountCode: p.discountCode|| '',
      status:       'pending',
      time:         p.time        || new Date().toISOString()
    });
    var out = JSON.stringify(result);
    if (callback) {
      return ContentService.createTextOutput(callback + '(' + out + ')')
        .setMimeType(ContentService.MimeType.JAVASCRIPT);
    }
    return ContentService.createTextOutput(out).setMimeType(ContentService.MimeType.JSON);
  }

  // default — إرجاع كل التسجيلات للأدمن
  const data = getRegistrations();
  const json = JSON.stringify(data);

  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

// ── POST — لإضافة تسجيل أو تحديث الحالة ──
function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const action = data.action || 'add';
    if (action === 'add')          return jsonOut(addRegistration(data));
    if (action === 'updateStatus') return jsonOut(updateStatus(data.rowIndex, data.status));
    return jsonOut({ error: 'unknown action' });
  } catch(err) {
    return jsonOut({ error: err.message });
  }
}

function jsonOut(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── الشيت ──
function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.getRange(1, 1, 1, 10).setValues([[
      'الاسم الكامل', 'الاسم الأول', 'اسم العائلة',
      'الإيميل', 'الواتساب', 'طريقة الدفع',
      'المبلغ', 'كود الخصم', 'الحالة', 'التوقيت'
    ]]);
    sh.getRange(1, 1, 1, 10).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function getRegistrations() {
  const sh   = getSheet();
  const data = sh.getDataRange().getValues();
  // تخطي السطر الأول (headers)
  return data.slice(1).map(function(row) {
    return row.map(function(c) {
      return (c === null || c === undefined) ? '' : String(c);
    });
  });
}

function addRegistration(d) {
  getSheet().appendRow([
    d.name        || '',
    d.first       || '',
    d.last        || '',
    d.email       || '',
    d.phone       || '',
    d.payMethod   || '',
    d.finalPrice  || 2500,
    d.discountCode|| '',
    d.status      || 'pending',
    d.time        || new Date().toISOString()
  ]);

  // إشعار بالإيميل (اختياري)
  try {
    MailApp.sendEmail({
      to: Session.getActiveUser().getEmail(),
      subject: '🎬 تسجيل جديد: ' + (d.name || ''),
      body: [
        'الاسم: '      + (d.name      || ''),
        'الواتساب: '   + (d.phone     || ''),
        'الإيميل: '    + (d.email     || ''),
        'الدفع: '      + (d.payMethod || '') + ' — ' + (d.finalPrice || 2500) + ' جنيه',
        'الكود: '      + (d.discountCode || 'لا يوجد'),
      ].join('\n')
    });
  } catch(_) {}

  return { success: true };
}

function updateStatus(rowIndex, status) {
  if (!rowIndex || !status) return { error: 'missing params' };
  getSheet().getRange(rowIndex, 9).setValue(status);
  return { success: true };
}
