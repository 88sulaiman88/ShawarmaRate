/**
 * شاورما ريت — استقبال الاقتراحات وحفظها في Google Sheet
 * انسخ هذا الملف كامل في محرر Apps Script داخل الجدول.
 */

// غيّرها لنفس كلمة مرور لوحة الأدمن في الموقع
const READ_KEY = 'shawarma';

const HEADERS = ['id', 'التاريخ', 'المطعم', 'المقترح', 'ملاحظات'];

function sheet_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
  if (sh.getLastRow() === 0) sh.appendRow(HEADERS);
  return sh;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/** إضافة اقتراح (من نموذج الموقع) أو حذف (من لوحة الأدمن) */
function doPost(e) {
  try {
    const body = JSON.parse(e.postData.contents);

    if (body.action === 'delete') {
      if (body.key !== READ_KEY) return json_({ error: 'unauthorized' });
      const sh = sheet_();
      const ids = sh.getRange(1, 1, sh.getLastRow(), 1).getValues();
      for (let i = ids.length - 1; i >= 1; i--) {
        if (String(ids[i][0]) === String(body.id)) { sh.deleteRow(i + 1); break; }
      }
      return json_({ ok: true });
    }

    const place = String(body.place || '').trim().slice(0, 80);
    const who   = String(body.who   || '').trim().slice(0, 40);
    const note  = String(body.note  || '').trim().slice(0, 400);
    if (!place || !who) return json_({ error: 'missing' });

    sheet_().appendRow([Date.now(), new Date(), place, who, note]);
    return json_({ ok: true });

  } catch (err) {
    return json_({ error: String(err) });
  }
}

/** قراءة الاقتراحات — تتطلب المفتاح، فلا يقدر أحد يقرأها إلا أنت */
function doGet(e) {
  if (!e.parameter || e.parameter.key !== READ_KEY) return json_({ error: 'unauthorized' });

  const sh = sheet_();
  if (sh.getLastRow() < 2) return json_({ items: [] });

  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues();
  const items = rows.map(function (r) {
    return { id: String(r[0]), t: Number(r[0]), place: r[2], who: r[3], note: r[4] };
  });
  return json_({ items: items });
}
