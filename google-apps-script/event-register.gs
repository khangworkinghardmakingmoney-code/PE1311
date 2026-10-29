/**
 * Nhận dữ liệu form đăng ký từ index.html và ghi vào Google Sheet.
 *
 * Cách cài (một lần):
 *  1. Mở Google Sheet sẽ chứa dữ liệu → Extensions (Tiện ích) → Apps Script.
 *  2. Dán toàn bộ file này vào Code.gs → Save.
 *  3. Deploy → New deployment → loại "Web app":
 *       Execute as: Me   ·   Who has access: Anyone
 *  4. Cấp quyền khi được hỏi, rồi copy "Web app URL" (dạng https://script.google.com/macros/s/.../exec)
 *     dán vào biến REGISTER_ENDPOINT trong index.html.
 *  Sửa code sau này: Deploy → Manage deployments → Edit → Version: New version (giữ nguyên URL).
 */

const SHEET_NAME = 'Đăng ký';

const HEADERS = [
  'Thời gian', 'Họ và tên', 'Email', 'Số điện thoại',
  'Công ty / Tổ chức', 'Vị trí', 'Hình thức', 'Đồng ý dữ liệu', 'Biết đến qua',
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = e.parameter;
    // Ô bẫy: người thật không thấy nên để trống, bot thì hay điền vào
    if (p.website) return json_({ ok: true });

    if (!p.name || !p.email || !p.phone || !p.company) {
      return json_({ ok: false, error: 'missing_fields' });
    }

    const sheet = getSheet_();
    sheet.appendRow([
      new Date(),
      p.name,
      p.email,
      // Thêm ' để Sheets không cắt số 0 đầu của số điện thoại
      "'" + p.phone,
      p.company,
      p.position || '',
      p.mode === 'online' ? 'Tham gia Online' : 'Tham gia Trực tiếp',
      p.consent ? 'Có' : 'Không',
      p.source || '',
    ]);
    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else {
    // Sheet tạo trước khi thêm cột mới: ghi lại hàng tiêu đề cho đủ cột
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  }
  sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
  return sheet;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
