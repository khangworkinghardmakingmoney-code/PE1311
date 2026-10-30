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
  'Công ty / Tổ chức', 'Vị trí', 'Hình thức', 'Đồng ý dữ liệu', 'Biết đến qua', 'Câu hỏi / Khó khăn',
];

const FREE_EMAIL_DOMAINS = [
  'gmail.com', 'googlemail.com', 'yahoo.com', 'yahoo.com.vn', 'ymail.com', 'rocketmail.com',
  'outlook.com', 'outlook.com.vn', 'hotmail.com', 'live.com', 'msn.com', 'icloud.com', 'me.com', 'mac.com',
  'aol.com', 'proton.me', 'protonmail.com', 'zoho.com', 'mail.com', 'gmx.com', 'yandex.com', 'qq.com',
];

function doPost(e) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const p = e.parameter;
    // Ô bẫy: người thật không thấy nên để trống, bot thì hay điền vào
    if (p.website) return json_({ ok: true });

    // Mọi trường trên form đều bắt buộc, trừ câu hỏi (question)
    const required = ['name', 'email', 'phone', 'company', 'position', 'source', 'mode', 'consent'];
    const missing = required.filter(function (k) { return !String(p[k] || '').trim(); });
    if (missing.length) {
      return json_({ ok: false, error: 'missing_fields', fields: missing });
    }

    // Chỉ nhận email doanh nghiệp (giữ giống danh sách FREE_EMAIL_DOMAINS trong index.html)
    const domain = String(p.email).trim().toLowerCase().split('@')[1] || '';
    if (FREE_EMAIL_DOMAINS.indexOf(domain) !== -1) {
      return json_({ ok: false, error: 'free_email' });
    }

    // Email phải đúng dạng ten@tenmien.xx
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(p.email).trim())) {
      return json_({ ok: false, error: 'invalid_email' });
    }

    const sheet = getSheet_();
    sheet.appendRow([
      new Date(),
      cell_(p.name, 100),
      cell_(p.email, 150),
      // Luôn thêm ' để Sheets không cắt số 0 đầu của số điện thoại (và không coi là công thức)
      "'" + clean_(p.phone, 20),
      cell_(p.company, 150),
      cell_(p.position, 100),
      p.mode === 'online' ? 'Tham gia Online' : 'Tham gia Trực tiếp',
      p.consent ? 'Có' : 'Không',
      cell_(p.source, 100),
      cell_(p.question, 1000),
    ]);
    return json_({ ok: true });
  } catch (err) {
    // Ghi lỗi vào mục Executions, không trả chi tiết lỗi ra ngoài
    console.error(err);
    return json_({ ok: false, error: 'server_error' });
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

// Bỏ khoảng trắng thừa, ký tự điều khiển và cắt theo độ dài tối đa
function clean_(value, maxLength) {
  return String(value || '')
    .replace(/[\u0000-\u001F\u007F]/g, ' ')
    .trim()
    .slice(0, maxLength);
}

// Chống chèn công thức: Sheets coi ô bắt đầu bằng = + - @ là công thức (vd =IMPORTXML(...)
// có thể gửi dữ liệu các dòng khác ra ngoài khi mở Sheet). Thêm ' để ô luôn là chữ thường.
function cell_(value, maxLength) {
  const text = clean_(value, maxLength);
  return /^[=+\-@]/.test(text) ? "'" + text : text;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
