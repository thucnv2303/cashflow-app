/**
 * Hướng dẫn cài đặt (Setup Instructions):
 * 1. Mở Google Sheets, tạo một Spreadsheet mới hoặc mở một Spreadsheet có sẵn.
 * 2. Trên menu, chọn Tiện ích mở rộng (Extensions) > Apps Script.
 * 3. Xóa code có sẵn và dán toàn bộ nội dung file này vào.
 * 4. Chạy hàm `setupSheet` (chọn setupSheet trên thanh công cụ rồi nhấn Run). Cấp quyền nếu được yêu cầu.
 * 5. Bấm Deploy > New deployment.
 *    - Select type: Web app
 *    - Execute as: Me (Tài khoản của bạn)
 *    - Who has access: Anyone (Bất kỳ ai)
 * 6. Bấm Deploy, sao chép Web app URL và dán vào Cài đặt trong ứng dụng CashFlow của bạn.
 */

function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Transactions');
  
  if (!sheet) {
    sheet = ss.insertSheet('Transactions');
    // Tạo headers
    sheet.appendRow(['ID', 'Loại', 'Số tiền', 'Danh mục', 'Ghi chú', 'Ngày', 'Ngày tạo']);
    // Cố định hàng đầu tiên
    sheet.setFrozenRows(1);
    // Làm đậm hàng đầu
    sheet.getRange('A1:G1').setFontWeight('bold');
  }
}

function getSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Transactions');
  if (!sheet) {
    setupSheet();
    sheet = ss.getSheetByName('Transactions');
  }
  return sheet;
}

function findRowById(sheet, id) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) { // Bỏ qua header
    if (data[i][0] === id) {
      return i + 1; // getRange bắt đầu từ 1
    }
  }
  return -1;
}

function rowToObject(row) {
  return {
    id: row[0],
    type: row[1],
    amount: Number(row[2]),
    category: row[3],
    note: row[4],
    date: row[5],
    createdAt: row[6]
  };
}

function objectToRow(obj) {
  return [
    obj.id || '',
    obj.type || '',
    obj.amount || 0,
    obj.category || '',
    obj.note || '',
    obj.date || '',
    obj.createdAt || new Date().toISOString()
  ];
}

function responseJson(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  // Lấy parameter 'action', mặc định là 'getAll'
  const action = e.parameter.action || 'getAll';
  
  if (action === 'ping') {
    return responseJson({ status: 'ok', message: 'Kết nối thành công!' });
  }
  
  const lock = LockService.getScriptLock();
  try {
    // Chờ tối đa 10 giây để lấy lock
    lock.waitLock(10000);
    
    const sheet = getSheet();
    
    if (action === 'getAll') {
      const data = sheet.getDataRange().getValues();
      const transactions = [];
      for (let i = 1; i < data.length; i++) {
        if (data[i][0]) { // Kiểm tra nếu ID không rỗng
          transactions.push(rowToObject(data[i]));
        }
      }
      return responseJson({ success: true, data: transactions });
    }
    
    else if (action === 'add') {
      const dataStr = e.parameter.data;
      if (!dataStr) throw new Error('Thiếu dữ liệu (data)');
      const obj = JSON.parse(dataStr);
      
      sheet.appendRow(objectToRow(obj));
      return responseJson({ success: true, message: 'Đã thêm giao dịch' });
    }
    
    else if (action === 'update') {
      const id = e.parameter.id;
      const dataStr = e.parameter.data;
      if (!id || !dataStr) throw new Error('Thiếu ID hoặc dữ liệu');
      const obj = JSON.parse(dataStr);
      
      const rowNum = findRowById(sheet, id);
      if (rowNum > -1) {
        // Ghi đè dòng hiện tại
        const newRow = objectToRow(obj);
        sheet.getRange(rowNum, 1, 1, newRow.length).setValues([newRow]);
        return responseJson({ success: true, message: 'Đã cập nhật giao dịch' });
      } else {
        throw new Error('Không tìm thấy giao dịch với ID này');
      }
    }
    
    else if (action === 'delete') {
      const id = e.parameter.id;
      if (!id) throw new Error('Thiếu ID');
      
      const rowNum = findRowById(sheet, id);
      if (rowNum > -1) {
        sheet.deleteRow(rowNum);
        return responseJson({ success: true, message: 'Đã xóa giao dịch' });
      } else {
        throw new Error('Không tìm thấy giao dịch với ID này');
      }
    }
    
    else if (action === 'sync') {
      const dataStr = e.parameter.data;
      if (!dataStr) throw new Error('Thiếu dữ liệu đồng bộ');
      const transactions = JSON.parse(dataStr);
      
      // Xóa tất cả dữ liệu cũ (giữ lại header)
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      }
      
      // Thêm tất cả dữ liệu mới
      if (transactions && transactions.length > 0) {
        const rows = transactions.map(t => objectToRow(t));
        sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
      }
      return responseJson({ success: true, message: 'Đã đồng bộ toàn bộ dữ liệu' });
    }
    
    else {
      return responseJson({ success: false, error: 'Hành động không hợp lệ' });
    }
    
  } catch (error) {
    return responseJson({ success: false, error: error.toString() });
  } finally {
    lock.releaseLock();
  }
}
