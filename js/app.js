// ==================== APPS SCRIPT CODE (for copy button) ====================
const APPS_SCRIPT_CODE = `/**
 * CashFlow - Google Apps Script API
 * Dán code này vào Apps Script của Google Sheet.
 * Chạy hàm setupSheet() trước khi deploy.
 */
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Transactions');
  if (!sheet) {
    sheet = ss.insertSheet('Transactions');
    sheet.appendRow(['ID','Loại','Số tiền','Danh mục','Ghi chú','Ngày','Thành viên','Ngày tạo','Chi cho ai']);
    sheet.setFrozenRows(1);
    sheet.getRange('A1:I1').setFontWeight('bold');
  }
  let loanSheet = ss.getSheetByName('Loans');
  if (!loanSheet) {
    loanSheet = ss.insertSheet('Loans');
    loanSheet.appendRow(['ID','Tên','Emoji','Loại','Gốc vay','Lãi suất','Kỳ hạn','Trả/tháng','Ngày BĐ','Ghi chú','Ngày tạo']);
    loanSheet.setFrozenRows(1);
    loanSheet.getRange('A1:K1').setFontWeight('bold');
  }
  let memberSheet = ss.getSheetByName('Members');
  if (!memberSheet) {
    memberSheet = ss.insertSheet('Members');
    memberSheet.appendRow(['ID','Tên','Avatar','Màu','AvatarImg','AvatarId']);
    memberSheet.setFrozenRows(1);
    memberSheet.getRange('A1:F1').setFontWeight('bold');
  }
  let budgetSheet = ss.getSheetByName('Budgets');
  if (!budgetSheet) {
    budgetSheet = ss.insertSheet('Budgets');
    budgetSheet.appendRow(['CategoryID','Số tiền','Cập nhật']);
    budgetSheet.setFrozenRows(1);
    budgetSheet.getRange('A1:C1').setFontWeight('bold');
  }
  let catSheet = ss.getSheetByName('CustomCategories');
  if (!catSheet) {
    catSheet = ss.insertSheet('CustomCategories');
    catSheet.appendRow(['ID','Tên','Emoji','Loại','Ngày tạo']);
    catSheet.setFrozenRows(1);
    catSheet.getRange('A1:E1').setFontWeight('bold');
  }
  let savingsSheet = ss.getSheetByName('SavingsGoals');
  if (!savingsSheet) {
    savingsSheet = ss.insertSheet('SavingsGoals');
    savingsSheet.appendRow(['ID','Tên','Emoji','Mục tiêu','Hiện có','Hạn ngày','Thành viên','Ngày tạo']);
    savingsSheet.setFrozenRows(1);
    savingsSheet.getRange('A1:H1').setFontWeight('bold');
  let logsSheet = ss.getSheetByName('SavingsLogs');
  if (!logsSheet) {
    logsSheet = ss.insertSheet('SavingsLogs');
    logsSheet.appendRow(['ID','GoalID','GoalName','Loại','Số tiền','Thành viên','Ngày','Ghi chú','Ngày tạo']);
    logsSheet.setFrozenRows(1);
    logsSheet.getRange('A1:I1').setFontWeight('bold');
  }
  let pendingSheet = ss.getSheetByName('PendingTransactions');
  if (!pendingSheet) {
    pendingSheet = ss.insertSheet('PendingTransactions');
    pendingSheet.appendRow(['ID','Ngân hàng','Loại','Số tiền','Nội dung thô','Danh mục gợi ý','Thành viên','Ngày giờ','Trạng thái','Ngày tạo']);
    pendingSheet.setFrozenRows(1);
    pendingSheet.getRange('A1:J1').setFontWeight('bold');
  }
}
function getSheet(name) {
  return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(name);
}
function findRow(sheet,id){var d=sheet.getDataRange().getValues();for(var i=1;i<d.length;i++){if(d[i][0]===id)return i+1;}return -1;}
function responseJson(data, callback){
  var json = JSON.stringify(data);
  if (callback) {
    return ContentService.createTextOutput(callback + '(' + json + ')').setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}
function doPost(e) {
  if (e && e.postData && e.postData.contents) {
    try {
      var body = JSON.parse(e.postData.contents);
      e.parameter = e.parameter || {};
      for (var k in body) { e.parameter[k] = body[k]; }
    } catch(err) {}
  }
  return doGet(e);
}
function classifyBankNotification(rawText, title, bank) {
  var fullText = (title + ' ' + rawText + ' ' + (bank || '')).toLowerCase();
  var amount = 0;
  var amtMatch = rawText.match(/(?:[\+\-]|gd:?\s*[\+\-]?|ps:?\s*[\+\-]?)?\s*([\d\.\,]{3,15})\s*(?:vnd|vnđ|đ|d\b)/i) || rawText.match(/([\d\.\,]{4,15})\s*(?:vnd|vnđ|đ|d\b)/i);
  if (amtMatch && amtMatch[1]) { amount = Number(amtMatch[1].replace(/[\.\,]/g, '')) || 0; }
  var type = 'expense';
  if (rawText.includes('+') || fullText.includes('nhan tien') || fullText.includes('nhận tiền') || fullText.includes('cong tien') || fullText.includes('cộng tiền') || fullText.includes('credit')) {
    type = 'income';
  } else if (rawText.includes('-') || fullText.includes('tru tien') || fullText.includes('trừ tiền') || fullText.includes('thanh toan') || fullText.includes('thanh toán') || fullText.includes('debit')) {
    type = 'expense';
  }
  var category = type === 'income' ? 'salary' : 'other_expense';
  if (type === 'expense') {
    if (/highlands|phở|bún|cơm|quán|cafe|cà phê|trà sữa|starbucks|kfc|lotteria|pizza|food|shopeefood|grabfood|befood|baemin|tokyo deli|haidilao|gogi|kichi|nhà hàng|bánh mì|ăn sáng|ăn trưa|ăn tối|lẩu|nướng/i.test(fullText)) { category = 'food'; }
    else if (/grab|be |xanh sm|taxi|xăng|petrolimex|pvoil|gửi xe|giữ xe|vé xe|cầu đường|epass|vetc|đỗ xe/i.test(fullText)) { category = 'transport'; }
    else if (/shopee|lazada|tiki|sendo|winmart|co\.?op|bách hóa|bach hoa|siêu thị|uniqlo|zara|h&m|mua sắm|shop|store|mall|quần áo|giày|mỹ phẩm/i.test(fullText)) { category = 'shopping'; }
    else if (/evn|điện lực|tiền điện|tiền nước|cấp nước|viettel|vnpt|fpt|internet|wifi|mobifone|vinaphone|chung cư|phí quản lý|vệ sinh|rác/i.test(fullText)) { category = 'bills'; }
    else if (/thuốc|pharmacity|long châu|an khang|bệnh viện|phòng khám|bác sĩ|nha khoa|răng|khám bệnh|y tế/i.test(fullText)) { category = 'health'; }
    else if (/học phí|trường|mầm non|tiểu học|tiếng anh|ila|vus|apollo|sách|vở|dụng cụ học tập/i.test(fullText)) { category = 'education'; }
    else if (/cgv|bhd|lotte cinema|rạp|vé xem phim|netflix|spotify|youtube|steam|game|playstation|du lịch|khách sạn|resort|vé máy bay/i.test(fullText)) { category = 'entertainment'; }
    else if (/nội thất|điện máy|điện thoại|laptop|sửa nhà|decor|gia dụng/i.test(fullText)) { category = 'house'; }
    else if (/chứng khoán|cổ phiếu|ssi|vps|tcbs|vndirect|tiết kiệm|gửi tiền|vàng|sjc|doji/i.test(fullText)) { category = 'invest'; }
  } else {
    if (/lương|salary|payroll|thu nhập|thưởng|bonus/i.test(fullText)) { category = 'salary'; }
    else if (/đầu tư|cổ tức|lãi|tiết kiệm|interest/i.test(fullText)) { category = 'investment'; }
    else { category = 'other_income'; }
  }
  var cleanNote = rawText.replace(/\r?\n|\r/g, ' ').slice(0, 100);
  return { amount: amount, type: type, category: category, note: cleanNote };
}
function doGet(e) {
  e = e || {};
  e.parameter = e.parameter || {};
  var action = e.parameter.action || 'getAll';
  var cb = e.parameter.callback || null;
  if (action==='ping') return responseJson({status:'ok',message:'Kết nối thành công!'}, cb);
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (action==='getAll') {
      var sheet=getSheet('Transactions');if(!sheet)return responseJson({success:true,data:[]}, cb);
      var d=sheet.getDataRange().getValues(),r=[];
      for(var i=1;i<d.length;i++){if(d[i][0])r.push({id:String(d[i][0]),type:String(d[i][1]||'expense'),amount:Number(d[i][2]||0),category:String(d[i][3]||'other'),note:String(d[i][4]||''),date:String(d[i][5]||''),memberId:String(d[i][6]||''),createdAt:String(d[i][7]||''),beneficiaryId:String(d[i][8]||'')});}
      return responseJson({success:true,data:r}, cb);
    }
    if (action==='add'){var o=JSON.parse(e.parameter.data);getSheet('Transactions').appendRow([o.id||'',o.type||'',o.amount||0,o.category||'',o.note||'',o.date||'',o.memberId||'',o.createdAt||'',o.beneficiaryId||'']);return responseJson({success:true}, cb);}
    if (action==='update'){var id=e.parameter.id,o=JSON.parse(e.parameter.data),s=getSheet('Transactions'),r=findRow(s,id);if(r>-1){s.getRange(r,1,1,9).setValues([[o.id||id,o.type||'',o.amount||0,o.category||'',o.note||'',o.date||'',o.memberId||'',o.createdAt||'',o.beneficiaryId||'']]);return responseJson({success:true}, cb);}throw new Error('Not found');}
    if (action==='delete'){var s=getSheet('Transactions'),r=findRow(s,e.parameter.id);if(r>-1){s.deleteRow(r);return responseJson({success:true}, cb);}throw new Error('Not found');}
    if (action==='sync'){var ts=JSON.parse(e.parameter.data),s=getSheet('Transactions'),lr=s.getLastRow();if(lr>1)s.getRange(2,1,lr-1,s.getLastColumn()).clearContent();if(ts&&ts.length>0){var rows=ts.map(function(t){return[t.id||'',t.type||'',t.amount||0,t.category||'',t.note||'',t.date||'',t.memberId||'',t.createdAt||'',t.beneficiaryId||''];});s.getRange(2,1,rows.length,9).setValues(rows);}return responseJson({success:true}, cb);}
    if (action==='getLoans'){var s=getSheet('Loans');if(!s)return responseJson({success:true,data:[]}, cb);var d=s.getDataRange().getValues(),r=[];for(var i=1;i<d.length;i++){if(d[i][0])r.push({id:d[i][0],name:d[i][1],emoji:d[i][2],loanType:d[i][3],principal:Number(d[i][4]),interestRate:Number(d[i][5]),termMonths:Number(d[i][6]),monthlyPayment:Number(d[i][7]),startDate:d[i][8],note:d[i][9],createdAt:d[i][10]});}return responseJson({success:true,data:r}, cb);}
    if (action==='syncLoans'){var ls=JSON.parse(e.parameter.data),s=getSheet('Loans'),lr=s.getLastRow();if(lr>1)s.getRange(2,1,lr-1,s.getLastColumn()).clearContent();if(ls&&ls.length>0){var rows=ls.map(function(l){return[l.id||'',l.name||'',l.emoji||'',l.loanType||'',l.principal||0,l.interestRate||0,l.termMonths||0,l.monthlyPayment||0,l.startDate||'',l.note||'',l.createdAt||''];});s.getRange(2,1,rows.length,11).setValues(rows);}return responseJson({success:true}, cb);}
    if (action==='getMembers'){var s=getSheet('Members');if(!s)return responseJson({success:true,data:[]}, cb);var d=s.getDataRange().getValues(),r=[];for(var i=1;i<d.length;i++){if(d[i][0])r.push({id:String(d[i][0]),name:String(d[i][1]||''),avatar:String(d[i][2]||'👤'),color:String(d[i][3]||'#e77d3e'),avatarImg:String(d[i][4]||''),avatarId:String(d[i][5]||'')});}return responseJson({success:true,data:r}, cb);}
    if (action==='syncMembers'){var ms=JSON.parse(e.parameter.data),s=getSheet('Members'),lr=s.getLastRow();if(lr>1)s.getRange(2,1,lr-1,s.getLastColumn()).clearContent();if(ms&&ms.length>0){var rows=ms.map(function(m){return[m.id||'',m.name||'',m.avatar||'',m.color||'',m.avatarImg||'',m.avatarId||''];});s.getRange(2,1,rows.length,6).setValues(rows);}return responseJson({success:true}, cb);}
    if (action==='getBudgets'){var s=getSheet('Budgets');if(!s)return responseJson({success:true,data:{}}, cb);var d=s.getDataRange().getValues(),b={};for(var i=1;i<d.length;i++){if(d[i][0])b[String(d[i][0])]=Number(d[i][1]||0);}return responseJson({success:true,data:b}, cb);}
    if (action==='syncBudgets'){var bs=JSON.parse(e.parameter.data),s=getSheet('Budgets'),lr=s.getLastRow();if(lr>1)s.getRange(2,1,lr-1,s.getLastColumn()).clearContent();if(bs&&typeof bs==='object'){var rows=Object.keys(bs).map(function(k){return[k,Number(bs[k]||0),new Date().toISOString()];});if(rows.length>0)s.getRange(2,1,rows.length,3).setValues(rows);}return responseJson({success:true,message:'Đã đồng bộ ngân sách'}, cb);}
    if (action==='getCustomCats'){var s=getSheet('CustomCategories');if(!s)return responseJson({success:true,data:[]}, cb);var d=s.getDataRange().getValues(),r=[];for(var i=1;i<d.length;i++){if(d[i][0])r.push({id:String(d[i][0]),label:String(d[i][1]||''),emoji:String(d[i][2]||'📦'),type:String(d[i][3]||'expense'),createdAt:String(d[i][4]||'')});}return responseJson({success:true,data:r}, cb);}
    if (action==='syncCustomCats'){var cs=JSON.parse(e.parameter.data),s=getSheet('CustomCategories'),lr=s.getLastRow();if(lr>1)s.getRange(2,1,lr-1,s.getLastColumn()).clearContent();if(cs&&cs.length>0){var rows=cs.map(function(c){return[c.id||'',c.label||'',c.emoji||'📦',c.type||'expense',c.createdAt||''];});s.getRange(2,1,rows.length,5).setValues(rows);}return responseJson({success:true,message:'Đã đồng bộ danh mục tùy chỉnh'}, cb);}
    if (action==='getSavings'){var gs=getSheet('SavingsGoals'),ls=getSheet('SavingsLogs'),goals=[],logs=[];if(gs){var d=gs.getDataRange().getValues();for(var i=1;i<d.length;i++){if(d[i][0])goals.push({id:String(d[i][0]),name:String(d[i][1]||''),emoji:String(d[i][2]||'🐷'),targetAmount:Number(d[i][3]||0),currentAmount:Number(d[i][4]||0),targetDate:String(d[i][5]||''),memberId:String(d[i][6]||'family'),createdAt:String(d[i][7]||'')});}}if(ls){var d=ls.getDataRange().getValues();for(var i=1;i<d.length;i++){if(d[i][0])logs.push({id:String(d[i][0]),goalId:String(d[i][1]||''),goalName:String(d[i][2]||''),type:String(d[i][3]||'deposit'),amount:Number(d[i][4]||0),memberId:String(d[i][5]||'family'),date:String(d[i][6]||''),note:String(d[i][7]||''),createdAt:String(d[i][8]||'')});}}return responseJson({success:true,data:{goals:goals,logs:logs}}, cb);}
    if (action==='syncSavings'){var goals=JSON.parse(e.parameter.goals||'[]'),logs=JSON.parse(e.parameter.logs||'[]');var gs=getSheet('SavingsGoals');if(gs){var lr=gs.getLastRow();if(lr>1)gs.getRange(2,1,lr-1,gs.getLastColumn()).clearContent();if(goals&&goals.length>0){var rows=goals.map(function(g){return[g.id||'',g.name||'',g.emoji||'🐷',Number(g.targetAmount||0),Number(g.currentAmount||0),g.targetDate||'',g.memberId||'family',g.createdAt||''];});gs.getRange(2,1,rows.length,8).setValues(rows);}}var ls=getSheet('SavingsLogs');if(ls){var lr=ls.getLastRow();if(lr>1)ls.getRange(2,1,lr-1,ls.getLastColumn()).clearContent();if(logs&&logs.length>0){var rows=logs.map(function(l){return[l.id||'',l.goalId||'',l.goalName||'',l.type||'deposit',Number(l.amount||0),l.memberId||'family',l.date||'',l.note||'',l.createdAt||''];});ls.getRange(2,1,rows.length,9).setValues(rows);}}return responseJson({success:true,message:'Đã đồng bộ tiết kiệm'}, cb);}
    if (action==='bankNotification'){var rawText=String(e.parameter.text||e.parameter.body||e.parameter.content||e.parameter.message||''),title=String(e.parameter.title||e.parameter.sender||''),bank=String(e.parameter.bank||title||'Ngân hàng'),member=String(e.parameter.member||'dad');if(!rawText&&!title)return responseJson({success:false,error:'Thiếu nội dung'}, cb);var cls=classifyBankNotification(rawText,title,bank),pSheet=getSheet('PendingTransactions');if(!pSheet)return responseJson({success:false,error:'Sheet not found'}, cb);var id='pend_'+new Date().getTime()+'_'+Math.floor(Math.random()*1000),nowStr=new Date().toISOString();pSheet.appendRow([id,bank,cls.type,cls.amount,cls.note,cls.category,member,nowStr,'pending',nowStr]);return responseJson({success:true,message:'Đã nhận thông báo',transaction:{id:id,bank:bank,type:cls.type,amount:cls.amount,note:cls.note,category:cls.category,memberId:member,date:nowStr,status:'pending'}}, cb);}
    if (action==='getPending'){var sheet=getSheet('PendingTransactions');if(!sheet)return responseJson({success:true,data:[]}, cb);var d=sheet.getDataRange().getValues(),p=[];for(var i=1;i<d.length;i++){if(d[i][0]&&String(d[i][8])==='pending')p.push({id:String(d[i][0]),bank:String(d[i][1]||'Ngân hàng'),type:String(d[i][2]||'expense'),amount:Number(d[i][3]||0),note:String(d[i][4]||''),category:String(d[i][5]||'other_expense'),memberId:String(d[i][6]||'family'),date:String(d[i][7]||''),status:String(d[i][8]||'pending'),createdAt:String(d[i][9]||'')});}return responseJson({success:true,data:p}, cb);}
    if (action==='approvePending'){var pId=String(e.parameter.pendingId||e.parameter.id),data=typeof e.parameter.data==='string'?JSON.parse(e.parameter.data):e.parameter.data,tSheet=getSheet('Transactions');if(tSheet&&data){tSheet.appendRow([data.id||('trans_'+new Date().getTime()),data.type||'expense',Number(data.amount||0),data.category||'other_expense',data.note||'',data.date||new Date().toISOString().slice(0,10),data.memberId||'dad',new Date().toISOString(),data.beneficiaryId||data.memberId||'family']);}var pSheet=getSheet('PendingTransactions');if(pSheet&&pId){var r=findRow(pSheet,pId);if(r>-1)pSheet.getRange(r,9).setValue('approved');}return responseJson({success:true,message:'Đã duyệt'}, cb);}
    if (action==='deletePending'){var pId=String(e.parameter.id||e.parameter.pendingId),pSheet=getSheet('PendingTransactions');if(pSheet&&pId){var r=findRow(pSheet,pId);if(r>-1)pSheet.getRange(r,9).setValue('rejected');}return responseJson({success:true,message:'Đã bỏ qua'}, cb);}
    if (action==='scanGmail'){var res=scanBankEmails();return responseJson({success:true,addedCount:res.addedCount,message:'Đã quét Gmail thành công! Đã thêm '+res.addedCount+' giao dịch.'}, cb);}
    return responseJson({success:false,error:'Unknown action'}, cb);
  } catch(err) {return responseJson({success:false,error:err.toString()}, cb);} finally {lock.releaseLock();}
}
function scanBankEmails(targetMember){var member=targetMember||'dad',ss=SpreadsheetApp.getActiveSpreadsheet(),pSheet=ss.getSheetByName('PendingTransactions');if(!pSheet){setupSheet();pSheet=ss.getSheetByName('PendingTransactions');}var label=GmailApp.getUserLabelByName('CashFlow_Processed');if(!label){label=GmailApp.createLabel('CashFlow_Processed');}var queries=['from:(vpbank.com.vn OR techcombank.com.vn OR mbbank.com.vn OR vietcombank.com.vn OR timo.vn OR acb.com.vn OR vib.com.vn OR tpbank.com.vn OR bidv.com.vn OR cake.vn) -label:CashFlow_Processed newer_than:3d','subject:("Biến động số dư" OR "Thông báo giao dịch" OR "Transaction Alert" OR "Thông báo thay đổi số dư") -label:CashFlow_Processed newer_than:3d'],addedCount=0,processedThreadIds={};for(var q=0;q<queries.length;q++){var threads=GmailApp.search(queries[q],0,15);for(var i=0;i<threads.length;i++){var thread=threads[i];if(processedThreadIds[thread.getId()])continue;processedThreadIds[thread.getId()]=true;var messages=thread.getMessages();for(var j=0;j<messages.length;j++){var msg=messages[j],subject=msg.getSubject()||'',sender=msg.getFrom()||'',bodyText=msg.getPlainBody()||'',date=msg.getDate(),bank='Ngân hàng';if(/vpbank/i.test(sender)||/vpbank/i.test(subject)||/vpbank/i.test(bodyText))bank='VPBank';else if(/techcombank/i.test(sender)||/techcombank/i.test(subject))bank='Techcombank';else if(/mbbank|mb bank/i.test(sender)||/mbbank|mb bank/i.test(subject))bank='MB Bank';else if(/vietcombank|vcb/i.test(sender)||/vietcombank/i.test(subject))bank='Vietcombank';else if(/timo/i.test(sender)||/timo/i.test(subject))bank='Timo';else if(/acb/i.test(sender)||/acb/i.test(subject))bank='ACB';else if(/vib/i.test(sender)||/vib/i.test(subject))bank='VIB';else if(/tpbank/i.test(sender)||/tpbank/i.test(subject))bank='TPBank';else if(/bidv/i.test(sender)||/bidv/i.test(subject))bank='BIDV';else if(/cake/i.test(sender)||/cake/i.test(subject))bank='CAKE';else if(/momo/i.test(sender)||/momo/i.test(subject))bank='MoMo';var cls=classifyBankNotification(bodyText,subject,bank);if(cls.amount>0){var id='email_'+new Date().getTime()+'_'+Math.floor(Math.random()*1000),dateStr=date?date.toISOString():new Date().toISOString();pSheet.appendRow([id,bank,cls.type,cls.amount,cls.note||subject,cls.category,member,dateStr,'pending',dateStr]);addedCount++;}}thread.addLabel(label);}}return{success:true,addedCount:addedCount};}
function scanMemberBankEmails(memberId){return scanBankEmails(memberId||'mom');}
function triggerScanMom(){return scanBankEmails('mom');}
function triggerScanDad(){return scanBankEmails('dad');}
function triggerScanChild(){return scanBankEmails('child');}
function triggerScanOther(){return scanBankEmails('other');}
function setupMemberTrigger(memberId){var funcName='triggerScan'+memberId.charAt(0).toUpperCase()+memberId.slice(1);var triggers=ScriptApp.getProjectTriggers();for(var i=0;i<triggers.length;i++){if(triggers[i].getHandlerFunction()===funcName){ScriptApp.deleteTrigger(triggers[i]);}}ScriptApp.newTrigger(funcName).timeBased().everyMinutes(5).create();return'OK';}
function setupGmailTrigger(){var triggers=ScriptApp.getProjectTriggers();for(var i=0;i<triggers.length;i++){if(triggers[i].getHandlerFunction()==='scanBankEmails'){ScriptApp.deleteTrigger(triggers[i]);}}ScriptApp.newTrigger('scanBankEmails').timeBased().everyMinutes(1).create();return'OK';}
function renderMemberPortalHtml(memberId){var label=memberId==='mom'?'Vợ 🌸':(memberId==='dad'?'Chồng 👔':(memberId==='child'?'Con 🧒':'Thành viên'));var html='<!DOCTYPE html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>CashFlow - Kích hoạt</title><style>body{font-family:-apple-system,BlinkMacSystemFont,sans-serif;background:#0f172a;color:#f8fafc;display:flex;align-items:center;justify-content:center;min-height:100vh;margin:0;padding:16px;box-sizing:border-box}.card{background:#1e293b;border:1px solid #334155;border-radius:20px;padding:28px;max-width:440px;width:100%;text-align:center;box-shadow:0 10px 30px rgba(0,0,0,0.5)}.emoji{font-size:3rem;margin-bottom:12px}h1{font-size:1.3rem;margin:0 0 8px;color:#f8fafc}p{font-size:0.85rem;color:#94a3b8;line-height:1.5;margin:0 0 20px}.btn{background:#e77d3e;color:#fff;border:none;border-radius:12px;padding:14px 24px;font-size:1rem;font-weight:700;cursor:pointer;width:100%;box-shadow:0 4px 14px rgba(231,125,62,0.4)}.btn:disabled{opacity:0.6;cursor:not-allowed}.status{margin-top:18px;padding:14px;border-radius:10px;font-size:0.85rem;display:none;line-height:1.5}.status.success{background:rgba(34,197,94,0.15);color:#4ade80;border:1px solid rgba(34,197,94,0.3)}.status.loading{background:rgba(59,130,246,0.15);color:#60a5fa;border:1px solid rgba(59,130,246,0.3)}.status.error{background:rgba(248,113,113,0.15);color:#f87171;border:1px solid rgba(248,113,113,0.3)}.steps{text-align:left;margin-top:14px;font-size:0.8rem;color:#94a3b8;line-height:1.6}.steps .done{color:#4ade80}.steps .active{color:#60a5fa;font-weight:600}</style></head><body><div class="card"><div class="emoji">🌸</div><h1>Kích hoạt Quét Email Ngân Hàng</h1><p>Xin chào <strong>'+label+'</strong>!<br>Bấm nút bên dưới để cho phép hệ thống đọc email biến động ngân hàng và đưa vào sổ chi tiêu gia đình.<br><em style="font-size:0.78rem">Chỉ đọc email ngân hàng, không đọc email cá nhân.</em></p><button class="btn" id="scanBtn" onclick="go()">🚀 Kích hoạt & Quét Email Ngân Hàng</button><div id="statusBox" class="status"></div><div id="stepsBox" class="steps" style="display:none"><div id="s1">⬜ Bước 1: Cấp quyền đọc Gmail...</div><div id="s2">⬜ Bước 2: Quét email ngân hàng...</div><div id="s3">⬜ Bước 3: Thiết lập quét tự động 24/7...</div></div></div><script>function go(){var b=document.getElementById("scanBtn"),x=document.getElementById("statusBox"),st=document.getElementById("stepsBox");b.disabled=true;b.textContent="Đang kích hoạt... ⏳";st.style.display="block";document.getElementById("s1").innerHTML="🔄 Bước 1: Đang xin cấp quyền...";document.getElementById("s1").className="active";x.style.display="block";x.className="status loading";x.textContent="Đang kết nối Gmail...";google.script.run.withSuccessHandler(function(r){document.getElementById("s1").innerHTML="✅ Bước 1: Đã cấp quyền";document.getElementById("s1").className="done";document.getElementById("s2").innerHTML="✅ Bước 2: Tìm thấy "+(r.addedCount||0)+" giao dịch mới";document.getElementById("s2").className="done";document.getElementById("s3").innerHTML="🔄 Bước 3: Đang thiết lập tự động...";document.getElementById("s3").className="active";x.textContent="Đang thiết lập quét tự động 24/7...";google.script.run.withSuccessHandler(function(){document.getElementById("s3").innerHTML="✅ Bước 3: Đã thiết lập quét tự động!";document.getElementById("s3").className="done";b.disabled=false;b.textContent="🔄 Quét lại";x.className="status success";x.innerHTML="🎉 <strong>Hoàn tất!</strong><br>Đã thêm "+(r.addedCount||0)+" giao dịch.<br>Hệ thống sẽ tự quét mỗi 5 phút.<br><br><strong>Bạn có thể đóng trang này.</strong>";}).withFailureHandler(function(){document.getElementById("s3").innerHTML="⚠️ Bước 3: Chưa tự động hóa được";b.disabled=false;b.textContent="🔄 Quét lại";x.className="status success";x.innerHTML="🎉 Đã quét "+(r.addedCount||0)+" giao dịch!<br>Bạn có thể bấm nút để quét thủ công.";}).setupMemberTrigger("'+memberId+'");}).withFailureHandler(function(e){document.getElementById("s1").innerHTML="✅ Bước 1: OK";document.getElementById("s1").className="done";document.getElementById("s2").innerHTML="❌ Lỗi";b.disabled=false;b.textContent="🔄 Thử lại";x.className="status error";x.textContent="Lỗi: "+e;}).scanMemberBankEmails("'+memberId+'");}<\/script></body></html>';return HtmlService.createHtmlOutput(html).setTitle('CashFlow - Kích hoạt').setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);}`;

// ==================== MAIN APP ====================
const App = {
  currentPage: 'dashboard',
  currentMonth: getCurrentMonth(),
  budgetMonth: getCurrentMonth(),
  analyticsMonth: getCurrentMonth(),
  dashboardScope: 'all',
  budgetScope: 'all',
  savingsScope: 'all',
  editingId: null,
  editingLoanId: null,
  editingSavingsGoalId: null,
  selectedType: 'expense',
  selectedCategory: null,
  selectedMemberId: null,
  selectedBeneficiaryId: 'family',
  selectedLoanType: null,

  // ==================== INIT ====================
  init() {
    this.initTheme();
    this.bindEvents();
    this.initFilterMonth();
    this.initSyncStatus();
    // Populate hidden script textarea
    const codeEl = document.getElementById('appsScriptCode');
    if (codeEl) codeEl.value = APPS_SCRIPT_CODE;
    // Check if setup is done
    if (!Storage.isSetupDone()) {
      this.showSetupModal();
    }
    this.renderCurrentPage();
    // Init notifications
    if ('Notification' in window && Notification.permission === 'granted') {
      this.registerServiceWorker();
      this.scheduleReminderCheck();
    }
  },

  // ==================== AVATAR SELECTION ====================
  showSetupModal() {
    this.setupMemberRows = [
      { name: 'Ba Dâu', avatarId: 'dad', avatarImg: AVATARS[0].img },
      { name: 'Mẹ Dâu', avatarId: 'mom', avatarImg: AVATARS[1].img }
    ];
    this.renderSetupRows();
    document.getElementById('setupModal')?.classList.add('active');
  },

  renderSetupRows() {
    const container = document.getElementById('setupMembersList');
    if (!container) return;
    container.innerHTML = this.setupMemberRows.map((row, i) => `
      <div class="setup-member-row" data-index="${i}">
        <button type="button" class="avatar-select-btn" data-index="${i}">
          <img src="${row.avatarImg}" alt="Avatar" class="avatar-img">
        </button>
        <input type="text" class="setup-member-name" value="${row.name}" placeholder="Tên thành viên" required>
        ${this.setupMemberRows.length > 1 ? `<button type="button" class="btn-remove-member" data-index="${i}">✕</button>` : ''}
      </div>
    `).join('');

    container.querySelectorAll('.avatar-select-btn').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const idx = parseInt(btn.dataset.index);
        this.showAvatarPicker(idx, btn);
      });
    });

    container.querySelectorAll('.setup-member-name').forEach((input, idx) => {
      input.addEventListener('input', (e) => {
        this.setupMemberRows[idx].name = e.target.value;
      });
    });

    container.querySelectorAll('.btn-remove-member').forEach(btn => {
      btn.addEventListener('click', () => {
        const idx = parseInt(btn.dataset.index);
        this.setupMemberRows.splice(idx, 1);
        this.renderSetupRows();
      });
    });
  },

  handleSetupSubmit(e) {
    e.preventDefault();
    const familyName = document.getElementById('setupFamilyName').value.trim() || 'Gia đình';
    const members = [];
    this.setupMemberRows.forEach((row, i) => {
      if (row.name.trim()) {
        const av = AVATARS.find(a => a.id === row.avatarId) || AVATARS[0];
        members.push({
          id: generateId(),
          name: row.name.trim(),
          avatarId: row.avatarId,
          avatar: av.emoji,
          avatarImg: av.img,
          color: MEMBER_COLORS[i % MEMBER_COLORS.length]
        });
      }
    });

    if (members.length === 0) { this.showToast('Thêm ít nhất 1 thành viên', 'error'); return; }

    Storage.setFamilyName(familyName);
    members.forEach(m => Storage.addMember(m));
    Storage.setSetupDone();

    const modal = document.getElementById('setupModal');
    if (modal) modal.classList.remove('active');
    this.showToast('Chào mừng ' + familyName + '! 🎉');
    this.renderCurrentPage();
  },

  showAvatarPicker(index, button) {
    const currentRow = this.setupMemberRows[index];
    const currentAv = AVATARS.find(a => a.id === currentRow.avatarId) || AVATARS[0];
    const currentImg = currentRow.avatarImg || currentAv.img;

    this.openAvatarModal({
      currentImg: currentImg,
      onSelect: (selected) => {
        this.setupMemberRows[index].avatarId = selected.id || 'custom';
        this.setupMemberRows[index].avatarImg = selected.img;
        this.setupMemberRows[index].avatar = selected.emoji || '👤';
        button.innerHTML = `<img src="${selected.img}" alt="Avatar" class="avatar-img">`;
      }
    });
  },

  // ==================== THEME ====================
  initTheme() {
    const theme = Storage.getTheme();
    document.documentElement.setAttribute('data-theme', theme);
    this.updateThemeToggleIcon(theme);
    Charts.setTheme(theme === 'dark');
  },
  updateThemeToggleIcon(theme) {
    const icon = theme === 'dark' ? '🌙' : '☀️';
    document.querySelectorAll('.theme-icon').forEach(el => { el.textContent = icon; });
  },
  toggleTheme() {
    const newTheme = Storage.getTheme() === 'light' ? 'dark' : 'light';
    Storage.setTheme(newTheme);
    document.documentElement.setAttribute('data-theme', newTheme);
    this.updateThemeToggleIcon(newTheme);
    Charts.setTheme(newTheme === 'dark');
    this.renderCurrentPage();
  },

  // ==================== NAVIGATION ====================
  navigate(page) {
    this.currentPage = page;
    document.querySelectorAll('.nav-item').forEach(el => el.classList.toggle('active', el.dataset.page === page));
    document.querySelectorAll('.page').forEach(el => el.classList.remove('active'));
    const pageEl = document.getElementById(`page-${page}`);
    if (pageEl) pageEl.classList.add('active');
    this.renderCurrentPage();
  },
  renderCurrentPage() {
    switch(this.currentPage) {
      case 'dashboard': this.renderDashboard(); break;
      case 'budget': this.renderBudgetPage(); break;
      case 'savings': this.renderSavingsPage(); break;
      case 'transactions': this.renderTransactions(); break;
      case 'loans': this.renderLoans(); break;
      case 'analytics': this.renderAnalytics(); break;
      case 'settings': this.renderSettings(); break;
    }
  },

  // ==================== DYNAMIC MEMBER SCOPE SELECTOR ====================
  renderScopeSelector(containerId, activeScope, onChange) {
    const container = document.getElementById(containerId);
    if (!container) return;
    const members = Storage.getMembers();
    
    let html = `
      <button type="button" class="scope-btn ${activeScope === 'all' ? 'active' : ''}" data-scope="all">Tất cả</button>
      <button type="button" class="scope-btn ${activeScope === 'family' ? 'active' : ''}" data-scope="family">👨‍👩‍👧‍👦 Cả nhà</button>
    `;

    members.forEach(m => {
      const imgSrc = m.avatarImg || (AVATARS.find(a => a.id === m.avatarId) || AVATARS[0]).img;
      const isAct = activeScope === m.id;
      html += `
        <button type="button" class="scope-btn ${isAct ? 'active' : ''}" data-scope="${m.id}" title="Chi riêng cho ${m.name}">
          <img src="${imgSrc}" class="avatar-img-xs" alt="${m.name}">
          <span>${m.name}</span>
        </button>
      `;
    });

    container.innerHTML = html;
    container.onclick = (e) => {
      const btn = e.target.closest('.scope-btn');
      if (btn && btn.dataset.scope) {
        onChange(btn.dataset.scope);
      }
    };
  },

  // ==================== UNIFIED SCOPE FILTER ====================
  filterTransactionsByScope(transactions, scope) {
    if (!scope || scope === 'all') return transactions;
    if (scope === 'family') {
      return transactions.filter(t => !t.beneficiaryId || t.beneficiaryId === 'family');
    }
    // Specific member ID (e.g. Ba Dâu)
    return transactions.filter(t => {
      if (t.type === 'income') {
        // Income is earned by member
        return t.memberId === scope;
      } else {
        // Expense belongs to member if:
        // 1. beneficiary is this member (chi riêng cho người này)
        // 2. OR this member is the payer (người này chi tiền)
        return t.beneficiaryId === scope || (!t.beneficiaryId && t.memberId === scope) || (t.beneficiaryId === 'family' && t.memberId === scope) || t.memberId === scope;
      }
    });
  },

  // ==================== BUDGET PAGE ====================
  renderBudgetPage() {
    const { year, month } = this.budgetMonth;
    const monthLabel = document.getElementById('budgetCurrentMonth');
    if (monthLabel) monthLabel.textContent = formatMonthLabel(year, month);

    this.renderScopeSelector('budgetScopeToggle', this.budgetScope, (scope) => {
      this.budgetScope = scope;
      this.renderBudgetPage();
    });

    let transactions = Storage.getByMonth(year, month);
    transactions = this.filterTransactionsByScope(transactions, this.budgetScope);

    this.renderBudgetSummary(transactions, this.budgetMonth, this.budgetScope);
  },

  // ==================== SAVINGS & GOALS PAGE ====================
  renderSavingsPage() {
    this.renderScopeSelector('savingsScopeToggle', this.savingsScope, (scope) => {
      this.savingsScope = scope;
      this.renderSavingsPage();
    });

    const allGoals = Storage.getSavingsGoals();
    let goals = allGoals;
    if (this.savingsScope === 'family') {
      goals = allGoals.filter(g => !g.memberId || g.memberId === 'family');
    } else if (this.savingsScope !== 'all') {
      goals = allGoals.filter(g => g.memberId === this.savingsScope);
    }

    // Stats calculations
    let totalCurrent = 0;
    let totalTarget = 0;
    goals.forEach(g => {
      totalCurrent += (g.currentAmount || 0);
      totalTarget += (g.targetAmount || 0);
    });
    const overallPercent = totalTarget > 0 ? Math.round((totalCurrent / totalTarget) * 100) : 0;

    // This month's deposits
    const now = new Date();
    const curYear = now.getFullYear();
    const curMonth = now.getMonth() + 1;
    const allLogs = Storage.getSavingsLogs();
    let thisMonthAdded = 0;
    allLogs.forEach(l => {
      if (l.date && l.type === 'deposit') {
        const d = new Date(l.date);
        if (d.getFullYear() === curYear && (d.getMonth() + 1) === curMonth) {
          if (this.savingsScope === 'all') thisMonthAdded += l.amount;
          else if (this.savingsScope === 'family' && (!l.memberId || l.memberId === 'family')) thisMonthAdded += l.amount;
          else if (l.memberId === this.savingsScope) thisMonthAdded += l.amount;
        }
      }
    });

    // Render Stats Banner
    const banner = document.getElementById('savingsOverviewBanner');
    if (banner) {
      banner.innerHTML = `
        <div class="savings-stats-row">
          <div class="savings-stat-card">
            <span class="savings-stat-title">💎 Tổng tài sản tích lũy</span>
            <span class="savings-stat-val">${formatCurrency(totalCurrent)}</span>
          </div>
          <div class="savings-stat-card">
            <span class="savings-stat-title">🌱 Nạp thêm tháng ${curMonth}</span>
            <span class="savings-stat-val">+${formatCurrency(thisMonthAdded)}</span>
          </div>
          <div class="savings-stat-card">
            <span class="savings-stat-title">🎯 Tiến độ tổng thể</span>
            <span class="savings-stat-val" style="color:var(--primary);">${overallPercent}%</span>
          </div>
        </div>

        <div class="budget-progress-track" style="height:9px; margin-top:4px;">
          <div class="budget-progress-fill safe" style="width: ${Math.min(overallPercent, 100)}%;"></div>
        </div>
      `;
    }

    // Render Goals Grid
    const grid = document.getElementById('savingsGoalsGrid');
    if (grid) {
      if (goals.length === 0) {
        grid.innerHTML = `<div class="empty-state" style="grid-column: 1 / -1;"><span class="empty-icon">🐷</span><p>Chưa có hũ tiết kiệm nào</p><p class="text-muted">Bấm "+ Tạo hũ tiết kiệm" để bắt đầu tích lũy</p></div>`;
      } else {
        grid.innerHTML = goals.map(g => {
          const cur = g.currentAmount || 0;
          const tar = g.targetAmount || 0;
          const pct = tar > 0 ? Math.round((cur / tar) * 100) : 0;
          const member = g.memberId && g.memberId !== 'family' ? Storage.getMemberById(g.memberId) : null;
          const memberName = member ? member.name : 'Cả nhà';
          const memberImg = member ? (member.avatarImg || (AVATARS.find(a => a.id === member.avatarId) || AVATARS[0]).img) : '';

          let dateLabel = '';
          if (g.targetDate) {
            const parts = g.targetDate.split('-');
            if (parts.length >= 2) dateLabel = `📅 Hạn: Th${parseInt(parts[1])}/${parts[0]}`;
          }

          return `
            <div class="savings-goal-card" data-goal-id="${g.id}">
              <div class="savings-card-header">
                <div class="savings-card-title-box">
                  <div class="savings-card-emoji">${g.emoji || '🐷'}</div>
                  <div>
                    <div class="savings-card-name">${g.name}</div>
                    <div class="savings-card-member">
                      ${member ? `<img src="${memberImg}" class="avatar-img-xs">` : '👨‍👩‍👧‍👦'} ${memberName}
                    </div>
                  </div>
                </div>
                <div style="display:flex; gap:4px;">
                  <button class="btn-savings-icon edit-goal-btn" data-id="${g.id}" title="Sửa mục tiêu">✏️</button>
                  <button class="btn-savings-icon delete-goal-btn" data-id="${g.id}" title="Xóa hũ">🗑️</button>
                </div>
              </div>

              <div class="savings-card-amounts">
                <span class="savings-current-val">${formatCurrency(cur)}</span>
                <span class="savings-target-val">/ ${formatCurrency(tar)}</span>
              </div>

              <div class="budget-progress-track">
                <div class="budget-progress-fill ${pct >= 100 ? 'safe' : (pct >= 50 ? 'warning' : '')}" style="width: ${Math.min(pct, 100)}%;"></div>
              </div>

              <div class="savings-card-meta">
                <span class="budget-cat-percent ${pct >= 100 ? 'safe' : (pct >= 50 ? 'warning' : 'safe')}">${pct}%</span>
                ${dateLabel ? `<span class="savings-target-badge">${dateLabel}</span>` : ''}
              </div>

              <div class="savings-card-actions">
                <button class="btn-savings-deposit deposit-btn" data-id="${g.id}">
                  <span>➕</span> Nạp tiền
                </button>
                <button class="btn-savings-withdraw withdraw-btn" data-id="${g.id}">
                  <span>➖</span> Rút tiền
                </button>
                <button class="btn-savings-icon history-btn" data-id="${g.id}" title="Xem lịch sử">
                  <span>📜</span>
                </button>
              </div>
            </div>
          `;
        }).join('');

        // Bind goal card action buttons
        grid.querySelectorAll('.edit-goal-btn').forEach(btn => {
          btn.addEventListener('click', (e) => {
            e.stopPropagation();
            this.openSavingsGoalModal(btn.dataset.id);
          });
        });
        grid.querySelectorAll('.delete-goal-btn').forEach(btn => {
          btn.addEventListener('click', (e) => {
            e.stopPropagation();
            this.deleteSavingsGoal(btn.dataset.id);
          });
        });
        grid.querySelectorAll('.deposit-btn').forEach(btn => {
          btn.addEventListener('click', (e) => {
            e.stopPropagation();
            this.openSavingsActionModal(btn.dataset.id, 'deposit');
          });
        });
        grid.querySelectorAll('.withdraw-btn').forEach(btn => {
          btn.addEventListener('click', (e) => {
            e.stopPropagation();
            this.openSavingsActionModal(btn.dataset.id, 'withdraw');
          });
        });
        grid.querySelectorAll('.history-btn').forEach(btn => {
          btn.addEventListener('click', (e) => {
            e.stopPropagation();
            this.openSavingsHistoryModal(btn.dataset.id);
          });
        });
      }
    }

    // Render Recent Logs
    const logsContainer = document.getElementById('savingsRecentLogs');
    if (logsContainer) {
      let logs = allLogs;
      if (this.savingsScope === 'family') {
        logs = allLogs.filter(l => !l.memberId || l.memberId === 'family');
      } else if (this.savingsScope !== 'all') {
        logs = allLogs.filter(l => l.memberId === this.savingsScope);
      }

      const recent = logs.slice(0, 10);
      if (recent.length === 0) {
        logsContainer.innerHTML = `<div class="empty-state"><span class="empty-icon">📝</span><p>Chưa có giao dịch nạp/rút tiết kiệm nào</p></div>`;
      } else {
        logsContainer.innerHTML = recent.map(l => {
          const isDep = l.type === 'deposit';
          const member = l.memberId && l.memberId !== 'family' ? Storage.getMemberById(l.memberId) : null;
          const memberText = member ? ` · 👤 ${member.name}` : ' · 👨‍👩‍👧‍👦 Cả nhà';

          return `
            <div class="savings-log-item">
              <div class="savings-log-left">
                <span class="savings-log-badge ${l.type}">${isDep ? 'Nạp tiền' : 'Rút tiền'}</span>
                <div class="savings-log-info">
                  <div class="savings-log-title">${l.goalName || 'Tiết kiệm'}${l.note ? ' · ' + l.note : ''}</div>
                  <div class="savings-log-sub">${formatDate(l.date)}${memberText}</div>
                </div>
              </div>
              <div class="savings-log-amount ${l.type}">
                ${isDep ? '+' : '-'}${formatCurrency(l.amount)}
              </div>
            </div>
          `;
        }).join('');
      }
    }
  },

  // ==================== SYNC ====================
  initSyncStatus() {
    this.updateSyncStatusBar();
    if (Storage.isOnline()) this.syncFromSheets(true);

    // Auto-sync when user switches back to app tab
    document.addEventListener('visibilitychange', () => {
      if (document.visibilityState === 'visible' && Storage.isOnline()) {
        this.syncFromSheets(true);
      }
    });
    window.addEventListener('focus', () => {
      if (Storage.isOnline()) {
        this.syncFromSheets(true);
      }
    });

    // Periodic auto-sync every 30 seconds
    setInterval(() => {
      if (Storage.isOnline() && document.visibilityState === 'visible') {
        this.syncFromSheets(true);
      }
    }, 30000);
  },
  updateSyncStatusBar() {
    const bar = document.getElementById('syncStatus');
    if (!bar) return;
    if (Storage.isOnline()) {
      bar.style.display = 'flex'; bar.className = 'sync-status connected';
      bar.querySelector('.sync-text').textContent = '🟢 Đã kết nối Google Sheets';
    } else { bar.style.display = 'none'; }
  },
  async syncFromSheets(silent = false) {
    if (!Storage.isOnline()) return;
    const bar = document.getElementById('syncStatus');
    if (bar) { bar.style.display = 'flex'; bar.className = 'sync-status'; bar.querySelector('.sync-text').textContent = '🔄 Đang đồng bộ...'; }
    try {
      await Storage.syncFromSheets();
      if (bar) { bar.className = 'sync-status connected'; bar.querySelector('.sync-text').textContent = '🟢 Đã đồng bộ'; }
      this.renderFamilyAvatars();
      this.renderCurrentPage();
      if (!silent) this.showToast('Đã đồng bộ ✅');
    } catch(e) {
      if (bar) { bar.className = 'sync-status offline'; bar.querySelector('.sync-text').textContent = '🔴 Lỗi đồng bộ'; }
      if (!silent) this.showToast('Lỗi đồng bộ', 'error');
    }
  },

  // ==================== INIT FILTERS ====================
  initFilterMonth() {
    const f = document.getElementById('filterMonth');
    if (f) { const d = new Date(); f.value = `${d.getFullYear()}-${(d.getMonth()+1).toString().padStart(2,'0')}`; }
    this.populateCategoryFilter();
  },
  populateCategoryFilter() {
    const f = document.getElementById('filterCategory');
    if (!f) return;
    f.innerHTML = '<option value="all">Tất cả</option>';
    getAllCategories().forEach(c => { f.innerHTML += `<option value="${c.id}">${c.emoji} ${c.label}</option>`; });
  },

  // ==================== DASHBOARD ====================
  renderDashboard() {
    this.checkPendingInbox();
    const { year, month } = this.currentMonth;
    let transactions = Storage.getByMonth(year, month);
    const prev = navigateMonth(year, month, -1);
    let prevTransactions = Storage.getByMonth(prev.year, prev.month);

    // Dynamic scope selector
    this.renderScopeSelector('dashboardScopeToggle', this.dashboardScope, (scope) => {
      this.dashboardScope = scope;
      this.renderDashboard();
    });

    // Apply unified dashboard scope filter (all | family | memberId)
    transactions = this.filterTransactionsByScope(transactions, this.dashboardScope);
    prevTransactions = this.filterTransactionsByScope(prevTransactions, this.dashboardScope);

    let income = 0, expense = 0;
    transactions.forEach(t => { if (t.type === 'income') income += t.amount; else expense += t.amount; });
    const balance = income - expense;
    const savingsRate = income > 0 ? ((income - expense) / income * 100) : 0;

    let prevIncome = 0, prevExpense = 0;
    prevTransactions.forEach(t => { if (t.type === 'income') prevIncome += t.amount; else prevExpense += t.amount; });

    // Greeting
    const greeting = document.getElementById('greetingText');
    if (greeting) greeting.textContent = `Xin chào, ${Storage.getFamilyName()}! 👋`;

    // Family avatars
    this.renderFamilyAvatars();

    // Mood
    const expenseChange = calcChange(expense, prevExpense);
    const mood = analyzeMood(savingsRate, expenseChange);
    const moodCard = document.getElementById('moodCard');
    if (moodCard) {
      moodCard.style.background = mood.bgColor;
      moodCard.style.borderColor = mood.color;
    }
    const moodEmoji = document.getElementById('moodEmoji');
    if (moodEmoji) moodEmoji.textContent = mood.emoji;
    const moodLabel = document.getElementById('moodLabel');
    if (moodLabel) { moodLabel.textContent = mood.label; moodLabel.style.color = mood.color; }
    const moodMsg = document.getElementById('moodMessage');
    if (moodMsg) moodMsg.textContent = mood.message;

    // Smart tips
    const tips = generateSmartTips(transactions, prevTransactions);
    const tipsCont = document.getElementById('tipsContainer');
    if (tipsCont) {
      if (tips.length > 0) {
        tipsCont.style.display = 'flex';
        tipsCont.innerHTML = tips.map(t => `<div class="tip-item ${t.type}"><span>${t.emoji}</span><span>${t.text}</span></div>`).join('');
      } else { tipsCont.style.display = 'none'; }
    }

    // Summary cards
    document.getElementById('totalIncome').textContent = formatCurrency(income);
    document.getElementById('totalExpense').textContent = formatCurrency(expense);
    document.getElementById('balance').textContent = formatCurrency(balance);
    document.getElementById('savingsRate').textContent = savingsRate.toFixed(1) + '%';
    this.updateChangeBadge('incomeChange', calcChange(income, prevIncome), false);
    this.updateChangeBadge('expenseChange', expenseChange, true);
    document.getElementById('currentMonth').textContent = formatMonthLabel(year, month);

    // Charts
    Charts.renderCategoryChart('categoryChart', transactions);
    const monthsData = Storage.getLastNMonths(6);
    Charts.renderMonthlyChart('monthlyChart', monthsData);
    Charts.renderTrendChart('trendChart', monthsData);
    this.renderRecentTransactions(transactions);
  },

  renderFamilyAvatars() {
    const container = document.getElementById('familyAvatars');
    if (!container) return;
    const members = Storage.getMembers();
    container.innerHTML = members.map(m => {
      const imgSrc = m.avatarImg || (AVATARS.find(a => a.id === m.avatarId) || AVATARS[0]).img;
      return `<div class="family-avatar" style="border-color: ${m.color}" title="${m.name}"><img src="${imgSrc}" alt="${m.name}" class="avatar-img"></div>`;
    }).join('');
  },

  updateChangeBadge(id, change, invert) {
    const el = document.getElementById(id);
    if (!el) return;
    el.textContent = `${change >= 0 ? '+' : ''}${change.toFixed(1)}%`;
    el.className = 'card-change';
    el.classList.add(invert ? (change <= 0 ? 'positive' : 'negative') : (change >= 0 ? 'positive' : 'negative'));
  },

  renderRecentTransactions(transactions) {
    const c = document.getElementById('recentTransactions');
    if (!c) return;
    const sorted = [...transactions].sort((a, b) => new Date(b.date) - new Date(a.date)).slice(0, 5);
    if (sorted.length === 0) {
      c.innerHTML = `<div class="empty-state"><span class="empty-icon">📝</span><p>Chưa có giao dịch nào</p><p class="text-muted">Bấm nút + để thêm giao dịch đầu tiên</p></div>`;
    } else {
      c.innerHTML = sorted.map(t => this.renderRecentItem(t)).join('');
    }
  },

  // ==================== TRANSACTIONS PAGE ====================
  renderTransactions() {
    let txs = Storage.getLocal();
    const fm = document.getElementById('filterMonth')?.value;
    const fb = document.getElementById('filterBeneficiary')?.value || 'all';
    const ft = document.getElementById('filterType')?.value || 'all';
    const fc = document.getElementById('filterCategory')?.value || 'all';
    const fs = document.getElementById('filterSort')?.value || 'date-desc';

    // Populate filterBeneficiary options dynamically
    const beneSelect = document.getElementById('filterBeneficiary');
    if (beneSelect && beneSelect.options.length <= 2) {
      const members = Storage.getMembers();
      beneSelect.innerHTML = '<option value="all">Tất cả</option><option value="family">👨‍👩‍👧‍👦 Cả nhà</option>';
      members.forEach(m => {
        beneSelect.innerHTML += `<option value="${m.id}">👤 ${m.name}</option>`;
      });
      if (fb) beneSelect.value = fb;
    }

    if (fm) { const [y,m] = fm.split('-').map(Number); txs = txs.filter(t => { const d=new Date(t.date); return d.getFullYear()===y&&(d.getMonth()+1)===m; }); }
    txs = this.filterTransactionsByScope(txs, fb);
    if (ft !== 'all') txs = txs.filter(t => t.type === ft);
    if (fc !== 'all') txs = txs.filter(t => t.category === fc);
    switch(fs) {
      case 'date-desc': txs.sort((a,b) => new Date(b.date)-new Date(a.date)); break;
      case 'date-asc': txs.sort((a,b) => new Date(a.date)-new Date(b.date)); break;
      case 'amount-desc': txs.sort((a,b) => b.amount-a.amount); break;
      case 'amount-asc': txs.sort((a,b) => a.amount-b.amount); break;
    }
    const tbody = document.getElementById('transactionsBody');
    const empty = document.getElementById('emptyTransactions');
    const table = document.getElementById('transactionsTable');
    if (!tbody) return;
    if (txs.length === 0) {
      tbody.innerHTML = ''; if (table) table.style.display = 'none'; if (empty) empty.style.display = 'flex';
    } else {
      if (table) table.style.display = ''; if (empty) empty.style.display = 'none';
      tbody.innerHTML = txs.map(t => this.renderTransactionRow(t)).join('');
      let ti=0, te=0; txs.forEach(t => { if(t.type==='income') ti+=t.amount; else te+=t.amount; });
      const net = ti-te;
      const tot = document.getElementById('tableTotal');
      if (tot) tot.innerHTML = `<strong class="${net>=0?'amount-col income':'amount-col expense'}">${formatCurrency(net)}</strong>`;
    }
  },

  // ==================== LOANS PAGE ====================
  renderLoans() {
    const loans = Storage.getLoans();
    let totalDebt = 0, totalMonthly = 0, totalInterest = 0;
    loans.forEach(l => {
      const status = calculateLoanStatus(l);
      totalDebt += status.remainingBalance;
      totalMonthly += l.monthlyPayment;
      totalInterest += status.totalInterestPaid;
    });

    document.getElementById('totalDebt').textContent = formatCurrency(totalDebt);
    document.getElementById('monthlyPaymentTotal').textContent = formatCurrency(totalMonthly);
    document.getElementById('totalInterestPaid').textContent = formatCurrency(totalInterest);

    const list = document.getElementById('loansList');
    if (!list) return;
    if (loans.length === 0) {
      list.innerHTML = `<div class="empty-state"><span class="empty-icon">💳</span><p>Chưa có khoản vay nào</p><p class="text-muted">Bấm "Thêm khoản vay" để bắt đầu theo dõi</p></div>`;
    } else {
      list.innerHTML = loans.map(l => this.renderLoanCard(l)).join('');
    }
  },

  renderLoanCard(loan) {
    const status = calculateLoanStatus(loan);
    const loanType = LOAN_TYPES.find(t => t.id === loan.loanType);
    const emoji = loan.emoji || (loanType ? loanType.emoji : '💳');
    const endDate = status.estimatedEndDate;
    const endDateStr = `${endDate.getMonth()+1}/${endDate.getFullYear()}`;

    return `
      <div class="loan-card">
        <div class="loan-header">
          <div class="loan-title">
            <span class="loan-emoji">${emoji}</span>
            <div>
              <h4>${loan.name}</h4>
              ${loan.note ? `<span class="loan-note">${loan.note}</span>` : ''}
            </div>
          </div>
          <div class="loan-actions">
            <button class="edit-btn edit-loan-btn" data-id="${loan.id}" title="Sửa">✏️</button>
            <button class="delete-btn delete-loan-btn" data-id="${loan.id}" title="Xóa">🗑️</button>
          </div>
        </div>
        <div class="loan-progress">
          <div class="loan-progress-bar"><div class="loan-progress-fill" style="width:${status.progressPercent.toFixed(1)}%"></div></div>
          <div class="loan-progress-text">
            <span>Đã trả ${status.progressPercent.toFixed(1)}%</span>
            <span>Dự kiến hết: ${endDateStr}</span>
          </div>
        </div>
        <div class="loan-stats">
          <div class="loan-stat"><div class="loan-stat-label">Gốc vay</div><div class="loan-stat-value">${formatCurrency(loan.principal)}</div></div>
          <div class="loan-stat"><div class="loan-stat-label">Dư nợ còn lại</div><div class="loan-stat-value">${formatCurrency(status.remainingBalance)}</div></div>
          <div class="loan-stat"><div class="loan-stat-label">Lãi suất</div><div class="loan-stat-value">${loan.interestRate}%/năm</div></div>
          <div class="loan-stat"><div class="loan-stat-label">Trả/tháng</div><div class="loan-stat-value">${formatCurrency(loan.monthlyPayment)}</div></div>
        </div>
      </div>`;
  },

  // ==================== ANALYTICS PAGE ====================
  renderAnalytics() {
    const { year, month } = this.analyticsMonth;
    const txs = Storage.getByMonth(year, month);
    const prev = navigateMonth(year, month, -1);
    const prevTxs = Storage.getByMonth(prev.year, prev.month);
    let inc=0, exp=0; txs.forEach(t => { if(t.type==='income') inc+=t.amount; else exp+=t.amount; });
    let pInc=0, pExp=0; prevTxs.forEach(t => { if(t.type==='income') pInc+=t.amount; else pExp+=t.amount; });
    const sav = inc-exp, sr = inc>0?((sav/inc)*100):0;
    const pSav = pInc-pExp, pSr = pInc>0?((pSav/pInc)*100):0;

    document.getElementById('analyticsIncome').textContent = formatCurrency(inc);
    document.getElementById('analyticsExpense').textContent = formatCurrency(exp);
    document.getElementById('analyticsSavings').textContent = formatCurrency(sav);
    document.getElementById('analyticsSavingsRate').textContent = sr.toFixed(1)+'%';
    document.getElementById('analyticsCurrentMonth').textContent = formatMonthLabel(year, month);

    this.updateAnalyticsChange('analyticsIncomeChange', inc, pInc, false);
    this.updateAnalyticsChange('analyticsExpenseChange', exp, pExp, true);
    this.updateAnalyticsChange('analyticsSavingsChange', sav, pSav, false);
    this.updateAnalyticsChange('analyticsSavingsRateChange', sr, pSr, false);
    this.renderTopCategories(txs);
    Charts.renderCategoryTrendChart('categoryTrendChart', Storage.getLastNMonths(6));
  },

  updateAnalyticsChange(cId, cur, prev, inv) {
    const c = document.getElementById(cId); if(!c) return;
    const ch = calcChange(cur, prev);
    const badge = c.querySelector('.change-badge');
    if(badge) { badge.textContent = `${ch>=0?'+':''}${ch.toFixed(1)}%`; badge.className='change-badge'; badge.classList.add(inv?(ch<=0?'up':'down'):(ch>=0?'up':'down')); }
  },

  renderTopCategories(txs) {
    const c = document.getElementById('topCategories'); if(!c) return;
    const exps = txs.filter(t => t.type==='expense');
    if(exps.length===0) { c.innerHTML = `<div class="empty-state"><span class="empty-icon">📊</span><p>Chưa có dữ liệu</p></div>`; return; }
    const totals = {}; exps.forEach(t => { totals[t.category]=(totals[t.category]||0)+t.amount; });
    const total = exps.reduce((s,t)=>s+t.amount,0);
    c.innerHTML = Object.entries(totals).sort((a,b)=>b[1]-a[1]).map(([id,amt]) => {
      const cat = getCategoryById(id); const pct = ((amt/total)*100).toFixed(1);
      return `<div class="top-category-item"><div class="top-cat-info"><span class="top-cat-emoji">${cat?cat.emoji:'📦'}</span><span class="top-cat-label">${cat?cat.label:'Khác'}</span></div><div class="top-cat-bar-wrapper"><div class="top-cat-bar" style="width:${pct}%"></div></div><div class="top-cat-values"><span class="top-cat-amount">${formatCurrency(amt)}</span><span class="top-cat-percent">${pct}%</span></div></div>`;
    }).join('');
  },

  // ==================== SETTINGS PAGE ====================
  renderSettings() {
    const st = document.getElementById('syncToggle');
    const ac = document.getElementById('apiConfig');
    const au = document.getElementById('apiUrl');
    if (st) st.checked = Storage.getSyncMode() === 'sheets';
    if (ac) ac.style.display = Storage.getSyncMode() === 'sheets' ? 'block' : 'none';
    if (Storage.isOnline()) { const sa = document.getElementById('syncActions'); if(sa) sa.style.display = 'flex'; }
    
    // Webhook URL
    const webhookInput = document.getElementById('webhookUrlInput');
    const url = Storage.getApiUrl();
    if (webhookInput) {
      webhookInput.value = url ? `${url}?action=bankNotification` : 'Vui lòng kết nối Google Sheets trước để nhận link Webhook cá nhân';
    }

    this.renderMemberEmailList();
    this.renderMemberManagement();
    this.renderReminderSettings();
  },

  renderMemberEmailList() {
    const container = document.getElementById('memberEmailList');
    if (!container) return;
    const memberEmails = JSON.parse(localStorage.getItem('cashflow_member_emails') || '[]');
    if (memberEmails.length === 0) {
      container.innerHTML = '';
      return;
    }
    const roleLabels = { mom: '🌸 Vợ', dad: '👔 Chồng', child: '🧒 Con', other: '👤 Khác' };
    container.innerHTML = memberEmails.map((m, i) => `
      <div style="display:flex; align-items:center; justify-content:space-between; padding:8px 10px; margin-bottom:4px; background:rgba(34,197,94,0.08); border:1px solid rgba(34,197,94,0.2); border-radius:8px; font-size:0.78rem;">
        <div style="flex:1;">
          <span style="font-weight:600;">${roleLabels[m.role] || '👤'}</span>
          <span style="color:var(--text-secondary); margin-left:6px;">${m.email}</span>
          <span style="color:#4ade80; font-size:0.68rem; margin-left:6px;">✅ Đã kích hoạt</span>
        </div>
        <button type="button" class="btn-ghost" style="font-size:0.7rem; padding:2px 8px; color:#f87171;" onclick="App.removeMemberEmail(${i})">✕</button>
      </div>
    `).join('');
  },

  removeMemberEmail(index) {
    const memberEmails = JSON.parse(localStorage.getItem('cashflow_member_emails') || '[]');
    memberEmails.splice(index, 1);
    localStorage.setItem('cashflow_member_emails', JSON.stringify(memberEmails));
    this.renderMemberEmailList();
    this.showToast('Đã xóa email thành viên');
  },

  // ==================== MEMBER MANAGEMENT ====================
  _editingMemberId: null,

  renderMemberManagement() {
    const list = document.getElementById('memberManagementList');
    if (!list) return;
    const members = Storage.getMembers();
    list.innerHTML = members.map((m, i) => {
      const imgSrc = m.avatarImg || (AVATARS.find(a => a.id === m.avatarId) || AVATARS[0]).img;
      const isEditing = this._editingMemberId === m.id;
      return `
        <div class="member-manage-row" data-id="${m.id}">
          <div class="member-manage-avatar" data-id="${m.id}" data-action="change-avatar" title="Đổi avatar">
            <img src="${imgSrc}" alt="${m.name}">
          </div>
          <div class="member-manage-info">
            ${isEditing
              ? `<input class="member-manage-name-input" value="${m.name}" data-id="${m.id}" autofocus>`
              : `<span class="member-manage-name">${m.name}</span>`
            }
            <span><span class="member-manage-color" style="background:${m.color}"></span> ${isEditing ? 'Đang sửa...' : ''}</span>
          </div>
          <div class="member-manage-actions">
            ${isEditing
              ? `<button class="save-member-btn" data-id="${m.id}" title="Lưu">✓</button>
                 <button data-id="${m.id}" data-action="cancel-edit" title="Hủy">✕</button>`
              : `<button data-id="${m.id}" data-action="edit-member" title="Sửa tên">✏️</button>
                 <button class="delete-member-btn" data-id="${m.id}" data-action="delete-member" title="Xóa">🗑️</button>`
            }
          </div>
        </div>`;
    }).join('');
  },

  handleMemberManagementClick(e) {
    const btn = e.target.closest('button[data-action], div[data-action]');
    if (!btn) {
      // Check save button
      const saveBtn = e.target.closest('.save-member-btn');
      if (saveBtn) {
        const id = saveBtn.dataset.id;
        const input = document.querySelector(`.member-manage-name-input[data-id="${id}"]`);
        if (input && input.value.trim()) {
          Storage.updateMember(id, { name: input.value.trim() });
          this._editingMemberId = null;
          this.renderMemberManagement();
          this.showToast('Đã cập nhật ✅');
          this.renderFamilyAvatars();
        }
        return;
      }
      return;
    }
    const action = btn.dataset.action;
    const id = btn.dataset.id;

    if (action === 'edit-member') {
      this._editingMemberId = id;
      this.renderMemberManagement();
      setTimeout(() => {
        document.querySelector(`.member-manage-name-input[data-id="${id}"]`)?.focus();
      }, 50);
    } else if (action === 'cancel-edit') {
      this._editingMemberId = null;
      this.renderMemberManagement();
    } else if (action === 'delete-member') {
      const member = Storage.getMemberById(id);
      const members = Storage.getMembers();
      if (members.length <= 1) {
        this.showToast('Cần ít nhất 1 thành viên', 'error');
        return;
      }
      if (confirm(`Xóa thành viên "${member?.name}"? Giao dịch cũ vẫn giữ nguyên.`)) {
        Storage.deleteMember(id);
        this.renderMemberManagement();
        this.renderFamilyAvatars();
        this.showToast('Đã xóa 🗑️');
      }
    } else if (action === 'change-avatar') {
      this.showMemberAvatarPicker(id, btn);
    }
  },

  showMemberAvatarPicker(memberId, button) {
    const member = Storage.getMemberById(memberId);
    if (!member) return;
    const currentAv = AVATARS.find(a => a.id === member.avatarId) || AVATARS[0];
    const currentImg = member.avatarImg || currentAv.img;

    this.openAvatarModal({
      currentImg: currentImg,
      onSelect: (selected) => {
        Storage.updateMember(memberId, {
          avatarId: selected.id || 'custom',
          avatarImg: selected.img,
          avatar: selected.emoji || '👤'
        });
        this.renderMemberManagement();
        this.renderFamilyAvatars();
        this.showToast('Đã đổi avatar thành công! ✅');
      }
    });
  },

  // ==================== AVATAR MODAL CONTROLLER ====================
  _avatarModalCallback: null,
  _currentSelectedAvatar: null,

  openAvatarModal({ currentImg, onSelect }) {
    const modal = document.getElementById('avatarModal');
    if (!modal) return;
    this._avatarModalCallback = onSelect;
    this._currentSelectedAvatar = { img: currentImg || AVATARS[0].img };

    // Update preview
    const previewImg = document.getElementById('uploadPreviewImg');
    if (previewImg) previewImg.src = this._currentSelectedAvatar.img;

    // Render presets
    this.renderAvatarPresets();

    modal.classList.add('active');
  },

  closeAvatarModal() {
    const modal = document.getElementById('avatarModal');
    if (modal) modal.classList.remove('active');
    this._avatarModalCallback = null;
    this._currentSelectedAvatar = null;
  },

  renderAvatarPresets() {
    const grid = document.getElementById('avatarPresetsGrid');
    if (!grid) return;
    grid.innerHTML = AVATARS.map(a => {
      const isSelected = this._currentSelectedAvatar && this._currentSelectedAvatar.img === a.img;
      return `
        <button type="button" class="avatar-preset-card ${isSelected ? 'selected' : ''}" data-avatar-id="${a.id}">
          <img src="${a.img}" alt="${a.label}" class="preset-img">
          <span class="preset-label">${a.label}</span>
        </button>
      `;
    }).join('');
  },

  handleAvatarFileSelect(e) {
    const file = e.target.files?.[0];
    if (!file) return;
    this.showToast('Đang xử lý ảnh...');
    processImageFile(file, 128).then(dataUrl => {
      this._currentSelectedAvatar = {
        id: 'uploaded_' + Date.now(),
        img: dataUrl,
        emoji: '📸',
        label: 'Ảnh tự tải'
      };
      const previewImg = document.getElementById('uploadPreviewImg');
      if (previewImg) previewImg.src = dataUrl;
      // Deselect presets
      document.querySelectorAll('.avatar-preset-card').forEach(c => c.classList.remove('selected'));
      this.showToast('Đã tải ảnh lên! Bấm Áp dụng để lưu ✅');
    }).catch(err => {
      this.showToast(err.message || 'Lỗi đọc ảnh', 'error');
    });
  },

  handlePresetSelect(e) {
    const card = e.target.closest('.avatar-preset-card');
    if (!card) return;
    const avId = card.dataset.avatarId;
    const av = AVATARS.find(a => a.id === avId);
    if (!av) return;

    this._currentSelectedAvatar = {
      id: av.id,
      img: av.img,
      emoji: av.emoji,
      label: av.label
    };

    const previewImg = document.getElementById('uploadPreviewImg');
    if (previewImg) previewImg.src = av.img;

    document.querySelectorAll('.avatar-preset-card').forEach(c => c.classList.remove('selected'));
    card.classList.add('selected');
  },

  confirmAvatarSelection() {
    if (this._currentSelectedAvatar && this._avatarModalCallback) {
      this._avatarModalCallback(this._currentSelectedAvatar);
    }
    this.closeAvatarModal();
  },

  handleAddMember() {
    const members = Storage.getMembers();
    if (members.length >= 6) {
      this.showToast('Tối đa 6 thành viên', 'error');
      return;
    }
    const name = prompt('Tên thành viên mới:');
    if (!name || !name.trim()) return;
    const avIndex = members.length % AVATARS.length;
    const av = AVATARS[avIndex];
    Storage.addMember({
      name: name.trim(),
      avatarId: av.id,
      avatar: av.emoji,
      avatarImg: av.img,
      color: MEMBER_COLORS[members.length % MEMBER_COLORS.length]
    });
    this.renderMemberManagement();
    this.renderFamilyAvatars();
    this.showToast('Đã thêm thành viên ✅');
  },

  // ==================== NOTIFICATIONS ====================
  renderReminderSettings() {
    const list = document.getElementById('reminderMemberList');
    const statusEl = document.getElementById('notificationStatus');
    const btn = document.getElementById('enableNotifBtn');
    if (!list) return;

    // Show notification permission status
    if ('Notification' in window) {
      const perm = Notification.permission;
      if (perm === 'granted') {
        if (statusEl) { statusEl.style.display = 'block'; statusEl.className = 'notification-status granted'; statusEl.textContent = '✅ Đã bật thông báo'; }
        if (btn) btn.style.display = 'none';
      } else if (perm === 'denied') {
        if (statusEl) { statusEl.style.display = 'block'; statusEl.className = 'notification-status denied'; statusEl.textContent = '❌ Thông báo bị chặn. Vào cài đặt trình duyệt để bật lại.'; }
        if (btn) btn.style.display = 'none';
      } else {
        if (statusEl) statusEl.style.display = 'none';
        if (btn) btn.style.display = 'block';
      }
    } else {
      if (statusEl) { statusEl.style.display = 'block'; statusEl.className = 'notification-status denied'; statusEl.textContent = 'Trình duyệt không hỗ trợ thông báo'; }
      if (btn) btn.style.display = 'none';
    }

    // Render member reminder rows
    const members = Storage.getMembers();
    const reminders = JSON.parse(localStorage.getItem(REMINDER_KEY) || '{}');
    list.innerHTML = members.map(m => {
      const reminder = reminders[m.id] || { enabled: false, time: '21:00' };
      const imgSrc = m.avatarImg || (AVATARS.find(a => a.id === m.avatarId) || AVATARS[0]).img;
      return `
        <div class="reminder-member-row">
          <div class="member-info">
            <img src="${imgSrc}" alt="${m.name}">
            <span>${m.name}</span>
          </div>
          <div class="reminder-time">
            <span>Nhắc lúc</span>
            <input type="time" value="${reminder.time}" data-member="${m.id}" class="reminder-time-input">
          </div>
          <label class="reminder-toggle">
            <input type="checkbox" ${reminder.enabled ? 'checked' : ''} data-member="${m.id}" class="reminder-toggle-input">
            <span class="slider"></span>
          </label>
        </div>`;
    }).join('');
  },

  async requestNotificationPermission() {
    if (!('Notification' in window)) {
      this.showToast('Trình duyệt không hỗ trợ', 'error');
      return;
    }
    const perm = await Notification.requestPermission();
    if (perm === 'granted') {
      this.showToast('Đã bật thông báo! 🔔');
      // Register service worker for background checks
      this.registerServiceWorker();
    } else {
      this.showToast('Thông báo bị từ chối', 'error');
    }
    this.renderReminderSettings();
  },

  async registerServiceWorker() {
    if ('serviceWorker' in navigator) {
      try {
        await navigator.serviceWorker.register('./sw.js');
      } catch (e) {
        console.warn('SW registration failed:', e);
      }
    }
  },

  saveReminderSettings(memberId, field, value) {
    const reminders = JSON.parse(localStorage.getItem(REMINDER_KEY) || '{}');
    if (!reminders[memberId]) reminders[memberId] = { enabled: false, time: '21:00' };
    reminders[memberId][field] = value;
    localStorage.setItem(REMINDER_KEY, JSON.stringify(reminders));
    // Schedule check
    this.scheduleReminderCheck();
  },

  scheduleReminderCheck() {
    // Clear existing timer
    if (this._reminderTimer) clearTimeout(this._reminderTimer);
    const reminders = JSON.parse(localStorage.getItem(REMINDER_KEY) || '{}');
    const now = new Date();
    let nearest = null;

    for (const memberId in reminders) {
      const r = reminders[memberId];
      if (!r.enabled) continue;
      const [h, min] = r.time.split(':').map(Number);
      const target = new Date(now);
      target.setHours(h, min, 0, 0);
      if (target <= now) target.setDate(target.getDate() + 1); // tomorrow
      if (!nearest || target < nearest.time) {
        nearest = { memberId, time: target };
      }
    }

    if (nearest) {
      const delay = nearest.time - now;
      this._reminderTimer = setTimeout(() => this.fireReminder(nearest.memberId), delay);
    }
  },

  fireReminder(memberId) {
    // Check if any transaction was added today by this member
    const today = new Date().toISOString().split('T')[0];
    const txs = Storage.getLocal();
    const hasTodayTx = txs.some(t => t.date === today && t.memberId === memberId);

    if (!hasTodayTx) {
      const member = Storage.getMemberById(memberId);
      const name = member ? member.name : 'Bạn';
      if ('Notification' in window && Notification.permission === 'granted') {
        new Notification('CashFlow - Nhắc nhở 📝', {
          body: `${name} ơi, hôm nay chưa nhập giao dịch nào. Đừng quên ghi chép nhé!`,
          icon: member?.avatarImg || 'data:image/svg+xml,<svg xmlns="http://www.w3.org/2000/svg" viewBox="0 0 100 100"><text y=".9em" font-size="90">💰</text></svg>'
        });
      }
    }
    // Reschedule for next day
    this.scheduleReminderCheck();
  },

  async handleTestConnection() {
    const url = document.getElementById('apiUrl')?.value?.trim();
    if (!url) { this.showToast('Nhập URL', 'error'); return; }
    Storage.setApiUrl(url);
    const se = document.getElementById('connectionStatus'), si = document.getElementById('statusIcon'), st = document.getElementById('statusText');
    if(se) { se.style.display='flex'; se.className='connection-status loading'; } if(si) si.textContent='⏳'; if(st) st.textContent='Đang kiểm tra...';
    const r = await Storage.testConnection();
    if (r.success) {
      if(se) se.className='connection-status success'; if(si) si.textContent='✅'; if(st) st.textContent=r.message;
      const sa = document.getElementById('syncActions'); if(sa) sa.style.display='flex';
      Storage.setSyncMode('sheets'); this.updateSyncStatusBar(); this.showToast('Kết nối thành công! ✅');
    } else {
      if(se) se.className='connection-status error'; if(si) si.textContent='❌'; if(st) st.textContent=r.message;
      this.showToast('Lỗi: '+r.message, 'error');
    }
  },

  handleSyncToggle() {
    const checked = document.getElementById('syncToggle')?.checked;
    Storage.setSyncMode(checked ? 'sheets' : 'local');
    const ac = document.getElementById('apiConfig'); if(ac) ac.style.display = checked ? 'block' : 'none';
    if (!checked) { const sa = document.getElementById('syncActions'); if(sa) sa.style.display='none'; }
    this.updateSyncStatusBar();
  },
  async handleSyncPull() { this.showToast('Đang tải...'); await this.syncFromSheets(false); },
  async handleSyncPush() {
    if (!confirm('Đẩy dữ liệu lên Sheet? Dữ liệu cũ trên Sheet sẽ bị ghi đè.')) return;
    try { await Storage.uploadAllToSheets(); await Storage.syncMembersToSheets(); await Storage.syncLoansToSheets(); this.showToast('Đã đẩy lên Sheet ✅'); }
    catch(e) { this.showToast('Lỗi: '+e.message,'error'); }
  },
  handleCopyScript() {
    const code = APPS_SCRIPT_CODE;
    navigator.clipboard.writeText(code).then(() => this.showToast('Đã copy! 📋')).catch(() => {
      const ta = document.createElement('textarea'); ta.value = code; document.body.appendChild(ta); ta.select(); document.execCommand('copy'); document.body.removeChild(ta); this.showToast('Đã copy! 📋');
    });
  },
  handleExportJson() {
    const data = { transactions: Storage.getLocal(), members: Storage.getMembers(), loans: Storage.getLoans(), familyName: Storage.getFamilyName() };
    const blob = new Blob([JSON.stringify(data,null,2)], {type:'application/json'});
    const link = document.createElement('a'); link.href = URL.createObjectURL(blob);
    link.download = `cashflow_backup_${new Date().toISOString().slice(0,10)}.json`;
    document.body.appendChild(link); link.click(); document.body.removeChild(link);
    this.showToast('Đã xuất JSON 📤');
  },
  handleImportJson(e) {
    const file = e.target.files[0]; if(!file) return;
    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const data = JSON.parse(evt.target.result);
        if (data.transactions) Storage.saveLocal(data.transactions);
        if (data.members) Storage.saveMembers(data.members);
        if (data.loans) Storage.saveLoans(data.loans);
        if (data.familyName) Storage.setFamilyName(data.familyName);
        this.showToast('Đã nhập dữ liệu! ✅'); this.renderCurrentPage();
      } catch(err) { this.showToast('File không hợp lệ','error'); }
    };
    reader.readAsText(file); e.target.value = '';
  },
  handleClearData() {
    if (!confirm('XÓA TOÀN BỘ dữ liệu? Không thể hoàn tác!')) return;
    if (!confirm('Xác nhận lần nữa?')) return;
    Storage.saveLocal([]); Storage.saveLoans([]); this.showToast('Đã xóa 🗑️'); this.renderCurrentPage();
  },

  // ==================== TRANSACTION MODAL ====================
  openModal(editId = null) {
    this.editingId = editId;
    const modal = document.getElementById('transactionModal'); if(!modal) return;
    const form = document.getElementById('transactionForm'); if(form) form.reset();

    if (editId) {
      document.getElementById('modalTitle').textContent = 'Sửa giao dịch';
      document.getElementById('editId').value = editId;
      const t = Storage.getLocal().find(x => x.id === editId);
      if (t) {
        this.selectedType = t.type;
        this.selectedCategory = t.category;
        this.selectedMemberId = t.memberId || null;
        this.selectedBeneficiaryId = t.beneficiaryId || 'family';
        document.getElementById('amount').value = formatNumberInput(t.amount);
        document.getElementById('date').value = t.date;
        document.getElementById('note').value = t.note || '';
        document.querySelectorAll('.type-btn').forEach(b => {
          b.classList.toggle('active', b.dataset.type === t.type);
        });
      }
    } else {
      document.getElementById('modalTitle').textContent = 'Thêm giao dịch';
      document.getElementById('editId').value = '';
      this.selectedType = 'expense';
      this.selectedCategory = null;
      const members = Storage.getMembers();
      this.selectedMemberId = members.length > 0 ? members[0].id : null;
      this.selectedBeneficiaryId = 'family';
      document.getElementById('amount').value = '';
      document.getElementById('date').value = new Date().toISOString().split('T')[0];
      document.querySelectorAll('.type-btn').forEach(b => {
        b.classList.toggle('active', b.dataset.type === 'expense');
      });
    }
    this.renderMemberSelector();
    this.renderBeneficiarySelector();
    this.renderCategories();
    modal.classList.add('active');
    setTimeout(() => document.getElementById('amount')?.focus(), 100);
  },

  renderMemberSelector() {
    const c = document.getElementById('memberSelector'); if(!c) return;
    const members = Storage.getMembers();
    c.innerHTML = members.map(m => {
      const imgSrc = m.avatarImg || (AVATARS.find(a => a.id === m.avatarId) || AVATARS[0]).img;
      return `
      <button type="button" class="member-btn ${this.selectedMemberId === m.id ? 'active' : ''}" data-id="${m.id}" style="--member-color:${m.color}">
        <span class="member-avatar"><img src="${imgSrc}" alt="${m.name}" class="avatar-img"></span>
        <span>${m.name}</span>
      </button>`;
    }).join('');
  },

  renderBeneficiarySelector() {
    const c = document.getElementById('beneficiarySelector'); if(!c) return;
    const members = Storage.getMembers();
    const isFamily = !this.selectedBeneficiaryId || this.selectedBeneficiaryId === 'family';
    let html = `
      <button type="button" class="beneficiary-btn ${isFamily ? 'active' : ''}" data-id="family">
        <span>👨‍👩‍👧‍👦 Cả nhà</span>
      </button>
    `;
    html += members.map(m => {
      const imgSrc = m.avatarImg || (AVATARS.find(a => a.id === m.avatarId) || AVATARS[0]).img;
      const isActive = this.selectedBeneficiaryId === m.id;
      return `
        <button type="button" class="beneficiary-btn ${isActive ? 'active' : ''}" data-id="${m.id}" style="--member-color:${m.color}">
          <img src="${imgSrc}" alt="${m.name}" class="avatar-img-sm" style="width:16px;height:16px;border-radius:50%;">
          <span>${m.name}</span>
        </button>
      `;
    }).join('');
    c.innerHTML = html;
  },

  closeModal() {
    document.getElementById('transactionModal')?.classList.remove('active');
    this.editingId = null;
  },

  handleSubmit(e) {
    e.preventDefault();
    const rawAmount = document.getElementById('amount').value;
    const amount = parseNumberInput(rawAmount);
    const date = document.getElementById('date').value;
    const note = document.getElementById('note').value;
    if (!amount || amount <= 0) { this.showToast('Nhập số tiền hợp lệ','error'); return; }
    if (!this.selectedCategory) { this.showToast('Chọn danh mục','error'); return; }
    if (!date) { this.showToast('Chọn ngày','error'); return; }
    const beneficiaryId = this.selectedBeneficiaryId || 'family';
    const data = {
      type: this.selectedType,
      amount,
      category: this.selectedCategory,
      date,
      note,
      memberId: this.selectedMemberId,
      beneficiaryId
    };
    if (this.editingId) { Storage.update(this.editingId, data); this.showToast('Đã cập nhật ✅'); }
    else { Storage.add(data); this.showToast('Đã thêm giao dịch ✅'); }
    this.closeModal(); this.renderCurrentPage();
  },

  deleteTransaction(id) {
    if (confirm('Xóa giao dịch này?')) { Storage.delete(id); this.showToast('Đã xóa 🗑️'); this.renderCurrentPage(); }
  },

  selectType(type) {
    this.selectedType = type; this.selectedCategory = null;
    document.querySelectorAll('.type-btn').forEach(b => { b.classList.remove('active'); if(b.dataset.type===type) b.classList.add('active'); });
    this.renderCategories();
  },
  selectCategory(catId) {
    this.selectedCategory = catId;
    document.querySelectorAll('.category-btn').forEach(el => el.classList.toggle('active', el.dataset.id===catId));
  },
  renderCategories() {
    const c = document.getElementById('categoryGrid'); if(!c) return;
    const cats = this.selectedType === 'expense' ? getExpenseCategories() : CATEGORIES.income;
    c.innerHTML = cats.map(cat =>
      `<button type="button" class="category-btn ${this.selectedCategory===cat.id?'active':''}" data-id="${cat.id}"><span class="cat-icon">${cat.emoji}</span><span>${cat.label}</span></button>`
    ).join('');
  },

  // ==================== KPI & CATEGORY DRILLDOWN MODALS ====================
  openKPIModal(type) {
    const { year, month } = this.currentMonth;
    let transactions = Storage.getByMonth(year, month);
    transactions = this.filterTransactionsByScope(transactions, this.dashboardScope);

    const modal = document.getElementById('kpiDrilldownModal');
    const titleEl = document.getElementById('kpiDrilldownTitle');
    const subEl = document.getElementById('kpiDrilldownSubtitle');
    const totalValEl = document.getElementById('kpiTotalVal');
    const listEl = document.getElementById('kpiDrilldownList');
    if (!modal || !listEl) return;

    let items = [];
    let total = 0;
    let title = '';
    let subtitle = `Tháng ${month}/${year}`;
    if (this.dashboardScope === 'family') subtitle += ' · 👨‍👩‍👧‍👦 Cả nhà';
    else if (this.dashboardScope !== 'all') {
      const m = Storage.getMemberById(this.dashboardScope);
      if (m) subtitle += ` · 👤 Cho ${m.name}`;
    }

    if (type === 'income') {
      title = '📥 Chi tiết Thu nhập';
      items = transactions.filter(t => t.type === 'income');
      total = items.reduce((s, t) => s + t.amount, 0);
    } else if (type === 'expense') {
      title = '📤 Chi tiết Chi tiêu';
      items = transactions.filter(t => t.type === 'expense');
      total = items.reduce((s, t) => s + t.amount, 0);
    } else if (type === 'balance') {
      title = '💎 Chi tiết Số dư';
      items = transactions;
      const inc = transactions.filter(t => t.type === 'income').reduce((s, t) => s + t.amount, 0);
      const exp = transactions.filter(t => t.type === 'expense').reduce((s, t) => s + t.amount, 0);
      total = inc - exp;
    } else if (type === 'savings') {
      title = '🎯 Tỷ lệ tiết kiệm';
      items = transactions.filter(t => t.type === 'expense');
      total = items.reduce((s, t) => s + t.amount, 0);
    }

    titleEl.textContent = title;
    subEl.textContent = subtitle;
    totalValEl.textContent = formatCurrency(total);
    totalValEl.className = 'kpi-total-val ' + (type === 'income' ? 'income' : (type === 'expense' ? 'expense' : ''));

    if (items.length === 0) {
      listEl.innerHTML = '<div class="kpi-empty-state">Chưa có giao dịch nào trong mục này 🍃</div>';
    } else {
      listEl.innerHTML = items.map(t => {
        const cat = getCategoryById(t.category);
        const payer = t.memberId ? Storage.getMemberById(t.memberId) : null;
        const bene = t.beneficiaryId && t.beneficiaryId !== 'family' ? Storage.getMemberById(t.beneficiaryId) : null;
        const isInc = t.type === 'income';
        
        const payerImg = payer ? (payer.avatarImg || (AVATARS.find(a => a.id === payer.avatarId) || AVATARS[0]).img) : '';
        const beneBadge = bene ? `👤 Cho ${bene.name}` : '👨‍👩‍👧‍👦 Cả nhà';

        return `
          <div class="kpi-drilldown-item">
            <div class="kpi-item-left">
              <div class="kpi-item-icon">${cat ? cat.emoji : '📦'}</div>
              <div class="kpi-item-info">
                <div class="kpi-item-title">${cat ? cat.label : 'Khác'}${t.note ? ' · ' + t.note : ''}</div>
                <div class="kpi-item-sub">
                  <span>${formatDate(t.date)}</span>
                  ${payer ? `<span class="kpi-item-badge"><img src="${payerImg}" style="width:12px;height:12px;border-radius:50%;vertical-align:middle;"> ${payer.name}</span>` : ''}
                  <span class="kpi-item-badge">${beneBadge}</span>
                </div>
              </div>
            </div>
            <div class="kpi-item-amount ${isInc ? 'income' : 'expense'}">${isInc ? '+' : '-'}${formatCurrency(t.amount)}</div>
          </div>
        `;
      }).join('');
    }

    modal.classList.add('active');
  },

  openCategoryDrilldownModal(categoryId) {
    const modal = document.getElementById('kpiDrilldownModal');
    if (!modal) return;
    const cat = getCategoryById(categoryId) || { emoji: '📦', label: categoryId };
    const { year, month } = this.budgetMonth;
    let transactions = Storage.getByMonth(year, month);
    transactions = this.filterTransactionsByScope(transactions, this.budgetScope);

    const filtered = transactions.filter(t => t.type === 'expense' && t.category === categoryId);
    filtered.sort((a, b) => new Date(b.date) - new Date(a.date));

    const total = filtered.reduce((s, t) => s + t.amount, 0);
    const budgetAmount = Storage.getCategoryBudget(categoryId);

    document.getElementById('kpiDrilldownTitle').textContent = `${cat.emoji} Chi tiêu: ${cat.label}`;
    let sub = `Tháng ${month}/${year}`;
    if (this.budgetScope === 'family') sub += ' · 👨‍👩‍👧‍👦 Cả nhà';
    else if (this.budgetScope !== 'all') {
      const mem = Storage.getMemberById(this.budgetScope);
      if (mem) sub += ` · 👤 Cho ${mem.name}`;
    }
    if (budgetAmount > 0) sub += ` · Hạn mức: ${formatCurrency(budgetAmount)}`;
    document.getElementById('kpiDrilldownSubtitle').textContent = sub;
    document.getElementById('kpiTotalVal').textContent = formatCurrency(total);
    document.getElementById('kpiTotalVal').className = 'kpi-total-val expense';

    const listEl = document.getElementById('kpiDrilldownList');
    if (listEl) {
      if (filtered.length === 0) {
        listEl.innerHTML = `<div class="kpi-empty-state">Chưa có khoản chi nào cho mục ${cat.label} trong tháng ${month}/${year} 🍃</div>`;
      } else {
        listEl.innerHTML = filtered.map(t => {
          const payer = t.memberId ? Storage.getMemberById(t.memberId) : null;
          const bene = t.beneficiaryId && t.beneficiaryId !== 'family' ? Storage.getMemberById(t.beneficiaryId) : null;
          const payerImg = payer ? (payer.avatarImg || (AVATARS.find(a => a.id === payer.avatarId) || AVATARS[0]).img) : '';
          const beneBadge = bene ? `👤 Cho ${bene.name}` : '👨‍👩‍👧‍👦 Cả nhà';

          return `
            <div class="kpi-drilldown-item">
              <div class="kpi-item-left">
                <div class="kpi-item-icon">${cat.emoji}</div>
                <div class="kpi-item-info">
                  <div class="kpi-item-title">${t.note ? t.note : cat.label}</div>
                  <div class="kpi-item-sub">
                    <span>${formatDate(t.date)}</span>
                    ${payer ? `<span class="kpi-item-badge"><img src="${payerImg}" style="width:12px;height:12px;border-radius:50%;vertical-align:middle;"> ${payer.name}</span>` : ''}
                    <span class="kpi-item-badge">${beneBadge}</span>
                  </div>
                </div>
              </div>
              <div class="kpi-item-amount expense">-${formatCurrency(t.amount)}</div>
            </div>
          `;
        }).join('');
      }
    }

    modal.classList.add('active');
  },

  closeKPIModal() {
    document.getElementById('kpiDrilldownModal')?.classList.remove('active');
  },

  // ==================== BUDGET PLANNING & DAILY BURN RATE ====================
  renderBudgetSummary(transactions, viewedMonth, scope) {
    const { year, month } = viewedMonth || this.budgetMonth;
    const daysRemaining = getDaysRemainingInMonth(year, month);
    const budgets = Storage.getBudgets();
    const expenseCats = getExpenseCategories();
    const savingsGoal = Storage.getSavingsGoal();

    // Remaining days label
    const daysLabel = document.getElementById('budgetRemainingDaysLabel');
    if (daysLabel) {
      if (daysRemaining === 0) {
        daysLabel.textContent = `Tháng ${month}/${year} đã kết thúc`;
      } else {
        const today = new Date();
        const isCurrentMonth = today.getFullYear() === year && (today.getMonth() + 1) === month;
        daysLabel.textContent = isCurrentMonth 
          ? `Hôm nay ${today.getDate()}/${month} · Còn ${daysRemaining} ngày`
          : `Còn ${daysRemaining} ngày trong tháng`;
      }
    }

    // Expense breakdown by category & Income calculation
    const catSpentMap = {};
    let totalExpense = 0;
    let totalIncome = 0;
    expenseCats.forEach(c => { catSpentMap[c.id] = 0; });
    transactions.forEach(t => {
      if (t.type === 'expense') {
        catSpentMap[t.category] = (catSpentMap[t.category] || 0) + t.amount;
        totalExpense += t.amount;
      } else if (t.type === 'income') {
        totalIncome += t.amount;
      }
    });

    // Total budget calculation
    let totalBudget = 0;
    expenseCats.forEach(c => {
      totalBudget += (budgets[c.id] || 0);
    });

    const overallInfo = calcCategoryBudgetInfo(totalExpense, totalBudget, daysRemaining);

    // Savings Calculation
    const actualSavings = Math.max(0, totalIncome - totalExpense);
    const savingsProgress = savingsGoal > 0 ? Math.round((actualSavings / savingsGoal) * 100) : 0;

    // Render Overview Banner
    const banner = document.getElementById('budgetOverviewBanner');
    if (banner) {
      const isOver = overallInfo.isOver;
      const percentStr = overallInfo.budget > 0 ? `${overallInfo.percent}%` : 'Chưa đặt';
      
      let adviceText = '';
      if (daysRemaining === 0) {
        adviceText = isOver ? `Đã chi vượt ngân sách ${formatCurrency(overallInfo.overAmount)} ⚠️` : `Đã tiết kiệm được ${formatCurrency(overallInfo.remaining)} 🎉`;
      } else if (isOver) {
        adviceText = `Đã chi vượt ngân sách tổng <strong>${formatCurrency(overallInfo.overAmount)}</strong>! Hãy thắt chặt chi tiêu trong ${daysRemaining} ngày tới ⚠️`;
      } else if (overallInfo.dailyAllowance > 0) {
        adviceText = `Hạn mức chi tiêu an toàn: <strong>${formatCurrency(overallInfo.dailyAllowance)}/ngày</strong> trong ${daysRemaining} ngày còn lại 💡`;
      } else {
        adviceText = `Đã sử dụng hết ngân sách tháng! Cần hạn chế phát sinh chi tiêu mới 🛑`;
      }

      banner.innerHTML = `
        <div class="budget-banner-header">
          <div>
            <span class="text-muted" style="font-size: 0.75rem; font-weight: 500;">Đã chi tiêu / Tổng ngân sách</span>
            <div class="budget-banner-stat">
              <span class="budget-banner-amount">${formatCurrency(overallInfo.spent)}</span>
              <span class="text-muted" style="font-size: 0.85rem;">/ ${formatCurrency(overallInfo.budget)}</span>
              <span class="budget-cat-percent ${overallInfo.status}">${percentStr}</span>
            </div>
          </div>
          <div style="text-align: right;">
            <span class="text-muted" style="font-size: 0.75rem; font-weight: 500;">Còn lại</span>
            <div style="font-size: 1.15rem; font-weight: 800; color: ${isOver ? 'var(--expense)' : 'var(--income)'};">
              ${isOver ? '-' + formatCurrency(overallInfo.overAmount) : formatCurrency(overallInfo.remaining)}
            </div>
          </div>
        </div>

        <div class="budget-progress-track">
          <div class="budget-progress-fill ${overallInfo.status}" style="width: ${Math.min(overallInfo.percent, 100)}%;"></div>
        </div>

        <div class="budget-daily-allowance-box">
          <span>${adviceText}</span>
        </div>
      `;
    }

    // Render Categories List
    const catList = document.getElementById('budgetCategoriesList');
    if (catList) {
      catList.innerHTML = expenseCats.map(c => {
        const spent = catSpentMap[c.id] || 0;
        const bAmount = budgets[c.id] || 0;
        const info = calcCategoryBudgetInfo(spent, bAmount, daysRemaining);
        
        let dailyText = '';
        if (daysRemaining === 0) {
          dailyText = info.isOver ? `Vượt ${formatCurrency(info.overAmount)}` : `Dư ${formatCurrency(info.remaining)}`;
        } else if (info.isOver) {
          dailyText = `⚠️ Vượt ${formatCurrency(info.overAmount)}`;
        } else if (info.dailyAllowance > 0) {
          dailyText = `Tối đa: <strong>${formatCurrency(info.dailyAllowance)}/ngày</strong>`;
        } else {
          dailyText = `Đã hết ngân sách`;
        }

        return `
          <div class="budget-cat-item" data-cat="${c.id}" role="button" title="Bấm xem chi tiết các khoản chi ${c.label}">
            <div class="budget-cat-header">
              <div class="budget-cat-info">
                <span>${c.emoji}</span>
                <span>${c.label}</span>
                <span class="budget-cat-hint">🔍</span>
              </div>
              <span class="budget-cat-percent ${info.status}">
                ${bAmount > 0 ? info.percent + '%' : 'Chưa đặt'}
              </span>
            </div>

            <div class="budget-progress-track">
              <div class="budget-progress-fill ${info.status}" style="width: ${Math.min(info.percent, 100)}%;"></div>
            </div>

            <div class="budget-cat-footer">
              <div class="budget-cat-amounts">
                <span>${formatCurrency(spent)}</span>
                <span class="budget-target">/ ${formatCurrency(bAmount)}</span>
              </div>
              <span class="budget-daily-badge ${info.isOver ? 'over' : ''}">${dailyText}</span>
            </div>
          </div>
        `;
      }).join('');

      // Add click listeners to category cards
      catList.querySelectorAll('.budget-cat-item').forEach(item => {
        item.addEventListener('click', () => {
          this.openCategoryDrilldownModal(item.dataset.cat);
        });
      });
    }
  },

  openBudgetModal() {
    const modal = document.getElementById('budgetModal');
    const container = document.getElementById('budgetFormItems');
    if (!modal || !container) return;

    // Set savings goal input
    const savingsInput = document.getElementById('budgetSavingsGoalInput');
    if (savingsInput) {
      savingsInput.value = formatNumberInput(Storage.getSavingsGoal());
      savingsInput.oninput = (e) => {
        e.target.value = formatNumberInput(e.target.value);
      };
    }

    const budgets = Storage.getBudgets();
    const expenseCats = getExpenseCategories();

    container.innerHTML = expenseCats.map(c => {
      const val = budgets[c.id] !== undefined ? budgets[c.id] : (DEFAULT_BUDGETS[c.id] || 0);
      const isCustom = c.id.startsWith('custom_');
      return `
        <div class="budget-form-row" data-cat-row="${c.id}">
          <div class="budget-form-label">
            <span style="font-size: 1.2rem;">${c.emoji}</span>
            <span>${c.label}</span>
            ${isCustom ? `<button type="button" class="btn-delete-cat" data-delete-cat="${c.id}" title="Xóa danh mục này">🗑️</button>` : ''}
          </div>
          <input type="text" class="budget-form-input" data-cat="${c.id}" value="${formatNumberInput(val)}" inputmode="numeric" placeholder="0">
        </div>
      `;
    }).join('');

    // Bind number inputs formatting & total recalculation
    container.querySelectorAll('.budget-form-input').forEach(input => {
      input.addEventListener('input', (e) => {
        e.target.value = formatNumberInput(e.target.value);
        this.updateBudgetModalTotal();
      });
    });

    // Bind custom category deletion
    container.querySelectorAll('[data-delete-cat]').forEach(btn => {
      btn.addEventListener('click', (e) => {
        e.stopPropagation();
        const catId = btn.dataset.deleteCat;
        if (confirm('Xóa danh mục này khỏi danh sách ngân sách?')) {
          Storage.deleteCustomCategory(catId);
          this.populateCategoryFilter();
          this.openBudgetModal();
        }
      });
    });

    // Reset and setup Custom Category form
    this.setupCustomCategoryForm();

    this.updateBudgetModalTotal();
    modal.classList.add('active');
  },

  setupCustomCategoryForm() {
    const showBtn = document.getElementById('showAddCatBtn');
    const form = document.getElementById('addCustomCatForm');
    const cancelBtn = document.getElementById('cancelAddCatBtn');
    const confirmBtn = document.getElementById('confirmAddCatBtn');
    const emojiBtn = document.getElementById('newCatEmojiBtn');
    const emojiPalette = document.getElementById('catEmojiPalette');
    const labelInput = document.getElementById('newCatLabelInput');
    const amountInput = document.getElementById('newCatAmountInput');

    if (!showBtn || !form) return;

    form.style.display = 'none';
    showBtn.style.display = 'block';
    if (emojiPalette) emojiPalette.style.display = 'none';

    let selectedEmoji = '📈';
    if (emojiBtn) emojiBtn.textContent = selectedEmoji;

    // Render preset emojis
    if (emojiPalette) {
      emojiPalette.innerHTML = CUSTOM_CATEGORY_PRESET_EMOJIS.map(em => `
        <div class="cat-emoji-item" data-emoji="${em}">${em}</div>
      `).join('');

      emojiPalette.onclick = (e) => {
        const item = e.target.closest('.cat-emoji-item');
        if (item) {
          selectedEmoji = item.dataset.emoji;
          if (emojiBtn) emojiBtn.textContent = selectedEmoji;
          emojiPalette.style.display = 'none';
        }
      };
    }

    if (emojiBtn) {
      emojiBtn.onclick = () => {
        if (emojiPalette) {
          emojiPalette.style.display = emojiPalette.style.display === 'none' ? 'flex' : 'none';
        }
      };
    }

    if (amountInput) {
      amountInput.oninput = (e) => {
        e.target.value = formatNumberInput(e.target.value);
      };
    }

    showBtn.onclick = () => {
      showBtn.style.display = 'none';
      form.style.display = 'block';
      if (labelInput) { labelInput.value = ''; labelInput.focus(); }
      if (amountInput) amountInput.value = '';
    };

    if (cancelBtn) {
      cancelBtn.onclick = () => {
        form.style.display = 'none';
        showBtn.style.display = 'block';
      };
    }

    if (confirmBtn) {
      confirmBtn.onclick = () => {
        const label = labelInput?.value.trim();
        if (!label) {
          alert('Vui lòng nhập tên danh mục');
          return;
        }
        const amount = parseNumberInput(amountInput?.value);
        const newCat = Storage.addCustomCategory({
          label,
          emoji: selectedEmoji
        });

        // Set its budget amount
        if (amount > 0) {
          const budgets = Storage.getBudgets();
          budgets[newCat.id] = amount;
          Storage.saveBudgets(budgets);
        }

        this.populateCategoryFilter();
        this.openBudgetModal();
        this.showToast(`Đã thêm danh mục ${selectedEmoji} ${label} ✨`);
      };
    }
  },

  closeBudgetModal() {
    document.getElementById('budgetModal')?.classList.remove('active');
  },

  updateBudgetModalTotal() {
    let total = 0;
    document.querySelectorAll('.budget-form-input').forEach(input => {
      total += parseNumberInput(input.value);
    });
    const totalEl = document.getElementById('budgetModalTotal');
    if (totalEl) totalEl.textContent = formatCurrency(total);
  },

  handleBudgetSubmit(e) {
    e.preventDefault();
    // Save savings goal
    const savingsInput = document.getElementById('budgetSavingsGoalInput');
    if (savingsInput) {
      Storage.saveSavingsGoal(parseNumberInput(savingsInput.value));
    }

    // Save category budgets
    const newBudgets = {};
    document.querySelectorAll('.budget-form-input').forEach(input => {
      const catId = input.dataset.cat;
      if (catId) newBudgets[catId] = parseNumberInput(input.value);
    });
    Storage.saveBudgets(newBudgets);

    this.closeBudgetModal();
    this.renderCurrentPage();
    this.showToast('Đã lưu cài đặt ngân sách & mục tiêu ✅');
  },

  applyPresetBudgets(type = 'default') {
    let presets = { ...DEFAULT_BUDGETS };
    if (type === '503020') {
      // 50% Nhu cầu thiết yếu
      presets = {
        food: 7000000,
        house: 3500000,
        bills: 2500000,
        transport: 1500000,
        health: 1000000,
        shopping: 2000000,
        entertainment: 1500000,
        education: 1500000,
        other_expense: 1000000
      };
    }
    document.querySelectorAll('.budget-form-input').forEach(input => {
      const catId = input.dataset.cat;
      if (catId && presets[catId] !== undefined) {
        input.value = formatNumberInput(presets[catId]);
      }
    });
    this.updateBudgetModalTotal();
  },

  // ==================== SAVINGS MODALS & ACTIONS ====================
  openSavingsGoalModal(editId = null) {
    this.editingSavingsGoalId = editId;
    const modal = document.getElementById('savingsGoalModal');
    const form = document.getElementById('savingsGoalForm');
    if (!modal || !form) return;
    form.reset();

    const titleEl = document.getElementById('savingsGoalModalTitle');
    const emojiBtn = document.getElementById('savingsGoalEmojiBtn');
    const emojiPalette = document.getElementById('savingsEmojiPalette');
    const nameInput = document.getElementById('savingsGoalName');
    const targetInput = document.getElementById('savingsTargetAmount');
    const initialInput = document.getElementById('savingsInitialAmount');
    const dateInput = document.getElementById('savingsTargetDate');
    const memberSelect = document.getElementById('savingsMemberSelect');

    // Fill members select
    if (memberSelect) {
      const members = Storage.getMembers();
      memberSelect.innerHTML = `
        <option value="family">👨‍👩‍👧‍👦 Cả nhà</option>
        ${members.map(m => `<option value="${m.id}">👤 ${m.name}</option>`).join('')}
      `;
    }

    let selectedEmoji = '🐷';
    if (emojiBtn) emojiBtn.textContent = selectedEmoji;
    if (emojiPalette) {
      emojiPalette.style.display = 'none';
      emojiPalette.innerHTML = ['🐷','🏦','💰','🪙','💎','🚗','🏠','✈️','👶','🎓','🛡️','💍','🏖️','🎁','🖥️','📱','🩺','📈','🌳','⭐'].map(em => `
        <div class="cat-emoji-item" data-emoji="${em}">${em}</div>
      `).join('');
      emojiPalette.onclick = (e) => {
        const item = e.target.closest('.cat-emoji-item');
        if (item) {
          selectedEmoji = item.dataset.emoji;
          if (emojiBtn) emojiBtn.textContent = selectedEmoji;
          emojiPalette.style.display = 'none';
        }
      };
    }
    if (emojiBtn) {
      emojiBtn.onclick = () => {
        if (emojiPalette) {
          emojiPalette.style.display = emojiPalette.style.display === 'none' ? 'flex' : 'none';
        }
      };
    }

    if (editId) {
      const goal = Storage.getSavingsGoalById(editId);
      if (goal) {
        if (titleEl) titleEl.textContent = 'Sửa Hũ Tiết Kiệm';
        document.getElementById('savingsGoalId').value = editId;
        if (nameInput) nameInput.value = goal.name;
        if (targetInput) targetInput.value = formatNumberInput(goal.targetAmount);
        if (initialInput) initialInput.value = formatNumberInput(goal.currentAmount);
        if (dateInput) dateInput.value = goal.targetDate || '';
        if (memberSelect) memberSelect.value = goal.memberId || 'family';
        selectedEmoji = goal.emoji || '🐷';
        if (emojiBtn) emojiBtn.textContent = selectedEmoji;
      }
    } else {
      if (titleEl) titleEl.textContent = 'Tạo Hũ Tiết Kiệm Mới';
      document.getElementById('savingsGoalId').value = '';
    }

    modal.classList.add('active');
  },

  closeSavingsGoalModal() {
    document.getElementById('savingsGoalModal')?.classList.remove('active');
    this.editingSavingsGoalId = null;
  },

  handleSavingsGoalSubmit(e) {
    e.preventDefault();
    const id = document.getElementById('savingsGoalId')?.value;
    const name = document.getElementById('savingsGoalName')?.value.trim();
    const emoji = document.getElementById('savingsGoalEmojiBtn')?.textContent || '🐷';
    const targetAmount = parseNumberInput(document.getElementById('savingsTargetAmount')?.value);
    const initialAmount = parseNumberInput(document.getElementById('savingsInitialAmount')?.value);
    const targetDate = document.getElementById('savingsTargetDate')?.value;
    const memberId = document.getElementById('savingsMemberSelect')?.value || 'family';

    if (!name) {
      this.showToast('Vui lòng nhập tên mục tiêu', 'error');
      return;
    }
    if (targetAmount <= 0) {
      this.showToast('Vui lòng nhập số tiền mục tiêu hợp lệ', 'error');
      return;
    }

    if (id) {
      Storage.updateSavingsGoal(id, {
        name,
        emoji,
        targetAmount,
        currentAmount: initialAmount,
        targetDate,
        memberId
      });
      this.showToast('Đã cập nhật hũ tiết kiệm ✨');
    } else {
      Storage.addSavingsGoal({
        name,
        emoji,
        targetAmount,
        currentAmount: initialAmount,
        targetDate,
        memberId
      });
      this.showToast(`Đã tạo hũ ${emoji} ${name} 🎉`);
    }

    this.closeSavingsGoalModal();
    this.renderSavingsPage();
  },

  deleteSavingsGoal(id) {
    const goal = Storage.getSavingsGoalById(id);
    if (!goal) return;
    if (confirm(`Bạn có chắc muốn xóa hũ "${goal.name}" và toàn bộ lịch sử nạp/rút của hũ này?`)) {
      Storage.deleteSavingsGoal(id);
      this.renderSavingsPage();
      this.showToast(`Đã xóa hũ ${goal.name}`);
    }
  },

  openSavingsActionModal(goalId, actionType = 'deposit') {
    const goal = Storage.getSavingsGoalById(goalId);
    if (!goal) return;
    const modal = document.getElementById('savingsActionModal');
    const form = document.getElementById('savingsActionForm');
    if (!modal || !form) return;
    form.reset();

    document.getElementById('savingsActionGoalId').value = goalId;
    document.getElementById('savingsActionType').value = actionType;

    const titleEl = document.getElementById('savingsActionTitle');
    const subEl = document.getElementById('savingsActionSubtitle');
    const submitBtn = document.getElementById('savingsActionSubmitBtn');
    const amountInput = document.getElementById('savingsActionAmount');
    const memberSelect = document.getElementById('savingsActionMember');
    const dateInput = document.getElementById('savingsActionDate');

    if (titleEl) titleEl.textContent = actionType === 'deposit' ? '➕ Nạp tiền vào Hũ' : '➖ Rút tiền từ Hũ';
    if (subEl) subEl.textContent = `${goal.emoji} ${goal.name} (Hiện có: ${formatCurrency(goal.currentAmount || 0)})`;
    if (submitBtn) {
      submitBtn.textContent = actionType === 'deposit' ? 'Xác nhận Nạp tiền' : 'Xác nhận Rút tiền';
      submitBtn.style.background = actionType === 'deposit' ? 'var(--income)' : 'var(--expense)';
    }

    if (memberSelect) {
      const members = Storage.getMembers();
      memberSelect.innerHTML = `
        <option value="family">👨‍👩‍👧‍👦 Cả nhà</option>
        ${members.map(m => `<option value="${m.id}">👤 ${m.name}</option>`).join('')}
      `;
      if (goal.memberId && goal.memberId !== 'family') {
        memberSelect.value = goal.memberId;
      }
    }

    if (dateInput) {
      dateInput.value = new Date().toISOString().slice(0, 10);
    }

    modal.classList.add('active');
    setTimeout(() => amountInput?.focus(), 50);
  },

  closeSavingsActionModal() {
    document.getElementById('savingsActionModal')?.classList.remove('active');
  },

  handleSavingsActionSubmit(e) {
    e.preventDefault();
    const goalId = document.getElementById('savingsActionGoalId')?.value;
    const actionType = document.getElementById('savingsActionType')?.value;
    const amount = parseNumberInput(document.getElementById('savingsActionAmount')?.value);
    const memberId = document.getElementById('savingsActionMember')?.value;
    const date = document.getElementById('savingsActionDate')?.value;
    const note = document.getElementById('savingsActionNote')?.value.trim();

    if (!amount || amount <= 0) {
      this.showToast('Vui lòng nhập số tiền hợp lệ', 'error');
      return;
    }

    try {
      Storage.addSavingsAction(goalId, {
        type: actionType,
        amount,
        memberId,
        date,
        note
      });

      this.closeSavingsActionModal();
      this.renderSavingsPage();
      this.showToast(actionType === 'deposit' ? `Đã nạp +${formatCurrency(amount)} vào hũ! 🌱` : `Đã rút -${formatCurrency(amount)} từ hũ! 💸`);
    } catch(err) {
      this.showToast(err.message || 'Lỗi xử lý', 'error');
    }
  },

  openSavingsHistoryModal(goalId) {
    const goal = Storage.getSavingsGoalById(goalId);
    if (!goal) return;
    const modal = document.getElementById('savingsHistoryModal');
    if (!modal) return;

    const titleEl = document.getElementById('savingsHistoryTitle');
    const subEl = document.getElementById('savingsHistorySubtitle');
    const summaryEl = document.getElementById('savingsHistorySummary');
    const listEl = document.getElementById('savingsHistoryList');

    if (titleEl) titleEl.textContent = `Lịch sử: ${goal.emoji} ${goal.name}`;
    if (subEl) subEl.textContent = `Mục tiêu: ${formatCurrency(goal.targetAmount)} · Hiện có: ${formatCurrency(goal.currentAmount || 0)}`;

    const logs = Storage.getSavingsLogs(goalId);

    if (summaryEl) {
      let totalDep = 0;
      let totalWith = 0;
      logs.forEach(l => {
        if (l.type === 'deposit') totalDep += l.amount;
        else if (l.type === 'withdraw') totalWith += l.amount;
      });

      summaryEl.innerHTML = `
        <div style="display:grid; grid-template-columns:1fr 1fr; gap:8px; padding:10px; background:var(--surface); border:1px solid var(--border); border-radius:10px; margin-bottom:10px;">
          <div style="font-size:0.8rem;"><span class="text-muted">Tổng đã nạp:</span> <strong style="color:var(--income);">+${formatCurrency(totalDep)}</strong></div>
          <div style="font-size:0.8rem; text-align:right;"><span class="text-muted">Tổng đã rút:</span> <strong style="color:var(--expense);">-${formatCurrency(totalWith)}</strong></div>
        </div>
      `;
    }

    if (listEl) {
      if (logs.length === 0) {
        listEl.innerHTML = `<div class="empty-state"><p>Chưa có lịch sử giao dịch nào cho hũ này</p></div>`;
      } else {
        listEl.innerHTML = logs.map(l => {
          const isDep = l.type === 'deposit';
          const member = l.memberId && l.memberId !== 'family' ? Storage.getMemberById(l.memberId) : null;
          const memberText = member ? ` · 👤 ${member.name}` : ' · 👨‍👩‍👧‍👦 Cả nhà';

          return `
            <div class="savings-log-item">
              <div class="savings-log-left">
                <span class="savings-log-badge ${l.type}">${isDep ? 'Nạp tiền' : 'Rút tiền'}</span>
                <div class="savings-log-info">
                  <div class="savings-log-title">${l.note || (isDep ? 'Nạp tiền tích lũy' : 'Rút tiền chi tiêu')}</div>
                  <div class="savings-log-sub">${formatDate(l.date)}${memberText}</div>
                </div>
              </div>
              <div class="savings-log-amount ${l.type}">
                ${isDep ? '+' : '-'}${formatCurrency(l.amount)}
              </div>
            </div>
          `;
        }).join('');
      }
    }

    modal.classList.add('active');
  },

  closeSavingsHistoryModal() {
    document.getElementById('savingsHistoryModal')?.classList.remove('active');
  },

  // ==================== BANK WEBHOOK & PENDING TRANSACTIONS ====================
  checkPendingInbox() {
    const pending = Storage.getPendingTransactions();
    const banner = document.getElementById('pendingInboxBanner');
    const countEl = document.getElementById('pendingBannerCount');
    if (banner && countEl) {
      if (pending.length > 0) {
        banner.style.display = 'flex';
        countEl.textContent = `${pending.length} giao dịch ngân hàng mới`;
      } else {
        banner.style.display = 'none';
      }
    }
  },

  openPendingInboxModal() {
    const modal = document.getElementById('pendingInboxModal');
    if (!modal) return;
    this.renderPendingInbox();
    modal.classList.add('active');
  },

  closePendingInboxModal() {
    document.getElementById('pendingInboxModal')?.classList.remove('active');
    this.checkPendingInbox();
  },

  renderPendingInbox() {
    const container = document.getElementById('pendingItemsList');
    if (!container) return;
    const pending = Storage.getPendingTransactions();
    const expenseCats = getExpenseCategories();
    const incomeCats = getIncomeCategories();
    const members = Storage.getMembers();

    if (pending.length === 0) {
      container.innerHTML = `
        <div class="empty-state">
          <span class="empty-icon">🎉</span>
          <p>Hộp thư trống!</p>
          <p class="text-muted">Tất cả giao dịch ngân hàng đã được phân bổ.</p>
        </div>
      `;
      return;
    }

    container.innerHTML = pending.map((p, idx) => {
      const isExp = p.type === 'expense';
      const availableCats = isExp ? expenseCats : incomeCats;
      const catOptions = availableCats.map(c => `
        <option value="${c.id}" ${c.id === p.category ? 'selected' : ''}>${c.emoji} ${c.label}</option>
      `).join('');

      const memberOptions = `
        <option value="family" ${p.memberId === 'family' ? 'selected' : ''}>👨‍👩‍👧‍👦 Cả nhà (Chung)</option>
        ${members.map(m => `
          <option value="${m.id}" ${m.id === p.memberId ? 'selected' : ''}>👤 ${m.name}</option>
        `).join('')}
      `;

      return `
        <div class="pending-card-item" data-pending-id="${p.id}">
          <div class="pending-card-header">
            <div class="pending-bank-badge">
              <span>🏦</span>
              <span>${p.bank || 'Ngân hàng'}</span>
            </div>
            <div class="pending-card-amount ${p.type}">
              ${isExp ? '-' : '+'}${formatCurrency(p.amount)}
            </div>
          </div>

          <div class="pending-raw-text">
            "${p.note || 'Biến động số dư'}"
          </div>

          <div class="pending-controls-row">
            <div>
              <label style="font-size:0.72rem; color:var(--text-secondary); margin-bottom:2px; display:block;">Danh mục phân bổ:</label>
              <select class="pending-select pending-cat-select" data-id="${p.id}">
                ${catOptions}
              </select>
            </div>
            <div>
              <label style="font-size:0.72rem; color:var(--text-secondary); margin-bottom:2px; display:block;">Chi cho ai:</label>
              <select class="pending-select pending-member-select" data-id="${p.id}">
                ${memberOptions}
              </select>
            </div>
          </div>

          <div class="pending-card-actions">
            <button type="button" class="btn-dismiss-pending" data-id="${p.id}">🗑️ Bỏ qua</button>
            <button type="button" class="btn-approve-pending" data-id="${p.id}">✅ Phân bổ & Lưu</button>
          </div>
        </div>
      `;
    }).join('');

    // Bind approve & dismiss buttons
    container.querySelectorAll('.btn-approve-pending').forEach(btn => {
      btn.addEventListener('click', (e) => {
        const id = btn.dataset.id;
        const itemEl = btn.closest('.pending-card-item');
        const catSelect = itemEl.querySelector('.pending-cat-select');
        const memSelect = itemEl.querySelector('.pending-member-select');
        const targetPending = pending.find(x => x.id === id);
        if (targetPending) {
          const selectedCat = catSelect ? catSelect.value : targetPending.category;
          const selectedMember = memSelect ? memSelect.value : targetPending.memberId;
          
          Storage.approvePendingTransaction(id, {
            type: targetPending.type,
            amount: targetPending.amount,
            category: selectedCat,
            note: targetPending.note,
            date: targetPending.date ? targetPending.date.slice(0, 10) : new Date().toISOString().slice(0, 10),
            memberId: selectedMember !== 'family' ? selectedMember : (members[0] ? members[0].id : 'dad'),
            beneficiaryId: selectedMember
          });

          this.showToast(`Đã lưu giao dịch ${formatCurrency(targetPending.amount)} vào sổ! ✅`);
          this.renderPendingInbox();
          this.renderCurrentPage();
        }
      });
    });

    container.querySelectorAll('.btn-dismiss-pending').forEach(btn => {
      btn.addEventListener('click', () => {
        const id = btn.dataset.id;
        Storage.deletePendingTransaction(id);
        this.showToast('Đã bỏ qua giao dịch');
        this.renderPendingInbox();
        this.renderCurrentPage();
      });
    });
  },

  handleTestWebhook() {
    const samples = [
      {
        text: 'TK 19036789 -55.000VND vao 15/08 tai HIGHLANDS COFFEE KEANGNAM',
        title: 'Techcombank Mobile',
        bank: 'Techcombank',
        member: 'dad'
      },
      {
        text: 'TK 001100456 -185.000 VND luc 15/08 thanh toan SHOPEE PAY',
        title: 'VCB Digibank',
        bank: 'Vietcombank',
        member: 'mom'
      },
      {
        text: 'Giao dich thanh cong: -32.000d tai GRAB BIKE vao 15/08',
        title: 'MB Bank',
        bank: 'MB Bank',
        member: 'dad'
      },
      {
        text: 'TK 19036789 +10.000.000VND luc 15/08 luong thang 8',
        title: 'Techcombank Mobile',
        bank: 'Techcombank',
        member: 'dad'
      }
    ];

    const sample = samples[Math.floor(Math.random() * samples.length)];

    if (Storage.isOnline()) {
      this.showToast('Đang gửi thông báo mẫu tới Webhook Google Sheets... ⏳');
      Storage._sheetApiCall('bankNotification', sample).then(res => {
        if (res && res.success && res.transaction) {
          Storage.addLocalPending(res.transaction);
          this.checkPendingInbox();
          this.showToast(`🔔 Đã nhận biến động: ${sample.bank} (${formatCurrency(res.transaction.amount)})!`);
        }
      }).catch(err => {
        console.warn(err);
        this._simulateLocalBankNotification(sample);
      });
    } else {
      this._simulateLocalBankNotification(sample);
    }
  },

  async handleScanGmail() {
    if (!Storage.isOnline()) {
      this.showToast('Vui lòng kết nối Google Sheets trước để quét Gmail! ⚠️');
      return;
    }

    this.showToast('Đang quét hộp thư Gmail tìm email ngân hàng... ⏳');
    try {
      const res = await Storage._sheetApiCall('scanGmail');
      if (res && res.success) {
        await Storage.fetchPendingFromSheets();
        this.checkPendingInbox();
        if (res.addedCount > 0) {
          this.showToast(`🎉 Đã tìm thấy ${res.addedCount} giao dịch ngân hàng mới từ Gmail!`);
          this.openPendingInboxModal();
        } else {
          this.showToast('Không có email biến động số dư mới nào trong Gmail.');
        }
      } else {
        this.showToast(res.error || 'Quét Gmail thất bại');
      }
    } catch(err) {
      console.warn(err);
      this.showToast('Lỗi khi quét Gmail: Hãy chạy hàm setupGmailTrigger trong Apps Script trước');
    }
  },

  // ==================== QUICK BANK ENTRY ====================
  _quickEntryType: 'expense',
  _quickEntryCat: 'food',

  openQuickEntry() {
    const modal = document.getElementById('quickEntryModal');
    if (modal) {
      modal.classList.add('active');
      document.body.style.overflow = 'hidden';
      setTimeout(() => { document.getElementById('quickAmount')?.focus(); }, 200);
    }
  },

  closeQuickEntry() {
    const modal = document.getElementById('quickEntryModal');
    if (modal) {
      modal.classList.remove('active');
      document.body.style.overflow = '';
      document.getElementById('quickAmount').value = '';
      document.getElementById('quickNote').value = '';
    }
  },

  submitQuickEntry() {
    const amountInput = document.getElementById('quickAmount');
    const noteInput = document.getElementById('quickNote');
    const amount = Number(amountInput?.value || 0);
    if (!amount || amount <= 0) {
      this.showToast('⚠️ Vui lòng nhập số tiền');
      amountInput?.focus();
      return;
    }

    const typeBtn = document.querySelector('.quick-type-btn.active');
    const catChip = document.querySelector('.quick-cat-chip.active');
    const type = typeBtn?.dataset?.qtype || 'expense';
    const category = catChip?.dataset?.qcat || 'other_expense';
    const note = noteInput?.value || '';

    const transaction = {
      id: 'quick_' + Date.now() + '_' + Math.floor(Math.random() * 1000),
      type,
      amount,
      category,
      note: note || `Nhập nhanh từ thông báo`,
      date: new Date().toISOString().slice(0, 10),
      memberId: Storage.getAllMembers()[0]?.id || 'dad',
      createdAt: new Date().toISOString(),
      beneficiaryId: ''
    };

    Storage.addTransaction(transaction);
    this.closeQuickEntry();
    this.showToast(`✅ Đã ghi nhận ${type === 'income' ? 'thu' : 'chi'} ${formatCurrency(amount)}!`);
    this.renderCurrentPage();

    // Also sync to Google Sheets if online
    if (Storage.isOnline()) {
      Storage._sheetApiCall('add', { data: JSON.stringify(transaction) }).catch(() => {});
    }
  },

  _simulateLocalBankNotification(sample) {
    const fullText = (sample.title + ' ' + sample.text + ' ' + sample.bank).toLowerCase();
    let amount = 55000;
    const amtMatch = sample.text.match(/(?:[\+\-]|gd:?\s*[\+\-]?|ps:?\s*[\+\-]?)?\s*([\d\.\,]{3,15})\s*(?:vnd|vnđ|đ|d\b)/i);
    if (amtMatch && amtMatch[1]) amount = Number(amtMatch[1].replace(/[\.\,]/g, '')) || 55000;
    const type = sample.text.includes('+') ? 'income' : 'expense';
    let cat = 'food';
    if (/shopee|lazada/i.test(fullText)) cat = 'shopping';
    else if (/grab|be/i.test(fullText)) cat = 'transport';
    else if (/luong|salary/i.test(fullText)) cat = 'salary';

    const item = {
      id: 'pend_' + new Date().getTime(),
      bank: sample.bank,
      type: type,
      amount: amount,
      note: sample.text,
      category: cat,
      memberId: sample.member,
      date: new Date().toISOString(),
      status: 'pending'
    };

    Storage.addLocalPending(item);
    this.checkPendingInbox();
    this.showToast(`🔔 Giả lập nhận biến động: ${sample.bank} (${formatCurrency(amount)})!`);
  },

  // ==================== LOAN MODAL ====================
  openLoanModal(editId = null) {
    this.editingLoanId = editId;
    const modal = document.getElementById('loanModal'); if(!modal) return;
    const form = document.getElementById('loanForm'); if(form) form.reset();
    this.renderLoanTypes();

    if (editId) {
      document.getElementById('loanModalTitle').textContent = 'Sửa khoản vay';
      document.getElementById('editLoanId').value = editId;
      const l = Storage.getLoans().find(x => x.id === editId);
      if (l) {
        this.selectedLoanType = l.loanType;
        document.getElementById('loanName').value = l.name;
        document.getElementById('loanPrincipal').value = formatNumberInput(l.principal);
        document.getElementById('loanRate').value = l.interestRate;
        document.getElementById('loanTerm').value = l.termMonths;
        document.getElementById('loanMonthly').value = formatNumberInput(l.monthlyPayment);
        document.getElementById('loanStartDate').value = l.startDate;
        document.getElementById('loanNote').value = l.note || '';
        this.renderLoanTypes();
      }
    } else {
      document.getElementById('loanModalTitle').textContent = 'Thêm khoản vay';
      document.getElementById('editLoanId').value = '';
      this.selectedLoanType = null;
      document.getElementById('loanStartDate').value = new Date().toISOString().split('T')[0];
    }
    modal.classList.add('active');
  },

  closeLoanModal() {
    document.getElementById('loanModal')?.classList.remove('active');
    this.editingLoanId = null;
  },

  renderLoanTypes() {
    const c = document.getElementById('loanTypeGrid'); if(!c) return;
    c.innerHTML = LOAN_TYPES.map(lt =>
      `<button type="button" class="loan-type-btn ${this.selectedLoanType===lt.id?'active':''}" data-id="${lt.id}"><span class="loan-type-emoji">${lt.emoji}</span><span>${lt.label}</span></button>`
    ).join('');
  },

  handleLoanSubmit(e) {
    e.preventDefault();
    const data = {
      name: document.getElementById('loanName').value,
      loanType: this.selectedLoanType,
      emoji: LOAN_TYPES.find(t => t.id === this.selectedLoanType)?.emoji || '💳',
      principal: parseNumberInput(document.getElementById('loanPrincipal').value),
      interestRate: parseFloat(document.getElementById('loanRate').value),
      termMonths: parseInt(document.getElementById('loanTerm').value),
      monthlyPayment: parseNumberInput(document.getElementById('loanMonthly').value),
      startDate: document.getElementById('loanStartDate').value,
      note: document.getElementById('loanNote').value
    };
    if (!data.name || !data.principal) { this.showToast('Nhập đầy đủ thông tin','error'); return; }
    if (this.editingLoanId) { Storage.updateLoan(this.editingLoanId, data); this.showToast('Đã cập nhật ✅'); }
    else { Storage.addLoan(data); this.showToast('Đã thêm khoản vay ✅'); }
    this.closeLoanModal(); this.renderLoans();
  },

  deleteLoan(id) {
    if (confirm('Xóa khoản vay này?')) { Storage.deleteLoan(id); this.showToast('Đã xóa 🗑️'); this.renderLoans(); }
  },

  // ==================== EXPORT CSV ====================
  exportCSV() {
    const csv = Storage.exportCSV();
    const blob = new Blob([csv], {type:'text/csv;charset=utf-8;'});
    const link = document.createElement('a'); link.href = URL.createObjectURL(blob);
    link.download = `cashflow_${new Date().toISOString().slice(0,10)}.csv`;
    document.body.appendChild(link); link.click(); document.body.removeChild(link);
    this.showToast('Đã xuất CSV 📥');
  },

  // ==================== TOAST ====================
  showToast(msg, type='success') {
    const c = document.getElementById('toastContainer'); if(!c) return;
    const t = document.createElement('div'); t.className = `toast toast-${type}`;
    t.innerHTML = `<span>${msg}</span>`; c.appendChild(t);
    setTimeout(() => { t.classList.add('hide'); setTimeout(() => t.remove(), 300); }, 3000);
  },

  // ==================== RENDER HELPERS ====================
  renderTransactionRow(t) {
    const cat = getCategoryById(t.category);
    const member = t.memberId ? Storage.getMemberById(t.memberId) : null;
    const bene = t.beneficiaryId && t.beneficiaryId !== 'family' ? Storage.getMemberById(t.beneficiaryId) : null;
    const memberBadge = member ? `<span class="family-avatar" style="border-color:${member.color};width:28px;height:28px;display:inline-flex"><img src="${member.avatarImg || (AVATARS.find(a => a.id === member.avatarId) || AVATARS[0]).img}" alt="${member.name}" class="avatar-img"></span>` : '';
    const beneTag = bene ? `<span style="font-size:0.7rem;color:var(--text-secondary);margin-left:4px;">(Cho ${bene.name})</span>` : '';

    return `<tr>
      <td>${formatDate(t.date)}</td>
      <td>${memberBadge}</td>
      <td><span class="badge">${t.type==='income'?'Thu':'Chi'}</span></td>
      <td>${cat?cat.emoji+' '+cat.label:'📦 Khác'} ${beneTag}</td>
      <td>${t.note||'—'}</td>
      <td class="amount-col ${t.type==='income'?'income':'expense'}">${t.type==='income'?'+':'-'}${formatCurrency(t.amount)}</td>
      <td><div class="action-btns"><button class="edit-btn" data-id="${t.id}">✏️</button><button class="delete-btn" data-id="${t.id}">🗑️</button></div></td>
    </tr>`;
  },

  renderRecentItem(t) {
    const cat = getCategoryById(t.category);
    const member = t.memberId ? Storage.getMemberById(t.memberId) : null;
    const bene = t.beneficiaryId && t.beneficiaryId !== 'family' ? Storage.getMemberById(t.beneficiaryId) : null;
    const isInc = t.type === 'income';
    const beneText = bene ? ` · Cho ${bene.name}` : '';

    return `<div class="recent-item">
      <div class="recent-icon">${cat?cat.emoji:'📦'}</div>
      <div class="recent-info">
        <div class="recent-title">${member?`<img src="${member.avatarImg || (AVATARS.find(a => a.id === member.avatarId) || AVATARS[0]).img}" class="avatar-img-sm"> `:''} ${cat?cat.label:'Khác'}${beneText}${t.note?' · '+t.note:''}</div>
        <div class="recent-date">${formatDate(t.date)}</div>
      </div>
      <div class="recent-amount ${isInc?'income':'expense'}">${isInc?'+':'-'}${formatCurrency(t.amount)}</div>
    </div>`;
  },

  // ==================== EVENT BINDINGS ====================
  bindEvents() {
    // Nav
    document.addEventListener('click', (e) => { const n=e.target.closest('[data-page]'); if(n){e.preventDefault();this.navigate(n.dataset.page);} });
    
    // Dashboard scope filter
    document.getElementById('dashboardScopeToggle')?.addEventListener('click', (e) => {
      const btn = e.target.closest('.scope-btn');
      if (btn) {
        this.dashboardScope = btn.dataset.scope;
        this.renderDashboard();
      }
    });

    // KPI Drilldown Cards click
    document.getElementById('kpiIncomeCard')?.addEventListener('click', () => this.openKPIModal('income'));
    document.getElementById('kpiExpenseCard')?.addEventListener('click', () => this.openKPIModal('expense'));
    document.getElementById('kpiBalanceCard')?.addEventListener('click', () => this.openKPIModal('balance'));
    document.getElementById('kpiSavingsCard')?.addEventListener('click', () => this.openKPIModal('savings'));
    document.getElementById('kpiDrilldownClose')?.addEventListener('click', () => this.closeKPIModal());
    document.getElementById('kpiDrilldownModal')?.addEventListener('click', (e) => {
      if (e.target.id === 'kpiDrilldownModal') this.closeKPIModal();
    });

    // Dashboard month nav
    document.getElementById('prevMonth')?.addEventListener('click', () => { this.currentMonth=navigateMonth(this.currentMonth.year,this.currentMonth.month,-1); this.renderDashboard(); });
    document.getElementById('nextMonth')?.addEventListener('click', () => { this.currentMonth=navigateMonth(this.currentMonth.year,this.currentMonth.month,1); this.renderDashboard(); });
    
    // Budget month nav
    document.getElementById('budgetPrevMonth')?.addEventListener('click', () => {
      this.budgetMonth = navigateMonth(this.budgetMonth.year, this.budgetMonth.month, -1);
      this.renderBudgetPage();
    });
    document.getElementById('budgetNextMonth')?.addEventListener('click', () => {
      this.budgetMonth = navigateMonth(this.budgetMonth.year, this.budgetMonth.month, 1);
      this.renderBudgetPage();
    });

    // Analytics month nav
    document.getElementById('analyticsPrevMonth')?.addEventListener('click', () => { this.analyticsMonth=navigateMonth(this.analyticsMonth.year,this.analyticsMonth.month,-1); this.renderAnalytics(); });
    document.getElementById('analyticsNextMonth')?.addEventListener('click', () => { this.analyticsMonth=navigateMonth(this.analyticsMonth.year,this.analyticsMonth.month,1); this.renderAnalytics(); });
    
    // Currency inputs auto-formatter (e.g. 5000 -> 5.000)
    ['amount', 'loanPrincipal', 'loanMonthly', 'savingsTargetAmount', 'savingsInitialAmount', 'savingsActionAmount'].forEach(id => {
      const el = document.getElementById(id);
      if (el) {
        el.addEventListener('input', (e) => {
          e.target.value = formatNumberInput(e.target.value);
        });
      }
    });

    // Savings page & modals
    document.getElementById('openAddSavingsGoalBtn')?.addEventListener('click', () => this.openSavingsGoalModal());
    document.getElementById('savingsGoalModalClose')?.addEventListener('click', () => this.closeSavingsGoalModal());
    document.getElementById('savingsGoalModal')?.addEventListener('click', (e) => { if (e.target.id === 'savingsGoalModal') this.closeSavingsGoalModal(); });
    document.getElementById('savingsGoalForm')?.addEventListener('submit', (e) => this.handleSavingsGoalSubmit(e));
    
    document.getElementById('savingsActionClose')?.addEventListener('click', () => this.closeSavingsActionModal());
    document.getElementById('savingsActionModal')?.addEventListener('click', (e) => { if (e.target.id === 'savingsActionModal') this.closeSavingsActionModal(); });
    document.getElementById('savingsActionForm')?.addEventListener('submit', (e) => this.handleSavingsActionSubmit(e));

    document.getElementById('savingsHistoryClose')?.addEventListener('click', () => this.closeSavingsHistoryModal());
    document.getElementById('savingsHistoryModal')?.addEventListener('click', (e) => { if (e.target.id === 'savingsHistoryModal') this.closeSavingsHistoryModal(); });

    // Pending bank transactions inbox
    document.getElementById('openPendingInboxBtn')?.addEventListener('click', () => this.openPendingInboxModal());
    document.getElementById('pendingInboxClose')?.addEventListener('click', () => this.closePendingInboxModal());
    document.getElementById('pendingInboxModal')?.addEventListener('click', (e) => { if (e.target.id === 'pendingInboxModal') this.closePendingInboxModal(); });
    
    // Webhook settings actions
    document.getElementById('copyWebhookBtn')?.addEventListener('click', () => {
      const input = document.getElementById('webhookUrlInput');
      if (input && input.value) {
        navigator.clipboard.writeText(input.value).then(() => {
          this.showToast('Đã sao chép link Webhook! 📋');
        }).catch(() => {
          input.select();
          document.execCommand('copy');
          this.showToast('Đã sao chép link Webhook! 📋');
        });
      }
    });

    document.getElementById('activateMemberEmailBtn')?.addEventListener('click', () => {
      const emailInput = document.getElementById('memberEmailInput');
      const roleSelect = document.getElementById('memberRoleSelect');
      const email = (emailInput?.value || '').trim();
      const role = roleSelect?.value || 'mom';
      if (!email || !email.includes('@')) {
        this.showToast('⚠️ Vui lòng nhập địa chỉ email hợp lệ');
        emailInput?.focus();
        return;
      }
      const url = Storage.getApiUrl();
      if (!url) {
        this.showToast('⚠️ Vui lòng kết nối Google Sheets trước');
        return;
      }
      // Save member email to local storage
      const memberEmails = JSON.parse(localStorage.getItem('cashflow_member_emails') || '[]');
      const existing = memberEmails.findIndex(m => m.email === email);
      if (existing > -1) {
        memberEmails[existing].role = role;
        memberEmails[existing].activatedAt = new Date().toISOString();
      } else {
        memberEmails.push({ email, role, activatedAt: new Date().toISOString(), status: 'pending' });
      }
      localStorage.setItem('cashflow_member_emails', JSON.stringify(memberEmails));
      // Also save to Google Sheet
      fetch(`${url}?action=registerMemberEmail&email=${encodeURIComponent(email)}&member=${role}`);
      // Open portal for 1-time authorization
      const portalUrl = `${url}?action=portal&member=${role}`;
      const roleLabels = { mom: 'Vợ 🌸', dad: 'Chồng 👔', child: 'Con 🧒', other: 'Thành viên 👤' };
      this.showToast(`✅ Đã đăng ký email ${email} cho ${roleLabels[role] || role}! Đang mở trang xác nhận...`);
      setTimeout(() => {
        window.open(portalUrl, '_blank');
      }, 800);
      emailInput.value = '';
      this.renderMemberEmailList();
    });

    document.getElementById('testWebhookBtn')?.addEventListener('click', () => this.handleTestWebhook());
    document.getElementById('scanGmailBtn')?.addEventListener('click', () => this.handleScanGmail());

    // Quick Bank Entry
    document.getElementById('quickEntryBtn')?.addEventListener('click', () => this.openQuickEntry());
    document.getElementById('quickEntryClose')?.addEventListener('click', () => this.closeQuickEntry());
    document.getElementById('quickEntryModal')?.addEventListener('click', (e) => { if (e.target.id === 'quickEntryModal') this.closeQuickEntry(); });
    document.querySelectorAll('.quick-type-btn').forEach(b => b.addEventListener('click', () => {
      document.querySelectorAll('.quick-type-btn').forEach(x => { x.classList.remove('active'); x.style.borderColor = 'var(--border)'; x.style.background = 'transparent'; x.style.color = 'var(--text-secondary)'; });
      b.classList.add('active'); b.style.borderColor = 'var(--primary)'; b.style.background = 'rgba(231,125,62,0.15)'; b.style.color = 'var(--primary)';
    }));
    document.querySelectorAll('.quick-cat-chip').forEach(c => c.addEventListener('click', () => {
      document.querySelectorAll('.quick-cat-chip').forEach(x => x.classList.remove('active'));
      c.classList.add('active');
    }));
    document.getElementById('quickNote')?.addEventListener('input', (e) => {
      const txt = (e.target.value || '').toLowerCase();
      const map = {
        food: ['ăn','cơm','phở','bún','cafe','cà phê','coffee','trà','sữa','bánh','kem','nhậu','bia','nước','lẩu','pizza','gà','bò','heo','thịt','rau','trái cây','hải sản','chè','milk','tea','highland','starbuck'],
        shopping: ['mua','sắm','shopee','lazada','tiki','sendo','quần','áo','giày','dép','túi','mỹ phẩm','son','kem','đồ dùng','siêu thị','coopmart','winmart','bach hoa','bách hoá'],
        transport: ['grab','be','xăng','xe','gojek','taxi','uber','bus','parking','đỗ xe','vé','gửi xe','đi lại','di chuyển','toll'],
        bills: ['điện','nước','internet','wifi','gas','truyền hình','thuê nhà','phí','bảo hiểm','thuế','phí dịch vụ','cước','viettel','vnpt','fpt'],
        health: ['thuốc','bệnh viện','khám','bs','bác sĩ','y tế','nha','răng','mắt','vitamin','bổ sung','clinic','spa'],
        education: ['học','sách','khóa','trường','lớp','gia sư','course','udemy','skillshare','tiếng anh','đào tạo'],
        entertainment: ['phim','game','nhạc','karaoke','du lịch','netflix','spotify','youtube','vé xem','chơi','billiard','bowling']
      };
      for (const [cat, keywords] of Object.entries(map)) {
        if (keywords.some(k => txt.includes(k))) {
          document.querySelectorAll('.quick-cat-chip').forEach(x => x.classList.remove('active'));
          document.querySelector(`.quick-cat-chip[data-qcat="${cat}"]`)?.classList.add('active');
          return;
        }
      }
    });
    document.getElementById('quickEntrySubmit')?.addEventListener('click', () => this.submitQuickEntry());

    // Transaction modal
    document.getElementById('addTransactionBtn')?.addEventListener('click', (e) => { e.stopPropagation(); this.openModal(); });
    document.getElementById('fabAdd')?.addEventListener('click', () => this.openModal());
    document.getElementById('modalClose')?.addEventListener('click', () => this.closeModal());
    document.getElementById('transactionModal')?.addEventListener('click', (e) => { if(e.target.id==='transactionModal') this.closeModal(); });
    document.getElementById('transactionForm')?.addEventListener('submit', (e) => this.handleSubmit(e));
    document.querySelectorAll('.type-btn').forEach(b => b.addEventListener('click', (e) => { e.preventDefault(); this.selectType(e.currentTarget.dataset.type); }));
    document.getElementById('categoryGrid')?.addEventListener('click', (e) => { const b=e.target.closest('.category-btn'); if(b) this.selectCategory(b.dataset.id); });
    
    // Member selector in form (Ai chi?)
    document.getElementById('memberSelector')?.addEventListener('click', (e) => {
      const b = e.target.closest('.member-btn');
      if (b) { this.selectedMemberId = b.dataset.id; this.renderMemberSelector(); }
    });

    // Beneficiary selector in form (Chi cho ai?)
    document.getElementById('beneficiarySelector')?.addEventListener('click', (e) => {
      const b = e.target.closest('.beneficiary-btn');
      if (b) { this.selectedBeneficiaryId = b.dataset.id; this.renderBeneficiarySelector(); }
    });

    // Loan modal
    document.getElementById('addLoanBtn')?.addEventListener('click', () => this.openLoanModal());
    document.getElementById('loanModalClose')?.addEventListener('click', () => this.closeLoanModal());
    document.getElementById('loanModal')?.addEventListener('click', (e) => { if(e.target.id==='loanModal') this.closeLoanModal(); });
    document.getElementById('loanForm')?.addEventListener('submit', (e) => this.handleLoanSubmit(e));
    document.getElementById('loanTypeGrid')?.addEventListener('click', (e) => { const b=e.target.closest('.loan-type-btn'); if(b){this.selectedLoanType=b.dataset.id;this.renderLoanTypes();} });
    
    // Loan list edit/delete
    document.getElementById('loansList')?.addEventListener('click', (e) => {
      const eb=e.target.closest('.edit-loan-btn'); const db=e.target.closest('.delete-loan-btn');
      if(eb) this.openLoanModal(eb.dataset.id);
      if(db) this.deleteLoan(db.dataset.id);
    });
    
    // Filters
    ['filterMonth', 'filterBeneficiary', 'filterType', 'filterCategory', 'filterSort'].forEach(id => {
      document.getElementById(id)?.addEventListener('change', () => this.renderTransactions());
    });
    
    // Theme
    document.getElementById('themeToggle')?.addEventListener('click', () => this.toggleTheme());
    document.getElementById('themeToggleMobile')?.addEventListener('click', () => this.toggleTheme());
    
    // Export CSV
    document.getElementById('exportBtn')?.addEventListener('click', () => this.exportCSV());
    document.getElementById('exportBtnMobile')?.addEventListener('click', () => this.exportCSV());
    
    // Table edit/delete
    document.getElementById('transactionsBody')?.addEventListener('click', (e) => {
      const eb=e.target.closest('.edit-btn'); const db=e.target.closest('.delete-btn');
      if(eb) this.openModal(eb.dataset.id); if(db) this.deleteTransaction(db.dataset.id);
    });
    
    // Settings
    document.getElementById('syncToggle')?.addEventListener('change', () => this.handleSyncToggle());
    document.getElementById('testConnectionBtn')?.addEventListener('click', () => this.handleTestConnection());
    document.getElementById('syncPullBtn')?.addEventListener('click', () => this.handleSyncPull());
    document.getElementById('syncPushBtn')?.addEventListener('click', () => this.handleSyncPush());
    document.getElementById('copyScriptBtn')?.addEventListener('click', () => this.handleCopyScript());
    document.getElementById('exportJsonBtn')?.addEventListener('click', () => this.handleExportJson());
    document.getElementById('importJsonInput')?.addEventListener('change', (e) => this.handleImportJson(e));
    document.getElementById('clearDataBtn')?.addEventListener('click', () => this.handleClearData());
    
    // Setup
    document.getElementById('setupForm')?.addEventListener('submit', (e) => this.handleSetupSubmit(e));
    document.getElementById('addSetupMember')?.addEventListener('click', () => {
      if (this.setupMemberRows.length < 6) {
        this.setupMemberRows.push({ avatarId: AVATARS[this.setupMemberRows.length % AVATARS.length].id, name: '' });
        this.renderSetupMembers();
      } else { this.showToast('Tối đa 6 thành viên','error'); }
    });
    
    // Delegated: setup member list
    document.getElementById('setupMemberList')?.addEventListener('click', (e) => {
      const picker = e.target.closest('.avatar-picker');
      if (picker) this.showAvatarPicker(parseInt(picker.dataset.index), picker);
      const rm = e.target.closest('.remove-member-btn');
      if (rm) { this.setupMemberRows.splice(parseInt(rm.dataset.index), 1); this.renderSetupMembers(); }
    });
    document.getElementById('setupMemberList')?.addEventListener('input', (e) => {
      if (e.target.classList.contains('setup-member-name')) {
        this.setupMemberRows[parseInt(e.target.dataset.index)].name = e.target.value;
      }
    });
    
    // Member management
    document.getElementById('memberManagementList')?.addEventListener('click', (e) => this.handleMemberManagementClick(e));
    document.getElementById('memberManagementList')?.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' && e.target.classList.contains('member-manage-name-input')) {
        const id = e.target.dataset.id;
        if (e.target.value.trim()) {
          Storage.updateMember(id, { name: e.target.value.trim() });
          this._editingMemberId = null;
          this.renderMemberManagement();
          this.showToast('Đã cập nhật ✅');
          this.renderFamilyAvatars();
        }
      }
    });
    document.getElementById('addMemberBtn')?.addEventListener('click', () => this.handleAddMember());
    
    // Notification settings
    document.getElementById('enableNotifBtn')?.addEventListener('click', () => this.requestNotificationPermission());
    document.getElementById('reminderMemberList')?.addEventListener('change', (e) => {
      if (e.target.classList.contains('reminder-time-input')) {
        this.saveReminderSettings(e.target.dataset.member, 'time', e.target.value);
      }
      if (e.target.classList.contains('reminder-toggle-input')) {
        this.saveReminderSettings(e.target.dataset.member, 'enabled', e.target.checked);
      }
    });
    
    // Avatar Modal
    document.getElementById('avatarModalClose')?.addEventListener('click', () => this.closeAvatarModal());
    document.getElementById('avatarFileInput')?.addEventListener('change', (e) => this.handleAvatarFileSelect(e));
    document.getElementById('avatarPresetsGrid')?.addEventListener('click', (e) => this.handlePresetSelect(e));
    document.getElementById('confirmAvatarBtn')?.addEventListener('click', () => this.confirmAvatarSelection());
    
    // Budget Modal
    document.getElementById('openBudgetModalBtn')?.addEventListener('click', () => this.openBudgetModal());
    document.getElementById('budgetModalClose')?.addEventListener('click', () => this.closeBudgetModal());
    document.getElementById('budgetModal')?.addEventListener('click', (e) => {
      if (e.target.id === 'budgetModal') this.closeBudgetModal();
    });
    document.getElementById('budgetForm')?.addEventListener('submit', (e) => this.handleBudgetSubmit(e));
    document.getElementById('preset503020Btn')?.addEventListener('click', () => this.applyPresetBudgets('503020'));
    document.getElementById('resetBudgetBtn')?.addEventListener('click', () => this.applyPresetBudgets('default'));
    document.getElementById('budgetFormItems')?.addEventListener('input', (e) => {
      if (e.target.classList.contains('budget-form-input')) {
        e.target.value = formatNumberInput(e.target.value);
        this.updateBudgetModalTotal();
      }
    });

    // Escape key
    document.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') {
        this.closeModal();
        this.closeLoanModal();
        this.closeAvatarModal();
        this.closeKPIModal();
        this.closeBudgetModal();
      }
    });
  }
};

document.addEventListener('DOMContentLoaded', () => App.init());
