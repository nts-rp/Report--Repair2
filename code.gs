function doGet(e) {
  if (e.parameter.action === 'read') {
    try {
      // ກວດສອບຊື່ຜູ້ໃຊ້ຖ້າມີການສົ່ງມາ ເພື່ອບັງຄັບອອກຈາກລະບົບຖ້າຊື່ຖືກລຶບ
      if (e.parameter.user_name) {
        validateLogin(e.parameter.user_name);
      }
      const data = getData();
      return ContentService.createTextOutput(JSON.stringify({ status: 'success', data: data }))
        .setMimeType(ContentService.MimeType.JSON);
    } catch (error) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
        .setMimeType(ContentService.MimeType.JSON);
    }
  }
  
  // ກວດສອບວ່າຕ້ອງການເປີດໜ້າໃດ (Default ແມ່ນ index)
  let page = e.parameter.page === 'repair' ? 'repair' : 'index';
  let title = page === 'repair' ? 'ລະບົບຊ່າງຊ້ອມແປງ' : 'ລະບົບລາຍງານບັນຫາ';

  return HtmlService.createHtmlOutputFromFile(page)
    .setTitle(title)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ຮັບການສົ່ງຂໍ້ມູນຈາກພາຍນອກຜ່ານ POST
function doPost(e) {
  try {
    const params = JSON.parse(e.postData.contents);
    let message;
    
    if (params.action === 'save') {
      message = saveReport(params.data);
    } else if (params.action === 'update') {
      message = updateRepair(params.id, params.repairer, params.statusAction);
    } else if (params.action === 'edit') {
      message = editReport(params.id, params.issue, params.location, params.reporterName);
    } else if (params.action === 'delete') {
      message = deleteReport(params.id, params.reporterName);
    } else if (params.action === 'login') {
      message = validateLogin(params.name);
    } else {
      throw new Error("Invalid action.");
    }
    return ContentService.createTextOutput(JSON.stringify({ status: 'success', message: message }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ຟັງຊັນກວດສອບການເຂົ້າລະບົບຈາກລາຍຊື່ໃນ Sheet "header"
function validateLogin(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("SheetName");
  if (!sheet) throw new Error("ບໍ່ພົບຂໍ້ມູນລາຍຊື່ໃນລະບົບ (Sheet 'SheetName' ບໍ່ມີຢູ່).");
  
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) throw new Error("ບໍ່ມີລາຍຊື່ໃນລະບົບ.");
  
  const data = sheet.getRange(2, 2, lastRow - 1, 1).getValues();
  const inputName = name.trim().toLowerCase();
  
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] && data[i][0].toString().trim().toLowerCase() === inputName) {
      return "ເຂົ້າສູ່ລະບົບສຳເລັດ!";
    }
  }
  throw new Error("ຊື່ ແລະ ນາມສະກຸນ ບໍ່ຖືກຕ້ອງ ຫຼື ບໍ່ມີໃນລະບົບ.");
}

// ຟັງຊັນສຳລັບລຶບລາຍງານໂດຍຜູ້ລາຍງານ
function deleteReport(id, reporterName) {
  try {
    const sheet = initSheet();
    const data = sheet.getDataRange().getValues();
    Logger.log(`Attempting to delete report with ID: ${id}, Reporter: ${reporterName}`);

    for (let i = 1; i < data.length; i++) {
      const rowId = data[i][0];
      const rowReporter = data[i][4];
      const rowStatus = data[i][8];
      Logger.log(`Checking row ${i+1}: ID=${rowId}, Reporter=${rowReporter}, Status=${rowStatus}`);

      if (rowId == id && rowReporter === reporterName && rowStatus === 'ລໍຖ້າຊ້ອມແປງ') {
        sheet.deleteRow(i + 1);
        Logger.log(`Successfully deleted row ${i+1} for ID: ${id}`);
        return "ລຶບລາຍງານສຳເລັດ!";
      } else if (rowId == id) { // Found the ID, but conditions not met
        if (rowReporter !== reporterName) {
          Logger.log(`Deletion failed for ID ${id}: Reporter mismatch. Expected ${reporterName}, got ${rowReporter}`);
          throw new Error("ທ່ານບໍ່ມີສິດລຶບລາຍງານນີ້.");
        } else if (rowStatus !== 'ລໍຖ້າຊ້ອມແປງ') {
          Logger.log(`Deletion failed for ID ${id}: Status mismatch. Expected 'ລໍຖ້າຊ້ອມແປງ', got ${rowStatus}`);
          throw new Error("ບໍ່ສາມາດລຶບໄດ້, ລາຍງານນີ້ຖືກຮັບຊ້ອມແປງແລ້ວ.");
        }
      }
    }
    Logger.log(`Deletion failed: Report with ID ${id} not found or conditions not met.`);
    throw new Error("ບໍ່ພົບລາຍງານທີ່ຕ້ອງການລຶບ.");
  } catch (e) {
    Logger.log(`Error in deleteReport: ${e.message}`);
    throw new Error("ຂໍ້ຜິດພາດໃນການລຶບ: " + e.message);
  }
}

// ຟັງຊັນສຳລັບແກ້ໄຂລາຍງານໂດຍຜູ້ລາຍງານ
function editReport(id, issue, location, reporterName) {
  try {
    const sheet = initSheet();
    const data = sheet.getDataRange().getValues();
    Logger.log(`Attempting to edit report with ID: ${id}, Reporter: ${reporterName}`);

    for (let i = 1; i < data.length; i++) {
      const rowId = data[i][0];
      const rowReporter = data[i][4];
      const rowStatus = data[i][8];
      Logger.log(`Checking row ${i+1}: ID=${rowId}, Reporter=${rowReporter}, Status=${rowStatus}`);

      if (rowId == id && rowReporter === reporterName && rowStatus === 'ລໍຖ້າຊ້ອມແປງ') {
        sheet.getRange(i + 1, 2).setValue(issue); // ອັບເດດລາຍການບັນຫາ
        sheet.getRange(i + 1, 3).setValue(location); // ອັບເດດສະຖານທີ່
        Logger.log(`Successfully edited row ${i+1} for ID: ${id}`);
        return "ແກ້ໄຂລາຍງານສຳເລັດ!";
      } else if (rowId == id) { // Found the ID, but conditions not met
        if (rowReporter !== reporterName) {
          Logger.log(`Edit failed for ID ${id}: Reporter mismatch. Expected ${reporterName}, got ${rowReporter}`);
          throw new Error("ທ່ານບໍ່ມີສິດແກ້ໄຂລາຍງານນີ້.");
        } else if (rowStatus !== 'ລໍຖ້າຊ້ອມແປງ') {
          Logger.log(`Edit failed for ID ${id}: Status mismatch. Expected 'ລໍຖ້າຊ້ອມແປງ', got ${rowStatus}`);
          throw new Error("ບໍ່ສາມາດແກ້ໄຂໄດ້, ລາຍງານນີ້ຖືກຮັບຊ້ອມແປງແລ້ວ.");
        }
      }
    }
    Logger.log(`Edit failed: Report with ID ${id} not found or conditions not met.`);
    throw new Error("ບໍ່ພົບລາຍງານທີ່ຕ້ອງການແກ້ໄຂ.");
  } catch (e) {
    Logger.log(`Error in editReport: ${e.message}`);
    throw new Error("ຂໍ້ຜິດພາດໃນການແກ້ໄຂ: " + e.message);
  }
}

// ຟັງຊັນກວດສອບ ແລະ ສ້າງ Header ຖ້າຍັງບໍ່ມີ
function initSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName("Sheet1"); // ໃຊ້ Sheet1 ສຳລັບຂໍ້ມູນລາຍງານ
  if (!sheet) {
    sheet = ss.insertSheet("Sheet1"); // ສ້າງ Sheet1 ຖ້າບໍ່ມີ
  }
  if (sheet.getLastRow() === 0) {
    const headers = ["ລຳດັບ", "ລາຍການບັນຫາ", "ສະຖານທີ່", "ວັນທີລາຍງານ", "ຜູ້ລາຍງານ", "ວັນທີຮັບຊ້ອມແປງ", "ວັນທີຊ້ອມແປງສຳເລັດ", "ຜູ້ຊ້ອມແປງ", "ສະຖານະ"];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
  }
  return sheet;
}

// ບັນທຶກຂໍ້ມູນໃໝ່ (ຜູ້ລາຍງານ)
function saveReport(data) {
  try {
    const sheet = initSheet();
    const id = new Date().getTime();
    const reportDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");
    
    sheet.appendRow([
      id, 
      data.issue, 
      data.location, 
      reportDate, 
      data.reporter, 
      '', '', '',
      'ລໍຖ້າຊ້ອມແປງ'
    ]);
    return "ບັນທຶກຂໍ້ມູນສຳເລັດ!";
  } catch (e) {
    throw new Error("ບໍ່ສາມາດບັນທຶກຂໍ້ມູນ: " + e.message);
  }
}

// ດຶງຂໍ້ມູນທັງໝົດມາສະແດງ
function getData() {
  try {
    const sheet = initSheet();
    // ປ່ຽນຈາກ getValues() ເປັນ getDisplayValues() ເພື່ອໃຫ້ໄດ້ຂໍ້ຄວາມຕາມທີ່ສະແດງໃນ Sheet
    return sheet.getDataRange().getDisplayValues().slice(1);
  } catch (e) {
    throw new Error("ບໍ່ສາມາດດຶງຂໍ້ມູນໄດ້");
  }
}

// ອັບເດດການຊ້ອມແປງ
function updateRepair(id, repairer, statusAction) {
  try {
  const sheet = initSheet();
  const data = sheet.getDataRange().getValues();
  const dateStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MM/dd/yyyy HH:mm:ss");

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      const currentRepairer = data[i][7];
      const currentStatus = data[i][8];

      if (statusAction === 'accept') {
        if (currentStatus !== 'ລໍຖ້າຊ້ອມແປງ') {
          Logger.log(`Accept failed for ID ${id}: Status is ${currentStatus}, not 'ລໍຖ້າຊ້ອມແປງ'.`);
          throw new Error("ວຽກນີ້ຖືກຮັບໄປແລ້ວ ຫຼື ບໍ່ຢູ່ໃນສະຖານະລໍຖ້າຊ້ອມແປງ.");
        }
        sheet.getRange(i + 1, 6).setValue(dateStr); // ວັນທີຮັບຊ້ອມ
        sheet.getRange(i + 1, 8).setValue(repairer); // ຜູ້ຊ້ອມແປງ
        sheet.getRange(i + 1, 9).setValue("ກຳລັງຊ້ອມແປງ");
        Logger.log(`Accepted report ID ${id} by ${repairer}.`);
        return "ຮັບຊ້ອມແປງສຳເລັດ";
      } else if (statusAction === 'complete') {
        if (currentRepairer !== repairer) {
          Logger.log(`Complete failed for ID ${id}: Repairer mismatch. Expected ${repairer}, got ${currentRepairer}.`);
          throw new Error("ທ່ານບໍ່ມີສິດປິດວຽກທີ່ທ່ານບໍ່ໄດ້ຮັບ.");
        }
        sheet.getRange(i + 1, 7).setValue(dateStr); // ວັນທີສຳເລັດ
        sheet.getRange(i + 1, 9).setValue("ສຳເລັດແລ້ວ");
        Logger.log(`Completed report ID ${id} by ${repairer}.`);
        return "ຊ້ອມແປງສຳເລັດແລ້ວ";
      }
    }
  }
  } catch (e) {
    throw new Error("ຂໍ້ຜິດພາດ: " + e.message);
  }
}
