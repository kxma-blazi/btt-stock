/**
 * ======================================================
 * ระบบจัดการ Stock - BTT (Backend Code)
 * ======================================================
 */

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("ระบบจัดการ Stock - BTT")
    .addMetaTag("viewport", "width=device-width, initial-scale=1")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function logActivity(userEmail, actionType, description) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const logSheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS && CONFIG.SHEETS.LOGS) ? CONFIG.SHEETS.LOGS : "Logs";
    let logSheet = ss.getSheetByName(logSheetName);

    if (!logSheet) {
      logSheet = ss.insertSheet(logSheetName);
      logSheet.appendRow(["ลำดับ", "วัน-เวลา (Timestamp)", "ผู้ใช้งาน", "กิจกรรม (Action)", "รายละเอียด"]);
      logSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#f3f3f3");
    }

    const lastRow = logSheet.getLastRow();
    const newSeq = (lastRow <= 1) ? 1 : lastRow;
    const currentUser = userEmail || Session.getActiveUser().getEmail() || "Guest/WebUser";
    const timestampNow = formatDateValue(new Date());

    logSheet.appendRow([newSeq, timestampNow, currentUser, actionType, description]);
  } catch (e) {
    Logger.log("Error logging activity: " + e.toString());
  }
}

function sendTelegramAlert(message) {
  try {
    if (typeof CONFIG === 'undefined' || !CONFIG.TELEGRAM_TOKEN || CONFIG.TELEGRAM_TOKEN === "YOUR_TELEGRAM_BOT_TOKEN") {
      return;
    }
    const token = CONFIG.TELEGRAM_TOKEN;
    const chatId = CONFIG.TELEGRAM_CHAT_ID;
    const url = `https://api.telegram.org/bot${token}/sendMessage`;

    const payload = {
      chat_id: chatId,
      text: message,
      parse_mode: "HTML"
    };

    UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    });
  } catch (e) {
    Logger.log("Telegram Alert Error: " + e.toString());
  }
}

function getInitialData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const assetSheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS && CONFIG.SHEETS.ASSETS) ? CONFIG.SHEETS.ASSETS : "Asset";
    let assetSheet = ss.getSheetByName(assetSheetName) || ss.getSheetByName("Asset") || ss.getSheetByName("Assets");

    if (!assetSheet) {
      assetSheet = ss.insertSheet(assetSheetName);
      assetSheet.appendRow([
        "QR-Code", "Time-Stamp", "รูปภาพ", "รหัส", "ชื่อรายการ", "หมวดหมู่", 
        "สถานที่จัดเก็บ", "Stock Max", "Stock Min", "ราคา/หน่วย", "BTT", "TKE", "TOTAL", "สถานะ", "ต้องซื้อเพิ่ม", "หมายเหตุ"
      ]);
    }

    let assetData = [];
    let categories = new Set();
    let locations = new Set();
    let lowStockItems = [];
    let totalValuation = 0;
    const threshold = (typeof CONFIG !== 'undefined' && CONFIG.LOW_STOCK_THRESHOLD !== undefined) ? CONFIG.LOW_STOCK_THRESHOLD : 2;

    const lastAssetRow = assetSheet.getLastRow();
    if (lastAssetRow > 1) {
      const rawAssets = assetSheet.getRange(2, 1, lastAssetRow - 1, 16).getValues();
      const validAssets = rawAssets.filter(row => String(row[3]).trim() !== "" || String(row[4]).trim() !== "");

      assetData = validAssets.map(row => {
        const itemCode = String(row[3] || "").trim();
        const itemName = String(row[4] || "").trim();
        const cat = String(row[5] || "").trim();
        const loc = String(row[6] || "").trim();

        const max = parseInt(row[7]) || 0;
        const min = parseInt(row[8]) || 0;
        const price = parseFloat(row[9]) || 0;

        const btt = parseInt(row[10]) || 0;
        const tke = parseInt(row[11]) || 0;
        const totalQty = btt + tke;

        totalValuation += (totalQty * price);

        if (cat) categories.add(cat);
        if (loc) locations.add(loc);

        let needBuy = row[14];
        if (!needBuy || String(needBuy).trim() === "") {
          needBuy = (totalQty <= (min || threshold)) ? "ต้องซื้อเพิ่ม" : "เพียงพอ";
        }

        const itemObj = [
          itemCode,                     // 0: รหัส
          itemName,                     // 1: ชื่อรายการ
          cat,                          // 2: หมวดหมู่
          loc,                          // 3: สถานที่จัดเก็บ
          max,                          // 4: Stock Max
          min,                          // 5: Stock Min
          price,                        // 6: ราคา/หน่วย
          btt,                          // 7: BTT
          tke,                          // 8: TKE
          totalQty,                     // 9: TOTAL
          String(row[13] || "พร้อมใช้งาน").trim(), // 10: สถานะ
          needBuy,                      // 11: ต้องซื้อเพิ่ม
          String(row[15] || "").trim(), // 12: หมายเหตุ
          String(row[2] || "").trim()   // 13: รูปภาพ URL
        ];

        if (needBuy === "ต้องซื้อเพิ่ม" || totalQty <= threshold) {
          lowStockItems.push(itemObj);
        }

        return itemObj;
      });
    }

    const borrowSheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS && CONFIG.SHEETS.BORROW) ? CONFIG.SHEETS.BORROW : "BorrowData";
    let borrowSheet = ss.getSheetByName(borrowSheetName) || ss.getSheetByName("BorrowData");
    if (!borrowSheet) {
      borrowSheet = ss.insertSheet(borrowSheetName);
      borrowSheet.appendRow(["ลำดับ", "รหัสอุปกรณ์", "ชื่อโครงการ", "จำนวน", "วันที่เบิก", "วันที่คืน", "ชื่อผู้เบิก", "ตำแหน่งที่อยู่"]);
    }

    let borrowData = [];
    const lastBorrowRow = borrowSheet.getLastRow();
    if (lastBorrowRow > 1) {
      const rawBorrows = borrowSheet.getRange(2, 1, lastBorrowRow - 1, 8).getValues();
      const validBorrows = rawBorrows.filter(row => String(row[1]).trim() !== "");

      borrowData = validBorrows.map((row, idx) => {
        const itemCode = String(row[1] || "").trim();
        const returnDate = formatDateValue(row[5]);
        let borrowStatus = (returnDate !== "-" && returnDate !== "") ? "คืนแล้ว" : "กำลังยืม";

        return [
          String(row[0] || (idx + 1)),
          itemCode,
          String(row[2] || "-").trim(),
          parseInt(row[3]) || 1,
          formatDateValue(row[4]),
          returnDate,
          String(row[6] || "").trim(),
          String(row[7] || "-").trim(),
          borrowStatus
        ];
      });
    }

    let logData = [];
    const logSheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS && CONFIG.SHEETS.LOGS) ? CONFIG.SHEETS.LOGS : "Logs";
    let logSheet = ss.getSheetByName(logSheetName);
    if (logSheet) {
      const lastLogRow = logSheet.getLastRow();
      if (lastLogRow > 1) {
        logData = logSheet.getRange(2, 1, lastLogRow - 1, 5).getValues().reverse();
      }
    }

    return {
      success: true,
      assets: assetData,
      borrows: borrowData,
      logs: logData,
      lowStockItems: lowStockItems,
      categories: Array.from(categories),
      locations: Array.from(locations),
      summary: {
        total: assetData.length,
        available: assetData.filter(r => r[10] === "พร้อมใช้งาน").length,
        borrowed: assetData.filter(r => r[10] === "ถูกยืม").length,
        repair: assetData.filter(r => r[10] === "ชำรุด" || r[10] === "ส่งซ่อม").length,
        totalValuation: totalValuation,
        lowStockCount: lowStockItems.length
      }
    };
  } catch (error) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.toString() };
  }
}

/**
 * บันทึกหรืออัปเดตข้อมูลอุปกรณ์จากหน้าเว็บ (UI)
 */
function saveAssetData(formObject) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!formObject || !formObject.assetId || !formObject.assetName) {
      return { success: false, message: "กรุณากรอกรหัสและชื่อรายการ" };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS && CONFIG.SHEETS.ASSETS) ? CONFIG.SHEETS.ASSETS : "Asset";
    let sheet = ss.getSheetByName(sheetName) || ss.getSheetByName("Asset") || ss.getSheetByName("Assets");
    if (!sheet) { sheet = ss.insertSheet(sheetName); }

    const existingData = sheet.getDataRange().getValues();
    const newId = String(formObject.assetId).trim();
    const oldId = String(formObject.oldAssetId || "").trim();
    const category = String(formObject.assetCategory || "").trim();
    const location = String(formObject.assetLocation || "").trim();
    const btt = parseInt(formObject.assetBtt) || 0;
    const tke = parseInt(formObject.assetTke) || 0;
    const totalQty = btt + tke;
    const min = parseInt(formObject.assetMin) || 0;
    const threshold = (typeof CONFIG !== 'undefined' && CONFIG.LOW_STOCK_THRESHOLD !== undefined) ? CONFIG.LOW_STOCK_THRESHOLD : 2;
    const needBuy = (totalQty <= (min || threshold)) ? "ต้องซื้อเพิ่ม" : "เพียงพอ";
    const imageUrl = String(formObject.assetImageUrl || "").trim();
    const nowStamp = formatDateValue(new Date());

    let targetRow = -1;
    for (let i = 1; i < existingData.length; i++) {
      const cellId = String(existingData[i][3]).trim().toLowerCase();
      if ((oldId && cellId === oldId.toLowerCase()) || cellId === newId.toLowerCase()) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow > 0) {
      sheet.getRange(targetRow, 1, 1, 16).setValues([[
        existingData[targetRow - 1][0] || "",
        nowStamp,
        imageUrl,
        newId,
        String(formObject.assetName).trim(),
        category,
        location,
        parseInt(formObject.assetMax) || 0,
        min,
        parseFloat(formObject.assetPrice) || 0,
        btt,
        tke,
        `=SUM(K${targetRow}:L${targetRow})`,
        String(formObject.assetStatus || "พร้อมใช้งาน").trim(),
        needBuy,
        String(formObject.assetRemark || "").trim()
      ]]);

      appendIfNew(ss, CONFIG.SHEETS ? CONFIG.SHEETS.CATEGORIES : "Categories", category);
      appendIfNew(ss, CONFIG.SHEETS ? CONFIG.SHEETS.LOCATIONS : "Locations", location);
      logActivity("", "UPDATE_ASSET", `อัปเดตข้อมูลอุปกรณ์ ${newId} - ${formObject.assetName}`);
      return { success: true, message: "อัปเดตรายการสำเร็จ" };
    } else {
      const nextRow = sheet.getLastRow() + 1;
      sheet.appendRow([
        "",
        nowStamp,
        imageUrl,
        newId,
        String(formObject.assetName).trim(),
        category,
        location,
        parseInt(formObject.assetMax) || 0,
        min,
        parseFloat(formObject.assetPrice) || 0,
        btt,
        tke,
        `=SUM(K${nextRow}:L${nextRow})`,
        String(formObject.assetStatus || "พร้อมใช้งาน").trim(),
        needBuy,
        String(formObject.assetRemark || "").trim()
      ]);

      appendIfNew(ss, CONFIG.SHEETS ? CONFIG.SHEETS.CATEGORIES : "Categories", category);
      appendIfNew(ss, CONFIG.SHEETS ? CONFIG.SHEETS.LOCATIONS : "Locations", location);
      logActivity("", "CREATE_ASSET", `เพิ่มอุปกรณ์ใหม่ ${newId} - ${formObject.assetName}`);
      return { success: true, message: "บันทึกรายการสำเร็จ" };
    }
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function saveBatchBorrowData(borrowPayload) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!borrowPayload || !borrowPayload.items || borrowPayload.items.length === 0 || !borrowPayload.borrowerName) {
      return { success: false, message: "ข้อมูลการยืมไม่ครบถ้วน กรุณาเลือกรายการและระบุชื่อผู้เบิก" };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const borrowSheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS && CONFIG.SHEETS.BORROW) ? CONFIG.SHEETS.BORROW : "BorrowData";
    let borrowSheet = ss.getSheetByName(borrowSheetName) || ss.getSheetByName("BorrowData");
    
    const assetSheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS && CONFIG.SHEETS.ASSETS) ? CONFIG.SHEETS.ASSETS : "Asset";
    let assetSheet = ss.getSheetByName(assetSheetName) || ss.getSheetByName("Asset") || ss.getSheetByName("Assets");
    const assetData = assetSheet ? assetSheet.getDataRange().getValues() : [];

    const borrower = String(borrowPayload.borrowerName).trim();
    const project = String(borrowPayload.projectName || "-").trim();
    const location = String(borrowPayload.locationAddress || "-").trim();
    const borrowDateStr = formatDateValue(new Date());

    let successCount = 0;
    let summaryText = [];

    borrowPayload.items.forEach(item => {
      const targetCode = String(item.itemCode).trim();
      const qty = parseInt(item.quantity) || 1;
      const lastRow = borrowSheet.getLastRow();
      const newSeq = (lastRow <= 1) ? 1 : lastRow;

      borrowSheet.appendRow([
        newSeq,
        targetCode,
        project,
        qty,
        borrowDateStr,
        "-",
        borrower,
        location
      ]);

      let itemName = targetCode;
      if (assetSheet) {
        const codeLower = targetCode.toLowerCase();
        for (let i = 1; i < assetData.length; i++) {
          if (String(assetData[i][3]).trim().toLowerCase() === codeLower) {
            assetSheet.getRange(i + 1, 14).setValue("ถูกยืม");
            itemName = assetData[i][4] || targetCode;
            break;
          }
        }
      }

      successCount++;
      summaryText.push(`- ${itemName} (${targetCode}) จำนวน ${qty} ชิ้น`);
    });

    logActivity(borrower, "BATCH_BORROW", `เบิกอุปกรณ์รวม ${successCount} รายการ โครงการ: ${project}`);
    sendTelegramAlert(`📤 <b>บันทึกการเบิก/ยืมอุปกรณ์ (หลายรายการ)</b>\n👤 ผู้เบิก: ${borrower}\n📂 โครงการ: ${project}\n📦 รายการ:\n${summaryText.join('\n')}`);

    return { success: true, message: `บันทึกการยืมสำเร็จ ${successCount} รายการ` };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function returnAsset(itemCode, borrowRowIndex) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const sheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS && CONFIG.SHEETS.ASSETS) ? CONFIG.SHEETS.ASSETS : "Asset";
    let assetSheet = ss.getSheetByName(sheetName) || ss.getSheetByName("Asset") || ss.getSheetByName("Assets");
    let itemName = itemCode;

    if (assetSheet) {
      const assetData = assetSheet.getDataRange().getValues();
      const codeLower = String(itemCode).trim().toLowerCase();
      for (let i = 1; i < assetData.length; i++) {
        if (String(assetData[i][3]).trim().toLowerCase() === codeLower) {
          assetSheet.getRange(i + 1, 14).setValue("พร้อมใช้งาน");
          itemName = assetData[i][4] || itemCode;
          break;
        }
      }
    }

    const borrowSheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS && CONFIG.SHEETS.BORROW) ? CONFIG.SHEETS.BORROW : "BorrowData";
    let borrowSheet = ss.getSheetByName(borrowSheetName) || ss.getSheetByName("BorrowData");
    if (borrowSheet && borrowRowIndex !== undefined && borrowRowIndex !== null) {
      const actualRow = Number(borrowRowIndex) + 2;
      borrowSheet.getRange(actualRow, 6).setValue(formatDateValue(new Date()));
    }

    logActivity("", "RETURN_ASSET", `คืนอุปกรณ์ ${itemCode} (${itemName}) เรียบร้อยแล้ว`);
    sendTelegramAlert(`📥 <b>บันทึกการคืนอุปกรณ์</b>\n📦 อุปกรณ์: ${itemName} (${itemCode})\n✅ สถานะ: คืนเรียบร้อยแล้ว`);

    return { success: true, message: "บันทึกการคืนอุปกรณ์สำเร็จ" };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function deleteAssetData(itemCode) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS && CONFIG.SHEETS.ASSETS) ? CONFIG.SHEETS.ASSETS : "Asset";
    let sheet = ss.getSheetByName(sheetName) || ss.getSheetByName("Asset") || ss.getSheetByName("Assets");

    if (!sheet) return { success: false, message: "ไม่พบตารางข้อมูล" };

    const data = sheet.getDataRange().getValues();
    const targetCode = String(itemCode).trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][3]).trim().toLowerCase() === targetCode) {
        const itemName = data[i][4];
        sheet.deleteRow(i + 1);
        logActivity("", "DELETE_ASSET", `ลบอุปกรณ์รหัส ${itemCode} (${itemName}) ออกจากระบบ`);
        return { success: true, message: `ลบรายการรหัส "${itemCode}" สำเร็จ` };
      }
    }
    return { success: false, message: "ไม่พบรายการที่ต้องการลบ" };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

function appendIfNew(ss, sheetName, value) {
  if (!value) return;
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(["ชื่อ"]);
  }
  const data = sheet.getDataRange().getValues().flat().map(v => String(v).trim().toLowerCase());
  if (!data.includes(value.toLowerCase())) {
    sheet.appendRow([value]);
  }
}

/**
 * ฟังก์ชันจัดรูปแบบวันที่และเวลา เป็น ว/ด/ปปปป ชม:นาที:วินาที
 */
function formatDateValue(val) {
  if (!val || val === "" || val === "-") return "-";
  try {
    var d = (val instanceof Date) ? val : new Date(val);
    if (isNaN(d.getTime())) return String(val).trim();

    var day = String(d.getDate()).padStart(2, '0');
    var month = String(d.getMonth() + 1).padStart(2, '0');
    var year = d.getFullYear();
    var hours = String(d.getHours()).padStart(2, '0');
    var minutes = String(d.getMinutes()).padStart(2, '0');
    var seconds = String(d.getSeconds()).padStart(2, '0');

    return `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
  } catch (e) {
    return "-";
  }
}

// Backward-compat: รองรับ Index.html เดิมที่ยังเรียก saveBorrowData(form)
function saveBorrowData(formObject) {
  return saveBatchBorrowData({
    items: [{ itemCode: formObject.itemCode, quantity: formObject.quantity }],
    borrowerName: formObject.borrowerName,
    projectName: formObject.projectName,
    locationAddress: formObject.locationAddress
  });
}

/**
 * อัปเดตเวลา Time-Stamp อัตโนมัติเมื่อมีการแก้ไขข้อมูลโดยตรงใน Google Sheets
 */
function onEdit(e) {
  const sheet = e.range.getSheet();
  const assetSheetName = (typeof CONFIG !== 'undefined' && CONFIG.SHEETS && CONFIG.SHEETS.ASSETS) ? CONFIG.SHEETS.ASSETS : "Asset";
  if (sheet.getName() !== assetSheetName) return;

  const row = e.range.getRow();
  const col = e.range.getColumn();
  
  // ข้ามแถวหัวตาราง (แถว 1) และไม่ทำงานหากแก้ไขช่อง Time-Stamp (คอลัมน์ 2) โดยตรงเพื่อป้องกันลูป
  if (row === 1 || col === 2) return;

  const stampCell = sheet.getRange(row, 2);
  const formattedTime = formatDateValue(new Date());
  stampCell.setValue(formattedTime);
}


/**
 * ฟังก์ชันสำหรับรับไฟล์รูปภาพแบบ Base64 แล้วเซฟลง Google Drive
 */
function uploadImageToDrive(base64Data, fileName) {
  try {
    const splitBase = base64Data.split(',');
    const type = splitBase[0].split(';')[0].replace('data:', '');
    const byteCharacters = Utilities.base64Decode(splitBase[1]);
    const blob = Utilities.newBlob(byteCharacters, type, fileName);

    // ชื่อโฟลเดอร์สำหรับเก็บรูปใน Google Drive (สร้างให้อัตโนมัติหากยังไม่มี)
    const folderName = "Stock_Asset_Images";
    const folders = DriveApp.getFoldersByName(folderName);
    let folder;
    
    if (folders.hasNext()) {
      folder = folders.next();
    } else {
      folder = DriveApp.createFolder(folderName);
    }

    // เซฟไฟล์ลง Drive และปรับสิทธิ์เป็นสาธารณะ (เปิดดูรูปได้)
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // สร้าง Direct URL สำหรับดึงรูปมาแสดงผลบนเว็บได้ทันที
    const directUrl = "https://drive.google.com/thumbnail?id=" + file.getId() + "&sz=w1000";

    return { success: true, url: directUrl };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}

function uploadImageToDrive(base64Data, fileName) {
  try {
    var splitData = base64Data.split(",");
    var contentType = splitData[0].match(/:(.*?);/)[1];
    var bytes = Utilities.base64Decode(splitData[1]);
    var blob = Utilities.newBlob(bytes, contentType, fileName);

    // สร้างหรือดึงโฟลเดอร์สำหรับเก็บรูปภาพใน Drive
    var folderName = "Stock_Images";
    var folders = DriveApp.getFoldersByName(folderName);
    var folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(folderName);

    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    // ดึง Direct View URL สำหรับนำไปแสดงใน <img>
    var fileId = file.getId();
    var fileUrl = "https://lh3.googleusercontent.com/d/" + fileId;

    return { success: true, url: fileUrl };
  } catch (e) {
    return { success: false, message: e.toString() };
  }
}