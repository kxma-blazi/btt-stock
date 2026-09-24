/**
 * ======================================================
 * ระบบจัดการ Stock - BTT
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
    const logSheetName =
      typeof CONFIG !== "undefined" && CONFIG.SHEETS && CONFIG.SHEETS.LOGS
        ? CONFIG.SHEETS.LOGS
        : "Logs";
    let logSheet = ss.getSheetByName(logSheetName);

    if (!logSheet) {
      logSheet = ss.insertSheet(logSheetName);
      logSheet.appendRow([
        "ลำดับ",
        "วัน-เวลา (Timestamp)",
        "ผู้ใช้งาน",
        "กิจกรรม (Action)",
        "รายละเอียด",
      ]);
      logSheet.getRange("A1:E1").setFontWeight("bold").setBackground("#f3f3f3");
    }

    const lastRow = logSheet.getLastRow();
    const newSeq = lastRow <= 1 ? 1 : lastRow;
    const currentUser =
      userEmail || Session.getActiveUser().getEmail() || "Guest/WebUser";
    const timestampNow = formatDateValue(new Date());

    logSheet.appendRow([
      newSeq,
      timestampNow,
      currentUser,
      actionType,
      description,
    ]);
  } catch (e) {
    Logger.log("Error logging activity: " + e.toString());
  }
}

function sendTelegramAlert(message) {
  try {
    if (
      typeof CONFIG === "undefined" ||
      !CONFIG.TELEGRAM_TOKEN ||
      CONFIG.TELEGRAM_TOKEN === "YOUR_TELEGRAM_BOT_TOKEN"
    ) {
      return;
    }
    const token = CONFIG.TELEGRAM_TOKEN;
    const chatId = CONFIG.TELEGRAM_CHAT_ID;
    const url = `https://api.telegram.org/bot${token}/sendMessage`;

    const payload = {
      chat_id: chatId,
      text: message,
      parse_mode: "HTML",
    };

    UrlFetchApp.fetch(url, {
      method: "post",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true,
    });
  } catch (e) {
    Logger.log("Telegram Alert Error: " + e.toString());
  }
}

/**
 * ดึงข้อมูลตามโครงสร้างคอลัมน์ A-P ใน Google Sheets
 */
function getInitialData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const assetSheetName =
      typeof CONFIG !== "undefined" && CONFIG.SHEETS && CONFIG.SHEETS.ASSETS
        ? CONFIG.SHEETS.ASSETS
        : "Asset";
    let assetSheet =
      ss.getSheetByName(assetSheetName) ||
      ss.getSheetByName("Asset") ||
      ss.getSheetByName("Assets");

    if (!assetSheet) {
      assetSheet = ss.insertSheet(assetSheetName);
      assetSheet.appendRow([
        "QR-Code",
        "Time-Stamp",
        "รูปภาพ",
        "รหัส",
        "ชื่อรายการ",
        "หมวดหมู่",
        "สถานที่จัดเก็บ",
        "Stock Max",
        "Stock Min",
        "ราคา/หน่วย",
        "BTT",
        "TKE",
        "TOTAL",
        "สถานะ",
        "ต้องซื้อเพิ่ม",
        "หมายเหตุ",
      ]);
    }

    let assetData = [];
    let categories = new Set();
    let locations = new Set();
    let lowStockItems = [];
    let totalValuation = 0;
    const threshold =
      typeof CONFIG !== "undefined" && CONFIG.LOW_STOCK_THRESHOLD !== undefined
        ? CONFIG.LOW_STOCK_THRESHOLD
        : 2;

    const lastAssetRow = assetSheet.getLastRow();
    if (lastAssetRow > 1) {
      // ดึงข้อมูลตั้งแต่ A2 ถึง P (16 คอลัมน์)
      const rawAssets = assetSheet
        .getRange(2, 1, lastAssetRow - 1, 16)
        .getValues();

      const validAssets = rawAssets.filter(
        (row) => String(row[3]).trim() !== "" || String(row[4]).trim() !== "", // เช็คจาก Col D (รหัส) หรือ Col E (ชื่อ)
      );

      assetData = validAssets.map((row) => {
        const itemCode = String(row[3] || "").trim(); // Col D: รหัส
        const itemName = String(row[4] || "").trim(); // Col E: ชื่อรายการ
        const cat = String(row[5] || "").trim(); // Col F: หมวดหมู่
        const loc = String(row[6] || "").trim(); // Col G: สถานที่จัดเก็บ

        const max = parseInt(row[7]) || 0; // Col H: Stock Max
        const min = parseInt(row[8]) || 0; // Col I: Stock Min
        const price = parseFloat(row[9]) || 0; // Col J: ราคา/หน่วย

        const btt = parseInt(row[10]) || 0; // Col K: BTT
        const tke = parseInt(row[11]) || 0; // Col L: TKE
        const totalQty = btt + tke; // Col M: TOTAL

        totalValuation += totalQty * price;

        if (cat) categories.add(cat);
        if (loc) locations.add(loc);

        let needBuy = row[14]; // Col O: ต้องซื้อเพิ่ม
        if (!needBuy || String(needBuy).trim() === "") {
          needBuy =
            totalQty <= (min || threshold) ? "ต้องซื้อเพิ่ม" : "เพียงพอ";
        }

        const itemObj = [
          itemCode, // 0: รหัส
          itemName, // 1: ชื่อรายการ
          cat, // 2: หมวดหมู่
          loc, // 3: สถานที่จัดเก็บ
          max, // 4: Stock Max
          min, // 5: Stock Min
          price, // 6: ราคา/หน่วย
          btt, // 7: BTT
          tke, // 8: TKE
          totalQty, // 9: TOTAL
          String(row[13] || "พร้อมใช้งาน").trim(), // 10: สถานะ (Col N)
          needBuy, // 11: ต้องซื้อเพิ่ม (Col O)
          String(row[15] || "").trim(), // 12: หมายเหตุ (Col P)
          String(row[2] || "").trim(), // 13: รูปภาพ URL (Col C)
        ];

        if (needBuy === "ต้องซื้อเพิ่ม" || totalQty <= threshold) {
          lowStockItems.push(itemObj);
        }

        return itemObj;
      });
    }

    // ดึงข้อมูลประวัติการยืม
    const borrowSheetName =
      typeof CONFIG !== "undefined" && CONFIG.SHEETS && CONFIG.SHEETS.BORROW
        ? CONFIG.SHEETS.BORROW
        : "BorrowData";
    let borrowSheet =
      ss.getSheetByName(borrowSheetName) || ss.getSheetByName("BorrowData");
    if (!borrowSheet) {
      borrowSheet = ss.insertSheet(borrowSheetName);
      borrowSheet.appendRow([
        "ลำดับ",
        "รหัสอุปกรณ์",
        "ชื่อโครงการ",
        "จำนวน",
        "วันที่เบิก",
        "วันที่คืน",
        "ชื่อผู้เบิก",
        "ตำแหน่งที่อยู่",
      ]);
    }

    let borrowData = [];
    const lastBorrowRow = borrowSheet.getLastRow();
    if (lastBorrowRow > 1) {
      const rawBorrows = borrowSheet
        .getRange(2, 1, lastBorrowRow - 1, 8)
        .getValues();
      const validBorrows = rawBorrows.filter(
        (row) => String(row[1]).trim() !== "",
      );

      borrowData = validBorrows.map((row, idx) => {
        const itemCode = String(row[1] || "").trim();
        const returnDate = formatDateValue(row[5]);
        let borrowStatus =
          returnDate !== "-" && returnDate !== "" ? "คืนแล้ว" : "กำลังยืม";

        return [
          String(row[0] || idx + 1),
          itemCode,
          String(row[2] || "-").trim(),
          parseInt(row[3]) || 1,
          formatDateValue(row[4]),
          returnDate,
          String(row[6] || "").trim(),
          String(row[7] || "-").trim(),
          borrowStatus,
        ];
      });
    }

    // ดึงข้อมูล Logs
    let logData = [];
    const logSheetName =
      typeof CONFIG !== "undefined" && CONFIG.SHEETS && CONFIG.SHEETS.LOGS
        ? CONFIG.SHEETS.LOGS
        : "Logs";
    let logSheet = ss.getSheetByName(logSheetName);
    if (logSheet) {
      const lastLogRow = logSheet.getLastRow();
      if (lastLogRow > 1) {
        logData = logSheet
          .getRange(2, 1, lastLogRow - 1, 5)
          .getValues()
          .reverse();
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
        available: assetData.filter((r) => r[10] === "พร้อมใช้งาน").length,
        borrowed: assetData.filter((r) => r[10] === "ถูกยืม").length,
        repair: assetData.filter(
          (r) => r[10] === "ชำรุด" || r[10] === "ส่งซ่อม",
        ).length,
        totalValuation: totalValuation,
        lowStockCount: lowStockItems.length,
      },
    };
  } catch (error) {
    return { success: false, message: "เกิดข้อผิดพลาด: " + error.toString() };
  }
}

/**
 * บันทึก / แก้ไข ข้อมูล Asset ลงตามคอลัมน์ A-P
 */
function saveAssetData(formObject) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!formObject || !formObject.assetId || !formObject.assetName) {
      return { success: false, message: "กรุณากรอกรหัสและชื่อรายการ" };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName =
      typeof CONFIG !== "undefined" && CONFIG.SHEETS && CONFIG.SHEETS.ASSETS
        ? CONFIG.SHEETS.ASSETS
        : "Asset";
    let sheet =
      ss.getSheetByName(sheetName) ||
      ss.getSheetByName("Asset") ||
      ss.getSheetByName("Assets");
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }

    const existingData = sheet.getDataRange().getValues();
    const newId = String(formObject.assetId).trim();
    const oldId = String(formObject.oldAssetId || "").trim();
    const category = String(formObject.assetCategory || "").trim();
    const location = String(formObject.assetLocation || "").trim();
    const btt = parseInt(formObject.assetBtt) || 0;
    const tke = parseInt(formObject.assetTke) || 0;
    const totalQty = btt + tke;
    const min = parseInt(formObject.assetMin) || 0;
    const threshold =
      typeof CONFIG !== "undefined" && CONFIG.LOW_STOCK_THRESHOLD !== undefined
        ? CONFIG.LOW_STOCK_THRESHOLD
        : 2;
    const needBuy =
      totalQty <= (min || threshold) ? "ต้องซื้อเพิ่ม" : "เพียงพอ";
    const imageUrl = String(formObject.assetImageUrl || "").trim();
    const nowStamp = formatDateValue(new Date());

    let targetRow = -1;
    for (let i = 1; i < existingData.length; i++) {
      const cellId = String(existingData[i][3]).trim().toLowerCase(); // ค้นหาจาก Col D (Index 3)
      if (
        (oldId && cellId === oldId.toLowerCase()) ||
        cellId === newId.toLowerCase()
      ) {
        targetRow = i + 1;
        break;
      }
    }

    if (targetRow > 0) {
      // อัปเดตแถวเดิม คอลัมน์ A ถึง P
      sheet.getRange(targetRow, 1, 1, 16).setValues([
        [
          existingData[targetRow - 1][0] || "", // Col A: QR-Code เดิม
          nowStamp, // Col B: Time-Stamp อัปเดตล่าสุด
          imageUrl, // Col C: รูปภาพ URL
          newId, // Col D: รหัส
          String(formObject.assetName).trim(), // Col E: ชื่อรายการ
          category, // Col F: หมวดหมู่
          location, // Col G: สถานที่จัดเก็บ
          parseInt(formObject.assetMax) || 0, // Col H: Stock Max
          min, // Col I: Stock Min
          parseFloat(formObject.assetPrice) || 0, // Col J: ราคา/หน่วย
          btt, // Col K: BTT
          tke, // Col L: TKE
          `=SUM(K${targetRow}:L${targetRow})`, // Col M: TOTAL (สูตรซัม K+L)
          String(formObject.assetStatus || "พร้อมใช้งาน").trim(), // Col N: สถานะ
          needBuy, // Col O: ต้องซื้อเพิ่ม
          String(formObject.assetRemark || "").trim(), // Col P: หมายเหตุ
        ],
      ]);

      appendIfNew(
        ss,
        CONFIG.SHEETS ? CONFIG.SHEETS.CATEGORIES : "Categories",
        category,
      );
      appendIfNew(
        ss,
        CONFIG.SHEETS ? CONFIG.SHEETS.LOCATIONS : "Locations",
        location,
      );

      logActivity(
        "",
        "UPDATE_ASSET",
        `อัปเดตข้อมูลอุปกรณ์ ${newId} - ${formObject.assetName}`,
      );

      return { success: true, message: "อัปเดตรายการสำเร็จ" };
    } else {
      // เพิ่มรายการใหม่
      const nextRow = sheet.getLastRow() + 1;
      sheet.appendRow([
        "", // Col A: QR-Code (เว้นว่างไว้)
        nowStamp, // Col B: Time-Stamp บันทึกใหม่
        imageUrl, // Col C: รูปภาพ URL
        newId, // Col D: รหัส
        String(formObject.assetName).trim(), // Col E: ชื่อรายการ
        category, // Col F: หมวดหมู่
        location, // Col G: สถานที่จัดเก็บ
        parseInt(formObject.assetMax) || 0, // Col H: Stock Max
        min, // Col I: Stock Min
        parseFloat(formObject.assetPrice) || 0, // Col J: ราคา/หน่วย
        btt, // Col K: BTT
        tke, // Col L: TKE
        `=SUM(K${nextRow}:L${nextRow})`, // Col M: TOTAL (สูตรซัม K+L)
        String(formObject.assetStatus || "พร้อมใช้งาน").trim(), // Col N: สถานะ
        needBuy, // Col O: ต้องซื้อเพิ่ม
        String(formObject.assetRemark || "").trim(), // Col P: หมายเหตุ
      ]);

      appendIfNew(
        ss,
        CONFIG.SHEETS ? CONFIG.SHEETS.CATEGORIES : "Categories",
        category,
      );
      appendIfNew(
        ss,
        CONFIG.SHEETS ? CONFIG.SHEETS.LOCATIONS : "Locations",
        location,
      );

      logActivity(
        "",
        "CREATE_ASSET",
        `เพิ่มอุปกรณ์ใหม่ ${newId} - ${formObject.assetName}`,
      );

      return { success: true, message: "บันทึกรายการสำเร็จ" };
    }
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * บันทึกการยืมอุปกรณ์
 */
function saveBorrowData(formObject) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    if (!formObject || !formObject.itemCode || !formObject.borrowerName) {
      return {
        success: false,
        message: "กรุณาเลือกรหัสอุปกรณ์และระบุชื่อผู้เบิก",
      };
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const borrowSheetName =
      typeof CONFIG !== "undefined" && CONFIG.SHEETS && CONFIG.SHEETS.BORROW
        ? CONFIG.SHEETS.BORROW
        : "BorrowData";
    let borrowSheet =
      ss.getSheetByName(borrowSheetName) || ss.getSheetByName("BorrowData");

    const lastRow = borrowSheet.getLastRow();
    const newSeq = lastRow <= 1 ? 1 : lastRow;
    const targetCode = String(formObject.itemCode).trim();
    const borrower = String(formObject.borrowerName).trim();
    const project = String(formObject.projectName || "-").trim();
    const qty = parseInt(formObject.quantity) || 1;
    const location = String(formObject.locationAddress || "-").trim();

    borrowSheet.appendRow([
      newSeq,
      targetCode,
      project,
      qty,
      formatDateValue(new Date()),
      "-",
      borrower,
      location,
    ]);

    const sheetName =
      typeof CONFIG !== "undefined" && CONFIG.SHEETS && CONFIG.SHEETS.ASSETS
        ? CONFIG.SHEETS.ASSETS
        : "Asset";
    let assetSheet =
      ss.getSheetByName(sheetName) ||
      ss.getSheetByName("Asset") ||
      ss.getSheetByName("Assets");
    let itemName = targetCode;

    if (assetSheet) {
      const assetData = assetSheet.getDataRange().getValues();
      const codeLower = targetCode.toLowerCase();
      for (let i = 1; i < assetData.length; i++) {
        if (String(assetData[i][3]).trim().toLowerCase() === codeLower) {
          // อัปเดต Col N (สถานะ) อ้างอิงจาก Col D (รหัส)
          assetSheet.getRange(i + 1, 14).setValue("ถูกยืม");
          itemName = assetData[i][4] || targetCode;
          break;
        }
      }
    }

    logActivity(
      borrower,
      "BORROW_ASSET",
      `เบิกอุปกรณ์ ${targetCode} (${itemName}) จำนวน ${qty} ชิ้น โครงการ: ${project}`,
    );

    sendTelegramAlert(
      `📤 <b>บันทึกการเบิก/ยืมอุปกรณ์</b>\n📦 อุปกรณ์: ${itemName} (${targetCode})\n👤 ผู้เบิก: ${borrower}\n📂 โครงการ: ${project}\n🔢 จำนวน: ${qty}`,
    );

    return { success: true, message: "บันทึกการยืมเรียบร้อยแล้ว" };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * บันทึกการคืนอุปกรณ์
 */
function returnAsset(itemCode, borrowRowIndex) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const sheetName =
      typeof CONFIG !== "undefined" && CONFIG.SHEETS && CONFIG.SHEETS.ASSETS
        ? CONFIG.SHEETS.ASSETS
        : "Asset";
    let assetSheet =
      ss.getSheetByName(sheetName) ||
      ss.getSheetByName("Asset") ||
      ss.getSheetByName("Assets");
    let itemName = itemCode;

    if (assetSheet) {
      const assetData = assetSheet.getDataRange().getValues();
      const codeLower = String(itemCode).trim().toLowerCase();
      for (let i = 1; i < assetData.length; i++) {
        if (String(assetData[i][3]).trim().toLowerCase() === codeLower) {
          // Col D = รหัส
          assetSheet.getRange(i + 1, 14).setValue("พร้อมใช้งาน"); // Col N = สถานะ
          itemName = assetData[i][4] || itemCode;
          break;
        }
      }
    }

    const borrowSheetName =
      typeof CONFIG !== "undefined" && CONFIG.SHEETS && CONFIG.SHEETS.BORROW
        ? CONFIG.SHEETS.BORROW
        : "BorrowData";
    let borrowSheet =
      ss.getSheetByName(borrowSheetName) || ss.getSheetByName("BorrowData");
    if (
      borrowSheet &&
      borrowRowIndex !== undefined &&
      borrowRowIndex !== null
    ) {
      const actualRow = Number(borrowRowIndex) + 2;
      borrowSheet.getRange(actualRow, 6).setValue(formatDateValue(new Date()));
    }

    logActivity(
      "",
      "RETURN_ASSET",
      `คืนอุปกรณ์ ${itemCode} (${itemName}) เรียบร้อยแล้ว`,
    );

    sendTelegramAlert(
      `📥 <b>บันทึกการคืนอุปกรณ์</b>\n📦 อุปกรณ์: ${itemName} (${itemCode})\n✅ สถานะ: คืนเรียบร้อยแล้ว`,
    );

    return { success: true, message: "บันทึกการคืนอุปกรณ์สำเร็จ" };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

/**
 * ลบรายการ Asset
 */
function deleteAssetData(itemCode) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName =
      typeof CONFIG !== "undefined" && CONFIG.SHEETS && CONFIG.SHEETS.ASSETS
        ? CONFIG.SHEETS.ASSETS
        : "Asset";
    let sheet =
      ss.getSheetByName(sheetName) ||
      ss.getSheetByName("Asset") ||
      ss.getSheetByName("Assets");

    if (!sheet) return { success: false, message: "ไม่พบตารางข้อมูล" };

    const data = sheet.getDataRange().getValues();
    const targetCode = String(itemCode).trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      if (String(data[i][3]).trim().toLowerCase() === targetCode) {
        // ค้นหาที่ Col D (รหัส)
        const itemName = data[i][4];
        sheet.deleteRow(i + 1);

        logActivity(
          "",
          "DELETE_ASSET",
          `ลบอุปกรณ์รหัส ${itemCode} (${itemName}) ออกจากระบบ`,
        );

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
  const data = sheet
    .getDataRange()
    .getValues()
    .flat()
    .map((v) => String(v).trim().toLowerCase());
  if (!data.includes(value.toLowerCase())) {
    sheet.appendRow([value]);
  }
}

function formatDateValue(val) {
  if (!val || val === "" || val === "-") return "-";
  try {
    var d = val instanceof Date ? val : new Date(val);
    if (isNaN(d.getTime())) return String(val).trim();

    var day = String(d.getDate()).padStart(2, "0");
    var month = String(d.getMonth() + 1).padStart(2, "0");
    var year = d.getFullYear();
    var hours = String(d.getHours()).padStart(2, "0");
    var minutes = String(d.getMinutes()).padStart(2, "0");
    var seconds = String(d.getSeconds()).padStart(2, "0");

    return `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
  } catch (e) {
    return "-";
  }
}
