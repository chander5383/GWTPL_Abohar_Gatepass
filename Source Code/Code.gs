// =============================
// ✅ FINAL Code.gs (Updated for v5 Fields)
// =============================

const SHEET_ID = "1rt1-OgNmACT4zro2gSZjNDUB83V3hfvDTNMSzEqBFfA";
const SHEET_NAME = "GatePassLog";

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle("GWTPL Gate Pass")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function toYYYYMMDD(date) {
  if (!date) return "";
  try {
    var d = new Date(date);
    if (isNaN(d.getTime())) return "";
    var year = d.getFullYear();
    var month = String(d.getMonth() + 1).padStart(2, '0');
    var day = String(d.getDate()).padStart(2, '0');
    return [year, month, day].join('-');
  } catch (e) {
    return "";
  }
}

function getSheetHelper() {
  var ss;
  try {
    ss = SpreadsheetApp.openById(SHEET_ID);
  } catch(e) {
    // Fallback: Agar ID galat ho toh Active Sheet use karein
    try {
      ss = SpreadsheetApp.getActiveSpreadsheet();
    } catch(err) {
      throw new Error("Spreadsheet ID galat hai ya permission nahi hai.");
    }
  }

  var sheet = ss.getSheetByName(SHEET_NAME);
  
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow([
      "Gate Pass No", // A
      "Date",         // B
      "Consignor",    // C
      "Person",       // D
      "Vehicle",      // E
      "Auth Person",  // F
      "Type",         // G
      "Items JSON",   // H
      "Auth Name1",   // I
      "Auth Desig1",  // J
      "Remarks",      // K
      "Mobile",       // L
      "Site to Site", // M
      "Auth Name2",   // N (New)
      "Auth Desig2",  // O (New)
      "Outward No",   // P (New)
      "Inward No",    // Q (New)
      "Security Date",// R (New)
      "Timestamp"     // S
    ]);
  }
  return sheet;
}

// === YEAR RESET LOGIC IS HERE ===
function getNextGatePassNumber() {
  try {
    const sheet = getSheetHelper();
    const prefix = "GWTPL/ABO";
    const currentYear = new Date().getFullYear(); 
    const lastRow = sheet.getLastRow();
    let nextNum = 1;

    if (lastRow > 1) {
      const lastGP = sheet.getRange(lastRow, 1).getValue(); 
      const parts = String(lastGP).split('/');
      
      // Agar pichla GP usi saal ka hai, toh count badhao
      if (parts.length === 4 && parseInt(parts[2], 10) === currentYear) {
        nextNum = parseInt(parts[3], 10) + 1;
      }
      // Agar saal alag hai (New Year), toh nextNum 1 hi rahega (Reset)
    }
    
    return `${prefix}/${currentYear}/${String(nextNum).padStart(4, "0")}`;
  } catch (e) {
    return "Error: " + e.message;
  }
}

function saveData(data) {
  try {
    const sheet = getSheetHelper();
    const prefix = "GWTPL/ABO";
    const currentYear = new Date().getFullYear();
    const lastRow = sheet.getLastRow();
    let nextNum = 1;

    // Same Year Check Logic for Saving
    if (lastRow > 1) {
      const lastGP = sheet.getRange(lastRow, 1).getValue(); 
      const parts = String(lastGP).split('/');
      if (parts.length === 4 && parseInt(parts[2], 10) === currentYear) {
        nextNum = parseInt(parts[3], 10) + 1;
      }
    }

    const newGP = `${prefix}/${currentYear}/${String(nextNum).padStart(4, "0")}`;
    
    // Updated Row Data with New Fields
    const rowData = [
      newGP,                  
      data.date || "",            
      data.consignor || "",   
      data.person_carrying || "", 
      data.vehicle_no || "",    
      data.auth_person || "",   
      data.type || "",          
      JSON.stringify(data.items || []), 
      data.authName1 || "",     
      data.authDesig1 || "",    
      data.remarks || "",       
      data.person_mobile || "", 
      data.is_site_to_site || false,
      data.authName2 || "",     // New
      data.authDesig2 || "",    // New
      data.outward1 || "",      // New
      data.inward2 || "",       // New
      data.secDate1 || "",      // New
      new Date()                
    ];

    sheet.appendRow(rowData);
    return { success: true, gate_pass_no: newGP };
    
  } catch (err) {
    return { success: false, error: err.message };
  }
}

function fetchRecord(gpNo) {
  try {
    const sheet = getSheetHelper();
    if (!gpNo) return { success: false, error: "Missing GP No" };

    const data = sheet.getDataRange().getValues();
    let found = null;

    // Loop to find GP Number (Column A is Index 0)
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][0]).trim() === gpNo.trim()) {
        found = data[i];
        break;
      }
    }

    if (!found) return { success: false, error: "Record not found" };

    return {
      gate_pass_no: found[0],   
      date: toYYYYMMDD(found[1]), 
      consignor: found[2],    
      person_carrying: found[3], 
      vehicle_no: found[4],   
      authorised_person: found[5], 
      type: found[6],        
      items: found[7],       
      authName1: found[8] || "",    
      authDesig1: found[9] || "",    
      remarks: found[10] || "",    
      person_mobile: found[11] || "", 
      is_site_to_site: found[12],
      authName2: found[13] || "", // New Fetch
      authDesig2: found[14] || "", // New Fetch
      outward1: found[15] || "",    // New Fetch
      inward2: found[16] || "",     // New Fetch
      secDate1: toYYYYMMDD(found[17]), // New Fetch
      success: true,
    };
  } catch (err) {
    return { success: false, error: err.message };
  }
}
