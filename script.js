/**
 * REVISI: Flip -> Bank Jago
 */
const SPREADSHEET_ID = '1gmOhPU60RHrofJZ8ooOhiYcFlndHAENrhsT43D_ageQ'; 

// Nama Sheet Diubah
const TRANSACTION_SHEET_NAMES = ["Bank Jago", "Gopay", "Buku Agen", "QRIS"]; 
const BALANCE_BANK_JAGO_SHEET_NAME = 'Balance_Bank_Jago'; 
const BALANCE_CASH_SHEET_NAME = 'Balance_Cash';

// Index Kolom
const TIMESTAMP_COL_INDEX = 1;
const CHANNEL_COL_INDEX = 2;
const PRODUK_COL_INDEX = 3;
const IDENTITAS_COL_INDEX = 4;
const EWALLET_COL_INDEX = 5;
const STATUS_COL_INDEX = 6;
const PRICE_CHANNEL_COL_INDEX = 7;
const PRICE_SELL_COL_INDEX = 8;
const CASH_RECEIVE_COL_INDEX = 9;
const ONLINE_IN_COL_INDEX = 10;
const DEBT_AMOUNT_COL_INDEX = 11;

function doGet(e) {
  try {
    const action = e.parameter.action;

    if (action === 'getTotalDebt') {
      return jsonResponse({ status: "SUCCESS", total: calculateTotalDebt() });
    }
    
    // Action Bank Jago
    if (action === 'getBankJagoBalance') {
      return jsonResponse({ status: "SUCCESS", balance: calculateBankJagoBalance() });
    }
    
    if (action === 'getCashBalance') {
      return jsonResponse({ status: "SUCCESS", balance: calculateCashBalance() });
    }
    
    if (action === 'getActiveDebts') {
        return jsonResponse({ status: "SUCCESS", data: getActiveDebts(e.parameter.limit || 50) });
    }
    
    if (action === 'searchDebt') {
      return jsonResponse({ status: "SUCCESS", data: searchForDebt(e.parameter.identitas, e.parameter.date) });
    }

    throw new Error('Aksi tidak valid.');
  } catch (error) {
    return jsonResponse({ status: "ERROR", error: error.message });
  }
}

function doPost(e) {
  try {
    let params = e.parameter; 
    if (e.postData && e.postData.type === "application/json") {
      params = JSON.parse(e.postData.contents);
    }
    
    const action = params.action;

    if (action === "submitData") {
      submitTransaction(params);
      return jsonResponse({ status: "SUCCESS", message: "Data tersimpan" });
    }
    
    if (action === "updateDebt") {
        const result = updateDebtEntry(params.debtId, params.mode, params.amount);
        return jsonResponse({ status: "SUCCESS", result: result });
    }
    
    // Action Tambah Saldo Bank Jago
    if (action === "addBankJagoBalance") {
        return jsonResponse(addBankJagoBalance(params.amount));
    }
    
    if (action === "addCashBalance") {
        return jsonResponse(addCashBalance(params.amount));
    }
    
    throw new Error('Aksi POST tidak valid.');
  } catch (error) {
    return jsonResponse({ status: "ERROR", error: error.message });
  }
}

// LOGIKA SALDO BANK JAGO
function calculateBankJagoBalance() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let balance = 0;

    const bSheet = ss.getSheetByName(BALANCE_BANK_JAGO_SHEET_NAME);
    if (bSheet && bSheet.getLastRow() > 1) {
        const vals = bSheet.getRange(2, 2, bSheet.getLastRow() - 1, 1).getValues();
        vals.forEach(r => balance += cleanRupiahAndParse(r[0]));
    }

    const jagoSheet = ss.getSheetByName('Bank Jago');
    if (jagoSheet && jagoSheet.getLastRow() > 1) {
        const modalVals = jagoSheet.getRange(2, PRICE_CHANNEL_COL_INDEX, jagoSheet.getLastRow() - 1, 1).getValues();
        modalVals.forEach(r => balance -= cleanRupiahAndParse(r[0]));
    }
    return balance;
}

function addBankJagoBalance(amount) {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(BALANCE_BANK_JAGO_SHEET_NAME);
    const nominal = cleanRupiahAndParse(amount);
    const ts = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');
    sheet.appendRow([ts, nominal, "Tambah Saldo Bank Jago"]);
    return { status: "SUCCESS", message: "Saldo Bank Jago terupdate" };
}

// Helper & Core Logic Lainnya (Hutang, Cash)
function calculateTotalDebt() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let total = 0;
    TRANSACTION_SHEET_NAMES.forEach(name => {
        const s = ss.getSheetByName(name);
        if (s && s.getLastRow() > 1) {
            const data = s.getRange(2, STATUS_COL_INDEX, s.getLastRow()-1, DEBT_AMOUNT_COL_INDEX - STATUS_COL_INDEX + 1).getValues();
            data.forEach(r => {
                if (String(r[0]).toUpperCase() === "TERHUTANG") total += cleanRupiahAndParse(r[r.length-1]);
            });
        }
    });
    return total;
}

function calculateCashBalance() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let balance = 0;
    const cSheet = ss.getSheetByName(BALANCE_CASH_SHEET_NAME);
    if (cSheet && cSheet.getLastRow() > 1) {
        const v = cSheet.getRange(2, 2, cSheet.getLastRow()-1, 1).getValues();
        v.forEach(r => balance += cleanRupiahAndParse(r[0]));
    }
    TRANSACTION_SHEET_NAMES.forEach(name => {
        const s = ss.getSheetByName(name);
        if (s && s.getLastRow() > 1) {
            const d = s.getRange(2, PRODUK_COL_INDEX, s.getLastRow()-1, CASH_RECEIVE_COL_INDEX - PRODUK_COL_INDEX + 1).getValues();
            d.forEach(r => {
                balance += cleanRupiahAndParse(r[CASH_RECEIVE_COL_INDEX - PRODUK_COL_INDEX]);
                if (String(r[0]).toLowerCase().includes("tarik tunai")) balance -= cleanRupiahAndParse(r[PRICE_CHANNEL_COL_INDEX - PRODUK_COL_INDEX]);
            });
        }
    });
    return balance;
}

function submitTransaction(p) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName(p.sheetName);
  const row = [
    Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss'),
    p.channel, p.produk, p.identitas, p.ewallet || '', p.status,
    cleanRupiahAndParse(p.hargaChannel), cleanRupiahAndParse(p.hargaJual),
    cleanRupiahAndParse(p.cashDiterimaAgen), cleanRupiahAndParse(p.onlineMasuk),
    cleanRupiahAndParse(p.jumlahTerhutang)
  ];
  s.appendRow(row);
}

function cleanRupiahAndParse(raw) {
    if (!raw) return 0;
    if (typeof raw === 'number') return raw;
    let c = String(raw).replace(/[^0-9.-]/g, '');
    return parseFloat(c) || 0;
}

function jsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

function setupSheets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    [BALANCE_BANK_JAGO_SHEET_NAME, BALANCE_CASH_SHEET_NAME].forEach(n => {
        if (!ss.getSheetByName(n)) ss.insertSheet(n).appendRow(['Timestamp', 'Nominal', 'Tipe']);
    });
}
