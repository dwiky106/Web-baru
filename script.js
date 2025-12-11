/**
 * PENTING: Ganti SPREADSHEET_ID_HERE dengan ID Spreadsheet Google Anda.
 * CATATAN: Karena menggunakan getActiveSpreadsheet(), ID ini tidak terlalu penting
 * kecuali Anda menggunakan file Apps Script terpisah.
 */
const SPREADSHEET_ID = '1gmOhPU60RHrofJZ8ooOhiYcFlndHAENrhsT43D_ageQ'; 

// LOG_SHEET_NAME telah dihapus
const TRANSACTION_SHEET_NAMES = ["Flip", "Gopay", "Buku Agen", "QRIS"]; 
const BALANCE_FLIP_SHEET_NAME = 'Balance_Flip'; 
const BALANCE_CASH_SHEET_NAME = 'Balance_Cash';

// Index Kolom (1-based)
const TIMESTAMP_COL_INDEX = 1;    // Kolom A
const CHANNEL_COL_INDEX = 2;      // Kolom B
const PRODUK_COL_INDEX = 3;       // Kolom C (Produk)
const IDENTITAS_COL_INDEX = 4;    // Kolom D
const EWALLET_COL_INDEX = 5;      // Kolom E
const STATUS_COL_INDEX = 6;       // Kolom F (Terhutang/Lunas)
const PRICE_CHANNEL_COL_INDEX = 7;// Kolom G (Harga Channel/Modal)
const PRICE_SELL_COL_INDEX = 8;   // Kolom H (Harga Jual)
const CASH_RECEIVE_COL_INDEX = 9; // Kolom I (Cash Diterima Agen)
const ONLINE_IN_COL_INDEX = 10;   // Kolom J (Online Masuk)
const DEBT_AMOUNT_COL_INDEX = 11; // Kolom K (Jumlah Terhutang)

// =========================================================================
// >>>>> FUNGSI UTAMA (doGet, doPost) <<<<<
// =========================================================================

function doGet(e) {
  // Menangani permintaan OPTIONS preflight CORS secara eksplisit
  if (e.parameter && e.parameter.action === 'options') {
    return jsonResponse({ status: "CORS OK" });
  }

  try {
    const action = e.parameter.action;

    if (action === 'getTotalDebt') {
      const totalDebt = calculateTotalDebt();
      return jsonResponse({ status: "SUCCESS", total: totalDebt });
    }
    
    if (action === 'getFlipBalance') {
      const balance = calculateFlipBalance();
      return jsonResponse({ status: "SUCCESS", balance: balance });
    }
    
    if (action === 'getCashBalance') {
      const balance = calculateCashBalance();
      return jsonResponse({ status: "SUCCESS", balance: balance });
    }
    
    // --- AKSI BARU: Mengambil daftar utang terbaru untuk Slideshow ---
    if (action === 'getRecentDebts') {
        const limit = e.parameter.limit ? parseInt(e.parameter.limit) : 20; // Default 20
        const debts = getRecentDebts(limit);
        return jsonResponse({ status: "SUCCESS", data: debts });
    }
    // -----------------------------------------------------------------
    
    if (action === 'searchDebt') {
      const identitas = e.parameter.identitas;
      const date = e.parameter.date;
      const results = searchForDebt(identitas, date);
      return jsonResponse({ status: "SUCCESS", data: results });
    }

    throw new Error('Aksi tidak valid untuk permintaan GET.');
    
  } catch (error) {
    // Mengganti logError dengan Logger.log
    Logger.log(`Error di doGet: ${error.message} - Stack: ${error.stack}`);
    return jsonResponse({ status: "ERROR", error: error.message });
  }
}

function doPost(e) {
  try {
    // Mendapatkan parameter baik dari URL (e.parameter) atau payload JSON (e.postData.contents)
    let params = e.parameter; 
    if (e.postData && e.postData.type === "application/json") {
      try {
        params = JSON.parse(e.postData.contents);
      } catch(err) {
        throw new Error("Gagal parsing JSON body.");
      }
    }
    
    const action = params.action;

    if (action === "submitData") {
      submitTransaction(params);
      return jsonResponse({ status: "SUCCESS", message: "Data berhasil disimpan" });
    }
    
    if (action === "updateDebt") {
        const mode = params.mode;
        const debtId = params.debtId; 
        const amount = params.amount; 
        
        if (!debtId || !mode) {
              throw new Error("Parameter ID dan mode wajib ada untuk update hutang.");
        }

        const result = updateDebtEntry(debtId, mode, amount);
        return jsonResponse({ status: "SUCCESS", result: result, message: result.message });
    }
    
    if (action === "addFlipBalance") {
        const amount = params.amount;
        const result = addFlipBalance(amount);
        return jsonResponse({ status: "SUCCESS", message: result.message });
    }
    
    if (action === "addCashBalance") {
        const amount = params.amount;
        const result = addCashBalance(amount);
        return jsonResponse({ status: "SUCCESS", message: result.message });
    }
    
    throw new Error('Aksi tidak valid untuk permintaan POST.');

  } catch (error) {
    // Mengganti logError dengan Logger.log
    Logger.log(`Error di doPost: ${error.message} - Stack: ${error.stack}`);
    return jsonResponse({ status: "ERROR", error: error.message });
  }
}

// =========================================================================
// >>>>> FUNGSI SETUP, LOGGING, & HELPER <<<<<
// =========================================================================

/**
 * Menyiapkan sheet saldo jika belum ada.
 */
function setupSheets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetNames = [BALANCE_FLIP_SHEET_NAME, BALANCE_CASH_SHEET_NAME]; 
    const headers = ['Timestamp', 'Nominal', 'Tipe Transaksi'];

    sheetNames.forEach(sheetName => {
        let sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            sheet = ss.insertSheet(sheetName);
            sheet.appendRow(headers);
        } else if (sheet.getLastRow() === 0) {
            // Pastikan header ada jika sheet kosong
            sheet.appendRow(headers);
        }
    });
}

// Fungsi logError telah dihapus

/**
 * Fungsi pembantu untuk mengembalikan respons JSON yang valid.
 */
function jsonResponse(data) {
  const output = ContentService.createTextOutput(JSON.stringify(data));
  output.setMimeType(ContentService.MimeType.JSON);
  return output;
}

/**
 * Membersihkan nilai string dari format Rupiah/angka dan mengembalikan angka murni (number).
 */
function cleanRupiahAndParse(raw) {
    if (typeof raw === 'number') {
        return raw;
    }
    if (typeof raw !== 'string' && typeof raw !== 'number') {
        return 0;
    }

    let cleanedStr = String(raw).trim();
    
    let parts = cleanedStr.split(',');
    if (parts.length > 1) {
        cleanedStr = parts[0].replace(/\./g, '') + '.' + parts[1];
    } else {
        cleanedStr = cleanedStr.replace(/[^0-9.-]/g, '');
        cleanedStr = cleanedStr.replace(/\./g, '');
    }
    
    cleanedStr = cleanedStr.replace(/[^\d.-]/g, '');

    return parseFloat(cleanedStr) || 0;
}


// =========================================================================
// >>>>> FUNGSI CORE LOGIC: SALDO & HUTANG <<<<<
// =========================================================================

// ------------------- LOGIC: TOTAL HUTANG -------------------

/**
 * Menghitung dan mengembalikan total utang kumulatif dari semua sheet transaksi.
 */
function calculateTotalDebt() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let totalCumulativeDebt = 0;

    for (const sheetName of TRANSACTION_SHEET_NAMES) {
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet || sheet.getLastRow() <= 1) continue; 
        
        // Ambil range dari kolom F (Status) sampai K (Jumlah Terhutang)
        const range = sheet.getRange(2, STATUS_COL_INDEX, sheet.getLastRow() - 1, DEBT_AMOUNT_COL_INDEX - STATUS_COL_INDEX + 1);
        const values = range.getValues();

        values.forEach(row => {
            const status = String(row[0] || '').trim().toUpperCase(); 
            const rawDebtAmount = row[row.length - 1]; // Kolom K
            const debtAmount = cleanRupiahAndParse(rawDebtAmount);

            if (status === "TERHUTANG" && debtAmount > 0) {
                totalCumulativeDebt += debtAmount;
            }
        });
    }
    return totalCumulativeDebt;
}

// ------------------- LOGIC: SALDO FLIP -------------------

/**
 * Menghitung dan mengembalikan Saldo Flip saat ini.
 * Saldo = Total Penambahan Manual (Balance_Flip) - Total Modal Channel (Kolom G) dari sheet "Flip".
 */
function calculateFlipBalance() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let balance = 0;

    // --- 1. Ambil Data dari Balance_Flip (Penambahan Saldo Manual) ---
    const balanceSheet = spreadsheet.getSheetByName(BALANCE_FLIP_SHEET_NAME);
    if (balanceSheet && balanceSheet.getLastRow() > 1) {
        // Ambil semua nominal penambahan/saldo awal dari Kolom B (Index 2)
        const balanceValues = balanceSheet.getRange(2, 2, balanceSheet.getLastRow() - 1, 1).getValues();
        balanceValues.forEach(row => {
            const nominal = cleanRupiahAndParse(row[0]);
            balance += nominal;
        });
    }

    // --- 2. Kurangi dengan Transaksi Modal dari Channel 'Flip' saja ---
    const flipSheet = spreadsheet.getSheetByName('Flip');
    if (flipSheet && flipSheet.getLastRow() > 1) {
        
        // Ambil semua Harga Channel (Modal) dari Kolom G di sheet "Flip"
        const priceValues = flipSheet.getRange(2, PRICE_CHANNEL_COL_INDEX, flipSheet.getLastRow() - 1, 1).getValues();
        
        priceValues.forEach(row => {
            const modal = cleanRupiahAndParse(row[0]);
            // Saldo berkurang, jadi dikurangi modal yang keluar (Kolom G)
            balance -= modal;
        });
    }
    
    return balance;
}

/**
 * Menambahkan Saldo Flip secara manual oleh user melalui modal "Tambah Saldo".
 */
function addFlipBalance(amount) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const balanceSheet = spreadsheet.getSheetByName(BALANCE_FLIP_SHEET_NAME);

    if (!balanceSheet) {
        throw new Error(`Sheet '${BALANCE_FLIP_SHEET_NAME}' tidak ditemukan. Mohon buat sheet tersebut (Kolom A: Tanggal, Kolom B: Nominal, Kolom C: Tipe Transaksi).`);
    }
    
    const nominal = cleanRupiahAndParse(amount);

    if (nominal <= 0) {
        throw new Error("Nominal penambahan saldo harus lebih besar dari 0.");
    }
    
    // Menggunakan Utilities.formatDate untuk timestamp yang konsisten
    const timestamp = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    // Kolom A: Tanggal, Kolom B: Nominal, Kolom C: Tipe Transaksi
    balanceSheet.appendRow([timestamp, nominal, "Tambah Saldo"]);
    
    return { 
        status: "SUCCESS", 
        message: `Penambahan saldo sebesar ${nominal.toLocaleString('id-ID')} berhasil dicatat.`
    };
}


// ------------------- LOGIC: CASH DITANGAN -------------------

/**
 * Menghitung dan mengembalikan Saldo Cash Ditangan saat ini.
 * Saldo = Total Penambahan Manual (Balance_Cash)
 * + Total Cash Diterima Agen (Kolom I) dari SEMUA sheet transaksi
 * - Total Modal (Kolom G) yang bertransaksi Tarik Tunai
 */
function calculateCashBalance() {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let balance = 0;

    // --- 1. Ambil Data dari Balance_Cash (Penambahan Saldo Cash Manual) ---
    const cashSheet = spreadsheet.getSheetByName(BALANCE_CASH_SHEET_NAME);
    if (cashSheet && cashSheet.getLastRow() > 1) {
        // Ambil semua nominal penambahan/saldo awal dari Kolom B (Index 2)
        const balanceValues = cashSheet.getRange(2, 2, cashSheet.getLastRow() - 1, 1).getValues();
        balanceValues.forEach(row => {
            const nominal = cleanRupiahAndParse(row[0]);
            balance += nominal;
        });
    }

    // --- 2. Iterasi SEMUA Sheet Transaksi untuk Menambah dan Mengurangi Cash ---
    for (const sheetName of TRANSACTION_SHEET_NAMES) {
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet || sheet.getLastRow() <= 1) continue; 
        
        // Ambil range dari kolom C (Produk) sampai I (Cash Diterima Agen)
        // dan tambahkan Kolom G (Harga Channel)
        const range = sheet.getRange(2, PRODUK_COL_INDEX, sheet.getLastRow() - 1, CASH_RECEIVE_COL_INDEX - PRODUK_COL_INDEX + 1);
        const values = range.getValues();

        values.forEach(row => {
            const produk = String(row[0] || '').trim().toLowerCase(); 
            
            // CASH_RECEIVE_COL_INDEX (Kolom I) = index 6 relatif ke PRODUK_COL_INDEX (Kolom C)
            const cashReceived = cleanRupiahAndParse(row[CASH_RECEIVE_COL_INDEX - PRODUK_COL_INDEX]);
            
            // Kolom G (Harga Channel/Modal) = index 4 relatif ke PRODUK_COL_INDEX (Kolom C)
            const modal = cleanRupiahAndParse(row[PRICE_CHANNEL_COL_INDEX - PRODUK_COL_INDEX]);
            
            // A. Penambahan: Tambahkan semua Cash Diterima Agen (Kolom I)
            if (cashReceived > 0) {
                balance += cashReceived;
            }
            
            // B. Pengurangan: Kurangi modal (Kolom G) hanya jika transaksi adalah 'Tarik Tunai'
            if (produk.includes("tarik tunai") && modal > 0) {
                balance -= modal;
            }
        });
    }
    
    return balance;
}

/**
 * Menambahkan Saldo Cash Ditangan secara manual oleh user.
 */
function addCashBalance(amount) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const balanceSheet = spreadsheet.getSheetByName(BALANCE_CASH_SHEET_NAME);

    if (!balanceSheet) {
        throw new Error(`Sheet '${BALANCE_CASH_SHEET_NAME}' tidak ditemukan. Mohon buat sheet tersebut (Kolom A: Tanggal, Kolom B: Nominal, Kolom C: Tipe Transaksi).`);
    }
    
    const nominal = cleanRupiahAndParse(amount);

    if (nominal <= 0) {
        throw new Error("Nominal penambahan saldo cash harus lebih besar dari 0.");
    }
    
    const timestamp = Utilities.formatDate(new Date(), spreadsheet.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');

    // Kolom A: Tanggal, Kolom B: Nominal, Kolom C: Tipe Transaksi
    balanceSheet.appendRow([timestamp, nominal, "Tambah Cash"]);
    
    return { 
        status: "SUCCESS", 
        message: `Penambahan cash sebesar ${nominal.toLocaleString('id-ID')} berhasil dicatat.`
    };
}

// ------------------- LOGIC: DATA UTANG TERBARU (SLIDESHOW) -------------------

/**
 * Mengambil daftar N utang terbaru dari semua sheet transaksi yang masih berstatus 'Terhutang'.
 */
function getRecentDebts(limit = 20) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    let allDebts = [];

    for (const sheetName of TRANSACTION_SHEET_NAMES) {
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet || sheet.getLastRow() <= 1) continue; 
        
        // Ambil range dari Kolom A (Timestamp) sampai K (Jumlah Terhutang)
        // Batasi hanya mengambil data terbaru (misalnya, 500 baris terakhir)
        const lastRow = sheet.getLastRow();
        const startRow = Math.max(2, lastRow - 500); // Mulai dari baris data atau 500 baris ke atas
        const numRows = lastRow - startRow + 1;

        if (numRows <= 0) continue;

        // Ambil semua kolom yang relevan
        const range = sheet.getRange(startRow, TIMESTAMP_COL_INDEX, numRows, DEBT_AMOUNT_COL_INDEX);
        const values = range.getValues();

        values.forEach((row, index) => {
            const rowNumber = startRow + index; 
            
            const rowTimestamp = row[TIMESTAMP_COL_INDEX - 1]; // Kolom A
            const rowProduk = row[PRODUK_COL_INDEX - 1];        // Kolom C
            const rowIdentitas = String(row[IDENTITAS_COL_INDEX - 1] || '').trim(); // Kolom D
            const rowStatus = String(row[STATUS_COL_INDEX - 1] || '').trim().toUpperCase(); // Kolom F
            const rawDebtAmount = row[DEBT_AMOUNT_COL_INDEX - 1]; // Kolom K

            const debtAmount = cleanRupiahAndParse(rawDebtAmount);
            
            if (rowStatus === "TERHUTANG" && debtAmount > 0) {
                let rowDate = '';
                if (rowTimestamp instanceof Date) {
                    rowDate = Utilities.formatDate(rowTimestamp, spreadsheet.getSpreadsheetTimeZone(), 'dd MMM yyyy HH:mm');
                }
                
                allDebts.push({
                    id: `${sheetName}-${rowNumber}`, 
                    tanggal: rowDate,
                    identitas: rowIdentitas,
                    produk: rowProduk,
                    nominal: debtAmount,
                    sheet: sheetName,
                    // Tambahkan timestamp mentah untuk sorting
                    timestampRaw: rowTimestamp instanceof Date ? rowTimestamp.getTime() : 0 
                });
            }
        });
    }

    // Urutkan berdasarkan timestamp (terbaru di atas)
    allDebts.sort((a, b) => b.timestampRaw - a.timestampRaw);
    
    // Kembalikan data sesuai limit
    return allDebts.slice(0, limit).map(debt => {
        // Hapus timestampRaw dari output final
        delete debt.timestampRaw; 
        return debt;
    });
}


// ------------------- LOGIC: SUBMIT & UPDATE TRANSAKSI/UTANG -------------------

function submitTransaction(params) {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const sheetName = params.sheetName;
  const sheet = spreadsheet.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`Sheet dengan nama '${sheetName}' tidak ditemukan. Mohon buat sheet ini.`);
  }
  
  // Memastikan format angka yang benar
  const hargaChannel = cleanRupiahAndParse(params.hargaChannel);
  const hargaJual = cleanRupiahAndParse(params.hargaJual);
  const cashDiterimaAgen = cleanRupiahAndParse(params.cashDiterimaAgen);
  const onlineMasuk = cleanRupiahAndParse(params.onlineMasuk);
  const jumlahTerhutang = cleanRupiahAndParse(params.jumlahTerhutang);

  // Menyiapkan data baris baru
  const timestampValue = params.timestamp ? new Date(params.timestamp) : new Date();
  const formattedTimestamp = Utilities.formatDate(timestampValue, spreadsheet.getSpreadsheetTimeZone(), 'yyyy-MM-dd HH:mm:ss');

  const row = [
    formattedTimestamp, // A: Timestamp
    params.channel || '',          // B: Channel
    params.produk || '',           // C: Produk
    params.identitas || '',        // D: Identitas
    params.ewallet || '',          // E: E-Wallet
    params.status || 'Selesai',    // F: Status
    hargaChannel, // G: Harga Channel
    hargaJual,    // H: Harga Jual
    cashDiterimaAgen, // I: Cash Diterima Agen
    onlineMasuk,      // J: Online Masuk
    jumlahTerhutang  // K: Jumlah Terhutang
  ];

  sheet.appendRow(row);
}

function searchForDebt(identitas, date) {
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const results = [];

    const searchIdentitas = identitas ? String(identitas).trim().toLowerCase() : '';

    for (const sheetName of TRANSACTION_SHEET_NAMES) {
        const sheet = spreadsheet.getSheetByName(sheetName);
        if (!sheet || sheet.getLastRow() <= 1) continue;

        // Ambil semua kolom yang relevan
        const range = sheet.getRange(2, TIMESTAMP_COL_INDEX, sheet.getLastRow() - 1, DEBT_AMOUNT_COL_INDEX);
        const values = range.getValues();

        values.forEach((row, index) => {
            const rowNumber = index + 2; 
            
            const rowTimestamp = row[TIMESTAMP_COL_INDEX - 1]; // Kolom A
            const rowProduk = row[PRODUK_COL_INDEX - 1];        // Kolom C
            const rowIdentitas = String(row[IDENTITAS_COL_INDEX - 1] || '').trim().toLowerCase(); // Kolom D
            const rowStatus = String(row[STATUS_COL_INDEX - 1] || '').trim(); // Kolom F
            const rawDebtAmount = row[DEBT_AMOUNT_COL_INDEX - 1]; // Kolom K

            const debtAmount = cleanRupiahAndParse(rawDebtAmount);
            
            let rowDate = '';
            if (rowTimestamp instanceof Date) {
                // Format tanggal untuk perbandingan
                rowDate = Utilities.formatDate(rowTimestamp, spreadsheet.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
            }
            
            const identitasMatch = !searchIdentitas || rowIdentitas.includes(searchIdentitas);
            const dateMatch = !date || rowDate === date;
            const isDebt = rowStatus.toUpperCase() === "TERHUTANG" && debtAmount > 0;
            
            if (identitasMatch && dateMatch && isDebt) {
                results.push({
                    // ID unik untuk update
                    id: `${sheetName}-${rowNumber}`, 
                    identitas: row[IDENTITAS_COL_INDEX - 1],
                    produk: rowProduk,
                    nominal: debtAmount,
                    tanggal: rowDate,
                    sheet: sheetName,
                    row: rowNumber
                });
            }
        });
    }

    return results;
}

function updateDebtEntry(debtId, mode, amount) {
    const [sheetName, rowNumberStr] = debtId.split('-');
    const rowNumber = parseInt(rowNumberStr);
    const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = spreadsheet.getSheetByName(sheetName);

    if (!sheet || isNaN(rowNumber) || rowNumber <= 1) {
        throw new Error("ID Hutang tidak valid atau sheet tidak ditemukan.");
    }
    
    // Range untuk Update
    const debtRange = sheet.getRange(rowNumber, DEBT_AMOUNT_COL_INDEX);
    const cashRange = sheet.getRange(rowNumber, CASH_RECEIVE_COL_INDEX);
    const statusRange = sheet.getRange(rowNumber, STATUS_COL_INDEX);

    let currentDebt = cleanRupiahAndParse(debtRange.getValue());
    let currentCash = cleanRupiahAndParse(cashRange.getValue());

    if (currentDebt <= 0) {
        throw new Error("Hutang ini sudah lunas atau nominalnya nol.");
    }
    
    let actionNominal = 0;
    let finalNominal = currentDebt;
    let message = '';

    if (mode === 'lunas') {
        actionNominal = currentDebt;
        finalNominal = 0;
        
        // Update nilai
        cashRange.setValue(currentCash + actionNominal);
        debtRange.setValue(0); 
        statusRange.setValue("Lunas");
        
        message = "Hutang di set Lunas.";
    } 
    else if (mode === 'partial') {
        const paymentAmount = cleanRupiahAndParse(amount);

        if (isNaN(paymentAmount) || paymentAmount <= 0) {
            throw new Error("Nominal pembayaran sebagian harus lebih besar dari 0.");
        }
        if (paymentAmount > currentDebt) {
            throw new Error("Nominal pembayaran melebihi sisa hutang saat ini.");
        }

        actionNominal = paymentAmount;
        finalNominal = currentDebt - paymentAmount;
        
        // Update nilai
        debtRange.setValue(finalNominal);
        cashRange.setValue(currentCash + actionNominal);
        
        if (finalNominal <= 0) {
            statusRange.setValue("Lunas");
        }
        
        message = `Pembayaran sebagian sebesar ${actionNominal.toLocaleString('id-ID')} berhasil dicatat.`;
    } else {
        throw new Error("Mode update tidak valid.");
    }
    
    return { 
        message: message, 
        newDebt: finalNominal, 
        action: mode 
    };
}

// =========================================================================
// >>>>> FUNGSI TAMBAHAN (SETUP) <<<<<
// =========================================================================

/**
 * Fungsi untuk menginisialisasi sheet saat pertama kali skrip di-deploy atau di-run.
 * Anda bisa menjalankannya secara manual dari editor Apps Script (Run > setupSheets).
 */
function onOpen() {
    setupSheets();
}
