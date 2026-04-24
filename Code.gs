/**
 * Sistem Manajemen Arsip - UPDATED VERSION
 */

const APP_NAME = "LEMATANG";
const SCRIPT_PROP = PropertiesService.getScriptProperties();

const DIVISION_SHEETS = [
  "ArsipMasuk", "ArsipKeluar", "ArsipFoto", "ArsipKKPR", "ArsipAdministrasi"
];

const MAIN_HEADERS = {
  Users: ["ID", "Nama", "Username", "Password", "Role", "Divisi"],
  KategoriArsip: ["Divisi", "Kategori Arsip", "Masa Retensi Aktif (Tahun)", "Masa Retensi Inaktif (Tahun)", "Keterangan"],
  RiwayatDibagikan: ["ID Arsip", "Dibagikan Kepada", "Tanggal Dibagikan", "Dibagikan Oleh"],
  LogAkses: ["User", "Aksi", "Tanggal", "Detail"]
};

const ARSIP_HEADERS = [
  "ID", "Divisi", "Kategori Arsip", "Nama Arsip", "Kode Arsip", 
  "Tanggal Input", "Masa Retensi Aktif", "Masa Retensi Inaktif", 
  "Keterangan", "Jenis File", "File ID", "File URL", "Status"
];

// ==========================================
// SERVING HTML
// ==========================================
function doGet(e) {
  return HtmlService.createTemplateFromFile('index.html')
    .evaluate()
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL); 
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet(e) {
  if (e.parameter.page == "manifest") {
    return HtmlService.createHtmlOutputFromFile("manifest")
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  return HtmlService.createHtmlOutputFromFile("index");
}

// ==========================================
// SETUP & INITIALIZATION
// ==========================================
function getOrCreateMainSpreadsheet() {
  const id = SCRIPT_PROP.getProperty('SPREADSHEET_ID');
  
  if (id) {
    try {
      return SpreadsheetApp.openById(id);
    } catch (e) {
      console.log("Spreadsheet ID Error: " + e.toString() + ". Resetting ID...");
      SCRIPT_PROP.deleteProperty('SPREADSHEET_ID');
    }
  }
  
  const ss = SpreadsheetApp.create(APP_NAME + " DB");
  SCRIPT_PROP.setProperty('SPREADSHEET_ID', ss.getId());
  
  for (const [name, headers] of Object.entries(MAIN_HEADERS)) {
    const sheet = ss.insertSheet(name);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#E8F5E9");
  }
  
  const masterSheet = ss.insertSheet("ArsipMasuk");
  masterSheet.appendRow(ARSIP_HEADERS);
  masterSheet.getRange(1, 1, 1, ARSIP_HEADERS.length).setFontWeight("bold").setBackground("#E8F5E9");
  
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet) defaultSheet.deleteSheet();
  
  seedDefaultData(ss);
  return ss;
}

function seedDefaultData(ss) {
  const userSheet = ss.getSheetByName("Users");
  if (userSheet.getLastRow() <= 1) {
      userSheet.appendRow([1, "Administrator", "admin", "admin123", "Super Admin", ""]);
  }
  
  const catSheet = ss.getSheetByName("KategoriArsip");
  if (catSheet.getLastRow() <= 1) {
      const categories = [
        ["ArsipMasuk", "Dokumen Masuk", "5", "10", "Penting"],
        ["ArsipKeluar", "Dokumen Keluar", "10", "20", "Sangat Rahasia"],
        ["ArsipFoto", "Dokumentasi Kegiatan", "2", "5", "Teknis"]
      ];
      categories.forEach(cat => catSheet.appendRow(cat));
  }
}

function ensureDefaultAdmin() {
  try {
    const ss = getOrCreateMainSpreadsheet();
    const sheet = ss.getSheetByName("Users");
    if (sheet.getLastRow() <= 1) {
      sheet.appendRow([1, "Administrator", "admin", "admin123", "Super Admin", ""]);
      return { success: true, message: "User Admin berhasil dibuat otomatis." };
    } else {
      return { success: false, message: "Database User sudah memiliki data." };
    }
  } catch (e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

// ==========================================
// GOOGLE DRIVE LOGIC
// ==========================================
function getOrCreateDivisionFolder(divisionName) {
  const folders = DriveApp.getFoldersByName(divisionName);
  if (folders.hasNext()) return folders.next();
  return DriveApp.getRootFolder().createFolder(divisionName);
}

function uploadFileToDrive(base64Data, fileName, folderId) {
  try {
    const split = base64Data.split('base64,');
    const type = split[0].split(';')[0].replace('data:', '');
    const data = Utilities.base64Decode(split[1]);
    const blob = Utilities.newBlob(data, type, fileName);
    
    const folder = DriveApp.getFolderById(folderId);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    return { id: file.getId(), url: file.getUrl(), name: file.getName() };
  } catch (e) {
    throw new Error("Gagal upload file: " + e.toString());
  }
}

// ==========================================
// AUTHENTICATION
// ==========================================
function authenticateUser(username, password) {
  const ss = getOrCreateMainSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    // Pastikan password ada (index 3)
    if (data[i][2] == username && data[i][3] == password) {
      // AMBIL DIVISI DARI KOLOM TERAKHIR (index 5)
      const userDivisi = data[i].length > 5 ? data[i][5] : ""; 
      
      logAkses(data[i][2], "Login", "User login berhasil");
      return {
        status: true,
        user: { 
          id: data[i][0], 
          name: data[i][1], 
          username: data[i][2], 
          role: data[i][4],
          divisi: userDivisi // PENTING: Kirim divisi user
        }
      };
    }
  }
  return { status: false, message: "Username atau Password salah" };
}

function logAkses(user, aksi, detail) {
  const ss = getOrCreateMainSpreadsheet();
  const sheet = ss.getSheetByName("LogAkses");
  sheet.appendRow([user, aksi, new Date(), detail]);
}

// ==========================================
// CORE LOGIC: CRUD ARSIP
// ==========================================
// UPDATE FUNGSI INI (Perbaikan Filter Keamanan)
function getArchives(filter) {
  try {
    const ss = getOrCreateMainSpreadsheet();
    const sheet = ss.getSheetByName("ArsipMasuk");
    if (!sheet) throw new Error("Sheet ArsipMasuk tidak ditemukan");
    
    // SECURITY CHECK
    if (filter.userRole !== "Super Admin") {
      filter.divisi = filter.userDivisi;
    }

    const data = sheet.getDataRange().getValues();
    const results = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      
      if (!row[0] || row[0] == "") continue; 

      // --- PERBAIKAN 3: Validasi Tanggal di getArchives ---
      let dateObj = row[5];
      if (typeof dateObj === 'string' || !(dateObj instanceof Date)) {
        dateObj = new Date(dateObj);
      }
      if (!dateObj || isNaN(dateObj.getTime())) continue; // Skip jika tanggal rusak
      // ----------------------------------------------------

      let status = "Permanen";
      try {
        const aktif = parseInt(row[6]) || 0;
        const inaktif = parseInt(row[7]) || 0;
        status = calculateArchiveStatus(dateObj, aktif, inaktif);
      } catch(e) { status = "Error"; }

      // Gunakan dateObj yang sudah pasti Object Date di sini
      results.push({
        id: String(row[0] || ""),
        divisi: String(row[1] || ""),
        kategori: String(row[2] || ""),
        nama: String(row[3] || ""),
        kode: String(row[4] || ""),
        tanggalDisplay: Utilities.formatDate(dateObj, "id-ID", "dd MMM yyyy"), // Aman sekarang
        status: status,
        keterangan: String(row[8] || "-"),
        jenisFile: String(row[9] || "-"), 
        fileId: String(row[10] || ""),
        fileUrl: String(row[11] || ""),
        retensiAktif: row[6],
        retensiInaktif: row[7]
      });
    }

    // Filter Logika
    let finalData = results;
    if (filter) {
      finalData = results.filter(item => {
        let matchDiv = true;
        let matchSearch = true;
        let matchDate = true;

        if (filter.divisi && filter.divisi !== "All") {
          if (item.divisi.trim() !== filter.divisi.trim()) matchDiv = false;
        }
        
        if (filter.search && filter.search.trim() !== "") {
          const s = filter.search.toLowerCase();
          if (!item.nama.toLowerCase().includes(s) && !item.kode.toLowerCase().includes(s)) matchSearch = false;
        }

        if (filter.startDate || filter.endDate) {
        }

        return matchDiv && matchSearch && matchDate;
      });
    }
    
    return finalData;
    
  } catch (error) {
    throw new Error("Gagal Get Archives: " + error.toString());
  }
}

function calculateArchiveStatus(inputDate, activeYears, inactiveYears) {
  if (!inputDate || isNaN(inputDate.getTime())) return "Permanen";
  const now = new Date();
  const activeDate = new Date(inputDate);
  activeDate.setFullYear(activeDate.getFullYear() + parseInt(activeYears));
  const inactiveDate = new Date(activeDate);
  inactiveDate.setFullYear(inactiveDate.getFullYear() + parseInt(inactiveYears));
  
  if (now < activeDate) return "Aktif";
  if (now >= activeDate && now < inactiveDate) return "Inaktif";
  if (now >= inactiveDate) return "Harus Dimusnahkan";
  return "Permanen"; 
}

function addArchive(formData) {
  const ss = getOrCreateMainSpreadsheet();
  const sheet = ss.getSheetByName("ArsipMasuk"); 
  if (!sheet) throw new Error("Sheet ArsipMasuk tidak ditemukan");
  
  const newId = Utilities.getUuid();
  const newCode = generateCodeByDivision(formData.divisi);
  
  let fileData = { id: "", url: "" };
  if (formData.fileData) {
    const originalName = formData.fileName;
    const ext = originalName.substring(originalName.lastindexOf("."));
    const newFileName = formData.nama + ext;
    const folder = getOrCreateDivisionFolder(formData.divisi);
    fileData = uploadFileToDrive(formData.fileData, newFileName, folder.getId());
  }
  
  const inputDate = new Date(); 
  const status = calculateArchiveStatus(inputDate, formData.retensiAktif, formData.retensiInaktif);
  
  sheet.appendRow([
    newId, formData.divisi, formData.kategori, formData.nama, newCode, 
    inputDate, formData.retensiAktif, formData.retensiInaktif, formData.keterangan, 
    formData.fileType, fileData.id, fileData.url, status
  ]);
  
  logAkses(formData.userName, "Tambah Arsip", "Menambahkan " + newCode);
  return { success: true, code: newCode };
}

function editArchive(id, archive) {
  const ss = getOrCreateMainSpreadsheet();
  const sheet = ss.getSheetByName("ArsipMasuk");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      let fileData = { id: data[i][10], url: data[i][11] };
      if (archive.fileData) {
        try { DriveApp.getFileById(fileData.id).setTrashed(true); } catch(e){}
        const ext = archive.fileName.substring(archive.fileName.lastindexOf("."));
        const newFileName = archive.nama + ext;
        const folder = getOrCreateDivisionFolder(archive.divisi);
        fileData = uploadFileToDrive(archive.fileData, newFileName, folder.getId());
      }
      
      sheet.getRange(i + 1, 1, 1, 12).setValues([[
        data[i][0], archive.divisi, archive.kategori, archive.nama, 
        data[i][4], data[i][5], data[i][6], data[i][7], 
        archive.keterangan, archive.fileType || data[i][9], fileData.id, fileData.url
      ]]);
      
      logAkses(archive.userName, "Update Arsip", "Update ID " + id);
      return { success: true };
    }
  }
  throw new Error("Data tidak ditemukan di ArsipMasuk");
}

function deleteArchive(id, division, userName) {
  const ss = getOrCreateMainSpreadsheet();
  const sheet = ss.getSheetByName("ArsipMasuk");
  const data = sheet.getDataRange().getValues();
  
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] == id) {
      const fileId = data[i][10];
      if (fileId) {
        try { DriveApp.getFileById(fileId).setTrashed(true); } catch (e) {}
      }
      
      sheet.deleteRow(i + 1);
      logAkses(userName, "Hapus Arsip", "Menghapus ID " + id);
      return { success: true };
    }
  }
  throw new Error("Data tidak ditemukan di ArsipMasuk");
}

function getCategories(filter) {
  const ss = getOrCreateMainSpreadsheet();
  const sheet = ss.getSheetByName("KategoriArsip");
  if (!sheet) throw new Error("Sheet 'KategoriArsip' tidak ditemukan.");
  
  const data = sheet.getDataRange().getValues();
  const result = [];
  
  const userRole = (filter && filter.userRole) ? filter.userRole : "User";
  const userDivisi = (filter && filter.userDivisi) ? filter.userDivisi : "";
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowDivisi = row[0];

    if (userRole !== "Super Admin") {
      if (rowDivisi !== userDivisi) continue; 
    }

    result.push({
      rowindex: i + 1,
      divisi: row[0],
      kategori: row[1],
      retensiAktif: row[2],
      retensiInaktif: row[3],
      keterangan: row[4]
    });
  }
  return result;
}

function addCategory(data) {
  const ss = getOrCreateMainSpreadsheet();
  const sheet = ss.getSheetByName("KategoriArsip");
  sheet.appendRow([data.divisi, data.kategori, data.aktif, data.inaktif, data.keterangan]);
  return { success: true };
}

function updateCategory(rowindex, data) {
  const ss = getOrCreateMainSpreadsheet();
  const sheet = ss.getSheetByName("KategoriArsip");
  sheet.getRange(rowindex, 1, 1, 5).setValues([[data.divisi, data.kategori, data.aktif, data.inaktif, data.keterangan]]);
  return { success: true };
}

function deleteCategory(rowindex) {
  const ss = getOrCreateMainSpreadsheet();
  const sheet = ss.getSheetByName("KategoriArsip");
  sheet.deleteRow(rowindex);
  return { success: true };
}

// ==========================================
// DASHBOARD & REPORTS
// ==========================================
function getDashboardData(filter) {
  const ss = getOrCreateMainSpreadsheet();
  const sheet = ss.getSheetByName("ArsipMasuk");
  
  // 1. Ambil Info Role & Divisi dari Parameter
  const userRole = (filter && filter.userRole) ? filter.userRole : "User";
  const userDivisi = (filter && filter.userDivisi) ? filter.userDivisi : "";

  // 2. Persiapan Filter User untuk Log
  let allowedUsernames = [];
  if (userRole !== "Super Admin") {
    const userSheet = ss.getSheetByName("Users");
    const userData = userSheet.getDataRange().getValues();
    // Ambil list username yang divisinya sama dengan userDivisi
    allowedUsernames = userData
      .filter(row => row[5] === userDivisi) // row[5] adalah Divisi
      .map(row => row[2]);                  // row[2] adalah Username
  }

  // 3. Hitung Statistik Arsip & Chart Data
  let totalArsip = 0, totalAktif = 0, totalInaktif = 0, totalMusnah = 0;
  const divCounts = {};
  const archiveDates = []; 

  if (sheet && sheet.getLastRow() > 1) {
    const data = sheet.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const div = row[1]; // Kolom Divisi (index 1)
      let date = row[5];  // Kolom Tanggal (index 5)
      
      // --- PERBAIKAN 1: Pastikan date adalah Object Date ---
      if (typeof date === 'string' || !(date instanceof Date)) {
        date = new Date(date);
      }
      // Jika tanggal invalid, skip agar tidak error
      if (isNaN(date.getTime())) continue; 
      // ----------------------------------------------

      // FILTER DATA BERDASARKAN ROLE
      if (userRole !== "Super Admin") {
        if (div !== userDivisi) continue; 
      }

      totalArsip++;
      divCounts[div] = (divCounts[div] || 0) + 1;
      
      const status = calculateArchiveStatus(date, row[6], row[7]);
      if (status === "Aktif") totalAktif++;
      else if (status === "Inaktif") totalInaktif++;
      else if (status === "Harus Dimusnahkan") totalMusnah++;
      
      // Karena sudah dipastikan 'date' adalah Object Date, ini aman:
      archiveDates.push(Utilities.formatDate(date, "en-CA", "yyyy-MM-dd")); 
    }
  }

  // 4. Hitung Log Akses (DIFILTER)
  let recentLogs = [];
  const logSheet = ss.getSheetByName("LogAkses");
  if (logSheet && logSheet.getLastRow() > 1) {
    const logData = logSheet.getDataRange().getValues();
    
    // Filter logs berdasarkan daftar user di divisi
    const filteredLogs = logData.filter(logRow => {
      const logUser = logRow[0]; 
      if (userRole === "Super Admin") return true;
      return allowedUsernames.includes(logUser); 
    });

    const slicedLogs = filteredLogs.slice(-5);
    
    slicedLogs.reverse().forEach(row => {
      let logDate = row[2]; // Kolom Tanggal Log
      
      // --- PERBAIKAN 2: Validasi Tanggal Log ---
      if (typeof logDate === 'string' || !(logDate instanceof Date)) {
        logDate = new Date(logDate);
      }
      // ---------------------------------------

      // Pastikan valid sebelum format
      let dateStr = "-";
      if (!isNaN(logDate.getTime())) {
        dateStr = Utilities.formatDate(logDate, "Asia/Jakarta", "dd MMM yyyy, HH:mm"); 
      }

      recentLogs.push({
        user: row[0],
        action: row[1],
        date: dateStr,
        detail: row[3]
      });
    });
  }

  return {
    totalArsip, totalAktif, totalInaktif, totalMusnah, totalUsers: 2, 
    divCounts, 
    chartData: { dates: archiveDates },
    recentLogs: recentLogs
  };
}
function getUsers() {
  const ss = getOrCreateMainSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const data = sheet.getDataRange().getValues();
  const result = [];
  for (let i = 1; i < data.length; i++) {
    result.push({
      rowindex: i + 1,
      id: data[i][0],
      nama: data[i][1],
      username: data[i][2],
      role: data[i][4],
      divisi: data[i][5]
    });
  }
  return result;
}

function addUser(data) {
  const ss = getOrCreateMainSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  const newId = sheet.getLastRow();
  sheet.appendRow([newId, data.nama, data.username, data.password, data.role, data.divisi]);
  return { success: true };
}

function updateUser(rowindex, data) {
  const ss = getOrCreateMainSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  sheet.getRange(rowindex, 1, 1, 6).setValues([[data.id, data.nama, data.username, data.password, data.role, data.divisi]]);
  return { success: true };
}

function deleteUser(rowindex) {
  const ss = getOrCreateMainSpreadsheet();
  const sheet = ss.getSheetByName("Users");
  sheet.deleteRow(rowindex);
  return { success: true };
}

function exportToExcel(data) {
  try {
    const ss = SpreadsheetApp.create("Laporan_Arsip_" + new Date().getTime());
    const sheet = ss.getSheets()[0];
    sheet.appendRow(["ID", "Divisi", "Nama Arsip", "Kode", "Tanggal", "Status", "Keterangan"]);
    
    const outputValues = [];
    if (data && data.length > 0) {
      data.forEach(item => {
        outputValues.push([
          item.id || "",
          item.divisi || "",
          item.nama || "",
          item.kode || "",
          item.tanggalDisplay || "-",
          item.status || "-",          
          item.keterangan || "-"
        ]);
      });
      sheet.getRange(2, 1, outputValues.length, 7).setValues(outputValues);
      for (let i = 1; i <= 7; i++) sheet.autoResizeColumn(i);
    }

    const fileId = ss.getId();
    const file = DriveApp.getFileById(fileId);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    Utilities.sleep(2000); 
    return fileId; 
    
  } catch (e) {
    throw new Error("Gagal Export Excel: " + e.toString());
  }
}

function generateCodeByDivision(divisionName) {
  const ss = getOrCreateMainSpreadsheet();
  const sheet = ss.getSheetByName("ArsipMasuk");
  if (!sheet) return "ERR001";
  const data = sheet.getDataRange().getValues();
  let count = 0;
  for (let i = 1; i < data.length; i++) {
    if (data[i][1] == divisionName) count++;
  }
  let prefix = divisionName.replace("Arsip", "").substring(0,3).toUpperCase();
  return prefix + String(count + 1).padStart(3, '0');
}

function setSession() {
  return { status: "ok" };
}
