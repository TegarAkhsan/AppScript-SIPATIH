const SPREADSHEET_ID = '1UFEXxqUh5OKyPVEQh4fmPo0Y2i3FdOd7xXLlyBLN5XU';
const FOLDER_MASUK_ID = '1nSmgCSCslt3_cG3AS4S63gqSnCbEQNNe';
const FOLDER_KELUAR_ID = '1P5i0zcrBWa97aAY5fkjMPsqRTpB_bZGX';
const DRIVE_FOLDER_ID = '1GCrXpEb70cnsTvyTpnglyaSqyto0m_PC'; // Default fallback

function doGet() {
  checkAndInitializeBidang();
  return HtmlService.createTemplateFromFile('index').evaluate()
      .setTitle('SIPATIH - Sistem Arsip Digital Desa Kepatihan')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename)
      .getContent();
}

function getSS() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

// ----------------------------------------------------
// Setup & Clean Up (Bisa dijalankan manual dari Editor)
// ----------------------------------------------------
function cleanUpData() {
  const ss = getSS();
  const sheets = ['Pengguna', 'Aktivitas', 'Naskah Masuk', 'Naskah Keluar', 'Klasifikasi Arsip'];
  
  sheets.forEach(sheetName => {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
    }
    // Hapus isi tabel dari baris 2 kebawah agar tabel kosong (tetapi header tidak dihapus)
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getMaxColumns()).clearContent();
    }
  });

  // Setup default headers and first data if necessary
  setupHeaders();
}

function setupHeaders() {
  const ss = getSS();
  
  // Helper to ensure sheet exists
  const getOrCreateSheet = (name) => {
    let s = ss.getSheetByName(name);
    if (!s) s = ss.insertSheet(name);
    return s;
  };

  // Pengguna
  let sh = getOrCreateSheet('Pengguna');
  if (sh.getLastRow() <= 1) {
    if (sh.getLastRow() === 0) {
      sh.appendRow(['ID', 'Username', 'Password', 'Role', 'Status', 'Bidang']);
    }
    // Pastikan tidak menduplikasi jika sudah ada data tapi kurang dari 2 baris
    if (sh.getLastRow() === 1) {
      sh.appendRow(['1', 'admin', 'admin', 'Admin', 'Aktif', 'Semua']);
      sh.appendRow(['2', 'user', 'user', 'Perangkat Desa', 'Aktif', 'Pemerintahan']);
    }
  } else {
    // Migrasi kolom Bidang jika belum ada
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    if (!headers.includes('Bidang')) {
      const colIndex = sh.getLastColumn() + 1;
      sh.getRange(1, colIndex).setValue('Bidang');
      
      // Berikan nilai default 'Semua' untuk baris data yang ada
      const lastRow = sh.getLastRow();
      if (lastRow > 1) {
        const defaultValues = Array(lastRow - 1).fill(['Semua']);
        sh.getRange(2, colIndex, lastRow - 1, 1).setValues(defaultValues);
      }
    }
  }
  
  // Tambahkan/seeding akun bidang secara dinamis jika belum ada
  const currentLastRow = sh.getLastRow();
  if (currentLastRow > 1) {
    const existingUsers = sh.getRange(2, 2, currentLastRow - 1, 1).getValues().map(r => r[0].toString().toLowerCase().trim());
    const defaultAccounts = [
      { username: 'pemerintahan', password: 'pemerintahan', role: 'Perangkat Desa', bidang: 'Pemerintahan' },
      { username: 'kesejahteraan', password: 'kesejahteraan', role: 'Perangkat Desa', bidang: 'Kesejahteraan' },
      { username: 'kependudukan', password: 'kependudukan', role: 'Perangkat Desa', bidang: 'Kependudukan' },
      { username: 'pertanahan', password: 'pertanahan', role: 'Perangkat Desa', bidang: 'Pertanahan' },
      { username: 'umum', password: 'umum', role: 'Perangkat Desa', bidang: 'Umum' }
    ];

    defaultAccounts.forEach(acc => {
      if (!existingUsers.includes(acc.username)) {
        const nextId = (sh.getLastRow()).toString();
        sh.appendRow([nextId, acc.username, acc.password, acc.role, 'Aktif', acc.bidang]);
      }
    });
  }

  
  // Aktivitas
  sh = getOrCreateSheet('Aktivitas');
  if(sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp', 'Username', 'Aktivitas', 'Detail']);
  }

  // Notifikasi
  sh = getOrCreateSheet('Notifikasi');
  if(sh.getLastRow() === 0) {
    sh.appendRow(['ID', 'Timestamp', 'Username', 'Bidang', 'Pesan', 'Status']);
  }

  // Naskah Masuk
  sh = getOrCreateSheet('Naskah Masuk');
  if(sh.getLastRow() === 0) {
    sh.appendRow(['Nomor Indeks', 'Nomor Agenda', 'Kode Klasifikasi', 'Nomor Surat', 'Tanggal Surat', 'Tanggal Terima', 'Asal Instansi', 'Perihal', 'Sifat Surat', 'Bidang', 'Disposisi', 'Status Arsip', 'Keterangan', 'Diinput Oleh', 'Link File', 'Jenis Surat', 'Keperluan', 'Tanda Tangan']);
  } else {
    // Check if headers have the new fields, if not append them to row 1
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const missing = [];
    if (!headers.includes('Jenis Surat')) missing.push('Jenis Surat');
    if (!headers.includes('Keperluan')) missing.push('Keperluan');
    if (!headers.includes('Tanda Tangan')) missing.push('Tanda Tangan');
    if (missing.length > 0) {
      const startCol = sh.getLastColumn() + 1;
      sh.getRange(1, startCol, 1, missing.length).setValues([missing]);
    }
  }

  // Naskah Keluar
  sh = getOrCreateSheet('Naskah Keluar');
  if(sh.getLastRow() === 0) {
    sh.appendRow(['Nomor Indeks', 'Nomor Agenda', 'Kode Klasifikasi', 'Nomor Surat', 'Tanggal Surat', 'Tujuan Instansi', 'Perihal', 'Sifat Surat', 'Bidang', 'Status Arsip', 'Keterangan', 'Diinput Oleh', 'Link File', 'Jenis Surat', 'Keperluan', 'Tanda Tangan']);
  } else {
    // Check if headers have the new fields, if not append them to row 1
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    const missing = [];
    if (!headers.includes('Jenis Surat')) missing.push('Jenis Surat');
    if (!headers.includes('Keperluan')) missing.push('Keperluan');
    if (!headers.includes('Tanda Tangan')) missing.push('Tanda Tangan');
    if (missing.length > 0) {
      const startCol = sh.getLastColumn() + 1;
      sh.getRange(1, startCol, 1, missing.length).setValues([missing]);
    }
  }

  // Klasifikasi Arsip (Sesuai Permintaan User)
  sh = getOrCreateSheet('Klasifikasi Arsip');
  if(sh.getLastRow() <= 1) {
    if(sh.getLastRow() === 0) {
      sh.appendRow(['Baris', 'Kode', 'Nama Bidang', 'Jenis Arsip', 'Retensi Aktif', 'Retensi Inaktif']);
    }
    if(sh.getLastRow() === 1) {
      const klasifikasiAwal = [
        ['2', 'PEM', 'Pemerintahan', '-', '', ''],
        ['3', 'KES', 'Kesejahteraan', '-', '', ''],
        ['4', 'KEP', 'Kependudukan', '-', '', ''],
        ['5', 'UMU', 'Umum', '-', '', '']
      ];
      klasifikasiAwal.forEach(row => sh.appendRow(row));
    }
  }
}

// ----------------------------------------------------
// Authentication & Tracking
// ----------------------------------------------------
function login(username, password) {
  try {
    const sheet = getSS().getSheetByName('Pengguna');
    if(!sheet) return { success: false, message: 'Sheet Pengguna belum dibuat.' };
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      // Jika tabel pengguna kosong sama sekali (tidak termasuk header), izinkan login bypass sementara (atau admin default)
      if(username === 'admin' && password === 'admin') {
        logActivity('admin', 'Login Bypass', 'Login saat tabel pengguna kosong');
        return { success: true, role: 'Admin', username: 'admin', bidang: 'Semua' };
      }
      return { success: false, message: 'Tidak ada data pengguna dalam sistem.' };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    
    for (let i = 0; i < data.length; i++) {
      let row = data[i];
      let dbUsername = row[1];
      let dbPassword = row[2];
      let role = row[3];
      let status = row[4];
      let bidang = row[5] || 'Semua';
      
      if (dbUsername === username && dbPassword === password) {
        if (status && status.toString().toLowerCase() !== 'aktif') {
          return { success: false, message: 'Akun Anda dinonaktifkan!' };
        }
        logActivity(username, 'Login', 'Berhasil login ke sistem');
        return { success: true, role: role, username: username, bidang: bidang };
      }
    }
    return { success: false, message: 'Username atau password salah!' };
  } catch (e) {
    return { success: false, message: 'Terjadi kesalahan sistem: ' + e.message };
  }
}

function logActivity(username, aktivitas, detail) {
  try {
    const sheet = getSS().getSheetByName('Aktivitas');
    if(sheet) {
      const timestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
      sheet.appendRow([timestamp, username, aktivitas, detail]);
    }
    return true;
  } catch(e) {
    return false;
  }
}

// ----------------------------------------------------
// Data Retrieval (Dashboard & Table)
// ----------------------------------------------------
function getDashboardData(username) {
  try {
    checkAndInitializeBidang();
    setupHeaders(); // Pastikan header dan data awal tersedia
    const ss = getSS();
    
    // Get user details to check Bidang restriction
    let userBidang = "Semua";
    let userRole = "Admin";
    if (username) {
      const uSheet = ss.getSheetByName('Pengguna');
      if (uSheet && uSheet.getLastRow() > 1) {
        const uData = uSheet.getRange(2, 2, uSheet.getLastRow() - 1, 5).getValues();
        for (let i = 0; i < uData.length; i++) {
          if (uData[i][0] === username) {
            userRole = uData[i][2]; // Column D is index 2 from Col B
            userBidang = uData[i][4] || "Semua"; // Column F is index 4 from Col B
            break;
          }
        }
      }
    }
    
    const arsipMasukSheet = ss.getSheetByName('Naskah Masuk');
    const arsipKeluarSheet = ss.getSheetByName('Naskah Keluar');
    const klasifikasiSheet = ss.getSheetByName('Klasifikasi Arsip');
    
    let bidangMap = {};
    let bidangCodes = {};
    if (klasifikasiSheet && klasifikasiSheet.getLastRow() > 1) {
      const klasData = klasifikasiSheet.getRange(2, 1, klasifikasiSheet.getLastRow() - 1, 2).getValues(); // Col A (Nama Bidang) dan Col B (Kode)
      klasData.forEach(row => {
        const b = row[0] ? row[0].toString().trim() : "";
        const code = row[1] ? row[1].toString().trim() : "";
        
        // Bidang induk diidentifikasi jika kode memiliki panjang tepat 3 karakter (misal PEM, KES, KEP, PER, UMU)
        if (code.length === 3 && b) {
          if (userRole === "Admin" || userRole === "Kepala Desa" || userBidang === "Semua" || b === userBidang) {
            bidangMap[b] = 0;
            bidangCodes[b] = code.toUpperCase();
          }
        }
      });
    }

    let masukCount = 0, keluarCount = 0;
    let aktifCount = 0, inaktifCount = 0;
    
    if (arsipMasukSheet && arsipMasukSheet.getLastRow() > 1) {
      const allMasuk = arsipMasukSheet.getRange(2, 1, arsipMasukSheet.getLastRow() - 1, 18).getValues();
      allMasuk.forEach(row => {
        const b = row[9] ? row[9].toString().trim() : "";
        if (userRole === "Admin" || userRole === "Kepala Desa" || userBidang === "Semua" || b === userBidang) {
          masukCount++;
          const val = row[11] ? row[11].toString().trim().toLowerCase() : "";
          if (val === 'non aktif' || val === 'inaktif') {
            inaktifCount++;
          } else {
            aktifCount++;
          }
          if (b && bidangMap[b] !== undefined) {
            bidangMap[b] = (bidangMap[b] || 0) + 1;
          }
        }
      });
    }
    
    if (arsipKeluarSheet && arsipKeluarSheet.getLastRow() > 1) {
      const allKeluar = arsipKeluarSheet.getRange(2, 1, arsipKeluarSheet.getLastRow() - 1, 16).getValues();
      allKeluar.forEach(row => {
        const b = row[8] ? row[8].toString().trim() : "";
        if (userRole === "Admin" || userRole === "Kepala Desa" || userBidang === "Semua" || b === userBidang) {
          keluarCount++;
          const val = row[9] ? row[9].toString().trim().toLowerCase() : "";
          if (val === 'non aktif' || val === 'inaktif') {
            inaktifCount++;
          } else {
            aktifCount++;
          }
          if (b && bidangMap[b] !== undefined) {
            bidangMap[b] = (bidangMap[b] || 0) + 1;
          }
        }
      });
    }

    const totalArsip = masukCount + keluarCount;

    const bidangList = Object.keys(bidangMap).map(k => ({
      kode: bidangCodes[k] || k.substring(0, 3).toUpperCase(), 
      nama: k,
      jumlah: bidangMap[k]
    }));

    return {
      totalArsip: totalArsip,
      totalBidang: Object.keys(bidangMap).length,
      arsipAktif: aktifCount,
      arsipInaktif: inaktifCount,
      bidangList: bidangList
    };
  } catch (e) {
    return { error: true, message: e.message };
  }
}

// Helper untuk memformat objek Date atau String tanggal menjadi format string dd-MM-yyyy secara aman pada backend
function formatDateToString(val) {
  if (!val) return '-';
  
  // Jika ini adalah objek Date (atau bertipe objek dengan getMonth)
  if (val instanceof Date || (val && typeof val === 'object' && typeof val.getMonth === 'function')) {
    try {
      return Utilities.formatDate(val, 'Asia/Jakarta', 'dd-MM-yyyy');
    } catch(err) {
      // Fallback manual jika Utilities.formatDate gagal
      const y = val.getFullYear();
      const m = String(val.getMonth() + 1).padStart(2, '0');
      const d = String(val.getDate()).padStart(2, '0');
      return `${d}-${m}-${y}`;
    }
  }
  
  const str = val.toString().trim();
  if (!str || str === '-') return '-';
  
  // Jika formatnya ISO (mengandung T), kita parsing ke Date dulu lalu format ke dd-MM-yyyy
  if (str.indexOf('T') !== -1) {
    const d = new Date(str);
    if (!isNaN(d.getTime())) {
      try {
        return Utilities.formatDate(d, 'Asia/Jakarta', 'dd-MM-yyyy');
      } catch(err) {
        const y = d.getFullYear();
        const m = String(d.getMonth() + 1).padStart(2, '0');
        const day = String(d.getDate()).padStart(2, '0');
        return `${day}-${m}-${y}`;
      }
    }
  }
  
  // Jika formatnya yyyy-MM-dd atau yyyy/MM/dd, ubah menjadi dd-MM-yyyy
  let parts = str.split(/[-/]/);
  if (parts.length === 3) {
    const p0 = parts[0].trim();
    const p1 = parts[1].trim();
    const p2 = parts[2].trim();
    if (p0.length === 4) { // yyyy-MM-dd
      return `${p2.padStart(2, '0')}-${p1.padStart(2, '0')}-${p0}`;
    } else if (p2.length === 4) { // dd-MM-yyyy
      return `${p0.padStart(2, '0')}-${p1.padStart(2, '0')}-${p2}`;
    }
  }
  
  return str;
}

function getArsipData(username) {
  try {
    const ss = getSS();
    let result = [];
    
    // Get user details to check Bidang restriction
    let userBidang = "Semua";
    let userRole = "Admin";
    if (username) {
      const uSheet = ss.getSheetByName('Pengguna');
      if (uSheet && uSheet.getLastRow() > 1) {
        const uData = uSheet.getRange(2, 2, uSheet.getLastRow() - 1, 5).getValues();
        for (let i = 0; i < uData.length; i++) {
          if (uData[i][0] && uData[i][0].toString().toLowerCase() === username.toString().toLowerCase()) {
            userRole = uData[i][2]; // Column D is index 2 from Col B
            userBidang = uData[i][4] || "Semua"; // Column F is index 4 from Col B
            break;
          }
        }
      }
    }
    
    const masukSheet = ss.getSheetByName('Naskah Masuk');
    if (masukSheet && masukSheet.getLastRow() > 1) {
      const lastCol = Math.max(masukSheet.getLastColumn(), 18);
      const mData = masukSheet.getRange(2, 1, masukSheet.getLastRow() - 1, lastCol).getValues();
      mData.forEach(row => {
        if(!row[0] && !row[3]) return;
        const b = row[9] ? row[9].toString().trim() : "";
        if (userRole === "Admin" || userRole === "Kepala Desa" || userBidang === "Semua" || b === userBidang) {
          result.push({
            jenis: 'Masuk',
            noIndeks: row[0],
            noAgenda: row[1],
            kodeKlasifikasi: row[2],
            noSurat: row[3],
            tglSurat: formatDateToString(row[4]),
            tglTerima: formatDateToString(row[5]),
            asalInstansi: row[6],
            instansi: row[6],
            perihal: row[7],
            sifatSurat: row[8],
            bidang: row[9],
            disposisi: row[10],
            statusArsip: row[11] || 'Aktif',
            keterangan: row[12],
            linkFile: row[14],
            jenisSurat: row[15] || '-',
            keperluan: row[16] || '-',
            tandaTangan: row[17] || '-'
          });
        }
      });
    }

    const keluarSheet = ss.getSheetByName('Naskah Keluar');
    if (keluarSheet && keluarSheet.getLastRow() > 1) {
      const lastCol = Math.max(keluarSheet.getLastColumn(), 16);
      const kData = keluarSheet.getRange(2, 1, keluarSheet.getLastRow() - 1, lastCol).getValues();
      kData.forEach(row => {
        if(!row[0] && !row[3]) return;
        const b = row[8] ? row[8].toString().trim() : "";
        if (userRole === "Admin" || userRole === "Kepala Desa" || userBidang === "Semua" || b === userBidang) {
          result.push({
            jenis: 'Keluar',
            noIndeks: row[0],
            noAgenda: row[1],
            kodeKlasifikasi: row[2],
            noSurat: row[3],
            tglSurat: formatDateToString(row[4]),
            tujuanInstansi: row[5],
            instansi: row[5],
            perihal: row[6],
            sifatSurat: row[7],
            bidang: row[8],
            statusArsip: row[9] || 'Aktif',
            keterangan: row[10],
            linkFile: row[12],
            jenisSurat: row[13] || '-',
            keperluan: row[14] || '-',
            tandaTangan: row[15] || '-'
          });
        }
      });
    }
    
    return result;
  } catch (e) {
    return [];
  }
}

function getAktivitas() {
  try {
    const sheet = getSS().getSheetByName('Aktivitas');
    if (!sheet || sheet.getLastRow() <= 1) return [];
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
    return data.reverse().map(row => ({
      timestamp: row[0] instanceof Date ? Utilities.formatDate(row[0], 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss') : (row[0] || '-'),
      username: row[1],
      aktivitas: row[2],
      detail: row[3]
    }));
  } catch (e) {
    return [];
  }
}

function getUsers() {
  try {
    const sheet = getSS().getSheetByName('Pengguna');
    if (!sheet || sheet.getLastRow() <= 1) return [];
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    return data.map(row => ({
      id: row[0],
      username: row[1],
      password: row[2],
      role: row[3],
      status: row[4] || 'Aktif',
      bidang: row[5] || 'Semua'
    })).filter(u => u.username);
  } catch (e) {
    return [];
  }
}

// Helper untuk mem-parse tanggal secara fleksibel dari berbagai tipe (Date object atau String)
function parseDate(val) {
  if (!val) return null;
  if (val instanceof Date) {
    return new Date(val.getFullYear(), val.getMonth(), val.getDate());
  }
  const str = val.toString().trim();
  if (!str || str === '-') return null;
  
  // Format yyyy-mm-dd
  let parts = str.split('-');
  if (parts.length === 3) {
    if (parts[0].length === 4) { // yyyy-mm-dd
      return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    } else if (parts[2].length === 4) { // dd-mm-yyyy
      return new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
    }
  }
  
  // Format dd/mm/yyyy atau yyyy/mm/dd
  parts = str.split('/');
  if (parts.length === 3) {
    if (parts[0].length === 4) { // yyyy/mm/dd
      return new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10));
    } else if (parts[2].length === 4) { // dd/mm/yyyy
      return new Date(parseInt(parts[2], 10), parseInt(parts[1], 10) - 1, parseInt(parts[0], 10));
    }
  }
  
  const parsed = Date.parse(str);
  if (!isNaN(parsed)) {
    const d = new Date(parsed);
    return new Date(d.getFullYear(), d.getMonth(), d.getDate());
  }
  return null;
}

// Helper untuk mengonversi tanggal menjadi format ISO string yyyy-MM-dd secara aman & timezone-robust
function formatDateISO(val) {
  if (!val) return "";
  if (val instanceof Date) {
    const y = val.getFullYear();
    const m = String(val.getMonth() + 1).padStart(2, '0');
    const d = String(val.getDate()).padStart(2, '0');
    return `${y}-${m}-${d}`;
  }
  
  const str = val.toString().trim();
  if (!str || str === '-') return "";
  
  // Format yyyy-mm-dd atau dd-mm-yyyy
  let parts = str.split('-');
  if (parts.length === 3) {
    const p0 = parts[0].trim();
    const p1 = parts[1].trim();
    const p2 = parts[2].trim();
    if (p0.length === 4) { // yyyy-mm-dd
      return `${p0}-${p1.padStart(2, '0')}-${p2.padStart(2, '0')}`;
    } else if (p2.length === 4) { // dd-mm-yyyy
      return `${p2}-${p1.padStart(2, '0')}-${p0.padStart(2, '0')}`;
    }
  }
  
  // Format dd/mm/yyyy atau yyyy/mm/dd
  parts = str.split('/');
  if (parts.length === 3) {
    const p0 = parts[0].trim();
    const p1 = parts[1].trim();
    const p2 = parts[2].trim();
    if (p0.length === 4) { // yyyy/mm/dd
      return `${p0}-${p1.padStart(2, '0')}-${p2.padStart(2, '0')}`;
    } else if (p2.length === 4) { // dd/mm/yyyy
      return `${p2}-${p1.padStart(2, '0')}-${p0.padStart(2, '0')}`;
    }
  }
  
  const parsed = Date.parse(str);
  if (!isNaN(parsed)) {
    const d = new Date(parsed);
    const y = d.getFullYear();
    const m = String(d.getMonth() + 1).padStart(2, '0');
    const day = String(d.getDate()).padStart(2, '0');
    return `${y}-${m}-${day}`;
  }
  return "";
}

// Membersihkan berkas spreadsheet rekap temp yang berumur lebih dari 10 menit
function cleanupTempFiles() {
  try {
    const threshold = new Date(Date.now() - 10 * 60 * 1000); // 10 menit lalu
    const files = DriveApp.searchFiles("title contains 'SIPATIH_TEMP_REKAP_' and mimeType = '" + MimeType.GOOGLE_SHEETS + "'");
    while (files.hasNext()) {
      const file = files.next();
      if (file.getLastUpdated() < threshold) {
        file.setTrashed(true);
        console.log('Menghapus file temp rekap lama: ' + file.getName());
      }
    }
  } catch (e) {
    console.error('Error saat membersihkan file temp: ' + e.message);
  }
}

function downloadDataArsip(username, startDate, endDate) {
  try {
    // 1. Bersihkan file temp lama
    cleanupTempFiles();
    
    const ss = getSS();
    
    // 2. Cek izin bidang user
    let userBidang = "Semua";
    let userRole = "Admin";
    if (username) {
      const uSheet = ss.getSheetByName('Pengguna');
      if (uSheet && uSheet.getLastRow() > 1) {
        const uData = uSheet.getRange(2, 2, uSheet.getLastRow() - 1, 5).getValues();
        for (let i = 0; i < uData.length; i++) {
          if (uData[i][0] === username) {
            userRole = uData[i][2];
            userBidang = uData[i][4] || "Semua";
            break;
          }
        }
      }
    }
    
    // Parse tanggal batas jika ada
    let startLimitStr = startDate ? formatDateISO(startDate) : "";
    let endLimitStr = endDate ? formatDateISO(endDate) : "";
    
    // Implementasi fallback satu hari (single-day fallback)
    if (startLimitStr && !endLimitStr) {
      endLimitStr = startLimitStr;
    } else if (endLimitStr && !startLimitStr) {
      startLimitStr = endLimitStr;
    }
    
    // 3. Tarik & Filter Naskah Masuk
    const masukSheet = ss.getSheetByName('Naskah Masuk');
    const masukDataFiltered = [];
    let masukHeaders = [];
    if (masukSheet && masukSheet.getLastRow() > 0) {
      const allRows = masukSheet.getDataRange().getValues();
      masukHeaders = allRows[0];
      const dataRows = allRows.slice(1);
      
      dataRows.forEach(row => {
        if (!row[0] && !row[3]) return; // Lewati baris kosong
        
        // Filter Bidang
        const b = row[9] ? row[9].toString().trim() : "";
        if (userRole !== "Admin" && userRole !== "Kepala Desa" && userBidang !== "Semua" && b !== userBidang) {
          return; // Bidang tidak cocok
        }
        
        // Filter Tanggal Surat (row[4] is Tanggal Surat) - string-based comparison
        const docDateStr = formatDateISO(row[4]);
        if (docDateStr) {
          if (startLimitStr && docDateStr < startLimitStr) return;
          if (endLimitStr && docDateStr > endLimitStr) return;
        } else if (startLimitStr || endLimitStr) {
          return; // Jika ada filter tanggal tapi tanggal surat kosong, lewati
        }
        
        // Format object Date kembali menjadi string/nilai yang bagus untuk excel
        const formattedRow = row.map((cell, idx) => {
          if (cell instanceof Date) {
            return Utilities.formatDate(cell, 'Asia/Jakarta', 'yyyy-MM-dd');
          }
          return cell;
        });
        
        masukDataFiltered.push(formattedRow);
      });
    }
    
    // 4. Tarik & Filter Naskah Keluar
    const keluarSheet = ss.getSheetByName('Naskah Keluar');
    const keluarDataFiltered = [];
    let keluarHeaders = [];
    if (keluarSheet && keluarSheet.getLastRow() > 0) {
      const allRows = keluarSheet.getDataRange().getValues();
      keluarHeaders = allRows[0];
      const dataRows = allRows.slice(1);
      
      dataRows.forEach(row => {
        if (!row[0] && !row[3]) return; // Lewati baris kosong
        
        // Filter Bidang
        const b = row[8] ? row[8].toString().trim() : "";
        if (userRole !== "Admin" && userRole !== "Kepala Desa" && userBidang !== "Semua" && b !== userBidang) {
          return; // Bidang tidak cocok
        }
        
        // Filter Tanggal Surat (row[4] is Tanggal Surat) - string-based comparison
        const docDateStr = formatDateISO(row[4]);
        if (docDateStr) {
          if (startLimitStr && docDateStr < startLimitStr) return;
          if (endLimitStr && docDateStr > endLimitStr) return;
        } else if (startLimitStr || endLimitStr) {
          return; // Jika ada filter tanggal tapi tanggal surat kosong, lewati
        }
        
        // Format object Date kembali
        const formattedRow = row.map((cell, idx) => {
          if (cell instanceof Date) {
            return Utilities.formatDate(cell, 'Asia/Jakarta', 'yyyy-MM-dd');
          }
          return cell;
        });
        
        keluarDataFiltered.push(formattedRow);
      });
    }
    
    // 5. Buat Spreadsheet Temp Baru
    const timestampStr = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyyMMdd_HHmmss');
    const tempFileName = "SIPATIH_TEMP_REKAP_" + timestampStr;
    const tempSS = SpreadsheetApp.create(tempFileName);
    
    // Tulis Naskah Masuk
    const tempMasukSheet = tempSS.getActiveSheet();
    tempMasukSheet.setName("Naskah Masuk");
    if (masukHeaders.length > 0) {
      tempMasukSheet.appendRow(masukHeaders);
      if (masukDataFiltered.length > 0) {
        tempMasukSheet.getRange(2, 1, masukDataFiltered.length, masukHeaders.length).setValues(masukDataFiltered);
      }
      tempMasukSheet.autoResizeColumns(1, masukHeaders.length);
    }
    
    // Tulis Naskah Keluar
    if (keluarHeaders.length > 0) {
      const tempKeluarSheet = tempSS.insertSheet("Naskah Keluar");
      tempKeluarSheet.appendRow(keluarHeaders);
      if (keluarDataFiltered.length > 0) {
        tempKeluarSheet.getRange(2, 1, keluarDataFiltered.length, keluarHeaders.length).setValues(keluarDataFiltered);
      }
      tempKeluarSheet.autoResizeColumns(1, keluarHeaders.length);
    }
    
    // 6. Set Sharing file agar bisa diunduh
    try {
      const file = DriveApp.getFileById(tempSS.getId());
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (sharingErr) {
      console.warn("Gagal setSharing pada spreadsheet temp: " + sharingErr.message);
    }
    
    // 7. Log Aktivitas
    let detail = 'Mendownload rekap data arsip format Excel';
    if (startDate && endDate) {
      detail += ` (${startDate} s/d ${endDate})`;
    } else if (startDate) {
      detail += ` (Mulai ${startDate})`;
    } else if (endDate) {
      detail += ` (Hingga ${endDate})`;
    }
    detail += ` - Berhasil memfilter ${masukDataFiltered.length} Naskah Masuk & ${keluarDataFiltered.length} Naskah Keluar.`;
    
    logActivity(username || 'Sistem', 'Download Excel', detail);
    
    // 8. Kembalikan URL export
    return "https://docs.google.com/spreadsheets/d/" + tempSS.getId() + "/export?format=xlsx";
    
  } catch (e) {
    console.error("Error pada downloadDataArsip: " + e.message);
    throw new Error("Gagal membuat rekap Excel: " + e.message);
  }
}

// ----------------------------------------------------
// SAVE / UPDATE FUNCTIONS (Flowchart Implementation)
// ----------------------------------------------------
function saveArsip(type, data, username) {
  try {
    const sheetName = type === 'Masuk' ? 'Naskah Masuk' : 'Naskah Keluar';
    const sheet = getSS().getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Sheet ' + sheetName + ' tidak ditemukan.' };

    const tglSuratObj = parseDate(data.tglSurat) || "-";
    const tglTerimaObj = type === 'Masuk' ? (parseDate(data.tglTerima) || "-") : "";

    const row = type === 'Masuk' ? 
      [data.noIndeks || '-', data.noAgenda, data.kodeKlasifikasi, data.noSurat, tglSuratObj, tglTerimaObj, data.asalInstansi, data.perihal, data.sifatSurat, data.bidang, data.disposisi || '-', data.statusArsip || 'Aktif', data.keterangan || '-', username, data.linkFile || '-', data.jenisSurat || '-', data.keperluan || '-', data.tandaTangan || '-'] :
      [data.noIndeks || '-', data.noAgenda, data.kodeKlasifikasi, data.noSurat, tglSuratObj, data.tujuanInstansi, data.perihal, data.sifatSurat, data.bidang, data.statusArsip || 'Aktif', data.keterangan || '-', username, data.linkFile || '-', data.jenisSurat || '-', data.keperluan || '-', data.tandaTangan || '-'];

    sheet.appendRow(row);
    logActivity(username, 'Simpan Naskah ' + type, 'Berhasil menyimpan naskah: ' + data.noSurat);
    createNotification(type, data, username);
    
    // Pembuatan naskah keluar otomatis jika dicentang
    if (type === 'Masuk' && data.buatKeluarOtomatis) {
      const kSheet = getSS().getSheetByName('Naskah Keluar');
      if (kSheet) {
        const nextKeluarAgenda = getNextNoAgenda('Keluar');
        const keluarRow = [
          data.noIndeks || '-',
          nextKeluarAgenda,
          data.kodeKlasifikasi,
          data.noSurat,
          tglSuratObj,
          data.asalInstansi, // Tujuan sama dengan asal
          data.perihal,
          data.sifatSurat,
          data.bidang,
          data.statusArsip || 'Aktif',
          (data.keterangan || '-') + ' (Dibuat otomatis dari Naskah Masuk)',
          username,
          data.linkFile || '-',
          data.jenisSurat || '-',
          data.keperluan || '-',
          data.tandaTangan || '-'
        ];
        kSheet.appendRow(keluarRow);
        logActivity(username, 'Simpan Naskah Keluar (Otomatis)', 'Berhasil menyimpan naskah keluar otomatis untuk naskah: ' + data.noSurat);
        createNotification('Keluar', { noSurat: data.noSurat, perihal: data.perihal, bidang: data.bidang }, username);
      }
    }
    
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

/**
 * Consolidated function to upload file and save archive data in one request.
 * File upload errors are NON-FATAL: data is always saved to sheet.
 * Returns { success: true, warning: '...' } if upload fails but data saved.
 */
function saveArsipWithFile(type, data, username, base64Data, fileName) {
  let fileUrl = data.linkFile || '';
  let uploadWarning = '';

  // === STEP 1: Attempt File Upload (non-fatal) ===
  if (base64Data && fileName) {
    try {
      console.log('Memulai upload file untuk ' + type + ': ' + fileName);
      
      // Determine target folder based on type
      const folderId = type === 'Masuk' ? FOLDER_MASUK_ID : FOLDER_KELUAR_ID;
      const folder = DriveApp.getFolderById(folderId);
      
      const contentType = base64Data.substring(5, base64Data.indexOf(';'));
      const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
      const blob = Utilities.newBlob(bytes, contentType, fileName);
      
      const file = folder.createFile(blob);
      
      // Try to set sharing, but don't fail if permission denied
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (sharingErr) {
        console.warn('setSharing gagal (file tetap terupload): ' + sharingErr.message);
      }
      
      fileUrl = file.getUrl();
      console.log('Upload berhasil: ' + fileUrl);
      
    } catch (uploadErr) {
      // Upload failed — log it but CONTINUE saving data
      console.error('Upload file gagal (data tetap disimpan): ' + uploadErr.message);
      uploadWarning = 'File tidak dapat diunggah ke Drive (' + uploadErr.message + '). Data tetap tersimpan.';
    }
  }

  // === STEP 2: Save data to Sheet (always executed) ===
  try {
    const sheetName = type === 'Masuk' ? 'Naskah Masuk' : 'Naskah Keluar';
    const sheet = getSS().getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Sheet ' + sheetName + ' tidak ditemukan.' };

    const tglSuratObj = parseDate(data.tglSurat) || "-";
    const tglTerimaObj = type === 'Masuk' ? (parseDate(data.tglTerima) || "-") : "";

    const row = type === 'Masuk'
      ? [data.noIndeks || '-', data.noAgenda, data.kodeKlasifikasi, data.noSurat, tglSuratObj, tglTerimaObj, data.instansi, data.perihal, data.sifatSurat, data.bidang, data.disposisi || '-', data.statusArsip || 'Aktif', data.keterangan || '-', username, fileUrl, data.jenisSurat || '-', data.keperluan || '-', data.tandaTangan || '-']
      : [data.noIndeks || '-', data.noAgenda, data.kodeKlasifikasi, data.noSurat, tglSuratObj, data.instansi, data.perihal, data.sifatSurat, data.bidang, data.statusArsip || 'Aktif', data.keterangan || '-', username, fileUrl, data.jenisSurat || '-', data.keperluan || '-', data.tandaTangan || '-'];

    sheet.appendRow(row);
    logActivity(username, 'Simpan Naskah ' + type, 'Berhasil menyimpan naskah: ' + data.noSurat + (fileUrl ? ' dengan file' : ' (tanpa file)'));
    createNotification(type, data, username);

    // Pembuatan naskah keluar otomatis jika dicentang
    if (type === 'Masuk' && data.buatKeluarOtomatis) {
      const kSheet = getSS().getSheetByName('Naskah Keluar');
      if (kSheet) {
        const nextKeluarAgenda = getNextNoAgenda('Keluar');
        const keluarRow = [
          data.noIndeks || '-',
          nextKeluarAgenda,
          data.kodeKlasifikasi,
          data.noSurat,
          tglSuratObj,
          data.instansi, // Tujuan sama dengan asal
          data.perihal,
          data.sifatSurat,
          data.bidang,
          data.statusArsip || 'Aktif',
          (data.keterangan || '-') + ' (Dibuat otomatis dari Naskah Masuk)',
          username,
          fileUrl,
          data.jenisSurat || '-',
          data.keperluan || '-',
          data.tandaTangan || '-'
        ];
        kSheet.appendRow(keluarRow);
        logActivity(username, 'Simpan Naskah Keluar (Otomatis)', 'Berhasil menyimpan naskah keluar otomatis untuk naskah: ' + data.noSurat);
        createNotification('Keluar', { noSurat: data.noSurat, perihal: data.perihal, bidang: data.bidang }, username);
      }
    }

    return { success: true, url: fileUrl, warning: uploadWarning };
    
  } catch (saveErr) {
    console.error('Gagal menyimpan ke sheet: ' + saveErr.message);
    return { success: false, message: 'Gagal menyimpan data: ' + saveErr.message };
  }
}

function updateArsipWithFile(type, oldNoSurat, data, username, base64Data, fileName) {
  let fileUrl = data.linkFile || '';
  let uploadWarning = '';

  // === STEP 1: Attempt File Upload if new file provided ===
  if (base64Data && fileName) {
    try {
      console.log('Memulai upload file baru untuk ' + type + ': ' + fileName);
      const folderId = type === 'Masuk' ? FOLDER_MASUK_ID : FOLDER_KELUAR_ID;
      const folder = DriveApp.getFolderById(folderId);
      
      const contentType = base64Data.substring(5, base64Data.indexOf(';'));
      const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
      const blob = Utilities.newBlob(bytes, contentType, fileName);
      
      const file = folder.createFile(blob);
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (sharingErr) {
        console.warn('setSharing gagal: ' + sharingErr.message);
      }
      fileUrl = file.getUrl();
    } catch (uploadErr) {
      console.error('Upload file gagal: ' + uploadErr.message);
      uploadWarning = 'File tidak dapat diunggah (' + uploadErr.message + '). Menggunakan file lama.';
    }
  }

  // === STEP 2: Find row and Update ===
  try {
    const sheetName = type === 'Masuk' ? 'Naskah Masuk' : 'Naskah Keluar';
    const sheet = getSS().getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Sheet ' + sheetName + ' tidak ditemukan.' };
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'Tidak ada data arsip.' };
    
    const allNoSurat = sheet.getRange(2, 4, lastRow - 1, 1).getValues(); // Kolom D: Nomor Surat
    let rowIndex = -1;
    for (let i = 0; i < allNoSurat.length; i++) {
      if (allNoSurat[i][0] === oldNoSurat) {
        rowIndex = i + 2;
        break;
      }
    }
    
    if (rowIndex === -1) return { success: false, message: 'Naskah dengan nomor surat lama tidak ditemukan.' };
    
    const tglSuratObj = parseDate(data.tglSurat) || "-";
    const tglTerimaObj = type === 'Masuk' ? (parseDate(data.tglTerima) || "-") : "";

    // Update the row values
    const rowRange = sheet.getRange(rowIndex, 1, 1, type === 'Masuk' ? 18 : 16);
    const rowValues = type === 'Masuk'
      ? [[data.noIndeks || '-', data.noAgenda, data.kodeKlasifikasi, data.noSurat, tglSuratObj, tglTerimaObj, data.instansi, data.perihal, data.sifatSurat, data.bidang, data.disposisi || '-', data.statusArsip || 'Aktif', data.keterangan || '-', username, fileUrl, data.jenisSurat || '-', data.keperluan || '-', data.tandaTangan || '-']]
      : [[data.noIndeks || '-', data.noAgenda, data.kodeKlasifikasi, data.noSurat, tglSuratObj, data.instansi, data.perihal, data.sifatSurat, data.bidang, data.statusArsip || 'Aktif', data.keterangan || '-', username, fileUrl, data.jenisSurat || '-', data.keperluan || '-', data.tandaTangan || '-']];
    
    rowRange.setValues(rowValues);
    logActivity(username, 'Update Naskah ' + type, 'Berhasil memperbarui naskah: ' + data.noSurat);
    
    return { success: true, url: fileUrl, warning: uploadWarning };
  } catch (err) {
    console.error('Gagal memperbarui ke sheet: ' + err.message);
    return { success: false, message: 'Gagal memperbarui data: ' + err.message };
  }
}

function saveUser(data, adminUsername) {
  try {
    const sheet = getSS().getSheetByName('Pengguna');
    const id = sheet.getLastRow();
    sheet.appendRow([id, data.username, data.password, data.role, 'Aktif', data.bidang || 'Semua']);
    logActivity(adminUsername, 'Tambah Pengguna', 'Menambahkan user baru: ' + data.username);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateUser(oldUsername, data, adminUsername) {
  try {
    const sheet = getSS().getSheetByName('Pengguna');
    const dbData = sheet.getDataRange().getValues();
    for (let i = 1; i < dbData.length; i++) {
      if (dbData[i][1] === oldUsername) {
        sheet.getRange(i + 1, 2).setValue(data.username);
        sheet.getRange(i + 1, 3).setValue(data.password);
        sheet.getRange(i + 1, 4).setValue(data.role);
        sheet.getRange(i + 1, 5).setValue(data.status || 'Aktif');
        sheet.getRange(i + 1, 6).setValue(data.bidang || 'Semua');
        logActivity(adminUsername, 'Update Pengguna', 'Mengupdate user: ' + oldUsername + ' menjadi ' + data.username);
        return { success: true };
      }
    }
    return { success: false, message: 'User tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function deleteUser(usernameToDelete, adminUsername) {
  try {
    if (usernameToDelete === adminUsername) {
      return { success: false, message: 'Tidak dapat menghapus akun Anda sendiri yang sedang aktif!' };
    }
    const sheet = getSS().getSheetByName('Pengguna');
    const dbData = sheet.getDataRange().getValues();
    for (let i = 1; i < dbData.length; i++) {
      if (dbData[i][1] === usernameToDelete) {
        sheet.deleteRow(i + 1);
        logActivity(adminUsername, 'Hapus Pengguna', 'Menghapus user: ' + usernameToDelete);
        return { success: true };
      }
    }
    return { success: false, message: 'User tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function saveBidang(data, username) {
  try {
    const sheet = getSS().getSheetByName('Klasifikasi Arsip');
    sheet.appendRow([data.nama, data.kode.toUpperCase(), '', '-', '']); // Col A: Nama Bidang, Col B: Kode, Col C: Baris, Col D: Jenis Arsip, Col E: Retensi
    logActivity(username, 'Tambah Bidang', 'Menambahkan bidang baru: ' + data.nama);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateProfile(newPassword, username) {
  try {
    const sheet = getSS().getSheetByName('Pengguna');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][1] === username) {
        sheet.getRange(i + 1, 3).setValue(newPassword);
        logActivity(username, 'Update Profil', 'Berhasil memperbarui password');
        return { success: true };
      }
    }
    return { success: false, message: 'User tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function uploadFile(base64Data, fileName) {
  try {
    console.log('Memulai upload file: ' + fileName);
    const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    const contentType = base64Data.substring(5, base64Data.indexOf(';'));
    const bytes = Utilities.base64Decode(base64Data.split(',')[1]);
    const blob = Utilities.newBlob(bytes, contentType, fileName);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    console.log('Upload berhasil: ' + file.getUrl());
    return { success: true, url: file.getUrl() };
  } catch (e) {
    console.error('Error uploadFile: ' + e.message);
    return { success: false, message: e.message };
  }
}
function deleteBidang(nama, username) {
  try {
    const sheet = getSS().getSheetByName('Klasifikasi Arsip');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === nama) { // Col A is index 0
        sheet.deleteRow(i + 1);
        logActivity(username, 'Hapus Bidang', 'Menghapus bidang: ' + nama);
        return { success: true };
      }
    }
    return { success: false, message: 'Bidang tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function updateBidang(oldNama, newData, username) {
  try {
    const sheet = getSS().getSheetByName('Klasifikasi Arsip');
    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (data[i][0] === oldNama) { // Col A is index 0
        sheet.getRange(i + 1, 1).setValue(newData.nama);
        sheet.getRange(i + 1, 2).setValue(newData.kode.toUpperCase());
        logActivity(username, 'Update Bidang', 'Mengupdate bidang: ' + oldNama + ' menjadi ' + newData.nama);
        return { success: true };
      }
    }
    return { success: false, message: 'Bidang tidak ditemukan' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// Helper to initialize custom/required bidang list directly
function initializeBidangBaru() {
  try {
    const ss = getSS();
    let sh = ss.getSheetByName('Klasifikasi Arsip');
    if (!sh) {
      sh = ss.insertSheet('Klasifikasi Arsip');
    }
    
    // Clear and reset the classification sheet with the new column layout and user-defined classifications
    sh.clear();
    sh.appendRow(['Nama Bidang', 'Kode', 'Baris', 'Jenis Arsip', 'Retensi Aktif / Inaktif']);
    
    const klasifikasiAwal = [
      ['Pemerintahan', 'PEM', '2', '-', ''],
      ['', 'PEM.01', '', 'Surat Umum', ''],
      ['', 'PEM.02', '', 'Sosialisasi dan Undangan', ''],
      
      ['Kesejahteraan', 'KES', '3', '-', ''],
      ['', 'KES.01', '', 'Perkawinan', ''],
      ['', 'KES.02', '', 'Kesejahteraan Sosial', ''],
      
      ['Kependudukan', 'KEP', '4', '-', ''],
      ['', 'KEP.01', '', 'Pindah Keluar / Masuk', ''],
      ['', 'KEP.02', '', 'Kematian / Akte Kematian', ''],
      ['', 'KEP.03', '', 'Kelahiran / Akte Lahir', ''],
      ['', 'KEP.04', '', 'Pembuatan / Perubahan KK', ''],
      
      ['Pertanahan', 'PER', '5', '-', ''],
      ['', 'PER.01', '', 'Mutasi Tanah (Jual beli / waris/ hibah)', ''],
      ['', 'PER.02', '', 'Mutasi PBB', ''],
      ['', 'PER.03', '', 'Keterangan Tanah Lainnya', ''],
      
      ['Umum', 'UMU', '6', '-', ''],
      ['', 'UMU.01', '', 'Surat Lainnya', '']
    ];
    klasifikasiAwal.forEach(row => sh.appendRow(row));
    
    logActivity('Sistem', 'Reset Bidang', 'Inisialisasi bidang dokumen lengkap');
    return true;
  } catch (e) {
    console.error('Gagal inisialisasi bidang baru: ' + e.message);
    return false;
  }
}

function checkAndInitializeBidang() {
  try {
    const props = PropertiesService.getScriptProperties();
    const initialized = props.getProperty('bidang_initialized_v5');
    if (initialized !== 'true') {
      const success = initializeBidangBaru();
      if (success) {
        props.setProperty('bidang_initialized_v5', 'true');
      }
    }
  } catch (e) {
    console.error('Error di checkAndInitializeBidang: ' + e.message);
  }
}

// ----------------------------------------------------
// Fitur Lonceng Notifikasi Realtime
// ----------------------------------------------------
function createNotification(type, data, username) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName('Notifikasi');
    if (!sheet) return;
    
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss');
    const id = sheet.getLastRow();
    
    const isMasuk = type === 'Masuk';
    const tipeText = isMasuk ? 'Surat Masuk' : 'Surat Keluar';
    
    const pesan = `Naskah ${tipeText} baru nomor ${data.noSurat} perihal "${data.perihal}" bidang ${data.bidang} telah diinput oleh ${username}.`;
    
    sheet.appendRow([id, timestamp, username, data.bidang, pesan, 'Belum Dibaca']);
  } catch (e) {
    console.error('Gagal membuat notifikasi: ' + e.message);
  }
}

function getNotifications(username) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName('Notifikasi');
    if (!sheet || sheet.getLastRow() <= 1) return [];
    
    // Get user details
    let userBidang = "Semua";
    let userRole = "Admin";
    if (username) {
      const uSheet = ss.getSheetByName('Pengguna');
      if (uSheet && uSheet.getLastRow() > 1) {
        const uData = uSheet.getRange(2, 2, uSheet.getLastRow() - 1, 5).getValues();
        for (let i = 0; i < uData.length; i++) {
          if (uData[i][0] === username) {
            userRole = uData[i][2];
            userBidang = uData[i][4] || "Semua";
            break;
          }
        }
      }
    }

    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
    const result = [];
    
    data.forEach((row, index) => {
      const id = row[0];
      const timestamp = row[1] instanceof Date ? Utilities.formatDate(row[1], 'Asia/Jakarta', 'yyyy-MM-dd HH:mm:ss') : row[1];
      const targetUser = row[2];
      const targetBidang = row[3];
      const pesan = row[4];
      const status = row[5];
      
      // Admin/Kepala Desa melihat semua notifikasi
      if (userRole === "Admin" || userRole === "Kepala Desa") {
        result.push({ id, timestamp, pesan, status });
      } 
      // User biasa melihat notifikasi yang sesuai bidang mereka atau yang diinput oleh mereka sendiri
      else if (targetBidang === userBidang || targetUser === username) {
        result.push({ id, timestamp, pesan, status });
      }
    });
    
    return result.reverse().slice(0, 30); // Kembalikan 30 notifikasi terbaru, terbalik (terbaru pertama)
  } catch (e) {
    return [];
  }
}

function markNotificationsAsRead(username) {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName('Notifikasi');
    if (!sheet || sheet.getLastRow() <= 1) return { success: true };
    
    // Get user details
    let userBidang = "Semua";
    let userRole = "Admin";
    if (username) {
      const uSheet = ss.getSheetByName('Pengguna');
      if (uSheet && uSheet.getLastRow() > 1) {
        const uData = uSheet.getRange(2, 2, uSheet.getLastRow() - 1, 5).getValues();
        for (let i = 0; i < uData.length; i++) {
          if (uData[i][0] === username) {
            userRole = uData[i][2];
            userBidang = uData[i][4] || "Semua";
            break;
          }
        }
      }
    }

    const lastRow = sheet.getLastRow();
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    
    for (let i = 0; i < data.length; i++) {
      const targetUser = data[i][2];
      const targetBidang = data[i][3];
      const status = data[i][5];
      
      if (status === 'Belum Dibaca') {
        if (userRole === "Admin" || userRole === "Kepala Desa" || targetBidang === userBidang || targetUser === username) {
          sheet.getRange(i + 2, 6).setValue('Dibaca');
        }
      }
    }
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ----------------------------------------------------
// Fitur Retensi Surat (Aktif / Non Aktif)
// ----------------------------------------------------
function setArsipStatus(type, noSurat, status, username) {
  try {
    const sheetName = type === 'Masuk' ? 'Naskah Masuk' : 'Naskah Keluar';
    const sheet = getSS().getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Sheet ' + sheetName + ' tidak ditemukan.' };
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: false, message: 'Tidak ada data arsip.' };
    
    const data = sheet.getRange(2, 4, lastRow - 1, 1).getValues(); // Kolom D: Nomor Surat
    const statusColIndex = type === 'Masuk' ? 12 : 10; // Kolom L untuk Masuk, Kolom J untuk Keluar
    
    for (let i = 0; i < data.length; i++) {
      if (data[i][0] === noSurat) {
        sheet.getRange(i + 2, statusColIndex).setValue(status);
        logActivity(username, 'Update Status Arsip', `Mengubah status naskah ${type} nomor ${noSurat} menjadi ${status}`);
        return { success: true };
      }
    }
    return { success: false, message: 'Naskah tidak ditemukan.' };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

// ----------------------------------------------------
// Fitur Nomor Agenda Otomatis
// ----------------------------------------------------
function getNextNoAgenda(type) {
  try {
    const sheetName = type === 'Masuk' ? 'Naskah Masuk' : 'Naskah Keluar';
    const sheet = getSS().getSheetByName(sheetName);
    const lastRow = sheet ? sheet.getLastRow() : 1;
    const nextSeq = lastRow > 0 ? lastRow : 1; // Jika hanya header (lastRow=1), data=0, seq berikutnya=1
    
    const paddedSeq = nextSeq.toString().padStart(3, '0');
    
    const now = new Date();
    const month = (now.getMonth() + 1).toString().padStart(2, '0');
    const year = now.getFullYear();
    
    const code = type === 'Masuk' ? 'SM' : 'SK';
    
    return `${paddedSeq}/${code}/${month}/${year}`;
  } catch (e) {
    return "";
  }
}

// ----------------------------------------------------
// Fitur Kode Klasifikasi Otomatis
// ----------------------------------------------------
function getKlasifikasiOptions() {
  try {
    const ss = getSS();
    const sheet = ss.getSheetByName('Klasifikasi Arsip');
    if (!sheet) return [];
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    
    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    const result = [];
    
    let currentBidang = "";
    
    data.forEach(row => {
      const bName = row[0] ? row[0].toString().trim() : "";
      const code = row[1] ? row[1].toString().trim() : "";
      const baris = row[2] ? row[2].toString().trim() : "";
      const jenisArsip = row[3] ? row[3].toString().trim() : "";
      
      // Update bidang induk jika kodenya 3 karakter (PEM, KES, KEP, PER, UMU)
      if (code.length === 3 && bName) {
        currentBidang = bName;
      }
      
      // Ambil jenis arsip jika kodenya 6 karakter (PEM.01, KES.01, dll.)
      if (code.length === 6 && jenisArsip && jenisArsip !== "-") {
        result.push({
          bidang: currentBidang || bName,
          kode: code,
          baris: baris,
          jenisArsip: jenisArsip
        });
      }
    });
    
    return result;
  } catch (e) {
    return [];
  }
}

function getNextKodeKlasifikasi(bidang, jenisArsip) {
  try {
    const ss = getSS();
    const kOptions = getKlasifikasiOptions();
    let parentKode = "";
    
    for (let i = 0; i < kOptions.length; i++) {
      if (kOptions[i].bidang.toLowerCase() === bidang.toLowerCase() && kOptions[i].jenisArsip.toLowerCase() === jenisArsip.toLowerCase()) {
        parentKode = kOptions[i].kode;
        break;
      }
    }
    
    if (!parentKode) return "";
    
    let docCount = 0;
    
    const mSheet = ss.getSheetByName('Naskah Masuk');
    if (mSheet && mSheet.getLastRow() > 1) {
      const mCodes = mSheet.getRange(2, 3, mSheet.getLastRow() - 1, 1).getValues();
      mCodes.forEach(row => {
        const codeVal = row[0] ? row[0].toString().trim() : "";
        if (codeVal.startsWith(parentKode)) {
          docCount++;
        }
      });
    }
    
    const kSheetDoc = ss.getSheetByName('Naskah Keluar');
    if (kSheetDoc && kSheetDoc.getLastRow() > 1) {
      const kCodes = kSheetDoc.getRange(2, 3, kSheetDoc.getLastRow() - 1, 1).getValues();
      kCodes.forEach(row => {
        const codeVal = row[0] ? row[0].toString().trim() : "";
        if (codeVal.startsWith(parentKode)) {
          docCount++;
        }
      });
    }
    
    const nextIndex = docCount + 1;
    const paddedIndex = nextIndex.toString().padStart(2, '0');
    
    return `${parentKode}.${paddedIndex}`;
  } catch (e) {
    return "";
  }
}
