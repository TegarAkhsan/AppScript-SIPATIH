const SPREADSHEET_ID = '1UFEXxqUh5OKyPVEQh4fmPo0Y2i3FdOd7xXLlyBLN5XU';
const FOLDER_MASUK_ID = '1nSmgCSCslt3_cG3AS4S63gqSnCbEQNNe';
const FOLDER_KELUAR_ID = '1P5i0zcrBWa97aAY5fkjMPsqRTpB_bZGX';
const DRIVE_FOLDER_ID = '1GCrXpEb70cnsTvyTpnglyaSqyto0m_PC'; // Default fallback

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('SIPATIH - Sistem Arsip Digital Desa Kepatihan')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
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
      sh.appendRow(['ID', 'Username', 'Password', 'Role', 'Status']);
    }
    // Pastikan tidak menduplikasi jika sudah ada data tapi kurang dari 2 baris
    if (sh.getLastRow() === 1) {
      sh.appendRow(['1', 'admin', 'admin', 'Admin', 'Aktif']);
      sh.appendRow(['2', 'user', 'user', 'Perangkat Desa', 'Aktif']);
    }
  }
  
  // Aktivitas
  sh = getOrCreateSheet('Aktivitas');
  if(sh.getLastRow() === 0) {
    sh.appendRow(['Timestamp', 'Username', 'Aktivitas', 'Detail']);
  }

  // Naskah Masuk
  sh = getOrCreateSheet('Naskah Masuk');
  if(sh.getLastRow() === 0) {
    sh.appendRow(['Nomor Indeks', 'Nomor Agenda', 'Kode Klasifikasi', 'Nomor Surat', 'Tanggal Surat', 'Tanggal Terima', 'Asal Instansi', 'Perihal', 'Sifat Surat', 'Bidang', 'Disposisi', 'Status Arsip', 'Keterangan', 'Diinput Oleh', 'Link File']);
  }

  // Naskah Keluar
  sh = getOrCreateSheet('Naskah Keluar');
  if(sh.getLastRow() === 0) {
    sh.appendRow(['Nomor Indeks', 'Nomor Agenda', 'Kode Klasifikasi', 'Nomor Surat', 'Tanggal Surat', 'Tujuan Instansi', 'Perihal', 'Sifat Surat', 'Bidang', 'Status Arsip', 'Keterangan', 'Diinput Oleh', 'Link File']);
  }

  // Klasifikasi Arsip (Sesuai Permintaan User)
  sh = getOrCreateSheet('Klasifikasi Arsip');
  if(sh.getLastRow() <= 1) {
    if(sh.getLastRow() === 0) {
      sh.appendRow(['Baris', 'Kode', 'Nama Bidang', 'Jenis Arsip', 'Retensi Aktif', 'Retensi Inaktif']);
    }
    if(sh.getLastRow() === 1) {
      const klasifikasiAwal = [
        ['2', '470.1', 'Kependudukan', 'Kelahiran / Akte Lahir', '', ''],
        ['3', '470.2', 'Kependudukan', 'Kematian / Akte Kematian', '', ''],
        ['4', '470.3', 'Kependudukan', 'Perkawinan / Akte Nikah', '', ''],
        ['5', '470.4', 'Kependudukan', 'Perceraian', '', ''],
        ['6', '470.5', 'Kependudukan', 'KTP / KK / Domisili', '', ''],
        ['7', '590.1', 'Pertanahan', 'Setifikat / Kepemilikan', '', ''],
        ['8', '590.2', 'Pertanahan', 'Hibah / Waris Tanah', '', '']
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
        return { success: true, role: 'Admin', username: 'admin' };
      }
      return { success: false, message: 'Tidak ada data pengguna dalam sistem.' };
    }

    const data = sheet.getRange(2, 1, lastRow - 1, 5).getValues();
    
    for (let i = 0; i < data.length; i++) {
      let row = data[i];
      let dbUsername = row[1];
      let dbPassword = row[2];
      let role = row[3];
      let status = row[4];
      
      if (dbUsername === username && dbPassword === password) {
        if (status && status.toString().toLowerCase() !== 'aktif') {
          return { success: false, message: 'Akun Anda dinonaktifkan!' };
        }
        logActivity(username, 'Login', 'Berhasil login ke sistem');
        return { success: true, role: role, username: username };
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
function getDashboardData() {
  try {
    setupHeaders(); // Pastikan header dan data awal tersedia
    const ss = getSS();
    const arsipMasukSheet = ss.getSheetByName('Naskah Masuk');
    const arsipKeluarSheet = ss.getSheetByName('Naskah Keluar');
    const klasifikasiSheet = ss.getSheetByName('Klasifikasi Arsip');
    
    let masukCount = 0, keluarCount = 0;
    if (arsipMasukSheet && arsipMasukSheet.getLastRow() > 1) {
      masukCount = arsipMasukSheet.getLastRow() - 1;
    }
    if (arsipKeluarSheet && arsipKeluarSheet.getLastRow() > 1) {
      keluarCount = arsipKeluarSheet.getLastRow() - 1;
    }
    const totalArsip = masukCount + keluarCount;

    let bidangMap = {};
    if (klasifikasiSheet && klasifikasiSheet.getLastRow() > 1) {
      const klasData = klasifikasiSheet.getRange(2, 3, klasifikasiSheet.getLastRow() - 1, 1).getValues(); 
      klasData.forEach(row => {
        let b = row[0];
        if(b) {
          b = b.toString().trim();
          if(b !== "") bidangMap[b] = 0;
        }
      });
    }

    if (masukCount > 0) {
      const masukBidang = arsipMasukSheet.getRange(2, 10, masukCount, 1).getValues(); // J column
      masukBidang.forEach(row => {
        let b = row[0] ? row[0].toString().trim() : "";
        if(b) bidangMap[b] = (bidangMap[b] || 0) + 1;
      });
    }
    
    if (keluarCount > 0) {
      const keluarBidang = arsipKeluarSheet.getRange(2, 9, keluarCount, 1).getValues(); // I column
      keluarBidang.forEach(row => {
        let b = row[0] ? row[0].toString().trim() : "";
        if(b) bidangMap[b] = (bidangMap[b] || 0) + 1;
      });
    }

    const bidangList = Object.keys(bidangMap).map(k => ({
      kode: k.substring(0, 3).toUpperCase(), 
      nama: k,
      jumlah: bidangMap[k]
    }));

    return {
      totalArsip: totalArsip,
      totalBidang: Object.keys(bidangMap).length,
      arsipAktif: totalArsip, // Anggap semua aktif untuk dashboard sederhana
      arsipInaktif: 0,
      bidangList: bidangList
    };
  } catch (e) {
    return { error: true, message: e.message };
  }
}

function getArsipData() {
  try {
    const ss = getSS();
    let result = [];
    
    const masukSheet = ss.getSheetByName('Naskah Masuk');
    if (masukSheet && masukSheet.getLastRow() > 1) {
      const mData = masukSheet.getRange(2, 1, masukSheet.getLastRow() - 1, 15).getValues();
      mData.forEach(row => {
        if(!row[0] && !row[3]) return;
        result.push({
          jenis: 'Masuk',
          noIndeks: row[0],
          noSurat: row[3],
          tglSurat: row[4] instanceof Date ? Utilities.formatDate(row[4], 'Asia/Jakarta', 'dd-MM-yyyy') : (row[4] || '-'),
          perihal: row[7],
          sifatSurat: row[8],
          bidang: row[9],
          asalInstansi: row[6],
          keterangan: row[12],
          linkFile: row[14]
        });
      });
    }

    const keluarSheet = ss.getSheetByName('Naskah Keluar');
    if (keluarSheet && keluarSheet.getLastRow() > 1) {
      const kData = keluarSheet.getRange(2, 1, keluarSheet.getLastRow() - 1, 13).getValues();
      kData.forEach(row => {
        if(!row[0] && !row[3]) return;
        result.push({
          jenis: 'Keluar',
          noIndeks: row[0],
          noSurat: row[3],
          tglSurat: row[4] instanceof Date ? Utilities.formatDate(row[4], 'Asia/Jakarta', 'dd-MM-yyyy') : (row[4] || '-'),
          perihal: row[6],
          sifatSurat: row[7],
          bidang: row[8],
          tujuanInstansi: row[5],
          keterangan: row[10],
          linkFile: row[12]
        });
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
    
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 5).getValues();
    return data.map(row => ({
      id: row[0],
      username: row[1],
      role: row[3],
      status: row[4] || 'Aktif'
    })).filter(u => u.username);
  } catch (e) {
    return [];
  }
}

function downloadDataArsip(username, startDate, endDate) {
  let detail = 'Mendownload rekap data arsip format Excel';
  if (startDate && endDate) detail += ` (${startDate} s/d ${endDate})`;
  logActivity(username || 'Sistem', 'Download Excel', detail);
  return "https://docs.google.com/spreadsheets/d/" + SPREADSHEET_ID + "/export?format=xlsx";
}

function downloadDataZip(username, startDate, endDate) {
  let detail = 'Mencoba download ZIP arsip (Redirect ke Spreadsheet)';
  if (startDate && endDate) detail += ` (${startDate} s/d ${endDate})`;
  logActivity(username || 'Sistem', 'Download ZIP', detail);
  return "https://docs.google.com/spreadsheets/d/" + SPREADSHEET_ID + "/export?format=xlsx";
}

// ----------------------------------------------------
// SAVE / UPDATE FUNCTIONS (Flowchart Implementation)
// ----------------------------------------------------
function saveArsip(type, data, username) {
  try {
    const sheetName = type === 'Masuk' ? 'Naskah Masuk' : 'Naskah Keluar';
    const sheet = getSS().getSheetByName(sheetName);
    if (!sheet) return { success: false, message: 'Sheet ' + sheetName + ' tidak ditemukan.' };

    const row = type === 'Masuk' ? 
      [data.noIndeks, data.noAgenda, data.kodeKlasifikasi, data.noSurat, data.tglSurat, data.tglTerima, data.asalInstansi, data.perihal, data.sifatSurat, data.bidang, data.disposisi, data.statusArsip, data.keterangan, username, data.linkFile] :
      [data.noIndeks, data.noAgenda, data.kodeKlasifikasi, data.noSurat, data.tglSurat, data.tujuanInstansi, data.perihal, data.sifatSurat, data.bidang, data.statusArsip, data.keterangan, username, data.linkFile];

    sheet.appendRow(row);
    logActivity(username, 'Simpan Naskah ' + type, 'Berhasil menyimpan naskah: ' + data.noSurat);
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

    const row = type === 'Masuk'
      ? [data.noIndeks || '-', data.noAgenda, data.kodeKlasifikasi, data.noSurat, data.tglSurat, data.tglTerima, data.instansi, data.perihal, data.sifatSurat, data.bidang, data.disposisi || '-', data.statusArsip || 'Aktif', data.keterangan || '-', username, fileUrl]
      : [data.noIndeks || '-', data.noAgenda, data.kodeKlasifikasi, data.noSurat, data.tglSurat, data.instansi, data.perihal, data.sifatSurat, data.bidang, data.statusArsip || 'Aktif', data.keterangan || '-', username, fileUrl];

    sheet.appendRow(row);
    logActivity(username, 'Simpan Naskah ' + type, 'Berhasil menyimpan naskah: ' + data.noSurat + (fileUrl ? ' dengan file' : ' (tanpa file)'));

    return { success: true, url: fileUrl, warning: uploadWarning };
    
  } catch (saveErr) {
    console.error('Gagal menyimpan ke sheet: ' + saveErr.message);
    return { success: false, message: 'Gagal menyimpan data: ' + saveErr.message };
  }
}

function saveUser(data, adminUsername) {
  try {
    const sheet = getSS().getSheetByName('Pengguna');
    const id = sheet.getLastRow();
    sheet.appendRow([id, data.username, data.password, data.role, 'Aktif']);
    logActivity(adminUsername, 'Tambah Pengguna', 'Menambahkan user baru: ' + data.username);
    return { success: true };
  } catch (e) {
    return { success: false, message: e.message };
  }
}

function saveBidang(data, username) {
  try {
    const sheet = getSS().getSheetByName('Klasifikasi Arsip');
    const baris = sheet.getLastRow();
    sheet.appendRow([baris, data.kode, data.nama, data.jenis || '-', '-', '-']);
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
