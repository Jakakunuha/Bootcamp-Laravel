const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();
const SHEET_NAMES = { 
  USERS: 'users', 
  MODULES: 'modules', 
  LESSONS: 'lessons' 
};

function respond(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function setupDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Buat sheet jika belum ada beserta header kolomnya
  createSheetIfNotExists(ss, SHEET_NAMES.USERS, ['id', 'username', 'password', 'role', 'created_at']);
  createSheetIfNotExists(ss, SHEET_NAMES.MODULES, ['id', 'kategori', 'judul_materi', 'deskripsi', 'urutan', 'created_at']);
  createSheetIfNotExists(ss, SHEET_NAMES.LESSONS, ['id', 'module_id', 'judul_sub', 'konten_html', 'video_url', 'urutan', 'created_at']);

  // Seed default admin & user jika sheet users kosong
  const usersSheet = ss.getSheetByName(SHEET_NAMES.USERS);
  if (usersSheet.getLastRow() <= 1) {
    usersSheet.appendRow([Date.now().toString(), 'admin', 'adminpassword', 'admin', new Date().toISOString()]);
    usersSheet.appendRow([(Date.now() + 1).toString(), 'peserta', 'user123', 'user', new Date().toISOString()]);
  }
}

function createSheetIfNotExists(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
  }
  return sheet;
}

function doGet(e) { 
  setupDatabase();
  return respond({ status: 'success', message: 'GAS Backend Web Materi Pembelajaran Aktif & Terhubung' }); 
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const postData = JSON.parse(e.postData.contents || "{}");
    const action = postData.action;
    const payload = postData.payload || {};

    switch (action) {
      case 'login':    return handleLogin(payload);
      case 'getAll':   return handleGetAll(payload);
      case 'create':   return handleCreate(payload);
      case 'update':   return handleUpdate(payload);
      case 'delete':   return handleDelete(payload);
      case 'upload':   return handleUpload(payload);
      default:         return respond({ status: 'error', message: 'Action tidak valid: ' + action });
    }
  } catch (err) {
    return respond({ status: 'error', message: 'Internal Server Error: ' + err.message });
  } finally {
    lock.releaseLock();
  }
}

function handleLogin(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.USERS);
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return respond({ status: 'error', message: 'Data pengguna tidak ditemukan' });

  const headers = rows[0];
  const idxId = headers.indexOf('id');
  const idxUsername = headers.indexOf('username');
  const idxPassword = headers.indexOf('password');
  const idxRole = headers.indexOf('role');

  for (let i = 1; i < rows.length; i++) {
    const row = rows[i];
    if (row[idxUsername] === payload.username && row[idxPassword] === payload.password) {
      return respond({
        status: 'success',
        message: 'Login berhasil',
        data: {
          id: row[idxId],
          username: row[idxUsername],
          role: row[idxRole]
        }
      });
    }
  }
  return respond({ status: 'error', message: 'Username atau password salah!' });
}

function handleGetAll(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Ambil Data Modules
  const modSheet = ss.getSheetByName(SHEET_NAMES.MODULES);
  const modRows = modSheet.getDataRange().getValues();
  let modules = [];
  if (modRows.length > 1) {
    const modHeaders = modRows[0];
    for (let i = 1; i < modRows.length; i++) {
      let obj = {};
      modHeaders.forEach((h, idx) => { obj[h] = modRows[i][idx]; });
      modules.push(obj);
    }
  }

  // Ambil Data Lessons
  const lesSheet = ss.getSheetByName(SHEET_NAMES.LESSONS);
  const lesRows = lesSheet.getDataRange().getValues();
  let lessons = [];
  if (lesRows.length > 1) {
    const lesHeaders = lesRows[0];
    for (let i = 1; i < lesRows.length; i++) {
      let obj = {};
      lesHeaders.forEach((h, idx) => { obj[h] = lesRows[i][idx]; });
      lessons.push(obj);
    }
  }

  return respond({
    status: 'success',
    message: 'Data kurikulum berhasil dimuat',
    data: { modules, lessons }
  });
}

function handleCreate(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const type = payload.type; // 'module' atau 'lesson'
  
  if (type === 'module') {
    const sheet = ss.getSheetByName(SHEET_NAMES.MODULES);
    const id = "mod-" + Date.now();
    const kategori = payload.kategori || "";
    const judul_materi = payload.judul_materi || "";
    const deskripsi = payload.deskripsi || "";
    const urutan = payload.urutan || "1";
    const created_at = new Date().toISOString();

    sheet.appendRow([id, kategori, judul_materi, deskripsi, urutan, created_at]);
    return respond({ status: 'success', message: 'Modul berhasil ditambahkan!' });
  } 
  else if (type === 'lesson') {
    const sheet = ss.getSheetByName(SHEET_NAMES.LESSONS);
    const id = "les-" + Date.now();
    const module_id = payload.module_id || "";
    const judul_sub = payload.judul_sub || "";
    const konten_html = payload.konten_html || "";
    const video_url = payload.video_url || "";
    const urutan = payload.urutan || "1";
    const created_at = new Date().toISOString();

    sheet.appendRow([id, module_id, judul_sub, konten_html, video_url, urutan, created_at]);
    return respond({ status: 'success', message: 'Sub-materi berhasil ditambahkan!' });
  }

  return respond({ status: 'error', message: 'Tipe entitas create tidak dikenali' });
}

function handleUpdate(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const type = payload.type;
  const id = payload.id;
  
  const sheetName = type === 'module' ? SHEET_NAMES.MODULES : SHEET_NAMES.LESSONS;
  const sheet = ss.getSheetByName(sheetName);
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idxId = headers.indexOf('id');

  for (let i = 1; i < rows.length; i++) {
    if (String(rows[i][idxId]) === String(id)) {
      headers.forEach((h, colIdx) => {
        if (payload[h] !== undefined) {
          sheet.getRange(i + 1, colIdx + 1).setValue(payload[h]);
        }
      });
      return respond({ status: 'success', message: 'Data berhasil diperbarui!' });
    }
  }
  return respond({ status: 'error', message: 'ID tidak ditemukan untuk pembaruan' });
}

function handleDelete(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const type = payload.type;
  const id = payload.id;

  if (type === 'module') {
    // Hapus modul sekaligus lesson di dalamnya
    const modSheet = ss.getSheetByName(SHEET_NAMES.MODULES);
    deleteRowById(modSheet, id);

    const lesSheet = ss.getSheetByName(SHEET_NAMES.LESSONS);
    deleteRowsByForeignKey(lesSheet, 'module_id', id);

    return respond({ status: 'success', message: 'Modul dan sub-materi terkait berhasil dihapus' });
  } else if (type === 'lesson') {
    const lesSheet = ss.getSheetByName(SHEET_NAMES.LESSONS);
    deleteRowById(lesSheet, id);
    return respond({ status: 'success', message: 'Sub-materi berhasil dihapus' });
  }

  return respond({ status: 'error', message: 'Tipe delete tidak valid' });
}

function deleteRowById(sheet, id) {
  const rows = sheet.getDataRange().getValues();
  const headers = rows[0];
  const idxId = headers.indexOf('id');
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][idxId]) === String(id)) {
      sheet.deleteRow(i + 1);
      break;
    }
  }
}

function deleteRowsByForeignKey(sheet, foreignKeyColName, foreignKeyVal) {
  const rows = sheet.getDataRange().getValues();
  if (rows.length <= 1) return;
  const headers = rows[0];
  const idxCol = headers.indexOf(foreignKeyColName);
  for (let i = rows.length - 1; i >= 1; i--) {
    if (String(rows[i][idxCol]) === String(foreignKeyVal)) {
      sheet.deleteRow(i + 1);
    }
  }
}

function handleUpload(payload) {
  try {
    const base64Data = payload.fileData.split(',')[1];
    const contentType = payload.fileData.split(';')[0].split(':')[1];
    const decoded = Utilities.base64Decode(base64Data);
    const blob = Utilities.newBlob(decoded, contentType, payload.fileName);
    
    // Simpan ke Google Drive
    const file = DriveApp.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    
    const fileUrl = "https://lh3.googleusercontent.com/d/" + file.getId();
    return respond({ status: 'success', message: 'Upload berhasil', data: { url: fileUrl } });
  } catch (err) {
    return respond({ status: 'error', message: 'Gagal mengunggah file: ' + err.message });
  }
}