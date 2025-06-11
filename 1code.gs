/**
 * @OnlyCurrentDoc
 *
 * ULTIMATE Google Drive Folder Copier
 * Versi dengan Logging yang Disempurnakan
 * Fitur:
 * - Menyalin folder (termasuk folder bersama)
 * - Progress bar dan status real-time
 * - Logging detail ke Spreadsheet (dengan path/lokasi folder)
 * - Tombol STOP untuk menghentikan proses
 * - Fitur RESUME untuk melanjutkan proses yang terhenti/timeout
 */

const a_props = PropertiesService.getScriptProperties();

// Fungsi ini berjalan otomatis saat spreadsheet dibuka
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Drive Copier Ultimate')
      .addItem('▶️ Buka Alat Copy', 'showCopyDialog')
      .addToUi();
}

/**
 * Menampilkan dialog HTML.
 */
function showCopyDialog() {
  const html = HtmlService.createHtmlOutputFromFile('DialogUI')
      .setWidth(480)
      .setHeight(500);
  SpreadsheetApp.getUi().showModalDialog(html, 'Drive Copier Ultimate');
}

// =================================================================
// SECTION: FUNGSI MANAJEMEN STATUS & STATE
// =================================================================

/**
 * Dipanggil oleh UI saat pertama kali dimuat untuk memeriksa apakah ada pekerjaan yang bisa dilanjutkan.
 * @returns {object} Objek status pekerjaan saat ini.
 */
function getInitialState() {
  const state = a_props.getProperties();
  if (state.resume_possible === 'true') {
    return {
      resumable: true,
      newFolderName: state.newFolderName,
      processed: parseInt(state.processedItems || '0'),
      total: parseInt(state.totalItems || '0')
    };
  }
  return { resumable: false };
}

/**
 * Membersihkan semua state pekerjaan sebelumnya untuk memulai dari awal.
 */
function clearPreviousJob() {
  a_props.deleteAllProperties();
}

/**
 * Dipanggil oleh tombol STOP di UI.
 */
function requestStopProcess() {
  a_props.setProperty('stopRequested', 'true');
}

/**
 * Fungsi yang dipanggil oleh klien untuk mendapatkan status progres saat ini.
 * @returns {object} Objek yang berisi status progres.
 */
function getCopyProgress() {
  const total = parseInt(a_props.getProperty('totalItems') || '0');
  const processed = parseInt(a_props.getProperty('processedItems') || '0');
  const statusMessage = a_props.getProperty('statusMessage') || 'Menunggu...';
  
  return {
    percentage: total > 0 ? (processed / total) * 100 : 0,
    statusMessage: statusMessage,
    processed: processed,
    total: total,
    stopRequested: a_props.getProperty('stopRequested') === 'true'
  };
}

// =================================================================
// SECTION: FUNGSI INTI PROSES COPY
// =================================================================

/**
 * Mempersiapkan dan memulai proses penyalinan BARU.
 * @returns {string} ID target folder yang baru dibuat.
 */
function prepareAndStartCopyProcess(sourceFolderUrl, newFolderName, destinationFolderUrl) {
  clearPreviousJob();
  
  const sourceFolderId = getFolderIdFromUrl(sourceFolderUrl);
  const sourceFolder = DriveApp.getFolderById(sourceFolderId);
  
  let destinationParentFolder;
  if (destinationFolderUrl) {
    destinationParentFolder = DriveApp.getFolderById(getFolderIdFromUrl(destinationFolderUrl));
  } else {
    destinationParentFolder = DriveApp.getRootFolder();
  }
  
  updateStatus('Menghitung total item...');
  const counts = countItemsRecursive(sourceFolder);
  const totalItems = counts.files + counts.folders; 

  updateStatus(`Membuat folder utama: ${newFolderName}`);
  const newMainFolder = destinationParentFolder.createFolder(newFolderName);
  
  // Simpan state awal
  a_props.setProperties({
    'sourceFolderId': sourceFolderId,
    'targetFolderId': newMainFolder.getId(),
    'newFolderName': newFolderName,
    'totalItems': totalItems.toString(),
    'processedItems': '0',
    'logData': JSON.stringify([]),
    'resume_possible': 'true'
  });
  
  // Log pembuatan folder root
  addToLog(newMainFolder.getName(), 'Folder', '/', newMainFolder.getUrl());
  
  return startCopyExecution();
}

/**
 * Melanjutkan proses copy yang sudah ada.
 */
function resumeCopyProcess() {
  a_props.setProperty('stopRequested', 'false');
  return startCopyExecution();
}

/**
 * Mesin eksekusi utama.
 */
function startCopyExecution() {
  const state = a_props.getProperties();
  const sourceFolder = DriveApp.getFolderById(state.sourceFolderId);
  const targetFolder = DriveApp.getFolderById(state.targetFolderId);

  try {
    // Mulai/Lanjutkan proses rekursif dari path root ('/')
    copyFolderRecursive(sourceFolder, targetFolder, '/');

    if (a_props.getProperty('stopRequested') !== 'true') {
        updateStatus('Menulis log ke spreadsheet...');
        const logSheetUrl = writeLogToSheet();
        a_props.deleteAllProperties();
        return {
            status: 'completed',
            newFolderUrl: targetFolder.getUrl(),
            logSheetUrl: logSheetUrl
        };
    } else {
        return { status: 'stopped' };
    }
  } catch (e) {
    Logger.log(`Error: ${e.toString()} \nStack: ${e.stack}`);
    throw new Error(`Proses berhenti karena error: ${e.message}. Anda bisa mencoba melanjutkannya.`);
  }
}

/**
 * Fungsi rekursif yang disempurnakan dengan path tracking.
 * @param {Folder} source - Folder sumber.
 * @param {Folder} target - Folder tujuan.
 * @param {string} currentPath - Path relatif saat ini.
 */
function copyFolderRecursive(source, target, currentPath) {
  if (a_props.getProperty('stopRequested') === 'true') return;
  
  const existingItems = new Map();
  const targetFiles = target.getFiles(); while(targetFiles.hasNext()){ existingItems.set(targetFiles.next().getName(), 'file'); }
  const targetFolders = target.getFolders(); while(targetFolders.hasNext()){ const f = targetFolders.next(); existingItems.set(f.getName(), f); }

  const files = source.getFiles();
  while (files.hasNext()) {
    if (a_props.getProperty('stopRequested') === 'true') return;
    const file = files.next();
    const fileName = file.getName();
    
    if (!existingItems.has(fileName)) {
      updateStatus(`Menyalin file: ${currentPath}${fileName}`);
      try {
        const copiedFile = file.makeCopy(fileName, target);
        addToLog(fileName, 'File', currentPath, copiedFile.getUrl());
        incrementProcessedCount();
      } catch(e) {
        Logger.log(`Gagal menyalin file ${fileName}: ${e.message}`);
        addToLog(`${fileName} (GAGAL)`, 'File', currentPath, '#');
        incrementProcessedCount();
      }
    }
  }

  const subFolders = source.getFolders();
  while (subFolders.hasNext()) {
    if (a_props.getProperty('stopRequested') === 'true') return;
    const subFolder = subFolders.next();
    const folderName = subFolder.getName();
    const newPath = `${currentPath}${folderName}/`;
    
    let nextTargetFolder;
    if (existingItems.has(folderName) && existingItems.get(folderName) !== 'file') {
      nextTargetFolder = existingItems.get(folderName);
    } else {
      updateStatus(`Membuat folder: ${newPath}`);
      nextTargetFolder = target.createFolder(folderName);
      addToLog(folderName, 'Folder', currentPath, nextTargetFolder.getUrl());
      incrementProcessedCount();
    }
    copyFolderRecursive(subFolder, nextTargetFolder, newPath);
  }
}

// =================================================================
// SECTION: FUNGSI HELPER & LOGGING
// =================================================================

function countItemsRecursive(folder) {
  let counts = { files: 0, folders: 0 };
  const files = folder.getFiles(); while(files.hasNext()){ files.next(); counts.files++; }
  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    const sub = subFolders.next();
    counts.folders++;
    const subCounts = countItemsRecursive(sub);
    counts.files += subCounts.files;
    counts.folders += subCounts.folders;
  }
  return counts;
}

function updateStatus(message) {
  a_props.setProperty('statusMessage', message);
}

function incrementProcessedCount() {
  const current = parseInt(a_props.getProperty('processedItems') || '0');
  a_props.setProperty('processedItems', (current + 1).toString());
}

/**
 * Menambahkan entri ke log.
 * @param {string} name Nama item.
 * @param {string} type Tipe item ('File' atau 'Folder').
 * @param {string} path Lokasi/path di folder tujuan.
 * @param {string} url URL langsung ke item.
 */
function addToLog(name, type, path, url) {
  const logData = JSON.parse(a_props.getProperty('logData'));
  logData.push([
    new Date().toLocaleString('id-ID', { timeZone: "Asia/Jakarta" }), 
    name, 
    type, 
    path, 
    url
  ]);
  a_props.setProperty('logData', JSON.stringify(logData));
}

/**
 * Menulis data log ke spreadsheet dengan format yang disempurnakan.
 * @returns {string} URL dari sheet log.
 */
function writeLogToSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const folderName = a_props.getProperty('newFolderName').replace(/[^a-zA-Z0-9]/g, '_'); // Sanitasi nama folder
  const dateStr = new Date().toISOString().slice(0, 10);
  const sheetName = `Log_Copy_${folderName}_${dateStr}`.substring(0, 100);
  
  let sheet = ss.getSheetByName(sheetName);
  if (sheet) { 
    sheet.clear(); 
  } else { 
    sheet = ss.insertSheet(sheetName, 0); // Insert di paling depan
  }
  ss.setActiveSheet(sheet);
  
  const headers = ["Timestamp", "Nama Item", "Tipe", "Lokasi di Folder Baru", "Link Langsung"];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight("bold")
       .setBackground("#e6f4ea")
       .setHorizontalAlignment("center");
  
  const logData = JSON.parse(a_props.getProperty('logData') || '[]');
  if (logData.length > 0) {
    sheet.getRange(2, 1, logData.length, headers.length).setValues(logData);
  }

  // Atur lebar kolom
  sheet.setColumnWidth(1, 150); // Timestamp
  sheet.setColumnWidth(2, 250); // Nama Item
  sheet.setColumnWidth(3, 80);  // Tipe
  sheet.setColumnWidth(4, 300); // Lokasi
  sheet.setColumnWidth(5, 300); // Link
  
  // Freeze baris header
  sheet.setFrozenRows(1);
  
  return ss.getUrl() + "#gid=" + sheet.getSheetId();
}

function getFolderIdFromUrl(urlOrId) {
    if (!urlOrId) throw new Error("URL/ID tidak boleh kosong.");
    let match = urlOrId.match(/folders\/([a-zA-Z0-9_-]{15,})/);
    if (match) return match[1];
    match = urlOrId.match(/id=([a-zA-Z0-9_-]{15,})/);
    if (match) return match[1];
    if (urlOrId.length > 15 && !urlOrId.includes('/')) return urlOrId;
    throw new Error("Format URL atau ID Folder tidak valid.");
}
