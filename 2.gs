/**
 * @OnlyCurrentDoc
 * 
 * PLATINUM EDITION - Google Drive Folder Copier
 * Versi "Mission Control" dengan UI modern, dashboard canggih, 
 * dan fitur profesional seperti Mode Validasi (Dry Run).
 * 
 * -- DISEMPURNAKAN --
 * - UI Dialog dengan Live Progress Bar.
 * - Penjelasan proses latar belakang (aman untuk menutup dialog).
 * - Perbaikan logika dan peningkatan stabilitas.
 */

// --- KONFIGURASI GLOBAL ---
const a_props = PropertiesService.getScriptProperties();
const a_cache = CacheService.getScriptCache();
const DASHBOARD_NAME = '🚀 Mission Control';
const ss = SpreadsheetApp.getActiveSpreadsheet();

// --- FUNGSI UTAMA & MENU ---
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('⭐ Drive Copier Platinum')
      .addItem('🛰️ Buka Remote Control', 'showDialog')
      .addSeparator()
      .addItem('🧹 Reset Mission Control', 'resetDashboard')
      .addToUi();
}

function showDialog() {
  const html = HtmlService.createHtmlOutputFromFile('DialogUI').setWidth(500).setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Drive Copier - Remote Control');
}

// =================================================================
// SECTION: MANAJEMEN DASHBOARD, STATE, & KONTROL
// =================================================================

function getInitialState() {
  const state = a_props.getProperties();
  if (state.resume_possible === 'true') {
    return {
      resumable: true,
      jobData: {
        newFolderName: state.newFolderName,
        sourceFolderName: state.sourceFolderName,
        isDryRun: state.isDryRun === 'true'
      }
    };
  }
  return { resumable: false };
}

// FUNGSI BARU: Untuk di-polling oleh UI Dialog
function getJobProgress() {
  const state = a_props.getProperties();
  return {
    processed: parseInt(state.processedItems || '0'),
    total: parseInt(state.totalItems || '1'), // Hindari pembagian dengan nol
    statusMessage: state.statusMessage || 'Menunggu...',
    stopRequested: state.stopRequested === 'true'
  };
}

function requestStopProcess() {
  a_props.setProperty('stopRequested', 'true');
  updateDashboardStatus('🚦 MENGHENTIKAN...', '#ff9800');
}

function resetDashboard() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Konfirmasi Reset', 'Yakin ingin menghapus Mission Control dan semua progres tersimpan?', ui.ButtonSet.YES_NO);
  if (response == ui.Button.YES) {
    clearPreviousJob(); // Hapus properties dan cache
    const sheet = ss.getSheetByName(DASHBOARD_NAME);
    if (sheet) {
      ss.deleteSheet(sheet);
    }
    ui.alert('Mission Control telah direset.');
  }
}

function resetFromDialog() {
  clearPreviousJob();
  return { success: true };
}


// =================================================================
// SECTION: INTI LOGIKA PENYALINAN
// =================================================================

function startCopyJob(jobConfig) {
  clearPreviousJob();
  
  const { sourceFolderUrl, newFolderName, destinationFolderUrl, skipExisting, isDryRun } = jobConfig;
  
  const sourceFolder = DriveApp.getFolderById(getFolderIdFromUrl(sourceFolderUrl));
  let destParentFolder = destinationFolderUrl ? DriveApp.getFolderById(getFolderIdFromUrl(destinationFolderUrl)) : DriveApp.getRootFolder();
  
  updateStatus('Menghitung total item...');
  const counts = countItemsRecursive(sourceFolder);
  
  let newMainFolder;
  if (isDryRun) {
    updateStatus('Mode Validasi: Tidak membuat folder.');
  } else {
    updateStatus(`Membuat folder utama: ${newFolderName}`);
    newMainFolder = destParentFolder.createFolder(newFolderName);
  }
  
  const targetFolderId = isDryRun ? 'DRY_RUN_MODE' : newMainFolder.getId();

  a_props.setProperties({
    'sourceFolderId': sourceFolder.getId(),
    'sourceFolderName': sourceFolder.getName(),
    'targetFolderId': targetFolderId,
    'newFolderName': newFolderName,
    'totalItems': counts.total.toString(),
    'processedItems': '0',
    'stats_success': '0', 'stats_skipped': '0', 'stats_failed': '0',
    'skipExisting': skipExisting.toString(),
    'isDryRun': isDryRun.toString(),
    'resume_possible': 'true',
    'stopRequested': 'false'
  });
  
  setupDashboard(sourceFolder, newMainFolder, counts.total, isDryRun);
  
  return executeCopyJob();
}

function resumeCopyJob() {
  a_props.setProperty('stopRequested', 'false');
  updateDashboardStatus('🛰️ MELANJUTKAN...', '#4285f4');
  return executeCopyJob();
}

function executeCopyJob() {
  const state = a_props.getProperties();
  const sourceFolder = DriveApp.getFolderById(state.sourceFolderId);
  const isDryRun = state.isDryRun === 'true';

  let targetFolder;
  if (!isDryRun) {
      targetFolder = DriveApp.getFolderById(state.targetFolderId);
  }

  try {
    copyFolderRecursive(sourceFolder, targetFolder, '/', isDryRun, state.skipExisting === 'true');

    if (a_props.getProperty('stopRequested') !== 'true') {
        const finalStatus = isDryRun ? '✅ VALIDASI SELESAI' : '✅ PROSES SELESAI';
        updateDashboardStatus(finalStatus, '#0f9d58');
        const url = isDryRun ? '' : targetFolder.getUrl();
        clearPreviousJob(); // Hapus properti setelah selesai
        return { status: 'completed', newFolderUrl: url, isDryRun: isDryRun };
    } else {
        updateDashboardStatus('🟡 DIHENTIKAN', '#ff9800');
        // Jangan hapus properti agar bisa dilanjutkan
        return { status: 'stopped' };
    }
  } catch (e) {
    Logger.log(`Error: ${e.toString()}\nStack: ${e.stack}`);
    updateDashboardStatus('⛔ ERROR KRITIS', '#db4437');
    logToDashboard([['❌ GAGAL', 'PROSES UTAMA', 'Sistem', '/', '', '', `Error: ${e.message}`]]);
    return { status: 'error', message: e.message }; // Kembalikan status error ke client
  }
}

function copyFolderRecursive(source, target, currentPath, isDryRun, skipExisting) {
  if (a_props.getProperty('stopRequested') === 'true') return;

  const existingItems = getExistingItems(target, isDryRun);
  
  // Proses Files
  const files = source.getFiles();
  const fileLogs = [];
  while(files.hasNext()){
    if (a_props.getProperty('stopRequested') === 'true') break;
    const file = files.next();
    const fileName = file.getName();

    if (skipExisting && existingItems.has(fileName)) {
        fileLogs.push(['⏩ Dilewati', fileName, 'File', currentPath, file.getUrl(), 'Sudah ada']);
        incrementStats('skipped');
    } else {
        if (isDryRun) {
            fileLogs.push(['✔️ Akan Disalin', fileName, 'File', currentPath, file.getUrl()]);
            incrementStats('success');
        } else {
            updateStatus(`Menyalin file: ${currentPath}${fileName}`);
            try {
                const copiedFile = file.makeCopy(fileName, target);
                fileLogs.push(['✅ Sukses', fileName, 'File', currentPath, file.getUrl(), copiedFile.getUrl()]);
                incrementStats('success');
            } catch(e) {
                fileLogs.push(['❌ Gagal', fileName, 'File', currentPath, file.getUrl(), '', e.message]);
                incrementStats('failed');
            }
        }
    }
  }
  
  if (fileLogs.length > 0) {
      logToDashboard(fileLogs);
      incrementProcessedCount(fileLogs.length);
  }

  // Proses Folders
  const subFolders = source.getFolders();
  while(subFolders.hasNext()){
      if (a_props.getProperty('stopRequested') === 'true') break;
      const subFolder = subFolders.next();
      const folderName = subFolder.getName();
      const newPath = `${currentPath}${folderName}/`;

      let nextTargetFolder;
      if (skipExisting && existingItems.has(folderName)) {
          // **LOGIKA DIPERBAIKI**: Hanya skip pembuatan folder, tetap proses isinya.
          logToDashboard([['⏩ Dilewati', folderName, 'Folder', currentPath, subFolder.getUrl(), 'Sudah ada (tidak dibuat ulang)']]);
          incrementStats('skipped');
          incrementProcessedCount(); // Hanya hitung folder ini saja
          if (!isDryRun) {
            nextTargetFolder = target.getFoldersByName(folderName).next();
            copyFolderRecursive(subFolder, nextTargetFolder, newPath, isDryRun, skipExisting);
          } else {
            copyFolderRecursive(subFolder, null, newPath, isDryRun, skipExisting);
          }
      } else {
          if (isDryRun) {
              logToDashboard([['✔️ Akan Dibuat', folderName, 'Folder', currentPath, subFolder.getUrl()]]);
              incrementStats('success');
              incrementProcessedCount();
              copyFolderRecursive(subFolder, null, newPath, isDryRun, skipExisting);
          } else {
              updateStatus(`Membuat folder: ${newPath}`);
              try {
                  nextTargetFolder = target.createFolder(folderName);
                  logToDashboard([['✅ Sukses', folderName, 'Folder', currentPath, subFolder.getUrl(), nextTargetFolder.getUrl()]]);
                  incrementStats('success');
                  incrementProcessedCount();
                  copyFolderRecursive(subFolder, nextTargetFolder, newPath, isDryRun, skipExisting);
              } catch(e) {
                  logToDashboard([['❌ Gagal', folderName, 'Folder', currentPath, subFolder.getUrl(), '', e.message]]);
                  incrementStats('failed');
                  incrementProcessedCount();
              }
          }
      }
  }
}


// =================================================================
// SECTION: FUNGSI HELPER, CACHING, & LOGGING
// =================================================================

function clearPreviousJob() {
  a_props.deleteAllProperties();
  a_cache.removeAll(['existingItemsCache']); // Hapus cache terkait juga
}

function updateStatus(message) { a_props.setProperty('statusMessage', message); }

function countItemsRecursive(folder) {
  let counts = { files: 0, folders: 0, total: 0 };
  const files = folder.getFiles(); while(files.hasNext()){ files.next(); counts.files++; }
  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    const sub = subFolders.next();
    counts.folders++;
    const subCounts = countItemsRecursive(sub);
    counts.files += subCounts.files;
    counts.folders += subCounts.folders;
  }
  counts.total = counts.files + counts.folders;
  return counts;
}

function getExistingItems(targetFolder, isDryRun) {
  if (isDryRun || !targetFolder) return new Map();
  const cacheKey = 'existingItemsCache_' + targetFolder.getId();
  const cached = a_cache.get(cacheKey);
  if (cached) return new Map(JSON.parse(cached));

  const items = new Map();
  const files = targetFolder.getFiles(); while (files.hasNext()) { items.set(files.next().getName(), 'file'); }
  const folders = targetFolder.getFolders(); while (folders.hasNext()) { const f = folders.next(); items.set(f.getName(), 'folder'); }
  
  a_cache.put(cacheKey, JSON.stringify(Array.from(items.entries())), 300); // Cache for 5 mins
  return items;
}

function logToDashboard(dataRows) {
  const sheet = ss.getSheetByName(DASHBOARD_NAME);
  if (!sheet || dataRows.length === 0) return;
  
  // Pastikan dataRows selalu 2D array
  const data = Array.isArray(dataRows[0]) ? dataRows : [dataRows];
  
  const lastRow = sheet.getLastRow() + 1;
  const numRows = data.length;
  
  const now = new Date();
  const rowsToInsert = data.map(row => {
    const newRow = [...row];
    while (newRow.length < 7) newRow.push(''); // Pad empty cells if needed
    newRow.push(now); // Add timestamp
    return newRow;
  });
  
  sheet.getRange(lastRow, 1, numRows, rowsToInsert[0].length).setValues(rowsToInsert);
}


function incrementProcessedCount(count = 1) {
  try {
    const current = parseInt(a_props.getProperty('processedItems') || '0') + count;
    a_props.setProperty('processedItems', current.toString());
    const sheet = ss.getSheetByName(DASHBOARD_NAME);
    if (sheet) sheet.getRange('E6').setValue(current);
  } catch (e) {
    // Abaikan error jika sheet tidak bisa diakses, properti tetap tersimpan
  }
}

function incrementStats(type, count = 1) {
  const key = 'stats_' + type;
  try {
    const current = parseInt(a_props.getProperty(key) || '0') + count;
    a_props.setProperty(key, current.toString());
    const sheet = ss.getSheetByName(DASHBOARD_NAME);
    if (sheet) {
        const cellMap = { success: 'F3', skipped: 'F4', failed: 'F5' };
        if (cellMap[type]) {
            sheet.getRange(cellMap[type]).setValue(current);
        }
    }
  } catch(e) {
    // Abaikan error jika sheet tidak bisa diakses, properti tetap tersimpan
  }
}

// --- Dashboard Setup ---
function setupDashboard(sourceFolder, targetFolder, totalItems, isDryRun) {
    let sheet = ss.getSheetByName(DASHBOARD_NAME);
    if (sheet) { sheet.clear(); } else { sheet = ss.insertSheet(DASHBOARD_NAME, 0); }
    ss.setActiveSheet(sheet);

    sheet.setFrozenRows(8);
    sheet.getRange('A1:I8').setFontFamily('Google Sans').setVerticalAlignment('middle');
    
    // Header
    const mode = isDryRun ? 'MODE VALIDASI (DRY RUN)' : 'MODE PENYALINAN LIVE';
    sheet.getRange('A1').setValue('🚀 MISSION CONTROL').setFontSize(18).setFontWeight('bold');
    sheet.getRange('A2').setValue(mode).setFontColor('#db4437').setFontWeight('bold');
    
    // Info
    sheet.getRange('B3').setValue('Sumber:');
    sheet.getRange('C3').setValue(sourceFolder.getName()).setHyperlink(sourceFolder.getUrl());
    sheet.getRange('B4').setValue('Tujuan:');
    if (!isDryRun) {
      sheet.getRange('C4').setValue(targetFolder.getName()).setHyperlink(targetFolder.getUrl());
    } else {
      sheet.getRange('C4').setValue('(Simulasi)');
    }


    // Progress Panel
    sheet.getRange('B6').setValue('Status:');
    sheet.getRange('D6').setValue('Progres Total:').setHorizontalAlignment('right');
    sheet.getRange('B7').setFontSize(14).setFontWeight('bold');
    sheet.getRange('E6:G6').merge();
    sheet.getRange('E6').setFormula('=SPARKLINE(E7, {"charttype","bar";"max",F7;"color1","#4285f4"})');
    sheet.getRange('E7').setValue(0); // Processed
    sheet.getRange('F7').setValue(totalItems); // Total
    sheet.getRange('G7').setFormula('=IFERROR(E7/F7, 0)').setNumberFormat('0.00%');
    
    // Stats Panel
    sheet.getRange('E2').setValue('📊 STATISTIK').setFontWeight('bold');
    sheet.getRange('E3:E5').setValues([['✅ Sukses:'], ['⏩ Dilewati:'], ['❌ Gagal:']]).setHorizontalAlignment('right');
    sheet.getRange('F3:F5').setValue(0);

    // Pie Chart
    const chartRange = sheet.getRange('E3:F5');
    const chart = sheet.newChart()
        .setChartType(Charts.ChartType.PIE)
        .addRange(chartRange)
        .setOption('title', 'Komposisi Hasil')
        .setOption('pieHole', 0.4)
        .setOption('colors', ['#0f9d58', '#999999', '#db4437'])
        .setPosition(2, 8, 0, 0)
        .build();
    sheet.insertChart(chart);

    // Log Headers
    const headers = ['Status', 'Nama Item', 'Tipe', 'Lokasi Relatif', 'Link Asli', 'Link Hasil Copy', 'Pesan', 'Timestamp'];
    sheet.getRange(9, 1, 1, headers.length).setValues([headers]).setBackground('#f3f3f3').setFontWeight('bold');
    
    // Conditional Formatting
    const dataRange = sheet.getRange('A10:H' + sheet.getMaxRows());
    sheet.clearConditionalFormatRules();
    const rules = [
      SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('❌').setBackground('#fbe5d6').setRanges([dataRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('⏩').setFontColor('#757575').setRanges([dataRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('✔️').setFontColor('#0b804b').setRanges([dataRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('✅').setFontColor('#0b804b').setRanges([dataRange]).build()
    ];
    sheet.setConditionalFormatRules(rules);

    // Column Widths
    sheet.setColumnWidths(1, 1, 130); // Status
    sheet.setColumnWidth(2, 250); // Nama
    sheet.setColumnWidth(3, 80);  // Tipe
    sheet.setColumnWidth(4, 250); // Lokasi
    sheet.setColumnWidths(5, 2, 80); // Links
    sheet.setColumnWidth(7, 200); // Pesan
    sheet.setColumnWidth(8, 150); // Timestamp
    
    updateDashboardStatus('🚦 MEMULAI...', '#fbbc04');
}

function updateDashboardStatus(text, color) {
  const sheet = ss.getSheetByName(DASHBOARD_NAME);
  if (sheet) {
    sheet.getRange('B7').setValue(text).setFontColor(color);
    SpreadsheetApp.flush();
  }
}

function getFolderIdFromUrl(urlOrId) {
    if (!urlOrId) throw new Error("URL/ID tidak boleh kosong.");
    // Mencocokkan ID dari URL format baru
    let match = urlOrId.match(/folders\/([a-zA-Z0-9_-]{28,})/);
    if (match) return match[1];
    // Mencocokkan ID dari URL format lama
    match = urlOrId.match(/id=([a-zA-Z0-9_-]{28,})/);
    if (match) return match[1];
    // Anggap sebagai ID jika tidak mengandung '/' dan panjangnya sesuai
    if (urlOrId.length > 25 && !urlOrId.includes('/')) return urlOrId;
    throw new Error("Format URL atau ID Folder tidak valid. Pastikan Anda menyalin URL lengkap dari address bar atau hanya ID foldernya.");
}
