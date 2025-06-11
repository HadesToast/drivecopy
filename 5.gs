/**
 * @OnlyCurrentDoc
 * 
 * PLATINUM EDITION - Google Drive Folder Copier
 * Versi "Mission Control" dengan UI modern, dashboard canggih, 
 * dan fitur profesional seperti Mode Validasi (Dry Run) dan Notifikasi Telegram.
 * 
 * -- REVISI FINAL v3 (Integrasi Telegram) --
 * - FITUR BARU: Notifikasi status pekerjaan (mulai, selesai, berhenti, error) via Bot Telegram.
 * - FITUR BARU: Menu 'Konfigurasi Telegram' untuk setup API Token dan Chat ID.
 * - PENYEMPURNAAN: Kode diorganisir dengan seksi baru untuk integrasi eksternal.
 * - PERBAIKAN: Mengatasi error 'Jumlah baris tidak cocok' pada setupDashboard.
 * - PERBAIKAN: Mengatasi 'Error mengurai formula' dengan sintaks yang lebih robust.
 */

// --- KONFIGURASI GLOBAL ---
const a_props = PropertiesService.getScriptProperties();
const a_cache = CacheService.getScriptCache();
const DASHBOARD_NAME = '🚀 Mission Control';
const ss = SpreadsheetApp.getActiveSpreadsheet();
const TELEGRAM_API_URL = 'https://api.telegram.org/bot';

// --- FUNGSI UTAMA & MENU ---
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('⭐ Drive Copier Platinum')
      .addItem('🛰️ Buka Remote Control', 'showDialog')
      .addSeparator()
      .addItem('⚙️ Konfigurasi Telegram', 'configureTelegram')
      .addSeparator()
      .addItem('🧹 Reset Mission Control', 'resetDashboard')
      .addToUi();
}

function showDialog() {
  const html = HtmlService.createHtmlOutputFromFile('DialogUI').setWidth(500).setHeight(550);
  SpreadsheetApp.getUi().showModalDialog(html, 'Drive Copier - Remote Control');
}

// =================================================================
// SECTION: INTEGRASI TELEGRAM
// =================================================================

/**
 * Memunculkan dialog untuk mengkonfigurasi kredensial Bot Telegram.
 */
function configureTelegram() {
  const ui = SpreadsheetApp.getUi();
  const currentToken = a_props.getProperty('telegramBotToken') || '';
  const currentChatId = a_props.getProperty('telegramChatId') || '';

  const tokenResponse = ui.prompt('Konfigurasi Telegram - Langkah 1/2', `Masukkan API Token Bot Telegram Anda:\n(Biarkan kosong untuk tidak mengubah)`, ui.ButtonSet.OK_CANCEL);
  if (tokenResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Konfigurasi dibatalkan.');
    return;
  }
  const botToken = tokenResponse.getResponseText().trim();
  
  const chatIdResponse = ui.prompt('Konfigurasi Telegram - Langkah 2/2', `Masukkan Chat ID Anda (user, grup, atau channel):\n(Biarkan kosong untuk tidak mengubah)`, ui.ButtonSet.OK_CANCEL);
  if (chatIdResponse.getSelectedButton() !== ui.Button.OK) {
    ui.alert('Konfigurasi dibatalkan.');
    return;
  }
  const chatId = chatIdResponse.getResponseText().trim();

  if (botToken) {
    a_props.setProperty('telegramBotToken', botToken);
  }
  if (chatId) {
    a_props.setProperty('telegramChatId', chatId);
  }
  
  if (botToken || chatId) {
    ui.alert('Sukses!', 'Konfigurasi Telegram telah disimpan.', ui.ButtonSet.OK);
  } else {
    ui.alert('Tidak ada perubahan disimpan.');
  }
}

/**
 * Mengirim pesan notifikasi ke chat Telegram yang dikonfigurasi.
 * @param {string} message - Teks pesan yang akan dikirim. Mendukung format HTML.
 */
function sendTelegramNotification(message) {
  const token = a_props.getProperty('telegramBotToken');
  const chatId = a_props.getProperty('telegramChatId');

  if (!token || !chatId) {
    Logger.log('Kredensial Telegram tidak diatur, notifikasi dilewati.');
    return;
  }

  try {
    const url = `${TELEGRAM_API_URL}${token}/sendMessage`;
    const payload = {
      'chat_id': chatId,
      'text': message,
      'parse_mode': 'HTML'
    };
    
    const options = {
      'method': 'post',
      'contentType': 'application/json',
      'payload': JSON.stringify(payload),
      'muteHttpExceptions': true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    Logger.log(`Respons Notifikasi Telegram: ${response.getContentText()}`);
  } catch (e) {
    Logger.log(`Gagal mengirim notifikasi Telegram: ${e.toString()}`);
  }
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

function getJobProgress() {
  const state = a_props.getProperties();
  return {
    processed: parseInt(state.processedItems || '0'),
    total: parseInt(state.totalItems || '1'),
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
    clearPreviousJob();
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
  
  try {
    const { sourceFolderUrl, newFolderName, destinationFolderUrl, skipExisting, isDryRun } = jobConfig;
    
    const sourceFolder = DriveApp.getFolderById(getFolderIdFromUrl(sourceFolderUrl));
    let destParentFolder = destinationFolderUrl ? DriveApp.getFolderById(getFolderIdFromUrl(destinationFolderUrl)) : DriveApp.getRootFolder();
    
    const mode = isDryRun ? '🔬 Validasi (Dry Run)' : '🛰️ Penyalinan Live';
    const startMessage = `
<b>🚀 Misi Dimulai: ${newFolderName}</b>
--------------------------------------
<b>Mode:</b> ${mode}
<b>Sumber:</b> ${sourceFolder.getName()}
<i>Proses sedang berjalan di server...</i>
    `;
    sendTelegramNotification(startMessage);

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
  } catch (e) {
    Logger.log(`Error on startCopyJob: ${e.toString()}\nStack: ${e.stack}`);
    const errorMessage = `<b>⛔️ Gagal Memulai Misi</b>\n\nPesan Error: <code>${e.message}</code>\n\nPastikan URL folder sumber valid dan Anda memiliki akses.`;
    sendTelegramNotification(errorMessage);
    return { status: 'error', message: `Gagal memulai: ${e.message}` };
  }
}

function resumeCopyJob() {
  a_props.setProperty('stopRequested', 'false');
  const jobName = a_props.getProperty('newFolderName') || 'Pekerjaan Sebelumnya';
  sendTelegramNotification(`<b>▶️ Melanjutkan Misi: ${jobName}</b>`);
  updateDashboardStatus('🛰️ MELANJUTKAN...', '#4285f4');
  return executeCopyJob();
}

function executeCopyJob() {
  try {
    const state = a_props.getProperties();
    const sourceFolder = DriveApp.getFolderById(state.sourceFolderId);
    const isDryRun = state.isDryRun === 'true';

    let targetFolder;
    if (!isDryRun) {
        targetFolder = DriveApp.getFolderById(state.targetFolderId);
    }
  
    copyFolderRecursive(sourceFolder, targetFolder, '/', isDryRun, state.skipExisting === 'true');

    if (a_props.getProperty('stopRequested') !== 'true') {
        const finalStatus = isDryRun ? '✅ VALIDASI SELESAI' : '✅ PROSES SELESAI';
        updateDashboardStatus(finalStatus, '#0f9d58');
        const url = isDryRun ? '' : targetFolder.getUrl();
        
        const stats = a_props.getProperties();
        const successCount = stats.stats_success || 0;
        const skippedCount = stats.stats_skipped || 0;
        const failedCount = stats.stats_failed || 0;
        let completionMessage;
        if(isDryRun) {
          completionMessage = `
<b>✅ Validasi Selesai: ${state.newFolderName}</b>
--------------------------------------
<b>Hasil Simulasi:</b>
✔️ Sukses: ${successCount}
⏩ Dilewati: ${skippedCount}
❌ Gagal: ${failedCount}
<i>Tidak ada file/folder yang disalin. Periksa sheet untuk detail.</i>`;
        } else {
          completionMessage = `
<b>✅ Misi Selesai: ${state.newFolderName}</b>
--------------------------------------
<b>Statistik:</b>
✔️ Sukses: ${successCount}
⏩ Dilewati: ${skippedCount}
❌ Gagal: ${failedCount}

<a href="${url}"><b>Buka Folder Hasil Salinan</b></a>`;
        }
        sendTelegramNotification(completionMessage);

        clearPreviousJob();
        return { status: 'completed', newFolderUrl: url, isDryRun: isDryRun };
    } else {
        updateDashboardStatus('🟡 DIHENTIKAN', '#ff9800');
        sendTelegramNotification(`<b>🟡 Misi Dihentikan: ${state.newFolderName}</b>\n\nProses dihentikan oleh pengguna. Progres telah disimpan dan dapat dilanjutkan nanti.`);
        return { status: 'stopped' };
    }
  } catch (e) {
    Logger.log(`Error during executeCopyJob: ${e.toString()}\nStack: ${e.stack}`);
    updateDashboardStatus('⛔ ERROR KRITIS', '#db4437');
    logToDashboard([['❌ GAGAL', 'PROSES UTAMA', 'Sistem', '/', '', '', `Error: ${e.message}`]]);
    
    const jobName = a_props.getProperty('newFolderName') || 'Pekerjaan Saat Ini';
    const errorMessage = `<b>⛔️ ERROR KRITIS pada Misi: ${jobName}</b>\n\nProses penyalinan gagal total karena error tak terduga.\n\nPesan: <code>${e.message}</code>\n\nSilakan periksa log di sheet Mission Control.`;
    sendTelegramNotification(errorMessage);

    return { status: 'error', message: `Kesalahan pada proses utama: ${e.message}` };
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
          logToDashboard([['⏩ Dilewati', folderName, 'Folder', currentPath, subFolder.getUrl(), 'Sudah ada (tidak dibuat ulang)']]);
          incrementStats('skipped');
          incrementProcessedCount(); 
          if (!isDryRun) {
            nextTargetFolder = target.getFoldersByName(folderName).next();
            copyFolderRecursive(subFolder, nextTargetFolder, newPath, isDryRun, skipExisting);
          } else {
            copyFolderRecursive(subFolder, null, newPath, isDryRun, skipExisting);
          }
      } else {
          incrementProcessedCount();
          if (isDryRun) {
              logToDashboard([['✔️ Akan Dibuat', folderName, 'Folder', currentPath, subFolder.getUrl()]]);
              incrementStats('success');
              copyFolderRecursive(subFolder, null, newPath, isDryRun, skipExisting);
          } else {
              updateStatus(`Membuat folder: ${newPath}`);
              try {
                  nextTargetFolder = target.createFolder(folderName);
                  logToDashboard([['✅ Sukses', folderName, 'Folder', currentPath, subFolder.getUrl(), nextTargetFolder.getUrl()]]);
                  incrementStats('success');
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
  const propsToKeep = ['telegramBotToken', 'telegramChatId'];
  const allProps = a_props.getProperties();
  for (const key in allProps) {
    if (!propsToKeep.includes(key)) {
      a_props.deleteProperty(key);
    }
  }
  a_cache.removeAll(['existingItemsCache']);
}

function updateStatus(message) { a_props.setProperty('statusMessage', message); }

function countItemsRecursive(folder) {
  let counts = { files: 0, folders: 0, total: 0 };
  try {
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
  } catch (e) {
    Logger.log(`Tidak dapat mengakses item di dalam folder ${folder.getName()}. Error: ${e.message}`);
  }
  return counts;
}

function getExistingItems(targetFolder, isDryRun) {
  if (isDryRun || !targetFolder) return new Map();
  const cacheKey = 'existingItemsCache_' + targetFolder.getId();
  const cached = a_cache.get(cacheKey);
  if (cached) return new Map(JSON.parse(cached));

  const items = new Map();
  try {
    const files = targetFolder.getFiles(); while (files.hasNext()) { items.set(files.next().getName(), 'file'); }
    const folders = targetFolder.getFolders(); while (folders.hasNext()) { items.set(folders.next().getName(), 'folder'); }
    a_cache.put(cacheKey, JSON.stringify(Array.from(items.entries())), 300);
  } catch(e) {
    Logger.log(`Gagal mendapatkan item dari folder tujuan: ${e.message}`);
  }
  return items;
}

function logToDashboard(dataRows) {
  try {
    const sheet = ss.getSheetByName(DASHBOARD_NAME);
    if (!sheet || dataRows.length === 0) return;
    
    const data = Array.isArray(dataRows[0]) ? dataRows : [dataRows];
    
    const lastRow = sheet.getLastRow() + 1;
    const numRows = data.length;
    
    const now = new Date();
    const rowsToInsert = data.map(row => {
      const newRow = [...row];
      while (newRow.length < 7) newRow.push('');
      newRow.push(now);
      return newRow;
    });
    
    sheet.getRange(lastRow, 1, numRows, rowsToInsert[0].length).setValues(rowsToInsert);
  } catch (e) {
    Logger.log(`Gagal menulis log ke dashboard: ${e.message}`);
  }
}

function incrementProcessedCount(count = 1) {
  try {
    const current = parseInt(a_props.getProperty('processedItems') || '0') + count;
    a_props.setProperty('processedItems', current.toString());
    const sheet = ss.getSheetByName(DASHBOARD_NAME);
    if (sheet) sheet.getRange('F7').setValue(current);
  } catch (e) { /* Abaikan jika sheet sibuk */ }
}

function incrementStats(type, count = 1) {
  const key = 'stats_' + type;
  try {
    const current = parseInt(a_props.getProperty(key) || '0') + count;
    a_props.setProperty(key, current.toString());
    const sheet = ss.getSheetByName(DASHBOARD_NAME);
    if (sheet) {
        const cellMap = { success: 'J3', skipped: 'J4', failed: 'J5' };
        if (cellMap[type]) {
            sheet.getRange(cellMap[type]).setValue(current);
        }
    }
  } catch(e) { /* Abaikan jika sheet sibuk */ }
}

function setupDashboard(sourceFolder, targetFolder, totalItems, isDryRun) {
    let sheet = ss.getSheetByName(DASHBOARD_NAME);
    if (sheet) { sheet.clear(); } else { sheet = ss.insertSheet(DASHBOARD_NAME, 0); }
    ss.setActiveSheet(sheet);

    sheet.setFrozenRows(9); // Disesuaikan menjadi 9 untuk header log
    sheet.getRange('A1:J9').setFontFamily('Google Sans').setVerticalAlignment('middle');
    
    // Header
    sheet.getRange('A1:F1').merge().setValue('🚀 MISSION CONTROL').setFontSize(18).setFontWeight('bold');
    const mode = isDryRun ? 'MODE VALIDASI (DRY RUN)' : 'MODE PENYALINAN LIVE';
    sheet.getRange('A2:F2').merge().setValue(mode).setFontColor('#db4437').setFontWeight('bold');
    
    // Info Panel
    sheet.getRange('A4').setValue('Sumber:').setFontWeight('bold');
    const sourceLink = SpreadsheetApp.newRichTextValue().setText(sourceFolder.getName()).setLinkUrl(sourceFolder.getUrl()).build();
    sheet.getRange('B4').setRichTextValue(sourceLink);

    sheet.getRange('A5').setValue('Tujuan:').setFontWeight('bold');
    if (isDryRun) {
      sheet.getRange('B5').setValue('(Simulasi)');
    } else {
      const targetLink = SpreadsheetApp.newRichTextValue().setText(targetFolder.getName()).setLinkUrl(targetFolder.getUrl()).build();
      sheet.getRange('B5').setRichTextValue(targetLink);
    }
    
    // Progress Panel
    sheet.getRange('A7').setValue('Status:').setFontWeight('bold');
    sheet.getRange('B7').setFontSize(14).setFontWeight('bold');
    sheet.getRange('D7').setValue('Progres:').setHorizontalAlignment('right').setFontWeight('bold');
    sheet.getRange('E7').setFormula('=SPARKLINE(F7, {"charttype", "bar"; "max", G7; "color1", "#4285f4"})');
    sheet.getRange('F7').setValue(0); // Processed
    sheet.getRange('G7').setValue(totalItems); // Total
    sheet.getRange('H7').setFormula('=IFERROR(F7/G7, 0)');
    sheet.getRange('H7').setNumberFormat('0.00%');
    
    // Stats Panel
    sheet.getRange('I2').setValue('📊 STATISTIK').setFontWeight('bold');
    sheet.getRange('I3:J5').setValues([
      ['✅ Sukses:', 0],
      ['⏩ Dilewati:', 0],
      ['❌ Gagal:', 0]
    ]).setHorizontalAlignments([['right', 'left']]);

    // Pie Chart
    const chartRange = sheet.getRange('I3:J5');
    const chart = sheet.newChart().setChartType(Charts.ChartType.PIE)
        .addRange(chartRange).setOption('title', 'Komposisi Hasil')
        .setOption('pieHole', 0.4).setOption('legend', { position: 'right' })
        .setOption('colors', ['#0f9d58', '#9e9e9e', '#db4437'])
        .setPosition(2, 4, 15, 15).build();
    sheet.insertChart(chart);

    // Log Headers
    const headers = ['Status', 'Nama Item', 'Tipe', 'Lokasi Relatif', 'Link Asli', 'Link Hasil Copy', 'Pesan', 'Timestamp'];
    sheet.getRange(9, 1, 1, headers.length).setValues([headers]).setBackground('#f3f3f3').setFontWeight('bold');
    
    // Formatting
    const dataRange = sheet.getRange('A10:H' + sheet.getMaxRows());
    sheet.clearConditionalFormatRules();
    const rules = [
      SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('❌').setBackground('#fce8e6').setRanges([dataRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('⏩').setFontColor('#757575').setRanges([dataRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('✔️').setFontColor('#188038').setRanges([dataRange]).build(),
      SpreadsheetApp.newConditionalFormatRule().whenTextStartsWith('✅').setFontColor('#188038').setRanges([dataRange]).build()
    ];
    sheet.setConditionalFormatRules(rules);

    // Column Widths
    sheet.setColumnWidths(1, 1, 130); sheet.setColumnWidth(2, 250);
    sheet.setColumnWidth(3, 80); sheet.setColumnWidth(4, 250);
    sheet.setColumnWidths(5, 2, 80); sheet.setColumnWidth(7, 200);
    sheet.setColumnWidth(8, 150);
    
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
    let match = urlOrId.match(/folders\/([a-zA-Z0-9_-]{28,})/);
    if (match) return match[1];
    match = urlOrId.match(/id=([a-zA-Z0-9_-]{28,})/);
    if (match) return match[1];
    if (urlOrId.length > 25 && !urlOrId.includes('/')) return urlOrId;
    throw new Error("Format URL atau ID Folder tidak valid. Salin URL lengkap dari address bar atau hanya ID foldernya.");
}