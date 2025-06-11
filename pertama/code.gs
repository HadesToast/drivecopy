/**
 * @OnlyCurrentDoc
 *
 * Skrip ini menambahkan menu kustom ke Google Sheet untuk menyalin folder Google Drive,
 * termasuk folder yang dibagikan (shared folder).
 */

// Fungsi ini berjalan otomatis saat spreadsheet dibuka, untuk membuat menu UI.
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Drive Copier')
      .addItem('▶️ Mulai Copy Folder', 'showCopyDialog')
      .addToUi();
}

/**
 * Menampilkan dialog HTML untuk mendapatkan input dari pengguna.
 */
function showCopyDialog() {
  const html = HtmlService.createHtmlOutputFromFile('DialogUI')
      .setWidth(450)
      .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Copy Folder Google Drive');
}

/**
 * Fungsi utama yang dipanggil dari dialog HTML untuk memulai proses penyalinan.
 * @param {string} sourceFolderUrl - URL atau ID folder sumber.
 * @param {string} newFolderName - Nama untuk folder hasil salinan.
 * @param {string} destinationFolderUrl - (Opsional) URL atau ID folder tujuan.
 * @returns {string} URL dari folder baru yang berhasil dibuat.
 */
function startCopyProcess(sourceFolderUrl, newFolderName, destinationFolderUrl) {
  try {
    // Validasi input
    if (!sourceFolderUrl || !newFolderName) {
      throw new Error("Folder Sumber dan Nama Folder Baru tidak boleh kosong.");
    }

    const sourceFolderId = getFolderIdFromUrl(sourceFolderUrl);
    if (!sourceFolderId) {
      throw new Error("URL/ID Folder Sumber tidak valid.");
    }
    
    const sourceFolder = DriveApp.getFolderById(sourceFolderId);
    
    let destinationParentFolder;
    // Tentukan folder tujuan, jika tidak diisi, gunakan root 'My Drive'
    if (destinationFolderUrl) {
      const destId = getFolderIdFromUrl(destinationFolderUrl);
      if (!destId) throw new Error("URL/ID Folder Tujuan tidak valid.");
      destinationParentFolder = DriveApp.getFolderById(destId);
    } else {
      destinationParentFolder = DriveApp.getRootFolder();
    }
    
    Logger.log(`Memulai penyalinan folder '${sourceFolder.getName()}' ke folder baru bernama '${newFolderName}'`);
    
    // 1. Buat folder utama baru di tujuan
    const newMainFolder = destinationParentFolder.createFolder(newFolderName);
    
    // 2. Mulai proses penyalinan rekursif
    copyFolderRecursive(sourceFolder, newMainFolder);
    
    Logger.log(`Selesai! Folder baru tersedia di: ${newMainFolder.getUrl()}`);
    return newMainFolder.getUrl();

  } catch (e) {
    Logger.log(e.toString());
    // Melemparkan kembali error agar bisa ditangkap oleh .withFailureHandler di client-side
    throw new Error(`Terjadi kesalahan: ${e.message}`);
  }
}

/**
 * Fungsi rekursif untuk menyalin konten folder.
 * @param {Folder} source - Objek folder sumber.
 * @param {Folder} target - Objek folder tujuan tempat konten akan disalin.
 */
function copyFolderRecursive(source, target) {
  // Salin semua file di dalam folder saat ini
  const files = source.getFiles();
  while (files.hasNext()) {
    const file = files.next();
    try {
      file.makeCopy(file.getName(), target);
      Logger.log(`File disalin: ${file.getName()}`);
    } catch (e) {
      Logger.log(`Gagal menyalin file: ${file.getName()}. Error: ${e.message}`);
    }
  }
  
  // Buat ulang dan telusuri semua sub-folder
  const subFolders = source.getFolders();
  while (subFolders.hasNext()) {
    const subFolder = subFolders.next();
    // Buat sub-folder baru di dalam target
    const newSubFolder = target.createFolder(subFolder.getName());
    Logger.log(`Folder dibuat: ${subFolder.getName()}`);
    // Panggil fungsi ini lagi untuk menyalin konten sub-folder (rekursi)
    copyFolderRecursive(subFolder, newSubFolder);
  }
}

/**
 * Helper function untuk mengekstrak ID folder dari URL atau ID mentah.
 * @param {string} urlOrId - String yang bisa berupa URL lengkap atau ID folder.
 * @returns {string|null} ID folder atau null jika tidak ditemukan.
 */
function getFolderIdFromUrl(urlOrId) {
  if (urlOrId.includes("folders/")) {
    const match = urlOrId.match(/folders\/([a-zA-Z0-9_-]+)/);
    return match ? match[1] : null;
  } else if (urlOrId.includes("id=")) {
    const match = urlOrId.match(/id=([a-zA-Z0-9_-]+)/);
    return match ? match[1] : null;
  }
  // Jika input tidak mengandung format URL, anggap itu sudah ID
  else if (urlOrId.length > 15) { // ID folder biasanya panjang
    return urlOrId;
  }
  return null;
}
