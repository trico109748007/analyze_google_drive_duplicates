function scanAllFoldersForDuplicates() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('全硬碟重複檢查');
  if (sheet) { sheet.clear(); } else { sheet = ss.insertSheet('全硬碟重複檢查'); }

  // 設定標題
  sheet.appendRow(["狀態", "資料夾路徑", "檔案名稱", "檔案大小 (Bytes)", "MD5 雜湊值", "建立時間", "檔案連結"]);
  sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#f3f3f3");

  const fileMap = {}; // 用於記錄： ParentID_MD5_Size
  
  // 從根目錄開始遞迴掃描
  const root = DriveApp.getRootFolder();
  processFolder(root, "/", fileMap, sheet);

  sheet.autoResizeColumns(1, 7);
  SpreadsheetApp.getUi().alert("掃描完成！");
}

/**
 * 遞迴處理資料夾函數
 */
function processFolder(folder, path, fileMap, sheet) {
  const folderId = folder.getId();
  const files = folder.getFiles();

  while (files.hasNext()) {
    const file = files.next();
    const name = file.getName();
    const size = file.getSize();
    const date = file.getDateCreated();
    const url = file.getUrl();
    
    // 取得 MD5 雜湊值 (Google 格式檔案如 Doc/Sheet 可能無法取得，給予預設值)
    let md5 = "";
    try {
      md5 = file.getBlob().getHash();
    } catch (e) {
      md5 = "GOOGLE_FORMAT";
    }

    // 關鍵邏輯：唯一鍵值包含 FolderID，確保只有「同資料夾內」且「內容相同」才會判定為重複
    const signature = folderId + "_" + md5 + "_" + size;

    if (fileMap[signature]) {
      sheet.appendRow(["重複副本", path, name, size, md5, date, url]);
    } else {
      fileMap[signature] = true;
      // 若想連原始檔案都列出來以便比對，可取消下行註解
      // sheet.appendRow(["原始檔案", path, name, size, md5, date, url]);
    }
  }

  // 遞迴處理子資料夾
  const subFolders = folder.getFolders();
  while (subFolders.hasNext()) {
    const subFolder = subFolders.next();
    processFolder(subFolder, path + subFolder.getName() + "/", fileMap, sheet);
  }
}
