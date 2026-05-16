/**
 * SIMPEG API BACKEND
 * Endpoint untuk REST API dari Netlify
 */

const SPREADSHEET_NAME = "DB_SIMPEG_OFFICIAL";
const ROOT_FOLDER_NAME = "SIMPEG_DATA_MASTER";

// --- DATABASE HELPERS ---
function getDb() {
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    const files = DriveApp.getFilesByName(SPREADSHEET_NAME);
    if (files.hasNext()) ss = SpreadsheetApp.open(files.next());
    else ss = SpreadsheetApp.create(SPREADSHEET_NAME);
  }
  return ss;
}

function getRootFolder() {
  const folders = DriveApp.getFoldersByName(ROOT_FOLDER_NAME);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(ROOT_FOLDER_NAME);
}

// --- CORS & RESPONSE HELPER ---
function createResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// --- REST API ENDPOINTS ---
function doGet(e) {
  return createResponse({ status: "success", message: "SIMPEG API is running" });
}

function doPost(e) {
  try {
    const postData = JSON.parse(e.postData.contents);
    const action = postData.action;
    const payload = postData.payload;

    switch (action) {
      case 'getAppData':
        return createResponse(getAppData());
      case 'addPegawai':
        return createResponse(addPegawai(payload));
      case 'removePegawai':
        return createResponse(removePegawai(payload.id, payload.folderId));
      case 'uploadFile':
        return createResponse(uploadFileToDrive(payload.base64, payload.fileName, payload.pegawaiId, payload.folderId));
      case 'deleteFile':
        return createResponse(deleteSingleFile(payload.driveId));
      default:
        return createResponse({ status: "error", message: "Action not found" });
    }
  } catch (err) {
    return createResponse({ status: "error", message: err.toString() });
  }
}

// --- LOGIC FUNCTIONS (Refactored from previous version) ---
function getAppData() {
  const ss = getDb();
  const sheetP = ss.getSheetByName("Pegawai") || ss.insertSheet("Pegawai");
  const sheetF = ss.getSheetByName("Files") || ss.insertSheet("Files");
  
  if (sheetP.getLastRow() === 0) sheetP.appendRow(["ID", "Nama", "NIP", "Jabatan", "Dept", "FolderID", "CreatedAt"]);
  if (sheetF.getLastRow() === 0) sheetF.appendRow(["PegawaiID", "FileName", "DriveID", "URL", "Date"]);

  const dataP = sheetP.getDataRange().getValues();
  const dataF = sheetF.getDataRange().getValues();
  
  const pegawai = dataP.slice(1).map(row => {
    return {
      id: row[0],
      nama: row[1],
      nip: row[2],
      jabatan: row[3],
      dept: row[4],
      folderId: row[5],
      files: dataF.slice(1)
        .filter(f => f[0] == row[0])
        .map(f => ({ name: f[1], driveId: f[2], url: f[3], at: f[4] }))
    };
  });
  
  return { pegawai };
}

function addPegawai(obj) {
  const ss = getDb();
  const id = "P-" + new Date().getTime();
  const folder = getRootFolder().createFolder(obj.nama + " (" + obj.nip + ")");
  
  ss.getSheetByName("Pegawai").appendRow([
    id, obj.nama, obj.nip, obj.jabatan, obj.dept, folder.getId(), new Date()
  ]);
  
  return { id: id, folderId: folder.getId() };
}

function uploadFileToDrive(base64, fileName, pegawaiId, folderId) {
  const folder = DriveApp.getFolderById(folderId);
  const contentType = base64.substring(5, base64.indexOf(';'));
  const bytes = Utilities.base64Decode(base64.split(',')[1]);
  const blob = Utilities.newBlob(bytes, contentType, fileName);
  const file = folder.createFile(blob);
  
  getDb().getSheetByName("Files").appendRow([pegawaiId, fileName, file.getId(), file.getUrl(), new Date()]);
  
  return { 
    name: fileName, 
    url: file.getUrl(), 
    driveId: file.getId(), 
    at: new Date().getTime() 
  };
}

function deleteSingleFile(driveId) {
  DriveApp.getFileById(driveId).setTrashed(true);
  const sheet = getDb().getSheetByName("Files");
  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (data[i][2] == driveId) sheet.deleteRow(i + 1);
  }
  return { status: "success" };
}

function removePegawai(id, folderId) {
  const ss = getDb();
  try { DriveApp.getFolderById(folderId).setTrashed(true); } catch(e) {}
  
  const sheetP = ss.getSheetByName("Pegawai");
  const dataP = sheetP.getDataRange().getValues();
  for(let i = dataP.length - 1; i >= 1; i--) { 
    if(dataP[i][0] == id) sheetP.deleteRow(i + 1); 
  }
  
  const sheetF = ss.getSheetByName("Files");
  const dataF = sheetF.getDataRange().getValues();
  for(let i = dataF.length - 1; i >= 1; i--) {
    if(dataF[i][0] == id) sheetF.deleteRow(i + 1);
  }
  return { status: "success" };
}