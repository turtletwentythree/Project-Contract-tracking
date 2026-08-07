const DEFAULT_FOLDER_ID = "13d9ZNEE0ijV5JfU7chg0wHrTNp500AgR";
const EMAIL_SENDER_NAME = "T23 Contract Tracking";

function doPost(e) {
  try {
    const payload = JSON.parse((e.postData && e.postData.contents) || "{}");
    const mode = payload.mode || (payload.to ? "sendStatusEmail" : "uploadAttachment");
    if (mode === "sendStatusEmail") return sendStatusEmail_(payload);
    if (mode === "saveDriveDatabase") return saveDriveDatabase_(payload);
    return jsonResponse({ success: true, files: [saveAttachment_(payload)] });
  } catch (error) {
    return jsonResponse({ success: false, error: errorMessage_(error) });
  }
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  const callback = String(params.callback || "").trim();
  try {
    if (params.mode === "loadDriveDatabase") return jsonpResponse(loadDriveDatabase_(params), callback);
    return jsonpResponse({
      success: true,
      message: "T23 attachment upload, status email, and Drive database endpoint is running.",
      folderId: DEFAULT_FOLDER_ID
    }, callback);
  } catch (error) {
    return jsonpResponse({ success: false, error: errorMessage_(error) }, callback);
  }
}

function loadDriveDatabase_(params) {
  const folder = DriveApp.getFolderById(params.folderId || DEFAULT_FOLDER_ID);
  return {
    success: true,
    folderId: folder.getId(),
    loadedAt: new Date().toISOString(),
    contractsCsvText: readTextFileByName_(folder, params.contractsCsv || "tracking_contracts_contracts_db.csv"),
    logsCsvText: readTextFileByName_(folder, params.logsCsv || "tracking_contracts_log_db.csv"),
    typeMasterCsvText: readTextFileByName_(folder, params.typeMasterCsv || "tracking_contracts_type_master_db.csv"),
	    departmentMasterCsvText: readTextFileByName_(folder, params.departmentMasterCsv || "tracking_contracts_department_master_db.csv"),
	    peopleMasterCsvText: readTextFileByName_(folder, params.peopleMasterCsv || "tracking_contracts_people_master_db.csv"),
	    contractTemplateCsvText: readTextFileByName_(folder, params.contractTemplateCsv || "tracking_contracts_contract_template_master_db.csv"),
	    actionSlaCsvText: readTextFileByName_(folder, params.actionSlaCsv || "tracking_contracts_action_sla_master_db.csv")
  };
}

function saveDriveDatabase_(payload) {
  const folder = DriveApp.getFolderById(payload.folderId || DEFAULT_FOLDER_ID);
  const files = {
    contracts: upsertTextFileByName_(folder, payload.contractsCsv || "tracking_contracts_contracts_db.csv", payload.contractsCsvText || ""),
    logs: upsertTextFileByName_(folder, payload.logsCsv || "tracking_contracts_log_db.csv", payload.logsCsvText || ""),
    typeMaster: upsertTextFileByName_(folder, payload.typeMasterCsv || "tracking_contracts_type_master_db.csv", payload.typeMasterCsvText || ""),
	    departments: upsertTextFileByName_(folder, payload.departmentMasterCsv || "tracking_contracts_department_master_db.csv", payload.departmentMasterCsvText || ""),
	    people: upsertTextFileByName_(folder, payload.peopleMasterCsv || "tracking_contracts_people_master_db.csv", payload.peopleMasterCsvText || ""),
	    contractTemplates: upsertTextFileByName_(folder, payload.contractTemplateCsv || "tracking_contracts_contract_template_master_db.csv", payload.contractTemplateCsvText || ""),
	    actionSla: upsertTextFileByName_(folder, payload.actionSlaCsv || "tracking_contracts_action_sla_master_db.csv", payload.actionSlaCsvText || "")
  };
  return jsonResponse({
    success: true,
    saved: true,
    savedAt: new Date().toISOString(),
    folderId: folder.getId(),
    files: files
  });
}

function readTextFileByName_(folder, fileName) {
  const files = folder.getFilesByName(fileName);
  if (!files.hasNext()) return "";
  return files.next().getBlob().getDataAsString("UTF-8").replace(/^\uFEFF/, "");
}

function upsertTextFileByName_(folder, fileName, text) {
  const name = cleanFileName_(fileName || "database.csv");
  const content = String(text || "");
  const files = folder.getFilesByName(name);
  const file = files.hasNext()
    ? files.next().setContent(content)
    : folder.createFile(name, content, MimeType.CSV);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {
    id: file.getId(),
    name: file.getName(),
    url: file.getUrl(),
    downloadUrl: "https://drive.google.com/uc?export=download&id=" + file.getId()
  };
}

function sendStatusEmail_(payload) {
  const to = String(payload.to || "").trim();
  if (!to) throw new Error("Missing recipient email.");
  if (!isValidEmailList_(to)) throw new Error("Invalid recipient email.");

  const files = (payload.attachments || []).map(function(attachment) {
    return saveAttachment_(Object.assign({ folderId: payload.folderId }, attachment));
  });
  const mailAttachments = buildMailAttachmentBlobs_(payload.attachments || []);
  const body = buildEmailBody_(payload.body || "", files, payload.folderUrl || "");
  const cc = normalizeCc_(payload.cc || payload.ccText || "");
  if (cc && !isValidEmailList_(cc)) throw new Error("Invalid CC email.");
  const options = {
    to: to,
    subject: payload.subject || "Contract Status Update",
    body: body,
    name: EMAIL_SENDER_NAME,
    htmlBody: buildHtmlBody_(body)
  };
  if (cc) options.cc = cc;
  if (mailAttachments.length) options.attachments = mailAttachments;

  MailApp.sendEmail(options);
  return jsonResponse({
    success: true,
    sent: true,
    sentAt: new Date().toISOString(),
    to: to,
    cc: cc,
    files: files,
    attachedFiles: mailAttachments.map(function(blob) { return blob.getName(); })
  });
}

function normalizeCc_(value) {
  const list = Array.isArray(value) ? value : String(value || "").split(/[;,\n]+/);
  const seen = {};
  return list
    .map(function(item) {
      return String(item && item.email ? item.email : item || "").trim();
    })
    .filter(function(email) {
      if (!email || seen[email.toLowerCase()]) return false;
      seen[email.toLowerCase()] = true;
      return true;
    })
    .join(", ");
}

function isValidEmailList_(value) {
  const emails = String(value || "").split(/[;,\n]+/).map(function(item) {
    return item.trim();
  }).filter(Boolean);
  if (!emails.length) return false;
  return emails.every(function(email) {
    return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
  });
}

function buildMailAttachmentBlobs_(attachments) {
  return attachments
    .filter(function(attachment) { return attachment && attachment.base64; })
    .map(function(attachment) {
      const fileName = cleanFileName_(attachment.originalFileName || attachment.fileName || "attachment");
      const mimeType = attachment.mimeType || "application/octet-stream";
      return Utilities.newBlob(Utilities.base64Decode(attachment.base64), mimeType, fileName);
    });
}

function saveAttachment_(payload) {
  const existingUrl = reusableDriveFileUrl_(payload.cloudUrl || payload.url || "") || "";
  if (existingUrl) {
    return {
      id: payload.cloudFileId || "",
      fileName: payload.fileName || payload.originalFileName || "attachment",
      originalFileName: payload.originalFileName || payload.fileName || "attachment",
      mimeType: payload.mimeType || "",
      fileSize: payload.fileSize || 0,
      url: existingUrl,
      downloadUrl: payload.downloadUrl || existingUrl,
      reused: true
    };
  }

  const base64 = payload.base64 || "";
  if (!base64) {
    return {
      id: "",
      fileName: payload.fileName || payload.originalFileName || "attachment",
      originalFileName: payload.originalFileName || payload.fileName || "attachment",
      mimeType: payload.mimeType || "",
      fileSize: payload.fileSize || 0,
      url: "",
      downloadUrl: "",
      skipped: true
    };
  }

  const folder = DriveApp.getFolderById(payload.folderId || DEFAULT_FOLDER_ID);
  const originalFileName = payload.originalFileName || payload.fileName || "attachment";
  const fileName = cleanFileName_(payload.fileName || originalFileName);
  const mimeType = payload.mimeType || "application/octet-stream";
  const blob = Utilities.newBlob(Utilities.base64Decode(base64), mimeType, fileName);
  const file = folder.createFile(blob);
  file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  return {
    id: file.getId(),
    fileName: file.getName(),
    originalFileName: originalFileName,
    mimeType: mimeType,
    fileSize: payload.fileSize || blob.getBytes().length,
    url: file.getUrl(),
    downloadUrl: "https://drive.google.com/uc?export=download&id=" + file.getId(),
    reused: false
  };
}

function buildEmailBody_(baseBody, files, folderUrl) {
  const usableFiles = files.filter(function(file) { return file.url; });
  const lines = [String(baseBody || "")];
  if (usableFiles.length) {
    lines.push("", "Cloud Attachments / ไฟล์แนบบน Cloud");
    usableFiles.forEach(function(file, index) {
      lines.push(
        String(index + 1) + ". " + (file.originalFileName || file.fileName),
        "   Open: " + file.url,
        "   Download: " + (file.downloadUrl || file.url)
      );
    });
  }
  return lines.join("\n");
}

function buildHtmlBody_(body) {
  return String(body || "")
    .split("\n")
    .map(function(line) {
      return escapeHtml_(line).replace(/(https:\/\/[^\s<>"']+)/g, function(url) {
        return '<a href="' + url + '" target="_blank" style="color:#1155cc;text-decoration:underline;">' + url + '</a>';
      });
    })
    .join("<br>");
}

function escapeHtml_(value) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");
}

function cleanFileName_(name) {
  const cleaned = String(name || "attachment").replace(/[\\/:*?"<>|]+/g, "_").slice(0, 180);
  return cleaned || "attachment";
}

function reusableDriveFileUrl_(url) {
  const text = String(url || "").trim();
  if (!text) return "";
  if (/^https:\/\/drive\.google\.com\/file\/d\//.test(text)) return text;
  if (/^https:\/\/drive\.google\.com\/uc\?/.test(text)) return text;
  if (/^https:\/\/docs\.google\.com\//.test(text)) return text;
  return "";
}

function errorMessage_(error) {
  return String((error && error.message) || error || "Unknown error");
}

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonpResponse(data, callback) {
  const safeCallback = String(callback || "").trim();
  if (safeCallback && /^[A-Za-z_$][\w$]*(\.[A-Za-z_$][\w$]*)*$/.test(safeCallback)) {
    return ContentService
      .createTextOutput(safeCallback + "(" + JSON.stringify(data) + ");")
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }
  return jsonResponse(data);
}
