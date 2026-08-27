const DEFAULT_FOLDER_ID = "1sU2_6KlvRSWZ3Rv-9bF9AEU7PvYBF4pJ";
const DEFAULT_DATABASE_FOLDER_ID = "1JH3z-QrsjhiHxc2h8IUGKTf-jxML1igj";
const DEFAULT_BACKUP_FOLDER_ID = "1JH3z-QrsjhiHxc2h8IUGKTf-jxML1igj";
const EMAIL_SENDER_NAME = "T23 Contract Tracking";
const LINE_CHANNEL_ACCESS_TOKEN_PROPERTY = "LINE_CHANNEL_ACCESS_TOKEN";
const LINE_GROUP_ID_PROPERTY = "LINE_GROUP_ID";
const LINE_WEBHOOK_KEY_PROPERTY = "LINE_WEBHOOK_KEY";
const LINE_NOTIFICATION_TIMEZONE = "Asia/Bangkok";
const LINE_NOTIFICATION_HOUR = 9;
const LINE_NOTIFICATION_MINUTE = 30;
const LINE_NOTIFICATION_HANDLER = "runLineStatusNotificationsScheduled";

function doPost(e) {
  let requestId = "";
  try {
    const rawPayload = (e.postData && e.postData.contents) || (e.parameter && e.parameter.payload) || "{}";
    const payload = JSON.parse(rawPayload);
    if (Array.isArray(payload.events)) return jsonResponse(handleLineWebhook_(payload, e));
    requestId = String(payload.requestId || "").trim();
    const mode = payload.mode || (payload.to ? "sendStatusEmail" : "uploadAttachment");
    if (mode === "sendStatusEmail") {
      setEmailRequestStatus_(requestId, { success: true, state: "processing" });
      const result = sendStatusEmail_(payload);
      setEmailRequestStatus_(requestId, Object.assign({ state: "sent" }, result));
      return jsonResponse(result);
    }
    if (mode === "sendLineStatusNotification") return jsonResponse(sendLineStatusNotification_(payload));
    if (mode === "runLineStatusNotifications") return jsonResponse(runLineStatusNotifications(payload));
    if (mode === "installLineStatusNotificationTrigger") return jsonResponse(installLineStatusNotificationTrigger_());
    if (mode === "saveDriveDatabase") return saveDriveDatabase_(payload);
    if (mode === "backupDriveDatabase") return backupDriveDatabase_(payload);
    if (mode === "installDailyBackup") return installDailyBackupTrigger_(payload);
    return jsonResponse({ success: true, files: [saveAttachment_(payload)] });
  } catch (error) {
    setEmailRequestStatus_(requestId, { success: false, state: "failed", error: errorMessage_(error) });
    console.error("T23 doPost failed: " + errorMessage_(error));
    return jsonResponse({ success: false, error: errorMessage_(error) });
  }
}

function doGet(e) {
  const params = (e && e.parameter) || {};
  const callback = String(params.callback || "").trim();
  try {
    if (params.mode === "loadDriveDatabase") return jsonpResponse(loadDriveDatabase_(params), callback);
    if (params.mode === "emailRequestStatus") return jsonpResponse(emailRequestStatus_(params.requestId), callback);
    if (params.mode === "healthCheck") return jsonpResponse(endpointHealthCheck_(), callback);
    return jsonpResponse({
      success: true,
      message: "T23 attachment upload, status email, and Drive database endpoint is running.",
      folderId: DEFAULT_FOLDER_ID,
      databaseFolderId: DEFAULT_DATABASE_FOLDER_ID,
      backupFolderId: DEFAULT_BACKUP_FOLDER_ID
    }, callback);
  } catch (error) {
    return jsonpResponse({ success: false, error: errorMessage_(error) }, callback);
  }
}

function emailRequestStatus_(requestId) {
  const id = String(requestId || "").trim();
  if (!id) return { success: false, state: "failed", error: "Missing request ID." };
  const cached = CacheService.getScriptCache().get("t23_email_" + id);
  return cached ? JSON.parse(cached) : { success: true, state: "pending" };
}

function setEmailRequestStatus_(requestId, status) {
  const id = String(requestId || "").trim();
  if (!id) return;
  CacheService.getScriptCache().put("t23_email_" + id, JSON.stringify(status || {}), 600);
}

function endpointHealthCheck_() {
  const attachmentFolder = DriveApp.getFolderById(DEFAULT_FOLDER_ID);
  const scriptProperties = PropertiesService.getScriptProperties();
  return {
    success: true,
    state: "ready",
    attachmentFolderId: attachmentFolder.getId(),
    attachmentFolderName: attachmentFolder.getName(),
    remainingMailQuota: MailApp.getRemainingDailyQuota(),
    lineConfigured: Boolean(scriptProperties.getProperty(LINE_CHANNEL_ACCESS_TOKEN_PROPERTY)),
    lineGroupConfigured: Boolean(scriptProperties.getProperty(LINE_GROUP_ID_PROPERTY)),
    lineWebhookConfigured: Boolean(scriptProperties.getProperty(LINE_WEBHOOK_KEY_PROPERTY)),
    checkedAt: new Date().toISOString()
  };
}

function handleLineWebhook_(payload, event) {
  const properties = PropertiesService.getScriptProperties();
  const expectedKey = String(properties.getProperty(LINE_WEBHOOK_KEY_PROPERTY) || "").trim();
  const providedKey = String(event && event.parameter && event.parameter.key || "").trim();
  if (!expectedKey || providedKey !== expectedKey) throw new Error("Invalid LINE webhook key.");

  let groupCaptured = false;
  (payload.events || []).forEach(function(lineEvent) {
    const source = lineEvent && lineEvent.source || {};
    const groupId = source.type === "group" ? String(source.groupId || "").trim() : "";
    if (!groupId) return;
    properties.setProperty(LINE_GROUP_ID_PROPERTY, groupId);
    properties.setProperty("LINE_GROUP_ID_CAPTURED_AT", new Date().toISOString());
    groupCaptured = true;
  });
  return {
    success: true,
    received: (payload.events || []).length,
    groupCaptured: groupCaptured
  };
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
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
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
  } finally {
    lock.releaseLock();
  }
}

function backupDriveDatabase_(payload) {
  const sourceFolder = DriveApp.getFolderById(payload.folderId || DEFAULT_DATABASE_FOLDER_ID);
  const backupFolder = DriveApp.getFolderById(payload.backupFolderId || DEFAULT_BACKUP_FOLDER_ID);
  const stamp = Utilities.formatDate(new Date(), "Etc/UTC", "yyyyMMdd-HHmmss");
  const prefix = payload.prefix || "daily_backup";
  const files = {
    contracts: backupTextFileByName_(sourceFolder, backupFolder, payload.contractsCsv || "tracking_contracts_contracts_db.csv", stamp, prefix),
    logs: backupTextFileByName_(sourceFolder, backupFolder, payload.logsCsv || "tracking_contracts_log_db.csv", stamp, prefix),
    typeMaster: backupTextFileByName_(sourceFolder, backupFolder, payload.typeMasterCsv || "tracking_contracts_type_master_db.csv", stamp, prefix),
    departments: backupTextFileByName_(sourceFolder, backupFolder, payload.departmentMasterCsv || "tracking_contracts_department_master_db.csv", stamp, prefix),
    people: backupTextFileByName_(sourceFolder, backupFolder, payload.peopleMasterCsv || "tracking_contracts_people_master_db.csv", stamp, prefix),
    contractTemplates: backupTextFileByName_(sourceFolder, backupFolder, payload.contractTemplateCsv || "tracking_contracts_contract_template_master_db.csv", stamp, prefix),
    actionSla: backupTextFileByName_(sourceFolder, backupFolder, payload.actionSlaCsv || "tracking_contracts_action_sla_master_db.csv", stamp, prefix)
  };
  return jsonResponse({
    success: true,
    backedUp: true,
    backedUpAt: new Date().toISOString(),
    sourceFolderId: sourceFolder.getId(),
    backupFolderId: backupFolder.getId(),
    files: files
  });
}

function runDailyBackup() {
  return backupDriveDatabase_({
    folderId: DEFAULT_DATABASE_FOLDER_ID,
    backupFolderId: DEFAULT_BACKUP_FOLDER_ID,
    prefix: "daily_backup"
  });
}

function installDailyBackupTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction && trigger.getHandlerFunction() === "runDailyBackup") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  const trigger = ScriptApp.newTrigger("runDailyBackup")
    .timeBased()
    .everyDays(1)
    .atHour(2)
    .create();
  return jsonResponse({
    success: true,
    installed: true,
    handlerFunction: trigger.getHandlerFunction(),
    backupFolderId: DEFAULT_BACKUP_FOLDER_ID
  });
}

function readTextFileByName_(folder, fileName) {
  const file = latestFileByName_(folder, fileName);
  if (!file) return "";
  return file.getBlob().getDataAsString("UTF-8").replace(/^\uFEFF/, "");
}

function latestFileByName_(folder, fileName) {
  const files = folder.getFilesByName(cleanFileName_(fileName));
  let latest = null;
  while (files.hasNext()) {
    const candidate = files.next();
    if (!latest || candidate.getLastUpdated().getTime() > latest.getLastUpdated().getTime()) latest = candidate;
  }
  return latest;
}

function backupTextFileByName_(sourceFolder, backupFolder, fileName, stamp, prefix) {
  const name = cleanFileName_(fileName || "database.csv");
  const sourceFile = latestFileByName_(sourceFolder, name);
  if (!sourceFile) {
    return {
      sourceName: name,
      skipped: true,
      reason: "Source file not found"
    };
  }
  const backupName = cleanFileName_([prefix, stamp, name].join("_"));
  const backupFile = backupFolder.createFile(backupName, sourceFile.getBlob().getDataAsString("UTF-8"), MimeType.CSV);
  applyBestEffortFileSharing_(backupFile);
  return {
    id: backupFile.getId(),
    sourceName: name,
    name: backupFile.getName(),
    url: backupFile.getUrl(),
    downloadUrl: "https://drive.google.com/uc?export=download&id=" + backupFile.getId()
  };
}

function upsertTextFileByName_(folder, fileName, text) {
  const name = cleanFileName_(fileName || "database.csv");
  const content = String(text || "");
  const files = folder.getFilesByName(name);
  const matches = [];
  while (files.hasNext()) matches.push(files.next());
  matches.sort(function(a, b) { return b.getLastUpdated().getTime() - a.getLastUpdated().getTime(); });
  const file = matches.length ? matches[0].setContent(content) : folder.createFile(name, content, MimeType.CSV);
  matches.slice(1).forEach(function(duplicate) {
    try { duplicate.setTrashed(true); } catch (error) { console.warn("Duplicate CSV could not be trashed: " + errorMessage_(error)); }
  });
  applyBestEffortFileSharing_(file);
  return {
    id: file.getId(),
    name: file.getName(),
    updatedAt: file.getLastUpdated().toISOString(),
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
  return {
    success: true,
    sent: true,
    sentAt: new Date().toISOString(),
    to: to,
    cc: cc,
    files: files,
    attachedFiles: mailAttachments.map(function(blob) { return blob.getName(); })
  };
}

function sendLineStatusNotification_(payload) {
  const to = String(payload.to || payload.lineUserId || "").trim();
  const message = String(payload.message || "").trim();
  if (!to) throw new Error("Missing LINE User ID.");
  if (!message) throw new Error("Missing LINE notification message.");
  return pushLineMessage_(to, message);
}

function pushLineMessage_(to, message) {
  const token = String(PropertiesService.getScriptProperties().getProperty(LINE_CHANNEL_ACCESS_TOKEN_PROPERTY) || "").trim();
  if (!token) throw new Error("LINE_CHANNEL_ACCESS_TOKEN is not configured in Script Properties.");
  const messages = Array.isArray(message) ? message : [message];
  const normalizedMessages = messages.map(function(item) {
    if (item && typeof item === "object") return item;
    return { type: "text", text: String(item || "").slice(0, 5000) };
  });
  const response = UrlFetchApp.fetch("https://api.line.me/v2/bot/message/push", {
    method: "post",
    contentType: "application/json",
    headers: { Authorization: "Bearer " + token },
    payload: JSON.stringify({
      to: String(to || "").trim(),
      messages: normalizedMessages
    }),
    muteHttpExceptions: true
  });
  const responseCode = response.getResponseCode();
  if (responseCode < 200 || responseCode >= 300) {
    throw new Error("LINE Messaging API returned HTTP " + responseCode + ": " + response.getContentText());
  }
  return {
    success: true,
    sent: true,
    sentAt: new Date().toISOString(),
    to: String(to || "").trim()
  };
}

function runLineStatusNotifications(options) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) {
    return { success: true, sent: 0, skipped: 0, failed: 0, busy: true, reason: "Another LINE notification run is in progress." };
  }
  try {
    return runLineStatusNotificationsUnlocked_(options);
  } finally {
    lock.releaseLock();
  }
}

function runLineStatusNotificationsUnlocked_(options) {
  const settings = options || {};
  const dryRun = settings.dryRun === true || String(settings.dryRun || "").toLowerCase() === "true";
  const forceSend = settings.forceSend === true || String(settings.forceSend || "").toLowerCase() === "true";
  const folder = DriveApp.getFolderById(settings.folderId || DEFAULT_DATABASE_FOLDER_ID);
  const contracts = csvObjects_(readTextFileByName_(folder, settings.contractsCsv || "tracking_contracts_contracts_db.csv"));
  const people = csvObjects_(readTextFileByName_(folder, settings.peopleMasterCsv || "tracking_contracts_people_master_db.csv"));
  const owners = lineOwnerMap_(people);
  const properties = PropertiesService.getScriptProperties();
  const lineGroupId = String(properties.getProperty(LINE_GROUP_ID_PROPERTY) || "").trim();
  const today = Utilities.formatDate(new Date(), LINE_NOTIFICATION_TIMEZONE, "yyyy-MM-dd");
  const results = [];
  const candidates = [];
  let sent = 0;
  let skipped = 0;
  let failed = 0;

  contracts.forEach(function(contract) {
    if (lineContractIsClosed_(contract)) return;
    const statusCode = lineStatusCode_(contract["Status Update"]);
    if (statusCode !== "Y" && statusCode !== "R") return;

    const contractId = String(contract["Contract ID"] || "").trim();
    const ownerName = String(contract["Contract Owner"] || "").trim();
    const owner = owners[normalizeLineLookup_(ownerName)] || null;
    const lineUserId = String(owner && (owner.lineUserId || owner["LINE User ID"]) || "").trim();
    const lineEnabled = owner && masterFlagEnabled_(owner.lineNotifications || owner["LINE Alert"] || "Yes");
    const lineRecipient = lineGroupId || lineUserId;
    const recipientType = lineGroupId ? "group" : "owner";
    const dedupeKey = "t23_line_status_" + contractId;
    const dedupeValue = today;

    if (!contractId) {
      skipped += 1;
      results.push({ contractId: "", statusCode: statusCode, sent: false, reason: "Missing Contract ID." });
      return;
    }
    if (!lineGroupId && !owner) {
      skipped += 1;
      results.push({ contractId: contractId, statusCode: statusCode, sent: false, reason: "Contract Owner not found in People Master." });
      return;
    }
    if (!lineGroupId && !lineEnabled) {
      skipped += 1;
      results.push({ contractId: contractId, statusCode: statusCode, sent: false, reason: "LINE Alert disabled for Contract Owner." });
      return;
    }
    if (!lineRecipient) {
      skipped += 1;
      results.push({ contractId: contractId, statusCode: statusCode, sent: false, reason: "Missing LINE Group ID or LINE User ID." });
      return;
    }
    // A contract can generate at most one LINE notification per calendar day.
    if (!dryRun && !forceSend && properties.getProperty(dedupeKey) === dedupeValue) {
      skipped += 1;
      results.push({ contractId: contractId, statusCode: statusCode, sent: false, reason: "This contract was already sent today." });
      return;
    }

    candidates.push({
      contract: contract,
      contractId: contractId,
      statusCode: statusCode,
      ownerName: ownerName || "Unassigned",
      to: lineRecipient,
      recipientType: recipientType,
      dedupeKey: dedupeKey,
      dedupeValue: dedupeValue
    });
  });

  const batches = lineFlexNotificationBatches_(candidates);
  if (dryRun) {
    candidates.forEach(function(candidate) {
      results.push({
        contractId: candidate.contractId,
        statusCode: candidate.statusCode,
        sent: false,
        dryRun: true,
        to: candidate.to,
        recipientType: candidate.recipientType,
        format: "flex"
      });
    });
  } else {
    batches.forEach(function(batch) {
      try {
        pushLineMessage_(batch.to, batch.messages || batch.message);
        batch.candidates.forEach(function(candidate) {
          properties.setProperty(candidate.dedupeKey, candidate.dedupeValue);
          sent += 1;
          results.push({
            contractId: candidate.contractId,
            statusCode: candidate.statusCode,
            sent: true,
            to: candidate.to,
            recipientType: candidate.recipientType,
            format: "flex"
          });
        });
      } catch (error) {
        batch.candidates.forEach(function(candidate) {
          failed += 1;
          results.push({ contractId: candidate.contractId, statusCode: candidate.statusCode, sent: false, reason: errorMessage_(error) });
        });
      }
    });
  }

  return {
    success: failed === 0,
    dryRun: dryRun,
    forceSend: forceSend,
    sent: sent,
    skipped: skipped,
    failed: failed,
    checked: results.length,
    flexBatches: batches.length,
    flexMessages: batches.reduce(function(total, batch) { return total + ((batch.messages && batch.messages.length) || 1); }, 0),
    runAt: new Date().toISOString(),
    results: results
  };
}

function lineFlexNotificationBatches_(candidates) {
  const byRecipient = {};
  (candidates || []).forEach(function(candidate) {
    const key = candidate.to;
    if (!byRecipient[key]) byRecipient[key] = { recipientType: candidate.recipientType, candidates: [] };
    byRecipient[key].candidates.push(candidate);
  });

  const batches = [];
  Object.keys(byRecipient).forEach(function(to) {
    const recipient = byRecipient[to];
    const grouped = {};
    recipient.candidates.forEach(function(candidate) {
      const groupKey = candidate.statusCode + "|" + candidate.ownerName;
      if (!grouped[groupKey]) grouped[groupKey] = { statusCode: candidate.statusCode, ownerName: candidate.ownerName, candidates: [] };
      grouped[groupKey].candidates.push(candidate);
    });

    const pages = [];
    Object.keys(grouped).sort().forEach(function(groupKey) {
      const group = grouped[groupKey];
      group.candidates.sort(function(a, b) {
        return Number(b.contract["Days Used"] || 0) - Number(a.contract["Days Used"] || 0) || a.contractId.localeCompare(b.contractId);
      });
      const pageCount = Math.ceil(group.candidates.length / 4);
      for (let index = 0; index < group.candidates.length; index += 4) {
        const pageCandidates = group.candidates.slice(index, index + 4);
        pages.push({
          bubble: lineFlexStatusBubble_(group.statusCode, group.ownerName, pageCandidates, Math.floor(index / 4) + 1, pageCount, recipient.candidates),
          candidates: pageCandidates
        });
      }
    });

    for (let pageIndex = 0; pageIndex < pages.length; pageIndex += 10) {
      const pageChunk = pages.slice(pageIndex, pageIndex + 10);
      const batchCandidates = [];
      pageChunk.forEach(function(page) { Array.prototype.push.apply(batchCandidates, page.candidates); });
      const statusCodes = {};
      batchCandidates.forEach(function(candidate) { statusCodes[candidate.statusCode] = true; });
      const statusText = Object.keys(statusCodes).sort().join("/");
      const tableMessage = {
        type: "flex",
        altText: "[" + statusText + "] Contract Status Update - " + batchCandidates.length + " contract(s)",
        contents: { type: "carousel", contents: pageChunk.map(function(page) { return page.bubble; }) }
      };
      const messages = [tableMessage];
      if (pageIndex + 10 >= pages.length) messages.push(lineFlexOwnerStatusSummaryMessage_(recipient.candidates));
      batches.push({
        to: to,
        recipientType: recipient.recipientType,
        candidates: batchCandidates,
        messages: messages
      });
    }
  });
  return batches;
}

function lineFlexOwnerStatusSummaryMessage_(candidates) {
  const grouped = {};
  (candidates || []).forEach(function(candidate) {
    const ownerName = candidate.ownerName || "Unassigned";
    if (!grouped[ownerName]) grouped[ownerName] = { ownerName: ownerName, delayed: 0, risk: 0 };
    if (candidate.statusCode === "Y") grouped[ownerName].delayed += 1;
    if (candidate.statusCode === "R") grouped[ownerName].risk += 1;
  });
  const owners = Object.keys(grouped).map(function(key) { return grouped[key]; }).sort(function(a, b) {
    return (b.delayed + b.risk) - (a.delayed + a.risk) || b.risk - a.risk || a.ownerName.localeCompare(b.ownerName);
  });
  const maxTotal = Math.max.apply(null, owners.map(function(owner) { return owner.delayed + owner.risk; }).concat([1]));
  const bubbles = [];
  for (let index = 0; index < owners.length; index += 6) {
    bubbles.push(lineFlexOwnerStatusSummaryBubble_(owners.slice(index, index + 6), maxTotal, Math.floor(index / 6) + 1, Math.ceil(owners.length / 6)));
  }
  return {
    type: "flex",
    altText: "By Person - Station Owner Summary - Delayed and At Risk",
    contents: { type: "carousel", contents: bubbles }
  };
}

function lineFlexOwnerStatusSummaryBubble_(owners, maxTotal, pageNumber, pageCount) {
  const rows = [];
  (owners || []).forEach(function(owner, index) {
    if (index) rows.push({ type: "separator", color: "#ECEEEF", margin: "md" });
    rows.push({
      type: "box",
      layout: "vertical",
      margin: index ? "md" : "none",
      contents: [
        { type: "text", text: lineFlexText_(owner.ownerName, 80), size: "xs", weight: "bold", color: "#202124", wrap: true, maxLines: 2 },
        { type: "box", layout: "horizontal", height: "26px", margin: "sm", backgroundColor: "#E5EAEE", cornerRadius: "md", contents: lineFlexOwnerStatusBar_(owner, maxTotal) }
      ]
    });
  });
  return {
    type: "bubble",
    size: "giga",
    header: {
      type: "box",
      layout: "vertical",
      paddingAll: "16px",
      backgroundColor: "#F5F6F7",
      contents: [
        { type: "text", text: "By Person — Station Owner Summary", size: "lg", weight: "bold", color: "#202124", wrap: true },
        { type: "text", text: "สรุปสถานะตาม Station Owner", size: "xs", color: "#6F7478", margin: "sm" }
      ]
    },
    body: { type: "box", layout: "vertical", paddingAll: "16px", contents: rows },
    footer: {
      type: "box",
      layout: "horizontal",
      paddingAll: "12px",
      contents: [
        { type: "text", text: "Y = Delayed", size: "xxs", color: "#C58A00", weight: "bold", flex: 1 },
        { type: "text", text: "R = At Risk", size: "xxs", color: "#CE3D34", weight: "bold", flex: 1 },
        { type: "text", text: pageCount > 1 ? pageNumber + "/" + pageCount : "Summary", size: "xxs", color: "#777C80", align: "end", flex: 1 }
      ]
    }
  };
}

function lineFlexOwnerStatusBar_(owner, maxTotal) {
  const contents = [];
  if (owner.delayed > 0) contents.push({
    type: "box", layout: "vertical", flex: owner.delayed, backgroundColor: "#C58A00", justifyContent: "center",
    contents: [{ type: "text", text: String(owner.delayed), size: "xxs", color: "#FFFFFF", weight: "bold", align: "center" }]
  });
  if (owner.risk > 0) contents.push({
    type: "box", layout: "vertical", flex: owner.risk, backgroundColor: "#CE3D34", justifyContent: "center",
    contents: [{ type: "text", text: String(owner.risk), size: "xxs", color: "#FFFFFF", weight: "bold", align: "center" }]
  });
  const remainder = Math.max(0, maxTotal - owner.delayed - owner.risk);
  if (remainder > 0) contents.push({ type: "box", layout: "vertical", flex: remainder, contents: [] });
  return contents;
}

function lineFlexStatusBubble_(statusCode, ownerName, pageCandidates, pageNumber, pageCount, allCandidates) {
  const isOverdue = statusCode === "R";
  const statusEnglish = isOverdue ? "Overdue Contracts" : "Delayed Contracts";
  const statusThai = isOverdue ? "สัญญาที่เกินกำหนด" : "สัญญาที่ถึงช่วงติดตาม";
  const accent = isOverdue ? "#C62828" : "#C88A00";
  const soft = isOverdue ? "#FFF0F0" : "#FFF8E1";
  const statusTotal = (allCandidates || []).filter(function(candidate) { return candidate.statusCode === statusCode; }).length;
  const rows = [];
  rows.push(lineFlexTableHeader_());
  pageCandidates.forEach(function(candidate) {
    rows.push({ type: "separator", color: "#ECEEEF" });
    rows.push(lineFlexContractRow_(candidate, accent));
  });
  return {
    type: "bubble",
    size: "giga",
    header: {
      type: "box",
      layout: "horizontal",
      backgroundColor: soft,
      paddingAll: "16px",
      contents: [
        {
          type: "box",
          layout: "vertical",
          flex: 1,
          contents: [
            { type: "text", text: "[" + statusCode + "] " + statusEnglish, color: accent, weight: "bold", size: "lg" },
            { type: "text", text: statusThai, color: accent, size: "xs", margin: "sm" }
          ]
        },
        {
          type: "box",
          layout: "vertical",
          width: "40px",
          flex: 0,
          backgroundColor: accent,
          cornerRadius: "md",
          paddingAll: "6px",
          justifyContent: "center",
          contents: [{ type: "text", text: String(statusTotal), color: "#FFFFFF", weight: "bold", size: "sm", align: "center" }]
        }
      ]
    },
    body: {
      type: "box",
      layout: "vertical",
      paddingAll: "0px",
      contents: [
        {
          type: "box",
          layout: "vertical",
          paddingAll: "14px",
          contents: [
            { type: "text", text: "CONTRACT OWNER", color: "#777C80", size: "xxs", weight: "bold" },
            { type: "text", text: lineFlexText_(ownerName, 80), color: "#202124", size: "sm", weight: "bold", margin: "sm", wrap: true }
          ]
        },
        { type: "separator", color: "#DDE1E4" },
        { type: "box", layout: "vertical", contents: rows }
      ]
    },
    footer: {
      type: "box",
      layout: "horizontal",
      paddingAll: "12px",
      contents: [
        { type: "text", text: "Accumulated Days = Working Days", size: "xxs", color: "#777C80", flex: 1 },
        { type: "text", text: pageCount > 1 ? pageNumber + "/" + pageCount : "Status Summary", size: "xxs", color: "#777C80", align: "end" }
      ]
    }
  };
}

function lineFlexTableHeader_() {
  return {
    type: "box",
    layout: "horizontal",
    backgroundColor: "#F5F6F7",
    paddingAll: "9px",
    spacing: "sm",
    contents: [
      { type: "text", text: "CONTRACT ID", size: "xxs", color: "#777C80", weight: "bold", flex: 3 },
      { type: "text", text: "CONTRACT NAME", size: "xxs", color: "#777C80", weight: "bold", flex: 5 },
      { type: "text", text: "DAYS", size: "xxs", color: "#777C80", weight: "bold", align: "end", flex: 2 },
      { type: "text", text: "DUE DATE", size: "xxs", color: "#777C80", weight: "bold", align: "end", flex: 3 }
    ]
  };
}

function lineFlexContractRow_(candidate, accent) {
  const contract = candidate.contract || {};
  const confidential = /confidential|สัญญาลับ/i.test([contract["Access Level"], contract.Visibility, contract.Category].join(" "));
  const contractName = confidential ? "Confidential Contract / สัญญาลับ" : String(contract["Contract Name"] || "-");
  return {
    type: "box",
    layout: "horizontal",
    paddingAll: "10px",
    spacing: "sm",
    contents: [
      { type: "text", text: lineFlexText_(candidate.contractId, 28), size: "xxs", color: "#1667A8", weight: "bold", wrap: true, flex: 3 },
      { type: "text", text: lineFlexText_(contractName, 100), size: "xxs", color: "#202124", weight: "bold", wrap: true, maxLines: 3, flex: 5 },
      { type: "text", text: String(Number(contract["Days Used"] || 0)), size: "xxs", color: accent, weight: "bold", align: "end", flex: 2 },
      { type: "text", text: lineFlexText_(contract["Due Date"] || "-", 20), size: "xxs", color: "#4D5357", align: "end", wrap: true, flex: 3 }
    ]
  };
}

function lineFlexText_(value, limit) {
  return String(value == null ? "" : value).replace(/[\u0000-\u001F\u007F]/g, " ").trim().slice(0, limit || 100);
}

function previewLineStatusNotifications() {
  return runLineStatusNotifications({ dryRun: true });
}

function runLineStatusNotificationsScheduled() {
  const weekday = Number(Utilities.formatDate(new Date(), LINE_NOTIFICATION_TIMEZONE, "u"));
  if (weekday > 5) {
    return {
      success: true,
      scheduled: true,
      skippedWeekend: true,
      sent: 0,
      skipped: 0,
      failed: 0,
      runAt: new Date().toISOString()
    };
  }
  return runLineStatusNotifications({ source: "weekdayFallback0930" });
}

function installLineStatusNotificationTrigger() {
  return installLineStatusNotificationTrigger_();
}

function installLineStatusNotificationTrigger_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (!trigger.getHandlerFunction) return;
    const handler = trigger.getHandlerFunction();
    if (handler === LINE_NOTIFICATION_HANDLER || handler === "runLineStatusNotifications") ScriptApp.deleteTrigger(trigger);
  });
  const trigger = ScriptApp.newTrigger(LINE_NOTIFICATION_HANDLER)
    .timeBased()
    .everyDays(1)
    .atHour(LINE_NOTIFICATION_HOUR)
    .nearMinute(LINE_NOTIFICATION_MINUTE)
    .inTimezone(LINE_NOTIFICATION_TIMEZONE)
    .create();
  return {
    success: true,
    installed: true,
    handlerFunction: trigger.getHandlerFunction(),
    hour: LINE_NOTIFICATION_HOUR,
    minute: LINE_NOTIFICATION_MINUTE,
    timezone: LINE_NOTIFICATION_TIMEZONE,
    weekdaysOnly: true
  };
}

function csvObjects_(text) {
  const source = String(text || "").replace(/^\uFEFF/, "");
  if (!source.trim()) return [];
  const rows = Utilities.parseCsv(source);
  if (!rows.length) return [];
  const headers = rows.shift().map(function(header) { return String(header || "").trim(); });
  return rows.filter(function(row) {
    return row.some(function(value) { return String(value || "").trim(); });
  }).map(function(row) {
    const object = {};
    headers.forEach(function(header, index) { object[header] = row[index] == null ? "" : row[index]; });
    return object;
  });
}

function lineOwnerMap_(people) {
  const map = {};
  (people || []).forEach(function(person) {
    const key = normalizeLineLookup_(person.name || person.Name);
    if (key && masterFlagEnabled_(person.active || person.Active || "Yes")) map[key] = person;
  });
  return map;
}

function normalizeLineLookup_(value) {
  return String(value || "").trim().replace(/\s+/g, " ").toLowerCase();
}

function masterFlagEnabled_(value) {
  return !/^(no|false|0|inactive|disabled)$/i.test(String(value == null ? "Yes" : value).trim());
}

function lineContractIsClosed_(contract) {
  const text = [contract.Stage, contract["Status Update"]].join(" ");
  return /\b(cancelled|completed|signed)\b/i.test(text) || lineStatusCode_(contract["Status Update"]) === "B";
}

function lineStatusCode_(value) {
  const text = String(value || "").trim();
  if (/\b(cancelled|completed|signed)\b/i.test(text) || /(^|[^A-Z])B\s*=/i.test(text)) return "B";
  if (/\boverdue\b|\bred\b/i.test(text) || /(^|[^A-Z])R\s*=/i.test(text)) return "R";
  if (/\bdelayed\b|\byellow\b/i.test(text) || /(^|[^A-Z])Y\s*=/i.test(text)) return "Y";
  return "G";
}

function lineNotificationMessage_(contract, statusCode) {
  const confidential = /confidential|สัญญาลับ/i.test([contract["Access Level"], contract.Visibility, contract.Category].join(" "));
  const statusEnglish = statusCode === "R" ? "Overdue" : "Delayed";
  const statusThai = statusCode === "R" ? "เกิน SLA รวม" : "ถึงช่วงติดตาม SLA รวม";
  const contractName = confidential ? "Confidential Contract / สัญญาลับ" : String(contract["Contract Name"] || "-");
  return [
    "[" + statusCode + "] Contract Status Update: " + statusEnglish,
    "สถานะสัญญา: " + statusThai,
    "",
    "Contract ID: " + String(contract["Contract ID"] || "-"),
    "Contract Name: " + contractName,
    "Contract Owner: " + String(contract["Contract Owner"] || "-"),
    "Due Date: " + String(contract["Due Date"] || "-"),
    "",
    statusCode === "R"
      ? "Please update the action plan immediately. / กรุณาอัปเดตแผนดำเนินการทันที"
      : "Please review and update the contract status. / กรุณาตรวจสอบและอัปเดตสถานะสัญญา"
  ].join("\n");
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
  applyBestEffortFileSharing_(file);
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

function applyBestEffortFileSharing_(file) {
  try {
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    return "anyone";
  } catch (publicSharingError) {
    try {
      file.setSharing(DriveApp.Access.DOMAIN_WITH_LINK, DriveApp.Permission.VIEW);
      return "domain";
    } catch (domainSharingError) {
      console.warn("File created without changing sharing policy: " + errorMessage_(domainSharingError));
      return "existing";
    }
  }
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
