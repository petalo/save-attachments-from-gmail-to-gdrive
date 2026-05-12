/**
 * Attachment processing functions for Gmail Attachment Organizer
 */

/**
 * Builds description metadata to persist source linkage in the Drive file.
 * Stored for human readability; NOT used for dedup search (description is
 * not a valid Drive API query field — use custom properties instead).
 *
 * @param {Date} emailDate - Original email date
 * @param {string|null} sourceAttachmentId - Deterministic source attachment ID
 * @returns {string} Metadata description string
 */
function buildAttachmentMetadata(emailDate, sourceAttachmentId) {
  const parts = [`email_date=${emailDate.toISOString()}`];
  if (sourceAttachmentId) {
    parts.push(`source_attachment_id=${sourceAttachmentId}`);
  }
  return parts.join("; ");
}

/**
 * Stamps a saved Drive file with two custom properties so future runs can
 * locate it via a queryable Drive property search:
 *   sid_th  — the threadId (hex, colon-free) used as the Drive query term
 *   sid_pre — the "threadId:msgId:index:" prefix used for in-memory matching
 *
 * Drive's `description contains` query is NOT a valid query field (v3 API),
 * but `properties has { key='K' and value='V' }` IS supported.
 *
 * @param {DriveFile} file - The Drive file to annotate
 * @param {string|null} sourceAttachmentId - Full attachment ID
 */
function setSourceProperties(file, sourceAttachmentId) {
  if (!sourceAttachmentId) return;
  const parts = sourceAttachmentId.split(":");
  if (parts.length < 3) return;
  file.setProperty("sid_th", parts[0]);
  file.setProperty("sid_pre", parts.slice(0, 3).join(":") + ":");
}

/**
 * Searches a Drive folder for a file saved from this exact source attachment,
 * using Drive custom properties (sid_th / sid_pre) set by setSourceProperties.
 *
 * Two-step process:
 *  1. Drive query by sid_th (pure hex threadId, no Lucene-special chars)
 *  2. In-memory filter by sid_pre (the threadId:msgId:index: prefix)
 *
 * NOTE: Drive API v3's `description contains` is NOT a valid query field and
 * raises `Invalid argument: q`. Custom properties ARE queryable.
 *
 * @param {string|null} sourceAttachmentId - Deterministic source attachment ID
 * @param {Folder} folder - Drive folder to search in
 * @returns {DriveFile|null} Matching file or null
 */
function findFileBySourceId(sourceAttachmentId, folder) {
  if (!sourceAttachmentId) return null;
  try {
    const parts = sourceAttachmentId.split(":");
    const threadId = parts[0]; // hex only — safe for Drive query
    const safePrefix = parts.slice(0, 3).join(":") + ":";
    // `properties has` is the correct query term for GAS File.setProperty().
    // Only the threadId (hex) is in the query; the colon-containing safePrefix
    // is compared in-memory to avoid Lucene colon-as-field-separator issues.
    const query = `'${folder.getId()}' in parents and properties has { key='sid_th' and value='${threadId}' }`;
    logWithUser(
      `findFileBySourceId: threadId=${threadId} safePrefix=${safePrefix}`,
      "DEBUG"
    );
    const results = DriveApp.searchFiles(query);
    while (results.hasNext()) {
      const file = results.next();
      const storedPrefix = file.getProperty("sid_pre") || "";
      if (storedPrefix === safePrefix) {
        logWithUser(
          `findFileBySourceId: matched ${file.getName()} (sid_pre=${storedPrefix})`,
          "DEBUG"
        );
        return file;
      }
    }
    logWithUser(`findFileBySourceId: no match for safePrefix=${safePrefix}`, "DEBUG");
    return null;
  } catch (e) {
    logWithUser(
      `findFileBySourceId: search failed: ${e.message} | rawId=${sourceAttachmentId}`,
      "WARNING"
    );
    return null;
  }
}

/**
 * Saves an attachment to a Drive folder with two-stage duplicate detection.
 *
 * Dedup strategy (in order of cost):
 * 1. Filename + size match in folder → duplicate, skip (cheap: one Drive folder scan)
 * 2. Drive property search by source attachment ID → duplicate, skip
 *    (only reached on name collision — uncommon)
 * 3. No duplicate found → save as new file
 *
 * Every saved file gets two custom Drive properties (sid_th, sid_pre) so
 * future runs can locate it even after a filename rename.
 *
 * @param {GmailAttachment} attachment - The email attachment
 * @param {GmailMessage} message - The email message containing the attachment
 * @param {Folder} domainFolder - The Google Drive folder for the domain
 * @param {Object} options - Optional settings (sourceAttachmentId)
 * @returns {Object} Result object: { success, duplicate, file } or { success: false, error }
 */
function saveAttachment(attachment, message, domainFolder, options = {}) {
  try {
    const attachmentName = attachment.getName();
    const attachmentBytes = attachment.getSize();
    const sourceAttachmentId = options.sourceAttachmentId || null;
    const emailDate = message.getDate();

    logWithUser(
      `Processing attachment: ${attachmentName} (${Math.round(attachmentBytes / 1024)}KB)`,
      "DEBUG"
    );
    logWithUser(`Email date: ${emailDate.toISOString()}`, "DEBUG");

    // --- Stage 1: filename + size match (fast path) ---
    const existingFiles = domainFolder.getFilesByName(attachmentName);
    let nameCollision = false;
    while (existingFiles.hasNext()) {
      nameCollision = true;
      const existingFile = existingFiles.next();

      if (existingFile.getSize() === attachmentBytes) {
        logWithUser(
          `Duplicate detected by name+size: ${attachmentName}`,
          "INFO"
        );
        return { success: true, duplicate: true, file: existingFile };
      }
    }

    if (nameCollision) {
      // Name collision (same name, different size): check if this exact
      // attachment was already saved under a renamed filename.
      const renamedFile = findFileBySourceId(sourceAttachmentId, domainFolder);
      if (renamedFile) {
        logWithUser(
          `Duplicate detected by sid_pre (renamed file): ${renamedFile.getName()}`,
          "INFO"
        );
        return { success: true, duplicate: true, file: renamedFile };
      }

      // Genuine new attachment with a name collision → rename and save.
      const newName = getUniqueFilename(attachmentName, domainFolder);
      logWithUser(`Name collision, saving as: ${newName}`, "INFO");
      const savedFile = domainFolder.createFile(
        attachment.copyBlob().setName(newName)
      );
      savedFile.setDescription(buildAttachmentMetadata(emailDate, sourceAttachmentId));
      setSourceProperties(savedFile, sourceAttachmentId);
      logWithUser(
        `Successfully saved: ${newName} in ${domainFolder.getName()}`,
        "INFO"
      );
      return { success: true, duplicate: false, file: savedFile };
    }

    // --- Stage 2: no file by that name — property search as safety net ---
    // Handles edge cases where the file exists under a different name.
    const fileBySourceId = findFileBySourceId(sourceAttachmentId, domainFolder);
    if (fileBySourceId) {
      logWithUser(
        `Duplicate detected by sid_pre (different name): ${fileBySourceId.getName()}`,
        "INFO"
      );
      return { success: true, duplicate: true, file: fileBySourceId };
    }

    // --- Stage 3: no duplicate found → save normally ---
    const savedFile = domainFolder.createFile(attachment);
    savedFile.setDescription(buildAttachmentMetadata(emailDate, sourceAttachmentId));
    setSourceProperties(savedFile, sourceAttachmentId);
    logWithUser(
      `Successfully saved: ${attachmentName} in ${domainFolder.getName()}`,
      "INFO"
    );
    return { success: true, duplicate: false, file: savedFile };
  } catch (error) {
    logWithUser(
      `Error saving attachment ${attachment.getName()}: ${error.message}`,
      "ERROR"
    );
    return { success: false, error: error.message };
  }
}

/**
 * Compatibility wrapper for the old saveAttachment signature
 * This ensures backward compatibility with existing code
 *
 * @param {GmailAttachment} attachment - The attachment to save
 * @param {DriveFolder} folder - The folder to save to
 * @param {Date} messageDate - The message date for timestamp
 * @param {Object} options - Optional settings (sourceAttachmentId)
 * @returns {DriveFile|null} The saved file or null
 */
function saveAttachmentLegacy(attachment, folder, messageDate, options = {}) {
  const mockMessage = {
    getDate: function () {
      return messageDate || new Date();
    },
  };
  const result = saveAttachment(attachment, mockMessage, folder, options);
  return result.success ? result.file : null;
}
