/**
 * Attachment processing functions for Gmail Attachment Organizer
 */

/**
 * Builds description metadata to persist source linkage in the Drive file.
 * Stored for human readability only — not used for dedup search because
 * neither `description contains` nor `properties has` work as Drive query
 * fields on Shared Drive files via DriveApp.
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
 * Builds a deterministic renamed filename for name-collision cases.
 *
 * Appends a stable suffix derived from the first 8 hex chars of threadId,
 * attachment index, and byte size. Deliberately omits messageId so that
 * identical content appearing in multiple messages of the same thread
 * (e.g. inline signature images) always maps to the same filename and is
 * detected as a duplicate by getFilesByName on subsequent runs.
 *
 * Example: "image.png" (101569 bytes, index 0, thread 19c4c18a…)
 *          → "image__19c4c18a_0_101569.png"
 *
 * Fallback (no sourceAttachmentId): timestamp suffix (previous behaviour).
 *
 * @param {string} attachmentName - Original attachment filename
 * @param {string|null} sourceAttachmentId - Full attachment ID (threadId:msgId:index:...)
 * @param {number} attachmentBytes - Exact byte size of the attachment
 * @returns {string} Renamed filename
 */
function buildStableRename(attachmentName, sourceAttachmentId, attachmentBytes) {
  let suffix;
  if (sourceAttachmentId) {
    const parts = sourceAttachmentId.split(":");
    if (parts.length >= 3) {
      // threadId (8 hex) + attachmentIndex + size — unique per (thread, slot, content)
      suffix = `${parts[0].slice(0, 8)}_${parts[2]}_${attachmentBytes}`;
    }
  }
  if (!suffix) {
    suffix = Date.now().toString();
  }
  const dotIdx = attachmentName.lastIndexOf(".");
  return dotIdx >= 0
    ? `${attachmentName.slice(0, dotIdx)}__${suffix}${attachmentName.slice(dotIdx)}`
    : `${attachmentName}__${suffix}`;
}

/**
 * Saves an attachment to a Drive folder with two-stage duplicate detection.
 *
 * Dedup strategy:
 * 1. Filename + size match in folder → duplicate, skip  (cheap: exact name lookup)
 * 2a. Name collision (same name, different size): look up stable renamed file
 *     by its deterministic name → duplicate, skip if found; save under stable
 *     name otherwise
 * 2b. No name collision: save normally
 *
 * NOTE on Drive query limitations: DriveApp.searchFiles() on Shared Drive
 * files does not support `description contains` or `properties has` query
 * terms — both raise `Invalid argument: q`. DriveFile.setProperty() is
 * also unavailable on Shared Drive files. Dedup therefore relies entirely
 * on getFilesByName() (exact lookup, always works).
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

    // --- Stage 1: exact filename + size match (fast path) ---
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
      // Same name, different size: a different attachment already holds that
      // filename. Check whether THIS attachment was already saved under its
      // stable renamed filename (deterministic from sourceAttachmentId).
      const stableName = buildStableRename(attachmentName, sourceAttachmentId, attachmentBytes);
      const stableFiles = domainFolder.getFilesByName(stableName);
      if (stableFiles.hasNext()) {
        const existingRenamed = stableFiles.next();
        logWithUser(
          `Duplicate detected by stable rename: ${stableName}`,
          "INFO"
        );
        return { success: true, duplicate: true, file: existingRenamed };
      }

      // Genuine new attachment → save under stable renamed name.
      logWithUser(`Name collision, saving as: ${stableName}`, "INFO");
      const savedFile = domainFolder.createFile(
        attachment.copyBlob().setName(stableName)
      );
      savedFile.setDescription(
        buildAttachmentMetadata(emailDate, sourceAttachmentId)
      );
      logWithUser(
        `Successfully saved: ${stableName} in ${domainFolder.getName()}`,
        "INFO"
      );
      return { success: true, duplicate: false, file: savedFile };
    }

    // --- No name collision → save normally ---
    const savedFile = domainFolder.createFile(attachment);
    savedFile.setDescription(
      buildAttachmentMetadata(emailDate, sourceAttachmentId)
    );
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
