/**
 * Attachment processing functions for Gmail Attachment Organizer
 */

/**
 * Builds description metadata to persist source linkage in the Drive file.
 * The source_attachment_id field is the primary dedup key across runs,
 * including renamed files.
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
 * Searches a Drive folder for a file whose description contains the given
 * sourceAttachmentId. Used as a fallback dedup check for files that were
 * renamed due to name collisions in a previous run.
 *
 * @param {string|null} sourceAttachmentId - Deterministic source attachment ID
 * @param {Folder} folder - Drive folder to search in
 * @returns {DriveFile|null} Matching file or null
 */
function findFileBySourceId(sourceAttachmentId, folder) {
  if (!sourceAttachmentId) return null;
  try {
    const results = DriveApp.searchFiles(
      `'${folder.getId()}' in parents and description contains 'source_attachment_id=${sourceAttachmentId}'`
    );
    return results.hasNext() ? results.next() : null;
  } catch (e) {
    logWithUser(
      `findFileBySourceId: Drive search failed: ${e.message}`,
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
 * 2. Drive description search by source_attachment_id → duplicate, skip
 *    (only reached on name collision or missing file — uncommon)
 * 3. No duplicate found → save as new file
 *
 * The source_attachment_id is always persisted in the file description so
 * that future runs can detect duplicates even if the filename changes.
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
    const attachmentSize = Math.round(attachment.getSize() / 1024);
    const sourceAttachmentId = options.sourceAttachmentId || null;
    const emailDate = message.getDate();

    logWithUser(
      `Processing attachment: ${attachmentName} (${attachmentSize}KB)`,
      "DEBUG"
    );
    logWithUser(`Email date: ${emailDate.toISOString()}`, "DEBUG");

    // --- Stage 1: filename + size match (fast path) ---
    const existingFiles = domainFolder.getFilesByName(attachmentName);
    if (existingFiles.hasNext()) {
      const existingFile = existingFiles.next();
      const existingFileSize = Math.round(existingFile.getSize() / 1024);

      if (existingFileSize === attachmentSize) {
        logWithUser(
          `Duplicate detected by name+size: ${attachmentName}`,
          "INFO"
        );
        return { success: true, duplicate: true, file: existingFile };
      }

      // Name collision (same name, different size): check if this exact
      // attachment was already saved under a renamed filename.
      const renamedFile = findFileBySourceId(sourceAttachmentId, domainFolder);
      if (renamedFile) {
        logWithUser(
          `Duplicate detected by source_attachment_id (renamed file): ${renamedFile.getName()}`,
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
      savedFile.setDescription(
        buildAttachmentMetadata(emailDate, sourceAttachmentId)
      );
      logWithUser(
        `Successfully saved: ${newName} in ${domainFolder.getName()}`,
        "INFO"
      );
      return { success: true, duplicate: false, file: savedFile };
    }

    // --- Stage 2: no file by that name — check description as safety net ---
    // Handles edge cases where the file exists under a different name.
    const fileBySourceId = findFileBySourceId(sourceAttachmentId, domainFolder);
    if (fileBySourceId) {
      logWithUser(
        `Duplicate detected by source_attachment_id (different name): ${fileBySourceId.getName()}`,
        "INFO"
      );
      return { success: true, duplicate: true, file: fileBySourceId };
    }

    // --- Stage 3: no duplicate found → save normally ---
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
