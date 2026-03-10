/**
 * Attachment processing functions for Gmail Attachment Organizer
 */

/**
 * Builds a stable script-property key for an attachment source ID in a folder.
 *
 * @param {string} sourceAttachmentId - Deterministic source attachment ID
 * @param {string} folderId - Destination folder ID
 * @returns {string|null} Property key or null if inputs are missing
 */
function buildAttachmentSourceIndexKey(sourceAttachmentId, folderId) {
  if (!sourceAttachmentId || !folderId) return null;
  const raw = `${sourceAttachmentId}|${folderId}`;
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, raw);
  const encoded = Utilities.base64EncodeWebSafe(digest).replace(/=+$/, "");
  return `ATTACHMENT_SOURCE_${encoded}`;
}

/**
 * Builds description metadata to persist source linkage.
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
 * Returns an indexed file for a sourceAttachmentId/folder if it exists and is valid.
 *
 * @param {string|null} sourceAttachmentId - Deterministic source attachment ID
 * @param {Folder} folder - Destination folder
 * @returns {DriveFile|null} Existing file from index or null
 */
function resolveIndexedAttachment(sourceAttachmentId, folder) {
  if (!sourceAttachmentId) return null;

  const folderId = folder.getId();
  const key = buildAttachmentSourceIndexKey(sourceAttachmentId, folderId);
  if (!key) return null;

  const scriptProperties = PropertiesService.getScriptProperties();
  const fileId = scriptProperties.getProperty(key);
  if (!fileId) return null;

  try {
    const file = DriveApp.getFileById(fileId);
    const parents = file.getParents();
    while (parents.hasNext()) {
      if (parents.next().getId() === folderId) {
        return file;
      }
    }

    // Indexed file is no longer in folder, clear stale index.
    scriptProperties.deleteProperty(key);
    return null;
  } catch (e) {
    // Indexed file was deleted or inaccessible, clear stale index.
    logWithUser(`resolveIndexedAttachment: cleared stale index for key ${key}: ${e.message}`, "DEBUG");
    scriptProperties.deleteProperty(key);
    return null;
  }
}

/**
 * Registers sourceAttachmentId -> fileId mapping for a folder.
 *
 * @param {string|null} sourceAttachmentId - Deterministic source attachment ID
 * @param {Folder} folder - Destination folder
 * @param {DriveFile} file - Saved or existing file
 */
function registerAttachmentSource(sourceAttachmentId, folder, file) {
  if (!sourceAttachmentId || !file) return;

  const key = buildAttachmentSourceIndexKey(sourceAttachmentId, folder.getId());
  if (!key) return;

  PropertiesService.getScriptProperties().setProperty(key, file.getId());
}

/**
 * Saves an attachment to the appropriate folder based on the sender's domain
 *
 * @param {GmailAttachment} attachment - The email attachment
 * @param {GmailMessage} message - The email message containing the attachment
 * @param {Folder} domainFolder - The Google Drive folder for the domain
 * @param {Object} options - Optional settings (sourceAttachmentId)
 * @returns {Object} Result object with success status and saved file (if successful)
 *
 * The function follows this flow:
 * 1. Extracts attachment name and size for logging
 * 2. Checks if a file with the same name already exists in the domain folder
 * 3. If a file exists with the same name:
 *    - Compares file sizes to detect duplicates
 *    - If sizes match, considers it a duplicate and returns the existing file
 *    - If sizes differ, generates a unique filename to avoid collision
 * 4. Creates the file in Google Drive (either with original or unique name)
 * 5. Returns a detailed result object with success status and file reference
 *
 * This function handles duplicate detection and collision avoidance to ensure
 * no attachments are lost when processing emails.
 */
function saveAttachment(attachment, message, domainFolder, options = {}) {
  try {
    const attachmentName = attachment.getName();
    const attachmentSize = Math.round(attachment.getSize() / 1024);
    const sourceAttachmentId = options.sourceAttachmentId || null;

    // Log once at start, but don't repeat filter checks in logs
    logWithUser(
      `Processing attachment: ${attachmentName} (${attachmentSize}KB)`,
      "DEBUG"
    );

    // Get the date of the email for logging purposes
    const emailDate = message.getDate();
    logWithUser(`Email date: ${emailDate.toISOString()}`, "DEBUG");

    // Fast checkpoint lookup by deterministic source ID + folder
    const indexedFile = resolveIndexedAttachment(sourceAttachmentId, domainFolder);
    if (indexedFile) {
      logWithUser(
        `Attachment already indexed for source ID, skipping save: ${attachmentName}`,
        "INFO"
      );
      return { success: true, duplicate: true, file: indexedFile };
    }

    // Skip filter logging details here - we've already decided to save this file

    // Check if file already exists in the domain folder
    const existingFiles = domainFolder.getFilesByName(attachmentName);

    if (existingFiles.hasNext()) {
      // Check if it's exactly the same file (size-based check for simplicity)
      const existingFile = existingFiles.next();
      const existingFileSize = Math.round(existingFile.getSize() / 1024);

      if (existingFileSize === attachmentSize) {
        logWithUser(
          `File already exists with same size: ${attachmentName}`,
          "INFO"
        );
        registerAttachmentSource(sourceAttachmentId, domainFolder, existingFile);
        return { success: true, duplicate: true, file: existingFile };
      } else {
        // If sizes don't match, rename with timestamp to avoid collision
        const newName = getUniqueFilename(attachmentName, domainFolder);
        logWithUser(`Renaming to avoid collision: ${newName}`, "INFO");
        const savedFile = domainFolder.createFile(
          attachment.copyBlob().setName(newName)
        );

        // Add date info to the description for reference
        savedFile.setDescription(
          buildAttachmentMetadata(emailDate, sourceAttachmentId)
        );
        registerAttachmentSource(sourceAttachmentId, domainFolder, savedFile);

        logWithUser(
          `Successfully saved: ${newName} in ${domainFolder.getName()}`,
          "INFO"
        );
        return { success: true, duplicate: false, file: savedFile };
      }
    } else {
      // Save the file normally
      const savedFile = domainFolder.createFile(attachment);

      // Add date info to the description for reference
      savedFile.setDescription(buildAttachmentMetadata(emailDate, sourceAttachmentId));
      registerAttachmentSource(sourceAttachmentId, domainFolder, savedFile);

      logWithUser(
        `Successfully saved: ${attachmentName} in ${domainFolder.getName()}`,
        "INFO"
      );

      return { success: true, duplicate: false, file: savedFile };
    }
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
  // Create a mock message object with a getDate method
  const mockMessage = {
    getDate: function () {
      return messageDate || new Date();
    },
  };

  // Call the new version with the right parameters
  const result = saveAttachment(attachment, mockMessage, folder, options);

  // Return the file or null for compatibility
  return result.success ? result.file : null;
}
