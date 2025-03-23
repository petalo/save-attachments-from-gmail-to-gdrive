/**
 * Attachment processing functions for Gmail Attachment Organizer
 */

/**
 * Saves an attachment to the appropriate folder based on the sender's domain
 *
 * @param {GmailAttachment} attachment - The email attachment
 * @param {GmailMessage} message - The email message containing the attachment
 * @param {Folder} domainFolder - The Google Drive folder for the domain
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
function saveAttachment(attachment, message, domainFolder) {
  try {
    const attachmentName = attachment.getName();
    const attachmentSize = Math.round(attachment.getSize() / 1024);

    // Log once at start, but don't repeat filter checks in logs
    logWithUser(
      `Processing attachment: ${attachmentName} (${attachmentSize}KB)`,
      "INFO"
    );

    // Get the date of the email for logging purposes
    const emailDate = message.getDate();
    logWithUser(`Email date: ${emailDate.toISOString()}`, "INFO");

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
        return { success: true, duplicate: true, file: existingFile };
      } else {
        // If sizes don't match, rename with timestamp to avoid collision
        const newName = getUniqueFilename(attachmentName, domainFolder);
        logWithUser(`Renaming to avoid collision: ${newName}`, "INFO");
        const savedFile = domainFolder.createFile(
          attachment.copyBlob().setName(newName)
        );

        // Add date info to the description for reference
        savedFile.setDescription(`email_date=${emailDate.toISOString()}`);

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
      savedFile.setDescription(`email_date=${emailDate.toISOString()}`);

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
 * @returns {DriveFile|null} The saved file or null
 */
function saveAttachmentLegacy(attachment, folder, messageDate) {
  // Create a mock message object with a getDate method
  const mockMessage = {
    getDate: function () {
      return messageDate || new Date();
    },
  };

  // Call the new version with the right parameters
  const result = saveAttachment(attachment, mockMessage, folder);

  // Return the file or null for compatibility
  return result.success ? result.file : null;
}
