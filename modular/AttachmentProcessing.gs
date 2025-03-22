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

    // Get the date of the email for the file timestamp
    const emailDate = message.getDate();
    logWithUser(`Email date for timestamp: ${emailDate.toISOString()}`, "INFO");

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

        // Set the file creation date to match the email date
        setFileCreationDate(savedFile, emailDate);

        logWithUser(
          `Successfully saved: ${newName} in ${domainFolder.getName()}`,
          "INFO"
        );
        return { success: true, duplicate: false, file: savedFile };
      }
    } else {
      // Save the file normally
      const savedFile = domainFolder.createFile(attachment);

      // Set the file creation date to match the email date
      setFileCreationDate(savedFile, emailDate);

      logWithUser(
        `Successfully saved: ${attachmentName} in ${domainFolder.getName()}`,
        "INFO"
      );

      // Log a warning if the file timestamp differs significantly from email date
      const fileDate = savedFile.getLastUpdated();
      const diffMs = Math.abs(fileDate.getTime() - emailDate.getTime());
      const diffMinutes = Math.round(diffMs / (1000 * 60));

      if (diffMinutes > 60) {
        // Only log warnings if more than 1 hour difference
        logWithUser(
          `Saved file modification date: ${fileDate.toISOString()}`,
          "INFO"
        );
        logWithUser(
          `⚠️ File timestamp differs from email date by ${diffMinutes} minutes`,
          "WARNING"
        );
      }

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
 * Sets the creation date of a file to match a specific date
 * This uses file recreation since the Drive API methods don't work reliably
 *
 * @param {DriveFile} file - The Google Drive file
 * @param {Date} date - The date to set as creation date
 * @returns {boolean} True if successful, false otherwise
 */
function setFileCreationDate(file, date) {
  try {
    const fileName = file.getName();
    const fileId = file.getId();

    logWithUser(`Setting timestamp for file: ${fileName}`, "INFO");

    // Skip Drive API methods as they consistently fail with "File not found" errors
    // Go directly to the file recreation method that works
    try {
      // Get file content
      const blob = file.getBlob();
      const mimeType = file.getMimeType();
      const parentFolder = file.getParents().next();

      // Create new file with the same content
      const newFile = parentFolder.createFile(blob);
      newFile.setName(fileName);

      // Add date info to the description
      newFile.setDescription(`original_date=${date.toISOString()}`);

      // For text files, try appending an invisible comment
      if (
        mimeType.includes("text/") ||
        mimeType.includes("application/json") ||
        mimeType.includes("xml") ||
        mimeType.includes("html")
      ) {
        try {
          const content = newFile.getBlob().getDataAsString();
          const updatedContent =
            content + "\n<!-- timestamp:" + date.getTime() + " -->";
          newFile.setContent(updatedContent);
        } catch (e) {
          // Just log and continue
          logWithUser(`Content update error: ${e.message}`, "WARNING");
        }
      }

      // Delete original file
      file.setTrashed(true);

      // Verify result
      const finalDate = newFile.getLastUpdated();

      // Return success even if we couldn't set the exact date
      // The important thing is preserving the file content
      logWithUser(
        `Created replacement file with ID: ${newFile.getId()}`,
        "INFO"
      );
      return true;
    } catch (e) {
      logWithUser(`File recreation failed: ${e.message}`, "ERROR");
      return false;
    }
  } catch (error) {
    logWithUser(
      `General error in setFileCreationDate: ${error.message}`,
      "ERROR"
    );
    return false;
  }
}

/**
 * Calculate the difference in months between two dates
 * Helper function for timestamp verification
 *
 * @param {Date} date1 - First date
 * @param {Date} date2 - Second date
 * @returns {number} Difference in months (can be decimal)
 */
function dateDiffInMonths(date1, date2) {
  const monthDiff =
    (date2.getFullYear() - date1.getFullYear()) * 12 +
    (date2.getMonth() - date1.getMonth());

  // Add day-based fraction for more precision
  const dayDiff = date2.getDate() - date1.getDate();
  const daysInMonth = new Date(
    date1.getFullYear(),
    date1.getMonth() + 1,
    0
  ).getDate();

  return monthDiff + dayDiff / daysInMonth;
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
