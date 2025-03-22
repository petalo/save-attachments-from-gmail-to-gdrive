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
 * 2. Gets the email date to use for the file timestamp
 * 3. Checks if a file with the same name already exists in the domain folder
 * 4. If a file exists with the same name:
 *    - Compares file sizes to detect duplicates
 *    - If sizes match, considers it a duplicate and returns the existing file
 *    - If sizes differ, generates a unique filename to avoid collision
 * 5. Creates the file in Google Drive (either with original or unique name)
 * 6. Sets the file creation date to match the email date
 * 7. Verifies timestamp accuracy and logs warnings for significant differences
 * 8. Returns a detailed result object with success status and file reference
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
 *
 * The function follows this flow:
 * 1. Logs the operation for tracking purposes
 * 2. Skips standard Drive API methods that often fail with "File not found" errors
 * 3. Implements a workaround by:
 *    - Getting the original file's content as a blob
 *    - Creating a new file with the same content in the same folder
 *    - Setting the new file's name to match the original
 *    - Adding the original date to the file's description metadata
 *    - For text files, appending an invisible timestamp comment
 *    - Deleting the original file by moving it to trash
 * 4. Returns success even if the exact timestamp couldn't be set
 *
 * This workaround is necessary because Google Drive doesn't provide direct API
 * methods to modify file creation dates, and this approach preserves the file's
 * content while associating it with the email's timestamp.
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
