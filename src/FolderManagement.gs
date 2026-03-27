/**
 * Folder management functions for Gmail Attachment Organizer
 */

/**
 * Get or create a folder for the sender's domain
 *
 * Why it uses locks:
 * - Prevents race conditions when multiple executions attempt to create
 *   the same folder simultaneously
 * - Without locks, multiple executions might check that a folder doesn't exist
 *   and then try to create it, resulting in duplicate folders
 *
 * @param {string} sender - The sender's email address
 * @param {DriveFolder} mainFolder - The main folder to create the domain folder in
 * @returns {DriveFolder} The domain folder
 *
 * The function follows this flow:
 * 1. Extracts the domain from the sender's email address using regex
 * 2. Acquires a lock to prevent race conditions during folder creation
 * 3. Checks if a folder for this domain already exists
 * 4. If the folder exists, returns it immediately
 * 5. If not, performs a double-check to handle edge cases where another
 *    execution might have created the folder between checks
 * 6. If still not found, creates a new folder for the domain
 * 7. If any errors occur, falls back to using an "unknown" folder
 * 8. As a last resort, returns the main folder if all else fails
 *
 * The function includes robust error handling and fallback mechanisms to ensure
 * attachments are always saved somewhere, even if the ideal domain folder
 * cannot be created or accessed.
 */
function getDomainFolder(sender, mainFolder) {
  try {
    // Extract the domain from the sender's email address
    const domain = extractDomain(sender);

    // Use a lock to prevent race conditions when creating folders
    const lock = LockService.getScriptLock();
    try {
      if (!lock.tryLock(10000)) {
        throw new Error(
          `Could not acquire lock for domain folder "${domain}" — skipping to avoid duplicates`
        );
      }

      // Check if the folder exists
      const folders = withRetry(
        () => mainFolder.getFoldersByName(domain),
        "getting domain folder"
      );

      if (folders.hasNext()) {
        const folder = folders.next();
        logWithUser(`Using existing domain folder: ${domain}`);
        return folder;
      }

      // Brief pause to mitigate Drive eventual consistency before creating
      Utilities.sleep(1000);

      const foldersRecheck = withRetry(
        () => mainFolder.getFoldersByName(domain),
        "rechecking domain folder after delay"
      );

      if (foldersRecheck.hasNext()) {
        const folder = foldersRecheck.next();
        logWithUser(
          `Using existing domain folder (after recheck): ${domain}`
        );
        return folder;
      }

      const newFolder = withRetry(
        () => mainFolder.createFolder(domain),
        "creating domain folder"
      );
      logWithUser(`Created new domain folder: ${domain}`);
      return newFolder;
    } finally {
      // Always release the lock
      if (lock.hasLock()) {
        lock.releaseLock();
      }
    }
  } catch (error) {
    logWithUser(
      `Error getting domain folder for ${sender}: ${error.message}`,
      "ERROR"
    );

    // Use a similar approach for the unknown folder
    try {
      const lock = LockService.getScriptLock();
      try {
        if (!lock.tryLock(10000)) {
          throw new Error(
            `Could not acquire lock for unknown folder — skipping to avoid duplicates`
          );
        }

        const unknownFolders = withRetry(
          () => mainFolder.getFoldersByName("unknown"),
          "getting unknown folder"
        );

        if (unknownFolders.hasNext()) {
          const folder = unknownFolders.next();
          logWithUser(`Using fallback 'unknown' folder for ${sender}`);
          return folder;
        }

        Utilities.sleep(1000);

        const unknownRecheck = withRetry(
          () => mainFolder.getFoldersByName("unknown"),
          "rechecking unknown folder after delay"
        );

        if (unknownRecheck.hasNext()) {
          const folder = unknownRecheck.next();
          logWithUser(
            `Using fallback 'unknown' folder (after recheck) for ${sender}`
          );
          return folder;
        }

        const newFolder = withRetry(
          () => mainFolder.createFolder("unknown"),
          "creating unknown folder"
        );
        logWithUser(`Created fallback 'unknown' folder for ${sender}`);
        return newFolder;
      } finally {
        // Always release the lock
        if (lock.hasLock()) {
          lock.releaseLock();
        }
      }
    } catch (fallbackError) {
      logWithUser(
        `Failed to create fallback folder: ${fallbackError.message}. Using main folder as fallback.`,
        "ERROR"
      );
      // Ultimate fallback: return the main folder
      return mainFolder;
    }
  }
}
