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
      lock.tryLock(10000); // Wait up to 10 seconds for the lock

      // First check if the folder exists
      const folders = withRetry(
        () => mainFolder.getFoldersByName(domain),
        "getting domain folder"
      );

      if (folders.hasNext()) {
        const folder = folders.next();
        logWithUser(`Using existing domain folder: ${domain}`);
        return folder;
      } else {
        // Double-check that the folder still doesn't exist
        // This helps in cases where another execution created it just now
        const doubleCheckFolders = withRetry(
          () => mainFolder.getFoldersByName(domain),
          "double-checking domain folder"
        );

        if (doubleCheckFolders.hasNext()) {
          const folder = doubleCheckFolders.next();
          logWithUser(
            `Using existing domain folder (after double-check): ${domain}`
          );
          return folder;
        }

        // If we're still here, we can safely create the folder
        const newFolder = withRetry(
          () => mainFolder.createFolder(domain),
          "creating domain folder"
        );
        logWithUser(`Created new domain folder: ${domain}`);
        return newFolder;
      }
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
        lock.tryLock(10000);

        // Check for unknown folder
        const unknownFolders = withRetry(
          () => mainFolder.getFoldersByName("unknown"),
          "getting unknown folder"
        );

        if (unknownFolders.hasNext()) {
          const folder = unknownFolders.next();
          logWithUser(`Using fallback 'unknown' folder for ${sender}`);
          return folder;
        } else {
          // Double-check that the unknown folder still doesn't exist
          const doubleCheckUnknown = withRetry(
            () => mainFolder.getFoldersByName("unknown"),
            "double-checking unknown folder"
          );

          if (doubleCheckUnknown.hasNext()) {
            const folder = doubleCheckUnknown.next();
            logWithUser(
              `Using fallback 'unknown' folder (after double-check) for ${sender}`
            );
            return folder;
          }

          // Create the unknown folder
          const newFolder = withRetry(
            () => mainFolder.createFolder("unknown"),
            "creating unknown folder"
          );
          logWithUser(`Created fallback 'unknown' folder for ${sender}`);
          return newFolder;
        }
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
