/**
 * Utility functions for Gmail Attachment Organizer
 */

/**
 * Helper function to create consistent log entries with user information
 *
 * @param {string} message - The message to log
 * @param {string} level - The log level (INFO, WARNING, ERROR)
 */
function logWithUser(message, level = "INFO") {
  const userEmail = Session.getEffectiveUser().getEmail();
  const timestamp = new Date().toISOString();
  const logMessage = `[${timestamp}] [${level}] [${userEmail}] ${message}`;
  Logger.log(logMessage);
}

/**
 * Helper function to get user-specific settings
 *
 * @returns {Object} The user settings
 */
function getUserSettings() {
  const userProperties = PropertiesService.getUserProperties();
  const defaultSettings = {
    maxFileSize: CONFIG.maxFileSize,
    batchSize: CONFIG.batchSize,
    skipDomains: [], // Domains to skip processing
    preferredFolder: CONFIG.mainFolderId,
  };

  try {
    const savedSettings = userProperties.getProperty("userSettings");
    return savedSettings ? JSON.parse(savedSettings) : defaultSettings;
  } catch (e) {
    logWithUser(`Error reading user settings: ${e.message}`, "WARNING");
    return defaultSettings;
  }
}

/**
 * Helper function to save user-specific settings
 *
 * @param {Object} settings - The settings to save
 */
function saveUserSettings(settings) {
  const userProperties = PropertiesService.getUserProperties();
  try {
    userProperties.setProperty("userSettings", JSON.stringify(settings));
    logWithUser("User settings saved successfully");
  } catch (e) {
    logWithUser(`Error saving user settings: ${e.message}`, "ERROR");
  }
}

/**
 * Helper function to implement retry logic with exponential backoff
 *
 * Why it's necessary:
 * - Google Apps Script may encounter transient errors when interacting with Gmail and Drive
 * - These errors are often automatically resolved with a retry
 * - Exponential backoff prevents overloading the services
 * - Improves script robustness by handling temporary network conditions or quota limits
 *
 * @param {Function} operation - The operation to retry
 * @param {string} context - A description of the operation for logging
 * @returns {*} The result of the operation
 */
function withRetry(operation, context = "") {
  let lastError;
  let delay = CONFIG.retryDelay;

  for (let attempt = 1; attempt <= CONFIG.maxRetries; attempt++) {
    try {
      return operation();
    } catch (error) {
      lastError = error;
      logWithUser(
        `Attempt ${attempt}/${CONFIG.maxRetries} failed for ${context}: ${error.message}`,
        "WARNING"
      );

      if (attempt < CONFIG.maxRetries) {
        Utilities.sleep(delay);
        delay = Math.min(delay * 2, CONFIG.maxRetryDelay); // Exponential backoff
      }
    }
  }

  throw lastError;
}

/**
 * Prevents multiple simultaneous executions of the script by different users
 * Uses both LockService and Script Properties for robust locking
 *
 * @param {string} userEmail - The email of the user acquiring the lock
 * @returns {boolean} True if the lock was acquired, false otherwise
 */
function acquireExecutionLock(userEmail) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const lockKey = "EXECUTION_LOCK";

  // Try to acquire LockService lock first
  const lock = LockService.getScriptLock();
  try {
    if (lock.tryLock(30000)) {
      // Try to lock for 30 seconds
      // Also set script property lock as backup
      const lockData = {
        user: userEmail,
        timestamp: new Date().getTime(),
      };
      scriptProperties.setProperty(lockKey, JSON.stringify(lockData));
      logWithUser(`Lock acquired by ${userEmail} using LockService`);
      return true;
    }
  } catch (e) {
    logWithUser(`LockService failed: ${e.message}`, "WARNING");
  }

  // Fallback to script properties lock
  const lockInfo = scriptProperties.getProperty(lockKey);
  const now = new Date().getTime();

  if (lockInfo) {
    try {
      const lockData = JSON.parse(lockInfo);
      const lockTime = lockData.timestamp;
      const lockUser = lockData.user;

      // Check if the lock has expired (allow for CONFIG.executionLockTime minutes)
      if (now - lockTime < CONFIG.executionLockTime * 60 * 1000) {
        logWithUser(
          `Script is locked by ${lockUser} until ${new Date(
            lockTime + CONFIG.executionLockTime * 60 * 1000
          )}`,
          "WARNING"
        );
        return false;
      }

      // Lock has expired, we can take over
      logWithUser(`Found expired lock from ${lockUser}, acquiring new lock`);
    } catch (e) {
      // Invalid lock data, we can acquire the lock
      logWithUser("Found invalid lock data, acquiring new lock");
    }
  }

  // Acquire the lock
  const lockData = {
    user: userEmail,
    timestamp: now,
  };

  scriptProperties.setProperty(lockKey, JSON.stringify(lockData));
  logWithUser(`Lock acquired by ${userEmail} using Script Properties`);
  return true;
}

/**
 * Releases the execution lock
 *
 * @param {string} userEmail - The email of the user releasing the lock
 */
function releaseExecutionLock(userEmail) {
  const scriptProperties = PropertiesService.getScriptProperties();
  const lockKey = "EXECUTION_LOCK";

  // Try to release LockService lock first
  const lock = LockService.getScriptLock();
  try {
    if (lock.hasLock()) {
      lock.releaseLock();
      logWithUser("Released LockService lock");
    }
  } catch (e) {
    logWithUser(`Error releasing LockService lock: ${e.message}`, "WARNING");
  }

  // Release script properties lock
  const lockInfo = scriptProperties.getProperty(lockKey);
  if (lockInfo) {
    try {
      const lockData = JSON.parse(lockInfo);
      const lockUser = lockData.user;
      const currentUser = Session.getEffectiveUser().getEmail();

      if (lockUser === currentUser) {
        scriptProperties.deleteProperty(lockKey);
        logWithUser(`Lock released by ${currentUser}`);
      } else {
        logWithUser(`Cannot release lock owned by ${lockUser}`, "WARNING");
      }
    } catch (e) {
      // Invalid lock data, delete it
      scriptProperties.deleteProperty(lockKey);
      logWithUser("Deleted invalid lock data");
    }
  }
}

/**
 * Generates a unique filename to avoid collisions in the same folder
 *
 * @param {string} originalFilename - The original file name
 * @param {Folder} folder - The Google Drive folder
 * @returns {string} A unique filename that doesn't exist in the folder
 */
function getUniqueFilename(originalFilename, folder) {
  // Extract base name and extension
  const lastDotIndex = originalFilename.lastIndexOf(".");
  const baseName =
    lastDotIndex > 0
      ? originalFilename.substring(0, lastDotIndex)
      : originalFilename;
  const extension =
    lastDotIndex > 0 ? originalFilename.substring(lastDotIndex) : "";

  // First try adding a timestamp
  const timestamp = new Date().getTime();
  let newName = `${baseName}_${timestamp}${extension}`;

  // Check if this name exists
  if (!folder.getFilesByName(newName).hasNext()) {
    return newName;
  }

  // If timestamp wasn't enough, add a random string too
  const randomString = Utilities.getUuid().substring(0, 8);
  return `${baseName}_${timestamp}_${randomString}${extension}`;
}
