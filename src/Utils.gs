/**
 * Utility functions for Gmail Attachment Organizer
 */

/**
 * Helper function to create consistent log entries with user information
 *
 * This function creates standardized log entries that include timestamp, log level,
 * and the current user's email address.
 *
 * @param {string} message - The message content to log
 * @param {string} level - The log level (INFO, WARNING, ERROR), defaults to INFO if not specified
 * @returns {void} This function doesn't return a value, but writes to the Apps Script log
 *
 * The function follows this flow:
 * 1. Get the current user's email address using Session.getEffectiveUser()
 * 2. Generate a timestamp in ISO format using new Date().toISOString()
 * 3. Format the log message with timestamp, level, user email, and the provided message
 * 4. Write the formatted message to the Apps Script log using Logger.log()
 */
function logWithUser(message, level = "INFO") {
  // Handle undefined or null message
  if (message === undefined || message === null) {
    message = "[No message provided]";
  }

  const userEmail = Session.getEffectiveUser().getEmail();
  const timestamp = new Date().toISOString();
  const logMessage = `[${timestamp}] [${level}] [${userEmail}] ${message}`;
  Logger.log(logMessage);
}

/**
 * Helper function to get user-specific settings from the user's properties store
 *
 * This function retrieves user-specific settings or returns defaults if none are found.
 * The flow is:
 * 1. Access the user's properties store
 * 2. Define default settings based on global CONFIG
 * 3. Try to retrieve saved settings
 * 4. Parse and return saved settings if they exist, otherwise return defaults
 *
 * @returns {Object} The user settings object containing properties like maxFileSize,
 *                   batchSize, skipDomains, and preferredFolder with either saved values
 *                   or defaults from the CONFIG object
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
 * Helper function to save user-specific settings to the user's properties store
 *
 * This function persists user settings by serializing them to JSON and storing them
 * in the user's properties store. The flow is:
 * 1. Access the user's properties store
 * 2. Convert the settings object to a JSON string
 * 3. Save the JSON string to the "userSettings" property
 * 4. Log success or failure
 *
 * @param {Object} settings - The settings object to save, containing properties like
 *                           maxFileSize, batchSize, skipDomains, and preferredFolder
 * @returns {void} This function doesn't return a value, but logs the result of the operation
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
 * This function attempts to execute an operation multiple times with increasing delays
 * between attempts if failures occur. The flow is:
 * 1. Try to execute the provided operation function
 * 2. If successful, return the result immediately
 * 3. If it fails, log the error and wait before retrying
 * 4. For each retry, double the delay time (exponential backoff)
 * 5. After all retries are exhausted, throw the last error encountered
 *
 * Why it's necessary:
 * - Google Apps Script may encounter transient errors when interacting with Gmail and Drive
 * - These errors are often automatically resolved with a retry
 * - Exponential backoff prevents overloading the services
 * - Improves script robustness by handling temporary network conditions or quota limits
 *
 * @param {Function} operation - The operation function to retry
 * @param {string} context - A description of the operation for logging purposes
 * @returns {*} The result of the operation if successful
 * @throws {Error} The last error encountered if all retries fail
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
 * @param {string} userEmail - The email of the user acquiring the lock. If not provided, uses current user.
 * @returns {boolean} True if the lock was acquired, false otherwise
 *
 * The function follows this flow:
 * 1. First attempts to acquire a lock using LockService with a 30-second timeout
 * 2. If successful, also creates a backup lock in Script Properties with user and timestamp
 * 3. If LockService fails, checks for an existing lock in Script Properties
 * 4. If a lock exists, checks if it has expired (based on CONFIG.executionLockTime)
 * 5. If no lock exists or it has expired, acquires a new lock in Script Properties
 *
 * This dual locking mechanism provides redundancy in case either locking system fails.
 */
function acquireExecutionLock(userEmail) {
  // If userEmail is not provided, use the current user's email
  if (!userEmail) {
    userEmail = Session.getEffectiveUser().getEmail();
    logWithUser(
      `No user email provided for lock, using current user: ${userEmail}`,
      "INFO"
    );
  }
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
 * Releases the execution lock previously acquired by acquireExecutionLock
 *
 * This function attempts to release both the LockService lock and the Script Properties lock.
 * The flow is:
 * 1. Try to release the LockService lock if the current execution has it
 * 2. Check if there's a lock record in Script Properties
 * 3. If the lock belongs to the current user, delete the lock property
 * 4. If the lock belongs to another user, log a warning but don't delete it
 * 5. If the lock data is invalid, delete it as a cleanup measure
 *
 * @param {string} userEmail - The email of the user releasing the lock. If not provided, uses current user.
 * @returns {void} This function doesn't return a value, but logs the result of the operation
 */
function releaseExecutionLock(userEmail) {
  // If userEmail is not provided, use the current user's email
  if (!userEmail) {
    userEmail = Session.getEffectiveUser().getEmail();
    logWithUser(
      `No user email provided for lock release, using current user: ${userEmail}`,
      "INFO"
    );
  }
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

/**
 * Extracts the domain from an email address
 *
 * @param {string} email - The email address to extract the domain from
 * @returns {string} The domain part of the email address, or "unknown" if not found
 */
function extractDomain(email) {
  const domainMatch = email.match(/@([\w.-]+)/);
  return domainMatch ? domainMatch[1] : "unknown";
}

/**
 * Test function to verify if the folder ID is valid
 * This function can be run directly from the Apps Script editor
 * to check if the configured folder ID is correct
 */
function testFolderId() {
  try {
    const folder = DriveApp.getFolderById(CONFIG.mainFolderId);
    const folderName = folder.getName();
    Logger.log(
      `Successfully found folder: ${folderName} with ID: ${CONFIG.mainFolderId}`
    );
    return {
      success: true,
      folderName: folderName,
      folderId: CONFIG.mainFolderId,
    };
  } catch (e) {
    Logger.log(
      `Error accessing folder with ID ${CONFIG.mainFolderId}: ${e.message}`
    );
    return { success: false, error: e.message, folderId: CONFIG.mainFolderId };
  }
}
