/**
 * Gmail Attachment Organizer
 *
 * This script automatically organizes Gmail attachments in Google Drive
 * by the sender's email domain. It processes unread emails with attachments,
 * extracts the attachments, and saves them to Google Drive folders organized
 * by the sender's domain.
 *
 * Key features:
 * - Automatic organization by sender domain
 * - Configurable filters for file types and sizes
 * - Multiple invoice detection methods (AI-powered and email-based)
 * - Privacy-focused metadata analysis for invoice detection
 * - Scheduled processing via time-based triggers
 * - Multi-user support with permission management
 * - Robust error handling with retry logic
 * - Duplicate file detection
 * - Timestamp preservation from original emails
 */


//=============================================================================
// CONFIG - CONFIGURATION
//=============================================================================

// Configuration constants
const CONFIG = {
  //=============================================================================
  // CORE SETTINGS - Essential configuration for basic functionality
  //=============================================================================

  // Google Drive folder where attachments will be saved
  // This is the main parent folder that will contain domain subfolders
  mainFolderId: "__FOLDER_ID__", // Replace with your Google Drive shared folder's ID

  // Gmail label applied to threads after processing
  // This prevents the same emails from being processed multiple times
  processedLabelName: "GDrive_Processed",
  processingLabelName: "GDrive_Processing",
  errorLabelName: "GDrive_Error",
  permanentErrorLabelName: "GDrive_Error_Permanent",
  tooLargeLabelName: "GDrive_TooLarge",

  // Execution settings
  // RECOMMENDED VALUES:
  // - triggerIntervalMinutes: 10-30 for normal use, 5 for high-volume environments
  // - batchSize: 5-10 for most environments, lower for complex processing
  // - executionLockTime: Should be at least 2x the expected execution time
  // INTERDEPENDENCY: batchSize directly affects execution time - higher values may hit the 6-minute limit
  triggerIntervalMinutes: 10, // How often the script runs (in minutes)
  batchSize: 20, // Number of threads to process per execution (prevents hitting 6-minute limit)
  executionLockTime: 10, // Maximum time in minutes to wait for lock release
  executionSoftLimitMs: 270000, // Stop early before hard timeout (4.5 minutes)
  processingStateTtlMinutes: 180, // Stale Processing state threshold for safe recovery
  staleRecoveryBatchSize: 10, // Max stale Processing threads to recover per execution
  maxThreadFailureRetries: 3, // Escalate repeated failures to permanent state
  threadFailureStateTtlDays: 30, // Auto-expire historical failure state
  executionModel: "effective_user_only", // Supported model: one execution processes only effective user mailbox

  //=============================================================================
  // FILTERING SETTINGS - Control which emails and attachments are processed
  //=============================================================================

  // Domain filtering
  // Emails from these domains will be completely skipped during processing
  skipDomains: ["example.com", "noreply.com"],

  // Attachment filtering
  skipFileTypes: [".ics", ".ical", ".pkpass", ".vcf", ".vcard"], // File types to skip (e.g., calendar invitations)
  maxFileSize: 25 * 1024 * 1024, // Maximum attachment size (25MB)

  // Small image filtering (helps avoid saving email signatures and tiny embedded images)
  // INTERDEPENDENCY: skipSmallImages only works when smallImageExtensions and smallImageMaxSize are properly set
  // RECOMMENDED VALUES:
  // - 20KB is good for filtering most email signatures and tiny embedded images
  // - Increase to 50KB to be more aggressive in filtering out small images
  // - Decrease to 10KB to only filter the smallest images
  skipSmallImages: true, // Whether to skip small images
  smallImageExtensions: [".jpg", ".jpeg", ".png", ".gif", ".bmp"], // Image types to check for size
  smallImageMaxSize: 20 * 1024, // Size threshold (20KB) - images smaller than this will be skipped

  // Attachment whitelist - MIME types that should always be saved
  // These are considered "real" attachments rather than embedded content
  attachmentTypesWhitelist: [
    "application/pdf",
    "application/msword",
    "application/vnd.openxmlformats-officedocument.wordprocessingml.document", // docx
    "application/vnd.ms-excel",
    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", // xlsx
    "application/vnd.ms-powerpoint",
    "application/vnd.openxmlformats-officedocument.presentationml.presentation", // pptx
    "application/zip",
    "application/x-zip-compressed",
    "application/vnd.rar",
    "text/csv",
    "text/plain",
    "application/json",
  ],

  //=============================================================================
  // TECHNICAL SETTINGS - Advanced settings for script operation
  //=============================================================================

  // Retry logic for handling transient errors
  // RECOMMENDED VALUES:
  // - maxRetries: 3-5 for most environments
  // - retryDelay: 1000ms (1 second) is a good starting point
  // - maxRetryDelay: 10000-30000ms (10-30 seconds) depending on operation criticality
  // INTERDEPENDENCY: These settings work together to implement exponential backoff
  maxRetries: 3, // Maximum number of retry attempts for operations
  retryDelay: 1000, // Initial delay in milliseconds before first retry
  maxRetryDelay: 10000, // Maximum delay between retries (for exponential backoff)
  logLevel: "INFO", // DEBUG | INFO | WARNING | ERROR
};

//=============================================================================
// UTILS - UTILITIES
//=============================================================================

/**
 * Builds a per-user lock key for script properties.
 *
 * @param {string} userEmail - User email used to scope the lock
 * @returns {string} Stable script property key for this user
 */
function getExecutionLockKey(userEmail) {
  const email = (userEmail || Session.getEffectiveUser().getEmail()).toLowerCase();
  const normalizedEmail = email.replace(/[^a-z0-9]/g, "_");
  return `EXECUTION_LOCK_${normalizedEmail}`;
}

/**
 * Cleans up the legacy global lock key when it is invalid or expired.
 * Kept temporarily for backward compatibility during migration to per-user locks.
 */
function cleanupLegacyExecutionLock() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const legacyKey = "EXECUTION_LOCK";
  const legacyInfo = scriptProperties.getProperty(legacyKey);

  if (!legacyInfo) return;

  try {
    const lockData = JSON.parse(legacyInfo);
    const lockTime = lockData.timestamp;
    const now = new Date().getTime();
    const isExpired = now - lockTime >= CONFIG.executionLockTime * 60 * 1000;

    if (isExpired) {
      scriptProperties.deleteProperty(legacyKey);
      logWithUser("Removed expired legacy global execution lock", "INFO");
    }
  } catch (e) {
    scriptProperties.deleteProperty(legacyKey);
    logWithUser("Removed invalid legacy global execution lock", "INFO");
  }
}

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
const LOG_LEVEL_PRIORITY = {
  DEBUG: 10,
  INFO: 20,
  WARNING: 30,
  ERROR: 40,
};

/**
 * Checks if a log level should be emitted under current configuration.
 *
 * @param {string} level - Candidate log level
 * @returns {boolean} True when message should be logged
 */
function shouldLogLevel(level) {
  const configured = String(CONFIG.logLevel || "INFO").toUpperCase();
  const threshold = LOG_LEVEL_PRIORITY[configured] || LOG_LEVEL_PRIORITY.INFO;
  const candidate = LOG_LEVEL_PRIORITY[String(level || "INFO").toUpperCase()];
  return (candidate || LOG_LEVEL_PRIORITY.INFO) >= threshold;
}

function logWithUser(message, level = "INFO") {
  // Handle undefined or null message
  if (message === undefined || message === null) {
    message = "[No message provided]";
  }
  const normalizedLevel = String(level || "INFO").toUpperCase();
  if (!shouldLogLevel(normalizedLevel)) return;

  const userEmail = Session.getEffectiveUser().getEmail();
  const timestamp = new Date().toISOString();
  const logMessage = `[${timestamp}] [${normalizedLevel}] [${userEmail}] ${message}`;
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
  const lockKey = getExecutionLockKey(userEmail);

  // Migration helper: clean old global lock key if it's stale/invalid.
  cleanupLegacyExecutionLock();

  // Try to acquire a per-user lock first (doesn't block other users).
  const lock = LockService.getUserLock();
  try {
    if (lock.tryLock(30000)) {
      // Try to lock for 30 seconds
      // Also set script property lock as backup
      const lockData = {
        user: userEmail,
        timestamp: new Date().getTime(),
      };
      scriptProperties.setProperty(lockKey, JSON.stringify(lockData));
      logWithUser(`Lock acquired by ${userEmail} using UserLock`);
      return true;
    }
  } catch (e) {
    logWithUser(`UserLock failed: ${e.message}`, "WARNING");
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
  const lockKey = getExecutionLockKey(userEmail);

  // Try to release per-user LockService lock first
  const lock = LockService.getUserLock();
  try {
    if (lock.hasLock()) {
      lock.releaseLock();
      logWithUser("Released UserLock lock");
    }
  } catch (e) {
    logWithUser(`Error releasing UserLock lock: ${e.message}`, "WARNING");
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
 * Handles both simple email addresses and those with display names
 *
 * @param {string} email - The email address to extract the domain from
 * @returns {string} The domain part of the email address, or "unknown" if not found
 */
function extractDomain(email) {
  if (!email) return "unknown";

  // First, try to extract email from "Display Name <email@domain.com>" format
  const angleMatch = email.match(/<([^>]+)>/);
  const cleanEmail = angleMatch ? angleMatch[1] : email;

  // Now extract the domain from the clean email
  const domainMatch = cleanEmail.match(/@([\w.-]+)/);
  return domainMatch ? domainMatch[1] : "unknown";
}

/**
 * Extracts unique external recipient domains from a Gmail message.
 *
 * Parses To: and CC: fields, removes the sender's own domain and any
 * domain in CONFIG.skipDomains, and returns deduplicated external domains.
 * Used to route sent-email attachments to recipient domain folders.
 *
 * @param {GmailMessage} message - The Gmail message to inspect
 * @param {string} ownDomain - The current user's domain to exclude
 * @returns {string[]} Unique external domain strings (may be empty)
 */
function extractExternalRecipientDomains(message, ownDomain) {
  const toField = message.getTo() || "";
  const ccField = message.getCc() || "";
  const combined = [toField, ccField].join(",");

  if (!combined.trim()) return [];

  const seen = new Set();
  // Gmail normalises To:/CC: headers before returning them, so quoted display-name
  // commas (RFC 5322) are not a concern in practice.
  combined.split(",").forEach(function(addr) {
    const domain = extractDomain(addr.trim());
    if (
      domain !== "unknown" &&
      domain !== ownDomain &&
      // Uses script-level CONFIG.skipDomains (not per-user skipDomains), consistent
      // with sender-domain filtering elsewhere in GmailProcessing.gs.
      !CONFIG.skipDomains.includes(domain)
    ) {
      seen.add(domain);
    }
  });

  return Array.from(seen);
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

//=============================================================================
// USERMANAGEMENT - USER MANAGEMENT
//=============================================================================

/**
 * Verify if a user has granted all required permissions
 *
 * @param {string} userEmail - The user's email address to verify. If not provided, uses current user.
 * @returns {boolean} True if the user has all required permissions
 *
 * The function follows this flow:
 * 1. Checks if the user is the current user or another user
 * 2. Looks for cached permission results to avoid redundant checks
 * 3. For the current user:
 *    - Attempts to access Gmail and Drive services, which triggers permission prompts
 *    - Returns true only if both services are accessible
 * 4. For other users:
 *    - Checks if they have already granted Gmail and Drive permissions
 *    - Cannot trigger permission prompts for other users
 * 5. Caches the result for future checks within the same script execution
 *
 * This verification is crucial for multi-user scripts to ensure each user
 * has granted the necessary permissions before attempting to process their data.
 */
function verifyUserPermissions(userEmail) {
  // If userEmail is not provided, use the current user's email
  if (!userEmail) {
    userEmail = Session.getEffectiveUser().getEmail();
    logWithUser(
      `No user email provided, using current user: ${userEmail}`,
      "INFO"
    );
  }
  const currentUser = Session.getEffectiveUser().getEmail();
  const isCurrentUser = currentUser === userEmail;

  // Use script properties to track users we've already checked
  const scriptProperties = PropertiesService.getScriptProperties();
  const checkedUsersKey = "CHECKED_USERS_CACHE";

  try {
    // Try to get the cache of checked users
    let checkedUsersCache = scriptProperties.getProperty(checkedUsersKey);
    let checkedUsers = checkedUsersCache ? JSON.parse(checkedUsersCache) : {};

    // If we've checked this user recently (within this script run), use cached result
    if (checkedUsers[userEmail] !== undefined) {
      return checkedUsers[userEmail];
    }

    let result = false;

    if (isCurrentUser) {
      // For the current user, we'll attempt to force permission requests if needed
      try {
        // Try to access Gmail - this should trigger permission request if needed
        const labels = GmailApp.getUserLabels();

        // Try to access Drive - this should trigger permission request if needed
        const rootFolder = DriveApp.getRootFolder();

        logWithUser(`User ${userEmail} has all required permissions`, "INFO");
        result = true;
      } catch (e) {
        // If we get here with the current user, it means permissions weren't granted even after prompting
        logWithUser(
          `User ${userEmail} denied or has not granted required permissions: ${e.message}`,
          "WARNING"
        );
        result = false;
      }
    } else {
      // For other users, we can only check if they already have permissions
      // Try to access user's Gmail
      let hasGmailAccess = false;
      let hasDriveAccess = false;

      try {
        const userGmail = GmailApp.getUserLabelByName("INBOX");
        hasGmailAccess = true;
      } catch (e) {
        logWithUser(
          `User ${userEmail} has not granted Gmail permissions`,
          "WARNING"
        );
      }

      try {
        const userDrive = DriveApp.getRootFolder();
        hasDriveAccess = true;
      } catch (e) {
        logWithUser(
          `User ${userEmail} has not granted Drive permissions`,
          "WARNING"
        );
      }

      result = hasGmailAccess && hasDriveAccess;

      if (result) {
        logWithUser(`User ${userEmail} has all required permissions`, "INFO");
      } else {
        logWithUser(
          `User ${userEmail} is missing some required permissions`,
          "WARNING"
        );
      }
    }

    // Cache the result
    checkedUsers[userEmail] = result;
    scriptProperties.setProperty(checkedUsersKey, JSON.stringify(checkedUsers));

    // Log the final result when called directly
    if (!userEmail || userEmail === currentUser) {
      logWithUser(
        `Permission verification result for current user: ${
          result ? "GRANTED" : "DENIED"
        }`,
        "INFO"
      );
    }

    return result;
  } catch (e) {
    logWithUser(
      `Error verifying permissions for ${userEmail}: ${e.message}`,
      "ERROR"
    );
    return false;
  }
}

/**
 * Explicitly request all permissions needed by the script
 * This function should be run manually by each user to grant permissions
 *
 * @returns {boolean} True if permissions were granted successfully
 *
 * The function follows this flow:
 * 1. Attempts to access Gmail, which triggers the permission prompt
 * 2. Attempts to access Drive, which triggers another permission prompt
 * 3. If the main folder ID is configured, verifies access to that folder
 * 4. Verifies all permissions using verifyUserPermissions()
 * 5. If all permissions are granted, adds the user to the authorized users list
 *
 * This is typically the first function a new user should run before using the script,
 * as it ensures all necessary permissions are granted and the user is properly registered.
 */
function requestPermissions() {
  try {
    const userEmail = Session.getEffectiveUser().getEmail();
    logWithUser("Requesting permissions for Gmail and Drive...", "INFO");

    // Try to access Gmail - this will trigger the permission prompt
    try {
      const gmailLabels = GmailApp.getUserLabels();
      logWithUser("Gmail permissions granted", "INFO");
    } catch (gmailError) {
      logWithUser(`Error accessing Gmail: ${gmailError.message}`, "ERROR");
      return false;
    }

    // Try to access Drive - this will trigger the permission prompt
    try {
      const rootFolder = DriveApp.getRootFolder();
      logWithUser("Drive permissions granted", "INFO");
    } catch (driveError) {
      logWithUser(`Error accessing Drive: ${driveError.message}`, "ERROR");
      return false;
    }

    // Verify that we can access the main folder
    if (CONFIG.mainFolderId !== "YOUR_SHARED_FOLDER_ID") {
      try {
        const mainFolder = DriveApp.getFolderById(CONFIG.mainFolderId);
        logWithUser(
          `Successfully accessed main folder: ${mainFolder.getName()}`,
          "INFO"
        );
      } catch (folderError) {
        // This is not a critical error, just log it
        logWithUser(
          `Error accessing main folder: ${folderError.message}. Please check the folder ID.`,
          "ERROR"
        );
        // Continue with the rest of the function
      }
    } else {
      logWithUser(
        "Please configure the mainFolderId in the CONFIG object before continuing.",
        "WARNING"
      );
    }

    // Let's add the current user to the manual list if they have all permissions
    try {
      const hasPermissions = verifyUserPermissions(userEmail);
      if (hasPermissions) {
        addUserToList(userEmail);
        logWithUser(
          "All required permissions have been granted successfully!",
          "INFO"
        );
        logWithUser(
          "You have been added to the users list and the script is ready to process your emails.",
          "INFO"
        );
      } else {
        logWithUser(
          "Some permissions appear to be missing. Please run this function again.",
          "WARNING"
        );
      }
      return hasPermissions;
    } catch (permissionError) {
      logWithUser(
        `Error verifying permissions: ${permissionError.message}`,
        "ERROR"
      );
      return false;
    }
  } catch (e) {
    logWithUser(`Error requesting permissions: ${e.message}`, "ERROR");
    return false;
  }
}

/**
 * Get all users who have access to the script and verify their permissions
 *
 * @returns {string[]} Array of user email addresses
 */
function getAuthorizedUsers() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let users = [];

  try {
    const usersProperty = scriptProperties.getProperty("REGISTERED_USERS");
    if (usersProperty) {
      users = JSON.parse(usersProperty);
    }
  } catch (e) {
    logWithUser(`Error parsing registered users: ${e.message}`, "ERROR");
  }

  const currentUser = Session.getEffectiveUser().getEmail();

  if (users.length === 0) {
    logWithUser(
      `No registered users found. Using current user: ${currentUser}`,
      "INFO"
    );
    users = [currentUser];
  } else {
    logWithUser(`Found ${users.length} registered users`, "INFO");
  }

  return users;
}

/**
 * Add a user to the authorized users list
 *
 * @param {string} email - Email address to add
 * @returns {boolean} True if the user was added successfully
 */
function addUserToList(email) {
  if (!email || !email.includes("@")) {
    logWithUser("Invalid email format", "ERROR");
    return false;
  }

  const scriptProperties = PropertiesService.getScriptProperties();
  let users = [];

  try {
    const usersProperty = scriptProperties.getProperty("REGISTERED_USERS");
    if (usersProperty) {
      users = JSON.parse(usersProperty);
    }

    if (!users.includes(email)) {
      users.push(email);
      scriptProperties.setProperty("REGISTERED_USERS", JSON.stringify(users));
      logWithUser(`User ${email} added to registered users list`, "INFO");
    } else {
      logWithUser(
        `User ${email} is already in the registered users list`,
        "INFO"
      );
    }

    return true;
  } catch (e) {
    logWithUser(`Error adding user ${email}: ${e.message}`, "ERROR");
    return false;
  }
}

/**
 * Remove a user from the authorized users list
 *
 * @param {string} email - Email address to remove
 * @returns {boolean} True if the user was removed successfully
 */
function removeUserFromList(email) {
  if (!email || !email.includes("@")) {
    logWithUser("Invalid email format", "ERROR");
    return false;
  }

  const scriptProperties = PropertiesService.getScriptProperties();
  let users = [];

  try {
    const usersProperty = scriptProperties.getProperty("REGISTERED_USERS");
    if (usersProperty) {
      users = JSON.parse(usersProperty);
    }

    const index = users.indexOf(email);
    if (index > -1) {
      users.splice(index, 1);
      scriptProperties.setProperty("REGISTERED_USERS", JSON.stringify(users));
      logWithUser(`User ${email} removed from registered users list`, "INFO");
    } else {
      logWithUser(`User ${email} not found in registered users list`, "INFO");
    }

    return true;
  } catch (e) {
    logWithUser(`Error removing user ${email}: ${e.message}`, "ERROR");
    return false;
  }
}

/**
 * List all registered users
 *
 * @returns {string[]} Array of user email addresses
 */
function listUsers() {
  const users = getAuthorizedUsers();

  if (users.length === 0) {
    logWithUser("No registered users found", "INFO");
  } else {
    logWithUser(`Registered users (${users.length}):`, "INFO");
    users.forEach((email, index) => {
      logWithUser(`${index + 1}. ${email}`, "INFO");
    });
  }

  return users;
}

/**
 * Get the next user in the queue
 *
 * @returns {string} Email address of the next user to process
 */
function getNextUserInQueue() {
  const currentUser = Session.getEffectiveUser().getEmail();
  logWithUser(
    `Execution model "${CONFIG.executionModel}" active: queue disabled, using effective user ${currentUser}`,
    "INFO"
  );
  return currentUser;
}

/**
 * Function to verify all users' permissions
 */
function verifyAllUsersPermissions() {
  // Clear the checked users cache to ensure fresh checks
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty("CHECKED_USERS_CACHE");

  // First check if the current user has permissions
  const currentUser = Session.getEffectiveUser().getEmail();
  const currentUserHasPermissions = verifyUserPermissions(currentUser);

  if (!currentUserHasPermissions) {
    logWithUser(
      "You need to grant necessary permissions first. Please run the 'requestPermissions' function.",
      "WARNING"
    );
    return;
  }

  // Now check all users in the list
  const users = getAuthorizedUsers();

  if (users.length === 0) {
    logWithUser(
      "No authorized users found. You may need to add users using the 'addUserToList' function.",
      "INFO"
    );
    // Automatically add the current user if they're not in the list but have permissions
    if (currentUserHasPermissions) {
      addUserToList(currentUser);
      logWithUser(
        `Added current user ${currentUser} to the registered users list`,
        "INFO"
      );
    }
  } else if (users.length > 1) {
    // Only display individual user status if there are more than one user
    logWithUser(`Checking permissions for ${users.length} users...`, "INFO");

    users.forEach((email) => {
      // Skip logging for current user to avoid redundancy
      if (email !== currentUser) {
        verifyUserPermissions(email); // This will log the results
      }
    });
  }

  logWithUser("Permission verification completed", "INFO");
}

//=============================================================================
// LABELMANAGEMENT - LABEL MANAGEMENT
//=============================================================================

/**
 * Gets or creates a Gmail label by name
 *
 * @param {string} labelName - Label name to fetch/create
 * @returns {GmailLabel} Gmail label instance
 */
function getOrCreateLabel(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);

  if (!label) {
    label = GmailApp.createLabel(labelName);
    logWithUser(`Created new Gmail label: ${labelName}`);
  } else {
    logWithUser(`Using existing Gmail label: ${labelName}`);
  }

  return label;
}

/**
 * Gets or creates the processed label
 *
 * @returns {GmailLabel} The Gmail label used to mark processed threads
 */
function getProcessedLabel() {
  return getOrCreateLabel(CONFIG.processedLabelName);
}

/**
 * Gets or creates the processing label
 *
 * @returns {GmailLabel} Label used while a thread is being processed
 */
function getProcessingLabel() {
  return getOrCreateLabel(CONFIG.processingLabelName);
}

/**
 * Gets or creates the error label
 *
 * @returns {GmailLabel} Label used when processing fails
 */
function getErrorLabel() {
  return getOrCreateLabel(CONFIG.errorLabelName);
}

/**
 * Gets or creates the permanent error label
 *
 * @returns {GmailLabel} Label used for non-retriable failures
 */
function getPermanentErrorLabel() {
  return getOrCreateLabel(CONFIG.permanentErrorLabelName);
}

/**
 * Gets or creates the too-large label
 *
 * @returns {GmailLabel} Label used when attachments exceed max size
 */
function getTooLargeLabel() {
  return getOrCreateLabel(CONFIG.tooLargeLabelName);
}

//=============================================================================
// THREADSTATE - THREAD STATE
//=============================================================================

// ---------------------------------------------------------------------------
// Processing state (checkpoint / stale-recovery)
// ---------------------------------------------------------------------------

/**
 * Builds script-property key for per-thread processing state.
 *
 * @param {string} threadId - Gmail thread ID
 * @returns {string} Script property key
 */
function buildThreadProcessingStateKey(threadId) {
  return `THREAD_PROCESSING_${threadId}`;
}

/**
 * Stores processing start state for a thread.
 *
 * @param {string} threadId - Gmail thread ID
 * @param {string} userEmail - Effective execution user
 */
function markThreadProcessingState(threadId, userEmail) {
  if (!threadId) return;
  const state = {
    user: userEmail || Session.getEffectiveUser().getEmail(),
    timestamp: new Date().getTime(),
  };
  PropertiesService.getScriptProperties().setProperty(
    buildThreadProcessingStateKey(threadId),
    JSON.stringify(state)
  );
}

/**
 * Clears processing state checkpoint for a thread.
 *
 * @param {string} threadId - Gmail thread ID
 */
function clearThreadProcessingState(threadId) {
  if (!threadId) return;
  PropertiesService.getScriptProperties().deleteProperty(
    buildThreadProcessingStateKey(threadId)
  );
}

/**
 * Reads processing state checkpoint for a thread.
 *
 * @param {string} threadId - Gmail thread ID
 * @returns {Object|null} Parsed state or null
 */
function getThreadProcessingState(threadId) {
  if (!threadId) return null;
  const raw = PropertiesService.getScriptProperties().getProperty(
    buildThreadProcessingStateKey(threadId)
  );
  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (e) {
    clearThreadProcessingState(threadId);
    return null;
  }
}

/**
 * Recovers stale Processing states left by interrupted executions.
 *
 * This runs before the normal processing loop and only scans a small page
 * per execution to keep runtime/cost predictable.
 *
 * @param {string} userEmail - Effective execution user
 * @param {number|null} deadlineMs - Unix timestamp (ms) soft deadline
 * @returns {Object} Recovery stats
 */
function recoverStaleProcessingThreads(userEmail, deadlineMs = null) {
  const processingLabel = getProcessingLabel();
  const searchCriteria = `label:${CONFIG.processingLabelName} -label:${CONFIG.processedLabelName}`;
  const pageSize = Math.max(1, CONFIG.staleRecoveryBatchSize || CONFIG.batchSize);
  const staleThresholdMs =
    Math.max(1, CONFIG.processingStateTtlMinutes) * 60 * 1000;
  const now = new Date().getTime();

  let checked = 0;
  let recovered = 0;
  let skippedRecent = 0;

  try {
    const threads = GmailApp.search(searchCriteria, 0, pageSize);
    if (threads.length === 0) {
      return { checked, recovered, skippedRecent };
    }

    for (let i = 0; i < threads.length; i++) {
      if (deadlineMs && new Date().getTime() >= deadlineMs) {
        logWithUser(
          "Soft deadline reached during stale recovery; resuming on next run",
          "WARNING"
        );
        break;
      }

      const thread = threads[i];
      const threadId = thread.getId();
      const state = getThreadProcessingState(threadId);
      checked++;

      let isStale = false;
      if (!state || typeof state.timestamp !== "number") {
        isStale = true;
      } else if (
        state.user &&
        userEmail &&
        state.user !== userEmail &&
        now - state.timestamp < staleThresholdMs
      ) {
        skippedRecent++;
        continue;
      } else if (now - state.timestamp >= staleThresholdMs) {
        isStale = true;
      }

      if (!isStale) {
        skippedRecent++;
        continue;
      }

      withRetry(
        () => thread.removeLabel(processingLabel),
        "removing stale processing label"
      );
      clearThreadProcessingState(threadId);
      recovered++;
    }

    if (checked > 0) {
      logWithUser(
        `Stale recovery checked ${checked} thread(s): ${recovered} recovered, ${skippedRecent} still recent`,
        "INFO"
      );
    }

    return { checked, recovered, skippedRecent };
  } catch (error) {
    logWithUser(`Error recovering stale Processing threads: ${error.message}`, "WARNING");
    return { checked, recovered, skippedRecent };
  }
}

// ---------------------------------------------------------------------------
// Failure state (retry counting + TTL expiry)
// ---------------------------------------------------------------------------

/**
 * Builds script-property key for thread failure state.
 *
 * @param {string} threadId - Gmail thread ID
 * @returns {string} Stable key for thread failure state
 */
function buildThreadFailureStateKey(threadId) {
  return `THREAD_FAILURE_${threadId}`;
}

/**
 * Reads and validates per-thread failure state.
 *
 * @param {string} threadId - Gmail thread ID
 * @returns {Object|null} Failure state or null
 */
function getThreadFailureState(threadId) {
  if (!threadId) return null;

  const key = buildThreadFailureStateKey(threadId);
  const scriptProperties = PropertiesService.getScriptProperties();
  const raw = scriptProperties.getProperty(key);
  if (!raw) return null;

  try {
    const state = JSON.parse(raw);
    const now = new Date().getTime();
    const ttlMs = Math.max(1, CONFIG.threadFailureStateTtlDays) * 24 * 60 * 60 * 1000;
    const lastFailureAt = Number(state.lastFailureAt || 0);

    if (lastFailureAt > 0 && now - lastFailureAt > ttlMs) {
      scriptProperties.deleteProperty(key);
      return null;
    }

    return state;
  } catch (e) {
    logWithUser(`getThreadFailureState: cleared corrupt state for thread ${threadId}: ${e.message}`, "DEBUG");
    scriptProperties.deleteProperty(key);
    return null;
  }
}

/**
 * Clears per-thread failure state.
 *
 * @param {string} threadId - Gmail thread ID
 */
function clearThreadFailureState(threadId) {
  if (!threadId) return;
  PropertiesService.getScriptProperties().deleteProperty(
    buildThreadFailureStateKey(threadId)
  );
}

/**
 * Classifies a processing failure as transient/permanent/too_large.
 *
 * @param {string} message - Error message
 * @param {string} codeHint - Optional code hint
 * @returns {{category: string, code: string}} Classification output
 */
function classifyProcessingFailure(message, codeHint = "") {
  const normalized = `${codeHint || ""} ${message || ""}`.toLowerCase();

  if (
    normalized.includes("too large") ||
    normalized.includes("exceeds max size") ||
    normalized.includes("file too big") ||
    normalized.includes("entity too large") ||
    normalized.includes("attachment_too_large")
  ) {
    return { category: "too_large", code: "too_large" };
  }

  if (
    normalized.includes("invalid or inaccessible folder") ||
    normalized.includes("file not found") ||
    normalized.includes("cannot find") ||
    normalized.includes("permission") ||
    normalized.includes("forbidden") ||
    normalized.includes("not authorized") ||
    normalized.includes("insufficient")
  ) {
    return { category: "permanent", code: "permission_or_config" };
  }

  if (
    normalized.includes("service invoked too many times") ||
    normalized.includes("rate limit") ||
    normalized.includes("quota") ||
    normalized.includes("timed out") ||
    normalized.includes("internal error") ||
    normalized.includes("backend error") ||
    normalized.includes("temporar") ||
    normalized.includes("try again")
  ) {
    return { category: "transient", code: "service_transient" };
  }

  return { category: "transient", code: "unknown_failure" };
}

/**
 * Registers a failure state for a thread with retry classification.
 *
 * @param {string} threadId - Gmail thread ID
 * @param {Object} failure - Failure details
 * @returns {Object|null} Updated state
 */
function registerThreadFailure(threadId, failure = {}) {
  if (!threadId) return null;

  const previous = getThreadFailureState(threadId) || {};
  const now = new Date().getTime();
  const inferred = classifyProcessingFailure(failure.message || "", failure.code || "");

  let category = failure.category || inferred.category;
  let code = failure.code || inferred.code;
  const attempts = (Number(previous.attempts) || 0) + 1;

  if (
    category !== "permanent" &&
    category !== "too_large" &&
    attempts >= Math.max(1, CONFIG.maxThreadFailureRetries)
  ) {
    category = "permanent";
    code = "max_retries_exceeded";
  }

  const state = {
    attempts: attempts,
    category: category,
    code: code,
    context: failure.context || previous.context || "unspecified",
    message: String(failure.message || "").substring(0, 1000),
    attachmentName: failure.attachmentName || null,
    attachmentSize: Number(failure.attachmentSize || 0) || null,
    user:
      failure.userEmail ||
      previous.user ||
      Session.getEffectiveUser().getEmail(),
    firstFailureAt: Number(previous.firstFailureAt) || now,
    lastFailureAt: now,
  };

  PropertiesService.getScriptProperties().setProperty(
    buildThreadFailureStateKey(threadId),
    JSON.stringify(state)
  );

  logWithUser(
    `Thread failure registered: threadId=${threadId}, category=${state.category}, code=${state.code}, attempts=${state.attempts}`,
    state.category === "permanent" ? "ERROR" : "WARNING"
  );

  return state;
}

//=============================================================================
// ATTACHMENTFILTERS - ATTACHMENT FILTERS
//=============================================================================

/**
 * Checks if a file should be skipped based on filter rules
 *
 * Why it's necessary:
 * - Avoids storing unnecessary files such as:
 *   1. Small images that are usually email signatures or icons
 *   2. Specific file types like calendar invitations (.ics)
 *   3. Embedded/inline images from email body
 * - Improves storage efficiency and keeps folders organized
 * - Reduces processing time by filtering irrelevant files early
 *
 * @param {String} fileName - The name of the file
 * @param {Number} fileSize - The size of the file in bytes
 * @param {GmailAttachment} attachment - The Gmail attachment object (optional)
 * @returns {Boolean} - True if the file should be skipped, false otherwise
 */
function shouldSkipFile(fileName, fileSize, attachment = null) {
  // Check mime type if attachment is provided
  if (attachment) {
    const mimeType = attachment.getContentType();

    // Always keep files with these MIME types (like documents and archives)
    if (CONFIG.attachmentTypesWhitelist.includes(mimeType)) {
      logWithUser(
        `Keeping file with important MIME type ${mimeType}: ${fileName}`,
        "DEBUG"
      );
      return false; // Don't skip these types
    }

    // Check if it's an inline image (as opposed to a real attachment)
    try {
      // We can check some internal properties to try to determine if this is an inline image
      const contentDisposition = attachment.getContentDisposition
        ? attachment.getContentDisposition()
        : null;

      // If it's explicitly marked as "inline" and it's an image type
      if (
        contentDisposition &&
        contentDisposition.toLowerCase().includes("inline") &&
        mimeType.startsWith("image/")
      ) {
        logWithUser(
          `Skipping inline image with Content-Disposition "${contentDisposition}": ${fileName}`,
          "DEBUG"
        );
        return true;
      }
    } catch (e) {
      // Can't determine content disposition, continue with other checks
    }
  }

  // Check if the filename appears to be a Gmail embedded image URL or other email provider URLs
  if (
    (fileName.includes("mail.google.com/mail") &&
      (fileName.includes("view=fimg") || fileName.includes("disp=emb"))) ||
    // Outlook and Exchange embedded image URLs
    (fileName.includes("outlook.office") && fileName.includes("attachment")) ||
    // Yahoo Mail embedded images
    (fileName.includes("yimg.com") && fileName.includes("mail")) ||
    // General URL patterns often seen in HTML emails
    (fileName.startsWith("https://") &&
      (fileName.includes("?cid=") ||
        fileName.includes("&cid=") ||
        fileName.includes("&disp=") ||
        fileName.includes("&view=")))
  ) {
    logWithUser(`Skipping embedded email image URL: ${fileName}`, "DEBUG");
    return true;
  }

  // Common image and icon names used in email services
  const commonEmbeddedNames = [
    "box_logo",
    "logo",
    "icon",
    "banner",
    "header",
    "footer",
    "signature",
    "badge",
    "avatar",
    "profile",
    "divider",
    "spacer",
    "separator",
    "pixel",
    "background",
    "bg",
    "bullet",
    "social",
    "facebook",
    "twitter",
    "linkedin",
    "instagram",
  ];

  // Check for common embedded image names without extensions
  for (const name of commonEmbeddedNames) {
    if (
      fileName.toLowerCase() === name.toLowerCase() ||
      fileName.toLowerCase().includes(name.toLowerCase() + "_")
    ) {
      logWithUser(`Skipping common embedded element: ${fileName}`, "DEBUG");
      return true;
    }
  }

  // Detect common patterns for embedded/inline images in email body
  const embeddedImagePatterns = [
    /^image\d+\.(png|jpg|gif|jpeg)$/i, // Common Gmail embedded image pattern (image001.png)
    /^inline-/i, // Inline images often start with "inline-"
    /^Outlook-/i, // Outlook embedded images prefix
    /^emb_(image|embed)\d+/i, // Common embedded image pattern
    /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}(\.(png|jpg|gif|jpeg))?$/i, // UUID-style names often used for embedded images
    /^part_\d+\.\d+\.\d+$/i, // Gmail inline image format
    /^att\d+\.\d+$/i, // Another common email attachment format
  ];

  // Check if the filename matches any of the embedded image patterns
  for (const pattern of embeddedImagePatterns) {
    if (pattern.test(fileName)) {
      logWithUser(`Skipping embedded image: ${fileName}`, "DEBUG");
      return true;
    }
  }

  // Extract file extension - handle files without extensions
  const lastDotIndex = fileName.lastIndexOf(".");
  const fileExtension =
    lastDotIndex !== -1 ? fileName.substring(lastDotIndex).toLowerCase() : "";

  // Common document extensions we want to keep
  const keepExtensions = [
    ".pdf",
    ".doc",
    ".docx",
    ".xls",
    ".xlsx",
    ".ppt",
    ".pptx",
    ".zip",
    ".rar",
    ".7z",
    ".csv",
    ".txt",
  ];

  // Always keep documents regardless of size
  if (keepExtensions.includes(fileExtension)) {
    logWithUser(`Keeping document file: ${fileName}`, "DEBUG");
    return false; // Don't skip these types
  }

  // If no extension and likely an embedded image/content (often binary content)
  // Most email service logos, icons and embedded HTML images are under 50KB
  if (fileExtension === "" && fileSize < 50 * 1024) {
    logWithUser(
      `Skipping potential embedded content without extension: ${fileName} (${Math.round(
        fileSize / 1024
      )}KB)`,
      "DEBUG"
    );
    return true;
  }

  // Check if we should skip files by extension (like calendar invitations)
  if (CONFIG.skipFileTypes && CONFIG.skipFileTypes.includes(fileExtension)) {
    logWithUser(`Skipping file type: ${fileName} (${fileExtension})`, "DEBUG");
    return true;
  }

  // Check if we should skip small images
  if (CONFIG.skipSmallImages) {
    // If it's an image extension in our filter list and smaller than the threshold
    if (
      CONFIG.smallImageExtensions.includes(fileExtension) &&
      fileSize <= CONFIG.smallImageMaxSize
    ) {
      logWithUser(
        `Skipping small image file: ${fileName} (${Math.round(
          fileSize / 1024
        )}KB)`,
        "DEBUG"
      );
      return true;
    }
  }

  return false;
}

//=============================================================================
// FOLDERMANAGEMENT - FOLDER MANAGEMENT
//=============================================================================

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

//=============================================================================
// ATTACHMENTPROCESSING - ATTACHMENT PROCESSING
//=============================================================================

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

//=============================================================================
// GMAILPROCESSING - GMAIL PROCESSING
//=============================================================================

/**
 * Builds a deterministic source ID for an attachment processing attempt.
 *
 * @param {string} threadId - Gmail thread ID
 * @param {string} messageId - Gmail message ID
 * @param {number} attachmentIndex - Attachment index within filtered message list
 * @param {GmailAttachment} attachment - Attachment object
 * @returns {string} Deterministic source ID
 */
function buildSourceAttachmentId(
  threadId,
  messageId,
  attachmentIndex,
  attachment
) {
  const name = attachment.getName();
  const size = attachment.getSize();
  return `${threadId}:${messageId}:${attachmentIndex}:${name}:${size}`;
}

/**
 * Processes emails for a specific user with optimized batch handling
 *
 * This function processes a single paginated page of unprocessed threads
 * per execution to keep runtime predictable and reduce resource usage.
 *
 * @param {string} userEmail - Email of the user to process
 * @param {boolean} oldestFirst - Whether to process oldest emails first (default: true)
 * @param {number|null} deadlineMs - Unix timestamp (ms) to stop safely before hard timeout
 * @returns {boolean} True if processing was successful, false if an error occurred
 *
 * The function follows this flow:
 * 1. Accesses the main Google Drive folder specified in CONFIG
 * 2. Gets or creates the Gmail label used to mark processed threads
 * 3. Builds search criteria to find unprocessed threads with attachments
 * 4. Retrieves one paginated page of matching threads using CONFIG.batchSize
 * 5. Optionally sorts threads in that page by date (oldest first)
 * 6. Processes that single page and applies processed labels as needed
 * 7. Returns success/failure status and logs processing statistics
 *
 * This paginated approach helps avoid hitting the 6-minute execution limit
 * by processing a controlled number of threads per execution.
 */
function processUserEmails(userEmail, oldestFirst = true, deadlineMs = null) {
  try {
    logWithUser(`Processing emails for user: ${userEmail}`, "INFO");

    // Get the main folder
    let mainFolder;
    try {
      mainFolder = DriveApp.getFolderById(CONFIG.mainFolderId);
      logWithUser(`Successfully accessed main folder: ${mainFolder.getName()}`);
    } catch (e) {
      throw new Error(
        `Invalid or inaccessible folder ID: ${CONFIG.mainFolderId}. Error: ${e.message}`
      );
    }

    // Get or create processing state labels
    let processedLabel = getProcessedLabel();
    let processingLabel = getProcessingLabel();
    let errorLabel = getErrorLabel();
    let permanentErrorLabel = getPermanentErrorLabel();
    let tooLargeLabel = getTooLargeLabel();

    // Build search criteria using the processedLabelName from config
    const searchCriteria =
      `has:attachment -label:${CONFIG.processedLabelName}` +
      ` -label:${CONFIG.permanentErrorLabelName}` +
      ` -label:${CONFIG.tooLargeLabelName}`;
    const pageSize = Math.max(1, CONFIG.batchSize);

    // Single pagination cycle per execution to keep runtime predictable
    const threads = GmailApp.search(searchCriteria, 0, pageSize);
    logWithUser(
      `Retrieved ${threads.length} threads from paginated search (offset=0, limit=${pageSize})`,
      "INFO"
    );

    if (threads.length === 0) {
      logWithUser(
        "No unprocessed threads with attachments found, skipping processing",
        "INFO"
      );
      return true;
    }

    // Preserve oldest-first preference within the current page.
    if (oldestFirst && threads.length > 1) {
      try {
        threads.sort(function (a, b) {
          const dateA = a.getLastMessageDate();
          const dateB = b.getLastMessageDate();
          return dateA - dateB;
        });
        logWithUser("Sorted paginated threads by date (oldest first)", "INFO");
      } catch (e) {
        logWithUser(
          `Error sorting paginated threads: ${e.message}. Using default order.`,
          "WARNING"
        );
      }
    }

    // Debug: Log subject of first thread for troubleshooting
    logWithUser(
      `First paginated thread subject: "${threads[0].getFirstMessageSubject()}" from ${threads[0]
        .getLastMessageDate()
        .toISOString()}`,
      "INFO"
    );

    // Process exactly one page per execution
    const result = processThreadsWithCounting(
      threads,
      mainFolder,
      processedLabel,
      processingLabel,
      errorLabel,
      permanentErrorLabel,
      tooLargeLabel,
      deadlineMs,
      userEmail
    );
    const threadsProcessed = result.processedThreads;
    const threadsWithValidAttachments = result.threadsWithAttachments;

    logWithUser(
      `Completed paginated processing: ${threadsProcessed} threads, ${threadsWithValidAttachments} with valid attachments for user: ${userEmail}`,
      "INFO"
    );
    if (result.stoppedByDeadline) {
      logWithUser(
        `Stopped early due to execution soft limit; remaining threads will continue in next run`,
        "WARNING"
      );
    }
    return true;
  } catch (error) {
    logWithUser(
      `Failed to process emails for user ${userEmail}: ${error.message}`,
      "ERROR"
    );
    if (error.stack) {
      logWithUser(`Stack trace: ${error.stack}`, "ERROR");
    }
    return false;
  }
}

/**
 * Process Gmail threads and count those that have valid attachments
 * This is a modified version of processThreads that counts threads with valid attachments
 *
 * @param {GmailThread[]} threads - Gmail threads to process
 * @param {DriveFolder} mainFolder - The main folder to save attachments to
 * @param {GmailLabel} processedLabel - The label to apply to processed threads
 * @param {GmailLabel} processingLabel - The label to apply while processing
 * @param {GmailLabel} errorLabel - The label to apply on processing failure
 * @param {GmailLabel} permanentErrorLabel - The label for non-retriable failures
 * @param {GmailLabel} tooLargeLabel - The label for too-large attachments
 * @param {number|null} deadlineMs - Unix timestamp (ms) to stop safely before hard timeout
 * @returns {Object} Object containing count of threads with valid attachments and processing status
 *
 * The function follows this flow:
 * 1. Iterates through each thread in the provided array
 * 2. Checks if the thread is already processed (has the processed label)
 * 3. For unprocessed threads:
 *    - Examines each message in the thread for attachments
 *    - Filters attachments based on sender domain and attachment properties
 *    - Creates domain folders as needed and saves valid attachments
 *    - Verifies file timestamps match email dates
 * 4. Marks threads as processed only when safe (or when all attachments were filtered out)
 * 5. Counts and returns processing statistics for the page
 *
 * This function is optimized for batch processing to handle the 6-minute execution limit
 * of Google Apps Script by focusing on counting threads with valid attachments.
 */
function processThreadsWithCounting(
  threads,
  mainFolder,
  processedLabel,
  processingLabel,
  errorLabel,
  permanentErrorLabel,
  tooLargeLabel,
  deadlineMs = null,
  userEmail = null
) {
  let threadsWithAttachments = 0;
  let processedThreads = 0;
  let stoppedByDeadline = false;
  const resolvedUserEmail =
    userEmail || Session.getEffectiveUser().getEmail();

  for (let i = 0; i < threads.length; i++) {
    if (deadlineMs && new Date().getTime() >= deadlineMs) {
      stoppedByDeadline = true;
      logWithUser(
        "Execution soft limit reached; stopping thread loop for safe resume",
        "WARNING"
      );
      break;
    }

    const thread = threads[i];
    let threadId = null;
    let processingLabelApplied = false;
    try {
      // Check if the thread is already processed
      const threadLabels = thread.getLabels();
      const isAlreadyProcessed = threadLabels.some(
        (label) => label.getName() === processedLabel.getName()
      );

      if (!isAlreadyProcessed) {
        threadId = thread.getId();
        const previousFailure = getThreadFailureState(threadId);
        if (previousFailure && previousFailure.category === "permanent") {
          withRetry(
            () => thread.addLabel(permanentErrorLabel),
            "adding permanent error label from stored state"
          );
          logWithUser(
            `Skipping thread with permanent failure state: ${threadId} (${previousFailure.code})`,
            "WARNING"
          );
          processedThreads++;
          continue;
        }

        const messages = thread.getMessages();
        let threadProcessed = false;
        let threadHasAttachments = false;
        let threadHadValidAttachments = false;
        let threadHadSaveFailures = false;
        let threadHasTooLargeAttachments = false;
        let threadMarkedPermanentFailure = false;
        const threadSubject = thread.getFirstMessageSubject();
        logWithUser(`Processing thread: ${threadSubject}`);
        withRetry(
          () => thread.addLabel(processingLabel),
          "adding processing label"
        );
        processingLabelApplied = true;
        try {
          markThreadProcessingState(
            threadId,
            Session.getEffectiveUser().getEmail()
          );
        } catch (processingStateError) {
          logWithUser(
            `Failed to store processing checkpoint: ${processingStateError.message}`,
            "WARNING"
          );
        }

        // Process each message in the thread
        for (let j = 0; j < messages.length; j++) {
          const message = messages[j];
          const attachments = message.getAttachments();
          const messageId = message.getId();
          const sender = message.getFrom();
          // Get the message date to use for file creation date
          const messageDate = message.getDate();

          // Check if sender's domain should be skipped
          const senderDomain = sender.match(/@([\w.-]+)/)?.[1];
          if (senderDomain && CONFIG.skipDomains.includes(senderDomain)) {
            logWithUser(`Skipping domain: ${senderDomain}`, "INFO");
            continue;
          }

          if (attachments.length > 0) {
            // Mark that this thread has attachments, even if they're filtered out
            threadHasAttachments = true;

            logWithUser(
              `Found ${attachments.length} attachments in message from ${sender}${extractDomain(sender) === extractDomain(resolvedUserEmail) ? " (sent by me)" : ""}`,
              "DEBUG"
            );

            // Log MIME types to help diagnose what types of attachments are found
            attachments.forEach((att) => {
              try {
                logWithUser(
                  `Attachment: ${att.getName()}, Type: ${att.getContentType()}, Size: ${Math.round(
                    att.getSize() / 1024
                  )}KB`,
                  "DEBUG"
                );
              } catch (e) {
                // Skip if we can't get content type
              }
            });

            // Pre-filter attachments to see if any valid ones exist
            const validAttachments = [];
            for (let k = 0; k < attachments.length; k++) {
              const attachment = attachments[k];
              const fileName = attachment.getName();
              const fileSize = attachment.getSize();

              if (fileSize > CONFIG.maxFileSize) {
                threadHasTooLargeAttachments = true;
                registerThreadFailure(threadId, {
                  category: "too_large",
                  code: "too_large",
                  context: "attachment_filter",
                  message: `Attachment exceeds max size (${fileSize} bytes > ${CONFIG.maxFileSize})`,
                  attachmentName: fileName,
                  attachmentSize: fileSize,
                });
                logWithUser(
                  `Attachment too large, skipping and labeling thread: ${fileName}`,
                  "WARNING"
                );
                continue;
              }

              // Skip if it matches filter criteria
              if (shouldSkipFile(fileName, fileSize, attachment)) {
                logWithUser(`Skipping attachment: ${fileName}`, "DEBUG");
                continue;
              }

              validAttachments.push(attachment);
            }

            // Only create domain folder if we have valid attachments to save
            if (validAttachments.length > 0) {
              // For sent messages: route to each external recipient domain.
              // For received messages: route to the sender's domain (existing behavior).
              const isSentByMe =
                extractDomain(sender) === extractDomain(resolvedUserEmail);
              const targetDomains = isSentByMe
                ? extractExternalRecipientDomains(
                    message,
                    extractDomain(resolvedUserEmail)
                  // getDomainFolder calls extractDomain() internally, which requires an "@" to
                  // parse a domain. Prefix bare domain strings (e.g. "acme.com" → "@acme.com")
                  // since extractExternalRecipientDomains returns plain domain strings.
                  ).map((d) => "@" + d)
                : [sender];

              if (targetDomains.length === 0) {
                logWithUser(
                  `Skipping sent message with no external recipients: ${threadSubject}`,
                  "INFO"
                );
                continue;
              }

              threadHadValidAttachments = true;

              for (const domainTarget of targetDomains) {
                const domainFolder = getDomainFolder(domainTarget, mainFolder);

                // Process each valid attachment
                for (let k = 0; k < validAttachments.length; k++) {
                  const attachment = validAttachments[k];
                  const sourceAttachmentId = buildSourceAttachmentId(
                    threadId,
                    messageId,
                    k,
                    attachment
                  );
                  logWithUser(
                    `Processing attachment: ${attachment.getName()} (${Math.round(
                      attachment.getSize() / 1024
                    )}KB)`,
                    "DEBUG"
                  );
                  // Log the message date that will be used for the file timestamp
                  logWithUser(
                    `Email date for timestamp: ${messageDate.toISOString()}`,
                    "DEBUG"
                  );

                  // The same sourceAttachmentId is used across all domain targets for this attachment.
                  // Deduplication safety is maintained because saveAttachment checks by filename+size
                  // and then by source_attachment_id in the file description — the folderId scope
                  // differs per domain folder, so each domain copy is checked independently.
                  const saveResult = saveAttachment(attachment, message, domainFolder, {
                    sourceAttachmentId: `${sourceAttachmentId}:domain`,
                  });

                  // Process the result object
                  if (saveResult.success) {
                    const savedFile = saveResult.file;
                    threadProcessed = true;
                    // If the file was saved, verify its timestamp was set correctly
                    try {
                      // Get the file's timestamp using DriveApp
                      const updatedDate = savedFile.getLastUpdated();
                      logWithUser(
                        `Saved file modification date: ${updatedDate.toISOString()}`,
                        "DEBUG"
                      );

                      // Compare with the message date
                      const messageDateTime = messageDate.getTime();
                      const fileDateTime = updatedDate.getTime();
                      const diffInMinutes =
                        Math.abs(messageDateTime - fileDateTime) / (1000 * 60);

                      if (diffInMinutes < 5) {
                        logWithUser(
                          "✅ File timestamp matches email date (within 5 minutes)",
                          "INFO"
                        );
                      } else {
                        logWithUser(
                          `⚠️ File timestamp differs from email date by ${Math.round(
                            diffInMinutes
                          )} minutes`,
                          "WARNING"
                        );
                      }
                    } catch (e) {
                      // Just log the error but don't stop processing
                      logWithUser(
                        `Error verifying file timestamp: ${e.message}`,
                        "WARNING"
                      );
                    }
                  } else {
                    threadHadSaveFailures = true;
                    const failureState = registerThreadFailure(threadId, {
                      context: "save_attachment",
                      message:
                        saveResult.error ||
                        `Failed to save attachment ${attachment.getName()}`,
                      attachmentName: attachment.getName(),
                      attachmentSize: attachment.getSize(),
                    });
                    if (failureState && failureState.category === "permanent") {
                      threadMarkedPermanentFailure = true;
                    }
                    logWithUser(
                      `Failed to save valid attachment: ${attachment.getName()}`,
                      "WARNING"
                    );
                  }
                } // end for (let k ...)
              } // end for (const domainTarget of targetDomains)
            } else {
              logWithUser(
                `No valid attachments to save in message from ${sender}`,
                "INFO"
              );
            }
          }
        }

        // Mark as processed only for clean outcomes:
        // 1. Valid attachments saved without save failures or too-large items
        // 2. Thread had attachments but none were valid, and no too-large items
        if (threadProcessed && !threadHadSaveFailures && !threadHasTooLargeAttachments) {
          withRetry(
            () => thread.addLabel(processedLabel),
            "adding processed label"
          );
          withRetry(
            () => thread.removeLabel(errorLabel),
            "removing error label after successful processing"
          );
          withRetry(
            () => thread.removeLabel(permanentErrorLabel),
            "removing permanent error label after successful processing"
          );
          withRetry(
            () => thread.removeLabel(tooLargeLabel),
            "removing too-large label after successful processing"
          );
          clearThreadFailureState(threadId);
          logWithUser(
            `Thread "${threadSubject}" processed with valid attachments and labeled`
          );
          threadsWithAttachments++;
        } else if (
          threadHasAttachments &&
          !threadHadValidAttachments &&
          !threadHasTooLargeAttachments
        ) {
          withRetry(
            () => thread.addLabel(processedLabel),
            "adding processed label"
          );
          withRetry(
            () => thread.removeLabel(errorLabel),
            "removing error label after filtered processing"
          );
          withRetry(
            () => thread.removeLabel(permanentErrorLabel),
            "removing permanent error label after filtered processing"
          );
          withRetry(
            () => thread.removeLabel(tooLargeLabel),
            "removing too-large label after filtered processing"
          );
          clearThreadFailureState(threadId);
          logWithUser(
            `Thread "${threadSubject}" had attachments but none were valid; marked as processed`,
            "INFO"
          );
        } else if (threadHadValidAttachments && threadHadSaveFailures) {
          withRetry(() => thread.addLabel(errorLabel), "adding error label");
          if (threadMarkedPermanentFailure) {
            withRetry(
              () => thread.addLabel(permanentErrorLabel),
              "adding permanent error label"
            );
          }
          if (threadHasTooLargeAttachments) {
            withRetry(
              () => thread.addLabel(tooLargeLabel),
              "adding too-large label alongside save failures"
            );
          }
          logWithUser(
            `Thread "${threadSubject}" had valid attachments with save failures; not marking as processed for retry`,
            "WARNING"
          );
        } else if (threadHasTooLargeAttachments) {
          withRetry(() => thread.addLabel(tooLargeLabel), "adding too-large label");
          withRetry(
            () => thread.removeLabel(errorLabel),
            "removing error label for too-large-only thread"
          );
          withRetry(
            () => thread.removeLabel(permanentErrorLabel),
            "removing permanent error label for too-large-only thread"
          );
          logWithUser(
            `Thread "${threadSubject}" contains attachments above max size; marked as TooLarge`,
            "WARNING"
          );
        } else {
          logWithUser(
            `No attachments found in thread "${threadSubject}"`,
            "INFO"
          );
        }
      }
    } catch (error) {
      const failedThreadId = threadId || thread.getId();
      const failureState = registerThreadFailure(failedThreadId, {
        context: "thread_exception",
        message: error.message,
      });
      logWithUser(
        `Error processing thread "${thread.getFirstMessageSubject()}": ${
          error.message
        }`,
        "ERROR"
      );
      try {
        withRetry(
          () => thread.addLabel(errorLabel),
          "adding error label after thread exception"
        );
        if (failureState && failureState.category === "permanent") {
          withRetry(
            () => thread.addLabel(permanentErrorLabel),
            "adding permanent error label after thread exception"
          );
        }
      } catch (errorLabelError) {
        logWithUser(
          `Failed to add error label: ${errorLabelError.message}`,
          "WARNING"
        );
      }
      // Continue with next thread instead of stopping execution
    } finally {
      if (processingLabelApplied) {
        try {
          withRetry(
            () => thread.removeLabel(processingLabel),
            "removing processing label"
          );
        } catch (processingCleanupError) {
          logWithUser(
            `Failed to remove processing label: ${processingCleanupError.message}`,
            "WARNING"
          );
        }
      }
      if (threadId) {
        try {
          clearThreadProcessingState(threadId);
        } catch (processingStateCleanupError) {
          logWithUser(
            `Failed to clear processing state: ${processingStateCleanupError.message}`,
            "WARNING"
          );
        }
      }
    }

    processedThreads++;
  }

  return { threadsWithAttachments, processedThreads, stoppedByDeadline };
}

/**
 * Process the messages in a thread, saving attachments if any
 *
 * @param {GmailThread} thread - The Gmail thread to process
 * @param {GmailLabel} processedLabel - The label to apply to processed threads
 * @param {DriveFolder} mainFolder - The main Google Drive folder
 * @returns {Object} Processing results with counts of processed attachments
 *
 * The function follows this flow:
 * 1. Retrieves all messages in the thread
 * 2. For each message:
 *    - Gets all attachments and filters out those that should be skipped
 *    - Extracts the sender's domain to determine the target folder
 *    - Creates or uses an existing domain folder
 *    - Saves each valid attachment to the appropriate folder
 *    - Tracks statistics (saved, duplicates, errors, etc.)
 * 3. Applies the processed label to the thread regardless of outcome
 * 4. Returns a detailed result object with processing statistics
 *
 * This function is used by processThreadsWithCounting but provides more detailed
 * statistics about the processing results.
 */
function processMessages(thread, processedLabel, mainFolder) {
  try {
    const messages = thread.getMessages();
    const threadId = thread.getId();
    let result = {
      totalAttachments: 0,
      savedAttachments: 0,
      savedSize: 0,
      skippedAttachments: 0,
      errors: 0,
      duplicates: 0,
    };

    for (const message of messages) {
      try {
        const messageId = message.getId();
        const attachments = message.getAttachments();
        const validAttachments = attachments.filter(
          (att) => !shouldSkipFile(att.getName(), att.getSize(), att)
        );

        if (validAttachments.length === 0) {
          continue; // Skip messages with no valid attachments
        }

        // Get sender details
        const sender = message.getFrom();
        const domain = extractDomain(sender);

        // Get or create domain folder
        const domainFolder = getDomainFolder(sender, mainFolder);

        if (!domainFolder) {
          logWithUser(
            `Error: Could not find or create folder for domain ${domain}`,
            "ERROR"
          );
          result.errors++;
          continue;
        }

        result.totalAttachments += validAttachments.length;

        // Process each valid attachment
        for (let attachmentIndex = 0; attachmentIndex < validAttachments.length; attachmentIndex++) {
          const attachment = validAttachments[attachmentIndex];
          const sourceAttachmentId = buildSourceAttachmentId(
            threadId,
            messageId,
            attachmentIndex,
            attachment
          );

          // Save to domain folder
          const saveResult = saveAttachment(attachment, message, domainFolder, {
            sourceAttachmentId: `${sourceAttachmentId}:domain`,
          });

          if (saveResult.success) {
            if (saveResult.duplicate) {
              result.duplicates++;
            } else {
              result.savedAttachments++;
              result.savedSize += attachment.getSize();
            }
          } else {
            result.skippedAttachments++;
            result.errors++;
          }
        }
      } catch (messageError) {
        logWithUser(
          `Error processing message: ${messageError.message}`,
          "ERROR"
        );
        result.errors++;
      }
    }

    // Apply the processed label only when safe:
    // - no errors during valid-attachment saving, or
    // - there were no valid attachments to save
    if (result.errors === 0 || result.totalAttachments === 0) {
      thread.addLabel(processedLabel);
    } else {
      logWithUser(
        `Thread "${thread.getFirstMessageSubject()}" had save errors; not marked as processed`,
        "WARNING"
      );
    }

    // Log a summary for the thread if it had attachments
    if (result.totalAttachments > 0) {
      logWithUser(
        `Thread processed: ${result.savedAttachments} saved, ${result.duplicates} duplicates, ${result.skippedAttachments} skipped`,
        "INFO"
      );
    }

    return result;
  } catch (error) {
    logWithUser(`Error in processMessages: ${error.message}`, "ERROR");
    throw error;
  }
}

//=============================================================================
// MAIN - MAIN FUNCTIONS
//=============================================================================

/**
 * Validates the configuration settings to ensure all required values are properly set
 *
 * @return {boolean} True if the configuration is valid, throws an error otherwise
 */
function validateConfig() {
  // Check essential folder and label settings
  if (!CONFIG.mainFolderId || CONFIG.mainFolderId === "__FOLDER_ID__") {
    throw new Error(
      "Configuration error: mainFolderId is not set. Please set a valid Google Drive folder ID."
    );
  }

  if (
    !CONFIG.processedLabelName ||
    !CONFIG.processingLabelName ||
    !CONFIG.errorLabelName ||
    !CONFIG.permanentErrorLabelName ||
    !CONFIG.tooLargeLabelName
  ) {
    throw new Error(
      "Configuration error: processedLabelName, processingLabelName, errorLabelName, permanentErrorLabelName, and tooLargeLabelName must all be set."
    );
  }

  const labelNames = [
    CONFIG.processedLabelName,
    CONFIG.processingLabelName,
    CONFIG.errorLabelName,
    CONFIG.permanentErrorLabelName,
    CONFIG.tooLargeLabelName,
  ];
  if (
    new Set(labelNames).size !== labelNames.length
  ) {
    throw new Error(
      "Configuration error: processedLabelName, processingLabelName, errorLabelName, permanentErrorLabelName, and tooLargeLabelName must be different."
    );
  }

  if (CONFIG.batchSize < 1) {
    throw new Error("Configuration error: batchSize must be at least 1.");
  }

  if (CONFIG.triggerIntervalMinutes < 1) {
    throw new Error(
      "Configuration error: triggerIntervalMinutes must be at least 1."
    );
  }

  if (CONFIG.executionSoftLimitMs < 1000) {
    throw new Error(
      "Configuration error: executionSoftLimitMs must be at least 1000 milliseconds."
    );
  }

  if (CONFIG.processingStateTtlMinutes < 1) {
    throw new Error(
      "Configuration error: processingStateTtlMinutes must be at least 1 minute."
    );
  }

  if (CONFIG.staleRecoveryBatchSize < 1) {
    throw new Error(
      "Configuration error: staleRecoveryBatchSize must be at least 1."
    );
  }

  if (CONFIG.maxThreadFailureRetries < 1) {
    throw new Error(
      "Configuration error: maxThreadFailureRetries must be at least 1."
    );
  }

  if (CONFIG.threadFailureStateTtlDays < 1) {
    throw new Error(
      "Configuration error: threadFailureStateTtlDays must be at least 1."
    );
  }

  const allowedLogLevels = ["DEBUG", "INFO", "WARNING", "ERROR"];
  if (!allowedLogLevels.includes(String(CONFIG.logLevel || "").toUpperCase())) {
    throw new Error(
      "Configuration error: logLevel must be one of DEBUG, INFO, WARNING, ERROR."
    );
  }

  if (CONFIG.executionModel !== "effective_user_only") {
    throw new Error(
      "Configuration error: executionModel must be \"effective_user_only\"."
    );
  }

  return true;
}

/**
 * Main function that processes Gmail attachments for the effective user
 *
 * This function:
 * 1. Validates the configuration
 * 2. Acquires an execution lock to prevent concurrent runs
 * 3. Computes a safe execution deadline
 * 4. Processes emails for the current execution user
 * 5. Logs completion
 *
 * @return {boolean} True if processing completed successfully, false if an error occurred.
 * The function follows a structured flow:
 * - It first validates the configuration to ensure all required settings are properly set.
 * - It attempts to acquire a lock to ensure no concurrent executions.
 * - It computes a soft execution deadline to avoid hard timeouts.
 * - It processes the current user's emails to save attachments to Google Drive.
 * - Logs are generated throughout the process to provide detailed information on the execution status.
 */
function saveAttachmentsToDrive() {
  let currentUser = null;
  let lockAcquired = false;
  const executionStartMs = new Date().getTime();

  try {
    logWithUser("Starting attachment processing", "INFO");

    // Validate configuration
    validateConfig();
    logWithUser("Configuration validated successfully", "INFO");

    // Acquire lock to prevent concurrent executions
    currentUser = Session.getEffectiveUser().getEmail();
    if (!acquireExecutionLock(currentUser)) {
      logWithUser("Another instance is already running. Exiting.", "WARNING");
      return false;
    }
    lockAcquired = true;

    // Process only the effective user for this execution.
    // GmailApp runs in the current execution context, so iterating over a
    // registered user list would reprocess the same mailbox.
    const deadlineMs = executionStartMs + CONFIG.executionSoftLimitMs;
    recoverStaleProcessingThreads(currentUser, deadlineMs);
    logWithUser(`Processing attachments for current user: ${currentUser}`, "INFO");
    const processed = processUserEmails(currentUser, true, deadlineMs);
    if (!processed) {
      logWithUser(`Processing failed for user: ${currentUser}`, "ERROR");
      return false;
    }

    logWithUser("Attachment processing completed successfully", "INFO");
    return true;
  } catch (error) {
    logWithUser(`Error in saveAttachmentsToDrive: ${error.message}`, "ERROR");
    return false;
  } finally {
    if (lockAcquired) {
      try {
        releaseExecutionLock(currentUser);
      } catch (releaseError) {
        logWithUser(
          `Error releasing execution lock: ${releaseError.message}`,
          "WARNING"
        );
      }
    }
  }
}

/**
 * Creates a time-based trigger to run saveAttachmentsToDrive at specified intervals
 * If a trigger already exists, it will be deleted first
 *
 * @return {Trigger} The created trigger object, which represents the scheduled execution of the function.
 * The function performs the following steps:
 * - Deletes any existing triggers for the saveAttachmentsToDrive function to avoid duplicates.
 * - Creates a new time-based trigger using the interval specified in the CONFIG object.
 * - Logs the creation of the new trigger and returns the trigger object.
 */
function createTrigger() {
  try {
    // Delete existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === "saveAttachmentsToDrive") {
        ScriptApp.deleteTrigger(trigger);
        logWithUser("Deleted existing trigger", "INFO");
      }
    }

    // Create new trigger
    const trigger = ScriptApp.newTrigger("saveAttachmentsToDrive")
      .timeBased()
      .everyMinutes(CONFIG.triggerIntervalMinutes)
      .create();

    logWithUser(
      `Created new trigger to run every ${CONFIG.triggerIntervalMinutes} minutes`,
      "INFO"
    );
    return trigger;
  } catch (error) {
    logWithUser(`Error creating trigger: ${error.message}`, "ERROR");
    throw error;
  }
}
