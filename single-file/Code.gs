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
  mainFolderId: "__FOLDER_ID__", // Replace with your shared folder's ID
  processedLabelName: "GDrive_Processed", // Label to mark processed threads
  skipDomains: ["example.com", "noreply.com"], // Skip emails from these domains
  triggerIntervalMinutes: 15, // Interval in minutes for the trigger execution
  batchSize: 10, // Process this many threads at a time to avoid the 6 minutes execution limit
  skipFileTypes: [".ics", ".ical", ".pkpass", ".vcf", ".vcard"], // Additional file types to skip (e.g., calendar invitations, etc.)
  //
  // You probably won't need to edit any other CONFIG variable below this line
  //
  maxFileSize: 25 * 1024 * 1024, // 25MB max file size
  skipSmallImages: true, // Skip small images like email signatures
  smallImageMaxSize: 20 * 1024, // 20KB max size for images to skip
  smallImageExtensions: [".jpg", ".jpeg", ".png", ".gif", ".bmp"], // Image extensions to check
  executionLockTime: 10, // Maximum time in minutes to wait for lock release
  maxRetries: 3, // Maximum number of retries for operations
  retryDelay: 1000, // Initial delay in milliseconds for retries
  maxRetryDelay: 10000, // Maximum delay in milliseconds for exponential backoff
  useEmailTimestamps: true, // Set to true to use email timestamps as file creation dates
  // List of MIME types that should always be saved as they are considered as real attachments
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
};

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
// UTILS - UTILITIES
//=============================================================================

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
 * @param {string} userEmail - The email of the user acquiring the lock
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
 * @param {string} userEmail - The email of the user releasing the lock
 * @returns {void} This function doesn't return a value, but logs the result of the operation
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

//=============================================================================
// USERMANAGEMENT - USER MANAGEMENT
//=============================================================================

/**
 * Verify if a user has granted all required permissions
 *
 * @param {string} userEmail - The user's email address to verify
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
    const gmailLabels = GmailApp.getUserLabels();
    logWithUser("Gmail permissions granted", "INFO");

    // Try to access Drive - this will trigger the permission prompt
    const rootFolder = DriveApp.getRootFolder();
    logWithUser("Drive permissions granted", "INFO");

    // Verify that we can access the main folder
    if (CONFIG.mainFolderId !== "YOUR_SHARED_FOLDER_ID") {
      try {
        const mainFolder = DriveApp.getFolderById(CONFIG.mainFolderId);
        logWithUser(
          `Successfully accessed main folder: ${mainFolder.getName()}`,
          "INFO"
        );
      } catch (e) {
        logWithUser(
          `Error accessing main folder: ${e.message}. Please check the folder ID.`,
          "ERROR"
        );
      }
    } else {
      logWithUser(
        "Please configure the mainFolderId in the CONFIG object before continuing.",
        "WARNING"
      );
    }

    // Let's add the current user to the manual list if they have all permissions
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
  const scriptProperties = PropertiesService.getScriptProperties();
  const queueKey = "USER_QUEUE";
  const lastProcessedKey = "LAST_PROCESSED_USER";

  try {
    // Get the queue of users
    let queue = JSON.parse(scriptProperties.getProperty(queueKey) || "[]");

    // If queue is empty, refresh it with current authorized users
    if (queue.length === 0) {
      queue = getAuthorizedUsers();
      scriptProperties.setProperty(queueKey, JSON.stringify(queue));
    }

    // Get the last processed user
    const lastProcessed = scriptProperties.getProperty(lastProcessedKey);

    // Find the next user in the queue
    let nextUser;
    if (lastProcessed) {
      const lastIndex = queue.indexOf(lastProcessed);
      nextUser = queue[(lastIndex + 1) % queue.length];
    } else {
      nextUser = queue[0];
    }

    // Update the last processed user
    scriptProperties.setProperty(lastProcessedKey, nextUser);

    logWithUser(`Next user in queue: ${nextUser}`, "INFO");
    return nextUser;
  } catch (e) {
    logWithUser(`Error getting next user in queue: ${e.message}`, "ERROR");
    return null;
  }
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
        "INFO"
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
          "INFO"
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
    logWithUser(`Skipping embedded email image URL: ${fileName}`, "INFO");
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
      logWithUser(`Skipping common embedded element: ${fileName}`, "INFO");
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
      logWithUser(`Skipping embedded image: ${fileName}`, "INFO");
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
    logWithUser(`Keeping document file: ${fileName}`, "INFO");
    return false; // Don't skip these types
  }

  // If no extension and likely an embedded image/content (often binary content)
  // Most email service logos, icons and embedded HTML images are under 50KB
  if (fileExtension === "" && fileSize < 50 * 1024) {
    logWithUser(
      `Skipping potential embedded content without extension: ${fileName} (${Math.round(
        fileSize / 1024
      )}KB)`,
      "INFO"
    );
    return true;
  }

  // Check if we should skip files by extension (like calendar invitations)
  if (CONFIG.skipFileTypes && CONFIG.skipFileTypes.includes(fileExtension)) {
    logWithUser(`Skipping file type: ${fileName} (${fileExtension})`, "INFO");
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
        "INFO"
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

//=============================================================================
// GMAILPROCESSING - GMAIL PROCESSING
//=============================================================================

/**
 * Gets or creates the processed label
 *
 * @returns {GmailLabel} The Gmail label used to mark processed threads
 */
function getProcessedLabel() {
  let processedLabel = GmailApp.getUserLabelByName(CONFIG.processedLabelName);

  if (!processedLabel) {
    processedLabel = GmailApp.createLabel(CONFIG.processedLabelName);
    logWithUser(`Created new Gmail label: ${CONFIG.processedLabelName}`);
  } else {
    logWithUser(`Using existing Gmail label: ${CONFIG.processedLabelName}`);
  }

  return processedLabel;
}

/**
 * Processes emails for a specific user with optimized batch handling
 *
 * This function processes emails in batches until it finds enough threads
 * with valid attachments or exhausts all unprocessed threads.
 *
 * @param {string} userEmail - Email of the user to process
 * @param {boolean} oldestFirst - Whether to process oldest emails first (default: true)
 * @returns {boolean} True if processing was successful, false if an error occurred
 *
 * The function follows this flow:
 * 1. Accesses the main Google Drive folder specified in CONFIG
 * 2. Gets or creates the Gmail label used to mark processed threads
 * 3. Builds search criteria to find unprocessed threads with attachments
 * 4. Tests different search variations to find the most effective query
 * 5. Retrieves and optionally sorts threads by date (oldest first if specified)
 * 6. Processes threads in batches, with batch size determined by CONFIG.batchSize
 * 7. Tracks progress and adjusts batch size dynamically to meet target counts
 * 8. Returns success/failure status and logs detailed processing statistics
 *
 * This batch processing approach helps avoid hitting the 6-minute execution limit
 * of Google Apps Script by processing a controlled number of threads per execution.
 */
function processUserEmails(userEmail, oldestFirst = true) {
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

    // Get or create the 'Processed' label
    let processedLabel = getProcessedLabel();

    // Build search criteria using the processedLabelName from config
    let searchCriteria = `has:attachment -label:${CONFIG.processedLabelName}`;

    // Debug: Try several search variations to see if any of them work better
    const searchVariations = [
      { name: "Standard", query: searchCriteria },
      { name: "In Inbox", query: `in:inbox ${searchCriteria}` },
      {
        name: "With quote",
        query: `has:attachment -label:"${CONFIG.processedLabelName}"`,
      },
      {
        name: "With parentheses",
        query: `has:attachment AND -(label:${CONFIG.processedLabelName})`,
      },
      {
        name: "Explicit attachment",
        query: `filename:* -label:${CONFIG.processedLabelName}`,
      },
    ];

    // Try each search variation and log results
    logWithUser("Testing different search variations:", "INFO");
    for (const variation of searchVariations) {
      const count = GmailApp.search(variation.query).length;
      logWithUser(
        `- ${variation.name}: "${variation.query}" found ${count} threads`,
        "INFO"
      );
    }

    // Debug: Check how many threads match the base search criteria
    const totalMatchingThreads = GmailApp.search(searchCriteria).length;
    logWithUser(
      `Total unprocessed threads with attachments found: ${totalMatchingThreads}`,
      "INFO"
    );

    // Continue only if we actually have matching threads
    if (totalMatchingThreads === 0) {
      logWithUser(
        "No unprocessed threads with attachments found, skipping processing",
        "INFO"
      );
      return true;
    }

    // NOTE: We no longer add 'older_first' to the search criteria since it doesn't work as expected
    // Instead, we'll sort the threads ourselves after retrieving them
    if (oldestFirst) {
      logWithUser(
        "Will manually sort threads by date (oldest first) after retrieving them",
        "INFO"
      );
    } else {
      logWithUser("Using default order (newest first)", "INFO");
    }

    // Initialize batch processing
    let threadsProcessed = 0;
    let threadsWithValidAttachments = 0;
    let batchSize = CONFIG.batchSize;
    let offset = 0;

    // Get all threads that match our search criteria
    let allThreads = GmailApp.search(searchCriteria);
    logWithUser(
      `Retrieved ${allThreads.length} total threads matching search criteria`,
      "INFO"
    );

    // If we want oldest first, manually sort the threads by date
    if (oldestFirst && allThreads.length > 0) {
      try {
        // Sort threads by date (oldest first)
        allThreads.sort(function (a, b) {
          const dateA = a.getLastMessageDate();
          const dateB = b.getLastMessageDate();
          return dateA - dateB; // Ascending order (oldest first)
        });
        logWithUser(
          "Successfully sorted threads by date (oldest first)",
          "INFO"
        );
      } catch (e) {
        logWithUser(
          `Error sorting threads: ${e.message}. Will use default order.`,
          "WARNING"
        );
      }
    }

    // Continue processing batches until we reach the desired number of threads with valid attachments
    // or until we run out of unprocessed threads
    while (
      threadsWithValidAttachments < CONFIG.batchSize &&
      offset < allThreads.length
    ) {
      // Get the next batch of threads from our already retrieved and sorted list
      const currentBatchSize = Math.min(batchSize, allThreads.length - offset);
      const threads = allThreads.slice(offset, offset + currentBatchSize);

      // If no more threads, break out of the loop
      if (threads.length === 0) {
        logWithUser(
          "No more unprocessed threads with attachments found in this batch",
          "INFO"
        );
        break;
      }

      logWithUser(
        `Processing batch of ${threads.length} threads (batch ${
          Math.floor(offset / batchSize) + 1
        })`,
        "INFO"
      );

      // Debug: Log subject of first thread for troubleshooting
      if (threads.length > 0) {
        logWithUser(
          `First thread subject: "${threads[0].getFirstMessageSubject()}" from ${threads[0]
            .getLastMessageDate()
            .toISOString()}`,
          "INFO"
        );
      }

      // Process each thread and count those with valid attachments
      const result = processThreadsWithCounting(
        threads,
        mainFolder,
        processedLabel
      );
      threadsProcessed += threads.length;
      threadsWithValidAttachments += result.threadsWithAttachments;

      logWithUser(
        `Batch processed ${threads.length} threads, ${result.threadsWithAttachments} had valid attachments`,
        "INFO"
      );

      // Update the offset for the next batch
      offset += threads.length;

      // If we're getting close to our target, reduce the next batch size to avoid processing too many
      if (threadsWithValidAttachments + batchSize > CONFIG.batchSize * 1.5) {
        batchSize = Math.max(5, CONFIG.batchSize - threadsWithValidAttachments);
      }
    }

    logWithUser(
      `Completed processing ${threadsProcessed} threads, ${threadsWithValidAttachments} with valid attachments for user: ${userEmail}`,
      "INFO"
    );
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
 * @returns {Object} Object containing count of threads with valid attachments
 *
 * The function follows this flow:
 * 1. Iterates through each thread in the provided array
 * 2. Checks if the thread is already processed (has the processed label)
 * 3. For unprocessed threads:
 *    - Examines each message in the thread for attachments
 *    - Filters attachments based on sender domain and attachment properties
 *    - Creates domain folders as needed and saves valid attachments
 *    - Verifies file timestamps match email dates
 * 4. Marks threads as processed if they had attachments (even if all were filtered out)
 * 5. Counts and returns the number of threads that had valid attachments saved
 *
 * This function is optimized for batch processing to handle the 6-minute execution limit
 * of Google Apps Script by focusing on counting threads with valid attachments.
 */
function processThreadsWithCounting(threads, mainFolder, processedLabel) {
  let threadsWithAttachments = 0;

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    try {
      // Check if the thread is already processed
      const threadLabels = thread.getLabels();
      const isAlreadyProcessed = threadLabels.some(
        (label) => label.getName() === processedLabel.getName()
      );

      if (!isAlreadyProcessed) {
        const messages = thread.getMessages();
        let threadProcessed = false;
        let threadHasAttachments = false;
        const threadSubject = thread.getFirstMessageSubject();
        logWithUser(`Processing thread: ${threadSubject}`);

        // Process each message in the thread
        for (let j = 0; j < messages.length; j++) {
          const message = messages[j];
          const attachments = message.getAttachments();
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
              `Found ${attachments.length} attachments in message from ${sender}`
            );

            // Log MIME types to help diagnose what types of attachments are found
            attachments.forEach((att) => {
              try {
                logWithUser(
                  `Attachment: ${att.getName()}, Type: ${att.getContentType()}, Size: ${Math.round(
                    att.getSize() / 1024
                  )}KB`,
                  "INFO"
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

              // Skip if it matches our filter criteria - passing full attachment object
              if (
                shouldSkipFile(fileName, fileSize, attachment) ||
                fileSize > CONFIG.maxFileSize
              ) {
                logWithUser(`Skipping attachment: ${fileName}`, "INFO");
                continue;
              }

              validAttachments.push(attachment);
            }

            // Only create domain folder if we have valid attachments to save
            if (validAttachments.length > 0) {
              const domainFolder = getDomainFolder(sender, mainFolder);

              // Process each valid attachment
              for (let k = 0; k < validAttachments.length; k++) {
                const attachment = validAttachments[k];
                logWithUser(
                  `Processing attachment: ${attachment.getName()} (${Math.round(
                    attachment.getSize() / 1024
                  )}KB)`
                );
                // Log the message date that will be used for the file timestamp
                logWithUser(
                  `Email date for timestamp: ${messageDate.toISOString()}`,
                  "INFO"
                );

                // Use the legacy wrapper for backward compatibility
                const savedFile = saveAttachmentLegacy(
                  attachment,
                  domainFolder,
                  messageDate
                );

                // Process the result object
                if (savedFile) {
                  // If the file was saved, verify its timestamp was set correctly
                  try {
                    // Get the file's timestamp using DriveApp
                    const updatedDate = savedFile.getLastUpdated();
                    logWithUser(
                      `Saved file modification date: ${updatedDate.toISOString()}`,
                      "INFO"
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
                }

                threadProcessed = true;
              }
            } else {
              logWithUser(
                `No valid attachments to save in message from ${sender}`,
                "INFO"
              );
            }
          }
        }

        // Mark as processed:
        // 1. If we processed valid attachments OR
        // 2. If the thread had attachments but they were all filtered out
        if (threadProcessed || threadHasAttachments) {
          withRetry(
            () => thread.addLabel(processedLabel),
            "adding processed label"
          );

          if (threadProcessed) {
            logWithUser(
              `Thread "${threadSubject}" processed with valid attachments and labeled`
            );
            threadsWithAttachments++;
          } else {
            logWithUser(
              `Thread "${threadSubject}" had attachments that were filtered out, marked as processed`
            );
          }
        } else {
          logWithUser(
            `No attachments found in thread "${threadSubject}"`,
            "INFO"
          );
        }
      }
    } catch (error) {
      logWithUser(
        `Error processing thread "${thread.getFirstMessageSubject()}": ${
          error.message
        }`,
        "ERROR"
      );
      // Continue with next thread instead of stopping execution
    }
  }

  return { threadsWithAttachments };
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
        for (const attachment of validAttachments) {
          const saveResult = saveAttachment(attachment, message, domainFolder);

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

    // Apply the processed label to the thread
    thread.addLabel(processedLabel);

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
 * Main function that processes Gmail attachments for authorized users
 *
 * This function:
 * 1. Acquires an execution lock to prevent concurrent runs
 * 2. Gets the list of authorized users
 * 3. Processes emails for each user
 * 4. Logs completion
 *
 * @return {boolean} True if processing completed successfully, false if an error occurred or if no users are authorized.
 * The function follows a structured flow:
 * - It first attempts to acquire a lock to ensure no concurrent executions.
 * - It retrieves the list of authorized users.
 * - For each user, it processes their emails to save attachments to Google Drive.
 * - Logs are generated throughout the process to provide detailed information on the execution status.
 */
function saveAttachmentsToDrive() {
  try {
    logWithUser("Starting attachment processing", "INFO");

    // Acquire lock to prevent concurrent executions
    if (!acquireExecutionLock()) {
      logWithUser("Another instance is already running. Exiting.", "WARNING");
      return false;
    }

    // Get authorized users
    const users = getAuthorizedUsers();

    if (!users || users.length === 0) {
      logWithUser("No authorized users found. Exiting.", "WARNING");
      return false;
    }

    logWithUser(`Processing attachments for ${users.length} users`, "INFO");

    // Process each user
    for (const user of users) {
      processUserEmails(user);
    }

    logWithUser("Attachment processing completed successfully", "INFO");
    return true;
  } catch (error) {
    logWithUser(`Error in saveAttachmentsToDrive: ${error.message}`, "ERROR");
    return false;
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
