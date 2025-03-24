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
 * - AI-based invoice detection using Google Gemini or OpenAI
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

  // Execution settings
  // RECOMMENDED VALUES:
  // - triggerIntervalMinutes: 10-30 for normal use, 5 for high-volume environments
  // - batchSize: 5-10 for most environments, lower for complex processing
  // - executionLockTime: Should be at least 2x the expected execution time
  // INTERDEPENDENCY: batchSize directly affects execution time - higher values may hit the 6-minute limit
  triggerIntervalMinutes: 10, // How often the script runs (in minutes)
  batchSize: 10, // Number of threads to process per execution (prevents hitting 6-minute limit)
  executionLockTime: 10, // Maximum time in minutes to wait for lock release

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
  // INVOICE DETECTION - Settings for identifying and organizing invoices
  //=============================================================================

  // Main invoice detection settings
  // RECOMMENDED VALUES:
  // - invoiceDetection: "gemini" for best privacy and accuracy, "openai" as alternative, false to disable
  // - invoicesFolderName: Using a prefix like "aaa_" ensures the folder appears at the top
  // INTERDEPENDENCY: When invoiceDetection is enabled, detected invoices are saved to invoicesFolderName
  invoiceDetection: "gemini", // AI provider to use: "gemini" (recommended), "openai", or false to disable
  invoicesFolderName: "aaa_Facturas", // Special folder for invoices (prefix ensures it appears at top)

  // Invoice file types and keywords
  // INTERDEPENDENCY: invoiceFileTypes is used by onlyAnalyzePDFs in SHARED AI SETTINGS
  // INTERDEPENDENCY: invoiceKeywords is used by fallbackToKeywords in SHARED AI SETTINGS
  // RECOMMENDED VALUES:
  // - Add common invoice-related terms in your language(s)
  // - Include variations of terms (e.g., "factura", "facturación")
  invoiceFileTypes: [".pdf"], // File extensions considered as potential invoices
  invoiceKeywords: ["factura", "invoice", "receipt", "recibo", "pago"], // Keywords for basic detection

  //=============================================================================
  // GEMINI AI SETTINGS - Configuration for Google's Gemini AI (recommended)
  //=============================================================================

  // API configuration
  // INTERDEPENDENCY: Only used when invoiceDetection = "gemini"
  // The API key can be set in three ways:
  // 1. Directly in this file (not recommended for security)
  // 2. Through environment variables during build
  // 3. Stored in Script Properties using the property name below
  geminiApiKey: "__GEMINI_API_KEY__", // Will be replaced during build process
  geminiApiKeyPropertyName: "gemini_api_key", // Property name for storing the key in Script Properties

  // Model settings
  // RECOMMENDED VALUES:
  // - geminiModel: "gemini-2.0-flash" for fastest response, "gemini-pro" for older deployments
  // - geminiMaxTokens: 10 is sufficient since we only need a confidence score
  // - geminiTemperature: 0.05-0.1 for consistent responses, higher values introduce more variability
  geminiModel: "gemini-2.0-flash", // Gemini model to use
  geminiMaxTokens: 10, // Maximum tokens for response (low since we only need a confidence score)
  geminiTemperature: 0.05, // Very low temperature for consistent, conservative responses

  //=============================================================================
  // OPENAI SETTINGS - Configuration for OpenAI (alternative to Gemini)
  //=============================================================================

  // API configuration
  // INTERDEPENDENCY: Only used when invoiceDetection = "openai"
  // The API key can be set in three ways:
  // 1. Directly in this file (not recommended for security)
  // 2. Through environment variables during build
  // 3. Stored in Script Properties using the property name below
  openAIApiKey: "__OPENAI_API_KEY__", // Will be replaced during build process
  openAIApiKeyPropertyName: "openai_api_key", // Property name for storing the key in Script Properties

  // Model settings
  // RECOMMENDED VALUES:
  // - openAIModel: "gpt-3.5-turbo" offers good balance of performance and cost
  // - openAIMaxTokens: 100 is more than needed but provides flexibility
  // - openAITemperature: 0.05-0.1 for consistent responses
  openAIModel: "gpt-3.5-turbo", // OpenAI model to use
  openAIMaxTokens: 100, // Maximum tokens for response
  openAITemperature: 0.05, // Very low temperature for consistent, conservative responses

  //=============================================================================
  // SHARED AI SETTINGS - Settings that apply to both Gemini and OpenAI
  //=============================================================================

  // Domain exclusions for AI processing
  // Emails from these domains will skip AI analysis and use keyword detection instead
  skipAIForDomains: ["newsletter.com", "marketing.com"],

  // PDF-only option - only analyze emails with PDF attachments
  // This reduces unnecessary API calls and improves privacy
  // INTERDEPENDENCY: When onlyAnalyzePDFs=true, only emails with PDF attachments are sent to AI
  // INTERDEPENDENCY: When strictPdfCheck=true, both file extension and MIME type must match for PDFs
  onlyAnalyzePDFs: true, // Only send emails with PDF attachments to AI
  strictPdfCheck: true, // Check both file extension and MIME type for PDFs (more secure)

  // Fallback and confidence settings
  // INTERDEPENDENCY: When fallbackToKeywords=true, keyword detection is used if AI fails
  // INTERDEPENDENCY: This works with invoiceKeywords setting in the INVOICE DETECTION section
  fallbackToKeywords: true, // Use keyword detection if AI fails or is unavailable

  // AI confidence threshold
  // RECOMMENDED VALUES:
  // - 0.9+ for environments where false positives are costly (current setting)
  // - 0.7-0.8 for balanced precision/recall
  // - 0.5-0.6 for maximum detection (more false positives)
  aiConfidenceThreshold: 0.9, // Threshold (0.0-1.0) - higher values reduce false positives

  //=============================================================================
  // HISTORICAL PATTERN ANALYSIS - Learn from previously identified invoices
  //=============================================================================

  // Settings for analyzing patterns in previously identified invoices
  // INTERDEPENDENCY: useHistoricalPatterns requires manuallyLabeledInvoicesLabel to be set
  // INTERDEPENDENCY: This feature works with the AI invoice detection to improve accuracy
  // RECOMMENDED VALUES:
  // - maxHistoricalEmails: 8-12 for balanced analysis, more emails provide better patterns
  //   but increase processing time
  manuallyLabeledInvoicesLabel: "tickets/facturas", // Gmail label for manually identified invoices
  useHistoricalPatterns: true, // Whether to use historical pattern analysis
  maxHistoricalEmails: 12, // Maximum number of historical emails to analyze per sender

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

  // No additional technical settings at this time
};

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

//=============================================================================
// HISTORICALPATTERNS - HISTORICAL PATTERN ANALYSIS
//=============================================================================

/**
 * Gets historical invoice patterns from emails with the same sender
 * that have been manually labeled with the configured label
 *
 * @param {string} senderEmail - Email address of the sender
 * @returns {Object|null} Patterns detected in historical invoices, or null if disabled
 */
function getHistoricalInvoicePatterns(senderEmail) {
  // Only proceed if feature is enabled and label is configured
  if (!CONFIG.useHistoricalPatterns || !CONFIG.manuallyLabeledInvoicesLabel) {
    return null;
  }

  try {
    // Use the full sender email for more accurate internal Gmail search
    const searchQuery = `from:(${senderEmail}) label:${CONFIG.manuallyLabeledInvoicesLabel}`;
    const threads = GmailApp.search(searchQuery, 0, CONFIG.maxHistoricalEmails);

    logWithUser(
      `Found ${threads.length} historical emails from ${senderEmail} with label "${CONFIG.manuallyLabeledInvoicesLabel}"`,
      "INFO"
    );

    if (threads.length === 0) {
      return null;
    }

    // Collect metadata from these historical invoices
    const subjects = [];
    const dates = [];

    for (const thread of threads) {
      const messages = thread.getMessages();
      if (messages.length > 0) {
        const message = messages[0]; // Get the first message in each thread
        subjects.push(message.getSubject());
        dates.push(message.getDate());
      }
    }

    // Extract patterns
    const patterns = {
      count: threads.length,
      subjectPatterns: extractSubjectPatterns(subjects),
      datePatterns: extractDatePatterns(dates),
      // For internal use only (not sent to AI)
      rawSubjects: subjects,
      rawDates: dates.map((d) => d.toISOString()),
    };

    return patterns;
  } catch (error) {
    logWithUser(
      `Error getting historical invoice patterns: ${error.message}`,
      "ERROR"
    );
    return null;
  }
}

/**
 * Extracts patterns from email subjects
 *
 * @param {string[]} subjects - Array of email subjects
 * @returns {Object} Patterns found in subjects
 */
function extractSubjectPatterns(subjects) {
  if (!subjects || subjects.length === 0) {
    return {};
  }

  try {
    // Find common prefixes
    let commonPrefix = subjects[0];
    for (let i = 1; i < subjects.length; i++) {
      let j = 0;
      while (
        j < commonPrefix.length &&
        j < subjects[i].length &&
        commonPrefix.charAt(j) === subjects[i].charAt(j)
      ) {
        j++;
      }
      commonPrefix = commonPrefix.substring(0, j);
    }

    // Find common suffixes
    let commonSuffix = subjects[0];
    for (let i = 1; i < subjects.length; i++) {
      let j = 0;
      while (
        j < commonSuffix.length &&
        j < subjects[i].length &&
        commonSuffix.charAt(commonSuffix.length - 1 - j) ===
          subjects[i].charAt(subjects[i].length - 1 - j)
      ) {
        j++;
      }
      commonSuffix = commonSuffix.substring(commonSuffix.length - j);
    }

    // Check for invoice keywords
    const containsInvoiceTerms = subjects.some((subject) =>
      CONFIG.invoiceKeywords.some((keyword) =>
        subject.toLowerCase().includes(keyword.toLowerCase())
      )
    );

    // Check for numeric patterns (like invoice numbers)
    const numericPatterns = [];
    const numericRegex = /\d+/g;

    subjects.forEach((subject) => {
      const matches = subject.match(numericRegex);
      if (matches) {
        numericPatterns.push(...matches);
      }
    });

    return {
      commonPrefix: commonPrefix.length > 3 ? commonPrefix.trim() : "",
      commonSuffix: commonSuffix.length > 3 ? commonSuffix.trim() : "",
      containsInvoiceTerms: containsInvoiceTerms,
      hasNumericPatterns: numericPatterns.length > 0,
    };
  } catch (error) {
    logWithUser(`Error extracting subject patterns: ${error.message}`, "ERROR");
    return {};
  }
}

/**
 * Extracts patterns from email dates
 *
 * @param {Date[]} dates - Array of email dates
 * @returns {Object} Patterns found in dates
 */
function extractDatePatterns(dates) {
  if (!dates || dates.length < 2) {
    return { frequency: "unknown" };
  }

  try {
    // Sort dates chronologically
    dates.sort((a, b) => a - b);

    // Calculate intervals between dates in days
    const intervals = [];
    for (let i = 1; i < dates.length; i++) {
      const diffTime = Math.abs(dates[i] - dates[i - 1]);
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      intervals.push(diffDays);
    }

    // Analyze frequency
    const avgInterval =
      intervals.reduce((sum, days) => sum + days, 0) / intervals.length;

    let frequency = "irregular";

    if (avgInterval >= 25 && avgInterval <= 35) {
      frequency = "monthly";
    } else if (avgInterval >= 85 && avgInterval <= 95) {
      frequency = "quarterly";
    } else if (avgInterval >= 175 && avgInterval <= 190) {
      frequency = "biannual";
    } else if (avgInterval >= 350 && avgInterval <= 380) {
      frequency = "annual";
    } else if (avgInterval >= 13 && avgInterval <= 16) {
      frequency = "biweekly";
    } else if (avgInterval >= 6 && avgInterval <= 8) {
      frequency = "weekly";
    }

    // Check if dates fall on same day of month
    const daysOfMonth = dates.map((d) => d.getDate());
    const uniqueDaysOfMonth = [...new Set(daysOfMonth)];
    const sameDayOfMonth = uniqueDaysOfMonth.length === 1;

    return {
      frequency: frequency,
      averageIntervalDays: Math.round(avgInterval),
      sameDayOfMonth: sameDayOfMonth,
      dayOfMonth: sameDayOfMonth ? daysOfMonth[0] : null,
    };
  } catch (error) {
    logWithUser(`Error extracting date patterns: ${error.message}`, "ERROR");
    return { frequency: "unknown" };
  }
}

/**
 * Formats historical patterns into a human-readable description
 * for inclusion in AI prompts
 *
 * @param {Object} patterns - The patterns object from getHistoricalInvoicePatterns
 * @returns {string} Human-readable description of patterns
 */
function formatHistoricalPatternsForPrompt(patterns) {
  if (!patterns || patterns.count === 0) {
    return "";
  }

  try {
    let description = `\nHistorical context: The sender domain has ${patterns.count} previous emails that were manually labeled as invoices.\n\n`;

    // Subject patterns
    if (patterns.subjectPatterns) {
      description += "Subject patterns:\n";

      if (patterns.subjectPatterns.commonPrefix) {
        description += `- Subjects often start with: "${patterns.subjectPatterns.commonPrefix}"\n`;
      }

      if (patterns.subjectPatterns.commonSuffix) {
        description += `- Subjects often end with: "${patterns.subjectPatterns.commonSuffix}"\n`;
      }

      if (patterns.subjectPatterns.containsInvoiceTerms) {
        description += "- Subjects typically contain invoice-related terms\n";
      }

      if (patterns.subjectPatterns.hasNumericPatterns) {
        description +=
          "- Subjects often contain numeric patterns (like invoice numbers)\n";
      }
    }

    // Date patterns
    if (patterns.datePatterns) {
      description += "\nTiming patterns:\n";

      if (
        patterns.datePatterns.frequency !== "unknown" &&
        patterns.datePatterns.frequency !== "irregular"
      ) {
        description += `- Emails are typically sent ${patterns.datePatterns.frequency}\n`;
      }

      if (patterns.datePatterns.sameDayOfMonth) {
        description += `- Emails are typically sent on day ${patterns.datePatterns.dayOfMonth} of the month\n`;
      }

      if (patterns.datePatterns.frequency === "irregular") {
        description += `- Average interval between emails: ${patterns.datePatterns.averageIntervalDays} days\n`;
      }
    }

    return description;
  } catch (error) {
    logWithUser(
      `Error formatting historical patterns: ${error.message}`,
      "ERROR"
    );
    return "";
  }
}

// Make functions available to other modules
var HistoricalPatterns = {
  getHistoricalInvoicePatterns: getHistoricalInvoicePatterns,
  formatHistoricalPatternsForPrompt: formatHistoricalPatternsForPrompt,
};

//=============================================================================
// OPENAIDETECTION - OPENAI DETECTION
//=============================================================================

/**
 * Makes an API call to OpenAI
 *
 * @param {string} content - The content to analyze
 * @param {Object} options - Additional options for the API call
 * @returns {Object} The API response
 */
function callOpenAIAPI(content, options = {}) {
  try {
    // Get the API key
    const apiKey = getOpenAIApiKey();
    if (!apiKey) {
      throw new Error("OpenAI API key not found");
    }

    // Default options
    const defaultOptions = {
      model: CONFIG.openAIModel,
      temperature: CONFIG.openAITemperature,
      max_tokens: CONFIG.openAIMaxTokens,
    };

    // Merge default options with provided options
    const finalOptions = { ...defaultOptions, ...options };

    // Prepare the request payload
    const payload = {
      model: finalOptions.model,
      messages: [
        {
          role: "system",
          content:
            "You are an assistant that analyzes emails to determine if they contain invoices or bills. Be very precise and conservative in your analysis - only respond with 'yes' if you are highly confident the email is specifically about an invoice, bill, or receipt that requires payment.\n\nCheck the content in both English and Spanish languages.\n\nAn invoice typically contains:\n- A clear request for payment\n- An invoice number or reference\n- A specific amount to be paid\n- Payment instructions or terms\n\nJust mentioning words like 'invoice', 'bill', 'receipt', 'factura', 'recibo', or 'pago' is NOT enough to classify as an invoice. The email must be specifically about a payment document.\n\nRespond with 'yes' ONLY if the email is clearly about an actual invoice or bill. Otherwise, respond with 'no'.",
        },
        {
          role: "user",
          content: content,
        },
      ],
      temperature: finalOptions.temperature,
      max_tokens: finalOptions.max_tokens,
    };

    // Make the API request
    const response = UrlFetchApp.fetch(
      "https://api.openai.com/v1/chat/completions",
      {
        method: "post",
        headers: {
          Authorization: "Bearer " + apiKey,
          "Content-Type": "application/json",
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      }
    );

    // Parse and return the response
    const responseData = JSON.parse(response.getContentText());

    // Check for errors in the response
    if (response.getResponseCode() !== 200) {
      throw new Error(
        `API Error: ${responseData.error?.message || "Unknown error"}`
      );
    }

    return responseData;
  } catch (error) {
    logWithUser(`OpenAI API call failed: ${error.message}`, "ERROR");
    throw error;
  }
}

/**
 * Analyzes an email to determine if it contains an invoice
 *
 * @param {GmailMessage} message - The Gmail message to analyze
 * @returns {boolean} True if the message likely contains an invoice
 */
function analyzeEmail(message) {
  try {
    // Get the content to analyze
    const content = getTextContentToAnalyze(message);

    // Format the prompt
    const formattedContent = formatPrompt(content);

    // Call the OpenAI API
    const response = callOpenAIAPI(formattedContent);

    // Parse the response
    return parseResponse(response);
  } catch (error) {
    logWithUser(`Email analysis failed: ${error.message}`, "ERROR");
    throw error;
  }
}

/**
 * Main function to determine if a message contains an invoice using OpenAI
 *
 * @param {GmailMessage} message - The Gmail message to analyze
 * @returns {boolean} True if the message likely contains an invoice
 */
function isInvoiceWithOpenAI(message) {
  try {
    // Check if OpenAI is enabled
    if (CONFIG.invoiceDetection !== "openai") {
      logWithUser(
        "OpenAI invoice detection is disabled in configuration",
        "INFO"
      );
      return false;
    }

    // Log the message subject for debugging
    const subject = message.getSubject() || "(no subject)";
    logWithUser(
      `Analyzing message with subject: "${subject}" using OpenAI`,
      "INFO"
    );

    // Check if sender domain should be skipped
    const sender = message.getFrom();
    const domain = extractDomain(sender);
    logWithUser(`Message sender: ${sender}, domain: ${domain}`, "INFO");

    if (CONFIG.skipAIForDomains && CONFIG.skipAIForDomains.includes(domain)) {
      logWithUser(
        `Skipping OpenAI for domain ${domain} (in skipAIForDomains list)`,
        "INFO"
      );
      return false;
    }

    // Check for PDF attachments if configured
    if (CONFIG.onlyAnalyzePDFs) {
      const attachments = message.getAttachments();
      logWithUser(`Message has ${attachments.length} attachments`, "INFO");

      let hasPDF = false;
      for (const attachment of attachments) {
        const fileName = attachment.getName().toLowerCase();
        const contentType = attachment.getContentType().toLowerCase();
        logWithUser(
          `Checking attachment: ${fileName}, type: ${contentType}`,
          "INFO"
        );

        if (CONFIG.strictPdfCheck) {
          if (fileName.endsWith(".pdf") && contentType.includes("pdf")) {
            hasPDF = true;
            logWithUser(`Found PDF attachment: ${fileName}`, "INFO");
            break;
          }
        } else if (fileName.endsWith(".pdf")) {
          hasPDF = true;
          logWithUser(`Found PDF attachment: ${fileName}`, "INFO");
          break;
        }
      }

      if (!hasPDF) {
        logWithUser(
          `No PDF attachments found, skipping OpenAI analysis`,
          "INFO"
        );
        return CONFIG.fallbackToKeywords ? checkKeywords(message) : false;
      }
    }

    // Analyze the email
    logWithUser(`Starting OpenAI analysis for message: "${subject}"`, "INFO");
    const isInvoice = analyzeEmail(message);
    logWithUser(`OpenAI invoice detection result: ${isInvoice}`, "INFO");

    return isInvoice;
  } catch (error) {
    logWithUser(`OpenAI invoice detection failed: ${error.message}`, "ERROR");
    // Return false on error, caller can decide to fall back to keywords
    return false;
  }
}

/**
 * Extracts relevant text content from a Gmail message for analysis
 *
 * @param {GmailMessage} message - The Gmail message to extract content from
 * @returns {Object} Object containing subject, body, and historical patterns
 */
function getTextContentToAnalyze(message) {
  try {
    const subject = message.getSubject() || "";
    let body = "";

    // Try to get plain text body
    try {
      body = message.getPlainBody() || "";

      // Truncate body if it's too long (to save tokens)
      if (body.length > 1500) {
        body = body.substring(0, 1500) + "...";
      }
    } catch (e) {
      logWithUser(`Could not get message body: ${e.message}`, "WARNING");
    }

    // Get historical patterns if enabled
    let historicalPatterns = null;
    if (CONFIG.useHistoricalPatterns && CONFIG.manuallyLabeledInvoicesLabel) {
      historicalPatterns = HistoricalPatterns.getHistoricalInvoicePatterns(
        message.getFrom()
      );
    }

    return {
      subject: subject,
      body: body,
      sender: message.getFrom() || "",
      date: message.getDate().toISOString(),
      historicalPatterns: historicalPatterns,
    };
  } catch (error) {
    logWithUser(`Error extracting message content: ${error.message}`, "ERROR");
    // Return minimal content on error
    return {
      subject: message.getSubject() || "",
      body: "",
      sender: message.getFrom() || "",
      date: message.getDate().toISOString(),
    };
  }
}

/**
 * Formats the email content into a prompt for the OpenAI API
 *
 * @param {Object} content - The email content to format
 * @returns {string} Formatted prompt
 */
function formatPrompt(content) {
  try {
    // Get historical patterns if available
    let historicalContext = "";
    if (content.historicalPatterns && content.historicalPatterns.count > 0) {
      historicalContext = HistoricalPatterns.formatHistoricalPatternsForPrompt(
        content.historicalPatterns
      );
    }

    return `
Please analyze this email and determine if it contains an invoice or bill.
Respond with only 'yes' or 'no'.
${historicalContext}
From: ${content.sender}
Date: ${content.date}
Subject: ${content.subject}

${content.body}
`;
  } catch (error) {
    logWithUser(`Error formatting prompt: ${error.message}`, "ERROR");
    // Return a simplified prompt on error
    return `Subject: ${content.subject}\n\nIs this an invoice or bill? Answer yes or no.`;
  }
}

/**
 * Parses the OpenAI API response to determine if the message contains an invoice
 *
 * @param {Object} response - The API response to parse
 * @returns {boolean} True if the message likely contains an invoice
 */
function parseResponse(response) {
  try {
    // Extract the response text
    const responseText = response.choices[0].message.content
      .trim()
      .toLowerCase();

    // Log the raw response for debugging
    logWithUser(`OpenAI raw response: "${responseText}"`, "INFO");

    // Check if the response indicates an invoice
    // We're looking for "yes" or variations like "yes, it is an invoice"
    return responseText.includes("yes");
  } catch (error) {
    logWithUser(`Error parsing API response: ${error.message}`, "ERROR");
    throw error;
  }
}

/**
 * Verifies if there is an API key available, using first the configuration and then
 * the script properties as a backup
 *
 * @returns {string|null} The API key if found, null otherwise
 */
function getOpenAIApiKey() {
  // First check if the key is in CONFIG and not the placeholder
  if (CONFIG.openAIApiKey && CONFIG.openAIApiKey !== "__OPENAI_API_KEY__") {
    return CONFIG.openAIApiKey;
  }

  // Then try to get it from script properties
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty(
      CONFIG.openAIApiKeyPropertyName
    );
    return apiKey;
  } catch (error) {
    logWithUser(`Error retrieving API key: ${error.message}`, "ERROR");
    return null;
  }
}

/**
 * Stores the OpenAI API key securely in Script Properties
 *
 * @param {string} apiKey - The OpenAI API key to store
 * @returns {boolean} True if successful, false otherwise
 */
function storeOpenAIApiKey(apiKey) {
  try {
    PropertiesService.getScriptProperties().setProperty(
      CONFIG.openAIApiKeyPropertyName,
      apiKey
    );
    return true;
  } catch (error) {
    logWithUser(`Error storing API key: ${error.message}`, "ERROR");
    return false;
  }
}

/**
 * Test function to verify OpenAI API connectivity
 * This function can be run directly from the Apps Script editor
 * to check if the configured API key is correct
 */
function testOpenAIConnection() {
  try {
    const testPrompt = "Is this a test?";
    const response = callOpenAIAPI(testPrompt);

    Logger.log(
      `Successfully connected to OpenAI API. Response: ${JSON.stringify(
        response
      )}`
    );
    return {
      success: true,
      response: response,
    };
  } catch (e) {
    Logger.log(`Error connecting to OpenAI API: ${e.message}`);
    return {
      success: false,
      error: e.message,
    };
  }
}

// Make functions available to other modules
var OpenAIDetection = {
  isInvoiceWithOpenAI: isInvoiceWithOpenAI,
  testOpenAIConnection: testOpenAIConnection,
  storeOpenAIApiKey: storeOpenAIApiKey,
};

//=============================================================================
// GEMINIDETECTION - GEMINI DETECTION
//=============================================================================

/**
 * Makes an API call to Google Gemini
 *
 * @param {string} prompt - The prompt to send to Gemini
 * @param {Object} options - Additional options for the API call
 * @returns {Object} The API response
 */
function callGeminiAPI(prompt, options = {}) {
  try {
    // Get the API key
    const apiKey = getGeminiApiKey();
    if (!apiKey) {
      throw new Error("Gemini API key not found");
    }

    // Default options
    const defaultOptions = {
      model: CONFIG.geminiModel,
      temperature: CONFIG.geminiTemperature,
      maxOutputTokens: CONFIG.geminiMaxTokens,
    };

    // Merge default options with provided options
    const finalOptions = { ...defaultOptions, ...options };

    // Prepare the request payload
    const payload = {
      contents: [
        {
          parts: [{ text: prompt }],
        },
      ],
      generationConfig: {
        temperature: finalOptions.temperature,
        maxOutputTokens: finalOptions.maxOutputTokens,
      },
    };

    // Try v1 API first (newer version)
    let url =
      "https://generativelanguage.googleapis.com/v1/models/" +
      finalOptions.model +
      ":generateContent";

    try {
      // Make the API request to v1 endpoint
      const response = UrlFetchApp.fetch(`${url}?key=${apiKey}`, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      });

      // Parse the response
      const responseData = JSON.parse(response.getContentText());

      // Check for errors in the response
      if (response.getResponseCode() !== 200) {
        throw new Error(
          `API Error: ${responseData.error?.message || "Unknown error"}`
        );
      }

      // Extract the text response
      const responseText = responseData.candidates[0].content.parts[0].text;
      logWithUser("Successfully used Gemini API v1 endpoint", "INFO");
      return responseText;
    } catch (v1Error) {
      // Log the error but don't throw yet
      logWithUser(
        `Gemini API v1 endpoint failed: ${v1Error.message}, trying v1beta...`,
        "WARNING"
      );

      // Fall back to v1beta API
      url =
        "https://generativelanguage.googleapis.com/v1beta/models/" +
        finalOptions.model +
        ":generateContent";

      // Make the API request to v1beta endpoint
      const response = UrlFetchApp.fetch(`${url}?key=${apiKey}`, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      });

      // Parse the response
      const responseData = JSON.parse(response.getContentText());

      // Check for errors in the response
      if (response.getResponseCode() !== 200) {
        throw new Error(
          `API Error: ${responseData.error?.message || "Unknown error"}`
        );
      }

      // Extract the text response
      const responseText = responseData.candidates[0].content.parts[0].text;
      logWithUser("Successfully used Gemini API v1beta endpoint", "INFO");
      return responseText;
    }
  } catch (error) {
    logWithUser(`Gemini API call failed: ${error.message}`, "ERROR");
    throw error;
  }
}

/**
 * Extracts relevant keywords from email body without sending full content
 *
 * @param {string} body - The email body text
 * @returns {string[]} Array of extracted keywords and patterns
 */
function extractKeywords(body) {
  try {
    // Extract only relevant terms without sending the full body
    const relevantTerms = [];
    CONFIG.invoiceKeywords.forEach((keyword) => {
      if (body.toLowerCase().includes(keyword.toLowerCase())) {
        relevantTerms.push(keyword);
      }
    });

    // Look for numeric patterns that might be amounts or references
    const amountPatterns =
      body.match(/\$\s*\d+[.,]\d{2}|€\s*\d+[.,]\d{2}/g) || [];
    const refPatterns = body.match(/ref\w*\s*:?\s*[A-Z0-9-]+/gi) || [];
    const invoiceNumPatterns = body.match(/inv\w*\s*:?\s*[A-Z0-9-]+/gi) || [];

    return [
      ...relevantTerms,
      ...amountPatterns,
      ...refPatterns,
      ...invoiceNumPatterns,
    ];
  } catch (error) {
    logWithUser(`Error extracting keywords: ${error.message}`, "ERROR");
    return [];
  }
}

/**
 * Extracts metadata from a Gmail message for AI analysis
 * without sending the full content
 *
 * @param {GmailMessage} message - The Gmail message to extract metadata from
 * @returns {Object} Metadata object with privacy-safe information
 */
function extractMetadata(message) {
  try {
    const subject = message.getSubject() || "";
    const body = message.getPlainBody() || "";
    const attachments = message.getAttachments() || [];

    // Get attachment info
    const attachmentInfo = attachments.map((attachment) => {
      return {
        name: attachment.getName(),
        extension: attachment.getName().split(".").pop().toLowerCase(),
        contentType: attachment.getContentType(),
        size: attachment.getSize(),
      };
    });

    // Get historical patterns if enabled
    let historicalPatterns = null;
    if (CONFIG.useHistoricalPatterns && CONFIG.manuallyLabeledInvoicesLabel) {
      historicalPatterns = HistoricalPatterns.getHistoricalInvoicePatterns(
        message.getFrom()
      );
    }

    // Get sender information
    const sender = message.getFrom() || "";
    const senderDomain = extractDomain(sender);

    // Anonymize sender email (only keep domain)
    const anonymizedSender = `user@${senderDomain}`;

    // Create metadata object with only necessary information
    // Avoid sending any personally identifiable information
    return {
      subject: subject,
      senderDomain: senderDomain, // Only send domain, not full email
      date: message.getDate().toISOString(),
      hasAttachments: attachments.length > 0,
      attachmentTypes: attachmentInfo.map((a) => a.extension),
      attachmentContentTypes: attachmentInfo.map((a) => a.contentType),
      keywordsFound: extractKeywords(body),
      historicalPatterns: historicalPatterns,
    };
  } catch (error) {
    logWithUser(`Error extracting message metadata: ${error.message}`, "ERROR");
    // Return minimal metadata on error, still ensuring privacy
    const errorSenderDomain = extractDomain(message.getFrom() || "");
    return {
      subject: message.getSubject() || "",
      senderDomain: errorSenderDomain, // Only include domain, not full email
      hasAttachments: false,
      attachmentTypes: [],
      attachmentContentTypes: [],
      keywordsFound: [],
    };
  }
}

/**
 * Formats the metadata into a prompt for the Gemini API
 *
 * @param {Object} metadata - The email metadata to format
 * @returns {string} Formatted prompt
 */
function formatPrompt(metadata) {
  try {
    // Get historical patterns if available
    let historicalContext = "";
    if (metadata.historicalPatterns && metadata.historicalPatterns.count > 0) {
      historicalContext = HistoricalPatterns.formatHistoricalPatternsForPrompt(
        metadata.historicalPatterns
      );
    }

    return `
Based on these email metadata, assess the likelihood that this contains an invoice.
You don't have access to the full content for privacy reasons.
${historicalContext}
Metadata: ${JSON.stringify(metadata, null, 2)}

An invoice typically contains:
- A clear request for payment
- An invoice number or reference
- A specific amount to be paid
- Payment instructions or terms

Just mentioning words like 'invoice', 'bill', 'receipt', etc. is NOT enough to classify as an invoice.
The email must be specifically about a payment document.

On a scale from 0.0 to 1.0, where:
- 0.0 means definitely NOT an invoice
- 1.0 means definitely IS an invoice

Provide ONLY a single number between 0.0 and 1.0 representing your confidence.
Example responses: "0.2", "0.85", "0.99"
`;
  } catch (error) {
    logWithUser(`Error formatting prompt: ${error.message}`, "ERROR");
    // Return a simplified prompt on error
    return `Based on this subject: "${metadata.subject}", on a scale from 0.0 to 1.0, what's the likelihood this is an invoice? Respond with only a number.`;
  }
}

/**
 * Parses the Gemini API response to extract confidence score
 *
 * @param {string} response - The API response to parse
 * @returns {number} Confidence score between 0 and 1
 */
function parseGeminiResponse(response) {
  try {
    // Extract the text response and clean it
    const responseText = response.trim();

    // Try to extract a number from the response
    const confidenceMatch = responseText.match(/(\d+\.\d+|\d+)/);

    if (confidenceMatch) {
      const confidence = parseFloat(confidenceMatch[0]);

      // Validate that it's a number between 0 and 1
      if (!isNaN(confidence) && confidence >= 0 && confidence <= 1) {
        logWithUser(`Gemini confidence score: ${confidence}`, "INFO");
        return confidence;
      }
    }

    // If we couldn't extract a valid confidence score, log warning and return 0
    logWithUser(
      `Could not extract valid confidence score from response: "${responseText}"`,
      "WARNING"
    );
    return 0;
  } catch (error) {
    logWithUser(`Error parsing Gemini response: ${error.message}`, "ERROR");
    return 0;
  }
}

/**
 * Analyzes an email to determine if it contains an invoice
 *
 * @param {GmailMessage} message - The Gmail message to analyze
 * @returns {number} Confidence score between 0 and 1
 */
function analyzeEmail(message) {
  try {
    // Extract metadata (not full content)
    const metadata = extractMetadata(message);

    // Format the prompt
    const formattedPrompt = formatPrompt(metadata);

    // Log the formatted prompt
    logWithUser(
      `Formatted prompt with metadata for Gemini: ${formattedPrompt}`,
      "DEBUG"
    );

    // Call the Gemini API
    const response = callGeminiAPI(formattedPrompt);

    // Parse the response to get confidence score
    return parseGeminiResponse(response);
  } catch (error) {
    logWithUser(`Email analysis failed: ${error.message}`, "ERROR");
    return 0; // Return 0 confidence on error
  }
}

/**
 * Main function to determine if a message contains an invoice using Gemini
 *
 * @param {GmailMessage} message - The Gmail message to analyze
 * @returns {boolean} True if the message likely contains an invoice
 */
function isInvoiceWithGemini(message) {
  try {
    // Check if Gemini is enabled
    if (CONFIG.invoiceDetection !== "gemini") {
      logWithUser(
        "Gemini invoice detection is disabled in configuration",
        "INFO"
      );
      return false;
    }

    // Log the message subject for debugging
    const subject = message.getSubject() || "(no subject)";
    logWithUser(
      `Analyzing message with subject: "${subject}" using Gemini`,
      "INFO"
    );

    // Check if sender domain should be skipped
    const sender = message.getFrom();
    const domain = extractDomain(sender);

    // Log only the domain, not the full email address
    logWithUser(`Message domain: ${domain}`, "INFO");

    if (CONFIG.skipAIForDomains && CONFIG.skipAIForDomains.includes(domain)) {
      logWithUser(
        `Skipping Gemini for domain ${domain} (in skipAIForDomains list)`,
        "INFO"
      );
      return false;
    }

    // Check for PDF attachments if configured
    if (CONFIG.onlyAnalyzePDFs) {
      const attachments = message.getAttachments();
      logWithUser(`Message has ${attachments.length} attachments`, "INFO");

      let hasPDF = false;
      for (const attachment of attachments) {
        const fileName = attachment.getName().toLowerCase();
        const contentType = attachment.getContentType().toLowerCase();
        logWithUser(
          `Checking attachment: ${fileName}, type: ${contentType}`,
          "INFO"
        );

        if (CONFIG.strictPdfCheck) {
          if (fileName.endsWith(".pdf") && contentType.includes("pdf")) {
            hasPDF = true;
            logWithUser(`Found PDF attachment: ${fileName}`, "INFO");
            break;
          }
        } else if (fileName.endsWith(".pdf")) {
          hasPDF = true;
          logWithUser(`Found PDF attachment: ${fileName}`, "INFO");
          break;
        }
      }

      if (!hasPDF) {
        logWithUser(
          `No PDF attachments found, skipping Gemini analysis`,
          "INFO"
        );
        return CONFIG.fallbackToKeywords ? checkKeywords(message) : false;
      }
    }

    // Analyze the email to get confidence score
    logWithUser(`Starting Gemini analysis for message: "${subject}"`, "INFO");
    const confidence = analyzeEmail(message);

    // Compare with threshold
    const isInvoice = confidence >= CONFIG.aiConfidenceThreshold;
    logWithUser(
      `Gemini invoice detection result: ${isInvoice} (confidence: ${confidence}, threshold: ${CONFIG.aiConfidenceThreshold})`,
      "INFO"
    );

    return isInvoice;
  } catch (error) {
    logWithUser(`Gemini invoice detection failed: ${error.message}`, "ERROR");
    // Return false on error, caller can decide to fall back to keywords
    return false;
  }
}

/**
 * Verifies if there is an API key available, using first the configuration and then
 * the script properties as a backup
 *
 * @returns {string|null} The API key if found, null otherwise
 */
function getGeminiApiKey() {
  // First check if the key is in CONFIG and not the placeholder
  if (CONFIG.geminiApiKey && CONFIG.geminiApiKey !== "__GEMINI_API_KEY__") {
    return CONFIG.geminiApiKey;
  }

  // Then try to get it from script properties
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty(
      CONFIG.geminiApiKeyPropertyName
    );
    return apiKey;
  } catch (error) {
    logWithUser(`Error retrieving Gemini API key: ${error.message}`, "ERROR");
    return null;
  }
}

/**
 * Stores the Gemini API key securely in Script Properties
 *
 * @param {string} apiKey - The Gemini API key to store
 * @returns {boolean} True if successful, false otherwise
 */
function storeGeminiApiKey(apiKey) {
  try {
    PropertiesService.getScriptProperties().setProperty(
      CONFIG.geminiApiKeyPropertyName,
      apiKey
    );
    return true;
  } catch (error) {
    logWithUser(`Error storing Gemini API key: ${error.message}`, "ERROR");
    return false;
  }
}

/**
 * Test function to verify Gemini API connectivity
 * This function can be run directly from the Apps Script editor
 * to check if the configured API key is correct
 */
function testGeminiConnection() {
  try {
    const testPrompt =
      "On a scale from 0.0 to 1.0, how likely is this a test? Respond with only a number.";
    const response = callGeminiAPI(testPrompt);

    Logger.log(`Successfully connected to Gemini API. Response: ${response}`);
    return {
      success: true,
      response: response,
    };
  } catch (e) {
    Logger.log(`Error connecting to Gemini API: ${e.message}`);
    return {
      success: false,
      error: e.message,
    };
  }
}

// Make functions available to other modules
var GeminiDetection = {
  isInvoiceWithGemini: isInvoiceWithGemini,
  testGeminiConnection: testGeminiConnection,
  storeGeminiApiKey: storeGeminiApiKey,
};

//=============================================================================
// VERIFYAPIKEYS - API KEY VERIFICATION
//=============================================================================

/**
 * Verifies if the Gemini API key is valid and working
 *
 * @returns {Object} Object with success status, details, and API information
 */
function verifyGeminiAPIKey() {
  try {
    logWithUser("Starting Gemini API key verification...", "INFO");

    // Step 1: Check if the API key exists in configuration or script properties
    const apiKey = getGeminiApiKey();
    if (!apiKey) {
      return {
        success: false,
        message:
          "Gemini API key not found in configuration or script properties",
        details: {
          configValue:
            CONFIG.geminiApiKey === "__GEMINI_API_KEY__"
              ? "Not set (placeholder)"
              : "Set in CONFIG",
          scriptPropertyName: CONFIG.geminiApiKeyPropertyName,
          scriptPropertyValue: "Not available (null)",
        },
      };
    }

    // Step 2: Mask the API key for logging (show only first 4 and last 4 characters)
    const maskedKey = maskAPIKey(apiKey);
    logWithUser(`Found Gemini API key: ${maskedKey}`, "INFO");

    // Step 3: Test the API connection
    logWithUser("Testing Gemini API connection...", "INFO");
    const testResult = GeminiDetection.testGeminiConnection();

    if (testResult.success) {
      return {
        success: true,
        message: "Gemini API key is valid and working correctly",
        details: {
          apiKeySource:
            CONFIG.geminiApiKey !== "__GEMINI_API_KEY__"
              ? "CONFIG"
              : "Script Properties",
          apiKeyLength: apiKey.length,
          maskedKey: maskedKey,
          model: CONFIG.geminiModel,
          testResponse: testResult.response,
        },
      };
    } else {
      return {
        success: false,
        message: "Gemini API key validation failed",
        details: {
          apiKeySource:
            CONFIG.geminiApiKey !== "__GEMINI_API_KEY__"
              ? "CONFIG"
              : "Script Properties",
          apiKeyLength: apiKey.length,
          maskedKey: maskedKey,
          model: CONFIG.geminiModel,
          error: testResult.error,
        },
      };
    }
  } catch (error) {
    logWithUser(`Error verifying Gemini API key: ${error.message}`, "ERROR");
    return {
      success: false,
      message: `Error verifying Gemini API key: ${error.message}`,
      details: {
        error: error.message,
        stack: error.stack,
      },
    };
  }
}

/**
 * Verifies if the OpenAI API key is valid and working
 *
 * @returns {Object} Object with success status, details, and API information
 */
function verifyOpenAIAPIKey() {
  try {
    logWithUser("Starting OpenAI API key verification...", "INFO");

    // Step 1: Check if the API key exists in configuration or script properties
    const apiKey = getOpenAIApiKey();
    if (!apiKey) {
      return {
        success: false,
        message:
          "OpenAI API key not found in configuration or script properties",
        details: {
          configValue:
            CONFIG.openAIApiKey === "__OPENAI_API_KEY__"
              ? "Not set (placeholder)"
              : "Set in CONFIG",
          scriptPropertyName: CONFIG.openAIApiKeyPropertyName,
          scriptPropertyValue: "Not available (null)",
        },
      };
    }

    // Step 2: Mask the API key for logging (show only first 4 and last 4 characters)
    const maskedKey = maskAPIKey(apiKey);
    logWithUser(`Found OpenAI API key: ${maskedKey}`, "INFO");

    // Step 3: Test the API connection
    logWithUser("Testing OpenAI API connection...", "INFO");
    const testResult = OpenAIDetection.testOpenAIConnection();

    if (testResult.success) {
      return {
        success: true,
        message: "OpenAI API key is valid and working correctly",
        details: {
          apiKeySource:
            CONFIG.openAIApiKey !== "__OPENAI_API_KEY__"
              ? "CONFIG"
              : "Script Properties",
          apiKeyLength: apiKey.length,
          maskedKey: maskedKey,
          model: CONFIG.openAIModel,
          testResponse: "Response available (not shown for brevity)",
        },
      };
    } else {
      return {
        success: false,
        message: "OpenAI API key validation failed",
        details: {
          apiKeySource:
            CONFIG.openAIApiKey !== "__OPENAI_API_KEY__"
              ? "CONFIG"
              : "Script Properties",
          apiKeyLength: apiKey.length,
          maskedKey: maskedKey,
          model: CONFIG.openAIModel,
          error: testResult.error,
        },
      };
    }
  } catch (error) {
    logWithUser(`Error verifying OpenAI API key: ${error.message}`, "ERROR");
    return {
      success: false,
      message: `Error verifying OpenAI API key: ${error.message}`,
      details: {
        error: error.message,
        stack: error.stack,
      },
    };
  }
}

/**
 * Verifies both Gemini and OpenAI API keys
 *
 * @returns {Object} Object with verification results for both APIs
 */
function verifyAllAPIKeys() {
  const results = {
    gemini: verifyGeminiAPIKey(),
    openai: verifyOpenAIAPIKey(),
    timestamp: new Date().toISOString(),
    config: {
      invoiceDetection: CONFIG.invoiceDetection,
      geminiModel: CONFIG.geminiModel,
      openAIModel: CONFIG.openAIModel,
    },
  };

  // Log a summary of the results
  if (results.gemini.success && results.openai.success) {
    logWithUser(
      "✅ Both Gemini and OpenAI API keys are valid and working",
      "INFO"
    );
  } else if (results.gemini.success) {
    logWithUser(
      "⚠️ Gemini API key is valid, but OpenAI API key validation failed",
      "WARNING"
    );
  } else if (results.openai.success) {
    logWithUser(
      "⚠️ OpenAI API key is valid, but Gemini API key validation failed",
      "WARNING"
    );
  } else {
    logWithUser(
      "❌ Both Gemini and OpenAI API key validations failed",
      "ERROR"
    );
  }

  return results;
}

/**
 * Helper function to mask an API key for safe logging
 * Shows only the first 4 and last 4 characters, with the rest masked
 *
 * @param {string} apiKey - The API key to mask
 * @returns {string} The masked API key
 */
function maskAPIKey(apiKey) {
  if (!apiKey || apiKey.length < 8) {
    return "Invalid key (too short)";
  }

  const firstFour = apiKey.substring(0, 4);
  const lastFour = apiKey.substring(apiKey.length - 4);
  const maskedPortion = "*".repeat(Math.min(apiKey.length - 8, 10));

  return `${firstFour}${maskedPortion}${lastFour} (${apiKey.length} chars)`;
}

// Make functions available to other modules
var VerifyAPIKeys = {
  verifyGeminiAPIKey: verifyGeminiAPIKey,
  verifyOpenAIAPIKey: verifyOpenAIAPIKey,
  verifyAllAPIKeys: verifyAllAPIKeys,
};

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
 * Get or create the special folder for invoices
 *
 * This function creates or retrieves a special folder for storing invoices,
 * with subfolders for each sender's domain. It uses the same locking mechanism as
 * getDomainFolder to prevent race conditions.
 *
 * @param {DriveFolder} mainFolder - The main folder to create the invoices folder in
 * @param {string} domain - The sender's domain to create a subfolder for (optional)
 * @returns {DriveFolder} The invoices folder or domain subfolder
 *
 * The function follows this flow:
 * 1. Gets the folder name from CONFIG.invoicesFolderName
 * 2. Acquires a lock to prevent race conditions during folder creation
 * 3. Checks if the invoices folder already exists
 * 4. If the folder exists, returns it immediately if no domain is specified
 * 5. If domain is specified, gets or creates a subfolder for that domain
 * 6. If any errors occur, falls back to using the main folder
 */
function getInvoicesFolder(mainFolder, domain = null) {
  try {
    const folderName = CONFIG.invoicesFolderName;

    // Use a lock to prevent race conditions when creating folders
    const lock = LockService.getScriptLock();
    try {
      lock.tryLock(10000); // Wait up to 10 seconds for the lock

      // First check if the main invoices folder exists
      const folders = withRetry(
        () => mainFolder.getFoldersByName(folderName),
        "getting invoices folder"
      );

      let invoicesFolder;
      if (folders.hasNext()) {
        invoicesFolder = folders.next();
        logWithUser(`Using existing invoices folder: ${folderName}`);
      } else {
        // Double-check that the folder still doesn't exist
        // This helps in cases where another execution created it just now
        const doubleCheckFolders = withRetry(
          () => mainFolder.getFoldersByName(folderName),
          "double-checking invoices folder"
        );

        if (doubleCheckFolders.hasNext()) {
          invoicesFolder = doubleCheckFolders.next();
          logWithUser(
            `Using existing invoices folder (after double-check): ${folderName}`
          );
        } else {
          // If we're still here, we can safely create the folder
          invoicesFolder = withRetry(
            () => mainFolder.createFolder(folderName),
            "creating invoices folder"
          );
          logWithUser(`Created new invoices folder: ${folderName}`);
        }
      }

      // If no domain is specified, return the main invoices folder
      if (!domain) {
        return invoicesFolder;
      }

      // If domain is specified, get or create a subfolder for that domain
      const domainFolders = withRetry(
        () => invoicesFolder.getFoldersByName(domain),
        `getting domain subfolder in invoices folder: ${domain}`
      );

      if (domainFolders.hasNext()) {
        const domainFolder = domainFolders.next();
        logWithUser(`Using existing domain subfolder in invoices: ${domain}`);
        return domainFolder;
      } else {
        // Double-check for the domain subfolder
        const doubleCheckDomainFolders = withRetry(
          () => invoicesFolder.getFoldersByName(domain),
          `double-checking domain subfolder in invoices: ${domain}`
        );

        if (doubleCheckDomainFolders.hasNext()) {
          const domainFolder = doubleCheckDomainFolders.next();
          logWithUser(
            `Using existing domain subfolder in invoices (after double-check): ${domain}`
          );
          return domainFolder;
        }

        // Create the domain subfolder
        const newDomainFolder = withRetry(
          () => invoicesFolder.createFolder(domain),
          `creating domain subfolder in invoices: ${domain}`
        );
        logWithUser(`Created new domain subfolder in invoices: ${domain}`);
        return newDomainFolder;
      }
    } finally {
      // Always release the lock
      if (lock.hasLock()) {
        lock.releaseLock();
      }
    }
  } catch (error) {
    logWithUser(
      `Error getting invoices folder: ${error.message}. Using main folder as fallback.`,
      "ERROR"
    );
    // Ultimate fallback: return the main folder
    return mainFolder;
  }
}

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

//=============================================================================
// GMAILPROCESSING - GMAIL PROCESSING
//=============================================================================

/**
 * Determines if a message appears to contain an invoice using AI or keywords
 *
 * @param {GmailMessage} message - The Gmail message to analyze
 * @returns {boolean} True if the message likely contains an invoice
 */
function isInvoiceMessage(message) {
  // Early return if invoice detection is disabled
  if (CONFIG.invoiceDetection === false) return false;

  try {
    // Check if any AI detection is enabled
    if (
      CONFIG.invoiceDetection === "gemini" ||
      CONFIG.invoiceDetection === "openai"
    ) {
      // Check if message has attachments and if any are PDFs (with stricter checking)
      if (CONFIG.onlyAnalyzePDFs) {
        const attachments = message.getAttachments();
        let hasPDF = false;

        for (const attachment of attachments) {
          const fileName = attachment.getName().toLowerCase();
          const contentType = attachment.getContentType().toLowerCase();

          // Strict PDF check - both extension and MIME type
          if (CONFIG.strictPdfCheck) {
            if (fileName.endsWith(".pdf") && contentType.includes("pdf")) {
              hasPDF = true;
              break;
            }
          } else {
            // Legacy check - just extension
            if (fileName.endsWith(".pdf")) {
              hasPDF = true;
              break;
            }
          }
        }

        // Skip AI if there are no PDF attachments
        if (!hasPDF) {
          logWithUser("No PDF attachments found, skipping AI analysis", "INFO");
          // Don't fall back to keywords when there are no PDFs if onlyAnalyzePDFs is true
          // This prevents non-PDF attachments from being classified as invoices based on keywords
          return false;
        }
      }

      // Check if sender domain should be skipped
      const sender = message.getFrom();
      const domain = extractDomain(sender);
      if (CONFIG.skipAIForDomains && CONFIG.skipAIForDomains.includes(domain)) {
        logWithUser(`Skipping AI for domain ${domain}, using keywords`, "INFO");
        return checkKeywords(message);
      }

      // Try Gemini detection first if enabled
      if (CONFIG.invoiceDetection === "gemini") {
        try {
          const isInvoice = GeminiDetection.isInvoiceWithGemini(message);
          logWithUser(`Gemini invoice detection result: ${isInvoice}`, "INFO");
          return isInvoice;
        } catch (geminiError) {
          logWithUser(
            `Gemini detection error: ${geminiError.message}, trying fallback options`,
            "WARNING"
          );

          // Try OpenAI if enabled as fallback
          if (CONFIG.invoiceDetection === "openai") {
            try {
              const isInvoice = OpenAIDetection.isInvoiceWithOpenAI(message);
              logWithUser(
                `OpenAI invoice detection result: ${isInvoice}`,
                "INFO"
              );
              return isInvoice;
            } catch (openaiError) {
              logWithUser(
                `OpenAI detection error: ${openaiError.message}, falling back to keywords`,
                "WARNING"
              );
              return CONFIG.fallbackToKeywords ? checkKeywords(message) : false;
            }
          } else if (CONFIG.fallbackToKeywords) {
            return checkKeywords(message);
          } else {
            return false;
          }
        }
      }
      // Try OpenAI if Gemini is disabled but OpenAI is enabled
      else if (CONFIG.invoiceDetection === "openai") {
        try {
          const isInvoice = OpenAIDetection.isInvoiceWithOpenAI(message);
          logWithUser(`OpenAI invoice detection result: ${isInvoice}`, "INFO");
          return isInvoice;
        } catch (openaiError) {
          logWithUser(
            `OpenAI detection error: ${openaiError.message}, falling back to keywords`,
            "WARNING"
          );
          return CONFIG.fallbackToKeywords ? checkKeywords(message) : false;
        }
      }
    }

    // Use keyword detection if no AI is enabled
    return checkKeywords(message);
  } catch (error) {
    logWithUser(`Error in invoice detection: ${error.message}`, "ERROR");
    return false;
  }
}

/**
 * Helper function to check for invoice keywords in message subject and body
 *
 * @param {GmailMessage} message - The Gmail message to check
 * @returns {boolean} True if invoice keywords are found
 */
function checkKeywords(message) {
  try {
    // Check subject for invoice keywords
    const subject = message.getSubject().toLowerCase();
    for (const keyword of CONFIG.invoiceKeywords) {
      if (subject.includes(keyword.toLowerCase())) {
        logWithUser(
          `Invoice keyword "${keyword}" found in subject: "${subject}"`,
          "INFO"
        );
        return true;
      }
    }

    // Check body for invoice keywords (optional, may affect performance)
    try {
      const body = message.getPlainBody().toLowerCase();
      for (const keyword of CONFIG.invoiceKeywords) {
        if (body.includes(keyword.toLowerCase())) {
          logWithUser(
            `Invoice keyword "${keyword}" found in message body`,
            "INFO"
          );
          return true;
        }
      }
    } catch (e) {
      // If we can't get the body, just log and continue
      logWithUser(`Could not check message body: ${e.message}`, "WARNING");
    }

    return false;
  } catch (error) {
    logWithUser(
      `Error checking for invoice keywords: ${error.message}`,
      "ERROR"
    );
    return false;
  }
}

/**
 * Determines if an attachment appears to be an invoice based on file extension and content type
 *
 * @param {GmailAttachment} attachment - The attachment to analyze
 * @returns {boolean} True if the attachment likely is an invoice
 */
function isInvoiceAttachment(attachment) {
  if (CONFIG.invoiceDetection === false) return false;

  try {
    const fileName = attachment.getName().toLowerCase();
    const contentType = attachment.getContentType().toLowerCase();

    // Check file extension against invoice file types with stricter checking
    for (const ext of CONFIG.invoiceFileTypes) {
      const lowerExt = ext.toLowerCase();

      // Strict PDF check - both extension and MIME type
      if (CONFIG.strictPdfCheck) {
        if (
          lowerExt === ".pdf" &&
          fileName.endsWith(lowerExt) &&
          contentType.includes("pdf")
        ) {
          return true;
        }
      } else {
        // Legacy check - just extension
        if (fileName.endsWith(lowerExt)) {
          return true;
        }
      }
    }

    return false;
  } catch (error) {
    logWithUser(
      `Error checking if attachment is invoice: ${error.message}`,
      "ERROR"
    );
    return false;
  }
}

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

              // Check if this message might contain invoices
              const isInvoice = isInvoiceMessage(message);
              let invoicesFolder = null;

              // If invoice detection is enabled and this might be an invoice,
              // get the invoices folder with domain subfolder
              if (CONFIG.invoiceDetection !== false && isInvoice) {
                logWithUser(`Message appears to contain invoice(s)`, "INFO");
                const senderDomain = extractDomain(sender);
                invoicesFolder = getInvoicesFolder(mainFolder, senderDomain);
              }

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

                    // If this is an invoice or the attachment looks like an invoice,
                    // also save it to the invoices folder
                    if (invoicesFolder || isInvoiceAttachment(attachment)) {
                      if (!invoicesFolder) {
                        logWithUser(
                          `Attachment appears to be an invoice based on file type`,
                          "INFO"
                        );
                        const senderDomain = extractDomain(sender);
                        invoicesFolder = getInvoicesFolder(
                          mainFolder,
                          senderDomain
                        );
                      }

                      // Save a copy to the invoices folder
                      try {
                        logWithUser(
                          `Saving copy to invoices folder: ${CONFIG.invoicesFolderName}`,
                          "INFO"
                        );
                        const invoiceFile = saveAttachmentLegacy(
                          attachment,
                          invoicesFolder,
                          messageDate
                        );

                        if (invoiceFile) {
                          logWithUser(
                            `Successfully saved invoice copy: ${attachment.getName()}`,
                            "INFO"
                          );
                        }
                      } catch (invoiceError) {
                        logWithUser(
                          `Error saving to invoices folder: ${invoiceError.message}`,
                          "ERROR"
                        );
                        // Continue processing even if saving to invoices folder fails
                      }
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

        // Check if this message might contain invoices
        const isInvoice = isInvoiceMessage(message);
        let invoicesFolder = null;

        // If invoice detection is enabled and this might be an invoice,
        // get the invoices folder with domain subfolder
        if (CONFIG.invoiceDetection !== false && isInvoice) {
          logWithUser(`Message appears to contain invoice(s)`, "INFO");
          invoicesFolder = getInvoicesFolder(mainFolder, domain);
        }

        result.totalAttachments += validAttachments.length;

        // Process each valid attachment
        for (const attachment of validAttachments) {
          // Save to domain folder
          const saveResult = saveAttachment(attachment, message, domainFolder);

          if (saveResult.success) {
            if (saveResult.duplicate) {
              result.duplicates++;
            } else {
              result.savedAttachments++;
              result.savedSize += attachment.getSize();

              // If this is an invoice or the attachment looks like an invoice,
              // also save it to the invoices folder
              if (invoicesFolder || isInvoiceAttachment(attachment)) {
                if (!invoicesFolder) {
                  logWithUser(
                    `Attachment appears to be an invoice based on file type`,
                    "INFO"
                  );
                  invoicesFolder = getInvoicesFolder(mainFolder, domain);
                }

                // Save a copy to the invoices folder
                try {
                  logWithUser(
                    `Saving copy to invoices folder: ${CONFIG.invoicesFolderName}`,
                    "INFO"
                  );
                  const invoiceResult = saveAttachment(
                    attachment,
                    message,
                    invoicesFolder
                  );

                  if (invoiceResult.success && !invoiceResult.duplicate) {
                    logWithUser(
                      `Successfully saved invoice copy: ${attachment.getName()}`,
                      "INFO"
                    );
                  }
                } catch (invoiceError) {
                  logWithUser(
                    `Error saving to invoices folder: ${invoiceError.message}`,
                    "ERROR"
                  );
                  // Continue processing even if saving to invoices folder fails
                }
              }
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

  // Validate AI configuration if enabled
  if (CONFIG.invoiceDetection === "gemini") {
    if (
      CONFIG.geminiApiKey === "__GEMINI_API_KEY__" &&
      !PropertiesService.getScriptProperties().getProperty(
        CONFIG.geminiApiKeyPropertyName
      )
    ) {
      throw new Error(
        "Configuration error: Gemini API key is not set but Gemini invoice detection is enabled."
      );
    }
  } else if (CONFIG.invoiceDetection === "openai") {
    if (
      CONFIG.openAIApiKey === "__OPENAI_API_KEY__" &&
      !PropertiesService.getScriptProperties().getProperty(
        CONFIG.openAIApiKeyPropertyName
      )
    ) {
      throw new Error(
        "Configuration error: OpenAI API key is not set but OpenAI invoice detection is enabled."
      );
    }
  }

  // Validate numerical values are within acceptable ranges
  if (CONFIG.aiConfidenceThreshold < 0 || CONFIG.aiConfidenceThreshold > 1) {
    throw new Error(
      "Configuration error: aiConfidenceThreshold must be between 0 and 1."
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

  // Validate interdependent settings
  if (CONFIG.onlyAnalyzePDFs && !CONFIG.invoiceFileTypes.includes(".pdf")) {
    throw new Error(
      "Configuration error: onlyAnalyzePDFs is enabled but .pdf is not in invoiceFileTypes."
    );
  }

  return true;
}

/**
 * Main function that processes Gmail attachments for authorized users
 *
 * This function:
 * 1. Validates the configuration
 * 2. Acquires an execution lock to prevent concurrent runs
 * 3. Gets the list of authorized users
 * 4. Processes emails for each user
 * 5. Logs completion
 *
 * @return {boolean} True if processing completed successfully, false if an error occurred or if no users are authorized.
 * The function follows a structured flow:
 * - It first validates the configuration to ensure all required settings are properly set.
 * - It attempts to acquire a lock to ensure no concurrent executions.
 * - It retrieves the list of authorized users.
 * - For each user, it processes their emails to save attachments to Google Drive.
 * - Logs are generated throughout the process to provide detailed information on the execution status.
 */
function saveAttachmentsToDrive() {
  try {
    logWithUser("Starting attachment processing", "INFO");

    // Validate configuration
    validateConfig();
    logWithUser("Configuration validated successfully", "INFO");

    // Acquire lock to prevent concurrent executions
    const currentUser = Session.getEffectiveUser().getEmail();
    if (!acquireExecutionLock(currentUser)) {
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
