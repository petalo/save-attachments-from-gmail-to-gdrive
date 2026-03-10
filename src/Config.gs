/**
 * Configuration for Gmail Attachment Organizer
 *
 * This file contains all configurable settings for the script, organized into logical sections.
 * Each section groups related settings to make configuration more intuitive and maintainable.
 */

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
