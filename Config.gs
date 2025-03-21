/**
 * Configuration for Gmail Attachment Organizer
 *
 * This file contains all configurable settings for the script.
 */

// Configuration constants
const CONFIG = {
  mainFolderId: "YOUR_SHARED_FOLDER_ID", // Replace with your shared folder's ID
  processedLabelName: "GDrive_Processed", // Label to mark processed threads
  skipDomains: ["example.com", "noreply.com"], // Skip emails from these domains
  maxFileSize: 25 * 1024 * 1024, // 25MB max file size
  batchSize: 10, // Process this many threads at a time to avoid the 6 minutes execution limit
  skipSmallImages: true, // Skip small images like email signatures
  smallImageMaxSize: 20 * 1024, // 20KB max size for images to skip
  smallImageExtensions: [".jpg", ".jpeg", ".png", ".gif", ".bmp"], // Image extensions to check
  skipFileTypes: [".ics", ".ical", ".pkpass", ".vcf", ".vcard"], // Additional file types to skip (e.g., calendar invitations, etc.)
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
