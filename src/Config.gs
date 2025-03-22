/**
 * Configuration for Gmail Attachment Organizer
 *
 * This file contains all configurable settings for the script.
 */

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
  // Invoice detection and special folder configuration
  invoicesFolderName: "Facturas", // Name of the special folder for invoices
  invoiceDetection: true, // Enable/disable invoice detection feature
  invoiceKeywords: ["factura", "invoice", "receipt", "recibo", "pago"], // Keywords to search in subject/body
  invoiceFileTypes: [".pdf"], // File types considered as invoices
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
