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
  invoiceDetection: false, // AI provider to use: "gemini" (recommended), "openai", or false to disable
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
