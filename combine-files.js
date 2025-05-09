#!/usr/bin/env node
/**
 * Script to combine all .gs files into a single Code.gs file
 * maintaining a logical order for the code structure.
 *
 * This script generates two output files:
 * 1. A "Code.gs" file in the build/ directory with environment variables
 *    substituted, ready to be pushed to Google Apps Script
 * 2. A "Code.gs" file in the single-file/ directory without sensitive
 *    environment variables, suitable for sharing publicly
 *
 * The script maintains a logical order of the source files to ensure
 * proper code structure and dependencies in the combined output.
 *
 * Environment support:
 * - Use --env=prod for production environment (default)
 * - Use --env=test for test environment
 */

const fs = require("fs");
const path = require("path");
const dotenv = require("dotenv");

// Parse command line arguments for environment
const args = process.argv.slice(2);
const envArg = args.find((arg) => arg.startsWith("--env="));
const env = envArg ? envArg.split("=")[1] : "prod"; // Default to prod if not specified

console.log(`Building for environment: ${env}`);

// Load environment variables from the appropriate .env file
let envPath = path.resolve(__dirname, `.env.${env}`);

// If the environment-specific file doesn't exist, try the default .env file
if (!fs.existsSync(envPath)) {
  console.log(`Environment file .env.${env} not found, trying .env instead`);
  envPath = path.resolve(__dirname, ".env");

  if (!fs.existsSync(envPath)) {
    console.error(`Error: No environment file found (.env.${env} or .env)!`);
    process.exit(1);
  }
}

console.log(`Loading environment variables from: ${envPath}`);
// Check if the file exists and is readable
try {
  const envContent = fs.readFileSync(envPath, "utf8");
  console.log(`Environment file content:\n${envContent}`);
} catch (error) {
  console.error(`Error reading environment file: ${error.message}`);
}

// Try loading with dotenv
dotenv.config({ path: envPath });

// Debug: Print loaded environment variables
console.log("Loaded environment variables:");
console.log(`FOLDER_ID: ${process.env.FOLDER_ID || "not set"}`);
console.log(
  `OPENAI_API_KEY: ${process.env.OPENAI_API_KEY ? "set (masked)" : "not set"}`
);
console.log(
  `GEMINI_API_KEY: ${process.env.GEMINI_API_KEY ? "set (masked)" : "not set"}`
);
console.log(
  `PROCESSED_LABEL_NAME: ${process.env.PROCESSED_LABEL_NAME || "not set"}`
);
console.log(
  `INVOICES_FOLDER_NAME: ${process.env.INVOICES_FOLDER_NAME || "not set"}`
);

// Logical order of files
const fileOrder = [
  "Config.gs", // Configuration first
  "Utils.gs", // General utilities
  "HistoricalPatterns.gs", // Historical pattern analysis for invoice detection
  "OpenAIDetection.gs", // OpenAI Detection for invoice analysis (legacy)
  "GeminiDetection.gs", // Gemini Detection for invoice analysis (recommended)
  "InvoiceDetection.gs", // Email-based invoice detection logic
  "VerifyAPIKeys.gs", // API key verification utilities
  "UserManagement.gs", // User management
  "AttachmentFilters.gs", // Attachment filters
  "FolderManagement.gs", // Folder management
  "AttachmentProcessing.gs", // Attachment processing
  "GmailProcessing.gs", // Gmail processing
  "Main.gs", // Main functions
  // Debug files are intentionally not included
];

// Files to copy separately (not merged into Code.gs)
// This includes both .gs files that should remain separate and other file types
const filesToCopyDirectly = [
  "InvoiceSenders.gs", // List of invoice senders for email-based detection
  "appsscript.json", // Project manifest
  // Add any other files that should be copied directly here
];

// Function to get project timezone from src/appsscript.json
function getProjectTimezone() {
  try {
    const appsscriptPath = path.join(__dirname, "src", "appsscript.json");
    if (fs.existsSync(appsscriptPath)) {
      const appsscriptContent = fs.readFileSync(appsscriptPath, "utf8");
      const appsscriptJson = JSON.parse(appsscriptContent);
      return appsscriptJson.timeZone || "UTC";
    }
  } catch (error) {
    console.warn(
      `Warning: Could not read timezone from src/appsscript.json: ${error.message}`
    );
  }
  return "UTC";
}

// Get the current timestamp in the project timezone
function getFormattedTimestamp() {
  const timezone = getProjectTimezone();
  const now = new Date();
  const currentYear = now.getFullYear(); // Obtenemos el año actual

  // Format: MM/DD/YYYY HH:MM:SS Timezone
  const options = {
    year: "numeric",
    month: "2-digit",
    day: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit",
    hour12: false,
    timeZone: timezone,
  };

  try {
    // Obtener la fecha formateada
    let formattedDate = new Intl.DateTimeFormat("en-US", options).format(now);

    // Asegurarnos de que el año es correcto (forzar año actual)
    // La fecha formateada tiene formato "MM/DD/YYYY, HH:MM:SS"
    // Reemplazamos el año en la fecha formateada
    formattedDate = formattedDate.replace(/\d{4}/, currentYear);

    return `${formattedDate} (${timezone})`;
  } catch (error) {
    console.warn(
      `Warning: Could not format date with timezone ${timezone}: ${error.message}`
    );
    // Fallback seguro con el año correcto
    const month = String(now.getMonth() + 1).padStart(2, "0");
    const day = String(now.getDate()).padStart(2, "0");
    const hours = String(now.getHours()).padStart(2, "0");
    const minutes = String(now.getMinutes()).padStart(2, "0");
    const seconds = String(now.getSeconds()).padStart(2, "0");

    return `${month}/${day}/${currentYear}, ${hours}:${minutes}:${seconds} (fallback from ${timezone})`;
  }
}

// Header for the combined file
const header = `/**
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

`;

// Create the single-file directory if it doesn't exist
const singleFileDir = path.join(__dirname, "single-file");
if (!fs.existsSync(singleFileDir)) {
  fs.mkdirSync(singleFileDir, { recursive: true });
}

// Create the build directory if it doesn't exist
const buildDir = path.join(__dirname, "build");
if (!fs.existsSync(buildDir)) {
  fs.mkdirSync(buildDir, { recursive: true });
}

// Output files
const outputFileSingle = path.join(singleFileDir, "Code.gs");
const outputFileBuild = path.join(buildDir, "Code.gs");

// Start with the header for single-file
fs.writeFileSync(outputFileSingle, header);

// Get build timestamp
const buildTimestamp = getFormattedTimestamp();

// Start with the header and timestamp comment for build version
fs.writeFileSync(
  outputFileBuild,
  `// Build generated on: ${buildTimestamp}\n${header}`
);

// Process each file in the specified order
fileOrder.forEach((filename) => {
  const filePath = path.join(__dirname, "src", filename);

  if (fs.existsSync(filePath)) {
    console.log(`Processing ${filename}...`);

    // Read the file content
    let content = fs.readFileSync(filePath, "utf8");

    // Extract the module name from the filename (without extension)
    const moduleName = path.basename(filename, ".gs").toUpperCase();

    // Add a section separator before the content
    content = `\n//=============================================================================
// ${moduleName} - ${getModuleDescription(moduleName)}
//=============================================================================\n
${content}`;

    // Remove header comments from the original file if they exist
    content = content.replace(/\/\*\*[\s\S]*?\*\/\s*/, "");

    // Create a version for single-file (unchanged)
    fs.appendFileSync(outputFileSingle, content);

    // Create a version for build (with environment variables from .env)
    let buildContent = content;
    if (filename === "Config.gs") {
      // Replace configuration values with environment variables using regex
      // This approach is more robust as it doesn't rely on specific default values

      // Replace FOLDER_ID if available
      if (process.env.FOLDER_ID) {
        const mainFolderIdRegex = /(mainFolderId:\s*")[^"]*(")/;
        buildContent = buildContent.replace(
          mainFolderIdRegex,
          `$1${process.env.FOLDER_ID}$2`
        );
      }

      // Replace OPENAI_API_KEY if available
      if (process.env.OPENAI_API_KEY) {
        const openAIApiKeyRegex = /(openAIApiKey:\s*")[^"]*(")/;
        buildContent = buildContent.replace(
          openAIApiKeyRegex,
          `$1${process.env.OPENAI_API_KEY}$2`
        );
      }

      // Replace GEMINI_API_KEY if available
      if (process.env.GEMINI_API_KEY) {
        const geminiApiKeyRegex = /(geminiApiKey:\s*")[^"]*(")/;
        buildContent = buildContent.replace(
          geminiApiKeyRegex,
          `$1${process.env.GEMINI_API_KEY}$2`
        );
      }

      // Replace processedLabelName if available
      if (process.env.PROCESSED_LABEL_NAME) {
        const processedLabelNameRegex = /(processedLabelName:\s*")[^"]*(")/;
        buildContent = buildContent.replace(
          processedLabelNameRegex,
          `$1${process.env.PROCESSED_LABEL_NAME}$2`
        );
      }

      // Replace invoicesFolderName if available
      if (process.env.INVOICES_FOLDER_NAME) {
        const invoicesFolderNameRegex = /(invoicesFolderName:\s*")[^"]*(")/;
        buildContent = buildContent.replace(
          invoicesFolderNameRegex,
          `$1${process.env.INVOICES_FOLDER_NAME}$2`
        );
      }
    }
    fs.appendFileSync(outputFileBuild, buildContent);
  } else {
    console.warn(
      `Warning! The file src/${filename} does not exist and will be skipped.`
    );
  }
});

console.log(`\nCombined files created:`);
console.log(`- single-file/Code.gs (with PLACEHOLDERS)`);
console.log(`- build/Code.gs (with VALUES)`);

// Function to get an appropriate description for each module
function getModuleDescription(moduleName) {
  const descriptions = {
    CONFIG: "CONFIGURATION",
    UTILS: "UTILITIES",
    HISTORICALPATTERNS: "HISTORICAL PATTERN ANALYSIS",
    OPENAIDETECTION: "OPENAI DETECTION",
    GEMINIDETECTION: "GEMINI DETECTION",
    INVOICEDETECTION: "EMAIL-BASED INVOICE DETECTION",
    VERIFYAPIKEYS: "API KEY VERIFICATION",
    USERMANAGEMENT: "USER MANAGEMENT",
    ATTACHMENTFILTERS: "ATTACHMENT FILTERS",
    FOLDERMANAGEMENT: "FOLDER MANAGEMENT",
    ATTACHMENTPROCESSING: "ATTACHMENT PROCESSING",
    GMAILPROCESSING: "GMAIL PROCESSING",
    MAIN: "MAIN FUNCTIONS",
  };

  return descriptions[moduleName] || moduleName;
}

// Generate .clasp.json file with the correct scriptId
if (process.env.SCRIPT_ID) {
  const claspConfig = {
    scriptId: process.env.SCRIPT_ID,
    rootDir: "build",
    scriptExtensions: [".js", ".gs"],
    htmlExtensions: [".html"],
    jsonExtensions: [".json"],
    filePushOrder: [],
    skipSubdirectories: false,
  };

  fs.writeFileSync(
    path.join(__dirname, ".clasp.json"),
    JSON.stringify(claspConfig, null, 2)
  );

  console.log(`\nUpdated .clasp.json for ${env} environment`);
} else {
  console.warn(
    "Warning! SCRIPT_ID not found in environment variables. .clasp.json not updated."
  );
}

// Copy files that should remain separate (not merged into Code.gs)
console.log("\nCopying files directly (not merged into Code.gs):");
filesToCopyDirectly.forEach((filename) => {
  try {
    const filePath = path.join(__dirname, "src", filename);

    if (fs.existsSync(filePath)) {
      // Copy to single-file directory
      fs.copyFileSync(filePath, path.join(singleFileDir, filename));

      // Copy to build directory
      fs.copyFileSync(filePath, path.join(buildDir, filename));

      console.log(`- ${filename} copied to both directories`);
    } else {
      console.warn(
        `Warning! The file src/${filename} does not exist and will be skipped.`
      );
    }
  } catch (error) {
    console.error(`Error copying ${filename}: ${error.message}`);
  }
});
