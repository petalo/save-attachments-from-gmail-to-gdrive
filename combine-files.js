#!/usr/bin/env node

/**
 * Script to combine all .gs files into a single Code.gs file
 * maintaining a logical order for the code structure.
 */

const fs = require("fs");
const path = require("path");

// Logical order of files
const fileOrder = [
  "Config.gs", // Configuration first
  "Utils.gs", // General utilities
  "UserManagement.gs", // User management
  "AttachmentFilters.gs", // Attachment filters
  "FolderManagement.gs", // Folder management
  "AttachmentProcessing.gs", // Attachment processing
  "GmailProcessing.gs", // Gmail processing
  "Main.gs", // Main functions
  // Debug files are intentionally not included
];

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

// Output file
const outputFile = path.join(singleFileDir, "Code.gs");

// Start with the header
fs.writeFileSync(outputFile, header);

// Process each file in the specified order
fileOrder.forEach((filename) => {
  const filePath = path.join(__dirname, filename);

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

    // Append to the combined file
    fs.appendFileSync(outputFile, content);
  } else {
    console.warn(
      `Warning! The file ${filename} does not exist and will be skipped.`
    );
  }
});

console.log(`\nCombined file created: ${outputFile}`);

// Function to get an appropriate description for each module
function getModuleDescription(moduleName) {
  const descriptions = {
    CONFIG: "CONFIGURATION",
    UTILS: "UTILITIES",
    USERMANAGEMENT: "USER MANAGEMENT",
    ATTACHMENTFILTERS: "ATTACHMENT FILTERS",
    FOLDERMANAGEMENT: "FOLDER MANAGEMENT",
    ATTACHMENTPROCESSING: "ATTACHMENT PROCESSING",
    GMAILPROCESSING: "GMAIL PROCESSING",
    MAIN: "MAIN FUNCTIONS",
  };

  return descriptions[moduleName] || moduleName;
}
