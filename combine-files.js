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
 */

const fs = require("fs");
const path = require("path");
const dotenv = require("dotenv");

// Load environment variables from .env file
dotenv.config();

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

// Function to get project timezone from appsscript.json
function getProjectTimezone() {
  try {
    const appsscriptPath = path.join(__dirname, "appsscript.json");
    if (fs.existsSync(appsscriptPath)) {
      const appsscriptContent = fs.readFileSync(appsscriptPath, "utf8");
      const appsscriptJson = JSON.parse(appsscriptContent);
      return appsscriptJson.timeZone || "UTC";
    }
  } catch (error) {
    console.warn(
      `Warning: Could not read timezone from appsscript.json: ${error.message}`
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

    // Create a version for build (with FOLDER_ID from .env)
    let buildContent = content;
    if (filename === "Config.gs" && process.env.FOLDER_ID) {
      buildContent = buildContent.replace(
        'mainFolderId: "__FOLDER_ID__"',
        `mainFolderId: "${process.env.FOLDER_ID}"`
      );
    }
    fs.appendFileSync(outputFileBuild, buildContent);
  } else {
    console.warn(
      `Warning! The file src/${filename} does not exist and will be skipped.`
    );
  }
});

console.log(`\nCombined files created:`);
console.log(`- ${outputFileSingle} (with placeholder FOLDER_ID)`);

// Ocultar parte del FOLDER_ID, mostrar solo 4 caracteres al principio y al final
let displayFolderId = "not found";
if (process.env.FOLDER_ID) {
  const id = process.env.FOLDER_ID;
  if (id.length > 8) {
    displayFolderId = `${id.substring(0, 4)}...${id.substring(id.length - 4)}`;
  } else {
    displayFolderId = id;
  }
  console.log(
    `- ${outputFileBuild} (with FOLDER_ID from .env: ${displayFolderId})`
  );
} else {
  console.log(`- ${outputFileBuild} (FOLDER_ID not found in .env file)`);
}

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

// Copy appsscript.json to both directories
try {
  const appsscriptPath = path.join(__dirname, "appsscript.json");
  if (fs.existsSync(appsscriptPath)) {
    // Copy to single-file directory
    fs.copyFileSync(
      appsscriptPath,
      path.join(singleFileDir, "appsscript.json")
    );

    // Copy to build directory
    fs.copyFileSync(appsscriptPath, path.join(buildDir, "appsscript.json"));

    console.log("appsscript.json copied to both directories");
  } else {
    console.warn("Warning! appsscript.json not found in workspace root");
  }
} catch (error) {
  console.error(`Error copying appsscript.json: ${error.message}`);
}
