# Gmail Attachment Organizer for Google Drive <!-- omit in toc -->

This Google Apps Script automatically saves Gmail attachments to Google Drive, organizing them by the sender's email domain. This makes it easy to track and manage attachments based on their source.

- [Simple Installation (Recommended)](#simple-installation-recommended)
- [Other Installation Options](#other-installation-options)
  - [Pre-configured Installation (Build Version)](#pre-configured-installation-build-version)
  - [Modular Installation (Multiple Files)](#modular-installation-multiple-files)
    - [Option 1: Manual Installation](#option-1-manual-installation)
    - [Option 2: Using Clasp (Recommended for Developers)](#option-2-using-clasp-recommended-for-developers)
- [Features](#features)
- [Execution Flow](#execution-flow)
  - [High-Level Overview](#high-level-overview)
  - [Detailed Breakdown](#detailed-breakdown)
- [Thread Labeling System](#thread-labeling-system)
- [Attachment Filtering System](#attachment-filtering-system)
- [Required Permissions](#required-permissions)
- [Managing Users](#managing-users)
- [Configuration Options](#configuration-options)
  - [AI Configuration](#ai-configuration)
    - [Gemini Configuration (Recommended)](#gemini-configuration-recommended)
    - [OpenAI Configuration (Legacy)](#openai-configuration-legacy)
    - [Shared AI Settings](#shared-ai-settings)
- [Advanced Usage](#advanced-usage)
  - [Custom Processing Options](#custom-processing-options)
- [Performance Considerations](#performance-considerations)
- [Customizing Attachment Filtering](#customizing-attachment-filtering)
- [File Structure](#file-structure)
- [Troubleshooting](#troubleshooting)
- [Available Commands](#available-commands)
- [License](#license)

**Key Features:**

- Domain-based folder organization
- AI-powered invoice detection (Gemini and OpenAI)
- Smart attachment filtering
- Batch processing and duplicate prevention
- Multi-user support and robust error handling
- Email date preservation

## Simple Installation (Recommended)

This is the quickest way to get the script running.

**Purpose:**

- Quick and easy setup for most users.

**Steps:**

1. Create a new project in Google Apps Script at \[script.google.com](<https://script.google.com>)
2. Copy the content from the `single-file/Code.gs` file and paste it into the editor.
3. Replace `__FOLDER_ID__` with your Google Drive folder ID.
4. Save and run the `saveAttachmentsToDrive` function.

**Ideal for:**

- Users who want a fast setup without managing multiple files.

**For detailed steps, refer to the \[Installation Steps](./INSTALLATION_STEPS.md).**

## Other Installation Options

- **Pre-configured Installation (Build Version):** Ideal for development or deploying to multiple environments. See details below.
- **Modular Installation (Multiple Files):** Recommended for long-term maintenance and complex projects. See details below.

### Pre-configured Installation (Build Version)

**Purpose:**

- Installation with `FOLDER_ID` pre-configured from your environment.

**Steps:**

1. Create a `.env` file in the root directory with your `FOLDER_ID` (see `.env.example`).
2. Run `npm install` to install dependencies.
3. Run `npm run build` to generate the pre-configured script in the `build` directory.
4. Run `npm run deploy` to upload the script to Google Apps Script.
5. Run `npm run open` to open the script in the Google Apps Script editor.

**Ideal for:**

- Development and deployment to multiple environments.

### Modular Installation (Multiple Files)

**Purpose:**

- Organized file structure for long-term maintenance.

#### Option 1: Manual Installation

1. Create a new project in Google Apps Script at \[script.google.com](<https://script.google.com>)
2. Create the following files in your project:
    - `Config.gs`
    - `Utils.gs`
    - `AIDetection.gs`
    - `GeminiDetection.gs`
    - `UserManagement.gs`
    - `AttachmentFilters.gs`
    - `FolderManagement.gs`
    - `AttachmentProcessing.gs`
    - `GmailProcessing.gs`
    - `Main.gs`
    - `Debug.gs` (optional)
    - `appsscript.json`
3. Copy the content of each file from the `src/` directory.
4. Update the `mainFolderId` in `Config.gs` with your Google Drive folder ID.
5. Save and run the `saveAttachmentsToDrive` function.

#### Option 2: Using Clasp (Recommended for Developers)

[Clasp](https://github.com/google/clasp) is Google's Command Line Apps Script Projects tool, which makes it easier to develop and manage Apps Script projects.

1. Install Clasp globally (requires Node.js):

   ```bash
   npm install -g @google/clasp
   ```

2. Login to your Google account:

   ```bash
   clasp login
   ```

3. Clone this repository:

   ```bash
   git clone https://github.com/yourusername/gmail-attachment-organizer.git
   cd gmail-attachment-organizer
   ```

4. Create a new Apps Script project:

   ```bash
   clasp create --title "Gmail Attachment Organizer" --rootDir ./src
   ```

5. Update the `mainFolderId` in `src/Config.gs` with your Google Drive folder ID.

6. Push the code to Google Apps Script:

   ```bash
   clasp push
   ```

7. Open the project in the browser:

   ```bash
   clasp open
   ```

8. Run the `saveAttachmentsToDrive` function from the Apps Script editor.

**Ideal for:**

- Ongoing development and maintenance.
- Developers who prefer working with version control and local editing.

## Features

This script provides the following features:

- **Domain-Based Organization:** Automatically creates folders based on sender domains.
- **AI-Powered Invoice Detection:** Uses Google Gemini or OpenAI to identify invoices with high accuracy.
  - **Gemini Details:** \[GEMINI\_INTEGRATION.md](./GEMINI_INTEGRATION.md)
  - **OpenAI Details:** \[OPENAI\_INTEGRATION.md](./OPENAI_INTEGRATION.md)
  - **PDF-Only Processing:** AI invoice detection is activated only for emails with PDF attachments.
  - **Privacy-Focused Analysis (Gemini):** Gemini integration analyzes only metadata, not full email content.
  - **Personal Data Protection (Gemini):** Email addresses are anonymized; only domain names are shared with AI.
- **Smart Attachment Filtering:** Identifies and skips embedded images and email signatures.
- **Batch Processing:** Processes emails in batches to avoid timeout issues.
- **Duplicate Prevention:** Prevents saving duplicate files.
- **Multi-User Support:** Processes emails for multiple users and maintains a user queue.
- **Robust Error Handling:** Comprehensive try/catch blocks with logging.
- **Oldest-First Processing:** Processes emails from oldest to newest by default (configurable).
- **Email Date Preservation:** Stores original email date in file descriptions.

## Execution Flow

### High-Level Overview

1. Script Initialization
2. User Processing
3. Email Discovery
4. Attachment Processing
5. Error Handling

### Detailed Breakdown

1. **Script Initialization:**
    - Validates configuration settings.
    - Obtains reference to the main Google Drive folder.
    - Gets or creates the Gmail label for marking processed emails.
    - Gets the next user in the queue to process.
2. **User Processing:**
    - Verifies user permissions for Gmail and Drive access.
    - Switches to user context for processing.
    - Maintains a queue of users to process in turn.
3. **Email Discovery:**
    - Searches Gmail for unprocessed emails with attachments.
    - Processes emails from oldest to newest by default (configurable).
    - Limits processing to a configurable batch size to prevent timeouts.
4. **Attachment Processing:**
    - For each email thread:
        - Extracts all messages in the thread.
        - For each message with attachments:
            - Extracts the sender's domain from their email address.
            - Filters out unwanted attachments (small images, calendar invitations).
            - Creates domain-specific folder only when necessary.
            - For each valid attachment:
                - Ensures filename uniqueness.
                - Saves the attachment to the appropriate folder.
        - Marks the thread as processed by applying the Gmail label.
5. **Error Handling:**
    - Comprehensive try/catch blocks at multiple levels.
    - Detailed logging for troubleshooting.
    - Graceful failure handling to prevent script termination on single-item errors.
    - Retry logic with exponential backoff for transient errors.

## Thread Labeling System

The script uses labels at the thread level to track processed emails.

**Label Application:**

- The "GDrive_Processed" label is applied to entire Gmail threads, not individual messages.
- A thread is labeled when:
  - At least one valid attachment is saved from any message in the thread, OR
  - All messages in the thread have attachments, but they were all filtered out (e.g., embedded images).
- The script never labels threads that have no attachments at all.

**Handling New Messages in Processed Threads:**

- New messages added to a thread that already has the "GDrive_Processed" label will NOT be processed automatically by this script.
- If a new message with an important attachment is added to a previously processed thread:
  - You can manually remove the "GDrive_Processed" label from the thread to force reprocessing in the next script run.
  - Forward the message to yourself in a new thread (creating a new thread without the label).

**Rationale for Thread-Level Labeling:**

- Gmail's API is optimized for thread-level operations.
- Most email conversations maintain context in a thread.
- Reduces processing overhead by avoiding repeated analysis of related messages.
- Prevents running into quota limits for Gmail API calls.

**Note:** If you frequently receive important new attachments in existing threads, consider these options:

1. Adjust email sending/replying behavior to create new threads for important attachments.
2. Implement a more sophisticated Message-ID based tracking system (would require significant code changes).
3. Run the script more frequently and maintain a separate record of processed message IDs.

## Attachment Filtering System

The script uses a sophisticated filtering system to distinguish between real attachments and embedded elements like email signatures, logos, and inline images.

To configure attachment filtering, you can modify the settings described below and in the "Configuration Options" section.

**Filtering Methods:**

The script employs a multi-tiered approach to ensure accurate attachment filtering:

1. **MIME Type Whitelist:**
    - A list of MIME types (`attachmentTypesWhitelist`) that are always saved (e.g., PDF, Word, Excel).
    - If an attachment has a MIME type on this list, it is always kept regardless of other filters.
    - If a file doesn't match the whitelist, it will be evaluated by the additional criteria below.

2. **Content-Disposition Analysis:**
    - Checks if an image is specifically marked as "inline" in its `content-disposition` header.
    - Inline images are typically embedded in the email body rather than explicit attachments.

3. **HTML Email Image URL Detection:**
    - Identifies images referenced by URLs in HTML emails from various email providers.
    - Detects Gmail embedded image URLs with parameters like `view=fimg` or `disp=emb`.
    - Recognizes Outlook, Yahoo Mail, and other common email image URL patterns.
    - Filters out image URLs that use common patterns like `cid=` or specific domains.

4. **Filename Pattern Recognition:**
    - Detects common patterns for embedded images (e.g., `image001.png`, `inline-`, `Outlook-`).
    - Checks against a list of common embedded element names (logos, icons, banners, etc.).

5. **File Extension Filtering:**
    - Special handling for common document extensions (`.pdf`, `.doc`, `.docx`, etc.) - always kept.
    - Option to filter unwanted file types (e.g., calendar invitations - `.ics` files).
    - Option to filter small images that are likely signatures or icons.

6. **Size-Based Filtering:**
    - Files without extensions under a certain size (e.g., 50KB) are likely embedded content.
    - Configuration to skip images below a certain size threshold.

These combined approaches ensure that only legitimate attachments are saved while filtering out embedded elements that would clutter your Drive storage.

## Required Permissions

This script requires the following authorization scopes. You'll be prompted to grant these permissions during the script's initial setup.

- `<https://www.googleapis.com/auth/gmail.modify>` (for reading emails and applying labels)
- `<https://www.googleapis.com/auth/drive>` (for creating folders and files in Google Drive)
- `<https://www.googleapis.com/auth/script.scriptapp>` (for creating triggers)
- `<https://www.googleapis.com/auth/script.external_request>` (for external API calls if needed)

## Managing Users

The script includes functions for managing user access, primarily useful in multi-user environments.

**Available User Management Functions:**

- `addUserToList('email@domain.com')`: Adds a user to the list of processed users.
- `removeUserFromList('email@domain.com')`: Removes a user from the list.
- `listUsers()`: Displays a list of the currently managed users.
- `verifyUserPermissions('email@domain.com')`: Checks if a specific user has the necessary permissions.

## Configuration Options

The `Config.gs` file contains all configurable options, allowing you to tailor the script's behavior to your specific needs.

**General Configuration Options:**

- `mainFolderId`: ID of the main Google Drive folder where attachments will be saved.
- `processedLabelName`: Name of the Gmail label to apply to processed messages.
- `skipDomains`: Array of email domains to exclude from processing.
- `skipSmallImages`: Set to `true` to avoid saving small images like email signatures.
- `smallImageMaxSize`: Maximum size in bytes for images to be skipped (default: 20KB).
- `smallImageExtensions`: File extensions to consider as images for filtering.
- `skipFileTypes`: Additional file types to skip (e.g., calendar invitations, etc.).
- `attachmentTypesWhitelist`: List of MIME types that should always be saved.
- `batchSize`: Number of threads to process in each execution (default: 20).

### AI Configuration

The script integrates with AI services (Gemini and OpenAI) to provide invoice detection. You can configure the AI behavior using the following options.

#### Gemini Configuration (Recommended)

These options control the Gemini AI integration.

- `geminiEnabled`: Enable/disable Gemini API for invoice detection.
- `geminiApiKey`: API key for Gemini (set via .env file or Script Properties).
- `geminiModel`: Model to use (default: "gemini-2.0-flash").
- `geminiMaxTokens`: Maximum tokens for the AI response.
- `geminiTemperature`: Temperature setting for response determinism (0.0 - 1.0).

**Sample Metadata Sent to Gemini:**

To ensure transparency and demonstrate our commitment to user privacy, here's an example of the metadata sent to the Gemini API:

```json
{
  "subject": "Invoice #12345 - Example Company",
  "senderDomain": "example.com",
  "date": "2025-03-15T10:30:00.000Z",
  "hasAttachments": true,
  "attachmentTypes": [
    "pdf"
  ],
  "attachmentContentTypes": [
    "application/pdf"
  ],
  "keywordsFound": [
    "invoice",
    "#12345",
    "payment",
    "$299.99"
  ],
  "historicalPatterns": {
    "count": 3,
    "subjectPatterns": {
      "commonPrefix": "Invoice",
      "commonSuffix": "Example Company",
      "containsInvoiceTerms": true,
      "hasNumericPatterns": true
    },
    "datePatterns": {
      "frequency": "monthly",
      "averageIntervalDays": 30,
      "sameDayOfMonth": true
    },
    "rawSubjects": [
      "Invoice #12342 - Example Company",
      "Invoice #12343 - Example Company",
      "Invoice #12344 - Example Company"
    ],
    "rawDates": [
      "2024-12-15T10:30:00.000Z",
      "2025-01-15T10:30:00.000Z",
      "2025-02-15T10:30:00.000Z"
    ]
  }
}
```

**Important Privacy Note:**

- Only domain names are sent, not full email addresses, to protect user privacy.

#### OpenAI Configuration (Legacy)

These options are for the (now legacy) OpenAI integration.

- `openAIEnabled`: Enable/disable OpenAI API for invoice detection.
- `openAIApiKey`: API key for OpenAI (set via .env file or Script Properties).
- `openAIModel`: Model to use (default: "gpt-3.5-turbo").
- `openAIMaxTokens`: Maximum tokens for the AI response.
- `openAITemperature`: Temperature setting for response determinism (0.0 - 1.0).

#### Shared AI Settings

These settings are shared between the Gemini and OpenAI integrations.

- `skipAIForDomains`: Domains to exclude from AI analysis.
- `onlyAnalyzePDFs`: Only process emails with PDF attachments for AI analysis.
- `strictPdfCheck`: Check both file extension and MIME type for PDFs.
- `fallbackToKeywords`: Fall back to keyword detection if AI fails.
- `aiConfidenceThreshold`: Confidence threshold for AI detection (0.0-1.0).

**Additional AI Information:**

- \[GEMINI\_INTEGRATION.md](./GEMINI_INTEGRATION.md) for Gemini details
- \[OPENAI\_INTEGRATION.md](./OPENAI_INTEGRATION.md) for OpenAI details

## Advanced Usage

This section covers advanced usage patterns and customization options.

### Custom Processing Options

The main function `saveAttachmentsToDrive()` supports optional parameters to modify its behavior:

```javascript
//   Process with default options (oldest emails first)
saveAttachmentsToDrive();

//   Process newest emails first
saveAttachmentsToDrive({ oldestFirst: false });
```

By default, the script processes emails from oldest to newest. You can use the `oldestFirst` parameter to change this.

## Performance Considerations

The script incorporates several optimizations to ensure efficient processing, even with large volumes of emails.

**Tips for Handling Large Volumes of Emails:**

- **Adjust Batch Size:** Increase the `batchSize` if your script execution time allows, or decrease it for more frequent, smaller runs.
- **Logging Verbosity:** Adjust logging levels in the configuration to reduce output.

## Customizing Attachment Filtering

You can customize the script's attachment filtering behavior to suit your needs.

**Customization Methods:**

1. **Always Saving Specific MIME Types:**
    - Add the MIME type to the `attachmentTypesWhitelist` array in `Config.gs`.
    - Example: To always save PNG files, add `"image/png"` to the list.

2. **Skipping Specific File Types:**
    - Add the file extensions to the `skipFileTypes` array in `Config.gs`.
    - Example: To skip all text files, add `".txt"` to the list.

3. **Adjusting Image Filtering:**
    - Modify the `smallImageMaxSize` and `smallImageExtensions` settings in `Config.gs`.
    - Increasing `smallImageMaxSize` will cause more images to be skipped.

4. **Troubleshooting Skipped Files:**
    - Check the script's execution logs to understand why files are being filtered.
    - Adjust the relevant settings based on the log information.

## File Structure

Understanding the file structure can be helpful for debugging or extending the script.

- `Config.gs`: Configuration settings and constants.
- `Utils.gs`: Utility functions for logging, retries, and user settings.
- `AIDetection.gs`: OpenAI API integration for invoice detection (legacy).
- `GeminiDetection.gs`: Gemini API integration for invoice detection (recommended).
- `UserManagement.gs`: User authorization and permission management.
- `AttachmentFilters.gs`: Functions for determining which attachments to process.
- `FolderManagement.gs`: Google Drive folder creation and management.
- `AttachmentProcessing.gs`: Functions for saving attachments to Drive.
- `GmailProcessing.gs`: Gmail thread and message processing.
- `Main.gs`: Entry points and main execution flow.
- `appsscript.json`: Script manifest with required OAuth scopes.

## Troubleshooting

This section provides guidance on resolving common issues.

**Troubleshooting Common Issues:**

- **Permissions Issues:**
  - Run `verifyAllUsersPermissions()` to check for permission problems.

- **Missing Attachments:**
  - Check the execution logs to determine if attachments were filtered.

- **Script Timeout:**
  - Reduce the `batchSize` in the `Config.gs` configuration file.

- **Duplicate Folders:**
  - The script includes lock logic to prevent duplicate folder creation, but it can occur during simultaneous executions.

- **Unexpected Thread Processing:**
  - If a thread is processed but attachments are missing, review the logs to see if they were filtered (e.g., small images, embedded content).

- **Gemini API Issues:**
  - If Gemini invoice detection is not working:
    - Run `GeminiDetection.testGeminiConnection()` to verify API connectivity.
    - Ensure your Gemini API key is valid.
    - Run `npm run test:gemini` locally to test the API integration.
    - Adjust the `aiConfidenceThreshold` if you experience false positives or negatives.
    - Review logs in the `logs` directory for detailed error information.

- **OpenAI API Issues:**
  - If OpenAI invoice detection is not working:
    - Run `AIDetection.testOpenAIConnection()` to verify API connectivity.
    - Ensure your OpenAI API key is valid and has sufficient credits.
    - Run `npm run test:openai` locally to test the API integration.
    - Review logs in the `logs` directory for detailed error information.

## Available Commands

These commands are available when dependencies are installed and can be used to manage the script.

```bash
#   Generate files in single-file and build directories
npm run build

#   Generate files and upload them to Google Apps Script
npm run deploy

#   Force upload to Google Apps Script (overwrites any changes)
npm run deploy:force

#   Login to Google
npm run login

#   Logout
npm run logout

#   View status of files
npm run status

#   Open the script in Google Apps Script editor
npm run open

#   Download the latest version of the script from Google Apps Script
npm run pull

#   Test if the folder ID is valid
npm run test

#   Create a new version of the script
npm run version

#   Upload and create version in a single step
npm run deploy:version
```

## License

\[MIT License](./LICENSE.md)
