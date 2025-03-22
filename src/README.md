# Gmail Attachment Organizer - Modular Version

This is the modular version of the Gmail Attachment Organizer for Google Drive. The code is organized into multiple files to facilitate maintenance and debugging.

## File Structure

- `Config.gs`: Configuration and constants
- `Utils.gs`: Utility functions
- `UserManagement.gs`: User management and permissions
- `AttachmentFilters.gs`: Attachment filtering logic
- `FolderManagement.gs`: Drive folder management
- `AttachmentProcessing.gs`: Attachment processing
- `GmailProcessing.gs`: Gmail thread and message processing
- `Main.gs`: Entry points and main flow
- `Debug.gs`: Diagnostic functions (optional)
- `appsscript.json`: Manifest with OAuth permissions

## Installation

1. Create a new project in Google Apps Script at [script.google.com](https://script.google.com)
2. Create each of the files listed above in your project
3. Copy the content of each file from this directory
4. Edit `Config.gs` and update the `mainFolderId` value with your Google Drive folder ID
5. Save the project with a descriptive name
6. Run the `requestPermissions` function to grant the necessary permissions
7. Run the `createTrigger` function to set up automatic execution every 15 minutes

## Main Features

- Domain-based organization: Creates folders based on sender domain
- Smart filtering: Detects and skips embedded images and email signatures
- Batch processing: Processes emails in batches to avoid timeout issues
- Duplicate prevention: Avoids saving duplicate files
- Multi-user support: Processes emails for multiple users
- Robust error handling: Complete try/catch blocks with detailed logging
- Oldest-first processing: By default, processes emails from oldest to newest

## Diagnostics and Testing

The `Debug.gs` file contains several useful functions for testing and diagnosing issues:

- `testRun()`: Runs the script with a smaller set of threads
- `getConfig()`: Displays the current configuration
- `diagnoseMissingEmails()`: Diagnoses issues with email search
- `testTimestamps()`: Tests the timestamp functionality

## For More Information

See the [complete repository](https://github.com/your-username/gmail-attachment-organizer) for more detailed documentation. If you prefer a simpler installation, consider using the single-file version.
