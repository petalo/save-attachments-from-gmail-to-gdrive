# Gmail Attachment Organizer - Single File Version

This is the simplified (monolithic) version of the Gmail Attachment Organizer for Google Drive. All code is contained in a single file for easy installation.

## Quick Installation

1. Create a new project in Google Apps Script at [script.google.com](https://script.google.com)
2. Copy all the content from the `Code.gs` file and paste it into the editor
3. Replace `YOUR_SHARED_FOLDER_ID` with your Google Drive folder ID (the long string that appears after `/folders/` in your folder's URL)
4. Save the project with a descriptive name
5. Run the `requestPermissions` function to grant the necessary permissions
6. Run the `createTrigger` function to set up automatic execution every 15 minutes
7. That's it! The script will start processing your emails and saving attachments

## Main Features

- Domain-based organization: Creates folders based on sender domain
- Smart filtering: Detects and skips embedded images and email signatures
- Batch processing: Processes emails in batches to avoid timeout issues
- Duplicate prevention: Avoids saving duplicate files
- Multi-user support: Processes emails for multiple users
- Robust error handling: Complete try/catch blocks with detailed logging
- Oldest-first processing: By default, processes emails from oldest to newest

## Configuration

The script can be customized by editing the `CONFIG` object at the top of the file:

```javascript
const CONFIG = {
  mainFolderId: "YOUR_SHARED_FOLDER_ID", // Replace with your folder ID
  processedLabelName: "GDrive_Processed", // Label for processed threads
  skipDomains: ["example.com", "noreply.com"], // Skip these domains
  batchSize: 10, // Number of threads to process at once
  skipSmallImages: true, // Skip small images like signatures
  // ... other configuration options
};
```

## For more information

See the [complete repository](https://github.com/petalo/save-attachments-from-gmail-to-gdrive) for more detailed documentation and advanced debugging features.
