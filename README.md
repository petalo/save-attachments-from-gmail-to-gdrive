# Gmail Attachment Organizer for Google Drive

This Google Apps Script automatically saves Gmail attachments to Google Drive, organized by the domain of the sender's email address. It creates a folder structure where each sender's domain gets its own folder, making it easy to track and organize attachments based on their source.

## Installation Options

> **IMPORTANT**: This script offers two installation options, choose the one that best suits your needs:

### Option 1: Simple Installation (Single File)

For a detailed step-by-step guide, refer to the [Installation Steps](./INSTALLATION_STEPS.md).

If you want the quickest way to implement the script:

1. Create a new project in Google Apps Script at [script.google.com](https://script.google.com)
2. Copy all the content from the `single-file/Code.gs` file and paste it into the editor
3. Replace `YOUR_SHARED_FOLDER_ID` with your Google Drive folder ID
4. Save and run the `saveAttachmentsToDrive` function

This option is ideal for users who want a quick installation without worrying about multiple files.

### Option 2: Modular Installation (Multiple Files)

If you prefer a more organized structure for long-term maintenance:

1. Create a new project in Google Apps Script at [script.google.com](https://script.google.com)
2. Create the following files in your project:
   - `Config.gs`
   - `Utils.gs`
   - `UserManagement.gs`
   - `AttachmentFilters.gs`
   - `FolderManagement.gs`
   - `AttachmentProcessing.gs`
   - `GmailProcessing.gs`
   - `Main.gs`
   - `Debug.gs` (optional)
   - `appsscript.json`
3. Copy the content of each file from the `modular/` directory
4. Update the `mainFolderId` in `Config.gs` with your Google Drive folder ID
5. Save and run the `saveAttachmentsToDrive` function

This option is better for ongoing development and maintenance.

## Features

- **Domain-Based Organization**: Automatically creates folders based on sender domains
- **Smart Attachment Filtering**: Identifies and skips embedded images and email signatures
- **Batch Processing**: Processes emails in batches to avoid timeout issues
- **Duplicate Prevention**: Prevents saving duplicate files
- **Multi-User Support**: Processes emails for multiple users and maintains a user queue
- **Robust Error Handling**: Comprehensive try/catch blocks with logging
- **Oldest-First Processing**: Processes emails from oldest to newest by default (configurable)
- **Email Date Preservation**: Stores original email date in file descriptions

## Execution Flow

1. **Script Initialization**:
   - Validates configuration settings
   - Obtains reference to the main Google Drive folder
   - Gets or creates the Gmail label for marking processed emails
   - Gets the next user in the queue to process

2. **User Processing**:
   - Verifies user permissions for Gmail and Drive access
   - Switches to user context for processing
   - Maintains a queue of users to process in turn

3. **Email Discovery**:
   - Searches Gmail for unprocessed emails with attachments
   - Processes emails from oldest to newest by default (configurable)
   - Limits processing to a configurable batch size to prevent timeouts

4. **Attachment Processing**:
   - For each email thread:
     - Extracts all messages in the thread
     - For each message with attachments:
       - Extracts the sender's domain from their email address
       - Filters out unwanted attachments (small images, calendar invitations)
       - Creates domain-specific folder only when necessary
       - For each valid attachment:
         - Ensures filename uniqueness
         - Saves the attachment to the appropriate folder
     - Marks the thread as processed by applying the Gmail label

5. **Error Handling**:
   - Comprehensive try/catch blocks at multiple levels
   - Detailed logging for troubleshooting
   - Graceful failure handling to prevent script termination on single-item errors
   - Retry logic with exponential backoff for transient errors

## Thread Labeling System

The script uses labels at the thread level to track processed emails:

1. **Label Application**:
   - The "GDrive_Processed" label is applied to entire Gmail threads, not individual messages
   - A thread is labeled when:
     a) At least one valid attachment is saved from any message in the thread, OR
     b) All messages in the thread have attachments, but they were all filtered out (e.g., embedded images)
   - The script never labels threads that have no attachments at all

2. **Handling New Messages in Processed Threads**:
   - New messages added to a thread that already has the "GDrive_Processed" label will NOT be processed automatically by this script
   - If a new message with an important attachment is added to a previously processed thread:
     a) You can manually remove the "GDrive_Processed" label from the thread to force reprocessing in the next script run
     b) Forward the message to yourself in a new thread (creating a new thread without the label)

3. **Rationale for Thread-Level Labeling**:
   - Gmail's API is optimized for thread-level operations
   - Most email conversations maintain context in a thread
   - Reduces processing overhead by avoiding repeated analysis of related messages
   - Prevents running into quota limits for Gmail API calls

If you frequently receive important new attachments in existing threads, consider one of these options:

1. Adjust email sending/replying behavior to create new threads for important attachments
2. Implement a more sophisticated Message-ID based tracking system (would require significant code changes)
3. Run the script more frequently and maintain a separate record of processed message IDs

## Attachment Filtering System

The script uses a sophisticated filtering system to distinguish between real attachments and embedded elements like email signatures, logos, and inline images:

1. **MIME Type Whitelist**:
   - A list of MIME types (attachmentTypesWhitelist) that are always saved (e.g., PDF, Word, Excel)
   - If an attachment has a MIME type on this list, it is always kept regardless of other filters
   - If a file doesn't match the whitelist, it will be evaluated by the additional criteria below

2. **Content-Disposition Analysis**:
   - Checks if an image is specifically marked as "inline" in its content-disposition header
   - Inline images are typically embedded in the email body rather than conscious attachments

3. **HTML Email Image URL Detection**:
   - Identifies images referenced by URLs in HTML emails from various email providers
   - Detects Gmail embedded image URLs with parameters like "view=fimg" or "disp=emb"
   - Recognizes Outlook, Yahoo Mail, and other common email image URL patterns
   - Filters out image URLs that use common patterns like "cid=" or specific domains

4. **Filename Pattern Recognition**:
   - Detects common patterns for embedded images (e.g., "image001.png", "inline-", "Outlook-")
   - Checks against a list of common embedded element names (logos, icons, banners, etc.)

5. **File Extension Filtering**:
   - Special handling for common document extensions (.pdf, .doc, .docx, etc.) - always kept
   - Option to filter unwanted file types (e.g., calendar invitations - .ics files)
   - Option to filter small images that are likely signatures or icons

6. **Size-Based Filtering**:
   - Files without extensions under a certain size (e.g., 50KB) are likely embedded content
   - Configuration to skip images below a certain size threshold

These combined approaches ensure that only legitimate attachments are saved while filtering out embedded elements that would clutter your Drive storage.

## Required Permissions

This script requires the following authorization scopes:

- <https://www.googleapis.com/auth/gmail.modify> (for reading emails and applying labels)
- <https://www.googleapis.com/auth/drive> (for creating folders and files in Google Drive)
- <https://www.googleapis.com/auth/script.scriptapp> (for creating triggers)
- <https://www.googleapis.com/auth/script.external_request> (for external API calls if needed)

## Managing Users
- To remove users: Use `removeUserFromList('email@domain.com')`
- To check user permissions: Run `verifyUserPermissions('
- To add users manually: Use `addUserToList('email@domain.com')`
- To remove users: Use `removeUserFromList('email@domain.com')`
- To view current users: Run `listUsers()`

## Configuration Options

The `Config.gs` file contains all configurable options:

- `mainFolderId`: ID of the main Google Drive folder
- `processedLabelName`: Name of the Gmail label to apply to processed messages
- `skipDomains`: Array of email domains to skip processing
- `skipSmallImages`: Set to true to avoid saving small images like email signatures
- `smallImageMaxSize`: Maximum size in bytes for images to be skipped (default: 20KB)
- `smallImageExtensions`: File extensions to consider as images for filtering
- `skipFileTypes`: Additional file types to skip (e.g., calendar invitations, etc.)
- `attachmentTypesWhitelist`: List of MIME types that should always be saved
- `batchSize`: Number of threads to process in each execution (default: 20)

## Advanced Usage

### Custom Processing Options

The main function `saveAttachmentsToDrive()` supports optional parameters:

```javascript
// Process with default options (oldest emails first)
saveAttachmentsToDrive();

// Process newest emails first
saveAttachmentsToDrive({ oldestFirst: false });
```

By default, the script processes emails from oldest to newest, which helps ensure older attachments are saved first. This behavior can be changed if needed.

### Email Timestamps for Files

The script preserves the original email timestamps by storing the email date in the file's description. Due to Google Drive API limitations, it's not possible to directly set arbitrary creation dates for files. Instead, the script:

1. Stores the original email date in the file's description as `original_date=YYYY-MM-DDTHH:MM:SS.SSSZ`
2. For text files, adds an invisible comment with the timestamp
3. Uses file recreation when necessary to preserve content and metadata

This approach allows you to:

- See when the email was actually received
- Sort or filter files by the original date using the description field
- Maintain the relationship between files and their source emails

The timestamp functionality is optimized for performance, avoiding unnecessary API calls that would slow down processing.

```javascript
// In Config.gs
useEmailTimestamps: true, // Set to true to include email timestamps in file descriptions
```

## Performance Considerations

The script includes several optimizations to improve performance:

1. **Efficient Timestamp Handling**: Uses direct file creation with descriptive metadata rather than multiple API attempts
2. **Batch Processing**: Processes emails in configurable batch sizes to prevent timeouts
3. **Smart Logging**: Reduces excessive logging for better performance
4. **Duplicate Detection**: Uses efficient size-based comparison to identify duplicates

If you're processing large volumes of emails, consider:

- Increasing the batch size if your script has enough execution time available
- Running the script more frequently with smaller batch sizes
- Adjusting logging verbosity in the configuration

## Customizing Attachment Filtering

To adjust what files are saved or skipped:

1. To always save certain MIME types: Add the MIME type to `attachmentTypesWhitelist`
   - For example, to always save PNG files: add "image/png" to the list

2. To skip specific file types: Add extensions to `skipFileTypes`
   - For example, to skip all text files: add ".txt" to the list

3. To adjust image filtering: Modify `smallImageMaxSize` and `smallImageExtensions`
   - A larger size threshold will skip more images, a smaller one will save more

4. If legitimate files are being skipped: Check the logs to identify why they're being filtered out and adjust the appropriate settings

## File Structure

- `Config.gs`: Configuration settings and constants
- `Utils.gs`: Utility functions for logging, retries, and user settings
- `UserManagement.gs`: User authorization and permission management
- `AttachmentFilters.gs`: Functions for determining which attachments to process
- `FolderManagement.gs`: Google Drive folder creation and management
- `AttachmentProcessing.gs`: Functions for saving attachments to Drive
- `GmailProcessing.gs`: Gmail thread and message processing
- `Main.gs`: Entry points and main execution flow
- `appsscript.json`: Script manifest with required OAuth scopes

## Troubleshooting

- **Permissions Issues**: Run `verifyAllUsersPermissions()` to check for permission problems
- **Missing Attachments**: Check the execution logs for filtering reasons
- **Script Timeout**: Reduce the `batchSize` in the configuration
- **Duplicate Folders**: This can happen if multiple executions occur simultaneously; the script includes lock logic to handle this case
- **Thread Processing**: If a thread is processed but you can't find attachments, check if they were filtered out due to being small images or embedded content

## License
[MIT License](./LICENSE.md)

