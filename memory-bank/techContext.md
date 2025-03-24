# Technical Context: Gmail Attachment Organizer

## Technologies Used

### Core Technologies

1. **Google Apps Script (GAS)**
   - Server-side JavaScript runtime environment
   - Provides direct integration with Google services
   - Runs in Google's cloud infrastructure
   - Supports time-based triggers for automation

2. **Gmail API**
   - Used to search, read, and label emails
   - Provides access to email metadata and attachments
   - Supports thread-based operations

3. **Google Drive API**
   - Used to create folders and save files
   - Manages file metadata and organization
   - Handles file deduplication and versioning

4. **JavaScript (ECMAScript)**
   - Primary programming language for Google Apps Script
   - ES6+ features with some limitations
   - Synchronous execution model in GAS

5. **AI Services**
   - Google Gemini API for privacy-focused invoice detection
   - OpenAI API for alternative invoice detection
   - Pattern recognition for improved accuracy

### Development Tools

1. **clasp (Command Line Apps Script Projects)**
   - CLI tool for local development of Apps Script projects
   - Enables version control integration
   - Supports pushing/pulling code to/from Google Apps Script

2. **Node.js**
   - Used for build scripts and local development
   - Manages dependencies via npm
   - Runs the file combination script for deployment

3. **dotenv**
   - Manages environment variables for local development
   - Allows configuration of folder IDs without hardcoding

## Development Setup

### Local Development Environment

1. **Repository Structure**

   ```text
   save-attachments-from-gmail-to-gdrive/
   ├── .clasp.json          # clasp configuration
   ├── .claspignore         # Files to ignore when pushing
   ├── .env.example         # Example environment variables
   ├── .gitignore           # Git ignore file
   ├── appsscript.json      # Apps Script manifest
   ├── Code.js              # Combined script output
   ├── combine-files.js     # Build script
   ├── package.json         # npm package configuration
   ├── README.md            # Documentation
   ├── INSTALLATION_STEPS.md # Installation guide
   ├── LICENSE.md           # License information
   ├── assets/              # Documentation images
   ├── build/               # Build output directory
   ├── logs/                # Test logs directory
   ├── memory-bank/         # Documentation and context
   ├── single-file/         # Single-file version
   └── src/                 # Source code directory
       ├── appsscript.json  # Apps Script manifest
       ├── AttachmentFilters.gs
       ├── AttachmentProcessing.gs
       ├── Config.gs
       ├── Debug.gs
       ├── FolderManagement.gs
       ├── GeminiDetection.gs # Gemini AI integration
       ├── GmailProcessing.gs
       ├── HistoricalPatterns.gs # Pattern analysis for AI
       ├── Main.gs
       ├── OpenAIDetection.gs # OpenAI integration
       ├── README.md
       ├── UserManagement.gs
       └── Utils.gs
   ```

2. **Build Process**
   - `combine-files.js` script merges modular source files into:
     - A single `Code.js` file for direct deployment
     - A version in the `build/` directory with environment variables applied
     - A version in the `single-file/` directory for manual copy-paste installation
   - Supports environment selection via `--env` parameter:
     - Production environment (default): `--env=prod`
     - Test environment: `--env=test`
   - Dynamically generates `.clasp.json` based on selected environment

3. **Deployment Process**
   - Local development using clasp
   - Push to Google Apps Script using `npm run deploy`
   - Version management using `npm run version`
   - Manual installation option via copy-paste from `single-file/Code.gs`

## Technical Constraints

### Google Apps Script Limitations

1. **Execution Time Limits**
   - 6-minute maximum execution time per run
   - Necessitates batch processing approach
   - Requires careful management of API calls

2. **Quota Limits**
   - Gmail API has daily quotas for read operations
   - Drive API has quotas for file creation
   - Script properties have size limitations

3. **Authorization Scopes**
   - Requires specific OAuth scopes:
     - `https://www.googleapis.com/auth/gmail.modify` - For reading and labeling emails
     - `https://www.googleapis.com/auth/drive` - For saving attachments to Drive
     - `https://www.googleapis.com/auth/script.scriptapp` - For managing triggers
     - `https://www.googleapis.com/auth/script.external_request` - For AI API calls
   - Users must explicitly grant these permissions
   - The external request scope is specifically needed for AI integration

4. **JavaScript Environment**
   - Limited ES6+ support
   - No direct file system access
   - No external npm dependencies in runtime
   - Synchronous execution model

### Gmail-Specific Constraints

1. **Thread-Based Operations**
   - Gmail organizes emails into threads
   - Labels apply to entire threads, not individual messages
   - New messages in labeled threads won't be automatically processed

2. **Attachment Handling**
   - Embedded images often appear as attachments
   - Content-Disposition headers may be inconsistent
   - Size limits on attachments (25MB)

### Google Drive Constraints

1. **File Metadata Limitations**
   - Cannot directly set arbitrary creation dates
   - Limited custom metadata options
   - File descriptions used for storing email timestamps

2. **Folder Structure**
   - No built-in way to enforce unique folder names
   - Need custom logic to prevent duplicate folders
   - Folder operations can be slow with many items

## Dependencies

### Runtime Dependencies

1. **Google Services**
   - GmailApp - Core Gmail service
   - DriveApp - Core Drive service
   - PropertiesService - For storing script state and API keys
   - ScriptApp - For managing triggers
   - Logger - For logging and debugging
   - UrlFetchApp - For making external API calls

2. **External API Services**
   - Google Gemini API - For AI-powered invoice detection
   - OpenAI API - Alternative AI provider for invoice detection
   - Both require API keys stored securely in Script Properties

### Development Dependencies

1. **@google/clasp (v2.4.2)**
   - Command-line tool for Apps Script development
   - Enables local development and version control
   - Handles authentication with Google

2. **dotenv (v16.4.7)**
   - Loads environment variables from .env file
   - Used during build process to configure folder IDs
   - Not used at runtime

3. **googleapis (v148.0.0)**
   - Used for local testing and development
   - Not used in the deployed script

## Configuration Management

1. **Environment Variables**
   - Environment-specific configuration files:
     - `.env.prod` for production environment (mandatory)
     - `.env.test` for test environment (optional)
   - Contains sensitive information like folder IDs, script IDs, and API keys
   - Includes environment-specific settings like `PROCESSED_LABEL_NAME`
   - Not committed to version control

2. **Configuration Object**
   - Centralized `CONFIG` object in `Config.gs`
   - Contains all adjustable settings including AI configuration
   - Documented with comments for each option
   - AI-specific settings include:
     - `invoiceDetection`: Selects between "gemini", "openai", or false
     - API keys and model settings for both AI providers
     - Shared settings like confidence threshold and domain exclusions

3. **Build-Time Configuration**
   - Environment variables are injected during build
   - Placeholders are replaced with actual values:
     - `__FOLDER_ID__` - Main folder ID
     - `__GEMINI_API_KEY__` - Google Gemini API key
     - `__OPENAI_API_KEY__` - OpenAI API key
   - Environment-specific values are also replaced:
     - `processedLabelName` - Custom label name for each environment
   - Dynamically generates `.clasp.json` with the correct `scriptId`
   - Enables different configurations for production and test environments
   - API keys can be provided at build time or stored in Script Properties

## Security Considerations

1. **Authentication and Authorization**
   - Uses Google's OAuth 2.0 for authentication
   - Requires explicit user consent for required scopes
   - Runs with the permissions of the authenticated user

2. **Data Handling**
   - Processes potentially sensitive email attachments
   - Original emails remain untouched in Gmail
   - Secure API key storage using Script Properties

3. **AI Privacy Protections**
   - Gemini implementation only sends metadata, not full content
   - OpenAI implementation truncates content to minimize exposure
   - Domain exclusion list prevents processing sensitive domains
   - PDF-only option reduces unnecessary AI processing
   - Historical pattern analysis reduces reliance on external AI

4. **User Management**
   - Support for multiple authorized users
   - Permission verification before processing
   - User-specific processing queue

5. **Error Handling**
   - Comprehensive logging for troubleshooting
   - Graceful failure handling
   - No exposure of sensitive information in error messages
