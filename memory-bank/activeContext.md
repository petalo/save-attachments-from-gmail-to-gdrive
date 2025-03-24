# Active Context: Gmail Attachment Organizer

## Current Work Focus

The Gmail Attachment Organizer project is currently focused on enhancing its invoice detection capabilities through AI integration. The focus is on:

1. **AI-Powered Invoice Detection**: Implementing both Google Gemini and OpenAI for improved invoice identification
2. **Privacy-First AI Integration**: Ensuring minimal data exposure while maximizing detection accuracy
3. **Flexible AI Configuration**: Providing options to choose between AI providers or disable AI completely
4. **Documentation and Onboarding**: Maintaining comprehensive documentation for both users and developers

## Recent Changes

### 1. Deployment Improvements

- Added support for production and test environments via `.env.prod` and `.env.test` files
- Enhanced build process with environment selection via `--env` parameter
- Implemented dynamic `.clasp.json` generation based on environment
- Created a single-file version for easier installation

### 2. Attachment Filtering Enhancements

- Improved detection of embedded images vs. real attachments
- Added MIME type whitelist for important document types
- Enhanced size-based filtering for small images

### 3. Multi-User Support

- Implemented user queue for processing multiple users
- Added permission verification before processing
- Created user management functions for adding/removing users

### 4. Documentation Updates

- Created detailed installation guide with screenshots
- Enhanced README with comprehensive feature descriptions
- Added explanations for thread labeling and attachment filtering systems

## Next Steps

### 1. Short-Term Priorities

- **Testing**: Comprehensive testing across different Gmail account types and email volumes
- **User Feedback**: Gather feedback from initial users to identify pain points
- **Performance Optimization**: Identify and address any performance bottlenecks

### 2. Medium-Term Improvements

- **Enhanced Logging**: Implement more detailed logging for troubleshooting
- **User Interface**: Consider adding a simple UI for configuration and monitoring
- **Alternative Organization Schemes**: Support additional organization methods beyond domain-based folders

### 3. Long-Term Vision

- **Message-Level Processing**: Consider implementing message-level tracking instead of thread-level
- **Advanced Filtering**: Add more sophisticated attachment filtering options
- **Integration with Other Services**: Explore integration with other Google services or external platforms

## Active Decisions and Considerations

### 1. Thread vs. Message Processing

**Current Decision**: Process at the thread level for simplicity and efficiency
**Considerations**:

- Thread-level processing is more efficient for Gmail API usage
- New messages in processed threads won't be automatically processed
- Users can manually remove labels to force reprocessing

### 2. Attachment Filtering Strategy

**Current Decision**: Use a multi-faceted approach combining MIME types, size, and pattern recognition
**Considerations**:

- Balance between catching all legitimate attachments and filtering out embedded content
- Different email clients embed content in different ways
- Need to avoid false positives (skipping real attachments)

### 3. Folder Organization Structure

**Current Decision**: Organize by sender domain
**Considerations**:

- Domain-based organization provides a logical structure
- Could consider additional organization levels (date, attachment type)
- Balance between too flat and too deep folder structures

### 4. Execution Frequency and Batch Size

**Current Decision**: Default 15-minute intervals with 10 threads per batch
**Considerations**:

- More frequent execution provides quicker processing but uses more resources
- Larger batch sizes process more emails but risk hitting execution time limits
- Need to balance responsiveness with resource usage

### 5. Deployment Strategy

**Current Decision**: Support multiple deployment options (single-file, modular, build-based) with environment-specific configurations
**Considerations**:

- Single-file version is easier for non-technical users
- Modular version is better for development and maintenance
- Build-based version allows for environment-specific configuration
- Production and test environments can use different:
  - Google Drive folders
  - Google Apps Script projects
  - Gmail labels for processed threads
- **Important**: Avoid associating scripts with Google Cloud Platform (GCP) projects, as this can cause permission issues

## Current Implementation Status

| Component             | Status   | Notes                                  |
| --------------------- | -------- | -------------------------------------- |
| Core Email Processing | Complete | Fully functional with batch processing |
| Attachment Filtering  | Complete | Multi-faceted approach implemented     |
| Folder Management     | Complete | Domain-based organization working      |
| User Management       | Complete | Multi-user support with queue          |
| Error Handling        | Complete | Comprehensive try/catch with logging   |
| Documentation         | Complete | Installation guide and README          |
| Build Process         | Complete | Environment-based configuration        |
| Deployment Options    | Complete | Single-file and modular versions       |

## Open Questions and Challenges

1. **Handling New Messages in Processed Threads**: What's the best approach for users who frequently receive important new attachments in existing threads?

2. **Performance with Large Email Volumes**: How does the script perform with very large numbers of emails and attachments?

3. **Folder Structure Scalability**: How well does the domain-based folder structure scale with many different domains?

4. **User Experience for Configuration**: Is the current configuration approach (editing Config.gs) user-friendly enough?

5. **Error Notification**: Should we implement a notification system for errors rather than just logging?
