# Project Brief: Gmail Attachment Organizer for Google Drive

## Project Overview

The Gmail Attachment Organizer is a Google Apps Script application that automatically saves Gmail attachments to Google Drive, organizing them by the domain of the sender's email address. It creates a folder structure where each sender's domain gets its own folder, making it easy to track and organize attachments based on their source.

## Core Requirements

1. **Automated Attachment Saving**
   - Automatically detect and save attachments from Gmail to Google Drive
   - Preserve original email timestamps in file descriptions
   - Process emails in batches to avoid timeout issues

2. **Intelligent Organization**
   - Create a folder structure based on sender domains
   - Skip specified domains that should not be processed
   - Filter out unwanted attachments (embedded images, signatures, etc.)

3. **Multi-User Support**
   - Process emails for multiple authorized users
   - Maintain a user queue for processing
   - Verify user permissions for Gmail and Drive access

4. **Robust Error Handling**
   - Implement comprehensive try/catch blocks with logging
   - Prevent duplicate file saving
   - Include retry logic with exponential backoff for transient errors

5. **Flexible Deployment Options**
   - Provide a single-file version for simple installation
   - Offer a modular version for better maintenance and development
   - Support environment-based configuration for different deployment scenarios

## Project Goals

1. **Efficiency**
   - Minimize manual effort in organizing email attachments
   - Reduce time spent searching for specific attachments
   - Process emails in a timely manner without hitting Google's API limits

2. **Organization**
   - Create a logical, domain-based folder structure
   - Ensure attachments are easily findable based on their source
   - Maintain clean separation between different sender domains

3. **Reliability**
   - Ensure all valid attachments are properly saved
   - Prevent duplicate processing of emails
   - Handle errors gracefully without script termination

4. **Usability**
   - Make installation and configuration straightforward
   - Provide clear documentation for setup and customization
   - Support different user preferences through configuration options

5. **Maintainability**
   - Structure code in a modular, maintainable way
   - Include comprehensive documentation
   - Support ongoing development and enhancement

## Success Criteria

1. Attachments from Gmail are automatically saved to Google Drive
2. Files are organized in folders based on sender domains
3. Only legitimate attachments are saved (embedded images and signatures are filtered out)
4. The script runs reliably on a schedule without manual intervention
5. Multiple users can benefit from the script with proper authorization
6. The solution is easy to install, configure, and maintain
