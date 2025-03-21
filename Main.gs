/**
 * Main module for Gmail Attachment Organizer
 *
 * This module provides the entry point functions for the application,
 * including the main function to save attachments and trigger creation.
 */

/**
 * Main function that processes Gmail attachments for authorized users
 *
 * This function:
 * 1. Acquires an execution lock to prevent concurrent runs
 * 2. Gets the list of authorized users
 * 3. Processes emails for each user
 * 4. Logs completion
 *
 * @return {boolean} True if processing completed successfully
 */
function saveAttachmentsToDrive() {
  try {
    logWithUser("Starting attachment processing", "INFO");

    // Acquire lock to prevent concurrent executions
    if (!acquireExecutionLock()) {
      logWithUser("Another instance is already running. Exiting.", "WARNING");
      return false;
    }

    // Get authorized users
    const users = getAuthorizedUsers();

    if (!users || users.length === 0) {
      logWithUser("No authorized users found. Exiting.", "WARNING");
      return false;
    }

    logWithUser(`Processing attachments for ${users.length} users`, "INFO");

    // Process each user
    for (const user of users) {
      processUserEmails(user);
    }

    logWithUser("Attachment processing completed successfully", "INFO");
    return true;
  } catch (error) {
    logWithUser(`Error in saveAttachmentsToDrive: ${error.message}`, "ERROR");
    return false;
  }
}

/**
 * Creates a time-based trigger to run saveAttachmentsToDrive every 15 minutes
 * If a trigger already exists, it will be deleted first
 *
 * @return {Trigger} The created trigger
 */
function createTrigger() {
  try {
    // Delete existing triggers
    const triggers = ScriptApp.getProjectTriggers();
    for (const trigger of triggers) {
      if (trigger.getHandlerFunction() === "saveAttachmentsToDrive") {
        ScriptApp.deleteTrigger(trigger);
        logWithUser("Deleted existing trigger", "INFO");
      }
    }

    // Create new trigger
    const trigger = ScriptApp.newTrigger("saveAttachmentsToDrive")
      .timeBased()
      .everyMinutes(15)
      .create();

    logWithUser("Created new trigger to run every 15 minutes", "INFO");
    return trigger;
  } catch (error) {
    logWithUser(`Error creating trigger: ${error.message}`, "ERROR");
    throw error;
  }
}
