/**
 * Main module for Gmail Attachment Organizer
 *
 * This module provides the entry point functions for the application,
 * including the main function to save attachments and trigger creation.
 */

/**
 * Validates the configuration settings to ensure all required values are properly set
 *
 * @return {boolean} True if the configuration is valid, throws an error otherwise
 */
function validateConfig() {
  // Check essential folder and label settings
  if (!CONFIG.mainFolderId || CONFIG.mainFolderId === "__FOLDER_ID__") {
    throw new Error(
      "Configuration error: mainFolderId is not set. Please set a valid Google Drive folder ID."
    );
  }

  // Validate AI configuration if enabled
  if (CONFIG.invoiceDetection === "gemini") {
    if (
      CONFIG.geminiApiKey === "__GEMINI_API_KEY__" &&
      !PropertiesService.getScriptProperties().getProperty(
        CONFIG.geminiApiKeyPropertyName
      )
    ) {
      throw new Error(
        "Configuration error: Gemini API key is not set but Gemini invoice detection is enabled."
      );
    }
  } else if (CONFIG.invoiceDetection === "openai") {
    if (
      CONFIG.openAIApiKey === "__OPENAI_API_KEY__" &&
      !PropertiesService.getScriptProperties().getProperty(
        CONFIG.openAIApiKeyPropertyName
      )
    ) {
      throw new Error(
        "Configuration error: OpenAI API key is not set but OpenAI invoice detection is enabled."
      );
    }
  }

  // Validate numerical values are within acceptable ranges
  if (CONFIG.aiConfidenceThreshold < 0 || CONFIG.aiConfidenceThreshold > 1) {
    throw new Error(
      "Configuration error: aiConfidenceThreshold must be between 0 and 1."
    );
  }

  if (CONFIG.batchSize < 1) {
    throw new Error("Configuration error: batchSize must be at least 1.");
  }

  if (CONFIG.triggerIntervalMinutes < 1) {
    throw new Error(
      "Configuration error: triggerIntervalMinutes must be at least 1."
    );
  }

  // Validate interdependent settings
  if (CONFIG.onlyAnalyzePDFs && !CONFIG.invoiceFileTypes.includes(".pdf")) {
    throw new Error(
      "Configuration error: onlyAnalyzePDFs is enabled but .pdf is not in invoiceFileTypes."
    );
  }

  return true;
}

/**
 * Main function that processes Gmail attachments for authorized users
 *
 * This function:
 * 1. Validates the configuration
 * 2. Acquires an execution lock to prevent concurrent runs
 * 3. Gets the list of authorized users
 * 4. Processes emails for each user
 * 5. Logs completion
 *
 * @return {boolean} True if processing completed successfully, false if an error occurred or if no users are authorized.
 * The function follows a structured flow:
 * - It first validates the configuration to ensure all required settings are properly set.
 * - It attempts to acquire a lock to ensure no concurrent executions.
 * - It retrieves the list of authorized users.
 * - For each user, it processes their emails to save attachments to Google Drive.
 * - Logs are generated throughout the process to provide detailed information on the execution status.
 */
function saveAttachmentsToDrive() {
  try {
    logWithUser("Starting attachment processing", "INFO");

    // Validate configuration
    validateConfig();
    logWithUser("Configuration validated successfully", "INFO");

    // Acquire lock to prevent concurrent executions
    const currentUser = Session.getEffectiveUser().getEmail();
    if (!acquireExecutionLock(currentUser)) {
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
 * Creates a time-based trigger to run saveAttachmentsToDrive at specified intervals
 * If a trigger already exists, it will be deleted first
 *
 * @return {Trigger} The created trigger object, which represents the scheduled execution of the function.
 * The function performs the following steps:
 * - Deletes any existing triggers for the saveAttachmentsToDrive function to avoid duplicates.
 * - Creates a new time-based trigger using the interval specified in the CONFIG object.
 * - Logs the creation of the new trigger and returns the trigger object.
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
      .everyMinutes(CONFIG.triggerIntervalMinutes)
      .create();

    logWithUser(
      `Created new trigger to run every ${CONFIG.triggerIntervalMinutes} minutes`,
      "INFO"
    );
    return trigger;
  } catch (error) {
    logWithUser(`Error creating trigger: ${error.message}`, "ERROR");
    throw error;
  }
}
