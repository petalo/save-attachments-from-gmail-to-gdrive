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

  if (
    !CONFIG.processedLabelName ||
    !CONFIG.processingLabelName ||
    !CONFIG.errorLabelName ||
    !CONFIG.permanentErrorLabelName ||
    !CONFIG.tooLargeLabelName
  ) {
    throw new Error(
      "Configuration error: processedLabelName, processingLabelName, errorLabelName, permanentErrorLabelName, and tooLargeLabelName must all be set."
    );
  }

  const labelNames = [
    CONFIG.processedLabelName,
    CONFIG.processingLabelName,
    CONFIG.errorLabelName,
    CONFIG.permanentErrorLabelName,
    CONFIG.tooLargeLabelName,
  ];
  if (
    new Set(labelNames).size !== labelNames.length
  ) {
    throw new Error(
      "Configuration error: processedLabelName, processingLabelName, errorLabelName, permanentErrorLabelName, and tooLargeLabelName must be different."
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

  if (CONFIG.executionSoftLimitMs < 1000) {
    throw new Error(
      "Configuration error: executionSoftLimitMs must be at least 1000 milliseconds."
    );
  }

  if (CONFIG.processingStateTtlMinutes < 1) {
    throw new Error(
      "Configuration error: processingStateTtlMinutes must be at least 1 minute."
    );
  }

  if (CONFIG.staleRecoveryBatchSize < 1) {
    throw new Error(
      "Configuration error: staleRecoveryBatchSize must be at least 1."
    );
  }

  if (CONFIG.maxThreadFailureRetries < 1) {
    throw new Error(
      "Configuration error: maxThreadFailureRetries must be at least 1."
    );
  }

  if (CONFIG.threadFailureStateTtlDays < 1) {
    throw new Error(
      "Configuration error: threadFailureStateTtlDays must be at least 1."
    );
  }

  const allowedLogLevels = ["DEBUG", "INFO", "WARNING", "ERROR"];
  if (!allowedLogLevels.includes(String(CONFIG.logLevel || "").toUpperCase())) {
    throw new Error(
      "Configuration error: logLevel must be one of DEBUG, INFO, WARNING, ERROR."
    );
  }

  if (CONFIG.executionModel !== "effective_user_only") {
    throw new Error(
      "Configuration error: executionModel must be \"effective_user_only\"."
    );
  }

  return true;
}

/**
 * Main function that processes Gmail attachments for the effective user
 *
 * This function:
 * 1. Validates the configuration
 * 2. Acquires an execution lock to prevent concurrent runs
 * 3. Computes a safe execution deadline
 * 4. Processes emails for the current execution user
 * 5. Logs completion
 *
 * @return {boolean} True if processing completed successfully, false if an error occurred.
 * The function follows a structured flow:
 * - It first validates the configuration to ensure all required settings are properly set.
 * - It attempts to acquire a lock to ensure no concurrent executions.
 * - It computes a soft execution deadline to avoid hard timeouts.
 * - It processes the current user's emails to save attachments to Google Drive.
 * - Logs are generated throughout the process to provide detailed information on the execution status.
 */
function saveAttachmentsToDrive() {
  let currentUser = null;
  let lockAcquired = false;
  const executionStartMs = new Date().getTime();

  try {
    logWithUser("Starting attachment processing", "INFO");

    // Validate configuration
    validateConfig();
    logWithUser("Configuration validated successfully", "INFO");

    // Acquire lock to prevent concurrent executions
    currentUser = Session.getEffectiveUser().getEmail();
    if (!acquireExecutionLock(currentUser)) {
      logWithUser("Another instance is already running. Exiting.", "WARNING");
      return false;
    }
    lockAcquired = true;

    // Process only the effective user for this execution.
    // GmailApp runs in the current execution context, so iterating over a
    // registered user list would reprocess the same mailbox.
    const deadlineMs = executionStartMs + CONFIG.executionSoftLimitMs;
    recoverStaleProcessingThreads(currentUser, deadlineMs);
    logWithUser(`Processing attachments for current user: ${currentUser}`, "INFO");
    const processed = processUserEmails(currentUser, true, deadlineMs);
    if (!processed) {
      logWithUser(`Processing failed for user: ${currentUser}`, "ERROR");
      return false;
    }

    logWithUser("Attachment processing completed successfully", "INFO");
    return true;
  } catch (error) {
    logWithUser(`Error in saveAttachmentsToDrive: ${error.message}`, "ERROR");
    return false;
  } finally {
    if (lockAcquired) {
      try {
        releaseExecutionLock(currentUser);
      } catch (releaseError) {
        logWithUser(
          `Error releasing execution lock: ${releaseError.message}`,
          "WARNING"
        );
      }
    }
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
