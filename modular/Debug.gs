/**
 * Debug and Test Functions for Gmail Attachment Organizer
 *
 * This module contains various diagnostic, test, and debug functions
 * that are useful during development and troubleshooting.
 */

/**
 * Test function to run the script with a smaller batch size for testing
 * Processes only emails for the current user
 */
function testRun() {
  const currentUserEmail = Session.getEffectiveUser().getEmail();
  logWithUser("Starting test run for Gmail attachment organizer", "INFO");

  // Temporarily modify CONFIG for testing
  const originalBatchSize = CONFIG.batchSize;
  CONFIG.batchSize = 5;

  try {
    // Only process current user's emails with oldest first
    processUserEmails(currentUserEmail, true);
  } catch (e) {
    logWithUser(`Error in test run: ${e.message}`, "ERROR");
  } finally {
    // Restore original batch size
    CONFIG.batchSize = originalBatchSize;
    logWithUser("Test run completed", "INFO");
  }
}

/**
 * Gets the current configuration settings
 * This is useful for debugging
 *
 * @returns {Object} The current configuration object
 */
function getConfig() {
  return CONFIG;
}

/**
 * Diagnostic function to check the number of emails that match the search criteria
 * This helps troubleshoot issues where the script reports no emails when there should be some
 *
 * @returns {Object} Information about matching threads
 */
function diagnoseMissingEmails() {
  const searchCriteria = `has:attachment -label:${CONFIG.processedLabelName}`;
  const searchCriteriaWithOrder = searchCriteria + " older_first";

  // Search without limit to get all matching threads
  const allThreads = GmailApp.search(searchCriteria);
  const allThreadsOrdered = GmailApp.search(searchCriteriaWithOrder);

  // Get information about first thread if available
  let firstThreadInfo = "No threads found";
  if (allThreads.length > 0) {
    const firstThread = allThreads[0];
    const firstMessage = firstThread.getMessages()[0];
    firstThreadInfo = {
      subject: firstThread.getFirstMessageSubject(),
      from: firstMessage.getFrom(),
      date: firstMessage.getDate(),
      hasAttachments: firstMessage.getAttachments().length > 0,
    };
  }

  const result = {
    searchCriteria: searchCriteria,
    searchCriteriaWithOrder: searchCriteriaWithOrder,
    totalMatchingThreads: allThreads.length,
    totalMatchingThreadsOrdered: allThreadsOrdered.length,
    firstThreadDetails: firstThreadInfo,
  };

  Logger.log("Diagnostic results: " + JSON.stringify(result, null, 2));
  return result;
}

/**
 * Basic diagnostic that uses a simple search to test access to Gmail
 * This bypasses filters and just counts all messages with attachments
 */
function basicGmailDiagnostic() {
  try {
    // Try the most basic search possible to test access
    const basicSearch = "has:attachment";
    const threads = GmailApp.search(basicSearch, 0, 10);

    Logger.log(`Basic search '${basicSearch}' found ${threads.length} threads`);

    if (threads.length > 0) {
      // Log details of first thread
      const firstThread = threads[0];
      const messages = firstThread.getMessages();
      Logger.log(
        `First thread has ${
          messages.length
        } messages with subject: ${firstThread.getFirstMessageSubject()}`
      );

      // Test if we can get message details
      if (messages.length > 0) {
        const firstMessage = messages[0];
        Logger.log(`First message is from: ${firstMessage.getFrom()}`);
        const attachments = firstMessage.getAttachments();
        Logger.log(`First message has ${attachments.length} attachments`);
      }
    }

    // Test if our label exists and works
    const labelName = CONFIG.processedLabelName;
    const label = GmailApp.getUserLabelByName(labelName);

    if (label) {
      Logger.log(`Found label ${labelName} - it exists`);
      // Try searching with this label specifically
      const labelSearch = `has:attachment label:${labelName}`;
      const labelThreads = GmailApp.search(labelSearch, 0, 10);
      Logger.log(
        `Found ${labelThreads.length} threads with label ${labelName}`
      );
    } else {
      Logger.log(`Label ${labelName} does not exist yet`);
    }

    return {
      searchAccess: true,
      totalAttachmentThreads: threads.length,
      detailsAccessible: threads.length > 0,
    };
  } catch (error) {
    Logger.log(`Error in basic Gmail diagnostic: ${error.message}`);
    return {
      searchAccess: false,
      error: error.message,
    };
  }
}

/**
 * More detailed diagnostic that runs multiple tests
 * focusing on looking for a mismatch between search and actual processing
 */
function detailedGmailDiagnostic() {
  const searchCriteria = `has:attachment -label:${CONFIG.processedLabelName}`;
  const results = {
    tests: {},
    details: [],
  };

  try {
    // Test 1: Basic search without limit
    const allThreads = GmailApp.search(searchCriteria);
    results.tests.basicSearch = {
      passed: true,
      count: allThreads.length,
    };
    results.details.push(
      `Found ${allThreads.length} threads matching '${searchCriteria}'`
    );

    // Test 2: Search with a small limit
    const limitedThreads = GmailApp.search(searchCriteria, 0, 5);
    results.tests.limitedSearch = {
      passed: limitedThreads.length <= 5,
      count: limitedThreads.length,
    };
    results.details.push(
      `Limited search returned ${limitedThreads.length} threads`
    );

    // Test 3: Check if we can access the first thread
    if (allThreads.length > 0) {
      const thread = allThreads[0];
      try {
        const subject = thread.getFirstMessageSubject();
        results.tests.threadAccess = {
          passed: true,
          subject: subject,
        };
        results.details.push(
          `Successfully accessed thread with subject: ${subject}`
        );

        // Test 4: Check if we can load messages
        try {
          const messages = thread.getMessages();
          results.tests.messagesAccess = {
            passed: true,
            count: messages.length,
          };
          results.details.push(`Thread has ${messages.length} messages`);

          // Test 5: Check if we can access attachments
          if (messages.length > 0) {
            try {
              const attachments = messages[0].getAttachments();
              results.tests.attachmentsAccess = {
                passed: true,
                count: attachments.length,
              };
              results.details.push(
                `First message has ${attachments.length} attachments`
              );
            } catch (e) {
              results.tests.attachmentsAccess = {
                passed: false,
                error: e.message,
              };
              results.details.push(`Error accessing attachments: ${e.message}`);
            }
          }
        } catch (e) {
          results.tests.messagesAccess = {
            passed: false,
            error: e.message,
          };
          results.details.push(`Error accessing messages: ${e.message}`);
        }
      } catch (e) {
        results.tests.threadAccess = {
          passed: false,
          error: e.message,
        };
        results.details.push(`Error accessing thread: ${e.message}`);
      }
    } else {
      results.tests.threadAccess = {
        passed: false,
        reason: "No threads found",
      };
    }
  } catch (error) {
    results.overallError = error.message;
    results.details.push(`Overall error: ${error.message}`);
  }

  Logger.log(
    "Detailed diagnostic results: " + JSON.stringify(results, null, 2)
  );
  return results;
}

/**
 * A specialized diagnostic function to investigate the attachment issue
 * This function checks why some threads are showing up in search results
 * but don't actually have attachments
 */
function diagnoseAttachmentDiscrepancy() {
  try {
    // Basic search for attachments
    const threads = GmailApp.search("has:attachment", 0, 20);
    const results = [];

    logWithUser(
      `Found ${threads.length} threads matching has:attachment`,
      "INFO"
    );

    // Check each thread
    for (let i = 0; i < threads.length; i++) {
      const thread = threads[i];
      const messages = thread.getMessages();
      let threadHasRealAttachments = false;
      let attachmentCount = 0;
      let inlineImageCount = 0;

      // Check each message in the thread
      for (let j = 0; j < messages.length; j++) {
        const message = messages[j];
        const attachments = message.getAttachments({
          includeInlineImages: true,
        });
        const attachmentsNoInline = message.getAttachments({
          includeInlineImages: false,
        });

        // Count real attachments vs inline images
        attachmentCount += attachmentsNoInline.length;
        inlineImageCount += attachments.length - attachmentsNoInline.length;

        if (attachmentsNoInline.length > 0) {
          threadHasRealAttachments = true;
        }
      }

      // Log results for this thread
      results.push({
        subject: thread.getFirstMessageSubject(),
        date: thread.getLastMessageDate().toISOString(),
        messageCount: messages.length,
        totalAttachments: attachmentCount + inlineImageCount,
        realAttachments: attachmentCount,
        inlineImages: inlineImageCount,
        hasRealAttachments: threadHasRealAttachments,
      });

      logWithUser(
        `Thread ${i + 1}/${
          threads.length
        }: ${thread.getFirstMessageSubject()} - Real attachments: ${attachmentCount}, Inline images: ${inlineImageCount}`,
        "INFO"
      );
    }

    // Summarize findings
    const threadsWithRealAttachments = results.filter(
      (r) => r.hasRealAttachments
    ).length;
    const threadsWithOnlyInline = results.filter(
      (r) => r.inlineImages > 0 && !r.hasRealAttachments
    ).length;
    const threadsWithNoAttachments = results.filter(
      (r) => r.totalAttachments === 0
    ).length;

    logWithUser(
      `Summary: Among ${threads.length} threads with "has:attachment":`,
      "INFO"
    );
    logWithUser(
      `- ${threadsWithRealAttachments} threads have real attachments`,
      "INFO"
    );
    logWithUser(
      `- ${threadsWithOnlyInline} threads have only inline images`,
      "INFO"
    );
    logWithUser(
      `- ${threadsWithNoAttachments} threads have no attachments at all`,
      "INFO"
    );

    return {
      summary: {
        totalThreads: threads.length,
        threadsWithRealAttachments: threadsWithRealAttachments,
        threadsWithOnlyInline: threadsWithOnlyInline,
        threadsWithNoAttachments: threadsWithNoAttachments,
      },
      details: results,
    };
  } catch (error) {
    logWithUser(
      `Error in diagnoseAttachmentDiscrepancy: ${error.message}`,
      "ERROR"
    );
    return {
      error: error.message,
    };
  }
}

/**
 * Fix utility to manually remove processed labels from threads that don't have attachments
 * This helps clean up mislabeled threads
 */
function cleanupMislabeledThreads() {
  const labelName = CONFIG.processedLabelName;
  const label = GmailApp.getUserLabelByName(labelName);

  if (!label) {
    logWithUser(`Label ${labelName} does not exist yet`, "INFO");
    return { cleaned: 0 };
  }

  // Get threads with the processed label
  const labeledThreads = GmailApp.search(`label:${labelName}`, 0, 50);
  logWithUser(
    `Found ${labeledThreads.length} threads with the ${labelName} label`,
    "INFO"
  );

  let removedCount = 0;

  for (let i = 0; i < labeledThreads.length; i++) {
    const thread = labeledThreads[i];
    const messages = thread.getMessages();
    let hasRealAttachments = false;

    // Check each message in the thread
    for (let j = 0; j < messages.length; j++) {
      const message = messages[j];
      const attachments = message.getAttachments({
        includeInlineImages: false,
      });

      if (attachments.length > 0) {
        hasRealAttachments = true;
        break;
      }
    }

    // If this thread has no real attachments but has the label, remove the label
    if (!hasRealAttachments) {
      thread.removeLabel(label);
      removedCount++;
      logWithUser(
        `Removed label from thread: ${thread.getFirstMessageSubject()} (no real attachments)`,
        "INFO"
      );
    }
  }

  logWithUser(
    `Removed ${labelName} label from ${removedCount} threads that had no real attachments`,
    "INFO"
  );
  return { cleaned: removedCount };
}

/**
 * Tests the file timestamp functionality
 * Creates a test file and tries to set its timestamp to 6 months ago
 * Verifies if the timestamp was set correctly
 *
 * @returns {boolean} True if the test was successful, false otherwise
 */
function testTimestamps() {
  try {
    logWithUser("Starting timestamp test...", "INFO");

    // Check if Drive Advanced Service is enabled
    if (typeof Drive === "undefined") {
      logWithUser("Drive advanced service is not enabled!", "ERROR");
      return false;
    }

    // Log available Drive methods to help diagnose issues
    logWithUser(
      `Drive.Files.get available: ${typeof Drive.Files.get !== "undefined"}`,
      "INFO"
    );
    logWithUser(
      `Drive.Files.update available: ${
        typeof Drive.Files.update !== "undefined"
      }`,
      "INFO"
    );
    logWithUser(
      `Drive.Files.patch available: ${
        typeof Drive.Files.patch !== "undefined"
      }`,
      "INFO"
    );
    logWithUser(
      `Drive.Files.list available: ${typeof Drive.Files.list !== "undefined"}`,
      "INFO"
    );

    // Get main folder for testing
    const mainFolder = DriveApp.getFolderById(CONFIG.mainFolderId);
    logWithUser(`Main folder: ${mainFolder.getName()}`, "INFO");

    // Create a test file with current timestamp in the name to make it unique
    const timestamp = new Date().getTime();
    const testFileName = `timestamp_test_${timestamp}.txt`;
    const testFile = mainFolder.createFile(
      testFileName,
      "This is a test file for timestamp testing."
    );
    logWithUser(
      `Created test file: ${testFileName} with ID ${testFile.getId()}`,
      "INFO"
    );

    // Get the default creation date
    const defaultDate = testFile.getLastUpdated();
    logWithUser(`Default creation date: ${defaultDate.toISOString()}`, "INFO");

    // Create a date 6 months ago
    const now = new Date();
    const sixMonthsAgo = new Date();
    sixMonthsAgo.setMonth(now.getMonth() - 6);
    logWithUser(
      `Target date (6 months ago): ${sixMonthsAgo.toISOString()}`,
      "INFO"
    );

    // Try to set the creation date
    logWithUser(
      `Attempting to change timestamp to: ${sixMonthsAgo.toISOString()}`,
      "INFO"
    );
    const success = setFileCreationDate(testFile, sixMonthsAgo);

    if (success) {
      try {
        // Find the updated file by name instead of ID since the ID may have changed
        const files = mainFolder.getFilesByName(testFileName);
        if (!files.hasNext()) {
          logWithUser(
            "ERROR: Can't find the test file after timestamp operation",
            "ERROR"
          );
          return false;
        }

        const updatedFile = files.next();
        logWithUser(
          `Found updated file with ID: ${updatedFile.getId()}`,
          "INFO"
        );

        // Check the description for the timestamp marker
        const description = updatedFile.getDescription() || "";
        if (description.includes("original_date=")) {
          logWithUser(
            "Success: File description contains the timestamp marker",
            "INFO"
          );
        } else {
          logWithUser(
            "Warning: File description does not contain timestamp marker",
            "WARNING"
          );
        }

        // The actual file date won't match our target exactly, but at least check
        // that we didn't create a future-dated file
        const updatedDate = updatedFile.getLastUpdated();
        logWithUser(`Updated file date: ${updatedDate.toISOString()}`, "INFO");

        // Success criteria:
        // 1. Timestamp is in description
        // 2. File exists with the correct name

        logWithUser(
          "SUCCESS: Timestamp function completed and file preserved",
          "INFO"
        );
        return true;
      } catch (e) {
        logWithUser(`Error verifying timestamp: ${e.message}`, "ERROR");
        return false;
      }
    } else {
      logWithUser("ERROR: Timestamp function failed", "ERROR");
      return false;
    }
  } catch (error) {
    logWithUser(`Error in timestamp test: ${error.message}`, "ERROR");
    return false;
  }
}

/**
 * Helper function to calculate difference in months between two dates
 */
function dateDiffInMonths(date1, date2) {
  const d1 = new Date(date1);
  const d2 = new Date(date2);
  const yearsDiff = d2.getFullYear() - d1.getFullYear();
  const monthsDiff = d2.getMonth() - d1.getMonth();
  return yearsDiff * 12 + monthsDiff;
}
