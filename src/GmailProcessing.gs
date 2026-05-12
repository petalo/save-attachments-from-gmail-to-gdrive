/**
 * Gmail processing functions for Gmail Attachment Organizer
 *
 * Label helpers → LabelManagement.gs
 * Thread state  → ThreadState.gs
 */

/**
 * Builds a deterministic source ID for an attachment processing attempt.
 *
 * @param {string} threadId - Gmail thread ID
 * @param {string} messageId - Gmail message ID
 * @param {number} attachmentIndex - Attachment index within filtered message list
 * @param {GmailAttachment} attachment - Attachment object
 * @returns {string} Deterministic source ID
 */
function buildSourceAttachmentId(
  threadId,
  messageId,
  attachmentIndex,
  attachment
) {
  const name = attachment.getName();
  const size = attachment.getSize();
  return `${threadId}:${messageId}:${attachmentIndex}:${name}:${size}`;
}

/**
 * Processes emails for a specific user with optimized batch handling
 *
 * This function processes a single paginated page of unprocessed threads
 * per execution to keep runtime predictable and reduce resource usage.
 *
 * @param {string} userEmail - Email of the user to process
 * @param {boolean} oldestFirst - Whether to process oldest emails first (default: true)
 * @param {number|null} deadlineMs - Unix timestamp (ms) to stop safely before hard timeout
 * @returns {boolean} True if processing was successful, false if an error occurred
 *
 * The function follows this flow:
 * 1. Accesses the main Google Drive folder specified in CONFIG
 * 2. Gets or creates the Gmail label used to mark processed threads
 * 3. Builds search criteria to find unprocessed threads with attachments
 * 4. Retrieves one paginated page of matching threads using CONFIG.batchSize
 * 5. Optionally sorts threads in that page by date (oldest first)
 * 6. Processes that single page and applies processed labels as needed
 * 7. Returns success/failure status and logs processing statistics
 *
 * This paginated approach helps avoid hitting the 6-minute execution limit
 * by processing a controlled number of threads per execution.
 */
function processUserEmails(userEmail, oldestFirst = true, deadlineMs = null) {
  try {
    logWithUser(`Processing emails for user: ${userEmail}`, "INFO");

    // Get the main folder
    let mainFolder;
    try {
      mainFolder = DriveApp.getFolderById(CONFIG.mainFolderId);
      logWithUser(`Successfully accessed main folder: ${mainFolder.getName()}`);
    } catch (e) {
      throw new Error(
        `Invalid or inaccessible folder ID: ${CONFIG.mainFolderId}. Error: ${e.message}`
      );
    }

    // Get or create processing state labels
    let processingLabel = getProcessingLabel();
    let errorLabel = getErrorLabel();
    let permanentErrorLabel = getPermanentErrorLabel();
    let tooLargeLabel = getTooLargeLabel();

    // Cursor-based search replaces label-based filtering.
    //
    // Two modes depending on how far behind the cursor is:
    //
    //   Catch-up  (cursor older than bufferDays from now):
    //     Bounded window [cursor, cursor + windowDays] so Gmail returns a
    //     complete slice. Cursor advances by windowDays only when the window
    //     yields fewer threads than batchSize (window is exhausted).
    //
    //   Incremental (cursor within bufferDays of now):
    //     Open-ended window [now - bufferDays, ∞). Re-scans the recent window
    //     on every run to catch new replies in already-processed threads.
    //     Cursor advances to now after each successful batch.
    //
    // Gmail `after:` and `before:` accept Unix epoch seconds (no PST-midnight
    // ambiguity that YYYY/MM/DD causes for non-PST users).
    //
    // source_attachment_id stored in Drive file descriptions provides fine-grained
    // dedup: already-saved attachments are detected and skipped in O(1).
    const now = Math.floor(Date.now() / 1000);
    const { epoch: cursorEpoch, offset: cursorOffset } = getUserCursorState(userEmail);
    const windowEndEpoch = cursorEpoch + CONFIG.cursorWindowDays * 86400;
    const isIncremental = now - cursorEpoch <= CONFIG.cursorWindowBufferDays * 86400;
    const daysBehind = Math.round((now - cursorEpoch) / 86400);
    logWithUser(
      `Cursor: ${new Date(cursorEpoch * 1000).toISOString()} | offset=${cursorOffset} | ${daysBehind} days behind | Mode: ${isIncremental ? "incremental" : "catch-up"}`,
      "INFO"
    );

    let searchCriteria;
    if (isIncremental) {
      const bufferEpoch = now - CONFIG.cursorWindowBufferDays * 86400;
      searchCriteria =
        `has:attachment after:${bufferEpoch}` +
        ` -label:${CONFIG.permanentErrorLabelName}` +
        ` -label:${CONFIG.tooLargeLabelName}`;
      logWithUser(
        `Incremental mode: scanning threads after ${new Date(bufferEpoch * 1000).toISOString()}`,
        "INFO"
      );
    } else {
      searchCriteria =
        `has:attachment after:${cursorEpoch} before:${windowEndEpoch}` +
        ` -label:${CONFIG.permanentErrorLabelName}` +
        ` -label:${CONFIG.tooLargeLabelName}`;
      logWithUser(
        `Catch-up mode: window ${new Date(cursorEpoch * 1000).toISOString()} → ${new Date(windowEndEpoch * 1000).toISOString()}`,
        "INFO"
      );
    }

    const pageSize = Math.max(1, CONFIG.batchSize);
    const searchOffset = cursorOffset;
    const threads = GmailApp.search(searchCriteria, searchOffset, pageSize);
    logWithUser(
      `Retrieved ${threads.length} threads (limit=${pageSize}, incremental=${isIncremental})`,
      "INFO"
    );

    if (threads.length === 0) {
      if (!isIncremental) {
        // Empty catch-up window (or offset past all results) → advance to next window
        setUserCursorState(userEmail, windowEndEpoch, 0);
        logWithUser(
          `Empty catch-up window, cursor advanced to ${new Date(windowEndEpoch * 1000).toISOString()}`,
          "INFO"
        );
      }
      logWithUser("No threads with attachments found in current window", "INFO");
      return true;
    }

    // Preserve oldest-first preference within the current page.
    if (oldestFirst && threads.length > 1) {
      try {
        threads.sort(function (a, b) {
          const dateA = a.getLastMessageDate();
          const dateB = b.getLastMessageDate();
          return dateA - dateB;
        });
        logWithUser("Sorted paginated threads by date (oldest first)", "INFO");
      } catch (e) {
        logWithUser(
          `Error sorting paginated threads: ${e.message}. Using default order.`,
          "WARNING"
        );
      }
    }

    // Debug: Log subject of first thread for troubleshooting
    logWithUser(
      `First paginated thread subject: "${threads[0].getFirstMessageSubject()}" from ${threads[0]
        .getLastMessageDate()
        .toISOString()}`,
      "INFO"
    );

    // Process exactly one page per execution
    const result = processThreadsWithCounting(
      threads,
      mainFolder,
      processingLabel,
      errorLabel,
      permanentErrorLabel,
      tooLargeLabel,
      deadlineMs,
      userEmail
    );
    const threadsProcessed = result.processedThreads;
    const threadsWithValidAttachments = result.threadsWithAttachments;

    logWithUser(
      `Completed paginated processing: ${threadsProcessed} threads, ${threadsWithValidAttachments} with valid attachments for user: ${userEmail}`,
      "INFO"
    );
    if (result.stoppedByDeadline) {
      logWithUser(
        `Stopped early due to execution soft limit; remaining threads will continue in next run`,
        "WARNING"
      );
    } else {
      // Advance cursor only on full batch completion (not on timeout).
      if (isIncremental && threads.length < pageSize) {
        // Incremental buffer exhausted → advance cursor to now, reset offset.
        setUserCursorState(userEmail, now, 0);
      } else if (!isIncremental && threads.length < pageSize) {
        // Catch-up window exhausted → move to next window, reset offset.
        setUserCursorState(userEmail, windowEndEpoch, 0);
      } else {
        // Window has more threads than batchSize → advance offset to fetch next page.
        setUserCursorState(userEmail, cursorEpoch, cursorOffset + pageSize);
      }
    }
    return true;
  } catch (error) {
    logWithUser(
      `Failed to process emails for user ${userEmail}: ${error.message}`,
      "ERROR"
    );
    if (error.stack) {
      logWithUser(`Stack trace: ${error.stack}`, "ERROR");
    }
    return false;
  }
}

/**
 * Process Gmail threads and count those that have valid attachments
 * This is a modified version of processThreads that counts threads with valid attachments
 *
 * @param {GmailThread[]} threads - Gmail threads to process
 * @param {DriveFolder} mainFolder - The main folder to save attachments to
 * @param {GmailLabel} processingLabel - The label to apply while processing
 * @param {GmailLabel} errorLabel - The label to apply on processing failure
 * @param {GmailLabel} permanentErrorLabel - The label for non-retriable failures
 * @param {GmailLabel} tooLargeLabel - The label for too-large attachments
 * @param {number|null} deadlineMs - Unix timestamp (ms) to stop safely before hard timeout
 * @returns {Object} Object containing count of threads with valid attachments and processing status
 *
 * The function follows this flow:
 * 1. Iterates through each thread in the provided array
 * 2. Checks if the thread is already processed (has the processed label)
 * 3. For unprocessed threads:
 *    - Examines each message in the thread for attachments
 *    - Filters attachments based on sender domain and attachment properties
 *    - Creates domain folders as needed and saves valid attachments
 *    - Verifies file timestamps match email dates
 * 4. Marks threads as processed only when safe (or when all attachments were filtered out)
 * 5. Counts and returns processing statistics for the page
 *
 * This function is optimized for batch processing to handle the 6-minute execution limit
 * of Google Apps Script by focusing on counting threads with valid attachments.
 */
function processThreadsWithCounting(
  threads,
  mainFolder,
  processingLabel,
  errorLabel,
  permanentErrorLabel,
  tooLargeLabel,
  deadlineMs = null,
  userEmail = null
) {
  let threadsWithAttachments = 0;
  let processedThreads = 0;
  let stoppedByDeadline = false;
  const resolvedUserEmail =
    userEmail || Session.getEffectiveUser().getEmail();

  for (let i = 0; i < threads.length; i++) {
    if (deadlineMs && new Date().getTime() >= deadlineMs) {
      stoppedByDeadline = true;
      logWithUser(
        "Execution soft limit reached; stopping thread loop for safe resume",
        "WARNING"
      );
      break;
    }

    const thread = threads[i];
    let threadId = null;
    let processingLabelApplied = false;
    try {
      threadId = thread.getId();
        const previousFailure = getThreadFailureState(threadId);
        if (previousFailure && previousFailure.category === "permanent") {
          withRetry(
            () => thread.addLabel(permanentErrorLabel),
            "adding permanent error label from stored state"
          );
          logWithUser(
            `Skipping thread with permanent failure state: ${threadId} (${previousFailure.code})`,
            "WARNING"
          );
          processedThreads++;
          continue;
        }

        const messages = thread.getMessages();
        let threadProcessed = false;
        let threadHasAttachments = false;
        let threadHadValidAttachments = false;
        let threadHadSaveFailures = false;
        let threadHasTooLargeAttachments = false;
        let threadMarkedPermanentFailure = false;
        const threadSubject = thread.getFirstMessageSubject();
        logWithUser(`Processing thread: ${threadSubject}`);
        withRetry(
          () => thread.addLabel(processingLabel),
          "adding processing label"
        );
        processingLabelApplied = true;
        try {
          markThreadProcessingState(
            threadId,
            Session.getEffectiveUser().getEmail()
          );
        } catch (processingStateError) {
          logWithUser(
            `Failed to store processing checkpoint: ${processingStateError.message}`,
            "WARNING"
          );
        }

        // Process each message in the thread
        for (let j = 0; j < messages.length; j++) {
          const message = messages[j];
          const attachments = message.getAttachments();
          const messageId = message.getId();
          const sender = message.getFrom();
          // Get the message date to use for file creation date
          const messageDate = message.getDate();

          // Check if sender's domain should be skipped
          const senderDomain = sender.match(/@([\w.-]+)/)?.[1];
          if (senderDomain && CONFIG.skipDomains.includes(senderDomain)) {
            logWithUser(`Skipping domain: ${senderDomain}`, "INFO");
            continue;
          }

          if (attachments.length > 0) {
            // Mark that this thread has attachments, even if they're filtered out
            threadHasAttachments = true;

            logWithUser(
              `Found ${attachments.length} attachments in message from ${sender}${extractDomain(sender) === extractDomain(resolvedUserEmail) ? " (sent by me)" : ""}`,
              "DEBUG"
            );

            // Log MIME types to help diagnose what types of attachments are found
            attachments.forEach((att) => {
              try {
                logWithUser(
                  `Attachment: ${att.getName()}, Type: ${att.getContentType()}, Size: ${Math.round(
                    att.getSize() / 1024
                  )}KB`,
                  "DEBUG"
                );
              } catch (e) {
                // Skip if we can't get content type
              }
            });

            // Pre-filter attachments to see if any valid ones exist.
            // Each entry retains its rawIndex (position in message.getAttachments())
            // so sourceAttachmentId remains stable regardless of filter rule changes.
            const validAttachments = [];
            for (let rawIndex = 0; rawIndex < attachments.length; rawIndex++) {
              const attachment = attachments[rawIndex];
              const fileName = attachment.getName();
              const fileSize = attachment.getSize();

              if (fileSize > CONFIG.maxFileSize) {
                threadHasTooLargeAttachments = true;
                registerThreadFailure(threadId, {
                  category: "too_large",
                  code: "too_large",
                  context: "attachment_filter",
                  message: `Attachment exceeds max size (${fileSize} bytes > ${CONFIG.maxFileSize})`,
                  attachmentName: fileName,
                  attachmentSize: fileSize,
                });
                logWithUser(
                  `Attachment too large, skipping and labeling thread: ${fileName}`,
                  "WARNING"
                );
                continue;
              }

              // Skip if it matches filter criteria
              if (shouldSkipFile(fileName, fileSize, attachment)) {
                logWithUser(`Skipping attachment: ${fileName}`, "DEBUG");
                continue;
              }

              validAttachments.push({ attachment, rawIndex });
            }

            // Only create domain folder if we have valid attachments to save
            if (validAttachments.length > 0) {
              // For sent messages: route to each external recipient domain.
              // For received messages: route to the sender's domain (existing behavior).
              const isSentByMe =
                extractDomain(sender) === extractDomain(resolvedUserEmail);
              const targetDomains = isSentByMe
                ? extractExternalRecipientDomains(
                    message,
                    extractDomain(resolvedUserEmail)
                  // getDomainFolder calls extractDomain() internally, which requires an "@" to
                  // parse a domain. Prefix bare domain strings (e.g. "acme.com" → "@acme.com")
                  // since extractExternalRecipientDomains returns plain domain strings.
                  ).map((d) => "@" + d)
                : [sender];

              if (targetDomains.length === 0) {
                logWithUser(
                  `Skipping sent message with no external recipients: ${threadSubject}`,
                  "INFO"
                );
                continue;
              }

              threadHadValidAttachments = true;

              for (const domainTarget of targetDomains) {
                const domainFolder = getDomainFolder(domainTarget, mainFolder);

                // Process each valid attachment
                for (let k = 0; k < validAttachments.length; k++) {
                  const { attachment, rawIndex } = validAttachments[k];
                  const sourceAttachmentId = buildSourceAttachmentId(
                    threadId,
                    messageId,
                    rawIndex,
                    attachment
                  );
                  logWithUser(
                    `Processing attachment: ${attachment.getName()} (${Math.round(
                      attachment.getSize() / 1024
                    )}KB)`,
                    "DEBUG"
                  );
                  // Log the message date that will be used for the file timestamp
                  logWithUser(
                    `Email date for timestamp: ${messageDate.toISOString()}`,
                    "DEBUG"
                  );

                  // The same sourceAttachmentId is used across all domain targets for this attachment.
                  // Deduplication safety is maintained because saveAttachment checks by filename+size
                  // and then by source_attachment_id in the file description — the folderId scope
                  // differs per domain folder, so each domain copy is checked independently.
                  const saveResult = saveAttachment(attachment, message, domainFolder, {
                    sourceAttachmentId,
                  });

                  // Process the result object
                  if (saveResult.success) {
                    const savedFile = saveResult.file;
                    threadProcessed = true;
                    // If the file was saved, verify its timestamp was set correctly
                    try {
                      // Get the file's timestamp using DriveApp
                      const updatedDate = savedFile.getLastUpdated();
                      logWithUser(
                        `Saved file modification date: ${updatedDate.toISOString()}`,
                        "DEBUG"
                      );

                      // Compare with the message date
                      const messageDateTime = messageDate.getTime();
                      const fileDateTime = updatedDate.getTime();
                      const diffInMinutes =
                        Math.abs(messageDateTime - fileDateTime) / (1000 * 60);

                      if (diffInMinutes < 5) {
                        logWithUser(
                          "✅ File timestamp matches email date (within 5 minutes)",
                          "INFO"
                        );
                      } else {
                        logWithUser(
                          `⚠️ File timestamp differs from email date by ${Math.round(
                            diffInMinutes
                          )} minutes`,
                          "WARNING"
                        );
                      }
                    } catch (e) {
                      // Just log the error but don't stop processing
                      logWithUser(
                        `Error verifying file timestamp: ${e.message}`,
                        "WARNING"
                      );
                    }
                  } else {
                    threadHadSaveFailures = true;
                    const failureState = registerThreadFailure(threadId, {
                      context: "save_attachment",
                      message:
                        saveResult.error ||
                        `Failed to save attachment ${attachment.getName()}`,
                      attachmentName: attachment.getName(),
                      attachmentSize: attachment.getSize(),
                    });
                    if (failureState && failureState.category === "permanent") {
                      threadMarkedPermanentFailure = true;
                    }
                    logWithUser(
                      `Failed to save valid attachment: ${attachment.getName()}`,
                      "WARNING"
                    );
                  }
                } // end for (let k ...)
              } // end for (const domainTarget of targetDomains)
            } else {
              logWithUser(
                `No valid attachments to save in message from ${sender}`,
                "INFO"
              );
            }
          }
        }

        // Mark as processed only for clean outcomes:
        // 1. Valid attachments saved without save failures or too-large items
        // 2. Thread had attachments but none were valid, and no too-large items
        if (threadProcessed && !threadHadSaveFailures && !threadHasTooLargeAttachments) {
          withRetry(
            () => thread.removeLabel(errorLabel),
            "removing error label after successful processing"
          );
          withRetry(
            () => thread.removeLabel(permanentErrorLabel),
            "removing permanent error label after successful processing"
          );
          withRetry(
            () => thread.removeLabel(tooLargeLabel),
            "removing too-large label after successful processing"
          );
          clearThreadFailureState(threadId);
          logWithUser(
            `Thread "${threadSubject}" processed with valid attachments and labeled`
          );
          threadsWithAttachments++;
        } else if (
          threadHasAttachments &&
          !threadHadValidAttachments &&
          !threadHasTooLargeAttachments
        ) {
          withRetry(
            () => thread.removeLabel(errorLabel),
            "removing error label after filtered processing"
          );
          withRetry(
            () => thread.removeLabel(permanentErrorLabel),
            "removing permanent error label after filtered processing"
          );
          withRetry(
            () => thread.removeLabel(tooLargeLabel),
            "removing too-large label after filtered processing"
          );
          clearThreadFailureState(threadId);
          logWithUser(
            `Thread "${threadSubject}" had attachments but none were valid; skipped`,
            "INFO"
          );
        } else if (threadHadValidAttachments && threadHadSaveFailures) {
          withRetry(() => thread.addLabel(errorLabel), "adding error label");
          if (threadMarkedPermanentFailure) {
            withRetry(
              () => thread.addLabel(permanentErrorLabel),
              "adding permanent error label"
            );
          }
          if (threadHasTooLargeAttachments) {
            withRetry(
              () => thread.addLabel(tooLargeLabel),
              "adding too-large label alongside save failures"
            );
          }
          logWithUser(
            `Thread "${threadSubject}" had valid attachments with save failures; not marking as processed for retry`,
            "WARNING"
          );
        } else if (threadHasTooLargeAttachments) {
          withRetry(() => thread.addLabel(tooLargeLabel), "adding too-large label");
          withRetry(
            () => thread.removeLabel(errorLabel),
            "removing error label for too-large-only thread"
          );
          withRetry(
            () => thread.removeLabel(permanentErrorLabel),
            "removing permanent error label for too-large-only thread"
          );
          logWithUser(
            `Thread "${threadSubject}" contains attachments above max size; marked as TooLarge`,
            "WARNING"
          );
        } else {
          logWithUser(
            `No attachments found in thread "${threadSubject}"`,
            "INFO"
          );
        }
    } catch (error) {
      const failedThreadId = threadId || thread.getId();
      const failureState = registerThreadFailure(failedThreadId, {
        context: "thread_exception",
        message: error.message,
      });
      logWithUser(
        `Error processing thread "${thread.getFirstMessageSubject()}": ${
          error.message
        }`,
        "ERROR"
      );
      try {
        withRetry(
          () => thread.addLabel(errorLabel),
          "adding error label after thread exception"
        );
        if (failureState && failureState.category === "permanent") {
          withRetry(
            () => thread.addLabel(permanentErrorLabel),
            "adding permanent error label after thread exception"
          );
        }
      } catch (errorLabelError) {
        logWithUser(
          `Failed to add error label: ${errorLabelError.message}`,
          "WARNING"
        );
      }
      // Continue with next thread instead of stopping execution
    } finally {
      if (processingLabelApplied) {
        try {
          withRetry(
            () => thread.removeLabel(processingLabel),
            "removing processing label"
          );
        } catch (processingCleanupError) {
          logWithUser(
            `Failed to remove processing label: ${processingCleanupError.message}`,
            "WARNING"
          );
        }
      }
      if (threadId) {
        try {
          clearThreadProcessingState(threadId);
        } catch (processingStateCleanupError) {
          logWithUser(
            `Failed to clear processing state: ${processingStateCleanupError.message}`,
            "WARNING"
          );
        }
      }
    }

    processedThreads++;
  }

  return { threadsWithAttachments, processedThreads, stoppedByDeadline };
}

