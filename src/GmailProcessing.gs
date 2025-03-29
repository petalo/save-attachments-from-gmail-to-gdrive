/**
 * Gmail processing functions for Gmail Attachment Organizer
 */

/**
 * Determines if a message appears to contain an invoice using AI or keywords
 *
 * @param {GmailMessage} message - The Gmail message to analyze
 * @returns {boolean} True if the message likely contains an invoice
 */
function isInvoiceMessage(message) {
  // Early return if invoice detection is disabled
  if (CONFIG.invoiceDetection === false) return false;

  try {
    // Check if any detection method is enabled
    if (
      CONFIG.invoiceDetection === "gemini" ||
      CONFIG.invoiceDetection === "openai" ||
      CONFIG.invoiceDetection === "email"
    ) {
      // Check if message has attachments and if any are PDFs (with stricter checking)
      if (CONFIG.onlyAnalyzePDFs) {
        const attachments = message.getAttachments();
        let hasPDF = false;

        for (const attachment of attachments) {
          const fileName = attachment.getName().toLowerCase();
          const contentType = attachment.getContentType().toLowerCase();

          // Strict PDF check - both extension and MIME type
          if (CONFIG.strictPdfCheck) {
            if (fileName.endsWith(".pdf") && contentType.includes("pdf")) {
              hasPDF = true;
              break;
            }
          } else {
            // Legacy check - just extension
            if (fileName.endsWith(".pdf")) {
              hasPDF = true;
              break;
            }
          }
        }

        // Skip AI if there are no PDF attachments
        if (!hasPDF) {
          logWithUser("No PDF attachments found, skipping AI analysis", "INFO");
          // Don't fall back to keywords when there are no PDFs if onlyAnalyzePDFs is true
          // This prevents non-PDF attachments from being classified as invoices based on keywords
          return false;
        }
      }

      // Check if sender domain should be skipped
      const sender = message.getFrom();
      const domain = extractDomain(sender);
      if (CONFIG.skipAIForDomains && CONFIG.skipAIForDomains.includes(domain)) {
        logWithUser(`Skipping AI for domain ${domain}, using keywords`, "INFO");
        return checkKeywords(message);
      }

      // NEW: Check for email-based detection
      if (CONFIG.invoiceDetection === "email") {
        const isInvoice = InvoiceDetection.isInvoiceSender(sender);
        logWithUser(
          `Email-based invoice detection result: ${isInvoice}`,
          "INFO"
        );
        return isInvoice;
      }

      // Try Gemini detection first if enabled
      if (CONFIG.invoiceDetection === "gemini") {
        try {
          const isInvoice = GeminiDetection.isInvoiceWithGemini(message);
          logWithUser(`Gemini invoice detection result: ${isInvoice}`, "INFO");
          return isInvoice;
        } catch (geminiError) {
          logWithUser(
            `Gemini detection error: ${geminiError.message}, trying fallback options`,
            "WARNING"
          );

          // Try OpenAI if enabled as fallback
          if (CONFIG.invoiceDetection === "openai") {
            try {
              const isInvoice = OpenAIDetection.isInvoiceWithOpenAI(message);
              logWithUser(
                `OpenAI invoice detection result: ${isInvoice}`,
                "INFO"
              );
              return isInvoice;
            } catch (openaiError) {
              logWithUser(
                `OpenAI detection error: ${openaiError.message}, falling back to keywords`,
                "WARNING"
              );
              return CONFIG.fallbackToKeywords ? checkKeywords(message) : false;
            }
          } else if (CONFIG.fallbackToKeywords) {
            return checkKeywords(message);
          } else {
            return false;
          }
        }
      }
      // Try OpenAI if Gemini is disabled but OpenAI is enabled
      else if (CONFIG.invoiceDetection === "openai") {
        try {
          const isInvoice = OpenAIDetection.isInvoiceWithOpenAI(message);
          logWithUser(`OpenAI invoice detection result: ${isInvoice}`, "INFO");
          return isInvoice;
        } catch (openaiError) {
          logWithUser(
            `OpenAI detection error: ${openaiError.message}, falling back to keywords`,
            "WARNING"
          );
          return CONFIG.fallbackToKeywords ? checkKeywords(message) : false;
        }
      }
    }

    // Use keyword detection if no AI is enabled
    return checkKeywords(message);
  } catch (error) {
    logWithUser(`Error in invoice detection: ${error.message}`, "ERROR");
    return false;
  }
}

/**
 * Helper function to check for invoice keywords in message subject and body
 *
 * @param {GmailMessage} message - The Gmail message to check
 * @returns {boolean} True if invoice keywords are found
 */
function checkKeywords(message) {
  try {
    // Check subject for invoice keywords
    const subject = message.getSubject().toLowerCase();
    for (const keyword of CONFIG.invoiceKeywords) {
      if (subject.includes(keyword.toLowerCase())) {
        logWithUser(
          `Invoice keyword "${keyword}" found in subject: "${subject}"`,
          "INFO"
        );
        return true;
      }
    }

    // Check body for invoice keywords (optional, may affect performance)
    try {
      const body = message.getPlainBody().toLowerCase();
      for (const keyword of CONFIG.invoiceKeywords) {
        if (body.includes(keyword.toLowerCase())) {
          logWithUser(
            `Invoice keyword "${keyword}" found in message body`,
            "INFO"
          );
          return true;
        }
      }
    } catch (e) {
      // If we can't get the body, just log and continue
      logWithUser(`Could not check message body: ${e.message}`, "WARNING");
    }

    return false;
  } catch (error) {
    logWithUser(
      `Error checking for invoice keywords: ${error.message}`,
      "ERROR"
    );
    return false;
  }
}

/**
 * Determines if an attachment appears to be an invoice based on file extension and content type
 *
 * @param {GmailAttachment} attachment - The attachment to analyze
 * @returns {boolean} True if the attachment likely is an invoice
 */
function isInvoiceAttachment(attachment) {
  if (CONFIG.invoiceDetection === false) return false;

  try {
    const fileName = attachment.getName().toLowerCase();
    const contentType = attachment.getContentType().toLowerCase();

    // Check file extension against invoice file types with stricter checking
    for (const ext of CONFIG.invoiceFileTypes) {
      const lowerExt = ext.toLowerCase();

      // Strict PDF check - both extension and MIME type
      if (CONFIG.strictPdfCheck) {
        if (
          lowerExt === ".pdf" &&
          fileName.endsWith(lowerExt) &&
          contentType.includes("pdf")
        ) {
          return true;
        }
      } else {
        // Legacy check - just extension
        if (fileName.endsWith(lowerExt)) {
          return true;
        }
      }
    }

    return false;
  } catch (error) {
    logWithUser(
      `Error checking if attachment is invoice: ${error.message}`,
      "ERROR"
    );
    return false;
  }
}

/**
 * Gets or creates the processed label
 *
 * @returns {GmailLabel} The Gmail label used to mark processed threads
 */
function getProcessedLabel() {
  let processedLabel = GmailApp.getUserLabelByName(CONFIG.processedLabelName);

  if (!processedLabel) {
    processedLabel = GmailApp.createLabel(CONFIG.processedLabelName);
    logWithUser(`Created new Gmail label: ${CONFIG.processedLabelName}`);
  } else {
    logWithUser(`Using existing Gmail label: ${CONFIG.processedLabelName}`);
  }

  return processedLabel;
}

/**
 * Processes emails for a specific user with optimized batch handling
 *
 * This function processes emails in batches until it finds enough threads
 * with valid attachments or exhausts all unprocessed threads.
 *
 * @param {string} userEmail - Email of the user to process
 * @param {boolean} oldestFirst - Whether to process oldest emails first (default: true)
 * @returns {boolean} True if processing was successful, false if an error occurred
 *
 * The function follows this flow:
 * 1. Accesses the main Google Drive folder specified in CONFIG
 * 2. Gets or creates the Gmail label used to mark processed threads
 * 3. Builds search criteria to find unprocessed threads with attachments
 * 4. Tests different search variations to find the most effective query
 * 5. Retrieves and optionally sorts threads by date (oldest first if specified)
 * 6. Processes threads in batches, with batch size determined by CONFIG.batchSize
 * 7. Tracks progress and adjusts batch size dynamically to meet target counts
 * 8. Returns success/failure status and logs detailed processing statistics
 *
 * This batch processing approach helps avoid hitting the 6-minute execution limit
 * of Google Apps Script by processing a controlled number of threads per execution.
 */
function processUserEmails(userEmail, oldestFirst = true) {
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

    // Get or create the 'Processed' label
    let processedLabel = getProcessedLabel();

    // Build search criteria using the processedLabelName from config
    let searchCriteria = `has:attachment -label:${CONFIG.processedLabelName}`;

    // Debug: Try several search variations to see if any of them work better
    const searchVariations = [
      { name: "Standard", query: searchCriteria },
      { name: "In Inbox", query: `in:inbox ${searchCriteria}` },
      {
        name: "With quote",
        query: `has:attachment -label:"${CONFIG.processedLabelName}"`,
      },
      {
        name: "With parentheses",
        query: `has:attachment AND -(label:${CONFIG.processedLabelName})`,
      },
      {
        name: "Explicit attachment",
        query: `filename:* -label:${CONFIG.processedLabelName}`,
      },
    ];

    // Try each search variation and log results
    logWithUser("Testing different search variations:", "INFO");
    for (const variation of searchVariations) {
      const count = GmailApp.search(variation.query).length;
      logWithUser(
        `- ${variation.name}: "${variation.query}" found ${count} threads`,
        "INFO"
      );
    }

    // Debug: Check how many threads match the base search criteria
    const totalMatchingThreads = GmailApp.search(searchCriteria).length;
    logWithUser(
      `Total unprocessed threads with attachments found: ${totalMatchingThreads}`,
      "INFO"
    );

    // Continue only if we actually have matching threads
    if (totalMatchingThreads === 0) {
      logWithUser(
        "No unprocessed threads with attachments found, skipping processing",
        "INFO"
      );
      return true;
    }

    // NOTE: We no longer add 'older_first' to the search criteria since it doesn't work as expected
    // Instead, we'll sort the threads ourselves after retrieving them
    if (oldestFirst) {
      logWithUser(
        "Will manually sort threads by date (oldest first) after retrieving them",
        "INFO"
      );
    } else {
      logWithUser("Using default order (newest first)", "INFO");
    }

    // Initialize batch processing
    let threadsProcessed = 0;
    let threadsWithValidAttachments = 0;
    let batchSize = CONFIG.batchSize;
    let offset = 0;

    // Get all threads that match our search criteria
    let allThreads = GmailApp.search(searchCriteria);
    logWithUser(
      `Retrieved ${allThreads.length} total threads matching search criteria`,
      "INFO"
    );

    // If we want oldest first, manually sort the threads by date
    if (oldestFirst && allThreads.length > 0) {
      try {
        // Sort threads by date (oldest first)
        allThreads.sort(function (a, b) {
          const dateA = a.getLastMessageDate();
          const dateB = b.getLastMessageDate();
          return dateA - dateB; // Ascending order (oldest first)
        });
        logWithUser(
          "Successfully sorted threads by date (oldest first)",
          "INFO"
        );
      } catch (e) {
        logWithUser(
          `Error sorting threads: ${e.message}. Will use default order.`,
          "WARNING"
        );
      }
    }

    // Continue processing batches until we reach the desired number of threads with valid attachments
    // or until we run out of unprocessed threads
    while (
      threadsWithValidAttachments < CONFIG.batchSize &&
      offset < allThreads.length
    ) {
      // Get the next batch of threads from our already retrieved and sorted list
      const currentBatchSize = Math.min(batchSize, allThreads.length - offset);
      const threads = allThreads.slice(offset, offset + currentBatchSize);

      // If no more threads, break out of the loop
      if (threads.length === 0) {
        logWithUser(
          "No more unprocessed threads with attachments found in this batch",
          "INFO"
        );
        break;
      }

      logWithUser(
        `Processing batch of ${threads.length} threads (batch ${
          Math.floor(offset / batchSize) + 1
        })`,
        "INFO"
      );

      // Debug: Log subject of first thread for troubleshooting
      if (threads.length > 0) {
        logWithUser(
          `First thread subject: "${threads[0].getFirstMessageSubject()}" from ${threads[0]
            .getLastMessageDate()
            .toISOString()}`,
          "INFO"
        );
      }

      // Process each thread and count those with valid attachments
      const result = processThreadsWithCounting(
        threads,
        mainFolder,
        processedLabel
      );
      threadsProcessed += threads.length;
      threadsWithValidAttachments += result.threadsWithAttachments;

      logWithUser(
        `Batch processed ${threads.length} threads, ${result.threadsWithAttachments} had valid attachments`,
        "INFO"
      );

      // Update the offset for the next batch
      offset += threads.length;

      // If we're getting close to our target, reduce the next batch size to avoid processing too many
      if (threadsWithValidAttachments + batchSize > CONFIG.batchSize * 1.5) {
        batchSize = Math.max(5, CONFIG.batchSize - threadsWithValidAttachments);
      }
    }

    logWithUser(
      `Completed processing ${threadsProcessed} threads, ${threadsWithValidAttachments} with valid attachments for user: ${userEmail}`,
      "INFO"
    );
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
 * @param {GmailLabel} processedLabel - The label to apply to processed threads
 * @returns {Object} Object containing count of threads with valid attachments
 *
 * The function follows this flow:
 * 1. Iterates through each thread in the provided array
 * 2. Checks if the thread is already processed (has the processed label)
 * 3. For unprocessed threads:
 *    - Examines each message in the thread for attachments
 *    - Filters attachments based on sender domain and attachment properties
 *    - Creates domain folders as needed and saves valid attachments
 *    - Verifies file timestamps match email dates
 * 4. Marks threads as processed if they had attachments (even if all were filtered out)
 * 5. Counts and returns the number of threads that had valid attachments saved
 *
 * This function is optimized for batch processing to handle the 6-minute execution limit
 * of Google Apps Script by focusing on counting threads with valid attachments.
 */
function processThreadsWithCounting(threads, mainFolder, processedLabel) {
  let threadsWithAttachments = 0;

  for (let i = 0; i < threads.length; i++) {
    const thread = threads[i];
    try {
      // Check if the thread is already processed
      const threadLabels = thread.getLabels();
      const isAlreadyProcessed = threadLabels.some(
        (label) => label.getName() === processedLabel.getName()
      );

      if (!isAlreadyProcessed) {
        const messages = thread.getMessages();
        let threadProcessed = false;
        let threadHasAttachments = false;
        const threadSubject = thread.getFirstMessageSubject();
        logWithUser(`Processing thread: ${threadSubject}`);

        // Process each message in the thread
        for (let j = 0; j < messages.length; j++) {
          const message = messages[j];
          const attachments = message.getAttachments();
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
              `Found ${attachments.length} attachments in message from ${sender}`
            );

            // Log MIME types to help diagnose what types of attachments are found
            attachments.forEach((att) => {
              try {
                logWithUser(
                  `Attachment: ${att.getName()}, Type: ${att.getContentType()}, Size: ${Math.round(
                    att.getSize() / 1024
                  )}KB`,
                  "INFO"
                );
              } catch (e) {
                // Skip if we can't get content type
              }
            });

            // Pre-filter attachments to see if any valid ones exist
            const validAttachments = [];
            for (let k = 0; k < attachments.length; k++) {
              const attachment = attachments[k];
              const fileName = attachment.getName();
              const fileSize = attachment.getSize();

              // Skip if it matches our filter criteria - passing full attachment object
              if (
                shouldSkipFile(fileName, fileSize, attachment) ||
                fileSize > CONFIG.maxFileSize
              ) {
                logWithUser(`Skipping attachment: ${fileName}`, "INFO");
                continue;
              }

              validAttachments.push(attachment);
            }

            // Only create domain folder if we have valid attachments to save
            if (validAttachments.length > 0) {
              const domainFolder = getDomainFolder(sender, mainFolder);

              // Check if this message might contain invoices
              const isInvoice = isInvoiceMessage(message);
              let invoicesFolder = null;

              // If invoice detection is enabled and this might be an invoice,
              // get the invoices folder with domain subfolder
              if (CONFIG.invoiceDetection !== false && isInvoice) {
                logWithUser(`Message appears to contain invoice(s)`, "INFO");
                const senderDomain = extractDomain(sender);
                invoicesFolder = getInvoicesFolder(mainFolder, senderDomain);
              }

              // Process each valid attachment
              for (let k = 0; k < validAttachments.length; k++) {
                const attachment = validAttachments[k];
                logWithUser(
                  `Processing attachment: ${attachment.getName()} (${Math.round(
                    attachment.getSize() / 1024
                  )}KB)`
                );
                // Log the message date that will be used for the file timestamp
                logWithUser(
                  `Email date for timestamp: ${messageDate.toISOString()}`,
                  "INFO"
                );

                // Use the legacy wrapper for backward compatibility
                const savedFile = saveAttachmentLegacy(
                  attachment,
                  domainFolder,
                  messageDate
                );

                // Process the result object
                if (savedFile) {
                  // If the file was saved, verify its timestamp was set correctly
                  try {
                    // Get the file's timestamp using DriveApp
                    const updatedDate = savedFile.getLastUpdated();
                    logWithUser(
                      `Saved file modification date: ${updatedDate.toISOString()}`,
                      "INFO"
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

                    // If this is an invoice or the attachment looks like an invoice,
                    // also save it to the invoices folder
                    if (invoicesFolder || isInvoiceAttachment(attachment)) {
                      if (!invoicesFolder) {
                        logWithUser(
                          `Attachment appears to be an invoice based on file type`,
                          "INFO"
                        );
                        const senderDomain = extractDomain(sender);
                        invoicesFolder = getInvoicesFolder(
                          mainFolder,
                          senderDomain
                        );
                      }

                      // Save a copy to the invoices folder
                      try {
                        logWithUser(
                          `Saving copy to invoices folder: ${CONFIG.invoicesFolderName}`,
                          "INFO"
                        );
                        const invoiceFile = saveAttachmentLegacy(
                          attachment,
                          invoicesFolder,
                          messageDate
                        );

                        if (invoiceFile) {
                          logWithUser(
                            `Successfully saved invoice copy: ${attachment.getName()}`,
                            "INFO"
                          );
                        }
                      } catch (invoiceError) {
                        logWithUser(
                          `Error saving to invoices folder: ${invoiceError.message}`,
                          "ERROR"
                        );
                        // Continue processing even if saving to invoices folder fails
                      }
                    }
                  } catch (e) {
                    // Just log the error but don't stop processing
                    logWithUser(
                      `Error verifying file timestamp: ${e.message}`,
                      "WARNING"
                    );
                  }
                }

                threadProcessed = true;
              }
            } else {
              logWithUser(
                `No valid attachments to save in message from ${sender}`,
                "INFO"
              );
            }
          }
        }

        // Mark as processed:
        // 1. If we processed valid attachments OR
        // 2. If the thread had attachments but they were all filtered out
        if (threadProcessed || threadHasAttachments) {
          withRetry(
            () => thread.addLabel(processedLabel),
            "adding processed label"
          );

          if (threadProcessed) {
            logWithUser(
              `Thread "${threadSubject}" processed with valid attachments and labeled`
            );
            threadsWithAttachments++;
          } else {
            logWithUser(
              `Thread "${threadSubject}" had attachments that were filtered out, marked as processed`
            );
          }
        } else {
          logWithUser(
            `No attachments found in thread "${threadSubject}"`,
            "INFO"
          );
        }
      }
    } catch (error) {
      logWithUser(
        `Error processing thread "${thread.getFirstMessageSubject()}": ${
          error.message
        }`,
        "ERROR"
      );
      // Continue with next thread instead of stopping execution
    }
  }

  return { threadsWithAttachments };
}

/**
 * Process the messages in a thread, saving attachments if any
 *
 * @param {GmailThread} thread - The Gmail thread to process
 * @param {GmailLabel} processedLabel - The label to apply to processed threads
 * @param {DriveFolder} mainFolder - The main Google Drive folder
 * @returns {Object} Processing results with counts of processed attachments
 *
 * The function follows this flow:
 * 1. Retrieves all messages in the thread
 * 2. For each message:
 *    - Gets all attachments and filters out those that should be skipped
 *    - Extracts the sender's domain to determine the target folder
 *    - Creates or uses an existing domain folder
 *    - Saves each valid attachment to the appropriate folder
 *    - Tracks statistics (saved, duplicates, errors, etc.)
 * 3. Applies the processed label to the thread regardless of outcome
 * 4. Returns a detailed result object with processing statistics
 *
 * This function is used by processThreadsWithCounting but provides more detailed
 * statistics about the processing results.
 */
function processMessages(thread, processedLabel, mainFolder) {
  try {
    const messages = thread.getMessages();
    let result = {
      totalAttachments: 0,
      savedAttachments: 0,
      savedSize: 0,
      skippedAttachments: 0,
      errors: 0,
      duplicates: 0,
    };

    for (const message of messages) {
      try {
        const attachments = message.getAttachments();
        const validAttachments = attachments.filter(
          (att) => !shouldSkipFile(att.getName(), att.getSize(), att)
        );

        if (validAttachments.length === 0) {
          continue; // Skip messages with no valid attachments
        }

        // Get sender details
        const sender = message.getFrom();
        const domain = extractDomain(sender);

        // Get or create domain folder
        const domainFolder = getDomainFolder(sender, mainFolder);

        if (!domainFolder) {
          logWithUser(
            `Error: Could not find or create folder for domain ${domain}`,
            "ERROR"
          );
          result.errors++;
          continue;
        }

        // Check if this message might contain invoices
        const isInvoice = isInvoiceMessage(message);
        let invoicesFolder = null;

        // If invoice detection is enabled and this might be an invoice,
        // get the invoices folder with domain subfolder
        if (CONFIG.invoiceDetection !== false && isInvoice) {
          logWithUser(`Message appears to contain invoice(s)`, "INFO");
          invoicesFolder = getInvoicesFolder(mainFolder, domain);
        }

        result.totalAttachments += validAttachments.length;

        // Process each valid attachment
        for (const attachment of validAttachments) {
          // Save to domain folder
          const saveResult = saveAttachment(attachment, message, domainFolder);

          if (saveResult.success) {
            if (saveResult.duplicate) {
              result.duplicates++;
            } else {
              result.savedAttachments++;
              result.savedSize += attachment.getSize();

              // If this is an invoice or the attachment looks like an invoice,
              // also save it to the invoices folder
              if (invoicesFolder || isInvoiceAttachment(attachment)) {
                if (!invoicesFolder) {
                  logWithUser(
                    `Attachment appears to be an invoice based on file type`,
                    "INFO"
                  );
                  invoicesFolder = getInvoicesFolder(mainFolder, domain);
                }

                // Save a copy to the invoices folder
                try {
                  logWithUser(
                    `Saving copy to invoices folder: ${CONFIG.invoicesFolderName}`,
                    "INFO"
                  );
                  const invoiceResult = saveAttachment(
                    attachment,
                    message,
                    invoicesFolder
                  );

                  if (invoiceResult.success && !invoiceResult.duplicate) {
                    logWithUser(
                      `Successfully saved invoice copy: ${attachment.getName()}`,
                      "INFO"
                    );
                  }
                } catch (invoiceError) {
                  logWithUser(
                    `Error saving to invoices folder: ${invoiceError.message}`,
                    "ERROR"
                  );
                  // Continue processing even if saving to invoices folder fails
                }
              }
            }
          } else {
            result.skippedAttachments++;
            result.errors++;
          }
        }
      } catch (messageError) {
        logWithUser(
          `Error processing message: ${messageError.message}`,
          "ERROR"
        );
        result.errors++;
      }
    }

    // Apply the processed label to the thread
    thread.addLabel(processedLabel);

    // Log a summary for the thread if it had attachments
    if (result.totalAttachments > 0) {
      logWithUser(
        `Thread processed: ${result.savedAttachments} saved, ${result.duplicates} duplicates, ${result.skippedAttachments} skipped`,
        "INFO"
      );
    }

    return result;
  } catch (error) {
    logWithUser(`Error in processMessages: ${error.message}`, "ERROR");
    throw error;
  }
}
