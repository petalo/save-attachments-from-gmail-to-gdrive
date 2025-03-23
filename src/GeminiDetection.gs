/**
 * Gemini Detection functions for Gmail Attachment Organizer
 *
 * This module provides functions to interact with the Google Gemini API
 * for improved invoice detection in emails with privacy-focused approach.
 */

/**
 * Makes an API call to Google Gemini
 *
 * @param {string} prompt - The prompt to send to Gemini
 * @param {Object} options - Additional options for the API call
 * @returns {Object} The API response
 */
function callGeminiAPI(prompt, options = {}) {
  try {
    // Get the API key
    const apiKey = getGeminiApiKey();
    if (!apiKey) {
      throw new Error("Gemini API key not found");
    }

    // Default options
    const defaultOptions = {
      model: CONFIG.geminiModel,
      temperature: CONFIG.geminiTemperature,
      maxOutputTokens: CONFIG.geminiMaxTokens,
    };

    // Merge default options with provided options
    const finalOptions = { ...defaultOptions, ...options };

    // Prepare the request payload
    const payload = {
      contents: [
        {
          parts: [{ text: prompt }],
        },
      ],
      generationConfig: {
        temperature: finalOptions.temperature,
        maxOutputTokens: finalOptions.maxOutputTokens,
      },
    };

    // Try v1 API first (newer version)
    let url =
      "https://generativelanguage.googleapis.com/v1/models/" +
      finalOptions.model +
      ":generateContent";

    try {
      // Make the API request to v1 endpoint
      const response = UrlFetchApp.fetch(`${url}?key=${apiKey}`, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      });

      // Parse the response
      const responseData = JSON.parse(response.getContentText());

      // Check for errors in the response
      if (response.getResponseCode() !== 200) {
        throw new Error(
          `API Error: ${responseData.error?.message || "Unknown error"}`
        );
      }

      // Extract the text response
      const responseText = responseData.candidates[0].content.parts[0].text;
      logWithUser("Successfully used Gemini API v1 endpoint", "INFO");
      return responseText;
    } catch (v1Error) {
      // Log the error but don't throw yet
      logWithUser(
        `Gemini API v1 endpoint failed: ${v1Error.message}, trying v1beta...`,
        "WARNING"
      );

      // Fall back to v1beta API
      url =
        "https://generativelanguage.googleapis.com/v1beta/models/" +
        finalOptions.model +
        ":generateContent";

      // Make the API request to v1beta endpoint
      const response = UrlFetchApp.fetch(`${url}?key=${apiKey}`, {
        method: "post",
        contentType: "application/json",
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      });

      // Parse the response
      const responseData = JSON.parse(response.getContentText());

      // Check for errors in the response
      if (response.getResponseCode() !== 200) {
        throw new Error(
          `API Error: ${responseData.error?.message || "Unknown error"}`
        );
      }

      // Extract the text response
      const responseText = responseData.candidates[0].content.parts[0].text;
      logWithUser("Successfully used Gemini API v1beta endpoint", "INFO");
      return responseText;
    }
  } catch (error) {
    logWithUser(`Gemini API call failed: ${error.message}`, "ERROR");
    throw error;
  }
}

/**
 * Extracts relevant keywords from email body without sending full content
 *
 * @param {string} body - The email body text
 * @returns {string[]} Array of extracted keywords and patterns
 */
function extractKeywords(body) {
  try {
    // Extract only relevant terms without sending the full body
    const relevantTerms = [];
    CONFIG.invoiceKeywords.forEach((keyword) => {
      if (body.toLowerCase().includes(keyword.toLowerCase())) {
        relevantTerms.push(keyword);
      }
    });

    // Look for numeric patterns that might be amounts or references
    const amountPatterns =
      body.match(/\$\s*\d+[.,]\d{2}|€\s*\d+[.,]\d{2}/g) || [];
    const refPatterns = body.match(/ref\w*\s*:?\s*[A-Z0-9-]+/gi) || [];
    const invoiceNumPatterns = body.match(/inv\w*\s*:?\s*[A-Z0-9-]+/gi) || [];

    return [
      ...relevantTerms,
      ...amountPatterns,
      ...refPatterns,
      ...invoiceNumPatterns,
    ];
  } catch (error) {
    logWithUser(`Error extracting keywords: ${error.message}`, "ERROR");
    return [];
  }
}

/**
 * Extracts metadata from a Gmail message for AI analysis
 * without sending the full content
 *
 * @param {GmailMessage} message - The Gmail message to extract metadata from
 * @returns {Object} Metadata object with privacy-safe information
 */
function extractMetadata(message) {
  try {
    const subject = message.getSubject() || "";
    const body = message.getPlainBody() || "";
    const attachments = message.getAttachments() || [];

    // Get attachment info
    const attachmentInfo = attachments.map((attachment) => {
      return {
        name: attachment.getName(),
        extension: attachment.getName().split(".").pop().toLowerCase(),
        contentType: attachment.getContentType(),
        size: attachment.getSize(),
      };
    });

    // Get historical patterns if enabled
    let historicalPatterns = null;
    if (CONFIG.useHistoricalPatterns && CONFIG.manuallyLabeledInvoicesLabel) {
      historicalPatterns = HistoricalPatterns.getHistoricalInvoicePatterns(
        message.getFrom()
      );
    }

    // Get sender information
    const sender = message.getFrom() || "";
    const senderDomain = extractDomain(sender);

    // Anonymize sender email (only keep domain)
    const anonymizedSender = `user@${senderDomain}`;

    // Create metadata object with only necessary information
    // Avoid sending any personally identifiable information
    return {
      subject: subject,
      senderDomain: senderDomain, // Only send domain, not full email
      date: message.getDate().toISOString(),
      hasAttachments: attachments.length > 0,
      attachmentTypes: attachmentInfo.map((a) => a.extension),
      attachmentContentTypes: attachmentInfo.map((a) => a.contentType),
      keywordsFound: extractKeywords(body),
      historicalPatterns: historicalPatterns,
    };
  } catch (error) {
    logWithUser(`Error extracting message metadata: ${error.message}`, "ERROR");
    // Return minimal metadata on error, still ensuring privacy
    const errorSenderDomain = extractDomain(message.getFrom() || "");
    return {
      subject: message.getSubject() || "",
      senderDomain: errorSenderDomain, // Only include domain, not full email
      hasAttachments: false,
      attachmentTypes: [],
      attachmentContentTypes: [],
      keywordsFound: [],
    };
  }
}

/**
 * Formats the metadata into a prompt for the Gemini API
 *
 * @param {Object} metadata - The email metadata to format
 * @returns {string} Formatted prompt
 */
function formatPrompt(metadata) {
  try {
    // Get historical patterns if available
    let historicalContext = "";
    if (metadata.historicalPatterns && metadata.historicalPatterns.count > 0) {
      historicalContext = HistoricalPatterns.formatHistoricalPatternsForPrompt(
        metadata.historicalPatterns
      );
    }

    return `
Based on these email metadata, assess the likelihood that this contains an invoice.
You don't have access to the full content for privacy reasons.
${historicalContext}
Metadata: ${JSON.stringify(metadata, null, 2)}

An invoice typically contains:
- A clear request for payment
- An invoice number or reference
- A specific amount to be paid
- Payment instructions or terms

Just mentioning words like 'invoice', 'bill', 'receipt', etc. is NOT enough to classify as an invoice.
The email must be specifically about a payment document.

On a scale from 0.0 to 1.0, where:
- 0.0 means definitely NOT an invoice
- 1.0 means definitely IS an invoice

Provide ONLY a single number between 0.0 and 1.0 representing your confidence.
Example responses: "0.2", "0.85", "0.99"
`;
  } catch (error) {
    logWithUser(`Error formatting prompt: ${error.message}`, "ERROR");
    // Return a simplified prompt on error
    return `Based on this subject: "${metadata.subject}", on a scale from 0.0 to 1.0, what's the likelihood this is an invoice? Respond with only a number.`;
  }
}

/**
 * Parses the Gemini API response to extract confidence score
 *
 * @param {string} response - The API response to parse
 * @returns {number} Confidence score between 0 and 1
 */
function parseGeminiResponse(response) {
  try {
    // Extract the text response and clean it
    const responseText = response.trim();

    // Try to extract a number from the response
    const confidenceMatch = responseText.match(/(\d+\.\d+|\d+)/);

    if (confidenceMatch) {
      const confidence = parseFloat(confidenceMatch[0]);

      // Validate that it's a number between 0 and 1
      if (!isNaN(confidence) && confidence >= 0 && confidence <= 1) {
        logWithUser(`Gemini confidence score: ${confidence}`, "INFO");
        return confidence;
      }
    }

    // If we couldn't extract a valid confidence score, log warning and return 0
    logWithUser(
      `Could not extract valid confidence score from response: "${responseText}"`,
      "WARNING"
    );
    return 0;
  } catch (error) {
    logWithUser(`Error parsing Gemini response: ${error.message}`, "ERROR");
    return 0;
  }
}

/**
 * Analyzes an email to determine if it contains an invoice
 *
 * @param {GmailMessage} message - The Gmail message to analyze
 * @returns {number} Confidence score between 0 and 1
 */
function analyzeEmail(message) {
  try {
    // Extract metadata (not full content)
    const metadata = extractMetadata(message);

    // Format the prompt
    const formattedPrompt = formatPrompt(metadata);

    // Log the formatted prompt
    logWithUser(
      `Formatted prompt with metadata for Gemini: ${formattedPrompt}`,
      "DEBUG"
    );

    // Call the Gemini API
    const response = callGeminiAPI(formattedPrompt);

    // Parse the response to get confidence score
    return parseGeminiResponse(response);
  } catch (error) {
    logWithUser(`Email analysis failed: ${error.message}`, "ERROR");
    return 0; // Return 0 confidence on error
  }
}

/**
 * Main function to determine if a message contains an invoice using Gemini
 *
 * @param {GmailMessage} message - The Gmail message to analyze
 * @returns {boolean} True if the message likely contains an invoice
 */
function isInvoiceWithGemini(message) {
  try {
    // Check if Gemini is enabled
    if (CONFIG.invoiceDetection !== "gemini") {
      logWithUser(
        "Gemini invoice detection is disabled in configuration",
        "INFO"
      );
      return false;
    }

    // Log the message subject for debugging
    const subject = message.getSubject() || "(no subject)";
    logWithUser(
      `Analyzing message with subject: "${subject}" using Gemini`,
      "INFO"
    );

    // Check if sender domain should be skipped
    const sender = message.getFrom();
    const domain = extractDomain(sender);

    // Log only the domain, not the full email address
    logWithUser(`Message domain: ${domain}`, "INFO");

    if (CONFIG.skipAIForDomains && CONFIG.skipAIForDomains.includes(domain)) {
      logWithUser(
        `Skipping Gemini for domain ${domain} (in skipAIForDomains list)`,
        "INFO"
      );
      return false;
    }

    // Check for PDF attachments if configured
    if (CONFIG.onlyAnalyzePDFs) {
      const attachments = message.getAttachments();
      logWithUser(`Message has ${attachments.length} attachments`, "INFO");

      let hasPDF = false;
      for (const attachment of attachments) {
        const fileName = attachment.getName().toLowerCase();
        const contentType = attachment.getContentType().toLowerCase();
        logWithUser(
          `Checking attachment: ${fileName}, type: ${contentType}`,
          "INFO"
        );

        if (CONFIG.strictPdfCheck) {
          if (fileName.endsWith(".pdf") && contentType.includes("pdf")) {
            hasPDF = true;
            logWithUser(`Found PDF attachment: ${fileName}`, "INFO");
            break;
          }
        } else if (fileName.endsWith(".pdf")) {
          hasPDF = true;
          logWithUser(`Found PDF attachment: ${fileName}`, "INFO");
          break;
        }
      }

      if (!hasPDF) {
        logWithUser(
          `No PDF attachments found, skipping Gemini analysis`,
          "INFO"
        );
        return CONFIG.fallbackToKeywords ? checkKeywords(message) : false;
      }
    }

    // Analyze the email to get confidence score
    logWithUser(`Starting Gemini analysis for message: "${subject}"`, "INFO");
    const confidence = analyzeEmail(message);

    // Compare with threshold
    const isInvoice = confidence >= CONFIG.aiConfidenceThreshold;
    logWithUser(
      `Gemini invoice detection result: ${isInvoice} (confidence: ${confidence}, threshold: ${CONFIG.aiConfidenceThreshold})`,
      "INFO"
    );

    return isInvoice;
  } catch (error) {
    logWithUser(`Gemini invoice detection failed: ${error.message}`, "ERROR");
    // Return false on error, caller can decide to fall back to keywords
    return false;
  }
}

/**
 * Verifies if there is an API key available, using first the configuration and then
 * the script properties as a backup
 *
 * @returns {string|null} The API key if found, null otherwise
 */
function getGeminiApiKey() {
  // First check if the key is in CONFIG and not the placeholder
  if (CONFIG.geminiApiKey && CONFIG.geminiApiKey !== "__GEMINI_API_KEY__") {
    return CONFIG.geminiApiKey;
  }

  // Then try to get it from script properties
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty(
      CONFIG.geminiApiKeyPropertyName
    );
    return apiKey;
  } catch (error) {
    logWithUser(`Error retrieving Gemini API key: ${error.message}`, "ERROR");
    return null;
  }
}

/**
 * Stores the Gemini API key securely in Script Properties
 *
 * @param {string} apiKey - The Gemini API key to store
 * @returns {boolean} True if successful, false otherwise
 */
function storeGeminiApiKey(apiKey) {
  try {
    PropertiesService.getScriptProperties().setProperty(
      CONFIG.geminiApiKeyPropertyName,
      apiKey
    );
    return true;
  } catch (error) {
    logWithUser(`Error storing Gemini API key: ${error.message}`, "ERROR");
    return false;
  }
}

/**
 * Test function to verify Gemini API connectivity
 * This function can be run directly from the Apps Script editor
 * to check if the configured API key is correct
 */
function testGeminiConnection() {
  try {
    const testPrompt =
      "On a scale from 0.0 to 1.0, how likely is this a test? Respond with only a number.";
    const response = callGeminiAPI(testPrompt);

    Logger.log(`Successfully connected to Gemini API. Response: ${response}`);
    return {
      success: true,
      response: response,
    };
  } catch (e) {
    Logger.log(`Error connecting to Gemini API: ${e.message}`);
    return {
      success: false,
      error: e.message,
    };
  }
}

// Make functions available to other modules
var GeminiDetection = {
  isInvoiceWithGemini: isInvoiceWithGemini,
  testGeminiConnection: testGeminiConnection,
  storeGeminiApiKey: storeGeminiApiKey,
};
