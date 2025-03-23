/**
 * OpenAI Detection functions for Gmail Attachment Organizer
 *
 * This module provides functions to interact with the OpenAI API
 * for improved invoice detection in emails.
 */

/**
 * Makes an API call to OpenAI
 *
 * @param {string} content - The content to analyze
 * @param {Object} options - Additional options for the API call
 * @returns {Object} The API response
 */
function callOpenAIAPI(content, options = {}) {
  try {
    // Get the API key
    const apiKey = getOpenAIApiKey();
    if (!apiKey) {
      throw new Error("OpenAI API key not found");
    }

    // Default options
    const defaultOptions = {
      model: CONFIG.openAIModel,
      temperature: CONFIG.openAITemperature,
      max_tokens: CONFIG.openAIMaxTokens,
    };

    // Merge default options with provided options
    const finalOptions = { ...defaultOptions, ...options };

    // Prepare the request payload
    const payload = {
      model: finalOptions.model,
      messages: [
        {
          role: "system",
          content:
            "You are an assistant that analyzes emails to determine if they contain invoices or bills. Be very precise and conservative in your analysis - only respond with 'yes' if you are highly confident the email is specifically about an invoice, bill, or receipt that requires payment.\n\nCheck the content in both English and Spanish languages.\n\nAn invoice typically contains:\n- A clear request for payment\n- An invoice number or reference\n- A specific amount to be paid\n- Payment instructions or terms\n\nJust mentioning words like 'invoice', 'bill', 'receipt', 'factura', 'recibo', or 'pago' is NOT enough to classify as an invoice. The email must be specifically about a payment document.\n\nRespond with 'yes' ONLY if the email is clearly about an actual invoice or bill. Otherwise, respond with 'no'.",
        },
        {
          role: "user",
          content: content,
        },
      ],
      temperature: finalOptions.temperature,
      max_tokens: finalOptions.max_tokens,
    };

    // Make the API request
    const response = UrlFetchApp.fetch(
      "https://api.openai.com/v1/chat/completions",
      {
        method: "post",
        headers: {
          Authorization: "Bearer " + apiKey,
          "Content-Type": "application/json",
        },
        payload: JSON.stringify(payload),
        muteHttpExceptions: true,
      }
    );

    // Parse and return the response
    const responseData = JSON.parse(response.getContentText());

    // Check for errors in the response
    if (response.getResponseCode() !== 200) {
      throw new Error(
        `API Error: ${responseData.error?.message || "Unknown error"}`
      );
    }

    return responseData;
  } catch (error) {
    logWithUser(`OpenAI API call failed: ${error.message}`, "ERROR");
    throw error;
  }
}

/**
 * Analyzes an email to determine if it contains an invoice
 *
 * @param {GmailMessage} message - The Gmail message to analyze
 * @returns {boolean} True if the message likely contains an invoice
 */
function analyzeEmail(message) {
  try {
    // Get the content to analyze
    const content = getTextContentToAnalyze(message);

    // Format the prompt
    const formattedContent = formatPrompt(content);

    // Call the OpenAI API
    const response = callOpenAIAPI(formattedContent);

    // Parse the response
    return parseResponse(response);
  } catch (error) {
    logWithUser(`Email analysis failed: ${error.message}`, "ERROR");
    throw error;
  }
}

/**
 * Main function to determine if a message contains an invoice using OpenAI
 *
 * @param {GmailMessage} message - The Gmail message to analyze
 * @returns {boolean} True if the message likely contains an invoice
 */
function isInvoiceWithOpenAI(message) {
  try {
    // Check if OpenAI is enabled
    if (CONFIG.invoiceDetection !== "openai") {
      logWithUser(
        "OpenAI invoice detection is disabled in configuration",
        "INFO"
      );
      return false;
    }

    // Log the message subject for debugging
    const subject = message.getSubject() || "(no subject)";
    logWithUser(
      `Analyzing message with subject: "${subject}" using OpenAI`,
      "INFO"
    );

    // Check if sender domain should be skipped
    const sender = message.getFrom();
    const domain = extractDomain(sender);
    logWithUser(`Message sender: ${sender}, domain: ${domain}`, "INFO");

    if (CONFIG.skipAIForDomains && CONFIG.skipAIForDomains.includes(domain)) {
      logWithUser(
        `Skipping OpenAI for domain ${domain} (in skipAIForDomains list)`,
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
          `No PDF attachments found, skipping OpenAI analysis`,
          "INFO"
        );
        return CONFIG.fallbackToKeywords ? checkKeywords(message) : false;
      }
    }

    // Analyze the email
    logWithUser(`Starting OpenAI analysis for message: "${subject}"`, "INFO");
    const isInvoice = analyzeEmail(message);
    logWithUser(`OpenAI invoice detection result: ${isInvoice}`, "INFO");

    return isInvoice;
  } catch (error) {
    logWithUser(`OpenAI invoice detection failed: ${error.message}`, "ERROR");
    // Return false on error, caller can decide to fall back to keywords
    return false;
  }
}

/**
 * Extracts relevant text content from a Gmail message for analysis
 *
 * @param {GmailMessage} message - The Gmail message to extract content from
 * @returns {Object} Object containing subject, body, and historical patterns
 */
function getTextContentToAnalyze(message) {
  try {
    const subject = message.getSubject() || "";
    let body = "";

    // Try to get plain text body
    try {
      body = message.getPlainBody() || "";

      // Truncate body if it's too long (to save tokens)
      if (body.length > 1500) {
        body = body.substring(0, 1500) + "...";
      }
    } catch (e) {
      logWithUser(`Could not get message body: ${e.message}`, "WARNING");
    }

    // Get historical patterns if enabled
    let historicalPatterns = null;
    if (CONFIG.useHistoricalPatterns && CONFIG.manuallyLabeledInvoicesLabel) {
      historicalPatterns = HistoricalPatterns.getHistoricalInvoicePatterns(
        message.getFrom()
      );
    }

    return {
      subject: subject,
      body: body,
      sender: message.getFrom() || "",
      date: message.getDate().toISOString(),
      historicalPatterns: historicalPatterns,
    };
  } catch (error) {
    logWithUser(`Error extracting message content: ${error.message}`, "ERROR");
    // Return minimal content on error
    return {
      subject: message.getSubject() || "",
      body: "",
      sender: message.getFrom() || "",
      date: message.getDate().toISOString(),
    };
  }
}

/**
 * Formats the email content into a prompt for the OpenAI API
 *
 * @param {Object} content - The email content to format
 * @returns {string} Formatted prompt
 */
function formatPrompt(content) {
  try {
    // Get historical patterns if available
    let historicalContext = "";
    if (content.historicalPatterns && content.historicalPatterns.count > 0) {
      historicalContext = HistoricalPatterns.formatHistoricalPatternsForPrompt(
        content.historicalPatterns
      );
    }

    return `
Please analyze this email and determine if it contains an invoice or bill.
Respond with only 'yes' or 'no'.
${historicalContext}
From: ${content.sender}
Date: ${content.date}
Subject: ${content.subject}

${content.body}
`;
  } catch (error) {
    logWithUser(`Error formatting prompt: ${error.message}`, "ERROR");
    // Return a simplified prompt on error
    return `Subject: ${content.subject}\n\nIs this an invoice or bill? Answer yes or no.`;
  }
}

/**
 * Parses the OpenAI API response to determine if the message contains an invoice
 *
 * @param {Object} response - The API response to parse
 * @returns {boolean} True if the message likely contains an invoice
 */
function parseResponse(response) {
  try {
    // Extract the response text
    const responseText = response.choices[0].message.content
      .trim()
      .toLowerCase();

    // Log the raw response for debugging
    logWithUser(`OpenAI raw response: "${responseText}"`, "INFO");

    // Check if the response indicates an invoice
    // We're looking for "yes" or variations like "yes, it is an invoice"
    return responseText.includes("yes");
  } catch (error) {
    logWithUser(`Error parsing API response: ${error.message}`, "ERROR");
    throw error;
  }
}

/**
 * Verifies if there is an API key available, using first the configuration and then
 * the script properties as a backup
 *
 * @returns {string|null} The API key if found, null otherwise
 */
function getOpenAIApiKey() {
  // First check if the key is in CONFIG and not the placeholder
  if (CONFIG.openAIApiKey && CONFIG.openAIApiKey !== "__OPENAI_API_KEY__") {
    return CONFIG.openAIApiKey;
  }

  // Then try to get it from script properties
  try {
    const apiKey = PropertiesService.getScriptProperties().getProperty(
      CONFIG.openAIApiKeyPropertyName
    );
    return apiKey;
  } catch (error) {
    logWithUser(`Error retrieving API key: ${error.message}`, "ERROR");
    return null;
  }
}

/**
 * Stores the OpenAI API key securely in Script Properties
 *
 * @param {string} apiKey - The OpenAI API key to store
 * @returns {boolean} True if successful, false otherwise
 */
function storeOpenAIApiKey(apiKey) {
  try {
    PropertiesService.getScriptProperties().setProperty(
      CONFIG.openAIApiKeyPropertyName,
      apiKey
    );
    return true;
  } catch (error) {
    logWithUser(`Error storing API key: ${error.message}`, "ERROR");
    return false;
  }
}

/**
 * Test function to verify OpenAI API connectivity
 * This function can be run directly from the Apps Script editor
 * to check if the configured API key is correct
 */
function testOpenAIConnection() {
  try {
    const testPrompt = "Is this a test?";
    const response = callOpenAIAPI(testPrompt);

    Logger.log(
      `Successfully connected to OpenAI API. Response: ${JSON.stringify(
        response
      )}`
    );
    return {
      success: true,
      response: response,
    };
  } catch (e) {
    Logger.log(`Error connecting to OpenAI API: ${e.message}`);
    return {
      success: false,
      error: e.message,
    };
  }
}

// Make functions available to other modules
var OpenAIDetection = {
  isInvoiceWithOpenAI: isInvoiceWithOpenAI,
  testOpenAIConnection: testOpenAIConnection,
  storeOpenAIApiKey: storeOpenAIApiKey,
};
