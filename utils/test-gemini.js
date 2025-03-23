#!/usr/bin/env node
/**
 * Test script for Gemini API integration
 *
 * This script tests the Gemini API integration locally by:
 * 1. Loading the API key from the .env file
 * 2. Making a test call to the Gemini API
 * 3. Logging the results
 *
 * Usage: node test-gemini.js
 */

const dotenv = require("dotenv");
const axios = require("axios");
const fs = require("fs");

// Load environment variables from .env file
dotenv.config();

// Check if Gemini API key is available
const apiKey = process.env.GEMINI_API_KEY;
if (!apiKey) {
  console.error("Error: GEMINI_API_KEY not found in .env file");
  console.log("Please add your Gemini API key to the .env file:");
  console.log("GEMINI_API_KEY=your_api_key_here");
  process.exit(1);
}

// Create logs directory if it doesn't exist
const logsDir = "./logs";
if (!fs.existsSync(logsDir)) {
  fs.mkdirSync(logsDir);
}

// Create log file with timestamp
const timestamp = new Date().toISOString().replace(/:/g, "-");
const logFile = `${logsDir}/gemini-test-${timestamp}.log`;

// Function to log messages to console and file
function log(message, type = "INFO") {
  const logMessage = `[${new Date().toISOString()}] [${type}] ${message}`;
  console.log(logMessage);
  fs.appendFileSync(logFile, logMessage + "\n");
}

// Sample email metadata to test
const testEmailMetadata = {
  subject: "Your Invoice #12345 from Acme Corp",
  sender: "billing@acmecorp.com",
  senderDomain: "acmecorp.com",
  date: new Date().toISOString(),
  hasAttachments: true,
  attachmentTypes: ["pdf"],
  attachmentContentTypes: ["application/pdf"],
  keywordsFound: ["invoice", "#12345", "payment", "$299.99"],
};

// Format the prompt
function formatPrompt(metadata) {
  return `
Based on these email metadata, assess the likelihood that this contains an invoice.
You don't have access to the full content for privacy reasons.

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
}

// Make API call to Gemini
async function callGeminiAPI(prompt) {
  log("Making API call to Gemini...");

  try {
    // Default options
    const options = {
      model: "gemini-2.0-flash", // Try a newer model
      temperature: 0.05,
      maxOutputTokens: 10,
    };

    // Prepare the request payload
    const payload = {
      contents: [
        {
          parts: [{ text: prompt }],
        },
      ],
      generationConfig: {
        temperature: options.temperature,
        maxOutputTokens: options.maxOutputTokens,
      },
    };

    // Try v1 API first (newer version)
    let url = `https://generativelanguage.googleapis.com/v1/models/${options.model}:generateContent?key=${apiKey}`;

    log(`Trying v1 API endpoint with model: ${options.model}`);

    try {
      // Make the API request to v1 endpoint
      const response = await axios.post(url, payload, {
        headers: {
          "Content-Type": "application/json",
        },
      });

      // Log the response
      log("API call successful (v1 API)");
      log(`Response status: ${response.status}`);
      log(`Response data: ${JSON.stringify(response.data, null, 2)}`);

      // Extract the text response
      const responseText =
        response.data.candidates[0].content.parts[0].text.trim();
      log(`Response text: "${responseText}"`);

      // Try to extract a confidence score
      const confidenceMatch = responseText.match(/(\d+\.\d+|\d+)/);
      let confidence = null;

      if (confidenceMatch) {
        confidence = parseFloat(confidenceMatch[0]);

        // Validate that it's a number between 0 and 1
        if (!isNaN(confidence) && confidence >= 0 && confidence <= 1) {
          log(`Confidence score: ${confidence}`);
        } else {
          log(`Invalid confidence score: ${confidence}`, "WARNING");
          confidence = null;
        }
      } else {
        log(
          `Could not extract confidence score from response: "${responseText}"`,
          "WARNING"
        );
      }

      return {
        success: true,
        confidence: confidence,
        rawResponse: response.data,
        apiVersion: "v1",
      };
    } catch (v1Error) {
      // Log the error but don't throw yet
      log(`v1 API call failed: ${v1Error.message}`, "WARNING");
      if (v1Error.response) {
        log(`v1 Response status: ${v1Error.response.status}`, "WARNING");
        log(
          `v1 Response data: ${JSON.stringify(v1Error.response.data, null, 2)}`,
          "WARNING"
        );
      }

      // Try v1beta API as fallback
      log("Falling back to v1beta API endpoint...");
      url = `https://generativelanguage.googleapis.com/v1beta/models/${options.model}:generateContent?key=${apiKey}`;

      log(`Trying v1beta API endpoint with model: ${options.model}`);

      try {
        // Make the API request to v1beta endpoint
        const response = await axios.post(url, payload, {
          headers: {
            "Content-Type": "application/json",
          },
        });

        // Log the response
        log("API call successful (v1beta API)");
        log(`Response status: ${response.status}`);
        log(`Response data: ${JSON.stringify(response.data, null, 2)}`);

        // Extract the text response
        const responseText =
          response.data.candidates[0].content.parts[0].text.trim();
        log(`Response text: "${responseText}"`);

        // Try to extract a confidence score
        const confidenceMatch = responseText.match(/(\d+\.\d+|\d+)/);
        let confidence = null;

        if (confidenceMatch) {
          confidence = parseFloat(confidenceMatch[0]);

          // Validate that it's a number between 0 and 1
          if (!isNaN(confidence) && confidence >= 0 && confidence <= 1) {
            log(`Confidence score: ${confidence}`);
          } else {
            log(`Invalid confidence score: ${confidence}`, "WARNING");
            confidence = null;
          }
        } else {
          log(
            `Could not extract confidence score from response: "${responseText}"`,
            "WARNING"
          );
        }

        return {
          success: true,
          confidence: confidence,
          rawResponse: response.data,
          apiVersion: "v1beta",
        };
      } catch (v1betaError) {
        // Log the error
        log(`v1beta API call also failed: ${v1betaError.message}`, "ERROR");
        if (v1betaError.response) {
          log(
            `v1beta Response status: ${v1betaError.response.status}`,
            "ERROR"
          );
          log(
            `v1beta Response data: ${JSON.stringify(
              v1betaError.response.data,
              null,
              2
            )}`,
            "ERROR"
          );
        }

        // Re-throw the error since both API versions failed
        throw v1betaError;
      }
    }
  } catch (error) {
    log(`API call failed: ${error.message}`, "ERROR");
    if (error.response) {
      log(`Response status: ${error.response.status}`, "ERROR");
      log(
        `Response data: ${JSON.stringify(error.response.data, null, 2)}`,
        "ERROR"
      );
    }
    return {
      success: false,
      error: error.message,
      details: error.response ? error.response.data : null,
    };
  }
}

// Test with a positive example (invoice)
async function testPositiveExample() {
  log("Testing with positive example (invoice metadata)...");
  const formattedPrompt = formatPrompt(testEmailMetadata);
  log(`Formatted prompt:\n${formattedPrompt}`);
  const result = await callGeminiAPI(formattedPrompt);
  log(`Test result: ${JSON.stringify(result, null, 2)}`);
  return result;
}

// Test with a negative example (not an invoice)
async function testNegativeExample() {
  log("Testing with negative example (not an invoice)...");
  const nonInvoiceMetadata = {
    ...testEmailMetadata,
    subject: "Team meeting tomorrow",
    keywordsFound: ["meeting", "agenda", "tomorrow", "10am"],
  };
  const formattedPrompt = formatPrompt(nonInvoiceMetadata);
  log(`Formatted prompt:\n${formattedPrompt}`);
  const result = await callGeminiAPI(formattedPrompt);
  log(`Test result: ${JSON.stringify(result, null, 2)}`);
  return result;
}

// Test with a Spanish positive example (invoice)
async function testSpanishPositiveExample() {
  log("Testing with Spanish positive example (invoice)...");
  const spanishInvoiceMetadata = {
    subject: "Tu Factura #12345 de Empresa SA",
    sender: "facturacion@empresasa.es",
    senderDomain: "empresasa.es",
    date: new Date().toISOString(),
    hasAttachments: true,
    attachmentTypes: ["pdf"],
    attachmentContentTypes: ["application/pdf"],
    keywordsFound: ["factura", "#12345", "pago", "€299,99"],
  };
  const formattedPrompt = formatPrompt(spanishInvoiceMetadata);
  log(`Formatted prompt:\n${formattedPrompt}`);
  const result = await callGeminiAPI(formattedPrompt);
  log(`Test result: ${JSON.stringify(result, null, 2)}`);
  return result;
}

// Test with a borderline example
async function testBorderlineExample() {
  log("Testing with borderline example...");
  const borderlineMetadata = {
    subject: "Information about your recent purchase",
    sender: "info@somestore.com",
    senderDomain: "somestore.com",
    date: new Date().toISOString(),
    hasAttachments: true,
    attachmentTypes: ["pdf"],
    attachmentContentTypes: ["application/pdf"],
    keywordsFound: ["purchase", "order", "receipt", "information"],
  };
  const formattedPrompt = formatPrompt(borderlineMetadata);
  log(`Formatted prompt:\n${formattedPrompt}`);
  const result = await callGeminiAPI(formattedPrompt);
  log(`Test result: ${JSON.stringify(result, null, 2)}`);
  return result;
}

// Run the tests
async function runTests() {
  log("Starting Gemini API integration test");
  log(
    `API Key: ${apiKey.substring(0, 3)}...${apiKey.substring(
      apiKey.length - 3
    )}`
  );

  try {
    // Test with positive example
    const positiveResult = await testPositiveExample();

    // Test with negative example
    const negativeResult = await testNegativeExample();

    // Test with Spanish positive example
    const spanishPositiveResult = await testSpanishPositiveExample();

    // Test with borderline example
    const borderlineResult = await testBorderlineExample();

    // Log summary
    log("\nTest Summary:");

    // For positive example (should have high confidence)
    const positiveConfidence = positiveResult.confidence;
    log(
      `Positive example (should be invoice): ${
        positiveResult.success
          ? positiveConfidence !== null
            ? positiveConfidence >= 0.7
              ? `PASS ✅ (${positiveConfidence})`
              : `FAIL ❌ (${positiveConfidence} - too low)`
            : "FAIL ❌ (no valid confidence score)"
          : "ERROR ❌"
      }`
    );

    // For negative example (should have low confidence)
    const negativeConfidence = negativeResult.confidence;
    log(
      `Negative example (should not be invoice): ${
        negativeResult.success
          ? negativeConfidence !== null
            ? negativeConfidence <= 0.3
              ? `PASS ✅ (${negativeConfidence})`
              : `FAIL ❌ (${negativeConfidence} - too high)`
            : "FAIL ❌ (no valid confidence score)"
          : "ERROR ❌"
      }`
    );

    // For Spanish positive example (should have high confidence)
    const spanishConfidence = spanishPositiveResult.confidence;
    log(
      `Spanish positive example (should be invoice): ${
        spanishPositiveResult.success
          ? spanishConfidence !== null
            ? spanishConfidence >= 0.7
              ? `PASS ✅ (${spanishConfidence})`
              : `FAIL ❌ (${spanishConfidence} - too low)`
            : "FAIL ❌ (no valid confidence score)"
          : "ERROR ❌"
      }`
    );

    // For borderline example (just informational)
    const borderlineConfidence = borderlineResult.confidence;
    log(
      `Borderline example (informational): ${
        borderlineResult.success
          ? borderlineConfidence !== null
            ? `Score: ${borderlineConfidence}`
            : "No valid confidence score"
          : "ERROR"
      }`
    );

    log("\nTest completed. See logs in: " + logFile);
  } catch (error) {
    log(`Test failed with error: ${error.message}`, "ERROR");
    if (error.stack) {
      log(`Stack trace: ${error.stack}`, "ERROR");
    }
  }
}

// Run the tests
runTests();
