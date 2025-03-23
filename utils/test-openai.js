#!/usr/bin/env node
/**
 * Test script for OpenAI API integration
 *
 * This script tests the OpenAI API integration locally by:
 * 1. Loading the API key from the .env file
 * 2. Making a test call to the OpenAI API
 * 3. Logging the results
 *
 * Usage: node test-openai.js
 */

const dotenv = require("dotenv");
const axios = require("axios");
const fs = require("fs");

// Load environment variables from .env file
dotenv.config();

// Check if OpenAI API key is available
const apiKey = process.env.OPENAI_API_KEY;
if (!apiKey) {
  console.error("Error: OPENAI_API_KEY not found in .env file");
  console.log("Please add your OpenAI API key to the .env file:");
  console.log("OPENAI_API_KEY=your_api_key_here");
  process.exit(1);
}

// Create logs directory if it doesn't exist
const logsDir = "./logs";
if (!fs.existsSync(logsDir)) {
  fs.mkdirSync(logsDir);
}

// Create log file with timestamp
const timestamp = new Date().toISOString().replace(/:/g, "-");
const logFile = `${logsDir}/openai-test-${timestamp}.log`;

// Function to log messages to console and file
function log(message, type = "INFO") {
  const logMessage = `[${new Date().toISOString()}] [${type}] ${message}`;
  console.log(logMessage);
  fs.appendFileSync(logFile, logMessage + "\n");
}

// Sample email content to test
const testEmail = {
  subject: "Your Invoice #12345 from Acme Corp",
  body: "Dear Customer,\n\nAttached is your invoice #12345 for your recent purchase. Please remit payment at your earliest convenience.\n\nThank you for your business.\n\nRegards,\nAcme Corp",
  sender: "billing@acmecorp.com",
  date: new Date().toISOString(),
};

// Format the prompt
function formatPrompt(content) {
  return `
Please analyze this email and determine if it contains an invoice or bill.
Respond with only 'yes' or 'no'.

From: ${content.sender}
Date: ${content.date}
Subject: ${content.subject}

${content.body}
`;
}

// Make API call to OpenAI
async function callOpenAI(content) {
  log("Making API call to OpenAI...");

  try {
    // Default options
    const options = {
      model: "gpt-3.5-turbo",
      temperature: 0.1,
      max_tokens: 100,
    };

    // Prepare the request payload
    const payload = {
      model: options.model,
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
      temperature: options.temperature,
      max_tokens: options.max_tokens,
    };

    // Make the API request
    log(`Using model: ${options.model}`);
    const response = await axios.post(
      "https://api.openai.com/v1/chat/completions",
      payload,
      {
        headers: {
          Authorization: `Bearer ${apiKey}`,
          "Content-Type": "application/json",
        },
      }
    );

    // Log the response
    log("API call successful");
    log(`Response status: ${response.status}`);
    log(`Response data: ${JSON.stringify(response.data, null, 2)}`);

    // Extract and return the response text
    const responseText = response.data.choices[0].message.content
      .trim()
      .toLowerCase();
    log(`Response text: "${responseText}"`);

    // Determine if it's an invoice
    const isInvoice = responseText.includes("yes");
    log(`Is invoice: ${isInvoice}`);

    return {
      success: true,
      isInvoice: isInvoice,
      rawResponse: response.data,
    };
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
  log("Testing with positive example (invoice)...");
  const formattedPrompt = formatPrompt(testEmail);
  log(`Formatted prompt:\n${formattedPrompt}`);
  const result = await callOpenAI(formattedPrompt);
  log(`Test result: ${JSON.stringify(result, null, 2)}`);
  return result;
}

// Test with a negative example (not an invoice)
async function testNegativeExample() {
  log("Testing with negative example (not an invoice)...");
  const nonInvoiceEmail = {
    ...testEmail,
    subject: "Team meeting tomorrow",
    body: "Hi team,\n\nJust a reminder that we have a team meeting tomorrow at 10am.\n\nBest regards,\nJohn",
  };
  const formattedPrompt = formatPrompt(nonInvoiceEmail);
  log(`Formatted prompt:\n${formattedPrompt}`);
  const result = await callOpenAI(formattedPrompt);
  log(`Test result: ${JSON.stringify(result, null, 2)}`);
  return result;
}

// Test with a Spanish positive example (invoice)
async function testSpanishPositiveExample() {
  log("Testing with Spanish positive example (invoice)...");
  const spanishInvoiceEmail = {
    subject: "Tu Factura #12345 de Empresa SA",
    body: "Estimado Cliente,\n\nAdjuntamos la factura #12345 por su reciente compra. Por favor, realice el pago a la mayor brevedad posible.\n\nGracias por su confianza.\n\nSaludos,\nEmpresa SA",
    sender: "facturacion@empresasa.es",
    date: new Date().toISOString(),
  };
  const formattedPrompt = formatPrompt(spanishInvoiceEmail);
  log(`Formatted prompt:\n${formattedPrompt}`);
  const result = await callOpenAI(formattedPrompt);
  log(`Test result: ${JSON.stringify(result, null, 2)}`);
  return result;
}

// Run the tests
async function runTests() {
  log("Starting OpenAI API integration test");
  log(
    `API Key: ${apiKey.substring(0, 3)}...${apiKey.substring(
      apiKey.length - 3
    )}`
  );

  try {
    // Test with English positive example
    const positiveResult = await testPositiveExample();

    // Test with negative example
    const negativeResult = await testNegativeExample();

    // Test with Spanish positive example
    const spanishPositiveResult = await testSpanishPositiveExample();

    // Log summary
    log("\nTest Summary:");
    log(
      `English positive example (should be invoice): ${
        positiveResult.success
          ? positiveResult.isInvoice
            ? "PASS ✅"
            : "FAIL ❌"
          : "ERROR ❌"
      }`
    );
    log(
      `Negative example (should not be invoice): ${
        negativeResult.success
          ? !negativeResult.isInvoice
            ? "PASS ✅"
            : "FAIL ❌"
          : "ERROR ❌"
      }`
    );
    log(
      `Spanish positive example (should be invoice): ${
        spanishPositiveResult.success
          ? spanishPositiveResult.isInvoice
            ? "PASS ✅"
            : "FAIL ❌"
          : "ERROR ❌"
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
