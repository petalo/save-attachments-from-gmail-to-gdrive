/**
 * API Key Verification functions for Gmail Attachment Organizer
 *
 * This module provides functions to verify API keys for both Gemini and OpenAI.
 * It can be used to check if the API keys are valid and properly configured.
 */

/**
 * Verifies if the Gemini API key is valid and working
 *
 * @returns {Object} Object with success status, details, and API information
 */
function verifyGeminiAPIKey() {
  try {
    logWithUser("Starting Gemini API key verification...", "INFO");

    // Step 1: Check if the API key exists in configuration or script properties
    const apiKey = getGeminiApiKey();
    if (!apiKey) {
      return {
        success: false,
        message:
          "Gemini API key not found in configuration or script properties",
        details: {
          configValue:
            CONFIG.geminiApiKey === "__GEMINI_API_KEY__"
              ? "Not set (placeholder)"
              : "Set in CONFIG",
          scriptPropertyName: CONFIG.geminiApiKeyPropertyName,
          scriptPropertyValue: "Not available (null)",
        },
      };
    }

    // Step 2: Mask the API key for logging (show only first 4 and last 4 characters)
    const maskedKey = maskAPIKey(apiKey);
    logWithUser(`Found Gemini API key: ${maskedKey}`, "INFO");

    // Step 3: Test the API connection
    logWithUser("Testing Gemini API connection...", "INFO");
    const testResult = GeminiDetection.testGeminiConnection();

    if (testResult.success) {
      return {
        success: true,
        message: "Gemini API key is valid and working correctly",
        details: {
          apiKeySource:
            CONFIG.geminiApiKey !== "__GEMINI_API_KEY__"
              ? "CONFIG"
              : "Script Properties",
          apiKeyLength: apiKey.length,
          maskedKey: maskedKey,
          model: CONFIG.geminiModel,
          testResponse: testResult.response,
        },
      };
    } else {
      return {
        success: false,
        message: "Gemini API key validation failed",
        details: {
          apiKeySource:
            CONFIG.geminiApiKey !== "__GEMINI_API_KEY__"
              ? "CONFIG"
              : "Script Properties",
          apiKeyLength: apiKey.length,
          maskedKey: maskedKey,
          model: CONFIG.geminiModel,
          error: testResult.error,
        },
      };
    }
  } catch (error) {
    logWithUser(`Error verifying Gemini API key: ${error.message}`, "ERROR");
    return {
      success: false,
      message: `Error verifying Gemini API key: ${error.message}`,
      details: {
        error: error.message,
        stack: error.stack,
      },
    };
  }
}

/**
 * Verifies if the OpenAI API key is valid and working
 *
 * @returns {Object} Object with success status, details, and API information
 */
function verifyOpenAIAPIKey() {
  try {
    logWithUser("Starting OpenAI API key verification...", "INFO");

    // Step 1: Check if the API key exists in configuration or script properties
    const apiKey = getOpenAIApiKey();
    if (!apiKey) {
      return {
        success: false,
        message:
          "OpenAI API key not found in configuration or script properties",
        details: {
          configValue:
            CONFIG.openAIApiKey === "__OPENAI_API_KEY__"
              ? "Not set (placeholder)"
              : "Set in CONFIG",
          scriptPropertyName: CONFIG.openAIApiKeyPropertyName,
          scriptPropertyValue: "Not available (null)",
        },
      };
    }

    // Step 2: Mask the API key for logging (show only first 4 and last 4 characters)
    const maskedKey = maskAPIKey(apiKey);
    logWithUser(`Found OpenAI API key: ${maskedKey}`, "INFO");

    // Step 3: Test the API connection
    logWithUser("Testing OpenAI API connection...", "INFO");
    const testResult = OpenAIDetection.testOpenAIConnection();

    if (testResult.success) {
      return {
        success: true,
        message: "OpenAI API key is valid and working correctly",
        details: {
          apiKeySource:
            CONFIG.openAIApiKey !== "__OPENAI_API_KEY__"
              ? "CONFIG"
              : "Script Properties",
          apiKeyLength: apiKey.length,
          maskedKey: maskedKey,
          model: CONFIG.openAIModel,
          testResponse: "Response available (not shown for brevity)",
        },
      };
    } else {
      return {
        success: false,
        message: "OpenAI API key validation failed",
        details: {
          apiKeySource:
            CONFIG.openAIApiKey !== "__OPENAI_API_KEY__"
              ? "CONFIG"
              : "Script Properties",
          apiKeyLength: apiKey.length,
          maskedKey: maskedKey,
          model: CONFIG.openAIModel,
          error: testResult.error,
        },
      };
    }
  } catch (error) {
    logWithUser(`Error verifying OpenAI API key: ${error.message}`, "ERROR");
    return {
      success: false,
      message: `Error verifying OpenAI API key: ${error.message}`,
      details: {
        error: error.message,
        stack: error.stack,
      },
    };
  }
}

/**
 * Verifies both Gemini and OpenAI API keys
 *
 * @returns {Object} Object with verification results for both APIs
 */
function verifyAllAPIKeys() {
  const results = {
    gemini: verifyGeminiAPIKey(),
    openai: verifyOpenAIAPIKey(),
    timestamp: new Date().toISOString(),
    config: {
      invoiceDetection: CONFIG.invoiceDetection,
      geminiModel: CONFIG.geminiModel,
      openAIModel: CONFIG.openAIModel,
    },
  };

  // Log a summary of the results
  if (results.gemini.success && results.openai.success) {
    logWithUser(
      "✅ Both Gemini and OpenAI API keys are valid and working",
      "INFO"
    );
  } else if (results.gemini.success) {
    logWithUser(
      "⚠️ Gemini API key is valid, but OpenAI API key validation failed",
      "WARNING"
    );
  } else if (results.openai.success) {
    logWithUser(
      "⚠️ OpenAI API key is valid, but Gemini API key validation failed",
      "WARNING"
    );
  } else {
    logWithUser(
      "❌ Both Gemini and OpenAI API key validations failed",
      "ERROR"
    );
  }

  return results;
}

/**
 * Helper function to mask an API key for safe logging
 * Shows only the first 4 and last 4 characters, with the rest masked
 *
 * @param {string} apiKey - The API key to mask
 * @returns {string} The masked API key
 */
function maskAPIKey(apiKey) {
  if (!apiKey || apiKey.length < 8) {
    return "Invalid key (too short)";
  }

  const firstFour = apiKey.substring(0, 4);
  const lastFour = apiKey.substring(apiKey.length - 4);
  const maskedPortion = "*".repeat(Math.min(apiKey.length - 8, 10));

  return `${firstFour}${maskedPortion}${lastFour} (${apiKey.length} chars)`;
}

// Make functions available to other modules
var VerifyAPIKeys = {
  verifyGeminiAPIKey: verifyGeminiAPIKey,
  verifyOpenAIAPIKey: verifyOpenAIAPIKey,
  verifyAllAPIKeys: verifyAllAPIKeys,
};
