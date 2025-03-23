/**
 * Test script to verify API keys for both Gemini and OpenAI
 *
 * This script can be run directly with Node.js to test API keys
 * without needing to deploy to Google Apps Script.
 *
 * Usage:
 *   node test-api-keys.js
 *
 * Make sure you have the API keys in your .env file:
 *   OPENAI_API_KEY=your_openai_key
 *   GEMINI_API_KEY=your_gemini_key
 */

require("dotenv").config();
const axios = require("axios");

// Colors for console output
const colors = {
  reset: "\x1b[0m",
  red: "\x1b[31m",
  green: "\x1b[32m",
  yellow: "\x1b[33m",
  blue: "\x1b[34m",
  magenta: "\x1b[35m",
  cyan: "\x1b[36m",
  white: "\x1b[37m",
};

// Log with color
function log(message, color = colors.white) {
  console.log(`${color}${message}${colors.reset}`);
}

// Mask API key for safe logging
function maskAPIKey(apiKey) {
  if (!apiKey || apiKey.length < 8) {
    return "Invalid key (too short)";
  }

  const firstFour = apiKey.substring(0, 4);
  const lastFour = apiKey.substring(apiKey.length - 4);
  const maskedPortion = "*".repeat(Math.min(apiKey.length - 8, 10));

  return `${firstFour}${maskedPortion}${lastFour} (${apiKey.length} chars)`;
}

// Test OpenAI API key
async function testOpenAIKey() {
  log("\n🔑 Testing OpenAI API Key...", colors.cyan);

  const apiKey = process.env.OPENAI_API_KEY;
  if (!apiKey) {
    log("❌ OpenAI API key not found in .env file", colors.red);
    return false;
  }

  log(`📝 Found OpenAI API key: ${maskAPIKey(apiKey)}`, colors.blue);

  try {
    const response = await axios.post(
      "https://api.openai.com/v1/chat/completions",
      {
        model: "gpt-3.5-turbo",
        messages: [
          {
            role: "system",
            content:
              "You are a helpful assistant that responds with only 'yes' or 'no'.",
          },
          {
            role: "user",
            content: "Is this a test?",
          },
        ],
        temperature: 0.05,
        max_tokens: 10,
      },
      {
        headers: {
          Authorization: `Bearer ${apiKey}`,
          "Content-Type": "application/json",
        },
      }
    );

    if (response.status === 200) {
      const responseText = response.data.choices[0].message.content.trim();
      log(
        `✅ OpenAI API key is valid! Response: "${responseText}"`,
        colors.green
      );
      return true;
    } else {
      log(`❌ OpenAI API returned status ${response.status}`, colors.red);
      return false;
    }
  } catch (error) {
    log(`❌ OpenAI API test failed: ${error.message}`, colors.red);
    if (error.response) {
      log(`Error details: ${JSON.stringify(error.response.data)}`, colors.red);
    }
    return false;
  }
}

// Test Gemini API key
async function testGeminiKey() {
  log("\n🔑 Testing Gemini API Key...", colors.magenta);

  const apiKey = process.env.GEMINI_API_KEY;
  if (!apiKey) {
    log("❌ Gemini API key not found in .env file", colors.red);
    return false;
  }

  log(`📝 Found Gemini API key: ${maskAPIKey(apiKey)}`, colors.blue);

  try {
    // Try with v1 API first (newer version)
    const model = "gemini-2.0-flash";
    let url = `https://generativelanguage.googleapis.com/v1/models/${model}:generateContent?key=${apiKey}`;

    try {
      const response = await axios.post(
        url,
        {
          contents: [
            {
              parts: [
                {
                  text: "On a scale from 0.0 to 1.0, how likely is this a test? Respond with only a number.",
                },
              ],
            },
          ],
          generationConfig: {
            temperature: 0.05,
            maxOutputTokens: 10,
          },
        },
        {
          headers: {
            "Content-Type": "application/json",
          },
        }
      );

      if (response.status === 200) {
        const responseText =
          response.data.candidates[0].content.parts[0].text.trim();
        log(
          `✅ Gemini API key is valid! Response: "${responseText}" (using v1 API)`,
          colors.green
        );
        return true;
      }
    } catch (v1Error) {
      log(`Note: v1 API attempt failed: ${v1Error.message}`, colors.yellow);
      log(`Trying v1beta API endpoint instead...`, colors.yellow);

      // Fall back to v1beta API
      url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;

      try {
        const response = await axios.post(
          url,
          {
            contents: [
              {
                parts: [
                  {
                    text: "On a scale from 0.0 to 1.0, how likely is this a test? Respond with only a number.",
                  },
                ],
              },
            ],
            generationConfig: {
              temperature: 0.05,
              maxOutputTokens: 10,
            },
          },
          {
            headers: {
              "Content-Type": "application/json",
            },
          }
        );

        if (response.status === 200) {
          const responseText =
            response.data.candidates[0].content.parts[0].text.trim();
          log(
            `✅ Gemini API key is valid! Response: "${responseText}" (using v1beta API)`,
            colors.green
          );
          return true;
        } else {
          log(`❌ Gemini API returned status ${response.status}`, colors.red);
          return false;
        }
      } catch (v1betaError) {
        log(
          `❌ v1beta API attempt also failed: ${v1betaError.message}`,
          colors.red
        );
        if (v1betaError.response) {
          log(
            `Error details: ${JSON.stringify(v1betaError.response.data)}`,
            colors.red
          );
        }
        return false;
      }
    }
  } catch (error) {
    log(`❌ Gemini API test failed: ${error.message}`, colors.red);
    if (error.response) {
      log(`Error details: ${JSON.stringify(error.response.data)}`, colors.red);
    }
    return false;
  }
}

// Main function
async function main() {
  log("🔍 API Key Verification Tool 🔍", colors.cyan);
  log("================================", colors.cyan);

  // Check if axios is installed
  try {
    require.resolve("axios");
  } catch (e) {
    log("❌ The 'axios' package is required but not installed.", colors.red);
    log("Please run: npm install axios", colors.yellow);
    return;
  }

  // Check if dotenv is installed
  try {
    require.resolve("dotenv");
  } catch (e) {
    log("❌ The 'dotenv' package is required but not installed.", colors.red);
    log("Please run: npm install dotenv", colors.yellow);
    return;
  }

  // Test both API keys
  const openaiResult = await testOpenAIKey();
  const geminiResult = await testGeminiKey();

  // Summary
  log("\n📋 Summary:", colors.cyan);
  log("================================", colors.cyan);

  if (openaiResult) {
    log("✅ OpenAI API key: Valid", colors.green);
  } else {
    log("❌ OpenAI API key: Invalid or not found", colors.red);
  }

  if (geminiResult) {
    log("✅ Gemini API key: Valid", colors.green);
  } else {
    log("❌ Gemini API key: Invalid or not found", colors.red);
  }

  log(
    "\n💡 Tip: Make sure your API keys are correctly set in the .env file:",
    colors.yellow
  );
  log("OPENAI_API_KEY=your_openai_key", colors.yellow);
  log("GEMINI_API_KEY=your_gemini_key", colors.yellow);
}

// Run the main function
main().catch((error) => {
  log(`❌ Unexpected error: ${error.message}`, colors.red);
});
