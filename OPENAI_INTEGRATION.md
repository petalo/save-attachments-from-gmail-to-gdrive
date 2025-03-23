# OpenAI API Integration for Invoice Detection

This document explains how to use the OpenAI API integration for improved invoice detection in the Gmail Attachment Organizer project.

## Overview

The Gmail Attachment Organizer now includes AI-powered invoice detection using the OpenAI API. This feature analyzes email content to determine if it contains an invoice, providing more accurate detection than the previous keyword-based approach.

## Configuration

The OpenAI integration can be configured in the `CONFIG` object in `Config.gs`:

```javascript
// OpenAI API configuration for invoice detection
openAIEnabled: true, // Enable/disable OpenAI API for invoice detection
openAIApiKey: "__OPENAI_API_KEY__", // Will be replaced during build
openAIApiKeyPropertyName: "openai_api_key", // Property name to store the API key
openAIModel: "gpt-3.5-turbo", // Model to use
openAIMaxTokens: 100, // Maximum tokens for response
openAITemperature: 0.1, // Lower temperature for more deterministic responses
skipAIForDomains: ["newsletter.com", "marketing.com"], // Skip AI for these domains
onlyAnalyzePDFs: true, // Only send emails with PDF attachments to AI
fallbackToKeywords: true, // Fall back to keyword detection if AI fails
aiConfidenceThreshold: 0.7, // Confidence threshold for AI detection
```

## API Key Setup

There are two ways to set up your OpenAI API key:

### 1. Using the .env File (Recommended for Development)

1. Create a `.env` file in the project root (or copy from `.env.example`)
2. Add your OpenAI API key:

   ```text
   OPENAI_API_KEY=your_api_key_here
   ```

3. Run the build process to inject the API key into the script:

   ```bash
   npm run build
   ```

### 2. Using Script Properties (Recommended for Production)

You can store the API key in Google Apps Script's Script Properties:

1. Deploy the script to Google Apps Script
2. Run the `storeOpenAIApiKey` function with your API key:

   ```javascript
   AIDetection.storeOpenAIApiKey('your_api_key_here');
   ```

## Testing the Integration

You can test the OpenAI integration locally before deploying:

1. Make sure you have an OpenAI API key in your `.env` file
2. Install the required dependencies:

   ```bash
   npm install
   ```

3. Run the test script:

   ```bash
   npm run test:openai
   ```

The test script will:

- Make test calls to the OpenAI API with sample emails
- Log the results to the console and to a file in the `logs` directory
- Provide a summary of the test results

## How It Works

The AI-powered invoice detection works as follows:

1. When processing an email, the system checks if OpenAI integration is enabled
2. If enabled, it checks if the email has PDF attachments (if `onlyAnalyzePDFs` is true)
3. It checks if the sender's domain is in the `skipAIForDomains` list
4. If all checks pass, it extracts relevant content from the email (subject, body, sender)
5. It formats the content into a prompt for the OpenAI API
6. It sends the prompt to the OpenAI API with instructions to check in both English and Spanish
7. If the API indicates the email contains an invoice, the system processes it accordingly
8. If the API call fails, it falls back to keyword detection (if `fallbackToKeywords` is true)

### Multilingual Support

The system is designed to detect invoices in both English and Spanish:

- The AI prompt explicitly instructs the model to check for invoice-related terms in both languages
- It looks for terms like "invoice", "bill", "receipt" in English and "factura", "recibo", "pago" in Spanish
- This makes the system effective for international users who receive invoices in different languages
- The test suite includes examples in both English and Spanish to verify multilingual detection

## Optimizations

The implementation includes several optimizations to reduce API costs and improve performance:

1. **PDF-Only Processing**: Only emails with PDF attachments are sent to the API (configurable)
2. **Domain Exclusion**: Emails from specific domains can be excluded from AI analysis
3. **Fallback Mechanism**: If the API call fails, the system can fall back to keyword detection
4. **Token Optimization**: Email content is truncated to reduce token usage
5. **Caching**: The system could be extended to cache results for similar emails

## Troubleshooting

If you encounter issues with the OpenAI integration:

1. Check the logs for error messages
2. Verify that your OpenAI API key is valid and has sufficient credits
3. Make sure the API key is correctly set up in the `.env` file or Script Properties
4. Test the integration locally using the `test:openai` script
5. Check the OpenAI API status at <https://status.openai.com/>

## Security Considerations

The OpenAI integration includes several security features:

1. API keys are never hardcoded in the script
2. API keys can be stored securely in Script Properties
3. API keys in the `.env` file are not committed to version control
4. The system uses HTTPS for all API calls
5. Error messages do not expose sensitive information
