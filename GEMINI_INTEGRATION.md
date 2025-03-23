# Gemini API Integration for Invoice Detection

This document describes the integration of Google's Gemini API for improved invoice detection in the Gmail Attachment Organizer project. The integration provides a privacy-focused approach to invoice detection by analyzing email metadata rather than sending full email content to external AI services.

## Overview

The Gemini integration uses Google's Gemini Pro model to analyze email metadata and determine if an email likely contains an invoice. This approach offers several advantages:

1. **Privacy-focused**: Only sends metadata, not full email content
2. **Improved accuracy**: Uses confidence scores for more precise detection
3. **Multilingual support**: Works with both English and Spanish invoices
4. **Fallback mechanisms**: Can fall back to OpenAI or keyword detection if needed

## Configuration

The Gemini integration can be configured in the `Config.gs` file:

```javascript
// Gemini settings (recommended)
geminiEnabled: true,                   // Enable/disable Gemini API for invoice detection
geminiApiKey: "__GEMINI_API_KEY__",    // Will be replaced during build
geminiApiKeyPropertyName: "gemini_api_key", // Property name to store the API key
geminiModel: "gemini-pro",             // Model to use
geminiMaxTokens: 10,                   // Maximum tokens for response (we only need a number)
geminiTemperature: 0.05,               // Very low temperature for more conservative responses

// Shared AI settings
skipAIForDomains: ["newsletter.com", "marketing.com"], // Skip AI for these domains
onlyAnalyzePDFs: true,                 // Only send emails with PDF attachments to AI
strictPdfCheck: true,                  // Check both file extension and MIME type for PDFs
fallbackToKeywords: true,              // Fall back to keyword detection if AI fails
aiConfidenceThreshold: 0.9,            // High confidence threshold to reduce false positives (0.0-1.0)
```

## How It Works

The Gemini integration follows these steps:

1. **Extract Metadata**: Instead of sending the full email content, the system extracts only necessary metadata:
   - Subject line
   - Sender domain
   - Attachment types
   - Keywords found in the email (without sending the full body)

2. **Format Prompt**: The metadata is formatted into a prompt that asks Gemini to assess the likelihood that the email contains an invoice.

3. **Get Confidence Score**: Gemini returns a confidence score between 0.0 and 1.0.

4. **Apply Threshold**: The system compares the confidence score with the configured threshold (`aiConfidenceThreshold`) to determine if the email contains an invoice.

5. **Process Attachments**: If the email is determined to contain an invoice, the attachments are processed accordingly.

## Privacy Considerations

The Gemini integration is designed with privacy in mind:

- Only sends metadata, not the full email content
- Anonymizes personal identifiers like email addresses
- Only shares domain names, not full email addresses
- Extracts only relevant keywords from the email body
- Does not send any attachment content to the API
- Can be configured to skip specific domains

## Sample Metadata Sent to Gemini

Here's an example of the JSON metadata that is sent to Gemini for invoice detection, including historical patterns data:

```json
{
  "subject": "Invoice #12345 - Example Company",
  "senderDomain": "example.com",
  "date": "2025-03-15T10:30:00.000Z",
  "hasAttachments": true,
  "attachmentTypes": [
    "pdf"
  ],
  "attachmentContentTypes": [
    "application/pdf"
  ],
  "keywordsFound": [
    "invoice",
    "#12345",
    "payment",
    "$299.99"
  ],
  "historicalPatterns": {
    "count": 3,
    "subjectPatterns": {
      "commonPrefix": "Invoice",
      "commonSuffix": "Example Company",
      "containsInvoiceTerms": true,
      "hasNumericPatterns": true
    },
    "datePatterns": {
      "frequency": "monthly",
      "averageIntervalDays": 30,
      "sameDayOfMonth": true,
      "dayOfMonth": 15
    },
    "rawSubjects": [
      "Invoice #12342 - Example Company",
      "Invoice #12343 - Example Company",
      "Invoice #12344 - Example Company"
    ],
    "rawDates": [
      "2024-12-15T10:30:00.000Z",
      "2025-01-15T10:30:00.000Z",
      "2025-02-15T10:30:00.000Z"
    ]
  }
}
```

Note that the full email address is never sent to Gemini, only the domain name. This ensures that personal identifiers are protected while still allowing the AI to make accurate determinations about whether the email contains an invoice.

The `historicalPatterns` section contains valuable information about previous emails from the same sender that were identified as invoices, including:

- `count`: Number of historical invoices from this sender
- `subjectPatterns`: Analysis of common patterns in subject lines
- `datePatterns`: Analysis of timing patterns, including average interval between invoices
- `rawSubjects`: Anonymized list of previous invoice subject lines
- `rawDates`: Timestamps of previous invoices

This historical data helps improve the accuracy of invoice detection without compromising privacy.

## Testing

You can test the Gemini integration locally using the provided test script:

1. Add your Gemini API key to the `.env` file:

   ```text
   GEMINI_API_KEY=your_api_key_here
   ```

2. Run the test script:

   ```bash
   npm run test:gemini
   ```

The test script will make API calls with sample email metadata in both English and Spanish, and log the results to the console and to a file in the `logs` directory.

## Obtaining a Gemini API Key

To use the Gemini integration, you need to obtain a Gemini API key:

1. Go to the [Google AI Studio](https://makersuite.google.com/app/apikey)
2. Sign in with your Google account
3. Create a new API key
4. Add the API key to your `.env` file as `GEMINI_API_KEY=your_api_key_here`

## Fallback Mechanisms

The Gemini integration includes several fallback mechanisms:

1. **OpenAI Fallback**: If Gemini fails, the system can fall back to OpenAI if enabled.
2. **Keyword Fallback**: If both AI services fail, the system can fall back to keyword detection.
3. **Domain Skipping**: Specific domains can be skipped from AI analysis.

## Troubleshooting

If you encounter issues with the Gemini integration:

1. **Check API Key**: Ensure your Gemini API key is correctly set in the `.env` file.
2. **Check Logs**: Review the logs in the `logs` directory for detailed error messages.
3. **Adjust Confidence Threshold**: If you're getting too many false positives or negatives, adjust the `aiConfidenceThreshold` value in `Config.gs`.
4. **Test Locally**: Use the `test:gemini` script to test the integration locally before deploying.

## Comparison with OpenAI Integration

| Feature          | Gemini           | OpenAI             |
| ---------------- | ---------------- | ------------------ |
| Privacy          | Metadata only    | Full email content |
| Confidence Score | Yes              | No (binary yes/no) |
| Multilingual     | Yes              | Yes                |
| Cost             | Generally lower  | Generally higher   |
| Integration      | Google ecosystem | Third-party        |
| Fallback Options | Multiple         | Limited            |

## Implementation Details

The Gemini integration is implemented in the following files:

- `GeminiDetection.gs`: Main implementation of the Gemini API integration
- `Config.gs`: Configuration options for the Gemini integration
- `GmailProcessing.gs`: Integration with the email processing workflow
- `test-gemini.js`: Test script for local testing

The implementation follows the same patterns as the rest of the codebase, with comprehensive error handling and logging.
