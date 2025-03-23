/**
 * Historical Pattern Analysis for Gmail Attachment Organizer
 *
 * This module provides functions to analyze historical patterns in emails
 * that have been manually labeled as invoices, to improve invoice detection.
 */

/**
 * Gets historical invoice patterns from emails with the same sender
 * that have been manually labeled with the configured label
 *
 * @param {string} senderEmail - Email address of the sender
 * @returns {Object|null} Patterns detected in historical invoices, or null if disabled
 */
function getHistoricalInvoicePatterns(senderEmail) {
  // Only proceed if feature is enabled and label is configured
  if (!CONFIG.useHistoricalPatterns || !CONFIG.manuallyLabeledInvoicesLabel) {
    return null;
  }

  try {
    // Use the full sender email for more accurate internal Gmail search
    const searchQuery = `from:(${senderEmail}) label:${CONFIG.manuallyLabeledInvoicesLabel}`;
    const threads = GmailApp.search(searchQuery, 0, CONFIG.maxHistoricalEmails);

    logWithUser(
      `Found ${threads.length} historical emails from ${senderEmail} with label "${CONFIG.manuallyLabeledInvoicesLabel}"`,
      "INFO"
    );

    if (threads.length === 0) {
      return null;
    }

    // Collect metadata from these historical invoices
    const subjects = [];
    const dates = [];

    for (const thread of threads) {
      const messages = thread.getMessages();
      if (messages.length > 0) {
        const message = messages[0]; // Get the first message in each thread
        subjects.push(message.getSubject());
        dates.push(message.getDate());
      }
    }

    // Extract patterns
    const patterns = {
      count: threads.length,
      subjectPatterns: extractSubjectPatterns(subjects),
      datePatterns: extractDatePatterns(dates),
      // For internal use only (not sent to AI)
      rawSubjects: subjects,
      rawDates: dates.map((d) => d.toISOString()),
    };

    return patterns;
  } catch (error) {
    logWithUser(
      `Error getting historical invoice patterns: ${error.message}`,
      "ERROR"
    );
    return null;
  }
}

/**
 * Extracts patterns from email subjects
 *
 * @param {string[]} subjects - Array of email subjects
 * @returns {Object} Patterns found in subjects
 */
function extractSubjectPatterns(subjects) {
  if (!subjects || subjects.length === 0) {
    return {};
  }

  try {
    // Find common prefixes
    let commonPrefix = subjects[0];
    for (let i = 1; i < subjects.length; i++) {
      let j = 0;
      while (
        j < commonPrefix.length &&
        j < subjects[i].length &&
        commonPrefix.charAt(j) === subjects[i].charAt(j)
      ) {
        j++;
      }
      commonPrefix = commonPrefix.substring(0, j);
    }

    // Find common suffixes
    let commonSuffix = subjects[0];
    for (let i = 1; i < subjects.length; i++) {
      let j = 0;
      while (
        j < commonSuffix.length &&
        j < subjects[i].length &&
        commonSuffix.charAt(commonSuffix.length - 1 - j) ===
          subjects[i].charAt(subjects[i].length - 1 - j)
      ) {
        j++;
      }
      commonSuffix = commonSuffix.substring(commonSuffix.length - j);
    }

    // Check for invoice keywords
    const containsInvoiceTerms = subjects.some((subject) =>
      CONFIG.invoiceKeywords.some((keyword) =>
        subject.toLowerCase().includes(keyword.toLowerCase())
      )
    );

    // Check for numeric patterns (like invoice numbers)
    const numericPatterns = [];
    const numericRegex = /\d+/g;

    subjects.forEach((subject) => {
      const matches = subject.match(numericRegex);
      if (matches) {
        numericPatterns.push(...matches);
      }
    });

    return {
      commonPrefix: commonPrefix.length > 3 ? commonPrefix.trim() : "",
      commonSuffix: commonSuffix.length > 3 ? commonSuffix.trim() : "",
      containsInvoiceTerms: containsInvoiceTerms,
      hasNumericPatterns: numericPatterns.length > 0,
    };
  } catch (error) {
    logWithUser(`Error extracting subject patterns: ${error.message}`, "ERROR");
    return {};
  }
}

/**
 * Extracts patterns from email dates
 *
 * @param {Date[]} dates - Array of email dates
 * @returns {Object} Patterns found in dates
 */
function extractDatePatterns(dates) {
  if (!dates || dates.length < 2) {
    return { frequency: "unknown" };
  }

  try {
    // Sort dates chronologically
    dates.sort((a, b) => a - b);

    // Calculate intervals between dates in days
    const intervals = [];
    for (let i = 1; i < dates.length; i++) {
      const diffTime = Math.abs(dates[i] - dates[i - 1]);
      const diffDays = Math.ceil(diffTime / (1000 * 60 * 60 * 24));
      intervals.push(diffDays);
    }

    // Analyze frequency
    const avgInterval =
      intervals.reduce((sum, days) => sum + days, 0) / intervals.length;

    let frequency = "irregular";

    if (avgInterval >= 25 && avgInterval <= 35) {
      frequency = "monthly";
    } else if (avgInterval >= 85 && avgInterval <= 95) {
      frequency = "quarterly";
    } else if (avgInterval >= 175 && avgInterval <= 190) {
      frequency = "biannual";
    } else if (avgInterval >= 350 && avgInterval <= 380) {
      frequency = "annual";
    } else if (avgInterval >= 13 && avgInterval <= 16) {
      frequency = "biweekly";
    } else if (avgInterval >= 6 && avgInterval <= 8) {
      frequency = "weekly";
    }

    // Check if dates fall on same day of month
    const daysOfMonth = dates.map((d) => d.getDate());
    const uniqueDaysOfMonth = [...new Set(daysOfMonth)];
    const sameDayOfMonth = uniqueDaysOfMonth.length === 1;

    return {
      frequency: frequency,
      averageIntervalDays: Math.round(avgInterval),
      sameDayOfMonth: sameDayOfMonth,
      dayOfMonth: sameDayOfMonth ? daysOfMonth[0] : null,
    };
  } catch (error) {
    logWithUser(`Error extracting date patterns: ${error.message}`, "ERROR");
    return { frequency: "unknown" };
  }
}

/**
 * Formats historical patterns into a human-readable description
 * for inclusion in AI prompts
 *
 * @param {Object} patterns - The patterns object from getHistoricalInvoicePatterns
 * @returns {string} Human-readable description of patterns
 */
function formatHistoricalPatternsForPrompt(patterns) {
  if (!patterns || patterns.count === 0) {
    return "";
  }

  try {
    let description = `\nHistorical context: The sender domain has ${patterns.count} previous emails that were manually labeled as invoices.\n\n`;

    // Subject patterns
    if (patterns.subjectPatterns) {
      description += "Subject patterns:\n";

      if (patterns.subjectPatterns.commonPrefix) {
        description += `- Subjects often start with: "${patterns.subjectPatterns.commonPrefix}"\n`;
      }

      if (patterns.subjectPatterns.commonSuffix) {
        description += `- Subjects often end with: "${patterns.subjectPatterns.commonSuffix}"\n`;
      }

      if (patterns.subjectPatterns.containsInvoiceTerms) {
        description += "- Subjects typically contain invoice-related terms\n";
      }

      if (patterns.subjectPatterns.hasNumericPatterns) {
        description +=
          "- Subjects often contain numeric patterns (like invoice numbers)\n";
      }
    }

    // Date patterns
    if (patterns.datePatterns) {
      description += "\nTiming patterns:\n";

      if (
        patterns.datePatterns.frequency !== "unknown" &&
        patterns.datePatterns.frequency !== "irregular"
      ) {
        description += `- Emails are typically sent ${patterns.datePatterns.frequency}\n`;
      }

      if (patterns.datePatterns.sameDayOfMonth) {
        description += `- Emails are typically sent on day ${patterns.datePatterns.dayOfMonth} of the month\n`;
      }

      if (patterns.datePatterns.frequency === "irregular") {
        description += `- Average interval between emails: ${patterns.datePatterns.averageIntervalDays} days\n`;
      }
    }

    return description;
  } catch (error) {
    logWithUser(
      `Error formatting historical patterns: ${error.message}`,
      "ERROR"
    );
    return "";
  }
}

// Make functions available to other modules
var HistoricalPatterns = {
  getHistoricalInvoicePatterns: getHistoricalInvoicePatterns,
  formatHistoricalPatternsForPrompt: formatHistoricalPatternsForPrompt,
};
