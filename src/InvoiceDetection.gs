/**
 * Invoice detection logic for Gmail Attachment Organizer
 *
 * This file contains the logic for matching email senders against the invoice senders list.
 * It implements pattern matching with wildcards and domain-based matching.
 */

/**
 * Checks if a sender is in the invoice senders list
 *
 * @param {string} sender - The email address of the sender
 * @returns {boolean} True if the sender is in the list
 */
function isInvoiceSender(sender) {
  try {
    if (!sender) return false;

    // Extract email from "Display Name <email@domain.com>" format if needed
    const angleMatch = sender.match(/<([^>]+)>/);
    const cleanSender = angleMatch ? angleMatch[1] : sender;

    // Normalize the email to lowercase
    const normalizedSender = cleanSender.toLowerCase();

    // Extract the domain
    const domain = extractDomain(normalizedSender);

    // Extract the username (part before @)
    const atIndex = normalizedSender.indexOf("@");
    const username = atIndex > 0 ? normalizedSender.substring(0, atIndex) : "";

    // Check for matches
    for (const entry of INVOICE_SENDERS) {
      const normalizedEntry = entry.toLowerCase();

      // Case 1: Exact email match
      if (normalizedSender === normalizedEntry) {
        logWithUser(`Exact invoice sender match: ${sender}`, "INFO");
        return true;
      }

      // Case 2: Domain-only match (entry is just a domain without @)
      if (!normalizedEntry.includes("@") && domain === normalizedEntry) {
        logWithUser(`Domain match for invoice sender: ${domain}`, "INFO");
        return true;
      }

      // Case 3: Full wildcard match (*@domain.com)
      if (
        normalizedEntry.startsWith("*@") &&
        domain === normalizedEntry.substring(2)
      ) {
        logWithUser(
          `Wildcard domain match for invoice sender: ${domain}`,
          "INFO"
        );
        return true;
      }

      // Case 4: Pattern matching with wildcards
      if (normalizedEntry.includes("*") && normalizedEntry.includes("@")) {
        // Extract pattern parts
        const entryAtIndex = normalizedEntry.indexOf("@");
        const entryUsername = normalizedEntry.substring(0, entryAtIndex);
        const entryDomain = normalizedEntry.substring(entryAtIndex + 1);

        // Only proceed if domains match
        if (domain === entryDomain) {
          // Handle prefix*@domain.com
          if (entryUsername.endsWith("*") && !entryUsername.startsWith("*")) {
            const prefix = entryUsername.substring(0, entryUsername.length - 1);
            if (username.startsWith(prefix)) {
              logWithUser(
                `Prefix wildcard match for invoice sender: ${sender}`,
                "INFO"
              );
              return true;
            }
          }
          // Handle *suffix@domain.com
          else if (
            entryUsername.startsWith("*") &&
            !entryUsername.endsWith("*")
          ) {
            const suffix = entryUsername.substring(1);
            if (username.endsWith(suffix)) {
              logWithUser(
                `Suffix wildcard match for invoice sender: ${sender}`,
                "INFO"
              );
              return true;
            }
          }
          // Handle prefix*suffix@domain.com
          else if (entryUsername.includes("*")) {
            const parts = entryUsername.split("*");
            const prefix = parts[0];
            const suffix = parts[1];

            if (username.startsWith(prefix) && username.endsWith(suffix)) {
              logWithUser(
                `Prefix-suffix wildcard match for invoice sender: ${sender}`,
                "INFO"
              );
              return true;
            }
          }
        }
      }
    }

    return false;
  } catch (error) {
    logWithUser(`Error checking invoice sender: ${error.message}`, "ERROR");
    return false;
  }
}

// Export the function for use in other modules
var InvoiceDetection = {
  isInvoiceSender: isInvoiceSender,
};
