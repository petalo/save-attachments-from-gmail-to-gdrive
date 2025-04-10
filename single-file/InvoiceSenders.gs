/**
 * Invoice Senders for Gmail Attachment Organizer
 *
 * This file contains the list of known email senders that send invoices.
 * When CONFIG.invoiceDetection is set to "email", the script will check
 * if the sender of an email is in this list to determine if it's an invoice.
 *
 * ## Supported Formats
 *
 * The list supports various matching patterns:
 *
 * 1. Full domains (match any email from this domain):
 *    - "example.com"
 *
 * 2. Explicit wildcard domains (same as above):
 *    - "*@example.com"
 *
 * 3. Specific email addresses (exact match):
 *    - "billing@example.com"
 *
 * 4. Prefix wildcards (match emails starting with a prefix):
 *    - "invoice*@example.com" - Matches invoice123@example.com, invoice-info@example.com, etc.
 *
 * 5. Suffix wildcards (match emails ending with a suffix):
 *    - "*-noreply@example.com" - Matches billing-noreply@example.com, payment-noreply@example.com, etc.
 *
 * 6. Prefix-suffix wildcards (match emails starting with prefix and ending with suffix):
 *    - "invoice*info@example.com" - Matches invoice-user-info@example.com, invoice_additional_info@example.com, etc.
 *
 * ## How to Modify This List
 *
 * To add a new invoice sender, simply add an entry to the INVOICE_SENDERS array below.
 * The matching is case-insensitive, so "Example.com" will match "example.com".
 *
 * IMPORTANT: You only need to modify this file to update the list of invoice senders.
 * Do not modify the main Code.gs file.
 */

// List of invoice senders - MODIFY THIS LIST AS NEEDED
const INVOICE_SENDERS = [
  // Full domains (match any email from this domain)
  "pipedrivebilling.com",

  // Specific email addresses
  "billing@box.com",
  "facturacion@cuentica.com",
  "fiscal@tandemasesores.es",
  "help@paddle.com",
  "invoice@travelperk.com",
  "no_reply@am.atlassian.com",
  "payments-noreply@google.com",

  // Special cases with domain patterns
  "@cm.order.email.ikea.com", // Any email from this subdomain

  // Pattern matching with wildcards
  "invoice+statements+*@stripe.com", // Matches any email that starts with "invoice+statements+"

  // Add more entries as needed
];
