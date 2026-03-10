/**
 * Gmail label management for Gmail Attachment Organizer
 */

/**
 * Gets or creates a Gmail label by name
 *
 * @param {string} labelName - Label name to fetch/create
 * @returns {GmailLabel} Gmail label instance
 */
function getOrCreateLabel(labelName) {
  let label = GmailApp.getUserLabelByName(labelName);

  if (!label) {
    label = GmailApp.createLabel(labelName);
    logWithUser(`Created new Gmail label: ${labelName}`);
  } else {
    logWithUser(`Using existing Gmail label: ${labelName}`);
  }

  return label;
}

/**
 * Gets or creates the processed label
 *
 * @returns {GmailLabel} The Gmail label used to mark processed threads
 */
function getProcessedLabel() {
  return getOrCreateLabel(CONFIG.processedLabelName);
}

/**
 * Gets or creates the processing label
 *
 * @returns {GmailLabel} Label used while a thread is being processed
 */
function getProcessingLabel() {
  return getOrCreateLabel(CONFIG.processingLabelName);
}

/**
 * Gets or creates the error label
 *
 * @returns {GmailLabel} Label used when processing fails
 */
function getErrorLabel() {
  return getOrCreateLabel(CONFIG.errorLabelName);
}

/**
 * Gets or creates the permanent error label
 *
 * @returns {GmailLabel} Label used for non-retriable failures
 */
function getPermanentErrorLabel() {
  return getOrCreateLabel(CONFIG.permanentErrorLabelName);
}

/**
 * Gets or creates the too-large label
 *
 * @returns {GmailLabel} Label used when attachments exceed max size
 */
function getTooLargeLabel() {
  return getOrCreateLabel(CONFIG.tooLargeLabelName);
}
