/**
 * Thread state management for Gmail Attachment Organizer
 *
 * Covers two complementary state types stored in Script Properties:
 *   - Processing state  (THREAD_PROCESSING_*): short-lived checkpoint written at the
 *     start of each thread and cleared on completion; used for stale-state recovery.
 *   - Failure state     (THREAD_FAILURE_*):    persistent record of past failures with
 *     retry counting and TTL-based expiry.
 */

// ---------------------------------------------------------------------------
// Processing state (checkpoint / stale-recovery)
// ---------------------------------------------------------------------------

/**
 * Builds script-property key for per-thread processing state.
 *
 * @param {string} threadId - Gmail thread ID
 * @returns {string} Script property key
 */
function buildThreadProcessingStateKey(threadId) {
  return `THREAD_PROCESSING_${threadId}`;
}

/**
 * Stores processing start state for a thread.
 *
 * @param {string} threadId - Gmail thread ID
 * @param {string} userEmail - Effective execution user
 */
function markThreadProcessingState(threadId, userEmail) {
  if (!threadId) return;
  const state = {
    user: userEmail || Session.getEffectiveUser().getEmail(),
    timestamp: new Date().getTime(),
  };
  PropertiesService.getScriptProperties().setProperty(
    buildThreadProcessingStateKey(threadId),
    JSON.stringify(state)
  );
}

/**
 * Clears processing state checkpoint for a thread.
 *
 * @param {string} threadId - Gmail thread ID
 */
function clearThreadProcessingState(threadId) {
  if (!threadId) return;
  PropertiesService.getScriptProperties().deleteProperty(
    buildThreadProcessingStateKey(threadId)
  );
}

/**
 * Reads processing state checkpoint for a thread.
 *
 * @param {string} threadId - Gmail thread ID
 * @returns {Object|null} Parsed state or null
 */
function getThreadProcessingState(threadId) {
  if (!threadId) return null;
  const raw = PropertiesService.getScriptProperties().getProperty(
    buildThreadProcessingStateKey(threadId)
  );
  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (e) {
    clearThreadProcessingState(threadId);
    return null;
  }
}

/**
 * Recovers stale Processing states left by interrupted executions.
 *
 * This runs before the normal processing loop and only scans a small page
 * per execution to keep runtime/cost predictable.
 *
 * @param {string} userEmail - Effective execution user
 * @param {number|null} deadlineMs - Unix timestamp (ms) soft deadline
 * @returns {Object} Recovery stats
 */
function recoverStaleProcessingThreads(userEmail, deadlineMs = null) {
  const processingLabel = getProcessingLabel();
  const searchCriteria = `label:${CONFIG.processingLabelName} -label:${CONFIG.processedLabelName}`;
  const pageSize = Math.max(1, CONFIG.staleRecoveryBatchSize || CONFIG.batchSize);
  const staleThresholdMs =
    Math.max(1, CONFIG.processingStateTtlMinutes) * 60 * 1000;
  const now = new Date().getTime();

  let checked = 0;
  let recovered = 0;
  let skippedRecent = 0;

  try {
    const threads = GmailApp.search(searchCriteria, 0, pageSize);
    if (threads.length === 0) {
      return { checked, recovered, skippedRecent };
    }

    for (let i = 0; i < threads.length; i++) {
      if (deadlineMs && new Date().getTime() >= deadlineMs) {
        logWithUser(
          "Soft deadline reached during stale recovery; resuming on next run",
          "WARNING"
        );
        break;
      }

      const thread = threads[i];
      const threadId = thread.getId();
      const state = getThreadProcessingState(threadId);
      checked++;

      let isStale = false;
      if (!state || typeof state.timestamp !== "number") {
        isStale = true;
      } else if (
        state.user &&
        userEmail &&
        state.user !== userEmail &&
        now - state.timestamp < staleThresholdMs
      ) {
        skippedRecent++;
        continue;
      } else if (now - state.timestamp >= staleThresholdMs) {
        isStale = true;
      }

      if (!isStale) {
        skippedRecent++;
        continue;
      }

      withRetry(
        () => thread.removeLabel(processingLabel),
        "removing stale processing label"
      );
      clearThreadProcessingState(threadId);
      recovered++;
    }

    if (checked > 0) {
      logWithUser(
        `Stale recovery checked ${checked} thread(s): ${recovered} recovered, ${skippedRecent} still recent`,
        "INFO"
      );
    }

    return { checked, recovered, skippedRecent };
  } catch (error) {
    logWithUser(`Error recovering stale Processing threads: ${error.message}`, "WARNING");
    return { checked, recovered, skippedRecent };
  }
}

// ---------------------------------------------------------------------------
// Failure state (retry counting + TTL expiry)
// ---------------------------------------------------------------------------

/**
 * Builds script-property key for thread failure state.
 *
 * @param {string} threadId - Gmail thread ID
 * @returns {string} Stable key for thread failure state
 */
function buildThreadFailureStateKey(threadId) {
  return `THREAD_FAILURE_${threadId}`;
}

/**
 * Reads and validates per-thread failure state.
 *
 * @param {string} threadId - Gmail thread ID
 * @returns {Object|null} Failure state or null
 */
function getThreadFailureState(threadId) {
  if (!threadId) return null;

  const key = buildThreadFailureStateKey(threadId);
  const scriptProperties = PropertiesService.getScriptProperties();
  const raw = scriptProperties.getProperty(key);
  if (!raw) return null;

  try {
    const state = JSON.parse(raw);
    const now = new Date().getTime();
    const ttlMs = Math.max(1, CONFIG.threadFailureStateTtlDays) * 24 * 60 * 60 * 1000;
    const lastFailureAt = Number(state.lastFailureAt || 0);

    if (lastFailureAt > 0 && now - lastFailureAt > ttlMs) {
      scriptProperties.deleteProperty(key);
      return null;
    }

    return state;
  } catch (e) {
    logWithUser(`getThreadFailureState: cleared corrupt state for thread ${threadId}: ${e.message}`, "DEBUG");
    scriptProperties.deleteProperty(key);
    return null;
  }
}

/**
 * Clears per-thread failure state.
 *
 * @param {string} threadId - Gmail thread ID
 */
function clearThreadFailureState(threadId) {
  if (!threadId) return;
  PropertiesService.getScriptProperties().deleteProperty(
    buildThreadFailureStateKey(threadId)
  );
}

/**
 * Classifies a processing failure as transient/permanent/too_large.
 *
 * @param {string} message - Error message
 * @param {string} codeHint - Optional code hint
 * @returns {{category: string, code: string}} Classification output
 */
function classifyProcessingFailure(message, codeHint = "") {
  const normalized = `${codeHint || ""} ${message || ""}`.toLowerCase();

  if (
    normalized.includes("too large") ||
    normalized.includes("exceeds max size") ||
    normalized.includes("file too big") ||
    normalized.includes("entity too large") ||
    normalized.includes("attachment_too_large")
  ) {
    return { category: "too_large", code: "too_large" };
  }

  if (
    normalized.includes("invalid or inaccessible folder") ||
    normalized.includes("file not found") ||
    normalized.includes("cannot find") ||
    normalized.includes("permission") ||
    normalized.includes("forbidden") ||
    normalized.includes("not authorized") ||
    normalized.includes("insufficient")
  ) {
    return { category: "permanent", code: "permission_or_config" };
  }

  if (
    normalized.includes("service invoked too many times") ||
    normalized.includes("rate limit") ||
    normalized.includes("quota") ||
    normalized.includes("timed out") ||
    normalized.includes("internal error") ||
    normalized.includes("backend error") ||
    normalized.includes("temporar") ||
    normalized.includes("try again")
  ) {
    return { category: "transient", code: "service_transient" };
  }

  return { category: "transient", code: "unknown_failure" };
}

/**
 * Registers a failure state for a thread with retry classification.
 *
 * @param {string} threadId - Gmail thread ID
 * @param {Object} failure - Failure details
 * @returns {Object|null} Updated state
 */
function registerThreadFailure(threadId, failure = {}) {
  if (!threadId) return null;

  const previous = getThreadFailureState(threadId) || {};
  const now = new Date().getTime();
  const inferred = classifyProcessingFailure(failure.message || "", failure.code || "");

  let category = failure.category || inferred.category;
  let code = failure.code || inferred.code;
  const attempts = (Number(previous.attempts) || 0) + 1;

  if (
    category !== "permanent" &&
    category !== "too_large" &&
    attempts >= Math.max(1, CONFIG.maxThreadFailureRetries)
  ) {
    category = "permanent";
    code = "max_retries_exceeded";
  }

  const state = {
    attempts: attempts,
    category: category,
    code: code,
    context: failure.context || previous.context || "unspecified",
    message: String(failure.message || "").substring(0, 1000),
    attachmentName: failure.attachmentName || null,
    attachmentSize: Number(failure.attachmentSize || 0) || null,
    user:
      failure.userEmail ||
      previous.user ||
      Session.getEffectiveUser().getEmail(),
    firstFailureAt: Number(previous.firstFailureAt) || now,
    lastFailureAt: now,
  };

  PropertiesService.getScriptProperties().setProperty(
    buildThreadFailureStateKey(threadId),
    JSON.stringify(state)
  );

  logWithUser(
    `Thread failure registered: threadId=${threadId}, category=${state.category}, code=${state.code}, attempts=${state.attempts}`,
    state.category === "permanent" ? "ERROR" : "WARNING"
  );

  return state;
}
