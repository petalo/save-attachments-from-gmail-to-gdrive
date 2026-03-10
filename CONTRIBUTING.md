# Contributing to Gmail Attachment Organizer

## Architecture

The codebase follows a modular structure where each `.gs` file has a single, clear responsibility. Dependencies flow in one direction:

```
Main.gs
├── Config.gs              (no dependencies — read by everyone)
├── GmailProcessing.gs
│   ├── AttachmentProcessing.gs
│   │   ├── AttachmentFilters.gs
│   │   └── FolderManagement.gs
│   └── Utils.gs
└── UserManagement.gs
    └── Utils.gs
```

`Utils.gs` and `Config.gs` are leaf dependencies — they do not import anything else.

### Module Responsibilities

| File | Role |
|------|------|
| `Main.gs` | Entry point, execution lock, trigger creation |
| `Config.gs` | Centralized `CONFIG` object — all settings live here |
| `GmailProcessing.gs` | Thread/message search and processing loop |
| `AttachmentProcessing.gs` | Saves attachments to Drive, deduplication |
| `AttachmentFilters.gs` | Decides which attachments to keep or skip |
| `FolderManagement.gs` | Gets or creates domain-based Drive folders |
| `UserManagement.gs` | User onboarding, permission verification |
| `Utils.gs` | `withRetry()`, `logWithUser()`, `extractDomain()` |
| `Debug.gs` | Manual debug helpers, not used in production flow |

---

## Design Patterns

### Module Pattern

Each file encapsulates its own state and exposes only public functions. There are no explicit imports — Google Apps Script loads all `.gs` files into a single global scope, so functions are available across files without `require`.

### CONFIG Object

All configurable values live in the `CONFIG` object in `Config.gs`. Logic files must never hardcode values; they always read from `CONFIG`.

```javascript
// Good
if (fileSize > CONFIG.maxFileSize) { ... }

// Bad
if (fileSize > 26214400) { ... }
```

### Retry with Exponential Backoff

All Drive and Gmail API calls go through `withRetry()` in `Utils.gs`. This handles transient quota errors automatically.

```javascript
const folder = withRetry(
  () => mainFolder.getFoldersByName(domain),
  "getting domain folder"
);
```

### Effective User Resolution

Each trigger execution processes only the mailbox of the currently authenticated user — no queue rotation, no impersonation. This is intentional for simplicity and OAuth scope correctness.

```javascript
const currentUser = Session.getEffectiveUser().getEmail();
processUserEmails(currentUser, true, deadlineMs);
```

---

## Technical Constraints

These are hard constraints imposed by Google Apps Script — keep them in mind when adding features:

| Constraint | Detail |
|------------|--------|
| **6-minute execution limit** | The script uses `executionSoftLimitMs` (default 4.5 min) to stop early and resume on the next trigger run |
| **Thread-level Gmail labels** | Labels attach to threads, not individual messages. New messages in a labeled thread won't be re-processed automatically |
| **No file creation date** | Drive API can't set arbitrary creation dates. Email dates are stored in the file description instead |
| **No npm at runtime** | Only Google Apps Script built-in services are available; `UrlFetchApp`, `DriveApp`, `GmailApp`, etc. |
| **Limited ES6+** | Arrow functions and template literals work, but avoid advanced features like optional chaining (`?.`) — GAS runtime support varies |
| **Script Properties size** | `PropertiesService` values are limited in size; avoid storing large objects |
| **25MB attachment limit** | This is a Gmail API ceiling, not configurable in code |

---

## Concurrency Control

Multiple users can each have their own time-based trigger running in parallel safely. The mechanism:

1. **Per-user execution lock** — key `EXECUTION_LOCK_{email}` in Script Properties via `LockService.getUserLock()`. If a lock is already held, the execution exits immediately.
2. **Soft deadline** — `executionSoftLimitMs` ensures the script never hits the 6-minute hard wall.
3. **Stale state recovery** — threads stuck in `GDrive_Processing` state (e.g. from a crashed execution) are automatically recovered after `processingStateTtlMinutes`.

---

## Code Conventions

- **Function names**: `camelCase`
- **Constants**: `UPPER_SNAKE_CASE`
- **JSDoc**: every function has a `@param` / `@returns` block
- **Error handling**: always `try/catch` + `logWithUser(message, "ERROR")` — never a silent catch
- **Logging levels**: `DEBUG` for verbose tracing, `INFO` for key events, `WARNING` for recoverable issues, `ERROR` for failures

```javascript
// Correct error handling pattern
try {
  const file = DriveApp.getFileById(fileId);
  return file;
} catch (e) {
  logWithUser(`Could not access file ${fileId}: ${e.message}`, "ERROR");
  return null;
}
```

---

## Extension Points

### Adding a new attachment filter

Add a condition to `shouldSkipFile()` in `src/AttachmentFilters.gs`. The function returns `true` to skip a file, `false` to keep it.

### Changing folder organization

Modify `getDomainFolder()` in `src/FolderManagement.gs` to implement a different folder naming or nesting strategy. The function receives the full sender string and the main Drive folder.

### Storing extra metadata per file

`AttachmentProcessing.gs` calls `savedFile.setDescription(...)` with email date info. Additional fields can be appended to that string.

---

## Development Workflow

### Prerequisites

- Node.js (for build scripts)
- [clasp](https://github.com/google/clasp) for deployment

### Setup

```bash
npm install
clasp login
```

### Environment files

Create `.env.prod` (required) and optionally `.env.test` based on `.env.example`. These are not committed to version control.

Required variables:

- `FOLDER_ID` — Google Drive folder ID for saved attachments
- `SCRIPT_ID` — Google Apps Script project ID
- `PROCESSED_LABEL_NAME` — Gmail label name (allows different labels per environment)

### Build and deploy

```bash
npm run build:prod    # Generates single-file/Code.gs with env vars injected
npm run deploy:prod   # Pushes to Google Apps Script via clasp
```

The build script (`combine-files.js`) replaces placeholders like `__FOLDER_ID__` with values from the `.env.*` file.

### Important: avoid GCP project association

Do **not** associate the Apps Script project with a GCP (Google Cloud Platform) project. Doing so introduces OAuth permission issues that are hard to diagnose. Create the script directly at [script.google.com](https://script.google.com) without linking to a GCP project.

---

## Known Issues

| Issue | Workaround |
|-------|-----------|
| New messages in processed threads are skipped | Remove the `GDrive_Processed` label from the thread to force reprocessing |
| Rare duplicate Drive folders under high concurrency | Lock logic (`LockService`) reduces but doesn't fully eliminate this — Drive's folder API is eventually consistent |
| Attachments over 25MB silently skipped | Gmail API limit; thread gets `GDrive_TooLarge` label |
