/**
 * Attachment filter functions for Gmail Attachment Organizer
 */

/**
 * Checks if a file should be skipped based on filter rules
 *
 * Why it's necessary:
 * - Avoids storing unnecessary files such as:
 *   1. Small images that are usually email signatures or icons
 *   2. Specific file types like calendar invitations (.ics)
 *   3. Embedded/inline images from email body
 * - Improves storage efficiency and keeps folders organized
 * - Reduces processing time by filtering irrelevant files early
 *
 * @param {String} fileName - The name of the file
 * @param {Number} fileSize - The size of the file in bytes
 * @param {GmailAttachment} attachment - The Gmail attachment object (optional)
 * @returns {Boolean} - True if the file should be skipped, false otherwise
 */
function shouldSkipFile(fileName, fileSize, attachment = null) {
  // Check mime type if attachment is provided
  if (attachment) {
    const mimeType = attachment.getContentType();

    // Always keep files with these MIME types (like documents and archives)
    if (CONFIG.attachmentTypesWhitelist.includes(mimeType)) {
      logWithUser(
        `Keeping file with important MIME type ${mimeType}: ${fileName}`,
        "INFO"
      );
      return false; // Don't skip these types
    }

    // Check if it's an inline image (as opposed to a real attachment)
    try {
      // We can check some internal properties to try to determine if this is an inline image
      const contentDisposition = attachment.getContentDisposition
        ? attachment.getContentDisposition()
        : null;

      // If it's explicitly marked as "inline" and it's an image type
      if (
        contentDisposition &&
        contentDisposition.toLowerCase().includes("inline") &&
        mimeType.startsWith("image/")
      ) {
        logWithUser(
          `Skipping inline image with Content-Disposition "${contentDisposition}": ${fileName}`,
          "INFO"
        );
        return true;
      }
    } catch (e) {
      // Can't determine content disposition, continue with other checks
    }
  }

  // Check if the filename appears to be a Gmail embedded image URL or other email provider URLs
  if (
    (fileName.includes("mail.google.com/mail") &&
      (fileName.includes("view=fimg") || fileName.includes("disp=emb"))) ||
    // Outlook and Exchange embedded image URLs
    (fileName.includes("outlook.office") && fileName.includes("attachment")) ||
    // Yahoo Mail embedded images
    (fileName.includes("yimg.com") && fileName.includes("mail")) ||
    // General URL patterns often seen in HTML emails
    (fileName.startsWith("https://") &&
      (fileName.includes("?cid=") ||
        fileName.includes("&cid=") ||
        fileName.includes("&disp=") ||
        fileName.includes("&view=")))
  ) {
    logWithUser(`Skipping embedded email image URL: ${fileName}`, "INFO");
    return true;
  }

  // Common image and icon names used in email services
  const commonEmbeddedNames = [
    "box_logo",
    "logo",
    "icon",
    "banner",
    "header",
    "footer",
    "signature",
    "badge",
    "avatar",
    "profile",
    "divider",
    "spacer",
    "separator",
    "pixel",
    "background",
    "bg",
    "bullet",
    "social",
    "facebook",
    "twitter",
    "linkedin",
    "instagram",
  ];

  // Check for common embedded image names without extensions
  for (const name of commonEmbeddedNames) {
    if (
      fileName.toLowerCase() === name.toLowerCase() ||
      fileName.toLowerCase().includes(name.toLowerCase() + "_")
    ) {
      logWithUser(`Skipping common embedded element: ${fileName}`, "INFO");
      return true;
    }
  }

  // Detect common patterns for embedded/inline images in email body
  const embeddedImagePatterns = [
    /^image\d+\.(png|jpg|gif|jpeg)$/i, // Common Gmail embedded image pattern (image001.png)
    /^inline-/i, // Inline images often start with "inline-"
    /^Outlook-/i, // Outlook embedded images prefix
    /^emb_(image|embed)\d+/i, // Common embedded image pattern
    /^[a-f0-9]{8}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{4}-[a-f0-9]{12}(\.(png|jpg|gif|jpeg))?$/i, // UUID-style names often used for embedded images
    /^part_\d+\.\d+\.\d+$/i, // Gmail inline image format
    /^att\d+\.\d+$/i, // Another common email attachment format
  ];

  // Check if the filename matches any of the embedded image patterns
  for (const pattern of embeddedImagePatterns) {
    if (pattern.test(fileName)) {
      logWithUser(`Skipping embedded image: ${fileName}`, "INFO");
      return true;
    }
  }

  // Extract file extension - handle files without extensions
  const lastDotIndex = fileName.lastIndexOf(".");
  const fileExtension =
    lastDotIndex !== -1 ? fileName.substring(lastDotIndex).toLowerCase() : "";

  // Common document extensions we want to keep
  const keepExtensions = [
    ".pdf",
    ".doc",
    ".docx",
    ".xls",
    ".xlsx",
    ".ppt",
    ".pptx",
    ".zip",
    ".rar",
    ".7z",
    ".csv",
    ".txt",
  ];

  // Always keep documents regardless of size
  if (keepExtensions.includes(fileExtension)) {
    logWithUser(`Keeping document file: ${fileName}`, "INFO");
    return false; // Don't skip these types
  }

  // If no extension and likely an embedded image/content (often binary content)
  // Most email service logos, icons and embedded HTML images are under 50KB
  if (fileExtension === "" && fileSize < 50 * 1024) {
    logWithUser(
      `Skipping potential embedded content without extension: ${fileName} (${Math.round(
        fileSize / 1024
      )}KB)`,
      "INFO"
    );
    return true;
  }

  // Check if we should skip files by extension (like calendar invitations)
  if (CONFIG.skipFileTypes && CONFIG.skipFileTypes.includes(fileExtension)) {
    logWithUser(`Skipping file type: ${fileName} (${fileExtension})`, "INFO");
    return true;
  }

  // Check if we should skip small images
  if (CONFIG.skipSmallImages) {
    // If it's an image extension in our filter list and smaller than the threshold
    if (
      CONFIG.smallImageExtensions.includes(fileExtension) &&
      fileSize <= CONFIG.smallImageMaxSize
    ) {
      logWithUser(
        `Skipping small image file: ${fileName} (${Math.round(
          fileSize / 1024
        )}KB)`,
        "INFO"
      );
      return true;
    }
  }

  return false;
}
