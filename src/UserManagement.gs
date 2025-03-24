/**
 * User management functions for Gmail Attachment Organizer
 */

/**
 * Verify if a user has granted all required permissions
 *
 * @param {string} userEmail - The user's email address to verify. If not provided, uses current user.
 * @returns {boolean} True if the user has all required permissions
 *
 * The function follows this flow:
 * 1. Checks if the user is the current user or another user
 * 2. Looks for cached permission results to avoid redundant checks
 * 3. For the current user:
 *    - Attempts to access Gmail and Drive services, which triggers permission prompts
 *    - Returns true only if both services are accessible
 * 4. For other users:
 *    - Checks if they have already granted Gmail and Drive permissions
 *    - Cannot trigger permission prompts for other users
 * 5. Caches the result for future checks within the same script execution
 *
 * This verification is crucial for multi-user scripts to ensure each user
 * has granted the necessary permissions before attempting to process their data.
 */
function verifyUserPermissions(userEmail) {
  // If userEmail is not provided, use the current user's email
  if (!userEmail) {
    userEmail = Session.getEffectiveUser().getEmail();
    logWithUser(
      `No user email provided, using current user: ${userEmail}`,
      "INFO"
    );
  }
  const currentUser = Session.getEffectiveUser().getEmail();
  const isCurrentUser = currentUser === userEmail;

  // Use script properties to track users we've already checked
  const scriptProperties = PropertiesService.getScriptProperties();
  const checkedUsersKey = "CHECKED_USERS_CACHE";

  try {
    // Try to get the cache of checked users
    let checkedUsersCache = scriptProperties.getProperty(checkedUsersKey);
    let checkedUsers = checkedUsersCache ? JSON.parse(checkedUsersCache) : {};

    // If we've checked this user recently (within this script run), use cached result
    if (checkedUsers[userEmail] !== undefined) {
      return checkedUsers[userEmail];
    }

    let result = false;

    if (isCurrentUser) {
      // For the current user, we'll attempt to force permission requests if needed
      try {
        // Try to access Gmail - this should trigger permission request if needed
        const labels = GmailApp.getUserLabels();

        // Try to access Drive - this should trigger permission request if needed
        const rootFolder = DriveApp.getRootFolder();

        logWithUser(`User ${userEmail} has all required permissions`, "INFO");
        result = true;
      } catch (e) {
        // If we get here with the current user, it means permissions weren't granted even after prompting
        logWithUser(
          `User ${userEmail} denied or has not granted required permissions: ${e.message}`,
          "WARNING"
        );
        result = false;
      }
    } else {
      // For other users, we can only check if they already have permissions
      // Try to access user's Gmail
      let hasGmailAccess = false;
      let hasDriveAccess = false;

      try {
        const userGmail = GmailApp.getUserLabelByName("INBOX");
        hasGmailAccess = true;
      } catch (e) {
        logWithUser(
          `User ${userEmail} has not granted Gmail permissions`,
          "WARNING"
        );
      }

      try {
        const userDrive = DriveApp.getRootFolder();
        hasDriveAccess = true;
      } catch (e) {
        logWithUser(
          `User ${userEmail} has not granted Drive permissions`,
          "WARNING"
        );
      }

      result = hasGmailAccess && hasDriveAccess;

      if (result) {
        logWithUser(`User ${userEmail} has all required permissions`, "INFO");
      } else {
        logWithUser(
          `User ${userEmail} is missing some required permissions`,
          "WARNING"
        );
      }
    }

    // Cache the result
    checkedUsers[userEmail] = result;
    scriptProperties.setProperty(checkedUsersKey, JSON.stringify(checkedUsers));

    // Log the final result when called directly
    if (!userEmail || userEmail === currentUser) {
      logWithUser(
        `Permission verification result for current user: ${
          result ? "GRANTED" : "DENIED"
        }`,
        "INFO"
      );
    }

    return result;
  } catch (e) {
    logWithUser(
      `Error verifying permissions for ${userEmail}: ${e.message}`,
      "ERROR"
    );
    return false;
  }
}

/**
 * Explicitly request all permissions needed by the script
 * This function should be run manually by each user to grant permissions
 *
 * @returns {boolean} True if permissions were granted successfully
 *
 * The function follows this flow:
 * 1. Attempts to access Gmail, which triggers the permission prompt
 * 2. Attempts to access Drive, which triggers another permission prompt
 * 3. If the main folder ID is configured, verifies access to that folder
 * 4. Verifies all permissions using verifyUserPermissions()
 * 5. If all permissions are granted, adds the user to the authorized users list
 *
 * This is typically the first function a new user should run before using the script,
 * as it ensures all necessary permissions are granted and the user is properly registered.
 */
function requestPermissions() {
  try {
    const userEmail = Session.getEffectiveUser().getEmail();
    logWithUser("Requesting permissions for Gmail and Drive...", "INFO");

    // Try to access Gmail - this will trigger the permission prompt
    try {
      const gmailLabels = GmailApp.getUserLabels();
      logWithUser("Gmail permissions granted", "INFO");
    } catch (gmailError) {
      logWithUser(`Error accessing Gmail: ${gmailError.message}`, "ERROR");
      return false;
    }

    // Try to access Drive - this will trigger the permission prompt
    try {
      const rootFolder = DriveApp.getRootFolder();
      logWithUser("Drive permissions granted", "INFO");
    } catch (driveError) {
      logWithUser(`Error accessing Drive: ${driveError.message}`, "ERROR");
      return false;
    }

    // Verify that we can access the main folder
    if (CONFIG.mainFolderId !== "YOUR_SHARED_FOLDER_ID") {
      try {
        const mainFolder = DriveApp.getFolderById(CONFIG.mainFolderId);
        logWithUser(
          `Successfully accessed main folder: ${mainFolder.getName()}`,
          "INFO"
        );
      } catch (folderError) {
        // This is not a critical error, just log it
        logWithUser(
          `Error accessing main folder: ${folderError.message}. Please check the folder ID.`,
          "ERROR"
        );
        // Continue with the rest of the function
      }
    } else {
      logWithUser(
        "Please configure the mainFolderId in the CONFIG object before continuing.",
        "WARNING"
      );
    }

    // Let's add the current user to the manual list if they have all permissions
    try {
      const hasPermissions = verifyUserPermissions(userEmail);
      if (hasPermissions) {
        addUserToList(userEmail);
        logWithUser(
          "All required permissions have been granted successfully!",
          "INFO"
        );
        logWithUser(
          "You have been added to the users list and the script is ready to process your emails.",
          "INFO"
        );
      } else {
        logWithUser(
          "Some permissions appear to be missing. Please run this function again.",
          "WARNING"
        );
      }
      return hasPermissions;
    } catch (permissionError) {
      logWithUser(
        `Error verifying permissions: ${permissionError.message}`,
        "ERROR"
      );
      return false;
    }
  } catch (e) {
    logWithUser(`Error requesting permissions: ${e.message}`, "ERROR");
    return false;
  }
}

/**
 * Get all users who have access to the script and verify their permissions
 *
 * @returns {string[]} Array of user email addresses
 */
function getAuthorizedUsers() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let users = [];

  try {
    const usersProperty = scriptProperties.getProperty("REGISTERED_USERS");
    if (usersProperty) {
      users = JSON.parse(usersProperty);
    }
  } catch (e) {
    logWithUser(`Error parsing registered users: ${e.message}`, "ERROR");
  }

  const currentUser = Session.getEffectiveUser().getEmail();

  if (users.length === 0) {
    logWithUser(
      `No registered users found. Using current user: ${currentUser}`,
      "INFO"
    );
    users = [currentUser];
  } else {
    logWithUser(`Found ${users.length} registered users`, "INFO");
  }

  return users;
}

/**
 * Add a user to the authorized users list
 *
 * @param {string} email - Email address to add
 * @returns {boolean} True if the user was added successfully
 */
function addUserToList(email) {
  if (!email || !email.includes("@")) {
    logWithUser("Invalid email format", "ERROR");
    return false;
  }

  const scriptProperties = PropertiesService.getScriptProperties();
  let users = [];

  try {
    const usersProperty = scriptProperties.getProperty("REGISTERED_USERS");
    if (usersProperty) {
      users = JSON.parse(usersProperty);
    }

    if (!users.includes(email)) {
      users.push(email);
      scriptProperties.setProperty("REGISTERED_USERS", JSON.stringify(users));
      logWithUser(`User ${email} added to registered users list`, "INFO");
    } else {
      logWithUser(
        `User ${email} is already in the registered users list`,
        "INFO"
      );
    }

    return true;
  } catch (e) {
    logWithUser(`Error adding user ${email}: ${e.message}`, "ERROR");
    return false;
  }
}

/**
 * Remove a user from the authorized users list
 *
 * @param {string} email - Email address to remove
 * @returns {boolean} True if the user was removed successfully
 */
function removeUserFromList(email) {
  if (!email || !email.includes("@")) {
    logWithUser("Invalid email format", "ERROR");
    return false;
  }

  const scriptProperties = PropertiesService.getScriptProperties();
  let users = [];

  try {
    const usersProperty = scriptProperties.getProperty("REGISTERED_USERS");
    if (usersProperty) {
      users = JSON.parse(usersProperty);
    }

    const index = users.indexOf(email);
    if (index > -1) {
      users.splice(index, 1);
      scriptProperties.setProperty("REGISTERED_USERS", JSON.stringify(users));
      logWithUser(`User ${email} removed from registered users list`, "INFO");
    } else {
      logWithUser(`User ${email} not found in registered users list`, "INFO");
    }

    return true;
  } catch (e) {
    logWithUser(`Error removing user ${email}: ${e.message}`, "ERROR");
    return false;
  }
}

/**
 * List all registered users
 *
 * @returns {string[]} Array of user email addresses
 */
function listUsers() {
  const users = getAuthorizedUsers();

  if (users.length === 0) {
    logWithUser("No registered users found", "INFO");
  } else {
    logWithUser(`Registered users (${users.length}):`, "INFO");
    users.forEach((email, index) => {
      logWithUser(`${index + 1}. ${email}`, "INFO");
    });
  }

  return users;
}

/**
 * Get the next user in the queue
 *
 * @returns {string} Email address of the next user to process
 */
function getNextUserInQueue() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const queueKey = "USER_QUEUE";
  const lastProcessedKey = "LAST_PROCESSED_USER";

  try {
    // Get the queue of users
    let queue = JSON.parse(scriptProperties.getProperty(queueKey) || "[]");

    // If queue is empty, refresh it with current authorized users
    if (queue.length === 0) {
      queue = getAuthorizedUsers();
      scriptProperties.setProperty(queueKey, JSON.stringify(queue));
    }

    // Get the last processed user
    const lastProcessed = scriptProperties.getProperty(lastProcessedKey);

    // Find the next user in the queue
    let nextUser;
    if (lastProcessed) {
      const lastIndex = queue.indexOf(lastProcessed);
      nextUser = queue[(lastIndex + 1) % queue.length];
    } else {
      nextUser = queue[0];
    }

    // Update the last processed user
    scriptProperties.setProperty(lastProcessedKey, nextUser);

    logWithUser(`Next user in queue: ${nextUser}`, "INFO");
    return nextUser;
  } catch (e) {
    logWithUser(`Error getting next user in queue: ${e.message}`, "ERROR");
    return null;
  }
}

/**
 * Function to verify all users' permissions
 */
function verifyAllUsersPermissions() {
  // Clear the checked users cache to ensure fresh checks
  const scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.deleteProperty("CHECKED_USERS_CACHE");

  // First check if the current user has permissions
  const currentUser = Session.getEffectiveUser().getEmail();
  const currentUserHasPermissions = verifyUserPermissions(currentUser);

  if (!currentUserHasPermissions) {
    logWithUser(
      "You need to grant necessary permissions first. Please run the 'requestPermissions' function.",
      "WARNING"
    );
    return;
  }

  // Now check all users in the list
  const users = getAuthorizedUsers();

  if (users.length === 0) {
    logWithUser(
      "No authorized users found. You may need to add users using the 'addUserToList' function.",
      "INFO"
    );
    // Automatically add the current user if they're not in the list but have permissions
    if (currentUserHasPermissions) {
      addUserToList(currentUser);
      logWithUser(
        `Added current user ${currentUser} to the registered users list`,
        "INFO"
      );
    }
  } else if (users.length > 1) {
    // Only display individual user status if there are more than one user
    logWithUser(`Checking permissions for ${users.length} users...`, "INFO");

    users.forEach((email) => {
      // Skip logging for current user to avoid redundancy
      if (email !== currentUser) {
        verifyUserPermissions(email); // This will log the results
      }
    });
  }

  logWithUser("Permission verification completed", "INFO");
}
