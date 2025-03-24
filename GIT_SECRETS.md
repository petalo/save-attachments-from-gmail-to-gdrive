# Git-Secrets Configuration

This project uses [git-secrets](https://github.com/awslabs/git-secrets) to prevent committing sensitive information to the repository. Git-secrets scans commits for sensitive information like API keys, passwords, and other credentials before they are committed.

## Setup

After installing the project dependencies with `npm install`, you need to run the setup script to configure git-secrets:

```bash
npm run setup-git-secrets
```

This only needs to be done once after cloning the repository. The script will configure git-secrets with the patterns defined in the `.git-secrets-patterns` file.

## Configuration Files

This project uses two main files for git-secrets configuration:

1. `.git-secrets-patterns` - Contains patterns of sensitive information to detect
2. `.gitallowed` - Contains patterns of false positives to allow

## What Git-Secrets Checks For

The current configuration checks for the following patterns:

- API keys with actual values (OpenAI, Gemini)
- Folder IDs and Script IDs with actual values
- OpenAI API keys (starting with "sk-" followed by 48 characters)
- Passwords, tokens, secrets, and credentials in string literals

The patterns are designed to detect actual sensitive information while minimizing false positives. They focus on detecting values rather than variable assignments.

## How Git-Secrets Works

Git-secrets stores its configuration in the Git configuration system. When you run the setup script, it adds patterns to your local Git configuration that will be checked before each commit.

The patterns are stored in the Git configuration under the `secrets.patterns` key. You can view the current patterns with:

```bash
git config --get-all secrets.patterns
```

## Sharing Configuration

The git-secrets configuration is stored in your local Git configuration and is not automatically shared when someone clones the repository. To share the configuration with other developers, we use:

1. `.git-secrets-patterns` file - Contains the patterns to detect
2. `utils/setup_git_secrets.sh` script - Installs the patterns from the file
3. Package.json scripts - Automatically run the setup script during installation

This ensures that everyone who installs the project dependencies will also get the git-secrets configuration.

## Allowed Patterns

Some patterns might be false positives, such as template strings in code. These patterns are added to the allowed list to prevent false positives. The current allowed patterns are stored in:

1. Git configuration under the `secrets.allowed` key
2. The `.gitallowed` file in the repository root

The setup script automatically adds these allowed patterns using:

```bash
git-secrets --add-provider -- cat .gitallowed
```

## Pre-commit Hook

The pre-commit hook is configured to run:

1. Git-secrets to scan for sensitive information
2. The existing test script

If any of these checks fail, the commit will be blocked.

## Manual Checking

You can manually check for secrets in your staged changes without making a commit by running:

```bash
npm run check-secrets
```

This will scan only the files that are staged for commit, focusing on actual changes rather than scanning the entire repository (which could produce false positives in documentation files).

## Bypassing Checks (Not Recommended)

In rare cases where you need to bypass the git-secrets check (e.g., for a false positive that can't be fixed), you can use:

```bash
git commit --no-verify
```

However, this is not recommended as it bypasses all pre-commit hooks, including important checks for sensitive information.

## Adding Custom Patterns

If you need to add custom patterns to check for specific sensitive information in your project:

1. Add the pattern to the `.git-secrets-patterns` file
2. Run `npm run setup-git-secrets` to update your configuration

For literal strings (not regular expressions), add them without any special syntax. For regular expressions, add them as-is.

## Adding Allowed Patterns

If you encounter false positives that should be allowed:

1. Add the pattern to the `.gitallowed` file
2. Run `npm run setup-git-secrets` to update your configuration
