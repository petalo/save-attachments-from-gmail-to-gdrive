#!/bin/bash
# Script to set up git-secrets for this repository

# Check if git-secrets is installed
if ! command -v git-secrets &>/dev/null; then
    echo "git-secrets is not installed. Please install it first."
    echo "For macOS: brew install git-secrets"
    echo "For other systems: https://github.com/awslabs/git-secrets#installing-git-secrets"
    exit 1
fi

# Install git-secrets hooks
git-secrets --install

# Add patterns from .git-secrets-patterns file
echo "Adding patterns from .git-secrets-patterns..."
while IFS= read -r line || [[ -n "$line" ]]; do
    # Skip comments and empty lines
    if [[ ! "$line" =~ ^# ]] && [[ -n "$line" ]]; then
        git-secrets --add "$line"
        echo "Added pattern: $line"
    fi
done <.git-secrets-patterns

# Add allowed patterns from .gitallowed file if it exists
if [ -f .gitallowed ]; then
    echo "Adding allowed patterns from .gitallowed..."
    git-secrets --add-provider -- cat .gitallowed
fi

echo "git-secrets has been configured for this repository."
echo "The pre-commit hook will now check for secrets before each commit."
