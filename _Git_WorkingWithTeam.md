Here is the step-by-step Git and GitHub CLI (gh) workflow for working on a team with pull requests in VS Code, including automatic commit messages.

1. Start a New Feature or Bug Fix

Always create a fresh branch from an up-to-date main branch so your work stays isolated from your teammates' code.
bash
git checkout main
git pull origin main
git checkout -b feat/your-feature-name
Use code with caution.

2. Stage & Auto-Generate Your Commit Message

Once you finish coding in VS Code:
bash
git add .
Use code with caution.
• Option A (VS Code UI + Copilot): Open the Source Control tab (Ctrl + Shift + G), click the sparkle icon (✨) in the message box to generate a summary, and click Commit.
• Option B (Terminal CLI via OpenCommit): Run oco to auto-generate a Conventional Commit message (feat: ... or fix: ...), which teams prefer for clean PR histories.

3. Keep Your Branch Updated with main

Before opening a PR (or while waiting for reviews), pull any new changes your teammates merged into main to avoid merge conflicts:
bash
git fetch origin
git merge origin/main
Use code with caution.
(If VS Code highlights a merge conflict, use the built-in Merge Editor buttons to accept incoming/current changes, then run git add . and git commit.)

4. Push & Create a Pull Request

Push your branch to GitHub and open a PR directly from the VS Code terminal using the GitHub CLI (gh) (or the GitHub Pull Requests VS Code extension):
bash
git push -u origin feat/your-feature-name
gh pr create --fill
Use code with caution.
• --fill: Automatically uses your commit titles and descriptions to populate the Pull Request title and body.

5. Address Code Review Feedback

If a teammate requests changes on your PR:
1. Make the edits on the same branch (feat/your-feature-name).
2. Stage, auto-commit, and push again—your open PR updates automatically:bash
git add .
oco   # or commit via VS Code ✨ icon
git push
Use code with caution.

6. Clean Up After Merge

Once your PR is approved and merged on GitHub, switch back to main and delete your old local branch:
bash
git checkout main
git pull origin main
git branch -d feat/your-feature-name
Use code with caution.
Would you like instructions on installing and setting up the GitHub CLI (gh), or tips on resolving merge conflicts in VS Code?
