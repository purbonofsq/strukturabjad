To work with GitHub in Visual Studio Code, you will primarily use Git CLI commands in the integrated terminal (Ctrl + ~ or Cmd + ~) alongside a few VS Code-specific terminal shortcuts. [1] (https://code.visualstudio.com/docs/sourcecontrol/github), [2] (https://code.visualstudio.com/docs/configure/command-line), [3] (https://education.github.com/git-cheat-sheet-education.pdf)

1. Setup & Starting a Project

• git init: Initializes a new local Git repository in your current folder.
• git clone <repository-url>: Downloads an existing GitHub repository to your local machine.
• git config --global user.name "Your Name": Sets your username for commits.
• git config --global user.email "you@example.com": Sets your email address for commits. [1] (https://education.github.com/git-cheat-sheet-education.pdf)

2. Daily Workflow (Save & Sync)

• git status: Shows modified, staged, and untracked files in your workspace.
• git add .: Stages all modified and new files for the next commit. (Use git add <file> to stage a single file).
• git commit -m "Your message": Saves your staged changes with a descriptive message.
• git push: Uploads your local commits to your GitHub repository. (Use git push -u origin <branch-name> the first time you push a new branch).
• git pull: Fetches and merges the latest changes from GitHub into your local branch. [1] (https://www.youtube.com/watch?v=8ANnZn_8_Qg&t=303), [2] (https://education.github.com/git-cheat-sheet-education.pdf), [3] (https://dhanar98.hashnode.dev/working-with-github-basic-commands-in-visual-studio-code-part-1)

3. Branching & Merging

• git branch: Lists all local branches and highlights the current one.
• git checkout -b <branch-name>: Creates a new branch and immediately switches to it.
• git switch <branch-name> (or git checkout <branch-name>): Switches to an existing branch.
• git merge <branch-name>: Merges the specified branch into your currently active branch. [1] (https://amiladevin.medium.com/working-with-github-in-vs-code-7186b5c8f50d), [2] (https://www.youtube.com/watch?v=8ANnZn_8_Qg&t=303), [3] (https://education.github.com/git-cheat-sheet-education.pdf)

4. Inspection & Undoing Mistakes

• git log --oneline: Displays a compact history of past commits.
• git diff: Shows line-by-line unstaged changes in your files.
• git stash: Temporarily shelves uncommitted changes so you can switch branches cleanly.
• git stash pop: Restores your most recently stashed changes.
• git restore <file>: Discards uncommitted changes in a specific file. [1] (https://www.youtube.com/watch?v=8ANnZn_8_Qg&t=303), [2] (https://education.github.com/git-cheat-sheet-education.pdf)

5. VS Code Terminal Shortcuts

• code .: Opens the current terminal folder inside VS Code.
• code -d <file1> <file2>: Opens two files side-by-side in VS Code to compare differences. [1] (https://code.visualstudio.com/docs/configure/command-line)
If you tell me whether you are starting a new project, cloning an existing repo, or working on a team with pull requests, I can share the exact step-by-step sequence of commands for your workflow.

