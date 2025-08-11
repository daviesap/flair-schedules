To run local server :
cd "/Users/apndavies/Coding/Flair Schedules/flair-schedules"
firebase emulators:start --only functions

To run curl
curl -X POST "http://127.0.0.1:5001/flair-schedules/europe-west2/FlairScheduleHelper?action=mealsPivot" \
  -H "Content-Type: application/json" \
  --data-binary "@test-meals-large.json"



## Git Cheat Sheet

## Finishing work

# 1. Stage all changes
git add .

# 2. Commit with a useful message
git commit -m "WIP: brief description of what you did"

# 3. Push your branch to GitHub (usually main)
git push origin main

## Starting work
# 1. Make sure you're on the correct branch
git checkout main

# 2. Pull the latest changes from GitHub
git pull origin main




### Branch Management
```sh
# List all local branches
git branch

# List all remote branches
git branch -r

# Delete a branch locally (replace NEW_BRANCH with your branch name)
git branch -d NEW_BRANCH          # deletes locally (safe)
git branch -D NEW_BRANCH          # force delete locally (not merged)

# Delete a branch remotely (on GitHub)
git push origin --delete NEW_BRANCH

# Remove references to deleted remote branches from your local list
git fetch -p

# Show all branches (local and remote)
git branch -a
```

### Basic Git Commands
```sh
# Clone a repository
git clone <repo_url>

# Check the status of your working directory
git status

# Add all changes in current directory and subdirectories
git add .

# Commit staged changes with a message
git commit -m "Your commit message"

# Push committed changes to the remote repository
git push

# Pull latest changes from the remote repository
git pull
```

## Merging or Rebasing a Feature Branch into Main

> **Note:** Before merging or rebasing, make sure all your changes on `NEW_BRANCH` are committed.  
> Run `git status` — it should say "nothing to commit, working tree clean".

### Fast-Forward Merge Workflow
```sh
# 1. Switch to main branch
git checkout main

# 2. Make sure main is up to date with GitHub
git pull origin main

# 3. Merge your feature branch into main (fast-forward if possible)
git merge --ff-only NEW_BRANCH

# 4. Push the updated main back to GitHub
git push origin main
```

### Rebase Workflow (Keeping History Linear)
```sh
# 1. Update local refs from GitHub
git fetch origin

# 2. Move your branch on top of the latest main
git checkout NEW_BRANCH
git rebase origin/main
# If there are conflicts:
#   - Fix files, then:
#       git add <file>    # for each resolved file
#       git rebase --continue
#   - To bail out completely: git rebase --abort

# 3. Switch to main and fast-forward merge
git checkout main
git pull origin main
git merge --ff-only NEW_BRANCH

# 4. Push the updated main to GitHub
git push origin main
```

### Deleting a Feature Branch (After Merge)
```sh
# Delete branch locally
git branch -d NEW_BRANCH

# Delete branch remotely (on GitHub)
git push origin --delete NEW_BRANCH
```