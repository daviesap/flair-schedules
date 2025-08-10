To run local server :
cd "/Users/apndavies/Coding/Flair Schedules/flair-schedules"
firebase emulators:start --only functions

To run curl
curl -X POST "http://127.0.0.1:5001/flair-schedules/europe-west2/FlairScheduleHelper?action=mealsPivot" \
  -H "Content-Type: application/json" \
  --data-binary "@test-meals-large.json"


## Git cheat sheet 
  # 1. See all local branches
git branch

# 2. See all remote branches
git branch -r

# 3. Delete a branch locally
git branch -d branch_name
# (use -D instead of -d to force delete if it's not merged)

# 4. Delete a branch remotely (on GitHub)
git push origin --delete branch_name

# 5. Remove references to deleted remote branches from your local list
git fetch -p

# 6. Check again (should be gone)
git branch -a


## Basic Git Commands
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


# Move back from a branch to MAIN

#Make sure everything is commited to the branch first

# 1. Switch to main
git checkout main

# 2. Make sure main is up to date with GitHub
git pull origin main

# 3. Merge your branch into main
git merge change_date_format

# 4. Push the updated main back to GitHub
git push origin main

#If you want to delete the branch after merging (optional):
git branch -d change_date_format           # deletes locally
git push origin --delete change_date_format  # deletes on GitHub

# Before starting:
# ✅ Make sure all your changes on NEW_BRANCH are committed.
#    Run `git status` — it should say "nothing to commit, working tree clean".

# 1) Update local refs from GitHub
git fetch origin

# 2) Move your branch on top of the latest main
git checkout NEW_BRANCH
git rebase origin/main
# If there are conflicts:
#   - fix files, then:
#       git add <file>    # for each resolved file
#       git rebase --continue
#   - to bail out completely: git rebase --abort

# 3) Switch to main and fast-forward merge
git checkout main
git pull origin main
git merge --ff-only NEW_BRANCH

# 4) Push the updated main to GitHub
git push origin main

# 5) (Optional) Delete the branch
git branch -d NEW_BRANCH
git push origin --delete NEW_BRANCH