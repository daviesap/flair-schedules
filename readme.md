To run local server :
cd /Users/apndavies/Coding/Flair Schedules/flair-schedules
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