To run local server :
cd /Users/apndavies/Coding/Flair Schedules/flair-schedules
firebase emulators:start --only functions

To run curl
curl -X POST "http://127.0.0.1:5001/flair-schedules/europe-west2/FlairScheduleHelper?action=mealsPivot" \
  -H "Content-Type: application/json" \
  --data-binary "@test-meals-large.json"