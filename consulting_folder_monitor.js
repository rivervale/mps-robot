function weeklyTriggers() { // Run once to create the triggers
  // Start folder monitoring at 6.30pm on Mondays
  ScriptApp.newTrigger('perMinuteTrigger')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.MONDAY)
    .atHour(18)
    .nearMinute(30)
    .inTimezone('Asia/Singapore')
    .create();

  // End folder monitoring at 12.30am on Tuesdays
  ScriptApp.newTrigger('terminatePerMinuteTrigger')
    .timeBased()
    .onWeekDay(ScriptApp.WeekDay.TUESDAY)
    .atHour(0)
    .nearMinute(30)
    .inTimezone('Asia/Singapore')
    .create();
}

// Run checkConsultingFolder() every 2 minutes
function perMinuteTrigger() { 
  ScriptApp.newTrigger('checkConsultingFolder')
    .timeBased()
    .everyMinutes(2)
    .create();
}

function terminatePerMinuteTrigger() {
  let triggers = getProjectTriggersByFunctionName('checkConsultingFolder');
  for (let i = 0; i < triggers.length; ++i)
    ScriptApp.deleteTrigger(triggers[i]);
}

function getProjectTriggersByFunctionName(functionName) {
  return ScriptApp.getProjectTriggers().filter(trigger => trigger.getHandlerFunction() === functionName);
}

function checkConsultingFolder() {

}