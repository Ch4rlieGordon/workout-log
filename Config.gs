// =============================================================================
// Config.gs — Cycle Schedule Configuration
// =============================================================================
// The schedule is now managed through the app's Program Editor UI,
// which writes to the "Day Schedule" sheet. This file provides the
// fallback default for first-time use before the user sets up their program.
//
// The default maps to a standard Monday–Sunday week:
//   Day 1 = Monday, Day 7 = Sunday
// =============================================================================

function getDaySchedule() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName('Day Schedule');

  if (sheet && sheet.getLastRow() >= 2) {
    var data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
    var schedule = data
      .filter(function(r) { return r[0] && r[1]; })
      .map(function(r) {
        return {
          day: Number(r[0]),
          type: String(r[1]).trim(),
          label: String(r[2]).trim() || ('Day ' + r[0]),
        };
      });
    if (schedule.length > 0) return schedule;
  }

  // Default 7-day schedule (Mon–Sun)
  return [
    { day: 1, type: 'Workout',     label: 'Day 1' },
    { day: 2, type: 'Workout',     label: 'Day 2' },
    { day: 3, type: 'Active Rest', label: 'Day 3' },
    { day: 4, type: 'Workout',     label: 'Day 4' },
    { day: 5, type: 'Workout',     label: 'Day 5' },
    { day: 6, type: 'Active Rest', label: 'Day 6' },
    { day: 7, type: 'Rest',        label: 'Day 7' },
  ];
}/**
 * Returns the training cycle schedule.
 * Edit this array to change your cycle structure.
 */
function getDaySchedule() {
  return [
    { day: 1, type: 'Workout',     label: 'Day 1' },
    { day: 2, type: 'Active Rest', label: 'Day 2' },
    { day: 3, type: 'Workout',     label: 'Day 3' },
    { day: 4, type: 'Active Rest', label: 'Day 4' },
    { day: 5, type: 'Workout',     label: 'Day 5' },
    { day: 6, type: 'Rest',        label: 'Day 6' },
    { day: 7, type: 'Rest',        label: 'Day 7' },
  ];
}
