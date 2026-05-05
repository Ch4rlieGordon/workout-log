// =============================================================================
// Config.gs — Default schedule fallback
// =============================================================================
// All real schedule data lives in the Day Schedule sheet, keyed by
// (Block ID, Week Number). This file only provides a default 7-day schedule
// used when the user is creating their very first block in the editor.
// =============================================================================

function getDefaultDaySchedule() {
  return [
    { day: 1, type: 'Workout', label: 'Day 1' },
    { day: 2, type: 'Active Rest', label: 'Day 2' },
    { day: 3, type: 'Workout', label: 'Day 3' },
    { day: 4, type: 'Active Rest', label: 'Day 4' },
    { day: 5, type: 'Workout', label: 'Day 5' },
    { day: 6, type: 'Rest', label: 'Day 6' },
    { day: 7, type: 'Rest', label: 'Day 7' },
  ];
}
