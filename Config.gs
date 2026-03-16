// =============================================================================
// Config.gs — Cycle Schedule Configuration
// =============================================================================
// This is the only file you need to edit to change your training cycle
// structure. Modify the DAY_SCHEDULE array below to match your program.
//
// Rules:
//   - Each entry needs: day (1-based number), type, and label
//   - Valid types: "Workout", "Active Rest", "Rest"
//   - The array length defines your cycle length (e.g. 7 entries = 7-day cycle)
//   - Day numbers must be sequential starting from 1
//
// Examples:
//
//   — Default 7-day cycle (3 workout, 2 active rest, 2 rest):
//       See below.
//
//   — 6-day PPL cycle (Push/Pull/Legs, repeated):
//       { day: 1, type: 'Workout', label: 'Push A' },
//       { day: 2, type: 'Workout', label: 'Pull A' },
//       { day: 3, type: 'Workout', label: 'Legs A' },
//       { day: 4, type: 'Workout', label: 'Push B' },
//       { day: 5, type: 'Workout', label: 'Pull B' },
//       { day: 6, type: 'Workout', label: 'Legs B' },
//
//   — 5-day Upper/Lower + Full Body:
//       { day: 1, type: 'Workout',     label: 'Upper' },
//       { day: 2, type: 'Workout',     label: 'Lower' },
//       { day: 3, type: 'Active Rest', label: 'Recovery' },
//       { day: 4, type: 'Workout',     label: 'Full Body' },
//       { day: 5, type: 'Rest',        label: 'Rest' },
//
// After editing, update your Program sheet's "Day Label" column to match
// the label values you define here (e.g. "Push A", "Pull A", etc.).
// =============================================================================

/**
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
