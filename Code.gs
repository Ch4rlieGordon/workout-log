// =============================================================================
// Code.gs — Workout Log Tool (Server-Side)
// =============================================================================

const SHEET_LOG       = 'Workout Log';
const SHEET_EXERCISES = 'Exercise Selection';
const SHEET_PROGRAM   = 'Program';
const SHEET_SETTINGS  = 'Settings';

// =============================================================================
// Entry Points
// =============================================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Workout Log')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Workout Log')
    .addItem('Open Sidebar', 'openSidebar')
    .addItem('Setup Sheets', 'setupAllSheets')
    .addToUi();
}

function openSidebar() {
  const html = HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Workout Log');
  SpreadsheetApp.getUi().showSidebar(html);
}

// =============================================================================
// Sheet Setup — Run Once
// =============================================================================

function setupAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupExerciseSheet_(ss);
  setupLogSheet_(ss);
  setupProgramSheet_(ss);
  setupSettingsSheet_(ss);
  SpreadsheetApp.getUi().alert('All sheets created and populated.');
}

function setupExerciseSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_EXERCISES);
  if (!sheet) sheet = ss.insertSheet(SHEET_EXERCISES);
  sheet.clear();

  const headers = ['Exercise Name', 'Primary Muscle Group', 'Equipment', 'Notes'];
  const data = [
    headers,
    ['Paused Bench Press',          'Chest',     'Barbell',      'Pause at chest for 1-2s'],
    ['Larsen Bench Press',          'Chest',     'Barbell',      'Legs straight, feet off floor'],
    ['Bench Press TnG',             'Chest',     'Barbell',      'Touch and go, controlled'],
    ['Overhead Press',              'Shoulders', 'Barbell',      'Strict press, no leg drive'],
    ['Lateral Raises',              'Shoulders', 'Dumbbell',     'Slight bend in elbows'],
    ['Rear Delt Flies',             'Shoulders', 'Dumbbell',     'Bent over or incline bench'],
    ['Chin Ups',                    'Back',      'Bodyweight',   'Supinated grip'],
    ['Pull Ups',                    'Back',      'Bodyweight',   'Pronated grip'],
    ['Cable Rows',                  'Back',      'Cable',        'Squeeze at contraction'],
    ['Seal Rows',                   'Back',      'Barbell',      'Chest-supported on bench'],
    ['Hammer Curls',                'Biceps',    'Dumbbell',     'Neutral grip throughout'],
    ['Preacher Hammer Curls',       'Biceps',    'Dumbbell',     'Neutral grip on preacher pad'],
    ['Concentration Curl',          'Biceps',    'Dumbbell',     'Elbow braced on inner thigh'],
    ['Overhead Triceps Extensions', 'Triceps',   'Cable',        'Rope or bar attachment'],
    ['Triceps Pushdown',            'Triceps',   'Cable',        'Elbows pinned to sides'],
    ['Dips',                        'Triceps',   'Bodyweight',   'Slight forward lean for chest'],
    ['Reverse Curl',                'Forearms',  'Barbell',      'Pronated grip, EZ bar ok'],
    ['Dumbbell Wrist Curl',         'Forearms',  'Dumbbell',     'Forearms on thighs, palms up'],
    ['Dumbbell Wrist Extensions',   'Forearms',  'Dumbbell',     'Forearms on thighs, palms down'],
    ['Wrist Roller',                'Forearms',  'Wrist Roller', 'Roll up and down slowly'],
    ['Kettlebell Farmer Walks',     'Forearms',  'Kettlebell',   'Tall posture, squeeze grip'],
    ['Abs Decline Situps',          'Core',      'Bodyweight',   'Hold plate for added resistance'],
    ['Abs Rope Pushdown',           'Core',      'Cable',        'Kneeling cable crunch'],
    ['Abs Wheel',                   'Core',      'Ab Wheel',     'Controlled rollout and return'],
    ['High Bar Squat',              'Legs',      'Barbell',      'Bar on traps, upright torso'],
    ['Low Bar Squat',               'Legs',      'Barbell',      'Bar on rear delts, hip hinge'],
    ['Hyperextensions',             'Legs',      'Bodyweight',   'Targets glutes and lower back'],
    ['Romanian Deadlift',           'Legs',      'Barbell',      'Hinge at hips, slight knee bend'],
    ['Seated Leg Curl',             'Legs',      'Machine',      'Full ROM, squeeze at bottom'],
    ['Lying Leg Curl',              'Legs',      'Machine',      'Hips flat on pad'],
  ];

  sheet.getRange(1, 1, data.length, headers.length).setValues(data);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

function setupLogSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet) sheet = ss.insertSheet(SHEET_LOG);

  if (sheet.getLastRow() === 0) {
    const headers = [
      'Date', 'Day Label', 'Day Type', 'Exercise Name',
      'Set Number', 'Weight (kg)', 'Reps', 'RPE', 'Timestamp'
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function setupProgramSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_PROGRAM);
  if (!sheet) sheet = ss.insertSheet(SHEET_PROGRAM);
  sheet.clear();

  const headers = [
    'Day Label', 'Exercise Name', 'Target Sets',
    'Target Reps', 'Target RPE', 'Exercise Order'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);
}

function setupSettingsSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) sheet = ss.insertSheet(SHEET_SETTINGS);
  sheet.clear();

  sheet.getRange(1, 1, 2, 2).setValues([
    ['Key', 'Value'],
    ['startDate', '']   // User fills in, e.g. 2026-03-09
  ]);
  sheet.setFrozenRows(1);
}

// =============================================================================
// Data Access — Initial Payload
// =============================================================================

/**
 * Returns everything needed at page load:
 *  - cycleInfo:   startDate, current defaults, schedule
 *  - exercises:   full exercise database
 *  - fullProgram: program template for ALL days
 *  - allLoggedExercises: for the progress dropdown
 *  - loggedDates: array of YYYY-MM-DD strings that have log entries
 *                 (lets the client show "Modify" vs "Start" instantly)
 */
function getInitData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  return {
    cycleInfo:           getCycleInfo_(ss),
    exercises:           getExercises_(ss),
    fullProgram:         getFullProgram_(ss),
    allLoggedExercises:  getLoggedExerciseNames_(ss),
    loggedDates:         getLoggedDates_(ss),
  };
}

/**
 * Called when the user picks a day and taps Start/Modify.
 * Returns:
 *  - lastSession:  previous session reference per exercise (EXCLUDING dateStr)
 *  - existingLog:  data already logged for dateStr (empty object if none)
 *
 * @param {string} dayLabel  — e.g. "Day 1" (used to find program exercises)
 * @param {string} dateStr   — YYYY-MM-DD of the session being edited
 */
function getSessionSetup(dayLabel, dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Gather exercise names from the program for this day
  const programNames = getFullProgram_(ss)
    .filter(p => p.dayLabel === dayLabel)
    .map(p => p.exerciseName);

  // Also find any extra exercises already logged for this date
  // (user might have added exercises outside the template)
  const existingLog = getExistingLogForDate_(ss, dateStr);
  const loggedNames = Object.keys(existingLog);

  // Union of program + logged exercise names (deduped)
  const allNames = [...new Set([...programNames, ...loggedNames])];

  // Get last-session reference data, explicitly excluding the date being edited
  const lastSession = getLastSessionData_(ss, allNames, dateStr);

  return { lastSession, existingLog };
}

// =============================================================================
// Internal Helpers
// =============================================================================

/**
 * Robustly parse the start date from the Settings sheet.
 * Handles both native Date objects (from formatted cells) and strings.
 */
function getCycleInfo_(ss) {
  const tz = Session.getScriptTimeZone();
  const settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  let startDateStr = '';

  if (settingsSheet && settingsSheet.getLastRow() >= 2) {
    const raw = settingsSheet.getRange(2, 2).getValue();

    if (raw instanceof Date) {
      // Cell formatted as a date — Sheets gives us a Date object directly.
      // This avoids any DD/MM vs MM/DD locale ambiguity.
      startDateStr = Utilities.formatDate(raw, tz, 'yyyy-MM-dd');
    } else if (raw) {
      // Plain string — assume YYYY-MM-DD or try to parse whatever it is
      const s = String(raw).trim();
      const parsed = new Date(s);
      if (!isNaN(parsed.getTime())) {
        startDateStr = Utilities.formatDate(parsed, tz, 'yyyy-MM-dd');
      }
    }
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);

  if (!startDateStr) {
    // First use: anchor cycle to today and save
    startDateStr = Utilities.formatDate(today, tz, 'yyyy-MM-dd');
    if (settingsSheet) {
      settingsSheet.getRange(2, 2).setValue(startDateStr);
    }
  }

  const startDate = new Date(startDateStr + 'T00:00:00');
  const msPerDay = 86400000;
  const daysElapsed = Math.floor((today - startDate) / msPerDay);

  const schedule = getDaySchedule();
  const cycleLength = schedule.length;

  // Guard against start date being in the future
  const safeDays = Math.max(daysElapsed, 0);
  const currentDayIndex = safeDays % cycleLength;               // 0-based
  const currentCycleNumber = Math.floor(safeDays / cycleLength) + 1;

  return {
    startDate: startDateStr,
    currentCycleNumber,
    currentDayNumber: currentDayIndex + 1,  // 1-based
    todayStr: Utilities.formatDate(today, tz, 'yyyy-MM-dd'),
    daySchedule: schedule,
  };
}

function getExercises_(ss) {
  const sheet = ss.getSheetByName(SHEET_EXERCISES);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  return data.map(r => ({
    name: r[0],
    muscleGroup: r[1],
    equipment: r[2],
    notes: r[3] || '',
  }));
}

function getFullProgram_(ss) {
  const sheet = ss.getSheetByName(SHEET_PROGRAM);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 6).getValues();
  return data.map(r => ({
    dayLabel: r[0],
    exerciseName: r[1],
    targetSets: r[2],
    targetReps: String(r[3]),
    targetRPE: r[4] || '',
    order: r[5] || 0,
  }));
}

/**
 * Get all unique dates (YYYY-MM-DD) that have at least one log entry.
 * The client uses this to decide "Start Workout" vs "Modify Workout Log".
 */
function getLoggedDates_(ss) {
  const sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const tz = Session.getScriptTimeZone();
  const col = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
  const dates = new Set();
  for (const v of col) {
    if (v instanceof Date) {
      dates.add(Utilities.formatDate(v, tz, 'yyyy-MM-dd'));
    } else if (v) {
      dates.add(String(v).trim());
    }
  }
  return [...dates].sort();
}

/**
 * Return all log entries for a specific date, grouped by exercise name.
 * Each exercise maps to an array of { setNumber, weight, reps, rpe }.
 */
function getExistingLogForDate_(ss, dateStr) {
  const sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet || sheet.getLastRow() < 2) return {};

  const tz = Session.getScriptTimeZone();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  const result = {};

  for (const r of data) {
    let rowDate;
    if (r[0] instanceof Date) {
      rowDate = Utilities.formatDate(r[0], tz, 'yyyy-MM-dd');
    } else {
      rowDate = String(r[0]).trim();
    }

    if (rowDate !== dateStr) continue;

    const exName = r[3];
    if (!result[exName]) result[exName] = [];
    result[exName].push({
      setNumber: r[4],
      weight: r[5],
      reps: r[6],
      rpe: r[7],
    });
  }

  // Sort each exercise's sets by set number
  for (const name of Object.keys(result)) {
    result[name].sort((a, b) => a.setNumber - b.setNumber);
  }

  return result;
}

/**
 * For each exercise, find the most recent session's sets,
 * EXCLUDING a specific date (the one being edited).
 */
function getLastSessionData_(ss, exerciseNames, excludeDateStr) {
  if (!exerciseNames || !exerciseNames.length) return {};

  const sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet || sheet.getLastRow() < 2) return {};

  const tz = Session.getScriptTimeZone();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  const result = {};

  for (const name of exerciseNames) {
    const rows = data
      .filter(r => {
        if (r[3] !== name) return false;
        // Exclude the date being edited so "Prev" always means a different session
        let rowDate;
        if (r[0] instanceof Date) {
          rowDate = Utilities.formatDate(r[0], tz, 'yyyy-MM-dd');
        } else {
          rowDate = String(r[0]).trim();
        }
        return rowDate !== excludeDateStr;
      })
      .sort((a, b) => {
        const dateComp = new Date(b[0]) - new Date(a[0]);
        return dateComp !== 0 ? dateComp : a[4] - b[4];
      });

    if (rows.length === 0) { result[name] = []; continue; }

    const latestDateStr = rows[0][0] instanceof Date
      ? Utilities.formatDate(rows[0][0], tz, 'yyyy-MM-dd')
      : String(rows[0][0]).trim();

    result[name] = rows
      .filter(r => {
        const rd = r[0] instanceof Date
          ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd')
          : String(r[0]).trim();
        return rd === latestDateStr;
      })
      .map(r => ({
        setNumber: r[4],
        weight: r[5],
        reps: r[6],
        rpe: r[7],
      }));
  }

  return result;
}

function getLoggedExerciseNames_(ss) {
  const sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const col = sheet.getRange(2, 4, sheet.getLastRow() - 1, 1).getValues().flat();
  return [...new Set(col)].sort();
}

// =============================================================================
// Write — Save (or Overwrite) a Workout Session
// =============================================================================

/**
 * Save a workout session. If entries already exist for this date,
 * delete them first (overwrite), then write the new data.
 * This makes "Modify" safe — no duplicate rows.
 */
function saveWorkout(sessionData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG);
  const tz = Session.getScriptTimeZone();
  const timestamp = new Date();

  // --- 1. Delete any existing rows for this date ---
  if (sheet.getLastRow() >= 2) {
    const dates = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    // Walk backwards so row deletions don't shift indices
    for (let i = dates.length - 1; i >= 0; i--) {
      let rowDate;
      if (dates[i] instanceof Date) {
        rowDate = Utilities.formatDate(dates[i], tz, 'yyyy-MM-dd');
      } else {
        rowDate = String(dates[i]).trim();
      }
      if (rowDate === sessionData.date) {
        sheet.deleteRow(i + 2); // +2 because data starts at row 2
      }
    }
  }

  // --- 2. Write new rows ---
  const rows = sessionData.sets.map(s => [
    sessionData.date,
    sessionData.dayLabel,
    sessionData.dayType,
    s.exercise,
    s.setNumber,
    s.weight,
    s.reps,
    s.rpe || '',
    timestamp,
  ]);

  if (rows.length === 0) {
    return { success: false, message: 'No sets to save.' };
  }

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, 9).setValues(rows);

  return { success: true, rowsWritten: rows.length };
}

// =============================================================================
// Delete — Remove All Log Entries for a Date
// =============================================================================

/**
 * Delete every row in the Workout Log that matches the given date.
 * Used for "Delete Entire Session".
 * @param {string} dateStr — YYYY-MM-DD
 * @returns {Object} { success, rowsDeleted }
 */
function deleteWorkout(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG);
  const tz = Session.getScriptTimeZone();

  if (!sheet || sheet.getLastRow() < 2) {
    return { success: false, message: 'No log data found.' };
  }

  const dates = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
  let deletedCount = 0;

  for (let i = dates.length - 1; i >= 0; i--) {
    const rowDate = normalizeDateCell_(dates[i], tz);
    if (rowDate === dateStr) {
      sheet.deleteRow(i + 2);
      deletedCount++;
    }
  }

  return { success: true, rowsDeleted: deletedCount };
}

/**
 * Delete rows for a SINGLE exercise on a specific date.
 * Used for granular per-exercise deletion.
 * @param {string} dateStr      — YYYY-MM-DD
 * @param {string} exerciseName — exact exercise name to remove
 * @returns {Object} { success, rowsDeleted, remainingExercises }
 *   remainingExercises: how many distinct exercises still have entries on this date
 *   (lets the client know if the day is now fully empty)
 */
function deleteSingleExercise(dateStr, exerciseName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG);
  const tz = Session.getScriptTimeZone();

  if (!sheet || sheet.getLastRow() < 2) {
    return { success: false, message: 'No log data found.' };
  }

  // Read date (col 1) and exercise name (col 4) for all data rows
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  let deletedCount = 0;
  const remainingNames = new Set();

  // Walk backwards for safe row deletion
  for (let i = data.length - 1; i >= 0; i--) {
    const rowDate = normalizeDateCell_(data[i][0], tz);
    if (rowDate !== dateStr) continue;

    const rowExercise = data[i][3];

    if (rowExercise === exerciseName) {
      sheet.deleteRow(i + 2);
      deletedCount++;
    } else {
      // Track which other exercises remain for this date
      remainingNames.add(rowExercise);
    }
  }

  return {
    success: true,
    rowsDeleted: deletedCount,
    remainingExercises: remainingNames.size,
  };
}

/**
 * Fetch the list of exercises (with set counts) logged on a given date.
 * Powers the delete-picker modal so the user can choose what to remove.
 * @param {string} dateStr — YYYY-MM-DD
 * @returns {Array<{name, setCount}>} sorted by the order they appear in the log
 */
function getLoggedExercisesForDate(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG);
  const tz = Session.getScriptTimeZone();

  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 4).getValues();
  const counts = {};       // exerciseName → set count
  const firstSeen = {};    // exerciseName → row index (for ordering)

  for (let i = 0; i < data.length; i++) {
    const rowDate = normalizeDateCell_(data[i][0], tz);
    if (rowDate !== dateStr) continue;

    const exName = data[i][3];
    counts[exName] = (counts[exName] || 0) + 1;
    if (!(exName in firstSeen)) firstSeen[exName] = i;
  }

  return Object.keys(counts)
    .sort((a, b) => firstSeen[a] - firstSeen[b])
    .map(name => ({ name: name, setCount: counts[name] }));
}

/**
 * Helper: normalize a date cell value (could be Date object or string)
 * to a consistent YYYY-MM-DD string.
 */
function normalizeDateCell_(cellValue, tz) {
  if (cellValue instanceof Date) {
    return Utilities.formatDate(cellValue, tz, 'yyyy-MM-dd');
  }
  return String(cellValue).trim();
}

// =============================================================================
// Progress Data
// =============================================================================

function getProgressData(exerciseName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet || sheet.getLastRow() < 2) return { sessions: [], bestE1RM: null };

  const tz = Session.getScriptTimeZone();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  const rows = data.filter(r => r[3] === exerciseName);

  if (rows.length === 0) return { sessions: [], bestE1RM: null };

  const byDate = {};
  for (const r of rows) {
    const dateStr = r[0] instanceof Date
      ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd')
      : String(r[0]).trim();
    if (!byDate[dateStr]) byDate[dateStr] = [];
    byDate[dateStr].push({ weight: r[5], reps: r[6] });
  }

  let bestE1RM = 0;
  let bestDate = '';
  const sessions = [];

  for (const [date, sets] of Object.entries(byDate).sort()) {
    let maxE1RM = 0;
    let totalVolume = 0;

    for (const s of sets) {
      const w = Number(s.weight) || 0;
      const r = Number(s.reps) || 0;
      totalVolume += w * r;
      const e1rm = r > 0 ? w * (1 + r / 30) : 0;
      if (e1rm > maxE1RM) maxE1RM = e1rm;
    }

    if (maxE1RM > bestE1RM) {
      bestE1RM = maxE1RM;
      bestDate = date;
    }

    sessions.push({
      date,
      e1rm: Math.round(maxE1RM * 10) / 10,
      volume: Math.round(totalVolume),
      setCount: sets.length,
    });
  }

  let volumeChange = null;
  if (sessions.length >= 2) {
    const curr = sessions[sessions.length - 1].volume;
    const prev = sessions[sessions.length - 2].volume;
    volumeChange = prev > 0 ? Math.round(((curr - prev) / prev) * 1000) / 10 : null;
  }

  return {
    sessions,
    bestE1RM: { value: Math.round(bestE1RM * 10) / 10, date: bestDate },
    totalSessions: sessions.length,
    volumeChange,
  };
}

// =============================================================================
// Draft Auto-Save
// =============================================================================

function saveDraft(draftJson) {
  PropertiesService.getScriptProperties().setProperty('workoutDraft', draftJson);
}

function loadDraft() {
  return PropertiesService.getScriptProperties().getProperty('workoutDraft') || null;
}

function clearDraft() {
  PropertiesService.getScriptProperties().deleteProperty('workoutDraft');
}
