// =============================================================================
// Code.gs — Workout Log Tool (Server-Side)
// =============================================================================

const SHEET_LOG       = 'Workout Log';
const SHEET_EXERCISES = 'Exercise Selection';
const SHEET_PROGRAM   = 'Program';
const SHEET_SETTINGS  = 'Settings';
const SHEET_BODY      = 'Body Log';
const SHEET_SCHEDULE  = 'Day Schedule';

const LOG_HEADERS = [
  'Date', 'Day Label', 'Day Type', 'Exercise Name', 'Set Number',
  'Weight (kg)', 'Reps', 'Minutes', 'Seconds', 'Distance (km)',
  'RPE', 'Notes', 'Timestamp'
];
const LOG_COLS = LOG_HEADERS.length; // 13

const EX_HEADERS = [
  'Exercise Name', 'Primary Muscle Group', 'Equipment', 'Notes',
  'Measurement Type', 'Cardio Modality'
];
const EX_COLS = EX_HEADERS.length; // 6

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
  const html = HtmlService.createHtmlOutputFromFile('Index').setTitle('Workout Log');
  SpreadsheetApp.getUi().showSidebar(html);
}

// =============================================================================
// Sheet Setup / Migration
// =============================================================================

function setupAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupExerciseSheet_(ss);
  setupLogSheet_(ss);
  setupProgramSheet_(ss);
  setupSettingsSheet_(ss);
  setupBodyLogSheet_(ss);
  setupDayScheduleSheet_(ss);
  SpreadsheetApp.getUi().alert('All sheets created and migrated.');
}

function setupExerciseSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_EXERCISES);
  if (!sheet) sheet = ss.insertSheet(SHEET_EXERCISES);

  // Migrate existing sheet if columns are missing
  if (sheet.getLastRow() > 0) {
    const currentCols = sheet.getLastColumn();
    if (currentCols < EX_COLS) {
      // Insert new columns after col 4 (Notes)
      sheet.insertColumnsAfter(4, EX_COLS - currentCols);
      sheet.getRange(1, 1, 1, EX_COLS).setValues([EX_HEADERS]);
      // Default existing rows to weight_reps
      if (sheet.getLastRow() >= 2) {
        const rows = sheet.getLastRow() - 1;
        const defaults = Array(rows).fill(['weight_reps', '']);
        sheet.getRange(2, 5, rows, 2).setValues(defaults);
      }
    }
    return;
  }

  // Fresh setup with seed data
  const data = [
    EX_HEADERS,
    ['Paused Bench Press',          'Chest',     'Barbell',      'Pause at chest for 1-2s',      'weight_reps', ''],
    ['Larsen Bench Press',          'Chest',     'Barbell',      'Legs straight, feet off floor','weight_reps', ''],
    ['Bench Press TnG',             'Chest',     'Barbell',      'Touch and go, controlled',     'weight_reps', ''],
    ['Overhead Press',              'Shoulders', 'Barbell',      'Strict press, no leg drive',   'weight_reps', ''],
    ['Lateral Raises',              'Shoulders', 'Dumbbell',     'Slight bend in elbows',        'weight_reps', ''],
    ['Rear Delt Flies',             'Shoulders', 'Dumbbell',     'Bent over or incline bench',   'weight_reps', ''],
    ['Chin Ups',                    'Back',      'Bodyweight',   'Supinated grip',               'weight_reps', ''],
    ['Pull Ups',                    'Back',      'Bodyweight',   'Pronated grip',                'weight_reps', ''],
    ['Cable Rows',                  'Back',      'Cable',        'Squeeze at contraction',       'weight_reps', ''],
    ['Seal Rows',                   'Back',      'Barbell',      'Chest-supported on bench',     'weight_reps', ''],
    ['Hammer Curls',                'Biceps',    'Dumbbell',     'Neutral grip throughout',      'weight_reps', ''],
    ['Preacher Hammer Curls',       'Biceps',    'Dumbbell',     'Neutral grip on preacher pad', 'weight_reps', ''],
    ['Concentration Curl',          'Biceps',    'Dumbbell',     'Elbow braced on inner thigh',  'weight_reps', ''],
    ['Overhead Triceps Extensions', 'Triceps',   'Cable',        'Rope or bar attachment',       'weight_reps', ''],
    ['Triceps Pushdown',            'Triceps',   'Cable',        'Elbows pinned to sides',       'weight_reps', ''],
    ['Dips',                        'Triceps',   'Bodyweight',   'Slight forward lean for chest','weight_reps', ''],
    ['Reverse Curl',                'Forearms',  'Barbell',      'Pronated grip, EZ bar ok',     'weight_reps', ''],
    ['Dumbbell Wrist Curl',         'Forearms',  'Dumbbell',     'Forearms on thighs, palms up', 'weight_reps', ''],
    ['Dumbbell Wrist Extensions',   'Forearms',  'Dumbbell',     'Forearms on thighs, palms dn', 'weight_reps', ''],
    ['Wrist Roller',                'Forearms',  'Wrist Roller', 'Roll up and down slowly',      'weight_reps', ''],
    ['Kettlebell Farmer Walks',     'Forearms',  'Kettlebell',   'Tall posture, squeeze grip',   'weight_time', ''],
    ['Plank',                       'Core',      'Bodyweight',   'Neutral spine, tight glutes',  'time_only',   ''],
    ['Abs Decline Situps',          'Core',      'Bodyweight',   'Hold plate for resistance',    'weight_reps', ''],
    ['Abs Rope Pushdown',           'Core',      'Cable',        'Kneeling cable crunch',        'weight_reps', ''],
    ['Abs Wheel',                   'Core',      'Ab Wheel',     'Controlled rollout',           'weight_reps', ''],
    ['High Bar Squat',              'Legs',      'Barbell',      'Bar on traps, upright torso',  'weight_reps', ''],
    ['Low Bar Squat',               'Legs',      'Barbell',      'Bar on rear delts, hip hinge', 'weight_reps', ''],
    ['Hyperextensions',             'Legs',      'Bodyweight',   'Glutes and lower back',        'weight_reps', ''],
    ['Romanian Deadlift',           'Legs',      'Barbell',      'Hinge at hips, slight knee',   'weight_reps', ''],
    ['Seated Leg Curl',             'Legs',      'Machine',      'Full ROM, squeeze at bottom',  'weight_reps', ''],
    ['Lying Leg Curl',              'Legs',      'Machine',      'Hips flat on pad',             'weight_reps', ''],
    ['Easy Run',                    'Cardio',    'Treadmill',    'Conversation pace',            'cardio',      'steady_slow'],
    ['HIIT Cardio',                 'Cardio',    'Any',          'Work/rest intervals',          'cardio',      'hiit'],
  ];
  sheet.getRange(1, 1, data.length, EX_COLS).setValues(data);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, EX_COLS);
}

function setupLogSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet) sheet = ss.insertSheet(SHEET_LOG);

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, LOG_COLS).setValues([LOG_HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }

  // Migrate from 9-col schema to 13-col schema if needed
  const currentCols = sheet.getLastColumn();
  if (currentCols < LOG_COLS) {
    // Old: ..., Reps (7), RPE (8), Timestamp (9)
    // New: ..., Reps (7), Minutes (8), Seconds (9), Distance (10), RPE (11), Notes (12), Timestamp (13)
    sheet.insertColumnsAfter(7, 3);  // Minutes, Seconds, Distance after Reps
    sheet.insertColumnsAfter(11, 1); // Notes after new RPE position
    sheet.getRange(1, 1, 1, LOG_COLS).setValues([LOG_HEADERS]);
  }
}

function setupProgramSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_PROGRAM);
  if (!sheet) sheet = ss.insertSheet(SHEET_PROGRAM);
  if (sheet.getLastRow() === 0) {
    const headers = ['Day Label', 'Exercise Name', 'Target Sets', 'Target Reps', 'Target RPE', 'Exercise Order'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }
}

function setupSettingsSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) sheet = ss.insertSheet(SHEET_SETTINGS);

  if (sheet.getLastRow() >= 2) {
    const keys = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().flat();
    const needed = [['heightCm', ''], ['gender', 'male'], ['birthYear', ''], ['activityLevel', 'moderate']];
    for (const [key, def] of needed) {
      if (!keys.includes(key)) {
        const nr = sheet.getLastRow() + 1;
        sheet.getRange(nr, 1, 1, 2).setValues([[key, def]]);
      }
    }
    return;
  }
  sheet.clear();
  sheet.getRange(1, 1, 6, 2).setValues([
    ['Key', 'Value'], ['startDate', ''], ['heightCm', ''],
    ['gender', 'male'], ['birthYear', ''], ['activityLevel', 'moderate'],
  ]);
  sheet.setFrozenRows(1);
}

function setupBodyLogSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_BODY);
  if (!sheet) sheet = ss.insertSheet(SHEET_BODY);
  if (sheet.getLastRow() === 0) {
    const headers = ['Date', 'Weight (kg)', 'Waist (cm)', 'Neck (cm)', 'Hips (cm)', 'Chest (cm)',
      'Bicep L (cm)', 'Bicep R (cm)', 'Quad L (cm)', 'Quad R (cm)', 'Calves (cm)', 'Shoulders (cm)', 'Timestamp'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
}

function setupDayScheduleSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_SCHEDULE);
  if (!sheet) sheet = ss.insertSheet(SHEET_SCHEDULE);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 3).setValues([['Day', 'Type', 'Label']]);
    sheet.setFrozenRows(1);
  }
}

// =============================================================================
// Initial Payload
// =============================================================================

function getInitData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  // Ensure schema is up to date (idempotent — no-op if already migrated)
  setupLogSheet_(ss);
  setupExerciseSheet_(ss);

  const fullProgram = getFullProgram_(ss);
  return JSON.stringify({
    cycleInfo:           getCycleInfo_(ss),
    exercises:           getExercises_(ss),
    fullProgram:         fullProgram,
    allLoggedExercises:  getLoggedExerciseNames_(ss),
    loggedDates:         getLoggedDates_(ss),
    programExists:       fullProgram.length > 0,
    scheduleHistory:     getScheduleHistory_(),
  });
}

function getSessionSetup(dayLabel, dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const programNames = getFullProgram_(ss)
    .filter(p => p.dayLabel === dayLabel)
    .map(p => p.exerciseName);

  const existingResult = getExistingLogForDate_(ss, dateStr);
  const loggedNames = Object.keys(existingResult.sets);
  const allNames = [...new Set([...programNames, ...loggedNames])];
  const lastSession = getLastSessionData_(ss, allNames, dateStr);

  return JSON.stringify({
    lastSession,
    existingLog: existingResult.sets,
    existingNotes: existingResult.notes,
  });
}

// =============================================================================
// Create New Exercise (inline)
// =============================================================================

function createExercise(exData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_EXERCISES);
  if (!sheet) return JSON.stringify({ success: false, message: 'Exercise sheet missing.' });

  const name = String(exData.name || '').trim();
  if (!name) return JSON.stringify({ success: false, message: 'Name is required.' });

  // Check uniqueness
  if (sheet.getLastRow() >= 2) {
    const existing = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat()
      .map(n => String(n).trim().toLowerCase());
    if (existing.includes(name.toLowerCase())) {
      return JSON.stringify({ success: false, message: 'An exercise with that name already exists.' });
    }
  }

  const row = [
    name,
    String(exData.muscleGroup || '').trim(),
    String(exData.equipment || '').trim(),
    String(exData.notes || '').trim(),
    String(exData.measurementType || 'weight_reps').trim(),
    String(exData.cardioModality || '').trim(),
  ];
  const nextRow = sheet.getLastRow() + 1;
  sheet.getRange(nextRow, 1, 1, EX_COLS).setValues([row]);

  return JSON.stringify({
    success: true,
    exercise: {
      name: row[0], muscleGroup: row[1], equipment: row[2],
      notes: row[3], measurementType: row[4], cardioModality: row[5],
    },
  });
}

// =============================================================================
// Program Save (editor)
// =============================================================================

function saveProgramSetup(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  archiveCurrentSchedule_(ss, today);

  // Day Schedule sheet
  let schedSheet = ss.getSheetByName(SHEET_SCHEDULE);
  if (!schedSheet) schedSheet = ss.insertSheet(SHEET_SCHEDULE);
  schedSheet.clear();
  schedSheet.getRange(1, 1, 1, 3).setValues([['Day', 'Type', 'Label']]);
  const schedRows = payload.schedule.map(d => [d.day, d.type, d.label]);
  if (schedRows.length > 0) schedSheet.getRange(2, 1, schedRows.length, 3).setValues(schedRows);
  schedSheet.setFrozenRows(1);

  // Program sheet
  const progSheet = ss.getSheetByName(SHEET_PROGRAM);
  progSheet.clear();
  const headers = ['Day Label', 'Exercise Name', 'Target Sets', 'Target Reps', 'Target RPE', 'Exercise Order'];
  progSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  const progRows = payload.exercises.map(ex =>
    [ex.dayLabel, ex.name, ex.sets, ex.reps, ex.rpe || '', ex.order]);
  if (progRows.length > 0) {
    progSheet.getRange(2, 1, progRows.length, headers.length).setValues(progRows);
    progSheet.getRange(2, 4, progRows.length, 1).setNumberFormat('@');
  }
  progSheet.setFrozenRows(1);

  // Ensure startDate set
  const settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  if (settingsSheet) {
    const startVal = settingsSheet.getRange(2, 2).getValue();
    if (!startVal || String(startVal).trim() === '') {
      settingsSheet.getRange(2, 2).setValue(today);
    }
  }

  return JSON.stringify({ success: true });
}

// =============================================================================
// Schedule Archive
// =============================================================================

function archiveCurrentSchedule_(ss, endDate) {
  const currentSchedule = getDaySchedule();
  const currentProgram  = getFullProgram_(ss);
  if (!currentSchedule || currentSchedule.length === 0) return;

  const schedSheet = ss.getSheetByName(SHEET_SCHEDULE);
  if ((!schedSheet || schedSheet.getLastRow() < 2) && currentProgram.length === 0) return;

  const props = PropertiesService.getScriptProperties();
  let history = [];
  try { const raw = props.getProperty('scheduleHistory'); if (raw) history = JSON.parse(raw); } catch (_) {}

  let startDate = '';
  if (history.length > 0) {
    const lastEnd = history[history.length - 1].endDate;
    const d = new Date(lastEnd + 'T00:00:00');
    d.setDate(d.getDate() + 1);
    startDate = d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
  } else {
    const settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
    if (settingsSheet && settingsSheet.getLastRow() >= 2) {
      const raw = settingsSheet.getRange(2, 2).getValue();
      if (raw instanceof Date) startDate = Utilities.formatDate(raw, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      else if (raw) startDate = String(raw).trim();
    }
  }
  if (!startDate) return;

  const endD = new Date(endDate + 'T00:00:00');
  endD.setDate(endD.getDate() - 1);
  const archiveEnd = endD.getFullYear() + '-' + String(endD.getMonth() + 1).padStart(2, '0') + '-' + String(endD.getDate()).padStart(2, '0');
  if (startDate > archiveEnd) return;

  history.push({ startDate, endDate: archiveEnd, schedule: currentSchedule });
  props.setProperty('scheduleHistory', JSON.stringify(history));
}

function getScheduleHistory_() {
  try { const raw = PropertiesService.getScriptProperties().getProperty('scheduleHistory'); if (raw) return JSON.parse(raw); } catch (_) {}
  return [];
}

// =============================================================================
// Last Working Sets
// =============================================================================

function getLastWorkingSets(exerciseName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet || sheet.getLastRow() < 2) return JSON.stringify({ sets: [], date: null });

  const tz = Session.getScriptTimeZone();
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();

  const rows = data
    .filter(r => r[3] === exerciseName)
    .map(r => {
      const dateStr = r[0] instanceof Date
        ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd')
        : String(r[0]).trim();
      return readLogRow_(r, dateStr, lastCol);
    })
    .sort((a, b) => b.date.localeCompare(a.date));

  if (rows.length === 0) return JSON.stringify({ sets: [], date: null });

  const latestDate = rows[0].date;
  const sets = rows
    .filter(r => r.date === latestDate)
    .sort((a, b) => a.setNumber - b.setNumber)
    .map(r => ({
      setNumber: r.setNumber, weight: r.weight, reps: r.reps,
      minutes: r.minutes, seconds: r.seconds, distance: r.distance, rpe: r.rpe,
    }));

  const note = rows.find(r => r.date === latestDate && r.notes) || {};
  return JSON.stringify({ sets, date: latestDate, notes: note.notes || '' });
}

/** Read a log row into a structured object, handling old + new schemas. */
function readLogRow_(r, dateStr, lastCol) {
  if (lastCol >= 13) {
    return {
      date: dateStr, setNumber: r[4], weight: r[5], reps: r[6],
      minutes: r[7], seconds: r[8], distance: r[9], rpe: r[10], notes: r[11] || '',
    };
  }
  // Old 9-col schema: Date(0), DayLbl, DayTyp, Ex, Set#(4), Weight(5), Reps(6), RPE(7), Timestamp(8)
  return {
    date: dateStr, setNumber: r[4], weight: r[5], reps: r[6],
    minutes: '', seconds: '', distance: '', rpe: r[7], notes: '',
  };
}

// =============================================================================
// Data Readers
// =============================================================================

function getCycleInfo_(ss) {
  const tz = Session.getScriptTimeZone();
  const settingsSheet = ss.getSheetByName(SHEET_SETTINGS);
  let startDateStr = '';

  if (settingsSheet && settingsSheet.getLastRow() >= 2) {
    const raw = settingsSheet.getRange(2, 2).getValue();
    if (raw instanceof Date) startDateStr = Utilities.formatDate(raw, tz, 'yyyy-MM-dd');
    else if (raw) {
      const parsed = new Date(String(raw).trim());
      if (!isNaN(parsed.getTime())) startDateStr = Utilities.formatDate(parsed, tz, 'yyyy-MM-dd');
    }
  }

  const today = new Date();
  today.setHours(0, 0, 0, 0);
  if (!startDateStr) {
    startDateStr = Utilities.formatDate(today, tz, 'yyyy-MM-dd');
    if (settingsSheet) settingsSheet.getRange(2, 2).setValue(startDateStr);
  }

  const startDate = new Date(startDateStr + 'T00:00:00');
  const daysElapsed = Math.floor((today - startDate) / 86400000);
  const schedule = getDaySchedule();
  const cycleLength = schedule.length;
  const safeDays = Math.max(daysElapsed, 0);

  return {
    startDate: startDateStr,
    currentCycleNumber: Math.floor(safeDays / cycleLength) + 1,
    currentDayNumber: (safeDays % cycleLength) + 1,
    todayStr: Utilities.formatDate(today, tz, 'yyyy-MM-dd'),
    daySchedule: schedule,
  };
}

function getExercises_(ss) {
  const sheet = ss.getSheetByName(SHEET_EXERCISES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const lastCol = Math.max(sheet.getLastColumn(), 4);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  return data.map(r => ({
    name: r[0],
    muscleGroup: r[1],
    equipment: r[2],
    notes: r[3] || '',
    measurementType: (r[4] && String(r[4]).trim()) || 'weight_reps',
    cardioModality: r[5] ? String(r[5]).trim() : '',
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
    targetReps: r[3] instanceof Date
      ? (r[3].getMonth() + 1) + '-' + r[3].getDate()
      : String(r[3]),
    targetRPE: r[4] || '',
    order: r[5] || 0,
  }));
}

function getLoggedDates_(ss) {
  const sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet || sheet.getLastRow() < 2) return {};
  const tz = Session.getScriptTimeZone();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 3).getValues();
  const result = {};
  for (const row of data) {
    const dateStr = row[0] instanceof Date
      ? Utilities.formatDate(row[0], tz, 'yyyy-MM-dd') : String(row[0]).trim();
    if (!result[dateStr]) result[dateStr] = { dayLabel: String(row[1]).trim(), dayType: String(row[2]).trim() };
  }
  return result;
}

function getExistingLogForDate_(ss, dateStr) {
  const sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet || sheet.getLastRow() < 2) return { sets: {}, notes: {} };

  const tz = Session.getScriptTimeZone();
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  const sets = {}, notes = {};

  for (const r of data) {
    const rowDate = r[0] instanceof Date
      ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') : String(r[0]).trim();
    if (rowDate !== dateStr) continue;

    const exName = r[3];
    if (!sets[exName]) sets[exName] = [];
    const parsed = readLogRow_(r, rowDate, lastCol);
    sets[exName].push({
      setNumber: parsed.setNumber, weight: parsed.weight, reps: parsed.reps,
      minutes: parsed.minutes, seconds: parsed.seconds, distance: parsed.distance, rpe: parsed.rpe,
    });
    if (parsed.notes && !notes[exName]) notes[exName] = parsed.notes;
  }

  for (const name of Object.keys(sets)) sets[name].sort((a, b) => a.setNumber - b.setNumber);
  return { sets, notes };
}

function getLastSessionData_(ss, exerciseNames, excludeDateStr) {
  if (!exerciseNames || !exerciseNames.length) return {};
  const sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet || sheet.getLastRow() < 2) return {};

  const tz = Session.getScriptTimeZone();
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  const result = {};

  for (const name of exerciseNames) {
    const rows = data
      .filter(r => {
        if (r[3] !== name) return false;
        const rowDate = r[0] instanceof Date
          ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') : String(r[0]).trim();
        return rowDate !== excludeDateStr;
      })
      .sort((a, b) => {
        const dc = new Date(b[0]) - new Date(a[0]);
        return dc !== 0 ? dc : a[4] - b[4];
      });

    if (rows.length === 0) { result[name] = []; continue; }

    const latestDateStr = rows[0][0] instanceof Date
      ? Utilities.formatDate(rows[0][0], tz, 'yyyy-MM-dd') : String(rows[0][0]).trim();

    result[name] = rows
      .filter(r => {
        const rd = r[0] instanceof Date
          ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') : String(r[0]).trim();
        return rd === latestDateStr;
      })
      .map(r => {
        const parsed = readLogRow_(r, latestDateStr, lastCol);
        return {
          setNumber: parsed.setNumber, weight: parsed.weight, reps: parsed.reps,
          minutes: parsed.minutes, seconds: parsed.seconds, distance: parsed.distance, rpe: parsed.rpe,
        };
      });
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
// Save Session
// =============================================================================

function saveWorkout(sessionData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG);
  setupLogSheet_(ss); // ensure migrated
  const tz = Session.getScriptTimeZone();
  const timestamp = new Date();

  if (sheet.getLastRow() >= 2) {
    const dates = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    for (let i = dates.length - 1; i >= 0; i--) {
      const rowDate = dates[i] instanceof Date
        ? Utilities.formatDate(dates[i], tz, 'yyyy-MM-dd') : String(dates[i]).trim();
      if (rowDate === sessionData.date) sheet.deleteRow(i + 2);
    }
  }

  const notes = sessionData.exerciseNotes || {};
  const rows = sessionData.sets.map(s => [
    sessionData.date, sessionData.dayLabel, sessionData.dayType,
    s.exercise, s.setNumber,
    s.weight !== '' && s.weight !== undefined ? s.weight : '',
    s.reps !== '' && s.reps !== undefined ? s.reps : '',
    s.minutes !== '' && s.minutes !== undefined ? s.minutes : '',
    s.seconds !== '' && s.seconds !== undefined ? s.seconds : '',
    s.distance !== '' && s.distance !== undefined ? s.distance : '',
    s.rpe !== '' && s.rpe !== undefined ? s.rpe : '',
    notes[s.exercise] || '',
    timestamp,
  ]);

  if (rows.length === 0) return JSON.stringify({ success: false, message: 'No sets to save.' });

  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, rows.length, LOG_COLS).setValues(rows);
  return JSON.stringify({ success: true, rowsWritten: rows.length });
}

// =============================================================================
// Delete
// =============================================================================

function deleteWorkout(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG);
  const tz = Session.getScriptTimeZone();
  if (!sheet || sheet.getLastRow() < 2) return JSON.stringify({ success: false, message: 'No log data.' });

  const dates = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
  let deleted = 0;
  for (let i = dates.length - 1; i >= 0; i--) {
    const rowDate = dates[i] instanceof Date
      ? Utilities.formatDate(dates[i], tz, 'yyyy-MM-dd') : String(dates[i]).trim();
    if (rowDate === dateStr) { sheet.deleteRow(i + 2); deleted++; }
  }
  return JSON.stringify({ success: true, rowsDeleted: deleted });
}

// =============================================================================
// Body Log (unchanged)
// =============================================================================

function getBodySettings_(ss) {
  const sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet || sheet.getLastRow() < 2) return {};
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const map = {};
  for (const [k, v] of data) map[String(k).trim()] = v;
  return {
    heightCm: Number(map.heightCm) || 0,
    gender: String(map.gender || 'male').trim().toLowerCase(),
    birthYear: Number(map.birthYear) || 0,
    activityLevel: String(map.activityLevel || 'moderate').trim().toLowerCase(),
  };
}

function saveBodyEntry(entry) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_BODY);
  const tz = Session.getScriptTimeZone();
  const timestamp = new Date();

  if (sheet.getLastRow() >= 2) {
    const dates = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat();
    for (let i = dates.length - 1; i >= 0; i--) {
      const rowDate = dates[i] instanceof Date
        ? Utilities.formatDate(dates[i], tz, 'yyyy-MM-dd') : String(dates[i]).trim();
      if (rowDate === entry.date) sheet.deleteRow(i + 2);
    }
  }

  const row = [entry.date, entry.weight || '', entry.waist || '', entry.neck || '',
    entry.hips || '', entry.chest || '', entry.bicepL || '', entry.bicepR || '',
    entry.quadL || '', entry.quadR || '', entry.calves || '', entry.shoulders || '', timestamp];
  const startRow = sheet.getLastRow() + 1;
  sheet.getRange(startRow, 1, 1, row.length).setValues([row]);
  return JSON.stringify({ success: true });
}

function getBodyLogData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_BODY);
  const tz = Session.getScriptTimeZone();
  const settings = getBodySettings_(ss);
  if (!sheet || sheet.getLastRow() < 2) return JSON.stringify({ entries: [], settings });

  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 12).getValues();
  const entries = data.map(r => ({
    date: r[0] instanceof Date ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') : String(r[0]).trim(),
    weight: r[1] !== '' ? Number(r[1]) : null,
    waist: r[2] !== '' ? Number(r[2]) : null,
    neck: r[3] !== '' ? Number(r[3]) : null,
    hips: r[4] !== '' ? Number(r[4]) : null,
    chest: r[5] !== '' ? Number(r[5]) : null,
    bicepL: r[6] !== '' ? Number(r[6]) : null,
    bicepR: r[7] !== '' ? Number(r[7]) : null,
    quadL: r[8] !== '' ? Number(r[8]) : null,
    quadR: r[9] !== '' ? Number(r[9]) : null,
    calves: r[10] !== '' ? Number(r[10]) : null,
    shoulders: r[11] !== '' ? Number(r[11]) : null,
  })).sort((a, b) => a.date.localeCompare(b.date));
  return JSON.stringify({ entries, settings });
}

// =============================================================================
// Progress Data — supports all measurement types
// =============================================================================

function getProgressData(exerciseName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet || sheet.getLastRow() < 2) return JSON.stringify({ sessions: [], measurementType: 'weight_reps' });

  const exercises = getExercises_(ss);
  const ex = exercises.find(e => e.name === exerciseName);
  const mt = ex ? ex.measurementType : 'weight_reps';

  const tz = Session.getScriptTimeZone();
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  const rows = data.filter(r => r[3] === exerciseName);
  if (rows.length === 0) return JSON.stringify({ sessions: [], measurementType: mt });

  const byDate = {};
  for (const r of rows) {
    const dateStr = r[0] instanceof Date
      ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') : String(r[0]).trim();
    if (!byDate[dateStr]) byDate[dateStr] = [];
    byDate[dateStr].push(readLogRow_(r, dateStr, lastCol));
  }

  const sessions = [];
  let bestE1RM = 0, bestDate = '';

  for (const [date, sets] of Object.entries(byDate).sort()) {
    const sorted = sets.slice().sort((a, b) => a.setNumber - b.setNumber);
    const first = sorted[0];
    const session = { date };

    if (mt === 'weight_reps') {
      let maxE1RM = 0, totalVolume = 0;
      for (const s of sorted) {
        const w = Number(s.weight) || 0, r = Number(s.reps) || 0;
        totalVolume += w * r;
        const e1rm = r > 0 ? w * (1 + r / 30) : 0;
        if (e1rm > maxE1RM) maxE1RM = e1rm;
      }
      if (maxE1RM > bestE1RM) { bestE1RM = maxE1RM; bestDate = date; }
      session.e1rm = Math.round(maxE1RM * 10) / 10;
      session.volume = Math.round(totalVolume);
      session.firstSetWeight = Number(first.weight) || 0;
    } else if (mt === 'weight_time') {
      const totalSec = sorted.reduce((sum, s) =>
        sum + (Number(s.minutes) || 0) * 60 + (Number(s.seconds) || 0), 0);
      session.firstSetWeight = Number(first.weight) || 0;
      session.totalDurationMin = Math.round(totalSec / 6) / 10;
      session.maxWeight = Math.max(...sorted.map(s => Number(s.weight) || 0));
    } else if (mt === 'time_only') {
      const firstSec = (Number(first.minutes) || 0) * 60 + (Number(first.seconds) || 0);
      const maxSec = Math.max(...sorted.map(s =>
        (Number(s.minutes) || 0) * 60 + (Number(s.seconds) || 0)));
      session.firstSetDurationMin = Math.round(firstSec / 6) / 10;
      session.maxDurationMin = Math.round(maxSec / 6) / 10;
    } else if (mt === 'cardio') {
      const totalDist = sorted.reduce((sum, s) => sum + (Number(s.distance) || 0), 0);
      const totalMin = sorted.reduce((sum, s) => sum + (Number(s.minutes) || 0), 0);
      session.distance = Math.round(totalDist * 100) / 100;
      session.durationMin = totalMin;
    }
    sessions.push(session);
  }

  let volumeChange = null;
  if (mt === 'weight_reps' && sessions.length >= 2) {
    const curr = sessions[sessions.length - 1].volume;
    const prev = sessions[sessions.length - 2].volume;
    volumeChange = prev > 0 ? Math.round(((curr - prev) / prev) * 1000) / 10 : null;
  }

  return JSON.stringify({
    measurementType: mt,
    sessions,
    bestE1RM: { value: Math.round(bestE1RM * 10) / 10, date: bestDate },
    totalSessions: sessions.length,
    volumeChange,
  });
}

// =============================================================================
// Draft
// =============================================================================

function saveDraft(draftJson) { PropertiesService.getScriptProperties().setProperty('workoutDraft', draftJson); }
function loadDraft() { return PropertiesService.getScriptProperties().getProperty('workoutDraft') || null; }
function clearDraft() { PropertiesService.getScriptProperties().deleteProperty('workoutDraft'); }// =============================================================================

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
