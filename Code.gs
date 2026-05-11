// =============================================================================
// Code.gs — Workout Log Tool (Server-Side)
//
// Architecture:
//   - Persistent program: a single 7-day weekly template stored with
//     block_id = 0 in Program and Day Schedule sheets. Always exists.
//   - Blocks: optional periodized overlays. When a block covers a date, its
//     schedule and exercises take priority over the persistent program.
//   - Global week numbering: weeks indexed from the persistent program's
//     start date (in Settings).
// =============================================================================

const SHEET_LOG       = 'Workout Log';
const SHEET_EXERCISES = 'Exercise Selection';
const SHEET_PROGRAM   = 'Program';
const SHEET_SETTINGS  = 'Settings';
const SHEET_BODY      = 'Body Log';
const SHEET_SCHEDULE  = 'Day Schedule';
const SHEET_BLOCKS    = 'Blocks';

const PERSISTENT_BLOCK_ID = 0; // Sentinel for the persistent program.

const LOG_HEADERS = [
  'Date', 'Day Label', 'Day Type', 'Exercise Name', 'Set Number',
  'Weight (kg)', 'Reps', 'Minutes', 'Seconds', 'Distance (km)',
  'RPE', 'Notes', 'Timestamp'
];
const LOG_COLS = LOG_HEADERS.length;

const EX_HEADERS = [
  'Exercise Name', 'Primary Muscle Group', 'Equipment', 'Notes',
  'Measurement Type', 'Cardio Modality'
];
const EX_COLS = EX_HEADERS.length;

const PROGRAM_HEADERS = [
  'Block ID', 'Week Number', 'Day Label', 'Exercise Name',
  'Target Sets', 'Target Reps', 'Target RPE', 'Exercise Order', 'Superset Group'
];
const PROGRAM_COLS = PROGRAM_HEADERS.length;

const SCHEDULE_HEADERS = ['Block ID', 'Week Number', 'Day', 'Type', 'Label'];
const SCHEDULE_COLS = SCHEDULE_HEADERS.length;

const BLOCK_HEADERS = ['Block ID', 'Name', 'Length Weeks', 'Start Date', 'End Date', 'Status'];
const BLOCK_COLS = BLOCK_HEADERS.length;

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
// Setup / Migration
// =============================================================================

function setupAllSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupExerciseSheet_(ss);
  setupLogSheet_(ss);
  setupBlocksSheet_(ss);
  setupProgramSheet_(ss);
  setupDayScheduleSheet_(ss);
  setupSettingsSheet_(ss);
  setupBodyLogSheet_(ss);
  migrateLegacyBlock1ToPersistent_(ss);
  SpreadsheetApp.getUi().alert('All sheets created and migrated.');
}

function setupExerciseSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_EXERCISES);
  if (!sheet) sheet = ss.insertSheet(SHEET_EXERCISES);
  if (sheet.getLastRow() > 0) {
    const currentCols = sheet.getLastColumn();
    if (currentCols < EX_COLS) {
      sheet.insertColumnsAfter(4, EX_COLS - currentCols);
      sheet.getRange(1, 1, 1, EX_COLS).setValues([EX_HEADERS]);
      if (sheet.getLastRow() >= 2) {
        const rows = sheet.getLastRow() - 1;
        sheet.getRange(2, 5, rows, 2).setValues(Array(rows).fill(['weight_reps', '']));
      }
    }
    return;
  }
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
  const currentCols = sheet.getLastColumn();
  if (currentCols < LOG_COLS) {
    sheet.insertColumnsAfter(7, 3);
    sheet.insertColumnsAfter(11, 1);
    sheet.getRange(1, 1, 1, LOG_COLS).setValues([LOG_HEADERS]);
  }
}

function setupBlocksSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_BLOCKS);
  if (!sheet) sheet = ss.insertSheet(SHEET_BLOCKS);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, BLOCK_COLS).setValues([BLOCK_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, BLOCK_COLS);
  }
}

function setupProgramSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_PROGRAM);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_PROGRAM);
    sheet.getRange(1, 1, 1, PROGRAM_COLS).setValues([PROGRAM_HEADERS]);
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, PROGRAM_COLS);
    return;
  }
  const currentCols = sheet.getLastColumn();
  if (currentCols === 6) {
    sheet.insertColumnsBefore(1, 2);
    if (sheet.getLastRow() >= 2) {
      const rows = sheet.getLastRow() - 1;
      sheet.getRange(2, 1, rows, 2).setValues(Array(rows).fill([1, 1]));
      sheet.getRange(2, 6, rows, 1).setNumberFormat('@');
    }
  }
  sheet.getRange(1, 1, 1, PROGRAM_COLS).setValues([PROGRAM_HEADERS]);
}

function setupDayScheduleSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_SCHEDULE);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_SCHEDULE);
    sheet.getRange(1, 1, 1, SCHEDULE_COLS).setValues([SCHEDULE_HEADERS]);
    sheet.setFrozenRows(1);
    return;
  }
  if (sheet.getLastColumn() < SCHEDULE_COLS) {
    sheet.insertColumnsBefore(1, 2);
    sheet.getRange(1, 1, 1, SCHEDULE_COLS).setValues([SCHEDULE_HEADERS]);
    if (sheet.getLastRow() >= 2) {
      const rows = sheet.getLastRow() - 1;
      sheet.getRange(2, 1, rows, 2).setValues(Array(rows).fill([1, 1]));
    }
  }
}

function setupSettingsSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet) sheet = ss.insertSheet(SHEET_SETTINGS);
  if (sheet.getLastRow() >= 2) {
    const keys = sheet.getRange(1, 1, sheet.getLastRow(), 1).getValues().flat();
    const needed = [['heightCm', ''], ['gender', 'male'], ['birthYear', ''],
      ['activityLevel', 'moderate'], ['supersetsEnabled', 'false'], ['theme', 'blue']];
    for (const [key, def] of needed) {
      if (!keys.includes(key)) {
        const nr = sheet.getLastRow() + 1;
        sheet.getRange(nr, 1, 1, 2).setValues([[key, def]]);
      }
    }
    return;
  }
  sheet.clear();
    sheet.getRange(1, 1, 8, 2).setValues([
      ['Key', 'Value'], ['startDate', ''], ['heightCm', ''],
      ['gender', 'male'], ['birthYear', ''], ['activityLevel', 'moderate'],
      ['supersetsEnabled', 'false'], ['theme', 'blue'],
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

/**
 * If the only block is a legacy auto-created Block 1 (lengthWeeks=1, active),
 * convert its rows into the persistent program (block_id = 0) and remove the
 * Block 1 metadata. This runs at most once per spreadsheet.
 */
function migrateLegacyBlock1ToPersistent_(ss) {
  const blocksSheet = ss.getSheetByName(SHEET_BLOCKS);
  if (!blocksSheet || blocksSheet.getLastRow() < 2) return;

  const allBlocks = getAllBlocks_(ss);
  if (allBlocks.length !== 1) return;
  const b = allBlocks[0];
  if (b.id !== 1 || b.lengthWeeks !== 1 || b.status !== 'active') return;

  // Migrate Day Schedule rows: change block_id from 1 to 0.
  const sched = ss.getSheetByName(SHEET_SCHEDULE);
  if (sched && sched.getLastRow() >= 2) {
    const data = sched.getRange(2, 1, sched.getLastRow() - 1, SCHEDULE_COLS).getValues();
    for (let i = 0; i < data.length; i++) {
      if (Number(data[i][0]) === 1) {
        sched.getRange(i + 2, 1).setValue(PERSISTENT_BLOCK_ID);
      }
    }
  }

  // Migrate Program rows similarly.
  const prog = ss.getSheetByName(SHEET_PROGRAM);
  if (prog && prog.getLastRow() >= 2) {
    const data = prog.getRange(2, 1, prog.getLastRow() - 1, PROGRAM_COLS).getValues();
    for (let i = 0; i < data.length; i++) {
      if (Number(data[i][0]) === 1) {
        prog.getRange(i + 2, 1).setValue(PERSISTENT_BLOCK_ID);
      }
    }
  }

  // Remove the Block 1 metadata row.
  const rows = blocksSheet.getRange(2, 1, blocksSheet.getLastRow() - 1, 1).getValues();
  for (let i = rows.length - 1; i >= 0; i--) {
    if (Number(rows[i][0]) === 1) blocksSheet.deleteRow(i + 2);
  }
}

// =============================================================================
// Block Reads
// =============================================================================

function getAllBlocks_(ss) {
  const sheet = ss.getSheetByName(SHEET_BLOCKS);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const tz = Session.getScriptTimeZone();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, BLOCK_COLS).getValues();
  return data
    .filter(r => r[0] !== '' && r[0] !== null)
    .map(r => ({
      id: Number(r[0]),
      name: String(r[1] || '').trim() || ('Block ' + r[0]),
      lengthWeeks: Number(r[2]) || 1,
      startDate: r[3] instanceof Date ? Utilities.formatDate(r[3], tz, 'yyyy-MM-dd') : String(r[3] || '').trim(),
      endDate: r[4] instanceof Date ? Utilities.formatDate(r[4], tz, 'yyyy-MM-dd') : String(r[4] || '').trim(),
      status: String(r[5] || 'active').trim().toLowerCase(),
    }))
    .sort((a, b) => a.id - b.id);
}

function getActiveBlock_(ss) {
  const blocks = getAllBlocks_(ss);
  const active = blocks.filter(b => b.status === 'active');
  if (active.length === 0) return null;
  return active.sort((a, b) => b.id - a.id)[0];
}

/** Read schedule rows for a given blockId (use 0 for persistent program). */
function getScheduleForBlock_(ss, blockId, lengthWeeks) {
  const sheet = ss.getSheetByName(SHEET_SCHEDULE);
  const weeks = [];
  for (let w = 1; w <= lengthWeeks; w++) weeks.push({ weekNumber: w, schedule: [] });
  if (!sheet || sheet.getLastRow() < 2) return weeks;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, SCHEDULE_COLS).getValues();
  for (const r of data) {
    if (Number(r[0]) !== blockId) continue;
    const w = Number(r[1]);
    if (w < 1 || w > lengthWeeks) continue;
    weeks[w - 1].schedule.push({
      day: Number(r[2]),
      type: String(r[3] || '').trim(),
      label: String(r[4] || '').trim() || ('Day ' + r[2]),
    });
  }
  for (const w of weeks) w.schedule.sort((a, b) => a.day - b.day);
  return weeks;
}

/** Read program rows for a given blockId. Returns { 'weekNum': { dayLabel: [exercises] } }. */
function getProgramForBlock_(ss, blockId) {
  const sheet = ss.getSheetByName(SHEET_PROGRAM);
  const program = {};
  if (!sheet || sheet.getLastRow() < 2) return program;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, PROGRAM_COLS).getValues();
  for (const r of data) {
    if (Number(r[0]) !== blockId) continue;
    const week = String(Number(r[1]));
    const dayLabel = String(r[2] || '').trim();
    if (!program[week]) program[week] = {};
    if (!program[week][dayLabel]) program[week][dayLabel] = [];
    program[week][dayLabel].push({
      exerciseName: String(r[3] || '').trim(),
      targetSets: r[4],
      targetReps: r[5] instanceof Date
        ? (r[5].getMonth() + 1) + '-' + r[5].getDate()
        : String(r[5] || ''),
      targetRPE: r[6] || '',
      order: Number(r[7]) || 0,
      group: String(r[8] || '').trim(),
    });
  }
  for (const w of Object.keys(program)) {
    for (const d of Object.keys(program[w])) {
      program[w][d].sort((a, b) => a.order - b.order);
    }
  }
  return program;
}

function hydrateBlock_(ss, block) {
  return {
    id: block.id, name: block.name, lengthWeeks: block.lengthWeeks,
    startDate: block.startDate, endDate: block.endDate, status: block.status,
    weeks: getScheduleForBlock_(ss, block.id, block.lengthWeeks),
    program: getProgramForBlock_(ss, block.id),
  };
}

function getPersistentProgram_(ss) {
  // Persistent program is always "1 week long" — same week template repeats.
  const weeks = getScheduleForBlock_(ss, PERSISTENT_BLOCK_ID, 1);
  return {
    schedule: weeks[0] ? weeks[0].schedule : [],
    program: getProgramForBlock_(ss, PERSISTENT_BLOCK_ID),
  };
}

/** Returns null or the resolved block context. Active block wins; among
 * completed blocks where startDate <= date <= endDate, highest ID wins. */
function resolveBlockContext_(ss, dateStr) {
  const blocks = getAllBlocks_(ss);
  let chosen = null;
  for (const b of blocks) {
    if (b.status !== 'active' || !b.startDate) continue;
    if (dateStr < b.startDate) continue;
    const days = daysBetween_(b.startDate, dateStr);
    if (days >= b.lengthWeeks * 7) continue;
    chosen = b;
    break;
  }
  if (!chosen) {
    const candidates = blocks.filter(b =>
      b.status !== 'active' && b.startDate && b.endDate &&
      dateStr >= b.startDate && dateStr <= b.endDate);
    if (candidates.length > 0) chosen = candidates.sort((a, b) => b.id - a.id)[0];
  }
  if (!chosen) return null;
  const days = daysBetween_(chosen.startDate, dateStr);
  const weekInBlock = Math.floor(days / 7) + 1;
  const dayInWeek = (days % 7) + 1;
  const hydrated = hydrateBlock_(ss, chosen);
  const week = hydrated.weeks.find(w => w.weekNumber === weekInBlock) || { schedule: [] };
  const dayEntry = week.schedule.find(s => s.day === dayInWeek) || null;
  const dayLabel = dayEntry ? dayEntry.label : null;
  const dayType = dayEntry ? dayEntry.type : null;
  const programForDay = (dayLabel && hydrated.program[String(weekInBlock)] && hydrated.program[String(weekInBlock)][dayLabel])
    || [];
  return { block: hydrated, weekInBlock, dayInWeek, dayLabel, dayType, programForDay };
}

function daysBetween_(startStr, endStr) {
  const sParts = startStr.split('-').map(Number);
  const eParts = endStr.split('-').map(Number);
  const s = Date.UTC(sParts[0], sParts[1] - 1, sParts[2]);
  const e = Date.UTC(eParts[0], eParts[1] - 1, eParts[2]);
  return Math.floor((e - s) / 86400000);
}

// =============================================================================
// Persistent Program Start Date
// =============================================================================

function getPersistentStartDate_(ss) {
  const settings = ss.getSheetByName(SHEET_SETTINGS);
  if (!settings || settings.getLastRow() < 2) return null;
  const tz = Session.getScriptTimeZone();
  const raw = settings.getRange(2, 2).getValue();
  if (raw instanceof Date) return Utilities.formatDate(raw, tz, 'yyyy-MM-dd');
  if (raw) {
    const s = String(raw).trim();
    if (/^\d{4}-\d{2}-\d{2}$/.test(s)) return s;
    const parsed = new Date(s);
    if (!isNaN(parsed.getTime())) return Utilities.formatDate(parsed, tz, 'yyyy-MM-dd');
  }
  return null;
}

function setPersistentStartDate_(ss, dateStr) {
  const settings = ss.getSheetByName(SHEET_SETTINGS);
  if (settings) settings.getRange(2, 2).setValue(dateStr);
}

// =============================================================================
// Block Writes
// =============================================================================

function nextBlockId_(ss) {
  const blocks = getAllBlocks_(ss);
  if (blocks.length === 0) return 2; // Start at 2 — id 0 is persistent, and we keep 1 reserved historically.
  return Math.max(...blocks.map(b => b.id), 1) + 1;
}

function saveBlock(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupBlocksSheet_(ss);
  setupProgramSheet_(ss);
  setupDayScheduleSheet_(ss);

  const isNew = !payload.id;
  let blockId = payload.id;
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  for (const w of payload.weeks) {
    const workoutLabels = (w.schedule || [])
      .filter(d => d.type !== 'Rest')
      .map(d => String(d.label || '').trim());
    const seen = new Set();
    for (const l of workoutLabels) {
      if (seen.has(l)) {
        return JSON.stringify({ success: false, message:
          'In Week ' + w.weekNumber + ', workout days must have unique labels (duplicate: "' + l + '").' });
      }
      seen.add(l);
    }
  }

  if (isNew) {
    const active = getActiveBlock_(ss);
    if (active) endBlockById_(ss, active.id, today);
    blockId = nextBlockId_(ss);
    const startDate = payload.startDate || today;
    const blocksSheet = ss.getSheetByName(SHEET_BLOCKS);
    const nextRow = blocksSheet.getLastRow() + 1;
    blocksSheet.getRange(nextRow, 1, 1, BLOCK_COLS).setValues([
      [blockId, payload.name || ('Block ' + blockId), payload.lengthWeeks, startDate, '', 'active']
    ]);
  } else {
    const blocksSheet = ss.getSheetByName(SHEET_BLOCKS);
    const rows = blocksSheet.getRange(2, 1, blocksSheet.getLastRow() - 1, BLOCK_COLS).getValues();
    for (let i = 0; i < rows.length; i++) {
      if (Number(rows[i][0]) === blockId) {
        blocksSheet.getRange(i + 2, 2).setValue(payload.name || rows[i][1]);
        blocksSheet.getRange(i + 2, 3).setValue(payload.lengthWeeks);
        break;
      }
    }
  }
  rewriteRowsForBlock_(ss, SHEET_SCHEDULE, blockId, buildScheduleRows_(blockId, payload.weeks));
  rewriteRowsForBlock_(ss, SHEET_PROGRAM, blockId, buildProgramRows_(blockId, payload.weeks));
  return JSON.stringify({ success: true, blockId });
}

/**
 * Save the persistent program. payload = { schedule: [{day,type,label}], exercises: { dayLabel: [...] } }
 */
function savePersistentProgram(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupProgramSheet_(ss);
  setupDayScheduleSheet_(ss);

  const workoutLabels = (payload.schedule || [])
    .filter(d => d.type !== 'Rest')
    .map(d => String(d.label || '').trim());
  const seen = new Set();
  for (const l of workoutLabels) {
    if (seen.has(l)) {
      return JSON.stringify({ success: false, message: 'Workout days must have unique labels (duplicate: "' + l + '").' });
    }
    seen.add(l);
  }

  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  if (!getPersistentStartDate_(ss)) setPersistentStartDate_(ss, today);

  const weeks = [{ weekNumber: 1, schedule: payload.schedule || [], exercises: payload.exercises || {} }];
  rewriteRowsForBlock_(ss, SHEET_SCHEDULE, PERSISTENT_BLOCK_ID, buildScheduleRows_(PERSISTENT_BLOCK_ID, weeks));
  rewriteRowsForBlock_(ss, SHEET_PROGRAM, PERSISTENT_BLOCK_ID, buildProgramRows_(PERSISTENT_BLOCK_ID, weeks));
  return JSON.stringify({ success: true });
}

function buildScheduleRows_(blockId, weeks) {
  const rows = [];
  for (const w of weeks) {
    for (const d of (w.schedule || [])) {
      rows.push([blockId, w.weekNumber, d.day, d.type, d.label]);
    }
  }
  return rows;
}

function buildProgramRows_(blockId, weeks) {
  const rows = [];
  for (const w of weeks) {
    const exs = w.exercises || {};
    for (const [dayLabel, list] of Object.entries(exs)) {
      list.forEach((ex, i) => {
        rows.push([blockId, w.weekNumber, dayLabel, ex.name, ex.sets,
          ex.reps || '', ex.rpe || '', i + 1, (ex.group || '').toString().trim()]);
      });
    }
  }
  return rows;
}

function rewriteRowsForBlock_(ss, sheetName, blockId, newRows) {
  const sheet = ss.getSheetByName(sheetName);
  const cols = sheetName === SHEET_PROGRAM ? PROGRAM_COLS : SCHEDULE_COLS;
  if (sheet.getLastRow() >= 2) {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
    for (let i = data.length - 1; i >= 0; i--) {
      if (Number(data[i][0]) === blockId) sheet.deleteRow(i + 2);
    }
  }
  if (newRows.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, newRows.length, cols).setValues(newRows);
    if (sheetName === SHEET_PROGRAM) {
      sheet.getRange(startRow, 6, newRows.length, 1).setNumberFormat('@');
    }
  }
}

function endBlockById_(ss, blockId, endDate) {
  const sheet = ss.getSheetByName(SHEET_BLOCKS);
  if (!sheet || sheet.getLastRow() < 2) return;
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, BLOCK_COLS).getValues();
  for (let i = 0; i < data.length; i++) {
    if (Number(data[i][0]) === blockId) {
      sheet.getRange(i + 2, 5).setValue(endDate);
      sheet.getRange(i + 2, 6).setValue('completed');
      return;
    }
  }
}

function endCurrentBlock() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const active = getActiveBlock_(ss);
  if (!active) return JSON.stringify({ success: false, message: 'No active block.' });
  const tz = Session.getScriptTimeZone();
  const today = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const endDate = today < active.startDate ? active.startDate : today;
  endBlockById_(ss, active.id, endDate);
  return JSON.stringify({ success: true, ended: active.id, status: 'completed' });
}

// =============================================================================
// Initial Payload
// =============================================================================

function getInitData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupLogSheet_(ss);
  setupExerciseSheet_(ss);
  setupBlocksSheet_(ss);
  setupProgramSheet_(ss);
  setupDayScheduleSheet_(ss);
  setupSettingsSheet_(ss);
  migrateLegacyBlock1ToPersistent_(ss);

  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  const blocks = getAllBlocks_(ss);
  const active = blocks.find(b => b.status === 'active') || null;
  const past = blocks.filter(b => b.status !== 'active').map(b => hydrateBlock_(ss, b));
  const currentBlock = active ? hydrateBlock_(ss, active) : null;
  const persistent = getPersistentProgram_(ss);
  const persistentStartDate = getPersistentStartDate_(ss);
  const persistentExists = persistent.schedule.length > 0;

  return JSON.stringify({
    todayStr,
    persistentStartDate,
    persistentProgram: persistent,
    persistentExists,
    currentBlock,
    pastBlocks: past,
    settings:           getSettings_(ss),
    exercises:          getExercises_(ss),
    allLoggedExercises: getLoggedExerciseNames_(ss),
    loggedDates:        getLoggedDates_(ss),
  });
}

/**
 * Resolve the program for a date: block context first, falling back to the
 * persistent program for that day. The dayLabel/dayType come from whichever
 * source covers the date.
 */
function getSessionSetup(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const blockCtx = resolveBlockContext_(ss, dateStr);

  let programForDay = [];
  let dayLabel = null, dayType = null;
  let blockContextOut = null;

  if (blockCtx) {
    programForDay = blockCtx.programForDay;
    dayLabel = blockCtx.dayLabel;
    dayType = blockCtx.dayType;
    blockContextOut = {
      blockId: blockCtx.block.id,
      blockName: blockCtx.block.name,
      blockStatus: blockCtx.block.status,
      lengthWeeks: blockCtx.block.lengthWeeks,
      weekInBlock: blockCtx.weekInBlock,
      dayInWeek: blockCtx.dayInWeek,
      dayLabel: blockCtx.dayLabel,
      dayType: blockCtx.dayType,
    };
  } else {
    // Fall back to persistent program. We need to know which "day of week" within
    // the persistent week this date represents — driven by persistentStartDate.
    const persistentStart = getPersistentStartDate_(ss);
    if (persistentStart) {
      const persistent = getPersistentProgram_(ss);
      const days = daysBetween_(persistentStart, dateStr);
      if (days >= 0) {
        const dayInWeek = ((days % 7) + 7) % 7 + 1; // robust against negative
        const dayEntry = persistent.schedule.find(s => s.day === dayInWeek);
        if (dayEntry) {
          dayLabel = dayEntry.label;
          dayType = dayEntry.type;
          programForDay = (persistent.program['1'] && persistent.program['1'][dayLabel]) || [];
        }
      }
    }
  }

  const existing = getExistingLogForDate_(ss, dateStr);
  const programNames = programForDay.map(p => p.exerciseName);
  const loggedNames = Object.keys(existing.sets);
  const allNames = [...new Set([...programNames, ...loggedNames])];
  const lastSession = getLastSessionData_(ss, allNames, dateStr);

  return JSON.stringify({
    lastSession,
    existingLog: existing.sets,
    existingNotes: existing.notes,
    programForDay,
    dayLabel, dayType,
    blockContext: blockContextOut,
  });
}

// =============================================================================
// Exercise CRUD
// =============================================================================

function createExercise(exData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_EXERCISES);
  if (!sheet) return JSON.stringify({ success: false, message: 'Exercise sheet missing.' });
  const name = String(exData.name || '').trim();
  if (!name) return JSON.stringify({ success: false, message: 'Name is required.' });
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
// Last Working Sets
// =============================================================================

function getLastWorkingSets(exerciseName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet || sheet.getLastRow() < 2) return JSON.stringify({ sessions: [] });
  const tz = Session.getScriptTimeZone();
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();

  const byDate = {};
  for (const r of data) {
    if (r[3] !== exerciseName) continue;
    const dateStr = r[0] instanceof Date
      ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd')
      : String(r[0]).trim();
    if (!byDate[dateStr]) byDate[dateStr] = [];
    byDate[dateStr].push(readLogRow_(r, dateStr, lastCol));
  }

  const sessions = Object.keys(byDate)
    .sort((a, b) => b.localeCompare(a))
    .slice(0, 3)
    .map(date => {
      const sets = byDate[date].sort((a, b) => a.setNumber - b.setNumber);
      const noted = sets.find(s => s.notes);
      return {
        date,
        sets: sets.map(s => ({
          setNumber: s.setNumber, weight: s.weight, reps: s.reps,
          minutes: s.minutes, seconds: s.seconds, distance: s.distance, rpe: s.rpe,
        })),
        notes: noted ? noted.notes : '',
      };
    });

  return JSON.stringify({ sessions });
}

function readLogRow_(r, dateStr, lastCol) {
  if (lastCol >= 13) {
    return {
      date: dateStr, setNumber: r[4], weight: r[5], reps: r[6],
      minutes: r[7], seconds: r[8], distance: r[9], rpe: r[10], notes: r[11] || '',
    };
  }
  return {
    date: dateStr, setNumber: r[4], weight: r[5], reps: r[6],
    minutes: '', seconds: '', distance: '', rpe: r[7], notes: '',
  };
}

// =============================================================================
// Data Readers
// =============================================================================

function getExercises_(ss) {
  const sheet = ss.getSheetByName(SHEET_EXERCISES);
  if (!sheet || sheet.getLastRow() < 2) return [];
  const lastCol = Math.max(sheet.getLastColumn(), 4);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  return data.map(r => ({
    name: r[0], muscleGroup: r[1], equipment: r[2],
    notes: r[3] || '',
    measurementType: (r[4] && String(r[4]).trim()) || 'weight_reps',
    cardioModality: r[5] ? String(r[5]).trim() : '',
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
// Save / Delete Workout
// =============================================================================

function saveWorkout(sessionData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG);
  setupLogSheet_(ss);
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

function getSettings_(ss) {
  const sheet = ss.getSheetByName(SHEET_SETTINGS);
  if (!sheet || sheet.getLastRow() < 2) return {};
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const map = {};
  for (const [k, v] of data) map[String(k).trim()] = v;
  const tz = Session.getScriptTimeZone();
  let startDate = '';
  if (map.startDate instanceof Date) startDate = Utilities.formatDate(map.startDate, tz, 'yyyy-MM-dd');
  else startDate = String(map.startDate || '').trim();
  return {
    startDate,
    heightCm: Number(map.heightCm) || 0,
    gender: String(map.gender || 'male').trim().toLowerCase(),
    birthYear: Number(map.birthYear) || 0,
    activityLevel: String(map.activityLevel || 'moderate').trim().toLowerCase(),
    supersetsEnabled: String(map.supersetsEnabled || 'false').trim().toLowerCase() === 'true',
    theme: String(map.theme || 'blue').trim().toLowerCase(),
  };
}

function saveSettings(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  setupSettingsSheet_(ss);
  const sheet = ss.getSheetByName(SHEET_SETTINGS);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 2).getValues();
  const updates = {
    heightCm: payload.heightCm !== undefined ? payload.heightCm : null,
    gender: payload.gender !== undefined ? payload.gender : null,
    birthYear: payload.birthYear !== undefined ? payload.birthYear : null,
    activityLevel: payload.activityLevel !== undefined ? payload.activityLevel : null,
    supersetsEnabled: payload.supersetsEnabled !== undefined
      ? (payload.supersetsEnabled ? 'true' : 'false') : null,
    startDate: payload.startDate !== undefined ? payload.startDate : null,
    theme: payload.theme !== undefined ? payload.theme : null,
  };
  for (let i = 0; i < data.length; i++) {
    const key = String(data[i][0]).trim();
    if (updates[key] !== null && updates[key] !== undefined) {
      sheet.getRange(i + 2, 2).setValue(updates[key]);
    }
  }
  return JSON.stringify({ success: true });
}

// =============================================================================
// Progress (with optional date-range filter)
// =============================================================================

function getProgressData(exerciseName, dateRange) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_LOG);
  if (!sheet || sheet.getLastRow() < 2) return JSON.stringify({ sessions: [], measurementType: 'weight_reps' });
  const exercises = getExercises_(ss);
  const ex = exercises.find(e => e.name === exerciseName);
  const mt = ex ? ex.measurementType : 'weight_reps';
  const tz = Session.getScriptTimeZone();
  const lastCol = sheet.getLastColumn();
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, lastCol).getValues();
  const rows = data.filter(r => {
    if (r[3] !== exerciseName) return false;
    if (!dateRange) return true;
    const rd = r[0] instanceof Date
      ? Utilities.formatDate(r[0], tz, 'yyyy-MM-dd') : String(r[0]).trim();
    if (dateRange.start && rd < dateRange.start) return false;
    if (dateRange.end && rd > dateRange.end) return false;
    return true;
  });
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
      let maxE1RM = 0, totalVolume = 0, totalWeight = 0;
      for (const s of sorted) {
        const w = Number(s.weight) || 0, r = Number(s.reps) || 0;
        totalVolume += w * r;
        totalWeight += w;
        const e1rm = r > 0 ? w * (1 + r / 30) : 0;
        if (e1rm > maxE1RM) maxE1RM = e1rm;
      }
      if (maxE1RM > bestE1RM) { bestE1RM = maxE1RM; bestDate = date; }
      session.e1rm = Math.round(maxE1RM * 10) / 10;
      session.volume = Math.round(totalVolume);
      session.firstSetWeight = Number(first.weight) || 0;
      session.avgWeight = Math.round((totalWeight / sorted.length) * 10) / 10;
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
    measurementType: mt, sessions,
    bestE1RM: { value: Math.round(bestE1RM * 10) / 10, date: bestDate },
    totalSessions: sessions.length, volumeChange,
  });
}

// =============================================================================
// Drafts
// =============================================================================

function saveDraft(draftJson) { PropertiesService.getScriptProperties().setProperty('workoutDraft', draftJson); }
function loadDraft() { return PropertiesService.getScriptProperties().getProperty('workoutDraft') || null; }
function clearDraft() { PropertiesService.getScriptProperties().deleteProperty('workoutDraft'); }
