# Workout Log Tool

A mobile-friendly workout logger built on Google Apps Script. All data lives in a Google Spreadsheet you own. No app install needed — just a bookmarked URL.

---

## Setup

1. **Create a spreadsheet**: Go to [Google Sheets](https://sheets.google.com), create a new blank spreadsheet.

2. **Open the script editor**: Extensions → Apps Script.

3. **Add the files**: Replace the default `Code.gs` contents. Then click **+** next to Files and add:
   - Script file → `Config` (paste `Config.gs` contents)
   - HTML file → `Index` (paste `Index.html` contents)

4. **Run setup**: Select `setupAllSheets` from the function dropdown, click ▶ Run. Authorize when prompted. This creates and populates four sheets: Exercise Selection, Workout Log, Program, and Settings.

5. **Set your start date**: In the **Settings** sheet, enter your cycle start date in cell B2 as `YYYY-MM-DD` (e.g., `2026-03-09`). Leave blank to auto-set to today on first use.

6. **Fill in your program**: In the **Program** sheet, add rows defining your routine. Columns: `Day Label` (must match labels in Config.gs), `Exercise Name` (must match Exercise Selection), `Target Sets`, `Target Reps`, `Target RPE`, `Exercise Order`.

7. **Deploy**: Deploy → New deployment → Web app. Set "Execute as" to Me, "Who has access" to Only myself. Bookmark the URL on your phone.

### Updating after code changes

Deploy → Manage deployments → edit → set version to "New version" → Deploy. Or use the test URL (Deploy → Test deployments) during development — it always runs your latest saved code without redeploying.

---

## Customizing the Cycle

Edit `Config.gs` to change the training cycle structure. The default is a 7-day cycle (3 workout, 2 active rest, 2 rest). You can use any cycle length and any combination of day types (`Workout`, `Active Rest`, `Rest`). Examples are included in the file comments. After changing labels, update the Program sheet's "Day Label" column to match.

---

## Features

- **Day selection**: Pick any week and day to log. Days with existing data show a green dot.
- **Program templates**: Session forms are pre-built from your Program sheet with target sets/reps/RPE.
- **Previous session reference**: Each exercise shows your last session's heaviest set; a copy button pulls those numbers in.
- **Per-set logging**: Weight (kg), reps, and RPE (optional) for each set. Add extra sets beyond the template.
- **Swap exercise**: Replace any exercise with another from the database (clears the set data).
- **Clear / Undo Clear**: Wipe all fields for one exercise. Undo restores the original values (disappears once you start typing new data or swap the exercise).
- **Remove / Undo Remove**: Soft-delete an exercise from the session with an undo placeholder to bring it back.
- **Add exercises**: Add exercises beyond the template from the full exercise database.
- **Modify existing sessions**: Reopen and edit any previously logged day.
- **Delete sessions**: Remove an entire day's log entries with confirmation.
- **Draft auto-save**: In-progress sessions are saved every 15 seconds and restored if the app is reopened.
- **Progress charts**: Estimated 1RM trend (Epley formula) and session volume over time, per exercise, with summary stats (best 1RM, session count, volume change).
