# 5K Race Timer

A Streamlit app for timing a charity 5K race. Handles pre-registration, day-of sign-ups, live finish line timing, and automatic leaderboards by gender and age group.

---

## Requirements

```
pip install streamlit pandas openpyxl
```

## Running the app

```
streamlit run race_timer.py
```

---

## How to use it

### Before race day

**Registration tab → Import from Excel**

Upload your pre-registration spreadsheet. The app will auto-detect columns named Bib Number, Name, Age, and Gender (case-insensitive). Ages like `32.0` and gender values like `M`, `female`, or `MALE` are all normalized automatically.

**Registration tab → Age Groups**

At the bottom of the Registration tab, define the age group brackets you want to appear in the leaderboard. The app starts with a default set (U13, 13–20, 20–35, 35–50, 50–65, 65+), but you can delete any of them, add your own with custom labels and age ranges, or reset back to defaults. Changes take effect on the leaderboard immediately — no restart needed.

---

### Race day morning

**Registration tab → Add Day-Of Registrant**

Enter a bib number, name, age, and gender for any runner who registers on the day. They are added to the roster instantly and the backup file is updated.

---

### During the race

**Race Timing tab**

1. Press **Start Race** to begin the clock.
2. Press **Runner Crossed Finish Line** each time someone crosses. This records the exact elapsed time and wall clock time for that place.
3. In the finisher table, type each runner's bib number into the **Bib** column. Name, age, and gender populate automatically from the roster.
4. The leaderboard at the bottom of the page updates live as bib numbers are entered.

**Stopping the race**

Press **Stop Race** when the last runner finishes. You will be asked to confirm before the clock stops. After stopping, the finisher table and leaderboard remain fully editable — you can still enter bib numbers and correct any records.

**Removing a finisher row**

Each row in the finisher table has a **✕** button on the right. Clicking it prompts a confirmation before removing the record. Places are automatically renumbered after a deletion.

---

### Leaderboard rules

- **Overall Male** and **Overall Female** show the top 3 finishers across all ages for each gender.
- The overall winner for each gender is **excluded from all age group categories** — they win the overall title only.
- All other age group categories show the top 3 eligible finishers within that age range.

---

### Exports and backups

The app writes two files to disk automatically alongside `race_timer.py`:

| File | When it's written |
|---|---|
| `roster_backup.xlsx` | Every time a runner is imported or added |
| `finishers_backup.xlsx` | Every time a runner crosses the finish line or a bib number is entered |

These are safety backups in case the browser tab is closed or the app crashes.

From the Race Timing tab you can also manually download:

- **Finisher List (Excel)** — the full finisher table with place-based row highlighting
- **Finisher List (CSV)** — plain flat file
- **Full Leaderboard (Excel)** — a multi-sheet workbook with one sheet per age group, each showing the top 3 male and female finishers, plus an All Finishers sheet

---

### Notes

- All data lives in memory for the current session. Refreshing the browser will reset the app — use the backup files to recover if needed.
- The Reset Race button at the bottom of the Race Timing tab clears the timer and all finisher records, but does not touch the roster or age groups.
- Gender values accepted from Excel: `Male`, `Female`, `M`, `F` (any capitalization).
