# Offline Hours Summary

Turns an Outlook calendar CSV export into the monthly **Offline Hours Report**
(`.md` + `.pdf`).

## Every month, three steps

**1. Export the calendar from Outlook**

> File → Save As → *Comma Separated Values (.csv)*
> (or File → Open & Export → Import/Export → Export to a file → CSV)

Set the date range to cover the **whole month**. Make sure you export the real
calendar, not an *Availability only* view — that strips every event title.

**2. Run it**

Double-click **`Offline Hours Summary.bat`**, type your name in the box at the
top, then drag the CSV onto the window. You can also drag the CSV straight onto
the `.bat` file.

The name is what appears as **Employee** on the report, so it is required — no
name, no report. It is remembered per Windows user and filled in for you next
time; the box stays editable, so change it whenever you like. Once a name is
remembered, dropping a CSV builds the report immediately as before. Otherwise
the file is held and the **Generate report** button does it once you have typed
a name.

**3. Collect the report**

`Offline_Hours_<Month><Year>.md` and `.pdf` are written **next to the CSV**.

---

## The rules it applies

| Situation | Treatment |
|---|---|
| **Busy** | Counted |
| **Working Elsewhere** | Counted |
| **Out of Office** | Leave — **not** counted |
| All-day event (e.g. Leave) | Not counted |
| Meeting inside a longer block | Shown, marked *concurrent*, counted **once** |

That last one matters: on a site-visit day where the standing meetings fall
inside the visit, adding everything up would over-count. The report charges the
elapsed time once and flags the nested entries.

## Two Outlook export bugs it works around

Outlook's CSV export mishandles recurring series, and it does so silently:

- **Duplicates** — the same instance written several times on one date. These
  are removed automatically.
- **Dropped instances** — a weekly series missing from some weeks entirely.
  The tool spots the gaps and restores them, marking each one *(restored)* in
  the report. Untick the checkbox if you would rather it left them alone.

Both happened in July 2026, to the same recurring meeting.

## Notes

- Date format is detected automatically (`yyyy/m/d`, `d/m/y` or `m/d/y`).
- If the CSV spans more than one month, the month with the most events wins and
  the rest are ignored — the count is shown in the log.
- If the PDF cannot be written, it is almost always because the previous one is
  **open in Adobe Acrobat**. Close it and run again.
- Drag-and-drop needs `tkinterdnd2` (`pip install tkinterdnd2`). Without it the
  window still works — click the box to browse.
- The remembered name lives in `%APPDATA%\OfflineHoursSummary\settings.json`, not
  next to the script — so a shared copy of this folder does not carry one
  person's name to everybody else. Delete that file to be asked again.

## Adjusting it

Open `offline_hours.py`; the settings sit at the top under `CONFIG`:

- `LEAVE_STATUS_CODES` — which *Show time as* codes mean Leave. Currently `{"4"}`,
  which is the code the July 2026 `Leave` day used.
- `CATEGORIES` — the keyword buckets for the Activity Breakdown. First match wins.

Command line, if you ever want it:

```
python offline_hours.py "MyCalendar.csv" --nogui --name "Your Name"
```

`--name` is only needed the first time — after that the remembered name is used.
Passing it also updates what is remembered.
