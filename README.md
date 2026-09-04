COPY-PASTE MANAGER v5.4 — TOP 5 IMPROVEMENTS
===========================================

Implemented:
1. Resume / Skip Completed
2. Dry Run + Task Summary
3. Multiple status concepts in the UI
4. Safe Mode / approval gates
5. Clipboard reliability settings and stop/resume behavior

1. RESUME / SKIP COMPLETED
--------------------------
Processing modes:
- Skip Completed (recommended/default)
- Copy All
- Pending Only
- Resume

The app reads the status column and treats "Completed" as already finished.
The default mode does not re-copy those rows.

2. DRY RUN + TASK SUMMARY
-------------------------
Before clipboard processing, the app shows:
- selected file
- worksheet
- data column
- row range
- status column
- rows with data
- already completed rows
- new values to process

The user must approve the final dry-run confirmation before copying starts.

3. STATUS TRACKING
------------------
The result table tracks:
- Sheet Row
- Content
- Clipboard
- Spreadsheet Status

A copied item is shown as "✓ Copied".
The spreadsheet is not marked Completed automatically.

4. SAFE MODE
------------
The UI explicitly states local-only operation and keeps the read and write
operations separate.

Before copying:
- explicit approval required.

Before spreadsheet modification:
- all selected values must be copied;
- explicit second approval required.

Existing status cells:
- the current implementation protects completed rows through Skip Completed.
- When the app is about to modify the status column, the user sees a clear
  warning that the spreadsheet will be modified.

5. CLIPBOARD RELIABILITY
------------------------
Clipboard delay is configurable from 100 to 3000 ms.
Start at 500 ms. If Windows misses entries, try 750 or 1000 ms.

The STOP button stops between clipboard operations.
The app does not mark the spreadsheet Completed after a partial stop.

IMPORTANT PRIVACY NOTE
----------------------
The app itself is local and does not upload data, call a cloud service, use
telemetry, create a database, scan unrelated files, or make Registry/startup
changes.

However, Win+V is Windows Clipboard History. Windows controls its retention.

For sensitive data, turn OFF Windows Clipboard Sync:
Settings -> System -> Clipboard -> Clipboard sync -> Off

SPREADSHEET STATUS
------------------
The status column is intentionally written only after explicit completion
approval. The app writes "Completed" for rows in the selected range whose data
column contains a non-empty value.

For .xlsx/.xlsm, the same selected workbook is saved after approval.
For CSV, the same selected CSV is updated after approval.

INSTALL
-------
Run install.bat once, then run run.bat.

STANDALONE EXE
--------------
Run build_exe.bat.
The resulting application is:
dist\Copy-Paste-Manager.exe


v5.4 UI FIXES:
- START PROCESSING is fixed in a bottom action bar and remains visible.
- Main content is vertically scrollable.
- Preview has a dedicated Preview / Validate button and scrollbar.
- Action buttons remain accessible at normal Windows display scaling.


v5.4 CLIPBOARD RELIABILITY FIXES:
- Default clipboard delay increased to 1500 ms.
- Each clipboard write is read back and verified before continuing.
- Clipboard opening retries for up to 2 seconds.
- A short settle interval is used to give Windows Clipboard History time to capture each item.
- A failed verification stops the task rather than silently moving to the next row.
- Delay range is 500–5000 ms.

IMPORTANT: Windows Clipboard History is asynchronous. No application can guarantee its internal history database records every item instantly. Identical repeated values may also be de-duplicated by Windows. For sensitive data, keep Clipboard Sync OFF.
