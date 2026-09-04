# Copy-Paste Manager v5 — Security Design

## Data minimization
The program reads only the user-selected spreadsheet and only the selected
worksheet, data column, status column, and row range necessary for the task.

## Local processing
No spreadsheet data is uploaded. The application source contains no HTTP
client, cloud API, telemetry, analytics, sockets or remote endpoints.

## Persistence
The application does not intentionally maintain a database, application log,
telemetry file, cache, preference store, or task history.

The spreadsheet is modified only after explicit approval.

## Separation of operations
Reading and clipboard loading are separate from spreadsheet writing.

Clipboard loading never automatically means task completion.

## Approval gates
A final dry-run confirmation is required before clipboard loading.

After the clipboard task finishes, the user must explicitly approve a second
operation before the spreadsheet is modified.

## Resume / Skip Completed
The status column is used to identify rows already marked Completed. The
recommended default mode skips those rows, reducing accidental duplicate work.

## Clipboard
The requested Win+V workflow requires writing selected values to the Windows
clipboard. Windows Clipboard History may retain those values.

For sensitive data, disable Clipboard Sync.

## Source file
Excel workbooks are initially opened read-only. On approved completion,
the selected workbook is reopened for a controlled status-column write and
saved back to the same selected file.

## Windows changes
The application does not intentionally edit the Registry, create startup
entries, scheduled tasks, services, desktop shortcuts, or Windows security
settings.

## Privileges
No administrator privileges are required.


## v5.2 UI and workflow fix
The main content is scrollable while the primary action bar remains fixed and visible. Spreadsheet completion writes are limited to rows in the active processed task.
