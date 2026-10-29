# Partlink Daily Backup Task

This document describes the Windows Task Scheduler job that creates the RDSA Partlink backup and the separate GoodSync handoff to OneDrive.

## What runs

The scheduled task is registered at the root of Task Scheduler:

```text
\Partlink_daily-backup-and-restore
```

It runs daily at 6:00 PM as `SYSTEM` with highest privileges. Its action is:

```text
cmd /c cd /d C:\rdsaparts-1 && npm run Partlink_daily-backup-and-restore
```

The npm script launches:

```text
scripts/backup-and-restore/backup-and-restore.js
```

That script backs up the Sydney Supabase database tables and storage buckets, writes a timestamped local backup, and uploads the same run to the Frankfurt Supabase backup bucket.

The local backup location is:

```text
C:\Supabase - RDSA Partlink_Sydney\backups\<timestamp>\
```

GoodSync is a separate process. It watches or syncs this local backup directory to OneDrive; it is not part of the scheduled task.

## Prerequisites on a new PC

Before importing the task, ensure the new PC has:

- The repository at `C:\rdsaparts-1`.
- Node.js installed for all users, with `node` and `npm` available to the `SYSTEM` account.
- Project dependencies installed with `npm install`.
- A correctly configured `.env.local` containing the required Sydney and Frankfurt Supabase credentials.
- Permission to create and write to `C:\Supabase - RDSA Partlink_Sydney`.
- GoodSync configured separately for the backup directory and OneDrive destination.

Do not copy `.env.local` into source control or share it in the task export.

## Export the task

On the existing PC, open PowerShell as Administrator and export the task XML:

```powershell
Export-ScheduledTask `
  -TaskPath '\' `
  -TaskName 'Partlink_daily-backup-and-restore' |
  Set-Content `
    -Path "$env:USERPROFILE\Desktop\Partlink_daily-backup-and-restore.xml" `
    -Encoding Unicode
```

Copy the XML file to the new PC. The XML contains the schedule and command, but not the repository, Node.js installation, dependencies, or `.env.local`.

## Import the task

On the new PC, place the XML on the desktop and open PowerShell as Administrator:

```powershell
Register-ScheduledTask `
  -TaskName 'Partlink_daily-backup-and-restore' `
  -TaskPath '\' `
  -Xml (Get-Content "$env:USERPROFILE\Desktop\Partlink_daily-backup-and-restore.xml" -Raw) `
  -Force
```

Confirm the imported task:

```powershell
Get-ScheduledTask `
  -TaskPath '\' `
  -TaskName 'Partlink_daily-backup-and-restore' |
  Format-List TaskName, State, Principal, Actions, Triggers
```

The action must still point to `C:\rdsaparts-1`, and the principal should be `SYSTEM`.

## Test before waiting for the schedule

Run the task manually:

```powershell
Start-ScheduledTask `
  -TaskPath '\' `
  -TaskName 'Partlink_daily-backup-and-restore'
```

Check the result:

```powershell
Get-ScheduledTaskInfo `
  -TaskPath '\' `
  -TaskName 'Partlink_daily-backup-and-restore' |
  Format-List LastRunTime, LastTaskResult, NextRunTime
```

`LastTaskResult` of `0` indicates success. Also confirm that a new timestamped directory appears below:

```text
C:\Supabase - RDSA Partlink_Sydney\backups\
```

If the task fails under Task Scheduler but succeeds in an interactive PowerShell window, check that Node.js is installed system-wide and that the `SYSTEM` account can access the repository, `.env.local`, destination folder, and network endpoints.

## Retrieve or inspect the task later

```powershell
Get-ScheduledTask -TaskPath '\' -TaskName 'Partlink_daily-backup-and-restore' |
  Format-List *
```

For the verbose Task Scheduler view:

```powershell
schtasks /query /tn "\Partlink_daily-backup-and-restore" /fo LIST /v
```