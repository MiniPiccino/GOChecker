@echo off
set TASK_NAME=GO Weekly Snapshot
set TASK_CMD=C:\Users\RZbas\Projects\MCPClients\GetGO\fetch_weekly_snapshot.bat

:: Creates a weekly task for Mondays at 09:30 for the current user (no password required).
schtasks /Create /F /TN "%TASK_NAME%" /TR "%TASK_CMD%" /SC WEEKLY /D MON /ST 09:30 /RL LIMITED

if %ERRORLEVEL%==0 (
  echo Task created: %TASK_NAME%
) else (
  echo Failed to create task. You may need to run this as your user in an elevated prompt.
)
