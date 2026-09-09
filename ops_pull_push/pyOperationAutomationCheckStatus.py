git log --oneline -1        # should match the source machine's top commit
git status                  # should say "working tree clean" (ignoring credentials/.venv/output)


# ================= Schedule Check =======================
Get-ScheduledTask -TaskName "IntelliBI Operations Automation" | Get-ScheduledTaskInfo
# Reschedule the Operations:
$trigger = New-ScheduledTaskTrigger -Daily -At "11:15AM"
Set-ScheduledTask -TaskName "IntelliBI Operations Automation" -Trigger $trigger
# Start it manually.
Start-ScheduledTask -TaskName "IntelliBI Operations Automation"