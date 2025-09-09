Dim nextReminderTime As Date

Sub StartSaveReminder()
    nextReminderTime = Now + TimeValue("00:05:00")
    Application.OnTime nextReminderTime, "ShowSaveReminder"
End Sub

Sub ShowSaveReminder()
    MsgBox "Reminder: Save your work!", vbInformation, "Save Reminder"
    StartSaveReminder ' Reschedule next reminder
End Sub

Sub StopSaveReminder()
    On Error Resume Next
    Application.OnTime nextReminderTime, "ShowSaveReminder", , False
End Sub

