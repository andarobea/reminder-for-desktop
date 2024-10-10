Sub addScheduledTask()
    Dim strFileName, strTaskName, objShell, objFSO
    strFileName = "C:\Users\sorin\OneDrive\Desktop\code girl summer\reminder box\reminder.vbs"
    strTaskName = "ReminderTask"

    ' Check if reminder.vbs exists
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    If Not objFSO.FileExists(strFileName) Then
        MsgBox "reminder.vbs does not exist", 16, "Error"
        Exit Sub
    End If

    ' Create the scheduled task using schtasks command
    Set objShell = CreateObject("WScript.Shell")
    objShell.Run "schtasks /create /tn """ & strTaskName & """ /tr ""wscript.exe """ & strFileName & """ /sc once /st 21:37 /f /rl highest /it", 0, True

    ' Verify if the task was created
    If Err.Number = 0 Then
        MsgBox "Scheduled task created successfully!", 64, "Success"
    Else
        MsgBox "Failed to create scheduled task. Error code: " & Err.Number, 16, "Error"
    End If
End Sub

' Run the addScheduledTask subroutine
addScheduledTask