Attribute VB_Name = "Utilities"
Option Explicit


Public Function ConnectedToDB(DB As Database, Path As String, Optional ReadOnly As Boolean, Optional ConnectString As String, Optional ByVal ConnectRegardless As Boolean) As Boolean
On Error GoTo ErrorHandler
Dim strName As String
Dim secAttempt As Boolean

If Not ConnectRegardless Then
    If Valid(DB) Then ConnectRegardless = DB.Name = "" Else ConnectRegardless = True                              'If we get here we need to, you guessed it, connect regardless...
End If
If ConnectRegardless Then
ct:     Set DB = Workspaces(0).OpenDatabase(Path, False, False, ConnectString)
    ConnectedToDB = Valid(DB)
Else
    ConnectedToDB = True
End If
Exit Function
ErrorHandler:
Select Case Err
    Case 91
        ConnectRegardless = False
        Resume Next
    Case 3044 'Not a valid path.
        If Path <> "" Then
        strName = Path
            Do Until InStr(strName, "\") = 0
                strName = Mid$(strName, InStr(strName, "\") + 1)
            Loop
        End If
        MsgBox "Can't find Database.  Path is invalid or not longer exits." & vbCr & vbCr & "[" & strName & "]" & vbCr, vbOKOnly + vbInformation, App.Title & " - Path Not Found"
    Case 3031
        ''Requires a password, so crack it
        'ConnectString = "MS Access;pwd=" & GetPassword(Path)
        'Resume
    Case Else
        If Not secAttempt Then
            secAttempt = True
            Resume ct
        End If
End Select
Exit Function
End Function

Public Function AlreadyRunning() As Boolean
    'Check for previously loaded instances of the program'''''''''''
    If App.PrevInstance = True Then
        Beep
        MsgBox "Program cancelled." & vbCr & vbCr & _
            "There is a previous copy of " & App.Title & " already running.  " & vbCr & "Please" _
            & " check your currently running applications in " & vbCr & "the task manager and try again." _
              & vbCr & vbCr & "(Task Manager:  Ctrl+Alt+Del)", vbOKOnly + vbExclamation, App.Title & " already running!"
        End
        AlreadyRunning = True
    End If
End Function

Public Function Valid(Thing As Object) As Boolean
Valid = Not Thing Is Nothing
End Function

Public Sub xBeep(Optional lTimes&)
On Error Resume Next
Dim lx&
If lTimes = 0 Then lTimes = 10
For lx = 1 To lTimes
    Beep
Next
End Sub
