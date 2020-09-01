Attribute VB_Name = "GetUNC"
Global WroteNPSHr As Boolean

Public Function GetUNCFromLetter(DriveLetter As String) As String
    Dim PID As Long
    Dim hProcess As Long
    Dim str As String

    Dim dirname As String
    dirname = App.Path & "\NetUse.txt"

    PID = Shell("cmd /c net use f: | cmd /c find ""Remote""  > " & Chr(34) & dirname & Chr(34))
    If PID = 0 Then
         '
         'Handle Error, Shell Didn't Work
         '
    Else
         hProcess = OpenProcess(&H100000, True, PID)
         WaitForSingleObject hProcess, -1
         CloseHandle hProcess
    End If

    Open App.Path & "\NetUse.txt" For Input As #1
    Input #1, str

    PID = InStr(1, str, "\")


    GetUNCFromLetter = Right$(str, Len(str) - PID + 1)

    Close #1
  
End Function


