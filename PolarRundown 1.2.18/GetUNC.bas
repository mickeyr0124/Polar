Attribute VB_Name = "GetUNC"
Global WroteNPSHr As Boolean

' <VB WATCH>
Const VBWMODULE = "GetUNC"
' </VB WATCH>

Public Function GetUNCFromLetter(DriveLetter As String) As String
' <VB WATCH>
1          On Error GoTo vbwErrHandler
' </VB WATCH>
2          Dim PID As Long
3          Dim hProcess As Long
4          Dim str As String

5          Dim dirname As String
6          dirname = App.Path & "\NetUse.txt"

7          PID = Shell("cmd /c net use f: | cmd /c find ""Remote""  > " & Chr(34) & dirname & Chr(34))
8          If PID = 0 Then
                '
                'Handle Error, Shell Didn't Work
                '
9          Else
10              hProcess = OpenProcess(&H100000, True, PID)
11              WaitForSingleObject hProcess, -1
12              CloseHandle hProcess
13         End If

14         Open App.Path & "\NetUse.txt" For Input As #1
15         Input #1, str

16         PID = InStr(1, str, "\")


17         GetUNCFromLetter = Right$(str, Len(str) - PID + 1)

18         Close #1

' <VB WATCH>
19         Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetUNCFromLetter"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
    End Select
' </VB WATCH>
End Function



