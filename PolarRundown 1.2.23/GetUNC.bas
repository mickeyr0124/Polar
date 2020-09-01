Attribute VB_Name = "GetUNC"
Global WroteNPSHr As Boolean

' <VB WATCH>
Const VBWMODULE = "GetUNC"
' </VB WATCH>

Public Function GetUNCFromLetter(DriveLetter As String) As String
' <VB WATCH>
1          On Error GoTo vbwErrHandler
2          Const VBWPROCNAME = "GetUNC.GetUNCFromLetter"
3          If vbwProtector.vbwTraceProc Then
4              Dim vbwProtectorParameterString As String
5              If vbwProtector.vbwTraceParameters Then
6                  vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("DriveLetter", DriveLetter) & ") "
7              End If
8              vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
9          End If
' </VB WATCH>
10         Dim PID As Long
11         Dim hProcess As Long
12         Dim str As String

13         Dim dirname As String
14         dirname = App.Path & "\NetUse.txt"

15         PID = Shell("cmd /c net use f: | cmd /c find ""Remote""  > " & Chr(34) & dirname & Chr(34))
16         If PID = 0 Then
                '
                'Handle Error, Shell Didn't Work
                '
17         Else
18              hProcess = OpenProcess(&H100000, True, PID)
19              WaitForSingleObject hProcess, -1
20              CloseHandle hProcess
21         End If

22         Open App.Path & "\NetUse.txt" For Input As #1
23         Input #1, str

24         PID = InStr(1, str, "\")


25         GetUNCFromLetter = Right$(str, Len(str) - PID + 1)

26         Close #1

' <VB WATCH>
27         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
28         Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "DriveLetter", DriveLetter
            vbwReportVariable "PID", PID
            vbwReportVariable "hProcess", hProcess
            vbwReportVariable "str", str
            vbwReportVariable "dirname", dirname
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function



' <VB WATCH> <VBWATCHFINALPROC>
' Procedures added by VB Watch for variable dump


Private Sub vbwReportModuleVariables()
    vbwReportToFile VBW_MODULE_STRING
End Sub
' </VB WATCH>
