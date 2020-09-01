Attribute VB_Name = "vbwProtector"
 ' vbwProtector.bas file - Location: \VB Watch 2\Templates\VB6\Protector\ '
'                                                                        '
' This module contains all procedures common to the VB Watch tools.      '
' It will be added to every project instrumented with VB Watch.          '
'                                                                        '
' ************************* WARNING *******************************      '
' You should not modify it unles you know what you are doing.            '
' To modify it, remove the read-only attribute of vbwProtector.bas.      '
 ' WARNING: modifications of this file will apply to all error handling   '
'          plans !!!                                                     '

Option Explicit

' Options '
Public vbwCatchException As Boolean
Public vbwTraceProc As Boolean
Public vbwTraceParameters As Boolean
Public vbwTraceLine As Boolean
Public vbwCallStack As Boolean
Public vbwEmailRecipientAdress As String
Public vbwDumpStringMaxLength As Long
Public vbwSystemInfo As Boolean
Public vbwScreenshot As Boolean

' Variables for use with vbwFunctions.dll '
Public vbwAdvancedFunctions As Object          ' this will be used only if vbwFunctions.dll is installed on the enduser machine '
Public fIsVbwFunctionsInitialized As Boolean   ' true if vbwFunctions.dll is installed and instanciated                         '

' Call Stack '
Public vbwStackCalls() As String     ' array containing each call of the stack '
Public vbwStackCallsNumber As Long   ' number of calls = Ubound(vbwStackCalls) '

' Trace '
Public vbwTraceCallsNumber As Long   ' number of calls '

' Log File
Dim fIsLogInitialize As Boolean
Public vbwLogFile As String
Public vbwLogTraceToFile As Boolean
Dim fLogFileOpen As Boolean
Dim lLogFileNumber As Long
Dim lLogFileOffset As Long
' file I/O
Private Const GENERIC_WRITE = &H40000000
Private Const FILE_SHARE_READ = &H1
Private Const FILE_FLAG_OVERLAPPED = &H40000000
Private Const OPEN_ALWAYS = 4
Private Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        OffSet As Long
        OffsetHigh As Long
        hEvent As Long
End Type
Private Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Any, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
Private Declare Function WriteFileEx Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpOverlapped As OVERLAPPED, ByVal lpCompletionRoutine As Long) As Long

' Var Dump
Const VBW_STRING = "**************************"
Global Const VBW_LOCAL_STRING = vbCrLf & VBW_STRING & vbCrLf & "* LOCAL LEVEL VARIABLES  *" & vbCrLf & VBW_STRING & vbCrLf
Global Const VBW_MODULE_STRING = vbCrLf & VBW_STRING & vbCrLf & "* MODULE LEVEL VARIABLES *" & vbCrLf & VBW_STRING & vbCrLf
Global Const VBW_GLOBAL_STRING = vbCrLf & VBW_STRING & vbCrLf & "* GLOBAL LEVEL VARIABLES *" & vbCrLf & VBW_STRING & vbCrLf
Global Const VBW_TYPE_STRING = " (User Defined Type Array)"
Global Const VBW_UNKNOWN_STRING = " = {Unknown Type}"
Global Const VBW_LOCAL_NOT_REPORTED = "Local Variables: not reported"
Global Const VBW_MODULE_NOT_REPORTED = "Module Variables: not reported"
Global Const VBW_GLOBAL_NOT_REPORTED = "Global Variables: not reported"
Global Const VBW_NO_LOCAL_VARIABLES = "No Local Variables"
Global vbwDumpFile As String
Global vbwDumpFileNum As Long

' Thread & processes
Private Declare Function GetCurrentThread Lib "kernel32" () As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Private Declare Sub ExitThread Lib "kernel32" (ByVal dwExitCode As Long)
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function GetExitCodeProcess Lib "kernel32" (ByVal hProcess As Long, lpExitCode As Long) As Long
Private Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Private Declare Function GetModuleFileName Lib "kernel32" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long

' Exception handling declarations
Private Declare Function SetUnhandledExceptionFilter Lib "kernel32" (ByVal lpTopLevelExceptionFilter As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Const EXCEPTION_CONTINUE_EXECUTION = -1
Private Const EXCEPTION_MAXIMUM_PARAMETERS = 15
Private Type EXCEPTION_RECORD
    ExceptionCode As Long
    ExceptionFlags As Long
    pExceptionRecord As Long    ' Pointer to an EXCEPTION_RECORD structure
    ExceptionAddress As Long
    NumberParameters As Long
    ExceptionInformation(EXCEPTION_MAXIMUM_PARAMETERS) As Long
End Type
Private Type EXCEPTION_DEBUG_INFO
        pExceptionRecord As EXCEPTION_RECORD
        dwFirstChance As Long
End Type
Private Type CONTEXT
    dblVar(66) As Double ' The real structure is more complex
    lngVar(6) As Long    ' but we don't need those details
End Type
Private Type EXCEPTION_POINTERS
    pExceptionRecord As EXCEPTION_RECORD
    ContextRecord As CONTEXT
End Type
Private Const EXCEPTION_ACCESS_VIOLATION = &HC0000005
Private Const EXCEPTION_DATATYPE_MISALIGNMENT = &H80000002
Private Const EXCEPTION_BREAKPOINT = &H80000003
Private Const EXCEPTION_SINGLE_STEP = &H80000004
Private Const EXCEPTION_ARRAY_BOUNDS_EXCEEDED = &HC000008C
Private Const EXCEPTION_FLT_DENORMAL_OPERAND = &HC000008D
Private Const EXCEPTION_FLT_DIVIDE_BY_ZERO = &HC000008E
Private Const EXCEPTION_FLT_INEXACT_RESULT = &HC000008F
Private Const EXCEPTION_FLT_INVALID_OPERATION = &HC0000090
Private Const EXCEPTION_FLT_OVERFLOW = &HC0000091
Private Const EXCEPTION_FLT_STACK_CHECK = &HC0000092
Private Const EXCEPTION_FLT_UNDERFLOW = &HC0000093
Private Const EXCEPTION_INT_DIVIDE_BY_ZERO = &HC0000094
Private Const EXCEPTION_INT_OVERFLOW = &HC0000095
Private Const EXCEPTION_PRIV_INSTRUCTION = &HC0000096
Private Const EXCEPTION_IN_PAGE_ERROR = &HC0000006
Private Const EXCEPTION_ILLEGAL_INSTRUCTION = &HC000001D
Private Const EXCEPTION_NONCONTINUABLE_EXCEPTION = &HC0000025
Private Const EXCEPTION_STACK_OVERFLOW = &HC00000FD
Private Const EXCEPTION_INVALID_DISPOSITION = &HC0000026
Private Const EXCEPTION_GUARD_PAGE = &H80000001
Private Const EXCEPTION_INVALID_HANDLE = &HC0000008
Private Const CONTROL_C_EXIT = &HC000013A

' Variable to Save the Err object
Dim ErrObjectDescription As String
Dim ErrObjectHelpContext As Long
Dim ErrObjectHelpFile As String
Dim ErrObjectLastDllError As Long
Dim ErrObjectNumber As Long
Dim ErrObjectSource As String
Dim ErrLine As Long
Private Declare Function FormatMessage Lib "kernel32" Alias "FormatMessageA" (ByVal dwFlags As Long, lpSource As Any, ByVal dwMessageId As Long, ByVal dwLanguageId As Long, ByVal lpBuffer As String, ByVal nSize As Long, Arguments As Long) As Long

Public VBWPROTECTOR_EMPTY As Variant ' for use with vbwExecuteLine() in IIf structures

'Const VBW_EXE_EXTENSION = ".exe" ' this line will be rewritten by VB Watch with the right extension

' vbwNoTraceProc vbwNoTraceLine ' don't remove this !

#Const PROJECT = "PolarRundown.vbp"
' <VB WATCH>
Const VBWMODULE = "vbwProtector"
Global Const VBWPROJECT = "PolarRundown"
Global Const VBW_EXE_EXTENSION = ".exe"
' </VB WATCH>

Sub vbwInitializeProtector()
' <VB WATCH>
1          On Error GoTo vbwErrHandler
' </VB WATCH>

2          Static vbwIsInitialized As Boolean

3          If vbwIsInitialized Then
4              Exit Sub
5          End If

       ' Don't remove the following comments !                                                         '
       ' VB Watch will replace next line with the initialization code as set in the plan being applied '
6      vbwEmailRecipientAdress = "mrosenbaum@chempump.com" ' this will be replaced with the value found in the 'General options' tab below
7      vbwCatchException = True ' set to False if you have already an exception catcher
8      vbwTraceLine = InStr(Command$, "/trace") > 0 Or GetSetting(App.title, "Init", "vbwTrace", "") = "1"
9      vbwTraceProc = InStr(Command$, "/trace") > 0 Or GetSetting(App.title, "Init", "vbwTrace", "") = "1"
10     vbwTraceParameters = vbwTraceProc
11     vbwCallStack = True
12     vbwSystemInfo = True
13     vbwScreenshot = True

14         vbwLogTraceToFile = vbwTraceProc Or vbwTraceLine
15         If vbwCallStack Then ' needed to track call stack
16              vbwTraceProc = True
17         End If

18         vbwLogFile = App.Path & "\vbw" & App.EXEName & VBW_EXE_EXTENSION & ".log"
19         vbwDumpFile = App.Path & "\vbw" & App.EXEName & VBW_EXE_EXTENSION & ".dmp"

20         If vbwCatchException Then
21             vbwHandleException
22         End If

23         vbwDumpStringMaxLength = 128 ' change this value to suit your need - make it 0 to remove the size check (to use with caution)

24         vbwIsInitialized = True

' <VB WATCH>
25         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwInitializeProtector"

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
            vbwReportVariable "vbwIsInitialized", vbwIsInitialized
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
' </VB WATCH>
End Sub

Sub vbwReportVariable(ByVal lName As String, ByVal lValue As Variant, Optional ByVal lTab As Long)
       ' vbwNoErrorHandler ' don't remove this !
26         Dim i As Long, j As Long, k As Long, L As Long
27         Dim tDim As Long

28         On Error GoTo ErrDump

29         If InStr(1, TypeName(lValue), "()", vbBinaryCompare) Then
               ' array '
30             tDim = GetArrayDimension(lValue)
31             Select Case tDim
                   Case 1
32                     vbwReportToFile String$(lTab, vbTab) & vbwEncryptString("Array " & lName & "(" & LBound(lValue, 1) & " To " & UBound(lValue, 1) & ") As " & TypeName(lValue))
33                     For i = LBound(lValue) To UBound(lValue)
34                         vbwReportVariable lName & "(" & i & ")", lValue(i), lTab + 1
35                     Next i
36                 Case 2
37                     vbwReportToFile String$(lTab, vbTab) & vbwEncryptString("Array " & lName & "(" & LBound(lValue, 1) & " To " & UBound(lValue, 1) & "," & LBound(lValue, 2) & " To " & UBound(lValue, 2) & ") As " & TypeName(lValue))
38                     For j = LBound(lValue, 2) To UBound(lValue, 2)
39                         For i = LBound(lValue, 1) To UBound(lValue, 1)
40                             vbwReportVariable lName & "(" & i & "," & j & ")", lValue(i, j), lTab + 1
41                         Next i
42                     Next j
43                 Case 3
44                     vbwReportToFile String$(lTab, vbTab) & vbwEncryptString("Array " & lName & "(" & LBound(lValue, 1) & " To " & UBound(lValue, 1) & "," & LBound(lValue, 2) & " To " & UBound(lValue, 2) & "," & LBound(lValue, 3) & " To " & UBound(lValue, 3) & ") As " & TypeName(lValue))
45                     For k = LBound(lValue, 3) To UBound(lValue, 3)
46                         For j = LBound(lValue, 2) To UBound(lValue, 2)
47                             For i = LBound(lValue, 1) To UBound(lValue, 1)
48                                 vbwReportVariable lName & "(" & i & "," & j & "," & k & ")", lValue(i, j, k), lTab + 1
49                             Next i
50                         Next j
51                     Next k
52                 Case 4
53                     vbwReportToFile String$(lTab, vbTab) & vbwEncryptString("Array " & lName & "(" & LBound(lValue, 1) & " To " & UBound(lValue, 1) & "," & LBound(lValue, 2) & " To " & UBound(lValue, 2) & "," & LBound(lValue, 3) & " To " & UBound(lValue, 3) & "," & LBound(lValue, 4) & " To " & UBound(lValue, 4) & ") As " & TypeName(lValue))
54                     For L = LBound(lValue, 4) To UBound(lValue, 4)
55                         For k = LBound(lValue, 3) To UBound(lValue, 3)
56                             For j = LBound(lValue, 2) To UBound(lValue, 2)
57                                 For i = LBound(lValue, 1) To UBound(lValue, 1)
58                                     vbwReportVariable lName & "(" & i & "," & j & "," & k & "," & L & ")", lValue(i, j, k, L), lTab + 1
59                                 Next i
60                             Next j
61                         Next k
62                     Next L
63                 Case Else
64                     vbwReportToFile String$(lTab, vbTab) & vbwEncryptString("Array " & lName & "() not processed: " & tDim & " dimensions")
65             End Select
66         Else
               ' non-array '
67             If IsObject(lValue) Then
68                 vbwReportObject lName, lValue, lTab
69             Else
70                 If VarType(lValue) = vbString Then
71                     lValue = FormatString(lValue)
72                 End If
73                 vbwReportToFile String$(lTab, vbTab) & vbwEncryptString(lName & " = " & lValue & " (" & TypeName(lValue) & ")")
74             End If
75         End If
76         Exit Sub

77     ErrDump:
78         Err.Clear
79         vbwReportToFile String$(lTab, vbTab) & vbwEncryptString(lName & " = {Variable Dumping Error}")
End Sub

Public Sub vbwReportObject(lName As String, ByVal lObject As Object, Optional ByVal lTab As Long)
       ' vbwNoErrorHandler ' don't remove this !

80         On Error GoTo ErrDump

81         If TypeName(lObject) <> "ErrObject" Then
82             If fIsVbwFunctionsInitialized Then
                   ' this should be executed only if you are using a global error handler '
                   ' that prepares properly the vbwAdvancedFunctions for object dumping   '
83                 vbwCloseDumpFile       ' close it because vbwAdvancedFunctions uses its own file writing routines '
84                 vbwAdvancedFunctions.ReportObject lName, lObject, lTab, TypeOf lObject Is Form, TypeOf lObject Is MDIForm
85                 vbwOpenDumpFile
86             Else
                   ' no vbwFunctions.dll available                         '
                   ' only report the default value of objects and controls '
87                 If TypeOf lObject Is Form Or TypeOf lObject Is MDIForm Then
88                    On Error Resume Next
89                    vbwReportToFile vbwEncryptString("Form " & lName)
90                    Dim c As Control
91                    For Each c In lObject.Controls
92                        vbwReportObject c.Name & vbwGetIndex(c), c, 1
93                    Next c
94                 Else
95                     If IsNumeric(lObject) Then
96                         vbwReportVariable lName, CDbl(lObject), lTab
97                     Else
98                         vbwReportVariable lName, CStr(lObject), lTab
99                     End If
100                End If
101            End If
102        Else
103            vbwReportToFile vbCrLf & vbwEncryptString("**** ErrObject Err ****")
104            vbwReportVariable "Err.Number", ErrObjectNumber
105            vbwReportVariable "Err.Source", ErrObjectSource
106            vbwReportVariable "Err.Description", ErrObjectDescription
107            vbwReportVariable "Err.HelpContext", ErrObjectHelpContext
108            vbwReportVariable "Err.HelpFile", ErrObjectHelpFile
109            If ErrObjectLastDllError = 0 Then
110                vbwReportVariable "Err.LastDllError", ErrObjectLastDllError
111            Else
                   ' get the API error description from the system
112                Dim sBuffer As String * 512
113                Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
114                FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, Null, ErrObjectLastDllError, 0, sBuffer, 512, 0
115                If InStr(sBuffer, Chr(0)) Then
116                    vbwReportVariable "Err.LastDllError", ErrObjectLastDllError & " (" & Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1) & ")"
117                Else
118                    vbwReportVariable "Err.LastDllError", ErrObjectLastDllError
119                End If
120            End If
121        End If

122        Exit Sub

123    ErrDump:
124        Err.Clear
125        vbwReportToFile String$(lTab, vbTab) & vbwEncryptString(lName & ".Value = {No Value Property}")
End Sub

Public Function vbwEncryptString(ByRef sString As String, Optional sKey) As String
       ' vbwNoErrorHandler ' don't remove this ! '

126        On Error Resume Next
127        If fIsVbwFunctionsInitialized = False Then
               ' no encryption without vbwFunctions.dll            '
               ' you may want to write your own encryption routine '
128            vbwEncryptString = sString
129        Else
130            If IsMissing(sKey) Then
                   ' If you filled the vardump encryption key in the VB Watch Options, your key will   '
                   ' be already embeded in the vbwAdvancedFunctions.ObjectInfo property, so you do not '
                   ' have to care about provideing a key                                               '
131                vbwEncryptString = vbwAdvancedFunctions.EncryptString(sString)
132            Else
                   ' Yet if you wish to overide  your default encryption key, simply pass it '
                   ' in the sKey parameter                                                   '
133                vbwEncryptString = vbwAdvancedFunctions.EncryptString(sString, sKey)
134            End If
135        End If
End Function

Function vbwReportParameter(ByVal lName As String, ByRef lValue As Variant) As String
       ' vbwNoErrorHandler ' don't remove this !
136        Dim i As Long, j As Long, k As Long
137        Dim tDim As Long
138        Dim retString As String

139        On Error GoTo ErrDump

140        If InStr(1, TypeName(lValue), "()", vbBinaryCompare) Then
               ' array '
141            tDim = GetArrayDimension(lValue)
142            If tDim Then
143                retString = lName & "("
144                For i = 1 To tDim
145                    retString = retString & LBound(lValue, i) & " To " & UBound(lValue, i) & ","
146                Next i
147                Mid$(retString, Len(retString)) = ")"   ' Close the brackets by overwriting the last comma '
148            Else
149                retString = lName & "(Undimensioned Array)"
150            End If
151        Else
               ' non-array '
152            If IsObject(lValue) Then
                   ' object
153                On Error Resume Next
154                retString = TypeName(lValue) & " " & lName & " = " & CStr(lValue)
155                If Err.Number Then
156                    On Error GoTo ErrDump
157                    retString = TypeName(lValue) & " " & lName & " = " & lValue.Name & vbwGetIndex(lValue)
158                End If
159            Else
                   ' non-object
160                If VarType(lValue) = vbString Then
161                   retString = lName & " = " & FormatString(lValue)
162                Else
163                   retString = lName & " = " & lValue
164                End If
165            End If
166        End If

167        vbwReportParameter = retString
168        Exit Function

169    ErrDump:
170        Err.Clear
171        vbwReportParameter = lName & " = {" & TypeName(lValue) & ": Parameter Dumping Error}"
End Function

Function vbwReportParameterByVal(ByVal lName As String, ByVal lValue As Variant) As String
       ' vbwNoErrorHandler ' don't remove this !
172        Dim i As Long, j As Long, k As Long
173        Dim tDim As Long
174        Dim retString As String

175        On Error GoTo ErrDump

176        If InStr(1, TypeName(lValue), "()", vbBinaryCompare) Then
               ' array '
177            tDim = GetArrayDimension(lValue)
178            If tDim Then
179                retString = lName & "("
180                For i = 1 To tDim
181                    retString = retString & LBound(lValue, i) & " To " & UBound(lValue, i) & ","
182                Next i
183                Mid$(retString, Len(retString)) = ")"   ' Close the brackets by overwriting the last comma '
184            Else
185                retString = lName & "(Undimensioned Array)"
186            End If
187        Else
               ' non-array '
188            If IsObject(lValue) Then
                   ' object
189                On Error Resume Next
190                retString = TypeName(lValue) & " " & lName & " = " & CStr(lValue)
191                If Err.Number Then
192                    On Error GoTo ErrDump
193                    retString = TypeName(lValue) & " " & lName & " = " & lValue.Name & vbwGetIndex(lValue)
194                End If
195            Else
                   ' non-object
196                If VarType(lValue) = vbString Then
197                   retString = lName & " = " & FormatString(lValue)
198                Else
199                   retString = lName & " = " & lValue
200                End If
201            End If
202        End If

203        vbwReportParameterByVal = retString
204        Exit Function

205    ErrDump:
206        Err.Clear
207        vbwReportParameterByVal = lName & " = {" & TypeName(lValue) & ": Parameter Dumping Error}"
End Function

Sub vbwReportToFile(ByRef lString As String)
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
208        On Error GoTo vbwErrHandler
' </VB WATCH>
209         If vbwDumpFileNum = 0 Then
210              vbwOpenDumpFile
211         End If
212         On Error Resume Next
213         Print #vbwDumpFileNum, lString
214         If Err = 52 Then
215            vbwCloseDumpFile
216            vbwOpenDumpFile
217            Print #vbwDumpFileNum, lString
218         End If
' <VB WATCH>
219        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwReportToFile"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            ' <Dump>
            Err.Number = -1
            Goto vbwErrHandler
    End Select
' </VB WATCH>
End Sub

Sub vbwOpenDumpFile()
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
220        On Error GoTo vbwErrHandler
' </VB WATCH>
221       vbwDumpFileNum = FreeFile
222       Open vbwDumpFile For Append As #vbwDumpFileNum
' <VB WATCH>
223        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwOpenDumpFile"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            ' <Dump>
            Err.Number = -1
            Goto vbwErrHandler
    End Select
' </VB WATCH>
End Sub

Sub vbwCloseDumpFile()
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
224        On Error GoTo vbwErrHandler
' </VB WATCH>
225       Close #vbwDumpFileNum
' <VB WATCH>
226        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwCloseDumpFile"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            ' <Dump>
            Err.Number = -1
            Goto vbwErrHandler
    End Select
' </VB WATCH>
End Sub

Private Function GetArrayDimension(ByRef arg As Variant) As Long
       ' vbwNoErrorHandler ' don't remove this !
227        Dim i As Long, j As Long
228        On Error Resume Next
229        i = 0
230        Do
231            i = i + 1
232            j = LBound(arg, i)
233        Loop Until Err.Number
234        GetArrayDimension = i - 1
End Function

Function vbwGetIndex(tObject As Variant) As String
       ' vbwNoErrorHandler ' don't remove this !
235        On Error Resume Next
236        vbwGetIndex = "(" & tObject.Index & ")"
End Function

Private Function FormatString(ByVal arg As String) As String
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
237        On Error GoTo vbwErrHandler
' </VB WATCH>

238        If Right$(arg, 1) = "}" Then ' probably a VB Watch built-in message
239             FormatString = arg
240             Exit Function
241        End If

           ' 1. truncate according to the vbwDumpStringMaxLength value
242        If vbwDumpStringMaxLength Then
243            If Len(arg) > vbwDumpStringMaxLength Then
244                arg = Left$(arg, vbwDumpStringMaxLength + 1)   ' +1: avoids to cut inside a vbCrLf '
245                If Right$(arg, 2) = vbCrLf Then
                       ' don't cut inside a vbCrLf
246                Else
247                    arg = Left$(arg, vbwDumpStringMaxLength)
248                End If
249                arg = arg & "{...}" ' truncated
250            End If
251        End If

           ' 2. make sure string isn't multiline
252        arg = Replace(arg, vbCrLf, "<CrLf>", , , vbBinaryCompare)
253        arg = Replace(arg, Chr(13), "<Cr>", , , vbBinaryCompare)
254        arg = Replace(arg, Chr(10), "<Lf>", , , vbBinaryCompare)

           ' 3. add quotes
255        FormatString = Chr(34) & arg & Chr(34)
' <VB WATCH>
256        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FormatString"

    Select Case vbwErrorHandler(Err.Number, Err.Description, VBWPROJECT, VBWMODULE, VBWPROCEDURE, Erl)
        Case vbwEnd
            End
        Case vbwRetry
            Resume
        Case vbwIgnoreLine
            Resume Next
        Case vbwDoDumpVariable
            ' <Dump>
            Err.Number = -1
            Goto vbwErrHandler
    End Select
' </VB WATCH>
End Function

Sub vbwProcIn(ByRef lProc As String, Optional ByRef lParameters As String)
' <VB WATCH>
257        On Error GoTo vbwErrHandler
' </VB WATCH>

258        vbwTraceCallsNumber = vbwTraceCallsNumber + 1

259        vbwStackCallsNumber = vbwStackCallsNumber + 1
260        ReDim Preserve vbwStackCalls(1 To vbwStackCallsNumber)
261        vbwStackCalls(vbwStackCallsNumber) = lProc

262        Dim lString As String
263        lString = String$(vbwTraceCallsNumber - 1, vbTab) & lProc

264        If vbwLogTraceToFile Then
265             vbwSendLog lString & lParameters
266        End If

' <VB WATCH>
267        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwProcIn"

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
            vbwReportVariable "lProc", lProc
            vbwReportVariable "lParameters", lParameters
            vbwReportVariable "lString", lString
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
' </VB WATCH>
End Sub

Sub vbwProcOut(ByRef lProc As String)
' <VB WATCH>
268        On Error GoTo vbwErrHandler
' </VB WATCH>

269        If vbwTraceCallsNumber > 0 Then ' should always be true
270           vbwTraceCallsNumber = vbwTraceCallsNumber - 1
271        End If

272        If vbwStackCallsNumber > 0 Then ' should always be true
273           vbwStackCallsNumber = vbwStackCallsNumber - 1
274        End If

' <VB WATCH>
275        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwProcOut"

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
            vbwReportVariable "lProc", lProc
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
' </VB WATCH>
End Sub


Function vbwExecuteLine(ByRef fEncrypted As String, ByRef lLine As String) As Boolean
' <VB WATCH>
276        On Error GoTo vbwErrHandler
' </VB WATCH>

277        If vbwTraceLine Then

278            If fEncrypted Then
279                lLine = "<CRY>" & lLine & "</CRY>"
280            End If

281            If vbwLogTraceToFile Then
282                If vbwTraceCallsNumber > 0 Then
283                    vbwSendLog String$(vbwTraceCallsNumber - 1, vbTab) & " -> " & lLine
284                Else
285                    vbwSendLog " -> " & lLine
286                End If
287            End If

288        End If

           ' This function always returns false
' <VB WATCH>
289        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwExecuteLine"

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
            vbwReportVariable "fEncrypted", fEncrypted
            vbwReportVariable "lLine", lLine
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
' </VB WATCH>
End Function

Function vbwGetStack() As String
' <VB WATCH>
290        On Error GoTo vbwErrHandler
' </VB WATCH>

291        If vbwTraceProc = False Then
292            vbwGetStack = "{Unavailable}"
293            Exit Function
294        End If

295        Dim vbwStackString As String
296        Dim i As Long

297        For i = vbwStackCallsNumber To 1 Step -1
298            vbwStackString = vbwStackString & String$(i - 1, vbTab) & vbwStackCalls(i) & vbCrLf
299        Next i
300        vbwGetStack = IIf(vbwStackString <> "", vbwStackString, "{Empty}")
' <VB WATCH>
301        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwGetStack"

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
            vbwReportVariable "vbwStackString", vbwStackString
            vbwReportVariable "i", i
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
' </VB WATCH>
End Function

Sub vbwSendLog(ByRef tMsg As String)
' <VB WATCH>
302        On Error GoTo vbwErrHandler
' </VB WATCH>
303        If Err.Number Then
               ' Save Err object before being cleared by "On Error Resume Next"
304            Dim ErrDescription As String, ErrHelpFile As String, ErrSource As String
305            Dim ErrHelpContext As Long, ErrNumber As Long
306            ErrDescription = Err.Description
307            ErrHelpContext = Err.HelpContext
308            ErrHelpFile = Err.HelpFile
309            ErrNumber = Err.Number
310            ErrSource = Err.Source
311        End If

312        On Error Resume Next

313        If Not fLogFileOpen Then
314            fLogFileOpen = True
315            Dim suffix As Long
316            Do
317                Kill vbwLogFile
318                lLogFileNumber = CreateFile(vbwLogFile, GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, OPEN_ALWAYS, FILE_FLAG_OVERLAPPED, 0)
319                If lLogFileNumber < 0 Then
                       ' under some circumstances (retained in memory applications or while in the IDE)
                       ' the previous log file might not have been freed yet, so we must use another one
320                    suffix = suffix + 1
321                    vbwLogFile = App.Path & "\vbw" & App.EXEName & VBW_EXE_EXTENSION & suffix & ".log"
322                End If
323            Loop Until lLogFileNumber >= 0 Or suffix > 1000
324        End If

325        If Not fIsLogInitialize Then
               ' init file '
326            fIsLogInitialize = True
327            WriteToLogFile "Tracing " & App.Title
328            WriteToLogFile "Session started " & Now
329            WriteToLogFile ""
330        End If

           ' log to file
331        WriteToLogFile tMsg

332       If ErrNumber Then
               ' Restore Err object if cleared by "On Error Resume Next"
333            Err.Description = ErrDescription
334            Err.HelpContext = ErrHelpContext
335            Err.HelpFile = ErrHelpFile
336            Err.Number = ErrNumber
337            Err.Source = ErrSource
338        End If

' <VB WATCH>
339        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwSendLog"

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
            vbwReportVariable "tMsg", tMsg
            vbwReportVariable "ErrDescription", ErrDescription
            vbwReportVariable "ErrHelpFile", ErrHelpFile
            vbwReportVariable "ErrSource", ErrSource
            vbwReportVariable "ErrHelpContext", ErrHelpContext
            vbwReportVariable "ErrNumber", ErrNumber
            vbwReportVariable "suffix", suffix
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
' </VB WATCH>
End Sub

' Writes Str as a new line in the log file (adding a vbCrLf to the end)
Private Function WriteToLogFile(Str As String) As Long
' <VB WATCH>
340        On Error GoTo vbwErrHandler
' </VB WATCH>
341        Dim ol As OVERLAPPED
342        Dim bBytes() As Byte, StrLength As Long
343        StrLength = Len(Str) + 2
344        ReDim bBytes(0 To StrLength - 1)
345        CopyMemory bBytes(0), ByVal Str & vbCrLf, StrLength
346        ol.OffSet = lLogFileOffset
347        WriteToLogFile = WriteFileEx(lLogFileNumber, bBytes(0), StrLength, ol, ByVal 0&)
348        lLogFileOffset = lLogFileOffset + StrLength
' <VB WATCH>
349        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "WriteToLogFile"

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
            vbwReportVariable "Str", Str
            vbwReport_vbwProtector_OVERLAPPED "ol", ol
            vbwReportVariable "bBytes", bBytes
            vbwReportVariable "StrLength", StrLength
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
' </VB WATCH>
End Function


' Ends a component's thread. If this was the last active thread, ends the component's process.
Public Sub vbwExitThread()
' <VB WATCH>
350        On Error GoTo vbwErrHandler
' </VB WATCH>
351        If vbwIsInIDE Then
               ' Executing ExitThread within the IDE will terminate VB without ceremony !
352            Stop ' Press the End button now
353        Else
354            Dim lpExitCode As Long
355            If GetExitCodeThread(GetCurrentThread(), lpExitCode) Then
356                ExitThread lpExitCode
357            End If
358        End If
' <VB WATCH>
359        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwExitThread"

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
            vbwReportVariable "lpExitCode", lpExitCode
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
' </VB WATCH>
End Sub

' Ends a component's process. Equivalent to the End statement.
Public Sub vbwExitProcess()
' <VB WATCH>
360        On Error GoTo vbwErrHandler
' </VB WATCH>
361        If vbwIsInIDE Then
               ' Executing ExitProcess within the IDE will terminate VB without ceremony !
362            Stop ' Press the End button now
363        Else
364            Dim lpExitCode As Long
365            If GetExitCodeProcess(GetCurrentProcess(), lpExitCode) Then
366                ExitProcess lpExitCode
367            End If
368        End If
' <VB WATCH>
369        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwExitProcess"

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
            vbwReportVariable "lpExitCode", lpExitCode
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
' </VB WATCH>
End Sub

' determines if the program is running in the IDE or an EXE File
Private Function vbwIsInIDE() As Boolean
' <VB WATCH>
370        On Error GoTo vbwErrHandler
' </VB WATCH>

371        Dim strFileName As String
372        Dim lngCount As Long

373        strFileName = String(255, 0)
374        lngCount = GetModuleFileName(App.hInstance, strFileName, 255)
375        strFileName = Left(strFileName, lngCount)

376        vbwIsInIDE = UCase$(Right$(strFileName, 8)) Like "\VB#.EXE"

' <VB WATCH>
377        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwIsInIDE"

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
            vbwReportVariable "strFileName", strFileName
            vbwReportVariable "lngCount", lngCount
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
' </VB WATCH>
End Function

' Exception handling stuff
Public Sub vbwHandleException()
           ' Exceptions will be caught and redirected to the failing procedure
' <VB WATCH>
378        On Error GoTo vbwErrHandler
' </VB WATCH>
379        SetUnhandledExceptionFilter AddressOf vbwExceptionFilter
' <VB WATCH>
380        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwHandleException"

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
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
' </VB WATCH>
End Sub

' Exception handling stuff
Public Sub vbwUnHandleException()
           ' Exceptions are no longer caught and will cause Exceptions
           ' Whenever possible, call this procedure before returning to the VB's IDE
' <VB WATCH>
381        On Error GoTo vbwErrHandler
' </VB WATCH>
382        SetUnhandledExceptionFilter 0
' <VB WATCH>
383        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "vbwUnHandleException"

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
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
' </VB WATCH>
End Sub

' Exception handling stuff
Public Function vbwExceptionFilter(ByRef pExceptionInfo As EXCEPTION_POINTERS) As Long
       'vbwNoErrorHandler ' DO NOT remove this !!!

384        Dim ExceptionRecord As EXCEPTION_RECORD
385        ExceptionRecord = pExceptionInfo.pExceptionRecord

386        Do While ExceptionRecord.pExceptionRecord ' Empties the exceptions stack
387            CopyMemory ExceptionRecord, ByVal ExceptionRecord.pExceptionRecord, Len(ExceptionRecord)
388        Loop

389        vbwExceptionFilter = EXCEPTION_CONTINUE_EXECUTION

       'vbwExitProc ' because the next instruction causes to exit the function ' ' DO NOT remove this !!!

           ' Convert the exception to a normal VB error and go back to the failing procedure '
390        Err.Raise 65535, , ExceptionDescription(ExceptionRecord.ExceptionCode)

End Function

' Exception handling stuff
Private Function ExceptionDescription(ByVal ExceptionCode As Long) As String
       ' vbwNoErrorHandler ' don't remove this !
391        Select Case ExceptionCode
               Case EXCEPTION_ACCESS_VIOLATION
392                ExceptionDescription = "Exception: Access Violation"
393            Case EXCEPTION_DATATYPE_MISALIGNMENT
394                ExceptionDescription = "Exception: Datatype Misalignment"
395            Case EXCEPTION_BREAKPOINT
396                ExceptionDescription = "Exception: Breakpoint"
397            Case EXCEPTION_SINGLE_STEP
398                ExceptionDescription = "Exception: Single Step"
399            Case EXCEPTION_ARRAY_BOUNDS_EXCEEDED
400                ExceptionDescription = "Exception: Array Bounds Exceeded"
401            Case EXCEPTION_FLT_DENORMAL_OPERAND
402                ExceptionDescription = "Exception: Float Denormal Operand"
403            Case EXCEPTION_FLT_DIVIDE_BY_ZERO
404                ExceptionDescription = "Exception: Float Divide By Zero"
405            Case EXCEPTION_FLT_INEXACT_RESULT
406                ExceptionDescription = "Exception: Float Inexact Result"
407            Case EXCEPTION_FLT_INVALID_OPERATION
408                ExceptionDescription = "Exception: Float Invalid Operation"
409            Case EXCEPTION_FLT_OVERFLOW
410                ExceptionDescription = "Exception: Float Overflow"
411            Case EXCEPTION_FLT_STACK_CHECK
412                ExceptionDescription = "Exception: Float Stack Check"
413            Case EXCEPTION_FLT_UNDERFLOW
414                ExceptionDescription = "Exception: Float Underflow"
415            Case EXCEPTION_INT_DIVIDE_BY_ZERO
416                ExceptionDescription = "Exception: Integer Divide By Zero"
417            Case EXCEPTION_INT_OVERFLOW
418                ExceptionDescription = "Exception: Integer Overflow"
419            Case EXCEPTION_PRIV_INSTRUCTION
420                ExceptionDescription = "Exception: Priv Instruction"
421            Case EXCEPTION_IN_PAGE_ERROR
422                ExceptionDescription = "Exception: In Page Error"
423            Case EXCEPTION_ILLEGAL_INSTRUCTION
424                ExceptionDescription = "Exception: Illegal Instruction"
425            Case EXCEPTION_NONCONTINUABLE_EXCEPTION
426                ExceptionDescription = "Exception: Non Continuable Exception"
427            Case EXCEPTION_STACK_OVERFLOW
428                ExceptionDescription = "Exception: Stack Overflow"
429            Case EXCEPTION_INVALID_DISPOSITION
430                ExceptionDescription = "Exception: Invalid Disposition"
431            Case EXCEPTION_GUARD_PAGE
432                ExceptionDescription = "Exception: Guard Page"
433            Case EXCEPTION_INVALID_HANDLE
434                ExceptionDescription = "Exception: Invalid Handle"
435            Case CONTROL_C_EXIT
436                ExceptionDescription = "Exception: Control C Exit"
437            Case Else
438                ExceptionDescription = "Unknown Exception"
439        End Select

End Function

Public Sub vbwSaveErrObject()
       ' vbwNoErrorHandler ' don't remove this !
440        ErrObjectDescription = Err.Description
441        ErrObjectHelpContext = Err.HelpContext
442        ErrObjectHelpFile = Err.HelpFile
443        ErrObjectLastDllError = Err.LastDllError
444        ErrObjectNumber = Err.Number
445        ErrObjectSource = Err.Source
End Sub

Public Sub vbwRestoreErrObject()
       ' vbwNoErrorHandler ' don't remove this !
446       Err.Description = ErrObjectDescription
447       Err.HelpContext = ErrObjectHelpContext
448       Err.HelpFile = ErrObjectHelpFile
449       Err.Number = ErrObjectNumber
450       Err.Source = ErrObjectSource
End Sub


' <VB WATCH> <VBWATCHFINALPROC>
' Procedures added by VB Watch for variable dump
Private Sub vbwReport_vbwProtector_OVERLAPPED(lName As String, lUDT As vbwProtector.OVERLAPPED, Optional ByVal lTab As Long)
    vbwReportVariable lName & ".Internal", lUDT.Internal, lTab
    vbwReportVariable lName & ".InternalHigh", lUDT.InternalHigh, lTab
    vbwReportVariable lName & ".OffSet", lUDT.OffSet, lTab
    vbwReportVariable lName & ".OffsetHigh", lUDT.OffsetHigh, lTab
    vbwReportVariable lName & ".hEvent", lUDT.hEvent, lTab
End Sub
Private Sub vbwReport_vbwProtector_EXCEPTION_RECORD(lName As String, lUDT As vbwProtector.EXCEPTION_RECORD, Optional ByVal lTab As Long)
    vbwReportVariable lName & ".ExceptionCode", lUDT.ExceptionCode, lTab
    vbwReportVariable lName & ".ExceptionFlags", lUDT.ExceptionFlags, lTab
    vbwReportVariable lName & ".pExceptionRecord", lUDT.pExceptionRecord, lTab
    vbwReportVariable lName & ".ExceptionAddress", lUDT.ExceptionAddress, lTab
    vbwReportVariable lName & ".NumberParameters", lUDT.NumberParameters, lTab
    vbwReportVariable lName & ".ExceptionInformation", lUDT.ExceptionInformation, lTab
End Sub
Private Sub vbwReport_vbwProtector_EXCEPTION_DEBUG_INFO(lName As String, lUDT As vbwProtector.EXCEPTION_DEBUG_INFO, Optional ByVal lTab As Long)
    vbwReport_vbwProtector_EXCEPTION_RECORD lName & ".pExceptionRecord", lUDT.pExceptionRecord, lTab
    vbwReportVariable lName & ".dwFirstChance", lUDT.dwFirstChance, lTab
End Sub
Private Sub vbwReport_vbwProtector_CONTEXT(lName As String, lUDT As vbwProtector.CONTEXT, Optional ByVal lTab As Long)
    vbwReportVariable lName & ".dblVar", lUDT.dblVar, lTab
    vbwReportVariable lName & ".lngVar", lUDT.lngVar, lTab
End Sub
Private Sub vbwReport_vbwProtector_EXCEPTION_POINTERS(lName As String, lUDT As vbwProtector.EXCEPTION_POINTERS, Optional ByVal lTab As Long)
    vbwReport_vbwProtector_EXCEPTION_RECORD lName & ".pExceptionRecord", lUDT.pExceptionRecord, lTab
    vbwReport_vbwProtector_CONTEXT lName & ".ContextRecord", lUDT.ContextRecord, lTab
End Sub

Private Sub vbwReportModuleVariables()
    vbwReportToFile VBW_MODULE_STRING
    vbwReportVariable "fIsLogInitialize", fIsLogInitialize
    vbwReportVariable "fLogFileOpen", fLogFileOpen
    vbwReportVariable "lLogFileNumber", lLogFileNumber
    vbwReportVariable "lLogFileOffset", lLogFileOffset
    vbwReportVariable "ErrObjectDescription", ErrObjectDescription
    vbwReportVariable "ErrObjectHelpContext", ErrObjectHelpContext
    vbwReportVariable "ErrObjectHelpFile", ErrObjectHelpFile
    vbwReportVariable "ErrObjectLastDllError", ErrObjectLastDllError
    vbwReportVariable "ErrObjectNumber", ErrObjectNumber
    vbwReportVariable "ErrObjectSource", ErrObjectSource
    vbwReportVariable "ErrLine", ErrLine
End Sub
Public Sub vbwReport_AccessRoutines_DataResult(lName As String, lUDT As AccessRoutines.DataResult, Optional ByVal lTab As Long)
    vbwReportVariable lName & ".HP", lUDT.HP, lTab
    vbwReportVariable lName & ".Speed", lUDT.Speed, lTab
End Sub
Public Sub vbwReport_AccessRoutines_DataResult_Array(lName As String, lArray() As AccessRoutines.DataResult, Optional ByVal lTab As Long)

    ' get array dimension number
    Dim i As Long, j As Long, k As Long, l As Long
    Dim DimNumber As Long
    On Error Resume Next
    DimNumber = 0
    Do
        DimNumber = DimNumber + 1
        j = LBound(lArray, DimNumber)
    Loop Until Err.Number
    DimNumber = DimNumber - 1

    ' report each member
    Select Case DimNumber
        Case 1
            vbwReportToFile String$(lTab, vbTab) & lName & VBW_TYPE_STRING
            For i = LBound(lArray) To UBound(lArray)
                vbwReport_AccessRoutines_DataResult lName & "(" & i & ")", lArray(i), lTab + 1
            Next i
        Case 2
            vbwReportToFile String$(lTab, vbTab) & lName & VBW_TYPE_STRING
            For j = LBound(lArray, 2) To UBound(lArray, 2)
                For i = LBound(lArray, 1) To UBound(lArray, 1)
                    vbwReport_AccessRoutines_DataResult lName & "(" & i & "," & j & ")", lArray(i, j), lTab + 1
                Next i
            Next j
        Case 3
            vbwReportToFile String$(lTab, vbTab) & lName & VBW_TYPE_STRING
            For k = LBound(lArray, 3) To UBound(lArray, 3)
                For j = LBound(lArray, 2) To UBound(lArray, 2)
                    For i = LBound(lArray, 1) To UBound(lArray, 1)
                        vbwReport_AccessRoutines_DataResult lName & "(" & i & "," & j & "," & k & ")", lArray(i, j, k), lTab + 1
                    Next i
                Next j
            Next k
        Case 4
            vbwReportToFile String$(lTab, vbTab) & lName & VBW_TYPE_STRING
             For l = LBound(lArray, 4) To UBound(lArray, 4)
                For k = LBound(lArray, 3) To UBound(lArray, 3)
                    For j = LBound(lArray, 2) To UBound(lArray, 2)
                        For i = LBound(lArray, 1) To UBound(lArray, 1)
                            vbwReport_AccessRoutines_DataResult lName & "(" & i & "," & j & "," & k & "," & l & ")", lArray(i, j, k, l), lTab + 1
                        Next i
                    Next j
                Next k
            Next l
        Case Else
            vbwReportToFile lName & "() not processed: " & DimNumber & " dimensions" & VBW_TYPE_STRING
    End Select

End Sub
 
Public Sub vbwReport_AccessRoutines_DataSet(lName As String, lUDT As AccessRoutines.DataSet, Optional ByVal lTab As Long)
    vbwReportVariable lName & ".Flow", lUDT.Flow, lTab
    vbwReportVariable lName & ".SuctionPressure", lUDT.SuctionPressure, lTab
    vbwReportVariable lName & ".DischargePressure", lUDT.DischargePressure, lTab
    vbwReportVariable lName & ".Temperature", lUDT.Temperature, lTab
    vbwReportVariable lName & ".SuctionPipeDia", lUDT.SuctionPipeDia, lTab
    vbwReportVariable lName & ".DischargePipeDia", lUDT.DischargePipeDia, lTab
    vbwReportVariable lName & ".SuctionHeight", lUDT.SuctionHeight, lTab
    vbwReportVariable lName & ".DischargeHeight", lUDT.DischargeHeight, lTab
    vbwReportVariable lName & ".BarometricPressure", lUDT.BarometricPressure, lTab
    vbwReportVariable lName & ".HDCorr", lUDT.HDCorr, lTab
    vbwReportVariable lName & ".SuctionInHg", lUDT.SuctionInHg, lTab
    vbwReportVariable lName & ".MotorType", lUDT.MotorType, lTab
    vbwReportVariable lName & ".StatorFill", lUDT.StatorFill, lTab
    vbwReportVariable lName & ".VoltageA", lUDT.VoltageA, lTab
    vbwReportVariable lName & ".VoltageB", lUDT.VoltageB, lTab
    vbwReportVariable lName & ".VoltageC", lUDT.VoltageC, lTab
    vbwReportVariable lName & ".CurrentA", lUDT.CurrentA, lTab
    vbwReportVariable lName & ".CurrentB", lUDT.CurrentB, lTab
    vbwReportVariable lName & ".CurrentC", lUDT.CurrentC, lTab
    vbwReportVariable lName & ".PowerA", lUDT.PowerA, lTab
    vbwReportVariable lName & ".PowerB", lUDT.PowerB, lTab
    vbwReportVariable lName & ".PowerC", lUDT.PowerC, lTab
    vbwReportVariable lName & ".PowerFactor", lUDT.PowerFactor, lTab
    vbwReportVariable lName & ".VelocityHead", lUDT.VelocityHead, lTab
    vbwReportVariable lName & ".TDH", lUDT.TDH, lTab
    vbwReportVariable lName & ".OverallEfficiency", lUDT.OverallEfficiency, lTab
    vbwReportVariable lName & ".MotorEfficiency", lUDT.MotorEfficiency, lTab
    vbwReportVariable lName & ".HydraulicEfficiency", lUDT.HydraulicEfficiency, lTab
    vbwReportVariable lName & ".CalcPowerFactor", lUDT.CalcPowerFactor, lTab
    vbwReportVariable lName & ".CalcVelocityHead", lUDT.CalcVelocityHead, lTab
    vbwReportVariable lName & ".CalcTDH", lUDT.CalcTDH, lTab
    vbwReportVariable lName & ".CalcOverallEfficiency", lUDT.CalcOverallEfficiency, lTab
    vbwReportVariable lName & ".CalcMotorEfficiency", lUDT.CalcMotorEfficiency, lTab
    vbwReportVariable lName & ".CalcHydraulicEfficiency", lUDT.CalcHydraulicEfficiency, lTab
End Sub
Public Sub vbwReport_AccessRoutines_DataSet_Array(lName As String, lArray() As AccessRoutines.DataSet, Optional ByVal lTab As Long)

    Dim i As Long

    ' report each member '
    vbwReportToFile String$(lTab, vbTab) & lName & VBW_TYPE_STRING
    For i = LBound(lArray) To UBound(lArray)
        vbwReport_AccessRoutines_DataSet lName & "(" & i & ")", lArray(i), lTab + 1
    Next i

End Sub
  
Public Sub vbwReport_EpicorRoutines_SNRecord(lName As String, lUDT As EpicorRoutines.SNRecord, Optional ByVal lTab As Long)
    vbwReportVariable lName & ".SONumber", lUDT.SONumber, lTab
    vbwReportVariable lName & ".SOLine", lUDT.SOLine, lTab
    vbwReportVariable lName & ".ModelNo", lUDT.ModelNo, lTab
    vbwReportVariable lName & ".MotorSize", lUDT.MotorSize, lTab
    vbwReportVariable lName & ".PartNum", lUDT.PartNum, lTab
    vbwReportVariable lName & ".Customer", lUDT.Customer, lTab
    vbwReportVariable lName & ".ShipTo", lUDT.ShipTo, lTab
    vbwReportVariable lName & ".CustNum", lUDT.CustNum, lTab
    vbwReportVariable lName & ".ShipToNum", lUDT.ShipToNum, lTab
    vbwReportVariable lName & ".TDH", lUDT.TDH, lTab
    vbwReportVariable lName & ".Flow", lUDT.Flow, lTab
    vbwReportVariable lName & ".ImpellerDiameter", lUDT.ImpellerDiameter, lTab
    vbwReportVariable lName & ".SuctionPressure", lUDT.SuctionPressure, lTab
    vbwReportVariable lName & ".SpGr", lUDT.SpGr, lTab
    vbwReportVariable lName & ".Fluid", lUDT.Fluid, lTab
    vbwReportVariable lName & ".PumpTemperature", lUDT.PumpTemperature, lTab
    vbwReportVariable lName & ".Viscosity", lUDT.Viscosity, lTab
    vbwReportVariable lName & ".VaporPressure", lUDT.VaporPressure, lTab
    vbwReportVariable lName & ".SuctFlangeSize", lUDT.SuctFlangeSize, lTab
    vbwReportVariable lName & ".DischFlangeSize", lUDT.DischFlangeSize, lTab
    vbwReportVariable lName & ".RPM", lUDT.RPM, lTab
    vbwReportVariable lName & ".Voltage", lUDT.Voltage, lTab
    vbwReportVariable lName & ".StatorFill", lUDT.StatorFill, lTab
    vbwReportVariable lName & ".CirculationPath", lUDT.CirculationPath, lTab
    vbwReportVariable lName & ".TestProcedure", lUDT.TestProcedure, lTab
    vbwReportVariable lName & ".DesignPressure", lUDT.DesignPressure, lTab
    vbwReportVariable lName & ".Frequency", lUDT.Frequency, lTab
    vbwReportVariable lName & ".XPartNum", lUDT.XPartNum, lTab
    vbwReportVariable lName & ".Phases", lUDT.Phases, lTab
    vbwReportVariable lName & ".NPSHr", lUDT.NPSHr, lTab
    vbwReportVariable lName & ".RatedInputPower", lUDT.RatedInputPower, lTab
    vbwReportVariable lName & ".FLCurrent", lUDT.FLCurrent, lTab
    vbwReportVariable lName & ".ThermalClass", lUDT.ThermalClass, lTab
    vbwReportVariable lName & ".ExpClass", lUDT.ExpClass, lTab
    vbwReportVariable lName & ".LiquidTemp", lUDT.LiquidTemp, lTab
    vbwReportVariable lName & ".JobNumber", lUDT.JobNumber, lTab
    vbwReportVariable lName & ".CustomerPO", lUDT.CustomerPO, lTab
End Sub

Public Sub vbwReportGlobalVariables()
    vbwReportToFile VBW_GLOBAL_STRING
    vbwReportVariable "vbwProtector.vbwCatchException", vbwCatchException
    vbwReportVariable "vbwProtector.vbwTraceProc", vbwTraceProc
    vbwReportVariable "vbwProtector.vbwTraceParameters", vbwTraceParameters
    vbwReportVariable "vbwProtector.vbwTraceLine", vbwTraceLine
    vbwReportVariable "vbwProtector.vbwCallStack", vbwCallStack
    vbwReportVariable "vbwProtector.vbwEmailRecipientAdress", vbwEmailRecipientAdress
    vbwReportVariable "vbwProtector.vbwDumpStringMaxLength", vbwDumpStringMaxLength
    vbwReportVariable "vbwProtector.vbwSystemInfo", vbwSystemInfo
    vbwReportVariable "vbwProtector.vbwScreenshot", vbwScreenshot
    vbwReportVariable "vbwProtector.fIsVbwFunctionsInitialized", fIsVbwFunctionsInitialized
    vbwReportVariable "vbwProtector.vbwStackCalls", vbwStackCalls
    vbwReportVariable "vbwProtector.vbwStackCallsNumber", vbwStackCallsNumber
    vbwReportVariable "vbwProtector.vbwTraceCallsNumber", vbwTraceCallsNumber
    vbwReportVariable "vbwProtector.vbwLogFile", vbwLogFile
    vbwReportVariable "vbwProtector.vbwLogTraceToFile", vbwLogTraceToFile
    vbwReportVariable "vbwProtector.vbwDumpFile", vbwDumpFile
    vbwReportVariable "vbwProtector.vbwDumpFileNum", vbwDumpFileNum
    vbwReportVariable "vbwProtector.VBWPROTECTOR_EMPTY", VBWPROTECTOR_EMPTY
    vbwReportVariable "vbwErrHandler.vbwRetCode", vbwRetCode
    vbwReportVariable "vbwErrHandler.vbwMessageString", vbwMessageString
    vbwReportVariable "vbwErrHandler.vbwCircumstancesString", vbwCircumstancesString
    vbwReportVariable "vbwErrHandler.vbwErrorPath", vbwErrorPath
    vbwReportVariable "vbwErrHandler.vbwfHasReported", vbwfHasReported
    vbwReportVariable "AccessRoutines.boCanApprove", boCanApprove
    vbwReportVariable "AccessRoutines.LogInInitials", LogInInitials
    vbwReport_AccessRoutines_DataResult_Array "AccessRoutines.results", results
    vbwReport_AccessRoutines_DataSet_Array "AccessRoutines.DataSets", DataSets
    vbwReport_AccessRoutines_DataSet "AccessRoutines.UseDataset", UseDataset
    vbwReportVariable "AccessRoutines.Calibrating", Calibrating
    vbwReportVariable "AccessRoutines.sServerName", sServerName
    vbwReportVariable "AccessRoutines.sCalibrateDatabaseName", sCalibrateDatabaseName
    vbwReportVariable "AccessRoutines.sCalibrateSaveFileName", sCalibrateSaveFileName
    vbwReportVariable "AccessRoutines.CalibrateWorkSheetName", CalibrateWorkSheetName
    vbwReportVariable "AccessRoutines.WritingToCalFile", WritingToCalFile
    vbwReportVariable "AccessRoutines.PipeDiameters", PipeDiameters
    vbwReportVariable "AccessRoutines.VaporPressure", VaporPressure
    vbwReportVariable "AccessRoutines.TempCorrection", TempCorrection
    vbwReportVariable "AccessRoutines.TEMCForceViscosity", TEMCForceViscosity
    vbwReportVariable "PLCInterface.rc", rc
    vbwReport_PLCInterface_HEITransport "PLCInterface.TP", TP
    vbwReportVariable "PLCInterface.NetworkOK", NetworkOK
    vbwReport_PLCInterface_HEIDevice_Array "PLCInterface.aDevices", aDevices
    vbwReportVariable "PLCInterface.DeviceCount", DeviceCount
    vbwReportVariable "PLCInterface.DeviceOpen", DeviceOpen
    vbwReportVariable "PLCInterface.tDevice", tDevice
    vbwReportVariable "PLCInterface.bWrite", bWrite
    vbwReportVariable "PLCInterface.DataType", DataType
    vbwReportVariable "PLCInterface.DataAddress", DataAddress
    vbwReportVariable "PLCInterface.DataLength", DataLength
    vbwReportVariable "PLCInterface.ByteBuffer", ByteBuffer
    vbwReportVariable "PLCInterface.Description", Description
    vbwReportVariable "NIGLOBAL.ibsta", ibsta
    vbwReportVariable "NIGLOBAL.iberr", iberr
    vbwReportVariable "NIGLOBAL.ibcnt", ibcnt
    vbwReportVariable "NIGLOBAL.ibcntl", ibcntl
    vbwReportVariable "NIGLOBAL.Longibsta", Longibsta
    vbwReportVariable "NIGLOBAL.Longiberr", Longiberr
    vbwReportVariable "NIGLOBAL.Longibcnt", Longibcnt
    vbwReportVariable "NIGLOBAL.GPIBglobalsRegistered", GPIBglobalsRegistered
    vbwReportVariable "NIGLOBAL.buf", buf
    vbwReportVariable "NIGLOBAL.bytebuf", bytebuf
    vbwReportVariable "GetUNC.WroteNPSHr", WroteNPSHr
    vbwReportVariable "MagtrolRoutinesProLogix.vResponse", vResponse
    vbwReportVariable "MagtrolRoutinesProLogix.sData", sData
    vbwReportVariable "MagtrolRoutinesProLogix.iUD", iUD
    vbwReportVariable "MagtrolRoutinesProLogix.vPlot", vPlot
    vbwReportVariable "vbwProtector.vbwAdvancedFunctions", vbwAdvancedFunctions
    vbwReportVariable "AccessRoutines.cnPumpData", cnPumpData
    vbwReportVariable "AccessRoutines.cnEffData", cnEffData
    vbwReportVariable "AccessRoutines.cnCalibrate", cnCalibrate
    vbwReportVariable "AccessRoutines.rsCalibrate", rsCalibrate
    vbwReportVariable "AccessRoutines.xlApp", xlApp
    vbwReportVariable "AccessRoutines.xlBook", xlBook
    vbwReportVariable "MagtrolRoutinesProLogix.TCP5", TCP5
    vbwReportVariable "MagtrolRoutinesProLogix.TCP6", TCP6
    vbwReportVariable "MagtrolRoutinesProLogix.TCP", TCP
    Dim f As Form
    For Each f In Forms
        vbwReportObject f.Name, f
    Next f

    vbwReportObject "App", App
    vbwReportObject "Err", Err
    vbwReportObject "Screen", Screen
    vbwReportObject "Printer", Printer
End Sub
' </VB WATCH>
