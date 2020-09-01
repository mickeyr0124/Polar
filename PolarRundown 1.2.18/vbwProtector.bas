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
       ' Generated by VB Watch 2 Wizard 1/17/2018 10:18:13 AM
6      vbwCatchException = False
7      vbwSystemInfo = False
8      vbwScreenshot = False
9      vbwEmailRecipientAdress = "mrosenbaum@teikokupumps.com"


10         vbwLogTraceToFile = vbwTraceProc Or vbwTraceLine
11         If vbwCallStack Then ' needed to track call stack
12              vbwTraceProc = True
13         End If

14         vbwLogFile = Replace(App.Path & "\", ":\\", ":\") &  "vbw" & App.EXEName & VBW_EXE_EXTENSION & ".log"
15         vbwDumpFile = Replace(App.Path & "\", ":\\", ":\") &  "vbw" & App.EXEName & VBW_EXE_EXTENSION & ".dmp"

16         If vbwCatchException Then
17             vbwHandleException
18         End If

19         vbwDumpStringMaxLength = 128 ' change this value to suit your need - make it 0 to remove the size check (to use with caution)

20         vbwIsInitialized = True

' <VB WATCH>
21         Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Sub vbwReportVariable(ByVal lName As String, ByVal lValue As Variant, Optional ByVal lTab As Long)
       ' vbwNoErrorHandler ' don't remove this !
22         Dim i As Long, j As Long, k As Long, L As Long
23         Dim tDim As Long

24         On Error GoTo ErrDump

25         If InStr(1, TypeName(lValue), "()", vbBinaryCompare) Then
               ' array '
26             tDim = GetArrayDimension(lValue)
27             Select Case tDim
                   Case 1
28                     vbwReportToFile String$(lTab, vbTab) & vbwEncryptString("Array " & lName & "(" & LBound(lValue, 1) & " To " & UBound(lValue, 1) & ") As " & TypeName(lValue))
29                     For i = LBound(lValue) To UBound(lValue)
30                         vbwReportVariable lName & "(" & i & ")", lValue(i), lTab + 1
31                     Next i
32                 Case 2
33                     vbwReportToFile String$(lTab, vbTab) & vbwEncryptString("Array " & lName & "(" & LBound(lValue, 1) & " To " & UBound(lValue, 1) & "," & LBound(lValue, 2) & " To " & UBound(lValue, 2) & ") As " & TypeName(lValue))
34                     For j = LBound(lValue, 2) To UBound(lValue, 2)
35                         For i = LBound(lValue, 1) To UBound(lValue, 1)
36                             vbwReportVariable lName & "(" & i & "," & j & ")", lValue(i, j), lTab + 1
37                         Next i
38                     Next j
39                 Case 3
40                     vbwReportToFile String$(lTab, vbTab) & vbwEncryptString("Array " & lName & "(" & LBound(lValue, 1) & " To " & UBound(lValue, 1) & "," & LBound(lValue, 2) & " To " & UBound(lValue, 2) & "," & LBound(lValue, 3) & " To " & UBound(lValue, 3) & ") As " & TypeName(lValue))
41                     For k = LBound(lValue, 3) To UBound(lValue, 3)
42                         For j = LBound(lValue, 2) To UBound(lValue, 2)
43                             For i = LBound(lValue, 1) To UBound(lValue, 1)
44                                 vbwReportVariable lName & "(" & i & "," & j & "," & k & ")", lValue(i, j, k), lTab + 1
45                             Next i
46                         Next j
47                     Next k
48                 Case 4
49                     vbwReportToFile String$(lTab, vbTab) & vbwEncryptString("Array " & lName & "(" & LBound(lValue, 1) & " To " & UBound(lValue, 1) & "," & LBound(lValue, 2) & " To " & UBound(lValue, 2) & "," & LBound(lValue, 3) & " To " & UBound(lValue, 3) & "," & LBound(lValue, 4) & " To " & UBound(lValue, 4) & ") As " & TypeName(lValue))
50                     For L = LBound(lValue, 4) To UBound(lValue, 4)
51                         For k = LBound(lValue, 3) To UBound(lValue, 3)
52                             For j = LBound(lValue, 2) To UBound(lValue, 2)
53                                 For i = LBound(lValue, 1) To UBound(lValue, 1)
54                                     vbwReportVariable lName & "(" & i & "," & j & "," & k & "," & L & ")", lValue(i, j, k, L), lTab + 1
55                                 Next i
56                             Next j
57                         Next k
58                     Next L
59                 Case Else
60                     vbwReportToFile String$(lTab, vbTab) & vbwEncryptString("Array " & lName & "() not processed: " & tDim & " dimensions")
61             End Select
62         Else
               ' non-array '
63             If IsObject(lValue) Then
64                 vbwReportObject lName, lValue, lTab
65             Else
66                 If VarType(lValue) = vbString Then
67                     lValue = FormatString(lValue)
68                 End If
69                 vbwReportToFile String$(lTab, vbTab) & vbwEncryptString(lName & " = " & lValue & " (" & TypeName(lValue) & ")")
70             End If
71         End If
72         Exit Sub

73     ErrDump:
74         Err.Clear
75         vbwReportToFile String$(lTab, vbTab) & vbwEncryptString(lName & " = {Variable Dumping Error}")
End Sub

Public Sub vbwReportObject(lName As String, ByVal lObject As Object, Optional ByVal lTab As Long)
       ' vbwNoErrorHandler ' don't remove this !

76         On Error GoTo ErrDump

77         If TypeName(lObject) <> "ErrObject" Then
78             If fIsVbwFunctionsInitialized Then
                   ' this should be executed only if you are using a global error handler '
                   ' that prepares properly the vbwAdvancedFunctions for object dumping   '
79                 vbwCloseDumpFile       ' close it because vbwAdvancedFunctions uses its own file writing routines '
80                 vbwAdvancedFunctions.ReportObject lName, lObject, lTab, TypeOf lObject Is Form, TypeOf lObject Is MDIForm
81                 vbwOpenDumpFile
82             Else
                   ' no vbwFunctions.dll available                         '
                   ' only report the default value of objects and controls '
83                 If TypeOf lObject Is Form Or TypeOf lObject Is MDIForm Then
84                    On Error Resume Next
85                    vbwReportToFile vbwEncryptString("Form " & lName)
86                    Dim c As Control
87                    For Each c In lObject.Controls
88                        vbwReportObject c.Name & vbwGetIndex(c), c, 1
89                    Next c
90                 Else
91                     If IsNumeric(lObject) Then
92                         vbwReportVariable lName, CDbl(lObject), lTab
93                     Else
94                         vbwReportVariable lName, CStr(lObject), lTab
95                     End If
96                 End If
97             End If
98         Else
99             vbwReportToFile vbCrLf & vbwEncryptString("**** ErrObject Err ****")
100            vbwReportVariable "Err.Number", ErrObjectNumber
101            vbwReportVariable "Err.Source", ErrObjectSource
102            vbwReportVariable "Err.Description", ErrObjectDescription
103            vbwReportVariable "Err.HelpContext", ErrObjectHelpContext
104            vbwReportVariable "Err.HelpFile", ErrObjectHelpFile
105            If ErrObjectLastDllError = 0 Then
106                vbwReportVariable "Err.LastDllError", ErrObjectLastDllError
107            Else
                   ' get the API error description from the system
108                Dim sBuffer As String * 512
109                Const FORMAT_MESSAGE_FROM_SYSTEM = &H1000
110                FormatMessage FORMAT_MESSAGE_FROM_SYSTEM, Null, ErrObjectLastDllError, 0, sBuffer, 512, 0
111                If InStr(sBuffer, Chr(0)) Then
112                    vbwReportVariable "Err.LastDllError", ErrObjectLastDllError & " (" & Left$(sBuffer, InStr(sBuffer, Chr(0)) - 1) & ")"
113                Else
114                    vbwReportVariable "Err.LastDllError", ErrObjectLastDllError
115                End If
116            End If
117        End If

118        Exit Sub

119    ErrDump:
120        Err.Clear
121        vbwReportToFile String$(lTab, vbTab) & vbwEncryptString(lName & ".Value = {No Value Property}")
End Sub

Public Function vbwEncryptString(ByRef sString As String, Optional sKey) As String
       ' vbwNoErrorHandler ' don't remove this ! '

122        On Error Resume Next
123        If fIsVbwFunctionsInitialized = False Then
               ' no encryption without vbwFunctions.dll            '
               ' you may want to write your own encryption routine '
124            vbwEncryptString = sString
125        Else
126            If IsMissing(sKey) Then
                   ' If you filled the vardump encryption key in the VB Watch Options, your key will   '
                   ' be already embeded in the vbwAdvancedFunctions.ObjectInfo property, so you do not '
                   ' have to care about provideing a key                                               '
127                vbwEncryptString = vbwAdvancedFunctions.EncryptString(sString)
128            Else
                   ' Yet if you wish to overide  your default encryption key, simply pass it '
                   ' in the sKey parameter                                                   '
129                vbwEncryptString = vbwAdvancedFunctions.EncryptString(sString, sKey)
130            End If
131        End If
End Function

Function vbwReportParameter(ByVal lName As String, ByRef lValue As Variant) As String
       ' vbwNoErrorHandler ' don't remove this !
132        Dim i As Long, j As Long, k As Long
133        Dim tDim As Long
134        Dim retString As String

135        On Error GoTo ErrDump

136        If InStr(1, TypeName(lValue), "()", vbBinaryCompare) Then
               ' array '
137            tDim = GetArrayDimension(lValue)
138            If tDim Then
139                retString = lName & "("
140                For i = 1 To tDim
141                    retString = retString & LBound(lValue, i) & " To " & UBound(lValue, i) & ","
142                Next i
143                Mid$(retString, Len(retString)) = ")"   ' Close the brackets by overwriting the last comma '
144            Else
145                retString = lName & "(Undimensioned Array)"
146            End If
147        Else
               ' non-array '
148            If IsObject(lValue) Then
                   ' object
149                On Error Resume Next
150                retString = TypeName(lValue) & " " & lName & " = " & CStr(lValue)
151                If Err.Number Then
152                    On Error GoTo ErrDump
153                    retString = TypeName(lValue) & " " & lName & " = " & lValue.Name & vbwGetIndex(lValue)
154                End If
155            Else
                   ' non-object
156                If VarType(lValue) = vbString Then
157                   retString = lName & " = " & FormatString(lValue)
158                Else
159                   retString = lName & " = " & lValue
160                End If
161            End If
162        End If

163        vbwReportParameter = retString
164        Exit Function

165    ErrDump:
166        Err.Clear
167        vbwReportParameter = lName & " = {" & TypeName(lValue) & ": Parameter Dumping Error}"
End Function

Function vbwReportParameterByVal(ByVal lName As String, ByVal lValue As Variant) As String
       ' vbwNoErrorHandler ' don't remove this !
168        Dim i As Long, j As Long, k As Long
169        Dim tDim As Long
170        Dim retString As String

171        On Error GoTo ErrDump

172        If InStr(1, TypeName(lValue), "()", vbBinaryCompare) Then
               ' array '
173            tDim = GetArrayDimension(lValue)
174            If tDim Then
175                retString = lName & "("
176                For i = 1 To tDim
177                    retString = retString & LBound(lValue, i) & " To " & UBound(lValue, i) & ","
178                Next i
179                Mid$(retString, Len(retString)) = ")"   ' Close the brackets by overwriting the last comma '
180            Else
181                retString = lName & "(Undimensioned Array)"
182            End If
183        Else
               ' non-array '
184            If IsObject(lValue) Then
                   ' object
185                On Error Resume Next
186                retString = TypeName(lValue) & " " & lName & " = " & CStr(lValue)
187                If Err.Number Then
188                    On Error GoTo ErrDump
189                    retString = TypeName(lValue) & " " & lName & " = " & lValue.Name & vbwGetIndex(lValue)
190                End If
191            Else
                   ' non-object
192                If VarType(lValue) = vbString Then
193                   retString = lName & " = " & FormatString(lValue)
194                Else
195                   retString = lName & " = " & lValue
196                End If
197            End If
198        End If

199        vbwReportParameterByVal = retString
200        Exit Function

201    ErrDump:
202        Err.Clear
203        vbwReportParameterByVal = lName & " = {" & TypeName(lValue) & ": Parameter Dumping Error}"
End Function

Sub vbwReportToFile(ByRef lString As String)
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
204        On Error GoTo vbwErrHandler
' </VB WATCH>
205         If vbwDumpFileNum = 0 Then
206              vbwOpenDumpFile
207         End If
208         On Error Resume Next
209         Print #vbwDumpFileNum, lString
210         If Err = 52 Then
211            vbwCloseDumpFile
212            vbwOpenDumpFile
213            Print #vbwDumpFileNum, lString
214         End If
' <VB WATCH>
215        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Sub vbwOpenDumpFile()
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
216        On Error GoTo vbwErrHandler
' </VB WATCH>
217       vbwDumpFileNum = FreeFile
218       Open vbwDumpFile For Append As #vbwDumpFileNum
' <VB WATCH>
219        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Sub vbwCloseDumpFile()
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
220        On Error GoTo vbwErrHandler
' </VB WATCH>
221       Close #vbwDumpFileNum
' <VB WATCH>
222        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Function GetArrayDimension(ByRef arg As Variant) As Long
       ' vbwNoErrorHandler ' don't remove this !
223        Dim i As Long, j As Long
224        On Error Resume Next
225        i = 0
226        Do
227            i = i + 1
228            j = LBound(arg, i)
229        Loop Until Err.Number
230        GetArrayDimension = i - 1
End Function

Function vbwGetIndex(tObject As Variant) As String
       ' vbwNoErrorHandler ' don't remove this !
231        On Error Resume Next
232        vbwGetIndex = "(" & tObject.Index & ")"
End Function

Private Function FormatString(ByVal arg As String) As String
       ' vbwNoVariableDump ' don't remove this !
' <VB WATCH>
233        On Error GoTo vbwErrHandler
' </VB WATCH>

234        If Right$(arg, 1) = "}" Then ' probably a VB Watch built-in message
235             FormatString = arg
236             Exit Function
237        End If

           ' 1. truncate according to the vbwDumpStringMaxLength value
238        If vbwDumpStringMaxLength Then
239            If Len(arg) > vbwDumpStringMaxLength Then
240                arg = Left$(arg, vbwDumpStringMaxLength + 1)   ' +1: avoids to cut inside a vbCrLf '
241                If Right$(arg, 2) = vbCrLf Then
                       ' don't cut inside a vbCrLf
242                Else
243                    arg = Left$(arg, vbwDumpStringMaxLength)
244                End If
245                arg = arg & "{...}" ' truncated
246            End If
247        End If

           ' 2. make sure string isn't multiline
248        arg = Replace(arg, vbCrLf, "<CrLf>", , , vbBinaryCompare)
249        arg = Replace(arg, Chr(13), "<Cr>", , , vbBinaryCompare)
250        arg = Replace(arg, Chr(10), "<Lf>", , , vbBinaryCompare)

           ' 3. add quotes
251        FormatString = Chr(34) & arg & Chr(34)
' <VB WATCH>
252        Exit Function
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
    End Select
' </VB WATCH>
End Function

Sub vbwProcIn(ByRef lProc As String, Optional ByRef lParameters As String)
' <VB WATCH>
253        On Error GoTo vbwErrHandler
' </VB WATCH>

254        vbwTraceCallsNumber = vbwTraceCallsNumber + 1

255        vbwStackCallsNumber = vbwStackCallsNumber + 1
256        ReDim Preserve vbwStackCalls(1 To vbwStackCallsNumber)
257        vbwStackCalls(vbwStackCallsNumber) = lProc

258        Dim lString As String
259        lString = String$(vbwTraceCallsNumber - 1, vbTab) & lProc

260        If vbwLogTraceToFile Then
261             vbwSendLog lString & lParameters
262        End If

' <VB WATCH>
263        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Sub vbwProcOut(ByRef lProc As String)
' <VB WATCH>
264        On Error GoTo vbwErrHandler
' </VB WATCH>

265        If vbwTraceCallsNumber > 0 Then ' should always be true
266           vbwTraceCallsNumber = vbwTraceCallsNumber - 1
267        End If

268        If vbwStackCallsNumber > 0 Then ' should always be true
269           vbwStackCallsNumber = vbwStackCallsNumber - 1
270        End If

' <VB WATCH>
271        Exit Sub
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
    End Select
' </VB WATCH>
End Sub


Function vbwExecuteLine(ByRef fEncrypted As String, ByRef lLine As String) As Boolean
' <VB WATCH>
272        On Error GoTo vbwErrHandler
' </VB WATCH>

273        If vbwTraceLine Then

274            If fEncrypted Then
275                lLine = "<CRY>" & lLine & "</CRY>"
276            End If

277            If vbwLogTraceToFile Then
278                If vbwTraceCallsNumber > 0 Then
279                    vbwSendLog String$(vbwTraceCallsNumber - 1, vbTab) & " -> " & lLine
280                Else
281                    vbwSendLog " -> " & lLine
282                End If
283            End If

284        End If

           ' This function always returns false
' <VB WATCH>
285        Exit Function
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
    End Select
' </VB WATCH>
End Function

Function vbwGetStack() As String
' <VB WATCH>
286        On Error GoTo vbwErrHandler
' </VB WATCH>

287        If vbwTraceProc = False Then
288            vbwGetStack = "{Unavailable}"
289            Exit Function
290        End If

291        Dim vbwStackString As String
292        Dim i As Long

293        For i = vbwStackCallsNumber To 1 Step -1
294            vbwStackString = vbwStackString & String$(i - 1, vbTab) & vbwStackCalls(i) & vbCrLf
295        Next i
296        vbwGetStack = IIf(vbwStackString <> "", vbwStackString, "{Empty}")
' <VB WATCH>
297        Exit Function
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
    End Select
' </VB WATCH>
End Function

Sub vbwSendLog(ByRef tMsg As String)
' <VB WATCH>
298        On Error GoTo vbwErrHandler
' </VB WATCH>
299        If Err.Number Then
               ' Save Err object before being cleared by "On Error Resume Next"
300            Dim ErrDescription As String, ErrHelpFile As String, ErrSource As String
301            Dim ErrHelpContext As Long, ErrNumber As Long
302            ErrDescription = Err.Description
303            ErrHelpContext = Err.HelpContext
304            ErrHelpFile = Err.HelpFile
305            ErrNumber = Err.Number
306            ErrSource = Err.Source
307        End If

308        On Error Resume Next

309        If Not fLogFileOpen Then
310            fLogFileOpen = True
311            Dim suffix As Long
312            Do
313                Kill vbwLogFile
314                lLogFileNumber = CreateFile(vbwLogFile, GENERIC_WRITE, FILE_SHARE_READ, ByVal 0&, OPEN_ALWAYS, FILE_FLAG_OVERLAPPED, 0)
315                If lLogFileNumber < 0 Then
                       ' under some circumstances (retained in memory applications or while in the IDE)
                       ' the previous log file might not have been freed yet, so we must use another one
316                    suffix = suffix + 1
317                    vbwLogFile = Replace(App.Path & "\", ":\\", ":\") & "vbw" & App.EXEName & VBW_EXE_EXTENSION & suffix & ".log"
318                End If
319            Loop Until lLogFileNumber >= 0 Or suffix > 1000
320        End If

321        If Not fIsLogInitialize Then
               ' init file '
322            fIsLogInitialize = True
323            WriteToLogFile "Tracing " & App.Title
324            WriteToLogFile "Session started " & Now
325            WriteToLogFile ""
326        End If

           ' log to file
327        WriteToLogFile tMsg

328       If ErrNumber Then
               ' Restore Err object if cleared by "On Error Resume Next"
329            Err.Description = ErrDescription
330            Err.HelpContext = ErrHelpContext
331            Err.HelpFile = ErrHelpFile
332            Err.Number = ErrNumber
333            Err.Source = ErrSource
334        End If

' <VB WATCH>
335        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

' Writes Str as a new line in the log file (adding a vbCrLf to the end)
Private Function WriteToLogFile(Str As String) As Long
' <VB WATCH>
336        On Error GoTo vbwErrHandler
' </VB WATCH>
337        Dim ol As OVERLAPPED
338        Dim bBytes() As Byte, StrLength As Long
339        StrLength = Len(Str) + 2
340        ReDim bBytes(0 To StrLength - 1)
341        CopyMemory bBytes(0), ByVal Str & vbCrLf, StrLength
342        ol.OffSet = lLogFileOffset
343        WriteToLogFile = WriteFileEx(lLogFileNumber, bBytes(0), StrLength, ol, ByVal 0&)
344        lLogFileOffset = lLogFileOffset + StrLength
' <VB WATCH>
345        Exit Function
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
    End Select
' </VB WATCH>
End Function


' Ends a component's thread. If this was the last active thread, ends the component's process.
Public Sub vbwExitThread()
' <VB WATCH>
346        On Error GoTo vbwErrHandler
' </VB WATCH>
347        If vbwIsInIDE Then
               ' Executing ExitThread within the IDE will terminate VB without ceremony !
348            Stop ' Press the End button now
349        Else
350            Dim lpExitCode As Long
351            If GetExitCodeThread(GetCurrentThread(), lpExitCode) Then
352                ExitThread lpExitCode
353            End If
354        End If
' <VB WATCH>
355        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

' Ends a component's process. Equivalent to the End statement.
Public Sub vbwExitProcess()
' <VB WATCH>
356        On Error GoTo vbwErrHandler
' </VB WATCH>
357        If vbwIsInIDE Then
               ' Executing ExitProcess within the IDE will terminate VB without ceremony !
358            Stop ' Press the End button now
359        Else
360            Dim lpExitCode As Long
361            If GetExitCodeProcess(GetCurrentProcess(), lpExitCode) Then
362                ExitProcess lpExitCode
363            End If
364        End If
' <VB WATCH>
365        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

' determines if the program is running in the IDE or an EXE File
Private Function vbwIsInIDE() As Boolean
' <VB WATCH>
366        On Error GoTo vbwErrHandler
' </VB WATCH>

367        Dim strFileName As String
368        Dim lngCount As Long

369        strFileName = String(255, 0)
370        lngCount = GetModuleFileName(App.hInstance, strFileName, 255)
371        strFileName = Left(strFileName, lngCount)

372        vbwIsInIDE = UCase$(Right$(strFileName, 8)) Like "\VB#.EXE"

' <VB WATCH>
373        Exit Function
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
    End Select
' </VB WATCH>
End Function

' Exception handling stuff
Public Sub vbwHandleException()
           ' Exceptions will be caught and redirected to the failing procedure
' <VB WATCH>
374        On Error GoTo vbwErrHandler
' </VB WATCH>
375        SetUnhandledExceptionFilter AddressOf vbwExceptionFilter
' <VB WATCH>
376        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

' Exception handling stuff
Public Sub vbwUnHandleException()
           ' Exceptions are no longer caught and will cause Exceptions
           ' Whenever possible, call this procedure before returning to the VB's IDE
' <VB WATCH>
377        On Error GoTo vbwErrHandler
' </VB WATCH>
378        SetUnhandledExceptionFilter 0
' <VB WATCH>
379        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

' Exception handling stuff
Public Function vbwExceptionFilter(ByRef pExceptionInfo As EXCEPTION_POINTERS) As Long
       'vbwNoErrorHandler ' DO NOT remove this !!!

380        Dim ExceptionRecord As EXCEPTION_RECORD
381        ExceptionRecord = pExceptionInfo.pExceptionRecord

382        Do While ExceptionRecord.pExceptionRecord ' Empties the exceptions stack
383            CopyMemory ExceptionRecord, ByVal ExceptionRecord.pExceptionRecord, Len(ExceptionRecord)
384        Loop

385        vbwExceptionFilter = EXCEPTION_CONTINUE_EXECUTION

       'vbwExitProc ' because the next instruction causes to exit the function ' ' DO NOT remove this !!!

           ' Convert the exception to a normal VB error and go back to the failing procedure '
386        Err.Raise 65535, , ExceptionDescription(ExceptionRecord.ExceptionCode)

End Function

' Exception handling stuff
Private Function ExceptionDescription(ByVal ExceptionCode As Long) As String
       ' vbwNoErrorHandler ' don't remove this !
387        Select Case ExceptionCode
               Case EXCEPTION_ACCESS_VIOLATION
388                ExceptionDescription = "Exception: Access Violation"
389            Case EXCEPTION_DATATYPE_MISALIGNMENT
390                ExceptionDescription = "Exception: Datatype Misalignment"
391            Case EXCEPTION_BREAKPOINT
392                ExceptionDescription = "Exception: Breakpoint"
393            Case EXCEPTION_SINGLE_STEP
394                ExceptionDescription = "Exception: Single Step"
395            Case EXCEPTION_ARRAY_BOUNDS_EXCEEDED
396                ExceptionDescription = "Exception: Array Bounds Exceeded"
397            Case EXCEPTION_FLT_DENORMAL_OPERAND
398                ExceptionDescription = "Exception: Float Denormal Operand"
399            Case EXCEPTION_FLT_DIVIDE_BY_ZERO
400                ExceptionDescription = "Exception: Float Divide By Zero"
401            Case EXCEPTION_FLT_INEXACT_RESULT
402                ExceptionDescription = "Exception: Float Inexact Result"
403            Case EXCEPTION_FLT_INVALID_OPERATION
404                ExceptionDescription = "Exception: Float Invalid Operation"
405            Case EXCEPTION_FLT_OVERFLOW
406                ExceptionDescription = "Exception: Float Overflow"
407            Case EXCEPTION_FLT_STACK_CHECK
408                ExceptionDescription = "Exception: Float Stack Check"
409            Case EXCEPTION_FLT_UNDERFLOW
410                ExceptionDescription = "Exception: Float Underflow"
411            Case EXCEPTION_INT_DIVIDE_BY_ZERO
412                ExceptionDescription = "Exception: Integer Divide By Zero"
413            Case EXCEPTION_INT_OVERFLOW
414                ExceptionDescription = "Exception: Integer Overflow"
415            Case EXCEPTION_PRIV_INSTRUCTION
416                ExceptionDescription = "Exception: Priv Instruction"
417            Case EXCEPTION_IN_PAGE_ERROR
418                ExceptionDescription = "Exception: In Page Error"
419            Case EXCEPTION_ILLEGAL_INSTRUCTION
420                ExceptionDescription = "Exception: Illegal Instruction"
421            Case EXCEPTION_NONCONTINUABLE_EXCEPTION
422                ExceptionDescription = "Exception: Non Continuable Exception"
423            Case EXCEPTION_STACK_OVERFLOW
424                ExceptionDescription = "Exception: Stack Overflow"
425            Case EXCEPTION_INVALID_DISPOSITION
426                ExceptionDescription = "Exception: Invalid Disposition"
427            Case EXCEPTION_GUARD_PAGE
428                ExceptionDescription = "Exception: Guard Page"
429            Case EXCEPTION_INVALID_HANDLE
430                ExceptionDescription = "Exception: Invalid Handle"
431            Case CONTROL_C_EXIT
432                ExceptionDescription = "Exception: Control C Exit"
433            Case Else
434                ExceptionDescription = "Unknown Exception"
435        End Select

End Function

Public Sub vbwSaveErrObject()
       ' vbwNoErrorHandler ' don't remove this !
436        ErrObjectDescription = Err.Description
437        ErrObjectHelpContext = Err.HelpContext
438        ErrObjectHelpFile = Err.HelpFile
439        ErrObjectLastDllError = Err.LastDllError
440        ErrObjectNumber = Err.Number
441        ErrObjectSource = Err.Source
End Sub

Public Sub vbwRestoreErrObject()
       ' vbwNoErrorHandler ' don't remove this !
442       Err.Description = ErrObjectDescription
443       Err.HelpContext = ErrObjectHelpContext
444       Err.HelpFile = ErrObjectHelpFile
445       Err.Number = ErrObjectNumber
446       Err.Source = ErrObjectSource
End Sub


