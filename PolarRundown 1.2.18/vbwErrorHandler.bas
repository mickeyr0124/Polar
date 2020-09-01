Attribute VB_Name = "vbwErrHandler"
Option Explicit

Public Enum vbwEnumRetCode
    vbwEnd = vbAbort           ' = 3 '
    vbwRetry = vbRetry         ' = 4 '
    vbwIgnoreLine = vbIgnore   ' = 5 '
    vbwAlwaysIgnore
    vbwDebug ' not used here
    vbwDoDumpVariable
    vbwCollapse ' not used here
End Enum
Public vbwRetCode As vbwEnumRetCode

Global vbwMessageString As String
Global vbwCircumstancesString As String
Global vbwErrorPath As String ' the directory where we will store files to zip and email
Global vbwfHasReported As Boolean

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Const SW_SHOWNORMAL = 1

Dim vbwTurnOffTimersString As String
Private Declare Function MessageBox Lib "user32" Alias "MessageBoxA" (ByVal hwnd As Long, ByVal lpText As String, ByVal lpCaption As String, ByVal wType As Long) As Long
Private Const vbMsgBoxSetTopMost = &H40000

Dim vbwTypeInfo As String
Dim vbwExcludedProperties As String

Private Declare Function GetForegroundWindow Lib "user32" () As Long
Private Declare Function SetForegroundWindow Lib "user32" (ByVal hwnd As Long) As Long
Private Declare Function GetDesktopWindow Lib "user32" () As Long
Dim hWndForeground As Long

Dim StrRememberErrors As String
Dim StrLastError As String

' vbwNoTraceProc vbwNoTraceLine ' don't remove this !
' vbwNoErrorHandler ' don't remove this !


' <VB WATCH>
Const VBWMODULE = "vbwErrHandler"
' </VB WATCH>

Public Function vbwErrorHandler( _
        ByVal ErrNumber As Long, _
        ByVal ErrDescription As String, _
        ByVal ErrProject As String, _
        ByVal ErrSource As String, _
        ByVal ErrProcedure As String, _
        ByVal ErrLine As Long) As vbwEnumRetCode

1          vbwRetCode = 0

2          If ErrNumber <> -1 Then ' ErrNumber = -1 would indicate that we're just back from dumping

               ' check if the user doesn't want to see this error anymore '
3              StrLastError = Chr(0) & ErrSource & Chr(1) & ErrNumber & Chr(1) & ErrLine & Chr(0)
4              If InStr(StrRememberErrors, StrLastError) Then
5                 vbwRetCode = vbwIgnoreLine
6                 Exit Function
7              End If

               ' save current Err object for later reporting
8              vbwSaveErrObject

               ' don't let timer events interfere with the error handling process
9              TurnOffTimers

               ' store which window was foreground in case we want to dump a screenshot of it later '
10             hWndForeground = GetForegroundWindow()


11             vbwMessageString = "An unexpected error has occurred in this application: " & vbCrLf & vbCrLf & _
                       "Error " & ErrNumber & vbCrLf & _
                       "Description: " & ErrDescription & vbCrLf & _
                       "Project: " & ErrProject & vbCrLf & _
                       "Module: " & ErrSource & vbCrLf & _
                       "Procedure: " & ErrProcedure & vbCrLf & _
                       "Line: " & ErrLine & vbCrLf & _
                       "Version: " & App.Major & "." & App.Minor & "." & App.Revision & vbCrLf

12             vbwfHasReported = False
13             vbwFrmErrHandler.Show 1

14             If vbwRetCode = vbwDoDumpVariable Then

                   ' init variable dump
15                 On Error Resume Next
                   ' prepare directory where we will store files to zip
16                 vbwErrorPath = Replace(App.Path & "\", ":\\", ":\") & "Error Report Files\"
17                 Kill vbwErrorPath & "*.*"
18                 RmDir vbwErrorPath
19                 MkDir vbwErrorPath
20                 On Error GoTo 0

21                 vbwDumpFile = vbwErrorPath & "Variables.txt"

                   ' instanciate a vbwFunctionsVB6.AdvancedFunctions object for future use
22                 On Error Resume Next
23                 Set vbwAdvancedFunctions = CreateObject("vbwFunctionsVB6.AdvancedFunctions")
24                 fIsVbwFunctionsInitialized = (Err.Number = 0)

25                 If fIsVbwFunctionsInitialized Then
                       ' set this to False if you don't wish to display the status to the user
26                     vbwAdvancedFunctions.ShowStatus = True
                       ' Init the VB Watch Functions Library object                                 '
                       ' Please refer to the help file under the VB Watch Protector Reference topic '
27                     vbwAdvancedFunctions.DumpFileName = vbwDumpFile
28                     vbwAdvancedFunctions.DumpMaxStringLength = vbwDumpStringMaxLength
29                     vbwAdvancedFunctions.DumpMaxObjectLength = 1000000
                       ' The following is internal data used to dump objects found in your project - Don't modify !'
30                 vbwExcludedProperties = ""

31                     vbwAdvancedFunctions.ExcludedProperties = vbwExcludedProperties ' VB Watch will replace <VBW EXCLUDED PROPERTIES> with the correct value of vbwExcludedProperties

32                     vbwAdvancedFunctions.ObjectInfo = vbwTypeInfo ' VB Watch will replace <VBW OBJECT FILE INFO> with the correct value of vbwTypeInfo
33                 End If
34             End If

35         Else
               ' Continuing from Sub vbwFrmErrHandler.cmdSendMail_click '

36             vbwfHasReported = True

               ' Now the variable dump file is in vbwErrorPath & "Variables.txt" '
               ' Let's include some other useful infos                           '

37             Dim ff As Long
38             On Error Resume Next
39             If vbwTraceProc Then
                   ' write stack and trace data
40                 If vbwCallStack Then
41                     ff = FreeFile
42                     Open vbwErrorPath & "Stack.txt" For Output As #ff
43                     Print #ff, vbwGetStack()
44                     Close #ff
45                 End If

46                 If vbwLogTraceToFile Then
47                     If FileExist(vbwLogFile) Then
48                       Open vbwErrorPath & "Trace.txt" For Output As #ff
49                       Print #ff, LoadFile(vbwLogFile)
50                       Close #ff
51                     End If
52                 End If
53             End If

               ' write error data
54             ff = FreeFile
55             Open vbwErrorPath & "Error.txt" For Output As #ff
56                 Print #ff, vbwMessageString & vbCrLf & vbCrLf & "Circumstances:" & vbwCircumstancesString
57             Close #ff

58             Dim success As Long
59             Dim msg As String, i As Long
60             If fIsVbwFunctionsInitialized Then
61                 If vbwSystemInfo Then
                      ' write file version data '
62                    vbwAdvancedFunctions.ExeInstance = App.hInstance
63                    vbwAdvancedFunctions.ExeCommandLine = Command$
64                    vbwAdvancedFunctions.VBRuntimeFile = "MSVBVM60.DLL"     ' VB Watch will replace with the correct value E.g. "MSVBVM60.DLL"     '
65                    vbwAdvancedFunctions.ComponentsUsed = "StdFormat;DAO;MSBind;MSDataReportLib;DERuntimeObjects;COMSVCSLib;Excel;TabDlg;MSDataGridLib;MSAdodcLib;MSDataListLib;MSChart20Lib;MSComDlg;MSFlexGridLib;MSHierarchicalFlexGridLib;RichTextLib;MSForms;MSComCtl2;FORMS;MSDATAREPORTRUNTIMELIB;MSSTDFMT;RICHTEXT;DERUNTIME"   ' VB Watch will replace with the correct value E.g. "MSGrid;MSComCtl2" '
66                    Open vbwErrorPath & "Version.txt" For Output As #ff
67                    Print #ff, vbwAdvancedFunctions.GetVersionInfo
68                    Close #ff

                      ' write user's system data '
69                    Open vbwErrorPath & "System.txt" For Output As #ff
70                    Print #ff, vbwAdvancedFunctions.GetSystemInfo
71                    Close #ff
72                 End If
73                 If vbwScreenshot Then
                      ' captures screenshot of the active window                                                                             '
                      ' (you may also use vbwAdvancedFunctions.GetWindowScreenShot(GetDesktopWindow) for a screenshot of a the full desktop) '
74                    If hWndForeground Then
75                         SetForegroundWindow hWndForeground
76                    End If
77                    SavePicture vbwAdvancedFunctions.GetWindowScreenShot(hWndForeground), vbwErrorPath & "Screenshot.bmp"
78                 End If

79                 Dim title As String
80                 title = "Error report for " & ErrProject & " " & App.Major & "." & App.Minor & "." & App.Revision

                   ' zip files to send '
81                 If vbwAdvancedFunctions.ZipDirectory(vbwErrorPath, title & ".zip") = 0 Then
                       ' successful                                                          '
                       ' this will create a button allowing the user to open the report file '
82                     vbwFrmErrHandler.Report = vbwErrorPath & title & ".zip"
83                 End If

                   ' email error report '
84                 success = vbwAdvancedFunctions.SendMessage(vbwEmailRecipientAdress, title, _
                       vbwMessageString & vbCrLf & vbCrLf & "Circumstances:" & vbwCircumstancesString _
                           & vbCrLf & vbCrLf & "(Note to sender: the attached file is " & vbwErrorPath & title & ".zip - Size: " _
                           & VBA.FileLen(vbwErrorPath & title & ".zip") \ 1024 & " Kb)", _
                       vbwErrorPath & title & ".zip")

85                 If success <> 0 Then
                       ' the SendMessage method failed (see error codes in the help file)            '
                       ' so we have to try an alternative method:                                    '
                       ' command line: mailto:<address>?subject=<Message Subject>&body=<Message Body>'
                       ' Warning ! command lines of more than 256 characters may be truncated !      '

86                     success = ShellExecute(0&, "open", _
                               "mailto:" & vbwEmailRecipientAdress & _
                               "?subject=" & title & _
                               "&body=" & "(Please attach: " & vbwErrorPath & title & ".zip - Size: " & VBA.FileLen(vbwErrorPath & title & ".zip") \ 1024 & " Kb to this message.)", _
                               "", "c:\", SW_SHOWNORMAL)
87                     If success > 32 Then
                           ' was successful to execute the previous command      '
                           ' but sometimes this isn't enough to send the message '
88                         If MessageBox(0&, "An email message should be displayed on your computer now, ready for sending." & vbCrLf & "Have you been able to attach the required file and send the message successfully ?", App.title, vbYesNo + vbQuestion + vbMsgBoxSetTopMost) = vbNo Then   '
89                             success = 32  ' simulates fail          '
90                         End If
91                     End If

92                     If success <= 32 Then
                           ' failed again !                         '
                           ' let the user send the message manually '
93                         msg = "Impossible to send the error report automatically !" & vbCrLf & vbCrLf
94                         msg = msg & "Please open your email messenger and send a message with the following parameters:" & vbCrLf
95                         msg = msg & "Address: " & vbwEmailRecipientAdress & vbCrLf
96                         msg = msg & "Subject: " & title & vbCrLf
97                         msg = msg & "Message: Please open the attached file." & vbCrLf
98                         msg = msg & "Attached file: " & vbwErrorPath & title & ".zip (Size: " & VBA.FileLen(vbwErrorPath & title & ".zip") \ 1024 & " Kb)" & vbCrLf & vbCrLf
99                         msg = msg & "Would you like to copy this message to the clipboard ?" & vbCrLf
100                        If MessageBox(0&, msg, App.title, vbYesNo + vbQuestion + vbMsgBoxSetTopMost) = vbYes Then
101                           For i = 1 To 10
102                              Clipboard.Clear
103                              Clipboard.SetText msg, vbCFText
104                           Next i
105                        End If
106                    End If
107                End If
108            Else
109                msg = "Error:" & vbCrLf & LoadFile(vbwErrorPath & "Error.txt") & vbCrLf & _
                         "Stack:" & vbCrLf & LoadFile(vbwErrorPath & "Stack.txt") & vbCrLf & _
                         "Variables:" & vbCrLf & LoadFile(vbwErrorPath & "Variables.txt")
110                For i = 1 To 10
111                    Clipboard.Clear
112                    Clipboard.SetText msg, vbCFText
113                Next i
114                success = ShellExecute(0&, "open", _
                           "mailto:" & vbwEmailRecipientAdress & _
                           "?subject=" & title & _
                           "&body=" & "Please paste the contents of the clipboard below (Press CTRL+V):", _
                           "", "c:\", SW_SHOWNORMAL)

115                If success > 32 Then
                       ' was successful to execute the previous command      '
                       ' but sometimes this isn't enough to send the message '
116                    If MessageBox(0&, "Have you been able to send the message successfully ?", App.title, vbYesNo + vbQuestion + vbMsgBoxSetTopMost) = vbNo Then   '
117                        success = 32  ' simulates fail          '
118                    End If
119                End If

120                If success <= 32 Then
                         ' failed again !                         '
                         ' let the user send the message manually '
121                      msg = "Please open your email messenger and send a message with the following parameters:" & vbCrLf & _
                               "Address: " & vbwEmailRecipientAdress & vbCrLf & _
                               "Subject: " & title & vbCrLf & _
                               "Message: " & vbCrLf & _
                               msg
122                      For i = 1 To 10
123                          Clipboard.Clear
124                          Clipboard.SetText msg, vbCFText
125                      Next i
126                      MessageBox 0&, "Impossible to send the error report automatically !" & vbCrLf & vbCrLf & _
                                        "Please send the message manually." & vbCrLf & _
                                        "Full instructions have been copied to the clipboard." & vbCrLf & _
                                        "Thank you.", App.title, 0&
127                End If
128            End If

               ' success - remove temp files '
129            On Error Resume Next
130            Kill vbwErrorPath & "Screenshot.bmp"
131            Kill vbwErrorPath & "System.txt"
132            Kill vbwErrorPath & "Version.txt"
133            Kill vbwErrorPath & "Error.txt"
134            Kill vbwErrorPath & "Trace.txt"
135            Kill vbwErrorPath & "Stack.txt"
136            Kill vbwErrorPath & "Variables.txt"
137            On Error GoTo 0

138            Set vbwAdvancedFunctions = Nothing

139            vbwFrmErrHandler.Show 1
140        End If

141        If vbwRetCode = vbwAlwaysIgnore Then
               ' remember this error to ignore it next time '
142            StrRememberErrors = StrRememberErrors & StrLastError
143            vbwRetCode = vbwIgnoreLine
144        End If

145        vbwErrorHandler = vbwRetCode

146        If vbwRetCode <> vbwDoDumpVariable And vbwRetCode <> vbwEnd Then
147            TurnOnTimers
148        End If

End Function

' gets the whole file in a string
Private Function LoadFile(ByVal sFileName As String) As String
149    On Error GoTo erreur
150        Dim nFile As Integer, sText As String
151        nFile = FreeFile
           'Open sFileName For Input As nFile ' Don't do this!!!
152        If Not FileExist(sFileName) Then
153             Exit Function
154        End If
           ' Let others read but not write
155        Open sFileName For Binary As nFile
156        sText = String$(LOF(nFile), 0)
157        Get nFile, 1, sText
158        Close nFile
159        LoadFile = sText
160    erreur:
End Function

' function FileExist not requiring a dir$ call
Private Function FileExist(ByVal tFile As String) As Boolean
161        On Error Resume Next
162        Call VBA.FileLen(tFile)
163        FileExist = (Err = 0)
End Function

' This disables all active timers because they might raise events during error handling time, '
' which would lead to unexpected results.                                                     '
' If you have other "outside" sources of events in your project, such as custom timers or     '
' subclassing code, this is the place to disable them temporarily.                            '
Private Sub TurnOffTimers()
164        Dim f As Form
165        Dim c As Control

166        On Error Resume Next
167        vbwTurnOffTimersString = Chr(0) ' Chr(0) = delimiter
168        For Each f In Forms
169            For Each c In f.Controls
170                If TypeName(c) = "Timer" Then
171                    If c.Enabled = True Then
                           ' remember its state by way of its object address
172                        vbwTurnOffTimersString = vbwTurnOffTimersString & ObjPtr(c) & Chr(0)
173                        c.Enabled = False
174                    End If
175                End If
176            Next c
177        Next f
End Sub

Private Sub TurnOnTimers()
178        Dim f As Form
179        Dim c As Control

180        On Error Resume Next
181        For Each f In Forms
182            For Each c In f.Controls
183                If TypeName(c) = "Timer" Then
184                    If InStr(vbwTurnOffTimersString, Chr(0) & ObjPtr(c) & Chr(0)) > 0 Then
185                        If ObjPtr(c) > 0 Then
186                            c.Enabled = True
187                        End If
188                    End If
189                End If
190            Next c
191        Next f
End Sub


