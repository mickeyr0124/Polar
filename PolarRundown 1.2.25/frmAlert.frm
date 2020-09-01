VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmAlert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Special Electrical Setup Required"
   ClientHeight    =   3165
   ClientLeft      =   6030
   ClientTop       =   2370
   ClientWidth     =   3315
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   3315
   Begin VB.TextBox txtAlert 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1695
      Left            =   192
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmAlert.frx":0000
      Top             =   240
      Width           =   2775
   End
   Begin MSForms.CommandButton cmdAlertOK 
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   2400
      Width           =   1455
      ForeColor       =   255
      Caption         =   "OK"
      Size            =   "2561;868"
      FontEffects     =   1073741825
      FontHeight      =   216
      FontCharSet     =   0
      FontPitchAndFamily=   2
      ParagraphAlign  =   3
      FontWeight      =   700
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000FF&
      BorderWidth     =   4
      Height          =   2055
      Left            =   75
      Top             =   120
      Width           =   3015
   End
End
Attribute VB_Name = "frmAlert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' <VB WATCH>
Const VBWMODULE = "frmAlert"
' </VB WATCH>

Private Sub cmdAlertOK_Click()
' <VB WATCH>
1          On Error GoTo vbwErrHandler
2          Const VBWPROCNAME = "frmAlert.cmdAlertOK_Click"
3          If vbwProtector.vbwTraceProc Then
4              Dim vbwProtectorParameterString As String
5              If vbwProtector.vbwTraceParameters Then
6                  vbwProtectorParameterString = "()"
7              End If
8              vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
9          End If
' </VB WATCH>
10         Me.Hide
' <VB WATCH>
11         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
12         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdAlertOK_Click"

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
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub Form_Activate()
' <VB WATCH>
13         On Error GoTo vbwErrHandler
14         Const VBWPROCNAME = "frmAlert.Form_Activate"
15         If vbwProtector.vbwTraceProc Then
16             Dim vbwProtectorParameterString As String
17             If vbwProtector.vbwTraceParameters Then
18                 vbwProtectorParameterString = "()"
19             End If
20             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
21         End If
' </VB WATCH>
22         Const HWND_TOPMOST As Integer = -1
23         Const SWP_NOSIZE As Integer = &H1
24         Const SWP_NOMOVE As Integer = &H2
25         Const SWP_NOACTIVATE As Integer = &H10
26         Const SWP_SHOWWINDOW As Integer = &H40

27         SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

       '    SetWindowPos Me.hWnd, -1, 0, 0, 520, 400, &H40
' <VB WATCH>
28         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
29         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Form_Activate"

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
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub


' <VB WATCH> <VBWATCHFINALPROC>
' Procedures added by VB Watch for variable dump


Private Sub vbwReportModuleVariables()
    vbwReportToFile VBW_MODULE_STRING
End Sub
' </VB WATCH>
