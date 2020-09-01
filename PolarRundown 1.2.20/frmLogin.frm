VERSION 5.00
Begin VB.Form frmLogin 
   Caption         =   "frmLogIn"
   ClientHeight    =   2976
   ClientLeft      =   5112
   ClientTop       =   5052
   ClientWidth     =   4860
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   2976
   ScaleWidth      =   4860
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   495
      Left            =   1740
      TabIndex        =   2
      Top             =   1800
      Width           =   1335
   End
   Begin VB.TextBox txtInitials 
      Height          =   375
      Left            =   1980
      TabIndex        =   1
      Top             =   840
      Width           =   855
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "Please Login by Entering Your Initials"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmLogin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' <VB WATCH>
Const VBWMODULE = "frmLogin"
' </VB WATCH>

Private Sub Command1_Click()
           'Enter initials...compare to ApproveIntials (Admin).  if they are the same,
           '  allow Approval of data and deletion of test dates and/or pumps
           '  also, put initials in the "Operator" field
' <VB WATCH>
1          On Error GoTo vbwErrHandler
2          Const VBWPROCNAME = "frmLogin.Command1_Click"
3          If vbwProtector.vbwTraceProc Then
4              Dim vbwProtectorParameterString As String
5              If vbwProtector.vbwTraceParameters Then
6                  vbwProtectorParameterString = "()"
7              End If
8              vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
9          End If
' </VB WATCH>

10         boCanApprove = False
11         If IsNull(txtInitials.Text) Or LenB(txtInitials.Text) = 0 Then
12             MsgBox "Please Enter Your Initials", vbOKOnly, "Please Enter Your Initials"
13         Else
14             LogInInitials = txtInitials.Text
15             If LogInInitials = strApproveInitials Then
16                 boCanApprove = True
17                 frmPLCData.cmdDeletePump.Visible = True
18                 frmPLCData.cmdApprovePump.Visible = True
19                 frmPLCData.cmdDeleteTestDate.Visible = True
20                 frmPLCData.cmdApproveTestDate.Visible = True
                   'frmPLCData.optReport(7).Visible = True
21                 frmPLCData.cmdAddNewBalanceHoles.Visible = True
22                 frmPLCData.cmdCalibrate.Visible = True
23             End If
24             frmPLCData.txtWho = LogInInitials
25             Me.Hide
26         End If
' <VB WATCH>
27         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
28         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Command1_Click"

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
29         On Error GoTo vbwErrHandler
30         Const VBWPROCNAME = "frmLogin.Form_Activate"
31         If vbwProtector.vbwTraceProc Then
32             Dim vbwProtectorParameterString As String
33             If vbwProtector.vbwTraceParameters Then
34                 vbwProtectorParameterString = "()"
35             End If
36             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
37         End If
' </VB WATCH>
38         Const HWND_TOPMOST As Integer = -1
           'Const HWND_NOTOPMOST As Integer = -2
39         Const SWP_NOSIZE As Integer = &H1
40         Const SWP_NOMOVE As Integer = &H2
41         Const SWP_NOACTIVATE As Integer = &H10
42         Const SWP_SHOWWINDOW As Integer = &H40

43         SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

       '    SetWindowPos Me.hWnd, -1, 0, 0, 520, 400, &H40
' <VB WATCH>
44         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
45         Exit Sub
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
