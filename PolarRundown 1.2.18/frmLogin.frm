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
' </VB WATCH>

2          boCanApprove = False
3          If IsNull(txtInitials.Text) Or LenB(txtInitials.Text) = 0 Then
4              MsgBox "Please Enter Your Initials", vbOKOnly, "Please Enter Your Initials"
5          Else
6              LogInInitials = txtInitials.Text
7              If LogInInitials = strApproveInitials Then
8                  boCanApprove = True
9                  frmPLCData.cmdDeletePump.Visible = True
10                 frmPLCData.cmdApprovePump.Visible = True
11                 frmPLCData.cmdDeleteTestDate.Visible = True
12                 frmPLCData.cmdApproveTestDate.Visible = True
                   'frmPLCData.optReport(7).Visible = True
13                 frmPLCData.cmdAddNewBalanceHoles.Visible = True
14                 frmPLCData.cmdCalibrate.Visible = True
15             End If
16             frmPLCData.txtWho = LogInInitials
17             Me.Hide
18         End If
' <VB WATCH>
19         Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub Form_Activate()
' <VB WATCH>
20         On Error GoTo vbwErrHandler
' </VB WATCH>
21         Const HWND_TOPMOST As Integer = -1
           'Const HWND_NOTOPMOST As Integer = -2
22         Const SWP_NOSIZE As Integer = &H1
23         Const SWP_NOMOVE As Integer = &H2
24         Const SWP_NOACTIVATE As Integer = &H10
25         Const SWP_SHOWWINDOW As Integer = &H40

26         SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

       '    SetWindowPos Me.hWnd, -1, 0, 0, 520, 400, &H40
' <VB WATCH>
27         Exit Sub
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
    End Select
' </VB WATCH>
End Sub





