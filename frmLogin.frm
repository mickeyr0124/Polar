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

Private Sub Command1_Click()
    'Enter initials...compare to ApproveIntials (Admin).  if they are the same,
    '  allow Approval of data and deletion of test dates and/or pumps
    '  also, put initials in the "Operator" field

    boCanApprove = False
    If IsNull(txtInitials.Text) Or LenB(txtInitials.Text) = 0 Then
        MsgBox "Please Enter Your Initials", vbOKOnly, "Please Enter Your Initials"
    Else
        LogInInitials = txtInitials.Text
        If LogInInitials = strApproveInitials Then
            boCanApprove = True
            frmPLCData.cmdDeletePump.Visible = True
            frmPLCData.cmdApprovePump.Visible = True
            frmPLCData.cmdDeleteTestDate.Visible = True
            frmPLCData.cmdApproveTestDate.Visible = True
            'frmPLCData.optReport(7).Visible = True
            frmPLCData.cmdAddNewBalanceHoles.Visible = True
            frmPLCData.cmdCalibrate.Visible = True
        End If
        frmPLCData.txtWho = LogInInitials
        Me.Hide
    End If
End Sub

Private Sub Form_Activate()
    Const HWND_TOPMOST As Integer = -1
    'Const HWND_NOTOPMOST As Integer = -2
    Const SWP_NOSIZE As Integer = &H1
    Const SWP_NOMOVE As Integer = &H2
    Const SWP_NOACTIVATE As Integer = &H10
    Const SWP_SHOWWINDOW As Integer = &H40

    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

'    SetWindowPos Me.hWnd, -1, 0, 0, 520, 400, &H40
End Sub




