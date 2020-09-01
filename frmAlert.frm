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
Private Sub cmdAlertOK_Click()
    Me.Hide
End Sub
Private Sub Form_Activate()
    Const HWND_TOPMOST As Integer = -1
    Const SWP_NOSIZE As Integer = &H1
    Const SWP_NOMOVE As Integer = &H2
    Const SWP_NOACTIVATE As Integer = &H10
    Const SWP_SHOWWINDOW As Integer = &H40

    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

'    SetWindowPos Me.hWnd, -1, 0, 0, 520, 400, &H40
End Sub

