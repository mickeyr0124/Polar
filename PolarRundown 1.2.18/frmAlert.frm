VERSION 5.00
Object = "{0D452EE1-E08F-101A-852E-02608C4D0BB4}#2.0#0"; "FM20.DLL"
Begin VB.Form frmAlert 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Special Electrical Setup Required"
   ClientHeight    =   2340
   ClientLeft      =   6036
   ClientTop       =   2364
   ClientWidth     =   2796
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2340
   ScaleWidth      =   2796
   Begin VB.TextBox txtAlert 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.6
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   612
      Left            =   192
      MultiLine       =   -1  'True
      TabIndex        =   1
      Text            =   "frmAlert.frx":0000
      Top             =   240
      Width           =   2412
   End
   Begin MSForms.CommandButton cmdAlertOK 
      Height          =   492
      Left            =   672
      TabIndex        =   0
      Top             =   1560
      Width           =   1452
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
      Height          =   852
      Left            =   72
      Top             =   120
      Width           =   2652
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
' </VB WATCH>
2          Me.Hide
' <VB WATCH>
3          Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub Form_Activate()
' <VB WATCH>
4          On Error GoTo vbwErrHandler
' </VB WATCH>
5          Const HWND_TOPMOST As Integer = -1
6          Const SWP_NOSIZE As Integer = &H1
7          Const SWP_NOMOVE As Integer = &H2
8          Const SWP_NOACTIVATE As Integer = &H10
9          Const SWP_SHOWWINDOW As Integer = &H40

10         SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

       '    SetWindowPos Me.hWnd, -1, 0, 0, 520, 400, &H40
' <VB WATCH>
11         Exit Sub
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


