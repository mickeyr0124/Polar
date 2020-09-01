VERSION 5.00
Begin VB.Form FrmSODetails 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Sales Order Details"
   ClientHeight    =   5268
   ClientLeft      =   6888
   ClientTop       =   5532
   ClientWidth     =   7584
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5268
   ScaleWidth      =   7584
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   375
      Left            =   3120
      TabIndex        =   1
      Top             =   4800
      Width           =   1335
   End
   Begin VB.TextBox txtSOData 
      Height          =   3615
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   1080
      Width           =   7335
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   $"FrmSODetails.frx":0000
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
      Height          =   855
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   7455
   End
End
Attribute VB_Name = "FrmSODetails"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Me.Hide
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



