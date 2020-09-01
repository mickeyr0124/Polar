VERSION 5.00
Begin VB.Form frmDiagram 
   BackColor       =   &H00FFFFC0&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   7365
   ClientLeft      =   6645
   ClientTop       =   1755
   ClientWidth     =   13815
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7365
   ScaleWidth      =   13815
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   6300
      TabIndex        =   1
      Top             =   6720
      Width           =   1215
   End
   Begin VB.Image Image1 
      Height          =   6360
      Left            =   240
      Picture         =   "frmDiagram.frx":0000
      Stretch         =   -1  'True
      Top             =   120
      Width           =   13035
   End
   Begin VB.Label Label1 
      BackColor       =   &H00FFFFC0&
      Caption         =   "Cathy is going to draw a figure that goes here.  It will show where and how the transducers should be connected."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2655
      Left            =   1800
      TabIndex        =   0
      Top             =   1200
      Width           =   4215
   End
End
Attribute VB_Name = "frmDiagram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdClose_Click()
    Me.Hide
End Sub



