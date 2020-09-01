VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmEmail 
   Caption         =   "Form1"
   ClientHeight    =   5352
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   13212
   LinkTopic       =   "Form1"
   ScaleHeight     =   5352
   ScaleWidth      =   13212
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   6480
      TabIndex        =   20
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   4680
      TabIndex        =   19
      Top             =   4560
      Width           =   1455
   End
   Begin VB.CommandButton cmdUpdate 
      Caption         =   "Update"
      Height          =   495
      Left            =   2880
      TabIndex        =   18
      Top             =   4560
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Send E-mail When There Is A Change To:"
      Height          =   3495
      Left            =   2880
      TabIndex        =   3
      Top             =   840
      Width           =   9495
      Begin VB.Frame Frame3 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Performance Modifications"
         Height          =   2655
         Left            =   360
         TabIndex        =   11
         Top             =   480
         Width           =   4695
         Begin VB.CheckBox chkDischDia 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Discharge Orifice Value"
            Height          =   255
            Left            =   240
            TabIndex        =   16
            Top             =   1920
            Width           =   2775
         End
         Begin VB.CheckBox chkImpDia 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Impeller Diameter Value"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   1560
            Width           =   2655
         End
         Begin VB.CheckBox chkPumpDisch 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pump Discharge Orifice Check Box"
            Height          =   255
            Left            =   240
            TabIndex        =   14
            Top             =   1200
            Width           =   3015
         End
         Begin VB.CheckBox chkImpTrimmed 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Impeller Trimmed Check Box"
            Height          =   255
            Left            =   240
            TabIndex        =   13
            Top             =   840
            Width           =   2415
         End
         Begin VB.CheckBox chkImpFeathered 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Impeller Feathered Check Box"
            Height          =   255
            Left            =   240
            TabIndex        =   12
            Top             =   480
            Width           =   2895
         End
      End
      Begin VB.Frame Frame2 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Thrust Balance Modifications"
         Height          =   2655
         Left            =   5400
         TabIndex        =   4
         Top             =   480
         Width           =   3855
         Begin VB.CheckBox chkOtherMods 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Other Mods"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   2280
            Width           =   1815
         End
         Begin VB.CheckBox chkCircFlowDia 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Circulation Flow Orifice Value"
            Height          =   255
            Left            =   240
            TabIndex        =   9
            Top             =   1920
            Width           =   3135
         End
         Begin VB.CheckBox ChkBalHoleMod 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Balance Holes Modifications"
            Height          =   255
            Left            =   240
            TabIndex        =   8
            Top             =   1560
            Width           =   3015
         End
         Begin VB.CheckBox chkCircOrifice 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Circulation Flow Orifice Check Box"
            Height          =   255
            Left            =   240
            TabIndex        =   7
            Top             =   1200
            Width           =   2775
         End
         Begin VB.CheckBox ChkBalHoles 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Balance Holes Modified Check Box"
            Height          =   255
            Left            =   240
            TabIndex        =   6
            Top             =   840
            Width           =   2895
         End
         Begin VB.CheckBox chkEndPlay 
            BackColor       =   &H00FFFFC0&
            Caption         =   "End Play Check Box"
            Height          =   255
            Left            =   240
            TabIndex        =   5
            Top             =   480
            Width           =   1815
         End
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid MSHFlexGrid1 
      Height          =   2895
      Left            =   240
      TabIndex        =   1
      Top             =   960
      Width           =   2175
      _ExtentX        =   3831
      _ExtentY        =   5101
      _Version        =   393216
      Cols            =   1
      FixedCols       =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   1
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   2880
      TabIndex        =   0
      Top             =   360
      Width           =   2775
   End
   Begin VB.Label Label2 
      Caption         =   "List of E-mails"
      Height          =   255
      Left            =   240
      TabIndex        =   17
      Top             =   600
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "E-mail"
      Height          =   255
      Left            =   2880
      TabIndex        =   2
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "frmEmail"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim rsData As New ADODB.Recordset
Dim qyData As New ADODB.Command
Dim rsData1 As New ADODB.Recordset
Dim qyData1 As New ADODB.Command

Private Sub cmdAdd_Click()

    rsData1.AddNew "email", "New Addition"
    qyData.CommandText = "SELECT Email FROM ModificationEmail ORDER BY Email;"
    Set MSHFlexGrid1.DataSource = rsData
    MShFlexGrid1_Click
End Sub

Private Sub cmdDelete_Click()
    rsData1.Delete
    qyData.CommandText = "SELECT Email FROM ModificationEmail ORDER BY Email;"
    Set MSHFlexGrid1.DataSource = rsData
    MShFlexGrid1_Click
End Sub

Private Sub cmdUpdate_Click()
    rsData1.Update
    qyData.CommandText = "SELECT Email FROM ModificationEmail ORDER BY Email;"
    Set MSHFlexGrid1.DataSource = rsData
    MShFlexGrid1_Click
End Sub

Private Sub Form_Load()
    MSHFlexGrid1.ColWidth(0) = 3000
    MSHFlexGrid1.RowSel = 1

    qyData.ActiveConnection = cnPumpData
    qyData1.ActiveConnection = cnPumpData
    qyData.CommandText = "SELECT Email FROM ModificationEmail ORDER BY Email;"
    rsData.CursorType = adOpenStatic
    rsData.CursorLocation = adUseClient
    rsData1.CursorType = adOpenStatic
    rsData1.CursorLocation = adUseClient
    rsData1.LockType = adLockPessimistic
    rsData.Open qyData

    Set MSHFlexGrid1.DataSource = rsData
    MSHFlexGrid1.RowSel = 1
    MShFlexGrid1_Click
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If rsData1.State = adStateOpen Then
        rsData1.Close
        Set rsData1 = Nothing
        Set qyData1 = Nothing
    End If
    If rsData.State = adStateOpen Then
        rsData.Close
        Set rsData = Nothing
        Set qyData = Nothing
    End If
  
End Sub

Private Sub MShFlexGrid1_Click()


    MSHFlexGrid1.Row = MSHFlexGrid1.RowSel
    MSHFlexGrid1.Col = 0
    FixHighlight (MSHFlexGrid1.Row)

    If rsData1.State = 1 Then
        rsData1.Close
    End If

    qyData1.CommandText = "SELECT * FROM ModificationEmail WHERE ModificationEmail.Email = '" & MSHFlexGrid1.Text & "' ORDER BY Email ;"
    rsData1.Open qyData1

    Set Text1.DataSource = rsData1
    Text1.DataField = "Email"
    Set chkImpFeathered.DataSource = rsData1
    chkImpFeathered.DataField = "ImpellerFeathered"
    Set chkImpTrimmed.DataSource = rsData1
    chkImpTrimmed.DataField = "ImpellerTrimmed"
    Set chkPumpDisch.DataSource = rsData1
    chkPumpDisch.DataField = "DischargeOrifice"
    Set chkImpDia.DataSource = rsData1
    chkImpDia.DataField = "ImpellerDiameter"
    Set chkDischDia.DataSource = rsData1
    chkDischDia.DataField = "OrificeDiameter"

    Set chkEndPlay.DataSource = rsData1
    chkEndPlay.DataField = "Endplay"
    Set ChkBalHoles.DataSource = rsData1
    ChkBalHoles.DataField = "BalanceHolesModified"
    Set chkCircOrifice.DataSource = rsData1
    chkCircOrifice.DataField = "CirculationFlowOrifice"
    Set chkCircFlowDia.DataSource = rsData1
    chkCircFlowDia.DataField = "CirculationFlowDiameter"
    Set chkOtherMods.DataSource = rsData1
    chkOtherMods.DataField = "OtherMods"
    Set ChkBalHoleMod.DataSource = rsData1
    ChkBalHoleMod.DataField = "BalanceHoleModifications"
  
End Sub
Sub FixHighlight(Row As Integer)
    Dim I As Integer
    Const PresentColor = -2147483643

    MSHFlexGrid1.Col = 0

    For I = 1 To MSHFlexGrid1.Rows - 1
        MSHFlexGrid1.Row = I
        MSHFlexGrid1.CellBackColor = PresentColor
    Next I

    MSHFlexGrid1.Row = Row
    MSHFlexGrid1.CellBackColor = vbYellow
End Sub


