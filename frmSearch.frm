VERSION 5.00
Object = "{F0D2F211-CCB0-11D0-A316-00AA00688B10}#1.0#0"; "MSDATLST.OCX"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Begin VB.Form frmSearch 
   Caption         =   "Search for Pumps"
   ClientHeight    =   12420
   ClientLeft      =   192
   ClientTop       =   516
   ClientWidth     =   16548
   LinkTopic       =   "Form1"
   ScaleHeight     =   12420
   ScaleWidth      =   16548
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame frmWildCard 
      Caption         =   "Search By Portion Of Model Number"
      Height          =   1335
      Left            =   0
      TabIndex        =   28
      Top             =   9960
      Width           =   15015
      Begin VB.TextBox txtModelNumberString 
         Height          =   375
         Left            =   480
         TabIndex        =   30
         Text            =   "Enter Characters and Search with Return"
         Top             =   360
         Width           =   3135
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgWildCard 
         Height          =   975
         Left            =   3840
         TabIndex        =   29
         Top             =   240
         Width           =   10935
         _ExtentX        =   19283
         _ExtentY        =   1715
         _Version        =   393216
         FixedCols       =   0
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.CommandButton cmdResetSizes 
      Caption         =   "Reset Sizes"
      Height          =   375
      Left            =   3480
      TabIndex        =   27
      Top             =   11460
      Width           =   1575
   End
   Begin VB.Frame frmShipTo 
      Caption         =   "Search By Ship To Customer"
      Height          =   1335
      Left            =   0
      TabIndex        =   17
      Top             =   8520
      Width           =   15015
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgShipTo 
         Height          =   975
         Left            =   3840
         TabIndex        =   19
         Top             =   240
         Width           =   10935
         _ExtentX        =   19283
         _ExtentY        =   1715
         _Version        =   393216
         FixedCols       =   0
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSDataListLib.DataCombo cmbSearchShipTo 
         Height          =   315
         Left            =   600
         TabIndex        =   18
         Top             =   600
         Width           =   3015
         _ExtentX        =   5313
         _ExtentY        =   508
         _Version        =   393216
         Text            =   "Select Customer"
      End
   End
   Begin VB.Frame frmCustomer 
      Caption         =   "Search By Bill To Customer"
      Height          =   1335
      Left            =   0
      TabIndex        =   12
      Top             =   7200
      Width           =   15015
      Begin MSDataListLib.DataCombo cmbSearchCustomer 
         Height          =   315
         Left            =   600
         TabIndex        =   13
         Top             =   600
         Width           =   3015
         _ExtentX        =   5313
         _ExtentY        =   508
         _Version        =   393216
         Text            =   "Select Customer"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgCustomer 
         Height          =   975
         Left            =   3840
         TabIndex        =   20
         Top             =   240
         Width           =   10935
         _ExtentX        =   19283
         _ExtentY        =   1715
         _Version        =   393216
         FixedCols       =   0
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame frmSalesOrder 
      Caption         =   "Search By Sales Order"
      Height          =   1335
      Left            =   0
      TabIndex        =   10
      Top             =   1920
      Width           =   15015
      Begin MSDataListLib.DataCombo cmbSearchSalesOrder 
         Height          =   315
         Left            =   600
         TabIndex        =   11
         Top             =   600
         Width           =   3015
         _ExtentX        =   5313
         _ExtentY        =   508
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "Select Sales Order"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgSalesOrder 
         Height          =   975
         Left            =   3840
         TabIndex        =   24
         Top             =   240
         Width           =   10935
         _ExtentX        =   19283
         _ExtentY        =   1715
         _Version        =   393216
         Rows            =   5
         FixedCols       =   0
         ScrollTrack     =   -1  'True
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame frmTEMCFrameNo 
      Caption         =   "Search By Teikoku Frame Number"
      Height          =   1335
      Left            =   0
      TabIndex        =   8
      Top             =   5880
      Width           =   15015
      Begin MSDataListLib.DataCombo cmbSearchTEMCFrameNumber 
         Height          =   315
         Left            =   600
         TabIndex        =   9
         Top             =   600
         Width           =   3015
         _ExtentX        =   5313
         _ExtentY        =   508
         _Version        =   393216
         Text            =   "Select TEMC Frame No"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgTEMCFrameNo 
         Height          =   975
         Left            =   3840
         TabIndex        =   21
         Top             =   240
         Width           =   10935
         _ExtentX        =   19283
         _ExtentY        =   1715
         _Version        =   393216
         FixedCols       =   0
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.Frame frmModel 
      Caption         =   "Search By TEMC Hydraulics"
      Height          =   1335
      Left            =   0
      TabIndex        =   6
      Top             =   4560
      Width           =   15015
      Begin VB.ComboBox cmbSearchModel 
         Height          =   315
         ItemData        =   "frmSearch.frx":0000
         Left            =   600
         List            =   "frmSearch.frx":0002
         TabIndex        =   7
         Text            =   "Select TEMC Hydraulics"
         Top             =   600
         Width           =   3015
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgModel 
         Height          =   972
         Left            =   3960
         TabIndex        =   22
         Top             =   240
         Width           =   10932
         _ExtentX        =   19283
         _ExtentY        =   1715
         _Version        =   393216
         FixedCols       =   0
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
         _Band(0).GridLinesBand=   1
         _Band(0).TextStyleBand=   0
         _Band(0).TextStyleHeader=   0
      End
   End
   Begin VB.Frame frmDate 
      Caption         =   "Search By Date"
      Height          =   1935
      Left            =   0
      TabIndex        =   5
      Top             =   0
      Width           =   15015
      Begin VB.ComboBox cmbStartDate 
         Height          =   315
         Left            =   600
         TabIndex        =   26
         Text            =   "Select Start Date"
         Top             =   720
         Width           =   3015
      End
      Begin MSDataListLib.DataCombo cmbSearchEndDate 
         Height          =   315
         Left            =   600
         TabIndex        =   14
         Top             =   1440
         Width           =   3015
         _ExtentX        =   5313
         _ExtentY        =   508
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "Select End Date"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgDate 
         Height          =   1575
         Left            =   3840
         TabIndex        =   25
         Top             =   240
         Width           =   10935
         _ExtentX        =   19283
         _ExtentY        =   2773
         _Version        =   393216
         FixedCols       =   0
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label2 
         Caption         =   "End Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   600
         TabIndex        =   16
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Start Date"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.4
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   600
         TabIndex        =   15
         Top             =   480
         Width           =   1095
      End
   End
   Begin VB.Frame frmSN 
      Caption         =   "Search By Serial Number"
      Height          =   1335
      Left            =   0
      TabIndex        =   3
      Top             =   3240
      Width           =   15015
      Begin MSDataListLib.DataCombo cmbSearchSN 
         Height          =   315
         Left            =   600
         TabIndex        =   4
         Top             =   600
         Width           =   3015
         _ExtentX        =   5313
         _ExtentY        =   508
         _Version        =   393216
         MatchEntry      =   -1  'True
         Text            =   "Select Serial Number"
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid fgSN 
         Height          =   975
         Left            =   3840
         TabIndex        =   23
         Top             =   240
         Width           =   10935
         _ExtentX        =   19283
         _ExtentY        =   1715
         _Version        =   393216
         FixedCols       =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   7440
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   11460
      Visible         =   0   'False
      Width           =   2895
   End
   Begin VB.CommandButton cmdClose 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Close"
      Height          =   375
      Left            =   960
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   11460
      Width           =   1335
   End
   Begin VB.Label lblNoOfPumps 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.4
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2400
      TabIndex        =   2
      Top             =   11640
      Visible         =   0   'False
      Width           =   1935
   End
End
Attribute VB_Name = "frmSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim rsData As New ADODB.Recordset
Dim qyData As New ADODB.Command
Dim rsData1 As New ADODB.Recordset
Dim qyData1 As New ADODB.Command
Dim qyData2 As New ADODB.Command
Dim rsData2 As New ADODB.Recordset
Dim rsDataDate As New ADODB.Recordset
Dim qyDataDate As New ADODB.Command
Dim rsDataModel As New ADODB.Recordset
Dim qyDataModel As New ADODB.Command
Dim rsDataSN As New ADODB.Recordset
Dim qyDataSN As New ADODB.Command
Dim rsDataSalesOrder As New ADODB.Recordset
Dim qySalesOrderData As New ADODB.Command
Dim rsSalesOrderData As New ADODB.Recordset
Dim qyDataSalesOrder As New ADODB.Command
Dim rsDataTEMCModel As New ADODB.Recordset
Dim qyDataTEMCModel As New ADODB.Command
Dim rsDataTEMCFrameNumber As New ADODB.Recordset
Dim qyDataTEMCFrameNumber As New ADODB.Command
Dim rsDataCustomer As New ADODB.Recordset
Dim qyDataCustomer As New ADODB.Command
Dim rsCustomerData As New ADODB.Recordset
Dim qyCustomerData As New ADODB.Command
Dim rsDataShipTo As New ADODB.Recordset
Dim qyDataShipto As New ADODB.Command
Dim rsShipToData As New ADODB.Recordset
Dim qyShipToData As New ADODB.Command
Private Sub cmbSalesOrder_Click(Area As Integer)
  
End Sub

Private Sub cmbSearchCustomer_Click(Area As Integer)
    If rsCustomerData.State = adStateOpen Then
        rsCustomerData.Close
    End If

    If cmbSearchCustomer.SelectedItem = 1 Then
        Exit Sub
    End If

    qyCustomerData.CommandText = "SELECT DISTINCT " & _
                  " [TempPumpData]![SerialNumber], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], TempPumpData.ShiptoCustomer " & _
                  " FROM (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
                  " WHERE (((TempPumpData.BillToCustomer)= '" & cmbSearchCustomer.BoundText & "'));"

    rsCustomerData.Open qyCustomerData

    If rsCustomerData.RecordCount = 0 Then
        Exit Sub
    End If

    'bind the datalist to the recordset
    Set fgCustomer.DataSource = rsCustomerData

    fgCustomer.ColWidth(0) = 1400
    fgCustomer.ColWidth(1) = 2000
    fgCustomer.ColWidth(2) = 1200
    fgCustomer.ColWidth(3) = 2000
    fgCustomer.ColWidth(4) = 3200
    fgCustomer.TextMatrix(0, 0) = "S/N"
    fgCustomer.TextMatrix(0, 1) = "Date"
    fgCustomer.TextMatrix(0, 2) = "Sales Order"
    fgCustomer.TextMatrix(0, 3) = "Model No"
    fgCustomer.TextMatrix(0, 4) = "Ship To"

    frmModel.Visible = False
    frmTEMCFrameNo.Visible = False
    frmCustomer.Top = 7200 - (4000 - 1335)
    frmCustomer.Height = 4000
    fgCustomer.Height = 4000 - 360
    frmCustomer.FontBold = True
  
End Sub


Private Sub cmbSearchEndDate_Click(Area As Integer)
    cmbStartDate_Click
End Sub

Private Sub cmbSearchModel_Click()
    If rsDataModel.State = adStateOpen Then
        rsDataModel.Close
    End If

    qyDataModel.CommandText = "SELECT DISTINCT " & _
                  " [TempPumpData]![SerialNumber], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], IIF(TempTestSetupData.ImpTrimmed=0, val(TempPumpData!ImpellerDia), val(TempTestSetupData!ImpTrimmed)) as ImpDia ,  " & _
                  "  TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer" & _
                  " FROM Motor INNER JOIN ((Model INNER JOIN TempPumpData ON Model.Model = TempPumpData.Model) INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) ON Motor.Motor = TempPumpData.Motor" & _
                  " WHERE (((TempPumpData.ModelNumber) LIKE '%" & cmbSearchModel.List(cmbSearchModel.ListIndex) & "%'));"
'                  " WHERE (((TempPumpData.Model)= " & cmbSearchModel.ItemData(cmbSearchModel.ListIndex) & "));"

    rsDataModel.Open qyDataModel

    If rsDataModel.RecordCount = 0 Then
        Exit Sub
    End If

    Set fgModel.DataSource = rsDataModel

    Dim f As String

    f = "<S/N     |<Date      |<Sales Order |<Model No     |^Imp Dia  |<Bill To     |<Ship To  "
    fgModel.FormatString = f
    fgModel.ColAlignment(4) = flexAlignCenterTop
    'fgModel.ColAlignment(5) = flexAlignCenterTop
    fgModel.ColWidth(0) = 1200
    fgModel.ColWidth(1) = 2000
    fgModel.ColWidth(2) = 1200
    fgModel.ColWidth(3) = 2000
    fgModel.ColWidth(4) = 1200
    'fgModel.ColWidth(5) = 1200
    fgModel.ColWidth(5) = 3200
    fgModel.ColWidth(6) = 3200
    fgModel.TextMatrix(0, 0) = "S/N"
    fgModel.TextMatrix(0, 1) = "Date"
    fgModel.TextMatrix(0, 2) = "Sales Order"
    fgModel.TextMatrix(0, 3) = "Model No"
    fgModel.TextMatrix(0, 4) = "Imp Dia"
    'fgModel.TextMatrix(0, 5) = "Motor"
    fgModel.TextMatrix(0, 5) = "Bill To"
    fgModel.TextMatrix(0, 6) = "Ship To"

    Dim x As Long
    With fgModel
        For x = .FixedRows To .Rows - 1
        .TextMatrix(x, 4) = Format(.TextMatrix(x, 4), "#0.000")
        Next x
    End With


    frmTEMCFrameNo.Visible = False
    frmCustomer.Visible = False
    frmModel.Height = 4000
    fgModel.Height = 4000 - 360
    frmModel.FontBold = True
  
End Sub

Private Sub cmbSearchSalesOrder_Change()

    Text1.Text = cmbSearchSalesOrder.BoundText

    If rsSalesOrderData.State = adStateOpen Then
        rsSalesOrderData.Close
    End If

    'find all dates and models for the selected serial number
    qySalesOrderData.CommandText = "SELECT DISTINCT " & _
                  " [TempPumpData]![SerialNumber], [TempPumpData]![ChempumpPump], [TempTestSetupData]![Date], [TempPumpData]![ModelNumber], [TempPumpData]![TEMCFrameNumber], TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer " & _
                  " FROM (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
                  " WHERE (((TempPumpData.SalesOrderNumber)= '" & cmbSearchSalesOrder.BoundText & "'));"

    rsSalesOrderData.Open qySalesOrderData

    If rsSalesOrderData.RecordCount = 0 Then
        Exit Sub
    End If

'    'bind the datalist to the other recordset
    Set fgSalesOrder.DataSource = rsSalesOrderData

    fgSalesOrder.ColWidth(0) = 1400               'serial number
    fgSalesOrder.ColWidth(1) = 0               'chempumppump
    fgSalesOrder.ColWidth(2) = 2000
    If rsSalesOrderData.Fields(1) = True Then       'show model number
        fgSalesOrder.ColWidth(3) = 1800
        fgSalesOrder.ColWidth(4) = 0
    Else                                    'else, show TEMC Frame number
        fgSalesOrder.ColWidth(3) = 0
        fgSalesOrder.ColWidth(4) = 1800
    End If
    fgSalesOrder.ColWidth(5) = 3200
    fgSalesOrder.ColWidth(6) = 3200
    fgSalesOrder.TextMatrix(0, 0) = "S/N"
    fgSalesOrder.TextMatrix(0, 2) = "Date"
    fgSalesOrder.TextMatrix(0, 3) = "Model No"
    fgSalesOrder.TextMatrix(0, 4) = "TEMC Frame"
    fgSalesOrder.TextMatrix(0, 5) = "Bill To"
    fgSalesOrder.TextMatrix(0, 6) = "Ship To"

    frmSN.Visible = False
    frmSalesOrder.Height = 4000
    fgSalesOrder.Height = 4000 - 360
    frmSalesOrder.FontBold = True
  
End Sub

Private Sub cmbSearchShipTo_Click(Area As Integer)

    If rsShipToData.State = adStateOpen Then
        rsShipToData.Close
    End If

    If cmbSearchShipTo.SelectedItem = 1 Then
        Exit Sub
    End If

    qyShipToData.CommandText = "SELECT DISTINCT " & _
                  " [TempPumpData]![SerialNumber],[TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], TempPumpData.ModelNumber, TempPumpData.BillToCustomer " & _
                  " FROM  (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
                  " WHERE (((TempPumpData.ShipToCustomer)= '" & cmbSearchShipTo.BoundText & "'));"

    rsShipToData.Open qyShipToData

    If rsShipToData.RecordCount = 0 Then
        Exit Sub
    End If

    Set fgShipTo.DataSource = rsShipToData


    fgShipTo.ColWidth(0) = 1400
    fgShipTo.ColWidth(1) = 2000
    fgShipTo.ColWidth(2) = 1200
    fgShipTo.ColWidth(3) = 2000
    fgShipTo.ColWidth(4) = 3200
    fgShipTo.TextMatrix(0, 0) = "S/N"
    fgShipTo.TextMatrix(0, 1) = "Date"
    fgShipTo.TextMatrix(0, 2) = "Sales Order"
    fgShipTo.TextMatrix(0, 3) = "Model No"
    fgShipTo.TextMatrix(0, 4) = "Bill To"

    frmTEMCFrameNo.Visible = False
    frmCustomer.Visible = False
    frmShipTo.Top = 8520 - (4000 - 1335)
    frmShipTo.Height = 4000
    fgShipTo.Height = 4000 - 360
    frmShipTo.FontBold = True
  
End Sub

Private Sub cmbSearchSN_Change()

    Text1.Text = cmbSearchSN.BoundText

    If rsDataSN.State = adStateOpen Then
        rsDataSN.Close
    End If

    'find all dates and models for the selected serial number
    qyDataSN.CommandText = "SELECT DISTINCT TempPumpData.SerialNumber," & _
                  " [TempPumpData]![ChempumpPump],[TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], [TempPumpData]![TEMCFrameNumber], TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer " & _
                  " FROM (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
                  " WHERE (((TempPumpData.SerialNumber)= '" & cmbSearchSN.BoundText & "'));"

    rsDataSN.Open qyDataSN

    'if we didn't find any records, see if we have any serial numbers that are close
    If rsDataSN.RecordCount = 0 Then
        rsDataSN.Close
        qyDataSN.CommandText = "SELECT DISTINCT TempPumpData.SerialNumber," & _
                      " [TempPumpData]![ChempumpPump],[TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], [TempPumpData]![TEMCFrameNumber], TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer " & _
                      " FROM (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
                      " WHERE (((TempPumpData.SerialNumber)= '" & cmbSearchSN.BoundText & "%'));"
        rsDataSN.Open qyDataSN
    End If

    If rsDataSN.RecordCount = 0 Then
        Exit Sub
    End If

    lblNoOfPumps.Caption = rsDataSN.RecordCount & " Pumps Found"

    'bind the datalist to the other recordset
    Set fgSN.DataSource = rsDataSN

    fgSN.ColWidth(0) = 0               'serial number
    fgSN.ColWidth(1) = 0               'chempumppump
    fgSN.ColWidth(2) = 2000
    fgSN.ColWidth(3) = 1200
    If rsDataSN.Fields(1) = True Then       'show model number
        fgSN.ColWidth(4) = 1800
        fgSN.ColWidth(5) = 0
    Else                                    'else, show TEMC Frame number
        fgSN.ColWidth(4) = 0
        fgSN.ColWidth(5) = 1800
    End If
    fgSN.ColWidth(6) = 3000
    fgSN.ColWidth(7) = 3000
    fgSN.TextMatrix(0, 2) = "Date"
    fgSN.TextMatrix(0, 3) = "Sales Order"
    fgSN.TextMatrix(0, 4) = "Model No"
    fgSN.TextMatrix(0, 5) = "TEMC Frame"
    fgSN.TextMatrix(0, 6) = "Bill To"
    fgSN.TextMatrix(0, 7) = "Ship To"

    'put the serial number into the find pump textbox
    frmPLCData.txtSN.Text = rsDataSN.Fields("SerialNumber")

    frmModel.Visible = False
    frmSN.Height = 4000
    fgSN.Height = 4000 - 360
    frmSN.FontBold = True
  
End Sub

Private Sub ORIGINALcmbSearchTEMCFrameNumber_Click(Area As Integer)
    If rsDataTEMCFrameNumber.State = adStateOpen Then
        rsDataTEMCFrameNumber.Close
    End If

    'find all pumps with the selected temc frame number
    qyDataTEMCFrameNumber.CommandText = "SELECT DISTINCT " & _
                  " [TempPumpData]![SerialNumber], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer " & _
                  " FROM TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber" & _
                  " WHERE (((TempPumpData.TemcFrameNumber)= '" & cmbSearchTEMCFrameNumber.BoundText & "'));"

    rsDataTEMCFrameNumber.Open qyDataTEMCFrameNumber

    If rsDataTEMCFrameNumber.RecordCount = 0 Then
        Exit Sub
    End If

    'bind the datalist to the recordset
    Set fgTEMCFrameNo.DataSource = rsDataTEMCFrameNumber

    fgTEMCFrameNo.TextMatrix(0, 0) = "S/N"
    fgTEMCFrameNo.TextMatrix(0, 1) = "Date"
    fgTEMCFrameNo.TextMatrix(0, 2) = "Sales Order"
    fgTEMCFrameNo.TextMatrix(0, 3) = "Bill To"
    fgTEMCFrameNo.TextMatrix(0, 4) = "Ship To"
    fgTEMCFrameNo.ColWidth(0) = 1400
    fgTEMCFrameNo.ColWidth(1) = 2000
    fgTEMCFrameNo.ColWidth(2) = 1200
    fgTEMCFrameNo.ColWidth(3) = 3200
    fgTEMCFrameNo.ColWidth(4) = 3200

    frmCustomer.Visible = False
    frmShipTo.Visible = False
    frmTEMCFrameNo.Height = 4000
    fgTEMCFrameNo.Height = 4000 - 360
    frmTEMCFrameNo.FontBold = True
  
End Sub
Private Sub cmbSearchTEMCFrameNumber_Click(Area As Integer)
    If rsDataTEMCFrameNumber.State = adStateOpen Then
        rsDataTEMCFrameNumber.Close
    End If

    'find all pumps with the selected temc frame number
    qyDataTEMCFrameNumber.CommandText = "SELECT DISTINCT " & _
                  " [TempPumpData]![SerialNumber], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], IIF(TempTestSetupData.ImpTrimmed=0, val(TempPumpData!ImpellerDia), val(TempTestSetupData!ImpTrimmed)) as ImpDia ,  " & _
                  "  TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer" & _
                  " FROM Motor INNER JOIN ((Model INNER JOIN TempPumpData ON Model.Model = TempPumpData.Model) INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) ON Motor.Motor = TempPumpData.Motor" & _
                  " WHERE (((TempPumpData.TemcFrameNumber)= '" & cmbSearchTEMCFrameNumber.BoundText & "'));"

    rsDataTEMCFrameNumber.Open qyDataTEMCFrameNumber

    If rsDataTEMCFrameNumber.RecordCount = 0 Then
        Exit Sub
    End If

    Set fgTEMCFrameNo.DataSource = rsDataTEMCFrameNumber

    Dim f As String

    f = "<S/N     |<Date      |<Sales Order |<Model No     |^Imp Dia  |<Bill To     |<Ship To  "
    fgTEMCFrameNo.FormatString = f
    fgTEMCFrameNo.ColAlignment(4) = flexAlignCenterTop
    'fgModel.ColAlignment(5) = flexAlignCenterTop
    fgTEMCFrameNo.ColWidth(0) = 1200
    fgTEMCFrameNo.ColWidth(1) = 2000
    fgTEMCFrameNo.ColWidth(2) = 1200
    fgTEMCFrameNo.ColWidth(3) = 2000
    fgTEMCFrameNo.ColWidth(4) = 1200
    'fgModel.ColWidth(5) = 1200
    fgTEMCFrameNo.ColWidth(5) = 3200
    fgTEMCFrameNo.ColWidth(6) = 3200
    fgTEMCFrameNo.TextMatrix(0, 0) = "S/N"
    fgTEMCFrameNo.TextMatrix(0, 1) = "Date"
    fgTEMCFrameNo.TextMatrix(0, 2) = "Sales Order"
    fgTEMCFrameNo.TextMatrix(0, 3) = "Model No"
    fgTEMCFrameNo.TextMatrix(0, 4) = "Imp Dia"
    'fgModel.TextMatrix(0, 5) = "Motor"
    fgTEMCFrameNo.TextMatrix(0, 5) = "Bill To"
    fgTEMCFrameNo.TextMatrix(0, 6) = "Ship To"

    Dim x As Long
    With fgTEMCFrameNo
        For x = .FixedRows To .Rows - 1
        .TextMatrix(x, 4) = Format(.TextMatrix(x, 4), "#0.000")
        Next x
    End With


    frmCustomer.Visible = False
    frmShipTo.Visible = False
    frmTEMCFrameNo.Height = 4000
    fgTEMCFrameNo.Height = 4000 - 360
    frmTEMCFrameNo.FontBold = True
  
End Sub

Private Sub cmbStartDate_Click()
    Dim StartDate As Date
    Dim EndDate As Date
    Dim I As Integer

    Text1.Text = cmbStartDate.List(cmbStartDate.ListIndex)
    If rsDataDate.State = adStateOpen Then
        rsDataDate.Close
    End If

    StartDate = FormatDateTime(Text1.Text)

    'see if there's and end date
    If Left$(cmbSearchEndDate.BoundText, 1) = "S" Then
        EndDate = StartDate
    Else
        EndDate = FormatDateTime(cmbSearchEndDate.BoundText)
    End If


    I = InStr(EndDate, " ")
    If I <> 0 Then
        EndDate = Left$(EndDate, I)
    End If
    StartDate = StartDate & " 00:00:00"
    EndDate = EndDate & " 23:59:59"

    'look for all tests that were done on that date, regardless of the time
    qyDataDate.CommandText = "SELECT " & _
                      " [TempPumpData]![SerialNumber], [TempPumpData]![ChempumpPump], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], [TempPumpData]![TEMCFrameNumber], TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer " & _
                      " FROM (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
                      " WHERE ((TempTestSetupData.Date >= #" & StartDate & "#) AND (TempTestSetupData.Date <= #" & EndDate & "#)) " & _
                      " ORDER BY TempTestSetupData.Date;"
    qyData2.CommandText = "SELECT DISTINCT TempTestSetupData.Date, IIf(InStr(2,[TempTestSetupData]![Date],"" "")<>0,Left$([TempTestSetupData]![Date],InStr(2,[TempTestSetupData]![Date],"" "")),[TempTestSetupData]![Date]) AS [Expr2]  " & _
                     " FROM TempTestSetupData " & _
                     " WHERE Date >= #" & StartDate & "#" & _
                     " ORDER BY Date;"
'    qyData2.CommandText = "SELECT DISTINCT TempPumpData.SerialNumber, TempTestSetupData.Date, TempPumpData.SalesOrderNumber, TempPumpData.ModelNumber " & _
'       " FROM TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber " & _
'       " WHERE Date >= #" & StartDate & "#" & _
'       " ORDER BY Date;"

    If rsData2.State = adStateOpen Then
        rsData2.Close
    End If

    rsData2.Open qyData2
    Set cmbSearchEndDate.DataSource = rsData2
    cmbSearchEndDate.ListField = "Expr2"
    Set cmbSearchEndDate.RowSource = rsData2

    cmbSearchEndDate.Enabled = True

    rsDataDate.Open qyDataDate

    If rsDataDate.RecordCount = 0 Then
        Exit Sub
    End If

    lblNoOfPumps.Caption = rsDataDate.RecordCount & " Pumps Found"

    'bind the dldate datalist to the recordset
    Set fgDate.DataSource = rsDataDate

    fgDate.TextMatrix(0, 0) = "S/N"
    fgDate.TextMatrix(0, 2) = "Date"
    fgDate.TextMatrix(0, 3) = "Sales Order"
    fgDate.TextMatrix(0, 4) = "Model No"
    fgDate.TextMatrix(0, 5) = "TEMC Frame"
    fgDate.TextMatrix(0, 6) = "Bill To"
    fgDate.TextMatrix(0, 7) = "Ship To"
    fgDate.ColWidth(0) = 1400               'serial number
    fgDate.ColWidth(1) = 0               'chempumppump
    fgDate.ColWidth(2) = 2000
    fgDate.ColWidth(3) = 1200
    If rsDataDate.Fields(1) = True Then       'show model number
        fgDate.ColWidth(4) = 1800
        fgDate.ColWidth(5) = 0
    Else                                    'else, show TEMC Frame number
        fgDate.ColWidth(4) = 0
        fgDate.ColWidth(5) = 1800
    End If
    fgDate.ColWidth(6) = 3000
    fgDate.ColWidth(7) = 3000

    frmSalesOrder.Visible = False
    frmSN.Visible = False
    frmDate.Height = 4000
    fgDate.Height = 4000 - 360
    frmDate.FontBold = True
  
End Sub

Private Sub cmdClose_Click()
    Unload Me   'unload the form
End Sub

Private Sub cmdResetSizes_Click()
    frmDate.Top = 0
    frmDate.Height = 1935
    fgDate.Height = 1575
    frmDate.Visible = True
    frmDate.FontBold = False

    frmSalesOrder.Top = 1920
    frmSalesOrder.Height = 1335
    fgSalesOrder.Height = 975
    frmSalesOrder.Visible = True
    frmSalesOrder.FontBold = False

    frmSN.Top = 3240
    frmSN.Height = 1335
    fgSN.Height = 975
    frmSN.Visible = True
    frmSN.FontBold = False

    frmModel.Top = 4560
    frmModel.Height = 1335
    fgModel.Height = 975
    frmModel.Visible = True
    frmModel.FontBold = False

    frmTEMCFrameNo.Top = 5880
    frmTEMCFrameNo.Height = 1335
    fgTEMCFrameNo.Height = 975
    frmTEMCFrameNo.Visible = True
    frmTEMCFrameNo.FontBold = False

    frmCustomer.Top = 7200
    frmCustomer.Height = 1335
    fgCustomer.Height = 975
    frmCustomer.Visible = True
    frmCustomer.FontBold = False

    frmShipTo.Top = 8520
    frmShipTo.Height = 1335
    fgShipTo.Height = 1335
    frmShipTo.Visible = True
    frmShipTo.FontBold = False

    frmWildCard.Top = 9960
    frmWildCard.Height = 1335
    fgWildCard.Height = 1335
    frmWildCard.Visible = True
    frmWildCard.FontBold = False
    txtModelNumberString.Text = "Enter Characters and Search with Return"
  
End Sub

Private Sub fgCustomer_Click()
    fgCustomer.Col = 0
    frmPLCData.txtSN.Text = fgCustomer.Text
End Sub

Private Sub fgDate_Click()
    fgDate.Col = 0
    frmPLCData.txtSN.Text = fgDate.Text
End Sub
Private Sub fgmodel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

   'MSFlexGrid has the strange feature of not being able to recognize
   'when the heading of 1st fixed column is clicked on, it calls it Row 1,
   'the Row below this also returns as Row 1.
   'This bit of code below singles out the heading row which is required in
   'this app for sorting the data.

    Static LastCol As Integer       'last column clicked on
    Static direction As Boolean     'sorting ascending or descending

    If y < fgModel.RowHeight(fgModel.Row) Then  'if user clicked in header row
        fgModel.Row = 1
        fgModel.RowSel = 1
        fgModel.ColSel = fgModel.Col
        If LastCol = fgModel.Col Then   'if user clicked on same column, reverse sort
            direction = Not direction
        Else
            direction = True            'if new column, sort ascending
        End If

        If direction Then
            fgModel.Sort = flexSortGenericAscending
        Else
            fgModel.Sort = flexSortGenericDescending
        End If

        LastCol = fgModel.Col   'save column number
    Else                            'user did not click on header, select serial number for main screen
        fgModel.Col = 0
        frmPLCData.txtSN.Text = fgModel.Text
    End If
  
End Sub

Private Sub fgSalesOrder_Click()
    fgSalesOrder.Col = 0
    frmPLCData.txtSN.Text = fgSalesOrder.Text
End Sub



Private Sub fgshipto_Click()
    fgShipTo.Col = 0
    frmPLCData.txtSN.Text = fgShipTo.Text
End Sub

Private Sub fgSN_Click()
    fgSN.Col = 0
    frmPLCData.txtSN.Text = fgSN.Text
End Sub

Private Sub fgTEMCFrameNo_Click()
    fgTEMCFrameNo.Col = 0
    frmPLCData.txtSN.Text = fgTEMCFrameNo.Text
End Sub

Private Sub Form_Activate()
    Const HWND_TOPMOST As Integer = -1
    Const SWP_NOSIZE As Integer = &H1
    Const SWP_NOMOVE As Integer = &H2
    Const SWP_NOACTIVATE As Integer = &H10
    Const SWP_SHOWWINDOW As Integer = &H40

    'window always on top
'    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE
  
End Sub

Private Sub Form_Load()
    'open several recordsets for searching

    'qydata/rsdata is for the serial number dropdown
    qyData.ActiveConnection = cnPumpData
    qyData.CommandText = "SELECT DISTINCT SerialNumber FROM TempPumpData ORDER BY SerialNumber;"
    rsData.CursorType = adOpenStatic
    rsData.CursorLocation = adUseClient
    rsData.Index = "SerialNumber"
    rsData.Open qyData

    'bind the serial number dropdown
    Set cmbSearchSN.DataSource = rsData
    cmbSearchSN.ListField = "SerialNumber"
    Set cmbSearchSN.RowSource = rsData

    'qydata1 and 2/rsdata1 and 2 is for the date dropdown
    qyData1.ActiveConnection = cnPumpData
    rsData1.CursorType = adOpenStatic
    rsData1.CursorLocation = adUseClient
    rsData1.Index = "SerialNumber"

    qyData2.ActiveConnection = cnPumpData
    rsData2.CursorType = adOpenStatic
    rsData2.CursorLocation = adUseClient
    rsData2.Index = "SerialNumber"

    'find dates without times
'    qyData1.CommandText = "SELECT DISTINCT TempPumpData.SerialNumber, TempPumpData.Model, TempTestSetupData.Date, IIf(InStr(2,[TempTestSetupData]![Date],"" "")<>0,Left$([TempTestSetupData]![Date],InStr(2,[TempTestSetupData]![Date],"" "")),[TempTestSetupData]![Date]) AS [Expr2]" & _
'        " FROM TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber ORDER BY Date;"
    qyData1.CommandText = "SELECT DISTINCT TempTestSetupData.Date, IIf(InStr(2,[TempTestSetupData]![Date],"" "")<>0,Left$([TempTestSetupData]![Date],InStr(2,[TempTestSetupData]![Date],"" "")),[TempTestSetupData]![Date]) AS [Expr2] " & _
                      " FROM TempTestSetupData ORDER BY Date;"
    rsData1.Open qyData1

    Dim I As Integer
    Dim TempDate As Date
    Dim LastDate As Date
    If Not rsData1.BOF Then
        rsData1.MoveFirst
    End If

    LastDate = FormatDateTime(Now(), vbShortDate)
    For I = 1 To rsData1.RecordCount
        TempDate = FormatDateTime(rsData1.Fields(0), vbShortDate)
        If TempDate <> LastDate Then
            cmbStartDate.AddItem TempDate
            LastDate = TempDate
        End If
        rsData1.MoveNext
    Next I

    'qydatadate/rsdatadate for date datalist
    qyDataDate.ActiveConnection = cnPumpData
    rsDataDate.CursorType = adOpenStatic
    rsDataDate.CursorLocation = adUseClient

    'qydatamodel/rsdatamodel for model dropdown
    qyDataModel.ActiveConnection = cnPumpData
    rsDataModel.CursorType = adOpenStatic
    rsDataModel.CursorLocation = adUseClient

    'qydatasalesorder/rsdatasalesorder for sales order dropdown
    qyDataSalesOrder.ActiveConnection = cnPumpData
    rsDataSalesOrder.CursorType = adOpenStatic
    rsDataSalesOrder.CursorLocation = adUseClient
    qyDataSalesOrder.CommandText = "SELECT DISTINCT TempPumpData.SalesOrderNumber FROM TempPumpData ORDER BY TempPumpData.SalesOrderNumber;"
    rsDataSalesOrder.Open qyDataSalesOrder

    qySalesOrderData.ActiveConnection = cnPumpData
    rsSalesOrderData.CursorType = adOpenStatic
    rsSalesOrderData.CursorLocation = adUseClient

    'bind to temc frame number dropdown
    Set cmbSearchSalesOrder.RowSource = rsDataSalesOrder
    cmbSearchSalesOrder.ListField = "SalesOrderNumber"
    Set cmbSearchSalesOrder.RowSource = rsDataSalesOrder

    'qydatasn/rsdatasn for serial numbers
    qyDataSN.ActiveConnection = cnPumpData
    rsDataSN.CursorType = adOpenStatic
    rsDataSN.CursorLocation = adUseClient

    'qydatatemcmodel/rsdatatemcmodel for temc frame number
    qyDataTEMCModel.ActiveConnection = cnPumpData
    rsDataTEMCModel.CursorType = adOpenStatic
    rsDataTEMCModel.CursorLocation = adUseClient
    qyDataTEMCModel.CommandText = "SELECT DISTINCT TempPumpData.TEMCFrameNumber FROM TempPumpData WHERE (TempPumpData.ChempumpPump = FALSE) ORDER BY TempPumpData.TEMCFrameNumber;"
    rsDataTEMCModel.Open qyDataTEMCModel

    'bind to temc frame number dropdown
    Set cmbSearchTEMCFrameNumber.RowSource = rsDataTEMCModel
    cmbSearchTEMCFrameNumber.ListField = "TEMCFrameNumber"
    Set cmbSearchTEMCFrameNumber.RowSource = rsDataTEMCModel

    qyDataTEMCFrameNumber.ActiveConnection = cnPumpData
    rsDataTEMCFrameNumber.CursorType = adOpenStatic
    rsDataTEMCFrameNumber.CursorLocation = adUseClient

    'customer
    qyDataCustomer.ActiveConnection = cnPumpData
    rsDataCustomer.CursorType = adOpenStatic
    rsDataCustomer.CursorLocation = adUseClient

    qyDataCustomer.CommandText = "SELECT DISTINCT TempPumpData.BillToCustomer FROM TempPumpData ORDER BY TempPumpData.BillToCustomer;"
    rsDataCustomer.Open qyDataCustomer

    qyCustomerData.ActiveConnection = cnPumpData
    rsCustomerData.CursorType = adOpenStatic
    rsCustomerData.CursorLocation = adUseClient

    'bind to customer dropdown
    Set cmbSearchCustomer.RowSource = rsDataCustomer
    cmbSearchCustomer.ListField = "BillToCustomer"

    ' ship to customer
    qyDataShipto.ActiveConnection = cnPumpData
    rsDataShipTo.CursorType = adOpenStatic
    rsDataShipTo.CursorLocation = adUseClient

    qyDataShipto.CommandText = "SELECT DISTINCT TempPumpData.shipToCustomer FROM TempPumpData ORDER BY TempPumpData.shipToCustomer;"
    rsDataShipTo.Open qyDataShipto

    qyShipToData.ActiveConnection = cnPumpData
    rsShipToData.CursorType = adOpenStatic
    rsShipToData.CursorLocation = adUseClient

    'bind to customer dropdown
    Set cmbSearchShipTo.RowSource = rsDataShipTo
    cmbSearchShipTo.ListField = "ShipToCustomer"

    cmbSearchEndDate.Enabled = False
  
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'close all the datasets and release the connections
    If rsData.State = adStateOpen Then
        rsData.Close
    End If
    If rsData1.State = adStateOpen Then
        rsData1.Close
    End If
    If rsData2.State = adStateOpen Then
        rsData2.Close
    End If
    If rsDataDate.State = adStateOpen Then
        rsDataDate.Close
    End If
    If rsDataModel.State = adStateOpen Then
        rsDataModel.Close
    End If
    If rsDataSalesOrder.State = adStateOpen Then
        rsDataSalesOrder.Close
    End If
    If rsSalesOrderData.State = adStateOpen Then
        rsSalesOrderData.Close
    End If
    If rsDataSN.State = adStateOpen Then
        rsDataSN.Close
    End If
    If rsDataTEMCModel.State = adStateOpen Then
        rsDataTEMCModel.Close
    End If
    If rsDataTEMCFrameNumber.State = adStateOpen Then
        rsDataTEMCFrameNumber.Close
    End If
    If rsDataCustomer.State = adStateOpen Then
        rsDataCustomer.Close
    End If
    If rsCustomerData.State = adStateOpen Then
        rsCustomerData.Close
    End If
    If rsDataShipTo.State = adStateOpen Then
        rsDataShipTo.Close
    End If
    If rsShipToData.State = adStateOpen Then
        rsShipToData.Close
    End If

    Set rsData = Nothing
    Set rsData1 = Nothing
    Set rsData2 = Nothing
    Set rsDataDate = Nothing
    Set rsDataModel = Nothing
    Set rsDataSalesOrder = Nothing
    Set rsSalesOrderData = Nothing
    Set rsDataSN = Nothing
    Set rsDataTEMCModel = Nothing
    Set rsDataTEMCFrameNumber = Nothing
    Set rsDataCustomer = Nothing
    Set rsCustomerData = Nothing
    Set rsDataShipTo = Nothing
    Set rsShipToData = Nothing
  
End Sub

Private Sub WildCardSearch()
    If rsDataModel.State = adStateOpen Then
        rsDataModel.Close
    End If

    qyDataModel.CommandText = "SELECT DISTINCT " & _
                  " [TempPumpData]![SerialNumber], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], IIF(TempTestSetupData.ImpTrimmed=0, val(TempPumpData!ImpellerDia), val(TempTestSetupData!ImpTrimmed)) as ImpDia ,  " & _
                  "  TempPumpData.BillToCustomer,  Model.Description " & _
                  " FROM Motor INNER JOIN ((Model INNER JOIN TempPumpData ON Model.Model = TempPumpData.Model) INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) ON Motor.Motor = TempPumpData.Motor" & _
                  " WHERE (((TempPumpData.ModelNumber) LIKE '%" & txtModelNumberString.Text & "%'));"


    rsDataModel.Open qyDataModel

    If rsDataModel.RecordCount = 0 Then
        Exit Sub
    End If

    Set fgWildCard.DataSource = rsDataModel

    Dim f As String

    f = "<S/N     |<Date      |<Sales Order |<Model No     |^Imp Dia  |<Bill To     |<Ship To  "
    fgWildCard.FormatString = f
    fgWildCard.ColAlignment(4) = flexAlignCenterTop
    'fgModel.ColAlignment(5) = flexAlignCenterTop
    fgWildCard.ColWidth(0) = 1200
    fgWildCard.ColWidth(1) = 2000
    fgWildCard.ColWidth(2) = 1200
    fgWildCard.ColWidth(3) = 2000
    fgWildCard.ColWidth(4) = 1200
    'fgModel.ColWidth(5) = 1200
    fgWildCard.ColWidth(5) = 3200
    fgWildCard.ColWidth(6) = 3200
    fgWildCard.TextMatrix(0, 0) = "S/N"
    fgWildCard.TextMatrix(0, 1) = "Date"
    fgWildCard.TextMatrix(0, 2) = "Sales Order"
    fgWildCard.TextMatrix(0, 3) = "Model No"
    fgWildCard.TextMatrix(0, 4) = "Imp Dia"
    'fgModel.TextMatrix(0, 5) = "Motor"
    fgWildCard.TextMatrix(0, 5) = "Bill To"
    fgWildCard.TextMatrix(0, 6) = "Ship To"

    Dim x As Long
    With fgWildCard
        For x = .FixedRows To .Rows - 1
        .TextMatrix(x, 4) = Format(.TextMatrix(x, 4), "#0.000")
        Next x
    End With


    frmTEMCFrameNo.Visible = False
    frmCustomer.Visible = False
    frmWildCard.Top = 4560
    frmWildCard.Height = 4000
    fgWildCard.Height = 4000 - 360
    frmWildCard.FontBold = True
  
End Sub

Private Sub txtModelNumberString_KeyPress(KeyAscii As Integer)
    If KeyAscii = 13 Then
        WildCardSearch
    End If
End Sub

Private Sub fgwildcard_Click()
    fgWildCard.Col = 0
    frmPLCData.txtSN.Text = fgWildCard.Text
End Sub

Private Sub txtModelNumberString_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If txtModelNumberString.Text = "Enter Characters and Search with Return" Then
        txtModelNumberString.Text = ""
    End If
  
End Sub
Private Sub fgDate_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If this is not row 0, do nothing.
    If fgDate.MouseRow <> 0 Then Exit Sub
    Static direction As Boolean
    With fgDate
        .ColSel = fgDate.MouseCol

        If direction = True Then
            .Sort = flexSortStringAscending
        Else
            .Sort = flexSortStringDescending
        End If
        direction = Not direction
    End With
 
End Sub
Private Sub fgSalesOrder_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If this is not row 0, do nothing.
    If fgSalesOrder.MouseRow <> 0 Then Exit Sub
    Static direction As Boolean
    With fgSalesOrder
        .ColSel = fgSalesOrder.MouseCol
        If direction = True Then
            .Sort = flexSortStringAscending
        Else
            .Sort = flexSortStringDescending
        End If
        direction = Not direction
    End With
 
End Sub

Private Sub fgSN_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If this is not row 0, do nothing.
    If fgSN.MouseRow <> 0 Then Exit Sub
    Static direction As Boolean
    With fgSN
        .ColSel = fgSN.MouseCol
        If direction = True Then
            .Sort = flexSortStringAscending
        Else
            .Sort = flexSortStringDescending
        End If
        direction = Not direction
    End With
 
End Sub

Private Sub fgModel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If this is not row 0, do nothing.
    If fgModel.MouseRow <> 0 Then Exit Sub
    Static direction As Boolean
    With fgModel
        .ColSel = fgModel.MouseCol
        If direction = True Then
            .Sort = flexSortStringAscending
        Else
            .Sort = flexSortStringDescending
        End If
        direction = Not direction
    End With
 
End Sub

Private Sub fgTEMCFrameNo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If this is not row 0, do nothing.
    If fgTEMCFrameNo.MouseRow <> 0 Then Exit Sub
    Static direction As Boolean
    Dim HeaderOrg As String
    
    With fgTEMCFrameNo
        .ColSel = fgTEMCFrameNo.MouseCol
'        .TextMatrix(0, .ColSel) = Replace(.TextMatrix(0, .ColSel), " " & ChrW(&H25B2), "")
'        .TextMatrix(0, .ColSel) = Replace(.TextMatrix(0, .ColSel), " " & ChrW(&H25BC), "")
        
        If direction = True Then
            .Sort = flexSortStringAscending
'            .TextMatrix(0, .ColSel) = .TextMatrix(0, .ColSel) & " " & ChrW(&H25BC)
        Else
            .Sort = flexSortStringDescending
'            .TextMatrix(0, .ColSel) = .TextMatrix(0, .ColSel) & " " & ChrW(&H25B2)
        End If
        direction = Not direction
    End With
 
End Sub

Private Sub fgCustomer_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If this is not row 0, do nothing.
    If fgCustomer.MouseRow <> 0 Then Exit Sub
    Static direction As Boolean
    With fgCustomer
        .ColSel = fgCustomer.MouseCol
        If direction = True Then
            .Sort = flexSortStringAscending
        Else
            .Sort = flexSortStringDescending
        End If
        direction = Not direction
    End With
 
End Sub

Private Sub fgShipTo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If this is not row 0, do nothing.
    If fgShipTo.MouseRow <> 0 Then Exit Sub
    Static direction As Boolean
    With fgShipTo
        .ColSel = fgShipTo.MouseCol
        If direction = True Then
            .Sort = flexSortStringAscending
        Else
            .Sort = flexSortStringDescending
        End If
        direction = Not direction
    End With
 
End Sub

Private Sub fgWildCard_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    ' If this is not row 0, do nothing.
    If fgWildCard.MouseRow <> 0 Then Exit Sub
    Static direction As Boolean
    With fgWildCard
        .ColSel = fgWildCard.MouseCol
        If direction = True Then
            .Sort = flexSortStringAscending
        Else
            .Sort = flexSortStringDescending
        End If
        direction = Not direction
    End With
 
End Sub

