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
' <VB WATCH>
Const VBWMODULE = "frmSearch"
' </VB WATCH>

Private Sub cmbSalesOrder_Click(Area As Integer)
' <VB WATCH>
1          On Error GoTo vbwErrHandler
' </VB WATCH>

' <VB WATCH>
2          Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbSalesOrder_Click"

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

Private Sub cmbSearchCustomer_Click(Area As Integer)
' <VB WATCH>
3          On Error GoTo vbwErrHandler
' </VB WATCH>
4          If rsCustomerData.State = adStateOpen Then
5              rsCustomerData.Close
6          End If

7          If cmbSearchCustomer.SelectedItem = 1 Then
8              Exit Sub
9          End If

10         qyCustomerData.CommandText = "SELECT DISTINCT " & _
                         " [TempPumpData]![SerialNumber], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], TempPumpData.ShiptoCustomer " & _
                         " FROM (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
                         " WHERE (((TempPumpData.BillToCustomer)= '" & cmbSearchCustomer.BoundText & "'));"

11         rsCustomerData.Open qyCustomerData

12         If rsCustomerData.RecordCount = 0 Then
13             Exit Sub
14         End If

           'bind the datalist to the recordset
15         Set fgCustomer.DataSource = rsCustomerData

16         fgCustomer.ColWidth(0) = 1400
17         fgCustomer.ColWidth(1) = 2000
18         fgCustomer.ColWidth(2) = 1200
19         fgCustomer.ColWidth(3) = 2000
20         fgCustomer.ColWidth(4) = 3200
21         fgCustomer.TextMatrix(0, 0) = "S/N"
22         fgCustomer.TextMatrix(0, 1) = "Date"
23         fgCustomer.TextMatrix(0, 2) = "Sales Order"
24         fgCustomer.TextMatrix(0, 3) = "Model No"
25         fgCustomer.TextMatrix(0, 4) = "Ship To"

26         frmModel.Visible = False
27         frmTEMCFrameNo.Visible = False
28         frmCustomer.Top = 7200 - (4000 - 1335)
29         frmCustomer.Height = 4000
30         fgCustomer.Height = 4000 - 360
31         frmCustomer.FontBold = True

' <VB WATCH>
32         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbSearchCustomer_Click"

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


Private Sub cmbSearchEndDate_Click(Area As Integer)
' <VB WATCH>
33         On Error GoTo vbwErrHandler
' </VB WATCH>
34         cmbStartDate_Click
' <VB WATCH>
35         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbSearchEndDate_Click"

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

Private Sub cmbSearchModel_Click()
' <VB WATCH>
36         On Error GoTo vbwErrHandler
' </VB WATCH>
37         If rsDataModel.State = adStateOpen Then
38             rsDataModel.Close
39         End If

40         qyDataModel.CommandText = "SELECT DISTINCT " & _
                         " [TempPumpData]![SerialNumber], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], IIF(TempTestSetupData.ImpTrimmed=0, val(TempPumpData!ImpellerDia), val(TempTestSetupData!ImpTrimmed)) as ImpDia ,  " & _
                         "  TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer" & _
                         " FROM Motor INNER JOIN ((Model INNER JOIN TempPumpData ON Model.Model = TempPumpData.Model) INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) ON Motor.Motor = TempPumpData.Motor" & _
                         " WHERE (((TempPumpData.ModelNumber) LIKE '%" & cmbSearchModel.List(cmbSearchModel.ListIndex) & "%'));"
       '                  " WHERE (((TempPumpData.Model)= " & cmbSearchModel.ItemData(cmbSearchModel.ListIndex) & "));"

41         rsDataModel.Open qyDataModel

42         If rsDataModel.RecordCount = 0 Then
43             Exit Sub
44         End If

45         Set fgModel.DataSource = rsDataModel

46         Dim f As String

47         f = "<S/N     |<Date      |<Sales Order |<Model No     |^Imp Dia  |<Bill To     |<Ship To  "
48         fgModel.FormatString = f
49         fgModel.ColAlignment(4) = flexAlignCenterTop
           'fgModel.ColAlignment(5) = flexAlignCenterTop
50         fgModel.ColWidth(0) = 1200
51         fgModel.ColWidth(1) = 2000
52         fgModel.ColWidth(2) = 1200
53         fgModel.ColWidth(3) = 2000
54         fgModel.ColWidth(4) = 1200
           'fgModel.ColWidth(5) = 1200
55         fgModel.ColWidth(5) = 3200
56         fgModel.ColWidth(6) = 3200
57         fgModel.TextMatrix(0, 0) = "S/N"
58         fgModel.TextMatrix(0, 1) = "Date"
59         fgModel.TextMatrix(0, 2) = "Sales Order"
60         fgModel.TextMatrix(0, 3) = "Model No"
61         fgModel.TextMatrix(0, 4) = "Imp Dia"
           'fgModel.TextMatrix(0, 5) = "Motor"
62         fgModel.TextMatrix(0, 5) = "Bill To"
63         fgModel.TextMatrix(0, 6) = "Ship To"

64         Dim x As Long
65         With fgModel
66             For x = .FixedRows To .Rows - 1
67             .TextMatrix(x, 4) = Format(.TextMatrix(x, 4), "#0.000")
68             Next x
69         End With


70         frmTEMCFrameNo.Visible = False
71         frmCustomer.Visible = False
72         frmModel.Height = 4000
73         fgModel.Height = 4000 - 360
74         frmModel.FontBold = True

' <VB WATCH>
75         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbSearchModel_Click"

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

Private Sub cmbSearchSalesOrder_Change()
' <VB WATCH>
76         On Error GoTo vbwErrHandler
' </VB WATCH>

77         Text1.Text = cmbSearchSalesOrder.BoundText

78         If rsSalesOrderData.State = adStateOpen Then
79             rsSalesOrderData.Close
80         End If

           'find all dates and models for the selected serial number
81         qySalesOrderData.CommandText = "SELECT DISTINCT " & _
                         " [TempPumpData]![SerialNumber], [TempPumpData]![ChempumpPump], [TempTestSetupData]![Date], [TempPumpData]![ModelNumber], [TempPumpData]![TEMCFrameNumber], TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer " & _
                         " FROM (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
                         " WHERE (((TempPumpData.SalesOrderNumber)= '" & cmbSearchSalesOrder.BoundText & "'));"

82         rsSalesOrderData.Open qySalesOrderData

83         If rsSalesOrderData.RecordCount = 0 Then
84             Exit Sub
85         End If

       '    'bind the datalist to the other recordset
86         Set fgSalesOrder.DataSource = rsSalesOrderData

87         fgSalesOrder.ColWidth(0) = 1400               'serial number
88         fgSalesOrder.ColWidth(1) = 0               'chempumppump
89         fgSalesOrder.ColWidth(2) = 2000
90         If rsSalesOrderData.Fields(1) = True Then       'show model number
91             fgSalesOrder.ColWidth(3) = 1800
92             fgSalesOrder.ColWidth(4) = 0
93         Else                                    'else, show TEMC Frame number
94             fgSalesOrder.ColWidth(3) = 0
95             fgSalesOrder.ColWidth(4) = 1800
96         End If
97         fgSalesOrder.ColWidth(5) = 3200
98         fgSalesOrder.ColWidth(6) = 3200
99         fgSalesOrder.TextMatrix(0, 0) = "S/N"
100        fgSalesOrder.TextMatrix(0, 2) = "Date"
101        fgSalesOrder.TextMatrix(0, 3) = "Model No"
102        fgSalesOrder.TextMatrix(0, 4) = "TEMC Frame"
103        fgSalesOrder.TextMatrix(0, 5) = "Bill To"
104        fgSalesOrder.TextMatrix(0, 6) = "Ship To"

105        frmSN.Visible = False
106        frmSalesOrder.Height = 4000
107        fgSalesOrder.Height = 4000 - 360
108        frmSalesOrder.FontBold = True

' <VB WATCH>
109        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbSearchSalesOrder_Change"

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

Private Sub cmbSearchShipTo_Click(Area As Integer)
' <VB WATCH>
110        On Error GoTo vbwErrHandler
' </VB WATCH>

111        If rsShipToData.State = adStateOpen Then
112            rsShipToData.Close
113        End If

114        If cmbSearchShipTo.SelectedItem = 1 Then
115            Exit Sub
116        End If

117        qyShipToData.CommandText = "SELECT DISTINCT " & _
                         " [TempPumpData]![SerialNumber],[TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], TempPumpData.ModelNumber, TempPumpData.BillToCustomer " & _
                         " FROM  (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
                         " WHERE (((TempPumpData.ShipToCustomer)= '" & cmbSearchShipTo.BoundText & "'));"

118        rsShipToData.Open qyShipToData

119        If rsShipToData.RecordCount = 0 Then
120            Exit Sub
121        End If

122        Set fgShipTo.DataSource = rsShipToData


123        fgShipTo.ColWidth(0) = 1400
124        fgShipTo.ColWidth(1) = 2000
125        fgShipTo.ColWidth(2) = 1200
126        fgShipTo.ColWidth(3) = 2000
127        fgShipTo.ColWidth(4) = 3200
128        fgShipTo.TextMatrix(0, 0) = "S/N"
129        fgShipTo.TextMatrix(0, 1) = "Date"
130        fgShipTo.TextMatrix(0, 2) = "Sales Order"
131        fgShipTo.TextMatrix(0, 3) = "Model No"
132        fgShipTo.TextMatrix(0, 4) = "Bill To"

133        frmTEMCFrameNo.Visible = False
134        frmCustomer.Visible = False
135        frmShipTo.Top = 8520 - (4000 - 1335)
136        frmShipTo.Height = 4000
137        fgShipTo.Height = 4000 - 360
138        frmShipTo.FontBold = True

' <VB WATCH>
139        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbSearchShipTo_Click"

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

Private Sub cmbSearchSN_Change()
' <VB WATCH>
140        On Error GoTo vbwErrHandler
' </VB WATCH>

141        Text1.Text = cmbSearchSN.BoundText

142        If rsDataSN.State = adStateOpen Then
143            rsDataSN.Close
144        End If

           'find all dates and models for the selected serial number
145        qyDataSN.CommandText = "SELECT DISTINCT TempPumpData.SerialNumber," & _
                         " [TempPumpData]![ChempumpPump],[TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], [TempPumpData]![TEMCFrameNumber], TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer " & _
                         " FROM (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
                         " WHERE (((TempPumpData.SerialNumber)= '" & cmbSearchSN.BoundText & "'));"

146        rsDataSN.Open qyDataSN

           'if we didn't find any records, see if we have any serial numbers that are close
147        If rsDataSN.RecordCount = 0 Then
148            rsDataSN.Close
149            qyDataSN.CommandText = "SELECT DISTINCT TempPumpData.SerialNumber," & _
                             " [TempPumpData]![ChempumpPump],[TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], [TempPumpData]![TEMCFrameNumber], TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer " & _
                             " FROM (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
                             " WHERE (((TempPumpData.SerialNumber)= '" & cmbSearchSN.BoundText & "%'));"
150            rsDataSN.Open qyDataSN
151        End If

152        If rsDataSN.RecordCount = 0 Then
153            Exit Sub
154        End If

155        lblNoOfPumps.Caption = rsDataSN.RecordCount & " Pumps Found"

           'bind the datalist to the other recordset
156        Set fgSN.DataSource = rsDataSN

157        fgSN.ColWidth(0) = 0               'serial number
158        fgSN.ColWidth(1) = 0               'chempumppump
159        fgSN.ColWidth(2) = 2000
160        fgSN.ColWidth(3) = 1200
161        If rsDataSN.Fields(1) = True Then       'show model number
162            fgSN.ColWidth(4) = 1800
163            fgSN.ColWidth(5) = 0
164        Else                                    'else, show TEMC Frame number
165            fgSN.ColWidth(4) = 0
166            fgSN.ColWidth(5) = 1800
167        End If
168        fgSN.ColWidth(6) = 3000
169        fgSN.ColWidth(7) = 3000
170        fgSN.TextMatrix(0, 2) = "Date"
171        fgSN.TextMatrix(0, 3) = "Sales Order"
172        fgSN.TextMatrix(0, 4) = "Model No"
173        fgSN.TextMatrix(0, 5) = "TEMC Frame"
174        fgSN.TextMatrix(0, 6) = "Bill To"
175        fgSN.TextMatrix(0, 7) = "Ship To"

           'put the serial number into the find pump textbox
176        frmPLCData.txtSN.Text = rsDataSN.Fields("SerialNumber")

177        frmModel.Visible = False
178        frmSN.Height = 4000
179        fgSN.Height = 4000 - 360
180        frmSN.FontBold = True

' <VB WATCH>
181        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbSearchSN_Change"

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

Private Sub ORIGINALcmbSearchTEMCFrameNumber_Click(Area As Integer)
' <VB WATCH>
182        On Error GoTo vbwErrHandler
' </VB WATCH>
183        If rsDataTEMCFrameNumber.State = adStateOpen Then
184            rsDataTEMCFrameNumber.Close
185        End If

           'find all pumps with the selected temc frame number
186        qyDataTEMCFrameNumber.CommandText = "SELECT DISTINCT " & _
                         " [TempPumpData]![SerialNumber], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer " & _
                         " FROM TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber" & _
                         " WHERE (((TempPumpData.TemcFrameNumber)= '" & cmbSearchTEMCFrameNumber.BoundText & "'));"

187        rsDataTEMCFrameNumber.Open qyDataTEMCFrameNumber

188        If rsDataTEMCFrameNumber.RecordCount = 0 Then
189            Exit Sub
190        End If

           'bind the datalist to the recordset
191        Set fgTEMCFrameNo.DataSource = rsDataTEMCFrameNumber

192        fgTEMCFrameNo.TextMatrix(0, 0) = "S/N"
193        fgTEMCFrameNo.TextMatrix(0, 1) = "Date"
194        fgTEMCFrameNo.TextMatrix(0, 2) = "Sales Order"
195        fgTEMCFrameNo.TextMatrix(0, 3) = "Bill To"
196        fgTEMCFrameNo.TextMatrix(0, 4) = "Ship To"
197        fgTEMCFrameNo.ColWidth(0) = 1400
198        fgTEMCFrameNo.ColWidth(1) = 2000
199        fgTEMCFrameNo.ColWidth(2) = 1200
200        fgTEMCFrameNo.ColWidth(3) = 3200
201        fgTEMCFrameNo.ColWidth(4) = 3200

202        frmCustomer.Visible = False
203        frmShipTo.Visible = False
204        frmTEMCFrameNo.Height = 4000
205        fgTEMCFrameNo.Height = 4000 - 360
206        frmTEMCFrameNo.FontBold = True

' <VB WATCH>
207        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ORIGINALcmbSearchTEMCFrameNumber_Click"

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
Private Sub cmbSearchTEMCFrameNumber_Click(Area As Integer)
' <VB WATCH>
208        On Error GoTo vbwErrHandler
' </VB WATCH>
209        If rsDataTEMCFrameNumber.State = adStateOpen Then
210            rsDataTEMCFrameNumber.Close
211        End If

           'find all pumps with the selected temc frame number
212        qyDataTEMCFrameNumber.CommandText = "SELECT DISTINCT " & _
                         " [TempPumpData]![SerialNumber], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], IIF(TempTestSetupData.ImpTrimmed=0, val(TempPumpData!ImpellerDia), val(TempTestSetupData!ImpTrimmed)) as ImpDia ,  " & _
                         "  TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer" & _
                         " FROM Motor INNER JOIN ((Model INNER JOIN TempPumpData ON Model.Model = TempPumpData.Model) INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) ON Motor.Motor = TempPumpData.Motor" & _
                         " WHERE (((TempPumpData.TemcFrameNumber)= '" & cmbSearchTEMCFrameNumber.BoundText & "'));"

213        rsDataTEMCFrameNumber.Open qyDataTEMCFrameNumber

214        If rsDataTEMCFrameNumber.RecordCount = 0 Then
215            Exit Sub
216        End If

217        Set fgTEMCFrameNo.DataSource = rsDataTEMCFrameNumber

218        Dim f As String

219        f = "<S/N     |<Date      |<Sales Order |<Model No     |^Imp Dia  |<Bill To     |<Ship To  "
220        fgTEMCFrameNo.FormatString = f
221        fgTEMCFrameNo.ColAlignment(4) = flexAlignCenterTop
           'fgModel.ColAlignment(5) = flexAlignCenterTop
222        fgTEMCFrameNo.ColWidth(0) = 1200
223        fgTEMCFrameNo.ColWidth(1) = 2000
224        fgTEMCFrameNo.ColWidth(2) = 1200
225        fgTEMCFrameNo.ColWidth(3) = 2000
226        fgTEMCFrameNo.ColWidth(4) = 1200
           'fgModel.ColWidth(5) = 1200
227        fgTEMCFrameNo.ColWidth(5) = 3200
228        fgTEMCFrameNo.ColWidth(6) = 3200
229        fgTEMCFrameNo.TextMatrix(0, 0) = "S/N"
230        fgTEMCFrameNo.TextMatrix(0, 1) = "Date"
231        fgTEMCFrameNo.TextMatrix(0, 2) = "Sales Order"
232        fgTEMCFrameNo.TextMatrix(0, 3) = "Model No"
233        fgTEMCFrameNo.TextMatrix(0, 4) = "Imp Dia"
           'fgModel.TextMatrix(0, 5) = "Motor"
234        fgTEMCFrameNo.TextMatrix(0, 5) = "Bill To"
235        fgTEMCFrameNo.TextMatrix(0, 6) = "Ship To"

236        Dim x As Long
237        With fgTEMCFrameNo
238            For x = .FixedRows To .Rows - 1
239            .TextMatrix(x, 4) = Format(.TextMatrix(x, 4), "#0.000")
240            Next x
241        End With


242        frmCustomer.Visible = False
243        frmShipTo.Visible = False
244        frmTEMCFrameNo.Height = 4000
245        fgTEMCFrameNo.Height = 4000 - 360
246        frmTEMCFrameNo.FontBold = True

' <VB WATCH>
247        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbSearchTEMCFrameNumber_Click"

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

Private Sub cmbStartDate_Click()
' <VB WATCH>
248        On Error GoTo vbwErrHandler
' </VB WATCH>
249        Dim StartDate As Date
250        Dim EndDate As Date
251        Dim I As Integer

252        Text1.Text = cmbStartDate.List(cmbStartDate.ListIndex)
253        If rsDataDate.State = adStateOpen Then
254            rsDataDate.Close
255        End If

256        StartDate = FormatDateTime(Text1.Text)

           'see if there's and end date
257        If Left$(cmbSearchEndDate.BoundText, 1) = "S" Then
258            EndDate = StartDate
259        Else
260            EndDate = FormatDateTime(cmbSearchEndDate.BoundText)
261        End If


262        I = InStr(EndDate, " ")
263        If I <> 0 Then
264            EndDate = Left$(EndDate, I)
265        End If
266        StartDate = StartDate & " 00:00:00"
267        EndDate = EndDate & " 23:59:59"

           'look for all tests that were done on that date, regardless of the time
268        qyDataDate.CommandText = "SELECT " & _
                             " [TempPumpData]![SerialNumber], [TempPumpData]![ChempumpPump], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], [TempPumpData]![TEMCFrameNumber], TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer " & _
                             " FROM (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
                             " WHERE ((TempTestSetupData.Date >= #" & StartDate & "#) AND (TempTestSetupData.Date <= #" & EndDate & "#)) " & _
                             " ORDER BY TempTestSetupData.Date;"
269        qyData2.CommandText = "SELECT DISTINCT TempTestSetupData.Date, IIf(InStr(2,[TempTestSetupData]![Date],"" "")<>0,Left$([TempTestSetupData]![Date],InStr(2,[TempTestSetupData]![Date],"" "")),[TempTestSetupData]![Date]) AS [Expr2]  " & _
                            " FROM TempTestSetupData " & _
                            " WHERE Date >= #" & StartDate & "#" & _
                            " ORDER BY Date;"
       '    qyData2.CommandText = "SELECT DISTINCT TempPumpData.SerialNumber, TempTestSetupData.Date, TempPumpData.SalesOrderNumber, TempPumpData.ModelNumber " & _
'       " FROM TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber " & _
'       " WHERE Date >= #" & StartDate & "#" & _
'       " ORDER BY Date;"

270        If rsData2.State = adStateOpen Then
271            rsData2.Close
272        End If

273        rsData2.Open qyData2
274        Set cmbSearchEndDate.DataSource = rsData2
275        cmbSearchEndDate.ListField = "Expr2"
276        Set cmbSearchEndDate.RowSource = rsData2

277        cmbSearchEndDate.Enabled = True

278        rsDataDate.Open qyDataDate

279        If rsDataDate.RecordCount = 0 Then
280            Exit Sub
281        End If

282        lblNoOfPumps.Caption = rsDataDate.RecordCount & " Pumps Found"

           'bind the dldate datalist to the recordset
283        Set fgDate.DataSource = rsDataDate

284        fgDate.TextMatrix(0, 0) = "S/N"
285        fgDate.TextMatrix(0, 2) = "Date"
286        fgDate.TextMatrix(0, 3) = "Sales Order"
287        fgDate.TextMatrix(0, 4) = "Model No"
288        fgDate.TextMatrix(0, 5) = "TEMC Frame"
289        fgDate.TextMatrix(0, 6) = "Bill To"
290        fgDate.TextMatrix(0, 7) = "Ship To"
291        fgDate.ColWidth(0) = 1400               'serial number
292        fgDate.ColWidth(1) = 0               'chempumppump
293        fgDate.ColWidth(2) = 2000
294        fgDate.ColWidth(3) = 1200
295        If rsDataDate.Fields(1) = True Then       'show model number
296            fgDate.ColWidth(4) = 1800
297            fgDate.ColWidth(5) = 0
298        Else                                    'else, show TEMC Frame number
299            fgDate.ColWidth(4) = 0
300            fgDate.ColWidth(5) = 1800
301        End If
302        fgDate.ColWidth(6) = 3000
303        fgDate.ColWidth(7) = 3000

304        frmSalesOrder.Visible = False
305        frmSN.Visible = False
306        frmDate.Height = 4000
307        fgDate.Height = 4000 - 360
308        frmDate.FontBold = True

' <VB WATCH>
309        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbStartDate_Click"

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

Private Sub cmdClose_Click()
' <VB WATCH>
310        On Error GoTo vbwErrHandler
' </VB WATCH>
311        Unload Me   'unload the form
' <VB WATCH>
312        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdClose_Click"

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

Private Sub cmdResetSizes_Click()
' <VB WATCH>
313        On Error GoTo vbwErrHandler
' </VB WATCH>
314        frmDate.Top = 0
315        frmDate.Height = 1935
316        fgDate.Height = 1575
317        frmDate.Visible = True
318        frmDate.FontBold = False

319        frmSalesOrder.Top = 1920
320        frmSalesOrder.Height = 1335
321        fgSalesOrder.Height = 975
322        frmSalesOrder.Visible = True
323        frmSalesOrder.FontBold = False

324        frmSN.Top = 3240
325        frmSN.Height = 1335
326        fgSN.Height = 975
327        frmSN.Visible = True
328        frmSN.FontBold = False

329        frmModel.Top = 4560
330        frmModel.Height = 1335
331        fgModel.Height = 975
332        frmModel.Visible = True
333        frmModel.FontBold = False

334        frmTEMCFrameNo.Top = 5880
335        frmTEMCFrameNo.Height = 1335
336        fgTEMCFrameNo.Height = 975
337        frmTEMCFrameNo.Visible = True
338        frmTEMCFrameNo.FontBold = False

339        frmCustomer.Top = 7200
340        frmCustomer.Height = 1335
341        fgCustomer.Height = 975
342        frmCustomer.Visible = True
343        frmCustomer.FontBold = False

344        frmShipTo.Top = 8520
345        frmShipTo.Height = 1335
346        fgShipTo.Height = 1335
347        frmShipTo.Visible = True
348        frmShipTo.FontBold = False

349        frmWildCard.Top = 9960
350        frmWildCard.Height = 1335
351        fgWildCard.Height = 1335
352        frmWildCard.Visible = True
353        frmWildCard.FontBold = False
354        txtModelNumberString.Text = "Enter Characters and Search with Return"

' <VB WATCH>
355        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdResetSizes_Click"

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

Private Sub fgCustomer_Click()
' <VB WATCH>
356        On Error GoTo vbwErrHandler
' </VB WATCH>
357        fgCustomer.Col = 0
358        frmPLCData.txtSN.Text = fgCustomer.Text
' <VB WATCH>
359        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgCustomer_Click"

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

Private Sub fgDate_Click()
' <VB WATCH>
360        On Error GoTo vbwErrHandler
' </VB WATCH>
361        fgDate.Col = 0
362        frmPLCData.txtSN.Text = fgDate.Text
' <VB WATCH>
363        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgDate_Click"

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
Private Sub fgmodel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' <VB WATCH>
364        On Error GoTo vbwErrHandler
' </VB WATCH>

          'MSFlexGrid has the strange feature of not being able to recognize
          'when the heading of 1st fixed column is clicked on, it calls it Row 1,
          'the Row below this also returns as Row 1.
          'This bit of code below singles out the heading row which is required in
          'this app for sorting the data.

365        Static LastCol As Integer       'last column clicked on
366        Static direction As Boolean     'sorting ascending or descending

367        If y < fgModel.RowHeight(fgModel.Row) Then  'if user clicked in header row
368            fgModel.Row = 1
369            fgModel.RowSel = 1
370            fgModel.ColSel = fgModel.Col
371            If LastCol = fgModel.Col Then   'if user clicked on same column, reverse sort
372                direction = Not direction
373            Else
374                direction = True            'if new column, sort ascending
375            End If

376            If direction Then
377                fgModel.Sort = flexSortGenericAscending
378            Else
379                fgModel.Sort = flexSortGenericDescending
380            End If

381            LastCol = fgModel.Col   'save column number
382        Else                            'user did not click on header, select serial number for main screen
383            fgModel.Col = 0
384            frmPLCData.txtSN.Text = fgModel.Text
385        End If

' <VB WATCH>
386        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgmodel_MouseDown"

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

Private Sub fgSalesOrder_Click()
' <VB WATCH>
387        On Error GoTo vbwErrHandler
' </VB WATCH>
388        fgSalesOrder.Col = 0
389        frmPLCData.txtSN.Text = fgSalesOrder.Text
' <VB WATCH>
390        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgSalesOrder_Click"

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



Private Sub fgshipto_Click()
' <VB WATCH>
391        On Error GoTo vbwErrHandler
' </VB WATCH>
392        fgShipTo.Col = 0
393        frmPLCData.txtSN.Text = fgShipTo.Text
' <VB WATCH>
394        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgshipto_Click"

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

Private Sub fgSN_Click()
' <VB WATCH>
395        On Error GoTo vbwErrHandler
' </VB WATCH>
396        fgSN.Col = 0
397        frmPLCData.txtSN.Text = fgSN.Text
' <VB WATCH>
398        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgSN_Click"

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

Private Sub fgTEMCFrameNo_Click()
' <VB WATCH>
399        On Error GoTo vbwErrHandler
' </VB WATCH>
400        fgTEMCFrameNo.Col = 0
401        frmPLCData.txtSN.Text = fgTEMCFrameNo.Text
' <VB WATCH>
402        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgTEMCFrameNo_Click"

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
403        On Error GoTo vbwErrHandler
' </VB WATCH>
404        Const HWND_TOPMOST As Integer = -1
405        Const SWP_NOSIZE As Integer = &H1
406        Const SWP_NOMOVE As Integer = &H2
407        Const SWP_NOACTIVATE As Integer = &H10
408        Const SWP_SHOWWINDOW As Integer = &H40

           'window always on top
       '    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

' <VB WATCH>
409        Exit Sub
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

Private Sub Form_Load()
           'open several recordsets for searching
' <VB WATCH>
410        On Error GoTo vbwErrHandler
' </VB WATCH>

           'qydata/rsdata is for the serial number dropdown
411        qyData.ActiveConnection = cnPumpData
412        qyData.CommandText = "SELECT DISTINCT SerialNumber FROM TempPumpData ORDER BY SerialNumber;"
413        rsData.CursorType = adOpenStatic
414        rsData.CursorLocation = adUseClient
415        rsData.Index = "SerialNumber"
416        rsData.Open qyData

           'bind the serial number dropdown
417        Set cmbSearchSN.DataSource = rsData
418        cmbSearchSN.ListField = "SerialNumber"
419        Set cmbSearchSN.RowSource = rsData

           'qydata1 and 2/rsdata1 and 2 is for the date dropdown
420        qyData1.ActiveConnection = cnPumpData
421        rsData1.CursorType = adOpenStatic
422        rsData1.CursorLocation = adUseClient
423        rsData1.Index = "SerialNumber"

424        qyData2.ActiveConnection = cnPumpData
425        rsData2.CursorType = adOpenStatic
426        rsData2.CursorLocation = adUseClient
427        rsData2.Index = "SerialNumber"

           'find dates without times
       '    qyData1.CommandText = "SELECT DISTINCT TempPumpData.SerialNumber, TempPumpData.Model, TempTestSetupData.Date, IIf(InStr(2,[TempTestSetupData]![Date],"" "")<>0,Left$([TempTestSetupData]![Date],InStr(2,[TempTestSetupData]![Date],"" "")),[TempTestSetupData]![Date]) AS [Expr2]" & _
'        " FROM TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber ORDER BY Date;"
428        qyData1.CommandText = "SELECT DISTINCT TempTestSetupData.Date, IIf(InStr(2,[TempTestSetupData]![Date],"" "")<>0,Left$([TempTestSetupData]![Date],InStr(2,[TempTestSetupData]![Date],"" "")),[TempTestSetupData]![Date]) AS [Expr2] " & _
                             " FROM TempTestSetupData ORDER BY Date;"
429        rsData1.Open qyData1

430        Dim I As Integer
431        Dim TempDate As Date
432        Dim LastDate As Date
433        If Not rsData1.BOF Then
434            rsData1.MoveFirst
435        End If

436        LastDate = FormatDateTime(Now(), vbShortDate)
437        For I = 1 To rsData1.RecordCount
438            TempDate = FormatDateTime(rsData1.Fields(0), vbShortDate)
439            If TempDate <> LastDate Then
440                cmbStartDate.AddItem TempDate
441                LastDate = TempDate
442            End If
443            rsData1.MoveNext
444        Next I

           'qydatadate/rsdatadate for date datalist
445        qyDataDate.ActiveConnection = cnPumpData
446        rsDataDate.CursorType = adOpenStatic
447        rsDataDate.CursorLocation = adUseClient

           'qydatamodel/rsdatamodel for model dropdown
448        qyDataModel.ActiveConnection = cnPumpData
449        rsDataModel.CursorType = adOpenStatic
450        rsDataModel.CursorLocation = adUseClient

           'qydatasalesorder/rsdatasalesorder for sales order dropdown
451        qyDataSalesOrder.ActiveConnection = cnPumpData
452        rsDataSalesOrder.CursorType = adOpenStatic
453        rsDataSalesOrder.CursorLocation = adUseClient
454        qyDataSalesOrder.CommandText = "SELECT DISTINCT TempPumpData.SalesOrderNumber FROM TempPumpData ORDER BY TempPumpData.SalesOrderNumber;"
455        rsDataSalesOrder.Open qyDataSalesOrder

456        qySalesOrderData.ActiveConnection = cnPumpData
457        rsSalesOrderData.CursorType = adOpenStatic
458        rsSalesOrderData.CursorLocation = adUseClient

           'bind to temc frame number dropdown
459        Set cmbSearchSalesOrder.RowSource = rsDataSalesOrder
460        cmbSearchSalesOrder.ListField = "SalesOrderNumber"
461        Set cmbSearchSalesOrder.RowSource = rsDataSalesOrder

           'qydatasn/rsdatasn for serial numbers
462        qyDataSN.ActiveConnection = cnPumpData
463        rsDataSN.CursorType = adOpenStatic
464        rsDataSN.CursorLocation = adUseClient

           'qydatatemcmodel/rsdatatemcmodel for temc frame number
465        qyDataTEMCModel.ActiveConnection = cnPumpData
466        rsDataTEMCModel.CursorType = adOpenStatic
467        rsDataTEMCModel.CursorLocation = adUseClient
468        qyDataTEMCModel.CommandText = "SELECT DISTINCT TempPumpData.TEMCFrameNumber FROM TempPumpData WHERE (TempPumpData.ChempumpPump = FALSE) ORDER BY TempPumpData.TEMCFrameNumber;"
469        rsDataTEMCModel.Open qyDataTEMCModel

           'bind to temc frame number dropdown
470        Set cmbSearchTEMCFrameNumber.RowSource = rsDataTEMCModel
471        cmbSearchTEMCFrameNumber.ListField = "TEMCFrameNumber"
472        Set cmbSearchTEMCFrameNumber.RowSource = rsDataTEMCModel

473        qyDataTEMCFrameNumber.ActiveConnection = cnPumpData
474        rsDataTEMCFrameNumber.CursorType = adOpenStatic
475        rsDataTEMCFrameNumber.CursorLocation = adUseClient

           'customer
476        qyDataCustomer.ActiveConnection = cnPumpData
477        rsDataCustomer.CursorType = adOpenStatic
478        rsDataCustomer.CursorLocation = adUseClient

479        qyDataCustomer.CommandText = "SELECT DISTINCT TempPumpData.BillToCustomer FROM TempPumpData ORDER BY TempPumpData.BillToCustomer;"
480        rsDataCustomer.Open qyDataCustomer

481        qyCustomerData.ActiveConnection = cnPumpData
482        rsCustomerData.CursorType = adOpenStatic
483        rsCustomerData.CursorLocation = adUseClient

           'bind to customer dropdown
484        Set cmbSearchCustomer.RowSource = rsDataCustomer
485        cmbSearchCustomer.ListField = "BillToCustomer"

           ' ship to customer
486        qyDataShipto.ActiveConnection = cnPumpData
487        rsDataShipTo.CursorType = adOpenStatic
488        rsDataShipTo.CursorLocation = adUseClient

489        qyDataShipto.CommandText = "SELECT DISTINCT TempPumpData.shipToCustomer FROM TempPumpData ORDER BY TempPumpData.shipToCustomer;"
490        rsDataShipTo.Open qyDataShipto

491        qyShipToData.ActiveConnection = cnPumpData
492        rsShipToData.CursorType = adOpenStatic
493        rsShipToData.CursorLocation = adUseClient

           'bind to customer dropdown
494        Set cmbSearchShipTo.RowSource = rsDataShipTo
495        cmbSearchShipTo.ListField = "ShipToCustomer"

496        cmbSearchEndDate.Enabled = False

' <VB WATCH>
497        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Form_Load"

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

Private Sub Form_Unload(Cancel As Integer)
' <VB WATCH>
498        On Error GoTo vbwErrHandler
' </VB WATCH>

           'close all the datasets and release the connections
499        If rsData.State = adStateOpen Then
500            rsData.Close
501        End If
502        If rsData1.State = adStateOpen Then
503            rsData1.Close
504        End If
505        If rsData2.State = adStateOpen Then
506            rsData2.Close
507        End If
508        If rsDataDate.State = adStateOpen Then
509            rsDataDate.Close
510        End If
511        If rsDataModel.State = adStateOpen Then
512            rsDataModel.Close
513        End If
514        If rsDataSalesOrder.State = adStateOpen Then
515            rsDataSalesOrder.Close
516        End If
517        If rsSalesOrderData.State = adStateOpen Then
518            rsSalesOrderData.Close
519        End If
520        If rsDataSN.State = adStateOpen Then
521            rsDataSN.Close
522        End If
523        If rsDataTEMCModel.State = adStateOpen Then
524            rsDataTEMCModel.Close
525        End If
526        If rsDataTEMCFrameNumber.State = adStateOpen Then
527            rsDataTEMCFrameNumber.Close
528        End If
529        If rsDataCustomer.State = adStateOpen Then
530            rsDataCustomer.Close
531        End If
532        If rsCustomerData.State = adStateOpen Then
533            rsCustomerData.Close
534        End If
535        If rsDataShipTo.State = adStateOpen Then
536            rsDataShipTo.Close
537        End If
538        If rsShipToData.State = adStateOpen Then
539            rsShipToData.Close
540        End If

541        Set rsData = Nothing
542        Set rsData1 = Nothing
543        Set rsData2 = Nothing
544        Set rsDataDate = Nothing
545        Set rsDataModel = Nothing
546        Set rsDataSalesOrder = Nothing
547        Set rsSalesOrderData = Nothing
548        Set rsDataSN = Nothing
549        Set rsDataTEMCModel = Nothing
550        Set rsDataTEMCFrameNumber = Nothing
551        Set rsDataCustomer = Nothing
552        Set rsCustomerData = Nothing
553        Set rsDataShipTo = Nothing
554        Set rsShipToData = Nothing

' <VB WATCH>
555        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Form_Unload"

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

Private Sub WildCardSearch()
' <VB WATCH>
556        On Error GoTo vbwErrHandler
' </VB WATCH>
557        If rsDataModel.State = adStateOpen Then
558            rsDataModel.Close
559        End If

560        qyDataModel.CommandText = "SELECT DISTINCT " & _
                         " [TempPumpData]![SerialNumber], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], IIF(TempTestSetupData.ImpTrimmed=0, val(TempPumpData!ImpellerDia), val(TempTestSetupData!ImpTrimmed)) as ImpDia ,  " & _
                         "  TempPumpData.BillToCustomer,  Model.Description " & _
                         " FROM Motor INNER JOIN ((Model INNER JOIN TempPumpData ON Model.Model = TempPumpData.Model) INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) ON Motor.Motor = TempPumpData.Motor" & _
                         " WHERE (((TempPumpData.ModelNumber) LIKE '%" & txtModelNumberString.Text & "%'));"


561        rsDataModel.Open qyDataModel

562        If rsDataModel.RecordCount = 0 Then
563            Exit Sub
564        End If

565        Set fgWildCard.DataSource = rsDataModel

566        Dim f As String

567        f = "<S/N     |<Date      |<Sales Order |<Model No     |^Imp Dia  |<Bill To     |<Ship To  "
568        fgWildCard.FormatString = f
569        fgWildCard.ColAlignment(4) = flexAlignCenterTop
           'fgModel.ColAlignment(5) = flexAlignCenterTop
570        fgWildCard.ColWidth(0) = 1200
571        fgWildCard.ColWidth(1) = 2000
572        fgWildCard.ColWidth(2) = 1200
573        fgWildCard.ColWidth(3) = 2000
574        fgWildCard.ColWidth(4) = 1200
           'fgModel.ColWidth(5) = 1200
575        fgWildCard.ColWidth(5) = 3200
576        fgWildCard.ColWidth(6) = 3200
577        fgWildCard.TextMatrix(0, 0) = "S/N"
578        fgWildCard.TextMatrix(0, 1) = "Date"
579        fgWildCard.TextMatrix(0, 2) = "Sales Order"
580        fgWildCard.TextMatrix(0, 3) = "Model No"
581        fgWildCard.TextMatrix(0, 4) = "Imp Dia"
           'fgModel.TextMatrix(0, 5) = "Motor"
582        fgWildCard.TextMatrix(0, 5) = "Bill To"
583        fgWildCard.TextMatrix(0, 6) = "Ship To"

584        Dim x As Long
585        With fgWildCard
586            For x = .FixedRows To .Rows - 1
587            .TextMatrix(x, 4) = Format(.TextMatrix(x, 4), "#0.000")
588            Next x
589        End With


590        frmTEMCFrameNo.Visible = False
591        frmCustomer.Visible = False
592        frmWildCard.Top = 4560
593        frmWildCard.Height = 4000
594        fgWildCard.Height = 4000 - 360
595        frmWildCard.FontBold = True

' <VB WATCH>
596        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "WildCardSearch"

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

Private Sub txtModelNumberString_KeyPress(KeyAscii As Integer)
' <VB WATCH>
597        On Error GoTo vbwErrHandler
' </VB WATCH>
598        If KeyAscii = 13 Then
599            WildCardSearch
600        End If
' <VB WATCH>
601        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtModelNumberString_KeyPress"

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

Private Sub fgwildcard_Click()
' <VB WATCH>
602        On Error GoTo vbwErrHandler
' </VB WATCH>
603        fgWildCard.Col = 0
604        frmPLCData.txtSN.Text = fgWildCard.Text
' <VB WATCH>
605        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgwildcard_Click"

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

Private Sub txtModelNumberString_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' <VB WATCH>
606        On Error GoTo vbwErrHandler
' </VB WATCH>
607        If txtModelNumberString.Text = "Enter Characters and Search with Return" Then
608            txtModelNumberString.Text = ""
609        End If

' <VB WATCH>
610        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtModelNumberString_MouseDown"

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
Private Sub fgDate_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
           ' If this is not row 0, do nothing.
' <VB WATCH>
611        On Error GoTo vbwErrHandler
' </VB WATCH>
612        If fgDate.MouseRow <> 0 Then
613             Exit Sub
614        End If
615        Static direction As Boolean
616        With fgDate
617            .ColSel = fgDate.MouseCol

618            If direction = True Then
619                .Sort = flexSortStringAscending
620            Else
621                .Sort = flexSortStringDescending
622            End If
623            direction = Not direction
624        End With

' <VB WATCH>
625        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgDate_MouseUp"

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
Private Sub fgSalesOrder_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
           ' If this is not row 0, do nothing.
' <VB WATCH>
626        On Error GoTo vbwErrHandler
' </VB WATCH>
627        If fgSalesOrder.MouseRow <> 0 Then
628             Exit Sub
629        End If
630        Static direction As Boolean
631        With fgSalesOrder
632            .ColSel = fgSalesOrder.MouseCol
633            If direction = True Then
634                .Sort = flexSortStringAscending
635            Else
636                .Sort = flexSortStringDescending
637            End If
638            direction = Not direction
639        End With

' <VB WATCH>
640        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgSalesOrder_MouseUp"

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

Private Sub fgSN_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
           ' If this is not row 0, do nothing.
' <VB WATCH>
641        On Error GoTo vbwErrHandler
' </VB WATCH>
642        If fgSN.MouseRow <> 0 Then
643             Exit Sub
644        End If
645        Static direction As Boolean
646        With fgSN
647            .ColSel = fgSN.MouseCol
648            If direction = True Then
649                .Sort = flexSortStringAscending
650            Else
651                .Sort = flexSortStringDescending
652            End If
653            direction = Not direction
654        End With

' <VB WATCH>
655        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgSN_MouseUp"

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

Private Sub fgModel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
           ' If this is not row 0, do nothing.
' <VB WATCH>
656        On Error GoTo vbwErrHandler
' </VB WATCH>
657        If fgModel.MouseRow <> 0 Then
658             Exit Sub
659        End If
660        Static direction As Boolean
661        With fgModel
662            .ColSel = fgModel.MouseCol
663            If direction = True Then
664                .Sort = flexSortStringAscending
665            Else
666                .Sort = flexSortStringDescending
667            End If
668            direction = Not direction
669        End With

' <VB WATCH>
670        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgModel_MouseUp"

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

Private Sub fgTEMCFrameNo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
           ' If this is not row 0, do nothing.
' <VB WATCH>
671        On Error GoTo vbwErrHandler
' </VB WATCH>
672        If fgTEMCFrameNo.MouseRow <> 0 Then
673             Exit Sub
674        End If
675        Static direction As Boolean
676        Dim HeaderOrg As String

677        With fgTEMCFrameNo
678            .ColSel = fgTEMCFrameNo.MouseCol
       '        .TextMatrix(0, .ColSel) = Replace(.TextMatrix(0, .ColSel), " " & ChrW(&H25B2), "")
       '        .TextMatrix(0, .ColSel) = Replace(.TextMatrix(0, .ColSel), " " & ChrW(&H25BC), "")

679            If direction = True Then
680                .Sort = flexSortStringAscending
       '            .TextMatrix(0, .ColSel) = .TextMatrix(0, .ColSel) & " " & ChrW(&H25BC)
681            Else
682                .Sort = flexSortStringDescending
       '            .TextMatrix(0, .ColSel) = .TextMatrix(0, .ColSel) & " " & ChrW(&H25B2)
683            End If
684            direction = Not direction
685        End With

' <VB WATCH>
686        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgTEMCFrameNo_MouseUp"

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

Private Sub fgCustomer_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
           ' If this is not row 0, do nothing.
' <VB WATCH>
687        On Error GoTo vbwErrHandler
' </VB WATCH>
688        If fgCustomer.MouseRow <> 0 Then
689             Exit Sub
690        End If
691        Static direction As Boolean
692        With fgCustomer
693            .ColSel = fgCustomer.MouseCol
694            If direction = True Then
695                .Sort = flexSortStringAscending
696            Else
697                .Sort = flexSortStringDescending
698            End If
699            direction = Not direction
700        End With

' <VB WATCH>
701        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgCustomer_MouseUp"

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

Private Sub fgShipTo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
           ' If this is not row 0, do nothing.
' <VB WATCH>
702        On Error GoTo vbwErrHandler
' </VB WATCH>
703        If fgShipTo.MouseRow <> 0 Then
704             Exit Sub
705        End If
706        Static direction As Boolean
707        With fgShipTo
708            .ColSel = fgShipTo.MouseCol
709            If direction = True Then
710                .Sort = flexSortStringAscending
711            Else
712                .Sort = flexSortStringDescending
713            End If
714            direction = Not direction
715        End With

' <VB WATCH>
716        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgShipTo_MouseUp"

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

Private Sub fgWildCard_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
           ' If this is not row 0, do nothing.
' <VB WATCH>
717        On Error GoTo vbwErrHandler
' </VB WATCH>
718        If fgWildCard.MouseRow <> 0 Then
719             Exit Sub
720        End If
721        Static direction As Boolean
722        With fgWildCard
723            .ColSel = fgWildCard.MouseCol
724            If direction = True Then
725                .Sort = flexSortStringAscending
726            Else
727                .Sort = flexSortStringDescending
728            End If
729            direction = Not direction
730        End With

' <VB WATCH>
731        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "fgWildCard_MouseUp"

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


