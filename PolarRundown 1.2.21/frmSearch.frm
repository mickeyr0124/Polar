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
2          Const VBWPROCNAME = "frmSearch.cmbSalesOrder_Click"
3          If vbwProtector.vbwTraceProc Then
4              Dim vbwProtectorParameterString As String
5              If vbwProtector.vbwTraceParameters Then
6                  vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Area", Area) & ") "
7              End If
8              vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
9          End If
' </VB WATCH>

' <VB WATCH>
10         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
11         Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Area", Area
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmbSearchCustomer_Click(Area As Integer)
' <VB WATCH>
12         On Error GoTo vbwErrHandler
13         Const VBWPROCNAME = "frmSearch.cmbSearchCustomer_Click"
14         If vbwProtector.vbwTraceProc Then
15             Dim vbwProtectorParameterString As String
16             If vbwProtector.vbwTraceParameters Then
17                 vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Area", Area) & ") "
18             End If
19             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
20         End If
' </VB WATCH>
21         If rsCustomerData.State = adStateOpen Then
22             rsCustomerData.Close
23         End If

24         If cmbSearchCustomer.SelectedItem = 1 Then
' <VB WATCH>
25         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
26             Exit Sub
27         End If

28         qyCustomerData.CommandText = "SELECT DISTINCT " & _
                         " [TempPumpData]![SerialNumber], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], TempPumpData.ShiptoCustomer " & _
                         " FROM (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
                         " WHERE (((TempPumpData.BillToCustomer)= '" & cmbSearchCustomer.BoundText & "'));"

29         rsCustomerData.Open qyCustomerData

30         If rsCustomerData.RecordCount = 0 Then
' <VB WATCH>
31         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
32             Exit Sub
33         End If

           'bind the datalist to the recordset
34         Set fgCustomer.DataSource = rsCustomerData

35         fgCustomer.ColWidth(0) = 1400
36         fgCustomer.ColWidth(1) = 2000
37         fgCustomer.ColWidth(2) = 1200
38         fgCustomer.ColWidth(3) = 2000
39         fgCustomer.ColWidth(4) = 3200
40         fgCustomer.TextMatrix(0, 0) = "S/N"
41         fgCustomer.TextMatrix(0, 1) = "Date"
42         fgCustomer.TextMatrix(0, 2) = "Sales Order"
43         fgCustomer.TextMatrix(0, 3) = "Model No"
44         fgCustomer.TextMatrix(0, 4) = "Ship To"

45         frmModel.Visible = False
46         frmTEMCFrameNo.Visible = False
47         frmCustomer.Top = 7200 - (4000 - 1335)
48         frmCustomer.Height = 4000
49         fgCustomer.Height = 4000 - 360
50         frmCustomer.FontBold = True

' <VB WATCH>
51         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
52         Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Area", Area
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub


Private Sub cmbSearchEndDate_Click(Area As Integer)
' <VB WATCH>
53         On Error GoTo vbwErrHandler
54         Const VBWPROCNAME = "frmSearch.cmbSearchEndDate_Click"
55         If vbwProtector.vbwTraceProc Then
56             Dim vbwProtectorParameterString As String
57             If vbwProtector.vbwTraceParameters Then
58                 vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Area", Area) & ") "
59             End If
60             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
61         End If
' </VB WATCH>
62         cmbStartDate_Click
' <VB WATCH>
63         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
64         Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Area", Area
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmbSearchModel_Click()
' <VB WATCH>
65         On Error GoTo vbwErrHandler
66         Const VBWPROCNAME = "frmSearch.cmbSearchModel_Click"
67         If vbwProtector.vbwTraceProc Then
68             Dim vbwProtectorParameterString As String
69             If vbwProtector.vbwTraceParameters Then
70                 vbwProtectorParameterString = "()"
71             End If
72             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
73         End If
' </VB WATCH>
74         If rsDataModel.State = adStateOpen Then
75             rsDataModel.Close
76         End If

77         qyDataModel.CommandText = "SELECT DISTINCT " & _
                         " [TempPumpData]![SerialNumber], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], IIF(TempTestSetupData.ImpTrimmed=0, val(TempPumpData!ImpellerDia), val(TempTestSetupData!ImpTrimmed)) as ImpDia ,  " & _
                         "  TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer" & _
                         " FROM Motor INNER JOIN ((Model INNER JOIN TempPumpData ON Model.Model = TempPumpData.Model) INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) ON Motor.Motor = TempPumpData.Motor" & _
                         " WHERE (((TempPumpData.ModelNumber) LIKE '%" & cmbSearchModel.List(cmbSearchModel.ListIndex) & "%'));"
       '                  " WHERE (((TempPumpData.Model)= " & cmbSearchModel.ItemData(cmbSearchModel.ListIndex) & "));"

78         rsDataModel.Open qyDataModel

79         If rsDataModel.RecordCount = 0 Then
' <VB WATCH>
80         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
81             Exit Sub
82         End If

83         Set fgModel.DataSource = rsDataModel

84         Dim f As String

85         f = "<S/N     |<Date      |<Sales Order |<Model No     |^Imp Dia  |<Bill To     |<Ship To  "
86         fgModel.FormatString = f
87         fgModel.ColAlignment(4) = flexAlignCenterTop
           'fgModel.ColAlignment(5) = flexAlignCenterTop
88         fgModel.ColWidth(0) = 1200
89         fgModel.ColWidth(1) = 2000
90         fgModel.ColWidth(2) = 1200
91         fgModel.ColWidth(3) = 2000
92         fgModel.ColWidth(4) = 1200
           'fgModel.ColWidth(5) = 1200
93         fgModel.ColWidth(5) = 3200
94         fgModel.ColWidth(6) = 3200
95         fgModel.TextMatrix(0, 0) = "S/N"
96         fgModel.TextMatrix(0, 1) = "Date"
97         fgModel.TextMatrix(0, 2) = "Sales Order"
98         fgModel.TextMatrix(0, 3) = "Model No"
99         fgModel.TextMatrix(0, 4) = "Imp Dia"
           'fgModel.TextMatrix(0, 5) = "Motor"
100        fgModel.TextMatrix(0, 5) = "Bill To"
101        fgModel.TextMatrix(0, 6) = "Ship To"

102        Dim x As Long
103        With fgModel
104            For x = .FixedRows To .Rows - 1
105            .TextMatrix(x, 4) = Format(.TextMatrix(x, 4), "#0.000")
106            Next x
107        End With


108        frmTEMCFrameNo.Visible = False
109        frmCustomer.Visible = False
110        frmModel.Height = 4000
111        fgModel.Height = 4000 - 360
112        frmModel.FontBold = True

' <VB WATCH>
113        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
114        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "f", f
            vbwReportVariable "x", x
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmbSearchSalesOrder_Change()
' <VB WATCH>
115        On Error GoTo vbwErrHandler
116        Const VBWPROCNAME = "frmSearch.cmbSearchSalesOrder_Change"
117        If vbwProtector.vbwTraceProc Then
118            Dim vbwProtectorParameterString As String
119            If vbwProtector.vbwTraceParameters Then
120                vbwProtectorParameterString = "()"
121            End If
122            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
123        End If
' </VB WATCH>

124        Text1.Text = cmbSearchSalesOrder.BoundText

125        If rsSalesOrderData.State = adStateOpen Then
126            rsSalesOrderData.Close
127        End If

           'find all dates and models for the selected serial number
128        qySalesOrderData.CommandText = "SELECT DISTINCT " & _
                         " [TempPumpData]![SerialNumber], [TempPumpData]![ChempumpPump], [TempTestSetupData]![Date], [TempPumpData]![ModelNumber], [TempPumpData]![TEMCFrameNumber], TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer " & _
                         " FROM (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
                         " WHERE (((TempPumpData.SalesOrderNumber)= '" & cmbSearchSalesOrder.BoundText & "'));"

129        rsSalesOrderData.Open qySalesOrderData

130        If rsSalesOrderData.RecordCount = 0 Then
' <VB WATCH>
131        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
132            Exit Sub
133        End If

       '    'bind the datalist to the other recordset
134        Set fgSalesOrder.DataSource = rsSalesOrderData

135        fgSalesOrder.ColWidth(0) = 1400               'serial number
136        fgSalesOrder.ColWidth(1) = 0               'chempumppump
137        fgSalesOrder.ColWidth(2) = 2000
138        If rsSalesOrderData.Fields(1) = True Then       'show model number
139            fgSalesOrder.ColWidth(3) = 1800
140            fgSalesOrder.ColWidth(4) = 0
141        Else                                    'else, show TEMC Frame number
142            fgSalesOrder.ColWidth(3) = 0
143            fgSalesOrder.ColWidth(4) = 1800
144        End If
145        fgSalesOrder.ColWidth(5) = 3200
146        fgSalesOrder.ColWidth(6) = 3200
147        fgSalesOrder.TextMatrix(0, 0) = "S/N"
148        fgSalesOrder.TextMatrix(0, 2) = "Date"
149        fgSalesOrder.TextMatrix(0, 3) = "Model No"
150        fgSalesOrder.TextMatrix(0, 4) = "TEMC Frame"
151        fgSalesOrder.TextMatrix(0, 5) = "Bill To"
152        fgSalesOrder.TextMatrix(0, 6) = "Ship To"

153        frmSN.Visible = False
154        frmSalesOrder.Height = 4000
155        fgSalesOrder.Height = 4000 - 360
156        frmSalesOrder.FontBold = True

' <VB WATCH>
157        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
158        Exit Sub
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

Private Sub cmbSearchShipTo_Click(Area As Integer)
' <VB WATCH>
159        On Error GoTo vbwErrHandler
160        Const VBWPROCNAME = "frmSearch.cmbSearchShipTo_Click"
161        If vbwProtector.vbwTraceProc Then
162            Dim vbwProtectorParameterString As String
163            If vbwProtector.vbwTraceParameters Then
164                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Area", Area) & ") "
165            End If
166            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
167        End If
' </VB WATCH>

168        If rsShipToData.State = adStateOpen Then
169            rsShipToData.Close
170        End If

171        If cmbSearchShipTo.SelectedItem = 1 Then
' <VB WATCH>
172        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
173            Exit Sub
174        End If

175        qyShipToData.CommandText = "SELECT DISTINCT " & _
                         " [TempPumpData]![SerialNumber],[TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], TempPumpData.ModelNumber, TempPumpData.BillToCustomer " & _
                         " FROM  (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
                         " WHERE (((TempPumpData.ShipToCustomer)= '" & cmbSearchShipTo.BoundText & "'));"

176        rsShipToData.Open qyShipToData

177        If rsShipToData.RecordCount = 0 Then
' <VB WATCH>
178        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
179            Exit Sub
180        End If

181        Set fgShipTo.DataSource = rsShipToData


182        fgShipTo.ColWidth(0) = 1400
183        fgShipTo.ColWidth(1) = 2000
184        fgShipTo.ColWidth(2) = 1200
185        fgShipTo.ColWidth(3) = 2000
186        fgShipTo.ColWidth(4) = 3200
187        fgShipTo.TextMatrix(0, 0) = "S/N"
188        fgShipTo.TextMatrix(0, 1) = "Date"
189        fgShipTo.TextMatrix(0, 2) = "Sales Order"
190        fgShipTo.TextMatrix(0, 3) = "Model No"
191        fgShipTo.TextMatrix(0, 4) = "Bill To"

192        frmTEMCFrameNo.Visible = False
193        frmCustomer.Visible = False
194        frmShipTo.Top = 8520 - (4000 - 1335)
195        frmShipTo.Height = 4000
196        fgShipTo.Height = 4000 - 360
197        frmShipTo.FontBold = True

' <VB WATCH>
198        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
199        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Area", Area
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmbSearchSN_Change()
' <VB WATCH>
200        On Error GoTo vbwErrHandler
201        Const VBWPROCNAME = "frmSearch.cmbSearchSN_Change"
202        If vbwProtector.vbwTraceProc Then
203            Dim vbwProtectorParameterString As String
204            If vbwProtector.vbwTraceParameters Then
205                vbwProtectorParameterString = "()"
206            End If
207            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
208        End If
' </VB WATCH>

209        Text1.Text = cmbSearchSN.BoundText

210        If rsDataSN.State = adStateOpen Then
211            rsDataSN.Close
212        End If

           'find all dates and models for the selected serial number
213        qyDataSN.CommandText = "SELECT DISTINCT TempPumpData.SerialNumber," & _
                         " [TempPumpData]![ChempumpPump],[TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], [TempPumpData]![TEMCFrameNumber], TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer " & _
                         " FROM (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
                         " WHERE (((TempPumpData.SerialNumber)= '" & cmbSearchSN.BoundText & "'));"

214        rsDataSN.Open qyDataSN

           'if we didn't find any records, see if we have any serial numbers that are close
215        If rsDataSN.RecordCount = 0 Then
216            rsDataSN.Close
217            qyDataSN.CommandText = "SELECT DISTINCT TempPumpData.SerialNumber," & _
                             " [TempPumpData]![ChempumpPump],[TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], [TempPumpData]![TEMCFrameNumber], TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer " & _
                             " FROM (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
                             " WHERE (((TempPumpData.SerialNumber)= '" & cmbSearchSN.BoundText & "%'));"
218            rsDataSN.Open qyDataSN
219        End If

220        If rsDataSN.RecordCount = 0 Then
' <VB WATCH>
221        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
222            Exit Sub
223        End If

224        lblNoOfPumps.Caption = rsDataSN.RecordCount & " Pumps Found"

           'bind the datalist to the other recordset
225        Set fgSN.DataSource = rsDataSN

226        fgSN.ColWidth(0) = 0               'serial number
227        fgSN.ColWidth(1) = 0               'chempumppump
228        fgSN.ColWidth(2) = 2000
229        fgSN.ColWidth(3) = 1200
230        If rsDataSN.Fields(1) = True Then       'show model number
231            fgSN.ColWidth(4) = 1800
232            fgSN.ColWidth(5) = 0
233        Else                                    'else, show TEMC Frame number
234            fgSN.ColWidth(4) = 0
235            fgSN.ColWidth(5) = 1800
236        End If
237        fgSN.ColWidth(6) = 3000
238        fgSN.ColWidth(7) = 3000
239        fgSN.TextMatrix(0, 2) = "Date"
240        fgSN.TextMatrix(0, 3) = "Sales Order"
241        fgSN.TextMatrix(0, 4) = "Model No"
242        fgSN.TextMatrix(0, 5) = "TEMC Frame"
243        fgSN.TextMatrix(0, 6) = "Bill To"
244        fgSN.TextMatrix(0, 7) = "Ship To"

           'put the serial number into the find pump textbox
245        frmPLCData.txtSN.Text = rsDataSN.Fields("SerialNumber")

246        frmModel.Visible = False
247        frmSN.Height = 4000
248        fgSN.Height = 4000 - 360
249        frmSN.FontBold = True

' <VB WATCH>
250        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
251        Exit Sub
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

Private Sub ORIGINALcmbSearchTEMCFrameNumber_Click(Area As Integer)
' <VB WATCH>
252        On Error GoTo vbwErrHandler
253        Const VBWPROCNAME = "frmSearch.ORIGINALcmbSearchTEMCFrameNumber_Click"
254        If vbwProtector.vbwTraceProc Then
255            Dim vbwProtectorParameterString As String
256            If vbwProtector.vbwTraceParameters Then
257                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Area", Area) & ") "
258            End If
259            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
260        End If
' </VB WATCH>
261        If rsDataTEMCFrameNumber.State = adStateOpen Then
262            rsDataTEMCFrameNumber.Close
263        End If

           'find all pumps with the selected temc frame number
264        qyDataTEMCFrameNumber.CommandText = "SELECT DISTINCT " & _
                         " [TempPumpData]![SerialNumber], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer " & _
                         " FROM TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber" & _
                         " WHERE (((TempPumpData.TemcFrameNumber)= '" & cmbSearchTEMCFrameNumber.BoundText & "'));"

265        rsDataTEMCFrameNumber.Open qyDataTEMCFrameNumber

266        If rsDataTEMCFrameNumber.RecordCount = 0 Then
' <VB WATCH>
267        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
268            Exit Sub
269        End If

           'bind the datalist to the recordset
270        Set fgTEMCFrameNo.DataSource = rsDataTEMCFrameNumber

271        fgTEMCFrameNo.TextMatrix(0, 0) = "S/N"
272        fgTEMCFrameNo.TextMatrix(0, 1) = "Date"
273        fgTEMCFrameNo.TextMatrix(0, 2) = "Sales Order"
274        fgTEMCFrameNo.TextMatrix(0, 3) = "Bill To"
275        fgTEMCFrameNo.TextMatrix(0, 4) = "Ship To"
276        fgTEMCFrameNo.ColWidth(0) = 1400
277        fgTEMCFrameNo.ColWidth(1) = 2000
278        fgTEMCFrameNo.ColWidth(2) = 1200
279        fgTEMCFrameNo.ColWidth(3) = 3200
280        fgTEMCFrameNo.ColWidth(4) = 3200

281        frmCustomer.Visible = False
282        frmShipTo.Visible = False
283        frmTEMCFrameNo.Height = 4000
284        fgTEMCFrameNo.Height = 4000 - 360
285        frmTEMCFrameNo.FontBold = True

' <VB WATCH>
286        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
287        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Area", Area
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub cmbSearchTEMCFrameNumber_Click(Area As Integer)
' <VB WATCH>
288        On Error GoTo vbwErrHandler
289        Const VBWPROCNAME = "frmSearch.cmbSearchTEMCFrameNumber_Click"
290        If vbwProtector.vbwTraceProc Then
291            Dim vbwProtectorParameterString As String
292            If vbwProtector.vbwTraceParameters Then
293                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Area", Area) & ") "
294            End If
295            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
296        End If
' </VB WATCH>
297        If rsDataTEMCFrameNumber.State = adStateOpen Then
298            rsDataTEMCFrameNumber.Close
299        End If

           'find all pumps with the selected temc frame number
300        qyDataTEMCFrameNumber.CommandText = "SELECT DISTINCT " & _
                         " [TempPumpData]![SerialNumber], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], IIF(TempTestSetupData.ImpTrimmed=0, val(TempPumpData!ImpellerDia), val(TempTestSetupData!ImpTrimmed)) as ImpDia ,  " & _
                         "  TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer" & _
                         " FROM Motor INNER JOIN ((Model INNER JOIN TempPumpData ON Model.Model = TempPumpData.Model) INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) ON Motor.Motor = TempPumpData.Motor" & _
                         " WHERE (((TempPumpData.TemcFrameNumber)= '" & cmbSearchTEMCFrameNumber.BoundText & "'));"

301        rsDataTEMCFrameNumber.Open qyDataTEMCFrameNumber

302        If rsDataTEMCFrameNumber.RecordCount = 0 Then
' <VB WATCH>
303        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
304            Exit Sub
305        End If

306        Set fgTEMCFrameNo.DataSource = rsDataTEMCFrameNumber

307        Dim f As String

308        f = "<S/N     |<Date      |<Sales Order |<Model No     |^Imp Dia  |<Bill To     |<Ship To  "
309        fgTEMCFrameNo.FormatString = f
310        fgTEMCFrameNo.ColAlignment(4) = flexAlignCenterTop
           'fgModel.ColAlignment(5) = flexAlignCenterTop
311        fgTEMCFrameNo.ColWidth(0) = 1200
312        fgTEMCFrameNo.ColWidth(1) = 2000
313        fgTEMCFrameNo.ColWidth(2) = 1200
314        fgTEMCFrameNo.ColWidth(3) = 2000
315        fgTEMCFrameNo.ColWidth(4) = 1200
           'fgModel.ColWidth(5) = 1200
316        fgTEMCFrameNo.ColWidth(5) = 3200
317        fgTEMCFrameNo.ColWidth(6) = 3200
318        fgTEMCFrameNo.TextMatrix(0, 0) = "S/N"
319        fgTEMCFrameNo.TextMatrix(0, 1) = "Date"
320        fgTEMCFrameNo.TextMatrix(0, 2) = "Sales Order"
321        fgTEMCFrameNo.TextMatrix(0, 3) = "Model No"
322        fgTEMCFrameNo.TextMatrix(0, 4) = "Imp Dia"
           'fgModel.TextMatrix(0, 5) = "Motor"
323        fgTEMCFrameNo.TextMatrix(0, 5) = "Bill To"
324        fgTEMCFrameNo.TextMatrix(0, 6) = "Ship To"

325        Dim x As Long
326        With fgTEMCFrameNo
327            For x = .FixedRows To .Rows - 1
328            .TextMatrix(x, 4) = Format(.TextMatrix(x, 4), "#0.000")
329            Next x
330        End With


331        frmCustomer.Visible = False
332        frmShipTo.Visible = False
333        frmTEMCFrameNo.Height = 4000
334        fgTEMCFrameNo.Height = 4000 - 360
335        frmTEMCFrameNo.FontBold = True

' <VB WATCH>
336        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
337        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Area", Area
            vbwReportVariable "f", f
            vbwReportVariable "x", x
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmbStartDate_Click()
' <VB WATCH>
338        On Error GoTo vbwErrHandler
339        Const VBWPROCNAME = "frmSearch.cmbStartDate_Click"
340        If vbwProtector.vbwTraceProc Then
341            Dim vbwProtectorParameterString As String
342            If vbwProtector.vbwTraceParameters Then
343                vbwProtectorParameterString = "()"
344            End If
345            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
346        End If
' </VB WATCH>
347        Dim StartDate As Date
348        Dim EndDate As Date
349        Dim I As Integer

350        Text1.Text = cmbStartDate.List(cmbStartDate.ListIndex)
351        If rsDataDate.State = adStateOpen Then
352            rsDataDate.Close
353        End If

354        StartDate = FormatDateTime(Text1.Text)

           'see if there's and end date
355        If Left$(cmbSearchEndDate.BoundText, 1) = "S" Then
356            EndDate = StartDate
357        Else
358            EndDate = FormatDateTime(cmbSearchEndDate.BoundText)
359        End If


360        I = InStr(EndDate, " ")
361        If I <> 0 Then
362            EndDate = Left$(EndDate, I)
363        End If
364        StartDate = StartDate & " 00:00:00"
365        EndDate = EndDate & " 23:59:59"

           'look for all tests that were done on that date, regardless of the time
366        qyDataDate.CommandText = "SELECT " & _
                             " [TempPumpData]![SerialNumber], [TempPumpData]![ChempumpPump], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], [TempPumpData]![TEMCFrameNumber], TempPumpData.BillToCustomer, TempPumpData.ShiptoCustomer " & _
                             " FROM (TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) " & _
                             " WHERE ((TempTestSetupData.Date >= #" & StartDate & "#) AND (TempTestSetupData.Date <= #" & EndDate & "#)) " & _
                             " ORDER BY TempTestSetupData.Date;"
367        qyData2.CommandText = "SELECT DISTINCT TempTestSetupData.Date, IIf(InStr(2,[TempTestSetupData]![Date],"" "")<>0,Left$([TempTestSetupData]![Date],InStr(2,[TempTestSetupData]![Date],"" "")),[TempTestSetupData]![Date]) AS [Expr2]  " & _
                            " FROM TempTestSetupData " & _
                            " WHERE Date >= #" & StartDate & "#" & _
                            " ORDER BY Date;"
       '    qyData2.CommandText = "SELECT DISTINCT TempPumpData.SerialNumber, TempTestSetupData.Date, TempPumpData.SalesOrderNumber, TempPumpData.ModelNumber " & _
'       " FROM TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber " & _
'       " WHERE Date >= #" & StartDate & "#" & _
'       " ORDER BY Date;"

368        If rsData2.State = adStateOpen Then
369            rsData2.Close
370        End If

371        rsData2.Open qyData2
372        Set cmbSearchEndDate.DataSource = rsData2
373        cmbSearchEndDate.ListField = "Expr2"
374        Set cmbSearchEndDate.RowSource = rsData2

375        cmbSearchEndDate.Enabled = True

376        rsDataDate.Open qyDataDate

377        If rsDataDate.RecordCount = 0 Then
' <VB WATCH>
378        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
379            Exit Sub
380        End If

381        lblNoOfPumps.Caption = rsDataDate.RecordCount & " Pumps Found"

           'bind the dldate datalist to the recordset
382        Set fgDate.DataSource = rsDataDate

383        fgDate.TextMatrix(0, 0) = "S/N"
384        fgDate.TextMatrix(0, 2) = "Date"
385        fgDate.TextMatrix(0, 3) = "Sales Order"
386        fgDate.TextMatrix(0, 4) = "Model No"
387        fgDate.TextMatrix(0, 5) = "TEMC Frame"
388        fgDate.TextMatrix(0, 6) = "Bill To"
389        fgDate.TextMatrix(0, 7) = "Ship To"
390        fgDate.ColWidth(0) = 1400               'serial number
391        fgDate.ColWidth(1) = 0               'chempumppump
392        fgDate.ColWidth(2) = 2000
393        fgDate.ColWidth(3) = 1200
394        If rsDataDate.Fields(1) = True Then       'show model number
395            fgDate.ColWidth(4) = 1800
396            fgDate.ColWidth(5) = 0
397        Else                                    'else, show TEMC Frame number
398            fgDate.ColWidth(4) = 0
399            fgDate.ColWidth(5) = 1800
400        End If
401        fgDate.ColWidth(6) = 3000
402        fgDate.ColWidth(7) = 3000

403        frmSalesOrder.Visible = False
404        frmSN.Visible = False
405        frmDate.Height = 4000
406        fgDate.Height = 4000 - 360
407        frmDate.FontBold = True

' <VB WATCH>
408        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
409        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "StartDate", StartDate
            vbwReportVariable "EndDate", EndDate
            vbwReportVariable "I", I
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmdClose_Click()
' <VB WATCH>
410        On Error GoTo vbwErrHandler
411        Const VBWPROCNAME = "frmSearch.cmdClose_Click"
412        If vbwProtector.vbwTraceProc Then
413            Dim vbwProtectorParameterString As String
414            If vbwProtector.vbwTraceParameters Then
415                vbwProtectorParameterString = "()"
416            End If
417            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
418        End If
' </VB WATCH>
419        Unload Me   'unload the form
' <VB WATCH>
420        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
421        Exit Sub
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

Private Sub cmdResetSizes_Click()
' <VB WATCH>
422        On Error GoTo vbwErrHandler
423        Const VBWPROCNAME = "frmSearch.cmdResetSizes_Click"
424        If vbwProtector.vbwTraceProc Then
425            Dim vbwProtectorParameterString As String
426            If vbwProtector.vbwTraceParameters Then
427                vbwProtectorParameterString = "()"
428            End If
429            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
430        End If
' </VB WATCH>
431        frmDate.Top = 0
432        frmDate.Height = 1935
433        fgDate.Height = 1575
434        frmDate.Visible = True
435        frmDate.FontBold = False

436        frmSalesOrder.Top = 1920
437        frmSalesOrder.Height = 1335
438        fgSalesOrder.Height = 975
439        frmSalesOrder.Visible = True
440        frmSalesOrder.FontBold = False

441        frmSN.Top = 3240
442        frmSN.Height = 1335
443        fgSN.Height = 975
444        frmSN.Visible = True
445        frmSN.FontBold = False

446        frmModel.Top = 4560
447        frmModel.Height = 1335
448        fgModel.Height = 975
449        frmModel.Visible = True
450        frmModel.FontBold = False

451        frmTEMCFrameNo.Top = 5880
452        frmTEMCFrameNo.Height = 1335
453        fgTEMCFrameNo.Height = 975
454        frmTEMCFrameNo.Visible = True
455        frmTEMCFrameNo.FontBold = False

456        frmCustomer.Top = 7200
457        frmCustomer.Height = 1335
458        fgCustomer.Height = 975
459        frmCustomer.Visible = True
460        frmCustomer.FontBold = False

461        frmShipTo.Top = 8520
462        frmShipTo.Height = 1335
463        fgShipTo.Height = 1335
464        frmShipTo.Visible = True
465        frmShipTo.FontBold = False

466        frmWildCard.Top = 9960
467        frmWildCard.Height = 1335
468        fgWildCard.Height = 1335
469        frmWildCard.Visible = True
470        frmWildCard.FontBold = False
471        txtModelNumberString.Text = "Enter Characters and Search with Return"

' <VB WATCH>
472        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
473        Exit Sub
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

Private Sub fgCustomer_Click()
' <VB WATCH>
474        On Error GoTo vbwErrHandler
475        Const VBWPROCNAME = "frmSearch.fgCustomer_Click"
476        If vbwProtector.vbwTraceProc Then
477            Dim vbwProtectorParameterString As String
478            If vbwProtector.vbwTraceParameters Then
479                vbwProtectorParameterString = "()"
480            End If
481            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
482        End If
' </VB WATCH>
483        fgCustomer.Col = 0
484        frmPLCData.txtSN.Text = fgCustomer.Text
' <VB WATCH>
485        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
486        Exit Sub
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

Private Sub fgDate_Click()
' <VB WATCH>
487        On Error GoTo vbwErrHandler
488        Const VBWPROCNAME = "frmSearch.fgDate_Click"
489        If vbwProtector.vbwTraceProc Then
490            Dim vbwProtectorParameterString As String
491            If vbwProtector.vbwTraceParameters Then
492                vbwProtectorParameterString = "()"
493            End If
494            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
495        End If
' </VB WATCH>
496        fgDate.Col = 0
497        frmPLCData.txtSN.Text = fgDate.Text
' <VB WATCH>
498        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
499        Exit Sub
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
Private Sub fgmodel_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' <VB WATCH>
500        On Error GoTo vbwErrHandler
501        Const VBWPROCNAME = "frmSearch.fgmodel_MouseDown"
502        If vbwProtector.vbwTraceProc Then
503            Dim vbwProtectorParameterString As String
504            If vbwProtector.vbwTraceParameters Then
505                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Button", Button) & ", "
506                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("Shift", Shift) & ", "
507                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("x", x) & ", "
508                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("y", y) & ") "
509            End If
510            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
511        End If
' </VB WATCH>

          'MSFlexGrid has the strange feature of not being able to recognize
          'when the heading of 1st fixed column is clicked on, it calls it Row 1,
          'the Row below this also returns as Row 1.
          'This bit of code below singles out the heading row which is required in
          'this app for sorting the data.

512        Static LastCol As Integer       'last column clicked on
513        Static direction As Boolean     'sorting ascending or descending

514        If y < fgModel.RowHeight(fgModel.Row) Then  'if user clicked in header row
515            fgModel.Row = 1
516            fgModel.RowSel = 1
517            fgModel.ColSel = fgModel.Col
518            If LastCol = fgModel.Col Then   'if user clicked on same column, reverse sort
519                direction = Not direction
520            Else
521                direction = True            'if new column, sort ascending
522            End If

523            If direction Then
524                fgModel.Sort = flexSortGenericAscending
525            Else
526                fgModel.Sort = flexSortGenericDescending
527            End If

528            LastCol = fgModel.Col   'save column number
529        Else                            'user did not click on header, select serial number for main screen
530            fgModel.Col = 0
531            frmPLCData.txtSN.Text = fgModel.Text
532        End If

' <VB WATCH>
533        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
534        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Button", Button
            vbwReportVariable "Shift", Shift
            vbwReportVariable "x", x
            vbwReportVariable "y", y
            vbwReportVariable "LastCol", LastCol
            vbwReportVariable "direction", direction
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub fgSalesOrder_Click()
' <VB WATCH>
535        On Error GoTo vbwErrHandler
536        Const VBWPROCNAME = "frmSearch.fgSalesOrder_Click"
537        If vbwProtector.vbwTraceProc Then
538            Dim vbwProtectorParameterString As String
539            If vbwProtector.vbwTraceParameters Then
540                vbwProtectorParameterString = "()"
541            End If
542            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
543        End If
' </VB WATCH>
544        fgSalesOrder.Col = 0
545        frmPLCData.txtSN.Text = fgSalesOrder.Text
' <VB WATCH>
546        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
547        Exit Sub
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



Private Sub fgshipto_Click()
' <VB WATCH>
548        On Error GoTo vbwErrHandler
549        Const VBWPROCNAME = "frmSearch.fgshipto_Click"
550        If vbwProtector.vbwTraceProc Then
551            Dim vbwProtectorParameterString As String
552            If vbwProtector.vbwTraceParameters Then
553                vbwProtectorParameterString = "()"
554            End If
555            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
556        End If
' </VB WATCH>
557        fgShipTo.Col = 0
558        frmPLCData.txtSN.Text = fgShipTo.Text
' <VB WATCH>
559        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
560        Exit Sub
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

Private Sub fgSN_Click()
' <VB WATCH>
561        On Error GoTo vbwErrHandler
562        Const VBWPROCNAME = "frmSearch.fgSN_Click"
563        If vbwProtector.vbwTraceProc Then
564            Dim vbwProtectorParameterString As String
565            If vbwProtector.vbwTraceParameters Then
566                vbwProtectorParameterString = "()"
567            End If
568            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
569        End If
' </VB WATCH>
570        fgSN.Col = 0
571        frmPLCData.txtSN.Text = fgSN.Text
' <VB WATCH>
572        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
573        Exit Sub
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

Private Sub fgTEMCFrameNo_Click()
' <VB WATCH>
574        On Error GoTo vbwErrHandler
575        Const VBWPROCNAME = "frmSearch.fgTEMCFrameNo_Click"
576        If vbwProtector.vbwTraceProc Then
577            Dim vbwProtectorParameterString As String
578            If vbwProtector.vbwTraceParameters Then
579                vbwProtectorParameterString = "()"
580            End If
581            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
582        End If
' </VB WATCH>
583        fgTEMCFrameNo.Col = 0
584        frmPLCData.txtSN.Text = fgTEMCFrameNo.Text
' <VB WATCH>
585        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
586        Exit Sub
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
587        On Error GoTo vbwErrHandler
588        Const VBWPROCNAME = "frmSearch.Form_Activate"
589        If vbwProtector.vbwTraceProc Then
590            Dim vbwProtectorParameterString As String
591            If vbwProtector.vbwTraceParameters Then
592                vbwProtectorParameterString = "()"
593            End If
594            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
595        End If
' </VB WATCH>
596        Const HWND_TOPMOST As Integer = -1
597        Const SWP_NOSIZE As Integer = &H1
598        Const SWP_NOMOVE As Integer = &H2
599        Const SWP_NOACTIVATE As Integer = &H10
600        Const SWP_SHOWWINDOW As Integer = &H40

           'window always on top
       '    SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_SHOWWINDOW Or SWP_NOMOVE Or SWP_NOSIZE

' <VB WATCH>
601        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
602        Exit Sub
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

Private Sub Form_Load()
           'open several recordsets for searching
' <VB WATCH>
603        On Error GoTo vbwErrHandler
604        Const VBWPROCNAME = "frmSearch.Form_Load"
605        If vbwProtector.vbwTraceProc Then
606            Dim vbwProtectorParameterString As String
607            If vbwProtector.vbwTraceParameters Then
608                vbwProtectorParameterString = "()"
609            End If
610            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
611        End If
' </VB WATCH>

           'qydata/rsdata is for the serial number dropdown
612        qyData.ActiveConnection = cnPumpData
613        qyData.CommandText = "SELECT DISTINCT SerialNumber FROM TempPumpData ORDER BY SerialNumber;"
614        rsData.CursorType = adOpenStatic
615        rsData.CursorLocation = adUseClient
616        rsData.Index = "SerialNumber"
617        rsData.Open qyData

           'bind the serial number dropdown
618        Set cmbSearchSN.DataSource = rsData
619        cmbSearchSN.ListField = "SerialNumber"
620        Set cmbSearchSN.RowSource = rsData

           'qydata1 and 2/rsdata1 and 2 is for the date dropdown
621        qyData1.ActiveConnection = cnPumpData
622        rsData1.CursorType = adOpenStatic
623        rsData1.CursorLocation = adUseClient
624        rsData1.Index = "SerialNumber"

625        qyData2.ActiveConnection = cnPumpData
626        rsData2.CursorType = adOpenStatic
627        rsData2.CursorLocation = adUseClient
628        rsData2.Index = "SerialNumber"

           'find dates without times
       '    qyData1.CommandText = "SELECT DISTINCT TempPumpData.SerialNumber, TempPumpData.Model, TempTestSetupData.Date, IIf(InStr(2,[TempTestSetupData]![Date],"" "")<>0,Left$([TempTestSetupData]![Date],InStr(2,[TempTestSetupData]![Date],"" "")),[TempTestSetupData]![Date]) AS [Expr2]" & _
'        " FROM TempPumpData INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber ORDER BY Date;"
629        qyData1.CommandText = "SELECT DISTINCT TempTestSetupData.Date, IIf(InStr(2,[TempTestSetupData]![Date],"" "")<>0,Left$([TempTestSetupData]![Date],InStr(2,[TempTestSetupData]![Date],"" "")),[TempTestSetupData]![Date]) AS [Expr2] " & _
                             " FROM TempTestSetupData ORDER BY Date;"
630        rsData1.Open qyData1

631        Dim I As Integer
632        Dim TempDate As Date
633        Dim LastDate As Date
634        If Not rsData1.BOF Then
635            rsData1.MoveFirst
636        End If

637        LastDate = FormatDateTime(Now(), vbShortDate)
638        For I = 1 To rsData1.RecordCount
639            TempDate = FormatDateTime(rsData1.Fields(0), vbShortDate)
640            If TempDate <> LastDate Then
641                cmbStartDate.AddItem TempDate
642                LastDate = TempDate
643            End If
644            rsData1.MoveNext
645        Next I

           'qydatadate/rsdatadate for date datalist
646        qyDataDate.ActiveConnection = cnPumpData
647        rsDataDate.CursorType = adOpenStatic
648        rsDataDate.CursorLocation = adUseClient

           'qydatamodel/rsdatamodel for model dropdown
649        qyDataModel.ActiveConnection = cnPumpData
650        rsDataModel.CursorType = adOpenStatic
651        rsDataModel.CursorLocation = adUseClient

           'qydatasalesorder/rsdatasalesorder for sales order dropdown
652        qyDataSalesOrder.ActiveConnection = cnPumpData
653        rsDataSalesOrder.CursorType = adOpenStatic
654        rsDataSalesOrder.CursorLocation = adUseClient
655        qyDataSalesOrder.CommandText = "SELECT DISTINCT TempPumpData.SalesOrderNumber FROM TempPumpData ORDER BY TempPumpData.SalesOrderNumber;"
656        rsDataSalesOrder.Open qyDataSalesOrder

657        qySalesOrderData.ActiveConnection = cnPumpData
658        rsSalesOrderData.CursorType = adOpenStatic
659        rsSalesOrderData.CursorLocation = adUseClient

           'bind to temc frame number dropdown
660        Set cmbSearchSalesOrder.RowSource = rsDataSalesOrder
661        cmbSearchSalesOrder.ListField = "SalesOrderNumber"
662        Set cmbSearchSalesOrder.RowSource = rsDataSalesOrder

           'qydatasn/rsdatasn for serial numbers
663        qyDataSN.ActiveConnection = cnPumpData
664        rsDataSN.CursorType = adOpenStatic
665        rsDataSN.CursorLocation = adUseClient

           'qydatatemcmodel/rsdatatemcmodel for temc frame number
666        qyDataTEMCModel.ActiveConnection = cnPumpData
667        rsDataTEMCModel.CursorType = adOpenStatic
668        rsDataTEMCModel.CursorLocation = adUseClient
669        qyDataTEMCModel.CommandText = "SELECT DISTINCT TempPumpData.TEMCFrameNumber FROM TempPumpData WHERE (TempPumpData.ChempumpPump = FALSE) ORDER BY TempPumpData.TEMCFrameNumber;"
670        rsDataTEMCModel.Open qyDataTEMCModel

           'bind to temc frame number dropdown
671        Set cmbSearchTEMCFrameNumber.RowSource = rsDataTEMCModel
672        cmbSearchTEMCFrameNumber.ListField = "TEMCFrameNumber"
673        Set cmbSearchTEMCFrameNumber.RowSource = rsDataTEMCModel

674        qyDataTEMCFrameNumber.ActiveConnection = cnPumpData
675        rsDataTEMCFrameNumber.CursorType = adOpenStatic
676        rsDataTEMCFrameNumber.CursorLocation = adUseClient

           'customer
677        qyDataCustomer.ActiveConnection = cnPumpData
678        rsDataCustomer.CursorType = adOpenStatic
679        rsDataCustomer.CursorLocation = adUseClient

680        qyDataCustomer.CommandText = "SELECT DISTINCT TempPumpData.BillToCustomer FROM TempPumpData ORDER BY TempPumpData.BillToCustomer;"
681        rsDataCustomer.Open qyDataCustomer

682        qyCustomerData.ActiveConnection = cnPumpData
683        rsCustomerData.CursorType = adOpenStatic
684        rsCustomerData.CursorLocation = adUseClient

           'bind to customer dropdown
685        Set cmbSearchCustomer.RowSource = rsDataCustomer
686        cmbSearchCustomer.ListField = "BillToCustomer"

           ' ship to customer
687        qyDataShipto.ActiveConnection = cnPumpData
688        rsDataShipTo.CursorType = adOpenStatic
689        rsDataShipTo.CursorLocation = adUseClient

690        qyDataShipto.CommandText = "SELECT DISTINCT TempPumpData.shipToCustomer FROM TempPumpData ORDER BY TempPumpData.shipToCustomer;"
691        rsDataShipTo.Open qyDataShipto

692        qyShipToData.ActiveConnection = cnPumpData
693        rsShipToData.CursorType = adOpenStatic
694        rsShipToData.CursorLocation = adUseClient

           'bind to customer dropdown
695        Set cmbSearchShipTo.RowSource = rsDataShipTo
696        cmbSearchShipTo.ListField = "ShipToCustomer"

697        cmbSearchEndDate.Enabled = False

' <VB WATCH>
698        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
699        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "I", I
            vbwReportVariable "TempDate", TempDate
            vbwReportVariable "LastDate", LastDate
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub Form_Unload(Cancel As Integer)
' <VB WATCH>
700        On Error GoTo vbwErrHandler
701        Const VBWPROCNAME = "frmSearch.Form_Unload"
702        If vbwProtector.vbwTraceProc Then
703            Dim vbwProtectorParameterString As String
704            If vbwProtector.vbwTraceParameters Then
705                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Cancel", Cancel) & ") "
706            End If
707            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
708        End If
' </VB WATCH>

           'close all the datasets and release the connections
709        If rsData.State = adStateOpen Then
710            rsData.Close
711        End If
712        If rsData1.State = adStateOpen Then
713            rsData1.Close
714        End If
715        If rsData2.State = adStateOpen Then
716            rsData2.Close
717        End If
718        If rsDataDate.State = adStateOpen Then
719            rsDataDate.Close
720        End If
721        If rsDataModel.State = adStateOpen Then
722            rsDataModel.Close
723        End If
724        If rsDataSalesOrder.State = adStateOpen Then
725            rsDataSalesOrder.Close
726        End If
727        If rsSalesOrderData.State = adStateOpen Then
728            rsSalesOrderData.Close
729        End If
730        If rsDataSN.State = adStateOpen Then
731            rsDataSN.Close
732        End If
733        If rsDataTEMCModel.State = adStateOpen Then
734            rsDataTEMCModel.Close
735        End If
736        If rsDataTEMCFrameNumber.State = adStateOpen Then
737            rsDataTEMCFrameNumber.Close
738        End If
739        If rsDataCustomer.State = adStateOpen Then
740            rsDataCustomer.Close
741        End If
742        If rsCustomerData.State = adStateOpen Then
743            rsCustomerData.Close
744        End If
745        If rsDataShipTo.State = adStateOpen Then
746            rsDataShipTo.Close
747        End If
748        If rsShipToData.State = adStateOpen Then
749            rsShipToData.Close
750        End If

751        Set rsData = Nothing
752        Set rsData1 = Nothing
753        Set rsData2 = Nothing
754        Set rsDataDate = Nothing
755        Set rsDataModel = Nothing
756        Set rsDataSalesOrder = Nothing
757        Set rsSalesOrderData = Nothing
758        Set rsDataSN = Nothing
759        Set rsDataTEMCModel = Nothing
760        Set rsDataTEMCFrameNumber = Nothing
761        Set rsDataCustomer = Nothing
762        Set rsCustomerData = Nothing
763        Set rsDataShipTo = Nothing
764        Set rsShipToData = Nothing

' <VB WATCH>
765        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
766        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Cancel", Cancel
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub WildCardSearch()
' <VB WATCH>
767        On Error GoTo vbwErrHandler
768        Const VBWPROCNAME = "frmSearch.WildCardSearch"
769        If vbwProtector.vbwTraceProc Then
770            Dim vbwProtectorParameterString As String
771            If vbwProtector.vbwTraceParameters Then
772                vbwProtectorParameterString = "()"
773            End If
774            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
775        End If
' </VB WATCH>
776        If rsDataModel.State = adStateOpen Then
777            rsDataModel.Close
778        End If

779        qyDataModel.CommandText = "SELECT DISTINCT " & _
                         " [TempPumpData]![SerialNumber], [TempTestSetupData]![Date], [TempPumpData]![SalesOrderNumber], [TempPumpData]![ModelNumber], IIF(TempTestSetupData.ImpTrimmed=0, val(TempPumpData!ImpellerDia), val(TempTestSetupData!ImpTrimmed)) as ImpDia ,  " & _
                         "  TempPumpData.BillToCustomer,  Model.Description " & _
                         " FROM Motor INNER JOIN ((Model INNER JOIN TempPumpData ON Model.Model = TempPumpData.Model) INNER JOIN TempTestSetupData ON TempPumpData.SerialNumber = TempTestSetupData.SerialNumber) ON Motor.Motor = TempPumpData.Motor" & _
                         " WHERE (((TempPumpData.ModelNumber) LIKE '%" & txtModelNumberString.Text & "%'));"


780        rsDataModel.Open qyDataModel

781        If rsDataModel.RecordCount = 0 Then
' <VB WATCH>
782        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
783            Exit Sub
784        End If

785        Set fgWildCard.DataSource = rsDataModel

786        Dim f As String

787        f = "<S/N     |<Date      |<Sales Order |<Model No     |^Imp Dia  |<Bill To     |<Ship To  "
788        fgWildCard.FormatString = f
789        fgWildCard.ColAlignment(4) = flexAlignCenterTop
           'fgModel.ColAlignment(5) = flexAlignCenterTop
790        fgWildCard.ColWidth(0) = 1200
791        fgWildCard.ColWidth(1) = 2000
792        fgWildCard.ColWidth(2) = 1200
793        fgWildCard.ColWidth(3) = 2000
794        fgWildCard.ColWidth(4) = 1200
           'fgModel.ColWidth(5) = 1200
795        fgWildCard.ColWidth(5) = 3200
796        fgWildCard.ColWidth(6) = 3200
797        fgWildCard.TextMatrix(0, 0) = "S/N"
798        fgWildCard.TextMatrix(0, 1) = "Date"
799        fgWildCard.TextMatrix(0, 2) = "Sales Order"
800        fgWildCard.TextMatrix(0, 3) = "Model No"
801        fgWildCard.TextMatrix(0, 4) = "Imp Dia"
           'fgModel.TextMatrix(0, 5) = "Motor"
802        fgWildCard.TextMatrix(0, 5) = "Bill To"
803        fgWildCard.TextMatrix(0, 6) = "Ship To"

804        Dim x As Long
805        With fgWildCard
806            For x = .FixedRows To .Rows - 1
807            .TextMatrix(x, 4) = Format(.TextMatrix(x, 4), "#0.000")
808            Next x
809        End With


810        frmTEMCFrameNo.Visible = False
811        frmCustomer.Visible = False
812        frmWildCard.Top = 4560
813        frmWildCard.Height = 4000
814        fgWildCard.Height = 4000 - 360
815        frmWildCard.FontBold = True

' <VB WATCH>
816        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
817        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "f", f
            vbwReportVariable "x", x
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub txtModelNumberString_KeyPress(KeyAscii As Integer)
' <VB WATCH>
818        On Error GoTo vbwErrHandler
819        Const VBWPROCNAME = "frmSearch.txtModelNumberString_KeyPress"
820        If vbwProtector.vbwTraceProc Then
821            Dim vbwProtectorParameterString As String
822            If vbwProtector.vbwTraceParameters Then
823                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("KeyAscii", KeyAscii) & ") "
824            End If
825            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
826        End If
' </VB WATCH>
827        If KeyAscii = 13 Then
828            WildCardSearch
829        End If
' <VB WATCH>
830        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
831        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "KeyAscii", KeyAscii
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub fgwildcard_Click()
' <VB WATCH>
832        On Error GoTo vbwErrHandler
833        Const VBWPROCNAME = "frmSearch.fgwildcard_Click"
834        If vbwProtector.vbwTraceProc Then
835            Dim vbwProtectorParameterString As String
836            If vbwProtector.vbwTraceParameters Then
837                vbwProtectorParameterString = "()"
838            End If
839            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
840        End If
' </VB WATCH>
841        fgWildCard.Col = 0
842        frmPLCData.txtSN.Text = fgWildCard.Text
' <VB WATCH>
843        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
844        Exit Sub
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

Private Sub txtModelNumberString_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
' <VB WATCH>
845        On Error GoTo vbwErrHandler
846        Const VBWPROCNAME = "frmSearch.txtModelNumberString_MouseDown"
847        If vbwProtector.vbwTraceProc Then
848            Dim vbwProtectorParameterString As String
849            If vbwProtector.vbwTraceParameters Then
850                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Button", Button) & ", "
851                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("Shift", Shift) & ", "
852                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("x", x) & ", "
853                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("y", y) & ") "
854            End If
855            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
856        End If
' </VB WATCH>
857        If txtModelNumberString.Text = "Enter Characters and Search with Return" Then
858            txtModelNumberString.Text = ""
859        End If

' <VB WATCH>
860        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
861        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Button", Button
            vbwReportVariable "Shift", Shift
            vbwReportVariable "x", x
            vbwReportVariable "y", y
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub fgDate_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
           ' If this is not row 0, do nothing.
' <VB WATCH>
862        On Error GoTo vbwErrHandler
863        Const VBWPROCNAME = "frmSearch.fgDate_MouseUp"
864        If vbwProtector.vbwTraceProc Then
865            Dim vbwProtectorParameterString As String
866            If vbwProtector.vbwTraceParameters Then
867                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Button", Button) & ", "
868                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("Shift", Shift) & ", "
869                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("x", x) & ", "
870                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("y", y) & ") "
871            End If
872            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
873        End If
' </VB WATCH>
874        If fgDate.MouseRow <> 0 Then
' <VB WATCH>
875        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
876             Exit Sub
877        End If
878        Static direction As Boolean
879        With fgDate
880            .ColSel = fgDate.MouseCol

881            If direction = True Then
882                .Sort = flexSortStringAscending
883            Else
884                .Sort = flexSortStringDescending
885            End If
886            direction = Not direction
887        End With

' <VB WATCH>
888        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
889        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Button", Button
            vbwReportVariable "Shift", Shift
            vbwReportVariable "x", x
            vbwReportVariable "y", y
            vbwReportVariable "direction", direction
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub fgSalesOrder_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
           ' If this is not row 0, do nothing.
' <VB WATCH>
890        On Error GoTo vbwErrHandler
891        Const VBWPROCNAME = "frmSearch.fgSalesOrder_MouseUp"
892        If vbwProtector.vbwTraceProc Then
893            Dim vbwProtectorParameterString As String
894            If vbwProtector.vbwTraceParameters Then
895                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Button", Button) & ", "
896                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("Shift", Shift) & ", "
897                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("x", x) & ", "
898                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("y", y) & ") "
899            End If
900            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
901        End If
' </VB WATCH>
902        If fgSalesOrder.MouseRow <> 0 Then
' <VB WATCH>
903        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
904             Exit Sub
905        End If
906        Static direction As Boolean
907        With fgSalesOrder
908            .ColSel = fgSalesOrder.MouseCol
909            If direction = True Then
910                .Sort = flexSortStringAscending
911            Else
912                .Sort = flexSortStringDescending
913            End If
914            direction = Not direction
915        End With

' <VB WATCH>
916        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
917        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Button", Button
            vbwReportVariable "Shift", Shift
            vbwReportVariable "x", x
            vbwReportVariable "y", y
            vbwReportVariable "direction", direction
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub fgSN_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
           ' If this is not row 0, do nothing.
' <VB WATCH>
918        On Error GoTo vbwErrHandler
919        Const VBWPROCNAME = "frmSearch.fgSN_MouseUp"
920        If vbwProtector.vbwTraceProc Then
921            Dim vbwProtectorParameterString As String
922            If vbwProtector.vbwTraceParameters Then
923                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Button", Button) & ", "
924                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("Shift", Shift) & ", "
925                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("x", x) & ", "
926                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("y", y) & ") "
927            End If
928            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
929        End If
' </VB WATCH>
930        If fgSN.MouseRow <> 0 Then
' <VB WATCH>
931        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
932             Exit Sub
933        End If
934        Static direction As Boolean
935        With fgSN
936            .ColSel = fgSN.MouseCol
937            If direction = True Then
938                .Sort = flexSortStringAscending
939            Else
940                .Sort = flexSortStringDescending
941            End If
942            direction = Not direction
943        End With

' <VB WATCH>
944        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
945        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Button", Button
            vbwReportVariable "Shift", Shift
            vbwReportVariable "x", x
            vbwReportVariable "y", y
            vbwReportVariable "direction", direction
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub fgModel_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
           ' If this is not row 0, do nothing.
' <VB WATCH>
946        On Error GoTo vbwErrHandler
947        Const VBWPROCNAME = "frmSearch.fgModel_MouseUp"
948        If vbwProtector.vbwTraceProc Then
949            Dim vbwProtectorParameterString As String
950            If vbwProtector.vbwTraceParameters Then
951                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Button", Button) & ", "
952                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("Shift", Shift) & ", "
953                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("x", x) & ", "
954                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("y", y) & ") "
955            End If
956            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
957        End If
' </VB WATCH>
958        If fgModel.MouseRow <> 0 Then
' <VB WATCH>
959        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
960             Exit Sub
961        End If
962        Static direction As Boolean
963        With fgModel
964            .ColSel = fgModel.MouseCol
965            If direction = True Then
966                .Sort = flexSortStringAscending
967            Else
968                .Sort = flexSortStringDescending
969            End If
970            direction = Not direction
971        End With

' <VB WATCH>
972        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
973        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Button", Button
            vbwReportVariable "Shift", Shift
            vbwReportVariable "x", x
            vbwReportVariable "y", y
            vbwReportVariable "direction", direction
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub fgTEMCFrameNo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
           ' If this is not row 0, do nothing.
' <VB WATCH>
974        On Error GoTo vbwErrHandler
975        Const VBWPROCNAME = "frmSearch.fgTEMCFrameNo_MouseUp"
976        If vbwProtector.vbwTraceProc Then
977            Dim vbwProtectorParameterString As String
978            If vbwProtector.vbwTraceParameters Then
979                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Button", Button) & ", "
980                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("Shift", Shift) & ", "
981                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("x", x) & ", "
982                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("y", y) & ") "
983            End If
984            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
985        End If
' </VB WATCH>
986        If fgTEMCFrameNo.MouseRow <> 0 Then
' <VB WATCH>
987        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
988             Exit Sub
989        End If
990        Static direction As Boolean
991        Dim HeaderOrg As String

992        With fgTEMCFrameNo
993            .ColSel = fgTEMCFrameNo.MouseCol
       '        .TextMatrix(0, .ColSel) = Replace(.TextMatrix(0, .ColSel), " " & ChrW(&H25B2), "")
       '        .TextMatrix(0, .ColSel) = Replace(.TextMatrix(0, .ColSel), " " & ChrW(&H25BC), "")

994            If direction = True Then
995                .Sort = flexSortStringAscending
       '            .TextMatrix(0, .ColSel) = .TextMatrix(0, .ColSel) & " " & ChrW(&H25BC)
996            Else
997                .Sort = flexSortStringDescending
       '            .TextMatrix(0, .ColSel) = .TextMatrix(0, .ColSel) & " " & ChrW(&H25B2)
998            End If
999            direction = Not direction
1000       End With

' <VB WATCH>
1001       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1002       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Button", Button
            vbwReportVariable "Shift", Shift
            vbwReportVariable "x", x
            vbwReportVariable "y", y
            vbwReportVariable "direction", direction
            vbwReportVariable "HeaderOrg", HeaderOrg
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub fgCustomer_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
           ' If this is not row 0, do nothing.
' <VB WATCH>
1003       On Error GoTo vbwErrHandler
1004       Const VBWPROCNAME = "frmSearch.fgCustomer_MouseUp"
1005       If vbwProtector.vbwTraceProc Then
1006           Dim vbwProtectorParameterString As String
1007           If vbwProtector.vbwTraceParameters Then
1008               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Button", Button) & ", "
1009               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("Shift", Shift) & ", "
1010               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("x", x) & ", "
1011               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("y", y) & ") "
1012           End If
1013           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1014       End If
' </VB WATCH>
1015       If fgCustomer.MouseRow <> 0 Then
' <VB WATCH>
1016       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1017            Exit Sub
1018       End If
1019       Static direction As Boolean
1020       With fgCustomer
1021           .ColSel = fgCustomer.MouseCol
1022           If direction = True Then
1023               .Sort = flexSortStringAscending
1024           Else
1025               .Sort = flexSortStringDescending
1026           End If
1027           direction = Not direction
1028       End With

' <VB WATCH>
1029       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1030       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Button", Button
            vbwReportVariable "Shift", Shift
            vbwReportVariable "x", x
            vbwReportVariable "y", y
            vbwReportVariable "direction", direction
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub fgShipTo_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
           ' If this is not row 0, do nothing.
' <VB WATCH>
1031       On Error GoTo vbwErrHandler
1032       Const VBWPROCNAME = "frmSearch.fgShipTo_MouseUp"
1033       If vbwProtector.vbwTraceProc Then
1034           Dim vbwProtectorParameterString As String
1035           If vbwProtector.vbwTraceParameters Then
1036               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Button", Button) & ", "
1037               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("Shift", Shift) & ", "
1038               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("x", x) & ", "
1039               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("y", y) & ") "
1040           End If
1041           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1042       End If
' </VB WATCH>
1043       If fgShipTo.MouseRow <> 0 Then
' <VB WATCH>
1044       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1045            Exit Sub
1046       End If
1047       Static direction As Boolean
1048       With fgShipTo
1049           .ColSel = fgShipTo.MouseCol
1050           If direction = True Then
1051               .Sort = flexSortStringAscending
1052           Else
1053               .Sort = flexSortStringDescending
1054           End If
1055           direction = Not direction
1056       End With

' <VB WATCH>
1057       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1058       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Button", Button
            vbwReportVariable "Shift", Shift
            vbwReportVariable "x", x
            vbwReportVariable "y", y
            vbwReportVariable "direction", direction
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub fgWildCard_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
           ' If this is not row 0, do nothing.
' <VB WATCH>
1059       On Error GoTo vbwErrHandler
1060       Const VBWPROCNAME = "frmSearch.fgWildCard_MouseUp"
1061       If vbwProtector.vbwTraceProc Then
1062           Dim vbwProtectorParameterString As String
1063           If vbwProtector.vbwTraceParameters Then
1064               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Button", Button) & ", "
1065               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("Shift", Shift) & ", "
1066               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("x", x) & ", "
1067               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("y", y) & ") "
1068           End If
1069           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1070       End If
' </VB WATCH>
1071       If fgWildCard.MouseRow <> 0 Then
' <VB WATCH>
1072       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1073            Exit Sub
1074       End If
1075       Static direction As Boolean
1076       With fgWildCard
1077           .ColSel = fgWildCard.MouseCol
1078           If direction = True Then
1079               .Sort = flexSortStringAscending
1080           Else
1081               .Sort = flexSortStringDescending
1082           End If
1083           direction = Not direction
1084       End With

' <VB WATCH>
1085       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1086       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Button", Button
            vbwReportVariable "Shift", Shift
            vbwReportVariable "x", x
            vbwReportVariable "y", y
            vbwReportVariable "direction", direction
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
    vbwReportVariable "rsData", rsData
    vbwReportVariable "qyData", qyData
    vbwReportVariable "rsData1", rsData1
    vbwReportVariable "qyData1", qyData1
    vbwReportVariable "qyData2", qyData2
    vbwReportVariable "rsData2", rsData2
    vbwReportVariable "rsDataDate", rsDataDate
    vbwReportVariable "qyDataDate", qyDataDate
    vbwReportVariable "rsDataModel", rsDataModel
    vbwReportVariable "qyDataModel", qyDataModel
    vbwReportVariable "rsDataSN", rsDataSN
    vbwReportVariable "qyDataSN", qyDataSN
    vbwReportVariable "rsDataSalesOrder", rsDataSalesOrder
    vbwReportVariable "qySalesOrderData", qySalesOrderData
    vbwReportVariable "rsSalesOrderData", rsSalesOrderData
    vbwReportVariable "qyDataSalesOrder", qyDataSalesOrder
    vbwReportVariable "rsDataTEMCModel", rsDataTEMCModel
    vbwReportVariable "qyDataTEMCModel", qyDataTEMCModel
    vbwReportVariable "rsDataTEMCFrameNumber", rsDataTEMCFrameNumber
    vbwReportVariable "qyDataTEMCFrameNumber", qyDataTEMCFrameNumber
    vbwReportVariable "rsDataCustomer", rsDataCustomer
    vbwReportVariable "qyDataCustomer", qyDataCustomer
    vbwReportVariable "rsCustomerData", rsCustomerData
    vbwReportVariable "qyCustomerData", qyCustomerData
    vbwReportVariable "rsDataShipTo", rsDataShipTo
    vbwReportVariable "qyDataShipto", qyDataShipto
    vbwReportVariable "rsShipToData", rsShipToData
    vbwReportVariable "qyShipToData", qyShipToData
End Sub
' </VB WATCH>
