VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form frmCalibrate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Software Calibration"
   ClientHeight    =   3732
   ClientLeft      =   48
   ClientTop       =   336
   ClientWidth     =   7140
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3732
   ScaleWidth      =   7140
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdRunCalibration 
      Caption         =   "Run Calibration"
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   3120
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit Calibration"
      Height          =   495
      Left            =   5760
      TabIndex        =   1
      Top             =   3120
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   1095
      Left            =   840
      TabIndex        =   0
      Top             =   720
      Width           =   5535
      _ExtentX        =   9758
      _ExtentY        =   1926
      _Version        =   393216
      Rows            =   4
      Cols            =   5
      AllowBigSelection=   0   'False
      FocusRect       =   0
      HighLight       =   0
      ScrollBars      =   0
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      Caption         =   "This Hydraulic Rundown program will automatically close after the calibration is performed."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.2
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   1095
      Left            =   1320
      TabIndex        =   4
      Top             =   1920
      Width           =   4695
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Calibration Data Set Input Values"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   840
      TabIndex        =   2
      Top             =   360
      Width           =   5415
   End
End
Attribute VB_Name = "frmCalibrate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' <VB WATCH>
Const VBWMODULE = "frmCalibrate"
' </VB WATCH>

Private Sub cmdExit_Click()
' <VB WATCH>
1          On Error GoTo vbwErrHandler
2          Const VBWPROCNAME = "frmCalibrate.cmdExit_Click"
3          If vbwProtector.vbwTraceProc Then
4              Dim vbwProtectorParameterString As String
5              If vbwProtector.vbwTraceParameters Then
6                  vbwProtectorParameterString = "()"
7              End If
8              vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
9          End If
' </VB WATCH>
10         Dim I As Integer

11         If rsCalibrate.State = adStateOpen Then
12             rsCalibrate.Close
13         End If
14         If cnCalibrate.State = adStateOpen Then
15             cnCalibrate.Close
16         End If

17         Unload Me
18         Calibrating = False
19         End

' <VB WATCH>
20         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
21         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdExit_Click"

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
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmdRunCalibration_Click()
' <VB WATCH>
22         On Error GoTo vbwErrHandler
23         Const VBWPROCNAME = "frmCalibrate.cmdRunCalibration_Click"
24         If vbwProtector.vbwTraceProc Then
25             Dim vbwProtectorParameterString As String
26             If vbwProtector.vbwTraceParameters Then
27                 vbwProtectorParameterString = "()"
28             End If
29             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
30         End If
' </VB WATCH>
31         Dim X As Integer

32         cmdRunCalibration.Visible = False

           ' Create the Excel App Object so we can store our data
33         Set xlApp = CreateObject("Excel.Application")

34         OpenCalibrateFile

35         If Not WritingToCalFile Then
' <VB WATCH>
36         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
37             Exit Sub
38         End If

39         WriteCalHeader

40         For X = 0 To 2
41             UseDataset = DataSets(X)
42             With MSFlexGrid1
43                 .Row = X + 1
44                 .RowSel = X + 1
45                 .Col = 0
46                 .ColSel = .Cols - 1
47                 .Highlight = flexHighlightAlways
48             End With
49             Calibrating = True

50             DoCalibrationCalcs
51             WriteCalData (X)
52         Next X

53         MSFlexGrid1.Highlight = flexHighlightNever
54         xlApp.ActiveWorkbook.Save             'save the file

55         xlApp.Application.Quit
56         Set xlApp = Nothing

57         cmdExit_Click
' <VB WATCH>
58         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
59         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdRunCalibration_Click"

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
            vbwReportVariable "X", X
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
' <VB WATCH>
60         On Error GoTo vbwErrHandler
61         Const VBWPROCNAME = "frmCalibrate.Form_Load"
62         If vbwProtector.vbwTraceProc Then
63             Dim vbwProtectorParameterString As String
64             If vbwProtector.vbwTraceParameters Then
65                 vbwProtectorParameterString = "()"
66             End If
67             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
68         End If
' </VB WATCH>

69         Dim X As Long
70         Dim Count As Long

71         sCalibrateDatabaseName = App.Path & "\CalibrateData.mdb"
72         With cnCalibrate
73             .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sCalibrateDatabaseName & ";Persist Security Info=False"
74             .Open
75         End With
76         rsCalibrate.Open "Data", cnCalibrate, adOpenStatic, adLockOptimistic, adCmdTable

77         With MSFlexGrid1

78             .Redraw = False
79             .Clear
80             .Row = 0

81             .Col = 0
82             .ColWidth(0) = 750
83             .Text = "Data Set"

84             .Col = 1
85             .ColWidth(1) = 1200
86             .Text = "Flow"
87             .ColAlignment(1) = flexAlignCenterCenter

88             .Col = 2
89             .ColWidth(2) = 1200
90             .Text = "Disch Press"
91             .ColAlignment(2) = flexAlignCenterCenter

92             .Col = 3
93             .ColWidth(3) = 1200
94             .Text = "Suction Press"
95             .ColAlignment(3) = flexAlignCenterCenter

96             .Col = 4
97             .ColWidth(4) = 1200
98             .Text = "Temperature"
99             .ColAlignment(4) = flexAlignCenterCenter

               'setup the minimum number of rows & add column headers
100            .Rows = 2
101            .FixedRows = 1
102            .Row = 0
103            For X = 2 To 5
104                .Col = X - 2 + 1
105                .Text = rsCalibrate.Fields(X).Name
106                .ColData(X - 2 + 1) = rsCalibrate.Fields(X).Type
107            Next

108            .Rows = rsCalibrate.RecordCount + 1
109            For Count = 1 To rsCalibrate.RecordCount

110                .TextMatrix(Count, 0) = Count    'assign line number
111                For X = 0 To 3
                       'we use Variant conversion to avoid any possible NULL errors
112                    .TextMatrix(Count, X + 1) = "" & CVar(rsCalibrate.Fields(X + 2).value)
113                Next
114                rsCalibrate.MoveNext
115            Next

116            .Redraw = True
117        End With

118        rsCalibrate.MoveFirst

119        For X = 0 To 2
120            DataSets(X).Flow = rsCalibrate.Fields("Flow")
121            DataSets(X).SuctionPressure = rsCalibrate.Fields("SuctPress")
122            DataSets(X).DischargePressure = rsCalibrate.Fields("DischPress")
123            DataSets(X).Temperature = rsCalibrate.Fields("temp")
124            DataSets(X).SuctionPipeDia = rsCalibrate.Fields("SuctPipeDia")
125            DataSets(X).DischargePipeDia = rsCalibrate.Fields("DischPipeDia")
126            DataSets(X).SuctionHeight = rsCalibrate.Fields("SuctHeight")
127            DataSets(X).DischargeHeight = rsCalibrate.Fields("DischHeight")
128            DataSets(X).BarometricPressure = rsCalibrate.Fields("BaroPress")
129            DataSets(X).HDCorr = rsCalibrate.Fields("HDCorr")
130            DataSets(X).SuctionInHg = rsCalibrate.Fields("SuctionInHg")
131            DataSets(X).MotorType = rsCalibrate.Fields("MotorType")
132            DataSets(X).StatorFill = rsCalibrate.Fields("StatorFill")
133            DataSets(X).VoltageA = rsCalibrate.Fields("VoltageA")
134            DataSets(X).VoltageB = rsCalibrate.Fields("VoltageB")
135            DataSets(X).VoltageC = rsCalibrate.Fields("VoltageC")
136            DataSets(X).CurrentA = rsCalibrate.Fields("CurrentA")
137            DataSets(X).CurrentB = rsCalibrate.Fields("CurrentB")
138            DataSets(X).CurrentC = rsCalibrate.Fields("CurrentC")
139            DataSets(X).PowerA = rsCalibrate.Fields("PowerA")
140            DataSets(X).PowerB = rsCalibrate.Fields("PowerB")
141            DataSets(X).PowerC = rsCalibrate.Fields("PowerC")
142            DataSets(X).PowerFactor = rsCalibrate.Fields("PowerFactor")
143            DataSets(X).VelocityHead = rsCalibrate.Fields("VelocityHead")
144            DataSets(X).TDH = rsCalibrate.Fields("TDH")
145            DataSets(X).OverallEfficiency = rsCalibrate.Fields("OverallEfficiency")
146            DataSets(X).MotorEfficiency = rsCalibrate.Fields("MotorEfficiency")
147            DataSets(X).HydraulicEfficiency = rsCalibrate.Fields("HydraulicEfficiency")
148            rsCalibrate.MoveNext
149        Next X

150        rsCalibrate.Close

' <VB WATCH>
151        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
152        Exit Sub
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
            vbwReportVariable "X", X
            vbwReportVariable "Count", Count
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub OpenCalibrateFile()
' <VB WATCH>
153        On Error GoTo vbwErrHandler
154        Const VBWPROCNAME = "frmCalibrate.OpenCalibrateFile"
155        If vbwProtector.vbwTraceProc Then
156            Dim vbwProtectorParameterString As String
157            If vbwProtector.vbwTraceParameters Then
158                vbwProtectorParameterString = "()"
159            End If
160            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
161        End If
' </VB WATCH>
162            frmPLCData.CommonDialog1.CancelError = True        'in case the user
163            On Error GoTo ErrHandler                '  chooses the cancel button

               'set up dialog box
164            frmPLCData.CommonDialog1.DialogTitle = "Open Excel Calibration Files"
165            frmPLCData.CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|"  'show Excel files
166            frmPLCData.CommonDialog1.InitDir = sServerName & sCalibrateDirectoryName & "\Software Calibration"    'in this directory
167            frmPLCData.CommonDialog1.ShowOpen                              'open the file selection dialog box

168            If Dir(frmPLCData.CommonDialog1.filename) = "" Then            'if the file name does not exist yet
169                sCalibrateSaveFileName = frmPLCData.CommonDialog1.filename           'get the name of the file
170                If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
171                     xlApp.Workbooks.Close
172                End If
                   ' Create the Excel Workbook Object.
173    On Error GoTo vbwErrHandler
174                Set xlBook = xlApp.Workbooks.Add                'add a workbook
175                NewWorkBook                                     'do some stuff for the new workbook
176                xlApp.ActiveWorkbook.SaveAs filename:=sCalibrateSaveFileName, _
                                     FileFormat:=xlNormal                        'save the file
177                MsgBox frmPLCData.CommonDialog1.filename & " has been opened for writing.", vbOKOnly, "File Opened"    'tell the user that file is open
178            Else                                                'the file name already exists
179                sCalibrateSaveFileName = frmPLCData.CommonDialog1.filename
                   ' Create the Excel Workbook Object.
180                If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
181                     xlApp.Workbooks.Close
182                End If
183                Set xlBook = xlApp.Workbooks.Open(sCalibrateSaveFileName)             'get the file name selected
184                If GetWorksheetTabs = vbNo Then     'ask the user if he/she wants a new tab.
185                    MsgBox "File not overwritten.", vbOKOnly, "File not Opened"
' <VB WATCH>
186        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
187                    Exit Sub
188                Else
189                    MsgBox frmPLCData.CommonDialog1.filename & " has been opened for writing.", vbOKOnly, "File Opened"
190                End If
191            End If

192    On Error GoTo vbwErrHandler

193        WritingToCalFile = True

' <VB WATCH>
194        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
195        Exit Sub

196    ErrHandler:
           'User pressed the Cancel button

197        WritingToCalFile = False

' <VB WATCH>
198        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
199        Exit Sub

' <VB WATCH>
200        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
201        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "OpenCalibrateFile"

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
Public Sub WriteCalHeader()
' <VB WATCH>
202        On Error GoTo vbwErrHandler
203        Const VBWPROCNAME = "frmCalibrate.WriteCalHeader"
204        If vbwProtector.vbwTraceProc Then
205            Dim vbwProtectorParameterString As String
206            If vbwProtector.vbwTraceParameters Then
207                vbwProtectorParameterString = "()"
208            End If
209            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
210        End If
' </VB WATCH>
211        Dim TextToWrite As String
212        Dim RowNo As Integer

               'write the header to the file
213        With xlApp
214            .Range("B1").Select
215            .ActiveCell.FormulaR1C1 = "Hydraulic Rundown Calibration"
216            .Selection.HorizontalAlignment = xlCenter

217            .Range("A3").Select
218            .ActiveCell.FormulaR1C1 = "Date - "

219            .Range("B3").Select
220            .ActiveCell.FormulaR1C1 = Now

221             .Range("A4").Select
222            .ActiveCell.FormulaR1C1 = "Data Set"

223            .Range("C4:E4").Select
224            .Selection.Merge
225            .ActiveCell.FormulaR1C1 = "1"

226            .Range("C5").Select
227            .ActiveCell.FormulaR1C1 = "Input"
228            .Range("D5").Select
229            .ActiveCell.FormulaR1C1 = "Correct"
230            .Range("E5").Select
231            .ActiveCell.FormulaR1C1 = "Calculated"

232            .Range("F4:H4").Select
233            .Selection.Merge
234            .ActiveCell.FormulaR1C1 = "2"

235            .Range("F5").Select
236            .ActiveCell.FormulaR1C1 = "Input"
237            .Range("G5").Select
238            .ActiveCell.FormulaR1C1 = "Correct"
239            .Range("H5").Select
240            .ActiveCell.FormulaR1C1 = "Calculated"

241            .Range("I4:K4").Select
242            .Selection.Merge
243            .ActiveCell.FormulaR1C1 = "3"

244            .Range("I5").Select
245            .ActiveCell.FormulaR1C1 = "Input"
246            .Range("J5").Select
247            .ActiveCell.FormulaR1C1 = "Correct"
248            .Range("K5").Select
249            .ActiveCell.FormulaR1C1 = "Calculated"

250            .Range("C4:K5").Select
251            .Selection.HorizontalAlignment = xlCenter

252            .Range("A6").Select
253            .ActiveCell.FormulaR1C1 = "Inputs"
254            .Selection.Font.Bold = True

255            .Range("A7").Select
256            .ActiveCell.FormulaR1C1 = "Flow"

257            .Range("A8").Select
258            .ActiveCell.FormulaR1C1 = "Suction Pressure"

259             .Range("A9").Select
260            .ActiveCell.FormulaR1C1 = "Discharge Pressure"

261            .Range("A10").Select
262            .ActiveCell.FormulaR1C1 = "Temperature"

263            .Range("A11").Select
264            .ActiveCell.FormulaR1C1 = "Suction Pipe Dia"

265            .Range("A12").Select
266            .ActiveCell.FormulaR1C1 = "Discharge Pipe Dia"

267            .Range("A13").Select
268            .ActiveCell.FormulaR1C1 = "Suction Gauge Height"

269            .Range("A14").Select
270            .ActiveCell.FormulaR1C1 = "Discharge Gauge Height"

271            .Range("A15").Select
272            .ActiveCell.FormulaR1C1 = "Barometric Pressure"

273            .Range("A16").Select
274            .ActiveCell.FormulaR1C1 = "HDCorr"

275            .Range("A17").Select
276            .ActiveCell.FormulaR1C1 = "Suction (InHg)"

277            .Range("A18").Select
278            .ActiveCell.FormulaR1C1 = "Motor Type"

279            .Range("A19").Select
280            .ActiveCell.FormulaR1C1 = "Voltage A"

281            .Range("A20").Select
282            .ActiveCell.FormulaR1C1 = "Voltage B"

283            .Range("A21").Select
284            .ActiveCell.FormulaR1C1 = "Voltage C"

285            .Range("A22").Select
286            .ActiveCell.FormulaR1C1 = "Current A"

287            .Range("A23").Select
288            .ActiveCell.FormulaR1C1 = "Current B"

289            .Range("A24").Select
290            .ActiveCell.FormulaR1C1 = "Current C"

291            .Range("A25").Select
292            .ActiveCell.FormulaR1C1 = "Power A"

293            .Range("A26").Select
294            .ActiveCell.FormulaR1C1 = "Power B"

295            .Range("A27").Select
296            .ActiveCell.FormulaR1C1 = "Power C"

297            .Range("A28").Select
298            .ActiveCell.FormulaR1C1 = "Stator Fill"

299            .Range("A30").Select
300            .ActiveCell.FormulaR1C1 = "Calculated Values"
301            .Selection.Font.Bold = True

302            .Range("A31").Select
303            .ActiveCell.FormulaR1C1 = "Velocity Head"

304            .Range("A32").Select
305            .ActiveCell.FormulaR1C1 = "TDH"

306            .Range("A33").Select
307            .ActiveCell.FormulaR1C1 = "Overall Eff"

308            .Range("A34").Select
309            .ActiveCell.FormulaR1C1 = "Motor Eff"

310            .Range("A35").Select
311            .ActiveCell.FormulaR1C1 = "Hydraulic Eff"

312            .Range("A36").Select
313            .ActiveCell.FormulaR1C1 = "Power Factor"


314            .Range("D30").Select
315            .ActiveCell.FormulaR1C1 = "Correct"

316            .Range("E30").Select
317            .ActiveCell.FormulaR1C1 = "Calculated"

318            .Range("G30").Select
319            .ActiveCell.FormulaR1C1 = "Correct"

320            .Range("H30").Select
321            .ActiveCell.FormulaR1C1 = "Calculated"

322            .Range("J30").Select
323            .ActiveCell.FormulaR1C1 = "Correct"

324            .Range("K30").Select
325            .ActiveCell.FormulaR1C1 = "Calculated"

326            .Range("C7:K36").Select
327            .Selection.NumberFormat = "0.00"

328            Range("D30:E36").Select
329            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
330            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
331            With Selection.Borders(xlEdgeLeft)
332                .LineStyle = xlContinuous
333                .Weight = xlThin
334                .ColorIndex = xlAutomatic
335            End With
336            With Selection.Borders(xlEdgeTop)
337                .LineStyle = xlContinuous
338                .Weight = xlThin
339                .ColorIndex = xlAutomatic
340            End With
341            With Selection.Borders(xlEdgeBottom)
342                .LineStyle = xlContinuous
343                .Weight = xlThin
344                .ColorIndex = xlAutomatic
345            End With
346            With Selection.Borders(xlEdgeRight)
347                .LineStyle = xlContinuous
348                .Weight = xlThin
349                .ColorIndex = xlAutomatic
350            End With
351            With Selection.Borders(xlInsideVertical)
352                .LineStyle = xlContinuous
353                .Weight = xlThin
354                .ColorIndex = xlAutomatic
355            End With
356            With Selection.Borders(xlInsideHorizontal)
357                .LineStyle = xlContinuous
358                .Weight = xlThin
359                .ColorIndex = xlAutomatic
360            End With

361            Range("G30:H36").Select
362            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
363            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
364            With Selection.Borders(xlEdgeLeft)
365                .LineStyle = xlContinuous
366                .Weight = xlThin
367                .ColorIndex = xlAutomatic
368            End With
369            With Selection.Borders(xlEdgeTop)
370                .LineStyle = xlContinuous
371                .Weight = xlThin
372                .ColorIndex = xlAutomatic
373            End With
374            With Selection.Borders(xlEdgeBottom)
375                .LineStyle = xlContinuous
376                .Weight = xlThin
377                .ColorIndex = xlAutomatic
378            End With
379            With Selection.Borders(xlEdgeRight)
380                .LineStyle = xlContinuous
381                .Weight = xlThin
382                .ColorIndex = xlAutomatic
383            End With
384            With Selection.Borders(xlInsideVertical)
385                .LineStyle = xlContinuous
386                .Weight = xlThin
387                .ColorIndex = xlAutomatic
388            End With
389            With Selection.Borders(xlInsideHorizontal)
390                .LineStyle = xlContinuous
391                .Weight = xlThin
392                .ColorIndex = xlAutomatic
393            End With

394            Range("J30:K36").Select
395            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
396            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
397            With Selection.Borders(xlEdgeLeft)
398                .LineStyle = xlContinuous
399                .Weight = xlThin
400                .ColorIndex = xlAutomatic
401            End With
402            With Selection.Borders(xlEdgeTop)
403                .LineStyle = xlContinuous
404                .Weight = xlThin
405                .ColorIndex = xlAutomatic
406            End With
407            With Selection.Borders(xlEdgeBottom)
408                .LineStyle = xlContinuous
409                .Weight = xlThin
410                .ColorIndex = xlAutomatic
411            End With
412            With Selection.Borders(xlEdgeRight)
413                .LineStyle = xlContinuous
414                .Weight = xlThin
415                .ColorIndex = xlAutomatic
416            End With
417            With Selection.Borders(xlInsideVertical)
418                .LineStyle = xlContinuous
419                .Weight = xlThin
420                .ColorIndex = xlAutomatic
421            End With
422            With Selection.Borders(xlInsideHorizontal)
423                .LineStyle = xlContinuous
424                .Weight = xlThin
425                .ColorIndex = xlAutomatic
426            End With

427            .Range("B35").Select
428            .ActiveCell.FormulaR1C1 = "For formulas see:"
429            .Selection.Font.Bold = True

430            .Range("B36").Select
431            ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
                             sServerName & "EN\GROUPS\SHARED\Calibration and Rundown\Hydraulic Rundown Calibration\Software Calibration\Calibration Reference Sheet.xls" _
                             , TextToDisplay:="Calibration Reference Sheet"

432            With ActiveSheet.PageSetup
433                .Orientation = xlLandscape
434            End With

435        End With
' <VB WATCH>
436        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
437        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "WriteCalHeader"

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
            vbwReportVariable "TextToWrite", TextToWrite
            vbwReportVariable "RowNo", RowNo
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Public Sub WriteCalData(DatasetNumber As Integer)
' <VB WATCH>
438        On Error GoTo vbwErrHandler
439        Const VBWPROCNAME = "frmCalibrate.WriteCalData"
440        If vbwProtector.vbwTraceProc Then
441            Dim vbwProtectorParameterString As String
442            If vbwProtector.vbwTraceParameters Then
443                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("DatasetNumber", DatasetNumber) & ") "
444            End If
445            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
446        End If
' </VB WATCH>
447        Dim Col As String
448        Dim Row As Integer
449        Dim cell As String

450        Select Case DatasetNumber
               Case 0
451                Col = "C"
452            Case 1
453                Col = "F"
454            Case 2
455                Col = "I"
456            Case Else
457        End Select

458        With xlApp
459            For Row = 7 To 28
460                cell = Col & Trim(str(Row))
461                .Range(cell).Select
462                Select Case Row
                       Case Is = 7
463                        .ActiveCell.FormulaR1C1 = UseDataset.Flow
464                    Case Is = 8
465                        .ActiveCell.FormulaR1C1 = UseDataset.SuctionPressure
466                    Case Is = 9
467                        .ActiveCell.FormulaR1C1 = UseDataset.DischargePressure
468                    Case Is = 10
469                        .ActiveCell.FormulaR1C1 = UseDataset.Temperature
470                    Case Is = 11
471                        .ActiveCell.FormulaR1C1 = frmPLCData.cmbSuctDia.List(UseDataset.SuctionPipeDia - 1)
472                    Case Is = 12
473                        .ActiveCell.FormulaR1C1 = frmPLCData.cmbDischDia.List(UseDataset.DischargePipeDia - 1)
474                    Case Is = 13
475                        .ActiveCell.FormulaR1C1 = UseDataset.SuctionHeight
476                    Case Is = 14
477                        .ActiveCell.FormulaR1C1 = UseDataset.DischargeHeight
478                    Case Is = 15
479                        .ActiveCell.FormulaR1C1 = UseDataset.BarometricPressure
480                    Case Is = 16
481                        .ActiveCell.FormulaR1C1 = UseDataset.HDCorr
482                    Case Is = 17
483                        .ActiveCell.FormulaR1C1 = UseDataset.SuctionInHg
484                    Case Is = 18
485                    Dim I As Integer
486                For I = 0 To frmPLCData.cmbMotor.ListCount - 1
487                If frmPLCData.cmbMotor.ItemData(I) = UseDataset.MotorType Then
488                    .ActiveCell.FormulaR1C1 = frmPLCData.cmbMotor.List(I)
489                    Exit For
490                End If
491            Next I

       '                    .ActiveCell.FormulaR1C1 = frmPLCData.cmbMotor.ItemData(UseDataset.MotorType)
492                    Case Is = 19
493                        .ActiveCell.FormulaR1C1 = UseDataset.VoltageA
494                    Case Is = 20
495                        .ActiveCell.FormulaR1C1 = UseDataset.VoltageB
496                    Case Is = 21
497                        .ActiveCell.FormulaR1C1 = UseDataset.VoltageC
498                    Case Is = 22
499                        .ActiveCell.FormulaR1C1 = UseDataset.CurrentA
500                    Case Is = 23
501                        .ActiveCell.FormulaR1C1 = UseDataset.CurrentB
502                    Case Is = 24
503                        .ActiveCell.FormulaR1C1 = UseDataset.CurrentC
504                    Case Is = 25
505                        .ActiveCell.FormulaR1C1 = UseDataset.PowerA
506                    Case Is = 26
507                        .ActiveCell.FormulaR1C1 = UseDataset.PowerB
508                    Case Is = 27
509                        .ActiveCell.FormulaR1C1 = UseDataset.PowerC
510                    Case Is = 28
511                        If UseDataset.StatorFill = 1 Then
512                            .ActiveCell.FormulaR1C1 = "No"
513                        Else
514                            .ActiveCell.FormulaR1C1 = "Yes"
515                        End If
       '                    .ActiveCell.FormulaR1C1 = frmPLCData.cmbStatorFill.List(UseDataset.StatorFill)
516                End Select
517            Next Row

518            Col = Chr(Asc(Col) + 1)
519            For Row = 31 To 36
520                cell = Col & Trim(str(Row))
521                .Range(cell).Select
522                Select Case Row
                       Case Is = 31
523                        .ActiveCell.FormulaR1C1 = UseDataset.VelocityHead
524                    Case Is = 32
525                       .ActiveCell.FormulaR1C1 = UseDataset.TDH
526                    Case Is = 33
527                        .ActiveCell.FormulaR1C1 = UseDataset.OverallEfficiency
528                    Case Is = 34
529                        .ActiveCell.FormulaR1C1 = UseDataset.MotorEfficiency
530                    Case Is = 35
531                        .ActiveCell.FormulaR1C1 = UseDataset.HydraulicEfficiency
532                    Case Is = 36
533                        .ActiveCell.FormulaR1C1 = UseDataset.PowerFactor
534                End Select
535            Next Row

536            Col = Chr(Asc(Col) + 1)
537            For Row = 31 To 36
538                cell = Col & Trim(str(Row))
539                .Range(cell).Select
540                Select Case Row
                       Case Is = 31
541                        .ActiveCell.FormulaR1C1 = UseDataset.CalcVelocityHead
542                    Case Is = 32
543                       .ActiveCell.FormulaR1C1 = UseDataset.CalcTDH
544                    Case Is = 33
545                        .ActiveCell.FormulaR1C1 = UseDataset.CalcOverallEfficiency
546                    Case Is = 34
547                        .ActiveCell.FormulaR1C1 = UseDataset.CalcMotorEfficiency
548                    Case Is = 35
549                        .ActiveCell.FormulaR1C1 = UseDataset.CalcHydraulicEfficiency
550                    Case Is = 36
551                        .ActiveCell.FormulaR1C1 = UseDataset.CalcPowerFactor
552                End Select
553            Next Row

554            .Columns("A:K").Select
555            .Selection.Columns.AutoFit
556        End With

' <VB WATCH>
557        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
558        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "WriteCalData"

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
            vbwReportVariable "DatasetNumber", DatasetNumber
            vbwReportVariable "Col", Col
            vbwReportVariable "Row", Row
            vbwReportVariable "cell", cell
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

Private Sub Form_Unload(Cancel As Integer)
' <VB WATCH>
559        On Error GoTo vbwErrHandler
560        Const VBWPROCNAME = "frmCalibrate.Form_Unload"
561        If vbwProtector.vbwTraceProc Then
562            Dim vbwProtectorParameterString As String
563            If vbwProtector.vbwTraceParameters Then
564                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Cancel", Cancel) & ") "
565            End If
566            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
567        End If
' </VB WATCH>
568        cmdExit_Click
' <VB WATCH>
569        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
570        Exit Sub
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

Public Sub NewWorkBook()
' <VB WATCH>
571        On Error GoTo vbwErrHandler
572        Const VBWPROCNAME = "frmCalibrate.NewWorkBook"
573        If vbwProtector.vbwTraceProc Then
574            Dim vbwProtectorParameterString As String
575            If vbwProtector.vbwTraceParameters Then
576                vbwProtectorParameterString = "()"
577            End If
578            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
579        End If
' </VB WATCH>

           'we've just added a new workbook, delete sheet1, sheet2, etc
580        xlApp.DisplayAlerts = False
581        While xlApp.Worksheets.Count > 1
582            xlApp.Worksheets(1).Delete          'delete the sheet
583        Wend
584        xlApp.DisplayAlerts = True

585        CalibrateWorkSheetName = InputBox("Enter Title Worksheet Name for this run.")    'get the desired name
586        xlApp.Worksheets(1).Name = CalibrateWorkSheetName    'and name the sheet

' <VB WATCH>
587        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
588        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "NewWorkBook"

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
Public Function GetWorksheetTabs()
' <VB WATCH>
589        On Error GoTo vbwErrHandler
590        Const VBWPROCNAME = "frmCalibrate.GetWorksheetTabs"
591        If vbwProtector.vbwTraceProc Then
592            Dim vbwProtectorParameterString As String
593            If vbwProtector.vbwTraceParameters Then
594                vbwProtectorParameterString = "()"
595            End If
596            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
597        End If
' </VB WATCH>

           'see what worksheet tabs alread exist in the excel worksheet

598        Dim intSheets As Integer    'number of sheets in the workbook
599        Dim I As Integer
600        Dim S As String
601        Dim ans
602        Dim NameOK As Boolean

603        intSheets = xlApp.Worksheets.Count      'how many sheets are there?

           'define a crlf string
604        S = vbCrLf

605        For I = 1 To intSheets
606            S = S & xlApp.Worksheets(I).Name & vbCrLf   'add in the worksheet name
607        Next I

           'tell the user the names so far and ask if he/she wants to add another
608        ans = MsgBox("You have the following Worksheet Names in " & sCalibrateSaveFileName & ": " & S & "Do you want to add another sheet to this file?", vbYesNo, "Sheets in Excel File")

           'get the answer
609        If ans = vbNo Then
610            GetWorksheetTabs = vbNo     'set up flag for when we return to the calling subroutine
' <VB WATCH>
611        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
612            Exit Function
613        End If

           'get worksheet name from user and check to see that it's not already used

614        NameOK = False  'start assuming that the name is bad

615        While Not NameOK    'as long as it's bad, stay in this loop
616            CalibrateWorkSheetName = InputBox("Enter Worksheet Name for this run.")  'ask for name

617            If CalibrateWorkSheetName = "" Then      'if we get a nul return or user presses cancel
618                GetWorksheetTabs = vbNo
' <VB WATCH>
619        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
620                Exit Function
621            End If

622            For I = 1 To xlApp.Worksheets.Count     'go through all of the existing sheets
623                If CalibrateWorkSheetName = xlApp.Worksheets(I).Name Then        'if the names are the same
624                    MsgBox "The name " & CalibrateWorkSheetName & " already exists for a Worksheet.  Please try again.", vbOKOnly, "Bad Worksheet Name"  'tell the user
625                    NameOK = False
626                    Exit For
627                End If
628                NameOK = True       'if we make it thru say the name is ok
629            Next I
630        Wend

631        xlApp.Worksheets.Add , xlApp.Worksheets(xlApp.Worksheets.Count)     'add a worksheer
632        xlApp.Worksheets(xlApp.Worksheets.Count).Name = CalibrateWorkSheetName       'give it the desired name
633        GetWorksheetTabs = vbYes                                            'say that the results were ok

' <VB WATCH>
634        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
635        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetWorksheetTabs"

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
            vbwReportVariable "intSheets", intSheets
            vbwReportVariable "I", I
            vbwReportVariable "S", S
            vbwReportVariable "ans", ans
            vbwReportVariable "NameOK", NameOK
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Private Sub DoCalibrationCalcs()
' <VB WATCH>
636        On Error GoTo vbwErrHandler
637        Const VBWPROCNAME = "frmCalibrate.DoCalibrationCalcs"
638        If vbwProtector.vbwTraceProc Then
639            Dim vbwProtectorParameterString As String
640            If vbwProtector.vbwTraceParameters Then
641                vbwProtectorParameterString = "()"
642            End If
643            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
644        End If
' </VB WATCH>
645        Dim KW As Single, VI As Single, VITemp As Single
646        Dim Vave As Single, Iave As Single
647        Dim I As Integer
648        Dim j As Integer
649        Dim HeightDiff As Single

650        If Not IsNull(UseDataset.PowerA) Then
651            KW = UseDataset.PowerA
652        End If
653        If Not IsNull(UseDataset.PowerB) Then
654            KW = KW + UseDataset.PowerB
655        End If
656        If Not IsNull(UseDataset.PowerC) Then
657            KW = KW + UseDataset.PowerC
658        End If

659        I = 0
660        Vave = 0
661        Iave = 0
662        If Not IsNull(UseDataset.VoltageA) And Not IsNull(UseDataset.CurrentA) Then
663            VI = UseDataset.VoltageA * UseDataset.CurrentA
664            Vave = UseDataset.VoltageA
665            Iave = UseDataset.CurrentA
666            If VI <> 0 Then
667                I = I + 1
668            End If
669        End If
670        If Not IsNull(UseDataset.VoltageB) And Not IsNull(UseDataset.CurrentB) Then
671            VITemp = UseDataset.VoltageB * UseDataset.CurrentB
672            If VITemp <> 0 Then
673                I = I + 1
674                VI = VI + VITemp
675                Vave = Vave + UseDataset.VoltageB
676                Iave = Iave + UseDataset.CurrentB
677            End If
678        End If
679        If Not IsNull(UseDataset.VoltageC) And Not IsNull(UseDataset.CurrentC) Then
680            VITemp = UseDataset.VoltageC * UseDataset.CurrentC
681            If VITemp <> 0 Then
682                I = I + 1
683                VI = VI + VITemp
684                Vave = Vave + UseDataset.VoltageC
685                Iave = Iave + UseDataset.CurrentC
686            End If
687        End If
688        If VI <> 0 Then
689            UseDataset.CalcPowerFactor = 1000 * I * KW / (VI * Sqr(3))
690            UseDataset.CalcPowerFactor = 100 * UseDataset.CalcPowerFactor
691        Else
692            UseDataset.CalcPowerFactor = 0
693        End If

694        UseDataset.CalcMotorEfficiency = Format$(Round(MotorEfficiency(KW, UseDataset.MotorType, UseDataset.StatorFill), 1), "00.0")

695        Dim sHDCor As Single
696        Dim sDisc As Single
697        Dim sSuct As Single

698        sDisc = UseDataset.DischargeHeight
699        sSuct = UseDataset.SuctionHeight

700        HeightDiff = UseDataset.HDCorr + sDisc / 12 - sSuct / 12

701        UseDataset.CalcVelocityHead = CalcVelHead(UseDataset.Flow, UseDataset.DischargePipeDia, UseDataset.SuctionPipeDia)

702        UseDataset.CalcTDH = CalcTDH(UseDataset.DischargePressure, UseDataset.SuctionPressure, UseDataset.SuctionInHg, UseDataset.CalcVelocityHead, HeightDiff, UseDataset.Temperature)

703        If Int(UseDataset.Temperature) >= 40 Then
704            If (DLookupA(TDHColNo, TempCorrection, TempColNo, Int(UseDataset.Temperature)) <> 0 And KW <> 0) Then
705                UseDataset.CalcOverallEfficiency = (0.189 * UseDataset.Flow * UseDataset.CalcTDH * DLookupA(TDHColNo, TempCorrection, TempColNo, 68)) / (10 * KW * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(UseDataset.Temperature)))
706                If UseDataset.CalcMotorEfficiency <> 0 Then
707                    UseDataset.CalcHydraulicEfficiency = 100 * UseDataset.CalcOverallEfficiency / UseDataset.CalcMotorEfficiency
708                Else
709                    UseDataset.CalcHydraulicEfficiency = 0
710                End If
711            Else
712                UseDataset.CalcOverallEfficiency = 0
713            End If
714        Else
       '        rsEff.Fields("LiquidHP") = 0
715            UseDataset.CalcOverallEfficiency = 0
716        End If

' <VB WATCH>
717        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
718        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DoCalibrationCalcs"

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
            vbwReportVariable "KW", KW
            vbwReportVariable "VI", VI
            vbwReportVariable "VITemp", VITemp
            vbwReportVariable "Vave", Vave
            vbwReportVariable "Iave", Iave
            vbwReportVariable "I", I
            vbwReportVariable "j", j
            vbwReportVariable "HeightDiff", HeightDiff
            vbwReportVariable "sHDCor", sHDCor
            vbwReportVariable "sDisc", sDisc
            vbwReportVariable "sSuct", sSuct
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
End Sub
' </VB WATCH>
