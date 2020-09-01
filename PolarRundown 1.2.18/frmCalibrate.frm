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
' </VB WATCH>
2          Dim I As Integer

3          If rsCalibrate.State = adStateOpen Then
4              rsCalibrate.Close
5          End If
6          If cnCalibrate.State = adStateOpen Then
7              cnCalibrate.Close
8          End If

9          Unload Me
10         Calibrating = False
11         End

' <VB WATCH>
12         Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub cmdRunCalibration_Click()
' <VB WATCH>
13         On Error GoTo vbwErrHandler
' </VB WATCH>
14         Dim X As Integer

15         cmdRunCalibration.Visible = False

           ' Create the Excel App Object so we can store our data
16         Set xlApp = CreateObject("Excel.Application")

17         OpenCalibrateFile

18         If Not WritingToCalFile Then
19             Exit Sub
20         End If

21         WriteCalHeader

22         For X = 0 To 2
23             UseDataset = DataSets(X)
24             With MSFlexGrid1
25                 .Row = X + 1
26                 .RowSel = X + 1
27                 .Col = 0
28                 .ColSel = .Cols - 1
29                 .Highlight = flexHighlightAlways
30             End With
31             Calibrating = True

32             DoCalibrationCalcs
33             WriteCalData (X)
34         Next X

35         MSFlexGrid1.Highlight = flexHighlightNever
36         xlApp.ActiveWorkbook.Save             'save the file

37         xlApp.Application.Quit
38         Set xlApp = Nothing

39         cmdExit_Click
' <VB WATCH>
40         Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub Form_Load()
' <VB WATCH>
41         On Error GoTo vbwErrHandler
' </VB WATCH>

42         Dim X As Long
43         Dim Count As Long

44         sCalibrateDatabaseName = App.Path & "\CalibrateData.mdb"
45         With cnCalibrate
46             .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sCalibrateDatabaseName & ";Persist Security Info=False"
47             .Open
48         End With
49         rsCalibrate.Open "Data", cnCalibrate, adOpenStatic, adLockOptimistic, adCmdTable

50         With MSFlexGrid1

51             .Redraw = False
52             .Clear
53             .Row = 0

54             .Col = 0
55             .ColWidth(0) = 750
56             .Text = "Data Set"

57             .Col = 1
58             .ColWidth(1) = 1200
59             .Text = "Flow"
60             .ColAlignment(1) = flexAlignCenterCenter

61             .Col = 2
62             .ColWidth(2) = 1200
63             .Text = "Disch Press"
64             .ColAlignment(2) = flexAlignCenterCenter

65             .Col = 3
66             .ColWidth(3) = 1200
67             .Text = "Suction Press"
68             .ColAlignment(3) = flexAlignCenterCenter

69             .Col = 4
70             .ColWidth(4) = 1200
71             .Text = "Temperature"
72             .ColAlignment(4) = flexAlignCenterCenter

               'setup the minimum number of rows & add column headers
73             .Rows = 2
74             .FixedRows = 1
75             .Row = 0
76             For X = 2 To 5
77                 .Col = X - 2 + 1
78                 .Text = rsCalibrate.Fields(X).Name
79                 .ColData(X - 2 + 1) = rsCalibrate.Fields(X).Type
80             Next

81             .Rows = rsCalibrate.RecordCount + 1
82             For Count = 1 To rsCalibrate.RecordCount

83                 .TextMatrix(Count, 0) = Count    'assign line number
84                 For X = 0 To 3
                       'we use Variant conversion to avoid any possible NULL errors
85                     .TextMatrix(Count, X + 1) = "" & CVar(rsCalibrate.Fields(X + 2).value)
86                 Next
87                 rsCalibrate.MoveNext
88             Next

89             .Redraw = True
90         End With

91         rsCalibrate.MoveFirst

92         For X = 0 To 2
93             DataSets(X).Flow = rsCalibrate.Fields("Flow")
94             DataSets(X).SuctionPressure = rsCalibrate.Fields("SuctPress")
95             DataSets(X).DischargePressure = rsCalibrate.Fields("DischPress")
96             DataSets(X).Temperature = rsCalibrate.Fields("temp")
97             DataSets(X).SuctionPipeDia = rsCalibrate.Fields("SuctPipeDia")
98             DataSets(X).DischargePipeDia = rsCalibrate.Fields("DischPipeDia")
99             DataSets(X).SuctionHeight = rsCalibrate.Fields("SuctHeight")
100            DataSets(X).DischargeHeight = rsCalibrate.Fields("DischHeight")
101            DataSets(X).BarometricPressure = rsCalibrate.Fields("BaroPress")
102            DataSets(X).HDCorr = rsCalibrate.Fields("HDCorr")
103            DataSets(X).SuctionInHg = rsCalibrate.Fields("SuctionInHg")
104            DataSets(X).MotorType = rsCalibrate.Fields("MotorType")
105            DataSets(X).StatorFill = rsCalibrate.Fields("StatorFill")
106            DataSets(X).VoltageA = rsCalibrate.Fields("VoltageA")
107            DataSets(X).VoltageB = rsCalibrate.Fields("VoltageB")
108            DataSets(X).VoltageC = rsCalibrate.Fields("VoltageC")
109            DataSets(X).CurrentA = rsCalibrate.Fields("CurrentA")
110            DataSets(X).CurrentB = rsCalibrate.Fields("CurrentB")
111            DataSets(X).CurrentC = rsCalibrate.Fields("CurrentC")
112            DataSets(X).PowerA = rsCalibrate.Fields("PowerA")
113            DataSets(X).PowerB = rsCalibrate.Fields("PowerB")
114            DataSets(X).PowerC = rsCalibrate.Fields("PowerC")
115            DataSets(X).PowerFactor = rsCalibrate.Fields("PowerFactor")
116            DataSets(X).VelocityHead = rsCalibrate.Fields("VelocityHead")
117            DataSets(X).TDH = rsCalibrate.Fields("TDH")
118            DataSets(X).OverallEfficiency = rsCalibrate.Fields("OverallEfficiency")
119            DataSets(X).MotorEfficiency = rsCalibrate.Fields("MotorEfficiency")
120            DataSets(X).HydraulicEfficiency = rsCalibrate.Fields("HydraulicEfficiency")
121            rsCalibrate.MoveNext
122        Next X

123        rsCalibrate.Close

' <VB WATCH>
124        Exit Sub
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
Private Sub OpenCalibrateFile()
' <VB WATCH>
125        On Error GoTo vbwErrHandler
' </VB WATCH>
126            frmPLCData.CommonDialog1.CancelError = True        'in case the user
127            On Error GoTo ErrHandler                '  chooses the cancel button

               'set up dialog box
128            frmPLCData.CommonDialog1.DialogTitle = "Open Excel Calibration Files"
129            frmPLCData.CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|"  'show Excel files
130            frmPLCData.CommonDialog1.InitDir = sServerName & sCalibrateDirectoryName & "\Software Calibration"    'in this directory
131            frmPLCData.CommonDialog1.ShowOpen                              'open the file selection dialog box

132            If Dir(frmPLCData.CommonDialog1.filename) = "" Then            'if the file name does not exist yet
133                sCalibrateSaveFileName = frmPLCData.CommonDialog1.filename           'get the name of the file
134                If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
135                     xlApp.Workbooks.Close
136                End If
                   ' Create the Excel Workbook Object.
137    On Error GoTo vbwErrHandler
138                Set xlBook = xlApp.Workbooks.Add                'add a workbook
139                NewWorkBook                                     'do some stuff for the new workbook
140                xlApp.ActiveWorkbook.SaveAs filename:=sCalibrateSaveFileName, _
                                     FileFormat:=xlNormal                        'save the file
141                MsgBox frmPLCData.CommonDialog1.filename & " has been opened for writing.", vbOKOnly, "File Opened"    'tell the user that file is open
142            Else                                                'the file name already exists
143                sCalibrateSaveFileName = frmPLCData.CommonDialog1.filename
                   ' Create the Excel Workbook Object.
144                If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
145                     xlApp.Workbooks.Close
146                End If
147                Set xlBook = xlApp.Workbooks.Open(sCalibrateSaveFileName)             'get the file name selected
148                If GetWorksheetTabs = vbNo Then     'ask the user if he/she wants a new tab.
149                    MsgBox "File not overwritten.", vbOKOnly, "File not Opened"
150                    Exit Sub
151                Else
152                    MsgBox frmPLCData.CommonDialog1.filename & " has been opened for writing.", vbOKOnly, "File Opened"
153                End If
154            End If

155    On Error GoTo vbwErrHandler

156        WritingToCalFile = True

157        Exit Sub

158    ErrHandler:
           'User pressed the Cancel button

159        WritingToCalFile = False

160        Exit Sub

' <VB WATCH>
161        Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Public Sub WriteCalHeader()
' <VB WATCH>
162        On Error GoTo vbwErrHandler
' </VB WATCH>
163        Dim TextToWrite As String
164        Dim RowNo As Integer

               'write the header to the file
165        With xlApp
166            .Range("B1").Select
167            .ActiveCell.FormulaR1C1 = "Hydraulic Rundown Calibration"
168            .Selection.HorizontalAlignment = xlCenter

169            .Range("A3").Select
170            .ActiveCell.FormulaR1C1 = "Date - "

171            .Range("B3").Select
172            .ActiveCell.FormulaR1C1 = Now

173             .Range("A4").Select
174            .ActiveCell.FormulaR1C1 = "Data Set"

175            .Range("C4:E4").Select
176            .Selection.Merge
177            .ActiveCell.FormulaR1C1 = "1"

178            .Range("C5").Select
179            .ActiveCell.FormulaR1C1 = "Input"
180            .Range("D5").Select
181            .ActiveCell.FormulaR1C1 = "Correct"
182            .Range("E5").Select
183            .ActiveCell.FormulaR1C1 = "Calculated"

184            .Range("F4:H4").Select
185            .Selection.Merge
186            .ActiveCell.FormulaR1C1 = "2"

187            .Range("F5").Select
188            .ActiveCell.FormulaR1C1 = "Input"
189            .Range("G5").Select
190            .ActiveCell.FormulaR1C1 = "Correct"
191            .Range("H5").Select
192            .ActiveCell.FormulaR1C1 = "Calculated"

193            .Range("I4:K4").Select
194            .Selection.Merge
195            .ActiveCell.FormulaR1C1 = "3"

196            .Range("I5").Select
197            .ActiveCell.FormulaR1C1 = "Input"
198            .Range("J5").Select
199            .ActiveCell.FormulaR1C1 = "Correct"
200            .Range("K5").Select
201            .ActiveCell.FormulaR1C1 = "Calculated"

202            .Range("C4:K5").Select
203            .Selection.HorizontalAlignment = xlCenter

204            .Range("A6").Select
205            .ActiveCell.FormulaR1C1 = "Inputs"
206            .Selection.Font.Bold = True

207            .Range("A7").Select
208            .ActiveCell.FormulaR1C1 = "Flow"

209            .Range("A8").Select
210            .ActiveCell.FormulaR1C1 = "Suction Pressure"

211             .Range("A9").Select
212            .ActiveCell.FormulaR1C1 = "Discharge Pressure"

213            .Range("A10").Select
214            .ActiveCell.FormulaR1C1 = "Temperature"

215            .Range("A11").Select
216            .ActiveCell.FormulaR1C1 = "Suction Pipe Dia"

217            .Range("A12").Select
218            .ActiveCell.FormulaR1C1 = "Discharge Pipe Dia"

219            .Range("A13").Select
220            .ActiveCell.FormulaR1C1 = "Suction Gauge Height"

221            .Range("A14").Select
222            .ActiveCell.FormulaR1C1 = "Discharge Gauge Height"

223            .Range("A15").Select
224            .ActiveCell.FormulaR1C1 = "Barometric Pressure"

225            .Range("A16").Select
226            .ActiveCell.FormulaR1C1 = "HDCorr"

227            .Range("A17").Select
228            .ActiveCell.FormulaR1C1 = "Suction (InHg)"

229            .Range("A18").Select
230            .ActiveCell.FormulaR1C1 = "Motor Type"

231            .Range("A19").Select
232            .ActiveCell.FormulaR1C1 = "Voltage A"

233            .Range("A20").Select
234            .ActiveCell.FormulaR1C1 = "Voltage B"

235            .Range("A21").Select
236            .ActiveCell.FormulaR1C1 = "Voltage C"

237            .Range("A22").Select
238            .ActiveCell.FormulaR1C1 = "Current A"

239            .Range("A23").Select
240            .ActiveCell.FormulaR1C1 = "Current B"

241            .Range("A24").Select
242            .ActiveCell.FormulaR1C1 = "Current C"

243            .Range("A25").Select
244            .ActiveCell.FormulaR1C1 = "Power A"

245            .Range("A26").Select
246            .ActiveCell.FormulaR1C1 = "Power B"

247            .Range("A27").Select
248            .ActiveCell.FormulaR1C1 = "Power C"

249            .Range("A28").Select
250            .ActiveCell.FormulaR1C1 = "Stator Fill"

251            .Range("A30").Select
252            .ActiveCell.FormulaR1C1 = "Calculated Values"
253            .Selection.Font.Bold = True

254            .Range("A31").Select
255            .ActiveCell.FormulaR1C1 = "Velocity Head"

256            .Range("A32").Select
257            .ActiveCell.FormulaR1C1 = "TDH"

258            .Range("A33").Select
259            .ActiveCell.FormulaR1C1 = "Overall Eff"

260            .Range("A34").Select
261            .ActiveCell.FormulaR1C1 = "Motor Eff"

262            .Range("A35").Select
263            .ActiveCell.FormulaR1C1 = "Hydraulic Eff"

264            .Range("A36").Select
265            .ActiveCell.FormulaR1C1 = "Power Factor"


266            .Range("D30").Select
267            .ActiveCell.FormulaR1C1 = "Correct"

268            .Range("E30").Select
269            .ActiveCell.FormulaR1C1 = "Calculated"

270            .Range("G30").Select
271            .ActiveCell.FormulaR1C1 = "Correct"

272            .Range("H30").Select
273            .ActiveCell.FormulaR1C1 = "Calculated"

274            .Range("J30").Select
275            .ActiveCell.FormulaR1C1 = "Correct"

276            .Range("K30").Select
277            .ActiveCell.FormulaR1C1 = "Calculated"

278            .Range("C7:K36").Select
279            .Selection.NumberFormat = "0.00"

280            Range("D30:E36").Select
281            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
282            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
283            With Selection.Borders(xlEdgeLeft)
284                .LineStyle = xlContinuous
285                .Weight = xlThin
286                .ColorIndex = xlAutomatic
287            End With
288            With Selection.Borders(xlEdgeTop)
289                .LineStyle = xlContinuous
290                .Weight = xlThin
291                .ColorIndex = xlAutomatic
292            End With
293            With Selection.Borders(xlEdgeBottom)
294                .LineStyle = xlContinuous
295                .Weight = xlThin
296                .ColorIndex = xlAutomatic
297            End With
298            With Selection.Borders(xlEdgeRight)
299                .LineStyle = xlContinuous
300                .Weight = xlThin
301                .ColorIndex = xlAutomatic
302            End With
303            With Selection.Borders(xlInsideVertical)
304                .LineStyle = xlContinuous
305                .Weight = xlThin
306                .ColorIndex = xlAutomatic
307            End With
308            With Selection.Borders(xlInsideHorizontal)
309                .LineStyle = xlContinuous
310                .Weight = xlThin
311                .ColorIndex = xlAutomatic
312            End With

313            Range("G30:H36").Select
314            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
315            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
316            With Selection.Borders(xlEdgeLeft)
317                .LineStyle = xlContinuous
318                .Weight = xlThin
319                .ColorIndex = xlAutomatic
320            End With
321            With Selection.Borders(xlEdgeTop)
322                .LineStyle = xlContinuous
323                .Weight = xlThin
324                .ColorIndex = xlAutomatic
325            End With
326            With Selection.Borders(xlEdgeBottom)
327                .LineStyle = xlContinuous
328                .Weight = xlThin
329                .ColorIndex = xlAutomatic
330            End With
331            With Selection.Borders(xlEdgeRight)
332                .LineStyle = xlContinuous
333                .Weight = xlThin
334                .ColorIndex = xlAutomatic
335            End With
336            With Selection.Borders(xlInsideVertical)
337                .LineStyle = xlContinuous
338                .Weight = xlThin
339                .ColorIndex = xlAutomatic
340            End With
341            With Selection.Borders(xlInsideHorizontal)
342                .LineStyle = xlContinuous
343                .Weight = xlThin
344                .ColorIndex = xlAutomatic
345            End With

346            Range("J30:K36").Select
347            Selection.Borders(xlDiagonalDown).LineStyle = xlNone
348            Selection.Borders(xlDiagonalUp).LineStyle = xlNone
349            With Selection.Borders(xlEdgeLeft)
350                .LineStyle = xlContinuous
351                .Weight = xlThin
352                .ColorIndex = xlAutomatic
353            End With
354            With Selection.Borders(xlEdgeTop)
355                .LineStyle = xlContinuous
356                .Weight = xlThin
357                .ColorIndex = xlAutomatic
358            End With
359            With Selection.Borders(xlEdgeBottom)
360                .LineStyle = xlContinuous
361                .Weight = xlThin
362                .ColorIndex = xlAutomatic
363            End With
364            With Selection.Borders(xlEdgeRight)
365                .LineStyle = xlContinuous
366                .Weight = xlThin
367                .ColorIndex = xlAutomatic
368            End With
369            With Selection.Borders(xlInsideVertical)
370                .LineStyle = xlContinuous
371                .Weight = xlThin
372                .ColorIndex = xlAutomatic
373            End With
374            With Selection.Borders(xlInsideHorizontal)
375                .LineStyle = xlContinuous
376                .Weight = xlThin
377                .ColorIndex = xlAutomatic
378            End With

379            .Range("B35").Select
380            .ActiveCell.FormulaR1C1 = "For formulas see:"
381            .Selection.Font.Bold = True

382            .Range("B36").Select
383            ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
                             sServerName & "EN\GROUPS\SHARED\Calibration and Rundown\Hydraulic Rundown Calibration\Software Calibration\Calibration Reference Sheet.xls" _
                             , TextToDisplay:="Calibration Reference Sheet"

384            With ActiveSheet.PageSetup
385                .Orientation = xlLandscape
386            End With

387        End With
' <VB WATCH>
388        Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Public Sub WriteCalData(DatasetNumber As Integer)
' <VB WATCH>
389        On Error GoTo vbwErrHandler
' </VB WATCH>
390        Dim Col As String
391        Dim Row As Integer
392        Dim cell As String

393        Select Case DatasetNumber
               Case 0
394                Col = "C"
395            Case 1
396                Col = "F"
397            Case 2
398                Col = "I"
399            Case Else
400        End Select

401        With xlApp
402            For Row = 7 To 28
403                cell = Col & Trim(str(Row))
404                .Range(cell).Select
405                Select Case Row
                       Case Is = 7
406                        .ActiveCell.FormulaR1C1 = UseDataset.Flow
407                    Case Is = 8
408                        .ActiveCell.FormulaR1C1 = UseDataset.SuctionPressure
409                    Case Is = 9
410                        .ActiveCell.FormulaR1C1 = UseDataset.DischargePressure
411                    Case Is = 10
412                        .ActiveCell.FormulaR1C1 = UseDataset.Temperature
413                    Case Is = 11
414                        .ActiveCell.FormulaR1C1 = frmPLCData.cmbSuctDia.List(UseDataset.SuctionPipeDia - 1)
415                    Case Is = 12
416                        .ActiveCell.FormulaR1C1 = frmPLCData.cmbDischDia.List(UseDataset.DischargePipeDia - 1)
417                    Case Is = 13
418                        .ActiveCell.FormulaR1C1 = UseDataset.SuctionHeight
419                    Case Is = 14
420                        .ActiveCell.FormulaR1C1 = UseDataset.DischargeHeight
421                    Case Is = 15
422                        .ActiveCell.FormulaR1C1 = UseDataset.BarometricPressure
423                    Case Is = 16
424                        .ActiveCell.FormulaR1C1 = UseDataset.HDCorr
425                    Case Is = 17
426                        .ActiveCell.FormulaR1C1 = UseDataset.SuctionInHg
427                    Case Is = 18
428                    Dim I As Integer
429                For I = 0 To frmPLCData.cmbMotor.ListCount - 1
430                If frmPLCData.cmbMotor.ItemData(I) = UseDataset.MotorType Then
431                    .ActiveCell.FormulaR1C1 = frmPLCData.cmbMotor.List(I)
432                    Exit For
433                End If
434            Next I

       '                    .ActiveCell.FormulaR1C1 = frmPLCData.cmbMotor.ItemData(UseDataset.MotorType)
435                    Case Is = 19
436                        .ActiveCell.FormulaR1C1 = UseDataset.VoltageA
437                    Case Is = 20
438                        .ActiveCell.FormulaR1C1 = UseDataset.VoltageB
439                    Case Is = 21
440                        .ActiveCell.FormulaR1C1 = UseDataset.VoltageC
441                    Case Is = 22
442                        .ActiveCell.FormulaR1C1 = UseDataset.CurrentA
443                    Case Is = 23
444                        .ActiveCell.FormulaR1C1 = UseDataset.CurrentB
445                    Case Is = 24
446                        .ActiveCell.FormulaR1C1 = UseDataset.CurrentC
447                    Case Is = 25
448                        .ActiveCell.FormulaR1C1 = UseDataset.PowerA
449                    Case Is = 26
450                        .ActiveCell.FormulaR1C1 = UseDataset.PowerB
451                    Case Is = 27
452                        .ActiveCell.FormulaR1C1 = UseDataset.PowerC
453                    Case Is = 28
454                        If UseDataset.StatorFill = 1 Then
455                            .ActiveCell.FormulaR1C1 = "No"
456                        Else
457                            .ActiveCell.FormulaR1C1 = "Yes"
458                        End If
       '                    .ActiveCell.FormulaR1C1 = frmPLCData.cmbStatorFill.List(UseDataset.StatorFill)
459                End Select
460            Next Row

461            Col = Chr(Asc(Col) + 1)
462            For Row = 31 To 36
463                cell = Col & Trim(str(Row))
464                .Range(cell).Select
465                Select Case Row
                       Case Is = 31
466                        .ActiveCell.FormulaR1C1 = UseDataset.VelocityHead
467                    Case Is = 32
468                       .ActiveCell.FormulaR1C1 = UseDataset.TDH
469                    Case Is = 33
470                        .ActiveCell.FormulaR1C1 = UseDataset.OverallEfficiency
471                    Case Is = 34
472                        .ActiveCell.FormulaR1C1 = UseDataset.MotorEfficiency
473                    Case Is = 35
474                        .ActiveCell.FormulaR1C1 = UseDataset.HydraulicEfficiency
475                    Case Is = 36
476                        .ActiveCell.FormulaR1C1 = UseDataset.PowerFactor
477                End Select
478            Next Row

479            Col = Chr(Asc(Col) + 1)
480            For Row = 31 To 36
481                cell = Col & Trim(str(Row))
482                .Range(cell).Select
483                Select Case Row
                       Case Is = 31
484                        .ActiveCell.FormulaR1C1 = UseDataset.CalcVelocityHead
485                    Case Is = 32
486                       .ActiveCell.FormulaR1C1 = UseDataset.CalcTDH
487                    Case Is = 33
488                        .ActiveCell.FormulaR1C1 = UseDataset.CalcOverallEfficiency
489                    Case Is = 34
490                        .ActiveCell.FormulaR1C1 = UseDataset.CalcMotorEfficiency
491                    Case Is = 35
492                        .ActiveCell.FormulaR1C1 = UseDataset.CalcHydraulicEfficiency
493                    Case Is = 36
494                        .ActiveCell.FormulaR1C1 = UseDataset.CalcPowerFactor
495                End Select
496            Next Row

497            .Columns("A:K").Select
498            .Selection.Columns.AutoFit
499        End With

' <VB WATCH>
500        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub Form_Unload(Cancel As Integer)
' <VB WATCH>
501        On Error GoTo vbwErrHandler
' </VB WATCH>
502        cmdExit_Click
' <VB WATCH>
503        Exit Sub
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

Public Sub NewWorkBook()
' <VB WATCH>
504        On Error GoTo vbwErrHandler
' </VB WATCH>

           'we've just added a new workbook, delete sheet1, sheet2, etc
505        xlApp.DisplayAlerts = False
506        While xlApp.Worksheets.Count > 1
507            xlApp.Worksheets(1).Delete          'delete the sheet
508        Wend
509        xlApp.DisplayAlerts = True

510        CalibrateWorkSheetName = InputBox("Enter Title Worksheet Name for this run.")    'get the desired name
511        xlApp.Worksheets(1).Name = CalibrateWorkSheetName    'and name the sheet

' <VB WATCH>
512        Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Public Function GetWorksheetTabs()
' <VB WATCH>
513        On Error GoTo vbwErrHandler
' </VB WATCH>

           'see what worksheet tabs alread exist in the excel worksheet

514        Dim intSheets As Integer    'number of sheets in the workbook
515        Dim I As Integer
516        Dim S As String
517        Dim ans
518        Dim NameOK As Boolean

519        intSheets = xlApp.Worksheets.Count      'how many sheets are there?

           'define a crlf string
520        S = vbCrLf

521        For I = 1 To intSheets
522            S = S & xlApp.Worksheets(I).Name & vbCrLf   'add in the worksheet name
523        Next I

           'tell the user the names so far and ask if he/she wants to add another
524        ans = MsgBox("You have the following Worksheet Names in " & sCalibrateSaveFileName & ": " & S & "Do you want to add another sheet to this file?", vbYesNo, "Sheets in Excel File")

           'get the answer
525        If ans = vbNo Then
526            GetWorksheetTabs = vbNo     'set up flag for when we return to the calling subroutine
527            Exit Function
528        End If

           'get worksheet name from user and check to see that it's not already used

529        NameOK = False  'start assuming that the name is bad

530        While Not NameOK    'as long as it's bad, stay in this loop
531            CalibrateWorkSheetName = InputBox("Enter Worksheet Name for this run.")  'ask for name

532            If CalibrateWorkSheetName = "" Then      'if we get a nul return or user presses cancel
533                GetWorksheetTabs = vbNo
534                Exit Function
535            End If

536            For I = 1 To xlApp.Worksheets.Count     'go through all of the existing sheets
537                If CalibrateWorkSheetName = xlApp.Worksheets(I).Name Then        'if the names are the same
538                    MsgBox "The name " & CalibrateWorkSheetName & " already exists for a Worksheet.  Please try again.", vbOKOnly, "Bad Worksheet Name"  'tell the user
539                    NameOK = False
540                    Exit For
541                End If
542                NameOK = True       'if we make it thru say the name is ok
543            Next I
544        Wend

545        xlApp.Worksheets.Add , xlApp.Worksheets(xlApp.Worksheets.Count)     'add a worksheer
546        xlApp.Worksheets(xlApp.Worksheets.Count).Name = CalibrateWorkSheetName       'give it the desired name
547        GetWorksheetTabs = vbYes                                            'say that the results were ok

' <VB WATCH>
548        Exit Function
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
    End Select
' </VB WATCH>
End Function
Private Sub DoCalibrationCalcs()
' <VB WATCH>
549        On Error GoTo vbwErrHandler
' </VB WATCH>
550        Dim KW As Single, VI As Single, VITemp As Single
551        Dim Vave As Single, Iave As Single
552        Dim I As Integer
553        Dim j As Integer
554        Dim HeightDiff As Single

555        If Not IsNull(UseDataset.PowerA) Then
556            KW = UseDataset.PowerA
557        End If
558        If Not IsNull(UseDataset.PowerB) Then
559            KW = KW + UseDataset.PowerB
560        End If
561        If Not IsNull(UseDataset.PowerC) Then
562            KW = KW + UseDataset.PowerC
563        End If

564        I = 0
565        Vave = 0
566        Iave = 0
567        If Not IsNull(UseDataset.VoltageA) And Not IsNull(UseDataset.CurrentA) Then
568            VI = UseDataset.VoltageA * UseDataset.CurrentA
569            Vave = UseDataset.VoltageA
570            Iave = UseDataset.CurrentA
571            If VI <> 0 Then
572                I = I + 1
573            End If
574        End If
575        If Not IsNull(UseDataset.VoltageB) And Not IsNull(UseDataset.CurrentB) Then
576            VITemp = UseDataset.VoltageB * UseDataset.CurrentB
577            If VITemp <> 0 Then
578                I = I + 1
579                VI = VI + VITemp
580                Vave = Vave + UseDataset.VoltageB
581                Iave = Iave + UseDataset.CurrentB
582            End If
583        End If
584        If Not IsNull(UseDataset.VoltageC) And Not IsNull(UseDataset.CurrentC) Then
585            VITemp = UseDataset.VoltageC * UseDataset.CurrentC
586            If VITemp <> 0 Then
587                I = I + 1
588                VI = VI + VITemp
589                Vave = Vave + UseDataset.VoltageC
590                Iave = Iave + UseDataset.CurrentC
591            End If
592        End If
593        If VI <> 0 Then
594            UseDataset.CalcPowerFactor = 1000 * I * KW / (VI * Sqr(3))
595            UseDataset.CalcPowerFactor = 100 * UseDataset.CalcPowerFactor
596        Else
597            UseDataset.CalcPowerFactor = 0
598        End If

599        UseDataset.CalcMotorEfficiency = Format$(Round(MotorEfficiency(KW, UseDataset.MotorType, UseDataset.StatorFill), 1), "00.0")

600        Dim sHDCor As Single
601        Dim sDisc As Single
602        Dim sSuct As Single

603        sDisc = UseDataset.DischargeHeight
604        sSuct = UseDataset.SuctionHeight

605        HeightDiff = UseDataset.HDCorr + sDisc / 12 - sSuct / 12

606        UseDataset.CalcVelocityHead = CalcVelHead(UseDataset.Flow, UseDataset.DischargePipeDia, UseDataset.SuctionPipeDia)

607        UseDataset.CalcTDH = CalcTDH(UseDataset.DischargePressure, UseDataset.SuctionPressure, UseDataset.SuctionInHg, UseDataset.CalcVelocityHead, HeightDiff, UseDataset.Temperature)

608        If Int(UseDataset.Temperature) >= 40 Then
609            If (DLookupA(TDHColNo, TempCorrection, TempColNo, Int(UseDataset.Temperature)) <> 0 And KW <> 0) Then
610                UseDataset.CalcOverallEfficiency = (0.189 * UseDataset.Flow * UseDataset.CalcTDH * DLookupA(TDHColNo, TempCorrection, TempColNo, 68)) / (10 * KW * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(UseDataset.Temperature)))
611                If UseDataset.CalcMotorEfficiency <> 0 Then
612                    UseDataset.CalcHydraulicEfficiency = 100 * UseDataset.CalcOverallEfficiency / UseDataset.CalcMotorEfficiency
613                Else
614                    UseDataset.CalcHydraulicEfficiency = 0
615                End If
616            Else
617                UseDataset.CalcOverallEfficiency = 0
618            End If
619        Else
       '        rsEff.Fields("LiquidHP") = 0
620            UseDataset.CalcOverallEfficiency = 0
621        End If

' <VB WATCH>
622        Exit Sub
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
    End Select
' </VB WATCH>
End Sub





