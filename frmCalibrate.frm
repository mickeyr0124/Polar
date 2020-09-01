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

Private Sub cmdExit_Click()
    Dim I As Integer

    If rsCalibrate.State = adStateOpen Then
        rsCalibrate.Close
    End If
    If cnCalibrate.State = adStateOpen Then
        cnCalibrate.Close
    End If

    Unload Me
    Calibrating = False
    End
  
End Sub

Private Sub cmdRunCalibration_Click()
    Dim X As Integer

    cmdRunCalibration.Visible = False

    ' Create the Excel App Object so we can store our data
    Set xlApp = CreateObject("Excel.Application")

    OpenCalibrateFile

    If Not WritingToCalFile Then
        Exit Sub
    End If

    WriteCalHeader

    For X = 0 To 2
        UseDataset = DataSets(X)
        With MSFlexGrid1
            .Row = X + 1
            .RowSel = X + 1
            .Col = 0
            .ColSel = .Cols - 1
            .Highlight = flexHighlightAlways
        End With
        Calibrating = True

        DoCalibrationCalcs
        WriteCalData (X)
    Next X

    MSFlexGrid1.Highlight = flexHighlightNever
    xlApp.ActiveWorkbook.Save             'save the file

    xlApp.Application.Quit
    Set xlApp = Nothing

    cmdExit_Click
End Sub

Private Sub Form_Load()

    Dim X As Long
    Dim Count As Long

    sCalibrateDatabaseName = App.Path & "\CalibrateData.mdb"
    With cnCalibrate
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sCalibrateDatabaseName & ";Persist Security Info=False"
        .Open
    End With
    rsCalibrate.Open "Data", cnCalibrate, adOpenStatic, adLockOptimistic, adCmdTable

    With MSFlexGrid1

        .Redraw = False
        .Clear
        .Row = 0

        .Col = 0
        .ColWidth(0) = 750
        .Text = "Data Set"

        .Col = 1
        .ColWidth(1) = 1200
        .Text = "Flow"
        .ColAlignment(1) = flexAlignCenterCenter

        .Col = 2
        .ColWidth(2) = 1200
        .Text = "Disch Press"
        .ColAlignment(2) = flexAlignCenterCenter

        .Col = 3
        .ColWidth(3) = 1200
        .Text = "Suction Press"
        .ColAlignment(3) = flexAlignCenterCenter

        .Col = 4
        .ColWidth(4) = 1200
        .Text = "Temperature"
        .ColAlignment(4) = flexAlignCenterCenter

        'setup the minimum number of rows & add column headers
        .Rows = 2
        .FixedRows = 1
        .Row = 0
        For X = 2 To 5
            .Col = X - 2 + 1
            .Text = rsCalibrate.Fields(X).Name
            .ColData(X - 2 + 1) = rsCalibrate.Fields(X).Type
        Next

        .Rows = rsCalibrate.RecordCount + 1
        For Count = 1 To rsCalibrate.RecordCount

            .TextMatrix(Count, 0) = Count    'assign line number
            For X = 0 To 3
                'we use Variant conversion to avoid any possible NULL errors
                .TextMatrix(Count, X + 1) = "" & CVar(rsCalibrate.Fields(X + 2).value)
            Next
            rsCalibrate.MoveNext
        Next

        .Redraw = True
    End With

    rsCalibrate.MoveFirst

    For X = 0 To 2
        DataSets(X).Flow = rsCalibrate.Fields("Flow")
        DataSets(X).SuctionPressure = rsCalibrate.Fields("SuctPress")
        DataSets(X).DischargePressure = rsCalibrate.Fields("DischPress")
        DataSets(X).Temperature = rsCalibrate.Fields("temp")
        DataSets(X).SuctionPipeDia = rsCalibrate.Fields("SuctPipeDia")
        DataSets(X).DischargePipeDia = rsCalibrate.Fields("DischPipeDia")
        DataSets(X).SuctionHeight = rsCalibrate.Fields("SuctHeight")
        DataSets(X).DischargeHeight = rsCalibrate.Fields("DischHeight")
        DataSets(X).BarometricPressure = rsCalibrate.Fields("BaroPress")
        DataSets(X).HDCorr = rsCalibrate.Fields("HDCorr")
        DataSets(X).SuctionInHg = rsCalibrate.Fields("SuctionInHg")
        DataSets(X).MotorType = rsCalibrate.Fields("MotorType")
        DataSets(X).StatorFill = rsCalibrate.Fields("StatorFill")
        DataSets(X).VoltageA = rsCalibrate.Fields("VoltageA")
        DataSets(X).VoltageB = rsCalibrate.Fields("VoltageB")
        DataSets(X).VoltageC = rsCalibrate.Fields("VoltageC")
        DataSets(X).CurrentA = rsCalibrate.Fields("CurrentA")
        DataSets(X).CurrentB = rsCalibrate.Fields("CurrentB")
        DataSets(X).CurrentC = rsCalibrate.Fields("CurrentC")
        DataSets(X).PowerA = rsCalibrate.Fields("PowerA")
        DataSets(X).PowerB = rsCalibrate.Fields("PowerB")
        DataSets(X).PowerC = rsCalibrate.Fields("PowerC")
        DataSets(X).PowerFactor = rsCalibrate.Fields("PowerFactor")
        DataSets(X).VelocityHead = rsCalibrate.Fields("VelocityHead")
        DataSets(X).TDH = rsCalibrate.Fields("TDH")
        DataSets(X).OverallEfficiency = rsCalibrate.Fields("OverallEfficiency")
        DataSets(X).MotorEfficiency = rsCalibrate.Fields("MotorEfficiency")
        DataSets(X).HydraulicEfficiency = rsCalibrate.Fields("HydraulicEfficiency")
        rsCalibrate.MoveNext
    Next X

    rsCalibrate.Close
  
End Sub
Private Sub OpenCalibrateFile()
        frmPLCData.CommonDialog1.CancelError = True        'in case the user
        On Error GoTo ErrHandler                '  chooses the cancel button

        'set up dialog box
        frmPLCData.CommonDialog1.DialogTitle = "Open Excel Calibration Files"
        frmPLCData.CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|"  'show Excel files
        frmPLCData.CommonDialog1.InitDir = sServerName & sCalibrateDirectoryName & "\Software Calibration"    'in this directory
        frmPLCData.CommonDialog1.ShowOpen                              'open the file selection dialog box

        If Dir(frmPLCData.CommonDialog1.filename) = "" Then            'if the file name does not exist yet
            sCalibrateSaveFileName = frmPLCData.CommonDialog1.filename           'get the name of the file
            If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
                 xlApp.Workbooks.Close
            End If
            ' Create the Excel Workbook Object.
On Error GoTo 0
            Set xlBook = xlApp.Workbooks.Add                'add a workbook
            NewWorkBook                                     'do some stuff for the new workbook
            xlApp.ActiveWorkbook.SaveAs filename:=sCalibrateSaveFileName, _
                              FileFormat:=xlNormal                        'save the file
            MsgBox frmPLCData.CommonDialog1.filename & " has been opened for writing.", vbOKOnly, "File Opened"    'tell the user that file is open
        Else                                                'the file name already exists
            sCalibrateSaveFileName = frmPLCData.CommonDialog1.filename
            ' Create the Excel Workbook Object.
            If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
                 xlApp.Workbooks.Close
            End If
            Set xlBook = xlApp.Workbooks.Open(sCalibrateSaveFileName)             'get the file name selected
            If GetWorksheetTabs = vbNo Then     'ask the user if he/she wants a new tab.
                MsgBox "File not overwritten.", vbOKOnly, "File not Opened"
                Exit Sub
            Else
                MsgBox frmPLCData.CommonDialog1.filename & " has been opened for writing.", vbOKOnly, "File Opened"
            End If
        End If

On Error GoTo 0

    WritingToCalFile = True

    Exit Sub

ErrHandler:
    'User pressed the Cancel button

    WritingToCalFile = False

    Exit Sub
  
End Sub
Public Sub WriteCalHeader()
    Dim TextToWrite As String
    Dim RowNo As Integer

        'write the header to the file
    With xlApp
        .Range("B1").Select
        .ActiveCell.FormulaR1C1 = "Hydraulic Rundown Calibration"
        .Selection.HorizontalAlignment = xlCenter

        .Range("A3").Select
        .ActiveCell.FormulaR1C1 = "Date - "

        .Range("B3").Select
        .ActiveCell.FormulaR1C1 = Now

         .Range("A4").Select
        .ActiveCell.FormulaR1C1 = "Data Set"

        .Range("C4:E4").Select
        .Selection.Merge
        .ActiveCell.FormulaR1C1 = "1"

        .Range("C5").Select
        .ActiveCell.FormulaR1C1 = "Input"
        .Range("D5").Select
        .ActiveCell.FormulaR1C1 = "Correct"
        .Range("E5").Select
        .ActiveCell.FormulaR1C1 = "Calculated"

        .Range("F4:H4").Select
        .Selection.Merge
        .ActiveCell.FormulaR1C1 = "2"

        .Range("F5").Select
        .ActiveCell.FormulaR1C1 = "Input"
        .Range("G5").Select
        .ActiveCell.FormulaR1C1 = "Correct"
        .Range("H5").Select
        .ActiveCell.FormulaR1C1 = "Calculated"

        .Range("I4:K4").Select
        .Selection.Merge
        .ActiveCell.FormulaR1C1 = "3"

        .Range("I5").Select
        .ActiveCell.FormulaR1C1 = "Input"
        .Range("J5").Select
        .ActiveCell.FormulaR1C1 = "Correct"
        .Range("K5").Select
        .ActiveCell.FormulaR1C1 = "Calculated"

        .Range("C4:K5").Select
        .Selection.HorizontalAlignment = xlCenter

        .Range("A6").Select
        .ActiveCell.FormulaR1C1 = "Inputs"
        .Selection.Font.Bold = True

        .Range("A7").Select
        .ActiveCell.FormulaR1C1 = "Flow"

        .Range("A8").Select
        .ActiveCell.FormulaR1C1 = "Suction Pressure"

         .Range("A9").Select
        .ActiveCell.FormulaR1C1 = "Discharge Pressure"

        .Range("A10").Select
        .ActiveCell.FormulaR1C1 = "Temperature"

        .Range("A11").Select
        .ActiveCell.FormulaR1C1 = "Suction Pipe Dia"

        .Range("A12").Select
        .ActiveCell.FormulaR1C1 = "Discharge Pipe Dia"

        .Range("A13").Select
        .ActiveCell.FormulaR1C1 = "Suction Gauge Height"

        .Range("A14").Select
        .ActiveCell.FormulaR1C1 = "Discharge Gauge Height"

        .Range("A15").Select
        .ActiveCell.FormulaR1C1 = "Barometric Pressure"

        .Range("A16").Select
        .ActiveCell.FormulaR1C1 = "HDCorr"

        .Range("A17").Select
        .ActiveCell.FormulaR1C1 = "Suction (InHg)"

        .Range("A18").Select
        .ActiveCell.FormulaR1C1 = "Motor Type"

        .Range("A19").Select
        .ActiveCell.FormulaR1C1 = "Voltage A"

        .Range("A20").Select
        .ActiveCell.FormulaR1C1 = "Voltage B"

        .Range("A21").Select
        .ActiveCell.FormulaR1C1 = "Voltage C"

        .Range("A22").Select
        .ActiveCell.FormulaR1C1 = "Current A"

        .Range("A23").Select
        .ActiveCell.FormulaR1C1 = "Current B"

        .Range("A24").Select
        .ActiveCell.FormulaR1C1 = "Current C"

        .Range("A25").Select
        .ActiveCell.FormulaR1C1 = "Power A"

        .Range("A26").Select
        .ActiveCell.FormulaR1C1 = "Power B"

        .Range("A27").Select
        .ActiveCell.FormulaR1C1 = "Power C"

        .Range("A28").Select
        .ActiveCell.FormulaR1C1 = "Stator Fill"

        .Range("A30").Select
        .ActiveCell.FormulaR1C1 = "Calculated Values"
        .Selection.Font.Bold = True

        .Range("A31").Select
        .ActiveCell.FormulaR1C1 = "Velocity Head"

        .Range("A32").Select
        .ActiveCell.FormulaR1C1 = "TDH"

        .Range("A33").Select
        .ActiveCell.FormulaR1C1 = "Overall Eff"

        .Range("A34").Select
        .ActiveCell.FormulaR1C1 = "Motor Eff"

        .Range("A35").Select
        .ActiveCell.FormulaR1C1 = "Hydraulic Eff"

        .Range("A36").Select
        .ActiveCell.FormulaR1C1 = "Power Factor"


        .Range("D30").Select
        .ActiveCell.FormulaR1C1 = "Correct"

        .Range("E30").Select
        .ActiveCell.FormulaR1C1 = "Calculated"

        .Range("G30").Select
        .ActiveCell.FormulaR1C1 = "Correct"

        .Range("H30").Select
        .ActiveCell.FormulaR1C1 = "Calculated"

        .Range("J30").Select
        .ActiveCell.FormulaR1C1 = "Correct"

        .Range("K30").Select
        .ActiveCell.FormulaR1C1 = "Calculated"

        .Range("C7:K36").Select
        .Selection.NumberFormat = "0.00"

        Range("D30:E36").Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With

        Range("G30:H36").Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With

        Range("J30:K36").Select
        Selection.Borders(xlDiagonalDown).LineStyle = xlNone
        Selection.Borders(xlDiagonalUp).LineStyle = xlNone
        With Selection.Borders(xlEdgeLeft)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeTop)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeBottom)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlEdgeRight)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlInsideVertical)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With
        With Selection.Borders(xlInsideHorizontal)
            .LineStyle = xlContinuous
            .Weight = xlThin
            .ColorIndex = xlAutomatic
        End With

        .Range("B35").Select
        .ActiveCell.FormulaR1C1 = "For formulas see:"
        .Selection.Font.Bold = True

        .Range("B36").Select
        ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:= _
                      sServerName & "EN\GROUPS\SHARED\Calibration and Rundown\Hydraulic Rundown Calibration\Software Calibration\Calibration Reference Sheet.xls" _
                      , TextToDisplay:="Calibration Reference Sheet"

        With ActiveSheet.PageSetup
            .Orientation = xlLandscape
        End With

    End With
End Sub
Public Sub WriteCalData(DatasetNumber As Integer)
    Dim Col As String
    Dim Row As Integer
    Dim cell As String

    Select Case DatasetNumber
        Case 0
            Col = "C"
        Case 1
            Col = "F"
        Case 2
            Col = "I"
        Case Else
    End Select

    With xlApp
        For Row = 7 To 28
            cell = Col & Trim(str(Row))
            .Range(cell).Select
            Select Case Row
                Case Is = 7
                    .ActiveCell.FormulaR1C1 = UseDataset.Flow
                Case Is = 8
                    .ActiveCell.FormulaR1C1 = UseDataset.SuctionPressure
                Case Is = 9
                    .ActiveCell.FormulaR1C1 = UseDataset.DischargePressure
                Case Is = 10
                    .ActiveCell.FormulaR1C1 = UseDataset.Temperature
                Case Is = 11
                    .ActiveCell.FormulaR1C1 = frmPLCData.cmbSuctDia.List(UseDataset.SuctionPipeDia - 1)
                Case Is = 12
                    .ActiveCell.FormulaR1C1 = frmPLCData.cmbDischDia.List(UseDataset.DischargePipeDia - 1)
                Case Is = 13
                    .ActiveCell.FormulaR1C1 = UseDataset.SuctionHeight
                Case Is = 14
                    .ActiveCell.FormulaR1C1 = UseDataset.DischargeHeight
                Case Is = 15
                    .ActiveCell.FormulaR1C1 = UseDataset.BarometricPressure
                Case Is = 16
                    .ActiveCell.FormulaR1C1 = UseDataset.HDCorr
                Case Is = 17
                    .ActiveCell.FormulaR1C1 = UseDataset.SuctionInHg
                Case Is = 18
                Dim I As Integer
            For I = 0 To frmPLCData.cmbMotor.ListCount - 1
            If frmPLCData.cmbMotor.ItemData(I) = UseDataset.MotorType Then
                .ActiveCell.FormulaR1C1 = frmPLCData.cmbMotor.List(I)
                Exit For
            End If
        Next I

'                    .ActiveCell.FormulaR1C1 = frmPLCData.cmbMotor.ItemData(UseDataset.MotorType)
                Case Is = 19
                    .ActiveCell.FormulaR1C1 = UseDataset.VoltageA
                Case Is = 20
                    .ActiveCell.FormulaR1C1 = UseDataset.VoltageB
                Case Is = 21
                    .ActiveCell.FormulaR1C1 = UseDataset.VoltageC
                Case Is = 22
                    .ActiveCell.FormulaR1C1 = UseDataset.CurrentA
                Case Is = 23
                    .ActiveCell.FormulaR1C1 = UseDataset.CurrentB
                Case Is = 24
                    .ActiveCell.FormulaR1C1 = UseDataset.CurrentC
                Case Is = 25
                    .ActiveCell.FormulaR1C1 = UseDataset.PowerA
                Case Is = 26
                    .ActiveCell.FormulaR1C1 = UseDataset.PowerB
                Case Is = 27
                    .ActiveCell.FormulaR1C1 = UseDataset.PowerC
                Case Is = 28
                    If UseDataset.StatorFill = 1 Then
                        .ActiveCell.FormulaR1C1 = "No"
                    Else
                        .ActiveCell.FormulaR1C1 = "Yes"
                    End If
'                    .ActiveCell.FormulaR1C1 = frmPLCData.cmbStatorFill.List(UseDataset.StatorFill)
            End Select
        Next Row

        Col = Chr(Asc(Col) + 1)
        For Row = 31 To 36
            cell = Col & Trim(str(Row))
            .Range(cell).Select
            Select Case Row
                Case Is = 31
                    .ActiveCell.FormulaR1C1 = UseDataset.VelocityHead
                Case Is = 32
                   .ActiveCell.FormulaR1C1 = UseDataset.TDH
                Case Is = 33
                    .ActiveCell.FormulaR1C1 = UseDataset.OverallEfficiency
                Case Is = 34
                    .ActiveCell.FormulaR1C1 = UseDataset.MotorEfficiency
                Case Is = 35
                    .ActiveCell.FormulaR1C1 = UseDataset.HydraulicEfficiency
                Case Is = 36
                    .ActiveCell.FormulaR1C1 = UseDataset.PowerFactor
            End Select
        Next Row

        Col = Chr(Asc(Col) + 1)
        For Row = 31 To 36
            cell = Col & Trim(str(Row))
            .Range(cell).Select
            Select Case Row
                Case Is = 31
                    .ActiveCell.FormulaR1C1 = UseDataset.CalcVelocityHead
                Case Is = 32
                   .ActiveCell.FormulaR1C1 = UseDataset.CalcTDH
                Case Is = 33
                    .ActiveCell.FormulaR1C1 = UseDataset.CalcOverallEfficiency
                Case Is = 34
                    .ActiveCell.FormulaR1C1 = UseDataset.CalcMotorEfficiency
                Case Is = 35
                    .ActiveCell.FormulaR1C1 = UseDataset.CalcHydraulicEfficiency
                Case Is = 36
                    .ActiveCell.FormulaR1C1 = UseDataset.CalcPowerFactor
            End Select
        Next Row

        .Columns("A:K").Select
        .Selection.Columns.AutoFit
    End With
  
End Sub

Private Sub Form_Unload(Cancel As Integer)
    cmdExit_Click
End Sub

Public Sub NewWorkBook()

    'we've just added a new workbook, delete sheet1, sheet2, etc
    xlApp.DisplayAlerts = False
    While xlApp.Worksheets.Count > 1
        xlApp.Worksheets(1).Delete          'delete the sheet
    Wend
    xlApp.DisplayAlerts = True

    CalibrateWorkSheetName = InputBox("Enter Title Worksheet Name for this run.")    'get the desired name
    xlApp.Worksheets(1).Name = CalibrateWorkSheetName    'and name the sheet
  
End Sub
Public Function GetWorksheetTabs()

    'see what worksheet tabs alread exist in the excel worksheet

    Dim intSheets As Integer    'number of sheets in the workbook
    Dim I As Integer
    Dim S As String
    Dim ans
    Dim NameOK As Boolean

    intSheets = xlApp.Worksheets.Count      'how many sheets are there?

    'define a crlf string
    S = vbCrLf

    For I = 1 To intSheets
        S = S & xlApp.Worksheets(I).Name & vbCrLf   'add in the worksheet name
    Next I

    'tell the user the names so far and ask if he/she wants to add another
    ans = MsgBox("You have the following Worksheet Names in " & sCalibrateSaveFileName & ": " & S & "Do you want to add another sheet to this file?", vbYesNo, "Sheets in Excel File")

    'get the answer
    If ans = vbNo Then
        GetWorksheetTabs = vbNo     'set up flag for when we return to the calling subroutine
        Exit Function
    End If

    'get worksheet name from user and check to see that it's not already used

    NameOK = False  'start assuming that the name is bad

    While Not NameOK    'as long as it's bad, stay in this loop
        CalibrateWorkSheetName = InputBox("Enter Worksheet Name for this run.")  'ask for name

        If CalibrateWorkSheetName = "" Then      'if we get a nul return or user presses cancel
            GetWorksheetTabs = vbNo
            Exit Function
        End If

        For I = 1 To xlApp.Worksheets.Count     'go through all of the existing sheets
            If CalibrateWorkSheetName = xlApp.Worksheets(I).Name Then        'if the names are the same
                MsgBox "The name " & CalibrateWorkSheetName & " already exists for a Worksheet.  Please try again.", vbOKOnly, "Bad Worksheet Name"  'tell the user
                NameOK = False
                Exit For
            End If
            NameOK = True       'if we make it thru say the name is ok
        Next I
    Wend

    xlApp.Worksheets.Add , xlApp.Worksheets(xlApp.Worksheets.Count)     'add a worksheer
    xlApp.Worksheets(xlApp.Worksheets.Count).Name = CalibrateWorkSheetName       'give it the desired name
    GetWorksheetTabs = vbYes                                            'say that the results were ok
  
End Function
Private Sub DoCalibrationCalcs()
    Dim KW As Single, VI As Single, VITemp As Single
    Dim Vave As Single, Iave As Single
    Dim I As Integer
    Dim j As Integer
    Dim HeightDiff As Single

    If Not IsNull(UseDataset.PowerA) Then
        KW = UseDataset.PowerA
    End If
    If Not IsNull(UseDataset.PowerB) Then
        KW = KW + UseDataset.PowerB
    End If
    If Not IsNull(UseDataset.PowerC) Then
        KW = KW + UseDataset.PowerC
    End If

    I = 0
    Vave = 0
    Iave = 0
    If Not IsNull(UseDataset.VoltageA) And Not IsNull(UseDataset.CurrentA) Then
        VI = UseDataset.VoltageA * UseDataset.CurrentA
        Vave = UseDataset.VoltageA
        Iave = UseDataset.CurrentA
        If VI <> 0 Then
            I = I + 1
        End If
    End If
    If Not IsNull(UseDataset.VoltageB) And Not IsNull(UseDataset.CurrentB) Then
        VITemp = UseDataset.VoltageB * UseDataset.CurrentB
        If VITemp <> 0 Then
            I = I + 1
            VI = VI + VITemp
            Vave = Vave + UseDataset.VoltageB
            Iave = Iave + UseDataset.CurrentB
        End If
    End If
    If Not IsNull(UseDataset.VoltageC) And Not IsNull(UseDataset.CurrentC) Then
        VITemp = UseDataset.VoltageC * UseDataset.CurrentC
        If VITemp <> 0 Then
            I = I + 1
            VI = VI + VITemp
            Vave = Vave + UseDataset.VoltageC
            Iave = Iave + UseDataset.CurrentC
        End If
    End If
    If VI <> 0 Then
        UseDataset.CalcPowerFactor = 1000 * I * KW / (VI * Sqr(3))
        UseDataset.CalcPowerFactor = 100 * UseDataset.CalcPowerFactor
    Else
        UseDataset.CalcPowerFactor = 0
    End If

    UseDataset.CalcMotorEfficiency = Format$(Round(MotorEfficiency(KW, UseDataset.MotorType, UseDataset.StatorFill), 1), "00.0")

    Dim sHDCor As Single
    Dim sDisc As Single
    Dim sSuct As Single

    sDisc = UseDataset.DischargeHeight
    sSuct = UseDataset.SuctionHeight

    HeightDiff = UseDataset.HDCorr + sDisc / 12 - sSuct / 12

    UseDataset.CalcVelocityHead = CalcVelHead(UseDataset.Flow, UseDataset.DischargePipeDia, UseDataset.SuctionPipeDia)

    UseDataset.CalcTDH = CalcTDH(UseDataset.DischargePressure, UseDataset.SuctionPressure, UseDataset.SuctionInHg, UseDataset.CalcVelocityHead, HeightDiff, UseDataset.Temperature)

    If Int(UseDataset.Temperature) >= 40 Then
        If (DLookupA(TDHColNo, TempCorrection, TempColNo, Int(UseDataset.Temperature)) <> 0 And KW <> 0) Then
            UseDataset.CalcOverallEfficiency = (0.189 * UseDataset.Flow * UseDataset.CalcTDH * DLookupA(TDHColNo, TempCorrection, TempColNo, 68)) / (10 * KW * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(UseDataset.Temperature)))
            If UseDataset.CalcMotorEfficiency <> 0 Then
                UseDataset.CalcHydraulicEfficiency = 100 * UseDataset.CalcOverallEfficiency / UseDataset.CalcMotorEfficiency
            Else
                UseDataset.CalcHydraulicEfficiency = 0
            End If
        Else
            UseDataset.CalcOverallEfficiency = 0
        End If
    Else
'        rsEff.Fields("LiquidHP") = 0
        UseDataset.CalcOverallEfficiency = 0
    End If
  
End Sub




