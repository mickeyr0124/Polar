Attribute VB_Name = "AccessRoutines"
Option Explicit
Global cnPumpData As New ADODB.Connection      'Pump Database connection
Global cnEffData As New ADODB.Connection       'local efficiency database connection

'initials to approve and delete testing
Global Const strApproveInitials As String = "Admin"
Global boCanApprove As Boolean
Global LogInInitials As String

Public Type DataResult
    HP As Double
    Speed As Double
End Type
Global results() As DataResult

Type DataSet
    Flow As Single                  'input flow
    SuctionPressure As Single       'input suct press
    DischargePressure As Single     'input disch press
    Temperature As Single           'input temp
    SuctionPipeDia As Integer       'input suct pipe dia
    DischargePipeDia As Integer     'input disch pipe dia
    SuctionHeight As Integer        'input suction gage height
    DischargeHeight As Integer      'input disch gage height
    BarometricPressure As Single    'input barometric pressure
    HDCorr As Single                'input HDCorr
    SuctionInHg As Single           'input suction in inHg
    MotorType As Long               'input motor type
    StatorFill As Long              'input stator fill type
    VoltageA As Single              'input voltage
    VoltageB As Single              'input voltage
    VoltageC As Single              'input voltage
    CurrentA As Single              'input current
    CurrentB As Single              'input current
    CurrentC As Single              'input current
    PowerA As Single                'input power
    PowerB As Single                'input power
    PowerC As Single                'input power
    PowerFactor As Single           'input power factor
    VelocityHead As Single          'output velocity head
    TDH As Single                   'output TDH
    OverallEfficiency As Single     'output Overall Efficiancy
    MotorEfficiency As Single       'output motor efficiency
    HydraulicEfficiency As Single   'output Hydraulic efficiency
    CalcPowerFactor As Single
    CalcVelocityHead As Single          'output velocity head
    CalcTDH As Single                   'output TDH
    CalcOverallEfficiency As Single     'output Overall Efficiancy
    CalcMotorEfficiency As Single       'output motor efficiency
    CalcHydraulicEfficiency As Single   'output Hydraulic efficiency
End Type

    Global DataSets(2) As DataSet
    Global UseDataset As DataSet
    Global Calibrating As Boolean          'in the process of calibrating
    Global sServerName As String

    Global Const sCalibrateDirectoryName = "EN\GROUPS\SHARED\Calibration and Rundown\Hydraulic Rundown Calibration"

    Global sCalibrateDatabaseName As String
    Global sCalibrateSaveFileName As String
    Global cnCalibrate As New ADODB.Connection
    Global rsCalibrate As New ADODB.Recordset

    Global xlApp As Excel.Application  ' Excel Application Object
    Global xlBook As Excel.Workbook    ' Excel Workbook Object

    Global CalibrateWorkSheetName As String         'Worksheet Tab Name
    Global WritingToCalFile As Boolean

    'Arrays for DLookup
    Public PipeDiameters As Variant
    Public VaporPressure As Variant
    Public TempCorrection As Variant
    Public TEMCForceViscosity As Variant

    'Column number constants
    Public Const IDColNo As Integer = 0
    Public Const NominalColNo As Integer = 1
    Public Const ActualColNo As Integer = 2
    Public Const TempColNo As Integer = 1
    Public Const VaporPressureColNo As Integer = 2
    Public Const SpecificVolumeColNo As Integer = 3
    Public Const TDHColNo As Integer = 3

    Public Declare Function OpenProcess _
            Lib "kernel32" _
            (ByVal dwDesiredAccess As Long, _
             ByVal bInheritHandle As Long, _
             ByVal dwProcessId As Long) As Long
    Public Declare Function CloseHandle _
            Lib "kernel32" _
            (ByVal hObject As Long) As Long
    Public Declare Function WaitForSingleObject _
            Lib "kernel32" _
            (ByVal hHandle As Long, _
             ByVal dwMilliseconds As Long) As Long





Public Function DLookup(sField As String, sDomain As String, Optional sCriteria As String) As Variant

    Dim oRs As New ADODB.Recordset
    Dim qy As New ADODB.Command

    DLookup = Empty

    qy.ActiveConnection = cnPumpData

    qy.CommandText = "SELECT " & sField & " FROM " & sDomain
    If LenB(sCriteria) <> 0 Then
        qy.CommandText = qy.CommandText & " WHERE " & sCriteria
    End If

    oRs.Open qy
    If Not oRs.EOF Then
        oRs.MoveFirst
        DLookup = oRs.Fields(sField).value
    End If
    oRs.Close
    Set oRs = Nothing
    Exit Function
  
End Function
Public Function DLookupA(ReturnColumnNo As Integer, ArrayName As Variant, FindColNo As Integer, FindValue As Variant) As Variant
    Dim I As Integer

    If FindValue = -1 Or IsNull(FindValue) Then
        DLookupA = Empty
        Exit Function
    End If

    DLookupA = 0
    For I = 0 To UBound(ArrayName, 2)
        If ArrayName(FindColNo, I) = FindValue Then
            DLookupA = ArrayName(ReturnColumnNo, I)
            Exit For
        End If
    Next I
  
End Function
Function MotorEfficiency(KW As Single, Motor As Long, StatorFill As Long)
    Dim eff0 As Single, eff1 As Single, eff2 As Single, eff3 As Single, eff4 As Single, eff5 As Single
    Dim kw0 As Single, kw1 As Single, kw2 As Single, kw3 As Single, kw4 As Single, kw5 As Single
    Dim qy As New ADODB.Command
    Dim rs As New ADODB.Recordset

    'select the testsetup data for the serial number
    qy.ActiveConnection = cnPumpData
    If StatorFill = 1 Then  'dry stator
        qy.CommandText = "SELECT * FROM MotorEfficiencies WHERE (((MotorEfficiencies.MotorKey)=" & Motor & ") AND ((MotorEfficiencies.Fill)='No')) OR (((MotorEfficiencies.MotorKey)=" & Motor & ") AND ((MotorEfficiencies.Fill)='Both'));"
    Else
        qy.CommandText = "SELECT * FROM MotorEfficiencies WHERE (((MotorEfficiencies.MotorKey)=" & Motor & ") AND ((MotorEfficiencies.Fill)='Yes')) OR (((MotorEfficiencies.MotorKey)=" & Motor & ") AND ((MotorEfficiencies.Fill)='Both'));"

    End If

    With rs     'open the recordset for the query
        .CursorLocation = adUseServer
        .CursorType = adOpenDynamic
        .Open qy
    End With

    If rs.BOF = True And rs.EOF = True Then
        MotorEfficiency = 0
        Exit Function
    End If

    If rs!in125 <> 0 Then
        kw5 = rs!in125
        eff5 = rs!eff125
    Else
        kw5 = rs!in100
        eff5 = rs!eff100
    End If

    kw4 = rs!in100
    kw3 = rs!in75
    kw2 = rs!in50
    kw1 = rs!in25
    kw0 = rs!in0

    eff4 = rs!eff100
    eff3 = rs!eff75
    eff2 = rs!eff50
    eff1 = rs!eff25
    eff0 = rs!eff0

    Select Case KW
        Case Is >= kw5
            MotorEfficiency = eff5      'trap at highest table entry

        Case Is >= kw4
            MotorEfficiency = Interpolate(eff5, eff4, kw5, kw4, KW)

        Case Is >= kw3
            MotorEfficiency = Interpolate(eff4, eff3, kw4, kw3, KW)

        Case Is >= kw2
            MotorEfficiency = Interpolate(eff3, eff2, kw3, kw2, KW)

        Case Is >= kw1
            MotorEfficiency = Interpolate(eff2, eff1, kw2, kw1, KW)

        Case Is < kw1
            MotorEfficiency = Interpolate(eff1, eff0, kw1, kw0, KW)

        Case Else
            MotorEfficiency = " "
    End Select
End Function
Function TEMCMotorEfficiency(KW As Single, ModelNumber As String, Voltage As String, RatedKW As Single)
    Dim eff0 As Single, eff1 As Single, eff2 As Single, eff3 As Single, eff4 As Single
    Dim kw0 As Single, kw1 As Single, kw2 As Single, kw3 As Single, kw4 As Single
    Dim qy As New ADODB.Command
    Dim rs As New ADODB.Recordset

    If ModelNumber = "" Then
        TEMCMotorEfficiency = 0
        RatedKW = 999
        Exit Function
    End If

    'select the testsetup data for the serial number
    qy.ActiveConnection = cnPumpData
    qy.CommandText = "SELECT TEMCMotorEfficienciesNew.* From TEMCMotorEfficienciesNew " & _
                      "WHERE ((TEMCMotorEfficienciesNew.ModelNumber)= " & ModelNumber & _
                      ") ;"
'        ") AND ((TEMCMotorEfficiencies.Voltage)= " & Voltage & "));"

    With rs     'open the recordset for the query
        .CursorLocation = adUseServer
        .CursorType = adOpenDynamic
        .Open qy
    End With

    If rs.BOF = True And rs.EOF = True Then
        TEMCMotorEfficiency = 0
        RatedKW = 999
        Exit Function
    End If

    kw4 = rs!in100
    kw3 = rs!in75
    kw2 = rs!in50
    kw1 = rs!in25
    kw0 = rs!in0
    eff4 = 100 * rs!eff100 / 100
    eff3 = 100 * rs!eff75 / 100
    eff2 = 100 * rs!eff50 / 100
    eff1 = 100 * rs!eff25 / 100
    eff0 = 100 * rs!eff0 / 100

    Select Case KW
        Case Is >= kw4
            TEMCMotorEfficiency = eff4          'trap at highest table entry

        Case Is >= kw3
            TEMCMotorEfficiency = Interpolate(eff4, eff3, kw4, kw3, KW)

        Case Is >= kw2
            TEMCMotorEfficiency = Interpolate(eff3, eff2, kw3, kw2, KW)

        Case Is >= kw1
            TEMCMotorEfficiency = Interpolate(eff2, eff1, kw2, kw1, KW)

        Case Is < kw1
            TEMCMotorEfficiency = Interpolate(eff1, eff0, kw1, kw0, KW)

        Case Else
            TEMCMotorEfficiency = " "
    End Select
    If rs!RatedOutput <> 0 Then
        RatedKW = rs!RatedOutput
    Else
        RatedKW = 999
    End If
  
End Function
Function Interpolate(HiEff, LowEff, HiKW, LowKW, ActualKW) As Single
    Dim PctKw As Single

    PctKw = (ActualKW - LowKW) / (HiKW - LowKW)
    Interpolate = PctKw * (HiEff - LowEff) + LowEff
  
End Function
Function CalculateSuctionPressure(SuctPress, SuctInHg)
    Dim sp As Single

    If (Not IsNumeric(SuctPress)) Then
        sp = 0
    Else
        sp = SuctPress
    End If

    CalculateSuctionPressure = sp - 0.4893 * SuctInHg
End Function

Function CalcVelHead(Flow, DischDiam, SuctDiam)
    If Not (DischDiam = 0 Or SuctDiam = 0) Then
        If Not ((SuctDiam = -1 Or DischDiam = -1) Or DLookupA(ActualColNo, PipeDiameters, IDColNo, SuctDiam) = 0) Then
            CalcVelHead = (0.00259 * Flow ^ 2 / DLookupA(ActualColNo, PipeDiameters, IDColNo, DischDiam) ^ 4) - (0.00259 * Flow ^ 2 / DLookupA(ActualColNo, PipeDiameters, IDColNo, SuctDiam) ^ 4)
        End If
    End If
End Function

Function CalcTDH(DischargePressure, SuctionPressure, SuctionInHg, VelHead, HDCorr, SuctTemp)
    If IsNull(HDCorr) Then
        HDCorr = 0
    End If
    If SuctTemp < 40 Or IsNull(SuctTemp) Then
        CalcTDH = 0
        Exit Function
    End If
'    CalcTDH = (DischargePressure - CalculateSuctionPressure(SuctionPressure, SuctionInHg)) * 144 * DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(SuctTemp)) + VelHead + HDCorr
    CalcTDH = (DischargePressure - CalculateSuctionPressure(SuctionPressure, SuctionInHg)) * 144 * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(SuctTemp)) + VelHead + HDCorr
  
End Function

Function FillArrays()

    'fill the arrays for dlookup
    Dim rsTemp As New ADODB.Recordset

    rsTemp.Open "PipeDiameters", cnPumpData, adOpenStatic, adLockReadOnly
    PipeDiameters = rsTemp.GetRows()
    rsTemp.Close
    rsTemp.Open "VaporPressure", cnPumpData, adOpenStatic, adLockReadOnly
    VaporPressure = rsTemp.GetRows()
    rsTemp.Close
    rsTemp.Open "TempCorrection", cnPumpData, adOpenStatic, adLockReadOnly
    TempCorrection = rsTemp.GetRows()
    rsTemp.Close
    rsTemp.Open "TEMCForceViscosity", cnPumpData, adOpenStatic, adLockReadOnly
    TEMCForceViscosity = rsTemp.GetRows()
    rsTemp.Close
    Set rsTemp = Nothing
End Function
Public Function PingSilent(strComputer) As Integer
    Dim PID As Long
    Dim hProcess As Long
    Dim str As String

    str = Environ$("comspec") & " /c ping -n 2 -w 300 " & strComputer & " | find /c ""Reply"" > """ & App.Path & "\pingdata.txt"""

    PID = Shell(str, vbHide)


    If PID = 0 Then
         '
         'Handle Error, Shell Didn't Work
         '
    Else
         hProcess = OpenProcess(&H100000, True, PID)
         WaitForSingleObject hProcess, -1
         CloseHandle hProcess
    End If

    Open App.Path & "\pingdata.txt" For Input As #1
    Input #1, str

    PingSilent = Val(str)

    Close #1
  
End Function



