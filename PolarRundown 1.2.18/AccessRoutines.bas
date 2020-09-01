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





' <VB WATCH>
Const VBWMODULE = "AccessRoutines"
' </VB WATCH>

Public Function DLookup(sField As String, sDomain As String, Optional sCriteria As String) As Variant
' <VB WATCH>
1          On Error GoTo vbwErrHandler
' </VB WATCH>

2          Dim oRs As New ADODB.Recordset
3          Dim qy As New ADODB.Command

4          DLookup = Empty

5          qy.ActiveConnection = cnPumpData

6          qy.CommandText = "SELECT " & sField & " FROM " & sDomain
7          If LenB(sCriteria) <> 0 Then
8              qy.CommandText = qy.CommandText & " WHERE " & sCriteria
9          End If

10         oRs.Open qy
11         If Not oRs.EOF Then
12             oRs.MoveFirst
13             DLookup = oRs.Fields(sField).value
14         End If
15         oRs.Close
16         Set oRs = Nothing
17         Exit Function

' <VB WATCH>
18         Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DLookup"

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
Public Function DLookupA(ReturnColumnNo As Integer, ArrayName As Variant, FindColNo As Integer, FindValue As Variant) As Variant
' <VB WATCH>
19         On Error GoTo vbwErrHandler
' </VB WATCH>
20         Dim I As Integer

21         If FindValue = -1 Or IsNull(FindValue) Then
22             DLookupA = Empty
23             Exit Function
24         End If

25         DLookupA = 0
26         For I = 0 To UBound(ArrayName, 2)
27             If ArrayName(FindColNo, I) = FindValue Then
28                 DLookupA = ArrayName(ReturnColumnNo, I)
29                 Exit For
30             End If
31         Next I

' <VB WATCH>
32         Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DLookupA"

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
Function MotorEfficiency(KW As Single, Motor As Long, StatorFill As Long)
' <VB WATCH>
33         On Error GoTo vbwErrHandler
' </VB WATCH>
34         Dim eff0 As Single, eff1 As Single, eff2 As Single, eff3 As Single, eff4 As Single, eff5 As Single
35         Dim kw0 As Single, kw1 As Single, kw2 As Single, kw3 As Single, kw4 As Single, kw5 As Single
36         Dim qy As New ADODB.Command
37         Dim rs As New ADODB.Recordset

           'select the testsetup data for the serial number
38         qy.ActiveConnection = cnPumpData
39         If StatorFill = 1 Then  'dry stator
40             qy.CommandText = "SELECT * FROM MotorEfficiencies WHERE (((MotorEfficiencies.MotorKey)=" & Motor & ") AND ((MotorEfficiencies.Fill)='No')) OR (((MotorEfficiencies.MotorKey)=" & Motor & ") AND ((MotorEfficiencies.Fill)='Both'));"
41         Else
42             qy.CommandText = "SELECT * FROM MotorEfficiencies WHERE (((MotorEfficiencies.MotorKey)=" & Motor & ") AND ((MotorEfficiencies.Fill)='Yes')) OR (((MotorEfficiencies.MotorKey)=" & Motor & ") AND ((MotorEfficiencies.Fill)='Both'));"

43         End If

44         With rs     'open the recordset for the query
45             .CursorLocation = adUseServer
46             .CursorType = adOpenDynamic
47             .Open qy
48         End With

49         If rs.BOF = True And rs.EOF = True Then
50             MotorEfficiency = 0
51             Exit Function
52         End If

53         If rs!in125 <> 0 Then
54             kw5 = rs!in125
55             eff5 = rs!eff125
56         Else
57             kw5 = rs!in100
58             eff5 = rs!eff100
59         End If

60         kw4 = rs!in100
61         kw3 = rs!in75
62         kw2 = rs!in50
63         kw1 = rs!in25
64         kw0 = rs!in0

65         eff4 = rs!eff100
66         eff3 = rs!eff75
67         eff2 = rs!eff50
68         eff1 = rs!eff25
69         eff0 = rs!eff0

70         Select Case KW
               Case Is >= kw5
71                 MotorEfficiency = eff5      'trap at highest table entry

72             Case Is >= kw4
73                 MotorEfficiency = Interpolate(eff5, eff4, kw5, kw4, KW)

74             Case Is >= kw3
75                 MotorEfficiency = Interpolate(eff4, eff3, kw4, kw3, KW)

76             Case Is >= kw2
77                 MotorEfficiency = Interpolate(eff3, eff2, kw3, kw2, KW)

78             Case Is >= kw1
79                 MotorEfficiency = Interpolate(eff2, eff1, kw2, kw1, KW)

80             Case Is < kw1
81                 MotorEfficiency = Interpolate(eff1, eff0, kw1, kw0, KW)

82             Case Else
83                 MotorEfficiency = " "
84         End Select
' <VB WATCH>
85         Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "MotorEfficiency"

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
Function TEMCMotorEfficiency(KW As Single, ModelNumber As String, Voltage As String, RatedKW As Single)
' <VB WATCH>
86         On Error GoTo vbwErrHandler
' </VB WATCH>
87         Dim eff0 As Single, eff1 As Single, eff2 As Single, eff3 As Single, eff4 As Single
88         Dim kw0 As Single, kw1 As Single, kw2 As Single, kw3 As Single, kw4 As Single
89         Dim qy As New ADODB.Command
90         Dim rs As New ADODB.Recordset

91         If ModelNumber = "" Then
92             TEMCMotorEfficiency = 0
93             RatedKW = 999
94             Exit Function
95         End If

           'select the testsetup data for the serial number
96         qy.ActiveConnection = cnPumpData
97         qy.CommandText = "SELECT TEMCMotorEfficienciesNew.* From TEMCMotorEfficienciesNew " & _
                             "WHERE ((TEMCMotorEfficienciesNew.ModelNumber)= " & ModelNumber & _
                             ") ;"
       '        ") AND ((TEMCMotorEfficiencies.Voltage)= " & Voltage & "));"

98         With rs     'open the recordset for the query
99             .CursorLocation = adUseServer
100            .CursorType = adOpenDynamic
101            .Open qy
102        End With

103        If rs.BOF = True And rs.EOF = True Then
104            TEMCMotorEfficiency = 0
105            RatedKW = 999
106            Exit Function
107        End If

108        kw4 = rs!in100
109        kw3 = rs!in75
110        kw2 = rs!in50
111        kw1 = rs!in25
112        kw0 = rs!in0
113        eff4 = 100 * rs!eff100 / 100
114        eff3 = 100 * rs!eff75 / 100
115        eff2 = 100 * rs!eff50 / 100
116        eff1 = 100 * rs!eff25 / 100
117        eff0 = 100 * rs!eff0 / 100

118        Select Case KW
               Case Is >= kw4
119                TEMCMotorEfficiency = eff4          'trap at highest table entry

120            Case Is >= kw3
121                TEMCMotorEfficiency = Interpolate(eff4, eff3, kw4, kw3, KW)

122            Case Is >= kw2
123                TEMCMotorEfficiency = Interpolate(eff3, eff2, kw3, kw2, KW)

124            Case Is >= kw1
125                TEMCMotorEfficiency = Interpolate(eff2, eff1, kw2, kw1, KW)

126            Case Is < kw1
127                TEMCMotorEfficiency = Interpolate(eff1, eff0, kw1, kw0, KW)

128            Case Else
129                TEMCMotorEfficiency = " "
130        End Select
131        If rs!RatedOutput <> 0 Then
132            RatedKW = rs!RatedOutput
133        Else
134            RatedKW = 999
135        End If

' <VB WATCH>
136        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "TEMCMotorEfficiency"

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
Function Interpolate(HiEff, LowEff, HiKW, LowKW, ActualKW) As Single
' <VB WATCH>
137        On Error GoTo vbwErrHandler
' </VB WATCH>
138        Dim PctKw As Single

139        PctKw = (ActualKW - LowKW) / (HiKW - LowKW)
140        Interpolate = PctKw * (HiEff - LowEff) + LowEff

' <VB WATCH>
141        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Interpolate"

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
Function CalculateSuctionPressure(SuctPress, SuctInHg)
' <VB WATCH>
142        On Error GoTo vbwErrHandler
' </VB WATCH>
143        Dim sp As Single

144        If (Not IsNumeric(SuctPress)) Then
145            sp = 0
146        Else
147            sp = SuctPress
148        End If

149        CalculateSuctionPressure = sp - 0.4893 * SuctInHg
' <VB WATCH>
150        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "CalculateSuctionPressure"

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

Function CalcVelHead(Flow, DischDiam, SuctDiam)
' <VB WATCH>
151        On Error GoTo vbwErrHandler
' </VB WATCH>
152        If Not (DischDiam = 0 Or SuctDiam = 0) Then
153            If Not ((SuctDiam = -1 Or DischDiam = -1) Or DLookupA(ActualColNo, PipeDiameters, IDColNo, SuctDiam) = 0) Then
154                CalcVelHead = (0.00259 * Flow ^ 2 / DLookupA(ActualColNo, PipeDiameters, IDColNo, DischDiam) ^ 4) - (0.00259 * Flow ^ 2 / DLookupA(ActualColNo, PipeDiameters, IDColNo, SuctDiam) ^ 4)
155            End If
156        End If
' <VB WATCH>
157        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "CalcVelHead"

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

Function CalcTDH(DischargePressure, SuctionPressure, SuctionInHg, VelHead, HDCorr, SuctTemp)
' <VB WATCH>
158        On Error GoTo vbwErrHandler
' </VB WATCH>
159        If IsNull(HDCorr) Then
160            HDCorr = 0
161        End If
162        If SuctTemp < 40 Or IsNull(SuctTemp) Then
163            CalcTDH = 0
164            Exit Function
165        End If
       '    CalcTDH = (DischargePressure - CalculateSuctionPressure(SuctionPressure, SuctionInHg)) * 144 * DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(SuctTemp)) + VelHead + HDCorr
166        CalcTDH = (DischargePressure - CalculateSuctionPressure(SuctionPressure, SuctionInHg)) * 144 * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(SuctTemp)) + VelHead + HDCorr

' <VB WATCH>
167        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "CalcTDH"

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

Function FillArrays()
' <VB WATCH>
168        On Error GoTo vbwErrHandler
' </VB WATCH>

           'fill the arrays for dlookup
169        Dim rsTemp As New ADODB.Recordset

170        rsTemp.Open "PipeDiameters", cnPumpData, adOpenStatic, adLockReadOnly
171        PipeDiameters = rsTemp.GetRows()
172        rsTemp.Close
173        rsTemp.Open "VaporPressure", cnPumpData, adOpenStatic, adLockReadOnly
174        VaporPressure = rsTemp.GetRows()
175        rsTemp.Close
176        rsTemp.Open "TempCorrection", cnPumpData, adOpenStatic, adLockReadOnly
177        TempCorrection = rsTemp.GetRows()
178        rsTemp.Close
179        rsTemp.Open "TEMCForceViscosity", cnPumpData, adOpenStatic, adLockReadOnly
180        TEMCForceViscosity = rsTemp.GetRows()
181        rsTemp.Close
182        Set rsTemp = Nothing
' <VB WATCH>
183        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FillArrays"

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
Public Function PingSilent(strComputer) As Integer
' <VB WATCH>
184        On Error GoTo vbwErrHandler
' </VB WATCH>
185        Dim PID As Long
186        Dim hProcess As Long
187        Dim str As String

188        str = Environ$("comspec") & " /c ping -n 2 -w 300 " & strComputer & " | find /c ""Reply"" > """ & App.Path & "\pingdata.txt"""

189        PID = Shell(str, vbHide)


190        If PID = 0 Then
                '
                'Handle Error, Shell Didn't Work
                '
191        Else
192             hProcess = OpenProcess(&H100000, True, PID)
193             WaitForSingleObject hProcess, -1
194             CloseHandle hProcess
195        End If

196        Open App.Path & "\pingdata.txt" For Input As #1
197        Input #1, str

198        PingSilent = Val(str)

199        Close #1

' <VB WATCH>
200        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "PingSilent"

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




