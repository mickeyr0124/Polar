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
2          Const VBWPROCNAME = "AccessRoutines.DLookup"
3          If vbwProtector.vbwTraceProc Then
4              Dim vbwProtectorParameterString As String
5              If vbwProtector.vbwTraceParameters Then
6                  vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("sField", sField) & ", "
7                  vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sDomain", sDomain) & ", "
8                  vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sCriteria", sCriteria) & ") "
9              End If
10             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
11         End If
' </VB WATCH>

12         Dim oRs As New ADODB.Recordset
13         Dim qy As New ADODB.Command

14         DLookup = Empty

15         qy.ActiveConnection = cnPumpData

16         qy.CommandText = "SELECT " & sField & " FROM " & sDomain
17         If LenB(sCriteria) <> 0 Then
18             qy.CommandText = qy.CommandText & " WHERE " & sCriteria
19         End If

20         oRs.Open qy
21         If Not oRs.EOF Then
22             oRs.MoveFirst
23             DLookup = oRs.Fields(sField).value
24         End If
25         oRs.Close
26         Set oRs = Nothing
' <VB WATCH>
27         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
28         Exit Function

' <VB WATCH>
29         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
30         Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "sField", sField
            vbwReportVariable "sDomain", sDomain
            vbwReportVariable "sCriteria", sCriteria
            vbwReportVariable "oRs", oRs
            vbwReportVariable "qy", qy
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Public Function DLookupA(ReturnColumnNo As Integer, ArrayName As Variant, FindColNo As Integer, FindValue As Variant) As Variant
' <VB WATCH>
31         On Error GoTo vbwErrHandler
32         Const VBWPROCNAME = "AccessRoutines.DLookupA"
33         If vbwProtector.vbwTraceProc Then
34             Dim vbwProtectorParameterString As String
35             If vbwProtector.vbwTraceParameters Then
36                 vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ReturnColumnNo", ReturnColumnNo) & ", "
37                 vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ArrayName", ArrayName) & ", "
38                 vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("FindColNo", FindColNo) & ", "
39                 vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("FindValue", FindValue) & ") "
40             End If
41             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
42         End If
' </VB WATCH>
43         Dim I As Integer

44         If FindValue = -1 Or IsNull(FindValue) Then
45             DLookupA = Empty
' <VB WATCH>
46         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
47             Exit Function
48         End If

49         DLookupA = 0
50         For I = 0 To UBound(ArrayName, 2)
51             If ArrayName(FindColNo, I) = FindValue Then
52                 DLookupA = ArrayName(ReturnColumnNo, I)
53                 Exit For
54             End If
55         Next I

' <VB WATCH>
56         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
57         Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ReturnColumnNo", ReturnColumnNo
            vbwReportVariable "ArrayName", ArrayName
            vbwReportVariable "FindColNo", FindColNo
            vbwReportVariable "FindValue", FindValue
            vbwReportVariable "I", I
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Function MotorEfficiency(KW As Single, Motor As Long, StatorFill As Long)
' <VB WATCH>
58         On Error GoTo vbwErrHandler
59         Const VBWPROCNAME = "AccessRoutines.MotorEfficiency"
60         If vbwProtector.vbwTraceProc Then
61             Dim vbwProtectorParameterString As String
62             If vbwProtector.vbwTraceParameters Then
63                 vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("KW", KW) & ", "
64                 vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("Motor", Motor) & ", "
65                 vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("StatorFill", StatorFill) & ") "
66             End If
67             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
68         End If
' </VB WATCH>
69         Dim eff0 As Single, eff1 As Single, eff2 As Single, eff3 As Single, eff4 As Single, eff5 As Single
70         Dim kw0 As Single, kw1 As Single, kw2 As Single, kw3 As Single, kw4 As Single, kw5 As Single
71         Dim qy As New ADODB.Command
72         Dim rs As New ADODB.Recordset

           'select the testsetup data for the serial number
73         qy.ActiveConnection = cnPumpData
74         If StatorFill = 1 Then  'dry stator
75             qy.CommandText = "SELECT * FROM MotorEfficiencies WHERE (((MotorEfficiencies.MotorKey)=" & Motor & ") AND ((MotorEfficiencies.Fill)='No')) OR (((MotorEfficiencies.MotorKey)=" & Motor & ") AND ((MotorEfficiencies.Fill)='Both'));"
76         Else
77             qy.CommandText = "SELECT * FROM MotorEfficiencies WHERE (((MotorEfficiencies.MotorKey)=" & Motor & ") AND ((MotorEfficiencies.Fill)='Yes')) OR (((MotorEfficiencies.MotorKey)=" & Motor & ") AND ((MotorEfficiencies.Fill)='Both'));"

78         End If

79         With rs     'open the recordset for the query
80             .CursorLocation = adUseServer
81             .CursorType = adOpenDynamic
82             .Open qy
83         End With

84         If rs.BOF = True And rs.EOF = True Then
85             MotorEfficiency = 0
' <VB WATCH>
86         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
87             Exit Function
88         End If

89         If rs!in125 <> 0 Then
90             kw5 = rs!in125
91             eff5 = rs!eff125
92         Else
93             kw5 = rs!in100
94             eff5 = rs!eff100
95         End If

96         kw4 = rs!in100
97         kw3 = rs!in75
98         kw2 = rs!in50
99         kw1 = rs!in25
100        kw0 = rs!in0

101        eff4 = rs!eff100
102        eff3 = rs!eff75
103        eff2 = rs!eff50
104        eff1 = rs!eff25
105        eff0 = rs!eff0

106        Select Case KW
               Case Is >= kw5
107                MotorEfficiency = eff5      'trap at highest table entry

108            Case Is >= kw4
109                MotorEfficiency = Interpolate(eff5, eff4, kw5, kw4, KW)

110            Case Is >= kw3
111                MotorEfficiency = Interpolate(eff4, eff3, kw4, kw3, KW)

112            Case Is >= kw2
113                MotorEfficiency = Interpolate(eff3, eff2, kw3, kw2, KW)

114            Case Is >= kw1
115                MotorEfficiency = Interpolate(eff2, eff1, kw2, kw1, KW)

116            Case Is < kw1
117                MotorEfficiency = Interpolate(eff1, eff0, kw1, kw0, KW)

118            Case Else
119                MotorEfficiency = " "
120        End Select
' <VB WATCH>
121        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
122        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "KW", KW
            vbwReportVariable "Motor", Motor
            vbwReportVariable "StatorFill", StatorFill
            vbwReportVariable "eff0", eff0
            vbwReportVariable "eff1", eff1
            vbwReportVariable "eff2", eff2
            vbwReportVariable "eff3", eff3
            vbwReportVariable "eff4", eff4
            vbwReportVariable "eff5", eff5
            vbwReportVariable "kw0", kw0
            vbwReportVariable "kw1", kw1
            vbwReportVariable "kw2", kw2
            vbwReportVariable "kw3", kw3
            vbwReportVariable "kw4", kw4
            vbwReportVariable "kw5", kw5
            vbwReportVariable "qy", qy
            vbwReportVariable "rs", rs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Function TEMCMotorEfficiency(KW As Single, ModelNumber As String, Voltage As String, RatedKW As Single)
' <VB WATCH>
123        On Error GoTo vbwErrHandler
124        Const VBWPROCNAME = "AccessRoutines.TEMCMotorEfficiency"
125        If vbwProtector.vbwTraceProc Then
126            Dim vbwProtectorParameterString As String
127            If vbwProtector.vbwTraceParameters Then
128                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("KW", KW) & ", "
129                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ModelNumber", ModelNumber) & ", "
130                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("Voltage", Voltage) & ", "
131                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("RatedKW", RatedKW) & ") "
132            End If
133            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
134        End If
' </VB WATCH>
135        Dim eff0 As Single, eff1 As Single, eff2 As Single, eff3 As Single, eff4 As Single
136        Dim kw0 As Single, kw1 As Single, kw2 As Single, kw3 As Single, kw4 As Single
137        Dim qy As New ADODB.Command
138        Dim rs As New ADODB.Recordset

139        If ModelNumber = "" Then
140            TEMCMotorEfficiency = 0
141            RatedKW = 999
' <VB WATCH>
142        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
143            Exit Function
144        End If

           'select the testsetup data for the serial number
145        qy.ActiveConnection = cnPumpData
146        qy.CommandText = "SELECT TEMCMotorEfficienciesNew.* From TEMCMotorEfficienciesNew " & _
                             "WHERE ((TEMCMotorEfficienciesNew.ModelNumber)= " & ModelNumber & _
                             ") ;"
       '        ") AND ((TEMCMotorEfficiencies.Voltage)= " & Voltage & "));"

147        With rs     'open the recordset for the query
148            .CursorLocation = adUseServer
149            .CursorType = adOpenDynamic
150            .Open qy
151        End With

152        If rs.BOF = True And rs.EOF = True Then
153            TEMCMotorEfficiency = 0
154            RatedKW = 999
' <VB WATCH>
155        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
156            Exit Function
157        End If

158        kw4 = rs!in100
159        kw3 = rs!in75
160        kw2 = rs!in50
161        kw1 = rs!in25
162        kw0 = rs!in0
163        eff4 = 100 * rs!eff100 / 100
164        eff3 = 100 * rs!eff75 / 100
165        eff2 = 100 * rs!eff50 / 100
166        eff1 = 100 * rs!eff25 / 100
167        eff0 = 100 * rs!eff0 / 100

168        Select Case KW
               Case Is >= kw4
169                TEMCMotorEfficiency = eff4          'trap at highest table entry

170            Case Is >= kw3
171                TEMCMotorEfficiency = Interpolate(eff4, eff3, kw4, kw3, KW)

172            Case Is >= kw2
173                TEMCMotorEfficiency = Interpolate(eff3, eff2, kw3, kw2, KW)

174            Case Is >= kw1
175                TEMCMotorEfficiency = Interpolate(eff2, eff1, kw2, kw1, KW)

176            Case Is < kw1
177                TEMCMotorEfficiency = Interpolate(eff1, eff0, kw1, kw0, KW)

178            Case Else
179                TEMCMotorEfficiency = " "
180        End Select
181        If rs!RatedOutput <> 0 Then
182            RatedKW = rs!RatedOutput
183        Else
184            RatedKW = 999
185        End If

' <VB WATCH>
186        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
187        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "KW", KW
            vbwReportVariable "ModelNumber", ModelNumber
            vbwReportVariable "Voltage", Voltage
            vbwReportVariable "RatedKW", RatedKW
            vbwReportVariable "eff0", eff0
            vbwReportVariable "eff1", eff1
            vbwReportVariable "eff2", eff2
            vbwReportVariable "eff3", eff3
            vbwReportVariable "eff4", eff4
            vbwReportVariable "kw0", kw0
            vbwReportVariable "kw1", kw1
            vbwReportVariable "kw2", kw2
            vbwReportVariable "kw3", kw3
            vbwReportVariable "kw4", kw4
            vbwReportVariable "qy", qy
            vbwReportVariable "rs", rs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Function Interpolate(HiEff, LowEff, HiKW, LowKW, ActualKW) As Single
' <VB WATCH>
188        On Error GoTo vbwErrHandler
189        Const VBWPROCNAME = "AccessRoutines.Interpolate"
190        If vbwProtector.vbwTraceProc Then
191            Dim vbwProtectorParameterString As String
192            If vbwProtector.vbwTraceParameters Then
193                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("HiEff", HiEff) & ", "
194                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("LowEff", LowEff) & ", "
195                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("HiKW", HiKW) & ", "
196                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("LowKW", LowKW) & ", "
197                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ActualKW", ActualKW) & ") "
198            End If
199            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
200        End If
' </VB WATCH>
201        Dim PctKw As Single

202        PctKw = (ActualKW - LowKW) / (HiKW - LowKW)
203        Interpolate = PctKw * (HiEff - LowEff) + LowEff

' <VB WATCH>
204        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
205        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "HiEff", HiEff
            vbwReportVariable "LowEff", LowEff
            vbwReportVariable "HiKW", HiKW
            vbwReportVariable "LowKW", LowKW
            vbwReportVariable "ActualKW", ActualKW
            vbwReportVariable "PctKw", PctKw
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Function CalculateSuctionPressure(SuctPress, SuctInHg)
' <VB WATCH>
206        On Error GoTo vbwErrHandler
207        Const VBWPROCNAME = "AccessRoutines.CalculateSuctionPressure"
208        If vbwProtector.vbwTraceProc Then
209            Dim vbwProtectorParameterString As String
210            If vbwProtector.vbwTraceParameters Then
211                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("SuctPress", SuctPress) & ", "
212                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("SuctInHg", SuctInHg) & ") "
213            End If
214            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
215        End If
' </VB WATCH>
216        Dim sp As Single

217        If (Not IsNumeric(SuctPress)) Then
218            sp = 0
219        Else
220            sp = SuctPress
221        End If

222        CalculateSuctionPressure = sp - 0.4893 * SuctInHg
' <VB WATCH>
223        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
224        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "SuctPress", SuctPress
            vbwReportVariable "SuctInHg", SuctInHg
            vbwReportVariable "sp", sp
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function CalcVelHead(Flow, DischDiam, SuctDiam)
' <VB WATCH>
225        On Error GoTo vbwErrHandler
226        Const VBWPROCNAME = "AccessRoutines.CalcVelHead"
227        If vbwProtector.vbwTraceProc Then
228            Dim vbwProtectorParameterString As String
229            If vbwProtector.vbwTraceParameters Then
230                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Flow", Flow) & ", "
231                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("DischDiam", DischDiam) & ", "
232                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("SuctDiam", SuctDiam) & ") "
233            End If
234            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
235        End If
' </VB WATCH>
236        If Not (DischDiam = 0 Or SuctDiam = 0) Then
237            If Not ((SuctDiam = -1 Or DischDiam = -1) Or DLookupA(ActualColNo, PipeDiameters, IDColNo, SuctDiam) = 0) Then
238                CalcVelHead = (0.00259 * Flow ^ 2 / DLookupA(ActualColNo, PipeDiameters, IDColNo, DischDiam) ^ 4) - (0.00259 * Flow ^ 2 / DLookupA(ActualColNo, PipeDiameters, IDColNo, SuctDiam) ^ 4)
239            End If
240        End If
' <VB WATCH>
241        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
242        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Flow", Flow
            vbwReportVariable "DischDiam", DischDiam
            vbwReportVariable "SuctDiam", SuctDiam
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function CalcTDH(DischargePressure, SuctionPressure, SuctionInHg, VelHead, HDCorr, SuctTemp)
' <VB WATCH>
243        On Error GoTo vbwErrHandler
244        Const VBWPROCNAME = "AccessRoutines.CalcTDH"
245        If vbwProtector.vbwTraceProc Then
246            Dim vbwProtectorParameterString As String
247            If vbwProtector.vbwTraceParameters Then
248                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("DischargePressure", DischargePressure) & ", "
249                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("SuctionPressure", SuctionPressure) & ", "
250                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("SuctionInHg", SuctionInHg) & ", "
251                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("VelHead", VelHead) & ", "
252                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("HDCorr", HDCorr) & ", "
253                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("SuctTemp", SuctTemp) & ") "
254            End If
255            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
256        End If
' </VB WATCH>
257        If IsNull(HDCorr) Then
258            HDCorr = 0
259        End If
260        If SuctTemp < 40 Or IsNull(SuctTemp) Then
261            CalcTDH = 0
' <VB WATCH>
262        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
263            Exit Function
264        End If
       '    CalcTDH = (DischargePressure - CalculateSuctionPressure(SuctionPressure, SuctionInHg)) * 144 * DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(SuctTemp)) + VelHead + HDCorr
265        CalcTDH = (DischargePressure - CalculateSuctionPressure(SuctionPressure, SuctionInHg)) * 144 * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(SuctTemp)) + VelHead + HDCorr

' <VB WATCH>
266        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
267        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "DischargePressure", DischargePressure
            vbwReportVariable "SuctionPressure", SuctionPressure
            vbwReportVariable "SuctionInHg", SuctionInHg
            vbwReportVariable "VelHead", VelHead
            vbwReportVariable "HDCorr", HDCorr
            vbwReportVariable "SuctTemp", SuctTemp
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function FillArrays()
' <VB WATCH>
268        On Error GoTo vbwErrHandler
269        Const VBWPROCNAME = "AccessRoutines.FillArrays"
270        If vbwProtector.vbwTraceProc Then
271            Dim vbwProtectorParameterString As String
272            If vbwProtector.vbwTraceParameters Then
273                vbwProtectorParameterString = "()"
274            End If
275            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
276        End If
' </VB WATCH>

           'fill the arrays for dlookup
277        Dim rsTemp As New ADODB.Recordset

278        rsTemp.Open "PipeDiameters", cnPumpData, adOpenStatic, adLockReadOnly
279        PipeDiameters = rsTemp.GetRows()
280        rsTemp.Close
281        rsTemp.Open "VaporPressure", cnPumpData, adOpenStatic, adLockReadOnly
282        VaporPressure = rsTemp.GetRows()
283        rsTemp.Close
284        rsTemp.Open "TempCorrection", cnPumpData, adOpenStatic, adLockReadOnly
285        TempCorrection = rsTemp.GetRows()
286        rsTemp.Close
287        rsTemp.Open "TEMCForceViscosity", cnPumpData, adOpenStatic, adLockReadOnly
288        TEMCForceViscosity = rsTemp.GetRows()
289        rsTemp.Close
290        Set rsTemp = Nothing
' <VB WATCH>
291        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
292        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "rsTemp", rsTemp
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Public Function PingSilent(strComputer) As Integer
' <VB WATCH>
293        On Error GoTo vbwErrHandler
294        Const VBWPROCNAME = "AccessRoutines.PingSilent"
295        If vbwProtector.vbwTraceProc Then
296            Dim vbwProtectorParameterString As String
297            If vbwProtector.vbwTraceParameters Then
298                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("strComputer", strComputer) & ") "
299            End If
300            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
301        End If
' </VB WATCH>
302        Dim PID As Long
303        Dim hProcess As Long
304        Dim str As String

305        str = Environ$("comspec") & " /c ping -n 2 -w 300 " & strComputer & " | find /c ""Reply"" > """ & App.Path & "\pingdata.txt"""

306        PID = Shell(str, vbHide)


307        If PID = 0 Then
                '
                'Handle Error, Shell Didn't Work
                '
308        Else
309             hProcess = OpenProcess(&H100000, True, PID)
310             WaitForSingleObject hProcess, -1
311             CloseHandle hProcess
312        End If

313        Open App.Path & "\pingdata.txt" For Input As #1
314        Input #1, str

315        PingSilent = Val(str)

316        Close #1

' <VB WATCH>
317        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
318        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "strComputer", strComputer
            vbwReportVariable "PID", PID
            vbwReportVariable "hProcess", hProcess
            vbwReportVariable "str", str
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function




' <VB WATCH> <VBWATCHFINALPROC>
' Procedures added by VB Watch for variable dump


Private Sub vbwReportModuleVariables()
    vbwReportToFile VBW_MODULE_STRING
End Sub
' </VB WATCH>
