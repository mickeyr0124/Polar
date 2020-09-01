Attribute VB_Name = "EpicorRoutines"
    Option Explicit
    Public Type SNRecord
        SONumber As String
        SOLine As String
        ModelNo As String
        MotorSize As String
        PartNum As String
        Customer As String
        ShipTo As String
        CustNum As String
        ShipToNum As String
        TDH As String
        Flow As String
        ImpellerDiameter As String
        SuctionPressure As String
        SpGr As String
        Fluid As String
        PumpTemperature As String
        Viscosity As String
        VaporPressure As String
        SuctFlangeSize As String
        DischFlangeSize As String
        RPM As String
        Voltage As String
        StatorFill As String
        CirculationPath As String
        TestProcedure As String
        DesignPressure As String
        Frequency As String
        XPartNum As String
        '
        Phases As String
        NPSHr As String
        RatedInputPower As String
        FLCurrent As String
        ThermalClass As String
        ExpClass As String
        LiquidTemp As String
        JobNumber As String
        CustomerPO As String
    End Type

' <VB WATCH>
Const VBWMODULE = "EpicorRoutines"
' </VB WATCH>

Public Function GetEpicorODBCData(SerialNumber As String, EpicorConnectionString As String) As SNRecord
' <VB WATCH>
1          On Error GoTo vbwErrHandler
2          Const VBWPROCNAME = "EpicorRoutines.GetEpicorODBCData"
3          If vbwProtector.vbwTraceProc Then
4              Dim vbwProtectorParameterString As String
5              If vbwProtector.vbwTraceParameters Then
6                  vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("SerialNumber", SerialNumber) & ", "
7                  vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("EpicorConnectionString", EpicorConnectionString) & ") "
8              End If
9              vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
10         End If
' </VB WATCH>
11         Dim conConn As New ADODB.Connection
12         Dim cmdCommand As New ADODB.Command
13         Dim rstRecordSet As New ADODB.Recordset
14         Dim SQLString As String

15         Dim MyRecord As SNRecord

           'construct connection string
16         conConn.Open EpicorConnectionString
       '    conConn.Open "Driver={SQL Server};" & _
'                    "Database=Epicor905;" & _
'                    "Server=ERP-DB-01;" & _
'                    "UID=MRosenbaum;" & _
'                    "PWD=Maple;"

       '   first see if there is an order number in the job file.  if there is, it is a make direct
       '       job and we bring back all of the information from Epicor as normal.
       '       if there is no order number, it is a make to stock job (supermarket), and we want
       '       to return the job number and part number only.  there is a table in the database
       '       that will get referenced for the supermarket data to put into temppumpdata

17         SQLString = "SELECT"
18         SQLString = SQLString & " SerialNo.JobNum       AS JobNum,"
19         SQLString = SQLString & " SerialNo.PartNum    AS PartNum,"
20         SQLString = SQLString & " SerialNo.SerialNumber AS SerialNo,"
21         SQLString = SQLString & " JobProd.OrderNum     AS SONumber, "
22         SQLString = SQLString & " JobProd.JobNum AS  JobProdJobNum "
23         SQLString = SQLString & " FROM Erp.SerialNo, Erp.JobProd "
24         SQLString = SQLString & " WHERE SerialNo.SerialNumber = '" & SerialNumber & "' "
25         SQLString = SQLString & " AND JobProd.JobNum = SerialNo.JobNum "
26         SQLString = SQLString & ";"

27         With cmdCommand
28             .ActiveConnection = conConn
29             .CommandText = SQLString
30             .CommandType = adCmdText
31         End With

32         With rstRecordSet
33            .CursorType = adOpenStatic
34            .CursorLocation = adUseClient
35            .LockType = adLockBatchOptimistic
36            .Open cmdCommand
37          End With

           'if we have a record, save the data, else tell user and leave
38         If rstRecordSet.RecordCount > 0 Then    'there is no order no
39             If rstRecordSet.Fields("SONumber") = 0 Then
40                 rstRecordSet.MoveFirst
41                 MyRecord.PartNum = rstRecordSet.Fields("PartNum")
42                 MyRecord.JobNumber = rstRecordSet.Fields("Jobnum")
43                 MyRecord.SONumber = 0
                   'close the recordset and connection
44                 rstRecordSet.Close
45                 conConn.Close

46                 Set rstRecordSet = Nothing
47                 Set conConn = Nothing

48                 GetEpicorODBCData = MyRecord
' <VB WATCH>
49         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
50                 Exit Function
51             End If
52         End If

           'get job number, order number, order line and misc data from serial number  and order detail tables
53         SQLString = "SELECT"
54         SQLString = SQLString & " SerialNo.JobNum       AS JobNum,"
55         SQLString = SQLString & " SerialNo.PartNum    AS PartNum,"
56         SQLString = SQLString & " SerialNo.SerialNumber AS SerialNo,"
57         SQLString = SQLString & " JobProd.OrderNum     AS SONumber,"
58         SQLString = SQLString & " JobProd.OrderLine    AS SOLine,"
           'SQLString = SQLString & " OrderDtl.Character01  AS ModelNo, "
59         SQLString = SQLString & " OrderDtl.XPartNum  AS XPartNum, "
60         SQLString = SQLString & " OrderHed.CustNum  AS CustNum, "
61         SQLString = SQLString & " OrderHed.ShiptoNum AS ShipToNum, "
62         SQLString = SQLString & " OrderHed.PONum AS CustPONum, "
63         SQLString = SQLString & " Customer.Name AS CustomerName,  "
64         SQLString = SQLString & " ShipTo.Name AS ShipToName  "
65         SQLString = SQLString & " FROM Erp.OrderDtl AS OrderDtl, Erp.SerialNo, Erp.JobProd, Erp.OrderHed, Erp.Customer, Erp.ShipTo "
66         SQLString = SQLString & " WHERE SerialNo.SerialNumber = '" & SerialNumber & "' "
67         SQLString = SQLString & " AND JobProd.JobNum = SerialNo.JobNum "
68         SQLString = SQLString & " AND OrderDtl.OrderNum = JobProd.OrderNum"
69         SQLString = SQLString & " AND OrderDtl.OrderLine = JobProd.OrderLine"
70         SQLString = SQLString & " AND OrderHed.OrderNum = JobProd.OrderNum"
71         SQLString = SQLString & " AND Customer.CustNum = OrderHed.CustNum"
72         SQLString = SQLString & " AND ShipTo.ShipToNum = OrderHed.ShipToNum"
73         SQLString = SQLString & ";"

74         With cmdCommand
75             .ActiveConnection = conConn
76             .CommandText = SQLString
77             .CommandType = adCmdText
78         End With

79         With rstRecordSet
80             If rstRecordSet.State = adStateOpen Then
81                 .Close
82             End If
83            .CursorType = adOpenStatic
84            .CursorLocation = adUseClient
85            .LockType = adLockBatchOptimistic
86            .Open cmdCommand
87          End With

           'if we have a record, save the data, else tell user and leave
88         If rstRecordSet.RecordCount > 0 Then
89             rstRecordSet.MoveFirst
90             MyRecord.SONumber = rstRecordSet.Fields("SONumber")
91             MyRecord.SOLine = rstRecordSet.Fields("SOLine")
       '        MyRecord.ModelNo = rstRecordSet.Fields("ModelNo")
92             MyRecord.PartNum = rstRecordSet.Fields("PartNum")
93             MyRecord.CustNum = rstRecordSet.Fields("CustNum")
94             MyRecord.CustomerPO = rstRecordSet.Fields("CustPONum")
95             MyRecord.ShipToNum = rstRecordSet.Fields("ShipToNum")
96             MyRecord.JobNumber = rstRecordSet.Fields("Jobnum")
97             MyRecord.Customer = rstRecordSet.Fields("CustomerName")
98             MyRecord.XPartNum = rstRecordSet.Fields("XPartNum")
99             MyRecord.ShipTo = IIf(MyRecord.ShipToNum = "", rstRecordSet.Fields("CustomerName"), rstRecordSet.Fields("ShipToName"))
100        Else
101            MsgBox ("No Records found for Serial Number = " & SerialNumber)
' <VB WATCH>
102        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
103            Exit Function
104        End If

           'get ud02 data
105        SQLString = "SELECT"
106        SQLString = SQLString & " UD02.Number01         AS TDH,"
107        SQLString = SQLString & " UD02.Number02         AS Flow,"
108        SQLString = SQLString & " UD02.Number07         AS ImpellerDiameter,"
109        SQLString = SQLString & " UD02.Number03         AS SuctionPressure,"
110        SQLString = SQLString & " UD02.Number17         AS DesignPressure,"
111        SQLString = SQLString & " UD02.Number14         AS NPSHr"
112        SQLString = SQLString & " FROM Ice.UD02"
113        SQLString = SQLString & " WHERE UD02.Key1 = '" & MyRecord.SONumber & "' "
114        SQLString = SQLString & " AND UD02.Key2 = '" & MyRecord.SOLine & "' "

115        With rstRecordSet
116            .Close
117            cmdCommand.CommandText = SQLString
118            .Open cmdCommand
119        End With

120        If rstRecordSet.RecordCount > 0 Then
121            rstRecordSet.MoveFirst
122            MyRecord.TDH = rstRecordSet.Fields("TDH")
123            MyRecord.Flow = rstRecordSet.Fields("Flow")
124            MyRecord.ImpellerDiameter = rstRecordSet.Fields("ImpellerDiameter")
125            MyRecord.SuctionPressure = rstRecordSet.Fields("SuctionPressure")
126            MyRecord.DesignPressure = rstRecordSet.Fields("DesignPressure")
127            MyRecord.NPSHr = rstRecordSet.Fields("NPSHr")
128        End If

           'get ud03 data
129        SQLString = "SELECT"
130        SQLString = SQLString & " UD03.Number09         AS SpGr,"
131        SQLString = SQLString & " UD03.Character02      AS Fluid,"
132        SQLString = SQLString & " UD03.Number07         AS PumpTemperature,"
133        SQLString = SQLString & " UD03.Number11         AS Viscosity,"
134        SQLString = SQLString & " UD03.Number13         AS VaporPressure,"
135        SQLString = SQLString & " UD03.Number07           As LiquidTemp"
136        SQLString = SQLString & " FROM ice.UD03"
137        SQLString = SQLString & " WHERE UD03.Key1 = '" & MyRecord.SONumber & "' "
138        SQLString = SQLString & " AND UD03.Key2 = '" & MyRecord.SOLine & "' "

139        With rstRecordSet
140            .Close
141            cmdCommand.CommandText = SQLString
142            .Open cmdCommand
143        End With

144        If rstRecordSet.RecordCount > 0 Then
145            rstRecordSet.MoveFirst
146            MyRecord.SpGr = rstRecordSet.Fields("SpGr")
147            MyRecord.Fluid = rstRecordSet.Fields("Fluid")
148            MyRecord.PumpTemperature = rstRecordSet.Fields("PumpTemperature")
149            MyRecord.Viscosity = rstRecordSet.Fields("Viscosity")
150            MyRecord.VaporPressure = rstRecordSet.Fields("VaporPressure")
151            MyRecord.LiquidTemp = rstRecordSet.Fields("LiquidTemp")
152        End If

           'get ud04 data
153        SQLString = "SELECT"
154        SQLString = SQLString & " UD04.Character01      AS SuctFlangeSize,"
155        SQLString = SQLString & " UD04.Character04      AS DischFlangeSize"
156        SQLString = SQLString & " FROM ice.UD04"
157        SQLString = SQLString & " WHERE UD04.Key1 = '" & MyRecord.SONumber & "' "
158        SQLString = SQLString & " AND UD04.Key2 = '" & MyRecord.SOLine & "' "

159        With rstRecordSet
160            .Close
161            cmdCommand.CommandText = SQLString
162            .Open cmdCommand
163        End With

164        If rstRecordSet.RecordCount > 0 Then
165            rstRecordSet.MoveFirst
166            MyRecord.SuctFlangeSize = rstRecordSet.Fields("SuctFlangeSize")
167            MyRecord.DischFlangeSize = rstRecordSet.Fields("DischFlangeSize")
168        End If

           'get ud05 data
169        SQLString = "SELECT"
170        SQLString = SQLString & " UD05.Character05      AS RPM,"
171        SQLString = SQLString & " UD05.Character01      AS Voltage,"
172        SQLString = SQLString & " UD05.Character08      AS StatorFill,"
173        SQLString = SQLString & " UD05.Character02      AS Frequency,"
174        SQLString = SQLString & " UD05.Character03      AS Phases,"
175        SQLString = SQLString & " UD05.Character06        As ThermalClass,"
176        SQLString = SQLString & " UD05.Number01         AS RatedInputPower,"
177        SQLString = SQLString & " UD05.Number02         AS FLCurrent"
178        SQLString = SQLString & " FROM ice.UD05"
179        SQLString = SQLString & " WHERE UD05.Key1 = '" & MyRecord.SONumber & "' "
180        SQLString = SQLString & " AND UD05.Key2 = '" & MyRecord.SOLine & "' "

181        With rstRecordSet
182            .Close
183            cmdCommand.CommandText = SQLString
184            .Open cmdCommand
185        End With

186        If rstRecordSet.RecordCount > 0 Then
187            rstRecordSet.MoveFirst
188            MyRecord.RPM = rstRecordSet.Fields("RPM")
189            MyRecord.Voltage = rstRecordSet.Fields("Voltage")
190            MyRecord.StatorFill = rstRecordSet.Fields("StatorFill")
191            MyRecord.Frequency = rstRecordSet.Fields("Frequency")
192            MyRecord.Phases = rstRecordSet.Fields("Phases")
193            MyRecord.RatedInputPower = rstRecordSet.Fields("RatedInputPower")
194            MyRecord.FLCurrent = rstRecordSet.Fields("FLCurrent")
195            MyRecord.ThermalClass = rstRecordSet.Fields("ThermalClass")
196        End If

           'get ud07 data
197        SQLString = "SELECT"
198        SQLString = SQLString & " UD07.Character01      AS CirculationPath"
199        SQLString = SQLString & " FROM ice.UD07"
200        SQLString = SQLString & " WHERE UD07.Key1 = '" & MyRecord.SONumber & "' "
201        SQLString = SQLString & " AND UD07.Key2 = '" & MyRecord.SOLine & "' "

202        With rstRecordSet
203            .Close
204            cmdCommand.CommandText = SQLString
205            .Open cmdCommand
206        End With

207        If rstRecordSet.RecordCount > 0 Then
208            rstRecordSet.MoveFirst
209            MyRecord.CirculationPath = rstRecordSet.Fields("CirculationPath")
210        End If

           'get ud08 data
211        SQLString = "SELECT"
212        SQLString = SQLString & " UD08.ShortChar01      AS EXPRating"
213        SQLString = SQLString & " FROM ice.UD08"
214        SQLString = SQLString & " WHERE UD08.Key1 = '" & MyRecord.SONumber & "' "
215        SQLString = SQLString & " AND UD08.Key2 = '" & MyRecord.SOLine & "' "

216        With rstRecordSet
217            .Close
218            cmdCommand.CommandText = SQLString
219            .Open cmdCommand
220        End With


221        If rstRecordSet.RecordCount > 0 Then
222            rstRecordSet.MoveFirst
223            MyRecord.ExpClass = rstRecordSet.Fields("EXPRating")
224        End If

           'get ud09 data
225        SQLString = "SELECT"
226        SQLString = SQLString & " UD09.Character01      AS TestProcedure"
227        SQLString = SQLString & " FROM ice.UD09"
228        SQLString = SQLString & " WHERE UD09.Key1 = '" & MyRecord.SONumber & "' "
229        SQLString = SQLString & " AND UD09.Key2 = '" & MyRecord.SOLine & "' "

230        With rstRecordSet
231            .Close
232            cmdCommand.CommandText = SQLString
233            .Open cmdCommand
234        End With

235        If rstRecordSet.RecordCount > 0 Then
236            rstRecordSet.MoveFirst
237            MyRecord.TestProcedure = rstRecordSet.Fields("TestProcedure")
238        End If

           'get part data
239        SQLString = "SELECT"
           'SQLString = SQLString & " Part.Character06      AS MotorSize"
240        SQLString = SQLString & " FROM erp.Part"
241        SQLString = SQLString & " WHERE Part.PartNum = '" & MyRecord.PartNum & "' "

242        With rstRecordSet
        '       .Close
        '       cmdCommand.CommandText = SQLString
        '       .Open cmdCommand
243        End With

       '    If rstRecordSet.RecordCount > 0 Then
       '        rstRecordSet.MoveFirst
       '        MyRecord.MotorSize = rstRecordSet.Fields("MotorSize")
       '    End If

           'get Customer
244        SQLString = "SELECT"
245        SQLString = SQLString & " Customer.Name      AS Customer"
246        SQLString = SQLString & " FROM Customer"
247        SQLString = SQLString & " WHERE Customer.CustNum = '" & MyRecord.CustNum & "' "

248        With rstRecordSet
249            .Close
250            cmdCommand.CommandText = SQLString
251            .Open cmdCommand
252        End With

253        If rstRecordSet.RecordCount > 0 Then
254            rstRecordSet.MoveFirst
255            MyRecord.Customer = rstRecordSet.Fields("Customer")
256        End If

           'get ShipTo
257        If MyRecord.ShipToNum <> "" Then
258            SQLString = "SELECT"
259            SQLString = SQLString & " ShipTo.Name      AS ShipTo"
260            SQLString = SQLString & " FROM ShipTo"
261            SQLString = SQLString & " WHERE ShipTo.ShipToNum = '" & MyRecord.ShipToNum & "' "

262            With rstRecordSet
263                .Close
264                cmdCommand.CommandText = SQLString
265                .Open cmdCommand
266            End With

267            If rstRecordSet.RecordCount > 0 Then
268                rstRecordSet.MoveFirst
269                MyRecord.ShipTo = rstRecordSet.Fields("ShipTo")
270            End If
271        End If

           'close the recordset and connection
272        rstRecordSet.Close
273        conConn.Close

274        Set rstRecordSet = Nothing
275        Set conConn = Nothing

276        GetEpicorODBCData = MyRecord
' <VB WATCH>
277        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
278        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetEpicorODBCData"

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
            vbwReportVariable "SerialNumber", SerialNumber
            vbwReportVariable "EpicorConnectionString", EpicorConnectionString
            vbwReportVariable "SQLString", SQLString
            vbwReport_EpicorRoutines_SNRecord "MyRecord", MyRecord
            vbwReportVariable "conConn", conConn
            vbwReportVariable "cmdCommand", cmdCommand
            vbwReportVariable "rstRecordSet", rstRecordSet
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
