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
' </VB WATCH>
2          Dim conConn As New ADODB.Connection
3          Dim cmdCommand As New ADODB.Command
4          Dim rstRecordSet As New ADODB.Recordset
5          Dim SQLString As String

6          Dim MyRecord As SNRecord

           'construct connection string
7          conConn.Open EpicorConnectionString
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

8          SQLString = "SELECT"
9          SQLString = SQLString & " SerialNo.JobNum       AS JobNum,"
10         SQLString = SQLString & " SerialNo.PartNum    AS PartNum,"
11         SQLString = SQLString & " SerialNo.SerialNumber AS SerialNo,"
12         SQLString = SQLString & " JobProd.OrderNum     AS SONumber, "
13         SQLString = SQLString & " JobProd.JobNum AS  JobProdJobNum "
14         SQLString = SQLString & " FROM Erp.SerialNo, Erp.JobProd "
15         SQLString = SQLString & " WHERE SerialNo.SerialNumber = '" & SerialNumber & "' "
16         SQLString = SQLString & " AND JobProd.JobNum = SerialNo.JobNum "
17         SQLString = SQLString & ";"

18         With cmdCommand
19             .ActiveConnection = conConn
20             .CommandText = SQLString
21             .CommandType = adCmdText
22         End With

23         With rstRecordSet
24            .CursorType = adOpenStatic
25            .CursorLocation = adUseClient
26            .LockType = adLockBatchOptimistic
27            .Open cmdCommand
28          End With

           'if we have a record, save the data, else tell user and leave
29         If rstRecordSet.RecordCount > 0 Then    'there is no order no
30             If rstRecordSet.Fields("SONumber") = 0 Then
31                 rstRecordSet.MoveFirst
32                 MyRecord.PartNum = rstRecordSet.Fields("PartNum")
33                 MyRecord.JobNumber = rstRecordSet.Fields("Jobnum")
34                 MyRecord.SONumber = 0
                   'close the recordset and connection
35                 rstRecordSet.Close
36                 conConn.Close

37                 Set rstRecordSet = Nothing
38                 Set conConn = Nothing

39                 GetEpicorODBCData = MyRecord
40                 Exit Function
41             End If
42         End If

           'get job number, order number, order line and misc data from serial number  and order detail tables
43         SQLString = "SELECT"
44         SQLString = SQLString & " SerialNo.JobNum       AS JobNum,"
45         SQLString = SQLString & " SerialNo.PartNum    AS PartNum,"
46         SQLString = SQLString & " SerialNo.SerialNumber AS SerialNo,"
47         SQLString = SQLString & " JobProd.OrderNum     AS SONumber,"
48         SQLString = SQLString & " JobProd.OrderLine    AS SOLine,"
           'SQLString = SQLString & " OrderDtl.Character01  AS ModelNo, "
49         SQLString = SQLString & " OrderDtl.XPartNum  AS XPartNum, "
50         SQLString = SQLString & " OrderHed.CustNum  AS CustNum, "
51         SQLString = SQLString & " OrderHed.ShiptoNum AS ShipToNum, "
52         SQLString = SQLString & " OrderHed.PONum AS CustPONum, "
53         SQLString = SQLString & " Customer.Name AS CustomerName,  "
54         SQLString = SQLString & " ShipTo.Name AS ShipToName  "
55         SQLString = SQLString & " FROM Erp.OrderDtl AS OrderDtl, Erp.SerialNo, Erp.JobProd, Erp.OrderHed, Erp.Customer, Erp.ShipTo "
56         SQLString = SQLString & " WHERE SerialNo.SerialNumber = '" & SerialNumber & "' "
57         SQLString = SQLString & " AND JobProd.JobNum = SerialNo.JobNum "
58         SQLString = SQLString & " AND OrderDtl.OrderNum = JobProd.OrderNum"
59         SQLString = SQLString & " AND OrderDtl.OrderLine = JobProd.OrderLine"
60         SQLString = SQLString & " AND OrderHed.OrderNum = JobProd.OrderNum"
61         SQLString = SQLString & " AND Customer.CustNum = OrderHed.CustNum"
62         SQLString = SQLString & " AND ShipTo.ShipToNum = OrderHed.ShipToNum"
63         SQLString = SQLString & ";"

64         With cmdCommand
65             .ActiveConnection = conConn
66             .CommandText = SQLString
67             .CommandType = adCmdText
68         End With

69         With rstRecordSet
70             If rstRecordSet.State = adStateOpen Then
71                 .Close
72             End If
73            .CursorType = adOpenStatic
74            .CursorLocation = adUseClient
75            .LockType = adLockBatchOptimistic
76            .Open cmdCommand
77          End With

           'if we have a record, save the data, else tell user and leave
78         If rstRecordSet.RecordCount > 0 Then
79             rstRecordSet.MoveFirst
80             MyRecord.SONumber = rstRecordSet.Fields("SONumber")
81             MyRecord.SOLine = rstRecordSet.Fields("SOLine")
       '        MyRecord.ModelNo = rstRecordSet.Fields("ModelNo")
82             MyRecord.PartNum = rstRecordSet.Fields("PartNum")
83             MyRecord.CustNum = rstRecordSet.Fields("CustNum")
84             MyRecord.CustomerPO = rstRecordSet.Fields("CustPONum")
85             MyRecord.ShipToNum = rstRecordSet.Fields("ShipToNum")
86             MyRecord.JobNumber = rstRecordSet.Fields("Jobnum")
87             MyRecord.Customer = rstRecordSet.Fields("CustomerName")
88             MyRecord.XPartNum = rstRecordSet.Fields("XPartNum")
89             MyRecord.ShipTo = IIf(MyRecord.ShipToNum = "", rstRecordSet.Fields("CustomerName"), rstRecordSet.Fields("ShipToName"))
90         Else
91             MsgBox ("No Records found for Serial Number = " & SerialNumber)
92             Exit Function
93         End If

           'get ud02 data
94         SQLString = "SELECT"
95         SQLString = SQLString & " UD02.Number01         AS TDH,"
96         SQLString = SQLString & " UD02.Number02         AS Flow,"
97         SQLString = SQLString & " UD02.Number07         AS ImpellerDiameter,"
98         SQLString = SQLString & " UD02.Number03         AS SuctionPressure,"
99         SQLString = SQLString & " UD02.Number17         AS DesignPressure,"
100        SQLString = SQLString & " UD02.Number14         AS NPSHr"
101        SQLString = SQLString & " FROM Ice.UD02"
102        SQLString = SQLString & " WHERE UD02.Key1 = '" & MyRecord.SONumber & "' "
103        SQLString = SQLString & " AND UD02.Key2 = '" & MyRecord.SOLine & "' "

104        With rstRecordSet
105            .Close
106            cmdCommand.CommandText = SQLString
107            .Open cmdCommand
108        End With

109        If rstRecordSet.RecordCount > 0 Then
110            rstRecordSet.MoveFirst
111            MyRecord.TDH = rstRecordSet.Fields("TDH")
112            MyRecord.Flow = rstRecordSet.Fields("Flow")
113            MyRecord.ImpellerDiameter = rstRecordSet.Fields("ImpellerDiameter")
114            MyRecord.SuctionPressure = rstRecordSet.Fields("SuctionPressure")
115            MyRecord.DesignPressure = rstRecordSet.Fields("DesignPressure")
116            MyRecord.NPSHr = rstRecordSet.Fields("NPSHr")
117        End If

           'get ud03 data
118        SQLString = "SELECT"
119        SQLString = SQLString & " UD03.Number09         AS SpGr,"
120        SQLString = SQLString & " UD03.Character02      AS Fluid,"
121        SQLString = SQLString & " UD03.Number07         AS PumpTemperature,"
122        SQLString = SQLString & " UD03.Number11         AS Viscosity,"
123        SQLString = SQLString & " UD03.Number13         AS VaporPressure,"
124        SQLString = SQLString & " UD03.Number07           As LiquidTemp"
125        SQLString = SQLString & " FROM ice.UD03"
126        SQLString = SQLString & " WHERE UD03.Key1 = '" & MyRecord.SONumber & "' "
127        SQLString = SQLString & " AND UD03.Key2 = '" & MyRecord.SOLine & "' "

128        With rstRecordSet
129            .Close
130            cmdCommand.CommandText = SQLString
131            .Open cmdCommand
132        End With

133        If rstRecordSet.RecordCount > 0 Then
134            rstRecordSet.MoveFirst
135            MyRecord.SpGr = rstRecordSet.Fields("SpGr")
136            MyRecord.Fluid = rstRecordSet.Fields("Fluid")
137            MyRecord.PumpTemperature = rstRecordSet.Fields("PumpTemperature")
138            MyRecord.Viscosity = rstRecordSet.Fields("Viscosity")
139            MyRecord.VaporPressure = rstRecordSet.Fields("VaporPressure")
140            MyRecord.LiquidTemp = rstRecordSet.Fields("LiquidTemp")
141        End If

           'get ud04 data
142        SQLString = "SELECT"
143        SQLString = SQLString & " UD04.Character01      AS SuctFlangeSize,"
144        SQLString = SQLString & " UD04.Character04      AS DischFlangeSize"
145        SQLString = SQLString & " FROM ice.UD04"
146        SQLString = SQLString & " WHERE UD04.Key1 = '" & MyRecord.SONumber & "' "
147        SQLString = SQLString & " AND UD04.Key2 = '" & MyRecord.SOLine & "' "

148        With rstRecordSet
149            .Close
150            cmdCommand.CommandText = SQLString
151            .Open cmdCommand
152        End With

153        If rstRecordSet.RecordCount > 0 Then
154            rstRecordSet.MoveFirst
155            MyRecord.SuctFlangeSize = rstRecordSet.Fields("SuctFlangeSize")
156            MyRecord.DischFlangeSize = rstRecordSet.Fields("DischFlangeSize")
157        End If

           'get ud05 data
158        SQLString = "SELECT"
159        SQLString = SQLString & " UD05.Character05      AS RPM,"
160        SQLString = SQLString & " UD05.Character01      AS Voltage,"
161        SQLString = SQLString & " UD05.Character08      AS StatorFill,"
162        SQLString = SQLString & " UD05.Character02      AS Frequency,"
163        SQLString = SQLString & " UD05.Character03      AS Phases,"
164        SQLString = SQLString & " UD05.Character06        As ThermalClass,"
165        SQLString = SQLString & " UD05.Number01         AS RatedInputPower,"
166        SQLString = SQLString & " UD05.Number02         AS FLCurrent"
167        SQLString = SQLString & " FROM ice.UD05"
168        SQLString = SQLString & " WHERE UD05.Key1 = '" & MyRecord.SONumber & "' "
169        SQLString = SQLString & " AND UD05.Key2 = '" & MyRecord.SOLine & "' "

170        With rstRecordSet
171            .Close
172            cmdCommand.CommandText = SQLString
173            .Open cmdCommand
174        End With

175        If rstRecordSet.RecordCount > 0 Then
176            rstRecordSet.MoveFirst
177            MyRecord.RPM = rstRecordSet.Fields("RPM")
178            MyRecord.Voltage = rstRecordSet.Fields("Voltage")
179            MyRecord.StatorFill = rstRecordSet.Fields("StatorFill")
180            MyRecord.Frequency = rstRecordSet.Fields("Frequency")
181            MyRecord.Phases = rstRecordSet.Fields("Phases")
182            MyRecord.RatedInputPower = rstRecordSet.Fields("RatedInputPower")
183            MyRecord.FLCurrent = rstRecordSet.Fields("FLCurrent")
184            MyRecord.ThermalClass = rstRecordSet.Fields("ThermalClass")
185        End If

           'get ud07 data
186        SQLString = "SELECT"
187        SQLString = SQLString & " UD07.Character01      AS CirculationPath"
188        SQLString = SQLString & " FROM ice.UD07"
189        SQLString = SQLString & " WHERE UD07.Key1 = '" & MyRecord.SONumber & "' "
190        SQLString = SQLString & " AND UD07.Key2 = '" & MyRecord.SOLine & "' "

191        With rstRecordSet
192            .Close
193            cmdCommand.CommandText = SQLString
194            .Open cmdCommand
195        End With

196        If rstRecordSet.RecordCount > 0 Then
197            rstRecordSet.MoveFirst
198            MyRecord.CirculationPath = rstRecordSet.Fields("CirculationPath")
199        End If

           'get ud08 data
200        SQLString = "SELECT"
201        SQLString = SQLString & " UD08.ShortChar01      AS EXPRating"
202        SQLString = SQLString & " FROM ice.UD08"
203        SQLString = SQLString & " WHERE UD08.Key1 = '" & MyRecord.SONumber & "' "
204        SQLString = SQLString & " AND UD08.Key2 = '" & MyRecord.SOLine & "' "

205        With rstRecordSet
206            .Close
207            cmdCommand.CommandText = SQLString
208            .Open cmdCommand
209        End With


210        If rstRecordSet.RecordCount > 0 Then
211            rstRecordSet.MoveFirst
212            MyRecord.ExpClass = rstRecordSet.Fields("EXPRating")
213        End If

           'get ud09 data
214        SQLString = "SELECT"
215        SQLString = SQLString & " UD09.Character01      AS TestProcedure"
216        SQLString = SQLString & " FROM ice.UD09"
217        SQLString = SQLString & " WHERE UD09.Key1 = '" & MyRecord.SONumber & "' "
218        SQLString = SQLString & " AND UD09.Key2 = '" & MyRecord.SOLine & "' "

219        With rstRecordSet
220            .Close
221            cmdCommand.CommandText = SQLString
222            .Open cmdCommand
223        End With

224        If rstRecordSet.RecordCount > 0 Then
225            rstRecordSet.MoveFirst
226            MyRecord.TestProcedure = rstRecordSet.Fields("TestProcedure")
227        End If

           'get part data
228        SQLString = "SELECT"
           'SQLString = SQLString & " Part.Character06      AS MotorSize"
229        SQLString = SQLString & " FROM erp.Part"
230        SQLString = SQLString & " WHERE Part.PartNum = '" & MyRecord.PartNum & "' "

231        With rstRecordSet
        '       .Close
        '       cmdCommand.CommandText = SQLString
        '       .Open cmdCommand
232        End With

       '    If rstRecordSet.RecordCount > 0 Then
       '        rstRecordSet.MoveFirst
       '        MyRecord.MotorSize = rstRecordSet.Fields("MotorSize")
       '    End If

           'get Customer
233        SQLString = "SELECT"
234        SQLString = SQLString & " Customer.Name      AS Customer"
235        SQLString = SQLString & " FROM Customer"
236        SQLString = SQLString & " WHERE Customer.CustNum = '" & MyRecord.CustNum & "' "

237        With rstRecordSet
238            .Close
239            cmdCommand.CommandText = SQLString
240            .Open cmdCommand
241        End With

242        If rstRecordSet.RecordCount > 0 Then
243            rstRecordSet.MoveFirst
244            MyRecord.Customer = rstRecordSet.Fields("Customer")
245        End If

           'get ShipTo
246        If MyRecord.ShipToNum <> "" Then
247            SQLString = "SELECT"
248            SQLString = SQLString & " ShipTo.Name      AS ShipTo"
249            SQLString = SQLString & " FROM ShipTo"
250            SQLString = SQLString & " WHERE ShipTo.ShipToNum = '" & MyRecord.ShipToNum & "' "

251            With rstRecordSet
252                .Close
253                cmdCommand.CommandText = SQLString
254                .Open cmdCommand
255            End With

256            If rstRecordSet.RecordCount > 0 Then
257                rstRecordSet.MoveFirst
258                MyRecord.ShipTo = rstRecordSet.Fields("ShipTo")
259            End If
260        End If

           'close the recordset and connection
261        rstRecordSet.Close
262        conConn.Close

263        Set rstRecordSet = Nothing
264        Set conConn = Nothing

265        GetEpicorODBCData = MyRecord
' <VB WATCH>
266        Exit Function
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
    End Select
' </VB WATCH>
End Function



