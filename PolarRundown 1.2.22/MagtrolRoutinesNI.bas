Attribute VB_Name = "MagtrolRoutinesProLogix"
'    Global vResponse As Variant        'Parsed response from Magtrol
    Global vResponse() As Double
    Global sData As String             'string response from Magtrol
    Global iUD As Integer              'GPIB address of Magtrol
    Global vPlot(2, 100) As Variant    'arrays for mini graph
    Global TCP5 As New TCPIPLibrary.Routines    'for GPIB5
    Global TCP6 As New TCPIPLibrary.Routines    'for GPIB6
    Global TCP As New TCPIPLibrary.Routines     'temp
    Global Const UsingNatInst = True

' <VB WATCH>
Const VBWMODULE = "MagtrolRoutinesProLogix"
' </VB WATCH>

Sub FindMagtrols()
' <VB WATCH>
1          On Error GoTo vbwErrHandler
2          Const VBWPROCNAME = "MagtrolRoutinesProLogix.FindMagtrols"
3          If vbwProtector.vbwTraceProc Then
4              Dim vbwProtectorParameterString As String
5              If vbwProtector.vbwTraceParameters Then
6                  vbwProtectorParameterString = "()"
7              End If
8              vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
9          End If
' </VB WATCH>
10         Dim I As Integer
11         Dim j As Integer
12         Dim MagtrolModel As String

13         Dim rs As New ADODB.Recordset

14         Do While frmPLCData.cmbMagtrol.ListCount > 0
15             frmPLCData.cmbMagtrol.RemoveItem frmPLCData.cmbMagtrol.ListCount - 1
16         Loop

       '==============
17         Dim sGPIBAddress As String
18         Dim sGPIBName As String
19         rs.Open "GPIBAddresses", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTableDirect

20         rs.MoveFirst                                'goto the top
21         For I = 0 To rs.RecordCount - 1             'go through the whole recordset
22             sGPIBAddress = rs.Fields("IPAddress")        'get the description
23             sGPIBName = rs.Fields("GPIBName")                      'get the index number - promary key
24             j = PingSilent(sGPIBAddress)
25             If j <> 0 Then
                   'also get the type of magtrol (5300 or 6530) from CheckMagtrolModel
26                 MagtrolModel = CheckMagtrolModel(sGPIBAddress, sGPIBName)
27                 sGPIBName = sGPIBName & MagtrolModel
28                 If MagtrolModel <> "" Then
29                     frmPLCData.cmbMagtrol.AddItem sGPIBName
30                     frmPLCData.cmbMagtrol.ItemData(frmPLCData.cmbMagtrol.NewIndex) = Val(Mid(sGPIBName, 5, 1))
31                 End If
32             End If
33             rs.MoveNext                             'get the next record
34         Next I
35         rs.Close
36         Set rs = Nothing

37         frmPLCData.cmbMagtrol.AddItem "Add Manually"
38         frmPLCData.cmbMagtrol.ItemData(frmPLCData.cmbMagtrol.NewIndex) = 99
39         frmPLCData.cmbMagtrol.ListIndex = 0

' <VB WATCH>
40         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
41         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FindMagtrols"

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
            vbwReportVariable "j", j
            vbwReportVariable "MagtrolModel", MagtrolModel
            vbwReportVariable "sGPIBAddress", sGPIBAddress
            vbwReportVariable "sGPIBName", sGPIBName
            vbwReportVariable "rs", rs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Function CheckMagtrolModel(GPIBAddress As String, GPIBName As String) As String
' <VB WATCH>
42         On Error GoTo vbwErrHandler
43         Const VBWPROCNAME = "MagtrolRoutinesProLogix.CheckMagtrolModel"
44         If vbwProtector.vbwTraceProc Then
45             Dim vbwProtectorParameterString As String
46             If vbwProtector.vbwTraceParameters Then
47                 vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("GPIBAddress", GPIBAddress) & ", "
48                 vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("GPIBName", GPIBName) & ") "
49             End If
50             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
51         End If
' </VB WATCH>
52         Dim I As Integer
53         Dim strRead As String
54         Dim sSendStr As String

55         If Not UsingNatInst Then
56            strRead = Space$(182)
57            Dim Answer As String
58            If GPIBName = "GPIB5" Then
59             Set TCP = TCP5
60            Else
61             Set TCP = TCP6
62            End If

63            If TCP.ServerAddress <> GPIBAddress Then
64                If TCP.Connected Then
65                    TCP.Disconnect
66                End If
67                TCP.ServerAddress = GPIBAddress
68                TCP.ServerPort = "1234"
69                Answer = TCP.Connect & vbCrLf
70                If Answer = "False" & vbCrLf Then
71                 CheckMagtrolModel = ""
' <VB WATCH>
72         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
73                 Exit Function
74                End If
75                 Answer = TCP.SendGetData("++addr")
76                If Answer <> "14" & vbCrLf Then
77                    TCP.SendGetData ("++addr 14 0")
78                End If
79                Answer = TCP.SendGetData("++eos")
80                If Answer <> "0" & vbCrLf Then
81                    TCP.SendGetData ("++eos 0")
82                End If
83                Answer = TCP.SendGetData("++mode")
84                If Answer <> "1" & vbCrLf Then
85                    TCP.SendGetData ("++mode 1")
86                End If
87                Answer = TCP.SendGetData("++eoi")
88                If Answer <> "1" & vbCrLf Then
89                    TCP.SendGetData ("++eoi 1")
90                End If
91                Answer = TCP.SendGetData("++eot_enable")
92                If Answer <> "1" & vbCrLf Then
93                    TCP.SendGetData ("++eot_enable 1")
94                End If
95                Answer = TCP.SendGetData("++eot_char")
96                If Answer <> "10" & vbCrLf Then
97                    TCP.SendGetData ("++eot_char 10")
98                End If
99                TCP.SendGetData ("++read_tmo_ms 3000")

100           End If

101           Answer = TCP.SendGetData("*IDN?")
102        Else
103            strRead = Space$(182)
104            If GPIBAddress = "192.0.0.145" Then
105                GPIBNo = 5
106            Else
107                GPIBNo = 6
108            End If
               'if we're talking to a magtrol, close the connection
109            If iUD <> 0 Then
110                ibonl iUD, 0
           '        UnregisterGPIBGlobals
111                iUD = 0
112            End If

               'open a new connection to the magtrol:
                   'primary address = 14
                   'secondary address = 0
                   'timeout = 3 second
                   'eoi mode = 1
                   'stop reading when line feed character is received - 0x10
                   'and return iUD

113            ibdev GPIBNo, 14, 0, 11, 1, &H140A, iUD

114            If iberr Then
115                I = 0
           '        Debug.Print GPIBNo & " - i=" & iberr
116                CheckMagtrolModel = ""
117            Else    'if no error
                   'ask who it is
118                sSendStr = "*IDN?" & vbCrLf
119                ibwrt iUD, sSendStr

120                Sleep (1000)

                   'see what the Magtrol says
121                ibrd iUD, strRead
                   '6530 will return a string like 6530 R 1.16"
                   '5300 will return measurement data
122            End If
123            Answer = strRead
124        End If


125            If Left(Answer, 4) = "6530" Then
126                CheckMagtrolModel = " - 6530"
127            ElseIf Left(stAnswerrRead, 2) = "A=" Then
128                CheckMagtrolModel = " - 5300"
129            Else
130                CheckMagtrolModel = " - Unknown"
131            End If
       '        Debug.Print GPIBNo & " - " & strRead


' <VB WATCH>
132        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
133        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "CheckMagtrolModel"

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
            vbwReportVariable "GPIBAddress", GPIBAddress
            vbwReportVariable "GPIBName", GPIBName
            vbwReportVariable "I", I
            vbwReportVariable "strRead", strRead
            vbwReportVariable "sSendStr", sSendStr
            vbwReportVariable "Answer", Answer
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Public Sub SetupMagtrols(MagtrolName As String, GPIBNo As Integer)
' <VB WATCH>
134        On Error GoTo vbwErrHandler
135        Const VBWPROCNAME = "MagtrolRoutinesProLogix.SetupMagtrols"
136        If vbwProtector.vbwTraceProc Then
137            Dim vbwProtectorParameterString As String
138            If vbwProtector.vbwTraceParameters Then
139                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("MagtrolName", MagtrolName) & ", "
140                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("GPIBNo", GPIBNo) & ") "
141            End If
142            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
143        End If
' </VB WATCH>
144        If Not UsingNatInst Then
145            Dim ipaddress As String

146            If GPIBNo = 5 Then
147                ipaddress = "192.0.0.145"
148            Else
149                ipaddress = "192.0.0.146"
150            End If
151            If TCP.ServerAddress <> ipaddress Then
152                If TCP.Connected Then
153                    TCP.Disconnect
154                End If
155                TCP.ServerAddress = ipaddress
156                TCP.ServerPort = "1234"
157                TCP.Connect
158            End If

159            Connected = TCP.Connected
160        Else

           'if we are already talking to a magtrol, close the connection
161        If iUD <> 0 Then
162            ibonl iUD, 0
163        End If

164        ibdev GPIBNo, 14, 0, 11, 1, &H140A, iUD

165        If iberr Then   'if we have an error
166            GPIBNo = 0
167        Else
168            If Right(MagtrolName, 4) = "5300" Then
                   'tell the magtrol that we want full data
169                sSendStr = "FULL" & vbCrLf
170                ibwrt iUD, sSendStr
                   'tell the magtrol that we don't want to wait for data
171                sSendStr = "OPEN" & vbCrLf
172                ibwrt iUD, sSendStr
173            Else
174            End If
175        End If
       '
176        End If
' <VB WATCH>
177        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
178        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SetupMagtrols"

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
            vbwReportVariable "MagtrolName", MagtrolName
            vbwReportVariable "GPIBNo", GPIBNo
            vbwReportVariable "ipaddress", ipaddress
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
