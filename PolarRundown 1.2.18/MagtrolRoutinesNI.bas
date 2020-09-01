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
' </VB WATCH>
2          Dim I As Integer
3          Dim j As Integer
4          Dim MagtrolModel As String

5          Dim rs As New ADODB.Recordset

6          Do While frmPLCData.cmbMagtrol.ListCount > 0
7              frmPLCData.cmbMagtrol.RemoveItem frmPLCData.cmbMagtrol.ListCount - 1
8          Loop

       '==============
9          Dim sGPIBAddress As String
10         Dim sGPIBName As String
11         rs.Open "GPIBAddresses", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTableDirect

12         rs.MoveFirst                                'goto the top
13         For I = 0 To rs.RecordCount - 1             'go through the whole recordset
14             sGPIBAddress = rs.Fields("IPAddress")        'get the description
15             sGPIBName = rs.Fields("GPIBName")                      'get the index number - promary key
16             j = PingSilent(sGPIBAddress)
17             If j <> 0 Then
                   'also get the type of magtrol (5300 or 6530) from CheckMagtrolModel
18                 MagtrolModel = CheckMagtrolModel(sGPIBAddress, sGPIBName)
19                 sGPIBName = sGPIBName & MagtrolModel
20                 If MagtrolModel <> "" Then
21                     frmPLCData.cmbMagtrol.AddItem sGPIBName
22                     frmPLCData.cmbMagtrol.ItemData(frmPLCData.cmbMagtrol.NewIndex) = Val(Mid(sGPIBName, 5, 1))
23                 End If
24             End If
25             rs.MoveNext                             'get the next record
26         Next I
27         rs.Close
28         Set rs = Nothing

29         frmPLCData.cmbMagtrol.AddItem "Add Manually"
30         frmPLCData.cmbMagtrol.ItemData(frmPLCData.cmbMagtrol.NewIndex) = 99
31         frmPLCData.cmbMagtrol.ListIndex = 0

' <VB WATCH>
32         Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Function CheckMagtrolModel(GPIBAddress As String, GPIBName As String) As String
' <VB WATCH>
33         On Error GoTo vbwErrHandler
' </VB WATCH>
34         Dim I As Integer
35         Dim strRead As String
36         Dim sSendStr As String

37         If Not UsingNatInst Then
38            strRead = Space$(182)
39            Dim Answer As String
40            If GPIBName = "GPIB5" Then
41             Set TCP = TCP5
42            Else
43             Set TCP = TCP6
44            End If

45            If TCP.ServerAddress <> GPIBAddress Then
46                If TCP.Connected Then
47                    TCP.Disconnect
48                End If
49                TCP.ServerAddress = GPIBAddress
50                TCP.ServerPort = "1234"
51                Answer = TCP.Connect & vbCrLf
52                If Answer = "False" & vbCrLf Then
53                 CheckMagtrolModel = ""
54                 Exit Function
55                End If
56                 Answer = TCP.SendGetData("++addr")
57                If Answer <> "14" & vbCrLf Then
58                    TCP.SendGetData ("++addr 14 0")
59                End If
60                Answer = TCP.SendGetData("++eos")
61                If Answer <> "0" & vbCrLf Then
62                    TCP.SendGetData ("++eos 0")
63                End If
64                Answer = TCP.SendGetData("++mode")
65                If Answer <> "1" & vbCrLf Then
66                    TCP.SendGetData ("++mode 1")
67                End If
68                Answer = TCP.SendGetData("++eoi")
69                If Answer <> "1" & vbCrLf Then
70                    TCP.SendGetData ("++eoi 1")
71                End If
72                Answer = TCP.SendGetData("++eot_enable")
73                If Answer <> "1" & vbCrLf Then
74                    TCP.SendGetData ("++eot_enable 1")
75                End If
76                Answer = TCP.SendGetData("++eot_char")
77                If Answer <> "10" & vbCrLf Then
78                    TCP.SendGetData ("++eot_char 10")
79                End If
80                TCP.SendGetData ("++read_tmo_ms 3000")

81            End If

82            Answer = TCP.SendGetData("*IDN?")
83         Else
84             strRead = Space$(182)
85             If GPIBAddress = "192.0.0.145" Then
86                 GPIBNo = 5
87             Else
88                 GPIBNo = 6
89             End If
               'if we're talking to a magtrol, close the connection
90             If iUD <> 0 Then
91                 ibonl iUD, 0
           '        UnregisterGPIBGlobals
92                 iUD = 0
93             End If

               'open a new connection to the magtrol:
                   'primary address = 14
                   'secondary address = 0
                   'timeout = 3 second
                   'eoi mode = 1
                   'stop reading when line feed character is received - 0x10
                   'and return iUD

94             ibdev GPIBNo, 14, 0, 11, 1, &H140A, iUD

95             If iberr Then
96                 I = 0
           '        Debug.Print GPIBNo & " - i=" & iberr
97                 CheckMagtrolModel = ""
98             Else    'if no error
                   'ask who it is
99                 sSendStr = "*IDN?" & vbCrLf
100                ibwrt iUD, sSendStr

101                Sleep (1000)

                   'see what the Magtrol says
102                ibrd iUD, strRead
                   '6530 will return a string like 6530 R 1.16"
                   '5300 will return measurement data
103            End If
104            Answer = strRead
105        End If


106            If Left(Answer, 4) = "6530" Then
107                CheckMagtrolModel = " - 6530"
108            ElseIf Left(stAnswerrRead, 2) = "A=" Then
109                CheckMagtrolModel = " - 5300"
110            Else
111                CheckMagtrolModel = " - Unknown"
112            End If
       '        Debug.Print GPIBNo & " - " & strRead


' <VB WATCH>
113        Exit Function
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
    End Select
' </VB WATCH>
End Function

Public Sub SetupMagtrols(MagtrolName As String, GPIBNo As Integer)
' <VB WATCH>
114        On Error GoTo vbwErrHandler
' </VB WATCH>
115        If Not UsingNatInst Then
116            Dim ipaddress As String

117            If GPIBNo = 5 Then
118                ipaddress = "192.0.0.145"
119            Else
120                ipaddress = "192.0.0.146"
121            End If
122            If TCP.ServerAddress <> ipaddress Then
123                If TCP.Connected Then
124                    TCP.Disconnect
125                End If
126                TCP.ServerAddress = ipaddress
127                TCP.ServerPort = "1234"
128                TCP.Connect
129            End If

130            Connected = TCP.Connected
131        Else

           'if we are already talking to a magtrol, close the connection
132        If iUD <> 0 Then
133            ibonl iUD, 0
134        End If

135        ibdev GPIBNo, 14, 0, 11, 1, &H140A, iUD

136        If iberr Then   'if we have an error
137            GPIBNo = 0
138        Else
139            If Right(MagtrolName, 4) = "5300" Then
                   'tell the magtrol that we want full data
140                sSendStr = "FULL" & vbCrLf
141                ibwrt iUD, sSendStr
                   'tell the magtrol that we don't want to wait for data
142                sSendStr = "OPEN" & vbCrLf
143                ibwrt iUD, sSendStr
144            Else
145            End If
146        End If
       '
147        End If
' <VB WATCH>
148        Exit Sub
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
    End Select
' </VB WATCH>
End Sub



