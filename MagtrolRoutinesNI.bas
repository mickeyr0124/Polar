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

Sub FindMagtrols()
    Dim I As Integer
    Dim j As Integer
    Dim MagtrolModel As String

    Dim rs As New ADODB.Recordset

    Do While frmPLCData.cmbMagtrol.ListCount > 0
        frmPLCData.cmbMagtrol.RemoveItem frmPLCData.cmbMagtrol.ListCount - 1
    Loop

'==============
    Dim sGPIBAddress As String
    Dim sGPIBName As String
    rs.Open "GPIBAddresses", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTableDirect

    rs.MoveFirst                                'goto the top
    For I = 0 To rs.RecordCount - 1             'go through the whole recordset
        sGPIBAddress = rs.Fields("IPAddress")        'get the description
        sGPIBName = rs.Fields("GPIBName")                      'get the index number - promary key
        j = PingSilent(sGPIBAddress)
        If j <> 0 Then
            'also get the type of magtrol (5300 or 6530) from CheckMagtrolModel
            MagtrolModel = CheckMagtrolModel(sGPIBAddress, sGPIBName)
            sGPIBName = sGPIBName & MagtrolModel
            If MagtrolModel <> "" Then
                frmPLCData.cmbMagtrol.AddItem sGPIBName
                frmPLCData.cmbMagtrol.ItemData(frmPLCData.cmbMagtrol.NewIndex) = Val(Mid(sGPIBName, 5, 1))
            End If
        End If
        rs.MoveNext                             'get the next record
    Next I
    rs.Close
    Set rs = Nothing

    frmPLCData.cmbMagtrol.AddItem "Add Manually"
    frmPLCData.cmbMagtrol.ItemData(frmPLCData.cmbMagtrol.NewIndex) = 99
    frmPLCData.cmbMagtrol.ListIndex = 0
  
End Sub
Private Function CheckMagtrolModel(GPIBAddress As String, GPIBName As String) As String
    Dim I As Integer
    Dim strRead As String
    Dim sSendStr As String

    If Not UsingNatInst Then
       strRead = Space$(182)
       Dim Answer As String
       If GPIBName = "GPIB5" Then
        Set TCP = TCP5
       Else
        Set TCP = TCP6
       End If

       If TCP.ServerAddress <> GPIBAddress Then
           If TCP.Connected Then
               TCP.Disconnect
           End If
           TCP.ServerAddress = GPIBAddress
           TCP.ServerPort = "1234"
           Answer = TCP.Connect & vbCrLf
           If Answer = "False" & vbCrLf Then
            CheckMagtrolModel = ""
            Exit Function
           End If
            Answer = TCP.SendGetData("++addr")
           If Answer <> "14" & vbCrLf Then
               TCP.SendGetData ("++addr 14 0")
           End If
           Answer = TCP.SendGetData("++eos")
           If Answer <> "0" & vbCrLf Then
               TCP.SendGetData ("++eos 0")
           End If
           Answer = TCP.SendGetData("++mode")
           If Answer <> "1" & vbCrLf Then
               TCP.SendGetData ("++mode 1")
           End If
           Answer = TCP.SendGetData("++eoi")
           If Answer <> "1" & vbCrLf Then
               TCP.SendGetData ("++eoi 1")
           End If
           Answer = TCP.SendGetData("++eot_enable")
           If Answer <> "1" & vbCrLf Then
               TCP.SendGetData ("++eot_enable 1")
           End If
           Answer = TCP.SendGetData("++eot_char")
           If Answer <> "10" & vbCrLf Then
               TCP.SendGetData ("++eot_char 10")
           End If
           TCP.SendGetData ("++read_tmo_ms 3000")

       End If

       Answer = TCP.SendGetData("*IDN?")
    Else
        strRead = Space$(182)
        If GPIBAddress = "192.0.0.145" Then
            GPIBNo = 5
        Else
            GPIBNo = 6
        End If
        'if we're talking to a magtrol, close the connection
        If iUD <> 0 Then
            ibonl iUD, 0
    '        UnregisterGPIBGlobals
            iUD = 0
        End If

        'open a new connection to the magtrol:
            'primary address = 14
            'secondary address = 0
            'timeout = 3 second
            'eoi mode = 1
            'stop reading when line feed character is received - 0x10
            'and return iUD

        ibdev GPIBNo, 14, 0, 11, 1, &H140A, iUD

        If iberr Then
            I = 0
    '        Debug.Print GPIBNo & " - i=" & iberr
            CheckMagtrolModel = ""
        Else    'if no error
            'ask who it is
            sSendStr = "*IDN?" & vbCrLf
            ibwrt iUD, sSendStr

            Sleep (1000)

            'see what the Magtrol says
            ibrd iUD, strRead
            '6530 will return a string like 6530 R 1.16"
            '5300 will return measurement data
        End If
        Answer = strRead
    End If


        If Left(Answer, 4) = "6530" Then
            CheckMagtrolModel = " - 6530"
        ElseIf Left(stAnswerrRead, 2) = "A=" Then
            CheckMagtrolModel = " - 5300"
        Else
            CheckMagtrolModel = " - Unknown"
        End If
'        Debug.Print GPIBNo & " - " & strRead

  
End Function

Public Sub SetupMagtrols(MagtrolName As String, GPIBNo As Integer)
    If Not UsingNatInst Then
        Dim ipaddress As String

        If GPIBNo = 5 Then
            ipaddress = "192.0.0.145"
        Else
            ipaddress = "192.0.0.146"
        End If
        If TCP.ServerAddress <> ipaddress Then
            If TCP.Connected Then
                TCP.Disconnect
            End If
            TCP.ServerAddress = ipaddress
            TCP.ServerPort = "1234"
            TCP.Connect
        End If

        Connected = TCP.Connected
    Else

    'if we are already talking to a magtrol, close the connection
    If iUD <> 0 Then
        ibonl iUD, 0
    End If

    ibdev GPIBNo, 14, 0, 11, 1, &H140A, iUD

    If iberr Then   'if we have an error
        GPIBNo = 0
    Else
        If Right(MagtrolName, 4) = "5300" Then
            'tell the magtrol that we want full data
            sSendStr = "FULL" & vbCrLf
            ibwrt iUD, sSendStr
            'tell the magtrol that we don't want to wait for data
            sSendStr = "OPEN" & vbCrLf
            ibwrt iUD, sSendStr
        Else
        End If
    End If
'
    End If
End Sub


