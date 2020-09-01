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

Public Function GetEpicorODBCData(SerialNumber As String, EpicorConnectionString As String) As SNRecord
    Dim conConn As New ADODB.Connection
    Dim cmdCommand As New ADODB.Command
    Dim rstRecordSet As New ADODB.Recordset
    Dim SQLString As String

    Dim MyRecord As SNRecord

    'construct connection string
    conConn.Open EpicorConnectionString
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

    SQLString = "SELECT"
    SQLString = SQLString & " SerialNo.JobNum       AS JobNum,"
    SQLString = SQLString & " SerialNo.PartNum    AS PartNum,"
    SQLString = SQLString & " SerialNo.SerialNumber AS SerialNo,"
    SQLString = SQLString & " JobProd.OrderNum     AS SONumber, "
    SQLString = SQLString & " JobProd.JobNum AS  JobProdJobNum "
    SQLString = SQLString & " FROM Erp.SerialNo, Erp.JobProd "
    SQLString = SQLString & " WHERE SerialNo.SerialNumber = '" & SerialNumber & "' "
    SQLString = SQLString & " AND JobProd.JobNum = SerialNo.JobNum "
    SQLString = SQLString & ";"
    
    With cmdCommand
        .ActiveConnection = conConn
        .CommandText = SQLString
        .CommandType = adCmdText
    End With

    With rstRecordSet
       .CursorType = adOpenStatic
       .CursorLocation = adUseClient
       .LockType = adLockBatchOptimistic
       .Open cmdCommand
     End With

    'if we have a record, save the data, else tell user and leave
    If rstRecordSet.RecordCount > 0 Then    'there is no order no
        If rstRecordSet.Fields("SONumber") = 0 Then
            rstRecordSet.MoveFirst
            MyRecord.PartNum = rstRecordSet.Fields("PartNum")
            MyRecord.JobNumber = rstRecordSet.Fields("Jobnum")
            MyRecord.SONumber = 0
            'close the recordset and connection
            rstRecordSet.Close
            conConn.Close
        
            Set rstRecordSet = Nothing
            Set conConn = Nothing
        
            GetEpicorODBCData = MyRecord
            Exit Function
        End If
    End If
    
    'get job number, order number, order line and misc data from serial number  and order detail tables
    SQLString = "SELECT"
    SQLString = SQLString & " SerialNo.JobNum       AS JobNum,"
    SQLString = SQLString & " SerialNo.PartNum    AS PartNum,"
    SQLString = SQLString & " SerialNo.SerialNumber AS SerialNo,"
    SQLString = SQLString & " JobProd.OrderNum     AS SONumber,"
    SQLString = SQLString & " JobProd.OrderLine    AS SOLine,"
    'SQLString = SQLString & " OrderDtl.Character01  AS ModelNo, "
    SQLString = SQLString & " OrderDtl.XPartNum  AS XPartNum, "
    SQLString = SQLString & " OrderHed.CustNum  AS CustNum, "
    SQLString = SQLString & " OrderHed.ShiptoNum AS ShipToNum, "
    SQLString = SQLString & " OrderHed.PONum AS CustPONum, "
    SQLString = SQLString & " Customer.Name AS CustomerName,  "
    SQLString = SQLString & " ShipTo.Name AS ShipToName  "
    SQLString = SQLString & " FROM Erp.OrderDtl AS OrderDtl, Erp.SerialNo, Erp.JobProd, Erp.OrderHed, Erp.Customer, Erp.ShipTo "
    SQLString = SQLString & " WHERE SerialNo.SerialNumber = '" & SerialNumber & "' "
    SQLString = SQLString & " AND JobProd.JobNum = SerialNo.JobNum "
    SQLString = SQLString & " AND OrderDtl.OrderNum = JobProd.OrderNum"
    SQLString = SQLString & " AND OrderDtl.OrderLine = JobProd.OrderLine"
    SQLString = SQLString & " AND OrderHed.OrderNum = JobProd.OrderNum"
    SQLString = SQLString & " AND Customer.CustNum = OrderHed.CustNum"
    SQLString = SQLString & " AND ShipTo.ShipToNum = OrderHed.ShipToNum"
    SQLString = SQLString & ";"

    With cmdCommand
        .ActiveConnection = conConn
        .CommandText = SQLString
        .CommandType = adCmdText
    End With

    With rstRecordSet
        If rstRecordSet.State = adStateOpen Then
            .Close
        End If
       .CursorType = adOpenStatic
       .CursorLocation = adUseClient
       .LockType = adLockBatchOptimistic
       .Open cmdCommand
     End With

    'if we have a record, save the data, else tell user and leave
    If rstRecordSet.RecordCount > 0 Then
        rstRecordSet.MoveFirst
        MyRecord.SONumber = rstRecordSet.Fields("SONumber")
        MyRecord.SOLine = rstRecordSet.Fields("SOLine")
'        MyRecord.ModelNo = rstRecordSet.Fields("ModelNo")
        MyRecord.PartNum = rstRecordSet.Fields("PartNum")
        MyRecord.CustNum = rstRecordSet.Fields("CustNum")
        MyRecord.CustomerPO = rstRecordSet.Fields("CustPONum")
        MyRecord.ShipToNum = rstRecordSet.Fields("ShipToNum")
        MyRecord.JobNumber = rstRecordSet.Fields("Jobnum")
        MyRecord.Customer = rstRecordSet.Fields("CustomerName")
        MyRecord.XPartNum = rstRecordSet.Fields("XPartNum")
        MyRecord.ShipTo = IIf(MyRecord.ShipToNum = "", rstRecordSet.Fields("CustomerName"), rstRecordSet.Fields("ShipToName"))
    Else
        MsgBox ("No Records found for Serial Number = " & SerialNumber)
        Exit Function
    End If

    'get ud02 data
    SQLString = "SELECT"
    SQLString = SQLString & " UD02.Number01         AS TDH,"
    SQLString = SQLString & " UD02.Number02         AS Flow,"
    SQLString = SQLString & " UD02.Number07         AS ImpellerDiameter,"
    SQLString = SQLString & " UD02.Number03         AS SuctionPressure,"
    SQLString = SQLString & " UD02.Number17         AS DesignPressure,"
    SQLString = SQLString & " UD02.Number14         AS NPSHr"
    SQLString = SQLString & " FROM Ice.UD02"
    SQLString = SQLString & " WHERE UD02.Key1 = '" & MyRecord.SONumber & "' "
    SQLString = SQLString & " AND UD02.Key2 = '" & MyRecord.SOLine & "' "

    With rstRecordSet
        .Close
        cmdCommand.CommandText = SQLString
        .Open cmdCommand
    End With

    If rstRecordSet.RecordCount > 0 Then
        rstRecordSet.MoveFirst
        MyRecord.TDH = rstRecordSet.Fields("TDH")
        MyRecord.Flow = rstRecordSet.Fields("Flow")
        MyRecord.ImpellerDiameter = rstRecordSet.Fields("ImpellerDiameter")
        MyRecord.SuctionPressure = rstRecordSet.Fields("SuctionPressure")
        MyRecord.DesignPressure = rstRecordSet.Fields("DesignPressure")
        MyRecord.NPSHr = rstRecordSet.Fields("NPSHr")
    End If

    'get ud03 data
    SQLString = "SELECT"
    SQLString = SQLString & " UD03.Number09         AS SpGr,"
    SQLString = SQLString & " UD03.Character02      AS Fluid,"
    SQLString = SQLString & " UD03.Number07         AS PumpTemperature,"
    SQLString = SQLString & " UD03.Number11         AS Viscosity,"
    SQLString = SQLString & " UD03.Number13         AS VaporPressure,"
    SQLString = SQLString & " UD03.Number07           As LiquidTemp"
    SQLString = SQLString & " FROM ice.UD03"
    SQLString = SQLString & " WHERE UD03.Key1 = '" & MyRecord.SONumber & "' "
    SQLString = SQLString & " AND UD03.Key2 = '" & MyRecord.SOLine & "' "

    With rstRecordSet
        .Close
        cmdCommand.CommandText = SQLString
        .Open cmdCommand
    End With

    If rstRecordSet.RecordCount > 0 Then
        rstRecordSet.MoveFirst
        MyRecord.SpGr = rstRecordSet.Fields("SpGr")
        MyRecord.Fluid = rstRecordSet.Fields("Fluid")
        MyRecord.PumpTemperature = rstRecordSet.Fields("PumpTemperature")
        MyRecord.Viscosity = rstRecordSet.Fields("Viscosity")
        MyRecord.VaporPressure = rstRecordSet.Fields("VaporPressure")
        MyRecord.LiquidTemp = rstRecordSet.Fields("LiquidTemp")
    End If

    'get ud04 data
    SQLString = "SELECT"
    SQLString = SQLString & " UD04.Character01      AS SuctFlangeSize,"
    SQLString = SQLString & " UD04.Character04      AS DischFlangeSize"
    SQLString = SQLString & " FROM ice.UD04"
    SQLString = SQLString & " WHERE UD04.Key1 = '" & MyRecord.SONumber & "' "
    SQLString = SQLString & " AND UD04.Key2 = '" & MyRecord.SOLine & "' "

    With rstRecordSet
        .Close
        cmdCommand.CommandText = SQLString
        .Open cmdCommand
    End With

    If rstRecordSet.RecordCount > 0 Then
        rstRecordSet.MoveFirst
        MyRecord.SuctFlangeSize = rstRecordSet.Fields("SuctFlangeSize")
        MyRecord.DischFlangeSize = rstRecordSet.Fields("DischFlangeSize")
    End If

    'get ud05 data
    SQLString = "SELECT"
    SQLString = SQLString & " UD05.Character05      AS RPM,"
    SQLString = SQLString & " UD05.Character01      AS Voltage,"
    SQLString = SQLString & " UD05.Character08      AS StatorFill,"
    SQLString = SQLString & " UD05.Character02      AS Frequency,"
    SQLString = SQLString & " UD05.Character03      AS Phases,"
    SQLString = SQLString & " UD05.Character06        As ThermalClass,"
    SQLString = SQLString & " UD05.Number01         AS RatedInputPower,"
    SQLString = SQLString & " UD05.Number02         AS FLCurrent"
    SQLString = SQLString & " FROM ice.UD05"
    SQLString = SQLString & " WHERE UD05.Key1 = '" & MyRecord.SONumber & "' "
    SQLString = SQLString & " AND UD05.Key2 = '" & MyRecord.SOLine & "' "

    With rstRecordSet
        .Close
        cmdCommand.CommandText = SQLString
        .Open cmdCommand
    End With

    If rstRecordSet.RecordCount > 0 Then
        rstRecordSet.MoveFirst
        MyRecord.RPM = rstRecordSet.Fields("RPM")
        MyRecord.Voltage = rstRecordSet.Fields("Voltage")
        MyRecord.StatorFill = rstRecordSet.Fields("StatorFill")
        MyRecord.Frequency = rstRecordSet.Fields("Frequency")
        MyRecord.Phases = rstRecordSet.Fields("Phases")
        MyRecord.RatedInputPower = rstRecordSet.Fields("RatedInputPower")
        MyRecord.FLCurrent = rstRecordSet.Fields("FLCurrent")
        MyRecord.ThermalClass = rstRecordSet.Fields("ThermalClass")
    End If

    'get ud07 data
    SQLString = "SELECT"
    SQLString = SQLString & " UD07.Character01      AS CirculationPath"
    SQLString = SQLString & " FROM ice.UD07"
    SQLString = SQLString & " WHERE UD07.Key1 = '" & MyRecord.SONumber & "' "
    SQLString = SQLString & " AND UD07.Key2 = '" & MyRecord.SOLine & "' "

    With rstRecordSet
        .Close
        cmdCommand.CommandText = SQLString
        .Open cmdCommand
    End With

    If rstRecordSet.RecordCount > 0 Then
        rstRecordSet.MoveFirst
        MyRecord.CirculationPath = rstRecordSet.Fields("CirculationPath")
    End If

    'get ud08 data
    SQLString = "SELECT"
    SQLString = SQLString & " UD08.ShortChar01      AS EXPRating"
    SQLString = SQLString & " FROM ice.UD08"
    SQLString = SQLString & " WHERE UD08.Key1 = '" & MyRecord.SONumber & "' "
    SQLString = SQLString & " AND UD08.Key2 = '" & MyRecord.SOLine & "' "

    With rstRecordSet
        .Close
        cmdCommand.CommandText = SQLString
        .Open cmdCommand
    End With


    If rstRecordSet.RecordCount > 0 Then
        rstRecordSet.MoveFirst
        MyRecord.ExpClass = rstRecordSet.Fields("EXPRating")
    End If

    'get ud09 data
    SQLString = "SELECT"
    SQLString = SQLString & " UD09.Character01      AS TestProcedure"
    SQLString = SQLString & " FROM ice.UD09"
    SQLString = SQLString & " WHERE UD09.Key1 = '" & MyRecord.SONumber & "' "
    SQLString = SQLString & " AND UD09.Key2 = '" & MyRecord.SOLine & "' "

    With rstRecordSet
        .Close
        cmdCommand.CommandText = SQLString
        .Open cmdCommand
    End With

    If rstRecordSet.RecordCount > 0 Then
        rstRecordSet.MoveFirst
        MyRecord.TestProcedure = rstRecordSet.Fields("TestProcedure")
    End If

    'get part data
    SQLString = "SELECT"
    'SQLString = SQLString & " Part.Character06      AS MotorSize"
    SQLString = SQLString & " FROM erp.Part"
    SQLString = SQLString & " WHERE Part.PartNum = '" & MyRecord.PartNum & "' "

    With rstRecordSet
 '       .Close
 '       cmdCommand.CommandText = SQLString
 '       .Open cmdCommand
    End With

'    If rstRecordSet.RecordCount > 0 Then
'        rstRecordSet.MoveFirst
'        MyRecord.MotorSize = rstRecordSet.Fields("MotorSize")
'    End If

    'get Customer
    SQLString = "SELECT"
    SQLString = SQLString & " Customer.Name      AS Customer"
    SQLString = SQLString & " FROM Customer"
    SQLString = SQLString & " WHERE Customer.CustNum = '" & MyRecord.CustNum & "' "

    With rstRecordSet
        .Close
        cmdCommand.CommandText = SQLString
        .Open cmdCommand
    End With

    If rstRecordSet.RecordCount > 0 Then
        rstRecordSet.MoveFirst
        MyRecord.Customer = rstRecordSet.Fields("Customer")
    End If

    'get ShipTo
    If MyRecord.ShipToNum <> "" Then
        SQLString = "SELECT"
        SQLString = SQLString & " ShipTo.Name      AS ShipTo"
        SQLString = SQLString & " FROM ShipTo"
        SQLString = SQLString & " WHERE ShipTo.ShipToNum = '" & MyRecord.ShipToNum & "' "

        With rstRecordSet
            .Close
            cmdCommand.CommandText = SQLString
            .Open cmdCommand
        End With

        If rstRecordSet.RecordCount > 0 Then
            rstRecordSet.MoveFirst
            MyRecord.ShipTo = rstRecordSet.Fields("ShipTo")
        End If
    End If

    'close the recordset and connection
    rstRecordSet.Close
    conConn.Close

    Set rstRecordSet = Nothing
    Set conConn = Nothing

    GetEpicorODBCData = MyRecord
End Function


