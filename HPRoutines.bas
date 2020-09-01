Attribute VB_Name = "HPRoutines"
Option Explicit

'Global Definitions for the Database Connections
Global cnHP As New ADODB.Connection
Global rsHP As New ADODB.Recordset
Global QyHP As New ADODB.Command
Global cnHPOpen As Boolean      'status

Global rsHPDetail As New ADODB.Recordset
Global rsHPLineNo As New ADODB.Recordset

'Global Definitions for the parameters from the sifil
Global strShipTo As String
Global strBillTo As String
Global strModelNo() As String
Global strSerialNo() As String
Global strCapacity() As String
Global strTDH() As String
Global strImpellers() As String
Global strRPM() As String
Global strSpGr() As String
Global strFluid() As String
Global strPumpTemp() As String
Global strViscosity() As String
Global strVaporPress() As String
Global strSuctPress() As String
Global strDesignPress() As String
Global strSuctFlg() As String
Global strDischFlg() As String
Global strStatorFill() As String
Global strTestProcedure() As String
Global strVoltage() As String
Dim boolUseArchive As Boolean       'true - use the archive files, false - use the current files
Dim boUsingHP As Boolean
Dim boUsingArchive As Boolean
Global intMaxEntries As Integer
Global intLineNo As Integer
Global LogInInitials As String

'initials to approve and delete testing
Global Const strApproveInitials As String = "Admin"
Global boCanApprove As Boolean


Function intFindTheValue(strLineText As String) As String
    'This function will find the value of the parameter on a line of text.
    'The line will consist of a parameter name, a colon, some whitespace and the value.

    'The input, strLineText, will be TDH(FT):       70
    'The output intFindTheValue will be 70.  This will be a string.  Any conversion,
    '  say to a number, must be made by the calling program.

    'If there is no value, an empty string "" will be returned

    Dim strTemp As String
    Dim lngWhere As Long

    'Find the :
    lngWhere = InStr(strLineText, ":")

    'if it's not there, return an empty string
    If IsNull(lngWhere) Then
        intFindTheValue = vbNullString
        Exit Function
    End If

    'Else, get the string to the right of the :
    strTemp = Right$(strLineText, Len(strLineText) - lngWhere)

    'Get rid of the white space around it, and return it
    intFindTheValue = Trim$(strTemp)
  
End Function

Function SearchSalesOrder(strSerialNumber As String)

    'The function writes parameters from the sales order data to the file, PumpDataFromManMan
    'This is called with a sales order number

    Dim I As Integer
    Dim j As Integer
    Dim c As String
    Dim strTemp As String

    Dim strSalesOrderNumber As String

    strSalesOrderNumber = Left$(strSerialNumber, 7)

'    'First, get the bill to and ship to info from soefil

    boolUseArchive = False    'start by using the current files
    GetBillToShipTo strSalesOrderNumber, strShipTo, strBillTo

'
    'if there is no entry, skip the rest
    If Len(strShipTo) < 1 Or IsNull(strShipTo) Then
        I = SaveAnError(strSalesOrderNumber, "Bad Number", strShipTo)
        Exit Function
    End If


    'Get the Detail Data from sifil into a local recordset
    GetDetail (strSalesOrderNumber)



    'get each piece of data

    Dim intNoOfModelNo As Integer
    intNoOfModelNo = intFindData(strModelNo, "MODEL NO:")
    intMaxEntries = intMax(0, intNoOfModelNo)

    Dim intNoOfSerialNo As Integer
    intNoOfSerialNo = intFindData(strSerialNo, "SERIAL NO:")
    intMaxEntries = intMax(intMaxEntries, intNoOfSerialNo)

    Dim intNoOfCapacity As Integer
    intNoOfCapacity = intFindData(strCapacity, "CAPACITY(GPM):")
    intMaxEntries = intMax(intMaxEntries, intNoOfCapacity)

    Dim intNoOfTDH As Integer
    intNoOfTDH = intFindData(strTDH, "TDH(FT):")
    intMaxEntries = intMax(intMaxEntries, intNoOfTDH)

    Dim intNoOfImpeller As Integer
    intNoOfImpeller = intFindData(strImpellers, "*IMPELLER DIA")    'just look for impeller
    intMaxEntries = intMax(intMaxEntries, intNoOfImpeller)         ' could say dia or dia[ in]

    Dim intNoOfRPM As Integer
    intNoOfRPM = intFindData(strRPM, "SPEED:")
    intMaxEntries = intMax(intMaxEntries, intNoOfRPM)

    Dim intNoOfSpGr As Integer
    intNoOfSpGr = intFindData(strSpGr, "SPECIFIC GRAVITY:")
    intMaxEntries = intMax(intMaxEntries, intNoOfSpGr)

    If intNoOfSpGr = 0 Then
        intNoOfSpGr = intFindData(strSpGr, "SP GR:")
        intMaxEntries = intMax(intMaxEntries, intNoOfSpGr)
    End If

    Dim intNoOfFluid As Integer
    intNoOfFluid = intFindData(strFluid, "FLUID:")
    intMaxEntries = intMax(intMaxEntries, intNoOfFluid)

    Dim intNoOfPumpTemp As Integer
    intNoOfPumpTemp = intFindData(strPumpTemp, "PUMPING TEMP:")
    intMaxEntries = intMax(intMaxEntries, intNoOfPumpTemp)

    Dim intNoOfViscosity As Integer
    intNoOfViscosity = intFindData(strViscosity, "VISCOSITY:")
    intMaxEntries = intMax(intMaxEntries, intNoOfViscosity)

    Dim intNoOfVaporPress As Integer
    intNoOfVaporPress = intFindData(strVaporPress, "VAPOR PRESS:")
    intMaxEntries = intMax(intMaxEntries, intNoOfVaporPress)

    Dim intNoOfSuctPress As Integer
    intNoOfSuctPress = intFindData(strSuctPress, "SUCT PRESS:")
    intMaxEntries = intMax(intMaxEntries, intNoOfSuctPress)

    Dim intNoOfDesignPress As Integer
    intNoOfDesignPress = intFindData(strDesignPress, "DESIGN PRESS")
    intMaxEntries = intMax(intMaxEntries, intNoOfDesignPress)

    Dim intNoOfSuctFlg As Integer
    intNoOfSuctFlg = intFindData(strSuctFlg, "SUCT FLG:")
    intMaxEntries = intMax(intMaxEntries, intNoOfSuctFlg)

    Dim intNoOfDischFlg As Integer
    intNoOfDischFlg = intFindData(strDischFlg, "DISCHARGE FLG:")
    intMaxEntries = intMax(intMaxEntries, intNoOfDischFlg)

    Dim intNoOfStatorFill As Integer
    intNoOfStatorFill = intFindData(strStatorFill, "STATOR FILL:")
    intMaxEntries = intMax(intMaxEntries, intNoOfStatorFill)

    Dim intNoOfVoltages As Integer
    intNoOfVoltages = intFindData(strVoltage, "*VOLT:")
    intMaxEntries = intMax(intMaxEntries, intNoOfVoltages)

    Dim intNoOfTestProcedure As Integer
    intNoOfTestProcedure = intFindData(strTestProcedure, "*15852*")
    intMaxEntries = intMax(intMaxEntries, intNoOfTestProcedure)
    I = 15852

    If intNoOfTestProcedure = 0 Then
        intNoOfTestProcedure = intFindData(strTestProcedure, "*15605*")
        intMaxEntries = intMax(intMaxEntries, intNoOfTestProcedure)
        I = 15605
    End If

    If intNoOfTestProcedure = 0 Then
        intNoOfTestProcedure = intFindData(strTestProcedure, "*19021*")
        intMaxEntries = intMax(intMaxEntries, intNoOfTestProcedure)
        I = 19021
    End If

    If intNoOfTestProcedure = 0 Then
        intNoOfTestProcedure = intFindData(strTestProcedure, "*19530*")
        intMaxEntries = intMax(intMaxEntries, intNoOfTestProcedure)
        I = 19530
    End If

    If intNoOfTestProcedure = 0 Then
        intNoOfTestProcedure = intFindData(strTestProcedure, "*17550*")
        intMaxEntries = intMax(intMaxEntries, intNoOfTestProcedure)
        I = 17550
    End If

    If I <> 0 Then
        For j = 1 To 1000                               'get rid of any non-numbers
            If Len(strTestProcedure(1, j)) <> 0 Then
                strTestProcedure(1, j) = Trim$(str$((I)))
            End If
        Next j
    End If

    For I = 1 To 1000                               'get rid of any non-numbers
        If Len(strRPM(1, I)) <> 0 Then
            strTemp = vbNullString
            For j = 1 To Len(strRPM(1, I))
                c = Mid$(strRPM(1, I), j, 1)
                If c >= "0" And c <= "9" Then
                    strTemp = strTemp & c
                Else
                    j = Len(strRPM(1, I))
                End If
            Next j
            strRPM(1, I) = strTemp
        End If
    Next I

    'Open the SOtoShipToBillTo table, and lookup the bill to and ship to names


    If intMaxEntries = 0 Then       'no data found
        I = SaveAnError(strSalesOrderNumber, "No Data Found", strShipTo)
        Exit Function
    End If

    'Get the line numbers and quantities from SODFIL and store them in rsHPLineNo
    GetLineNoQuan (strSalesOrderNumber)


    'can we directly find the serial number that we're looking for?


    With rsHPDetail
        .Filter = "SICOM like '*" & strSerialNumber & "*'"
        If .BOF = True And .EOF = True Then
            intLineNo = 0
        Else
            intLineNo = Int(rsHPDetail.Fields(1))
        End If
    End With

  
End Function
Function GetBillToShipTo(strSalesOrderNumber As String, strShipTo As String, strBillTo As String)
    'enter with the Sales Order Number and return the Ship To and Bill To Names
    ' sets the boUsingHP if we found some data
    ' sets boUsingArchive if we found it in the Archive

    Dim sBillTo As String
    Dim sShipTo As String

    boUsingArchive = False  'assume that we're not usint the archive data

    'define the connection and set up the query
    QyHP.ActiveConnection = cnHP
    QyHP.CommandText = "SELECT SONUM, SHPNO, BILNO FROM SOEFIL WHERE SONUM = '" & _
                              strSalesOrderNumber & "'"

    'execute the query and get back a record set
    Set rsHP = QyHP.Execute()

    'if the record set is empty, try the archive data
    If rsHP.BOF = True And rsHP.EOF = True Then
        boUsingArchive = True
        QyHP.CommandText = "SELECT SONUM, SHPNO, BILNO FROM XSOEFIL WHERE SONUM = '" & _
                                 strSalesOrderNumber & "'"
        Set rsHP = QyHP.Execute()
    End If

    'if this recordset is empty, we can't find the pump
    ' return null data for the bill to and ship to names
    'set boolean to say we can't find the pump
    If rsHP.BOF = True And rsHP.EOF = True Then
        boUsingHP = False
'        MsgBox ("Record Not Found")
        strShipTo = vbNullString
        strBillTo = vbNullString
        Exit Function
    Else
        'else say we found it in the archive and we found the pump
        boUsingHP = True
    End If

    'get the numbers
    sShipTo = rsHP.Fields(1)
    sBillTo = rsHP.Fields(2)

    'get the data
    'set up the query for bill to

    QyHP.CommandText = "SELECT BILNO, BILNAM FROM FINDB.BILMAS WHERE BILNO = '" & _
                             sBillTo & "'"

    'get the recordset -- should be one record
    Set rsHP = QyHP.Execute()

    'extract the name
    strBillTo = rsHP.Fields("BILNAM")

    'do the same for the ship to
    QyHP.CommandText = "SELECT SHPNO, CUSNAM FROM FINDB.CUSFIL WHERE SHPNO = '" & _
                             sShipTo & "'"

    Set rsHP = QyHP.Execute()

    'get the name
    strShipTo = rsHP.Fields("CUSNAM")
  
End Function
Function GetLineNoQuan(strSalesOrderNumber As String)

    'Select the correct file

    Dim strFileName As String
    Dim qry As String

    If boUsingArchive = False Then
        strFileName = "SODFIL"
    Else
        strFileName = "XSODFIL"
    End If

    'Get the Line Numbers and Quantities

    qry = "SELECT SONUM, LINE, SODQO " & _
                                        "FROM " & strFileName & " " & _
                                         "WHERE " & _
                                         " SONUM= '" & strSalesOrderNumber & "' " & _
                                         "ORDER BY LINE;"

    'put it in a new recordset, rsHPDetail

    If rsHPLineNo.State = adStateOpen Then
        rsHPLineNo.Close
    End If

    rsHPLineNo.CursorLocation = adUseClient
    rsHPLineNo.Open qry, cnHP, adOpenStatic

  
End Function
Function GetDetail(strSalesOrderNumber As String)

    'Select the correct file

    Dim strFileName As String

    If boUsingArchive = False Then
        strFileName = "SIFIL"
    Else
        strFileName = "XSIFIL"
    End If

    'Get the details from SIFIL and write them into the SODetails table

    'get the data from the file

    Dim qry As String
    qry = "SELECT SONUM, SILIN, SICOM " & _
                                        "FROM " & strFileName & _
                                         " WHERE " & _
                                         " SONUM = '" & strSalesOrderNumber & _
                                         "' ORDER BY SONUM, SILIN;"

    'put it in a new recordset, rsHPDetail

    If rsHPDetail.State = adStateOpen Then
        rsHPDetail.Close
    End If

    rsHPDetail.CursorLocation = adUseClient
    rsHPDetail.Open qry, cnHP, adOpenStatic
  
End Function
Function intFindData(strArray() As String, strParameter As String) As Integer
    'This will load the array strArray with the
    '  data found in the rst recordset using the strCriteria.
    '  The array is 2 dimensional, the first entry is the line item and the second is the data.

    'The function returns the number of entries found for the criteria

    Dim strLocalData As String      'local variable for data
    Dim strCriteria As String       'criteria to search for
    Dim intCounter As Integer       'counts how many we found
    Dim boolFoundAll As Boolean     'says we found all instances
    Dim intMaxEntry As Integer      'maximum line item number entered into array
    Dim intLineNo As Integer        'line number from sifil
    Dim boolFirstLook As Boolean    'first time we're looking?

    ReDim strArray(1, 1000)            'start with an array of 1x1000

    intCounter = 0                  'none found yet
    intMaxEntry = 0
    boolFoundAll = False            'haven't found all instances

    boolFirstLook = True
    strCriteria = "SICOM like '" & strParameter & "*'"
    With rsHPDetail
        .Filter = 0
        .Filter = strCriteria
        If .BOF = True And .EOF = True Then
            boolFoundAll = True
        Else
            .MoveFirst
        End If
    End With

    While Not boolFoundAll
        With rsHPDetail
            If Not .EOF Then                                'if we found one
                intLineNo = Int(.Fields(1))                     'int portion of line number
                strLocalData = intFindTheValue(.Fields(2))         'strip the model no data
                strLocalData = Replace(strLocalData, "'", " ")
                strArray(1, intLineNo) = strLocalData          'store it in the array
                strArray(0, intLineNo) = str$(Int(.Fields(1)))
                If intLineNo > intMaxEntry Then
                    intMaxEntry = intLineNo
                End If
                intCounter = intCounter + 1                     'inc the counter
                .MoveNext
            Else
                boolFoundAll = True
            End If
        End With
    Wend
    intFindData = intCounter       'return the number of ModelNos found
End Function
Function intMax(intOne As Integer, intTwo As Integer) As Integer
'Returns the greater of the 2 integers sent to it

    If intOne > intTwo Then
        intMax = intOne
    Else
        intMax = intTwo
    End If
End Function
Function SaveAnError(strSalesOrderNumber As String, ErrorType As String, strShipTo As String)
    Dim strSQLStatement As String

    strSQLStatement = "INSERT INTO PumpDataFromManMan ([ErrorType], [SalesOrderNumber], [ShipTo], [BillTo], [ModelNo], [SerialNumber], " & _
                                    "[Capacity], [TDH], [ImpellerDia], [RPM], [SpGr], [Fluid], [PumpTemp], [Viscosity], [VaporPress], " & _
                                    "[SuctPress], [DesignPress], [SuctFlange], [DischFlange], [StatorFill], [TestProcedure]) " & _
                                    "VALUES ('" & ErrorType & "', '" & strSalesOrderNumber & "', '" & _
                                    strShipTo & "', '', '" & _
                                    "', '', '" & _
                                    "', '', '" & _
                                    "', '', '" & _
                                    "', '', '" & _
                                    "', '', '" & _
                                    "', '', '" & _
                                    "', '', '" & _
                                    "', '', '" & _
                                    "');"

  
End Function


Function strFindTheNumber(strInput As String, intStart As Integer) As String
    'Enter with a string and a starting location
    'Return the substring containing the numeric value

    'Ex.  enter with "PHASE: 3 CY: 60  VOLTS: 460" and Start=10
    '  the C of CY.  Return 60.

    Dim L As Integer
    Dim I As Integer
    Dim j As Integer


    If intStart = 0 Then
        strFindTheNumber = vbNullString
        Exit Function
    End If

    L = Len(strInput)
    For I = intStart To L
        'move to the first numeric character
        If IsNumeric(Mid$(strInput, I, 1)) Then
            j = I
            Exit For
        End If
    Next I

    'we're at a number or at the end of the string
    'if its the end of the string, return a null string
    If I - 1 = L Then
        strFindTheNumber = vbNullString
        Exit Function
    End If

    'else we're at a number
    For I = j To L
        If IsNumeric(Mid$(strInput, I, 1)) Then
            strFindTheNumber = strFindTheNumber & Mid$(strInput, I, 1)
        Else
            Exit For
        End If
    Next I
  
End Function



