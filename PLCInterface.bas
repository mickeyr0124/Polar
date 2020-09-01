Attribute VB_Name = "PLCInterface"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Public Declare Sub SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long)
Public Const MAX_COMPUTERNAME_LENGTH As Long = 15&
Public Declare Function GetComputerName Lib "kernel32" Alias "GetComputerNameA" (ByVal lpBuffer As String, nSize As Long) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' HEI API Defines
'
    Const HEIAPIVersion As Byte = 3
    Const HEIP_IP As Integer = 3
    Const HEIT_WINSOCK As Integer = 4

    Const DefDevTimeout As Integer = 50                        ' value in milliseconds
    Const DefDevRetrys As Byte = 3

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'
    Private Type Encryption
        Algorithm As Byte ' Algorithm to use for encryption: 0= No encryption, 1= private key encryption
        Unused1(2) As Byte              ' Reserved
        Key(59) As Byte                 ' Encryption key (null terminated)
    End Type

    Private Type EnetAddress
        Address(19) As Byte
    End Type


    Private Type HEITransport
        Transport As Integer
        Protocol As Integer
        Encrypt As Encryption
        SourceAddress As EnetAddress
        Reserved(47) As Byte
    End Type

    Private Type HEIDevice
        Address(125) As Byte             ' 126-byte byte array (VB packs on 4-byte boundaries)
    End Type

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Host Ethernet APIs
'
    Private Declare Function PASCAL_HEIOpen Lib "hei_pas" ( _
        ByVal HEIAPIVersion As Integer _
    ) As Long

    Private Declare Function PASCAL_HEIClose Lib "hei_pas" ( _
    ) As Long

    Private Declare Function PASCAL_HEIOpenTransport Lib "hei_pas" ( _
        ByRef pTransport As HEITransport, _
        ByVal HEIAPIVersion As Integer, _
        ByVal EnetAdress As Long _
    ) As Long

    Private Declare Function PASCAL_HEICloseTransport Lib "hei_pas" ( _
        ByRef pTransport As HEITransport _
    ) As Long

    Private Declare Function PASCAL_HEIOpenDevice Lib "hei_pas" ( _
        ByRef pTransport As HEITransport, _
        ByRef pDevice As HEIDevice, _
        ByVal HEIAPIVersion As Integer, _
        ByVal Timeout As Integer, _
        ByVal Retrys As Integer, _
        ByVal UseAddressedBroadcast As Boolean _
    ) As Long

    Private Declare Function PASCAL_HEICloseDevice Lib "hei_pas" ( _
        ByRef pDevice As HEIDevice _
    ) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' ECOM Specific APIs
'
    Private Declare Function PASCAL_HEICCMRequest Lib "hei_pas" ( _
        ByRef pDevice As HEIDevice, _
        ByVal bWrite As Integer, _
        ByVal DataType As Byte, _
        ByVal Address As Integer, _
        ByVal pDataLen As Integer, _
        ByRef pData As Byte _
    ) As Long

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Global variables
'
    ' return code from the SDK API calls
    Global rc As Long

    ' Ethernet protocol transport
    Global TP As HEITransport

    ' true if the network interface can be initialized using the selected protocol
    Global NetworkOK As Boolean

    ' maximum number of devices you want to allow
    Const MAXDEVICES As Integer = 100

    ' array of Host Ethernet devices
    Global aDevices(MAXDEVICES) As HEIDevice

    ' number of Host Ethernet devices found on the network
    Global DeviceCount As Long

    ' set to true if any Host Ethernet device is already open
    Global DeviceOpen As Boolean

    ' this is the device the user selected from the list
    Global tDevice As Integer

    ' this is the type of device the user selcted
'    Global tDeviceType As String

    ' detail line that gets displayed in the listbox
'    Global DetailLine As String

    Global bWrite As Long
    Global DataType As Byte
    Global DataAddress As Integer
    Global DataLength As Integer
    Global ByteBuffer(255) As Byte

    Global Description(MAXDEVICES) As String

Function NetWorkInitialize() As Long
    'return rc from pascal calls
    '  0 says ok

    ' if the network interface has already been opened, close it
    '
    If NetworkOK = True Then
        rc = PASCAL_HEICloseTransport(TP)
        rc = PASCAL_HEIClose()
    End If

    '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
    ' Initialize the Ethernet Driver
    '
    rc = PASCAL_HEIOpen(HEIAPIVersion)
    If rc <> 0 Then
        NetWorkInitialize = rc  'return error code
        Exit Function
    Else
        '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
        ' Initiaizize the Winsock protocol transport
        '
        TP.Transport = HEIT_WINSOCK

        TP.Protocol = HEIP_IP

        rc = PASCAL_HEIOpenTransport(TP, HEIAPIVersion, 0)

        NetWorkInitialize = rc

    End If
  
End Function
'***********************************************************************
' Since the PASCAL_HEIxxxx calls require a byte buffer, convert the user
' entered strings to byte arrays
'
Function StringToByteArray(ByVal inString As String, ByRef Buffer() As Byte) As Integer
Dim I As Integer
Dim U() As Byte

    'Make sure all alpha characters are uppercase
    U = StrConv(inString, vbUpperCase)

    'skip over the Unicode byte
    For I = 0 To (Len(inString) - 1)
        Buffer(I) = U(I * 2)
    Next I

    StringToByteArray = I
  
End Function

'***********************************************************************
' Swap successive entries in a byte array
'
Function ByteSwap(ByRef Buffer() As Byte, Count As Integer) As Integer
Dim I As Integer
Dim temp As Byte

    For I = 0 To Count - 1 Step 2
        temp = Buffer(I)
        Buffer(I) = Buffer(I + 1)
        Buffer(I + 1) = temp
    Next I

    ByteSwap = I
  
End Function

'************************************************************************
' Convert a byte array of character codes to a packed array of characters
'
Function HexConvert(ByRef Buffer() As Byte, Count As Integer) As Integer
Dim I As Integer

    'convert each character code
    For I = 0 To (Count * 2) - 1

        'have to manually process HEX character digits
        If (Buffer(I) > 64) And (Buffer(I) < 71) Then
            Select Case Buffer(I)
                Case 65 'A
                    Buffer(I) = 10
                Case 66 'B
                    Buffer(I) = 11
                Case 67 'C
                    Buffer(I) = 12
                Case 68 'D
                    Buffer(I) = 13
                Case 69 'E
                    Buffer(I) = 14
                Case 70 'F
                    Buffer(I) = 15
            End Select

        Else
            'numeric digits are much easier
            Buffer(I) = ChrW$(Buffer(I))

        End If

    Next I

    'Now pack two HEX characters into a byte
    Dim Z As Integer
    Z = 0
    For I = 0 To (Count * 2) - 1 Step 2
        Buffer(Z) = (Buffer(I) * 16) + Buffer(I + 1)
        Z = Z + 1
    Next I

    'Now clear the remainder of the byte array - just to be neat and complete
    For I = Z To (Count * 2) - 1
        Buffer(I) = 0
    Next I

    HexConvert = Z
  
End Function

'***********************************************************************
' Brute force method of converting a 4 character string to a HEX number
'
Function StringToHexInt(inData As String) As Integer

    Dim I As Integer, j As Integer
    Dim t(4) As Byte

    'convert from octal

    j = 0
    For I = 1 To Len(inData)
        j = j + Val(Mid$(inData, Len(inData) - I + 1, 1)) * (8 ^ (I - 1))
    Next I

    StringToHexInt = j + 1
    Exit Function

    inData = Hex$(j + 1)

    I = StringToByteArray(inData, t)

    'convert each character code
    For I = 0 To (Len(inData) - 1)

        'have to manually process HEX characters digits
        If (t(I) > 64) And (t(I) < 71) Then
            Select Case t(I)
                Case 65 'A
                    t(I) = 10
                Case 66 'B
                    t(I) = 11
                Case 67 'C
                    t(I) = 12
                Case 68 'D
                    t(I) = 13
                Case 69 'E
                    t(I) = 14
                Case 70 'F
                    t(I) = 15
            End Select

        Else
            'numeric digits are much easier
            t(I) = ChrW$(t(I))

        End If
    Next I

    Select Case Len(inData)
        Case 0
            StringToHexInt = 0
        Case 1
            StringToHexInt = t(0)
        Case 2
            StringToHexInt = (t(0) * 16) + t(1)
        Case 3
            StringToHexInt = (t(0) * 256) + (t(1) * 16) + t(2)
        Case 4
            StringToHexInt = (t(0) * 4096) + (t(1) * 256) + (t(2) * 16) + t(3)
    End Select
  
End Function
Function ConnectToPLC(DeviceNo As Integer) As String
    ' Open the device
    '
    rc = PASCAL_HEIOpenDevice(TP, aDevices(DeviceNo), HEIAPIVersion, DefDevTimeout, DefDevRetrys, False)
    If rc <> 0 Then
        DeviceOpen = False
    Else
        DeviceOpen = True
    End If
        ConnectToPLC = rc
  
End Function

Function DisconnectPLC() As String
    rc = PASCAL_HEICloseDevice(aDevices(tDevice))
    DisconnectPLC = rc
End Function
Function GetData() As String
    GetData = PASCAL_HEICCMRequest(aDevices(tDevice), bWrite, DataType, DataAddress, DataLength, ByteBuffer(0))
End Function
'Function ConvertToReal(ByteBuffer() As Byte) As Single
Function ConvertToReal(Address As String) As Single
    Dim sFloat As Single
    Dim lngNum As Long
    Dim blnSign As Boolean
    Dim I As Integer

    DataType = &H31
    DataLength = 4
    DataAddress = StringToHexInt(Address)
    rc = GetData
    lngNum = 0

    If ByteBuffer(3) > 127 Then
        ByteBuffer(3) = ByteBuffer(3) - 128
        blnSign = True
    End If
    For I = 0 To 3
        lngNum = lngNum + (ByteBuffer(I) * 256 ^ I)
    Next I

    CopyMemory sFloat, lngNum, 4

    If blnSign Then
        sFloat = -sFloat
    End If

    ConvertToReal = sFloat
  
End Function
'Function ConvertToLong(ByteBuffer() As Byte) As Long
Function ConvertToLong(Address As String) As Long
    Dim I As Integer
    Dim S As String

    DataType = &H31
    DataLength = 2
    DataAddress = StringToHexInt(Address)
    rc = GetData

    rc = ByteSwap(ByteBuffer, 2)

    S = vbNullString
    For I = 0 To DataLength - 1
        S = S + Format$(Hex$(ByteBuffer(I)), "00")
    Next I
    ConvertToLong = Val(S)
End Function
Public Function GetMachineName() As String

    Dim plngSize As Long
    Dim pstrBuffer As String

    pstrBuffer = Space$(MAX_COMPUTERNAME_LENGTH + 1)

    plngSize = Len(pstrBuffer)

    If GetComputerName(pstrBuffer, plngSize) Then
        GetMachineName = Left$(pstrBuffer, plngSize)
    End If
  
End Function




