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

' <VB WATCH>
Const VBWMODULE = "PLCInterface"
' </VB WATCH>

Function NetWorkInitialize() As Long
           'return rc from pascal calls
           '  0 says ok
' <VB WATCH>
1          On Error GoTo vbwErrHandler
' </VB WATCH>

           ' if the network interface has already been opened, close it
           '
2          If NetworkOK = True Then
3              rc = PASCAL_HEICloseTransport(TP)
4              rc = PASCAL_HEIClose()
5          End If

           '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           ' Initialize the Ethernet Driver
           '
6          rc = PASCAL_HEIOpen(HEIAPIVersion)
7          If rc <> 0 Then
8              NetWorkInitialize = rc  'return error code
9              Exit Function
10         Else
               '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
               ' Initiaizize the Winsock protocol transport
               '
11             TP.Transport = HEIT_WINSOCK

12             TP.Protocol = HEIP_IP

13             rc = PASCAL_HEIOpenTransport(TP, HEIAPIVersion, 0)

14             NetWorkInitialize = rc

15         End If

' <VB WATCH>
16         Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "NetWorkInitialize"

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
'***********************************************************************
' Since the PASCAL_HEIxxxx calls require a byte buffer, convert the user
' entered strings to byte arrays
'
Function StringToByteArray(ByVal inString As String, ByRef Buffer() As Byte) As Integer
' <VB WATCH>
17         On Error GoTo vbwErrHandler
' </VB WATCH>
18     Dim I As Integer
19     Dim U() As Byte

           'Make sure all alpha characters are uppercase
20         U = StrConv(inString, vbUpperCase)

           'skip over the Unicode byte
21         For I = 0 To (Len(inString) - 1)
22             Buffer(I) = U(I * 2)
23         Next I

24         StringToByteArray = I

' <VB WATCH>
25         Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "StringToByteArray"

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

'***********************************************************************
' Swap successive entries in a byte array
'
Function ByteSwap(ByRef Buffer() As Byte, Count As Integer) As Integer
' <VB WATCH>
26         On Error GoTo vbwErrHandler
' </VB WATCH>
27     Dim I As Integer
28     Dim temp As Byte

29         For I = 0 To Count - 1 Step 2
30             temp = Buffer(I)
31             Buffer(I) = Buffer(I + 1)
32             Buffer(I + 1) = temp
33         Next I

34         ByteSwap = I

' <VB WATCH>
35         Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ByteSwap"

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

'************************************************************************
' Convert a byte array of character codes to a packed array of characters
'
Function HexConvert(ByRef Buffer() As Byte, Count As Integer) As Integer
' <VB WATCH>
36         On Error GoTo vbwErrHandler
' </VB WATCH>
37     Dim I As Integer

           'convert each character code
38         For I = 0 To (Count * 2) - 1

               'have to manually process HEX character digits
39             If (Buffer(I) > 64) And (Buffer(I) < 71) Then
40                 Select Case Buffer(I)
                       Case 65 'A
41                         Buffer(I) = 10
42                     Case 66 'B
43                         Buffer(I) = 11
44                     Case 67 'C
45                         Buffer(I) = 12
46                     Case 68 'D
47                         Buffer(I) = 13
48                     Case 69 'E
49                         Buffer(I) = 14
50                     Case 70 'F
51                         Buffer(I) = 15
52                 End Select

53             Else
                   'numeric digits are much easier
54                 Buffer(I) = ChrW$(Buffer(I))

55             End If

56         Next I

           'Now pack two HEX characters into a byte
57         Dim Z As Integer
58         Z = 0
59         For I = 0 To (Count * 2) - 1 Step 2
60             Buffer(Z) = (Buffer(I) * 16) + Buffer(I + 1)
61             Z = Z + 1
62         Next I

           'Now clear the remainder of the byte array - just to be neat and complete
63         For I = Z To (Count * 2) - 1
64             Buffer(I) = 0
65         Next I

66         HexConvert = Z

' <VB WATCH>
67         Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "HexConvert"

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

'***********************************************************************
' Brute force method of converting a 4 character string to a HEX number
'
Function StringToHexInt(inData As String) As Integer
' <VB WATCH>
68         On Error GoTo vbwErrHandler
' </VB WATCH>

69         Dim I As Integer, j As Integer
70         Dim t(4) As Byte

           'convert from octal

71         j = 0
72         For I = 1 To Len(inData)
73             j = j + Val(Mid$(inData, Len(inData) - I + 1, 1)) * (8 ^ (I - 1))
74         Next I

75         StringToHexInt = j + 1
76         Exit Function

77         inData = Hex$(j + 1)

78         I = StringToByteArray(inData, t)

           'convert each character code
79         For I = 0 To (Len(inData) - 1)

               'have to manually process HEX characters digits
80             If (t(I) > 64) And (t(I) < 71) Then
81                 Select Case t(I)
                       Case 65 'A
82                         t(I) = 10
83                     Case 66 'B
84                         t(I) = 11
85                     Case 67 'C
86                         t(I) = 12
87                     Case 68 'D
88                         t(I) = 13
89                     Case 69 'E
90                         t(I) = 14
91                     Case 70 'F
92                         t(I) = 15
93                 End Select

94             Else
                   'numeric digits are much easier
95                 t(I) = ChrW$(t(I))

96             End If
97         Next I

98         Select Case Len(inData)
               Case 0
99                 StringToHexInt = 0
100            Case 1
101                StringToHexInt = t(0)
102            Case 2
103                StringToHexInt = (t(0) * 16) + t(1)
104            Case 3
105                StringToHexInt = (t(0) * 256) + (t(1) * 16) + t(2)
106            Case 4
107                StringToHexInt = (t(0) * 4096) + (t(1) * 256) + (t(2) * 16) + t(3)
108        End Select

' <VB WATCH>
109        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "StringToHexInt"

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
Function ConnectToPLC(DeviceNo As Integer) As String
           ' Open the device
           '
' <VB WATCH>
110        On Error GoTo vbwErrHandler
' </VB WATCH>
111        rc = PASCAL_HEIOpenDevice(TP, aDevices(DeviceNo), HEIAPIVersion, DefDevTimeout, DefDevRetrys, False)
112        If rc <> 0 Then
113            DeviceOpen = False
114        Else
115            DeviceOpen = True
116        End If
117            ConnectToPLC = rc

' <VB WATCH>
118        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ConnectToPLC"

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

Function DisconnectPLC() As String
' <VB WATCH>
119        On Error GoTo vbwErrHandler
' </VB WATCH>
120        rc = PASCAL_HEICloseDevice(aDevices(tDevice))
121        DisconnectPLC = rc
' <VB WATCH>
122        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DisconnectPLC"

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
Function GetData() As String
' <VB WATCH>
123        On Error GoTo vbwErrHandler
' </VB WATCH>
124        GetData = PASCAL_HEICCMRequest(aDevices(tDevice), bWrite, DataType, DataAddress, DataLength, ByteBuffer(0))
' <VB WATCH>
125        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetData"

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
'Function ConvertToReal(ByteBuffer() As Byte) As Single
Function ConvertToReal(Address As String) As Single
' <VB WATCH>
126        On Error GoTo vbwErrHandler
' </VB WATCH>
127        Dim sFloat As Single
128        Dim lngNum As Long
129        Dim blnSign As Boolean
130        Dim I As Integer

131        DataType = &H31
132        DataLength = 4
133        DataAddress = StringToHexInt(Address)
134        rc = GetData
135        lngNum = 0

136        If ByteBuffer(3) > 127 Then
137            ByteBuffer(3) = ByteBuffer(3) - 128
138            blnSign = True
139        End If
140        For I = 0 To 3
141            lngNum = lngNum + (ByteBuffer(I) * 256 ^ I)
142        Next I

143        CopyMemory sFloat, lngNum, 4

144        If blnSign Then
145            sFloat = -sFloat
146        End If

147        ConvertToReal = sFloat

' <VB WATCH>
148        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ConvertToReal"

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
'Function ConvertToLong(ByteBuffer() As Byte) As Long
Function ConvertToLong(Address As String) As Long
' <VB WATCH>
149        On Error GoTo vbwErrHandler
' </VB WATCH>
150        Dim I As Integer
151        Dim S As String

152        DataType = &H31
153        DataLength = 2
154        DataAddress = StringToHexInt(Address)
155        rc = GetData

156        rc = ByteSwap(ByteBuffer, 2)

157        S = vbNullString
158        For I = 0 To DataLength - 1
159            S = S + Format$(Hex$(ByteBuffer(I)), "00")
160        Next I
161        ConvertToLong = Val(S)
' <VB WATCH>
162        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ConvertToLong"

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
Public Function GetMachineName() As String
' <VB WATCH>
163        On Error GoTo vbwErrHandler
' </VB WATCH>

164        Dim plngSize As Long
165        Dim pstrBuffer As String

166        pstrBuffer = Space$(MAX_COMPUTERNAME_LENGTH + 1)

167        plngSize = Len(pstrBuffer)

168        If GetComputerName(pstrBuffer, plngSize) Then
169            GetMachineName = Left$(pstrBuffer, plngSize)
170        End If

' <VB WATCH>
171        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetMachineName"

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





