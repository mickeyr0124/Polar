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
2          Const VBWPROCNAME = "PLCInterface.NetWorkInitialize"
3          If vbwProtector.vbwTraceProc Then
4              Dim vbwProtectorParameterString As String
5              If vbwProtector.vbwTraceParameters Then
6                  vbwProtectorParameterString = "()"
7              End If
8              vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
9          End If
' </VB WATCH>

           ' if the network interface has already been opened, close it
           '
10         If NetworkOK = True Then
11             rc = PASCAL_HEICloseTransport(TP)
12             rc = PASCAL_HEIClose()
13         End If

           '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
           ' Initialize the Ethernet Driver
           '
14         rc = PASCAL_HEIOpen(HEIAPIVersion)
15         If rc <> 0 Then
16             NetWorkInitialize = rc  'return error code
' <VB WATCH>
17         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
18             Exit Function
19         Else
               '''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
               ' Initiaizize the Winsock protocol transport
               '
20             TP.Transport = HEIT_WINSOCK

21             TP.Protocol = HEIP_IP

22             rc = PASCAL_HEIOpenTransport(TP, HEIAPIVersion, 0)

23             NetWorkInitialize = rc

24         End If

' <VB WATCH>
25         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
26         Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
'***********************************************************************
' Since the PASCAL_HEIxxxx calls require a byte buffer, convert the user
' entered strings to byte arrays
'
Function StringToByteArray(ByVal inString As String, ByRef Buffer() As Byte) As Integer
' <VB WATCH>
27         On Error GoTo vbwErrHandler
28         Const VBWPROCNAME = "PLCInterface.StringToByteArray"
29         If vbwProtector.vbwTraceProc Then
30             Dim vbwProtectorParameterString As String
31             If vbwProtector.vbwTraceParameters Then
32                 vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("inString", inString) & ", "
33                 vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("Buffer", Buffer) & ") "
34             End If
35             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
36         End If
' </VB WATCH>
37     Dim I As Integer
38     Dim U() As Byte

           'Make sure all alpha characters are uppercase
39         U = StrConv(inString, vbUpperCase)

           'skip over the Unicode byte
40         For I = 0 To (Len(inString) - 1)
41             Buffer(I) = U(I * 2)
42         Next I

43         StringToByteArray = I

' <VB WATCH>
44         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
45         Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "inString", inString
            vbwReportVariable "Buffer", Buffer
            vbwReportVariable "I", I
            vbwReportVariable "U", U
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

'***********************************************************************
' Swap successive entries in a byte array
'
Function ByteSwap(ByRef Buffer() As Byte, Count As Integer) As Integer
' <VB WATCH>
46         On Error GoTo vbwErrHandler
47         Const VBWPROCNAME = "PLCInterface.ByteSwap"
48         If vbwProtector.vbwTraceProc Then
49             Dim vbwProtectorParameterString As String
50             If vbwProtector.vbwTraceParameters Then
51                 vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Buffer", Buffer) & ", "
52                 vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("Count", Count) & ") "
53             End If
54             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
55         End If
' </VB WATCH>
56     Dim I As Integer
57     Dim temp As Byte

58         For I = 0 To Count - 1 Step 2
59             temp = Buffer(I)
60             Buffer(I) = Buffer(I + 1)
61             Buffer(I + 1) = temp
62         Next I

63         ByteSwap = I

' <VB WATCH>
64         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
65         Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Buffer", Buffer
            vbwReportVariable "Count", Count
            vbwReportVariable "I", I
            vbwReportVariable "temp", temp
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

'************************************************************************
' Convert a byte array of character codes to a packed array of characters
'
Function HexConvert(ByRef Buffer() As Byte, Count As Integer) As Integer
' <VB WATCH>
66         On Error GoTo vbwErrHandler
67         Const VBWPROCNAME = "PLCInterface.HexConvert"
68         If vbwProtector.vbwTraceProc Then
69             Dim vbwProtectorParameterString As String
70             If vbwProtector.vbwTraceParameters Then
71                 vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Buffer", Buffer) & ", "
72                 vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("Count", Count) & ") "
73             End If
74             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
75         End If
' </VB WATCH>
76     Dim I As Integer

           'convert each character code
77         For I = 0 To (Count * 2) - 1

               'have to manually process HEX character digits
78             If (Buffer(I) > 64) And (Buffer(I) < 71) Then
79                 Select Case Buffer(I)
                       Case 65 'A
80                         Buffer(I) = 10
81                     Case 66 'B
82                         Buffer(I) = 11
83                     Case 67 'C
84                         Buffer(I) = 12
85                     Case 68 'D
86                         Buffer(I) = 13
87                     Case 69 'E
88                         Buffer(I) = 14
89                     Case 70 'F
90                         Buffer(I) = 15
91                 End Select

92             Else
                   'numeric digits are much easier
93                 Buffer(I) = ChrW$(Buffer(I))

94             End If

95         Next I

           'Now pack two HEX characters into a byte
96         Dim Z As Integer
97         Z = 0
98         For I = 0 To (Count * 2) - 1 Step 2
99             Buffer(Z) = (Buffer(I) * 16) + Buffer(I + 1)
100            Z = Z + 1
101        Next I

           'Now clear the remainder of the byte array - just to be neat and complete
102        For I = Z To (Count * 2) - 1
103            Buffer(I) = 0
104        Next I

105        HexConvert = Z

' <VB WATCH>
106        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
107        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Buffer", Buffer
            vbwReportVariable "Count", Count
            vbwReportVariable "I", I
            vbwReportVariable "Z", Z
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

'***********************************************************************
' Brute force method of converting a 4 character string to a HEX number
'
Function StringToHexInt(inData As String) As Integer
' <VB WATCH>
108        On Error GoTo vbwErrHandler
109        Const VBWPROCNAME = "PLCInterface.StringToHexInt"
110        If vbwProtector.vbwTraceProc Then
111            Dim vbwProtectorParameterString As String
112            If vbwProtector.vbwTraceParameters Then
113                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("inData", inData) & ") "
114            End If
115            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
116        End If
' </VB WATCH>

117        Dim I As Integer, j As Integer
118        Dim t(4) As Byte

           'convert from octal

119        j = 0
120        For I = 1 To Len(inData)
121            j = j + Val(Mid$(inData, Len(inData) - I + 1, 1)) * (8 ^ (I - 1))
122        Next I

123        StringToHexInt = j + 1
' <VB WATCH>
124        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
125        Exit Function

126        inData = Hex$(j + 1)

127        I = StringToByteArray(inData, t)

           'convert each character code
128        For I = 0 To (Len(inData) - 1)

               'have to manually process HEX characters digits
129            If (t(I) > 64) And (t(I) < 71) Then
130                Select Case t(I)
                       Case 65 'A
131                        t(I) = 10
132                    Case 66 'B
133                        t(I) = 11
134                    Case 67 'C
135                        t(I) = 12
136                    Case 68 'D
137                        t(I) = 13
138                    Case 69 'E
139                        t(I) = 14
140                    Case 70 'F
141                        t(I) = 15
142                End Select

143            Else
                   'numeric digits are much easier
144                t(I) = ChrW$(t(I))

145            End If
146        Next I

147        Select Case Len(inData)
               Case 0
148                StringToHexInt = 0
149            Case 1
150                StringToHexInt = t(0)
151            Case 2
152                StringToHexInt = (t(0) * 16) + t(1)
153            Case 3
154                StringToHexInt = (t(0) * 256) + (t(1) * 16) + t(2)
155            Case 4
156                StringToHexInt = (t(0) * 4096) + (t(1) * 256) + (t(2) * 16) + t(3)
157        End Select

' <VB WATCH>
158        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
159        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "inData", inData
            vbwReportVariable "I", I
            vbwReportVariable "j", j
            vbwReportVariable "t", t
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Function ConnectToPLC(DeviceNo As Integer) As String
           ' Open the device
           '
' <VB WATCH>
160        On Error GoTo vbwErrHandler
161        Const VBWPROCNAME = "PLCInterface.ConnectToPLC"
162        If vbwProtector.vbwTraceProc Then
163            Dim vbwProtectorParameterString As String
164            If vbwProtector.vbwTraceParameters Then
165                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("DeviceNo", DeviceNo) & ") "
166            End If
167            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
168        End If
' </VB WATCH>
169        rc = PASCAL_HEIOpenDevice(TP, aDevices(DeviceNo), HEIAPIVersion, DefDevTimeout, DefDevRetrys, False)
170        If rc <> 0 Then
171            DeviceOpen = False
172        Else
173            DeviceOpen = True
174        End If
175            ConnectToPLC = rc

' <VB WATCH>
176        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
177        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "DeviceNo", DeviceNo
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function DisconnectPLC() As String
' <VB WATCH>
178        On Error GoTo vbwErrHandler
179        Const VBWPROCNAME = "PLCInterface.DisconnectPLC"
180        If vbwProtector.vbwTraceProc Then
181            Dim vbwProtectorParameterString As String
182            If vbwProtector.vbwTraceParameters Then
183                vbwProtectorParameterString = "()"
184            End If
185            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
186        End If
' </VB WATCH>
187        rc = PASCAL_HEICloseDevice(aDevices(tDevice))
188        DisconnectPLC = rc
' <VB WATCH>
189        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
190        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Function GetData() As String
' <VB WATCH>
191        On Error GoTo vbwErrHandler
192        Const VBWPROCNAME = "PLCInterface.GetData"
193        If vbwProtector.vbwTraceProc Then
194            Dim vbwProtectorParameterString As String
195            If vbwProtector.vbwTraceParameters Then
196                vbwProtectorParameterString = "()"
197            End If
198            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
199        End If
' </VB WATCH>
200        GetData = PASCAL_HEICCMRequest(aDevices(tDevice), bWrite, DataType, DataAddress, DataLength, ByteBuffer(0))
' <VB WATCH>
201        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
202        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
'Function ConvertToReal(ByteBuffer() As Byte) As Single
Function ConvertToReal(Address As String) As Single
' <VB WATCH>
203        On Error GoTo vbwErrHandler
204        Const VBWPROCNAME = "PLCInterface.ConvertToReal"
205        If vbwProtector.vbwTraceProc Then
206            Dim vbwProtectorParameterString As String
207            If vbwProtector.vbwTraceParameters Then
208                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Address", Address) & ") "
209            End If
210            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
211        End If
' </VB WATCH>
212        Dim sFloat As Single
213        Dim lngNum As Long
214        Dim blnSign As Boolean
215        Dim I As Integer

216        DataType = &H31
217        DataLength = 4
218        DataAddress = StringToHexInt(Address)
219        rc = GetData
220        lngNum = 0

221        If ByteBuffer(3) > 127 Then
222            ByteBuffer(3) = ByteBuffer(3) - 128
223            blnSign = True
224        End If
225        For I = 0 To 3
226            lngNum = lngNum + (ByteBuffer(I) * 256 ^ I)
227        Next I

228        CopyMemory sFloat, lngNum, 4

229        If blnSign Then
230            sFloat = -sFloat
231        End If

232        ConvertToReal = sFloat

' <VB WATCH>
233        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
234        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Address", Address
            vbwReportVariable "sFloat", sFloat
            vbwReportVariable "lngNum", lngNum
            vbwReportVariable "blnSign", blnSign
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
'Function ConvertToLong(ByteBuffer() As Byte) As Long
Function ConvertToLong(Address As String) As Long
' <VB WATCH>
235        On Error GoTo vbwErrHandler
236        Const VBWPROCNAME = "PLCInterface.ConvertToLong"
237        If vbwProtector.vbwTraceProc Then
238            Dim vbwProtectorParameterString As String
239            If vbwProtector.vbwTraceParameters Then
240                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Address", Address) & ") "
241            End If
242            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
243        End If
' </VB WATCH>
244        Dim I As Integer
245        Dim S As String

246        DataType = &H31
247        DataLength = 2
248        DataAddress = StringToHexInt(Address)
249        rc = GetData

250        rc = ByteSwap(ByteBuffer, 2)

251        S = vbNullString
252        For I = 0 To DataLength - 1
253            S = S + Format$(Hex$(ByteBuffer(I)), "00")
254        Next I
255        ConvertToLong = Val(S)
' <VB WATCH>
256        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
257        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "Address", Address
            vbwReportVariable "I", I
            vbwReportVariable "S", S
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Public Function GetMachineName() As String
' <VB WATCH>
258        On Error GoTo vbwErrHandler
259        Const VBWPROCNAME = "PLCInterface.GetMachineName"
260        If vbwProtector.vbwTraceProc Then
261            Dim vbwProtectorParameterString As String
262            If vbwProtector.vbwTraceParameters Then
263                vbwProtectorParameterString = "()"
264            End If
265            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
266        End If
' </VB WATCH>

267        Dim plngSize As Long
268        Dim pstrBuffer As String

269        pstrBuffer = Space$(MAX_COMPUTERNAME_LENGTH + 1)

270        plngSize = Len(pstrBuffer)

271        If GetComputerName(pstrBuffer, plngSize) Then
272            GetMachineName = Left$(pstrBuffer, plngSize)
273        End If

' <VB WATCH>
274        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
275        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "plngSize", plngSize
            vbwReportVariable "pstrBuffer", pstrBuffer
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
Private Sub vbwReport_PLCInterface_Encryption(lName As String, lUDT As PLCInterface.Encryption, Optional ByVal lTab As Long)
    vbwReportVariable lName & ".Algorithm", lUDT.Algorithm, lTab
    vbwReportVariable lName & ".Unused1", lUDT.Unused1, lTab
    vbwReportVariable lName & ".Key", lUDT.Key, lTab
End Sub
Private Sub vbwReport_PLCInterface_EnetAddress(lName As String, lUDT As PLCInterface.EnetAddress, Optional ByVal lTab As Long)
    vbwReportVariable lName & ".Address", lUDT.Address, lTab
End Sub
Private Sub vbwReport_PLCInterface_HEITransport(lName As String, lUDT As PLCInterface.HEITransport, Optional ByVal lTab As Long)
    vbwReportVariable lName & ".Transport", lUDT.Transport, lTab
    vbwReportVariable lName & ".Protocol", lUDT.Protocol, lTab
    vbwReport_PLCInterface_Encryption lName & ".Encrypt", lUDT.Encrypt, lTab
    vbwReport_PLCInterface_EnetAddress lName & ".SourceAddress", lUDT.SourceAddress, lTab
    vbwReportVariable lName & ".Reserved", lUDT.Reserved, lTab
End Sub
Private Sub vbwReport_PLCInterface_HEIDevice(lName As String, lUDT As PLCInterface.HEIDevice, Optional ByVal lTab As Long)
    vbwReportVariable lName & ".Address", lUDT.Address, lTab
End Sub
Private Sub vbwReport_PLCInterface_HEIDevice_Array(lName As String, lArray() As PLCInterface.HEIDevice, Optional ByVal lTab As Long)

    Dim i As Long

    ' report each member '
    vbwReportToFile String$(lTab, vbTab) & lName & VBW_TYPE_STRING
    For i = LBound(lArray) To UBound(lArray)
        vbwReport_PLCInterface_HEIDevice lName & "(" & i & ")", lArray(i), lTab + 1
    Next i

End Sub
  

Private Sub vbwReportModuleVariables()
    vbwReportToFile VBW_MODULE_STRING
End Sub
' </VB WATCH>
