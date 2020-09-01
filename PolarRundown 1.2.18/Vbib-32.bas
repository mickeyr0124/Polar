Attribute VB_Name = "VBIB32"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 32-bit Visual Basic Language Interface
' Version 1.81
' Copyright 2001 National Instruments Corporation.
' All Rights Reserved.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   This module contains the subroutine declarations,
'   function declarations and constants required to use
'   the National Instruments GPIB Dynamic Link Library
'   (DLL) for controlling IEEE-488 instrumentation.  This
'   file must be 'added' to your Visual Basic project
'   (by choosing Add File from the File menu or pressing
'   CTRL+F12) so that you can access the NI-488.2
'   subroutines and functions.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   NI-488.2 DLL entry function declarations

Declare Function ibask32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibask" (ByVal ud As Long, ByVal opt As Long, value As Long) As Long
Declare Function ibbna32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibbnaA" (ByVal ud As Long, sstr As Any) As Long
Declare Function ibcac32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibcac" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibclr32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibclr" (ByVal ud As Long) As Long
Declare Function ibcmd32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibcmd" (ByVal ud As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibcmda32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibcmda" (ByVal ud As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibconfig32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibconfig" (ByVal ud As Long, ByVal opt As Long, ByVal v As Long) As Long
Declare Function ibdev32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibdev" (ByVal bdid As Long, ByVal pad As Long, ByVal sad As Long, ByVal tmo As Long, ByVal eot As Long, ByVal eos As Long) As Long
Declare Function ibdma32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibdma" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibeos32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibeos" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibeot32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibeot" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibfind32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibfindA" (sstr As Any) As Long
Declare Function ibgts32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibgts" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibist32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibist" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function iblck32 Lib "c:\windows\system32\gpib-32.dll" Alias "iblck" (ByVal ud As Long, ByVal v As Long, ByVal LockWaitTime As Long, arg1 As Any) As Long
Declare Function iblines32 Lib "c:\windows\system32\gpib-32.dll" Alias "iblines" (ByVal ud As Long, v As Long) As Long
Declare Function ibln32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibln" (ByVal ud As Long, ByVal pad As Long, ByVal sad As Long, ln As Long) As Long
Declare Function ibloc32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibloc" (ByVal ud As Long) As Long
Declare Function iblock32 Lib "c:\windows\system32\gpib-32.dll" Alias "iblock" (ByVal ud As Long) As Long
Declare Function ibonl32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibonl" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibpad32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibpad" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibpct32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibpct" (ByVal ud As Long) As Long
Declare Function ibppc32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibppc" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibrd32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibrd" (ByVal ud As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibrda32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibrda" (ByVal ud As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibrdf32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibrdfA" (ByVal ud As Long, sstr As Any) As Long
Declare Function ibrpp32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibrpp" (ByVal ud As Long, sstr As Any) As Long
Declare Function ibrsc32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibrsc" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibrsp32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibrsp" (ByVal ud As Long, sstr As Any) As Long
Declare Function ibrsv32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibrsv" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibsad32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibsad" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibsic32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibsic" (ByVal ud As Long) As Long
Declare Function ibsre32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibsre" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibstop32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibstop" (ByVal ud As Long) As Long
Declare Function ibtmo32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibtmo" (ByVal ud As Long, ByVal v As Long) As Long
Declare Function ibtrg32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibtrg" (ByVal ud As Long) As Long
Declare Function ibunlock32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibunlock" (ByVal ud As Long) As Long
Declare Function ibwait32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibwait" (ByVal ud As Long, ByVal mask As Long) As Long
Declare Function ibwrt32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibwrt" (ByVal ud As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibwrta32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibwrta" (ByVal ud As Long, sstr As Any, ByVal cnt As Long) As Long
Declare Function ibwrtf32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibwrtfA" (ByVal ud As Long, sstr As Any) As Long
Declare Sub AllSpoll32 Lib "c:\windows\system32\gpib-32.dll" Alias "AllSpoll" (ByVal boardID As Long, arg1 As Any, arg2 As Any)
Declare Sub DevClear32 Lib "c:\windows\system32\gpib-32.dll" Alias "DevClear" (ByVal boardID As Long, ByVal v As Long)
Declare Sub DevClearList32 Lib "c:\windows\system32\gpib-32.dll" Alias "DevClearList" (ByVal boardID As Long, arg1 As Any)
Declare Sub EnableLocal32 Lib "c:\windows\system32\gpib-32.dll" Alias "EnableLocal" (ByVal boardID As Long, arg1 As Any)
Declare Sub EnableRemote32 Lib "c:\windows\system32\gpib-32.dll" Alias "EnableRemote" (ByVal boardID As Long, arg1 As Any)
Declare Sub FindLstn32 Lib "c:\windows\system32\gpib-32.dll" Alias "FindLstn" (ByVal boardID As Long, arg1 As Any, arg2 As Any, ByVal limit As Long)
Declare Sub FindRQS32 Lib "c:\windows\system32\gpib-32.dll" Alias "FindRQS" (ByVal boardID As Long, arg1 As Any, result As Long)
Declare Sub PassControl32 Lib "c:\windows\system32\gpib-32.dll" Alias "PassControl" (ByVal boardID As Long, ByVal addr As Long)
Declare Sub PPoll32 Lib "c:\windows\system32\gpib-32.dll" Alias "PPoll" (ByVal boardID As Long, result As Long)
Declare Sub PPollConfig32 Lib "c:\windows\system32\gpib-32.dll" Alias "PPollConfig" (ByVal boardID As Long, ByVal addr As Long, ByVal line As Long, ByVal sense As Long)
Declare Sub PPollUnconfig32 Lib "c:\windows\system32\gpib-32.dll" Alias "PPollUnconfig" (ByVal boardID As Long, arg1 As Any)
Declare Sub RcvRespMsg32 Lib "c:\windows\system32\gpib-32.dll" Alias "RcvRespMsg" (ByVal boardID As Long, arg1 As Any, ByVal cnt As Long, ByVal term As Long)
Declare Sub ReadStatusByte32 Lib "c:\windows\system32\gpib-32.dll" Alias "ReadStatusByte" (ByVal boardID As Long, ByVal addr As Long, result As Long)
Declare Sub Receive32 Lib "c:\windows\system32\gpib-32.dll" Alias "Receive" (ByVal boardID As Long, ByVal addr As Long, arg1 As Any, ByVal cnt As Long, ByVal term As Long)
Declare Sub ReceiveSetup32 Lib "c:\windows\system32\gpib-32.dll" Alias "ReceiveSetup" (ByVal boardID As Long, ByVal addr As Long)
Declare Sub ResetSys32 Lib "c:\windows\system32\gpib-32.dll" Alias "ResetSys" (ByVal boardID As Long, arg1 As Any)
Declare Sub Send32 Lib "c:\windows\system32\gpib-32.dll" Alias "Send" (ByVal boardID As Long, ByVal addr As Long, sstr As Any, ByVal cnt As Long, ByVal term As Long)
Declare Sub SendCmds32 Lib "c:\windows\system32\gpib-32.dll" Alias "SendCmds" (ByVal boardID As Long, sstr As Any, ByVal cnt As Long)
Declare Sub SendDataBytes32 Lib "c:\windows\system32\gpib-32.dll" Alias "SendDataBytes" (ByVal boardID As Long, sstr As Any, ByVal cnt As Long, ByVal term As Long)
Declare Sub SendIFC32 Lib "c:\windows\system32\gpib-32.dll" Alias "SendIFC" (ByVal boardID As Long)
Declare Sub SendList32 Lib "c:\windows\system32\gpib-32.dll" Alias "SendList" (ByVal boardID As Long, arg1 As Any, arg2 As Any, ByVal cnt As Long, ByVal term As Long)
Declare Sub SendLLO32 Lib "c:\windows\system32\gpib-32.dll" Alias "SendLLO" (ByVal boardID As Long)
Declare Sub SendSetup32 Lib "c:\windows\system32\gpib-32.dll" Alias "SendSetup" (ByVal boardID As Long, arg1 As Any)
Declare Sub SetRWLS32 Lib "c:\windows\system32\gpib-32.dll" Alias "SetRWLS" (ByVal boardID As Long, arg1 As Any)
Declare Sub TestSys32 Lib "c:\windows\system32\gpib-32.dll" Alias "TestSys" (ByVal boardID As Long, arg1 As Any, arg2 As Any)
Declare Sub Trigger32 Lib "c:\windows\system32\gpib-32.dll" Alias "Trigger" (ByVal boardID As Long, ByVal addr As Long)
Declare Sub TriggerList32 Lib "c:\windows\system32\gpib-32.dll" Alias "TriggerList" (ByVal boardID As Long, arg1 As Any)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   DLL entry function declarations needed for GPIB global variables

Declare Function RegisterGpibGlobalsForThread Lib "c:\windows\system32\gpib-32.dll" (Longibsta As Long, Longiberr As Long, Longibcnt As Long, ibcntl As Long) As Long
Declare Function UnregisterGpibGlobalsForThread Lib "c:\windows\system32\gpib-32.dll" () As Long
Declare Function ThreadIbsta32 Lib "c:\windows\system32\gpib-32.dll" Alias "ThreadIbsta" () As Long
Declare Function ThreadIbcnt32 Lib "c:\windows\system32\gpib-32.dll" Alias "ThreadIbcnt" () As Long
Declare Function ThreadIbcntl32 Lib "c:\windows\system32\gpib-32.dll" Alias "ThreadIbcntl" () As Long
Declare Function ThreadIberr32 Lib "c:\windows\system32\gpib-32.dll" Alias "ThreadIberr" () As Long

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   DLL entry function declarations needed for GPIBnotify OLE control

Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   DLL entry function declarations needed for GPIB-ENET functions

Declare Function iblockx32 Lib "c:\windows\system32\gpib-32.dll" Alias "iblockxA" (ByVal ud As Long, ByVal LockWaitTime As Long, arg1 As Any) As Long
Declare Function ibunlockx32 Lib "c:\windows\system32\gpib-32.dll" Alias "ibunlockx" (ByVal ud As Long) As Long


' <VB WATCH>
Const VBWMODULE = "VBIB32"
' </VB WATCH>

Sub AllSpoll(ByVal boardID As Integer, addrs() As Integer, results() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1          On Error GoTo vbwErrHandler
' </VB WATCH>
2          If (GPIBglobalsRegistered = 0) Then
3            Call RegisterGPIBGlobals
4          End If

       ' Call the 32-bit DLL.
5          Call AllSpoll32(boardID, addrs(0), results(0))

6          Call copy_ibvars
' <VB WATCH>
7          Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "AllSpoll"

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

Sub copy_ibvars()
' <VB WATCH>
8          On Error GoTo vbwErrHandler
' </VB WATCH>
9          ibsta = ConvertLongToInt(Longibsta)
10         iberr = CInt(Longiberr)
11         ibcnt = ConvertLongToInt(ibcntl)
' <VB WATCH>
12         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "copy_ibvars"

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

Sub DevClear(ByVal boardID As Integer, ByVal addr As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
13         On Error GoTo vbwErrHandler
' </VB WATCH>
14         If (GPIBglobalsRegistered = 0) Then
15           Call RegisterGPIBGlobals
16         End If

       ' Call the 32-bit DLL.
17         Call DevClear32(boardID, addr)

18         Call copy_ibvars
' <VB WATCH>
19         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DevClear"

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

Sub DevClearList(ByVal boardID As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
20         On Error GoTo vbwErrHandler
' </VB WATCH>
21         If (GPIBglobalsRegistered = 0) Then
22           Call RegisterGPIBGlobals
23         End If

       ' Call the 32-bit DLL.
24         Call DevClearList32(boardID, addrs(0))

25         Call copy_ibvars
' <VB WATCH>
26         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DevClearList"

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

Sub EnableLocal(ByVal boardID As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
27         On Error GoTo vbwErrHandler
' </VB WATCH>
28         If (GPIBglobalsRegistered = 0) Then
29           Call RegisterGPIBGlobals
30         End If

       ' Call the 32-bit DLL.
31         Call EnableLocal32(boardID, addrs(0))

32         Call copy_ibvars
' <VB WATCH>
33         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "EnableLocal"

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

Sub EnableRemote(ByVal boardID As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
34         On Error GoTo vbwErrHandler
' </VB WATCH>
35         If (GPIBglobalsRegistered = 0) Then
36           Call RegisterGPIBGlobals
37         End If

       ' Call the 32-bit DLL.
38         Call EnableRemote32(boardID, addrs(0))

39         Call copy_ibvars
' <VB WATCH>
40         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "EnableRemote"

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

Sub FindLstn(ByVal boardID As Integer, addrs() As Integer, results() As Integer, ByVal limit As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
41         On Error GoTo vbwErrHandler
' </VB WATCH>
42         If (GPIBglobalsRegistered = 0) Then
43           Call RegisterGPIBGlobals
44         End If

       ' Call the 32-bit DLL.
45         Call FindLstn32(boardID, addrs(0), results(0), limit)

46         Call copy_ibvars
' <VB WATCH>
47         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FindLstn"

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

Sub FindRQS(ByVal boardID As Integer, addrs() As Integer, result As Integer)
' <VB WATCH>
48         On Error GoTo vbwErrHandler
' </VB WATCH>
49        Dim tmpresult As Long

       ' Check to see if GPIB Global variables are registered
50         If (GPIBglobalsRegistered = 0) Then
51           Call RegisterGPIBGlobals
52         End If

       ' Call the 32-bit DLL.
53         Call FindRQS32(boardID, addrs(0), tmpresult)

54         result = ConvertLongToInt(tmpresult)

55         Call copy_ibvars
' <VB WATCH>
56         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FindRQS"

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

Sub ibask(ByVal ud As Integer, ByVal opt As Integer, rval As Integer)
' <VB WATCH>
57         On Error GoTo vbwErrHandler
' </VB WATCH>
58       Dim tmprval As Long

       ' Check to see if GPIB Global variables are registered
59         If (GPIBglobalsRegistered = 0) Then
60           Call RegisterGPIBGlobals
61         End If

       ' Call the 32-bit DLL.
62         Call ibask32(ud, opt, tmprval)

63         rval = ConvertLongToInt(tmprval)

64         Call copy_ibvars
' <VB WATCH>
65         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibask"

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

Sub ibbna(ByVal ud As Integer, ByVal udname As String)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
66         On Error GoTo vbwErrHandler
' </VB WATCH>
67         If (GPIBglobalsRegistered = 0) Then
68           Call RegisterGPIBGlobals
69         End If

       ' Call the 32-bit DLL.
70         Call ibbna32(ud, ByVal udname)

71         Call copy_ibvars
' <VB WATCH>
72         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibbna"

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

Sub ibcac(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
73         On Error GoTo vbwErrHandler
' </VB WATCH>
74         If (GPIBglobalsRegistered = 0) Then
75           Call RegisterGPIBGlobals
76         End If

       ' Call the 32-bit DLL.
77         Call ibcac32(ud, v)

78         Call copy_ibvars
' <VB WATCH>
79         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibcac"

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

Sub ibclr(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
80         On Error GoTo vbwErrHandler
' </VB WATCH>
81         If (GPIBglobalsRegistered = 0) Then
82           Call RegisterGPIBGlobals
83         End If

       ' Call the 32-bit DLL.
84         Call ibclr32(ud)

85         Call copy_ibvars
' <VB WATCH>
86         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibclr"

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

Sub ibcmd(ByVal ud As Integer, ByVal buf As String)
' <VB WATCH>
87         On Error GoTo vbwErrHandler
' </VB WATCH>
88        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
89         If (GPIBglobalsRegistered = 0) Then
90           Call RegisterGPIBGlobals
91         End If

92         cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
93         Call ibcmd32(ud, ByVal buf, cnt)

94         Call copy_ibvars
' <VB WATCH>
95         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibcmd"

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

Sub ibcmda(ByVal ud As Integer, ByVal buf As String)
' <VB WATCH>
96         On Error GoTo vbwErrHandler
' </VB WATCH>
97         Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
98         If (GPIBglobalsRegistered = 0) Then
99           Call RegisterGPIBGlobals
100        End If

101        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
102        Call ibcmd32(ud, ByVal buf, cnt)

       ' When Visual Basic remapping buffer problem solved, then use:
       '    call ibcmda32(ud, ByVal buf, cnt)

103        Call copy_ibvars
' <VB WATCH>
104        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibcmda"

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

Sub ibconfig(ByVal bdid As Integer, ByVal opt As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
105        On Error GoTo vbwErrHandler
' </VB WATCH>
106        If (GPIBglobalsRegistered = 0) Then
107          Call RegisterGPIBGlobals
108        End If

       ' Call the 32-bit DLL.
109        Call ibconfig32(bdid, opt, v)

110        Call copy_ibvars
' <VB WATCH>
111        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibconfig"

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

Sub ibdev(ByVal bdid As Integer, ByVal pad As Integer, ByVal sad As Integer, ByVal tmo As Integer, ByVal eot As Integer, ByVal eos As Integer, ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
112        On Error GoTo vbwErrHandler
' </VB WATCH>
113        If (GPIBglobalsRegistered = 0) Then
114          Call RegisterGPIBGlobals
115        End If

       ' Call the 32-bit DLL.
116        ud = ConvertLongToInt(ibdev32(bdid, pad, sad, tmo, eot, eos))

117        Call copy_ibvars
' <VB WATCH>
118        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibdev"

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

Sub ibdma(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
119        On Error GoTo vbwErrHandler
' </VB WATCH>
120        If (GPIBglobalsRegistered = 0) Then
121          Call RegisterGPIBGlobals
122        End If

       ' Call the 32-bit DLL.
123        Call ibdma32(ud, v)

124        Call copy_ibvars
' <VB WATCH>
125        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibdma"

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

Sub ibeos(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
126        On Error GoTo vbwErrHandler
' </VB WATCH>
127        If (GPIBglobalsRegistered = 0) Then
128          Call RegisterGPIBGlobals
129        End If

       ' Call the 32-bit DLL.
130        Call ibeos32(ud, v)

131        Call copy_ibvars
' <VB WATCH>
132        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibeos"

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

Sub ibeot(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
133        On Error GoTo vbwErrHandler
' </VB WATCH>
134        If (GPIBglobalsRegistered = 0) Then
135          Call RegisterGPIBGlobals
136        End If

       ' Call the 32-bit DLL.
137        Call ibeot32(ud, v)

138        Call copy_ibvars
' <VB WATCH>
139        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibeot"

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

Sub ibfind(ByVal udname As String, ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
140        On Error GoTo vbwErrHandler
' </VB WATCH>
141        If (GPIBglobalsRegistered = 0) Then
142          Call RegisterGPIBGlobals
143        End If

       ' Call the 32-bit DLL.
144        ud = ConvertLongToInt(ibfind32(ByVal udname))

145        Call copy_ibvars
' <VB WATCH>
146        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibfind"

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

Sub ibgts(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
147        On Error GoTo vbwErrHandler
' </VB WATCH>
148        If (GPIBglobalsRegistered = 0) Then
149          Call RegisterGPIBGlobals
150        End If

       ' Call the 32-bit DLL.
151        Call ibgts32(ud, v)

152        Call copy_ibvars
' <VB WATCH>
153        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibgts"

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

Sub ibist(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
154        On Error GoTo vbwErrHandler
' </VB WATCH>
155        If (GPIBglobalsRegistered = 0) Then
156          Call RegisterGPIBGlobals
157        End If

       ' Call the 32-bit DLL.
158        Call ibist32(ud, v)

159        Call copy_ibvars
' <VB WATCH>
160        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibist"

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

Sub iblines(ByVal ud As Integer, lines As Integer)
' <VB WATCH>
161        On Error GoTo vbwErrHandler
' </VB WATCH>
162       Dim tmplines As Long

       ' Check to see if GPIB Global variables are registered
163        If (GPIBglobalsRegistered = 0) Then
164          Call RegisterGPIBGlobals
165        End If

       ' Call the 32-bit DLL.
166        Call iblines32(ud, tmplines)

167        lines = ConvertLongToInt(tmplines)

168        Call copy_ibvars
' <VB WATCH>
169        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "iblines"

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

Sub ibln(ByVal ud As Integer, ByVal pad As Integer, ByVal sad As Integer, ln As Integer)
' <VB WATCH>
170        On Error GoTo vbwErrHandler
' </VB WATCH>
171        Dim tmpln As Long

       ' Check to see if GPIB Global variables are registered
172        If (GPIBglobalsRegistered = 0) Then
173          Call RegisterGPIBGlobals
174        End If

       ' Call the 32-bit DLL.
175        Call ibln32(ud, pad, sad, tmpln)

176        ln = ConvertLongToInt(tmpln)

177        Call copy_ibvars
' <VB WATCH>
178        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibln"

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

Sub ibloc(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
179        On Error GoTo vbwErrHandler
' </VB WATCH>
180        If (GPIBglobalsRegistered = 0) Then
181          Call RegisterGPIBGlobals
182        End If

       ' Call the 32-bit DLL.
183        Call ibloc32(ud)

184        Call copy_ibvars
' <VB WATCH>
185        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibloc"

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

Sub iblck(ByVal ud As Integer, ByVal v As Integer, ByVal LockWaitTime As Long)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
186        On Error GoTo vbwErrHandler
' </VB WATCH>
187        If (GPIBglobalsRegistered = 0) Then
188          Call RegisterGPIBGlobals
189        End If

       ' Call the 32-bit DLL.
190        Call iblck32(ud, v, LockWaitTime, ByVal 0)

191        Call copy_ibvars
' <VB WATCH>
192        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "iblck"

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

Sub ibonl(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
193        On Error GoTo vbwErrHandler
' </VB WATCH>
194        If (GPIBglobalsRegistered = 0) Then
195          Call RegisterGPIBGlobals
196        End If

       ' Call the 32-bit DLL.
197        Call ibonl32(ud, v)

198        Call copy_ibvars
' <VB WATCH>
199        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibonl"

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

Sub ibpad(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
200        On Error GoTo vbwErrHandler
' </VB WATCH>
201        If (GPIBglobalsRegistered = 0) Then
202          Call RegisterGPIBGlobals
203        End If

       ' Call the 32-bit DLL.
204        Call ibpad32(ud, v)

205        Call copy_ibvars
' <VB WATCH>
206        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibpad"

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

Sub ibpct(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
207        On Error GoTo vbwErrHandler
' </VB WATCH>
208        If (GPIBglobalsRegistered = 0) Then
209          Call RegisterGPIBGlobals
210        End If

       ' Call the 32-bit DLL.
211        Call ibpct32(ud)

212        Call copy_ibvars
' <VB WATCH>
213        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibpct"

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

Sub ibppc(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
214        On Error GoTo vbwErrHandler
' </VB WATCH>
215        If (GPIBglobalsRegistered = 0) Then
216          Call RegisterGPIBGlobals
217        End If

       ' Call the 32-bit DLL.
218        Call ibppc32(ud, v)

219        Call copy_ibvars
' <VB WATCH>
220        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibppc"

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

Sub ibrd(ByVal ud As Integer, buf As String)
' <VB WATCH>
221        On Error GoTo vbwErrHandler
' </VB WATCH>
222        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
223        If (GPIBglobalsRegistered = 0) Then
224          Call RegisterGPIBGlobals
225        End If

226        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
227        Call ibrd32(ud, ByVal buf, cnt)

228        Call copy_ibvars
' <VB WATCH>
229        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibrd"

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

Sub ibrda(ByVal ud As Integer, buf As String)
' <VB WATCH>
230        On Error GoTo vbwErrHandler
' </VB WATCH>
231        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
232        If (GPIBglobalsRegistered = 0) Then
233          Call RegisterGPIBGlobals
234        End If

235        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
236        Call ibrd32(ud, ByVal buf, cnt)

       ' When Visual Basic remapping buffer problem solved, use this:
       '    Call ibrda32(ud, ByVal buf, cnt)

237        Call copy_ibvars
' <VB WATCH>
238        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibrda"

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

Sub ibrdf(ByVal ud As Integer, ByVal filename As String)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
239        On Error GoTo vbwErrHandler
' </VB WATCH>
240        If (GPIBglobalsRegistered = 0) Then
241          Call RegisterGPIBGlobals
242        End If

       ' Call the 32-bit DLL.
243        Call ibrdf32(ud, ByVal filename)

244        Call copy_ibvars
' <VB WATCH>
245        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibrdf"

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

Sub ibrdi(ByVal ud As Integer, ibuf() As Integer, ByVal cnt As Long)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
246        On Error GoTo vbwErrHandler
' </VB WATCH>
247        If (GPIBglobalsRegistered = 0) Then
248          Call RegisterGPIBGlobals
249        End If

       ' Call the 32-bit DLL.
250        Call ibrd32(ud, ibuf(0), cnt)

251        Call copy_ibvars
' <VB WATCH>
252        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibrdi"

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

Sub ibrdia(ByVal ud As Integer, ibuf() As Integer, ByVal cnt As Long)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
253        On Error GoTo vbwErrHandler
' </VB WATCH>
254        If (GPIBglobalsRegistered = 0) Then
255          Call RegisterGPIBGlobals
256        End If

       ' Call the 32-bit DLL.
257        Call ibrd32(ud, ibuf(0), cnt)

       ' When Visual Basic remapping buffer problem is solved, then use:
       '    Call ibrda32(u, ibuf(0), cnt)

258        Call copy_ibvars
' <VB WATCH>
259        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibrdia"

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

Sub ibrpp(ByVal ud As Integer, ppr As Integer)
' <VB WATCH>
260        On Error GoTo vbwErrHandler
' </VB WATCH>
261        Static tmp_str As String * 2

       ' Check to see if GPIB Global variables are registered
262        If (GPIBglobalsRegistered = 0) Then
263          Call RegisterGPIBGlobals
264        End If

       ' Call the 32-bit DLL.
265        Call ibrpp32(ud, ByVal tmp_str)

266        ppr = Asc(tmp_str)

267        Call copy_ibvars
' <VB WATCH>
268        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibrpp"

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

Sub ibrsc(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
269        On Error GoTo vbwErrHandler
' </VB WATCH>
270        If (GPIBglobalsRegistered = 0) Then
271          Call RegisterGPIBGlobals
272        End If

       ' Call the 32-bit DLL.
273        Call ibrsc32(ud, v)

274        Call copy_ibvars
' <VB WATCH>
275        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibrsc"

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

Sub ibrsp(ByVal ud As Integer, spr As Integer)
' <VB WATCH>
276        On Error GoTo vbwErrHandler
' </VB WATCH>
277        Static tmp_str As String * 2

       ' Check to see if GPIB Global variables are registered
278        If (GPIBglobalsRegistered = 0) Then
279          Call RegisterGPIBGlobals
280        End If

       ' Call the 32-bit DLL
281        Call ibrsp32(ud, ByVal tmp_str)

282        spr = Asc(tmp_str)

283        Call copy_ibvars
' <VB WATCH>
284        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibrsp"

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

Sub ibrsv(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
285        On Error GoTo vbwErrHandler
' </VB WATCH>
286        If (GPIBglobalsRegistered = 0) Then
287          Call RegisterGPIBGlobals
288        End If

       ' Call the 32-bit DLL.
289        Call ibrsv32(ud, v)

290        Call copy_ibvars
' <VB WATCH>
291        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibrsv"

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

Sub ibsad(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
292        On Error GoTo vbwErrHandler
' </VB WATCH>
293        If (GPIBglobalsRegistered = 0) Then
294          Call RegisterGPIBGlobals
295        End If

       ' Call the 32-bit DLL.
296        Call ibsad32(ud, v)

297        Call copy_ibvars
' <VB WATCH>
298        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibsad"

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

Sub ibsic(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
299        On Error GoTo vbwErrHandler
' </VB WATCH>
300        If (GPIBglobalsRegistered = 0) Then
301          Call RegisterGPIBGlobals
302        End If

       ' Call the 32-bit DLL.
303        Call ibsic32(ud)

304        Call copy_ibvars
' <VB WATCH>
305        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibsic"

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

Sub ibsre(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
306        On Error GoTo vbwErrHandler
' </VB WATCH>
307        If (GPIBglobalsRegistered = 0) Then
308          Call RegisterGPIBGlobals
309        End If

       ' Call the 32-bit DLL.
310        Call ibsre32(ud, v)

311        Call copy_ibvars
' <VB WATCH>
312        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibsre"

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

Sub ibstop(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
313        On Error GoTo vbwErrHandler
' </VB WATCH>
314        If (GPIBglobalsRegistered = 0) Then
315          Call RegisterGPIBGlobals
316        End If

       ' Call the 32-bit DLL.
317        Call ibstop32(ud)

318        Call copy_ibvars
' <VB WATCH>
319        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibstop"

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

Sub ibtmo(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
320        On Error GoTo vbwErrHandler
' </VB WATCH>
321        If (GPIBglobalsRegistered = 0) Then
322          Call RegisterGPIBGlobals
323        End If

       ' Call the 32-bit DLL.
324        Call ibtmo32(ud, v)

325        Call copy_ibvars
' <VB WATCH>
326        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibtmo"

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

Sub ibtrg(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
327        On Error GoTo vbwErrHandler
' </VB WATCH>
328        If (GPIBglobalsRegistered = 0) Then
329          Call RegisterGPIBGlobals
330        End If

       ' Call 32-bit DLL.
331        Call ibtrg32(ud)

332        Call copy_ibvars
' <VB WATCH>
333        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibtrg"

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

Sub ibwait(ByVal ud As Integer, ByVal mask As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
334        On Error GoTo vbwErrHandler
' </VB WATCH>
335        If (GPIBglobalsRegistered = 0) Then
336          Call RegisterGPIBGlobals
337        End If

       ' Call the 32-bit DLL.
338        Call ibwait32(ud, mask)

339        Call copy_ibvars
' <VB WATCH>
340        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibwait"

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

Sub ibwrt(ByVal ud As Integer, ByVal buf As String)
' <VB WATCH>
341        On Error GoTo vbwErrHandler
' </VB WATCH>
342        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
343        If (GPIBglobalsRegistered = 0) Then
344          Call RegisterGPIBGlobals
345        End If

346        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
347        Call ibwrt32(ud, ByVal buf, cnt)

348        Call copy_ibvars
' <VB WATCH>
349        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibwrt"

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

Sub ibwrta(ByVal ud As Integer, ByVal buf As String)
' <VB WATCH>
350        On Error GoTo vbwErrHandler
' </VB WATCH>
351        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
352        If (GPIBglobalsRegistered = 0) Then
353          Call RegisterGPIBGlobals
354        End If

355        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
356        Call ibwrt32(ud, ByVal buf, cnt)

       ' When Visual Basic remapping buffer problem is solved, use this:
       '    Call ibwrta32(ud, ByVal buf, cnt)

357        Call copy_ibvars
' <VB WATCH>
358        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibwrta"

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

Sub ibwrtf(ByVal ud As Integer, ByVal filename As String)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
359        On Error GoTo vbwErrHandler
' </VB WATCH>
360        If (GPIBglobalsRegistered = 0) Then
361          Call RegisterGPIBGlobals
362        End If

       ' Call the 32-bit DLL.
363        Call ibwrtf32(ud, ByVal filename)

364        Call copy_ibvars
' <VB WATCH>
365        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibwrtf"

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

Sub ibwrti(ByVal ud As Integer, ByRef ibuf() As Integer, ByVal cnt As Long)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
366        On Error GoTo vbwErrHandler
' </VB WATCH>
367        If (GPIBglobalsRegistered = 0) Then
368          Call RegisterGPIBGlobals
369        End If

       ' Call the 32-bit DLL.
370        Call ibwrt32(ud, ibuf(0), cnt)

371        Call copy_ibvars
' <VB WATCH>
372        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibwrti"

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

Sub ibwrtia(ByVal ud As Integer, ByRef ibuf() As Integer, ByVal cnt As Long)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
373        On Error GoTo vbwErrHandler
' </VB WATCH>
374        If (GPIBglobalsRegistered = 0) Then
375          Call RegisterGPIBGlobals
376        End If

       ' Call the 32-bit DLL.
377        Call ibwrt32(ud, ibuf(0), cnt)

       ' When Visual Basic remapping buffer problem is solved, use this:
       '    Call ibwrta32(ud, ibuf(0), cnt)

378        Call copy_ibvars
' <VB WATCH>
379        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibwrtia"

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

Function ilask(ByVal ud As Integer, ByVal opt As Integer, rval As Integer) As Integer
' <VB WATCH>
380        On Error GoTo vbwErrHandler
' </VB WATCH>
381        Dim tmprval As Long

       ' Check to see if GPIB Global variables are registered
382        If (GPIBglobalsRegistered = 0) Then
383          Call RegisterGPIBGlobals
384        End If

       ' Call the 32-bit DLL.
385        ilask = ConvertLongToInt(ibask32(ud, opt, tmprval))

386        rval = ConvertLongToInt(tmprval)

387        Call copy_ibvars
' <VB WATCH>
388        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilask"

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

Function ilbna(ByVal ud As Integer, ByVal udname As String) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
389        On Error GoTo vbwErrHandler
' </VB WATCH>
390        If (GPIBglobalsRegistered = 0) Then
391          Call RegisterGPIBGlobals
392        End If

       ' Call the 32-bit DLL.
393        ilbna = ConvertLongToInt(ibbna32(ud, ByVal udname))

394        Call copy_ibvars
' <VB WATCH>
395        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilbna"

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

Function ilcac(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
396        On Error GoTo vbwErrHandler
' </VB WATCH>
397        If (GPIBglobalsRegistered = 0) Then
398          Call RegisterGPIBGlobals
399        End If

       ' Call the 32-bit DLL.
400        ilcac = ConvertLongToInt(ibcac32(ud, v))

401        Call copy_ibvars
' <VB WATCH>
402        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilcac"

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

Function ilclr(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
403        On Error GoTo vbwErrHandler
' </VB WATCH>
404        If (GPIBglobalsRegistered = 0) Then
405          Call RegisterGPIBGlobals
406        End If

       ' Call the 32-bit DLL.
407        ilclr = ConvertLongToInt(ibclr32(ud))

408        Call copy_ibvars
' <VB WATCH>
409        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilclr"

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

Function ilcmd(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
410        On Error GoTo vbwErrHandler
' </VB WATCH>
411        If (GPIBglobalsRegistered = 0) Then
412          Call RegisterGPIBGlobals
413        End If

       ' Call the 32-bit DLL.
414        ilcmd = ConvertLongToInt(ibcmd32(ud, ByVal buf, cnt))

415        Call copy_ibvars
' <VB WATCH>
416        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilcmd"

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

Function ilcmda(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
417        On Error GoTo vbwErrHandler
' </VB WATCH>
418        If (GPIBglobalsRegistered = 0) Then
419          Call RegisterGPIBGlobals
420        End If

       ' Call the 32-bit DLL.
421        ilcmda = ConvertLongToInt(ibcmd32(ud, ByVal buf, cnt))

       ' When Visual Basic remapping buffer problem is solved, use this:
       '    ilcmda = ConvertLongToInt(ibcmda32(ud, ByVal buf, cnt))

422        Call copy_ibvars
' <VB WATCH>
423        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilcmda"

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

Function ilconfig(ByVal bdid As Integer, ByVal opt As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
424        On Error GoTo vbwErrHandler
' </VB WATCH>
425        If (GPIBglobalsRegistered = 0) Then
426          Call RegisterGPIBGlobals
427        End If

       ' Call the 32-bit DLL.
428        ilconfig = ConvertLongToInt(ibconfig32(bdid, opt, v))

429        Call copy_ibvars
' <VB WATCH>
430        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilconfig"

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

Function ildev(ByVal bdid As Integer, ByVal pad As Integer, ByVal sad As Integer, ByVal tmo As Integer, ByVal eot As Integer, ByVal eos As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
431        On Error GoTo vbwErrHandler
' </VB WATCH>
432        If (GPIBglobalsRegistered = 0) Then
433          Call RegisterGPIBGlobals
434        End If

       ' Call the 32-bit DLL.
435        ildev = ConvertLongToInt(ibdev32(bdid, pad, sad, tmo, eot, eos))

436        Call copy_ibvars
' <VB WATCH>
437        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ildev"

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

Function ildma(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
438        On Error GoTo vbwErrHandler
' </VB WATCH>
439        If (GPIBglobalsRegistered = 0) Then
440          Call RegisterGPIBGlobals
441        End If

       ' Call the 32-bit DLL.
442        ildma = ConvertLongToInt(ibdma32(ud, v))

443        Call copy_ibvars
' <VB WATCH>
444        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ildma"

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

Function ileos(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
445        On Error GoTo vbwErrHandler
' </VB WATCH>
446        If (GPIBglobalsRegistered = 0) Then
447          Call RegisterGPIBGlobals
448        End If

       ' Call the 32-bit DLL.
449        ileos = ConvertLongToInt(ibeos32(ud, v))

450        Call copy_ibvars
' <VB WATCH>
451        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ileos"

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

Function ileot(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
452        On Error GoTo vbwErrHandler
' </VB WATCH>
453        If (GPIBglobalsRegistered = 0) Then
454          Call RegisterGPIBGlobals
455        End If

       ' Call the 32-bit DLL.
456        ileot = ConvertLongToInt(ibeot32(ud, v))

457        Call copy_ibvars
' <VB WATCH>
458        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ileot"

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

Function ilfind(ByVal udname As String) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
459        On Error GoTo vbwErrHandler
' </VB WATCH>
460        If (GPIBglobalsRegistered = 0) Then
461          Call RegisterGPIBGlobals
462        End If

       ' Call the 32-bit DLL.
463        ilfind = ConvertLongToInt(ibfind32(ByVal udname))

464        Call copy_ibvars
' <VB WATCH>
465        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilfind"

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

Function ilgts(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
466        On Error GoTo vbwErrHandler
' </VB WATCH>
467        If (GPIBglobalsRegistered = 0) Then
468          Call RegisterGPIBGlobals
469        End If

       ' Call the 32-bit DLL.
470        ilgts = ConvertLongToInt(ibgts32(ud, v))

471        Call copy_ibvars
' <VB WATCH>
472        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilgts"

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

Function ilist(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
473        On Error GoTo vbwErrHandler
' </VB WATCH>
474        If (GPIBglobalsRegistered = 0) Then
475          Call RegisterGPIBGlobals
476        End If

       ' Call the 32-bit DLL.
477        ilist = ConvertLongToInt(ibist32(ud, v))

478        Call copy_ibvars
' <VB WATCH>
479        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilist"

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

Function illck(ByVal ud As Integer, ByVal v As Integer, ByVal LockWaitTime As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
480        On Error GoTo vbwErrHandler
' </VB WATCH>
481        If (GPIBglobalsRegistered = 0) Then
482          Call RegisterGPIBGlobals
483        End If

       ' Call the 32-bit DLL.
484        illck = ConvertLongToInt(iblck32(ud, v, LockWaitTime, ByVal 0))

485        Call copy_ibvars
' <VB WATCH>
486        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "illck"

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
       
Function illines(ByVal ud As Integer, lines As Integer) As Integer
' <VB WATCH>
487        On Error GoTo vbwErrHandler
' </VB WATCH>
488        Dim tmplines As Long

       ' Check to see if GPIB Global variables are registered
489        If (GPIBglobalsRegistered = 0) Then
490          Call RegisterGPIBGlobals
491        End If

       ' Call the 32-bit DLL.
492        illines = ConvertLongToInt(iblines32(ud, tmplines))

493        lines = ConvertLongToInt(tmplines)

494        Call copy_ibvars
' <VB WATCH>
495        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "illines"

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

Function illn(ByVal ud As Integer, ByVal pad As Integer, ByVal sad As Integer, ln As Integer) As Integer
' <VB WATCH>
496        On Error GoTo vbwErrHandler
' </VB WATCH>
497        Dim tmpln As Long

       ' Check to see if GPIB Global variables are registered
498        If (GPIBglobalsRegistered = 0) Then
499          Call RegisterGPIBGlobals
500        End If

       ' Call the 32-bit DLL.
501        illn = ConvertLongToInt(ibln32(ud, pad, sad, tmpln))

502        ln = ConvertLongToInt(tmpln)

503        Call copy_ibvars
' <VB WATCH>
504        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "illn"

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

Function illoc(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
505        On Error GoTo vbwErrHandler
' </VB WATCH>
506        If (GPIBglobalsRegistered = 0) Then
507          Call RegisterGPIBGlobals
508        End If

       ' Call the 32-bit DLL.
509        illoc = ConvertLongToInt(ibloc32(ud))

510        Call copy_ibvars
' <VB WATCH>
511        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "illoc"

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

Function ilonl(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
512        On Error GoTo vbwErrHandler
' </VB WATCH>
513        If (GPIBglobalsRegistered = 0) Then
514          Call RegisterGPIBGlobals
515        End If

       ' Call the 32-bit DLL.
516        ilonl = ConvertLongToInt(ibonl32(ud, v))

517        Call copy_ibvars
' <VB WATCH>
518        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilonl"

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

Function ilpad(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
519        On Error GoTo vbwErrHandler
' </VB WATCH>
520        If (GPIBglobalsRegistered = 0) Then
521          Call RegisterGPIBGlobals
522        End If

       ' Call the 32-bit DLL.
523        ilpad = ConvertLongToInt(ibpad32(ud, v))

524        Call copy_ibvars
' <VB WATCH>
525        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilpad"

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

Function ilpct(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
526        On Error GoTo vbwErrHandler
' </VB WATCH>
527        If (GPIBglobalsRegistered = 0) Then
528          Call RegisterGPIBGlobals
529        End If

       ' Call the 32-bit DLL.
530        ilpct = ConvertLongToInt(ibpct32(ud))

531        Call copy_ibvars
' <VB WATCH>
532        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilpct"

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

Function ilppc(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
533        On Error GoTo vbwErrHandler
' </VB WATCH>
534        If (GPIBglobalsRegistered = 0) Then
535          Call RegisterGPIBGlobals
536        End If

       ' Call the 32-bit DLL.
537        ilppc = ConvertLongToInt(ibppc32(ud, v))

538        Call copy_ibvars
' <VB WATCH>
539        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilppc"

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

Function ilrd(ByVal ud As Integer, buf As String, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
540        On Error GoTo vbwErrHandler
' </VB WATCH>
541        If (GPIBglobalsRegistered = 0) Then
542          Call RegisterGPIBGlobals
543        End If

       ' Call the 32-bit DLL.
544        ilrd = ConvertLongToInt(ibrd32(ud, ByVal buf, cnt))

545        Call copy_ibvars
' <VB WATCH>
546        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilrd"

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

Function ilrda(ByVal ud As Integer, buf As String, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
547        On Error GoTo vbwErrHandler
' </VB WATCH>
548        If (GPIBglobalsRegistered = 0) Then
549          Call RegisterGPIBGlobals
550        End If

       ' Call the 32-bit DLL.
551        ilrda = ConvertLongToInt(ibrd32(ud, ByVal buf, cnt))

       ' When Visual Basic remapping buffer problem solved, use this:
       '    ilrda = ConvertLongToInt(ibrda32(ud, ByVal buf, cnt))

552        Call copy_ibvars
' <VB WATCH>
553        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilrda"

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

Function ilrdf(ByVal ud As Integer, ByVal filename As String) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
554        On Error GoTo vbwErrHandler
' </VB WATCH>
555        If (GPIBglobalsRegistered = 0) Then
556          Call RegisterGPIBGlobals
557        End If

       ' Call the 32-bit DLL.
558        ilrdf = ConvertLongToInt(ibrdf32(ud, ByVal filename))

559        Call copy_ibvars
' <VB WATCH>
560        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilrdf"

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

Function ilrdi(ByVal ud As Integer, ibuf() As Integer, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
561        On Error GoTo vbwErrHandler
' </VB WATCH>
562        If (GPIBglobalsRegistered = 0) Then
563          Call RegisterGPIBGlobals
564        End If

       ' Call the 32-bit DLL.
565        ilrdi = ConvertLongToInt(ibrd32(ud, ibuf(0), cnt))

566        Call copy_ibvars
' <VB WATCH>
567        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilrdi"

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

Function ilrdia(ByVal ud As Integer, ibuf() As Integer, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
568        On Error GoTo vbwErrHandler
' </VB WATCH>
569        If (GPIBglobalsRegistered = 0) Then
570          Call RegisterGPIBGlobals
571        End If

       ' Call the 32-bit DLL.
572        ilrdia = ConvertLongToInt(ibrd32(ud, ibuf(0), cnt))

       ' When Visual Basic remapping buffer problem solved, use this:
       '    ilrdia = ConvertLongToInt(ibrda32(ud, ibuf(0), cnt))

573        Call copy_ibvars
' <VB WATCH>
574        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilrdia"

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

Function ilrpp(ByVal ud As Integer, ppr As Integer) As Integer
' <VB WATCH>
575        On Error GoTo vbwErrHandler
' </VB WATCH>
576        Static tmp_str As String * 2

       ' Check to see if GPIB Global variables are registered
577        If (GPIBglobalsRegistered = 0) Then
578          Call RegisterGPIBGlobals
579        End If

       ' Call the 32-bit DLL.
580        ilrpp = ConvertLongToInt(ibrpp32(ud, ByVal tmp_str))

581        ppr = Asc(tmp_str)

582        Call copy_ibvars
' <VB WATCH>
583        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilrpp"

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

Function ilrsc(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
584        On Error GoTo vbwErrHandler
' </VB WATCH>
585        If (GPIBglobalsRegistered = 0) Then
586          Call RegisterGPIBGlobals
587        End If

       '  Call the 32-bit DLL.
588        ilrsc = ConvertLongToInt(ibrsc32(ud, v))

589        Call copy_ibvars
' <VB WATCH>
590        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilrsc"

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

Function ilrsp(ByVal ud As Integer, spr As Integer) As Integer
' <VB WATCH>
591        On Error GoTo vbwErrHandler
' </VB WATCH>
592        Static tmp_str As String * 2

       ' Check to see if GPIB Global variables are registered
593        If (GPIBglobalsRegistered = 0) Then
594          Call RegisterGPIBGlobals
595        End If

       ' Call the 32-bit DLL
596        ilrsp = ConvertLongToInt(ibrsp32(ud, ByVal tmp_str))

597        spr = Asc(tmp_str)

598        Call copy_ibvars
' <VB WATCH>
599        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilrsp"

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

Function ilrsv(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
600        On Error GoTo vbwErrHandler
' </VB WATCH>
601        If (GPIBglobalsRegistered = 0) Then
602          Call RegisterGPIBGlobals
603        End If

       ' Call the 32-bit DLL.
604        ilrsv = ConvertLongToInt(ibrsv32(ud, v))

605        Call copy_ibvars
' <VB WATCH>
606        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilrsv"

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

Function ilsad(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
607        On Error GoTo vbwErrHandler
' </VB WATCH>
608        If (GPIBglobalsRegistered = 0) Then
609          Call RegisterGPIBGlobals
610        End If

       '  Call the 32-bit DLL.
611        ilsad = ConvertLongToInt(ibsad32(ud, v))

612        Call copy_ibvars
' <VB WATCH>
613        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilsad"

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

Function ilsic(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
614        On Error GoTo vbwErrHandler
' </VB WATCH>
615        If (GPIBglobalsRegistered = 0) Then
616          Call RegisterGPIBGlobals
617        End If

       '  Call the 32-bit DLL.
618        ilsic = ConvertLongToInt(ibsic32(ud))

619        Call copy_ibvars
' <VB WATCH>
620        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilsic"

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

Function ilsre(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
621        On Error GoTo vbwErrHandler
' </VB WATCH>
622        If (GPIBglobalsRegistered = 0) Then
623          Call RegisterGPIBGlobals
624        End If

       '  Call the 32-bit DLL.
625        ilsre = ConvertLongToInt(ibsre32(ud, v))

626        Call copy_ibvars
' <VB WATCH>
627        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilsre"

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

Function ilstop(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
628        On Error GoTo vbwErrHandler
' </VB WATCH>
629        If (GPIBglobalsRegistered = 0) Then
630          Call RegisterGPIBGlobals
631        End If

       '  Call the 32-bit DLL.
632        ilstop = ConvertLongToInt(ibstop32(ud))

633        Call copy_ibvars
' <VB WATCH>
634        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilstop"

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

Function iltmo(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
635        On Error GoTo vbwErrHandler
' </VB WATCH>
636        If (GPIBglobalsRegistered = 0) Then
637          Call RegisterGPIBGlobals
638        End If

       '  Call the 32-bit DLL.
639        iltmo = ConvertLongToInt(ibtmo32(ud, v))

640        Call copy_ibvars
' <VB WATCH>
641        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "iltmo"

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

Function iltrg(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
642        On Error GoTo vbwErrHandler
' </VB WATCH>
643        If (GPIBglobalsRegistered = 0) Then
644          Call RegisterGPIBGlobals
645        End If

       ' Call 32-bit DLL.
646        iltrg = ConvertLongToInt(ibtrg32(ud))

647        Call copy_ibvars
' <VB WATCH>
648        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "iltrg"

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

Function ilwait(ByVal ud As Integer, ByVal mask As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
649        On Error GoTo vbwErrHandler
' </VB WATCH>
650        If (GPIBglobalsRegistered = 0) Then
651          Call RegisterGPIBGlobals
652        End If

       ' Call the 32-bit DLL.
653        ilwait = ConvertLongToInt(ibwait32(ud, mask))

654        Call copy_ibvars
' <VB WATCH>
655        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilwait"

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

Function ilwrt(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
656        On Error GoTo vbwErrHandler
' </VB WATCH>
657        If (GPIBglobalsRegistered = 0) Then
658          Call RegisterGPIBGlobals
659        End If

       ' Call the 32-bit DLL.
660        ilwrt = ConvertLongToInt(ibwrt32(ud, ByVal buf, cnt))

661        Call copy_ibvars
' <VB WATCH>
662        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilwrt"

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

Function ilwrta(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
663        On Error GoTo vbwErrHandler
' </VB WATCH>
664        If (GPIBglobalsRegistered = 0) Then
665          Call RegisterGPIBGlobals
666        End If

       ' Call the 32-bit DLL.
667        ilwrta = ConvertLongToInt(ibwrt32(ud, ByVal buf, cnt))

       ' When the Visual Basic remapping solved, use this:
       '    ilwrta = ConvertLongToInt(ibwrta32(ud, ByVal buf, cnt))

668        Call copy_ibvars

' <VB WATCH>
669        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilwrta"

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

Function ilwrtf(ByVal ud As Integer, ByVal filename As String) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
670        On Error GoTo vbwErrHandler
' </VB WATCH>
671        If (GPIBglobalsRegistered = 0) Then
672          Call RegisterGPIBGlobals
673        End If

       ' Call the 32-bit DLL.
674        ilwrtf = ConvertLongToInt(ibwrtf32(ud, ByVal filename))

675        Call copy_ibvars
' <VB WATCH>
676        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilwrtf"

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

Function ilwrti(ByVal ud As Integer, ByRef ibuf() As Integer, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
677        On Error GoTo vbwErrHandler
' </VB WATCH>
678        If (GPIBglobalsRegistered = 0) Then
679          Call RegisterGPIBGlobals
680        End If

       ' Call the 32-bit DLL.
681        ilwrti = ConvertLongToInt(ibwrt32(ud, ibuf(0), cnt))

682        Call copy_ibvars
' <VB WATCH>
683        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilwrti"

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

Function ilwrtia(ByVal ud As Integer, ByRef ibuf() As Integer, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
684        On Error GoTo vbwErrHandler
' </VB WATCH>
685        If (GPIBglobalsRegistered = 0) Then
686          Call RegisterGPIBGlobals
687        End If

       ' Call the 32-bit DLL.
688        ilwrtia = ConvertLongToInt(ibwrt32(ud, ibuf(0), cnt))

       ' When Visual Basic remapping buffer problem solved, use this:
       '    ilwrtia = ConvertLongToInt(ibwrta32(ud, ibuf(0), cnt))

689        Call copy_ibvars
' <VB WATCH>
690        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilwrtia"

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

Sub PassControl(ByVal boardID As Integer, ByVal addr As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
691        On Error GoTo vbwErrHandler
' </VB WATCH>
692        If (GPIBglobalsRegistered = 0) Then
693          Call RegisterGPIBGlobals
694        End If

       ' Call the 32-bit DLL.
695        Call PassControl32(boardID, addr)

696        Call copy_ibvars
' <VB WATCH>
697        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "PassControl"

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

Sub Ppoll(ByVal boardID As Integer, result As Integer)
' <VB WATCH>
698        On Error GoTo vbwErrHandler
' </VB WATCH>
699        Dim tmpresult As Long

       ' Check to see if GPIB Global variables are registered
700        If (GPIBglobalsRegistered = 0) Then
701          Call RegisterGPIBGlobals
702        End If

       ' Call the 32-bit DLL.
703        Call PPoll32(boardID, tmpresult)

704        result = ConvertLongToInt(tmpresult)

705        Call copy_ibvars
' <VB WATCH>
706        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Ppoll"

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

Sub PpollConfig(ByVal boardID As Integer, ByVal addr As Integer, ByVal lline As Integer, ByVal sense As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
707        On Error GoTo vbwErrHandler
' </VB WATCH>
708        If (GPIBglobalsRegistered = 0) Then
709          Call RegisterGPIBGlobals
710        End If

       ' Call the 32-bit DLL.
711        Call PPollConfig32(boardID, addr, lline, sense)

712        Call copy_ibvars
' <VB WATCH>
713        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "PpollConfig"

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

Sub PpollUnconfig(ByVal boardID As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
714        On Error GoTo vbwErrHandler
' </VB WATCH>
715        If (GPIBglobalsRegistered = 0) Then
716          Call RegisterGPIBGlobals
717        End If

       ' Call the 32-bit DLL.
718        Call PPollUnconfig32(boardID, addrs(0))

719        Call copy_ibvars
' <VB WATCH>
720        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "PpollUnconfig"

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

Sub RcvRespMsg(ByVal boardID As Integer, buf As String, ByVal term As Integer)
' <VB WATCH>
721        On Error GoTo vbwErrHandler
' </VB WATCH>
722        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
723        If (GPIBglobalsRegistered = 0) Then
724          Call RegisterGPIBGlobals
725        End If

726        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
727        Call RcvRespMsg32(boardID, ByVal buf, cnt, term)

728        Call copy_ibvars
' <VB WATCH>
729        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "RcvRespMsg"

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

Sub ReadStatusByte(ByVal boardID As Integer, ByVal addr As Integer, result As Integer)
' <VB WATCH>
730        On Error GoTo vbwErrHandler
' </VB WATCH>
731        Dim tmpresult As Long

       ' Check to see if GPIB Global variables are registered
732        If (GPIBglobalsRegistered = 0) Then
733          Call RegisterGPIBGlobals
734        End If

       ' Call the 32-bit DLL.
735        Call ReadStatusByte32(boardID, addr, tmpresult)

736        result = ConvertLongToInt(tmpresult)

737        Call copy_ibvars
' <VB WATCH>
738        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ReadStatusByte"

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

Sub Receive(ByVal boardID As Integer, ByVal addr As Integer, buf As String, ByVal term As Integer)
' <VB WATCH>
739        On Error GoTo vbwErrHandler
' </VB WATCH>
740        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
741        If (GPIBglobalsRegistered = 0) Then
742          Call RegisterGPIBGlobals
743        End If

744        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
745        Call Receive32(boardID, addr, ByVal buf, cnt, term)

746        Call copy_ibvars
' <VB WATCH>
747        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Receive"

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

Sub ReceiveSetup(ByVal boardID As Integer, ByVal addr As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
748        On Error GoTo vbwErrHandler
' </VB WATCH>
749        If (GPIBglobalsRegistered = 0) Then
750          Call RegisterGPIBGlobals
751        End If

       ' Call the 32-bit DLL.
752        Call ReceiveSetup32(boardID, addr)

753        Call copy_ibvars
' <VB WATCH>
754        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ReceiveSetup"

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

Sub ResetSys(ByVal boardID As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
755        On Error GoTo vbwErrHandler
' </VB WATCH>
756        If (GPIBglobalsRegistered = 0) Then
757          Call RegisterGPIBGlobals
758        End If

       ' Call the 32-bit DLL.
759        Call ResetSys32(boardID, addrs(0))

760        Call copy_ibvars
' <VB WATCH>
761        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ResetSys"

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

Sub Send(ByVal boardID As Integer, ByVal addr As Integer, ByVal buf As String, ByVal term As Integer)
' <VB WATCH>
762        On Error GoTo vbwErrHandler
' </VB WATCH>
763        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
764        If (GPIBglobalsRegistered = 0) Then
765          Call RegisterGPIBGlobals
766        End If

767        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
768        Call Send32(boardID, addr, ByVal buf, cnt, term)

769        Call copy_ibvars
' <VB WATCH>
770        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Send"

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

Sub SendCmds(ByVal boardID As Integer, ByVal cmdbuf As String)
' <VB WATCH>
771        On Error GoTo vbwErrHandler
' </VB WATCH>
772        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
773        If (GPIBglobalsRegistered = 0) Then
774          Call RegisterGPIBGlobals
775        End If

776        cnt = CLng(Len(cmdbuf))

       ' Call the 32-bit DLL.
777        Call SendCmds32(boardID, ByVal cmdbuf, cnt)

778        Call copy_ibvars
' <VB WATCH>
779        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendCmds"

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

Sub SendDataBytes(ByVal boardID As Integer, ByVal buf As String, ByVal term As Integer)
' <VB WATCH>
780        On Error GoTo vbwErrHandler
' </VB WATCH>
781        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
782        If (GPIBglobalsRegistered = 0) Then
783          Call RegisterGPIBGlobals
784        End If

785        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
786        Call SendDataBytes32(boardID, ByVal buf, cnt, term)

787        Call copy_ibvars
' <VB WATCH>
788        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendDataBytes"

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

Sub SendIFC(ByVal boardID As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
789        On Error GoTo vbwErrHandler
' </VB WATCH>
790        If (GPIBglobalsRegistered = 0) Then
791          Call RegisterGPIBGlobals
792        End If

       ' Call the 32-bit DLL.
793        Call SendIFC32(boardID)

794        Call copy_ibvars
' <VB WATCH>
795        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendIFC"

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

Sub SendList(ByVal boardID As Integer, addr() As Integer, ByVal buf As String, ByVal term As Integer)
' <VB WATCH>
796        On Error GoTo vbwErrHandler
' </VB WATCH>
797        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
798        If (GPIBglobalsRegistered = 0) Then
799          Call RegisterGPIBGlobals
800        End If

801        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
802        Call SendList32(boardID, addr(0), ByVal buf, cnt, term)

803        Call copy_ibvars
' <VB WATCH>
804        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendList"

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

Sub SendLLO(ByVal boardID As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
805        On Error GoTo vbwErrHandler
' </VB WATCH>
806        If (GPIBglobalsRegistered = 0) Then
807          Call RegisterGPIBGlobals
808        End If

       ' Call the 32-bit DLL.
809        Call SendLLO32(boardID)

810        Call copy_ibvars
' <VB WATCH>
811        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendLLO"

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

Sub SendSetup(ByVal boardID As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
812        On Error GoTo vbwErrHandler
' </VB WATCH>
813        If (GPIBglobalsRegistered = 0) Then
814          Call RegisterGPIBGlobals
815        End If

       ' Call the 32-bit DLL.
816        Call SendSetup32(boardID, addrs(0))

817        Call copy_ibvars
' <VB WATCH>
818        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SendSetup"

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

Sub SetRWLS(ByVal boardID As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
819        On Error GoTo vbwErrHandler
' </VB WATCH>
820        If (GPIBglobalsRegistered = 0) Then
821          Call RegisterGPIBGlobals
822        End If

       ' Call the 32-bit DLL.
823        Call SetRWLS32(boardID, addrs(0))

824        Call copy_ibvars
' <VB WATCH>
825        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SetRWLS"

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

Sub TestSRQ(ByVal boardID As Integer, result As Integer)
' <VB WATCH>
826        On Error GoTo vbwErrHandler
' </VB WATCH>
827        Call ibwait(boardID, 0)

828        If ibsta And &H1000 Then
829            result = 1
830        Else
831            result = 0
832        End If

' <VB WATCH>
833        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "TestSRQ"

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

Sub TestSys(ByVal boardID As Integer, addrs() As Integer, results() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
834        On Error GoTo vbwErrHandler
' </VB WATCH>
835        If (GPIBglobalsRegistered = 0) Then
836          Call RegisterGPIBGlobals
837        End If

       ' Call the 32-bit DLL.
838        Call TestSys32(boardID, addrs(0), results(0))

839        Call copy_ibvars
' <VB WATCH>
840        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "TestSys"

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

Sub Trigger(ByVal boardID As Integer, ByVal addr As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
841        On Error GoTo vbwErrHandler
' </VB WATCH>
842        If (GPIBglobalsRegistered = 0) Then
843          Call RegisterGPIBGlobals
844        End If

       ' Call the 32-bit DLL.
845        Call Trigger32(boardID, addr)

846        Call copy_ibvars
' <VB WATCH>
847        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Trigger"

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

Sub TriggerList(ByVal boardID As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
848        On Error GoTo vbwErrHandler
' </VB WATCH>
849        If (GPIBglobalsRegistered = 0) Then
850          Call RegisterGPIBGlobals
851        End If

       ' Call the 32-bit DLL.
852        Call TriggerList32(boardID, addrs(0))

853        Call copy_ibvars
' <VB WATCH>
854        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "TriggerList"

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

Sub WaitSRQ(ByVal boardID As Integer, result As Integer)
' <VB WATCH>
855        On Error GoTo vbwErrHandler
' </VB WATCH>
856        Call ibwait(boardID, &H5000)

857        If ibsta And &H1000 Then
858            result = 1
859        Else
860            result = 0
861        End If
' <VB WATCH>
862        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "WaitSRQ"

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


Private Function ConvertLongToInt(LongNumb As Long) As Integer
' <VB WATCH>
863        On Error GoTo vbwErrHandler
' </VB WATCH>

864      If (LongNumb And &H8000&) = 0 Then
865          ConvertLongToInt = LongNumb And &HFFFF&
866      Else
867        ConvertLongToInt = &H8000 Or (LongNumb And &H7FFF&)
868      End If

' <VB WATCH>
869        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ConvertLongToInt"

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

Public Sub RegisterGPIBGlobals()
' <VB WATCH>
870        On Error GoTo vbwErrHandler
' </VB WATCH>
871        Dim rc As Long

872        rc = RegisterGpibGlobalsForThread(Longibsta, Longiberr, Longibcnt, ibcntl)
873        If (rc = 0) Then
874          GPIBglobalsRegistered = 1
875        ElseIf (rc = 1) Then
876          rc = UnregisterGpibGlobalsForThread
877          rc = RegisterGpibGlobalsForThread(Longibsta, Longiberr, Longibcnt, ibcntl)
878          GPIBglobalsRegistered = 1
879        ElseIf (rc = 2) Then
880          rc = UnregisterGpibGlobalsForThread
881          ibsta = &H8000
882          iberr = EDVR
883          ibcntl = &HDEAD37F0
884        ElseIf (rc = 3) Then
885          rc = UnregisterGpibGlobalsForThread
886          ibsta = &H8000
887          iberr = EDVR
888          ibcntl = &HDEAD37F0
889        Else
890          ibsta = &H8000
891          iberr = EDVR
892          ibcntl = &HDEAD37F0
893        End If
' <VB WATCH>
894        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "RegisterGPIBGlobals"

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

Public Sub UnregisterGPIBGlobals()
' <VB WATCH>
895        On Error GoTo vbwErrHandler
' </VB WATCH>
896        Dim rc As Long

897        rc = UnregisterGpibGlobalsForThread
898        GPIBglobalsRegistered = 0

' <VB WATCH>
899        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "UnregisterGPIBGlobals"

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



Public Function ThreadIbsta() As Integer
       ' Call the 32-bit DLL.
' <VB WATCH>
900        On Error GoTo vbwErrHandler
' </VB WATCH>
901        ThreadIbsta = ConvertLongToInt(ThreadIbsta32())
' <VB WATCH>
902        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ThreadIbsta"

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

Public Function ThreadIberr() As Integer
       ' Call the 32-bit DLL.
' <VB WATCH>
903        On Error GoTo vbwErrHandler
' </VB WATCH>
904        ThreadIberr = ConvertLongToInt(ThreadIberr32())
' <VB WATCH>
905        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ThreadIberr"

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

Public Function ThreadIbcnt() As Integer
       ' Call the 32-bit DLL.
' <VB WATCH>
906        On Error GoTo vbwErrHandler
' </VB WATCH>
907        ThreadIbcnt = ConvertLongToInt(ThreadIbcnt32())
' <VB WATCH>
908        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ThreadIbcnt"

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

Public Function ThreadIbcntl() As Long
       ' Call the 32-bit DLL.
' <VB WATCH>
909        On Error GoTo vbwErrHandler
' </VB WATCH>
910        ThreadIbcntl = ThreadIbcntl32()
' <VB WATCH>
911        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ThreadIbcntl"

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

Public Function illock(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
912        On Error GoTo vbwErrHandler
' </VB WATCH>
913        If (GPIBglobalsRegistered = 0) Then
914          Call RegisterGPIBGlobals
915        End If

       ' Call the 32-bit DLL.
916        illock = ConvertLongToInt(iblock32(ud))

917        Call copy_ibvars
' <VB WATCH>
918        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "illock"

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

Public Function ilunlock(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
919        On Error GoTo vbwErrHandler
' </VB WATCH>
920        If (GPIBglobalsRegistered = 0) Then
921          Call RegisterGPIBGlobals
922        End If

       ' Call the 32-bit DLL.
923        ilunlock = ConvertLongToInt(ibunlock32(ud))

924        Call copy_ibvars
' <VB WATCH>
925        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilunlock"

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

Public Sub iblock(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
926        On Error GoTo vbwErrHandler
' </VB WATCH>
927        If (GPIBglobalsRegistered = 0) Then
928          Call RegisterGPIBGlobals
929        End If

       ' Call the 32-bit DLL.
930        Call iblock32(ud)

931        Call copy_ibvars
' <VB WATCH>
932        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "iblock"

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

Public Sub ibunlock(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
933        On Error GoTo vbwErrHandler
' </VB WATCH>
934        If (GPIBglobalsRegistered = 0) Then
935          Call RegisterGPIBGlobals
936        End If

       ' Call the 32-bit DLL.
937        Call ibunlock32(ud)

938        Call copy_ibvars
' <VB WATCH>
939        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibunlock"

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

Public Function illockx(ByVal ud As Integer, ByVal LockWaitTime As Integer, ByVal buf As String) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
940        On Error GoTo vbwErrHandler
' </VB WATCH>
941        If (GPIBglobalsRegistered = 0) Then
942          Call RegisterGPIBGlobals
943        End If

       ' Call the 32-bit DLL.
944        illockx = ConvertLongToInt(iblockx32(ud, LockWaitTime, buf))

945        Call copy_ibvars
' <VB WATCH>
946        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "illockx"

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

Public Function ilunlockx(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
947        On Error GoTo vbwErrHandler
' </VB WATCH>
948        If (GPIBglobalsRegistered = 0) Then
949          Call RegisterGPIBGlobals
950        End If

       ' Call the 32-bit DLL.
951        ilunlockx = ConvertLongToInt(ibunlockx32(ud))

952        Call copy_ibvars
' <VB WATCH>
953        Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ilunlockx"

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

Public Sub iblockx(ByVal ud As Integer, ByVal LockWaitTime As Integer, ByVal buf As String)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
954        On Error GoTo vbwErrHandler
' </VB WATCH>
955        If (GPIBglobalsRegistered = 0) Then
956          Call RegisterGPIBGlobals
957        End If

       ' Call the 32-bit DLL.
958        Call iblockx32(ud, LockWaitTime, buf)

959        Call copy_ibvars
' <VB WATCH>
960        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "iblockx"

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

Public Sub ibunlockx(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
961        On Error GoTo vbwErrHandler
' </VB WATCH>
962        If (GPIBglobalsRegistered = 0) Then
963          Call RegisterGPIBGlobals
964        End If

       ' Call the 32-bit DLL.
965        Call ibunlockx32(ud)

966        Call copy_ibvars
' <VB WATCH>
967        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ibunlockx"

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



