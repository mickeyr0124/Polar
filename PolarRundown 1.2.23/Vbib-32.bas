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
2          Const VBWPROCNAME = "VBIB32.AllSpoll"
3          If vbwProtector.vbwTraceProc Then
4              Dim vbwProtectorParameterString As String
5              If vbwProtector.vbwTraceParameters Then
6                  vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
7                  vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ", "
8                  vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("results", results) & ") "
9              End If
10             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
11         End If
' </VB WATCH>
12         If (GPIBglobalsRegistered = 0) Then
13           Call RegisterGPIBGlobals
14         End If

       ' Call the 32-bit DLL.
15         Call AllSpoll32(boardID, addrs(0), results(0))

16         Call copy_ibvars
' <VB WATCH>
17         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
18         Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addrs", addrs
            vbwReportVariable "results", results
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub copy_ibvars()
' <VB WATCH>
19         On Error GoTo vbwErrHandler
20         Const VBWPROCNAME = "VBIB32.copy_ibvars"
21         If vbwProtector.vbwTraceProc Then
22             Dim vbwProtectorParameterString As String
23             If vbwProtector.vbwTraceParameters Then
24                 vbwProtectorParameterString = "()"
25             End If
26             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
27         End If
' </VB WATCH>
28         ibsta = ConvertLongToInt(Longibsta)
29         iberr = CInt(Longiberr)
30         ibcnt = ConvertLongToInt(ibcntl)
' <VB WATCH>
31         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
32         Exit Sub
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
End Sub

Sub DevClear(ByVal boardID As Integer, ByVal addr As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
33         On Error GoTo vbwErrHandler
34         Const VBWPROCNAME = "VBIB32.DevClear"
35         If vbwProtector.vbwTraceProc Then
36             Dim vbwProtectorParameterString As String
37             If vbwProtector.vbwTraceParameters Then
38                 vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
39                 vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addr", addr) & ") "
40             End If
41             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
42         End If
' </VB WATCH>
43         If (GPIBglobalsRegistered = 0) Then
44           Call RegisterGPIBGlobals
45         End If

       ' Call the 32-bit DLL.
46         Call DevClear32(boardID, addr)

47         Call copy_ibvars
' <VB WATCH>
48         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
49         Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addr", addr
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub DevClearList(ByVal boardID As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
50         On Error GoTo vbwErrHandler
51         Const VBWPROCNAME = "VBIB32.DevClearList"
52         If vbwProtector.vbwTraceProc Then
53             Dim vbwProtectorParameterString As String
54             If vbwProtector.vbwTraceParameters Then
55                 vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
56                 vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ") "
57             End If
58             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
59         End If
' </VB WATCH>
60         If (GPIBglobalsRegistered = 0) Then
61           Call RegisterGPIBGlobals
62         End If

       ' Call the 32-bit DLL.
63         Call DevClearList32(boardID, addrs(0))

64         Call copy_ibvars
' <VB WATCH>
65         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
66         Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addrs", addrs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub EnableLocal(ByVal boardID As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
67         On Error GoTo vbwErrHandler
68         Const VBWPROCNAME = "VBIB32.EnableLocal"
69         If vbwProtector.vbwTraceProc Then
70             Dim vbwProtectorParameterString As String
71             If vbwProtector.vbwTraceParameters Then
72                 vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
73                 vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ") "
74             End If
75             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
76         End If
' </VB WATCH>
77         If (GPIBglobalsRegistered = 0) Then
78           Call RegisterGPIBGlobals
79         End If

       ' Call the 32-bit DLL.
80         Call EnableLocal32(boardID, addrs(0))

81         Call copy_ibvars
' <VB WATCH>
82         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
83         Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addrs", addrs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub EnableRemote(ByVal boardID As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
84         On Error GoTo vbwErrHandler
85         Const VBWPROCNAME = "VBIB32.EnableRemote"
86         If vbwProtector.vbwTraceProc Then
87             Dim vbwProtectorParameterString As String
88             If vbwProtector.vbwTraceParameters Then
89                 vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
90                 vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ") "
91             End If
92             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
93         End If
' </VB WATCH>
94         If (GPIBglobalsRegistered = 0) Then
95           Call RegisterGPIBGlobals
96         End If

       ' Call the 32-bit DLL.
97         Call EnableRemote32(boardID, addrs(0))

98         Call copy_ibvars
' <VB WATCH>
99         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
100        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addrs", addrs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub FindLstn(ByVal boardID As Integer, addrs() As Integer, results() As Integer, ByVal limit As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
101        On Error GoTo vbwErrHandler
102        Const VBWPROCNAME = "VBIB32.FindLstn"
103        If vbwProtector.vbwTraceProc Then
104            Dim vbwProtectorParameterString As String
105            If vbwProtector.vbwTraceParameters Then
106                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
107                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ", "
108                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("results", results) & ", "
109                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("limit", limit) & ") "
110            End If
111            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
112        End If
' </VB WATCH>
113        If (GPIBglobalsRegistered = 0) Then
114          Call RegisterGPIBGlobals
115        End If

       ' Call the 32-bit DLL.
116        Call FindLstn32(boardID, addrs(0), results(0), limit)

117        Call copy_ibvars
' <VB WATCH>
118        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
119        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addrs", addrs
            vbwReportVariable "results", results
            vbwReportVariable "limit", limit
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub FindRQS(ByVal boardID As Integer, addrs() As Integer, result As Integer)
' <VB WATCH>
120        On Error GoTo vbwErrHandler
121        Const VBWPROCNAME = "VBIB32.FindRQS"
122        If vbwProtector.vbwTraceProc Then
123            Dim vbwProtectorParameterString As String
124            If vbwProtector.vbwTraceParameters Then
125                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
126                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ", "
127                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("result", result) & ") "
128            End If
129            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
130        End If
' </VB WATCH>
131       Dim tmpresult As Long

       ' Check to see if GPIB Global variables are registered
132        If (GPIBglobalsRegistered = 0) Then
133          Call RegisterGPIBGlobals
134        End If

       ' Call the 32-bit DLL.
135        Call FindRQS32(boardID, addrs(0), tmpresult)

136        result = ConvertLongToInt(tmpresult)

137        Call copy_ibvars
' <VB WATCH>
138        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
139        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addrs", addrs
            vbwReportVariable "result", result
            vbwReportVariable "tmpresult", tmpresult
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibask(ByVal ud As Integer, ByVal opt As Integer, rval As Integer)
' <VB WATCH>
140        On Error GoTo vbwErrHandler
141        Const VBWPROCNAME = "VBIB32.ibask"
142        If vbwProtector.vbwTraceProc Then
143            Dim vbwProtectorParameterString As String
144            If vbwProtector.vbwTraceParameters Then
145                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
146                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("opt", opt) & ", "
147                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("rval", rval) & ") "
148            End If
149            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
150        End If
' </VB WATCH>
151      Dim tmprval As Long

       ' Check to see if GPIB Global variables are registered
152        If (GPIBglobalsRegistered = 0) Then
153          Call RegisterGPIBGlobals
154        End If

       ' Call the 32-bit DLL.
155        Call ibask32(ud, opt, tmprval)

156        rval = ConvertLongToInt(tmprval)

157        Call copy_ibvars
' <VB WATCH>
158        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
159        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "opt", opt
            vbwReportVariable "rval", rval
            vbwReportVariable "tmprval", tmprval
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibbna(ByVal ud As Integer, ByVal udname As String)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
160        On Error GoTo vbwErrHandler
161        Const VBWPROCNAME = "VBIB32.ibbna"
162        If vbwProtector.vbwTraceProc Then
163            Dim vbwProtectorParameterString As String
164            If vbwProtector.vbwTraceParameters Then
165                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
166                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("udname", udname) & ") "
167            End If
168            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
169        End If
' </VB WATCH>
170        If (GPIBglobalsRegistered = 0) Then
171          Call RegisterGPIBGlobals
172        End If

       ' Call the 32-bit DLL.
173        Call ibbna32(ud, ByVal udname)

174        Call copy_ibvars
' <VB WATCH>
175        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
176        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "udname", udname
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibcac(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
177        On Error GoTo vbwErrHandler
178        Const VBWPROCNAME = "VBIB32.ibcac"
179        If vbwProtector.vbwTraceProc Then
180            Dim vbwProtectorParameterString As String
181            If vbwProtector.vbwTraceParameters Then
182                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
183                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
184            End If
185            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
186        End If
' </VB WATCH>
187        If (GPIBglobalsRegistered = 0) Then
188          Call RegisterGPIBGlobals
189        End If

       ' Call the 32-bit DLL.
190        Call ibcac32(ud, v)

191        Call copy_ibvars
' <VB WATCH>
192        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
193        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibclr(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
194        On Error GoTo vbwErrHandler
195        Const VBWPROCNAME = "VBIB32.ibclr"
196        If vbwProtector.vbwTraceProc Then
197            Dim vbwProtectorParameterString As String
198            If vbwProtector.vbwTraceParameters Then
199                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
200            End If
201            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
202        End If
' </VB WATCH>
203        If (GPIBglobalsRegistered = 0) Then
204          Call RegisterGPIBGlobals
205        End If

       ' Call the 32-bit DLL.
206        Call ibclr32(ud)

207        Call copy_ibvars
' <VB WATCH>
208        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
209        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibcmd(ByVal ud As Integer, ByVal buf As String)
' <VB WATCH>
210        On Error GoTo vbwErrHandler
211        Const VBWPROCNAME = "VBIB32.ibcmd"
212        If vbwProtector.vbwTraceProc Then
213            Dim vbwProtectorParameterString As String
214            If vbwProtector.vbwTraceParameters Then
215                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
216                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ") "
217            End If
218            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
219        End If
' </VB WATCH>
220       Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
221        If (GPIBglobalsRegistered = 0) Then
222          Call RegisterGPIBGlobals
223        End If

224        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
225        Call ibcmd32(ud, ByVal buf, cnt)

226        Call copy_ibvars
' <VB WATCH>
227        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
228        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibcmda(ByVal ud As Integer, ByVal buf As String)
' <VB WATCH>
229        On Error GoTo vbwErrHandler
230        Const VBWPROCNAME = "VBIB32.ibcmda"
231        If vbwProtector.vbwTraceProc Then
232            Dim vbwProtectorParameterString As String
233            If vbwProtector.vbwTraceParameters Then
234                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
235                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ") "
236            End If
237            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
238        End If
' </VB WATCH>
239        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
240        If (GPIBglobalsRegistered = 0) Then
241          Call RegisterGPIBGlobals
242        End If

243        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
244        Call ibcmd32(ud, ByVal buf, cnt)

       ' When Visual Basic remapping buffer problem solved, then use:
       '    call ibcmda32(ud, ByVal buf, cnt)

245        Call copy_ibvars
' <VB WATCH>
246        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
247        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibconfig(ByVal bdid As Integer, ByVal opt As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
248        On Error GoTo vbwErrHandler
249        Const VBWPROCNAME = "VBIB32.ibconfig"
250        If vbwProtector.vbwTraceProc Then
251            Dim vbwProtectorParameterString As String
252            If vbwProtector.vbwTraceParameters Then
253                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("bdid", bdid) & ", "
254                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("opt", opt) & ", "
255                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
256            End If
257            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
258        End If
' </VB WATCH>
259        If (GPIBglobalsRegistered = 0) Then
260          Call RegisterGPIBGlobals
261        End If

       ' Call the 32-bit DLL.
262        Call ibconfig32(bdid, opt, v)

263        Call copy_ibvars
' <VB WATCH>
264        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
265        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "bdid", bdid
            vbwReportVariable "opt", opt
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibdev(ByVal bdid As Integer, ByVal pad As Integer, ByVal sad As Integer, ByVal tmo As Integer, ByVal eot As Integer, ByVal eos As Integer, ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
266        On Error GoTo vbwErrHandler
267        Const VBWPROCNAME = "VBIB32.ibdev"
268        If vbwProtector.vbwTraceProc Then
269            Dim vbwProtectorParameterString As String
270            If vbwProtector.vbwTraceParameters Then
271                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("bdid", bdid) & ", "
272                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("pad", pad) & ", "
273                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sad", sad) & ", "
274                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("tmo", tmo) & ", "
275                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("eot", eot) & ", "
276                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("eos", eos) & ", "
277                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ud", ud) & ") "
278            End If
279            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
280        End If
' </VB WATCH>
281        If (GPIBglobalsRegistered = 0) Then
282          Call RegisterGPIBGlobals
283        End If

       ' Call the 32-bit DLL.
284        ud = ConvertLongToInt(ibdev32(bdid, pad, sad, tmo, eot, eos))

285        Call copy_ibvars
' <VB WATCH>
286        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
287        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "bdid", bdid
            vbwReportVariable "pad", pad
            vbwReportVariable "sad", sad
            vbwReportVariable "tmo", tmo
            vbwReportVariable "eot", eot
            vbwReportVariable "eos", eos
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibdma(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
288        On Error GoTo vbwErrHandler
289        Const VBWPROCNAME = "VBIB32.ibdma"
290        If vbwProtector.vbwTraceProc Then
291            Dim vbwProtectorParameterString As String
292            If vbwProtector.vbwTraceParameters Then
293                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
294                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
295            End If
296            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
297        End If
' </VB WATCH>
298        If (GPIBglobalsRegistered = 0) Then
299          Call RegisterGPIBGlobals
300        End If

       ' Call the 32-bit DLL.
301        Call ibdma32(ud, v)

302        Call copy_ibvars
' <VB WATCH>
303        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
304        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibeos(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
305        On Error GoTo vbwErrHandler
306        Const VBWPROCNAME = "VBIB32.ibeos"
307        If vbwProtector.vbwTraceProc Then
308            Dim vbwProtectorParameterString As String
309            If vbwProtector.vbwTraceParameters Then
310                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
311                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
312            End If
313            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
314        End If
' </VB WATCH>
315        If (GPIBglobalsRegistered = 0) Then
316          Call RegisterGPIBGlobals
317        End If

       ' Call the 32-bit DLL.
318        Call ibeos32(ud, v)

319        Call copy_ibvars
' <VB WATCH>
320        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
321        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibeot(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
322        On Error GoTo vbwErrHandler
323        Const VBWPROCNAME = "VBIB32.ibeot"
324        If vbwProtector.vbwTraceProc Then
325            Dim vbwProtectorParameterString As String
326            If vbwProtector.vbwTraceParameters Then
327                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
328                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
329            End If
330            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
331        End If
' </VB WATCH>
332        If (GPIBglobalsRegistered = 0) Then
333          Call RegisterGPIBGlobals
334        End If

       ' Call the 32-bit DLL.
335        Call ibeot32(ud, v)

336        Call copy_ibvars
' <VB WATCH>
337        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
338        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibfind(ByVal udname As String, ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
339        On Error GoTo vbwErrHandler
340        Const VBWPROCNAME = "VBIB32.ibfind"
341        If vbwProtector.vbwTraceProc Then
342            Dim vbwProtectorParameterString As String
343            If vbwProtector.vbwTraceParameters Then
344                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("udname", udname) & ", "
345                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ud", ud) & ") "
346            End If
347            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
348        End If
' </VB WATCH>
349        If (GPIBglobalsRegistered = 0) Then
350          Call RegisterGPIBGlobals
351        End If

       ' Call the 32-bit DLL.
352        ud = ConvertLongToInt(ibfind32(ByVal udname))

353        Call copy_ibvars
' <VB WATCH>
354        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
355        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "udname", udname
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibgts(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
356        On Error GoTo vbwErrHandler
357        Const VBWPROCNAME = "VBIB32.ibgts"
358        If vbwProtector.vbwTraceProc Then
359            Dim vbwProtectorParameterString As String
360            If vbwProtector.vbwTraceParameters Then
361                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
362                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
363            End If
364            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
365        End If
' </VB WATCH>
366        If (GPIBglobalsRegistered = 0) Then
367          Call RegisterGPIBGlobals
368        End If

       ' Call the 32-bit DLL.
369        Call ibgts32(ud, v)

370        Call copy_ibvars
' <VB WATCH>
371        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
372        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibist(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
373        On Error GoTo vbwErrHandler
374        Const VBWPROCNAME = "VBIB32.ibist"
375        If vbwProtector.vbwTraceProc Then
376            Dim vbwProtectorParameterString As String
377            If vbwProtector.vbwTraceParameters Then
378                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
379                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
380            End If
381            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
382        End If
' </VB WATCH>
383        If (GPIBglobalsRegistered = 0) Then
384          Call RegisterGPIBGlobals
385        End If

       ' Call the 32-bit DLL.
386        Call ibist32(ud, v)

387        Call copy_ibvars
' <VB WATCH>
388        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
389        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub iblines(ByVal ud As Integer, lines As Integer)
' <VB WATCH>
390        On Error GoTo vbwErrHandler
391        Const VBWPROCNAME = "VBIB32.iblines"
392        If vbwProtector.vbwTraceProc Then
393            Dim vbwProtectorParameterString As String
394            If vbwProtector.vbwTraceParameters Then
395                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
396                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("lines", lines) & ") "
397            End If
398            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
399        End If
' </VB WATCH>
400       Dim tmplines As Long

       ' Check to see if GPIB Global variables are registered
401        If (GPIBglobalsRegistered = 0) Then
402          Call RegisterGPIBGlobals
403        End If

       ' Call the 32-bit DLL.
404        Call iblines32(ud, tmplines)

405        lines = ConvertLongToInt(tmplines)

406        Call copy_ibvars
' <VB WATCH>
407        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
408        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "lines", lines
            vbwReportVariable "tmplines", tmplines
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibln(ByVal ud As Integer, ByVal pad As Integer, ByVal sad As Integer, ln As Integer)
' <VB WATCH>
409        On Error GoTo vbwErrHandler
410        Const VBWPROCNAME = "VBIB32.ibln"
411        If vbwProtector.vbwTraceProc Then
412            Dim vbwProtectorParameterString As String
413            If vbwProtector.vbwTraceParameters Then
414                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
415                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("pad", pad) & ", "
416                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sad", sad) & ", "
417                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ln", ln) & ") "
418            End If
419            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
420        End If
' </VB WATCH>
421        Dim tmpln As Long

       ' Check to see if GPIB Global variables are registered
422        If (GPIBglobalsRegistered = 0) Then
423          Call RegisterGPIBGlobals
424        End If

       ' Call the 32-bit DLL.
425        Call ibln32(ud, pad, sad, tmpln)

426        ln = ConvertLongToInt(tmpln)

427        Call copy_ibvars
' <VB WATCH>
428        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
429        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "pad", pad
            vbwReportVariable "sad", sad
            vbwReportVariable "ln", ln
            vbwReportVariable "tmpln", tmpln
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibloc(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
430        On Error GoTo vbwErrHandler
431        Const VBWPROCNAME = "VBIB32.ibloc"
432        If vbwProtector.vbwTraceProc Then
433            Dim vbwProtectorParameterString As String
434            If vbwProtector.vbwTraceParameters Then
435                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
436            End If
437            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
438        End If
' </VB WATCH>
439        If (GPIBglobalsRegistered = 0) Then
440          Call RegisterGPIBGlobals
441        End If

       ' Call the 32-bit DLL.
442        Call ibloc32(ud)

443        Call copy_ibvars
' <VB WATCH>
444        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
445        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub iblck(ByVal ud As Integer, ByVal v As Integer, ByVal LockWaitTime As Long)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
446        On Error GoTo vbwErrHandler
447        Const VBWPROCNAME = "VBIB32.iblck"
448        If vbwProtector.vbwTraceProc Then
449            Dim vbwProtectorParameterString As String
450            If vbwProtector.vbwTraceParameters Then
451                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
452                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ", "
453                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("LockWaitTime", LockWaitTime) & ") "
454            End If
455            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
456        End If
' </VB WATCH>
457        If (GPIBglobalsRegistered = 0) Then
458          Call RegisterGPIBGlobals
459        End If

       ' Call the 32-bit DLL.
460        Call iblck32(ud, v, LockWaitTime, ByVal 0)

461        Call copy_ibvars
' <VB WATCH>
462        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
463        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportVariable "LockWaitTime", LockWaitTime
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibonl(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
464        On Error GoTo vbwErrHandler
465        Const VBWPROCNAME = "VBIB32.ibonl"
466        If vbwProtector.vbwTraceProc Then
467            Dim vbwProtectorParameterString As String
468            If vbwProtector.vbwTraceParameters Then
469                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
470                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
471            End If
472            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
473        End If
' </VB WATCH>
474        If (GPIBglobalsRegistered = 0) Then
475          Call RegisterGPIBGlobals
476        End If

       ' Call the 32-bit DLL.
477        Call ibonl32(ud, v)

478        Call copy_ibvars
' <VB WATCH>
479        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
480        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibpad(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
481        On Error GoTo vbwErrHandler
482        Const VBWPROCNAME = "VBIB32.ibpad"
483        If vbwProtector.vbwTraceProc Then
484            Dim vbwProtectorParameterString As String
485            If vbwProtector.vbwTraceParameters Then
486                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
487                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
488            End If
489            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
490        End If
' </VB WATCH>
491        If (GPIBglobalsRegistered = 0) Then
492          Call RegisterGPIBGlobals
493        End If

       ' Call the 32-bit DLL.
494        Call ibpad32(ud, v)

495        Call copy_ibvars
' <VB WATCH>
496        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
497        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibpct(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
498        On Error GoTo vbwErrHandler
499        Const VBWPROCNAME = "VBIB32.ibpct"
500        If vbwProtector.vbwTraceProc Then
501            Dim vbwProtectorParameterString As String
502            If vbwProtector.vbwTraceParameters Then
503                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
504            End If
505            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
506        End If
' </VB WATCH>
507        If (GPIBglobalsRegistered = 0) Then
508          Call RegisterGPIBGlobals
509        End If

       ' Call the 32-bit DLL.
510        Call ibpct32(ud)

511        Call copy_ibvars
' <VB WATCH>
512        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
513        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibppc(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
514        On Error GoTo vbwErrHandler
515        Const VBWPROCNAME = "VBIB32.ibppc"
516        If vbwProtector.vbwTraceProc Then
517            Dim vbwProtectorParameterString As String
518            If vbwProtector.vbwTraceParameters Then
519                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
520                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
521            End If
522            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
523        End If
' </VB WATCH>
524        If (GPIBglobalsRegistered = 0) Then
525          Call RegisterGPIBGlobals
526        End If

       ' Call the 32-bit DLL.
527        Call ibppc32(ud, v)

528        Call copy_ibvars
' <VB WATCH>
529        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
530        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibrd(ByVal ud As Integer, buf As String)
' <VB WATCH>
531        On Error GoTo vbwErrHandler
532        Const VBWPROCNAME = "VBIB32.ibrd"
533        If vbwProtector.vbwTraceProc Then
534            Dim vbwProtectorParameterString As String
535            If vbwProtector.vbwTraceParameters Then
536                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
537                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ") "
538            End If
539            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
540        End If
' </VB WATCH>
541        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
542        If (GPIBglobalsRegistered = 0) Then
543          Call RegisterGPIBGlobals
544        End If

545        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
546        Call ibrd32(ud, ByVal buf, cnt)

547        Call copy_ibvars
' <VB WATCH>
548        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
549        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibrda(ByVal ud As Integer, buf As String)
' <VB WATCH>
550        On Error GoTo vbwErrHandler
551        Const VBWPROCNAME = "VBIB32.ibrda"
552        If vbwProtector.vbwTraceProc Then
553            Dim vbwProtectorParameterString As String
554            If vbwProtector.vbwTraceParameters Then
555                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
556                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ") "
557            End If
558            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
559        End If
' </VB WATCH>
560        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
561        If (GPIBglobalsRegistered = 0) Then
562          Call RegisterGPIBGlobals
563        End If

564        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
565        Call ibrd32(ud, ByVal buf, cnt)

       ' When Visual Basic remapping buffer problem solved, use this:
       '    Call ibrda32(ud, ByVal buf, cnt)

566        Call copy_ibvars
' <VB WATCH>
567        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
568        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibrdf(ByVal ud As Integer, ByVal filename As String)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
569        On Error GoTo vbwErrHandler
570        Const VBWPROCNAME = "VBIB32.ibrdf"
571        If vbwProtector.vbwTraceProc Then
572            Dim vbwProtectorParameterString As String
573            If vbwProtector.vbwTraceParameters Then
574                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
575                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("filename", filename) & ") "
576            End If
577            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
578        End If
' </VB WATCH>
579        If (GPIBglobalsRegistered = 0) Then
580          Call RegisterGPIBGlobals
581        End If

       ' Call the 32-bit DLL.
582        Call ibrdf32(ud, ByVal filename)

583        Call copy_ibvars
' <VB WATCH>
584        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
585        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "filename", filename
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibrdi(ByVal ud As Integer, ibuf() As Integer, ByVal cnt As Long)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
586        On Error GoTo vbwErrHandler
587        Const VBWPROCNAME = "VBIB32.ibrdi"
588        If vbwProtector.vbwTraceProc Then
589            Dim vbwProtectorParameterString As String
590            If vbwProtector.vbwTraceParameters Then
591                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
592                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ibuf", ibuf) & ", "
593                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
594            End If
595            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
596        End If
' </VB WATCH>
597        If (GPIBglobalsRegistered = 0) Then
598          Call RegisterGPIBGlobals
599        End If

       ' Call the 32-bit DLL.
600        Call ibrd32(ud, ibuf(0), cnt)

601        Call copy_ibvars
' <VB WATCH>
602        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
603        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "ibuf", ibuf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibrdia(ByVal ud As Integer, ibuf() As Integer, ByVal cnt As Long)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
604        On Error GoTo vbwErrHandler
605        Const VBWPROCNAME = "VBIB32.ibrdia"
606        If vbwProtector.vbwTraceProc Then
607            Dim vbwProtectorParameterString As String
608            If vbwProtector.vbwTraceParameters Then
609                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
610                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ibuf", ibuf) & ", "
611                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
612            End If
613            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
614        End If
' </VB WATCH>
615        If (GPIBglobalsRegistered = 0) Then
616          Call RegisterGPIBGlobals
617        End If

       ' Call the 32-bit DLL.
618        Call ibrd32(ud, ibuf(0), cnt)

       ' When Visual Basic remapping buffer problem is solved, then use:
       '    Call ibrda32(u, ibuf(0), cnt)

619        Call copy_ibvars
' <VB WATCH>
620        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
621        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "ibuf", ibuf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibrpp(ByVal ud As Integer, ppr As Integer)
' <VB WATCH>
622        On Error GoTo vbwErrHandler
623        Const VBWPROCNAME = "VBIB32.ibrpp"
624        If vbwProtector.vbwTraceProc Then
625            Dim vbwProtectorParameterString As String
626            If vbwProtector.vbwTraceParameters Then
627                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
628                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ppr", ppr) & ") "
629            End If
630            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
631        End If
' </VB WATCH>
632        Static tmp_str As String * 2

       ' Check to see if GPIB Global variables are registered
633        If (GPIBglobalsRegistered = 0) Then
634          Call RegisterGPIBGlobals
635        End If

       ' Call the 32-bit DLL.
636        Call ibrpp32(ud, ByVal tmp_str)

637        ppr = Asc(tmp_str)

638        Call copy_ibvars
' <VB WATCH>
639        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
640        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "ppr", ppr
            vbwReportVariable "tmp_str", tmp_str
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibrsc(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
641        On Error GoTo vbwErrHandler
642        Const VBWPROCNAME = "VBIB32.ibrsc"
643        If vbwProtector.vbwTraceProc Then
644            Dim vbwProtectorParameterString As String
645            If vbwProtector.vbwTraceParameters Then
646                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
647                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
648            End If
649            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
650        End If
' </VB WATCH>
651        If (GPIBglobalsRegistered = 0) Then
652          Call RegisterGPIBGlobals
653        End If

       ' Call the 32-bit DLL.
654        Call ibrsc32(ud, v)

655        Call copy_ibvars
' <VB WATCH>
656        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
657        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibrsp(ByVal ud As Integer, spr As Integer)
' <VB WATCH>
658        On Error GoTo vbwErrHandler
659        Const VBWPROCNAME = "VBIB32.ibrsp"
660        If vbwProtector.vbwTraceProc Then
661            Dim vbwProtectorParameterString As String
662            If vbwProtector.vbwTraceParameters Then
663                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
664                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("spr", spr) & ") "
665            End If
666            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
667        End If
' </VB WATCH>
668        Static tmp_str As String * 2

       ' Check to see if GPIB Global variables are registered
669        If (GPIBglobalsRegistered = 0) Then
670          Call RegisterGPIBGlobals
671        End If

       ' Call the 32-bit DLL
672        Call ibrsp32(ud, ByVal tmp_str)

673        spr = Asc(tmp_str)

674        Call copy_ibvars
' <VB WATCH>
675        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
676        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "spr", spr
            vbwReportVariable "tmp_str", tmp_str
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibrsv(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
677        On Error GoTo vbwErrHandler
678        Const VBWPROCNAME = "VBIB32.ibrsv"
679        If vbwProtector.vbwTraceProc Then
680            Dim vbwProtectorParameterString As String
681            If vbwProtector.vbwTraceParameters Then
682                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
683                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
684            End If
685            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
686        End If
' </VB WATCH>
687        If (GPIBglobalsRegistered = 0) Then
688          Call RegisterGPIBGlobals
689        End If

       ' Call the 32-bit DLL.
690        Call ibrsv32(ud, v)

691        Call copy_ibvars
' <VB WATCH>
692        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
693        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibsad(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
694        On Error GoTo vbwErrHandler
695        Const VBWPROCNAME = "VBIB32.ibsad"
696        If vbwProtector.vbwTraceProc Then
697            Dim vbwProtectorParameterString As String
698            If vbwProtector.vbwTraceParameters Then
699                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
700                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
701            End If
702            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
703        End If
' </VB WATCH>
704        If (GPIBglobalsRegistered = 0) Then
705          Call RegisterGPIBGlobals
706        End If

       ' Call the 32-bit DLL.
707        Call ibsad32(ud, v)

708        Call copy_ibvars
' <VB WATCH>
709        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
710        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibsic(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
711        On Error GoTo vbwErrHandler
712        Const VBWPROCNAME = "VBIB32.ibsic"
713        If vbwProtector.vbwTraceProc Then
714            Dim vbwProtectorParameterString As String
715            If vbwProtector.vbwTraceParameters Then
716                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
717            End If
718            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
719        End If
' </VB WATCH>
720        If (GPIBglobalsRegistered = 0) Then
721          Call RegisterGPIBGlobals
722        End If

       ' Call the 32-bit DLL.
723        Call ibsic32(ud)

724        Call copy_ibvars
' <VB WATCH>
725        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
726        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibsre(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
727        On Error GoTo vbwErrHandler
728        Const VBWPROCNAME = "VBIB32.ibsre"
729        If vbwProtector.vbwTraceProc Then
730            Dim vbwProtectorParameterString As String
731            If vbwProtector.vbwTraceParameters Then
732                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
733                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
734            End If
735            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
736        End If
' </VB WATCH>
737        If (GPIBglobalsRegistered = 0) Then
738          Call RegisterGPIBGlobals
739        End If

       ' Call the 32-bit DLL.
740        Call ibsre32(ud, v)

741        Call copy_ibvars
' <VB WATCH>
742        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
743        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibstop(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
744        On Error GoTo vbwErrHandler
745        Const VBWPROCNAME = "VBIB32.ibstop"
746        If vbwProtector.vbwTraceProc Then
747            Dim vbwProtectorParameterString As String
748            If vbwProtector.vbwTraceParameters Then
749                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
750            End If
751            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
752        End If
' </VB WATCH>
753        If (GPIBglobalsRegistered = 0) Then
754          Call RegisterGPIBGlobals
755        End If

       ' Call the 32-bit DLL.
756        Call ibstop32(ud)

757        Call copy_ibvars
' <VB WATCH>
758        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
759        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibtmo(ByVal ud As Integer, ByVal v As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
760        On Error GoTo vbwErrHandler
761        Const VBWPROCNAME = "VBIB32.ibtmo"
762        If vbwProtector.vbwTraceProc Then
763            Dim vbwProtectorParameterString As String
764            If vbwProtector.vbwTraceParameters Then
765                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
766                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
767            End If
768            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
769        End If
' </VB WATCH>
770        If (GPIBglobalsRegistered = 0) Then
771          Call RegisterGPIBGlobals
772        End If

       ' Call the 32-bit DLL.
773        Call ibtmo32(ud, v)

774        Call copy_ibvars
' <VB WATCH>
775        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
776        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibtrg(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
777        On Error GoTo vbwErrHandler
778        Const VBWPROCNAME = "VBIB32.ibtrg"
779        If vbwProtector.vbwTraceProc Then
780            Dim vbwProtectorParameterString As String
781            If vbwProtector.vbwTraceParameters Then
782                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
783            End If
784            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
785        End If
' </VB WATCH>
786        If (GPIBglobalsRegistered = 0) Then
787          Call RegisterGPIBGlobals
788        End If

       ' Call 32-bit DLL.
789        Call ibtrg32(ud)

790        Call copy_ibvars
' <VB WATCH>
791        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
792        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibwait(ByVal ud As Integer, ByVal mask As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
793        On Error GoTo vbwErrHandler
794        Const VBWPROCNAME = "VBIB32.ibwait"
795        If vbwProtector.vbwTraceProc Then
796            Dim vbwProtectorParameterString As String
797            If vbwProtector.vbwTraceParameters Then
798                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
799                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("mask", mask) & ") "
800            End If
801            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
802        End If
' </VB WATCH>
803        If (GPIBglobalsRegistered = 0) Then
804          Call RegisterGPIBGlobals
805        End If

       ' Call the 32-bit DLL.
806        Call ibwait32(ud, mask)

807        Call copy_ibvars
' <VB WATCH>
808        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
809        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "mask", mask
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibwrt(ByVal ud As Integer, ByVal buf As String)
' <VB WATCH>
810        On Error GoTo vbwErrHandler
811        Const VBWPROCNAME = "VBIB32.ibwrt"
812        If vbwProtector.vbwTraceProc Then
813            Dim vbwProtectorParameterString As String
814            If vbwProtector.vbwTraceParameters Then
815                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
816                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ") "
817            End If
818            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
819        End If
' </VB WATCH>
820        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
821        If (GPIBglobalsRegistered = 0) Then
822          Call RegisterGPIBGlobals
823        End If

824        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
825        Call ibwrt32(ud, ByVal buf, cnt)

826        Call copy_ibvars
' <VB WATCH>
827        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
828        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibwrta(ByVal ud As Integer, ByVal buf As String)
' <VB WATCH>
829        On Error GoTo vbwErrHandler
830        Const VBWPROCNAME = "VBIB32.ibwrta"
831        If vbwProtector.vbwTraceProc Then
832            Dim vbwProtectorParameterString As String
833            If vbwProtector.vbwTraceParameters Then
834                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
835                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ") "
836            End If
837            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
838        End If
' </VB WATCH>
839        Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
840        If (GPIBglobalsRegistered = 0) Then
841          Call RegisterGPIBGlobals
842        End If

843        cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
844        Call ibwrt32(ud, ByVal buf, cnt)

       ' When Visual Basic remapping buffer problem is solved, use this:
       '    Call ibwrta32(ud, ByVal buf, cnt)

845        Call copy_ibvars
' <VB WATCH>
846        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
847        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibwrtf(ByVal ud As Integer, ByVal filename As String)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
848        On Error GoTo vbwErrHandler
849        Const VBWPROCNAME = "VBIB32.ibwrtf"
850        If vbwProtector.vbwTraceProc Then
851            Dim vbwProtectorParameterString As String
852            If vbwProtector.vbwTraceParameters Then
853                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
854                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("filename", filename) & ") "
855            End If
856            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
857        End If
' </VB WATCH>
858        If (GPIBglobalsRegistered = 0) Then
859          Call RegisterGPIBGlobals
860        End If

       ' Call the 32-bit DLL.
861        Call ibwrtf32(ud, ByVal filename)

862        Call copy_ibvars
' <VB WATCH>
863        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
864        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "filename", filename
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibwrti(ByVal ud As Integer, ByRef ibuf() As Integer, ByVal cnt As Long)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
865        On Error GoTo vbwErrHandler
866        Const VBWPROCNAME = "VBIB32.ibwrti"
867        If vbwProtector.vbwTraceProc Then
868            Dim vbwProtectorParameterString As String
869            If vbwProtector.vbwTraceParameters Then
870                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
871                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ibuf", ibuf) & ", "
872                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
873            End If
874            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
875        End If
' </VB WATCH>
876        If (GPIBglobalsRegistered = 0) Then
877          Call RegisterGPIBGlobals
878        End If

       ' Call the 32-bit DLL.
879        Call ibwrt32(ud, ibuf(0), cnt)

880        Call copy_ibvars
' <VB WATCH>
881        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
882        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "ibuf", ibuf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ibwrtia(ByVal ud As Integer, ByRef ibuf() As Integer, ByVal cnt As Long)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
883        On Error GoTo vbwErrHandler
884        Const VBWPROCNAME = "VBIB32.ibwrtia"
885        If vbwProtector.vbwTraceProc Then
886            Dim vbwProtectorParameterString As String
887            If vbwProtector.vbwTraceParameters Then
888                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
889                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ibuf", ibuf) & ", "
890                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
891            End If
892            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
893        End If
' </VB WATCH>
894        If (GPIBglobalsRegistered = 0) Then
895          Call RegisterGPIBGlobals
896        End If

       ' Call the 32-bit DLL.
897        Call ibwrt32(ud, ibuf(0), cnt)

       ' When Visual Basic remapping buffer problem is solved, use this:
       '    Call ibwrta32(ud, ibuf(0), cnt)

898        Call copy_ibvars
' <VB WATCH>
899        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
900        Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "ibuf", ibuf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Function ilask(ByVal ud As Integer, ByVal opt As Integer, rval As Integer) As Integer
' <VB WATCH>
901        On Error GoTo vbwErrHandler
902        Const VBWPROCNAME = "VBIB32.ilask"
903        If vbwProtector.vbwTraceProc Then
904            Dim vbwProtectorParameterString As String
905            If vbwProtector.vbwTraceParameters Then
906                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
907                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("opt", opt) & ", "
908                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("rval", rval) & ") "
909            End If
910            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
911        End If
' </VB WATCH>
912        Dim tmprval As Long

       ' Check to see if GPIB Global variables are registered
913        If (GPIBglobalsRegistered = 0) Then
914          Call RegisterGPIBGlobals
915        End If

       ' Call the 32-bit DLL.
916        ilask = ConvertLongToInt(ibask32(ud, opt, tmprval))

917        rval = ConvertLongToInt(tmprval)

918        Call copy_ibvars
' <VB WATCH>
919        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
920        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "opt", opt
            vbwReportVariable "rval", rval
            vbwReportVariable "tmprval", tmprval
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilbna(ByVal ud As Integer, ByVal udname As String) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
921        On Error GoTo vbwErrHandler
922        Const VBWPROCNAME = "VBIB32.ilbna"
923        If vbwProtector.vbwTraceProc Then
924            Dim vbwProtectorParameterString As String
925            If vbwProtector.vbwTraceParameters Then
926                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
927                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("udname", udname) & ") "
928            End If
929            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
930        End If
' </VB WATCH>
931        If (GPIBglobalsRegistered = 0) Then
932          Call RegisterGPIBGlobals
933        End If

       ' Call the 32-bit DLL.
934        ilbna = ConvertLongToInt(ibbna32(ud, ByVal udname))

935        Call copy_ibvars
' <VB WATCH>
936        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
937        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "udname", udname
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilcac(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
938        On Error GoTo vbwErrHandler
939        Const VBWPROCNAME = "VBIB32.ilcac"
940        If vbwProtector.vbwTraceProc Then
941            Dim vbwProtectorParameterString As String
942            If vbwProtector.vbwTraceParameters Then
943                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
944                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
945            End If
946            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
947        End If
' </VB WATCH>
948        If (GPIBglobalsRegistered = 0) Then
949          Call RegisterGPIBGlobals
950        End If

       ' Call the 32-bit DLL.
951        ilcac = ConvertLongToInt(ibcac32(ud, v))

952        Call copy_ibvars
' <VB WATCH>
953        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
954        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilclr(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
955        On Error GoTo vbwErrHandler
956        Const VBWPROCNAME = "VBIB32.ilclr"
957        If vbwProtector.vbwTraceProc Then
958            Dim vbwProtectorParameterString As String
959            If vbwProtector.vbwTraceParameters Then
960                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
961            End If
962            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
963        End If
' </VB WATCH>
964        If (GPIBglobalsRegistered = 0) Then
965          Call RegisterGPIBGlobals
966        End If

       ' Call the 32-bit DLL.
967        ilclr = ConvertLongToInt(ibclr32(ud))

968        Call copy_ibvars
' <VB WATCH>
969        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
970        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilcmd(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
971        On Error GoTo vbwErrHandler
972        Const VBWPROCNAME = "VBIB32.ilcmd"
973        If vbwProtector.vbwTraceProc Then
974            Dim vbwProtectorParameterString As String
975            If vbwProtector.vbwTraceParameters Then
976                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
977                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
978                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
979            End If
980            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
981        End If
' </VB WATCH>
982        If (GPIBglobalsRegistered = 0) Then
983          Call RegisterGPIBGlobals
984        End If

       ' Call the 32-bit DLL.
985        ilcmd = ConvertLongToInt(ibcmd32(ud, ByVal buf, cnt))

986        Call copy_ibvars
' <VB WATCH>
987        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
988        Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilcmda(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
989        On Error GoTo vbwErrHandler
990        Const VBWPROCNAME = "VBIB32.ilcmda"
991        If vbwProtector.vbwTraceProc Then
992            Dim vbwProtectorParameterString As String
993            If vbwProtector.vbwTraceParameters Then
994                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
995                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
996                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
997            End If
998            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
999        End If
' </VB WATCH>
1000       If (GPIBglobalsRegistered = 0) Then
1001         Call RegisterGPIBGlobals
1002       End If

       ' Call the 32-bit DLL.
1003       ilcmda = ConvertLongToInt(ibcmd32(ud, ByVal buf, cnt))

       ' When Visual Basic remapping buffer problem is solved, use this:
       '    ilcmda = ConvertLongToInt(ibcmda32(ud, ByVal buf, cnt))

1004       Call copy_ibvars
' <VB WATCH>
1005       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1006       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilconfig(ByVal bdid As Integer, ByVal opt As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1007       On Error GoTo vbwErrHandler
1008       Const VBWPROCNAME = "VBIB32.ilconfig"
1009       If vbwProtector.vbwTraceProc Then
1010           Dim vbwProtectorParameterString As String
1011           If vbwProtector.vbwTraceParameters Then
1012               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("bdid", bdid) & ", "
1013               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("opt", opt) & ", "
1014               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1015           End If
1016           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1017       End If
' </VB WATCH>
1018       If (GPIBglobalsRegistered = 0) Then
1019         Call RegisterGPIBGlobals
1020       End If

       ' Call the 32-bit DLL.
1021       ilconfig = ConvertLongToInt(ibconfig32(bdid, opt, v))

1022       Call copy_ibvars
' <VB WATCH>
1023       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1024       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "bdid", bdid
            vbwReportVariable "opt", opt
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ildev(ByVal bdid As Integer, ByVal pad As Integer, ByVal sad As Integer, ByVal tmo As Integer, ByVal eot As Integer, ByVal eos As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1025       On Error GoTo vbwErrHandler
1026       Const VBWPROCNAME = "VBIB32.ildev"
1027       If vbwProtector.vbwTraceProc Then
1028           Dim vbwProtectorParameterString As String
1029           If vbwProtector.vbwTraceParameters Then
1030               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("bdid", bdid) & ", "
1031               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("pad", pad) & ", "
1032               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sad", sad) & ", "
1033               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("tmo", tmo) & ", "
1034               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("eot", eot) & ", "
1035               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("eos", eos) & ") "
1036           End If
1037           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1038       End If
' </VB WATCH>
1039       If (GPIBglobalsRegistered = 0) Then
1040         Call RegisterGPIBGlobals
1041       End If

       ' Call the 32-bit DLL.
1042       ildev = ConvertLongToInt(ibdev32(bdid, pad, sad, tmo, eot, eos))

1043       Call copy_ibvars
' <VB WATCH>
1044       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1045       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "bdid", bdid
            vbwReportVariable "pad", pad
            vbwReportVariable "sad", sad
            vbwReportVariable "tmo", tmo
            vbwReportVariable "eot", eot
            vbwReportVariable "eos", eos
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ildma(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1046       On Error GoTo vbwErrHandler
1047       Const VBWPROCNAME = "VBIB32.ildma"
1048       If vbwProtector.vbwTraceProc Then
1049           Dim vbwProtectorParameterString As String
1050           If vbwProtector.vbwTraceParameters Then
1051               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1052               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1053           End If
1054           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1055       End If
' </VB WATCH>
1056       If (GPIBglobalsRegistered = 0) Then
1057         Call RegisterGPIBGlobals
1058       End If

       ' Call the 32-bit DLL.
1059       ildma = ConvertLongToInt(ibdma32(ud, v))

1060       Call copy_ibvars
' <VB WATCH>
1061       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1062       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ileos(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1063       On Error GoTo vbwErrHandler
1064       Const VBWPROCNAME = "VBIB32.ileos"
1065       If vbwProtector.vbwTraceProc Then
1066           Dim vbwProtectorParameterString As String
1067           If vbwProtector.vbwTraceParameters Then
1068               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1069               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1070           End If
1071           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1072       End If
' </VB WATCH>
1073       If (GPIBglobalsRegistered = 0) Then
1074         Call RegisterGPIBGlobals
1075       End If

       ' Call the 32-bit DLL.
1076       ileos = ConvertLongToInt(ibeos32(ud, v))

1077       Call copy_ibvars
' <VB WATCH>
1078       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1079       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ileot(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1080       On Error GoTo vbwErrHandler
1081       Const VBWPROCNAME = "VBIB32.ileot"
1082       If vbwProtector.vbwTraceProc Then
1083           Dim vbwProtectorParameterString As String
1084           If vbwProtector.vbwTraceParameters Then
1085               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1086               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1087           End If
1088           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1089       End If
' </VB WATCH>
1090       If (GPIBglobalsRegistered = 0) Then
1091         Call RegisterGPIBGlobals
1092       End If

       ' Call the 32-bit DLL.
1093       ileot = ConvertLongToInt(ibeot32(ud, v))

1094       Call copy_ibvars
' <VB WATCH>
1095       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1096       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilfind(ByVal udname As String) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1097       On Error GoTo vbwErrHandler
1098       Const VBWPROCNAME = "VBIB32.ilfind"
1099       If vbwProtector.vbwTraceProc Then
1100           Dim vbwProtectorParameterString As String
1101           If vbwProtector.vbwTraceParameters Then
1102               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("udname", udname) & ") "
1103           End If
1104           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1105       End If
' </VB WATCH>
1106       If (GPIBglobalsRegistered = 0) Then
1107         Call RegisterGPIBGlobals
1108       End If

       ' Call the 32-bit DLL.
1109       ilfind = ConvertLongToInt(ibfind32(ByVal udname))

1110       Call copy_ibvars
' <VB WATCH>
1111       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1112       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "udname", udname
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilgts(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1113       On Error GoTo vbwErrHandler
1114       Const VBWPROCNAME = "VBIB32.ilgts"
1115       If vbwProtector.vbwTraceProc Then
1116           Dim vbwProtectorParameterString As String
1117           If vbwProtector.vbwTraceParameters Then
1118               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1119               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1120           End If
1121           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1122       End If
' </VB WATCH>
1123       If (GPIBglobalsRegistered = 0) Then
1124         Call RegisterGPIBGlobals
1125       End If

       ' Call the 32-bit DLL.
1126       ilgts = ConvertLongToInt(ibgts32(ud, v))

1127       Call copy_ibvars
' <VB WATCH>
1128       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1129       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilist(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1130       On Error GoTo vbwErrHandler
1131       Const VBWPROCNAME = "VBIB32.ilist"
1132       If vbwProtector.vbwTraceProc Then
1133           Dim vbwProtectorParameterString As String
1134           If vbwProtector.vbwTraceParameters Then
1135               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1136               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1137           End If
1138           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1139       End If
' </VB WATCH>
1140       If (GPIBglobalsRegistered = 0) Then
1141         Call RegisterGPIBGlobals
1142       End If

       ' Call the 32-bit DLL.
1143       ilist = ConvertLongToInt(ibist32(ud, v))

1144       Call copy_ibvars
' <VB WATCH>
1145       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1146       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function illck(ByVal ud As Integer, ByVal v As Integer, ByVal LockWaitTime As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1147       On Error GoTo vbwErrHandler
1148       Const VBWPROCNAME = "VBIB32.illck"
1149       If vbwProtector.vbwTraceProc Then
1150           Dim vbwProtectorParameterString As String
1151           If vbwProtector.vbwTraceParameters Then
1152               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1153               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ", "
1154               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("LockWaitTime", LockWaitTime) & ") "
1155           End If
1156           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1157       End If
' </VB WATCH>
1158       If (GPIBglobalsRegistered = 0) Then
1159         Call RegisterGPIBGlobals
1160       End If

       ' Call the 32-bit DLL.
1161       illck = ConvertLongToInt(iblck32(ud, v, LockWaitTime, ByVal 0))

1162       Call copy_ibvars
' <VB WATCH>
1163       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1164       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportVariable "LockWaitTime", LockWaitTime
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
       
Function illines(ByVal ud As Integer, lines As Integer) As Integer
' <VB WATCH>
1165       On Error GoTo vbwErrHandler
1166       Const VBWPROCNAME = "VBIB32.illines"
1167       If vbwProtector.vbwTraceProc Then
1168           Dim vbwProtectorParameterString As String
1169           If vbwProtector.vbwTraceParameters Then
1170               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1171               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("lines", lines) & ") "
1172           End If
1173           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1174       End If
' </VB WATCH>
1175       Dim tmplines As Long

       ' Check to see if GPIB Global variables are registered
1176       If (GPIBglobalsRegistered = 0) Then
1177         Call RegisterGPIBGlobals
1178       End If

       ' Call the 32-bit DLL.
1179       illines = ConvertLongToInt(iblines32(ud, tmplines))

1180       lines = ConvertLongToInt(tmplines)

1181       Call copy_ibvars
' <VB WATCH>
1182       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1183       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "lines", lines
            vbwReportVariable "tmplines", tmplines
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function illn(ByVal ud As Integer, ByVal pad As Integer, ByVal sad As Integer, ln As Integer) As Integer
' <VB WATCH>
1184       On Error GoTo vbwErrHandler
1185       Const VBWPROCNAME = "VBIB32.illn"
1186       If vbwProtector.vbwTraceProc Then
1187           Dim vbwProtectorParameterString As String
1188           If vbwProtector.vbwTraceParameters Then
1189               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1190               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("pad", pad) & ", "
1191               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sad", sad) & ", "
1192               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ln", ln) & ") "
1193           End If
1194           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1195       End If
' </VB WATCH>
1196       Dim tmpln As Long

       ' Check to see if GPIB Global variables are registered
1197       If (GPIBglobalsRegistered = 0) Then
1198         Call RegisterGPIBGlobals
1199       End If

       ' Call the 32-bit DLL.
1200       illn = ConvertLongToInt(ibln32(ud, pad, sad, tmpln))

1201       ln = ConvertLongToInt(tmpln)

1202       Call copy_ibvars
' <VB WATCH>
1203       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1204       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "pad", pad
            vbwReportVariable "sad", sad
            vbwReportVariable "ln", ln
            vbwReportVariable "tmpln", tmpln
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function illoc(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1205       On Error GoTo vbwErrHandler
1206       Const VBWPROCNAME = "VBIB32.illoc"
1207       If vbwProtector.vbwTraceProc Then
1208           Dim vbwProtectorParameterString As String
1209           If vbwProtector.vbwTraceParameters Then
1210               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
1211           End If
1212           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1213       End If
' </VB WATCH>
1214       If (GPIBglobalsRegistered = 0) Then
1215         Call RegisterGPIBGlobals
1216       End If

       ' Call the 32-bit DLL.
1217       illoc = ConvertLongToInt(ibloc32(ud))

1218       Call copy_ibvars
' <VB WATCH>
1219       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1220       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilonl(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1221       On Error GoTo vbwErrHandler
1222       Const VBWPROCNAME = "VBIB32.ilonl"
1223       If vbwProtector.vbwTraceProc Then
1224           Dim vbwProtectorParameterString As String
1225           If vbwProtector.vbwTraceParameters Then
1226               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1227               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1228           End If
1229           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1230       End If
' </VB WATCH>
1231       If (GPIBglobalsRegistered = 0) Then
1232         Call RegisterGPIBGlobals
1233       End If

       ' Call the 32-bit DLL.
1234       ilonl = ConvertLongToInt(ibonl32(ud, v))

1235       Call copy_ibvars
' <VB WATCH>
1236       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1237       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilpad(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1238       On Error GoTo vbwErrHandler
1239       Const VBWPROCNAME = "VBIB32.ilpad"
1240       If vbwProtector.vbwTraceProc Then
1241           Dim vbwProtectorParameterString As String
1242           If vbwProtector.vbwTraceParameters Then
1243               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1244               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1245           End If
1246           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1247       End If
' </VB WATCH>
1248       If (GPIBglobalsRegistered = 0) Then
1249         Call RegisterGPIBGlobals
1250       End If

       ' Call the 32-bit DLL.
1251       ilpad = ConvertLongToInt(ibpad32(ud, v))

1252       Call copy_ibvars
' <VB WATCH>
1253       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1254       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilpct(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1255       On Error GoTo vbwErrHandler
1256       Const VBWPROCNAME = "VBIB32.ilpct"
1257       If vbwProtector.vbwTraceProc Then
1258           Dim vbwProtectorParameterString As String
1259           If vbwProtector.vbwTraceParameters Then
1260               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
1261           End If
1262           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1263       End If
' </VB WATCH>
1264       If (GPIBglobalsRegistered = 0) Then
1265         Call RegisterGPIBGlobals
1266       End If

       ' Call the 32-bit DLL.
1267       ilpct = ConvertLongToInt(ibpct32(ud))

1268       Call copy_ibvars
' <VB WATCH>
1269       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1270       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilppc(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1271       On Error GoTo vbwErrHandler
1272       Const VBWPROCNAME = "VBIB32.ilppc"
1273       If vbwProtector.vbwTraceProc Then
1274           Dim vbwProtectorParameterString As String
1275           If vbwProtector.vbwTraceParameters Then
1276               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1277               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1278           End If
1279           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1280       End If
' </VB WATCH>
1281       If (GPIBglobalsRegistered = 0) Then
1282         Call RegisterGPIBGlobals
1283       End If

       ' Call the 32-bit DLL.
1284       ilppc = ConvertLongToInt(ibppc32(ud, v))

1285       Call copy_ibvars
' <VB WATCH>
1286       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1287       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilrd(ByVal ud As Integer, buf As String, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1288       On Error GoTo vbwErrHandler
1289       Const VBWPROCNAME = "VBIB32.ilrd"
1290       If vbwProtector.vbwTraceProc Then
1291           Dim vbwProtectorParameterString As String
1292           If vbwProtector.vbwTraceParameters Then
1293               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1294               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
1295               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
1296           End If
1297           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1298       End If
' </VB WATCH>
1299       If (GPIBglobalsRegistered = 0) Then
1300         Call RegisterGPIBGlobals
1301       End If

       ' Call the 32-bit DLL.
1302       ilrd = ConvertLongToInt(ibrd32(ud, ByVal buf, cnt))

1303       Call copy_ibvars
' <VB WATCH>
1304       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1305       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilrda(ByVal ud As Integer, buf As String, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1306       On Error GoTo vbwErrHandler
1307       Const VBWPROCNAME = "VBIB32.ilrda"
1308       If vbwProtector.vbwTraceProc Then
1309           Dim vbwProtectorParameterString As String
1310           If vbwProtector.vbwTraceParameters Then
1311               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1312               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
1313               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
1314           End If
1315           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1316       End If
' </VB WATCH>
1317       If (GPIBglobalsRegistered = 0) Then
1318         Call RegisterGPIBGlobals
1319       End If

       ' Call the 32-bit DLL.
1320       ilrda = ConvertLongToInt(ibrd32(ud, ByVal buf, cnt))

       ' When Visual Basic remapping buffer problem solved, use this:
       '    ilrda = ConvertLongToInt(ibrda32(ud, ByVal buf, cnt))

1321       Call copy_ibvars
' <VB WATCH>
1322       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1323       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilrdf(ByVal ud As Integer, ByVal filename As String) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1324       On Error GoTo vbwErrHandler
1325       Const VBWPROCNAME = "VBIB32.ilrdf"
1326       If vbwProtector.vbwTraceProc Then
1327           Dim vbwProtectorParameterString As String
1328           If vbwProtector.vbwTraceParameters Then
1329               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1330               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("filename", filename) & ") "
1331           End If
1332           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1333       End If
' </VB WATCH>
1334       If (GPIBglobalsRegistered = 0) Then
1335         Call RegisterGPIBGlobals
1336       End If

       ' Call the 32-bit DLL.
1337       ilrdf = ConvertLongToInt(ibrdf32(ud, ByVal filename))

1338       Call copy_ibvars
' <VB WATCH>
1339       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1340       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "filename", filename
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilrdi(ByVal ud As Integer, ibuf() As Integer, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1341       On Error GoTo vbwErrHandler
1342       Const VBWPROCNAME = "VBIB32.ilrdi"
1343       If vbwProtector.vbwTraceProc Then
1344           Dim vbwProtectorParameterString As String
1345           If vbwProtector.vbwTraceParameters Then
1346               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1347               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ibuf", ibuf) & ", "
1348               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
1349           End If
1350           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1351       End If
' </VB WATCH>
1352       If (GPIBglobalsRegistered = 0) Then
1353         Call RegisterGPIBGlobals
1354       End If

       ' Call the 32-bit DLL.
1355       ilrdi = ConvertLongToInt(ibrd32(ud, ibuf(0), cnt))

1356       Call copy_ibvars
' <VB WATCH>
1357       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1358       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "ibuf", ibuf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilrdia(ByVal ud As Integer, ibuf() As Integer, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1359       On Error GoTo vbwErrHandler
1360       Const VBWPROCNAME = "VBIB32.ilrdia"
1361       If vbwProtector.vbwTraceProc Then
1362           Dim vbwProtectorParameterString As String
1363           If vbwProtector.vbwTraceParameters Then
1364               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1365               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ibuf", ibuf) & ", "
1366               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
1367           End If
1368           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1369       End If
' </VB WATCH>
1370       If (GPIBglobalsRegistered = 0) Then
1371         Call RegisterGPIBGlobals
1372       End If

       ' Call the 32-bit DLL.
1373       ilrdia = ConvertLongToInt(ibrd32(ud, ibuf(0), cnt))

       ' When Visual Basic remapping buffer problem solved, use this:
       '    ilrdia = ConvertLongToInt(ibrda32(ud, ibuf(0), cnt))

1374       Call copy_ibvars
' <VB WATCH>
1375       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1376       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "ibuf", ibuf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilrpp(ByVal ud As Integer, ppr As Integer) As Integer
' <VB WATCH>
1377       On Error GoTo vbwErrHandler
1378       Const VBWPROCNAME = "VBIB32.ilrpp"
1379       If vbwProtector.vbwTraceProc Then
1380           Dim vbwProtectorParameterString As String
1381           If vbwProtector.vbwTraceParameters Then
1382               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1383               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ppr", ppr) & ") "
1384           End If
1385           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1386       End If
' </VB WATCH>
1387       Static tmp_str As String * 2

       ' Check to see if GPIB Global variables are registered
1388       If (GPIBglobalsRegistered = 0) Then
1389         Call RegisterGPIBGlobals
1390       End If

       ' Call the 32-bit DLL.
1391       ilrpp = ConvertLongToInt(ibrpp32(ud, ByVal tmp_str))

1392       ppr = Asc(tmp_str)

1393       Call copy_ibvars
' <VB WATCH>
1394       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1395       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "ppr", ppr
            vbwReportVariable "tmp_str", tmp_str
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilrsc(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1396       On Error GoTo vbwErrHandler
1397       Const VBWPROCNAME = "VBIB32.ilrsc"
1398       If vbwProtector.vbwTraceProc Then
1399           Dim vbwProtectorParameterString As String
1400           If vbwProtector.vbwTraceParameters Then
1401               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1402               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1403           End If
1404           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1405       End If
' </VB WATCH>
1406       If (GPIBglobalsRegistered = 0) Then
1407         Call RegisterGPIBGlobals
1408       End If

       '  Call the 32-bit DLL.
1409       ilrsc = ConvertLongToInt(ibrsc32(ud, v))

1410       Call copy_ibvars
' <VB WATCH>
1411       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1412       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilrsp(ByVal ud As Integer, spr As Integer) As Integer
' <VB WATCH>
1413       On Error GoTo vbwErrHandler
1414       Const VBWPROCNAME = "VBIB32.ilrsp"
1415       If vbwProtector.vbwTraceProc Then
1416           Dim vbwProtectorParameterString As String
1417           If vbwProtector.vbwTraceParameters Then
1418               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1419               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("spr", spr) & ") "
1420           End If
1421           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1422       End If
' </VB WATCH>
1423       Static tmp_str As String * 2

       ' Check to see if GPIB Global variables are registered
1424       If (GPIBglobalsRegistered = 0) Then
1425         Call RegisterGPIBGlobals
1426       End If

       ' Call the 32-bit DLL
1427       ilrsp = ConvertLongToInt(ibrsp32(ud, ByVal tmp_str))

1428       spr = Asc(tmp_str)

1429       Call copy_ibvars
' <VB WATCH>
1430       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1431       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "spr", spr
            vbwReportVariable "tmp_str", tmp_str
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilrsv(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1432       On Error GoTo vbwErrHandler
1433       Const VBWPROCNAME = "VBIB32.ilrsv"
1434       If vbwProtector.vbwTraceProc Then
1435           Dim vbwProtectorParameterString As String
1436           If vbwProtector.vbwTraceParameters Then
1437               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1438               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1439           End If
1440           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1441       End If
' </VB WATCH>
1442       If (GPIBglobalsRegistered = 0) Then
1443         Call RegisterGPIBGlobals
1444       End If

       ' Call the 32-bit DLL.
1445       ilrsv = ConvertLongToInt(ibrsv32(ud, v))

1446       Call copy_ibvars
' <VB WATCH>
1447       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1448       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilsad(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1449       On Error GoTo vbwErrHandler
1450       Const VBWPROCNAME = "VBIB32.ilsad"
1451       If vbwProtector.vbwTraceProc Then
1452           Dim vbwProtectorParameterString As String
1453           If vbwProtector.vbwTraceParameters Then
1454               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1455               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1456           End If
1457           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1458       End If
' </VB WATCH>
1459       If (GPIBglobalsRegistered = 0) Then
1460         Call RegisterGPIBGlobals
1461       End If

       '  Call the 32-bit DLL.
1462       ilsad = ConvertLongToInt(ibsad32(ud, v))

1463       Call copy_ibvars
' <VB WATCH>
1464       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1465       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilsic(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1466       On Error GoTo vbwErrHandler
1467       Const VBWPROCNAME = "VBIB32.ilsic"
1468       If vbwProtector.vbwTraceProc Then
1469           Dim vbwProtectorParameterString As String
1470           If vbwProtector.vbwTraceParameters Then
1471               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
1472           End If
1473           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1474       End If
' </VB WATCH>
1475       If (GPIBglobalsRegistered = 0) Then
1476         Call RegisterGPIBGlobals
1477       End If

       '  Call the 32-bit DLL.
1478       ilsic = ConvertLongToInt(ibsic32(ud))

1479       Call copy_ibvars
' <VB WATCH>
1480       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1481       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilsre(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1482       On Error GoTo vbwErrHandler
1483       Const VBWPROCNAME = "VBIB32.ilsre"
1484       If vbwProtector.vbwTraceProc Then
1485           Dim vbwProtectorParameterString As String
1486           If vbwProtector.vbwTraceParameters Then
1487               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1488               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1489           End If
1490           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1491       End If
' </VB WATCH>
1492       If (GPIBglobalsRegistered = 0) Then
1493         Call RegisterGPIBGlobals
1494       End If

       '  Call the 32-bit DLL.
1495       ilsre = ConvertLongToInt(ibsre32(ud, v))

1496       Call copy_ibvars
' <VB WATCH>
1497       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1498       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilstop(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1499       On Error GoTo vbwErrHandler
1500       Const VBWPROCNAME = "VBIB32.ilstop"
1501       If vbwProtector.vbwTraceProc Then
1502           Dim vbwProtectorParameterString As String
1503           If vbwProtector.vbwTraceParameters Then
1504               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
1505           End If
1506           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1507       End If
' </VB WATCH>
1508       If (GPIBglobalsRegistered = 0) Then
1509         Call RegisterGPIBGlobals
1510       End If

       '  Call the 32-bit DLL.
1511       ilstop = ConvertLongToInt(ibstop32(ud))

1512       Call copy_ibvars
' <VB WATCH>
1513       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1514       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function iltmo(ByVal ud As Integer, ByVal v As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1515       On Error GoTo vbwErrHandler
1516       Const VBWPROCNAME = "VBIB32.iltmo"
1517       If vbwProtector.vbwTraceProc Then
1518           Dim vbwProtectorParameterString As String
1519           If vbwProtector.vbwTraceParameters Then
1520               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1521               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("v", v) & ") "
1522           End If
1523           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1524       End If
' </VB WATCH>
1525       If (GPIBglobalsRegistered = 0) Then
1526         Call RegisterGPIBGlobals
1527       End If

       '  Call the 32-bit DLL.
1528       iltmo = ConvertLongToInt(ibtmo32(ud, v))

1529       Call copy_ibvars
' <VB WATCH>
1530       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1531       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "v", v
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function iltrg(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1532       On Error GoTo vbwErrHandler
1533       Const VBWPROCNAME = "VBIB32.iltrg"
1534       If vbwProtector.vbwTraceProc Then
1535           Dim vbwProtectorParameterString As String
1536           If vbwProtector.vbwTraceParameters Then
1537               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
1538           End If
1539           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1540       End If
' </VB WATCH>
1541       If (GPIBglobalsRegistered = 0) Then
1542         Call RegisterGPIBGlobals
1543       End If

       ' Call 32-bit DLL.
1544       iltrg = ConvertLongToInt(ibtrg32(ud))

1545       Call copy_ibvars
' <VB WATCH>
1546       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1547       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilwait(ByVal ud As Integer, ByVal mask As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1548       On Error GoTo vbwErrHandler
1549       Const VBWPROCNAME = "VBIB32.ilwait"
1550       If vbwProtector.vbwTraceProc Then
1551           Dim vbwProtectorParameterString As String
1552           If vbwProtector.vbwTraceParameters Then
1553               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1554               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("mask", mask) & ") "
1555           End If
1556           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1557       End If
' </VB WATCH>
1558       If (GPIBglobalsRegistered = 0) Then
1559         Call RegisterGPIBGlobals
1560       End If

       ' Call the 32-bit DLL.
1561       ilwait = ConvertLongToInt(ibwait32(ud, mask))

1562       Call copy_ibvars
' <VB WATCH>
1563       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1564       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "mask", mask
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilwrt(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1565       On Error GoTo vbwErrHandler
1566       Const VBWPROCNAME = "VBIB32.ilwrt"
1567       If vbwProtector.vbwTraceProc Then
1568           Dim vbwProtectorParameterString As String
1569           If vbwProtector.vbwTraceParameters Then
1570               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1571               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
1572               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
1573           End If
1574           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1575       End If
' </VB WATCH>
1576       If (GPIBglobalsRegistered = 0) Then
1577         Call RegisterGPIBGlobals
1578       End If

       ' Call the 32-bit DLL.
1579       ilwrt = ConvertLongToInt(ibwrt32(ud, ByVal buf, cnt))

1580       Call copy_ibvars
' <VB WATCH>
1581       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1582       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilwrta(ByVal ud As Integer, ByVal buf As String, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1583       On Error GoTo vbwErrHandler
1584       Const VBWPROCNAME = "VBIB32.ilwrta"
1585       If vbwProtector.vbwTraceProc Then
1586           Dim vbwProtectorParameterString As String
1587           If vbwProtector.vbwTraceParameters Then
1588               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1589               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
1590               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
1591           End If
1592           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1593       End If
' </VB WATCH>
1594       If (GPIBglobalsRegistered = 0) Then
1595         Call RegisterGPIBGlobals
1596       End If

       ' Call the 32-bit DLL.
1597       ilwrta = ConvertLongToInt(ibwrt32(ud, ByVal buf, cnt))

       ' When the Visual Basic remapping solved, use this:
       '    ilwrta = ConvertLongToInt(ibwrta32(ud, ByVal buf, cnt))

1598       Call copy_ibvars

' <VB WATCH>
1599       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1600       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "buf", buf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilwrtf(ByVal ud As Integer, ByVal filename As String) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1601       On Error GoTo vbwErrHandler
1602       Const VBWPROCNAME = "VBIB32.ilwrtf"
1603       If vbwProtector.vbwTraceProc Then
1604           Dim vbwProtectorParameterString As String
1605           If vbwProtector.vbwTraceParameters Then
1606               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1607               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("filename", filename) & ") "
1608           End If
1609           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1610       End If
' </VB WATCH>
1611       If (GPIBglobalsRegistered = 0) Then
1612         Call RegisterGPIBGlobals
1613       End If

       ' Call the 32-bit DLL.
1614       ilwrtf = ConvertLongToInt(ibwrtf32(ud, ByVal filename))

1615       Call copy_ibvars
' <VB WATCH>
1616       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1617       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "filename", filename
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilwrti(ByVal ud As Integer, ByRef ibuf() As Integer, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1618       On Error GoTo vbwErrHandler
1619       Const VBWPROCNAME = "VBIB32.ilwrti"
1620       If vbwProtector.vbwTraceProc Then
1621           Dim vbwProtectorParameterString As String
1622           If vbwProtector.vbwTraceParameters Then
1623               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1624               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ibuf", ibuf) & ", "
1625               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
1626           End If
1627           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1628       End If
' </VB WATCH>
1629       If (GPIBglobalsRegistered = 0) Then
1630         Call RegisterGPIBGlobals
1631       End If

       ' Call the 32-bit DLL.
1632       ilwrti = ConvertLongToInt(ibwrt32(ud, ibuf(0), cnt))

1633       Call copy_ibvars
' <VB WATCH>
1634       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1635       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "ibuf", ibuf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Function ilwrtia(ByVal ud As Integer, ByRef ibuf() As Integer, ByVal cnt As Long) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1636       On Error GoTo vbwErrHandler
1637       Const VBWPROCNAME = "VBIB32.ilwrtia"
1638       If vbwProtector.vbwTraceProc Then
1639           Dim vbwProtectorParameterString As String
1640           If vbwProtector.vbwTraceParameters Then
1641               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
1642               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ibuf", ibuf) & ", "
1643               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cnt", cnt) & ") "
1644           End If
1645           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1646       End If
' </VB WATCH>
1647       If (GPIBglobalsRegistered = 0) Then
1648         Call RegisterGPIBGlobals
1649       End If

       ' Call the 32-bit DLL.
1650       ilwrtia = ConvertLongToInt(ibwrt32(ud, ibuf(0), cnt))

       ' When Visual Basic remapping buffer problem solved, use this:
       '    ilwrtia = ConvertLongToInt(ibwrta32(ud, ibuf(0), cnt))

1651       Call copy_ibvars
' <VB WATCH>
1652       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1653       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "ibuf", ibuf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Sub PassControl(ByVal boardID As Integer, ByVal addr As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1654       On Error GoTo vbwErrHandler
1655       Const VBWPROCNAME = "VBIB32.PassControl"
1656       If vbwProtector.vbwTraceProc Then
1657           Dim vbwProtectorParameterString As String
1658           If vbwProtector.vbwTraceParameters Then
1659               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
1660               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addr", addr) & ") "
1661           End If
1662           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1663       End If
' </VB WATCH>
1664       If (GPIBglobalsRegistered = 0) Then
1665         Call RegisterGPIBGlobals
1666       End If

       ' Call the 32-bit DLL.
1667       Call PassControl32(boardID, addr)

1668       Call copy_ibvars
' <VB WATCH>
1669       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1670       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addr", addr
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub Ppoll(ByVal boardID As Integer, result As Integer)
' <VB WATCH>
1671       On Error GoTo vbwErrHandler
1672       Const VBWPROCNAME = "VBIB32.Ppoll"
1673       If vbwProtector.vbwTraceProc Then
1674           Dim vbwProtectorParameterString As String
1675           If vbwProtector.vbwTraceParameters Then
1676               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
1677               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("result", result) & ") "
1678           End If
1679           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1680       End If
' </VB WATCH>
1681       Dim tmpresult As Long

       ' Check to see if GPIB Global variables are registered
1682       If (GPIBglobalsRegistered = 0) Then
1683         Call RegisterGPIBGlobals
1684       End If

       ' Call the 32-bit DLL.
1685       Call PPoll32(boardID, tmpresult)

1686       result = ConvertLongToInt(tmpresult)

1687       Call copy_ibvars
' <VB WATCH>
1688       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1689       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "result", result
            vbwReportVariable "tmpresult", tmpresult
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub PpollConfig(ByVal boardID As Integer, ByVal addr As Integer, ByVal lline As Integer, ByVal sense As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1690       On Error GoTo vbwErrHandler
1691       Const VBWPROCNAME = "VBIB32.PpollConfig"
1692       If vbwProtector.vbwTraceProc Then
1693           Dim vbwProtectorParameterString As String
1694           If vbwProtector.vbwTraceParameters Then
1695               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
1696               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addr", addr) & ", "
1697               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("lline", lline) & ", "
1698               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sense", sense) & ") "
1699           End If
1700           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1701       End If
' </VB WATCH>
1702       If (GPIBglobalsRegistered = 0) Then
1703         Call RegisterGPIBGlobals
1704       End If

       ' Call the 32-bit DLL.
1705       Call PPollConfig32(boardID, addr, lline, sense)

1706       Call copy_ibvars
' <VB WATCH>
1707       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1708       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addr", addr
            vbwReportVariable "lline", lline
            vbwReportVariable "sense", sense
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub PpollUnconfig(ByVal boardID As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1709       On Error GoTo vbwErrHandler
1710       Const VBWPROCNAME = "VBIB32.PpollUnconfig"
1711       If vbwProtector.vbwTraceProc Then
1712           Dim vbwProtectorParameterString As String
1713           If vbwProtector.vbwTraceParameters Then
1714               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
1715               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ") "
1716           End If
1717           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1718       End If
' </VB WATCH>
1719       If (GPIBglobalsRegistered = 0) Then
1720         Call RegisterGPIBGlobals
1721       End If

       ' Call the 32-bit DLL.
1722       Call PPollUnconfig32(boardID, addrs(0))

1723       Call copy_ibvars
' <VB WATCH>
1724       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1725       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addrs", addrs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub RcvRespMsg(ByVal boardID As Integer, buf As String, ByVal term As Integer)
' <VB WATCH>
1726       On Error GoTo vbwErrHandler
1727       Const VBWPROCNAME = "VBIB32.RcvRespMsg"
1728       If vbwProtector.vbwTraceProc Then
1729           Dim vbwProtectorParameterString As String
1730           If vbwProtector.vbwTraceParameters Then
1731               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
1732               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
1733               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("term", term) & ") "
1734           End If
1735           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1736       End If
' </VB WATCH>
1737       Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
1738       If (GPIBglobalsRegistered = 0) Then
1739         Call RegisterGPIBGlobals
1740       End If

1741       cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
1742       Call RcvRespMsg32(boardID, ByVal buf, cnt, term)

1743       Call copy_ibvars
' <VB WATCH>
1744       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1745       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "buf", buf
            vbwReportVariable "term", term
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ReadStatusByte(ByVal boardID As Integer, ByVal addr As Integer, result As Integer)
' <VB WATCH>
1746       On Error GoTo vbwErrHandler
1747       Const VBWPROCNAME = "VBIB32.ReadStatusByte"
1748       If vbwProtector.vbwTraceProc Then
1749           Dim vbwProtectorParameterString As String
1750           If vbwProtector.vbwTraceParameters Then
1751               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
1752               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addr", addr) & ", "
1753               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("result", result) & ") "
1754           End If
1755           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1756       End If
' </VB WATCH>
1757       Dim tmpresult As Long

       ' Check to see if GPIB Global variables are registered
1758       If (GPIBglobalsRegistered = 0) Then
1759         Call RegisterGPIBGlobals
1760       End If

       ' Call the 32-bit DLL.
1761       Call ReadStatusByte32(boardID, addr, tmpresult)

1762       result = ConvertLongToInt(tmpresult)

1763       Call copy_ibvars
' <VB WATCH>
1764       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1765       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addr", addr
            vbwReportVariable "result", result
            vbwReportVariable "tmpresult", tmpresult
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub Receive(ByVal boardID As Integer, ByVal addr As Integer, buf As String, ByVal term As Integer)
' <VB WATCH>
1766       On Error GoTo vbwErrHandler
1767       Const VBWPROCNAME = "VBIB32.Receive"
1768       If vbwProtector.vbwTraceProc Then
1769           Dim vbwProtectorParameterString As String
1770           If vbwProtector.vbwTraceParameters Then
1771               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
1772               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addr", addr) & ", "
1773               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
1774               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("term", term) & ") "
1775           End If
1776           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1777       End If
' </VB WATCH>
1778       Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
1779       If (GPIBglobalsRegistered = 0) Then
1780         Call RegisterGPIBGlobals
1781       End If

1782       cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
1783       Call Receive32(boardID, addr, ByVal buf, cnt, term)

1784       Call copy_ibvars
' <VB WATCH>
1785       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1786       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addr", addr
            vbwReportVariable "buf", buf
            vbwReportVariable "term", term
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ReceiveSetup(ByVal boardID As Integer, ByVal addr As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1787       On Error GoTo vbwErrHandler
1788       Const VBWPROCNAME = "VBIB32.ReceiveSetup"
1789       If vbwProtector.vbwTraceProc Then
1790           Dim vbwProtectorParameterString As String
1791           If vbwProtector.vbwTraceParameters Then
1792               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
1793               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addr", addr) & ") "
1794           End If
1795           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1796       End If
' </VB WATCH>
1797       If (GPIBglobalsRegistered = 0) Then
1798         Call RegisterGPIBGlobals
1799       End If

       ' Call the 32-bit DLL.
1800       Call ReceiveSetup32(boardID, addr)

1801       Call copy_ibvars
' <VB WATCH>
1802       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1803       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addr", addr
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub ResetSys(ByVal boardID As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1804       On Error GoTo vbwErrHandler
1805       Const VBWPROCNAME = "VBIB32.ResetSys"
1806       If vbwProtector.vbwTraceProc Then
1807           Dim vbwProtectorParameterString As String
1808           If vbwProtector.vbwTraceParameters Then
1809               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
1810               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ") "
1811           End If
1812           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1813       End If
' </VB WATCH>
1814       If (GPIBglobalsRegistered = 0) Then
1815         Call RegisterGPIBGlobals
1816       End If

       ' Call the 32-bit DLL.
1817       Call ResetSys32(boardID, addrs(0))

1818       Call copy_ibvars
' <VB WATCH>
1819       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1820       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addrs", addrs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub Send(ByVal boardID As Integer, ByVal addr As Integer, ByVal buf As String, ByVal term As Integer)
' <VB WATCH>
1821       On Error GoTo vbwErrHandler
1822       Const VBWPROCNAME = "VBIB32.Send"
1823       If vbwProtector.vbwTraceProc Then
1824           Dim vbwProtectorParameterString As String
1825           If vbwProtector.vbwTraceParameters Then
1826               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
1827               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addr", addr) & ", "
1828               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
1829               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("term", term) & ") "
1830           End If
1831           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1832       End If
' </VB WATCH>
1833       Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
1834       If (GPIBglobalsRegistered = 0) Then
1835         Call RegisterGPIBGlobals
1836       End If

1837       cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
1838       Call Send32(boardID, addr, ByVal buf, cnt, term)

1839       Call copy_ibvars
' <VB WATCH>
1840       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1841       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addr", addr
            vbwReportVariable "buf", buf
            vbwReportVariable "term", term
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub SendCmds(ByVal boardID As Integer, ByVal cmdbuf As String)
' <VB WATCH>
1842       On Error GoTo vbwErrHandler
1843       Const VBWPROCNAME = "VBIB32.SendCmds"
1844       If vbwProtector.vbwTraceProc Then
1845           Dim vbwProtectorParameterString As String
1846           If vbwProtector.vbwTraceParameters Then
1847               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
1848               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("cmdbuf", cmdbuf) & ") "
1849           End If
1850           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1851       End If
' </VB WATCH>
1852       Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
1853       If (GPIBglobalsRegistered = 0) Then
1854         Call RegisterGPIBGlobals
1855       End If

1856       cnt = CLng(Len(cmdbuf))

       ' Call the 32-bit DLL.
1857       Call SendCmds32(boardID, ByVal cmdbuf, cnt)

1858       Call copy_ibvars
' <VB WATCH>
1859       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1860       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "cmdbuf", cmdbuf
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub SendDataBytes(ByVal boardID As Integer, ByVal buf As String, ByVal term As Integer)
' <VB WATCH>
1861       On Error GoTo vbwErrHandler
1862       Const VBWPROCNAME = "VBIB32.SendDataBytes"
1863       If vbwProtector.vbwTraceProc Then
1864           Dim vbwProtectorParameterString As String
1865           If vbwProtector.vbwTraceParameters Then
1866               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
1867               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
1868               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("term", term) & ") "
1869           End If
1870           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1871       End If
' </VB WATCH>
1872       Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
1873       If (GPIBglobalsRegistered = 0) Then
1874         Call RegisterGPIBGlobals
1875       End If

1876       cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
1877       Call SendDataBytes32(boardID, ByVal buf, cnt, term)

1878       Call copy_ibvars
' <VB WATCH>
1879       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1880       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "buf", buf
            vbwReportVariable "term", term
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub SendIFC(ByVal boardID As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1881       On Error GoTo vbwErrHandler
1882       Const VBWPROCNAME = "VBIB32.SendIFC"
1883       If vbwProtector.vbwTraceProc Then
1884           Dim vbwProtectorParameterString As String
1885           If vbwProtector.vbwTraceParameters Then
1886               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ") "
1887           End If
1888           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1889       End If
' </VB WATCH>
1890       If (GPIBglobalsRegistered = 0) Then
1891         Call RegisterGPIBGlobals
1892       End If

       ' Call the 32-bit DLL.
1893       Call SendIFC32(boardID)

1894       Call copy_ibvars
' <VB WATCH>
1895       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1896       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub SendList(ByVal boardID As Integer, addr() As Integer, ByVal buf As String, ByVal term As Integer)
' <VB WATCH>
1897       On Error GoTo vbwErrHandler
1898       Const VBWPROCNAME = "VBIB32.SendList"
1899       If vbwProtector.vbwTraceProc Then
1900           Dim vbwProtectorParameterString As String
1901           If vbwProtector.vbwTraceParameters Then
1902               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
1903               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addr", addr) & ", "
1904               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ", "
1905               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("term", term) & ") "
1906           End If
1907           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1908       End If
' </VB WATCH>
1909       Dim cnt As Long

       ' Check to see if GPIB Global variables are registered
1910       If (GPIBglobalsRegistered = 0) Then
1911         Call RegisterGPIBGlobals
1912       End If

1913       cnt = CLng(Len(buf))

       ' Call the 32-bit DLL.
1914       Call SendList32(boardID, addr(0), ByVal buf, cnt, term)

1915       Call copy_ibvars
' <VB WATCH>
1916       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1917       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addr", addr
            vbwReportVariable "buf", buf
            vbwReportVariable "term", term
            vbwReportVariable "cnt", cnt
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub SendLLO(ByVal boardID As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1918       On Error GoTo vbwErrHandler
1919       Const VBWPROCNAME = "VBIB32.SendLLO"
1920       If vbwProtector.vbwTraceProc Then
1921           Dim vbwProtectorParameterString As String
1922           If vbwProtector.vbwTraceParameters Then
1923               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ") "
1924           End If
1925           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1926       End If
' </VB WATCH>
1927       If (GPIBglobalsRegistered = 0) Then
1928         Call RegisterGPIBGlobals
1929       End If

       ' Call the 32-bit DLL.
1930       Call SendLLO32(boardID)

1931       Call copy_ibvars
' <VB WATCH>
1932       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1933       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub SendSetup(ByVal boardID As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1934       On Error GoTo vbwErrHandler
1935       Const VBWPROCNAME = "VBIB32.SendSetup"
1936       If vbwProtector.vbwTraceProc Then
1937           Dim vbwProtectorParameterString As String
1938           If vbwProtector.vbwTraceParameters Then
1939               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
1940               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ") "
1941           End If
1942           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1943       End If
' </VB WATCH>
1944       If (GPIBglobalsRegistered = 0) Then
1945         Call RegisterGPIBGlobals
1946       End If

       ' Call the 32-bit DLL.
1947       Call SendSetup32(boardID, addrs(0))

1948       Call copy_ibvars
' <VB WATCH>
1949       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1950       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addrs", addrs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub SetRWLS(ByVal boardID As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1951       On Error GoTo vbwErrHandler
1952       Const VBWPROCNAME = "VBIB32.SetRWLS"
1953       If vbwProtector.vbwTraceProc Then
1954           Dim vbwProtectorParameterString As String
1955           If vbwProtector.vbwTraceParameters Then
1956               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
1957               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ") "
1958           End If
1959           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1960       End If
' </VB WATCH>
1961       If (GPIBglobalsRegistered = 0) Then
1962         Call RegisterGPIBGlobals
1963       End If

       ' Call the 32-bit DLL.
1964       Call SetRWLS32(boardID, addrs(0))

1965       Call copy_ibvars
' <VB WATCH>
1966       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1967       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addrs", addrs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub TestSRQ(ByVal boardID As Integer, result As Integer)
' <VB WATCH>
1968       On Error GoTo vbwErrHandler
1969       Const VBWPROCNAME = "VBIB32.TestSRQ"
1970       If vbwProtector.vbwTraceProc Then
1971           Dim vbwProtectorParameterString As String
1972           If vbwProtector.vbwTraceParameters Then
1973               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
1974               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("result", result) & ") "
1975           End If
1976           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1977       End If
' </VB WATCH>
1978       Call ibwait(boardID, 0)

1979       If ibsta And &H1000 Then
1980           result = 1
1981       Else
1982           result = 0
1983       End If

' <VB WATCH>
1984       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1985       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "result", result
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub TestSys(ByVal boardID As Integer, addrs() As Integer, results() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
1986       On Error GoTo vbwErrHandler
1987       Const VBWPROCNAME = "VBIB32.TestSys"
1988       If vbwProtector.vbwTraceProc Then
1989           Dim vbwProtectorParameterString As String
1990           If vbwProtector.vbwTraceParameters Then
1991               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
1992               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ", "
1993               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("results", results) & ") "
1994           End If
1995           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1996       End If
' </VB WATCH>
1997       If (GPIBglobalsRegistered = 0) Then
1998         Call RegisterGPIBGlobals
1999       End If

       ' Call the 32-bit DLL.
2000       Call TestSys32(boardID, addrs(0), results(0))

2001       Call copy_ibvars
' <VB WATCH>
2002       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2003       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addrs", addrs
            vbwReportVariable "results", results
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub Trigger(ByVal boardID As Integer, ByVal addr As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
2004       On Error GoTo vbwErrHandler
2005       Const VBWPROCNAME = "VBIB32.Trigger"
2006       If vbwProtector.vbwTraceProc Then
2007           Dim vbwProtectorParameterString As String
2008           If vbwProtector.vbwTraceParameters Then
2009               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
2010               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addr", addr) & ") "
2011           End If
2012           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2013       End If
' </VB WATCH>
2014       If (GPIBglobalsRegistered = 0) Then
2015         Call RegisterGPIBGlobals
2016       End If

       ' Call the 32-bit DLL.
2017       Call Trigger32(boardID, addr)

2018       Call copy_ibvars
' <VB WATCH>
2019       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2020       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addr", addr
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub TriggerList(ByVal boardID As Integer, addrs() As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
2021       On Error GoTo vbwErrHandler
2022       Const VBWPROCNAME = "VBIB32.TriggerList"
2023       If vbwProtector.vbwTraceProc Then
2024           Dim vbwProtectorParameterString As String
2025           If vbwProtector.vbwTraceParameters Then
2026               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
2027               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("addrs", addrs) & ") "
2028           End If
2029           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2030       End If
' </VB WATCH>
2031       If (GPIBglobalsRegistered = 0) Then
2032         Call RegisterGPIBGlobals
2033       End If

       ' Call the 32-bit DLL.
2034       Call TriggerList32(boardID, addrs(0))

2035       Call copy_ibvars
' <VB WATCH>
2036       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2037       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "addrs", addrs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub WaitSRQ(ByVal boardID As Integer, result As Integer)
' <VB WATCH>
2038       On Error GoTo vbwErrHandler
2039       Const VBWPROCNAME = "VBIB32.WaitSRQ"
2040       If vbwProtector.vbwTraceProc Then
2041           Dim vbwProtectorParameterString As String
2042           If vbwProtector.vbwTraceParameters Then
2043               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("boardID", boardID) & ", "
2044               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("result", result) & ") "
2045           End If
2046           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2047       End If
' </VB WATCH>
2048       Call ibwait(boardID, &H5000)

2049       If ibsta And &H1000 Then
2050           result = 1
2051       Else
2052           result = 0
2053       End If
' <VB WATCH>
2054       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2055       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "boardID", boardID
            vbwReportVariable "result", result
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub


Private Function ConvertLongToInt(LongNumb As Long) As Integer
' <VB WATCH>
2056       On Error GoTo vbwErrHandler
2057       Const VBWPROCNAME = "VBIB32.ConvertLongToInt"
2058       If vbwProtector.vbwTraceProc Then
2059           Dim vbwProtectorParameterString As String
2060           If vbwProtector.vbwTraceParameters Then
2061               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("LongNumb", LongNumb) & ") "
2062           End If
2063           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2064       End If
' </VB WATCH>

2065     If (LongNumb And &H8000&) = 0 Then
2066         ConvertLongToInt = LongNumb And &HFFFF&
2067     Else
2068       ConvertLongToInt = &H8000 Or (LongNumb And &H7FFF&)
2069     End If

' <VB WATCH>
2070       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2071       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "LongNumb", LongNumb
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Public Sub RegisterGPIBGlobals()
' <VB WATCH>
2072       On Error GoTo vbwErrHandler
2073       Const VBWPROCNAME = "VBIB32.RegisterGPIBGlobals"
2074       If vbwProtector.vbwTraceProc Then
2075           Dim vbwProtectorParameterString As String
2076           If vbwProtector.vbwTraceParameters Then
2077               vbwProtectorParameterString = "()"
2078           End If
2079           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2080       End If
' </VB WATCH>
2081       Dim rc As Long

2082       rc = RegisterGpibGlobalsForThread(Longibsta, Longiberr, Longibcnt, ibcntl)
2083       If (rc = 0) Then
2084         GPIBglobalsRegistered = 1
2085       ElseIf (rc = 1) Then
2086         rc = UnregisterGpibGlobalsForThread
2087         rc = RegisterGpibGlobalsForThread(Longibsta, Longiberr, Longibcnt, ibcntl)
2088         GPIBglobalsRegistered = 1
2089       ElseIf (rc = 2) Then
2090         rc = UnregisterGpibGlobalsForThread
2091         ibsta = &H8000
2092         iberr = EDVR
2093         ibcntl = &HDEAD37F0
2094       ElseIf (rc = 3) Then
2095         rc = UnregisterGpibGlobalsForThread
2096         ibsta = &H8000
2097         iberr = EDVR
2098         ibcntl = &HDEAD37F0
2099       Else
2100         ibsta = &H8000
2101         iberr = EDVR
2102         ibcntl = &HDEAD37F0
2103       End If
' <VB WATCH>
2104       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2105       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "rc", rc
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Public Sub UnregisterGPIBGlobals()
' <VB WATCH>
2106       On Error GoTo vbwErrHandler
2107       Const VBWPROCNAME = "VBIB32.UnregisterGPIBGlobals"
2108       If vbwProtector.vbwTraceProc Then
2109           Dim vbwProtectorParameterString As String
2110           If vbwProtector.vbwTraceParameters Then
2111               vbwProtectorParameterString = "()"
2112           End If
2113           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2114       End If
' </VB WATCH>
2115       Dim rc As Long

2116       rc = UnregisterGpibGlobalsForThread
2117       GPIBglobalsRegistered = 0

' <VB WATCH>
2118       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2119       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "rc", rc
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub



Public Function ThreadIbsta() As Integer
       ' Call the 32-bit DLL.
' <VB WATCH>
2120       On Error GoTo vbwErrHandler
2121       Const VBWPROCNAME = "VBIB32.ThreadIbsta"
2122       If vbwProtector.vbwTraceProc Then
2123           Dim vbwProtectorParameterString As String
2124           If vbwProtector.vbwTraceParameters Then
2125               vbwProtectorParameterString = "()"
2126           End If
2127           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2128       End If
' </VB WATCH>
2129       ThreadIbsta = ConvertLongToInt(ThreadIbsta32())
' <VB WATCH>
2130       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2131       Exit Function
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

Public Function ThreadIberr() As Integer
       ' Call the 32-bit DLL.
' <VB WATCH>
2132       On Error GoTo vbwErrHandler
2133       Const VBWPROCNAME = "VBIB32.ThreadIberr"
2134       If vbwProtector.vbwTraceProc Then
2135           Dim vbwProtectorParameterString As String
2136           If vbwProtector.vbwTraceParameters Then
2137               vbwProtectorParameterString = "()"
2138           End If
2139           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2140       End If
' </VB WATCH>
2141       ThreadIberr = ConvertLongToInt(ThreadIberr32())
' <VB WATCH>
2142       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2143       Exit Function
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

Public Function ThreadIbcnt() As Integer
       ' Call the 32-bit DLL.
' <VB WATCH>
2144       On Error GoTo vbwErrHandler
2145       Const VBWPROCNAME = "VBIB32.ThreadIbcnt"
2146       If vbwProtector.vbwTraceProc Then
2147           Dim vbwProtectorParameterString As String
2148           If vbwProtector.vbwTraceParameters Then
2149               vbwProtectorParameterString = "()"
2150           End If
2151           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2152       End If
' </VB WATCH>
2153       ThreadIbcnt = ConvertLongToInt(ThreadIbcnt32())
' <VB WATCH>
2154       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2155       Exit Function
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

Public Function ThreadIbcntl() As Long
       ' Call the 32-bit DLL.
' <VB WATCH>
2156       On Error GoTo vbwErrHandler
2157       Const VBWPROCNAME = "VBIB32.ThreadIbcntl"
2158       If vbwProtector.vbwTraceProc Then
2159           Dim vbwProtectorParameterString As String
2160           If vbwProtector.vbwTraceParameters Then
2161               vbwProtectorParameterString = "()"
2162           End If
2163           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2164       End If
' </VB WATCH>
2165       ThreadIbcntl = ThreadIbcntl32()
' <VB WATCH>
2166       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2167       Exit Function
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

Public Function illock(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
2168       On Error GoTo vbwErrHandler
2169       Const VBWPROCNAME = "VBIB32.illock"
2170       If vbwProtector.vbwTraceProc Then
2171           Dim vbwProtectorParameterString As String
2172           If vbwProtector.vbwTraceParameters Then
2173               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
2174           End If
2175           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2176       End If
' </VB WATCH>
2177       If (GPIBglobalsRegistered = 0) Then
2178         Call RegisterGPIBGlobals
2179       End If

       ' Call the 32-bit DLL.
2180       illock = ConvertLongToInt(iblock32(ud))

2181       Call copy_ibvars
' <VB WATCH>
2182       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2183       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Public Function ilunlock(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
2184       On Error GoTo vbwErrHandler
2185       Const VBWPROCNAME = "VBIB32.ilunlock"
2186       If vbwProtector.vbwTraceProc Then
2187           Dim vbwProtectorParameterString As String
2188           If vbwProtector.vbwTraceParameters Then
2189               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
2190           End If
2191           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2192       End If
' </VB WATCH>
2193       If (GPIBglobalsRegistered = 0) Then
2194         Call RegisterGPIBGlobals
2195       End If

       ' Call the 32-bit DLL.
2196       ilunlock = ConvertLongToInt(ibunlock32(ud))

2197       Call copy_ibvars
' <VB WATCH>
2198       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2199       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Public Sub iblock(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
2200       On Error GoTo vbwErrHandler
2201       Const VBWPROCNAME = "VBIB32.iblock"
2202       If vbwProtector.vbwTraceProc Then
2203           Dim vbwProtectorParameterString As String
2204           If vbwProtector.vbwTraceParameters Then
2205               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
2206           End If
2207           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2208       End If
' </VB WATCH>
2209       If (GPIBglobalsRegistered = 0) Then
2210         Call RegisterGPIBGlobals
2211       End If

       ' Call the 32-bit DLL.
2212       Call iblock32(ud)

2213       Call copy_ibvars
' <VB WATCH>
2214       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2215       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Public Sub ibunlock(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
2216       On Error GoTo vbwErrHandler
2217       Const VBWPROCNAME = "VBIB32.ibunlock"
2218       If vbwProtector.vbwTraceProc Then
2219           Dim vbwProtectorParameterString As String
2220           If vbwProtector.vbwTraceParameters Then
2221               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
2222           End If
2223           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2224       End If
' </VB WATCH>
2225       If (GPIBglobalsRegistered = 0) Then
2226         Call RegisterGPIBGlobals
2227       End If

       ' Call the 32-bit DLL.
2228       Call ibunlock32(ud)

2229       Call copy_ibvars
' <VB WATCH>
2230       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2231       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Public Function illockx(ByVal ud As Integer, ByVal LockWaitTime As Integer, ByVal buf As String) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
2232       On Error GoTo vbwErrHandler
2233       Const VBWPROCNAME = "VBIB32.illockx"
2234       If vbwProtector.vbwTraceProc Then
2235           Dim vbwProtectorParameterString As String
2236           If vbwProtector.vbwTraceParameters Then
2237               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
2238               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("LockWaitTime", LockWaitTime) & ", "
2239               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ") "
2240           End If
2241           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2242       End If
' </VB WATCH>
2243       If (GPIBglobalsRegistered = 0) Then
2244         Call RegisterGPIBGlobals
2245       End If

       ' Call the 32-bit DLL.
2246       illockx = ConvertLongToInt(iblockx32(ud, LockWaitTime, buf))

2247       Call copy_ibvars
' <VB WATCH>
2248       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2249       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "LockWaitTime", LockWaitTime
            vbwReportVariable "buf", buf
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Public Function ilunlockx(ByVal ud As Integer) As Integer
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
2250       On Error GoTo vbwErrHandler
2251       Const VBWPROCNAME = "VBIB32.ilunlockx"
2252       If vbwProtector.vbwTraceProc Then
2253           Dim vbwProtectorParameterString As String
2254           If vbwProtector.vbwTraceParameters Then
2255               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
2256           End If
2257           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2258       End If
' </VB WATCH>
2259       If (GPIBglobalsRegistered = 0) Then
2260         Call RegisterGPIBGlobals
2261       End If

       ' Call the 32-bit DLL.
2262       ilunlockx = ConvertLongToInt(ibunlockx32(ud))

2263       Call copy_ibvars
' <VB WATCH>
2264       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2265       Exit Function
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Public Sub iblockx(ByVal ud As Integer, ByVal LockWaitTime As Integer, ByVal buf As String)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
2266       On Error GoTo vbwErrHandler
2267       Const VBWPROCNAME = "VBIB32.iblockx"
2268       If vbwProtector.vbwTraceProc Then
2269           Dim vbwProtectorParameterString As String
2270           If vbwProtector.vbwTraceParameters Then
2271               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ", "
2272               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("LockWaitTime", LockWaitTime) & ", "
2273               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("buf", buf) & ") "
2274           End If
2275           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2276       End If
' </VB WATCH>
2277       If (GPIBglobalsRegistered = 0) Then
2278         Call RegisterGPIBGlobals
2279       End If

       ' Call the 32-bit DLL.
2280       Call iblockx32(ud, LockWaitTime, buf)

2281       Call copy_ibvars
' <VB WATCH>
2282       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2283       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
            vbwReportVariable "LockWaitTime", LockWaitTime
            vbwReportVariable "buf", buf
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Public Sub ibunlockx(ByVal ud As Integer)
       ' Check to see if GPIB Global variables are registered
' <VB WATCH>
2284       On Error GoTo vbwErrHandler
2285       Const VBWPROCNAME = "VBIB32.ibunlockx"
2286       If vbwProtector.vbwTraceProc Then
2287           Dim vbwProtectorParameterString As String
2288           If vbwProtector.vbwTraceParameters Then
2289               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ud", ud) & ") "
2290           End If
2291           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2292       End If
' </VB WATCH>
2293       If (GPIBglobalsRegistered = 0) Then
2294         Call RegisterGPIBGlobals
2295       End If

       ' Call the 32-bit DLL.
2296       Call ibunlockx32(ud)

2297       Call copy_ibvars
' <VB WATCH>
2298       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2299       Exit Sub
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
        Case vbwDoDumpVariable
            vbwOpenDumpFile
            vbwReportToFile VBW_LOCAL_STRING
            vbwReportVariable "ud", ud
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
