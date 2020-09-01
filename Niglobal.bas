Attribute VB_Name = "NIGLOBAL"
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' 32-bit Visual Basic Language Interface
' Version 1.9
' Copyright 2001 National Instruments Corporation.
' All Rights Reserved.
'
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'   This module contains the variable  declarations,
'   constant definitions, and type information that
'   is recognized by the entire application.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
Option Explicit

Global ibsta As Integer
Global iberr As Integer
Global ibcnt As Integer
Global ibcntl As Long

' Needed to register for GPIB Global Thread.
Global Longibsta As Long
Global Longiberr As Long
Global Longibcnt As Long
Global GPIBglobalsRegistered As Integer

Global buf As String

Global bytebuf() As Byte

Global Const UNL = &H3F     ' GPIB unlisten command
Global Const UNT = &H5F     ' GPIB untalk command
Global Const GTL = &H1      ' GPIB go to local
Global Const SDC = &H4      ' GPIB selected device clear
Global Const PPC = &H5      ' GPIB parallel poll configure
Global Const GGET = &H8     ' GPIB group execute trigger
Global Const TCT = &H9      ' GPIB take control
Global Const LLO = &H11     ' GPIB local lock out
Global Const DCL = &H14     ' GPIB device clear
Global Const PPU = &H15     ' GPIB parallel poll unconfigure
Global Const SPE = &H18     ' GPIB serial poll enable
Global Const SPD = &H19     ' GPIB serial poll disable
Global Const PPE = &H60     ' GPIB parallel poll enable
Global Const PPD = &H70     ' GPIB parallel poll disable

  ' GPIB status bit vector :
'       status variable ibsta and wait mask

Global Const EERR = &H8000      ' Error detected
Global Const TIMO = &H4000      ' Timeout
Global Const EEND = &H2000      ' EOI or EOS detected
Global Const SRQI = &H1000      ' SRQ detected by CIC
Global Const RQS = &H800        ' Device requesting service
Global Const CMPL = &H100       ' I/O completed
Global Const LOK = &H80         ' Local lockout state
Global Const RREM = &H40        ' Remote state
Global Const CIC = &H20         ' Controller-in-Charge
Global Const AATN = &H10        ' Attention asserted
Global Const TACS = &H8         ' Talker active
Global Const LACS = &H4         ' Listener active
Global Const DTAS = &H2         ' Device trigger state
Global Const DCAS = &H1         ' Device clear state

' Error messages returned in global variable iberr

Global Const EDVR = 0      ' System error
Global Const ECIC = 1      ' Function requires GPIB board to be CIC
Global Const ENOL = 2      ' Write function detected no listeners
Global Const EADR = 3      ' Interface board not addressed correctly
Global Const EARG = 4      ' Invalid argument to function call
Global Const ESAC = 5      ' Function requires GPIB board to be SAC
Global Const EABO = 6      ' I/O operation aborted
Global Const ENEB = 7      ' Non-existent interface board
Global Const EDMA = 8      ' DMA Error
Global Const EOIP = 10     ' I/O operation started before previous
                           ' operation completed
Global Const ECAP = 11     ' No capability for intended operation
Global Const EFSO = 12     ' File system operation error
Global Const EBUS = 14     ' Command error during device call
Global Const ESTB = 15     ' Serial poll status byte lost
Global Const ESRQ = 16     ' SRQ remains asserted
Global Const ETAB = 20     ' The return buffer is full
Global Const ELCK = 21     ' Address or board is locked
Global Const EARM = 22     ' ibnotify's asynchronous event
                           ' notification failed to rearm
Global Const EHDL = 23     ' Invalid GPIB handle type for function
Global Const WCFG = 24     ' Configuration warning
Global Const ECFG = 24     ' Configuration warning
Global Const EWIP = 26     ' Wait already in progress on input ud
Global Const ERST = 27     ' The event notification was cancelled
                           ' due to a reset of the interface
Global Const EPWR = 28     ' The interface lost power


' EOS mode bits

Global Const BIN = &H1000   ' Eight bit compare
Global Const XEOS = &H800   ' Send EOI with EOS byte
Global Const REOS = &H400   ' Terminate read on EOS

' Timeout values and meanings

Global Const TNONE = 0      ' Infinite timeout (disabled)
Global Const T10us = 1      ' Timeout of 10 us (ideal)
Global Const T30us = 2      ' Timeout of 30 us (ideal)
Global Const T100us = 3     ' Timeout of 100 us (ideal)
Global Const T300us = 4     ' Timeout of 300 us (ideal)
Global Const T1ms = 5       ' Timeout of 1 ms (ideal)
Global Const T3ms = 6       ' Timeout of 3 ms (ideal)
Global Const T10ms = 7      ' Timeout of 10 ms (ideal)
Global Const T30ms = 8      ' Timeout of 30 ms (ideal)
Global Const T100ms = 9     ' Timeout of 100 ms (ideal)
Global Const T300ms = 10    ' Timeout of 300 ms (ideal)
Global Const T1s = 11       ' Timeout of 1 s (ideal)
Global Const T3s = 12       ' Timeout of 3 s (ideal)
Global Const T10s = 13      ' Timeout of 10 s (ideal)
Global Const T30s = 14      ' Timeout of 30 s (ideal)
Global Const T100s = 15     ' Timeout of 100 s (ideal)
Global Const T300s = 16     ' Timeout of 300 s (ideal)
Global Const T1000s = 17    ' Timeout of 1000 s (maximum)

' IBLN constants

Global Const ALL_SAD = -1
Global Const NO_SAD = 0

' The following constants are used for the second parameter of the
' ibconfig function.  They are the "option" selection codes.

Global Const IbcPAD = &H1             ' Primary Address
Global Const IbcSAD = &H2             ' Secondary Address
Global Const IbcTMO = &H3             ' Timeout Value
Global Const IbcEOT = &H4             ' Send EOI with last data byte?
Global Const IbcPPC = &H5             ' Parallel Poll Configure
Global Const IbcREADDR = &H6          ' Repeat Addressing
Global Const IbcAUTOPOLL = &H7        ' Disable Auto Serial Polling
Global Const IbcCICPROT = &H8         ' Use the CIC Protocol?
Global Const IbcIRQ = &H9             ' Use PIO for I/O
Global Const IbcSC = &HA              ' Board is System Controller.
Global Const IbcSRE = &HB             ' Assert SRE on device calls?
Global Const IbcEOSrd = &HC           ' Terminate reads on EOS.
Global Const IbcEOSwrt = &HD          ' Send EOI with EOS character.
Global Const IbcEOScmp = &HE          ' Use 7 or 8-bit EOS compare.
Global Const IbcEOSchar = &HF         ' The EOS character.
Global Const IbcPP2 = &H10            ' Use Parallel Poll Mode 2.
Global Const IbcTIMING = &H11         ' NORMAL, HIGH, or VERY_HIGH timing.
Global Const IbcDMA = &H12            ' Use DMA for I/O.
Global Const IbcReadAdjust = &H13     ' Swap bytes during an ibrd.
Global Const IbcWriteAdjust = &H14    ' Swap bytes during an ibwrt.
Global Const IbcSendLLO = &H17        ' Enable/disable the sending of LLO.
Global Const IbcSPollTime = &H18      ' Set the timeout value for serial polls.
Global Const IbcPPollTime = &H19      ' Set the parallel poll length period
Global Const IbcEndBitIsNormal = &H1A    ' Remove EOS from END bit of IBSTA.
Global Const IbcUnAddr = &H1B            ' Enable/disable device unaddressing.
Global Const IbcSignalNumber = &H1C      ' Set UNIX signal number - unsupported
Global Const IbcBlockIfLocked = &H1D     ' Enable/disable blocking for locked boards/devices
Global Const IbcHSCableLength = &H1F     ' Enable/disable high-speed handshaking.
Global Const IbcIst = &H20               ' Set the IST bit
Global Const IbcRsv = &H21               ' Set the RSV bit
Global Const IbcLON = &H22               ' Enable listen only mode.


'   Constants that can be used (in addition to the ibconfig constants)
'   when calling the IBASK function.

Global Const IbaPAD = &H1             ' Primary Address
Global Const IbaSAD = &H2             ' Secondary Address
Global Const IbaTMO = &H3             ' Timeout Value
Global Const IbaEOT = &H4             ' Send EOI with last data byte?
Global Const IbaPPC = &H5             ' Parallel Poll Configure
Global Const IbaREADDR = &H6          ' Repeat Addressing
Global Const IbaAUTOPOLL = &H7        ' Disable Auto Serial Polling
Global Const IbaCICPROT = &H8         ' Use the CIC Protocol?
Global Const IbaIRQ = &H9             ' Use PIO for I/O
Global Const IbaSC = &HA              ' Board is System Controller.
Global Const IbaSRE = &HB             ' Assert SRE on device calls?
Global Const IbaEOSrd = &HC           ' Terminate reads on EOS.
Global Const IbaEOSwrt = &HD          ' Send EOI with EOS character.
Global Const IbaEOScmp = &HE          ' Use 7 or 8-bit EOS compare.
Global Const IbaEOSchar = &HF         ' The EOS character.
Global Const IbaPP2 = &H10            ' Use Parallel Poll Mode 2.
Global Const IbaTIMING = &H11         ' NORMAL, HIGH, or VERY_HIGH timing.
Global Const IbaDMA = &H12            ' Use DMA for I/O.
Global Const IbaReadAdjust = &H13     ' Swap bytes during an ibrd.
Global Const IbaWriteAdjust = &H14    ' Swap bytes during an ibwrt.
Global Const IbaSendLLO = &H17        ' Enable/disable the sending of LLO.
Global Const IbaSPollTime = &H18      ' Set the timeout value for serial polls.
Global Const IbaPPollTime = &H19      ' Set the parallel poll length period
Global Const IbaEndBitIsNormal = &H1A ' Remove EOS from END bit of IBSTA.
Global Const IbaUnAddr = &H1B         ' Enable/disable device unaddressing.
Global Const IbaSignalNumber = &H1C   ' Set UNIX signal number - unsupported
Global Const IbaBlockIfLocked = &H1D  ' Enable/disable blocking for locked boards/devices.
Global Const IbaHSCableLength = &H1F  ' Enable/disable high-speed handshaking.
Global Const IbaIst = &H20            ' Set the IST bit
Global Const IbaRsv = &H21            ' Set the RSV bit
Global Const IbaLON = &H22            ' Enable listen only mode.
Global Const IbaBNA = &H200           ' A device's access board.
Global Const IbaBaseAddr = &H201      ' A GPIB board's base I/O address.
Global Const IbaDmaChannel = &H202    ' A GPIB board's DMA channel.
Global Const IbaIrqLevel = &H203      ' A GPIB board's IRQ level.
Global Const IbaBaud = &H204          ' Baud rate used to communicate to CT box.
Global Const IbaParity = &H205        ' Parity setting for CT box.
Global Const IbaStopBits = &H206      ' Stop bits used for communicating to CT.
Global Const IbaDataBits = &H207      ' Data bits used for communicating to CT.
Global Const IbaComPort = &H208       ' System COM port used for CT box.
Global Const IbaComIrqLevel = &H209   ' System COM port's interrupt level.
Global Const IbaComPortBase = &H20A   ' System COM port's base I/O address.
Global Const IbaSingleCycleDma = &H20B   ' Does the board use single cycle DMA?
Global Const IbaSlotNumber = &H20C       ' Board's slot number.
Global Const IbaLPTNumber = &H20D        ' Parallel port number
Global Const IbaLPTType = &H20E          ' Parallel port protocol

' These are the values used by the 488.2 Send command

Global Const NULLend = &H0            ' Do nothing at the end of a transfer
Global Const NLend = &H1              ' Send NL with EOI after a transfer
Global Const DABend = &H2             ' Send EOI with the last DAB

' This value is useds by the 488.2 Receive command

Global Const STOPend = &H100          ' Stop the read on EOI

' The following values are used by the iblines function.  The integer
  ' returned by iblines contains:
'       The lower byte will contain a "monitor" bit mask.  If a bit
'               is set (1) in this mask, then the corresponding line
'               can be monitored by the driver.  If the bit is clear (0),
'               then the line cannot be monitored.
'       The upper byte will contain the status of the bus lines.
'               Each bit corresponds to a certain bus line, and has
'               a corresponding "monitor" bit in the lower byte.

Global Const ValidEOI = &H80
Global Const ValidATN = &H40
Global Const ValidSRQ = &H20
Global Const ValidREN = &H10
Global Const ValidIFC = &H8
Global Const ValidNRFD = &H4
Global Const ValidNDAC = &H2
Global Const ValidDAV = &H1
Global Const BusEOI = &H8000
Global Const BusATN = &H4000
Global Const BusSRQ = &H2000
Global Const BusREN = &H1000
Global Const BusIFC = &H800
Global Const BusNRFD = &H400
Global Const BusNDAC = &H200
Global Const BusDAV = &H100

' This value is used to terminate an address list.  It should be
' assigned to the last entry.  (488.2)

Global Const NOADDR = &HFFFF

' This value is defined for when the GPIBnotify Callback
' has failed to rearm.

Global Const IBNOTIFY_REARM_FAILED = &HE00A003F

' These constants are for use with iblockx/ibunlockx
' functions for GPIB-ENET

Global Const TIMMEDIATE = -1
Global Const TINFINITE = -2
Global Const MAX_LOCKSHARENAME_LENGTH = 64


