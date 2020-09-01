VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmPLCData 
   Caption         =   "Polar Rundown"
   ClientHeight    =   11400
   ClientLeft      =   4110
   ClientTop       =   1650
   ClientWidth     =   16635
   LinkTopic       =   "Form1"
   ScaleHeight     =   11400
   ScaleWidth      =   16635
   WindowState     =   2  'Maximized
   Begin VB.Timer tmrNPSHr 
      Interval        =   5000
      Left            =   10920
      Top             =   0
   End
   Begin VB.CommandButton cmdSearchForPump 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Search for Pump"
      Height          =   375
      Left            =   3960
      Style           =   1  'Graphical
      TabIndex        =   263
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdCalibrate 
      Caption         =   "Calibrate Software"
      Height          =   495
      Left            =   9360
      TabIndex        =   188
      Top             =   600
      Visible         =   0   'False
      Width           =   1215
   End
   Begin VB.CommandButton cmdExit 
      BackColor       =   &H00C0C0FF&
      Caption         =   "Exit"
      Height          =   375
      Left            =   9360
      Style           =   1  'Graphical
      TabIndex        =   109
      Top             =   120
      Width           =   1215
   End
   Begin VB.ComboBox cmbTestDate 
      Height          =   315
      Left            =   6720
      TabIndex        =   43
      Top             =   120
      Width           =   2055
   End
   Begin MSAdodcLib.Adodc adohp 
      Height          =   330
      Left            =   11160
      Top             =   1200
      Visible         =   0   'False
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   582
      ConnectMode     =   0
      CursorLocation  =   3
      IsolationLevel  =   -1
      ConnectionTimeout=   15
      CommandTimeout  =   30
      CursorType      =   3
      LockType        =   3
      CommandType     =   8
      CursorOptions   =   0
      CacheSize       =   50
      MaxRecords      =   0
      BOFAction       =   0
      EOFAction       =   0
      ConnectStringType=   1
      Appearance      =   1
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      Orientation     =   0
      Enabled         =   -1
      Connect         =   "Provider=MSDASQL.1;Persist Security Info=True;Data Source=HP-3000/32;Mode=Read"
      OLEDBString     =   "Provider=MSDASQL.1;Persist Security Info=True;Data Source=HP-3000/32;Mode=Read"
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "Adodc1"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin VB.CommandButton cmdFindPump 
      BackColor       =   &H00C0FFC0&
      Caption         =   "Find Pump"
      Height          =   255
      Left            =   4080
      Style           =   1  'Graphical
      TabIndex        =   45
      Top             =   120
      Width           =   975
   End
   Begin VB.TextBox txtSN 
      Height          =   285
      Left            =   1800
      TabIndex        =   0
      Top             =   120
      Width           =   2175
   End
   Begin VB.Timer tmrStartUp 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   10440
      Top             =   0
   End
   Begin VB.TextBox txtUpdateInterval 
      Height          =   495
      Left            =   8760
      TabIndex        =   42
      Text            =   "Text1"
      Top             =   0
      Visible         =   0   'False
      Width           =   615
   End
   Begin VB.Timer tmrGetDDE 
      Interval        =   2000
      Left            =   9960
      Top             =   0
   End
   Begin TabDlg.SSTab SSTab1 
      Height          =   10200
      Left            =   1560
      TabIndex        =   9
      Top             =   1320
      Width           =   14928
      _ExtentX        =   26326
      _ExtentY        =   17992
      _Version        =   393216
      Tabs            =   4
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "Pump Data"
      TabPicture(0)   =   "Main.frx":0000
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "lbltab1(1)"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "lbltab1(2)"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "lbltab1(3)"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "lbltab1(11)"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "lbltab1(12)"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "lbltab1(13)"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).Control(6)=   "lbltab1(0)"
      Tab(0).Control(6).Enabled=   0   'False
      Tab(0).Control(7)=   "lbltab1(10)"
      Tab(0).Control(7).Enabled=   0   'False
      Tab(0).Control(8)=   "lbltab1(44)"
      Tab(0).Control(8).Enabled=   0   'False
      Tab(0).Control(9)=   "lbltab1(46)"
      Tab(0).Control(9).Enabled=   0   'False
      Tab(0).Control(10)=   "lbltab1(47)"
      Tab(0).Control(10).Enabled=   0   'False
      Tab(0).Control(11)=   "lbltab1(48)"
      Tab(0).Control(11).Enabled=   0   'False
      Tab(0).Control(12)=   "lbltab1(49)"
      Tab(0).Control(12).Enabled=   0   'False
      Tab(0).Control(13)=   "lbltab1(50)"
      Tab(0).Control(13).Enabled=   0   'False
      Tab(0).Control(14)=   "frmChempump"
      Tab(0).Control(14).Enabled=   0   'False
      Tab(0).Control(15)=   "txtBilNo"
      Tab(0).Control(15).Enabled=   0   'False
      Tab(0).Control(16)=   "txtShpNo"
      Tab(0).Control(16).Enabled=   0   'False
      Tab(0).Control(17)=   "txtModelNo"
      Tab(0).Control(17).Enabled=   0   'False
      Tab(0).Control(18)=   "txtDesignFlow"
      Tab(0).Control(18).Enabled=   0   'False
      Tab(0).Control(19)=   "txtDesignTDH"
      Tab(0).Control(19).Enabled=   0   'False
      Tab(0).Control(20)=   "txtRemarks"
      Tab(0).Control(20).Enabled=   0   'False
      Tab(0).Control(21)=   "cmdEnterPumpData"
      Tab(0).Control(21).Enabled=   0   'False
      Tab(0).Control(22)=   "txtSalesOrderNumber"
      Tab(0).Control(22).Enabled=   0   'False
      Tab(0).Control(23)=   "cmdDeletePump"
      Tab(0).Control(23).Enabled=   0   'False
      Tab(0).Control(24)=   "cmdApprovePump"
      Tab(0).Control(24).Enabled=   0   'False
      Tab(0).Control(25)=   "frmMfr"
      Tab(0).Control(25).Enabled=   0   'False
      Tab(0).Control(26)=   "cmdClearPumpData"
      Tab(0).Control(26).Enabled=   0   'False
      Tab(0).Control(27)=   "txtImpellerDia"
      Tab(0).Control(27).Enabled=   0   'False
      Tab(0).Control(28)=   "frmMiscPumpData"
      Tab(0).Control(28).Enabled=   0   'False
      Tab(0).Control(29)=   "txtLineNumber"
      Tab(0).Control(29).Enabled=   0   'False
      Tab(0).Control(30)=   "frmTEMC"
      Tab(0).Control(30).Enabled=   0   'False
      Tab(0).Control(31)=   "CommonDialog2"
      Tab(0).Control(31).Enabled=   0   'False
      Tab(0).Control(32)=   "grpSupermarket"
      Tab(0).Control(32).Enabled=   0   'False
      Tab(0).Control(33)=   "chkSuperMarketFeathered"
      Tab(0).Control(33).Enabled=   0   'False
      Tab(0).Control(34)=   "txtRVSPartNo"
      Tab(0).Control(34).Enabled=   0   'False
      Tab(0).Control(35)=   "txtXPartNum"
      Tab(0).Control(35).Enabled=   0   'False
      Tab(0).Control(36)=   "txtCustPONum"
      Tab(0).Control(36).Enabled=   0   'False
      Tab(0).ControlCount=   37
      TabCaption(1)   =   "Test Setup"
      TabPicture(1)   =   "Main.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "lbltab2(0)"
      Tab(1).Control(1)=   "lbltab2(1)"
      Tab(1).Control(2)=   "lbltab2(65)"
      Tab(1).Control(3)=   "lbltab2(88)"
      Tab(1).Control(4)=   "cmbTestSpec"
      Tab(1).Control(5)=   "cmdEnterTestSetupData"
      Tab(1).Control(6)=   "txtWho"
      Tab(1).Control(7)=   "cmdAddNewTestDate"
      Tab(1).Control(8)=   "txtTestSetupRemarks"
      Tab(1).Control(9)=   "frmInstrumentTags"
      Tab(1).Control(10)=   "frmLoopAndXducer"
      Tab(1).Control(11)=   "frmElecData"
      Tab(1).Control(12)=   "frmThrustBalMods"
      Tab(1).Control(13)=   "frmPerfMods"
      Tab(1).Control(14)=   "frmOtherFiles"
      Tab(1).Control(15)=   "CommonDialog1"
      Tab(1).Control(16)=   "cmdDeleteTestDate"
      Tab(1).Control(17)=   "cmdApproveTestDate"
      Tab(1).Control(18)=   "frmTAndI"
      Tab(1).Control(19)=   "Command1"
      Tab(1).Control(20)=   "txtRMA"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Test Data"
      TabPicture(2)   =   "Main.frx":0038
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "txtUpDn2"
      Tab(2).Control(1)=   "MSChart1"
      Tab(2).Control(2)=   "txtUpDn1"
      Tab(2).Control(3)=   "UpDown1"
      Tab(2).Control(4)=   "frmNPSH"
      Tab(2).Control(5)=   "btnRunNPSH"
      Tab(2).Control(6)=   "frmMagtrol"
      Tab(2).Control(7)=   "txtTDH"
      Tab(2).Control(8)=   "cmdReport"
      Tab(2).Control(9)=   "txtNPSHa"
      Tab(2).Control(10)=   "frmPumpData"
      Tab(2).Control(11)=   "frmPLCMisc"
      Tab(2).Control(12)=   "DataGrid2"
      Tab(2).Control(13)=   "cmdEnterTestData"
      Tab(2).Control(14)=   "fmrMiscTestData"
      Tab(2).Control(15)=   "frmThermocouples"
      Tab(2).Control(16)=   "frmAI"
      Tab(2).Control(17)=   "cmbPLCLoop"
      Tab(2).Control(18)=   "DataGrid1"
      Tab(2).Control(19)=   "UpDown2"
      Tab(2).Control(20)=   "shpGetPLCData"
      Tab(2).Control(21)=   "lbltab2(54)"
      Tab(2).Control(22)=   "lbltab2(53)"
      Tab(2).Control(23)=   "lbltab2(64)"
      Tab(2).Control(24)=   "lbltab2(63)"
      Tab(2).Control(25)=   "lbltab2(55)"
      Tab(2).ControlCount=   26
      TabCaption(3)   =   "Charts"
      TabPicture(3)   =   "Main.frx":0054
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      Begin VB.TextBox txtUpDn2 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   16.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   516
         Left            =   -60840
         TabIndex        =   425
         Text            =   "8"
         Top             =   5520
         Width           =   285
      End
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   2775
         Left            =   -67920
         OleObjectBlob   =   "Main.frx":0070
         TabIndex        =   424
         Top             =   3000
         Width           =   5655
      End
      Begin VB.TextBox txtUpDn1 
         Alignment       =   2  'Center
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   16.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   516
         Left            =   -74040
         TabIndex        =   423
         Text            =   "1"
         Top             =   8880
         Width           =   285
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   504
         Left            =   -74280
         TabIndex        =   422
         Top             =   8880
         Width           =   252
         _ExtentX        =   450
         _ExtentY        =   900
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtUpDn1"
         BuddyDispid     =   196620
         OrigLeft        =   600
         OrigTop         =   8880
         OrigRight       =   852
         OrigBottom      =   9372
         Max             =   8
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.TextBox txtCustPONum 
         Height          =   315
         Left            =   7800
         TabIndex        =   420
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtXPartNum 
         Height          =   315
         Left            =   1440
         TabIndex        =   418
         Top             =   1080
         Width           =   4932
      End
      Begin VB.TextBox txtRVSPartNo 
         Height          =   315
         Left            =   5520
         TabIndex        =   417
         Top             =   4349
         Width           =   1932
      End
      Begin VB.CheckBox chkSuperMarketFeathered 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   252
         Left            =   14400
         TabIndex        =   413
         Top             =   4320
         Width           =   252
      End
      Begin VB.Frame frmNPSH 
         Caption         =   "TDH (ft)"
         Height          =   1812
         Left            =   -63720
         TabIndex        =   400
         Top             =   480
         Visible         =   0   'False
         Width           =   3372
         Begin VB.TextBox txtNPSH 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   5
            Left            =   240
            TabIndex        =   411
            Text            =   "2"
            Top             =   1440
            Width           =   492
         End
         Begin VB.TextBox txtNPSH 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   264
            Index           =   1
            Left            =   2280
            TabIndex        =   410
            Top             =   480
            Width           =   732
         End
         Begin VB.TextBox txtNPSH 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   2
            Left            =   2280
            TabIndex        =   409
            Top             =   840
            Width           =   732
         End
         Begin VB.TextBox txtNPSH 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   288
            Index           =   4
            Left            =   2280
            TabIndex        =   407
            Top             =   1200
            Width           =   732
         End
         Begin VB.TextBox txtNPSH 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   264
            Index           =   3
            Left            =   1080
            TabIndex        =   406
            Top             =   840
            Width           =   732
         End
         Begin VB.TextBox txtNPSH 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   264
            Index           =   0
            Left            =   1080
            TabIndex        =   401
            Top             =   480
            Width           =   732
         End
         Begin VB.Label lbltab4 
            Alignment       =   2  'Center
            Caption         =   "% TDH Var"
            Height          =   252
            Index           =   5
            Left            =   120
            TabIndex        =   412
            Top             =   1200
            Width           =   732
         End
         Begin VB.Label lbltab4 
            Alignment       =   1  'Right Justify
            Caption         =   "TDH (ft)"
            Height          =   252
            Index           =   4
            Left            =   120
            TabIndex        =   408
            Top             =   840
            Width           =   852
         End
         Begin VB.Label lbltab4 
            Alignment       =   1  'Right Justify
            Caption         =   "Flow (GPM)"
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   405
            Top             =   480
            Width           =   852
         End
         Begin VB.Label lbltab4 
            Alignment       =   2  'Center
            Caption         =   "%"
            Height          =   252
            Index           =   2
            Left            =   2400
            TabIndex        =   404
            Top             =   240
            Width           =   492
         End
         Begin VB.Label lbltab4 
            Alignment       =   2  'Center
            Caption         =   "Start"
            Height          =   252
            Index           =   1
            Left            =   1080
            TabIndex        =   403
            Top             =   240
            Width           =   732
         End
         Begin VB.Label lbltab4 
            Alignment       =   1  'Right Justify
            Caption         =   "NPSHr"
            Height          =   252
            Index           =   3
            Left            =   1440
            TabIndex        =   402
            Top             =   1200
            Width           =   732
         End
      End
      Begin VB.Frame grpSupermarket 
         Caption         =   "Supermarket Pumps"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H000000FF&
         Height          =   1452
         Left            =   2880
         TabIndex        =   396
         Top             =   4800
         Visible         =   0   'False
         Width           =   9972
         Begin VB.ComboBox cmbSupermarketModel 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   384
            Left            =   3600
            TabIndex        =   398
            Top             =   600
            Width           =   3972
         End
         Begin VB.CommandButton cmdSelectSupermarket 
            BackColor       =   &H000000FF&
            Caption         =   "Cancel Supermarket Selection"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   492
            Left            =   7800
            MaskColor       =   &H000000FF&
            Style           =   1  'Graphical
            TabIndex        =   397
            Top             =   480
            UseMaskColor    =   -1  'True
            Width           =   1812
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Select Supermarket Model ==>"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   10.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ForeColor       =   &H000000FF&
            Height          =   252
            Index           =   45
            Left            =   240
            TabIndex        =   399
            Top             =   600
            Width           =   3252
         End
      End
      Begin VB.CommandButton btnRunNPSH 
         Caption         =   "Run NPSH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   732
         Left            =   -61920
         Style           =   1  'Graphical
         TabIndex        =   388
         Top             =   2400
         Width           =   1332
      End
      Begin VB.TextBox txtRMA 
         Height          =   315
         Left            =   -69960
         TabIndex        =   384
         Top             =   540
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   120
         Top             =   3480
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Frame frmTEMC 
         Caption         =   "TEMC Pump Data"
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   3735
         Left            =   240
         TabIndex        =   189
         Top             =   4680
         Visible         =   0   'False
         Width           =   14535
         Begin VB.ComboBox cmbTEMCNominalSuctionSize 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   241
            Top             =   600
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCNominalDischargeSize 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   240
            Top             =   240
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCVoltage 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   239
            Top             =   2400
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCDesignPressure 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   236
            Top             =   1320
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCCirculation 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   235
            Top             =   3120
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCModel 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   233
            Top             =   240
            Width           =   5445
         End
         Begin VB.TextBox txtTEMCFrameNumber 
            Height          =   315
            Left            =   1680
            TabIndex        =   213
            Top             =   1680
            Width           =   855
         End
         Begin VB.ComboBox cmbTEMCTRG 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   210
            Top             =   3120
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCPumpStages 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   208
            Top             =   2040
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCOtherMotor 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   206
            Top             =   2760
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCNominalImpSize 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   204
            Top             =   960
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCMaterials 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   202
            Top             =   960
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCJacketGasket 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   200
            Top             =   2400
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCInsulation 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   198
            Top             =   2040
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCImpellerType 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   196
            Top             =   1320
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCDivisionType 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   194
            Top             =   1680
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCAdditions 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   192
            Top             =   2760
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCAdapter 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   190
            Top             =   600
            Width           =   5445
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Suction Size:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   34
            Left            =   7200
            TabIndex        =   244
            Top             =   626
            Width           =   1215
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Voltage:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   32
            Left            =   480
            TabIndex        =   243
            Top             =   2433
            Width           =   1095
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Discharge Size:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   31
            Left            =   6840
            TabIndex        =   242
            Top             =   270
            Width           =   1575
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Design Pressure:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   30
            Left            =   0
            TabIndex        =   238
            Top             =   1359
            Width           =   1575
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Circulation:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   29
            Left            =   7200
            TabIndex        =   237
            Top             =   3120
            Width           =   1215
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Type:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Index           =   28
            Left            =   360
            TabIndex        =   234
            Top             =   255
            Width           =   1215
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Frame No:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   27
            Left            =   480
            TabIndex        =   212
            Top             =   1717
            Width           =   1095
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "TRG:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   26
            Left            =   480
            TabIndex        =   211
            Top             =   3150
            Width           =   1095
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "No. of Stages:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   25
            Left            =   7200
            TabIndex        =   209
            Top             =   2050
            Width           =   1215
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Other:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   24
            Left            =   480
            TabIndex        =   207
            Top             =   2791
            Width           =   1095
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Nom Imp Size:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   7200
            TabIndex        =   205
            Top             =   982
            Width           =   1215
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Materials:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   600
            TabIndex        =   203
            Top             =   1001
            Width           =   975
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Jacket/Gasket:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   7200
            TabIndex        =   201
            Top             =   2406
            Width           =   1215
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Insulation:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   240
            TabIndex        =   199
            Top             =   2075
            Width           =   1335
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Impeller Type:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   19
            Left            =   7080
            TabIndex        =   197
            Top             =   1338
            Width           =   1335
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Division Type:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   18
            Left            =   7200
            TabIndex        =   195
            Top             =   1694
            Width           =   1215
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Additions:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   17
            Left            =   7560
            TabIndex        =   193
            Top             =   2762
            Width           =   855
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            BackStyle       =   0  'Transparent
            Caption         =   "Adapter:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   720
            TabIndex        =   191
            Top             =   643
            Width           =   855
         End
      End
      Begin VB.TextBox txtLineNumber 
         Height          =   315
         Left            =   5280
         TabIndex        =   376
         Top             =   420
         Width           =   615
      End
      Begin VB.Frame frmMiscPumpData 
         Caption         =   "Pump Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   240
         TabIndex        =   350
         Top             =   2160
         Width           =   14535
         Begin VB.Frame Frame1 
            Caption         =   "NPSH Data File Directory"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   8040
            TabIndex        =   382
            Top             =   1200
            Visible         =   0   'False
            Width           =   6375
            Begin VB.TextBox txtNPSHFileLocation 
               Height          =   315
               Left            =   120
               TabIndex        =   383
               Top             =   240
               Width           =   5895
            End
         End
         Begin VB.TextBox txtLiquid 
            Height          =   315
            Left            =   2400
            TabIndex        =   370
            Top             =   1440
            Width           =   5415
         End
         Begin VB.TextBox txtJobNum 
            Height          =   315
            Left            =   12840
            TabIndex        =   369
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtSpGr 
            Height          =   315
            Left            =   5880
            TabIndex        =   365
            Top             =   990
            Width           =   1335
         End
         Begin VB.TextBox txtRatedInputPower 
            Height          =   315
            Left            =   2400
            TabIndex        =   364
            Top             =   1020
            Width           =   1335
         End
         Begin VB.TextBox txtLiquidTemperature 
            Height          =   315
            Left            =   12840
            TabIndex        =   363
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtNPSHr 
            Height          =   315
            Left            =   2400
            TabIndex        =   359
            Top             =   630
            Width           =   1335
         End
         Begin VB.TextBox txtThermalClass 
            Height          =   315
            Left            =   5880
            TabIndex        =   358
            Top             =   630
            Width           =   1335
         End
         Begin VB.TextBox txtExpClass 
            Height          =   315
            Left            =   9120
            TabIndex        =   357
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtAmps 
            Height          =   315
            Left            =   5880
            TabIndex        =   354
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtViscosity 
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.00"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Height          =   315
            Left            =   9120
            TabIndex        =   353
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtNoPhases 
            Height          =   315
            Left            =   2400
            TabIndex        =   351
            Top             =   240
            Width           =   1335
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Liquid:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   37
            Left            =   960
            TabIndex        =   372
            Top             =   1470
            Width           =   1335
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Job Number:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   43
            Left            =   11400
            TabIndex        =   371
            Top             =   630
            Width           =   1335
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Specific Gravity:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   4200
            TabIndex        =   368
            Top             =   1050
            Width           =   1575
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Rated Input Power:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   36
            Left            =   480
            TabIndex        =   367
            Top             =   1056
            Width           =   1812
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Liquid Temperature:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   40
            Left            =   10800
            TabIndex        =   366
            Top             =   330
            Width           =   1935
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            Caption         =   "NPSHr:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   35
            Left            =   960
            TabIndex        =   362
            Top             =   660
            Width           =   1335
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Thermal Class:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   39
            Left            =   4440
            TabIndex        =   361
            Top             =   660
            Width           =   1335
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            Caption         =   "EXP Class:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   42
            Left            =   7680
            TabIndex        =   360
            Top             =   660
            Width           =   1335
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Amps:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   38
            Left            =   4440
            TabIndex        =   356
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Viscosity:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   41
            Left            =   7680
            TabIndex        =   355
            Top             =   270
            Width           =   1335
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Phases:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   33
            Left            =   1440
            TabIndex        =   352
            Top             =   240
            Width           =   852
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   -67320
         TabIndex        =   349
         Top             =   6600
         Visible         =   0   'False
         Width           =   495
      End
      Begin VB.Frame frmTAndI 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Test and Inspection Report Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   5775
         Left            =   -66480
         TabIndex        =   296
         Top             =   3600
         Visible         =   0   'False
         Width           =   6135
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   14
            Left            =   3720
            TabIndex        =   348
            Top             =   5160
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   13
            Left            =   5640
            TabIndex        =   344
            Top             =   4800
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   12
            Left            =   5640
            TabIndex        =   343
            Top             =   4440
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   11
            Left            =   5640
            TabIndex        =   342
            Top             =   4080
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   10
            Left            =   5640
            TabIndex        =   341
            Top             =   3720
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   9
            Left            =   5640
            TabIndex        =   340
            Top             =   3360
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   8
            Left            =   2400
            TabIndex        =   339
            Top             =   4800
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   7
            Left            =   2400
            TabIndex        =   338
            Top             =   4440
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   6
            Left            =   2400
            TabIndex        =   337
            Top             =   4080
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   5
            Left            =   2400
            TabIndex        =   336
            Top             =   3720
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   4
            Left            =   2400
            TabIndex        =   335
            Top             =   3360
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   3
            Left            =   5580
            TabIndex        =   323
            Top             =   2490
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   2
            Left            =   5580
            TabIndex        =   322
            Top             =   1890
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   1
            Left            =   5580
            TabIndex        =   321
            Top             =   1290
            Width           =   255
         End
         Begin VB.CheckBox TestAndInspectionGood 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   0
            Left            =   5580
            TabIndex        =   320
            Top             =   690
            Width           =   255
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   6
            Left            =   840
            TabIndex        =   315
            Top             =   2520
            Width           =   735
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   7
            Left            =   3120
            TabIndex        =   314
            Top             =   2520
            Width           =   735
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   4
            Left            =   840
            TabIndex        =   313
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   5
            Left            =   3120
            TabIndex        =   312
            Top             =   1920
            Width           =   735
         End
         Begin VB.ComboBox cmbTestAndInspection 
            Height          =   315
            Index           =   1
            ItemData        =   "Main.frx":1CE9
            Left            =   1680
            List            =   "Main.frx":1CF3
            TabIndex        =   311
            Top             =   2520
            Width           =   975
         End
         Begin VB.ComboBox cmbTestAndInspection 
            Height          =   315
            Index           =   0
            ItemData        =   "Main.frx":1D03
            Left            =   1680
            List            =   "Main.frx":1D0D
            TabIndex        =   310
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   3
            Left            =   3120
            TabIndex        =   307
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   2
            Left            =   840
            TabIndex        =   305
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   1
            Left            =   3120
            TabIndex        =   303
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   0
            Left            =   840
            TabIndex        =   297
            Top             =   720
            Width           =   735
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Supervisor Approval?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   85
            Left            =   1680
            TabIndex        =   347
            Top             =   5220
            Width           =   1935
         End
         Begin VB.Line Line3 
            X1              =   120
            X2              =   6000
            Y1              =   3000
            Y2              =   3000
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Good?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   84
            Left            =   5400
            TabIndex        =   346
            Top             =   3120
            Width           =   615
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Good?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   83
            Left            =   2220
            TabIndex        =   345
            Top             =   3120
            Width           =   615
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Nameplate Check?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   82
            Left            =   3600
            TabIndex        =   334
            Top             =   4860
            Width           =   1935
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Paint Check?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   81
            Left            =   3360
            TabIndex        =   333
            Top             =   4500
            Width           =   2175
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Clean, Purge and Seal?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   80
            Left            =   3600
            TabIndex        =   332
            Top             =   4140
            Width           =   1935
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "NPSH Test?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   79
            Left            =   3600
            TabIndex        =   331
            Top             =   3780
            Width           =   1935
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Hydraulic Test?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   78
            Left            =   3600
            TabIndex        =   330
            Top             =   3420
            Width           =   1935
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Hydrostatic Test?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   77
            Left            =   360
            TabIndex        =   329
            Top             =   4860
            Width           =   1935
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Motor Locked Rotor Test?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   76
            Left            =   120
            TabIndex        =   328
            Top             =   4500
            Width           =   2175
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Motor No-load Test?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   75
            Left            =   360
            TabIndex        =   327
            Top             =   4140
            Width           =   1935
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Outline Dimensions?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   74
            Left            =   360
            TabIndex        =   326
            Top             =   3780
            Width           =   1935
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "General Appearance?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   73
            Left            =   360
            TabIndex        =   325
            Top             =   3420
            Width           =   1935
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Good?"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   72
            Left            =   5400
            TabIndex        =   324
            Top             =   480
            Width           =   615
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   71
            Left            =   2640
            TabIndex        =   319
            Top             =   2520
            Width           =   495
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   70
            Left            =   2640
            TabIndex        =   318
            Top             =   1950
            Width           =   495
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "minutes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   69
            Left            =   3960
            TabIndex        =   317
            Top             =   2550
            Width           =   735
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "minutes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   3960
            TabIndex        =   316
            Top             =   1950
            Width           =   735
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "AC"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   68
            Left            =   240
            TabIndex        =   309
            Top             =   1350
            Width           =   495
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "minutes"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   67
            Left            =   3960
            TabIndex        =   308
            Top             =   1350
            Width           =   735
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "V       X"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   27
            Left            =   1680
            TabIndex        =   306
            Top             =   1350
            Width           =   615
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "MOhms Above"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   26
            Left            =   3960
            TabIndex        =   304
            Top             =   750
            Width           =   1335
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pneumatic Test for N2 Gas:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   25
            Left            =   120
            TabIndex        =   302
            Top             =   2280
            Width           =   2535
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Hydrostatic Test:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   24
            Left            =   120
            TabIndex        =   301
            Top             =   1680
            Width           =   1935
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Dielectric Strength:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   120
            TabIndex        =   300
            Top             =   1080
            Width           =   1935
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Insulation Resistance:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   120
            TabIndex        =   299
            Top             =   480
            Width           =   1935
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "V Megger"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   1680
            TabIndex        =   298
            Top             =   750
            Width           =   975
         End
      End
      Begin VB.TextBox txtImpellerDia 
         Height          =   315
         Left            =   9840
         TabIndex        =   264
         Top             =   4349
         Width           =   1335
      End
      Begin VB.CommandButton cmdClearPumpData 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Clear Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   11160
         Style           =   1  'Graphical
         TabIndex        =   262
         Top             =   480
         Width           =   1695
      End
      Begin VB.Frame frmMfr 
         Caption         =   "Manufacturer"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   615
         Left            =   6120
         TabIndex        =   214
         Top             =   360
         Width           =   2655
         Begin VB.OptionButton optMfr 
            Caption         =   "TEMC"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   1680
            TabIndex        =   216
            Top             =   240
            Width           =   855
         End
         Begin VB.OptionButton optMfr 
            Caption         =   "Chempump"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   215
            Top             =   240
            Value           =   -1  'True
            Width           =   1455
         End
      End
      Begin VB.Frame frmMagtrol 
         BackColor       =   &H8000000A&
         Caption         =   "Magtrol"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -74880
         TabIndex        =   66
         Top             =   3960
         Width           =   6735
         Begin VB.OptionButton optKW 
            Caption         =   "Use Ana In 4"
            Height          =   195
            Index           =   2
            Left            =   5280
            TabIndex        =   276
            Top             =   1680
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.OptionButton optKW 
            Caption         =   "Enter KW"
            Height          =   195
            Index           =   1
            Left            =   5280
            TabIndex        =   275
            Top             =   1440
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.OptionButton optKW 
            Caption         =   "Add 3 powers"
            Height          =   195
            Index           =   0
            Left            =   5280
            TabIndex        =   274
            Top             =   1200
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton cmdFindMagtrols 
            Caption         =   "Find Magtrols"
            Height          =   255
            Left            =   5040
            TabIndex        =   187
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox cmbMagtrol 
            Height          =   315
            Left            =   2520
            TabIndex        =   185
            Top             =   240
            Width           =   2055
         End
         Begin VB.TextBox txtV1 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   77
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtV2 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   76
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtV3 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1080
            TabIndex        =   75
            Top             =   1560
            Width           =   855
         End
         Begin VB.TextBox txtI1 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            TabIndex        =   74
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtI2 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            TabIndex        =   73
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtI3 
            Alignment       =   1  'Right Justify
            BeginProperty DataFormat 
               Type            =   1
               Format          =   "0.0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2160
            TabIndex        =   72
            Top             =   1560
            Width           =   855
         End
         Begin VB.TextBox txtP1 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3240
            TabIndex        =   71
            Top             =   840
            Width           =   975
         End
         Begin VB.TextBox txtP2 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3240
            TabIndex        =   70
            Top             =   1200
            Width           =   975
         End
         Begin VB.TextBox txtP3 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3240
            TabIndex        =   69
            Top             =   1560
            Width           =   975
         End
         Begin VB.TextBox txtPF 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   4320
            TabIndex        =   68
            Top             =   840
            Width           =   855
         End
         Begin VB.TextBox txtKW 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5520
            TabIndex        =   67
            Top             =   840
            Width           =   855
         End
         Begin VB.Shape shpGetMagtrolData 
            FillColor       =   &H0000FF00&
            FillStyle       =   0  'Solid
            Height          =   252
            Left            =   120
            Shape           =   3  'Circle
            Top             =   240
            Width           =   252
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            Caption         =   "Magtrol Select"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   9.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   -1  'True
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   28
            Left            =   600
            TabIndex        =   186
            Top             =   270
            Width           =   1815
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            Caption         =   "Voltage"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   29
            Left            =   1080
            TabIndex        =   85
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            Caption         =   "Current"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   30
            Left            =   2160
            TabIndex        =   84
            Top             =   600
            Width           =   855
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            Caption         =   "Power"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   31
            Left            =   3240
            TabIndex        =   83
            Top             =   600
            Width           =   975
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            Caption         =   "Phase 1"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   32
            Left            =   240
            TabIndex        =   82
            Top             =   840
            Width           =   735
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            Caption         =   "Phase 2"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   33
            Left            =   240
            TabIndex        =   81
            Top             =   1260
            Width           =   735
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            Caption         =   "Phase 3"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   34
            Left            =   240
            TabIndex        =   80
            Top             =   1620
            Width           =   735
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            Caption         =   "PF"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   35
            Left            =   4440
            TabIndex        =   79
            Top             =   600
            Width           =   615
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            Caption         =   "Total KW"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   36
            Left            =   5520
            TabIndex        =   78
            Top             =   600
            Width           =   855
         End
      End
      Begin VB.CommandButton cmdApproveTestDate 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Approve/Unapprove This Test Date"
         Height          =   615
         Left            =   -62760
         Style           =   1  'Graphical
         TabIndex        =   182
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdApprovePump 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Approve/Unapprove This Pump"
         Height          =   615
         Left            =   13080
         Style           =   1  'Graphical
         TabIndex        =   181
         Top             =   480
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdDeleteTestDate 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Delete This Test Date"
         Height          =   615
         Left            =   -64440
         Style           =   1  'Graphical
         TabIndex        =   178
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdDeletePump 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Delete This Pump"
         Height          =   615
         Left            =   13080
         Style           =   1  'Graphical
         TabIndex        =   177
         Top             =   1080
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSComDlg.CommonDialog CommonDialog1 
         Left            =   -60960
         Top             =   600
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
         DialogTitle     =   "Browse for File"
      End
      Begin VB.TextBox txtTDH 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -61920
         TabIndex        =   172
         Top             =   4200
         Width           =   1455
      End
      Begin VB.Frame frmOtherFiles 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Other Files"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1335
         Left            =   -69840
         TabIndex        =   171
         Top             =   5040
         Visible         =   0   'False
         Width           =   3255
         Begin VB.TextBox txtNPSHFile 
            Height          =   285
            Left            =   2040
            TabIndex        =   35
            Top             =   240
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox chkNPSH 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "NPSH Data:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   34
            Top             =   195
            Width           =   1455
         End
         Begin VB.CheckBox chkPictures 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pictures:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   36
            Top             =   555
            Width           =   1455
         End
         Begin VB.CheckBox chkVibration 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "FFT Vibration"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   240
            TabIndex        =   38
            Top             =   915
            Width           =   1455
         End
         Begin VB.TextBox txtPicturesFile 
            Height          =   285
            Left            =   2040
            TabIndex        =   37
            Top             =   600
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.TextBox txtVibrationFile 
            Height          =   285
            Left            =   2040
            TabIndex        =   39
            Top             =   960
            Visible         =   0   'False
            Width           =   1095
         End
      End
      Begin VB.Frame frmPerfMods 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Performance Modifications"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1935
         Left            =   -69840
         TabIndex        =   168
         Top             =   3000
         Width           =   3255
         Begin VB.TextBox txtOrifice 
            Height          =   285
            Left            =   1680
            TabIndex        =   31
            Top             =   1365
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.TextBox txtImpTrim 
            Height          =   285
            Left            =   1680
            TabIndex        =   29
            Top             =   825
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.CheckBox chkOrifice 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Pump Discharge Orifice:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Left            =   240
            TabIndex        =   30
            Top             =   1200
            Width           =   1215
         End
         Begin VB.CheckBox chkTrimmed 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Impeller Trimmed:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   28
            Top             =   720
            Width           =   1215
         End
         Begin VB.CheckBox chkFeathered 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Impeller Feathered:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   240
            TabIndex        =   27
            Top             =   240
            Width           =   1215
         End
         Begin VB.Label lblOrifice 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Orifice Diameter"
            Height          =   255
            Left            =   1560
            TabIndex        =   170
            Top             =   1200
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.Label lblImpTrim 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Impeller Diameter"
            Height          =   255
            Left            =   1620
            TabIndex        =   169
            Top             =   600
            Visible         =   0   'False
            Width           =   1335
         End
      End
      Begin VB.Frame frmThrustBalMods 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Thrust Balance Modifications"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   -74880
         TabIndex        =   165
         Top             =   6480
         Width           =   7215
         Begin VB.TextBox txtGGap 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1200
            TabIndex        =   386
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdModifyBalanceHoleData 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Modify Balance Hole Data"
            Height          =   495
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   295
            Top             =   120
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.TextBox txtCircOrifice 
            Height          =   405
            Left            =   4080
            TabIndex        =   22
            Top             =   1680
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.CheckBox chkCircOrifice 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Circulation Flow Orifice:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   21
            Top             =   1680
            Width           =   1815
         End
         Begin VB.CommandButton cmdAddNewBalanceHoles 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Add New Balance Hole Data"
            Height          =   495
            Left            =   3000
            Style           =   1  'Graphical
            TabIndex        =   175
            Top             =   120
            Width           =   1575
         End
         Begin VB.CheckBox chkBalanceHoles 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Balance Holes Modified:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Left            =   360
            TabIndex        =   20
            Top             =   1080
            Width           =   1815
         End
         Begin VB.TextBox txtOtherMods 
            Height          =   555
            Left            =   1680
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   23
            Top             =   2280
            Width           =   5055
         End
         Begin VB.TextBox txtEndPlay 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1200
            TabIndex        =   19
            Top             =   330
            Width           =   975
         End
         Begin MSDataGridLib.DataGrid dgBalanceHoles 
            Height          =   975
            Left            =   2400
            TabIndex        =   174
            ToolTipText     =   "Click left column (where arrow is) to select to modify or delete. Choose date in Test Data above to add new data."
            Top             =   600
            Visible         =   0   'False
            Width           =   4695
            _ExtentX        =   8281
            _ExtentY        =   1720
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777152
            Enabled         =   -1  'True
            ForeColor       =   0
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            ColumnCount     =   2
            BeginProperty Column00 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   1
                  Format          =   "0"
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   1
               EndProperty
            EndProperty
            BeginProperty Column01 
               DataField       =   ""
               Caption         =   ""
               BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
                  Type            =   0
                  Format          =   ""
                  HaveTrueFalseNull=   0
                  FirstDayOfWeek  =   0
                  FirstWeekOfYear =   0
                  LCID            =   1033
                  SubFormatType   =   0
               EndProperty
            EndProperty
            SplitCount      =   1
            BeginProperty Split0 
               BeginProperty Column00 
               EndProperty
               BeginProperty Column01 
               EndProperty
            EndProperty
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "G-Gap:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   89
            Left            =   240
            TabIndex        =   387
            Top             =   750
            Width           =   855
         End
         Begin VB.Label lblCircOrifice 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Orifice Diameter:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   2400
            TabIndex        =   176
            Top             =   1800
            Visible         =   0   'False
            Width           =   1455
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Other Mods:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   12
            Left            =   240
            TabIndex        =   167
            Top             =   2400
            Width           =   1335
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "End Play:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   11
            Left            =   240
            TabIndex        =   166
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame frmElecData 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Electrical Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -69840
         TabIndex        =   159
         Top             =   1440
         Width           =   3255
         Begin VB.TextBox txtVFDFreq 
            Height          =   315
            Left            =   2040
            TabIndex        =   373
            Top             =   600
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.TextBox txtKWMult 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2040
            TabIndex        =   25
            Top             =   1080
            Visible         =   0   'False
            Width           =   735
         End
         Begin VB.ComboBox cmbFrequency 
            Height          =   315
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   600
            Width           =   1175
         End
         Begin VB.ComboBox cmbVoltage 
            Height          =   315
            Left            =   600
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   240
            Width           =   1175
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "VFD Freq:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   86
            Left            =   1920
            TabIndex        =   374
            Top             =   240
            Visible         =   0   'False
            Width           =   975
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "KW Multiplier:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   720
            TabIndex        =   162
            Top             =   1110
            Visible         =   0   'False
            Width           =   1215
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Freq:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   10
            Left            =   120
            TabIndex        =   161
            Top             =   630
            Width           =   495
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Volt:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   160
            Top             =   270
            Width           =   375
         End
      End
      Begin VB.Frame frmLoopAndXducer 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Loop and Transducer (Gauge) Setup"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   4935
         Left            =   -74880
         TabIndex        =   154
         Top             =   1440
         Width           =   4935
         Begin VB.ComboBox cmbMounting 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   277
            Top             =   1080
            Width           =   1335
         End
         Begin VB.ComboBox cmbLoopNumber 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cmbOrificeNumber 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   720
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.TextBox txtHDCor 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   2520
            TabIndex        =   18
            Top             =   3360
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.ComboBox cmbSuctDia 
            Height          =   288
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   14
            Top             =   2460
            Width           =   1335
         End
         Begin VB.ComboBox cmbDischDia 
            Height          =   288
            Left            =   2040
            Style           =   2  'Dropdown List
            TabIndex        =   15
            Top             =   2940
            Width           =   1335
         End
         Begin VB.TextBox txtDischHeight 
            Height          =   375
            Left            =   3600
            TabIndex        =   17
            Top             =   2880
            Width           =   615
         End
         Begin VB.TextBox txtSuctHeight 
            Height          =   375
            Left            =   3600
            TabIndex        =   16
            Top             =   2400
            Width           =   615
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Mounting:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   66
            Left            =   720
            TabIndex        =   278
            Top             =   1110
            Width           =   1095
         End
         Begin VB.Label Label15 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Click here for a diagram"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   13.5
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Left            =   720
            TabIndex        =   183
            Top             =   4080
            Width           =   3615
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Orifice Number:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   240
            TabIndex        =   164
            Top             =   750
            Visible         =   0   'False
            Width           =   1575
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Loop Number:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   163
            Top             =   390
            Width           =   1575
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "HD Cor:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   1560
            TabIndex        =   158
            Top             =   3390
            Visible         =   0   'False
            Width           =   855
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Suction Diameter:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   5
            Left            =   360
            TabIndex        =   157
            Top             =   2496
            Width           =   1572
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Discharge Diameter:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   252
            Index           =   6
            Left            =   120
            TabIndex        =   156
            Top             =   2976
            Width           =   1812
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Transducer Height (in Inches)"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   615
            Index           =   4
            Left            =   3360
            TabIndex        =   155
            Top             =   1680
            Width           =   1095
         End
      End
      Begin VB.Frame frmInstrumentTags 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Instrument Identification (Tags)"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2055
         Left            =   -66480
         TabIndex        =   146
         Top             =   1440
         Width           =   6135
         Begin VB.ComboBox cmbCirculationFlowMeter 
            Height          =   315
            Left            =   4560
            TabIndex        =   393
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cmbTemperatureTransducer 
            Height          =   315
            Left            =   1440
            TabIndex        =   392
            Top             =   1560
            Width           =   1335
         End
         Begin VB.ComboBox cmbDischargePressureTransducer 
            Height          =   315
            Left            =   1440
            TabIndex        =   391
            Top             =   1160
            Width           =   1335
         End
         Begin VB.ComboBox cmbSuctionPressureTransducer 
            Height          =   315
            Left            =   1440
            TabIndex        =   390
            Top             =   760
            Width           =   1335
         End
         Begin VB.ComboBox cmbFlowMeter 
            Height          =   315
            Left            =   1440
            TabIndex        =   389
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cmbPLCNo 
            Height          =   315
            ItemData        =   "Main.frx":1D1D
            Left            =   4560
            List            =   "Main.frx":1D1F
            Style           =   2  'Dropdown List
            TabIndex        =   377
            Top             =   1620
            Width           =   1335
         End
         Begin VB.ComboBox cmbTachID 
            Height          =   315
            Left            =   4560
            Style           =   2  'Dropdown List
            TabIndex        =   32
            Top             =   780
            Width           =   1335
         End
         Begin VB.ComboBox cmbAnalyzerNo 
            Height          =   315
            ItemData        =   "Main.frx":1D21
            Left            =   4560
            List            =   "Main.frx":1D23
            Style           =   2  'Dropdown List
            TabIndex        =   33
            Top             =   1200
            Width           =   1335
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "PLC:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   87
            Left            =   3480
            TabIndex        =   378
            Top             =   1680
            Width           =   975
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Flowmeter:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   165
            Index           =   13
            Left            =   360
            TabIndex        =   153
            Top             =   435
            Width           =   975
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Suction:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   240
            Index           =   14
            Left            =   600
            TabIndex        =   152
            Top             =   810
            Width           =   735
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Discharge:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   15
            Left            =   360
            TabIndex        =   151
            Top             =   1185
            Width           =   975
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Temperature:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   16
            Left            =   120
            TabIndex        =   150
            Top             =   1590
            Width           =   1215
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Circulation Flowmeter:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   435
            Index           =   61
            Left            =   3360
            TabIndex        =   149
            Top             =   300
            Width           =   1095
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Tach:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   62
            Left            =   3720
            TabIndex        =   148
            Top             =   810
            Width           =   615
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Analyzer:"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   17
            Left            =   3480
            TabIndex        =   147
            Top             =   1230
            Width           =   975
         End
      End
      Begin VB.TextBox txtTestSetupRemarks 
         Height          =   375
         Left            =   -72360
         MaxLength       =   150
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   40
         Top             =   9720
         Width           =   8655
      End
      Begin VB.CommandButton cmdReport 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Output Hydraulic Test Report To Excel"
         Height          =   615
         Left            =   -62040
         Style           =   1  'Graphical
         TabIndex        =   143
         Top             =   4800
         Width           =   1812
      End
      Begin VB.TextBox txtNPSHa 
         Alignment       =   2  'Center
         Enabled         =   0   'False
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   495
         Left            =   -61920
         TabIndex        =   141
         Top             =   3480
         Width           =   1455
      End
      Begin VB.Frame frmPumpData 
         Caption         =   "Transducers"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74880
         TabIndex        =   128
         Top             =   720
         Width           =   6735
         Begin VB.TextBox txtTemperatureDisplay 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5280
            TabIndex        =   136
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtFlowDisplay 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   480
            TabIndex        =   135
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtDischargeDisplay 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3675
            TabIndex        =   134
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtSuctionDisplay 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2085
            TabIndex        =   133
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtTemperature 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6360
            TabIndex        =   132
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtDischarge 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6360
            TabIndex        =   131
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtSuction 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6360
            TabIndex        =   130
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtFlow 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6360
            TabIndex        =   129
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.Label lblAutoMan 
            Alignment       =   1  'Right Justify
            Caption         =   "Auto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   3
            Left            =   4920
            TabIndex        =   269
            Top             =   720
            Width           =   300
         End
         Begin VB.Label lblAutoMan 
            Alignment       =   1  'Right Justify
            Caption         =   "Auto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   2
            Left            =   3320
            TabIndex        =   268
            Top             =   720
            Width           =   300
         End
         Begin VB.Label lblAutoMan 
            Alignment       =   1  'Right Justify
            Caption         =   "Auto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   1
            Left            =   1720
            TabIndex        =   267
            Top             =   720
            Width           =   300
         End
         Begin VB.Label lblAutoMan 
            Alignment       =   1  'Right Justify
            Caption         =   "Auto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   0
            Left            =   120
            TabIndex        =   266
            Top             =   720
            Width           =   300
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            Caption         =   "Temperature"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   19
            Left            =   5280
            TabIndex        =   140
            Top             =   360
            Width           =   1095
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            Caption         =   "Discharge Pressure"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   18
            Left            =   3675
            TabIndex        =   139
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblTab3 
            Alignment       =   2  'Center
            Caption         =   "Suction Pressure"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   1
            Left            =   2085
            TabIndex        =   138
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lblTab3 
            Alignment       =   2  'Center
            Caption         =   "Flow"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   480
            TabIndex        =   137
            Top             =   360
            Width           =   1095
         End
      End
      Begin VB.Frame frmPLCMisc 
         Caption         =   "PLC"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1695
         Left            =   -68040
         TabIndex        =   110
         Top             =   1020
         Width           =   4095
         Begin VB.TextBox txtManualLamp 
            Height          =   285
            Left            =   2520
            TabIndex        =   123
            Text            =   "Text1"
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtWriteSP 
            Height          =   375
            Left            =   2880
            TabIndex        =   122
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtWriteSPData 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1200
            TabIndex        =   121
            Text            =   "0"
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtSetPointDisplay 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   1200
            TabIndex        =   120
            Top             =   330
            Width           =   615
         End
         Begin VB.TextBox txtValvePositionDisplay 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3240
            TabIndex        =   119
            Top             =   300
            Width           =   615
         End
         Begin VB.TextBox txtValvePosition 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   3120
            TabIndex        =   118
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtDCoef 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2400
            TabIndex        =   117
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtICoef 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   3000
            TabIndex        =   116
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtPCoef 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2160
            TabIndex        =   115
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtSetPoint 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2280
            TabIndex        =   114
            Text            =   "0"
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.CommandButton cmdWriteSP 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Write SP"
            Height          =   405
            Left            =   960
            Style           =   1  'Graphical
            TabIndex        =   113
            Top             =   1200
            Width           =   855
         End
         Begin VB.TextBox txtInHgDisplay 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3240
            TabIndex        =   112
            Top             =   720
            Width           =   615
         End
         Begin VB.TextBox txtInHg 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2640
            TabIndex        =   111
            Top             =   1200
            Visible         =   0   'False
            Width           =   180
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            Caption         =   "Valve Position"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   495
            Index           =   44
            Left            =   2280
            TabIndex        =   127
            Top             =   240
            Width           =   735
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            Caption         =   "Set Point"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   42
            Left            =   240
            TabIndex        =   126
            Top             =   360
            Width           =   855
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            Caption         =   "SP to Write"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   375
            Index           =   43
            Left            =   120
            TabIndex        =   125
            Top             =   720
            Width           =   855
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            Caption         =   "In Hg"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   45
            Left            =   2160
            TabIndex        =   124
            Top             =   720
            Width           =   855
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   2415
         Left            =   -66840
         TabIndex        =   56
         Top             =   7620
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   4260
         _Version        =   393216
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   1
               Format          =   "0"
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   1
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.CommandButton cmdAddNewTestDate 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Add New Test Date"
         Height          =   615
         Left            =   -68280
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   660
         Width           =   2055
      End
      Begin VB.TextBox txtWho 
         Height          =   315
         Left            =   -73080
         TabIndex        =   11
         Top             =   900
         Width           =   1695
      End
      Begin VB.TextBox txtSalesOrderNumber 
         Height          =   315
         Left            =   1920
         TabIndex        =   1
         Top             =   420
         Width           =   1575
      End
      Begin VB.CommandButton cmdEnterTestSetupData 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Enter Test Setup Data"
         Height          =   615
         Left            =   -66000
         Style           =   1  'Graphical
         TabIndex        =   106
         Top             =   660
         Width           =   1215
      End
      Begin VB.CommandButton cmdEnterPumpData 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Enter Pump Data"
         Height          =   615
         Left            =   9360
         MaskColor       =   &H00E0E0E0&
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   480
         Width           =   1695
      End
      Begin VB.CommandButton cmdEnterTestData 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Enter Test Data"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   855
         Left            =   -74760
         Style           =   1  'Graphical
         TabIndex        =   104
         Top             =   7800
         Width           =   1575
      End
      Begin VB.Frame fmrMiscTestData 
         BackColor       =   &H00FFFFC0&
         Caption         =   "Miscellaneous"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1455
         Left            =   -74760
         TabIndex        =   97
         Top             =   6000
         Width           =   14535
         Begin VB.TextBox txtTEMCTRGReading 
            Height          =   285
            Left            =   7080
            TabIndex        =   381
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtVibAx 
            Height          =   285
            Left            =   6000
            TabIndex        =   380
            Top             =   480
            Width           =   855
         End
         Begin VB.Frame frmTEMCData 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Thrust Measurement"
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1335
            Left            =   8400
            TabIndex        =   245
            Top             =   120
            Visible         =   0   'False
            Width           =   6015
            Begin VB.TextBox txtThrustBal 
               Height          =   285
               Left            =   2800
               TabIndex        =   427
               Top             =   390
               Width           =   855
            End
            Begin VB.TextBox txtRevHead 
               Height          =   285
               Left            =   2800
               TabIndex        =   394
               Top             =   960
               Width           =   855
            End
            Begin VB.TextBox txtTEMCPVValue 
               Height          =   285
               Left            =   4080
               TabIndex        =   258
               Top             =   960
               Width           =   855
            End
            Begin VB.TextBox txtTEMCCalcForce 
               Height          =   285
               Left            =   4080
               TabIndex        =   256
               Top             =   390
               Width           =   855
            End
            Begin VB.TextBox txtTEMCViscosity 
               Height          =   225
               Left            =   2520
               TabIndex        =   254
               Top             =   990
               Visible         =   0   'False
               Width           =   150
            End
            Begin VB.TextBox txtTEMCThrustRigPressure 
               Height          =   285
               Left            =   1520
               TabIndex        =   252
               Top             =   990
               Width           =   855
            End
            Begin VB.TextBox txtTEMCMomentArm 
               Height          =   285
               Left            =   1520
               TabIndex        =   249
               Top             =   390
               Width           =   855
            End
            Begin VB.TextBox txtTEMCRearThrust 
               Height          =   285
               Left            =   240
               TabIndex        =   248
               Top             =   990
               Width           =   855
            End
            Begin VB.TextBox txtTEMCFrontThrust 
               Height          =   285
               Left            =   240
               TabIndex        =   247
               Top             =   390
               Width           =   855
            End
            Begin VB.Label lbltab2 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "Rig Pressure"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   51
               Left            =   1320
               TabIndex        =   253
               Top             =   750
               Width           =   1215
            End
            Begin VB.Label lbltab2 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "Axial Position"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   38
               Left            =   2640
               TabIndex        =   428
               Top             =   165
               Width           =   1215
            End
            Begin VB.Label lbltab2 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "HR (ft)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   90
               Left            =   2760
               TabIndex        =   395
               Top             =   720
               Width           =   855
            End
            Begin VB.Label lblTEMCFrontRear 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "Label1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5040
               TabIndex        =   261
               Top             =   390
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label lblTEMCPassFail 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "Label1"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   12
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Left            =   5040
               TabIndex        =   260
               Top             =   750
               Visible         =   0   'False
               Width           =   855
            End
            Begin VB.Label txtTEMCPV 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "PV (SG)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   53
               Left            =   4080
               TabIndex        =   259
               Top             =   720
               Width           =   855
            End
            Begin VB.Label lbltab2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "Calc Force (SG)"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   47
               Left            =   3780
               TabIndex        =   257
               Top             =   187
               Width           =   1455
            End
            Begin VB.Label lbltab2 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "Viscosity"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   195
               Index           =   52
               Left            =   2520
               TabIndex        =   255
               Top             =   780
               Visible         =   0   'False
               Width           =   255
            End
            Begin VB.Label lbltab2 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "Front Thrust"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   50
               Left            =   120
               TabIndex        =   251
               Top             =   165
               Width           =   1095
            End
            Begin VB.Label lbltab2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "Moment Arm"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   49
               Left            =   1320
               TabIndex        =   250
               Top             =   187
               Width           =   1215
            End
            Begin VB.Label lbltab2 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "Rear Thrust"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   210
               Index           =   48
               Left            =   120
               TabIndex        =   246
               Top             =   720
               Width           =   1095
            End
         End
         Begin VB.TextBox txtRPM 
            Height          =   285
            Left            =   4680
            TabIndex        =   88
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtVibRad 
            Height          =   285
            Left            =   6000
            TabIndex        =   100
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtTestRemarks 
            Height          =   855
            Left            =   360
            MaxLength       =   80
            MultiLine       =   -1  'True
            ScrollBars      =   2  'Vertical
            TabIndex        =   98
            Top             =   480
            Width           =   3855
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "TRG"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   46
            Left            =   7200
            TabIndex        =   379
            Top             =   240
            Width           =   615
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "RPM"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   39
            Left            =   4860
            TabIndex        =   103
            Top             =   840
            Width           =   495
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "Y Vibration"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   41
            Left            =   5880
            TabIndex        =   102
            Top             =   840
            Width           =   1095
         End
         Begin VB.Label lbltab2 
            Alignment       =   2  'Center
            BackColor       =   &H00FFFFC0&
            Caption         =   "X Vibration"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   40
            Left            =   5880
            TabIndex        =   101
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lbltab2 
            BackColor       =   &H00FFFFC0&
            Caption         =   "Remarks"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   37
            Left            =   360
            TabIndex        =   99
            Top             =   240
            Width           =   1095
         End
      End
      Begin VB.ComboBox cmbTestSpec 
         Height          =   315
         Left            =   -73080
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   540
         Width           =   2055
      End
      Begin VB.TextBox txtRemarks 
         Height          =   555
         Left            =   3360
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   8700
         Width           =   7695
      End
      Begin VB.TextBox txtDesignTDH 
         Height          =   315
         Left            =   11160
         TabIndex        =   6
         Top             =   1770
         Width           =   1335
      End
      Begin VB.TextBox txtDesignFlow 
         Height          =   315
         Left            =   11160
         TabIndex        =   5
         Top             =   1410
         Width           =   1335
      End
      Begin VB.TextBox txtModelNo 
         Height          =   315
         Left            =   1320
         TabIndex        =   4
         Top             =   4349
         Width           =   2532
      End
      Begin VB.TextBox txtShpNo 
         Height          =   315
         Left            =   1440
         TabIndex        =   2
         Top             =   1410
         Width           =   4935
      End
      Begin VB.TextBox txtBilNo 
         Height          =   315
         Left            =   1440
         TabIndex        =   3
         Top             =   1770
         Width           =   4935
      End
      Begin VB.Frame frmThermocouples 
         Caption         =   "Thermocouples"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74880
         TabIndex        =   57
         Top             =   1800
         Width           =   6735
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   3552
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   286
            Text            =   "TC 3"
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   3552
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   285
            Text            =   "(F)"
            Top             =   480
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   5152
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   284
            Text            =   "TC 4"
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   5152
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   283
            Text            =   "(F)"
            Top             =   480
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   3
            Left            =   1952
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   282
            Text            =   "(F)"
            Top             =   480
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   2
            Left            =   1952
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   281
            Text            =   "TC 2"
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   352
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   280
            Text            =   "(F)"
            Top             =   480
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   0
            Left            =   352
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   279
            Text            =   "TC 1"
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox txtTC4 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6480
            TabIndex        =   65
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtTC3 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6480
            TabIndex        =   64
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtTC2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6480
            TabIndex        =   63
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtTC1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6480
            TabIndex        =   62
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtTC1Display 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   480
            TabIndex        =   61
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtTC2Display 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2080
            TabIndex        =   60
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtTC3Display 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3680
            TabIndex        =   59
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtTC4Display 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5280
            TabIndex        =   58
            Top             =   720
            Width           =   1095
         End
      End
      Begin VB.Frame frmAI 
         Caption         =   "Analog Inputs"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   1095
         Left            =   -74880
         TabIndex        =   47
         Top             =   2880
         Width           =   6735
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   24
            Left            =   3554
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   294
            Text            =   "P2"
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   25
            Left            =   3554
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   293
            Text            =   "(psig)"
            Top             =   480
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   26
            Left            =   5152
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   292
            Text            =   "AI 4"
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   27
            Left            =   5152
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   291
            Top             =   480
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   20
            Left            =   352
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   290
            Text            =   "Circ Flow"
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   21
            Left            =   352
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   289
            Text            =   "(GPM)"
            Top             =   480
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   22
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   288
            Text            =   "P1"
            Top             =   240
            Width           =   1350
         End
         Begin VB.TextBox txtTitle 
            Alignment       =   2  'Center
            Appearance      =   0  'Flat
            BackColor       =   &H00FFFFC0&
            BorderStyle     =   0  'None
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   23
            Left            =   1920
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   287
            Text            =   "(psig)"
            Top             =   480
            Width           =   1350
         End
         Begin VB.TextBox txtAI1 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6480
            TabIndex        =   55
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtAI2 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6480
            TabIndex        =   54
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtAI3 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6480
            TabIndex        =   53
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtAI4 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6480
            TabIndex        =   52
            Top             =   480
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtAI4Display 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   5280
            TabIndex        =   51
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtAI3Display 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   3682
            TabIndex        =   50
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtAI2Display 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   2085
            TabIndex        =   49
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtAI1Display 
            Alignment       =   1  'Right Justify
            Enabled         =   0   'False
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   285
            Left            =   480
            TabIndex        =   48
            Top             =   720
            Width           =   1095
         End
         Begin VB.Label lblAutoMan 
            Alignment       =   1  'Right Justify
            Caption         =   "Auto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   7
            Left            =   4920
            TabIndex        =   273
            Top             =   720
            Width           =   300
         End
         Begin VB.Label lblAutoMan 
            Alignment       =   1  'Right Justify
            Caption         =   "Man"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   6
            Left            =   3320
            TabIndex        =   272
            Top             =   720
            Width           =   300
         End
         Begin VB.Label lblAutoMan 
            Alignment       =   1  'Right Justify
            Caption         =   "Man"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   5
            Left            =   1720
            TabIndex        =   271
            Top             =   720
            Width           =   300
         End
         Begin VB.Label lblAutoMan 
            Alignment       =   1  'Right Justify
            Caption         =   "Auto"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   6.75
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   135
            Index           =   4
            Left            =   120
            TabIndex        =   270
            Top             =   720
            Width           =   300
         End
      End
      Begin VB.ComboBox cmbPLCLoop 
         Height          =   315
         ItemData        =   "Main.frx":1D25
         Left            =   -72480
         List            =   "Main.frx":1D27
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   420
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2415
         Left            =   -72960
         TabIndex        =   86
         Top             =   7620
         Width           =   6015
         _ExtentX        =   10610
         _ExtentY        =   4260
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ColumnCount     =   2
         BeginProperty Column00 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         BeginProperty Column01 
            DataField       =   ""
            Caption         =   ""
            BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
               Type            =   0
               Format          =   ""
               HaveTrueFalseNull=   0
               FirstDayOfWeek  =   0
               FirstWeekOfYear =   0
               LCID            =   1033
               SubFormatType   =   0
            EndProperty
         EndProperty
         SplitCount      =   1
         BeginProperty Split0 
            BeginProperty Column00 
            EndProperty
            BeginProperty Column01 
            EndProperty
         EndProperty
      End
      Begin VB.Frame frmChempump 
         Caption         =   "Chempump Pump Data"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2895
         Left            =   240
         TabIndex        =   217
         Top             =   5040
         Width           =   14415
         Begin VB.ComboBox cmbCirculationPath 
            Height          =   315
            ItemData        =   "Main.frx":1D29
            Left            =   1800
            List            =   "Main.frx":1D2B
            Style           =   2  'Dropdown List
            TabIndex        =   227
            Top             =   1860
            Width           =   3615
         End
         Begin VB.ComboBox cmbStatorFill 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   226
            Top             =   1140
            Width           =   3615
         End
         Begin VB.ComboBox cmbDesignPressure 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   225
            Top             =   1500
            Width           =   3615
         End
         Begin VB.ComboBox cmbRPM 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   224
            Top             =   780
            Width           =   3615
         End
         Begin VB.ComboBox cmbMotor 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   223
            Top             =   420
            Width           =   3615
         End
         Begin VB.Frame Frame7 
            BackColor       =   &H00FFFFC0&
            Caption         =   "User Entry - Model and Group"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   9
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   1575
            Left            =   5520
            TabIndex        =   218
            Top             =   360
            Width           =   5175
            Begin VB.ComboBox cmbModelGroup 
               Height          =   315
               Left            =   1440
               Style           =   2  'Dropdown List
               TabIndex        =   220
               Top             =   960
               Width           =   3615
            End
            Begin VB.ComboBox cmbModel 
               Height          =   315
               Left            =   1440
               Style           =   2  'Dropdown List
               TabIndex        =   219
               Top             =   480
               Width           =   3615
            End
            Begin VB.Label lbltab1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "Model:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   14
               Left            =   480
               TabIndex        =   222
               Top             =   540
               Width           =   855
            End
            Begin VB.Label lbltab1 
               Alignment       =   1  'Right Justify
               BackColor       =   &H00FFFFC0&
               Caption         =   "Model Group:"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   255
               Index           =   15
               Left            =   120
               TabIndex        =   221
               Top             =   1020
               Width           =   1215
            End
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Motor:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   840
            TabIndex        =   232
            Top             =   480
            Width           =   855
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Stator Fill:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   7
            Left            =   480
            TabIndex        =   231
            Top             =   1200
            Width           =   1215
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            Caption         =   "DesignPressure:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   8
            Left            =   120
            TabIndex        =   230
            Top             =   1560
            Width           =   1575
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Design RPM:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   360
            TabIndex        =   229
            Top             =   840
            Width           =   1335
         End
         Begin VB.Label lbltab1 
            Alignment       =   1  'Right Justify
            Caption         =   "Circulation Path:"
            BeginProperty Font 
               Name            =   "Arial"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   9
            Left            =   120
            TabIndex        =   228
            Top             =   1920
            Width           =   1575
         End
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   504
         Left            =   -61080
         TabIndex        =   426
         Top             =   5520
         Width           =   252
         _ExtentX        =   450
         _ExtentY        =   900
         _Version        =   393216
         Value           =   1
         BuddyControl    =   "txtUpDn2"
         BuddyDispid     =   196619
         OrigLeft        =   13920
         OrigTop         =   5520
         OrigRight       =   14172
         OrigBottom      =   6012
         Max             =   8
         Min             =   1
         SyncBuddy       =   -1  'True
         BuddyProperty   =   0
         Enabled         =   -1  'True
      End
      Begin VB.Shape shpGetPLCData 
         FillColor       =   &H0000FF00&
         FillStyle       =   0  'Solid
         Height          =   252
         Left            =   -74880
         Shape           =   3  'Circle
         Top             =   360
         Width           =   252
      End
      Begin VB.Label lbltab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Cust PO Num:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   50
         Left            =   6480
         TabIndex        =   421
         Top             =   1080
         Width           =   1212
      End
      Begin VB.Label lbltab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Customer PN:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   49
         Left            =   120
         TabIndex        =   419
         Top             =   1080
         Width           =   1212
      End
      Begin VB.Label lbltab1 
         Alignment       =   1  'Right Justify
         Caption         =   "RVS Part No:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   48
         Left            =   4200
         TabIndex        =   416
         Top             =   4380
         Width           =   1212
      End
      Begin VB.Label lbltab1 
         Caption         =   "SuperMarket Impeller Feathered:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   7.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   47
         Left            =   11760
         TabIndex        =   415
         Top             =   4320
         Width           =   2652
      End
      Begin VB.Label lbltab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Original Impeller Dia:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   46
         Left            =   9360
         TabIndex        =   414
         Top             =   3120
         Width           =   2052
      End
      Begin VB.Label lbltab2 
         Alignment       =   1  'Right Justify
         Caption         =   "RMA:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   88
         Left            =   -70920
         TabIndex        =   385
         Top             =   600
         Visible         =   0   'False
         Width           =   855
      End
      Begin VB.Label lbltab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Line Number:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   44
         Left            =   3960
         TabIndex        =   375
         Top             =   451
         Width           =   1212
      End
      Begin VB.Label lbltab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Original Impeller Dia:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   10
         Left            =   7680
         TabIndex        =   265
         Top             =   4380
         Width           =   2052
      End
      Begin VB.Label lbltab2 
         Alignment       =   2  'Center
         Caption         =   "Number of Points to Plot"
         Height          =   375
         Index           =   54
         Left            =   -62160
         TabIndex        =   184
         Top             =   5640
         Width           =   1095
      End
      Begin VB.Label lbltab2 
         Alignment       =   2  'Center
         Caption         =   "TDH"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   53
         Left            =   -61920
         TabIndex        =   173
         Top             =   3960
         Width           =   1452
      End
      Begin VB.Label lbltab2 
         Alignment       =   1  'Right Justify
         Caption         =   "Test Setup Remarks:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   65
         Left            =   -74400
         TabIndex        =   145
         Top             =   9780
         Width           =   1815
      End
      Begin VB.Label lbltab2 
         Alignment       =   2  'Center
         Caption         =   "NPSHa"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   64
         Left            =   -61920
         TabIndex        =   142
         Top             =   3240
         Width           =   1452
      End
      Begin VB.Label lbltab2 
         Alignment       =   1  'Right Justify
         Caption         =   "Operator:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   -74760
         TabIndex        =   108
         Top             =   930
         Width           =   1575
      End
      Begin VB.Label lbltab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Sales Order:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   720
         TabIndex        =   107
         Top             =   450
         Width           =   1095
      End
      Begin VB.Label lbltab2 
         Alignment       =   2  'Center
         Caption         =   "Test Number"
         Height          =   255
         Index           =   63
         Left            =   -74520
         TabIndex        =   105
         Top             =   9480
         Width           =   975
      End
      Begin VB.Label lbltab2 
         Alignment       =   1  'Right Justify
         Caption         =   "Test Specification:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   -74880
         TabIndex        =   95
         Top             =   570
         Width           =   1695
      End
      Begin VB.Label lbltab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Remarks:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   13
         Left            =   2400
         TabIndex        =   94
         Top             =   8850
         Width           =   855
      End
      Begin VB.Label lbltab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Design TDH:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   12
         Left            =   9720
         TabIndex        =   93
         Top             =   1800
         Width           =   1332
      End
      Begin VB.Label lbltab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Design Flow:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   11
         Left            =   9840
         TabIndex        =   92
         Top             =   1440
         Width           =   1212
      End
      Begin VB.Label lbltab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Model No:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   3
         Left            =   240
         TabIndex        =   91
         Top             =   4380
         Width           =   972
      End
      Begin VB.Label lbltab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Bill to:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   2
         Left            =   480
         TabIndex        =   90
         Top             =   1800
         Width           =   852
      End
      Begin VB.Label lbltab1 
         Alignment       =   1  'Right Justify
         Caption         =   "Ship to:"
         BeginProperty Font 
            Name            =   "Arial"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   252
         Index           =   1
         Left            =   480
         TabIndex        =   89
         Top             =   1440
         Width           =   852
      End
      Begin VB.Label lbltab2 
         Alignment       =   1  'Right Justify
         Caption         =   "PLC Select"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   -1  'True
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   55
         Left            =   -74160
         TabIndex        =   87
         Top             =   420
         Width           =   1455
      End
   End
   Begin VB.Label lblPumpApproved 
      Caption         =   "Pump Data Approved"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1800
      TabIndex        =   180
      Top             =   480
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label lblTestDateApproved 
      Caption         =   "Test Setup Data Approved"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   6480
      TabIndex        =   179
      Top             =   480
      Visible         =   0   'False
      Width           =   2535
   End
   Begin VB.Label lblVersion 
      Alignment       =   1  'Right Justify
      Caption         =   "Version 1.10"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   252
      Left            =   11640
      TabIndex        =   144
      Top             =   120
      Width           =   3492
   End
   Begin VB.Label lbltab2 
      Alignment       =   1  'Right Justify
      Caption         =   "Test Date:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   57
      Left            =   5640
      TabIndex        =   96
      Top             =   150
      Width           =   975
   End
   Begin VB.Label lbltab2 
      Alignment       =   1  'Right Justify
      Caption         =   "Serial No:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   56
      Left            =   480
      TabIndex        =   44
      Top             =   120
      Width           =   1215
   End
End
Attribute VB_Name = "frmPLCData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Polar V1.0.0 - MHR - 4/15/16
'   Release

'V1.0.1 - MHR - 6/14/16
'   Several fixes

'v1.0.2 - MHR - 6/14/16
'   Changed transducer text boxes to dropdowns

'v1.0.3 - MHR - 6/20/16
'   Misc changes
'   Added VBWatch

'v1.0.4 - MHR - 6/27/16
'   Modify Excel Output

'v1.0.5 - MHR- 7/5/16
'   fix pv calculation
'   only use number of entries for calculate BEP on excel sheet

'v1.0.6 - MHR - 7/21/16
'   added orderdtl.orderline as criterion in epicor select to get proper model number

'v1.0.7 - MHR - 7/29/16
'   if excel is running, set xlapp = excel.application
'   only allow one copy of rundown to run at a time
'   added supermarket model table and fill in data from table
'   created MagtrolRoutines module for easier change to Prologix when required
'   added hydraulic report to excel sheet

'v1.0.8 - uses prologix - branch from this

'v1.0.9 - MHR - 9/8/16
'   added single quotes around sheet name when writing to excel

'v1.0.10 - MHR - 10/27/16
'   Modified Excel Sheet

'v1.0.11 - MHR - 11/1/16
'   Changed Search from By Chempump Model to By TEMC Hydraulics
'   Added NPSH calcs and recording

'v1.0.12 - MHR - 11/4/16
'   Modified NPSHr

'v1.0.13 - MHR - 11/6/16
'   Added set save hydraulic test button in excel to macro

'v1.1.0 - MHR - 11/10/16
'   Eliminating local reports - only report to Excel Sheet
'   Adding NPSHr to Excel Sheet
'   Autoshow excel
'   autosave and autoshow hydraulic test report

'v1.1.1 - MHR - 11/14/16
'   Removed autoshow report on excel open
'   prevented continual update of NPSHr

'v1.1.2 - MHR - 11/15/16
'   Added NPSHr from Pump Tab to M8
'   Updated template to #5

'v1.1.3 - MHR - 11/21/16
'   removed npshr directory
'   modified grid2 columns
'   timer to close NPSHr test after NPSHr written

'v1.1.4 - MHR - 2/14/17
'   changed method for calculating maximum scale on excel chart
'   look at both uncorrected and corrected for maximum value

'v1.2.0 - MHR - 3/8/17
'   Added 380V to voltage dropdown
'   Replaced supermarket table with new, updated one
'   Added pop-up to alert user of special electrical set ups, like vfd or trg transformer
'   Modified RPM to account for frequencies other than 60Hz
'   Used new SG&Visc corrections sheet - rev 8
'   Get data from Epicor for supermarket sheets and fill in.  remove user dropdown to select supermarket pump
'   Add invisible Feathered checkbox on PumpData screen that is set from supermarket table.  use this box
'       to set feathered on pumpsetup screen
'   Change TEMCDesignPressure so that #4 is 600psi instead of 550psi
'   Scrub PolarPumpData

'v1.2.1 - MHR 3/9/2017
'   added save and restore supermarket feathered to/from TempPumpData

'v1.2.2 - MHR 3/10/2017
'   added RVS Part Numbers to Supermarket table and TempPumpData
'   added customer part number to temppumpdata
'   adjusted spreadsheet to list customer part number

'v1.2.3 - MHR
'   housekeeping

'v1.2.4 - MHR 3/17/16
'   allowed sn to either be aannnna or aannnna-n or aannnna-nn

'v1.2.5 - MHR 4/3/17
'   for frame number = 529, make frame = 420 for pv calculation

'v1.2.6 - MHR 4/3/17
'   Changed pv calculation to account for all frequencies

'v1.2.7 - MHR - 5/18/17
'   Removed efficiency from hyd rept graph and set max right hand scale based on current

'v1.2.8 - MHR - 5/19/17
'   removed draw eff curve on graph

'v1.2.9 - MHR - 5/31/17
'   removed efficiency text from right-hand axis

'v1.2.10 - MHR - 6/8/17
'   modified right hand axis scale to be 5 max, 10 max, or multiples of 25

'v1.2.11 - MHR - 8/18/17
'   Modified Epicor routine for E10

'v1.2.12 - MHR - 9/21/17
' Removed cwnumedit and replaced with up/down and removed cwchart and replaces with mschart
'   so we don't get error message about updating cw components

'v1.2.13 - MHR - 11/14/17
'   Changed updown2 text from 1 to 8
'   removed reference and code about interopdb
'   added prompt when "enter test setup data" for suct and disch diameter and transducer heights are not entered or are 0
'   added table, Transducers, to set dropdowns as loop is changed

'v1.2.14 - MHR - 12/13/17
'   removed remove series in chart in export to excel
'   was failing on excel 2016

'v1.2.15 - MHR - 12/19/17
'   added excel 16.0 library for Excel 2016

'v1.2.16 - MHR - 12/26/17
'   allow search grids to sort on columns

'v1.2.17 - MHR - 12/31/17
'   fixed parsing of model number to make circulation always [*]

'v1.2.18 - MHR - 1/16/18
'   recompiled with office 2010
'   had to remove button to write hydraulic test report
'   changed to late binding on excel

'v1.2.19 - MHR - 1/31/18
'sync'ed updown2.value with txtupdn2

'v1.2.20 - MHR - 3/14/18
'   Changed Supermarket table
'   Changed alerts for model number trg and additions
'   Added 380 to Voltage dropdown
'   resized head-flow chart
'   allowed plc selection on test setup page to set actual plc on test data page
'   allowed plc loop selection on test setup page to set plc and gpib

'v1.2.21 - MHR - 3/26/18
'   Fixed FixPointsToPlot routine to prevent setting number of points to 0
'   Fixed cmbCirculationFlowMeter to store proper data

'v1.2.22 - MHR - 5/11/18
'   Added figure to transducer diagram
'   changed A2 and A3 default to Man from Auto
'   Default Frequency to 60Hz

'v1.2.23 - MHR - 5/21/18
'   made txtTEMCviscosity.text = txtViscosity.text / SG
'   modified TEMC thrust panel layout

'v1.2.24 - MHR - 6/13/18
'   only plot current once
'   output number to excel sheet for impdia in M1

'v1.2.25 - MHR - 6/26/18
'   Added check for "4" in the third position of model number to alert for CO2 pump

'v1.2.26 - MHR - 8/16/18
'   Changed %End Play output to Excel to formatted to 1 dec place

'v1.2.27 - MHR - 10/10/19
'   Allow serial numbers for LE pumps like 2019010387

    Option Explicit

    Dim debugging As Integer        'debugging 1=true 0=false
    Dim sDataBaseName As String



'    Dim boUsingHP As Boolean            'We're using the HP database
    Dim boFoundPump As Boolean          'found the pump in database
    Dim boPumpIsApproved As Boolean     'pump data is approved
    Dim boTestDateIsApproved As Boolean 'data for this date is approved
    Dim boFoundTestSetup As Boolean     'found setup data
    Dim boFoundTestData As Boolean      'found test data
    Dim boUsingEpicor As Boolean        'search epicor for pump
    Dim boUsingSupermarketTable As Boolean 'load from supermarket table
    Dim boEpicorFound As Boolean        'epicor found the pump

    Dim boPLCOperating As Boolean       'is the PLC working?
    Dim boMagtrolOperating As Boolean   'is Magtrol working?
    Dim boGotBalanceHoles As Boolean              'do we have any balance hole data?

    'recordsets
    Dim rsPumpData As New ADODB.Recordset       'PumpData recordset
    Dim rsTestSetup As New ADODB.Recordset      'TestSetup recordset
    Dim rsTestData As New ADODB.Recordset       'Test Data recordset
    Dim rsMisc As New ADODB.Recordset           'Misc Parameters
    Dim rsEff As New ADODB.Recordset            'Efficiency Calcs
    Dim rsBalanceHoles As New ADODB.Recordset   'Balance holes
    Dim rsPumpParameters As New ADODB.Recordset 'Other parameters
    Dim rsSupermarketModel As New ADODB.Recordset

    'commands
    Dim qyPumpData As New ADODB.Command         'Query for PumpData
    Dim qyTestSetup As New ADODB.Command        'Query for TestSetup
    Dim qyBalanceHoles As New ADODB.Command     'query for Balance Holes
    Dim qySupermarketModel As New ADODB.Command 'query for supermarket pump
    Dim qyMisc As New ADODB.Command             'query for misc parameters

    'array for head/flow chart
    Dim HeadFlow(1, 7) As Single            'x and y
    Dim EffFlow(1, 7) As Single
    Dim KWFlow(1, 7) As Single
    Dim AmpsFlow(1, 7) As Single
    Dim FlowHead(7, 1) As Single


    Dim RatedKW As Single               'TEMC Motor rated output

    Dim blnEnabled As Boolean           'auto enabled

    Dim EpicorConnectionString As String
    Dim ParentDirectoryName As String

    Dim xlApp As Object  ' Excel Application Object
'    Dim xlApp As Excel.Application  ' Excel Application Object
'    Dim xlBook As Excel.Workbook    ' Excel Workbook Object
    Dim xlBook As Object    ' Excel Workbook Object
    'Efficiency Database Name
    Const sEffDataBaseName As String = "\eff.mdb"
   
    'Server Name Text File
    Const sServerNameTextFile = "C:\Server.txt"

    'HP Database Path
    Const sHPDataBaseName As String = "Provider=MSDASQL.1;Persist Security Info=False;Data Source=HP-3000/32"

      'mdb at f:\groups\dev\3393567 where database names and locations reside
      ' we're using f:\ instead of a fully qualified unc since the names of the servers change

    Const sDevelopmentDatabase = "\Groups\DEV\3393567\Development.mdb"
    Const sSGandViscSpreadsheetTemplate = "\Polar SG&Visc Correction12.xls"
    Const sSaveFileMacroFile = "\savefile.bas"

    Dim ProgramEnd As Boolean 'we want to end the program
    Dim Pressed As Boolean  'was cmdEnterTestData called because user pressed the button


Private Sub chkBalanceHoles_Click()
    'if the balance holes box is checked, show the datagrid
    If chkBalanceHoles.value = 1 Then
        dgBalanceHoles.Visible = True
    Else
        dgBalanceHoles.Visible = False
    End If
    If LenB(frmPLCData.txtSN.Text) = 0 Or LenB(cmbTestDate.Text) = 0 Then
        dgBalanceHoles.Visible = False
    End If
End Sub

Private Sub chkCircOrifice_Click()
    'if the CircOrifice box is checked, show the size
    If chkCircOrifice.value = 1 Then
        lblCircOrifice.Visible = True
        txtCircOrifice.Visible = True
    Else
        lblCircOrifice.Visible = False
        txtCircOrifice.Visible = False
    End If
End Sub

Private Sub chkNPSH_Click()
    'if the NPSH file box is checked, show the file name
    If chkNPSH.value = 1 Then
        txtNPSHFile.Visible = True
    Else
        txtNPSHFile.Visible = False
    End If
End Sub

Private Sub chkOrifice_Click()
    'if the orifice box is checked, show the size
    If chkOrifice.value = 1 Then
        lblOrifice.Visible = True
        txtOrifice.Visible = True
    Else
        lblOrifice.Visible = False
        txtOrifice.Visible = False
    End If
End Sub

Private Sub chkPictures_Click()
    'if the pictures box is checked, show the file name
    If chkPictures.value = 1 Then
        txtPicturesFile.Visible = True
    Else
        txtPicturesFile.Visible = False
    End If
End Sub

Private Sub chkTrimmed_Click()
    'if the trimmed box is checked, show the impeller size
    If chkTrimmed.value = 1 Then
        lblImpTrim.Visible = True
        txtImpTrim.Visible = True
    Else
        lblImpTrim.Visible = False
        txtImpTrim.Visible = False
    End If
End Sub

Private Sub chkVibration_Click()
    'if the vibration box is checked, show the file name
    If chkVibration.value = 1 Then
        txtVibrationFile.Visible = True
    Else
        txtVibrationFile.Visible = False
    End If
End Sub



Private Sub cmbFrequency_Click()
    If cmbFrequency.Text = "VFD" Then
        txtVFDFreq.Visible = True
        lbltab2(86).Visible = True
    Else
        txtVFDFreq.Visible = False
        lbltab2(86).Visible = False
    End If
End Sub


Private Sub cmbLoopNumber_Click()

    Dim I As Integer
    I = cmbLoopNumber.ListIndex

    Dim qyTransducers As New ADODB.Command
    Dim rsTransducers As New ADODB.Recordset
    qyTransducers.ActiveConnection = cnPumpData
    qyTransducers.CommandText = "SELECT * " & _
              "From Transducers " & _
              "Where LoopNumber  = " & I

    With rsTransducers     'open the recordset for the query
'        .Index = "FindData"
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .Open qyTransducers
    End With
    If rsTransducers.RecordCount = 1 Then
        Me.cmbFlowMeter.ListIndex = rsTransducers.Fields("FlowMeter")
        Me.cmbSuctionPressureTransducer.ListIndex = rsTransducers.Fields("SuctionPressure")
        Me.cmbDischargePressureTransducer.ListIndex = rsTransducers.Fields("DischargePressure")
        Me.cmbTemperatureTransducer.ListIndex = rsTransducers.Fields("Temperature")
        Me.cmbPLCNo.ListIndex = rsTransducers.Fields("PLC")
        Me.cmbAnalyzerNo.ListIndex = rsTransducers.Fields("Analyzer")
        Me.cmbCirculationFlowMeter.ListIndex = rsTransducers.Fields("CircFlowMeter")
    End If
    
'    If I < 2 Then
'        Me.cmbPLCNo.ListIndex = 0
'    Else
'        Me.cmbPLCNo.ListIndex = 1
'    End If
End Sub

Private Sub GetSuperMarketPump(SuperMarketPartNum As String, JobNumber As String)

    'get the data from the SupermarketPumpData table
    qySupermarketModel.ActiveConnection = cnPumpData
    qySupermarketModel.CommandText = "SELECT * " & _
              "From SupermarketPumpData " & _
              "Where Model  = '" & SuperMarketPartNum & "'"

              'cmbSupermarketModel.ItemData(cmbSupermarketModel.ListIndex)"

    If rsSupermarketModel.State = adStateOpen Then
        rsSupermarketModel.Close
    End If

    With rsSupermarketModel     'open the recordset for the query
'        .Index = "FindData"
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .Open qySupermarketModel
    End With
    If rsSupermarketModel.RecordCount = 1 Then
        txtSalesOrderNumber.Text = rsSupermarketModel.Fields("SalesOrder")
        txtLineNumber.Text = rsSupermarketModel.Fields("LineNumber")
        txtShpNo.Text = rsSupermarketModel.Fields("ShipTo")
        txtBilNo.Text = rsSupermarketModel.Fields("BillTo")
        txtDesignFlow.Text = rsSupermarketModel.Fields("DesignFlow")
        txtDesignTDH.Text = rsSupermarketModel.Fields("DesignTDH")
        txtNoPhases.Text = rsSupermarketModel.Fields("Phases")
        txtNPSHr.Text = rsSupermarketModel.Fields("NPSHr")
        txtRatedInputPower.Text = rsSupermarketModel.Fields("RatedInputPower")
        txtAmps.Text = rsSupermarketModel.Fields("RatedCurrent")
        txtThermalClass.Text = rsSupermarketModel.Fields("ThermalClass")
        txtSpGr.Text = rsSupermarketModel.Fields("SG")
        txtViscosity.Text = rsSupermarketModel.Fields("Viscosity")
        txtTEMCViscosity.Text = Format((Val(rsSupermarketModel.Fields("Viscosity")) / Val(txtSpGr.Text)), "000.00")
        txtExpClass.Text = rsSupermarketModel.Fields("EXPClass")
        txtLiquid.Text = rsSupermarketModel.Fields("Liquid")
        txtLiquidTemperature.Text = rsSupermarketModel.Fields("LiquidTemp")
        txtJobNum.Text = JobNumber
        txtImpellerDia.Text = rsSupermarketModel.Fields("ImpellerDiameter")
        txtModelNo.Text = rsSupermarketModel.Fields("Model")
        txtRVSPartNo.Text = rsSupermarketModel.Fields("RVSPartNo")
        cmdSelectSupermarket.Caption = "Save Data"
        If UCase(rsSupermarketModel.Fields("Feathered")) = "FEATHERED" Then
            Me.chkSuperMarketFeathered.value = Checked
        End If
    End If
    grpSupermarket.Visible = False
  
End Sub

Private Sub cmbPLCNo_Click()
    'cmbplc text either contains 8 or 9
    
    Dim I As Integer
    Dim PLCNo As Integer
    Dim MagtrolNo As String
    
    PLCNo = 0
    If InStr(cmbPLCNo.Text, "8") > 0 Then
        PLCNo = 8
        MagtrolNo = "GPIB6"
    End If
    If InStr(cmbPLCNo.Text, "9") > 0 Then
        PLCNo = 9
        MagtrolNo = "GPIB5"
    End If
    
    For I = 0 To cmbPLCLoop.ListCount - 1                     'go through the combobox entries
        If InStr(cmbPLCLoop.List(I), PLCNo) > 0 Then   'see when we find the desired index number
            cmbPLCLoop.ListIndex = I                                              'if we do, set the combo box
            Exit For                                            'and we're done
        End If
        'cmbPLCLoop.ListIndex = -1                             'else, remove any pointer
    Next I
    
    For I = 0 To cmbMagtrol.ListCount - 1
        If InStr(cmbMagtrol.List(I), MagtrolNo) > 0 Then   'see when we find the desired index number
            cmbMagtrol.ListIndex = I                                              'if we do, set the combo box
            Exit For                                            'and we're done
        End If
    Next I
End Sub

Private Sub cmbVoltage_click()
    If Me.cmbVoltage.ListIndex = 0 Then
        Me.cmbFrequency.ListIndex = 2
    End If
End Sub
Private Sub cmbMagtrol_Click()
    Dim I As Integer
    Dim sSendStr As String
    Dim sGPIBName As String
    Dim MagtrolName As String

    I = cmbMagtrol.ItemData(cmbMagtrol.ListIndex)
    sGPIBName = "GPIB" & I
    MagtrolName = cmbMagtrol.List(cmbMagtrol.ListIndex)

    If I = 99 Then      'manual entry
        boMagtrolOperating = False
        EnableMagtrolFields
        Exit Sub
    Else
        boMagtrolOperating = True
    End If

    SetupMagtrols MagtrolName, I
  
End Sub


Private Sub cmbPLCLoop_Click()
    'Change the PLC that we're looking at

    Dim RetVal As String

    'manual data entry selection
    If cmbPLCLoop.ListIndex = cmbPLCLoop.ListCount - 1 Then 'no plc
        boPLCOperating = False
        EnablePLCFields
        If DeviceOpen = True Then
            RetVal = DisconnectPLC()
        End If
        Exit Sub
    End If

    If DeviceOpen = True Then
        RetVal = DisconnectPLC()
    End If

    RetVal = ConnectToPLC(cmbPLCLoop.ItemData(cmbPLCLoop.ListIndex))
    If RetVal <> 0 Then
        MsgBox ("Can't connect to PLC - " & Description(cmbPLCLoop.ListIndex))
        boPLCOperating = False
        EnablePLCFields
    Else
        boPLCOperating = True
        tDevice = cmbPLCLoop.ItemData(cmbPLCLoop.ListIndex)
        DisablePLCFields
    End If
End Sub

Private Sub cmbTestDate_Click()
    'select a test date to show

    Dim sName As String
    Dim sParam As String
    Dim I As Integer
    Dim j As Integer
    Dim k As Integer
    Dim bSk As Boolean
    Dim sBC As Single
    Dim NOK() As Long

    cmdModifyBalanceHoleData.Visible = False


    If Not boFoundTestSetup Then    'if we don't have any TestSetup data written
        boFoundTestData = False
        Exit Sub
    End If


    'select the testsetup data for the serial number
    qyTestSetup.ActiveConnection = cnPumpData
    qyTestSetup.CommandText = "SELECT * " & _
                  "From TempTestSetupData " & _
                  "Where (((TempTestSetupData.SerialNumber) = '" & txtSN.Text & "') AND TempTestSetupData.Date = #" & cmbTestDate.List(cmbTestDate.ListIndex) & "#) " & _
                  "ORDER BY TempTestSetupData.Date;"

    If rsTestSetup.State = adStateOpen Then
        rsTestSetup.Close
    End If

    With rsTestSetup     'open the recordset for the query
'        .Index = "FindData"
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .Open qyTestSetup
    End With

    'move to the selected date
    If Not rsTestSetup.BOF Then
        rsTestSetup.MoveFirst
    End If
'
    'show the correct combo box entries for this record
    'SetComboTestSetup cmbOrificeNumber, "OrificeNumber", "OrificeNumber", rsTestSetup
    SetComboTestSetup cmbTestSpec, "TestSpec", "TestSpecification", rsTestSetup
    SetComboTestSetup cmbLoopNumber, "LoopNumber", "LoopNumber", rsTestSetup
    SetComboTestSetup cmbSuctDia, "SuctDiam", "SuctionDiameter", rsTestSetup
    SetComboTestSetup cmbDischDia, "DischDiam", "DischargeDiameter", rsTestSetup
    SetComboTestSetup cmbTachID, "TachID", "TachID", rsTestSetup
    SetComboTestSetup cmbAnalyzerNo, "AnalyzerNo", "AnalyzerNo", rsTestSetup
    SetComboTestSetup cmbVoltage, "Voltage", "Voltage", rsTestSetup
    SetComboTestSetup cmbFrequency, "Frequency", "Frequency", rsTestSetup
    SetComboTestSetup cmbMounting, "Mounting", "Mounting", rsTestSetup
    SetComboTestSetup cmbPLCNo, "PLCNo", "PLCNo", rsTestSetup
    SetComboTestSetup cmbFlowMeter, "FlowMeterID", "PumpFlowMeter", rsTestSetup
    SetComboTestSetup cmbSuctionPressureTransducer, "SuctionID", "SuctionPressureTransducer", rsTestSetup
    SetComboTestSetup cmbDischargePressureTransducer, "DischID", "DischargePressureTransducer", rsTestSetup
    SetComboTestSetup cmbTemperatureTransducer, "TemperatureID", "TemperatureTransducer", rsTestSetup
    SetComboTestSetup cmbCirculationFlowMeter, "MagFlowID", "CirculationFlowMeter", rsTestSetup

    sName = "HDCor"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtHDCor.Text = sParam

    sName = "KWMult"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtKWMult.Text = sParam

    sName = "Who"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtWho.Text = sParam

    sName = "RMA"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtRMA.Text = sParam

    sName = "Remarks"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtTestSetupRemarks.Text = sParam

    sName = "VFDFrequency"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtVFDFreq.Text = sParam

    sName = "SuctionGageHeight"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = 0
    End If
    txtSuctHeight.Text = sParam

    sName = "DischargeGageHeight"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = 0
    End If
    txtDischHeight.Text = sParam

    sName = "EndPlay"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtEndPlay.Text = sParam

    sName = "GGAP"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtGGap.Text = sParam

    sName = "OtherMods"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = vbNullString
    End If
    txtOtherMods.Text = sParam

    If rsTestSetup.Fields("ImpFeathered") Then
        chkFeathered.value = 1
    Else
        chkFeathered.value = 0
    End If

    If Val(rsTestSetup.Fields("ImpTrimmed")) = 0 Then
        chkTrimmed.value = 0
        txtImpTrim.Visible = False
        txtImpTrim.Text = rsTestSetup.Fields("Imptrimmed")
    Else
        chkTrimmed.value = 1
        txtImpTrim.Visible = True
        txtImpTrim.Text = rsTestSetup.Fields("Imptrimmed")
    End If

    If Val(rsTestSetup.Fields("PumpDischOrifice")) = 0 Then
        chkOrifice.value = 0
        txtOrifice.Visible = False
    Else
        chkOrifice.value = 1
        txtOrifice.Visible = True
        txtOrifice.Text = rsTestSetup.Fields("PumpDischOrifice")
    End If

    If Val(rsTestSetup.Fields("CircFlowOrifice")) = 0 Then
        chkCircOrifice.value = 0
        txtCircOrifice.Visible = False
    Else
        chkCircOrifice.value = 1
        txtCircOrifice.Visible = True
        txtCircOrifice.Text = rsTestSetup.Fields("CircFlowOrifice")
    End If

    If (IsNull(rsTestSetup.Fields("NPSHFile"))) Or (LenB(rsTestSetup.Fields("NPSHFile")) = 0) Then
        chkNPSH.value = 0
        txtNPSHFile.Visible = False
    Else
        chkNPSH.value = 1
        txtNPSHFile.Visible = True
        txtNPSHFile.Text = rsTestSetup.Fields("NPSHFile")
    End If

    If (IsNull(rsTestSetup.Fields("PictureFile"))) Or (LenB(rsTestSetup.Fields("PictureFile")) = 0) Then
        chkPictures.value = 0
        txtPicturesFile.Visible = False
    Else
        chkPictures.value = 1
        txtPicturesFile.Visible = True
        txtPicturesFile.Text = rsTestSetup.Fields("PictureFile")
    End If

    If (IsNull(rsTestSetup.Fields("VibrationFile"))) Or (LenB(rsTestSetup.Fields("VibrationFile")) = 0) Then
        chkVibration.value = 0
        txtVibrationFile.Visible = False
    Else
        chkVibration.value = 1
        txtVibrationFile.Visible = True
        txtVibrationFile.Text = rsTestSetup.Fields("VibrationFile")
    End If


    'for TEMC Inspection Report
    sName = "InsulationMeggerVolts"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = 0
    End If
    txtTestAndInspection(0).Text = sParam

    sName = "InsulationMegOhms"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = 0
    End If
    txtTestAndInspection(1).Text = sParam

    sName = "DielectricVolts"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = 0
    End If
    txtTestAndInspection(2).Text = sParam

    sName = "DielectricTime"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = 0
    End If
    txtTestAndInspection(3).Text = sParam

    sName = "HydrostaticValue"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = 0
    End If
    txtTestAndInspection(4).Text = sParam

    sName = "HydrostaticTime"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = 0
    End If
    txtTestAndInspection(5).Text = sParam

    sName = "PneumaticValue"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = 0
    End If
    txtTestAndInspection(6).Text = sParam

    sName = "PneumaticTime"
    If rsTestSetup.Fields(sName).ActualSize <> 0 Then
        sParam = rsTestSetup.Fields(sName)
    Else
        sParam = 0
    End If
    txtTestAndInspection(7).Text = sParam

    For I = 0 To cmbTestAndInspection(0).ListCount - 1
        If cmbTestAndInspection(0).Text = rsTestSetup.Fields("HydrostaticUnits") Then
                cmbTestAndInspection(0).ListIndex = I
                Exit For
        End If
        cmbTestAndInspection(0).ListIndex = -1
    Next I


    For I = 0 To cmbTestAndInspection(1).ListCount - 1
        If cmbTestAndInspection(1).Text = rsTestSetup.Fields("PneumaticUnits") Then
                cmbTestAndInspection(1).ListIndex = I
                Exit For
        End If
        cmbTestAndInspection(1).ListIndex = -1
    Next I

    TestAndInspectionGood(0).value = Abs(rsTestSetup!insulationgood)
    TestAndInspectionGood(1).value = Abs(rsTestSetup!DielectricGood)
    TestAndInspectionGood(2).value = Abs(rsTestSetup!HydrostaticGood)
    TestAndInspectionGood(3).value = Abs(rsTestSetup!PneumaticGood)
    TestAndInspectionGood(4).value = Abs(rsTestSetup!GeneralAppearanceGood)
    TestAndInspectionGood(5).value = Abs(rsTestSetup!OutlineDimensionsGood)
    TestAndInspectionGood(6).value = Abs(rsTestSetup!MotorNoLoadTestGood)
    TestAndInspectionGood(7).value = Abs(rsTestSetup!MotorLockedRotorTestGood)
    TestAndInspectionGood(8).value = Abs(rsTestSetup!HydrostaticTestGood)
    TestAndInspectionGood(9).value = Abs(rsTestSetup!HydraulicTestGood)
    TestAndInspectionGood(10).value = Abs(rsTestSetup!NPSHTestGood)
    TestAndInspectionGood(11).value = Abs(rsTestSetup!CleanPurgeSealGood)
    TestAndInspectionGood(12).value = Abs(rsTestSetup!PaintCheckGood)
    TestAndInspectionGood(13).value = Abs(rsTestSetup!NameplateGood)
    TestAndInspectionGood(14).value = Abs(rsTestSetup!SupervisorApproval)

    GetBalanceHoleData frmPLCData.txtSN.Text, cmbTestDate.Text

     If rsBalanceHoles.RecordCount = 0 Then
        chkBalanceHoles.value = 0
        dgBalanceHoles.Visible = False
        boGotBalanceHoles = False
    Else
        boGotBalanceHoles = True
        ReDim NOK(rsBalanceHoles.RecordCount)
        rsBalanceHoles.MoveLast
        For I = 1 To rsBalanceHoles.RecordCount
            NOK(I) = 0
        Next I

        For j = 1 To rsBalanceHoles.RecordCount - 1
            rsBalanceHoles.MoveFirst
            rsBalanceHoles.Move rsBalanceHoles.RecordCount - j
            sBC = rsBalanceHoles.Fields("BoltCircle")
            bSk = False
            For k = 1 To rsBalanceHoles.RecordCount
                If NOK(k) = rsBalanceHoles.Fields(0) Then
                    bSk = True
                End If
            Next k
            If Not bSk Then
                For I = rsBalanceHoles.RecordCount - j To 1 Step -1
                    rsBalanceHoles.MovePrevious
                    If rsBalanceHoles.Fields("BoltCircle") = sBC Then
                        NOK(I) = rsBalanceHoles.Fields(0)
                    End If
                Next I
            End If
        Next j

        Dim sFilt As String
        sFilt = ""
        For I = 1 To rsBalanceHoles.RecordCount
            If NOK(I) <> 0 Then
                sFilt = sFilt & "(BalanceHoleID <> " & NOK(I) & ") AND "
'                sFilt = sFilt & "(" & rsBalanceHoles.Filter & " AND BalanceHoleID <> " & NOK(I) & ") AND "
            End If
        Next I

        If Len(sFilt) > 4 Then
            sFilt = Left(sFilt, Len(sFilt) - 4)
            rsBalanceHoles.Filter = sFilt
        End If

        chkBalanceHoles.value = 1
        dgBalanceHoles.Visible = True
    End If
'
    'set the test date filter for the test data
    rsTestData.Filter = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"

    If rsTestData.RecordCount = 0 Then
        boFoundTestData = False
        AddTestData
        EnableTestDataControls
        MsgBox "No Test Data Exists for this Serial Number"
    Else
        boFoundTestData = True
        DisableTestDataControls                         'if it's in the real database, don't allow changes here
    End If

    If Not boTestDateIsApproved Then    'data approved?
        EnableTestDataControls
    End If

    If rsTestSetup.Fields("Approved") = True Then
        DisableTestDataControls                         'if it's in the real database, don't allow changes here
        lblTestDateApproved.Visible = True
        MsgBox ("Found pump.  Data cannot be modified.")
        If boCanApprove Then
            cmdApproveTestDate.Caption = "Unapprove this Test Date"
        End If
    Else
        EnableTestDataControls                          'it's in the temp database, allow changes
        lblTestDateApproved.Visible = False
        If boPumpIsApproved = True Then
            MsgBox ("Found pump.  Pump data cannot be modified, but test setup data and test data can be modified.")
        Else
            MsgBox ("Found pump.  Pump data, test setup data, and test data can be modified.")
        End If
        If boCanApprove Then
            If rsPumpData.Fields("Approved") = True Then
                cmdApproveTestDate.Enabled = True
                cmdApproveTestDate.Caption = "Approve this Test Date"
            Else
                cmdApproveTestDate.Caption = "You Must Approve Pump First"
                cmdApproveTestDate.Enabled = False
            End If
        End If
    End If

    rsEff.MoveFirst
    rsTestData.MoveFirst

    For I = 1 To rsTestData.RecordCount
        DoEfficiencyCalcs
        rsEff.MoveNext
        rsTestData.MoveNext
    Next I

   ' fix the datagrid
   Set DataGrid1.DataSource = rsTestData
   Set DataGrid2.DataSource = rsEff

   Dim c As Column
   For Each c In DataGrid1.Columns
      Select Case c.DataField
      Case "TestDataID"     'Hide some columns
         c.Visible = False
      Case "SerialNumber"
         c.Visible = False
      Case "Date"
         c.Visible = False
      Case Else             ' Show all other columns.
         c.Visible = True
         c.Alignment = dbgRight
      End Select
    Next c

    For Each c In DataGrid2.Columns
        c.Alignment = dbgCenter
        c.Width = 750
        Select Case c.ColIndex
            Case 1
                c.Caption = "Flow"
                c.NumberFormat = "###0.00"
            Case 2
                c.Caption = "TDH"
                c.NumberFormat = "##0.00"
            Case 3
                c.Caption = "Input Pwr"
                c.NumberFormat = "##0.00"
                c.Width = 850
            Case 4
                c.Caption = "Voltage"
                c.NumberFormat = "##0.00"
            Case 5
                c.Caption = "Current"
                c.NumberFormat = "##0.00"
            Case 6
                c.Caption = "Overall Eff"
                c.NumberFormat = "##0.00"
                c.Width = 850
            Case 7
                c.Caption = "NPSHr"
                c.NumberFormat = "#0.00"
            Case Else
                c.Visible = False
        End Select
    Next c
        FixPointsToPlot

    txtUpDn1.Text = 1

'unlock the text boxes
    For I = 0 To 7
        txtTitle(I).Locked = False
    Next I

    For I = 20 To 27
        txtTitle(I).Locked = False
    Next I

'look for titles for TCs and AIs
    Dim qy As New ADODB.Command
    Dim rs As New ADODB.Recordset

    qy.ActiveConnection = cnPumpData

    'see if we have an entry in the table
    qy.CommandText = "SELECT * FROM AITitles " & _
                      "WHERE (((AITitles.SerialNo)= '" & txtSN.Text & "') " & _
                      "AND ((AITitles.Date)= #" & cmbTestDate.Text & "#)); "

    With rs     'open the recordset for the query
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open qy
    End With

    If Not (rs.BOF = True And rs.EOF = True) Then   'update titles
        rs.MoveFirst
        Do While Not rs.EOF
            txtTitle(rs.Fields("Channel")).Text = rs.Fields("Title")
            rs.MoveNext
        Loop
    End If

    rs.Close
    Set rs = Nothing
    Set qy = Nothing
End Sub

Private Sub cmdAddNewBalanceHoles_Click()
    Dim strInput As String
    Dim I As Integer
    Dim sNumber As Integer
    Dim sDia As Single
    Dim sBC As Single

    'get the data for the balance holes
    strInput = InputBox("Enter Number of Holes")
    If strInput <> "" Then
        sNumber = CInt(strInput)
    Else
        GoTo CancelPressed
    End If

    strInput = InputBox("Enter Decimal Value of Hole Diameter or Slot (For Example, 0.675) ")
    If strInput <> "" Then
        If UCase(strInput) = "SLOT" Then
            strInput = 99
        End If
        sDia = CSng(strInput)
    Else
        GoTo CancelPressed
    End If

    strInput = InputBox("Enter Decimal Value of Bolt Circle or Unknown (For Example, 4.525)")
    If strInput <> "" Then
        If UCase(strInput) = "UNKNOWN" Then
            strInput = 99
        End If
        sBC = CSng(strInput)
    Else
        GoTo CancelPressed
    End If

    GetBalanceHoleData frmPLCData.txtSN.Text, cmbTestDate.Text

    rsBalanceHoles.AddNew
    rsBalanceHoles!SerialNo = txtSN.Text
    rsBalanceHoles!Date = cmbTestDate.Text
    rsBalanceHoles!Number = sNumber
    rsBalanceHoles!diameter = sDia
    rsBalanceHoles!boltcircle = sBC

    rsBalanceHoles.Update

    GetBalanceHoleData txtSN.Text, cmbTestDate.Text
    rsBalanceHoles.MoveLast
    dgBalanceHoles.Refresh
    chkBalanceHoles.value = 1

    Exit Sub

CancelPressed:
    MsgBox "No New Balance Hole Data Entered", vbOKOnly
End Sub

Private Sub cmdAddNewTestDate_Click()
    'add a new test date/time
    Dim I As Integer
    
    chkFeathered.value = chkSuperMarketFeathered.value
    
    For I = 1 To cmbTestDate.ListCount      'see if we already have today's date entered
        If cmbTestDate.List(I) = Date Then
            MsgBox "There is already an entry for today.  You can only have one entry for each Serial Number and a given date.  You may want to modify the Serial Number.", vbOKOnly
            Exit Sub
        End If
    Next I

    'we didn't find today's date entered, allow data entry
    boFoundTestSetup = False

    SetFrequencyCombo

    EnableTestSetupDataControls
    Pressed = False
    cmdEnterTestSetupData_Click
    cmdAddNewBalanceHoles.Visible = True
    txtWho.Text = LogInInitials
    MsgBox "New Test Date Added - " & cmbTestDate.List(cmbTestDate.ListCount - 1), vbOKOnly, "Added New Test Date"
End Sub

Private Sub cmdApprovePump_Click()
    'allow the pump data to be approved
    rsPumpData.Fields("Approved") = Not rsPumpData.Fields("Approved")
    rsPumpData.Update
    rsPumpData.Requery
    lblPumpApproved.Visible = rsPumpData.Fields("Approved")
    If rsPumpData.Fields("Approved") = True Then
        cmdApprovePump.Caption = "Unapprove This Pump"
        cmdApproveTestDate.Enabled = True
        If rsTestSetup.Fields("Approved") = True Then
            cmdApproveTestDate.Caption = "Unapprove This Test Date"
        Else
            cmdApproveTestDate.Caption = "Approve This Test Date"
        End If
    Else
        cmdApprovePump.Caption = "Approve This Pump"
        cmdApproveTestDate.Caption = "You Must Approve Pump First"
        cmdApproveTestDate.Enabled = False
    End If
End Sub

Private Sub cmdApproveTestDate_Click()
    'allow the test setup data to be approved
    rsTestSetup.Fields("Approved") = Not rsTestSetup.Fields("Approved")
    rsTestSetup.Update
    rsTestSetup.Requery
    lblTestDateApproved.Visible = rsTestSetup.Fields("Approved")
    If rsTestSetup.Fields("Approved") = True Then
        cmdApproveTestDate.Caption = "Unapprove This Test Date"
    Else
        cmdApproveTestDate.Caption = "Approve This Test Date"
    End If
End Sub

Private Sub cmdCalibrate_Click()
    Dim ans As Integer
    Dim I As Integer

    ans = MsgBox("You have selected to calibrate the software.  Do you want to continue?", vbYesNo, "Calibrate Software")
    If ans = vbNo Then
        Calibrating = False
        Exit Sub
    Else
        CalibrateSoftware
    End If
End Sub

Private Sub cmdClearPumpData_Click()
    BlankData
End Sub

Private Sub cmdDeletePump_Click()
    'delete this pump
    Dim Answer As Integer
    Answer = MsgBox("You are about to delete the following record: S/N = " & rsPumpData.Fields("SerialNumber") & "!  Do you want to continue?", vbCritical Or vbYesNo, "Ready to Delete")
    If Answer = vbYes Then
        rsPumpData.Delete
        rsPumpData.Update
        cmdFindPump_Click
    End If
End Sub

Private Sub cmdDeleteTestDate_Click()
    'delete this test date
    Dim Answer As Integer
    Answer = MsgBox("You are about to delete the following record: S/N = " & rsTestData.Fields("SerialNumber") & " and Test Date = " & rsTestSetup.Fields("Date") & "!  Do you want to continue?", vbCritical Or vbYesNo, "Ready to Delete")
    If Answer = vbYes Then
        rsTestSetup.Delete
        rsTestSetup.Update
        cmdFindPump_Click
    End If
End Sub

Private Sub cmdEnterPumpData_Click()
    'store the data on the screen to the pump (pumpdata)
    Dim d As Integer
    Dim sSearch As String
    Dim ans As Integer
    Dim boWriteDataWritten As Boolean


    'check for a serial number
    If LenB(txtSN.Text) = 0 Then
        MsgBox "You must have a Serial Number to enter data.  Data has not been saved."
        Exit Sub
    End If

    'check to make sure most entries are filled in
    If LenB(txtModelNo.Text) = 0 And optMfr(0).value = True Then
        MsgBox "You need to enter a MODEL NO before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If
    If LenB(txtSalesOrderNumber.Text) = 0 Then
        If InStr(1, txtSN.Text, "-") <> 0 Then
            txtSalesOrderNumber.Text = Mid$(txtSN.Text, 1, InStr(1, txtSN.Text, "-") - 1)
        End If
    End If
    If LenB(txtSalesOrderNumber.Text) = 0 Then
        MsgBox "You need to enter a SALES ORDER NUMBER before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbMotor.ListIndex = -1 And optMfr(0).value = True Then
        MsgBox "You need to pick a MOTOR before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbStatorFill.ListIndex = -1 And optMfr(0).value = True Then    'set default
        cmbStatorFill.ListIndex = 0
    End If

    If cmbModel.ListIndex = -1 And optMfr(0).value = True Then
        MsgBox "You need to pick a MODEL before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbModelGroup.ListIndex = -1 And optMfr(0).value = True Then
        MsgBox "You need to pick a MODEL GROUP before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If


    If cmbDesignPressure.ListIndex = -1 And optMfr(0).value = True Then
        MsgBox "You need to pick a DESIGN PRESSURE before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbCirculationPath.ListIndex = -1 And optMfr(0).value = True Then
        MsgBox "You need to pick a CIRCULATION PATH before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbRPM.ListIndex = -1 And optMfr(0).value = True Then
        MsgBox "You need to pick an RPM before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

'check TEMC dropdowns

    If cmbTEMCAdapter.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC ADAPTER before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCAdditions.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick TEMC ADDITIONS before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCCirculation.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC CIRCULATION before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCDesignPressure.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC DESIGN PRESSURE before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCDivisionType.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC DIVISION TYPE before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCImpellerType.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC IMPELLER TYPE before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCInsulation.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC INSULATION TYPE before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCJacketGasket.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC JACKET GASKET before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCMaterials.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick TEMC MATERIALS before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCModel.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC MODEL before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCNominalImpSize.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC NOMINAL IMPELLER SIZE before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCNominalDischargeSize.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC NOMINAL DISCHARGE SIZE before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCNominalSuctionSize.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC NOMINAL SUCTION SIZE before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCOtherMotor.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC OTHER MOTOR before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCPumpStages.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick TEMC NUMBER OF PUMP STAGES before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCTRG.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC TRG before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If cmbTEMCVoltage.ListIndex = -1 And optMfr(0).value = False Then
        MsgBox "You need to pick a TEMC VOLTAGE before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If

    If LenB(txtTEMCFrameNumber.Text) = 0 And optMfr(0).value = False Then
        MsgBox "You need to enter a TEMC FRAME NUMBER before saving data.  Data has not been saved.", vbOKOnly
        Exit Sub
    End If


    If Not boFoundPump Then     'if we havent found a pump in the database, add it
        rsPumpData.AddNew
        boWriteDataWritten = False
    Else    'else, find the entry
        sSearch = "Serialnumber = '" & frmPLCData.txtSN.Text & "'"
        rsPumpData.MoveFirst
        rsPumpData.Find sSearch, , adSearchForward, 1
        boWriteDataWritten = True
    End If

    If Not IsNull(rsPumpData!DataWritten) Or rsPumpData!DataWritten = True Then
        ans = MsgBox("You have already entered data for this pump.  Do you want to overwrite the data?", vbDefaultButton2 + vbYesNo, "Overwrite Data?")
        If ans = vbNo Then
            rsPumpData!DataWritten = True
            rsPumpData.Update   'update datawritten
            Exit Sub
        End If
    End If

    rsPumpData!SerialNumber = frmPLCData.txtSN.Text
    rsPumpData!ModelNumber = frmPLCData.txtModelNo.Text
    rsPumpData!SalesOrderNumber = frmPLCData.txtSalesOrderNumber.Text
    rsPumpData!ShipToCustomer = frmPLCData.txtShpNo.Text
    rsPumpData!BillToCustomer = frmPLCData.txtBilNo.Text
    rsPumpData!ApplicationFluid = frmPLCData.txtLiquid
    rsPumpData!NPSHFile = frmPLCData.txtNPSHFileLocation.Text
    rsPumpData!RVSPartNo = frmPLCData.txtRVSPartNo.Text
    rsPumpData!CustPN = frmPLCData.txtXPartNum.Text
    rsPumpData!CustPO = frmPLCData.txtCustPONum.Text
    
    If Len(frmPLCData.txtViscosity) <> 0 Then
        rsPumpData!ApplicationViscosity = frmPLCData.txtViscosity
    End If
    
    If frmPLCData.chkSuperMarketFeathered.value = Checked Then
        rsPumpData!Field1 = "Feathered"
    Else
        rsPumpData!Field1 = ""
    End If
    
    If LenB(txtSpGr.Text) <> 0 Then
        If Not IsNumeric(frmPLCData.txtSpGr.Text) Then
            MsgBox "Specific Gravity must be a number."
            Exit Sub
        End If
        rsPumpData!SpGr = frmPLCData.txtSpGr.Text
    End If
    If LenB(txtImpellerDia.Text) <> 0 Then
        If Not IsNumeric(frmPLCData.txtImpellerDia.Text) Then
            MsgBox "Impeller Diameter must be a number."
            Exit Sub
        End If
        rsPumpData!impellerdia = frmPLCData.txtImpellerDia.Text
    End If
    If LenB(txtDesignFlow.Text) <> 0 Then
        rsPumpData!designflow = frmPLCData.txtDesignFlow.Text
    End If
    If LenB(txtDesignTDH.Text) <> 0 Then
        rsPumpData!designtdh = frmPLCData.txtDesignTDH.Text
    End If
    If LenB(txtRemarks.Text) <> 0 Then
        rsPumpData!Remarks = txtRemarks.Text
    End If

    If optMfr(0).value = True Then
        d = cmbMotor.ItemData(cmbMotor.ListIndex)
        rsPumpData!Motor = d
        d = cmbStatorFill.ItemData(cmbStatorFill.ListIndex)
        rsPumpData!StatorFill = d
         d = cmbDesignPressure.ItemData(cmbDesignPressure.ListIndex)
        rsPumpData!DesignPressure = d
        d = cmbCirculationPath.ItemData(cmbCirculationPath.ListIndex)
        rsPumpData!CirculationPath = d
        d = cmbRPM.ItemData(cmbRPM.ListIndex)
        rsPumpData!RPM = d
        d = cmbModel.ItemData(cmbModel.ListIndex)
        rsPumpData!Model = d
        d = cmbModelGroup.ItemData(cmbModelGroup.ListIndex)
        rsPumpData!ModelGroup = d
    End If
'   TEMC fields
    If optMfr(0).value = False Then
        d = cmbTEMCAdapter.ItemData(cmbTEMCAdapter.ListIndex)
        rsPumpData!TEMCAdapter = d

        d = cmbTEMCAdditions.ItemData(cmbTEMCAdditions.ListIndex)
        rsPumpData!TEMCAdditions = d

        d = cmbTEMCCirculation.ItemData(cmbTEMCCirculation.ListIndex)
        rsPumpData!TEMCcirculation = d

        d = cmbTEMCDesignPressure.ItemData(cmbTEMCDesignPressure.ListIndex)
        rsPumpData!TEMCDesignpressure = d

        d = cmbTEMCDivisionType.ItemData(cmbTEMCDivisionType.ListIndex)
        rsPumpData!TEMCDivisionType = d

        d = cmbTEMCImpellerType.ItemData(cmbTEMCImpellerType.ListIndex)
        rsPumpData!TEMCImpellerType = d

        d = cmbTEMCInsulation.ItemData(cmbTEMCInsulation.ListIndex)
        rsPumpData!TEMCInsulation = d

        d = cmbTEMCJacketGasket.ItemData(cmbTEMCJacketGasket.ListIndex)
        rsPumpData!TEMCJacketGasket = d

        d = cmbTEMCMaterials.ItemData(cmbTEMCMaterials.ListIndex)
        rsPumpData!TEMCMaterials = d

        d = cmbTEMCModel.ItemData(cmbTEMCModel.ListIndex)
        rsPumpData!TEMCModel = d

        d = cmbTEMCNominalImpSize.ItemData(cmbTEMCNominalImpSize.ListIndex)
        rsPumpData!TEMCNominalImpSize = d

        d = cmbTEMCNominalDischargeSize.ItemData(cmbTEMCNominalDischargeSize.ListIndex)
        rsPumpData!TEMCNominalDischargeSize = d

        d = cmbTEMCNominalSuctionSize.ItemData(cmbTEMCNominalSuctionSize.ListIndex)
        rsPumpData!TEMCNominalSuctionSize = d

        d = cmbTEMCOtherMotor.ItemData(cmbTEMCOtherMotor.ListIndex)
        rsPumpData!TEMCOtherMotor = d

        d = cmbTEMCPumpStages.ItemData(cmbTEMCPumpStages.ListIndex)
        rsPumpData!TEMCPumpStages = d

        d = cmbTEMCTRG.ItemData(cmbTEMCTRG.ListIndex)
        rsPumpData!TEMCTRG = d

        d = cmbTEMCVoltage.ItemData(cmbTEMCVoltage.ListIndex)
        rsPumpData!TEMCVoltage = d

        If LenB(txtTEMCFrameNumber.Text) <> 0 Then
            rsPumpData!TEMCFrameNumber = frmPLCData.txtTEMCFrameNumber.Text
        End If
    End If

    rsPumpData!ChempumpPump = optMfr(0).value

    rsPumpData!Approved = False

'added from TEMC Inspection Report
    If Len(txtJobNum.Text) <> 0 Then
        rsPumpData!JobNumber = txtJobNum.Text
    End If

    If Len(txtNoPhases.Text) <> 0 Then
        rsPumpData!Phases = txtNoPhases.Text
    End If

    If Len(txtExpClass.Text) <> 0 Then
        rsPumpData!ExpClass = txtExpClass.Text
    End If

    If Len(txtThermalClass.Text) <> 0 Then
        rsPumpData!ThermalClass = txtThermalClass.Text
    End If

    rsPumpData!NPSHr = Val(txtNPSHr.Text)
    rsPumpData!LiquidTemperature = Val(txtLiquidTemperature.Text)
    rsPumpData!RatedInputPower = Val(txtRatedInputPower.Text)
    rsPumpData!FLCurrent = Val(txtAmps.Text)





    If boWriteDataWritten Then
        rsPumpData!DataWritten = True
    Else
        rsPumpData!DataWritten = False
    End If

    'write the data into the database
    rsPumpData.Update
    boFoundPump = True

    'enter a new test date if it's a new entry
    If Not boWriteDataWritten Then


        cmdAddNewTestDate_Click
    End If
End Sub
Private Sub cmdEnterTestData_Click()
    ' save the data on the screen to test data at the selected run
    Dim sSearch As String
    Dim ans As Integer

    'if we didn't find the test setup, can't enter test data
    If Not boFoundTestSetup Then
        MsgBox "You must enter Test Setup Data before entering the Test Data"
        Exit Sub
    End If

    'if we don't find data in the test database, add records
    If boFoundTestData = False Then     'add 8 records for 8 tests
        AddTestData
        rsTestData.MoveFirst
    Else        'find the data in the database
        sSearch = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"
        rsTestData.MoveFirst
        rsTestData.Filter = sSearch
    End If

    'find the desired record from the form
    rsTestData.MoveFirst
    rsTestData.Move UpDown1.value - 1

    If rsTestData!DataWritten = True Then
        ans = MsgBox("You have already entered data for this test.  Do you want to overwrite the data?", vbYesNo + vbDefaultButton2, "Data Already Entered")
        If ans = vbNo Then
            Exit Sub
        End If
    End If

    rsEff.MoveFirst
    rsEff.Move UpDown1.value - 1

    If LenB(txtV1.Text) <> 0 Then
        rsTestData!VoltageA = Val(txtV1.Text)
    End If

    If LenB(txtV2.Text) <> 0 Then
        rsTestData!VoltageB = Val(txtV2.Text)
    End If

    If LenB(txtV3.Text) <> 0 Then
        rsTestData!VoltageC = Val(txtV3.Text)
    End If

    If LenB(txtI1.Text) <> 0 Then
        rsTestData!CurrentA = Val(txtI1.Text)
    End If

    If LenB(txtI2.Text) <> 0 Then
        rsTestData!CurrentB = Val(txtI2.Text)
    End If

    If LenB(txtI3.Text) <> 0 Then
        rsTestData!CurrentC = Val(txtI3.Text)
    End If

    If LenB(txtP1.Text) <> 0 Then
        rsTestData!PowerA = Val(txtP1.Text)
    End If

    If LenB(txtP2.Text) <> 0 Then
        rsTestData!PowerB = Val(txtP2.Text)
    End If

    If LenB(txtP3.Text) <> 0 Then
        rsTestData!PowerC = Val(txtP3.Text)
    End If

    If LenB(txtKW.Text) <> 0 Then
        rsTestData!TotalPower = Val(txtKW.Text)
    End If

    rsTestData!Flow = Val(txtFlowDisplay.Text)
    rsTestData!DischargePressure = Val(txtDischargeDisplay.Text)
    rsTestData!SuctionPressure = Val(txtSuctionDisplay.Text)
    rsTestData!TemperatureSuction = Val(txtTemperatureDisplay.Text)

    rsTestData!TC1 = Val(txtTC1Display.Text)
    rsTestData!TC2 = Val(txtTC2Display.Text)
    rsTestData!TC3 = Val(txtTC3Display.Text)
    rsTestData!TC4 = Val(txtTC4Display.Text)

    rsTestData!CircFlow = Val(txtAI1Display.Text)
    rsTestData!RBHTemp = Val(txtAI2Display.Text)
    rsTestData!RBHPress = Val(txtAI3Display.Text)
    rsTestData!AI4 = Val(txtAI4Display.Text)

    rsTestData!ValvePosition = Val(txtValvePosition.Text)
    rsTestData!SetPoint = Val(txtSetPoint.Text)

    If LenB(txtThrustBal.Text) <> 0 Then
        rsTestData!ThrustBalance = txtThrustBal.Text
    End If

    If LenB(txtVibAx.Text) <> 0 Then
        rsTestData!VibrationX = txtVibAx.Text
    End If

    If LenB(txtVibRad.Text) <> 0 Then
        rsTestData!VibrationY = txtVibRad.Text
    End If

    If LenB(txtTEMCTRGReading.Text) <> 0 Then
        rsTestData!TEMCTRG = txtTEMCTRGReading.Text
    Else
        rsTestData!TEMCTRG = 0
    End If

    If LenB(txtRPM.Text) <> 0 Then
        rsTestData!RPM = txtRPM.Text
    End If

    If LenB(txtTestRemarks.Text) <> 0 Then
        rsTestData!Remarks = txtTestRemarks.Text
    Else
        rsTestData!Remarks = " "
    End If

    If LenB(txtTEMCTRGReading.Text) <> 0 Then
        rsTestData!TEMCTRG = txtTEMCTRGReading.Text
    End If

    If LenB(txtTEMCFrontThrust.Text) <> 0 Then
        rsTestData!TEMCFrontThrust = txtTEMCFrontThrust.Text
    End If

    If LenB(txtTEMCRearThrust.Text) <> 0 Then
        rsTestData!TEMCRearThrust = txtTEMCRearThrust.Text
    End If

    If LenB(txtTEMCMomentArm.Text) <> 0 Then
        rsTestData!TEMCMomentArm = txtTEMCMomentArm.Text
    End If

    If LenB(txtTEMCThrustRigPressure.Text) <> 0 Then
        rsTestData!TEMCThrustRigPressure = txtTEMCThrustRigPressure.Text
    End If

    If LenB(txtTEMCViscosity.Text) <> 0 Then
        rsTestData!TEMCViscosity = txtTEMCViscosity.Text
    End If

    If LenB(txtNPSHa.Text) <> 0 Then
        rsTestData!NPSHa = txtNPSHa.Text
    End If

    rsTestData!Approved = False

    rsTestData!DataWritten = True

    'update the database
    rsTestData.Update

    DoEfficiencyCalcs
    rsEff.Update

    'update the form
    DataGrid1.Refresh
    DataGrid2.Refresh

    FixPointsToPlot

End Sub
Private Sub cmdEnterTestSetupData_Click()
    'save the data on the screen to testsetupdata
    Dim I As Integer
    Dim d As Integer
    Dim sSearch As String
    Dim ans As Integer
    Dim boWriteDataWritten As Boolean

    'check for a serial number
    If LenB(txtSN.Text) = 0 Then
        MsgBox "You must have a Serial Number to enter data.", vbOKOnly + vbExclamation, "Cannot Enter Data"
        Exit Sub
    End If
    
    If Pressed = True Then
        If Me.cmbDischDia.ListIndex = -1 Or Me.cmbSuctDia.ListIndex = -1 Or Val(Me.txtSuctHeight.Text) = 0 Or Val(Me.txtDischHeight.Text) = 0 Then
            MsgBox "You must have Discharge Diameter AND Suction Diameter AND Suction Height AND Discharge Height entered to enter data.", vbOKOnly + vbExclamation, "Cannot Enter Data"
            Exit Sub
        End If
    End If

    Pressed = True
    If Not boFoundTestSetup Then    'if we didn't find any test setup, add a record
        rsTestSetup.AddNew
        cmbTestDate.AddItem Now
        cmbTestDate.ListIndex = cmbTestDate.NewIndex
        cmdAddNewBalanceHoles.Visible = True
        boFoundTestSetup = True
        boWriteDataWritten = False
        rsTestSetup!DataWritten = False
    Else    'find the record and display
        sSearch = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"
        rsTestSetup.MoveFirst
        rsTestSetup.Filter = sSearch
        If Not boCanApprove Then
'            cmdAddNewBalanceHoles.Visible = False
        End If
        boWriteDataWritten = True
    End If

    If rsTestSetup!DataWritten = True Then
        ans = MsgBox("Data has already been entered for this test date.  Do you want to overwrite it?", vbYesNo + vbDefaultButton2, "Data Exists")
        If ans = vbNo Then
            Exit Sub
        End If
    End If

    rsTestSetup!SerialNumber = txtSN
    rsTestSetup!Date = cmbTestDate.List(cmbTestDate.ListIndex)

    I = cmbFlowMeter.ListIndex
    If I = -1 Then
        d = 1
        rsTestSetup!FlowMeterID = d
    Else
        d = cmbLoopNumber.ItemData(I)
        rsTestSetup!FlowMeterID = d
    End If

    I = cmbSuctionPressureTransducer.ListIndex
    If I = -1 Then
        d = 1
        rsTestSetup!suctionid = d
    Else
        d = cmbLoopNumber.ItemData(I)
        rsTestSetup!suctionid = d
    End If

    I = cmbDischargePressureTransducer.ListIndex
    If I = -1 Then
        d = 1
        rsTestSetup!dischid = d
    Else
        d = cmbLoopNumber.ItemData(I)
        rsTestSetup!dischid = d
    End If

    I = cmbTemperatureTransducer.ListIndex
    If I = -1 Then
        d = 1
        rsTestSetup!temperatureid = d
    Else
        d = cmbLoopNumber.ItemData(I)
        rsTestSetup!temperatureid = d
    End If

    I = Me.cmbCirculationFlowMeter.ListIndex
    If I = -1 Or I > 4 Then
        d = 5
        rsTestSetup!MagFlowID = d
    Else
        d = cmbLoopNumber.ItemData(I) + 4
        rsTestSetup!MagFlowID = d
    End If


    If LenB(txtHDCor.Text) <> 0 Then
        rsTestSetup!HDCor = txtHDCor
    Else
        rsTestSetup!HDCor = 0
    End If
    If LenB(txtKWMult.Text) <> 0 Then
        rsTestSetup!kwmult = txtKWMult
    Else
        rsTestSetup!kwmult = 1
    End If
    If LenB(txtWho.Text) <> 0 Then
        rsTestSetup!who = txtWho
    Else
        rsTestSetup!who = vbNullString
    End If
    If LenB(txtRMA.Text) <> 0 Then
        rsTestSetup!RMA = txtRMA
    Else
        rsTestSetup!RMA = vbNullString
    End If
    If LenB(frmPLCData.txtDischHeight) <> 0 Then
        rsTestSetup!DischargeGageHeight = Val(txtDischHeight)
    Else
        rsTestSetup!DischargeGageHeight = 0
    End If
    If LenB(frmPLCData.txtSuctHeight) <> 0 Then
        rsTestSetup!SuctionGageHeight = Val(txtSuctHeight)
    Else
        rsTestSetup!SuctionGageHeight = 0
    End If
    If LenB(frmPLCData.txtTestSetupRemarks.Text) <> 0 Then
        rsTestSetup!Remarks = txtTestSetupRemarks.Text
    Else
        rsTestSetup!Remarks = vbNullString
    End If
    If LenB(frmPLCData.txtVFDFreq.Text) <> 0 Then
        rsTestSetup!VFDFrequency = txtVFDFreq.Text
    Else
        rsTestSetup!VFDFrequency = 0
    End If

    I = cmbOrificeNumber.ListIndex
    If I = -1 Then
        d = 18      'entry for None
    Else
        d = cmbOrificeNumber.ItemData(I)
    End If
    rsTestSetup!orificenumber = d

    If LenB(txtEndPlay.Text) <> 0 Then
        rsTestSetup!Endplay = Val(frmPLCData.txtEndPlay.Text)
    Else
        rsTestSetup!Endplay = 0
    End If

    If LenB(txtGGap.Text) <> 0 Then
        rsTestSetup!GGAP = Val(frmPLCData.txtGGap.Text)
    Else
        rsTestSetup!GGAP = 0
    End If

    If LenB(txtOtherMods.Text) <> 0 Then
        rsTestSetup!OtherMods = txtOtherMods.Text
    Else
        rsTestSetup!OtherMods = vbNullString
    End If

    rsTestSetup!Approved = False

    I = cmbLoopNumber.ListIndex
    If I = -1 Then
        d = 1
        rsTestSetup!loopnumber = d
    Else
        d = cmbLoopNumber.ItemData(I)
        rsTestSetup!loopnumber = d
    End If

    I = cmbSuctDia.ListIndex
    If I = -1 Then
        d = -1
    Else
        d = cmbSuctDia.ItemData(I)
        rsTestSetup!SuctDiam = d
    End If

    I = cmbDischDia.ListIndex
    If I = -1 Then
        d = -1
    Else
        d = cmbDischDia.ItemData(I)
        rsTestSetup!DischDiam = d
    End If

    I = cmbTachID.ListIndex
    If I = -1 Then
        d = 1
        rsTestSetup!tachid = d
    Else
        d = cmbTachID.ItemData(I)
        rsTestSetup!tachid = d
    End If

    I = cmbAnalyzerNo.ListIndex
    If I = -1 Then
        d = 1
    Else
        d = cmbAnalyzerNo.ItemData(I)
    End If
    rsTestSetup!analyzerno = d

    I = cmbTestSpec.ListIndex
    If I = -1 Then
        d = 1
    Else
        d = cmbTestSpec.ItemData(I)
    End If
    rsTestSetup!testspec = d

    I = cmbVoltage.ListIndex
    If I = -1 Then
        d = 1
    Else
        d = cmbVoltage.ItemData(I)
    End If
    rsTestSetup!Voltage = d

    I = cmbFrequency.ListIndex
    If I = -1 Then
        d = 1
    Else
        d = cmbFrequency.ItemData(I)
    End If
    rsTestSetup!Frequency = d

    I = cmbMounting.ListIndex
    If I = -1 Then
        d = 1
    Else
        d = cmbMounting.ItemData(I)
    End If
    rsTestSetup!Mounting = d

    I = cmbPLCNo.ListIndex
    If I = -1 Then
        d = 8
    Else
        d = cmbPLCNo.ItemData(I)
    End If
    rsTestSetup!PLCNo = d
 
    rsTestSetup!ImpFeathered = chkFeathered.value

    If chkTrimmed.value = 1 Then
        rsTestSetup!ImpTrimmed = Val(txtImpTrim)
    Else
        rsTestSetup!ImpTrimmed = 0
    End If
    chkTrimmed_Click

    If chkOrifice.value = 1 Then
        rsTestSetup!PumpDischOrifice = Val(txtOrifice)
    Else
        rsTestSetup!PumpDischOrifice = 0
    End If
    chkOrifice_Click

    If chkCircOrifice.value = 1 Then
        rsTestSetup!CircFlowOrifice = Val(txtCircOrifice)
    Else
        rsTestSetup!CircFlowOrifice = 0
    End If
    chkCircOrifice_Click

    chkBalanceHoles_Click

    If chkNPSH.value = 1 Then
        txtNPSHFile.Visible = True
        rsTestSetup!NPSHFile = txtNPSHFile
    Else
        rsTestSetup!NPSHFile = vbNullString
        txtNPSHFile.Visible = False
    End If

    If chkPictures.value = 1 Then
        txtPicturesFile.Visible = True
        rsTestSetup!PictureFile = txtPicturesFile
    Else
        rsTestSetup!PictureFile = vbNullString
        txtPicturesFile.Visible = False
    End If

    If chkVibration.value = 1 Then
        txtVibrationFile.Visible = True
        rsTestSetup!VibrationFile = txtVibrationFile
    Else
        rsTestSetup!VibrationFile = vbNullString
        txtVibrationFile.Visible = False
    End If

    If boWriteDataWritten Then
        rsTestSetup!DataWritten = True
    Else
        rsTestSetup!DataWritten = False
    End If

    'for TEMC Inspection Report
    If LenB(frmPLCData.txtTestAndInspection(0).Text) <> 0 Then
        rsTestSetup!InsulationMeggerVolts = frmPLCData.txtTestAndInspection(0).Text
    Else
        rsTestSetup!InsulationMeggerVolts = ""
    End If

    If LenB(frmPLCData.txtTestAndInspection(1).Text) <> 0 Then
        rsTestSetup!InsulationMegOhms = frmPLCData.txtTestAndInspection(1).Text
    Else
        rsTestSetup!InsulationMegOhms = ""
    End If

    If LenB(frmPLCData.txtTestAndInspection(2).Text) <> 0 Then
        rsTestSetup!DielectricVolts = frmPLCData.txtTestAndInspection(2).Text
    Else
        rsTestSetup!DielectricVolts = ""
    End If

    If LenB(frmPLCData.txtTestAndInspection(3).Text) <> 0 Then
        rsTestSetup!DielectricTime = frmPLCData.txtTestAndInspection(3).Text
    Else
        rsTestSetup!DielectricTime = ""
    End If

    If LenB(frmPLCData.txtTestAndInspection(4).Text) <> 0 Then
        rsTestSetup!HydrostaticValue = frmPLCData.txtTestAndInspection(4).Text
    Else
        rsTestSetup!HydrostaticValue = ""
    End If

    If LenB(frmPLCData.txtTestAndInspection(5).Text) <> 0 Then
        rsTestSetup!HydrostaticTime = frmPLCData.txtTestAndInspection(5).Text
    Else
        rsTestSetup!HydrostaticTime = ""
    End If

    If LenB(frmPLCData.txtTestAndInspection(6).Text) <> 0 Then
        rsTestSetup!PneumaticValue = frmPLCData.txtTestAndInspection(6).Text
    Else
        rsTestSetup!PneumaticValue = ""
    End If

    If LenB(frmPLCData.txtTestAndInspection(7).Text) <> 0 Then
        rsTestSetup!PneumaticTime = frmPLCData.txtTestAndInspection(7).Text
    Else
        rsTestSetup!PneumaticTime = ""
    End If

    I = cmbTestAndInspection(0).ListIndex
    If I = -1 Then
        rsTestSetup!HydrostaticUnits = ""
    Else
        rsTestSetup!HydrostaticUnits = cmbTestAndInspection(0).Text
    End If


    I = cmbTestAndInspection(1).ListIndex
    If I = -1 Then
        rsTestSetup!PneumaticUnits = ""
    Else
        rsTestSetup!PneumaticUnits = cmbTestAndInspection(1).Text
    End If

    'use abs to convert from 1 and 0 to boolean
    rsTestSetup!insulationgood = Abs(TestAndInspectionGood(0).value)
    rsTestSetup!DielectricGood = Abs(TestAndInspectionGood(1).value)
    rsTestSetup!HydrostaticGood = Abs(TestAndInspectionGood(2).value)
    rsTestSetup!PneumaticGood = Abs(TestAndInspectionGood(3).value)
    rsTestSetup!GeneralAppearanceGood = Abs(TestAndInspectionGood(4).value)
    rsTestSetup!OutlineDimensionsGood = Abs(TestAndInspectionGood(5).value)
    rsTestSetup!MotorNoLoadTestGood = Abs(TestAndInspectionGood(6).value)
    rsTestSetup!MotorLockedRotorTestGood = Abs(TestAndInspectionGood(7).value)
    rsTestSetup!HydrostaticTestGood = Abs(TestAndInspectionGood(8).value)
    rsTestSetup!HydraulicTestGood = Abs(TestAndInspectionGood(9).value)
    rsTestSetup!NPSHTestGood = Abs(TestAndInspectionGood(10).value)
    rsTestSetup!CleanPurgeSealGood = Abs(TestAndInspectionGood(11).value)
    rsTestSetup!PaintCheckGood = Abs(TestAndInspectionGood(12).value)
    rsTestSetup!NameplateGood = Abs(TestAndInspectionGood(13).value)
    rsTestSetup!SupervisorApproval = Abs(TestAndInspectionGood(14).value)

    'update the database
    rsTestSetup.Update

    If boFoundTestData = False Then     'add 8 records for 8 tests
        AddTestData
    End If

    rsTestSetup.Filter = vbNullString
End Sub
Private Sub cmdExit_Click()
    End
End Sub

Private Sub cmdFindMagtrols_Click()
    FindMagtrols
End Sub

Private Sub cmdFindPump_Click()
    ' find the pump whose sn is shown
    Dim sAns As String
    Dim sSO As String
    Dim sParam As String
    Dim sName As String

    Dim I As Integer
    
    'clear the data
    BlankData
    
    'set TC and AI labels with default values
    txtTitle(0).Text = "TC 1"
    txtTitle(1).Text = "(F)"
    txtTitle(2).Text = "TC 2"
    txtTitle(3).Text = "(F)"
    txtTitle(4).Text = "TC 3"
    txtTitle(5).Text = "(F)"
    txtTitle(6).Text = "TC 4"
    txtTitle(7).Text = "(F)"
    txtTitle(20).Text = "Circ Flow"
    txtTitle(21).Text = "(GPM)"
    txtTitle(22).Text = "P1"
    txtTitle(23).Text = "(psig)"
    txtTitle(24).Text = "P2"
    txtTitle(25).Text = "(psig)"
    txtTitle(26).Text = "AI 4"
    txtTitle(27).Text = ""


    For I = 0 To 7
        lblAutoMan(I).Caption = "Auto"
    Next I
    
    lblAutoMan(5).Caption = "Man"
    lblAutoMan(6).Caption = "Man"

    txtFlowDisplay.Enabled = False
    txtSuctionDisplay.Enabled = False
    txtDischargeDisplay.Enabled = False
    txtTemperatureDisplay.Enabled = False
    txtAI1Display.Enabled = False
    txtAI2Display.Enabled = False
    txtAI3Display.Enabled = False
    txtAI4Display.Enabled = False


    cmdFindPump.Default = False

    'set all found booleans to false
'    boUsingHP = False
    boFoundPump = False
    boPumpIsApproved = False
    boFoundTestSetup = False
    boFoundTestData = False


    'get rid of all test dates in combo box
    For I = cmbTestDate.ListCount - 1 To 0 Step -1
        cmbTestDate.RemoveItem 0
    Next I

    rsTestData.Filter = "SerialNumber = ''"

    DataGrid2.ClearFields
    ClearEff

    If rsPumpData.State = adStateOpen Then
        If rsPumpData.BOF = False Or rsPumpData.EOF = False Then
            rsPumpData.Update
        End If
        rsPumpData.Close
    End If

    'parse the serial number to make sure it is formed correctly
    Dim ok As Boolean
    ok = UCase(txtSN.Text) Like "[A-Z][A-Z][0-9][0-9][0-9][0-9][A-Z]" Or UCase(txtSN.Text) Like "[A-Z][A-Z][0-9][0-9][0-9][0-9][A-Z]-[0-9]" Or UCase(txtSN.Text) Like "[A-Z][A-Z][0-9][0-9][0-9][0-9][A-Z]-[0-9][0-9]" Or UCase(txtSN.Text) Like "[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]"
    If Not ok Then
        MsgBox "Serial Number must be 2 letters, 4 numbers, and 1 letter, or 10 numbers. Please re-enter.", vbOKOnly, "Serial Number not correctly formed."
        Exit Sub
    End If
    
    'find the pump listed in the Serial Number text box
    qyPumpData.ActiveConnection = cnPumpData
    qyPumpData.CommandText = "SELECT * From TempPumpData WHERE (((TempPumpData.SerialNumber)='" & _
                             txtSN.Text & "'))"
    rsPumpData.CursorType = adOpenStatic
    rsPumpData.CursorLocation = adUseClient
    rsPumpData.Index = "SerialNumber"
    rsPumpData.Open qyPumpData
    boEpicorFound = False

    If rsPumpData.BOF = True And rsPumpData.EOF = True Then
        'if the bof=eof, we have an empty recordset
        boFoundPump = False
    Else
        'we found it
        boFoundPump = True
    End If

    If boFoundPump = False Then
        'not found in either database, try HP?
        sAns = MsgBox("Pump Not Found in the Database.  Look in Epicor?", vbYesNo, "Can't Find Pump")
        If sAns = vbNo Then     'new pump - don't get data from HP
            boUsingEpicor = False
        Else
            boUsingEpicor = True
'            boUsingHP = False
        End If
'        If boUsingEpicor = False Then
'            sAns = MsgBox("Pump Not Found in the Database.  Look on the HP?", vbYesNo, "Can't Find Pump")
'            If sAns = vbNo Then     'new pump - don't get data from HP
'                 boUsingHP = False
'            Else
'                boUsingHP = True
'            End If
'        End If
        EnablePumpDataControls
        EnableTestSetupDataControls
        EnableTestDataControls
'        BlankData               'clear any data on the screen
        cmdAddNewBalanceHoles.Visible = True

    End If

    If boFoundPump = True Then    'found the pump
        If rsPumpData.Fields("Approved") = True Then
            DisablePumpDataControls                         'if it's in the real database, don't allow changes here
            boPumpIsApproved = True
            lblPumpApproved.Visible = True
            If boCanApprove Then
                cmdApprovePump.Caption = "Unapprove this pump"
            End If
            frmPLCData.cmdApproveTestDate.Enabled = True
        Else
            EnablePumpDataControls                          'it's in the temp database, allow changes
            boPumpIsApproved = False
            boTestDateIsApproved = False
            lblPumpApproved.Visible = False
            If boCanApprove Then
                cmdApprovePump.Caption = "Approve this pump"
            End If
            cmdApproveTestDate.Caption = "You Must Approve Pump First"
            frmPLCData.cmdApproveTestDate.Enabled = False
        End If

        'found the pump, show the data
        txtModelNo.Text = rsPumpData.Fields("ModelNumber")
        frmPLCData.optMfr(0).value = rsPumpData.Fields("ChempumpPump")

        If rsPumpData.Fields("ChempumpPump") = True Then
            SetCombo cmbMotor, "Motor", rsPumpData
            SetCombo cmbDesignPressure, "DesignPressure", rsPumpData
            SetCombo cmbRPM, "RPM", rsPumpData
            SetCombo cmbCirculationPath, "CirculationPath", rsPumpData
            SetCombo cmbStatorFill, "StatorFill", rsPumpData
            SetCombo cmbModel, "Model", rsPumpData
            SetCombo cmbModelGroup, "ModelGroup", rsPumpData
            RatedKW = 999
        End If

        'set the TEMC data
        If rsPumpData.Fields("ChempumpPump") = False Then
            SetCombo cmbTEMCAdapter, "TEMCAdapter", rsPumpData
            SetCombo cmbTEMCAdditions, "TEMCAdditions", rsPumpData
            SetCombo cmbTEMCCirculation, "TEMCCirculation", rsPumpData
            SetCombo cmbTEMCDesignPressure, "TEMCDesignPressure", rsPumpData
            SetCombo cmbTEMCNominalDischargeSize, "TEMCNominalDischargeSize", rsPumpData
            SetCombo cmbTEMCDivisionType, "TEMCDivisionType", rsPumpData
            SetCombo cmbTEMCImpellerType, "TEMCImpellerType", rsPumpData
            SetCombo cmbTEMCInsulation, "TEMCInsulation", rsPumpData
            SetCombo cmbTEMCJacketGasket, "TEMCJacketGasket", rsPumpData
            SetCombo cmbTEMCMaterials, "TEMCMaterials", rsPumpData
            SetCombo cmbTEMCModel, "TEMCModel", rsPumpData
            SetCombo cmbTEMCNominalImpSize, "TEMCNominalImpSize", rsPumpData
            SetCombo cmbTEMCOtherMotor, "TEMCOtherMotor", rsPumpData
            SetCombo cmbTEMCPumpStages, "TEMCPumpStages", rsPumpData
            SetCombo cmbTEMCNominalSuctionSize, "TEMCNominalSuctionSize", rsPumpData
            SetCombo cmbTEMCTRG, "TEMCTRG", rsPumpData
            SetCombo cmbTEMCVoltage, "TEMCVoltage", rsPumpData
        End If

        'write ship to and bill to info
        If Not IsNull(rsPumpData.Fields("ShipToCustomer")) Then
            txtShpNo.Text = rsPumpData.Fields("ShipToCustomer")
        Else
            txtShpNo.Text = vbNullString
        End If

        If Not IsNull(rsPumpData.Fields("BillToCustomer")) Then
            txtBilNo.Text = rsPumpData.Fields("BillToCustomer")
        Else
            txtBilNo.Text = vbNullString
        End If

        sName = "ImpellerDia"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtImpellerDia.Text = sParam

        sName = "DesignFlow"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtDesignFlow.Text = sParam

        sName = "DesignTDH"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtDesignTDH.Text = sParam

        sName = "SpGr"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtSpGr.Text = sParam

        sName = "Remarks"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtRemarks.Text = sParam

        sName = "SalesOrderNumber"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtSalesOrderNumber.Text = sParam

        sName = "ApplicationFluid"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtLiquid.Text = sParam

        sName = "NPSHFile"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtNPSHFileLocation.Text = sParam

        sName = "RVSPartNo"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtRVSPartNo.Text = sParam

        sName = "CustPN"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtXPartNum.Text = sParam

        sName = "CustPO"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtCustPONum.Text = sParam

        'make sure table has custpn - see if last three digits of model no are numeric
'        sName = "SalesOrderNumber"
'        If rsPumpData.Fields(sName).ActualSize <> 0 Then
'            If IsNumeric(Right(rsPumpData.Fields("ModelNumber"), 3)) Then 'no sales order no, must be supermarket
'                rsPumpData.Fields("CustPN") = rsPumpData.Fields("RVSPartNo")
'            Else
'                rsPumpData.Fields("CustPN") = rsPumpData.Fields("ModelNumber")
'            End If
'        End If

        sName = "ApplicationViscosity"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = Format(rsPumpData.Fields(sName), "#0.00")
        Else
            sParam = vbNullString
        End If
        txtViscosity.Text = sParam
        txtTEMCViscosity.Text = sParam
        

'added from TEMC Inspection Report
        sName = "JobNumber"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = ""
        End If
        txtJobNum.Text = sParam

        sName = "Phases"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtNoPhases.Text = sParam

        sName = "ThermalClass"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtThermalClass.Text = sParam

        sName = "ExpClass"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtExpClass.Text = sParam

        sName = "NPSHr"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtNPSHr.Text = sParam

        sName = "LiquidTemperature"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtLiquidTemperature.Text = sParam

        sName = "RatedInputPower"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtRatedInputPower.Text = sParam

        sName = "FLCurrent"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtAmps.Text = sParam

        sName = "TEMCFrameNumber"
        If rsPumpData.Fields(sName).ActualSize <> 0 Then
            sParam = rsPumpData.Fields(sName)
        Else
            sParam = vbNullString
        End If
        txtTEMCFrameNumber.Text = sParam

        optMfr(0).value = rsPumpData.Fields("ChempumpPump")
        optMfr(1).value = Not optMfr(0).value
        
        If rsPumpData.Fields("Field1") = "Feathered" Then
            Me.chkSuperMarketFeathered.value = Checked
        Else
            Me.chkSuperMarketFeathered.value = Unchecked
        End If
        
        'select the testsetup data
        qyTestSetup.ActiveConnection = cnPumpData
        qyTestSetup.CommandText = "SELECT * FROM TempTestSetupData WHERE (((TempTestSetupData.SerialNumber)='" & _
                             txtSN.Text & "')) ORDER BY Date"
'        qyTestSetup.CommandText = "SELECT * FROM TempTestSetupData WHERE (((TempTestSetupData.SerialNumber)='" & _
'               txtSN.Text & "'))"

        With rsTestSetup
            If .State = adStateOpen Then
                .Close
            End If
            .CursorLocation = adUseClient
            .CursorType = adOpenStatic
            .Index = "FindData"
            .Open qyTestSetup
        End With


        'add the selection of dates to the Test Date combo box
        If rsTestSetup.RecordCount <> 0 Then
            For I = 0 To cmbTestDate.ListCount - 1
                cmbTestDate.RemoveItem 0
            Next I
            rsTestSetup.MoveFirst
            For I = 1 To rsTestSetup.RecordCount
                cmbTestDate.AddItem rsTestSetup.Fields("Date")
                rsTestSetup.MoveNext
            Next I
            rsTestSetup.MoveFirst
            boFoundTestSetup = True

            If rsTestSetup.Fields("Approved") = True Then
                DisableTestSetupDataControls                         'if it's in the real database, don't allow changes here
                boTestDateIsApproved = True
                lblTestDateApproved.Visible = True
                If boCanApprove Then
                    cmdApproveTestDate.Caption = "Unapprove this Test Date"
                End If
            Else
                EnableTestSetupDataControls                          'it's in the temp database, allow changes
                lblTestDateApproved.Visible = False
                If boCanApprove Then
                    cmdApproveTestDate.Caption = "Approve this Test Date"
                End If
            End If
            cmbTestDate.ListIndex = 0
        Else
            MsgBox ("There is no Test Setup Data for Serial Number " & txtSN.Text)
            boFoundTestSetup = False        'didn't find any data
            boFoundTestData = False
            cmbTestDate.AddItem Date        'load with today
            cmbTestDate.ListIndex = 0       'show the entry
            EnableTestSetupDataControls
            txtTestRemarks.Text = ""
            txtVibAx.Text = ""
            txtVibRad.Text = ""
            txtThrustBal.Text = ""
            txtTEMCTRGReading.Text = ""
            txtTEMCFrontThrust.Text = ""
            txtTEMCRearThrust.Text = ""
            Exit Sub
        End If

        If cmbTestDate.ListCount = 1 Then       'if there's only one test date, select it
        End If
        Exit Sub
    End If


    Do While boUsingEpicor = True   'need a do loop to exit
        If boUsingEpicor = True Then
            'Dim MyRecord As SNRecord
            Dim MyRecord As SNRecord
    '            I = InStr(1, txtSN.Text, "-")
    '            If I > 0 Then
                MyRecord = GetEpicorODBCData(txtSN.Text, EpicorConnectionString)
    '            End If
            If MyRecord.SONumber = "" Then
                MsgBox ("Not found in Epicor")
                boUsingEpicor = False
                boEpicorFound = False
                Exit Do
            End If
            
            If MyRecord.SONumber = 0 Then
                boEpicorFound = False
                boUsingSupermarketTable = True
                boUsingEpicor = False
            Else
                boEpicorFound = True
                boUsingSupermarketTable = False
            End If
            
            If boEpicorFound = True Then
                boUsingEpicor = False
'                boEpicorFound = True
                txtSalesOrderNumber.Text = MyRecord.SONumber
                txtLineNumber.Text = MyRecord.SOLine
                txtBilNo.Text = MyRecord.Customer
                txtXPartNum.Text = MyRecord.XPartNum
                txtCustPONum.Text = MyRecord.CustomerPO
                
                If MyRecord.ShipTo = "" Then
                    txtShpNo.Text = MyRecord.Customer
                Else
                    txtShpNo.Text = MyRecord.ShipTo
                End If
                txtModelNo.Text = MyRecord.PartNum
                txtModelNo_Change
                txtDesignTDH.Text = MyRecord.TDH
                txtSpGr.Text = MyRecord.SpGr
                txtImpellerDia.Text = MyRecord.ImpellerDiameter
                txtDesignFlow.Text = MyRecord.Flow
                txtNoPhases.Text = MyRecord.Phases
                txtNPSHr.Text = MyRecord.NPSHr
                txtRatedInputPower.Text = MyRecord.RatedInputPower
                txtAmps.Text = MyRecord.FLCurrent
                txtThermalClass.Text = MyRecord.ThermalClass
                txtViscosity.Text = MyRecord.Viscosity
                txtTEMCViscosity.Text = Format((Val(MyRecord.Viscosity) / Val(MyRecord.SpGr)), "000.00")
                txtExpClass.Text = MyRecord.ExpClass
                txtLiquidTemperature.Text = MyRecord.LiquidTemp
                txtLiquid.Text = MyRecord.Fluid
                txtJobNum.Text = MyRecord.JobNumber
    
                For I = 0 To cmbStatorFill.ListCount - 1
                    If InStr(1, UCase$(MyRecord.StatorFill), UCase$(cmbStatorFill.List(I))) <> 0 Then
                        cmbStatorFill.ListIndex = I
                        Exit For
                    End If
                Next I
    
                For I = 0 To cmbCirculationPath.ListCount - 1
                    If InStr(1, UCase$(MyRecord.CirculationPath), UCase$(cmbCirculationPath.List(I))) <> 0 Then
                        cmbCirculationPath.ListIndex = I
                        Exit For
                    End If
                Next I
    
                For I = 0 To cmbDesignPressure.ListCount - 1
                    If InStr(1, MyRecord.DesignPressure, cmbDesignPressure.List(I)) <> 0 Then
                        cmbDesignPressure.ListIndex = I
                        Exit For
                    End If
                Next I
    
                For I = 0 To cmbVoltage.ListCount - 1
                    If InStr(1, MyRecord.Voltage, cmbVoltage.List(I)) <> 0 Then
                        cmbVoltage.ListIndex = I
                        Exit For
                    End If
                Next I
    
                For I = 0 To cmbFrequency.ListCount - 1
                    If InStr(1, MyRecord.Frequency, sName) <> 0 Then
                        cmbFrequency.ListIndex = I
                        Exit For
                    End If
                Next I
    
                For I = 0 To cmbRPM.ListCount - 1
                    If InStr(1, MyRecord.RPM, cmbRPM.List(I)) <> 0 Then
                        cmbRPM.ListIndex = I
                        Exit For
                    End If
                Next I
    
                For I = 0 To cmbSuctDia.ListCount - 1
                    If InStr(1, MyRecord.SuctFlangeSize, cmbSuctDia.List(I)) <> 0 Then
                        cmbSuctDia.ListIndex = I
                        Exit For
                    End If
                Next I
    
                For I = 0 To cmbDischDia.ListCount - 1
                    If InStr(1, MyRecord.DischFlangeSize, cmbDischDia.List(I)) <> 0 Then
                        cmbDischDia.ListIndex = I
                        Exit For
                    End If
                Next I
    
                For I = 0 To cmbTestSpec.ListCount - 1
                    If InStr(1, MyRecord.TestProcedure, cmbTestSpec.List(I)) <> 0 Then
                        cmbTestSpec.ListIndex = I
                        Exit For
                    End If
                Next I
    
                For I = 0 To cmbMotor.ListCount - 1
                    If InStr(1, MyRecord.MotorSize, cmbMotor.List(I)) <> 0 Then
                        cmbMotor.ListIndex = I
                        Exit For
                    End If
                Next I
    
    
            End If
        End If
    Loop

    If boUsingSupermarketTable = True Then
        GetSuperMarketPump MyRecord.PartNum, MyRecord.JobNumber
'        If Not boEpicorFound Then
'            sAns = MsgBox("Is this a supermarket pump?", vbYesNo, "Can't Find Pump")
'            If sAns = vbNo Then     'new pump - don't get data from HP
'                boUsingSupermarketTable = False
'            Else
'                boUsingSupermarketTable = True
'                grpSupermarket.Visible = False
'            End If
'        End If
'
'        If boUsingSupermarketTable = True Then
'            grpSupermarket.Visible = True
'            cmdSelectSupermarket.Caption = "Cancel Supermarket Selection"
'        End If 'boUsingSupermarketTable
    End If
End Sub

Private Sub cmdModifyBalanceHoleData_Click()
    Dim strInput As String
    Dim I As Integer
    Dim sNumber As Integer
    Dim sDia As String
    Dim sBC As String

    cmdModifyBalanceHoleData.Visible = False

    If dgBalanceHoles.SelBookmarks.Count = 0 Then
        cmdModifyBalanceHoleData.Visible = False
        Exit Sub
    End If

    rsBalanceHoles.MoveFirst
    rsBalanceHoles.Move dgBalanceHoles.SelBookmarks(0) - dgBalanceHoles.FirstRow

    sNumber = rsBalanceHoles!Number
    If rsBalanceHoles!diameter = 99 Then
        sDia = "Slot"
    Else
        sDia = str(rsBalanceHoles!diameter)
    End If
    If rsBalanceHoles!boltcircle = 99 Then
        sBC = "Unknown"
    Else
        sBC = str(rsBalanceHoles!boltcircle)
    End If


    'get the data for the balance holes
    strInput = InputBox("Enter Number of Holes (0 to delete entry)", , sNumber)
    If strInput = "" Then
        GoTo DeleteIt
    End If
    sNumber = CInt(strInput)
    If Val(sNumber) = 0 Then
        GoTo DeleteIt
    End If

    strInput = InputBox("Enter Decimal Value of Hole Diameter or 'Slot' (For Example, 0.675) ", , sDia)
    If strInput <> "" Then
        If UCase(strInput) = "SLOT" Then
            strInput = 99
        End If
        sDia = CSng(strInput)
    Else
        GoTo CancelPressed
    End If

    strInput = InputBox("Enter Decimal Value of Bolt Circle or 'Unknown' (For Example, 4.525)", , sBC)
    If strInput <> "" Then
        If UCase(strInput) = "UNKNOWN" Then
            strInput = 99
        End If
        sBC = CSng(strInput)
    Else
        GoTo CancelPressed
    End If

    rsBalanceHoles!Number = sNumber
    rsBalanceHoles!diameter = sDia
    rsBalanceHoles!boltcircle = sBC

    rsBalanceHoles.Update
    'rsBalanceHoles.Filter = "SerialNo = '" & frmPLCData.txtSN.Text & "'"

    GetBalanceHoleData txtSN.Text, cmbTestDate.Text
'    rsBalanceHoles.Requery
    rsBalanceHoles.MoveLast
    dgBalanceHoles.Refresh
    chkBalanceHoles.value = 1
    rsBalanceHoles.MoveFirst

    Exit Sub

CancelPressed:
    MsgBox "No New Balance Hole Data Entered", vbOKOnly

DeleteIt:
    If (MsgBox("Do you really want to delete this entry?", vbYesNo, "Deleting Balance Hole Data. . .")) = vbYes Then
        rsBalanceHoles.Delete
        rsBalanceHoles.Update
        GetBalanceHoleData txtSN.Text, cmbTestDate.Text
'        rsBalanceHoles.Requery
        If Not rsBalanceHoles.EOF Then
            rsBalanceHoles.MoveLast
        End If
        dgBalanceHoles.Refresh
        chkBalanceHoles.value = 1
        If Not rsBalanceHoles.BOF Then
            rsBalanceHoles.MoveFirst
        End If
    End If

  
End Sub

Private Sub cmdReport_Click()
    'view/print a report
    Dim I As Integer

    ExportToExcel

End Sub

Private Sub cmdSearchForPump_Click()
    LoadCombo frmSearch.cmbSearchModel, "TEMCHydraulics"
    
    frmSearch.Show
End Sub

Private Sub cmdSelectSupermarket_Click()
    grpSupermarket.Visible = False
End Sub

Private Sub cmdWriteSP_Click()
    'write the sp to the plc
    Dim rc As String
    Dim S As String

    'write the set point data to the PLC
        bWrite = True
        S = Right$("0000" & txtWriteSPData, 4)
        S = Right$(S, 2) & Left$(S, 2)
        rc = StringToByteArray(S, ByteBuffer)

        DataLength = HexConvert(ByteBuffer, 2)
        DataAddress = StringToHexInt("2005")

        rc = GetData

        bWrite = False
End Sub

'Private Sub Command1_Click()
'    Dim frmem As New InteropDBWithButtons.Form1
'    frmem.ConString = cnPumpData.ConnectionString
'    frmem.Caption = "Email Database Maintenance"
'    frmem.Show 1
'End Sub

Private Sub btnRunNPSH_Click()
    Static OriginalColor As Long
    If btnRunNPSH.Caption = "Run NPSH" Then
        btnRunNPSH.Caption = "Cancel NPSH Run"
        OriginalColor = btnRunNPSH.BackColor
        tmrNPSHr.Enabled = False
        btnRunNPSH.BackColor = vbRed
        If boCanApprove Then
            txtNPSH(5).Visible = True
            lbltab4(5).Visible = True
        Else
            txtNPSH(5).Visible = False
            lbltab4(5).Visible = False
        End If
        WroteNPSHr = False

        frmNPSH.Visible = True
        txtNPSH(5).Enabled = True
        If Val(txtTDH.Text) <= 10 Then
            MsgBox "This test will not work starting with this starting TDH.  Ending test...", vbOKOnly, "Flow is 0"
            btnRunNPSH.Caption = "Run NPSH"
            btnRunNPSH.BackColor = OriginalColor
            frmNPSH.Visible = False
            Exit Sub
        End If
        'load initial values
        If DataGrid2.Row = -1 Then
            MsgBox "You must write the normal test data to this row before you run NPSH.", vbOKOnly, "Nothing written for this row"
            btnRunNPSH.Caption = "Run NPSH"
            btnRunNPSH.BackColor = OriginalColor
            frmNPSH.Visible = False
            Exit Sub
        Else
            DataGrid2.Row = UpDown1.value - 1
        End If

        txtNPSH(0).Text = DataGrid2.Columns("Flow")
        txtNPSH(3).Text = DataGrid2.Columns("TDH")
        txtNPSH(4) = 0
        'txtNPSH(0).Text = txtFlow.Text
        'txtNPSH(3).Text = txtTDH.Text
        txtNPSH(4) = 0
    Else
        btnRunNPSH.Caption = "Run NPSH"
        btnRunNPSH.BackColor = OriginalColor
        frmNPSH.Visible = False
    End If
    
    'ReportToExcel
End Sub

    Private Sub updown1_change()
    Dim sName As String

    If Not rsTestData.BOF Then
        rsTestData.MoveFirst
    End If

    If Not rsTestData.BOF Or Not rsTestData.EOF Then
        rsTestData.Move UpDown1.value - 1
    End If

    sName = "VibrationX"
    If rsTestData.Fields(sName).ActualSize <> 0 Then
        txtVibAx.Text = rsTestData.Fields(sName)
    Else
'        txtVibAx.Text = vbNullString
    End If

    sName = "VibrationY"
    If rsTestData.Fields(sName).ActualSize <> 0 Then
        txtVibRad.Text = rsTestData.Fields(sName)
    Else
'        txtVibRad.Text = vbNullString
    End If

    sName = "Remarks"
    If rsTestData.Fields(sName).ActualSize <> 0 Then
        txtTestRemarks.Text = rsTestData.Fields(sName)
    Else
'        txtTestRemarks.Text = vbNullString
    End If

    sName = "ThrustBalance"
    If rsTestData.Fields(sName).ActualSize <> 0 Then
        txtThrustBal.Text = rsTestData.Fields(sName)
    Else
'        txtThrustBal.Text = vbNullString
    End If

    sName = "TEMCTRG"
    If rsTestData.Fields(sName).ActualSize <> 0 Then
        txtTEMCTRGReading.Text = rsTestData.Fields(sName)
    Else
        txtTEMCTRGReading.Text = 0
'        txtTEMCTRGReading.Text = vbNullString
    End If

    sName = "TEMCFrontThrust"
    If rsTestData.Fields(sName).ActualSize <> 0 Then
        txtTEMCFrontThrust.Text = rsTestData.Fields(sName)
    Else
'        txtTEMCFrontThrust.Text = vbNullString
    End If

    sName = "TEMCRearThrust"
    If rsTestData.Fields(sName).ActualSize <> 0 Then
        txtTEMCRearThrust.Text = rsTestData.Fields(sName)
    Else
'        txtTEMCRearThrust.Text = vbNullString
    End If
    sName = "TEMCMomentArm"
    If rsTestData.Fields(sName).ActualSize <> 0 Then
        txtTEMCMomentArm.Text = rsTestData.Fields(sName)
    Else
'        txtTEMCMomentArm.Text = vbNullString
    End If
    sName = "TEMCThrustRigPressure"
    If rsTestData.Fields(sName).ActualSize <> 0 Then
        txtTEMCThrustRigPressure.Text = rsTestData.Fields(sName)
    Else
'        txtTEMCThrustRigPressure.Text = vbNullString
    End If
    sName = "TEMCViscosity"
    If rsTestData.Fields(sName).ActualSize <> 0 And rsTestData.Fields(sName) <> 0 Then
        txtTEMCViscosity.Text = rsTestData.Fields(sName)
    Else
'        txtTEMCViscosity.Text = vbNullString
    End If

    CalculateTEMCForce

    rsEff.MoveFirst
    rsEff.Move UpDown1.value - 1
End Sub
Sub CalculateTEMCForce()
    Dim NoOfPoles As Integer
    Dim Frequency As Integer
    Dim Additions As String
    Dim Frame As String
    Dim VOverA As Double
    Dim Force As Double
    Dim Gravity As Double

    If Val(txtSpGr.Text) = 0 Then
        Gravity = 1
    Else
        Gravity = CDbl(Val(txtSpGr.Text))
    End If

    'show calculated values
    If Val(txtTEMCFrontThrust.Text) = 0 Then
        If Val(txtTEMCRearThrust.Text) = 0 Then
        'no thrust entered
            lblTEMCFrontRear.Visible = False
            txtTEMCCalcForce.Text = " "
        Else
            'rear thrust
            txtTEMCCalcForce.Text = Format(Gravity * (Val(txtTEMCRearThrust.Text) * Val(txtTEMCMomentArm.Text) - (Val(txtTEMCThrustRigPressure.Text) / 14.223) * 4.5), "##0.0")
            lblTEMCFrontRear.Caption = "REAR"
            lblTEMCFrontRear.Visible = True
        End If
    Else
        'front thrust
        txtTEMCCalcForce.Text = Format(Gravity * (Val(txtTEMCFrontThrust.Text) * Val(txtTEMCMomentArm.Text) + (Val(txtTEMCThrustRigPressure.Text) / 14.223) * 4.5), "##0.0")
        lblTEMCFrontRear.Caption = "FRONT"
        lblTEMCFrontRear.Visible = True
    End If

    If Val(txtTEMCCalcForce.Text) < 0 Then
        txtTEMCCalcForce.Text = -txtTEMCCalcForce
        lblTEMCFrontRear.Caption = "FRONT"
    End If

    'see how many poles we have, it's the next to last number in the frame size
    If Len(txtTEMCFrameNumber) > 2 Then
        NoOfPoles = 2 * Val(Left$(Right$(txtTEMCFrameNumber.Text, 2), 1))
    End If

    If cmbTEMCAdditions.ListIndex <> -1 Then
        Additions = Mid$(cmbTEMCAdditions.List(cmbTEMCAdditions.ListIndex), 2, 1)
        If Additions = "A" Or Additions = "E" Or Additions = "G" Or Additions = "J" Then
            Frequency = 60
        ElseIf Additions = "B" Or Additions = "F" Or Additions = "H" Or Additions = "K" Then
            Frequency = 50
        Else
            Frequency = 0
        End If
    End If

    If Len(txtTEMCFrameNumber.Text) = 3 Then
        If txtTEMCFrameNumber.Text = "529" Then
            Frame = "420"
        Else
            Frame = Left$(txtTEMCFrameNumber, 2) & "0"
        End If
    Else
        Frame = txtTEMCFrameNumber.Text
        If Right$(txtTEMCFrameNumber.Text, 1) = "5" Then
            Frame = Frame & Left$(lblTEMCFrontRear.Caption, 1)
        Else
        End If
    End If
    Force = DLookupA(3, TEMCForceViscosity, 1, Frame)
    If Frequency = 60 Then
        Force = Force / 1.2
    End If
    If Val(txtTEMCViscosity.Text) > 1# Then
        If (Val(txtTEMCCalcForce.Text) > 3 * Force) Then
            lblTEMCPassFail.Visible = True
            lblTEMCPassFail.ForeColor = vbRed
            lblTEMCPassFail.Caption = "FAIL"
        Else
            lblTEMCPassFail.Visible = True
            lblTEMCPassFail.ForeColor = vbGreen
            lblTEMCPassFail.Caption = "PASS"
        End If
    End If

    If (Val(txtTEMCViscosity.Text) > 0.5) And (Val(txtTEMCViscosity.Text) <= 1#) Then
        If (Val(txtTEMCCalcForce.Text) > 2 * Force) Then
            lblTEMCPassFail.Visible = True
            lblTEMCPassFail.ForeColor = vbRed
            lblTEMCPassFail.Caption = "FAIL"
        Else
            lblTEMCPassFail.Visible = True
            lblTEMCPassFail.ForeColor = vbGreen
            lblTEMCPassFail.Caption = "PASS"
        End If
    End If

    If (Val(txtTEMCViscosity.Text) > 0.3) And (Val(txtTEMCViscosity.Text) <= 0.5) Then
        If (Val(txtTEMCCalcForce.Text) > 1.5 * Force) Then
            lblTEMCPassFail.Visible = True
            lblTEMCPassFail.ForeColor = vbRed
            lblTEMCPassFail.Caption = "FAIL"
        Else
            lblTEMCPassFail.Visible = True
            lblTEMCPassFail.ForeColor = vbGreen
            lblTEMCPassFail.Caption = "PASS"
        End If
    End If

    If (Val(txtTEMCViscosity.Text) <= 0.3) Then
        If (Val(txtTEMCCalcForce.Text) > 1# * Force) Then
            lblTEMCPassFail.Visible = True
            lblTEMCPassFail.ForeColor = vbRed
            lblTEMCPassFail.Caption = "FAIL"
        Else
            lblTEMCPassFail.Visible = True
            lblTEMCPassFail.ForeColor = vbGreen
            lblTEMCPassFail.Caption = "PASS"
        End If
    End If
    If NoOfPoles <> 0 Then
        VOverA = (DLookupA(2, TEMCForceViscosity, 1, Frame)) / (NoOfPoles * 30 / Frequency)
    End If
'    If Frequency = 60 Then
'        VOverA = VOverA * 1.2
'    End If

    txtTEMCPVValue.Text = Format(Val(txtTEMCCalcForce.Text) * VOverA, "##0.0")

    If Val(txtTEMCFrontThrust.Text) = 0 And Val(txtTEMCRearThrust.Text) = 0 Then
        txtTEMCPVValue.Text = ""
        txtTEMCCalcForce.Text = ""
        lblTEMCPassFail.Visible = False
    End If


    'calculate reverse head
    txtRevHead.Text = Format(rsTestData.Fields("RBHPress") - rsTestData.Fields("SuctionPressure") * 2.31, "##0.0")
'    txtRevHead.Text = Format((CDbl(Val(txtAI3Display.Text)) - CDbl(Val(txtSuctionDisplay.Text))) * 2.31, "##0.0")
  
End Sub
    Private Sub updown2_change()
    Dim Plothead(1, 7) As Single
    Dim HeadPlot(7, 1) As Single

    Dim PlotEff() As Single
    Dim PlotKW() As Single
    Dim PlotAmps() As Single

    Dim j As Integer

    For j = 0 To UpDown2.value - 1
        Plothead(0, j) = HeadFlow(0, j)
        Plothead(1, j) = HeadFlow(1, j)
        HeadPlot(j, 0) = FlowHead(j, 0)
        HeadPlot(j, 1) = FlowHead(j, 1)
'        ReDim Preserve PlotEff(1, j)
'        PlotEff(0, j) = EffFlow(0, j)
'        PlotEff(1, j) = EffFlow(1, j)
'        ReDim Preserve PlotKW(1, j)
'        PlotKW(0, j) = KWFlow(0, j)
'        PlotKW(1, j) = KWFlow(1, j)
'        ReDim Preserve PlotAmps(1, j)
'        PlotAmps(0, j) = AmpsFlow(0, j)
'        PlotAmps(1, j) = AmpsFlow(1, j)
    Next j

    MSChart1 = HeadPlot
  
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
    DoEfficiencyCalcs
End Sub

Private Sub dgBalanceHoles_SelChange(Cancel As Integer)
    If dgBalanceHoles.SelBookmarks.Count = 0 Then
        cmdModifyBalanceHoleData.Visible = False
    Else
        cmdModifyBalanceHoleData.Visible = True
    End If
End Sub

Private Sub Form_Activate()
    If ProgramEnd = True Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    Dim RetVal As String
    Dim sSendStr As String
    Dim I As Integer
    Dim j As Integer
    Dim sTableName As String
    Dim WhichServer As String
    Dim WhichDatabase As String

    ProgramEnd = False
    Dim objWMIService As Object
    Dim colProcesses As Object
    Set objWMIService = GetObject("winmgmts:")
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name LIKE 'PolarRundown%'")
'    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name LIKE 'Excel%'")
    If colProcesses.Count > 1 Then
        MsgBox "There is already a copy of Polar Rundown running.  You can only have one copy running at a time", vbOKOnly, "Polar Rundown already running"
        Dim f As Form
        For Each f In Forms
            If f.Name <> Me.Name Then
                 Unload f
            End If
        Next
        ProgramEnd = True
        Exit Sub
    Else
    End If
    Set objWMIService = Nothing
    Set colProcesses = Nothing

    debugging = 0   'assume not debugging
    WhichServer = "Production"     'change to production server
    WhichDatabase = "Production"

    If UCase$(Left$(GetMachineName, 5)) = "MROSE" Or UCase$(Left$(GetMachineName, 5)) = "ITTES" Then  'if mickey, see if we want to be in debug
        I = MsgBox("Debug?", vbYesNo)
        If I = vbYes Then
            debugging = 1
            WhichServer = "Production"
            WhichDatabase = "Production"
        Else
        End If
    End If

    If debugging Then
'        GoTo temp
    End If
    'see if the mdb file is where it's supposed to be

    Dim developmentDatabase As String
    developmentDatabase = GetUNCFromLetter("F:") & sDevelopmentDatabase

    If Dir(developmentDatabase) = "" Then
        MsgBox "Development.mdb does not exist on F:, Please contact IT.", , "No Development Database"
        End
    End If

    'get the database info from the new mdb file
    Dim cnDevelopment As New ADODB.Connection
    Dim qyDevelopment As New ADODB.Command
    Dim rsDevelopment As New ADODB.Recordset

    On Error GoTo CannotConnect

    With cnDevelopment
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & developmentDatabase & ";Persist Security Info=False; Jet OLEDB:Database Password=Access7277word;"
        .ConnectionTimeout = 10
        .Open
    End With

On Error GoTo 0
    GoTo Connected

CannotConnect:
    MsgBox "Cannot connect with Development.mdb database.  Please contact IT.", , "Cannot find Connection data."
    End

Connected:

    'we're connected, get the data for the Epicor SQL server
    qyDevelopment.CommandText = "SELECT * FROM Connections WHERE Connections.WhichServer = '" & WhichServer & "' AND WhichDatabase = '" & WhichDatabase & "'"
    qyDevelopment.ActiveConnection = cnDevelopment

    rsDevelopment.CursorLocation = adUseClient
    rsDevelopment.CursorType = adOpenStatic
    rsDevelopment.LockType = adLockOptimistic

    On Error GoTo NoServerData

    rsDevelopment.Open qyDevelopment

On Error GoTo 0
    GoTo GotServerData

NoServerData:

    MsgBox "Cannot connect with Development.mdb database.  Please contact IT.", , "Cannot find Connection data."
    End

GotServerData:

    If rsDevelopment.RecordCount <> 1 Then
        GoTo NoServerData
    End If

    'construct Epicor connection string
    EpicorConnectionString = "Driver={" & rsDevelopment.Fields("ODBCDriver") & "};" & _
                                  "Database=" & rsDevelopment.Fields("DatabaseName") & ";" & _
                                  "Server=" & rsDevelopment.Fields("ServerName") & ";" & _
                                  "UID=" & rsDevelopment.Fields("UserName") & ";" & _
                                  "PWD=" & rsDevelopment.Fields("UserPassword") & ";"


    'make sure we can open the SQL database

    On Error GoTo CannotOpenEpicorSQLServer

    Dim cnTestEpicor As New ADODB.Connection
    cnTestEpicor.ConnectionString = EpicorConnectionString
    cnTestEpicor.Open
    cnTestEpicor.Close
    Set cnTestEpicor = Nothing
On Error GoTo 0

    GoTo FoundEpicorSQLServer

CannotOpenEpicorSQLServer:
    MsgBox "Cannot connect with the Epicor SQL server specified in Development.mdb.  Please contact IT.", , "Cannot connect with Epicor SQL Server"
    End

FoundEpicorSQLServer:
    'get data on rundown database
    rsDevelopment.Close
    qyDevelopment.CommandText = "SELECT * FROM Connections WHERE Connections.WhichServer = 'PolarRundown'"

    On Error GoTo NoRundownDatabase

    rsDevelopment.Open qyDevelopment

    GoTo FoundRundownDatabase

NoRundownDatabase:
    MsgBox "Cannot connect with the Pump Rundown database specified in Development.mdb.  Please contact IT.", , "Cannot connect with Epicor SQL Server"
    End

FoundRundownDatabase:
    If rsDevelopment.RecordCount <> 1 Then
        GoTo NoRundownDatabase
        End
    End If

temp:

    If debugging Then
        sDataBaseName = "c:\databases\PolarData.mdb"
    Else

       sDataBaseName = GetUNCFromLetter("F:") & "\Groups\Shared\databases\PolarData.mdb"
       
'        sDataBaseName = rsDevelopment.Fields("ServerName") & rsDevelopment.Fields("DatabaseName")

'        sDataBaseName = sServerName & "f\groups\shared\databases\PumpData 2k.mdb"
    End If

    Dim tempFSO As Object
    Set tempFSO = CreateObject("Scripting.FileSystemObject")
    ParentDirectoryName = tempFSO.getparentfoldername(sDataBaseName)
    Set tempFSO = Nothing

    'see if we can open the pump rundown database
    On Error GoTo NoRundownDatabase
    With cnPumpData
'        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataBaseName & ";Persist Security Info=False;Jet OLEDB:Database Password=185TitusAve"
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataBaseName & ";Persist Security Info=False;"
        .ConnectionTimeout = 10
        .Open
    End With
On Error GoTo 0


    If debugging = 0 Then
'        Printer.Orientation = vbPRORLandscape
    End If

    lblVersion = "Polar Rundown - Version " & App.Major & "." & App.Minor & "." & App.Revision
    frmPLCData.Caption = "Polar Rundown"

    boFoundPump = False

    Me.Show

    MSChart1.Plot.Axis(VtChAxisIdX).AxisTitle = "Flow"
    MSChart1.Plot.Axis(VtChAxisIdY).AxisTitle = "TDH"
    'MSChart1.Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen = True
    'MSChart1.Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen = True
    MSChart1.Plot.UniformAxis = False
    MSChart1.Plot.SeriesCollection.Item(1).SeriesMarker.Auto = False
    MSChart1.Plot.SeriesCollection.Item(1).Pen.Width = 5
    With MSChart1.Plot.SeriesCollection.Item(1).DataPoints.Item(-1).Marker
        .Visible = True
        .Size = 50
        .Style = VtMarkerStyleCircle
        .FillColor.Automatic = False
        .FillColor.Set 0, 0, 255
    End With
    MSChart1.Plot.AutoLayout = False
    MSChart1.Plot.LocationRect.Max.x = 5600
    MSChart1.Plot.LocationRect.Max.y = 2800
    MSChart1.Plot.LocationRect.Min.x = 0
    MSChart1.Plot.LocationRect.Min.y = 0
    
    'assure that the timers are off
    frmPLCData.tmrGetDDE.Enabled = False

    frmPLCData.tmrStartUp.Enabled = False

    'initialize the PLC network
    RetVal = NetWorkInitialize()
    If RetVal <> 0 Then
        MsgBox ("Can't Initialize Network. Exiting...")
        End
    End If

    If debugging = 0 Then
        'load array of plcs
        I = 0
        Open rsDevelopment.Fields("ServerName") & "PolarPLCAddresses.txt" For Input As 1
        While Not EOF(1)
            Input #1, Description(I)
            For j = 0 To 125
                Input #1, aDevices(I).Address(j)
            Next j
            Input #1, j
            I = I + 1
        Wend
        Close #1

        DeviceCount = I

        If Left$(GetMachineName, 2) = "WV" Then  'if in WV, put MWSC first in loop dropdown
            Dim k As Integer
            For k = 0 To DeviceCount - 1
                If InStr(Description(k), "MWSC") <> 0 Then
                    Exit For
                End If
            Next k
            Description(DeviceCount) = Description(0)
            Description(0) = Description(k)
            Description(k) = Description(DeviceCount)

            aDevices(DeviceCount) = aDevices(0)
            aDevices(0) = aDevices(k)
            aDevices(k) = aDevices(DeviceCount)

        End If

        Dim PLCAddress As String
        For I = 0 To DeviceCount - 1
            PLCAddress = aDevices(I).Address(4) & "." & aDevices(I).Address(5) & "." & aDevices(I).Address(6) & "." & aDevices(I).Address(7)
            RetVal = PingSilent(PLCAddress)
            If RetVal <> 0 Then
                frmPLCData.cmbPLCLoop.AddItem Description(I)
                frmPLCData.cmbPLCLoop.ItemData(frmPLCData.cmbPLCLoop.NewIndex) = I
            End If
        Next I
    End If

    frmPLCData.cmbPLCLoop.AddItem "Add PLC Data Manually"   'enable the controls for manual entry

    'turn on the PLC led

    frmPLCData.cmbPLCLoop.ListIndex = 0
    frmPLCData.tmrGetDDE.Enabled = True

    'hook up to the various databases

    'copy the template of the database here
    'see if it exists
    Dim fdrive As String
    fdrive = GetUNCFromLetter("F:")
    If Dir(fdrive & "\groups\shared\databases" & sEffDataBaseName) = "" Then
        MsgBox "File does not exist at " & fdrive & "\groups\shared\databases" & sEffDataBaseName & ". Please contact IT", vbOKOnly, "Eff.mdb does not exist"
    Else
        'Dim FSO As New FileSystemObject
        FileCopy fdrive & "\groups\shared\databases" & sEffDataBaseName, App.Path & sEffDataBaseName
    End If
    
    
    With cnEffData
        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & sEffDataBaseName & ";Persist Security Info=False"
        .Open
    End With

    'open some recordsets
    rsPumpData.Index = "SerialNumber"
    rsTestSetup.Index = "FindData"
    rsTestData.Index = "PrimaryKey"
    rsPumpData.Open "TempPumpData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
    rsTestSetup.Open "TempTestSetupData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
    rsTestData.Filter = "SerialNumber = ''"
    rsTestData.CursorLocation = adUseClient
    rsTestData.Open "TempTestData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
    rsEff.CursorLocation = adUseClient
    rsEff.Open "Efficiency", cnEffData, adOpenStatic, adLockOptimistic, adCmdTableDirect
    qyBalanceHoles.ActiveConnection = cnPumpData
    rsBalanceHoles.CursorLocation = adUseClient
    rsBalanceHoles.CursorType = adOpenStatic
    rsBalanceHoles.LockType = adLockOptimistic
    qyMisc.ActiveConnection = cnPumpData
    qyMisc.CommandText = "SELECT MiscParameters.ParameterName, MiscParameters.ParameterValue From MiscParameters WHERE (((MiscParameters.ParameterName)='AllowableTDHVariation'));"
    rsMisc.CursorLocation = adUseClient
    rsMisc.CursorType = adOpenStatic
    rsMisc.LockType = adLockBatchOptimistic
    rsMisc.Open qyMisc
    txtNPSH(5).Text = rsMisc!ParameterValue

    If debugging <> 1 Then
        FindMagtrols
    Else
        cmbMagtrol.AddItem "Add Manually"
        cmbMagtrol.ItemData(cmbMagtrol.NewIndex) = 99
        cmbMagtrol.ListIndex = 0
    End If
    optKW(1).value = True
    optKW_Click (1)


    'blank out data grid
    Set DataGrid1.DataSource = rsTestData

    'load the combo boxes
    LoadCombo cmbStatorFill, "StatorFill"
    LoadCombo cmbCirculationPath, "CirculationPath"
    LoadCombo cmbVoltage, "Voltage"
    LoadCombo cmbFrequency, "Frequency"
    LoadCombo cmbMotor, "Motor"
    LoadCombo cmbDesignPressure, "DesignPressure"
    LoadCombo cmbRPM, "RPM"
    LoadCombo cmbOrificeNumber, "OrificeNumber"
    LoadCombo cmbTestSpec, "TestSpecification"
    LoadCombo cmbLoopNumber, "LoopNumber"
    LoadCombo cmbSuctDia, "SuctionDiameter"
    LoadCombo cmbDischDia, "DischargeDiameter"
    LoadCombo cmbTachID, "TachID"
    LoadCombo cmbAnalyzerNo, "AnalyzerNo"
    LoadCombo cmbModel, "Model"
    LoadCombo cmbModelGroup, "ModelGroup"
    LoadCombo cmbMounting, "Mounting"
    LoadCombo cmbPLCNo, "PLCNo"
    LoadCombo cmbFlowMeter, "PumpFlowMeter"
    LoadCombo cmbSuctionPressureTransducer, "SuctionPressureTransducer"
    LoadCombo cmbDischargePressureTransducer, "DischargePressureTransducer"
    LoadCombo cmbTemperatureTransducer, "TemperatureTransducer"
    LoadCombo cmbCirculationFlowMeter, "CirculationFlowMeter"
    'LoadCombo cmbSupermarketModel, "SupermarketPumpData"

    SetFrequencyCombo
    'load the TEMC combo boxes, too
    LoadCombo cmbTEMCAdapter, "TEMCAdapter"
    LoadCombo cmbTEMCAdditions, "TEMCAdditions"
    LoadCombo cmbTEMCCirculation, "TEMCCirculation"
    LoadCombo cmbTEMCDesignPressure, "TEMCDesignPressure"
    LoadCombo cmbTEMCNominalDischargeSize, "TEMCNominalDischargeSize"
    LoadCombo cmbTEMCDivisionType, "TEMCDivisionType"
    LoadCombo cmbTEMCImpellerType, "TEMCImpellerType"
    LoadCombo cmbTEMCInsulation, "TEMCInsulation"
    LoadCombo cmbTEMCJacketGasket, "TEMCJacketGasket"
    LoadCombo cmbTEMCMaterials, "TEMCMaterials"
    LoadCombo cmbTEMCModel, "TEMCModel"
    LoadCombo cmbTEMCNominalImpSize, "TEMCNominalImpSize"
    LoadCombo cmbTEMCOtherMotor, "TEMCOtherMotor"
    LoadCombo cmbTEMCNominalSuctionSize, "TEMCNominalSuctionSize"
    LoadCombo cmbTEMCVoltage, "TEMCVoltage"
    LoadCombo cmbTEMCPumpStages, "TEMCPumpStages"
    LoadCombo cmbTEMCTRG, "TEMCTRG"

    'LoadCombo frmSearch.cmbSearchModel, "Model"

    'fill memory arrays for dlookups
    FillArrays

    'choose the first tab
    frmPLCData.SSTab1.Tab = 0

    'set the grid column names
    Dim c As Column
    For Each c In DataGrid1.Columns
        Select Case c.DataField
        Case "TestDataID"
            c.Visible = False
        Case "SerialNumber"
            c.Visible = False
        Case "Date"
            c.Visible = False
        Case Else ' Show all other columns.
            c.Visible = True
            c.Alignment = dbgRight
        End Select
    Next c

    Set dgBalanceHoles.DataSource = rsBalanceHoles

    For Each c In dgBalanceHoles.Columns
        Select Case c.DataField
        Case "BalanceHoleID"
            c.Visible = False
        Case "SerialNo"
            c.Visible = False
        Case "Date"
            c.Visible = True
            c.Alignment = dbgCenter
            c.Width = 2000
        Case "Number"
            c.Visible = True
            c.Alignment = dbgCenter
            c.Width = 700
        Case "Diameter"
            c.Visible = False
        Case "Diameter1"
            c.Caption = "Diameter"
            c.Visible = True
            c.Alignment = dbgCenter
            c.Width = 700
        Case "BoltCircle1"
            c.Caption = "Bolt Circle"
            c.Visible = True
            c.Alignment = dbgCenter
            c.Width = 800
        Case "BoltCircle"
            c.Visible = False
        Case "SetNo"
            c.Visible = False
        Case Else ' Show all other columns.
            c.Visible = False
        End Select
    Next c

    BlankData

'    If debugging <> 1 Then
        'get user initials
        frmLogin.Show
'    End If

  optMfr(1).value = True
  frmMfr.Visible = False
  
    Pressed = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    End
End Sub

Private Sub Label15_Click()
    frmDiagram.Show
End Sub

Private Sub lblAutoMan_Click(Index As Integer)
    '0 - Flow
    '1 - Suction
    '2 - Discharge
    '3 - Temperature
    '4 - A1 - Circ Flow
    '5 - A2 - RBH Temp
    '6 - A3 - RBH Press
    '7 - A4

    Dim blnEnabled As Boolean

    If lblAutoMan(Index).Caption = "Auto" Then
        lblAutoMan(Index).Caption = "Man"
        blnEnabled = True
    Else
        lblAutoMan(Index).Caption = "Auto"
        blnEnabled = False
    End If

    Select Case Index
        Case 0
            txtFlowDisplay.Enabled = blnEnabled
        Case 1
            txtSuctionDisplay.Enabled = blnEnabled
        Case 2
            txtDischargeDisplay.Enabled = blnEnabled
        Case 3
            txtTemperatureDisplay.Enabled = blnEnabled
        Case 4
            txtAI1Display.Enabled = blnEnabled
        Case 5
            txtAI2Display.Enabled = blnEnabled
        Case 6
            txtAI3Display.Enabled = blnEnabled
        Case 7
            txtAI4Display.Enabled = blnEnabled
    End Select
  
End Sub

Private Sub tmrNPSHr_Timer()
    tmrNPSHr.Enabled = False
    If frmNPSH.Visible = True Then
        btnRunNPSH_Click    'close test
    End If
End Sub

Private Sub txtNPSH_Change(Index As Integer)
    If Index = 5 Then
        If frmNPSH.Visible = True Then
            If rsMisc.State = adStateOpen Then
                rsMisc.Close
            End If
            rsMisc.CursorLocation = adUseClient
            rsMisc.Open "Select * from MiscParameters WHERE (ParameterName = 'AllowableTDHVariation');", cnPumpData, adOpenStatic, adLockOptimistic, adCmdText
            rsMisc.Fields("ParameterValue").value = txtNPSH(5).Text
            rsMisc.Update
        End If
    End If
End Sub

Private Sub txtNPSHFileLocation_Click()
    Dim sTempDir As String
    On Error Resume Next
    sTempDir = CurDir    'Remember the current active directory
    CommonDialog2.DialogTitle = "Select a directory" 'titlebar
    CommonDialog2.InitDir = "\\tei-main-01\f\en\groups\shared\calibration and rundown\npsh\" 'start dir, might be "C:\" or so also
    CommonDialog2.filename = "Select a Directory"  'Something in filenamebox
    CommonDialog2.Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
    CommonDialog2.Filter = "Directories|*.~#~" 'set files-filter to show dirs only
    CommonDialog2.CancelError = True 'allow escape key/cancel
    CommonDialog2.ShowSave   'show the dialog screen

    If Err <> 32755 Then    ' User didn't chose Cancel.
        'Me.SDir.Text = CurDir
    End If

'    ChDir sTempDir  'restore path to what it was at entering

Me.txtNPSHFileLocation.Text = CommonDialog2.filename
  
End Sub





Private Sub txtTitle_LostFocus(Index As Integer)

    ChangeTitles Index
  
End Sub
Private Sub ChangeTitles(ChannelNo As Integer)
    Dim I As Integer
    Dim S As String

    If txtTitle(ChannelNo).Locked = True Then
        Exit Sub
    End If

    Dim qy As New ADODB.Command
    Dim rs As New ADODB.Recordset

    qy.ActiveConnection = cnPumpData

    'see if we have an entry in the table
    qy.CommandText = "SELECT * FROM AITitles " & _
                      "WHERE (((AITitles.SerialNo)= '" & txtSN.Text & "') " & _
                      "AND ((AITitles.Date)= #" & cmbTestDate.Text & "#) " & _
                      "AND ((AITitles.Channel)=" & ChannelNo & "));"

    With rs     'open the recordset for the query
        .CursorLocation = adUseClient
        .CursorType = adOpenStatic
        .LockType = adLockOptimistic
        .Open qy
    End With

    If (rs.BOF = True And rs.EOF = True) Then  'new record
        rs.AddNew
        rs.Fields("SerialNo") = txtSN.Text
        rs.Fields("Date") = cmbTestDate.Text
        rs.Fields("Channel") = CByte(ChannelNo)
        rs.Fields("Title") = txtTitle(ChannelNo).Text
        rs.Update
    Else    'we have an entry, modify it
        rs.Fields("SerialNo") = txtSN.Text
        rs.Fields("Date") = cmbTestDate.Text
        rs.Fields("Channel") = CByte(ChannelNo)
        rs.Fields("Title") = txtTitle(ChannelNo).Text
        rs.Update
    End If

    rs.Close
    Set rs = Nothing
    Set qy = Nothing
  
End Sub

Private Sub optKW_Click(Index As Integer)
    Select Case Index
        Case 0  'add 3 powers
            txtKW.Enabled = False
        Case 1  'enter kw
            txtKW.Enabled = True
        Case 2  'use analog in 4
            txtKW.Enabled = False
    End Select
End Sub

Private Sub optMfr_Click(Index As Integer)
    frmTEMC.Visible = optMfr(1).value
    frmChempump.Visible = optMfr(0).value
    frmTEMCData.Visible = optMfr(1).value
    txtModelNo_Change
End Sub

Private Sub tmrGetDDE_Timer()

'get here every second... get plc and magtrol data

    Dim sSendStr As String
    Dim I As Integer
    Dim VoltMul As Double

    If Calibrating Then
        Exit Sub
    End If

    If debugging Then
        'Exit Sub
    End If


    If boPLCOperating = True Then
        frmPLCData.shpGetPLCData.Visible = True    'turn the PLC led on

        'convert the plc data into real numbers
        'the following data are type real
        txtFlow.Text = ConvertToReal("4050")
        txtSuction.Text = ConvertToReal("4052")
        txtDischarge.Text = ConvertToReal("4054")
        txtTemperature.Text = ConvertToReal("4056")

        txtValvePosition.Text = ConvertToLong("2004")

        frmPLCData.txtTC1.Text = ConvertToLong("2200")
        frmPLCData.txtTC2.Text = ConvertToLong("2202")
        frmPLCData.txtTC3.Text = ConvertToLong("2204")
        frmPLCData.txtTC4.Text = ConvertToLong("2206")

        frmPLCData.txtAI1.Text = ConvertToReal("4060")
        frmPLCData.txtAI2.Text = ConvertToReal("4062")
        frmPLCData.txtAI3.Text = ConvertToReal("4064")
        frmPLCData.txtAI4.Text = ConvertToReal("4066")

        frmPLCData.txtPCoef.Text = ConvertToLong("4036")
        frmPLCData.txtICoef.Text = ConvertToLong("4037")
        frmPLCData.txtDCoef.Text = ConvertToLong("4040")

        frmPLCData.txtSetPoint.Text = ConvertToLong("4035")
        frmPLCData.txtInHg.Text = ConvertToLong("1460")


        'modify the data from PLC format to format that we can use
        'and update the screen
        If txtFlowDisplay.Enabled = False Then
            frmPLCData.txtFlowDisplay = Format$(txtFlow.Text, "###0.00")
        End If
        If txtSuctionDisplay.Enabled = False Then
            frmPLCData.txtSuctionDisplay = Format$((txtSuction.Text) / 10, "##0.00")
        End If
        If txtDischargeDisplay.Enabled = False Then
            frmPLCData.txtDischargeDisplay = Format$(txtDischarge.Text, "##0.00")
        End If
        If txtTemperatureDisplay.Enabled = False Then
            frmPLCData.txtTemperatureDisplay = Format$(txtTemperature.Text, "##0.00")
        End If
        frmPLCData.txtValvePositionDisplay = (txtValvePosition.Text)

        frmPLCData.txtTC1Display = Format$((txtTC1.Text) / 10, "##0.0")
        frmPLCData.txtTC2Display = Format$((txtTC2.Text) / 10, "##0.0")
        frmPLCData.txtTC3Display = Format$((txtTC3.Text) / 10, "##0.0")
        frmPLCData.txtTC4Display = Format$((txtTC4.Text) / 10, "##0.0")

        If txtAI1Display.Enabled = False Then
            frmPLCData.txtAI1Display = Format$(txtAI1.Text, "##0.00")
        End If
        If txtAI2Display.Enabled = False Then
            frmPLCData.txtAI2Display = Format$(txtAI2.Text, "##0.00")
        End If
        If txtAI3Display.Enabled = False Then
            frmPLCData.txtAI3Display = Format$(txtAI3.Text, "##0.00")
        End If
        If txtAI4Display.Enabled = False Then
            frmPLCData.txtAI4Display = Format$(txtAI4.Text, "##0.00")
        End If

        frmPLCData.txtSetPointDisplay = (txtSetPoint.Text)

        frmPLCData.txtInHgDisplay = Format$(txtInHg.Text / 100, "00.00")

        frmPLCData.shpGetPLCData.Visible = False   'turn the PLC led off

        frmPLCData.shpGetMagtrolData.Visible = True 'turn the Magtrol led on
    End If

    If boMagtrolOperating = True Then


        'get the data from the Magtrol
        If Right(cmbMagtrol.List(cmbMagtrol.ListIndex), 4) = "5300" Then
            sSendStr = vbCrLf
            sData = Space$(68)
            VoltMul = Sqr(3)
        Else
            sSendStr = "OT" & vbCrLf
            sData = Space$(183)
            VoltMul = 1#
        End If

        On Error GoTo noresponse
        If UsingNatInst Then
            ibwrt iUD, sSendStr
            ibrd iUD, sData

            'parse the Magrol response
'            vResponse = CWGPIB1.Tasks("Number Parser").Parse(sData)
        Else
            'Dim Databack As String
            sData = TCP.SendGetData("OT")
        End If

            Dim vSplit() As String
            vSplit = Split(Right(sData, Len(sData) - 1), ",")
            If UBound(vSplit) > 0 Then
               ReDim vResponse(UBound(vSplit))
            End If
            For I = 0 To UBound(vSplit) - 1
                If Len(vSplit(I)) <> 0 Then
                    vResponse(I) = CDbl(vSplit(I))
                End If
            Next I

        'format the parsed response
        Dim dd As String
        dd = "- -"

        If Not IsEmpty(vResponse) Then
        '8 entries for 5300 and 12 for the 6530
            If UBound(vResponse) = 8 Or UBound(vResponse) = 12 Then
                'put the responses into the correct text box
                txtV1.Text = Format$(VoltMul * vResponse(1), "###0.0")   'we get back phase voltage and we want line voltage

                Select Case vResponse(0)
                    Case Is < 1
                        txtI1.Text = Format$(vResponse(0), "0.0000")
                    Case Is < 10
                        txtI1.Text = Format$(vResponse(0), "0.000")
                    Case Is < 100
                        txtI1.Text = Format$(vResponse(0), "00.00")
                    Case Else
                        txtI1.Text = Format$(vResponse(0), "000.0")
                End Select

                Select Case vResponse(3)
                    Case Is < 1
                        txtI2.Text = Format$(vResponse(3), "0.0000")
                    Case Is < 10
                        txtI2.Text = Format$(vResponse(3), "0.000")
                    Case Is < 100
                        txtI2.Text = Format$(vResponse(3), "00.00")
                    Case Else
                        txtI2.Text = Format$(vResponse(3), "000.0")
                End Select

                Select Case vResponse(6)
                    Case Is < 1
                        txtI3.Text = Format$(vResponse(6), "0.0000")
                    Case Is < 10
                        txtI3.Text = Format$(vResponse(6), "0.000")
                    Case Is < 100
                        txtI3.Text = Format$(vResponse(6), "00.00")
                    Case Else
                        txtI3.Text = Format$(vResponse(6), "000.0")
                End Select

                txtP1.Text = Format$(vResponse(2) / 1000, "##0.00")     '/ by 1000 to show kW
                txtV2.Text = Format$(VoltMul * vResponse(4), "###0.0")
                'txtI2.Text = Format$(vResponse(3), "###0.0")
                txtP2.Text = Format$(vResponse(5) / 1000, "##0.00")
                txtV3.Text = Format$(VoltMul * vResponse(7), "###0.0")
                'txtI3.Text = Format$(vResponse(6), "###0.0")
                txtP3.Text = Format$(vResponse(8) / 1000, "##0.00")
                If (vResponse(0) * vResponse(1) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7)) <> 0 Then
                    'if we have some measured current
                    'pf = sum of power/sum of VA
                    If Right(cmbMagtrol.List(cmbMagtrol.ListIndex), 4) = "5300" Then
                        'add kw responses and / by 1000 to get to kW
                        txtKW.Text = (vResponse(2) + vResponse(5) + vResponse(8)) / 1000
                        txtPF.Text = Format$(100 * (vResponse(2) + vResponse(5) + vResponse(8)) / (vResponse(1) * vResponse(0) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7)), "0.00")
                    Else
                        txtKW.Text = (vResponse(2) + vResponse(8)) / 1000
                        txtPF.Text = Format$(100 * (vResponse(2) + vResponse(8)) / ((Sqr(3) / 3) * (vResponse(1) * vResponse(0) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7))), "0.00")
                    End If
                    Select Case Val(txtKW.Text)
                        Case Is < 1
                            txtKW.Text = Format$(txtKW.Text, "0.00000")
                        Case Is < 10
                            txtKW.Text = Format$(txtKW.Text, "0.0000")
                        Case Is < 100
                            txtKW.Text = Format$(txtKW.Text, "00.000")
                        Case Else
                            txtKW.Text = Format$(txtKW.Text, "000.00")
                    End Select
                Else
                    txtPF = dd
                End If
            Else
                'no response, show all -- in text boxes
                txtV1.Text = dd
                txtI1.Text = dd
                txtP1.Text = dd
                txtV2.Text = dd
                txtI2.Text = dd
                txtP2.Text = dd
                txtV3.Text = dd
                txtI3.Text = dd
                txtP3.Text = dd
                txtPF = dd
                txtKW = dd
            End If
        End If
    Else    'magtrol not operating
        Dim dbl As Double

        If optKW(0).value = True Then   'add 3 powers
            txtKW.Text = Val(txtP1.Text) + Val(txtP2.Text) + Val(txtP3.Text)
        End If
        If optKW(1).value = True Then   'enter kw
            txtP1.Text = Val(txtKW.Text) / 3
            txtP2.Text = Val(txtKW.Text) / 3
            txtP3.Text = Val(txtKW.Text) / 3
        End If
        If optKW(2).value = True Then   'use ai4
            txtKW.Text = txtAI4Display.Text
            txtP1.Text = Val(txtKW.Text) / 3
            txtP2.Text = Val(txtKW.Text) / 3
            txtP3.Text = Val(txtKW.Text) / 3
        End If

        dbl = Val(txtV1.Text) * Val(txtI1.Text)
        dbl = dbl + Val(txtV2.Text) * Val(txtI2.Text)
        dbl = dbl + Val(txtV3.Text) * Val(txtI3.Text)
        If dbl <> 0 Then
            txtPF.Text = Format$((Val(txtKW.Text) * 1000 * 3 * 100 / (dbl * Sqr(3))), "0.00")
        End If
    End If

noresponse:
On Error GoTo 0
    frmPLCData.shpGetMagtrolData.Visible = False   'turn the Magtrol led off

    'update the little PLC chart
    For I = 1 To 99
        vPlot(0, I) = vPlot(0, I + 1)
        vPlot(1, I) = vPlot(1, I + 1)
    Next I
    vPlot(0, 100) = txtSetPointDisplay
    vPlot(1, 100) = txtFlowDisplay

    'do NPSH stuff
    Dim SuctVelHead As Single
    Dim DischVelHead As Single
    Dim Conversion As Single
    Dim SuctionPSIA As Single
    Dim DischargePSIA As Single
    Dim VaporPress As Single
    Dim SpecVolume As Single
    Dim NPSHa As Single
    Dim NPSHr As Single
    Dim TDH As Single
    Dim pd As Single


    'velocity head
    If cmbSuctDia.ListIndex = -1 Then   'if no suction diameter chosen
        SuctVelHead = 0
    Else
'        pd = DLookup("ActualDia", "PipeDiameters", "ID = " & cmbSuctDia.ListIndex + 1)
        pd = DLookupA(ActualColNo, PipeDiameters, IDColNo, cmbSuctDia.ItemData(cmbSuctDia.ListIndex) + 1)
        SuctVelHead = (0.002592 * Val(txtFlow) ^ 2) / (pd ^ 4)
    End If

    If cmbDischDia.ListIndex = -1 Then     'if no discharge diameter chosen
        DischVelHead = 0
    Else
'        pd = DLookup("ActualDia", "PipeDiameters", "ID = " & cmbDischDia.ListIndex + 1)
        pd = DLookupA(ActualColNo, PipeDiameters, IDColNo, cmbDischDia.ItemData(cmbDischDia.ListIndex) + 1)
        DischVelHead = (0.002592 * Val(txtFlow) ^ 2) / (pd ^ 4)
    End If

    'convert gauges to absolute
    If txtInHgDisplay.Text = "" Then
        Conversion = 0
    Else
        Conversion = txtInHgDisplay * 0.491
    End If

    SuctionPSIA = Val(txtSuctionDisplay) + Conversion
    DischargePSIA = Val(txtDischargeDisplay) + Conversion


    'lookup vapor pressure and specific volume in the arrays that we made
    'if temp is out of range, say so and exit
    If Val(txtTemperatureDisplay) < 40 Or Val(txtTemperatureDisplay) > 165 Then
        txtNPSHa = 0
        Exit Sub
    Else
        I = Val(txtTemperatureDisplay) - 40
'        VaporPress = DLookup("VaporPressure", "VaporPressure", "ID = " & I)
'        SpecVolume = DLookup("SpecificVolume", "VaporPressure", "ID= " & I)
        VaporPress = DLookupA(VaporPressureColNo, VaporPressure, IDColNo, I)
        SpecVolume = DLookupA(SpecificVolumeColNo, VaporPressure, IDColNo, I)
    End If

    If Not ((txtSuctHeight = "") Or (txtDischHeight = "") Or Not IsNumeric(txtSuctHeight) Or Not IsNumeric(txtDischHeight)) Then
        'NPSHa
        NPSHa = (144 * SpecVolume * (SuctionPSIA - VaporPress)) + (txtSuctHeight / 12) + SuctVelHead
'        NPSHa = CalcTDH(DischargePSIA, SuctionPSIA, 0, DischVelHead, 0, txtTemperature)
        txtNPSHa = Format$(NPSHa, "##0.00")

        'tdh
        TDH = CalcTDH(DischargePSIA, SuctionPSIA, 0, (DischVelHead - SuctVelHead), (txtDischHeight / 12) - (txtSuctHeight / 12), txtTemperatureDisplay)
        txtTDH = Format$(TDH, "##0.00")

        If frmNPSH.Visible = True Then
            If Val(txtTDH.Text) > 0 Then
                txtNPSH(2).Text = Format(100 * Val(txtTDH.Text) / Val(txtNPSH(3).Text), "##0.00")
                txtNPSH(1).Text = Format(100 * Val(txtFlow.Text) / Val(txtNPSH(0).Text), "##0.00")
                'check for tdh variation
                If Abs(Val(txtNPSH(1)) - 100) > Val(txtNPSH(5).Text) Then
                    MsgBox "The TDH value has varied more than " & txtNPSH(5) & " %. NPSHr data will NOT be written to the data table", vbOKOnly, "TDH variation too large"
                    btnRunNPSH_Click
                Else    'tdh variation small
                    If Val(txtNPSH(2).Text) <= 97 Then
                        'btnRunNPSH_Click
                        'write the npsh and save
                        If WroteNPSHr = False Then
                            txtNPSH(4).Text = txtNPSHa.Text
                            rsTestData!NPSHr = txtNPSHa.Text
                            rsTestData.Update
                            rsEff!NPSHr = txtNPSHa.Text
                            rsEff.Update
                            WroteNPSHr = True
                            tmrNPSHr.Interval = 5000
                            tmrNPSHr.Enabled = True
                        End If
                    End If  'val < 97
                End If  'check for tdh variation
            End If 'val tdh <=0
        Else    'frm not visible
            'txtNPSHa = Format$(0, "##0.00")
        End If  'if frm visible

    Else
        txtNPSHa = 0
    End If
End Sub
Private Sub tmrStartUp_Timer()
    'we waited for a while, disable the timer
    tmrStartUp.Enabled = False
End Sub
Public Function SetCombo(cmbComboName As ComboBox, sName As String, rs As ADODB.Recordset)
'set the pump parameter combo box to the right data based upon
'the number in the database

    Dim I As Integer
    Dim sParam As String
    Dim qy As New ADODB.Command
    Dim rs1 As New ADODB.Recordset

    If rs.Fields(sName).ActualSize <> 0 Then     'if there's an entry
        sParam = rs.Fields(sName)                'get the index number
        qy.ActiveConnection = cnPumpData
        qy.CommandText = "SELECT * FROM " & sName & " WHERE " & sName & " = " & sParam
        Set rs1 = qy.Execute()                                  'get the record for the index number

        If rs1.BOF = True And rs1.EOF = True Then
            cmbComboName.ListIndex = -1                             'else, remove any pointer
            Exit Function
        End If

        For I = 0 To cmbComboName.ListCount - 1                     'go through the combobox entries
            If cmbComboName.ItemData(I) = rs1.Fields(0) Then     'see when we find the desired index number
                cmbComboName.ListIndex = I                                              'if we do, set the combo box
                Exit For                                            'and we're done
            End If
            cmbComboName.ListIndex = -1                             'else, remove any pointer
        Next I
    Else
        cmbComboName.ListIndex = -1
    End If

    Exit Function
End Function
Private Function SetComboTestSetup(cmbComboName As ComboBox, sFieldName As String, sTableName As String, rs As ADODB.Recordset)
'set the pump parameter combo box to the right data based upon
'the number in the database

'same as setcombo, except here we also pass in the field name

    Dim I As Integer
    Dim sParam As String
    Dim qy As New ADODB.Command
    Dim rs1 As New ADODB.Recordset

    If rs.Fields(sFieldName).ActualSize <> 0 Then
        'if plc number, adjust plcaddress id numbers 1 and 2 to plc 8 and 9 respectively
        If sTableName = "CirculationFlowMeter" Then
            'sParam = rs.Fields(sFieldName) + 7
            sParam = rs.Fields(sFieldName)
            If Val(sParam) < 4 Then
                sParam = str(Val(sParam) + 4)
                rs.Fields(sFieldName) = sParam
            End If
        Else
            sParam = rs.Fields(sFieldName)
        End If
        qy.ActiveConnection = cnPumpData
        qy.CommandText = "SELECT * FROM " & sTableName & " WHERE " & sTableName & " = " & sParam
        Set rs1 = qy.Execute()

        For I = 0 To cmbComboName.ListCount - 1
            If cmbComboName.ItemData(I) = rs1.Fields(0) Then
                cmbComboName.ListIndex = I
                Exit For
            End If
            cmbComboName.ListIndex = -1
        Next I
    Else
        cmbComboName.ListIndex = -1
    End If

    Exit Function
End Function

Private Sub DisablePumpDataControls()
    'disable the pump data controls cause we're just showing what we found

    txtSalesOrderNumber.Enabled = False
    frmMfr.Enabled = False
    txtShpNo.Enabled = False
    txtBilNo.Enabled = False
    txtDesignFlow.Enabled = False
    txtDesignTDH.Enabled = False

    frmMiscPumpData.Enabled = False

    txtModelNo.Enabled = False
    txtImpellerDia.Enabled = False

    frmTEMC.Enabled = False
    frmChempump.Enabled = False

    txtRemarks.Enabled = False
    Me.cmdAddNewTestDate.Visible = False

    cmdEnterPumpData.Enabled = False
  
End Sub
Private Sub DisableTestSetupDataControls()

    cmbTestSpec.Enabled = False
    txtWho.Enabled = False
    txtRMA.Enabled = False

    frmLoopAndXducer.Enabled = False
    frmElecData.Enabled = False
    frmPerfMods.Enabled = False
    frmOtherFiles.Enabled = False
    frmInstrumentTags.Enabled = False
    frmTAndI.Enabled = False
    frmThrustBalMods.Enabled = False
    txtTestSetupRemarks.Enabled = False

    cmdEnterTestSetupData.Enabled = False
    cmbPLCNo.Enabled = False
End Sub
Private Sub DisableTestDataControls()

    cmbPLCLoop.Enabled = False
    frmPumpData.Enabled = False
    frmThermocouples.Enabled = False
    frmAI.Enabled = False
    frmMagtrol.Enabled = False
    fmrMiscTestData.Enabled = False
    frmPLCMisc.Enabled = False
    DataGrid1.Enabled = False
    DataGrid2.Enabled = False
    cmdEnterTestData.Enabled = False
  
End Sub
Private Sub EnableTestSetupDataControls()

    cmbTestSpec.Enabled = True
    txtWho.Enabled = True
    txtRMA.Enabled = True

    frmLoopAndXducer.Enabled = True
    frmElecData.Enabled = True
    frmPerfMods.Enabled = True
    frmOtherFiles.Enabled = True
    frmInstrumentTags.Enabled = True
    frmTAndI.Enabled = True
    frmThrustBalMods.Enabled = True
    txtTestSetupRemarks.Enabled = True

    cmdEnterTestSetupData.Enabled = True
    cmbPLCNo.Enabled = True
End Sub
Private Sub EnableTestDataControls()

    cmbPLCLoop.Enabled = True
    frmPumpData.Enabled = True
    frmThermocouples.Enabled = True
    frmAI.Enabled = True
    frmMagtrol.Enabled = True
    fmrMiscTestData.Enabled = True
    frmPLCMisc.Enabled = True
    DataGrid1.Enabled = True
    DataGrid2.Enabled = True
    cmdEnterTestData.Enabled = True
  
End Sub
Private Sub EnablePumpDataControls()
    'disable the pump data controls cause we're just showing what we found

    txtSalesOrderNumber.Enabled = True
    frmMfr.Enabled = True
    txtShpNo.Enabled = True
    txtBilNo.Enabled = True
    txtDesignFlow.Enabled = True
    txtDesignTDH.Enabled = True

    frmMiscPumpData.Enabled = True

    txtModelNo.Enabled = True
    txtImpellerDia.Enabled = True

    frmTEMC.Enabled = True
    frmChempump.Enabled = True

    txtRemarks.Enabled = True
    Me.cmdAddNewTestDate.Visible = True

    cmdEnterPumpData.Enabled = True
  
End Sub
Private Sub EnableMagtrolFields()
    txtV1.Enabled = True
    txtV2.Enabled = True
    txtV3.Enabled = True
    txtI1.Enabled = True
    txtI2.Enabled = True
    txtI3.Enabled = True
    txtP1.Enabled = True
    txtP2.Enabled = True
    txtP3.Enabled = True
    optKW(0).Visible = True
    optKW(1).Visible = True
    optKW(2).Visible = True
    optKW(1).value = True
    optKW_Click (1)
End Sub
Private Sub DisableMagtrolFields()
    txtV1.Enabled = False
    txtV2.Enabled = False
    txtV3.Enabled = False
    txtI1.Enabled = False
    txtI2.Enabled = False
    txtI3.Enabled = False
    txtP1.Enabled = False
    txtP2.Enabled = False
    txtP3.Enabled = False
    txtKW.Enabled = False
    optKW(0).Visible = False
    optKW(1).Visible = False
    optKW(2).Visible = False
End Sub
Private Sub EnablePLCFields()
    frmPLCData.txtAI1Display.Enabled = True
    frmPLCData.txtAI2Display.Enabled = True
    frmPLCData.txtAI3Display.Enabled = True
    frmPLCData.txtAI4Display.Enabled = True
    frmPLCData.txtTC1Display.Enabled = True
    frmPLCData.txtTC2Display.Enabled = True
    frmPLCData.txtTC3Display.Enabled = True
    frmPLCData.txtTC4Display.Enabled = True
    frmPLCData.txtFlowDisplay.Enabled = True
    frmPLCData.txtSuctionDisplay.Enabled = True
    frmPLCData.txtDischargeDisplay.Enabled = True
    frmPLCData.txtTemperatureDisplay.Enabled = True
    frmPLCData.txtInHgDisplay.Enabled = True
End Sub
Private Sub DisablePLCFields()
    frmPLCData.txtAI1Display.Enabled = False
    frmPLCData.txtAI2Display.Enabled = False
    frmPLCData.txtAI3Display.Enabled = False
    frmPLCData.txtAI4Display.Enabled = False
    frmPLCData.txtTC1Display.Enabled = False
    frmPLCData.txtTC2Display.Enabled = False
    frmPLCData.txtTC3Display.Enabled = False
    frmPLCData.txtTC4Display.Enabled = False
    frmPLCData.txtFlowDisplay.Enabled = False
    frmPLCData.txtSuctionDisplay.Enabled = False
    frmPLCData.txtDischargeDisplay.Enabled = False
    frmPLCData.txtTemperatureDisplay.Enabled = False
    frmPLCData.txtInHgDisplay.Enabled = False
End Sub
Private Sub BlankData()
    txtShpNo.Text = vbNullString
    txtBilNo.Text = vbNullString
    txtModelNo.Text = vbNullString
    cmbMotor.ListIndex = -1
    cmbStatorFill.ListIndex = -1
    cmbVoltage.ListIndex = -1
    cmbDesignPressure.ListIndex = -1
    cmbFrequency.ListIndex = -1
    cmbCirculationPath.ListIndex = -1
    cmbRPM.ListIndex = -1
    cmbModel.ListIndex = -1
    cmbModelGroup.ListIndex = -1
    txtSpGr.Text = vbNullString
    txtImpellerDia.Text = vbNullString
    txtEndPlay.Text = vbNullString
    txtGGap.Text = vbNullString
    txtDesignFlow.Text = vbNullString
    txtDesignTDH.Text = vbNullString
    txtOtherMods.Text = vbNullString
    txtRemarks.Text = vbNullString
    txtSalesOrderNumber.Text = vbNullString
    txtTestSetupRemarks.Text = vbNullString
    txtNPSHFile.Text = vbNullString
    txtPicturesFile.Text = vbNullString
    txtVibrationFile.Text = vbNullString
'    cmbOrificeNumber.ListIndex = 18

    SetFrequencyCombo

'    cmbTestSpec.ListIndex = 6       'default = Rev7
    cmbLoopNumber.ListIndex = -1
    cmbSuctDia.ListIndex = -1
    cmbDischDia.ListIndex = -1
    cmbTachID.ListIndex = -1
    cmbAnalyzerNo.ListIndex = -1
    txtTestRemarks.Text = vbNullString
    txtHDCor.Text = 0
    txtDischHeight.Text = 0
    txtSuctHeight.Text = 0
    txtKWMult.Text = 1
    txtWho.Text = LogInInitials
    txtRMA.Text = vbNullString
    frmPLCData.chkNPSH.value = 0
    frmPLCData.chkPictures.value = 0
    frmPLCData.chkVibration.value = 0
    cmbFlowMeter.ListIndex = -1
    cmbSuctionPressureTransducer.ListIndex = -1
    cmbDischargePressureTransducer.ListIndex = -1
    cmbTemperatureTransducer.ListIndex = -1
    cmbCirculationFlowMeter.ListIndex = -1
    frmPLCData.chkBalanceHoles.value = 0
    frmPLCData.chkCircOrifice.value = 0
    frmPLCData.txtCircOrifice = vbNullString
    frmPLCData.txtImpTrim = vbNullString
    frmPLCData.txtOrifice = vbNullString
    frmPLCData.chkFeathered.value = Unchecked
    frmPLCData.chkTrimmed.value = 0
    frmPLCData.chkCircOrifice.value = 0
    frmPLCData.txtThrustBal = vbNullString
    frmPLCData.txtRPM = vbNullString
    frmPLCData.txtVibAx = vbNullString
    frmPLCData.txtVibRad = vbNullString
    frmPLCData.txtTEMCTRGReading = vbNullString
    dgBalanceHoles.Visible = False
    Me.txtLineNumber.Text = vbNullString
    Me.txtNPSHr.Text = vbNullString
    Me.txtRatedInputPower.Text = vbNullString
    Me.txtAmps.Text = vbNullString
    Me.txtThermalClass.Text = vbNullString
    Me.txtViscosity.Text = vbNullString
    Me.txtTEMCViscosity.Text = vbNullString
    Me.txtExpClass.Text = vbNullString
    Me.txtNoPhases.Text = vbNullString
    Me.txtLiquidTemperature.Text = vbNullString
    Me.txtJobNum.Text = vbNullString
    Me.txtTEMCFrameNumber.Text = vbNullString
    Me.txtLiquid.Text = vbNullString
    Me.chkSuperMarketFeathered.value = Unchecked
    Me.txtRVSPartNo.Text = vbNullString
End Sub
Private Sub AddTestData()
    Dim I As Integer
    Dim sFilter As String

    ClearEff
    rsEff.MoveFirst

    For I = 1 To 8
        rsTestData.AddNew
        rsTestData!SerialNumber = txtSN
        rsTestData!Date = cmbTestDate.List(cmbTestDate.ListIndex)
        rsTestData!testnumber = I
        rsTestData!DataWritten = False
        rsTestData.Update
        DoEfficiencyCalcs
        rsEff.MoveNext
        rsTestData.MoveNext
    Next I
    boFoundTestData = True
    'rsTestData.Update
    rsTestData.Requery
    rsTestData.Resync

   'select the entries from testdata
    sFilter = "SerialNumber='" & txtSN.Text & "' AND Date=#" & cmbTestDate.Text & "#"

    rsTestData.Filter = sFilter

    Set DataGrid1.DataSource = rsTestData

    ' fix the datagrid

    Dim c As Column
    For Each c In DataGrid1.Columns
       Select Case c.DataField
       Case "TestDataID"
          c.Visible = False
       Case "SerialNumber"
          c.Visible = False
       Case "Date"
          c.Visible = False
       Case Else ' Hide all other columns.
          c.Visible = True
          c.Alignment = dbgRight
       End Select
    Next c

    rsEff.Requery
    DataGrid1.Refresh
    DataGrid2.Refresh

End Sub
Private Sub DoEfficiencyCalcs()
    Dim KW As Single, VI As Single, VITemp As Single
    Dim Vave As Single, Iave As Single
    Dim I As Integer
    Dim j As Integer
    Dim HeightDiff As Single
    
    If Not IsNull(rsTestData.Fields("TotalPower")) Then
        KW = rsTestData.Fields("TotalPower")
    Else
        'if we wrote data with an old version, we will not have written total power
        'if total power = 0 and the three individual powers are not 0, add them

        If rsTestData.Fields("PowerA") > 0 Then
            If rsTestData.Fields("PowerB") > 0 Then
                If rsTestData.Fields("PowerC") > 0 Then
                    KW = rsTestData.Fields("PowerA") + rsTestData.Fields("PowerB") + rsTestData.Fields("PowerC")
                End If
            End If
        End If
   End If

    I = 0
    Vave = 0
    Iave = 0
    If Not IsNull(rsTestData.Fields("VoltageA")) And Not IsNull(rsTestData.Fields("CurrentA")) Then
        VI = rsTestData.Fields("VoltageA") * rsTestData.Fields("CurrentA")
        Vave = rsTestData.Fields("VoltageA")
        Iave = rsTestData.Fields("CurrentA")
        If VI <> 0 Then
            I = I + 1
        End If
    End If
    If Not IsNull(rsTestData.Fields("VoltageB")) And Not IsNull(rsTestData.Fields("CurrentB")) Then
        VITemp = rsTestData.Fields("VoltageB") * rsTestData.Fields("CurrentB")
        If VITemp <> 0 Then
            I = I + 1
            VI = VI + VITemp
            Vave = Vave + rsTestData.Fields("VoltageB")
            Iave = Iave + rsTestData.Fields("CurrentB")
        End If
    End If
    If Not IsNull(rsTestData.Fields("VoltageC")) And Not IsNull(rsTestData.Fields("CurrentC")) Then
        VITemp = rsTestData.Fields("VoltageC") * rsTestData.Fields("CurrentC")
        If VITemp <> 0 Then
            I = I + 1
            VI = VI + VITemp
            Vave = Vave + rsTestData.Fields("VoltageC")
            Iave = Iave + rsTestData.Fields("CurrentC")
        End If
    End If
    If KW = 0 Then
        For j = 1 To rsEff.Fields.Count - 1
            rsEff.Fields(j) = 0
        Next j
'        Exit Sub
    End If
    If VI <> 0 Then
        rsEff.Fields("Volts") = Vave / I
        rsEff.Fields("Amps") = Iave / I
        rsEff.Fields("PowerFactor") = 1000 * I * KW / (VI * Sqr(3))
        rsEff.Fields("PowerFactor") = 100 * rsEff.Fields("PowerFactor")
    Else
        rsEff.Fields("PowerFactor") = 0
    End If

    If optMfr(0).value = True Then
        If cmbStatorFill.ListIndex = -1 Then
            rsEff.Fields("MotorEfficiency") = Format$(0, "0.00")

        Else
            rsEff.Fields("Motorefficiency") = Format$(Round(MotorEfficiency(KW, cmbMotor.ItemData(cmbMotor.ListIndex), cmbStatorFill.ItemData(cmbStatorFill.ListIndex)), 1), "00.0")
'            rsEff.Fields("Motorefficiency") = Format$(Round(MotorEfficiency(KW, cmbMotor.ListIndex, cmbStatorFill.ListIndex), 1), "00.0")
        End If
    Else
        rsEff.Fields("MotorEfficiency") = Format$(Round(TEMCMotorEfficiency(KW, txtTEMCFrameNumber.Text, 460, RatedKW), 1), "00.0")
    End If

    Dim sHDCor As Single
    Dim sDisc As Single
    Dim sSuct As Single
    If IsNull(rsTestSetup.Fields("HDCor")) Then
        sHDCor = 0
    Else
        sHDCor = rsTestSetup.Fields("HDCor")
    End If
    If IsNull(rsTestSetup.Fields("DischargeGageHeight")) Then
        sDisc = 0
    Else
        sDisc = rsTestSetup.Fields("DischargeGageHeight")
    End If
    If IsNull(rsTestSetup.Fields("SuctionGageHeight")) Then
        sSuct = 0
    Else
        sSuct = rsTestSetup.Fields("SuctionGageHeight")
    End If
    HeightDiff = sHDCor + sDisc / 12 - sSuct / 12
    If (cmbDischDia.ListIndex <> -1 And cmbSuctDia.ListIndex <> -1) Then
        rsEff.Fields("VelocityHead") = CalcVelHead(rsTestData.Fields("Flow"), cmbDischDia.ItemData(cmbDischDia.ListIndex) + 1, cmbSuctDia.ItemData(cmbSuctDia.ListIndex) + 1)
    End If
'    rsEff.Fields("VelocityHead") = CalcVelHead(rsTestData.Fields("Flow"), cmbDischDia.ListIndex + 1, cmbSuctDia.ListIndex + 1)
    rsEff.Fields("TDH") = CalcTDH(rsTestData.Fields("DischargePressure"), rsTestData.Fields("SuctionPressure"), rsTestData.Fields("SuctionInHg"), rsEff.Fields("VelocityHead"), HeightDiff, rsTestData.Fields("TemperatureSuction"))
    rsEff.Fields("ElecHP") = 1000 * KW / 746
'    If (DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))) <> 0 And KW <> 0) Then
        If Int(rsTestData.Fields("TemperatureSuction")) >= 40 Then
            If (DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))) <> 0 And KW <> 0) Then
    '        rsEff.Fields("LiquidHP") = (rsEff.Fields("TDH") * rsTestData.Fields("Flow") * DLookup("TDHCorr", "TempCorrection", "Temp = 68")) / (3960 * DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))))
            rsEff.Fields("LiquidHP") = (rsEff.Fields("TDH") * rsTestData.Fields("Flow") * DLookupA(TDHColNo, TempCorrection, TempColNo, 68)) / (3960 * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))))
    '        rsEff.Fields("OverallEfficiency") = (0.189 * rsTestData.Fields("Flow") * rsEff.Fields("TDH") * DLookup("TDHCorr", "TempCorrection", "Temp = 68")) / (10 * KW * DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))))
            rsEff.Fields("OverallEfficiency") = (0.189 * rsTestData.Fields("Flow") * rsEff.Fields("TDH") * DLookupA(TDHColNo, TempCorrection, TempColNo, 68)) / (10 * KW * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))))
            If rsEff.Fields("MotorEfficiency") <> 0 Then
                rsEff.Fields("HydraulicEfficiency") = 100 * rsEff.Fields("OverallEfficiency") / rsEff.Fields("MotorEfficiency")
            Else
                rsEff.Fields("HydraulicEfficiency") = 0
            End If
        Else
            rsEff.Fields("LiquidHP") = 0
            rsEff.Fields("OverallEfficiency") = 0
        End If
    
    Else
        rsEff.Fields("LiquidHP") = 0
        rsEff.Fields("OverallEfficiency") = 0
    End If


    I = rsEff.AbsolutePosition
    If Not IsNull(rsTestData.Fields("Flow")) Then
        rsEff.Fields("Flow") = rsTestData.Fields("Flow")
        HeadFlow(0, I - 1) = rsTestData.Fields("Flow")
        HeadFlow(1, I - 1) = rsEff.Fields("TDH")
        FlowHead(I - 1, 0) = rsTestData.Fields("Flow")
        FlowHead(I - 1, 1) = rsEff.Fields("TDH")
        
'        EffFlow(0, i - 1) = rsTestData.Fields("Flow")
'        EffFlow(1, i - 1) = rsEff.Fields("OverallEfficiency")
'        KWFlow(0, i - 1) = rsTestData.Fields("Flow")
'        KWFlow(1, i - 1) = KW
'        AmpsFlow(0, i - 1) = rsTestData.Fields("Flow")
'        AmpsFlow(1, i - 1) = rsEff.Fields("Amps")
    Else
        HeadFlow(0, I - 1) = 0
        HeadFlow(1, I - 1) = 0
        FlowHead(I - 1, 0) = 0
        FlowHead(I - 1, 1) = 0

'        EffFlow(0, i - 1) = 0
'        EffFlow(1, i - 1) = 0
'        KWFlow(0, i - 1) = 0
'        KWFlow(1, i - 1) = 0
'        AmpsFlow(0, i - 1) = 0
'        AmpsFlow(1, i - 1) = 0
    End If

    Dim Plothead(1, 7) As Single
    Dim HeadPlot(7, 1) As Single
    'ReDim Preserve Plothead(1, j)
    'ReDim Preserve HeadPlot(j, 1)

'    Dim PlotEff() As Single
'    Dim PlotKW() As Single
'    Dim PlotAmps() As Single
'    ReDim PlotHead(0, 0)
'    ReDim PlotEff(0, 0)
'    ReDim PlotKW(0, 0)
'
    For j = 0 To UpDown2.value - 1
'        If HeadFlow(1, j) <> 0 Then
'            ReDim Preserve Plothead(1, j)
'            ReDim Preserve HeadPlot(j, 1)
            Plothead(0, j) = HeadFlow(0, j)
            Plothead(1, j) = HeadFlow(1, j)
            HeadPlot(j, 0) = FlowHead(j, 0)
            HeadPlot(j, 1) = FlowHead(j, 1)
'            ReDim Preserve PlotEff(1, j)
'            PlotEff(0, j) = EffFlow(0, j)
'            PlotEff(1, j) = EffFlow(1, j)
'            ReDim Preserve PlotKW(1, j)
'            PlotKW(0, j) = KWFlow(0, j)
'            PlotKW(1, j) = KWFlow(1, j)
'            ReDim Preserve PlotAmps(1, j)
'            PlotAmps(0, j) = AmpsFlow(0, j)
'            PlotAmps(1, j) = AmpsFlow(1, j)
'        End If
    Next j




'    SetGraphMax (Plothead())
'    If UBound(PlotHead()) <> 0 Then

'fix 4/29/19

        MSChart1.ChartData = HeadPlot

'    End If

    'copy fields for reports
    rsEff.Fields("DischPress") = rsTestData.Fields("Dischargepressure")
    rsEff.Fields("SuctPress") = rsTestData.Fields("Suctionpressure")
'    rsEff.Fields("Volts") = rsTestData.Fields("VoltageA")
'    rsEff.Fields("Amps") = rsTestData.Fields("CurrentA")
    rsEff.Fields("KW") = KW
    rsEff.Fields("Freq") = rsTestData.Fields("VFDFrequency")
    rsEff.Fields("RPM") = rsTestData.Fields("RPM")
    rsEff.Fields("Pos") = rsTestData.Fields("ThrustBalance")
    rsEff.Fields("NPSHa") = rsTestData.Fields("NPSHa")
    rsEff.Fields("NPSHr") = rsTestData.Fields("NPSHr")
    rsEff.Fields("InputPower") = rsTestData.Fields("TotalPower")
    rsEff.Fields("Temperature") = rsTestData.Fields("TemperatureSuction")
    rsEff.Fields("CircFlow") = rsTestData.Fields("CircFlow")
    rsEff.Fields("VibrationX") = rsTestData.Fields("VibrationX")
    rsEff.Fields("VibrationY") = rsTestData.Fields("VibrationY")
    rsEff.Fields("CurrentA") = rsTestData.Fields("CurrentA")
    rsEff.Fields("CurrentB") = rsTestData.Fields("CurrentB")
    rsEff.Fields("CurrentC") = rsTestData.Fields("CurrentC")
    rsEff.Fields("VoltageA") = rsTestData.Fields("VoltageA")
    rsEff.Fields("VoltageB") = rsTestData.Fields("VoltageB")
    rsEff.Fields("VoltageC") = rsTestData.Fields("VoltageC")
    rsEff.Fields("TC1") = rsTestData.Fields("TC1")
    rsEff.Fields("TC2") = rsTestData.Fields("TC2")
    rsEff.Fields("TC3") = rsTestData.Fields("TC3")
    rsEff.Fields("TC4") = rsTestData.Fields("TC4")
    rsEff.Fields("RBHTemp") = rsTestData.Fields("RBHTemp")
    rsEff.Fields("RBHPress") = rsTestData.Fields("RBHPress")
    rsEff.Fields("AI4") = rsTestData.Fields("AI4")
    rsEff.Fields("Remarks") = rsTestData.Fields("Remarks")
    rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCFrontThrust")
    rsEff.Fields("TEMCRearThrust") = rsTestData.Fields("TEMCRearThrust")
    rsEff.Fields("TEMCTRG") = rsTestData.Fields("TEMCTRG")
    rsEff.Fields("TEMCThrustRigPressure") = rsTestData.Fields("TEMCThrustRigPressure")
    rsEff.Fields("TEMCMomentArm") = rsTestData.Fields("TEMCMomentArm")
    rsEff.Fields("TEMCViscosity") = rsTestData.Fields("TEMCViscosity")
    If Not IsNull(rsEff.Fields("TEMCFrontThrust")) Then
        txtTEMCFrontThrust.Text = rsEff.Fields("TEMCFrontThrust")
    End If
    If Not IsNull(rsEff.Fields("TEMCREarThrust")) Then
        txtTEMCRearThrust.Text = rsEff.Fields("TEMCREarThrust")
    End If
    If (Not IsNull(rsEff.Fields("TEMCViscosity"))) And (rsEff.Fields("TEMCViscosity") <> 0) Then
        txtTEMCViscosity.Text = rsEff.Fields("TEMCViscosity")
    End If
    If Not IsNull(rsTestData.Fields("TEMCThrustRigPressure")) Then
        txtTEMCThrustRigPressure.Text = rsTestData.Fields("TEMCThrustRigPressure")
    End If
    If Not IsNull(rsTestData.Fields("TEMCMomentArm")) Then
        txtTEMCMomentArm.Text = rsTestData.Fields("TEMCMomentArm")
    End If

 '   If Not IsNull(Me.txtAI3Display.Text) Then
 '       Me.txtAI3Display = rsTestData.Fields("RBHPress")
 '   End If

    CalculateTEMCForce

    If Not IsNull(txtTEMCCalcForce.Text) Then
        rsEff.Fields("TEMCCalculatedForce") = Val(txtTEMCCalcForce.Text)
    Else
        rsEff.Fields("TEMCCalculatedForce") = 0
    End If

    If Not IsNull(txtTEMCPVValue.Text) Then
        rsEff.Fields("TEMCPV") = Val(txtTEMCPVValue.Text)
    Else
        rsEff.Fields("TEMCPV") = 0
    End If

    If Val(txtTEMCFrontThrust.Text) <> 0 Then
        rsEff.Fields("TEMCFR") = "F"
'        rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCFrontThrust")
    Else
        If Val(txtTEMCRearThrust.Text) = 0 Then
            'no thrust
            rsEff.Fields("TEMCFR") = " "
            rsEff.Fields("TEMCFrontThrust") = 0
        Else
            rsEff.Fields("TEMCFR") = "R"
'            rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCRearThrust")
        End If
    End If

    rsEff.Fields("TEMCForceDirection") = Left(lblTEMCFrontRear.Caption, 1)

    rsEff.Update
    DataGrid2.Refresh

  
End Sub
Private Sub ClearEff()
'    Dim I As Integer, j As Integer
    Dim qy As New ADODB.Command

    If rsEff.State = adStateOpen Then
        If Not (rsEff.BOF = True Or rsEff.EOF = True) Then
            rsEff.CancelUpdate
        End If
        rsEff.Close
    End If
    qy.ActiveConnection = cnEffData
    qy.CommandText = "DROP TABLE Efficiency"
    rsEff.Open qy
    qy.CommandText = "SELECT EfficiencyOrg.* INTO Efficiency FROM EfficiencyOrg;"
    rsEff.Open qy
    rsEff.Open "Efficiency", cnEffData, adOpenStatic, adLockOptimistic, adCmdTableDirect

    rsEff.Requery
    DataGrid2.Refresh

    Dim c As Column
    For Each c In DataGrid2.Columns
        c.Alignment = dbgCenter
        c.Width = 750
        Select Case c.ColIndex
            Case 1
                c.Caption = "Flow"
                c.NumberFormat = "###0.00"
            Case 2
                c.Caption = "TDH"
                c.NumberFormat = "00.0"
            Case 3
                c.Caption = "Overall Eff"
                c.NumberFormat = "00.00"
                c.Width = 850
            Case 4
                c.Caption = "PF"
                c.NumberFormat = "00.0"
            Case 5
                c.Caption = "Vel Head"
                c.NumberFormat = "00.00"
            Case 6
                c.Caption = "Elec HP"
                c.NumberFormat = "#00.0"
            Case 7
                c.Caption = "Liq HP"
                c.NumberFormat = "#00.0"
            Case Else
                c.Visible = False
        End Select
    Next c
  
End Sub
Function JustAlphaNumeric(char As String) As String
    Select Case Asc(char)
        Case 42             ' *
            JustAlphaNumeric = char
        Case 48 To 57       ' 0 - 9
            JustAlphaNumeric = char
        Case 65 To 90       ' A - Z
            JustAlphaNumeric = char
        Case 97 To 122      ' a - z
            JustAlphaNumeric = UCase(char)
        Case Else
            JustAlphaNumeric = ""
    End Select
End Function



Private Sub txtI1_Change()
    txtI2.Text = txtI1.Text
    txtI3.Text = txtI1.Text
End Sub

Private Sub txtModelNo_Change()
    Dim I As Integer
    Dim S As String
    Dim sFull As String
    Dim boDone As Boolean
    Dim boRepeat As Boolean

    Static bo3Digits As Boolean         '3 digits in frame number
    Static bo2Digits As Boolean         '2 digits in stages

    If optMfr(0).value = True Then
        Exit Sub
    End If

    cmbTEMCAdapter.ListIndex = -1
    cmbTEMCAdditions.ListIndex = -1
    cmbTEMCCirculation.ListIndex = -1
    cmbTEMCDesignPressure.ListIndex = -1
    cmbTEMCNominalDischargeSize.ListIndex = -1
    cmbTEMCDivisionType.ListIndex = -1
    cmbTEMCImpellerType.ListIndex = -1
    cmbTEMCInsulation.ListIndex = -1
    cmbTEMCJacketGasket.ListIndex = -1
    cmbTEMCMaterials.ListIndex = -1
    cmbTEMCModel.ListIndex = -1
    cmbTEMCNominalImpSize.ListIndex = -1
    cmbTEMCOtherMotor.ListIndex = -1
    cmbTEMCPumpStages.ListIndex = -1
    cmbTEMCNominalSuctionSize.ListIndex = -1
    cmbTEMCTRG.ListIndex = -1
    cmbTEMCVoltage.ListIndex = -1


    'first, get rid of spaces, dashes, etc

    S = ""
    For I = 1 To Len(txtModelNo.Text)
        S = S & JustAlphaNumeric(Mid$(txtModelNo.Text, I, 1))
    Next I

    'next, fill out the model number to it's max length of 24 characters

    boDone = False
    boRepeat = False

    Do While Not boDone
        sFull = ""
        For I = 1 To Len(S)
            Select Case I
                Case 1
                    'type
                    sFull = sFull & Mid$(S, I, 1)
                Case 2
                    'adapter
                    If IsNumeric(Mid$(S, I, 1)) Then
                        S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
                        boRepeat = True
                        Exit For
                    Else
                        sFull = sFull & Mid$(S, I, 1)
                        boRepeat = False
                    End If
                Case 3
                    'materials
                    sFull = sFull & Mid$(S, I, 1)
                Case 4
                'design pressure
                    sFull = sFull & Mid$(S, I, 1)
                Case 5
                'motor frame number - digit 1
                    sFull = sFull & Mid$(S, I, 1)
                Case 6
                'motor frame number - digit 2
                    sFull = sFull & Mid$(S, I, 1)
                Case 7
                'motor frame number - digit 3
                    sFull = sFull & Mid$(S, I, 1)
                Case 8
                'motor frame number - digit 4
                    If IsNumeric(Mid$(S, I, 1)) Then
                        sFull = sFull & Mid$(S, I, 1)
                        boRepeat = False
                    Else    '3 digits
'                        s = Left$(s, i - 1) & "*" & Right$(s, Len(s) - i + 1)
                        S = Left$(S, I - 4) & "0" & Right$(S, Len(S) - I + 4)
                        boRepeat = True
                        Exit For
                    End If
                Case 9
                'insulation
                    sFull = sFull & Mid$(S, I, 1)
                Case 10
                'voltage
                    sFull = sFull & Mid$(S, I, 1)
                Case 11
                'other motor specs
                    If Mid$(S, I, 1) = "M" Or Mid$(S, I, 1) = "R" Or Mid$(S, I, 1) = "L" Or Mid$(S, I, 1) = "G" Or Mid$(S, I, 1) = "N" Then
                        S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
                        boRepeat = True
                        Exit For
                    Else
                        sFull = sFull & Mid$(S, I, 1)
                        boRepeat = False
                    End If
                Case 12
                ' TRG
                    sFull = sFull & Mid$(S, I, 1)
                Case 13
                'Nominal discharge - digit 1
                    sFull = sFull & Mid$(S, I, 1)
                Case 14
                'nominal discharge - digit 2
                    sFull = sFull & Mid$(S, I, 1)
                Case 15
                'nominal suction - digit 1
                    sFull = sFull & Mid$(S, I, 1)
                Case 16
                'nominal suction - digit 2
                    sFull = sFull & Mid$(S, I, 1)
                Case 17
                'nominal impeller size
                    sFull = sFull & Mid$(S, I, 1)
                Case 18
                'impeller type
                    If Mid$(S, I, 1) <> "*" Then
                        S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
                        boRepeat = True
                        Exit For
                    Else
                        sFull = sFull & Mid$(S, I, 1)
                        boRepeat = False
                    End If
                Case 19
                'Division type
                    If IsNumeric(Mid$(S, I, 1)) Then
                        S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
                        boRepeat = True
                        Exit For
                    Else
                        sFull = sFull & Mid$(S, I, 1)
                        boRepeat = False
                    End If
                Case 20
                'pump stages - digit 1
                    sFull = sFull & Mid$(S, I, 1)
                Case 21
                'pump jacket
                    If Mid$(S, I, 1) = "A" Or Mid$(S, I, 1) = "B" Or Mid$(S, I, 1) = "E" Or Mid$(S, I, 1) = "F" Or _
                                      Mid$(S, I, 1) = "G" Or Mid$(S, I, 1) = "H" Or Mid$(S, I, 1) = "J" Or Mid$(S, I, 1) = "K" Then
                        S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
                        boRepeat = True
                    Else
                        sFull = sFull & Mid$(S, I, 1)
                        boRepeat = False
                    End If
                Case 22
                'additions
                      sFull = sFull & Mid$(S, I, 1)
                Case 23
                'circulation
                      sFull = sFull & Mid$(S, I, 1)
            End Select
        Next I
        If Not boRepeat Then
            boDone = True
        End If
    Loop

    For I = 1 To Len(sFull)
        Select Case I
            Case 1
                ParseTEMCModelNo cmbTEMCModel, Mid$(sFull, I, 1)
            Case 2
                ParseTEMCModelNo cmbTEMCAdapter, Mid$(sFull, I, 1)
            Case 3
                ParseTEMCModelNo cmbTEMCMaterials, Mid$(sFull, I, 1)
            Case 4
                ParseTEMCModelNo cmbTEMCDesignPressure, Mid$(sFull, I, 1)
            Case 5
                    If Val(Mid$(sFull, I, 1)) = 0 Then
                        txtTEMCFrameNumber.Text = Mid$(sFull, 6, 3)
                    Else
                        txtTEMCFrameNumber.Text = Mid$(sFull, 5, 4)
                    End If
            Case 9
                    ParseTEMCModelNo cmbTEMCInsulation, Mid$(sFull, I, 1)
            Case 10
                    ParseTEMCModelNo cmbTEMCVoltage, Mid$(sFull, I, 1)
            Case 11
                    ParseTEMCModelNo cmbTEMCOtherMotor, Mid$(sFull, I, 1)
            Case 12
                    ParseTEMCModelNo cmbTEMCTRG, Mid$(sFull, I, 1)
            Case 13
                    ParseTEMCModelNo cmbTEMCNominalDischargeSize, Mid$(sFull, I, 2)
            Case 14
            Case 15
                    ParseTEMCModelNo cmbTEMCNominalSuctionSize, Mid$(sFull, I, 2)
            Case 16
            Case 17
                    ParseTEMCModelNo cmbTEMCNominalImpSize, Mid$(sFull, I, 1)
            Case 18
                    ParseTEMCModelNo cmbTEMCImpellerType, Mid$(sFull, I, 1)
            Case 19
                    ParseTEMCModelNo cmbTEMCDivisionType, Mid$(sFull, I, 1)
            Case 20
                    ParseTEMCModelNo cmbTEMCPumpStages, Mid$(sFull, I, 1)
            Case 21
                    ParseTEMCModelNo cmbTEMCJacketGasket, Mid$(sFull, I, 1)
            Case 22
                    ParseTEMCModelNo cmbTEMCAdditions, Mid$(sFull, I, 1)
                    ParseTEMCModelNo cmbTEMCCirculation, "*"
            Case 23
'                    ParseTEMCModelNo cmbTEMCCirculation, Mid$(sFull, I, 1)

        End Select
    Next I
    
    'give alerts on certain conditions
    Dim msg As String
    msg = ""
    
    'look for 4 in third digit of model number for CO2 pump
    If Mid(txtModelNo.Text, 3, 1) = "4" Then
        msg = "Special CO2 model. Requires special thrust rig and circulation flow setting. See Engineering"
    End If
    
    If Left(cmbTEMCVoltage, 3) = "[6]" Then
        msg = "575V transformer required for Rundown and TRG"
    End If
'    If Left(cmbTEMCTRG, 3) = "[L]" Or InStr("X[B][F][H][K]", Left(cmbTEMCAdditions, 3)) > 1 Then
    If Left(cmbTEMCTRG, 3) = "[L]" Then
        If msg = "" Then
            msg = "VFD required for Rundown and TRG"
        Else
            msg = msg & " and " & "VFD required for Rundown and TRG"
        End If
    End If
        
    If InStr("X[B][F][H][K]", Left(cmbTEMCAdditions, 3)) > 1 Then
        If msg = "" Then
            msg = "VFD required for Rundown, standard drive required for TRG"
        Else
            msg = msg & " and " & "VFD required for Rundown, standard drive required for TRG"
        End If
    End If
    
    If msg <> "" Then
        frmAlert.txtAlert.Text = msg
        frmAlert.Show
    End If
    
End Sub


Private Sub txtModelNo_Validate(Cancel As Boolean)
    Dim I As Integer
    Dim S As String

'    s = txtModelNo.Text
'    S = Replace(S, "-", "")
'    S = Replace(S, " ", "")
'    S = Replace(S, "/", "")

'    txtModelNo.Text = ""

'    For i = 1 To Len(s)
'        txtModelNo.Text = txtModelNo.Text & Mid(s, i, 1)
'    Next i
    txtModelNo_Change
  
End Sub

Private Sub txtNPSHFile_GotFocus()
    On Error GoTo FileCancel
    If LenB(txtNPSHFile.Text) <> 0 Then
        CommonDialog1.filename = txtNPSHFile.Text
    End If
    CommonDialog1.ShowOpen
    txtNPSHFile.Text = CommonDialog1.filename
    Exit Sub
FileCancel:
On Error GoTo 0
    CommonDialog1.CancelError = False
End Sub

Private Sub txtP1_Change()
    txtP2.Text = txtP1.Text
    txtP3.Text = txtP1.Text
End Sub

Private Sub txtPicturesFile_gotfocus()
    CommonDialog1.CancelError = True
    On Error GoTo FileCancel
    If LenB(txtPicturesFile.Text) <> 0 Then
        CommonDialog1.filename = txtPicturesFile.Text
    End If
    CommonDialog1.ShowOpen
    txtPicturesFile.Text = CommonDialog1.filename
    Exit Sub
FileCancel:
On Error GoTo 0
    CommonDialog1.CancelError = False
End Sub

Private Sub txtSN_Change()
    cmdFindPump.Default = True
End Sub

Private Sub txtTEMCFrontThrust_Change()
    CalculateTEMCForce
End Sub

Private Sub txtTEMCMomentArm_Change()
    CalculateTEMCForce
End Sub

Private Sub txtTEMCRearThrust_Change()
    CalculateTEMCForce
End Sub

Private Sub txtTEMCThrustRigPressure_Change()
    CalculateTEMCForce
End Sub

Private Sub txtTEMCViscosity_Change()
    'CalculateTEMCForce
End Sub



Private Sub txtV1_Change()
    txtV2.Text = txtV1.Text
    txtV3.Text = txtV1.Text
End Sub

Private Sub txtVibrationFile_gotfocus()
    On Error GoTo FileCancel
    If LenB(txtVibrationFile.Text) <> 0 Then
        CommonDialog1.filename = txtVibrationFile.Text
    End If
    CommonDialog1.ShowOpen
    txtVibrationFile.Text = CommonDialog1.filename
    Exit Sub
FileCancel:
On Error GoTo 0
    CommonDialog1.CancelError = False
End Sub
Private Sub ExportToExcel()

    Dim SaveFileName As String
    Dim WorkSheetName As String

    Dim I As Integer
    Dim iRowNo As Integer
    Dim sImp As String
    Dim ans As Integer

    Dim bCanShowSpeed As Boolean
    Dim CantShowReason As String

'close any running excel processes
    Dim objWMIService, colProcesses
    Set objWMIService = GetObject("winmgmts:")
    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name LIKE 'Excel%'")
    If colProcesses.Count > 0 Then
        Set xlApp = Excel.Application
    Else
        'use existing copy
'        Set xlApp = New Excel.Application
        Set xlApp = CreateObject("Excel.Application")
    End If


    CommonDialog1.CancelError = True        'in case the user
    On Error GoTo ErrHandler                '  chooses the cancel button

    'set up dialog box
    CommonDialog1.DialogTitle = "Open Excel Files"
    CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|"  'show Excel files
    CommonDialog1.InitDir = App.Path
'    CommonDialog1.InitDir = "C:\"    'in this directory
    CommonDialog1.ShowOpen                              'open the file selection dialog box

    If Dir(CommonDialog1.filename) = "" Then            'if the file name does not exist yet
        SaveFileName = CommonDialog1.filename           'get the name of the file
        If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
             xlApp.Workbooks.Close
        End If
        ' Create the Excel Workbook Object.
On Error GoTo 0
        Set xlBook = xlApp.Workbooks.Add                'add a workbook
        WorkSheetName = NewWorkBook                                     'do some stuff for the new workbook
        ActiveWorkbook.CheckCompatibility = False
        xlApp.ActiveWorkbook.SaveAs filename:=SaveFileName, _
                          FileFormat:=xlNormal                        'save the file
    Else                                                'the file name already exists
        SaveFileName = CommonDialog1.filename
        ' Create the Excel Workbook Object.
        If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
             xlApp.Workbooks.Close
        End If
        Set xlBook = xlApp.Workbooks.Open(SaveFileName)             'get the file name selected
        If GetWorksheetTabs(SaveFileName, WorkSheetName) = vbNo Then    'ask the user if he/she wants a new tab.
            MsgBox "File not overwritten.", vbOKOnly, "File not Opened"
            Exit Sub
        Else
        End If
    End If

On Error GoTo 0

    'see if we can export Speed and SG and if we can, ask user if s/he wants it
    'assume that we can show speed calcs

    bCanShowSpeed = False
'open the template and copy the data from the sheet
'  excel file resides in ParentDirectoryName + "\Polar SG&Visc Correction5.xls"
    'write the data to the spreadsheet
    With xlApp

    Dim xlTemplateName As String
    xlTemplateName = ParentDirectoryName & sSGandViscSpreadsheetTemplate
    Dim xlTemplate As Excel.Workbook
    Set xlTemplate = xlApp.Workbooks.Open(xlTemplateName)
    Dim TemplateWS As Excel.Worksheet
    Dim sheetName As String
    sheetName = xlTemplate.Sheets(1).Name
    xlTemplate.Sheets(1).Copy After:=xlBook.Sheets(WorkSheetName)

    xlTemplate.Close savechanges:=False

    Set xlTemplate = Nothing

    Application.DisplayAlerts = False
    ActiveWorkbook.Worksheets(WorkSheetName).Delete
    Application.DisplayAlerts = True
    ActiveWorkbook.Worksheets(sheetName).Name = WorkSheetName

    'WorkSheetName = sheetName

    'first see if there is an entry in CalculatedRPM table for this frame size and voltage.
    ' if there is, get the coefficients, else make the coefficients 0

        Dim ACoef As Double
        Dim BCoef As Double
        Dim CCoef As Double

        Dim qy As New ADODB.Command
        Dim rs As New ADODB.Recordset
        qy.ActiveConnection = cnPumpData
        Dim VoltageForLookup As Integer
        If cmbVoltage.List(cmbVoltage.ListIndex) = "380" And cmbFrequency.List(cmbFrequency.ListIndex) = "50 Hz" Then
            VoltageForLookup = 460
        ElseIf cmbVoltage.List(cmbVoltage.ListIndex) <> "380" Then
            VoltageForLookup = cmbVoltage.List(cmbVoltage.ListIndex)
        End If
        qy.CommandText = "SELECT * FROM CalculatedRPM WHERE FrameNumber = '" & txtTEMCFrameNumber.Text & _
                   "' AND Voltage = '" & VoltageForLookup & "'"

        rs.CursorLocation = adUseClient
        rs.CursorType = adOpenStatic

        rs.Open qy
        If rs.RecordCount = 0 Then
            ACoef = 0
            BCoef = 0
            CCoef = 0
            MsgBox ("Cannot find coefficient data for Frame Number " & txtTEMCFrameNumber.Text & _
                   " AND Voltage = " & cmbVoltage.List(cmbVoltage.ListIndex) & _
                   " AND Frequency = " & cmbFrequency.List(cmbFrequency.ListIndex))
        Else
            ACoef = rs.Fields("A")
            BCoef = rs.Fields("B")
            CCoef = rs.Fields("C")
        End If


    'write header data

        .Range("A2").Select
        .ActiveCell.FormulaR1C1 = "Serial Number"
        .Range("C2").Select
        .ActiveCell.FormulaR1C1 = txtSN

        .Range("F1").Select
        .ActiveCell.FormulaR1C1 = "Customer"
        .Range("H1").Select
        .ActiveCell.FormulaR1C1 = txtShpNo

        .Range("A3").Select
        .ActiveCell.FormulaR1C1 = "Model"
        .Range("C3").Select
        .ActiveCell.FormulaR1C1 = txtModelNo

        .Range("F2").Select
        .ActiveCell.FormulaR1C1 = "Sales Order"
        .Range("H2").Select
        .ActiveCell.FormulaR1C1 = txtSalesOrderNumber

        .Range("A9").Select
        .ActiveCell.FormulaR1C1 = "Design Flow"
        .Range("C9").Select
        .ActiveCell.FormulaR1C1 = Val(txtDesignFlow)

        .Range("A10").Select
        .ActiveCell.FormulaR1C1 = "Design Head"
        .Range("C10").Select
        .ActiveCell.FormulaR1C1 = Val(txtDesignTDH)

        .Range("P13").Select
        .ActiveCell.FormulaR1C1 = "Barometric Pressure"
        .Range("R13").Select
        .ActiveCell.FormulaR1C1 = Val(txtInHgDisplay)

        .Range("P11").Select
        .ActiveCell.FormulaR1C1 = "Suction Gage Height"
        .Range("R11").Select
        .ActiveCell.FormulaR1C1 = Val(txtSuctHeight)

        .Range("P12").Select
        .ActiveCell.FormulaR1C1 = "Discharge Gage Height"
        .Range("R12").Select
        .ActiveCell.FormulaR1C1 = Val(txtDischHeight)

        .Range("A1").Select
        .ActiveCell.FormulaR1C1 = "Run Date"
        .Range("C1").Select
        .ActiveCell.FormulaR1C1 = cmbTestDate.List(cmbTestDate.ListIndex)

        .Range("D10:E10").Select
        With xlApp.Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
            .WrapText = False
            .Orientation = 0
            .AddIndent = False
            .IndentLevel = 0
            .ShrinkToFit = False
            .ReadingOrder = xlContext
            .MergeCells = False
        End With
        xlApp.Selection.Merge

        'determine rpm

        Dim RPMvalue As String
        If Mid$(Me.txtTEMCFrameNumber.Text, 2, 1) = "1" Then
        '1 says 2 pole
            If Me.cmbFrequency.ListIndex = 0 Then
                '0 says 50Hz
                RPMvalue = "2900"
            ElseIf Me.cmbFrequency.ListIndex = 1 Then
                ' says 60Hz
                RPMvalue = "3450"
            Else
                'vfd or other, no rpm
                RPMvalue = ""
            End If
        Else
        '2 says 4 pole
            If Me.cmbFrequency.ListIndex = 0 Then
                '0 says 50Hz
                RPMvalue = "1450"
            ElseIf Me.cmbFrequency.ListIndex = 1 Then
                ' says 60Hz
                RPMvalue = "1750"
            Else
                'vfd or other, no rpm
                RPMvalue = ""
            End If
        End If

'        .Range("G1").Select
'        .ActiveCell.FormulaR1C1 = "RPM"
'        .Range("I1").Select
'        .ActiveCell.FormulaR1C1 = RPMvalue

        .Range("A5").Select
        .ActiveCell.FormulaR1C1 = "Sp Gravity"
        .Range("C5").Select
        .ActiveCell.FormulaR1C1 = txtSpGr

        .Range("A6").Select
        .ActiveCell.FormulaR1C1 = "Viscosity"
        .Range("C6").Select
        .ActiveCell.FormulaR1C1 = txtViscosity

        .Range("F4").Select
        .ActiveCell.FormulaR1C1 = "Motor"
        .Range("H4").Select
        .ActiveCell.FormulaR1C1 = txtTEMCFrameNumber.Text

        .Range("H12").Select
        .ActiveCell.FormulaR1C1 = Me.txtCustPONum.Text
        
        .Range("F5").Select
        .ActiveCell.FormulaR1C1 = "Voltage"
        .Range("H5").Select
        .ActiveCell.FormulaR1C1 = cmbVoltage.List(cmbVoltage.ListIndex)

        .Range("K6").Select
        .ActiveCell.FormulaR1C1 = "End Play"
        .Range("M6").Select
        .ActiveCell.FormulaR1C1 = Val(txtEndPlay)

        .Range("K7").Select
        .ActiveCell.FormulaR1C1 = "G-Gap"
        .Range("M7").Select
        .ActiveCell.FormulaR1C1 = txtGGap.Text
    
        .Range("A8").Select
        .ActiveCell.FormulaR1C1 = "Design Pressure"
        .Range("C8").Select
        Dim DesPress As String
        DesPress = cmbTEMCDesignPressure.List(cmbTEMCDesignPressure.ListIndex)
        Dim j As Integer
        j = InStrRev(DesPress, "-")
        .ActiveCell.FormulaR1C1 = Mid$(DesPress, j + 2)

'        .Range("G8").Select
'        .ActiveCell.FormulaR1C1 = "Stator Fill"
'        .Range("I8").Select
'        .ActiveCell.FormulaR1C1 = "Dry"

        .Range("K4").Select
        .ActiveCell.FormulaR1C1 = "Circulation Path"
        .Range("M4").Select
        .ActiveCell.FormulaR1C1 = cmbTEMCModel.List(cmbTEMCModel.ListIndex)

        .Range("M8").Select
        .ActiveCell.FormulaR1C1 = txtNPSHr.Text
        
        .Range("K1").Select
        .ActiveCell.FormulaR1C1 = "Impeller Dia"
        .Range("M1").Select
        

'        If LenB(txtImpTrim) <> 0 Then
'            .ActiveCell.FormulaR1C1 = Val(txtImpTrim)
'        Else
'            .ActiveCell.FormulaR1C1 = Val(txtImpellerDia)
'        End If
'
'        If chkTrimmed.value = 1 Then
'            If Val(txtImpTrim.Text) <> 0 Then
'                .ActiveCell.FormulaR1C1 = txtImpTrim
'            Else
'                .ActiveCell.FormulaR1C1 = txtImpellerDia
'            End If
'        Else
'            .ActiveCell.FormulaR1C1 = txtImpellerDia
'        End If

        If chkTrimmed.value = 1 Then
            If Val(txtImpTrim.Text) <> 0 Then
                .ActiveCell.FormulaR1C1 = Val(txtImpTrim.Text)
            Else
                .ActiveCell.FormulaR1C1 = Val(txtImpellerDia.Text)
            End If
        Else
            .ActiveCell.FormulaR1C1 = Val(txtImpellerDia.Text)
        End If



'        .Range("K1").Select
'        .ActiveCell.FormulaR1C1 = "KW Mult"
'        .Range("N1").Select
'        .ActiveCell.FormulaR1C1 = Val(txtKWMult)

'        .Range("K2").Select
'        .ActiveCell.FormulaR1C1 = "HD Cor"
'        .Range("N2").Select
'        If Val(txtHDCor) = 0 Then
'            .ActiveCell.FormulaR1C1 = 0
'        Else
'            .ActiveCell.FormulaR1C1 = Val(txtHDCor)
'        End If

        .Range("P9").Select
        .ActiveCell.FormulaR1C1 = "Suction Dia"
        .Range("R9").Select
        .ActiveCell.FormulaR1C1 = cmbSuctDia.List(cmbSuctDia.ListIndex)

        .Range("P10").Select
        .ActiveCell.FormulaR1C1 = "Discharge Dia"
        .Range("R10").Select
        .ActiveCell.FormulaR1C1 = cmbDischDia.List(cmbDischDia.ListIndex)

        .Range("A11").Select
        .ActiveCell.FormulaR1C1 = "Test Spec"
        .Range("C11").Select
        .ActiveCell.FormulaR1C1 = cmbTestSpec.List(cmbTestSpec.ListIndex)

        .Range("K3").Select
        .ActiveCell.FormulaR1C1 = "Impeller Feathered"
        .Range("M3").Select
        If chkFeathered.value = 1 Then
            .ActiveCell.FormulaR1C1 = "Yes"
        Else
            .ActiveCell.FormulaR1C1 = "No"
        End If

        .Range("K2").Select
        .ActiveCell.FormulaR1C1 = "Disch Orifice"
        .Range("M2").Select
        If chkOrifice.value = 1 Then
            .ActiveCell.FormulaR1C1 = Val(txtOrifice)
        Else
            .ActiveCell.FormulaR1C1 = "None"
        End If


        .Range("K5").Select
        .ActiveCell.FormulaR1C1 = "Circulation Orifice"
        .Range("M5").Select
        If chkCircOrifice.value = 1 Then
            .ActiveCell.FormulaR1C1 = Val(txtCircOrifice)
        Else
            .ActiveCell.FormulaR1C1 = "None"
        End If

        .Range("A13").Select
        .ActiveCell.FormulaR1C1 = "Other Mods"
        .Range("C13").Select
        .ActiveCell.FormulaR1C1 = txtOtherMods

        .Range("A14").Select
        .ActiveCell.FormulaR1C1 = "Remarks"
        .Range("C14").Select
        .ActiveCell.FormulaR1C1 = txtRemarks

        .Range("A15").Select
        .ActiveCell.FormulaR1C1 = "Test Setup Remarks"
        .Range("C15").Select
        .ActiveCell.FormulaR1C1 = txtTestSetupRemarks

        .Range("P1").Select
        .ActiveCell.FormulaR1C1 = "Suct ID"
        .Range("R1").Select
        .ActiveCell.FormulaR1C1 = cmbSuctionPressureTransducer.List(cmbSuctionPressureTransducer.ListIndex)

        .Range("P2").Select
        .ActiveCell.FormulaR1C1 = "Disch ID"
        .Range("R2").Select
        .ActiveCell.FormulaR1C1 = cmbDischargePressureTransducer.List(cmbDischargePressureTransducer.ListIndex)

        .Range("P3").Select
        .ActiveCell.FormulaR1C1 = "Temp ID"
        .Range("R3").Select
        .ActiveCell.FormulaR1C1 = cmbTemperatureTransducer.List(cmbTemperatureTransducer.ListIndex)

        .Range("P4").Select
        .ActiveCell.FormulaR1C1 = "Circ Flow ID"
        .Range("R4").Select
        .ActiveCell.FormulaR1C1 = cmbCirculationFlowMeter.List(cmbCirculationFlowMeter.ListIndex)

        .Range("P5").Select
        .ActiveCell.FormulaR1C1 = "Flow ID"
        .Range("R5").Select
        .ActiveCell.FormulaR1C1 = cmbFlowMeter.List(cmbFlowMeter.ListIndex)

        .Range("P6").Select
        .ActiveCell.FormulaR1C1 = "Analyzer ID"
        .Range("R6").Select
        .ActiveCell.FormulaR1C1 = cmbAnalyzerNo.List(cmbAnalyzerNo.ListIndex)

        .Range("P7").Select
        .ActiveCell.FormulaR1C1 = "Loop ID"
        .Range("R7").Select
        .ActiveCell.FormulaR1C1 = cmbLoopNumber.List(cmbLoopNumber.ListIndex)

        .Range("A4").Select
        .ActiveCell.FormulaR1C1 = "Fluid"
        .Range("C4").Select
        .ActiveCell.FormulaR1C1 = txtLiquid.Text

        .Range("F3").Select
        .ActiveCell.FormulaR1C1 = "Cust PN"
        .Range("H3").Select
'        .ActiveCell.FormulaR1C1 = txtRMA.Text
        If rsPumpData.Fields("RVSPartNo") <> "" Then
            .ActiveCell.FormulaR1C1 = rsPumpData.Fields("RVSPartNo")
        End If
        If rsPumpData.Fields("CustPN") <> "" Then
            .ActiveCell.FormulaR1C1 = rsPumpData.Fields("CustPN")
        End If
        
        .Range("A7").Select
        .ActiveCell.FormulaR1C1 = "Temperature"
        .Range("C7").Select
        .ActiveCell.FormulaR1C1 = txtLiquidTemperature.Text

        .Range("F6").Select
        .ActiveCell.FormulaR1C1 = "Frequency"
        .Range("H6").Select
        If UCase(cmbFrequency.List(cmbFrequency.ListIndex)) = "VFD" Then
            .ActiveCell.FormulaR1C1 = Val(Me.txtVFDFreq)
        Else
            .ActiveCell.FormulaR1C1 = Val(cmbFrequency.List(cmbFrequency.ListIndex))
        End If
'        .Range("K2").Select
'        .ActiveCell.FormulaR1C1 = "Disch Orifice"
'        .Range("M2").Select
'        .ActiveCell.FormulaR1C1 = txtOrifice.Text

'        .Range("K12").Select
'        .ActiveCell.FormulaR1C1 = "Flow Orifice"
'        .Range("L12").Select
'        .ActiveCell.FormulaR1C1 = txtCircOrifice.Text

        .Range("P8").Select
        .ActiveCell.FormulaR1C1 = "PLC No"
        .Range("R8").Select
        .ActiveCell.FormulaR1C1 = cmbPLCNo.List(cmbPLCNo.ListIndex)

        .Range("F7").Select
        .ActiveCell.FormulaR1C1 = "Phases"
        .Range("H7").Select
        .ActiveCell.FormulaR1C1 = txtNoPhases.Text

        .Range("F8").Select
        .ActiveCell.FormulaR1C1 = "Poles"
        .Range("H8").Select
        .ActiveCell.FormulaR1C1 = 2 * Val(Left$(Right$(txtTEMCFrameNumber.Text, 2), 1))

        .Range("F9").Select
        .ActiveCell.FormulaR1C1 = "Rated Current"
        .Range("H9").Select
        .ActiveCell.FormulaR1C1 = txtAmps.Text

        .Range("F10").Select
        .ActiveCell.FormulaR1C1 = "Rated Input Power"
        .Range("H10").Select
        .ActiveCell.FormulaR1C1 = txtRatedInputPower.Text

        .Range("F11").Select
        .ActiveCell.FormulaR1C1 = "Insulation Class"
        .Range("H11").Select
        .ActiveCell.FormulaR1C1 = txtThermalClass.Text

'        .Range("P8").Select
'        .ActiveCell.FormulaR1C1 = "Tach ID"
'        .Range("R8").Select
'        .ActiveCell.FormulaR1C1 = cmbTachID.List(cmbTachID.ListIndex)
'
'        .Range("P9").Select
'        .ActiveCell.FormulaR1C1 = "Orifice ID"
'        .Range("R9").Select
'        '.ActiveCell.FormulaR1C1 = cmbOrificeNumber.List(cmbOrificeNumber.ListIndex)

    'list the columns starting at row17

        .Range("A17").Select
        .ActiveCell.FormulaR1C1 = "Flow"
        .Range("A18").Select
        .ActiveCell.FormulaR1C1 = "(GPM)"

        .Range("B17").Select
        .ActiveCell.FormulaR1C1 = "TDH"
        .Range("B18").Select
        .ActiveCell.FormulaR1C1 = "(Ft)"

        .Range("C17").Select
        .ActiveCell.FormulaR1C1 = "KW"

        .Range("D17").Select
        .ActiveCell.FormulaR1C1 = "Ave"
        .Range("D18").Select
        .ActiveCell.FormulaR1C1 = "Volts"

        .Range("E17").Select
        .ActiveCell.FormulaR1C1 = "Ave"
        .Range("E18").Select
        .ActiveCell.FormulaR1C1 = "Amps"

        .Range("F17").Select
        .ActiveCell.FormulaR1C1 = "Power"
        .Range("F18").Select
        .ActiveCell.FormulaR1C1 = "Factor"

        .Range("G17").Select
        .ActiveCell.FormulaR1C1 = "Overall"
        .Range("G18").Select
        .ActiveCell.FormulaR1C1 = "Eff"

        .Range("H17").Select
        .ActiveCell.FormulaR1C1 = "Measured"
        .Range("H18").Select
        .ActiveCell.FormulaR1C1 = "RPM"

        .Range("I17").Select
        .ActiveCell.FormulaR1C1 = "Calculated"
        .Range("I18").Select
        .ActiveCell.FormulaR1C1 = "RPM"

        .Range("J17").Select
        .ActiveCell.FormulaR1C1 = "Suction"
        .Range("J18").Select
        .ActiveCell.FormulaR1C1 = "Temp(F)"

        .Range("K17").Select
        .ActiveCell.FormulaR1C1 = "Disch"
        .Range("K18").Select
        .ActiveCell.FormulaR1C1 = "Pressure"

        .Range("L17").Select
        .ActiveCell.FormulaR1C1 = "Suction"
        .Range("L18").Select
        .ActiveCell.FormulaR1C1 = "Pressure"

        .Range("M17").Select
        .ActiveCell.FormulaR1C1 = "Vel"
        .Range("M18").Select
        .ActiveCell.FormulaR1C1 = "Head"

        .Range("N17").Select
        .ActiveCell.FormulaR1C1 = "Axial"
        .Range("N18").Select
        .ActiveCell.FormulaR1C1 = "Position"

        .Range("O17").Select
        .ActiveCell.FormulaR1C1 = "Pct of"
        .Range("O18").Select
        .ActiveCell.FormulaR1C1 = "End Play"

        .Range("P17").Select
        .ActiveCell.FormulaR1C1 = "Hydraulic"
        .Range("P18").Select
        .ActiveCell.FormulaR1C1 = "Efficiency"

'        .Range("P17").Select
'        .ActiveCell.FormulaR1C1 = "Circ"
'        .Range("P18").Select
'        .ActiveCell.FormulaR1C1 = "Flow"

        .Range("Q17").Select
        .ActiveCell.FormulaR1C1 = "Motor"
        .Range("Q18").Select
        .ActiveCell.FormulaR1C1 = "Efficiency"

        .Range("S17").Select
        .ActiveCell.FormulaR1C1 = "NPSHa"

        .Range("T17").Select
        .ActiveCell.FormulaR1C1 = "Phase 1"
        .Range("T18").Select
        .ActiveCell.FormulaR1C1 = "Current"

        .Range("U17").Select
        .ActiveCell.FormulaR1C1 = "Phase 2"
        .Range("U18").Select
        .ActiveCell.FormulaR1C1 = "Current"

        .Range("V17").Select
        .ActiveCell.FormulaR1C1 = "Phase 3"
        .Range("V18").Select
        .ActiveCell.FormulaR1C1 = "Current"

        .Range("W17").Select
        .ActiveCell.FormulaR1C1 = "Phase 1"
        .Range("W18").Select
        .ActiveCell.FormulaR1C1 = "Voltage"

        .Range("X17").Select
        .ActiveCell.FormulaR1C1 = "Phase 2"
        .Range("X18").Select
        .ActiveCell.FormulaR1C1 = "Voltage"

        .Range("Y17").Select
        .ActiveCell.FormulaR1C1 = "Phase 3"
        .Range("Y18").Select
        .ActiveCell.FormulaR1C1 = "Voltage"

        .Range("Z17").Select
        .ActiveCell.FormulaR1C1 = "'" & txtTitle(20).Text

        .Range("Z18").Select
        .ActiveCell.FormulaR1C1 = "'" & txtTitle(21).Text

        .Range("AA17").Select
        .ActiveCell.FormulaR1C1 = "'" & txtTitle(22).Text

        .Range("AA18").Select
        .ActiveCell.FormulaR1C1 = "'" & txtTitle(23).Text

        .Range("AB17").Select
        .ActiveCell.FormulaR1C1 = "'" & txtTitle(24).Text

        .Range("AB18").Select
        .ActiveCell.FormulaR1C1 = "'" & txtTitle(25).Text

        .Range("AC17").Select
        .ActiveCell.FormulaR1C1 = "HR"

        .Range("AC18").Select
        .ActiveCell.FormulaR1C1 = "(ft)"

        .Range("AD17").Select
        .ActiveCell.FormulaR1C1 = "'" & txtTitle(26).Text

        .Range("AD18").Select
        .ActiveCell.FormulaR1C1 = "'" & txtTitle(27).Text

        .Range("AE17").Select
        .ActiveCell.FormulaR1C1 = "TRG"
        .Range("AE18").Select
        .ActiveCell.FormulaR1C1 = "Position"

        .Range("AF17").Select
        .ActiveCell.FormulaR1C1 = "Thrust"

        .Range("AG17").Select
        .ActiveCell.FormulaR1C1 = "F/R"

        .Range("AH17").Select
        .ActiveCell.FormulaR1C1 = "Moment"
        .Range("AH18").Select
        .ActiveCell.FormulaR1C1 = "Arm"

        .Range("AI17").Select
        .ActiveCell.FormulaR1C1 = "Rig"
        .Range("AI18").Select
        .ActiveCell.FormulaR1C1 = "Pressure"

'        .Range("AI17").Select
'        .ActiveCell.FormulaR1C1 = "Viscosity"

        .Range("AJ19").Select
        .ActiveCell.FormulaR1C1 = "Rear"
        .Range("AJ18").Select
        .ActiveCell.FormulaR1C1 = "Force"

        .Range("AK17").Select
        .ActiveCell.FormulaR1C1 = "PV"

        .Range("R17").Select
        .ActiveCell.FormulaR1C1 = "Shaft"
        .Range("R18").Select
        .ActiveCell.FormulaR1C1 = "Power"

'        .Range("AM17").Select
'        .ActiveCell.FormulaR1C1 = "Pct Full"
'        .Range("AM18").Select
'        .ActiveCell.FormulaR1C1 = "Scale"

        .Range("AL17").Select
        .ActiveCell.FormulaR1C1 = "NPSHr"

        .Range("AM17").Select
        .ActiveCell.FormulaR1C1 = "Remarks"




        'now output the data

        iRowNo = 20

        rsEff.MoveFirst
        For I = 1 To frmPLCData.UpDown2.value
            .Range("A" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("Flow")

            .Range("B" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("TDH")

            .Range("C" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("KW")

            .Range("D" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("Volts")

            .Range("E" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("Amps")

            .Range("F" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("PowerFactor")

            .Range("G" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("OverallEfficiency")

            .Range("H" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("RPM")

            .Range("I" & iRowNo).Select
            'use the coefficients from above to calculate rpm
            Dim f As Double
            f = .Range("H6").value
            .ActiveCell.FormulaR1C1 = (Val(f) / 60) * (ACoef * (rsEff.Fields("KW")) ^ 2 + BCoef * (rsEff.Fields("KW")) + CCoef)

            .Range("J" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("Temperature")

            .Range("K" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("DischPress")

            .Range("L" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("SuctPress")

            .Range("M" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("VelocityHead")

            .Range("N" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("Pos")

            .Range("O" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = Format((100 * rsEff.Fields("Pos") / Val(txtEndPlay)), "00.0")

            .Range("P" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("HydraulicEfficiency")

'            .Range("P" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("CircFlow")

            .Range("Q" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("MotorEfficiency")

            .Range("S" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("NPSHa")

            .Range("T" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentA")

            .Range("U" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentB")

            .Range("V" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentC")

            .Range("W" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageA")

            .Range("X" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageB")

            .Range("Y" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageC")

'            .Range("Y" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("TC1")
'
'            .Range("Z" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("TC2")
'
'            .Range("AA" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("TC3")
'
'            .Range("AB" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("TC4")

            .Range("Z" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("CircFlow")

            .Range("AA" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("RBHTemp")

            .Range("AB" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("RBHPress")

            .Range("AC" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = (rsEff.Fields("RBHPress") - rsEff.Fields("SuctPress")) * 2.31

            .Range("AD" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("AI4")

            .Range("AE" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCTRG")

            .Range("AF" & iRowNo).Select
            If rsEff.Fields("TEMCFrontThrust") = 0 Then
                If rsEff.Fields("TEMCRearThrust") = 0 Then
                    .ActiveCell.FormulaR1C1 = " "
                    .Range("AG" & iRowNo).Select
                    .ActiveCell.FormulaR1C1 = " "
                Else
                    .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCRearThrust")
                    .Range("AG" & iRowNo).Select
                    .ActiveCell.FormulaR1C1 = "R"
                End If
            Else
                .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCFrontThrust")
                .Range("AG" & iRowNo).Select
                .ActiveCell.FormulaR1C1 = "F"
            End If

            .Range("AH" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCMomentArm")

            .Range("AI" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCThrustRigPressure")

'            .Range("AJ" & iRowNo).Select
'            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCViscosity")

            .Range("AJ" & iRowNo).Select
            If rsEff.Fields("TEMCForceDirection") = "F" Then
                .ActiveCell.FormulaR1C1 = -rsEff.Fields("TEMCCalculatedForce")
            Else
                .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCCalculatedForce")
            End If

            .Range("AK" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCPV")

            .Range("R" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("KW") * rsEff.Fields("MotorEfficiency") / 100

            .Range("AL" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("NPSHr")
            
'            If RatedKW = 999 Then
'                .ActiveCell.FormulaR1C1 = ""
'            Else
'                .ActiveCell.FormulaR1C1 = (rsEff.Fields("KW") * rsEff.Fields("MotorEfficiency")) / (1 * RatedKW)
'            End If

            .Range("AM" & iRowNo).Select
            .ActiveCell.FormulaR1C1 = rsEff.Fields("Remarks")


            rsEff.MoveNext
            iRowNo = iRowNo + 1
        Next I

        .Range("A20:AS30").Select
        .Selection.NumberFormat = "0.00"
        
        'format AxPos to 3 dp
        .Range("N20:N30").Select
        .Selection.NumberFormat = "0.000"
        
        'format %EndPlay to 1 dp
        .Range("O20:O30").Select
        .Selection.NumberFormat = "0.0"

    'set up formulas to calculate BEP
    '  first, plot 2nd order polynomial for flow vs hydraulic efficiency
    '  the formulas for doing that are in E68, F68 and G68
    '  only want the formulas to point to the number of points in the test data, so use frmPLCData.CWNumEdit2.value
    '
    Dim AColumnRow As String
    Dim PColumnRow As String

    AColumnRow = "A" & Trim(str(19 + frmPLCData.UpDown2.value))
    PColumnRow = "P" & Trim(str(19 + frmPLCData.UpDown2.value))

        .Range("E68").Select
        .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1)"

        .Range("F68").Select
        .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1,2)"

        .Range("G68").Select
        .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1,3)"

    'export balance holes
    If boGotBalanceHoles Then
        If rsBalanceHoles.State = adStateClosed Then
            rsBalanceHoles.ActiveConnection = cnPumpData
            rsBalanceHoles.Open
        End If 'rsBalanceHoles.State = adStateClosed

        If rsBalanceHoles.RecordCount <> 0 Then

            .Range("K9:N9").Merge
            .Range("K9:N9").Formula = "Balance Hole Data"
            .Range("K9:N9").HorizontalAlignment = xlCenter

            .Range("K10").Select
            .ActiveCell.Formula = "Date"

            .Range("L10").Select
            .ActiveCell.Formula = "Number"

            .Range("M10").Select
            .ActiveCell.Formula = "Diameter"

            .Range("N10").Select
            .ActiveCell.Formula = "Bolt Circle"

            iRowNo = 11

            If rsBalanceHoles.RecordCount > 3 Then
                For I = 1 To rsBalanceHoles.RecordCount - 3
                    Rows("13:13").Select
                    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
                Next I
            End If

            rsBalanceHoles.MoveFirst
            For I = 1 To rsBalanceHoles.RecordCount

                .Range("K" & iRowNo).Select
                .ActiveCell.Formula = rsBalanceHoles.Fields("Date")
                .ActiveCell.NumberFormat = "m/d/yy h:mm AM/PM;@"
                .Range("L" & iRowNo).Select
                .ActiveCell = rsBalanceHoles.Fields("Number")
                .ActiveCell.NumberFormat = "0"
                .Range("M" & iRowNo).Select
                If IsNumeric(rsBalanceHoles.Fields("Diameter1")) Then
                    .ActiveCell = Val(rsBalanceHoles.Fields("Diameter1"))
                    .ActiveCell.NumberFormat = "0.0000"
                Else
                    .ActiveCell = rsBalanceHoles.Fields("Diameter1")
                End If

                .Range("N" & iRowNo).Select
                If IsNumeric(rsBalanceHoles.Fields("BoltCircle1")) Then
                    .ActiveCell = Val(rsBalanceHoles.Fields("BoltCircle1"))
                    .ActiveCell.NumberFormat = "0.0000"
                Else
                    .ActiveCell = rsBalanceHoles.Fields("BoltCircle1")
                End If

                rsBalanceHoles.MoveNext
                iRowNo = iRowNo + 1
            Next I
            .Range("K10:N" & iRowNo - 1).Select
            With .Selection.Interior
                .ColorIndex = 34
                .Pattern = xlSolid
            End With
        End If 'rsBalanceHoles.RecordCount <> 0
    End If ' boGotBalanceHoles

    'plot graphs

    Dim SeriesName As String
    Dim XVals As String
    Dim YVals As String
    Dim RowNo As Long
    Dim RowStr As String
    Dim LastPoint As Integer
    Dim LineType As String
    Dim AxisGroup As Integer
    Dim LabelPos As Integer
    Dim LineColor As Long

        .ActiveSheet.ChartObjects("HydRepChart").Activate
        Dim S As Series
        'For Each S In ActiveChart.SeriesCollection
        '    S.Delete
        'Next S

       'determine how many rows of data we have

'        Range("J86", "J93").Select
'        With Application.WorksheetFunction
'            LastPoint = .Match(.Max(Selection), Selection)
'            RowNo = LastPoint + 85
'        End With
'        RowStr = Trim(str(RowNo))

        'find max values to scale chart

        'first TDH
        Dim aq As Double
        Range("AQ56", "AQ71").Select
        aq = .Max(Selection)
        Dim ax As Double
        Range("AX56", "AX71").Select
        ax = .Max(Selection)
        
        'then current (as and az)
        Dim at As Double
        Range("AS56", "AS71").Select
        at = .Max(Selection)
        Dim ba As Double
        Range("AZ56", "AZ71").Select
        ba = .Max(Selection)

        Dim CurrentScaleMax As Integer
        Dim TDHScaleMax As Integer

        Dim MaxTDH As Integer
        With Application.WorksheetFunction
            If aq > ax Then
                MaxTDH = .Ceiling(aq, 25)
            Else
                MaxTDH = .Ceiling(ax, 25)
            End If
        End With

        Dim MaxCurrent As Integer
        With Application.WorksheetFunction
            If at > ba Then
                Select Case at
                    Case Is <= 5
                        CurrentScaleMax = 5
                        
                    Case Is <= 10
                        CurrentScaleMax = 10
                        
                    Case Else
                        CurrentScaleMax = 25
                End Select
                
                MaxCurrent = .Ceiling(at, CurrentScaleMax)
            Else
               Select Case ba
                    Case Is <= 5
                        CurrentScaleMax = 5
                        
                    Case Is <= 10
                        CurrentScaleMax = 10
                        
                    Case Else
                        CurrentScaleMax = 25
                End Select
                
                MaxCurrent = .Ceiling(ba, CurrentScaleMax)
            End If
        End With

        ActiveSheet.ChartObjects("HydRepChart").Activate
         Dim ShtName As String
         ShtName = "'" & ActiveSheet.Name & "'"

        Dim skipSeries As Boolean
        RowStr = 56 + 15
         For I = 1 To 8
            skipSeries = False
             Select Case I
                 Case 1
                     SeriesName = "=""TDH"""
                     XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
                     YVals = "=" & ShtName & "!$AQ$56:$AQ$" & RowStr
                     LineType = msoLineSolid
                     AxisGroup = 1
                     LabelPos = xlLabelPositionRight
                     LineColor = vbBlue
                 
                 Case 2
                     SeriesName = "=""Input Power"""
                     XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
                     YVals = "=" & ShtName & "!$AR$56:$AR$" & RowStr
                     LineType = msoLineSolid
                     AxisGroup = 2
                     LabelPos = xlLabelPositionRight
                     LineColor = vbRed

                 Case 3
                     SeriesName = "=""Current"""
                     XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
                     YVals = "=" & ShtName & "!$AS$56:$AS$" & RowStr
                     LineType = msoLineSolid
                     AxisGroup = 2
                     LabelPos = xlLabelPositionRight
                     LineColor = vbGreen
                 
                 Case 4
                      skipSeries = True
'                     SeriesName = "=""Overall Eff"""
'                     XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
'                     YVals = "=" & ShtName & "!$AT$56:$AT$" & RowStr
'                     LineType = msoLineSolid
'                     AxisGroup = 2
'                     LabelPos = xlLabelPositionRight
'                     LineColor = vbCyan

                 Case 5
                     SeriesName = "=""TDH (Adj)"""
                     XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
                     YVals = "=" & ShtName & "!$AX$56:$AX$" & RowStr
                     LineType = msoLineDash
                     AxisGroup = 1
                     LabelPos = xlLabelPositionBelow
                     LineColor = vbBlue

                 Case 6
                     SeriesName = "=""Input Power (Adj)"""
                     XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
                     YVals = "=" & ShtName & "!$AY$56:$AY$" & RowStr
                     LineType = msoLineDash
                     AxisGroup = 2
                     LabelPos = xlLabelPositionBelow
                     LineColor = vbRed

                 Case 7
                     SeriesName = "=""Current (Adj)"""
                     XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
                     YVals = "=" & ShtName & "!$AZ$56:$AZ$" & RowStr
                     LineType = msoLineDash
                     AxisGroup = 2
                     LabelPos = xlLabelPositionBelow
                     LineColor = vbGreen

                 Case 8
                    skipSeries = True
'                     SeriesName = "=""Overall Eff (Adj)"""
'                     XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
'                     YVals = "=" & ShtName & "!$BA$56:$BA$" & RowStr
'                     LineType = msoLineDash
'                     AxisGroup = 2
'                     LabelPos = xlLabelPositionBelow
'                     LineColor = vbCyan

            End Select
            
            If Not skipSeries Then
                LastPoint = 16
                ActiveChart.SeriesCollection.NewSeries
                ActiveChart.SeriesCollection(I).Name = SeriesName
                ActiveChart.SeriesCollection(I).XValues = XVals
                ActiveChart.SeriesCollection(I).Values = YVals
                ActiveChart.SeriesCollection(I).Select
                ActiveChart.SeriesCollection(I).Points(LastPoint).Select
                ActiveChart.SeriesCollection(I).Points(LastPoint).ApplyDataLabels
                ActiveChart.SeriesCollection(I).Points(LastPoint).DataLabel.Select
                If I < 5 Then
                    Selection.ShowSeriesName = True
                    Selection.Position = LabelPos
                Else
                    Selection.ShowSeriesName = False
                End If
                Selection.ShowValue = False
                ActiveChart.SeriesCollection(I).ChartType = xlXYScatterSmoothNoMarkers
                ActiveChart.SeriesCollection(I).Select
                With Selection.Format.line
                    .Visible = msoTrue
                    .DashStyle = LineType
                    .ForeColor.RGB = LineColor
                End With
    
                
                ActiveChart.SeriesCollection(I).AxisGroup = AxisGroup
                ActiveChart.SeriesCollection(I).DataLabels.Font.Size = 8
                ActiveChart.SeriesCollection(I).DataLabels.Font.Name = "Arial"
            End If  'not skip series
        Next I

        'show design point
        SeriesName = "=""Design Point"""
        XVals = "=" & ShtName & "!$L$63"
        YVals = "=" & ShtName & "!$L$64"
        LineType = msoLineSolid
        AxisGroup = 1
        ActiveChart.SeriesCollection.NewSeries
        ActiveChart.SeriesCollection(I).Name = SeriesName
        ActiveChart.SeriesCollection(I).XValues = XVals
        ActiveChart.SeriesCollection(I).Values = YVals
        ActiveChart.SeriesCollection(I).Select
       
        Selection.MarkerStyle = 4
        Selection.MarkerSize = 7
        With Selection.Format.line
            .Visible = msoTrue
            .Weight = 2.25
            .ForeColor.RGB = vbBlack
        End With


        ActiveChart.Axes(xlValue).Select
        ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
        ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True

        ActiveChart.Axes(xlValue).MaximumScale = MaxTDH
        ActiveChart.Axes(xlValue).MinimumScale = 0
        ActiveChart.Axes(xlValue).MajorUnit = Int(MaxTDH / 5)
        Selection.TickLabels.NumberFormat = "0"

        ActiveChart.Axes(xlValue, xlSecondary).Select
        ActiveChart.Axes(xlValue, xlSecondary).MinimumScaleIsAuto = True
        ActiveChart.Axes(xlValue, xlSecondary).MaximumScaleIsAuto = True
        
        ActiveChart.Axes(xlValue, xlSecondary).MaximumScale = MaxCurrent
        ActiveChart.Axes(xlValue, xlSecondary).MinimumScale = 0
        ActiveChart.Axes(xlValue, xlSecondary).MajorUnit = Int(MaxCurrent / 5)
        Selection.TickLabels.NumberFormat = "0"

        ActiveChart.Axes(xlValue, xlSecondary).HasTitle = True
        ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = "Input Power (kW)-Current (A)"
'        ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = "Input Power (kW)-Current (A)-Overall Efficiency (%)"
        ActiveChart.SetElement (msoElementSecondaryValueAxisTitleRotated)
        'ActiveSheet.PageSetup.PrintArea = "$CA$1:$CI$50"
        
        Range("A1").Select
     
        'delete all macros in the excel file
        
        ' Declare variables to access the macros in the workbook.
        Dim objProject As VBIDE.VBProject
        Dim objComponent As VBIDE.VBComponent
        Dim objCode As VBIDE.CodeModule
     
        ' Get the project details in the workbook.
        Set objProject = xlBook.VBProject
    
        ' Iterate through each component in the project.
        For Each objComponent In objProject.VBComponents
        
            ' Delete code modules
            Set objCode = objComponent.CodeModule
            objCode.DeleteLines 1, objCode.CountOfLines
                        
            Set objCode = Nothing
            Set objComponent = Nothing
        Next
    
        Set objProject = Nothing
          
     
        xlApp.Visible = True                    'show the sheet
        
        xlApp.VBE.ActiveVBProject.VBComponents.Import ParentDirectoryName & sSaveFileMacroFile
        xlApp.Run "AssignButton"
    End With

'    Exit Sub

ErrHandler:
    'User pressed the Cancel button

    On Error GoTo notopen
    If Not xlApp.ActiveWorkbook Is Nothing Then
        ActiveWorkbook.CheckCompatibility = False
        xlApp.ActiveWorkbook.Save               'save the workbook
        'xlApp.ActiveWorkbook.Close

    End If

notopen:

'    xlApp.Application.Quit

'    xlApp.Quit
'    Set xlApp = Nothing

'    If CommonDialog1.filename <> "" Then
'        MsgBox CommonDialog1.filename & " has been written.", vbOKOnly, "File Opened"
'    End If

On Error GoTo 0

    Exit Sub
End Sub

Function GetWorksheetTabs(filename As String, WorkSheetName As String)

    'see what worksheet tabs alread exist in the excel worksheet

    Dim intSheets As Integer    'number of sheets in the workbook
    Dim I As Integer
    Dim S As String
    Dim ans As Integer
    Dim NameOK As Boolean

    intSheets = xlApp.Worksheets.Count      'how many sheets are there?

    'define a crlf string
    S = vbCrLf

    For I = 1 To intSheets
        S = S & xlApp.Worksheets(I).Name & vbCrLf   'add in the worksheet name
    Next I

    'tell the user the names so far and ask if he/she wants to add another
    ans = MsgBox("You have the following Worksheet Names in " & filename & ": " & S & "Do you want to add another sheet to this file?", vbYesNo, "Sheets in Excel File")

    'get the answer
    If ans = vbNo Then
        GetWorksheetTabs = vbNo     'set up flag for when we return to the calling subroutine
        Exit Function
    End If

    'get worksheet name from user and check to see that it's not already used

    NameOK = False  'start assuming that the name is bad

    While Not NameOK    'as long as it's bad, stay in this loop
        WorkSheetName = InputBox("Enter Worksheet Name for this run.")  'ask for name

        If WorkSheetName = "" Then      'if we get a nul return or user presses cancel
            GetWorksheetTabs = vbNo
            Exit Function
        End If

        For I = 1 To xlApp.Worksheets.Count     'go through all of the existing sheets
            If WorkSheetName = xlApp.Worksheets(I).Name Then        'if the names are the same
                MsgBox "The name " & WorkSheetName & " already exists for a Worksheet.  Please try again.", vbOKOnly, "Bad Worksheet Name"  'tell the user
                NameOK = False
                Exit For
            End If
            NameOK = True       'if we make it thru say the name is ok
        Next I
    Wend

    xlApp.Worksheets.Add , xlApp.Worksheets(xlApp.Worksheets.Count)     'add a worksheer
    xlApp.Worksheets(xlApp.Worksheets.Count).Name = WorkSheetName       'give it the desired name
    GetWorksheetTabs = vbYes                                            'say that the results were ok
  
End Function
Function NewWorkBook() As String

    Dim WorkSheetName As String

    'we've just added a new workbook, delete sheet1, sheet2, etc
    xlApp.DisplayAlerts = False
    While xlApp.Worksheets.Count > 1
        xlApp.Worksheets(1).Delete          'delete the sheet
    Wend
    xlApp.DisplayAlerts = True

    WorkSheetName = InputBox("Enter Title Worksheet Name for this run.")    'get the desired name
    xlApp.Worksheets(1).Name = WorkSheetName    'and name the sheet

    NewWorkBook = WorkSheetName
  
End Function

Private Sub CalibrateSoftware()
        frmCalibrate.Show
        'Calibrating = True
  
End Sub

Function ParseTEMCModelNo(cmbComboName As ComboBox, ltr As String)
    Dim I As Integer
    Dim iStart As Integer
    Dim iStop As Integer
    Dim strCompare As String

    For I = 0 To cmbComboName.ListCount - 1                     'go through the combobox entries
        iStart = InStr(1, cmbComboName.List(I), "[")
        iStop = InStr(1, cmbComboName.List(I), "]")
        strCompare = Mid$(cmbComboName.List(I), iStart + 1, iStop - iStart - 1)
        If UCase(strCompare) = UCase(ltr) Then   'see when we find the desired index number
            cmbComboName.ListIndex = I                                              'if we do, set the combo box
            Exit For                                            'and we're done
        End If
'        cmbComboName.ListIndex = -1                             'else, remove any pointer
        cmbComboName.ListIndex = cmbComboName.ListCount - 1                           'else, remove any pointer
    Next I

    txtModelNo.Text = UCase(txtModelNo.Text)
    txtModelNo.SelStart = Len(txtModelNo.Text)
End Function
Public Function LoadCombo(cmbComboName As ComboBox, sTableName As String)
'load all of the pump parameter combo boxes from the tables on the database

    Dim I As Integer
    Dim sItem As String
    Dim iID As Integer
    Dim bUseDropdown As Boolean
    Dim qy As New ADODB.Command
    Dim rs As New ADODB.Recordset

'    rsPumpParameters.CursorLocation = adUseClient
'    If sTableName = "Model" Then
'        rsPumpParameters.Sort = "Model"
'    Else
'        rsPumpParameters.Sort = vbNullString
'    End If
'    rsPumpParameters.Open sTableName, cnPumpData, adOpenStatic, adLockOptimistic, adCmdTableDirect

    qy.ActiveConnection = cnPumpData
    If sTableName = "DischargeDiameter" Or sTableName = "SuctionDiameter" Then
        qy.CommandText = "SELECT * FROM " & sTableName & " ORDER BY Val(Description)"
    Else
        qy.CommandText = "SELECT * FROM " & sTableName & " ORDER BY Description"
    End If
    If sTableName = "SupermarketPumpData" Then
        qy.CommandText = "SELECT ID,Model AS Description FROM " & sTableName
    End If
    rs.CursorLocation = adUseClient
    rs.CursorType = adOpenStatic

    rs.Open qy


    On Error GoTo NoField
    bUseDropdown = True
    'sItem = rsPumpParameters.Fields("UseInDropdown")
'    If bUseDropdown Then
'        rsPumpParameters.Sort = "Description"
'    End If
    rs.MoveFirst                                'goto the top
    For I = 0 To rs.RecordCount - 1             'go through the whole recordset
        sItem = rs.Fields("Description")        'get the description
        iID = rs.Fields(0)                      'get the index number - primary key
        If bUseDropdown Then
'            If rsPumpParameters.Fields("UseInDropdown").value = True Then
                cmbComboName.AddItem sItem, I                                   'add the description to the combo box
'                cmbComboName.AddItem sItem                                   'add the description to the combo box
                cmbComboName.ItemData(cmbComboName.NewIndex) = iID              'add the key number into the item data
'            End If
        End If
        rs.MoveNext                             'get the next record
    Next I
    rs.Close
    cmbComboName.ListIndex = -1
On Error GoTo 0
    Set rs = Nothing
    Set qy = Nothing
    Exit Function

NoField:
    bUseDropdown = False
On Error GoTo 0
    Resume Next
  
End Function
Function SetGraphMax(Plothead) As Integer

    Dim I As Integer
    Dim m As Single

    m = 0
    For I = 0 To UBound(Plothead, 2)
        If Plothead(1, I) > m Then
            m = Plothead(1, I)
        End If
    Next I
    SetGraphMax = 10 * (Int((m / 10) + 0.5) + 1)
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10 * (Int((m / 10) + 0.5) + 1)
    MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0
    
End Function
Public Function CalculateSpeed(CoefSq As Double, CoefLin As Double, CoefConstant As Double, InputHP As Double, SG As Double) As Integer
    Dim I As Integer
    Dim OldResult As Double
    Dim NewResult As Double

    CalculateSpeed = 0

    If SG > 5 Or SG < 0.01 Then
        MsgBox "Bad value for SG...must be between 0.01 and 5.", vbOKOnly, "Bad SG Value"
        Exit Function
    End If

    OldResult = 1000
    NewResult = 0

    I = 1

    Do While Abs(NewResult - OldResult) > 0.1
        ReDim Preserve results(I)
        Select Case I
            Case 1
                results(I - 1).HP = InputHP
            Case 2
                results(I - 1).HP = results(I - 2).HP * SG
            Case Else
                results(I - 1).HP = results(I - 2).HP * (results(I - 2).Speed / results(I - 3).Speed) ^ 3
        End Select
        OldResult = NewResult
        results(I - 1).Speed = CalcPoly(CoefSq, CoefLin, CoefConstant, results(I - 1).HP)
        NewResult = results(I - 1).Speed
        If I > 15 Then
            If I = 0 Or I > 15 Then
                MsgBox "Over 15 calculations and no convergence", vbOKOnly, "Too many iterations"
                Exit Function
            End If
            Exit Function
        End If
        I = I + 1
    Loop
    CalculateSpeed = I - 1
End Function
Public Function CalcPoly(CoefSq As Double, CoefLin As Double, CoefConstant As Double, DataIn As Double) As Double
    CalcPoly = CoefSq * DataIn ^ 2 + CoefLin * DataIn + CoefConstant
End Function

Sub GetBalanceHoleData(SerialNumber As String, TestDate As String)
    If rsBalanceHoles.State = adStateOpen Then
        rsBalanceHoles.Close
    End If
    qyBalanceHoles.CommandText = "SELECT BalanceHoles.*, " & _
                      "IIf([Diameter]=99, 'Slot', [diameter]) as Diameter1, IIf([BoltCircle]=99, 'Unknown', [BoltCircle]) as BoltCircle1 " & _
                      "FROM BalanceHoles " & _
                      "WHERE [SerialNo] = '" & SerialNumber & "' AND [Date] <= #" & TestDate & "# " & _
                      "ORDER BY [Date], Val([BoltCircle]);"

    rsBalanceHoles.Open qyBalanceHoles
    rsBalanceHoles.Filter = ""

    Set dgBalanceHoles.DataSource = rsBalanceHoles

    Dim c As Column
    For Each c In dgBalanceHoles.Columns
        Select Case c.DataField
        Case "BalanceHoleID"
            c.Visible = False
        Case "SerialNo"
            c.Visible = False
        Case "Date"
            c.Visible = True
            c.Alignment = dbgCenter
            c.Width = 2000
        Case "Number"
            c.Visible = True
            c.Alignment = dbgCenter
            c.Width = 700
        Case "Diameter"
            c.Visible = False
        Case "Diameter1"
            c.Caption = "Diameter"
            c.Visible = True
            c.Alignment = dbgCenter
            c.Width = 700
        Case "BoltCircle1"
            c.Caption = "Bolt Circle"
            c.Visible = True
            c.Alignment = dbgCenter
            c.Width = 800
        Case "BoltCircle"
            c.Visible = False
        Case Else ' hide all other columns.
            c.Visible = False
        End Select
    Next c
  
End Sub

Public Sub FixPointsToPlot()
    'count valid data test entry and set points to plot
    If DataGrid2.Row = -1 Then
        Exit Sub
    End If
    Dim PresentGridRow As Integer
    PresentGridRow = DataGrid2.Row
    Dim GridIndex As Integer
    UpDown2.value = 8
    If DataGrid2.Row <> -1 Then
        For GridIndex = 0 To 7
            DataGrid2.Row = GridIndex
            If DataGrid2.Columns("Flow") = 0 And DataGrid2.Columns("TDH") = 0 Then
                txtUpDn2.Text = GridIndex
                If GridIndex = 0 Then
                    UpDown2.value = 8
                Else
                    UpDown2.value = GridIndex
                End If
                Exit Sub
            End If
        Next GridIndex
    End If
    DataGrid2.Row = PresentGridRow
End Sub

Sub SetFrequencyCombo()
    'set default for test spec
    Dim j As Integer
    For j = 0 To cmbFrequency.ListCount - 1
        If cmbFrequency.List(j) = "60 Hz" Then
            cmbFrequency.ListIndex = j
            Exit For
        End If
    Next

End Sub

