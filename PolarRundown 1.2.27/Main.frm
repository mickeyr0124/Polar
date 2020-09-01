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


' <VB WATCH>
Const VBWMODULE = "frmPLCData"
' </VB WATCH>

Private Sub chkBalanceHoles_Click()
           'if the balance holes box is checked, show the datagrid
' <VB WATCH>
1          On Error GoTo vbwErrHandler
2          Const VBWPROCNAME = "frmPLCData.chkBalanceHoles_Click"
3          If vbwProtector.vbwTraceProc Then
4              Dim vbwProtectorParameterString As String
5              If vbwProtector.vbwTraceParameters Then
6                  vbwProtectorParameterString = "()"
7              End If
8              vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
9          End If
' </VB WATCH>
10         If chkBalanceHoles.value = 1 Then
11             dgBalanceHoles.Visible = True
12         Else
13             dgBalanceHoles.Visible = False
14         End If
15         If LenB(frmPLCData.txtSN.Text) = 0 Or LenB(cmbTestDate.Text) = 0 Then
16             dgBalanceHoles.Visible = False
17         End If
' <VB WATCH>
18         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
19         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "chkBalanceHoles_Click"

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

Private Sub chkCircOrifice_Click()
           'if the CircOrifice box is checked, show the size
' <VB WATCH>
20         On Error GoTo vbwErrHandler
21         Const VBWPROCNAME = "frmPLCData.chkCircOrifice_Click"
22         If vbwProtector.vbwTraceProc Then
23             Dim vbwProtectorParameterString As String
24             If vbwProtector.vbwTraceParameters Then
25                 vbwProtectorParameterString = "()"
26             End If
27             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
28         End If
' </VB WATCH>
29         If chkCircOrifice.value = 1 Then
30             lblCircOrifice.Visible = True
31             txtCircOrifice.Visible = True
32         Else
33             lblCircOrifice.Visible = False
34             txtCircOrifice.Visible = False
35         End If
' <VB WATCH>
36         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
37         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "chkCircOrifice_Click"

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

Private Sub chkNPSH_Click()
           'if the NPSH file box is checked, show the file name
' <VB WATCH>
38         On Error GoTo vbwErrHandler
39         Const VBWPROCNAME = "frmPLCData.chkNPSH_Click"
40         If vbwProtector.vbwTraceProc Then
41             Dim vbwProtectorParameterString As String
42             If vbwProtector.vbwTraceParameters Then
43                 vbwProtectorParameterString = "()"
44             End If
45             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
46         End If
' </VB WATCH>
47         If chkNPSH.value = 1 Then
48             txtNPSHFile.Visible = True
49         Else
50             txtNPSHFile.Visible = False
51         End If
' <VB WATCH>
52         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
53         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "chkNPSH_Click"

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

Private Sub chkOrifice_Click()
           'if the orifice box is checked, show the size
' <VB WATCH>
54         On Error GoTo vbwErrHandler
55         Const VBWPROCNAME = "frmPLCData.chkOrifice_Click"
56         If vbwProtector.vbwTraceProc Then
57             Dim vbwProtectorParameterString As String
58             If vbwProtector.vbwTraceParameters Then
59                 vbwProtectorParameterString = "()"
60             End If
61             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
62         End If
' </VB WATCH>
63         If chkOrifice.value = 1 Then
64             lblOrifice.Visible = True
65             txtOrifice.Visible = True
66         Else
67             lblOrifice.Visible = False
68             txtOrifice.Visible = False
69         End If
' <VB WATCH>
70         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
71         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "chkOrifice_Click"

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

Private Sub chkPictures_Click()
           'if the pictures box is checked, show the file name
' <VB WATCH>
72         On Error GoTo vbwErrHandler
73         Const VBWPROCNAME = "frmPLCData.chkPictures_Click"
74         If vbwProtector.vbwTraceProc Then
75             Dim vbwProtectorParameterString As String
76             If vbwProtector.vbwTraceParameters Then
77                 vbwProtectorParameterString = "()"
78             End If
79             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
80         End If
' </VB WATCH>
81         If chkPictures.value = 1 Then
82             txtPicturesFile.Visible = True
83         Else
84             txtPicturesFile.Visible = False
85         End If
' <VB WATCH>
86         If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
87         Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "chkPictures_Click"

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

Private Sub chkTrimmed_Click()
           'if the trimmed box is checked, show the impeller size
' <VB WATCH>
88         On Error GoTo vbwErrHandler
89         Const VBWPROCNAME = "frmPLCData.chkTrimmed_Click"
90         If vbwProtector.vbwTraceProc Then
91             Dim vbwProtectorParameterString As String
92             If vbwProtector.vbwTraceParameters Then
93                 vbwProtectorParameterString = "()"
94             End If
95             vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
96         End If
' </VB WATCH>
97         If chkTrimmed.value = 1 Then
98             lblImpTrim.Visible = True
99             txtImpTrim.Visible = True
100        Else
101            lblImpTrim.Visible = False
102            txtImpTrim.Visible = False
103        End If
' <VB WATCH>
104        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
105        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "chkTrimmed_Click"

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

Private Sub chkVibration_Click()
           'if the vibration box is checked, show the file name
' <VB WATCH>
106        On Error GoTo vbwErrHandler
107        Const VBWPROCNAME = "frmPLCData.chkVibration_Click"
108        If vbwProtector.vbwTraceProc Then
109            Dim vbwProtectorParameterString As String
110            If vbwProtector.vbwTraceParameters Then
111                vbwProtectorParameterString = "()"
112            End If
113            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
114        End If
' </VB WATCH>
115        If chkVibration.value = 1 Then
116            txtVibrationFile.Visible = True
117        Else
118            txtVibrationFile.Visible = False
119        End If
' <VB WATCH>
120        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
121        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "chkVibration_Click"

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



Private Sub cmbFrequency_Click()
' <VB WATCH>
122        On Error GoTo vbwErrHandler
123        Const VBWPROCNAME = "frmPLCData.cmbFrequency_Click"
124        If vbwProtector.vbwTraceProc Then
125            Dim vbwProtectorParameterString As String
126            If vbwProtector.vbwTraceParameters Then
127                vbwProtectorParameterString = "()"
128            End If
129            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
130        End If
' </VB WATCH>
131        If cmbFrequency.Text = "VFD" Then
132            txtVFDFreq.Visible = True
133            lbltab2(86).Visible = True
134        Else
135            txtVFDFreq.Visible = False
136            lbltab2(86).Visible = False
137        End If
' <VB WATCH>
138        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
139        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbFrequency_Click"

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


Private Sub cmbLoopNumber_Click()
' <VB WATCH>
140        On Error GoTo vbwErrHandler
141        Const VBWPROCNAME = "frmPLCData.cmbLoopNumber_Click"
142        If vbwProtector.vbwTraceProc Then
143            Dim vbwProtectorParameterString As String
144            If vbwProtector.vbwTraceParameters Then
145                vbwProtectorParameterString = "()"
146            End If
147            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
148        End If
' </VB WATCH>

149        Dim I As Integer
150        I = cmbLoopNumber.ListIndex

151        Dim qyTransducers As New ADODB.Command
152        Dim rsTransducers As New ADODB.Recordset
153        qyTransducers.ActiveConnection = cnPumpData
154        qyTransducers.CommandText = "SELECT * " & _
                     "From Transducers " & _
                     "Where LoopNumber  = " & I

155        With rsTransducers     'open the recordset for the query
       '        .Index = "FindData"
156            .CursorLocation = adUseClient
157            .CursorType = adOpenStatic
158            .Open qyTransducers
159        End With
160        If rsTransducers.RecordCount = 1 Then
161            Me.cmbFlowMeter.ListIndex = rsTransducers.Fields("FlowMeter")
162            Me.cmbSuctionPressureTransducer.ListIndex = rsTransducers.Fields("SuctionPressure")
163            Me.cmbDischargePressureTransducer.ListIndex = rsTransducers.Fields("DischargePressure")
164            Me.cmbTemperatureTransducer.ListIndex = rsTransducers.Fields("Temperature")
165            Me.cmbPLCNo.ListIndex = rsTransducers.Fields("PLC")
166            Me.cmbAnalyzerNo.ListIndex = rsTransducers.Fields("Analyzer")
167            Me.cmbCirculationFlowMeter.ListIndex = rsTransducers.Fields("CircFlowMeter")
168        End If

       '    If I < 2 Then
       '        Me.cmbPLCNo.ListIndex = 0
       '    Else
       '        Me.cmbPLCNo.ListIndex = 1
       '    End If
' <VB WATCH>
169        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
170        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbLoopNumber_Click"

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
            vbwReportVariable "I", I
            vbwReportVariable "qyTransducers", qyTransducers
            vbwReportVariable "rsTransducers", rsTransducers
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub GetSuperMarketPump(SuperMarketPartNum As String, JobNumber As String)
' <VB WATCH>
171        On Error GoTo vbwErrHandler
172        Const VBWPROCNAME = "frmPLCData.GetSuperMarketPump"
173        If vbwProtector.vbwTraceProc Then
174            Dim vbwProtectorParameterString As String
175            If vbwProtector.vbwTraceParameters Then
176                vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("SuperMarketPartNum", SuperMarketPartNum) & ", "
177                vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("JobNumber", JobNumber) & ") "
178            End If
179            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
180        End If
' </VB WATCH>

           'get the data from the SupermarketPumpData table
181        qySupermarketModel.ActiveConnection = cnPumpData
182        qySupermarketModel.CommandText = "SELECT * " & _
                     "From SupermarketPumpData " & _
                     "Where Model  = '" & SuperMarketPartNum & "'"

                     'cmbSupermarketModel.ItemData(cmbSupermarketModel.ListIndex)"

183        If rsSupermarketModel.State = adStateOpen Then
184            rsSupermarketModel.Close
185        End If

186        With rsSupermarketModel     'open the recordset for the query
       '        .Index = "FindData"
187            .CursorLocation = adUseClient
188            .CursorType = adOpenStatic
189            .Open qySupermarketModel
190        End With
191        If rsSupermarketModel.RecordCount = 1 Then
192            txtSalesOrderNumber.Text = rsSupermarketModel.Fields("SalesOrder")
193            txtLineNumber.Text = rsSupermarketModel.Fields("LineNumber")
194            txtShpNo.Text = rsSupermarketModel.Fields("ShipTo")
195            txtBilNo.Text = rsSupermarketModel.Fields("BillTo")
196            txtDesignFlow.Text = rsSupermarketModel.Fields("DesignFlow")
197            txtDesignTDH.Text = rsSupermarketModel.Fields("DesignTDH")
198            txtNoPhases.Text = rsSupermarketModel.Fields("Phases")
199            txtNPSHr.Text = rsSupermarketModel.Fields("NPSHr")
200            txtRatedInputPower.Text = rsSupermarketModel.Fields("RatedInputPower")
201            txtAmps.Text = rsSupermarketModel.Fields("RatedCurrent")
202            txtThermalClass.Text = rsSupermarketModel.Fields("ThermalClass")
203            txtSpGr.Text = rsSupermarketModel.Fields("SG")
204            txtViscosity.Text = rsSupermarketModel.Fields("Viscosity")
205            txtTEMCViscosity.Text = Format((Val(rsSupermarketModel.Fields("Viscosity")) / Val(txtSpGr.Text)), "000.00")
206            txtExpClass.Text = rsSupermarketModel.Fields("EXPClass")
207            txtLiquid.Text = rsSupermarketModel.Fields("Liquid")
208            txtLiquidTemperature.Text = rsSupermarketModel.Fields("LiquidTemp")
209            txtJobNum.Text = JobNumber
210            txtImpellerDia.Text = rsSupermarketModel.Fields("ImpellerDiameter")
211            txtModelNo.Text = rsSupermarketModel.Fields("Model")
212            txtRVSPartNo.Text = rsSupermarketModel.Fields("RVSPartNo")
213            cmdSelectSupermarket.Caption = "Save Data"
214            If UCase(rsSupermarketModel.Fields("Feathered")) = "FEATHERED" Then
215                Me.chkSuperMarketFeathered.value = Checked
216            End If
217        End If
218        grpSupermarket.Visible = False

' <VB WATCH>
219        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
220        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetSuperMarketPump"

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
            vbwReportVariable "SuperMarketPartNum", SuperMarketPartNum
            vbwReportVariable "JobNumber", JobNumber
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmbPLCNo_Click()
           'cmbplc text either contains 8 or 9
' <VB WATCH>
221        On Error GoTo vbwErrHandler
222        Const VBWPROCNAME = "frmPLCData.cmbPLCNo_Click"
223        If vbwProtector.vbwTraceProc Then
224            Dim vbwProtectorParameterString As String
225            If vbwProtector.vbwTraceParameters Then
226                vbwProtectorParameterString = "()"
227            End If
228            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
229        End If
' </VB WATCH>

230        Dim I As Integer
231        Dim PLCNo As Integer
232        Dim MagtrolNo As String

233        PLCNo = 0
234        If InStr(cmbPLCNo.Text, "8") > 0 Then
235            PLCNo = 8
236            MagtrolNo = "GPIB6"
237        End If
238        If InStr(cmbPLCNo.Text, "9") > 0 Then
239            PLCNo = 9
240            MagtrolNo = "GPIB5"
241        End If

242        For I = 0 To cmbPLCLoop.ListCount - 1                     'go through the combobox entries
243            If InStr(cmbPLCLoop.List(I), PLCNo) > 0 Then   'see when we find the desired index number
244                cmbPLCLoop.ListIndex = I                                              'if we do, set the combo box
245                Exit For                                            'and we're done
246            End If
               'cmbPLCLoop.ListIndex = -1                             'else, remove any pointer
247        Next I

248        For I = 0 To cmbMagtrol.ListCount - 1
249            If InStr(cmbMagtrol.List(I), MagtrolNo) > 0 Then   'see when we find the desired index number
250                cmbMagtrol.ListIndex = I                                              'if we do, set the combo box
251                Exit For                                            'and we're done
252            End If
253        Next I
' <VB WATCH>
254        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
255        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbPLCNo_Click"

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
            vbwReportVariable "I", I
            vbwReportVariable "PLCNo", PLCNo
            vbwReportVariable "MagtrolNo", MagtrolNo
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmbVoltage_click()
' <VB WATCH>
256        On Error GoTo vbwErrHandler
257        Const VBWPROCNAME = "frmPLCData.cmbVoltage_click"
258        If vbwProtector.vbwTraceProc Then
259            Dim vbwProtectorParameterString As String
260            If vbwProtector.vbwTraceParameters Then
261                vbwProtectorParameterString = "()"
262            End If
263            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
264        End If
' </VB WATCH>
265        If Me.cmbVoltage.ListIndex = 0 Then
266            Me.cmbFrequency.ListIndex = 2
267        End If
' <VB WATCH>
268        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
269        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbVoltage_click"

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
Private Sub cmbMagtrol_Click()
' <VB WATCH>
270        On Error GoTo vbwErrHandler
271        Const VBWPROCNAME = "frmPLCData.cmbMagtrol_Click"
272        If vbwProtector.vbwTraceProc Then
273            Dim vbwProtectorParameterString As String
274            If vbwProtector.vbwTraceParameters Then
275                vbwProtectorParameterString = "()"
276            End If
277            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
278        End If
' </VB WATCH>
279        Dim I As Integer
280        Dim sSendStr As String
281        Dim sGPIBName As String
282        Dim MagtrolName As String

283        I = cmbMagtrol.ItemData(cmbMagtrol.ListIndex)
284        sGPIBName = "GPIB" & I
285        MagtrolName = cmbMagtrol.List(cmbMagtrol.ListIndex)

286        If I = 99 Then      'manual entry
287            boMagtrolOperating = False
288            EnableMagtrolFields
' <VB WATCH>
289        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
290            Exit Sub
291        Else
292            boMagtrolOperating = True
293        End If

294        SetupMagtrols MagtrolName, I

' <VB WATCH>
295        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
296        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbMagtrol_Click"

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
            vbwReportVariable "I", I
            vbwReportVariable "sSendStr", sSendStr
            vbwReportVariable "sGPIBName", sGPIBName
            vbwReportVariable "MagtrolName", MagtrolName
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub


Private Sub cmbPLCLoop_Click()
           'Change the PLC that we're looking at
' <VB WATCH>
297        On Error GoTo vbwErrHandler
298        Const VBWPROCNAME = "frmPLCData.cmbPLCLoop_Click"
299        If vbwProtector.vbwTraceProc Then
300            Dim vbwProtectorParameterString As String
301            If vbwProtector.vbwTraceParameters Then
302                vbwProtectorParameterString = "()"
303            End If
304            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
305        End If
' </VB WATCH>

306        Dim RetVal As String

           'manual data entry selection
307        If cmbPLCLoop.ListIndex = cmbPLCLoop.ListCount - 1 Then 'no plc
308            boPLCOperating = False
309            EnablePLCFields
310            If DeviceOpen = True Then
311                RetVal = DisconnectPLC()
312            End If
' <VB WATCH>
313        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
314            Exit Sub
315        End If

316        If DeviceOpen = True Then
317            RetVal = DisconnectPLC()
318        End If

319        RetVal = ConnectToPLC(cmbPLCLoop.ItemData(cmbPLCLoop.ListIndex))
320        If RetVal <> 0 Then
321            MsgBox ("Can't connect to PLC - " & Description(cmbPLCLoop.ListIndex))
322            boPLCOperating = False
323            EnablePLCFields
324        Else
325            boPLCOperating = True
326            tDevice = cmbPLCLoop.ItemData(cmbPLCLoop.ListIndex)
327            DisablePLCFields
328        End If
' <VB WATCH>
329        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
330        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbPLCLoop_Click"

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
            vbwReportVariable "RetVal", RetVal
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmbTestDate_Click()
           'select a test date to show
' <VB WATCH>
331        On Error GoTo vbwErrHandler
332        Const VBWPROCNAME = "frmPLCData.cmbTestDate_Click"
333        If vbwProtector.vbwTraceProc Then
334            Dim vbwProtectorParameterString As String
335            If vbwProtector.vbwTraceParameters Then
336                vbwProtectorParameterString = "()"
337            End If
338            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
339        End If
' </VB WATCH>

340        Dim sName As String
341        Dim sParam As String
342        Dim I As Integer
343        Dim j As Integer
344        Dim k As Integer
345        Dim bSk As Boolean
346        Dim sBC As Single
347        Dim NOK() As Long

348        cmdModifyBalanceHoleData.Visible = False


349        If Not boFoundTestSetup Then    'if we don't have any TestSetup data written
350            boFoundTestData = False
' <VB WATCH>
351        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
352            Exit Sub
353        End If


           'select the testsetup data for the serial number
354        qyTestSetup.ActiveConnection = cnPumpData
355        qyTestSetup.CommandText = "SELECT * " & _
                         "From TempTestSetupData " & _
                         "Where (((TempTestSetupData.SerialNumber) = '" & txtSN.Text & "') AND TempTestSetupData.Date = #" & cmbTestDate.List(cmbTestDate.ListIndex) & "#) " & _
                         "ORDER BY TempTestSetupData.Date;"

356        If rsTestSetup.State = adStateOpen Then
357            rsTestSetup.Close
358        End If

359        With rsTestSetup     'open the recordset for the query
       '        .Index = "FindData"
360            .CursorLocation = adUseClient
361            .CursorType = adOpenStatic
362            .Open qyTestSetup
363        End With

           'move to the selected date
364        If Not rsTestSetup.BOF Then
365            rsTestSetup.MoveFirst
366        End If
       '
           'show the correct combo box entries for this record
           'SetComboTestSetup cmbOrificeNumber, "OrificeNumber", "OrificeNumber", rsTestSetup
367        SetComboTestSetup cmbTestSpec, "TestSpec", "TestSpecification", rsTestSetup
368        SetComboTestSetup cmbLoopNumber, "LoopNumber", "LoopNumber", rsTestSetup
369        SetComboTestSetup cmbSuctDia, "SuctDiam", "SuctionDiameter", rsTestSetup
370        SetComboTestSetup cmbDischDia, "DischDiam", "DischargeDiameter", rsTestSetup
371        SetComboTestSetup cmbTachID, "TachID", "TachID", rsTestSetup
372        SetComboTestSetup cmbAnalyzerNo, "AnalyzerNo", "AnalyzerNo", rsTestSetup
373        SetComboTestSetup cmbVoltage, "Voltage", "Voltage", rsTestSetup
374        SetComboTestSetup cmbFrequency, "Frequency", "Frequency", rsTestSetup
375        SetComboTestSetup cmbMounting, "Mounting", "Mounting", rsTestSetup
376        SetComboTestSetup cmbPLCNo, "PLCNo", "PLCNo", rsTestSetup
377        SetComboTestSetup cmbFlowMeter, "FlowMeterID", "PumpFlowMeter", rsTestSetup
378        SetComboTestSetup cmbSuctionPressureTransducer, "SuctionID", "SuctionPressureTransducer", rsTestSetup
379        SetComboTestSetup cmbDischargePressureTransducer, "DischID", "DischargePressureTransducer", rsTestSetup
380        SetComboTestSetup cmbTemperatureTransducer, "TemperatureID", "TemperatureTransducer", rsTestSetup
381        SetComboTestSetup cmbCirculationFlowMeter, "MagFlowID", "CirculationFlowMeter", rsTestSetup

382        sName = "HDCor"
383        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
384            sParam = rsTestSetup.Fields(sName)
385        Else
386            sParam = vbNullString
387        End If
388        txtHDCor.Text = sParam

389        sName = "KWMult"
390        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
391            sParam = rsTestSetup.Fields(sName)
392        Else
393            sParam = vbNullString
394        End If
395        txtKWMult.Text = sParam

396        sName = "Who"
397        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
398            sParam = rsTestSetup.Fields(sName)
399        Else
400            sParam = vbNullString
401        End If
402        txtWho.Text = sParam

403        sName = "RMA"
404        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
405            sParam = rsTestSetup.Fields(sName)
406        Else
407            sParam = vbNullString
408        End If
409        txtRMA.Text = sParam

410        sName = "Remarks"
411        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
412            sParam = rsTestSetup.Fields(sName)
413        Else
414            sParam = vbNullString
415        End If
416        txtTestSetupRemarks.Text = sParam

417        sName = "VFDFrequency"
418        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
419            sParam = rsTestSetup.Fields(sName)
420        Else
421            sParam = vbNullString
422        End If
423        txtVFDFreq.Text = sParam

424        sName = "SuctionGageHeight"
425        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
426            sParam = rsTestSetup.Fields(sName)
427        Else
428            sParam = 0
429        End If
430        txtSuctHeight.Text = sParam

431        sName = "DischargeGageHeight"
432        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
433            sParam = rsTestSetup.Fields(sName)
434        Else
435            sParam = 0
436        End If
437        txtDischHeight.Text = sParam

438        sName = "EndPlay"
439        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
440            sParam = rsTestSetup.Fields(sName)
441        Else
442            sParam = vbNullString
443        End If
444        txtEndPlay.Text = sParam

445        sName = "GGAP"
446        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
447            sParam = rsTestSetup.Fields(sName)
448        Else
449            sParam = vbNullString
450        End If
451        txtGGap.Text = sParam

452        sName = "OtherMods"
453        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
454            sParam = rsTestSetup.Fields(sName)
455        Else
456            sParam = vbNullString
457        End If
458        txtOtherMods.Text = sParam

459        If rsTestSetup.Fields("ImpFeathered") Then
460            chkFeathered.value = 1
461        Else
462            chkFeathered.value = 0
463        End If

464        If Val(rsTestSetup.Fields("ImpTrimmed")) = 0 Then
465            chkTrimmed.value = 0
466            txtImpTrim.Visible = False
467            txtImpTrim.Text = rsTestSetup.Fields("Imptrimmed")
468        Else
469            chkTrimmed.value = 1
470            txtImpTrim.Visible = True
471            txtImpTrim.Text = rsTestSetup.Fields("Imptrimmed")
472        End If

473        If Val(rsTestSetup.Fields("PumpDischOrifice")) = 0 Then
474            chkOrifice.value = 0
475            txtOrifice.Visible = False
476        Else
477            chkOrifice.value = 1
478            txtOrifice.Visible = True
479            txtOrifice.Text = rsTestSetup.Fields("PumpDischOrifice")
480        End If

481        If Val(rsTestSetup.Fields("CircFlowOrifice")) = 0 Then
482            chkCircOrifice.value = 0
483            txtCircOrifice.Visible = False
484        Else
485            chkCircOrifice.value = 1
486            txtCircOrifice.Visible = True
487            txtCircOrifice.Text = rsTestSetup.Fields("CircFlowOrifice")
488        End If

489        If (IsNull(rsTestSetup.Fields("NPSHFile"))) Or (LenB(rsTestSetup.Fields("NPSHFile")) = 0) Then
490            chkNPSH.value = 0
491            txtNPSHFile.Visible = False
492        Else
493            chkNPSH.value = 1
494            txtNPSHFile.Visible = True
495            txtNPSHFile.Text = rsTestSetup.Fields("NPSHFile")
496        End If

497        If (IsNull(rsTestSetup.Fields("PictureFile"))) Or (LenB(rsTestSetup.Fields("PictureFile")) = 0) Then
498            chkPictures.value = 0
499            txtPicturesFile.Visible = False
500        Else
501            chkPictures.value = 1
502            txtPicturesFile.Visible = True
503            txtPicturesFile.Text = rsTestSetup.Fields("PictureFile")
504        End If

505        If (IsNull(rsTestSetup.Fields("VibrationFile"))) Or (LenB(rsTestSetup.Fields("VibrationFile")) = 0) Then
506            chkVibration.value = 0
507            txtVibrationFile.Visible = False
508        Else
509            chkVibration.value = 1
510            txtVibrationFile.Visible = True
511            txtVibrationFile.Text = rsTestSetup.Fields("VibrationFile")
512        End If


           'for TEMC Inspection Report
513        sName = "InsulationMeggerVolts"
514        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
515            sParam = rsTestSetup.Fields(sName)
516        Else
517            sParam = 0
518        End If
519        txtTestAndInspection(0).Text = sParam

520        sName = "InsulationMegOhms"
521        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
522            sParam = rsTestSetup.Fields(sName)
523        Else
524            sParam = 0
525        End If
526        txtTestAndInspection(1).Text = sParam

527        sName = "DielectricVolts"
528        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
529            sParam = rsTestSetup.Fields(sName)
530        Else
531            sParam = 0
532        End If
533        txtTestAndInspection(2).Text = sParam

534        sName = "DielectricTime"
535        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
536            sParam = rsTestSetup.Fields(sName)
537        Else
538            sParam = 0
539        End If
540        txtTestAndInspection(3).Text = sParam

541        sName = "HydrostaticValue"
542        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
543            sParam = rsTestSetup.Fields(sName)
544        Else
545            sParam = 0
546        End If
547        txtTestAndInspection(4).Text = sParam

548        sName = "HydrostaticTime"
549        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
550            sParam = rsTestSetup.Fields(sName)
551        Else
552            sParam = 0
553        End If
554        txtTestAndInspection(5).Text = sParam

555        sName = "PneumaticValue"
556        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
557            sParam = rsTestSetup.Fields(sName)
558        Else
559            sParam = 0
560        End If
561        txtTestAndInspection(6).Text = sParam

562        sName = "PneumaticTime"
563        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
564            sParam = rsTestSetup.Fields(sName)
565        Else
566            sParam = 0
567        End If
568        txtTestAndInspection(7).Text = sParam

569        For I = 0 To cmbTestAndInspection(0).ListCount - 1
570            If cmbTestAndInspection(0).Text = rsTestSetup.Fields("HydrostaticUnits") Then
571                    cmbTestAndInspection(0).ListIndex = I
572                    Exit For
573            End If
574            cmbTestAndInspection(0).ListIndex = -1
575        Next I


576        For I = 0 To cmbTestAndInspection(1).ListCount - 1
577            If cmbTestAndInspection(1).Text = rsTestSetup.Fields("PneumaticUnits") Then
578                    cmbTestAndInspection(1).ListIndex = I
579                    Exit For
580            End If
581            cmbTestAndInspection(1).ListIndex = -1
582        Next I

583        TestAndInspectionGood(0).value = Abs(rsTestSetup!insulationgood)
584        TestAndInspectionGood(1).value = Abs(rsTestSetup!DielectricGood)
585        TestAndInspectionGood(2).value = Abs(rsTestSetup!HydrostaticGood)
586        TestAndInspectionGood(3).value = Abs(rsTestSetup!PneumaticGood)
587        TestAndInspectionGood(4).value = Abs(rsTestSetup!GeneralAppearanceGood)
588        TestAndInspectionGood(5).value = Abs(rsTestSetup!OutlineDimensionsGood)
589        TestAndInspectionGood(6).value = Abs(rsTestSetup!MotorNoLoadTestGood)
590        TestAndInspectionGood(7).value = Abs(rsTestSetup!MotorLockedRotorTestGood)
591        TestAndInspectionGood(8).value = Abs(rsTestSetup!HydrostaticTestGood)
592        TestAndInspectionGood(9).value = Abs(rsTestSetup!HydraulicTestGood)
593        TestAndInspectionGood(10).value = Abs(rsTestSetup!NPSHTestGood)
594        TestAndInspectionGood(11).value = Abs(rsTestSetup!CleanPurgeSealGood)
595        TestAndInspectionGood(12).value = Abs(rsTestSetup!PaintCheckGood)
596        TestAndInspectionGood(13).value = Abs(rsTestSetup!NameplateGood)
597        TestAndInspectionGood(14).value = Abs(rsTestSetup!SupervisorApproval)

598        GetBalanceHoleData frmPLCData.txtSN.Text, cmbTestDate.Text

599         If rsBalanceHoles.RecordCount = 0 Then
600            chkBalanceHoles.value = 0
601            dgBalanceHoles.Visible = False
602            boGotBalanceHoles = False
603        Else
604            boGotBalanceHoles = True
605            ReDim NOK(rsBalanceHoles.RecordCount)
606            rsBalanceHoles.MoveLast
607            For I = 1 To rsBalanceHoles.RecordCount
608                NOK(I) = 0
609            Next I

610            For j = 1 To rsBalanceHoles.RecordCount - 1
611                rsBalanceHoles.MoveFirst
612                rsBalanceHoles.Move rsBalanceHoles.RecordCount - j
613                sBC = rsBalanceHoles.Fields("BoltCircle")
614                bSk = False
615                For k = 1 To rsBalanceHoles.RecordCount
616                    If NOK(k) = rsBalanceHoles.Fields(0) Then
617                        bSk = True
618                    End If
619                Next k
620                If Not bSk Then
621                    For I = rsBalanceHoles.RecordCount - j To 1 Step -1
622                        rsBalanceHoles.MovePrevious
623                        If rsBalanceHoles.Fields("BoltCircle") = sBC Then
624                            NOK(I) = rsBalanceHoles.Fields(0)
625                        End If
626                    Next I
627                End If
628            Next j

629            Dim sFilt As String
630            sFilt = ""
631            For I = 1 To rsBalanceHoles.RecordCount
632                If NOK(I) <> 0 Then
633                    sFilt = sFilt & "(BalanceHoleID <> " & NOK(I) & ") AND "
       '                sFilt = sFilt & "(" & rsBalanceHoles.Filter & " AND BalanceHoleID <> " & NOK(I) & ") AND "
634                End If
635            Next I

636            If Len(sFilt) > 4 Then
637                sFilt = Left(sFilt, Len(sFilt) - 4)
638                rsBalanceHoles.Filter = sFilt
639            End If

640            chkBalanceHoles.value = 1
641            dgBalanceHoles.Visible = True
642        End If
       '
           'set the test date filter for the test data
643        rsTestData.Filter = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"

644        If rsTestData.RecordCount = 0 Then
645            boFoundTestData = False
646            AddTestData
647            EnableTestDataControls
648            MsgBox "No Test Data Exists for this Serial Number"
649        Else
650            boFoundTestData = True
651            DisableTestDataControls                         'if it's in the real database, don't allow changes here
652        End If

653        If Not boTestDateIsApproved Then    'data approved?
654            EnableTestDataControls
655        End If

656        If rsTestSetup.Fields("Approved") = True Then
657            DisableTestDataControls                         'if it's in the real database, don't allow changes here
658            lblTestDateApproved.Visible = True
659            MsgBox ("Found pump.  Data cannot be modified.")
660            If boCanApprove Then
661                cmdApproveTestDate.Caption = "Unapprove this Test Date"
662            End If
663        Else
664            EnableTestDataControls                          'it's in the temp database, allow changes
665            lblTestDateApproved.Visible = False
666            If boPumpIsApproved = True Then
667                MsgBox ("Found pump.  Pump data cannot be modified, but test setup data and test data can be modified.")
668            Else
669                MsgBox ("Found pump.  Pump data, test setup data, and test data can be modified.")
670            End If
671            If boCanApprove Then
672                If rsPumpData.Fields("Approved") = True Then
673                    cmdApproveTestDate.Enabled = True
674                    cmdApproveTestDate.Caption = "Approve this Test Date"
675                Else
676                    cmdApproveTestDate.Caption = "You Must Approve Pump First"
677                    cmdApproveTestDate.Enabled = False
678                End If
679            End If
680        End If

681        rsEff.MoveFirst
682        rsTestData.MoveFirst

683        For I = 1 To rsTestData.RecordCount
684            DoEfficiencyCalcs
685            rsEff.MoveNext
686            rsTestData.MoveNext
687        Next I

          ' fix the datagrid
688       Set DataGrid1.DataSource = rsTestData
689       Set DataGrid2.DataSource = rsEff

690       Dim c As Column
691       For Each c In DataGrid1.Columns
692          Select Case c.DataField
             Case "TestDataID"     'Hide some columns
693             c.Visible = False
694          Case "SerialNumber"
695             c.Visible = False
696          Case "Date"
697             c.Visible = False
698          Case Else             ' Show all other columns.
699             c.Visible = True
700             c.Alignment = dbgRight
701          End Select
702        Next c

703        For Each c In DataGrid2.Columns
704            c.Alignment = dbgCenter
705            c.Width = 750
706            Select Case c.ColIndex
                   Case 1
707                    c.Caption = "Flow"
708                    c.NumberFormat = "###0.00"
709                Case 2
710                    c.Caption = "TDH"
711                    c.NumberFormat = "##0.00"
712                Case 3
713                    c.Caption = "Input Pwr"
714                    c.NumberFormat = "##0.00"
715                    c.Width = 850
716                Case 4
717                    c.Caption = "Voltage"
718                    c.NumberFormat = "##0.00"
719                Case 5
720                    c.Caption = "Current"
721                    c.NumberFormat = "##0.00"
722                Case 6
723                    c.Caption = "Overall Eff"
724                    c.NumberFormat = "##0.00"
725                    c.Width = 850
726                Case 7
727                    c.Caption = "NPSHr"
728                    c.NumberFormat = "#0.00"
729                Case Else
730                    c.Visible = False
731            End Select
732        Next c
733            FixPointsToPlot

734        txtUpDn1.Text = 1

       'unlock the text boxes
735        For I = 0 To 7
736            txtTitle(I).Locked = False
737        Next I

738        For I = 20 To 27
739            txtTitle(I).Locked = False
740        Next I

       'look for titles for TCs and AIs
741        Dim qy As New ADODB.Command
742        Dim rs As New ADODB.Recordset

743        qy.ActiveConnection = cnPumpData

           'see if we have an entry in the table
744        qy.CommandText = "SELECT * FROM AITitles " & _
                             "WHERE (((AITitles.SerialNo)= '" & txtSN.Text & "') " & _
                             "AND ((AITitles.Date)= #" & cmbTestDate.Text & "#)); "

745        With rs     'open the recordset for the query
746            .CursorLocation = adUseClient
747            .CursorType = adOpenStatic
748            .LockType = adLockOptimistic
749            .Open qy
750        End With

751        If Not (rs.BOF = True And rs.EOF = True) Then   'update titles
752            rs.MoveFirst
753            Do While Not rs.EOF
754                txtTitle(rs.Fields("Channel")).Text = rs.Fields("Title")
755                rs.MoveNext
756            Loop
757        End If

758        rs.Close
759        Set rs = Nothing
760        Set qy = Nothing
' <VB WATCH>
761        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
762        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmbTestDate_Click"

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
            vbwReportVariable "sName", sName
            vbwReportVariable "sParam", sParam
            vbwReportVariable "I", I
            vbwReportVariable "j", j
            vbwReportVariable "k", k
            vbwReportVariable "bSk", bSk
            vbwReportVariable "sBC", sBC
            vbwReportVariable "NOK", NOK
            vbwReportVariable "sFilt", sFilt
            vbwReportVariable "c", c
            vbwReportVariable "qy", qy
            vbwReportVariable "rs", rs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmdAddNewBalanceHoles_Click()
' <VB WATCH>
763        On Error GoTo vbwErrHandler
764        Const VBWPROCNAME = "frmPLCData.cmdAddNewBalanceHoles_Click"
765        If vbwProtector.vbwTraceProc Then
766            Dim vbwProtectorParameterString As String
767            If vbwProtector.vbwTraceParameters Then
768                vbwProtectorParameterString = "()"
769            End If
770            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
771        End If
' </VB WATCH>
772        Dim strInput As String
773        Dim I As Integer
774        Dim sNumber As Integer
775        Dim sDia As Single
776        Dim sBC As Single

           'get the data for the balance holes
777        strInput = InputBox("Enter Number of Holes")
778        If strInput <> "" Then
779            sNumber = CInt(strInput)
780        Else
781            GoTo CancelPressed
782        End If

783        strInput = InputBox("Enter Decimal Value of Hole Diameter or Slot (For Example, 0.675) ")
784        If strInput <> "" Then
785            If UCase(strInput) = "SLOT" Then
786                strInput = 99
787            End If
788            sDia = CSng(strInput)
789        Else
790            GoTo CancelPressed
791        End If

792        strInput = InputBox("Enter Decimal Value of Bolt Circle or Unknown (For Example, 4.525)")
793        If strInput <> "" Then
794            If UCase(strInput) = "UNKNOWN" Then
795                strInput = 99
796            End If
797            sBC = CSng(strInput)
798        Else
799            GoTo CancelPressed
800        End If

801        GetBalanceHoleData frmPLCData.txtSN.Text, cmbTestDate.Text

802        rsBalanceHoles.AddNew
803        rsBalanceHoles!SerialNo = txtSN.Text
804        rsBalanceHoles!Date = cmbTestDate.Text
805        rsBalanceHoles!Number = sNumber
806        rsBalanceHoles!diameter = sDia
807        rsBalanceHoles!boltcircle = sBC

808        rsBalanceHoles.Update

809        GetBalanceHoleData txtSN.Text, cmbTestDate.Text
810        rsBalanceHoles.MoveLast
811        dgBalanceHoles.Refresh
812        chkBalanceHoles.value = 1

' <VB WATCH>
813        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
814        Exit Sub

815    CancelPressed:
816        MsgBox "No New Balance Hole Data Entered", vbOKOnly
' <VB WATCH>
817        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
818        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdAddNewBalanceHoles_Click"

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
            vbwReportVariable "strInput", strInput
            vbwReportVariable "I", I
            vbwReportVariable "sNumber", sNumber
            vbwReportVariable "sDia", sDia
            vbwReportVariable "sBC", sBC
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmdAddNewTestDate_Click()
           'add a new test date/time
' <VB WATCH>
819        On Error GoTo vbwErrHandler
820        Const VBWPROCNAME = "frmPLCData.cmdAddNewTestDate_Click"
821        If vbwProtector.vbwTraceProc Then
822            Dim vbwProtectorParameterString As String
823            If vbwProtector.vbwTraceParameters Then
824                vbwProtectorParameterString = "()"
825            End If
826            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
827        End If
' </VB WATCH>
828        Dim I As Integer

829        chkFeathered.value = chkSuperMarketFeathered.value

830        For I = 1 To cmbTestDate.ListCount      'see if we already have today's date entered
831            If cmbTestDate.List(I) = Date Then
832                MsgBox "There is already an entry for today.  You can only have one entry for each Serial Number and a given date.  You may want to modify the Serial Number.", vbOKOnly
' <VB WATCH>
833        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
834                Exit Sub
835            End If
836        Next I

           'we didn't find today's date entered, allow data entry
837        boFoundTestSetup = False

838        SetFrequencyCombo

839        EnableTestSetupDataControls
840        Pressed = False
841        cmdEnterTestSetupData_Click
842        cmdAddNewBalanceHoles.Visible = True
843        txtWho.Text = LogInInitials
844        MsgBox "New Test Date Added - " & cmbTestDate.List(cmbTestDate.ListCount - 1), vbOKOnly, "Added New Test Date"
' <VB WATCH>
845        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
846        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdAddNewTestDate_Click"

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
            vbwReportVariable "I", I
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmdApprovePump_Click()
           'allow the pump data to be approved
' <VB WATCH>
847        On Error GoTo vbwErrHandler
848        Const VBWPROCNAME = "frmPLCData.cmdApprovePump_Click"
849        If vbwProtector.vbwTraceProc Then
850            Dim vbwProtectorParameterString As String
851            If vbwProtector.vbwTraceParameters Then
852                vbwProtectorParameterString = "()"
853            End If
854            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
855        End If
' </VB WATCH>
856        rsPumpData.Fields("Approved") = Not rsPumpData.Fields("Approved")
857        rsPumpData.Update
858        rsPumpData.Requery
859        lblPumpApproved.Visible = rsPumpData.Fields("Approved")
860        If rsPumpData.Fields("Approved") = True Then
861            cmdApprovePump.Caption = "Unapprove This Pump"
862            cmdApproveTestDate.Enabled = True
863            If rsTestSetup.Fields("Approved") = True Then
864                cmdApproveTestDate.Caption = "Unapprove This Test Date"
865            Else
866                cmdApproveTestDate.Caption = "Approve This Test Date"
867            End If
868        Else
869            cmdApprovePump.Caption = "Approve This Pump"
870            cmdApproveTestDate.Caption = "You Must Approve Pump First"
871            cmdApproveTestDate.Enabled = False
872        End If
' <VB WATCH>
873        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
874        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdApprovePump_Click"

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

Private Sub cmdApproveTestDate_Click()
           'allow the test setup data to be approved
' <VB WATCH>
875        On Error GoTo vbwErrHandler
876        Const VBWPROCNAME = "frmPLCData.cmdApproveTestDate_Click"
877        If vbwProtector.vbwTraceProc Then
878            Dim vbwProtectorParameterString As String
879            If vbwProtector.vbwTraceParameters Then
880                vbwProtectorParameterString = "()"
881            End If
882            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
883        End If
' </VB WATCH>
884        rsTestSetup.Fields("Approved") = Not rsTestSetup.Fields("Approved")
885        rsTestSetup.Update
886        rsTestSetup.Requery
887        lblTestDateApproved.Visible = rsTestSetup.Fields("Approved")
888        If rsTestSetup.Fields("Approved") = True Then
889            cmdApproveTestDate.Caption = "Unapprove This Test Date"
890        Else
891            cmdApproveTestDate.Caption = "Approve This Test Date"
892        End If
' <VB WATCH>
893        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
894        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdApproveTestDate_Click"

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

Private Sub cmdCalibrate_Click()
' <VB WATCH>
895        On Error GoTo vbwErrHandler
896        Const VBWPROCNAME = "frmPLCData.cmdCalibrate_Click"
897        If vbwProtector.vbwTraceProc Then
898            Dim vbwProtectorParameterString As String
899            If vbwProtector.vbwTraceParameters Then
900                vbwProtectorParameterString = "()"
901            End If
902            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
903        End If
' </VB WATCH>
904        Dim ans As Integer
905        Dim I As Integer

906        ans = MsgBox("You have selected to calibrate the software.  Do you want to continue?", vbYesNo, "Calibrate Software")
907        If ans = vbNo Then
908            Calibrating = False
' <VB WATCH>
909        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
910            Exit Sub
911        Else
912            CalibrateSoftware
913        End If
' <VB WATCH>
914        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
915        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdCalibrate_Click"

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
            vbwReportVariable "ans", ans
            vbwReportVariable "I", I
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmdClearPumpData_Click()
' <VB WATCH>
916        On Error GoTo vbwErrHandler
917        Const VBWPROCNAME = "frmPLCData.cmdClearPumpData_Click"
918        If vbwProtector.vbwTraceProc Then
919            Dim vbwProtectorParameterString As String
920            If vbwProtector.vbwTraceParameters Then
921                vbwProtectorParameterString = "()"
922            End If
923            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
924        End If
' </VB WATCH>
925        BlankData
' <VB WATCH>
926        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
927        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdClearPumpData_Click"

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

Private Sub cmdDeletePump_Click()
           'delete this pump
' <VB WATCH>
928        On Error GoTo vbwErrHandler
929        Const VBWPROCNAME = "frmPLCData.cmdDeletePump_Click"
930        If vbwProtector.vbwTraceProc Then
931            Dim vbwProtectorParameterString As String
932            If vbwProtector.vbwTraceParameters Then
933                vbwProtectorParameterString = "()"
934            End If
935            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
936        End If
' </VB WATCH>
937        Dim Answer As Integer
938        Answer = MsgBox("You are about to delete the following record: S/N = " & rsPumpData.Fields("SerialNumber") & "!  Do you want to continue?", vbCritical Or vbYesNo, "Ready to Delete")
939        If Answer = vbYes Then
940            rsPumpData.Delete
941            rsPumpData.Update
942            cmdFindPump_Click
943        End If
' <VB WATCH>
944        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
945        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdDeletePump_Click"

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
            vbwReportVariable "Answer", Answer
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmdDeleteTestDate_Click()
           'delete this test date
' <VB WATCH>
946        On Error GoTo vbwErrHandler
947        Const VBWPROCNAME = "frmPLCData.cmdDeleteTestDate_Click"
948        If vbwProtector.vbwTraceProc Then
949            Dim vbwProtectorParameterString As String
950            If vbwProtector.vbwTraceParameters Then
951                vbwProtectorParameterString = "()"
952            End If
953            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
954        End If
' </VB WATCH>
955        Dim Answer As Integer
956        Answer = MsgBox("You are about to delete the following record: S/N = " & rsTestData.Fields("SerialNumber") & " and Test Date = " & rsTestSetup.Fields("Date") & "!  Do you want to continue?", vbCritical Or vbYesNo, "Ready to Delete")
957        If Answer = vbYes Then
958            rsTestSetup.Delete
959            rsTestSetup.Update
960            cmdFindPump_Click
961        End If
' <VB WATCH>
962        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
963        Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdDeleteTestDate_Click"

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
            vbwReportVariable "Answer", Answer
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmdEnterPumpData_Click()
           'store the data on the screen to the pump (pumpdata)
' <VB WATCH>
964        On Error GoTo vbwErrHandler
965        Const VBWPROCNAME = "frmPLCData.cmdEnterPumpData_Click"
966        If vbwProtector.vbwTraceProc Then
967            Dim vbwProtectorParameterString As String
968            If vbwProtector.vbwTraceParameters Then
969                vbwProtectorParameterString = "()"
970            End If
971            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
972        End If
' </VB WATCH>
973        Dim d As Integer
974        Dim sSearch As String
975        Dim ans As Integer
976        Dim boWriteDataWritten As Boolean


           'check for a serial number
977        If LenB(txtSN.Text) = 0 Then
978            MsgBox "You must have a Serial Number to enter data.  Data has not been saved."
' <VB WATCH>
979        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
980            Exit Sub
981        End If

           'check to make sure most entries are filled in
982        If LenB(txtModelNo.Text) = 0 And optMfr(0).value = True Then
983            MsgBox "You need to enter a MODEL NO before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
984        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
985            Exit Sub
986        End If
987        If LenB(txtSalesOrderNumber.Text) = 0 Then
988            If InStr(1, txtSN.Text, "-") <> 0 Then
989                txtSalesOrderNumber.Text = Mid$(txtSN.Text, 1, InStr(1, txtSN.Text, "-") - 1)
990            End If
991        End If
992        If LenB(txtSalesOrderNumber.Text) = 0 Then
993            MsgBox "You need to enter a SALES ORDER NUMBER before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
994        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
995            Exit Sub
996        End If

997        If cmbMotor.ListIndex = -1 And optMfr(0).value = True Then
998            MsgBox "You need to pick a MOTOR before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
999        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1000           Exit Sub
1001       End If

1002       If cmbStatorFill.ListIndex = -1 And optMfr(0).value = True Then    'set default
1003           cmbStatorFill.ListIndex = 0
1004       End If

1005       If cmbModel.ListIndex = -1 And optMfr(0).value = True Then
1006           MsgBox "You need to pick a MODEL before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1007       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1008           Exit Sub
1009       End If

1010       If cmbModelGroup.ListIndex = -1 And optMfr(0).value = True Then
1011           MsgBox "You need to pick a MODEL GROUP before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1012       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1013           Exit Sub
1014       End If


1015       If cmbDesignPressure.ListIndex = -1 And optMfr(0).value = True Then
1016           MsgBox "You need to pick a DESIGN PRESSURE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1017       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1018           Exit Sub
1019       End If

1020       If cmbCirculationPath.ListIndex = -1 And optMfr(0).value = True Then
1021           MsgBox "You need to pick a CIRCULATION PATH before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1022       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1023           Exit Sub
1024       End If

1025       If cmbRPM.ListIndex = -1 And optMfr(0).value = True Then
1026           MsgBox "You need to pick an RPM before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1027       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1028           Exit Sub
1029       End If

       'check TEMC dropdowns

1030       If cmbTEMCAdapter.ListIndex = -1 And optMfr(0).value = False Then
1031           MsgBox "You need to pick a TEMC ADAPTER before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1032       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1033           Exit Sub
1034       End If

1035       If cmbTEMCAdditions.ListIndex = -1 And optMfr(0).value = False Then
1036           MsgBox "You need to pick TEMC ADDITIONS before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1037       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1038           Exit Sub
1039       End If

1040       If cmbTEMCCirculation.ListIndex = -1 And optMfr(0).value = False Then
1041           MsgBox "You need to pick a TEMC CIRCULATION before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1042       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1043           Exit Sub
1044       End If

1045       If cmbTEMCDesignPressure.ListIndex = -1 And optMfr(0).value = False Then
1046           MsgBox "You need to pick a TEMC DESIGN PRESSURE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1047       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1048           Exit Sub
1049       End If

1050       If cmbTEMCDivisionType.ListIndex = -1 And optMfr(0).value = False Then
1051           MsgBox "You need to pick a TEMC DIVISION TYPE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1052       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1053           Exit Sub
1054       End If

1055       If cmbTEMCImpellerType.ListIndex = -1 And optMfr(0).value = False Then
1056           MsgBox "You need to pick a TEMC IMPELLER TYPE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1057       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1058           Exit Sub
1059       End If

1060       If cmbTEMCInsulation.ListIndex = -1 And optMfr(0).value = False Then
1061           MsgBox "You need to pick a TEMC INSULATION TYPE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1062       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1063           Exit Sub
1064       End If

1065       If cmbTEMCJacketGasket.ListIndex = -1 And optMfr(0).value = False Then
1066           MsgBox "You need to pick a TEMC JACKET GASKET before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1067       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1068           Exit Sub
1069       End If

1070       If cmbTEMCMaterials.ListIndex = -1 And optMfr(0).value = False Then
1071           MsgBox "You need to pick TEMC MATERIALS before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1072       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1073           Exit Sub
1074       End If

1075       If cmbTEMCModel.ListIndex = -1 And optMfr(0).value = False Then
1076           MsgBox "You need to pick a TEMC MODEL before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1077       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1078           Exit Sub
1079       End If

1080       If cmbTEMCNominalImpSize.ListIndex = -1 And optMfr(0).value = False Then
1081           MsgBox "You need to pick a TEMC NOMINAL IMPELLER SIZE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1082       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1083           Exit Sub
1084       End If

1085       If cmbTEMCNominalDischargeSize.ListIndex = -1 And optMfr(0).value = False Then
1086           MsgBox "You need to pick a TEMC NOMINAL DISCHARGE SIZE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1087       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1088           Exit Sub
1089       End If

1090       If cmbTEMCNominalSuctionSize.ListIndex = -1 And optMfr(0).value = False Then
1091           MsgBox "You need to pick a TEMC NOMINAL SUCTION SIZE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1092       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1093           Exit Sub
1094       End If

1095       If cmbTEMCOtherMotor.ListIndex = -1 And optMfr(0).value = False Then
1096           MsgBox "You need to pick a TEMC OTHER MOTOR before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1097       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1098           Exit Sub
1099       End If

1100       If cmbTEMCPumpStages.ListIndex = -1 And optMfr(0).value = False Then
1101           MsgBox "You need to pick TEMC NUMBER OF PUMP STAGES before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1102       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1103           Exit Sub
1104       End If

1105       If cmbTEMCTRG.ListIndex = -1 And optMfr(0).value = False Then
1106           MsgBox "You need to pick a TEMC TRG before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1107       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1108           Exit Sub
1109       End If

1110       If cmbTEMCVoltage.ListIndex = -1 And optMfr(0).value = False Then
1111           MsgBox "You need to pick a TEMC VOLTAGE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1112       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1113           Exit Sub
1114       End If

1115       If LenB(txtTEMCFrameNumber.Text) = 0 And optMfr(0).value = False Then
1116           MsgBox "You need to enter a TEMC FRAME NUMBER before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1117       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1118           Exit Sub
1119       End If


1120       If Not boFoundPump Then     'if we havent found a pump in the database, add it
1121           rsPumpData.AddNew
1122           boWriteDataWritten = False
1123       Else    'else, find the entry
1124           sSearch = "Serialnumber = '" & frmPLCData.txtSN.Text & "'"
1125           rsPumpData.MoveFirst
1126           rsPumpData.Find sSearch, , adSearchForward, 1
1127           boWriteDataWritten = True
1128       End If

1129       If Not IsNull(rsPumpData!DataWritten) Or rsPumpData!DataWritten = True Then
1130           ans = MsgBox("You have already entered data for this pump.  Do you want to overwrite the data?", vbDefaultButton2 + vbYesNo, "Overwrite Data?")
1131           If ans = vbNo Then
1132               rsPumpData!DataWritten = True
1133               rsPumpData.Update   'update datawritten
' <VB WATCH>
1134       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1135               Exit Sub
1136           End If
1137       End If

1138       rsPumpData!SerialNumber = frmPLCData.txtSN.Text
1139       rsPumpData!ModelNumber = frmPLCData.txtModelNo.Text
1140       rsPumpData!SalesOrderNumber = frmPLCData.txtSalesOrderNumber.Text
1141       rsPumpData!ShipToCustomer = frmPLCData.txtShpNo.Text
1142       rsPumpData!BillToCustomer = frmPLCData.txtBilNo.Text
1143       rsPumpData!ApplicationFluid = frmPLCData.txtLiquid
1144       rsPumpData!NPSHFile = frmPLCData.txtNPSHFileLocation.Text
1145       rsPumpData!RVSPartNo = frmPLCData.txtRVSPartNo.Text
1146       rsPumpData!CustPN = frmPLCData.txtXPartNum.Text
1147       rsPumpData!CustPO = frmPLCData.txtCustPONum.Text

1148       If Len(frmPLCData.txtViscosity) <> 0 Then
1149           rsPumpData!ApplicationViscosity = frmPLCData.txtViscosity
1150       End If

1151       If frmPLCData.chkSuperMarketFeathered.value = Checked Then
1152           rsPumpData!Field1 = "Feathered"
1153       Else
1154           rsPumpData!Field1 = ""
1155       End If

1156       If LenB(txtSpGr.Text) <> 0 Then
1157           If Not IsNumeric(frmPLCData.txtSpGr.Text) Then
1158               MsgBox "Specific Gravity must be a number."
' <VB WATCH>
1159       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1160               Exit Sub
1161           End If
1162           rsPumpData!SpGr = frmPLCData.txtSpGr.Text
1163       End If
1164       If LenB(txtImpellerDia.Text) <> 0 Then
1165           If Not IsNumeric(frmPLCData.txtImpellerDia.Text) Then
1166               MsgBox "Impeller Diameter must be a number."
' <VB WATCH>
1167       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1168               Exit Sub
1169           End If
1170           rsPumpData!impellerdia = frmPLCData.txtImpellerDia.Text
1171       End If
1172       If LenB(txtDesignFlow.Text) <> 0 Then
1173           rsPumpData!designflow = frmPLCData.txtDesignFlow.Text
1174       End If
1175       If LenB(txtDesignTDH.Text) <> 0 Then
1176           rsPumpData!designtdh = frmPLCData.txtDesignTDH.Text
1177       End If
1178       If LenB(txtRemarks.Text) <> 0 Then
1179           rsPumpData!Remarks = txtRemarks.Text
1180       End If

1181       If optMfr(0).value = True Then
1182           d = cmbMotor.ItemData(cmbMotor.ListIndex)
1183           rsPumpData!Motor = d
1184           d = cmbStatorFill.ItemData(cmbStatorFill.ListIndex)
1185           rsPumpData!StatorFill = d
1186            d = cmbDesignPressure.ItemData(cmbDesignPressure.ListIndex)
1187           rsPumpData!DesignPressure = d
1188           d = cmbCirculationPath.ItemData(cmbCirculationPath.ListIndex)
1189           rsPumpData!CirculationPath = d
1190           d = cmbRPM.ItemData(cmbRPM.ListIndex)
1191           rsPumpData!RPM = d
1192           d = cmbModel.ItemData(cmbModel.ListIndex)
1193           rsPumpData!Model = d
1194           d = cmbModelGroup.ItemData(cmbModelGroup.ListIndex)
1195           rsPumpData!ModelGroup = d
1196       End If
       '   TEMC fields
1197       If optMfr(0).value = False Then
1198           d = cmbTEMCAdapter.ItemData(cmbTEMCAdapter.ListIndex)
1199           rsPumpData!TEMCAdapter = d

1200           d = cmbTEMCAdditions.ItemData(cmbTEMCAdditions.ListIndex)
1201           rsPumpData!TEMCAdditions = d

1202           d = cmbTEMCCirculation.ItemData(cmbTEMCCirculation.ListIndex)
1203           rsPumpData!TEMCcirculation = d

1204           d = cmbTEMCDesignPressure.ItemData(cmbTEMCDesignPressure.ListIndex)
1205           rsPumpData!TEMCDesignpressure = d

1206           d = cmbTEMCDivisionType.ItemData(cmbTEMCDivisionType.ListIndex)
1207           rsPumpData!TEMCDivisionType = d

1208           d = cmbTEMCImpellerType.ItemData(cmbTEMCImpellerType.ListIndex)
1209           rsPumpData!TEMCImpellerType = d

1210           d = cmbTEMCInsulation.ItemData(cmbTEMCInsulation.ListIndex)
1211           rsPumpData!TEMCInsulation = d

1212           d = cmbTEMCJacketGasket.ItemData(cmbTEMCJacketGasket.ListIndex)
1213           rsPumpData!TEMCJacketGasket = d

1214           d = cmbTEMCMaterials.ItemData(cmbTEMCMaterials.ListIndex)
1215           rsPumpData!TEMCMaterials = d

1216           d = cmbTEMCModel.ItemData(cmbTEMCModel.ListIndex)
1217           rsPumpData!TEMCModel = d

1218           d = cmbTEMCNominalImpSize.ItemData(cmbTEMCNominalImpSize.ListIndex)
1219           rsPumpData!TEMCNominalImpSize = d

1220           d = cmbTEMCNominalDischargeSize.ItemData(cmbTEMCNominalDischargeSize.ListIndex)
1221           rsPumpData!TEMCNominalDischargeSize = d

1222           d = cmbTEMCNominalSuctionSize.ItemData(cmbTEMCNominalSuctionSize.ListIndex)
1223           rsPumpData!TEMCNominalSuctionSize = d

1224           d = cmbTEMCOtherMotor.ItemData(cmbTEMCOtherMotor.ListIndex)
1225           rsPumpData!TEMCOtherMotor = d

1226           d = cmbTEMCPumpStages.ItemData(cmbTEMCPumpStages.ListIndex)
1227           rsPumpData!TEMCPumpStages = d

1228           d = cmbTEMCTRG.ItemData(cmbTEMCTRG.ListIndex)
1229           rsPumpData!TEMCTRG = d

1230           d = cmbTEMCVoltage.ItemData(cmbTEMCVoltage.ListIndex)
1231           rsPumpData!TEMCVoltage = d

1232           If LenB(txtTEMCFrameNumber.Text) <> 0 Then
1233               rsPumpData!TEMCFrameNumber = frmPLCData.txtTEMCFrameNumber.Text
1234           End If
1235       End If

1236       rsPumpData!ChempumpPump = optMfr(0).value

1237       rsPumpData!Approved = False

       'added from TEMC Inspection Report
1238       If Len(txtJobNum.Text) <> 0 Then
1239           rsPumpData!JobNumber = txtJobNum.Text
1240       End If

1241       If Len(txtNoPhases.Text) <> 0 Then
1242           rsPumpData!Phases = txtNoPhases.Text
1243       End If

1244       If Len(txtExpClass.Text) <> 0 Then
1245           rsPumpData!ExpClass = txtExpClass.Text
1246       End If

1247       If Len(txtThermalClass.Text) <> 0 Then
1248           rsPumpData!ThermalClass = txtThermalClass.Text
1249       End If

1250       rsPumpData!NPSHr = Val(txtNPSHr.Text)
1251       rsPumpData!LiquidTemperature = Val(txtLiquidTemperature.Text)
1252       rsPumpData!RatedInputPower = Val(txtRatedInputPower.Text)
1253       rsPumpData!FLCurrent = Val(txtAmps.Text)





1254       If boWriteDataWritten Then
1255           rsPumpData!DataWritten = True
1256       Else
1257           rsPumpData!DataWritten = False
1258       End If

           'write the data into the database
1259       rsPumpData.Update
1260       boFoundPump = True

           'enter a new test date if it's a new entry
1261       If Not boWriteDataWritten Then


1262           cmdAddNewTestDate_Click
1263       End If
' <VB WATCH>
1264       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1265       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdEnterPumpData_Click"

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
            vbwReportVariable "d", d
            vbwReportVariable "sSearch", sSearch
            vbwReportVariable "ans", ans
            vbwReportVariable "boWriteDataWritten", boWriteDataWritten
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub cmdEnterTestData_Click()
           ' save the data on the screen to test data at the selected run
' <VB WATCH>
1266       On Error GoTo vbwErrHandler
1267       Const VBWPROCNAME = "frmPLCData.cmdEnterTestData_Click"
1268       If vbwProtector.vbwTraceProc Then
1269           Dim vbwProtectorParameterString As String
1270           If vbwProtector.vbwTraceParameters Then
1271               vbwProtectorParameterString = "()"
1272           End If
1273           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1274       End If
' </VB WATCH>
1275       Dim sSearch As String
1276       Dim ans As Integer

           'if we didn't find the test setup, can't enter test data
1277       If Not boFoundTestSetup Then
1278           MsgBox "You must enter Test Setup Data before entering the Test Data"
' <VB WATCH>
1279       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1280           Exit Sub
1281       End If

           'if we don't find data in the test database, add records
1282       If boFoundTestData = False Then     'add 8 records for 8 tests
1283           AddTestData
1284           rsTestData.MoveFirst
1285       Else        'find the data in the database
1286           sSearch = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"
1287           rsTestData.MoveFirst
1288           rsTestData.Filter = sSearch
1289       End If

           'find the desired record from the form
1290       rsTestData.MoveFirst
1291       rsTestData.Move UpDown1.value - 1

1292       If rsTestData!DataWritten = True Then
1293           ans = MsgBox("You have already entered data for this test.  Do you want to overwrite the data?", vbYesNo + vbDefaultButton2, "Data Already Entered")
1294           If ans = vbNo Then
' <VB WATCH>
1295       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1296               Exit Sub
1297           End If
1298       End If

1299       rsEff.MoveFirst
1300       rsEff.Move UpDown1.value - 1

1301       If LenB(txtV1.Text) <> 0 Then
1302           rsTestData!VoltageA = Val(txtV1.Text)
1303       End If

1304       If LenB(txtV2.Text) <> 0 Then
1305           rsTestData!VoltageB = Val(txtV2.Text)
1306       End If

1307       If LenB(txtV3.Text) <> 0 Then
1308           rsTestData!VoltageC = Val(txtV3.Text)
1309       End If

1310       If LenB(txtI1.Text) <> 0 Then
1311           rsTestData!CurrentA = Val(txtI1.Text)
1312       End If

1313       If LenB(txtI2.Text) <> 0 Then
1314           rsTestData!CurrentB = Val(txtI2.Text)
1315       End If

1316       If LenB(txtI3.Text) <> 0 Then
1317           rsTestData!CurrentC = Val(txtI3.Text)
1318       End If

1319       If LenB(txtP1.Text) <> 0 Then
1320           rsTestData!PowerA = Val(txtP1.Text)
1321       End If

1322       If LenB(txtP2.Text) <> 0 Then
1323           rsTestData!PowerB = Val(txtP2.Text)
1324       End If

1325       If LenB(txtP3.Text) <> 0 Then
1326           rsTestData!PowerC = Val(txtP3.Text)
1327       End If

1328       If LenB(txtKW.Text) <> 0 Then
1329           rsTestData!TotalPower = Val(txtKW.Text)
1330       End If

1331       rsTestData!Flow = Val(txtFlowDisplay.Text)
1332       rsTestData!DischargePressure = Val(txtDischargeDisplay.Text)
1333       rsTestData!SuctionPressure = Val(txtSuctionDisplay.Text)
1334       rsTestData!TemperatureSuction = Val(txtTemperatureDisplay.Text)

1335       rsTestData!TC1 = Val(txtTC1Display.Text)
1336       rsTestData!TC2 = Val(txtTC2Display.Text)
1337       rsTestData!TC3 = Val(txtTC3Display.Text)
1338       rsTestData!TC4 = Val(txtTC4Display.Text)

1339       rsTestData!CircFlow = Val(txtAI1Display.Text)
1340       rsTestData!RBHTemp = Val(txtAI2Display.Text)
1341       rsTestData!RBHPress = Val(txtAI3Display.Text)
1342       rsTestData!AI4 = Val(txtAI4Display.Text)

1343       rsTestData!ValvePosition = Val(txtValvePosition.Text)
1344       rsTestData!SetPoint = Val(txtSetPoint.Text)

1345       If LenB(txtThrustBal.Text) <> 0 Then
1346           rsTestData!ThrustBalance = txtThrustBal.Text
1347       End If

1348       If LenB(txtVibAx.Text) <> 0 Then
1349           rsTestData!VibrationX = txtVibAx.Text
1350       End If

1351       If LenB(txtVibRad.Text) <> 0 Then
1352           rsTestData!VibrationY = txtVibRad.Text
1353       End If

1354       If LenB(txtTEMCTRGReading.Text) <> 0 Then
1355           rsTestData!TEMCTRG = txtTEMCTRGReading.Text
1356       Else
1357           rsTestData!TEMCTRG = 0
1358       End If

1359       If LenB(txtRPM.Text) <> 0 Then
1360           rsTestData!RPM = txtRPM.Text
1361       End If

1362       If LenB(txtTestRemarks.Text) <> 0 Then
1363           rsTestData!Remarks = txtTestRemarks.Text
1364       Else
1365           rsTestData!Remarks = " "
1366       End If

1367       If LenB(txtTEMCTRGReading.Text) <> 0 Then
1368           rsTestData!TEMCTRG = txtTEMCTRGReading.Text
1369       End If

1370       If LenB(txtTEMCFrontThrust.Text) <> 0 Then
1371           rsTestData!TEMCFrontThrust = txtTEMCFrontThrust.Text
1372       End If

1373       If LenB(txtTEMCRearThrust.Text) <> 0 Then
1374           rsTestData!TEMCRearThrust = txtTEMCRearThrust.Text
1375       End If

1376       If LenB(txtTEMCMomentArm.Text) <> 0 Then
1377           rsTestData!TEMCMomentArm = txtTEMCMomentArm.Text
1378       End If

1379       If LenB(txtTEMCThrustRigPressure.Text) <> 0 Then
1380           rsTestData!TEMCThrustRigPressure = txtTEMCThrustRigPressure.Text
1381       End If

1382       If LenB(txtTEMCViscosity.Text) <> 0 Then
1383           rsTestData!TEMCViscosity = txtTEMCViscosity.Text
1384       End If

1385       If LenB(txtNPSHa.Text) <> 0 Then
1386           rsTestData!NPSHa = txtNPSHa.Text
1387       End If

1388       rsTestData!Approved = False

1389       rsTestData!DataWritten = True

           'update the database
1390       rsTestData.Update

1391       DoEfficiencyCalcs
1392       rsEff.Update

           'update the form
1393       DataGrid1.Refresh
1394       DataGrid2.Refresh

1395       FixPointsToPlot

' <VB WATCH>
1396       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1397       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdEnterTestData_Click"

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
            vbwReportVariable "sSearch", sSearch
            vbwReportVariable "ans", ans
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub cmdEnterTestSetupData_Click()
           'save the data on the screen to testsetupdata
' <VB WATCH>
1398       On Error GoTo vbwErrHandler
1399       Const VBWPROCNAME = "frmPLCData.cmdEnterTestSetupData_Click"
1400       If vbwProtector.vbwTraceProc Then
1401           Dim vbwProtectorParameterString As String
1402           If vbwProtector.vbwTraceParameters Then
1403               vbwProtectorParameterString = "()"
1404           End If
1405           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1406       End If
' </VB WATCH>
1407       Dim I As Integer
1408       Dim d As Integer
1409       Dim sSearch As String
1410       Dim ans As Integer
1411       Dim boWriteDataWritten As Boolean

           'check for a serial number
1412       If LenB(txtSN.Text) = 0 Then
1413           MsgBox "You must have a Serial Number to enter data.", vbOKOnly + vbExclamation, "Cannot Enter Data"
' <VB WATCH>
1414       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1415           Exit Sub
1416       End If

1417       If Pressed = True Then
1418           If Me.cmbDischDia.ListIndex = -1 Or Me.cmbSuctDia.ListIndex = -1 Or Val(Me.txtSuctHeight.Text) = 0 Or Val(Me.txtDischHeight.Text) = 0 Then
1419               MsgBox "You must have Discharge Diameter AND Suction Diameter AND Suction Height AND Discharge Height entered to enter data.", vbOKOnly + vbExclamation, "Cannot Enter Data"
' <VB WATCH>
1420       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1421               Exit Sub
1422           End If
1423       End If

1424       Pressed = True
1425       If Not boFoundTestSetup Then    'if we didn't find any test setup, add a record
1426           rsTestSetup.AddNew
1427           cmbTestDate.AddItem Now
1428           cmbTestDate.ListIndex = cmbTestDate.NewIndex
1429           cmdAddNewBalanceHoles.Visible = True
1430           boFoundTestSetup = True
1431           boWriteDataWritten = False
1432           rsTestSetup!DataWritten = False
1433       Else    'find the record and display
1434           sSearch = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"
1435           rsTestSetup.MoveFirst
1436           rsTestSetup.Filter = sSearch
1437           If Not boCanApprove Then
       '            cmdAddNewBalanceHoles.Visible = False
1438           End If
1439           boWriteDataWritten = True
1440       End If

1441       If rsTestSetup!DataWritten = True Then
1442           ans = MsgBox("Data has already been entered for this test date.  Do you want to overwrite it?", vbYesNo + vbDefaultButton2, "Data Exists")
1443           If ans = vbNo Then
' <VB WATCH>
1444       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1445               Exit Sub
1446           End If
1447       End If

1448       rsTestSetup!SerialNumber = txtSN
1449       rsTestSetup!Date = cmbTestDate.List(cmbTestDate.ListIndex)

1450       I = cmbFlowMeter.ListIndex
1451       If I = -1 Then
1452           d = 1
1453           rsTestSetup!FlowMeterID = d
1454       Else
1455           d = cmbLoopNumber.ItemData(I)
1456           rsTestSetup!FlowMeterID = d
1457       End If

1458       I = cmbSuctionPressureTransducer.ListIndex
1459       If I = -1 Then
1460           d = 1
1461           rsTestSetup!suctionid = d
1462       Else
1463           d = cmbLoopNumber.ItemData(I)
1464           rsTestSetup!suctionid = d
1465       End If

1466       I = cmbDischargePressureTransducer.ListIndex
1467       If I = -1 Then
1468           d = 1
1469           rsTestSetup!dischid = d
1470       Else
1471           d = cmbLoopNumber.ItemData(I)
1472           rsTestSetup!dischid = d
1473       End If

1474       I = cmbTemperatureTransducer.ListIndex
1475       If I = -1 Then
1476           d = 1
1477           rsTestSetup!temperatureid = d
1478       Else
1479           d = cmbLoopNumber.ItemData(I)
1480           rsTestSetup!temperatureid = d
1481       End If

1482       I = Me.cmbCirculationFlowMeter.ListIndex
1483       If I = -1 Or I > 4 Then
1484           d = 5
1485           rsTestSetup!MagFlowID = d
1486       Else
1487           d = cmbLoopNumber.ItemData(I) + 4
1488           rsTestSetup!MagFlowID = d
1489       End If


1490       If LenB(txtHDCor.Text) <> 0 Then
1491           rsTestSetup!HDCor = txtHDCor
1492       Else
1493           rsTestSetup!HDCor = 0
1494       End If
1495       If LenB(txtKWMult.Text) <> 0 Then
1496           rsTestSetup!kwmult = txtKWMult
1497       Else
1498           rsTestSetup!kwmult = 1
1499       End If
1500       If LenB(txtWho.Text) <> 0 Then
1501           rsTestSetup!who = txtWho
1502       Else
1503           rsTestSetup!who = vbNullString
1504       End If
1505       If LenB(txtRMA.Text) <> 0 Then
1506           rsTestSetup!RMA = txtRMA
1507       Else
1508           rsTestSetup!RMA = vbNullString
1509       End If
1510       If LenB(frmPLCData.txtDischHeight) <> 0 Then
1511           rsTestSetup!DischargeGageHeight = Val(txtDischHeight)
1512       Else
1513           rsTestSetup!DischargeGageHeight = 0
1514       End If
1515       If LenB(frmPLCData.txtSuctHeight) <> 0 Then
1516           rsTestSetup!SuctionGageHeight = Val(txtSuctHeight)
1517       Else
1518           rsTestSetup!SuctionGageHeight = 0
1519       End If
1520       If LenB(frmPLCData.txtTestSetupRemarks.Text) <> 0 Then
1521           rsTestSetup!Remarks = txtTestSetupRemarks.Text
1522       Else
1523           rsTestSetup!Remarks = vbNullString
1524       End If
1525       If LenB(frmPLCData.txtVFDFreq.Text) <> 0 Then
1526           rsTestSetup!VFDFrequency = txtVFDFreq.Text
1527       Else
1528           rsTestSetup!VFDFrequency = 0
1529       End If

1530       I = cmbOrificeNumber.ListIndex
1531       If I = -1 Then
1532           d = 18      'entry for None
1533       Else
1534           d = cmbOrificeNumber.ItemData(I)
1535       End If
1536       rsTestSetup!orificenumber = d

1537       If LenB(txtEndPlay.Text) <> 0 Then
1538           rsTestSetup!Endplay = Val(frmPLCData.txtEndPlay.Text)
1539       Else
1540           rsTestSetup!Endplay = 0
1541       End If

1542       If LenB(txtGGap.Text) <> 0 Then
1543           rsTestSetup!GGAP = Val(frmPLCData.txtGGap.Text)
1544       Else
1545           rsTestSetup!GGAP = 0
1546       End If

1547       If LenB(txtOtherMods.Text) <> 0 Then
1548           rsTestSetup!OtherMods = txtOtherMods.Text
1549       Else
1550           rsTestSetup!OtherMods = vbNullString
1551       End If

1552       rsTestSetup!Approved = False

1553       I = cmbLoopNumber.ListIndex
1554       If I = -1 Then
1555           d = 1
1556           rsTestSetup!loopnumber = d
1557       Else
1558           d = cmbLoopNumber.ItemData(I)
1559           rsTestSetup!loopnumber = d
1560       End If

1561       I = cmbSuctDia.ListIndex
1562       If I = -1 Then
1563           d = -1
1564       Else
1565           d = cmbSuctDia.ItemData(I)
1566           rsTestSetup!SuctDiam = d
1567       End If

1568       I = cmbDischDia.ListIndex
1569       If I = -1 Then
1570           d = -1
1571       Else
1572           d = cmbDischDia.ItemData(I)
1573           rsTestSetup!DischDiam = d
1574       End If

1575       I = cmbTachID.ListIndex
1576       If I = -1 Then
1577           d = 1
1578           rsTestSetup!tachid = d
1579       Else
1580           d = cmbTachID.ItemData(I)
1581           rsTestSetup!tachid = d
1582       End If

1583       I = cmbAnalyzerNo.ListIndex
1584       If I = -1 Then
1585           d = 1
1586       Else
1587           d = cmbAnalyzerNo.ItemData(I)
1588       End If
1589       rsTestSetup!analyzerno = d

1590       I = cmbTestSpec.ListIndex
1591       If I = -1 Then
1592           d = 1
1593       Else
1594           d = cmbTestSpec.ItemData(I)
1595       End If
1596       rsTestSetup!testspec = d

1597       I = cmbVoltage.ListIndex
1598       If I = -1 Then
1599           d = 1
1600       Else
1601           d = cmbVoltage.ItemData(I)
1602       End If
1603       rsTestSetup!Voltage = d

1604       I = cmbFrequency.ListIndex
1605       If I = -1 Then
1606           d = 1
1607       Else
1608           d = cmbFrequency.ItemData(I)
1609       End If
1610       rsTestSetup!Frequency = d

1611       I = cmbMounting.ListIndex
1612       If I = -1 Then
1613           d = 1
1614       Else
1615           d = cmbMounting.ItemData(I)
1616       End If
1617       rsTestSetup!Mounting = d

1618       I = cmbPLCNo.ListIndex
1619       If I = -1 Then
1620           d = 8
1621       Else
1622           d = cmbPLCNo.ItemData(I)
1623       End If
1624       rsTestSetup!PLCNo = d

1625       rsTestSetup!ImpFeathered = chkFeathered.value

1626       If chkTrimmed.value = 1 Then
1627           rsTestSetup!ImpTrimmed = Val(txtImpTrim)
1628       Else
1629           rsTestSetup!ImpTrimmed = 0
1630       End If
1631       chkTrimmed_Click

1632       If chkOrifice.value = 1 Then
1633           rsTestSetup!PumpDischOrifice = Val(txtOrifice)
1634       Else
1635           rsTestSetup!PumpDischOrifice = 0
1636       End If
1637       chkOrifice_Click

1638       If chkCircOrifice.value = 1 Then
1639           rsTestSetup!CircFlowOrifice = Val(txtCircOrifice)
1640       Else
1641           rsTestSetup!CircFlowOrifice = 0
1642       End If
1643       chkCircOrifice_Click

1644       chkBalanceHoles_Click

1645       If chkNPSH.value = 1 Then
1646           txtNPSHFile.Visible = True
1647           rsTestSetup!NPSHFile = txtNPSHFile
1648       Else
1649           rsTestSetup!NPSHFile = vbNullString
1650           txtNPSHFile.Visible = False
1651       End If

1652       If chkPictures.value = 1 Then
1653           txtPicturesFile.Visible = True
1654           rsTestSetup!PictureFile = txtPicturesFile
1655       Else
1656           rsTestSetup!PictureFile = vbNullString
1657           txtPicturesFile.Visible = False
1658       End If

1659       If chkVibration.value = 1 Then
1660           txtVibrationFile.Visible = True
1661           rsTestSetup!VibrationFile = txtVibrationFile
1662       Else
1663           rsTestSetup!VibrationFile = vbNullString
1664           txtVibrationFile.Visible = False
1665       End If

1666       If boWriteDataWritten Then
1667           rsTestSetup!DataWritten = True
1668       Else
1669           rsTestSetup!DataWritten = False
1670       End If

           'for TEMC Inspection Report
1671       If LenB(frmPLCData.txtTestAndInspection(0).Text) <> 0 Then
1672           rsTestSetup!InsulationMeggerVolts = frmPLCData.txtTestAndInspection(0).Text
1673       Else
1674           rsTestSetup!InsulationMeggerVolts = ""
1675       End If

1676       If LenB(frmPLCData.txtTestAndInspection(1).Text) <> 0 Then
1677           rsTestSetup!InsulationMegOhms = frmPLCData.txtTestAndInspection(1).Text
1678       Else
1679           rsTestSetup!InsulationMegOhms = ""
1680       End If

1681       If LenB(frmPLCData.txtTestAndInspection(2).Text) <> 0 Then
1682           rsTestSetup!DielectricVolts = frmPLCData.txtTestAndInspection(2).Text
1683       Else
1684           rsTestSetup!DielectricVolts = ""
1685       End If

1686       If LenB(frmPLCData.txtTestAndInspection(3).Text) <> 0 Then
1687           rsTestSetup!DielectricTime = frmPLCData.txtTestAndInspection(3).Text
1688       Else
1689           rsTestSetup!DielectricTime = ""
1690       End If

1691       If LenB(frmPLCData.txtTestAndInspection(4).Text) <> 0 Then
1692           rsTestSetup!HydrostaticValue = frmPLCData.txtTestAndInspection(4).Text
1693       Else
1694           rsTestSetup!HydrostaticValue = ""
1695       End If

1696       If LenB(frmPLCData.txtTestAndInspection(5).Text) <> 0 Then
1697           rsTestSetup!HydrostaticTime = frmPLCData.txtTestAndInspection(5).Text
1698       Else
1699           rsTestSetup!HydrostaticTime = ""
1700       End If

1701       If LenB(frmPLCData.txtTestAndInspection(6).Text) <> 0 Then
1702           rsTestSetup!PneumaticValue = frmPLCData.txtTestAndInspection(6).Text
1703       Else
1704           rsTestSetup!PneumaticValue = ""
1705       End If

1706       If LenB(frmPLCData.txtTestAndInspection(7).Text) <> 0 Then
1707           rsTestSetup!PneumaticTime = frmPLCData.txtTestAndInspection(7).Text
1708       Else
1709           rsTestSetup!PneumaticTime = ""
1710       End If

1711       I = cmbTestAndInspection(0).ListIndex
1712       If I = -1 Then
1713           rsTestSetup!HydrostaticUnits = ""
1714       Else
1715           rsTestSetup!HydrostaticUnits = cmbTestAndInspection(0).Text
1716       End If


1717       I = cmbTestAndInspection(1).ListIndex
1718       If I = -1 Then
1719           rsTestSetup!PneumaticUnits = ""
1720       Else
1721           rsTestSetup!PneumaticUnits = cmbTestAndInspection(1).Text
1722       End If

           'use abs to convert from 1 and 0 to boolean
1723       rsTestSetup!insulationgood = Abs(TestAndInspectionGood(0).value)
1724       rsTestSetup!DielectricGood = Abs(TestAndInspectionGood(1).value)
1725       rsTestSetup!HydrostaticGood = Abs(TestAndInspectionGood(2).value)
1726       rsTestSetup!PneumaticGood = Abs(TestAndInspectionGood(3).value)
1727       rsTestSetup!GeneralAppearanceGood = Abs(TestAndInspectionGood(4).value)
1728       rsTestSetup!OutlineDimensionsGood = Abs(TestAndInspectionGood(5).value)
1729       rsTestSetup!MotorNoLoadTestGood = Abs(TestAndInspectionGood(6).value)
1730       rsTestSetup!MotorLockedRotorTestGood = Abs(TestAndInspectionGood(7).value)
1731       rsTestSetup!HydrostaticTestGood = Abs(TestAndInspectionGood(8).value)
1732       rsTestSetup!HydraulicTestGood = Abs(TestAndInspectionGood(9).value)
1733       rsTestSetup!NPSHTestGood = Abs(TestAndInspectionGood(10).value)
1734       rsTestSetup!CleanPurgeSealGood = Abs(TestAndInspectionGood(11).value)
1735       rsTestSetup!PaintCheckGood = Abs(TestAndInspectionGood(12).value)
1736       rsTestSetup!NameplateGood = Abs(TestAndInspectionGood(13).value)
1737       rsTestSetup!SupervisorApproval = Abs(TestAndInspectionGood(14).value)

           'update the database
1738       rsTestSetup.Update

1739       If boFoundTestData = False Then     'add 8 records for 8 tests
1740           AddTestData
1741       End If

1742       rsTestSetup.Filter = vbNullString
' <VB WATCH>
1743       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1744       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdEnterTestSetupData_Click"

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
            vbwReportVariable "I", I
            vbwReportVariable "d", d
            vbwReportVariable "sSearch", sSearch
            vbwReportVariable "ans", ans
            vbwReportVariable "boWriteDataWritten", boWriteDataWritten
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub cmdExit_Click()
' <VB WATCH>
1745       On Error GoTo vbwErrHandler
1746       Const VBWPROCNAME = "frmPLCData.cmdExit_Click"
1747       If vbwProtector.vbwTraceProc Then
1748           Dim vbwProtectorParameterString As String
1749           If vbwProtector.vbwTraceParameters Then
1750               vbwProtectorParameterString = "()"
1751           End If
1752           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1753       End If
' </VB WATCH>
1754       End
' <VB WATCH>
1755       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1756       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdExit_Click"

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

Private Sub cmdFindMagtrols_Click()
' <VB WATCH>
1757       On Error GoTo vbwErrHandler
1758       Const VBWPROCNAME = "frmPLCData.cmdFindMagtrols_Click"
1759       If vbwProtector.vbwTraceProc Then
1760           Dim vbwProtectorParameterString As String
1761           If vbwProtector.vbwTraceParameters Then
1762               vbwProtectorParameterString = "()"
1763           End If
1764           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1765       End If
' </VB WATCH>
1766       FindMagtrols
' <VB WATCH>
1767       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1768       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdFindMagtrols_Click"

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

Private Sub cmdFindPump_Click()
           ' find the pump whose sn is shown
' <VB WATCH>
1769       On Error GoTo vbwErrHandler
1770       Const VBWPROCNAME = "frmPLCData.cmdFindPump_Click"
1771       If vbwProtector.vbwTraceProc Then
1772           Dim vbwProtectorParameterString As String
1773           If vbwProtector.vbwTraceParameters Then
1774               vbwProtectorParameterString = "()"
1775           End If
1776           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1777       End If
' </VB WATCH>
1778       Dim sAns As String
1779       Dim sSO As String
1780       Dim sParam As String
1781       Dim sName As String

1782       Dim I As Integer

           'clear the data
1783       BlankData

           'set TC and AI labels with default values
1784       txtTitle(0).Text = "TC 1"
1785       txtTitle(1).Text = "(F)"
1786       txtTitle(2).Text = "TC 2"
1787       txtTitle(3).Text = "(F)"
1788       txtTitle(4).Text = "TC 3"
1789       txtTitle(5).Text = "(F)"
1790       txtTitle(6).Text = "TC 4"
1791       txtTitle(7).Text = "(F)"
1792       txtTitle(20).Text = "Circ Flow"
1793       txtTitle(21).Text = "(GPM)"
1794       txtTitle(22).Text = "P1"
1795       txtTitle(23).Text = "(psig)"
1796       txtTitle(24).Text = "P2"
1797       txtTitle(25).Text = "(psig)"
1798       txtTitle(26).Text = "AI 4"
1799       txtTitle(27).Text = ""


1800       For I = 0 To 7
1801           lblAutoMan(I).Caption = "Auto"
1802       Next I

1803       lblAutoMan(5).Caption = "Man"
1804       lblAutoMan(6).Caption = "Man"

1805       txtFlowDisplay.Enabled = False
1806       txtSuctionDisplay.Enabled = False
1807       txtDischargeDisplay.Enabled = False
1808       txtTemperatureDisplay.Enabled = False
1809       txtAI1Display.Enabled = False
1810       txtAI2Display.Enabled = False
1811       txtAI3Display.Enabled = False
1812       txtAI4Display.Enabled = False


1813       cmdFindPump.Default = False

           'set all found booleans to false
       '    boUsingHP = False
1814       boFoundPump = False
1815       boPumpIsApproved = False
1816       boFoundTestSetup = False
1817       boFoundTestData = False


           'get rid of all test dates in combo box
1818       For I = cmbTestDate.ListCount - 1 To 0 Step -1
1819           cmbTestDate.RemoveItem 0
1820       Next I

1821       rsTestData.Filter = "SerialNumber = ''"

1822       DataGrid2.ClearFields
1823       ClearEff

1824       If rsPumpData.State = adStateOpen Then
1825           If rsPumpData.BOF = False Or rsPumpData.EOF = False Then
1826               rsPumpData.Update
1827           End If
1828           rsPumpData.Close
1829       End If

           'parse the serial number to make sure it is formed correctly
1830       Dim ok As Boolean
1831       ok = UCase(txtSN.Text) Like "[A-Z][A-Z][0-9][0-9][0-9][0-9][A-Z]" Or UCase(txtSN.Text) Like "[A-Z][A-Z][0-9][0-9][0-9][0-9][A-Z]-[0-9]" Or UCase(txtSN.Text) Like "[A-Z][A-Z][0-9][0-9][0-9][0-9][A-Z]-[0-9][0-9]" Or UCase(txtSN.Text) Like "[0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9][0-9]"
1832       If Not ok Then
1833           MsgBox "Serial Number must be 2 letters, 4 numbers, and 1 letter, or 10 numbers. Please re-enter.", vbOKOnly, "Serial Number not correctly formed."
' <VB WATCH>
1834       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1835           Exit Sub
1836       End If

           'find the pump listed in the Serial Number text box
1837       qyPumpData.ActiveConnection = cnPumpData
1838       qyPumpData.CommandText = "SELECT * From TempPumpData WHERE (((TempPumpData.SerialNumber)='" & _
                                    txtSN.Text & "'))"
1839       rsPumpData.CursorType = adOpenStatic
1840       rsPumpData.CursorLocation = adUseClient
1841       rsPumpData.Index = "SerialNumber"
1842       rsPumpData.Open qyPumpData
1843       boEpicorFound = False

1844       If rsPumpData.BOF = True And rsPumpData.EOF = True Then
               'if the bof=eof, we have an empty recordset
1845           boFoundPump = False
1846       Else
               'we found it
1847           boFoundPump = True
1848       End If

1849       If boFoundPump = False Then
               'not found in either database, try HP?
1850           sAns = MsgBox("Pump Not Found in the Database.  Look in Epicor?", vbYesNo, "Can't Find Pump")
1851           If sAns = vbNo Then     'new pump - don't get data from HP
1852               boUsingEpicor = False
1853           Else
1854               boUsingEpicor = True
       '            boUsingHP = False
1855           End If
       '        If boUsingEpicor = False Then
       '            sAns = MsgBox("Pump Not Found in the Database.  Look on the HP?", vbYesNo, "Can't Find Pump")
       '            If sAns = vbNo Then     'new pump - don't get data from HP
       '                 boUsingHP = False
       '            Else
       '                boUsingHP = True
       '            End If
       '        End If
1856           EnablePumpDataControls
1857           EnableTestSetupDataControls
1858           EnableTestDataControls
       '        BlankData               'clear any data on the screen
1859           cmdAddNewBalanceHoles.Visible = True

1860       End If

1861       If boFoundPump = True Then    'found the pump
1862           If rsPumpData.Fields("Approved") = True Then
1863               DisablePumpDataControls                         'if it's in the real database, don't allow changes here
1864               boPumpIsApproved = True
1865               lblPumpApproved.Visible = True
1866               If boCanApprove Then
1867                   cmdApprovePump.Caption = "Unapprove this pump"
1868               End If
1869               frmPLCData.cmdApproveTestDate.Enabled = True
1870           Else
1871               EnablePumpDataControls                          'it's in the temp database, allow changes
1872               boPumpIsApproved = False
1873               boTestDateIsApproved = False
1874               lblPumpApproved.Visible = False
1875               If boCanApprove Then
1876                   cmdApprovePump.Caption = "Approve this pump"
1877               End If
1878               cmdApproveTestDate.Caption = "You Must Approve Pump First"
1879               frmPLCData.cmdApproveTestDate.Enabled = False
1880           End If

               'found the pump, show the data
1881           txtModelNo.Text = rsPumpData.Fields("ModelNumber")
1882           frmPLCData.optMfr(0).value = rsPumpData.Fields("ChempumpPump")

1883           If rsPumpData.Fields("ChempumpPump") = True Then
1884               SetCombo cmbMotor, "Motor", rsPumpData
1885               SetCombo cmbDesignPressure, "DesignPressure", rsPumpData
1886               SetCombo cmbRPM, "RPM", rsPumpData
1887               SetCombo cmbCirculationPath, "CirculationPath", rsPumpData
1888               SetCombo cmbStatorFill, "StatorFill", rsPumpData
1889               SetCombo cmbModel, "Model", rsPumpData
1890               SetCombo cmbModelGroup, "ModelGroup", rsPumpData
1891               RatedKW = 999
1892           End If

               'set the TEMC data
1893           If rsPumpData.Fields("ChempumpPump") = False Then
1894               SetCombo cmbTEMCAdapter, "TEMCAdapter", rsPumpData
1895               SetCombo cmbTEMCAdditions, "TEMCAdditions", rsPumpData
1896               SetCombo cmbTEMCCirculation, "TEMCCirculation", rsPumpData
1897               SetCombo cmbTEMCDesignPressure, "TEMCDesignPressure", rsPumpData
1898               SetCombo cmbTEMCNominalDischargeSize, "TEMCNominalDischargeSize", rsPumpData
1899               SetCombo cmbTEMCDivisionType, "TEMCDivisionType", rsPumpData
1900               SetCombo cmbTEMCImpellerType, "TEMCImpellerType", rsPumpData
1901               SetCombo cmbTEMCInsulation, "TEMCInsulation", rsPumpData
1902               SetCombo cmbTEMCJacketGasket, "TEMCJacketGasket", rsPumpData
1903               SetCombo cmbTEMCMaterials, "TEMCMaterials", rsPumpData
1904               SetCombo cmbTEMCModel, "TEMCModel", rsPumpData
1905               SetCombo cmbTEMCNominalImpSize, "TEMCNominalImpSize", rsPumpData
1906               SetCombo cmbTEMCOtherMotor, "TEMCOtherMotor", rsPumpData
1907               SetCombo cmbTEMCPumpStages, "TEMCPumpStages", rsPumpData
1908               SetCombo cmbTEMCNominalSuctionSize, "TEMCNominalSuctionSize", rsPumpData
1909               SetCombo cmbTEMCTRG, "TEMCTRG", rsPumpData
1910               SetCombo cmbTEMCVoltage, "TEMCVoltage", rsPumpData
1911           End If

               'write ship to and bill to info
1912           If Not IsNull(rsPumpData.Fields("ShipToCustomer")) Then
1913               txtShpNo.Text = rsPumpData.Fields("ShipToCustomer")
1914           Else
1915               txtShpNo.Text = vbNullString
1916           End If

1917           If Not IsNull(rsPumpData.Fields("BillToCustomer")) Then
1918               txtBilNo.Text = rsPumpData.Fields("BillToCustomer")
1919           Else
1920               txtBilNo.Text = vbNullString
1921           End If

1922           sName = "ImpellerDia"
1923           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1924               sParam = rsPumpData.Fields(sName)
1925           Else
1926               sParam = vbNullString
1927           End If
1928           txtImpellerDia.Text = sParam

1929           sName = "DesignFlow"
1930           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1931               sParam = rsPumpData.Fields(sName)
1932           Else
1933               sParam = vbNullString
1934           End If
1935           txtDesignFlow.Text = sParam

1936           sName = "DesignTDH"
1937           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1938               sParam = rsPumpData.Fields(sName)
1939           Else
1940               sParam = vbNullString
1941           End If
1942           txtDesignTDH.Text = sParam

1943           sName = "SpGr"
1944           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1945               sParam = rsPumpData.Fields(sName)
1946           Else
1947               sParam = vbNullString
1948           End If
1949           txtSpGr.Text = sParam

1950           sName = "Remarks"
1951           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1952               sParam = rsPumpData.Fields(sName)
1953           Else
1954               sParam = vbNullString
1955           End If
1956           txtRemarks.Text = sParam

1957           sName = "SalesOrderNumber"
1958           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1959               sParam = rsPumpData.Fields(sName)
1960           Else
1961               sParam = vbNullString
1962           End If
1963           txtSalesOrderNumber.Text = sParam

1964           sName = "ApplicationFluid"
1965           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1966               sParam = rsPumpData.Fields(sName)
1967           Else
1968               sParam = vbNullString
1969           End If
1970           txtLiquid.Text = sParam

1971           sName = "NPSHFile"
1972           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1973               sParam = rsPumpData.Fields(sName)
1974           Else
1975               sParam = vbNullString
1976           End If
1977           txtNPSHFileLocation.Text = sParam

1978           sName = "RVSPartNo"
1979           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1980               sParam = rsPumpData.Fields(sName)
1981           Else
1982               sParam = vbNullString
1983           End If
1984           txtRVSPartNo.Text = sParam

1985           sName = "CustPN"
1986           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1987               sParam = rsPumpData.Fields(sName)
1988           Else
1989               sParam = vbNullString
1990           End If
1991           txtXPartNum.Text = sParam

1992           sName = "CustPO"
1993           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1994               sParam = rsPumpData.Fields(sName)
1995           Else
1996               sParam = vbNullString
1997           End If
1998           txtCustPONum.Text = sParam

               'make sure table has custpn - see if last three digits of model no are numeric
       '        sName = "SalesOrderNumber"
       '        If rsPumpData.Fields(sName).ActualSize <> 0 Then
       '            If IsNumeric(Right(rsPumpData.Fields("ModelNumber"), 3)) Then 'no sales order no, must be supermarket
       '                rsPumpData.Fields("CustPN") = rsPumpData.Fields("RVSPartNo")
       '            Else
       '                rsPumpData.Fields("CustPN") = rsPumpData.Fields("ModelNumber")
       '            End If
       '        End If

1999           sName = "ApplicationViscosity"
2000           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2001               sParam = Format(rsPumpData.Fields(sName), "#0.00")
2002           Else
2003               sParam = vbNullString
2004           End If
2005           txtViscosity.Text = sParam
2006           txtTEMCViscosity.Text = sParam


       'added from TEMC Inspection Report
2007           sName = "JobNumber"
2008           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2009               sParam = rsPumpData.Fields(sName)
2010           Else
2011               sParam = ""
2012           End If
2013           txtJobNum.Text = sParam

2014           sName = "Phases"
2015           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2016               sParam = rsPumpData.Fields(sName)
2017           Else
2018               sParam = vbNullString
2019           End If
2020           txtNoPhases.Text = sParam

2021           sName = "ThermalClass"
2022           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2023               sParam = rsPumpData.Fields(sName)
2024           Else
2025               sParam = vbNullString
2026           End If
2027           txtThermalClass.Text = sParam

2028           sName = "ExpClass"
2029           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2030               sParam = rsPumpData.Fields(sName)
2031           Else
2032               sParam = vbNullString
2033           End If
2034           txtExpClass.Text = sParam

2035           sName = "NPSHr"
2036           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2037               sParam = rsPumpData.Fields(sName)
2038           Else
2039               sParam = vbNullString
2040           End If
2041           txtNPSHr.Text = sParam

2042           sName = "LiquidTemperature"
2043           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2044               sParam = rsPumpData.Fields(sName)
2045           Else
2046               sParam = vbNullString
2047           End If
2048           txtLiquidTemperature.Text = sParam

2049           sName = "RatedInputPower"
2050           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2051               sParam = rsPumpData.Fields(sName)
2052           Else
2053               sParam = vbNullString
2054           End If
2055           txtRatedInputPower.Text = sParam

2056           sName = "FLCurrent"
2057           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2058               sParam = rsPumpData.Fields(sName)
2059           Else
2060               sParam = vbNullString
2061           End If
2062           txtAmps.Text = sParam

2063           sName = "TEMCFrameNumber"
2064           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2065               sParam = rsPumpData.Fields(sName)
2066           Else
2067               sParam = vbNullString
2068           End If
2069           txtTEMCFrameNumber.Text = sParam

2070           optMfr(0).value = rsPumpData.Fields("ChempumpPump")
2071           optMfr(1).value = Not optMfr(0).value

2072           If rsPumpData.Fields("Field1") = "Feathered" Then
2073               Me.chkSuperMarketFeathered.value = Checked
2074           Else
2075               Me.chkSuperMarketFeathered.value = Unchecked
2076           End If

               'select the testsetup data
2077           qyTestSetup.ActiveConnection = cnPumpData
2078           qyTestSetup.CommandText = "SELECT * FROM TempTestSetupData WHERE (((TempTestSetupData.SerialNumber)='" & _
                                    txtSN.Text & "')) ORDER BY Date"
       '        qyTestSetup.CommandText = "SELECT * FROM TempTestSetupData WHERE (((TempTestSetupData.SerialNumber)='" & _
'               txtSN.Text & "'))"

2079           With rsTestSetup
2080               If .State = adStateOpen Then
2081                   .Close
2082               End If
2083               .CursorLocation = adUseClient
2084               .CursorType = adOpenStatic
2085               .Index = "FindData"
2086               .Open qyTestSetup
2087           End With


               'add the selection of dates to the Test Date combo box
2088           If rsTestSetup.RecordCount <> 0 Then
2089               For I = 0 To cmbTestDate.ListCount - 1
2090                   cmbTestDate.RemoveItem 0
2091               Next I
2092               rsTestSetup.MoveFirst
2093               For I = 1 To rsTestSetup.RecordCount
2094                   cmbTestDate.AddItem rsTestSetup.Fields("Date")
2095                   rsTestSetup.MoveNext
2096               Next I
2097               rsTestSetup.MoveFirst
2098               boFoundTestSetup = True

2099               If rsTestSetup.Fields("Approved") = True Then
2100                   DisableTestSetupDataControls                         'if it's in the real database, don't allow changes here
2101                   boTestDateIsApproved = True
2102                   lblTestDateApproved.Visible = True
2103                   If boCanApprove Then
2104                       cmdApproveTestDate.Caption = "Unapprove this Test Date"
2105                   End If
2106               Else
2107                   EnableTestSetupDataControls                          'it's in the temp database, allow changes
2108                   lblTestDateApproved.Visible = False
2109                   If boCanApprove Then
2110                       cmdApproveTestDate.Caption = "Approve this Test Date"
2111                   End If
2112               End If
2113               cmbTestDate.ListIndex = 0
2114           Else
2115               MsgBox ("There is no Test Setup Data for Serial Number " & txtSN.Text)
2116               boFoundTestSetup = False        'didn't find any data
2117               boFoundTestData = False
2118               cmbTestDate.AddItem Date        'load with today
2119               cmbTestDate.ListIndex = 0       'show the entry
2120               EnableTestSetupDataControls
2121               txtTestRemarks.Text = ""
2122               txtVibAx.Text = ""
2123               txtVibRad.Text = ""
2124               txtThrustBal.Text = ""
2125               txtTEMCTRGReading.Text = ""
2126               txtTEMCFrontThrust.Text = ""
2127               txtTEMCRearThrust.Text = ""
' <VB WATCH>
2128       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2129               Exit Sub
2130           End If

2131           If cmbTestDate.ListCount = 1 Then       'if there's only one test date, select it
2132           End If
' <VB WATCH>
2133       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2134           Exit Sub
2135       End If


2136       Do While boUsingEpicor = True   'need a do loop to exit
2137           If boUsingEpicor = True Then
                   'Dim MyRecord As SNRecord
2138               Dim MyRecord As SNRecord
           '            I = InStr(1, txtSN.Text, "-")
           '            If I > 0 Then
2139                   MyRecord = GetEpicorODBCData(txtSN.Text, EpicorConnectionString)
           '            End If
2140               If MyRecord.SONumber = "" Then
2141                   MsgBox ("Not found in Epicor")
2142                   boUsingEpicor = False
2143                   boEpicorFound = False
2144                   Exit Do
2145               End If

2146               If MyRecord.SONumber = 0 Then
2147                   boEpicorFound = False
2148                   boUsingSupermarketTable = True
2149                   boUsingEpicor = False
2150               Else
2151                   boEpicorFound = True
2152                   boUsingSupermarketTable = False
2153               End If

2154               If boEpicorFound = True Then
2155                   boUsingEpicor = False
       '                boEpicorFound = True
2156                   txtSalesOrderNumber.Text = MyRecord.SONumber
2157                   txtLineNumber.Text = MyRecord.SOLine
2158                   txtBilNo.Text = MyRecord.Customer
2159                   txtXPartNum.Text = MyRecord.XPartNum
2160                   txtCustPONum.Text = MyRecord.CustomerPO

2161                   If MyRecord.ShipTo = "" Then
2162                       txtShpNo.Text = MyRecord.Customer
2163                   Else
2164                       txtShpNo.Text = MyRecord.ShipTo
2165                   End If
2166                   txtModelNo.Text = MyRecord.PartNum
2167                   txtModelNo_Change
2168                   txtDesignTDH.Text = MyRecord.TDH
2169                   txtSpGr.Text = MyRecord.SpGr
2170                   txtImpellerDia.Text = MyRecord.ImpellerDiameter
2171                   txtDesignFlow.Text = MyRecord.Flow
2172                   txtNoPhases.Text = MyRecord.Phases
2173                   txtNPSHr.Text = MyRecord.NPSHr
2174                   txtRatedInputPower.Text = MyRecord.RatedInputPower
2175                   txtAmps.Text = MyRecord.FLCurrent
2176                   txtThermalClass.Text = MyRecord.ThermalClass
2177                   txtViscosity.Text = MyRecord.Viscosity
2178                   txtTEMCViscosity.Text = Format((Val(MyRecord.Viscosity) / Val(MyRecord.SpGr)), "000.00")
2179                   txtExpClass.Text = MyRecord.ExpClass
2180                   txtLiquidTemperature.Text = MyRecord.LiquidTemp
2181                   txtLiquid.Text = MyRecord.Fluid
2182                   txtJobNum.Text = MyRecord.JobNumber

2183                   For I = 0 To cmbStatorFill.ListCount - 1
2184                       If InStr(1, UCase$(MyRecord.StatorFill), UCase$(cmbStatorFill.List(I))) <> 0 Then
2185                           cmbStatorFill.ListIndex = I
2186                           Exit For
2187                       End If
2188                   Next I

2189                   For I = 0 To cmbCirculationPath.ListCount - 1
2190                       If InStr(1, UCase$(MyRecord.CirculationPath), UCase$(cmbCirculationPath.List(I))) <> 0 Then
2191                           cmbCirculationPath.ListIndex = I
2192                           Exit For
2193                       End If
2194                   Next I

2195                   For I = 0 To cmbDesignPressure.ListCount - 1
2196                       If InStr(1, MyRecord.DesignPressure, cmbDesignPressure.List(I)) <> 0 Then
2197                           cmbDesignPressure.ListIndex = I
2198                           Exit For
2199                       End If
2200                   Next I

2201                   For I = 0 To cmbVoltage.ListCount - 1
2202                       If InStr(1, MyRecord.Voltage, cmbVoltage.List(I)) <> 0 Then
2203                           cmbVoltage.ListIndex = I
2204                           Exit For
2205                       End If
2206                   Next I

2207                   For I = 0 To cmbFrequency.ListCount - 1
2208                       If InStr(1, MyRecord.Frequency, sName) <> 0 Then
2209                           cmbFrequency.ListIndex = I
2210                           Exit For
2211                       End If
2212                   Next I

2213                   For I = 0 To cmbRPM.ListCount - 1
2214                       If InStr(1, MyRecord.RPM, cmbRPM.List(I)) <> 0 Then
2215                           cmbRPM.ListIndex = I
2216                           Exit For
2217                       End If
2218                   Next I

2219                   For I = 0 To cmbSuctDia.ListCount - 1
2220                       If InStr(1, MyRecord.SuctFlangeSize, cmbSuctDia.List(I)) <> 0 Then
2221                           cmbSuctDia.ListIndex = I
2222                           Exit For
2223                       End If
2224                   Next I

2225                   For I = 0 To cmbDischDia.ListCount - 1
2226                       If InStr(1, MyRecord.DischFlangeSize, cmbDischDia.List(I)) <> 0 Then
2227                           cmbDischDia.ListIndex = I
2228                           Exit For
2229                       End If
2230                   Next I

2231                   For I = 0 To cmbTestSpec.ListCount - 1
2232                       If InStr(1, MyRecord.TestProcedure, cmbTestSpec.List(I)) <> 0 Then
2233                           cmbTestSpec.ListIndex = I
2234                           Exit For
2235                       End If
2236                   Next I

2237                   For I = 0 To cmbMotor.ListCount - 1
2238                       If InStr(1, MyRecord.MotorSize, cmbMotor.List(I)) <> 0 Then
2239                           cmbMotor.ListIndex = I
2240                           Exit For
2241                       End If
2242                   Next I


2243               End If
2244           End If
2245       Loop

2246       If boUsingSupermarketTable = True Then
2247           GetSuperMarketPump MyRecord.PartNum, MyRecord.JobNumber
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
2248       End If
' <VB WATCH>
2249       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2250       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdFindPump_Click"

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
            vbwReportVariable "sAns", sAns
            vbwReportVariable "sSO", sSO
            vbwReportVariable "sParam", sParam
            vbwReportVariable "sName", sName
            vbwReportVariable "I", I
            vbwReportVariable "ok", ok
            vbwReport_EpicorRoutines_SNRecord "MyRecord", MyRecord
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmdModifyBalanceHoleData_Click()
' <VB WATCH>
2251       On Error GoTo vbwErrHandler
2252       Const VBWPROCNAME = "frmPLCData.cmdModifyBalanceHoleData_Click"
2253       If vbwProtector.vbwTraceProc Then
2254           Dim vbwProtectorParameterString As String
2255           If vbwProtector.vbwTraceParameters Then
2256               vbwProtectorParameterString = "()"
2257           End If
2258           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2259       End If
' </VB WATCH>
2260       Dim strInput As String
2261       Dim I As Integer
2262       Dim sNumber As Integer
2263       Dim sDia As String
2264       Dim sBC As String

2265       cmdModifyBalanceHoleData.Visible = False

2266       If dgBalanceHoles.SelBookmarks.Count = 0 Then
2267           cmdModifyBalanceHoleData.Visible = False
' <VB WATCH>
2268       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2269           Exit Sub
2270       End If

2271       rsBalanceHoles.MoveFirst
2272       rsBalanceHoles.Move dgBalanceHoles.SelBookmarks(0) - dgBalanceHoles.FirstRow

2273       sNumber = rsBalanceHoles!Number
2274       If rsBalanceHoles!diameter = 99 Then
2275           sDia = "Slot"
2276       Else
2277           sDia = str(rsBalanceHoles!diameter)
2278       End If
2279       If rsBalanceHoles!boltcircle = 99 Then
2280           sBC = "Unknown"
2281       Else
2282           sBC = str(rsBalanceHoles!boltcircle)
2283       End If


           'get the data for the balance holes
2284       strInput = InputBox("Enter Number of Holes (0 to delete entry)", , sNumber)
2285       If strInput = "" Then
2286           GoTo DeleteIt
2287       End If
2288       sNumber = CInt(strInput)
2289       If Val(sNumber) = 0 Then
2290           GoTo DeleteIt
2291       End If

2292       strInput = InputBox("Enter Decimal Value of Hole Diameter or 'Slot' (For Example, 0.675) ", , sDia)
2293       If strInput <> "" Then
2294           If UCase(strInput) = "SLOT" Then
2295               strInput = 99
2296           End If
2297           sDia = CSng(strInput)
2298       Else
2299           GoTo CancelPressed
2300       End If

2301       strInput = InputBox("Enter Decimal Value of Bolt Circle or 'Unknown' (For Example, 4.525)", , sBC)
2302       If strInput <> "" Then
2303           If UCase(strInput) = "UNKNOWN" Then
2304               strInput = 99
2305           End If
2306           sBC = CSng(strInput)
2307       Else
2308           GoTo CancelPressed
2309       End If

2310       rsBalanceHoles!Number = sNumber
2311       rsBalanceHoles!diameter = sDia
2312       rsBalanceHoles!boltcircle = sBC

2313       rsBalanceHoles.Update
           'rsBalanceHoles.Filter = "SerialNo = '" & frmPLCData.txtSN.Text & "'"

2314       GetBalanceHoleData txtSN.Text, cmbTestDate.Text
       '    rsBalanceHoles.Requery
2315       rsBalanceHoles.MoveLast
2316       dgBalanceHoles.Refresh
2317       chkBalanceHoles.value = 1
2318       rsBalanceHoles.MoveFirst

' <VB WATCH>
2319       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2320       Exit Sub

2321   CancelPressed:
2322       MsgBox "No New Balance Hole Data Entered", vbOKOnly

2323   DeleteIt:
2324       If (MsgBox("Do you really want to delete this entry?", vbYesNo, "Deleting Balance Hole Data. . .")) = vbYes Then
2325           rsBalanceHoles.Delete
2326           rsBalanceHoles.Update
2327           GetBalanceHoleData txtSN.Text, cmbTestDate.Text
       '        rsBalanceHoles.Requery
2328           If Not rsBalanceHoles.EOF Then
2329               rsBalanceHoles.MoveLast
2330           End If
2331           dgBalanceHoles.Refresh
2332           chkBalanceHoles.value = 1
2333           If Not rsBalanceHoles.BOF Then
2334               rsBalanceHoles.MoveFirst
2335           End If
2336       End If


' <VB WATCH>
2337       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2338       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdModifyBalanceHoleData_Click"

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
            vbwReportVariable "strInput", strInput
            vbwReportVariable "I", I
            vbwReportVariable "sNumber", sNumber
            vbwReportVariable "sDia", sDia
            vbwReportVariable "sBC", sBC
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmdReport_Click()
           'view/print a report
' <VB WATCH>
2339       On Error GoTo vbwErrHandler
2340       Const VBWPROCNAME = "frmPLCData.cmdReport_Click"
2341       If vbwProtector.vbwTraceProc Then
2342           Dim vbwProtectorParameterString As String
2343           If vbwProtector.vbwTraceParameters Then
2344               vbwProtectorParameterString = "()"
2345           End If
2346           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2347       End If
' </VB WATCH>
2348       Dim I As Integer

2349       ExportToExcel

' <VB WATCH>
2350       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2351       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdReport_Click"

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
            vbwReportVariable "I", I
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub cmdSearchForPump_Click()
' <VB WATCH>
2352       On Error GoTo vbwErrHandler
2353       Const VBWPROCNAME = "frmPLCData.cmdSearchForPump_Click"
2354       If vbwProtector.vbwTraceProc Then
2355           Dim vbwProtectorParameterString As String
2356           If vbwProtector.vbwTraceParameters Then
2357               vbwProtectorParameterString = "()"
2358           End If
2359           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2360       End If
' </VB WATCH>
2361       LoadCombo frmSearch.cmbSearchModel, "TEMCHydraulics"

2362       frmSearch.Show
' <VB WATCH>
2363       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2364       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdSearchForPump_Click"

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

Private Sub cmdSelectSupermarket_Click()
' <VB WATCH>
2365       On Error GoTo vbwErrHandler
2366       Const VBWPROCNAME = "frmPLCData.cmdSelectSupermarket_Click"
2367       If vbwProtector.vbwTraceProc Then
2368           Dim vbwProtectorParameterString As String
2369           If vbwProtector.vbwTraceParameters Then
2370               vbwProtectorParameterString = "()"
2371           End If
2372           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2373       End If
' </VB WATCH>
2374       grpSupermarket.Visible = False
' <VB WATCH>
2375       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2376       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdSelectSupermarket_Click"

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

Private Sub cmdWriteSP_Click()
           'write the sp to the plc
' <VB WATCH>
2377       On Error GoTo vbwErrHandler
2378       Const VBWPROCNAME = "frmPLCData.cmdWriteSP_Click"
2379       If vbwProtector.vbwTraceProc Then
2380           Dim vbwProtectorParameterString As String
2381           If vbwProtector.vbwTraceParameters Then
2382               vbwProtectorParameterString = "()"
2383           End If
2384           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2385       End If
' </VB WATCH>
2386       Dim rc As String
2387       Dim S As String

           'write the set point data to the PLC
2388           bWrite = True
2389           S = Right$("0000" & txtWriteSPData, 4)
2390           S = Right$(S, 2) & Left$(S, 2)
2391           rc = StringToByteArray(S, ByteBuffer)

2392           DataLength = HexConvert(ByteBuffer, 2)
2393           DataAddress = StringToHexInt("2005")

2394           rc = GetData

2395           bWrite = False
' <VB WATCH>
2396       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2397       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "cmdWriteSP_Click"

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
            vbwReportVariable "S", S
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

'Private Sub Command1_Click()
'    Dim frmem As New InteropDBWithButtons.Form1
'    frmem.ConString = cnPumpData.ConnectionString
'    frmem.Caption = "Email Database Maintenance"
'    frmem.Show 1
'End Sub

Private Sub btnRunNPSH_Click()
' <VB WATCH>
2398       On Error GoTo vbwErrHandler
2399       Const VBWPROCNAME = "frmPLCData.btnRunNPSH_Click"
2400       If vbwProtector.vbwTraceProc Then
2401           Dim vbwProtectorParameterString As String
2402           If vbwProtector.vbwTraceParameters Then
2403               vbwProtectorParameterString = "()"
2404           End If
2405           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2406       End If
' </VB WATCH>
2407       Static OriginalColor As Long
2408       If btnRunNPSH.Caption = "Run NPSH" Then
2409           btnRunNPSH.Caption = "Cancel NPSH Run"
2410           OriginalColor = btnRunNPSH.BackColor
2411           tmrNPSHr.Enabled = False
2412           btnRunNPSH.BackColor = vbRed
2413           If boCanApprove Then
2414               txtNPSH(5).Visible = True
2415               lbltab4(5).Visible = True
2416           Else
2417               txtNPSH(5).Visible = False
2418               lbltab4(5).Visible = False
2419           End If
2420           WroteNPSHr = False

2421           frmNPSH.Visible = True
2422           txtNPSH(5).Enabled = True
2423           If Val(txtTDH.Text) <= 10 Then
2424               MsgBox "This test will not work starting with this starting TDH.  Ending test...", vbOKOnly, "Flow is 0"
2425               btnRunNPSH.Caption = "Run NPSH"
2426               btnRunNPSH.BackColor = OriginalColor
2427               frmNPSH.Visible = False
' <VB WATCH>
2428       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2429               Exit Sub
2430           End If
               'load initial values
2431           If DataGrid2.Row = -1 Then
2432               MsgBox "You must write the normal test data to this row before you run NPSH.", vbOKOnly, "Nothing written for this row"
2433               btnRunNPSH.Caption = "Run NPSH"
2434               btnRunNPSH.BackColor = OriginalColor
2435               frmNPSH.Visible = False
' <VB WATCH>
2436       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2437               Exit Sub
2438           Else
2439               DataGrid2.Row = UpDown1.value - 1
2440           End If

2441           txtNPSH(0).Text = DataGrid2.Columns("Flow")
2442           txtNPSH(3).Text = DataGrid2.Columns("TDH")
2443           txtNPSH(4) = 0
               'txtNPSH(0).Text = txtFlow.Text
               'txtNPSH(3).Text = txtTDH.Text
2444           txtNPSH(4) = 0
2445       Else
2446           btnRunNPSH.Caption = "Run NPSH"
2447           btnRunNPSH.BackColor = OriginalColor
2448           frmNPSH.Visible = False
2449       End If

           'ReportToExcel
' <VB WATCH>
2450       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2451       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "btnRunNPSH_Click"

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
            vbwReportVariable "OriginalColor", OriginalColor
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

    Private Sub updown1_change()
' <VB WATCH>
2452       On Error GoTo vbwErrHandler
2453       Const VBWPROCNAME = "frmPLCData.updown1_change"
2454       If vbwProtector.vbwTraceProc Then
2455           Dim vbwProtectorParameterString As String
2456           If vbwProtector.vbwTraceParameters Then
2457               vbwProtectorParameterString = "()"
2458           End If
2459           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2460       End If
' </VB WATCH>
2461       Dim sName As String

2462       If Not rsTestData.BOF Then
2463           rsTestData.MoveFirst
2464       End If

2465       If Not rsTestData.BOF Or Not rsTestData.EOF Then
2466           rsTestData.Move UpDown1.value - 1
2467       End If

2468       sName = "VibrationX"
2469       If rsTestData.Fields(sName).ActualSize <> 0 Then
2470           txtVibAx.Text = rsTestData.Fields(sName)
2471       Else
       '        txtVibAx.Text = vbNullString
2472       End If

2473       sName = "VibrationY"
2474       If rsTestData.Fields(sName).ActualSize <> 0 Then
2475           txtVibRad.Text = rsTestData.Fields(sName)
2476       Else
       '        txtVibRad.Text = vbNullString
2477       End If

2478       sName = "Remarks"
2479       If rsTestData.Fields(sName).ActualSize <> 0 Then
2480           txtTestRemarks.Text = rsTestData.Fields(sName)
2481       Else
       '        txtTestRemarks.Text = vbNullString
2482       End If

2483       sName = "ThrustBalance"
2484       If rsTestData.Fields(sName).ActualSize <> 0 Then
2485           txtThrustBal.Text = rsTestData.Fields(sName)
2486       Else
       '        txtThrustBal.Text = vbNullString
2487       End If

2488       sName = "TEMCTRG"
2489       If rsTestData.Fields(sName).ActualSize <> 0 Then
2490           txtTEMCTRGReading.Text = rsTestData.Fields(sName)
2491       Else
2492           txtTEMCTRGReading.Text = 0
       '        txtTEMCTRGReading.Text = vbNullString
2493       End If

2494       sName = "TEMCFrontThrust"
2495       If rsTestData.Fields(sName).ActualSize <> 0 Then
2496           txtTEMCFrontThrust.Text = rsTestData.Fields(sName)
2497       Else
       '        txtTEMCFrontThrust.Text = vbNullString
2498       End If

2499       sName = "TEMCRearThrust"
2500       If rsTestData.Fields(sName).ActualSize <> 0 Then
2501           txtTEMCRearThrust.Text = rsTestData.Fields(sName)
2502       Else
       '        txtTEMCRearThrust.Text = vbNullString
2503       End If
2504       sName = "TEMCMomentArm"
2505       If rsTestData.Fields(sName).ActualSize <> 0 Then
2506           txtTEMCMomentArm.Text = rsTestData.Fields(sName)
2507       Else
       '        txtTEMCMomentArm.Text = vbNullString
2508       End If
2509       sName = "TEMCThrustRigPressure"
2510       If rsTestData.Fields(sName).ActualSize <> 0 Then
2511           txtTEMCThrustRigPressure.Text = rsTestData.Fields(sName)
2512       Else
       '        txtTEMCThrustRigPressure.Text = vbNullString
2513       End If
2514       sName = "TEMCViscosity"
2515       If rsTestData.Fields(sName).ActualSize <> 0 And rsTestData.Fields(sName) <> 0 Then
2516           txtTEMCViscosity.Text = rsTestData.Fields(sName)
2517       Else
       '        txtTEMCViscosity.Text = vbNullString
2518       End If

2519       CalculateTEMCForce

2520       rsEff.MoveFirst
2521       rsEff.Move UpDown1.value - 1
' <VB WATCH>
2522       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2523       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "updown1_change"

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
            vbwReportVariable "sName", sName
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Sub CalculateTEMCForce()
' <VB WATCH>
2524       On Error GoTo vbwErrHandler
2525       Const VBWPROCNAME = "frmPLCData.CalculateTEMCForce"
2526       If vbwProtector.vbwTraceProc Then
2527           Dim vbwProtectorParameterString As String
2528           If vbwProtector.vbwTraceParameters Then
2529               vbwProtectorParameterString = "()"
2530           End If
2531           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2532       End If
' </VB WATCH>
2533       Dim NoOfPoles As Integer
2534       Dim Frequency As Integer
2535       Dim Additions As String
2536       Dim Frame As String
2537       Dim VOverA As Double
2538       Dim Force As Double
2539       Dim Gravity As Double

2540       If Val(txtSpGr.Text) = 0 Then
2541           Gravity = 1
2542       Else
2543           Gravity = CDbl(Val(txtSpGr.Text))
2544       End If

           'show calculated values
2545       If Val(txtTEMCFrontThrust.Text) = 0 Then
2546           If Val(txtTEMCRearThrust.Text) = 0 Then
               'no thrust entered
2547               lblTEMCFrontRear.Visible = False
2548               txtTEMCCalcForce.Text = " "
2549           Else
                   'rear thrust
2550               txtTEMCCalcForce.Text = Format(Gravity * (Val(txtTEMCRearThrust.Text) * Val(txtTEMCMomentArm.Text) - (Val(txtTEMCThrustRigPressure.Text) / 14.223) * 4.5), "##0.0")
2551               lblTEMCFrontRear.Caption = "REAR"
2552               lblTEMCFrontRear.Visible = True
2553           End If
2554       Else
               'front thrust
2555           txtTEMCCalcForce.Text = Format(Gravity * (Val(txtTEMCFrontThrust.Text) * Val(txtTEMCMomentArm.Text) + (Val(txtTEMCThrustRigPressure.Text) / 14.223) * 4.5), "##0.0")
2556           lblTEMCFrontRear.Caption = "FRONT"
2557           lblTEMCFrontRear.Visible = True
2558       End If

2559       If Val(txtTEMCCalcForce.Text) < 0 Then
2560           txtTEMCCalcForce.Text = -txtTEMCCalcForce
2561           lblTEMCFrontRear.Caption = "FRONT"
2562       End If

           'see how many poles we have, it's the next to last number in the frame size
2563       If Len(txtTEMCFrameNumber) > 2 Then
2564           NoOfPoles = 2 * Val(Left$(Right$(txtTEMCFrameNumber.Text, 2), 1))
2565       End If

2566       If cmbTEMCAdditions.ListIndex <> -1 Then
2567           Additions = Mid$(cmbTEMCAdditions.List(cmbTEMCAdditions.ListIndex), 2, 1)
2568           If Additions = "A" Or Additions = "E" Or Additions = "G" Or Additions = "J" Then
2569               Frequency = 60
2570           ElseIf Additions = "B" Or Additions = "F" Or Additions = "H" Or Additions = "K" Then
2571               Frequency = 50
2572           Else
2573               Frequency = 0
2574           End If
2575       End If

2576       If Len(txtTEMCFrameNumber.Text) = 3 Then
2577           If txtTEMCFrameNumber.Text = "529" Then
2578               Frame = "420"
2579           Else
2580               Frame = Left$(txtTEMCFrameNumber, 2) & "0"
2581           End If
2582       Else
2583           Frame = txtTEMCFrameNumber.Text
2584           If Right$(txtTEMCFrameNumber.Text, 1) = "5" Then
2585               Frame = Frame & Left$(lblTEMCFrontRear.Caption, 1)
2586           Else
2587           End If
2588       End If
2589       Force = DLookupA(3, TEMCForceViscosity, 1, Frame)
2590       If Frequency = 60 Then
2591           Force = Force / 1.2
2592       End If
2593       If Val(txtTEMCViscosity.Text) > 1# Then
2594           If (Val(txtTEMCCalcForce.Text) > 3 * Force) Then
2595               lblTEMCPassFail.Visible = True
2596               lblTEMCPassFail.ForeColor = vbRed
2597               lblTEMCPassFail.Caption = "FAIL"
2598           Else
2599               lblTEMCPassFail.Visible = True
2600               lblTEMCPassFail.ForeColor = vbGreen
2601               lblTEMCPassFail.Caption = "PASS"
2602           End If
2603       End If

2604       If (Val(txtTEMCViscosity.Text) > 0.5) And (Val(txtTEMCViscosity.Text) <= 1#) Then
2605           If (Val(txtTEMCCalcForce.Text) > 2 * Force) Then
2606               lblTEMCPassFail.Visible = True
2607               lblTEMCPassFail.ForeColor = vbRed
2608               lblTEMCPassFail.Caption = "FAIL"
2609           Else
2610               lblTEMCPassFail.Visible = True
2611               lblTEMCPassFail.ForeColor = vbGreen
2612               lblTEMCPassFail.Caption = "PASS"
2613           End If
2614       End If

2615       If (Val(txtTEMCViscosity.Text) > 0.3) And (Val(txtTEMCViscosity.Text) <= 0.5) Then
2616           If (Val(txtTEMCCalcForce.Text) > 1.5 * Force) Then
2617               lblTEMCPassFail.Visible = True
2618               lblTEMCPassFail.ForeColor = vbRed
2619               lblTEMCPassFail.Caption = "FAIL"
2620           Else
2621               lblTEMCPassFail.Visible = True
2622               lblTEMCPassFail.ForeColor = vbGreen
2623               lblTEMCPassFail.Caption = "PASS"
2624           End If
2625       End If

2626       If (Val(txtTEMCViscosity.Text) <= 0.3) Then
2627           If (Val(txtTEMCCalcForce.Text) > 1# * Force) Then
2628               lblTEMCPassFail.Visible = True
2629               lblTEMCPassFail.ForeColor = vbRed
2630               lblTEMCPassFail.Caption = "FAIL"
2631           Else
2632               lblTEMCPassFail.Visible = True
2633               lblTEMCPassFail.ForeColor = vbGreen
2634               lblTEMCPassFail.Caption = "PASS"
2635           End If
2636       End If
2637       If NoOfPoles <> 0 Then
2638           VOverA = (DLookupA(2, TEMCForceViscosity, 1, Frame)) / (NoOfPoles * 30 / Frequency)
2639       End If
       '    If Frequency = 60 Then
       '        VOverA = VOverA * 1.2
       '    End If

2640       txtTEMCPVValue.Text = Format(Val(txtTEMCCalcForce.Text) * VOverA, "##0.0")

2641       If Val(txtTEMCFrontThrust.Text) = 0 And Val(txtTEMCRearThrust.Text) = 0 Then
2642           txtTEMCPVValue.Text = ""
2643           txtTEMCCalcForce.Text = ""
2644           lblTEMCPassFail.Visible = False
2645       End If


           'calculate reverse head
2646       txtRevHead.Text = Format(rsTestData.Fields("RBHPress") - rsTestData.Fields("SuctionPressure") * 2.31, "##0.0")
       '    txtRevHead.Text = Format((CDbl(Val(txtAI3Display.Text)) - CDbl(Val(txtSuctionDisplay.Text))) * 2.31, "##0.0")

' <VB WATCH>
2647       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2648       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "CalculateTEMCForce"

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
            vbwReportVariable "NoOfPoles", NoOfPoles
            vbwReportVariable "Frequency", Frequency
            vbwReportVariable "Additions", Additions
            vbwReportVariable "Frame", Frame
            vbwReportVariable "VOverA", VOverA
            vbwReportVariable "Force", Force
            vbwReportVariable "Gravity", Gravity
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
    Private Sub updown2_change()
' <VB WATCH>
2649       On Error GoTo vbwErrHandler
2650       Const VBWPROCNAME = "frmPLCData.updown2_change"
2651       If vbwProtector.vbwTraceProc Then
2652           Dim vbwProtectorParameterString As String
2653           If vbwProtector.vbwTraceParameters Then
2654               vbwProtectorParameterString = "()"
2655           End If
2656           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2657       End If
' </VB WATCH>
2658       Dim Plothead(1, 7) As Single
2659       Dim HeadPlot(7, 1) As Single

2660       Dim PlotEff() As Single
2661       Dim PlotKW() As Single
2662       Dim PlotAmps() As Single

2663       Dim j As Integer

2664       For j = 0 To UpDown2.value - 1
2665           Plothead(0, j) = HeadFlow(0, j)
2666           Plothead(1, j) = HeadFlow(1, j)
2667           HeadPlot(j, 0) = FlowHead(j, 0)
2668           HeadPlot(j, 1) = FlowHead(j, 1)
       '        ReDim Preserve PlotEff(1, j)
       '        PlotEff(0, j) = EffFlow(0, j)
       '        PlotEff(1, j) = EffFlow(1, j)
       '        ReDim Preserve PlotKW(1, j)
       '        PlotKW(0, j) = KWFlow(0, j)
       '        PlotKW(1, j) = KWFlow(1, j)
       '        ReDim Preserve PlotAmps(1, j)
       '        PlotAmps(0, j) = AmpsFlow(0, j)
       '        PlotAmps(1, j) = AmpsFlow(1, j)
2669       Next j

2670       MSChart1 = HeadPlot

' <VB WATCH>
2671       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2672       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "updown2_change"

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
            vbwReportVariable "Plothead", Plothead
            vbwReportVariable "HeadPlot", HeadPlot
            vbwReportVariable "PlotEff", PlotEff
            vbwReportVariable "PlotKW", PlotKW
            vbwReportVariable "PlotAmps", PlotAmps
            vbwReportVariable "j", j
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
' <VB WATCH>
2673       On Error GoTo vbwErrHandler
2674       Const VBWPROCNAME = "frmPLCData.DataGrid1_AfterColUpdate"
2675       If vbwProtector.vbwTraceProc Then
2676           Dim vbwProtectorParameterString As String
2677           If vbwProtector.vbwTraceParameters Then
2678               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ColIndex", ColIndex) & ") "
2679           End If
2680           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2681       End If
' </VB WATCH>
2682       DoEfficiencyCalcs
' <VB WATCH>
2683       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2684       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DataGrid1_AfterColUpdate"

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
            vbwReportVariable "ColIndex", ColIndex
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub dgBalanceHoles_SelChange(Cancel As Integer)
' <VB WATCH>
2685       On Error GoTo vbwErrHandler
2686       Const VBWPROCNAME = "frmPLCData.dgBalanceHoles_SelChange"
2687       If vbwProtector.vbwTraceProc Then
2688           Dim vbwProtectorParameterString As String
2689           If vbwProtector.vbwTraceParameters Then
2690               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Cancel", Cancel) & ") "
2691           End If
2692           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2693       End If
' </VB WATCH>
2694       If dgBalanceHoles.SelBookmarks.Count = 0 Then
2695           cmdModifyBalanceHoleData.Visible = False
2696       Else
2697           cmdModifyBalanceHoleData.Visible = True
2698       End If
' <VB WATCH>
2699       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2700       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "dgBalanceHoles_SelChange"

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
            vbwReportVariable "Cancel", Cancel
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub Form_Activate()
' <VB WATCH>
2701       On Error GoTo vbwErrHandler
2702       Const VBWPROCNAME = "frmPLCData.Form_Activate"
2703       If vbwProtector.vbwTraceProc Then
2704           Dim vbwProtectorParameterString As String
2705           If vbwProtector.vbwTraceParameters Then
2706               vbwProtectorParameterString = "()"
2707           End If
2708           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2709       End If
' </VB WATCH>
2710       If ProgramEnd = True Then
2711           Unload Me
2712       End If
' <VB WATCH>
2713       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2714       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Form_Activate"

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

Private Sub Form_Load()
' <VB WATCH>
2715       On Error GoTo vbwErrHandler
2716       Const VBWPROCNAME = "frmPLCData.Form_Load"
2717       If vbwProtector.vbwTraceProc Then
2718           Dim vbwProtectorParameterString As String
2719           If vbwProtector.vbwTraceParameters Then
2720               vbwProtectorParameterString = "()"
2721           End If
2722           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2723       End If
' </VB WATCH>
2724       Dim RetVal As String
2725       Dim sSendStr As String
2726       Dim I As Integer
2727       Dim j As Integer
2728       Dim sTableName As String
2729       Dim WhichServer As String
2730       Dim WhichDatabase As String

2731       ProgramEnd = False
2732       Dim objWMIService As Object
2733       Dim colProcesses As Object
2734       Set objWMIService = GetObject("winmgmts:")
2735       Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name LIKE 'PolarRundown%'")
       '    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name LIKE 'Excel%'")
2736       If colProcesses.Count > 1 Then
2737           MsgBox "There is already a copy of Polar Rundown running.  You can only have one copy running at a time", vbOKOnly, "Polar Rundown already running"
2738           Dim f As Form
2739           For Each f In Forms
2740               If f.Name <> Me.Name Then
2741                    Unload f
2742               End If
2743           Next
2744           ProgramEnd = True
' <VB WATCH>
2745       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2746           Exit Sub
2747       Else
2748       End If
2749       Set objWMIService = Nothing
2750       Set colProcesses = Nothing

2751       debugging = 0   'assume not debugging
2752       WhichServer = "Production"     'change to production server
2753       WhichDatabase = "Production"

2754       If UCase$(Left$(GetMachineName, 5)) = "MROSE" Or UCase$(Left$(GetMachineName, 5)) = "ITTES" Then  'if mickey, see if we want to be in debug
2755           I = MsgBox("Debug?", vbYesNo)
2756           If I = vbYes Then
2757               debugging = 1
2758               WhichServer = "Production"
2759               WhichDatabase = "Production"
2760           Else
2761           End If
2762       End If

2763       If debugging Then
       '        GoTo temp
2764       End If
           'see if the mdb file is where it's supposed to be

2765       Dim developmentDatabase As String
2766       developmentDatabase = GetUNCFromLetter("F:") & sDevelopmentDatabase

2767       If Dir(developmentDatabase) = "" Then
2768           MsgBox "Development.mdb does not exist on F:, Please contact IT.", , "No Development Database"
2769           End
2770       End If

           'get the database info from the new mdb file
2771       Dim cnDevelopment As New ADODB.Connection
2772       Dim qyDevelopment As New ADODB.Command
2773       Dim rsDevelopment As New ADODB.Recordset

2774       On Error GoTo CannotConnect

2775       With cnDevelopment
2776           .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & developmentDatabase & ";Persist Security Info=False; Jet OLEDB:Database Password=Access7277word;"
2777           .ConnectionTimeout = 10
2778           .Open
2779       End With

2780   On Error GoTo vbwErrHandler
2781       GoTo Connected

2782   CannotConnect:
2783       MsgBox "Cannot connect with Development.mdb database.  Please contact IT.", , "Cannot find Connection data."
2784       End

2785   Connected:

           'we're connected, get the data for the Epicor SQL server
2786       qyDevelopment.CommandText = "SELECT * FROM Connections WHERE Connections.WhichServer = '" & WhichServer & "' AND WhichDatabase = '" & WhichDatabase & "'"
2787       qyDevelopment.ActiveConnection = cnDevelopment

2788       rsDevelopment.CursorLocation = adUseClient
2789       rsDevelopment.CursorType = adOpenStatic
2790       rsDevelopment.LockType = adLockOptimistic

2791       On Error GoTo NoServerData

2792       rsDevelopment.Open qyDevelopment

2793   On Error GoTo vbwErrHandler
2794       GoTo GotServerData

2795   NoServerData:

2796       MsgBox "Cannot connect with Development.mdb database.  Please contact IT.", , "Cannot find Connection data."
2797       End

2798   GotServerData:

2799       If rsDevelopment.RecordCount <> 1 Then
2800           GoTo NoServerData
2801       End If

           'construct Epicor connection string
2802       EpicorConnectionString = "Driver={" & rsDevelopment.Fields("ODBCDriver") & "};" & _
                                         "Database=" & rsDevelopment.Fields("DatabaseName") & ";" & _
                                         "Server=" & rsDevelopment.Fields("ServerName") & ";" & _
                                         "UID=" & rsDevelopment.Fields("UserName") & ";" & _
                                         "PWD=" & rsDevelopment.Fields("UserPassword") & ";"


           'make sure we can open the SQL database

2803       On Error GoTo CannotOpenEpicorSQLServer

2804       Dim cnTestEpicor As New ADODB.Connection
2805       cnTestEpicor.ConnectionString = EpicorConnectionString
2806       cnTestEpicor.Open
2807       cnTestEpicor.Close
2808       Set cnTestEpicor = Nothing
2809   On Error GoTo vbwErrHandler

2810       GoTo FoundEpicorSQLServer

2811   CannotOpenEpicorSQLServer:
2812       MsgBox "Cannot connect with the Epicor SQL server specified in Development.mdb.  Please contact IT.", , "Cannot connect with Epicor SQL Server"
2813       End

2814   FoundEpicorSQLServer:
           'get data on rundown database
2815       rsDevelopment.Close
2816       qyDevelopment.CommandText = "SELECT * FROM Connections WHERE Connections.WhichServer = 'PolarRundown'"

2817       On Error GoTo NoRundownDatabase

2818       rsDevelopment.Open qyDevelopment

2819       GoTo FoundRundownDatabase

2820   NoRundownDatabase:
2821       MsgBox "Cannot connect with the Pump Rundown database specified in Development.mdb.  Please contact IT.", , "Cannot connect with Epicor SQL Server"
2822       End

2823   FoundRundownDatabase:
2824       If rsDevelopment.RecordCount <> 1 Then
2825           GoTo NoRundownDatabase
2826           End
2827       End If

2828   temp:

2829       If debugging Then
2830           sDataBaseName = "c:\databases\PolarData.mdb"
2831       Else

2832          sDataBaseName = GetUNCFromLetter("F:") & "\Groups\Shared\databases\PolarData.mdb"

       '        sDataBaseName = rsDevelopment.Fields("ServerName") & rsDevelopment.Fields("DatabaseName")

       '        sDataBaseName = sServerName & "f\groups\shared\databases\PumpData 2k.mdb"
2833       End If

2834       Dim tempFSO As Object
2835       Set tempFSO = CreateObject("Scripting.FileSystemObject")
2836       ParentDirectoryName = tempFSO.getparentfoldername(sDataBaseName)
2837       Set tempFSO = Nothing

           'see if we can open the pump rundown database
2838       On Error GoTo NoRundownDatabase
2839       With cnPumpData
       '        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataBaseName & ";Persist Security Info=False;Jet OLEDB:Database Password=185TitusAve"
2840           .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataBaseName & ";Persist Security Info=False;"
2841           .ConnectionTimeout = 10
2842           .Open
2843       End With
2844   On Error GoTo vbwErrHandler


2845       If debugging = 0 Then
       '        Printer.Orientation = vbPRORLandscape
2846       End If

2847       lblVersion = "Polar Rundown - Version " & App.Major & "." & App.Minor & "." & App.Revision
2848       frmPLCData.Caption = "Polar Rundown"

2849       boFoundPump = False

2850       Me.Show

2851       MSChart1.Plot.Axis(VtChAxisIdX).AxisTitle = "Flow"
2852       MSChart1.Plot.Axis(VtChAxisIdY).AxisTitle = "TDH"
           'MSChart1.Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen = True
           'MSChart1.Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen = True
2853       MSChart1.Plot.UniformAxis = False
2854       MSChart1.Plot.SeriesCollection.Item(1).SeriesMarker.Auto = False
2855       MSChart1.Plot.SeriesCollection.Item(1).Pen.Width = 5
2856       With MSChart1.Plot.SeriesCollection.Item(1).DataPoints.Item(-1).Marker
2857           .Visible = True
2858           .Size = 50
2859           .Style = VtMarkerStyleCircle
2860           .FillColor.Automatic = False
2861           .FillColor.Set 0, 0, 255
2862       End With
2863       MSChart1.Plot.AutoLayout = False
2864       MSChart1.Plot.LocationRect.Max.x = 5600
2865       MSChart1.Plot.LocationRect.Max.y = 2800
2866       MSChart1.Plot.LocationRect.Min.x = 0
2867       MSChart1.Plot.LocationRect.Min.y = 0

           'assure that the timers are off
2868       frmPLCData.tmrGetDDE.Enabled = False

2869       frmPLCData.tmrStartUp.Enabled = False

           'initialize the PLC network
2870       RetVal = NetWorkInitialize()
2871       If RetVal <> 0 Then
2872           MsgBox ("Can't Initialize Network. Exiting...")
2873           End
2874       End If

2875       If debugging = 0 Then
               'load array of plcs
2876           I = 0
2877           Open rsDevelopment.Fields("ServerName") & "PolarPLCAddresses.txt" For Input As 1
2878           While Not EOF(1)
2879               Input #1, Description(I)
2880               For j = 0 To 125
2881                   Input #1, aDevices(I).Address(j)
2882               Next j
2883               Input #1, j
2884               I = I + 1
2885           Wend
2886           Close #1

2887           DeviceCount = I

2888           If Left$(GetMachineName, 2) = "WV" Then  'if in WV, put MWSC first in loop dropdown
2889               Dim k As Integer
2890               For k = 0 To DeviceCount - 1
2891                   If InStr(Description(k), "MWSC") <> 0 Then
2892                       Exit For
2893                   End If
2894               Next k
2895               Description(DeviceCount) = Description(0)
2896               Description(0) = Description(k)
2897               Description(k) = Description(DeviceCount)

2898               aDevices(DeviceCount) = aDevices(0)
2899               aDevices(0) = aDevices(k)
2900               aDevices(k) = aDevices(DeviceCount)

2901           End If

2902           Dim PLCAddress As String
2903           For I = 0 To DeviceCount - 1
2904               PLCAddress = aDevices(I).Address(4) & "." & aDevices(I).Address(5) & "." & aDevices(I).Address(6) & "." & aDevices(I).Address(7)
2905               RetVal = PingSilent(PLCAddress)
2906               If RetVal <> 0 Then
2907                   frmPLCData.cmbPLCLoop.AddItem Description(I)
2908                   frmPLCData.cmbPLCLoop.ItemData(frmPLCData.cmbPLCLoop.NewIndex) = I
2909               End If
2910           Next I
2911       End If

2912       frmPLCData.cmbPLCLoop.AddItem "Add PLC Data Manually"   'enable the controls for manual entry

           'turn on the PLC led

2913       frmPLCData.cmbPLCLoop.ListIndex = 0
2914       frmPLCData.tmrGetDDE.Enabled = True

           'hook up to the various databases

           'copy the template of the database here
           'see if it exists
2915       Dim fdrive As String
2916       fdrive = GetUNCFromLetter("F:")
2917       If Dir(fdrive & "\groups\shared\databases" & sEffDataBaseName) = "" Then
2918           MsgBox "File does not exist at " & fdrive & "\groups\shared\databases" & sEffDataBaseName & ". Please contact IT", vbOKOnly, "Eff.mdb does not exist"
2919       Else
               'Dim FSO As New FileSystemObject
2920           FileCopy fdrive & "\groups\shared\databases" & sEffDataBaseName, App.Path & sEffDataBaseName
2921       End If


2922       With cnEffData
2923           .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & sEffDataBaseName & ";Persist Security Info=False"
2924           .Open
2925       End With

           'open some recordsets
2926       rsPumpData.Index = "SerialNumber"
2927       rsTestSetup.Index = "FindData"
2928       rsTestData.Index = "PrimaryKey"
2929       rsPumpData.Open "TempPumpData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
2930       rsTestSetup.Open "TempTestSetupData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
2931       rsTestData.Filter = "SerialNumber = ''"
2932       rsTestData.CursorLocation = adUseClient
2933       rsTestData.Open "TempTestData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
2934       rsEff.CursorLocation = adUseClient
2935       rsEff.Open "Efficiency", cnEffData, adOpenStatic, adLockOptimistic, adCmdTableDirect
2936       qyBalanceHoles.ActiveConnection = cnPumpData
2937       rsBalanceHoles.CursorLocation = adUseClient
2938       rsBalanceHoles.CursorType = adOpenStatic
2939       rsBalanceHoles.LockType = adLockOptimistic
2940       qyMisc.ActiveConnection = cnPumpData
2941       qyMisc.CommandText = "SELECT MiscParameters.ParameterName, MiscParameters.ParameterValue From MiscParameters WHERE (((MiscParameters.ParameterName)='AllowableTDHVariation'));"
2942       rsMisc.CursorLocation = adUseClient
2943       rsMisc.CursorType = adOpenStatic
2944       rsMisc.LockType = adLockBatchOptimistic
2945       rsMisc.Open qyMisc
2946       txtNPSH(5).Text = rsMisc!ParameterValue

2947       If debugging <> 1 Then
2948           FindMagtrols
2949       Else
2950           cmbMagtrol.AddItem "Add Manually"
2951           cmbMagtrol.ItemData(cmbMagtrol.NewIndex) = 99
2952           cmbMagtrol.ListIndex = 0
2953       End If
2954       optKW(1).value = True
2955       optKW_Click (1)


           'blank out data grid
2956       Set DataGrid1.DataSource = rsTestData

           'load the combo boxes
2957       LoadCombo cmbStatorFill, "StatorFill"
2958       LoadCombo cmbCirculationPath, "CirculationPath"
2959       LoadCombo cmbVoltage, "Voltage"
2960       LoadCombo cmbFrequency, "Frequency"
2961       LoadCombo cmbMotor, "Motor"
2962       LoadCombo cmbDesignPressure, "DesignPressure"
2963       LoadCombo cmbRPM, "RPM"
2964       LoadCombo cmbOrificeNumber, "OrificeNumber"
2965       LoadCombo cmbTestSpec, "TestSpecification"
2966       LoadCombo cmbLoopNumber, "LoopNumber"
2967       LoadCombo cmbSuctDia, "SuctionDiameter"
2968       LoadCombo cmbDischDia, "DischargeDiameter"
2969       LoadCombo cmbTachID, "TachID"
2970       LoadCombo cmbAnalyzerNo, "AnalyzerNo"
2971       LoadCombo cmbModel, "Model"
2972       LoadCombo cmbModelGroup, "ModelGroup"
2973       LoadCombo cmbMounting, "Mounting"
2974       LoadCombo cmbPLCNo, "PLCNo"
2975       LoadCombo cmbFlowMeter, "PumpFlowMeter"
2976       LoadCombo cmbSuctionPressureTransducer, "SuctionPressureTransducer"
2977       LoadCombo cmbDischargePressureTransducer, "DischargePressureTransducer"
2978       LoadCombo cmbTemperatureTransducer, "TemperatureTransducer"
2979       LoadCombo cmbCirculationFlowMeter, "CirculationFlowMeter"
           'LoadCombo cmbSupermarketModel, "SupermarketPumpData"

2980       SetFrequencyCombo
           'load the TEMC combo boxes, too
2981       LoadCombo cmbTEMCAdapter, "TEMCAdapter"
2982       LoadCombo cmbTEMCAdditions, "TEMCAdditions"
2983       LoadCombo cmbTEMCCirculation, "TEMCCirculation"
2984       LoadCombo cmbTEMCDesignPressure, "TEMCDesignPressure"
2985       LoadCombo cmbTEMCNominalDischargeSize, "TEMCNominalDischargeSize"
2986       LoadCombo cmbTEMCDivisionType, "TEMCDivisionType"
2987       LoadCombo cmbTEMCImpellerType, "TEMCImpellerType"
2988       LoadCombo cmbTEMCInsulation, "TEMCInsulation"
2989       LoadCombo cmbTEMCJacketGasket, "TEMCJacketGasket"
2990       LoadCombo cmbTEMCMaterials, "TEMCMaterials"
2991       LoadCombo cmbTEMCModel, "TEMCModel"
2992       LoadCombo cmbTEMCNominalImpSize, "TEMCNominalImpSize"
2993       LoadCombo cmbTEMCOtherMotor, "TEMCOtherMotor"
2994       LoadCombo cmbTEMCNominalSuctionSize, "TEMCNominalSuctionSize"
2995       LoadCombo cmbTEMCVoltage, "TEMCVoltage"
2996       LoadCombo cmbTEMCPumpStages, "TEMCPumpStages"
2997       LoadCombo cmbTEMCTRG, "TEMCTRG"

           'LoadCombo frmSearch.cmbSearchModel, "Model"

           'fill memory arrays for dlookups
2998       FillArrays

           'choose the first tab
2999       frmPLCData.SSTab1.Tab = 0

           'set the grid column names
3000       Dim c As Column
3001       For Each c In DataGrid1.Columns
3002           Select Case c.DataField
               Case "TestDataID"
3003               c.Visible = False
3004           Case "SerialNumber"
3005               c.Visible = False
3006           Case "Date"
3007               c.Visible = False
3008           Case Else ' Show all other columns.
3009               c.Visible = True
3010               c.Alignment = dbgRight
3011           End Select
3012       Next c

3013       Set dgBalanceHoles.DataSource = rsBalanceHoles

3014       For Each c In dgBalanceHoles.Columns
3015           Select Case c.DataField
               Case "BalanceHoleID"
3016               c.Visible = False
3017           Case "SerialNo"
3018               c.Visible = False
3019           Case "Date"
3020               c.Visible = True
3021               c.Alignment = dbgCenter
3022               c.Width = 2000
3023           Case "Number"
3024               c.Visible = True
3025               c.Alignment = dbgCenter
3026               c.Width = 700
3027           Case "Diameter"
3028               c.Visible = False
3029           Case "Diameter1"
3030               c.Caption = "Diameter"
3031               c.Visible = True
3032               c.Alignment = dbgCenter
3033               c.Width = 700
3034           Case "BoltCircle1"
3035               c.Caption = "Bolt Circle"
3036               c.Visible = True
3037               c.Alignment = dbgCenter
3038               c.Width = 800
3039           Case "BoltCircle"
3040               c.Visible = False
3041           Case "SetNo"
3042               c.Visible = False
3043           Case Else ' Show all other columns.
3044               c.Visible = False
3045           End Select
3046       Next c

3047       BlankData

       '    If debugging <> 1 Then
               'get user initials
3048           frmLogin.Show
       '    End If

3049     optMfr(1).value = True
3050     frmMfr.Visible = False

3051       Pressed = True
' <VB WATCH>
3052       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3053       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Form_Load"

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
            vbwReportVariable "RetVal", RetVal
            vbwReportVariable "sSendStr", sSendStr
            vbwReportVariable "I", I
            vbwReportVariable "j", j
            vbwReportVariable "sTableName", sTableName
            vbwReportVariable "WhichServer", WhichServer
            vbwReportVariable "WhichDatabase", WhichDatabase
            vbwReportVariable "developmentDatabase", developmentDatabase
            vbwReportVariable "k", k
            vbwReportVariable "PLCAddress", PLCAddress
            vbwReportVariable "fdrive", fdrive
            vbwReportVariable "objWMIService", objWMIService
            vbwReportVariable "colProcesses", colProcesses
            vbwReportVariable "f", f
            vbwReportVariable "cnDevelopment", cnDevelopment
            vbwReportVariable "qyDevelopment", qyDevelopment
            vbwReportVariable "rsDevelopment", rsDevelopment
            vbwReportVariable "cnTestEpicor", cnTestEpicor
            vbwReportVariable "tempFSO", tempFSO
            vbwReportVariable "c", c
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub Form_Unload(Cancel As Integer)
' <VB WATCH>
3054       On Error GoTo vbwErrHandler
3055       Const VBWPROCNAME = "frmPLCData.Form_Unload"
3056       If vbwProtector.vbwTraceProc Then
3057           Dim vbwProtectorParameterString As String
3058           If vbwProtector.vbwTraceParameters Then
3059               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Cancel", Cancel) & ") "
3060           End If
3061           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3062       End If
' </VB WATCH>
3063       End
' <VB WATCH>
3064       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3065       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Form_Unload"

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
            vbwReportVariable "Cancel", Cancel
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub Label15_Click()
' <VB WATCH>
3066       On Error GoTo vbwErrHandler
3067       Const VBWPROCNAME = "frmPLCData.Label15_Click"
3068       If vbwProtector.vbwTraceProc Then
3069           Dim vbwProtectorParameterString As String
3070           If vbwProtector.vbwTraceParameters Then
3071               vbwProtectorParameterString = "()"
3072           End If
3073           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3074       End If
' </VB WATCH>
3075       frmDiagram.Show
' <VB WATCH>
3076       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3077       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "Label15_Click"

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

Private Sub lblAutoMan_Click(Index As Integer)
           '0 - Flow
           '1 - Suction
           '2 - Discharge
           '3 - Temperature
           '4 - A1 - Circ Flow
           '5 - A2 - RBH Temp
           '6 - A3 - RBH Press
           '7 - A4
' <VB WATCH>
3078       On Error GoTo vbwErrHandler
3079       Const VBWPROCNAME = "frmPLCData.lblAutoMan_Click"
3080       If vbwProtector.vbwTraceProc Then
3081           Dim vbwProtectorParameterString As String
3082           If vbwProtector.vbwTraceParameters Then
3083               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3084           End If
3085           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3086       End If
' </VB WATCH>

3087       Dim blnEnabled As Boolean

3088       If lblAutoMan(Index).Caption = "Auto" Then
3089           lblAutoMan(Index).Caption = "Man"
3090           blnEnabled = True
3091       Else
3092           lblAutoMan(Index).Caption = "Auto"
3093           blnEnabled = False
3094       End If

3095       Select Case Index
               Case 0
3096               txtFlowDisplay.Enabled = blnEnabled
3097           Case 1
3098               txtSuctionDisplay.Enabled = blnEnabled
3099           Case 2
3100               txtDischargeDisplay.Enabled = blnEnabled
3101           Case 3
3102               txtTemperatureDisplay.Enabled = blnEnabled
3103           Case 4
3104               txtAI1Display.Enabled = blnEnabled
3105           Case 5
3106               txtAI2Display.Enabled = blnEnabled
3107           Case 6
3108               txtAI3Display.Enabled = blnEnabled
3109           Case 7
3110               txtAI4Display.Enabled = blnEnabled
3111       End Select

' <VB WATCH>
3112       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3113       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "lblAutoMan_Click"

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
            vbwReportVariable "Index", Index
            vbwReportVariable "blnEnabled", blnEnabled
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub tmrNPSHr_Timer()
' <VB WATCH>
3114       On Error GoTo vbwErrHandler
3115       Const VBWPROCNAME = "frmPLCData.tmrNPSHr_Timer"
3116       If vbwProtector.vbwTraceProc Then
3117           Dim vbwProtectorParameterString As String
3118           If vbwProtector.vbwTraceParameters Then
3119               vbwProtectorParameterString = "()"
3120           End If
3121           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3122       End If
' </VB WATCH>
3123       tmrNPSHr.Enabled = False
3124       If frmNPSH.Visible = True Then
3125           btnRunNPSH_Click    'close test
3126       End If
' <VB WATCH>
3127       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3128       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tmrNPSHr_Timer"

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

Private Sub txtNPSH_Change(Index As Integer)
' <VB WATCH>
3129       On Error GoTo vbwErrHandler
3130       Const VBWPROCNAME = "frmPLCData.txtNPSH_Change"
3131       If vbwProtector.vbwTraceProc Then
3132           Dim vbwProtectorParameterString As String
3133           If vbwProtector.vbwTraceParameters Then
3134               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3135           End If
3136           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3137       End If
' </VB WATCH>
3138       If Index = 5 Then
3139           If frmNPSH.Visible = True Then
3140               If rsMisc.State = adStateOpen Then
3141                   rsMisc.Close
3142               End If
3143               rsMisc.CursorLocation = adUseClient
3144               rsMisc.Open "Select * from MiscParameters WHERE (ParameterName = 'AllowableTDHVariation');", cnPumpData, adOpenStatic, adLockOptimistic, adCmdText
3145               rsMisc.Fields("ParameterValue").value = txtNPSH(5).Text
3146               rsMisc.Update
3147           End If
3148       End If
' <VB WATCH>
3149       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3150       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtNPSH_Change"

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
            vbwReportVariable "Index", Index
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub txtNPSHFileLocation_Click()
' <VB WATCH>
3151       On Error GoTo vbwErrHandler
3152       Const VBWPROCNAME = "frmPLCData.txtNPSHFileLocation_Click"
3153       If vbwProtector.vbwTraceProc Then
3154           Dim vbwProtectorParameterString As String
3155           If vbwProtector.vbwTraceParameters Then
3156               vbwProtectorParameterString = "()"
3157           End If
3158           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3159       End If
' </VB WATCH>
3160       Dim sTempDir As String
3161       On Error Resume Next
3162       sTempDir = CurDir    'Remember the current active directory
3163       CommonDialog2.DialogTitle = "Select a directory" 'titlebar
3164       CommonDialog2.InitDir = "\\tei-main-01\f\en\groups\shared\calibration and rundown\npsh\" 'start dir, might be "C:\" or so also
3165       CommonDialog2.filename = "Select a Directory"  'Something in filenamebox
3166       CommonDialog2.Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
3167       CommonDialog2.Filter = "Directories|*.~#~" 'set files-filter to show dirs only
3168       CommonDialog2.CancelError = True 'allow escape key/cancel
3169       CommonDialog2.ShowSave   'show the dialog screen

3170       If Err <> 32755 Then    ' User didn't chose Cancel.
               'Me.SDir.Text = CurDir
3171       End If

       '    ChDir sTempDir  'restore path to what it was at entering

3172   Me.txtNPSHFileLocation.Text = CommonDialog2.filename

' <VB WATCH>
3173       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3174       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtNPSHFileLocation_Click"

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
            vbwReportVariable "sTempDir", sTempDir
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub





Private Sub txtTitle_LostFocus(Index As Integer)
' <VB WATCH>
3175       On Error GoTo vbwErrHandler
3176       Const VBWPROCNAME = "frmPLCData.txtTitle_LostFocus"
3177       If vbwProtector.vbwTraceProc Then
3178           Dim vbwProtectorParameterString As String
3179           If vbwProtector.vbwTraceParameters Then
3180               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3181           End If
3182           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3183       End If
' </VB WATCH>

3184       ChangeTitles Index

' <VB WATCH>
3185       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3186       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtTitle_LostFocus"

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
            vbwReportVariable "Index", Index
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub ChangeTitles(ChannelNo As Integer)
' <VB WATCH>
3187       On Error GoTo vbwErrHandler
3188       Const VBWPROCNAME = "frmPLCData.ChangeTitles"
3189       If vbwProtector.vbwTraceProc Then
3190           Dim vbwProtectorParameterString As String
3191           If vbwProtector.vbwTraceParameters Then
3192               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ChannelNo", ChannelNo) & ") "
3193           End If
3194           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3195       End If
' </VB WATCH>
3196       Dim I As Integer
3197       Dim S As String

3198       If txtTitle(ChannelNo).Locked = True Then
' <VB WATCH>
3199       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3200           Exit Sub
3201       End If

3202       Dim qy As New ADODB.Command
3203       Dim rs As New ADODB.Recordset

3204       qy.ActiveConnection = cnPumpData

           'see if we have an entry in the table
3205       qy.CommandText = "SELECT * FROM AITitles " & _
                             "WHERE (((AITitles.SerialNo)= '" & txtSN.Text & "') " & _
                             "AND ((AITitles.Date)= #" & cmbTestDate.Text & "#) " & _
                             "AND ((AITitles.Channel)=" & ChannelNo & "));"

3206       With rs     'open the recordset for the query
3207           .CursorLocation = adUseClient
3208           .CursorType = adOpenStatic
3209           .LockType = adLockOptimistic
3210           .Open qy
3211       End With

3212       If (rs.BOF = True And rs.EOF = True) Then  'new record
3213           rs.AddNew
3214           rs.Fields("SerialNo") = txtSN.Text
3215           rs.Fields("Date") = cmbTestDate.Text
3216           rs.Fields("Channel") = CByte(ChannelNo)
3217           rs.Fields("Title") = txtTitle(ChannelNo).Text
3218           rs.Update
3219       Else    'we have an entry, modify it
3220           rs.Fields("SerialNo") = txtSN.Text
3221           rs.Fields("Date") = cmbTestDate.Text
3222           rs.Fields("Channel") = CByte(ChannelNo)
3223           rs.Fields("Title") = txtTitle(ChannelNo).Text
3224           rs.Update
3225       End If

3226       rs.Close
3227       Set rs = Nothing
3228       Set qy = Nothing

' <VB WATCH>
3229       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3230       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ChangeTitles"

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
            vbwReportVariable "ChannelNo", ChannelNo
            vbwReportVariable "I", I
            vbwReportVariable "S", S
            vbwReportVariable "qy", qy
            vbwReportVariable "rs", rs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub optKW_Click(Index As Integer)
' <VB WATCH>
3231       On Error GoTo vbwErrHandler
3232       Const VBWPROCNAME = "frmPLCData.optKW_Click"
3233       If vbwProtector.vbwTraceProc Then
3234           Dim vbwProtectorParameterString As String
3235           If vbwProtector.vbwTraceParameters Then
3236               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3237           End If
3238           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3239       End If
' </VB WATCH>
3240       Select Case Index
               Case 0  'add 3 powers
3241               txtKW.Enabled = False
3242           Case 1  'enter kw
3243               txtKW.Enabled = True
3244           Case 2  'use analog in 4
3245               txtKW.Enabled = False
3246       End Select
' <VB WATCH>
3247       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3248       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "optKW_Click"

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
            vbwReportVariable "Index", Index
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub optMfr_Click(Index As Integer)
' <VB WATCH>
3249       On Error GoTo vbwErrHandler
3250       Const VBWPROCNAME = "frmPLCData.optMfr_Click"
3251       If vbwProtector.vbwTraceProc Then
3252           Dim vbwProtectorParameterString As String
3253           If vbwProtector.vbwTraceParameters Then
3254               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3255           End If
3256           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3257       End If
' </VB WATCH>
3258       frmTEMC.Visible = optMfr(1).value
3259       frmChempump.Visible = optMfr(0).value
3260       frmTEMCData.Visible = optMfr(1).value
3261       txtModelNo_Change
' <VB WATCH>
3262       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3263       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "optMfr_Click"

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
            vbwReportVariable "Index", Index
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Private Sub tmrGetDDE_Timer()
' <VB WATCH>
3264       On Error GoTo vbwErrHandler
3265       Const VBWPROCNAME = "frmPLCData.tmrGetDDE_Timer"
3266       If vbwProtector.vbwTraceProc Then
3267           Dim vbwProtectorParameterString As String
3268           If vbwProtector.vbwTraceParameters Then
3269               vbwProtectorParameterString = "()"
3270           End If
3271           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3272       End If
' </VB WATCH>

       'get here every second... get plc and magtrol data

3273       Dim sSendStr As String
3274       Dim I As Integer
3275       Dim VoltMul As Double

3276       If Calibrating Then
' <VB WATCH>
3277       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3278           Exit Sub
3279       End If

3280       If debugging Then
               'Exit Sub
3281       End If


3282       If boPLCOperating = True Then
3283           frmPLCData.shpGetPLCData.Visible = True    'turn the PLC led on

               'convert the plc data into real numbers
               'the following data are type real
3284           txtFlow.Text = ConvertToReal("4050")
3285           txtSuction.Text = ConvertToReal("4052")
3286           txtDischarge.Text = ConvertToReal("4054")
3287           txtTemperature.Text = ConvertToReal("4056")

3288           txtValvePosition.Text = ConvertToLong("2004")

3289           frmPLCData.txtTC1.Text = ConvertToLong("2200")
3290           frmPLCData.txtTC2.Text = ConvertToLong("2202")
3291           frmPLCData.txtTC3.Text = ConvertToLong("2204")
3292           frmPLCData.txtTC4.Text = ConvertToLong("2206")

3293           frmPLCData.txtAI1.Text = ConvertToReal("4060")
3294           frmPLCData.txtAI2.Text = ConvertToReal("4062")
3295           frmPLCData.txtAI3.Text = ConvertToReal("4064")
3296           frmPLCData.txtAI4.Text = ConvertToReal("4066")

3297           frmPLCData.txtPCoef.Text = ConvertToLong("4036")
3298           frmPLCData.txtICoef.Text = ConvertToLong("4037")
3299           frmPLCData.txtDCoef.Text = ConvertToLong("4040")

3300           frmPLCData.txtSetPoint.Text = ConvertToLong("4035")
3301           frmPLCData.txtInHg.Text = ConvertToLong("1460")


               'modify the data from PLC format to format that we can use
               'and update the screen
3302           If txtFlowDisplay.Enabled = False Then
3303               frmPLCData.txtFlowDisplay = Format$(txtFlow.Text, "###0.00")
3304           End If
3305           If txtSuctionDisplay.Enabled = False Then
3306               frmPLCData.txtSuctionDisplay = Format$((txtSuction.Text) / 10, "##0.00")
3307           End If
3308           If txtDischargeDisplay.Enabled = False Then
3309               frmPLCData.txtDischargeDisplay = Format$(txtDischarge.Text, "##0.00")
3310           End If
3311           If txtTemperatureDisplay.Enabled = False Then
3312               frmPLCData.txtTemperatureDisplay = Format$(txtTemperature.Text, "##0.00")
3313           End If
3314           frmPLCData.txtValvePositionDisplay = (txtValvePosition.Text)

3315           frmPLCData.txtTC1Display = Format$((txtTC1.Text) / 10, "##0.0")
3316           frmPLCData.txtTC2Display = Format$((txtTC2.Text) / 10, "##0.0")
3317           frmPLCData.txtTC3Display = Format$((txtTC3.Text) / 10, "##0.0")
3318           frmPLCData.txtTC4Display = Format$((txtTC4.Text) / 10, "##0.0")

3319           If txtAI1Display.Enabled = False Then
3320               frmPLCData.txtAI1Display = Format$(txtAI1.Text, "##0.00")
3321           End If
3322           If txtAI2Display.Enabled = False Then
3323               frmPLCData.txtAI2Display = Format$(txtAI2.Text, "##0.00")
3324           End If
3325           If txtAI3Display.Enabled = False Then
3326               frmPLCData.txtAI3Display = Format$(txtAI3.Text, "##0.00")
3327           End If
3328           If txtAI4Display.Enabled = False Then
3329               frmPLCData.txtAI4Display = Format$(txtAI4.Text, "##0.00")
3330           End If

3331           frmPLCData.txtSetPointDisplay = (txtSetPoint.Text)

3332           frmPLCData.txtInHgDisplay = Format$(txtInHg.Text / 100, "00.00")

3333           frmPLCData.shpGetPLCData.Visible = False   'turn the PLC led off

3334           frmPLCData.shpGetMagtrolData.Visible = True 'turn the Magtrol led on
3335       End If

3336       If boMagtrolOperating = True Then


               'get the data from the Magtrol
3337           If Right(cmbMagtrol.List(cmbMagtrol.ListIndex), 4) = "5300" Then
3338               sSendStr = vbCrLf
3339               sData = Space$(68)
3340               VoltMul = Sqr(3)
3341           Else
3342               sSendStr = "OT" & vbCrLf
3343               sData = Space$(183)
3344               VoltMul = 1#
3345           End If

3346           On Error GoTo noresponse
3347           If UsingNatInst Then
3348               ibwrt iUD, sSendStr
3349               ibrd iUD, sData

                   'parse the Magrol response
       '            vResponse = CWGPIB1.Tasks("Number Parser").Parse(sData)
3350           Else
                   'Dim Databack As String
3351               sData = TCP.SendGetData("OT")
3352           End If

3353               Dim vSplit() As String
3354               vSplit = Split(Right(sData, Len(sData) - 1), ",")
3355               If UBound(vSplit) > 0 Then
3356                  ReDim vResponse(UBound(vSplit))
3357               End If
3358               For I = 0 To UBound(vSplit) - 1
3359                   If Len(vSplit(I)) <> 0 Then
3360                       vResponse(I) = CDbl(vSplit(I))
3361                   End If
3362               Next I

               'format the parsed response
3363           Dim dd As String
3364           dd = "- -"

3365           If Not IsEmpty(vResponse) Then
               '8 entries for 5300 and 12 for the 6530
3366               If UBound(vResponse) = 8 Or UBound(vResponse) = 12 Then
                       'put the responses into the correct text box
3367                   txtV1.Text = Format$(VoltMul * vResponse(1), "###0.0")   'we get back phase voltage and we want line voltage

3368                   Select Case vResponse(0)
                           Case Is < 1
3369                           txtI1.Text = Format$(vResponse(0), "0.0000")
3370                       Case Is < 10
3371                           txtI1.Text = Format$(vResponse(0), "0.000")
3372                       Case Is < 100
3373                           txtI1.Text = Format$(vResponse(0), "00.00")
3374                       Case Else
3375                           txtI1.Text = Format$(vResponse(0), "000.0")
3376                   End Select

3377                   Select Case vResponse(3)
                           Case Is < 1
3378                           txtI2.Text = Format$(vResponse(3), "0.0000")
3379                       Case Is < 10
3380                           txtI2.Text = Format$(vResponse(3), "0.000")
3381                       Case Is < 100
3382                           txtI2.Text = Format$(vResponse(3), "00.00")
3383                       Case Else
3384                           txtI2.Text = Format$(vResponse(3), "000.0")
3385                   End Select

3386                   Select Case vResponse(6)
                           Case Is < 1
3387                           txtI3.Text = Format$(vResponse(6), "0.0000")
3388                       Case Is < 10
3389                           txtI3.Text = Format$(vResponse(6), "0.000")
3390                       Case Is < 100
3391                           txtI3.Text = Format$(vResponse(6), "00.00")
3392                       Case Else
3393                           txtI3.Text = Format$(vResponse(6), "000.0")
3394                   End Select

3395                   txtP1.Text = Format$(vResponse(2) / 1000, "##0.00")     '/ by 1000 to show kW
3396                   txtV2.Text = Format$(VoltMul * vResponse(4), "###0.0")
                       'txtI2.Text = Format$(vResponse(3), "###0.0")
3397                   txtP2.Text = Format$(vResponse(5) / 1000, "##0.00")
3398                   txtV3.Text = Format$(VoltMul * vResponse(7), "###0.0")
                       'txtI3.Text = Format$(vResponse(6), "###0.0")
3399                   txtP3.Text = Format$(vResponse(8) / 1000, "##0.00")
3400                   If (vResponse(0) * vResponse(1) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7)) <> 0 Then
                           'if we have some measured current
                           'pf = sum of power/sum of VA
3401                       If Right(cmbMagtrol.List(cmbMagtrol.ListIndex), 4) = "5300" Then
                               'add kw responses and / by 1000 to get to kW
3402                           txtKW.Text = (vResponse(2) + vResponse(5) + vResponse(8)) / 1000
3403                           txtPF.Text = Format$(100 * (vResponse(2) + vResponse(5) + vResponse(8)) / (vResponse(1) * vResponse(0) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7)), "0.00")
3404                       Else
3405                           txtKW.Text = (vResponse(2) + vResponse(8)) / 1000
3406                           txtPF.Text = Format$(100 * (vResponse(2) + vResponse(8)) / ((Sqr(3) / 3) * (vResponse(1) * vResponse(0) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7))), "0.00")
3407                       End If
3408                       Select Case Val(txtKW.Text)
                               Case Is < 1
3409                               txtKW.Text = Format$(txtKW.Text, "0.00000")
3410                           Case Is < 10
3411                               txtKW.Text = Format$(txtKW.Text, "0.0000")
3412                           Case Is < 100
3413                               txtKW.Text = Format$(txtKW.Text, "00.000")
3414                           Case Else
3415                               txtKW.Text = Format$(txtKW.Text, "000.00")
3416                       End Select
3417                   Else
3418                       txtPF = dd
3419                   End If
3420               Else
                       'no response, show all -- in text boxes
3421                   txtV1.Text = dd
3422                   txtI1.Text = dd
3423                   txtP1.Text = dd
3424                   txtV2.Text = dd
3425                   txtI2.Text = dd
3426                   txtP2.Text = dd
3427                   txtV3.Text = dd
3428                   txtI3.Text = dd
3429                   txtP3.Text = dd
3430                   txtPF = dd
3431                   txtKW = dd
3432               End If
3433           End If
3434       Else    'magtrol not operating
3435           Dim dbl As Double

3436           If optKW(0).value = True Then   'add 3 powers
3437               txtKW.Text = Val(txtP1.Text) + Val(txtP2.Text) + Val(txtP3.Text)
3438           End If
3439           If optKW(1).value = True Then   'enter kw
3440               txtP1.Text = Val(txtKW.Text) / 3
3441               txtP2.Text = Val(txtKW.Text) / 3
3442               txtP3.Text = Val(txtKW.Text) / 3
3443           End If
3444           If optKW(2).value = True Then   'use ai4
3445               txtKW.Text = txtAI4Display.Text
3446               txtP1.Text = Val(txtKW.Text) / 3
3447               txtP2.Text = Val(txtKW.Text) / 3
3448               txtP3.Text = Val(txtKW.Text) / 3
3449           End If

3450           dbl = Val(txtV1.Text) * Val(txtI1.Text)
3451           dbl = dbl + Val(txtV2.Text) * Val(txtI2.Text)
3452           dbl = dbl + Val(txtV3.Text) * Val(txtI3.Text)
3453           If dbl <> 0 Then
3454               txtPF.Text = Format$((Val(txtKW.Text) * 1000 * 3 * 100 / (dbl * Sqr(3))), "0.00")
3455           End If
3456       End If

3457   noresponse:
3458   On Error GoTo vbwErrHandler
3459       frmPLCData.shpGetMagtrolData.Visible = False   'turn the Magtrol led off

           'update the little PLC chart
3460       For I = 1 To 99
3461           vPlot(0, I) = vPlot(0, I + 1)
3462           vPlot(1, I) = vPlot(1, I + 1)
3463       Next I
3464       vPlot(0, 100) = txtSetPointDisplay
3465       vPlot(1, 100) = txtFlowDisplay

           'do NPSH stuff
3466       Dim SuctVelHead As Single
3467       Dim DischVelHead As Single
3468       Dim Conversion As Single
3469       Dim SuctionPSIA As Single
3470       Dim DischargePSIA As Single
3471       Dim VaporPress As Single
3472       Dim SpecVolume As Single
3473       Dim NPSHa As Single
3474       Dim NPSHr As Single
3475       Dim TDH As Single
3476       Dim pd As Single


           'velocity head
3477       If cmbSuctDia.ListIndex = -1 Then   'if no suction diameter chosen
3478           SuctVelHead = 0
3479       Else
       '        pd = DLookup("ActualDia", "PipeDiameters", "ID = " & cmbSuctDia.ListIndex + 1)
3480           pd = DLookupA(ActualColNo, PipeDiameters, IDColNo, cmbSuctDia.ItemData(cmbSuctDia.ListIndex) + 1)
3481           SuctVelHead = (0.002592 * Val(txtFlow) ^ 2) / (pd ^ 4)
3482       End If

3483       If cmbDischDia.ListIndex = -1 Then     'if no discharge diameter chosen
3484           DischVelHead = 0
3485       Else
       '        pd = DLookup("ActualDia", "PipeDiameters", "ID = " & cmbDischDia.ListIndex + 1)
3486           pd = DLookupA(ActualColNo, PipeDiameters, IDColNo, cmbDischDia.ItemData(cmbDischDia.ListIndex) + 1)
3487           DischVelHead = (0.002592 * Val(txtFlow) ^ 2) / (pd ^ 4)
3488       End If

           'convert gauges to absolute
3489       If txtInHgDisplay.Text = "" Then
3490           Conversion = 0
3491       Else
3492           Conversion = txtInHgDisplay * 0.491
3493       End If

3494       SuctionPSIA = Val(txtSuctionDisplay) + Conversion
3495       DischargePSIA = Val(txtDischargeDisplay) + Conversion


           'lookup vapor pressure and specific volume in the arrays that we made
           'if temp is out of range, say so and exit
3496       If Val(txtTemperatureDisplay) < 40 Or Val(txtTemperatureDisplay) > 165 Then
3497           txtNPSHa = 0
' <VB WATCH>
3498       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3499           Exit Sub
3500       Else
3501           I = Val(txtTemperatureDisplay) - 40
       '        VaporPress = DLookup("VaporPressure", "VaporPressure", "ID = " & I)
       '        SpecVolume = DLookup("SpecificVolume", "VaporPressure", "ID= " & I)
3502           VaporPress = DLookupA(VaporPressureColNo, VaporPressure, IDColNo, I)
3503           SpecVolume = DLookupA(SpecificVolumeColNo, VaporPressure, IDColNo, I)
3504       End If

3505       If Not ((txtSuctHeight = "") Or (txtDischHeight = "") Or Not IsNumeric(txtSuctHeight) Or Not IsNumeric(txtDischHeight)) Then
               'NPSHa
3506           NPSHa = (144 * SpecVolume * (SuctionPSIA - VaporPress)) + (txtSuctHeight / 12) + SuctVelHead
       '        NPSHa = CalcTDH(DischargePSIA, SuctionPSIA, 0, DischVelHead, 0, txtTemperature)
3507           txtNPSHa = Format$(NPSHa, "##0.00")

               'tdh
3508           TDH = CalcTDH(DischargePSIA, SuctionPSIA, 0, (DischVelHead - SuctVelHead), (txtDischHeight / 12) - (txtSuctHeight / 12), txtTemperatureDisplay)
3509           txtTDH = Format$(TDH, "##0.00")

3510           If frmNPSH.Visible = True Then
3511               If Val(txtTDH.Text) > 0 Then
3512                   txtNPSH(2).Text = Format(100 * Val(txtTDH.Text) / Val(txtNPSH(3).Text), "##0.00")
3513                   txtNPSH(1).Text = Format(100 * Val(txtFlow.Text) / Val(txtNPSH(0).Text), "##0.00")
                       'check for tdh variation
3514                   If Abs(Val(txtNPSH(1)) - 100) > Val(txtNPSH(5).Text) Then
3515                       MsgBox "The TDH value has varied more than " & txtNPSH(5) & " %. NPSHr data will NOT be written to the data table", vbOKOnly, "TDH variation too large"
3516                       btnRunNPSH_Click
3517                   Else    'tdh variation small
3518                       If Val(txtNPSH(2).Text) <= 97 Then
                               'btnRunNPSH_Click
                               'write the npsh and save
3519                           If WroteNPSHr = False Then
3520                               txtNPSH(4).Text = txtNPSHa.Text
3521                               rsTestData!NPSHr = txtNPSHa.Text
3522                               rsTestData.Update
3523                               rsEff!NPSHr = txtNPSHa.Text
3524                               rsEff.Update
3525                               WroteNPSHr = True
3526                               tmrNPSHr.Interval = 5000
3527                               tmrNPSHr.Enabled = True
3528                           End If
3529                       End If  'val < 97
3530                   End If  'check for tdh variation
3531               End If 'val tdh <=0
3532           Else    'frm not visible
                   'txtNPSHa = Format$(0, "##0.00")
3533           End If  'if frm visible

3534       Else
3535           txtNPSHa = 0
3536       End If
' <VB WATCH>
3537       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3538       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tmrGetDDE_Timer"

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
            vbwReportVariable "sSendStr", sSendStr
            vbwReportVariable "I", I
            vbwReportVariable "VoltMul", VoltMul
            vbwReportVariable "vSplit", vSplit
            vbwReportVariable "dd", dd
            vbwReportVariable "dbl", dbl
            vbwReportVariable "SuctVelHead", SuctVelHead
            vbwReportVariable "DischVelHead", DischVelHead
            vbwReportVariable "Conversion", Conversion
            vbwReportVariable "SuctionPSIA", SuctionPSIA
            vbwReportVariable "DischargePSIA", DischargePSIA
            vbwReportVariable "VaporPress", VaporPress
            vbwReportVariable "SpecVolume", SpecVolume
            vbwReportVariable "NPSHa", NPSHa
            vbwReportVariable "NPSHr", NPSHr
            vbwReportVariable "TDH", TDH
            vbwReportVariable "pd", pd
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub tmrStartUp_Timer()
           'we waited for a while, disable the timer
' <VB WATCH>
3539       On Error GoTo vbwErrHandler
3540       Const VBWPROCNAME = "frmPLCData.tmrStartUp_Timer"
3541       If vbwProtector.vbwTraceProc Then
3542           Dim vbwProtectorParameterString As String
3543           If vbwProtector.vbwTraceParameters Then
3544               vbwProtectorParameterString = "()"
3545           End If
3546           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3547       End If
' </VB WATCH>
3548       tmrStartUp.Enabled = False
' <VB WATCH>
3549       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3550       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "tmrStartUp_Timer"

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
Public Function SetCombo(cmbComboName As ComboBox, sName As String, rs As ADODB.Recordset)
       'set the pump parameter combo box to the right data based upon
       'the number in the database
' <VB WATCH>
3551       On Error GoTo vbwErrHandler
3552       Const VBWPROCNAME = "frmPLCData.SetCombo"
3553       If vbwProtector.vbwTraceProc Then
3554           Dim vbwProtectorParameterString As String
3555           If vbwProtector.vbwTraceParameters Then
3556               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
3557               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sName", sName) & ", "
3558               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("rs", rs) & ") "
3559           End If
3560           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3561       End If
' </VB WATCH>

3562       Dim I As Integer
3563       Dim sParam As String
3564       Dim qy As New ADODB.Command
3565       Dim rs1 As New ADODB.Recordset

3566       If rs.Fields(sName).ActualSize <> 0 Then     'if there's an entry
3567           sParam = rs.Fields(sName)                'get the index number
3568           qy.ActiveConnection = cnPumpData
3569           qy.CommandText = "SELECT * FROM " & sName & " WHERE " & sName & " = " & sParam
3570           Set rs1 = qy.Execute()                                  'get the record for the index number

3571           If rs1.BOF = True And rs1.EOF = True Then
3572               cmbComboName.ListIndex = -1                             'else, remove any pointer
' <VB WATCH>
3573       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3574               Exit Function
3575           End If

3576           For I = 0 To cmbComboName.ListCount - 1                     'go through the combobox entries
3577               If cmbComboName.ItemData(I) = rs1.Fields(0) Then     'see when we find the desired index number
3578                   cmbComboName.ListIndex = I                                              'if we do, set the combo box
3579                   Exit For                                            'and we're done
3580               End If
3581               cmbComboName.ListIndex = -1                             'else, remove any pointer
3582           Next I
3583       Else
3584           cmbComboName.ListIndex = -1
3585       End If

' <VB WATCH>
3586       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3587       Exit Function
' <VB WATCH>
3588       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3589       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SetCombo"

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
            vbwReportVariable "sName", sName
            vbwReportVariable "I", I
            vbwReportVariable "sParam", sParam
            vbwReportVariable "cmbComboName", cmbComboName
            vbwReportVariable "rs", rs
            vbwReportVariable "qy", qy
            vbwReportVariable "rs1", rs1
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Private Function SetComboTestSetup(cmbComboName As ComboBox, sFieldName As String, sTableName As String, rs As ADODB.Recordset)
       'set the pump parameter combo box to the right data based upon
       'the number in the database
' <VB WATCH>
3590       On Error GoTo vbwErrHandler
3591       Const VBWPROCNAME = "frmPLCData.SetComboTestSetup"
3592       If vbwProtector.vbwTraceProc Then
3593           Dim vbwProtectorParameterString As String
3594           If vbwProtector.vbwTraceParameters Then
3595               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
3596               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sFieldName", sFieldName) & ", "
3597               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sTableName", sTableName) & ", "
3598               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("rs", rs) & ") "
3599           End If
3600           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3601       End If
' </VB WATCH>

       'same as setcombo, except here we also pass in the field name

3602       Dim I As Integer
3603       Dim sParam As String
3604       Dim qy As New ADODB.Command
3605       Dim rs1 As New ADODB.Recordset

3606       If rs.Fields(sFieldName).ActualSize <> 0 Then
               'if plc number, adjust plcaddress id numbers 1 and 2 to plc 8 and 9 respectively
3607           If sTableName = "CirculationFlowMeter" Then
                   'sParam = rs.Fields(sFieldName) + 7
3608               sParam = rs.Fields(sFieldName)
3609               If Val(sParam) < 4 Then
3610                   sParam = str(Val(sParam) + 4)
3611                   rs.Fields(sFieldName) = sParam
3612               End If
3613           Else
3614               sParam = rs.Fields(sFieldName)
3615           End If
3616           qy.ActiveConnection = cnPumpData
3617           qy.CommandText = "SELECT * FROM " & sTableName & " WHERE " & sTableName & " = " & sParam
3618           Set rs1 = qy.Execute()

3619           For I = 0 To cmbComboName.ListCount - 1
3620               If cmbComboName.ItemData(I) = rs1.Fields(0) Then
3621                   cmbComboName.ListIndex = I
3622                   Exit For
3623               End If
3624               cmbComboName.ListIndex = -1
3625           Next I
3626       Else
3627           cmbComboName.ListIndex = -1
3628       End If

' <VB WATCH>
3629       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3630       Exit Function
' <VB WATCH>
3631       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3632       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SetComboTestSetup"

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
            vbwReportVariable "sFieldName", sFieldName
            vbwReportVariable "sTableName", sTableName
            vbwReportVariable "I", I
            vbwReportVariable "sParam", sParam
            vbwReportVariable "cmbComboName", cmbComboName
            vbwReportVariable "rs", rs
            vbwReportVariable "qy", qy
            vbwReportVariable "rs1", rs1
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Private Sub DisablePumpDataControls()
           'disable the pump data controls cause we're just showing what we found
' <VB WATCH>
3633       On Error GoTo vbwErrHandler
3634       Const VBWPROCNAME = "frmPLCData.DisablePumpDataControls"
3635       If vbwProtector.vbwTraceProc Then
3636           Dim vbwProtectorParameterString As String
3637           If vbwProtector.vbwTraceParameters Then
3638               vbwProtectorParameterString = "()"
3639           End If
3640           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3641       End If
' </VB WATCH>

3642       txtSalesOrderNumber.Enabled = False
3643       frmMfr.Enabled = False
3644       txtShpNo.Enabled = False
3645       txtBilNo.Enabled = False
3646       txtDesignFlow.Enabled = False
3647       txtDesignTDH.Enabled = False

3648       frmMiscPumpData.Enabled = False

3649       txtModelNo.Enabled = False
3650       txtImpellerDia.Enabled = False

3651       frmTEMC.Enabled = False
3652       frmChempump.Enabled = False

3653       txtRemarks.Enabled = False
3654       Me.cmdAddNewTestDate.Visible = False

3655       cmdEnterPumpData.Enabled = False

' <VB WATCH>
3656       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3657       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DisablePumpDataControls"

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
Private Sub DisableTestSetupDataControls()
' <VB WATCH>
3658       On Error GoTo vbwErrHandler
3659       Const VBWPROCNAME = "frmPLCData.DisableTestSetupDataControls"
3660       If vbwProtector.vbwTraceProc Then
3661           Dim vbwProtectorParameterString As String
3662           If vbwProtector.vbwTraceParameters Then
3663               vbwProtectorParameterString = "()"
3664           End If
3665           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3666       End If
' </VB WATCH>

3667       cmbTestSpec.Enabled = False
3668       txtWho.Enabled = False
3669       txtRMA.Enabled = False

3670       frmLoopAndXducer.Enabled = False
3671       frmElecData.Enabled = False
3672       frmPerfMods.Enabled = False
3673       frmOtherFiles.Enabled = False
3674       frmInstrumentTags.Enabled = False
3675       frmTAndI.Enabled = False
3676       frmThrustBalMods.Enabled = False
3677       txtTestSetupRemarks.Enabled = False

3678       cmdEnterTestSetupData.Enabled = False
3679       cmbPLCNo.Enabled = False
' <VB WATCH>
3680       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3681       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DisableTestSetupDataControls"

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
Private Sub DisableTestDataControls()
' <VB WATCH>
3682       On Error GoTo vbwErrHandler
3683       Const VBWPROCNAME = "frmPLCData.DisableTestDataControls"
3684       If vbwProtector.vbwTraceProc Then
3685           Dim vbwProtectorParameterString As String
3686           If vbwProtector.vbwTraceParameters Then
3687               vbwProtectorParameterString = "()"
3688           End If
3689           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3690       End If
' </VB WATCH>

3691       cmbPLCLoop.Enabled = False
3692       frmPumpData.Enabled = False
3693       frmThermocouples.Enabled = False
3694       frmAI.Enabled = False
3695       frmMagtrol.Enabled = False
3696       fmrMiscTestData.Enabled = False
3697       frmPLCMisc.Enabled = False
3698       DataGrid1.Enabled = False
3699       DataGrid2.Enabled = False
3700       cmdEnterTestData.Enabled = False

' <VB WATCH>
3701       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3702       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DisableTestDataControls"

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
Private Sub EnableTestSetupDataControls()
' <VB WATCH>
3703       On Error GoTo vbwErrHandler
3704       Const VBWPROCNAME = "frmPLCData.EnableTestSetupDataControls"
3705       If vbwProtector.vbwTraceProc Then
3706           Dim vbwProtectorParameterString As String
3707           If vbwProtector.vbwTraceParameters Then
3708               vbwProtectorParameterString = "()"
3709           End If
3710           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3711       End If
' </VB WATCH>

3712       cmbTestSpec.Enabled = True
3713       txtWho.Enabled = True
3714       txtRMA.Enabled = True

3715       frmLoopAndXducer.Enabled = True
3716       frmElecData.Enabled = True
3717       frmPerfMods.Enabled = True
3718       frmOtherFiles.Enabled = True
3719       frmInstrumentTags.Enabled = True
3720       frmTAndI.Enabled = True
3721       frmThrustBalMods.Enabled = True
3722       txtTestSetupRemarks.Enabled = True

3723       cmdEnterTestSetupData.Enabled = True
3724       cmbPLCNo.Enabled = True
' <VB WATCH>
3725       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3726       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "EnableTestSetupDataControls"

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
Private Sub EnableTestDataControls()
' <VB WATCH>
3727       On Error GoTo vbwErrHandler
3728       Const VBWPROCNAME = "frmPLCData.EnableTestDataControls"
3729       If vbwProtector.vbwTraceProc Then
3730           Dim vbwProtectorParameterString As String
3731           If vbwProtector.vbwTraceParameters Then
3732               vbwProtectorParameterString = "()"
3733           End If
3734           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3735       End If
' </VB WATCH>

3736       cmbPLCLoop.Enabled = True
3737       frmPumpData.Enabled = True
3738       frmThermocouples.Enabled = True
3739       frmAI.Enabled = True
3740       frmMagtrol.Enabled = True
3741       fmrMiscTestData.Enabled = True
3742       frmPLCMisc.Enabled = True
3743       DataGrid1.Enabled = True
3744       DataGrid2.Enabled = True
3745       cmdEnterTestData.Enabled = True

' <VB WATCH>
3746       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3747       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "EnableTestDataControls"

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
Private Sub EnablePumpDataControls()
           'disable the pump data controls cause we're just showing what we found
' <VB WATCH>
3748       On Error GoTo vbwErrHandler
3749       Const VBWPROCNAME = "frmPLCData.EnablePumpDataControls"
3750       If vbwProtector.vbwTraceProc Then
3751           Dim vbwProtectorParameterString As String
3752           If vbwProtector.vbwTraceParameters Then
3753               vbwProtectorParameterString = "()"
3754           End If
3755           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3756       End If
' </VB WATCH>

3757       txtSalesOrderNumber.Enabled = True
3758       frmMfr.Enabled = True
3759       txtShpNo.Enabled = True
3760       txtBilNo.Enabled = True
3761       txtDesignFlow.Enabled = True
3762       txtDesignTDH.Enabled = True

3763       frmMiscPumpData.Enabled = True

3764       txtModelNo.Enabled = True
3765       txtImpellerDia.Enabled = True

3766       frmTEMC.Enabled = True
3767       frmChempump.Enabled = True

3768       txtRemarks.Enabled = True
3769       Me.cmdAddNewTestDate.Visible = True

3770       cmdEnterPumpData.Enabled = True

' <VB WATCH>
3771       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3772       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "EnablePumpDataControls"

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
Private Sub EnableMagtrolFields()
' <VB WATCH>
3773       On Error GoTo vbwErrHandler
3774       Const VBWPROCNAME = "frmPLCData.EnableMagtrolFields"
3775       If vbwProtector.vbwTraceProc Then
3776           Dim vbwProtectorParameterString As String
3777           If vbwProtector.vbwTraceParameters Then
3778               vbwProtectorParameterString = "()"
3779           End If
3780           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3781       End If
' </VB WATCH>
3782       txtV1.Enabled = True
3783       txtV2.Enabled = True
3784       txtV3.Enabled = True
3785       txtI1.Enabled = True
3786       txtI2.Enabled = True
3787       txtI3.Enabled = True
3788       txtP1.Enabled = True
3789       txtP2.Enabled = True
3790       txtP3.Enabled = True
3791       optKW(0).Visible = True
3792       optKW(1).Visible = True
3793       optKW(2).Visible = True
3794       optKW(1).value = True
3795       optKW_Click (1)
' <VB WATCH>
3796       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3797       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "EnableMagtrolFields"

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
Private Sub DisableMagtrolFields()
' <VB WATCH>
3798       On Error GoTo vbwErrHandler
3799       Const VBWPROCNAME = "frmPLCData.DisableMagtrolFields"
3800       If vbwProtector.vbwTraceProc Then
3801           Dim vbwProtectorParameterString As String
3802           If vbwProtector.vbwTraceParameters Then
3803               vbwProtectorParameterString = "()"
3804           End If
3805           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3806       End If
' </VB WATCH>
3807       txtV1.Enabled = False
3808       txtV2.Enabled = False
3809       txtV3.Enabled = False
3810       txtI1.Enabled = False
3811       txtI2.Enabled = False
3812       txtI3.Enabled = False
3813       txtP1.Enabled = False
3814       txtP2.Enabled = False
3815       txtP3.Enabled = False
3816       txtKW.Enabled = False
3817       optKW(0).Visible = False
3818       optKW(1).Visible = False
3819       optKW(2).Visible = False
' <VB WATCH>
3820       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3821       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DisableMagtrolFields"

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
Private Sub EnablePLCFields()
' <VB WATCH>
3822       On Error GoTo vbwErrHandler
3823       Const VBWPROCNAME = "frmPLCData.EnablePLCFields"
3824       If vbwProtector.vbwTraceProc Then
3825           Dim vbwProtectorParameterString As String
3826           If vbwProtector.vbwTraceParameters Then
3827               vbwProtectorParameterString = "()"
3828           End If
3829           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3830       End If
' </VB WATCH>
3831       frmPLCData.txtAI1Display.Enabled = True
3832       frmPLCData.txtAI2Display.Enabled = True
3833       frmPLCData.txtAI3Display.Enabled = True
3834       frmPLCData.txtAI4Display.Enabled = True
3835       frmPLCData.txtTC1Display.Enabled = True
3836       frmPLCData.txtTC2Display.Enabled = True
3837       frmPLCData.txtTC3Display.Enabled = True
3838       frmPLCData.txtTC4Display.Enabled = True
3839       frmPLCData.txtFlowDisplay.Enabled = True
3840       frmPLCData.txtSuctionDisplay.Enabled = True
3841       frmPLCData.txtDischargeDisplay.Enabled = True
3842       frmPLCData.txtTemperatureDisplay.Enabled = True
3843       frmPLCData.txtInHgDisplay.Enabled = True
' <VB WATCH>
3844       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3845       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "EnablePLCFields"

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
Private Sub DisablePLCFields()
' <VB WATCH>
3846       On Error GoTo vbwErrHandler
3847       Const VBWPROCNAME = "frmPLCData.DisablePLCFields"
3848       If vbwProtector.vbwTraceProc Then
3849           Dim vbwProtectorParameterString As String
3850           If vbwProtector.vbwTraceParameters Then
3851               vbwProtectorParameterString = "()"
3852           End If
3853           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3854       End If
' </VB WATCH>
3855       frmPLCData.txtAI1Display.Enabled = False
3856       frmPLCData.txtAI2Display.Enabled = False
3857       frmPLCData.txtAI3Display.Enabled = False
3858       frmPLCData.txtAI4Display.Enabled = False
3859       frmPLCData.txtTC1Display.Enabled = False
3860       frmPLCData.txtTC2Display.Enabled = False
3861       frmPLCData.txtTC3Display.Enabled = False
3862       frmPLCData.txtTC4Display.Enabled = False
3863       frmPLCData.txtFlowDisplay.Enabled = False
3864       frmPLCData.txtSuctionDisplay.Enabled = False
3865       frmPLCData.txtDischargeDisplay.Enabled = False
3866       frmPLCData.txtTemperatureDisplay.Enabled = False
3867       frmPLCData.txtInHgDisplay.Enabled = False
' <VB WATCH>
3868       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3869       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DisablePLCFields"

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
Private Sub BlankData()
' <VB WATCH>
3870       On Error GoTo vbwErrHandler
3871       Const VBWPROCNAME = "frmPLCData.BlankData"
3872       If vbwProtector.vbwTraceProc Then
3873           Dim vbwProtectorParameterString As String
3874           If vbwProtector.vbwTraceParameters Then
3875               vbwProtectorParameterString = "()"
3876           End If
3877           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3878       End If
' </VB WATCH>
3879       txtShpNo.Text = vbNullString
3880       txtBilNo.Text = vbNullString
3881       txtModelNo.Text = vbNullString
3882       cmbMotor.ListIndex = -1
3883       cmbStatorFill.ListIndex = -1
3884       cmbVoltage.ListIndex = -1
3885       cmbDesignPressure.ListIndex = -1
3886       cmbFrequency.ListIndex = -1
3887       cmbCirculationPath.ListIndex = -1
3888       cmbRPM.ListIndex = -1
3889       cmbModel.ListIndex = -1
3890       cmbModelGroup.ListIndex = -1
3891       txtSpGr.Text = vbNullString
3892       txtImpellerDia.Text = vbNullString
3893       txtEndPlay.Text = vbNullString
3894       txtGGap.Text = vbNullString
3895       txtDesignFlow.Text = vbNullString
3896       txtDesignTDH.Text = vbNullString
3897       txtOtherMods.Text = vbNullString
3898       txtRemarks.Text = vbNullString
3899       txtSalesOrderNumber.Text = vbNullString
3900       txtTestSetupRemarks.Text = vbNullString
3901       txtNPSHFile.Text = vbNullString
3902       txtPicturesFile.Text = vbNullString
3903       txtVibrationFile.Text = vbNullString
       '    cmbOrificeNumber.ListIndex = 18

3904       SetFrequencyCombo

       '    cmbTestSpec.ListIndex = 6       'default = Rev7
3905       cmbLoopNumber.ListIndex = -1
3906       cmbSuctDia.ListIndex = -1
3907       cmbDischDia.ListIndex = -1
3908       cmbTachID.ListIndex = -1
3909       cmbAnalyzerNo.ListIndex = -1
3910       txtTestRemarks.Text = vbNullString
3911       txtHDCor.Text = 0
3912       txtDischHeight.Text = 0
3913       txtSuctHeight.Text = 0
3914       txtKWMult.Text = 1
3915       txtWho.Text = LogInInitials
3916       txtRMA.Text = vbNullString
3917       frmPLCData.chkNPSH.value = 0
3918       frmPLCData.chkPictures.value = 0
3919       frmPLCData.chkVibration.value = 0
3920       cmbFlowMeter.ListIndex = -1
3921       cmbSuctionPressureTransducer.ListIndex = -1
3922       cmbDischargePressureTransducer.ListIndex = -1
3923       cmbTemperatureTransducer.ListIndex = -1
3924       cmbCirculationFlowMeter.ListIndex = -1
3925       frmPLCData.chkBalanceHoles.value = 0
3926       frmPLCData.chkCircOrifice.value = 0
3927       frmPLCData.txtCircOrifice = vbNullString
3928       frmPLCData.txtImpTrim = vbNullString
3929       frmPLCData.txtOrifice = vbNullString
3930       frmPLCData.chkFeathered.value = Unchecked
3931       frmPLCData.chkTrimmed.value = 0
3932       frmPLCData.chkCircOrifice.value = 0
3933       frmPLCData.txtThrustBal = vbNullString
3934       frmPLCData.txtRPM = vbNullString
3935       frmPLCData.txtVibAx = vbNullString
3936       frmPLCData.txtVibRad = vbNullString
3937       frmPLCData.txtTEMCTRGReading = vbNullString
3938       dgBalanceHoles.Visible = False
3939       Me.txtLineNumber.Text = vbNullString
3940       Me.txtNPSHr.Text = vbNullString
3941       Me.txtRatedInputPower.Text = vbNullString
3942       Me.txtAmps.Text = vbNullString
3943       Me.txtThermalClass.Text = vbNullString
3944       Me.txtViscosity.Text = vbNullString
3945       Me.txtTEMCViscosity.Text = vbNullString
3946       Me.txtExpClass.Text = vbNullString
3947       Me.txtNoPhases.Text = vbNullString
3948       Me.txtLiquidTemperature.Text = vbNullString
3949       Me.txtJobNum.Text = vbNullString
3950       Me.txtTEMCFrameNumber.Text = vbNullString
3951       Me.txtLiquid.Text = vbNullString
3952       Me.chkSuperMarketFeathered.value = Unchecked
3953       Me.txtRVSPartNo.Text = vbNullString
' <VB WATCH>
3954       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3955       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "BlankData"

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
Private Sub AddTestData()
' <VB WATCH>
3956       On Error GoTo vbwErrHandler
3957       Const VBWPROCNAME = "frmPLCData.AddTestData"
3958       If vbwProtector.vbwTraceProc Then
3959           Dim vbwProtectorParameterString As String
3960           If vbwProtector.vbwTraceParameters Then
3961               vbwProtectorParameterString = "()"
3962           End If
3963           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3964       End If
' </VB WATCH>
3965       Dim I As Integer
3966       Dim sFilter As String

3967       ClearEff
3968       rsEff.MoveFirst

3969       For I = 1 To 8
3970           rsTestData.AddNew
3971           rsTestData!SerialNumber = txtSN
3972           rsTestData!Date = cmbTestDate.List(cmbTestDate.ListIndex)
3973           rsTestData!testnumber = I
3974           rsTestData!DataWritten = False
3975           rsTestData.Update
3976           DoEfficiencyCalcs
3977           rsEff.MoveNext
3978           rsTestData.MoveNext
3979       Next I
3980       boFoundTestData = True
           'rsTestData.Update
3981       rsTestData.Requery
3982       rsTestData.Resync

          'select the entries from testdata
3983       sFilter = "SerialNumber='" & txtSN.Text & "' AND Date=#" & cmbTestDate.Text & "#"

3984       rsTestData.Filter = sFilter

3985       Set DataGrid1.DataSource = rsTestData

           ' fix the datagrid

3986       Dim c As Column
3987       For Each c In DataGrid1.Columns
3988          Select Case c.DataField
              Case "TestDataID"
3989             c.Visible = False
3990          Case "SerialNumber"
3991             c.Visible = False
3992          Case "Date"
3993             c.Visible = False
3994          Case Else ' Hide all other columns.
3995             c.Visible = True
3996             c.Alignment = dbgRight
3997          End Select
3998       Next c

3999       rsEff.Requery
4000       DataGrid1.Refresh
4001       DataGrid2.Refresh

' <VB WATCH>
4002       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4003       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "AddTestData"

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
            vbwReportVariable "I", I
            vbwReportVariable "sFilter", sFilter
            vbwReportVariable "c", c
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub DoEfficiencyCalcs()
' <VB WATCH>
4004       On Error GoTo vbwErrHandler
4005       Const VBWPROCNAME = "frmPLCData.DoEfficiencyCalcs"
4006       If vbwProtector.vbwTraceProc Then
4007           Dim vbwProtectorParameterString As String
4008           If vbwProtector.vbwTraceParameters Then
4009               vbwProtectorParameterString = "()"
4010           End If
4011           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4012       End If
' </VB WATCH>
4013       Dim KW As Single, VI As Single, VITemp As Single
4014       Dim Vave As Single, Iave As Single
4015       Dim I As Integer
4016       Dim j As Integer
4017       Dim HeightDiff As Single

4018       If Not IsNull(rsTestData.Fields("TotalPower")) Then
4019           KW = rsTestData.Fields("TotalPower")
4020       Else
               'if we wrote data with an old version, we will not have written total power
               'if total power = 0 and the three individual powers are not 0, add them

4021           If rsTestData.Fields("PowerA") > 0 Then
4022               If rsTestData.Fields("PowerB") > 0 Then
4023                   If rsTestData.Fields("PowerC") > 0 Then
4024                       KW = rsTestData.Fields("PowerA") + rsTestData.Fields("PowerB") + rsTestData.Fields("PowerC")
4025                   End If
4026               End If
4027           End If
4028      End If

4029       I = 0
4030       Vave = 0
4031       Iave = 0
4032       If Not IsNull(rsTestData.Fields("VoltageA")) And Not IsNull(rsTestData.Fields("CurrentA")) Then
4033           VI = rsTestData.Fields("VoltageA") * rsTestData.Fields("CurrentA")
4034           Vave = rsTestData.Fields("VoltageA")
4035           Iave = rsTestData.Fields("CurrentA")
4036           If VI <> 0 Then
4037               I = I + 1
4038           End If
4039       End If
4040       If Not IsNull(rsTestData.Fields("VoltageB")) And Not IsNull(rsTestData.Fields("CurrentB")) Then
4041           VITemp = rsTestData.Fields("VoltageB") * rsTestData.Fields("CurrentB")
4042           If VITemp <> 0 Then
4043               I = I + 1
4044               VI = VI + VITemp
4045               Vave = Vave + rsTestData.Fields("VoltageB")
4046               Iave = Iave + rsTestData.Fields("CurrentB")
4047           End If
4048       End If
4049       If Not IsNull(rsTestData.Fields("VoltageC")) And Not IsNull(rsTestData.Fields("CurrentC")) Then
4050           VITemp = rsTestData.Fields("VoltageC") * rsTestData.Fields("CurrentC")
4051           If VITemp <> 0 Then
4052               I = I + 1
4053               VI = VI + VITemp
4054               Vave = Vave + rsTestData.Fields("VoltageC")
4055               Iave = Iave + rsTestData.Fields("CurrentC")
4056           End If
4057       End If
4058       If KW = 0 Then
4059           For j = 1 To rsEff.Fields.Count - 1
4060               rsEff.Fields(j) = 0
4061           Next j
       '        Exit Sub
4062       End If
4063       If VI <> 0 Then
4064           rsEff.Fields("Volts") = Vave / I
4065           rsEff.Fields("Amps") = Iave / I
4066           rsEff.Fields("PowerFactor") = 1000 * I * KW / (VI * Sqr(3))
4067           rsEff.Fields("PowerFactor") = 100 * rsEff.Fields("PowerFactor")
4068       Else
4069           rsEff.Fields("PowerFactor") = 0
4070       End If

4071       If optMfr(0).value = True Then
4072           If cmbStatorFill.ListIndex = -1 Then
4073               rsEff.Fields("MotorEfficiency") = Format$(0, "0.00")

4074           Else
4075               rsEff.Fields("Motorefficiency") = Format$(Round(MotorEfficiency(KW, cmbMotor.ItemData(cmbMotor.ListIndex), cmbStatorFill.ItemData(cmbStatorFill.ListIndex)), 1), "00.0")
       '            rsEff.Fields("Motorefficiency") = Format$(Round(MotorEfficiency(KW, cmbMotor.ListIndex, cmbStatorFill.ListIndex), 1), "00.0")
4076           End If
4077       Else
4078           rsEff.Fields("MotorEfficiency") = Format$(Round(TEMCMotorEfficiency(KW, txtTEMCFrameNumber.Text, 460, RatedKW), 1), "00.0")
4079       End If

4080       Dim sHDCor As Single
4081       Dim sDisc As Single
4082       Dim sSuct As Single
4083       If IsNull(rsTestSetup.Fields("HDCor")) Then
4084           sHDCor = 0
4085       Else
4086           sHDCor = rsTestSetup.Fields("HDCor")
4087       End If
4088       If IsNull(rsTestSetup.Fields("DischargeGageHeight")) Then
4089           sDisc = 0
4090       Else
4091           sDisc = rsTestSetup.Fields("DischargeGageHeight")
4092       End If
4093       If IsNull(rsTestSetup.Fields("SuctionGageHeight")) Then
4094           sSuct = 0
4095       Else
4096           sSuct = rsTestSetup.Fields("SuctionGageHeight")
4097       End If
4098       HeightDiff = sHDCor + sDisc / 12 - sSuct / 12
4099       If (cmbDischDia.ListIndex <> -1 And cmbSuctDia.ListIndex <> -1) Then
4100           rsEff.Fields("VelocityHead") = CalcVelHead(rsTestData.Fields("Flow"), cmbDischDia.ItemData(cmbDischDia.ListIndex) + 1, cmbSuctDia.ItemData(cmbSuctDia.ListIndex) + 1)
4101       End If
       '    rsEff.Fields("VelocityHead") = CalcVelHead(rsTestData.Fields("Flow"), cmbDischDia.ListIndex + 1, cmbSuctDia.ListIndex + 1)
4102       rsEff.Fields("TDH") = CalcTDH(rsTestData.Fields("DischargePressure"), rsTestData.Fields("SuctionPressure"), rsTestData.Fields("SuctionInHg"), rsEff.Fields("VelocityHead"), HeightDiff, rsTestData.Fields("TemperatureSuction"))
4103       rsEff.Fields("ElecHP") = 1000 * KW / 746
       '    If (DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))) <> 0 And KW <> 0) Then
4104           If Int(rsTestData.Fields("TemperatureSuction")) >= 40 Then
4105               If (DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))) <> 0 And KW <> 0) Then
           '        rsEff.Fields("LiquidHP") = (rsEff.Fields("TDH") * rsTestData.Fields("Flow") * DLookup("TDHCorr", "TempCorrection", "Temp = 68")) / (3960 * DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))))
4106               rsEff.Fields("LiquidHP") = (rsEff.Fields("TDH") * rsTestData.Fields("Flow") * DLookupA(TDHColNo, TempCorrection, TempColNo, 68)) / (3960 * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))))
           '        rsEff.Fields("OverallEfficiency") = (0.189 * rsTestData.Fields("Flow") * rsEff.Fields("TDH") * DLookup("TDHCorr", "TempCorrection", "Temp = 68")) / (10 * KW * DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))))
4107               rsEff.Fields("OverallEfficiency") = (0.189 * rsTestData.Fields("Flow") * rsEff.Fields("TDH") * DLookupA(TDHColNo, TempCorrection, TempColNo, 68)) / (10 * KW * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))))
4108               If rsEff.Fields("MotorEfficiency") <> 0 Then
4109                   rsEff.Fields("HydraulicEfficiency") = 100 * rsEff.Fields("OverallEfficiency") / rsEff.Fields("MotorEfficiency")
4110               Else
4111                   rsEff.Fields("HydraulicEfficiency") = 0
4112               End If
4113           Else
4114               rsEff.Fields("LiquidHP") = 0
4115               rsEff.Fields("OverallEfficiency") = 0
4116           End If

4117       Else
4118           rsEff.Fields("LiquidHP") = 0
4119           rsEff.Fields("OverallEfficiency") = 0
4120       End If


4121       I = rsEff.AbsolutePosition
4122       If Not IsNull(rsTestData.Fields("Flow")) Then
4123           rsEff.Fields("Flow") = rsTestData.Fields("Flow")
4124           HeadFlow(0, I - 1) = rsTestData.Fields("Flow")
4125           HeadFlow(1, I - 1) = rsEff.Fields("TDH")
4126           FlowHead(I - 1, 0) = rsTestData.Fields("Flow")
4127           FlowHead(I - 1, 1) = rsEff.Fields("TDH")

       '        EffFlow(0, i - 1) = rsTestData.Fields("Flow")
       '        EffFlow(1, i - 1) = rsEff.Fields("OverallEfficiency")
       '        KWFlow(0, i - 1) = rsTestData.Fields("Flow")
       '        KWFlow(1, i - 1) = KW
       '        AmpsFlow(0, i - 1) = rsTestData.Fields("Flow")
       '        AmpsFlow(1, i - 1) = rsEff.Fields("Amps")
4128       Else
4129           HeadFlow(0, I - 1) = 0
4130           HeadFlow(1, I - 1) = 0
4131           FlowHead(I - 1, 0) = 0
4132           FlowHead(I - 1, 1) = 0

       '        EffFlow(0, i - 1) = 0
       '        EffFlow(1, i - 1) = 0
       '        KWFlow(0, i - 1) = 0
       '        KWFlow(1, i - 1) = 0
       '        AmpsFlow(0, i - 1) = 0
       '        AmpsFlow(1, i - 1) = 0
4133       End If

4134       Dim Plothead(1, 7) As Single
4135       Dim HeadPlot(7, 1) As Single
           'ReDim Preserve Plothead(1, j)
           'ReDim Preserve HeadPlot(j, 1)

       '    Dim PlotEff() As Single
       '    Dim PlotKW() As Single
       '    Dim PlotAmps() As Single
       '    ReDim PlotHead(0, 0)
       '    ReDim PlotEff(0, 0)
       '    ReDim PlotKW(0, 0)
       '
4136       For j = 0 To UpDown2.value - 1
       '        If HeadFlow(1, j) <> 0 Then
       '            ReDim Preserve Plothead(1, j)
       '            ReDim Preserve HeadPlot(j, 1)
4137               Plothead(0, j) = HeadFlow(0, j)
4138               Plothead(1, j) = HeadFlow(1, j)
4139               HeadPlot(j, 0) = FlowHead(j, 0)
4140               HeadPlot(j, 1) = FlowHead(j, 1)
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
4141       Next j




       '    SetGraphMax (Plothead())
       '    If UBound(PlotHead()) <> 0 Then

       'fix 4/29/19

4142           MSChart1.ChartData = HeadPlot

       '    End If

           'copy fields for reports
4143       rsEff.Fields("DischPress") = rsTestData.Fields("Dischargepressure")
4144       rsEff.Fields("SuctPress") = rsTestData.Fields("Suctionpressure")
       '    rsEff.Fields("Volts") = rsTestData.Fields("VoltageA")
       '    rsEff.Fields("Amps") = rsTestData.Fields("CurrentA")
4145       rsEff.Fields("KW") = KW
4146       rsEff.Fields("Freq") = rsTestData.Fields("VFDFrequency")
4147       rsEff.Fields("RPM") = rsTestData.Fields("RPM")
4148       rsEff.Fields("Pos") = rsTestData.Fields("ThrustBalance")
4149       rsEff.Fields("NPSHa") = rsTestData.Fields("NPSHa")
4150       rsEff.Fields("NPSHr") = rsTestData.Fields("NPSHr")
4151       rsEff.Fields("InputPower") = rsTestData.Fields("TotalPower")
4152       rsEff.Fields("Temperature") = rsTestData.Fields("TemperatureSuction")
4153       rsEff.Fields("CircFlow") = rsTestData.Fields("CircFlow")
4154       rsEff.Fields("VibrationX") = rsTestData.Fields("VibrationX")
4155       rsEff.Fields("VibrationY") = rsTestData.Fields("VibrationY")
4156       rsEff.Fields("CurrentA") = rsTestData.Fields("CurrentA")
4157       rsEff.Fields("CurrentB") = rsTestData.Fields("CurrentB")
4158       rsEff.Fields("CurrentC") = rsTestData.Fields("CurrentC")
4159       rsEff.Fields("VoltageA") = rsTestData.Fields("VoltageA")
4160       rsEff.Fields("VoltageB") = rsTestData.Fields("VoltageB")
4161       rsEff.Fields("VoltageC") = rsTestData.Fields("VoltageC")
4162       rsEff.Fields("TC1") = rsTestData.Fields("TC1")
4163       rsEff.Fields("TC2") = rsTestData.Fields("TC2")
4164       rsEff.Fields("TC3") = rsTestData.Fields("TC3")
4165       rsEff.Fields("TC4") = rsTestData.Fields("TC4")
4166       rsEff.Fields("RBHTemp") = rsTestData.Fields("RBHTemp")
4167       rsEff.Fields("RBHPress") = rsTestData.Fields("RBHPress")
4168       rsEff.Fields("AI4") = rsTestData.Fields("AI4")
4169       rsEff.Fields("Remarks") = rsTestData.Fields("Remarks")
4170       rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCFrontThrust")
4171       rsEff.Fields("TEMCRearThrust") = rsTestData.Fields("TEMCRearThrust")
4172       rsEff.Fields("TEMCTRG") = rsTestData.Fields("TEMCTRG")
4173       rsEff.Fields("TEMCThrustRigPressure") = rsTestData.Fields("TEMCThrustRigPressure")
4174       rsEff.Fields("TEMCMomentArm") = rsTestData.Fields("TEMCMomentArm")
4175       rsEff.Fields("TEMCViscosity") = rsTestData.Fields("TEMCViscosity")
4176       If Not IsNull(rsEff.Fields("TEMCFrontThrust")) Then
4177           txtTEMCFrontThrust.Text = rsEff.Fields("TEMCFrontThrust")
4178       End If
4179       If Not IsNull(rsEff.Fields("TEMCREarThrust")) Then
4180           txtTEMCRearThrust.Text = rsEff.Fields("TEMCREarThrust")
4181       End If
4182       If (Not IsNull(rsEff.Fields("TEMCViscosity"))) And (rsEff.Fields("TEMCViscosity") <> 0) Then
4183           txtTEMCViscosity.Text = rsEff.Fields("TEMCViscosity")
4184       End If
4185       If Not IsNull(rsTestData.Fields("TEMCThrustRigPressure")) Then
4186           txtTEMCThrustRigPressure.Text = rsTestData.Fields("TEMCThrustRigPressure")
4187       End If
4188       If Not IsNull(rsTestData.Fields("TEMCMomentArm")) Then
4189           txtTEMCMomentArm.Text = rsTestData.Fields("TEMCMomentArm")
4190       End If

        '   If Not IsNull(Me.txtAI3Display.Text) Then
        '       Me.txtAI3Display = rsTestData.Fields("RBHPress")
        '   End If

4191       CalculateTEMCForce

4192       If Not IsNull(txtTEMCCalcForce.Text) Then
4193           rsEff.Fields("TEMCCalculatedForce") = Val(txtTEMCCalcForce.Text)
4194       Else
4195           rsEff.Fields("TEMCCalculatedForce") = 0
4196       End If

4197       If Not IsNull(txtTEMCPVValue.Text) Then
4198           rsEff.Fields("TEMCPV") = Val(txtTEMCPVValue.Text)
4199       Else
4200           rsEff.Fields("TEMCPV") = 0
4201       End If

4202       If Val(txtTEMCFrontThrust.Text) <> 0 Then
4203           rsEff.Fields("TEMCFR") = "F"
       '        rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCFrontThrust")
4204       Else
4205           If Val(txtTEMCRearThrust.Text) = 0 Then
                   'no thrust
4206               rsEff.Fields("TEMCFR") = " "
4207               rsEff.Fields("TEMCFrontThrust") = 0
4208           Else
4209               rsEff.Fields("TEMCFR") = "R"
       '            rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCRearThrust")
4210           End If
4211       End If

4212       rsEff.Fields("TEMCForceDirection") = Left(lblTEMCFrontRear.Caption, 1)

4213       rsEff.Update
4214       DataGrid2.Refresh


' <VB WATCH>
4215       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4216       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "DoEfficiencyCalcs"

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
            vbwReportVariable "KW", KW
            vbwReportVariable "VI", VI
            vbwReportVariable "VITemp", VITemp
            vbwReportVariable "Vave", Vave
            vbwReportVariable "Iave", Iave
            vbwReportVariable "I", I
            vbwReportVariable "j", j
            vbwReportVariable "HeightDiff", HeightDiff
            vbwReportVariable "sHDCor", sHDCor
            vbwReportVariable "sDisc", sDisc
            vbwReportVariable "sSuct", sSuct
            vbwReportVariable "Plothead", Plothead
            vbwReportVariable "HeadPlot", HeadPlot
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Private Sub ClearEff()
       '    Dim I As Integer, j As Integer
' <VB WATCH>
4217       On Error GoTo vbwErrHandler
4218       Const VBWPROCNAME = "frmPLCData.ClearEff"
4219       If vbwProtector.vbwTraceProc Then
4220           Dim vbwProtectorParameterString As String
4221           If vbwProtector.vbwTraceParameters Then
4222               vbwProtectorParameterString = "()"
4223           End If
4224           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4225       End If
' </VB WATCH>
4226       Dim qy As New ADODB.Command

4227       If rsEff.State = adStateOpen Then
4228           If Not (rsEff.BOF = True Or rsEff.EOF = True) Then
4229               rsEff.CancelUpdate
4230           End If
4231           rsEff.Close
4232       End If
4233       qy.ActiveConnection = cnEffData
4234       qy.CommandText = "DROP TABLE Efficiency"
4235       rsEff.Open qy
4236       qy.CommandText = "SELECT EfficiencyOrg.* INTO Efficiency FROM EfficiencyOrg;"
4237       rsEff.Open qy
4238       rsEff.Open "Efficiency", cnEffData, adOpenStatic, adLockOptimistic, adCmdTableDirect

4239       rsEff.Requery
4240       DataGrid2.Refresh

4241       Dim c As Column
4242       For Each c In DataGrid2.Columns
4243           c.Alignment = dbgCenter
4244           c.Width = 750
4245           Select Case c.ColIndex
                   Case 1
4246                   c.Caption = "Flow"
4247                   c.NumberFormat = "###0.00"
4248               Case 2
4249                   c.Caption = "TDH"
4250                   c.NumberFormat = "00.0"
4251               Case 3
4252                   c.Caption = "Overall Eff"
4253                   c.NumberFormat = "00.00"
4254                   c.Width = 850
4255               Case 4
4256                   c.Caption = "PF"
4257                   c.NumberFormat = "00.0"
4258               Case 5
4259                   c.Caption = "Vel Head"
4260                   c.NumberFormat = "00.00"
4261               Case 6
4262                   c.Caption = "Elec HP"
4263                   c.NumberFormat = "#00.0"
4264               Case 7
4265                   c.Caption = "Liq HP"
4266                   c.NumberFormat = "#00.0"
4267               Case Else
4268                   c.Visible = False
4269           End Select
4270       Next c

' <VB WATCH>
4271       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4272       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ClearEff"

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
            vbwReportVariable "qy", qy
            vbwReportVariable "c", c
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub
Function JustAlphaNumeric(char As String) As String
' <VB WATCH>
4273       On Error GoTo vbwErrHandler
4274       Const VBWPROCNAME = "frmPLCData.JustAlphaNumeric"
4275       If vbwProtector.vbwTraceProc Then
4276           Dim vbwProtectorParameterString As String
4277           If vbwProtector.vbwTraceParameters Then
4278               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("char", char) & ") "
4279           End If
4280           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4281       End If
' </VB WATCH>
4282       Select Case Asc(char)
               Case 42             ' *
4283               JustAlphaNumeric = char
4284           Case 48 To 57       ' 0 - 9
4285               JustAlphaNumeric = char
4286           Case 65 To 90       ' A - Z
4287               JustAlphaNumeric = char
4288           Case 97 To 122      ' a - z
4289               JustAlphaNumeric = UCase(char)
4290           Case Else
4291               JustAlphaNumeric = ""
4292       End Select
' <VB WATCH>
4293       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4294       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "JustAlphaNumeric"

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
            vbwReportVariable "char", char
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function



Private Sub txtI1_Change()
' <VB WATCH>
4295       On Error GoTo vbwErrHandler
4296       Const VBWPROCNAME = "frmPLCData.txtI1_Change"
4297       If vbwProtector.vbwTraceProc Then
4298           Dim vbwProtectorParameterString As String
4299           If vbwProtector.vbwTraceParameters Then
4300               vbwProtectorParameterString = "()"
4301           End If
4302           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4303       End If
' </VB WATCH>
4304       txtI2.Text = txtI1.Text
4305       txtI3.Text = txtI1.Text
' <VB WATCH>
4306       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4307       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtI1_Change"

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

Private Sub txtModelNo_Change()
' <VB WATCH>
4308       On Error GoTo vbwErrHandler
4309       Const VBWPROCNAME = "frmPLCData.txtModelNo_Change"
4310       If vbwProtector.vbwTraceProc Then
4311           Dim vbwProtectorParameterString As String
4312           If vbwProtector.vbwTraceParameters Then
4313               vbwProtectorParameterString = "()"
4314           End If
4315           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4316       End If
' </VB WATCH>
4317       Dim I As Integer
4318       Dim S As String
4319       Dim sFull As String
4320       Dim boDone As Boolean
4321       Dim boRepeat As Boolean

4322       Static bo3Digits As Boolean         '3 digits in frame number
4323       Static bo2Digits As Boolean         '2 digits in stages

4324       If optMfr(0).value = True Then
' <VB WATCH>
4325       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4326           Exit Sub
4327       End If

4328       cmbTEMCAdapter.ListIndex = -1
4329       cmbTEMCAdditions.ListIndex = -1
4330       cmbTEMCCirculation.ListIndex = -1
4331       cmbTEMCDesignPressure.ListIndex = -1
4332       cmbTEMCNominalDischargeSize.ListIndex = -1
4333       cmbTEMCDivisionType.ListIndex = -1
4334       cmbTEMCImpellerType.ListIndex = -1
4335       cmbTEMCInsulation.ListIndex = -1
4336       cmbTEMCJacketGasket.ListIndex = -1
4337       cmbTEMCMaterials.ListIndex = -1
4338       cmbTEMCModel.ListIndex = -1
4339       cmbTEMCNominalImpSize.ListIndex = -1
4340       cmbTEMCOtherMotor.ListIndex = -1
4341       cmbTEMCPumpStages.ListIndex = -1
4342       cmbTEMCNominalSuctionSize.ListIndex = -1
4343       cmbTEMCTRG.ListIndex = -1
4344       cmbTEMCVoltage.ListIndex = -1


           'first, get rid of spaces, dashes, etc

4345       S = ""
4346       For I = 1 To Len(txtModelNo.Text)
4347           S = S & JustAlphaNumeric(Mid$(txtModelNo.Text, I, 1))
4348       Next I

           'next, fill out the model number to it's max length of 24 characters

4349       boDone = False
4350       boRepeat = False

4351       Do While Not boDone
4352           sFull = ""
4353           For I = 1 To Len(S)
4354               Select Case I
                       Case 1
                           'type
4355                       sFull = sFull & Mid$(S, I, 1)
4356                   Case 2
                           'adapter
4357                       If IsNumeric(Mid$(S, I, 1)) Then
4358                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4359                           boRepeat = True
4360                           Exit For
4361                       Else
4362                           sFull = sFull & Mid$(S, I, 1)
4363                           boRepeat = False
4364                       End If
4365                   Case 3
                           'materials
4366                       sFull = sFull & Mid$(S, I, 1)
4367                   Case 4
                       'design pressure
4368                       sFull = sFull & Mid$(S, I, 1)
4369                   Case 5
                       'motor frame number - digit 1
4370                       sFull = sFull & Mid$(S, I, 1)
4371                   Case 6
                       'motor frame number - digit 2
4372                       sFull = sFull & Mid$(S, I, 1)
4373                   Case 7
                       'motor frame number - digit 3
4374                       sFull = sFull & Mid$(S, I, 1)
4375                   Case 8
                       'motor frame number - digit 4
4376                       If IsNumeric(Mid$(S, I, 1)) Then
4377                           sFull = sFull & Mid$(S, I, 1)
4378                           boRepeat = False
4379                       Else    '3 digits
       '                        s = Left$(s, i - 1) & "*" & Right$(s, Len(s) - i + 1)
4380                           S = Left$(S, I - 4) & "0" & Right$(S, Len(S) - I + 4)
4381                           boRepeat = True
4382                           Exit For
4383                       End If
4384                   Case 9
                       'insulation
4385                       sFull = sFull & Mid$(S, I, 1)
4386                   Case 10
                       'voltage
4387                       sFull = sFull & Mid$(S, I, 1)
4388                   Case 11
                       'other motor specs
4389                       If Mid$(S, I, 1) = "M" Or Mid$(S, I, 1) = "R" Or Mid$(S, I, 1) = "L" Or Mid$(S, I, 1) = "G" Or Mid$(S, I, 1) = "N" Then
4390                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4391                           boRepeat = True
4392                           Exit For
4393                       Else
4394                           sFull = sFull & Mid$(S, I, 1)
4395                           boRepeat = False
4396                       End If
4397                   Case 12
                       ' TRG
4398                       sFull = sFull & Mid$(S, I, 1)
4399                   Case 13
                       'Nominal discharge - digit 1
4400                       sFull = sFull & Mid$(S, I, 1)
4401                   Case 14
                       'nominal discharge - digit 2
4402                       sFull = sFull & Mid$(S, I, 1)
4403                   Case 15
                       'nominal suction - digit 1
4404                       sFull = sFull & Mid$(S, I, 1)
4405                   Case 16
                       'nominal suction - digit 2
4406                       sFull = sFull & Mid$(S, I, 1)
4407                   Case 17
                       'nominal impeller size
4408                       sFull = sFull & Mid$(S, I, 1)
4409                   Case 18
                       'impeller type
4410                       If Mid$(S, I, 1) <> "*" Then
4411                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4412                           boRepeat = True
4413                           Exit For
4414                       Else
4415                           sFull = sFull & Mid$(S, I, 1)
4416                           boRepeat = False
4417                       End If
4418                   Case 19
                       'Division type
4419                       If IsNumeric(Mid$(S, I, 1)) Then
4420                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4421                           boRepeat = True
4422                           Exit For
4423                       Else
4424                           sFull = sFull & Mid$(S, I, 1)
4425                           boRepeat = False
4426                       End If
4427                   Case 20
                       'pump stages - digit 1
4428                       sFull = sFull & Mid$(S, I, 1)
4429                   Case 21
                       'pump jacket
4430                       If Mid$(S, I, 1) = "A" Or Mid$(S, I, 1) = "B" Or Mid$(S, I, 1) = "E" Or Mid$(S, I, 1) = "F" Or _
                                             Mid$(S, I, 1) = "G" Or Mid$(S, I, 1) = "H" Or Mid$(S, I, 1) = "J" Or Mid$(S, I, 1) = "K" Then
4431                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4432                           boRepeat = True
4433                       Else
4434                           sFull = sFull & Mid$(S, I, 1)
4435                           boRepeat = False
4436                       End If
4437                   Case 22
                       'additions
4438                         sFull = sFull & Mid$(S, I, 1)
4439                   Case 23
                       'circulation
4440                         sFull = sFull & Mid$(S, I, 1)
4441               End Select
4442           Next I
4443           If Not boRepeat Then
4444               boDone = True
4445           End If
4446       Loop

4447       For I = 1 To Len(sFull)
4448           Select Case I
                   Case 1
4449                   ParseTEMCModelNo cmbTEMCModel, Mid$(sFull, I, 1)
4450               Case 2
4451                   ParseTEMCModelNo cmbTEMCAdapter, Mid$(sFull, I, 1)
4452               Case 3
4453                   ParseTEMCModelNo cmbTEMCMaterials, Mid$(sFull, I, 1)
4454               Case 4
4455                   ParseTEMCModelNo cmbTEMCDesignPressure, Mid$(sFull, I, 1)
4456               Case 5
4457                       If Val(Mid$(sFull, I, 1)) = 0 Then
4458                           txtTEMCFrameNumber.Text = Mid$(sFull, 6, 3)
4459                       Else
4460                           txtTEMCFrameNumber.Text = Mid$(sFull, 5, 4)
4461                       End If
4462               Case 9
4463                       ParseTEMCModelNo cmbTEMCInsulation, Mid$(sFull, I, 1)
4464               Case 10
4465                       ParseTEMCModelNo cmbTEMCVoltage, Mid$(sFull, I, 1)
4466               Case 11
4467                       ParseTEMCModelNo cmbTEMCOtherMotor, Mid$(sFull, I, 1)
4468               Case 12
4469                       ParseTEMCModelNo cmbTEMCTRG, Mid$(sFull, I, 1)
4470               Case 13
4471                       ParseTEMCModelNo cmbTEMCNominalDischargeSize, Mid$(sFull, I, 2)
4472               Case 14
4473               Case 15
4474                       ParseTEMCModelNo cmbTEMCNominalSuctionSize, Mid$(sFull, I, 2)
4475               Case 16
4476               Case 17
4477                       ParseTEMCModelNo cmbTEMCNominalImpSize, Mid$(sFull, I, 1)
4478               Case 18
4479                       ParseTEMCModelNo cmbTEMCImpellerType, Mid$(sFull, I, 1)
4480               Case 19
4481                       ParseTEMCModelNo cmbTEMCDivisionType, Mid$(sFull, I, 1)
4482               Case 20
4483                       ParseTEMCModelNo cmbTEMCPumpStages, Mid$(sFull, I, 1)
4484               Case 21
4485                       ParseTEMCModelNo cmbTEMCJacketGasket, Mid$(sFull, I, 1)
4486               Case 22
4487                       ParseTEMCModelNo cmbTEMCAdditions, Mid$(sFull, I, 1)
4488                       ParseTEMCModelNo cmbTEMCCirculation, "*"
4489               Case 23
       '                    ParseTEMCModelNo cmbTEMCCirculation, Mid$(sFull, I, 1)

4490           End Select
4491       Next I

           'give alerts on certain conditions
4492       Dim msg As String
4493       msg = ""

           'look for 4 in third digit of model number for CO2 pump
4494       If Mid(txtModelNo.Text, 3, 1) = "4" Then
4495           msg = "Special CO2 model. Requires special thrust rig and circulation flow setting. See Engineering"
4496       End If

4497       If Left(cmbTEMCVoltage, 3) = "[6]" Then
4498           msg = "575V transformer required for Rundown and TRG"
4499       End If
       '    If Left(cmbTEMCTRG, 3) = "[L]" Or InStr("X[B][F][H][K]", Left(cmbTEMCAdditions, 3)) > 1 Then
4500       If Left(cmbTEMCTRG, 3) = "[L]" Then
4501           If msg = "" Then
4502               msg = "VFD required for Rundown and TRG"
4503           Else
4504               msg = msg & " and " & "VFD required for Rundown and TRG"
4505           End If
4506       End If

4507       If InStr("X[B][F][H][K]", Left(cmbTEMCAdditions, 3)) > 1 Then
4508           If msg = "" Then
4509               msg = "VFD required for Rundown, standard drive required for TRG"
4510           Else
4511               msg = msg & " and " & "VFD required for Rundown, standard drive required for TRG"
4512           End If
4513       End If

4514       If msg <> "" Then
4515           frmAlert.txtAlert.Text = msg
4516           frmAlert.Show
4517       End If

' <VB WATCH>
4518       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4519       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtModelNo_Change"

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
            vbwReportVariable "I", I
            vbwReportVariable "S", S
            vbwReportVariable "sFull", sFull
            vbwReportVariable "boDone", boDone
            vbwReportVariable "boRepeat", boRepeat
            vbwReportVariable "bo3Digits", bo3Digits
            vbwReportVariable "bo2Digits", bo2Digits
            vbwReportVariable "msg", msg
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub


Private Sub txtModelNo_Validate(Cancel As Boolean)
' <VB WATCH>
4520       On Error GoTo vbwErrHandler
4521       Const VBWPROCNAME = "frmPLCData.txtModelNo_Validate"
4522       If vbwProtector.vbwTraceProc Then
4523           Dim vbwProtectorParameterString As String
4524           If vbwProtector.vbwTraceParameters Then
4525               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Cancel", Cancel) & ") "
4526           End If
4527           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4528       End If
' </VB WATCH>
4529       Dim I As Integer
4530       Dim S As String

       '    s = txtModelNo.Text
       '    S = Replace(S, "-", "")
       '    S = Replace(S, " ", "")
       '    S = Replace(S, "/", "")

       '    txtModelNo.Text = ""

       '    For i = 1 To Len(s)
       '        txtModelNo.Text = txtModelNo.Text & Mid(s, i, 1)
       '    Next i
4531       txtModelNo_Change

' <VB WATCH>
4532       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4533       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtModelNo_Validate"

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
            vbwReportVariable "Cancel", Cancel
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
End Sub

Private Sub txtNPSHFile_GotFocus()
' <VB WATCH>
4534       On Error GoTo vbwErrHandler
4535       Const VBWPROCNAME = "frmPLCData.txtNPSHFile_GotFocus"
4536       If vbwProtector.vbwTraceProc Then
4537           Dim vbwProtectorParameterString As String
4538           If vbwProtector.vbwTraceParameters Then
4539               vbwProtectorParameterString = "()"
4540           End If
4541           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4542       End If
' </VB WATCH>
4543       On Error GoTo FileCancel
4544       If LenB(txtNPSHFile.Text) <> 0 Then
4545           CommonDialog1.filename = txtNPSHFile.Text
4546       End If
4547       CommonDialog1.ShowOpen
4548       txtNPSHFile.Text = CommonDialog1.filename
' <VB WATCH>
4549       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4550       Exit Sub
4551   FileCancel:
4552   On Error GoTo vbwErrHandler
4553       CommonDialog1.CancelError = False
' <VB WATCH>
4554       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4555       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtNPSHFile_GotFocus"

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

Private Sub txtP1_Change()
' <VB WATCH>
4556       On Error GoTo vbwErrHandler
4557       Const VBWPROCNAME = "frmPLCData.txtP1_Change"
4558       If vbwProtector.vbwTraceProc Then
4559           Dim vbwProtectorParameterString As String
4560           If vbwProtector.vbwTraceParameters Then
4561               vbwProtectorParameterString = "()"
4562           End If
4563           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4564       End If
' </VB WATCH>
4565       txtP2.Text = txtP1.Text
4566       txtP3.Text = txtP1.Text
' <VB WATCH>
4567       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4568       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtP1_Change"

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

Private Sub txtPicturesFile_gotfocus()
' <VB WATCH>
4569       On Error GoTo vbwErrHandler
4570       Const VBWPROCNAME = "frmPLCData.txtPicturesFile_gotfocus"
4571       If vbwProtector.vbwTraceProc Then
4572           Dim vbwProtectorParameterString As String
4573           If vbwProtector.vbwTraceParameters Then
4574               vbwProtectorParameterString = "()"
4575           End If
4576           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4577       End If
' </VB WATCH>
4578       CommonDialog1.CancelError = True
4579       On Error GoTo FileCancel
4580       If LenB(txtPicturesFile.Text) <> 0 Then
4581           CommonDialog1.filename = txtPicturesFile.Text
4582       End If
4583       CommonDialog1.ShowOpen
4584       txtPicturesFile.Text = CommonDialog1.filename
' <VB WATCH>
4585       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4586       Exit Sub
4587   FileCancel:
4588   On Error GoTo vbwErrHandler
4589       CommonDialog1.CancelError = False
' <VB WATCH>
4590       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4591       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtPicturesFile_gotfocus"

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

Private Sub txtSN_Change()
' <VB WATCH>
4592       On Error GoTo vbwErrHandler
4593       Const VBWPROCNAME = "frmPLCData.txtSN_Change"
4594       If vbwProtector.vbwTraceProc Then
4595           Dim vbwProtectorParameterString As String
4596           If vbwProtector.vbwTraceParameters Then
4597               vbwProtectorParameterString = "()"
4598           End If
4599           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4600       End If
' </VB WATCH>
4601       cmdFindPump.Default = True
' <VB WATCH>
4602       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4603       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtSN_Change"

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

Private Sub txtTEMCFrontThrust_Change()
' <VB WATCH>
4604       On Error GoTo vbwErrHandler
4605       Const VBWPROCNAME = "frmPLCData.txtTEMCFrontThrust_Change"
4606       If vbwProtector.vbwTraceProc Then
4607           Dim vbwProtectorParameterString As String
4608           If vbwProtector.vbwTraceParameters Then
4609               vbwProtectorParameterString = "()"
4610           End If
4611           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4612       End If
' </VB WATCH>
4613       CalculateTEMCForce
' <VB WATCH>
4614       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4615       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtTEMCFrontThrust_Change"

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

Private Sub txtTEMCMomentArm_Change()
' <VB WATCH>
4616       On Error GoTo vbwErrHandler
4617       Const VBWPROCNAME = "frmPLCData.txtTEMCMomentArm_Change"
4618       If vbwProtector.vbwTraceProc Then
4619           Dim vbwProtectorParameterString As String
4620           If vbwProtector.vbwTraceParameters Then
4621               vbwProtectorParameterString = "()"
4622           End If
4623           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4624       End If
' </VB WATCH>
4625       CalculateTEMCForce
' <VB WATCH>
4626       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4627       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtTEMCMomentArm_Change"

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

Private Sub txtTEMCRearThrust_Change()
' <VB WATCH>
4628       On Error GoTo vbwErrHandler
4629       Const VBWPROCNAME = "frmPLCData.txtTEMCRearThrust_Change"
4630       If vbwProtector.vbwTraceProc Then
4631           Dim vbwProtectorParameterString As String
4632           If vbwProtector.vbwTraceParameters Then
4633               vbwProtectorParameterString = "()"
4634           End If
4635           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4636       End If
' </VB WATCH>
4637       CalculateTEMCForce
' <VB WATCH>
4638       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4639       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtTEMCRearThrust_Change"

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

Private Sub txtTEMCThrustRigPressure_Change()
' <VB WATCH>
4640       On Error GoTo vbwErrHandler
4641       Const VBWPROCNAME = "frmPLCData.txtTEMCThrustRigPressure_Change"
4642       If vbwProtector.vbwTraceProc Then
4643           Dim vbwProtectorParameterString As String
4644           If vbwProtector.vbwTraceParameters Then
4645               vbwProtectorParameterString = "()"
4646           End If
4647           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4648       End If
' </VB WATCH>
4649       CalculateTEMCForce
' <VB WATCH>
4650       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4651       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtTEMCThrustRigPressure_Change"

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

Private Sub txtTEMCViscosity_Change()
           'CalculateTEMCForce
' <VB WATCH>
4652       On Error GoTo vbwErrHandler
4653       Const VBWPROCNAME = "frmPLCData.txtTEMCViscosity_Change"
4654       If vbwProtector.vbwTraceProc Then
4655           Dim vbwProtectorParameterString As String
4656           If vbwProtector.vbwTraceParameters Then
4657               vbwProtectorParameterString = "()"
4658           End If
4659           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4660       End If
' </VB WATCH>
' <VB WATCH>
4661       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4662       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtTEMCViscosity_Change"

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



Private Sub txtV1_Change()
' <VB WATCH>
4663       On Error GoTo vbwErrHandler
4664       Const VBWPROCNAME = "frmPLCData.txtV1_Change"
4665       If vbwProtector.vbwTraceProc Then
4666           Dim vbwProtectorParameterString As String
4667           If vbwProtector.vbwTraceParameters Then
4668               vbwProtectorParameterString = "()"
4669           End If
4670           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4671       End If
' </VB WATCH>
4672       txtV2.Text = txtV1.Text
4673       txtV3.Text = txtV1.Text
' <VB WATCH>
4674       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4675       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtV1_Change"

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

Private Sub txtVibrationFile_gotfocus()
' <VB WATCH>
4676       On Error GoTo vbwErrHandler
4677       Const VBWPROCNAME = "frmPLCData.txtVibrationFile_gotfocus"
4678       If vbwProtector.vbwTraceProc Then
4679           Dim vbwProtectorParameterString As String
4680           If vbwProtector.vbwTraceParameters Then
4681               vbwProtectorParameterString = "()"
4682           End If
4683           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4684       End If
' </VB WATCH>
4685       On Error GoTo FileCancel
4686       If LenB(txtVibrationFile.Text) <> 0 Then
4687           CommonDialog1.filename = txtVibrationFile.Text
4688       End If
4689       CommonDialog1.ShowOpen
4690       txtVibrationFile.Text = CommonDialog1.filename
' <VB WATCH>
4691       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4692       Exit Sub
4693   FileCancel:
4694   On Error GoTo vbwErrHandler
4695       CommonDialog1.CancelError = False
' <VB WATCH>
4696       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4697       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "txtVibrationFile_gotfocus"

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
Private Sub ExportToExcel()
' <VB WATCH>
4698       On Error GoTo vbwErrHandler
4699       Const VBWPROCNAME = "frmPLCData.ExportToExcel"
4700       If vbwProtector.vbwTraceProc Then
4701           Dim vbwProtectorParameterString As String
4702           If vbwProtector.vbwTraceParameters Then
4703               vbwProtectorParameterString = "()"
4704           End If
4705           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4706       End If
' </VB WATCH>

4707       Dim SaveFileName As String
4708       Dim WorkSheetName As String

4709       Dim I As Integer
4710       Dim iRowNo As Integer
4711       Dim sImp As String
4712       Dim ans As Integer

4713       Dim bCanShowSpeed As Boolean
4714       Dim CantShowReason As String

       'close any running excel processes
4715       Dim objWMIService, colProcesses
4716       Set objWMIService = GetObject("winmgmts:")
4717       Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name LIKE 'Excel%'")
4718       If colProcesses.Count > 0 Then
4719           Set xlApp = Excel.Application
4720       Else
               'use existing copy
       '        Set xlApp = New Excel.Application
4721           Set xlApp = CreateObject("Excel.Application")
4722       End If


4723       CommonDialog1.CancelError = True        'in case the user
4724       On Error GoTo ErrHandler                '  chooses the cancel button

           'set up dialog box
4725       CommonDialog1.DialogTitle = "Open Excel Files"
4726       CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|"  'show Excel files
4727       CommonDialog1.InitDir = App.Path
       '    CommonDialog1.InitDir = "C:\"    'in this directory
4728       CommonDialog1.ShowOpen                              'open the file selection dialog box

4729       If Dir(CommonDialog1.filename) = "" Then            'if the file name does not exist yet
4730           SaveFileName = CommonDialog1.filename           'get the name of the file
4731           If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
4732                xlApp.Workbooks.Close
4733           End If
               ' Create the Excel Workbook Object.
4734   On Error GoTo vbwErrHandler
4735           Set xlBook = xlApp.Workbooks.Add                'add a workbook
4736           WorkSheetName = NewWorkBook                                     'do some stuff for the new workbook
4737           ActiveWorkbook.CheckCompatibility = False
4738           xlApp.ActiveWorkbook.SaveAs filename:=SaveFileName, _
                                 FileFormat:=xlNormal                        'save the file
4739       Else                                                'the file name already exists
4740           SaveFileName = CommonDialog1.filename
               ' Create the Excel Workbook Object.
4741           If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
4742                xlApp.Workbooks.Close
4743           End If
4744           Set xlBook = xlApp.Workbooks.Open(SaveFileName)             'get the file name selected
4745           If GetWorksheetTabs(SaveFileName, WorkSheetName) = vbNo Then    'ask the user if he/she wants a new tab.
4746               MsgBox "File not overwritten.", vbOKOnly, "File not Opened"
' <VB WATCH>
4747       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4748               Exit Sub
4749           Else
4750           End If
4751       End If

4752   On Error GoTo vbwErrHandler

           'see if we can export Speed and SG and if we can, ask user if s/he wants it
           'assume that we can show speed calcs

4753       bCanShowSpeed = False
       'open the template and copy the data from the sheet
       '  excel file resides in ParentDirectoryName + "\Polar SG&Visc Correction5.xls"
           'write the data to the spreadsheet
4754       With xlApp

4755       Dim xlTemplateName As String
4756       xlTemplateName = ParentDirectoryName & sSGandViscSpreadsheetTemplate
4757       Dim xlTemplate As Excel.Workbook
4758       Set xlTemplate = xlApp.Workbooks.Open(xlTemplateName)
4759       Dim TemplateWS As Excel.Worksheet
4760       Dim sheetName As String
4761       sheetName = xlTemplate.Sheets(1).Name
4762       xlTemplate.Sheets(1).Copy After:=xlBook.Sheets(WorkSheetName)

4763       xlTemplate.Close savechanges:=False

4764       Set xlTemplate = Nothing

4765       Application.DisplayAlerts = False
4766       ActiveWorkbook.Worksheets(WorkSheetName).Delete
4767       Application.DisplayAlerts = True
4768       ActiveWorkbook.Worksheets(sheetName).Name = WorkSheetName

           'WorkSheetName = sheetName

           'first see if there is an entry in CalculatedRPM table for this frame size and voltage.
           ' if there is, get the coefficients, else make the coefficients 0

4769           Dim ACoef As Double
4770           Dim BCoef As Double
4771           Dim CCoef As Double

4772           Dim qy As New ADODB.Command
4773           Dim rs As New ADODB.Recordset
4774           qy.ActiveConnection = cnPumpData
4775           Dim VoltageForLookup As Integer
4776           If cmbVoltage.List(cmbVoltage.ListIndex) = "380" And cmbFrequency.List(cmbFrequency.ListIndex) = "50 Hz" Then
4777               VoltageForLookup = 460
4778           ElseIf cmbVoltage.List(cmbVoltage.ListIndex) <> "380" Then
4779               VoltageForLookup = cmbVoltage.List(cmbVoltage.ListIndex)
4780           End If
4781           qy.CommandText = "SELECT * FROM CalculatedRPM WHERE FrameNumber = '" & txtTEMCFrameNumber.Text & _
                          "' AND Voltage = '" & VoltageForLookup & "'"

4782           rs.CursorLocation = adUseClient
4783           rs.CursorType = adOpenStatic

4784           rs.Open qy
4785           If rs.RecordCount = 0 Then
4786               ACoef = 0
4787               BCoef = 0
4788               CCoef = 0
4789               MsgBox ("Cannot find coefficient data for Frame Number " & txtTEMCFrameNumber.Text & _
                          " AND Voltage = " & cmbVoltage.List(cmbVoltage.ListIndex) & _
                          " AND Frequency = " & cmbFrequency.List(cmbFrequency.ListIndex))
4790           Else
4791               ACoef = rs.Fields("A")
4792               BCoef = rs.Fields("B")
4793               CCoef = rs.Fields("C")
4794           End If


           'write header data

4795           .Range("A2").Select
4796           .ActiveCell.FormulaR1C1 = "Serial Number"
4797           .Range("C2").Select
4798           .ActiveCell.FormulaR1C1 = txtSN

4799           .Range("F1").Select
4800           .ActiveCell.FormulaR1C1 = "Customer"
4801           .Range("H1").Select
4802           .ActiveCell.FormulaR1C1 = txtShpNo

4803           .Range("A3").Select
4804           .ActiveCell.FormulaR1C1 = "Model"
4805           .Range("C3").Select
4806           .ActiveCell.FormulaR1C1 = txtModelNo

4807           .Range("F2").Select
4808           .ActiveCell.FormulaR1C1 = "Sales Order"
4809           .Range("H2").Select
4810           .ActiveCell.FormulaR1C1 = txtSalesOrderNumber

4811           .Range("A9").Select
4812           .ActiveCell.FormulaR1C1 = "Design Flow"
4813           .Range("C9").Select
4814           .ActiveCell.FormulaR1C1 = Val(txtDesignFlow)

4815           .Range("A10").Select
4816           .ActiveCell.FormulaR1C1 = "Design Head"
4817           .Range("C10").Select
4818           .ActiveCell.FormulaR1C1 = Val(txtDesignTDH)

4819           .Range("P13").Select
4820           .ActiveCell.FormulaR1C1 = "Barometric Pressure"
4821           .Range("R13").Select
4822           .ActiveCell.FormulaR1C1 = Val(txtInHgDisplay)

4823           .Range("P11").Select
4824           .ActiveCell.FormulaR1C1 = "Suction Gage Height"
4825           .Range("R11").Select
4826           .ActiveCell.FormulaR1C1 = Val(txtSuctHeight)

4827           .Range("P12").Select
4828           .ActiveCell.FormulaR1C1 = "Discharge Gage Height"
4829           .Range("R12").Select
4830           .ActiveCell.FormulaR1C1 = Val(txtDischHeight)

4831           .Range("A1").Select
4832           .ActiveCell.FormulaR1C1 = "Run Date"
4833           .Range("C1").Select
4834           .ActiveCell.FormulaR1C1 = cmbTestDate.List(cmbTestDate.ListIndex)

4835           .Range("D10:E10").Select
4836           With xlApp.Selection
4837               .HorizontalAlignment = xlCenter
4838               .VerticalAlignment = xlBottom
4839               .WrapText = False
4840               .Orientation = 0
4841               .AddIndent = False
4842               .IndentLevel = 0
4843               .ShrinkToFit = False
4844               .ReadingOrder = xlContext
4845               .MergeCells = False
4846           End With
4847           xlApp.Selection.Merge

               'determine rpm

4848           Dim RPMvalue As String
4849           If Mid$(Me.txtTEMCFrameNumber.Text, 2, 1) = "1" Then
               '1 says 2 pole
4850               If Me.cmbFrequency.ListIndex = 0 Then
                       '0 says 50Hz
4851                   RPMvalue = "2900"
4852               ElseIf Me.cmbFrequency.ListIndex = 1 Then
                       ' says 60Hz
4853                   RPMvalue = "3450"
4854               Else
                       'vfd or other, no rpm
4855                   RPMvalue = ""
4856               End If
4857           Else
               '2 says 4 pole
4858               If Me.cmbFrequency.ListIndex = 0 Then
                       '0 says 50Hz
4859                   RPMvalue = "1450"
4860               ElseIf Me.cmbFrequency.ListIndex = 1 Then
                       ' says 60Hz
4861                   RPMvalue = "1750"
4862               Else
                       'vfd or other, no rpm
4863                   RPMvalue = ""
4864               End If
4865           End If

       '        .Range("G1").Select
       '        .ActiveCell.FormulaR1C1 = "RPM"
       '        .Range("I1").Select
       '        .ActiveCell.FormulaR1C1 = RPMvalue

4866           .Range("A5").Select
4867           .ActiveCell.FormulaR1C1 = "Sp Gravity"
4868           .Range("C5").Select
4869           .ActiveCell.FormulaR1C1 = txtSpGr

4870           .Range("A6").Select
4871           .ActiveCell.FormulaR1C1 = "Viscosity"
4872           .Range("C6").Select
4873           .ActiveCell.FormulaR1C1 = txtViscosity

4874           .Range("F4").Select
4875           .ActiveCell.FormulaR1C1 = "Motor"
4876           .Range("H4").Select
4877           .ActiveCell.FormulaR1C1 = txtTEMCFrameNumber.Text

4878           .Range("H12").Select
4879           .ActiveCell.FormulaR1C1 = Me.txtCustPONum.Text

4880           .Range("F5").Select
4881           .ActiveCell.FormulaR1C1 = "Voltage"
4882           .Range("H5").Select
4883           .ActiveCell.FormulaR1C1 = cmbVoltage.List(cmbVoltage.ListIndex)

4884           .Range("K6").Select
4885           .ActiveCell.FormulaR1C1 = "End Play"
4886           .Range("M6").Select
4887           .ActiveCell.FormulaR1C1 = Val(txtEndPlay)

4888           .Range("K7").Select
4889           .ActiveCell.FormulaR1C1 = "G-Gap"
4890           .Range("M7").Select
4891           .ActiveCell.FormulaR1C1 = txtGGap.Text

4892           .Range("A8").Select
4893           .ActiveCell.FormulaR1C1 = "Design Pressure"
4894           .Range("C8").Select
4895           Dim DesPress As String
4896           DesPress = cmbTEMCDesignPressure.List(cmbTEMCDesignPressure.ListIndex)
4897           Dim j As Integer
4898           j = InStrRev(DesPress, "-")
4899           .ActiveCell.FormulaR1C1 = Mid$(DesPress, j + 2)

       '        .Range("G8").Select
       '        .ActiveCell.FormulaR1C1 = "Stator Fill"
       '        .Range("I8").Select
       '        .ActiveCell.FormulaR1C1 = "Dry"

4900           .Range("K4").Select
4901           .ActiveCell.FormulaR1C1 = "Circulation Path"
4902           .Range("M4").Select
4903           .ActiveCell.FormulaR1C1 = cmbTEMCModel.List(cmbTEMCModel.ListIndex)

4904           .Range("M8").Select
4905           .ActiveCell.FormulaR1C1 = txtNPSHr.Text

4906           .Range("K1").Select
4907           .ActiveCell.FormulaR1C1 = "Impeller Dia"
4908           .Range("M1").Select


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

4909           If chkTrimmed.value = 1 Then
4910               If Val(txtImpTrim.Text) <> 0 Then
4911                   .ActiveCell.FormulaR1C1 = Val(txtImpTrim.Text)
4912               Else
4913                   .ActiveCell.FormulaR1C1 = Val(txtImpellerDia.Text)
4914               End If
4915           Else
4916               .ActiveCell.FormulaR1C1 = Val(txtImpellerDia.Text)
4917           End If



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

4918           .Range("P9").Select
4919           .ActiveCell.FormulaR1C1 = "Suction Dia"
4920           .Range("R9").Select
4921           .ActiveCell.FormulaR1C1 = cmbSuctDia.List(cmbSuctDia.ListIndex)

4922           .Range("P10").Select
4923           .ActiveCell.FormulaR1C1 = "Discharge Dia"
4924           .Range("R10").Select
4925           .ActiveCell.FormulaR1C1 = cmbDischDia.List(cmbDischDia.ListIndex)

4926           .Range("A11").Select
4927           .ActiveCell.FormulaR1C1 = "Test Spec"
4928           .Range("C11").Select
4929           .ActiveCell.FormulaR1C1 = cmbTestSpec.List(cmbTestSpec.ListIndex)

4930           .Range("K3").Select
4931           .ActiveCell.FormulaR1C1 = "Impeller Feathered"
4932           .Range("M3").Select
4933           If chkFeathered.value = 1 Then
4934               .ActiveCell.FormulaR1C1 = "Yes"
4935           Else
4936               .ActiveCell.FormulaR1C1 = "No"
4937           End If

4938           .Range("K2").Select
4939           .ActiveCell.FormulaR1C1 = "Disch Orifice"
4940           .Range("M2").Select
4941           If chkOrifice.value = 1 Then
4942               .ActiveCell.FormulaR1C1 = Val(txtOrifice)
4943           Else
4944               .ActiveCell.FormulaR1C1 = "None"
4945           End If


4946           .Range("K5").Select
4947           .ActiveCell.FormulaR1C1 = "Circulation Orifice"
4948           .Range("M5").Select
4949           If chkCircOrifice.value = 1 Then
4950               .ActiveCell.FormulaR1C1 = Val(txtCircOrifice)
4951           Else
4952               .ActiveCell.FormulaR1C1 = "None"
4953           End If

4954           .Range("A13").Select
4955           .ActiveCell.FormulaR1C1 = "Other Mods"
4956           .Range("C13").Select
4957           .ActiveCell.FormulaR1C1 = txtOtherMods

4958           .Range("A14").Select
4959           .ActiveCell.FormulaR1C1 = "Remarks"
4960           .Range("C14").Select
4961           .ActiveCell.FormulaR1C1 = txtRemarks

4962           .Range("A15").Select
4963           .ActiveCell.FormulaR1C1 = "Test Setup Remarks"
4964           .Range("C15").Select
4965           .ActiveCell.FormulaR1C1 = txtTestSetupRemarks

4966           .Range("P1").Select
4967           .ActiveCell.FormulaR1C1 = "Suct ID"
4968           .Range("R1").Select
4969           .ActiveCell.FormulaR1C1 = cmbSuctionPressureTransducer.List(cmbSuctionPressureTransducer.ListIndex)

4970           .Range("P2").Select
4971           .ActiveCell.FormulaR1C1 = "Disch ID"
4972           .Range("R2").Select
4973           .ActiveCell.FormulaR1C1 = cmbDischargePressureTransducer.List(cmbDischargePressureTransducer.ListIndex)

4974           .Range("P3").Select
4975           .ActiveCell.FormulaR1C1 = "Temp ID"
4976           .Range("R3").Select
4977           .ActiveCell.FormulaR1C1 = cmbTemperatureTransducer.List(cmbTemperatureTransducer.ListIndex)

4978           .Range("P4").Select
4979           .ActiveCell.FormulaR1C1 = "Circ Flow ID"
4980           .Range("R4").Select
4981           .ActiveCell.FormulaR1C1 = cmbCirculationFlowMeter.List(cmbCirculationFlowMeter.ListIndex)

4982           .Range("P5").Select
4983           .ActiveCell.FormulaR1C1 = "Flow ID"
4984           .Range("R5").Select
4985           .ActiveCell.FormulaR1C1 = cmbFlowMeter.List(cmbFlowMeter.ListIndex)

4986           .Range("P6").Select
4987           .ActiveCell.FormulaR1C1 = "Analyzer ID"
4988           .Range("R6").Select
4989           .ActiveCell.FormulaR1C1 = cmbAnalyzerNo.List(cmbAnalyzerNo.ListIndex)

4990           .Range("P7").Select
4991           .ActiveCell.FormulaR1C1 = "Loop ID"
4992           .Range("R7").Select
4993           .ActiveCell.FormulaR1C1 = cmbLoopNumber.List(cmbLoopNumber.ListIndex)

4994           .Range("A4").Select
4995           .ActiveCell.FormulaR1C1 = "Fluid"
4996           .Range("C4").Select
4997           .ActiveCell.FormulaR1C1 = txtLiquid.Text

4998           .Range("F3").Select
4999           .ActiveCell.FormulaR1C1 = "Cust PN"
5000           .Range("H3").Select
       '        .ActiveCell.FormulaR1C1 = txtRMA.Text
5001           If rsPumpData.Fields("RVSPartNo") <> "" Then
5002               .ActiveCell.FormulaR1C1 = rsPumpData.Fields("RVSPartNo")
5003           End If
5004           If rsPumpData.Fields("CustPN") <> "" Then
5005               .ActiveCell.FormulaR1C1 = rsPumpData.Fields("CustPN")
5006           End If

5007           .Range("A7").Select
5008           .ActiveCell.FormulaR1C1 = "Temperature"
5009           .Range("C7").Select
5010           .ActiveCell.FormulaR1C1 = txtLiquidTemperature.Text

5011           .Range("F6").Select
5012           .ActiveCell.FormulaR1C1 = "Frequency"
5013           .Range("H6").Select
5014           If UCase(cmbFrequency.List(cmbFrequency.ListIndex)) = "VFD" Then
5015               .ActiveCell.FormulaR1C1 = Val(Me.txtVFDFreq)
5016           Else
5017               .ActiveCell.FormulaR1C1 = Val(cmbFrequency.List(cmbFrequency.ListIndex))
5018           End If
       '        .Range("K2").Select
       '        .ActiveCell.FormulaR1C1 = "Disch Orifice"
       '        .Range("M2").Select
       '        .ActiveCell.FormulaR1C1 = txtOrifice.Text

       '        .Range("K12").Select
       '        .ActiveCell.FormulaR1C1 = "Flow Orifice"
       '        .Range("L12").Select
       '        .ActiveCell.FormulaR1C1 = txtCircOrifice.Text

5019           .Range("P8").Select
5020           .ActiveCell.FormulaR1C1 = "PLC No"
5021           .Range("R8").Select
5022           .ActiveCell.FormulaR1C1 = cmbPLCNo.List(cmbPLCNo.ListIndex)

5023           .Range("F7").Select
5024           .ActiveCell.FormulaR1C1 = "Phases"
5025           .Range("H7").Select
5026           .ActiveCell.FormulaR1C1 = txtNoPhases.Text

5027           .Range("F8").Select
5028           .ActiveCell.FormulaR1C1 = "Poles"
5029           .Range("H8").Select
5030           .ActiveCell.FormulaR1C1 = 2 * Val(Left$(Right$(txtTEMCFrameNumber.Text, 2), 1))

5031           .Range("F9").Select
5032           .ActiveCell.FormulaR1C1 = "Rated Current"
5033           .Range("H9").Select
5034           .ActiveCell.FormulaR1C1 = txtAmps.Text

5035           .Range("F10").Select
5036           .ActiveCell.FormulaR1C1 = "Rated Input Power"
5037           .Range("H10").Select
5038           .ActiveCell.FormulaR1C1 = txtRatedInputPower.Text

5039           .Range("F11").Select
5040           .ActiveCell.FormulaR1C1 = "Insulation Class"
5041           .Range("H11").Select
5042           .ActiveCell.FormulaR1C1 = txtThermalClass.Text

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

5043           .Range("A17").Select
5044           .ActiveCell.FormulaR1C1 = "Flow"
5045           .Range("A18").Select
5046           .ActiveCell.FormulaR1C1 = "(GPM)"

5047           .Range("B17").Select
5048           .ActiveCell.FormulaR1C1 = "TDH"
5049           .Range("B18").Select
5050           .ActiveCell.FormulaR1C1 = "(Ft)"

5051           .Range("C17").Select
5052           .ActiveCell.FormulaR1C1 = "KW"

5053           .Range("D17").Select
5054           .ActiveCell.FormulaR1C1 = "Ave"
5055           .Range("D18").Select
5056           .ActiveCell.FormulaR1C1 = "Volts"

5057           .Range("E17").Select
5058           .ActiveCell.FormulaR1C1 = "Ave"
5059           .Range("E18").Select
5060           .ActiveCell.FormulaR1C1 = "Amps"

5061           .Range("F17").Select
5062           .ActiveCell.FormulaR1C1 = "Power"
5063           .Range("F18").Select
5064           .ActiveCell.FormulaR1C1 = "Factor"

5065           .Range("G17").Select
5066           .ActiveCell.FormulaR1C1 = "Overall"
5067           .Range("G18").Select
5068           .ActiveCell.FormulaR1C1 = "Eff"

5069           .Range("H17").Select
5070           .ActiveCell.FormulaR1C1 = "Measured"
5071           .Range("H18").Select
5072           .ActiveCell.FormulaR1C1 = "RPM"

5073           .Range("I17").Select
5074           .ActiveCell.FormulaR1C1 = "Calculated"
5075           .Range("I18").Select
5076           .ActiveCell.FormulaR1C1 = "RPM"

5077           .Range("J17").Select
5078           .ActiveCell.FormulaR1C1 = "Suction"
5079           .Range("J18").Select
5080           .ActiveCell.FormulaR1C1 = "Temp(F)"

5081           .Range("K17").Select
5082           .ActiveCell.FormulaR1C1 = "Disch"
5083           .Range("K18").Select
5084           .ActiveCell.FormulaR1C1 = "Pressure"

5085           .Range("L17").Select
5086           .ActiveCell.FormulaR1C1 = "Suction"
5087           .Range("L18").Select
5088           .ActiveCell.FormulaR1C1 = "Pressure"

5089           .Range("M17").Select
5090           .ActiveCell.FormulaR1C1 = "Vel"
5091           .Range("M18").Select
5092           .ActiveCell.FormulaR1C1 = "Head"

5093           .Range("N17").Select
5094           .ActiveCell.FormulaR1C1 = "Axial"
5095           .Range("N18").Select
5096           .ActiveCell.FormulaR1C1 = "Position"

5097           .Range("O17").Select
5098           .ActiveCell.FormulaR1C1 = "Pct of"
5099           .Range("O18").Select
5100           .ActiveCell.FormulaR1C1 = "End Play"

5101           .Range("P17").Select
5102           .ActiveCell.FormulaR1C1 = "Hydraulic"
5103           .Range("P18").Select
5104           .ActiveCell.FormulaR1C1 = "Efficiency"

       '        .Range("P17").Select
       '        .ActiveCell.FormulaR1C1 = "Circ"
       '        .Range("P18").Select
       '        .ActiveCell.FormulaR1C1 = "Flow"

5105           .Range("Q17").Select
5106           .ActiveCell.FormulaR1C1 = "Motor"
5107           .Range("Q18").Select
5108           .ActiveCell.FormulaR1C1 = "Efficiency"

5109           .Range("S17").Select
5110           .ActiveCell.FormulaR1C1 = "NPSHa"

5111           .Range("T17").Select
5112           .ActiveCell.FormulaR1C1 = "Phase 1"
5113           .Range("T18").Select
5114           .ActiveCell.FormulaR1C1 = "Current"

5115           .Range("U17").Select
5116           .ActiveCell.FormulaR1C1 = "Phase 2"
5117           .Range("U18").Select
5118           .ActiveCell.FormulaR1C1 = "Current"

5119           .Range("V17").Select
5120           .ActiveCell.FormulaR1C1 = "Phase 3"
5121           .Range("V18").Select
5122           .ActiveCell.FormulaR1C1 = "Current"

5123           .Range("W17").Select
5124           .ActiveCell.FormulaR1C1 = "Phase 1"
5125           .Range("W18").Select
5126           .ActiveCell.FormulaR1C1 = "Voltage"

5127           .Range("X17").Select
5128           .ActiveCell.FormulaR1C1 = "Phase 2"
5129           .Range("X18").Select
5130           .ActiveCell.FormulaR1C1 = "Voltage"

5131           .Range("Y17").Select
5132           .ActiveCell.FormulaR1C1 = "Phase 3"
5133           .Range("Y18").Select
5134           .ActiveCell.FormulaR1C1 = "Voltage"

5135           .Range("Z17").Select
5136           .ActiveCell.FormulaR1C1 = "'" & txtTitle(20).Text

5137           .Range("Z18").Select
5138           .ActiveCell.FormulaR1C1 = "'" & txtTitle(21).Text

5139           .Range("AA17").Select
5140           .ActiveCell.FormulaR1C1 = "'" & txtTitle(22).Text

5141           .Range("AA18").Select
5142           .ActiveCell.FormulaR1C1 = "'" & txtTitle(23).Text

5143           .Range("AB17").Select
5144           .ActiveCell.FormulaR1C1 = "'" & txtTitle(24).Text

5145           .Range("AB18").Select
5146           .ActiveCell.FormulaR1C1 = "'" & txtTitle(25).Text

5147           .Range("AC17").Select
5148           .ActiveCell.FormulaR1C1 = "HR"

5149           .Range("AC18").Select
5150           .ActiveCell.FormulaR1C1 = "(ft)"

5151           .Range("AD17").Select
5152           .ActiveCell.FormulaR1C1 = "'" & txtTitle(26).Text

5153           .Range("AD18").Select
5154           .ActiveCell.FormulaR1C1 = "'" & txtTitle(27).Text

5155           .Range("AE17").Select
5156           .ActiveCell.FormulaR1C1 = "TRG"
5157           .Range("AE18").Select
5158           .ActiveCell.FormulaR1C1 = "Position"

5159           .Range("AF17").Select
5160           .ActiveCell.FormulaR1C1 = "Thrust"

5161           .Range("AG17").Select
5162           .ActiveCell.FormulaR1C1 = "F/R"

5163           .Range("AH17").Select
5164           .ActiveCell.FormulaR1C1 = "Moment"
5165           .Range("AH18").Select
5166           .ActiveCell.FormulaR1C1 = "Arm"

5167           .Range("AI17").Select
5168           .ActiveCell.FormulaR1C1 = "Rig"
5169           .Range("AI18").Select
5170           .ActiveCell.FormulaR1C1 = "Pressure"

       '        .Range("AI17").Select
       '        .ActiveCell.FormulaR1C1 = "Viscosity"

5171           .Range("AJ19").Select
5172           .ActiveCell.FormulaR1C1 = "Rear"
5173           .Range("AJ18").Select
5174           .ActiveCell.FormulaR1C1 = "Force"

5175           .Range("AK17").Select
5176           .ActiveCell.FormulaR1C1 = "PV"

5177           .Range("R17").Select
5178           .ActiveCell.FormulaR1C1 = "Shaft"
5179           .Range("R18").Select
5180           .ActiveCell.FormulaR1C1 = "Power"

       '        .Range("AM17").Select
       '        .ActiveCell.FormulaR1C1 = "Pct Full"
       '        .Range("AM18").Select
       '        .ActiveCell.FormulaR1C1 = "Scale"

5181           .Range("AL17").Select
5182           .ActiveCell.FormulaR1C1 = "NPSHr"

5183           .Range("AM17").Select
5184           .ActiveCell.FormulaR1C1 = "Remarks"




               'now output the data

5185           iRowNo = 20

5186           rsEff.MoveFirst
5187           For I = 1 To frmPLCData.UpDown2.value
5188               .Range("A" & iRowNo).Select
5189               .ActiveCell.FormulaR1C1 = rsEff.Fields("Flow")

5190               .Range("B" & iRowNo).Select
5191               .ActiveCell.FormulaR1C1 = rsEff.Fields("TDH")

5192               .Range("C" & iRowNo).Select
5193               .ActiveCell.FormulaR1C1 = rsEff.Fields("KW")

5194               .Range("D" & iRowNo).Select
5195               .ActiveCell.FormulaR1C1 = rsEff.Fields("Volts")

5196               .Range("E" & iRowNo).Select
5197               .ActiveCell.FormulaR1C1 = rsEff.Fields("Amps")

5198               .Range("F" & iRowNo).Select
5199               .ActiveCell.FormulaR1C1 = rsEff.Fields("PowerFactor")

5200               .Range("G" & iRowNo).Select
5201               .ActiveCell.FormulaR1C1 = rsEff.Fields("OverallEfficiency")

5202               .Range("H" & iRowNo).Select
5203               .ActiveCell.FormulaR1C1 = rsEff.Fields("RPM")

5204               .Range("I" & iRowNo).Select
                   'use the coefficients from above to calculate rpm
5205               Dim f As Double
5206               f = .Range("H6").value
5207               .ActiveCell.FormulaR1C1 = (Val(f) / 60) * (ACoef * (rsEff.Fields("KW")) ^ 2 + BCoef * (rsEff.Fields("KW")) + CCoef)

5208               .Range("J" & iRowNo).Select
5209               .ActiveCell.FormulaR1C1 = rsEff.Fields("Temperature")

5210               .Range("K" & iRowNo).Select
5211               .ActiveCell.FormulaR1C1 = rsEff.Fields("DischPress")

5212               .Range("L" & iRowNo).Select
5213               .ActiveCell.FormulaR1C1 = rsEff.Fields("SuctPress")

5214               .Range("M" & iRowNo).Select
5215               .ActiveCell.FormulaR1C1 = rsEff.Fields("VelocityHead")

5216               .Range("N" & iRowNo).Select
5217               .ActiveCell.FormulaR1C1 = rsEff.Fields("Pos")

5218               .Range("O" & iRowNo).Select
5219               .ActiveCell.FormulaR1C1 = Format((100 * rsEff.Fields("Pos") / Val(txtEndPlay)), "00.0")

5220               .Range("P" & iRowNo).Select
5221               .ActiveCell.FormulaR1C1 = rsEff.Fields("HydraulicEfficiency")

       '            .Range("P" & iRowNo).Select
       '            .ActiveCell.FormulaR1C1 = rsEff.Fields("CircFlow")

5222               .Range("Q" & iRowNo).Select
5223               .ActiveCell.FormulaR1C1 = rsEff.Fields("MotorEfficiency")

5224               .Range("S" & iRowNo).Select
5225               .ActiveCell.FormulaR1C1 = rsEff.Fields("NPSHa")

5226               .Range("T" & iRowNo).Select
5227               .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentA")

5228               .Range("U" & iRowNo).Select
5229               .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentB")

5230               .Range("V" & iRowNo).Select
5231               .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentC")

5232               .Range("W" & iRowNo).Select
5233               .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageA")

5234               .Range("X" & iRowNo).Select
5235               .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageB")

5236               .Range("Y" & iRowNo).Select
5237               .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageC")

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

5238               .Range("Z" & iRowNo).Select
5239               .ActiveCell.FormulaR1C1 = rsEff.Fields("CircFlow")

5240               .Range("AA" & iRowNo).Select
5241               .ActiveCell.FormulaR1C1 = rsEff.Fields("RBHTemp")

5242               .Range("AB" & iRowNo).Select
5243               .ActiveCell.FormulaR1C1 = rsEff.Fields("RBHPress")

5244               .Range("AC" & iRowNo).Select
5245               .ActiveCell.FormulaR1C1 = (rsEff.Fields("RBHPress") - rsEff.Fields("SuctPress")) * 2.31

5246               .Range("AD" & iRowNo).Select
5247               .ActiveCell.FormulaR1C1 = rsEff.Fields("AI4")

5248               .Range("AE" & iRowNo).Select
5249               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCTRG")

5250               .Range("AF" & iRowNo).Select
5251               If rsEff.Fields("TEMCFrontThrust") = 0 Then
5252                   If rsEff.Fields("TEMCRearThrust") = 0 Then
5253                       .ActiveCell.FormulaR1C1 = " "
5254                       .Range("AG" & iRowNo).Select
5255                       .ActiveCell.FormulaR1C1 = " "
5256                   Else
5257                       .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCRearThrust")
5258                       .Range("AG" & iRowNo).Select
5259                       .ActiveCell.FormulaR1C1 = "R"
5260                   End If
5261               Else
5262                   .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCFrontThrust")
5263                   .Range("AG" & iRowNo).Select
5264                   .ActiveCell.FormulaR1C1 = "F"
5265               End If

5266               .Range("AH" & iRowNo).Select
5267               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCMomentArm")

5268               .Range("AI" & iRowNo).Select
5269               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCThrustRigPressure")

       '            .Range("AJ" & iRowNo).Select
       '            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCViscosity")

5270               .Range("AJ" & iRowNo).Select
5271               If rsEff.Fields("TEMCForceDirection") = "F" Then
5272                   .ActiveCell.FormulaR1C1 = -rsEff.Fields("TEMCCalculatedForce")
5273               Else
5274                   .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCCalculatedForce")
5275               End If

5276               .Range("AK" & iRowNo).Select
5277               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCPV")

5278               .Range("R" & iRowNo).Select
5279               .ActiveCell.FormulaR1C1 = rsEff.Fields("KW") * rsEff.Fields("MotorEfficiency") / 100

5280               .Range("AL" & iRowNo).Select
5281               .ActiveCell.FormulaR1C1 = rsEff.Fields("NPSHr")

       '            If RatedKW = 999 Then
       '                .ActiveCell.FormulaR1C1 = ""
       '            Else
       '                .ActiveCell.FormulaR1C1 = (rsEff.Fields("KW") * rsEff.Fields("MotorEfficiency")) / (1 * RatedKW)
       '            End If

5282               .Range("AM" & iRowNo).Select
5283               .ActiveCell.FormulaR1C1 = rsEff.Fields("Remarks")


5284               rsEff.MoveNext
5285               iRowNo = iRowNo + 1
5286           Next I

5287           .Range("A20:AS30").Select
5288           .Selection.NumberFormat = "0.00"

               'format AxPos to 3 dp
5289           .Range("N20:N30").Select
5290           .Selection.NumberFormat = "0.000"

               'format %EndPlay to 1 dp
5291           .Range("O20:O30").Select
5292           .Selection.NumberFormat = "0.0"

           'set up formulas to calculate BEP
           '  first, plot 2nd order polynomial for flow vs hydraulic efficiency
           '  the formulas for doing that are in E68, F68 and G68
           '  only want the formulas to point to the number of points in the test data, so use frmPLCData.CWNumEdit2.value
           '
5293       Dim AColumnRow As String
5294       Dim PColumnRow As String

5295       AColumnRow = "A" & Trim(str(19 + frmPLCData.UpDown2.value))
5296       PColumnRow = "P" & Trim(str(19 + frmPLCData.UpDown2.value))

5297           .Range("E68").Select
5298           .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1)"

5299           .Range("F68").Select
5300           .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1,2)"

5301           .Range("G68").Select
5302           .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1,3)"

           'export balance holes
5303       If boGotBalanceHoles Then
5304           If rsBalanceHoles.State = adStateClosed Then
5305               rsBalanceHoles.ActiveConnection = cnPumpData
5306               rsBalanceHoles.Open
5307           End If 'rsBalanceHoles.State = adStateClosed

5308           If rsBalanceHoles.RecordCount <> 0 Then

5309               .Range("K9:N9").Merge
5310               .Range("K9:N9").Formula = "Balance Hole Data"
5311               .Range("K9:N9").HorizontalAlignment = xlCenter

5312               .Range("K10").Select
5313               .ActiveCell.Formula = "Date"

5314               .Range("L10").Select
5315               .ActiveCell.Formula = "Number"

5316               .Range("M10").Select
5317               .ActiveCell.Formula = "Diameter"

5318               .Range("N10").Select
5319               .ActiveCell.Formula = "Bolt Circle"

5320               iRowNo = 11

5321               If rsBalanceHoles.RecordCount > 3 Then
5322                   For I = 1 To rsBalanceHoles.RecordCount - 3
5323                       Rows("13:13").Select
5324                       Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
5325                   Next I
5326               End If

5327               rsBalanceHoles.MoveFirst
5328               For I = 1 To rsBalanceHoles.RecordCount

5329                   .Range("K" & iRowNo).Select
5330                   .ActiveCell.Formula = rsBalanceHoles.Fields("Date")
5331                   .ActiveCell.NumberFormat = "m/d/yy h:mm AM/PM;@"
5332                   .Range("L" & iRowNo).Select
5333                   .ActiveCell = rsBalanceHoles.Fields("Number")
5334                   .ActiveCell.NumberFormat = "0"
5335                   .Range("M" & iRowNo).Select
5336                   If IsNumeric(rsBalanceHoles.Fields("Diameter1")) Then
5337                       .ActiveCell = Val(rsBalanceHoles.Fields("Diameter1"))
5338                       .ActiveCell.NumberFormat = "0.0000"
5339                   Else
5340                       .ActiveCell = rsBalanceHoles.Fields("Diameter1")
5341                   End If

5342                   .Range("N" & iRowNo).Select
5343                   If IsNumeric(rsBalanceHoles.Fields("BoltCircle1")) Then
5344                       .ActiveCell = Val(rsBalanceHoles.Fields("BoltCircle1"))
5345                       .ActiveCell.NumberFormat = "0.0000"
5346                   Else
5347                       .ActiveCell = rsBalanceHoles.Fields("BoltCircle1")
5348                   End If

5349                   rsBalanceHoles.MoveNext
5350                   iRowNo = iRowNo + 1
5351               Next I
5352               .Range("K10:N" & iRowNo - 1).Select
5353               With .Selection.Interior
5354                   .ColorIndex = 34
5355                   .Pattern = xlSolid
5356               End With
5357           End If 'rsBalanceHoles.RecordCount <> 0
5358       End If ' boGotBalanceHoles

           'plot graphs

5359       Dim SeriesName As String
5360       Dim XVals As String
5361       Dim YVals As String
5362       Dim RowNo As Long
5363       Dim RowStr As String
5364       Dim LastPoint As Integer
5365       Dim LineType As String
5366       Dim AxisGroup As Integer
5367       Dim LabelPos As Integer
5368       Dim LineColor As Long

5369           .ActiveSheet.ChartObjects("HydRepChart").Activate
5370           Dim S As Series
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
5371           Dim aq As Double
5372           Range("AQ56", "AQ71").Select
5373           aq = .Max(Selection)
5374           Dim ax As Double
5375           Range("AX56", "AX71").Select
5376           ax = .Max(Selection)

               'then current (as and az)
5377           Dim at As Double
5378           Range("AS56", "AS71").Select
5379           at = .Max(Selection)
5380           Dim ba As Double
5381           Range("AZ56", "AZ71").Select
5382           ba = .Max(Selection)

5383           Dim CurrentScaleMax As Integer
5384           Dim TDHScaleMax As Integer

5385           Dim MaxTDH As Integer
5386           With Application.WorksheetFunction
5387               If aq > ax Then
5388                   MaxTDH = .Ceiling(aq, 25)
5389               Else
5390                   MaxTDH = .Ceiling(ax, 25)
5391               End If
5392           End With

5393           Dim MaxCurrent As Integer
5394           With Application.WorksheetFunction
5395               If at > ba Then
5396                   Select Case at
                           Case Is <= 5
5397                           CurrentScaleMax = 5

5398                       Case Is <= 10
5399                           CurrentScaleMax = 10

5400                       Case Else
5401                           CurrentScaleMax = 25
5402                   End Select

5403                   MaxCurrent = .Ceiling(at, CurrentScaleMax)
5404               Else
5405                  Select Case ba
                           Case Is <= 5
5406                           CurrentScaleMax = 5

5407                       Case Is <= 10
5408                           CurrentScaleMax = 10

5409                       Case Else
5410                           CurrentScaleMax = 25
5411                   End Select

5412                   MaxCurrent = .Ceiling(ba, CurrentScaleMax)
5413               End If
5414           End With

5415           ActiveSheet.ChartObjects("HydRepChart").Activate
5416            Dim ShtName As String
5417            ShtName = "'" & ActiveSheet.Name & "'"

5418           Dim skipSeries As Boolean
5419           RowStr = 56 + 15
5420            For I = 1 To 8
5421               skipSeries = False
5422                Select Case I
                        Case 1
5423                        SeriesName = "=""TDH"""
5424                        XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
5425                        YVals = "=" & ShtName & "!$AQ$56:$AQ$" & RowStr
5426                        LineType = msoLineSolid
5427                        AxisGroup = 1
5428                        LabelPos = xlLabelPositionRight
5429                        LineColor = vbBlue

5430                    Case 2
5431                        SeriesName = "=""Input Power"""
5432                        XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
5433                        YVals = "=" & ShtName & "!$AR$56:$AR$" & RowStr
5434                        LineType = msoLineSolid
5435                        AxisGroup = 2
5436                        LabelPos = xlLabelPositionRight
5437                        LineColor = vbRed

5438                    Case 3
5439                        SeriesName = "=""Current"""
5440                        XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
5441                        YVals = "=" & ShtName & "!$AS$56:$AS$" & RowStr
5442                        LineType = msoLineSolid
5443                        AxisGroup = 2
5444                        LabelPos = xlLabelPositionRight
5445                        LineColor = vbGreen

5446                    Case 4
5447                         skipSeries = True
       '                     SeriesName = "=""Overall Eff"""
       '                     XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
       '                     YVals = "=" & ShtName & "!$AT$56:$AT$" & RowStr
       '                     LineType = msoLineSolid
       '                     AxisGroup = 2
       '                     LabelPos = xlLabelPositionRight
       '                     LineColor = vbCyan

5448                    Case 5
5449                        SeriesName = "=""TDH (Adj)"""
5450                        XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
5451                        YVals = "=" & ShtName & "!$AX$56:$AX$" & RowStr
5452                        LineType = msoLineDash
5453                        AxisGroup = 1
5454                        LabelPos = xlLabelPositionBelow
5455                        LineColor = vbBlue

5456                    Case 6
5457                        SeriesName = "=""Input Power (Adj)"""
5458                        XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
5459                        YVals = "=" & ShtName & "!$AY$56:$AY$" & RowStr
5460                        LineType = msoLineDash
5461                        AxisGroup = 2
5462                        LabelPos = xlLabelPositionBelow
5463                        LineColor = vbRed

5464                    Case 7
5465                        SeriesName = "=""Current (Adj)"""
5466                        XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
5467                        YVals = "=" & ShtName & "!$AZ$56:$AZ$" & RowStr
5468                        LineType = msoLineDash
5469                        AxisGroup = 2
5470                        LabelPos = xlLabelPositionBelow
5471                        LineColor = vbGreen

5472                    Case 8
5473                       skipSeries = True
       '                     SeriesName = "=""Overall Eff (Adj)"""
       '                     XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
       '                     YVals = "=" & ShtName & "!$BA$56:$BA$" & RowStr
       '                     LineType = msoLineDash
       '                     AxisGroup = 2
       '                     LabelPos = xlLabelPositionBelow
       '                     LineColor = vbCyan

5474               End Select

5475               If Not skipSeries Then
5476                   LastPoint = 16
5477                   ActiveChart.SeriesCollection.NewSeries
5478                   ActiveChart.SeriesCollection(I).Name = SeriesName
5479                   ActiveChart.SeriesCollection(I).XValues = XVals
5480                   ActiveChart.SeriesCollection(I).Values = YVals
5481                   ActiveChart.SeriesCollection(I).Select
5482                   ActiveChart.SeriesCollection(I).Points(LastPoint).Select
5483                   ActiveChart.SeriesCollection(I).Points(LastPoint).ApplyDataLabels
5484                   ActiveChart.SeriesCollection(I).Points(LastPoint).DataLabel.Select
5485                   If I < 5 Then
5486                       Selection.ShowSeriesName = True
5487                       Selection.Position = LabelPos
5488                   Else
5489                       Selection.ShowSeriesName = False
5490                   End If
5491                   Selection.ShowValue = False
5492                   ActiveChart.SeriesCollection(I).ChartType = xlXYScatterSmoothNoMarkers
5493                   ActiveChart.SeriesCollection(I).Select
5494                   With Selection.Format.line
5495                       .Visible = msoTrue
5496                       .DashStyle = LineType
5497                       .ForeColor.RGB = LineColor
5498                   End With


5499                   ActiveChart.SeriesCollection(I).AxisGroup = AxisGroup
5500                   ActiveChart.SeriesCollection(I).DataLabels.Font.Size = 8
5501                   ActiveChart.SeriesCollection(I).DataLabels.Font.Name = "Arial"
5502               End If  'not skip series
5503           Next I

               'show design point
5504           SeriesName = "=""Design Point"""
5505           XVals = "=" & ShtName & "!$L$63"
5506           YVals = "=" & ShtName & "!$L$64"
5507           LineType = msoLineSolid
5508           AxisGroup = 1
5509           ActiveChart.SeriesCollection.NewSeries
5510           ActiveChart.SeriesCollection(I).Name = SeriesName
5511           ActiveChart.SeriesCollection(I).XValues = XVals
5512           ActiveChart.SeriesCollection(I).Values = YVals
5513           ActiveChart.SeriesCollection(I).Select

5514           Selection.MarkerStyle = 4
5515           Selection.MarkerSize = 7
5516           With Selection.Format.line
5517               .Visible = msoTrue
5518               .Weight = 2.25
5519               .ForeColor.RGB = vbBlack
5520           End With


5521           ActiveChart.Axes(xlValue).Select
5522           ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
5523           ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True

5524           ActiveChart.Axes(xlValue).MaximumScale = MaxTDH
5525           ActiveChart.Axes(xlValue).MinimumScale = 0
5526           ActiveChart.Axes(xlValue).MajorUnit = Int(MaxTDH / 5)
5527           Selection.TickLabels.NumberFormat = "0"

5528           ActiveChart.Axes(xlValue, xlSecondary).Select
5529           ActiveChart.Axes(xlValue, xlSecondary).MinimumScaleIsAuto = True
5530           ActiveChart.Axes(xlValue, xlSecondary).MaximumScaleIsAuto = True

5531           ActiveChart.Axes(xlValue, xlSecondary).MaximumScale = MaxCurrent
5532           ActiveChart.Axes(xlValue, xlSecondary).MinimumScale = 0
5533           ActiveChart.Axes(xlValue, xlSecondary).MajorUnit = Int(MaxCurrent / 5)
5534           Selection.TickLabels.NumberFormat = "0"

5535           ActiveChart.Axes(xlValue, xlSecondary).HasTitle = True
5536           ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = "Input Power (kW)-Current (A)"
       '        ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = "Input Power (kW)-Current (A)-Overall Efficiency (%)"
5537           ActiveChart.SetElement (msoElementSecondaryValueAxisTitleRotated)
               'ActiveSheet.PageSetup.PrintArea = "$CA$1:$CI$50"

5538           Range("A1").Select

               'delete all macros in the excel file

               ' Declare variables to access the macros in the workbook.
5539           Dim objProject As VBIDE.VBProject
5540           Dim objComponent As VBIDE.VBComponent
5541           Dim objCode As VBIDE.CodeModule

               ' Get the project details in the workbook.
5542           Set objProject = xlBook.VBProject

               ' Iterate through each component in the project.
5543           For Each objComponent In objProject.VBComponents

                   ' Delete code modules
5544               Set objCode = objComponent.CodeModule
5545               objCode.DeleteLines 1, objCode.CountOfLines

5546               Set objCode = Nothing
5547               Set objComponent = Nothing
5548           Next

5549           Set objProject = Nothing


5550           xlApp.Visible = True                    'show the sheet

5551           xlApp.VBE.ActiveVBProject.VBComponents.Import ParentDirectoryName & sSaveFileMacroFile
5552           xlApp.Run "AssignButton"
5553       End With

       '    Exit Sub

5554   ErrHandler:
           'User pressed the Cancel button

5555       On Error GoTo notopen
5556       If Not xlApp.ActiveWorkbook Is Nothing Then
5557           ActiveWorkbook.CheckCompatibility = False
5558           xlApp.ActiveWorkbook.Save               'save the workbook
               'xlApp.ActiveWorkbook.Close

5559       End If

5560   notopen:

       '    xlApp.Application.Quit

       '    xlApp.Quit
       '    Set xlApp = Nothing

       '    If CommonDialog1.filename <> "" Then
       '        MsgBox CommonDialog1.filename & " has been written.", vbOKOnly, "File Opened"
       '    End If

5561   On Error GoTo vbwErrHandler

' <VB WATCH>
5562       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5563       Exit Sub
' <VB WATCH>
5564       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5565       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ExportToExcel"

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
            vbwReportVariable "SaveFileName", SaveFileName
            vbwReportVariable "WorkSheetName", WorkSheetName
            vbwReportVariable "I", I
            vbwReportVariable "iRowNo", iRowNo
            vbwReportVariable "sImp", sImp
            vbwReportVariable "ans", ans
            vbwReportVariable "bCanShowSpeed", bCanShowSpeed
            vbwReportVariable "CantShowReason", CantShowReason
            vbwReportVariable "objWMIService", objWMIService
            vbwReportVariable "colProcesses", colProcesses
            vbwReportVariable "xlTemplateName", xlTemplateName
            vbwReportVariable "sheetName", sheetName
            vbwReportVariable "ACoef", ACoef
            vbwReportVariable "BCoef", BCoef
            vbwReportVariable "CCoef", CCoef
            vbwReportVariable "VoltageForLookup", VoltageForLookup
            vbwReportVariable "RPMvalue", RPMvalue
            vbwReportVariable "DesPress", DesPress
            vbwReportVariable "j", j
            vbwReportVariable "f", f
            vbwReportVariable "AColumnRow", AColumnRow
            vbwReportVariable "PColumnRow", PColumnRow
            vbwReportVariable "SeriesName", SeriesName
            vbwReportVariable "XVals", XVals
            vbwReportVariable "YVals", YVals
            vbwReportVariable "RowNo", RowNo
            vbwReportVariable "RowStr", RowStr
            vbwReportVariable "LastPoint", LastPoint
            vbwReportVariable "LineType", LineType
            vbwReportVariable "AxisGroup", AxisGroup
            vbwReportVariable "LabelPos", LabelPos
            vbwReportVariable "LineColor", LineColor
            vbwReportVariable "aq", aq
            vbwReportVariable "ax", ax
            vbwReportVariable "at", at
            vbwReportVariable "ba", ba
            vbwReportVariable "CurrentScaleMax", CurrentScaleMax
            vbwReportVariable "TDHScaleMax", TDHScaleMax
            vbwReportVariable "MaxTDH", MaxTDH
            vbwReportVariable "MaxCurrent", MaxCurrent
            vbwReportVariable "ShtName", ShtName
            vbwReportVariable "skipSeries", skipSeries
            vbwReportVariable "xlTemplate", xlTemplate
            vbwReportVariable "TemplateWS", TemplateWS
            vbwReportVariable "qy", qy
            vbwReportVariable "rs", rs
            vbwReportVariable "S", S
            vbwReportVariable "objProject", objProject
            vbwReportVariable "objComponent", objComponent
            vbwReportVariable "objCode", objCode
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Function GetWorksheetTabs(filename As String, WorkSheetName As String)
' <VB WATCH>
5566       On Error GoTo vbwErrHandler
5567       Const VBWPROCNAME = "frmPLCData.GetWorksheetTabs"
5568       If vbwProtector.vbwTraceProc Then
5569           Dim vbwProtectorParameterString As String
5570           If vbwProtector.vbwTraceParameters Then
5571               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("filename", filename) & ", "
5572               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("WorkSheetName", WorkSheetName) & ") "
5573           End If
5574           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5575       End If
' </VB WATCH>

           'see what worksheet tabs alread exist in the excel worksheet

5576       Dim intSheets As Integer    'number of sheets in the workbook
5577       Dim I As Integer
5578       Dim S As String
5579       Dim ans As Integer
5580       Dim NameOK As Boolean

5581       intSheets = xlApp.Worksheets.Count      'how many sheets are there?

           'define a crlf string
5582       S = vbCrLf

5583       For I = 1 To intSheets
5584           S = S & xlApp.Worksheets(I).Name & vbCrLf   'add in the worksheet name
5585       Next I

           'tell the user the names so far and ask if he/she wants to add another
5586       ans = MsgBox("You have the following Worksheet Names in " & filename & ": " & S & "Do you want to add another sheet to this file?", vbYesNo, "Sheets in Excel File")

           'get the answer
5587       If ans = vbNo Then
5588           GetWorksheetTabs = vbNo     'set up flag for when we return to the calling subroutine
' <VB WATCH>
5589       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5590           Exit Function
5591       End If

           'get worksheet name from user and check to see that it's not already used

5592       NameOK = False  'start assuming that the name is bad

5593       While Not NameOK    'as long as it's bad, stay in this loop
5594           WorkSheetName = InputBox("Enter Worksheet Name for this run.")  'ask for name

5595           If WorkSheetName = "" Then      'if we get a nul return or user presses cancel
5596               GetWorksheetTabs = vbNo
' <VB WATCH>
5597       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5598               Exit Function
5599           End If

5600           For I = 1 To xlApp.Worksheets.Count     'go through all of the existing sheets
5601               If WorkSheetName = xlApp.Worksheets(I).Name Then        'if the names are the same
5602                   MsgBox "The name " & WorkSheetName & " already exists for a Worksheet.  Please try again.", vbOKOnly, "Bad Worksheet Name"  'tell the user
5603                   NameOK = False
5604                   Exit For
5605               End If
5606               NameOK = True       'if we make it thru say the name is ok
5607           Next I
5608       Wend

5609       xlApp.Worksheets.Add , xlApp.Worksheets(xlApp.Worksheets.Count)     'add a worksheer
5610       xlApp.Worksheets(xlApp.Worksheets.Count).Name = WorkSheetName       'give it the desired name
5611       GetWorksheetTabs = vbYes                                            'say that the results were ok

' <VB WATCH>
5612       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5613       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetWorksheetTabs"

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
            vbwReportVariable "filename", filename
            vbwReportVariable "WorkSheetName", WorkSheetName
            vbwReportVariable "intSheets", intSheets
            vbwReportVariable "I", I
            vbwReportVariable "S", S
            vbwReportVariable "ans", ans
            vbwReportVariable "NameOK", NameOK
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Function NewWorkBook() As String
' <VB WATCH>
5614       On Error GoTo vbwErrHandler
5615       Const VBWPROCNAME = "frmPLCData.NewWorkBook"
5616       If vbwProtector.vbwTraceProc Then
5617           Dim vbwProtectorParameterString As String
5618           If vbwProtector.vbwTraceParameters Then
5619               vbwProtectorParameterString = "()"
5620           End If
5621           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5622       End If
' </VB WATCH>

5623       Dim WorkSheetName As String

           'we've just added a new workbook, delete sheet1, sheet2, etc
5624       xlApp.DisplayAlerts = False
5625       While xlApp.Worksheets.Count > 1
5626           xlApp.Worksheets(1).Delete          'delete the sheet
5627       Wend
5628       xlApp.DisplayAlerts = True

5629       WorkSheetName = InputBox("Enter Title Worksheet Name for this run.")    'get the desired name
5630       xlApp.Worksheets(1).Name = WorkSheetName    'and name the sheet

5631       NewWorkBook = WorkSheetName

' <VB WATCH>
5632       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5633       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "NewWorkBook"

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
            vbwReportVariable "WorkSheetName", WorkSheetName
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Private Sub CalibrateSoftware()
' <VB WATCH>
5634       On Error GoTo vbwErrHandler
5635       Const VBWPROCNAME = "frmPLCData.CalibrateSoftware"
5636       If vbwProtector.vbwTraceProc Then
5637           Dim vbwProtectorParameterString As String
5638           If vbwProtector.vbwTraceParameters Then
5639               vbwProtectorParameterString = "()"
5640           End If
5641           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5642       End If
' </VB WATCH>
5643           frmCalibrate.Show
               'Calibrating = True

' <VB WATCH>
5644       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5645       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "CalibrateSoftware"

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

Function ParseTEMCModelNo(cmbComboName As ComboBox, ltr As String)
' <VB WATCH>
5646       On Error GoTo vbwErrHandler
5647       Const VBWPROCNAME = "frmPLCData.ParseTEMCModelNo"
5648       If vbwProtector.vbwTraceProc Then
5649           Dim vbwProtectorParameterString As String
5650           If vbwProtector.vbwTraceParameters Then
5651               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
5652               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ltr", ltr) & ") "
5653           End If
5654           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5655       End If
' </VB WATCH>
5656       Dim I As Integer
5657       Dim iStart As Integer
5658       Dim iStop As Integer
5659       Dim strCompare As String

5660       For I = 0 To cmbComboName.ListCount - 1                     'go through the combobox entries
5661           iStart = InStr(1, cmbComboName.List(I), "[")
5662           iStop = InStr(1, cmbComboName.List(I), "]")
5663           strCompare = Mid$(cmbComboName.List(I), iStart + 1, iStop - iStart - 1)
5664           If UCase(strCompare) = UCase(ltr) Then   'see when we find the desired index number
5665               cmbComboName.ListIndex = I                                              'if we do, set the combo box
5666               Exit For                                            'and we're done
5667           End If
       '        cmbComboName.ListIndex = -1                             'else, remove any pointer
5668           cmbComboName.ListIndex = cmbComboName.ListCount - 1                           'else, remove any pointer
5669       Next I

5670       txtModelNo.Text = UCase(txtModelNo.Text)
5671       txtModelNo.SelStart = Len(txtModelNo.Text)
' <VB WATCH>
5672       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5673       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "ParseTEMCModelNo"

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
            vbwReportVariable "ltr", ltr
            vbwReportVariable "I", I
            vbwReportVariable "iStart", iStart
            vbwReportVariable "iStop", iStop
            vbwReportVariable "strCompare", strCompare
            vbwReportVariable "cmbComboName", cmbComboName
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Public Function LoadCombo(cmbComboName As ComboBox, sTableName As String)
       'load all of the pump parameter combo boxes from the tables on the database
' <VB WATCH>
5674       On Error GoTo vbwErrHandler
5675       Const VBWPROCNAME = "frmPLCData.LoadCombo"
5676       If vbwProtector.vbwTraceProc Then
5677           Dim vbwProtectorParameterString As String
5678           If vbwProtector.vbwTraceParameters Then
5679               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
5680               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sTableName", sTableName) & ") "
5681           End If
5682           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5683       End If
' </VB WATCH>

5684       Dim I As Integer
5685       Dim sItem As String
5686       Dim iID As Integer
5687       Dim bUseDropdown As Boolean
5688       Dim qy As New ADODB.Command
5689       Dim rs As New ADODB.Recordset

       '    rsPumpParameters.CursorLocation = adUseClient
       '    If sTableName = "Model" Then
       '        rsPumpParameters.Sort = "Model"
       '    Else
       '        rsPumpParameters.Sort = vbNullString
       '    End If
       '    rsPumpParameters.Open sTableName, cnPumpData, adOpenStatic, adLockOptimistic, adCmdTableDirect

5690       qy.ActiveConnection = cnPumpData
5691       If sTableName = "DischargeDiameter" Or sTableName = "SuctionDiameter" Then
5692           qy.CommandText = "SELECT * FROM " & sTableName & " ORDER BY Val(Description)"
5693       Else
5694           qy.CommandText = "SELECT * FROM " & sTableName & " ORDER BY Description"
5695       End If
5696       If sTableName = "SupermarketPumpData" Then
5697           qy.CommandText = "SELECT ID,Model AS Description FROM " & sTableName
5698       End If
5699       rs.CursorLocation = adUseClient
5700       rs.CursorType = adOpenStatic

5701       rs.Open qy


5702       On Error GoTo NoField
5703       bUseDropdown = True
           'sItem = rsPumpParameters.Fields("UseInDropdown")
       '    If bUseDropdown Then
       '        rsPumpParameters.Sort = "Description"
       '    End If
5704       rs.MoveFirst                                'goto the top
5705       For I = 0 To rs.RecordCount - 1             'go through the whole recordset
5706           sItem = rs.Fields("Description")        'get the description
5707           iID = rs.Fields(0)                      'get the index number - primary key
5708           If bUseDropdown Then
       '            If rsPumpParameters.Fields("UseInDropdown").value = True Then
5709                   cmbComboName.AddItem sItem, I                                   'add the description to the combo box
       '                cmbComboName.AddItem sItem                                   'add the description to the combo box
5710                   cmbComboName.ItemData(cmbComboName.NewIndex) = iID              'add the key number into the item data
       '            End If
5711           End If
5712           rs.MoveNext                             'get the next record
5713       Next I
5714       rs.Close
5715       cmbComboName.ListIndex = -1
5716   On Error GoTo vbwErrHandler
5717       Set rs = Nothing
5718       Set qy = Nothing
' <VB WATCH>
5719       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5720       Exit Function

5721   NoField:
5722       bUseDropdown = False
5723   On Error GoTo vbwErrHandler
5724       Resume Next

' <VB WATCH>
5725       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5726       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "LoadCombo"

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
            vbwReportVariable "sTableName", sTableName
            vbwReportVariable "I", I
            vbwReportVariable "sItem", sItem
            vbwReportVariable "iID", iID
            vbwReportVariable "bUseDropdown", bUseDropdown
            vbwReportVariable "cmbComboName", cmbComboName
            vbwReportVariable "qy", qy
            vbwReportVariable "rs", rs
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Function SetGraphMax(Plothead) As Integer
' <VB WATCH>
5727       On Error GoTo vbwErrHandler
5728       Const VBWPROCNAME = "frmPLCData.SetGraphMax"
5729       If vbwProtector.vbwTraceProc Then
5730           Dim vbwProtectorParameterString As String
5731           If vbwProtector.vbwTraceParameters Then
5732               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Plothead", Plothead) & ") "
5733           End If
5734           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5735       End If
' </VB WATCH>

5736       Dim I As Integer
5737       Dim m As Single

5738       m = 0
5739       For I = 0 To UBound(Plothead, 2)
5740           If Plothead(1, I) > m Then
5741               m = Plothead(1, I)
5742           End If
5743       Next I
5744       SetGraphMax = 10 * (Int((m / 10) + 0.5) + 1)
5745       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
5746       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10 * (Int((m / 10) + 0.5) + 1)
5747       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0

' <VB WATCH>
5748       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5749       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SetGraphMax"

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
            vbwReportVariable "Plothead", Plothead
            vbwReportVariable "I", I
            vbwReportVariable "m", m
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Public Function CalculateSpeed(CoefSq As Double, CoefLin As Double, CoefConstant As Double, InputHP As Double, SG As Double) As Integer
' <VB WATCH>
5750       On Error GoTo vbwErrHandler
5751       Const VBWPROCNAME = "frmPLCData.CalculateSpeed"
5752       If vbwProtector.vbwTraceProc Then
5753           Dim vbwProtectorParameterString As String
5754           If vbwProtector.vbwTraceParameters Then
5755               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("CoefSq", CoefSq) & ", "
5756               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("CoefLin", CoefLin) & ", "
5757               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("CoefConstant", CoefConstant) & ", "
5758               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("InputHP", InputHP) & ", "
5759               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("SG", SG) & ") "
5760           End If
5761           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5762       End If
' </VB WATCH>
5763       Dim I As Integer
5764       Dim OldResult As Double
5765       Dim NewResult As Double

5766       CalculateSpeed = 0

5767       If SG > 5 Or SG < 0.01 Then
5768           MsgBox "Bad value for SG...must be between 0.01 and 5.", vbOKOnly, "Bad SG Value"
' <VB WATCH>
5769       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5770           Exit Function
5771       End If

5772       OldResult = 1000
5773       NewResult = 0

5774       I = 1

5775       Do While Abs(NewResult - OldResult) > 0.1
5776           ReDim Preserve results(I)
5777           Select Case I
                   Case 1
5778                   results(I - 1).HP = InputHP
5779               Case 2
5780                   results(I - 1).HP = results(I - 2).HP * SG
5781               Case Else
5782                   results(I - 1).HP = results(I - 2).HP * (results(I - 2).Speed / results(I - 3).Speed) ^ 3
5783           End Select
5784           OldResult = NewResult
5785           results(I - 1).Speed = CalcPoly(CoefSq, CoefLin, CoefConstant, results(I - 1).HP)
5786           NewResult = results(I - 1).Speed
5787           If I > 15 Then
5788               If I = 0 Or I > 15 Then
5789                   MsgBox "Over 15 calculations and no convergence", vbOKOnly, "Too many iterations"
' <VB WATCH>
5790       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5791                   Exit Function
5792               End If
' <VB WATCH>
5793       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5794               Exit Function
5795           End If
5796           I = I + 1
5797       Loop
5798       CalculateSpeed = I - 1
' <VB WATCH>
5799       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5800       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "CalculateSpeed"

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
            vbwReportVariable "CoefSq", CoefSq
            vbwReportVariable "CoefLin", CoefLin
            vbwReportVariable "CoefConstant", CoefConstant
            vbwReportVariable "InputHP", InputHP
            vbwReportVariable "SG", SG
            vbwReportVariable "I", I
            vbwReportVariable "OldResult", OldResult
            vbwReportVariable "NewResult", NewResult
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function
Public Function CalcPoly(CoefSq As Double, CoefLin As Double, CoefConstant As Double, DataIn As Double) As Double
' <VB WATCH>
5801       On Error GoTo vbwErrHandler
5802       Const VBWPROCNAME = "frmPLCData.CalcPoly"
5803       If vbwProtector.vbwTraceProc Then
5804           Dim vbwProtectorParameterString As String
5805           If vbwProtector.vbwTraceParameters Then
5806               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("CoefSq", CoefSq) & ", "
5807               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("CoefLin", CoefLin) & ", "
5808               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("CoefConstant", CoefConstant) & ", "
5809               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("DataIn", DataIn) & ") "
5810           End If
5811           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5812       End If
' </VB WATCH>
5813       CalcPoly = CoefSq * DataIn ^ 2 + CoefLin * DataIn + CoefConstant
' <VB WATCH>
5814       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5815       Exit Function
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "CalcPoly"

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
            vbwReportVariable "CoefSq", CoefSq
            vbwReportVariable "CoefLin", CoefLin
            vbwReportVariable "CoefConstant", CoefConstant
            vbwReportVariable "DataIn", DataIn
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Function

Sub GetBalanceHoleData(SerialNumber As String, TestDate As String)
' <VB WATCH>
5816       On Error GoTo vbwErrHandler
5817       Const VBWPROCNAME = "frmPLCData.GetBalanceHoleData"
5818       If vbwProtector.vbwTraceProc Then
5819           Dim vbwProtectorParameterString As String
5820           If vbwProtector.vbwTraceParameters Then
5821               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("SerialNumber", SerialNumber) & ", "
5822               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("TestDate", TestDate) & ") "
5823           End If
5824           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5825       End If
' </VB WATCH>
5826       If rsBalanceHoles.State = adStateOpen Then
5827           rsBalanceHoles.Close
5828       End If
5829       qyBalanceHoles.CommandText = "SELECT BalanceHoles.*, " & _
                             "IIf([Diameter]=99, 'Slot', [diameter]) as Diameter1, IIf([BoltCircle]=99, 'Unknown', [BoltCircle]) as BoltCircle1 " & _
                             "FROM BalanceHoles " & _
                             "WHERE [SerialNo] = '" & SerialNumber & "' AND [Date] <= #" & TestDate & "# " & _
                             "ORDER BY [Date], Val([BoltCircle]);"

5830       rsBalanceHoles.Open qyBalanceHoles
5831       rsBalanceHoles.Filter = ""

5832       Set dgBalanceHoles.DataSource = rsBalanceHoles

5833       Dim c As Column
5834       For Each c In dgBalanceHoles.Columns
5835           Select Case c.DataField
               Case "BalanceHoleID"
5836               c.Visible = False
5837           Case "SerialNo"
5838               c.Visible = False
5839           Case "Date"
5840               c.Visible = True
5841               c.Alignment = dbgCenter
5842               c.Width = 2000
5843           Case "Number"
5844               c.Visible = True
5845               c.Alignment = dbgCenter
5846               c.Width = 700
5847           Case "Diameter"
5848               c.Visible = False
5849           Case "Diameter1"
5850               c.Caption = "Diameter"
5851               c.Visible = True
5852               c.Alignment = dbgCenter
5853               c.Width = 700
5854           Case "BoltCircle1"
5855               c.Caption = "Bolt Circle"
5856               c.Visible = True
5857               c.Alignment = dbgCenter
5858               c.Width = 800
5859           Case "BoltCircle"
5860               c.Visible = False
5861           Case Else ' hide all other columns.
5862               c.Visible = False
5863           End Select
5864       Next c

' <VB WATCH>
5865       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5866       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "GetBalanceHoleData"

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
            vbwReportVariable "SerialNumber", SerialNumber
            vbwReportVariable "TestDate", TestDate
            vbwReportVariable "c", c
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Public Sub FixPointsToPlot()
           'count valid data test entry and set points to plot
' <VB WATCH>
5867       On Error GoTo vbwErrHandler
5868       Const VBWPROCNAME = "frmPLCData.FixPointsToPlot"
5869       If vbwProtector.vbwTraceProc Then
5870           Dim vbwProtectorParameterString As String
5871           If vbwProtector.vbwTraceParameters Then
5872               vbwProtectorParameterString = "()"
5873           End If
5874           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5875       End If
' </VB WATCH>
5876       If DataGrid2.Row = -1 Then
' <VB WATCH>
5877       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5878           Exit Sub
5879       End If
5880       Dim PresentGridRow As Integer
5881       PresentGridRow = DataGrid2.Row
5882       Dim GridIndex As Integer
5883       UpDown2.value = 8
5884       If DataGrid2.Row <> -1 Then
5885           For GridIndex = 0 To 7
5886               DataGrid2.Row = GridIndex
5887               If DataGrid2.Columns("Flow") = 0 And DataGrid2.Columns("TDH") = 0 Then
5888                   txtUpDn2.Text = GridIndex
5889                   If GridIndex = 0 Then
5890                       UpDown2.value = 8
5891                   Else
5892                       UpDown2.value = GridIndex
5893                   End If
' <VB WATCH>
5894       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5895                   Exit Sub
5896               End If
5897           Next GridIndex
5898       End If
5899       DataGrid2.Row = PresentGridRow
' <VB WATCH>
5900       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5901       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "FixPointsToPlot"

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
            vbwReportVariable "PresentGridRow", PresentGridRow
            vbwReportVariable "GridIndex", GridIndex
            vbwReportModuleVariables
            vbwReportGlobalVariables
            vbwCloseDumpFile
            Err.Number = -1
            Goto vbwErrHandler
    End Select
    If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
End Sub

Sub SetFrequencyCombo()
           'set default for test spec
' <VB WATCH>
5902       On Error GoTo vbwErrHandler
5903       Const VBWPROCNAME = "frmPLCData.SetFrequencyCombo"
5904       If vbwProtector.vbwTraceProc Then
5905           Dim vbwProtectorParameterString As String
5906           If vbwProtector.vbwTraceParameters Then
5907               vbwProtectorParameterString = "()"
5908           End If
5909           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5910       End If
' </VB WATCH>
5911       Dim j As Integer
5912       For j = 0 To cmbFrequency.ListCount - 1
5913           If cmbFrequency.List(j) = "60 Hz" Then
5914               cmbFrequency.ListIndex = j
5915               Exit For
5916           End If
5917       Next

' <VB WATCH>
5918       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5919       Exit Sub
    ' ----- Error Handler ------
vbwErrHandler:
    Const VBWPROCEDURE = "SetFrequencyCombo"

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
            vbwReportVariable "j", j
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
' Procedure added by VB Watch 'ID
Private Sub Form_Initialize() 'ID
    vbwInitializeProtector ' Initialize VB Watch 'ID
End Sub 'ID
' </VB WATCH>
' <VB WATCH> <VBWATCHFINALPROC>
' Procedures added by VB Watch for variable dump


Private Sub vbwReportModuleVariables()
    vbwReportToFile VBW_MODULE_STRING
    vbwReportVariable "debugging", debugging
    vbwReportVariable "sDataBaseName", sDataBaseName
    vbwReportVariable "boFoundPump", boFoundPump
    vbwReportVariable "boPumpIsApproved", boPumpIsApproved
    vbwReportVariable "boTestDateIsApproved", boTestDateIsApproved
    vbwReportVariable "boFoundTestSetup", boFoundTestSetup
    vbwReportVariable "boFoundTestData", boFoundTestData
    vbwReportVariable "boUsingEpicor", boUsingEpicor
    vbwReportVariable "boUsingSupermarketTable", boUsingSupermarketTable
    vbwReportVariable "boEpicorFound", boEpicorFound
    vbwReportVariable "boPLCOperating", boPLCOperating
    vbwReportVariable "boMagtrolOperating", boMagtrolOperating
    vbwReportVariable "boGotBalanceHoles", boGotBalanceHoles
    vbwReportVariable "HeadFlow", HeadFlow
    vbwReportVariable "EffFlow", EffFlow
    vbwReportVariable "KWFlow", KWFlow
    vbwReportVariable "AmpsFlow", AmpsFlow
    vbwReportVariable "FlowHead", FlowHead
    vbwReportVariable "RatedKW", RatedKW
    vbwReportVariable "blnEnabled", blnEnabled
    vbwReportVariable "EpicorConnectionString", EpicorConnectionString
    vbwReportVariable "ParentDirectoryName", ParentDirectoryName
    vbwReportVariable "ProgramEnd", ProgramEnd
    vbwReportVariable "Pressed", Pressed
    vbwReportVariable "rsPumpData", rsPumpData
    vbwReportVariable "rsTestSetup", rsTestSetup
    vbwReportVariable "rsTestData", rsTestData
    vbwReportVariable "rsMisc", rsMisc
    vbwReportVariable "rsEff", rsEff
    vbwReportVariable "rsBalanceHoles", rsBalanceHoles
    vbwReportVariable "rsPumpParameters", rsPumpParameters
    vbwReportVariable "rsSupermarketModel", rsSupermarketModel
    vbwReportVariable "qyPumpData", qyPumpData
    vbwReportVariable "qyTestSetup", qyTestSetup
    vbwReportVariable "qyBalanceHoles", qyBalanceHoles
    vbwReportVariable "qySupermarketModel", qySupermarketModel
    vbwReportVariable "qyMisc", qyMisc
    vbwReportVariable "xlApp", xlApp
    vbwReportVariable "xlBook", xlBook
End Sub
' </VB WATCH>
