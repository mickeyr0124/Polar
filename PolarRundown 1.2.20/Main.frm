VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
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
      TabIndex        =   265
      Top             =   480
      Width           =   1575
   End
   Begin VB.CommandButton cmdCalibrate 
      Caption         =   "Calibrate Software"
      Height          =   495
      Left            =   9360
      TabIndex        =   190
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
      TabIndex        =   111
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
      Tab             =   2
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
      Tab(0).ControlEnabled=   0   'False
      Tab(0).Control(0)=   "lbltab1(1)"
      Tab(0).Control(1)=   "lbltab1(2)"
      Tab(0).Control(2)=   "lbltab1(3)"
      Tab(0).Control(3)=   "lbltab1(11)"
      Tab(0).Control(4)=   "lbltab1(12)"
      Tab(0).Control(5)=   "lbltab1(13)"
      Tab(0).Control(6)=   "lbltab1(0)"
      Tab(0).Control(7)=   "lbltab1(10)"
      Tab(0).Control(8)=   "lbltab1(44)"
      Tab(0).Control(9)=   "lbltab1(46)"
      Tab(0).Control(10)=   "lbltab1(47)"
      Tab(0).Control(11)=   "lbltab1(48)"
      Tab(0).Control(12)=   "lbltab1(49)"
      Tab(0).Control(13)=   "lbltab1(50)"
      Tab(0).Control(14)=   "frmChempump"
      Tab(0).Control(15)=   "txtBilNo"
      Tab(0).Control(16)=   "txtShpNo"
      Tab(0).Control(17)=   "txtModelNo"
      Tab(0).Control(18)=   "txtDesignFlow"
      Tab(0).Control(19)=   "txtDesignTDH"
      Tab(0).Control(20)=   "txtRemarks"
      Tab(0).Control(21)=   "cmdEnterPumpData"
      Tab(0).Control(22)=   "txtSalesOrderNumber"
      Tab(0).Control(23)=   "cmdDeletePump"
      Tab(0).Control(24)=   "cmdApprovePump"
      Tab(0).Control(25)=   "frmMfr"
      Tab(0).Control(26)=   "cmdClearPumpData"
      Tab(0).Control(27)=   "txtImpellerDia"
      Tab(0).Control(28)=   "frmMiscPumpData"
      Tab(0).Control(29)=   "txtLineNumber"
      Tab(0).Control(30)=   "frmTEMC"
      Tab(0).Control(31)=   "CommonDialog2"
      Tab(0).Control(32)=   "grpSupermarket"
      Tab(0).Control(33)=   "chkSuperMarketFeathered"
      Tab(0).Control(34)=   "txtRVSPartNo"
      Tab(0).Control(35)=   "txtXPartNum"
      Tab(0).Control(36)=   "txtCustPONum"
      Tab(0).ControlCount=   37
      TabCaption(1)   =   "Test Setup"
      TabPicture(1)   =   "Main.frx":001C
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "txtRMA"
      Tab(1).Control(1)=   "Command1"
      Tab(1).Control(2)=   "frmTAndI"
      Tab(1).Control(3)=   "cmdApproveTestDate"
      Tab(1).Control(4)=   "cmdDeleteTestDate"
      Tab(1).Control(5)=   "CommonDialog1"
      Tab(1).Control(6)=   "frmOtherFiles"
      Tab(1).Control(7)=   "frmPerfMods"
      Tab(1).Control(8)=   "frmThrustBalMods"
      Tab(1).Control(9)=   "frmElecData"
      Tab(1).Control(10)=   "frmLoopAndXducer"
      Tab(1).Control(11)=   "frmInstrumentTags"
      Tab(1).Control(12)=   "txtTestSetupRemarks"
      Tab(1).Control(13)=   "cmdAddNewTestDate"
      Tab(1).Control(14)=   "txtWho"
      Tab(1).Control(15)=   "cmdEnterTestSetupData"
      Tab(1).Control(16)=   "cmbTestSpec"
      Tab(1).Control(17)=   "lbltab2(88)"
      Tab(1).Control(18)=   "lbltab2(65)"
      Tab(1).Control(19)=   "lbltab2(1)"
      Tab(1).Control(20)=   "lbltab2(0)"
      Tab(1).ControlCount=   21
      TabCaption(2)   =   "Test Data"
      TabPicture(2)   =   "Main.frx":0038
      Tab(2).ControlEnabled=   -1  'True
      Tab(2).Control(0)=   "lbltab2(55)"
      Tab(2).Control(0).Enabled=   0   'False
      Tab(2).Control(1)=   "lbltab2(63)"
      Tab(2).Control(1).Enabled=   0   'False
      Tab(2).Control(2)=   "lbltab2(64)"
      Tab(2).Control(2).Enabled=   0   'False
      Tab(2).Control(3)=   "lbltab2(53)"
      Tab(2).Control(3).Enabled=   0   'False
      Tab(2).Control(4)=   "lbltab2(54)"
      Tab(2).Control(4).Enabled=   0   'False
      Tab(2).Control(5)=   "shpGetPLCData"
      Tab(2).Control(5).Enabled=   0   'False
      Tab(2).Control(6)=   "UpDown2"
      Tab(2).Control(6).Enabled=   0   'False
      Tab(2).Control(7)=   "DataGrid1"
      Tab(2).Control(7).Enabled=   0   'False
      Tab(2).Control(8)=   "cmbPLCLoop"
      Tab(2).Control(8).Enabled=   0   'False
      Tab(2).Control(9)=   "frmAI"
      Tab(2).Control(9).Enabled=   0   'False
      Tab(2).Control(10)=   "frmThermocouples"
      Tab(2).Control(10).Enabled=   0   'False
      Tab(2).Control(11)=   "fmrMiscTestData"
      Tab(2).Control(11).Enabled=   0   'False
      Tab(2).Control(12)=   "cmdEnterTestData"
      Tab(2).Control(12).Enabled=   0   'False
      Tab(2).Control(13)=   "DataGrid2"
      Tab(2).Control(13).Enabled=   0   'False
      Tab(2).Control(14)=   "frmPLCMisc"
      Tab(2).Control(14).Enabled=   0   'False
      Tab(2).Control(15)=   "frmPumpData"
      Tab(2).Control(15).Enabled=   0   'False
      Tab(2).Control(16)=   "txtNPSHa"
      Tab(2).Control(16).Enabled=   0   'False
      Tab(2).Control(17)=   "cmdReport"
      Tab(2).Control(17).Enabled=   0   'False
      Tab(2).Control(18)=   "txtTDH"
      Tab(2).Control(18).Enabled=   0   'False
      Tab(2).Control(19)=   "frmMagtrol"
      Tab(2).Control(19).Enabled=   0   'False
      Tab(2).Control(20)=   "btnRunNPSH"
      Tab(2).Control(20).Enabled=   0   'False
      Tab(2).Control(21)=   "frmNPSH"
      Tab(2).Control(21).Enabled=   0   'False
      Tab(2).Control(22)=   "UpDown1"
      Tab(2).Control(22).Enabled=   0   'False
      Tab(2).Control(23)=   "txtUpDn1"
      Tab(2).Control(23).Enabled=   0   'False
      Tab(2).Control(24)=   "MSChart1"
      Tab(2).Control(24).Enabled=   0   'False
      Tab(2).Control(25)=   "txtUpDn2"
      Tab(2).Control(25).Enabled=   0   'False
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
         Left            =   14160
         TabIndex        =   427
         Text            =   "8"
         Top             =   5520
         Width           =   285
      End
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   2772
         Left            =   6960
         OleObjectBlob   =   "Main.frx":0070
         TabIndex        =   426
         Top             =   3000
         Width           =   5652
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
         Left            =   960
         TabIndex        =   425
         Text            =   "1"
         Top             =   8880
         Width           =   285
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   504
         Left            =   720
         TabIndex        =   424
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
         Left            =   -67200
         TabIndex        =   422
         Top             =   1080
         Width           =   1335
      End
      Begin VB.TextBox txtXPartNum 
         Height          =   315
         Left            =   -73560
         TabIndex        =   420
         Top             =   1080
         Width           =   4932
      End
      Begin VB.TextBox txtRVSPartNo 
         Height          =   315
         Left            =   -69480
         TabIndex        =   419
         Top             =   4349
         Width           =   1932
      End
      Begin VB.CheckBox chkSuperMarketFeathered 
         Caption         =   "Check1"
         Enabled         =   0   'False
         Height          =   252
         Left            =   -60600
         TabIndex        =   415
         Top             =   4320
         Width           =   252
      End
      Begin VB.Frame frmNPSH 
         Caption         =   "TDH (ft)"
         Height          =   1812
         Left            =   11280
         TabIndex        =   402
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
            TabIndex        =   413
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
            TabIndex        =   412
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
            TabIndex        =   411
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
            TabIndex        =   409
            Top             =   1200
            Width           =   732
         End
         Begin VB.TextBox txtNPSH 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   264
            Index           =   3
            Left            =   1080
            TabIndex        =   408
            Top             =   840
            Width           =   732
         End
         Begin VB.TextBox txtNPSH 
            Alignment       =   2  'Center
            Enabled         =   0   'False
            Height          =   264
            Index           =   0
            Left            =   1080
            TabIndex        =   403
            Top             =   480
            Width           =   732
         End
         Begin VB.Label lbltab4 
            Alignment       =   2  'Center
            Caption         =   "% TDH Var"
            Height          =   252
            Index           =   5
            Left            =   120
            TabIndex        =   414
            Top             =   1200
            Width           =   732
         End
         Begin VB.Label lbltab4 
            Alignment       =   1  'Right Justify
            Caption         =   "TDH (ft)"
            Height          =   252
            Index           =   4
            Left            =   120
            TabIndex        =   410
            Top             =   840
            Width           =   852
         End
         Begin VB.Label lbltab4 
            Alignment       =   1  'Right Justify
            Caption         =   "Flow (GPM)"
            Height          =   252
            Index           =   0
            Left            =   120
            TabIndex        =   407
            Top             =   480
            Width           =   852
         End
         Begin VB.Label lbltab4 
            Alignment       =   2  'Center
            Caption         =   "%"
            Height          =   252
            Index           =   2
            Left            =   2400
            TabIndex        =   406
            Top             =   240
            Width           =   492
         End
         Begin VB.Label lbltab4 
            Alignment       =   2  'Center
            Caption         =   "Start"
            Height          =   252
            Index           =   1
            Left            =   1080
            TabIndex        =   405
            Top             =   240
            Width           =   732
         End
         Begin VB.Label lbltab4 
            Alignment       =   1  'Right Justify
            Caption         =   "NPSHr"
            Height          =   252
            Index           =   3
            Left            =   1440
            TabIndex        =   404
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
         Left            =   -72120
         TabIndex        =   398
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
            TabIndex        =   400
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
            TabIndex        =   399
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
            TabIndex        =   401
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
         Left            =   13080
         Style           =   1  'Graphical
         TabIndex        =   390
         Top             =   2400
         Width           =   1332
      End
      Begin VB.TextBox txtRMA 
         Height          =   315
         Left            =   -69960
         TabIndex        =   386
         Top             =   540
         Visible         =   0   'False
         Width           =   1455
      End
      Begin MSComDlg.CommonDialog CommonDialog2 
         Left            =   -74880
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
         Left            =   -74760
         TabIndex        =   191
         Top             =   4680
         Visible         =   0   'False
         Width           =   14535
         Begin VB.ComboBox cmbTEMCNominalSuctionSize 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   243
            Top             =   600
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCNominalDischargeSize 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   242
            Top             =   240
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCVoltage 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   241
            Top             =   2400
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCDesignPressure 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   238
            Top             =   1320
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCCirculation 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   237
            Top             =   3120
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCModel 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   235
            Top             =   240
            Width           =   5445
         End
         Begin VB.TextBox txtTEMCFrameNumber 
            Height          =   315
            Left            =   1680
            TabIndex        =   215
            Top             =   1680
            Width           =   855
         End
         Begin VB.ComboBox cmbTEMCTRG 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   212
            Top             =   3120
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCPumpStages 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   210
            Top             =   2040
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCOtherMotor 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   208
            Top             =   2760
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCNominalImpSize 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   206
            Top             =   960
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCMaterials 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   204
            Top             =   960
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCJacketGasket 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   202
            Top             =   2400
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCInsulation 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   200
            Top             =   2040
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCImpellerType 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   198
            Top             =   1320
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCDivisionType 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   196
            Top             =   1680
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCAdditions 
            Height          =   315
            Left            =   8520
            Style           =   2  'Dropdown List
            TabIndex        =   194
            Top             =   2760
            Width           =   5445
         End
         Begin VB.ComboBox cmbTEMCAdapter 
            Height          =   315
            Left            =   1680
            Style           =   2  'Dropdown List
            TabIndex        =   192
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
            TabIndex        =   246
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
            TabIndex        =   245
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
            TabIndex        =   244
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
            TabIndex        =   240
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
            TabIndex        =   239
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
            TabIndex        =   236
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
            TabIndex        =   214
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
            TabIndex        =   213
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
            TabIndex        =   211
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
            TabIndex        =   209
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
            TabIndex        =   207
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
            TabIndex        =   205
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
            TabIndex        =   203
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
            TabIndex        =   201
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
            TabIndex        =   199
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
            TabIndex        =   197
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
            TabIndex        =   195
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
            TabIndex        =   193
            Top             =   643
            Width           =   855
         End
      End
      Begin VB.TextBox txtLineNumber 
         Height          =   315
         Left            =   -69720
         TabIndex        =   378
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
         Left            =   -74760
         TabIndex        =   352
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
            TabIndex        =   384
            Top             =   1200
            Visible         =   0   'False
            Width           =   6375
            Begin VB.TextBox txtNPSHFileLocation 
               Height          =   315
               Left            =   120
               TabIndex        =   385
               Top             =   240
               Width           =   5895
            End
         End
         Begin VB.TextBox txtLiquid 
            Height          =   315
            Left            =   2400
            TabIndex        =   372
            Top             =   1440
            Width           =   5415
         End
         Begin VB.TextBox txtJobNum 
            Height          =   315
            Left            =   12840
            TabIndex        =   371
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtSpGr 
            Height          =   315
            Left            =   5880
            TabIndex        =   367
            Top             =   990
            Width           =   1335
         End
         Begin VB.TextBox txtRatedInputPower 
            Height          =   315
            Left            =   2400
            TabIndex        =   366
            Top             =   1020
            Width           =   1335
         End
         Begin VB.TextBox txtLiquidTemperature 
            Height          =   315
            Left            =   12840
            TabIndex        =   365
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtNPSHr 
            Height          =   315
            Left            =   2400
            TabIndex        =   361
            Top             =   630
            Width           =   1335
         End
         Begin VB.TextBox txtThermalClass 
            Height          =   315
            Left            =   5880
            TabIndex        =   360
            Top             =   630
            Width           =   1335
         End
         Begin VB.TextBox txtExpClass 
            Height          =   315
            Left            =   9120
            TabIndex        =   359
            Top             =   600
            Width           =   1335
         End
         Begin VB.TextBox txtAmps 
            Height          =   315
            Left            =   5880
            TabIndex        =   356
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
            TabIndex        =   355
            Top             =   240
            Width           =   1335
         End
         Begin VB.TextBox txtNoPhases 
            Height          =   315
            Left            =   2400
            TabIndex        =   353
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
            TabIndex        =   374
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
            TabIndex        =   373
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
            TabIndex        =   370
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
            TabIndex        =   369
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
            TabIndex        =   368
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
            TabIndex        =   364
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
            TabIndex        =   363
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
            TabIndex        =   362
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
            TabIndex        =   358
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
            TabIndex        =   357
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
            TabIndex        =   354
            Top             =   240
            Width           =   852
         End
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Command1"
         Height          =   375
         Left            =   -67320
         TabIndex        =   351
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
         TabIndex        =   298
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
            TabIndex        =   350
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
            TabIndex        =   346
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
            TabIndex        =   345
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
            TabIndex        =   344
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
            TabIndex        =   343
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
            TabIndex        =   342
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
            TabIndex        =   341
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
            TabIndex        =   340
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
            TabIndex        =   339
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
            TabIndex        =   338
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
            TabIndex        =   337
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
            TabIndex        =   325
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
            TabIndex        =   324
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
            TabIndex        =   323
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
            TabIndex        =   322
            Top             =   690
            Width           =   255
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   6
            Left            =   840
            TabIndex        =   317
            Top             =   2520
            Width           =   735
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   7
            Left            =   3120
            TabIndex        =   316
            Top             =   2520
            Width           =   735
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   4
            Left            =   840
            TabIndex        =   315
            Top             =   1920
            Width           =   735
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   5
            Left            =   3120
            TabIndex        =   314
            Top             =   1920
            Width           =   735
         End
         Begin VB.ComboBox cmbTestAndInspection 
            Height          =   315
            Index           =   1
            ItemData        =   "Main.frx":1CE9
            Left            =   1680
            List            =   "Main.frx":1CF3
            TabIndex        =   313
            Top             =   2520
            Width           =   975
         End
         Begin VB.ComboBox cmbTestAndInspection 
            Height          =   315
            Index           =   0
            ItemData        =   "Main.frx":1D03
            Left            =   1680
            List            =   "Main.frx":1D0D
            TabIndex        =   312
            Top             =   1920
            Width           =   975
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   3
            Left            =   3120
            TabIndex        =   309
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   2
            Left            =   840
            TabIndex        =   307
            Top             =   1320
            Width           =   735
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   1
            Left            =   3120
            TabIndex        =   305
            Top             =   720
            Width           =   735
         End
         Begin VB.TextBox txtTestAndInspection 
            Alignment       =   1  'Right Justify
            Height          =   315
            Index           =   0
            Left            =   840
            TabIndex        =   299
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
            TabIndex        =   349
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
            TabIndex        =   348
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
            TabIndex        =   347
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
            TabIndex        =   336
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
            TabIndex        =   335
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
            TabIndex        =   334
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
            TabIndex        =   333
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
            TabIndex        =   332
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
            TabIndex        =   331
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
            TabIndex        =   330
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
            TabIndex        =   329
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
            TabIndex        =   328
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
            TabIndex        =   327
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
            TabIndex        =   326
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
            TabIndex        =   321
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
            TabIndex        =   320
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
            TabIndex        =   319
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
            TabIndex        =   318
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
            TabIndex        =   311
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
            TabIndex        =   310
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
            TabIndex        =   308
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
            TabIndex        =   306
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
            TabIndex        =   304
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
            TabIndex        =   303
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
            TabIndex        =   302
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
            TabIndex        =   301
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
            TabIndex        =   300
            Top             =   750
            Width           =   975
         End
      End
      Begin VB.TextBox txtImpellerDia 
         Height          =   315
         Left            =   -65160
         TabIndex        =   266
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
         Left            =   -63840
         Style           =   1  'Graphical
         TabIndex        =   264
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
         Left            =   -68880
         TabIndex        =   216
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
            TabIndex        =   218
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
            TabIndex        =   217
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
         Left            =   120
         TabIndex        =   66
         Top             =   3960
         Width           =   6735
         Begin VB.OptionButton optKW 
            Caption         =   "Use Ana In 4"
            Height          =   195
            Index           =   2
            Left            =   5280
            TabIndex        =   278
            Top             =   1680
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.OptionButton optKW 
            Caption         =   "Enter KW"
            Height          =   195
            Index           =   1
            Left            =   5280
            TabIndex        =   277
            Top             =   1440
            Visible         =   0   'False
            Width           =   1095
         End
         Begin VB.OptionButton optKW 
            Caption         =   "Add 3 powers"
            Height          =   195
            Index           =   0
            Left            =   5280
            TabIndex        =   276
            Top             =   1200
            Visible         =   0   'False
            Width           =   1335
         End
         Begin VB.CommandButton cmdFindMagtrols 
            Caption         =   "Find Magtrols"
            Height          =   255
            Left            =   5040
            TabIndex        =   189
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox cmbMagtrol 
            Height          =   315
            Left            =   2520
            TabIndex        =   187
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
            TabIndex        =   188
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
         TabIndex        =   184
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdApprovePump 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Approve/Unapprove This Pump"
         Height          =   615
         Left            =   -61920
         Style           =   1  'Graphical
         TabIndex        =   183
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
         TabIndex        =   180
         Top             =   720
         Visible         =   0   'False
         Width           =   1695
      End
      Begin VB.CommandButton cmdDeletePump 
         BackColor       =   &H00C0C0FF&
         Caption         =   "Delete This Pump"
         Height          =   615
         Left            =   -61920
         Style           =   1  'Graphical
         TabIndex        =   179
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
         Left            =   13080
         TabIndex        =   174
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
         TabIndex        =   173
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
         TabIndex        =   170
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
            TabIndex        =   172
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
            TabIndex        =   171
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
         TabIndex        =   167
         Top             =   6480
         Width           =   7215
         Begin VB.TextBox txtGGap 
            Alignment       =   1  'Right Justify
            Height          =   315
            Left            =   1200
            TabIndex        =   388
            Top             =   720
            Width           =   975
         End
         Begin VB.CommandButton cmdModifyBalanceHoleData 
            BackColor       =   &H00C0FFC0&
            Caption         =   "Modify Balance Hole Data"
            Height          =   495
            Left            =   5160
            Style           =   1  'Graphical
            TabIndex        =   297
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
            TabIndex        =   177
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
            TabIndex        =   176
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
            TabIndex        =   389
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
            TabIndex        =   178
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
            TabIndex        =   169
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
            TabIndex        =   168
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
         TabIndex        =   161
         Top             =   1440
         Width           =   3255
         Begin VB.TextBox txtVFDFreq 
            Height          =   315
            Left            =   2040
            TabIndex        =   375
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
            TabIndex        =   376
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
            TabIndex        =   164
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
            TabIndex        =   163
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
            TabIndex        =   162
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
         TabIndex        =   156
         Top             =   1440
         Width           =   4935
         Begin VB.ComboBox cmbMounting 
            Height          =   315
            Left            =   1920
            Style           =   2  'Dropdown List
            TabIndex        =   279
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
            TabIndex        =   280
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
            TabIndex        =   185
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
            TabIndex        =   166
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
            TabIndex        =   165
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
            TabIndex        =   160
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
            TabIndex        =   159
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
            TabIndex        =   158
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
            TabIndex        =   157
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
         TabIndex        =   148
         Top             =   1440
         Width           =   6135
         Begin VB.ComboBox cmbCirculationFlowMeter 
            Height          =   315
            Left            =   4560
            TabIndex        =   395
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cmbTemperatureTransducer 
            Height          =   315
            Left            =   1440
            TabIndex        =   394
            Top             =   1560
            Width           =   1335
         End
         Begin VB.ComboBox cmbDischargePressureTransducer 
            Height          =   315
            Left            =   1440
            TabIndex        =   393
            Top             =   1160
            Width           =   1335
         End
         Begin VB.ComboBox cmbSuctionPressureTransducer 
            Height          =   315
            Left            =   1440
            TabIndex        =   392
            Top             =   760
            Width           =   1335
         End
         Begin VB.ComboBox cmbFlowMeter 
            Height          =   315
            Left            =   1440
            TabIndex        =   391
            Top             =   360
            Width           =   1335
         End
         Begin VB.ComboBox cmbPLCNo 
            Height          =   315
            ItemData        =   "Main.frx":1D1D
            Left            =   4560
            List            =   "Main.frx":1D1F
            Style           =   2  'Dropdown List
            TabIndex        =   379
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
            TabIndex        =   380
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
            TabIndex        =   155
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
            TabIndex        =   154
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
            TabIndex        =   153
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
            TabIndex        =   152
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
            TabIndex        =   151
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
            TabIndex        =   150
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
            TabIndex        =   149
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
         Left            =   12960
         Style           =   1  'Graphical
         TabIndex        =   145
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
         Left            =   13080
         TabIndex        =   143
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
         Left            =   120
         TabIndex        =   130
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
            TabIndex        =   138
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
            TabIndex        =   137
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
            TabIndex        =   136
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
            TabIndex        =   135
            Top             =   720
            Width           =   1095
         End
         Begin VB.TextBox txtTemperature 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6360
            TabIndex        =   134
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtDischarge 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6360
            TabIndex        =   133
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtSuction 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6360
            TabIndex        =   132
            Top             =   600
            Visible         =   0   'False
            Width           =   255
         End
         Begin VB.TextBox txtFlow 
            Alignment       =   2  'Center
            Height          =   285
            Left            =   6360
            TabIndex        =   131
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
            Index           =   2
            Left            =   3320
            TabIndex        =   270
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
            Index           =   0
            Left            =   120
            TabIndex        =   268
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
            TabIndex        =   142
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
            TabIndex        =   141
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
            TabIndex        =   140
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
            TabIndex        =   139
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
         Left            =   6960
         TabIndex        =   112
         Top             =   1020
         Width           =   4095
         Begin VB.TextBox txtManualLamp 
            Height          =   285
            Left            =   2520
            TabIndex        =   125
            Text            =   "Text1"
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtWriteSP 
            Height          =   375
            Left            =   2880
            TabIndex        =   124
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtWriteSPData 
            Alignment       =   1  'Right Justify
            Height          =   285
            Left            =   1200
            TabIndex        =   123
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
            TabIndex        =   122
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
            TabIndex        =   121
            Top             =   300
            Width           =   615
         End
         Begin VB.TextBox txtValvePosition 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   3120
            TabIndex        =   120
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtDCoef 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2400
            TabIndex        =   119
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtICoef 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   3000
            TabIndex        =   118
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtPCoef 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2160
            TabIndex        =   117
            Top             =   1200
            Visible         =   0   'False
            Width           =   150
         End
         Begin VB.TextBox txtSetPoint 
            Alignment       =   2  'Center
            Height          =   375
            Left            =   2280
            TabIndex        =   116
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
            TabIndex        =   115
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
            TabIndex        =   114
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
            TabIndex        =   113
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
            TabIndex        =   129
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
            TabIndex        =   128
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
            TabIndex        =   127
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
            TabIndex        =   126
            Top             =   720
            Width           =   855
         End
      End
      Begin MSDataGridLib.DataGrid DataGrid2 
         Height          =   2415
         Left            =   8160
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
         Left            =   -73080
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
         TabIndex        =   108
         Top             =   660
         Width           =   1215
      End
      Begin VB.CommandButton cmdEnterPumpData 
         BackColor       =   &H00C0FFC0&
         Caption         =   "Enter Pump Data"
         Height          =   615
         Left            =   -65640
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
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   106
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
         Left            =   240
         TabIndex        =   97
         Top             =   6000
         Width           =   14535
         Begin VB.TextBox txtTEMCTRGReading 
            Height          =   285
            Left            =   7080
            TabIndex        =   383
            Top             =   480
            Width           =   855
         End
         Begin VB.TextBox txtVibAx 
            Height          =   285
            Left            =   6000
            TabIndex        =   382
            Top             =   480
            Width           =   855
         End
         Begin VB.Frame frmTEMCData 
            BackColor       =   &H00FFFFC0&
            Caption         =   "TEMC Data"
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
            TabIndex        =   247
            Top             =   120
            Visible         =   0   'False
            Width           =   6015
            Begin VB.TextBox txtRevHead 
               Height          =   285
               Left            =   2760
               TabIndex        =   396
               Top             =   960
               Width           =   855
            End
            Begin VB.TextBox txtTEMCPVValue 
               Height          =   285
               Left            =   4080
               TabIndex        =   260
               Top             =   960
               Width           =   855
            End
            Begin VB.TextBox txtTEMCCalcForce 
               Height          =   285
               Left            =   4080
               TabIndex        =   258
               Top             =   390
               Width           =   855
            End
            Begin VB.TextBox txtTEMCViscosity 
               Height          =   285
               Left            =   2760
               TabIndex        =   256
               Top             =   390
               Width           =   855
            End
            Begin VB.TextBox txtTEMCThrustRigPressure 
               Height          =   285
               Left            =   1560
               TabIndex        =   254
               Top             =   990
               Width           =   855
            End
            Begin VB.TextBox txtTEMCMomentArm 
               Height          =   285
               Left            =   1560
               TabIndex        =   251
               Top             =   390
               Width           =   855
            End
            Begin VB.TextBox txtTEMCRearThrust 
               Height          =   285
               Left            =   240
               TabIndex        =   250
               Top             =   990
               Width           =   855
            End
            Begin VB.TextBox txtTEMCFrontThrust 
               Height          =   285
               Left            =   240
               TabIndex        =   249
               Top             =   390
               Width           =   855
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
               TabIndex        =   397
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
               TabIndex        =   263
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
               TabIndex        =   262
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
               TabIndex        =   261
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
               TabIndex        =   259
               Top             =   180
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
               Height          =   210
               Index           =   52
               Left            =   2760
               TabIndex        =   257
               Top             =   180
               Width           =   855
            End
            Begin VB.Label lbltab2 
               Alignment       =   2  'Center
               BackColor       =   &H00FFFFC0&
               Caption         =   "Th Rig Pressure"
               BeginProperty Font 
                  Name            =   "Arial"
                  Size            =   8.25
                  Charset         =   0
                  Weight          =   700
                  Underline       =   0   'False
                  Italic          =   0   'False
                  Strikethrough   =   0   'False
               EndProperty
               Height          =   216
               Index           =   51
               Left            =   1560
               TabIndex        =   255
               Top             =   756
               Width           =   852
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
               TabIndex        =   253
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
               TabIndex        =   252
               Top             =   180
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
               Left            =   180
               TabIndex        =   248
               Top             =   750
               Width           =   975
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
            TabIndex        =   102
            Top             =   1080
            Width           =   855
         End
         Begin VB.TextBox txtThrustBal 
            Height          =   285
            Left            =   4680
            TabIndex        =   100
            Top             =   480
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
            TabIndex        =   381
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
            TabIndex        =   105
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
            TabIndex        =   104
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
            TabIndex        =   103
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label lbltab2 
            Alignment       =   1  'Right Justify
            BackColor       =   &H00FFFFC0&
            Caption         =   "Thrust Balance"
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
            TabIndex        =   101
            Top             =   240
            Width           =   1335
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
         Height          =   288
         Left            =   -73080
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   540
         Width           =   2055
      End
      Begin VB.TextBox txtRemarks 
         Height          =   555
         Left            =   -71640
         MaxLength       =   255
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Top             =   8700
         Width           =   7695
      End
      Begin VB.TextBox txtDesignTDH 
         Height          =   315
         Left            =   -63840
         TabIndex        =   6
         Top             =   1770
         Width           =   1335
      End
      Begin VB.TextBox txtDesignFlow 
         Height          =   315
         Left            =   -63840
         TabIndex        =   5
         Top             =   1410
         Width           =   1335
      End
      Begin VB.TextBox txtModelNo 
         Height          =   315
         Left            =   -73680
         TabIndex        =   4
         Top             =   4349
         Width           =   2532
      End
      Begin VB.TextBox txtShpNo 
         Height          =   315
         Left            =   -73560
         TabIndex        =   2
         Top             =   1410
         Width           =   4935
      End
      Begin VB.TextBox txtBilNo 
         Height          =   315
         Left            =   -73560
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
         Left            =   120
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
            TabIndex        =   288
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
            TabIndex        =   287
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
            TabIndex        =   286
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
            Index           =   3
            Left            =   1952
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   284
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
            TabIndex        =   283
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
            Index           =   0
            Left            =   352
            Locked          =   -1  'True
            MaxLength       =   15
            TabIndex        =   281
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
         Left            =   120
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
            TabIndex        =   296
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
            TabIndex        =   295
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
            TabIndex        =   294
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
            TabIndex        =   293
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
            TabIndex        =   292
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
            TabIndex        =   291
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
            TabIndex        =   290
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
            TabIndex        =   289
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
            TabIndex        =   275
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
            Index           =   6
            Left            =   3320
            TabIndex        =   274
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
            Index           =   5
            Left            =   1720
            TabIndex        =   273
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
            TabIndex        =   272
            Top             =   720
            Width           =   300
         End
      End
      Begin VB.ComboBox cmbPLCLoop 
         Height          =   288
         ItemData        =   "Main.frx":1D25
         Left            =   2520
         List            =   "Main.frx":1D27
         Style           =   2  'Dropdown List
         TabIndex        =   46
         Top             =   420
         Width           =   2895
      End
      Begin MSDataGridLib.DataGrid DataGrid1 
         Height          =   2415
         Left            =   2040
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
         Left            =   -74760
         TabIndex        =   219
         Top             =   5040
         Width           =   14415
         Begin VB.ComboBox cmbCirculationPath 
            Height          =   315
            ItemData        =   "Main.frx":1D29
            Left            =   1800
            List            =   "Main.frx":1D2B
            Style           =   2  'Dropdown List
            TabIndex        =   229
            Top             =   1860
            Width           =   3615
         End
         Begin VB.ComboBox cmbStatorFill 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   228
            Top             =   1140
            Width           =   3615
         End
         Begin VB.ComboBox cmbDesignPressure 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   227
            Top             =   1500
            Width           =   3615
         End
         Begin VB.ComboBox cmbRPM 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   226
            Top             =   780
            Width           =   3615
         End
         Begin VB.ComboBox cmbMotor 
            Height          =   315
            Left            =   1800
            Style           =   2  'Dropdown List
            TabIndex        =   225
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
            TabIndex        =   220
            Top             =   360
            Width           =   5175
            Begin VB.ComboBox cmbModelGroup 
               Height          =   315
               Left            =   1440
               Style           =   2  'Dropdown List
               TabIndex        =   222
               Top             =   960
               Width           =   3615
            End
            Begin VB.ComboBox cmbModel 
               Height          =   315
               Left            =   1440
               Style           =   2  'Dropdown List
               TabIndex        =   221
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
               TabIndex        =   224
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
               TabIndex        =   223
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
            TabIndex        =   234
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
            TabIndex        =   233
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
            TabIndex        =   232
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
            TabIndex        =   231
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
            TabIndex        =   230
            Top             =   1920
            Width           =   1575
         End
      End
      Begin MSComCtl2.UpDown UpDown2 
         Height          =   504
         Left            =   13920
         TabIndex        =   428
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
         Left            =   120
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
         Left            =   -68520
         TabIndex        =   423
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
         Left            =   -74880
         TabIndex        =   421
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
         Left            =   -70800
         TabIndex        =   418
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
         Left            =   -63240
         TabIndex        =   417
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
         Left            =   -65640
         TabIndex        =   416
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
         TabIndex        =   387
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
         Left            =   -71040
         TabIndex        =   377
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
         Left            =   -67320
         TabIndex        =   267
         Top             =   4380
         Width           =   2052
      End
      Begin VB.Label lbltab2 
         Alignment       =   2  'Center
         Caption         =   "Number of Points to Plot"
         Height          =   375
         Index           =   54
         Left            =   12840
         TabIndex        =   186
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
         Left            =   13080
         TabIndex        =   175
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
         TabIndex        =   147
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
         Left            =   13080
         TabIndex        =   144
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
         TabIndex        =   110
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
         Left            =   -74280
         TabIndex        =   109
         Top             =   450
         Width           =   1095
      End
      Begin VB.Label lbltab2 
         Alignment       =   2  'Center
         Caption         =   "Test Number"
         Height          =   255
         Index           =   63
         Left            =   480
         TabIndex        =   107
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
         Left            =   -72600
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
         Left            =   -65280
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
         Left            =   -65160
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
         Left            =   -74760
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
         Left            =   -74520
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
         Left            =   -74520
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
         Left            =   840
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
      TabIndex        =   182
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
      TabIndex        =   181
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
      TabIndex        =   146
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
165            Me.cmbCirculationFlowMeter.ListIndex = rsTransducers.Fields("CircFlowMeter")
166            Me.cmbPLCNo.ListIndex = rsTransducers.Fields("PLC")
167            Me.cmbAnalyzerNo.ListIndex = rsTransducers.Fields("Analyzer")
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
205            txtExpClass.Text = rsSupermarketModel.Fields("EXPClass")
206            txtLiquid.Text = rsSupermarketModel.Fields("Liquid")
207            txtLiquidTemperature.Text = rsSupermarketModel.Fields("LiquidTemp")
208            txtJobNum.Text = JobNumber
209            txtImpellerDia.Text = rsSupermarketModel.Fields("ImpellerDiameter")
210            txtModelNo.Text = rsSupermarketModel.Fields("Model")
211            txtRVSPartNo.Text = rsSupermarketModel.Fields("RVSPartNo")
212            cmdSelectSupermarket.Caption = "Save Data"
213            If UCase(rsSupermarketModel.Fields("Feathered")) = "FEATHERED" Then
214                Me.chkSuperMarketFeathered.value = Checked
215            End If
216        End If
217        grpSupermarket.Visible = False

' <VB WATCH>
218        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
219        Exit Sub
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
220        On Error GoTo vbwErrHandler
221        Const VBWPROCNAME = "frmPLCData.cmbPLCNo_Click"
222        If vbwProtector.vbwTraceProc Then
223            Dim vbwProtectorParameterString As String
224            If vbwProtector.vbwTraceParameters Then
225                vbwProtectorParameterString = "()"
226            End If
227            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
228        End If
' </VB WATCH>

229        Dim I As Integer
230        Dim PLCNo As Integer
231        Dim MagtrolNo As String

232        PLCNo = 0
233        If InStr(cmbPLCNo.Text, "8") > 0 Then
234            PLCNo = 8
235            MagtrolNo = "GPIB6"
236        End If
237        If InStr(cmbPLCNo.Text, "9") > 0 Then
238            PLCNo = 9
239            MagtrolNo = "GPIB5"
240        End If

241        For I = 0 To cmbPLCLoop.ListCount - 1                     'go through the combobox entries
242            If InStr(cmbPLCLoop.List(I), PLCNo) > 0 Then   'see when we find the desired index number
243                cmbPLCLoop.ListIndex = I                                              'if we do, set the combo box
244                Exit For                                            'and we're done
245            End If
               'cmbPLCLoop.ListIndex = -1                             'else, remove any pointer
246        Next I

247        For I = 0 To cmbMagtrol.ListCount - 1
248            If InStr(cmbMagtrol.List(I), MagtrolNo) > 0 Then   'see when we find the desired index number
249                cmbMagtrol.ListIndex = I                                              'if we do, set the combo box
250                Exit For                                            'and we're done
251            End If
252        Next I
' <VB WATCH>
253        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
254        Exit Sub
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
255        On Error GoTo vbwErrHandler
256        Const VBWPROCNAME = "frmPLCData.cmbVoltage_click"
257        If vbwProtector.vbwTraceProc Then
258            Dim vbwProtectorParameterString As String
259            If vbwProtector.vbwTraceParameters Then
260                vbwProtectorParameterString = "()"
261            End If
262            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
263        End If
' </VB WATCH>
264        If Me.cmbVoltage.ListIndex = 0 Then
265            Me.cmbFrequency.ListIndex = 2
266        End If
' <VB WATCH>
267        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
268        Exit Sub
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
269        On Error GoTo vbwErrHandler
270        Const VBWPROCNAME = "frmPLCData.cmbMagtrol_Click"
271        If vbwProtector.vbwTraceProc Then
272            Dim vbwProtectorParameterString As String
273            If vbwProtector.vbwTraceParameters Then
274                vbwProtectorParameterString = "()"
275            End If
276            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
277        End If
' </VB WATCH>
278        Dim I As Integer
279        Dim sSendStr As String
280        Dim sGPIBName As String
281        Dim MagtrolName As String

282        I = cmbMagtrol.ItemData(cmbMagtrol.ListIndex)
283        sGPIBName = "GPIB" & I
284        MagtrolName = cmbMagtrol.List(cmbMagtrol.ListIndex)

285        If I = 99 Then      'manual entry
286            boMagtrolOperating = False
287            EnableMagtrolFields
' <VB WATCH>
288        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
289            Exit Sub
290        Else
291            boMagtrolOperating = True
292        End If

293        SetupMagtrols MagtrolName, I

' <VB WATCH>
294        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
295        Exit Sub
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
296        On Error GoTo vbwErrHandler
297        Const VBWPROCNAME = "frmPLCData.cmbPLCLoop_Click"
298        If vbwProtector.vbwTraceProc Then
299            Dim vbwProtectorParameterString As String
300            If vbwProtector.vbwTraceParameters Then
301                vbwProtectorParameterString = "()"
302            End If
303            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
304        End If
' </VB WATCH>

305        Dim RetVal As String

           'manual data entry selection
306        If cmbPLCLoop.ListIndex = cmbPLCLoop.ListCount - 1 Then 'no plc
307            boPLCOperating = False
308            EnablePLCFields
309            If DeviceOpen = True Then
310                RetVal = DisconnectPLC()
311            End If
' <VB WATCH>
312        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
313            Exit Sub
314        End If

315        If DeviceOpen = True Then
316            RetVal = DisconnectPLC()
317        End If

318        RetVal = ConnectToPLC(cmbPLCLoop.ItemData(cmbPLCLoop.ListIndex))
319        If RetVal <> 0 Then
320            MsgBox ("Can't connect to PLC - " & Description(cmbPLCLoop.ListIndex))
321            boPLCOperating = False
322            EnablePLCFields
323        Else
324            boPLCOperating = True
325            tDevice = cmbPLCLoop.ItemData(cmbPLCLoop.ListIndex)
326            DisablePLCFields
327        End If
' <VB WATCH>
328        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
329        Exit Sub
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
330        On Error GoTo vbwErrHandler
331        Const VBWPROCNAME = "frmPLCData.cmbTestDate_Click"
332        If vbwProtector.vbwTraceProc Then
333            Dim vbwProtectorParameterString As String
334            If vbwProtector.vbwTraceParameters Then
335                vbwProtectorParameterString = "()"
336            End If
337            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
338        End If
' </VB WATCH>

339        Dim sName As String
340        Dim sParam As String
341        Dim I As Integer
342        Dim j As Integer
343        Dim k As Integer
344        Dim bSk As Boolean
345        Dim sBC As Single
346        Dim NOK() As Long

347        cmdModifyBalanceHoleData.Visible = False


348        If Not boFoundTestSetup Then    'if we don't have any TestSetup data written
349            boFoundTestData = False
' <VB WATCH>
350        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
351            Exit Sub
352        End If


           'select the testsetup data for the serial number
353        qyTestSetup.ActiveConnection = cnPumpData
354        qyTestSetup.CommandText = "SELECT * " & _
                         "From TempTestSetupData " & _
                         "Where (((TempTestSetupData.SerialNumber) = '" & txtSN.Text & "') AND TempTestSetupData.Date = #" & cmbTestDate.List(cmbTestDate.ListIndex) & "#) " & _
                         "ORDER BY TempTestSetupData.Date;"

355        If rsTestSetup.State = adStateOpen Then
356            rsTestSetup.Close
357        End If

358        With rsTestSetup     'open the recordset for the query
       '        .Index = "FindData"
359            .CursorLocation = adUseClient
360            .CursorType = adOpenStatic
361            .Open qyTestSetup
362        End With

           'move to the selected date
363        If Not rsTestSetup.BOF Then
364            rsTestSetup.MoveFirst
365        End If
       '
           'show the correct combo box entries for this record
           'SetComboTestSetup cmbOrificeNumber, "OrificeNumber", "OrificeNumber", rsTestSetup
366        SetComboTestSetup cmbTestSpec, "TestSpec", "TestSpecification", rsTestSetup
367        SetComboTestSetup cmbLoopNumber, "LoopNumber", "LoopNumber", rsTestSetup
368        SetComboTestSetup cmbSuctDia, "SuctDiam", "SuctionDiameter", rsTestSetup
369        SetComboTestSetup cmbDischDia, "DischDiam", "DischargeDiameter", rsTestSetup
370        SetComboTestSetup cmbTachID, "TachID", "TachID", rsTestSetup
371        SetComboTestSetup cmbAnalyzerNo, "AnalyzerNo", "AnalyzerNo", rsTestSetup
372        SetComboTestSetup cmbVoltage, "Voltage", "Voltage", rsTestSetup
373        SetComboTestSetup cmbFrequency, "Frequency", "Frequency", rsTestSetup
374        SetComboTestSetup cmbMounting, "Mounting", "Mounting", rsTestSetup
375        SetComboTestSetup cmbPLCNo, "PLCNo", "PLCNo", rsTestSetup
376        SetComboTestSetup cmbFlowMeter, "FlowMeterID", "PumpFlowMeter", rsTestSetup
377        SetComboTestSetup cmbSuctionPressureTransducer, "SuctionID", "SuctionPressureTransducer", rsTestSetup
378        SetComboTestSetup cmbDischargePressureTransducer, "DischID", "DischargePressureTransducer", rsTestSetup
379        SetComboTestSetup cmbTemperatureTransducer, "TemperatureID", "TemperatureTransducer", rsTestSetup
380        SetComboTestSetup cmbCirculationFlowMeter, "MagFlowID", "CirculationFlowMeter", rsTestSetup

381        sName = "HDCor"
382        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
383            sParam = rsTestSetup.Fields(sName)
384        Else
385            sParam = vbNullString
386        End If
387        txtHDCor.Text = sParam

388        sName = "KWMult"
389        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
390            sParam = rsTestSetup.Fields(sName)
391        Else
392            sParam = vbNullString
393        End If
394        txtKWMult.Text = sParam

395        sName = "Who"
396        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
397            sParam = rsTestSetup.Fields(sName)
398        Else
399            sParam = vbNullString
400        End If
401        txtWho.Text = sParam

402        sName = "RMA"
403        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
404            sParam = rsTestSetup.Fields(sName)
405        Else
406            sParam = vbNullString
407        End If
408        txtRMA.Text = sParam

409        sName = "Remarks"
410        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
411            sParam = rsTestSetup.Fields(sName)
412        Else
413            sParam = vbNullString
414        End If
415        txtTestSetupRemarks.Text = sParam

416        sName = "VFDFrequency"
417        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
418            sParam = rsTestSetup.Fields(sName)
419        Else
420            sParam = vbNullString
421        End If
422        txtVFDFreq.Text = sParam

423        sName = "SuctionGageHeight"
424        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
425            sParam = rsTestSetup.Fields(sName)
426        Else
427            sParam = 0
428        End If
429        txtSuctHeight.Text = sParam

430        sName = "DischargeGageHeight"
431        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
432            sParam = rsTestSetup.Fields(sName)
433        Else
434            sParam = 0
435        End If
436        txtDischHeight.Text = sParam

437        sName = "EndPlay"
438        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
439            sParam = rsTestSetup.Fields(sName)
440        Else
441            sParam = vbNullString
442        End If
443        txtEndPlay.Text = sParam

444        sName = "GGAP"
445        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
446            sParam = rsTestSetup.Fields(sName)
447        Else
448            sParam = vbNullString
449        End If
450        txtGGap.Text = sParam

451        sName = "OtherMods"
452        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
453            sParam = rsTestSetup.Fields(sName)
454        Else
455            sParam = vbNullString
456        End If
457        txtOtherMods.Text = sParam

458        If rsTestSetup.Fields("ImpFeathered") Then
459            chkFeathered.value = 1
460        Else
461            chkFeathered.value = 0
462        End If

463        If Val(rsTestSetup.Fields("ImpTrimmed")) = 0 Then
464            chkTrimmed.value = 0
465            txtImpTrim.Visible = False
466            txtImpTrim.Text = rsTestSetup.Fields("Imptrimmed")
467        Else
468            chkTrimmed.value = 1
469            txtImpTrim.Visible = True
470            txtImpTrim.Text = rsTestSetup.Fields("Imptrimmed")
471        End If

472        If Val(rsTestSetup.Fields("PumpDischOrifice")) = 0 Then
473            chkOrifice.value = 0
474            txtOrifice.Visible = False
475        Else
476            chkOrifice.value = 1
477            txtOrifice.Visible = True
478            txtOrifice.Text = rsTestSetup.Fields("PumpDischOrifice")
479        End If

480        If Val(rsTestSetup.Fields("CircFlowOrifice")) = 0 Then
481            chkCircOrifice.value = 0
482            txtCircOrifice.Visible = False
483        Else
484            chkCircOrifice.value = 1
485            txtCircOrifice.Visible = True
486            txtCircOrifice.Text = rsTestSetup.Fields("CircFlowOrifice")
487        End If

488        If (IsNull(rsTestSetup.Fields("NPSHFile"))) Or (LenB(rsTestSetup.Fields("NPSHFile")) = 0) Then
489            chkNPSH.value = 0
490            txtNPSHFile.Visible = False
491        Else
492            chkNPSH.value = 1
493            txtNPSHFile.Visible = True
494            txtNPSHFile.Text = rsTestSetup.Fields("NPSHFile")
495        End If

496        If (IsNull(rsTestSetup.Fields("PictureFile"))) Or (LenB(rsTestSetup.Fields("PictureFile")) = 0) Then
497            chkPictures.value = 0
498            txtPicturesFile.Visible = False
499        Else
500            chkPictures.value = 1
501            txtPicturesFile.Visible = True
502            txtPicturesFile.Text = rsTestSetup.Fields("PictureFile")
503        End If

504        If (IsNull(rsTestSetup.Fields("VibrationFile"))) Or (LenB(rsTestSetup.Fields("VibrationFile")) = 0) Then
505            chkVibration.value = 0
506            txtVibrationFile.Visible = False
507        Else
508            chkVibration.value = 1
509            txtVibrationFile.Visible = True
510            txtVibrationFile.Text = rsTestSetup.Fields("VibrationFile")
511        End If


           'for TEMC Inspection Report
512        sName = "InsulationMeggerVolts"
513        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
514            sParam = rsTestSetup.Fields(sName)
515        Else
516            sParam = 0
517        End If
518        txtTestAndInspection(0).Text = sParam

519        sName = "InsulationMegOhms"
520        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
521            sParam = rsTestSetup.Fields(sName)
522        Else
523            sParam = 0
524        End If
525        txtTestAndInspection(1).Text = sParam

526        sName = "DielectricVolts"
527        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
528            sParam = rsTestSetup.Fields(sName)
529        Else
530            sParam = 0
531        End If
532        txtTestAndInspection(2).Text = sParam

533        sName = "DielectricTime"
534        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
535            sParam = rsTestSetup.Fields(sName)
536        Else
537            sParam = 0
538        End If
539        txtTestAndInspection(3).Text = sParam

540        sName = "HydrostaticValue"
541        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
542            sParam = rsTestSetup.Fields(sName)
543        Else
544            sParam = 0
545        End If
546        txtTestAndInspection(4).Text = sParam

547        sName = "HydrostaticTime"
548        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
549            sParam = rsTestSetup.Fields(sName)
550        Else
551            sParam = 0
552        End If
553        txtTestAndInspection(5).Text = sParam

554        sName = "PneumaticValue"
555        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
556            sParam = rsTestSetup.Fields(sName)
557        Else
558            sParam = 0
559        End If
560        txtTestAndInspection(6).Text = sParam

561        sName = "PneumaticTime"
562        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
563            sParam = rsTestSetup.Fields(sName)
564        Else
565            sParam = 0
566        End If
567        txtTestAndInspection(7).Text = sParam

568        For I = 0 To cmbTestAndInspection(0).ListCount - 1
569            If cmbTestAndInspection(0).Text = rsTestSetup.Fields("HydrostaticUnits") Then
570                    cmbTestAndInspection(0).ListIndex = I
571                    Exit For
572            End If
573            cmbTestAndInspection(0).ListIndex = -1
574        Next I


575        For I = 0 To cmbTestAndInspection(1).ListCount - 1
576            If cmbTestAndInspection(1).Text = rsTestSetup.Fields("PneumaticUnits") Then
577                    cmbTestAndInspection(1).ListIndex = I
578                    Exit For
579            End If
580            cmbTestAndInspection(1).ListIndex = -1
581        Next I

582        TestAndInspectionGood(0).value = Abs(rsTestSetup!insulationgood)
583        TestAndInspectionGood(1).value = Abs(rsTestSetup!DielectricGood)
584        TestAndInspectionGood(2).value = Abs(rsTestSetup!HydrostaticGood)
585        TestAndInspectionGood(3).value = Abs(rsTestSetup!PneumaticGood)
586        TestAndInspectionGood(4).value = Abs(rsTestSetup!GeneralAppearanceGood)
587        TestAndInspectionGood(5).value = Abs(rsTestSetup!OutlineDimensionsGood)
588        TestAndInspectionGood(6).value = Abs(rsTestSetup!MotorNoLoadTestGood)
589        TestAndInspectionGood(7).value = Abs(rsTestSetup!MotorLockedRotorTestGood)
590        TestAndInspectionGood(8).value = Abs(rsTestSetup!HydrostaticTestGood)
591        TestAndInspectionGood(9).value = Abs(rsTestSetup!HydraulicTestGood)
592        TestAndInspectionGood(10).value = Abs(rsTestSetup!NPSHTestGood)
593        TestAndInspectionGood(11).value = Abs(rsTestSetup!CleanPurgeSealGood)
594        TestAndInspectionGood(12).value = Abs(rsTestSetup!PaintCheckGood)
595        TestAndInspectionGood(13).value = Abs(rsTestSetup!NameplateGood)
596        TestAndInspectionGood(14).value = Abs(rsTestSetup!SupervisorApproval)

597        GetBalanceHoleData frmPLCData.txtSN.Text, cmbTestDate.Text

598         If rsBalanceHoles.RecordCount = 0 Then
599            chkBalanceHoles.value = 0
600            dgBalanceHoles.Visible = False
601            boGotBalanceHoles = False
602        Else
603            boGotBalanceHoles = True
604            ReDim NOK(rsBalanceHoles.RecordCount)
605            rsBalanceHoles.MoveLast
606            For I = 1 To rsBalanceHoles.RecordCount
607                NOK(I) = 0
608            Next I

609            For j = 1 To rsBalanceHoles.RecordCount - 1
610                rsBalanceHoles.MoveFirst
611                rsBalanceHoles.Move rsBalanceHoles.RecordCount - j
612                sBC = rsBalanceHoles.Fields("BoltCircle")
613                bSk = False
614                For k = 1 To rsBalanceHoles.RecordCount
615                    If NOK(k) = rsBalanceHoles.Fields(0) Then
616                        bSk = True
617                    End If
618                Next k
619                If Not bSk Then
620                    For I = rsBalanceHoles.RecordCount - j To 1 Step -1
621                        rsBalanceHoles.MovePrevious
622                        If rsBalanceHoles.Fields("BoltCircle") = sBC Then
623                            NOK(I) = rsBalanceHoles.Fields(0)
624                        End If
625                    Next I
626                End If
627            Next j

628            Dim sFilt As String
629            sFilt = ""
630            For I = 1 To rsBalanceHoles.RecordCount
631                If NOK(I) <> 0 Then
632                    sFilt = sFilt & "(BalanceHoleID <> " & NOK(I) & ") AND "
       '                sFilt = sFilt & "(" & rsBalanceHoles.Filter & " AND BalanceHoleID <> " & NOK(I) & ") AND "
633                End If
634            Next I

635            If Len(sFilt) > 4 Then
636                sFilt = Left(sFilt, Len(sFilt) - 4)
637                rsBalanceHoles.Filter = sFilt
638            End If

639            chkBalanceHoles.value = 1
640            dgBalanceHoles.Visible = True
641        End If
       '
           'set the test date filter for the test data
642        rsTestData.Filter = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"

643        If rsTestData.RecordCount = 0 Then
644            boFoundTestData = False
645            AddTestData
646            EnableTestDataControls
647            MsgBox "No Test Data Exists for this Serial Number"
648        Else
649            boFoundTestData = True
650            DisableTestDataControls                         'if it's in the real database, don't allow changes here
651        End If

652        If Not boTestDateIsApproved Then    'data approved?
653            EnableTestDataControls
654        End If

655        If rsTestSetup.Fields("Approved") = True Then
656            DisableTestDataControls                         'if it's in the real database, don't allow changes here
657            lblTestDateApproved.Visible = True
658            MsgBox ("Found pump.  Data cannot be modified.")
659            If boCanApprove Then
660                cmdApproveTestDate.Caption = "Unapprove this Test Date"
661            End If
662        Else
663            EnableTestDataControls                          'it's in the temp database, allow changes
664            lblTestDateApproved.Visible = False
665            If boPumpIsApproved = True Then
666                MsgBox ("Found pump.  Pump data cannot be modified, but test setup data and test data can be modified.")
667            Else
668                MsgBox ("Found pump.  Pump data, test setup data, and test data can be modified.")
669            End If
670            If boCanApprove Then
671                If rsPumpData.Fields("Approved") = True Then
672                    cmdApproveTestDate.Enabled = True
673                    cmdApproveTestDate.Caption = "Approve this Test Date"
674                Else
675                    cmdApproveTestDate.Caption = "You Must Approve Pump First"
676                    cmdApproveTestDate.Enabled = False
677                End If
678            End If
679        End If

680        rsEff.MoveFirst
681        rsTestData.MoveFirst

682        For I = 1 To rsTestData.RecordCount
683            DoEfficiencyCalcs
684            rsEff.MoveNext
685            rsTestData.MoveNext
686        Next I

          ' fix the datagrid
687       Set DataGrid1.DataSource = rsTestData
688       Set DataGrid2.DataSource = rsEff

689       Dim c As Column
690       For Each c In DataGrid1.Columns
691          Select Case c.DataField
             Case "TestDataID"     'Hide some columns
692             c.Visible = False
693          Case "SerialNumber"
694             c.Visible = False
695          Case "Date"
696             c.Visible = False
697          Case Else             ' Show all other columns.
698             c.Visible = True
699             c.Alignment = dbgRight
700          End Select
701        Next c

702        For Each c In DataGrid2.Columns
703            c.Alignment = dbgCenter
704            c.Width = 750
705            Select Case c.ColIndex
                   Case 1
706                    c.Caption = "Flow"
707                    c.NumberFormat = "###0.00"
708                Case 2
709                    c.Caption = "TDH"
710                    c.NumberFormat = "##0.00"
711                Case 3
712                    c.Caption = "Input Pwr"
713                    c.NumberFormat = "##0.00"
714                    c.Width = 850
715                Case 4
716                    c.Caption = "Voltage"
717                    c.NumberFormat = "##0.00"
718                Case 5
719                    c.Caption = "Current"
720                    c.NumberFormat = "##0.00"
721                Case 6
722                    c.Caption = "Overall Eff"
723                    c.NumberFormat = "##0.00"
724                    c.Width = 850
725                Case 7
726                    c.Caption = "NPSHr"
727                    c.NumberFormat = "#0.00"
728                Case Else
729                    c.Visible = False
730            End Select
731        Next c
732            FixPointsToPlot

733        txtUpDn1.Text = 1

       'unlock the text boxes
734        For I = 0 To 7
735            txtTitle(I).Locked = False
736        Next I

737        For I = 20 To 27
738            txtTitle(I).Locked = False
739        Next I

       'look for titles for TCs and AIs
740        Dim qy As New ADODB.Command
741        Dim rs As New ADODB.Recordset

742        qy.ActiveConnection = cnPumpData

           'see if we have an entry in the table
743        qy.CommandText = "SELECT * FROM AITitles " & _
                             "WHERE (((AITitles.SerialNo)= '" & txtSN.Text & "') " & _
                             "AND ((AITitles.Date)= #" & cmbTestDate.Text & "#)); "

744        With rs     'open the recordset for the query
745            .CursorLocation = adUseClient
746            .CursorType = adOpenStatic
747            .LockType = adLockOptimistic
748            .Open qy
749        End With

750        If Not (rs.BOF = True And rs.EOF = True) Then   'update titles
751            rs.MoveFirst
752            Do While Not rs.EOF
753                txtTitle(rs.Fields("Channel")).Text = rs.Fields("Title")
754                rs.MoveNext
755            Loop
756        End If

757        rs.Close
758        Set rs = Nothing
759        Set qy = Nothing
' <VB WATCH>
760        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
761        Exit Sub
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
762        On Error GoTo vbwErrHandler
763        Const VBWPROCNAME = "frmPLCData.cmdAddNewBalanceHoles_Click"
764        If vbwProtector.vbwTraceProc Then
765            Dim vbwProtectorParameterString As String
766            If vbwProtector.vbwTraceParameters Then
767                vbwProtectorParameterString = "()"
768            End If
769            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
770        End If
' </VB WATCH>
771        Dim strInput As String
772        Dim I As Integer
773        Dim sNumber As Integer
774        Dim sDia As Single
775        Dim sBC As Single

           'get the data for the balance holes
776        strInput = InputBox("Enter Number of Holes")
777        If strInput <> "" Then
778            sNumber = CInt(strInput)
779        Else
780            GoTo CancelPressed
781        End If

782        strInput = InputBox("Enter Decimal Value of Hole Diameter or Slot (For Example, 0.675) ")
783        If strInput <> "" Then
784            If UCase(strInput) = "SLOT" Then
785                strInput = 99
786            End If
787            sDia = CSng(strInput)
788        Else
789            GoTo CancelPressed
790        End If

791        strInput = InputBox("Enter Decimal Value of Bolt Circle or Unknown (For Example, 4.525)")
792        If strInput <> "" Then
793            If UCase(strInput) = "UNKNOWN" Then
794                strInput = 99
795            End If
796            sBC = CSng(strInput)
797        Else
798            GoTo CancelPressed
799        End If

800        GetBalanceHoleData frmPLCData.txtSN.Text, cmbTestDate.Text

801        rsBalanceHoles.AddNew
802        rsBalanceHoles!SerialNo = txtSN.Text
803        rsBalanceHoles!Date = cmbTestDate.Text
804        rsBalanceHoles!Number = sNumber
805        rsBalanceHoles!diameter = sDia
806        rsBalanceHoles!boltcircle = sBC

807        rsBalanceHoles.Update

808        GetBalanceHoleData txtSN.Text, cmbTestDate.Text
809        rsBalanceHoles.MoveLast
810        dgBalanceHoles.Refresh
811        chkBalanceHoles.value = 1

' <VB WATCH>
812        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
813        Exit Sub

814    CancelPressed:
815        MsgBox "No New Balance Hole Data Entered", vbOKOnly
' <VB WATCH>
816        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
817        Exit Sub
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
818        On Error GoTo vbwErrHandler
819        Const VBWPROCNAME = "frmPLCData.cmdAddNewTestDate_Click"
820        If vbwProtector.vbwTraceProc Then
821            Dim vbwProtectorParameterString As String
822            If vbwProtector.vbwTraceParameters Then
823                vbwProtectorParameterString = "()"
824            End If
825            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
826        End If
' </VB WATCH>
827        Dim I As Integer

828        chkFeathered.value = chkSuperMarketFeathered.value

829        For I = 1 To cmbTestDate.ListCount      'see if we already have today's date entered
830            If cmbTestDate.List(I) = Date Then
831                MsgBox "There is already an entry for today.  You can only have one entry for each Serial Number and a given date.  You may want to modify the Serial Number.", vbOKOnly
' <VB WATCH>
832        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
833                Exit Sub
834            End If
835        Next I

           'we didn't find today's date entered, allow data entry
836        boFoundTestSetup = False

837        EnableTestSetupDataControls
838        Pressed = False
839        cmdEnterTestSetupData_Click
840        cmdAddNewBalanceHoles.Visible = True
841        txtWho.Text = LogInInitials
842        MsgBox "New Test Date Added - " & cmbTestDate.List(cmbTestDate.ListCount - 1), vbOKOnly, "Added New Test Date"
' <VB WATCH>
843        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
844        Exit Sub
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
845        On Error GoTo vbwErrHandler
846        Const VBWPROCNAME = "frmPLCData.cmdApprovePump_Click"
847        If vbwProtector.vbwTraceProc Then
848            Dim vbwProtectorParameterString As String
849            If vbwProtector.vbwTraceParameters Then
850                vbwProtectorParameterString = "()"
851            End If
852            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
853        End If
' </VB WATCH>
854        rsPumpData.Fields("Approved") = Not rsPumpData.Fields("Approved")
855        rsPumpData.Update
856        rsPumpData.Requery
857        lblPumpApproved.Visible = rsPumpData.Fields("Approved")
858        If rsPumpData.Fields("Approved") = True Then
859            cmdApprovePump.Caption = "Unapprove This Pump"
860            cmdApproveTestDate.Enabled = True
861            If rsTestSetup.Fields("Approved") = True Then
862                cmdApproveTestDate.Caption = "Unapprove This Test Date"
863            Else
864                cmdApproveTestDate.Caption = "Approve This Test Date"
865            End If
866        Else
867            cmdApprovePump.Caption = "Approve This Pump"
868            cmdApproveTestDate.Caption = "You Must Approve Pump First"
869            cmdApproveTestDate.Enabled = False
870        End If
' <VB WATCH>
871        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
872        Exit Sub
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
873        On Error GoTo vbwErrHandler
874        Const VBWPROCNAME = "frmPLCData.cmdApproveTestDate_Click"
875        If vbwProtector.vbwTraceProc Then
876            Dim vbwProtectorParameterString As String
877            If vbwProtector.vbwTraceParameters Then
878                vbwProtectorParameterString = "()"
879            End If
880            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
881        End If
' </VB WATCH>
882        rsTestSetup.Fields("Approved") = Not rsTestSetup.Fields("Approved")
883        rsTestSetup.Update
884        rsTestSetup.Requery
885        lblTestDateApproved.Visible = rsTestSetup.Fields("Approved")
886        If rsTestSetup.Fields("Approved") = True Then
887            cmdApproveTestDate.Caption = "Unapprove This Test Date"
888        Else
889            cmdApproveTestDate.Caption = "Approve This Test Date"
890        End If
' <VB WATCH>
891        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
892        Exit Sub
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
893        On Error GoTo vbwErrHandler
894        Const VBWPROCNAME = "frmPLCData.cmdCalibrate_Click"
895        If vbwProtector.vbwTraceProc Then
896            Dim vbwProtectorParameterString As String
897            If vbwProtector.vbwTraceParameters Then
898                vbwProtectorParameterString = "()"
899            End If
900            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
901        End If
' </VB WATCH>
902        Dim ans As Integer
903        Dim I As Integer

904        ans = MsgBox("You have selected to calibrate the software.  Do you want to continue?", vbYesNo, "Calibrate Software")
905        If ans = vbNo Then
906            Calibrating = False
' <VB WATCH>
907        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
908            Exit Sub
909        Else
910            CalibrateSoftware
911        End If
' <VB WATCH>
912        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
913        Exit Sub
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
914        On Error GoTo vbwErrHandler
915        Const VBWPROCNAME = "frmPLCData.cmdClearPumpData_Click"
916        If vbwProtector.vbwTraceProc Then
917            Dim vbwProtectorParameterString As String
918            If vbwProtector.vbwTraceParameters Then
919                vbwProtectorParameterString = "()"
920            End If
921            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
922        End If
' </VB WATCH>
923        BlankData
' <VB WATCH>
924        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
925        Exit Sub
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
926        On Error GoTo vbwErrHandler
927        Const VBWPROCNAME = "frmPLCData.cmdDeletePump_Click"
928        If vbwProtector.vbwTraceProc Then
929            Dim vbwProtectorParameterString As String
930            If vbwProtector.vbwTraceParameters Then
931                vbwProtectorParameterString = "()"
932            End If
933            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
934        End If
' </VB WATCH>
935        Dim Answer As Integer
936        Answer = MsgBox("You are about to delete the following record: S/N = " & rsPumpData.Fields("SerialNumber") & "!  Do you want to continue?", vbCritical Or vbYesNo, "Ready to Delete")
937        If Answer = vbYes Then
938            rsPumpData.Delete
939            rsPumpData.Update
940            cmdFindPump_Click
941        End If
' <VB WATCH>
942        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
943        Exit Sub
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
944        On Error GoTo vbwErrHandler
945        Const VBWPROCNAME = "frmPLCData.cmdDeleteTestDate_Click"
946        If vbwProtector.vbwTraceProc Then
947            Dim vbwProtectorParameterString As String
948            If vbwProtector.vbwTraceParameters Then
949                vbwProtectorParameterString = "()"
950            End If
951            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
952        End If
' </VB WATCH>
953        Dim Answer As Integer
954        Answer = MsgBox("You are about to delete the following record: S/N = " & rsTestData.Fields("SerialNumber") & " and Test Date = " & rsTestSetup.Fields("Date") & "!  Do you want to continue?", vbCritical Or vbYesNo, "Ready to Delete")
955        If Answer = vbYes Then
956            rsTestSetup.Delete
957            rsTestSetup.Update
958            cmdFindPump_Click
959        End If
' <VB WATCH>
960        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
961        Exit Sub
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
962        On Error GoTo vbwErrHandler
963        Const VBWPROCNAME = "frmPLCData.cmdEnterPumpData_Click"
964        If vbwProtector.vbwTraceProc Then
965            Dim vbwProtectorParameterString As String
966            If vbwProtector.vbwTraceParameters Then
967                vbwProtectorParameterString = "()"
968            End If
969            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
970        End If
' </VB WATCH>
971        Dim d As Integer
972        Dim sSearch As String
973        Dim ans As Integer
974        Dim boWriteDataWritten As Boolean


           'check for a serial number
975        If LenB(txtSN.Text) = 0 Then
976            MsgBox "You must have a Serial Number to enter data.  Data has not been saved."
' <VB WATCH>
977        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
978            Exit Sub
979        End If

           'check to make sure most entries are filled in
980        If LenB(txtModelNo.Text) = 0 And optMfr(0).value = True Then
981            MsgBox "You need to enter a MODEL NO before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
982        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
983            Exit Sub
984        End If
985        If LenB(txtSalesOrderNumber.Text) = 0 Then
986            If InStr(1, txtSN.Text, "-") <> 0 Then
987                txtSalesOrderNumber.Text = Mid$(txtSN.Text, 1, InStr(1, txtSN.Text, "-") - 1)
988            End If
989        End If
990        If LenB(txtSalesOrderNumber.Text) = 0 Then
991            MsgBox "You need to enter a SALES ORDER NUMBER before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
992        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
993            Exit Sub
994        End If

995        If cmbMotor.ListIndex = -1 And optMfr(0).value = True Then
996            MsgBox "You need to pick a MOTOR before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
997        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
998            Exit Sub
999        End If

1000       If cmbStatorFill.ListIndex = -1 And optMfr(0).value = True Then    'set default
1001           cmbStatorFill.ListIndex = 0
1002       End If

1003       If cmbModel.ListIndex = -1 And optMfr(0).value = True Then
1004           MsgBox "You need to pick a MODEL before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1005       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1006           Exit Sub
1007       End If

1008       If cmbModelGroup.ListIndex = -1 And optMfr(0).value = True Then
1009           MsgBox "You need to pick a MODEL GROUP before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1010       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1011           Exit Sub
1012       End If


1013       If cmbDesignPressure.ListIndex = -1 And optMfr(0).value = True Then
1014           MsgBox "You need to pick a DESIGN PRESSURE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1015       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1016           Exit Sub
1017       End If

1018       If cmbCirculationPath.ListIndex = -1 And optMfr(0).value = True Then
1019           MsgBox "You need to pick a CIRCULATION PATH before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1020       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1021           Exit Sub
1022       End If

1023       If cmbRPM.ListIndex = -1 And optMfr(0).value = True Then
1024           MsgBox "You need to pick an RPM before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1025       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1026           Exit Sub
1027       End If

       'check TEMC dropdowns

1028       If cmbTEMCAdapter.ListIndex = -1 And optMfr(0).value = False Then
1029           MsgBox "You need to pick a TEMC ADAPTER before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1030       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1031           Exit Sub
1032       End If

1033       If cmbTEMCAdditions.ListIndex = -1 And optMfr(0).value = False Then
1034           MsgBox "You need to pick TEMC ADDITIONS before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1035       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1036           Exit Sub
1037       End If

1038       If cmbTEMCCirculation.ListIndex = -1 And optMfr(0).value = False Then
1039           MsgBox "You need to pick a TEMC CIRCULATION before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1040       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1041           Exit Sub
1042       End If

1043       If cmbTEMCDesignPressure.ListIndex = -1 And optMfr(0).value = False Then
1044           MsgBox "You need to pick a TEMC DESIGN PRESSURE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1045       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1046           Exit Sub
1047       End If

1048       If cmbTEMCDivisionType.ListIndex = -1 And optMfr(0).value = False Then
1049           MsgBox "You need to pick a TEMC DIVISION TYPE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1050       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1051           Exit Sub
1052       End If

1053       If cmbTEMCImpellerType.ListIndex = -1 And optMfr(0).value = False Then
1054           MsgBox "You need to pick a TEMC IMPELLER TYPE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1055       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1056           Exit Sub
1057       End If

1058       If cmbTEMCInsulation.ListIndex = -1 And optMfr(0).value = False Then
1059           MsgBox "You need to pick a TEMC INSULATION TYPE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1060       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1061           Exit Sub
1062       End If

1063       If cmbTEMCJacketGasket.ListIndex = -1 And optMfr(0).value = False Then
1064           MsgBox "You need to pick a TEMC JACKET GASKET before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1065       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1066           Exit Sub
1067       End If

1068       If cmbTEMCMaterials.ListIndex = -1 And optMfr(0).value = False Then
1069           MsgBox "You need to pick TEMC MATERIALS before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1070       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1071           Exit Sub
1072       End If

1073       If cmbTEMCModel.ListIndex = -1 And optMfr(0).value = False Then
1074           MsgBox "You need to pick a TEMC MODEL before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1075       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1076           Exit Sub
1077       End If

1078       If cmbTEMCNominalImpSize.ListIndex = -1 And optMfr(0).value = False Then
1079           MsgBox "You need to pick a TEMC NOMINAL IMPELLER SIZE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1080       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1081           Exit Sub
1082       End If

1083       If cmbTEMCNominalDischargeSize.ListIndex = -1 And optMfr(0).value = False Then
1084           MsgBox "You need to pick a TEMC NOMINAL DISCHARGE SIZE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1085       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1086           Exit Sub
1087       End If

1088       If cmbTEMCNominalSuctionSize.ListIndex = -1 And optMfr(0).value = False Then
1089           MsgBox "You need to pick a TEMC NOMINAL SUCTION SIZE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1090       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1091           Exit Sub
1092       End If

1093       If cmbTEMCOtherMotor.ListIndex = -1 And optMfr(0).value = False Then
1094           MsgBox "You need to pick a TEMC OTHER MOTOR before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1095       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1096           Exit Sub
1097       End If

1098       If cmbTEMCPumpStages.ListIndex = -1 And optMfr(0).value = False Then
1099           MsgBox "You need to pick TEMC NUMBER OF PUMP STAGES before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1100       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1101           Exit Sub
1102       End If

1103       If cmbTEMCTRG.ListIndex = -1 And optMfr(0).value = False Then
1104           MsgBox "You need to pick a TEMC TRG before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1105       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1106           Exit Sub
1107       End If

1108       If cmbTEMCVoltage.ListIndex = -1 And optMfr(0).value = False Then
1109           MsgBox "You need to pick a TEMC VOLTAGE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1110       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1111           Exit Sub
1112       End If

1113       If LenB(txtTEMCFrameNumber.Text) = 0 And optMfr(0).value = False Then
1114           MsgBox "You need to enter a TEMC FRAME NUMBER before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1115       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1116           Exit Sub
1117       End If


1118       If Not boFoundPump Then     'if we havent found a pump in the database, add it
1119           rsPumpData.AddNew
1120           boWriteDataWritten = False
1121       Else    'else, find the entry
1122           sSearch = "Serialnumber = '" & frmPLCData.txtSN.Text & "'"
1123           rsPumpData.MoveFirst
1124           rsPumpData.Find sSearch, , adSearchForward, 1
1125           boWriteDataWritten = True
1126       End If

1127       If Not IsNull(rsPumpData!DataWritten) Or rsPumpData!DataWritten = True Then
1128           ans = MsgBox("You have already entered data for this pump.  Do you want to overwrite the data?", vbDefaultButton2 + vbYesNo, "Overwrite Data?")
1129           If ans = vbNo Then
1130               rsPumpData!DataWritten = True
1131               rsPumpData.Update   'update datawritten
' <VB WATCH>
1132       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1133               Exit Sub
1134           End If
1135       End If

1136       rsPumpData!SerialNumber = frmPLCData.txtSN.Text
1137       rsPumpData!ModelNumber = frmPLCData.txtModelNo.Text
1138       rsPumpData!SalesOrderNumber = frmPLCData.txtSalesOrderNumber.Text
1139       rsPumpData!ShipToCustomer = frmPLCData.txtShpNo.Text
1140       rsPumpData!BillToCustomer = frmPLCData.txtBilNo.Text
1141       rsPumpData!ApplicationFluid = frmPLCData.txtLiquid
1142       rsPumpData!NPSHFile = frmPLCData.txtNPSHFileLocation.Text
1143       rsPumpData!RVSPartNo = frmPLCData.txtRVSPartNo.Text
1144       rsPumpData!CustPN = frmPLCData.txtXPartNum.Text
1145       rsPumpData!CustPO = frmPLCData.txtCustPONum.Text

1146       If Len(frmPLCData.txtViscosity) <> 0 Then
1147           rsPumpData!ApplicationViscosity = frmPLCData.txtViscosity
1148       End If

1149       If frmPLCData.chkSuperMarketFeathered.value = Checked Then
1150           rsPumpData!Field1 = "Feathered"
1151       Else
1152           rsPumpData!Field1 = ""
1153       End If

1154       If LenB(txtSpGr.Text) <> 0 Then
1155           If Not IsNumeric(frmPLCData.txtSpGr.Text) Then
1156               MsgBox "Specific Gravity must be a number."
' <VB WATCH>
1157       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1158               Exit Sub
1159           End If
1160           rsPumpData!SpGr = frmPLCData.txtSpGr.Text
1161       End If
1162       If LenB(txtImpellerDia.Text) <> 0 Then
1163           If Not IsNumeric(frmPLCData.txtImpellerDia.Text) Then
1164               MsgBox "Impeller Diameter must be a number."
' <VB WATCH>
1165       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1166               Exit Sub
1167           End If
1168           rsPumpData!impellerdia = frmPLCData.txtImpellerDia.Text
1169       End If
1170       If LenB(txtDesignFlow.Text) <> 0 Then
1171           rsPumpData!designflow = frmPLCData.txtDesignFlow.Text
1172       End If
1173       If LenB(txtDesignTDH.Text) <> 0 Then
1174           rsPumpData!designtdh = frmPLCData.txtDesignTDH.Text
1175       End If
1176       If LenB(txtRemarks.Text) <> 0 Then
1177           rsPumpData!Remarks = txtRemarks.Text
1178       End If

1179       If optMfr(0).value = True Then
1180           d = cmbMotor.ItemData(cmbMotor.ListIndex)
1181           rsPumpData!Motor = d
1182           d = cmbStatorFill.ItemData(cmbStatorFill.ListIndex)
1183           rsPumpData!StatorFill = d
1184            d = cmbDesignPressure.ItemData(cmbDesignPressure.ListIndex)
1185           rsPumpData!DesignPressure = d
1186           d = cmbCirculationPath.ItemData(cmbCirculationPath.ListIndex)
1187           rsPumpData!CirculationPath = d
1188           d = cmbRPM.ItemData(cmbRPM.ListIndex)
1189           rsPumpData!RPM = d
1190           d = cmbModel.ItemData(cmbModel.ListIndex)
1191           rsPumpData!Model = d
1192           d = cmbModelGroup.ItemData(cmbModelGroup.ListIndex)
1193           rsPumpData!ModelGroup = d
1194       End If
       '   TEMC fields
1195       If optMfr(0).value = False Then
1196           d = cmbTEMCAdapter.ItemData(cmbTEMCAdapter.ListIndex)
1197           rsPumpData!TEMCAdapter = d

1198           d = cmbTEMCAdditions.ItemData(cmbTEMCAdditions.ListIndex)
1199           rsPumpData!TEMCAdditions = d

1200           d = cmbTEMCCirculation.ItemData(cmbTEMCCirculation.ListIndex)
1201           rsPumpData!TEMCcirculation = d

1202           d = cmbTEMCDesignPressure.ItemData(cmbTEMCDesignPressure.ListIndex)
1203           rsPumpData!TEMCDesignpressure = d

1204           d = cmbTEMCDivisionType.ItemData(cmbTEMCDivisionType.ListIndex)
1205           rsPumpData!TEMCDivisionType = d

1206           d = cmbTEMCImpellerType.ItemData(cmbTEMCImpellerType.ListIndex)
1207           rsPumpData!TEMCImpellerType = d

1208           d = cmbTEMCInsulation.ItemData(cmbTEMCInsulation.ListIndex)
1209           rsPumpData!TEMCInsulation = d

1210           d = cmbTEMCJacketGasket.ItemData(cmbTEMCJacketGasket.ListIndex)
1211           rsPumpData!TEMCJacketGasket = d

1212           d = cmbTEMCMaterials.ItemData(cmbTEMCMaterials.ListIndex)
1213           rsPumpData!TEMCMaterials = d

1214           d = cmbTEMCModel.ItemData(cmbTEMCModel.ListIndex)
1215           rsPumpData!TEMCModel = d

1216           d = cmbTEMCNominalImpSize.ItemData(cmbTEMCNominalImpSize.ListIndex)
1217           rsPumpData!TEMCNominalImpSize = d

1218           d = cmbTEMCNominalDischargeSize.ItemData(cmbTEMCNominalDischargeSize.ListIndex)
1219           rsPumpData!TEMCNominalDischargeSize = d

1220           d = cmbTEMCNominalSuctionSize.ItemData(cmbTEMCNominalSuctionSize.ListIndex)
1221           rsPumpData!TEMCNominalSuctionSize = d

1222           d = cmbTEMCOtherMotor.ItemData(cmbTEMCOtherMotor.ListIndex)
1223           rsPumpData!TEMCOtherMotor = d

1224           d = cmbTEMCPumpStages.ItemData(cmbTEMCPumpStages.ListIndex)
1225           rsPumpData!TEMCPumpStages = d

1226           d = cmbTEMCTRG.ItemData(cmbTEMCTRG.ListIndex)
1227           rsPumpData!TEMCTRG = d

1228           d = cmbTEMCVoltage.ItemData(cmbTEMCVoltage.ListIndex)
1229           rsPumpData!TEMCVoltage = d

1230           If LenB(txtTEMCFrameNumber.Text) <> 0 Then
1231               rsPumpData!TEMCFrameNumber = frmPLCData.txtTEMCFrameNumber.Text
1232           End If
1233       End If

1234       rsPumpData!ChempumpPump = optMfr(0).value

1235       rsPumpData!Approved = False

       'added from TEMC Inspection Report
1236       If Len(txtJobNum.Text) <> 0 Then
1237           rsPumpData!JobNumber = txtJobNum.Text
1238       End If

1239       If Len(txtNoPhases.Text) <> 0 Then
1240           rsPumpData!Phases = txtNoPhases.Text
1241       End If

1242       If Len(txtExpClass.Text) <> 0 Then
1243           rsPumpData!ExpClass = txtExpClass.Text
1244       End If

1245       If Len(txtThermalClass.Text) <> 0 Then
1246           rsPumpData!ThermalClass = txtThermalClass.Text
1247       End If

1248       rsPumpData!NPSHr = Val(txtNPSHr.Text)
1249       rsPumpData!LiquidTemperature = Val(txtLiquidTemperature.Text)
1250       rsPumpData!RatedInputPower = Val(txtRatedInputPower.Text)
1251       rsPumpData!FLCurrent = Val(txtAmps.Text)





1252       If boWriteDataWritten Then
1253           rsPumpData!DataWritten = True
1254       Else
1255           rsPumpData!DataWritten = False
1256       End If

           'write the data into the database
1257       rsPumpData.Update
1258       boFoundPump = True

           'enter a new test date if it's a new entry
1259       If Not boWriteDataWritten Then


1260           cmdAddNewTestDate_Click
1261       End If
' <VB WATCH>
1262       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1263       Exit Sub
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
1264       On Error GoTo vbwErrHandler
1265       Const VBWPROCNAME = "frmPLCData.cmdEnterTestData_Click"
1266       If vbwProtector.vbwTraceProc Then
1267           Dim vbwProtectorParameterString As String
1268           If vbwProtector.vbwTraceParameters Then
1269               vbwProtectorParameterString = "()"
1270           End If
1271           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1272       End If
' </VB WATCH>
1273       Dim sSearch As String
1274       Dim ans As Integer

           'if we didn't find the test setup, can't enter test data
1275       If Not boFoundTestSetup Then
1276           MsgBox "You must enter Test Setup Data before entering the Test Data"
' <VB WATCH>
1277       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1278           Exit Sub
1279       End If

           'if we don't find data in the test database, add records
1280       If boFoundTestData = False Then     'add 8 records for 8 tests
1281           AddTestData
1282           rsTestData.MoveFirst
1283       Else        'find the data in the database
1284           sSearch = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"
1285           rsTestData.MoveFirst
1286           rsTestData.Filter = sSearch
1287       End If

           'find the desired record from the form
1288       rsTestData.MoveFirst
1289       rsTestData.Move UpDown1.value - 1

1290       If rsTestData!DataWritten = True Then
1291           ans = MsgBox("You have already entered data for this test.  Do you want to overwrite the data?", vbYesNo + vbDefaultButton2, "Data Already Entered")
1292           If ans = vbNo Then
' <VB WATCH>
1293       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1294               Exit Sub
1295           End If
1296       End If

1297       rsEff.MoveFirst
1298       rsEff.Move UpDown1.value - 1

1299       If LenB(txtV1.Text) <> 0 Then
1300           rsTestData!VoltageA = Val(txtV1.Text)
1301       End If

1302       If LenB(txtV2.Text) <> 0 Then
1303           rsTestData!VoltageB = Val(txtV2.Text)
1304       End If

1305       If LenB(txtV3.Text) <> 0 Then
1306           rsTestData!VoltageC = Val(txtV3.Text)
1307       End If

1308       If LenB(txtI1.Text) <> 0 Then
1309           rsTestData!CurrentA = Val(txtI1.Text)
1310       End If

1311       If LenB(txtI2.Text) <> 0 Then
1312           rsTestData!CurrentB = Val(txtI2.Text)
1313       End If

1314       If LenB(txtI3.Text) <> 0 Then
1315           rsTestData!CurrentC = Val(txtI3.Text)
1316       End If

1317       If LenB(txtP1.Text) <> 0 Then
1318           rsTestData!PowerA = Val(txtP1.Text)
1319       End If

1320       If LenB(txtP2.Text) <> 0 Then
1321           rsTestData!PowerB = Val(txtP2.Text)
1322       End If

1323       If LenB(txtP3.Text) <> 0 Then
1324           rsTestData!PowerC = Val(txtP3.Text)
1325       End If

1326       If LenB(txtKW.Text) <> 0 Then
1327           rsTestData!TotalPower = Val(txtKW.Text)
1328       End If

1329       rsTestData!Flow = Val(txtFlowDisplay.Text)
1330       rsTestData!DischargePressure = Val(txtDischargeDisplay.Text)
1331       rsTestData!SuctionPressure = Val(txtSuctionDisplay.Text)
1332       rsTestData!TemperatureSuction = Val(txtTemperatureDisplay.Text)

1333       rsTestData!TC1 = Val(txtTC1Display.Text)
1334       rsTestData!TC2 = Val(txtTC2Display.Text)
1335       rsTestData!TC3 = Val(txtTC3Display.Text)
1336       rsTestData!TC4 = Val(txtTC4Display.Text)

1337       rsTestData!CircFlow = Val(txtAI1Display.Text)
1338       rsTestData!RBHTemp = Val(txtAI2Display.Text)
1339       rsTestData!RBHPress = Val(txtAI3Display.Text)
1340       rsTestData!AI4 = Val(txtAI4Display.Text)

1341       rsTestData!ValvePosition = Val(txtValvePosition.Text)
1342       rsTestData!SetPoint = Val(txtSetPoint.Text)

1343       If LenB(txtThrustBal.Text) <> 0 Then
1344           rsTestData!ThrustBalance = txtThrustBal.Text
1345       End If

1346       If LenB(txtVibAx.Text) <> 0 Then
1347           rsTestData!VibrationX = txtVibAx.Text
1348       End If

1349       If LenB(txtVibRad.Text) <> 0 Then
1350           rsTestData!VibrationY = txtVibRad.Text
1351       End If

1352       If LenB(txtTEMCTRGReading.Text) <> 0 Then
1353           rsTestData!TEMCTRG = txtTEMCTRGReading.Text
1354       Else
1355           rsTestData!TEMCTRG = 0
1356       End If

1357       If LenB(txtRPM.Text) <> 0 Then
1358           rsTestData!RPM = txtRPM.Text
1359       End If

1360       If LenB(txtTestRemarks.Text) <> 0 Then
1361           rsTestData!Remarks = txtTestRemarks.Text
1362       Else
1363           rsTestData!Remarks = " "
1364       End If

1365       If LenB(txtTEMCTRGReading.Text) <> 0 Then
1366           rsTestData!TEMCTRG = txtTEMCTRGReading.Text
1367       End If

1368       If LenB(txtTEMCFrontThrust.Text) <> 0 Then
1369           rsTestData!TEMCFrontThrust = txtTEMCFrontThrust.Text
1370       End If

1371       If LenB(txtTEMCRearThrust.Text) <> 0 Then
1372           rsTestData!TEMCRearThrust = txtTEMCRearThrust.Text
1373       End If

1374       If LenB(txtTEMCMomentArm.Text) <> 0 Then
1375           rsTestData!TEMCMomentArm = txtTEMCMomentArm.Text
1376       End If

1377       If LenB(txtTEMCThrustRigPressure.Text) <> 0 Then
1378           rsTestData!TEMCThrustRigPressure = txtTEMCThrustRigPressure.Text
1379       End If

1380       If LenB(txtTEMCViscosity.Text) <> 0 Then
1381           rsTestData!TEMCViscosity = txtTEMCViscosity.Text
1382       End If

1383       If LenB(txtNPSHa.Text) <> 0 Then
1384           rsTestData!NPSHa = txtNPSHa.Text
1385       End If

1386       rsTestData!Approved = False

1387       rsTestData!DataWritten = True

           'update the database
1388       rsTestData.Update

1389       DoEfficiencyCalcs
1390       rsEff.Update

           'update the form
1391       DataGrid1.Refresh
1392       DataGrid2.Refresh

1393       FixPointsToPlot

' <VB WATCH>
1394       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1395       Exit Sub
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
1396       On Error GoTo vbwErrHandler
1397       Const VBWPROCNAME = "frmPLCData.cmdEnterTestSetupData_Click"
1398       If vbwProtector.vbwTraceProc Then
1399           Dim vbwProtectorParameterString As String
1400           If vbwProtector.vbwTraceParameters Then
1401               vbwProtectorParameterString = "()"
1402           End If
1403           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1404       End If
' </VB WATCH>
1405       Dim I As Integer
1406       Dim d As Integer
1407       Dim sSearch As String
1408       Dim ans As Integer
1409       Dim boWriteDataWritten As Boolean

           'check for a serial number
1410       If LenB(txtSN.Text) = 0 Then
1411           MsgBox "You must have a Serial Number to enter data.", vbOKOnly + vbExclamation, "Cannot Enter Data"
' <VB WATCH>
1412       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1413           Exit Sub
1414       End If

1415       If Pressed = True Then
1416           If Me.cmbDischDia.ListIndex = -1 Or Me.cmbSuctDia.ListIndex = -1 Or Val(Me.txtSuctHeight.Text) = 0 Or Val(Me.txtDischHeight.Text) = 0 Then
1417               MsgBox "You must have Discharge Diameter AND Suction Diameter AND Suction Height AND Discharge Height entered to enter data.", vbOKOnly + vbExclamation, "Cannot Enter Data"
' <VB WATCH>
1418       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1419               Exit Sub
1420           End If
1421       End If

1422       Pressed = True
1423       If Not boFoundTestSetup Then    'if we didn't find any test setup, add a record
1424           rsTestSetup.AddNew
1425           cmbTestDate.AddItem Now
1426           cmbTestDate.ListIndex = cmbTestDate.NewIndex
1427           cmdAddNewBalanceHoles.Visible = True
1428           boFoundTestSetup = True
1429           boWriteDataWritten = False
1430           rsTestSetup!DataWritten = False
1431       Else    'find the record and display
1432           sSearch = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"
1433           rsTestSetup.MoveFirst
1434           rsTestSetup.Filter = sSearch
1435           If Not boCanApprove Then
       '            cmdAddNewBalanceHoles.Visible = False
1436           End If
1437           boWriteDataWritten = True
1438       End If

1439       If rsTestSetup!DataWritten = True Then
1440           ans = MsgBox("Data has already been entered for this test date.  Do you want to overwrite it?", vbYesNo + vbDefaultButton2, "Data Exists")
1441           If ans = vbNo Then
' <VB WATCH>
1442       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1443               Exit Sub
1444           End If
1445       End If

1446       rsTestSetup!SerialNumber = txtSN
1447       rsTestSetup!Date = cmbTestDate.List(cmbTestDate.ListIndex)

1448       I = cmbFlowMeter.ListIndex
1449       If I = -1 Then
1450           d = 1
1451           rsTestSetup!FlowMeterID = d
1452       Else
1453           d = cmbLoopNumber.ItemData(I)
1454           rsTestSetup!FlowMeterID = d
1455       End If

1456       I = cmbSuctionPressureTransducer.ListIndex
1457       If I = -1 Then
1458           d = 1
1459           rsTestSetup!suctionid = d
1460       Else
1461           d = cmbLoopNumber.ItemData(I)
1462           rsTestSetup!suctionid = d
1463       End If

1464       I = cmbDischargePressureTransducer.ListIndex
1465       If I = -1 Then
1466           d = 1
1467           rsTestSetup!dischid = d
1468       Else
1469           d = cmbLoopNumber.ItemData(I)
1470           rsTestSetup!dischid = d
1471       End If

1472       I = cmbTemperatureTransducer.ListIndex
1473       If I = -1 Then
1474           d = 1
1475           rsTestSetup!temperatureid = d
1476       Else
1477           d = cmbLoopNumber.ItemData(I)
1478           rsTestSetup!temperatureid = d
1479       End If

1480       I = Me.cmbCirculationFlowMeter.ListIndex
1481       If I = -1 Or I < 4 Then
1482           d = 5
1483           rsTestSetup!magflowid = d
1484       Else
1485           d = cmbLoopNumber.ItemData(I)
1486           rsTestSetup!magflowid = d
1487       End If


1488       If LenB(txtHDCor.Text) <> 0 Then
1489           rsTestSetup!HDCor = txtHDCor
1490       Else
1491           rsTestSetup!HDCor = 0
1492       End If
1493       If LenB(txtKWMult.Text) <> 0 Then
1494           rsTestSetup!kwmult = txtKWMult
1495       Else
1496           rsTestSetup!kwmult = 1
1497       End If
1498       If LenB(txtWho.Text) <> 0 Then
1499           rsTestSetup!who = txtWho
1500       Else
1501           rsTestSetup!who = vbNullString
1502       End If
1503       If LenB(txtRMA.Text) <> 0 Then
1504           rsTestSetup!RMA = txtRMA
1505       Else
1506           rsTestSetup!RMA = vbNullString
1507       End If
1508       If LenB(frmPLCData.txtDischHeight) <> 0 Then
1509           rsTestSetup!DischargeGageHeight = Val(txtDischHeight)
1510       Else
1511           rsTestSetup!DischargeGageHeight = 0
1512       End If
1513       If LenB(frmPLCData.txtSuctHeight) <> 0 Then
1514           rsTestSetup!SuctionGageHeight = Val(txtSuctHeight)
1515       Else
1516           rsTestSetup!SuctionGageHeight = 0
1517       End If
1518       If LenB(frmPLCData.txtTestSetupRemarks.Text) <> 0 Then
1519           rsTestSetup!Remarks = txtTestSetupRemarks.Text
1520       Else
1521           rsTestSetup!Remarks = vbNullString
1522       End If
1523       If LenB(frmPLCData.txtVFDFreq.Text) <> 0 Then
1524           rsTestSetup!VFDFrequency = txtVFDFreq.Text
1525       Else
1526           rsTestSetup!VFDFrequency = 0
1527       End If

1528       I = cmbOrificeNumber.ListIndex
1529       If I = -1 Then
1530           d = 18      'entry for None
1531       Else
1532           d = cmbOrificeNumber.ItemData(I)
1533       End If
1534       rsTestSetup!orificenumber = d

1535       If LenB(txtEndPlay.Text) <> 0 Then
1536           rsTestSetup!Endplay = Val(frmPLCData.txtEndPlay.Text)
1537       Else
1538           rsTestSetup!Endplay = 0
1539       End If

1540       If LenB(txtGGap.Text) <> 0 Then
1541           rsTestSetup!GGAP = Val(frmPLCData.txtGGap.Text)
1542       Else
1543           rsTestSetup!GGAP = 0
1544       End If

1545       If LenB(txtOtherMods.Text) <> 0 Then
1546           rsTestSetup!OtherMods = txtOtherMods.Text
1547       Else
1548           rsTestSetup!OtherMods = vbNullString
1549       End If

1550       rsTestSetup!Approved = False

1551       I = cmbLoopNumber.ListIndex
1552       If I = -1 Then
1553           d = 1
1554           rsTestSetup!loopnumber = d
1555       Else
1556           d = cmbLoopNumber.ItemData(I)
1557           rsTestSetup!loopnumber = d
1558       End If

1559       I = cmbSuctDia.ListIndex
1560       If I = -1 Then
1561           d = -1
1562       Else
1563           d = cmbSuctDia.ItemData(I)
1564           rsTestSetup!SuctDiam = d
1565       End If

1566       I = cmbDischDia.ListIndex
1567       If I = -1 Then
1568           d = -1
1569       Else
1570           d = cmbDischDia.ItemData(I)
1571           rsTestSetup!DischDiam = d
1572       End If

1573       I = cmbTachID.ListIndex
1574       If I = -1 Then
1575           d = 1
1576           rsTestSetup!tachid = d
1577       Else
1578           d = cmbTachID.ItemData(I)
1579           rsTestSetup!tachid = d
1580       End If

1581       I = cmbAnalyzerNo.ListIndex
1582       If I = -1 Then
1583           d = 1
1584       Else
1585           d = cmbAnalyzerNo.ItemData(I)
1586       End If
1587       rsTestSetup!analyzerno = d

1588       I = cmbTestSpec.ListIndex
1589       If I = -1 Then
1590           d = 1
1591       Else
1592           d = cmbTestSpec.ItemData(I)
1593       End If
1594       rsTestSetup!testspec = d

1595       I = cmbVoltage.ListIndex
1596       If I = -1 Then
1597           d = 1
1598       Else
1599           d = cmbVoltage.ItemData(I)
1600       End If
1601       rsTestSetup!Voltage = d

1602       I = cmbFrequency.ListIndex
1603       If I = -1 Then
1604           d = 1
1605       Else
1606           d = cmbFrequency.ItemData(I)
1607       End If
1608       rsTestSetup!Frequency = d

1609       I = cmbMounting.ListIndex
1610       If I = -1 Then
1611           d = 1
1612       Else
1613           d = cmbMounting.ItemData(I)
1614       End If
1615       rsTestSetup!Mounting = d

1616       I = cmbPLCNo.ListIndex
1617       If I = -1 Then
1618           d = 8
1619       Else
1620           d = cmbPLCNo.ItemData(I)
1621       End If
1622       rsTestSetup!PLCNo = d

1623       rsTestSetup!ImpFeathered = chkFeathered.value

1624       If chkTrimmed.value = 1 Then
1625           rsTestSetup!ImpTrimmed = Val(txtImpTrim)
1626       Else
1627           rsTestSetup!ImpTrimmed = 0
1628       End If
1629       chkTrimmed_Click

1630       If chkOrifice.value = 1 Then
1631           rsTestSetup!PumpDischOrifice = Val(txtOrifice)
1632       Else
1633           rsTestSetup!PumpDischOrifice = 0
1634       End If
1635       chkOrifice_Click

1636       If chkCircOrifice.value = 1 Then
1637           rsTestSetup!CircFlowOrifice = Val(txtCircOrifice)
1638       Else
1639           rsTestSetup!CircFlowOrifice = 0
1640       End If
1641       chkCircOrifice_Click

1642       chkBalanceHoles_Click

1643       If chkNPSH.value = 1 Then
1644           txtNPSHFile.Visible = True
1645           rsTestSetup!NPSHFile = txtNPSHFile
1646       Else
1647           rsTestSetup!NPSHFile = vbNullString
1648           txtNPSHFile.Visible = False
1649       End If

1650       If chkPictures.value = 1 Then
1651           txtPicturesFile.Visible = True
1652           rsTestSetup!PictureFile = txtPicturesFile
1653       Else
1654           rsTestSetup!PictureFile = vbNullString
1655           txtPicturesFile.Visible = False
1656       End If

1657       If chkVibration.value = 1 Then
1658           txtVibrationFile.Visible = True
1659           rsTestSetup!VibrationFile = txtVibrationFile
1660       Else
1661           rsTestSetup!VibrationFile = vbNullString
1662           txtVibrationFile.Visible = False
1663       End If

1664       If boWriteDataWritten Then
1665           rsTestSetup!DataWritten = True
1666       Else
1667           rsTestSetup!DataWritten = False
1668       End If

           'for TEMC Inspection Report
1669       If LenB(frmPLCData.txtTestAndInspection(0).Text) <> 0 Then
1670           rsTestSetup!InsulationMeggerVolts = frmPLCData.txtTestAndInspection(0).Text
1671       Else
1672           rsTestSetup!InsulationMeggerVolts = ""
1673       End If

1674       If LenB(frmPLCData.txtTestAndInspection(1).Text) <> 0 Then
1675           rsTestSetup!InsulationMegOhms = frmPLCData.txtTestAndInspection(1).Text
1676       Else
1677           rsTestSetup!InsulationMegOhms = ""
1678       End If

1679       If LenB(frmPLCData.txtTestAndInspection(2).Text) <> 0 Then
1680           rsTestSetup!DielectricVolts = frmPLCData.txtTestAndInspection(2).Text
1681       Else
1682           rsTestSetup!DielectricVolts = ""
1683       End If

1684       If LenB(frmPLCData.txtTestAndInspection(3).Text) <> 0 Then
1685           rsTestSetup!DielectricTime = frmPLCData.txtTestAndInspection(3).Text
1686       Else
1687           rsTestSetup!DielectricTime = ""
1688       End If

1689       If LenB(frmPLCData.txtTestAndInspection(4).Text) <> 0 Then
1690           rsTestSetup!HydrostaticValue = frmPLCData.txtTestAndInspection(4).Text
1691       Else
1692           rsTestSetup!HydrostaticValue = ""
1693       End If

1694       If LenB(frmPLCData.txtTestAndInspection(5).Text) <> 0 Then
1695           rsTestSetup!HydrostaticTime = frmPLCData.txtTestAndInspection(5).Text
1696       Else
1697           rsTestSetup!HydrostaticTime = ""
1698       End If

1699       If LenB(frmPLCData.txtTestAndInspection(6).Text) <> 0 Then
1700           rsTestSetup!PneumaticValue = frmPLCData.txtTestAndInspection(6).Text
1701       Else
1702           rsTestSetup!PneumaticValue = ""
1703       End If

1704       If LenB(frmPLCData.txtTestAndInspection(7).Text) <> 0 Then
1705           rsTestSetup!PneumaticTime = frmPLCData.txtTestAndInspection(7).Text
1706       Else
1707           rsTestSetup!PneumaticTime = ""
1708       End If

1709       I = cmbTestAndInspection(0).ListIndex
1710       If I = -1 Then
1711           rsTestSetup!HydrostaticUnits = ""
1712       Else
1713           rsTestSetup!HydrostaticUnits = cmbTestAndInspection(0).Text
1714       End If


1715       I = cmbTestAndInspection(1).ListIndex
1716       If I = -1 Then
1717           rsTestSetup!PneumaticUnits = ""
1718       Else
1719           rsTestSetup!PneumaticUnits = cmbTestAndInspection(1).Text
1720       End If

           'use abs to convert from 1 and 0 to boolean
1721       rsTestSetup!insulationgood = Abs(TestAndInspectionGood(0).value)
1722       rsTestSetup!DielectricGood = Abs(TestAndInspectionGood(1).value)
1723       rsTestSetup!HydrostaticGood = Abs(TestAndInspectionGood(2).value)
1724       rsTestSetup!PneumaticGood = Abs(TestAndInspectionGood(3).value)
1725       rsTestSetup!GeneralAppearanceGood = Abs(TestAndInspectionGood(4).value)
1726       rsTestSetup!OutlineDimensionsGood = Abs(TestAndInspectionGood(5).value)
1727       rsTestSetup!MotorNoLoadTestGood = Abs(TestAndInspectionGood(6).value)
1728       rsTestSetup!MotorLockedRotorTestGood = Abs(TestAndInspectionGood(7).value)
1729       rsTestSetup!HydrostaticTestGood = Abs(TestAndInspectionGood(8).value)
1730       rsTestSetup!HydraulicTestGood = Abs(TestAndInspectionGood(9).value)
1731       rsTestSetup!NPSHTestGood = Abs(TestAndInspectionGood(10).value)
1732       rsTestSetup!CleanPurgeSealGood = Abs(TestAndInspectionGood(11).value)
1733       rsTestSetup!PaintCheckGood = Abs(TestAndInspectionGood(12).value)
1734       rsTestSetup!NameplateGood = Abs(TestAndInspectionGood(13).value)
1735       rsTestSetup!SupervisorApproval = Abs(TestAndInspectionGood(14).value)

           'update the database
1736       rsTestSetup.Update

1737       If boFoundTestData = False Then     'add 8 records for 8 tests
1738           AddTestData
1739       End If

1740       rsTestSetup.Filter = vbNullString
' <VB WATCH>
1741       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1742       Exit Sub
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
1743       On Error GoTo vbwErrHandler
1744       Const VBWPROCNAME = "frmPLCData.cmdExit_Click"
1745       If vbwProtector.vbwTraceProc Then
1746           Dim vbwProtectorParameterString As String
1747           If vbwProtector.vbwTraceParameters Then
1748               vbwProtectorParameterString = "()"
1749           End If
1750           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1751       End If
' </VB WATCH>
1752       End
' <VB WATCH>
1753       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1754       Exit Sub
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
1755       On Error GoTo vbwErrHandler
1756       Const VBWPROCNAME = "frmPLCData.cmdFindMagtrols_Click"
1757       If vbwProtector.vbwTraceProc Then
1758           Dim vbwProtectorParameterString As String
1759           If vbwProtector.vbwTraceParameters Then
1760               vbwProtectorParameterString = "()"
1761           End If
1762           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1763       End If
' </VB WATCH>
1764       FindMagtrols
' <VB WATCH>
1765       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1766       Exit Sub
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
1767       On Error GoTo vbwErrHandler
1768       Const VBWPROCNAME = "frmPLCData.cmdFindPump_Click"
1769       If vbwProtector.vbwTraceProc Then
1770           Dim vbwProtectorParameterString As String
1771           If vbwProtector.vbwTraceParameters Then
1772               vbwProtectorParameterString = "()"
1773           End If
1774           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1775       End If
' </VB WATCH>
1776       Dim sAns As String
1777       Dim sSO As String
1778       Dim sParam As String
1779       Dim sName As String

1780       Dim I As Integer

           'clear the data
1781       BlankData

           'set TC and AI labels with default values
1782       txtTitle(0).Text = "TC 1"
1783       txtTitle(1).Text = "(F)"
1784       txtTitle(2).Text = "TC 2"
1785       txtTitle(3).Text = "(F)"
1786       txtTitle(4).Text = "TC 3"
1787       txtTitle(5).Text = "(F)"
1788       txtTitle(6).Text = "TC 4"
1789       txtTitle(7).Text = "(F)"
1790       txtTitle(20).Text = "Circ Flow"
1791       txtTitle(21).Text = "(GPM)"
1792       txtTitle(22).Text = "P1"
1793       txtTitle(23).Text = "(psig)"
1794       txtTitle(24).Text = "P2"
1795       txtTitle(25).Text = "(psig)"
1796       txtTitle(26).Text = "AI 4"
1797       txtTitle(27).Text = ""


1798       For I = 0 To 7
1799           lblAutoMan(I).Caption = "Auto"
1800       Next I

1801       txtFlowDisplay.Enabled = False
1802       txtSuctionDisplay.Enabled = False
1803       txtDischargeDisplay.Enabled = False
1804       txtTemperatureDisplay.Enabled = False
1805       txtAI1Display.Enabled = False
1806       txtAI2Display.Enabled = False
1807       txtAI3Display.Enabled = False
1808       txtAI4Display.Enabled = False


1809       cmdFindPump.Default = False

           'set all found booleans to false
       '    boUsingHP = False
1810       boFoundPump = False
1811       boPumpIsApproved = False
1812       boFoundTestSetup = False
1813       boFoundTestData = False


           'get rid of all test dates in combo box
1814       For I = cmbTestDate.ListCount - 1 To 0 Step -1
1815           cmbTestDate.RemoveItem 0
1816       Next I

1817       rsTestData.Filter = "SerialNumber = ''"

1818       DataGrid2.ClearFields
1819       ClearEff

1820       If rsPumpData.State = adStateOpen Then
1821           If rsPumpData.BOF = False Or rsPumpData.EOF = False Then
1822               rsPumpData.Update
1823           End If
1824           rsPumpData.Close
1825       End If

           'parse the serial number to make sure it is formed correctly
1826       Dim ok As Boolean
1827       ok = UCase(txtSN.Text) Like "[A-Z][A-Z][0-9][0-9][0-9][0-9][A-Z]" Or UCase(txtSN.Text) Like "[A-Z][A-Z][0-9][0-9][0-9][0-9][A-Z]-[0-9]" Or UCase(txtSN.Text) Like "[A-Z][A-Z][0-9][0-9][0-9][0-9][A-Z]-[0-9][0-9]"
1828       If Not ok Then
1829           MsgBox "Serial Number must be 2 letters, 4 numbers, and 1 letter. Please re-enter.", vbOKOnly, "Serial Number not correctly formed."
' <VB WATCH>
1830       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1831           Exit Sub
1832       End If

           'find the pump listed in the Serial Number text box
1833       qyPumpData.ActiveConnection = cnPumpData
1834       qyPumpData.CommandText = "SELECT * From TempPumpData WHERE (((TempPumpData.SerialNumber)='" & _
                                    txtSN.Text & "'))"
1835       rsPumpData.CursorType = adOpenStatic
1836       rsPumpData.CursorLocation = adUseClient
1837       rsPumpData.Index = "SerialNumber"
1838       rsPumpData.Open qyPumpData
1839       boEpicorFound = False

1840       If rsPumpData.BOF = True And rsPumpData.EOF = True Then
               'if the bof=eof, we have an empty recordset
1841           boFoundPump = False
1842       Else
               'we found it
1843           boFoundPump = True
1844       End If

1845       If boFoundPump = False Then
               'not found in either database, try HP?
1846           sAns = MsgBox("Pump Not Found in the Database.  Look in Epicor?", vbYesNo, "Can't Find Pump")
1847           If sAns = vbNo Then     'new pump - don't get data from HP
1848               boUsingEpicor = False
1849           Else
1850               boUsingEpicor = True
       '            boUsingHP = False
1851           End If
       '        If boUsingEpicor = False Then
       '            sAns = MsgBox("Pump Not Found in the Database.  Look on the HP?", vbYesNo, "Can't Find Pump")
       '            If sAns = vbNo Then     'new pump - don't get data from HP
       '                 boUsingHP = False
       '            Else
       '                boUsingHP = True
       '            End If
       '        End If
1852           EnablePumpDataControls
1853           EnableTestSetupDataControls
1854           EnableTestDataControls
       '        BlankData               'clear any data on the screen
1855           cmdAddNewBalanceHoles.Visible = True

1856       End If

1857       If boFoundPump = True Then    'found the pump
1858           If rsPumpData.Fields("Approved") = True Then
1859               DisablePumpDataControls                         'if it's in the real database, don't allow changes here
1860               boPumpIsApproved = True
1861               lblPumpApproved.Visible = True
1862               If boCanApprove Then
1863                   cmdApprovePump.Caption = "Unapprove this pump"
1864               End If
1865               frmPLCData.cmdApproveTestDate.Enabled = True
1866           Else
1867               EnablePumpDataControls                          'it's in the temp database, allow changes
1868               boPumpIsApproved = False
1869               boTestDateIsApproved = False
1870               lblPumpApproved.Visible = False
1871               If boCanApprove Then
1872                   cmdApprovePump.Caption = "Approve this pump"
1873               End If
1874               cmdApproveTestDate.Caption = "You Must Approve Pump First"
1875               frmPLCData.cmdApproveTestDate.Enabled = False
1876           End If

               'found the pump, show the data
1877           txtModelNo.Text = rsPumpData.Fields("ModelNumber")
1878           frmPLCData.optMfr(0).value = rsPumpData.Fields("ChempumpPump")

1879           If rsPumpData.Fields("ChempumpPump") = True Then
1880               SetCombo cmbMotor, "Motor", rsPumpData
1881               SetCombo cmbDesignPressure, "DesignPressure", rsPumpData
1882               SetCombo cmbRPM, "RPM", rsPumpData
1883               SetCombo cmbCirculationPath, "CirculationPath", rsPumpData
1884               SetCombo cmbStatorFill, "StatorFill", rsPumpData
1885               SetCombo cmbModel, "Model", rsPumpData
1886               SetCombo cmbModelGroup, "ModelGroup", rsPumpData
1887               RatedKW = 999
1888           End If

               'set the TEMC data
1889           If rsPumpData.Fields("ChempumpPump") = False Then
1890               SetCombo cmbTEMCAdapter, "TEMCAdapter", rsPumpData
1891               SetCombo cmbTEMCAdditions, "TEMCAdditions", rsPumpData
1892               SetCombo cmbTEMCCirculation, "TEMCCirculation", rsPumpData
1893               SetCombo cmbTEMCDesignPressure, "TEMCDesignPressure", rsPumpData
1894               SetCombo cmbTEMCNominalDischargeSize, "TEMCNominalDischargeSize", rsPumpData
1895               SetCombo cmbTEMCDivisionType, "TEMCDivisionType", rsPumpData
1896               SetCombo cmbTEMCImpellerType, "TEMCImpellerType", rsPumpData
1897               SetCombo cmbTEMCInsulation, "TEMCInsulation", rsPumpData
1898               SetCombo cmbTEMCJacketGasket, "TEMCJacketGasket", rsPumpData
1899               SetCombo cmbTEMCMaterials, "TEMCMaterials", rsPumpData
1900               SetCombo cmbTEMCModel, "TEMCModel", rsPumpData
1901               SetCombo cmbTEMCNominalImpSize, "TEMCNominalImpSize", rsPumpData
1902               SetCombo cmbTEMCOtherMotor, "TEMCOtherMotor", rsPumpData
1903               SetCombo cmbTEMCPumpStages, "TEMCPumpStages", rsPumpData
1904               SetCombo cmbTEMCNominalSuctionSize, "TEMCNominalSuctionSize", rsPumpData
1905               SetCombo cmbTEMCTRG, "TEMCTRG", rsPumpData
1906               SetCombo cmbTEMCVoltage, "TEMCVoltage", rsPumpData
1907           End If

               'write ship to and bill to info
1908           If Not IsNull(rsPumpData.Fields("ShipToCustomer")) Then
1909               txtShpNo.Text = rsPumpData.Fields("ShipToCustomer")
1910           Else
1911               txtShpNo.Text = vbNullString
1912           End If

1913           If Not IsNull(rsPumpData.Fields("BillToCustomer")) Then
1914               txtBilNo.Text = rsPumpData.Fields("BillToCustomer")
1915           Else
1916               txtBilNo.Text = vbNullString
1917           End If

1918           sName = "ImpellerDia"
1919           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1920               sParam = rsPumpData.Fields(sName)
1921           Else
1922               sParam = vbNullString
1923           End If
1924           txtImpellerDia.Text = sParam

1925           sName = "DesignFlow"
1926           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1927               sParam = rsPumpData.Fields(sName)
1928           Else
1929               sParam = vbNullString
1930           End If
1931           txtDesignFlow.Text = sParam

1932           sName = "DesignTDH"
1933           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1934               sParam = rsPumpData.Fields(sName)
1935           Else
1936               sParam = vbNullString
1937           End If
1938           txtDesignTDH.Text = sParam

1939           sName = "SpGr"
1940           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1941               sParam = rsPumpData.Fields(sName)
1942           Else
1943               sParam = vbNullString
1944           End If
1945           txtSpGr.Text = sParam

1946           sName = "Remarks"
1947           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1948               sParam = rsPumpData.Fields(sName)
1949           Else
1950               sParam = vbNullString
1951           End If
1952           txtRemarks.Text = sParam

1953           sName = "SalesOrderNumber"
1954           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1955               sParam = rsPumpData.Fields(sName)
1956           Else
1957               sParam = vbNullString
1958           End If
1959           txtSalesOrderNumber.Text = sParam

1960           sName = "ApplicationFluid"
1961           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1962               sParam = rsPumpData.Fields(sName)
1963           Else
1964               sParam = vbNullString
1965           End If
1966           txtLiquid.Text = sParam

1967           sName = "NPSHFile"
1968           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1969               sParam = rsPumpData.Fields(sName)
1970           Else
1971               sParam = vbNullString
1972           End If
1973           txtNPSHFileLocation.Text = sParam

1974           sName = "RVSPartNo"
1975           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1976               sParam = rsPumpData.Fields(sName)
1977           Else
1978               sParam = vbNullString
1979           End If
1980           txtRVSPartNo.Text = sParam

1981           sName = "CustPN"
1982           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1983               sParam = rsPumpData.Fields(sName)
1984           Else
1985               sParam = vbNullString
1986           End If
1987           txtXPartNum.Text = sParam

1988           sName = "CustPO"
1989           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1990               sParam = rsPumpData.Fields(sName)
1991           Else
1992               sParam = vbNullString
1993           End If
1994           txtCustPONum.Text = sParam

               'make sure table has custpn - see if last three digits of model no are numeric
       '        sName = "SalesOrderNumber"
       '        If rsPumpData.Fields(sName).ActualSize <> 0 Then
       '            If IsNumeric(Right(rsPumpData.Fields("ModelNumber"), 3)) Then 'no sales order no, must be supermarket
       '                rsPumpData.Fields("CustPN") = rsPumpData.Fields("RVSPartNo")
       '            Else
       '                rsPumpData.Fields("CustPN") = rsPumpData.Fields("ModelNumber")
       '            End If
       '        End If

1995           sName = "ApplicationViscosity"
1996           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1997               sParam = Format(rsPumpData.Fields(sName), "#0.00")
1998           Else
1999               sParam = vbNullString
2000           End If
2001           txtViscosity.Text = sParam

       'added from TEMC Inspection Report
2002           sName = "JobNumber"
2003           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2004               sParam = rsPumpData.Fields(sName)
2005           Else
2006               sParam = ""
2007           End If
2008           txtJobNum.Text = sParam

2009           sName = "Phases"
2010           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2011               sParam = rsPumpData.Fields(sName)
2012           Else
2013               sParam = vbNullString
2014           End If
2015           txtNoPhases.Text = sParam

2016           sName = "ThermalClass"
2017           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2018               sParam = rsPumpData.Fields(sName)
2019           Else
2020               sParam = vbNullString
2021           End If
2022           txtThermalClass.Text = sParam

2023           sName = "ExpClass"
2024           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2025               sParam = rsPumpData.Fields(sName)
2026           Else
2027               sParam = vbNullString
2028           End If
2029           txtExpClass.Text = sParam

2030           sName = "NPSHr"
2031           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2032               sParam = rsPumpData.Fields(sName)
2033           Else
2034               sParam = vbNullString
2035           End If
2036           txtNPSHr.Text = sParam

2037           sName = "LiquidTemperature"
2038           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2039               sParam = rsPumpData.Fields(sName)
2040           Else
2041               sParam = vbNullString
2042           End If
2043           txtLiquidTemperature.Text = sParam

2044           sName = "RatedInputPower"
2045           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2046               sParam = rsPumpData.Fields(sName)
2047           Else
2048               sParam = vbNullString
2049           End If
2050           txtRatedInputPower.Text = sParam

2051           sName = "FLCurrent"
2052           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2053               sParam = rsPumpData.Fields(sName)
2054           Else
2055               sParam = vbNullString
2056           End If
2057           txtAmps.Text = sParam

2058           sName = "TEMCFrameNumber"
2059           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2060               sParam = rsPumpData.Fields(sName)
2061           Else
2062               sParam = vbNullString
2063           End If
2064           txtTEMCFrameNumber.Text = sParam

2065           optMfr(0).value = rsPumpData.Fields("ChempumpPump")
2066           optMfr(1).value = Not optMfr(0).value

2067           If rsPumpData.Fields("Field1") = "Feathered" Then
2068               Me.chkSuperMarketFeathered.value = Checked
2069           Else
2070               Me.chkSuperMarketFeathered.value = Unchecked
2071           End If

               'select the testsetup data
2072           qyTestSetup.ActiveConnection = cnPumpData
2073           qyTestSetup.CommandText = "SELECT * FROM TempTestSetupData WHERE (((TempTestSetupData.SerialNumber)='" & _
                                    txtSN.Text & "')) ORDER BY Date"
       '        qyTestSetup.CommandText = "SELECT * FROM TempTestSetupData WHERE (((TempTestSetupData.SerialNumber)='" & _
'               txtSN.Text & "'))"

2074           With rsTestSetup
2075               If .State = adStateOpen Then
2076                   .Close
2077               End If
2078               .CursorLocation = adUseClient
2079               .CursorType = adOpenStatic
2080               .Index = "FindData"
2081               .Open qyTestSetup
2082           End With


               'add the selection of dates to the Test Date combo box
2083           If rsTestSetup.RecordCount <> 0 Then
2084               For I = 0 To cmbTestDate.ListCount - 1
2085                   cmbTestDate.RemoveItem 0
2086               Next I
2087               rsTestSetup.MoveFirst
2088               For I = 1 To rsTestSetup.RecordCount
2089                   cmbTestDate.AddItem rsTestSetup.Fields("Date")
2090                   rsTestSetup.MoveNext
2091               Next I
2092               rsTestSetup.MoveFirst
2093               boFoundTestSetup = True

2094               If rsTestSetup.Fields("Approved") = True Then
2095                   DisableTestSetupDataControls                         'if it's in the real database, don't allow changes here
2096                   boTestDateIsApproved = True
2097                   lblTestDateApproved.Visible = True
2098                   If boCanApprove Then
2099                       cmdApproveTestDate.Caption = "Unapprove this Test Date"
2100                   End If
2101               Else
2102                   EnableTestSetupDataControls                          'it's in the temp database, allow changes
2103                   lblTestDateApproved.Visible = False
2104                   If boCanApprove Then
2105                       cmdApproveTestDate.Caption = "Approve this Test Date"
2106                   End If
2107               End If
2108               cmbTestDate.ListIndex = 0
2109           Else
2110               MsgBox ("There is no Test Setup Data for Serial Number " & txtSN.Text)
2111               boFoundTestSetup = False        'didn't find any data
2112               boFoundTestData = False
2113               cmbTestDate.AddItem Date        'load with today
2114               cmbTestDate.ListIndex = 0       'show the entry
2115               EnableTestSetupDataControls
2116               txtTestRemarks.Text = ""
2117               txtVibAx.Text = ""
2118               txtVibRad.Text = ""
2119               txtThrustBal.Text = ""
2120               txtTEMCTRGReading.Text = ""
2121               txtTEMCFrontThrust.Text = ""
2122               txtTEMCRearThrust.Text = ""
' <VB WATCH>
2123       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2124               Exit Sub
2125           End If

2126           If cmbTestDate.ListCount = 1 Then       'if there's only one test date, select it
2127           End If
' <VB WATCH>
2128       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2129           Exit Sub
2130       End If


2131       Do While boUsingEpicor = True   'need a do loop to exit
2132           If boUsingEpicor = True Then
                   'Dim MyRecord As SNRecord
2133               Dim MyRecord As SNRecord
           '            I = InStr(1, txtSN.Text, "-")
           '            If I > 0 Then
2134                   MyRecord = GetEpicorODBCData(txtSN.Text, EpicorConnectionString)
           '            End If
2135               If MyRecord.SONumber = "" Then
2136                   MsgBox ("Not found in Epicor")
2137                   boUsingEpicor = False
2138                   boEpicorFound = False
2139                   Exit Do
2140               End If

2141               If MyRecord.SONumber = 0 Then
2142                   boEpicorFound = False
2143                   boUsingSupermarketTable = True
2144                   boUsingEpicor = False
2145               Else
2146                   boEpicorFound = True
2147                   boUsingSupermarketTable = False
2148               End If

2149               If boEpicorFound = True Then
2150                   boUsingEpicor = False
       '                boEpicorFound = True
2151                   txtSalesOrderNumber.Text = MyRecord.SONumber
2152                   txtLineNumber.Text = MyRecord.SOLine
2153                   txtBilNo.Text = MyRecord.Customer
2154                   txtXPartNum.Text = MyRecord.XPartNum
2155                   txtCustPONum.Text = MyRecord.CustomerPO

2156                   If MyRecord.ShipTo = "" Then
2157                       txtShpNo.Text = MyRecord.Customer
2158                   Else
2159                       txtShpNo.Text = MyRecord.ShipTo
2160                   End If
2161                   txtModelNo.Text = MyRecord.PartNum
2162                   txtModelNo_Change
2163                   txtDesignTDH.Text = MyRecord.TDH
2164                   txtSpGr.Text = MyRecord.SpGr
2165                   txtImpellerDia.Text = MyRecord.ImpellerDiameter
2166                   txtDesignFlow.Text = MyRecord.Flow
2167                   txtNoPhases.Text = MyRecord.Phases
2168                   txtNPSHr.Text = MyRecord.NPSHr
2169                   txtRatedInputPower.Text = MyRecord.RatedInputPower
2170                   txtAmps.Text = MyRecord.FLCurrent
2171                   txtThermalClass.Text = MyRecord.ThermalClass
2172                   txtViscosity.Text = MyRecord.Viscosity
2173                   txtExpClass.Text = MyRecord.ExpClass
2174                   txtLiquidTemperature.Text = MyRecord.LiquidTemp
2175                   txtLiquid.Text = MyRecord.Fluid
2176                   txtJobNum.Text = MyRecord.JobNumber

2177                   For I = 0 To cmbStatorFill.ListCount - 1
2178                       If InStr(1, UCase$(MyRecord.StatorFill), UCase$(cmbStatorFill.List(I))) <> 0 Then
2179                           cmbStatorFill.ListIndex = I
2180                           Exit For
2181                       End If
2182                   Next I

2183                   For I = 0 To cmbCirculationPath.ListCount - 1
2184                       If InStr(1, UCase$(MyRecord.CirculationPath), UCase$(cmbCirculationPath.List(I))) <> 0 Then
2185                           cmbCirculationPath.ListIndex = I
2186                           Exit For
2187                       End If
2188                   Next I

2189                   For I = 0 To cmbDesignPressure.ListCount - 1
2190                       If InStr(1, MyRecord.DesignPressure, cmbDesignPressure.List(I)) <> 0 Then
2191                           cmbDesignPressure.ListIndex = I
2192                           Exit For
2193                       End If
2194                   Next I

2195                   For I = 0 To cmbVoltage.ListCount - 1
2196                       If InStr(1, MyRecord.Voltage, cmbVoltage.List(I)) <> 0 Then
2197                           cmbVoltage.ListIndex = I
2198                           Exit For
2199                       End If
2200                   Next I

2201                   For I = 0 To cmbFrequency.ListCount - 1
2202                       If InStr(1, MyRecord.Frequency, sName) <> 0 Then
2203                           cmbFrequency.ListIndex = I
2204                           Exit For
2205                       End If
2206                   Next I

2207                   For I = 0 To cmbRPM.ListCount - 1
2208                       If InStr(1, MyRecord.RPM, cmbRPM.List(I)) <> 0 Then
2209                           cmbRPM.ListIndex = I
2210                           Exit For
2211                       End If
2212                   Next I

2213                   For I = 0 To cmbSuctDia.ListCount - 1
2214                       If InStr(1, MyRecord.SuctFlangeSize, cmbSuctDia.List(I)) <> 0 Then
2215                           cmbSuctDia.ListIndex = I
2216                           Exit For
2217                       End If
2218                   Next I

2219                   For I = 0 To cmbDischDia.ListCount - 1
2220                       If InStr(1, MyRecord.DischFlangeSize, cmbDischDia.List(I)) <> 0 Then
2221                           cmbDischDia.ListIndex = I
2222                           Exit For
2223                       End If
2224                   Next I

2225                   For I = 0 To cmbTestSpec.ListCount - 1
2226                       If InStr(1, MyRecord.TestProcedure, cmbTestSpec.List(I)) <> 0 Then
2227                           cmbTestSpec.ListIndex = I
2228                           Exit For
2229                       End If
2230                   Next I

2231                   For I = 0 To cmbMotor.ListCount - 1
2232                       If InStr(1, MyRecord.MotorSize, cmbMotor.List(I)) <> 0 Then
2233                           cmbMotor.ListIndex = I
2234                           Exit For
2235                       End If
2236                   Next I


2237               End If
2238           End If
2239       Loop

2240       If boUsingSupermarketTable = True Then
2241           GetSuperMarketPump MyRecord.PartNum, MyRecord.JobNumber
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
2242       End If
' <VB WATCH>
2243       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2244       Exit Sub
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
2245       On Error GoTo vbwErrHandler
2246       Const VBWPROCNAME = "frmPLCData.cmdModifyBalanceHoleData_Click"
2247       If vbwProtector.vbwTraceProc Then
2248           Dim vbwProtectorParameterString As String
2249           If vbwProtector.vbwTraceParameters Then
2250               vbwProtectorParameterString = "()"
2251           End If
2252           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2253       End If
' </VB WATCH>
2254       Dim strInput As String
2255       Dim I As Integer
2256       Dim sNumber As Integer
2257       Dim sDia As String
2258       Dim sBC As String

2259       cmdModifyBalanceHoleData.Visible = False

2260       If dgBalanceHoles.SelBookmarks.Count = 0 Then
2261           cmdModifyBalanceHoleData.Visible = False
' <VB WATCH>
2262       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2263           Exit Sub
2264       End If

2265       rsBalanceHoles.MoveFirst
2266       rsBalanceHoles.Move dgBalanceHoles.SelBookmarks(0) - dgBalanceHoles.FirstRow

2267       sNumber = rsBalanceHoles!Number
2268       If rsBalanceHoles!diameter = 99 Then
2269           sDia = "Slot"
2270       Else
2271           sDia = str(rsBalanceHoles!diameter)
2272       End If
2273       If rsBalanceHoles!boltcircle = 99 Then
2274           sBC = "Unknown"
2275       Else
2276           sBC = str(rsBalanceHoles!boltcircle)
2277       End If


           'get the data for the balance holes
2278       strInput = InputBox("Enter Number of Holes (0 to delete entry)", , sNumber)
2279       If strInput = "" Then
2280           GoTo DeleteIt
2281       End If
2282       sNumber = CInt(strInput)
2283       If Val(sNumber) = 0 Then
2284           GoTo DeleteIt
2285       End If

2286       strInput = InputBox("Enter Decimal Value of Hole Diameter or 'Slot' (For Example, 0.675) ", , sDia)
2287       If strInput <> "" Then
2288           If UCase(strInput) = "SLOT" Then
2289               strInput = 99
2290           End If
2291           sDia = CSng(strInput)
2292       Else
2293           GoTo CancelPressed
2294       End If

2295       strInput = InputBox("Enter Decimal Value of Bolt Circle or 'Unknown' (For Example, 4.525)", , sBC)
2296       If strInput <> "" Then
2297           If UCase(strInput) = "UNKNOWN" Then
2298               strInput = 99
2299           End If
2300           sBC = CSng(strInput)
2301       Else
2302           GoTo CancelPressed
2303       End If

2304       rsBalanceHoles!Number = sNumber
2305       rsBalanceHoles!diameter = sDia
2306       rsBalanceHoles!boltcircle = sBC

2307       rsBalanceHoles.Update
           'rsBalanceHoles.Filter = "SerialNo = '" & frmPLCData.txtSN.Text & "'"

2308       GetBalanceHoleData txtSN.Text, cmbTestDate.Text
       '    rsBalanceHoles.Requery
2309       rsBalanceHoles.MoveLast
2310       dgBalanceHoles.Refresh
2311       chkBalanceHoles.value = 1
2312       rsBalanceHoles.MoveFirst

' <VB WATCH>
2313       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2314       Exit Sub

2315   CancelPressed:
2316       MsgBox "No New Balance Hole Data Entered", vbOKOnly

2317   DeleteIt:
2318       If (MsgBox("Do you really want to delete this entry?", vbYesNo, "Deleting Balance Hole Data. . .")) = vbYes Then
2319           rsBalanceHoles.Delete
2320           rsBalanceHoles.Update
2321           GetBalanceHoleData txtSN.Text, cmbTestDate.Text
       '        rsBalanceHoles.Requery
2322           If Not rsBalanceHoles.EOF Then
2323               rsBalanceHoles.MoveLast
2324           End If
2325           dgBalanceHoles.Refresh
2326           chkBalanceHoles.value = 1
2327           If Not rsBalanceHoles.BOF Then
2328               rsBalanceHoles.MoveFirst
2329           End If
2330       End If


' <VB WATCH>
2331       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2332       Exit Sub
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
2333       On Error GoTo vbwErrHandler
2334       Const VBWPROCNAME = "frmPLCData.cmdReport_Click"
2335       If vbwProtector.vbwTraceProc Then
2336           Dim vbwProtectorParameterString As String
2337           If vbwProtector.vbwTraceParameters Then
2338               vbwProtectorParameterString = "()"
2339           End If
2340           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2341       End If
' </VB WATCH>
2342       Dim I As Integer

2343       ExportToExcel

' <VB WATCH>
2344       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2345       Exit Sub
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
2346       On Error GoTo vbwErrHandler
2347       Const VBWPROCNAME = "frmPLCData.cmdSearchForPump_Click"
2348       If vbwProtector.vbwTraceProc Then
2349           Dim vbwProtectorParameterString As String
2350           If vbwProtector.vbwTraceParameters Then
2351               vbwProtectorParameterString = "()"
2352           End If
2353           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2354       End If
' </VB WATCH>
2355       LoadCombo frmSearch.cmbSearchModel, "TEMCHydraulics"

2356       frmSearch.Show
' <VB WATCH>
2357       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2358       Exit Sub
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
2359       On Error GoTo vbwErrHandler
2360       Const VBWPROCNAME = "frmPLCData.cmdSelectSupermarket_Click"
2361       If vbwProtector.vbwTraceProc Then
2362           Dim vbwProtectorParameterString As String
2363           If vbwProtector.vbwTraceParameters Then
2364               vbwProtectorParameterString = "()"
2365           End If
2366           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2367       End If
' </VB WATCH>
2368       grpSupermarket.Visible = False
' <VB WATCH>
2369       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2370       Exit Sub
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
2371       On Error GoTo vbwErrHandler
2372       Const VBWPROCNAME = "frmPLCData.cmdWriteSP_Click"
2373       If vbwProtector.vbwTraceProc Then
2374           Dim vbwProtectorParameterString As String
2375           If vbwProtector.vbwTraceParameters Then
2376               vbwProtectorParameterString = "()"
2377           End If
2378           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2379       End If
' </VB WATCH>
2380       Dim rc As String
2381       Dim S As String

           'write the set point data to the PLC
2382           bWrite = True
2383           S = Right$("0000" & txtWriteSPData, 4)
2384           S = Right$(S, 2) & Left$(S, 2)
2385           rc = StringToByteArray(S, ByteBuffer)

2386           DataLength = HexConvert(ByteBuffer, 2)
2387           DataAddress = StringToHexInt("2005")

2388           rc = GetData

2389           bWrite = False
' <VB WATCH>
2390       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2391       Exit Sub
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
2392       On Error GoTo vbwErrHandler
2393       Const VBWPROCNAME = "frmPLCData.btnRunNPSH_Click"
2394       If vbwProtector.vbwTraceProc Then
2395           Dim vbwProtectorParameterString As String
2396           If vbwProtector.vbwTraceParameters Then
2397               vbwProtectorParameterString = "()"
2398           End If
2399           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2400       End If
' </VB WATCH>
2401       Static OriginalColor As Long
2402       If btnRunNPSH.Caption = "Run NPSH" Then
2403           btnRunNPSH.Caption = "Cancel NPSH Run"
2404           OriginalColor = btnRunNPSH.BackColor
2405           tmrNPSHr.Enabled = False
2406           btnRunNPSH.BackColor = vbRed
2407           If boCanApprove Then
2408               txtNPSH(5).Visible = True
2409               lbltab4(5).Visible = True
2410           Else
2411               txtNPSH(5).Visible = False
2412               lbltab4(5).Visible = False
2413           End If
2414           WroteNPSHr = False

2415           frmNPSH.Visible = True
2416           txtNPSH(5).Enabled = True
2417           If Val(txtTDH.Text) <= 10 Then
2418               MsgBox "This test will not work starting with this starting TDH.  Ending test...", vbOKOnly, "Flow is 0"
2419               btnRunNPSH.Caption = "Run NPSH"
2420               btnRunNPSH.BackColor = OriginalColor
2421               frmNPSH.Visible = False
' <VB WATCH>
2422       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2423               Exit Sub
2424           End If
               'load initial values
2425           If DataGrid2.Row = -1 Then
2426               MsgBox "You must write the normal test data to this row before you run NPSH.", vbOKOnly, "Nothing written for this row"
2427               btnRunNPSH.Caption = "Run NPSH"
2428               btnRunNPSH.BackColor = OriginalColor
2429               frmNPSH.Visible = False
' <VB WATCH>
2430       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2431               Exit Sub
2432           Else
2433               DataGrid2.Row = UpDown1.value - 1
2434           End If

2435           txtNPSH(0).Text = DataGrid2.Columns("Flow")
2436           txtNPSH(3).Text = DataGrid2.Columns("TDH")
2437           txtNPSH(4) = 0
               'txtNPSH(0).Text = txtFlow.Text
               'txtNPSH(3).Text = txtTDH.Text
2438           txtNPSH(4) = 0
2439       Else
2440           btnRunNPSH.Caption = "Run NPSH"
2441           btnRunNPSH.BackColor = OriginalColor
2442           frmNPSH.Visible = False
2443       End If

           'ReportToExcel
' <VB WATCH>
2444       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2445       Exit Sub
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
2446       On Error GoTo vbwErrHandler
2447       Const VBWPROCNAME = "frmPLCData.updown1_change"
2448       If vbwProtector.vbwTraceProc Then
2449           Dim vbwProtectorParameterString As String
2450           If vbwProtector.vbwTraceParameters Then
2451               vbwProtectorParameterString = "()"
2452           End If
2453           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2454       End If
' </VB WATCH>
2455       Dim sName As String

2456       If Not rsTestData.BOF Then
2457           rsTestData.MoveFirst
2458       End If

2459       If Not rsTestData.BOF Or Not rsTestData.EOF Then
2460           rsTestData.Move UpDown1.value - 1
2461       End If

2462       sName = "VibrationX"
2463       If rsTestData.Fields(sName).ActualSize <> 0 Then
2464           txtVibAx.Text = rsTestData.Fields(sName)
2465       Else
       '        txtVibAx.Text = vbNullString
2466       End If

2467       sName = "VibrationY"
2468       If rsTestData.Fields(sName).ActualSize <> 0 Then
2469           txtVibRad.Text = rsTestData.Fields(sName)
2470       Else
       '        txtVibRad.Text = vbNullString
2471       End If

2472       sName = "Remarks"
2473       If rsTestData.Fields(sName).ActualSize <> 0 Then
2474           txtTestRemarks.Text = rsTestData.Fields(sName)
2475       Else
       '        txtTestRemarks.Text = vbNullString
2476       End If

2477       sName = "ThrustBalance"
2478       If rsTestData.Fields(sName).ActualSize <> 0 Then
2479           txtThrustBal.Text = rsTestData.Fields(sName)
2480       Else
       '        txtThrustBal.Text = vbNullString
2481       End If

2482       sName = "TEMCTRG"
2483       If rsTestData.Fields(sName).ActualSize <> 0 Then
2484           txtTEMCTRGReading.Text = rsTestData.Fields(sName)
2485       Else
2486           txtTEMCTRGReading.Text = 0
       '        txtTEMCTRGReading.Text = vbNullString
2487       End If

2488       sName = "TEMCFrontThrust"
2489       If rsTestData.Fields(sName).ActualSize <> 0 Then
2490           txtTEMCFrontThrust.Text = rsTestData.Fields(sName)
2491       Else
       '        txtTEMCFrontThrust.Text = vbNullString
2492       End If

2493       sName = "TEMCRearThrust"
2494       If rsTestData.Fields(sName).ActualSize <> 0 Then
2495           txtTEMCRearThrust.Text = rsTestData.Fields(sName)
2496       Else
       '        txtTEMCRearThrust.Text = vbNullString
2497       End If
2498       sName = "TEMCMomentArm"
2499       If rsTestData.Fields(sName).ActualSize <> 0 Then
2500           txtTEMCMomentArm.Text = rsTestData.Fields(sName)
2501       Else
       '        txtTEMCMomentArm.Text = vbNullString
2502       End If
2503       sName = "TEMCThrustRigPressure"
2504       If rsTestData.Fields(sName).ActualSize <> 0 Then
2505           txtTEMCThrustRigPressure.Text = rsTestData.Fields(sName)
2506       Else
       '        txtTEMCThrustRigPressure.Text = vbNullString
2507       End If
2508       sName = "TEMCViscosity"
2509       If rsTestData.Fields(sName).ActualSize <> 0 And rsTestData.Fields(sName) <> 0 Then
2510           txtTEMCViscosity.Text = rsTestData.Fields(sName)
2511       Else
       '        txtTEMCViscosity.Text = vbNullString
2512       End If

2513       CalculateTEMCForce

2514       rsEff.MoveFirst
2515       rsEff.Move UpDown1.value - 1
' <VB WATCH>
2516       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2517       Exit Sub
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
2518       On Error GoTo vbwErrHandler
2519       Const VBWPROCNAME = "frmPLCData.CalculateTEMCForce"
2520       If vbwProtector.vbwTraceProc Then
2521           Dim vbwProtectorParameterString As String
2522           If vbwProtector.vbwTraceParameters Then
2523               vbwProtectorParameterString = "()"
2524           End If
2525           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2526       End If
' </VB WATCH>
2527       Dim NoOfPoles As Integer
2528       Dim Frequency As Integer
2529       Dim Additions As String
2530       Dim Frame As String
2531       Dim VOverA As Double
2532       Dim Force As Double
2533       Dim Gravity As Double

2534       If Val(txtSpGr.Text) = 0 Then
2535           Gravity = 1
2536       Else
2537           Gravity = CDbl(Val(txtSpGr.Text))
2538       End If

           'show calculated values
2539       If Val(txtTEMCFrontThrust.Text) = 0 Then
2540           If Val(txtTEMCRearThrust.Text) = 0 Then
               'no thrust entered
2541               lblTEMCFrontRear.Visible = False
2542               txtTEMCCalcForce.Text = " "
2543           Else
                   'rear thrust
2544               txtTEMCCalcForce.Text = Format(Gravity * (Val(txtTEMCRearThrust.Text) * Val(txtTEMCMomentArm.Text) - (Val(txtTEMCThrustRigPressure.Text) / 14.223) * 4.5), "##0.0")
2545               lblTEMCFrontRear.Caption = "REAR"
2546               lblTEMCFrontRear.Visible = True
2547           End If
2548       Else
               'front thrust
2549           txtTEMCCalcForce.Text = Format(Gravity * (Val(txtTEMCFrontThrust.Text) * Val(txtTEMCMomentArm.Text) + (Val(txtTEMCThrustRigPressure.Text) / 14.223) * 4.5), "##0.0")
2550           lblTEMCFrontRear.Caption = "FRONT"
2551           lblTEMCFrontRear.Visible = True
2552       End If

2553       If Val(txtTEMCCalcForce.Text) < 0 Then
2554           txtTEMCCalcForce.Text = -txtTEMCCalcForce
2555           lblTEMCFrontRear.Caption = "FRONT"
2556       End If

           'see how many poles we have, it's the next to last number in the frame size
2557       If Len(txtTEMCFrameNumber) > 2 Then
2558           NoOfPoles = 2 * Val(Left$(Right$(txtTEMCFrameNumber.Text, 2), 1))
2559       End If

2560       If cmbTEMCAdditions.ListIndex <> -1 Then
2561           Additions = Mid$(cmbTEMCAdditions.List(cmbTEMCAdditions.ListIndex), 2, 1)
2562           If Additions = "A" Or Additions = "E" Or Additions = "G" Or Additions = "J" Then
2563               Frequency = 60
2564           ElseIf Additions = "B" Or Additions = "F" Or Additions = "H" Or Additions = "K" Then
2565               Frequency = 50
2566           Else
2567               Frequency = 0
2568           End If
2569       End If

2570       If Len(txtTEMCFrameNumber.Text) = 3 Then
2571           If txtTEMCFrameNumber.Text = "529" Then
2572               Frame = "420"
2573           Else
2574               Frame = Left$(txtTEMCFrameNumber, 2) & "0"
2575           End If
2576       Else
2577           Frame = txtTEMCFrameNumber.Text
2578           If Right$(txtTEMCFrameNumber.Text, 1) = "5" Then
2579               Frame = Frame & Left$(lblTEMCFrontRear.Caption, 1)
2580           Else
2581           End If
2582       End If
2583       Force = DLookupA(3, TEMCForceViscosity, 1, Frame)
2584       If Frequency = 60 Then
2585           Force = Force / 1.2
2586       End If
2587       If Val(txtTEMCViscosity.Text) > 1# Then
2588           If (Val(txtTEMCCalcForce.Text) > 3 * Force) Then
2589               lblTEMCPassFail.Visible = True
2590               lblTEMCPassFail.ForeColor = vbRed
2591               lblTEMCPassFail.Caption = "FAIL"
2592           Else
2593               lblTEMCPassFail.Visible = True
2594               lblTEMCPassFail.ForeColor = vbGreen
2595               lblTEMCPassFail.Caption = "PASS"
2596           End If
2597       End If

2598       If (Val(txtTEMCViscosity.Text) > 0.5) And (Val(txtTEMCViscosity.Text) <= 1#) Then
2599           If (Val(txtTEMCCalcForce.Text) > 2 * Force) Then
2600               lblTEMCPassFail.Visible = True
2601               lblTEMCPassFail.ForeColor = vbRed
2602               lblTEMCPassFail.Caption = "FAIL"
2603           Else
2604               lblTEMCPassFail.Visible = True
2605               lblTEMCPassFail.ForeColor = vbGreen
2606               lblTEMCPassFail.Caption = "PASS"
2607           End If
2608       End If

2609       If (Val(txtTEMCViscosity.Text) > 0.3) And (Val(txtTEMCViscosity.Text) <= 0.5) Then
2610           If (Val(txtTEMCCalcForce.Text) > 1.5 * Force) Then
2611               lblTEMCPassFail.Visible = True
2612               lblTEMCPassFail.ForeColor = vbRed
2613               lblTEMCPassFail.Caption = "FAIL"
2614           Else
2615               lblTEMCPassFail.Visible = True
2616               lblTEMCPassFail.ForeColor = vbGreen
2617               lblTEMCPassFail.Caption = "PASS"
2618           End If
2619       End If

2620       If (Val(txtTEMCViscosity.Text) <= 0.3) Then
2621           If (Val(txtTEMCCalcForce.Text) > 1# * Force) Then
2622               lblTEMCPassFail.Visible = True
2623               lblTEMCPassFail.ForeColor = vbRed
2624               lblTEMCPassFail.Caption = "FAIL"
2625           Else
2626               lblTEMCPassFail.Visible = True
2627               lblTEMCPassFail.ForeColor = vbGreen
2628               lblTEMCPassFail.Caption = "PASS"
2629           End If
2630       End If
2631       If NoOfPoles <> 0 Then
2632           VOverA = (DLookupA(2, TEMCForceViscosity, 1, Frame)) / (NoOfPoles * 30 / Frequency)
2633       End If
       '    If Frequency = 60 Then
       '        VOverA = VOverA * 1.2
       '    End If

2634       txtTEMCPVValue.Text = Format(Val(txtTEMCCalcForce.Text) * VOverA, "##0.0")

2635       If Val(txtTEMCFrontThrust.Text) = 0 And Val(txtTEMCRearThrust.Text) = 0 Then
2636           txtTEMCPVValue.Text = ""
2637           txtTEMCCalcForce.Text = ""
2638           lblTEMCPassFail.Visible = False
2639       End If


           'calculate reverse head
2640       txtRevHead.Text = Format(rsTestData.Fields("RBHPress") - rsTestData.Fields("SuctionPressure") * 2.31, "##0.0")
       '    txtRevHead.Text = Format((CDbl(Val(txtAI3Display.Text)) - CDbl(Val(txtSuctionDisplay.Text))) * 2.31, "##0.0")

' <VB WATCH>
2641       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2642       Exit Sub
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
2643       On Error GoTo vbwErrHandler
2644       Const VBWPROCNAME = "frmPLCData.updown2_change"
2645       If vbwProtector.vbwTraceProc Then
2646           Dim vbwProtectorParameterString As String
2647           If vbwProtector.vbwTraceParameters Then
2648               vbwProtectorParameterString = "()"
2649           End If
2650           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2651       End If
' </VB WATCH>
2652       Dim Plothead(1, 7) As Single
2653       Dim HeadPlot(7, 1) As Single

2654       Dim PlotEff() As Single
2655       Dim PlotKW() As Single
2656       Dim PlotAmps() As Single

2657       Dim j As Integer

2658       For j = 0 To UpDown2.value - 1
2659           Plothead(0, j) = HeadFlow(0, j)
2660           Plothead(1, j) = HeadFlow(1, j)
2661           HeadPlot(j, 0) = FlowHead(j, 0)
2662           HeadPlot(j, 1) = FlowHead(j, 1)
       '        ReDim Preserve PlotEff(1, j)
       '        PlotEff(0, j) = EffFlow(0, j)
       '        PlotEff(1, j) = EffFlow(1, j)
       '        ReDim Preserve PlotKW(1, j)
       '        PlotKW(0, j) = KWFlow(0, j)
       '        PlotKW(1, j) = KWFlow(1, j)
       '        ReDim Preserve PlotAmps(1, j)
       '        PlotAmps(0, j) = AmpsFlow(0, j)
       '        PlotAmps(1, j) = AmpsFlow(1, j)
2663       Next j

2664       MSChart1 = HeadPlot

' <VB WATCH>
2665       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2666       Exit Sub
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
2667       On Error GoTo vbwErrHandler
2668       Const VBWPROCNAME = "frmPLCData.DataGrid1_AfterColUpdate"
2669       If vbwProtector.vbwTraceProc Then
2670           Dim vbwProtectorParameterString As String
2671           If vbwProtector.vbwTraceParameters Then
2672               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ColIndex", ColIndex) & ") "
2673           End If
2674           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2675       End If
' </VB WATCH>
2676       DoEfficiencyCalcs
' <VB WATCH>
2677       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2678       Exit Sub
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
2679       On Error GoTo vbwErrHandler
2680       Const VBWPROCNAME = "frmPLCData.dgBalanceHoles_SelChange"
2681       If vbwProtector.vbwTraceProc Then
2682           Dim vbwProtectorParameterString As String
2683           If vbwProtector.vbwTraceParameters Then
2684               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Cancel", Cancel) & ") "
2685           End If
2686           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2687       End If
' </VB WATCH>
2688       If dgBalanceHoles.SelBookmarks.Count = 0 Then
2689           cmdModifyBalanceHoleData.Visible = False
2690       Else
2691           cmdModifyBalanceHoleData.Visible = True
2692       End If
' <VB WATCH>
2693       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2694       Exit Sub
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
2695       On Error GoTo vbwErrHandler
2696       Const VBWPROCNAME = "frmPLCData.Form_Activate"
2697       If vbwProtector.vbwTraceProc Then
2698           Dim vbwProtectorParameterString As String
2699           If vbwProtector.vbwTraceParameters Then
2700               vbwProtectorParameterString = "()"
2701           End If
2702           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2703       End If
' </VB WATCH>
2704       If ProgramEnd = True Then
2705           Unload Me
2706       End If
' <VB WATCH>
2707       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2708       Exit Sub
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
2709       On Error GoTo vbwErrHandler
2710       Const VBWPROCNAME = "frmPLCData.Form_Load"
2711       If vbwProtector.vbwTraceProc Then
2712           Dim vbwProtectorParameterString As String
2713           If vbwProtector.vbwTraceParameters Then
2714               vbwProtectorParameterString = "()"
2715           End If
2716           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2717       End If
' </VB WATCH>
2718       Dim RetVal As String
2719       Dim sSendStr As String
2720       Dim I As Integer
2721       Dim j As Integer
2722       Dim sTableName As String
2723       Dim WhichServer As String
2724       Dim WhichDatabase As String

2725       ProgramEnd = False
2726       Dim objWMIService As Object
2727       Dim colProcesses As Object
2728       Set objWMIService = GetObject("winmgmts:")
2729       Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name LIKE 'PolarRundown%'")
       '    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name LIKE 'Excel%'")
2730       If colProcesses.Count > 1 Then
2731           MsgBox "There is already a copy of Polar Rundown running.  You can only have one copy running at a time", vbOKOnly, "Polar Rundown already running"
2732           Dim f As Form
2733           For Each f In Forms
2734               If f.Name <> Me.Name Then
2735                    Unload f
2736               End If
2737           Next
2738           ProgramEnd = True
' <VB WATCH>
2739       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2740           Exit Sub
2741       Else
2742       End If
2743       Set objWMIService = Nothing
2744       Set colProcesses = Nothing

2745       debugging = 0   'assume not debugging
2746       WhichServer = "Production"     'change to production server
2747       WhichDatabase = "Production"

2748       If UCase$(Left$(GetMachineName, 5)) = "MROSE" Or UCase$(Left$(GetMachineName, 5)) = "ITTES" Then  'if mickey, see if we want to be in debug
2749           I = MsgBox("Debug?", vbYesNo)
2750           If I = vbYes Then
2751               debugging = 1
2752               WhichServer = "Production"
2753               WhichDatabase = "Production"
2754           Else
2755           End If
2756       End If

2757       If debugging Then
       '        GoTo temp
2758       End If
           'see if the mdb file is where it's supposed to be

2759       Dim developmentDatabase As String
2760       developmentDatabase = GetUNCFromLetter("F:") & sDevelopmentDatabase

2761       If Dir(developmentDatabase) = "" Then
2762           MsgBox "Development.mdb does not exist on F:, Please contact IT.", , "No Development Database"
2763           End
2764       End If

           'get the database info from the new mdb file
2765       Dim cnDevelopment As New ADODB.Connection
2766       Dim qyDevelopment As New ADODB.Command
2767       Dim rsDevelopment As New ADODB.Recordset

2768       On Error GoTo CannotConnect

2769       With cnDevelopment
2770           .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & developmentDatabase & ";Persist Security Info=False; Jet OLEDB:Database Password=Access7277word;"
2771           .ConnectionTimeout = 10
2772           .Open
2773       End With

2774   On Error GoTo vbwErrHandler
2775       GoTo Connected

2776   CannotConnect:
2777       MsgBox "Cannot connect with Development.mdb database.  Please contact IT.", , "Cannot find Connection data."
2778       End

2779   Connected:

           'we're connected, get the data for the Epicor SQL server
2780       qyDevelopment.CommandText = "SELECT * FROM Connections WHERE Connections.WhichServer = '" & WhichServer & "' AND WhichDatabase = '" & WhichDatabase & "'"
2781       qyDevelopment.ActiveConnection = cnDevelopment

2782       rsDevelopment.CursorLocation = adUseClient
2783       rsDevelopment.CursorType = adOpenStatic
2784       rsDevelopment.LockType = adLockOptimistic

2785       On Error GoTo NoServerData

2786       rsDevelopment.Open qyDevelopment

2787   On Error GoTo vbwErrHandler
2788       GoTo GotServerData

2789   NoServerData:

2790       MsgBox "Cannot connect with Development.mdb database.  Please contact IT.", , "Cannot find Connection data."
2791       End

2792   GotServerData:

2793       If rsDevelopment.RecordCount <> 1 Then
2794           GoTo NoServerData
2795       End If

           'construct Epicor connection string
2796       EpicorConnectionString = "Driver={" & rsDevelopment.Fields("ODBCDriver") & "};" & _
                                         "Database=" & rsDevelopment.Fields("DatabaseName") & ";" & _
                                         "Server=" & rsDevelopment.Fields("ServerName") & ";" & _
                                         "UID=" & rsDevelopment.Fields("UserName") & ";" & _
                                         "PWD=" & rsDevelopment.Fields("UserPassword") & ";"


           'make sure we can open the SQL database

2797       On Error GoTo CannotOpenEpicorSQLServer

2798       Dim cnTestEpicor As New ADODB.Connection
2799       cnTestEpicor.ConnectionString = EpicorConnectionString
2800       cnTestEpicor.Open
2801       cnTestEpicor.Close
2802       Set cnTestEpicor = Nothing
2803   On Error GoTo vbwErrHandler

2804       GoTo FoundEpicorSQLServer

2805   CannotOpenEpicorSQLServer:
2806       MsgBox "Cannot connect with the Epicor SQL server specified in Development.mdb.  Please contact IT.", , "Cannot connect with Epicor SQL Server"
2807       End

2808   FoundEpicorSQLServer:
           'get data on rundown database
2809       rsDevelopment.Close
2810       qyDevelopment.CommandText = "SELECT * FROM Connections WHERE Connections.WhichServer = 'PolarRundown'"

2811       On Error GoTo NoRundownDatabase

2812       rsDevelopment.Open qyDevelopment

2813       GoTo FoundRundownDatabase

2814   NoRundownDatabase:
2815       MsgBox "Cannot connect with the Pump Rundown database specified in Development.mdb.  Please contact IT.", , "Cannot connect with Epicor SQL Server"
2816       End

2817   FoundRundownDatabase:
2818       If rsDevelopment.RecordCount <> 1 Then
2819           GoTo NoRundownDatabase
2820           End
2821       End If

2822   temp:

2823       If debugging Then
2824           sDataBaseName = "c:\databases\PolarData.mdb"
2825       Else

2826          sDataBaseName = GetUNCFromLetter("F:") & "\Groups\Shared\databases\PolarData.mdb"

       '        sDataBaseName = rsDevelopment.Fields("ServerName") & rsDevelopment.Fields("DatabaseName")

       '        sDataBaseName = sServerName & "f\groups\shared\databases\PumpData 2k.mdb"
2827       End If

2828       Dim tempFSO As Object
2829       Set tempFSO = CreateObject("Scripting.FileSystemObject")
2830       ParentDirectoryName = tempFSO.getparentfoldername(sDataBaseName)
2831       Set tempFSO = Nothing

           'see if we can open the pump rundown database
2832       On Error GoTo NoRundownDatabase
2833       With cnPumpData
       '        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataBaseName & ";Persist Security Info=False;Jet OLEDB:Database Password=185TitusAve"
2834           .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataBaseName & ";Persist Security Info=False;"
2835           .ConnectionTimeout = 10
2836           .Open
2837       End With
2838   On Error GoTo vbwErrHandler


2839       If debugging = 0 Then
       '        Printer.Orientation = vbPRORLandscape
2840       End If

2841       lblVersion = "Polar Rundown - Version " & App.Major & "." & App.Minor & "." & App.Revision
2842       frmPLCData.Caption = "Polar Rundown"

2843       boFoundPump = False

2844       Me.Show

2845       MSChart1.Plot.Axis(VtChAxisIdX).AxisTitle = "Flow"
2846       MSChart1.Plot.Axis(VtChAxisIdY).AxisTitle = "TDH"
           'MSChart1.Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen = True
           'MSChart1.Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen = True
2847       MSChart1.Plot.UniformAxis = False
2848       MSChart1.Plot.SeriesCollection.Item(1).SeriesMarker.Auto = False
2849       MSChart1.Plot.SeriesCollection.Item(1).Pen.Width = 5
2850       With MSChart1.Plot.SeriesCollection.Item(1).DataPoints.Item(-1).Marker
2851           .Visible = True
2852           .Size = 50
2853           .Style = VtMarkerStyleCircle
2854           .FillColor.Automatic = False
2855           .FillColor.Set 0, 0, 255
2856       End With
2857       MSChart1.Plot.AutoLayout = False
2858       MSChart1.Plot.LocationRect.Max.x = 5600
2859       MSChart1.Plot.LocationRect.Max.y = 2800
2860       MSChart1.Plot.LocationRect.Min.x = 0
2861       MSChart1.Plot.LocationRect.Min.y = 0

           'assure that the timers are off
2862       frmPLCData.tmrGetDDE.Enabled = False

2863       frmPLCData.tmrStartUp.Enabled = False

           'initialize the PLC network
2864       RetVal = NetWorkInitialize()
2865       If RetVal <> 0 Then
2866           MsgBox ("Can't Initialize Network. Exiting...")
2867           End
2868       End If

2869       If debugging = 0 Then
               'load array of plcs
2870           I = 0
2871           Open rsDevelopment.Fields("ServerName") & "PolarPLCAddresses.txt" For Input As 1
2872           While Not EOF(1)
2873               Input #1, Description(I)
2874               For j = 0 To 125
2875                   Input #1, aDevices(I).Address(j)
2876               Next j
2877               Input #1, j
2878               I = I + 1
2879           Wend
2880           Close #1

2881           DeviceCount = I

2882           If Left$(GetMachineName, 2) = "WV" Then  'if in WV, put MWSC first in loop dropdown
2883               Dim k As Integer
2884               For k = 0 To DeviceCount - 1
2885                   If InStr(Description(k), "MWSC") <> 0 Then
2886                       Exit For
2887                   End If
2888               Next k
2889               Description(DeviceCount) = Description(0)
2890               Description(0) = Description(k)
2891               Description(k) = Description(DeviceCount)

2892               aDevices(DeviceCount) = aDevices(0)
2893               aDevices(0) = aDevices(k)
2894               aDevices(k) = aDevices(DeviceCount)

2895           End If

2896           Dim PLCAddress As String
2897           For I = 0 To DeviceCount - 1
2898               PLCAddress = aDevices(I).Address(4) & "." & aDevices(I).Address(5) & "." & aDevices(I).Address(6) & "." & aDevices(I).Address(7)
2899               RetVal = PingSilent(PLCAddress)
2900               If RetVal <> 0 Then
2901                   frmPLCData.cmbPLCLoop.AddItem Description(I)
2902                   frmPLCData.cmbPLCLoop.ItemData(frmPLCData.cmbPLCLoop.NewIndex) = I
2903               End If
2904           Next I
2905       End If

2906       frmPLCData.cmbPLCLoop.AddItem "Add PLC Data Manually"   'enable the controls for manual entry

           'turn on the PLC led

2907       frmPLCData.cmbPLCLoop.ListIndex = 0
2908       frmPLCData.tmrGetDDE.Enabled = True

           'hook up to the various databases

           'copy the template of the database here
           'see if it exists
2909       Dim fdrive As String
2910       fdrive = GetUNCFromLetter("F:")
2911       If Dir(fdrive & "\groups\shared\databases" & sEffDataBaseName) = "" Then
2912           MsgBox "File does not exist at " & fdrive & "\groups\shared\databases" & sEffDataBaseName & ". Please contact IT", vbOKOnly, "Eff.mdb does not exist"
2913       Else
               'Dim FSO As New FileSystemObject
2914           FileCopy fdrive & "\groups\shared\databases" & sEffDataBaseName, App.Path & sEffDataBaseName
2915       End If


2916       With cnEffData
2917           .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & sEffDataBaseName & ";Persist Security Info=False"
2918           .Open
2919       End With

           'open some recordsets
2920       rsPumpData.Index = "SerialNumber"
2921       rsTestSetup.Index = "FindData"
2922       rsTestData.Index = "PrimaryKey"
2923       rsPumpData.Open "TempPumpData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
2924       rsTestSetup.Open "TempTestSetupData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
2925       rsTestData.Filter = "SerialNumber = ''"
2926       rsTestData.CursorLocation = adUseClient
2927       rsTestData.Open "TempTestData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
2928       rsEff.CursorLocation = adUseClient
2929       rsEff.Open "Efficiency", cnEffData, adOpenStatic, adLockOptimistic, adCmdTableDirect
2930       qyBalanceHoles.ActiveConnection = cnPumpData
2931       rsBalanceHoles.CursorLocation = adUseClient
2932       rsBalanceHoles.CursorType = adOpenStatic
2933       rsBalanceHoles.LockType = adLockOptimistic
2934       qyMisc.ActiveConnection = cnPumpData
2935       qyMisc.CommandText = "SELECT MiscParameters.ParameterName, MiscParameters.ParameterValue From MiscParameters WHERE (((MiscParameters.ParameterName)='AllowableTDHVariation'));"
2936       rsMisc.CursorLocation = adUseClient
2937       rsMisc.CursorType = adOpenStatic
2938       rsMisc.LockType = adLockBatchOptimistic
2939       rsMisc.Open qyMisc
2940       txtNPSH(5).Text = rsMisc!ParameterValue

2941       If debugging <> 1 Then
2942           FindMagtrols
2943       Else
2944           cmbMagtrol.AddItem "Add Manually"
2945           cmbMagtrol.ItemData(cmbMagtrol.NewIndex) = 99
2946           cmbMagtrol.ListIndex = 0
2947       End If
2948       optKW(1).value = True
2949       optKW_Click (1)


           'blank out data grid
2950       Set DataGrid1.DataSource = rsTestData

           'load the combo boxes
2951       LoadCombo cmbStatorFill, "StatorFill"
2952       LoadCombo cmbCirculationPath, "CirculationPath"
2953       LoadCombo cmbVoltage, "Voltage"
2954       LoadCombo cmbFrequency, "Frequency"
2955       LoadCombo cmbMotor, "Motor"
2956       LoadCombo cmbDesignPressure, "DesignPressure"
2957       LoadCombo cmbRPM, "RPM"
2958       LoadCombo cmbOrificeNumber, "OrificeNumber"
2959       LoadCombo cmbTestSpec, "TestSpecification"
2960       LoadCombo cmbLoopNumber, "LoopNumber"
2961       LoadCombo cmbSuctDia, "SuctionDiameter"
2962       LoadCombo cmbDischDia, "DischargeDiameter"
2963       LoadCombo cmbTachID, "TachID"
2964       LoadCombo cmbAnalyzerNo, "AnalyzerNo"
2965       LoadCombo cmbModel, "Model"
2966       LoadCombo cmbModelGroup, "ModelGroup"
2967       LoadCombo cmbMounting, "Mounting"
2968       LoadCombo cmbPLCNo, "PLCNo"
2969       LoadCombo cmbFlowMeter, "PumpFlowMeter"
2970       LoadCombo cmbSuctionPressureTransducer, "SuctionPressureTransducer"
2971       LoadCombo cmbDischargePressureTransducer, "DischargePressureTransducer"
2972       LoadCombo cmbTemperatureTransducer, "TemperatureTransducer"
2973       LoadCombo cmbCirculationFlowMeter, "CirculationFlowMeter"
           'LoadCombo cmbSupermarketModel, "SupermarketPumpData"

           'load the TEMC combo boxes, too
2974       LoadCombo cmbTEMCAdapter, "TEMCAdapter"
2975       LoadCombo cmbTEMCAdditions, "TEMCAdditions"
2976       LoadCombo cmbTEMCCirculation, "TEMCCirculation"
2977       LoadCombo cmbTEMCDesignPressure, "TEMCDesignPressure"
2978       LoadCombo cmbTEMCNominalDischargeSize, "TEMCNominalDischargeSize"
2979       LoadCombo cmbTEMCDivisionType, "TEMCDivisionType"
2980       LoadCombo cmbTEMCImpellerType, "TEMCImpellerType"
2981       LoadCombo cmbTEMCInsulation, "TEMCInsulation"
2982       LoadCombo cmbTEMCJacketGasket, "TEMCJacketGasket"
2983       LoadCombo cmbTEMCMaterials, "TEMCMaterials"
2984       LoadCombo cmbTEMCModel, "TEMCModel"
2985       LoadCombo cmbTEMCNominalImpSize, "TEMCNominalImpSize"
2986       LoadCombo cmbTEMCOtherMotor, "TEMCOtherMotor"
2987       LoadCombo cmbTEMCNominalSuctionSize, "TEMCNominalSuctionSize"
2988       LoadCombo cmbTEMCVoltage, "TEMCVoltage"
2989       LoadCombo cmbTEMCPumpStages, "TEMCPumpStages"
2990       LoadCombo cmbTEMCTRG, "TEMCTRG"

           'LoadCombo frmSearch.cmbSearchModel, "Model"

           'fill memory arrays for dlookups
2991       FillArrays

           'choose the first tab
2992       frmPLCData.SSTab1.Tab = 0

           'set the grid column names
2993       Dim c As Column
2994       For Each c In DataGrid1.Columns
2995           Select Case c.DataField
               Case "TestDataID"
2996               c.Visible = False
2997           Case "SerialNumber"
2998               c.Visible = False
2999           Case "Date"
3000               c.Visible = False
3001           Case Else ' Show all other columns.
3002               c.Visible = True
3003               c.Alignment = dbgRight
3004           End Select
3005       Next c

3006       Set dgBalanceHoles.DataSource = rsBalanceHoles

3007       For Each c In dgBalanceHoles.Columns
3008           Select Case c.DataField
               Case "BalanceHoleID"
3009               c.Visible = False
3010           Case "SerialNo"
3011               c.Visible = False
3012           Case "Date"
3013               c.Visible = True
3014               c.Alignment = dbgCenter
3015               c.Width = 2000
3016           Case "Number"
3017               c.Visible = True
3018               c.Alignment = dbgCenter
3019               c.Width = 700
3020           Case "Diameter"
3021               c.Visible = False
3022           Case "Diameter1"
3023               c.Caption = "Diameter"
3024               c.Visible = True
3025               c.Alignment = dbgCenter
3026               c.Width = 700
3027           Case "BoltCircle1"
3028               c.Caption = "Bolt Circle"
3029               c.Visible = True
3030               c.Alignment = dbgCenter
3031               c.Width = 800
3032           Case "BoltCircle"
3033               c.Visible = False
3034           Case "SetNo"
3035               c.Visible = False
3036           Case Else ' Show all other columns.
3037               c.Visible = False
3038           End Select
3039       Next c

3040       BlankData

       '    If debugging <> 1 Then
               'get user initials
3041           frmLogin.Show
       '    End If

3042     optMfr(1).value = True
3043     frmMfr.Visible = False

3044       Pressed = True
' <VB WATCH>
3045       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3046       Exit Sub
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
3047       On Error GoTo vbwErrHandler
3048       Const VBWPROCNAME = "frmPLCData.Form_Unload"
3049       If vbwProtector.vbwTraceProc Then
3050           Dim vbwProtectorParameterString As String
3051           If vbwProtector.vbwTraceParameters Then
3052               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Cancel", Cancel) & ") "
3053           End If
3054           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3055       End If
' </VB WATCH>
3056       End
' <VB WATCH>
3057       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3058       Exit Sub
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
3059       On Error GoTo vbwErrHandler
3060       Const VBWPROCNAME = "frmPLCData.Label15_Click"
3061       If vbwProtector.vbwTraceProc Then
3062           Dim vbwProtectorParameterString As String
3063           If vbwProtector.vbwTraceParameters Then
3064               vbwProtectorParameterString = "()"
3065           End If
3066           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3067       End If
' </VB WATCH>
3068       frmDiagram.Show
' <VB WATCH>
3069       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3070       Exit Sub
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
3071       On Error GoTo vbwErrHandler
3072       Const VBWPROCNAME = "frmPLCData.lblAutoMan_Click"
3073       If vbwProtector.vbwTraceProc Then
3074           Dim vbwProtectorParameterString As String
3075           If vbwProtector.vbwTraceParameters Then
3076               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3077           End If
3078           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3079       End If
' </VB WATCH>

3080       Dim blnEnabled As Boolean

3081       If lblAutoMan(Index).Caption = "Auto" Then
3082           lblAutoMan(Index).Caption = "Man"
3083           blnEnabled = True
3084       Else
3085           lblAutoMan(Index).Caption = "Auto"
3086           blnEnabled = False
3087       End If

3088       Select Case Index
               Case 0
3089               txtFlowDisplay.Enabled = blnEnabled
3090           Case 1
3091               txtSuctionDisplay.Enabled = blnEnabled
3092           Case 2
3093               txtDischargeDisplay.Enabled = blnEnabled
3094           Case 3
3095               txtTemperatureDisplay.Enabled = blnEnabled
3096           Case 4
3097               txtAI1Display.Enabled = blnEnabled
3098           Case 5
3099               txtAI2Display.Enabled = blnEnabled
3100           Case 6
3101               txtAI3Display.Enabled = blnEnabled
3102           Case 7
3103               txtAI4Display.Enabled = blnEnabled
3104       End Select

' <VB WATCH>
3105       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3106       Exit Sub
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
3107       On Error GoTo vbwErrHandler
3108       Const VBWPROCNAME = "frmPLCData.tmrNPSHr_Timer"
3109       If vbwProtector.vbwTraceProc Then
3110           Dim vbwProtectorParameterString As String
3111           If vbwProtector.vbwTraceParameters Then
3112               vbwProtectorParameterString = "()"
3113           End If
3114           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3115       End If
' </VB WATCH>
3116       tmrNPSHr.Enabled = False
3117       If frmNPSH.Visible = True Then
3118           btnRunNPSH_Click    'close test
3119       End If
' <VB WATCH>
3120       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3121       Exit Sub
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
3122       On Error GoTo vbwErrHandler
3123       Const VBWPROCNAME = "frmPLCData.txtNPSH_Change"
3124       If vbwProtector.vbwTraceProc Then
3125           Dim vbwProtectorParameterString As String
3126           If vbwProtector.vbwTraceParameters Then
3127               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3128           End If
3129           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3130       End If
' </VB WATCH>
3131       If Index = 5 Then
3132           If frmNPSH.Visible = True Then
3133               If rsMisc.State = adStateOpen Then
3134                   rsMisc.Close
3135               End If
3136               rsMisc.CursorLocation = adUseClient
3137               rsMisc.Open "Select * from MiscParameters WHERE (ParameterName = 'AllowableTDHVariation');", cnPumpData, adOpenStatic, adLockOptimistic, adCmdText
3138               rsMisc.Fields("ParameterValue").value = txtNPSH(5).Text
3139               rsMisc.Update
3140           End If
3141       End If
' <VB WATCH>
3142       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3143       Exit Sub
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
3144       On Error GoTo vbwErrHandler
3145       Const VBWPROCNAME = "frmPLCData.txtNPSHFileLocation_Click"
3146       If vbwProtector.vbwTraceProc Then
3147           Dim vbwProtectorParameterString As String
3148           If vbwProtector.vbwTraceParameters Then
3149               vbwProtectorParameterString = "()"
3150           End If
3151           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3152       End If
' </VB WATCH>
3153       Dim sTempDir As String
3154       On Error Resume Next
3155       sTempDir = CurDir    'Remember the current active directory
3156       CommonDialog2.DialogTitle = "Select a directory" 'titlebar
3157       CommonDialog2.InitDir = "\\tei-main-01\f\en\groups\shared\calibration and rundown\npsh\" 'start dir, might be "C:\" or so also
3158       CommonDialog2.filename = "Select a Directory"  'Something in filenamebox
3159       CommonDialog2.Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
3160       CommonDialog2.Filter = "Directories|*.~#~" 'set files-filter to show dirs only
3161       CommonDialog2.CancelError = True 'allow escape key/cancel
3162       CommonDialog2.ShowSave   'show the dialog screen

3163       If Err <> 32755 Then    ' User didn't chose Cancel.
               'Me.SDir.Text = CurDir
3164       End If

       '    ChDir sTempDir  'restore path to what it was at entering

3165   Me.txtNPSHFileLocation.Text = CommonDialog2.filename

' <VB WATCH>
3166       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3167       Exit Sub
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
3168       On Error GoTo vbwErrHandler
3169       Const VBWPROCNAME = "frmPLCData.txtTitle_LostFocus"
3170       If vbwProtector.vbwTraceProc Then
3171           Dim vbwProtectorParameterString As String
3172           If vbwProtector.vbwTraceParameters Then
3173               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3174           End If
3175           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3176       End If
' </VB WATCH>

3177       ChangeTitles Index

' <VB WATCH>
3178       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3179       Exit Sub
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
3180       On Error GoTo vbwErrHandler
3181       Const VBWPROCNAME = "frmPLCData.ChangeTitles"
3182       If vbwProtector.vbwTraceProc Then
3183           Dim vbwProtectorParameterString As String
3184           If vbwProtector.vbwTraceParameters Then
3185               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ChannelNo", ChannelNo) & ") "
3186           End If
3187           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3188       End If
' </VB WATCH>
3189       Dim I As Integer
3190       Dim S As String

3191       If txtTitle(ChannelNo).Locked = True Then
' <VB WATCH>
3192       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3193           Exit Sub
3194       End If

3195       Dim qy As New ADODB.Command
3196       Dim rs As New ADODB.Recordset

3197       qy.ActiveConnection = cnPumpData

           'see if we have an entry in the table
3198       qy.CommandText = "SELECT * FROM AITitles " & _
                             "WHERE (((AITitles.SerialNo)= '" & txtSN.Text & "') " & _
                             "AND ((AITitles.Date)= #" & cmbTestDate.Text & "#) " & _
                             "AND ((AITitles.Channel)=" & ChannelNo & "));"

3199       With rs     'open the recordset for the query
3200           .CursorLocation = adUseClient
3201           .CursorType = adOpenStatic
3202           .LockType = adLockOptimistic
3203           .Open qy
3204       End With

3205       If (rs.BOF = True And rs.EOF = True) Then  'new record
3206           rs.AddNew
3207           rs.Fields("SerialNo") = txtSN.Text
3208           rs.Fields("Date") = cmbTestDate.Text
3209           rs.Fields("Channel") = CByte(ChannelNo)
3210           rs.Fields("Title") = txtTitle(ChannelNo).Text
3211           rs.Update
3212       Else    'we have an entry, modify it
3213           rs.Fields("SerialNo") = txtSN.Text
3214           rs.Fields("Date") = cmbTestDate.Text
3215           rs.Fields("Channel") = CByte(ChannelNo)
3216           rs.Fields("Title") = txtTitle(ChannelNo).Text
3217           rs.Update
3218       End If

3219       rs.Close
3220       Set rs = Nothing
3221       Set qy = Nothing

' <VB WATCH>
3222       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3223       Exit Sub
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
3224       On Error GoTo vbwErrHandler
3225       Const VBWPROCNAME = "frmPLCData.optKW_Click"
3226       If vbwProtector.vbwTraceProc Then
3227           Dim vbwProtectorParameterString As String
3228           If vbwProtector.vbwTraceParameters Then
3229               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3230           End If
3231           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3232       End If
' </VB WATCH>
3233       Select Case Index
               Case 0  'add 3 powers
3234               txtKW.Enabled = False
3235           Case 1  'enter kw
3236               txtKW.Enabled = True
3237           Case 2  'use analog in 4
3238               txtKW.Enabled = False
3239       End Select
' <VB WATCH>
3240       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3241       Exit Sub
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
3242       On Error GoTo vbwErrHandler
3243       Const VBWPROCNAME = "frmPLCData.optMfr_Click"
3244       If vbwProtector.vbwTraceProc Then
3245           Dim vbwProtectorParameterString As String
3246           If vbwProtector.vbwTraceParameters Then
3247               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3248           End If
3249           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3250       End If
' </VB WATCH>
3251       frmTEMC.Visible = optMfr(1).value
3252       frmChempump.Visible = optMfr(0).value
3253       frmTEMCData.Visible = optMfr(1).value
3254       txtModelNo_Change
' <VB WATCH>
3255       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3256       Exit Sub
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
3257       On Error GoTo vbwErrHandler
3258       Const VBWPROCNAME = "frmPLCData.tmrGetDDE_Timer"
3259       If vbwProtector.vbwTraceProc Then
3260           Dim vbwProtectorParameterString As String
3261           If vbwProtector.vbwTraceParameters Then
3262               vbwProtectorParameterString = "()"
3263           End If
3264           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3265       End If
' </VB WATCH>

       'get here every second... get plc and magtrol data

3266       Dim sSendStr As String
3267       Dim I As Integer
3268       Dim VoltMul As Double

3269       If Calibrating Then
' <VB WATCH>
3270       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3271           Exit Sub
3272       End If

3273       If debugging Then
               'Exit Sub
3274       End If


3275       If boPLCOperating = True Then
3276           frmPLCData.shpGetPLCData.Visible = True    'turn the PLC led on

               'convert the plc data into real numbers
               'the following data are type real
3277           txtFlow.Text = ConvertToReal("4050")
3278           txtSuction.Text = ConvertToReal("4052")
3279           txtDischarge.Text = ConvertToReal("4054")
3280           txtTemperature.Text = ConvertToReal("4056")

3281           txtValvePosition.Text = ConvertToLong("2004")

3282           frmPLCData.txtTC1.Text = ConvertToLong("2200")
3283           frmPLCData.txtTC2.Text = ConvertToLong("2202")
3284           frmPLCData.txtTC3.Text = ConvertToLong("2204")
3285           frmPLCData.txtTC4.Text = ConvertToLong("2206")

3286           frmPLCData.txtAI1.Text = ConvertToReal("4060")
3287           frmPLCData.txtAI2.Text = ConvertToReal("4062")
3288           frmPLCData.txtAI3.Text = ConvertToReal("4064")
3289           frmPLCData.txtAI4.Text = ConvertToReal("4066")

3290           frmPLCData.txtPCoef.Text = ConvertToLong("4036")
3291           frmPLCData.txtICoef.Text = ConvertToLong("4037")
3292           frmPLCData.txtDCoef.Text = ConvertToLong("4040")

3293           frmPLCData.txtSetPoint.Text = ConvertToLong("4035")
3294           frmPLCData.txtInHg.Text = ConvertToLong("1460")


               'modify the data from PLC format to format that we can use
               'and update the screen
3295           If txtFlowDisplay.Enabled = False Then
3296               frmPLCData.txtFlowDisplay = Format$(txtFlow.Text, "###0.00")
3297           End If
3298           If txtSuctionDisplay.Enabled = False Then
3299               frmPLCData.txtSuctionDisplay = Format$((txtSuction.Text) / 10, "##0.00")
3300           End If
3301           If txtDischargeDisplay.Enabled = False Then
3302               frmPLCData.txtDischargeDisplay = Format$(txtDischarge.Text, "##0.00")
3303           End If
3304           If txtTemperatureDisplay.Enabled = False Then
3305               frmPLCData.txtTemperatureDisplay = Format$(txtTemperature.Text, "##0.00")
3306           End If
3307           frmPLCData.txtValvePositionDisplay = (txtValvePosition.Text)

3308           frmPLCData.txtTC1Display = Format$((txtTC1.Text) / 10, "##0.0")
3309           frmPLCData.txtTC2Display = Format$((txtTC2.Text) / 10, "##0.0")
3310           frmPLCData.txtTC3Display = Format$((txtTC3.Text) / 10, "##0.0")
3311           frmPLCData.txtTC4Display = Format$((txtTC4.Text) / 10, "##0.0")

3312           If txtAI1Display.Enabled = False Then
3313               frmPLCData.txtAI1Display = Format$(txtAI1.Text, "##0.00")
3314           End If
3315           If txtAI2Display.Enabled = False Then
3316               frmPLCData.txtAI2Display = Format$(txtAI2.Text, "##0.00")
3317           End If
3318           If txtAI3Display.Enabled = False Then
3319               frmPLCData.txtAI3Display = Format$(txtAI3.Text, "##0.00")
3320           End If
3321           If txtAI4Display.Enabled = False Then
3322               frmPLCData.txtAI4Display = Format$(txtAI4.Text, "##0.00")
3323           End If

3324           frmPLCData.txtSetPointDisplay = (txtSetPoint.Text)

3325           frmPLCData.txtInHgDisplay = Format$(txtInHg.Text / 100, "00.00")

3326           frmPLCData.shpGetPLCData.Visible = False   'turn the PLC led off

3327           frmPLCData.shpGetMagtrolData.Visible = True 'turn the Magtrol led on
3328       End If

3329       If boMagtrolOperating = True Then


               'get the data from the Magtrol
3330           If Right(cmbMagtrol.List(cmbMagtrol.ListIndex), 4) = "5300" Then
3331               sSendStr = vbCrLf
3332               sData = Space$(68)
3333               VoltMul = Sqr(3)
3334           Else
3335               sSendStr = "OT" & vbCrLf
3336               sData = Space$(183)
3337               VoltMul = 1#
3338           End If

3339           On Error GoTo noresponse
3340           If UsingNatInst Then
3341               ibwrt iUD, sSendStr
3342               ibrd iUD, sData

                   'parse the Magrol response
       '            vResponse = CWGPIB1.Tasks("Number Parser").Parse(sData)
3343           Else
                   'Dim Databack As String
3344               sData = TCP.SendGetData("OT")
3345           End If

3346               Dim vSplit() As String
3347               vSplit = Split(Right(sData, Len(sData) - 1), ",")
3348               ReDim vResponse(UBound(vSplit))
3349               For I = 0 To UBound(vSplit) - 1
3350                   vResponse(I) = CDbl(vSplit(I))
3351               Next I

               'format the parsed response
3352           Dim dd As String
3353           dd = "- -"

3354           If Not IsEmpty(vResponse) Then
               '8 entries for 5300 and 12 for the 6530
3355               If UBound(vResponse) = 8 Or UBound(vResponse) = 12 Then
                       'put the responses into the correct text box
3356                   txtV1.Text = Format$(VoltMul * vResponse(1), "###0.0")   'we get back phase voltage and we want line voltage

3357                   Select Case vResponse(0)
                           Case Is < 1
3358                           txtI1.Text = Format$(vResponse(0), "0.0000")
3359                       Case Is < 10
3360                           txtI1.Text = Format$(vResponse(0), "0.000")
3361                       Case Is < 100
3362                           txtI1.Text = Format$(vResponse(0), "00.00")
3363                       Case Else
3364                           txtI1.Text = Format$(vResponse(0), "000.0")
3365                   End Select

3366                   Select Case vResponse(3)
                           Case Is < 1
3367                           txtI2.Text = Format$(vResponse(3), "0.0000")
3368                       Case Is < 10
3369                           txtI2.Text = Format$(vResponse(3), "0.000")
3370                       Case Is < 100
3371                           txtI2.Text = Format$(vResponse(3), "00.00")
3372                       Case Else
3373                           txtI2.Text = Format$(vResponse(3), "000.0")
3374                   End Select

3375                   Select Case vResponse(6)
                           Case Is < 1
3376                           txtI3.Text = Format$(vResponse(6), "0.0000")
3377                       Case Is < 10
3378                           txtI3.Text = Format$(vResponse(6), "0.000")
3379                       Case Is < 100
3380                           txtI3.Text = Format$(vResponse(6), "00.00")
3381                       Case Else
3382                           txtI3.Text = Format$(vResponse(6), "000.0")
3383                   End Select

3384                   txtP1.Text = Format$(vResponse(2) / 1000, "##0.00")     '/ by 1000 to show kW
3385                   txtV2.Text = Format$(VoltMul * vResponse(4), "###0.0")
                       'txtI2.Text = Format$(vResponse(3), "###0.0")
3386                   txtP2.Text = Format$(vResponse(5) / 1000, "##0.00")
3387                   txtV3.Text = Format$(VoltMul * vResponse(7), "###0.0")
                       'txtI3.Text = Format$(vResponse(6), "###0.0")
3388                   txtP3.Text = Format$(vResponse(8) / 1000, "##0.00")
3389                   If (vResponse(0) * vResponse(1) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7)) <> 0 Then
                           'if we have some measured current
                           'pf = sum of power/sum of VA
3390                       If Right(cmbMagtrol.List(cmbMagtrol.ListIndex), 4) = "5300" Then
                               'add kw responses and / by 1000 to get to kW
3391                           txtKW.Text = (vResponse(2) + vResponse(5) + vResponse(8)) / 1000
3392                           txtPF.Text = Format$(100 * (vResponse(2) + vResponse(5) + vResponse(8)) / (vResponse(1) * vResponse(0) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7)), "0.00")
3393                       Else
3394                           txtKW.Text = (vResponse(2) + vResponse(8)) / 1000
3395                           txtPF.Text = Format$(100 * (vResponse(2) + vResponse(8)) / ((Sqr(3) / 3) * (vResponse(1) * vResponse(0) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7))), "0.00")
3396                       End If
3397                       Select Case Val(txtKW.Text)
                               Case Is < 1
3398                               txtKW.Text = Format$(txtKW.Text, "0.00000")
3399                           Case Is < 10
3400                               txtKW.Text = Format$(txtKW.Text, "0.0000")
3401                           Case Is < 100
3402                               txtKW.Text = Format$(txtKW.Text, "00.000")
3403                           Case Else
3404                               txtKW.Text = Format$(txtKW.Text, "000.00")
3405                       End Select
3406                   Else
3407                       txtPF = dd
3408                   End If
3409               Else
                       'no response, show all -- in text boxes
3410                   txtV1.Text = dd
3411                   txtI1.Text = dd
3412                   txtP1.Text = dd
3413                   txtV2.Text = dd
3414                   txtI2.Text = dd
3415                   txtP2.Text = dd
3416                   txtV3.Text = dd
3417                   txtI3.Text = dd
3418                   txtP3.Text = dd
3419                   txtPF = dd
3420                   txtKW = dd
3421               End If
3422           End If
3423       Else    'magtrol not operating
3424           Dim dbl As Double

3425           If optKW(0).value = True Then   'add 3 powers
3426               txtKW.Text = Val(txtP1.Text) + Val(txtP2.Text) + Val(txtP3.Text)
3427           End If
3428           If optKW(1).value = True Then   'enter kw
3429               txtP1.Text = Val(txtKW.Text) / 3
3430               txtP2.Text = Val(txtKW.Text) / 3
3431               txtP3.Text = Val(txtKW.Text) / 3
3432           End If
3433           If optKW(2).value = True Then   'use ai4
3434               txtKW.Text = txtAI4Display.Text
3435               txtP1.Text = Val(txtKW.Text) / 3
3436               txtP2.Text = Val(txtKW.Text) / 3
3437               txtP3.Text = Val(txtKW.Text) / 3
3438           End If

3439           dbl = Val(txtV1.Text) * Val(txtI1.Text)
3440           dbl = dbl + Val(txtV2.Text) * Val(txtI2.Text)
3441           dbl = dbl + Val(txtV3.Text) * Val(txtI3.Text)
3442           If dbl <> 0 Then
3443               txtPF.Text = Format$((Val(txtKW.Text) * 1000 * 3 * 100 / (dbl * Sqr(3))), "0.00")
3444           End If
3445       End If

3446   noresponse:
3447   On Error GoTo vbwErrHandler
3448       frmPLCData.shpGetMagtrolData.Visible = False   'turn the Magtrol led off

           'update the little PLC chart
3449       For I = 1 To 99
3450           vPlot(0, I) = vPlot(0, I + 1)
3451           vPlot(1, I) = vPlot(1, I + 1)
3452       Next I
3453       vPlot(0, 100) = txtSetPointDisplay
3454       vPlot(1, 100) = txtFlowDisplay

           'do NPSH stuff
3455       Dim SuctVelHead As Single
3456       Dim DischVelHead As Single
3457       Dim Conversion As Single
3458       Dim SuctionPSIA As Single
3459       Dim DischargePSIA As Single
3460       Dim VaporPress As Single
3461       Dim SpecVolume As Single
3462       Dim NPSHa As Single
3463       Dim NPSHr As Single
3464       Dim TDH As Single
3465       Dim pd As Single


           'velocity head
3466       If cmbSuctDia.ListIndex = -1 Then   'if no suction diameter chosen
3467           SuctVelHead = 0
3468       Else
       '        pd = DLookup("ActualDia", "PipeDiameters", "ID = " & cmbSuctDia.ListIndex + 1)
3469           pd = DLookupA(ActualColNo, PipeDiameters, IDColNo, cmbSuctDia.ItemData(cmbSuctDia.ListIndex) + 1)
3470           SuctVelHead = (0.002592 * Val(txtFlow) ^ 2) / (pd ^ 4)
3471       End If

3472       If cmbDischDia.ListIndex = -1 Then     'if no discharge diameter chosen
3473           DischVelHead = 0
3474       Else
       '        pd = DLookup("ActualDia", "PipeDiameters", "ID = " & cmbDischDia.ListIndex + 1)
3475           pd = DLookupA(ActualColNo, PipeDiameters, IDColNo, cmbDischDia.ItemData(cmbDischDia.ListIndex) + 1)
3476           DischVelHead = (0.002592 * Val(txtFlow) ^ 2) / (pd ^ 4)
3477       End If

           'convert gauges to absolute
3478       If txtInHgDisplay.Text = "" Then
3479           Conversion = 0
3480       Else
3481           Conversion = txtInHgDisplay * 0.491
3482       End If

3483       SuctionPSIA = Val(txtSuctionDisplay) + Conversion
3484       DischargePSIA = Val(txtDischargeDisplay) + Conversion


           'lookup vapor pressure and specific volume in the arrays that we made
           'if temp is out of range, say so and exit
3485       If Val(txtTemperatureDisplay) < 40 Or Val(txtTemperatureDisplay) > 165 Then
3486           txtNPSHa = 0
' <VB WATCH>
3487       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3488           Exit Sub
3489       Else
3490           I = Val(txtTemperatureDisplay) - 40
       '        VaporPress = DLookup("VaporPressure", "VaporPressure", "ID = " & I)
       '        SpecVolume = DLookup("SpecificVolume", "VaporPressure", "ID= " & I)
3491           VaporPress = DLookupA(VaporPressureColNo, VaporPressure, IDColNo, I)
3492           SpecVolume = DLookupA(SpecificVolumeColNo, VaporPressure, IDColNo, I)
3493       End If

3494       If Not ((txtSuctHeight = "") Or (txtDischHeight = "") Or Not IsNumeric(txtSuctHeight) Or Not IsNumeric(txtDischHeight)) Then
               'NPSHa
3495           NPSHa = (144 * SpecVolume * (SuctionPSIA - VaporPress)) + (txtSuctHeight / 12) + SuctVelHead
       '        NPSHa = CalcTDH(DischargePSIA, SuctionPSIA, 0, DischVelHead, 0, txtTemperature)
3496           txtNPSHa = Format$(NPSHa, "##0.00")

               'tdh
3497           TDH = CalcTDH(DischargePSIA, SuctionPSIA, 0, (DischVelHead - SuctVelHead), (txtDischHeight / 12) - (txtSuctHeight / 12), txtTemperatureDisplay)
3498           txtTDH = Format$(TDH, "##0.00")

3499           If frmNPSH.Visible = True Then
3500               If Val(txtTDH.Text) > 0 Then
3501                   txtNPSH(2).Text = Format(100 * Val(txtTDH.Text) / Val(txtNPSH(3).Text), "##0.00")
3502                   txtNPSH(1).Text = Format(100 * Val(txtFlow.Text) / Val(txtNPSH(0).Text), "##0.00")
                       'check for tdh variation
3503                   If Abs(Val(txtNPSH(1)) - 100) > Val(txtNPSH(5).Text) Then
3504                       MsgBox "The TDH value has varied more than " & txtNPSH(5) & " %. NPSHr data will NOT be written to the data table", vbOKOnly, "TDH variation too large"
3505                       btnRunNPSH_Click
3506                   Else    'tdh variation small
3507                       If Val(txtNPSH(2).Text) <= 97 Then
                               'btnRunNPSH_Click
                               'write the npsh and save
3508                           If WroteNPSHr = False Then
3509                               txtNPSH(4).Text = txtNPSHa.Text
3510                               rsTestData!NPSHr = txtNPSHa.Text
3511                               rsTestData.Update
3512                               rsEff!NPSHr = txtNPSHa.Text
3513                               rsEff.Update
3514                               WroteNPSHr = True
3515                               tmrNPSHr.Interval = 5000
3516                               tmrNPSHr.Enabled = True
3517                           End If
3518                       End If  'val < 97
3519                   End If  'check for tdh variation
3520               End If 'val tdh <=0
3521           Else    'frm not visible
                   'txtNPSHa = Format$(0, "##0.00")
3522           End If  'if frm visible

3523       Else
3524           txtNPSHa = 0
3525       End If
' <VB WATCH>
3526       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3527       Exit Sub
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
3528       On Error GoTo vbwErrHandler
3529       Const VBWPROCNAME = "frmPLCData.tmrStartUp_Timer"
3530       If vbwProtector.vbwTraceProc Then
3531           Dim vbwProtectorParameterString As String
3532           If vbwProtector.vbwTraceParameters Then
3533               vbwProtectorParameterString = "()"
3534           End If
3535           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3536       End If
' </VB WATCH>
3537       tmrStartUp.Enabled = False
' <VB WATCH>
3538       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3539       Exit Sub
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
3540       On Error GoTo vbwErrHandler
3541       Const VBWPROCNAME = "frmPLCData.SetCombo"
3542       If vbwProtector.vbwTraceProc Then
3543           Dim vbwProtectorParameterString As String
3544           If vbwProtector.vbwTraceParameters Then
3545               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
3546               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sName", sName) & ", "
3547               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("rs", rs) & ") "
3548           End If
3549           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3550       End If
' </VB WATCH>

3551       Dim I As Integer
3552       Dim sParam As String
3553       Dim qy As New ADODB.Command
3554       Dim rs1 As New ADODB.Recordset

3555       If rs.Fields(sName).ActualSize <> 0 Then     'if there's an entry
3556           sParam = rs.Fields(sName)                'get the index number
3557           qy.ActiveConnection = cnPumpData
3558           qy.CommandText = "SELECT * FROM " & sName & " WHERE " & sName & " = " & sParam
3559           Set rs1 = qy.Execute()                                  'get the record for the index number

3560           If rs1.BOF = True And rs1.EOF = True Then
3561               cmbComboName.ListIndex = -1                             'else, remove any pointer
' <VB WATCH>
3562       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3563               Exit Function
3564           End If

3565           For I = 0 To cmbComboName.ListCount - 1                     'go through the combobox entries
3566               If cmbComboName.ItemData(I) = rs1.Fields(0) Then     'see when we find the desired index number
3567                   cmbComboName.ListIndex = I                                              'if we do, set the combo box
3568                   Exit For                                            'and we're done
3569               End If
3570               cmbComboName.ListIndex = -1                             'else, remove any pointer
3571           Next I
3572       Else
3573           cmbComboName.ListIndex = -1
3574       End If

' <VB WATCH>
3575       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3576       Exit Function
' <VB WATCH>
3577       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3578       Exit Function
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
3579       On Error GoTo vbwErrHandler
3580       Const VBWPROCNAME = "frmPLCData.SetComboTestSetup"
3581       If vbwProtector.vbwTraceProc Then
3582           Dim vbwProtectorParameterString As String
3583           If vbwProtector.vbwTraceParameters Then
3584               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
3585               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sFieldName", sFieldName) & ", "
3586               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sTableName", sTableName) & ", "
3587               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("rs", rs) & ") "
3588           End If
3589           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3590       End If
' </VB WATCH>

       'same as setcombo, except here we also pass in the field name

3591       Dim I As Integer
3592       Dim sParam As String
3593       Dim qy As New ADODB.Command
3594       Dim rs1 As New ADODB.Recordset

3595       If rs.Fields(sFieldName).ActualSize <> 0 Then
               'if plc number, adjust plcaddress id numbers 1 and 2 to plc 8 and 9 respectively
3596           If sTableName = "CirculationFlowMeter" Then
                   'sParam = rs.Fields(sFieldName) + 7
3597               sParam = rs.Fields(sFieldName)
3598               If Val(sParam) < 4 Then
3599                   sParam = str(Val(sParam) + 4)
3600                   rs.Fields(sFieldName) = sParam
3601               End If
3602           Else
3603               sParam = rs.Fields(sFieldName)
3604           End If
3605           qy.ActiveConnection = cnPumpData
3606           qy.CommandText = "SELECT * FROM " & sTableName & " WHERE " & sTableName & " = " & sParam
3607           Set rs1 = qy.Execute()

3608           For I = 0 To cmbComboName.ListCount - 1
3609               If cmbComboName.ItemData(I) = rs1.Fields(0) Then
3610                   cmbComboName.ListIndex = I
3611                   Exit For
3612               End If
3613               cmbComboName.ListIndex = -1
3614           Next I
3615       Else
3616           cmbComboName.ListIndex = -1
3617       End If

' <VB WATCH>
3618       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3619       Exit Function
' <VB WATCH>
3620       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3621       Exit Function
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
3622       On Error GoTo vbwErrHandler
3623       Const VBWPROCNAME = "frmPLCData.DisablePumpDataControls"
3624       If vbwProtector.vbwTraceProc Then
3625           Dim vbwProtectorParameterString As String
3626           If vbwProtector.vbwTraceParameters Then
3627               vbwProtectorParameterString = "()"
3628           End If
3629           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3630       End If
' </VB WATCH>

3631       txtSalesOrderNumber.Enabled = False
3632       frmMfr.Enabled = False
3633       txtShpNo.Enabled = False
3634       txtBilNo.Enabled = False
3635       txtDesignFlow.Enabled = False
3636       txtDesignTDH.Enabled = False

3637       frmMiscPumpData.Enabled = False

3638       txtModelNo.Enabled = False
3639       txtImpellerDia.Enabled = False

3640       frmTEMC.Enabled = False
3641       frmChempump.Enabled = False

3642       txtRemarks.Enabled = False
3643       Me.cmdAddNewTestDate.Visible = False

3644       cmdEnterPumpData.Enabled = False

' <VB WATCH>
3645       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3646       Exit Sub
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
3647       On Error GoTo vbwErrHandler
3648       Const VBWPROCNAME = "frmPLCData.DisableTestSetupDataControls"
3649       If vbwProtector.vbwTraceProc Then
3650           Dim vbwProtectorParameterString As String
3651           If vbwProtector.vbwTraceParameters Then
3652               vbwProtectorParameterString = "()"
3653           End If
3654           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3655       End If
' </VB WATCH>

3656       cmbTestSpec.Enabled = False
3657       txtWho.Enabled = False
3658       txtRMA.Enabled = False

3659       frmLoopAndXducer.Enabled = False
3660       frmElecData.Enabled = False
3661       frmPerfMods.Enabled = False
3662       frmOtherFiles.Enabled = False
3663       frmInstrumentTags.Enabled = False
3664       frmTAndI.Enabled = False
3665       frmThrustBalMods.Enabled = False
3666       txtTestSetupRemarks.Enabled = False

3667       cmdEnterTestSetupData.Enabled = False
3668       cmbPLCNo.Enabled = False
' <VB WATCH>
3669       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3670       Exit Sub
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
3671       On Error GoTo vbwErrHandler
3672       Const VBWPROCNAME = "frmPLCData.DisableTestDataControls"
3673       If vbwProtector.vbwTraceProc Then
3674           Dim vbwProtectorParameterString As String
3675           If vbwProtector.vbwTraceParameters Then
3676               vbwProtectorParameterString = "()"
3677           End If
3678           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3679       End If
' </VB WATCH>

3680       cmbPLCLoop.Enabled = False
3681       frmPumpData.Enabled = False
3682       frmThermocouples.Enabled = False
3683       frmAI.Enabled = False
3684       frmMagtrol.Enabled = False
3685       fmrMiscTestData.Enabled = False
3686       frmPLCMisc.Enabled = False
3687       DataGrid1.Enabled = False
3688       DataGrid2.Enabled = False
3689       cmdEnterTestData.Enabled = False

' <VB WATCH>
3690       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3691       Exit Sub
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
3692       On Error GoTo vbwErrHandler
3693       Const VBWPROCNAME = "frmPLCData.EnableTestSetupDataControls"
3694       If vbwProtector.vbwTraceProc Then
3695           Dim vbwProtectorParameterString As String
3696           If vbwProtector.vbwTraceParameters Then
3697               vbwProtectorParameterString = "()"
3698           End If
3699           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3700       End If
' </VB WATCH>

3701       cmbTestSpec.Enabled = True
3702       txtWho.Enabled = True
3703       txtRMA.Enabled = True

3704       frmLoopAndXducer.Enabled = True
3705       frmElecData.Enabled = True
3706       frmPerfMods.Enabled = True
3707       frmOtherFiles.Enabled = True
3708       frmInstrumentTags.Enabled = True
3709       frmTAndI.Enabled = True
3710       frmThrustBalMods.Enabled = True
3711       txtTestSetupRemarks.Enabled = True

3712       cmdEnterTestSetupData.Enabled = True
3713       cmbPLCNo.Enabled = True
' <VB WATCH>
3714       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3715       Exit Sub
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
3716       On Error GoTo vbwErrHandler
3717       Const VBWPROCNAME = "frmPLCData.EnableTestDataControls"
3718       If vbwProtector.vbwTraceProc Then
3719           Dim vbwProtectorParameterString As String
3720           If vbwProtector.vbwTraceParameters Then
3721               vbwProtectorParameterString = "()"
3722           End If
3723           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3724       End If
' </VB WATCH>

3725       cmbPLCLoop.Enabled = True
3726       frmPumpData.Enabled = True
3727       frmThermocouples.Enabled = True
3728       frmAI.Enabled = True
3729       frmMagtrol.Enabled = True
3730       fmrMiscTestData.Enabled = True
3731       frmPLCMisc.Enabled = True
3732       DataGrid1.Enabled = True
3733       DataGrid2.Enabled = True
3734       cmdEnterTestData.Enabled = True

' <VB WATCH>
3735       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3736       Exit Sub
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
3737       On Error GoTo vbwErrHandler
3738       Const VBWPROCNAME = "frmPLCData.EnablePumpDataControls"
3739       If vbwProtector.vbwTraceProc Then
3740           Dim vbwProtectorParameterString As String
3741           If vbwProtector.vbwTraceParameters Then
3742               vbwProtectorParameterString = "()"
3743           End If
3744           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3745       End If
' </VB WATCH>

3746       txtSalesOrderNumber.Enabled = True
3747       frmMfr.Enabled = True
3748       txtShpNo.Enabled = True
3749       txtBilNo.Enabled = True
3750       txtDesignFlow.Enabled = True
3751       txtDesignTDH.Enabled = True

3752       frmMiscPumpData.Enabled = True

3753       txtModelNo.Enabled = True
3754       txtImpellerDia.Enabled = True

3755       frmTEMC.Enabled = True
3756       frmChempump.Enabled = True

3757       txtRemarks.Enabled = True
3758       Me.cmdAddNewTestDate.Visible = True

3759       cmdEnterPumpData.Enabled = True

' <VB WATCH>
3760       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3761       Exit Sub
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
3762       On Error GoTo vbwErrHandler
3763       Const VBWPROCNAME = "frmPLCData.EnableMagtrolFields"
3764       If vbwProtector.vbwTraceProc Then
3765           Dim vbwProtectorParameterString As String
3766           If vbwProtector.vbwTraceParameters Then
3767               vbwProtectorParameterString = "()"
3768           End If
3769           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3770       End If
' </VB WATCH>
3771       txtV1.Enabled = True
3772       txtV2.Enabled = True
3773       txtV3.Enabled = True
3774       txtI1.Enabled = True
3775       txtI2.Enabled = True
3776       txtI3.Enabled = True
3777       txtP1.Enabled = True
3778       txtP2.Enabled = True
3779       txtP3.Enabled = True
3780       optKW(0).Visible = True
3781       optKW(1).Visible = True
3782       optKW(2).Visible = True
3783       optKW(1).value = True
3784       optKW_Click (1)
' <VB WATCH>
3785       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3786       Exit Sub
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
3787       On Error GoTo vbwErrHandler
3788       Const VBWPROCNAME = "frmPLCData.DisableMagtrolFields"
3789       If vbwProtector.vbwTraceProc Then
3790           Dim vbwProtectorParameterString As String
3791           If vbwProtector.vbwTraceParameters Then
3792               vbwProtectorParameterString = "()"
3793           End If
3794           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3795       End If
' </VB WATCH>
3796       txtV1.Enabled = False
3797       txtV2.Enabled = False
3798       txtV3.Enabled = False
3799       txtI1.Enabled = False
3800       txtI2.Enabled = False
3801       txtI3.Enabled = False
3802       txtP1.Enabled = False
3803       txtP2.Enabled = False
3804       txtP3.Enabled = False
3805       txtKW.Enabled = False
3806       optKW(0).Visible = False
3807       optKW(1).Visible = False
3808       optKW(2).Visible = False
' <VB WATCH>
3809       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3810       Exit Sub
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
3811       On Error GoTo vbwErrHandler
3812       Const VBWPROCNAME = "frmPLCData.EnablePLCFields"
3813       If vbwProtector.vbwTraceProc Then
3814           Dim vbwProtectorParameterString As String
3815           If vbwProtector.vbwTraceParameters Then
3816               vbwProtectorParameterString = "()"
3817           End If
3818           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3819       End If
' </VB WATCH>
3820       frmPLCData.txtAI1Display.Enabled = True
3821       frmPLCData.txtAI2Display.Enabled = True
3822       frmPLCData.txtAI3Display.Enabled = True
3823       frmPLCData.txtAI4Display.Enabled = True
3824       frmPLCData.txtTC1Display.Enabled = True
3825       frmPLCData.txtTC2Display.Enabled = True
3826       frmPLCData.txtTC3Display.Enabled = True
3827       frmPLCData.txtTC4Display.Enabled = True
3828       frmPLCData.txtFlowDisplay.Enabled = True
3829       frmPLCData.txtSuctionDisplay.Enabled = True
3830       frmPLCData.txtDischargeDisplay.Enabled = True
3831       frmPLCData.txtTemperatureDisplay.Enabled = True
3832       frmPLCData.txtInHgDisplay.Enabled = True
' <VB WATCH>
3833       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3834       Exit Sub
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
3835       On Error GoTo vbwErrHandler
3836       Const VBWPROCNAME = "frmPLCData.DisablePLCFields"
3837       If vbwProtector.vbwTraceProc Then
3838           Dim vbwProtectorParameterString As String
3839           If vbwProtector.vbwTraceParameters Then
3840               vbwProtectorParameterString = "()"
3841           End If
3842           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3843       End If
' </VB WATCH>
3844       frmPLCData.txtAI1Display.Enabled = False
3845       frmPLCData.txtAI2Display.Enabled = False
3846       frmPLCData.txtAI3Display.Enabled = False
3847       frmPLCData.txtAI4Display.Enabled = False
3848       frmPLCData.txtTC1Display.Enabled = False
3849       frmPLCData.txtTC2Display.Enabled = False
3850       frmPLCData.txtTC3Display.Enabled = False
3851       frmPLCData.txtTC4Display.Enabled = False
3852       frmPLCData.txtFlowDisplay.Enabled = False
3853       frmPLCData.txtSuctionDisplay.Enabled = False
3854       frmPLCData.txtDischargeDisplay.Enabled = False
3855       frmPLCData.txtTemperatureDisplay.Enabled = False
3856       frmPLCData.txtInHgDisplay.Enabled = False
' <VB WATCH>
3857       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3858       Exit Sub
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
3859       On Error GoTo vbwErrHandler
3860       Const VBWPROCNAME = "frmPLCData.BlankData"
3861       If vbwProtector.vbwTraceProc Then
3862           Dim vbwProtectorParameterString As String
3863           If vbwProtector.vbwTraceParameters Then
3864               vbwProtectorParameterString = "()"
3865           End If
3866           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3867       End If
' </VB WATCH>
3868       txtShpNo.Text = vbNullString
3869       txtBilNo.Text = vbNullString
3870       txtModelNo.Text = vbNullString
3871       cmbMotor.ListIndex = -1
3872       cmbStatorFill.ListIndex = -1
3873       cmbVoltage.ListIndex = -1
3874       cmbDesignPressure.ListIndex = -1
3875       cmbFrequency.ListIndex = -1
3876       cmbCirculationPath.ListIndex = -1
3877       cmbRPM.ListIndex = -1
3878       cmbModel.ListIndex = -1
3879       cmbModelGroup.ListIndex = -1
3880       txtSpGr.Text = vbNullString
3881       txtImpellerDia.Text = vbNullString
3882       txtEndPlay.Text = vbNullString
3883       txtGGap.Text = vbNullString
3884       txtDesignFlow.Text = vbNullString
3885       txtDesignTDH.Text = vbNullString
3886       txtOtherMods.Text = vbNullString
3887       txtRemarks.Text = vbNullString
3888       txtSalesOrderNumber.Text = vbNullString
3889       txtTestSetupRemarks.Text = vbNullString
3890       txtNPSHFile.Text = vbNullString
3891       txtPicturesFile.Text = vbNullString
3892       txtVibrationFile.Text = vbNullString
       '    cmbOrificeNumber.ListIndex = 18
       '    cmbTestSpec.ListIndex = 6       'default = Rev7
3893       cmbLoopNumber.ListIndex = -1
3894       cmbSuctDia.ListIndex = -1
3895       cmbDischDia.ListIndex = -1
3896       cmbTachID.ListIndex = -1
3897       cmbAnalyzerNo.ListIndex = -1
3898       txtTestRemarks.Text = vbNullString
3899       txtHDCor.Text = 0
3900       txtDischHeight.Text = 0
3901       txtSuctHeight.Text = 0
3902       txtKWMult.Text = 1
3903       txtWho.Text = LogInInitials
3904       txtRMA.Text = vbNullString
3905       frmPLCData.chkNPSH.value = 0
3906       frmPLCData.chkPictures.value = 0
3907       frmPLCData.chkVibration.value = 0
3908       cmbFlowMeter.ListIndex = -1
3909       cmbSuctionPressureTransducer.ListIndex = -1
3910       cmbDischargePressureTransducer.ListIndex = -1
3911       cmbTemperatureTransducer.ListIndex = -1
3912       cmbCirculationFlowMeter.ListIndex = -1
3913       frmPLCData.chkBalanceHoles.value = 0
3914       frmPLCData.chkCircOrifice.value = 0
3915       frmPLCData.txtCircOrifice = vbNullString
3916       frmPLCData.txtImpTrim = vbNullString
3917       frmPLCData.txtOrifice = vbNullString
3918       frmPLCData.chkFeathered.value = Unchecked
3919       frmPLCData.chkTrimmed.value = 0
3920       frmPLCData.chkCircOrifice.value = 0
3921       frmPLCData.txtThrustBal = vbNullString
3922       frmPLCData.txtRPM = vbNullString
3923       frmPLCData.txtVibAx = vbNullString
3924       frmPLCData.txtVibRad = vbNullString
3925       frmPLCData.txtTEMCTRGReading = vbNullString
3926       dgBalanceHoles.Visible = False
3927       Me.txtLineNumber.Text = vbNullString
3928       Me.txtNPSHr.Text = vbNullString
3929       Me.txtRatedInputPower.Text = vbNullString
3930       Me.txtAmps.Text = vbNullString
3931       Me.txtThermalClass.Text = vbNullString
3932       Me.txtViscosity.Text = vbNullString
3933       Me.txtExpClass.Text = vbNullString
3934       Me.txtNoPhases.Text = vbNullString
3935       Me.txtLiquidTemperature.Text = vbNullString
3936       Me.txtJobNum.Text = vbNullString
3937       Me.txtTEMCFrameNumber.Text = vbNullString
3938       Me.txtLiquid.Text = vbNullString
3939       Me.chkSuperMarketFeathered.value = Unchecked
3940       Me.txtRVSPartNo.Text = vbNullString
' <VB WATCH>
3941       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3942       Exit Sub
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
3943       On Error GoTo vbwErrHandler
3944       Const VBWPROCNAME = "frmPLCData.AddTestData"
3945       If vbwProtector.vbwTraceProc Then
3946           Dim vbwProtectorParameterString As String
3947           If vbwProtector.vbwTraceParameters Then
3948               vbwProtectorParameterString = "()"
3949           End If
3950           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3951       End If
' </VB WATCH>
3952       Dim I As Integer
3953       Dim sFilter As String

3954       ClearEff
3955       rsEff.MoveFirst

3956       For I = 1 To 8
3957           rsTestData.AddNew
3958           rsTestData!SerialNumber = txtSN
3959           rsTestData!Date = cmbTestDate.List(cmbTestDate.ListIndex)
3960           rsTestData!testnumber = I
3961           rsTestData!DataWritten = False
3962           rsTestData.Update
3963           DoEfficiencyCalcs
3964           rsEff.MoveNext
3965           rsTestData.MoveNext
3966       Next I
3967       boFoundTestData = True
           'rsTestData.Update
3968       rsTestData.Requery
3969       rsTestData.Resync

          'select the entries from testdata
3970       sFilter = "SerialNumber='" & txtSN.Text & "' AND Date=#" & cmbTestDate.Text & "#"

3971       rsTestData.Filter = sFilter

3972       Set DataGrid1.DataSource = rsTestData

           ' fix the datagrid

3973       Dim c As Column
3974       For Each c In DataGrid1.Columns
3975          Select Case c.DataField
              Case "TestDataID"
3976             c.Visible = False
3977          Case "SerialNumber"
3978             c.Visible = False
3979          Case "Date"
3980             c.Visible = False
3981          Case Else ' Hide all other columns.
3982             c.Visible = True
3983             c.Alignment = dbgRight
3984          End Select
3985       Next c

3986       rsEff.Requery
3987       DataGrid1.Refresh
3988       DataGrid2.Refresh

' <VB WATCH>
3989       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3990       Exit Sub
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
3991       On Error GoTo vbwErrHandler
3992       Const VBWPROCNAME = "frmPLCData.DoEfficiencyCalcs"
3993       If vbwProtector.vbwTraceProc Then
3994           Dim vbwProtectorParameterString As String
3995           If vbwProtector.vbwTraceParameters Then
3996               vbwProtectorParameterString = "()"
3997           End If
3998           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3999       End If
' </VB WATCH>
4000       Dim KW As Single, VI As Single, VITemp As Single
4001       Dim Vave As Single, Iave As Single
4002       Dim I As Integer
4003       Dim j As Integer
4004       Dim HeightDiff As Single

4005       If Not IsNull(rsTestData.Fields("TotalPower")) Then
4006           KW = rsTestData.Fields("TotalPower")
4007       Else
               'if we wrote data with an old version, we will not have written total power
               'if total power = 0 and the three individual powers are not 0, add them

4008           If rsTestData.Fields("PowerA") > 0 Then
4009               If rsTestData.Fields("PowerB") > 0 Then
4010                   If rsTestData.Fields("PowerC") > 0 Then
4011                       KW = rsTestData.Fields("PowerA") + rsTestData.Fields("PowerB") + rsTestData.Fields("PowerC")
4012                   End If
4013               End If
4014           End If
4015      End If

4016       I = 0
4017       Vave = 0
4018       Iave = 0
4019       If Not IsNull(rsTestData.Fields("VoltageA")) And Not IsNull(rsTestData.Fields("CurrentA")) Then
4020           VI = rsTestData.Fields("VoltageA") * rsTestData.Fields("CurrentA")
4021           Vave = rsTestData.Fields("VoltageA")
4022           Iave = rsTestData.Fields("CurrentA")
4023           If VI <> 0 Then
4024               I = I + 1
4025           End If
4026       End If
4027       If Not IsNull(rsTestData.Fields("VoltageB")) And Not IsNull(rsTestData.Fields("CurrentB")) Then
4028           VITemp = rsTestData.Fields("VoltageB") * rsTestData.Fields("CurrentB")
4029           If VITemp <> 0 Then
4030               I = I + 1
4031               VI = VI + VITemp
4032               Vave = Vave + rsTestData.Fields("VoltageB")
4033               Iave = Iave + rsTestData.Fields("CurrentB")
4034           End If
4035       End If
4036       If Not IsNull(rsTestData.Fields("VoltageC")) And Not IsNull(rsTestData.Fields("CurrentC")) Then
4037           VITemp = rsTestData.Fields("VoltageC") * rsTestData.Fields("CurrentC")
4038           If VITemp <> 0 Then
4039               I = I + 1
4040               VI = VI + VITemp
4041               Vave = Vave + rsTestData.Fields("VoltageC")
4042               Iave = Iave + rsTestData.Fields("CurrentC")
4043           End If
4044       End If
4045       If KW = 0 Then
4046           For j = 1 To rsEff.Fields.Count - 1
4047               rsEff.Fields(j) = 0
4048           Next j
       '        Exit Sub
4049       End If
4050       If VI <> 0 Then
4051           rsEff.Fields("Volts") = Vave / I
4052           rsEff.Fields("Amps") = Iave / I
4053           rsEff.Fields("PowerFactor") = 1000 * I * KW / (VI * Sqr(3))
4054           rsEff.Fields("PowerFactor") = 100 * rsEff.Fields("PowerFactor")
4055       Else
4056           rsEff.Fields("PowerFactor") = 0
4057       End If

4058       If optMfr(0).value = True Then
4059           If cmbStatorFill.ListIndex = -1 Then
4060               rsEff.Fields("MotorEfficiency") = Format$(0, "0.00")

4061           Else
4062               rsEff.Fields("Motorefficiency") = Format$(Round(MotorEfficiency(KW, cmbMotor.ItemData(cmbMotor.ListIndex), cmbStatorFill.ItemData(cmbStatorFill.ListIndex)), 1), "00.0")
       '            rsEff.Fields("Motorefficiency") = Format$(Round(MotorEfficiency(KW, cmbMotor.ListIndex, cmbStatorFill.ListIndex), 1), "00.0")
4063           End If
4064       Else
4065           rsEff.Fields("MotorEfficiency") = Format$(Round(TEMCMotorEfficiency(KW, txtTEMCFrameNumber.Text, 460, RatedKW), 1), "00.0")
4066       End If

4067       Dim sHDCor As Single
4068       Dim sDisc As Single
4069       Dim sSuct As Single
4070       If IsNull(rsTestSetup.Fields("HDCor")) Then
4071           sHDCor = 0
4072       Else
4073           sHDCor = rsTestSetup.Fields("HDCor")
4074       End If
4075       If IsNull(rsTestSetup.Fields("DischargeGageHeight")) Then
4076           sDisc = 0
4077       Else
4078           sDisc = rsTestSetup.Fields("DischargeGageHeight")
4079       End If
4080       If IsNull(rsTestSetup.Fields("SuctionGageHeight")) Then
4081           sSuct = 0
4082       Else
4083           sSuct = rsTestSetup.Fields("SuctionGageHeight")
4084       End If
4085       HeightDiff = sHDCor + sDisc / 12 - sSuct / 12
4086       If (cmbDischDia.ListIndex <> -1 And cmbSuctDia.ListIndex <> -1) Then
4087           rsEff.Fields("VelocityHead") = CalcVelHead(rsTestData.Fields("Flow"), cmbDischDia.ItemData(cmbDischDia.ListIndex) + 1, cmbSuctDia.ItemData(cmbSuctDia.ListIndex) + 1)
4088       End If
       '    rsEff.Fields("VelocityHead") = CalcVelHead(rsTestData.Fields("Flow"), cmbDischDia.ListIndex + 1, cmbSuctDia.ListIndex + 1)
4089       rsEff.Fields("TDH") = CalcTDH(rsTestData.Fields("DischargePressure"), rsTestData.Fields("SuctionPressure"), rsTestData.Fields("SuctionInHg"), rsEff.Fields("VelocityHead"), HeightDiff, rsTestData.Fields("TemperatureSuction"))
4090       rsEff.Fields("ElecHP") = 1000 * KW / 746
       '    If (DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))) <> 0 And KW <> 0) Then
4091           If Int(rsTestData.Fields("TemperatureSuction")) >= 40 Then
4092               If (DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))) <> 0 And KW <> 0) Then
           '        rsEff.Fields("LiquidHP") = (rsEff.Fields("TDH") * rsTestData.Fields("Flow") * DLookup("TDHCorr", "TempCorrection", "Temp = 68")) / (3960 * DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))))
4093               rsEff.Fields("LiquidHP") = (rsEff.Fields("TDH") * rsTestData.Fields("Flow") * DLookupA(TDHColNo, TempCorrection, TempColNo, 68)) / (3960 * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))))
           '        rsEff.Fields("OverallEfficiency") = (0.189 * rsTestData.Fields("Flow") * rsEff.Fields("TDH") * DLookup("TDHCorr", "TempCorrection", "Temp = 68")) / (10 * KW * DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))))
4094               rsEff.Fields("OverallEfficiency") = (0.189 * rsTestData.Fields("Flow") * rsEff.Fields("TDH") * DLookupA(TDHColNo, TempCorrection, TempColNo, 68)) / (10 * KW * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))))
4095               If rsEff.Fields("MotorEfficiency") <> 0 Then
4096                   rsEff.Fields("HydraulicEfficiency") = 100 * rsEff.Fields("OverallEfficiency") / rsEff.Fields("MotorEfficiency")
4097               Else
4098                   rsEff.Fields("HydraulicEfficiency") = 0
4099               End If
4100           Else
4101               rsEff.Fields("LiquidHP") = 0
4102               rsEff.Fields("OverallEfficiency") = 0
4103           End If

4104       Else
4105           rsEff.Fields("LiquidHP") = 0
4106           rsEff.Fields("OverallEfficiency") = 0
4107       End If


4108       I = rsEff.AbsolutePosition
4109       If Not IsNull(rsTestData.Fields("Flow")) Then
4110           rsEff.Fields("Flow") = rsTestData.Fields("Flow")
4111           HeadFlow(0, I - 1) = rsTestData.Fields("Flow")
4112           HeadFlow(1, I - 1) = rsEff.Fields("TDH")
4113           FlowHead(I - 1, 0) = rsTestData.Fields("Flow")
4114           FlowHead(I - 1, 1) = rsEff.Fields("TDH")

       '        EffFlow(0, i - 1) = rsTestData.Fields("Flow")
       '        EffFlow(1, i - 1) = rsEff.Fields("OverallEfficiency")
       '        KWFlow(0, i - 1) = rsTestData.Fields("Flow")
       '        KWFlow(1, i - 1) = KW
       '        AmpsFlow(0, i - 1) = rsTestData.Fields("Flow")
       '        AmpsFlow(1, i - 1) = rsEff.Fields("Amps")
4115       Else
4116           HeadFlow(0, I - 1) = 0
4117           HeadFlow(1, I - 1) = 0
4118           FlowHead(I - 1, 0) = 0
4119           FlowHead(I - 1, 1) = 0

       '        EffFlow(0, i - 1) = 0
       '        EffFlow(1, i - 1) = 0
       '        KWFlow(0, i - 1) = 0
       '        KWFlow(1, i - 1) = 0
       '        AmpsFlow(0, i - 1) = 0
       '        AmpsFlow(1, i - 1) = 0
4120       End If

4121       Dim Plothead(1, 7) As Single
4122       Dim HeadPlot(7, 1) As Single
           'ReDim Preserve Plothead(1, j)
           'ReDim Preserve HeadPlot(j, 1)

       '    Dim PlotEff() As Single
       '    Dim PlotKW() As Single
       '    Dim PlotAmps() As Single
       '    ReDim PlotHead(0, 0)
       '    ReDim PlotEff(0, 0)
       '    ReDim PlotKW(0, 0)
       '
4123       For j = 0 To UpDown2.value - 1
       '        If HeadFlow(1, j) <> 0 Then
       '            ReDim Preserve Plothead(1, j)
       '            ReDim Preserve HeadPlot(j, 1)
4124               Plothead(0, j) = HeadFlow(0, j)
4125               Plothead(1, j) = HeadFlow(1, j)
4126               HeadPlot(j, 0) = FlowHead(j, 0)
4127               HeadPlot(j, 1) = FlowHead(j, 1)
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
4128       Next j




       '    SetGraphMax (Plothead())
       '    If UBound(PlotHead()) <> 0 Then

       'fix 4/29/19

4129           MSChart1.ChartData = HeadPlot

       '    End If

           'copy fields for reports
4130       rsEff.Fields("DischPress") = rsTestData.Fields("Dischargepressure")
4131       rsEff.Fields("SuctPress") = rsTestData.Fields("Suctionpressure")
       '    rsEff.Fields("Volts") = rsTestData.Fields("VoltageA")
       '    rsEff.Fields("Amps") = rsTestData.Fields("CurrentA")
4132       rsEff.Fields("KW") = KW
4133       rsEff.Fields("Freq") = rsTestData.Fields("VFDFrequency")
4134       rsEff.Fields("RPM") = rsTestData.Fields("RPM")
4135       rsEff.Fields("Pos") = rsTestData.Fields("ThrustBalance")
4136       rsEff.Fields("NPSHa") = rsTestData.Fields("NPSHa")
4137       rsEff.Fields("NPSHr") = rsTestData.Fields("NPSHr")
4138       rsEff.Fields("InputPower") = rsTestData.Fields("TotalPower")
4139       rsEff.Fields("Temperature") = rsTestData.Fields("TemperatureSuction")
4140       rsEff.Fields("CircFlow") = rsTestData.Fields("CircFlow")
4141       rsEff.Fields("VibrationX") = rsTestData.Fields("VibrationX")
4142       rsEff.Fields("VibrationY") = rsTestData.Fields("VibrationY")
4143       rsEff.Fields("CurrentA") = rsTestData.Fields("CurrentA")
4144       rsEff.Fields("CurrentB") = rsTestData.Fields("CurrentB")
4145       rsEff.Fields("CurrentC") = rsTestData.Fields("CurrentC")
4146       rsEff.Fields("VoltageA") = rsTestData.Fields("VoltageA")
4147       rsEff.Fields("VoltageB") = rsTestData.Fields("VoltageB")
4148       rsEff.Fields("VoltageC") = rsTestData.Fields("VoltageC")
4149       rsEff.Fields("TC1") = rsTestData.Fields("TC1")
4150       rsEff.Fields("TC2") = rsTestData.Fields("TC2")
4151       rsEff.Fields("TC3") = rsTestData.Fields("TC3")
4152       rsEff.Fields("TC4") = rsTestData.Fields("TC4")
4153       rsEff.Fields("RBHTemp") = rsTestData.Fields("RBHTemp")
4154       rsEff.Fields("RBHPress") = rsTestData.Fields("RBHPress")
4155       rsEff.Fields("AI4") = rsTestData.Fields("AI4")
4156       rsEff.Fields("Remarks") = rsTestData.Fields("Remarks")
4157       rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCFrontThrust")
4158       rsEff.Fields("TEMCRearThrust") = rsTestData.Fields("TEMCRearThrust")
4159       rsEff.Fields("TEMCTRG") = rsTestData.Fields("TEMCTRG")
4160       rsEff.Fields("TEMCThrustRigPressure") = rsTestData.Fields("TEMCThrustRigPressure")
4161       rsEff.Fields("TEMCMomentArm") = rsTestData.Fields("TEMCMomentArm")
4162       rsEff.Fields("TEMCViscosity") = rsTestData.Fields("TEMCViscosity")
4163       If Not IsNull(rsEff.Fields("TEMCFrontThrust")) Then
4164           txtTEMCFrontThrust.Text = rsEff.Fields("TEMCFrontThrust")
4165       End If
4166       If Not IsNull(rsEff.Fields("TEMCREarThrust")) Then
4167           txtTEMCRearThrust.Text = rsEff.Fields("TEMCREarThrust")
4168       End If
4169       If (Not IsNull(rsEff.Fields("TEMCViscosity"))) And (rsEff.Fields("TEMCViscosity") <> 0) Then
4170           txtTEMCViscosity.Text = rsEff.Fields("TEMCViscosity")
4171       End If
4172       If Not IsNull(rsTestData.Fields("TEMCThrustRigPressure")) Then
4173           txtTEMCThrustRigPressure.Text = rsTestData.Fields("TEMCThrustRigPressure")
4174       End If
4175       If Not IsNull(rsTestData.Fields("TEMCMomentArm")) Then
4176           txtTEMCMomentArm.Text = rsTestData.Fields("TEMCMomentArm")
4177       End If

        '   If Not IsNull(Me.txtAI3Display.Text) Then
        '       Me.txtAI3Display = rsTestData.Fields("RBHPress")
        '   End If

4178       CalculateTEMCForce

4179       If Not IsNull(txtTEMCCalcForce.Text) Then
4180           rsEff.Fields("TEMCCalculatedForce") = Val(txtTEMCCalcForce.Text)
4181       Else
4182           rsEff.Fields("TEMCCalculatedForce") = 0
4183       End If

4184       If Not IsNull(txtTEMCPVValue.Text) Then
4185           rsEff.Fields("TEMCPV") = Val(txtTEMCPVValue.Text)
4186       Else
4187           rsEff.Fields("TEMCPV") = 0
4188       End If

4189       If Val(txtTEMCFrontThrust.Text) <> 0 Then
4190           rsEff.Fields("TEMCFR") = "F"
       '        rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCFrontThrust")
4191       Else
4192           If Val(txtTEMCRearThrust.Text) = 0 Then
                   'no thrust
4193               rsEff.Fields("TEMCFR") = " "
4194               rsEff.Fields("TEMCFrontThrust") = 0
4195           Else
4196               rsEff.Fields("TEMCFR") = "R"
       '            rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCRearThrust")
4197           End If
4198       End If

4199       rsEff.Fields("TEMCForceDirection") = Left(lblTEMCFrontRear.Caption, 1)

4200       rsEff.Update
4201       DataGrid2.Refresh


' <VB WATCH>
4202       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4203       Exit Sub
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
4204       On Error GoTo vbwErrHandler
4205       Const VBWPROCNAME = "frmPLCData.ClearEff"
4206       If vbwProtector.vbwTraceProc Then
4207           Dim vbwProtectorParameterString As String
4208           If vbwProtector.vbwTraceParameters Then
4209               vbwProtectorParameterString = "()"
4210           End If
4211           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4212       End If
' </VB WATCH>
4213       Dim qy As New ADODB.Command

4214       If rsEff.State = adStateOpen Then
4215           If Not (rsEff.BOF = True Or rsEff.EOF = True) Then
4216               rsEff.CancelUpdate
4217           End If
4218           rsEff.Close
4219       End If
4220       qy.ActiveConnection = cnEffData
4221       qy.CommandText = "DROP TABLE Efficiency"
4222       rsEff.Open qy
4223       qy.CommandText = "SELECT EfficiencyOrg.* INTO Efficiency FROM EfficiencyOrg;"
4224       rsEff.Open qy
4225       rsEff.Open "Efficiency", cnEffData, adOpenStatic, adLockOptimistic, adCmdTableDirect

4226       rsEff.Requery
4227       DataGrid2.Refresh

4228       Dim c As Column
4229       For Each c In DataGrid2.Columns
4230           c.Alignment = dbgCenter
4231           c.Width = 750
4232           Select Case c.ColIndex
                   Case 1
4233                   c.Caption = "Flow"
4234                   c.NumberFormat = "###0.00"
4235               Case 2
4236                   c.Caption = "TDH"
4237                   c.NumberFormat = "00.0"
4238               Case 3
4239                   c.Caption = "Overall Eff"
4240                   c.NumberFormat = "00.00"
4241                   c.Width = 850
4242               Case 4
4243                   c.Caption = "PF"
4244                   c.NumberFormat = "00.0"
4245               Case 5
4246                   c.Caption = "Vel Head"
4247                   c.NumberFormat = "00.00"
4248               Case 6
4249                   c.Caption = "Elec HP"
4250                   c.NumberFormat = "#00.0"
4251               Case 7
4252                   c.Caption = "Liq HP"
4253                   c.NumberFormat = "#00.0"
4254               Case Else
4255                   c.Visible = False
4256           End Select
4257       Next c

' <VB WATCH>
4258       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4259       Exit Sub
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
4260       On Error GoTo vbwErrHandler
4261       Const VBWPROCNAME = "frmPLCData.JustAlphaNumeric"
4262       If vbwProtector.vbwTraceProc Then
4263           Dim vbwProtectorParameterString As String
4264           If vbwProtector.vbwTraceParameters Then
4265               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("char", char) & ") "
4266           End If
4267           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4268       End If
' </VB WATCH>
4269       Select Case Asc(char)
               Case 42             ' *
4270               JustAlphaNumeric = char
4271           Case 48 To 57       ' 0 - 9
4272               JustAlphaNumeric = char
4273           Case 65 To 90       ' A - Z
4274               JustAlphaNumeric = char
4275           Case 97 To 122      ' a - z
4276               JustAlphaNumeric = UCase(char)
4277           Case Else
4278               JustAlphaNumeric = ""
4279       End Select
' <VB WATCH>
4280       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4281       Exit Function
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
4282       On Error GoTo vbwErrHandler
4283       Const VBWPROCNAME = "frmPLCData.txtI1_Change"
4284       If vbwProtector.vbwTraceProc Then
4285           Dim vbwProtectorParameterString As String
4286           If vbwProtector.vbwTraceParameters Then
4287               vbwProtectorParameterString = "()"
4288           End If
4289           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4290       End If
' </VB WATCH>
4291       txtI2.Text = txtI1.Text
4292       txtI3.Text = txtI1.Text
' <VB WATCH>
4293       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4294       Exit Sub
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
4295       On Error GoTo vbwErrHandler
4296       Const VBWPROCNAME = "frmPLCData.txtModelNo_Change"
4297       If vbwProtector.vbwTraceProc Then
4298           Dim vbwProtectorParameterString As String
4299           If vbwProtector.vbwTraceParameters Then
4300               vbwProtectorParameterString = "()"
4301           End If
4302           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4303       End If
' </VB WATCH>
4304       Dim I As Integer
4305       Dim S As String
4306       Dim sFull As String
4307       Dim boDone As Boolean
4308       Dim boRepeat As Boolean

4309       Static bo3Digits As Boolean         '3 digits in frame number
4310       Static bo2Digits As Boolean         '2 digits in stages

4311       If optMfr(0).value = True Then
' <VB WATCH>
4312       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4313           Exit Sub
4314       End If

4315       cmbTEMCAdapter.ListIndex = -1
4316       cmbTEMCAdditions.ListIndex = -1
4317       cmbTEMCCirculation.ListIndex = -1
4318       cmbTEMCDesignPressure.ListIndex = -1
4319       cmbTEMCNominalDischargeSize.ListIndex = -1
4320       cmbTEMCDivisionType.ListIndex = -1
4321       cmbTEMCImpellerType.ListIndex = -1
4322       cmbTEMCInsulation.ListIndex = -1
4323       cmbTEMCJacketGasket.ListIndex = -1
4324       cmbTEMCMaterials.ListIndex = -1
4325       cmbTEMCModel.ListIndex = -1
4326       cmbTEMCNominalImpSize.ListIndex = -1
4327       cmbTEMCOtherMotor.ListIndex = -1
4328       cmbTEMCPumpStages.ListIndex = -1
4329       cmbTEMCNominalSuctionSize.ListIndex = -1
4330       cmbTEMCTRG.ListIndex = -1
4331       cmbTEMCVoltage.ListIndex = -1


           'first, get rid of spaces, dashes, etc

4332       S = ""
4333       For I = 1 To Len(txtModelNo.Text)
4334           S = S & JustAlphaNumeric(Mid$(txtModelNo.Text, I, 1))
4335       Next I

           'next, fill out the model number to it's max length of 24 characters

4336       boDone = False
4337       boRepeat = False

4338       Do While Not boDone
4339           sFull = ""
4340           For I = 1 To Len(S)
4341               Select Case I
                       Case 1
                           'type
4342                       sFull = sFull & Mid$(S, I, 1)
4343                   Case 2
                           'adapter
4344                       If IsNumeric(Mid$(S, I, 1)) Then
4345                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4346                           boRepeat = True
4347                           Exit For
4348                       Else
4349                           sFull = sFull & Mid$(S, I, 1)
4350                           boRepeat = False
4351                       End If
4352                   Case 3
                           'materials
4353                       sFull = sFull & Mid$(S, I, 1)
4354                   Case 4
                       'design pressure
4355                       sFull = sFull & Mid$(S, I, 1)
4356                   Case 5
                       'motor frame number - digit 1
4357                       sFull = sFull & Mid$(S, I, 1)
4358                   Case 6
                       'motor frame number - digit 2
4359                       sFull = sFull & Mid$(S, I, 1)
4360                   Case 7
                       'motor frame number - digit 3
4361                       sFull = sFull & Mid$(S, I, 1)
4362                   Case 8
                       'motor frame number - digit 4
4363                       If IsNumeric(Mid$(S, I, 1)) Then
4364                           sFull = sFull & Mid$(S, I, 1)
4365                           boRepeat = False
4366                       Else    '3 digits
       '                        s = Left$(s, i - 1) & "*" & Right$(s, Len(s) - i + 1)
4367                           S = Left$(S, I - 4) & "0" & Right$(S, Len(S) - I + 4)
4368                           boRepeat = True
4369                           Exit For
4370                       End If
4371                   Case 9
                       'insulation
4372                       sFull = sFull & Mid$(S, I, 1)
4373                   Case 10
                       'voltage
4374                       sFull = sFull & Mid$(S, I, 1)
4375                   Case 11
                       'other motor specs
4376                       If Mid$(S, I, 1) = "M" Or Mid$(S, I, 1) = "R" Or Mid$(S, I, 1) = "L" Or Mid$(S, I, 1) = "G" Or Mid$(S, I, 1) = "N" Then
4377                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4378                           boRepeat = True
4379                           Exit For
4380                       Else
4381                           sFull = sFull & Mid$(S, I, 1)
4382                           boRepeat = False
4383                       End If
4384                   Case 12
                       ' TRG
4385                       sFull = sFull & Mid$(S, I, 1)
4386                   Case 13
                       'Nominal discharge - digit 1
4387                       sFull = sFull & Mid$(S, I, 1)
4388                   Case 14
                       'nominal discharge - digit 2
4389                       sFull = sFull & Mid$(S, I, 1)
4390                   Case 15
                       'nominal suction - digit 1
4391                       sFull = sFull & Mid$(S, I, 1)
4392                   Case 16
                       'nominal suction - digit 2
4393                       sFull = sFull & Mid$(S, I, 1)
4394                   Case 17
                       'nominal impeller size
4395                       sFull = sFull & Mid$(S, I, 1)
4396                   Case 18
                       'impeller type
4397                       If Mid$(S, I, 1) <> "*" Then
4398                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4399                           boRepeat = True
4400                           Exit For
4401                       Else
4402                           sFull = sFull & Mid$(S, I, 1)
4403                           boRepeat = False
4404                       End If
4405                   Case 19
                       'Division type
4406                       If IsNumeric(Mid$(S, I, 1)) Then
4407                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4408                           boRepeat = True
4409                           Exit For
4410                       Else
4411                           sFull = sFull & Mid$(S, I, 1)
4412                           boRepeat = False
4413                       End If
4414                   Case 20
                       'pump stages - digit 1
4415                       sFull = sFull & Mid$(S, I, 1)
4416                   Case 21
                       'pump jacket
4417                       If Mid$(S, I, 1) = "A" Or Mid$(S, I, 1) = "B" Or Mid$(S, I, 1) = "E" Or Mid$(S, I, 1) = "F" Or _
                                             Mid$(S, I, 1) = "G" Or Mid$(S, I, 1) = "H" Or Mid$(S, I, 1) = "J" Or Mid$(S, I, 1) = "K" Then
4418                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4419                           boRepeat = True
4420                       Else
4421                           sFull = sFull & Mid$(S, I, 1)
4422                           boRepeat = False
4423                       End If
4424                   Case 22
                       'additions
4425                         sFull = sFull & Mid$(S, I, 1)
4426                   Case 23
                       'circulation
4427                         sFull = sFull & Mid$(S, I, 1)
4428               End Select
4429           Next I
4430           If Not boRepeat Then
4431               boDone = True
4432           End If
4433       Loop

4434       For I = 1 To Len(sFull)
4435           Select Case I
                   Case 1
4436                   ParseTEMCModelNo cmbTEMCModel, Mid$(sFull, I, 1)
4437               Case 2
4438                   ParseTEMCModelNo cmbTEMCAdapter, Mid$(sFull, I, 1)
4439               Case 3
4440                   ParseTEMCModelNo cmbTEMCMaterials, Mid$(sFull, I, 1)
4441               Case 4
4442                   ParseTEMCModelNo cmbTEMCDesignPressure, Mid$(sFull, I, 1)
4443               Case 5
4444                       If Val(Mid$(sFull, I, 1)) = 0 Then
4445                           txtTEMCFrameNumber.Text = Mid$(sFull, 6, 3)
4446                       Else
4447                           txtTEMCFrameNumber.Text = Mid$(sFull, 5, 4)
4448                       End If
4449               Case 9
4450                       ParseTEMCModelNo cmbTEMCInsulation, Mid$(sFull, I, 1)
4451               Case 10
4452                       ParseTEMCModelNo cmbTEMCVoltage, Mid$(sFull, I, 1)
4453               Case 11
4454                       ParseTEMCModelNo cmbTEMCOtherMotor, Mid$(sFull, I, 1)
4455               Case 12
4456                       ParseTEMCModelNo cmbTEMCTRG, Mid$(sFull, I, 1)
4457               Case 13
4458                       ParseTEMCModelNo cmbTEMCNominalDischargeSize, Mid$(sFull, I, 2)
4459               Case 14
4460               Case 15
4461                       ParseTEMCModelNo cmbTEMCNominalSuctionSize, Mid$(sFull, I, 2)
4462               Case 16
4463               Case 17
4464                       ParseTEMCModelNo cmbTEMCNominalImpSize, Mid$(sFull, I, 1)
4465               Case 18
4466                       ParseTEMCModelNo cmbTEMCImpellerType, Mid$(sFull, I, 1)
4467               Case 19
4468                       ParseTEMCModelNo cmbTEMCDivisionType, Mid$(sFull, I, 1)
4469               Case 20
4470                       ParseTEMCModelNo cmbTEMCPumpStages, Mid$(sFull, I, 1)
4471               Case 21
4472                       ParseTEMCModelNo cmbTEMCJacketGasket, Mid$(sFull, I, 1)
4473               Case 22
4474                       ParseTEMCModelNo cmbTEMCAdditions, Mid$(sFull, I, 1)
4475                       ParseTEMCModelNo cmbTEMCCirculation, "*"
4476               Case 23
       '                    ParseTEMCModelNo cmbTEMCCirculation, Mid$(sFull, I, 1)

4477           End Select
4478       Next I

           'give alerts on certain conditions
4479       Dim msg As String
4480       msg = ""
4481       If Left(cmbTEMCVoltage, 3) = "[6]" Then
4482           msg = "575V transformer required for Rundown and TRG"
4483       End If
       '    If Left(cmbTEMCTRG, 3) = "[L]" Or InStr("X[B][F][H][K]", Left(cmbTEMCAdditions, 3)) > 1 Then
4484       If Left(cmbTEMCTRG, 3) = "[L]" Then
4485           If msg = "" Then
4486               msg = "VFD required for Rundown and TRG"
4487           Else
4488               msg = msg & " and " & "VFD required for Rundown and TRG"
4489           End If
4490       End If

4491       If InStr("X[B][F][H][K]", Left(cmbTEMCAdditions, 3)) > 1 Then
4492           If msg = "" Then
4493               msg = "VFD required for Rundown, standard drive required for TRG"
4494           Else
4495               msg = msg & " and " & "VFD required for Rundown, standard drive required for TRG"
4496           End If
4497       End If

4498       If msg <> "" Then
4499           frmAlert.txtAlert.Text = msg
4500           frmAlert.Show
4501       End If

' <VB WATCH>
4502       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4503       Exit Sub
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
4504       On Error GoTo vbwErrHandler
4505       Const VBWPROCNAME = "frmPLCData.txtModelNo_Validate"
4506       If vbwProtector.vbwTraceProc Then
4507           Dim vbwProtectorParameterString As String
4508           If vbwProtector.vbwTraceParameters Then
4509               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Cancel", Cancel) & ") "
4510           End If
4511           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4512       End If
' </VB WATCH>
4513       Dim I As Integer
4514       Dim S As String

       '    s = txtModelNo.Text
       '    S = Replace(S, "-", "")
       '    S = Replace(S, " ", "")
       '    S = Replace(S, "/", "")

       '    txtModelNo.Text = ""

       '    For i = 1 To Len(s)
       '        txtModelNo.Text = txtModelNo.Text & Mid(s, i, 1)
       '    Next i
4515       txtModelNo_Change

' <VB WATCH>
4516       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4517       Exit Sub
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
4518       On Error GoTo vbwErrHandler
4519       Const VBWPROCNAME = "frmPLCData.txtNPSHFile_GotFocus"
4520       If vbwProtector.vbwTraceProc Then
4521           Dim vbwProtectorParameterString As String
4522           If vbwProtector.vbwTraceParameters Then
4523               vbwProtectorParameterString = "()"
4524           End If
4525           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4526       End If
' </VB WATCH>
4527       On Error GoTo FileCancel
4528       If LenB(txtNPSHFile.Text) <> 0 Then
4529           CommonDialog1.filename = txtNPSHFile.Text
4530       End If
4531       CommonDialog1.ShowOpen
4532       txtNPSHFile.Text = CommonDialog1.filename
' <VB WATCH>
4533       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4534       Exit Sub
4535   FileCancel:
4536   On Error GoTo vbwErrHandler
4537       CommonDialog1.CancelError = False
' <VB WATCH>
4538       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4539       Exit Sub
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
4540       On Error GoTo vbwErrHandler
4541       Const VBWPROCNAME = "frmPLCData.txtP1_Change"
4542       If vbwProtector.vbwTraceProc Then
4543           Dim vbwProtectorParameterString As String
4544           If vbwProtector.vbwTraceParameters Then
4545               vbwProtectorParameterString = "()"
4546           End If
4547           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4548       End If
' </VB WATCH>
4549       txtP2.Text = txtP1.Text
4550       txtP3.Text = txtP1.Text
' <VB WATCH>
4551       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4552       Exit Sub
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
4553       On Error GoTo vbwErrHandler
4554       Const VBWPROCNAME = "frmPLCData.txtPicturesFile_gotfocus"
4555       If vbwProtector.vbwTraceProc Then
4556           Dim vbwProtectorParameterString As String
4557           If vbwProtector.vbwTraceParameters Then
4558               vbwProtectorParameterString = "()"
4559           End If
4560           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4561       End If
' </VB WATCH>
4562       CommonDialog1.CancelError = True
4563       On Error GoTo FileCancel
4564       If LenB(txtPicturesFile.Text) <> 0 Then
4565           CommonDialog1.filename = txtPicturesFile.Text
4566       End If
4567       CommonDialog1.ShowOpen
4568       txtPicturesFile.Text = CommonDialog1.filename
' <VB WATCH>
4569       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4570       Exit Sub
4571   FileCancel:
4572   On Error GoTo vbwErrHandler
4573       CommonDialog1.CancelError = False
' <VB WATCH>
4574       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4575       Exit Sub
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
4576       On Error GoTo vbwErrHandler
4577       Const VBWPROCNAME = "frmPLCData.txtSN_Change"
4578       If vbwProtector.vbwTraceProc Then
4579           Dim vbwProtectorParameterString As String
4580           If vbwProtector.vbwTraceParameters Then
4581               vbwProtectorParameterString = "()"
4582           End If
4583           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4584       End If
' </VB WATCH>
4585       cmdFindPump.Default = True
' <VB WATCH>
4586       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4587       Exit Sub
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
4588       On Error GoTo vbwErrHandler
4589       Const VBWPROCNAME = "frmPLCData.txtTEMCFrontThrust_Change"
4590       If vbwProtector.vbwTraceProc Then
4591           Dim vbwProtectorParameterString As String
4592           If vbwProtector.vbwTraceParameters Then
4593               vbwProtectorParameterString = "()"
4594           End If
4595           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4596       End If
' </VB WATCH>
4597       CalculateTEMCForce
' <VB WATCH>
4598       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4599       Exit Sub
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
4600       On Error GoTo vbwErrHandler
4601       Const VBWPROCNAME = "frmPLCData.txtTEMCMomentArm_Change"
4602       If vbwProtector.vbwTraceProc Then
4603           Dim vbwProtectorParameterString As String
4604           If vbwProtector.vbwTraceParameters Then
4605               vbwProtectorParameterString = "()"
4606           End If
4607           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4608       End If
' </VB WATCH>
4609       CalculateTEMCForce
' <VB WATCH>
4610       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4611       Exit Sub
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
4612       On Error GoTo vbwErrHandler
4613       Const VBWPROCNAME = "frmPLCData.txtTEMCRearThrust_Change"
4614       If vbwProtector.vbwTraceProc Then
4615           Dim vbwProtectorParameterString As String
4616           If vbwProtector.vbwTraceParameters Then
4617               vbwProtectorParameterString = "()"
4618           End If
4619           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4620       End If
' </VB WATCH>
4621       CalculateTEMCForce
' <VB WATCH>
4622       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4623       Exit Sub
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
4624       On Error GoTo vbwErrHandler
4625       Const VBWPROCNAME = "frmPLCData.txtTEMCThrustRigPressure_Change"
4626       If vbwProtector.vbwTraceProc Then
4627           Dim vbwProtectorParameterString As String
4628           If vbwProtector.vbwTraceParameters Then
4629               vbwProtectorParameterString = "()"
4630           End If
4631           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4632       End If
' </VB WATCH>
4633       CalculateTEMCForce
' <VB WATCH>
4634       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4635       Exit Sub
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
' <VB WATCH>
4636       On Error GoTo vbwErrHandler
4637       Const VBWPROCNAME = "frmPLCData.txtTEMCViscosity_Change"
4638       If vbwProtector.vbwTraceProc Then
4639           Dim vbwProtectorParameterString As String
4640           If vbwProtector.vbwTraceParameters Then
4641               vbwProtectorParameterString = "()"
4642           End If
4643           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4644       End If
' </VB WATCH>
4645       CalculateTEMCForce
' <VB WATCH>
4646       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4647       Exit Sub
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
4648       On Error GoTo vbwErrHandler
4649       Const VBWPROCNAME = "frmPLCData.txtV1_Change"
4650       If vbwProtector.vbwTraceProc Then
4651           Dim vbwProtectorParameterString As String
4652           If vbwProtector.vbwTraceParameters Then
4653               vbwProtectorParameterString = "()"
4654           End If
4655           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4656       End If
' </VB WATCH>
4657       txtV2.Text = txtV1.Text
4658       txtV3.Text = txtV1.Text
' <VB WATCH>
4659       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4660       Exit Sub
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
4661       On Error GoTo vbwErrHandler
4662       Const VBWPROCNAME = "frmPLCData.txtVibrationFile_gotfocus"
4663       If vbwProtector.vbwTraceProc Then
4664           Dim vbwProtectorParameterString As String
4665           If vbwProtector.vbwTraceParameters Then
4666               vbwProtectorParameterString = "()"
4667           End If
4668           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4669       End If
' </VB WATCH>
4670       On Error GoTo FileCancel
4671       If LenB(txtVibrationFile.Text) <> 0 Then
4672           CommonDialog1.filename = txtVibrationFile.Text
4673       End If
4674       CommonDialog1.ShowOpen
4675       txtVibrationFile.Text = CommonDialog1.filename
' <VB WATCH>
4676       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4677       Exit Sub
4678   FileCancel:
4679   On Error GoTo vbwErrHandler
4680       CommonDialog1.CancelError = False
' <VB WATCH>
4681       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4682       Exit Sub
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
4683       On Error GoTo vbwErrHandler
4684       Const VBWPROCNAME = "frmPLCData.ExportToExcel"
4685       If vbwProtector.vbwTraceProc Then
4686           Dim vbwProtectorParameterString As String
4687           If vbwProtector.vbwTraceParameters Then
4688               vbwProtectorParameterString = "()"
4689           End If
4690           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4691       End If
' </VB WATCH>

4692       Dim SaveFileName As String
4693       Dim WorkSheetName As String

4694       Dim I As Integer
4695       Dim iRowNo As Integer
4696       Dim sImp As String
4697       Dim ans As Integer

4698       Dim bCanShowSpeed As Boolean
4699       Dim CantShowReason As String

       'close any running excel processes
4700       Dim objWMIService, colProcesses
4701       Set objWMIService = GetObject("winmgmts:")
4702       Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name LIKE 'Excel%'")
4703       If colProcesses.Count > 0 Then
4704           Set xlApp = Excel.Application
4705       Else
               'use existing copy
       '        Set xlApp = New Excel.Application
4706           Set xlApp = CreateObject("Excel.Application")
4707       End If


4708       CommonDialog1.CancelError = True        'in case the user
4709       On Error GoTo ErrHandler                '  chooses the cancel button

           'set up dialog box
4710       CommonDialog1.DialogTitle = "Open Excel Files"
4711       CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|"  'show Excel files
4712       CommonDialog1.InitDir = App.Path
       '    CommonDialog1.InitDir = "C:\"    'in this directory
4713       CommonDialog1.ShowOpen                              'open the file selection dialog box

4714       If Dir(CommonDialog1.filename) = "" Then            'if the file name does not exist yet
4715           SaveFileName = CommonDialog1.filename           'get the name of the file
4716           If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
4717                xlApp.Workbooks.Close
4718           End If
               ' Create the Excel Workbook Object.
4719   On Error GoTo vbwErrHandler
4720           Set xlBook = xlApp.Workbooks.Add                'add a workbook
4721           WorkSheetName = NewWorkBook                                     'do some stuff for the new workbook
4722           ActiveWorkbook.CheckCompatibility = False
4723           xlApp.ActiveWorkbook.SaveAs filename:=SaveFileName, _
                                 FileFormat:=xlNormal                        'save the file
4724       Else                                                'the file name already exists
4725           SaveFileName = CommonDialog1.filename
               ' Create the Excel Workbook Object.
4726           If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
4727                xlApp.Workbooks.Close
4728           End If
4729           Set xlBook = xlApp.Workbooks.Open(SaveFileName)             'get the file name selected
4730           If GetWorksheetTabs(SaveFileName, WorkSheetName) = vbNo Then    'ask the user if he/she wants a new tab.
4731               MsgBox "File not overwritten.", vbOKOnly, "File not Opened"
' <VB WATCH>
4732       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4733               Exit Sub
4734           Else
4735           End If
4736       End If

4737   On Error GoTo vbwErrHandler

           'see if we can export Speed and SG and if we can, ask user if s/he wants it
           'assume that we can show speed calcs

4738       bCanShowSpeed = False
       'open the template and copy the data from the sheet
       '  excel file resides in ParentDirectoryName + "\Polar SG&Visc Correction5.xls"
           'write the data to the spreadsheet
4739       With xlApp

4740       Dim xlTemplateName As String
4741       xlTemplateName = ParentDirectoryName & sSGandViscSpreadsheetTemplate
4742       Dim xlTemplate As Excel.Workbook
4743       Set xlTemplate = xlApp.Workbooks.Open(xlTemplateName)
4744       Dim TemplateWS As Excel.Worksheet
4745       Dim sheetName As String
4746       sheetName = xlTemplate.Sheets(1).Name
4747       xlTemplate.Sheets(1).Copy After:=xlBook.Sheets(WorkSheetName)

4748       xlTemplate.Close savechanges:=False

4749       Set xlTemplate = Nothing

4750       Application.DisplayAlerts = False
4751       ActiveWorkbook.Worksheets(WorkSheetName).Delete
4752       Application.DisplayAlerts = True
4753       ActiveWorkbook.Worksheets(sheetName).Name = WorkSheetName

           'WorkSheetName = sheetName

           'first see if there is an entry in CalculatedRPM table for this frame size and voltage.
           ' if there is, get the coefficients, else make the coefficients 0

4754           Dim ACoef As Double
4755           Dim BCoef As Double
4756           Dim CCoef As Double

4757           Dim qy As New ADODB.Command
4758           Dim rs As New ADODB.Recordset
4759           qy.ActiveConnection = cnPumpData
4760           Dim VoltageForLookup As Integer
4761           If cmbVoltage.List(cmbVoltage.ListIndex) = "380" And cmbFrequency.List(cmbFrequency.ListIndex) = "50 Hz" Then
4762               VoltageForLookup = 460
4763           ElseIf cmbVoltage.List(cmbVoltage.ListIndex) <> "380" Then
4764               VoltageForLookup = cmbVoltage.List(cmbVoltage.ListIndex)
4765           End If
4766           qy.CommandText = "SELECT * FROM CalculatedRPM WHERE FrameNumber = '" & txtTEMCFrameNumber.Text & _
                          "' AND Voltage = '" & VoltageForLookup & "'"

4767           rs.CursorLocation = adUseClient
4768           rs.CursorType = adOpenStatic

4769           rs.Open qy
4770           If rs.RecordCount = 0 Then
4771               ACoef = 0
4772               BCoef = 0
4773               CCoef = 0
4774               MsgBox ("Cannot find coefficient data for Frame Number " & txtTEMCFrameNumber.Text & _
                          " AND Voltage = " & cmbVoltage.List(cmbVoltage.ListIndex) & _
                          " AND Frequency = " & cmbFrequency.List(cmbFrequency.ListIndex))
4775           Else
4776               ACoef = rs.Fields("A")
4777               BCoef = rs.Fields("B")
4778               CCoef = rs.Fields("C")
4779           End If


           'write header data

4780           .Range("A2").Select
4781           .ActiveCell.FormulaR1C1 = "Serial Number"
4782           .Range("C2").Select
4783           .ActiveCell.FormulaR1C1 = txtSN

4784           .Range("F1").Select
4785           .ActiveCell.FormulaR1C1 = "Customer"
4786           .Range("H1").Select
4787           .ActiveCell.FormulaR1C1 = txtShpNo

4788           .Range("A3").Select
4789           .ActiveCell.FormulaR1C1 = "Model"
4790           .Range("C3").Select
4791           .ActiveCell.FormulaR1C1 = txtModelNo

4792           .Range("F2").Select
4793           .ActiveCell.FormulaR1C1 = "Sales Order"
4794           .Range("H2").Select
4795           .ActiveCell.FormulaR1C1 = txtSalesOrderNumber

4796           .Range("A9").Select
4797           .ActiveCell.FormulaR1C1 = "Design Flow"
4798           .Range("C9").Select
4799           .ActiveCell.FormulaR1C1 = Val(txtDesignFlow)

4800           .Range("A10").Select
4801           .ActiveCell.FormulaR1C1 = "Design Head"
4802           .Range("C10").Select
4803           .ActiveCell.FormulaR1C1 = Val(txtDesignTDH)

4804           .Range("P13").Select
4805           .ActiveCell.FormulaR1C1 = "Barometric Pressure"
4806           .Range("R13").Select
4807           .ActiveCell.FormulaR1C1 = Val(txtInHgDisplay)

4808           .Range("P11").Select
4809           .ActiveCell.FormulaR1C1 = "Suction Gage Height"
4810           .Range("R11").Select
4811           .ActiveCell.FormulaR1C1 = Val(txtSuctHeight)

4812           .Range("P12").Select
4813           .ActiveCell.FormulaR1C1 = "Discharge Gage Height"
4814           .Range("R12").Select
4815           .ActiveCell.FormulaR1C1 = Val(txtDischHeight)

4816           .Range("A1").Select
4817           .ActiveCell.FormulaR1C1 = "Run Date"
4818           .Range("C1").Select
4819           .ActiveCell.FormulaR1C1 = cmbTestDate.List(cmbTestDate.ListIndex)

4820           .Range("D10:E10").Select
4821           With xlApp.Selection
4822               .HorizontalAlignment = xlCenter
4823               .VerticalAlignment = xlBottom
4824               .WrapText = False
4825               .Orientation = 0
4826               .AddIndent = False
4827               .IndentLevel = 0
4828               .ShrinkToFit = False
4829               .ReadingOrder = xlContext
4830               .MergeCells = False
4831           End With
4832           xlApp.Selection.Merge

               'determine rpm

4833           Dim RPMvalue As String
4834           If Mid$(Me.txtTEMCFrameNumber.Text, 2, 1) = "1" Then
               '1 says 2 pole
4835               If Me.cmbFrequency.ListIndex = 0 Then
                       '0 says 50Hz
4836                   RPMvalue = "2900"
4837               ElseIf Me.cmbFrequency.ListIndex = 1 Then
                       ' says 60Hz
4838                   RPMvalue = "3450"
4839               Else
                       'vfd or other, no rpm
4840                   RPMvalue = ""
4841               End If
4842           Else
               '2 says 4 pole
4843               If Me.cmbFrequency.ListIndex = 0 Then
                       '0 says 50Hz
4844                   RPMvalue = "1450"
4845               ElseIf Me.cmbFrequency.ListIndex = 1 Then
                       ' says 60Hz
4846                   RPMvalue = "1750"
4847               Else
                       'vfd or other, no rpm
4848                   RPMvalue = ""
4849               End If
4850           End If

       '        .Range("G1").Select
       '        .ActiveCell.FormulaR1C1 = "RPM"
       '        .Range("I1").Select
       '        .ActiveCell.FormulaR1C1 = RPMvalue

4851           .Range("A5").Select
4852           .ActiveCell.FormulaR1C1 = "Sp Gravity"
4853           .Range("C5").Select
4854           .ActiveCell.FormulaR1C1 = txtSpGr

4855           .Range("A6").Select
4856           .ActiveCell.FormulaR1C1 = "Viscosity"
4857           .Range("C6").Select
4858           .ActiveCell.FormulaR1C1 = txtViscosity

4859           .Range("F4").Select
4860           .ActiveCell.FormulaR1C1 = "Motor"
4861           .Range("H4").Select
4862           .ActiveCell.FormulaR1C1 = txtTEMCFrameNumber.Text

4863           .Range("H12").Select
4864           .ActiveCell.FormulaR1C1 = Me.txtCustPONum.Text

4865           .Range("F5").Select
4866           .ActiveCell.FormulaR1C1 = "Voltage"
4867           .Range("H5").Select
4868           .ActiveCell.FormulaR1C1 = cmbVoltage.List(cmbVoltage.ListIndex)

4869           .Range("K6").Select
4870           .ActiveCell.FormulaR1C1 = "End Play"
4871           .Range("M6").Select
4872           .ActiveCell.FormulaR1C1 = Val(txtEndPlay)

4873           .Range("K7").Select
4874           .ActiveCell.FormulaR1C1 = "G-Gap"
4875           .Range("M7").Select
4876           .ActiveCell.FormulaR1C1 = txtGGap.Text

4877           .Range("A8").Select
4878           .ActiveCell.FormulaR1C1 = "Design Pressure"
4879           .Range("C8").Select
4880           Dim DesPress As String
4881           DesPress = cmbTEMCDesignPressure.List(cmbTEMCDesignPressure.ListIndex)
4882           Dim j As Integer
4883           j = InStrRev(DesPress, "-")
4884           .ActiveCell.FormulaR1C1 = Mid$(DesPress, j + 2)

       '        .Range("G8").Select
       '        .ActiveCell.FormulaR1C1 = "Stator Fill"
       '        .Range("I8").Select
       '        .ActiveCell.FormulaR1C1 = "Dry"

4885           .Range("K4").Select
4886           .ActiveCell.FormulaR1C1 = "Circulation Path"
4887           .Range("M4").Select
4888           .ActiveCell.FormulaR1C1 = cmbTEMCModel.List(cmbTEMCModel.ListIndex)

4889           .Range("M8").Select
4890           .ActiveCell.FormulaR1C1 = txtNPSHr.Text

4891           .Range("K1").Select
4892           .ActiveCell.FormulaR1C1 = "Impeller Dia"
4893           .Range("M1").Select


       '        If LenB(txtImpTrim) <> 0 Then
       '            .ActiveCell.FormulaR1C1 = Val(txtImpTrim)
       '        Else
       '            .ActiveCell.FormulaR1C1 = Val(txtImpellerDia)
       '        End If
       '
4894           If chkTrimmed.value = 1 Then
4895               If Val(txtImpTrim.Text) <> 0 Then
4896                   .ActiveCell.FormulaR1C1 = txtImpTrim
4897               Else
4898                   .ActiveCell.FormulaR1C1 = txtImpellerDia
4899               End If
4900           Else
4901               .ActiveCell.FormulaR1C1 = txtImpellerDia
4902           End If



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

4903           .Range("P9").Select
4904           .ActiveCell.FormulaR1C1 = "Suction Dia"
4905           .Range("R9").Select
4906           .ActiveCell.FormulaR1C1 = cmbSuctDia.List(cmbSuctDia.ListIndex)

4907           .Range("P10").Select
4908           .ActiveCell.FormulaR1C1 = "Discharge Dia"
4909           .Range("R10").Select
4910           .ActiveCell.FormulaR1C1 = cmbDischDia.List(cmbDischDia.ListIndex)

4911           .Range("A11").Select
4912           .ActiveCell.FormulaR1C1 = "Test Spec"
4913           .Range("C11").Select
4914           .ActiveCell.FormulaR1C1 = cmbTestSpec.List(cmbTestSpec.ListIndex)

4915           .Range("K3").Select
4916           .ActiveCell.FormulaR1C1 = "Impeller Feathered"
4917           .Range("M3").Select
4918           If chkFeathered.value = 1 Then
4919               .ActiveCell.FormulaR1C1 = "Yes"
4920           Else
4921               .ActiveCell.FormulaR1C1 = "No"
4922           End If

4923           .Range("K2").Select
4924           .ActiveCell.FormulaR1C1 = "Disch Orifice"
4925           .Range("M2").Select
4926           If chkOrifice.value = 1 Then
4927               .ActiveCell.FormulaR1C1 = Val(txtOrifice)
4928           Else
4929               .ActiveCell.FormulaR1C1 = "None"
4930           End If


4931           .Range("K5").Select
4932           .ActiveCell.FormulaR1C1 = "Circulation Orifice"
4933           .Range("M5").Select
4934           If chkCircOrifice.value = 1 Then
4935               .ActiveCell.FormulaR1C1 = Val(txtCircOrifice)
4936           Else
4937               .ActiveCell.FormulaR1C1 = "None"
4938           End If

4939           .Range("A13").Select
4940           .ActiveCell.FormulaR1C1 = "Other Mods"
4941           .Range("C13").Select
4942           .ActiveCell.FormulaR1C1 = txtOtherMods

4943           .Range("A14").Select
4944           .ActiveCell.FormulaR1C1 = "Remarks"
4945           .Range("C14").Select
4946           .ActiveCell.FormulaR1C1 = txtRemarks

4947           .Range("A15").Select
4948           .ActiveCell.FormulaR1C1 = "Test Setup Remarks"
4949           .Range("C15").Select
4950           .ActiveCell.FormulaR1C1 = txtTestSetupRemarks

4951           .Range("P1").Select
4952           .ActiveCell.FormulaR1C1 = "Suct ID"
4953           .Range("R1").Select
4954           .ActiveCell.FormulaR1C1 = cmbSuctionPressureTransducer.List(cmbSuctionPressureTransducer.ListIndex)

4955           .Range("P2").Select
4956           .ActiveCell.FormulaR1C1 = "Disch ID"
4957           .Range("R2").Select
4958           .ActiveCell.FormulaR1C1 = cmbDischargePressureTransducer.List(cmbDischargePressureTransducer.ListIndex)

4959           .Range("P3").Select
4960           .ActiveCell.FormulaR1C1 = "Temp ID"
4961           .Range("R3").Select
4962           .ActiveCell.FormulaR1C1 = cmbTemperatureTransducer.List(cmbTemperatureTransducer.ListIndex)

4963           .Range("P4").Select
4964           .ActiveCell.FormulaR1C1 = "Circ Flow ID"
4965           .Range("R4").Select
4966           .ActiveCell.FormulaR1C1 = cmbCirculationFlowMeter.List(cmbCirculationFlowMeter.ListIndex)

4967           .Range("P5").Select
4968           .ActiveCell.FormulaR1C1 = "Flow ID"
4969           .Range("R5").Select
4970           .ActiveCell.FormulaR1C1 = cmbFlowMeter.List(cmbFlowMeter.ListIndex)

4971           .Range("P6").Select
4972           .ActiveCell.FormulaR1C1 = "Analyzer ID"
4973           .Range("R6").Select
4974           .ActiveCell.FormulaR1C1 = cmbAnalyzerNo.List(cmbAnalyzerNo.ListIndex)

4975           .Range("P7").Select
4976           .ActiveCell.FormulaR1C1 = "Loop ID"
4977           .Range("R7").Select
4978           .ActiveCell.FormulaR1C1 = cmbLoopNumber.List(cmbLoopNumber.ListIndex)

4979           .Range("A4").Select
4980           .ActiveCell.FormulaR1C1 = "Fluid"
4981           .Range("C4").Select
4982           .ActiveCell.FormulaR1C1 = txtLiquid.Text

4983           .Range("F3").Select
4984           .ActiveCell.FormulaR1C1 = "Cust PN"
4985           .Range("H3").Select
       '        .ActiveCell.FormulaR1C1 = txtRMA.Text
4986           If rsPumpData.Fields("RVSPartNo") <> "" Then
4987               .ActiveCell.FormulaR1C1 = rsPumpData.Fields("RVSPartNo")
4988           End If
4989           If rsPumpData.Fields("CustPN") <> "" Then
4990               .ActiveCell.FormulaR1C1 = rsPumpData.Fields("CustPN")
4991           End If

4992           .Range("A7").Select
4993           .ActiveCell.FormulaR1C1 = "Temperature"
4994           .Range("C7").Select
4995           .ActiveCell.FormulaR1C1 = txtLiquidTemperature.Text

4996           .Range("F6").Select
4997           .ActiveCell.FormulaR1C1 = "Frequency"
4998           .Range("H6").Select
4999           If UCase(cmbFrequency.List(cmbFrequency.ListIndex)) = "VFD" Then
5000               .ActiveCell.FormulaR1C1 = Val(Me.txtVFDFreq)
5001           Else
5002               .ActiveCell.FormulaR1C1 = Val(cmbFrequency.List(cmbFrequency.ListIndex))
5003           End If
       '        .Range("K2").Select
       '        .ActiveCell.FormulaR1C1 = "Disch Orifice"
       '        .Range("M2").Select
       '        .ActiveCell.FormulaR1C1 = txtOrifice.Text

       '        .Range("K12").Select
       '        .ActiveCell.FormulaR1C1 = "Flow Orifice"
       '        .Range("L12").Select
       '        .ActiveCell.FormulaR1C1 = txtCircOrifice.Text

5004           .Range("P8").Select
5005           .ActiveCell.FormulaR1C1 = "PLC No"
5006           .Range("R8").Select
5007           .ActiveCell.FormulaR1C1 = cmbPLCNo.List(cmbPLCNo.ListIndex)

5008           .Range("F7").Select
5009           .ActiveCell.FormulaR1C1 = "Phases"
5010           .Range("H7").Select
5011           .ActiveCell.FormulaR1C1 = txtNoPhases.Text

5012           .Range("F8").Select
5013           .ActiveCell.FormulaR1C1 = "Poles"
5014           .Range("H8").Select
5015           .ActiveCell.FormulaR1C1 = 2 * Val(Left$(Right$(txtTEMCFrameNumber.Text, 2), 1))

5016           .Range("F9").Select
5017           .ActiveCell.FormulaR1C1 = "Rated Current"
5018           .Range("H9").Select
5019           .ActiveCell.FormulaR1C1 = txtAmps.Text

5020           .Range("F10").Select
5021           .ActiveCell.FormulaR1C1 = "Rated Input Power"
5022           .Range("H10").Select
5023           .ActiveCell.FormulaR1C1 = txtRatedInputPower.Text

5024           .Range("F11").Select
5025           .ActiveCell.FormulaR1C1 = "Insulation Class"
5026           .Range("H11").Select
5027           .ActiveCell.FormulaR1C1 = txtThermalClass.Text

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

5028           .Range("A17").Select
5029           .ActiveCell.FormulaR1C1 = "Flow"
5030           .Range("A18").Select
5031           .ActiveCell.FormulaR1C1 = "(GPM)"

5032           .Range("B17").Select
5033           .ActiveCell.FormulaR1C1 = "TDH"
5034           .Range("B18").Select
5035           .ActiveCell.FormulaR1C1 = "(Ft)"

5036           .Range("C17").Select
5037           .ActiveCell.FormulaR1C1 = "KW"

5038           .Range("D17").Select
5039           .ActiveCell.FormulaR1C1 = "Ave"
5040           .Range("D18").Select
5041           .ActiveCell.FormulaR1C1 = "Volts"

5042           .Range("E17").Select
5043           .ActiveCell.FormulaR1C1 = "Ave"
5044           .Range("E18").Select
5045           .ActiveCell.FormulaR1C1 = "Amps"

5046           .Range("F17").Select
5047           .ActiveCell.FormulaR1C1 = "Power"
5048           .Range("F18").Select
5049           .ActiveCell.FormulaR1C1 = "Factor"

5050           .Range("G17").Select
5051           .ActiveCell.FormulaR1C1 = "Overall"
5052           .Range("G18").Select
5053           .ActiveCell.FormulaR1C1 = "Eff"

5054           .Range("H17").Select
5055           .ActiveCell.FormulaR1C1 = "Measured"
5056           .Range("H18").Select
5057           .ActiveCell.FormulaR1C1 = "RPM"

5058           .Range("I17").Select
5059           .ActiveCell.FormulaR1C1 = "Calculated"
5060           .Range("I18").Select
5061           .ActiveCell.FormulaR1C1 = "RPM"

5062           .Range("J17").Select
5063           .ActiveCell.FormulaR1C1 = "Suction"
5064           .Range("J18").Select
5065           .ActiveCell.FormulaR1C1 = "Temp(F)"

5066           .Range("K17").Select
5067           .ActiveCell.FormulaR1C1 = "Disch"
5068           .Range("K18").Select
5069           .ActiveCell.FormulaR1C1 = "Pressure"

5070           .Range("L17").Select
5071           .ActiveCell.FormulaR1C1 = "Suction"
5072           .Range("L18").Select
5073           .ActiveCell.FormulaR1C1 = "Pressure"

5074           .Range("M17").Select
5075           .ActiveCell.FormulaR1C1 = "Vel"
5076           .Range("M18").Select
5077           .ActiveCell.FormulaR1C1 = "Head"

5078           .Range("N17").Select
5079           .ActiveCell.FormulaR1C1 = "Axial"
5080           .Range("N18").Select
5081           .ActiveCell.FormulaR1C1 = "Position"

5082           .Range("O17").Select
5083           .ActiveCell.FormulaR1C1 = "Pct of"
5084           .Range("O18").Select
5085           .ActiveCell.FormulaR1C1 = "End Play"

5086           .Range("P17").Select
5087           .ActiveCell.FormulaR1C1 = "Hydraulic"
5088           .Range("P18").Select
5089           .ActiveCell.FormulaR1C1 = "Efficiency"

       '        .Range("P17").Select
       '        .ActiveCell.FormulaR1C1 = "Circ"
       '        .Range("P18").Select
       '        .ActiveCell.FormulaR1C1 = "Flow"

5090           .Range("Q17").Select
5091           .ActiveCell.FormulaR1C1 = "Motor"
5092           .Range("Q18").Select
5093           .ActiveCell.FormulaR1C1 = "Efficiency"

5094           .Range("S17").Select
5095           .ActiveCell.FormulaR1C1 = "NPSHa"

5096           .Range("T17").Select
5097           .ActiveCell.FormulaR1C1 = "Phase 1"
5098           .Range("T18").Select
5099           .ActiveCell.FormulaR1C1 = "Current"

5100           .Range("U17").Select
5101           .ActiveCell.FormulaR1C1 = "Phase 2"
5102           .Range("U18").Select
5103           .ActiveCell.FormulaR1C1 = "Current"

5104           .Range("V17").Select
5105           .ActiveCell.FormulaR1C1 = "Phase 3"
5106           .Range("V18").Select
5107           .ActiveCell.FormulaR1C1 = "Current"

5108           .Range("W17").Select
5109           .ActiveCell.FormulaR1C1 = "Phase 1"
5110           .Range("W18").Select
5111           .ActiveCell.FormulaR1C1 = "Voltage"

5112           .Range("X17").Select
5113           .ActiveCell.FormulaR1C1 = "Phase 2"
5114           .Range("X18").Select
5115           .ActiveCell.FormulaR1C1 = "Voltage"

5116           .Range("Y17").Select
5117           .ActiveCell.FormulaR1C1 = "Phase 3"
5118           .Range("Y18").Select
5119           .ActiveCell.FormulaR1C1 = "Voltage"

5120           .Range("Z17").Select
5121           .ActiveCell.FormulaR1C1 = "'" & txtTitle(20).Text

5122           .Range("Z18").Select
5123           .ActiveCell.FormulaR1C1 = "'" & txtTitle(21).Text

5124           .Range("AA17").Select
5125           .ActiveCell.FormulaR1C1 = "'" & txtTitle(22).Text

5126           .Range("AA18").Select
5127           .ActiveCell.FormulaR1C1 = "'" & txtTitle(23).Text

5128           .Range("AB17").Select
5129           .ActiveCell.FormulaR1C1 = "'" & txtTitle(24).Text

5130           .Range("AB18").Select
5131           .ActiveCell.FormulaR1C1 = "'" & txtTitle(25).Text

5132           .Range("AC17").Select
5133           .ActiveCell.FormulaR1C1 = "HR"

5134           .Range("AC18").Select
5135           .ActiveCell.FormulaR1C1 = "(ft)"

5136           .Range("AD17").Select
5137           .ActiveCell.FormulaR1C1 = "'" & txtTitle(26).Text

5138           .Range("AD18").Select
5139           .ActiveCell.FormulaR1C1 = "'" & txtTitle(27).Text

5140           .Range("AE17").Select
5141           .ActiveCell.FormulaR1C1 = "TRG"
5142           .Range("AE18").Select
5143           .ActiveCell.FormulaR1C1 = "Position"

5144           .Range("AF17").Select
5145           .ActiveCell.FormulaR1C1 = "Thrust"

5146           .Range("AG17").Select
5147           .ActiveCell.FormulaR1C1 = "F/R"

5148           .Range("AH17").Select
5149           .ActiveCell.FormulaR1C1 = "Moment"
5150           .Range("AH18").Select
5151           .ActiveCell.FormulaR1C1 = "Arm"

5152           .Range("AI17").Select
5153           .ActiveCell.FormulaR1C1 = "Rig"
5154           .Range("AI18").Select
5155           .ActiveCell.FormulaR1C1 = "Pressure"

       '        .Range("AI17").Select
       '        .ActiveCell.FormulaR1C1 = "Viscosity"

5156           .Range("AJ19").Select
5157           .ActiveCell.FormulaR1C1 = "Rear"
5158           .Range("AJ18").Select
5159           .ActiveCell.FormulaR1C1 = "Force"

5160           .Range("AK17").Select
5161           .ActiveCell.FormulaR1C1 = "PV"

5162           .Range("R17").Select
5163           .ActiveCell.FormulaR1C1 = "Shaft"
5164           .Range("R18").Select
5165           .ActiveCell.FormulaR1C1 = "Power"

       '        .Range("AM17").Select
       '        .ActiveCell.FormulaR1C1 = "Pct Full"
       '        .Range("AM18").Select
       '        .ActiveCell.FormulaR1C1 = "Scale"

5166           .Range("AL17").Select
5167           .ActiveCell.FormulaR1C1 = "NPSHr"

5168           .Range("AM17").Select
5169           .ActiveCell.FormulaR1C1 = "Remarks"




               'now output the data

5170           iRowNo = 20

5171           rsEff.MoveFirst
5172           For I = 1 To frmPLCData.UpDown2.value
5173               .Range("A" & iRowNo).Select
5174               .ActiveCell.FormulaR1C1 = rsEff.Fields("Flow")

5175               .Range("B" & iRowNo).Select
5176               .ActiveCell.FormulaR1C1 = rsEff.Fields("TDH")

5177               .Range("C" & iRowNo).Select
5178               .ActiveCell.FormulaR1C1 = rsEff.Fields("KW")

5179               .Range("D" & iRowNo).Select
5180               .ActiveCell.FormulaR1C1 = rsEff.Fields("Volts")

5181               .Range("E" & iRowNo).Select
5182               .ActiveCell.FormulaR1C1 = rsEff.Fields("Amps")

5183               .Range("F" & iRowNo).Select
5184               .ActiveCell.FormulaR1C1 = rsEff.Fields("PowerFactor")

5185               .Range("G" & iRowNo).Select
5186               .ActiveCell.FormulaR1C1 = rsEff.Fields("OverallEfficiency")

5187               .Range("H" & iRowNo).Select
5188               .ActiveCell.FormulaR1C1 = rsEff.Fields("RPM")

5189               .Range("I" & iRowNo).Select
                   'use the coefficients from above to calculate rpm
5190               Dim f As Double
5191               f = .Range("H6").value
5192               .ActiveCell.FormulaR1C1 = (Val(f) / 60) * (ACoef * (rsEff.Fields("KW")) ^ 2 + BCoef * (rsEff.Fields("KW")) + CCoef)

5193               .Range("J" & iRowNo).Select
5194               .ActiveCell.FormulaR1C1 = rsEff.Fields("Temperature")

5195               .Range("K" & iRowNo).Select
5196               .ActiveCell.FormulaR1C1 = rsEff.Fields("DischPress")

5197               .Range("L" & iRowNo).Select
5198               .ActiveCell.FormulaR1C1 = rsEff.Fields("SuctPress")

5199               .Range("M" & iRowNo).Select
5200               .ActiveCell.FormulaR1C1 = rsEff.Fields("VelocityHead")

5201               .Range("N" & iRowNo).Select
5202               .ActiveCell.FormulaR1C1 = rsEff.Fields("Pos")

5203               .Range("O" & iRowNo).Select
5204               .ActiveCell.FormulaR1C1 = 100 * rsEff.Fields("Pos") / Val(txtEndPlay)

5205               .Range("P" & iRowNo).Select
5206               .ActiveCell.FormulaR1C1 = rsEff.Fields("HydraulicEfficiency")

       '            .Range("P" & iRowNo).Select
       '            .ActiveCell.FormulaR1C1 = rsEff.Fields("CircFlow")

5207               .Range("Q" & iRowNo).Select
5208               .ActiveCell.FormulaR1C1 = rsEff.Fields("MotorEfficiency")

5209               .Range("S" & iRowNo).Select
5210               .ActiveCell.FormulaR1C1 = rsEff.Fields("NPSHa")

5211               .Range("T" & iRowNo).Select
5212               .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentA")

5213               .Range("U" & iRowNo).Select
5214               .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentB")

5215               .Range("V" & iRowNo).Select
5216               .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentC")

5217               .Range("W" & iRowNo).Select
5218               .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageA")

5219               .Range("X" & iRowNo).Select
5220               .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageB")

5221               .Range("Y" & iRowNo).Select
5222               .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageC")

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

5223               .Range("Z" & iRowNo).Select
5224               .ActiveCell.FormulaR1C1 = rsEff.Fields("CircFlow")

5225               .Range("AA" & iRowNo).Select
5226               .ActiveCell.FormulaR1C1 = rsEff.Fields("RBHTemp")

5227               .Range("AB" & iRowNo).Select
5228               .ActiveCell.FormulaR1C1 = rsEff.Fields("RBHPress")

5229               .Range("AC" & iRowNo).Select
5230               .ActiveCell.FormulaR1C1 = (rsEff.Fields("RBHPress") - rsEff.Fields("SuctPress")) * 2.31

5231               .Range("AD" & iRowNo).Select
5232               .ActiveCell.FormulaR1C1 = rsEff.Fields("AI4")

5233               .Range("AE" & iRowNo).Select
5234               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCTRG")

5235               .Range("AF" & iRowNo).Select
5236               If rsEff.Fields("TEMCFrontThrust") = 0 Then
5237                   If rsEff.Fields("TEMCRearThrust") = 0 Then
5238                       .ActiveCell.FormulaR1C1 = " "
5239                       .Range("AG" & iRowNo).Select
5240                       .ActiveCell.FormulaR1C1 = " "
5241                   Else
5242                       .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCRearThrust")
5243                       .Range("AG" & iRowNo).Select
5244                       .ActiveCell.FormulaR1C1 = "R"
5245                   End If
5246               Else
5247                   .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCFrontThrust")
5248                   .Range("AG" & iRowNo).Select
5249                   .ActiveCell.FormulaR1C1 = "F"
5250               End If

5251               .Range("AH" & iRowNo).Select
5252               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCMomentArm")

5253               .Range("AI" & iRowNo).Select
5254               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCThrustRigPressure")

       '            .Range("AJ" & iRowNo).Select
       '            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCViscosity")

5255               .Range("AJ" & iRowNo).Select
5256               If rsEff.Fields("TEMCForceDirection") = "F" Then
5257                   .ActiveCell.FormulaR1C1 = -rsEff.Fields("TEMCCalculatedForce")
5258               Else
5259                   .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCCalculatedForce")
5260               End If

5261               .Range("AK" & iRowNo).Select
5262               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCPV")

5263               .Range("R" & iRowNo).Select
5264               .ActiveCell.FormulaR1C1 = rsEff.Fields("KW") * rsEff.Fields("MotorEfficiency") / 100

5265               .Range("AL" & iRowNo).Select
5266               .ActiveCell.FormulaR1C1 = rsEff.Fields("NPSHr")

       '            If RatedKW = 999 Then
       '                .ActiveCell.FormulaR1C1 = ""
       '            Else
       '                .ActiveCell.FormulaR1C1 = (rsEff.Fields("KW") * rsEff.Fields("MotorEfficiency")) / (1 * RatedKW)
       '            End If

5267               .Range("AM" & iRowNo).Select
5268               .ActiveCell.FormulaR1C1 = rsEff.Fields("Remarks")


5269               rsEff.MoveNext
5270               iRowNo = iRowNo + 1
5271           Next I

5272           .Range("A20:AS30").Select
5273           .Selection.NumberFormat = "0.00"

           'set up formulas to calculate BEP
           '  first, plot 2nd order polynomial for flow vs hydraulic efficiency
           '  the formulas for doing that are in E68, F68 and G68
           '  only want the formulas to point to the number of points in the test data, so use frmPLCData.CWNumEdit2.value
           '
5274       Dim AColumnRow As String
5275       Dim PColumnRow As String

5276       AColumnRow = "A" & Trim(str(19 + frmPLCData.UpDown2.value))
5277       PColumnRow = "P" & Trim(str(19 + frmPLCData.UpDown2.value))

5278           .Range("E68").Select
5279           .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1)"

5280           .Range("F68").Select
5281           .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1,2)"

5282           .Range("G68").Select
5283           .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1,3)"

           'export balance holes
5284       If boGotBalanceHoles Then
5285           If rsBalanceHoles.State = adStateClosed Then
5286               rsBalanceHoles.ActiveConnection = cnPumpData
5287               rsBalanceHoles.Open
5288           End If 'rsBalanceHoles.State = adStateClosed

5289           If rsBalanceHoles.RecordCount <> 0 Then

5290               .Range("K9:N9").Merge
5291               .Range("K9:N9").Formula = "Balance Hole Data"
5292               .Range("K9:N9").HorizontalAlignment = xlCenter

5293               .Range("K10").Select
5294               .ActiveCell.Formula = "Date"

5295               .Range("L10").Select
5296               .ActiveCell.Formula = "Number"

5297               .Range("M10").Select
5298               .ActiveCell.Formula = "Diameter"

5299               .Range("N10").Select
5300               .ActiveCell.Formula = "Bolt Circle"

5301               iRowNo = 11

5302               If rsBalanceHoles.RecordCount > 3 Then
5303                   For I = 1 To rsBalanceHoles.RecordCount - 3
5304                       Rows("13:13").Select
5305                       Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
5306                   Next I
5307               End If

5308               rsBalanceHoles.MoveFirst
5309               For I = 1 To rsBalanceHoles.RecordCount

5310                   .Range("K" & iRowNo).Select
5311                   .ActiveCell.Formula = rsBalanceHoles.Fields("Date")
5312                   .ActiveCell.NumberFormat = "m/d/yy h:mm AM/PM;@"
5313                   .Range("L" & iRowNo).Select
5314                   .ActiveCell = rsBalanceHoles.Fields("Number")
5315                   .ActiveCell.NumberFormat = "0"
5316                   .Range("M" & iRowNo).Select
5317                   If IsNumeric(rsBalanceHoles.Fields("Diameter1")) Then
5318                       .ActiveCell = Val(rsBalanceHoles.Fields("Diameter1"))
5319                       .ActiveCell.NumberFormat = "0.0000"
5320                   Else
5321                       .ActiveCell = rsBalanceHoles.Fields("Diameter1")
5322                   End If

5323                   .Range("N" & iRowNo).Select
5324                   If IsNumeric(rsBalanceHoles.Fields("BoltCircle1")) Then
5325                       .ActiveCell = Val(rsBalanceHoles.Fields("BoltCircle1"))
5326                       .ActiveCell.NumberFormat = "0.0000"
5327                   Else
5328                       .ActiveCell = rsBalanceHoles.Fields("BoltCircle1")
5329                   End If

5330                   rsBalanceHoles.MoveNext
5331                   iRowNo = iRowNo + 1
5332               Next I
5333               .Range("K10:N" & iRowNo - 1).Select
5334               With .Selection.Interior
5335                   .ColorIndex = 34
5336                   .Pattern = xlSolid
5337               End With
5338           End If 'rsBalanceHoles.RecordCount <> 0
5339       End If ' boGotBalanceHoles

           'plot graphs

5340       Dim SeriesName As String
5341       Dim XVals As String
5342       Dim YVals As String
5343       Dim RowNo As Long
5344       Dim RowStr As String
5345       Dim LastPoint As Integer
5346       Dim LineType As String
5347       Dim AxisGroup As Integer
5348       Dim LabelPos As Integer
5349       Dim LineColor As Long

5350           .ActiveSheet.ChartObjects("HydRepChart").Activate
5351           Dim S As Series
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
5352           Dim aq As Double
5353           Range("AQ56", "AQ71").Select
5354           aq = .Max(Selection)
5355           Dim ax As Double
5356           Range("AX56", "AX71").Select
5357           ax = .Max(Selection)

               'then current (as and az)
5358           Dim at As Double
5359           Range("AS56", "AS71").Select
5360           at = .Max(Selection)
5361           Dim ba As Double
5362           Range("AZ56", "AZ71").Select
5363           ba = .Max(Selection)

5364           Dim CurrentScaleMax As Integer
5365           Dim TDHScaleMax As Integer

5366           Dim MaxTDH As Integer
5367           With Application.WorksheetFunction
5368               If aq > ax Then
5369                   MaxTDH = .Ceiling(aq, 25)
5370               Else
5371                   MaxTDH = .Ceiling(ax, 25)
5372               End If
5373           End With

5374           Dim MaxCurrent As Integer
5375           With Application.WorksheetFunction
5376               If at > ba Then
5377                   Select Case at
                           Case Is <= 5
5378                           CurrentScaleMax = 5

5379                       Case Is <= 10
5380                           CurrentScaleMax = 10

5381                       Case Else
5382                           CurrentScaleMax = 25
5383                   End Select

5384                   MaxCurrent = .Ceiling(at, CurrentScaleMax)
5385               Else
5386                  Select Case ba
                           Case Is <= 5
5387                           CurrentScaleMax = 5

5388                       Case Is <= 10
5389                           CurrentScaleMax = 10

5390                       Case Else
5391                           CurrentScaleMax = 25
5392                   End Select

5393                   MaxCurrent = .Ceiling(ba, CurrentScaleMax)
5394               End If
5395           End With

5396           ActiveSheet.ChartObjects("HydRepChart").Activate
5397            Dim ShtName As String
5398            ShtName = "'" & ActiveSheet.Name & "'"

5399           RowStr = 56 + 15
5400            For I = 1 To 8

5401                Select Case I
                        Case 1
5402                        SeriesName = "=""TDH"""
5403                        XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
5404                        YVals = "=" & ShtName & "!$AQ$56:$AQ$" & RowStr
5405                        LineType = msoLineSolid
5406                        AxisGroup = 1
5407                        LabelPos = xlLabelPositionRight
5408                        LineColor = vbBlue

5409                    Case 2
5410                        SeriesName = "=""Input Power"""
5411                        XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
5412                        YVals = "=" & ShtName & "!$AR$56:$AR$" & RowStr
5413                        LineType = msoLineSolid
5414                        AxisGroup = 2
5415                        LabelPos = xlLabelPositionRight
5416                        LineColor = vbRed

5417                    Case 3
5418                        SeriesName = "=""Current"""
5419                        XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
5420                        YVals = "=" & ShtName & "!$AS$56:$AS$" & RowStr
5421                        LineType = msoLineSolid
5422                        AxisGroup = 2
5423                        LabelPos = xlLabelPositionRight
5424                        LineColor = vbGreen

5425                    Case 4
       '                     SeriesName = "=""Overall Eff"""
       '                     XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
       '                     YVals = "=" & ShtName & "!$AT$56:$AT$" & RowStr
       '                     LineType = msoLineSolid
       '                     AxisGroup = 2
       '                     LabelPos = xlLabelPositionRight
       '                     LineColor = vbCyan

5426                    Case 5
5427                        SeriesName = "=""TDH (Adj)"""
5428                        XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
5429                        YVals = "=" & ShtName & "!$AX$56:$AX$" & RowStr
5430                        LineType = msoLineDash
5431                        AxisGroup = 1
5432                        LabelPos = xlLabelPositionBelow
5433                        LineColor = vbBlue

5434                    Case 6
5435                        SeriesName = "=""Input Power (Adj)"""
5436                        XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
5437                        YVals = "=" & ShtName & "!$AY$56:$AY$" & RowStr
5438                        LineType = msoLineDash
5439                        AxisGroup = 2
5440                        LabelPos = xlLabelPositionBelow
5441                        LineColor = vbRed

5442                    Case 7
5443                        SeriesName = "=""Current (Adj)"""
5444                        XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
5445                        YVals = "=" & ShtName & "!$AZ$56:$AZ$" & RowStr
5446                        LineType = msoLineDash
5447                        AxisGroup = 2
5448                        LabelPos = xlLabelPositionBelow
5449                        LineColor = vbGreen

5450                    Case 8
       '                     SeriesName = "=""Overall Eff (Adj)"""
       '                     XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
       '                     YVals = "=" & ShtName & "!$BA$56:$BA$" & RowStr
       '                     LineType = msoLineDash
       '                     AxisGroup = 2
       '                     LabelPos = xlLabelPositionBelow
       '                     LineColor = vbCyan

5451               End Select
5452               LastPoint = 16
5453               ActiveChart.SeriesCollection.NewSeries
5454               ActiveChart.SeriesCollection(I).Name = SeriesName
5455               ActiveChart.SeriesCollection(I).XValues = XVals
5456               ActiveChart.SeriesCollection(I).Values = YVals
5457               ActiveChart.SeriesCollection(I).Select
5458               ActiveChart.SeriesCollection(I).Points(LastPoint).Select
5459               ActiveChart.SeriesCollection(I).Points(LastPoint).ApplyDataLabels
5460               ActiveChart.SeriesCollection(I).Points(LastPoint).DataLabel.Select
5461               If I < 5 Then
5462                   Selection.ShowSeriesName = True
5463                   Selection.Position = LabelPos
5464               Else
5465                   Selection.ShowSeriesName = False
5466               End If
5467               Selection.ShowValue = False
5468               ActiveChart.SeriesCollection(I).ChartType = xlXYScatterSmoothNoMarkers
5469               ActiveChart.SeriesCollection(I).Select
5470               With Selection.Format.line
5471                   .Visible = msoTrue
5472                   .DashStyle = LineType
5473                   .ForeColor.RGB = LineColor
5474               End With


5475               ActiveChart.SeriesCollection(I).AxisGroup = AxisGroup
5476               ActiveChart.SeriesCollection(I).DataLabels.Font.Size = 8
5477               ActiveChart.SeriesCollection(I).DataLabels.Font.Name = "Arial"
5478           Next I

               'show design point
5479           SeriesName = "=""Design Point"""
5480           XVals = "=" & ShtName & "!$L$63"
5481           YVals = "=" & ShtName & "!$L$64"
5482           LineType = msoLineSolid
5483           AxisGroup = 1
5484           ActiveChart.SeriesCollection.NewSeries
5485           ActiveChart.SeriesCollection(I).Name = SeriesName
5486           ActiveChart.SeriesCollection(I).XValues = XVals
5487           ActiveChart.SeriesCollection(I).Values = YVals
5488           ActiveChart.SeriesCollection(I).Select

5489           Selection.MarkerStyle = 4
5490           Selection.MarkerSize = 7
5491           With Selection.Format.line
5492               .Visible = msoTrue
5493               .Weight = 2.25
5494               .ForeColor.RGB = vbBlack
5495           End With


5496           ActiveChart.Axes(xlValue).Select
5497           ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
5498           ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True

5499           ActiveChart.Axes(xlValue).MaximumScale = MaxTDH
5500           ActiveChart.Axes(xlValue).MinimumScale = 0
5501           ActiveChart.Axes(xlValue).MajorUnit = Int(MaxTDH / 5)
5502           Selection.TickLabels.NumberFormat = "0"

5503           ActiveChart.Axes(xlValue, xlSecondary).Select
5504           ActiveChart.Axes(xlValue, xlSecondary).MinimumScaleIsAuto = True
5505           ActiveChart.Axes(xlValue, xlSecondary).MaximumScaleIsAuto = True

5506           ActiveChart.Axes(xlValue, xlSecondary).MaximumScale = MaxCurrent
5507           ActiveChart.Axes(xlValue, xlSecondary).MinimumScale = 0
5508           ActiveChart.Axes(xlValue, xlSecondary).MajorUnit = Int(MaxCurrent / 5)
5509           Selection.TickLabels.NumberFormat = "0"

5510           ActiveChart.Axes(xlValue, xlSecondary).HasTitle = True
5511           ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = "Input Power (kW)-Current (A)"
       '        ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = "Input Power (kW)-Current (A)-Overall Efficiency (%)"
5512           ActiveChart.SetElement (msoElementSecondaryValueAxisTitleRotated)
               'ActiveSheet.PageSetup.PrintArea = "$CA$1:$CI$50"

5513           Range("A1").Select

               'delete all macros in the excel file

               ' Declare variables to access the macros in the workbook.
5514           Dim objProject As VBIDE.VBProject
5515           Dim objComponent As VBIDE.VBComponent
5516           Dim objCode As VBIDE.CodeModule

               ' Get the project details in the workbook.
5517           Set objProject = xlBook.VBProject

               ' Iterate through each component in the project.
5518           For Each objComponent In objProject.VBComponents

                   ' Delete code modules
5519               Set objCode = objComponent.CodeModule
5520               objCode.DeleteLines 1, objCode.CountOfLines

5521               Set objCode = Nothing
5522               Set objComponent = Nothing
5523           Next

5524           Set objProject = Nothing


5525           xlApp.Visible = True                    'show the sheet

5526           xlApp.VBE.ActiveVBProject.VBComponents.Import ParentDirectoryName & sSaveFileMacroFile
5527           xlApp.Run "AssignButton"
5528       End With

       '    Exit Sub

5529   ErrHandler:
           'User pressed the Cancel button

5530       On Error GoTo notopen
5531       If Not xlApp.ActiveWorkbook Is Nothing Then
5532           ActiveWorkbook.CheckCompatibility = False
5533           xlApp.ActiveWorkbook.Save               'save the workbook
               'xlApp.ActiveWorkbook.Close

5534       End If

5535   notopen:

       '    xlApp.Application.Quit

       '    xlApp.Quit
       '    Set xlApp = Nothing

       '    If CommonDialog1.filename <> "" Then
       '        MsgBox CommonDialog1.filename & " has been written.", vbOKOnly, "File Opened"
       '    End If

5536   On Error GoTo vbwErrHandler

' <VB WATCH>
5537       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5538       Exit Sub
' <VB WATCH>
5539       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5540       Exit Sub
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
5541       On Error GoTo vbwErrHandler
5542       Const VBWPROCNAME = "frmPLCData.GetWorksheetTabs"
5543       If vbwProtector.vbwTraceProc Then
5544           Dim vbwProtectorParameterString As String
5545           If vbwProtector.vbwTraceParameters Then
5546               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("filename", filename) & ", "
5547               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("WorkSheetName", WorkSheetName) & ") "
5548           End If
5549           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5550       End If
' </VB WATCH>

           'see what worksheet tabs alread exist in the excel worksheet

5551       Dim intSheets As Integer    'number of sheets in the workbook
5552       Dim I As Integer
5553       Dim S As String
5554       Dim ans As Integer
5555       Dim NameOK As Boolean

5556       intSheets = xlApp.Worksheets.Count      'how many sheets are there?

           'define a crlf string
5557       S = vbCrLf

5558       For I = 1 To intSheets
5559           S = S & xlApp.Worksheets(I).Name & vbCrLf   'add in the worksheet name
5560       Next I

           'tell the user the names so far and ask if he/she wants to add another
5561       ans = MsgBox("You have the following Worksheet Names in " & filename & ": " & S & "Do you want to add another sheet to this file?", vbYesNo, "Sheets in Excel File")

           'get the answer
5562       If ans = vbNo Then
5563           GetWorksheetTabs = vbNo     'set up flag for when we return to the calling subroutine
' <VB WATCH>
5564       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5565           Exit Function
5566       End If

           'get worksheet name from user and check to see that it's not already used

5567       NameOK = False  'start assuming that the name is bad

5568       While Not NameOK    'as long as it's bad, stay in this loop
5569           WorkSheetName = InputBox("Enter Worksheet Name for this run.")  'ask for name

5570           If WorkSheetName = "" Then      'if we get a nul return or user presses cancel
5571               GetWorksheetTabs = vbNo
' <VB WATCH>
5572       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5573               Exit Function
5574           End If

5575           For I = 1 To xlApp.Worksheets.Count     'go through all of the existing sheets
5576               If WorkSheetName = xlApp.Worksheets(I).Name Then        'if the names are the same
5577                   MsgBox "The name " & WorkSheetName & " already exists for a Worksheet.  Please try again.", vbOKOnly, "Bad Worksheet Name"  'tell the user
5578                   NameOK = False
5579                   Exit For
5580               End If
5581               NameOK = True       'if we make it thru say the name is ok
5582           Next I
5583       Wend

5584       xlApp.Worksheets.Add , xlApp.Worksheets(xlApp.Worksheets.Count)     'add a worksheer
5585       xlApp.Worksheets(xlApp.Worksheets.Count).Name = WorkSheetName       'give it the desired name
5586       GetWorksheetTabs = vbYes                                            'say that the results were ok

' <VB WATCH>
5587       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5588       Exit Function
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
5589       On Error GoTo vbwErrHandler
5590       Const VBWPROCNAME = "frmPLCData.NewWorkBook"
5591       If vbwProtector.vbwTraceProc Then
5592           Dim vbwProtectorParameterString As String
5593           If vbwProtector.vbwTraceParameters Then
5594               vbwProtectorParameterString = "()"
5595           End If
5596           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5597       End If
' </VB WATCH>

5598       Dim WorkSheetName As String

           'we've just added a new workbook, delete sheet1, sheet2, etc
5599       xlApp.DisplayAlerts = False
5600       While xlApp.Worksheets.Count > 1
5601           xlApp.Worksheets(1).Delete          'delete the sheet
5602       Wend
5603       xlApp.DisplayAlerts = True

5604       WorkSheetName = InputBox("Enter Title Worksheet Name for this run.")    'get the desired name
5605       xlApp.Worksheets(1).Name = WorkSheetName    'and name the sheet

5606       NewWorkBook = WorkSheetName

' <VB WATCH>
5607       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5608       Exit Function
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
5609       On Error GoTo vbwErrHandler
5610       Const VBWPROCNAME = "frmPLCData.CalibrateSoftware"
5611       If vbwProtector.vbwTraceProc Then
5612           Dim vbwProtectorParameterString As String
5613           If vbwProtector.vbwTraceParameters Then
5614               vbwProtectorParameterString = "()"
5615           End If
5616           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5617       End If
' </VB WATCH>
5618           frmCalibrate.Show
               'Calibrating = True

' <VB WATCH>
5619       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5620       Exit Sub
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
5621       On Error GoTo vbwErrHandler
5622       Const VBWPROCNAME = "frmPLCData.ParseTEMCModelNo"
5623       If vbwProtector.vbwTraceProc Then
5624           Dim vbwProtectorParameterString As String
5625           If vbwProtector.vbwTraceParameters Then
5626               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
5627               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ltr", ltr) & ") "
5628           End If
5629           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5630       End If
' </VB WATCH>
5631       Dim I As Integer
5632       Dim iStart As Integer
5633       Dim iStop As Integer
5634       Dim strCompare As String

5635       For I = 0 To cmbComboName.ListCount - 1                     'go through the combobox entries
5636           iStart = InStr(1, cmbComboName.List(I), "[")
5637           iStop = InStr(1, cmbComboName.List(I), "]")
5638           strCompare = Mid$(cmbComboName.List(I), iStart + 1, iStop - iStart - 1)
5639           If UCase(strCompare) = UCase(ltr) Then   'see when we find the desired index number
5640               cmbComboName.ListIndex = I                                              'if we do, set the combo box
5641               Exit For                                            'and we're done
5642           End If
       '        cmbComboName.ListIndex = -1                             'else, remove any pointer
5643           cmbComboName.ListIndex = cmbComboName.ListCount - 1                           'else, remove any pointer
5644       Next I

5645       txtModelNo.Text = UCase(txtModelNo.Text)
5646       txtModelNo.SelStart = Len(txtModelNo.Text)
' <VB WATCH>
5647       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5648       Exit Function
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
5649       On Error GoTo vbwErrHandler
5650       Const VBWPROCNAME = "frmPLCData.LoadCombo"
5651       If vbwProtector.vbwTraceProc Then
5652           Dim vbwProtectorParameterString As String
5653           If vbwProtector.vbwTraceParameters Then
5654               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
5655               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sTableName", sTableName) & ") "
5656           End If
5657           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5658       End If
' </VB WATCH>

5659       Dim I As Integer
5660       Dim sItem As String
5661       Dim iID As Integer
5662       Dim bUseDropdown As Boolean
5663       Dim qy As New ADODB.Command
5664       Dim rs As New ADODB.Recordset

       '    rsPumpParameters.CursorLocation = adUseClient
       '    If sTableName = "Model" Then
       '        rsPumpParameters.Sort = "Model"
       '    Else
       '        rsPumpParameters.Sort = vbNullString
       '    End If
       '    rsPumpParameters.Open sTableName, cnPumpData, adOpenStatic, adLockOptimistic, adCmdTableDirect

5665       qy.ActiveConnection = cnPumpData
5666       If sTableName = "DischargeDiameter" Or sTableName = "SuctionDiameter" Then
5667           qy.CommandText = "SELECT * FROM " & sTableName & " ORDER BY Val(Description)"
5668       Else
5669           qy.CommandText = "SELECT * FROM " & sTableName & " ORDER BY Description"
5670       End If
5671       If sTableName = "SupermarketPumpData" Then
5672           qy.CommandText = "SELECT ID,Model AS Description FROM " & sTableName
5673       End If
5674       rs.CursorLocation = adUseClient
5675       rs.CursorType = adOpenStatic

5676       rs.Open qy


5677       On Error GoTo NoField
5678       bUseDropdown = True
           'sItem = rsPumpParameters.Fields("UseInDropdown")
       '    If bUseDropdown Then
       '        rsPumpParameters.Sort = "Description"
       '    End If
5679       rs.MoveFirst                                'goto the top
5680       For I = 0 To rs.RecordCount - 1             'go through the whole recordset
5681           sItem = rs.Fields("Description")        'get the description
5682           iID = rs.Fields(0)                      'get the index number - primary key
5683           If bUseDropdown Then
       '            If rsPumpParameters.Fields("UseInDropdown").value = True Then
5684                   cmbComboName.AddItem sItem, I                                   'add the description to the combo box
       '                cmbComboName.AddItem sItem                                   'add the description to the combo box
5685                   cmbComboName.ItemData(cmbComboName.NewIndex) = iID              'add the key number into the item data
       '            End If
5686           End If
5687           rs.MoveNext                             'get the next record
5688       Next I
5689       rs.Close
5690       cmbComboName.ListIndex = -1
5691   On Error GoTo vbwErrHandler
5692       Set rs = Nothing
5693       Set qy = Nothing
' <VB WATCH>
5694       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5695       Exit Function

5696   NoField:
5697       bUseDropdown = False
5698   On Error GoTo vbwErrHandler
5699       Resume Next

' <VB WATCH>
5700       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5701       Exit Function
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
5702       On Error GoTo vbwErrHandler
5703       Const VBWPROCNAME = "frmPLCData.SetGraphMax"
5704       If vbwProtector.vbwTraceProc Then
5705           Dim vbwProtectorParameterString As String
5706           If vbwProtector.vbwTraceParameters Then
5707               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Plothead", Plothead) & ") "
5708           End If
5709           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5710       End If
' </VB WATCH>

5711       Dim I As Integer
5712       Dim m As Single

5713       m = 0
5714       For I = 0 To UBound(Plothead, 2)
5715           If Plothead(1, I) > m Then
5716               m = Plothead(1, I)
5717           End If
5718       Next I
5719       SetGraphMax = 10 * (Int((m / 10) + 0.5) + 1)
5720       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
5721       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10 * (Int((m / 10) + 0.5) + 1)
5722       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0

' <VB WATCH>
5723       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5724       Exit Function
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
5725       On Error GoTo vbwErrHandler
5726       Const VBWPROCNAME = "frmPLCData.CalculateSpeed"
5727       If vbwProtector.vbwTraceProc Then
5728           Dim vbwProtectorParameterString As String
5729           If vbwProtector.vbwTraceParameters Then
5730               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("CoefSq", CoefSq) & ", "
5731               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("CoefLin", CoefLin) & ", "
5732               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("CoefConstant", CoefConstant) & ", "
5733               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("InputHP", InputHP) & ", "
5734               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("SG", SG) & ") "
5735           End If
5736           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5737       End If
' </VB WATCH>
5738       Dim I As Integer
5739       Dim OldResult As Double
5740       Dim NewResult As Double

5741       CalculateSpeed = 0

5742       If SG > 5 Or SG < 0.01 Then
5743           MsgBox "Bad value for SG...must be between 0.01 and 5.", vbOKOnly, "Bad SG Value"
' <VB WATCH>
5744       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5745           Exit Function
5746       End If

5747       OldResult = 1000
5748       NewResult = 0

5749       I = 1

5750       Do While Abs(NewResult - OldResult) > 0.1
5751           ReDim Preserve results(I)
5752           Select Case I
                   Case 1
5753                   results(I - 1).HP = InputHP
5754               Case 2
5755                   results(I - 1).HP = results(I - 2).HP * SG
5756               Case Else
5757                   results(I - 1).HP = results(I - 2).HP * (results(I - 2).Speed / results(I - 3).Speed) ^ 3
5758           End Select
5759           OldResult = NewResult
5760           results(I - 1).Speed = CalcPoly(CoefSq, CoefLin, CoefConstant, results(I - 1).HP)
5761           NewResult = results(I - 1).Speed
5762           If I > 15 Then
5763               If I = 0 Or I > 15 Then
5764                   MsgBox "Over 15 calculations and no convergence", vbOKOnly, "Too many iterations"
' <VB WATCH>
5765       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5766                   Exit Function
5767               End If
' <VB WATCH>
5768       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5769               Exit Function
5770           End If
5771           I = I + 1
5772       Loop
5773       CalculateSpeed = I - 1
' <VB WATCH>
5774       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5775       Exit Function
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
5776       On Error GoTo vbwErrHandler
5777       Const VBWPROCNAME = "frmPLCData.CalcPoly"
5778       If vbwProtector.vbwTraceProc Then
5779           Dim vbwProtectorParameterString As String
5780           If vbwProtector.vbwTraceParameters Then
5781               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("CoefSq", CoefSq) & ", "
5782               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("CoefLin", CoefLin) & ", "
5783               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("CoefConstant", CoefConstant) & ", "
5784               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("DataIn", DataIn) & ") "
5785           End If
5786           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5787       End If
' </VB WATCH>
5788       CalcPoly = CoefSq * DataIn ^ 2 + CoefLin * DataIn + CoefConstant
' <VB WATCH>
5789       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5790       Exit Function
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
5791       On Error GoTo vbwErrHandler
5792       Const VBWPROCNAME = "frmPLCData.GetBalanceHoleData"
5793       If vbwProtector.vbwTraceProc Then
5794           Dim vbwProtectorParameterString As String
5795           If vbwProtector.vbwTraceParameters Then
5796               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("SerialNumber", SerialNumber) & ", "
5797               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("TestDate", TestDate) & ") "
5798           End If
5799           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5800       End If
' </VB WATCH>
5801       If rsBalanceHoles.State = adStateOpen Then
5802           rsBalanceHoles.Close
5803       End If
5804       qyBalanceHoles.CommandText = "SELECT BalanceHoles.*, " & _
                             "IIf([Diameter]=99, 'Slot', [diameter]) as Diameter1, IIf([BoltCircle]=99, 'Unknown', [BoltCircle]) as BoltCircle1 " & _
                             "FROM BalanceHoles " & _
                             "WHERE [SerialNo] = '" & SerialNumber & "' AND [Date] <= #" & TestDate & "# " & _
                             "ORDER BY [Date], Val([BoltCircle]);"

5805       rsBalanceHoles.Open qyBalanceHoles
5806       rsBalanceHoles.Filter = ""

5807       Set dgBalanceHoles.DataSource = rsBalanceHoles

5808       Dim c As Column
5809       For Each c In dgBalanceHoles.Columns
5810           Select Case c.DataField
               Case "BalanceHoleID"
5811               c.Visible = False
5812           Case "SerialNo"
5813               c.Visible = False
5814           Case "Date"
5815               c.Visible = True
5816               c.Alignment = dbgCenter
5817               c.Width = 2000
5818           Case "Number"
5819               c.Visible = True
5820               c.Alignment = dbgCenter
5821               c.Width = 700
5822           Case "Diameter"
5823               c.Visible = False
5824           Case "Diameter1"
5825               c.Caption = "Diameter"
5826               c.Visible = True
5827               c.Alignment = dbgCenter
5828               c.Width = 700
5829           Case "BoltCircle1"
5830               c.Caption = "Bolt Circle"
5831               c.Visible = True
5832               c.Alignment = dbgCenter
5833               c.Width = 800
5834           Case "BoltCircle"
5835               c.Visible = False
5836           Case Else ' hide all other columns.
5837               c.Visible = False
5838           End Select
5839       Next c

' <VB WATCH>
5840       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5841       Exit Sub
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
5842       On Error GoTo vbwErrHandler
5843       Const VBWPROCNAME = "frmPLCData.FixPointsToPlot"
5844       If vbwProtector.vbwTraceProc Then
5845           Dim vbwProtectorParameterString As String
5846           If vbwProtector.vbwTraceParameters Then
5847               vbwProtectorParameterString = "()"
5848           End If
5849           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5850       End If
' </VB WATCH>
5851       If DataGrid2.Row = -1 Then
' <VB WATCH>
5852       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5853           Exit Sub
5854       End If
5855       Dim PresentGridRow As Integer
5856       PresentGridRow = DataGrid2.Row
5857       Dim GridIndex As Integer
5858       UpDown2.value = 8
5859       If DataGrid2.Row <> -1 Then
5860           For GridIndex = 0 To 7
5861               DataGrid2.Row = GridIndex
5862               If DataGrid2.Columns("Flow") = 0 And DataGrid2.Columns("TDH") = 0 Then
5863                   txtUpDn2.Text = GridIndex
5864                   UpDown2.value = GridIndex
' <VB WATCH>
5865       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5866                   Exit Sub
5867               End If
5868           Next GridIndex
5869       End If
5870       DataGrid2.Row = PresentGridRow
' <VB WATCH>
5871       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5872       Exit Sub
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
