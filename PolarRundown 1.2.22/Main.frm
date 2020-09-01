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
      Tab             =   1
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
      Tab(1).ControlEnabled=   -1  'True
      Tab(1).Control(0)=   "lbltab2(0)"
      Tab(1).Control(0).Enabled=   0   'False
      Tab(1).Control(1)=   "lbltab2(1)"
      Tab(1).Control(1).Enabled=   0   'False
      Tab(1).Control(2)=   "lbltab2(65)"
      Tab(1).Control(2).Enabled=   0   'False
      Tab(1).Control(3)=   "lbltab2(88)"
      Tab(1).Control(3).Enabled=   0   'False
      Tab(1).Control(4)=   "cmbTestSpec"
      Tab(1).Control(4).Enabled=   0   'False
      Tab(1).Control(5)=   "cmdEnterTestSetupData"
      Tab(1).Control(5).Enabled=   0   'False
      Tab(1).Control(6)=   "txtWho"
      Tab(1).Control(6).Enabled=   0   'False
      Tab(1).Control(7)=   "cmdAddNewTestDate"
      Tab(1).Control(7).Enabled=   0   'False
      Tab(1).Control(8)=   "txtTestSetupRemarks"
      Tab(1).Control(8).Enabled=   0   'False
      Tab(1).Control(9)=   "frmInstrumentTags"
      Tab(1).Control(9).Enabled=   0   'False
      Tab(1).Control(10)=   "frmLoopAndXducer"
      Tab(1).Control(10).Enabled=   0   'False
      Tab(1).Control(11)=   "frmElecData"
      Tab(1).Control(11).Enabled=   0   'False
      Tab(1).Control(12)=   "frmThrustBalMods"
      Tab(1).Control(12).Enabled=   0   'False
      Tab(1).Control(13)=   "frmPerfMods"
      Tab(1).Control(13).Enabled=   0   'False
      Tab(1).Control(14)=   "frmOtherFiles"
      Tab(1).Control(14).Enabled=   0   'False
      Tab(1).Control(15)=   "CommonDialog1"
      Tab(1).Control(15).Enabled=   0   'False
      Tab(1).Control(16)=   "cmdDeleteTestDate"
      Tab(1).Control(16).Enabled=   0   'False
      Tab(1).Control(17)=   "cmdApproveTestDate"
      Tab(1).Control(17).Enabled=   0   'False
      Tab(1).Control(18)=   "frmTAndI"
      Tab(1).Control(18).Enabled=   0   'False
      Tab(1).Control(19)=   "Command1"
      Tab(1).Control(19).Enabled=   0   'False
      Tab(1).Control(20)=   "txtRMA"
      Tab(1).Control(20).Enabled=   0   'False
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
         TabIndex        =   427
         Text            =   "8"
         Top             =   5520
         Width           =   285
      End
      Begin MSChart20Lib.MSChart MSChart1 
         Height          =   2772
         Left            =   -68040
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
         Left            =   -74040
         TabIndex        =   425
         Text            =   "1"
         Top             =   8880
         Width           =   285
      End
      Begin MSComCtl2.UpDown UpDown1 
         Height          =   504
         Left            =   -74280
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
         Left            =   -63720
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
         Left            =   -61920
         Style           =   1  'Graphical
         TabIndex        =   390
         Top             =   2400
         Width           =   1332
      End
      Begin VB.TextBox txtRMA 
         Height          =   315
         Left            =   5040
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
         Left            =   7680
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
         Left            =   8520
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
         Left            =   -74880
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
         Left            =   12240
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
         Left            =   10560
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
         Left            =   14040
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
         Left            =   5160
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
         Left            =   5160
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
         Left            =   120
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
         Left            =   5160
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
         Left            =   120
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
         Left            =   8520
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
         Left            =   2640
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
         Left            =   -61920
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
         Left            =   -74880
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
         Left            =   -68040
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
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   41
         Top             =   660
         Width           =   2055
      End
      Begin VB.TextBox txtWho 
         Height          =   315
         Left            =   1920
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
         Left            =   9000
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
         Left            =   -74760
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
         Left            =   -74760
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
         Height          =   315
         Left            =   1920
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
            TabIndex        =   274
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
         Left            =   -61080
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
         Left            =   4080
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
         Left            =   -62160
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
         Left            =   -61920
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
         Left            =   600
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
         Left            =   -61920
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
         Left            =   240
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
         Left            =   -74520
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
         Left            =   120
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

'v1.2.21 - MHR - 3/26/18
'   Fixed FixPointsToPlot routine to prevent setting number of points to 0
'   Fixed cmbCirculationFlowMeter to store proper data

'v1.2.2 - MHR - 5/11/18
'   Added figure to transducer diagram
'   changed A2 and A3 default to Man from Auto
'   Default Frequency to 60Hz

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

837        SetFrequencyCombo

838        EnableTestSetupDataControls
839        Pressed = False
840        cmdEnterTestSetupData_Click
841        cmdAddNewBalanceHoles.Visible = True
842        txtWho.Text = LogInInitials
843        MsgBox "New Test Date Added - " & cmbTestDate.List(cmbTestDate.ListCount - 1), vbOKOnly, "Added New Test Date"
' <VB WATCH>
844        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
845        Exit Sub
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
846        On Error GoTo vbwErrHandler
847        Const VBWPROCNAME = "frmPLCData.cmdApprovePump_Click"
848        If vbwProtector.vbwTraceProc Then
849            Dim vbwProtectorParameterString As String
850            If vbwProtector.vbwTraceParameters Then
851                vbwProtectorParameterString = "()"
852            End If
853            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
854        End If
' </VB WATCH>
855        rsPumpData.Fields("Approved") = Not rsPumpData.Fields("Approved")
856        rsPumpData.Update
857        rsPumpData.Requery
858        lblPumpApproved.Visible = rsPumpData.Fields("Approved")
859        If rsPumpData.Fields("Approved") = True Then
860            cmdApprovePump.Caption = "Unapprove This Pump"
861            cmdApproveTestDate.Enabled = True
862            If rsTestSetup.Fields("Approved") = True Then
863                cmdApproveTestDate.Caption = "Unapprove This Test Date"
864            Else
865                cmdApproveTestDate.Caption = "Approve This Test Date"
866            End If
867        Else
868            cmdApprovePump.Caption = "Approve This Pump"
869            cmdApproveTestDate.Caption = "You Must Approve Pump First"
870            cmdApproveTestDate.Enabled = False
871        End If
' <VB WATCH>
872        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
873        Exit Sub
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
874        On Error GoTo vbwErrHandler
875        Const VBWPROCNAME = "frmPLCData.cmdApproveTestDate_Click"
876        If vbwProtector.vbwTraceProc Then
877            Dim vbwProtectorParameterString As String
878            If vbwProtector.vbwTraceParameters Then
879                vbwProtectorParameterString = "()"
880            End If
881            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
882        End If
' </VB WATCH>
883        rsTestSetup.Fields("Approved") = Not rsTestSetup.Fields("Approved")
884        rsTestSetup.Update
885        rsTestSetup.Requery
886        lblTestDateApproved.Visible = rsTestSetup.Fields("Approved")
887        If rsTestSetup.Fields("Approved") = True Then
888            cmdApproveTestDate.Caption = "Unapprove This Test Date"
889        Else
890            cmdApproveTestDate.Caption = "Approve This Test Date"
891        End If
' <VB WATCH>
892        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
893        Exit Sub
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
894        On Error GoTo vbwErrHandler
895        Const VBWPROCNAME = "frmPLCData.cmdCalibrate_Click"
896        If vbwProtector.vbwTraceProc Then
897            Dim vbwProtectorParameterString As String
898            If vbwProtector.vbwTraceParameters Then
899                vbwProtectorParameterString = "()"
900            End If
901            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
902        End If
' </VB WATCH>
903        Dim ans As Integer
904        Dim I As Integer

905        ans = MsgBox("You have selected to calibrate the software.  Do you want to continue?", vbYesNo, "Calibrate Software")
906        If ans = vbNo Then
907            Calibrating = False
' <VB WATCH>
908        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
909            Exit Sub
910        Else
911            CalibrateSoftware
912        End If
' <VB WATCH>
913        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
914        Exit Sub
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
915        On Error GoTo vbwErrHandler
916        Const VBWPROCNAME = "frmPLCData.cmdClearPumpData_Click"
917        If vbwProtector.vbwTraceProc Then
918            Dim vbwProtectorParameterString As String
919            If vbwProtector.vbwTraceParameters Then
920                vbwProtectorParameterString = "()"
921            End If
922            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
923        End If
' </VB WATCH>
924        BlankData
' <VB WATCH>
925        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
926        Exit Sub
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
927        On Error GoTo vbwErrHandler
928        Const VBWPROCNAME = "frmPLCData.cmdDeletePump_Click"
929        If vbwProtector.vbwTraceProc Then
930            Dim vbwProtectorParameterString As String
931            If vbwProtector.vbwTraceParameters Then
932                vbwProtectorParameterString = "()"
933            End If
934            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
935        End If
' </VB WATCH>
936        Dim Answer As Integer
937        Answer = MsgBox("You are about to delete the following record: S/N = " & rsPumpData.Fields("SerialNumber") & "!  Do you want to continue?", vbCritical Or vbYesNo, "Ready to Delete")
938        If Answer = vbYes Then
939            rsPumpData.Delete
940            rsPumpData.Update
941            cmdFindPump_Click
942        End If
' <VB WATCH>
943        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
944        Exit Sub
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
945        On Error GoTo vbwErrHandler
946        Const VBWPROCNAME = "frmPLCData.cmdDeleteTestDate_Click"
947        If vbwProtector.vbwTraceProc Then
948            Dim vbwProtectorParameterString As String
949            If vbwProtector.vbwTraceParameters Then
950                vbwProtectorParameterString = "()"
951            End If
952            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
953        End If
' </VB WATCH>
954        Dim Answer As Integer
955        Answer = MsgBox("You are about to delete the following record: S/N = " & rsTestData.Fields("SerialNumber") & " and Test Date = " & rsTestSetup.Fields("Date") & "!  Do you want to continue?", vbCritical Or vbYesNo, "Ready to Delete")
956        If Answer = vbYes Then
957            rsTestSetup.Delete
958            rsTestSetup.Update
959            cmdFindPump_Click
960        End If
' <VB WATCH>
961        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
962        Exit Sub
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
963        On Error GoTo vbwErrHandler
964        Const VBWPROCNAME = "frmPLCData.cmdEnterPumpData_Click"
965        If vbwProtector.vbwTraceProc Then
966            Dim vbwProtectorParameterString As String
967            If vbwProtector.vbwTraceParameters Then
968                vbwProtectorParameterString = "()"
969            End If
970            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
971        End If
' </VB WATCH>
972        Dim d As Integer
973        Dim sSearch As String
974        Dim ans As Integer
975        Dim boWriteDataWritten As Boolean


           'check for a serial number
976        If LenB(txtSN.Text) = 0 Then
977            MsgBox "You must have a Serial Number to enter data.  Data has not been saved."
' <VB WATCH>
978        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
979            Exit Sub
980        End If

           'check to make sure most entries are filled in
981        If LenB(txtModelNo.Text) = 0 And optMfr(0).value = True Then
982            MsgBox "You need to enter a MODEL NO before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
983        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
984            Exit Sub
985        End If
986        If LenB(txtSalesOrderNumber.Text) = 0 Then
987            If InStr(1, txtSN.Text, "-") <> 0 Then
988                txtSalesOrderNumber.Text = Mid$(txtSN.Text, 1, InStr(1, txtSN.Text, "-") - 1)
989            End If
990        End If
991        If LenB(txtSalesOrderNumber.Text) = 0 Then
992            MsgBox "You need to enter a SALES ORDER NUMBER before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
993        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
994            Exit Sub
995        End If

996        If cmbMotor.ListIndex = -1 And optMfr(0).value = True Then
997            MsgBox "You need to pick a MOTOR before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
998        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
999            Exit Sub
1000       End If

1001       If cmbStatorFill.ListIndex = -1 And optMfr(0).value = True Then    'set default
1002           cmbStatorFill.ListIndex = 0
1003       End If

1004       If cmbModel.ListIndex = -1 And optMfr(0).value = True Then
1005           MsgBox "You need to pick a MODEL before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1006       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1007           Exit Sub
1008       End If

1009       If cmbModelGroup.ListIndex = -1 And optMfr(0).value = True Then
1010           MsgBox "You need to pick a MODEL GROUP before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1011       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1012           Exit Sub
1013       End If


1014       If cmbDesignPressure.ListIndex = -1 And optMfr(0).value = True Then
1015           MsgBox "You need to pick a DESIGN PRESSURE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1016       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1017           Exit Sub
1018       End If

1019       If cmbCirculationPath.ListIndex = -1 And optMfr(0).value = True Then
1020           MsgBox "You need to pick a CIRCULATION PATH before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1021       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1022           Exit Sub
1023       End If

1024       If cmbRPM.ListIndex = -1 And optMfr(0).value = True Then
1025           MsgBox "You need to pick an RPM before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1026       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1027           Exit Sub
1028       End If

       'check TEMC dropdowns

1029       If cmbTEMCAdapter.ListIndex = -1 And optMfr(0).value = False Then
1030           MsgBox "You need to pick a TEMC ADAPTER before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1031       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1032           Exit Sub
1033       End If

1034       If cmbTEMCAdditions.ListIndex = -1 And optMfr(0).value = False Then
1035           MsgBox "You need to pick TEMC ADDITIONS before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1036       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1037           Exit Sub
1038       End If

1039       If cmbTEMCCirculation.ListIndex = -1 And optMfr(0).value = False Then
1040           MsgBox "You need to pick a TEMC CIRCULATION before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1041       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1042           Exit Sub
1043       End If

1044       If cmbTEMCDesignPressure.ListIndex = -1 And optMfr(0).value = False Then
1045           MsgBox "You need to pick a TEMC DESIGN PRESSURE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1046       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1047           Exit Sub
1048       End If

1049       If cmbTEMCDivisionType.ListIndex = -1 And optMfr(0).value = False Then
1050           MsgBox "You need to pick a TEMC DIVISION TYPE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1051       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1052           Exit Sub
1053       End If

1054       If cmbTEMCImpellerType.ListIndex = -1 And optMfr(0).value = False Then
1055           MsgBox "You need to pick a TEMC IMPELLER TYPE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1056       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1057           Exit Sub
1058       End If

1059       If cmbTEMCInsulation.ListIndex = -1 And optMfr(0).value = False Then
1060           MsgBox "You need to pick a TEMC INSULATION TYPE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1061       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1062           Exit Sub
1063       End If

1064       If cmbTEMCJacketGasket.ListIndex = -1 And optMfr(0).value = False Then
1065           MsgBox "You need to pick a TEMC JACKET GASKET before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1066       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1067           Exit Sub
1068       End If

1069       If cmbTEMCMaterials.ListIndex = -1 And optMfr(0).value = False Then
1070           MsgBox "You need to pick TEMC MATERIALS before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1071       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1072           Exit Sub
1073       End If

1074       If cmbTEMCModel.ListIndex = -1 And optMfr(0).value = False Then
1075           MsgBox "You need to pick a TEMC MODEL before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1076       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1077           Exit Sub
1078       End If

1079       If cmbTEMCNominalImpSize.ListIndex = -1 And optMfr(0).value = False Then
1080           MsgBox "You need to pick a TEMC NOMINAL IMPELLER SIZE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1081       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1082           Exit Sub
1083       End If

1084       If cmbTEMCNominalDischargeSize.ListIndex = -1 And optMfr(0).value = False Then
1085           MsgBox "You need to pick a TEMC NOMINAL DISCHARGE SIZE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1086       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1087           Exit Sub
1088       End If

1089       If cmbTEMCNominalSuctionSize.ListIndex = -1 And optMfr(0).value = False Then
1090           MsgBox "You need to pick a TEMC NOMINAL SUCTION SIZE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1091       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1092           Exit Sub
1093       End If

1094       If cmbTEMCOtherMotor.ListIndex = -1 And optMfr(0).value = False Then
1095           MsgBox "You need to pick a TEMC OTHER MOTOR before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1096       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1097           Exit Sub
1098       End If

1099       If cmbTEMCPumpStages.ListIndex = -1 And optMfr(0).value = False Then
1100           MsgBox "You need to pick TEMC NUMBER OF PUMP STAGES before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1101       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1102           Exit Sub
1103       End If

1104       If cmbTEMCTRG.ListIndex = -1 And optMfr(0).value = False Then
1105           MsgBox "You need to pick a TEMC TRG before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1106       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1107           Exit Sub
1108       End If

1109       If cmbTEMCVoltage.ListIndex = -1 And optMfr(0).value = False Then
1110           MsgBox "You need to pick a TEMC VOLTAGE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1111       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1112           Exit Sub
1113       End If

1114       If LenB(txtTEMCFrameNumber.Text) = 0 And optMfr(0).value = False Then
1115           MsgBox "You need to enter a TEMC FRAME NUMBER before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1116       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1117           Exit Sub
1118       End If


1119       If Not boFoundPump Then     'if we havent found a pump in the database, add it
1120           rsPumpData.AddNew
1121           boWriteDataWritten = False
1122       Else    'else, find the entry
1123           sSearch = "Serialnumber = '" & frmPLCData.txtSN.Text & "'"
1124           rsPumpData.MoveFirst
1125           rsPumpData.Find sSearch, , adSearchForward, 1
1126           boWriteDataWritten = True
1127       End If

1128       If Not IsNull(rsPumpData!DataWritten) Or rsPumpData!DataWritten = True Then
1129           ans = MsgBox("You have already entered data for this pump.  Do you want to overwrite the data?", vbDefaultButton2 + vbYesNo, "Overwrite Data?")
1130           If ans = vbNo Then
1131               rsPumpData!DataWritten = True
1132               rsPumpData.Update   'update datawritten
' <VB WATCH>
1133       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1134               Exit Sub
1135           End If
1136       End If

1137       rsPumpData!SerialNumber = frmPLCData.txtSN.Text
1138       rsPumpData!ModelNumber = frmPLCData.txtModelNo.Text
1139       rsPumpData!SalesOrderNumber = frmPLCData.txtSalesOrderNumber.Text
1140       rsPumpData!ShipToCustomer = frmPLCData.txtShpNo.Text
1141       rsPumpData!BillToCustomer = frmPLCData.txtBilNo.Text
1142       rsPumpData!ApplicationFluid = frmPLCData.txtLiquid
1143       rsPumpData!NPSHFile = frmPLCData.txtNPSHFileLocation.Text
1144       rsPumpData!RVSPartNo = frmPLCData.txtRVSPartNo.Text
1145       rsPumpData!CustPN = frmPLCData.txtXPartNum.Text
1146       rsPumpData!CustPO = frmPLCData.txtCustPONum.Text

1147       If Len(frmPLCData.txtViscosity) <> 0 Then
1148           rsPumpData!ApplicationViscosity = frmPLCData.txtViscosity
1149       End If

1150       If frmPLCData.chkSuperMarketFeathered.value = Checked Then
1151           rsPumpData!Field1 = "Feathered"
1152       Else
1153           rsPumpData!Field1 = ""
1154       End If

1155       If LenB(txtSpGr.Text) <> 0 Then
1156           If Not IsNumeric(frmPLCData.txtSpGr.Text) Then
1157               MsgBox "Specific Gravity must be a number."
' <VB WATCH>
1158       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1159               Exit Sub
1160           End If
1161           rsPumpData!SpGr = frmPLCData.txtSpGr.Text
1162       End If
1163       If LenB(txtImpellerDia.Text) <> 0 Then
1164           If Not IsNumeric(frmPLCData.txtImpellerDia.Text) Then
1165               MsgBox "Impeller Diameter must be a number."
' <VB WATCH>
1166       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1167               Exit Sub
1168           End If
1169           rsPumpData!impellerdia = frmPLCData.txtImpellerDia.Text
1170       End If
1171       If LenB(txtDesignFlow.Text) <> 0 Then
1172           rsPumpData!designflow = frmPLCData.txtDesignFlow.Text
1173       End If
1174       If LenB(txtDesignTDH.Text) <> 0 Then
1175           rsPumpData!designtdh = frmPLCData.txtDesignTDH.Text
1176       End If
1177       If LenB(txtRemarks.Text) <> 0 Then
1178           rsPumpData!Remarks = txtRemarks.Text
1179       End If

1180       If optMfr(0).value = True Then
1181           d = cmbMotor.ItemData(cmbMotor.ListIndex)
1182           rsPumpData!Motor = d
1183           d = cmbStatorFill.ItemData(cmbStatorFill.ListIndex)
1184           rsPumpData!StatorFill = d
1185            d = cmbDesignPressure.ItemData(cmbDesignPressure.ListIndex)
1186           rsPumpData!DesignPressure = d
1187           d = cmbCirculationPath.ItemData(cmbCirculationPath.ListIndex)
1188           rsPumpData!CirculationPath = d
1189           d = cmbRPM.ItemData(cmbRPM.ListIndex)
1190           rsPumpData!RPM = d
1191           d = cmbModel.ItemData(cmbModel.ListIndex)
1192           rsPumpData!Model = d
1193           d = cmbModelGroup.ItemData(cmbModelGroup.ListIndex)
1194           rsPumpData!ModelGroup = d
1195       End If
       '   TEMC fields
1196       If optMfr(0).value = False Then
1197           d = cmbTEMCAdapter.ItemData(cmbTEMCAdapter.ListIndex)
1198           rsPumpData!TEMCAdapter = d

1199           d = cmbTEMCAdditions.ItemData(cmbTEMCAdditions.ListIndex)
1200           rsPumpData!TEMCAdditions = d

1201           d = cmbTEMCCirculation.ItemData(cmbTEMCCirculation.ListIndex)
1202           rsPumpData!TEMCcirculation = d

1203           d = cmbTEMCDesignPressure.ItemData(cmbTEMCDesignPressure.ListIndex)
1204           rsPumpData!TEMCDesignpressure = d

1205           d = cmbTEMCDivisionType.ItemData(cmbTEMCDivisionType.ListIndex)
1206           rsPumpData!TEMCDivisionType = d

1207           d = cmbTEMCImpellerType.ItemData(cmbTEMCImpellerType.ListIndex)
1208           rsPumpData!TEMCImpellerType = d

1209           d = cmbTEMCInsulation.ItemData(cmbTEMCInsulation.ListIndex)
1210           rsPumpData!TEMCInsulation = d

1211           d = cmbTEMCJacketGasket.ItemData(cmbTEMCJacketGasket.ListIndex)
1212           rsPumpData!TEMCJacketGasket = d

1213           d = cmbTEMCMaterials.ItemData(cmbTEMCMaterials.ListIndex)
1214           rsPumpData!TEMCMaterials = d

1215           d = cmbTEMCModel.ItemData(cmbTEMCModel.ListIndex)
1216           rsPumpData!TEMCModel = d

1217           d = cmbTEMCNominalImpSize.ItemData(cmbTEMCNominalImpSize.ListIndex)
1218           rsPumpData!TEMCNominalImpSize = d

1219           d = cmbTEMCNominalDischargeSize.ItemData(cmbTEMCNominalDischargeSize.ListIndex)
1220           rsPumpData!TEMCNominalDischargeSize = d

1221           d = cmbTEMCNominalSuctionSize.ItemData(cmbTEMCNominalSuctionSize.ListIndex)
1222           rsPumpData!TEMCNominalSuctionSize = d

1223           d = cmbTEMCOtherMotor.ItemData(cmbTEMCOtherMotor.ListIndex)
1224           rsPumpData!TEMCOtherMotor = d

1225           d = cmbTEMCPumpStages.ItemData(cmbTEMCPumpStages.ListIndex)
1226           rsPumpData!TEMCPumpStages = d

1227           d = cmbTEMCTRG.ItemData(cmbTEMCTRG.ListIndex)
1228           rsPumpData!TEMCTRG = d

1229           d = cmbTEMCVoltage.ItemData(cmbTEMCVoltage.ListIndex)
1230           rsPumpData!TEMCVoltage = d

1231           If LenB(txtTEMCFrameNumber.Text) <> 0 Then
1232               rsPumpData!TEMCFrameNumber = frmPLCData.txtTEMCFrameNumber.Text
1233           End If
1234       End If

1235       rsPumpData!ChempumpPump = optMfr(0).value

1236       rsPumpData!Approved = False

       'added from TEMC Inspection Report
1237       If Len(txtJobNum.Text) <> 0 Then
1238           rsPumpData!JobNumber = txtJobNum.Text
1239       End If

1240       If Len(txtNoPhases.Text) <> 0 Then
1241           rsPumpData!Phases = txtNoPhases.Text
1242       End If

1243       If Len(txtExpClass.Text) <> 0 Then
1244           rsPumpData!ExpClass = txtExpClass.Text
1245       End If

1246       If Len(txtThermalClass.Text) <> 0 Then
1247           rsPumpData!ThermalClass = txtThermalClass.Text
1248       End If

1249       rsPumpData!NPSHr = Val(txtNPSHr.Text)
1250       rsPumpData!LiquidTemperature = Val(txtLiquidTemperature.Text)
1251       rsPumpData!RatedInputPower = Val(txtRatedInputPower.Text)
1252       rsPumpData!FLCurrent = Val(txtAmps.Text)





1253       If boWriteDataWritten Then
1254           rsPumpData!DataWritten = True
1255       Else
1256           rsPumpData!DataWritten = False
1257       End If

           'write the data into the database
1258       rsPumpData.Update
1259       boFoundPump = True

           'enter a new test date if it's a new entry
1260       If Not boWriteDataWritten Then


1261           cmdAddNewTestDate_Click
1262       End If
' <VB WATCH>
1263       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1264       Exit Sub
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
1265       On Error GoTo vbwErrHandler
1266       Const VBWPROCNAME = "frmPLCData.cmdEnterTestData_Click"
1267       If vbwProtector.vbwTraceProc Then
1268           Dim vbwProtectorParameterString As String
1269           If vbwProtector.vbwTraceParameters Then
1270               vbwProtectorParameterString = "()"
1271           End If
1272           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1273       End If
' </VB WATCH>
1274       Dim sSearch As String
1275       Dim ans As Integer

           'if we didn't find the test setup, can't enter test data
1276       If Not boFoundTestSetup Then
1277           MsgBox "You must enter Test Setup Data before entering the Test Data"
' <VB WATCH>
1278       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1279           Exit Sub
1280       End If

           'if we don't find data in the test database, add records
1281       If boFoundTestData = False Then     'add 8 records for 8 tests
1282           AddTestData
1283           rsTestData.MoveFirst
1284       Else        'find the data in the database
1285           sSearch = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"
1286           rsTestData.MoveFirst
1287           rsTestData.Filter = sSearch
1288       End If

           'find the desired record from the form
1289       rsTestData.MoveFirst
1290       rsTestData.Move UpDown1.value - 1

1291       If rsTestData!DataWritten = True Then
1292           ans = MsgBox("You have already entered data for this test.  Do you want to overwrite the data?", vbYesNo + vbDefaultButton2, "Data Already Entered")
1293           If ans = vbNo Then
' <VB WATCH>
1294       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1295               Exit Sub
1296           End If
1297       End If

1298       rsEff.MoveFirst
1299       rsEff.Move UpDown1.value - 1

1300       If LenB(txtV1.Text) <> 0 Then
1301           rsTestData!VoltageA = Val(txtV1.Text)
1302       End If

1303       If LenB(txtV2.Text) <> 0 Then
1304           rsTestData!VoltageB = Val(txtV2.Text)
1305       End If

1306       If LenB(txtV3.Text) <> 0 Then
1307           rsTestData!VoltageC = Val(txtV3.Text)
1308       End If

1309       If LenB(txtI1.Text) <> 0 Then
1310           rsTestData!CurrentA = Val(txtI1.Text)
1311       End If

1312       If LenB(txtI2.Text) <> 0 Then
1313           rsTestData!CurrentB = Val(txtI2.Text)
1314       End If

1315       If LenB(txtI3.Text) <> 0 Then
1316           rsTestData!CurrentC = Val(txtI3.Text)
1317       End If

1318       If LenB(txtP1.Text) <> 0 Then
1319           rsTestData!PowerA = Val(txtP1.Text)
1320       End If

1321       If LenB(txtP2.Text) <> 0 Then
1322           rsTestData!PowerB = Val(txtP2.Text)
1323       End If

1324       If LenB(txtP3.Text) <> 0 Then
1325           rsTestData!PowerC = Val(txtP3.Text)
1326       End If

1327       If LenB(txtKW.Text) <> 0 Then
1328           rsTestData!TotalPower = Val(txtKW.Text)
1329       End If

1330       rsTestData!Flow = Val(txtFlowDisplay.Text)
1331       rsTestData!DischargePressure = Val(txtDischargeDisplay.Text)
1332       rsTestData!SuctionPressure = Val(txtSuctionDisplay.Text)
1333       rsTestData!TemperatureSuction = Val(txtTemperatureDisplay.Text)

1334       rsTestData!TC1 = Val(txtTC1Display.Text)
1335       rsTestData!TC2 = Val(txtTC2Display.Text)
1336       rsTestData!TC3 = Val(txtTC3Display.Text)
1337       rsTestData!TC4 = Val(txtTC4Display.Text)

1338       rsTestData!CircFlow = Val(txtAI1Display.Text)
1339       rsTestData!RBHTemp = Val(txtAI2Display.Text)
1340       rsTestData!RBHPress = Val(txtAI3Display.Text)
1341       rsTestData!AI4 = Val(txtAI4Display.Text)

1342       rsTestData!ValvePosition = Val(txtValvePosition.Text)
1343       rsTestData!SetPoint = Val(txtSetPoint.Text)

1344       If LenB(txtThrustBal.Text) <> 0 Then
1345           rsTestData!ThrustBalance = txtThrustBal.Text
1346       End If

1347       If LenB(txtVibAx.Text) <> 0 Then
1348           rsTestData!VibrationX = txtVibAx.Text
1349       End If

1350       If LenB(txtVibRad.Text) <> 0 Then
1351           rsTestData!VibrationY = txtVibRad.Text
1352       End If

1353       If LenB(txtTEMCTRGReading.Text) <> 0 Then
1354           rsTestData!TEMCTRG = txtTEMCTRGReading.Text
1355       Else
1356           rsTestData!TEMCTRG = 0
1357       End If

1358       If LenB(txtRPM.Text) <> 0 Then
1359           rsTestData!RPM = txtRPM.Text
1360       End If

1361       If LenB(txtTestRemarks.Text) <> 0 Then
1362           rsTestData!Remarks = txtTestRemarks.Text
1363       Else
1364           rsTestData!Remarks = " "
1365       End If

1366       If LenB(txtTEMCTRGReading.Text) <> 0 Then
1367           rsTestData!TEMCTRG = txtTEMCTRGReading.Text
1368       End If

1369       If LenB(txtTEMCFrontThrust.Text) <> 0 Then
1370           rsTestData!TEMCFrontThrust = txtTEMCFrontThrust.Text
1371       End If

1372       If LenB(txtTEMCRearThrust.Text) <> 0 Then
1373           rsTestData!TEMCRearThrust = txtTEMCRearThrust.Text
1374       End If

1375       If LenB(txtTEMCMomentArm.Text) <> 0 Then
1376           rsTestData!TEMCMomentArm = txtTEMCMomentArm.Text
1377       End If

1378       If LenB(txtTEMCThrustRigPressure.Text) <> 0 Then
1379           rsTestData!TEMCThrustRigPressure = txtTEMCThrustRigPressure.Text
1380       End If

1381       If LenB(txtTEMCViscosity.Text) <> 0 Then
1382           rsTestData!TEMCViscosity = txtTEMCViscosity.Text
1383       End If

1384       If LenB(txtNPSHa.Text) <> 0 Then
1385           rsTestData!NPSHa = txtNPSHa.Text
1386       End If

1387       rsTestData!Approved = False

1388       rsTestData!DataWritten = True

           'update the database
1389       rsTestData.Update

1390       DoEfficiencyCalcs
1391       rsEff.Update

           'update the form
1392       DataGrid1.Refresh
1393       DataGrid2.Refresh

1394       FixPointsToPlot

' <VB WATCH>
1395       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1396       Exit Sub
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
1397       On Error GoTo vbwErrHandler
1398       Const VBWPROCNAME = "frmPLCData.cmdEnterTestSetupData_Click"
1399       If vbwProtector.vbwTraceProc Then
1400           Dim vbwProtectorParameterString As String
1401           If vbwProtector.vbwTraceParameters Then
1402               vbwProtectorParameterString = "()"
1403           End If
1404           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1405       End If
' </VB WATCH>
1406       Dim I As Integer
1407       Dim d As Integer
1408       Dim sSearch As String
1409       Dim ans As Integer
1410       Dim boWriteDataWritten As Boolean

           'check for a serial number
1411       If LenB(txtSN.Text) = 0 Then
1412           MsgBox "You must have a Serial Number to enter data.", vbOKOnly + vbExclamation, "Cannot Enter Data"
' <VB WATCH>
1413       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1414           Exit Sub
1415       End If

1416       If Pressed = True Then
1417           If Me.cmbDischDia.ListIndex = -1 Or Me.cmbSuctDia.ListIndex = -1 Or Val(Me.txtSuctHeight.Text) = 0 Or Val(Me.txtDischHeight.Text) = 0 Then
1418               MsgBox "You must have Discharge Diameter AND Suction Diameter AND Suction Height AND Discharge Height entered to enter data.", vbOKOnly + vbExclamation, "Cannot Enter Data"
' <VB WATCH>
1419       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1420               Exit Sub
1421           End If
1422       End If

1423       Pressed = True
1424       If Not boFoundTestSetup Then    'if we didn't find any test setup, add a record
1425           rsTestSetup.AddNew
1426           cmbTestDate.AddItem Now
1427           cmbTestDate.ListIndex = cmbTestDate.NewIndex
1428           cmdAddNewBalanceHoles.Visible = True
1429           boFoundTestSetup = True
1430           boWriteDataWritten = False
1431           rsTestSetup!DataWritten = False
1432       Else    'find the record and display
1433           sSearch = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"
1434           rsTestSetup.MoveFirst
1435           rsTestSetup.Filter = sSearch
1436           If Not boCanApprove Then
       '            cmdAddNewBalanceHoles.Visible = False
1437           End If
1438           boWriteDataWritten = True
1439       End If

1440       If rsTestSetup!DataWritten = True Then
1441           ans = MsgBox("Data has already been entered for this test date.  Do you want to overwrite it?", vbYesNo + vbDefaultButton2, "Data Exists")
1442           If ans = vbNo Then
' <VB WATCH>
1443       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1444               Exit Sub
1445           End If
1446       End If

1447       rsTestSetup!SerialNumber = txtSN
1448       rsTestSetup!Date = cmbTestDate.List(cmbTestDate.ListIndex)

1449       I = cmbFlowMeter.ListIndex
1450       If I = -1 Then
1451           d = 1
1452           rsTestSetup!FlowMeterID = d
1453       Else
1454           d = cmbLoopNumber.ItemData(I)
1455           rsTestSetup!FlowMeterID = d
1456       End If

1457       I = cmbSuctionPressureTransducer.ListIndex
1458       If I = -1 Then
1459           d = 1
1460           rsTestSetup!suctionid = d
1461       Else
1462           d = cmbLoopNumber.ItemData(I)
1463           rsTestSetup!suctionid = d
1464       End If

1465       I = cmbDischargePressureTransducer.ListIndex
1466       If I = -1 Then
1467           d = 1
1468           rsTestSetup!dischid = d
1469       Else
1470           d = cmbLoopNumber.ItemData(I)
1471           rsTestSetup!dischid = d
1472       End If

1473       I = cmbTemperatureTransducer.ListIndex
1474       If I = -1 Then
1475           d = 1
1476           rsTestSetup!temperatureid = d
1477       Else
1478           d = cmbLoopNumber.ItemData(I)
1479           rsTestSetup!temperatureid = d
1480       End If

1481       I = Me.cmbCirculationFlowMeter.ListIndex
1482       If I = -1 Or I > 4 Then
1483           d = 5
1484           rsTestSetup!MagFlowID = d
1485       Else
1486           d = cmbLoopNumber.ItemData(I) + 4
1487           rsTestSetup!MagFlowID = d
1488       End If


1489       If LenB(txtHDCor.Text) <> 0 Then
1490           rsTestSetup!HDCor = txtHDCor
1491       Else
1492           rsTestSetup!HDCor = 0
1493       End If
1494       If LenB(txtKWMult.Text) <> 0 Then
1495           rsTestSetup!kwmult = txtKWMult
1496       Else
1497           rsTestSetup!kwmult = 1
1498       End If
1499       If LenB(txtWho.Text) <> 0 Then
1500           rsTestSetup!who = txtWho
1501       Else
1502           rsTestSetup!who = vbNullString
1503       End If
1504       If LenB(txtRMA.Text) <> 0 Then
1505           rsTestSetup!RMA = txtRMA
1506       Else
1507           rsTestSetup!RMA = vbNullString
1508       End If
1509       If LenB(frmPLCData.txtDischHeight) <> 0 Then
1510           rsTestSetup!DischargeGageHeight = Val(txtDischHeight)
1511       Else
1512           rsTestSetup!DischargeGageHeight = 0
1513       End If
1514       If LenB(frmPLCData.txtSuctHeight) <> 0 Then
1515           rsTestSetup!SuctionGageHeight = Val(txtSuctHeight)
1516       Else
1517           rsTestSetup!SuctionGageHeight = 0
1518       End If
1519       If LenB(frmPLCData.txtTestSetupRemarks.Text) <> 0 Then
1520           rsTestSetup!Remarks = txtTestSetupRemarks.Text
1521       Else
1522           rsTestSetup!Remarks = vbNullString
1523       End If
1524       If LenB(frmPLCData.txtVFDFreq.Text) <> 0 Then
1525           rsTestSetup!VFDFrequency = txtVFDFreq.Text
1526       Else
1527           rsTestSetup!VFDFrequency = 0
1528       End If

1529       I = cmbOrificeNumber.ListIndex
1530       If I = -1 Then
1531           d = 18      'entry for None
1532       Else
1533           d = cmbOrificeNumber.ItemData(I)
1534       End If
1535       rsTestSetup!orificenumber = d

1536       If LenB(txtEndPlay.Text) <> 0 Then
1537           rsTestSetup!Endplay = Val(frmPLCData.txtEndPlay.Text)
1538       Else
1539           rsTestSetup!Endplay = 0
1540       End If

1541       If LenB(txtGGap.Text) <> 0 Then
1542           rsTestSetup!GGAP = Val(frmPLCData.txtGGap.Text)
1543       Else
1544           rsTestSetup!GGAP = 0
1545       End If

1546       If LenB(txtOtherMods.Text) <> 0 Then
1547           rsTestSetup!OtherMods = txtOtherMods.Text
1548       Else
1549           rsTestSetup!OtherMods = vbNullString
1550       End If

1551       rsTestSetup!Approved = False

1552       I = cmbLoopNumber.ListIndex
1553       If I = -1 Then
1554           d = 1
1555           rsTestSetup!loopnumber = d
1556       Else
1557           d = cmbLoopNumber.ItemData(I)
1558           rsTestSetup!loopnumber = d
1559       End If

1560       I = cmbSuctDia.ListIndex
1561       If I = -1 Then
1562           d = -1
1563       Else
1564           d = cmbSuctDia.ItemData(I)
1565           rsTestSetup!SuctDiam = d
1566       End If

1567       I = cmbDischDia.ListIndex
1568       If I = -1 Then
1569           d = -1
1570       Else
1571           d = cmbDischDia.ItemData(I)
1572           rsTestSetup!DischDiam = d
1573       End If

1574       I = cmbTachID.ListIndex
1575       If I = -1 Then
1576           d = 1
1577           rsTestSetup!tachid = d
1578       Else
1579           d = cmbTachID.ItemData(I)
1580           rsTestSetup!tachid = d
1581       End If

1582       I = cmbAnalyzerNo.ListIndex
1583       If I = -1 Then
1584           d = 1
1585       Else
1586           d = cmbAnalyzerNo.ItemData(I)
1587       End If
1588       rsTestSetup!analyzerno = d

1589       I = cmbTestSpec.ListIndex
1590       If I = -1 Then
1591           d = 1
1592       Else
1593           d = cmbTestSpec.ItemData(I)
1594       End If
1595       rsTestSetup!testspec = d

1596       I = cmbVoltage.ListIndex
1597       If I = -1 Then
1598           d = 1
1599       Else
1600           d = cmbVoltage.ItemData(I)
1601       End If
1602       rsTestSetup!Voltage = d

1603       I = cmbFrequency.ListIndex
1604       If I = -1 Then
1605           d = 1
1606       Else
1607           d = cmbFrequency.ItemData(I)
1608       End If
1609       rsTestSetup!Frequency = d

1610       I = cmbMounting.ListIndex
1611       If I = -1 Then
1612           d = 1
1613       Else
1614           d = cmbMounting.ItemData(I)
1615       End If
1616       rsTestSetup!Mounting = d

1617       I = cmbPLCNo.ListIndex
1618       If I = -1 Then
1619           d = 8
1620       Else
1621           d = cmbPLCNo.ItemData(I)
1622       End If
1623       rsTestSetup!PLCNo = d

1624       rsTestSetup!ImpFeathered = chkFeathered.value

1625       If chkTrimmed.value = 1 Then
1626           rsTestSetup!ImpTrimmed = Val(txtImpTrim)
1627       Else
1628           rsTestSetup!ImpTrimmed = 0
1629       End If
1630       chkTrimmed_Click

1631       If chkOrifice.value = 1 Then
1632           rsTestSetup!PumpDischOrifice = Val(txtOrifice)
1633       Else
1634           rsTestSetup!PumpDischOrifice = 0
1635       End If
1636       chkOrifice_Click

1637       If chkCircOrifice.value = 1 Then
1638           rsTestSetup!CircFlowOrifice = Val(txtCircOrifice)
1639       Else
1640           rsTestSetup!CircFlowOrifice = 0
1641       End If
1642       chkCircOrifice_Click

1643       chkBalanceHoles_Click

1644       If chkNPSH.value = 1 Then
1645           txtNPSHFile.Visible = True
1646           rsTestSetup!NPSHFile = txtNPSHFile
1647       Else
1648           rsTestSetup!NPSHFile = vbNullString
1649           txtNPSHFile.Visible = False
1650       End If

1651       If chkPictures.value = 1 Then
1652           txtPicturesFile.Visible = True
1653           rsTestSetup!PictureFile = txtPicturesFile
1654       Else
1655           rsTestSetup!PictureFile = vbNullString
1656           txtPicturesFile.Visible = False
1657       End If

1658       If chkVibration.value = 1 Then
1659           txtVibrationFile.Visible = True
1660           rsTestSetup!VibrationFile = txtVibrationFile
1661       Else
1662           rsTestSetup!VibrationFile = vbNullString
1663           txtVibrationFile.Visible = False
1664       End If

1665       If boWriteDataWritten Then
1666           rsTestSetup!DataWritten = True
1667       Else
1668           rsTestSetup!DataWritten = False
1669       End If

           'for TEMC Inspection Report
1670       If LenB(frmPLCData.txtTestAndInspection(0).Text) <> 0 Then
1671           rsTestSetup!InsulationMeggerVolts = frmPLCData.txtTestAndInspection(0).Text
1672       Else
1673           rsTestSetup!InsulationMeggerVolts = ""
1674       End If

1675       If LenB(frmPLCData.txtTestAndInspection(1).Text) <> 0 Then
1676           rsTestSetup!InsulationMegOhms = frmPLCData.txtTestAndInspection(1).Text
1677       Else
1678           rsTestSetup!InsulationMegOhms = ""
1679       End If

1680       If LenB(frmPLCData.txtTestAndInspection(2).Text) <> 0 Then
1681           rsTestSetup!DielectricVolts = frmPLCData.txtTestAndInspection(2).Text
1682       Else
1683           rsTestSetup!DielectricVolts = ""
1684       End If

1685       If LenB(frmPLCData.txtTestAndInspection(3).Text) <> 0 Then
1686           rsTestSetup!DielectricTime = frmPLCData.txtTestAndInspection(3).Text
1687       Else
1688           rsTestSetup!DielectricTime = ""
1689       End If

1690       If LenB(frmPLCData.txtTestAndInspection(4).Text) <> 0 Then
1691           rsTestSetup!HydrostaticValue = frmPLCData.txtTestAndInspection(4).Text
1692       Else
1693           rsTestSetup!HydrostaticValue = ""
1694       End If

1695       If LenB(frmPLCData.txtTestAndInspection(5).Text) <> 0 Then
1696           rsTestSetup!HydrostaticTime = frmPLCData.txtTestAndInspection(5).Text
1697       Else
1698           rsTestSetup!HydrostaticTime = ""
1699       End If

1700       If LenB(frmPLCData.txtTestAndInspection(6).Text) <> 0 Then
1701           rsTestSetup!PneumaticValue = frmPLCData.txtTestAndInspection(6).Text
1702       Else
1703           rsTestSetup!PneumaticValue = ""
1704       End If

1705       If LenB(frmPLCData.txtTestAndInspection(7).Text) <> 0 Then
1706           rsTestSetup!PneumaticTime = frmPLCData.txtTestAndInspection(7).Text
1707       Else
1708           rsTestSetup!PneumaticTime = ""
1709       End If

1710       I = cmbTestAndInspection(0).ListIndex
1711       If I = -1 Then
1712           rsTestSetup!HydrostaticUnits = ""
1713       Else
1714           rsTestSetup!HydrostaticUnits = cmbTestAndInspection(0).Text
1715       End If


1716       I = cmbTestAndInspection(1).ListIndex
1717       If I = -1 Then
1718           rsTestSetup!PneumaticUnits = ""
1719       Else
1720           rsTestSetup!PneumaticUnits = cmbTestAndInspection(1).Text
1721       End If

           'use abs to convert from 1 and 0 to boolean
1722       rsTestSetup!insulationgood = Abs(TestAndInspectionGood(0).value)
1723       rsTestSetup!DielectricGood = Abs(TestAndInspectionGood(1).value)
1724       rsTestSetup!HydrostaticGood = Abs(TestAndInspectionGood(2).value)
1725       rsTestSetup!PneumaticGood = Abs(TestAndInspectionGood(3).value)
1726       rsTestSetup!GeneralAppearanceGood = Abs(TestAndInspectionGood(4).value)
1727       rsTestSetup!OutlineDimensionsGood = Abs(TestAndInspectionGood(5).value)
1728       rsTestSetup!MotorNoLoadTestGood = Abs(TestAndInspectionGood(6).value)
1729       rsTestSetup!MotorLockedRotorTestGood = Abs(TestAndInspectionGood(7).value)
1730       rsTestSetup!HydrostaticTestGood = Abs(TestAndInspectionGood(8).value)
1731       rsTestSetup!HydraulicTestGood = Abs(TestAndInspectionGood(9).value)
1732       rsTestSetup!NPSHTestGood = Abs(TestAndInspectionGood(10).value)
1733       rsTestSetup!CleanPurgeSealGood = Abs(TestAndInspectionGood(11).value)
1734       rsTestSetup!PaintCheckGood = Abs(TestAndInspectionGood(12).value)
1735       rsTestSetup!NameplateGood = Abs(TestAndInspectionGood(13).value)
1736       rsTestSetup!SupervisorApproval = Abs(TestAndInspectionGood(14).value)

           'update the database
1737       rsTestSetup.Update

1738       If boFoundTestData = False Then     'add 8 records for 8 tests
1739           AddTestData
1740       End If

1741       rsTestSetup.Filter = vbNullString
' <VB WATCH>
1742       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1743       Exit Sub
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
1744       On Error GoTo vbwErrHandler
1745       Const VBWPROCNAME = "frmPLCData.cmdExit_Click"
1746       If vbwProtector.vbwTraceProc Then
1747           Dim vbwProtectorParameterString As String
1748           If vbwProtector.vbwTraceParameters Then
1749               vbwProtectorParameterString = "()"
1750           End If
1751           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1752       End If
' </VB WATCH>
1753       End
' <VB WATCH>
1754       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1755       Exit Sub
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
1756       On Error GoTo vbwErrHandler
1757       Const VBWPROCNAME = "frmPLCData.cmdFindMagtrols_Click"
1758       If vbwProtector.vbwTraceProc Then
1759           Dim vbwProtectorParameterString As String
1760           If vbwProtector.vbwTraceParameters Then
1761               vbwProtectorParameterString = "()"
1762           End If
1763           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1764       End If
' </VB WATCH>
1765       FindMagtrols
' <VB WATCH>
1766       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1767       Exit Sub
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
1768       On Error GoTo vbwErrHandler
1769       Const VBWPROCNAME = "frmPLCData.cmdFindPump_Click"
1770       If vbwProtector.vbwTraceProc Then
1771           Dim vbwProtectorParameterString As String
1772           If vbwProtector.vbwTraceParameters Then
1773               vbwProtectorParameterString = "()"
1774           End If
1775           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1776       End If
' </VB WATCH>
1777       Dim sAns As String
1778       Dim sSO As String
1779       Dim sParam As String
1780       Dim sName As String

1781       Dim I As Integer

           'clear the data
1782       BlankData

           'set TC and AI labels with default values
1783       txtTitle(0).Text = "TC 1"
1784       txtTitle(1).Text = "(F)"
1785       txtTitle(2).Text = "TC 2"
1786       txtTitle(3).Text = "(F)"
1787       txtTitle(4).Text = "TC 3"
1788       txtTitle(5).Text = "(F)"
1789       txtTitle(6).Text = "TC 4"
1790       txtTitle(7).Text = "(F)"
1791       txtTitle(20).Text = "Circ Flow"
1792       txtTitle(21).Text = "(GPM)"
1793       txtTitle(22).Text = "P1"
1794       txtTitle(23).Text = "(psig)"
1795       txtTitle(24).Text = "P2"
1796       txtTitle(25).Text = "(psig)"
1797       txtTitle(26).Text = "AI 4"
1798       txtTitle(27).Text = ""


1799       For I = 0 To 7
1800           lblAutoMan(I).Caption = "Auto"
1801       Next I

1802       lblAutoMan(5).Caption = "Man"
1803       lblAutoMan(6).Caption = "Man"

1804       txtFlowDisplay.Enabled = False
1805       txtSuctionDisplay.Enabled = False
1806       txtDischargeDisplay.Enabled = False
1807       txtTemperatureDisplay.Enabled = False
1808       txtAI1Display.Enabled = False
1809       txtAI2Display.Enabled = False
1810       txtAI3Display.Enabled = False
1811       txtAI4Display.Enabled = False


1812       cmdFindPump.Default = False

           'set all found booleans to false
       '    boUsingHP = False
1813       boFoundPump = False
1814       boPumpIsApproved = False
1815       boFoundTestSetup = False
1816       boFoundTestData = False


           'get rid of all test dates in combo box
1817       For I = cmbTestDate.ListCount - 1 To 0 Step -1
1818           cmbTestDate.RemoveItem 0
1819       Next I

1820       rsTestData.Filter = "SerialNumber = ''"

1821       DataGrid2.ClearFields
1822       ClearEff

1823       If rsPumpData.State = adStateOpen Then
1824           If rsPumpData.BOF = False Or rsPumpData.EOF = False Then
1825               rsPumpData.Update
1826           End If
1827           rsPumpData.Close
1828       End If

           'parse the serial number to make sure it is formed correctly
1829       Dim ok As Boolean
1830       ok = UCase(txtSN.Text) Like "[A-Z][A-Z][0-9][0-9][0-9][0-9][A-Z]" Or UCase(txtSN.Text) Like "[A-Z][A-Z][0-9][0-9][0-9][0-9][A-Z]-[0-9]" Or UCase(txtSN.Text) Like "[A-Z][A-Z][0-9][0-9][0-9][0-9][A-Z]-[0-9][0-9]"
1831       If Not ok Then
1832           MsgBox "Serial Number must be 2 letters, 4 numbers, and 1 letter. Please re-enter.", vbOKOnly, "Serial Number not correctly formed."
' <VB WATCH>
1833       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1834           Exit Sub
1835       End If

           'find the pump listed in the Serial Number text box
1836       qyPumpData.ActiveConnection = cnPumpData
1837       qyPumpData.CommandText = "SELECT * From TempPumpData WHERE (((TempPumpData.SerialNumber)='" & _
                                    txtSN.Text & "'))"
1838       rsPumpData.CursorType = adOpenStatic
1839       rsPumpData.CursorLocation = adUseClient
1840       rsPumpData.Index = "SerialNumber"
1841       rsPumpData.Open qyPumpData
1842       boEpicorFound = False

1843       If rsPumpData.BOF = True And rsPumpData.EOF = True Then
               'if the bof=eof, we have an empty recordset
1844           boFoundPump = False
1845       Else
               'we found it
1846           boFoundPump = True
1847       End If

1848       If boFoundPump = False Then
               'not found in either database, try HP?
1849           sAns = MsgBox("Pump Not Found in the Database.  Look in Epicor?", vbYesNo, "Can't Find Pump")
1850           If sAns = vbNo Then     'new pump - don't get data from HP
1851               boUsingEpicor = False
1852           Else
1853               boUsingEpicor = True
       '            boUsingHP = False
1854           End If
       '        If boUsingEpicor = False Then
       '            sAns = MsgBox("Pump Not Found in the Database.  Look on the HP?", vbYesNo, "Can't Find Pump")
       '            If sAns = vbNo Then     'new pump - don't get data from HP
       '                 boUsingHP = False
       '            Else
       '                boUsingHP = True
       '            End If
       '        End If
1855           EnablePumpDataControls
1856           EnableTestSetupDataControls
1857           EnableTestDataControls
       '        BlankData               'clear any data on the screen
1858           cmdAddNewBalanceHoles.Visible = True

1859       End If

1860       If boFoundPump = True Then    'found the pump
1861           If rsPumpData.Fields("Approved") = True Then
1862               DisablePumpDataControls                         'if it's in the real database, don't allow changes here
1863               boPumpIsApproved = True
1864               lblPumpApproved.Visible = True
1865               If boCanApprove Then
1866                   cmdApprovePump.Caption = "Unapprove this pump"
1867               End If
1868               frmPLCData.cmdApproveTestDate.Enabled = True
1869           Else
1870               EnablePumpDataControls                          'it's in the temp database, allow changes
1871               boPumpIsApproved = False
1872               boTestDateIsApproved = False
1873               lblPumpApproved.Visible = False
1874               If boCanApprove Then
1875                   cmdApprovePump.Caption = "Approve this pump"
1876               End If
1877               cmdApproveTestDate.Caption = "You Must Approve Pump First"
1878               frmPLCData.cmdApproveTestDate.Enabled = False
1879           End If

               'found the pump, show the data
1880           txtModelNo.Text = rsPumpData.Fields("ModelNumber")
1881           frmPLCData.optMfr(0).value = rsPumpData.Fields("ChempumpPump")

1882           If rsPumpData.Fields("ChempumpPump") = True Then
1883               SetCombo cmbMotor, "Motor", rsPumpData
1884               SetCombo cmbDesignPressure, "DesignPressure", rsPumpData
1885               SetCombo cmbRPM, "RPM", rsPumpData
1886               SetCombo cmbCirculationPath, "CirculationPath", rsPumpData
1887               SetCombo cmbStatorFill, "StatorFill", rsPumpData
1888               SetCombo cmbModel, "Model", rsPumpData
1889               SetCombo cmbModelGroup, "ModelGroup", rsPumpData
1890               RatedKW = 999
1891           End If

               'set the TEMC data
1892           If rsPumpData.Fields("ChempumpPump") = False Then
1893               SetCombo cmbTEMCAdapter, "TEMCAdapter", rsPumpData
1894               SetCombo cmbTEMCAdditions, "TEMCAdditions", rsPumpData
1895               SetCombo cmbTEMCCirculation, "TEMCCirculation", rsPumpData
1896               SetCombo cmbTEMCDesignPressure, "TEMCDesignPressure", rsPumpData
1897               SetCombo cmbTEMCNominalDischargeSize, "TEMCNominalDischargeSize", rsPumpData
1898               SetCombo cmbTEMCDivisionType, "TEMCDivisionType", rsPumpData
1899               SetCombo cmbTEMCImpellerType, "TEMCImpellerType", rsPumpData
1900               SetCombo cmbTEMCInsulation, "TEMCInsulation", rsPumpData
1901               SetCombo cmbTEMCJacketGasket, "TEMCJacketGasket", rsPumpData
1902               SetCombo cmbTEMCMaterials, "TEMCMaterials", rsPumpData
1903               SetCombo cmbTEMCModel, "TEMCModel", rsPumpData
1904               SetCombo cmbTEMCNominalImpSize, "TEMCNominalImpSize", rsPumpData
1905               SetCombo cmbTEMCOtherMotor, "TEMCOtherMotor", rsPumpData
1906               SetCombo cmbTEMCPumpStages, "TEMCPumpStages", rsPumpData
1907               SetCombo cmbTEMCNominalSuctionSize, "TEMCNominalSuctionSize", rsPumpData
1908               SetCombo cmbTEMCTRG, "TEMCTRG", rsPumpData
1909               SetCombo cmbTEMCVoltage, "TEMCVoltage", rsPumpData
1910           End If

               'write ship to and bill to info
1911           If Not IsNull(rsPumpData.Fields("ShipToCustomer")) Then
1912               txtShpNo.Text = rsPumpData.Fields("ShipToCustomer")
1913           Else
1914               txtShpNo.Text = vbNullString
1915           End If

1916           If Not IsNull(rsPumpData.Fields("BillToCustomer")) Then
1917               txtBilNo.Text = rsPumpData.Fields("BillToCustomer")
1918           Else
1919               txtBilNo.Text = vbNullString
1920           End If

1921           sName = "ImpellerDia"
1922           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1923               sParam = rsPumpData.Fields(sName)
1924           Else
1925               sParam = vbNullString
1926           End If
1927           txtImpellerDia.Text = sParam

1928           sName = "DesignFlow"
1929           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1930               sParam = rsPumpData.Fields(sName)
1931           Else
1932               sParam = vbNullString
1933           End If
1934           txtDesignFlow.Text = sParam

1935           sName = "DesignTDH"
1936           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1937               sParam = rsPumpData.Fields(sName)
1938           Else
1939               sParam = vbNullString
1940           End If
1941           txtDesignTDH.Text = sParam

1942           sName = "SpGr"
1943           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1944               sParam = rsPumpData.Fields(sName)
1945           Else
1946               sParam = vbNullString
1947           End If
1948           txtSpGr.Text = sParam

1949           sName = "Remarks"
1950           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1951               sParam = rsPumpData.Fields(sName)
1952           Else
1953               sParam = vbNullString
1954           End If
1955           txtRemarks.Text = sParam

1956           sName = "SalesOrderNumber"
1957           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1958               sParam = rsPumpData.Fields(sName)
1959           Else
1960               sParam = vbNullString
1961           End If
1962           txtSalesOrderNumber.Text = sParam

1963           sName = "ApplicationFluid"
1964           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1965               sParam = rsPumpData.Fields(sName)
1966           Else
1967               sParam = vbNullString
1968           End If
1969           txtLiquid.Text = sParam

1970           sName = "NPSHFile"
1971           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1972               sParam = rsPumpData.Fields(sName)
1973           Else
1974               sParam = vbNullString
1975           End If
1976           txtNPSHFileLocation.Text = sParam

1977           sName = "RVSPartNo"
1978           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1979               sParam = rsPumpData.Fields(sName)
1980           Else
1981               sParam = vbNullString
1982           End If
1983           txtRVSPartNo.Text = sParam

1984           sName = "CustPN"
1985           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1986               sParam = rsPumpData.Fields(sName)
1987           Else
1988               sParam = vbNullString
1989           End If
1990           txtXPartNum.Text = sParam

1991           sName = "CustPO"
1992           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1993               sParam = rsPumpData.Fields(sName)
1994           Else
1995               sParam = vbNullString
1996           End If
1997           txtCustPONum.Text = sParam

               'make sure table has custpn - see if last three digits of model no are numeric
       '        sName = "SalesOrderNumber"
       '        If rsPumpData.Fields(sName).ActualSize <> 0 Then
       '            If IsNumeric(Right(rsPumpData.Fields("ModelNumber"), 3)) Then 'no sales order no, must be supermarket
       '                rsPumpData.Fields("CustPN") = rsPumpData.Fields("RVSPartNo")
       '            Else
       '                rsPumpData.Fields("CustPN") = rsPumpData.Fields("ModelNumber")
       '            End If
       '        End If

1998           sName = "ApplicationViscosity"
1999           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2000               sParam = Format(rsPumpData.Fields(sName), "#0.00")
2001           Else
2002               sParam = vbNullString
2003           End If
2004           txtViscosity.Text = sParam

       'added from TEMC Inspection Report
2005           sName = "JobNumber"
2006           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2007               sParam = rsPumpData.Fields(sName)
2008           Else
2009               sParam = ""
2010           End If
2011           txtJobNum.Text = sParam

2012           sName = "Phases"
2013           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2014               sParam = rsPumpData.Fields(sName)
2015           Else
2016               sParam = vbNullString
2017           End If
2018           txtNoPhases.Text = sParam

2019           sName = "ThermalClass"
2020           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2021               sParam = rsPumpData.Fields(sName)
2022           Else
2023               sParam = vbNullString
2024           End If
2025           txtThermalClass.Text = sParam

2026           sName = "ExpClass"
2027           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2028               sParam = rsPumpData.Fields(sName)
2029           Else
2030               sParam = vbNullString
2031           End If
2032           txtExpClass.Text = sParam

2033           sName = "NPSHr"
2034           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2035               sParam = rsPumpData.Fields(sName)
2036           Else
2037               sParam = vbNullString
2038           End If
2039           txtNPSHr.Text = sParam

2040           sName = "LiquidTemperature"
2041           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2042               sParam = rsPumpData.Fields(sName)
2043           Else
2044               sParam = vbNullString
2045           End If
2046           txtLiquidTemperature.Text = sParam

2047           sName = "RatedInputPower"
2048           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2049               sParam = rsPumpData.Fields(sName)
2050           Else
2051               sParam = vbNullString
2052           End If
2053           txtRatedInputPower.Text = sParam

2054           sName = "FLCurrent"
2055           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2056               sParam = rsPumpData.Fields(sName)
2057           Else
2058               sParam = vbNullString
2059           End If
2060           txtAmps.Text = sParam

2061           sName = "TEMCFrameNumber"
2062           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2063               sParam = rsPumpData.Fields(sName)
2064           Else
2065               sParam = vbNullString
2066           End If
2067           txtTEMCFrameNumber.Text = sParam

2068           optMfr(0).value = rsPumpData.Fields("ChempumpPump")
2069           optMfr(1).value = Not optMfr(0).value

2070           If rsPumpData.Fields("Field1") = "Feathered" Then
2071               Me.chkSuperMarketFeathered.value = Checked
2072           Else
2073               Me.chkSuperMarketFeathered.value = Unchecked
2074           End If

               'select the testsetup data
2075           qyTestSetup.ActiveConnection = cnPumpData
2076           qyTestSetup.CommandText = "SELECT * FROM TempTestSetupData WHERE (((TempTestSetupData.SerialNumber)='" & _
                                    txtSN.Text & "')) ORDER BY Date"
       '        qyTestSetup.CommandText = "SELECT * FROM TempTestSetupData WHERE (((TempTestSetupData.SerialNumber)='" & _
'               txtSN.Text & "'))"

2077           With rsTestSetup
2078               If .State = adStateOpen Then
2079                   .Close
2080               End If
2081               .CursorLocation = adUseClient
2082               .CursorType = adOpenStatic
2083               .Index = "FindData"
2084               .Open qyTestSetup
2085           End With


               'add the selection of dates to the Test Date combo box
2086           If rsTestSetup.RecordCount <> 0 Then
2087               For I = 0 To cmbTestDate.ListCount - 1
2088                   cmbTestDate.RemoveItem 0
2089               Next I
2090               rsTestSetup.MoveFirst
2091               For I = 1 To rsTestSetup.RecordCount
2092                   cmbTestDate.AddItem rsTestSetup.Fields("Date")
2093                   rsTestSetup.MoveNext
2094               Next I
2095               rsTestSetup.MoveFirst
2096               boFoundTestSetup = True

2097               If rsTestSetup.Fields("Approved") = True Then
2098                   DisableTestSetupDataControls                         'if it's in the real database, don't allow changes here
2099                   boTestDateIsApproved = True
2100                   lblTestDateApproved.Visible = True
2101                   If boCanApprove Then
2102                       cmdApproveTestDate.Caption = "Unapprove this Test Date"
2103                   End If
2104               Else
2105                   EnableTestSetupDataControls                          'it's in the temp database, allow changes
2106                   lblTestDateApproved.Visible = False
2107                   If boCanApprove Then
2108                       cmdApproveTestDate.Caption = "Approve this Test Date"
2109                   End If
2110               End If
2111               cmbTestDate.ListIndex = 0
2112           Else
2113               MsgBox ("There is no Test Setup Data for Serial Number " & txtSN.Text)
2114               boFoundTestSetup = False        'didn't find any data
2115               boFoundTestData = False
2116               cmbTestDate.AddItem Date        'load with today
2117               cmbTestDate.ListIndex = 0       'show the entry
2118               EnableTestSetupDataControls
2119               txtTestRemarks.Text = ""
2120               txtVibAx.Text = ""
2121               txtVibRad.Text = ""
2122               txtThrustBal.Text = ""
2123               txtTEMCTRGReading.Text = ""
2124               txtTEMCFrontThrust.Text = ""
2125               txtTEMCRearThrust.Text = ""
' <VB WATCH>
2126       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2127               Exit Sub
2128           End If

2129           If cmbTestDate.ListCount = 1 Then       'if there's only one test date, select it
2130           End If
' <VB WATCH>
2131       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2132           Exit Sub
2133       End If


2134       Do While boUsingEpicor = True   'need a do loop to exit
2135           If boUsingEpicor = True Then
                   'Dim MyRecord As SNRecord
2136               Dim MyRecord As SNRecord
           '            I = InStr(1, txtSN.Text, "-")
           '            If I > 0 Then
2137                   MyRecord = GetEpicorODBCData(txtSN.Text, EpicorConnectionString)
           '            End If
2138               If MyRecord.SONumber = "" Then
2139                   MsgBox ("Not found in Epicor")
2140                   boUsingEpicor = False
2141                   boEpicorFound = False
2142                   Exit Do
2143               End If

2144               If MyRecord.SONumber = 0 Then
2145                   boEpicorFound = False
2146                   boUsingSupermarketTable = True
2147                   boUsingEpicor = False
2148               Else
2149                   boEpicorFound = True
2150                   boUsingSupermarketTable = False
2151               End If

2152               If boEpicorFound = True Then
2153                   boUsingEpicor = False
       '                boEpicorFound = True
2154                   txtSalesOrderNumber.Text = MyRecord.SONumber
2155                   txtLineNumber.Text = MyRecord.SOLine
2156                   txtBilNo.Text = MyRecord.Customer
2157                   txtXPartNum.Text = MyRecord.XPartNum
2158                   txtCustPONum.Text = MyRecord.CustomerPO

2159                   If MyRecord.ShipTo = "" Then
2160                       txtShpNo.Text = MyRecord.Customer
2161                   Else
2162                       txtShpNo.Text = MyRecord.ShipTo
2163                   End If
2164                   txtModelNo.Text = MyRecord.PartNum
2165                   txtModelNo_Change
2166                   txtDesignTDH.Text = MyRecord.TDH
2167                   txtSpGr.Text = MyRecord.SpGr
2168                   txtImpellerDia.Text = MyRecord.ImpellerDiameter
2169                   txtDesignFlow.Text = MyRecord.Flow
2170                   txtNoPhases.Text = MyRecord.Phases
2171                   txtNPSHr.Text = MyRecord.NPSHr
2172                   txtRatedInputPower.Text = MyRecord.RatedInputPower
2173                   txtAmps.Text = MyRecord.FLCurrent
2174                   txtThermalClass.Text = MyRecord.ThermalClass
2175                   txtViscosity.Text = MyRecord.Viscosity
2176                   txtExpClass.Text = MyRecord.ExpClass
2177                   txtLiquidTemperature.Text = MyRecord.LiquidTemp
2178                   txtLiquid.Text = MyRecord.Fluid
2179                   txtJobNum.Text = MyRecord.JobNumber

2180                   For I = 0 To cmbStatorFill.ListCount - 1
2181                       If InStr(1, UCase$(MyRecord.StatorFill), UCase$(cmbStatorFill.List(I))) <> 0 Then
2182                           cmbStatorFill.ListIndex = I
2183                           Exit For
2184                       End If
2185                   Next I

2186                   For I = 0 To cmbCirculationPath.ListCount - 1
2187                       If InStr(1, UCase$(MyRecord.CirculationPath), UCase$(cmbCirculationPath.List(I))) <> 0 Then
2188                           cmbCirculationPath.ListIndex = I
2189                           Exit For
2190                       End If
2191                   Next I

2192                   For I = 0 To cmbDesignPressure.ListCount - 1
2193                       If InStr(1, MyRecord.DesignPressure, cmbDesignPressure.List(I)) <> 0 Then
2194                           cmbDesignPressure.ListIndex = I
2195                           Exit For
2196                       End If
2197                   Next I

2198                   For I = 0 To cmbVoltage.ListCount - 1
2199                       If InStr(1, MyRecord.Voltage, cmbVoltage.List(I)) <> 0 Then
2200                           cmbVoltage.ListIndex = I
2201                           Exit For
2202                       End If
2203                   Next I

2204                   For I = 0 To cmbFrequency.ListCount - 1
2205                       If InStr(1, MyRecord.Frequency, sName) <> 0 Then
2206                           cmbFrequency.ListIndex = I
2207                           Exit For
2208                       End If
2209                   Next I

2210                   For I = 0 To cmbRPM.ListCount - 1
2211                       If InStr(1, MyRecord.RPM, cmbRPM.List(I)) <> 0 Then
2212                           cmbRPM.ListIndex = I
2213                           Exit For
2214                       End If
2215                   Next I

2216                   For I = 0 To cmbSuctDia.ListCount - 1
2217                       If InStr(1, MyRecord.SuctFlangeSize, cmbSuctDia.List(I)) <> 0 Then
2218                           cmbSuctDia.ListIndex = I
2219                           Exit For
2220                       End If
2221                   Next I

2222                   For I = 0 To cmbDischDia.ListCount - 1
2223                       If InStr(1, MyRecord.DischFlangeSize, cmbDischDia.List(I)) <> 0 Then
2224                           cmbDischDia.ListIndex = I
2225                           Exit For
2226                       End If
2227                   Next I

2228                   For I = 0 To cmbTestSpec.ListCount - 1
2229                       If InStr(1, MyRecord.TestProcedure, cmbTestSpec.List(I)) <> 0 Then
2230                           cmbTestSpec.ListIndex = I
2231                           Exit For
2232                       End If
2233                   Next I

2234                   For I = 0 To cmbMotor.ListCount - 1
2235                       If InStr(1, MyRecord.MotorSize, cmbMotor.List(I)) <> 0 Then
2236                           cmbMotor.ListIndex = I
2237                           Exit For
2238                       End If
2239                   Next I


2240               End If
2241           End If
2242       Loop

2243       If boUsingSupermarketTable = True Then
2244           GetSuperMarketPump MyRecord.PartNum, MyRecord.JobNumber
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
2245       End If
' <VB WATCH>
2246       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2247       Exit Sub
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
2248       On Error GoTo vbwErrHandler
2249       Const VBWPROCNAME = "frmPLCData.cmdModifyBalanceHoleData_Click"
2250       If vbwProtector.vbwTraceProc Then
2251           Dim vbwProtectorParameterString As String
2252           If vbwProtector.vbwTraceParameters Then
2253               vbwProtectorParameterString = "()"
2254           End If
2255           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2256       End If
' </VB WATCH>
2257       Dim strInput As String
2258       Dim I As Integer
2259       Dim sNumber As Integer
2260       Dim sDia As String
2261       Dim sBC As String

2262       cmdModifyBalanceHoleData.Visible = False

2263       If dgBalanceHoles.SelBookmarks.Count = 0 Then
2264           cmdModifyBalanceHoleData.Visible = False
' <VB WATCH>
2265       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2266           Exit Sub
2267       End If

2268       rsBalanceHoles.MoveFirst
2269       rsBalanceHoles.Move dgBalanceHoles.SelBookmarks(0) - dgBalanceHoles.FirstRow

2270       sNumber = rsBalanceHoles!Number
2271       If rsBalanceHoles!diameter = 99 Then
2272           sDia = "Slot"
2273       Else
2274           sDia = str(rsBalanceHoles!diameter)
2275       End If
2276       If rsBalanceHoles!boltcircle = 99 Then
2277           sBC = "Unknown"
2278       Else
2279           sBC = str(rsBalanceHoles!boltcircle)
2280       End If


           'get the data for the balance holes
2281       strInput = InputBox("Enter Number of Holes (0 to delete entry)", , sNumber)
2282       If strInput = "" Then
2283           GoTo DeleteIt
2284       End If
2285       sNumber = CInt(strInput)
2286       If Val(sNumber) = 0 Then
2287           GoTo DeleteIt
2288       End If

2289       strInput = InputBox("Enter Decimal Value of Hole Diameter or 'Slot' (For Example, 0.675) ", , sDia)
2290       If strInput <> "" Then
2291           If UCase(strInput) = "SLOT" Then
2292               strInput = 99
2293           End If
2294           sDia = CSng(strInput)
2295       Else
2296           GoTo CancelPressed
2297       End If

2298       strInput = InputBox("Enter Decimal Value of Bolt Circle or 'Unknown' (For Example, 4.525)", , sBC)
2299       If strInput <> "" Then
2300           If UCase(strInput) = "UNKNOWN" Then
2301               strInput = 99
2302           End If
2303           sBC = CSng(strInput)
2304       Else
2305           GoTo CancelPressed
2306       End If

2307       rsBalanceHoles!Number = sNumber
2308       rsBalanceHoles!diameter = sDia
2309       rsBalanceHoles!boltcircle = sBC

2310       rsBalanceHoles.Update
           'rsBalanceHoles.Filter = "SerialNo = '" & frmPLCData.txtSN.Text & "'"

2311       GetBalanceHoleData txtSN.Text, cmbTestDate.Text
       '    rsBalanceHoles.Requery
2312       rsBalanceHoles.MoveLast
2313       dgBalanceHoles.Refresh
2314       chkBalanceHoles.value = 1
2315       rsBalanceHoles.MoveFirst

' <VB WATCH>
2316       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2317       Exit Sub

2318   CancelPressed:
2319       MsgBox "No New Balance Hole Data Entered", vbOKOnly

2320   DeleteIt:
2321       If (MsgBox("Do you really want to delete this entry?", vbYesNo, "Deleting Balance Hole Data. . .")) = vbYes Then
2322           rsBalanceHoles.Delete
2323           rsBalanceHoles.Update
2324           GetBalanceHoleData txtSN.Text, cmbTestDate.Text
       '        rsBalanceHoles.Requery
2325           If Not rsBalanceHoles.EOF Then
2326               rsBalanceHoles.MoveLast
2327           End If
2328           dgBalanceHoles.Refresh
2329           chkBalanceHoles.value = 1
2330           If Not rsBalanceHoles.BOF Then
2331               rsBalanceHoles.MoveFirst
2332           End If
2333       End If


' <VB WATCH>
2334       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2335       Exit Sub
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
2336       On Error GoTo vbwErrHandler
2337       Const VBWPROCNAME = "frmPLCData.cmdReport_Click"
2338       If vbwProtector.vbwTraceProc Then
2339           Dim vbwProtectorParameterString As String
2340           If vbwProtector.vbwTraceParameters Then
2341               vbwProtectorParameterString = "()"
2342           End If
2343           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2344       End If
' </VB WATCH>
2345       Dim I As Integer

2346       ExportToExcel

' <VB WATCH>
2347       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2348       Exit Sub
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
2349       On Error GoTo vbwErrHandler
2350       Const VBWPROCNAME = "frmPLCData.cmdSearchForPump_Click"
2351       If vbwProtector.vbwTraceProc Then
2352           Dim vbwProtectorParameterString As String
2353           If vbwProtector.vbwTraceParameters Then
2354               vbwProtectorParameterString = "()"
2355           End If
2356           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2357       End If
' </VB WATCH>
2358       LoadCombo frmSearch.cmbSearchModel, "TEMCHydraulics"

2359       frmSearch.Show
' <VB WATCH>
2360       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2361       Exit Sub
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
2362       On Error GoTo vbwErrHandler
2363       Const VBWPROCNAME = "frmPLCData.cmdSelectSupermarket_Click"
2364       If vbwProtector.vbwTraceProc Then
2365           Dim vbwProtectorParameterString As String
2366           If vbwProtector.vbwTraceParameters Then
2367               vbwProtectorParameterString = "()"
2368           End If
2369           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2370       End If
' </VB WATCH>
2371       grpSupermarket.Visible = False
' <VB WATCH>
2372       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2373       Exit Sub
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
2374       On Error GoTo vbwErrHandler
2375       Const VBWPROCNAME = "frmPLCData.cmdWriteSP_Click"
2376       If vbwProtector.vbwTraceProc Then
2377           Dim vbwProtectorParameterString As String
2378           If vbwProtector.vbwTraceParameters Then
2379               vbwProtectorParameterString = "()"
2380           End If
2381           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2382       End If
' </VB WATCH>
2383       Dim rc As String
2384       Dim S As String

           'write the set point data to the PLC
2385           bWrite = True
2386           S = Right$("0000" & txtWriteSPData, 4)
2387           S = Right$(S, 2) & Left$(S, 2)
2388           rc = StringToByteArray(S, ByteBuffer)

2389           DataLength = HexConvert(ByteBuffer, 2)
2390           DataAddress = StringToHexInt("2005")

2391           rc = GetData

2392           bWrite = False
' <VB WATCH>
2393       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2394       Exit Sub
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
2395       On Error GoTo vbwErrHandler
2396       Const VBWPROCNAME = "frmPLCData.btnRunNPSH_Click"
2397       If vbwProtector.vbwTraceProc Then
2398           Dim vbwProtectorParameterString As String
2399           If vbwProtector.vbwTraceParameters Then
2400               vbwProtectorParameterString = "()"
2401           End If
2402           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2403       End If
' </VB WATCH>
2404       Static OriginalColor As Long
2405       If btnRunNPSH.Caption = "Run NPSH" Then
2406           btnRunNPSH.Caption = "Cancel NPSH Run"
2407           OriginalColor = btnRunNPSH.BackColor
2408           tmrNPSHr.Enabled = False
2409           btnRunNPSH.BackColor = vbRed
2410           If boCanApprove Then
2411               txtNPSH(5).Visible = True
2412               lbltab4(5).Visible = True
2413           Else
2414               txtNPSH(5).Visible = False
2415               lbltab4(5).Visible = False
2416           End If
2417           WroteNPSHr = False

2418           frmNPSH.Visible = True
2419           txtNPSH(5).Enabled = True
2420           If Val(txtTDH.Text) <= 10 Then
2421               MsgBox "This test will not work starting with this starting TDH.  Ending test...", vbOKOnly, "Flow is 0"
2422               btnRunNPSH.Caption = "Run NPSH"
2423               btnRunNPSH.BackColor = OriginalColor
2424               frmNPSH.Visible = False
' <VB WATCH>
2425       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2426               Exit Sub
2427           End If
               'load initial values
2428           If DataGrid2.Row = -1 Then
2429               MsgBox "You must write the normal test data to this row before you run NPSH.", vbOKOnly, "Nothing written for this row"
2430               btnRunNPSH.Caption = "Run NPSH"
2431               btnRunNPSH.BackColor = OriginalColor
2432               frmNPSH.Visible = False
' <VB WATCH>
2433       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2434               Exit Sub
2435           Else
2436               DataGrid2.Row = UpDown1.value - 1
2437           End If

2438           txtNPSH(0).Text = DataGrid2.Columns("Flow")
2439           txtNPSH(3).Text = DataGrid2.Columns("TDH")
2440           txtNPSH(4) = 0
               'txtNPSH(0).Text = txtFlow.Text
               'txtNPSH(3).Text = txtTDH.Text
2441           txtNPSH(4) = 0
2442       Else
2443           btnRunNPSH.Caption = "Run NPSH"
2444           btnRunNPSH.BackColor = OriginalColor
2445           frmNPSH.Visible = False
2446       End If

           'ReportToExcel
' <VB WATCH>
2447       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2448       Exit Sub
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
2449       On Error GoTo vbwErrHandler
2450       Const VBWPROCNAME = "frmPLCData.updown1_change"
2451       If vbwProtector.vbwTraceProc Then
2452           Dim vbwProtectorParameterString As String
2453           If vbwProtector.vbwTraceParameters Then
2454               vbwProtectorParameterString = "()"
2455           End If
2456           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2457       End If
' </VB WATCH>
2458       Dim sName As String

2459       If Not rsTestData.BOF Then
2460           rsTestData.MoveFirst
2461       End If

2462       If Not rsTestData.BOF Or Not rsTestData.EOF Then
2463           rsTestData.Move UpDown1.value - 1
2464       End If

2465       sName = "VibrationX"
2466       If rsTestData.Fields(sName).ActualSize <> 0 Then
2467           txtVibAx.Text = rsTestData.Fields(sName)
2468       Else
       '        txtVibAx.Text = vbNullString
2469       End If

2470       sName = "VibrationY"
2471       If rsTestData.Fields(sName).ActualSize <> 0 Then
2472           txtVibRad.Text = rsTestData.Fields(sName)
2473       Else
       '        txtVibRad.Text = vbNullString
2474       End If

2475       sName = "Remarks"
2476       If rsTestData.Fields(sName).ActualSize <> 0 Then
2477           txtTestRemarks.Text = rsTestData.Fields(sName)
2478       Else
       '        txtTestRemarks.Text = vbNullString
2479       End If

2480       sName = "ThrustBalance"
2481       If rsTestData.Fields(sName).ActualSize <> 0 Then
2482           txtThrustBal.Text = rsTestData.Fields(sName)
2483       Else
       '        txtThrustBal.Text = vbNullString
2484       End If

2485       sName = "TEMCTRG"
2486       If rsTestData.Fields(sName).ActualSize <> 0 Then
2487           txtTEMCTRGReading.Text = rsTestData.Fields(sName)
2488       Else
2489           txtTEMCTRGReading.Text = 0
       '        txtTEMCTRGReading.Text = vbNullString
2490       End If

2491       sName = "TEMCFrontThrust"
2492       If rsTestData.Fields(sName).ActualSize <> 0 Then
2493           txtTEMCFrontThrust.Text = rsTestData.Fields(sName)
2494       Else
       '        txtTEMCFrontThrust.Text = vbNullString
2495       End If

2496       sName = "TEMCRearThrust"
2497       If rsTestData.Fields(sName).ActualSize <> 0 Then
2498           txtTEMCRearThrust.Text = rsTestData.Fields(sName)
2499       Else
       '        txtTEMCRearThrust.Text = vbNullString
2500       End If
2501       sName = "TEMCMomentArm"
2502       If rsTestData.Fields(sName).ActualSize <> 0 Then
2503           txtTEMCMomentArm.Text = rsTestData.Fields(sName)
2504       Else
       '        txtTEMCMomentArm.Text = vbNullString
2505       End If
2506       sName = "TEMCThrustRigPressure"
2507       If rsTestData.Fields(sName).ActualSize <> 0 Then
2508           txtTEMCThrustRigPressure.Text = rsTestData.Fields(sName)
2509       Else
       '        txtTEMCThrustRigPressure.Text = vbNullString
2510       End If
2511       sName = "TEMCViscosity"
2512       If rsTestData.Fields(sName).ActualSize <> 0 And rsTestData.Fields(sName) <> 0 Then
2513           txtTEMCViscosity.Text = rsTestData.Fields(sName)
2514       Else
       '        txtTEMCViscosity.Text = vbNullString
2515       End If

2516       CalculateTEMCForce

2517       rsEff.MoveFirst
2518       rsEff.Move UpDown1.value - 1
' <VB WATCH>
2519       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2520       Exit Sub
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
2521       On Error GoTo vbwErrHandler
2522       Const VBWPROCNAME = "frmPLCData.CalculateTEMCForce"
2523       If vbwProtector.vbwTraceProc Then
2524           Dim vbwProtectorParameterString As String
2525           If vbwProtector.vbwTraceParameters Then
2526               vbwProtectorParameterString = "()"
2527           End If
2528           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2529       End If
' </VB WATCH>
2530       Dim NoOfPoles As Integer
2531       Dim Frequency As Integer
2532       Dim Additions As String
2533       Dim Frame As String
2534       Dim VOverA As Double
2535       Dim Force As Double
2536       Dim Gravity As Double

2537       If Val(txtSpGr.Text) = 0 Then
2538           Gravity = 1
2539       Else
2540           Gravity = CDbl(Val(txtSpGr.Text))
2541       End If

           'show calculated values
2542       If Val(txtTEMCFrontThrust.Text) = 0 Then
2543           If Val(txtTEMCRearThrust.Text) = 0 Then
               'no thrust entered
2544               lblTEMCFrontRear.Visible = False
2545               txtTEMCCalcForce.Text = " "
2546           Else
                   'rear thrust
2547               txtTEMCCalcForce.Text = Format(Gravity * (Val(txtTEMCRearThrust.Text) * Val(txtTEMCMomentArm.Text) - (Val(txtTEMCThrustRigPressure.Text) / 14.223) * 4.5), "##0.0")
2548               lblTEMCFrontRear.Caption = "REAR"
2549               lblTEMCFrontRear.Visible = True
2550           End If
2551       Else
               'front thrust
2552           txtTEMCCalcForce.Text = Format(Gravity * (Val(txtTEMCFrontThrust.Text) * Val(txtTEMCMomentArm.Text) + (Val(txtTEMCThrustRigPressure.Text) / 14.223) * 4.5), "##0.0")
2553           lblTEMCFrontRear.Caption = "FRONT"
2554           lblTEMCFrontRear.Visible = True
2555       End If

2556       If Val(txtTEMCCalcForce.Text) < 0 Then
2557           txtTEMCCalcForce.Text = -txtTEMCCalcForce
2558           lblTEMCFrontRear.Caption = "FRONT"
2559       End If

           'see how many poles we have, it's the next to last number in the frame size
2560       If Len(txtTEMCFrameNumber) > 2 Then
2561           NoOfPoles = 2 * Val(Left$(Right$(txtTEMCFrameNumber.Text, 2), 1))
2562       End If

2563       If cmbTEMCAdditions.ListIndex <> -1 Then
2564           Additions = Mid$(cmbTEMCAdditions.List(cmbTEMCAdditions.ListIndex), 2, 1)
2565           If Additions = "A" Or Additions = "E" Or Additions = "G" Or Additions = "J" Then
2566               Frequency = 60
2567           ElseIf Additions = "B" Or Additions = "F" Or Additions = "H" Or Additions = "K" Then
2568               Frequency = 50
2569           Else
2570               Frequency = 0
2571           End If
2572       End If

2573       If Len(txtTEMCFrameNumber.Text) = 3 Then
2574           If txtTEMCFrameNumber.Text = "529" Then
2575               Frame = "420"
2576           Else
2577               Frame = Left$(txtTEMCFrameNumber, 2) & "0"
2578           End If
2579       Else
2580           Frame = txtTEMCFrameNumber.Text
2581           If Right$(txtTEMCFrameNumber.Text, 1) = "5" Then
2582               Frame = Frame & Left$(lblTEMCFrontRear.Caption, 1)
2583           Else
2584           End If
2585       End If
2586       Force = DLookupA(3, TEMCForceViscosity, 1, Frame)
2587       If Frequency = 60 Then
2588           Force = Force / 1.2
2589       End If
2590       If Val(txtTEMCViscosity.Text) > 1# Then
2591           If (Val(txtTEMCCalcForce.Text) > 3 * Force) Then
2592               lblTEMCPassFail.Visible = True
2593               lblTEMCPassFail.ForeColor = vbRed
2594               lblTEMCPassFail.Caption = "FAIL"
2595           Else
2596               lblTEMCPassFail.Visible = True
2597               lblTEMCPassFail.ForeColor = vbGreen
2598               lblTEMCPassFail.Caption = "PASS"
2599           End If
2600       End If

2601       If (Val(txtTEMCViscosity.Text) > 0.5) And (Val(txtTEMCViscosity.Text) <= 1#) Then
2602           If (Val(txtTEMCCalcForce.Text) > 2 * Force) Then
2603               lblTEMCPassFail.Visible = True
2604               lblTEMCPassFail.ForeColor = vbRed
2605               lblTEMCPassFail.Caption = "FAIL"
2606           Else
2607               lblTEMCPassFail.Visible = True
2608               lblTEMCPassFail.ForeColor = vbGreen
2609               lblTEMCPassFail.Caption = "PASS"
2610           End If
2611       End If

2612       If (Val(txtTEMCViscosity.Text) > 0.3) And (Val(txtTEMCViscosity.Text) <= 0.5) Then
2613           If (Val(txtTEMCCalcForce.Text) > 1.5 * Force) Then
2614               lblTEMCPassFail.Visible = True
2615               lblTEMCPassFail.ForeColor = vbRed
2616               lblTEMCPassFail.Caption = "FAIL"
2617           Else
2618               lblTEMCPassFail.Visible = True
2619               lblTEMCPassFail.ForeColor = vbGreen
2620               lblTEMCPassFail.Caption = "PASS"
2621           End If
2622       End If

2623       If (Val(txtTEMCViscosity.Text) <= 0.3) Then
2624           If (Val(txtTEMCCalcForce.Text) > 1# * Force) Then
2625               lblTEMCPassFail.Visible = True
2626               lblTEMCPassFail.ForeColor = vbRed
2627               lblTEMCPassFail.Caption = "FAIL"
2628           Else
2629               lblTEMCPassFail.Visible = True
2630               lblTEMCPassFail.ForeColor = vbGreen
2631               lblTEMCPassFail.Caption = "PASS"
2632           End If
2633       End If
2634       If NoOfPoles <> 0 Then
2635           VOverA = (DLookupA(2, TEMCForceViscosity, 1, Frame)) / (NoOfPoles * 30 / Frequency)
2636       End If
       '    If Frequency = 60 Then
       '        VOverA = VOverA * 1.2
       '    End If

2637       txtTEMCPVValue.Text = Format(Val(txtTEMCCalcForce.Text) * VOverA, "##0.0")

2638       If Val(txtTEMCFrontThrust.Text) = 0 And Val(txtTEMCRearThrust.Text) = 0 Then
2639           txtTEMCPVValue.Text = ""
2640           txtTEMCCalcForce.Text = ""
2641           lblTEMCPassFail.Visible = False
2642       End If


           'calculate reverse head
2643       txtRevHead.Text = Format(rsTestData.Fields("RBHPress") - rsTestData.Fields("SuctionPressure") * 2.31, "##0.0")
       '    txtRevHead.Text = Format((CDbl(Val(txtAI3Display.Text)) - CDbl(Val(txtSuctionDisplay.Text))) * 2.31, "##0.0")

' <VB WATCH>
2644       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2645       Exit Sub
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
2646       On Error GoTo vbwErrHandler
2647       Const VBWPROCNAME = "frmPLCData.updown2_change"
2648       If vbwProtector.vbwTraceProc Then
2649           Dim vbwProtectorParameterString As String
2650           If vbwProtector.vbwTraceParameters Then
2651               vbwProtectorParameterString = "()"
2652           End If
2653           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2654       End If
' </VB WATCH>
2655       Dim Plothead(1, 7) As Single
2656       Dim HeadPlot(7, 1) As Single

2657       Dim PlotEff() As Single
2658       Dim PlotKW() As Single
2659       Dim PlotAmps() As Single

2660       Dim j As Integer

2661       For j = 0 To UpDown2.value - 1
2662           Plothead(0, j) = HeadFlow(0, j)
2663           Plothead(1, j) = HeadFlow(1, j)
2664           HeadPlot(j, 0) = FlowHead(j, 0)
2665           HeadPlot(j, 1) = FlowHead(j, 1)
       '        ReDim Preserve PlotEff(1, j)
       '        PlotEff(0, j) = EffFlow(0, j)
       '        PlotEff(1, j) = EffFlow(1, j)
       '        ReDim Preserve PlotKW(1, j)
       '        PlotKW(0, j) = KWFlow(0, j)
       '        PlotKW(1, j) = KWFlow(1, j)
       '        ReDim Preserve PlotAmps(1, j)
       '        PlotAmps(0, j) = AmpsFlow(0, j)
       '        PlotAmps(1, j) = AmpsFlow(1, j)
2666       Next j

2667       MSChart1 = HeadPlot

' <VB WATCH>
2668       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2669       Exit Sub
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
2670       On Error GoTo vbwErrHandler
2671       Const VBWPROCNAME = "frmPLCData.DataGrid1_AfterColUpdate"
2672       If vbwProtector.vbwTraceProc Then
2673           Dim vbwProtectorParameterString As String
2674           If vbwProtector.vbwTraceParameters Then
2675               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ColIndex", ColIndex) & ") "
2676           End If
2677           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2678       End If
' </VB WATCH>
2679       DoEfficiencyCalcs
' <VB WATCH>
2680       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2681       Exit Sub
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
2682       On Error GoTo vbwErrHandler
2683       Const VBWPROCNAME = "frmPLCData.dgBalanceHoles_SelChange"
2684       If vbwProtector.vbwTraceProc Then
2685           Dim vbwProtectorParameterString As String
2686           If vbwProtector.vbwTraceParameters Then
2687               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Cancel", Cancel) & ") "
2688           End If
2689           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2690       End If
' </VB WATCH>
2691       If dgBalanceHoles.SelBookmarks.Count = 0 Then
2692           cmdModifyBalanceHoleData.Visible = False
2693       Else
2694           cmdModifyBalanceHoleData.Visible = True
2695       End If
' <VB WATCH>
2696       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2697       Exit Sub
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
2698       On Error GoTo vbwErrHandler
2699       Const VBWPROCNAME = "frmPLCData.Form_Activate"
2700       If vbwProtector.vbwTraceProc Then
2701           Dim vbwProtectorParameterString As String
2702           If vbwProtector.vbwTraceParameters Then
2703               vbwProtectorParameterString = "()"
2704           End If
2705           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2706       End If
' </VB WATCH>
2707       If ProgramEnd = True Then
2708           Unload Me
2709       End If
' <VB WATCH>
2710       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2711       Exit Sub
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
2712       On Error GoTo vbwErrHandler
2713       Const VBWPROCNAME = "frmPLCData.Form_Load"
2714       If vbwProtector.vbwTraceProc Then
2715           Dim vbwProtectorParameterString As String
2716           If vbwProtector.vbwTraceParameters Then
2717               vbwProtectorParameterString = "()"
2718           End If
2719           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2720       End If
' </VB WATCH>
2721       Dim RetVal As String
2722       Dim sSendStr As String
2723       Dim I As Integer
2724       Dim j As Integer
2725       Dim sTableName As String
2726       Dim WhichServer As String
2727       Dim WhichDatabase As String

2728       ProgramEnd = False
2729       Dim objWMIService As Object
2730       Dim colProcesses As Object
2731       Set objWMIService = GetObject("winmgmts:")
2732       Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name LIKE 'PolarRundown%'")
       '    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name LIKE 'Excel%'")
2733       If colProcesses.Count > 1 Then
2734           MsgBox "There is already a copy of Polar Rundown running.  You can only have one copy running at a time", vbOKOnly, "Polar Rundown already running"
2735           Dim f As Form
2736           For Each f In Forms
2737               If f.Name <> Me.Name Then
2738                    Unload f
2739               End If
2740           Next
2741           ProgramEnd = True
' <VB WATCH>
2742       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2743           Exit Sub
2744       Else
2745       End If
2746       Set objWMIService = Nothing
2747       Set colProcesses = Nothing

2748       debugging = 0   'assume not debugging
2749       WhichServer = "Production"     'change to production server
2750       WhichDatabase = "Production"

2751       If UCase$(Left$(GetMachineName, 5)) = "MROSE" Or UCase$(Left$(GetMachineName, 5)) = "ITTES" Then  'if mickey, see if we want to be in debug
2752           I = MsgBox("Debug?", vbYesNo)
2753           If I = vbYes Then
2754               debugging = 1
2755               WhichServer = "Production"
2756               WhichDatabase = "Production"
2757           Else
2758           End If
2759       End If

2760       If debugging Then
       '        GoTo temp
2761       End If
           'see if the mdb file is where it's supposed to be

2762       Dim developmentDatabase As String
2763       developmentDatabase = GetUNCFromLetter("F:") & sDevelopmentDatabase

2764       If Dir(developmentDatabase) = "" Then
2765           MsgBox "Development.mdb does not exist on F:, Please contact IT.", , "No Development Database"
2766           End
2767       End If

           'get the database info from the new mdb file
2768       Dim cnDevelopment As New ADODB.Connection
2769       Dim qyDevelopment As New ADODB.Command
2770       Dim rsDevelopment As New ADODB.Recordset

2771       On Error GoTo CannotConnect

2772       With cnDevelopment
2773           .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & developmentDatabase & ";Persist Security Info=False; Jet OLEDB:Database Password=Access7277word;"
2774           .ConnectionTimeout = 10
2775           .Open
2776       End With

2777   On Error GoTo vbwErrHandler
2778       GoTo Connected

2779   CannotConnect:
2780       MsgBox "Cannot connect with Development.mdb database.  Please contact IT.", , "Cannot find Connection data."
2781       End

2782   Connected:

           'we're connected, get the data for the Epicor SQL server
2783       qyDevelopment.CommandText = "SELECT * FROM Connections WHERE Connections.WhichServer = '" & WhichServer & "' AND WhichDatabase = '" & WhichDatabase & "'"
2784       qyDevelopment.ActiveConnection = cnDevelopment

2785       rsDevelopment.CursorLocation = adUseClient
2786       rsDevelopment.CursorType = adOpenStatic
2787       rsDevelopment.LockType = adLockOptimistic

2788       On Error GoTo NoServerData

2789       rsDevelopment.Open qyDevelopment

2790   On Error GoTo vbwErrHandler
2791       GoTo GotServerData

2792   NoServerData:

2793       MsgBox "Cannot connect with Development.mdb database.  Please contact IT.", , "Cannot find Connection data."
2794       End

2795   GotServerData:

2796       If rsDevelopment.RecordCount <> 1 Then
2797           GoTo NoServerData
2798       End If

           'construct Epicor connection string
2799       EpicorConnectionString = "Driver={" & rsDevelopment.Fields("ODBCDriver") & "};" & _
                                         "Database=" & rsDevelopment.Fields("DatabaseName") & ";" & _
                                         "Server=" & rsDevelopment.Fields("ServerName") & ";" & _
                                         "UID=" & rsDevelopment.Fields("UserName") & ";" & _
                                         "PWD=" & rsDevelopment.Fields("UserPassword") & ";"


           'make sure we can open the SQL database

2800       On Error GoTo CannotOpenEpicorSQLServer

2801       Dim cnTestEpicor As New ADODB.Connection
2802       cnTestEpicor.ConnectionString = EpicorConnectionString
2803       cnTestEpicor.Open
2804       cnTestEpicor.Close
2805       Set cnTestEpicor = Nothing
2806   On Error GoTo vbwErrHandler

2807       GoTo FoundEpicorSQLServer

2808   CannotOpenEpicorSQLServer:
2809       MsgBox "Cannot connect with the Epicor SQL server specified in Development.mdb.  Please contact IT.", , "Cannot connect with Epicor SQL Server"
2810       End

2811   FoundEpicorSQLServer:
           'get data on rundown database
2812       rsDevelopment.Close
2813       qyDevelopment.CommandText = "SELECT * FROM Connections WHERE Connections.WhichServer = 'PolarRundown'"

2814       On Error GoTo NoRundownDatabase

2815       rsDevelopment.Open qyDevelopment

2816       GoTo FoundRundownDatabase

2817   NoRundownDatabase:
2818       MsgBox "Cannot connect with the Pump Rundown database specified in Development.mdb.  Please contact IT.", , "Cannot connect with Epicor SQL Server"
2819       End

2820   FoundRundownDatabase:
2821       If rsDevelopment.RecordCount <> 1 Then
2822           GoTo NoRundownDatabase
2823           End
2824       End If

2825   temp:

2826       If debugging Then
2827           sDataBaseName = "c:\databases\PolarData.mdb"
2828       Else

2829          sDataBaseName = GetUNCFromLetter("F:") & "\Groups\Shared\databases\PolarData.mdb"

       '        sDataBaseName = rsDevelopment.Fields("ServerName") & rsDevelopment.Fields("DatabaseName")

       '        sDataBaseName = sServerName & "f\groups\shared\databases\PumpData 2k.mdb"
2830       End If

2831       Dim tempFSO As Object
2832       Set tempFSO = CreateObject("Scripting.FileSystemObject")
2833       ParentDirectoryName = tempFSO.getparentfoldername(sDataBaseName)
2834       Set tempFSO = Nothing

           'see if we can open the pump rundown database
2835       On Error GoTo NoRundownDatabase
2836       With cnPumpData
       '        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataBaseName & ";Persist Security Info=False;Jet OLEDB:Database Password=185TitusAve"
2837           .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataBaseName & ";Persist Security Info=False;"
2838           .ConnectionTimeout = 10
2839           .Open
2840       End With
2841   On Error GoTo vbwErrHandler


2842       If debugging = 0 Then
       '        Printer.Orientation = vbPRORLandscape
2843       End If

2844       lblVersion = "Polar Rundown - Version " & App.Major & "." & App.Minor & "." & App.Revision
2845       frmPLCData.Caption = "Polar Rundown"

2846       boFoundPump = False

2847       Me.Show

2848       MSChart1.Plot.Axis(VtChAxisIdX).AxisTitle = "Flow"
2849       MSChart1.Plot.Axis(VtChAxisIdY).AxisTitle = "TDH"
           'MSChart1.Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen = True
           'MSChart1.Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen = True
2850       MSChart1.Plot.UniformAxis = False
2851       MSChart1.Plot.SeriesCollection.Item(1).SeriesMarker.Auto = False
2852       MSChart1.Plot.SeriesCollection.Item(1).Pen.Width = 5
2853       With MSChart1.Plot.SeriesCollection.Item(1).DataPoints.Item(-1).Marker
2854           .Visible = True
2855           .Size = 50
2856           .Style = VtMarkerStyleCircle
2857           .FillColor.Automatic = False
2858           .FillColor.Set 0, 0, 255
2859       End With
2860       MSChart1.Plot.AutoLayout = False
2861       MSChart1.Plot.LocationRect.Max.x = 5600
2862       MSChart1.Plot.LocationRect.Max.y = 2800
2863       MSChart1.Plot.LocationRect.Min.x = 0
2864       MSChart1.Plot.LocationRect.Min.y = 0

           'assure that the timers are off
2865       frmPLCData.tmrGetDDE.Enabled = False

2866       frmPLCData.tmrStartUp.Enabled = False

           'initialize the PLC network
2867       RetVal = NetWorkInitialize()
2868       If RetVal <> 0 Then
2869           MsgBox ("Can't Initialize Network. Exiting...")
2870           End
2871       End If

2872       If debugging = 0 Then
               'load array of plcs
2873           I = 0
2874           Open rsDevelopment.Fields("ServerName") & "PolarPLCAddresses.txt" For Input As 1
2875           While Not EOF(1)
2876               Input #1, Description(I)
2877               For j = 0 To 125
2878                   Input #1, aDevices(I).Address(j)
2879               Next j
2880               Input #1, j
2881               I = I + 1
2882           Wend
2883           Close #1

2884           DeviceCount = I

2885           If Left$(GetMachineName, 2) = "WV" Then  'if in WV, put MWSC first in loop dropdown
2886               Dim k As Integer
2887               For k = 0 To DeviceCount - 1
2888                   If InStr(Description(k), "MWSC") <> 0 Then
2889                       Exit For
2890                   End If
2891               Next k
2892               Description(DeviceCount) = Description(0)
2893               Description(0) = Description(k)
2894               Description(k) = Description(DeviceCount)

2895               aDevices(DeviceCount) = aDevices(0)
2896               aDevices(0) = aDevices(k)
2897               aDevices(k) = aDevices(DeviceCount)

2898           End If

2899           Dim PLCAddress As String
2900           For I = 0 To DeviceCount - 1
2901               PLCAddress = aDevices(I).Address(4) & "." & aDevices(I).Address(5) & "." & aDevices(I).Address(6) & "." & aDevices(I).Address(7)
2902               RetVal = PingSilent(PLCAddress)
2903               If RetVal <> 0 Then
2904                   frmPLCData.cmbPLCLoop.AddItem Description(I)
2905                   frmPLCData.cmbPLCLoop.ItemData(frmPLCData.cmbPLCLoop.NewIndex) = I
2906               End If
2907           Next I
2908       End If

2909       frmPLCData.cmbPLCLoop.AddItem "Add PLC Data Manually"   'enable the controls for manual entry

           'turn on the PLC led

2910       frmPLCData.cmbPLCLoop.ListIndex = 0
2911       frmPLCData.tmrGetDDE.Enabled = True

           'hook up to the various databases

           'copy the template of the database here
           'see if it exists
2912       Dim fdrive As String
2913       fdrive = GetUNCFromLetter("F:")
2914       If Dir(fdrive & "\groups\shared\databases" & sEffDataBaseName) = "" Then
2915           MsgBox "File does not exist at " & fdrive & "\groups\shared\databases" & sEffDataBaseName & ". Please contact IT", vbOKOnly, "Eff.mdb does not exist"
2916       Else
               'Dim FSO As New FileSystemObject
2917           FileCopy fdrive & "\groups\shared\databases" & sEffDataBaseName, App.Path & sEffDataBaseName
2918       End If


2919       With cnEffData
2920           .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & sEffDataBaseName & ";Persist Security Info=False"
2921           .Open
2922       End With

           'open some recordsets
2923       rsPumpData.Index = "SerialNumber"
2924       rsTestSetup.Index = "FindData"
2925       rsTestData.Index = "PrimaryKey"
2926       rsPumpData.Open "TempPumpData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
2927       rsTestSetup.Open "TempTestSetupData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
2928       rsTestData.Filter = "SerialNumber = ''"
2929       rsTestData.CursorLocation = adUseClient
2930       rsTestData.Open "TempTestData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
2931       rsEff.CursorLocation = adUseClient
2932       rsEff.Open "Efficiency", cnEffData, adOpenStatic, adLockOptimistic, adCmdTableDirect
2933       qyBalanceHoles.ActiveConnection = cnPumpData
2934       rsBalanceHoles.CursorLocation = adUseClient
2935       rsBalanceHoles.CursorType = adOpenStatic
2936       rsBalanceHoles.LockType = adLockOptimistic
2937       qyMisc.ActiveConnection = cnPumpData
2938       qyMisc.CommandText = "SELECT MiscParameters.ParameterName, MiscParameters.ParameterValue From MiscParameters WHERE (((MiscParameters.ParameterName)='AllowableTDHVariation'));"
2939       rsMisc.CursorLocation = adUseClient
2940       rsMisc.CursorType = adOpenStatic
2941       rsMisc.LockType = adLockBatchOptimistic
2942       rsMisc.Open qyMisc
2943       txtNPSH(5).Text = rsMisc!ParameterValue

2944       If debugging <> 1 Then
2945           FindMagtrols
2946       Else
2947           cmbMagtrol.AddItem "Add Manually"
2948           cmbMagtrol.ItemData(cmbMagtrol.NewIndex) = 99
2949           cmbMagtrol.ListIndex = 0
2950       End If
2951       optKW(1).value = True
2952       optKW_Click (1)


           'blank out data grid
2953       Set DataGrid1.DataSource = rsTestData

           'load the combo boxes
2954       LoadCombo cmbStatorFill, "StatorFill"
2955       LoadCombo cmbCirculationPath, "CirculationPath"
2956       LoadCombo cmbVoltage, "Voltage"
2957       LoadCombo cmbFrequency, "Frequency"
2958       LoadCombo cmbMotor, "Motor"
2959       LoadCombo cmbDesignPressure, "DesignPressure"
2960       LoadCombo cmbRPM, "RPM"
2961       LoadCombo cmbOrificeNumber, "OrificeNumber"
2962       LoadCombo cmbTestSpec, "TestSpecification"
2963       LoadCombo cmbLoopNumber, "LoopNumber"
2964       LoadCombo cmbSuctDia, "SuctionDiameter"
2965       LoadCombo cmbDischDia, "DischargeDiameter"
2966       LoadCombo cmbTachID, "TachID"
2967       LoadCombo cmbAnalyzerNo, "AnalyzerNo"
2968       LoadCombo cmbModel, "Model"
2969       LoadCombo cmbModelGroup, "ModelGroup"
2970       LoadCombo cmbMounting, "Mounting"
2971       LoadCombo cmbPLCNo, "PLCNo"
2972       LoadCombo cmbFlowMeter, "PumpFlowMeter"
2973       LoadCombo cmbSuctionPressureTransducer, "SuctionPressureTransducer"
2974       LoadCombo cmbDischargePressureTransducer, "DischargePressureTransducer"
2975       LoadCombo cmbTemperatureTransducer, "TemperatureTransducer"
2976       LoadCombo cmbCirculationFlowMeter, "CirculationFlowMeter"
           'LoadCombo cmbSupermarketModel, "SupermarketPumpData"

2977       SetFrequencyCombo
           'load the TEMC combo boxes, too
2978       LoadCombo cmbTEMCAdapter, "TEMCAdapter"
2979       LoadCombo cmbTEMCAdditions, "TEMCAdditions"
2980       LoadCombo cmbTEMCCirculation, "TEMCCirculation"
2981       LoadCombo cmbTEMCDesignPressure, "TEMCDesignPressure"
2982       LoadCombo cmbTEMCNominalDischargeSize, "TEMCNominalDischargeSize"
2983       LoadCombo cmbTEMCDivisionType, "TEMCDivisionType"
2984       LoadCombo cmbTEMCImpellerType, "TEMCImpellerType"
2985       LoadCombo cmbTEMCInsulation, "TEMCInsulation"
2986       LoadCombo cmbTEMCJacketGasket, "TEMCJacketGasket"
2987       LoadCombo cmbTEMCMaterials, "TEMCMaterials"
2988       LoadCombo cmbTEMCModel, "TEMCModel"
2989       LoadCombo cmbTEMCNominalImpSize, "TEMCNominalImpSize"
2990       LoadCombo cmbTEMCOtherMotor, "TEMCOtherMotor"
2991       LoadCombo cmbTEMCNominalSuctionSize, "TEMCNominalSuctionSize"
2992       LoadCombo cmbTEMCVoltage, "TEMCVoltage"
2993       LoadCombo cmbTEMCPumpStages, "TEMCPumpStages"
2994       LoadCombo cmbTEMCTRG, "TEMCTRG"

           'LoadCombo frmSearch.cmbSearchModel, "Model"

           'fill memory arrays for dlookups
2995       FillArrays

           'choose the first tab
2996       frmPLCData.SSTab1.Tab = 0

           'set the grid column names
2997       Dim c As Column
2998       For Each c In DataGrid1.Columns
2999           Select Case c.DataField
               Case "TestDataID"
3000               c.Visible = False
3001           Case "SerialNumber"
3002               c.Visible = False
3003           Case "Date"
3004               c.Visible = False
3005           Case Else ' Show all other columns.
3006               c.Visible = True
3007               c.Alignment = dbgRight
3008           End Select
3009       Next c

3010       Set dgBalanceHoles.DataSource = rsBalanceHoles

3011       For Each c In dgBalanceHoles.Columns
3012           Select Case c.DataField
               Case "BalanceHoleID"
3013               c.Visible = False
3014           Case "SerialNo"
3015               c.Visible = False
3016           Case "Date"
3017               c.Visible = True
3018               c.Alignment = dbgCenter
3019               c.Width = 2000
3020           Case "Number"
3021               c.Visible = True
3022               c.Alignment = dbgCenter
3023               c.Width = 700
3024           Case "Diameter"
3025               c.Visible = False
3026           Case "Diameter1"
3027               c.Caption = "Diameter"
3028               c.Visible = True
3029               c.Alignment = dbgCenter
3030               c.Width = 700
3031           Case "BoltCircle1"
3032               c.Caption = "Bolt Circle"
3033               c.Visible = True
3034               c.Alignment = dbgCenter
3035               c.Width = 800
3036           Case "BoltCircle"
3037               c.Visible = False
3038           Case "SetNo"
3039               c.Visible = False
3040           Case Else ' Show all other columns.
3041               c.Visible = False
3042           End Select
3043       Next c

3044       BlankData

       '    If debugging <> 1 Then
               'get user initials
3045           frmLogin.Show
       '    End If

3046     optMfr(1).value = True
3047     frmMfr.Visible = False

3048       Pressed = True
' <VB WATCH>
3049       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3050       Exit Sub
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
3051       On Error GoTo vbwErrHandler
3052       Const VBWPROCNAME = "frmPLCData.Form_Unload"
3053       If vbwProtector.vbwTraceProc Then
3054           Dim vbwProtectorParameterString As String
3055           If vbwProtector.vbwTraceParameters Then
3056               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Cancel", Cancel) & ") "
3057           End If
3058           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3059       End If
' </VB WATCH>
3060       End
' <VB WATCH>
3061       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3062       Exit Sub
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
3063       On Error GoTo vbwErrHandler
3064       Const VBWPROCNAME = "frmPLCData.Label15_Click"
3065       If vbwProtector.vbwTraceProc Then
3066           Dim vbwProtectorParameterString As String
3067           If vbwProtector.vbwTraceParameters Then
3068               vbwProtectorParameterString = "()"
3069           End If
3070           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3071       End If
' </VB WATCH>
3072       frmDiagram.Show
' <VB WATCH>
3073       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3074       Exit Sub
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
3075       On Error GoTo vbwErrHandler
3076       Const VBWPROCNAME = "frmPLCData.lblAutoMan_Click"
3077       If vbwProtector.vbwTraceProc Then
3078           Dim vbwProtectorParameterString As String
3079           If vbwProtector.vbwTraceParameters Then
3080               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3081           End If
3082           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3083       End If
' </VB WATCH>

3084       Dim blnEnabled As Boolean

3085       If lblAutoMan(Index).Caption = "Auto" Then
3086           lblAutoMan(Index).Caption = "Man"
3087           blnEnabled = True
3088       Else
3089           lblAutoMan(Index).Caption = "Auto"
3090           blnEnabled = False
3091       End If

3092       Select Case Index
               Case 0
3093               txtFlowDisplay.Enabled = blnEnabled
3094           Case 1
3095               txtSuctionDisplay.Enabled = blnEnabled
3096           Case 2
3097               txtDischargeDisplay.Enabled = blnEnabled
3098           Case 3
3099               txtTemperatureDisplay.Enabled = blnEnabled
3100           Case 4
3101               txtAI1Display.Enabled = blnEnabled
3102           Case 5
3103               txtAI2Display.Enabled = blnEnabled
3104           Case 6
3105               txtAI3Display.Enabled = blnEnabled
3106           Case 7
3107               txtAI4Display.Enabled = blnEnabled
3108       End Select

' <VB WATCH>
3109       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3110       Exit Sub
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
3111       On Error GoTo vbwErrHandler
3112       Const VBWPROCNAME = "frmPLCData.tmrNPSHr_Timer"
3113       If vbwProtector.vbwTraceProc Then
3114           Dim vbwProtectorParameterString As String
3115           If vbwProtector.vbwTraceParameters Then
3116               vbwProtectorParameterString = "()"
3117           End If
3118           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3119       End If
' </VB WATCH>
3120       tmrNPSHr.Enabled = False
3121       If frmNPSH.Visible = True Then
3122           btnRunNPSH_Click    'close test
3123       End If
' <VB WATCH>
3124       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3125       Exit Sub
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
3126       On Error GoTo vbwErrHandler
3127       Const VBWPROCNAME = "frmPLCData.txtNPSH_Change"
3128       If vbwProtector.vbwTraceProc Then
3129           Dim vbwProtectorParameterString As String
3130           If vbwProtector.vbwTraceParameters Then
3131               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3132           End If
3133           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3134       End If
' </VB WATCH>
3135       If Index = 5 Then
3136           If frmNPSH.Visible = True Then
3137               If rsMisc.State = adStateOpen Then
3138                   rsMisc.Close
3139               End If
3140               rsMisc.CursorLocation = adUseClient
3141               rsMisc.Open "Select * from MiscParameters WHERE (ParameterName = 'AllowableTDHVariation');", cnPumpData, adOpenStatic, adLockOptimistic, adCmdText
3142               rsMisc.Fields("ParameterValue").value = txtNPSH(5).Text
3143               rsMisc.Update
3144           End If
3145       End If
' <VB WATCH>
3146       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3147       Exit Sub
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
3148       On Error GoTo vbwErrHandler
3149       Const VBWPROCNAME = "frmPLCData.txtNPSHFileLocation_Click"
3150       If vbwProtector.vbwTraceProc Then
3151           Dim vbwProtectorParameterString As String
3152           If vbwProtector.vbwTraceParameters Then
3153               vbwProtectorParameterString = "()"
3154           End If
3155           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3156       End If
' </VB WATCH>
3157       Dim sTempDir As String
3158       On Error Resume Next
3159       sTempDir = CurDir    'Remember the current active directory
3160       CommonDialog2.DialogTitle = "Select a directory" 'titlebar
3161       CommonDialog2.InitDir = "\\tei-main-01\f\en\groups\shared\calibration and rundown\npsh\" 'start dir, might be "C:\" or so also
3162       CommonDialog2.filename = "Select a Directory"  'Something in filenamebox
3163       CommonDialog2.Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
3164       CommonDialog2.Filter = "Directories|*.~#~" 'set files-filter to show dirs only
3165       CommonDialog2.CancelError = True 'allow escape key/cancel
3166       CommonDialog2.ShowSave   'show the dialog screen

3167       If Err <> 32755 Then    ' User didn't chose Cancel.
               'Me.SDir.Text = CurDir
3168       End If

       '    ChDir sTempDir  'restore path to what it was at entering

3169   Me.txtNPSHFileLocation.Text = CommonDialog2.filename

' <VB WATCH>
3170       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3171       Exit Sub
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
3172       On Error GoTo vbwErrHandler
3173       Const VBWPROCNAME = "frmPLCData.txtTitle_LostFocus"
3174       If vbwProtector.vbwTraceProc Then
3175           Dim vbwProtectorParameterString As String
3176           If vbwProtector.vbwTraceParameters Then
3177               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3178           End If
3179           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3180       End If
' </VB WATCH>

3181       ChangeTitles Index

' <VB WATCH>
3182       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3183       Exit Sub
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
3184       On Error GoTo vbwErrHandler
3185       Const VBWPROCNAME = "frmPLCData.ChangeTitles"
3186       If vbwProtector.vbwTraceProc Then
3187           Dim vbwProtectorParameterString As String
3188           If vbwProtector.vbwTraceParameters Then
3189               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ChannelNo", ChannelNo) & ") "
3190           End If
3191           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3192       End If
' </VB WATCH>
3193       Dim I As Integer
3194       Dim S As String

3195       If txtTitle(ChannelNo).Locked = True Then
' <VB WATCH>
3196       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3197           Exit Sub
3198       End If

3199       Dim qy As New ADODB.Command
3200       Dim rs As New ADODB.Recordset

3201       qy.ActiveConnection = cnPumpData

           'see if we have an entry in the table
3202       qy.CommandText = "SELECT * FROM AITitles " & _
                             "WHERE (((AITitles.SerialNo)= '" & txtSN.Text & "') " & _
                             "AND ((AITitles.Date)= #" & cmbTestDate.Text & "#) " & _
                             "AND ((AITitles.Channel)=" & ChannelNo & "));"

3203       With rs     'open the recordset for the query
3204           .CursorLocation = adUseClient
3205           .CursorType = adOpenStatic
3206           .LockType = adLockOptimistic
3207           .Open qy
3208       End With

3209       If (rs.BOF = True And rs.EOF = True) Then  'new record
3210           rs.AddNew
3211           rs.Fields("SerialNo") = txtSN.Text
3212           rs.Fields("Date") = cmbTestDate.Text
3213           rs.Fields("Channel") = CByte(ChannelNo)
3214           rs.Fields("Title") = txtTitle(ChannelNo).Text
3215           rs.Update
3216       Else    'we have an entry, modify it
3217           rs.Fields("SerialNo") = txtSN.Text
3218           rs.Fields("Date") = cmbTestDate.Text
3219           rs.Fields("Channel") = CByte(ChannelNo)
3220           rs.Fields("Title") = txtTitle(ChannelNo).Text
3221           rs.Update
3222       End If

3223       rs.Close
3224       Set rs = Nothing
3225       Set qy = Nothing

' <VB WATCH>
3226       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3227       Exit Sub
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
3228       On Error GoTo vbwErrHandler
3229       Const VBWPROCNAME = "frmPLCData.optKW_Click"
3230       If vbwProtector.vbwTraceProc Then
3231           Dim vbwProtectorParameterString As String
3232           If vbwProtector.vbwTraceParameters Then
3233               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3234           End If
3235           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3236       End If
' </VB WATCH>
3237       Select Case Index
               Case 0  'add 3 powers
3238               txtKW.Enabled = False
3239           Case 1  'enter kw
3240               txtKW.Enabled = True
3241           Case 2  'use analog in 4
3242               txtKW.Enabled = False
3243       End Select
' <VB WATCH>
3244       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3245       Exit Sub
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
3246       On Error GoTo vbwErrHandler
3247       Const VBWPROCNAME = "frmPLCData.optMfr_Click"
3248       If vbwProtector.vbwTraceProc Then
3249           Dim vbwProtectorParameterString As String
3250           If vbwProtector.vbwTraceParameters Then
3251               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3252           End If
3253           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3254       End If
' </VB WATCH>
3255       frmTEMC.Visible = optMfr(1).value
3256       frmChempump.Visible = optMfr(0).value
3257       frmTEMCData.Visible = optMfr(1).value
3258       txtModelNo_Change
' <VB WATCH>
3259       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3260       Exit Sub
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
3261       On Error GoTo vbwErrHandler
3262       Const VBWPROCNAME = "frmPLCData.tmrGetDDE_Timer"
3263       If vbwProtector.vbwTraceProc Then
3264           Dim vbwProtectorParameterString As String
3265           If vbwProtector.vbwTraceParameters Then
3266               vbwProtectorParameterString = "()"
3267           End If
3268           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3269       End If
' </VB WATCH>

       'get here every second... get plc and magtrol data

3270       Dim sSendStr As String
3271       Dim I As Integer
3272       Dim VoltMul As Double

3273       If Calibrating Then
' <VB WATCH>
3274       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3275           Exit Sub
3276       End If

3277       If debugging Then
               'Exit Sub
3278       End If


3279       If boPLCOperating = True Then
3280           frmPLCData.shpGetPLCData.Visible = True    'turn the PLC led on

               'convert the plc data into real numbers
               'the following data are type real
3281           txtFlow.Text = ConvertToReal("4050")
3282           txtSuction.Text = ConvertToReal("4052")
3283           txtDischarge.Text = ConvertToReal("4054")
3284           txtTemperature.Text = ConvertToReal("4056")

3285           txtValvePosition.Text = ConvertToLong("2004")

3286           frmPLCData.txtTC1.Text = ConvertToLong("2200")
3287           frmPLCData.txtTC2.Text = ConvertToLong("2202")
3288           frmPLCData.txtTC3.Text = ConvertToLong("2204")
3289           frmPLCData.txtTC4.Text = ConvertToLong("2206")

3290           frmPLCData.txtAI1.Text = ConvertToReal("4060")
3291           frmPLCData.txtAI2.Text = ConvertToReal("4062")
3292           frmPLCData.txtAI3.Text = ConvertToReal("4064")
3293           frmPLCData.txtAI4.Text = ConvertToReal("4066")

3294           frmPLCData.txtPCoef.Text = ConvertToLong("4036")
3295           frmPLCData.txtICoef.Text = ConvertToLong("4037")
3296           frmPLCData.txtDCoef.Text = ConvertToLong("4040")

3297           frmPLCData.txtSetPoint.Text = ConvertToLong("4035")
3298           frmPLCData.txtInHg.Text = ConvertToLong("1460")


               'modify the data from PLC format to format that we can use
               'and update the screen
3299           If txtFlowDisplay.Enabled = False Then
3300               frmPLCData.txtFlowDisplay = Format$(txtFlow.Text, "###0.00")
3301           End If
3302           If txtSuctionDisplay.Enabled = False Then
3303               frmPLCData.txtSuctionDisplay = Format$((txtSuction.Text) / 10, "##0.00")
3304           End If
3305           If txtDischargeDisplay.Enabled = False Then
3306               frmPLCData.txtDischargeDisplay = Format$(txtDischarge.Text, "##0.00")
3307           End If
3308           If txtTemperatureDisplay.Enabled = False Then
3309               frmPLCData.txtTemperatureDisplay = Format$(txtTemperature.Text, "##0.00")
3310           End If
3311           frmPLCData.txtValvePositionDisplay = (txtValvePosition.Text)

3312           frmPLCData.txtTC1Display = Format$((txtTC1.Text) / 10, "##0.0")
3313           frmPLCData.txtTC2Display = Format$((txtTC2.Text) / 10, "##0.0")
3314           frmPLCData.txtTC3Display = Format$((txtTC3.Text) / 10, "##0.0")
3315           frmPLCData.txtTC4Display = Format$((txtTC4.Text) / 10, "##0.0")

3316           If txtAI1Display.Enabled = False Then
3317               frmPLCData.txtAI1Display = Format$(txtAI1.Text, "##0.00")
3318           End If
3319           If txtAI2Display.Enabled = False Then
3320               frmPLCData.txtAI2Display = Format$(txtAI2.Text, "##0.00")
3321           End If
3322           If txtAI3Display.Enabled = False Then
3323               frmPLCData.txtAI3Display = Format$(txtAI3.Text, "##0.00")
3324           End If
3325           If txtAI4Display.Enabled = False Then
3326               frmPLCData.txtAI4Display = Format$(txtAI4.Text, "##0.00")
3327           End If

3328           frmPLCData.txtSetPointDisplay = (txtSetPoint.Text)

3329           frmPLCData.txtInHgDisplay = Format$(txtInHg.Text / 100, "00.00")

3330           frmPLCData.shpGetPLCData.Visible = False   'turn the PLC led off

3331           frmPLCData.shpGetMagtrolData.Visible = True 'turn the Magtrol led on
3332       End If

3333       If boMagtrolOperating = True Then


               'get the data from the Magtrol
3334           If Right(cmbMagtrol.List(cmbMagtrol.ListIndex), 4) = "5300" Then
3335               sSendStr = vbCrLf
3336               sData = Space$(68)
3337               VoltMul = Sqr(3)
3338           Else
3339               sSendStr = "OT" & vbCrLf
3340               sData = Space$(183)
3341               VoltMul = 1#
3342           End If

3343           On Error GoTo noresponse
3344           If UsingNatInst Then
3345               ibwrt iUD, sSendStr
3346               ibrd iUD, sData

                   'parse the Magrol response
       '            vResponse = CWGPIB1.Tasks("Number Parser").Parse(sData)
3347           Else
                   'Dim Databack As String
3348               sData = TCP.SendGetData("OT")
3349           End If

3350               Dim vSplit() As String
3351               vSplit = Split(Right(sData, Len(sData) - 1), ",")
3352               If UBound(vSplit) > 0 Then
3353                  ReDim vResponse(UBound(vSplit))
3354               End If
3355               For I = 0 To UBound(vSplit) - 1
3356                   If Len(vSplit(I)) <> 0 Then
3357                       vResponse(I) = CDbl(vSplit(I))
3358                   End If
3359               Next I

               'format the parsed response
3360           Dim dd As String
3361           dd = "- -"

3362           If Not IsEmpty(vResponse) Then
               '8 entries for 5300 and 12 for the 6530
3363               If UBound(vResponse) = 8 Or UBound(vResponse) = 12 Then
                       'put the responses into the correct text box
3364                   txtV1.Text = Format$(VoltMul * vResponse(1), "###0.0")   'we get back phase voltage and we want line voltage

3365                   Select Case vResponse(0)
                           Case Is < 1
3366                           txtI1.Text = Format$(vResponse(0), "0.0000")
3367                       Case Is < 10
3368                           txtI1.Text = Format$(vResponse(0), "0.000")
3369                       Case Is < 100
3370                           txtI1.Text = Format$(vResponse(0), "00.00")
3371                       Case Else
3372                           txtI1.Text = Format$(vResponse(0), "000.0")
3373                   End Select

3374                   Select Case vResponse(3)
                           Case Is < 1
3375                           txtI2.Text = Format$(vResponse(3), "0.0000")
3376                       Case Is < 10
3377                           txtI2.Text = Format$(vResponse(3), "0.000")
3378                       Case Is < 100
3379                           txtI2.Text = Format$(vResponse(3), "00.00")
3380                       Case Else
3381                           txtI2.Text = Format$(vResponse(3), "000.0")
3382                   End Select

3383                   Select Case vResponse(6)
                           Case Is < 1
3384                           txtI3.Text = Format$(vResponse(6), "0.0000")
3385                       Case Is < 10
3386                           txtI3.Text = Format$(vResponse(6), "0.000")
3387                       Case Is < 100
3388                           txtI3.Text = Format$(vResponse(6), "00.00")
3389                       Case Else
3390                           txtI3.Text = Format$(vResponse(6), "000.0")
3391                   End Select

3392                   txtP1.Text = Format$(vResponse(2) / 1000, "##0.00")     '/ by 1000 to show kW
3393                   txtV2.Text = Format$(VoltMul * vResponse(4), "###0.0")
                       'txtI2.Text = Format$(vResponse(3), "###0.0")
3394                   txtP2.Text = Format$(vResponse(5) / 1000, "##0.00")
3395                   txtV3.Text = Format$(VoltMul * vResponse(7), "###0.0")
                       'txtI3.Text = Format$(vResponse(6), "###0.0")
3396                   txtP3.Text = Format$(vResponse(8) / 1000, "##0.00")
3397                   If (vResponse(0) * vResponse(1) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7)) <> 0 Then
                           'if we have some measured current
                           'pf = sum of power/sum of VA
3398                       If Right(cmbMagtrol.List(cmbMagtrol.ListIndex), 4) = "5300" Then
                               'add kw responses and / by 1000 to get to kW
3399                           txtKW.Text = (vResponse(2) + vResponse(5) + vResponse(8)) / 1000
3400                           txtPF.Text = Format$(100 * (vResponse(2) + vResponse(5) + vResponse(8)) / (vResponse(1) * vResponse(0) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7)), "0.00")
3401                       Else
3402                           txtKW.Text = (vResponse(2) + vResponse(8)) / 1000
3403                           txtPF.Text = Format$(100 * (vResponse(2) + vResponse(8)) / ((Sqr(3) / 3) * (vResponse(1) * vResponse(0) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7))), "0.00")
3404                       End If
3405                       Select Case Val(txtKW.Text)
                               Case Is < 1
3406                               txtKW.Text = Format$(txtKW.Text, "0.00000")
3407                           Case Is < 10
3408                               txtKW.Text = Format$(txtKW.Text, "0.0000")
3409                           Case Is < 100
3410                               txtKW.Text = Format$(txtKW.Text, "00.000")
3411                           Case Else
3412                               txtKW.Text = Format$(txtKW.Text, "000.00")
3413                       End Select
3414                   Else
3415                       txtPF = dd
3416                   End If
3417               Else
                       'no response, show all -- in text boxes
3418                   txtV1.Text = dd
3419                   txtI1.Text = dd
3420                   txtP1.Text = dd
3421                   txtV2.Text = dd
3422                   txtI2.Text = dd
3423                   txtP2.Text = dd
3424                   txtV3.Text = dd
3425                   txtI3.Text = dd
3426                   txtP3.Text = dd
3427                   txtPF = dd
3428                   txtKW = dd
3429               End If
3430           End If
3431       Else    'magtrol not operating
3432           Dim dbl As Double

3433           If optKW(0).value = True Then   'add 3 powers
3434               txtKW.Text = Val(txtP1.Text) + Val(txtP2.Text) + Val(txtP3.Text)
3435           End If
3436           If optKW(1).value = True Then   'enter kw
3437               txtP1.Text = Val(txtKW.Text) / 3
3438               txtP2.Text = Val(txtKW.Text) / 3
3439               txtP3.Text = Val(txtKW.Text) / 3
3440           End If
3441           If optKW(2).value = True Then   'use ai4
3442               txtKW.Text = txtAI4Display.Text
3443               txtP1.Text = Val(txtKW.Text) / 3
3444               txtP2.Text = Val(txtKW.Text) / 3
3445               txtP3.Text = Val(txtKW.Text) / 3
3446           End If

3447           dbl = Val(txtV1.Text) * Val(txtI1.Text)
3448           dbl = dbl + Val(txtV2.Text) * Val(txtI2.Text)
3449           dbl = dbl + Val(txtV3.Text) * Val(txtI3.Text)
3450           If dbl <> 0 Then
3451               txtPF.Text = Format$((Val(txtKW.Text) * 1000 * 3 * 100 / (dbl * Sqr(3))), "0.00")
3452           End If
3453       End If

3454   noresponse:
3455   On Error GoTo vbwErrHandler
3456       frmPLCData.shpGetMagtrolData.Visible = False   'turn the Magtrol led off

           'update the little PLC chart
3457       For I = 1 To 99
3458           vPlot(0, I) = vPlot(0, I + 1)
3459           vPlot(1, I) = vPlot(1, I + 1)
3460       Next I
3461       vPlot(0, 100) = txtSetPointDisplay
3462       vPlot(1, 100) = txtFlowDisplay

           'do NPSH stuff
3463       Dim SuctVelHead As Single
3464       Dim DischVelHead As Single
3465       Dim Conversion As Single
3466       Dim SuctionPSIA As Single
3467       Dim DischargePSIA As Single
3468       Dim VaporPress As Single
3469       Dim SpecVolume As Single
3470       Dim NPSHa As Single
3471       Dim NPSHr As Single
3472       Dim TDH As Single
3473       Dim pd As Single


           'velocity head
3474       If cmbSuctDia.ListIndex = -1 Then   'if no suction diameter chosen
3475           SuctVelHead = 0
3476       Else
       '        pd = DLookup("ActualDia", "PipeDiameters", "ID = " & cmbSuctDia.ListIndex + 1)
3477           pd = DLookupA(ActualColNo, PipeDiameters, IDColNo, cmbSuctDia.ItemData(cmbSuctDia.ListIndex) + 1)
3478           SuctVelHead = (0.002592 * Val(txtFlow) ^ 2) / (pd ^ 4)
3479       End If

3480       If cmbDischDia.ListIndex = -1 Then     'if no discharge diameter chosen
3481           DischVelHead = 0
3482       Else
       '        pd = DLookup("ActualDia", "PipeDiameters", "ID = " & cmbDischDia.ListIndex + 1)
3483           pd = DLookupA(ActualColNo, PipeDiameters, IDColNo, cmbDischDia.ItemData(cmbDischDia.ListIndex) + 1)
3484           DischVelHead = (0.002592 * Val(txtFlow) ^ 2) / (pd ^ 4)
3485       End If

           'convert gauges to absolute
3486       If txtInHgDisplay.Text = "" Then
3487           Conversion = 0
3488       Else
3489           Conversion = txtInHgDisplay * 0.491
3490       End If

3491       SuctionPSIA = Val(txtSuctionDisplay) + Conversion
3492       DischargePSIA = Val(txtDischargeDisplay) + Conversion


           'lookup vapor pressure and specific volume in the arrays that we made
           'if temp is out of range, say so and exit
3493       If Val(txtTemperatureDisplay) < 40 Or Val(txtTemperatureDisplay) > 165 Then
3494           txtNPSHa = 0
' <VB WATCH>
3495       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3496           Exit Sub
3497       Else
3498           I = Val(txtTemperatureDisplay) - 40
       '        VaporPress = DLookup("VaporPressure", "VaporPressure", "ID = " & I)
       '        SpecVolume = DLookup("SpecificVolume", "VaporPressure", "ID= " & I)
3499           VaporPress = DLookupA(VaporPressureColNo, VaporPressure, IDColNo, I)
3500           SpecVolume = DLookupA(SpecificVolumeColNo, VaporPressure, IDColNo, I)
3501       End If

3502       If Not ((txtSuctHeight = "") Or (txtDischHeight = "") Or Not IsNumeric(txtSuctHeight) Or Not IsNumeric(txtDischHeight)) Then
               'NPSHa
3503           NPSHa = (144 * SpecVolume * (SuctionPSIA - VaporPress)) + (txtSuctHeight / 12) + SuctVelHead
       '        NPSHa = CalcTDH(DischargePSIA, SuctionPSIA, 0, DischVelHead, 0, txtTemperature)
3504           txtNPSHa = Format$(NPSHa, "##0.00")

               'tdh
3505           TDH = CalcTDH(DischargePSIA, SuctionPSIA, 0, (DischVelHead - SuctVelHead), (txtDischHeight / 12) - (txtSuctHeight / 12), txtTemperatureDisplay)
3506           txtTDH = Format$(TDH, "##0.00")

3507           If frmNPSH.Visible = True Then
3508               If Val(txtTDH.Text) > 0 Then
3509                   txtNPSH(2).Text = Format(100 * Val(txtTDH.Text) / Val(txtNPSH(3).Text), "##0.00")
3510                   txtNPSH(1).Text = Format(100 * Val(txtFlow.Text) / Val(txtNPSH(0).Text), "##0.00")
                       'check for tdh variation
3511                   If Abs(Val(txtNPSH(1)) - 100) > Val(txtNPSH(5).Text) Then
3512                       MsgBox "The TDH value has varied more than " & txtNPSH(5) & " %. NPSHr data will NOT be written to the data table", vbOKOnly, "TDH variation too large"
3513                       btnRunNPSH_Click
3514                   Else    'tdh variation small
3515                       If Val(txtNPSH(2).Text) <= 97 Then
                               'btnRunNPSH_Click
                               'write the npsh and save
3516                           If WroteNPSHr = False Then
3517                               txtNPSH(4).Text = txtNPSHa.Text
3518                               rsTestData!NPSHr = txtNPSHa.Text
3519                               rsTestData.Update
3520                               rsEff!NPSHr = txtNPSHa.Text
3521                               rsEff.Update
3522                               WroteNPSHr = True
3523                               tmrNPSHr.Interval = 5000
3524                               tmrNPSHr.Enabled = True
3525                           End If
3526                       End If  'val < 97
3527                   End If  'check for tdh variation
3528               End If 'val tdh <=0
3529           Else    'frm not visible
                   'txtNPSHa = Format$(0, "##0.00")
3530           End If  'if frm visible

3531       Else
3532           txtNPSHa = 0
3533       End If
' <VB WATCH>
3534       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3535       Exit Sub
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
3536       On Error GoTo vbwErrHandler
3537       Const VBWPROCNAME = "frmPLCData.tmrStartUp_Timer"
3538       If vbwProtector.vbwTraceProc Then
3539           Dim vbwProtectorParameterString As String
3540           If vbwProtector.vbwTraceParameters Then
3541               vbwProtectorParameterString = "()"
3542           End If
3543           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3544       End If
' </VB WATCH>
3545       tmrStartUp.Enabled = False
' <VB WATCH>
3546       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3547       Exit Sub
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
3548       On Error GoTo vbwErrHandler
3549       Const VBWPROCNAME = "frmPLCData.SetCombo"
3550       If vbwProtector.vbwTraceProc Then
3551           Dim vbwProtectorParameterString As String
3552           If vbwProtector.vbwTraceParameters Then
3553               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
3554               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sName", sName) & ", "
3555               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("rs", rs) & ") "
3556           End If
3557           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3558       End If
' </VB WATCH>

3559       Dim I As Integer
3560       Dim sParam As String
3561       Dim qy As New ADODB.Command
3562       Dim rs1 As New ADODB.Recordset

3563       If rs.Fields(sName).ActualSize <> 0 Then     'if there's an entry
3564           sParam = rs.Fields(sName)                'get the index number
3565           qy.ActiveConnection = cnPumpData
3566           qy.CommandText = "SELECT * FROM " & sName & " WHERE " & sName & " = " & sParam
3567           Set rs1 = qy.Execute()                                  'get the record for the index number

3568           If rs1.BOF = True And rs1.EOF = True Then
3569               cmbComboName.ListIndex = -1                             'else, remove any pointer
' <VB WATCH>
3570       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3571               Exit Function
3572           End If

3573           For I = 0 To cmbComboName.ListCount - 1                     'go through the combobox entries
3574               If cmbComboName.ItemData(I) = rs1.Fields(0) Then     'see when we find the desired index number
3575                   cmbComboName.ListIndex = I                                              'if we do, set the combo box
3576                   Exit For                                            'and we're done
3577               End If
3578               cmbComboName.ListIndex = -1                             'else, remove any pointer
3579           Next I
3580       Else
3581           cmbComboName.ListIndex = -1
3582       End If

' <VB WATCH>
3583       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3584       Exit Function
' <VB WATCH>
3585       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3586       Exit Function
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
3587       On Error GoTo vbwErrHandler
3588       Const VBWPROCNAME = "frmPLCData.SetComboTestSetup"
3589       If vbwProtector.vbwTraceProc Then
3590           Dim vbwProtectorParameterString As String
3591           If vbwProtector.vbwTraceParameters Then
3592               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
3593               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sFieldName", sFieldName) & ", "
3594               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sTableName", sTableName) & ", "
3595               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("rs", rs) & ") "
3596           End If
3597           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3598       End If
' </VB WATCH>

       'same as setcombo, except here we also pass in the field name

3599       Dim I As Integer
3600       Dim sParam As String
3601       Dim qy As New ADODB.Command
3602       Dim rs1 As New ADODB.Recordset

3603       If rs.Fields(sFieldName).ActualSize <> 0 Then
               'if plc number, adjust plcaddress id numbers 1 and 2 to plc 8 and 9 respectively
3604           If sTableName = "CirculationFlowMeter" Then
                   'sParam = rs.Fields(sFieldName) + 7
3605               sParam = rs.Fields(sFieldName)
3606               If Val(sParam) < 4 Then
3607                   sParam = str(Val(sParam) + 4)
3608                   rs.Fields(sFieldName) = sParam
3609               End If
3610           Else
3611               sParam = rs.Fields(sFieldName)
3612           End If
3613           qy.ActiveConnection = cnPumpData
3614           qy.CommandText = "SELECT * FROM " & sTableName & " WHERE " & sTableName & " = " & sParam
3615           Set rs1 = qy.Execute()

3616           For I = 0 To cmbComboName.ListCount - 1
3617               If cmbComboName.ItemData(I) = rs1.Fields(0) Then
3618                   cmbComboName.ListIndex = I
3619                   Exit For
3620               End If
3621               cmbComboName.ListIndex = -1
3622           Next I
3623       Else
3624           cmbComboName.ListIndex = -1
3625       End If

' <VB WATCH>
3626       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3627       Exit Function
' <VB WATCH>
3628       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3629       Exit Function
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
3630       On Error GoTo vbwErrHandler
3631       Const VBWPROCNAME = "frmPLCData.DisablePumpDataControls"
3632       If vbwProtector.vbwTraceProc Then
3633           Dim vbwProtectorParameterString As String
3634           If vbwProtector.vbwTraceParameters Then
3635               vbwProtectorParameterString = "()"
3636           End If
3637           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3638       End If
' </VB WATCH>

3639       txtSalesOrderNumber.Enabled = False
3640       frmMfr.Enabled = False
3641       txtShpNo.Enabled = False
3642       txtBilNo.Enabled = False
3643       txtDesignFlow.Enabled = False
3644       txtDesignTDH.Enabled = False

3645       frmMiscPumpData.Enabled = False

3646       txtModelNo.Enabled = False
3647       txtImpellerDia.Enabled = False

3648       frmTEMC.Enabled = False
3649       frmChempump.Enabled = False

3650       txtRemarks.Enabled = False
3651       Me.cmdAddNewTestDate.Visible = False

3652       cmdEnterPumpData.Enabled = False

' <VB WATCH>
3653       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3654       Exit Sub
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
3655       On Error GoTo vbwErrHandler
3656       Const VBWPROCNAME = "frmPLCData.DisableTestSetupDataControls"
3657       If vbwProtector.vbwTraceProc Then
3658           Dim vbwProtectorParameterString As String
3659           If vbwProtector.vbwTraceParameters Then
3660               vbwProtectorParameterString = "()"
3661           End If
3662           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3663       End If
' </VB WATCH>

3664       cmbTestSpec.Enabled = False
3665       txtWho.Enabled = False
3666       txtRMA.Enabled = False

3667       frmLoopAndXducer.Enabled = False
3668       frmElecData.Enabled = False
3669       frmPerfMods.Enabled = False
3670       frmOtherFiles.Enabled = False
3671       frmInstrumentTags.Enabled = False
3672       frmTAndI.Enabled = False
3673       frmThrustBalMods.Enabled = False
3674       txtTestSetupRemarks.Enabled = False

3675       cmdEnterTestSetupData.Enabled = False
3676       cmbPLCNo.Enabled = False
' <VB WATCH>
3677       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3678       Exit Sub
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
3679       On Error GoTo vbwErrHandler
3680       Const VBWPROCNAME = "frmPLCData.DisableTestDataControls"
3681       If vbwProtector.vbwTraceProc Then
3682           Dim vbwProtectorParameterString As String
3683           If vbwProtector.vbwTraceParameters Then
3684               vbwProtectorParameterString = "()"
3685           End If
3686           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3687       End If
' </VB WATCH>

3688       cmbPLCLoop.Enabled = False
3689       frmPumpData.Enabled = False
3690       frmThermocouples.Enabled = False
3691       frmAI.Enabled = False
3692       frmMagtrol.Enabled = False
3693       fmrMiscTestData.Enabled = False
3694       frmPLCMisc.Enabled = False
3695       DataGrid1.Enabled = False
3696       DataGrid2.Enabled = False
3697       cmdEnterTestData.Enabled = False

' <VB WATCH>
3698       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3699       Exit Sub
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
3700       On Error GoTo vbwErrHandler
3701       Const VBWPROCNAME = "frmPLCData.EnableTestSetupDataControls"
3702       If vbwProtector.vbwTraceProc Then
3703           Dim vbwProtectorParameterString As String
3704           If vbwProtector.vbwTraceParameters Then
3705               vbwProtectorParameterString = "()"
3706           End If
3707           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3708       End If
' </VB WATCH>

3709       cmbTestSpec.Enabled = True
3710       txtWho.Enabled = True
3711       txtRMA.Enabled = True

3712       frmLoopAndXducer.Enabled = True
3713       frmElecData.Enabled = True
3714       frmPerfMods.Enabled = True
3715       frmOtherFiles.Enabled = True
3716       frmInstrumentTags.Enabled = True
3717       frmTAndI.Enabled = True
3718       frmThrustBalMods.Enabled = True
3719       txtTestSetupRemarks.Enabled = True

3720       cmdEnterTestSetupData.Enabled = True
3721       cmbPLCNo.Enabled = True
' <VB WATCH>
3722       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3723       Exit Sub
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
3724       On Error GoTo vbwErrHandler
3725       Const VBWPROCNAME = "frmPLCData.EnableTestDataControls"
3726       If vbwProtector.vbwTraceProc Then
3727           Dim vbwProtectorParameterString As String
3728           If vbwProtector.vbwTraceParameters Then
3729               vbwProtectorParameterString = "()"
3730           End If
3731           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3732       End If
' </VB WATCH>

3733       cmbPLCLoop.Enabled = True
3734       frmPumpData.Enabled = True
3735       frmThermocouples.Enabled = True
3736       frmAI.Enabled = True
3737       frmMagtrol.Enabled = True
3738       fmrMiscTestData.Enabled = True
3739       frmPLCMisc.Enabled = True
3740       DataGrid1.Enabled = True
3741       DataGrid2.Enabled = True
3742       cmdEnterTestData.Enabled = True

' <VB WATCH>
3743       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3744       Exit Sub
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
3745       On Error GoTo vbwErrHandler
3746       Const VBWPROCNAME = "frmPLCData.EnablePumpDataControls"
3747       If vbwProtector.vbwTraceProc Then
3748           Dim vbwProtectorParameterString As String
3749           If vbwProtector.vbwTraceParameters Then
3750               vbwProtectorParameterString = "()"
3751           End If
3752           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3753       End If
' </VB WATCH>

3754       txtSalesOrderNumber.Enabled = True
3755       frmMfr.Enabled = True
3756       txtShpNo.Enabled = True
3757       txtBilNo.Enabled = True
3758       txtDesignFlow.Enabled = True
3759       txtDesignTDH.Enabled = True

3760       frmMiscPumpData.Enabled = True

3761       txtModelNo.Enabled = True
3762       txtImpellerDia.Enabled = True

3763       frmTEMC.Enabled = True
3764       frmChempump.Enabled = True

3765       txtRemarks.Enabled = True
3766       Me.cmdAddNewTestDate.Visible = True

3767       cmdEnterPumpData.Enabled = True

' <VB WATCH>
3768       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3769       Exit Sub
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
3770       On Error GoTo vbwErrHandler
3771       Const VBWPROCNAME = "frmPLCData.EnableMagtrolFields"
3772       If vbwProtector.vbwTraceProc Then
3773           Dim vbwProtectorParameterString As String
3774           If vbwProtector.vbwTraceParameters Then
3775               vbwProtectorParameterString = "()"
3776           End If
3777           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3778       End If
' </VB WATCH>
3779       txtV1.Enabled = True
3780       txtV2.Enabled = True
3781       txtV3.Enabled = True
3782       txtI1.Enabled = True
3783       txtI2.Enabled = True
3784       txtI3.Enabled = True
3785       txtP1.Enabled = True
3786       txtP2.Enabled = True
3787       txtP3.Enabled = True
3788       optKW(0).Visible = True
3789       optKW(1).Visible = True
3790       optKW(2).Visible = True
3791       optKW(1).value = True
3792       optKW_Click (1)
' <VB WATCH>
3793       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3794       Exit Sub
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
3795       On Error GoTo vbwErrHandler
3796       Const VBWPROCNAME = "frmPLCData.DisableMagtrolFields"
3797       If vbwProtector.vbwTraceProc Then
3798           Dim vbwProtectorParameterString As String
3799           If vbwProtector.vbwTraceParameters Then
3800               vbwProtectorParameterString = "()"
3801           End If
3802           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3803       End If
' </VB WATCH>
3804       txtV1.Enabled = False
3805       txtV2.Enabled = False
3806       txtV3.Enabled = False
3807       txtI1.Enabled = False
3808       txtI2.Enabled = False
3809       txtI3.Enabled = False
3810       txtP1.Enabled = False
3811       txtP2.Enabled = False
3812       txtP3.Enabled = False
3813       txtKW.Enabled = False
3814       optKW(0).Visible = False
3815       optKW(1).Visible = False
3816       optKW(2).Visible = False
' <VB WATCH>
3817       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3818       Exit Sub
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
3819       On Error GoTo vbwErrHandler
3820       Const VBWPROCNAME = "frmPLCData.EnablePLCFields"
3821       If vbwProtector.vbwTraceProc Then
3822           Dim vbwProtectorParameterString As String
3823           If vbwProtector.vbwTraceParameters Then
3824               vbwProtectorParameterString = "()"
3825           End If
3826           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3827       End If
' </VB WATCH>
3828       frmPLCData.txtAI1Display.Enabled = True
3829       frmPLCData.txtAI2Display.Enabled = True
3830       frmPLCData.txtAI3Display.Enabled = True
3831       frmPLCData.txtAI4Display.Enabled = True
3832       frmPLCData.txtTC1Display.Enabled = True
3833       frmPLCData.txtTC2Display.Enabled = True
3834       frmPLCData.txtTC3Display.Enabled = True
3835       frmPLCData.txtTC4Display.Enabled = True
3836       frmPLCData.txtFlowDisplay.Enabled = True
3837       frmPLCData.txtSuctionDisplay.Enabled = True
3838       frmPLCData.txtDischargeDisplay.Enabled = True
3839       frmPLCData.txtTemperatureDisplay.Enabled = True
3840       frmPLCData.txtInHgDisplay.Enabled = True
' <VB WATCH>
3841       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3842       Exit Sub
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
3843       On Error GoTo vbwErrHandler
3844       Const VBWPROCNAME = "frmPLCData.DisablePLCFields"
3845       If vbwProtector.vbwTraceProc Then
3846           Dim vbwProtectorParameterString As String
3847           If vbwProtector.vbwTraceParameters Then
3848               vbwProtectorParameterString = "()"
3849           End If
3850           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3851       End If
' </VB WATCH>
3852       frmPLCData.txtAI1Display.Enabled = False
3853       frmPLCData.txtAI2Display.Enabled = False
3854       frmPLCData.txtAI3Display.Enabled = False
3855       frmPLCData.txtAI4Display.Enabled = False
3856       frmPLCData.txtTC1Display.Enabled = False
3857       frmPLCData.txtTC2Display.Enabled = False
3858       frmPLCData.txtTC3Display.Enabled = False
3859       frmPLCData.txtTC4Display.Enabled = False
3860       frmPLCData.txtFlowDisplay.Enabled = False
3861       frmPLCData.txtSuctionDisplay.Enabled = False
3862       frmPLCData.txtDischargeDisplay.Enabled = False
3863       frmPLCData.txtTemperatureDisplay.Enabled = False
3864       frmPLCData.txtInHgDisplay.Enabled = False
' <VB WATCH>
3865       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3866       Exit Sub
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
3867       On Error GoTo vbwErrHandler
3868       Const VBWPROCNAME = "frmPLCData.BlankData"
3869       If vbwProtector.vbwTraceProc Then
3870           Dim vbwProtectorParameterString As String
3871           If vbwProtector.vbwTraceParameters Then
3872               vbwProtectorParameterString = "()"
3873           End If
3874           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3875       End If
' </VB WATCH>
3876       txtShpNo.Text = vbNullString
3877       txtBilNo.Text = vbNullString
3878       txtModelNo.Text = vbNullString
3879       cmbMotor.ListIndex = -1
3880       cmbStatorFill.ListIndex = -1
3881       cmbVoltage.ListIndex = -1
3882       cmbDesignPressure.ListIndex = -1
3883       cmbFrequency.ListIndex = -1
3884       cmbCirculationPath.ListIndex = -1
3885       cmbRPM.ListIndex = -1
3886       cmbModel.ListIndex = -1
3887       cmbModelGroup.ListIndex = -1
3888       txtSpGr.Text = vbNullString
3889       txtImpellerDia.Text = vbNullString
3890       txtEndPlay.Text = vbNullString
3891       txtGGap.Text = vbNullString
3892       txtDesignFlow.Text = vbNullString
3893       txtDesignTDH.Text = vbNullString
3894       txtOtherMods.Text = vbNullString
3895       txtRemarks.Text = vbNullString
3896       txtSalesOrderNumber.Text = vbNullString
3897       txtTestSetupRemarks.Text = vbNullString
3898       txtNPSHFile.Text = vbNullString
3899       txtPicturesFile.Text = vbNullString
3900       txtVibrationFile.Text = vbNullString
       '    cmbOrificeNumber.ListIndex = 18

3901       SetFrequencyCombo

       '    cmbTestSpec.ListIndex = 6       'default = Rev7
3902       cmbLoopNumber.ListIndex = -1
3903       cmbSuctDia.ListIndex = -1
3904       cmbDischDia.ListIndex = -1
3905       cmbTachID.ListIndex = -1
3906       cmbAnalyzerNo.ListIndex = -1
3907       txtTestRemarks.Text = vbNullString
3908       txtHDCor.Text = 0
3909       txtDischHeight.Text = 0
3910       txtSuctHeight.Text = 0
3911       txtKWMult.Text = 1
3912       txtWho.Text = LogInInitials
3913       txtRMA.Text = vbNullString
3914       frmPLCData.chkNPSH.value = 0
3915       frmPLCData.chkPictures.value = 0
3916       frmPLCData.chkVibration.value = 0
3917       cmbFlowMeter.ListIndex = -1
3918       cmbSuctionPressureTransducer.ListIndex = -1
3919       cmbDischargePressureTransducer.ListIndex = -1
3920       cmbTemperatureTransducer.ListIndex = -1
3921       cmbCirculationFlowMeter.ListIndex = -1
3922       frmPLCData.chkBalanceHoles.value = 0
3923       frmPLCData.chkCircOrifice.value = 0
3924       frmPLCData.txtCircOrifice = vbNullString
3925       frmPLCData.txtImpTrim = vbNullString
3926       frmPLCData.txtOrifice = vbNullString
3927       frmPLCData.chkFeathered.value = Unchecked
3928       frmPLCData.chkTrimmed.value = 0
3929       frmPLCData.chkCircOrifice.value = 0
3930       frmPLCData.txtThrustBal = vbNullString
3931       frmPLCData.txtRPM = vbNullString
3932       frmPLCData.txtVibAx = vbNullString
3933       frmPLCData.txtVibRad = vbNullString
3934       frmPLCData.txtTEMCTRGReading = vbNullString
3935       dgBalanceHoles.Visible = False
3936       Me.txtLineNumber.Text = vbNullString
3937       Me.txtNPSHr.Text = vbNullString
3938       Me.txtRatedInputPower.Text = vbNullString
3939       Me.txtAmps.Text = vbNullString
3940       Me.txtThermalClass.Text = vbNullString
3941       Me.txtViscosity.Text = vbNullString
3942       Me.txtExpClass.Text = vbNullString
3943       Me.txtNoPhases.Text = vbNullString
3944       Me.txtLiquidTemperature.Text = vbNullString
3945       Me.txtJobNum.Text = vbNullString
3946       Me.txtTEMCFrameNumber.Text = vbNullString
3947       Me.txtLiquid.Text = vbNullString
3948       Me.chkSuperMarketFeathered.value = Unchecked
3949       Me.txtRVSPartNo.Text = vbNullString
' <VB WATCH>
3950       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3951       Exit Sub
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
3952       On Error GoTo vbwErrHandler
3953       Const VBWPROCNAME = "frmPLCData.AddTestData"
3954       If vbwProtector.vbwTraceProc Then
3955           Dim vbwProtectorParameterString As String
3956           If vbwProtector.vbwTraceParameters Then
3957               vbwProtectorParameterString = "()"
3958           End If
3959           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3960       End If
' </VB WATCH>
3961       Dim I As Integer
3962       Dim sFilter As String

3963       ClearEff
3964       rsEff.MoveFirst

3965       For I = 1 To 8
3966           rsTestData.AddNew
3967           rsTestData!SerialNumber = txtSN
3968           rsTestData!Date = cmbTestDate.List(cmbTestDate.ListIndex)
3969           rsTestData!testnumber = I
3970           rsTestData!DataWritten = False
3971           rsTestData.Update
3972           DoEfficiencyCalcs
3973           rsEff.MoveNext
3974           rsTestData.MoveNext
3975       Next I
3976       boFoundTestData = True
           'rsTestData.Update
3977       rsTestData.Requery
3978       rsTestData.Resync

          'select the entries from testdata
3979       sFilter = "SerialNumber='" & txtSN.Text & "' AND Date=#" & cmbTestDate.Text & "#"

3980       rsTestData.Filter = sFilter

3981       Set DataGrid1.DataSource = rsTestData

           ' fix the datagrid

3982       Dim c As Column
3983       For Each c In DataGrid1.Columns
3984          Select Case c.DataField
              Case "TestDataID"
3985             c.Visible = False
3986          Case "SerialNumber"
3987             c.Visible = False
3988          Case "Date"
3989             c.Visible = False
3990          Case Else ' Hide all other columns.
3991             c.Visible = True
3992             c.Alignment = dbgRight
3993          End Select
3994       Next c

3995       rsEff.Requery
3996       DataGrid1.Refresh
3997       DataGrid2.Refresh

' <VB WATCH>
3998       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3999       Exit Sub
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
4000       On Error GoTo vbwErrHandler
4001       Const VBWPROCNAME = "frmPLCData.DoEfficiencyCalcs"
4002       If vbwProtector.vbwTraceProc Then
4003           Dim vbwProtectorParameterString As String
4004           If vbwProtector.vbwTraceParameters Then
4005               vbwProtectorParameterString = "()"
4006           End If
4007           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4008       End If
' </VB WATCH>
4009       Dim KW As Single, VI As Single, VITemp As Single
4010       Dim Vave As Single, Iave As Single
4011       Dim I As Integer
4012       Dim j As Integer
4013       Dim HeightDiff As Single

4014       If Not IsNull(rsTestData.Fields("TotalPower")) Then
4015           KW = rsTestData.Fields("TotalPower")
4016       Else
               'if we wrote data with an old version, we will not have written total power
               'if total power = 0 and the three individual powers are not 0, add them

4017           If rsTestData.Fields("PowerA") > 0 Then
4018               If rsTestData.Fields("PowerB") > 0 Then
4019                   If rsTestData.Fields("PowerC") > 0 Then
4020                       KW = rsTestData.Fields("PowerA") + rsTestData.Fields("PowerB") + rsTestData.Fields("PowerC")
4021                   End If
4022               End If
4023           End If
4024      End If

4025       I = 0
4026       Vave = 0
4027       Iave = 0
4028       If Not IsNull(rsTestData.Fields("VoltageA")) And Not IsNull(rsTestData.Fields("CurrentA")) Then
4029           VI = rsTestData.Fields("VoltageA") * rsTestData.Fields("CurrentA")
4030           Vave = rsTestData.Fields("VoltageA")
4031           Iave = rsTestData.Fields("CurrentA")
4032           If VI <> 0 Then
4033               I = I + 1
4034           End If
4035       End If
4036       If Not IsNull(rsTestData.Fields("VoltageB")) And Not IsNull(rsTestData.Fields("CurrentB")) Then
4037           VITemp = rsTestData.Fields("VoltageB") * rsTestData.Fields("CurrentB")
4038           If VITemp <> 0 Then
4039               I = I + 1
4040               VI = VI + VITemp
4041               Vave = Vave + rsTestData.Fields("VoltageB")
4042               Iave = Iave + rsTestData.Fields("CurrentB")
4043           End If
4044       End If
4045       If Not IsNull(rsTestData.Fields("VoltageC")) And Not IsNull(rsTestData.Fields("CurrentC")) Then
4046           VITemp = rsTestData.Fields("VoltageC") * rsTestData.Fields("CurrentC")
4047           If VITemp <> 0 Then
4048               I = I + 1
4049               VI = VI + VITemp
4050               Vave = Vave + rsTestData.Fields("VoltageC")
4051               Iave = Iave + rsTestData.Fields("CurrentC")
4052           End If
4053       End If
4054       If KW = 0 Then
4055           For j = 1 To rsEff.Fields.Count - 1
4056               rsEff.Fields(j) = 0
4057           Next j
       '        Exit Sub
4058       End If
4059       If VI <> 0 Then
4060           rsEff.Fields("Volts") = Vave / I
4061           rsEff.Fields("Amps") = Iave / I
4062           rsEff.Fields("PowerFactor") = 1000 * I * KW / (VI * Sqr(3))
4063           rsEff.Fields("PowerFactor") = 100 * rsEff.Fields("PowerFactor")
4064       Else
4065           rsEff.Fields("PowerFactor") = 0
4066       End If

4067       If optMfr(0).value = True Then
4068           If cmbStatorFill.ListIndex = -1 Then
4069               rsEff.Fields("MotorEfficiency") = Format$(0, "0.00")

4070           Else
4071               rsEff.Fields("Motorefficiency") = Format$(Round(MotorEfficiency(KW, cmbMotor.ItemData(cmbMotor.ListIndex), cmbStatorFill.ItemData(cmbStatorFill.ListIndex)), 1), "00.0")
       '            rsEff.Fields("Motorefficiency") = Format$(Round(MotorEfficiency(KW, cmbMotor.ListIndex, cmbStatorFill.ListIndex), 1), "00.0")
4072           End If
4073       Else
4074           rsEff.Fields("MotorEfficiency") = Format$(Round(TEMCMotorEfficiency(KW, txtTEMCFrameNumber.Text, 460, RatedKW), 1), "00.0")
4075       End If

4076       Dim sHDCor As Single
4077       Dim sDisc As Single
4078       Dim sSuct As Single
4079       If IsNull(rsTestSetup.Fields("HDCor")) Then
4080           sHDCor = 0
4081       Else
4082           sHDCor = rsTestSetup.Fields("HDCor")
4083       End If
4084       If IsNull(rsTestSetup.Fields("DischargeGageHeight")) Then
4085           sDisc = 0
4086       Else
4087           sDisc = rsTestSetup.Fields("DischargeGageHeight")
4088       End If
4089       If IsNull(rsTestSetup.Fields("SuctionGageHeight")) Then
4090           sSuct = 0
4091       Else
4092           sSuct = rsTestSetup.Fields("SuctionGageHeight")
4093       End If
4094       HeightDiff = sHDCor + sDisc / 12 - sSuct / 12
4095       If (cmbDischDia.ListIndex <> -1 And cmbSuctDia.ListIndex <> -1) Then
4096           rsEff.Fields("VelocityHead") = CalcVelHead(rsTestData.Fields("Flow"), cmbDischDia.ItemData(cmbDischDia.ListIndex) + 1, cmbSuctDia.ItemData(cmbSuctDia.ListIndex) + 1)
4097       End If
       '    rsEff.Fields("VelocityHead") = CalcVelHead(rsTestData.Fields("Flow"), cmbDischDia.ListIndex + 1, cmbSuctDia.ListIndex + 1)
4098       rsEff.Fields("TDH") = CalcTDH(rsTestData.Fields("DischargePressure"), rsTestData.Fields("SuctionPressure"), rsTestData.Fields("SuctionInHg"), rsEff.Fields("VelocityHead"), HeightDiff, rsTestData.Fields("TemperatureSuction"))
4099       rsEff.Fields("ElecHP") = 1000 * KW / 746
       '    If (DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))) <> 0 And KW <> 0) Then
4100           If Int(rsTestData.Fields("TemperatureSuction")) >= 40 Then
4101               If (DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))) <> 0 And KW <> 0) Then
           '        rsEff.Fields("LiquidHP") = (rsEff.Fields("TDH") * rsTestData.Fields("Flow") * DLookup("TDHCorr", "TempCorrection", "Temp = 68")) / (3960 * DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))))
4102               rsEff.Fields("LiquidHP") = (rsEff.Fields("TDH") * rsTestData.Fields("Flow") * DLookupA(TDHColNo, TempCorrection, TempColNo, 68)) / (3960 * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))))
           '        rsEff.Fields("OverallEfficiency") = (0.189 * rsTestData.Fields("Flow") * rsEff.Fields("TDH") * DLookup("TDHCorr", "TempCorrection", "Temp = 68")) / (10 * KW * DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))))
4103               rsEff.Fields("OverallEfficiency") = (0.189 * rsTestData.Fields("Flow") * rsEff.Fields("TDH") * DLookupA(TDHColNo, TempCorrection, TempColNo, 68)) / (10 * KW * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))))
4104               If rsEff.Fields("MotorEfficiency") <> 0 Then
4105                   rsEff.Fields("HydraulicEfficiency") = 100 * rsEff.Fields("OverallEfficiency") / rsEff.Fields("MotorEfficiency")
4106               Else
4107                   rsEff.Fields("HydraulicEfficiency") = 0
4108               End If
4109           Else
4110               rsEff.Fields("LiquidHP") = 0
4111               rsEff.Fields("OverallEfficiency") = 0
4112           End If

4113       Else
4114           rsEff.Fields("LiquidHP") = 0
4115           rsEff.Fields("OverallEfficiency") = 0
4116       End If


4117       I = rsEff.AbsolutePosition
4118       If Not IsNull(rsTestData.Fields("Flow")) Then
4119           rsEff.Fields("Flow") = rsTestData.Fields("Flow")
4120           HeadFlow(0, I - 1) = rsTestData.Fields("Flow")
4121           HeadFlow(1, I - 1) = rsEff.Fields("TDH")
4122           FlowHead(I - 1, 0) = rsTestData.Fields("Flow")
4123           FlowHead(I - 1, 1) = rsEff.Fields("TDH")

       '        EffFlow(0, i - 1) = rsTestData.Fields("Flow")
       '        EffFlow(1, i - 1) = rsEff.Fields("OverallEfficiency")
       '        KWFlow(0, i - 1) = rsTestData.Fields("Flow")
       '        KWFlow(1, i - 1) = KW
       '        AmpsFlow(0, i - 1) = rsTestData.Fields("Flow")
       '        AmpsFlow(1, i - 1) = rsEff.Fields("Amps")
4124       Else
4125           HeadFlow(0, I - 1) = 0
4126           HeadFlow(1, I - 1) = 0
4127           FlowHead(I - 1, 0) = 0
4128           FlowHead(I - 1, 1) = 0

       '        EffFlow(0, i - 1) = 0
       '        EffFlow(1, i - 1) = 0
       '        KWFlow(0, i - 1) = 0
       '        KWFlow(1, i - 1) = 0
       '        AmpsFlow(0, i - 1) = 0
       '        AmpsFlow(1, i - 1) = 0
4129       End If

4130       Dim Plothead(1, 7) As Single
4131       Dim HeadPlot(7, 1) As Single
           'ReDim Preserve Plothead(1, j)
           'ReDim Preserve HeadPlot(j, 1)

       '    Dim PlotEff() As Single
       '    Dim PlotKW() As Single
       '    Dim PlotAmps() As Single
       '    ReDim PlotHead(0, 0)
       '    ReDim PlotEff(0, 0)
       '    ReDim PlotKW(0, 0)
       '
4132       For j = 0 To UpDown2.value - 1
       '        If HeadFlow(1, j) <> 0 Then
       '            ReDim Preserve Plothead(1, j)
       '            ReDim Preserve HeadPlot(j, 1)
4133               Plothead(0, j) = HeadFlow(0, j)
4134               Plothead(1, j) = HeadFlow(1, j)
4135               HeadPlot(j, 0) = FlowHead(j, 0)
4136               HeadPlot(j, 1) = FlowHead(j, 1)
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
4137       Next j




       '    SetGraphMax (Plothead())
       '    If UBound(PlotHead()) <> 0 Then

       'fix 4/29/19

4138           MSChart1.ChartData = HeadPlot

       '    End If

           'copy fields for reports
4139       rsEff.Fields("DischPress") = rsTestData.Fields("Dischargepressure")
4140       rsEff.Fields("SuctPress") = rsTestData.Fields("Suctionpressure")
       '    rsEff.Fields("Volts") = rsTestData.Fields("VoltageA")
       '    rsEff.Fields("Amps") = rsTestData.Fields("CurrentA")
4141       rsEff.Fields("KW") = KW
4142       rsEff.Fields("Freq") = rsTestData.Fields("VFDFrequency")
4143       rsEff.Fields("RPM") = rsTestData.Fields("RPM")
4144       rsEff.Fields("Pos") = rsTestData.Fields("ThrustBalance")
4145       rsEff.Fields("NPSHa") = rsTestData.Fields("NPSHa")
4146       rsEff.Fields("NPSHr") = rsTestData.Fields("NPSHr")
4147       rsEff.Fields("InputPower") = rsTestData.Fields("TotalPower")
4148       rsEff.Fields("Temperature") = rsTestData.Fields("TemperatureSuction")
4149       rsEff.Fields("CircFlow") = rsTestData.Fields("CircFlow")
4150       rsEff.Fields("VibrationX") = rsTestData.Fields("VibrationX")
4151       rsEff.Fields("VibrationY") = rsTestData.Fields("VibrationY")
4152       rsEff.Fields("CurrentA") = rsTestData.Fields("CurrentA")
4153       rsEff.Fields("CurrentB") = rsTestData.Fields("CurrentB")
4154       rsEff.Fields("CurrentC") = rsTestData.Fields("CurrentC")
4155       rsEff.Fields("VoltageA") = rsTestData.Fields("VoltageA")
4156       rsEff.Fields("VoltageB") = rsTestData.Fields("VoltageB")
4157       rsEff.Fields("VoltageC") = rsTestData.Fields("VoltageC")
4158       rsEff.Fields("TC1") = rsTestData.Fields("TC1")
4159       rsEff.Fields("TC2") = rsTestData.Fields("TC2")
4160       rsEff.Fields("TC3") = rsTestData.Fields("TC3")
4161       rsEff.Fields("TC4") = rsTestData.Fields("TC4")
4162       rsEff.Fields("RBHTemp") = rsTestData.Fields("RBHTemp")
4163       rsEff.Fields("RBHPress") = rsTestData.Fields("RBHPress")
4164       rsEff.Fields("AI4") = rsTestData.Fields("AI4")
4165       rsEff.Fields("Remarks") = rsTestData.Fields("Remarks")
4166       rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCFrontThrust")
4167       rsEff.Fields("TEMCRearThrust") = rsTestData.Fields("TEMCRearThrust")
4168       rsEff.Fields("TEMCTRG") = rsTestData.Fields("TEMCTRG")
4169       rsEff.Fields("TEMCThrustRigPressure") = rsTestData.Fields("TEMCThrustRigPressure")
4170       rsEff.Fields("TEMCMomentArm") = rsTestData.Fields("TEMCMomentArm")
4171       rsEff.Fields("TEMCViscosity") = rsTestData.Fields("TEMCViscosity")
4172       If Not IsNull(rsEff.Fields("TEMCFrontThrust")) Then
4173           txtTEMCFrontThrust.Text = rsEff.Fields("TEMCFrontThrust")
4174       End If
4175       If Not IsNull(rsEff.Fields("TEMCREarThrust")) Then
4176           txtTEMCRearThrust.Text = rsEff.Fields("TEMCREarThrust")
4177       End If
4178       If (Not IsNull(rsEff.Fields("TEMCViscosity"))) And (rsEff.Fields("TEMCViscosity") <> 0) Then
4179           txtTEMCViscosity.Text = rsEff.Fields("TEMCViscosity")
4180       End If
4181       If Not IsNull(rsTestData.Fields("TEMCThrustRigPressure")) Then
4182           txtTEMCThrustRigPressure.Text = rsTestData.Fields("TEMCThrustRigPressure")
4183       End If
4184       If Not IsNull(rsTestData.Fields("TEMCMomentArm")) Then
4185           txtTEMCMomentArm.Text = rsTestData.Fields("TEMCMomentArm")
4186       End If

        '   If Not IsNull(Me.txtAI3Display.Text) Then
        '       Me.txtAI3Display = rsTestData.Fields("RBHPress")
        '   End If

4187       CalculateTEMCForce

4188       If Not IsNull(txtTEMCCalcForce.Text) Then
4189           rsEff.Fields("TEMCCalculatedForce") = Val(txtTEMCCalcForce.Text)
4190       Else
4191           rsEff.Fields("TEMCCalculatedForce") = 0
4192       End If

4193       If Not IsNull(txtTEMCPVValue.Text) Then
4194           rsEff.Fields("TEMCPV") = Val(txtTEMCPVValue.Text)
4195       Else
4196           rsEff.Fields("TEMCPV") = 0
4197       End If

4198       If Val(txtTEMCFrontThrust.Text) <> 0 Then
4199           rsEff.Fields("TEMCFR") = "F"
       '        rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCFrontThrust")
4200       Else
4201           If Val(txtTEMCRearThrust.Text) = 0 Then
                   'no thrust
4202               rsEff.Fields("TEMCFR") = " "
4203               rsEff.Fields("TEMCFrontThrust") = 0
4204           Else
4205               rsEff.Fields("TEMCFR") = "R"
       '            rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCRearThrust")
4206           End If
4207       End If

4208       rsEff.Fields("TEMCForceDirection") = Left(lblTEMCFrontRear.Caption, 1)

4209       rsEff.Update
4210       DataGrid2.Refresh


' <VB WATCH>
4211       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4212       Exit Sub
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
4213       On Error GoTo vbwErrHandler
4214       Const VBWPROCNAME = "frmPLCData.ClearEff"
4215       If vbwProtector.vbwTraceProc Then
4216           Dim vbwProtectorParameterString As String
4217           If vbwProtector.vbwTraceParameters Then
4218               vbwProtectorParameterString = "()"
4219           End If
4220           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4221       End If
' </VB WATCH>
4222       Dim qy As New ADODB.Command

4223       If rsEff.State = adStateOpen Then
4224           If Not (rsEff.BOF = True Or rsEff.EOF = True) Then
4225               rsEff.CancelUpdate
4226           End If
4227           rsEff.Close
4228       End If
4229       qy.ActiveConnection = cnEffData
4230       qy.CommandText = "DROP TABLE Efficiency"
4231       rsEff.Open qy
4232       qy.CommandText = "SELECT EfficiencyOrg.* INTO Efficiency FROM EfficiencyOrg;"
4233       rsEff.Open qy
4234       rsEff.Open "Efficiency", cnEffData, adOpenStatic, adLockOptimistic, adCmdTableDirect

4235       rsEff.Requery
4236       DataGrid2.Refresh

4237       Dim c As Column
4238       For Each c In DataGrid2.Columns
4239           c.Alignment = dbgCenter
4240           c.Width = 750
4241           Select Case c.ColIndex
                   Case 1
4242                   c.Caption = "Flow"
4243                   c.NumberFormat = "###0.00"
4244               Case 2
4245                   c.Caption = "TDH"
4246                   c.NumberFormat = "00.0"
4247               Case 3
4248                   c.Caption = "Overall Eff"
4249                   c.NumberFormat = "00.00"
4250                   c.Width = 850
4251               Case 4
4252                   c.Caption = "PF"
4253                   c.NumberFormat = "00.0"
4254               Case 5
4255                   c.Caption = "Vel Head"
4256                   c.NumberFormat = "00.00"
4257               Case 6
4258                   c.Caption = "Elec HP"
4259                   c.NumberFormat = "#00.0"
4260               Case 7
4261                   c.Caption = "Liq HP"
4262                   c.NumberFormat = "#00.0"
4263               Case Else
4264                   c.Visible = False
4265           End Select
4266       Next c

' <VB WATCH>
4267       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4268       Exit Sub
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
4269       On Error GoTo vbwErrHandler
4270       Const VBWPROCNAME = "frmPLCData.JustAlphaNumeric"
4271       If vbwProtector.vbwTraceProc Then
4272           Dim vbwProtectorParameterString As String
4273           If vbwProtector.vbwTraceParameters Then
4274               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("char", char) & ") "
4275           End If
4276           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4277       End If
' </VB WATCH>
4278       Select Case Asc(char)
               Case 42             ' *
4279               JustAlphaNumeric = char
4280           Case 48 To 57       ' 0 - 9
4281               JustAlphaNumeric = char
4282           Case 65 To 90       ' A - Z
4283               JustAlphaNumeric = char
4284           Case 97 To 122      ' a - z
4285               JustAlphaNumeric = UCase(char)
4286           Case Else
4287               JustAlphaNumeric = ""
4288       End Select
' <VB WATCH>
4289       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4290       Exit Function
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
4291       On Error GoTo vbwErrHandler
4292       Const VBWPROCNAME = "frmPLCData.txtI1_Change"
4293       If vbwProtector.vbwTraceProc Then
4294           Dim vbwProtectorParameterString As String
4295           If vbwProtector.vbwTraceParameters Then
4296               vbwProtectorParameterString = "()"
4297           End If
4298           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4299       End If
' </VB WATCH>
4300       txtI2.Text = txtI1.Text
4301       txtI3.Text = txtI1.Text
' <VB WATCH>
4302       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4303       Exit Sub
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
4304       On Error GoTo vbwErrHandler
4305       Const VBWPROCNAME = "frmPLCData.txtModelNo_Change"
4306       If vbwProtector.vbwTraceProc Then
4307           Dim vbwProtectorParameterString As String
4308           If vbwProtector.vbwTraceParameters Then
4309               vbwProtectorParameterString = "()"
4310           End If
4311           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4312       End If
' </VB WATCH>
4313       Dim I As Integer
4314       Dim S As String
4315       Dim sFull As String
4316       Dim boDone As Boolean
4317       Dim boRepeat As Boolean

4318       Static bo3Digits As Boolean         '3 digits in frame number
4319       Static bo2Digits As Boolean         '2 digits in stages

4320       If optMfr(0).value = True Then
' <VB WATCH>
4321       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4322           Exit Sub
4323       End If

4324       cmbTEMCAdapter.ListIndex = -1
4325       cmbTEMCAdditions.ListIndex = -1
4326       cmbTEMCCirculation.ListIndex = -1
4327       cmbTEMCDesignPressure.ListIndex = -1
4328       cmbTEMCNominalDischargeSize.ListIndex = -1
4329       cmbTEMCDivisionType.ListIndex = -1
4330       cmbTEMCImpellerType.ListIndex = -1
4331       cmbTEMCInsulation.ListIndex = -1
4332       cmbTEMCJacketGasket.ListIndex = -1
4333       cmbTEMCMaterials.ListIndex = -1
4334       cmbTEMCModel.ListIndex = -1
4335       cmbTEMCNominalImpSize.ListIndex = -1
4336       cmbTEMCOtherMotor.ListIndex = -1
4337       cmbTEMCPumpStages.ListIndex = -1
4338       cmbTEMCNominalSuctionSize.ListIndex = -1
4339       cmbTEMCTRG.ListIndex = -1
4340       cmbTEMCVoltage.ListIndex = -1


           'first, get rid of spaces, dashes, etc

4341       S = ""
4342       For I = 1 To Len(txtModelNo.Text)
4343           S = S & JustAlphaNumeric(Mid$(txtModelNo.Text, I, 1))
4344       Next I

           'next, fill out the model number to it's max length of 24 characters

4345       boDone = False
4346       boRepeat = False

4347       Do While Not boDone
4348           sFull = ""
4349           For I = 1 To Len(S)
4350               Select Case I
                       Case 1
                           'type
4351                       sFull = sFull & Mid$(S, I, 1)
4352                   Case 2
                           'adapter
4353                       If IsNumeric(Mid$(S, I, 1)) Then
4354                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4355                           boRepeat = True
4356                           Exit For
4357                       Else
4358                           sFull = sFull & Mid$(S, I, 1)
4359                           boRepeat = False
4360                       End If
4361                   Case 3
                           'materials
4362                       sFull = sFull & Mid$(S, I, 1)
4363                   Case 4
                       'design pressure
4364                       sFull = sFull & Mid$(S, I, 1)
4365                   Case 5
                       'motor frame number - digit 1
4366                       sFull = sFull & Mid$(S, I, 1)
4367                   Case 6
                       'motor frame number - digit 2
4368                       sFull = sFull & Mid$(S, I, 1)
4369                   Case 7
                       'motor frame number - digit 3
4370                       sFull = sFull & Mid$(S, I, 1)
4371                   Case 8
                       'motor frame number - digit 4
4372                       If IsNumeric(Mid$(S, I, 1)) Then
4373                           sFull = sFull & Mid$(S, I, 1)
4374                           boRepeat = False
4375                       Else    '3 digits
       '                        s = Left$(s, i - 1) & "*" & Right$(s, Len(s) - i + 1)
4376                           S = Left$(S, I - 4) & "0" & Right$(S, Len(S) - I + 4)
4377                           boRepeat = True
4378                           Exit For
4379                       End If
4380                   Case 9
                       'insulation
4381                       sFull = sFull & Mid$(S, I, 1)
4382                   Case 10
                       'voltage
4383                       sFull = sFull & Mid$(S, I, 1)
4384                   Case 11
                       'other motor specs
4385                       If Mid$(S, I, 1) = "M" Or Mid$(S, I, 1) = "R" Or Mid$(S, I, 1) = "L" Or Mid$(S, I, 1) = "G" Or Mid$(S, I, 1) = "N" Then
4386                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4387                           boRepeat = True
4388                           Exit For
4389                       Else
4390                           sFull = sFull & Mid$(S, I, 1)
4391                           boRepeat = False
4392                       End If
4393                   Case 12
                       ' TRG
4394                       sFull = sFull & Mid$(S, I, 1)
4395                   Case 13
                       'Nominal discharge - digit 1
4396                       sFull = sFull & Mid$(S, I, 1)
4397                   Case 14
                       'nominal discharge - digit 2
4398                       sFull = sFull & Mid$(S, I, 1)
4399                   Case 15
                       'nominal suction - digit 1
4400                       sFull = sFull & Mid$(S, I, 1)
4401                   Case 16
                       'nominal suction - digit 2
4402                       sFull = sFull & Mid$(S, I, 1)
4403                   Case 17
                       'nominal impeller size
4404                       sFull = sFull & Mid$(S, I, 1)
4405                   Case 18
                       'impeller type
4406                       If Mid$(S, I, 1) <> "*" Then
4407                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4408                           boRepeat = True
4409                           Exit For
4410                       Else
4411                           sFull = sFull & Mid$(S, I, 1)
4412                           boRepeat = False
4413                       End If
4414                   Case 19
                       'Division type
4415                       If IsNumeric(Mid$(S, I, 1)) Then
4416                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4417                           boRepeat = True
4418                           Exit For
4419                       Else
4420                           sFull = sFull & Mid$(S, I, 1)
4421                           boRepeat = False
4422                       End If
4423                   Case 20
                       'pump stages - digit 1
4424                       sFull = sFull & Mid$(S, I, 1)
4425                   Case 21
                       'pump jacket
4426                       If Mid$(S, I, 1) = "A" Or Mid$(S, I, 1) = "B" Or Mid$(S, I, 1) = "E" Or Mid$(S, I, 1) = "F" Or _
                                             Mid$(S, I, 1) = "G" Or Mid$(S, I, 1) = "H" Or Mid$(S, I, 1) = "J" Or Mid$(S, I, 1) = "K" Then
4427                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4428                           boRepeat = True
4429                       Else
4430                           sFull = sFull & Mid$(S, I, 1)
4431                           boRepeat = False
4432                       End If
4433                   Case 22
                       'additions
4434                         sFull = sFull & Mid$(S, I, 1)
4435                   Case 23
                       'circulation
4436                         sFull = sFull & Mid$(S, I, 1)
4437               End Select
4438           Next I
4439           If Not boRepeat Then
4440               boDone = True
4441           End If
4442       Loop

4443       For I = 1 To Len(sFull)
4444           Select Case I
                   Case 1
4445                   ParseTEMCModelNo cmbTEMCModel, Mid$(sFull, I, 1)
4446               Case 2
4447                   ParseTEMCModelNo cmbTEMCAdapter, Mid$(sFull, I, 1)
4448               Case 3
4449                   ParseTEMCModelNo cmbTEMCMaterials, Mid$(sFull, I, 1)
4450               Case 4
4451                   ParseTEMCModelNo cmbTEMCDesignPressure, Mid$(sFull, I, 1)
4452               Case 5
4453                       If Val(Mid$(sFull, I, 1)) = 0 Then
4454                           txtTEMCFrameNumber.Text = Mid$(sFull, 6, 3)
4455                       Else
4456                           txtTEMCFrameNumber.Text = Mid$(sFull, 5, 4)
4457                       End If
4458               Case 9
4459                       ParseTEMCModelNo cmbTEMCInsulation, Mid$(sFull, I, 1)
4460               Case 10
4461                       ParseTEMCModelNo cmbTEMCVoltage, Mid$(sFull, I, 1)
4462               Case 11
4463                       ParseTEMCModelNo cmbTEMCOtherMotor, Mid$(sFull, I, 1)
4464               Case 12
4465                       ParseTEMCModelNo cmbTEMCTRG, Mid$(sFull, I, 1)
4466               Case 13
4467                       ParseTEMCModelNo cmbTEMCNominalDischargeSize, Mid$(sFull, I, 2)
4468               Case 14
4469               Case 15
4470                       ParseTEMCModelNo cmbTEMCNominalSuctionSize, Mid$(sFull, I, 2)
4471               Case 16
4472               Case 17
4473                       ParseTEMCModelNo cmbTEMCNominalImpSize, Mid$(sFull, I, 1)
4474               Case 18
4475                       ParseTEMCModelNo cmbTEMCImpellerType, Mid$(sFull, I, 1)
4476               Case 19
4477                       ParseTEMCModelNo cmbTEMCDivisionType, Mid$(sFull, I, 1)
4478               Case 20
4479                       ParseTEMCModelNo cmbTEMCPumpStages, Mid$(sFull, I, 1)
4480               Case 21
4481                       ParseTEMCModelNo cmbTEMCJacketGasket, Mid$(sFull, I, 1)
4482               Case 22
4483                       ParseTEMCModelNo cmbTEMCAdditions, Mid$(sFull, I, 1)
4484                       ParseTEMCModelNo cmbTEMCCirculation, "*"
4485               Case 23
       '                    ParseTEMCModelNo cmbTEMCCirculation, Mid$(sFull, I, 1)

4486           End Select
4487       Next I

           'give alerts on certain conditions
4488       Dim msg As String
4489       msg = ""
4490       If Left(cmbTEMCVoltage, 3) = "[6]" Then
4491           msg = "575V transformer required for Rundown and TRG"
4492       End If
       '    If Left(cmbTEMCTRG, 3) = "[L]" Or InStr("X[B][F][H][K]", Left(cmbTEMCAdditions, 3)) > 1 Then
4493       If Left(cmbTEMCTRG, 3) = "[L]" Then
4494           If msg = "" Then
4495               msg = "VFD required for Rundown and TRG"
4496           Else
4497               msg = msg & " and " & "VFD required for Rundown and TRG"
4498           End If
4499       End If

4500       If InStr("X[B][F][H][K]", Left(cmbTEMCAdditions, 3)) > 1 Then
4501           If msg = "" Then
4502               msg = "VFD required for Rundown, standard drive required for TRG"
4503           Else
4504               msg = msg & " and " & "VFD required for Rundown, standard drive required for TRG"
4505           End If
4506       End If

4507       If msg <> "" Then
4508           frmAlert.txtAlert.Text = msg
4509           frmAlert.Show
4510       End If

' <VB WATCH>
4511       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4512       Exit Sub
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
4513       On Error GoTo vbwErrHandler
4514       Const VBWPROCNAME = "frmPLCData.txtModelNo_Validate"
4515       If vbwProtector.vbwTraceProc Then
4516           Dim vbwProtectorParameterString As String
4517           If vbwProtector.vbwTraceParameters Then
4518               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Cancel", Cancel) & ") "
4519           End If
4520           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4521       End If
' </VB WATCH>
4522       Dim I As Integer
4523       Dim S As String

       '    s = txtModelNo.Text
       '    S = Replace(S, "-", "")
       '    S = Replace(S, " ", "")
       '    S = Replace(S, "/", "")

       '    txtModelNo.Text = ""

       '    For i = 1 To Len(s)
       '        txtModelNo.Text = txtModelNo.Text & Mid(s, i, 1)
       '    Next i
4524       txtModelNo_Change

' <VB WATCH>
4525       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4526       Exit Sub
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
4527       On Error GoTo vbwErrHandler
4528       Const VBWPROCNAME = "frmPLCData.txtNPSHFile_GotFocus"
4529       If vbwProtector.vbwTraceProc Then
4530           Dim vbwProtectorParameterString As String
4531           If vbwProtector.vbwTraceParameters Then
4532               vbwProtectorParameterString = "()"
4533           End If
4534           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4535       End If
' </VB WATCH>
4536       On Error GoTo FileCancel
4537       If LenB(txtNPSHFile.Text) <> 0 Then
4538           CommonDialog1.filename = txtNPSHFile.Text
4539       End If
4540       CommonDialog1.ShowOpen
4541       txtNPSHFile.Text = CommonDialog1.filename
' <VB WATCH>
4542       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4543       Exit Sub
4544   FileCancel:
4545   On Error GoTo vbwErrHandler
4546       CommonDialog1.CancelError = False
' <VB WATCH>
4547       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4548       Exit Sub
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
4549       On Error GoTo vbwErrHandler
4550       Const VBWPROCNAME = "frmPLCData.txtP1_Change"
4551       If vbwProtector.vbwTraceProc Then
4552           Dim vbwProtectorParameterString As String
4553           If vbwProtector.vbwTraceParameters Then
4554               vbwProtectorParameterString = "()"
4555           End If
4556           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4557       End If
' </VB WATCH>
4558       txtP2.Text = txtP1.Text
4559       txtP3.Text = txtP1.Text
' <VB WATCH>
4560       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4561       Exit Sub
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
4562       On Error GoTo vbwErrHandler
4563       Const VBWPROCNAME = "frmPLCData.txtPicturesFile_gotfocus"
4564       If vbwProtector.vbwTraceProc Then
4565           Dim vbwProtectorParameterString As String
4566           If vbwProtector.vbwTraceParameters Then
4567               vbwProtectorParameterString = "()"
4568           End If
4569           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4570       End If
' </VB WATCH>
4571       CommonDialog1.CancelError = True
4572       On Error GoTo FileCancel
4573       If LenB(txtPicturesFile.Text) <> 0 Then
4574           CommonDialog1.filename = txtPicturesFile.Text
4575       End If
4576       CommonDialog1.ShowOpen
4577       txtPicturesFile.Text = CommonDialog1.filename
' <VB WATCH>
4578       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4579       Exit Sub
4580   FileCancel:
4581   On Error GoTo vbwErrHandler
4582       CommonDialog1.CancelError = False
' <VB WATCH>
4583       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4584       Exit Sub
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
4585       On Error GoTo vbwErrHandler
4586       Const VBWPROCNAME = "frmPLCData.txtSN_Change"
4587       If vbwProtector.vbwTraceProc Then
4588           Dim vbwProtectorParameterString As String
4589           If vbwProtector.vbwTraceParameters Then
4590               vbwProtectorParameterString = "()"
4591           End If
4592           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4593       End If
' </VB WATCH>
4594       cmdFindPump.Default = True
' <VB WATCH>
4595       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4596       Exit Sub
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
4597       On Error GoTo vbwErrHandler
4598       Const VBWPROCNAME = "frmPLCData.txtTEMCFrontThrust_Change"
4599       If vbwProtector.vbwTraceProc Then
4600           Dim vbwProtectorParameterString As String
4601           If vbwProtector.vbwTraceParameters Then
4602               vbwProtectorParameterString = "()"
4603           End If
4604           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4605       End If
' </VB WATCH>
4606       CalculateTEMCForce
' <VB WATCH>
4607       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4608       Exit Sub
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
4609       On Error GoTo vbwErrHandler
4610       Const VBWPROCNAME = "frmPLCData.txtTEMCMomentArm_Change"
4611       If vbwProtector.vbwTraceProc Then
4612           Dim vbwProtectorParameterString As String
4613           If vbwProtector.vbwTraceParameters Then
4614               vbwProtectorParameterString = "()"
4615           End If
4616           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4617       End If
' </VB WATCH>
4618       CalculateTEMCForce
' <VB WATCH>
4619       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4620       Exit Sub
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
4621       On Error GoTo vbwErrHandler
4622       Const VBWPROCNAME = "frmPLCData.txtTEMCRearThrust_Change"
4623       If vbwProtector.vbwTraceProc Then
4624           Dim vbwProtectorParameterString As String
4625           If vbwProtector.vbwTraceParameters Then
4626               vbwProtectorParameterString = "()"
4627           End If
4628           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4629       End If
' </VB WATCH>
4630       CalculateTEMCForce
' <VB WATCH>
4631       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4632       Exit Sub
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
4633       On Error GoTo vbwErrHandler
4634       Const VBWPROCNAME = "frmPLCData.txtTEMCThrustRigPressure_Change"
4635       If vbwProtector.vbwTraceProc Then
4636           Dim vbwProtectorParameterString As String
4637           If vbwProtector.vbwTraceParameters Then
4638               vbwProtectorParameterString = "()"
4639           End If
4640           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4641       End If
' </VB WATCH>
4642       CalculateTEMCForce
' <VB WATCH>
4643       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4644       Exit Sub
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
4645       On Error GoTo vbwErrHandler
4646       Const VBWPROCNAME = "frmPLCData.txtTEMCViscosity_Change"
4647       If vbwProtector.vbwTraceProc Then
4648           Dim vbwProtectorParameterString As String
4649           If vbwProtector.vbwTraceParameters Then
4650               vbwProtectorParameterString = "()"
4651           End If
4652           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4653       End If
' </VB WATCH>
4654       CalculateTEMCForce
' <VB WATCH>
4655       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4656       Exit Sub
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
4657       On Error GoTo vbwErrHandler
4658       Const VBWPROCNAME = "frmPLCData.txtV1_Change"
4659       If vbwProtector.vbwTraceProc Then
4660           Dim vbwProtectorParameterString As String
4661           If vbwProtector.vbwTraceParameters Then
4662               vbwProtectorParameterString = "()"
4663           End If
4664           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4665       End If
' </VB WATCH>
4666       txtV2.Text = txtV1.Text
4667       txtV3.Text = txtV1.Text
' <VB WATCH>
4668       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4669       Exit Sub
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
4670       On Error GoTo vbwErrHandler
4671       Const VBWPROCNAME = "frmPLCData.txtVibrationFile_gotfocus"
4672       If vbwProtector.vbwTraceProc Then
4673           Dim vbwProtectorParameterString As String
4674           If vbwProtector.vbwTraceParameters Then
4675               vbwProtectorParameterString = "()"
4676           End If
4677           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4678       End If
' </VB WATCH>
4679       On Error GoTo FileCancel
4680       If LenB(txtVibrationFile.Text) <> 0 Then
4681           CommonDialog1.filename = txtVibrationFile.Text
4682       End If
4683       CommonDialog1.ShowOpen
4684       txtVibrationFile.Text = CommonDialog1.filename
' <VB WATCH>
4685       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4686       Exit Sub
4687   FileCancel:
4688   On Error GoTo vbwErrHandler
4689       CommonDialog1.CancelError = False
' <VB WATCH>
4690       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4691       Exit Sub
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
4692       On Error GoTo vbwErrHandler
4693       Const VBWPROCNAME = "frmPLCData.ExportToExcel"
4694       If vbwProtector.vbwTraceProc Then
4695           Dim vbwProtectorParameterString As String
4696           If vbwProtector.vbwTraceParameters Then
4697               vbwProtectorParameterString = "()"
4698           End If
4699           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4700       End If
' </VB WATCH>

4701       Dim SaveFileName As String
4702       Dim WorkSheetName As String

4703       Dim I As Integer
4704       Dim iRowNo As Integer
4705       Dim sImp As String
4706       Dim ans As Integer

4707       Dim bCanShowSpeed As Boolean
4708       Dim CantShowReason As String

       'close any running excel processes
4709       Dim objWMIService, colProcesses
4710       Set objWMIService = GetObject("winmgmts:")
4711       Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name LIKE 'Excel%'")
4712       If colProcesses.Count > 0 Then
4713           Set xlApp = Excel.Application
4714       Else
               'use existing copy
       '        Set xlApp = New Excel.Application
4715           Set xlApp = CreateObject("Excel.Application")
4716       End If


4717       CommonDialog1.CancelError = True        'in case the user
4718       On Error GoTo ErrHandler                '  chooses the cancel button

           'set up dialog box
4719       CommonDialog1.DialogTitle = "Open Excel Files"
4720       CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|"  'show Excel files
4721       CommonDialog1.InitDir = App.Path
       '    CommonDialog1.InitDir = "C:\"    'in this directory
4722       CommonDialog1.ShowOpen                              'open the file selection dialog box

4723       If Dir(CommonDialog1.filename) = "" Then            'if the file name does not exist yet
4724           SaveFileName = CommonDialog1.filename           'get the name of the file
4725           If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
4726                xlApp.Workbooks.Close
4727           End If
               ' Create the Excel Workbook Object.
4728   On Error GoTo vbwErrHandler
4729           Set xlBook = xlApp.Workbooks.Add                'add a workbook
4730           WorkSheetName = NewWorkBook                                     'do some stuff for the new workbook
4731           ActiveWorkbook.CheckCompatibility = False
4732           xlApp.ActiveWorkbook.SaveAs filename:=SaveFileName, _
                                 FileFormat:=xlNormal                        'save the file
4733       Else                                                'the file name already exists
4734           SaveFileName = CommonDialog1.filename
               ' Create the Excel Workbook Object.
4735           If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
4736                xlApp.Workbooks.Close
4737           End If
4738           Set xlBook = xlApp.Workbooks.Open(SaveFileName)             'get the file name selected
4739           If GetWorksheetTabs(SaveFileName, WorkSheetName) = vbNo Then    'ask the user if he/she wants a new tab.
4740               MsgBox "File not overwritten.", vbOKOnly, "File not Opened"
' <VB WATCH>
4741       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4742               Exit Sub
4743           Else
4744           End If
4745       End If

4746   On Error GoTo vbwErrHandler

           'see if we can export Speed and SG and if we can, ask user if s/he wants it
           'assume that we can show speed calcs

4747       bCanShowSpeed = False
       'open the template and copy the data from the sheet
       '  excel file resides in ParentDirectoryName + "\Polar SG&Visc Correction5.xls"
           'write the data to the spreadsheet
4748       With xlApp

4749       Dim xlTemplateName As String
4750       xlTemplateName = ParentDirectoryName & sSGandViscSpreadsheetTemplate
4751       Dim xlTemplate As Excel.Workbook
4752       Set xlTemplate = xlApp.Workbooks.Open(xlTemplateName)
4753       Dim TemplateWS As Excel.Worksheet
4754       Dim sheetName As String
4755       sheetName = xlTemplate.Sheets(1).Name
4756       xlTemplate.Sheets(1).Copy After:=xlBook.Sheets(WorkSheetName)

4757       xlTemplate.Close savechanges:=False

4758       Set xlTemplate = Nothing

4759       Application.DisplayAlerts = False
4760       ActiveWorkbook.Worksheets(WorkSheetName).Delete
4761       Application.DisplayAlerts = True
4762       ActiveWorkbook.Worksheets(sheetName).Name = WorkSheetName

           'WorkSheetName = sheetName

           'first see if there is an entry in CalculatedRPM table for this frame size and voltage.
           ' if there is, get the coefficients, else make the coefficients 0

4763           Dim ACoef As Double
4764           Dim BCoef As Double
4765           Dim CCoef As Double

4766           Dim qy As New ADODB.Command
4767           Dim rs As New ADODB.Recordset
4768           qy.ActiveConnection = cnPumpData
4769           Dim VoltageForLookup As Integer
4770           If cmbVoltage.List(cmbVoltage.ListIndex) = "380" And cmbFrequency.List(cmbFrequency.ListIndex) = "50 Hz" Then
4771               VoltageForLookup = 460
4772           ElseIf cmbVoltage.List(cmbVoltage.ListIndex) <> "380" Then
4773               VoltageForLookup = cmbVoltage.List(cmbVoltage.ListIndex)
4774           End If
4775           qy.CommandText = "SELECT * FROM CalculatedRPM WHERE FrameNumber = '" & txtTEMCFrameNumber.Text & _
                          "' AND Voltage = '" & VoltageForLookup & "'"

4776           rs.CursorLocation = adUseClient
4777           rs.CursorType = adOpenStatic

4778           rs.Open qy
4779           If rs.RecordCount = 0 Then
4780               ACoef = 0
4781               BCoef = 0
4782               CCoef = 0
4783               MsgBox ("Cannot find coefficient data for Frame Number " & txtTEMCFrameNumber.Text & _
                          " AND Voltage = " & cmbVoltage.List(cmbVoltage.ListIndex) & _
                          " AND Frequency = " & cmbFrequency.List(cmbFrequency.ListIndex))
4784           Else
4785               ACoef = rs.Fields("A")
4786               BCoef = rs.Fields("B")
4787               CCoef = rs.Fields("C")
4788           End If


           'write header data

4789           .Range("A2").Select
4790           .ActiveCell.FormulaR1C1 = "Serial Number"
4791           .Range("C2").Select
4792           .ActiveCell.FormulaR1C1 = txtSN

4793           .Range("F1").Select
4794           .ActiveCell.FormulaR1C1 = "Customer"
4795           .Range("H1").Select
4796           .ActiveCell.FormulaR1C1 = txtShpNo

4797           .Range("A3").Select
4798           .ActiveCell.FormulaR1C1 = "Model"
4799           .Range("C3").Select
4800           .ActiveCell.FormulaR1C1 = txtModelNo

4801           .Range("F2").Select
4802           .ActiveCell.FormulaR1C1 = "Sales Order"
4803           .Range("H2").Select
4804           .ActiveCell.FormulaR1C1 = txtSalesOrderNumber

4805           .Range("A9").Select
4806           .ActiveCell.FormulaR1C1 = "Design Flow"
4807           .Range("C9").Select
4808           .ActiveCell.FormulaR1C1 = Val(txtDesignFlow)

4809           .Range("A10").Select
4810           .ActiveCell.FormulaR1C1 = "Design Head"
4811           .Range("C10").Select
4812           .ActiveCell.FormulaR1C1 = Val(txtDesignTDH)

4813           .Range("P13").Select
4814           .ActiveCell.FormulaR1C1 = "Barometric Pressure"
4815           .Range("R13").Select
4816           .ActiveCell.FormulaR1C1 = Val(txtInHgDisplay)

4817           .Range("P11").Select
4818           .ActiveCell.FormulaR1C1 = "Suction Gage Height"
4819           .Range("R11").Select
4820           .ActiveCell.FormulaR1C1 = Val(txtSuctHeight)

4821           .Range("P12").Select
4822           .ActiveCell.FormulaR1C1 = "Discharge Gage Height"
4823           .Range("R12").Select
4824           .ActiveCell.FormulaR1C1 = Val(txtDischHeight)

4825           .Range("A1").Select
4826           .ActiveCell.FormulaR1C1 = "Run Date"
4827           .Range("C1").Select
4828           .ActiveCell.FormulaR1C1 = cmbTestDate.List(cmbTestDate.ListIndex)

4829           .Range("D10:E10").Select
4830           With xlApp.Selection
4831               .HorizontalAlignment = xlCenter
4832               .VerticalAlignment = xlBottom
4833               .WrapText = False
4834               .Orientation = 0
4835               .AddIndent = False
4836               .IndentLevel = 0
4837               .ShrinkToFit = False
4838               .ReadingOrder = xlContext
4839               .MergeCells = False
4840           End With
4841           xlApp.Selection.Merge

               'determine rpm

4842           Dim RPMvalue As String
4843           If Mid$(Me.txtTEMCFrameNumber.Text, 2, 1) = "1" Then
               '1 says 2 pole
4844               If Me.cmbFrequency.ListIndex = 0 Then
                       '0 says 50Hz
4845                   RPMvalue = "2900"
4846               ElseIf Me.cmbFrequency.ListIndex = 1 Then
                       ' says 60Hz
4847                   RPMvalue = "3450"
4848               Else
                       'vfd or other, no rpm
4849                   RPMvalue = ""
4850               End If
4851           Else
               '2 says 4 pole
4852               If Me.cmbFrequency.ListIndex = 0 Then
                       '0 says 50Hz
4853                   RPMvalue = "1450"
4854               ElseIf Me.cmbFrequency.ListIndex = 1 Then
                       ' says 60Hz
4855                   RPMvalue = "1750"
4856               Else
                       'vfd or other, no rpm
4857                   RPMvalue = ""
4858               End If
4859           End If

       '        .Range("G1").Select
       '        .ActiveCell.FormulaR1C1 = "RPM"
       '        .Range("I1").Select
       '        .ActiveCell.FormulaR1C1 = RPMvalue

4860           .Range("A5").Select
4861           .ActiveCell.FormulaR1C1 = "Sp Gravity"
4862           .Range("C5").Select
4863           .ActiveCell.FormulaR1C1 = txtSpGr

4864           .Range("A6").Select
4865           .ActiveCell.FormulaR1C1 = "Viscosity"
4866           .Range("C6").Select
4867           .ActiveCell.FormulaR1C1 = txtViscosity

4868           .Range("F4").Select
4869           .ActiveCell.FormulaR1C1 = "Motor"
4870           .Range("H4").Select
4871           .ActiveCell.FormulaR1C1 = txtTEMCFrameNumber.Text

4872           .Range("H12").Select
4873           .ActiveCell.FormulaR1C1 = Me.txtCustPONum.Text

4874           .Range("F5").Select
4875           .ActiveCell.FormulaR1C1 = "Voltage"
4876           .Range("H5").Select
4877           .ActiveCell.FormulaR1C1 = cmbVoltage.List(cmbVoltage.ListIndex)

4878           .Range("K6").Select
4879           .ActiveCell.FormulaR1C1 = "End Play"
4880           .Range("M6").Select
4881           .ActiveCell.FormulaR1C1 = Val(txtEndPlay)

4882           .Range("K7").Select
4883           .ActiveCell.FormulaR1C1 = "G-Gap"
4884           .Range("M7").Select
4885           .ActiveCell.FormulaR1C1 = txtGGap.Text

4886           .Range("A8").Select
4887           .ActiveCell.FormulaR1C1 = "Design Pressure"
4888           .Range("C8").Select
4889           Dim DesPress As String
4890           DesPress = cmbTEMCDesignPressure.List(cmbTEMCDesignPressure.ListIndex)
4891           Dim j As Integer
4892           j = InStrRev(DesPress, "-")
4893           .ActiveCell.FormulaR1C1 = Mid$(DesPress, j + 2)

       '        .Range("G8").Select
       '        .ActiveCell.FormulaR1C1 = "Stator Fill"
       '        .Range("I8").Select
       '        .ActiveCell.FormulaR1C1 = "Dry"

4894           .Range("K4").Select
4895           .ActiveCell.FormulaR1C1 = "Circulation Path"
4896           .Range("M4").Select
4897           .ActiveCell.FormulaR1C1 = cmbTEMCModel.List(cmbTEMCModel.ListIndex)

4898           .Range("M8").Select
4899           .ActiveCell.FormulaR1C1 = txtNPSHr.Text

4900           .Range("K1").Select
4901           .ActiveCell.FormulaR1C1 = "Impeller Dia"
4902           .Range("M1").Select


       '        If LenB(txtImpTrim) <> 0 Then
       '            .ActiveCell.FormulaR1C1 = Val(txtImpTrim)
       '        Else
       '            .ActiveCell.FormulaR1C1 = Val(txtImpellerDia)
       '        End If
       '
4903           If chkTrimmed.value = 1 Then
4904               If Val(txtImpTrim.Text) <> 0 Then
4905                   .ActiveCell.FormulaR1C1 = txtImpTrim
4906               Else
4907                   .ActiveCell.FormulaR1C1 = txtImpellerDia
4908               End If
4909           Else
4910               .ActiveCell.FormulaR1C1 = txtImpellerDia
4911           End If



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

4912           .Range("P9").Select
4913           .ActiveCell.FormulaR1C1 = "Suction Dia"
4914           .Range("R9").Select
4915           .ActiveCell.FormulaR1C1 = cmbSuctDia.List(cmbSuctDia.ListIndex)

4916           .Range("P10").Select
4917           .ActiveCell.FormulaR1C1 = "Discharge Dia"
4918           .Range("R10").Select
4919           .ActiveCell.FormulaR1C1 = cmbDischDia.List(cmbDischDia.ListIndex)

4920           .Range("A11").Select
4921           .ActiveCell.FormulaR1C1 = "Test Spec"
4922           .Range("C11").Select
4923           .ActiveCell.FormulaR1C1 = cmbTestSpec.List(cmbTestSpec.ListIndex)

4924           .Range("K3").Select
4925           .ActiveCell.FormulaR1C1 = "Impeller Feathered"
4926           .Range("M3").Select
4927           If chkFeathered.value = 1 Then
4928               .ActiveCell.FormulaR1C1 = "Yes"
4929           Else
4930               .ActiveCell.FormulaR1C1 = "No"
4931           End If

4932           .Range("K2").Select
4933           .ActiveCell.FormulaR1C1 = "Disch Orifice"
4934           .Range("M2").Select
4935           If chkOrifice.value = 1 Then
4936               .ActiveCell.FormulaR1C1 = Val(txtOrifice)
4937           Else
4938               .ActiveCell.FormulaR1C1 = "None"
4939           End If


4940           .Range("K5").Select
4941           .ActiveCell.FormulaR1C1 = "Circulation Orifice"
4942           .Range("M5").Select
4943           If chkCircOrifice.value = 1 Then
4944               .ActiveCell.FormulaR1C1 = Val(txtCircOrifice)
4945           Else
4946               .ActiveCell.FormulaR1C1 = "None"
4947           End If

4948           .Range("A13").Select
4949           .ActiveCell.FormulaR1C1 = "Other Mods"
4950           .Range("C13").Select
4951           .ActiveCell.FormulaR1C1 = txtOtherMods

4952           .Range("A14").Select
4953           .ActiveCell.FormulaR1C1 = "Remarks"
4954           .Range("C14").Select
4955           .ActiveCell.FormulaR1C1 = txtRemarks

4956           .Range("A15").Select
4957           .ActiveCell.FormulaR1C1 = "Test Setup Remarks"
4958           .Range("C15").Select
4959           .ActiveCell.FormulaR1C1 = txtTestSetupRemarks

4960           .Range("P1").Select
4961           .ActiveCell.FormulaR1C1 = "Suct ID"
4962           .Range("R1").Select
4963           .ActiveCell.FormulaR1C1 = cmbSuctionPressureTransducer.List(cmbSuctionPressureTransducer.ListIndex)

4964           .Range("P2").Select
4965           .ActiveCell.FormulaR1C1 = "Disch ID"
4966           .Range("R2").Select
4967           .ActiveCell.FormulaR1C1 = cmbDischargePressureTransducer.List(cmbDischargePressureTransducer.ListIndex)

4968           .Range("P3").Select
4969           .ActiveCell.FormulaR1C1 = "Temp ID"
4970           .Range("R3").Select
4971           .ActiveCell.FormulaR1C1 = cmbTemperatureTransducer.List(cmbTemperatureTransducer.ListIndex)

4972           .Range("P4").Select
4973           .ActiveCell.FormulaR1C1 = "Circ Flow ID"
4974           .Range("R4").Select
4975           .ActiveCell.FormulaR1C1 = cmbCirculationFlowMeter.List(cmbCirculationFlowMeter.ListIndex)

4976           .Range("P5").Select
4977           .ActiveCell.FormulaR1C1 = "Flow ID"
4978           .Range("R5").Select
4979           .ActiveCell.FormulaR1C1 = cmbFlowMeter.List(cmbFlowMeter.ListIndex)

4980           .Range("P6").Select
4981           .ActiveCell.FormulaR1C1 = "Analyzer ID"
4982           .Range("R6").Select
4983           .ActiveCell.FormulaR1C1 = cmbAnalyzerNo.List(cmbAnalyzerNo.ListIndex)

4984           .Range("P7").Select
4985           .ActiveCell.FormulaR1C1 = "Loop ID"
4986           .Range("R7").Select
4987           .ActiveCell.FormulaR1C1 = cmbLoopNumber.List(cmbLoopNumber.ListIndex)

4988           .Range("A4").Select
4989           .ActiveCell.FormulaR1C1 = "Fluid"
4990           .Range("C4").Select
4991           .ActiveCell.FormulaR1C1 = txtLiquid.Text

4992           .Range("F3").Select
4993           .ActiveCell.FormulaR1C1 = "Cust PN"
4994           .Range("H3").Select
       '        .ActiveCell.FormulaR1C1 = txtRMA.Text
4995           If rsPumpData.Fields("RVSPartNo") <> "" Then
4996               .ActiveCell.FormulaR1C1 = rsPumpData.Fields("RVSPartNo")
4997           End If
4998           If rsPumpData.Fields("CustPN") <> "" Then
4999               .ActiveCell.FormulaR1C1 = rsPumpData.Fields("CustPN")
5000           End If

5001           .Range("A7").Select
5002           .ActiveCell.FormulaR1C1 = "Temperature"
5003           .Range("C7").Select
5004           .ActiveCell.FormulaR1C1 = txtLiquidTemperature.Text

5005           .Range("F6").Select
5006           .ActiveCell.FormulaR1C1 = "Frequency"
5007           .Range("H6").Select
5008           If UCase(cmbFrequency.List(cmbFrequency.ListIndex)) = "VFD" Then
5009               .ActiveCell.FormulaR1C1 = Val(Me.txtVFDFreq)
5010           Else
5011               .ActiveCell.FormulaR1C1 = Val(cmbFrequency.List(cmbFrequency.ListIndex))
5012           End If
       '        .Range("K2").Select
       '        .ActiveCell.FormulaR1C1 = "Disch Orifice"
       '        .Range("M2").Select
       '        .ActiveCell.FormulaR1C1 = txtOrifice.Text

       '        .Range("K12").Select
       '        .ActiveCell.FormulaR1C1 = "Flow Orifice"
       '        .Range("L12").Select
       '        .ActiveCell.FormulaR1C1 = txtCircOrifice.Text

5013           .Range("P8").Select
5014           .ActiveCell.FormulaR1C1 = "PLC No"
5015           .Range("R8").Select
5016           .ActiveCell.FormulaR1C1 = cmbPLCNo.List(cmbPLCNo.ListIndex)

5017           .Range("F7").Select
5018           .ActiveCell.FormulaR1C1 = "Phases"
5019           .Range("H7").Select
5020           .ActiveCell.FormulaR1C1 = txtNoPhases.Text

5021           .Range("F8").Select
5022           .ActiveCell.FormulaR1C1 = "Poles"
5023           .Range("H8").Select
5024           .ActiveCell.FormulaR1C1 = 2 * Val(Left$(Right$(txtTEMCFrameNumber.Text, 2), 1))

5025           .Range("F9").Select
5026           .ActiveCell.FormulaR1C1 = "Rated Current"
5027           .Range("H9").Select
5028           .ActiveCell.FormulaR1C1 = txtAmps.Text

5029           .Range("F10").Select
5030           .ActiveCell.FormulaR1C1 = "Rated Input Power"
5031           .Range("H10").Select
5032           .ActiveCell.FormulaR1C1 = txtRatedInputPower.Text

5033           .Range("F11").Select
5034           .ActiveCell.FormulaR1C1 = "Insulation Class"
5035           .Range("H11").Select
5036           .ActiveCell.FormulaR1C1 = txtThermalClass.Text

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

5037           .Range("A17").Select
5038           .ActiveCell.FormulaR1C1 = "Flow"
5039           .Range("A18").Select
5040           .ActiveCell.FormulaR1C1 = "(GPM)"

5041           .Range("B17").Select
5042           .ActiveCell.FormulaR1C1 = "TDH"
5043           .Range("B18").Select
5044           .ActiveCell.FormulaR1C1 = "(Ft)"

5045           .Range("C17").Select
5046           .ActiveCell.FormulaR1C1 = "KW"

5047           .Range("D17").Select
5048           .ActiveCell.FormulaR1C1 = "Ave"
5049           .Range("D18").Select
5050           .ActiveCell.FormulaR1C1 = "Volts"

5051           .Range("E17").Select
5052           .ActiveCell.FormulaR1C1 = "Ave"
5053           .Range("E18").Select
5054           .ActiveCell.FormulaR1C1 = "Amps"

5055           .Range("F17").Select
5056           .ActiveCell.FormulaR1C1 = "Power"
5057           .Range("F18").Select
5058           .ActiveCell.FormulaR1C1 = "Factor"

5059           .Range("G17").Select
5060           .ActiveCell.FormulaR1C1 = "Overall"
5061           .Range("G18").Select
5062           .ActiveCell.FormulaR1C1 = "Eff"

5063           .Range("H17").Select
5064           .ActiveCell.FormulaR1C1 = "Measured"
5065           .Range("H18").Select
5066           .ActiveCell.FormulaR1C1 = "RPM"

5067           .Range("I17").Select
5068           .ActiveCell.FormulaR1C1 = "Calculated"
5069           .Range("I18").Select
5070           .ActiveCell.FormulaR1C1 = "RPM"

5071           .Range("J17").Select
5072           .ActiveCell.FormulaR1C1 = "Suction"
5073           .Range("J18").Select
5074           .ActiveCell.FormulaR1C1 = "Temp(F)"

5075           .Range("K17").Select
5076           .ActiveCell.FormulaR1C1 = "Disch"
5077           .Range("K18").Select
5078           .ActiveCell.FormulaR1C1 = "Pressure"

5079           .Range("L17").Select
5080           .ActiveCell.FormulaR1C1 = "Suction"
5081           .Range("L18").Select
5082           .ActiveCell.FormulaR1C1 = "Pressure"

5083           .Range("M17").Select
5084           .ActiveCell.FormulaR1C1 = "Vel"
5085           .Range("M18").Select
5086           .ActiveCell.FormulaR1C1 = "Head"

5087           .Range("N17").Select
5088           .ActiveCell.FormulaR1C1 = "Axial"
5089           .Range("N18").Select
5090           .ActiveCell.FormulaR1C1 = "Position"

5091           .Range("O17").Select
5092           .ActiveCell.FormulaR1C1 = "Pct of"
5093           .Range("O18").Select
5094           .ActiveCell.FormulaR1C1 = "End Play"

5095           .Range("P17").Select
5096           .ActiveCell.FormulaR1C1 = "Hydraulic"
5097           .Range("P18").Select
5098           .ActiveCell.FormulaR1C1 = "Efficiency"

       '        .Range("P17").Select
       '        .ActiveCell.FormulaR1C1 = "Circ"
       '        .Range("P18").Select
       '        .ActiveCell.FormulaR1C1 = "Flow"

5099           .Range("Q17").Select
5100           .ActiveCell.FormulaR1C1 = "Motor"
5101           .Range("Q18").Select
5102           .ActiveCell.FormulaR1C1 = "Efficiency"

5103           .Range("S17").Select
5104           .ActiveCell.FormulaR1C1 = "NPSHa"

5105           .Range("T17").Select
5106           .ActiveCell.FormulaR1C1 = "Phase 1"
5107           .Range("T18").Select
5108           .ActiveCell.FormulaR1C1 = "Current"

5109           .Range("U17").Select
5110           .ActiveCell.FormulaR1C1 = "Phase 2"
5111           .Range("U18").Select
5112           .ActiveCell.FormulaR1C1 = "Current"

5113           .Range("V17").Select
5114           .ActiveCell.FormulaR1C1 = "Phase 3"
5115           .Range("V18").Select
5116           .ActiveCell.FormulaR1C1 = "Current"

5117           .Range("W17").Select
5118           .ActiveCell.FormulaR1C1 = "Phase 1"
5119           .Range("W18").Select
5120           .ActiveCell.FormulaR1C1 = "Voltage"

5121           .Range("X17").Select
5122           .ActiveCell.FormulaR1C1 = "Phase 2"
5123           .Range("X18").Select
5124           .ActiveCell.FormulaR1C1 = "Voltage"

5125           .Range("Y17").Select
5126           .ActiveCell.FormulaR1C1 = "Phase 3"
5127           .Range("Y18").Select
5128           .ActiveCell.FormulaR1C1 = "Voltage"

5129           .Range("Z17").Select
5130           .ActiveCell.FormulaR1C1 = "'" & txtTitle(20).Text

5131           .Range("Z18").Select
5132           .ActiveCell.FormulaR1C1 = "'" & txtTitle(21).Text

5133           .Range("AA17").Select
5134           .ActiveCell.FormulaR1C1 = "'" & txtTitle(22).Text

5135           .Range("AA18").Select
5136           .ActiveCell.FormulaR1C1 = "'" & txtTitle(23).Text

5137           .Range("AB17").Select
5138           .ActiveCell.FormulaR1C1 = "'" & txtTitle(24).Text

5139           .Range("AB18").Select
5140           .ActiveCell.FormulaR1C1 = "'" & txtTitle(25).Text

5141           .Range("AC17").Select
5142           .ActiveCell.FormulaR1C1 = "HR"

5143           .Range("AC18").Select
5144           .ActiveCell.FormulaR1C1 = "(ft)"

5145           .Range("AD17").Select
5146           .ActiveCell.FormulaR1C1 = "'" & txtTitle(26).Text

5147           .Range("AD18").Select
5148           .ActiveCell.FormulaR1C1 = "'" & txtTitle(27).Text

5149           .Range("AE17").Select
5150           .ActiveCell.FormulaR1C1 = "TRG"
5151           .Range("AE18").Select
5152           .ActiveCell.FormulaR1C1 = "Position"

5153           .Range("AF17").Select
5154           .ActiveCell.FormulaR1C1 = "Thrust"

5155           .Range("AG17").Select
5156           .ActiveCell.FormulaR1C1 = "F/R"

5157           .Range("AH17").Select
5158           .ActiveCell.FormulaR1C1 = "Moment"
5159           .Range("AH18").Select
5160           .ActiveCell.FormulaR1C1 = "Arm"

5161           .Range("AI17").Select
5162           .ActiveCell.FormulaR1C1 = "Rig"
5163           .Range("AI18").Select
5164           .ActiveCell.FormulaR1C1 = "Pressure"

       '        .Range("AI17").Select
       '        .ActiveCell.FormulaR1C1 = "Viscosity"

5165           .Range("AJ19").Select
5166           .ActiveCell.FormulaR1C1 = "Rear"
5167           .Range("AJ18").Select
5168           .ActiveCell.FormulaR1C1 = "Force"

5169           .Range("AK17").Select
5170           .ActiveCell.FormulaR1C1 = "PV"

5171           .Range("R17").Select
5172           .ActiveCell.FormulaR1C1 = "Shaft"
5173           .Range("R18").Select
5174           .ActiveCell.FormulaR1C1 = "Power"

       '        .Range("AM17").Select
       '        .ActiveCell.FormulaR1C1 = "Pct Full"
       '        .Range("AM18").Select
       '        .ActiveCell.FormulaR1C1 = "Scale"

5175           .Range("AL17").Select
5176           .ActiveCell.FormulaR1C1 = "NPSHr"

5177           .Range("AM17").Select
5178           .ActiveCell.FormulaR1C1 = "Remarks"




               'now output the data

5179           iRowNo = 20

5180           rsEff.MoveFirst
5181           For I = 1 To frmPLCData.UpDown2.value
5182               .Range("A" & iRowNo).Select
5183               .ActiveCell.FormulaR1C1 = rsEff.Fields("Flow")

5184               .Range("B" & iRowNo).Select
5185               .ActiveCell.FormulaR1C1 = rsEff.Fields("TDH")

5186               .Range("C" & iRowNo).Select
5187               .ActiveCell.FormulaR1C1 = rsEff.Fields("KW")

5188               .Range("D" & iRowNo).Select
5189               .ActiveCell.FormulaR1C1 = rsEff.Fields("Volts")

5190               .Range("E" & iRowNo).Select
5191               .ActiveCell.FormulaR1C1 = rsEff.Fields("Amps")

5192               .Range("F" & iRowNo).Select
5193               .ActiveCell.FormulaR1C1 = rsEff.Fields("PowerFactor")

5194               .Range("G" & iRowNo).Select
5195               .ActiveCell.FormulaR1C1 = rsEff.Fields("OverallEfficiency")

5196               .Range("H" & iRowNo).Select
5197               .ActiveCell.FormulaR1C1 = rsEff.Fields("RPM")

5198               .Range("I" & iRowNo).Select
                   'use the coefficients from above to calculate rpm
5199               Dim f As Double
5200               f = .Range("H6").value
5201               .ActiveCell.FormulaR1C1 = (Val(f) / 60) * (ACoef * (rsEff.Fields("KW")) ^ 2 + BCoef * (rsEff.Fields("KW")) + CCoef)

5202               .Range("J" & iRowNo).Select
5203               .ActiveCell.FormulaR1C1 = rsEff.Fields("Temperature")

5204               .Range("K" & iRowNo).Select
5205               .ActiveCell.FormulaR1C1 = rsEff.Fields("DischPress")

5206               .Range("L" & iRowNo).Select
5207               .ActiveCell.FormulaR1C1 = rsEff.Fields("SuctPress")

5208               .Range("M" & iRowNo).Select
5209               .ActiveCell.FormulaR1C1 = rsEff.Fields("VelocityHead")

5210               .Range("N" & iRowNo).Select
5211               .ActiveCell.FormulaR1C1 = rsEff.Fields("Pos")

5212               .Range("O" & iRowNo).Select
5213               .ActiveCell.FormulaR1C1 = 100 * rsEff.Fields("Pos") / Val(txtEndPlay)

5214               .Range("P" & iRowNo).Select
5215               .ActiveCell.FormulaR1C1 = rsEff.Fields("HydraulicEfficiency")

       '            .Range("P" & iRowNo).Select
       '            .ActiveCell.FormulaR1C1 = rsEff.Fields("CircFlow")

5216               .Range("Q" & iRowNo).Select
5217               .ActiveCell.FormulaR1C1 = rsEff.Fields("MotorEfficiency")

5218               .Range("S" & iRowNo).Select
5219               .ActiveCell.FormulaR1C1 = rsEff.Fields("NPSHa")

5220               .Range("T" & iRowNo).Select
5221               .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentA")

5222               .Range("U" & iRowNo).Select
5223               .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentB")

5224               .Range("V" & iRowNo).Select
5225               .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentC")

5226               .Range("W" & iRowNo).Select
5227               .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageA")

5228               .Range("X" & iRowNo).Select
5229               .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageB")

5230               .Range("Y" & iRowNo).Select
5231               .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageC")

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

5232               .Range("Z" & iRowNo).Select
5233               .ActiveCell.FormulaR1C1 = rsEff.Fields("CircFlow")

5234               .Range("AA" & iRowNo).Select
5235               .ActiveCell.FormulaR1C1 = rsEff.Fields("RBHTemp")

5236               .Range("AB" & iRowNo).Select
5237               .ActiveCell.FormulaR1C1 = rsEff.Fields("RBHPress")

5238               .Range("AC" & iRowNo).Select
5239               .ActiveCell.FormulaR1C1 = (rsEff.Fields("RBHPress") - rsEff.Fields("SuctPress")) * 2.31

5240               .Range("AD" & iRowNo).Select
5241               .ActiveCell.FormulaR1C1 = rsEff.Fields("AI4")

5242               .Range("AE" & iRowNo).Select
5243               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCTRG")

5244               .Range("AF" & iRowNo).Select
5245               If rsEff.Fields("TEMCFrontThrust") = 0 Then
5246                   If rsEff.Fields("TEMCRearThrust") = 0 Then
5247                       .ActiveCell.FormulaR1C1 = " "
5248                       .Range("AG" & iRowNo).Select
5249                       .ActiveCell.FormulaR1C1 = " "
5250                   Else
5251                       .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCRearThrust")
5252                       .Range("AG" & iRowNo).Select
5253                       .ActiveCell.FormulaR1C1 = "R"
5254                   End If
5255               Else
5256                   .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCFrontThrust")
5257                   .Range("AG" & iRowNo).Select
5258                   .ActiveCell.FormulaR1C1 = "F"
5259               End If

5260               .Range("AH" & iRowNo).Select
5261               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCMomentArm")

5262               .Range("AI" & iRowNo).Select
5263               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCThrustRigPressure")

       '            .Range("AJ" & iRowNo).Select
       '            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCViscosity")

5264               .Range("AJ" & iRowNo).Select
5265               If rsEff.Fields("TEMCForceDirection") = "F" Then
5266                   .ActiveCell.FormulaR1C1 = -rsEff.Fields("TEMCCalculatedForce")
5267               Else
5268                   .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCCalculatedForce")
5269               End If

5270               .Range("AK" & iRowNo).Select
5271               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCPV")

5272               .Range("R" & iRowNo).Select
5273               .ActiveCell.FormulaR1C1 = rsEff.Fields("KW") * rsEff.Fields("MotorEfficiency") / 100

5274               .Range("AL" & iRowNo).Select
5275               .ActiveCell.FormulaR1C1 = rsEff.Fields("NPSHr")

       '            If RatedKW = 999 Then
       '                .ActiveCell.FormulaR1C1 = ""
       '            Else
       '                .ActiveCell.FormulaR1C1 = (rsEff.Fields("KW") * rsEff.Fields("MotorEfficiency")) / (1 * RatedKW)
       '            End If

5276               .Range("AM" & iRowNo).Select
5277               .ActiveCell.FormulaR1C1 = rsEff.Fields("Remarks")


5278               rsEff.MoveNext
5279               iRowNo = iRowNo + 1
5280           Next I

5281           .Range("A20:AS30").Select
5282           .Selection.NumberFormat = "0.00"

           'set up formulas to calculate BEP
           '  first, plot 2nd order polynomial for flow vs hydraulic efficiency
           '  the formulas for doing that are in E68, F68 and G68
           '  only want the formulas to point to the number of points in the test data, so use frmPLCData.CWNumEdit2.value
           '
5283       Dim AColumnRow As String
5284       Dim PColumnRow As String

5285       AColumnRow = "A" & Trim(str(19 + frmPLCData.UpDown2.value))
5286       PColumnRow = "P" & Trim(str(19 + frmPLCData.UpDown2.value))

5287           .Range("E68").Select
5288           .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1)"

5289           .Range("F68").Select
5290           .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1,2)"

5291           .Range("G68").Select
5292           .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1,3)"

           'export balance holes
5293       If boGotBalanceHoles Then
5294           If rsBalanceHoles.State = adStateClosed Then
5295               rsBalanceHoles.ActiveConnection = cnPumpData
5296               rsBalanceHoles.Open
5297           End If 'rsBalanceHoles.State = adStateClosed

5298           If rsBalanceHoles.RecordCount <> 0 Then

5299               .Range("K9:N9").Merge
5300               .Range("K9:N9").Formula = "Balance Hole Data"
5301               .Range("K9:N9").HorizontalAlignment = xlCenter

5302               .Range("K10").Select
5303               .ActiveCell.Formula = "Date"

5304               .Range("L10").Select
5305               .ActiveCell.Formula = "Number"

5306               .Range("M10").Select
5307               .ActiveCell.Formula = "Diameter"

5308               .Range("N10").Select
5309               .ActiveCell.Formula = "Bolt Circle"

5310               iRowNo = 11

5311               If rsBalanceHoles.RecordCount > 3 Then
5312                   For I = 1 To rsBalanceHoles.RecordCount - 3
5313                       Rows("13:13").Select
5314                       Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
5315                   Next I
5316               End If

5317               rsBalanceHoles.MoveFirst
5318               For I = 1 To rsBalanceHoles.RecordCount

5319                   .Range("K" & iRowNo).Select
5320                   .ActiveCell.Formula = rsBalanceHoles.Fields("Date")
5321                   .ActiveCell.NumberFormat = "m/d/yy h:mm AM/PM;@"
5322                   .Range("L" & iRowNo).Select
5323                   .ActiveCell = rsBalanceHoles.Fields("Number")
5324                   .ActiveCell.NumberFormat = "0"
5325                   .Range("M" & iRowNo).Select
5326                   If IsNumeric(rsBalanceHoles.Fields("Diameter1")) Then
5327                       .ActiveCell = Val(rsBalanceHoles.Fields("Diameter1"))
5328                       .ActiveCell.NumberFormat = "0.0000"
5329                   Else
5330                       .ActiveCell = rsBalanceHoles.Fields("Diameter1")
5331                   End If

5332                   .Range("N" & iRowNo).Select
5333                   If IsNumeric(rsBalanceHoles.Fields("BoltCircle1")) Then
5334                       .ActiveCell = Val(rsBalanceHoles.Fields("BoltCircle1"))
5335                       .ActiveCell.NumberFormat = "0.0000"
5336                   Else
5337                       .ActiveCell = rsBalanceHoles.Fields("BoltCircle1")
5338                   End If

5339                   rsBalanceHoles.MoveNext
5340                   iRowNo = iRowNo + 1
5341               Next I
5342               .Range("K10:N" & iRowNo - 1).Select
5343               With .Selection.Interior
5344                   .ColorIndex = 34
5345                   .Pattern = xlSolid
5346               End With
5347           End If 'rsBalanceHoles.RecordCount <> 0
5348       End If ' boGotBalanceHoles

           'plot graphs

5349       Dim SeriesName As String
5350       Dim XVals As String
5351       Dim YVals As String
5352       Dim RowNo As Long
5353       Dim RowStr As String
5354       Dim LastPoint As Integer
5355       Dim LineType As String
5356       Dim AxisGroup As Integer
5357       Dim LabelPos As Integer
5358       Dim LineColor As Long

5359           .ActiveSheet.ChartObjects("HydRepChart").Activate
5360           Dim S As Series
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
5361           Dim aq As Double
5362           Range("AQ56", "AQ71").Select
5363           aq = .Max(Selection)
5364           Dim ax As Double
5365           Range("AX56", "AX71").Select
5366           ax = .Max(Selection)

               'then current (as and az)
5367           Dim at As Double
5368           Range("AS56", "AS71").Select
5369           at = .Max(Selection)
5370           Dim ba As Double
5371           Range("AZ56", "AZ71").Select
5372           ba = .Max(Selection)

5373           Dim CurrentScaleMax As Integer
5374           Dim TDHScaleMax As Integer

5375           Dim MaxTDH As Integer
5376           With Application.WorksheetFunction
5377               If aq > ax Then
5378                   MaxTDH = .Ceiling(aq, 25)
5379               Else
5380                   MaxTDH = .Ceiling(ax, 25)
5381               End If
5382           End With

5383           Dim MaxCurrent As Integer
5384           With Application.WorksheetFunction
5385               If at > ba Then
5386                   Select Case at
                           Case Is <= 5
5387                           CurrentScaleMax = 5

5388                       Case Is <= 10
5389                           CurrentScaleMax = 10

5390                       Case Else
5391                           CurrentScaleMax = 25
5392                   End Select

5393                   MaxCurrent = .Ceiling(at, CurrentScaleMax)
5394               Else
5395                  Select Case ba
                           Case Is <= 5
5396                           CurrentScaleMax = 5

5397                       Case Is <= 10
5398                           CurrentScaleMax = 10

5399                       Case Else
5400                           CurrentScaleMax = 25
5401                   End Select

5402                   MaxCurrent = .Ceiling(ba, CurrentScaleMax)
5403               End If
5404           End With

5405           ActiveSheet.ChartObjects("HydRepChart").Activate
5406            Dim ShtName As String
5407            ShtName = "'" & ActiveSheet.Name & "'"

5408           RowStr = 56 + 15
5409            For I = 1 To 8

5410                Select Case I
                        Case 1
5411                        SeriesName = "=""TDH"""
5412                        XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
5413                        YVals = "=" & ShtName & "!$AQ$56:$AQ$" & RowStr
5414                        LineType = msoLineSolid
5415                        AxisGroup = 1
5416                        LabelPos = xlLabelPositionRight
5417                        LineColor = vbBlue

5418                    Case 2
5419                        SeriesName = "=""Input Power"""
5420                        XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
5421                        YVals = "=" & ShtName & "!$AR$56:$AR$" & RowStr
5422                        LineType = msoLineSolid
5423                        AxisGroup = 2
5424                        LabelPos = xlLabelPositionRight
5425                        LineColor = vbRed

5426                    Case 3
5427                        SeriesName = "=""Current"""
5428                        XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
5429                        YVals = "=" & ShtName & "!$AS$56:$AS$" & RowStr
5430                        LineType = msoLineSolid
5431                        AxisGroup = 2
5432                        LabelPos = xlLabelPositionRight
5433                        LineColor = vbGreen

5434                    Case 4
       '                     SeriesName = "=""Overall Eff"""
       '                     XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
       '                     YVals = "=" & ShtName & "!$AT$56:$AT$" & RowStr
       '                     LineType = msoLineSolid
       '                     AxisGroup = 2
       '                     LabelPos = xlLabelPositionRight
       '                     LineColor = vbCyan

5435                    Case 5
5436                        SeriesName = "=""TDH (Adj)"""
5437                        XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
5438                        YVals = "=" & ShtName & "!$AX$56:$AX$" & RowStr
5439                        LineType = msoLineDash
5440                        AxisGroup = 1
5441                        LabelPos = xlLabelPositionBelow
5442                        LineColor = vbBlue

5443                    Case 6
5444                        SeriesName = "=""Input Power (Adj)"""
5445                        XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
5446                        YVals = "=" & ShtName & "!$AY$56:$AY$" & RowStr
5447                        LineType = msoLineDash
5448                        AxisGroup = 2
5449                        LabelPos = xlLabelPositionBelow
5450                        LineColor = vbRed

5451                    Case 7
5452                        SeriesName = "=""Current (Adj)"""
5453                        XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
5454                        YVals = "=" & ShtName & "!$AZ$56:$AZ$" & RowStr
5455                        LineType = msoLineDash
5456                        AxisGroup = 2
5457                        LabelPos = xlLabelPositionBelow
5458                        LineColor = vbGreen

5459                    Case 8
       '                     SeriesName = "=""Overall Eff (Adj)"""
       '                     XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
       '                     YVals = "=" & ShtName & "!$BA$56:$BA$" & RowStr
       '                     LineType = msoLineDash
       '                     AxisGroup = 2
       '                     LabelPos = xlLabelPositionBelow
       '                     LineColor = vbCyan

5460               End Select
5461               LastPoint = 16
5462               ActiveChart.SeriesCollection.NewSeries
5463               ActiveChart.SeriesCollection(I).Name = SeriesName
5464               ActiveChart.SeriesCollection(I).XValues = XVals
5465               ActiveChart.SeriesCollection(I).Values = YVals
5466               ActiveChart.SeriesCollection(I).Select
5467               ActiveChart.SeriesCollection(I).Points(LastPoint).Select
5468               ActiveChart.SeriesCollection(I).Points(LastPoint).ApplyDataLabels
5469               ActiveChart.SeriesCollection(I).Points(LastPoint).DataLabel.Select
5470               If I < 5 Then
5471                   Selection.ShowSeriesName = True
5472                   Selection.Position = LabelPos
5473               Else
5474                   Selection.ShowSeriesName = False
5475               End If
5476               Selection.ShowValue = False
5477               ActiveChart.SeriesCollection(I).ChartType = xlXYScatterSmoothNoMarkers
5478               ActiveChart.SeriesCollection(I).Select
5479               With Selection.Format.line
5480                   .Visible = msoTrue
5481                   .DashStyle = LineType
5482                   .ForeColor.RGB = LineColor
5483               End With


5484               ActiveChart.SeriesCollection(I).AxisGroup = AxisGroup
5485               ActiveChart.SeriesCollection(I).DataLabels.Font.Size = 8
5486               ActiveChart.SeriesCollection(I).DataLabels.Font.Name = "Arial"
5487           Next I

               'show design point
5488           SeriesName = "=""Design Point"""
5489           XVals = "=" & ShtName & "!$L$63"
5490           YVals = "=" & ShtName & "!$L$64"
5491           LineType = msoLineSolid
5492           AxisGroup = 1
5493           ActiveChart.SeriesCollection.NewSeries
5494           ActiveChart.SeriesCollection(I).Name = SeriesName
5495           ActiveChart.SeriesCollection(I).XValues = XVals
5496           ActiveChart.SeriesCollection(I).Values = YVals
5497           ActiveChart.SeriesCollection(I).Select

5498           Selection.MarkerStyle = 4
5499           Selection.MarkerSize = 7
5500           With Selection.Format.line
5501               .Visible = msoTrue
5502               .Weight = 2.25
5503               .ForeColor.RGB = vbBlack
5504           End With


5505           ActiveChart.Axes(xlValue).Select
5506           ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
5507           ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True

5508           ActiveChart.Axes(xlValue).MaximumScale = MaxTDH
5509           ActiveChart.Axes(xlValue).MinimumScale = 0
5510           ActiveChart.Axes(xlValue).MajorUnit = Int(MaxTDH / 5)
5511           Selection.TickLabels.NumberFormat = "0"

5512           ActiveChart.Axes(xlValue, xlSecondary).Select
5513           ActiveChart.Axes(xlValue, xlSecondary).MinimumScaleIsAuto = True
5514           ActiveChart.Axes(xlValue, xlSecondary).MaximumScaleIsAuto = True

5515           ActiveChart.Axes(xlValue, xlSecondary).MaximumScale = MaxCurrent
5516           ActiveChart.Axes(xlValue, xlSecondary).MinimumScale = 0
5517           ActiveChart.Axes(xlValue, xlSecondary).MajorUnit = Int(MaxCurrent / 5)
5518           Selection.TickLabels.NumberFormat = "0"

5519           ActiveChart.Axes(xlValue, xlSecondary).HasTitle = True
5520           ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = "Input Power (kW)-Current (A)"
       '        ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = "Input Power (kW)-Current (A)-Overall Efficiency (%)"
5521           ActiveChart.SetElement (msoElementSecondaryValueAxisTitleRotated)
               'ActiveSheet.PageSetup.PrintArea = "$CA$1:$CI$50"

5522           Range("A1").Select

               'delete all macros in the excel file

               ' Declare variables to access the macros in the workbook.
5523           Dim objProject As VBIDE.VBProject
5524           Dim objComponent As VBIDE.VBComponent
5525           Dim objCode As VBIDE.CodeModule

               ' Get the project details in the workbook.
5526           Set objProject = xlBook.VBProject

               ' Iterate through each component in the project.
5527           For Each objComponent In objProject.VBComponents

                   ' Delete code modules
5528               Set objCode = objComponent.CodeModule
5529               objCode.DeleteLines 1, objCode.CountOfLines

5530               Set objCode = Nothing
5531               Set objComponent = Nothing
5532           Next

5533           Set objProject = Nothing


5534           xlApp.Visible = True                    'show the sheet

5535           xlApp.VBE.ActiveVBProject.VBComponents.Import ParentDirectoryName & sSaveFileMacroFile
5536           xlApp.Run "AssignButton"
5537       End With

       '    Exit Sub

5538   ErrHandler:
           'User pressed the Cancel button

5539       On Error GoTo notopen
5540       If Not xlApp.ActiveWorkbook Is Nothing Then
5541           ActiveWorkbook.CheckCompatibility = False
5542           xlApp.ActiveWorkbook.Save               'save the workbook
               'xlApp.ActiveWorkbook.Close

5543       End If

5544   notopen:

       '    xlApp.Application.Quit

       '    xlApp.Quit
       '    Set xlApp = Nothing

       '    If CommonDialog1.filename <> "" Then
       '        MsgBox CommonDialog1.filename & " has been written.", vbOKOnly, "File Opened"
       '    End If

5545   On Error GoTo vbwErrHandler

' <VB WATCH>
5546       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5547       Exit Sub
' <VB WATCH>
5548       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5549       Exit Sub
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
5550       On Error GoTo vbwErrHandler
5551       Const VBWPROCNAME = "frmPLCData.GetWorksheetTabs"
5552       If vbwProtector.vbwTraceProc Then
5553           Dim vbwProtectorParameterString As String
5554           If vbwProtector.vbwTraceParameters Then
5555               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("filename", filename) & ", "
5556               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("WorkSheetName", WorkSheetName) & ") "
5557           End If
5558           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5559       End If
' </VB WATCH>

           'see what worksheet tabs alread exist in the excel worksheet

5560       Dim intSheets As Integer    'number of sheets in the workbook
5561       Dim I As Integer
5562       Dim S As String
5563       Dim ans As Integer
5564       Dim NameOK As Boolean

5565       intSheets = xlApp.Worksheets.Count      'how many sheets are there?

           'define a crlf string
5566       S = vbCrLf

5567       For I = 1 To intSheets
5568           S = S & xlApp.Worksheets(I).Name & vbCrLf   'add in the worksheet name
5569       Next I

           'tell the user the names so far and ask if he/she wants to add another
5570       ans = MsgBox("You have the following Worksheet Names in " & filename & ": " & S & "Do you want to add another sheet to this file?", vbYesNo, "Sheets in Excel File")

           'get the answer
5571       If ans = vbNo Then
5572           GetWorksheetTabs = vbNo     'set up flag for when we return to the calling subroutine
' <VB WATCH>
5573       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5574           Exit Function
5575       End If

           'get worksheet name from user and check to see that it's not already used

5576       NameOK = False  'start assuming that the name is bad

5577       While Not NameOK    'as long as it's bad, stay in this loop
5578           WorkSheetName = InputBox("Enter Worksheet Name for this run.")  'ask for name

5579           If WorkSheetName = "" Then      'if we get a nul return or user presses cancel
5580               GetWorksheetTabs = vbNo
' <VB WATCH>
5581       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5582               Exit Function
5583           End If

5584           For I = 1 To xlApp.Worksheets.Count     'go through all of the existing sheets
5585               If WorkSheetName = xlApp.Worksheets(I).Name Then        'if the names are the same
5586                   MsgBox "The name " & WorkSheetName & " already exists for a Worksheet.  Please try again.", vbOKOnly, "Bad Worksheet Name"  'tell the user
5587                   NameOK = False
5588                   Exit For
5589               End If
5590               NameOK = True       'if we make it thru say the name is ok
5591           Next I
5592       Wend

5593       xlApp.Worksheets.Add , xlApp.Worksheets(xlApp.Worksheets.Count)     'add a worksheer
5594       xlApp.Worksheets(xlApp.Worksheets.Count).Name = WorkSheetName       'give it the desired name
5595       GetWorksheetTabs = vbYes                                            'say that the results were ok

' <VB WATCH>
5596       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5597       Exit Function
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
5598       On Error GoTo vbwErrHandler
5599       Const VBWPROCNAME = "frmPLCData.NewWorkBook"
5600       If vbwProtector.vbwTraceProc Then
5601           Dim vbwProtectorParameterString As String
5602           If vbwProtector.vbwTraceParameters Then
5603               vbwProtectorParameterString = "()"
5604           End If
5605           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5606       End If
' </VB WATCH>

5607       Dim WorkSheetName As String

           'we've just added a new workbook, delete sheet1, sheet2, etc
5608       xlApp.DisplayAlerts = False
5609       While xlApp.Worksheets.Count > 1
5610           xlApp.Worksheets(1).Delete          'delete the sheet
5611       Wend
5612       xlApp.DisplayAlerts = True

5613       WorkSheetName = InputBox("Enter Title Worksheet Name for this run.")    'get the desired name
5614       xlApp.Worksheets(1).Name = WorkSheetName    'and name the sheet

5615       NewWorkBook = WorkSheetName

' <VB WATCH>
5616       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5617       Exit Function
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
5618       On Error GoTo vbwErrHandler
5619       Const VBWPROCNAME = "frmPLCData.CalibrateSoftware"
5620       If vbwProtector.vbwTraceProc Then
5621           Dim vbwProtectorParameterString As String
5622           If vbwProtector.vbwTraceParameters Then
5623               vbwProtectorParameterString = "()"
5624           End If
5625           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5626       End If
' </VB WATCH>
5627           frmCalibrate.Show
               'Calibrating = True

' <VB WATCH>
5628       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5629       Exit Sub
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
5630       On Error GoTo vbwErrHandler
5631       Const VBWPROCNAME = "frmPLCData.ParseTEMCModelNo"
5632       If vbwProtector.vbwTraceProc Then
5633           Dim vbwProtectorParameterString As String
5634           If vbwProtector.vbwTraceParameters Then
5635               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
5636               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ltr", ltr) & ") "
5637           End If
5638           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5639       End If
' </VB WATCH>
5640       Dim I As Integer
5641       Dim iStart As Integer
5642       Dim iStop As Integer
5643       Dim strCompare As String

5644       For I = 0 To cmbComboName.ListCount - 1                     'go through the combobox entries
5645           iStart = InStr(1, cmbComboName.List(I), "[")
5646           iStop = InStr(1, cmbComboName.List(I), "]")
5647           strCompare = Mid$(cmbComboName.List(I), iStart + 1, iStop - iStart - 1)
5648           If UCase(strCompare) = UCase(ltr) Then   'see when we find the desired index number
5649               cmbComboName.ListIndex = I                                              'if we do, set the combo box
5650               Exit For                                            'and we're done
5651           End If
       '        cmbComboName.ListIndex = -1                             'else, remove any pointer
5652           cmbComboName.ListIndex = cmbComboName.ListCount - 1                           'else, remove any pointer
5653       Next I

5654       txtModelNo.Text = UCase(txtModelNo.Text)
5655       txtModelNo.SelStart = Len(txtModelNo.Text)
' <VB WATCH>
5656       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5657       Exit Function
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
5658       On Error GoTo vbwErrHandler
5659       Const VBWPROCNAME = "frmPLCData.LoadCombo"
5660       If vbwProtector.vbwTraceProc Then
5661           Dim vbwProtectorParameterString As String
5662           If vbwProtector.vbwTraceParameters Then
5663               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
5664               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sTableName", sTableName) & ") "
5665           End If
5666           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5667       End If
' </VB WATCH>

5668       Dim I As Integer
5669       Dim sItem As String
5670       Dim iID As Integer
5671       Dim bUseDropdown As Boolean
5672       Dim qy As New ADODB.Command
5673       Dim rs As New ADODB.Recordset

       '    rsPumpParameters.CursorLocation = adUseClient
       '    If sTableName = "Model" Then
       '        rsPumpParameters.Sort = "Model"
       '    Else
       '        rsPumpParameters.Sort = vbNullString
       '    End If
       '    rsPumpParameters.Open sTableName, cnPumpData, adOpenStatic, adLockOptimistic, adCmdTableDirect

5674       qy.ActiveConnection = cnPumpData
5675       If sTableName = "DischargeDiameter" Or sTableName = "SuctionDiameter" Then
5676           qy.CommandText = "SELECT * FROM " & sTableName & " ORDER BY Val(Description)"
5677       Else
5678           qy.CommandText = "SELECT * FROM " & sTableName & " ORDER BY Description"
5679       End If
5680       If sTableName = "SupermarketPumpData" Then
5681           qy.CommandText = "SELECT ID,Model AS Description FROM " & sTableName
5682       End If
5683       rs.CursorLocation = adUseClient
5684       rs.CursorType = adOpenStatic

5685       rs.Open qy


5686       On Error GoTo NoField
5687       bUseDropdown = True
           'sItem = rsPumpParameters.Fields("UseInDropdown")
       '    If bUseDropdown Then
       '        rsPumpParameters.Sort = "Description"
       '    End If
5688       rs.MoveFirst                                'goto the top
5689       For I = 0 To rs.RecordCount - 1             'go through the whole recordset
5690           sItem = rs.Fields("Description")        'get the description
5691           iID = rs.Fields(0)                      'get the index number - primary key
5692           If bUseDropdown Then
       '            If rsPumpParameters.Fields("UseInDropdown").value = True Then
5693                   cmbComboName.AddItem sItem, I                                   'add the description to the combo box
       '                cmbComboName.AddItem sItem                                   'add the description to the combo box
5694                   cmbComboName.ItemData(cmbComboName.NewIndex) = iID              'add the key number into the item data
       '            End If
5695           End If
5696           rs.MoveNext                             'get the next record
5697       Next I
5698       rs.Close
5699       cmbComboName.ListIndex = -1
5700   On Error GoTo vbwErrHandler
5701       Set rs = Nothing
5702       Set qy = Nothing
' <VB WATCH>
5703       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5704       Exit Function

5705   NoField:
5706       bUseDropdown = False
5707   On Error GoTo vbwErrHandler
5708       Resume Next

' <VB WATCH>
5709       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5710       Exit Function
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
5711       On Error GoTo vbwErrHandler
5712       Const VBWPROCNAME = "frmPLCData.SetGraphMax"
5713       If vbwProtector.vbwTraceProc Then
5714           Dim vbwProtectorParameterString As String
5715           If vbwProtector.vbwTraceParameters Then
5716               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Plothead", Plothead) & ") "
5717           End If
5718           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5719       End If
' </VB WATCH>

5720       Dim I As Integer
5721       Dim m As Single

5722       m = 0
5723       For I = 0 To UBound(Plothead, 2)
5724           If Plothead(1, I) > m Then
5725               m = Plothead(1, I)
5726           End If
5727       Next I
5728       SetGraphMax = 10 * (Int((m / 10) + 0.5) + 1)
5729       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
5730       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10 * (Int((m / 10) + 0.5) + 1)
5731       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0

' <VB WATCH>
5732       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5733       Exit Function
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
5734       On Error GoTo vbwErrHandler
5735       Const VBWPROCNAME = "frmPLCData.CalculateSpeed"
5736       If vbwProtector.vbwTraceProc Then
5737           Dim vbwProtectorParameterString As String
5738           If vbwProtector.vbwTraceParameters Then
5739               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("CoefSq", CoefSq) & ", "
5740               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("CoefLin", CoefLin) & ", "
5741               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("CoefConstant", CoefConstant) & ", "
5742               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("InputHP", InputHP) & ", "
5743               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("SG", SG) & ") "
5744           End If
5745           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5746       End If
' </VB WATCH>
5747       Dim I As Integer
5748       Dim OldResult As Double
5749       Dim NewResult As Double

5750       CalculateSpeed = 0

5751       If SG > 5 Or SG < 0.01 Then
5752           MsgBox "Bad value for SG...must be between 0.01 and 5.", vbOKOnly, "Bad SG Value"
' <VB WATCH>
5753       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5754           Exit Function
5755       End If

5756       OldResult = 1000
5757       NewResult = 0

5758       I = 1

5759       Do While Abs(NewResult - OldResult) > 0.1
5760           ReDim Preserve results(I)
5761           Select Case I
                   Case 1
5762                   results(I - 1).HP = InputHP
5763               Case 2
5764                   results(I - 1).HP = results(I - 2).HP * SG
5765               Case Else
5766                   results(I - 1).HP = results(I - 2).HP * (results(I - 2).Speed / results(I - 3).Speed) ^ 3
5767           End Select
5768           OldResult = NewResult
5769           results(I - 1).Speed = CalcPoly(CoefSq, CoefLin, CoefConstant, results(I - 1).HP)
5770           NewResult = results(I - 1).Speed
5771           If I > 15 Then
5772               If I = 0 Or I > 15 Then
5773                   MsgBox "Over 15 calculations and no convergence", vbOKOnly, "Too many iterations"
' <VB WATCH>
5774       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5775                   Exit Function
5776               End If
' <VB WATCH>
5777       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5778               Exit Function
5779           End If
5780           I = I + 1
5781       Loop
5782       CalculateSpeed = I - 1
' <VB WATCH>
5783       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5784       Exit Function
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
5785       On Error GoTo vbwErrHandler
5786       Const VBWPROCNAME = "frmPLCData.CalcPoly"
5787       If vbwProtector.vbwTraceProc Then
5788           Dim vbwProtectorParameterString As String
5789           If vbwProtector.vbwTraceParameters Then
5790               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("CoefSq", CoefSq) & ", "
5791               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("CoefLin", CoefLin) & ", "
5792               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("CoefConstant", CoefConstant) & ", "
5793               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("DataIn", DataIn) & ") "
5794           End If
5795           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5796       End If
' </VB WATCH>
5797       CalcPoly = CoefSq * DataIn ^ 2 + CoefLin * DataIn + CoefConstant
' <VB WATCH>
5798       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5799       Exit Function
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
5800       On Error GoTo vbwErrHandler
5801       Const VBWPROCNAME = "frmPLCData.GetBalanceHoleData"
5802       If vbwProtector.vbwTraceProc Then
5803           Dim vbwProtectorParameterString As String
5804           If vbwProtector.vbwTraceParameters Then
5805               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("SerialNumber", SerialNumber) & ", "
5806               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("TestDate", TestDate) & ") "
5807           End If
5808           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5809       End If
' </VB WATCH>
5810       If rsBalanceHoles.State = adStateOpen Then
5811           rsBalanceHoles.Close
5812       End If
5813       qyBalanceHoles.CommandText = "SELECT BalanceHoles.*, " & _
                             "IIf([Diameter]=99, 'Slot', [diameter]) as Diameter1, IIf([BoltCircle]=99, 'Unknown', [BoltCircle]) as BoltCircle1 " & _
                             "FROM BalanceHoles " & _
                             "WHERE [SerialNo] = '" & SerialNumber & "' AND [Date] <= #" & TestDate & "# " & _
                             "ORDER BY [Date], Val([BoltCircle]);"

5814       rsBalanceHoles.Open qyBalanceHoles
5815       rsBalanceHoles.Filter = ""

5816       Set dgBalanceHoles.DataSource = rsBalanceHoles

5817       Dim c As Column
5818       For Each c In dgBalanceHoles.Columns
5819           Select Case c.DataField
               Case "BalanceHoleID"
5820               c.Visible = False
5821           Case "SerialNo"
5822               c.Visible = False
5823           Case "Date"
5824               c.Visible = True
5825               c.Alignment = dbgCenter
5826               c.Width = 2000
5827           Case "Number"
5828               c.Visible = True
5829               c.Alignment = dbgCenter
5830               c.Width = 700
5831           Case "Diameter"
5832               c.Visible = False
5833           Case "Diameter1"
5834               c.Caption = "Diameter"
5835               c.Visible = True
5836               c.Alignment = dbgCenter
5837               c.Width = 700
5838           Case "BoltCircle1"
5839               c.Caption = "Bolt Circle"
5840               c.Visible = True
5841               c.Alignment = dbgCenter
5842               c.Width = 800
5843           Case "BoltCircle"
5844               c.Visible = False
5845           Case Else ' hide all other columns.
5846               c.Visible = False
5847           End Select
5848       Next c

' <VB WATCH>
5849       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5850       Exit Sub
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
5851       On Error GoTo vbwErrHandler
5852       Const VBWPROCNAME = "frmPLCData.FixPointsToPlot"
5853       If vbwProtector.vbwTraceProc Then
5854           Dim vbwProtectorParameterString As String
5855           If vbwProtector.vbwTraceParameters Then
5856               vbwProtectorParameterString = "()"
5857           End If
5858           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5859       End If
' </VB WATCH>
5860       If DataGrid2.Row = -1 Then
' <VB WATCH>
5861       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5862           Exit Sub
5863       End If
5864       Dim PresentGridRow As Integer
5865       PresentGridRow = DataGrid2.Row
5866       Dim GridIndex As Integer
5867       UpDown2.value = 8
5868       If DataGrid2.Row <> -1 Then
5869           For GridIndex = 0 To 7
5870               DataGrid2.Row = GridIndex
5871               If DataGrid2.Columns("Flow") = 0 And DataGrid2.Columns("TDH") = 0 Then
5872                   txtUpDn2.Text = GridIndex
5873                   If GridIndex = 0 Then
5874                       UpDown2.value = 8
5875                   Else
5876                       UpDown2.value = GridIndex
5877                   End If
' <VB WATCH>
5878       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5879                   Exit Sub
5880               End If
5881           Next GridIndex
5882       End If
5883       DataGrid2.Row = PresentGridRow
' <VB WATCH>
5884       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5885       Exit Sub
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
5886       On Error GoTo vbwErrHandler
5887       Const VBWPROCNAME = "frmPLCData.SetFrequencyCombo"
5888       If vbwProtector.vbwTraceProc Then
5889           Dim vbwProtectorParameterString As String
5890           If vbwProtector.vbwTraceParameters Then
5891               vbwProtectorParameterString = "()"
5892           End If
5893           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5894       End If
' </VB WATCH>
5895       Dim j As Integer
5896       For j = 0 To cmbFrequency.ListCount - 1
5897           If cmbFrequency.List(j) = "60 Hz" Then
5898               cmbFrequency.ListIndex = j
5899               Exit For
5900           End If
5901       Next

' <VB WATCH>
5902       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5903       Exit Sub
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
