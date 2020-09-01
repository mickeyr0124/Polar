VERSION 5.00
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "TABCTL32.OCX"
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "comdlg32.ocx"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomct2.ocx"
Begin VB.Form frmPLCData 
   Caption         =   "Polar Rundown"
   ClientHeight    =   11400
   ClientLeft      =   4116
   ClientTop       =   1656
   ClientWidth     =   16632
   LinkTopic       =   "Form1"
   ScaleHeight     =   11400
   ScaleWidth      =   16632
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
      _ExtentX        =   2773
      _ExtentY        =   572
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
         Size            =   7.8
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
      _ExtentX        =   26331
      _ExtentY        =   17992
      _Version        =   393216
      Tabs            =   4
      Tab             =   2
      TabsPerRow      =   4
      TabHeight       =   520
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.4
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
            Size            =   16.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   504
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
            Size            =   16.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   504
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
         _ExtentX        =   445
         _ExtentY        =   910
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
            Size            =   9.6
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
               Size            =   9.6
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
               Size            =   7.8
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
               Size            =   10.8
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
            Size            =   9.6
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
            Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
            Size            =   9.6
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   9.6
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
            Size            =   13.2
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
            _ExtentX        =   8276
            _ExtentY        =   1715
            _Version        =   393216
            AllowUpdate     =   0   'False
            BackColor       =   16777152
            Enabled         =   -1  'True
            ForeColor       =   0
            HeadLines       =   1
            RowHeight       =   15
            BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
               Name            =   "MS Sans Serif"
               Size            =   7.8
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   13.2
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
            Size            =   13.2
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
            Size            =   9.6
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   6.6
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
               Size            =   6.6
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
               Size            =   6.6
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
               Size            =   6.6
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
            Size            =   9.6
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
         _ExtentX        =   10181
         _ExtentY        =   4255
         _Version        =   393216
         Enabled         =   0   'False
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
            Size            =   9.6
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
            Size            =   9.6
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
               Size            =   7.8
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
                  Size            =   8.4
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
                  Size            =   8.4
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
                  Size            =   8.4
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
                  Size            =   8.4
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
                  Size            =   8.4
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
                  Size            =   8.4
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
                  Size            =   8.4
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
                  Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
            Size            =   9.6
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
            Size            =   9.6
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   7.8
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
               Size            =   6.6
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
               Size            =   6.6
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
               Size            =   6.6
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
               Size            =   6.6
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
         _ExtentX        =   10605
         _ExtentY        =   4255
         _Version        =   393216
         AllowUpdate     =   -1  'True
         BackColor       =   16777215
         Enabled         =   -1  'True
         HeadLines       =   1
         RowHeight       =   15
         BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   7.8
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
                  Size            =   8.4
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
                  Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
               Size            =   8.4
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
         _ExtentX        =   445
         _ExtentY        =   910
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
            Size            =   8.4
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
            Size            =   8.4
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
            Size            =   8.4
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
            Size            =   7.8
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
            Size            =   8.4
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
            Size            =   8.4
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
            Size            =   8.4
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
            Size            =   8.4
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
            Size            =   8.4
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
            Size            =   8.4
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
            Size            =   8.4
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
            Size            =   8.4
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
            Size            =   8.4
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
            Size            =   8.4
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
            Size            =   8.4
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
            Size            =   8.4
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
            Size            =   8.4
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
            Size            =   8.4
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
            Size            =   8.4
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
            Size            =   8.4
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
            Size            =   9.6
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
         Size            =   7.8
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
         Size            =   7.8
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
         Size            =   9.6
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
         Size            =   8.4
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
         Size            =   9.6
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

Private Sub cmbVoltage_click()
' <VB WATCH>
220        On Error GoTo vbwErrHandler
221        Const VBWPROCNAME = "frmPLCData.cmbVoltage_click"
222        If vbwProtector.vbwTraceProc Then
223            Dim vbwProtectorParameterString As String
224            If vbwProtector.vbwTraceParameters Then
225                vbwProtectorParameterString = "()"
226            End If
227            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
228        End If
' </VB WATCH>
229        If Me.cmbVoltage.ListIndex = 0 Then
230            Me.cmbFrequency.ListIndex = 2
231        End If
' <VB WATCH>
232        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
233        Exit Sub
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
234        On Error GoTo vbwErrHandler
235        Const VBWPROCNAME = "frmPLCData.cmbMagtrol_Click"
236        If vbwProtector.vbwTraceProc Then
237            Dim vbwProtectorParameterString As String
238            If vbwProtector.vbwTraceParameters Then
239                vbwProtectorParameterString = "()"
240            End If
241            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
242        End If
' </VB WATCH>
243        Dim I As Integer
244        Dim sSendStr As String
245        Dim sGPIBName As String
246        Dim MagtrolName As String

247        I = cmbMagtrol.ItemData(cmbMagtrol.ListIndex)
248        sGPIBName = "GPIB" & I
249        MagtrolName = cmbMagtrol.List(cmbMagtrol.ListIndex)

250        If I = 99 Then      'manual entry
251            boMagtrolOperating = False
252            EnableMagtrolFields
' <VB WATCH>
253        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
254            Exit Sub
255        Else
256            boMagtrolOperating = True
257        End If

258        SetupMagtrols MagtrolName, I

' <VB WATCH>
259        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
260        Exit Sub
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
261        On Error GoTo vbwErrHandler
262        Const VBWPROCNAME = "frmPLCData.cmbPLCLoop_Click"
263        If vbwProtector.vbwTraceProc Then
264            Dim vbwProtectorParameterString As String
265            If vbwProtector.vbwTraceParameters Then
266                vbwProtectorParameterString = "()"
267            End If
268            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
269        End If
' </VB WATCH>

270        Dim RetVal As String

           'manual data entry selection
271        If cmbPLCLoop.ListIndex = cmbPLCLoop.ListCount - 1 Then 'no plc
272            boPLCOperating = False
273            EnablePLCFields
274            If DeviceOpen = True Then
275                RetVal = DisconnectPLC()
276            End If
' <VB WATCH>
277        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
278            Exit Sub
279        End If

280        If DeviceOpen = True Then
281            RetVal = DisconnectPLC()
282        End If

283        RetVal = ConnectToPLC(cmbPLCLoop.ItemData(cmbPLCLoop.ListIndex))
284        If RetVal <> 0 Then
285            MsgBox ("Can't connect to PLC - " & Description(cmbPLCLoop.ListIndex))
286            boPLCOperating = False
287            EnablePLCFields
288        Else
289            boPLCOperating = True
290            tDevice = cmbPLCLoop.ItemData(cmbPLCLoop.ListIndex)
291            DisablePLCFields
292        End If
' <VB WATCH>
293        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
294        Exit Sub
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
295        On Error GoTo vbwErrHandler
296        Const VBWPROCNAME = "frmPLCData.cmbTestDate_Click"
297        If vbwProtector.vbwTraceProc Then
298            Dim vbwProtectorParameterString As String
299            If vbwProtector.vbwTraceParameters Then
300                vbwProtectorParameterString = "()"
301            End If
302            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
303        End If
' </VB WATCH>

304        Dim sName As String
305        Dim sParam As String
306        Dim I As Integer
307        Dim j As Integer
308        Dim k As Integer
309        Dim bSk As Boolean
310        Dim sBC As Single
311        Dim NOK() As Long

312        cmdModifyBalanceHoleData.Visible = False


313        If Not boFoundTestSetup Then    'if we don't have any TestSetup data written
314            boFoundTestData = False
' <VB WATCH>
315        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
316            Exit Sub
317        End If


           'select the testsetup data for the serial number
318        qyTestSetup.ActiveConnection = cnPumpData
319        qyTestSetup.CommandText = "SELECT * " & _
                         "From TempTestSetupData " & _
                         "Where (((TempTestSetupData.SerialNumber) = '" & txtSN.Text & "') AND TempTestSetupData.Date = #" & cmbTestDate.List(cmbTestDate.ListIndex) & "#) " & _
                         "ORDER BY TempTestSetupData.Date;"

320        If rsTestSetup.State = adStateOpen Then
321            rsTestSetup.Close
322        End If

323        With rsTestSetup     'open the recordset for the query
       '        .Index = "FindData"
324            .CursorLocation = adUseClient
325            .CursorType = adOpenStatic
326            .Open qyTestSetup
327        End With

           'move to the selected date
328        If Not rsTestSetup.BOF Then
329            rsTestSetup.MoveFirst
330        End If
       '
           'show the correct combo box entries for this record
           'SetComboTestSetup cmbOrificeNumber, "OrificeNumber", "OrificeNumber", rsTestSetup
331        SetComboTestSetup cmbTestSpec, "TestSpec", "TestSpecification", rsTestSetup
332        SetComboTestSetup cmbLoopNumber, "LoopNumber", "LoopNumber", rsTestSetup
333        SetComboTestSetup cmbSuctDia, "SuctDiam", "SuctionDiameter", rsTestSetup
334        SetComboTestSetup cmbDischDia, "DischDiam", "DischargeDiameter", rsTestSetup
335        SetComboTestSetup cmbTachID, "TachID", "TachID", rsTestSetup
336        SetComboTestSetup cmbAnalyzerNo, "AnalyzerNo", "AnalyzerNo", rsTestSetup
337        SetComboTestSetup cmbVoltage, "Voltage", "Voltage", rsTestSetup
338        SetComboTestSetup cmbFrequency, "Frequency", "Frequency", rsTestSetup
339        SetComboTestSetup cmbMounting, "Mounting", "Mounting", rsTestSetup
340        SetComboTestSetup cmbPLCNo, "PLCNo", "PLCNo", rsTestSetup
341        SetComboTestSetup cmbFlowMeter, "FlowMeterID", "PumpFlowMeter", rsTestSetup
342        SetComboTestSetup cmbSuctionPressureTransducer, "SuctionID", "SuctionPressureTransducer", rsTestSetup
343        SetComboTestSetup cmbDischargePressureTransducer, "DischID", "DischargePressureTransducer", rsTestSetup
344        SetComboTestSetup cmbTemperatureTransducer, "TemperatureID", "TemperatureTransducer", rsTestSetup
345        SetComboTestSetup cmbCirculationFlowMeter, "MagFlowID", "CirculationFlowMeter", rsTestSetup

346        sName = "HDCor"
347        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
348            sParam = rsTestSetup.Fields(sName)
349        Else
350            sParam = vbNullString
351        End If
352        txtHDCor.Text = sParam

353        sName = "KWMult"
354        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
355            sParam = rsTestSetup.Fields(sName)
356        Else
357            sParam = vbNullString
358        End If
359        txtKWMult.Text = sParam

360        sName = "Who"
361        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
362            sParam = rsTestSetup.Fields(sName)
363        Else
364            sParam = vbNullString
365        End If
366        txtWho.Text = sParam

367        sName = "RMA"
368        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
369            sParam = rsTestSetup.Fields(sName)
370        Else
371            sParam = vbNullString
372        End If
373        txtRMA.Text = sParam

374        sName = "Remarks"
375        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
376            sParam = rsTestSetup.Fields(sName)
377        Else
378            sParam = vbNullString
379        End If
380        txtTestSetupRemarks.Text = sParam

381        sName = "VFDFrequency"
382        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
383            sParam = rsTestSetup.Fields(sName)
384        Else
385            sParam = vbNullString
386        End If
387        txtVFDFreq.Text = sParam

388        sName = "SuctionGageHeight"
389        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
390            sParam = rsTestSetup.Fields(sName)
391        Else
392            sParam = 0
393        End If
394        txtSuctHeight.Text = sParam

395        sName = "DischargeGageHeight"
396        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
397            sParam = rsTestSetup.Fields(sName)
398        Else
399            sParam = 0
400        End If
401        txtDischHeight.Text = sParam

402        sName = "EndPlay"
403        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
404            sParam = rsTestSetup.Fields(sName)
405        Else
406            sParam = vbNullString
407        End If
408        txtEndPlay.Text = sParam

409        sName = "GGAP"
410        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
411            sParam = rsTestSetup.Fields(sName)
412        Else
413            sParam = vbNullString
414        End If
415        txtGGap.Text = sParam

416        sName = "OtherMods"
417        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
418            sParam = rsTestSetup.Fields(sName)
419        Else
420            sParam = vbNullString
421        End If
422        txtOtherMods.Text = sParam

423        If rsTestSetup.Fields("ImpFeathered") Then
424            chkFeathered.value = 1
425        Else
426            chkFeathered.value = 0
427        End If

428        If Val(rsTestSetup.Fields("ImpTrimmed")) = 0 Then
429            chkTrimmed.value = 0
430            txtImpTrim.Visible = False
431            txtImpTrim.Text = rsTestSetup.Fields("Imptrimmed")
432        Else
433            chkTrimmed.value = 1
434            txtImpTrim.Visible = True
435            txtImpTrim.Text = rsTestSetup.Fields("Imptrimmed")
436        End If

437        If Val(rsTestSetup.Fields("PumpDischOrifice")) = 0 Then
438            chkOrifice.value = 0
439            txtOrifice.Visible = False
440        Else
441            chkOrifice.value = 1
442            txtOrifice.Visible = True
443            txtOrifice.Text = rsTestSetup.Fields("PumpDischOrifice")
444        End If

445        If Val(rsTestSetup.Fields("CircFlowOrifice")) = 0 Then
446            chkCircOrifice.value = 0
447            txtCircOrifice.Visible = False
448        Else
449            chkCircOrifice.value = 1
450            txtCircOrifice.Visible = True
451            txtCircOrifice.Text = rsTestSetup.Fields("CircFlowOrifice")
452        End If

453        If (IsNull(rsTestSetup.Fields("NPSHFile"))) Or (LenB(rsTestSetup.Fields("NPSHFile")) = 0) Then
454            chkNPSH.value = 0
455            txtNPSHFile.Visible = False
456        Else
457            chkNPSH.value = 1
458            txtNPSHFile.Visible = True
459            txtNPSHFile.Text = rsTestSetup.Fields("NPSHFile")
460        End If

461        If (IsNull(rsTestSetup.Fields("PictureFile"))) Or (LenB(rsTestSetup.Fields("PictureFile")) = 0) Then
462            chkPictures.value = 0
463            txtPicturesFile.Visible = False
464        Else
465            chkPictures.value = 1
466            txtPicturesFile.Visible = True
467            txtPicturesFile.Text = rsTestSetup.Fields("PictureFile")
468        End If

469        If (IsNull(rsTestSetup.Fields("VibrationFile"))) Or (LenB(rsTestSetup.Fields("VibrationFile")) = 0) Then
470            chkVibration.value = 0
471            txtVibrationFile.Visible = False
472        Else
473            chkVibration.value = 1
474            txtVibrationFile.Visible = True
475            txtVibrationFile.Text = rsTestSetup.Fields("VibrationFile")
476        End If


           'for TEMC Inspection Report
477        sName = "InsulationMeggerVolts"
478        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
479            sParam = rsTestSetup.Fields(sName)
480        Else
481            sParam = 0
482        End If
483        txtTestAndInspection(0).Text = sParam

484        sName = "InsulationMegOhms"
485        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
486            sParam = rsTestSetup.Fields(sName)
487        Else
488            sParam = 0
489        End If
490        txtTestAndInspection(1).Text = sParam

491        sName = "DielectricVolts"
492        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
493            sParam = rsTestSetup.Fields(sName)
494        Else
495            sParam = 0
496        End If
497        txtTestAndInspection(2).Text = sParam

498        sName = "DielectricTime"
499        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
500            sParam = rsTestSetup.Fields(sName)
501        Else
502            sParam = 0
503        End If
504        txtTestAndInspection(3).Text = sParam

505        sName = "HydrostaticValue"
506        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
507            sParam = rsTestSetup.Fields(sName)
508        Else
509            sParam = 0
510        End If
511        txtTestAndInspection(4).Text = sParam

512        sName = "HydrostaticTime"
513        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
514            sParam = rsTestSetup.Fields(sName)
515        Else
516            sParam = 0
517        End If
518        txtTestAndInspection(5).Text = sParam

519        sName = "PneumaticValue"
520        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
521            sParam = rsTestSetup.Fields(sName)
522        Else
523            sParam = 0
524        End If
525        txtTestAndInspection(6).Text = sParam

526        sName = "PneumaticTime"
527        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
528            sParam = rsTestSetup.Fields(sName)
529        Else
530            sParam = 0
531        End If
532        txtTestAndInspection(7).Text = sParam

533        For I = 0 To cmbTestAndInspection(0).ListCount - 1
534            If cmbTestAndInspection(0).Text = rsTestSetup.Fields("HydrostaticUnits") Then
535                    cmbTestAndInspection(0).ListIndex = I
536                    Exit For
537            End If
538            cmbTestAndInspection(0).ListIndex = -1
539        Next I


540        For I = 0 To cmbTestAndInspection(1).ListCount - 1
541            If cmbTestAndInspection(1).Text = rsTestSetup.Fields("PneumaticUnits") Then
542                    cmbTestAndInspection(1).ListIndex = I
543                    Exit For
544            End If
545            cmbTestAndInspection(1).ListIndex = -1
546        Next I

547        TestAndInspectionGood(0).value = Abs(rsTestSetup!insulationgood)
548        TestAndInspectionGood(1).value = Abs(rsTestSetup!DielectricGood)
549        TestAndInspectionGood(2).value = Abs(rsTestSetup!HydrostaticGood)
550        TestAndInspectionGood(3).value = Abs(rsTestSetup!PneumaticGood)
551        TestAndInspectionGood(4).value = Abs(rsTestSetup!GeneralAppearanceGood)
552        TestAndInspectionGood(5).value = Abs(rsTestSetup!OutlineDimensionsGood)
553        TestAndInspectionGood(6).value = Abs(rsTestSetup!MotorNoLoadTestGood)
554        TestAndInspectionGood(7).value = Abs(rsTestSetup!MotorLockedRotorTestGood)
555        TestAndInspectionGood(8).value = Abs(rsTestSetup!HydrostaticTestGood)
556        TestAndInspectionGood(9).value = Abs(rsTestSetup!HydraulicTestGood)
557        TestAndInspectionGood(10).value = Abs(rsTestSetup!NPSHTestGood)
558        TestAndInspectionGood(11).value = Abs(rsTestSetup!CleanPurgeSealGood)
559        TestAndInspectionGood(12).value = Abs(rsTestSetup!PaintCheckGood)
560        TestAndInspectionGood(13).value = Abs(rsTestSetup!NameplateGood)
561        TestAndInspectionGood(14).value = Abs(rsTestSetup!SupervisorApproval)

562        GetBalanceHoleData frmPLCData.txtSN.Text, cmbTestDate.Text

563         If rsBalanceHoles.RecordCount = 0 Then
564            chkBalanceHoles.value = 0
565            dgBalanceHoles.Visible = False
566            boGotBalanceHoles = False
567        Else
568            boGotBalanceHoles = True
569            ReDim NOK(rsBalanceHoles.RecordCount)
570            rsBalanceHoles.MoveLast
571            For I = 1 To rsBalanceHoles.RecordCount
572                NOK(I) = 0
573            Next I

574            For j = 1 To rsBalanceHoles.RecordCount - 1
575                rsBalanceHoles.MoveFirst
576                rsBalanceHoles.Move rsBalanceHoles.RecordCount - j
577                sBC = rsBalanceHoles.Fields("BoltCircle")
578                bSk = False
579                For k = 1 To rsBalanceHoles.RecordCount
580                    If NOK(k) = rsBalanceHoles.Fields(0) Then
581                        bSk = True
582                    End If
583                Next k
584                If Not bSk Then
585                    For I = rsBalanceHoles.RecordCount - j To 1 Step -1
586                        rsBalanceHoles.MovePrevious
587                        If rsBalanceHoles.Fields("BoltCircle") = sBC Then
588                            NOK(I) = rsBalanceHoles.Fields(0)
589                        End If
590                    Next I
591                End If
592            Next j

593            Dim sFilt As String
594            sFilt = ""
595            For I = 1 To rsBalanceHoles.RecordCount
596                If NOK(I) <> 0 Then
597                    sFilt = sFilt & "(BalanceHoleID <> " & NOK(I) & ") AND "
       '                sFilt = sFilt & "(" & rsBalanceHoles.Filter & " AND BalanceHoleID <> " & NOK(I) & ") AND "
598                End If
599            Next I

600            If Len(sFilt) > 4 Then
601                sFilt = Left(sFilt, Len(sFilt) - 4)
602                rsBalanceHoles.Filter = sFilt
603            End If

604            chkBalanceHoles.value = 1
605            dgBalanceHoles.Visible = True
606        End If
       '
           'set the test date filter for the test data
607        rsTestData.Filter = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"

608        If rsTestData.RecordCount = 0 Then
609            boFoundTestData = False
610            AddTestData
611            EnableTestDataControls
612            MsgBox "No Test Data Exists for this Serial Number"
613        Else
614            boFoundTestData = True
615            DisableTestDataControls                         'if it's in the real database, don't allow changes here
616        End If

617        If Not boTestDateIsApproved Then    'data approved?
618            EnableTestDataControls
619        End If

620        If rsTestSetup.Fields("Approved") = True Then
621            DisableTestDataControls                         'if it's in the real database, don't allow changes here
622            lblTestDateApproved.Visible = True
623            MsgBox ("Found pump.  Data cannot be modified.")
624            If boCanApprove Then
625                cmdApproveTestDate.Caption = "Unapprove this Test Date"
626            End If
627        Else
628            EnableTestDataControls                          'it's in the temp database, allow changes
629            lblTestDateApproved.Visible = False
630            If boPumpIsApproved = True Then
631                MsgBox ("Found pump.  Pump data cannot be modified, but test setup data and test data can be modified.")
632            Else
633                MsgBox ("Found pump.  Pump data, test setup data, and test data can be modified.")
634            End If
635            If boCanApprove Then
636                If rsPumpData.Fields("Approved") = True Then
637                    cmdApproveTestDate.Enabled = True
638                    cmdApproveTestDate.Caption = "Approve this Test Date"
639                Else
640                    cmdApproveTestDate.Caption = "You Must Approve Pump First"
641                    cmdApproveTestDate.Enabled = False
642                End If
643            End If
644        End If

645        rsEff.MoveFirst
646        rsTestData.MoveFirst

647        For I = 1 To rsTestData.RecordCount
648            DoEfficiencyCalcs
649            rsEff.MoveNext
650            rsTestData.MoveNext
651        Next I

          ' fix the datagrid
652       Set DataGrid1.DataSource = rsTestData
653       Set DataGrid2.DataSource = rsEff

654       Dim c As Column
655       For Each c In DataGrid1.Columns
656          Select Case c.DataField
             Case "TestDataID"     'Hide some columns
657             c.Visible = False
658          Case "SerialNumber"
659             c.Visible = False
660          Case "Date"
661             c.Visible = False
662          Case Else             ' Show all other columns.
663             c.Visible = True
664             c.Alignment = dbgRight
665          End Select
666        Next c

667        For Each c In DataGrid2.Columns
668            c.Alignment = dbgCenter
669            c.Width = 750
670            Select Case c.ColIndex
                   Case 1
671                    c.Caption = "Flow"
672                    c.NumberFormat = "###0.00"
673                Case 2
674                    c.Caption = "TDH"
675                    c.NumberFormat = "##0.00"
676                Case 3
677                    c.Caption = "Input Pwr"
678                    c.NumberFormat = "##0.00"
679                    c.Width = 850
680                Case 4
681                    c.Caption = "Voltage"
682                    c.NumberFormat = "##0.00"
683                Case 5
684                    c.Caption = "Current"
685                    c.NumberFormat = "##0.00"
686                Case 6
687                    c.Caption = "Overall Eff"
688                    c.NumberFormat = "##0.00"
689                    c.Width = 850
690                Case 7
691                    c.Caption = "NPSHr"
692                    c.NumberFormat = "#0.00"
693                Case Else
694                    c.Visible = False
695            End Select
696        Next c
697            FixPointsToPlot

698        txtUpDn1.Text = 1

       'unlock the text boxes
699        For I = 0 To 7
700            txtTitle(I).Locked = False
701        Next I

702        For I = 20 To 27
703            txtTitle(I).Locked = False
704        Next I

       'look for titles for TCs and AIs
705        Dim qy As New ADODB.Command
706        Dim rs As New ADODB.Recordset

707        qy.ActiveConnection = cnPumpData

           'see if we have an entry in the table
708        qy.CommandText = "SELECT * FROM AITitles " & _
                             "WHERE (((AITitles.SerialNo)= '" & txtSN.Text & "') " & _
                             "AND ((AITitles.Date)= #" & cmbTestDate.Text & "#)); "

709        With rs     'open the recordset for the query
710            .CursorLocation = adUseClient
711            .CursorType = adOpenStatic
712            .LockType = adLockOptimistic
713            .Open qy
714        End With

715        If Not (rs.BOF = True And rs.EOF = True) Then   'update titles
716            rs.MoveFirst
717            Do While Not rs.EOF
718                txtTitle(rs.Fields("Channel")).Text = rs.Fields("Title")
719                rs.MoveNext
720            Loop
721        End If

722        rs.Close
723        Set rs = Nothing
724        Set qy = Nothing
' <VB WATCH>
725        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
726        Exit Sub
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
727        On Error GoTo vbwErrHandler
728        Const VBWPROCNAME = "frmPLCData.cmdAddNewBalanceHoles_Click"
729        If vbwProtector.vbwTraceProc Then
730            Dim vbwProtectorParameterString As String
731            If vbwProtector.vbwTraceParameters Then
732                vbwProtectorParameterString = "()"
733            End If
734            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
735        End If
' </VB WATCH>
736        Dim strInput As String
737        Dim I As Integer
738        Dim sNumber As Integer
739        Dim sDia As Single
740        Dim sBC As Single

           'get the data for the balance holes
741        strInput = InputBox("Enter Number of Holes")
742        If strInput <> "" Then
743            sNumber = CInt(strInput)
744        Else
745            GoTo CancelPressed
746        End If

747        strInput = InputBox("Enter Decimal Value of Hole Diameter or Slot (For Example, 0.675) ")
748        If strInput <> "" Then
749            If UCase(strInput) = "SLOT" Then
750                strInput = 99
751            End If
752            sDia = CSng(strInput)
753        Else
754            GoTo CancelPressed
755        End If

756        strInput = InputBox("Enter Decimal Value of Bolt Circle or Unknown (For Example, 4.525)")
757        If strInput <> "" Then
758            If UCase(strInput) = "UNKNOWN" Then
759                strInput = 99
760            End If
761            sBC = CSng(strInput)
762        Else
763            GoTo CancelPressed
764        End If

765        GetBalanceHoleData frmPLCData.txtSN.Text, cmbTestDate.Text

766        rsBalanceHoles.AddNew
767        rsBalanceHoles!SerialNo = txtSN.Text
768        rsBalanceHoles!Date = cmbTestDate.Text
769        rsBalanceHoles!Number = sNumber
770        rsBalanceHoles!diameter = sDia
771        rsBalanceHoles!boltcircle = sBC

772        rsBalanceHoles.Update

773        GetBalanceHoleData txtSN.Text, cmbTestDate.Text
774        rsBalanceHoles.MoveLast
775        dgBalanceHoles.Refresh
776        chkBalanceHoles.value = 1

' <VB WATCH>
777        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
778        Exit Sub

779    CancelPressed:
780        MsgBox "No New Balance Hole Data Entered", vbOKOnly
' <VB WATCH>
781        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
782        Exit Sub
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
783        On Error GoTo vbwErrHandler
784        Const VBWPROCNAME = "frmPLCData.cmdAddNewTestDate_Click"
785        If vbwProtector.vbwTraceProc Then
786            Dim vbwProtectorParameterString As String
787            If vbwProtector.vbwTraceParameters Then
788                vbwProtectorParameterString = "()"
789            End If
790            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
791        End If
' </VB WATCH>
792        Dim I As Integer

793        chkFeathered.value = chkSuperMarketFeathered.value

794        For I = 1 To cmbTestDate.ListCount      'see if we already have today's date entered
795            If cmbTestDate.List(I) = Date Then
796                MsgBox "There is already an entry for today.  You can only have one entry for each Serial Number and a given date.  You may want to modify the Serial Number.", vbOKOnly
' <VB WATCH>
797        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
798                Exit Sub
799            End If
800        Next I

           'we didn't find today's date entered, allow data entry
801        boFoundTestSetup = False

802        EnableTestSetupDataControls
803        Pressed = False
804        cmdEnterTestSetupData_Click
805        cmdAddNewBalanceHoles.Visible = True
806        txtWho.Text = LogInInitials
807        MsgBox "New Test Date Added - " & cmbTestDate.List(cmbTestDate.ListCount - 1), vbOKOnly, "Added New Test Date"
' <VB WATCH>
808        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
809        Exit Sub
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
810        On Error GoTo vbwErrHandler
811        Const VBWPROCNAME = "frmPLCData.cmdApprovePump_Click"
812        If vbwProtector.vbwTraceProc Then
813            Dim vbwProtectorParameterString As String
814            If vbwProtector.vbwTraceParameters Then
815                vbwProtectorParameterString = "()"
816            End If
817            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
818        End If
' </VB WATCH>
819        rsPumpData.Fields("Approved") = Not rsPumpData.Fields("Approved")
820        rsPumpData.Update
821        rsPumpData.Requery
822        lblPumpApproved.Visible = rsPumpData.Fields("Approved")
823        If rsPumpData.Fields("Approved") = True Then
824            cmdApprovePump.Caption = "Unapprove This Pump"
825            cmdApproveTestDate.Enabled = True
826            If rsTestSetup.Fields("Approved") = True Then
827                cmdApproveTestDate.Caption = "Unapprove This Test Date"
828            Else
829                cmdApproveTestDate.Caption = "Approve This Test Date"
830            End If
831        Else
832            cmdApprovePump.Caption = "Approve This Pump"
833            cmdApproveTestDate.Caption = "You Must Approve Pump First"
834            cmdApproveTestDate.Enabled = False
835        End If
' <VB WATCH>
836        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
837        Exit Sub
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
838        On Error GoTo vbwErrHandler
839        Const VBWPROCNAME = "frmPLCData.cmdApproveTestDate_Click"
840        If vbwProtector.vbwTraceProc Then
841            Dim vbwProtectorParameterString As String
842            If vbwProtector.vbwTraceParameters Then
843                vbwProtectorParameterString = "()"
844            End If
845            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
846        End If
' </VB WATCH>
847        rsTestSetup.Fields("Approved") = Not rsTestSetup.Fields("Approved")
848        rsTestSetup.Update
849        rsTestSetup.Requery
850        lblTestDateApproved.Visible = rsTestSetup.Fields("Approved")
851        If rsTestSetup.Fields("Approved") = True Then
852            cmdApproveTestDate.Caption = "Unapprove This Test Date"
853        Else
854            cmdApproveTestDate.Caption = "Approve This Test Date"
855        End If
' <VB WATCH>
856        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
857        Exit Sub
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
858        On Error GoTo vbwErrHandler
859        Const VBWPROCNAME = "frmPLCData.cmdCalibrate_Click"
860        If vbwProtector.vbwTraceProc Then
861            Dim vbwProtectorParameterString As String
862            If vbwProtector.vbwTraceParameters Then
863                vbwProtectorParameterString = "()"
864            End If
865            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
866        End If
' </VB WATCH>
867        Dim ans As Integer
868        Dim I As Integer

869        ans = MsgBox("You have selected to calibrate the software.  Do you want to continue?", vbYesNo, "Calibrate Software")
870        If ans = vbNo Then
871            Calibrating = False
' <VB WATCH>
872        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
873            Exit Sub
874        Else
875            CalibrateSoftware
876        End If
' <VB WATCH>
877        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
878        Exit Sub
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
879        On Error GoTo vbwErrHandler
880        Const VBWPROCNAME = "frmPLCData.cmdClearPumpData_Click"
881        If vbwProtector.vbwTraceProc Then
882            Dim vbwProtectorParameterString As String
883            If vbwProtector.vbwTraceParameters Then
884                vbwProtectorParameterString = "()"
885            End If
886            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
887        End If
' </VB WATCH>
888        BlankData
' <VB WATCH>
889        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
890        Exit Sub
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
891        On Error GoTo vbwErrHandler
892        Const VBWPROCNAME = "frmPLCData.cmdDeletePump_Click"
893        If vbwProtector.vbwTraceProc Then
894            Dim vbwProtectorParameterString As String
895            If vbwProtector.vbwTraceParameters Then
896                vbwProtectorParameterString = "()"
897            End If
898            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
899        End If
' </VB WATCH>
900        Dim Answer As Integer
901        Answer = MsgBox("You are about to delete the following record: S/N = " & rsPumpData.Fields("SerialNumber") & "!  Do you want to continue?", vbCritical Or vbYesNo, "Ready to Delete")
902        If Answer = vbYes Then
903            rsPumpData.Delete
904            rsPumpData.Update
905            cmdFindPump_Click
906        End If
' <VB WATCH>
907        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
908        Exit Sub
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
909        On Error GoTo vbwErrHandler
910        Const VBWPROCNAME = "frmPLCData.cmdDeleteTestDate_Click"
911        If vbwProtector.vbwTraceProc Then
912            Dim vbwProtectorParameterString As String
913            If vbwProtector.vbwTraceParameters Then
914                vbwProtectorParameterString = "()"
915            End If
916            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
917        End If
' </VB WATCH>
918        Dim Answer As Integer
919        Answer = MsgBox("You are about to delete the following record: S/N = " & rsTestData.Fields("SerialNumber") & " and Test Date = " & rsTestSetup.Fields("Date") & "!  Do you want to continue?", vbCritical Or vbYesNo, "Ready to Delete")
920        If Answer = vbYes Then
921            rsTestSetup.Delete
922            rsTestSetup.Update
923            cmdFindPump_Click
924        End If
' <VB WATCH>
925        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
926        Exit Sub
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
927        On Error GoTo vbwErrHandler
928        Const VBWPROCNAME = "frmPLCData.cmdEnterPumpData_Click"
929        If vbwProtector.vbwTraceProc Then
930            Dim vbwProtectorParameterString As String
931            If vbwProtector.vbwTraceParameters Then
932                vbwProtectorParameterString = "()"
933            End If
934            vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
935        End If
' </VB WATCH>
936        Dim d As Integer
937        Dim sSearch As String
938        Dim ans As Integer
939        Dim boWriteDataWritten As Boolean


           'check for a serial number
940        If LenB(txtSN.Text) = 0 Then
941            MsgBox "You must have a Serial Number to enter data.  Data has not been saved."
' <VB WATCH>
942        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
943            Exit Sub
944        End If

           'check to make sure most entries are filled in
945        If LenB(txtModelNo.Text) = 0 And optMfr(0).value = True Then
946            MsgBox "You need to enter a MODEL NO before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
947        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
948            Exit Sub
949        End If
950        If LenB(txtSalesOrderNumber.Text) = 0 Then
951            If InStr(1, txtSN.Text, "-") <> 0 Then
952                txtSalesOrderNumber.Text = Mid$(txtSN.Text, 1, InStr(1, txtSN.Text, "-") - 1)
953            End If
954        End If
955        If LenB(txtSalesOrderNumber.Text) = 0 Then
956            MsgBox "You need to enter a SALES ORDER NUMBER before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
957        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
958            Exit Sub
959        End If

960        If cmbMotor.ListIndex = -1 And optMfr(0).value = True Then
961            MsgBox "You need to pick a MOTOR before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
962        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
963            Exit Sub
964        End If

965        If cmbStatorFill.ListIndex = -1 And optMfr(0).value = True Then    'set default
966            cmbStatorFill.ListIndex = 0
967        End If

968        If cmbModel.ListIndex = -1 And optMfr(0).value = True Then
969            MsgBox "You need to pick a MODEL before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
970        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
971            Exit Sub
972        End If

973        If cmbModelGroup.ListIndex = -1 And optMfr(0).value = True Then
974            MsgBox "You need to pick a MODEL GROUP before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
975        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
976            Exit Sub
977        End If


978        If cmbDesignPressure.ListIndex = -1 And optMfr(0).value = True Then
979            MsgBox "You need to pick a DESIGN PRESSURE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
980        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
981            Exit Sub
982        End If

983        If cmbCirculationPath.ListIndex = -1 And optMfr(0).value = True Then
984            MsgBox "You need to pick a CIRCULATION PATH before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
985        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
986            Exit Sub
987        End If

988        If cmbRPM.ListIndex = -1 And optMfr(0).value = True Then
989            MsgBox "You need to pick an RPM before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
990        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
991            Exit Sub
992        End If

       'check TEMC dropdowns

993        If cmbTEMCAdapter.ListIndex = -1 And optMfr(0).value = False Then
994            MsgBox "You need to pick a TEMC ADAPTER before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
995        If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
996            Exit Sub
997        End If

998        If cmbTEMCAdditions.ListIndex = -1 And optMfr(0).value = False Then
999            MsgBox "You need to pick TEMC ADDITIONS before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1000       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1001           Exit Sub
1002       End If

1003       If cmbTEMCCirculation.ListIndex = -1 And optMfr(0).value = False Then
1004           MsgBox "You need to pick a TEMC CIRCULATION before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1005       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1006           Exit Sub
1007       End If

1008       If cmbTEMCDesignPressure.ListIndex = -1 And optMfr(0).value = False Then
1009           MsgBox "You need to pick a TEMC DESIGN PRESSURE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1010       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1011           Exit Sub
1012       End If

1013       If cmbTEMCDivisionType.ListIndex = -1 And optMfr(0).value = False Then
1014           MsgBox "You need to pick a TEMC DIVISION TYPE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1015       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1016           Exit Sub
1017       End If

1018       If cmbTEMCImpellerType.ListIndex = -1 And optMfr(0).value = False Then
1019           MsgBox "You need to pick a TEMC IMPELLER TYPE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1020       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1021           Exit Sub
1022       End If

1023       If cmbTEMCInsulation.ListIndex = -1 And optMfr(0).value = False Then
1024           MsgBox "You need to pick a TEMC INSULATION TYPE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1025       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1026           Exit Sub
1027       End If

1028       If cmbTEMCJacketGasket.ListIndex = -1 And optMfr(0).value = False Then
1029           MsgBox "You need to pick a TEMC JACKET GASKET before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1030       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1031           Exit Sub
1032       End If

1033       If cmbTEMCMaterials.ListIndex = -1 And optMfr(0).value = False Then
1034           MsgBox "You need to pick TEMC MATERIALS before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1035       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1036           Exit Sub
1037       End If

1038       If cmbTEMCModel.ListIndex = -1 And optMfr(0).value = False Then
1039           MsgBox "You need to pick a TEMC MODEL before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1040       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1041           Exit Sub
1042       End If

1043       If cmbTEMCNominalImpSize.ListIndex = -1 And optMfr(0).value = False Then
1044           MsgBox "You need to pick a TEMC NOMINAL IMPELLER SIZE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1045       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1046           Exit Sub
1047       End If

1048       If cmbTEMCNominalDischargeSize.ListIndex = -1 And optMfr(0).value = False Then
1049           MsgBox "You need to pick a TEMC NOMINAL DISCHARGE SIZE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1050       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1051           Exit Sub
1052       End If

1053       If cmbTEMCNominalSuctionSize.ListIndex = -1 And optMfr(0).value = False Then
1054           MsgBox "You need to pick a TEMC NOMINAL SUCTION SIZE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1055       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1056           Exit Sub
1057       End If

1058       If cmbTEMCOtherMotor.ListIndex = -1 And optMfr(0).value = False Then
1059           MsgBox "You need to pick a TEMC OTHER MOTOR before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1060       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1061           Exit Sub
1062       End If

1063       If cmbTEMCPumpStages.ListIndex = -1 And optMfr(0).value = False Then
1064           MsgBox "You need to pick TEMC NUMBER OF PUMP STAGES before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1065       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1066           Exit Sub
1067       End If

1068       If cmbTEMCTRG.ListIndex = -1 And optMfr(0).value = False Then
1069           MsgBox "You need to pick a TEMC TRG before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1070       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1071           Exit Sub
1072       End If

1073       If cmbTEMCVoltage.ListIndex = -1 And optMfr(0).value = False Then
1074           MsgBox "You need to pick a TEMC VOLTAGE before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1075       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1076           Exit Sub
1077       End If

1078       If LenB(txtTEMCFrameNumber.Text) = 0 And optMfr(0).value = False Then
1079           MsgBox "You need to enter a TEMC FRAME NUMBER before saving data.  Data has not been saved.", vbOKOnly
' <VB WATCH>
1080       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1081           Exit Sub
1082       End If


1083       If Not boFoundPump Then     'if we havent found a pump in the database, add it
1084           rsPumpData.AddNew
1085           boWriteDataWritten = False
1086       Else    'else, find the entry
1087           sSearch = "Serialnumber = '" & frmPLCData.txtSN.Text & "'"
1088           rsPumpData.MoveFirst
1089           rsPumpData.Find sSearch, , adSearchForward, 1
1090           boWriteDataWritten = True
1091       End If

1092       If Not IsNull(rsPumpData!DataWritten) Or rsPumpData!DataWritten = True Then
1093           ans = MsgBox("You have already entered data for this pump.  Do you want to overwrite the data?", vbDefaultButton2 + vbYesNo, "Overwrite Data?")
1094           If ans = vbNo Then
1095               rsPumpData!DataWritten = True
1096               rsPumpData.Update   'update datawritten
' <VB WATCH>
1097       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1098               Exit Sub
1099           End If
1100       End If

1101       rsPumpData!SerialNumber = frmPLCData.txtSN.Text
1102       rsPumpData!ModelNumber = frmPLCData.txtModelNo.Text
1103       rsPumpData!SalesOrderNumber = frmPLCData.txtSalesOrderNumber.Text
1104       rsPumpData!ShipToCustomer = frmPLCData.txtShpNo.Text
1105       rsPumpData!BillToCustomer = frmPLCData.txtBilNo.Text
1106       rsPumpData!ApplicationFluid = frmPLCData.txtLiquid
1107       rsPumpData!NPSHFile = frmPLCData.txtNPSHFileLocation.Text
1108       rsPumpData!RVSPartNo = frmPLCData.txtRVSPartNo.Text
1109       rsPumpData!CustPN = frmPLCData.txtXPartNum.Text
1110       rsPumpData!CustPO = frmPLCData.txtCustPONum.Text

1111       If Len(frmPLCData.txtViscosity) <> 0 Then
1112           rsPumpData!ApplicationViscosity = frmPLCData.txtViscosity
1113       End If

1114       If frmPLCData.chkSuperMarketFeathered.value = Checked Then
1115           rsPumpData!Field1 = "Feathered"
1116       Else
1117           rsPumpData!Field1 = ""
1118       End If

1119       If LenB(txtSpGr.Text) <> 0 Then
1120           If Not IsNumeric(frmPLCData.txtSpGr.Text) Then
1121               MsgBox "Specific Gravity must be a number."
' <VB WATCH>
1122       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1123               Exit Sub
1124           End If
1125           rsPumpData!SpGr = frmPLCData.txtSpGr.Text
1126       End If
1127       If LenB(txtImpellerDia.Text) <> 0 Then
1128           If Not IsNumeric(frmPLCData.txtImpellerDia.Text) Then
1129               MsgBox "Impeller Diameter must be a number."
' <VB WATCH>
1130       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1131               Exit Sub
1132           End If
1133           rsPumpData!impellerdia = frmPLCData.txtImpellerDia.Text
1134       End If
1135       If LenB(txtDesignFlow.Text) <> 0 Then
1136           rsPumpData!designflow = frmPLCData.txtDesignFlow.Text
1137       End If
1138       If LenB(txtDesignTDH.Text) <> 0 Then
1139           rsPumpData!designtdh = frmPLCData.txtDesignTDH.Text
1140       End If
1141       If LenB(txtRemarks.Text) <> 0 Then
1142           rsPumpData!Remarks = txtRemarks.Text
1143       End If

1144       If optMfr(0).value = True Then
1145           d = cmbMotor.ItemData(cmbMotor.ListIndex)
1146           rsPumpData!Motor = d
1147           d = cmbStatorFill.ItemData(cmbStatorFill.ListIndex)
1148           rsPumpData!StatorFill = d
1149            d = cmbDesignPressure.ItemData(cmbDesignPressure.ListIndex)
1150           rsPumpData!DesignPressure = d
1151           d = cmbCirculationPath.ItemData(cmbCirculationPath.ListIndex)
1152           rsPumpData!CirculationPath = d
1153           d = cmbRPM.ItemData(cmbRPM.ListIndex)
1154           rsPumpData!RPM = d
1155           d = cmbModel.ItemData(cmbModel.ListIndex)
1156           rsPumpData!Model = d
1157           d = cmbModelGroup.ItemData(cmbModelGroup.ListIndex)
1158           rsPumpData!ModelGroup = d
1159       End If
       '   TEMC fields
1160       If optMfr(0).value = False Then
1161           d = cmbTEMCAdapter.ItemData(cmbTEMCAdapter.ListIndex)
1162           rsPumpData!TEMCAdapter = d

1163           d = cmbTEMCAdditions.ItemData(cmbTEMCAdditions.ListIndex)
1164           rsPumpData!TEMCAdditions = d

1165           d = cmbTEMCCirculation.ItemData(cmbTEMCCirculation.ListIndex)
1166           rsPumpData!TEMCcirculation = d

1167           d = cmbTEMCDesignPressure.ItemData(cmbTEMCDesignPressure.ListIndex)
1168           rsPumpData!TEMCDesignpressure = d

1169           d = cmbTEMCDivisionType.ItemData(cmbTEMCDivisionType.ListIndex)
1170           rsPumpData!TEMCDivisionType = d

1171           d = cmbTEMCImpellerType.ItemData(cmbTEMCImpellerType.ListIndex)
1172           rsPumpData!TEMCImpellerType = d

1173           d = cmbTEMCInsulation.ItemData(cmbTEMCInsulation.ListIndex)
1174           rsPumpData!TEMCInsulation = d

1175           d = cmbTEMCJacketGasket.ItemData(cmbTEMCJacketGasket.ListIndex)
1176           rsPumpData!TEMCJacketGasket = d

1177           d = cmbTEMCMaterials.ItemData(cmbTEMCMaterials.ListIndex)
1178           rsPumpData!TEMCMaterials = d

1179           d = cmbTEMCModel.ItemData(cmbTEMCModel.ListIndex)
1180           rsPumpData!TEMCModel = d

1181           d = cmbTEMCNominalImpSize.ItemData(cmbTEMCNominalImpSize.ListIndex)
1182           rsPumpData!TEMCNominalImpSize = d

1183           d = cmbTEMCNominalDischargeSize.ItemData(cmbTEMCNominalDischargeSize.ListIndex)
1184           rsPumpData!TEMCNominalDischargeSize = d

1185           d = cmbTEMCNominalSuctionSize.ItemData(cmbTEMCNominalSuctionSize.ListIndex)
1186           rsPumpData!TEMCNominalSuctionSize = d

1187           d = cmbTEMCOtherMotor.ItemData(cmbTEMCOtherMotor.ListIndex)
1188           rsPumpData!TEMCOtherMotor = d

1189           d = cmbTEMCPumpStages.ItemData(cmbTEMCPumpStages.ListIndex)
1190           rsPumpData!TEMCPumpStages = d

1191           d = cmbTEMCTRG.ItemData(cmbTEMCTRG.ListIndex)
1192           rsPumpData!TEMCTRG = d

1193           d = cmbTEMCVoltage.ItemData(cmbTEMCVoltage.ListIndex)
1194           rsPumpData!TEMCVoltage = d

1195           If LenB(txtTEMCFrameNumber.Text) <> 0 Then
1196               rsPumpData!TEMCFrameNumber = frmPLCData.txtTEMCFrameNumber.Text
1197           End If
1198       End If

1199       rsPumpData!ChempumpPump = optMfr(0).value

1200       rsPumpData!Approved = False

       'added from TEMC Inspection Report
1201       If Len(txtJobNum.Text) <> 0 Then
1202           rsPumpData!JobNumber = txtJobNum.Text
1203       End If

1204       If Len(txtNoPhases.Text) <> 0 Then
1205           rsPumpData!Phases = txtNoPhases.Text
1206       End If

1207       If Len(txtExpClass.Text) <> 0 Then
1208           rsPumpData!ExpClass = txtExpClass.Text
1209       End If

1210       If Len(txtThermalClass.Text) <> 0 Then
1211           rsPumpData!ThermalClass = txtThermalClass.Text
1212       End If

1213       rsPumpData!NPSHr = Val(txtNPSHr.Text)
1214       rsPumpData!LiquidTemperature = Val(txtLiquidTemperature.Text)
1215       rsPumpData!RatedInputPower = Val(txtRatedInputPower.Text)
1216       rsPumpData!FLCurrent = Val(txtAmps.Text)





1217       If boWriteDataWritten Then
1218           rsPumpData!DataWritten = True
1219       Else
1220           rsPumpData!DataWritten = False
1221       End If

           'write the data into the database
1222       rsPumpData.Update
1223       boFoundPump = True

           'enter a new test date if it's a new entry
1224       If Not boWriteDataWritten Then


1225           cmdAddNewTestDate_Click
1226       End If
' <VB WATCH>
1227       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1228       Exit Sub
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
1229       On Error GoTo vbwErrHandler
1230       Const VBWPROCNAME = "frmPLCData.cmdEnterTestData_Click"
1231       If vbwProtector.vbwTraceProc Then
1232           Dim vbwProtectorParameterString As String
1233           If vbwProtector.vbwTraceParameters Then
1234               vbwProtectorParameterString = "()"
1235           End If
1236           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1237       End If
' </VB WATCH>
1238       Dim sSearch As String
1239       Dim ans As Integer

           'if we didn't find the test setup, can't enter test data
1240       If Not boFoundTestSetup Then
1241           MsgBox "You must enter Test Setup Data before entering the Test Data"
' <VB WATCH>
1242       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1243           Exit Sub
1244       End If

           'if we don't find data in the test database, add records
1245       If boFoundTestData = False Then     'add 8 records for 8 tests
1246           AddTestData
1247           rsTestData.MoveFirst
1248       Else        'find the data in the database
1249           sSearch = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"
1250           rsTestData.MoveFirst
1251           rsTestData.Filter = sSearch
1252       End If

           'find the desired record from the form
1253       rsTestData.MoveFirst
1254       rsTestData.Move UpDown1.value - 1

1255       If rsTestData!DataWritten = True Then
1256           ans = MsgBox("You have already entered data for this test.  Do you want to overwrite the data?", vbYesNo + vbDefaultButton2, "Data Already Entered")
1257           If ans = vbNo Then
' <VB WATCH>
1258       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1259               Exit Sub
1260           End If
1261       End If

1262       rsEff.MoveFirst
1263       rsEff.Move UpDown1.value - 1

1264       If LenB(txtV1.Text) <> 0 Then
1265           rsTestData!VoltageA = Val(txtV1.Text)
1266       End If

1267       If LenB(txtV2.Text) <> 0 Then
1268           rsTestData!VoltageB = Val(txtV2.Text)
1269       End If

1270       If LenB(txtV3.Text) <> 0 Then
1271           rsTestData!VoltageC = Val(txtV3.Text)
1272       End If

1273       If LenB(txtI1.Text) <> 0 Then
1274           rsTestData!CurrentA = Val(txtI1.Text)
1275       End If

1276       If LenB(txtI2.Text) <> 0 Then
1277           rsTestData!CurrentB = Val(txtI2.Text)
1278       End If

1279       If LenB(txtI3.Text) <> 0 Then
1280           rsTestData!CurrentC = Val(txtI3.Text)
1281       End If

1282       If LenB(txtP1.Text) <> 0 Then
1283           rsTestData!PowerA = Val(txtP1.Text)
1284       End If

1285       If LenB(txtP2.Text) <> 0 Then
1286           rsTestData!PowerB = Val(txtP2.Text)
1287       End If

1288       If LenB(txtP3.Text) <> 0 Then
1289           rsTestData!PowerC = Val(txtP3.Text)
1290       End If

1291       If LenB(txtKW.Text) <> 0 Then
1292           rsTestData!TotalPower = Val(txtKW.Text)
1293       End If

1294       rsTestData!Flow = Val(txtFlowDisplay.Text)
1295       rsTestData!DischargePressure = Val(txtDischargeDisplay.Text)
1296       rsTestData!SuctionPressure = Val(txtSuctionDisplay.Text)
1297       rsTestData!TemperatureSuction = Val(txtTemperatureDisplay.Text)

1298       rsTestData!TC1 = Val(txtTC1Display.Text)
1299       rsTestData!TC2 = Val(txtTC2Display.Text)
1300       rsTestData!TC3 = Val(txtTC3Display.Text)
1301       rsTestData!TC4 = Val(txtTC4Display.Text)

1302       rsTestData!CircFlow = Val(txtAI1Display.Text)
1303       rsTestData!RBHTemp = Val(txtAI2Display.Text)
1304       rsTestData!RBHPress = Val(txtAI3Display.Text)
1305       rsTestData!AI4 = Val(txtAI4Display.Text)

1306       rsTestData!ValvePosition = Val(txtValvePosition.Text)
1307       rsTestData!SetPoint = Val(txtSetPoint.Text)

1308       If LenB(txtThrustBal.Text) <> 0 Then
1309           rsTestData!ThrustBalance = txtThrustBal.Text
1310       End If

1311       If LenB(txtVibAx.Text) <> 0 Then
1312           rsTestData!VibrationX = txtVibAx.Text
1313       End If

1314       If LenB(txtVibRad.Text) <> 0 Then
1315           rsTestData!VibrationY = txtVibRad.Text
1316       End If

1317       If LenB(txtTEMCTRGReading.Text) <> 0 Then
1318           rsTestData!TEMCTRG = txtTEMCTRGReading.Text
1319       Else
1320           rsTestData!TEMCTRG = 0
1321       End If

1322       If LenB(txtRPM.Text) <> 0 Then
1323           rsTestData!RPM = txtRPM.Text
1324       End If

1325       If LenB(txtTestRemarks.Text) <> 0 Then
1326           rsTestData!Remarks = txtTestRemarks.Text
1327       Else
1328           rsTestData!Remarks = " "
1329       End If

1330       If LenB(txtTEMCTRGReading.Text) <> 0 Then
1331           rsTestData!TEMCTRG = txtTEMCTRGReading.Text
1332       End If

1333       If LenB(txtTEMCFrontThrust.Text) <> 0 Then
1334           rsTestData!TEMCFrontThrust = txtTEMCFrontThrust.Text
1335       End If

1336       If LenB(txtTEMCRearThrust.Text) <> 0 Then
1337           rsTestData!TEMCRearThrust = txtTEMCRearThrust.Text
1338       End If

1339       If LenB(txtTEMCMomentArm.Text) <> 0 Then
1340           rsTestData!TEMCMomentArm = txtTEMCMomentArm.Text
1341       End If

1342       If LenB(txtTEMCThrustRigPressure.Text) <> 0 Then
1343           rsTestData!TEMCThrustRigPressure = txtTEMCThrustRigPressure.Text
1344       End If

1345       If LenB(txtTEMCViscosity.Text) <> 0 Then
1346           rsTestData!TEMCViscosity = txtTEMCViscosity.Text
1347       End If

1348       If LenB(txtNPSHa.Text) <> 0 Then
1349           rsTestData!NPSHa = txtNPSHa.Text
1350       End If

1351       rsTestData!Approved = False

1352       rsTestData!DataWritten = True

           'update the database
1353       rsTestData.Update

1354       DoEfficiencyCalcs
1355       rsEff.Update

           'update the form
1356       DataGrid1.Refresh
1357       DataGrid2.Refresh

1358       FixPointsToPlot

' <VB WATCH>
1359       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1360       Exit Sub
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
1361       On Error GoTo vbwErrHandler
1362       Const VBWPROCNAME = "frmPLCData.cmdEnterTestSetupData_Click"
1363       If vbwProtector.vbwTraceProc Then
1364           Dim vbwProtectorParameterString As String
1365           If vbwProtector.vbwTraceParameters Then
1366               vbwProtectorParameterString = "()"
1367           End If
1368           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1369       End If
' </VB WATCH>
1370       Dim I As Integer
1371       Dim d As Integer
1372       Dim sSearch As String
1373       Dim ans As Integer
1374       Dim boWriteDataWritten As Boolean

           'check for a serial number
1375       If LenB(txtSN.Text) = 0 Then
1376           MsgBox "You must have a Serial Number to enter data.", vbOKOnly + vbExclamation, "Cannot Enter Data"
' <VB WATCH>
1377       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1378           Exit Sub
1379       End If

1380       If Pressed = True Then
1381           If Me.cmbDischDia.ListIndex = -1 Or Me.cmbSuctDia.ListIndex = -1 Or Val(Me.txtSuctHeight.Text) = 0 Or Val(Me.txtDischHeight.Text) = 0 Then
1382               MsgBox "You must have Discharge Diameter AND Suction Diameter AND Suction Height AND Discharge Height entered to enter data.", vbOKOnly + vbExclamation, "Cannot Enter Data"
' <VB WATCH>
1383       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1384               Exit Sub
1385           End If
1386       End If

1387       Pressed = True
1388       If Not boFoundTestSetup Then    'if we didn't find any test setup, add a record
1389           rsTestSetup.AddNew
1390           cmbTestDate.AddItem Now
1391           cmbTestDate.ListIndex = cmbTestDate.NewIndex
1392           cmdAddNewBalanceHoles.Visible = True
1393           boFoundTestSetup = True
1394           boWriteDataWritten = False
1395           rsTestSetup!DataWritten = False
1396       Else    'find the record and display
1397           sSearch = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"
1398           rsTestSetup.MoveFirst
1399           rsTestSetup.Filter = sSearch
1400           If Not boCanApprove Then
       '            cmdAddNewBalanceHoles.Visible = False
1401           End If
1402           boWriteDataWritten = True
1403       End If

1404       If rsTestSetup!DataWritten = True Then
1405           ans = MsgBox("Data has already been entered for this test date.  Do you want to overwrite it?", vbYesNo + vbDefaultButton2, "Data Exists")
1406           If ans = vbNo Then
' <VB WATCH>
1407       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1408               Exit Sub
1409           End If
1410       End If

1411       rsTestSetup!SerialNumber = txtSN
1412       rsTestSetup!Date = cmbTestDate.List(cmbTestDate.ListIndex)

1413       I = cmbFlowMeter.ListIndex
1414       If I = -1 Then
1415           d = 1
1416           rsTestSetup!FlowMeterID = d
1417       Else
1418           d = cmbLoopNumber.ItemData(I)
1419           rsTestSetup!FlowMeterID = d
1420       End If

1421       I = cmbSuctionPressureTransducer.ListIndex
1422       If I = -1 Then
1423           d = 1
1424           rsTestSetup!suctionid = d
1425       Else
1426           d = cmbLoopNumber.ItemData(I)
1427           rsTestSetup!suctionid = d
1428       End If

1429       I = cmbDischargePressureTransducer.ListIndex
1430       If I = -1 Then
1431           d = 1
1432           rsTestSetup!dischid = d
1433       Else
1434           d = cmbLoopNumber.ItemData(I)
1435           rsTestSetup!dischid = d
1436       End If

1437       I = cmbTemperatureTransducer.ListIndex
1438       If I = -1 Then
1439           d = 1
1440           rsTestSetup!temperatureid = d
1441       Else
1442           d = cmbLoopNumber.ItemData(I)
1443           rsTestSetup!temperatureid = d
1444       End If

1445       I = Me.cmbCirculationFlowMeter.ListIndex
1446       If I = -1 Or I < 4 Then
1447           d = 5
1448           rsTestSetup!magflowid = d
1449       Else
1450           d = cmbLoopNumber.ItemData(I)
1451           rsTestSetup!magflowid = d
1452       End If


1453       If LenB(txtHDCor.Text) <> 0 Then
1454           rsTestSetup!HDCor = txtHDCor
1455       Else
1456           rsTestSetup!HDCor = 0
1457       End If
1458       If LenB(txtKWMult.Text) <> 0 Then
1459           rsTestSetup!kwmult = txtKWMult
1460       Else
1461           rsTestSetup!kwmult = 1
1462       End If
1463       If LenB(txtWho.Text) <> 0 Then
1464           rsTestSetup!who = txtWho
1465       Else
1466           rsTestSetup!who = vbNullString
1467       End If
1468       If LenB(txtRMA.Text) <> 0 Then
1469           rsTestSetup!RMA = txtRMA
1470       Else
1471           rsTestSetup!RMA = vbNullString
1472       End If
1473       If LenB(frmPLCData.txtDischHeight) <> 0 Then
1474           rsTestSetup!DischargeGageHeight = Val(txtDischHeight)
1475       Else
1476           rsTestSetup!DischargeGageHeight = 0
1477       End If
1478       If LenB(frmPLCData.txtSuctHeight) <> 0 Then
1479           rsTestSetup!SuctionGageHeight = Val(txtSuctHeight)
1480       Else
1481           rsTestSetup!SuctionGageHeight = 0
1482       End If
1483       If LenB(frmPLCData.txtTestSetupRemarks.Text) <> 0 Then
1484           rsTestSetup!Remarks = txtTestSetupRemarks.Text
1485       Else
1486           rsTestSetup!Remarks = vbNullString
1487       End If
1488       If LenB(frmPLCData.txtVFDFreq.Text) <> 0 Then
1489           rsTestSetup!VFDFrequency = txtVFDFreq.Text
1490       Else
1491           rsTestSetup!VFDFrequency = 0
1492       End If

1493       I = cmbOrificeNumber.ListIndex
1494       If I = -1 Then
1495           d = 18      'entry for None
1496       Else
1497           d = cmbOrificeNumber.ItemData(I)
1498       End If
1499       rsTestSetup!orificenumber = d

1500       If LenB(txtEndPlay.Text) <> 0 Then
1501           rsTestSetup!Endplay = Val(frmPLCData.txtEndPlay.Text)
1502       Else
1503           rsTestSetup!Endplay = 0
1504       End If

1505       If LenB(txtGGap.Text) <> 0 Then
1506           rsTestSetup!GGAP = Val(frmPLCData.txtGGap.Text)
1507       Else
1508           rsTestSetup!GGAP = 0
1509       End If

1510       If LenB(txtOtherMods.Text) <> 0 Then
1511           rsTestSetup!OtherMods = txtOtherMods.Text
1512       Else
1513           rsTestSetup!OtherMods = vbNullString
1514       End If

1515       rsTestSetup!Approved = False

1516       I = cmbLoopNumber.ListIndex
1517       If I = -1 Then
1518           d = 1
1519           rsTestSetup!loopnumber = d
1520       Else
1521           d = cmbLoopNumber.ItemData(I)
1522           rsTestSetup!loopnumber = d
1523       End If

1524       I = cmbSuctDia.ListIndex
1525       If I = -1 Then
1526           d = -1
1527       Else
1528           d = cmbSuctDia.ItemData(I)
1529           rsTestSetup!SuctDiam = d
1530       End If

1531       I = cmbDischDia.ListIndex
1532       If I = -1 Then
1533           d = -1
1534       Else
1535           d = cmbDischDia.ItemData(I)
1536           rsTestSetup!DischDiam = d
1537       End If

1538       I = cmbTachID.ListIndex
1539       If I = -1 Then
1540           d = 1
1541           rsTestSetup!tachid = d
1542       Else
1543           d = cmbTachID.ItemData(I)
1544           rsTestSetup!tachid = d
1545       End If

1546       I = cmbAnalyzerNo.ListIndex
1547       If I = -1 Then
1548           d = 1
1549       Else
1550           d = cmbAnalyzerNo.ItemData(I)
1551       End If
1552       rsTestSetup!analyzerno = d

1553       I = cmbTestSpec.ListIndex
1554       If I = -1 Then
1555           d = 1
1556       Else
1557           d = cmbTestSpec.ItemData(I)
1558       End If
1559       rsTestSetup!testspec = d

1560       I = cmbVoltage.ListIndex
1561       If I = -1 Then
1562           d = 1
1563       Else
1564           d = cmbVoltage.ItemData(I)
1565       End If
1566       rsTestSetup!Voltage = d

1567       I = cmbFrequency.ListIndex
1568       If I = -1 Then
1569           d = 1
1570       Else
1571           d = cmbFrequency.ItemData(I)
1572       End If
1573       rsTestSetup!Frequency = d

1574       I = cmbMounting.ListIndex
1575       If I = -1 Then
1576           d = 1
1577       Else
1578           d = cmbMounting.ItemData(I)
1579       End If
1580       rsTestSetup!Mounting = d

1581       I = cmbPLCNo.ListIndex
1582       If I = -1 Then
1583           d = 8
1584       Else
1585           d = cmbPLCNo.ItemData(I)
1586       End If
1587       rsTestSetup!PLCNo = d

1588       rsTestSetup!ImpFeathered = chkFeathered.value

1589       If chkTrimmed.value = 1 Then
1590           rsTestSetup!ImpTrimmed = Val(txtImpTrim)
1591       Else
1592           rsTestSetup!ImpTrimmed = 0
1593       End If
1594       chkTrimmed_Click

1595       If chkOrifice.value = 1 Then
1596           rsTestSetup!PumpDischOrifice = Val(txtOrifice)
1597       Else
1598           rsTestSetup!PumpDischOrifice = 0
1599       End If
1600       chkOrifice_Click

1601       If chkCircOrifice.value = 1 Then
1602           rsTestSetup!CircFlowOrifice = Val(txtCircOrifice)
1603       Else
1604           rsTestSetup!CircFlowOrifice = 0
1605       End If
1606       chkCircOrifice_Click

1607       chkBalanceHoles_Click

1608       If chkNPSH.value = 1 Then
1609           txtNPSHFile.Visible = True
1610           rsTestSetup!NPSHFile = txtNPSHFile
1611       Else
1612           rsTestSetup!NPSHFile = vbNullString
1613           txtNPSHFile.Visible = False
1614       End If

1615       If chkPictures.value = 1 Then
1616           txtPicturesFile.Visible = True
1617           rsTestSetup!PictureFile = txtPicturesFile
1618       Else
1619           rsTestSetup!PictureFile = vbNullString
1620           txtPicturesFile.Visible = False
1621       End If

1622       If chkVibration.value = 1 Then
1623           txtVibrationFile.Visible = True
1624           rsTestSetup!VibrationFile = txtVibrationFile
1625       Else
1626           rsTestSetup!VibrationFile = vbNullString
1627           txtVibrationFile.Visible = False
1628       End If

1629       If boWriteDataWritten Then
1630           rsTestSetup!DataWritten = True
1631       Else
1632           rsTestSetup!DataWritten = False
1633       End If

           'for TEMC Inspection Report
1634       If LenB(frmPLCData.txtTestAndInspection(0).Text) <> 0 Then
1635           rsTestSetup!InsulationMeggerVolts = frmPLCData.txtTestAndInspection(0).Text
1636       Else
1637           rsTestSetup!InsulationMeggerVolts = ""
1638       End If

1639       If LenB(frmPLCData.txtTestAndInspection(1).Text) <> 0 Then
1640           rsTestSetup!InsulationMegOhms = frmPLCData.txtTestAndInspection(1).Text
1641       Else
1642           rsTestSetup!InsulationMegOhms = ""
1643       End If

1644       If LenB(frmPLCData.txtTestAndInspection(2).Text) <> 0 Then
1645           rsTestSetup!DielectricVolts = frmPLCData.txtTestAndInspection(2).Text
1646       Else
1647           rsTestSetup!DielectricVolts = ""
1648       End If

1649       If LenB(frmPLCData.txtTestAndInspection(3).Text) <> 0 Then
1650           rsTestSetup!DielectricTime = frmPLCData.txtTestAndInspection(3).Text
1651       Else
1652           rsTestSetup!DielectricTime = ""
1653       End If

1654       If LenB(frmPLCData.txtTestAndInspection(4).Text) <> 0 Then
1655           rsTestSetup!HydrostaticValue = frmPLCData.txtTestAndInspection(4).Text
1656       Else
1657           rsTestSetup!HydrostaticValue = ""
1658       End If

1659       If LenB(frmPLCData.txtTestAndInspection(5).Text) <> 0 Then
1660           rsTestSetup!HydrostaticTime = frmPLCData.txtTestAndInspection(5).Text
1661       Else
1662           rsTestSetup!HydrostaticTime = ""
1663       End If

1664       If LenB(frmPLCData.txtTestAndInspection(6).Text) <> 0 Then
1665           rsTestSetup!PneumaticValue = frmPLCData.txtTestAndInspection(6).Text
1666       Else
1667           rsTestSetup!PneumaticValue = ""
1668       End If

1669       If LenB(frmPLCData.txtTestAndInspection(7).Text) <> 0 Then
1670           rsTestSetup!PneumaticTime = frmPLCData.txtTestAndInspection(7).Text
1671       Else
1672           rsTestSetup!PneumaticTime = ""
1673       End If

1674       I = cmbTestAndInspection(0).ListIndex
1675       If I = -1 Then
1676           rsTestSetup!HydrostaticUnits = ""
1677       Else
1678           rsTestSetup!HydrostaticUnits = cmbTestAndInspection(0).Text
1679       End If


1680       I = cmbTestAndInspection(1).ListIndex
1681       If I = -1 Then
1682           rsTestSetup!PneumaticUnits = ""
1683       Else
1684           rsTestSetup!PneumaticUnits = cmbTestAndInspection(1).Text
1685       End If

           'use abs to convert from 1 and 0 to boolean
1686       rsTestSetup!insulationgood = Abs(TestAndInspectionGood(0).value)
1687       rsTestSetup!DielectricGood = Abs(TestAndInspectionGood(1).value)
1688       rsTestSetup!HydrostaticGood = Abs(TestAndInspectionGood(2).value)
1689       rsTestSetup!PneumaticGood = Abs(TestAndInspectionGood(3).value)
1690       rsTestSetup!GeneralAppearanceGood = Abs(TestAndInspectionGood(4).value)
1691       rsTestSetup!OutlineDimensionsGood = Abs(TestAndInspectionGood(5).value)
1692       rsTestSetup!MotorNoLoadTestGood = Abs(TestAndInspectionGood(6).value)
1693       rsTestSetup!MotorLockedRotorTestGood = Abs(TestAndInspectionGood(7).value)
1694       rsTestSetup!HydrostaticTestGood = Abs(TestAndInspectionGood(8).value)
1695       rsTestSetup!HydraulicTestGood = Abs(TestAndInspectionGood(9).value)
1696       rsTestSetup!NPSHTestGood = Abs(TestAndInspectionGood(10).value)
1697       rsTestSetup!CleanPurgeSealGood = Abs(TestAndInspectionGood(11).value)
1698       rsTestSetup!PaintCheckGood = Abs(TestAndInspectionGood(12).value)
1699       rsTestSetup!NameplateGood = Abs(TestAndInspectionGood(13).value)
1700       rsTestSetup!SupervisorApproval = Abs(TestAndInspectionGood(14).value)

           'update the database
1701       rsTestSetup.Update

1702       If boFoundTestData = False Then     'add 8 records for 8 tests
1703           AddTestData
1704       End If

1705       rsTestSetup.Filter = vbNullString
' <VB WATCH>
1706       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1707       Exit Sub
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
1708       On Error GoTo vbwErrHandler
1709       Const VBWPROCNAME = "frmPLCData.cmdExit_Click"
1710       If vbwProtector.vbwTraceProc Then
1711           Dim vbwProtectorParameterString As String
1712           If vbwProtector.vbwTraceParameters Then
1713               vbwProtectorParameterString = "()"
1714           End If
1715           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1716       End If
' </VB WATCH>
1717       End
' <VB WATCH>
1718       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1719       Exit Sub
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
1720       On Error GoTo vbwErrHandler
1721       Const VBWPROCNAME = "frmPLCData.cmdFindMagtrols_Click"
1722       If vbwProtector.vbwTraceProc Then
1723           Dim vbwProtectorParameterString As String
1724           If vbwProtector.vbwTraceParameters Then
1725               vbwProtectorParameterString = "()"
1726           End If
1727           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1728       End If
' </VB WATCH>
1729       FindMagtrols
' <VB WATCH>
1730       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
1731       Exit Sub
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
1732       On Error GoTo vbwErrHandler
1733       Const VBWPROCNAME = "frmPLCData.cmdFindPump_Click"
1734       If vbwProtector.vbwTraceProc Then
1735           Dim vbwProtectorParameterString As String
1736           If vbwProtector.vbwTraceParameters Then
1737               vbwProtectorParameterString = "()"
1738           End If
1739           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
1740       End If
' </VB WATCH>
1741       Dim sAns As String
1742       Dim sSO As String
1743       Dim sParam As String
1744       Dim sName As String

1745       Dim I As Integer

           'clear the data
1746       BlankData

           'set TC and AI labels with default values
1747       txtTitle(0).Text = "TC 1"
1748       txtTitle(1).Text = "(F)"
1749       txtTitle(2).Text = "TC 2"
1750       txtTitle(3).Text = "(F)"
1751       txtTitle(4).Text = "TC 3"
1752       txtTitle(5).Text = "(F)"
1753       txtTitle(6).Text = "TC 4"
1754       txtTitle(7).Text = "(F)"
1755       txtTitle(20).Text = "Circ Flow"
1756       txtTitle(21).Text = "(GPM)"
1757       txtTitle(22).Text = "P1"
1758       txtTitle(23).Text = "(psig)"
1759       txtTitle(24).Text = "P2"
1760       txtTitle(25).Text = "(psig)"
1761       txtTitle(26).Text = "AI 4"
1762       txtTitle(27).Text = ""


1763       For I = 0 To 7
1764           lblAutoMan(I).Caption = "Auto"
1765       Next I

1766       txtFlowDisplay.Enabled = False
1767       txtSuctionDisplay.Enabled = False
1768       txtDischargeDisplay.Enabled = False
1769       txtTemperatureDisplay.Enabled = False
1770       txtAI1Display.Enabled = False
1771       txtAI2Display.Enabled = False
1772       txtAI3Display.Enabled = False
1773       txtAI4Display.Enabled = False


1774       cmdFindPump.Default = False

           'set all found booleans to false
       '    boUsingHP = False
1775       boFoundPump = False
1776       boPumpIsApproved = False
1777       boFoundTestSetup = False
1778       boFoundTestData = False


           'get rid of all test dates in combo box
1779       For I = cmbTestDate.ListCount - 1 To 0 Step -1
1780           cmbTestDate.RemoveItem 0
1781       Next I

1782       rsTestData.Filter = "SerialNumber = ''"

1783       DataGrid2.ClearFields
1784       ClearEff

1785       If rsPumpData.State = adStateOpen Then
1786           If rsPumpData.BOF = False Or rsPumpData.EOF = False Then
1787               rsPumpData.Update
1788           End If
1789           rsPumpData.Close
1790       End If

           'parse the serial number to make sure it is formed correctly
1791       Dim ok As Boolean
1792       ok = UCase(txtSN.Text) Like "[A-Z][A-Z][0-9][0-9][0-9][0-9][A-Z]" Or UCase(txtSN.Text) Like "[A-Z][A-Z][0-9][0-9][0-9][0-9][A-Z]-[0-9]" Or UCase(txtSN.Text) Like "[A-Z][A-Z][0-9][0-9][0-9][0-9][A-Z]-[0-9][0-9]"
1793       If Not ok Then
1794           MsgBox "Serial Number must be 2 letters, 4 numbers, and 1 letter. Please re-enter.", vbOKOnly, "Serial Number not correctly formed."
' <VB WATCH>
1795       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
1796           Exit Sub
1797       End If

           'find the pump listed in the Serial Number text box
1798       qyPumpData.ActiveConnection = cnPumpData
1799       qyPumpData.CommandText = "SELECT * From TempPumpData WHERE (((TempPumpData.SerialNumber)='" & _
                                    txtSN.Text & "'))"
1800       rsPumpData.CursorType = adOpenStatic
1801       rsPumpData.CursorLocation = adUseClient
1802       rsPumpData.Index = "SerialNumber"
1803       rsPumpData.Open qyPumpData
1804       boEpicorFound = False

1805       If rsPumpData.BOF = True And rsPumpData.EOF = True Then
               'if the bof=eof, we have an empty recordset
1806           boFoundPump = False
1807       Else
               'we found it
1808           boFoundPump = True
1809       End If

1810       If boFoundPump = False Then
               'not found in either database, try HP?
1811           sAns = MsgBox("Pump Not Found in the Database.  Look in Epicor?", vbYesNo, "Can't Find Pump")
1812           If sAns = vbNo Then     'new pump - don't get data from HP
1813               boUsingEpicor = False
1814           Else
1815               boUsingEpicor = True
       '            boUsingHP = False
1816           End If
       '        If boUsingEpicor = False Then
       '            sAns = MsgBox("Pump Not Found in the Database.  Look on the HP?", vbYesNo, "Can't Find Pump")
       '            If sAns = vbNo Then     'new pump - don't get data from HP
       '                 boUsingHP = False
       '            Else
       '                boUsingHP = True
       '            End If
       '        End If
1817           EnablePumpDataControls
1818           EnableTestSetupDataControls
1819           EnableTestDataControls
       '        BlankData               'clear any data on the screen
1820           cmdAddNewBalanceHoles.Visible = True

1821       End If

1822       If boFoundPump = True Then    'found the pump
1823           If rsPumpData.Fields("Approved") = True Then
1824               DisablePumpDataControls                         'if it's in the real database, don't allow changes here
1825               boPumpIsApproved = True
1826               lblPumpApproved.Visible = True
1827               If boCanApprove Then
1828                   cmdApprovePump.Caption = "Unapprove this pump"
1829               End If
1830               frmPLCData.cmdApproveTestDate.Enabled = True
1831           Else
1832               EnablePumpDataControls                          'it's in the temp database, allow changes
1833               boPumpIsApproved = False
1834               boTestDateIsApproved = False
1835               lblPumpApproved.Visible = False
1836               If boCanApprove Then
1837                   cmdApprovePump.Caption = "Approve this pump"
1838               End If
1839               cmdApproveTestDate.Caption = "You Must Approve Pump First"
1840               frmPLCData.cmdApproveTestDate.Enabled = False
1841           End If

               'found the pump, show the data
1842           txtModelNo.Text = rsPumpData.Fields("ModelNumber")
1843           frmPLCData.optMfr(0).value = rsPumpData.Fields("ChempumpPump")

1844           If rsPumpData.Fields("ChempumpPump") = True Then
1845               SetCombo cmbMotor, "Motor", rsPumpData
1846               SetCombo cmbDesignPressure, "DesignPressure", rsPumpData
1847               SetCombo cmbRPM, "RPM", rsPumpData
1848               SetCombo cmbCirculationPath, "CirculationPath", rsPumpData
1849               SetCombo cmbStatorFill, "StatorFill", rsPumpData
1850               SetCombo cmbModel, "Model", rsPumpData
1851               SetCombo cmbModelGroup, "ModelGroup", rsPumpData
1852               RatedKW = 999
1853           End If

               'set the TEMC data
1854           If rsPumpData.Fields("ChempumpPump") = False Then
1855               SetCombo cmbTEMCAdapter, "TEMCAdapter", rsPumpData
1856               SetCombo cmbTEMCAdditions, "TEMCAdditions", rsPumpData
1857               SetCombo cmbTEMCCirculation, "TEMCCirculation", rsPumpData
1858               SetCombo cmbTEMCDesignPressure, "TEMCDesignPressure", rsPumpData
1859               SetCombo cmbTEMCNominalDischargeSize, "TEMCNominalDischargeSize", rsPumpData
1860               SetCombo cmbTEMCDivisionType, "TEMCDivisionType", rsPumpData
1861               SetCombo cmbTEMCImpellerType, "TEMCImpellerType", rsPumpData
1862               SetCombo cmbTEMCInsulation, "TEMCInsulation", rsPumpData
1863               SetCombo cmbTEMCJacketGasket, "TEMCJacketGasket", rsPumpData
1864               SetCombo cmbTEMCMaterials, "TEMCMaterials", rsPumpData
1865               SetCombo cmbTEMCModel, "TEMCModel", rsPumpData
1866               SetCombo cmbTEMCNominalImpSize, "TEMCNominalImpSize", rsPumpData
1867               SetCombo cmbTEMCOtherMotor, "TEMCOtherMotor", rsPumpData
1868               SetCombo cmbTEMCPumpStages, "TEMCPumpStages", rsPumpData
1869               SetCombo cmbTEMCNominalSuctionSize, "TEMCNominalSuctionSize", rsPumpData
1870               SetCombo cmbTEMCTRG, "TEMCTRG", rsPumpData
1871               SetCombo cmbTEMCVoltage, "TEMCVoltage", rsPumpData
1872           End If

               'write ship to and bill to info
1873           If Not IsNull(rsPumpData.Fields("ShipToCustomer")) Then
1874               txtShpNo.Text = rsPumpData.Fields("ShipToCustomer")
1875           Else
1876               txtShpNo.Text = vbNullString
1877           End If

1878           If Not IsNull(rsPumpData.Fields("BillToCustomer")) Then
1879               txtBilNo.Text = rsPumpData.Fields("BillToCustomer")
1880           Else
1881               txtBilNo.Text = vbNullString
1882           End If

1883           sName = "ImpellerDia"
1884           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1885               sParam = rsPumpData.Fields(sName)
1886           Else
1887               sParam = vbNullString
1888           End If
1889           txtImpellerDia.Text = sParam

1890           sName = "DesignFlow"
1891           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1892               sParam = rsPumpData.Fields(sName)
1893           Else
1894               sParam = vbNullString
1895           End If
1896           txtDesignFlow.Text = sParam

1897           sName = "DesignTDH"
1898           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1899               sParam = rsPumpData.Fields(sName)
1900           Else
1901               sParam = vbNullString
1902           End If
1903           txtDesignTDH.Text = sParam

1904           sName = "SpGr"
1905           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1906               sParam = rsPumpData.Fields(sName)
1907           Else
1908               sParam = vbNullString
1909           End If
1910           txtSpGr.Text = sParam

1911           sName = "Remarks"
1912           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1913               sParam = rsPumpData.Fields(sName)
1914           Else
1915               sParam = vbNullString
1916           End If
1917           txtRemarks.Text = sParam

1918           sName = "SalesOrderNumber"
1919           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1920               sParam = rsPumpData.Fields(sName)
1921           Else
1922               sParam = vbNullString
1923           End If
1924           txtSalesOrderNumber.Text = sParam

1925           sName = "ApplicationFluid"
1926           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1927               sParam = rsPumpData.Fields(sName)
1928           Else
1929               sParam = vbNullString
1930           End If
1931           txtLiquid.Text = sParam

1932           sName = "NPSHFile"
1933           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1934               sParam = rsPumpData.Fields(sName)
1935           Else
1936               sParam = vbNullString
1937           End If
1938           txtNPSHFileLocation.Text = sParam

1939           sName = "RVSPartNo"
1940           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1941               sParam = rsPumpData.Fields(sName)
1942           Else
1943               sParam = vbNullString
1944           End If
1945           txtRVSPartNo.Text = sParam

1946           sName = "CustPN"
1947           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1948               sParam = rsPumpData.Fields(sName)
1949           Else
1950               sParam = vbNullString
1951           End If
1952           txtXPartNum.Text = sParam

1953           sName = "CustPO"
1954           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1955               sParam = rsPumpData.Fields(sName)
1956           Else
1957               sParam = vbNullString
1958           End If
1959           txtCustPONum.Text = sParam

               'make sure table has custpn - see if last three digits of model no are numeric
       '        sName = "SalesOrderNumber"
       '        If rsPumpData.Fields(sName).ActualSize <> 0 Then
       '            If IsNumeric(Right(rsPumpData.Fields("ModelNumber"), 3)) Then 'no sales order no, must be supermarket
       '                rsPumpData.Fields("CustPN") = rsPumpData.Fields("RVSPartNo")
       '            Else
       '                rsPumpData.Fields("CustPN") = rsPumpData.Fields("ModelNumber")
       '            End If
       '        End If

1960           sName = "ApplicationViscosity"
1961           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1962               sParam = Format(rsPumpData.Fields(sName), "#0.00")
1963           Else
1964               sParam = vbNullString
1965           End If
1966           txtViscosity.Text = sParam

       'added from TEMC Inspection Report
1967           sName = "JobNumber"
1968           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1969               sParam = rsPumpData.Fields(sName)
1970           Else
1971               sParam = ""
1972           End If
1973           txtJobNum.Text = sParam

1974           sName = "Phases"
1975           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1976               sParam = rsPumpData.Fields(sName)
1977           Else
1978               sParam = vbNullString
1979           End If
1980           txtNoPhases.Text = sParam

1981           sName = "ThermalClass"
1982           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1983               sParam = rsPumpData.Fields(sName)
1984           Else
1985               sParam = vbNullString
1986           End If
1987           txtThermalClass.Text = sParam

1988           sName = "ExpClass"
1989           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1990               sParam = rsPumpData.Fields(sName)
1991           Else
1992               sParam = vbNullString
1993           End If
1994           txtExpClass.Text = sParam

1995           sName = "NPSHr"
1996           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1997               sParam = rsPumpData.Fields(sName)
1998           Else
1999               sParam = vbNullString
2000           End If
2001           txtNPSHr.Text = sParam

2002           sName = "LiquidTemperature"
2003           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2004               sParam = rsPumpData.Fields(sName)
2005           Else
2006               sParam = vbNullString
2007           End If
2008           txtLiquidTemperature.Text = sParam

2009           sName = "RatedInputPower"
2010           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2011               sParam = rsPumpData.Fields(sName)
2012           Else
2013               sParam = vbNullString
2014           End If
2015           txtRatedInputPower.Text = sParam

2016           sName = "FLCurrent"
2017           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2018               sParam = rsPumpData.Fields(sName)
2019           Else
2020               sParam = vbNullString
2021           End If
2022           txtAmps.Text = sParam

2023           sName = "TEMCFrameNumber"
2024           If rsPumpData.Fields(sName).ActualSize <> 0 Then
2025               sParam = rsPumpData.Fields(sName)
2026           Else
2027               sParam = vbNullString
2028           End If
2029           txtTEMCFrameNumber.Text = sParam

2030           optMfr(0).value = rsPumpData.Fields("ChempumpPump")
2031           optMfr(1).value = Not optMfr(0).value

2032           If rsPumpData.Fields("Field1") = "Feathered" Then
2033               Me.chkSuperMarketFeathered.value = Checked
2034           Else
2035               Me.chkSuperMarketFeathered.value = Unchecked
2036           End If

               'select the testsetup data
2037           qyTestSetup.ActiveConnection = cnPumpData
2038           qyTestSetup.CommandText = "SELECT * FROM TempTestSetupData WHERE (((TempTestSetupData.SerialNumber)='" & _
                                    txtSN.Text & "')) ORDER BY Date"
       '        qyTestSetup.CommandText = "SELECT * FROM TempTestSetupData WHERE (((TempTestSetupData.SerialNumber)='" & _
'               txtSN.Text & "'))"

2039           With rsTestSetup
2040               If .State = adStateOpen Then
2041                   .Close
2042               End If
2043               .CursorLocation = adUseClient
2044               .CursorType = adOpenStatic
2045               .Index = "FindData"
2046               .Open qyTestSetup
2047           End With


               'add the selection of dates to the Test Date combo box
2048           If rsTestSetup.RecordCount <> 0 Then
2049               For I = 0 To cmbTestDate.ListCount - 1
2050                   cmbTestDate.RemoveItem 0
2051               Next I
2052               rsTestSetup.MoveFirst
2053               For I = 1 To rsTestSetup.RecordCount
2054                   cmbTestDate.AddItem rsTestSetup.Fields("Date")
2055                   rsTestSetup.MoveNext
2056               Next I
2057               rsTestSetup.MoveFirst
2058               boFoundTestSetup = True

2059               If rsTestSetup.Fields("Approved") = True Then
2060                   DisableTestSetupDataControls                         'if it's in the real database, don't allow changes here
2061                   boTestDateIsApproved = True
2062                   lblTestDateApproved.Visible = True
2063                   If boCanApprove Then
2064                       cmdApproveTestDate.Caption = "Unapprove this Test Date"
2065                   End If
2066               Else
2067                   EnableTestSetupDataControls                          'it's in the temp database, allow changes
2068                   lblTestDateApproved.Visible = False
2069                   If boCanApprove Then
2070                       cmdApproveTestDate.Caption = "Approve this Test Date"
2071                   End If
2072               End If
2073               cmbTestDate.ListIndex = 0
2074           Else
2075               MsgBox ("There is no Test Setup Data for Serial Number " & txtSN.Text)
2076               boFoundTestSetup = False        'didn't find any data
2077               boFoundTestData = False
2078               cmbTestDate.AddItem Date        'load with today
2079               cmbTestDate.ListIndex = 0       'show the entry
2080               EnableTestSetupDataControls
2081               txtTestRemarks.Text = ""
2082               txtVibAx.Text = ""
2083               txtVibRad.Text = ""
2084               txtThrustBal.Text = ""
2085               txtTEMCTRGReading.Text = ""
2086               txtTEMCFrontThrust.Text = ""
2087               txtTEMCRearThrust.Text = ""
' <VB WATCH>
2088       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2089               Exit Sub
2090           End If

2091           If cmbTestDate.ListCount = 1 Then       'if there's only one test date, select it
2092           End If
' <VB WATCH>
2093       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2094           Exit Sub
2095       End If


2096       Do While boUsingEpicor = True   'need a do loop to exit
2097           If boUsingEpicor = True Then
                   'Dim MyRecord As SNRecord
2098               Dim MyRecord As SNRecord
           '            I = InStr(1, txtSN.Text, "-")
           '            If I > 0 Then
2099                   MyRecord = GetEpicorODBCData(txtSN.Text, EpicorConnectionString)
           '            End If
2100               If MyRecord.SONumber = "" Then
2101                   MsgBox ("Not found in Epicor")
2102                   boUsingEpicor = False
2103                   boEpicorFound = False
2104                   Exit Do
2105               End If

2106               If MyRecord.SONumber = 0 Then
2107                   boEpicorFound = False
2108                   boUsingSupermarketTable = True
2109                   boUsingEpicor = False
2110               Else
2111                   boEpicorFound = True
2112                   boUsingSupermarketTable = False
2113               End If

2114               If boEpicorFound = True Then
2115                   boUsingEpicor = False
       '                boEpicorFound = True
2116                   txtSalesOrderNumber.Text = MyRecord.SONumber
2117                   txtLineNumber.Text = MyRecord.SOLine
2118                   txtBilNo.Text = MyRecord.Customer
2119                   txtXPartNum.Text = MyRecord.XPartNum
2120                   txtCustPONum.Text = MyRecord.CustomerPO

2121                   If MyRecord.ShipTo = "" Then
2122                       txtShpNo.Text = MyRecord.Customer
2123                   Else
2124                       txtShpNo.Text = MyRecord.ShipTo
2125                   End If
2126                   txtModelNo.Text = MyRecord.PartNum
2127                   txtModelNo_Change
2128                   txtDesignTDH.Text = MyRecord.TDH
2129                   txtSpGr.Text = MyRecord.SpGr
2130                   txtImpellerDia.Text = MyRecord.ImpellerDiameter
2131                   txtDesignFlow.Text = MyRecord.Flow
2132                   txtNoPhases.Text = MyRecord.Phases
2133                   txtNPSHr.Text = MyRecord.NPSHr
2134                   txtRatedInputPower.Text = MyRecord.RatedInputPower
2135                   txtAmps.Text = MyRecord.FLCurrent
2136                   txtThermalClass.Text = MyRecord.ThermalClass
2137                   txtViscosity.Text = MyRecord.Viscosity
2138                   txtExpClass.Text = MyRecord.ExpClass
2139                   txtLiquidTemperature.Text = MyRecord.LiquidTemp
2140                   txtLiquid.Text = MyRecord.Fluid
2141                   txtJobNum.Text = MyRecord.JobNumber

2142                   For I = 0 To cmbStatorFill.ListCount - 1
2143                       If InStr(1, UCase$(MyRecord.StatorFill), UCase$(cmbStatorFill.List(I))) <> 0 Then
2144                           cmbStatorFill.ListIndex = I
2145                           Exit For
2146                       End If
2147                   Next I

2148                   For I = 0 To cmbCirculationPath.ListCount - 1
2149                       If InStr(1, UCase$(MyRecord.CirculationPath), UCase$(cmbCirculationPath.List(I))) <> 0 Then
2150                           cmbCirculationPath.ListIndex = I
2151                           Exit For
2152                       End If
2153                   Next I

2154                   For I = 0 To cmbDesignPressure.ListCount - 1
2155                       If InStr(1, MyRecord.DesignPressure, cmbDesignPressure.List(I)) <> 0 Then
2156                           cmbDesignPressure.ListIndex = I
2157                           Exit For
2158                       End If
2159                   Next I

2160                   For I = 0 To cmbVoltage.ListCount - 1
2161                       If InStr(1, MyRecord.Voltage, cmbVoltage.List(I)) <> 0 Then
2162                           cmbVoltage.ListIndex = I
2163                           Exit For
2164                       End If
2165                   Next I

2166                   For I = 0 To cmbFrequency.ListCount - 1
2167                       If InStr(1, MyRecord.Frequency, sName) <> 0 Then
2168                           cmbFrequency.ListIndex = I
2169                           Exit For
2170                       End If
2171                   Next I

2172                   For I = 0 To cmbRPM.ListCount - 1
2173                       If InStr(1, MyRecord.RPM, cmbRPM.List(I)) <> 0 Then
2174                           cmbRPM.ListIndex = I
2175                           Exit For
2176                       End If
2177                   Next I

2178                   For I = 0 To cmbSuctDia.ListCount - 1
2179                       If InStr(1, MyRecord.SuctFlangeSize, cmbSuctDia.List(I)) <> 0 Then
2180                           cmbSuctDia.ListIndex = I
2181                           Exit For
2182                       End If
2183                   Next I

2184                   For I = 0 To cmbDischDia.ListCount - 1
2185                       If InStr(1, MyRecord.DischFlangeSize, cmbDischDia.List(I)) <> 0 Then
2186                           cmbDischDia.ListIndex = I
2187                           Exit For
2188                       End If
2189                   Next I

2190                   For I = 0 To cmbTestSpec.ListCount - 1
2191                       If InStr(1, MyRecord.TestProcedure, cmbTestSpec.List(I)) <> 0 Then
2192                           cmbTestSpec.ListIndex = I
2193                           Exit For
2194                       End If
2195                   Next I

2196                   For I = 0 To cmbMotor.ListCount - 1
2197                       If InStr(1, MyRecord.MotorSize, cmbMotor.List(I)) <> 0 Then
2198                           cmbMotor.ListIndex = I
2199                           Exit For
2200                       End If
2201                   Next I


2202               End If
2203           End If
2204       Loop

2205       If boUsingSupermarketTable = True Then
2206           GetSuperMarketPump MyRecord.PartNum, MyRecord.JobNumber
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
2207       End If
' <VB WATCH>
2208       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2209       Exit Sub
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
2210       On Error GoTo vbwErrHandler
2211       Const VBWPROCNAME = "frmPLCData.cmdModifyBalanceHoleData_Click"
2212       If vbwProtector.vbwTraceProc Then
2213           Dim vbwProtectorParameterString As String
2214           If vbwProtector.vbwTraceParameters Then
2215               vbwProtectorParameterString = "()"
2216           End If
2217           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2218       End If
' </VB WATCH>
2219       Dim strInput As String
2220       Dim I As Integer
2221       Dim sNumber As Integer
2222       Dim sDia As String
2223       Dim sBC As String

2224       cmdModifyBalanceHoleData.Visible = False

2225       If dgBalanceHoles.SelBookmarks.Count = 0 Then
2226           cmdModifyBalanceHoleData.Visible = False
' <VB WATCH>
2227       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2228           Exit Sub
2229       End If

2230       rsBalanceHoles.MoveFirst
2231       rsBalanceHoles.Move dgBalanceHoles.SelBookmarks(0) - dgBalanceHoles.FirstRow

2232       sNumber = rsBalanceHoles!Number
2233       If rsBalanceHoles!diameter = 99 Then
2234           sDia = "Slot"
2235       Else
2236           sDia = str(rsBalanceHoles!diameter)
2237       End If
2238       If rsBalanceHoles!boltcircle = 99 Then
2239           sBC = "Unknown"
2240       Else
2241           sBC = str(rsBalanceHoles!boltcircle)
2242       End If


           'get the data for the balance holes
2243       strInput = InputBox("Enter Number of Holes (0 to delete entry)", , sNumber)
2244       If strInput = "" Then
2245           GoTo DeleteIt
2246       End If
2247       sNumber = CInt(strInput)
2248       If Val(sNumber) = 0 Then
2249           GoTo DeleteIt
2250       End If

2251       strInput = InputBox("Enter Decimal Value of Hole Diameter or 'Slot' (For Example, 0.675) ", , sDia)
2252       If strInput <> "" Then
2253           If UCase(strInput) = "SLOT" Then
2254               strInput = 99
2255           End If
2256           sDia = CSng(strInput)
2257       Else
2258           GoTo CancelPressed
2259       End If

2260       strInput = InputBox("Enter Decimal Value of Bolt Circle or 'Unknown' (For Example, 4.525)", , sBC)
2261       If strInput <> "" Then
2262           If UCase(strInput) = "UNKNOWN" Then
2263               strInput = 99
2264           End If
2265           sBC = CSng(strInput)
2266       Else
2267           GoTo CancelPressed
2268       End If

2269       rsBalanceHoles!Number = sNumber
2270       rsBalanceHoles!diameter = sDia
2271       rsBalanceHoles!boltcircle = sBC

2272       rsBalanceHoles.Update
           'rsBalanceHoles.Filter = "SerialNo = '" & frmPLCData.txtSN.Text & "'"

2273       GetBalanceHoleData txtSN.Text, cmbTestDate.Text
       '    rsBalanceHoles.Requery
2274       rsBalanceHoles.MoveLast
2275       dgBalanceHoles.Refresh
2276       chkBalanceHoles.value = 1
2277       rsBalanceHoles.MoveFirst

' <VB WATCH>
2278       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2279       Exit Sub

2280   CancelPressed:
2281       MsgBox "No New Balance Hole Data Entered", vbOKOnly

2282   DeleteIt:
2283       If (MsgBox("Do you really want to delete this entry?", vbYesNo, "Deleting Balance Hole Data. . .")) = vbYes Then
2284           rsBalanceHoles.Delete
2285           rsBalanceHoles.Update
2286           GetBalanceHoleData txtSN.Text, cmbTestDate.Text
       '        rsBalanceHoles.Requery
2287           If Not rsBalanceHoles.EOF Then
2288               rsBalanceHoles.MoveLast
2289           End If
2290           dgBalanceHoles.Refresh
2291           chkBalanceHoles.value = 1
2292           If Not rsBalanceHoles.BOF Then
2293               rsBalanceHoles.MoveFirst
2294           End If
2295       End If


' <VB WATCH>
2296       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2297       Exit Sub
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
2298       On Error GoTo vbwErrHandler
2299       Const VBWPROCNAME = "frmPLCData.cmdReport_Click"
2300       If vbwProtector.vbwTraceProc Then
2301           Dim vbwProtectorParameterString As String
2302           If vbwProtector.vbwTraceParameters Then
2303               vbwProtectorParameterString = "()"
2304           End If
2305           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2306       End If
' </VB WATCH>
2307       Dim I As Integer

2308       ExportToExcel

' <VB WATCH>
2309       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2310       Exit Sub
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
2311       On Error GoTo vbwErrHandler
2312       Const VBWPROCNAME = "frmPLCData.cmdSearchForPump_Click"
2313       If vbwProtector.vbwTraceProc Then
2314           Dim vbwProtectorParameterString As String
2315           If vbwProtector.vbwTraceParameters Then
2316               vbwProtectorParameterString = "()"
2317           End If
2318           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2319       End If
' </VB WATCH>
2320       LoadCombo frmSearch.cmbSearchModel, "TEMCHydraulics"

2321       frmSearch.Show
' <VB WATCH>
2322       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2323       Exit Sub
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
2324       On Error GoTo vbwErrHandler
2325       Const VBWPROCNAME = "frmPLCData.cmdSelectSupermarket_Click"
2326       If vbwProtector.vbwTraceProc Then
2327           Dim vbwProtectorParameterString As String
2328           If vbwProtector.vbwTraceParameters Then
2329               vbwProtectorParameterString = "()"
2330           End If
2331           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2332       End If
' </VB WATCH>
2333       grpSupermarket.Visible = False
' <VB WATCH>
2334       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2335       Exit Sub
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
2336       On Error GoTo vbwErrHandler
2337       Const VBWPROCNAME = "frmPLCData.cmdWriteSP_Click"
2338       If vbwProtector.vbwTraceProc Then
2339           Dim vbwProtectorParameterString As String
2340           If vbwProtector.vbwTraceParameters Then
2341               vbwProtectorParameterString = "()"
2342           End If
2343           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2344       End If
' </VB WATCH>
2345       Dim rc As String
2346       Dim S As String

           'write the set point data to the PLC
2347           bWrite = True
2348           S = Right$("0000" & txtWriteSPData, 4)
2349           S = Right$(S, 2) & Left$(S, 2)
2350           rc = StringToByteArray(S, ByteBuffer)

2351           DataLength = HexConvert(ByteBuffer, 2)
2352           DataAddress = StringToHexInt("2005")

2353           rc = GetData

2354           bWrite = False
' <VB WATCH>
2355       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2356       Exit Sub
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
2357       On Error GoTo vbwErrHandler
2358       Const VBWPROCNAME = "frmPLCData.btnRunNPSH_Click"
2359       If vbwProtector.vbwTraceProc Then
2360           Dim vbwProtectorParameterString As String
2361           If vbwProtector.vbwTraceParameters Then
2362               vbwProtectorParameterString = "()"
2363           End If
2364           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2365       End If
' </VB WATCH>
2366       Static OriginalColor As Long
2367       If btnRunNPSH.Caption = "Run NPSH" Then
2368           btnRunNPSH.Caption = "Cancel NPSH Run"
2369           OriginalColor = btnRunNPSH.BackColor
2370           tmrNPSHr.Enabled = False
2371           btnRunNPSH.BackColor = vbRed
2372           If boCanApprove Then
2373               txtNPSH(5).Visible = True
2374               lbltab4(5).Visible = True
2375           Else
2376               txtNPSH(5).Visible = False
2377               lbltab4(5).Visible = False
2378           End If
2379           WroteNPSHr = False

2380           frmNPSH.Visible = True
2381           txtNPSH(5).Enabled = True
2382           If Val(txtTDH.Text) <= 10 Then
2383               MsgBox "This test will not work starting with this starting TDH.  Ending test...", vbOKOnly, "Flow is 0"
2384               btnRunNPSH.Caption = "Run NPSH"
2385               btnRunNPSH.BackColor = OriginalColor
2386               frmNPSH.Visible = False
' <VB WATCH>
2387       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2388               Exit Sub
2389           End If
               'load initial values
2390           If DataGrid2.Row = -1 Then
2391               MsgBox "You must write the normal test data to this row before you run NPSH.", vbOKOnly, "Nothing written for this row"
2392               btnRunNPSH.Caption = "Run NPSH"
2393               btnRunNPSH.BackColor = OriginalColor
2394               frmNPSH.Visible = False
' <VB WATCH>
2395       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2396               Exit Sub
2397           Else
2398               DataGrid2.Row = UpDown1.value - 1
2399           End If

2400           txtNPSH(0).Text = DataGrid2.Columns("Flow")
2401           txtNPSH(3).Text = DataGrid2.Columns("TDH")
2402           txtNPSH(4) = 0
               'txtNPSH(0).Text = txtFlow.Text
               'txtNPSH(3).Text = txtTDH.Text
2403           txtNPSH(4) = 0
2404       Else
2405           btnRunNPSH.Caption = "Run NPSH"
2406           btnRunNPSH.BackColor = OriginalColor
2407           frmNPSH.Visible = False
2408       End If

           'ReportToExcel
' <VB WATCH>
2409       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2410       Exit Sub
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
2411       On Error GoTo vbwErrHandler
2412       Const VBWPROCNAME = "frmPLCData.updown1_change"
2413       If vbwProtector.vbwTraceProc Then
2414           Dim vbwProtectorParameterString As String
2415           If vbwProtector.vbwTraceParameters Then
2416               vbwProtectorParameterString = "()"
2417           End If
2418           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2419       End If
' </VB WATCH>
2420       Dim sName As String

2421       If Not rsTestData.BOF Then
2422           rsTestData.MoveFirst
2423       End If

2424       If Not rsTestData.BOF Or Not rsTestData.EOF Then
2425           rsTestData.Move UpDown1.value - 1
2426       End If

2427       sName = "VibrationX"
2428       If rsTestData.Fields(sName).ActualSize <> 0 Then
2429           txtVibAx.Text = rsTestData.Fields(sName)
2430       Else
       '        txtVibAx.Text = vbNullString
2431       End If

2432       sName = "VibrationY"
2433       If rsTestData.Fields(sName).ActualSize <> 0 Then
2434           txtVibRad.Text = rsTestData.Fields(sName)
2435       Else
       '        txtVibRad.Text = vbNullString
2436       End If

2437       sName = "Remarks"
2438       If rsTestData.Fields(sName).ActualSize <> 0 Then
2439           txtTestRemarks.Text = rsTestData.Fields(sName)
2440       Else
       '        txtTestRemarks.Text = vbNullString
2441       End If

2442       sName = "ThrustBalance"
2443       If rsTestData.Fields(sName).ActualSize <> 0 Then
2444           txtThrustBal.Text = rsTestData.Fields(sName)
2445       Else
       '        txtThrustBal.Text = vbNullString
2446       End If

2447       sName = "TEMCTRG"
2448       If rsTestData.Fields(sName).ActualSize <> 0 Then
2449           txtTEMCTRGReading.Text = rsTestData.Fields(sName)
2450       Else
2451           txtTEMCTRGReading.Text = 0
       '        txtTEMCTRGReading.Text = vbNullString
2452       End If

2453       sName = "TEMCFrontThrust"
2454       If rsTestData.Fields(sName).ActualSize <> 0 Then
2455           txtTEMCFrontThrust.Text = rsTestData.Fields(sName)
2456       Else
       '        txtTEMCFrontThrust.Text = vbNullString
2457       End If

2458       sName = "TEMCRearThrust"
2459       If rsTestData.Fields(sName).ActualSize <> 0 Then
2460           txtTEMCRearThrust.Text = rsTestData.Fields(sName)
2461       Else
       '        txtTEMCRearThrust.Text = vbNullString
2462       End If
2463       sName = "TEMCMomentArm"
2464       If rsTestData.Fields(sName).ActualSize <> 0 Then
2465           txtTEMCMomentArm.Text = rsTestData.Fields(sName)
2466       Else
       '        txtTEMCMomentArm.Text = vbNullString
2467       End If
2468       sName = "TEMCThrustRigPressure"
2469       If rsTestData.Fields(sName).ActualSize <> 0 Then
2470           txtTEMCThrustRigPressure.Text = rsTestData.Fields(sName)
2471       Else
       '        txtTEMCThrustRigPressure.Text = vbNullString
2472       End If
2473       sName = "TEMCViscosity"
2474       If rsTestData.Fields(sName).ActualSize <> 0 And rsTestData.Fields(sName) <> 0 Then
2475           txtTEMCViscosity.Text = rsTestData.Fields(sName)
2476       Else
       '        txtTEMCViscosity.Text = vbNullString
2477       End If

2478       CalculateTEMCForce

2479       rsEff.MoveFirst
2480       rsEff.Move UpDown1.value - 1
' <VB WATCH>
2481       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2482       Exit Sub
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
2483       On Error GoTo vbwErrHandler
2484       Const VBWPROCNAME = "frmPLCData.CalculateTEMCForce"
2485       If vbwProtector.vbwTraceProc Then
2486           Dim vbwProtectorParameterString As String
2487           If vbwProtector.vbwTraceParameters Then
2488               vbwProtectorParameterString = "()"
2489           End If
2490           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2491       End If
' </VB WATCH>
2492       Dim NoOfPoles As Integer
2493       Dim Frequency As Integer
2494       Dim Additions As String
2495       Dim Frame As String
2496       Dim VOverA As Double
2497       Dim Force As Double
2498       Dim Gravity As Double

2499       If Val(txtSpGr.Text) = 0 Then
2500           Gravity = 1
2501       Else
2502           Gravity = CDbl(Val(txtSpGr.Text))
2503       End If

           'show calculated values
2504       If Val(txtTEMCFrontThrust.Text) = 0 Then
2505           If Val(txtTEMCRearThrust.Text) = 0 Then
               'no thrust entered
2506               lblTEMCFrontRear.Visible = False
2507               txtTEMCCalcForce.Text = " "
2508           Else
                   'rear thrust
2509               txtTEMCCalcForce.Text = Format(Gravity * (Val(txtTEMCRearThrust.Text) * Val(txtTEMCMomentArm.Text) - (Val(txtTEMCThrustRigPressure.Text) / 14.223) * 4.5), "##0.0")
2510               lblTEMCFrontRear.Caption = "REAR"
2511               lblTEMCFrontRear.Visible = True
2512           End If
2513       Else
               'front thrust
2514           txtTEMCCalcForce.Text = Format(Gravity * (Val(txtTEMCFrontThrust.Text) * Val(txtTEMCMomentArm.Text) + (Val(txtTEMCThrustRigPressure.Text) / 14.223) * 4.5), "##0.0")
2515           lblTEMCFrontRear.Caption = "FRONT"
2516           lblTEMCFrontRear.Visible = True
2517       End If

2518       If Val(txtTEMCCalcForce.Text) < 0 Then
2519           txtTEMCCalcForce.Text = -txtTEMCCalcForce
2520           lblTEMCFrontRear.Caption = "FRONT"
2521       End If

           'see how many poles we have, it's the next to last number in the frame size
2522       If Len(txtTEMCFrameNumber) > 2 Then
2523           NoOfPoles = 2 * Val(Left$(Right$(txtTEMCFrameNumber.Text, 2), 1))
2524       End If

2525       If cmbTEMCAdditions.ListIndex <> -1 Then
2526           Additions = Mid$(cmbTEMCAdditions.List(cmbTEMCAdditions.ListIndex), 2, 1)
2527           If Additions = "A" Or Additions = "E" Or Additions = "G" Or Additions = "J" Then
2528               Frequency = 60
2529           ElseIf Additions = "B" Or Additions = "F" Or Additions = "H" Or Additions = "K" Then
2530               Frequency = 50
2531           Else
2532               Frequency = 0
2533           End If
2534       End If

2535       If Len(txtTEMCFrameNumber.Text) = 3 Then
2536           If txtTEMCFrameNumber.Text = "529" Then
2537               Frame = "420"
2538           Else
2539               Frame = Left$(txtTEMCFrameNumber, 2) & "0"
2540           End If
2541       Else
2542           Frame = txtTEMCFrameNumber.Text
2543           If Right$(txtTEMCFrameNumber.Text, 1) = "5" Then
2544               Frame = Frame & Left$(lblTEMCFrontRear.Caption, 1)
2545           Else
2546           End If
2547       End If
2548       Force = DLookupA(3, TEMCForceViscosity, 1, Frame)
2549       If Frequency = 60 Then
2550           Force = Force / 1.2
2551       End If
2552       If Val(txtTEMCViscosity.Text) > 1# Then
2553           If (Val(txtTEMCCalcForce.Text) > 3 * Force) Then
2554               lblTEMCPassFail.Visible = True
2555               lblTEMCPassFail.ForeColor = vbRed
2556               lblTEMCPassFail.Caption = "FAIL"
2557           Else
2558               lblTEMCPassFail.Visible = True
2559               lblTEMCPassFail.ForeColor = vbGreen
2560               lblTEMCPassFail.Caption = "PASS"
2561           End If
2562       End If

2563       If (Val(txtTEMCViscosity.Text) > 0.5) And (Val(txtTEMCViscosity.Text) <= 1#) Then
2564           If (Val(txtTEMCCalcForce.Text) > 2 * Force) Then
2565               lblTEMCPassFail.Visible = True
2566               lblTEMCPassFail.ForeColor = vbRed
2567               lblTEMCPassFail.Caption = "FAIL"
2568           Else
2569               lblTEMCPassFail.Visible = True
2570               lblTEMCPassFail.ForeColor = vbGreen
2571               lblTEMCPassFail.Caption = "PASS"
2572           End If
2573       End If

2574       If (Val(txtTEMCViscosity.Text) > 0.3) And (Val(txtTEMCViscosity.Text) <= 0.5) Then
2575           If (Val(txtTEMCCalcForce.Text) > 1.5 * Force) Then
2576               lblTEMCPassFail.Visible = True
2577               lblTEMCPassFail.ForeColor = vbRed
2578               lblTEMCPassFail.Caption = "FAIL"
2579           Else
2580               lblTEMCPassFail.Visible = True
2581               lblTEMCPassFail.ForeColor = vbGreen
2582               lblTEMCPassFail.Caption = "PASS"
2583           End If
2584       End If

2585       If (Val(txtTEMCViscosity.Text) <= 0.3) Then
2586           If (Val(txtTEMCCalcForce.Text) > 1# * Force) Then
2587               lblTEMCPassFail.Visible = True
2588               lblTEMCPassFail.ForeColor = vbRed
2589               lblTEMCPassFail.Caption = "FAIL"
2590           Else
2591               lblTEMCPassFail.Visible = True
2592               lblTEMCPassFail.ForeColor = vbGreen
2593               lblTEMCPassFail.Caption = "PASS"
2594           End If
2595       End If
2596       If NoOfPoles <> 0 Then
2597           VOverA = (DLookupA(2, TEMCForceViscosity, 1, Frame)) / (NoOfPoles * 30 / Frequency)
2598       End If
       '    If Frequency = 60 Then
       '        VOverA = VOverA * 1.2
       '    End If

2599       txtTEMCPVValue.Text = Format(Val(txtTEMCCalcForce.Text) * VOverA, "##0.0")

2600       If Val(txtTEMCFrontThrust.Text) = 0 And Val(txtTEMCRearThrust.Text) = 0 Then
2601           txtTEMCPVValue.Text = ""
2602           txtTEMCCalcForce.Text = ""
2603           lblTEMCPassFail.Visible = False
2604       End If


           'calculate reverse head
2605       txtRevHead.Text = Format(rsTestData.Fields("RBHPress") - rsTestData.Fields("SuctionPressure") * 2.31, "##0.0")
       '    txtRevHead.Text = Format((CDbl(Val(txtAI3Display.Text)) - CDbl(Val(txtSuctionDisplay.Text))) * 2.31, "##0.0")

' <VB WATCH>
2606       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2607       Exit Sub
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
2608       On Error GoTo vbwErrHandler
2609       Const VBWPROCNAME = "frmPLCData.updown2_change"
2610       If vbwProtector.vbwTraceProc Then
2611           Dim vbwProtectorParameterString As String
2612           If vbwProtector.vbwTraceParameters Then
2613               vbwProtectorParameterString = "()"
2614           End If
2615           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2616       End If
' </VB WATCH>
2617       Dim Plothead(1, 7) As Single
2618       Dim HeadPlot(7, 1) As Single

2619       Dim PlotEff() As Single
2620       Dim PlotKW() As Single
2621       Dim PlotAmps() As Single

2622       Dim j As Integer

2623       For j = 0 To UpDown2.value - 1
2624           Plothead(0, j) = HeadFlow(0, j)
2625           Plothead(1, j) = HeadFlow(1, j)
2626           HeadPlot(j, 0) = FlowHead(j, 0)
2627           HeadPlot(j, 1) = FlowHead(j, 1)
       '        ReDim Preserve PlotEff(1, j)
       '        PlotEff(0, j) = EffFlow(0, j)
       '        PlotEff(1, j) = EffFlow(1, j)
       '        ReDim Preserve PlotKW(1, j)
       '        PlotKW(0, j) = KWFlow(0, j)
       '        PlotKW(1, j) = KWFlow(1, j)
       '        ReDim Preserve PlotAmps(1, j)
       '        PlotAmps(0, j) = AmpsFlow(0, j)
       '        PlotAmps(1, j) = AmpsFlow(1, j)
2628       Next j

2629       MSChart1 = HeadPlot

' <VB WATCH>
2630       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2631       Exit Sub
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
2632       On Error GoTo vbwErrHandler
2633       Const VBWPROCNAME = "frmPLCData.DataGrid1_AfterColUpdate"
2634       If vbwProtector.vbwTraceProc Then
2635           Dim vbwProtectorParameterString As String
2636           If vbwProtector.vbwTraceParameters Then
2637               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ColIndex", ColIndex) & ") "
2638           End If
2639           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2640       End If
' </VB WATCH>
2641       DoEfficiencyCalcs
' <VB WATCH>
2642       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2643       Exit Sub
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
2644       On Error GoTo vbwErrHandler
2645       Const VBWPROCNAME = "frmPLCData.dgBalanceHoles_SelChange"
2646       If vbwProtector.vbwTraceProc Then
2647           Dim vbwProtectorParameterString As String
2648           If vbwProtector.vbwTraceParameters Then
2649               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Cancel", Cancel) & ") "
2650           End If
2651           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2652       End If
' </VB WATCH>
2653       If dgBalanceHoles.SelBookmarks.Count = 0 Then
2654           cmdModifyBalanceHoleData.Visible = False
2655       Else
2656           cmdModifyBalanceHoleData.Visible = True
2657       End If
' <VB WATCH>
2658       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2659       Exit Sub
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
2660       On Error GoTo vbwErrHandler
2661       Const VBWPROCNAME = "frmPLCData.Form_Activate"
2662       If vbwProtector.vbwTraceProc Then
2663           Dim vbwProtectorParameterString As String
2664           If vbwProtector.vbwTraceParameters Then
2665               vbwProtectorParameterString = "()"
2666           End If
2667           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2668       End If
' </VB WATCH>
2669       If ProgramEnd = True Then
2670           Unload Me
2671       End If
' <VB WATCH>
2672       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
2673       Exit Sub
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
2674       On Error GoTo vbwErrHandler
2675       Const VBWPROCNAME = "frmPLCData.Form_Load"
2676       If vbwProtector.vbwTraceProc Then
2677           Dim vbwProtectorParameterString As String
2678           If vbwProtector.vbwTraceParameters Then
2679               vbwProtectorParameterString = "()"
2680           End If
2681           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
2682       End If
' </VB WATCH>
2683       Dim RetVal As String
2684       Dim sSendStr As String
2685       Dim I As Integer
2686       Dim j As Integer
2687       Dim sTableName As String
2688       Dim WhichServer As String
2689       Dim WhichDatabase As String

2690       ProgramEnd = False
2691       Dim objWMIService As Object
2692       Dim colProcesses As Object
2693       Set objWMIService = GetObject("winmgmts:")
2694       Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name LIKE 'PolarRundown%'")
       '    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name LIKE 'Excel%'")
2695       If colProcesses.Count > 1 Then
2696           MsgBox "There is already a copy of Polar Rundown running.  You can only have one copy running at a time", vbOKOnly, "Polar Rundown already running"
2697           Dim f As Form
2698           For Each f In Forms
2699               If f.Name <> Me.Name Then
2700                    Unload f
2701               End If
2702           Next
2703           ProgramEnd = True
' <VB WATCH>
2704       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
2705           Exit Sub
2706       Else
2707       End If
2708       Set objWMIService = Nothing
2709       Set colProcesses = Nothing

2710       debugging = 0   'assume not debugging
2711       WhichServer = "Production"     'change to production server
2712       WhichDatabase = "Production"

2713       If UCase$(Left$(GetMachineName, 5)) = "MROSE" Or UCase$(Left$(GetMachineName, 5)) = "ITTES" Then  'if mickey, see if we want to be in debug
2714           I = MsgBox("Debug?", vbYesNo)
2715           If I = vbYes Then
2716               debugging = 1
2717               WhichServer = "Production"
2718               WhichDatabase = "Production"
2719           Else
2720           End If
2721       End If

2722       If debugging Then
       '        GoTo temp
2723       End If
           'see if the mdb file is where it's supposed to be

2724       Dim developmentDatabase As String
2725       developmentDatabase = GetUNCFromLetter("F:") & sDevelopmentDatabase

2726       If Dir(developmentDatabase) = "" Then
2727           MsgBox "Development.mdb does not exist on F:, Please contact IT.", , "No Development Database"
2728           End
2729       End If

           'get the database info from the new mdb file
2730       Dim cnDevelopment As New ADODB.Connection
2731       Dim qyDevelopment As New ADODB.Command
2732       Dim rsDevelopment As New ADODB.Recordset

2733       On Error GoTo CannotConnect

2734       With cnDevelopment
2735           .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & developmentDatabase & ";Persist Security Info=False; Jet OLEDB:Database Password=Access7277word;"
2736           .ConnectionTimeout = 10
2737           .Open
2738       End With

2739   On Error GoTo vbwErrHandler
2740       GoTo Connected

2741   CannotConnect:
2742       MsgBox "Cannot connect with Development.mdb database.  Please contact IT.", , "Cannot find Connection data."
2743       End

2744   Connected:

           'we're connected, get the data for the Epicor SQL server
2745       qyDevelopment.CommandText = "SELECT * FROM Connections WHERE Connections.WhichServer = '" & WhichServer & "' AND WhichDatabase = '" & WhichDatabase & "'"
2746       qyDevelopment.ActiveConnection = cnDevelopment

2747       rsDevelopment.CursorLocation = adUseClient
2748       rsDevelopment.CursorType = adOpenStatic
2749       rsDevelopment.LockType = adLockOptimistic

2750       On Error GoTo NoServerData

2751       rsDevelopment.Open qyDevelopment

2752   On Error GoTo vbwErrHandler
2753       GoTo GotServerData

2754   NoServerData:

2755       MsgBox "Cannot connect with Development.mdb database.  Please contact IT.", , "Cannot find Connection data."
2756       End

2757   GotServerData:

2758       If rsDevelopment.RecordCount <> 1 Then
2759           GoTo NoServerData
2760       End If

           'construct Epicor connection string
2761       EpicorConnectionString = "Driver={" & rsDevelopment.Fields("ODBCDriver") & "};" & _
                                         "Database=" & rsDevelopment.Fields("DatabaseName") & ";" & _
                                         "Server=" & rsDevelopment.Fields("ServerName") & ";" & _
                                         "UID=" & rsDevelopment.Fields("UserName") & ";" & _
                                         "PWD=" & rsDevelopment.Fields("UserPassword") & ";"


           'make sure we can open the SQL database

2762       On Error GoTo CannotOpenEpicorSQLServer

2763       Dim cnTestEpicor As New ADODB.Connection
2764       cnTestEpicor.ConnectionString = EpicorConnectionString
2765       cnTestEpicor.Open
2766       cnTestEpicor.Close
2767       Set cnTestEpicor = Nothing
2768   On Error GoTo vbwErrHandler

2769       GoTo FoundEpicorSQLServer

2770   CannotOpenEpicorSQLServer:
2771       MsgBox "Cannot connect with the Epicor SQL server specified in Development.mdb.  Please contact IT.", , "Cannot connect with Epicor SQL Server"
2772       End

2773   FoundEpicorSQLServer:
           'get data on rundown database
2774       rsDevelopment.Close
2775       qyDevelopment.CommandText = "SELECT * FROM Connections WHERE Connections.WhichServer = 'PolarRundown'"

2776       On Error GoTo NoRundownDatabase

2777       rsDevelopment.Open qyDevelopment

2778       GoTo FoundRundownDatabase

2779   NoRundownDatabase:
2780       MsgBox "Cannot connect with the Pump Rundown database specified in Development.mdb.  Please contact IT.", , "Cannot connect with Epicor SQL Server"
2781       End

2782   FoundRundownDatabase:
2783       If rsDevelopment.RecordCount <> 1 Then
2784           GoTo NoRundownDatabase
2785           End
2786       End If

2787   temp:

2788       If debugging Then
2789           sDataBaseName = "c:\databases\PolarData.mdb"
2790       Else

2791          sDataBaseName = GetUNCFromLetter("F:") & "\Groups\Shared\databases\PolarData.mdb"

       '        sDataBaseName = rsDevelopment.Fields("ServerName") & rsDevelopment.Fields("DatabaseName")

       '        sDataBaseName = sServerName & "f\groups\shared\databases\PumpData 2k.mdb"
2792       End If

2793       Dim tempFSO As Object
2794       Set tempFSO = CreateObject("Scripting.FileSystemObject")
2795       ParentDirectoryName = tempFSO.getparentfoldername(sDataBaseName)
2796       Set tempFSO = Nothing

           'see if we can open the pump rundown database
2797       On Error GoTo NoRundownDatabase
2798       With cnPumpData
       '        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataBaseName & ";Persist Security Info=False;Jet OLEDB:Database Password=185TitusAve"
2799           .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataBaseName & ";Persist Security Info=False;"
2800           .ConnectionTimeout = 10
2801           .Open
2802       End With
2803   On Error GoTo vbwErrHandler


2804       If debugging = 0 Then
       '        Printer.Orientation = vbPRORLandscape
2805       End If

2806       lblVersion = "Polar Rundown - Version " & App.Major & "." & App.Minor & "." & App.Revision
2807       frmPLCData.Caption = "Polar Rundown"

2808       boFoundPump = False

2809       Me.Show

2810       MSChart1.Plot.Axis(VtChAxisIdX).AxisTitle = "Flow"
2811       MSChart1.Plot.Axis(VtChAxisIdY).AxisTitle = "TDH"
           'MSChart1.Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen = True
           'MSChart1.Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen = True
2812       MSChart1.Plot.UniformAxis = False
2813       MSChart1.Plot.SeriesCollection.Item(1).SeriesMarker.Auto = False
2814       MSChart1.Plot.SeriesCollection.Item(1).Pen.Width = 5
2815       With MSChart1.Plot.SeriesCollection.Item(1).DataPoints.Item(-1).Marker
2816           .Visible = True
2817           .Size = 50
2818           .Style = VtMarkerStyleCircle
2819           .FillColor.Automatic = False
2820           .FillColor.Set 0, 0, 255
2821       End With

           'assure that the timers are off
2822       frmPLCData.tmrGetDDE.Enabled = False

2823       frmPLCData.tmrStartUp.Enabled = False

           'initialize the PLC network
2824       RetVal = NetWorkInitialize()
2825       If RetVal <> 0 Then
2826           MsgBox ("Can't Initialize Network. Exiting...")
2827           End
2828       End If

2829       If debugging = 0 Then
               'load array of plcs
2830           I = 0
2831           Open rsDevelopment.Fields("ServerName") & "PolarPLCAddresses.txt" For Input As 1
2832           While Not EOF(1)
2833               Input #1, Description(I)
2834               For j = 0 To 125
2835                   Input #1, aDevices(I).Address(j)
2836               Next j
2837               Input #1, j
2838               I = I + 1
2839           Wend
2840           Close #1

2841           DeviceCount = I

2842           If Left$(GetMachineName, 2) = "WV" Then  'if in WV, put MWSC first in loop dropdown
2843               Dim k As Integer
2844               For k = 0 To DeviceCount - 1
2845                   If InStr(Description(k), "MWSC") <> 0 Then
2846                       Exit For
2847                   End If
2848               Next k
2849               Description(DeviceCount) = Description(0)
2850               Description(0) = Description(k)
2851               Description(k) = Description(DeviceCount)

2852               aDevices(DeviceCount) = aDevices(0)
2853               aDevices(0) = aDevices(k)
2854               aDevices(k) = aDevices(DeviceCount)

2855           End If

2856           Dim PLCAddress As String
2857           For I = 0 To DeviceCount - 1
2858               PLCAddress = aDevices(I).Address(4) & "." & aDevices(I).Address(5) & "." & aDevices(I).Address(6) & "." & aDevices(I).Address(7)
2859               RetVal = PingSilent(PLCAddress)
2860               If RetVal <> 0 Then
2861                   frmPLCData.cmbPLCLoop.AddItem Description(I)
2862                   frmPLCData.cmbPLCLoop.ItemData(frmPLCData.cmbPLCLoop.NewIndex) = I
2863               End If
2864           Next I
2865       End If

2866       frmPLCData.cmbPLCLoop.AddItem "Add PLC Data Manually"   'enable the controls for manual entry

           'turn on the PLC led

2867       frmPLCData.cmbPLCLoop.ListIndex = 0
2868       frmPLCData.tmrGetDDE.Enabled = True

           'hook up to the various databases

           'copy the template of the database here
           'see if it exists
2869       Dim fdrive As String
2870       fdrive = GetUNCFromLetter("F:")
2871       If Dir(fdrive & "\groups\shared\databases" & sEffDataBaseName) = "" Then
2872           MsgBox "File does not exist at " & fdrive & "\groups\shared\databases" & sEffDataBaseName & ". Please contact IT", vbOKOnly, "Eff.mdb does not exist"
2873       Else
               'Dim FSO As New FileSystemObject
2874           FileCopy fdrive & "\groups\shared\databases" & sEffDataBaseName, App.Path & sEffDataBaseName
2875       End If


2876       With cnEffData
2877           .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & sEffDataBaseName & ";Persist Security Info=False"
2878           .Open
2879       End With

           'open some recordsets
2880       rsPumpData.Index = "SerialNumber"
2881       rsTestSetup.Index = "FindData"
2882       rsTestData.Index = "PrimaryKey"
2883       rsPumpData.Open "TempPumpData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
2884       rsTestSetup.Open "TempTestSetupData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
2885       rsTestData.Filter = "SerialNumber = ''"
2886       rsTestData.CursorLocation = adUseClient
2887       rsTestData.Open "TempTestData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
2888       rsEff.CursorLocation = adUseClient
2889       rsEff.Open "Efficiency", cnEffData, adOpenStatic, adLockOptimistic, adCmdTableDirect
2890       qyBalanceHoles.ActiveConnection = cnPumpData
2891       rsBalanceHoles.CursorLocation = adUseClient
2892       rsBalanceHoles.CursorType = adOpenStatic
2893       rsBalanceHoles.LockType = adLockOptimistic
2894       qyMisc.ActiveConnection = cnPumpData
2895       qyMisc.CommandText = "SELECT MiscParameters.ParameterName, MiscParameters.ParameterValue From MiscParameters WHERE (((MiscParameters.ParameterName)='AllowableTDHVariation'));"
2896       rsMisc.CursorLocation = adUseClient
2897       rsMisc.CursorType = adOpenStatic
2898       rsMisc.LockType = adLockBatchOptimistic
2899       rsMisc.Open qyMisc
2900       txtNPSH(5).Text = rsMisc!ParameterValue

2901       If debugging <> 1 Then
2902           FindMagtrols
2903       Else
2904           cmbMagtrol.AddItem "Add Manually"
2905           cmbMagtrol.ItemData(cmbMagtrol.NewIndex) = 99
2906           cmbMagtrol.ListIndex = 0
2907       End If
2908       optKW(1).value = True
2909       optKW_Click (1)


           'blank out data grid
2910       Set DataGrid1.DataSource = rsTestData

           'load the combo boxes
2911       LoadCombo cmbStatorFill, "StatorFill"
2912       LoadCombo cmbCirculationPath, "CirculationPath"
2913       LoadCombo cmbVoltage, "Voltage"
2914       LoadCombo cmbFrequency, "Frequency"
2915       LoadCombo cmbMotor, "Motor"
2916       LoadCombo cmbDesignPressure, "DesignPressure"
2917       LoadCombo cmbRPM, "RPM"
2918       LoadCombo cmbOrificeNumber, "OrificeNumber"
2919       LoadCombo cmbTestSpec, "TestSpecification"
2920       LoadCombo cmbLoopNumber, "LoopNumber"
2921       LoadCombo cmbSuctDia, "SuctionDiameter"
2922       LoadCombo cmbDischDia, "DischargeDiameter"
2923       LoadCombo cmbTachID, "TachID"
2924       LoadCombo cmbAnalyzerNo, "AnalyzerNo"
2925       LoadCombo cmbModel, "Model"
2926       LoadCombo cmbModelGroup, "ModelGroup"
2927       LoadCombo cmbMounting, "Mounting"
2928       LoadCombo cmbPLCNo, "PLCNo"
2929       LoadCombo cmbFlowMeter, "PumpFlowMeter"
2930       LoadCombo cmbSuctionPressureTransducer, "SuctionPressureTransducer"
2931       LoadCombo cmbDischargePressureTransducer, "DischargePressureTransducer"
2932       LoadCombo cmbTemperatureTransducer, "TemperatureTransducer"
2933       LoadCombo cmbCirculationFlowMeter, "CirculationFlowMeter"
           'LoadCombo cmbSupermarketModel, "SupermarketPumpData"

           'load the TEMC combo boxes, too
2934       LoadCombo cmbTEMCAdapter, "TEMCAdapter"
2935       LoadCombo cmbTEMCAdditions, "TEMCAdditions"
2936       LoadCombo cmbTEMCCirculation, "TEMCCirculation"
2937       LoadCombo cmbTEMCDesignPressure, "TEMCDesignPressure"
2938       LoadCombo cmbTEMCNominalDischargeSize, "TEMCNominalDischargeSize"
2939       LoadCombo cmbTEMCDivisionType, "TEMCDivisionType"
2940       LoadCombo cmbTEMCImpellerType, "TEMCImpellerType"
2941       LoadCombo cmbTEMCInsulation, "TEMCInsulation"
2942       LoadCombo cmbTEMCJacketGasket, "TEMCJacketGasket"
2943       LoadCombo cmbTEMCMaterials, "TEMCMaterials"
2944       LoadCombo cmbTEMCModel, "TEMCModel"
2945       LoadCombo cmbTEMCNominalImpSize, "TEMCNominalImpSize"
2946       LoadCombo cmbTEMCOtherMotor, "TEMCOtherMotor"
2947       LoadCombo cmbTEMCNominalSuctionSize, "TEMCNominalSuctionSize"
2948       LoadCombo cmbTEMCVoltage, "TEMCVoltage"
2949       LoadCombo cmbTEMCPumpStages, "TEMCPumpStages"
2950       LoadCombo cmbTEMCTRG, "TEMCTRG"

           'LoadCombo frmSearch.cmbSearchModel, "Model"

           'fill memory arrays for dlookups
2951       FillArrays

           'choose the first tab
2952       frmPLCData.SSTab1.Tab = 0

           'set the grid column names
2953       Dim c As Column
2954       For Each c In DataGrid1.Columns
2955           Select Case c.DataField
               Case "TestDataID"
2956               c.Visible = False
2957           Case "SerialNumber"
2958               c.Visible = False
2959           Case "Date"
2960               c.Visible = False
2961           Case Else ' Show all other columns.
2962               c.Visible = True
2963               c.Alignment = dbgRight
2964           End Select
2965       Next c

2966       Set dgBalanceHoles.DataSource = rsBalanceHoles

2967       For Each c In dgBalanceHoles.Columns
2968           Select Case c.DataField
               Case "BalanceHoleID"
2969               c.Visible = False
2970           Case "SerialNo"
2971               c.Visible = False
2972           Case "Date"
2973               c.Visible = True
2974               c.Alignment = dbgCenter
2975               c.Width = 2000
2976           Case "Number"
2977               c.Visible = True
2978               c.Alignment = dbgCenter
2979               c.Width = 700
2980           Case "Diameter"
2981               c.Visible = False
2982           Case "Diameter1"
2983               c.Caption = "Diameter"
2984               c.Visible = True
2985               c.Alignment = dbgCenter
2986               c.Width = 700
2987           Case "BoltCircle1"
2988               c.Caption = "Bolt Circle"
2989               c.Visible = True
2990               c.Alignment = dbgCenter
2991               c.Width = 800
2992           Case "BoltCircle"
2993               c.Visible = False
2994           Case "SetNo"
2995               c.Visible = False
2996           Case Else ' Show all other columns.
2997               c.Visible = False
2998           End Select
2999       Next c

3000       BlankData

       '    If debugging <> 1 Then
               'get user initials
3001           frmLogin.Show
       '    End If

3002     optMfr(1).value = True
3003     frmMfr.Visible = False

3004       Pressed = True
' <VB WATCH>
3005       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3006       Exit Sub
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
3007       On Error GoTo vbwErrHandler
3008       Const VBWPROCNAME = "frmPLCData.Form_Unload"
3009       If vbwProtector.vbwTraceProc Then
3010           Dim vbwProtectorParameterString As String
3011           If vbwProtector.vbwTraceParameters Then
3012               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Cancel", Cancel) & ") "
3013           End If
3014           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3015       End If
' </VB WATCH>
3016       End
' <VB WATCH>
3017       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3018       Exit Sub
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
3019       On Error GoTo vbwErrHandler
3020       Const VBWPROCNAME = "frmPLCData.Label15_Click"
3021       If vbwProtector.vbwTraceProc Then
3022           Dim vbwProtectorParameterString As String
3023           If vbwProtector.vbwTraceParameters Then
3024               vbwProtectorParameterString = "()"
3025           End If
3026           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3027       End If
' </VB WATCH>
3028       frmDiagram.Show
' <VB WATCH>
3029       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3030       Exit Sub
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
3031       On Error GoTo vbwErrHandler
3032       Const VBWPROCNAME = "frmPLCData.lblAutoMan_Click"
3033       If vbwProtector.vbwTraceProc Then
3034           Dim vbwProtectorParameterString As String
3035           If vbwProtector.vbwTraceParameters Then
3036               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3037           End If
3038           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3039       End If
' </VB WATCH>

3040       Dim blnEnabled As Boolean

3041       If lblAutoMan(Index).Caption = "Auto" Then
3042           lblAutoMan(Index).Caption = "Man"
3043           blnEnabled = True
3044       Else
3045           lblAutoMan(Index).Caption = "Auto"
3046           blnEnabled = False
3047       End If

3048       Select Case Index
               Case 0
3049               txtFlowDisplay.Enabled = blnEnabled
3050           Case 1
3051               txtSuctionDisplay.Enabled = blnEnabled
3052           Case 2
3053               txtDischargeDisplay.Enabled = blnEnabled
3054           Case 3
3055               txtTemperatureDisplay.Enabled = blnEnabled
3056           Case 4
3057               txtAI1Display.Enabled = blnEnabled
3058           Case 5
3059               txtAI2Display.Enabled = blnEnabled
3060           Case 6
3061               txtAI3Display.Enabled = blnEnabled
3062           Case 7
3063               txtAI4Display.Enabled = blnEnabled
3064       End Select

' <VB WATCH>
3065       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3066       Exit Sub
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
3067       On Error GoTo vbwErrHandler
3068       Const VBWPROCNAME = "frmPLCData.tmrNPSHr_Timer"
3069       If vbwProtector.vbwTraceProc Then
3070           Dim vbwProtectorParameterString As String
3071           If vbwProtector.vbwTraceParameters Then
3072               vbwProtectorParameterString = "()"
3073           End If
3074           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3075       End If
' </VB WATCH>
3076       tmrNPSHr.Enabled = False
3077       If frmNPSH.Visible = True Then
3078           btnRunNPSH_Click    'close test
3079       End If
' <VB WATCH>
3080       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3081       Exit Sub
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
3082       On Error GoTo vbwErrHandler
3083       Const VBWPROCNAME = "frmPLCData.txtNPSH_Change"
3084       If vbwProtector.vbwTraceProc Then
3085           Dim vbwProtectorParameterString As String
3086           If vbwProtector.vbwTraceParameters Then
3087               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3088           End If
3089           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3090       End If
' </VB WATCH>
3091       If Index = 5 Then
3092           If frmNPSH.Visible = True Then
3093               If rsMisc.State = adStateOpen Then
3094                   rsMisc.Close
3095               End If
3096               rsMisc.CursorLocation = adUseClient
3097               rsMisc.Open "Select * from MiscParameters WHERE (ParameterName = 'AllowableTDHVariation');", cnPumpData, adOpenStatic, adLockOptimistic, adCmdText
3098               rsMisc.Fields("ParameterValue").value = txtNPSH(5).Text
3099               rsMisc.Update
3100           End If
3101       End If
' <VB WATCH>
3102       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3103       Exit Sub
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
3104       On Error GoTo vbwErrHandler
3105       Const VBWPROCNAME = "frmPLCData.txtNPSHFileLocation_Click"
3106       If vbwProtector.vbwTraceProc Then
3107           Dim vbwProtectorParameterString As String
3108           If vbwProtector.vbwTraceParameters Then
3109               vbwProtectorParameterString = "()"
3110           End If
3111           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3112       End If
' </VB WATCH>
3113       Dim sTempDir As String
3114       On Error Resume Next
3115       sTempDir = CurDir    'Remember the current active directory
3116       CommonDialog2.DialogTitle = "Select a directory" 'titlebar
3117       CommonDialog2.InitDir = "\\tei-main-01\f\en\groups\shared\calibration and rundown\npsh\" 'start dir, might be "C:\" or so also
3118       CommonDialog2.filename = "Select a Directory"  'Something in filenamebox
3119       CommonDialog2.Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
3120       CommonDialog2.Filter = "Directories|*.~#~" 'set files-filter to show dirs only
3121       CommonDialog2.CancelError = True 'allow escape key/cancel
3122       CommonDialog2.ShowSave   'show the dialog screen

3123       If Err <> 32755 Then    ' User didn't chose Cancel.
               'Me.SDir.Text = CurDir
3124       End If

       '    ChDir sTempDir  'restore path to what it was at entering

3125   Me.txtNPSHFileLocation.Text = CommonDialog2.filename

' <VB WATCH>
3126       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3127       Exit Sub
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
3128       On Error GoTo vbwErrHandler
3129       Const VBWPROCNAME = "frmPLCData.txtTitle_LostFocus"
3130       If vbwProtector.vbwTraceProc Then
3131           Dim vbwProtectorParameterString As String
3132           If vbwProtector.vbwTraceParameters Then
3133               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3134           End If
3135           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3136       End If
' </VB WATCH>

3137       ChangeTitles Index

' <VB WATCH>
3138       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3139       Exit Sub
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
3140       On Error GoTo vbwErrHandler
3141       Const VBWPROCNAME = "frmPLCData.ChangeTitles"
3142       If vbwProtector.vbwTraceProc Then
3143           Dim vbwProtectorParameterString As String
3144           If vbwProtector.vbwTraceParameters Then
3145               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("ChannelNo", ChannelNo) & ") "
3146           End If
3147           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3148       End If
' </VB WATCH>
3149       Dim I As Integer
3150       Dim S As String

3151       If txtTitle(ChannelNo).Locked = True Then
' <VB WATCH>
3152       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3153           Exit Sub
3154       End If

3155       Dim qy As New ADODB.Command
3156       Dim rs As New ADODB.Recordset

3157       qy.ActiveConnection = cnPumpData

           'see if we have an entry in the table
3158       qy.CommandText = "SELECT * FROM AITitles " & _
                             "WHERE (((AITitles.SerialNo)= '" & txtSN.Text & "') " & _
                             "AND ((AITitles.Date)= #" & cmbTestDate.Text & "#) " & _
                             "AND ((AITitles.Channel)=" & ChannelNo & "));"

3159       With rs     'open the recordset for the query
3160           .CursorLocation = adUseClient
3161           .CursorType = adOpenStatic
3162           .LockType = adLockOptimistic
3163           .Open qy
3164       End With

3165       If (rs.BOF = True And rs.EOF = True) Then  'new record
3166           rs.AddNew
3167           rs.Fields("SerialNo") = txtSN.Text
3168           rs.Fields("Date") = cmbTestDate.Text
3169           rs.Fields("Channel") = CByte(ChannelNo)
3170           rs.Fields("Title") = txtTitle(ChannelNo).Text
3171           rs.Update
3172       Else    'we have an entry, modify it
3173           rs.Fields("SerialNo") = txtSN.Text
3174           rs.Fields("Date") = cmbTestDate.Text
3175           rs.Fields("Channel") = CByte(ChannelNo)
3176           rs.Fields("Title") = txtTitle(ChannelNo).Text
3177           rs.Update
3178       End If

3179       rs.Close
3180       Set rs = Nothing
3181       Set qy = Nothing

' <VB WATCH>
3182       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3183       Exit Sub
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
3184       On Error GoTo vbwErrHandler
3185       Const VBWPROCNAME = "frmPLCData.optKW_Click"
3186       If vbwProtector.vbwTraceProc Then
3187           Dim vbwProtectorParameterString As String
3188           If vbwProtector.vbwTraceParameters Then
3189               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3190           End If
3191           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3192       End If
' </VB WATCH>
3193       Select Case Index
               Case 0  'add 3 powers
3194               txtKW.Enabled = False
3195           Case 1  'enter kw
3196               txtKW.Enabled = True
3197           Case 2  'use analog in 4
3198               txtKW.Enabled = False
3199       End Select
' <VB WATCH>
3200       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3201       Exit Sub
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
3202       On Error GoTo vbwErrHandler
3203       Const VBWPROCNAME = "frmPLCData.optMfr_Click"
3204       If vbwProtector.vbwTraceProc Then
3205           Dim vbwProtectorParameterString As String
3206           If vbwProtector.vbwTraceParameters Then
3207               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Index", Index) & ") "
3208           End If
3209           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3210       End If
' </VB WATCH>
3211       frmTEMC.Visible = optMfr(1).value
3212       frmChempump.Visible = optMfr(0).value
3213       frmTEMCData.Visible = optMfr(1).value
3214       txtModelNo_Change
' <VB WATCH>
3215       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3216       Exit Sub
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
3217       On Error GoTo vbwErrHandler
3218       Const VBWPROCNAME = "frmPLCData.tmrGetDDE_Timer"
3219       If vbwProtector.vbwTraceProc Then
3220           Dim vbwProtectorParameterString As String
3221           If vbwProtector.vbwTraceParameters Then
3222               vbwProtectorParameterString = "()"
3223           End If
3224           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3225       End If
' </VB WATCH>

       'get here every second... get plc and magtrol data

3226       Dim sSendStr As String
3227       Dim I As Integer
3228       Dim VoltMul As Double

3229       If Calibrating Then
' <VB WATCH>
3230       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3231           Exit Sub
3232       End If

3233       If debugging Then
               'Exit Sub
3234       End If


3235       If boPLCOperating = True Then
3236           frmPLCData.shpGetPLCData.Visible = True    'turn the PLC led on

               'convert the plc data into real numbers
               'the following data are type real
3237           txtFlow.Text = ConvertToReal("4050")
3238           txtSuction.Text = ConvertToReal("4052")
3239           txtDischarge.Text = ConvertToReal("4054")
3240           txtTemperature.Text = ConvertToReal("4056")

3241           txtValvePosition.Text = ConvertToLong("2004")

3242           frmPLCData.txtTC1.Text = ConvertToLong("2200")
3243           frmPLCData.txtTC2.Text = ConvertToLong("2202")
3244           frmPLCData.txtTC3.Text = ConvertToLong("2204")
3245           frmPLCData.txtTC4.Text = ConvertToLong("2206")

3246           frmPLCData.txtAI1.Text = ConvertToReal("4060")
3247           frmPLCData.txtAI2.Text = ConvertToReal("4062")
3248           frmPLCData.txtAI3.Text = ConvertToReal("4064")
3249           frmPLCData.txtAI4.Text = ConvertToReal("4066")

3250           frmPLCData.txtPCoef.Text = ConvertToLong("4036")
3251           frmPLCData.txtICoef.Text = ConvertToLong("4037")
3252           frmPLCData.txtDCoef.Text = ConvertToLong("4040")

3253           frmPLCData.txtSetPoint.Text = ConvertToLong("4035")
3254           frmPLCData.txtInHg.Text = ConvertToLong("1460")


               'modify the data from PLC format to format that we can use
               'and update the screen
3255           If txtFlowDisplay.Enabled = False Then
3256               frmPLCData.txtFlowDisplay = Format$(txtFlow.Text, "###0.00")
3257           End If
3258           If txtSuctionDisplay.Enabled = False Then
3259               frmPLCData.txtSuctionDisplay = Format$((txtSuction.Text) / 10, "##0.00")
3260           End If
3261           If txtDischargeDisplay.Enabled = False Then
3262               frmPLCData.txtDischargeDisplay = Format$(txtDischarge.Text, "##0.00")
3263           End If
3264           If txtTemperatureDisplay.Enabled = False Then
3265               frmPLCData.txtTemperatureDisplay = Format$(txtTemperature.Text, "##0.00")
3266           End If
3267           frmPLCData.txtValvePositionDisplay = (txtValvePosition.Text)

3268           frmPLCData.txtTC1Display = Format$((txtTC1.Text) / 10, "##0.0")
3269           frmPLCData.txtTC2Display = Format$((txtTC2.Text) / 10, "##0.0")
3270           frmPLCData.txtTC3Display = Format$((txtTC3.Text) / 10, "##0.0")
3271           frmPLCData.txtTC4Display = Format$((txtTC4.Text) / 10, "##0.0")

3272           If txtAI1Display.Enabled = False Then
3273               frmPLCData.txtAI1Display = Format$(txtAI1.Text, "##0.00")
3274           End If
3275           If txtAI2Display.Enabled = False Then
3276               frmPLCData.txtAI2Display = Format$(txtAI2.Text, "##0.00")
3277           End If
3278           If txtAI3Display.Enabled = False Then
3279               frmPLCData.txtAI3Display = Format$(txtAI3.Text, "##0.00")
3280           End If
3281           If txtAI4Display.Enabled = False Then
3282               frmPLCData.txtAI4Display = Format$(txtAI4.Text, "##0.00")
3283           End If

3284           frmPLCData.txtSetPointDisplay = (txtSetPoint.Text)

3285           frmPLCData.txtInHgDisplay = Format$(txtInHg.Text / 100, "00.00")

3286           frmPLCData.shpGetPLCData.Visible = False   'turn the PLC led off

3287           frmPLCData.shpGetMagtrolData.Visible = True 'turn the Magtrol led on
3288       End If

3289       If boMagtrolOperating = True Then


               'get the data from the Magtrol
3290           If Right(cmbMagtrol.List(cmbMagtrol.ListIndex), 4) = "5300" Then
3291               sSendStr = vbCrLf
3292               sData = Space$(68)
3293               VoltMul = Sqr(3)
3294           Else
3295               sSendStr = "OT" & vbCrLf
3296               sData = Space$(183)
3297               VoltMul = 1#
3298           End If

3299           On Error GoTo noresponse
3300           If UsingNatInst Then
3301               ibwrt iUD, sSendStr
3302               ibrd iUD, sData

                   'parse the Magrol response
       '            vResponse = CWGPIB1.Tasks("Number Parser").Parse(sData)
3303           Else
                   'Dim Databack As String
3304               sData = TCP.SendGetData("OT")
3305           End If

3306               Dim vSplit() As String
3307               vSplit = Split(Right(sData, Len(sData) - 1), ",")
3308               ReDim vResponse(UBound(vSplit))
3309               For I = 0 To UBound(vSplit) - 1
3310                   vResponse(I) = CDbl(vSplit(I))
3311               Next I

               'format the parsed response
3312           Dim dd As String
3313           dd = "- -"

3314           If Not IsEmpty(vResponse) Then
               '8 entries for 5300 and 12 for the 6530
3315               If UBound(vResponse) = 8 Or UBound(vResponse) = 12 Then
                       'put the responses into the correct text box
3316                   txtV1.Text = Format$(VoltMul * vResponse(1), "###0.0")   'we get back phase voltage and we want line voltage

3317                   Select Case vResponse(0)
                           Case Is < 1
3318                           txtI1.Text = Format$(vResponse(0), "0.0000")
3319                       Case Is < 10
3320                           txtI1.Text = Format$(vResponse(0), "0.000")
3321                       Case Is < 100
3322                           txtI1.Text = Format$(vResponse(0), "00.00")
3323                       Case Else
3324                           txtI1.Text = Format$(vResponse(0), "000.0")
3325                   End Select

3326                   Select Case vResponse(3)
                           Case Is < 1
3327                           txtI2.Text = Format$(vResponse(3), "0.0000")
3328                       Case Is < 10
3329                           txtI2.Text = Format$(vResponse(3), "0.000")
3330                       Case Is < 100
3331                           txtI2.Text = Format$(vResponse(3), "00.00")
3332                       Case Else
3333                           txtI2.Text = Format$(vResponse(3), "000.0")
3334                   End Select

3335                   Select Case vResponse(6)
                           Case Is < 1
3336                           txtI3.Text = Format$(vResponse(6), "0.0000")
3337                       Case Is < 10
3338                           txtI3.Text = Format$(vResponse(6), "0.000")
3339                       Case Is < 100
3340                           txtI3.Text = Format$(vResponse(6), "00.00")
3341                       Case Else
3342                           txtI3.Text = Format$(vResponse(6), "000.0")
3343                   End Select

3344                   txtP1.Text = Format$(vResponse(2) / 1000, "##0.00")     '/ by 1000 to show kW
3345                   txtV2.Text = Format$(VoltMul * vResponse(4), "###0.0")
                       'txtI2.Text = Format$(vResponse(3), "###0.0")
3346                   txtP2.Text = Format$(vResponse(5) / 1000, "##0.00")
3347                   txtV3.Text = Format$(VoltMul * vResponse(7), "###0.0")
                       'txtI3.Text = Format$(vResponse(6), "###0.0")
3348                   txtP3.Text = Format$(vResponse(8) / 1000, "##0.00")
3349                   If (vResponse(0) * vResponse(1) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7)) <> 0 Then
                           'if we have some measured current
                           'pf = sum of power/sum of VA
3350                       If Right(cmbMagtrol.List(cmbMagtrol.ListIndex), 4) = "5300" Then
                               'add kw responses and / by 1000 to get to kW
3351                           txtKW.Text = (vResponse(2) + vResponse(5) + vResponse(8)) / 1000
3352                           txtPF.Text = Format$(100 * (vResponse(2) + vResponse(5) + vResponse(8)) / (vResponse(1) * vResponse(0) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7)), "0.00")
3353                       Else
3354                           txtKW.Text = (vResponse(2) + vResponse(8)) / 1000
3355                           txtPF.Text = Format$(100 * (vResponse(2) + vResponse(8)) / ((Sqr(3) / 3) * (vResponse(1) * vResponse(0) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7))), "0.00")
3356                       End If
3357                       Select Case Val(txtKW.Text)
                               Case Is < 1
3358                               txtKW.Text = Format$(txtKW.Text, "0.00000")
3359                           Case Is < 10
3360                               txtKW.Text = Format$(txtKW.Text, "0.0000")
3361                           Case Is < 100
3362                               txtKW.Text = Format$(txtKW.Text, "00.000")
3363                           Case Else
3364                               txtKW.Text = Format$(txtKW.Text, "000.00")
3365                       End Select
3366                   Else
3367                       txtPF = dd
3368                   End If
3369               Else
                       'no response, show all -- in text boxes
3370                   txtV1.Text = dd
3371                   txtI1.Text = dd
3372                   txtP1.Text = dd
3373                   txtV2.Text = dd
3374                   txtI2.Text = dd
3375                   txtP2.Text = dd
3376                   txtV3.Text = dd
3377                   txtI3.Text = dd
3378                   txtP3.Text = dd
3379                   txtPF = dd
3380                   txtKW = dd
3381               End If
3382           End If
3383       Else    'magtrol not operating
3384           Dim dbl As Double

3385           If optKW(0).value = True Then   'add 3 powers
3386               txtKW.Text = Val(txtP1.Text) + Val(txtP2.Text) + Val(txtP3.Text)
3387           End If
3388           If optKW(1).value = True Then   'enter kw
3389               txtP1.Text = Val(txtKW.Text) / 3
3390               txtP2.Text = Val(txtKW.Text) / 3
3391               txtP3.Text = Val(txtKW.Text) / 3
3392           End If
3393           If optKW(2).value = True Then   'use ai4
3394               txtKW.Text = txtAI4Display.Text
3395               txtP1.Text = Val(txtKW.Text) / 3
3396               txtP2.Text = Val(txtKW.Text) / 3
3397               txtP3.Text = Val(txtKW.Text) / 3
3398           End If

3399           dbl = Val(txtV1.Text) * Val(txtI1.Text)
3400           dbl = dbl + Val(txtV2.Text) * Val(txtI2.Text)
3401           dbl = dbl + Val(txtV3.Text) * Val(txtI3.Text)
3402           If dbl <> 0 Then
3403               txtPF.Text = Format$((Val(txtKW.Text) * 1000 * 3 * 100 / (dbl * Sqr(3))), "0.00")
3404           End If
3405       End If

3406   noresponse:
3407   On Error GoTo vbwErrHandler
3408       frmPLCData.shpGetMagtrolData.Visible = False   'turn the Magtrol led off

           'update the little PLC chart
3409       For I = 1 To 99
3410           vPlot(0, I) = vPlot(0, I + 1)
3411           vPlot(1, I) = vPlot(1, I + 1)
3412       Next I
3413       vPlot(0, 100) = txtSetPointDisplay
3414       vPlot(1, 100) = txtFlowDisplay

           'do NPSH stuff
3415       Dim SuctVelHead As Single
3416       Dim DischVelHead As Single
3417       Dim Conversion As Single
3418       Dim SuctionPSIA As Single
3419       Dim DischargePSIA As Single
3420       Dim VaporPress As Single
3421       Dim SpecVolume As Single
3422       Dim NPSHa As Single
3423       Dim NPSHr As Single
3424       Dim TDH As Single
3425       Dim pd As Single


           'velocity head
3426       If cmbSuctDia.ListIndex = -1 Then   'if no suction diameter chosen
3427           SuctVelHead = 0
3428       Else
       '        pd = DLookup("ActualDia", "PipeDiameters", "ID = " & cmbSuctDia.ListIndex + 1)
3429           pd = DLookupA(ActualColNo, PipeDiameters, IDColNo, cmbSuctDia.ItemData(cmbSuctDia.ListIndex) + 1)
3430           SuctVelHead = (0.002592 * Val(txtFlow) ^ 2) / (pd ^ 4)
3431       End If

3432       If cmbDischDia.ListIndex = -1 Then     'if no discharge diameter chosen
3433           DischVelHead = 0
3434       Else
       '        pd = DLookup("ActualDia", "PipeDiameters", "ID = " & cmbDischDia.ListIndex + 1)
3435           pd = DLookupA(ActualColNo, PipeDiameters, IDColNo, cmbDischDia.ItemData(cmbDischDia.ListIndex) + 1)
3436           DischVelHead = (0.002592 * Val(txtFlow) ^ 2) / (pd ^ 4)
3437       End If

           'convert gauges to absolute
3438       If txtInHgDisplay.Text = "" Then
3439           Conversion = 0
3440       Else
3441           Conversion = txtInHgDisplay * 0.491
3442       End If

3443       SuctionPSIA = Val(txtSuctionDisplay) + Conversion
3444       DischargePSIA = Val(txtDischargeDisplay) + Conversion


           'lookup vapor pressure and specific volume in the arrays that we made
           'if temp is out of range, say so and exit
3445       If Val(txtTemperatureDisplay) < 40 Or Val(txtTemperatureDisplay) > 165 Then
3446           txtNPSHa = 0
' <VB WATCH>
3447       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3448           Exit Sub
3449       Else
3450           I = Val(txtTemperatureDisplay) - 40
       '        VaporPress = DLookup("VaporPressure", "VaporPressure", "ID = " & I)
       '        SpecVolume = DLookup("SpecificVolume", "VaporPressure", "ID= " & I)
3451           VaporPress = DLookupA(VaporPressureColNo, VaporPressure, IDColNo, I)
3452           SpecVolume = DLookupA(SpecificVolumeColNo, VaporPressure, IDColNo, I)
3453       End If

3454       If Not ((txtSuctHeight = "") Or (txtDischHeight = "") Or Not IsNumeric(txtSuctHeight) Or Not IsNumeric(txtDischHeight)) Then
               'NPSHa
3455           NPSHa = (144 * SpecVolume * (SuctionPSIA - VaporPress)) + (txtSuctHeight / 12) + SuctVelHead
       '        NPSHa = CalcTDH(DischargePSIA, SuctionPSIA, 0, DischVelHead, 0, txtTemperature)
3456           txtNPSHa = Format$(NPSHa, "##0.00")

               'tdh
3457           TDH = CalcTDH(DischargePSIA, SuctionPSIA, 0, (DischVelHead - SuctVelHead), (txtDischHeight / 12) - (txtSuctHeight / 12), txtTemperatureDisplay)
3458           txtTDH = Format$(TDH, "##0.00")

3459           If frmNPSH.Visible = True Then
3460               If Val(txtTDH.Text) > 0 Then
3461                   txtNPSH(2).Text = Format(100 * Val(txtTDH.Text) / Val(txtNPSH(3).Text), "##0.00")
3462                   txtNPSH(1).Text = Format(100 * Val(txtFlow.Text) / Val(txtNPSH(0).Text), "##0.00")
                       'check for tdh variation
3463                   If Abs(Val(txtNPSH(1)) - 100) > Val(txtNPSH(5).Text) Then
3464                       MsgBox "The TDH value has varied more than " & txtNPSH(5) & " %. NPSHr data will NOT be written to the data table", vbOKOnly, "TDH variation too large"
3465                       btnRunNPSH_Click
3466                   Else    'tdh variation small
3467                       If Val(txtNPSH(2).Text) <= 97 Then
                               'btnRunNPSH_Click
                               'write the npsh and save
3468                           If WroteNPSHr = False Then
3469                               txtNPSH(4).Text = txtNPSHa.Text
3470                               rsTestData!NPSHr = txtNPSHa.Text
3471                               rsTestData.Update
3472                               rsEff!NPSHr = txtNPSHa.Text
3473                               rsEff.Update
3474                               WroteNPSHr = True
3475                               tmrNPSHr.Interval = 5000
3476                               tmrNPSHr.Enabled = True
3477                           End If
3478                       End If  'val < 97
3479                   End If  'check for tdh variation
3480               End If 'val tdh <=0
3481           Else    'frm not visible
                   'txtNPSHa = Format$(0, "##0.00")
3482           End If  'if frm visible

3483       Else
3484           txtNPSHa = 0
3485       End If
' <VB WATCH>
3486       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3487       Exit Sub
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
3488       On Error GoTo vbwErrHandler
3489       Const VBWPROCNAME = "frmPLCData.tmrStartUp_Timer"
3490       If vbwProtector.vbwTraceProc Then
3491           Dim vbwProtectorParameterString As String
3492           If vbwProtector.vbwTraceParameters Then
3493               vbwProtectorParameterString = "()"
3494           End If
3495           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3496       End If
' </VB WATCH>
3497       tmrStartUp.Enabled = False
' <VB WATCH>
3498       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3499       Exit Sub
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
3500       On Error GoTo vbwErrHandler
3501       Const VBWPROCNAME = "frmPLCData.SetCombo"
3502       If vbwProtector.vbwTraceProc Then
3503           Dim vbwProtectorParameterString As String
3504           If vbwProtector.vbwTraceParameters Then
3505               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
3506               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sName", sName) & ", "
3507               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("rs", rs) & ") "
3508           End If
3509           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3510       End If
' </VB WATCH>

3511       Dim I As Integer
3512       Dim sParam As String
3513       Dim qy As New ADODB.Command
3514       Dim rs1 As New ADODB.Recordset

3515       If rs.Fields(sName).ActualSize <> 0 Then     'if there's an entry
3516           sParam = rs.Fields(sName)                'get the index number
3517           qy.ActiveConnection = cnPumpData
3518           qy.CommandText = "SELECT * FROM " & sName & " WHERE " & sName & " = " & sParam
3519           Set rs1 = qy.Execute()                                  'get the record for the index number

3520           If rs1.BOF = True And rs1.EOF = True Then
3521               cmbComboName.ListIndex = -1                             'else, remove any pointer
' <VB WATCH>
3522       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3523               Exit Function
3524           End If

3525           For I = 0 To cmbComboName.ListCount - 1                     'go through the combobox entries
3526               If cmbComboName.ItemData(I) = rs1.Fields(0) Then     'see when we find the desired index number
3527                   cmbComboName.ListIndex = I                                              'if we do, set the combo box
3528                   Exit For                                            'and we're done
3529               End If
3530               cmbComboName.ListIndex = -1                             'else, remove any pointer
3531           Next I
3532       Else
3533           cmbComboName.ListIndex = -1
3534       End If

' <VB WATCH>
3535       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3536       Exit Function
' <VB WATCH>
3537       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3538       Exit Function
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
3539       On Error GoTo vbwErrHandler
3540       Const VBWPROCNAME = "frmPLCData.SetComboTestSetup"
3541       If vbwProtector.vbwTraceProc Then
3542           Dim vbwProtectorParameterString As String
3543           If vbwProtector.vbwTraceParameters Then
3544               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
3545               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sFieldName", sFieldName) & ", "
3546               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sTableName", sTableName) & ", "
3547               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("rs", rs) & ") "
3548           End If
3549           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3550       End If
' </VB WATCH>

       'same as setcombo, except here we also pass in the field name

3551       Dim I As Integer
3552       Dim sParam As String
3553       Dim qy As New ADODB.Command
3554       Dim rs1 As New ADODB.Recordset

3555       If rs.Fields(sFieldName).ActualSize <> 0 Then
               'if plc number, adjust plcaddress id numbers 1 and 2 to plc 8 and 9 respectively
3556           If sTableName = "CirculationFlowMeter" Then
                   'sParam = rs.Fields(sFieldName) + 7
3557               sParam = rs.Fields(sFieldName)
3558               If Val(sParam) < 4 Then
3559                   sParam = str(Val(sParam) + 4)
3560                   rs.Fields(sFieldName) = sParam
3561               End If
3562           Else
3563               sParam = rs.Fields(sFieldName)
3564           End If
3565           qy.ActiveConnection = cnPumpData
3566           qy.CommandText = "SELECT * FROM " & sTableName & " WHERE " & sTableName & " = " & sParam
3567           Set rs1 = qy.Execute()

3568           For I = 0 To cmbComboName.ListCount - 1
3569               If cmbComboName.ItemData(I) = rs1.Fields(0) Then
3570                   cmbComboName.ListIndex = I
3571                   Exit For
3572               End If
3573               cmbComboName.ListIndex = -1
3574           Next I
3575       Else
3576           cmbComboName.ListIndex = -1
3577       End If

' <VB WATCH>
3578       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
3579       Exit Function
' <VB WATCH>
3580       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3581       Exit Function
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
3582       On Error GoTo vbwErrHandler
3583       Const VBWPROCNAME = "frmPLCData.DisablePumpDataControls"
3584       If vbwProtector.vbwTraceProc Then
3585           Dim vbwProtectorParameterString As String
3586           If vbwProtector.vbwTraceParameters Then
3587               vbwProtectorParameterString = "()"
3588           End If
3589           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3590       End If
' </VB WATCH>

3591       txtSalesOrderNumber.Enabled = False
3592       frmMfr.Enabled = False
3593       txtShpNo.Enabled = False
3594       txtBilNo.Enabled = False
3595       txtDesignFlow.Enabled = False
3596       txtDesignTDH.Enabled = False

3597       frmMiscPumpData.Enabled = False

3598       txtModelNo.Enabled = False
3599       txtImpellerDia.Enabled = False

3600       frmTEMC.Enabled = False
3601       frmChempump.Enabled = False

3602       txtRemarks.Enabled = False
3603       Me.cmdAddNewTestDate.Visible = False

3604       cmdEnterPumpData.Enabled = False

' <VB WATCH>
3605       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3606       Exit Sub
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
3607       On Error GoTo vbwErrHandler
3608       Const VBWPROCNAME = "frmPLCData.DisableTestSetupDataControls"
3609       If vbwProtector.vbwTraceProc Then
3610           Dim vbwProtectorParameterString As String
3611           If vbwProtector.vbwTraceParameters Then
3612               vbwProtectorParameterString = "()"
3613           End If
3614           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3615       End If
' </VB WATCH>

3616       cmbTestSpec.Enabled = False
3617       txtWho.Enabled = False
3618       txtRMA.Enabled = False

3619       frmLoopAndXducer.Enabled = False
3620       frmElecData.Enabled = False
3621       frmPerfMods.Enabled = False
3622       frmOtherFiles.Enabled = False
3623       frmInstrumentTags.Enabled = False
3624       frmTAndI.Enabled = False
3625       frmThrustBalMods.Enabled = False
3626       txtTestSetupRemarks.Enabled = False

3627       cmdEnterTestSetupData.Enabled = False
3628       cmbPLCNo.Enabled = False
' <VB WATCH>
3629       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3630       Exit Sub
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
3631       On Error GoTo vbwErrHandler
3632       Const VBWPROCNAME = "frmPLCData.DisableTestDataControls"
3633       If vbwProtector.vbwTraceProc Then
3634           Dim vbwProtectorParameterString As String
3635           If vbwProtector.vbwTraceParameters Then
3636               vbwProtectorParameterString = "()"
3637           End If
3638           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3639       End If
' </VB WATCH>

3640       cmbPLCLoop.Enabled = False
3641       frmPumpData.Enabled = False
3642       frmThermocouples.Enabled = False
3643       frmAI.Enabled = False
3644       frmMagtrol.Enabled = False
3645       fmrMiscTestData.Enabled = False
3646       frmPLCMisc.Enabled = False
3647       DataGrid1.Enabled = False
3648       DataGrid2.Enabled = False
3649       cmdEnterTestData.Enabled = False

' <VB WATCH>
3650       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3651       Exit Sub
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
3652       On Error GoTo vbwErrHandler
3653       Const VBWPROCNAME = "frmPLCData.EnableTestSetupDataControls"
3654       If vbwProtector.vbwTraceProc Then
3655           Dim vbwProtectorParameterString As String
3656           If vbwProtector.vbwTraceParameters Then
3657               vbwProtectorParameterString = "()"
3658           End If
3659           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3660       End If
' </VB WATCH>

3661       cmbTestSpec.Enabled = True
3662       txtWho.Enabled = True
3663       txtRMA.Enabled = True

3664       frmLoopAndXducer.Enabled = True
3665       frmElecData.Enabled = True
3666       frmPerfMods.Enabled = True
3667       frmOtherFiles.Enabled = True
3668       frmInstrumentTags.Enabled = True
3669       frmTAndI.Enabled = True
3670       frmThrustBalMods.Enabled = True
3671       txtTestSetupRemarks.Enabled = True

3672       cmdEnterTestSetupData.Enabled = True
3673       cmbPLCNo.Enabled = True
' <VB WATCH>
3674       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3675       Exit Sub
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
3676       On Error GoTo vbwErrHandler
3677       Const VBWPROCNAME = "frmPLCData.EnableTestDataControls"
3678       If vbwProtector.vbwTraceProc Then
3679           Dim vbwProtectorParameterString As String
3680           If vbwProtector.vbwTraceParameters Then
3681               vbwProtectorParameterString = "()"
3682           End If
3683           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3684       End If
' </VB WATCH>

3685       cmbPLCLoop.Enabled = True
3686       frmPumpData.Enabled = True
3687       frmThermocouples.Enabled = True
3688       frmAI.Enabled = True
3689       frmMagtrol.Enabled = True
3690       fmrMiscTestData.Enabled = True
3691       frmPLCMisc.Enabled = True
3692       DataGrid1.Enabled = True
3693       DataGrid2.Enabled = True
3694       cmdEnterTestData.Enabled = True

' <VB WATCH>
3695       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3696       Exit Sub
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
3697       On Error GoTo vbwErrHandler
3698       Const VBWPROCNAME = "frmPLCData.EnablePumpDataControls"
3699       If vbwProtector.vbwTraceProc Then
3700           Dim vbwProtectorParameterString As String
3701           If vbwProtector.vbwTraceParameters Then
3702               vbwProtectorParameterString = "()"
3703           End If
3704           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3705       End If
' </VB WATCH>

3706       txtSalesOrderNumber.Enabled = True
3707       frmMfr.Enabled = True
3708       txtShpNo.Enabled = True
3709       txtBilNo.Enabled = True
3710       txtDesignFlow.Enabled = True
3711       txtDesignTDH.Enabled = True

3712       frmMiscPumpData.Enabled = True

3713       txtModelNo.Enabled = True
3714       txtImpellerDia.Enabled = True

3715       frmTEMC.Enabled = True
3716       frmChempump.Enabled = True

3717       txtRemarks.Enabled = True
3718       Me.cmdAddNewTestDate.Visible = True

3719       cmdEnterPumpData.Enabled = True

' <VB WATCH>
3720       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3721       Exit Sub
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
3722       On Error GoTo vbwErrHandler
3723       Const VBWPROCNAME = "frmPLCData.EnableMagtrolFields"
3724       If vbwProtector.vbwTraceProc Then
3725           Dim vbwProtectorParameterString As String
3726           If vbwProtector.vbwTraceParameters Then
3727               vbwProtectorParameterString = "()"
3728           End If
3729           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3730       End If
' </VB WATCH>
3731       txtV1.Enabled = True
3732       txtV2.Enabled = True
3733       txtV3.Enabled = True
3734       txtI1.Enabled = True
3735       txtI2.Enabled = True
3736       txtI3.Enabled = True
3737       txtP1.Enabled = True
3738       txtP2.Enabled = True
3739       txtP3.Enabled = True
3740       optKW(0).Visible = True
3741       optKW(1).Visible = True
3742       optKW(2).Visible = True
3743       optKW(1).value = True
3744       optKW_Click (1)
' <VB WATCH>
3745       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3746       Exit Sub
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
3747       On Error GoTo vbwErrHandler
3748       Const VBWPROCNAME = "frmPLCData.DisableMagtrolFields"
3749       If vbwProtector.vbwTraceProc Then
3750           Dim vbwProtectorParameterString As String
3751           If vbwProtector.vbwTraceParameters Then
3752               vbwProtectorParameterString = "()"
3753           End If
3754           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3755       End If
' </VB WATCH>
3756       txtV1.Enabled = False
3757       txtV2.Enabled = False
3758       txtV3.Enabled = False
3759       txtI1.Enabled = False
3760       txtI2.Enabled = False
3761       txtI3.Enabled = False
3762       txtP1.Enabled = False
3763       txtP2.Enabled = False
3764       txtP3.Enabled = False
3765       txtKW.Enabled = False
3766       optKW(0).Visible = False
3767       optKW(1).Visible = False
3768       optKW(2).Visible = False
' <VB WATCH>
3769       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3770       Exit Sub
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
3771       On Error GoTo vbwErrHandler
3772       Const VBWPROCNAME = "frmPLCData.EnablePLCFields"
3773       If vbwProtector.vbwTraceProc Then
3774           Dim vbwProtectorParameterString As String
3775           If vbwProtector.vbwTraceParameters Then
3776               vbwProtectorParameterString = "()"
3777           End If
3778           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3779       End If
' </VB WATCH>
3780       frmPLCData.txtAI1Display.Enabled = True
3781       frmPLCData.txtAI2Display.Enabled = True
3782       frmPLCData.txtAI3Display.Enabled = True
3783       frmPLCData.txtAI4Display.Enabled = True
3784       frmPLCData.txtTC1Display.Enabled = True
3785       frmPLCData.txtTC2Display.Enabled = True
3786       frmPLCData.txtTC3Display.Enabled = True
3787       frmPLCData.txtTC4Display.Enabled = True
3788       frmPLCData.txtFlowDisplay.Enabled = True
3789       frmPLCData.txtSuctionDisplay.Enabled = True
3790       frmPLCData.txtDischargeDisplay.Enabled = True
3791       frmPLCData.txtTemperatureDisplay.Enabled = True
3792       frmPLCData.txtInHgDisplay.Enabled = True
' <VB WATCH>
3793       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3794       Exit Sub
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
3795       On Error GoTo vbwErrHandler
3796       Const VBWPROCNAME = "frmPLCData.DisablePLCFields"
3797       If vbwProtector.vbwTraceProc Then
3798           Dim vbwProtectorParameterString As String
3799           If vbwProtector.vbwTraceParameters Then
3800               vbwProtectorParameterString = "()"
3801           End If
3802           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3803       End If
' </VB WATCH>
3804       frmPLCData.txtAI1Display.Enabled = False
3805       frmPLCData.txtAI2Display.Enabled = False
3806       frmPLCData.txtAI3Display.Enabled = False
3807       frmPLCData.txtAI4Display.Enabled = False
3808       frmPLCData.txtTC1Display.Enabled = False
3809       frmPLCData.txtTC2Display.Enabled = False
3810       frmPLCData.txtTC3Display.Enabled = False
3811       frmPLCData.txtTC4Display.Enabled = False
3812       frmPLCData.txtFlowDisplay.Enabled = False
3813       frmPLCData.txtSuctionDisplay.Enabled = False
3814       frmPLCData.txtDischargeDisplay.Enabled = False
3815       frmPLCData.txtTemperatureDisplay.Enabled = False
3816       frmPLCData.txtInHgDisplay.Enabled = False
' <VB WATCH>
3817       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3818       Exit Sub
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
3819       On Error GoTo vbwErrHandler
3820       Const VBWPROCNAME = "frmPLCData.BlankData"
3821       If vbwProtector.vbwTraceProc Then
3822           Dim vbwProtectorParameterString As String
3823           If vbwProtector.vbwTraceParameters Then
3824               vbwProtectorParameterString = "()"
3825           End If
3826           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3827       End If
' </VB WATCH>
3828       txtShpNo.Text = vbNullString
3829       txtBilNo.Text = vbNullString
3830       txtModelNo.Text = vbNullString
3831       cmbMotor.ListIndex = -1
3832       cmbStatorFill.ListIndex = -1
3833       cmbVoltage.ListIndex = -1
3834       cmbDesignPressure.ListIndex = -1
3835       cmbFrequency.ListIndex = -1
3836       cmbCirculationPath.ListIndex = -1
3837       cmbRPM.ListIndex = -1
3838       cmbModel.ListIndex = -1
3839       cmbModelGroup.ListIndex = -1
3840       txtSpGr.Text = vbNullString
3841       txtImpellerDia.Text = vbNullString
3842       txtEndPlay.Text = vbNullString
3843       txtGGap.Text = vbNullString
3844       txtDesignFlow.Text = vbNullString
3845       txtDesignTDH.Text = vbNullString
3846       txtOtherMods.Text = vbNullString
3847       txtRemarks.Text = vbNullString
3848       txtSalesOrderNumber.Text = vbNullString
3849       txtTestSetupRemarks.Text = vbNullString
3850       txtNPSHFile.Text = vbNullString
3851       txtPicturesFile.Text = vbNullString
3852       txtVibrationFile.Text = vbNullString
       '    cmbOrificeNumber.ListIndex = 18
       '    cmbTestSpec.ListIndex = 6       'default = Rev7
3853       cmbLoopNumber.ListIndex = -1
3854       cmbSuctDia.ListIndex = -1
3855       cmbDischDia.ListIndex = -1
3856       cmbTachID.ListIndex = -1
3857       cmbAnalyzerNo.ListIndex = -1
3858       txtTestRemarks.Text = vbNullString
3859       txtHDCor.Text = 0
3860       txtDischHeight.Text = 0
3861       txtSuctHeight.Text = 0
3862       txtKWMult.Text = 1
3863       txtWho.Text = LogInInitials
3864       txtRMA.Text = vbNullString
3865       frmPLCData.chkNPSH.value = 0
3866       frmPLCData.chkPictures.value = 0
3867       frmPLCData.chkVibration.value = 0
3868       cmbFlowMeter.ListIndex = -1
3869       cmbSuctionPressureTransducer.ListIndex = -1
3870       cmbDischargePressureTransducer.ListIndex = -1
3871       cmbTemperatureTransducer.ListIndex = -1
3872       cmbCirculationFlowMeter.ListIndex = -1
3873       frmPLCData.chkBalanceHoles.value = 0
3874       frmPLCData.chkCircOrifice.value = 0
3875       frmPLCData.txtCircOrifice = vbNullString
3876       frmPLCData.txtImpTrim = vbNullString
3877       frmPLCData.txtOrifice = vbNullString
3878       frmPLCData.chkFeathered.value = Unchecked
3879       frmPLCData.chkTrimmed.value = 0
3880       frmPLCData.chkCircOrifice.value = 0
3881       frmPLCData.txtThrustBal = vbNullString
3882       frmPLCData.txtRPM = vbNullString
3883       frmPLCData.txtVibAx = vbNullString
3884       frmPLCData.txtVibRad = vbNullString
3885       frmPLCData.txtTEMCTRGReading = vbNullString
3886       dgBalanceHoles.Visible = False
3887       Me.txtLineNumber.Text = vbNullString
3888       Me.txtNPSHr.Text = vbNullString
3889       Me.txtRatedInputPower.Text = vbNullString
3890       Me.txtAmps.Text = vbNullString
3891       Me.txtThermalClass.Text = vbNullString
3892       Me.txtViscosity.Text = vbNullString
3893       Me.txtExpClass.Text = vbNullString
3894       Me.txtNoPhases.Text = vbNullString
3895       Me.txtLiquidTemperature.Text = vbNullString
3896       Me.txtJobNum.Text = vbNullString
3897       Me.txtTEMCFrameNumber.Text = vbNullString
3898       Me.txtLiquid.Text = vbNullString
3899       Me.chkSuperMarketFeathered.value = Unchecked
3900       Me.txtRVSPartNo.Text = vbNullString
' <VB WATCH>
3901       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3902       Exit Sub
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
3903       On Error GoTo vbwErrHandler
3904       Const VBWPROCNAME = "frmPLCData.AddTestData"
3905       If vbwProtector.vbwTraceProc Then
3906           Dim vbwProtectorParameterString As String
3907           If vbwProtector.vbwTraceParameters Then
3908               vbwProtectorParameterString = "()"
3909           End If
3910           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3911       End If
' </VB WATCH>
3912       Dim I As Integer
3913       Dim sFilter As String

3914       ClearEff
3915       rsEff.MoveFirst

3916       For I = 1 To 8
3917           rsTestData.AddNew
3918           rsTestData!SerialNumber = txtSN
3919           rsTestData!Date = cmbTestDate.List(cmbTestDate.ListIndex)
3920           rsTestData!testnumber = I
3921           rsTestData!DataWritten = False
3922           rsTestData.Update
3923           DoEfficiencyCalcs
3924           rsEff.MoveNext
3925           rsTestData.MoveNext
3926       Next I
3927       boFoundTestData = True
           'rsTestData.Update
3928       rsTestData.Requery
3929       rsTestData.Resync

          'select the entries from testdata
3930       sFilter = "SerialNumber='" & txtSN.Text & "' AND Date=#" & cmbTestDate.Text & "#"

3931       rsTestData.Filter = sFilter

3932       Set DataGrid1.DataSource = rsTestData

           ' fix the datagrid

3933       Dim c As Column
3934       For Each c In DataGrid1.Columns
3935          Select Case c.DataField
              Case "TestDataID"
3936             c.Visible = False
3937          Case "SerialNumber"
3938             c.Visible = False
3939          Case "Date"
3940             c.Visible = False
3941          Case Else ' Hide all other columns.
3942             c.Visible = True
3943             c.Alignment = dbgRight
3944          End Select
3945       Next c

3946       rsEff.Requery
3947       DataGrid1.Refresh
3948       DataGrid2.Refresh

' <VB WATCH>
3949       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
3950       Exit Sub
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
3951       On Error GoTo vbwErrHandler
3952       Const VBWPROCNAME = "frmPLCData.DoEfficiencyCalcs"
3953       If vbwProtector.vbwTraceProc Then
3954           Dim vbwProtectorParameterString As String
3955           If vbwProtector.vbwTraceParameters Then
3956               vbwProtectorParameterString = "()"
3957           End If
3958           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
3959       End If
' </VB WATCH>
3960       Dim KW As Single, VI As Single, VITemp As Single
3961       Dim Vave As Single, Iave As Single
3962       Dim I As Integer
3963       Dim j As Integer
3964       Dim HeightDiff As Single

3965       If Not IsNull(rsTestData.Fields("TotalPower")) Then
3966           KW = rsTestData.Fields("TotalPower")
3967       Else
               'if we wrote data with an old version, we will not have written total power
               'if total power = 0 and the three individual powers are not 0, add them

3968           If rsTestData.Fields("PowerA") > 0 Then
3969               If rsTestData.Fields("PowerB") > 0 Then
3970                   If rsTestData.Fields("PowerC") > 0 Then
3971                       KW = rsTestData.Fields("PowerA") + rsTestData.Fields("PowerB") + rsTestData.Fields("PowerC")
3972                   End If
3973               End If
3974           End If
3975      End If

3976       I = 0
3977       Vave = 0
3978       Iave = 0
3979       If Not IsNull(rsTestData.Fields("VoltageA")) And Not IsNull(rsTestData.Fields("CurrentA")) Then
3980           VI = rsTestData.Fields("VoltageA") * rsTestData.Fields("CurrentA")
3981           Vave = rsTestData.Fields("VoltageA")
3982           Iave = rsTestData.Fields("CurrentA")
3983           If VI <> 0 Then
3984               I = I + 1
3985           End If
3986       End If
3987       If Not IsNull(rsTestData.Fields("VoltageB")) And Not IsNull(rsTestData.Fields("CurrentB")) Then
3988           VITemp = rsTestData.Fields("VoltageB") * rsTestData.Fields("CurrentB")
3989           If VITemp <> 0 Then
3990               I = I + 1
3991               VI = VI + VITemp
3992               Vave = Vave + rsTestData.Fields("VoltageB")
3993               Iave = Iave + rsTestData.Fields("CurrentB")
3994           End If
3995       End If
3996       If Not IsNull(rsTestData.Fields("VoltageC")) And Not IsNull(rsTestData.Fields("CurrentC")) Then
3997           VITemp = rsTestData.Fields("VoltageC") * rsTestData.Fields("CurrentC")
3998           If VITemp <> 0 Then
3999               I = I + 1
4000               VI = VI + VITemp
4001               Vave = Vave + rsTestData.Fields("VoltageC")
4002               Iave = Iave + rsTestData.Fields("CurrentC")
4003           End If
4004       End If
4005       If KW = 0 Then
4006           For j = 1 To rsEff.Fields.Count - 1
4007               rsEff.Fields(j) = 0
4008           Next j
       '        Exit Sub
4009       End If
4010       If VI <> 0 Then
4011           rsEff.Fields("Volts") = Vave / I
4012           rsEff.Fields("Amps") = Iave / I
4013           rsEff.Fields("PowerFactor") = 1000 * I * KW / (VI * Sqr(3))
4014           rsEff.Fields("PowerFactor") = 100 * rsEff.Fields("PowerFactor")
4015       Else
4016           rsEff.Fields("PowerFactor") = 0
4017       End If

4018       If optMfr(0).value = True Then
4019           If cmbStatorFill.ListIndex = -1 Then
4020               rsEff.Fields("MotorEfficiency") = Format$(0, "0.00")

4021           Else
4022               rsEff.Fields("Motorefficiency") = Format$(Round(MotorEfficiency(KW, cmbMotor.ItemData(cmbMotor.ListIndex), cmbStatorFill.ItemData(cmbStatorFill.ListIndex)), 1), "00.0")
       '            rsEff.Fields("Motorefficiency") = Format$(Round(MotorEfficiency(KW, cmbMotor.ListIndex, cmbStatorFill.ListIndex), 1), "00.0")
4023           End If
4024       Else
4025           rsEff.Fields("MotorEfficiency") = Format$(Round(TEMCMotorEfficiency(KW, txtTEMCFrameNumber.Text, 460, RatedKW), 1), "00.0")
4026       End If

4027       Dim sHDCor As Single
4028       Dim sDisc As Single
4029       Dim sSuct As Single
4030       If IsNull(rsTestSetup.Fields("HDCor")) Then
4031           sHDCor = 0
4032       Else
4033           sHDCor = rsTestSetup.Fields("HDCor")
4034       End If
4035       If IsNull(rsTestSetup.Fields("DischargeGageHeight")) Then
4036           sDisc = 0
4037       Else
4038           sDisc = rsTestSetup.Fields("DischargeGageHeight")
4039       End If
4040       If IsNull(rsTestSetup.Fields("SuctionGageHeight")) Then
4041           sSuct = 0
4042       Else
4043           sSuct = rsTestSetup.Fields("SuctionGageHeight")
4044       End If
4045       HeightDiff = sHDCor + sDisc / 12 - sSuct / 12
4046       If (cmbDischDia.ListIndex <> -1 And cmbSuctDia.ListIndex <> -1) Then
4047           rsEff.Fields("VelocityHead") = CalcVelHead(rsTestData.Fields("Flow"), cmbDischDia.ItemData(cmbDischDia.ListIndex) + 1, cmbSuctDia.ItemData(cmbSuctDia.ListIndex) + 1)
4048       End If
       '    rsEff.Fields("VelocityHead") = CalcVelHead(rsTestData.Fields("Flow"), cmbDischDia.ListIndex + 1, cmbSuctDia.ListIndex + 1)
4049       rsEff.Fields("TDH") = CalcTDH(rsTestData.Fields("DischargePressure"), rsTestData.Fields("SuctionPressure"), rsTestData.Fields("SuctionInHg"), rsEff.Fields("VelocityHead"), HeightDiff, rsTestData.Fields("TemperatureSuction"))
4050       rsEff.Fields("ElecHP") = 1000 * KW / 746
       '    If (DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))) <> 0 And KW <> 0) Then
4051           If Int(rsTestData.Fields("TemperatureSuction")) >= 40 Then
4052               If (DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))) <> 0 And KW <> 0) Then
           '        rsEff.Fields("LiquidHP") = (rsEff.Fields("TDH") * rsTestData.Fields("Flow") * DLookup("TDHCorr", "TempCorrection", "Temp = 68")) / (3960 * DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))))
4053               rsEff.Fields("LiquidHP") = (rsEff.Fields("TDH") * rsTestData.Fields("Flow") * DLookupA(TDHColNo, TempCorrection, TempColNo, 68)) / (3960 * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))))
           '        rsEff.Fields("OverallEfficiency") = (0.189 * rsTestData.Fields("Flow") * rsEff.Fields("TDH") * DLookup("TDHCorr", "TempCorrection", "Temp = 68")) / (10 * KW * DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))))
4054               rsEff.Fields("OverallEfficiency") = (0.189 * rsTestData.Fields("Flow") * rsEff.Fields("TDH") * DLookupA(TDHColNo, TempCorrection, TempColNo, 68)) / (10 * KW * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))))
4055               If rsEff.Fields("MotorEfficiency") <> 0 Then
4056                   rsEff.Fields("HydraulicEfficiency") = 100 * rsEff.Fields("OverallEfficiency") / rsEff.Fields("MotorEfficiency")
4057               Else
4058                   rsEff.Fields("HydraulicEfficiency") = 0
4059               End If
4060           Else
4061               rsEff.Fields("LiquidHP") = 0
4062               rsEff.Fields("OverallEfficiency") = 0
4063           End If

4064       Else
4065           rsEff.Fields("LiquidHP") = 0
4066           rsEff.Fields("OverallEfficiency") = 0
4067       End If


4068       I = rsEff.AbsolutePosition
4069       If Not IsNull(rsTestData.Fields("Flow")) Then
4070           rsEff.Fields("Flow") = rsTestData.Fields("Flow")
4071           HeadFlow(0, I - 1) = rsTestData.Fields("Flow")
4072           HeadFlow(1, I - 1) = rsEff.Fields("TDH")
4073           FlowHead(I - 1, 0) = rsTestData.Fields("Flow")
4074           FlowHead(I - 1, 1) = rsEff.Fields("TDH")

       '        EffFlow(0, i - 1) = rsTestData.Fields("Flow")
       '        EffFlow(1, i - 1) = rsEff.Fields("OverallEfficiency")
       '        KWFlow(0, i - 1) = rsTestData.Fields("Flow")
       '        KWFlow(1, i - 1) = KW
       '        AmpsFlow(0, i - 1) = rsTestData.Fields("Flow")
       '        AmpsFlow(1, i - 1) = rsEff.Fields("Amps")
4075       Else
4076           HeadFlow(0, I - 1) = 0
4077           HeadFlow(1, I - 1) = 0
4078           FlowHead(I - 1, 0) = 0
4079           FlowHead(I - 1, 1) = 0

       '        EffFlow(0, i - 1) = 0
       '        EffFlow(1, i - 1) = 0
       '        KWFlow(0, i - 1) = 0
       '        KWFlow(1, i - 1) = 0
       '        AmpsFlow(0, i - 1) = 0
       '        AmpsFlow(1, i - 1) = 0
4080       End If

4081       Dim Plothead(1, 7) As Single
4082       Dim HeadPlot(7, 1) As Single
           'ReDim Preserve Plothead(1, j)
           'ReDim Preserve HeadPlot(j, 1)

       '    Dim PlotEff() As Single
       '    Dim PlotKW() As Single
       '    Dim PlotAmps() As Single
       '    ReDim PlotHead(0, 0)
       '    ReDim PlotEff(0, 0)
       '    ReDim PlotKW(0, 0)
       '
4083       For j = 0 To UpDown2.value - 1
       '        If HeadFlow(1, j) <> 0 Then
       '            ReDim Preserve Plothead(1, j)
       '            ReDim Preserve HeadPlot(j, 1)
4084               Plothead(0, j) = HeadFlow(0, j)
4085               Plothead(1, j) = HeadFlow(1, j)
4086               HeadPlot(j, 0) = FlowHead(j, 0)
4087               HeadPlot(j, 1) = FlowHead(j, 1)
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
4088       Next j




       '    SetGraphMax (Plothead())
       '    If UBound(PlotHead()) <> 0 Then

       'fix 4/29/19

4089           MSChart1.ChartData = HeadPlot

       '    End If

           'copy fields for reports
4090       rsEff.Fields("DischPress") = rsTestData.Fields("Dischargepressure")
4091       rsEff.Fields("SuctPress") = rsTestData.Fields("Suctionpressure")
       '    rsEff.Fields("Volts") = rsTestData.Fields("VoltageA")
       '    rsEff.Fields("Amps") = rsTestData.Fields("CurrentA")
4092       rsEff.Fields("KW") = KW
4093       rsEff.Fields("Freq") = rsTestData.Fields("VFDFrequency")
4094       rsEff.Fields("RPM") = rsTestData.Fields("RPM")
4095       rsEff.Fields("Pos") = rsTestData.Fields("ThrustBalance")
4096       rsEff.Fields("NPSHa") = rsTestData.Fields("NPSHa")
4097       rsEff.Fields("NPSHr") = rsTestData.Fields("NPSHr")
4098       rsEff.Fields("InputPower") = rsTestData.Fields("TotalPower")
4099       rsEff.Fields("Temperature") = rsTestData.Fields("TemperatureSuction")
4100       rsEff.Fields("CircFlow") = rsTestData.Fields("CircFlow")
4101       rsEff.Fields("VibrationX") = rsTestData.Fields("VibrationX")
4102       rsEff.Fields("VibrationY") = rsTestData.Fields("VibrationY")
4103       rsEff.Fields("CurrentA") = rsTestData.Fields("CurrentA")
4104       rsEff.Fields("CurrentB") = rsTestData.Fields("CurrentB")
4105       rsEff.Fields("CurrentC") = rsTestData.Fields("CurrentC")
4106       rsEff.Fields("VoltageA") = rsTestData.Fields("VoltageA")
4107       rsEff.Fields("VoltageB") = rsTestData.Fields("VoltageB")
4108       rsEff.Fields("VoltageC") = rsTestData.Fields("VoltageC")
4109       rsEff.Fields("TC1") = rsTestData.Fields("TC1")
4110       rsEff.Fields("TC2") = rsTestData.Fields("TC2")
4111       rsEff.Fields("TC3") = rsTestData.Fields("TC3")
4112       rsEff.Fields("TC4") = rsTestData.Fields("TC4")
4113       rsEff.Fields("RBHTemp") = rsTestData.Fields("RBHTemp")
4114       rsEff.Fields("RBHPress") = rsTestData.Fields("RBHPress")
4115       rsEff.Fields("AI4") = rsTestData.Fields("AI4")
4116       rsEff.Fields("Remarks") = rsTestData.Fields("Remarks")
4117       rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCFrontThrust")
4118       rsEff.Fields("TEMCRearThrust") = rsTestData.Fields("TEMCRearThrust")
4119       rsEff.Fields("TEMCTRG") = rsTestData.Fields("TEMCTRG")
4120       rsEff.Fields("TEMCThrustRigPressure") = rsTestData.Fields("TEMCThrustRigPressure")
4121       rsEff.Fields("TEMCMomentArm") = rsTestData.Fields("TEMCMomentArm")
4122       rsEff.Fields("TEMCViscosity") = rsTestData.Fields("TEMCViscosity")
4123       If Not IsNull(rsEff.Fields("TEMCFrontThrust")) Then
4124           txtTEMCFrontThrust.Text = rsEff.Fields("TEMCFrontThrust")
4125       End If
4126       If Not IsNull(rsEff.Fields("TEMCREarThrust")) Then
4127           txtTEMCRearThrust.Text = rsEff.Fields("TEMCREarThrust")
4128       End If
4129       If (Not IsNull(rsEff.Fields("TEMCViscosity"))) And (rsEff.Fields("TEMCViscosity") <> 0) Then
4130           txtTEMCViscosity.Text = rsEff.Fields("TEMCViscosity")
4131       End If
4132       If Not IsNull(rsTestData.Fields("TEMCThrustRigPressure")) Then
4133           txtTEMCThrustRigPressure.Text = rsTestData.Fields("TEMCThrustRigPressure")
4134       End If
4135       If Not IsNull(rsTestData.Fields("TEMCMomentArm")) Then
4136           txtTEMCMomentArm.Text = rsTestData.Fields("TEMCMomentArm")
4137       End If

        '   If Not IsNull(Me.txtAI3Display.Text) Then
        '       Me.txtAI3Display = rsTestData.Fields("RBHPress")
        '   End If

4138       CalculateTEMCForce

4139       If Not IsNull(txtTEMCCalcForce.Text) Then
4140           rsEff.Fields("TEMCCalculatedForce") = Val(txtTEMCCalcForce.Text)
4141       Else
4142           rsEff.Fields("TEMCCalculatedForce") = 0
4143       End If

4144       If Not IsNull(txtTEMCPVValue.Text) Then
4145           rsEff.Fields("TEMCPV") = Val(txtTEMCPVValue.Text)
4146       Else
4147           rsEff.Fields("TEMCPV") = 0
4148       End If

4149       If Val(txtTEMCFrontThrust.Text) <> 0 Then
4150           rsEff.Fields("TEMCFR") = "F"
       '        rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCFrontThrust")
4151       Else
4152           If Val(txtTEMCRearThrust.Text) = 0 Then
                   'no thrust
4153               rsEff.Fields("TEMCFR") = " "
4154               rsEff.Fields("TEMCFrontThrust") = 0
4155           Else
4156               rsEff.Fields("TEMCFR") = "R"
       '            rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCRearThrust")
4157           End If
4158       End If

4159       rsEff.Fields("TEMCForceDirection") = Left(lblTEMCFrontRear.Caption, 1)

4160       rsEff.Update
4161       DataGrid2.Refresh


' <VB WATCH>
4162       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4163       Exit Sub
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
4164       On Error GoTo vbwErrHandler
4165       Const VBWPROCNAME = "frmPLCData.ClearEff"
4166       If vbwProtector.vbwTraceProc Then
4167           Dim vbwProtectorParameterString As String
4168           If vbwProtector.vbwTraceParameters Then
4169               vbwProtectorParameterString = "()"
4170           End If
4171           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4172       End If
' </VB WATCH>
4173       Dim qy As New ADODB.Command

4174       If rsEff.State = adStateOpen Then
4175           If Not (rsEff.BOF = True Or rsEff.EOF = True) Then
4176               rsEff.CancelUpdate
4177           End If
4178           rsEff.Close
4179       End If
4180       qy.ActiveConnection = cnEffData
4181       qy.CommandText = "DROP TABLE Efficiency"
4182       rsEff.Open qy
4183       qy.CommandText = "SELECT EfficiencyOrg.* INTO Efficiency FROM EfficiencyOrg;"
4184       rsEff.Open qy
4185       rsEff.Open "Efficiency", cnEffData, adOpenStatic, adLockOptimistic, adCmdTableDirect

4186       rsEff.Requery
4187       DataGrid2.Refresh

4188       Dim c As Column
4189       For Each c In DataGrid2.Columns
4190           c.Alignment = dbgCenter
4191           c.Width = 750
4192           Select Case c.ColIndex
                   Case 1
4193                   c.Caption = "Flow"
4194                   c.NumberFormat = "###0.00"
4195               Case 2
4196                   c.Caption = "TDH"
4197                   c.NumberFormat = "00.0"
4198               Case 3
4199                   c.Caption = "Overall Eff"
4200                   c.NumberFormat = "00.00"
4201                   c.Width = 850
4202               Case 4
4203                   c.Caption = "PF"
4204                   c.NumberFormat = "00.0"
4205               Case 5
4206                   c.Caption = "Vel Head"
4207                   c.NumberFormat = "00.00"
4208               Case 6
4209                   c.Caption = "Elec HP"
4210                   c.NumberFormat = "#00.0"
4211               Case 7
4212                   c.Caption = "Liq HP"
4213                   c.NumberFormat = "#00.0"
4214               Case Else
4215                   c.Visible = False
4216           End Select
4217       Next c

' <VB WATCH>
4218       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4219       Exit Sub
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
4220       On Error GoTo vbwErrHandler
4221       Const VBWPROCNAME = "frmPLCData.JustAlphaNumeric"
4222       If vbwProtector.vbwTraceProc Then
4223           Dim vbwProtectorParameterString As String
4224           If vbwProtector.vbwTraceParameters Then
4225               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("char", char) & ") "
4226           End If
4227           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4228       End If
' </VB WATCH>
4229       Select Case Asc(char)
               Case 42             ' *
4230               JustAlphaNumeric = char
4231           Case 48 To 57       ' 0 - 9
4232               JustAlphaNumeric = char
4233           Case 65 To 90       ' A - Z
4234               JustAlphaNumeric = char
4235           Case 97 To 122      ' a - z
4236               JustAlphaNumeric = UCase(char)
4237           Case Else
4238               JustAlphaNumeric = ""
4239       End Select
' <VB WATCH>
4240       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4241       Exit Function
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
4242       On Error GoTo vbwErrHandler
4243       Const VBWPROCNAME = "frmPLCData.txtI1_Change"
4244       If vbwProtector.vbwTraceProc Then
4245           Dim vbwProtectorParameterString As String
4246           If vbwProtector.vbwTraceParameters Then
4247               vbwProtectorParameterString = "()"
4248           End If
4249           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4250       End If
' </VB WATCH>
4251       txtI2.Text = txtI1.Text
4252       txtI3.Text = txtI1.Text
' <VB WATCH>
4253       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4254       Exit Sub
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
4255       On Error GoTo vbwErrHandler
4256       Const VBWPROCNAME = "frmPLCData.txtModelNo_Change"
4257       If vbwProtector.vbwTraceProc Then
4258           Dim vbwProtectorParameterString As String
4259           If vbwProtector.vbwTraceParameters Then
4260               vbwProtectorParameterString = "()"
4261           End If
4262           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4263       End If
' </VB WATCH>
4264       Dim I As Integer
4265       Dim S As String
4266       Dim sFull As String
4267       Dim boDone As Boolean
4268       Dim boRepeat As Boolean

4269       Static bo3Digits As Boolean         '3 digits in frame number
4270       Static bo2Digits As Boolean         '2 digits in stages

4271       If optMfr(0).value = True Then
' <VB WATCH>
4272       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4273           Exit Sub
4274       End If

4275       cmbTEMCAdapter.ListIndex = -1
4276       cmbTEMCAdditions.ListIndex = -1
4277       cmbTEMCCirculation.ListIndex = -1
4278       cmbTEMCDesignPressure.ListIndex = -1
4279       cmbTEMCNominalDischargeSize.ListIndex = -1
4280       cmbTEMCDivisionType.ListIndex = -1
4281       cmbTEMCImpellerType.ListIndex = -1
4282       cmbTEMCInsulation.ListIndex = -1
4283       cmbTEMCJacketGasket.ListIndex = -1
4284       cmbTEMCMaterials.ListIndex = -1
4285       cmbTEMCModel.ListIndex = -1
4286       cmbTEMCNominalImpSize.ListIndex = -1
4287       cmbTEMCOtherMotor.ListIndex = -1
4288       cmbTEMCPumpStages.ListIndex = -1
4289       cmbTEMCNominalSuctionSize.ListIndex = -1
4290       cmbTEMCTRG.ListIndex = -1
4291       cmbTEMCVoltage.ListIndex = -1


           'first, get rid of spaces, dashes, etc

4292       S = ""
4293       For I = 1 To Len(txtModelNo.Text)
4294           S = S & JustAlphaNumeric(Mid$(txtModelNo.Text, I, 1))
4295       Next I

           'next, fill out the model number to it's max length of 24 characters

4296       boDone = False
4297       boRepeat = False

4298       Do While Not boDone
4299           sFull = ""
4300           For I = 1 To Len(S)
4301               Select Case I
                       Case 1
                           'type
4302                       sFull = sFull & Mid$(S, I, 1)
4303                   Case 2
                           'adapter
4304                       If IsNumeric(Mid$(S, I, 1)) Then
4305                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4306                           boRepeat = True
4307                           Exit For
4308                       Else
4309                           sFull = sFull & Mid$(S, I, 1)
4310                           boRepeat = False
4311                       End If
4312                   Case 3
                           'materials
4313                       sFull = sFull & Mid$(S, I, 1)
4314                   Case 4
                       'design pressure
4315                       sFull = sFull & Mid$(S, I, 1)
4316                   Case 5
                       'motor frame number - digit 1
4317                       sFull = sFull & Mid$(S, I, 1)
4318                   Case 6
                       'motor frame number - digit 2
4319                       sFull = sFull & Mid$(S, I, 1)
4320                   Case 7
                       'motor frame number - digit 3
4321                       sFull = sFull & Mid$(S, I, 1)
4322                   Case 8
                       'motor frame number - digit 4
4323                       If IsNumeric(Mid$(S, I, 1)) Then
4324                           sFull = sFull & Mid$(S, I, 1)
4325                           boRepeat = False
4326                       Else    '3 digits
       '                        s = Left$(s, i - 1) & "*" & Right$(s, Len(s) - i + 1)
4327                           S = Left$(S, I - 4) & "0" & Right$(S, Len(S) - I + 4)
4328                           boRepeat = True
4329                           Exit For
4330                       End If
4331                   Case 9
                       'insulation
4332                       sFull = sFull & Mid$(S, I, 1)
4333                   Case 10
                       'voltage
4334                       sFull = sFull & Mid$(S, I, 1)
4335                   Case 11
                       'other motor specs
4336                       If Mid$(S, I, 1) = "M" Or Mid$(S, I, 1) = "R" Or Mid$(S, I, 1) = "L" Or Mid$(S, I, 1) = "G" Or Mid$(S, I, 1) = "N" Then
4337                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4338                           boRepeat = True
4339                           Exit For
4340                       Else
4341                           sFull = sFull & Mid$(S, I, 1)
4342                           boRepeat = False
4343                       End If
4344                   Case 12
                       ' TRG
4345                       sFull = sFull & Mid$(S, I, 1)
4346                   Case 13
                       'Nominal discharge - digit 1
4347                       sFull = sFull & Mid$(S, I, 1)
4348                   Case 14
                       'nominal discharge - digit 2
4349                       sFull = sFull & Mid$(S, I, 1)
4350                   Case 15
                       'nominal suction - digit 1
4351                       sFull = sFull & Mid$(S, I, 1)
4352                   Case 16
                       'nominal suction - digit 2
4353                       sFull = sFull & Mid$(S, I, 1)
4354                   Case 17
                       'nominal impeller size
4355                       sFull = sFull & Mid$(S, I, 1)
4356                   Case 18
                       'impeller type
4357                       If Mid$(S, I, 1) <> "*" Then
4358                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4359                           boRepeat = True
4360                           Exit For
4361                       Else
4362                           sFull = sFull & Mid$(S, I, 1)
4363                           boRepeat = False
4364                       End If
4365                   Case 19
                       'Division type
4366                       If IsNumeric(Mid$(S, I, 1)) Then
4367                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4368                           boRepeat = True
4369                           Exit For
4370                       Else
4371                           sFull = sFull & Mid$(S, I, 1)
4372                           boRepeat = False
4373                       End If
4374                   Case 20
                       'pump stages - digit 1
4375                       sFull = sFull & Mid$(S, I, 1)
4376                   Case 21
                       'pump jacket
4377                       If Mid$(S, I, 1) = "A" Or Mid$(S, I, 1) = "B" Or Mid$(S, I, 1) = "E" Or Mid$(S, I, 1) = "F" Or _
                                             Mid$(S, I, 1) = "G" Or Mid$(S, I, 1) = "H" Or Mid$(S, I, 1) = "J" Or Mid$(S, I, 1) = "K" Then
4378                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
4379                           boRepeat = True
4380                       Else
4381                           sFull = sFull & Mid$(S, I, 1)
4382                           boRepeat = False
4383                       End If
4384                   Case 22
                       'additions
4385                         sFull = sFull & Mid$(S, I, 1)
4386                   Case 23
                       'circulation
4387                         sFull = sFull & Mid$(S, I, 1)
4388               End Select
4389           Next I
4390           If Not boRepeat Then
4391               boDone = True
4392           End If
4393       Loop

4394       For I = 1 To Len(sFull)
4395           Select Case I
                   Case 1
4396                   ParseTEMCModelNo cmbTEMCModel, Mid$(sFull, I, 1)
4397               Case 2
4398                   ParseTEMCModelNo cmbTEMCAdapter, Mid$(sFull, I, 1)
4399               Case 3
4400                   ParseTEMCModelNo cmbTEMCMaterials, Mid$(sFull, I, 1)
4401               Case 4
4402                   ParseTEMCModelNo cmbTEMCDesignPressure, Mid$(sFull, I, 1)
4403               Case 5
4404                       If Val(Mid$(sFull, I, 1)) = 0 Then
4405                           txtTEMCFrameNumber.Text = Mid$(sFull, 6, 3)
4406                       Else
4407                           txtTEMCFrameNumber.Text = Mid$(sFull, 5, 4)
4408                       End If
4409               Case 9
4410                       ParseTEMCModelNo cmbTEMCInsulation, Mid$(sFull, I, 1)
4411               Case 10
4412                       ParseTEMCModelNo cmbTEMCVoltage, Mid$(sFull, I, 1)
4413               Case 11
4414                       ParseTEMCModelNo cmbTEMCOtherMotor, Mid$(sFull, I, 1)
4415               Case 12
4416                       ParseTEMCModelNo cmbTEMCTRG, Mid$(sFull, I, 1)
4417               Case 13
4418                       ParseTEMCModelNo cmbTEMCNominalDischargeSize, Mid$(sFull, I, 2)
4419               Case 14
4420               Case 15
4421                       ParseTEMCModelNo cmbTEMCNominalSuctionSize, Mid$(sFull, I, 2)
4422               Case 16
4423               Case 17
4424                       ParseTEMCModelNo cmbTEMCNominalImpSize, Mid$(sFull, I, 1)
4425               Case 18
4426                       ParseTEMCModelNo cmbTEMCImpellerType, Mid$(sFull, I, 1)
4427               Case 19
4428                       ParseTEMCModelNo cmbTEMCDivisionType, Mid$(sFull, I, 1)
4429               Case 20
4430                       ParseTEMCModelNo cmbTEMCPumpStages, Mid$(sFull, I, 1)
4431               Case 21
4432                       ParseTEMCModelNo cmbTEMCJacketGasket, Mid$(sFull, I, 1)
4433               Case 22
4434                       ParseTEMCModelNo cmbTEMCAdditions, Mid$(sFull, I, 1)
4435                       ParseTEMCModelNo cmbTEMCCirculation, "*"
4436               Case 23
       '                    ParseTEMCModelNo cmbTEMCCirculation, Mid$(sFull, I, 1)

4437           End Select
4438       Next I

           'give alerts on certain conditions
4439       Dim msg As String
4440       msg = ""
4441       If Left(cmbTEMCVoltage, 3) = "[6]" Then
4442           msg = "Requires Transformer"
4443       End If
4444       If Left(cmbTEMCTRG, 3) = "[L]" Or InStr("X[B][F][H][K]", Left(cmbTEMCAdditions, 3)) > 1 Then
4445           If msg = "" Then
4446               msg = "Requires VFD"
4447           Else
4448               msg = msg & " and " & "Requires VFD"
4449           End If
4450       End If

4451       If msg <> "" Then
4452           frmAlert.txtAlert.Text = msg
4453           frmAlert.Show
4454       End If

' <VB WATCH>
4455       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4456       Exit Sub
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
4457       On Error GoTo vbwErrHandler
4458       Const VBWPROCNAME = "frmPLCData.txtModelNo_Validate"
4459       If vbwProtector.vbwTraceProc Then
4460           Dim vbwProtectorParameterString As String
4461           If vbwProtector.vbwTraceParameters Then
4462               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Cancel", Cancel) & ") "
4463           End If
4464           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4465       End If
' </VB WATCH>
4466       Dim I As Integer
4467       Dim S As String

       '    s = txtModelNo.Text
       '    S = Replace(S, "-", "")
       '    S = Replace(S, " ", "")
       '    S = Replace(S, "/", "")

       '    txtModelNo.Text = ""

       '    For i = 1 To Len(s)
       '        txtModelNo.Text = txtModelNo.Text & Mid(s, i, 1)
       '    Next i
4468       txtModelNo_Change

' <VB WATCH>
4469       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4470       Exit Sub
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
4471       On Error GoTo vbwErrHandler
4472       Const VBWPROCNAME = "frmPLCData.txtNPSHFile_GotFocus"
4473       If vbwProtector.vbwTraceProc Then
4474           Dim vbwProtectorParameterString As String
4475           If vbwProtector.vbwTraceParameters Then
4476               vbwProtectorParameterString = "()"
4477           End If
4478           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4479       End If
' </VB WATCH>
4480       On Error GoTo FileCancel
4481       If LenB(txtNPSHFile.Text) <> 0 Then
4482           CommonDialog1.filename = txtNPSHFile.Text
4483       End If
4484       CommonDialog1.ShowOpen
4485       txtNPSHFile.Text = CommonDialog1.filename
' <VB WATCH>
4486       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4487       Exit Sub
4488   FileCancel:
4489   On Error GoTo vbwErrHandler
4490       CommonDialog1.CancelError = False
' <VB WATCH>
4491       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4492       Exit Sub
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
4493       On Error GoTo vbwErrHandler
4494       Const VBWPROCNAME = "frmPLCData.txtP1_Change"
4495       If vbwProtector.vbwTraceProc Then
4496           Dim vbwProtectorParameterString As String
4497           If vbwProtector.vbwTraceParameters Then
4498               vbwProtectorParameterString = "()"
4499           End If
4500           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4501       End If
' </VB WATCH>
4502       txtP2.Text = txtP1.Text
4503       txtP3.Text = txtP1.Text
' <VB WATCH>
4504       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4505       Exit Sub
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
4506       On Error GoTo vbwErrHandler
4507       Const VBWPROCNAME = "frmPLCData.txtPicturesFile_gotfocus"
4508       If vbwProtector.vbwTraceProc Then
4509           Dim vbwProtectorParameterString As String
4510           If vbwProtector.vbwTraceParameters Then
4511               vbwProtectorParameterString = "()"
4512           End If
4513           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4514       End If
' </VB WATCH>
4515       CommonDialog1.CancelError = True
4516       On Error GoTo FileCancel
4517       If LenB(txtPicturesFile.Text) <> 0 Then
4518           CommonDialog1.filename = txtPicturesFile.Text
4519       End If
4520       CommonDialog1.ShowOpen
4521       txtPicturesFile.Text = CommonDialog1.filename
' <VB WATCH>
4522       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4523       Exit Sub
4524   FileCancel:
4525   On Error GoTo vbwErrHandler
4526       CommonDialog1.CancelError = False
' <VB WATCH>
4527       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4528       Exit Sub
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
4529       On Error GoTo vbwErrHandler
4530       Const VBWPROCNAME = "frmPLCData.txtSN_Change"
4531       If vbwProtector.vbwTraceProc Then
4532           Dim vbwProtectorParameterString As String
4533           If vbwProtector.vbwTraceParameters Then
4534               vbwProtectorParameterString = "()"
4535           End If
4536           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4537       End If
' </VB WATCH>
4538       cmdFindPump.Default = True
' <VB WATCH>
4539       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4540       Exit Sub
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
4541       On Error GoTo vbwErrHandler
4542       Const VBWPROCNAME = "frmPLCData.txtTEMCFrontThrust_Change"
4543       If vbwProtector.vbwTraceProc Then
4544           Dim vbwProtectorParameterString As String
4545           If vbwProtector.vbwTraceParameters Then
4546               vbwProtectorParameterString = "()"
4547           End If
4548           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4549       End If
' </VB WATCH>
4550       CalculateTEMCForce
' <VB WATCH>
4551       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4552       Exit Sub
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
4553       On Error GoTo vbwErrHandler
4554       Const VBWPROCNAME = "frmPLCData.txtTEMCMomentArm_Change"
4555       If vbwProtector.vbwTraceProc Then
4556           Dim vbwProtectorParameterString As String
4557           If vbwProtector.vbwTraceParameters Then
4558               vbwProtectorParameterString = "()"
4559           End If
4560           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4561       End If
' </VB WATCH>
4562       CalculateTEMCForce
' <VB WATCH>
4563       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4564       Exit Sub
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
4565       On Error GoTo vbwErrHandler
4566       Const VBWPROCNAME = "frmPLCData.txtTEMCRearThrust_Change"
4567       If vbwProtector.vbwTraceProc Then
4568           Dim vbwProtectorParameterString As String
4569           If vbwProtector.vbwTraceParameters Then
4570               vbwProtectorParameterString = "()"
4571           End If
4572           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4573       End If
' </VB WATCH>
4574       CalculateTEMCForce
' <VB WATCH>
4575       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4576       Exit Sub
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
4577       On Error GoTo vbwErrHandler
4578       Const VBWPROCNAME = "frmPLCData.txtTEMCThrustRigPressure_Change"
4579       If vbwProtector.vbwTraceProc Then
4580           Dim vbwProtectorParameterString As String
4581           If vbwProtector.vbwTraceParameters Then
4582               vbwProtectorParameterString = "()"
4583           End If
4584           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4585       End If
' </VB WATCH>
4586       CalculateTEMCForce
' <VB WATCH>
4587       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4588       Exit Sub
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
4589       On Error GoTo vbwErrHandler
4590       Const VBWPROCNAME = "frmPLCData.txtTEMCViscosity_Change"
4591       If vbwProtector.vbwTraceProc Then
4592           Dim vbwProtectorParameterString As String
4593           If vbwProtector.vbwTraceParameters Then
4594               vbwProtectorParameterString = "()"
4595           End If
4596           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4597       End If
' </VB WATCH>
4598       CalculateTEMCForce
' <VB WATCH>
4599       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4600       Exit Sub
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
4601       On Error GoTo vbwErrHandler
4602       Const VBWPROCNAME = "frmPLCData.txtV1_Change"
4603       If vbwProtector.vbwTraceProc Then
4604           Dim vbwProtectorParameterString As String
4605           If vbwProtector.vbwTraceParameters Then
4606               vbwProtectorParameterString = "()"
4607           End If
4608           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4609       End If
' </VB WATCH>
4610       txtV2.Text = txtV1.Text
4611       txtV3.Text = txtV1.Text
' <VB WATCH>
4612       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4613       Exit Sub
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
4614       On Error GoTo vbwErrHandler
4615       Const VBWPROCNAME = "frmPLCData.txtVibrationFile_gotfocus"
4616       If vbwProtector.vbwTraceProc Then
4617           Dim vbwProtectorParameterString As String
4618           If vbwProtector.vbwTraceParameters Then
4619               vbwProtectorParameterString = "()"
4620           End If
4621           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4622       End If
' </VB WATCH>
4623       On Error GoTo FileCancel
4624       If LenB(txtVibrationFile.Text) <> 0 Then
4625           CommonDialog1.filename = txtVibrationFile.Text
4626       End If
4627       CommonDialog1.ShowOpen
4628       txtVibrationFile.Text = CommonDialog1.filename
' <VB WATCH>
4629       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4630       Exit Sub
4631   FileCancel:
4632   On Error GoTo vbwErrHandler
4633       CommonDialog1.CancelError = False
' <VB WATCH>
4634       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
4635       Exit Sub
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
4636       On Error GoTo vbwErrHandler
4637       Const VBWPROCNAME = "frmPLCData.ExportToExcel"
4638       If vbwProtector.vbwTraceProc Then
4639           Dim vbwProtectorParameterString As String
4640           If vbwProtector.vbwTraceParameters Then
4641               vbwProtectorParameterString = "()"
4642           End If
4643           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
4644       End If
' </VB WATCH>

4645       Dim SaveFileName As String
4646       Dim WorkSheetName As String

4647       Dim I As Integer
4648       Dim iRowNo As Integer
4649       Dim sImp As String
4650       Dim ans As Integer

4651       Dim bCanShowSpeed As Boolean
4652       Dim CantShowReason As String

       'close any running excel processes
4653       Dim objWMIService, colProcesses
4654       Set objWMIService = GetObject("winmgmts:")
4655       Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name LIKE 'Excel%'")
4656       If colProcesses.Count > 0 Then
4657           Set xlApp = Excel.Application
4658       Else
               'use existing copy
       '        Set xlApp = New Excel.Application
4659           Set xlApp = CreateObject("Excel.Application")
4660       End If


4661       CommonDialog1.CancelError = True        'in case the user
4662       On Error GoTo ErrHandler                '  chooses the cancel button

           'set up dialog box
4663       CommonDialog1.DialogTitle = "Open Excel Files"
4664       CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|"  'show Excel files
4665       CommonDialog1.InitDir = App.Path
       '    CommonDialog1.InitDir = "C:\"    'in this directory
4666       CommonDialog1.ShowOpen                              'open the file selection dialog box

4667       If Dir(CommonDialog1.filename) = "" Then            'if the file name does not exist yet
4668           SaveFileName = CommonDialog1.filename           'get the name of the file
4669           If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
4670                xlApp.Workbooks.Close
4671           End If
               ' Create the Excel Workbook Object.
4672   On Error GoTo vbwErrHandler
4673           Set xlBook = xlApp.Workbooks.Add                'add a workbook
4674           WorkSheetName = NewWorkBook                                     'do some stuff for the new workbook
4675           ActiveWorkbook.CheckCompatibility = False
4676           xlApp.ActiveWorkbook.SaveAs filename:=SaveFileName, _
                                 FileFormat:=xlNormal                        'save the file
4677       Else                                                'the file name already exists
4678           SaveFileName = CommonDialog1.filename
               ' Create the Excel Workbook Object.
4679           If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
4680                xlApp.Workbooks.Close
4681           End If
4682           Set xlBook = xlApp.Workbooks.Open(SaveFileName)             'get the file name selected
4683           If GetWorksheetTabs(SaveFileName, WorkSheetName) = vbNo Then    'ask the user if he/she wants a new tab.
4684               MsgBox "File not overwritten.", vbOKOnly, "File not Opened"
' <VB WATCH>
4685       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
4686               Exit Sub
4687           Else
4688           End If
4689       End If

4690   On Error GoTo vbwErrHandler

           'see if we can export Speed and SG and if we can, ask user if s/he wants it
           'assume that we can show speed calcs

4691       bCanShowSpeed = False
       'open the template and copy the data from the sheet
       '  excel file resides in ParentDirectoryName + "\Polar SG&Visc Correction5.xls"
           'write the data to the spreadsheet
4692       With xlApp

4693       Dim xlTemplateName As String
4694       xlTemplateName = ParentDirectoryName & sSGandViscSpreadsheetTemplate
4695       Dim xlTemplate As Excel.Workbook
4696       Set xlTemplate = xlApp.Workbooks.Open(xlTemplateName)
4697       Dim TemplateWS As Excel.Worksheet
4698       Dim sheetName As String
4699       sheetName = xlTemplate.Sheets(1).Name
4700       xlTemplate.Sheets(1).Copy After:=xlBook.Sheets(WorkSheetName)

4701       xlTemplate.Close savechanges:=False

4702       Set xlTemplate = Nothing

4703       Application.DisplayAlerts = False
4704       ActiveWorkbook.Worksheets(WorkSheetName).Delete
4705       Application.DisplayAlerts = True
4706       ActiveWorkbook.Worksheets(sheetName).Name = WorkSheetName

           'WorkSheetName = sheetName

           'first see if there is an entry in CalculatedRPM table for this frame size and voltage.
           ' if there is, get the coefficients, else make the coefficients 0

4707           Dim ACoef As Double
4708           Dim BCoef As Double
4709           Dim CCoef As Double

4710           Dim qy As New ADODB.Command
4711           Dim rs As New ADODB.Recordset
4712           qy.ActiveConnection = cnPumpData
4713           Dim VoltageForLookup As Integer
4714           If cmbVoltage.List(cmbVoltage.ListIndex) = "380" And cmbFrequency.List(cmbFrequency.ListIndex) = "50 Hz" Then
4715               VoltageForLookup = 460
4716           ElseIf cmbVoltage.List(cmbVoltage.ListIndex) <> "380" Then
4717               VoltageForLookup = cmbVoltage.List(cmbVoltage.ListIndex)
4718           End If
4719           qy.CommandText = "SELECT * FROM CalculatedRPM WHERE FrameNumber = '" & txtTEMCFrameNumber.Text & _
                          "' AND Voltage = '" & VoltageForLookup & "'"

4720           rs.CursorLocation = adUseClient
4721           rs.CursorType = adOpenStatic

4722           rs.Open qy
4723           If rs.RecordCount = 0 Then
4724               ACoef = 0
4725               BCoef = 0
4726               CCoef = 0
4727               MsgBox ("Cannot find coefficient data for Frame Number " & txtTEMCFrameNumber.Text & _
                          " AND Voltage = " & cmbVoltage.List(cmbVoltage.ListIndex) & _
                          " AND Frequency = " & cmbFrequency.List(cmbFrequency.ListIndex))
4728           Else
4729               ACoef = rs.Fields("A")
4730               BCoef = rs.Fields("B")
4731               CCoef = rs.Fields("C")
4732           End If


           'write header data

4733           .Range("A2").Select
4734           .ActiveCell.FormulaR1C1 = "Serial Number"
4735           .Range("C2").Select
4736           .ActiveCell.FormulaR1C1 = txtSN

4737           .Range("F1").Select
4738           .ActiveCell.FormulaR1C1 = "Customer"
4739           .Range("H1").Select
4740           .ActiveCell.FormulaR1C1 = txtShpNo

4741           .Range("A3").Select
4742           .ActiveCell.FormulaR1C1 = "Model"
4743           .Range("C3").Select
4744           .ActiveCell.FormulaR1C1 = txtModelNo

4745           .Range("F2").Select
4746           .ActiveCell.FormulaR1C1 = "Sales Order"
4747           .Range("H2").Select
4748           .ActiveCell.FormulaR1C1 = txtSalesOrderNumber

4749           .Range("A9").Select
4750           .ActiveCell.FormulaR1C1 = "Design Flow"
4751           .Range("C9").Select
4752           .ActiveCell.FormulaR1C1 = Val(txtDesignFlow)

4753           .Range("A10").Select
4754           .ActiveCell.FormulaR1C1 = "Design Head"
4755           .Range("C10").Select
4756           .ActiveCell.FormulaR1C1 = Val(txtDesignTDH)

4757           .Range("P13").Select
4758           .ActiveCell.FormulaR1C1 = "Barometric Pressure"
4759           .Range("R13").Select
4760           .ActiveCell.FormulaR1C1 = Val(txtInHgDisplay)

4761           .Range("P11").Select
4762           .ActiveCell.FormulaR1C1 = "Suction Gage Height"
4763           .Range("R11").Select
4764           .ActiveCell.FormulaR1C1 = Val(txtSuctHeight)

4765           .Range("P12").Select
4766           .ActiveCell.FormulaR1C1 = "Discharge Gage Height"
4767           .Range("R12").Select
4768           .ActiveCell.FormulaR1C1 = Val(txtDischHeight)

4769           .Range("A1").Select
4770           .ActiveCell.FormulaR1C1 = "Run Date"
4771           .Range("C1").Select
4772           .ActiveCell.FormulaR1C1 = cmbTestDate.List(cmbTestDate.ListIndex)

4773           .Range("D10:E10").Select
4774           With xlApp.Selection
4775               .HorizontalAlignment = xlCenter
4776               .VerticalAlignment = xlBottom
4777               .WrapText = False
4778               .Orientation = 0
4779               .AddIndent = False
4780               .IndentLevel = 0
4781               .ShrinkToFit = False
4782               .ReadingOrder = xlContext
4783               .MergeCells = False
4784           End With
4785           xlApp.Selection.Merge

               'determine rpm

4786           Dim RPMvalue As String
4787           If Mid$(Me.txtTEMCFrameNumber.Text, 2, 1) = "1" Then
               '1 says 2 pole
4788               If Me.cmbFrequency.ListIndex = 0 Then
                       '0 says 50Hz
4789                   RPMvalue = "2900"
4790               ElseIf Me.cmbFrequency.ListIndex = 1 Then
                       ' says 60Hz
4791                   RPMvalue = "3450"
4792               Else
                       'vfd or other, no rpm
4793                   RPMvalue = ""
4794               End If
4795           Else
               '2 says 4 pole
4796               If Me.cmbFrequency.ListIndex = 0 Then
                       '0 says 50Hz
4797                   RPMvalue = "1450"
4798               ElseIf Me.cmbFrequency.ListIndex = 1 Then
                       ' says 60Hz
4799                   RPMvalue = "1750"
4800               Else
                       'vfd or other, no rpm
4801                   RPMvalue = ""
4802               End If
4803           End If

       '        .Range("G1").Select
       '        .ActiveCell.FormulaR1C1 = "RPM"
       '        .Range("I1").Select
       '        .ActiveCell.FormulaR1C1 = RPMvalue

4804           .Range("A5").Select
4805           .ActiveCell.FormulaR1C1 = "Sp Gravity"
4806           .Range("C5").Select
4807           .ActiveCell.FormulaR1C1 = txtSpGr

4808           .Range("A6").Select
4809           .ActiveCell.FormulaR1C1 = "Viscosity"
4810           .Range("C6").Select
4811           .ActiveCell.FormulaR1C1 = txtViscosity

4812           .Range("F4").Select
4813           .ActiveCell.FormulaR1C1 = "Motor"
4814           .Range("H4").Select
4815           .ActiveCell.FormulaR1C1 = txtTEMCFrameNumber.Text

4816           .Range("H12").Select
4817           .ActiveCell.FormulaR1C1 = Me.txtCustPONum.Text

4818           .Range("F5").Select
4819           .ActiveCell.FormulaR1C1 = "Voltage"
4820           .Range("H5").Select
4821           .ActiveCell.FormulaR1C1 = cmbVoltage.List(cmbVoltage.ListIndex)

4822           .Range("K6").Select
4823           .ActiveCell.FormulaR1C1 = "End Play"
4824           .Range("M6").Select
4825           .ActiveCell.FormulaR1C1 = Val(txtEndPlay)

4826           .Range("K7").Select
4827           .ActiveCell.FormulaR1C1 = "G-Gap"
4828           .Range("M7").Select
4829           .ActiveCell.FormulaR1C1 = txtGGap.Text

4830           .Range("A8").Select
4831           .ActiveCell.FormulaR1C1 = "Design Pressure"
4832           .Range("C8").Select
4833           Dim DesPress As String
4834           DesPress = cmbTEMCDesignPressure.List(cmbTEMCDesignPressure.ListIndex)
4835           Dim j As Integer
4836           j = InStrRev(DesPress, "-")
4837           .ActiveCell.FormulaR1C1 = Mid$(DesPress, j + 2)

       '        .Range("G8").Select
       '        .ActiveCell.FormulaR1C1 = "Stator Fill"
       '        .Range("I8").Select
       '        .ActiveCell.FormulaR1C1 = "Dry"

4838           .Range("K4").Select
4839           .ActiveCell.FormulaR1C1 = "Circulation Path"
4840           .Range("M4").Select
4841           .ActiveCell.FormulaR1C1 = cmbTEMCModel.List(cmbTEMCModel.ListIndex)

4842           .Range("M8").Select
4843           .ActiveCell.FormulaR1C1 = txtNPSHr.Text

4844           .Range("K1").Select
4845           .ActiveCell.FormulaR1C1 = "Impeller Dia"
4846           .Range("M1").Select


       '        If LenB(txtImpTrim) <> 0 Then
       '            .ActiveCell.FormulaR1C1 = Val(txtImpTrim)
       '        Else
       '            .ActiveCell.FormulaR1C1 = Val(txtImpellerDia)
       '        End If
       '
4847           If chkTrimmed.value = 1 Then
4848               If Val(txtImpTrim.Text) <> 0 Then
4849                   .ActiveCell.FormulaR1C1 = txtImpTrim
4850               Else
4851                   .ActiveCell.FormulaR1C1 = txtImpellerDia
4852               End If
4853           Else
4854               .ActiveCell.FormulaR1C1 = txtImpellerDia
4855           End If



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

4856           .Range("P9").Select
4857           .ActiveCell.FormulaR1C1 = "Suction Dia"
4858           .Range("R9").Select
4859           .ActiveCell.FormulaR1C1 = cmbSuctDia.List(cmbSuctDia.ListIndex)

4860           .Range("P10").Select
4861           .ActiveCell.FormulaR1C1 = "Discharge Dia"
4862           .Range("R10").Select
4863           .ActiveCell.FormulaR1C1 = cmbDischDia.List(cmbDischDia.ListIndex)

4864           .Range("A11").Select
4865           .ActiveCell.FormulaR1C1 = "Test Spec"
4866           .Range("C11").Select
4867           .ActiveCell.FormulaR1C1 = cmbTestSpec.List(cmbTestSpec.ListIndex)

4868           .Range("K3").Select
4869           .ActiveCell.FormulaR1C1 = "Impeller Feathered"
4870           .Range("M3").Select
4871           If chkFeathered.value = 1 Then
4872               .ActiveCell.FormulaR1C1 = "Yes"
4873           Else
4874               .ActiveCell.FormulaR1C1 = "No"
4875           End If

4876           .Range("K2").Select
4877           .ActiveCell.FormulaR1C1 = "Disch Orifice"
4878           .Range("M2").Select
4879           If chkOrifice.value = 1 Then
4880               .ActiveCell.FormulaR1C1 = Val(txtOrifice)
4881           Else
4882               .ActiveCell.FormulaR1C1 = "None"
4883           End If


4884           .Range("K5").Select
4885           .ActiveCell.FormulaR1C1 = "Circulation Orifice"
4886           .Range("M5").Select
4887           If chkCircOrifice.value = 1 Then
4888               .ActiveCell.FormulaR1C1 = Val(txtCircOrifice)
4889           Else
4890               .ActiveCell.FormulaR1C1 = "None"
4891           End If

4892           .Range("A13").Select
4893           .ActiveCell.FormulaR1C1 = "Other Mods"
4894           .Range("C13").Select
4895           .ActiveCell.FormulaR1C1 = txtOtherMods

4896           .Range("A14").Select
4897           .ActiveCell.FormulaR1C1 = "Remarks"
4898           .Range("C14").Select
4899           .ActiveCell.FormulaR1C1 = txtRemarks

4900           .Range("A15").Select
4901           .ActiveCell.FormulaR1C1 = "Test Setup Remarks"
4902           .Range("C15").Select
4903           .ActiveCell.FormulaR1C1 = txtTestSetupRemarks

4904           .Range("P1").Select
4905           .ActiveCell.FormulaR1C1 = "Suct ID"
4906           .Range("R1").Select
4907           .ActiveCell.FormulaR1C1 = cmbSuctionPressureTransducer.List(cmbSuctionPressureTransducer.ListIndex)

4908           .Range("P2").Select
4909           .ActiveCell.FormulaR1C1 = "Disch ID"
4910           .Range("R2").Select
4911           .ActiveCell.FormulaR1C1 = cmbDischargePressureTransducer.List(cmbDischargePressureTransducer.ListIndex)

4912           .Range("P3").Select
4913           .ActiveCell.FormulaR1C1 = "Temp ID"
4914           .Range("R3").Select
4915           .ActiveCell.FormulaR1C1 = cmbTemperatureTransducer.List(cmbTemperatureTransducer.ListIndex)

4916           .Range("P4").Select
4917           .ActiveCell.FormulaR1C1 = "Circ Flow ID"
4918           .Range("R4").Select
4919           .ActiveCell.FormulaR1C1 = cmbCirculationFlowMeter.List(cmbCirculationFlowMeter.ListIndex)

4920           .Range("P5").Select
4921           .ActiveCell.FormulaR1C1 = "Flow ID"
4922           .Range("R5").Select
4923           .ActiveCell.FormulaR1C1 = cmbFlowMeter.List(cmbFlowMeter.ListIndex)

4924           .Range("P6").Select
4925           .ActiveCell.FormulaR1C1 = "Analyzer ID"
4926           .Range("R6").Select
4927           .ActiveCell.FormulaR1C1 = cmbAnalyzerNo.List(cmbAnalyzerNo.ListIndex)

4928           .Range("P7").Select
4929           .ActiveCell.FormulaR1C1 = "Loop ID"
4930           .Range("R7").Select
4931           .ActiveCell.FormulaR1C1 = cmbLoopNumber.List(cmbLoopNumber.ListIndex)

4932           .Range("A4").Select
4933           .ActiveCell.FormulaR1C1 = "Fluid"
4934           .Range("C4").Select
4935           .ActiveCell.FormulaR1C1 = txtLiquid.Text

4936           .Range("F3").Select
4937           .ActiveCell.FormulaR1C1 = "Cust PN"
4938           .Range("H3").Select
       '        .ActiveCell.FormulaR1C1 = txtRMA.Text
4939           If rsPumpData.Fields("RVSPartNo") <> "" Then
4940               .ActiveCell.FormulaR1C1 = rsPumpData.Fields("RVSPartNo")
4941           End If
4942           If rsPumpData.Fields("CustPN") <> "" Then
4943               .ActiveCell.FormulaR1C1 = rsPumpData.Fields("CustPN")
4944           End If

4945           .Range("A7").Select
4946           .ActiveCell.FormulaR1C1 = "Temperature"
4947           .Range("C7").Select
4948           .ActiveCell.FormulaR1C1 = txtLiquidTemperature.Text

4949           .Range("F6").Select
4950           .ActiveCell.FormulaR1C1 = "Frequency"
4951           .Range("H6").Select
4952           If UCase(cmbFrequency.List(cmbFrequency.ListIndex)) = "VFD" Then
4953               .ActiveCell.FormulaR1C1 = Val(Me.txtVFDFreq)
4954           Else
4955               .ActiveCell.FormulaR1C1 = Val(cmbFrequency.List(cmbFrequency.ListIndex))
4956           End If
       '        .Range("K2").Select
       '        .ActiveCell.FormulaR1C1 = "Disch Orifice"
       '        .Range("M2").Select
       '        .ActiveCell.FormulaR1C1 = txtOrifice.Text

       '        .Range("K12").Select
       '        .ActiveCell.FormulaR1C1 = "Flow Orifice"
       '        .Range("L12").Select
       '        .ActiveCell.FormulaR1C1 = txtCircOrifice.Text

4957           .Range("P8").Select
4958           .ActiveCell.FormulaR1C1 = "PLC No"
4959           .Range("R8").Select
4960           .ActiveCell.FormulaR1C1 = cmbPLCNo.List(cmbPLCNo.ListIndex)

4961           .Range("F7").Select
4962           .ActiveCell.FormulaR1C1 = "Phases"
4963           .Range("H7").Select
4964           .ActiveCell.FormulaR1C1 = txtNoPhases.Text

4965           .Range("F8").Select
4966           .ActiveCell.FormulaR1C1 = "Poles"
4967           .Range("H8").Select
4968           .ActiveCell.FormulaR1C1 = 2 * Val(Left$(Right$(txtTEMCFrameNumber.Text, 2), 1))

4969           .Range("F9").Select
4970           .ActiveCell.FormulaR1C1 = "Rated Current"
4971           .Range("H9").Select
4972           .ActiveCell.FormulaR1C1 = txtAmps.Text

4973           .Range("F10").Select
4974           .ActiveCell.FormulaR1C1 = "Rated Input Power"
4975           .Range("H10").Select
4976           .ActiveCell.FormulaR1C1 = txtRatedInputPower.Text

4977           .Range("F11").Select
4978           .ActiveCell.FormulaR1C1 = "Insulation Class"
4979           .Range("H11").Select
4980           .ActiveCell.FormulaR1C1 = txtThermalClass.Text

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

4981           .Range("A17").Select
4982           .ActiveCell.FormulaR1C1 = "Flow"
4983           .Range("A18").Select
4984           .ActiveCell.FormulaR1C1 = "(GPM)"

4985           .Range("B17").Select
4986           .ActiveCell.FormulaR1C1 = "TDH"
4987           .Range("B18").Select
4988           .ActiveCell.FormulaR1C1 = "(Ft)"

4989           .Range("C17").Select
4990           .ActiveCell.FormulaR1C1 = "KW"

4991           .Range("D17").Select
4992           .ActiveCell.FormulaR1C1 = "Ave"
4993           .Range("D18").Select
4994           .ActiveCell.FormulaR1C1 = "Volts"

4995           .Range("E17").Select
4996           .ActiveCell.FormulaR1C1 = "Ave"
4997           .Range("E18").Select
4998           .ActiveCell.FormulaR1C1 = "Amps"

4999           .Range("F17").Select
5000           .ActiveCell.FormulaR1C1 = "Power"
5001           .Range("F18").Select
5002           .ActiveCell.FormulaR1C1 = "Factor"

5003           .Range("G17").Select
5004           .ActiveCell.FormulaR1C1 = "Overall"
5005           .Range("G18").Select
5006           .ActiveCell.FormulaR1C1 = "Eff"

5007           .Range("H17").Select
5008           .ActiveCell.FormulaR1C1 = "Measured"
5009           .Range("H18").Select
5010           .ActiveCell.FormulaR1C1 = "RPM"

5011           .Range("I17").Select
5012           .ActiveCell.FormulaR1C1 = "Calculated"
5013           .Range("I18").Select
5014           .ActiveCell.FormulaR1C1 = "RPM"

5015           .Range("J17").Select
5016           .ActiveCell.FormulaR1C1 = "Suction"
5017           .Range("J18").Select
5018           .ActiveCell.FormulaR1C1 = "Temp(F)"

5019           .Range("K17").Select
5020           .ActiveCell.FormulaR1C1 = "Disch"
5021           .Range("K18").Select
5022           .ActiveCell.FormulaR1C1 = "Pressure"

5023           .Range("L17").Select
5024           .ActiveCell.FormulaR1C1 = "Suction"
5025           .Range("L18").Select
5026           .ActiveCell.FormulaR1C1 = "Pressure"

5027           .Range("M17").Select
5028           .ActiveCell.FormulaR1C1 = "Vel"
5029           .Range("M18").Select
5030           .ActiveCell.FormulaR1C1 = "Head"

5031           .Range("N17").Select
5032           .ActiveCell.FormulaR1C1 = "Axial"
5033           .Range("N18").Select
5034           .ActiveCell.FormulaR1C1 = "Position"

5035           .Range("O17").Select
5036           .ActiveCell.FormulaR1C1 = "Pct of"
5037           .Range("O18").Select
5038           .ActiveCell.FormulaR1C1 = "End Play"

5039           .Range("P17").Select
5040           .ActiveCell.FormulaR1C1 = "Hydraulic"
5041           .Range("P18").Select
5042           .ActiveCell.FormulaR1C1 = "Efficiency"

       '        .Range("P17").Select
       '        .ActiveCell.FormulaR1C1 = "Circ"
       '        .Range("P18").Select
       '        .ActiveCell.FormulaR1C1 = "Flow"

5043           .Range("Q17").Select
5044           .ActiveCell.FormulaR1C1 = "Motor"
5045           .Range("Q18").Select
5046           .ActiveCell.FormulaR1C1 = "Efficiency"

5047           .Range("S17").Select
5048           .ActiveCell.FormulaR1C1 = "NPSHa"

5049           .Range("T17").Select
5050           .ActiveCell.FormulaR1C1 = "Phase 1"
5051           .Range("T18").Select
5052           .ActiveCell.FormulaR1C1 = "Current"

5053           .Range("U17").Select
5054           .ActiveCell.FormulaR1C1 = "Phase 2"
5055           .Range("U18").Select
5056           .ActiveCell.FormulaR1C1 = "Current"

5057           .Range("V17").Select
5058           .ActiveCell.FormulaR1C1 = "Phase 3"
5059           .Range("V18").Select
5060           .ActiveCell.FormulaR1C1 = "Current"

5061           .Range("W17").Select
5062           .ActiveCell.FormulaR1C1 = "Phase 1"
5063           .Range("W18").Select
5064           .ActiveCell.FormulaR1C1 = "Voltage"

5065           .Range("X17").Select
5066           .ActiveCell.FormulaR1C1 = "Phase 2"
5067           .Range("X18").Select
5068           .ActiveCell.FormulaR1C1 = "Voltage"

5069           .Range("Y17").Select
5070           .ActiveCell.FormulaR1C1 = "Phase 3"
5071           .Range("Y18").Select
5072           .ActiveCell.FormulaR1C1 = "Voltage"

5073           .Range("Z17").Select
5074           .ActiveCell.FormulaR1C1 = "'" & txtTitle(20).Text

5075           .Range("Z18").Select
5076           .ActiveCell.FormulaR1C1 = "'" & txtTitle(21).Text

5077           .Range("AA17").Select
5078           .ActiveCell.FormulaR1C1 = "'" & txtTitle(22).Text

5079           .Range("AA18").Select
5080           .ActiveCell.FormulaR1C1 = "'" & txtTitle(23).Text

5081           .Range("AB17").Select
5082           .ActiveCell.FormulaR1C1 = "'" & txtTitle(24).Text

5083           .Range("AB18").Select
5084           .ActiveCell.FormulaR1C1 = "'" & txtTitle(25).Text

5085           .Range("AC17").Select
5086           .ActiveCell.FormulaR1C1 = "HR"

5087           .Range("AC18").Select
5088           .ActiveCell.FormulaR1C1 = "(ft)"

5089           .Range("AD17").Select
5090           .ActiveCell.FormulaR1C1 = "'" & txtTitle(26).Text

5091           .Range("AD18").Select
5092           .ActiveCell.FormulaR1C1 = "'" & txtTitle(27).Text

5093           .Range("AE17").Select
5094           .ActiveCell.FormulaR1C1 = "TRG"
5095           .Range("AE18").Select
5096           .ActiveCell.FormulaR1C1 = "Position"

5097           .Range("AF17").Select
5098           .ActiveCell.FormulaR1C1 = "Thrust"

5099           .Range("AG17").Select
5100           .ActiveCell.FormulaR1C1 = "F/R"

5101           .Range("AH17").Select
5102           .ActiveCell.FormulaR1C1 = "Moment"
5103           .Range("AH18").Select
5104           .ActiveCell.FormulaR1C1 = "Arm"

5105           .Range("AI17").Select
5106           .ActiveCell.FormulaR1C1 = "Rig"
5107           .Range("AI18").Select
5108           .ActiveCell.FormulaR1C1 = "Pressure"

       '        .Range("AI17").Select
       '        .ActiveCell.FormulaR1C1 = "Viscosity"

5109           .Range("AJ19").Select
5110           .ActiveCell.FormulaR1C1 = "Rear"
5111           .Range("AJ18").Select
5112           .ActiveCell.FormulaR1C1 = "Force"

5113           .Range("AK17").Select
5114           .ActiveCell.FormulaR1C1 = "PV"

5115           .Range("R17").Select
5116           .ActiveCell.FormulaR1C1 = "Shaft"
5117           .Range("R18").Select
5118           .ActiveCell.FormulaR1C1 = "Power"

       '        .Range("AM17").Select
       '        .ActiveCell.FormulaR1C1 = "Pct Full"
       '        .Range("AM18").Select
       '        .ActiveCell.FormulaR1C1 = "Scale"

5119           .Range("AL17").Select
5120           .ActiveCell.FormulaR1C1 = "NPSHr"

5121           .Range("AM17").Select
5122           .ActiveCell.FormulaR1C1 = "Remarks"




               'now output the data

5123           iRowNo = 20

5124           rsEff.MoveFirst
5125           For I = 1 To frmPLCData.UpDown2.value
5126               .Range("A" & iRowNo).Select
5127               .ActiveCell.FormulaR1C1 = rsEff.Fields("Flow")

5128               .Range("B" & iRowNo).Select
5129               .ActiveCell.FormulaR1C1 = rsEff.Fields("TDH")

5130               .Range("C" & iRowNo).Select
5131               .ActiveCell.FormulaR1C1 = rsEff.Fields("KW")

5132               .Range("D" & iRowNo).Select
5133               .ActiveCell.FormulaR1C1 = rsEff.Fields("Volts")

5134               .Range("E" & iRowNo).Select
5135               .ActiveCell.FormulaR1C1 = rsEff.Fields("Amps")

5136               .Range("F" & iRowNo).Select
5137               .ActiveCell.FormulaR1C1 = rsEff.Fields("PowerFactor")

5138               .Range("G" & iRowNo).Select
5139               .ActiveCell.FormulaR1C1 = rsEff.Fields("OverallEfficiency")

5140               .Range("H" & iRowNo).Select
5141               .ActiveCell.FormulaR1C1 = rsEff.Fields("RPM")

5142               .Range("I" & iRowNo).Select
                   'use the coefficients from above to calculate rpm
5143               Dim f As Double
5144               f = .Range("H6").value
5145               .ActiveCell.FormulaR1C1 = (Val(f) / 60) * (ACoef * (rsEff.Fields("KW")) ^ 2 + BCoef * (rsEff.Fields("KW")) + CCoef)

5146               .Range("J" & iRowNo).Select
5147               .ActiveCell.FormulaR1C1 = rsEff.Fields("Temperature")

5148               .Range("K" & iRowNo).Select
5149               .ActiveCell.FormulaR1C1 = rsEff.Fields("DischPress")

5150               .Range("L" & iRowNo).Select
5151               .ActiveCell.FormulaR1C1 = rsEff.Fields("SuctPress")

5152               .Range("M" & iRowNo).Select
5153               .ActiveCell.FormulaR1C1 = rsEff.Fields("VelocityHead")

5154               .Range("N" & iRowNo).Select
5155               .ActiveCell.FormulaR1C1 = rsEff.Fields("Pos")

5156               .Range("O" & iRowNo).Select
5157               .ActiveCell.FormulaR1C1 = 100 * rsEff.Fields("Pos") / Val(txtEndPlay)

5158               .Range("P" & iRowNo).Select
5159               .ActiveCell.FormulaR1C1 = rsEff.Fields("HydraulicEfficiency")

       '            .Range("P" & iRowNo).Select
       '            .ActiveCell.FormulaR1C1 = rsEff.Fields("CircFlow")

5160               .Range("Q" & iRowNo).Select
5161               .ActiveCell.FormulaR1C1 = rsEff.Fields("MotorEfficiency")

5162               .Range("S" & iRowNo).Select
5163               .ActiveCell.FormulaR1C1 = rsEff.Fields("NPSHa")

5164               .Range("T" & iRowNo).Select
5165               .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentA")

5166               .Range("U" & iRowNo).Select
5167               .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentB")

5168               .Range("V" & iRowNo).Select
5169               .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentC")

5170               .Range("W" & iRowNo).Select
5171               .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageA")

5172               .Range("X" & iRowNo).Select
5173               .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageB")

5174               .Range("Y" & iRowNo).Select
5175               .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageC")

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

5176               .Range("Z" & iRowNo).Select
5177               .ActiveCell.FormulaR1C1 = rsEff.Fields("CircFlow")

5178               .Range("AA" & iRowNo).Select
5179               .ActiveCell.FormulaR1C1 = rsEff.Fields("RBHTemp")

5180               .Range("AB" & iRowNo).Select
5181               .ActiveCell.FormulaR1C1 = rsEff.Fields("RBHPress")

5182               .Range("AC" & iRowNo).Select
5183               .ActiveCell.FormulaR1C1 = (rsEff.Fields("RBHPress") - rsEff.Fields("SuctPress")) * 2.31

5184               .Range("AD" & iRowNo).Select
5185               .ActiveCell.FormulaR1C1 = rsEff.Fields("AI4")

5186               .Range("AE" & iRowNo).Select
5187               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCTRG")

5188               .Range("AF" & iRowNo).Select
5189               If rsEff.Fields("TEMCFrontThrust") = 0 Then
5190                   If rsEff.Fields("TEMCRearThrust") = 0 Then
5191                       .ActiveCell.FormulaR1C1 = " "
5192                       .Range("AG" & iRowNo).Select
5193                       .ActiveCell.FormulaR1C1 = " "
5194                   Else
5195                       .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCRearThrust")
5196                       .Range("AG" & iRowNo).Select
5197                       .ActiveCell.FormulaR1C1 = "R"
5198                   End If
5199               Else
5200                   .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCFrontThrust")
5201                   .Range("AG" & iRowNo).Select
5202                   .ActiveCell.FormulaR1C1 = "F"
5203               End If

5204               .Range("AH" & iRowNo).Select
5205               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCMomentArm")

5206               .Range("AI" & iRowNo).Select
5207               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCThrustRigPressure")

       '            .Range("AJ" & iRowNo).Select
       '            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCViscosity")

5208               .Range("AJ" & iRowNo).Select
5209               If rsEff.Fields("TEMCForceDirection") = "F" Then
5210                   .ActiveCell.FormulaR1C1 = -rsEff.Fields("TEMCCalculatedForce")
5211               Else
5212                   .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCCalculatedForce")
5213               End If

5214               .Range("AK" & iRowNo).Select
5215               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCPV")

5216               .Range("R" & iRowNo).Select
5217               .ActiveCell.FormulaR1C1 = rsEff.Fields("KW") * rsEff.Fields("MotorEfficiency") / 100

5218               .Range("AL" & iRowNo).Select
5219               .ActiveCell.FormulaR1C1 = rsEff.Fields("NPSHr")

       '            If RatedKW = 999 Then
       '                .ActiveCell.FormulaR1C1 = ""
       '            Else
       '                .ActiveCell.FormulaR1C1 = (rsEff.Fields("KW") * rsEff.Fields("MotorEfficiency")) / (1 * RatedKW)
       '            End If

5220               .Range("AM" & iRowNo).Select
5221               .ActiveCell.FormulaR1C1 = rsEff.Fields("Remarks")


5222               rsEff.MoveNext
5223               iRowNo = iRowNo + 1
5224           Next I

5225           .Range("A20:AS30").Select
5226           .Selection.NumberFormat = "0.00"

           'set up formulas to calculate BEP
           '  first, plot 2nd order polynomial for flow vs hydraulic efficiency
           '  the formulas for doing that are in E68, F68 and G68
           '  only want the formulas to point to the number of points in the test data, so use frmPLCData.CWNumEdit2.value
           '
5227       Dim AColumnRow As String
5228       Dim PColumnRow As String

5229       AColumnRow = "A" & Trim(str(19 + frmPLCData.UpDown2.value))
5230       PColumnRow = "P" & Trim(str(19 + frmPLCData.UpDown2.value))

5231           .Range("E68").Select
5232           .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1)"

5233           .Range("F68").Select
5234           .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1,2)"

5235           .Range("G68").Select
5236           .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1,3)"

           'export balance holes
5237       If boGotBalanceHoles Then
5238           If rsBalanceHoles.State = adStateClosed Then
5239               rsBalanceHoles.ActiveConnection = cnPumpData
5240               rsBalanceHoles.Open
5241           End If 'rsBalanceHoles.State = adStateClosed

5242           If rsBalanceHoles.RecordCount <> 0 Then

5243               .Range("K9:N9").Merge
5244               .Range("K9:N9").Formula = "Balance Hole Data"
5245               .Range("K9:N9").HorizontalAlignment = xlCenter

5246               .Range("K10").Select
5247               .ActiveCell.Formula = "Date"

5248               .Range("L10").Select
5249               .ActiveCell.Formula = "Number"

5250               .Range("M10").Select
5251               .ActiveCell.Formula = "Diameter"

5252               .Range("N10").Select
5253               .ActiveCell.Formula = "Bolt Circle"

5254               iRowNo = 11

5255               If rsBalanceHoles.RecordCount > 3 Then
5256                   For I = 1 To rsBalanceHoles.RecordCount - 3
5257                       Rows("13:13").Select
5258                       Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
5259                   Next I
5260               End If

5261               rsBalanceHoles.MoveFirst
5262               For I = 1 To rsBalanceHoles.RecordCount

5263                   .Range("K" & iRowNo).Select
5264                   .ActiveCell.Formula = rsBalanceHoles.Fields("Date")
5265                   .ActiveCell.NumberFormat = "m/d/yy h:mm AM/PM;@"
5266                   .Range("L" & iRowNo).Select
5267                   .ActiveCell = rsBalanceHoles.Fields("Number")
5268                   .ActiveCell.NumberFormat = "0"
5269                   .Range("M" & iRowNo).Select
5270                   If IsNumeric(rsBalanceHoles.Fields("Diameter1")) Then
5271                       .ActiveCell = Val(rsBalanceHoles.Fields("Diameter1"))
5272                       .ActiveCell.NumberFormat = "0.0000"
5273                   Else
5274                       .ActiveCell = rsBalanceHoles.Fields("Diameter1")
5275                   End If

5276                   .Range("N" & iRowNo).Select
5277                   If IsNumeric(rsBalanceHoles.Fields("BoltCircle1")) Then
5278                       .ActiveCell = Val(rsBalanceHoles.Fields("BoltCircle1"))
5279                       .ActiveCell.NumberFormat = "0.0000"
5280                   Else
5281                       .ActiveCell = rsBalanceHoles.Fields("BoltCircle1")
5282                   End If

5283                   rsBalanceHoles.MoveNext
5284                   iRowNo = iRowNo + 1
5285               Next I
5286               .Range("K10:N" & iRowNo - 1).Select
5287               With .Selection.Interior
5288                   .ColorIndex = 34
5289                   .Pattern = xlSolid
5290               End With
5291           End If 'rsBalanceHoles.RecordCount <> 0
5292       End If ' boGotBalanceHoles

           'plot graphs

5293       Dim SeriesName As String
5294       Dim XVals As String
5295       Dim YVals As String
5296       Dim RowNo As Long
5297       Dim RowStr As String
5298       Dim LastPoint As Integer
5299       Dim LineType As String
5300       Dim AxisGroup As Integer
5301       Dim LabelPos As Integer
5302       Dim LineColor As Long

5303           .ActiveSheet.ChartObjects("HydRepChart").Activate
5304           Dim S As Series
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
5305           Dim aq As Double
5306           Range("AQ56", "AQ71").Select
5307           aq = .Max(Selection)
5308           Dim ax As Double
5309           Range("AX56", "AX71").Select
5310           ax = .Max(Selection)

               'then current (as and az)
5311           Dim at As Double
5312           Range("AS56", "AS71").Select
5313           at = .Max(Selection)
5314           Dim ba As Double
5315           Range("AZ56", "AZ71").Select
5316           ba = .Max(Selection)

5317           Dim CurrentScaleMax As Integer
5318           Dim TDHScaleMax As Integer

5319           Dim MaxTDH As Integer
5320           With Application.WorksheetFunction
5321               If aq > ax Then
5322                   MaxTDH = .Ceiling(aq, 25)
5323               Else
5324                   MaxTDH = .Ceiling(ax, 25)
5325               End If
5326           End With

5327           Dim MaxCurrent As Integer
5328           With Application.WorksheetFunction
5329               If at > ba Then
5330                   Select Case at
                           Case Is <= 5
5331                           CurrentScaleMax = 5

5332                       Case Is <= 10
5333                           CurrentScaleMax = 10

5334                       Case Else
5335                           CurrentScaleMax = 25
5336                   End Select

5337                   MaxCurrent = .Ceiling(at, CurrentScaleMax)
5338               Else
5339                  Select Case ba
                           Case Is <= 5
5340                           CurrentScaleMax = 5

5341                       Case Is <= 10
5342                           CurrentScaleMax = 10

5343                       Case Else
5344                           CurrentScaleMax = 25
5345                   End Select

5346                   MaxCurrent = .Ceiling(ba, CurrentScaleMax)
5347               End If
5348           End With

5349           ActiveSheet.ChartObjects("HydRepChart").Activate
5350            Dim ShtName As String
5351            ShtName = "'" & ActiveSheet.Name & "'"

5352           RowStr = 56 + 15
5353            For I = 1 To 8

5354                Select Case I
                        Case 1
5355                        SeriesName = "=""TDH"""
5356                        XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
5357                        YVals = "=" & ShtName & "!$AQ$56:$AQ$" & RowStr
5358                        LineType = msoLineSolid
5359                        AxisGroup = 1
5360                        LabelPos = xlLabelPositionRight
5361                        LineColor = vbBlue

5362                    Case 2
5363                        SeriesName = "=""Input Power"""
5364                        XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
5365                        YVals = "=" & ShtName & "!$AR$56:$AR$" & RowStr
5366                        LineType = msoLineSolid
5367                        AxisGroup = 2
5368                        LabelPos = xlLabelPositionRight
5369                        LineColor = vbRed

5370                    Case 3
5371                        SeriesName = "=""Current"""
5372                        XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
5373                        YVals = "=" & ShtName & "!$AS$56:$AS$" & RowStr
5374                        LineType = msoLineSolid
5375                        AxisGroup = 2
5376                        LabelPos = xlLabelPositionRight
5377                        LineColor = vbGreen

5378                    Case 4
       '                     SeriesName = "=""Overall Eff"""
       '                     XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
       '                     YVals = "=" & ShtName & "!$AT$56:$AT$" & RowStr
       '                     LineType = msoLineSolid
       '                     AxisGroup = 2
       '                     LabelPos = xlLabelPositionRight
       '                     LineColor = vbCyan

5379                    Case 5
5380                        SeriesName = "=""TDH (Adj)"""
5381                        XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
5382                        YVals = "=" & ShtName & "!$AX$56:$AX$" & RowStr
5383                        LineType = msoLineDash
5384                        AxisGroup = 1
5385                        LabelPos = xlLabelPositionBelow
5386                        LineColor = vbBlue

5387                    Case 6
5388                        SeriesName = "=""Input Power (Adj)"""
5389                        XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
5390                        YVals = "=" & ShtName & "!$AY$56:$AY$" & RowStr
5391                        LineType = msoLineDash
5392                        AxisGroup = 2
5393                        LabelPos = xlLabelPositionBelow
5394                        LineColor = vbRed

5395                    Case 7
5396                        SeriesName = "=""Current (Adj)"""
5397                        XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
5398                        YVals = "=" & ShtName & "!$AZ$56:$AZ$" & RowStr
5399                        LineType = msoLineDash
5400                        AxisGroup = 2
5401                        LabelPos = xlLabelPositionBelow
5402                        LineColor = vbGreen

5403                    Case 8
       '                     SeriesName = "=""Overall Eff (Adj)"""
       '                     XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
       '                     YVals = "=" & ShtName & "!$BA$56:$BA$" & RowStr
       '                     LineType = msoLineDash
       '                     AxisGroup = 2
       '                     LabelPos = xlLabelPositionBelow
       '                     LineColor = vbCyan

5404               End Select
5405               LastPoint = 16
5406               ActiveChart.SeriesCollection.NewSeries
5407               ActiveChart.SeriesCollection(I).Name = SeriesName
5408               ActiveChart.SeriesCollection(I).XValues = XVals
5409               ActiveChart.SeriesCollection(I).Values = YVals
5410               ActiveChart.SeriesCollection(I).Select
5411               ActiveChart.SeriesCollection(I).Points(LastPoint).Select
5412               ActiveChart.SeriesCollection(I).Points(LastPoint).ApplyDataLabels
5413               ActiveChart.SeriesCollection(I).Points(LastPoint).DataLabel.Select
5414               If I < 5 Then
5415                   Selection.ShowSeriesName = True
5416                   Selection.Position = LabelPos
5417               Else
5418                   Selection.ShowSeriesName = False
5419               End If
5420               Selection.ShowValue = False
5421               ActiveChart.SeriesCollection(I).ChartType = xlXYScatterSmoothNoMarkers
5422               ActiveChart.SeriesCollection(I).Select
5423               With Selection.Format.line
5424                   .Visible = msoTrue
5425                   .DashStyle = LineType
5426                   .ForeColor.RGB = LineColor
5427               End With


5428               ActiveChart.SeriesCollection(I).AxisGroup = AxisGroup
5429               ActiveChart.SeriesCollection(I).DataLabels.Font.Size = 8
5430               ActiveChart.SeriesCollection(I).DataLabels.Font.Name = "Arial"
5431           Next I

               'show design point
5432           SeriesName = "=""Design Point"""
5433           XVals = "=" & ShtName & "!$L$63"
5434           YVals = "=" & ShtName & "!$L$64"
5435           LineType = msoLineSolid
5436           AxisGroup = 1
5437           ActiveChart.SeriesCollection.NewSeries
5438           ActiveChart.SeriesCollection(I).Name = SeriesName
5439           ActiveChart.SeriesCollection(I).XValues = XVals
5440           ActiveChart.SeriesCollection(I).Values = YVals
5441           ActiveChart.SeriesCollection(I).Select

5442           Selection.MarkerStyle = 4
5443           Selection.MarkerSize = 7
5444           With Selection.Format.line
5445               .Visible = msoTrue
5446               .Weight = 2.25
5447               .ForeColor.RGB = vbBlack
5448           End With


5449           ActiveChart.Axes(xlValue).Select
5450           ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
5451           ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True

5452           ActiveChart.Axes(xlValue).MaximumScale = MaxTDH
5453           ActiveChart.Axes(xlValue).MinimumScale = 0
5454           ActiveChart.Axes(xlValue).MajorUnit = Int(MaxTDH / 5)
5455           Selection.TickLabels.NumberFormat = "0"

5456           ActiveChart.Axes(xlValue, xlSecondary).Select
5457           ActiveChart.Axes(xlValue, xlSecondary).MinimumScaleIsAuto = True
5458           ActiveChart.Axes(xlValue, xlSecondary).MaximumScaleIsAuto = True

5459           ActiveChart.Axes(xlValue, xlSecondary).MaximumScale = MaxCurrent
5460           ActiveChart.Axes(xlValue, xlSecondary).MinimumScale = 0
5461           ActiveChart.Axes(xlValue, xlSecondary).MajorUnit = Int(MaxCurrent / 5)
5462           Selection.TickLabels.NumberFormat = "0"

5463           ActiveChart.Axes(xlValue, xlSecondary).HasTitle = True
5464           ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = "Input Power (kW)-Current (A)"
       '        ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = "Input Power (kW)-Current (A)-Overall Efficiency (%)"
5465           ActiveChart.SetElement (msoElementSecondaryValueAxisTitleRotated)
               'ActiveSheet.PageSetup.PrintArea = "$CA$1:$CI$50"

5466           Range("A1").Select

               'delete all macros in the excel file

               ' Declare variables to access the macros in the workbook.
5467           Dim objProject As VBIDE.VBProject
5468           Dim objComponent As VBIDE.VBComponent
5469           Dim objCode As VBIDE.CodeModule

               ' Get the project details in the workbook.
5470           Set objProject = xlBook.VBProject

               ' Iterate through each component in the project.
5471           For Each objComponent In objProject.VBComponents

                   ' Delete code modules
5472               Set objCode = objComponent.CodeModule
5473               objCode.DeleteLines 1, objCode.CountOfLines

5474               Set objCode = Nothing
5475               Set objComponent = Nothing
5476           Next

5477           Set objProject = Nothing


5478           xlApp.Visible = True                    'show the sheet

5479           xlApp.VBE.ActiveVBProject.VBComponents.Import ParentDirectoryName & sSaveFileMacroFile
5480           xlApp.Run "AssignButton"
5481       End With

       '    Exit Sub

5482   ErrHandler:
           'User pressed the Cancel button

5483       On Error GoTo notopen
5484       If Not xlApp.ActiveWorkbook Is Nothing Then
5485           ActiveWorkbook.CheckCompatibility = False
5486           xlApp.ActiveWorkbook.Save               'save the workbook
               'xlApp.ActiveWorkbook.Close

5487       End If

5488   notopen:

       '    xlApp.Application.Quit

       '    xlApp.Quit
       '    Set xlApp = Nothing

       '    If CommonDialog1.filename <> "" Then
       '        MsgBox CommonDialog1.filename & " has been written.", vbOKOnly, "File Opened"
       '    End If

5489   On Error GoTo vbwErrHandler

' <VB WATCH>
5490       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5491       Exit Sub
' <VB WATCH>
5492       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5493       Exit Sub
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
5494       On Error GoTo vbwErrHandler
5495       Const VBWPROCNAME = "frmPLCData.GetWorksheetTabs"
5496       If vbwProtector.vbwTraceProc Then
5497           Dim vbwProtectorParameterString As String
5498           If vbwProtector.vbwTraceParameters Then
5499               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("filename", filename) & ", "
5500               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("WorkSheetName", WorkSheetName) & ") "
5501           End If
5502           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5503       End If
' </VB WATCH>

           'see what worksheet tabs alread exist in the excel worksheet

5504       Dim intSheets As Integer    'number of sheets in the workbook
5505       Dim I As Integer
5506       Dim S As String
5507       Dim ans As Integer
5508       Dim NameOK As Boolean

5509       intSheets = xlApp.Worksheets.Count      'how many sheets are there?

           'define a crlf string
5510       S = vbCrLf

5511       For I = 1 To intSheets
5512           S = S & xlApp.Worksheets(I).Name & vbCrLf   'add in the worksheet name
5513       Next I

           'tell the user the names so far and ask if he/she wants to add another
5514       ans = MsgBox("You have the following Worksheet Names in " & filename & ": " & S & "Do you want to add another sheet to this file?", vbYesNo, "Sheets in Excel File")

           'get the answer
5515       If ans = vbNo Then
5516           GetWorksheetTabs = vbNo     'set up flag for when we return to the calling subroutine
' <VB WATCH>
5517       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5518           Exit Function
5519       End If

           'get worksheet name from user and check to see that it's not already used

5520       NameOK = False  'start assuming that the name is bad

5521       While Not NameOK    'as long as it's bad, stay in this loop
5522           WorkSheetName = InputBox("Enter Worksheet Name for this run.")  'ask for name

5523           If WorkSheetName = "" Then      'if we get a nul return or user presses cancel
5524               GetWorksheetTabs = vbNo
' <VB WATCH>
5525       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5526               Exit Function
5527           End If

5528           For I = 1 To xlApp.Worksheets.Count     'go through all of the existing sheets
5529               If WorkSheetName = xlApp.Worksheets(I).Name Then        'if the names are the same
5530                   MsgBox "The name " & WorkSheetName & " already exists for a Worksheet.  Please try again.", vbOKOnly, "Bad Worksheet Name"  'tell the user
5531                   NameOK = False
5532                   Exit For
5533               End If
5534               NameOK = True       'if we make it thru say the name is ok
5535           Next I
5536       Wend

5537       xlApp.Worksheets.Add , xlApp.Worksheets(xlApp.Worksheets.Count)     'add a worksheer
5538       xlApp.Worksheets(xlApp.Worksheets.Count).Name = WorkSheetName       'give it the desired name
5539       GetWorksheetTabs = vbYes                                            'say that the results were ok

' <VB WATCH>
5540       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5541       Exit Function
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
5542       On Error GoTo vbwErrHandler
5543       Const VBWPROCNAME = "frmPLCData.NewWorkBook"
5544       If vbwProtector.vbwTraceProc Then
5545           Dim vbwProtectorParameterString As String
5546           If vbwProtector.vbwTraceParameters Then
5547               vbwProtectorParameterString = "()"
5548           End If
5549           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5550       End If
' </VB WATCH>

5551       Dim WorkSheetName As String

           'we've just added a new workbook, delete sheet1, sheet2, etc
5552       xlApp.DisplayAlerts = False
5553       While xlApp.Worksheets.Count > 1
5554           xlApp.Worksheets(1).Delete          'delete the sheet
5555       Wend
5556       xlApp.DisplayAlerts = True

5557       WorkSheetName = InputBox("Enter Title Worksheet Name for this run.")    'get the desired name
5558       xlApp.Worksheets(1).Name = WorkSheetName    'and name the sheet

5559       NewWorkBook = WorkSheetName

' <VB WATCH>
5560       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5561       Exit Function
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
5562       On Error GoTo vbwErrHandler
5563       Const VBWPROCNAME = "frmPLCData.CalibrateSoftware"
5564       If vbwProtector.vbwTraceProc Then
5565           Dim vbwProtectorParameterString As String
5566           If vbwProtector.vbwTraceParameters Then
5567               vbwProtectorParameterString = "()"
5568           End If
5569           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5570       End If
' </VB WATCH>
5571           frmCalibrate.Show
               'Calibrating = True

' <VB WATCH>
5572       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5573       Exit Sub
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
5574       On Error GoTo vbwErrHandler
5575       Const VBWPROCNAME = "frmPLCData.ParseTEMCModelNo"
5576       If vbwProtector.vbwTraceProc Then
5577           Dim vbwProtectorParameterString As String
5578           If vbwProtector.vbwTraceParameters Then
5579               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
5580               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("ltr", ltr) & ") "
5581           End If
5582           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5583       End If
' </VB WATCH>
5584       Dim I As Integer
5585       Dim iStart As Integer
5586       Dim iStop As Integer
5587       Dim strCompare As String

5588       For I = 0 To cmbComboName.ListCount - 1                     'go through the combobox entries
5589           iStart = InStr(1, cmbComboName.List(I), "[")
5590           iStop = InStr(1, cmbComboName.List(I), "]")
5591           strCompare = Mid$(cmbComboName.List(I), iStart + 1, iStop - iStart - 1)
5592           If UCase(strCompare) = UCase(ltr) Then   'see when we find the desired index number
5593               cmbComboName.ListIndex = I                                              'if we do, set the combo box
5594               Exit For                                            'and we're done
5595           End If
       '        cmbComboName.ListIndex = -1                             'else, remove any pointer
5596           cmbComboName.ListIndex = cmbComboName.ListCount - 1                           'else, remove any pointer
5597       Next I

5598       txtModelNo.Text = UCase(txtModelNo.Text)
5599       txtModelNo.SelStart = Len(txtModelNo.Text)
' <VB WATCH>
5600       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5601       Exit Function
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
5602       On Error GoTo vbwErrHandler
5603       Const VBWPROCNAME = "frmPLCData.LoadCombo"
5604       If vbwProtector.vbwTraceProc Then
5605           Dim vbwProtectorParameterString As String
5606           If vbwProtector.vbwTraceParameters Then
5607               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("cmbComboName", cmbComboName) & ", "
5608               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("sTableName", sTableName) & ") "
5609           End If
5610           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5611       End If
' </VB WATCH>

5612       Dim I As Integer
5613       Dim sItem As String
5614       Dim iID As Integer
5615       Dim bUseDropdown As Boolean
5616       Dim qy As New ADODB.Command
5617       Dim rs As New ADODB.Recordset

       '    rsPumpParameters.CursorLocation = adUseClient
       '    If sTableName = "Model" Then
       '        rsPumpParameters.Sort = "Model"
       '    Else
       '        rsPumpParameters.Sort = vbNullString
       '    End If
       '    rsPumpParameters.Open sTableName, cnPumpData, adOpenStatic, adLockOptimistic, adCmdTableDirect

5618       qy.ActiveConnection = cnPumpData
5619       If sTableName = "DischargeDiameter" Or sTableName = "SuctionDiameter" Then
5620           qy.CommandText = "SELECT * FROM " & sTableName & " ORDER BY Val(Description)"
5621       Else
5622           qy.CommandText = "SELECT * FROM " & sTableName & " ORDER BY Description"
5623       End If
5624       If sTableName = "SupermarketPumpData" Then
5625           qy.CommandText = "SELECT ID,Model AS Description FROM " & sTableName
5626       End If
5627       rs.CursorLocation = adUseClient
5628       rs.CursorType = adOpenStatic

5629       rs.Open qy


5630       On Error GoTo NoField
5631       bUseDropdown = True
           'sItem = rsPumpParameters.Fields("UseInDropdown")
       '    If bUseDropdown Then
       '        rsPumpParameters.Sort = "Description"
       '    End If
5632       rs.MoveFirst                                'goto the top
5633       For I = 0 To rs.RecordCount - 1             'go through the whole recordset
5634           sItem = rs.Fields("Description")        'get the description
5635           iID = rs.Fields(0)                      'get the index number - primary key
5636           If bUseDropdown Then
       '            If rsPumpParameters.Fields("UseInDropdown").value = True Then
5637                   cmbComboName.AddItem sItem, I                                   'add the description to the combo box
       '                cmbComboName.AddItem sItem                                   'add the description to the combo box
5638                   cmbComboName.ItemData(cmbComboName.NewIndex) = iID              'add the key number into the item data
       '            End If
5639           End If
5640           rs.MoveNext                             'get the next record
5641       Next I
5642       rs.Close
5643       cmbComboName.ListIndex = -1
5644   On Error GoTo vbwErrHandler
5645       Set rs = Nothing
5646       Set qy = Nothing
' <VB WATCH>
5647       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5648       Exit Function

5649   NoField:
5650       bUseDropdown = False
5651   On Error GoTo vbwErrHandler
5652       Resume Next

' <VB WATCH>
5653       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5654       Exit Function
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
5655       On Error GoTo vbwErrHandler
5656       Const VBWPROCNAME = "frmPLCData.SetGraphMax"
5657       If vbwProtector.vbwTraceProc Then
5658           Dim vbwProtectorParameterString As String
5659           If vbwProtector.vbwTraceParameters Then
5660               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("Plothead", Plothead) & ") "
5661           End If
5662           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5663       End If
' </VB WATCH>

5664       Dim I As Integer
5665       Dim m As Single

5666       m = 0
5667       For I = 0 To UBound(Plothead, 2)
5668           If Plothead(1, I) > m Then
5669               m = Plothead(1, I)
5670           End If
5671       Next I
5672       SetGraphMax = 10 * (Int((m / 10) + 0.5) + 1)
5673       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
5674       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10 * (Int((m / 10) + 0.5) + 1)
5675       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0

' <VB WATCH>
5676       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5677       Exit Function
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
5678       On Error GoTo vbwErrHandler
5679       Const VBWPROCNAME = "frmPLCData.CalculateSpeed"
5680       If vbwProtector.vbwTraceProc Then
5681           Dim vbwProtectorParameterString As String
5682           If vbwProtector.vbwTraceParameters Then
5683               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("CoefSq", CoefSq) & ", "
5684               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("CoefLin", CoefLin) & ", "
5685               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("CoefConstant", CoefConstant) & ", "
5686               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("InputHP", InputHP) & ", "
5687               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("SG", SG) & ") "
5688           End If
5689           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5690       End If
' </VB WATCH>
5691       Dim I As Integer
5692       Dim OldResult As Double
5693       Dim NewResult As Double

5694       CalculateSpeed = 0

5695       If SG > 5 Or SG < 0.01 Then
5696           MsgBox "Bad value for SG...must be between 0.01 and 5.", vbOKOnly, "Bad SG Value"
' <VB WATCH>
5697       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5698           Exit Function
5699       End If

5700       OldResult = 1000
5701       NewResult = 0

5702       I = 1

5703       Do While Abs(NewResult - OldResult) > 0.1
5704           ReDim Preserve results(I)
5705           Select Case I
                   Case 1
5706                   results(I - 1).HP = InputHP
5707               Case 2
5708                   results(I - 1).HP = results(I - 2).HP * SG
5709               Case Else
5710                   results(I - 1).HP = results(I - 2).HP * (results(I - 2).Speed / results(I - 3).Speed) ^ 3
5711           End Select
5712           OldResult = NewResult
5713           results(I - 1).Speed = CalcPoly(CoefSq, CoefLin, CoefConstant, results(I - 1).HP)
5714           NewResult = results(I - 1).Speed
5715           If I > 15 Then
5716               If I = 0 Or I > 15 Then
5717                   MsgBox "Over 15 calculations and no convergence", vbOKOnly, "Too many iterations"
' <VB WATCH>
5718       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5719                   Exit Function
5720               End If
' <VB WATCH>
5721       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5722               Exit Function
5723           End If
5724           I = I + 1
5725       Loop
5726       CalculateSpeed = I - 1
' <VB WATCH>
5727       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5728       Exit Function
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
5729       On Error GoTo vbwErrHandler
5730       Const VBWPROCNAME = "frmPLCData.CalcPoly"
5731       If vbwProtector.vbwTraceProc Then
5732           Dim vbwProtectorParameterString As String
5733           If vbwProtector.vbwTraceParameters Then
5734               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("CoefSq", CoefSq) & ", "
5735               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("CoefLin", CoefLin) & ", "
5736               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("CoefConstant", CoefConstant) & ", "
5737               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("DataIn", DataIn) & ") "
5738           End If
5739           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5740       End If
' </VB WATCH>
5741       CalcPoly = CoefSq * DataIn ^ 2 + CoefLin * DataIn + CoefConstant
' <VB WATCH>
5742       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5743       Exit Function
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
5744       On Error GoTo vbwErrHandler
5745       Const VBWPROCNAME = "frmPLCData.GetBalanceHoleData"
5746       If vbwProtector.vbwTraceProc Then
5747           Dim vbwProtectorParameterString As String
5748           If vbwProtector.vbwTraceParameters Then
5749               vbwProtectorParameterString = "(" & vbwProtector.vbwReportParameter("SerialNumber", SerialNumber) & ", "
5750               vbwProtectorParameterString = vbwProtectorParameterString & vbwProtector.vbwReportParameter("TestDate", TestDate) & ") "
5751           End If
5752           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5753       End If
' </VB WATCH>
5754       If rsBalanceHoles.State = adStateOpen Then
5755           rsBalanceHoles.Close
5756       End If
5757       qyBalanceHoles.CommandText = "SELECT BalanceHoles.*, " & _
                             "IIf([Diameter]=99, 'Slot', [diameter]) as Diameter1, IIf([BoltCircle]=99, 'Unknown', [BoltCircle]) as BoltCircle1 " & _
                             "FROM BalanceHoles " & _
                             "WHERE [SerialNo] = '" & SerialNumber & "' AND [Date] <= #" & TestDate & "# " & _
                             "ORDER BY [Date], Val([BoltCircle]);"

5758       rsBalanceHoles.Open qyBalanceHoles
5759       rsBalanceHoles.Filter = ""

5760       Set dgBalanceHoles.DataSource = rsBalanceHoles

5761       Dim c As Column
5762       For Each c In dgBalanceHoles.Columns
5763           Select Case c.DataField
               Case "BalanceHoleID"
5764               c.Visible = False
5765           Case "SerialNo"
5766               c.Visible = False
5767           Case "Date"
5768               c.Visible = True
5769               c.Alignment = dbgCenter
5770               c.Width = 2000
5771           Case "Number"
5772               c.Visible = True
5773               c.Alignment = dbgCenter
5774               c.Width = 700
5775           Case "Diameter"
5776               c.Visible = False
5777           Case "Diameter1"
5778               c.Caption = "Diameter"
5779               c.Visible = True
5780               c.Alignment = dbgCenter
5781               c.Width = 700
5782           Case "BoltCircle1"
5783               c.Caption = "Bolt Circle"
5784               c.Visible = True
5785               c.Alignment = dbgCenter
5786               c.Width = 800
5787           Case "BoltCircle"
5788               c.Visible = False
5789           Case Else ' hide all other columns.
5790               c.Visible = False
5791           End Select
5792       Next c

' <VB WATCH>
5793       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5794       Exit Sub
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
5795       On Error GoTo vbwErrHandler
5796       Const VBWPROCNAME = "frmPLCData.FixPointsToPlot"
5797       If vbwProtector.vbwTraceProc Then
5798           Dim vbwProtectorParameterString As String
5799           If vbwProtector.vbwTraceParameters Then
5800               vbwProtectorParameterString = "()"
5801           End If
5802           vbwProtector.vbwProcIn VBWPROCNAME, vbwProtectorParameterString
5803       End If
' </VB WATCH>
5804       If DataGrid2.Row = -1 Then
' <VB WATCH>
5805       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5806           Exit Sub
5807       End If
5808       Dim PresentGridRow As Integer
5809       PresentGridRow = DataGrid2.Row
5810       Dim GridIndex As Integer
5811       UpDown2.value = 8
5812       If DataGrid2.Row <> -1 Then
5813           For GridIndex = 0 To 7
5814               DataGrid2.Row = GridIndex
5815               If DataGrid2.Columns("Flow") = 0 And DataGrid2.Columns("TDH") = 0 Then
5816                   txtUpDn2.Text = GridIndex
5817                   UpDown2.value = GridIndex
' <VB WATCH>
5818       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
' </VB WATCH>
5819                   Exit Sub
5820               End If
5821           Next GridIndex
5822       End If
5823       DataGrid2.Row = PresentGridRow
' <VB WATCH>
5824       If vbwProtector.vbwTraceProc Then vbwProtector.vbwProcOut VBWPROCNAME
5825       Exit Sub
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
