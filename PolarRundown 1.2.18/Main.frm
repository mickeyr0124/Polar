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
      Tab(0).Control(0)=   "txtCustPONum"
      Tab(0).Control(1)=   "txtXPartNum"
      Tab(0).Control(2)=   "txtRVSPartNo"
      Tab(0).Control(3)=   "chkSuperMarketFeathered"
      Tab(0).Control(4)=   "grpSupermarket"
      Tab(0).Control(5)=   "CommonDialog2"
      Tab(0).Control(6)=   "frmTEMC"
      Tab(0).Control(7)=   "txtLineNumber"
      Tab(0).Control(8)=   "frmMiscPumpData"
      Tab(0).Control(9)=   "txtImpellerDia"
      Tab(0).Control(10)=   "cmdClearPumpData"
      Tab(0).Control(11)=   "frmMfr"
      Tab(0).Control(12)=   "cmdApprovePump"
      Tab(0).Control(13)=   "cmdDeletePump"
      Tab(0).Control(14)=   "txtSalesOrderNumber"
      Tab(0).Control(15)=   "cmdEnterPumpData"
      Tab(0).Control(16)=   "txtRemarks"
      Tab(0).Control(17)=   "txtDesignTDH"
      Tab(0).Control(18)=   "txtDesignFlow"
      Tab(0).Control(19)=   "txtModelNo"
      Tab(0).Control(20)=   "txtShpNo"
      Tab(0).Control(21)=   "txtBilNo"
      Tab(0).Control(22)=   "frmChempump"
      Tab(0).Control(23)=   "lbltab1(50)"
      Tab(0).Control(24)=   "lbltab1(49)"
      Tab(0).Control(25)=   "lbltab1(48)"
      Tab(0).Control(26)=   "lbltab1(47)"
      Tab(0).Control(27)=   "lbltab1(46)"
      Tab(0).Control(28)=   "lbltab1(44)"
      Tab(0).Control(29)=   "lbltab1(10)"
      Tab(0).Control(30)=   "lbltab1(0)"
      Tab(0).Control(31)=   "lbltab1(13)"
      Tab(0).Control(32)=   "lbltab1(12)"
      Tab(0).Control(33)=   "lbltab1(11)"
      Tab(0).Control(34)=   "lbltab1(3)"
      Tab(0).Control(35)=   "lbltab1(2)"
      Tab(0).Control(36)=   "lbltab1(1)"
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
    Dim boGotBalanceHoles               'do we have any balance hole data?

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
    Dim xlBook As Excel.Workbook    ' Excel Workbook Object

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
' </VB WATCH>
2          If chkBalanceHoles.value = 1 Then
3              dgBalanceHoles.Visible = True
4          Else
5              dgBalanceHoles.Visible = False
6          End If
7          If LenB(frmPLCData.txtSN.Text) = 0 Or LenB(cmbTestDate.Text) = 0 Then
8              dgBalanceHoles.Visible = False
9          End If
' <VB WATCH>
10         Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub chkCircOrifice_Click()
           'if the CircOrifice box is checked, show the size
' <VB WATCH>
11         On Error GoTo vbwErrHandler
' </VB WATCH>
12         If chkCircOrifice.value = 1 Then
13             lblCircOrifice.Visible = True
14             txtCircOrifice.Visible = True
15         Else
16             lblCircOrifice.Visible = False
17             txtCircOrifice.Visible = False
18         End If
' <VB WATCH>
19         Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub chkNPSH_Click()
           'if the NPSH file box is checked, show the file name
' <VB WATCH>
20         On Error GoTo vbwErrHandler
' </VB WATCH>
21         If chkNPSH.value = 1 Then
22             txtNPSHFile.Visible = True
23         Else
24             txtNPSHFile.Visible = False
25         End If
' <VB WATCH>
26         Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub chkOrifice_Click()
           'if the orifice box is checked, show the size
' <VB WATCH>
27         On Error GoTo vbwErrHandler
' </VB WATCH>
28         If chkOrifice.value = 1 Then
29             lblOrifice.Visible = True
30             txtOrifice.Visible = True
31         Else
32             lblOrifice.Visible = False
33             txtOrifice.Visible = False
34         End If
' <VB WATCH>
35         Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub chkPictures_Click()
           'if the pictures box is checked, show the file name
' <VB WATCH>
36         On Error GoTo vbwErrHandler
' </VB WATCH>
37         If chkPictures.value = 1 Then
38             txtPicturesFile.Visible = True
39         Else
40             txtPicturesFile.Visible = False
41         End If
' <VB WATCH>
42         Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub chkTrimmed_Click()
           'if the trimmed box is checked, show the impeller size
' <VB WATCH>
43         On Error GoTo vbwErrHandler
' </VB WATCH>
44         If chkTrimmed.value = 1 Then
45             lblImpTrim.Visible = True
46             txtImpTrim.Visible = True
47         Else
48             lblImpTrim.Visible = False
49             txtImpTrim.Visible = False
50         End If
' <VB WATCH>
51         Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub chkVibration_Click()
           'if the vibration box is checked, show the file name
' <VB WATCH>
52         On Error GoTo vbwErrHandler
' </VB WATCH>
53         If chkVibration.value = 1 Then
54             txtVibrationFile.Visible = True
55         Else
56             txtVibrationFile.Visible = False
57         End If
' <VB WATCH>
58         Exit Sub
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
    End Select
' </VB WATCH>
End Sub



Private Sub cmbFrequency_Click()
' <VB WATCH>
59         On Error GoTo vbwErrHandler
' </VB WATCH>
60         If cmbFrequency.Text = "VFD" Then
61             txtVFDFreq.Visible = True
62             lbltab2(86).Visible = True
63         Else
64             txtVFDFreq.Visible = False
65             lbltab2(86).Visible = False
66         End If
' <VB WATCH>
67         Exit Sub
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
    End Select
' </VB WATCH>
End Sub


Private Sub cmbLoopNumber_Click()
' <VB WATCH>
68         On Error GoTo vbwErrHandler
' </VB WATCH>

69         Dim I As Integer
70         I = cmbLoopNumber.ListIndex

71         Dim qyTransducers As New ADODB.Command
72         Dim rsTransducers As New ADODB.Recordset
73         qyTransducers.ActiveConnection = cnPumpData
74         qyTransducers.CommandText = "SELECT * " & _
                     "From Transducers " & _
                     "Where LoopNumber  = " & I

75         With rsTransducers     'open the recordset for the query
       '        .Index = "FindData"
76             .CursorLocation = adUseClient
77             .CursorType = adOpenStatic
78             .Open qyTransducers
79         End With
80         If rsTransducers.RecordCount = 1 Then
81             Me.cmbFlowMeter.ListIndex = rsTransducers.Fields("FlowMeter")
82             Me.cmbSuctionPressureTransducer.ListIndex = rsTransducers.Fields("SuctionPressure")
83             Me.cmbDischargePressureTransducer.ListIndex = rsTransducers.Fields("DischargePressure")
84             Me.cmbTemperatureTransducer.ListIndex = rsTransducers.Fields("Temperature")
85             Me.cmbCirculationFlowMeter.ListIndex = rsTransducers.Fields("CircFlowMeter")
86             Me.cmbPLCNo.ListIndex = rsTransducers.Fields("PLC")
87             Me.cmbAnalyzerNo.ListIndex = rsTransducers.Fields("Analyzer")
88         End If

       '    If I < 2 Then
       '        Me.cmbPLCNo.ListIndex = 0
       '    Else
       '        Me.cmbPLCNo.ListIndex = 1
       '    End If
' <VB WATCH>
89         Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub GetSuperMarketPump(SuperMarketPartNum As String, JobNumber As String)
' <VB WATCH>
90         On Error GoTo vbwErrHandler
' </VB WATCH>

           'get the data from the SupermarketPumpData table
91         qySupermarketModel.ActiveConnection = cnPumpData
92         qySupermarketModel.CommandText = "SELECT * " & _
                     "From SupermarketPumpData " & _
                     "Where Model  = '" & SuperMarketPartNum & "'"

                     'cmbSupermarketModel.ItemData(cmbSupermarketModel.ListIndex)"

93         If rsSupermarketModel.State = adStateOpen Then
94             rsSupermarketModel.Close
95         End If

96         With rsSupermarketModel     'open the recordset for the query
       '        .Index = "FindData"
97             .CursorLocation = adUseClient
98             .CursorType = adOpenStatic
99             .Open qySupermarketModel
100        End With
101        If rsSupermarketModel.RecordCount = 1 Then
102            txtSalesOrderNumber.Text = rsSupermarketModel.Fields("SalesOrder")
103            txtLineNumber.Text = rsSupermarketModel.Fields("LineNumber")
104            txtShpNo.Text = rsSupermarketModel.Fields("ShipTo")
105            txtBilNo.Text = rsSupermarketModel.Fields("BillTo")
106            txtDesignFlow.Text = rsSupermarketModel.Fields("DesignFlow")
107            txtDesignTDH.Text = rsSupermarketModel.Fields("DesignTDH")
108            txtNoPhases.Text = rsSupermarketModel.Fields("Phases")
109            txtNPSHr.Text = rsSupermarketModel.Fields("NPSHr")
110            txtRatedInputPower.Text = rsSupermarketModel.Fields("RatedInputPower")
111            txtAmps.Text = rsSupermarketModel.Fields("RatedCurrent")
112            txtThermalClass.Text = rsSupermarketModel.Fields("ThermalClass")
113            txtSpGr.Text = rsSupermarketModel.Fields("SG")
114            txtViscosity.Text = rsSupermarketModel.Fields("Viscosity")
115            txtExpClass.Text = rsSupermarketModel.Fields("EXPClass")
116            txtLiquid.Text = rsSupermarketModel.Fields("Liquid")
117            txtLiquidTemperature.Text = rsSupermarketModel.Fields("LiquidTemp")
118            txtJobNum.Text = JobNumber
119            txtImpellerDia.Text = rsSupermarketModel.Fields("ImpellerDiameter")
120            txtModelNo.Text = rsSupermarketModel.Fields("Model")
121            txtRVSPartNo.Text = rsSupermarketModel.Fields("RVSPartNo")
122            cmdSelectSupermarket.Caption = "Save Data"
123            If UCase(rsSupermarketModel.Fields("Feathered")) = "FEATHERED" Then
124                Me.chkSuperMarketFeathered.value = Checked
125            End If
126        End If
127        grpSupermarket.Visible = False

' <VB WATCH>
128        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub cmbVoltage_click()
' <VB WATCH>
129        On Error GoTo vbwErrHandler
' </VB WATCH>
130        If Me.cmbVoltage.ListIndex = 0 Then
131            Me.cmbFrequency.ListIndex = 2
132        End If
' <VB WATCH>
133        Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub cmbMagtrol_Click()
' <VB WATCH>
134        On Error GoTo vbwErrHandler
' </VB WATCH>
135        Dim I As Integer
136        Dim sSendStr As String
137        Dim sGPIBName As String
138        Dim MagtrolName As String

139        I = cmbMagtrol.ItemData(cmbMagtrol.ListIndex)
140        sGPIBName = "GPIB" & I
141        MagtrolName = cmbMagtrol.List(cmbMagtrol.ListIndex)

142        If I = 99 Then      'manual entry
143            boMagtrolOperating = False
144            EnableMagtrolFields
145            Exit Sub
146        Else
147            boMagtrolOperating = True
148        End If

149        SetupMagtrols MagtrolName, I

' <VB WATCH>
150        Exit Sub
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
    End Select
' </VB WATCH>
End Sub


Private Sub cmbPLCLoop_Click()
           'Change the PLC that we're looking at
' <VB WATCH>
151        On Error GoTo vbwErrHandler
' </VB WATCH>

152        Dim RetVal As String

           'manual data entry selection
153        If cmbPLCLoop.ListIndex = cmbPLCLoop.ListCount - 1 Then 'no plc
154            boPLCOperating = False
155            EnablePLCFields
156            If DeviceOpen = True Then
157                RetVal = DisconnectPLC()
158            End If
159            Exit Sub
160        End If

161        If DeviceOpen = True Then
162            RetVal = DisconnectPLC()
163        End If

164        RetVal = ConnectToPLC(cmbPLCLoop.ItemData(cmbPLCLoop.ListIndex))
165        If RetVal <> 0 Then
166            MsgBox ("Can't connect to PLC - " & Description(cmbPLCLoop.ListIndex))
167            boPLCOperating = False
168            EnablePLCFields
169        Else
170            boPLCOperating = True
171            tDevice = cmbPLCLoop.ItemData(cmbPLCLoop.ListIndex)
172            DisablePLCFields
173        End If
' <VB WATCH>
174        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub cmbTestDate_Click()
           'select a test date to show
' <VB WATCH>
175        On Error GoTo vbwErrHandler
' </VB WATCH>

176        Dim sName As String
177        Dim sParam As String
178        Dim I As Integer
179        Dim j As Integer
180        Dim k As Integer
181        Dim bSk As Boolean
182        Dim sBC As Single
183        Dim NOK() As Long

184        cmdModifyBalanceHoleData.Visible = False


185        If Not boFoundTestSetup Then    'if we don't have any TestSetup data written
186            boFoundTestData = False
187            Exit Sub
188        End If


           'select the testsetup data for the serial number
189        qyTestSetup.ActiveConnection = cnPumpData
190        qyTestSetup.CommandText = "SELECT * " & _
                         "From TempTestSetupData " & _
                         "Where (((TempTestSetupData.SerialNumber) = '" & txtSN.Text & "') AND TempTestSetupData.Date = #" & cmbTestDate.List(cmbTestDate.ListIndex) & "#) " & _
                         "ORDER BY TempTestSetupData.Date;"

191        If rsTestSetup.State = adStateOpen Then
192            rsTestSetup.Close
193        End If

194        With rsTestSetup     'open the recordset for the query
       '        .Index = "FindData"
195            .CursorLocation = adUseClient
196            .CursorType = adOpenStatic
197            .Open qyTestSetup
198        End With

           'move to the selected date
199        If Not rsTestSetup.BOF Then
200            rsTestSetup.MoveFirst
201        End If
       '
           'show the correct combo box entries for this record
           'SetComboTestSetup cmbOrificeNumber, "OrificeNumber", "OrificeNumber", rsTestSetup
202        SetComboTestSetup cmbTestSpec, "TestSpec", "TestSpecification", rsTestSetup
203        SetComboTestSetup cmbLoopNumber, "LoopNumber", "LoopNumber", rsTestSetup
204        SetComboTestSetup cmbSuctDia, "SuctDiam", "SuctionDiameter", rsTestSetup
205        SetComboTestSetup cmbDischDia, "DischDiam", "DischargeDiameter", rsTestSetup
206        SetComboTestSetup cmbTachID, "TachID", "TachID", rsTestSetup
207        SetComboTestSetup cmbAnalyzerNo, "AnalyzerNo", "AnalyzerNo", rsTestSetup
208        SetComboTestSetup cmbVoltage, "Voltage", "Voltage", rsTestSetup
209        SetComboTestSetup cmbFrequency, "Frequency", "Frequency", rsTestSetup
210        SetComboTestSetup cmbMounting, "Mounting", "Mounting", rsTestSetup
211        SetComboTestSetup cmbPLCNo, "PLCNo", "PLCNo", rsTestSetup
212        SetComboTestSetup cmbFlowMeter, "FlowMeterID", "PumpFlowMeter", rsTestSetup
213        SetComboTestSetup cmbSuctionPressureTransducer, "SuctionID", "SuctionPressureTransducer", rsTestSetup
214        SetComboTestSetup cmbDischargePressureTransducer, "DischID", "DischargePressureTransducer", rsTestSetup
215        SetComboTestSetup cmbTemperatureTransducer, "TemperatureID", "TemperatureTransducer", rsTestSetup
216        SetComboTestSetup cmbCirculationFlowMeter, "MagFlowID", "CirculationFlowMeter", rsTestSetup

217        sName = "HDCor"
218        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
219            sParam = rsTestSetup.Fields(sName)
220        Else
221            sParam = vbNullString
222        End If
223        txtHDCor.Text = sParam

224        sName = "KWMult"
225        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
226            sParam = rsTestSetup.Fields(sName)
227        Else
228            sParam = vbNullString
229        End If
230        txtKWMult.Text = sParam

231        sName = "Who"
232        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
233            sParam = rsTestSetup.Fields(sName)
234        Else
235            sParam = vbNullString
236        End If
237        txtWho.Text = sParam

238        sName = "RMA"
239        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
240            sParam = rsTestSetup.Fields(sName)
241        Else
242            sParam = vbNullString
243        End If
244        txtRMA.Text = sParam

245        sName = "Remarks"
246        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
247            sParam = rsTestSetup.Fields(sName)
248        Else
249            sParam = vbNullString
250        End If
251        txtTestSetupRemarks.Text = sParam

252        sName = "VFDFrequency"
253        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
254            sParam = rsTestSetup.Fields(sName)
255        Else
256            sParam = vbNullString
257        End If
258        txtVFDFreq.Text = sParam

259        sName = "SuctionGageHeight"
260        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
261            sParam = rsTestSetup.Fields(sName)
262        Else
263            sParam = 0
264        End If
265        txtSuctHeight.Text = sParam

266        sName = "DischargeGageHeight"
267        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
268            sParam = rsTestSetup.Fields(sName)
269        Else
270            sParam = 0
271        End If
272        txtDischHeight.Text = sParam

273        sName = "EndPlay"
274        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
275            sParam = rsTestSetup.Fields(sName)
276        Else
277            sParam = vbNullString
278        End If
279        txtEndPlay.Text = sParam

280        sName = "GGAP"
281        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
282            sParam = rsTestSetup.Fields(sName)
283        Else
284            sParam = vbNullString
285        End If
286        txtGGap.Text = sParam

287        sName = "OtherMods"
288        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
289            sParam = rsTestSetup.Fields(sName)
290        Else
291            sParam = vbNullString
292        End If
293        txtOtherMods.Text = sParam

294        If rsTestSetup.Fields("ImpFeathered") Then
295            chkFeathered.value = 1
296        Else
297            chkFeathered.value = 0
298        End If

299        If Val(rsTestSetup.Fields("ImpTrimmed")) = 0 Then
300            chkTrimmed.value = 0
301            txtImpTrim.Visible = False
302            txtImpTrim.Text = rsTestSetup.Fields("Imptrimmed")
303        Else
304            chkTrimmed.value = 1
305            txtImpTrim.Visible = True
306            txtImpTrim.Text = rsTestSetup.Fields("Imptrimmed")
307        End If

308        If Val(rsTestSetup.Fields("PumpDischOrifice")) = 0 Then
309            chkOrifice.value = 0
310            txtOrifice.Visible = False
311        Else
312            chkOrifice.value = 1
313            txtOrifice.Visible = True
314            txtOrifice.Text = rsTestSetup.Fields("PumpDischOrifice")
315        End If

316        If Val(rsTestSetup.Fields("CircFlowOrifice")) = 0 Then
317            chkCircOrifice.value = 0
318            txtCircOrifice.Visible = False
319        Else
320            chkCircOrifice.value = 1
321            txtCircOrifice.Visible = True
322            txtCircOrifice.Text = rsTestSetup.Fields("CircFlowOrifice")
323        End If

324        If (IsNull(rsTestSetup.Fields("NPSHFile"))) Or (LenB(rsTestSetup.Fields("NPSHFile")) = 0) Then
325            chkNPSH.value = 0
326            txtNPSHFile.Visible = False
327        Else
328            chkNPSH.value = 1
329            txtNPSHFile.Visible = True
330            txtNPSHFile.Text = rsTestSetup.Fields("NPSHFile")
331        End If

332        If (IsNull(rsTestSetup.Fields("PictureFile"))) Or (LenB(rsTestSetup.Fields("PictureFile")) = 0) Then
333            chkPictures.value = 0
334            txtPicturesFile.Visible = False
335        Else
336            chkPictures.value = 1
337            txtPicturesFile.Visible = True
338            txtPicturesFile.Text = rsTestSetup.Fields("PictureFile")
339        End If

340        If (IsNull(rsTestSetup.Fields("VibrationFile"))) Or (LenB(rsTestSetup.Fields("VibrationFile")) = 0) Then
341            chkVibration.value = 0
342            txtVibrationFile.Visible = False
343        Else
344            chkVibration.value = 1
345            txtVibrationFile.Visible = True
346            txtVibrationFile.Text = rsTestSetup.Fields("VibrationFile")
347        End If


           'for TEMC Inspection Report
348        sName = "InsulationMeggerVolts"
349        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
350            sParam = rsTestSetup.Fields(sName)
351        Else
352            sParam = 0
353        End If
354        txtTestAndInspection(0).Text = sParam

355        sName = "InsulationMegOhms"
356        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
357            sParam = rsTestSetup.Fields(sName)
358        Else
359            sParam = 0
360        End If
361        txtTestAndInspection(1).Text = sParam

362        sName = "DielectricVolts"
363        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
364            sParam = rsTestSetup.Fields(sName)
365        Else
366            sParam = 0
367        End If
368        txtTestAndInspection(2).Text = sParam

369        sName = "DielectricTime"
370        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
371            sParam = rsTestSetup.Fields(sName)
372        Else
373            sParam = 0
374        End If
375        txtTestAndInspection(3).Text = sParam

376        sName = "HydrostaticValue"
377        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
378            sParam = rsTestSetup.Fields(sName)
379        Else
380            sParam = 0
381        End If
382        txtTestAndInspection(4).Text = sParam

383        sName = "HydrostaticTime"
384        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
385            sParam = rsTestSetup.Fields(sName)
386        Else
387            sParam = 0
388        End If
389        txtTestAndInspection(5).Text = sParam

390        sName = "PneumaticValue"
391        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
392            sParam = rsTestSetup.Fields(sName)
393        Else
394            sParam = 0
395        End If
396        txtTestAndInspection(6).Text = sParam

397        sName = "PneumaticTime"
398        If rsTestSetup.Fields(sName).ActualSize <> 0 Then
399            sParam = rsTestSetup.Fields(sName)
400        Else
401            sParam = 0
402        End If
403        txtTestAndInspection(7).Text = sParam

404        For I = 0 To cmbTestAndInspection(0).ListCount - 1
405            If cmbTestAndInspection(0).Text = rsTestSetup.Fields("HydrostaticUnits") Then
406                    cmbTestAndInspection(0).ListIndex = I
407                    Exit For
408            End If
409            cmbTestAndInspection(0).ListIndex = -1
410        Next I


411        For I = 0 To cmbTestAndInspection(1).ListCount - 1
412            If cmbTestAndInspection(1).Text = rsTestSetup.Fields("PneumaticUnits") Then
413                    cmbTestAndInspection(1).ListIndex = I
414                    Exit For
415            End If
416            cmbTestAndInspection(1).ListIndex = -1
417        Next I

418        TestAndInspectionGood(0).value = Abs(rsTestSetup!insulationgood)
419        TestAndInspectionGood(1).value = Abs(rsTestSetup!DielectricGood)
420        TestAndInspectionGood(2).value = Abs(rsTestSetup!HydrostaticGood)
421        TestAndInspectionGood(3).value = Abs(rsTestSetup!PneumaticGood)
422        TestAndInspectionGood(4).value = Abs(rsTestSetup!GeneralAppearanceGood)
423        TestAndInspectionGood(5).value = Abs(rsTestSetup!OutlineDimensionsGood)
424        TestAndInspectionGood(6).value = Abs(rsTestSetup!MotorNoLoadTestGood)
425        TestAndInspectionGood(7).value = Abs(rsTestSetup!MotorLockedRotorTestGood)
426        TestAndInspectionGood(8).value = Abs(rsTestSetup!HydrostaticTestGood)
427        TestAndInspectionGood(9).value = Abs(rsTestSetup!HydraulicTestGood)
428        TestAndInspectionGood(10).value = Abs(rsTestSetup!NPSHTestGood)
429        TestAndInspectionGood(11).value = Abs(rsTestSetup!CleanPurgeSealGood)
430        TestAndInspectionGood(12).value = Abs(rsTestSetup!PaintCheckGood)
431        TestAndInspectionGood(13).value = Abs(rsTestSetup!NameplateGood)
432        TestAndInspectionGood(14).value = Abs(rsTestSetup!SupervisorApproval)

433        GetBalanceHoleData frmPLCData.txtSN.Text, cmbTestDate.Text

434         If rsBalanceHoles.RecordCount = 0 Then
435            chkBalanceHoles.value = 0
436            dgBalanceHoles.Visible = False
437            boGotBalanceHoles = False
438        Else
439            boGotBalanceHoles = True
440            ReDim NOK(rsBalanceHoles.RecordCount)
441            rsBalanceHoles.MoveLast
442            For I = 1 To rsBalanceHoles.RecordCount
443                NOK(I) = 0
444            Next I

445            For j = 1 To rsBalanceHoles.RecordCount - 1
446                rsBalanceHoles.MoveFirst
447                rsBalanceHoles.Move rsBalanceHoles.RecordCount - j
448                sBC = rsBalanceHoles.Fields("BoltCircle")
449                bSk = False
450                For k = 1 To rsBalanceHoles.RecordCount
451                    If NOK(k) = rsBalanceHoles.Fields(0) Then
452                        bSk = True
453                    End If
454                Next k
455                If Not bSk Then
456                    For I = rsBalanceHoles.RecordCount - j To 1 Step -1
457                        rsBalanceHoles.MovePrevious
458                        If rsBalanceHoles.Fields("BoltCircle") = sBC Then
459                            NOK(I) = rsBalanceHoles.Fields(0)
460                        End If
461                    Next I
462                End If
463            Next j

464            Dim sFilt
465            sFilt = ""
466            For I = 1 To rsBalanceHoles.RecordCount
467                If NOK(I) <> 0 Then
468                    sFilt = sFilt & "(BalanceHoleID <> " & NOK(I) & ") AND "
       '                sFilt = sFilt & "(" & rsBalanceHoles.Filter & " AND BalanceHoleID <> " & NOK(I) & ") AND "
469                End If
470            Next I

471            If Len(sFilt) > 4 Then
472                sFilt = Left(sFilt, Len(sFilt) - 4)
473                rsBalanceHoles.Filter = sFilt
474            End If

475            chkBalanceHoles.value = 1
476            dgBalanceHoles.Visible = True
477        End If
       '
           'set the test date filter for the test data
478        rsTestData.Filter = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"

479        If rsTestData.RecordCount = 0 Then
480            boFoundTestData = False
481            AddTestData
482            EnableTestDataControls
483            MsgBox "No Test Data Exists for this Serial Number"
484        Else
485            boFoundTestData = True
486            DisableTestDataControls                         'if it's in the real database, don't allow changes here
487        End If

488        If Not boTestDateIsApproved Then    'data approved?
489            EnableTestDataControls
490        End If

491        If rsTestSetup.Fields("Approved") = True Then
492            DisableTestDataControls                         'if it's in the real database, don't allow changes here
493            lblTestDateApproved.Visible = True
494            MsgBox ("Found pump.  Data cannot be modified.")
495            If boCanApprove Then
496                cmdApproveTestDate.Caption = "Unapprove this Test Date"
497            End If
498        Else
499            EnableTestDataControls                          'it's in the temp database, allow changes
500            lblTestDateApproved.Visible = False
501            If boPumpIsApproved = True Then
502                MsgBox ("Found pump.  Pump data cannot be modified, but test setup data and test data can be modified.")
503            Else
504                MsgBox ("Found pump.  Pump data, test setup data, and test data can be modified.")
505            End If
506            If boCanApprove Then
507                If rsPumpData.Fields("Approved") = True Then
508                    cmdApproveTestDate.Enabled = True
509                    cmdApproveTestDate.Caption = "Approve this Test Date"
510                Else
511                    cmdApproveTestDate.Caption = "You Must Approve Pump First"
512                    cmdApproveTestDate.Enabled = False
513                End If
514            End If
515        End If

516        rsEff.MoveFirst
517        rsTestData.MoveFirst

518        For I = 1 To rsTestData.RecordCount
519            DoEfficiencyCalcs
520            rsEff.MoveNext
521            rsTestData.MoveNext
522        Next I

          ' fix the datagrid
523       Set DataGrid1.DataSource = rsTestData
524       Set DataGrid2.DataSource = rsEff

525       Dim c As Column
526       For Each c In DataGrid1.Columns
527          Select Case c.DataField
             Case "TestDataID"     'Hide some columns
528             c.Visible = False
529          Case "SerialNumber"
530             c.Visible = False
531          Case "Date"
532             c.Visible = False
533          Case Else             ' Show all other columns.
534             c.Visible = True
535             c.Alignment = dbgRight
536          End Select
537        Next c

538        For Each c In DataGrid2.Columns
539            c.Alignment = dbgCenter
540            c.Width = 750
541            Select Case c.ColIndex
                   Case 1
542                    c.Caption = "Flow"
543                    c.NumberFormat = "###0.00"
544                Case 2
545                    c.Caption = "TDH"
546                    c.NumberFormat = "##0.00"
547                Case 3
548                    c.Caption = "Input Pwr"
549                    c.NumberFormat = "##0.00"
550                    c.Width = 850
551                Case 4
552                    c.Caption = "Voltage"
553                    c.NumberFormat = "##0.00"
554                Case 5
555                    c.Caption = "Current"
556                    c.NumberFormat = "##0.00"
557                Case 6
558                    c.Caption = "Overall Eff"
559                    c.NumberFormat = "##0.00"
560                    c.Width = 850
561                Case 7
562                    c.Caption = "NPSHr"
563                    c.NumberFormat = "#0.00"
564                Case Else
565                    c.Visible = False
566            End Select
567        Next c
568            FixPointsToPlot

569        txtUpDn1.Text = 1

       'unlock the text boxes
570        For I = 0 To 7
571            txtTitle(I).Locked = False
572        Next I

573        For I = 20 To 27
574            txtTitle(I).Locked = False
575        Next I

       'look for titles for TCs and AIs
576        Dim qy As New ADODB.Command
577        Dim rs As New ADODB.Recordset

578        qy.ActiveConnection = cnPumpData

           'see if we have an entry in the table
579        qy.CommandText = "SELECT * FROM AITitles " & _
                             "WHERE (((AITitles.SerialNo)= '" & txtSN.Text & "') " & _
                             "AND ((AITitles.Date)= #" & cmbTestDate.Text & "#)); "

580        With rs     'open the recordset for the query
581            .CursorLocation = adUseClient
582            .CursorType = adOpenStatic
583            .LockType = adLockOptimistic
584            .Open qy
585        End With

586        If Not (rs.BOF = True And rs.EOF = True) Then   'update titles
587            rs.MoveFirst
588            Do While Not rs.EOF
589                txtTitle(rs.Fields("Channel")).Text = rs.Fields("Title")
590                rs.MoveNext
591            Loop
592        End If

593        rs.Close
594        Set rs = Nothing
595        Set qy = Nothing
' <VB WATCH>
596        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub cmdAddNewBalanceHoles_Click()
' <VB WATCH>
597        On Error GoTo vbwErrHandler
' </VB WATCH>
598        Dim strInput As String
599        Dim I As Integer
600        Dim sNumber As Integer
601        Dim sDia As Single
602        Dim sBC As Single

           'get the data for the balance holes
603        strInput = InputBox("Enter Number of Holes")
604        If strInput <> "" Then
605            sNumber = CInt(strInput)
606        Else
607            GoTo CancelPressed
608        End If

609        strInput = InputBox("Enter Decimal Value of Hole Diameter or Slot (For Example, 0.675) ")
610        If strInput <> "" Then
611            If UCase(strInput) = "SLOT" Then
612                strInput = 99
613            End If
614            sDia = CSng(strInput)
615        Else
616            GoTo CancelPressed
617        End If

618        strInput = InputBox("Enter Decimal Value of Bolt Circle or Unknown (For Example, 4.525)")
619        If strInput <> "" Then
620            If UCase(strInput) = "UNKNOWN" Then
621                strInput = 99
622            End If
623            sBC = CSng(strInput)
624        Else
625            GoTo CancelPressed
626        End If

627        GetBalanceHoleData frmPLCData.txtSN.Text, cmbTestDate.Text

628        rsBalanceHoles.AddNew
629        rsBalanceHoles!SerialNo = txtSN.Text
630        rsBalanceHoles!Date = cmbTestDate.Text
631        rsBalanceHoles!Number = sNumber
632        rsBalanceHoles!diameter = sDia
633        rsBalanceHoles!boltcircle = sBC

634        rsBalanceHoles.Update

635        GetBalanceHoleData txtSN.Text, cmbTestDate.Text
636        rsBalanceHoles.MoveLast
637        dgBalanceHoles.Refresh
638        chkBalanceHoles.value = 1

639        Exit Sub

640    CancelPressed:
641        MsgBox "No New Balance Hole Data Entered", vbOKOnly
' <VB WATCH>
642        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub cmdAddNewTestDate_Click()
           'add a new test date/time
' <VB WATCH>
643        On Error GoTo vbwErrHandler
' </VB WATCH>
644        Dim I As Integer

645        chkFeathered.value = chkSuperMarketFeathered.value

646        For I = 1 To cmbTestDate.ListCount      'see if we already have today's date entered
647            If cmbTestDate.List(I) = Date Then
648                MsgBox "There is already an entry for today.  You can only have one entry for each Serial Number and a given date.  You may want to modify the Serial Number.", vbOKOnly
649                Exit Sub
650            End If
651        Next I

           'we didn't find today's date entered, allow data entry
652        boFoundTestSetup = False

653        EnableTestSetupDataControls
654        Pressed = False
655        cmdEnterTestSetupData_Click
656        cmdAddNewBalanceHoles.Visible = True
657        txtWho.Text = LogInInitials
658        MsgBox "New Test Date Added - " & cmbTestDate.List(cmbTestDate.ListCount - 1), vbOKOnly, "Added New Test Date"
' <VB WATCH>
659        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub cmdApprovePump_Click()
           'allow the pump data to be approved
' <VB WATCH>
660        On Error GoTo vbwErrHandler
' </VB WATCH>
661        rsPumpData.Fields("Approved") = Not rsPumpData.Fields("Approved")
662        rsPumpData.Update
663        rsPumpData.Requery
664        lblPumpApproved.Visible = rsPumpData.Fields("Approved")
665        If rsPumpData.Fields("Approved") = True Then
666            cmdApprovePump.Caption = "Unapprove This Pump"
667            cmdApproveTestDate.Enabled = True
668            If rsTestSetup.Fields("Approved") = True Then
669                cmdApproveTestDate.Caption = "Unapprove This Test Date"
670            Else
671                cmdApproveTestDate.Caption = "Approve This Test Date"
672            End If
673        Else
674            cmdApprovePump.Caption = "Approve This Pump"
675            cmdApproveTestDate.Caption = "You Must Approve Pump First"
676            cmdApproveTestDate.Enabled = False
677        End If
' <VB WATCH>
678        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub cmdApproveTestDate_Click()
           'allow the test setup data to be approved
' <VB WATCH>
679        On Error GoTo vbwErrHandler
' </VB WATCH>
680        rsTestSetup.Fields("Approved") = Not rsTestSetup.Fields("Approved")
681        rsTestSetup.Update
682        rsTestSetup.Requery
683        lblTestDateApproved.Visible = rsTestSetup.Fields("Approved")
684        If rsTestSetup.Fields("Approved") = True Then
685            cmdApproveTestDate.Caption = "Unapprove This Test Date"
686        Else
687            cmdApproveTestDate.Caption = "Approve This Test Date"
688        End If
' <VB WATCH>
689        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub cmdCalibrate_Click()
' <VB WATCH>
690        On Error GoTo vbwErrHandler
' </VB WATCH>
691        Dim ans As Integer
692        Dim I As Integer

693        ans = MsgBox("You have selected to calibrate the software.  Do you want to continue?", vbYesNo, "Calibrate Software")
694        If ans = vbNo Then
695            Calibrating = False
696            Exit Sub
697        Else
698            CalibrateSoftware
699        End If
' <VB WATCH>
700        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub cmdClearPumpData_Click()
' <VB WATCH>
701        On Error GoTo vbwErrHandler
' </VB WATCH>
702        BlankData
' <VB WATCH>
703        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub cmdDeletePump_Click()
           'delete this pump
' <VB WATCH>
704        On Error GoTo vbwErrHandler
' </VB WATCH>
705        Dim Answer As Integer
706        Answer = MsgBox("You are about to delete the following record: S/N = " & rsPumpData.Fields("SerialNumber") & "!  Do you want to continue?", vbCritical Or vbYesNo, "Ready to Delete")
707        If Answer = vbYes Then
708            rsPumpData.Delete
709            rsPumpData.Update
710            cmdFindPump_Click
711        End If
' <VB WATCH>
712        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub cmdDeleteTestDate_Click()
           'delete this test date
' <VB WATCH>
713        On Error GoTo vbwErrHandler
' </VB WATCH>
714        Dim Answer As Integer
715        Answer = MsgBox("You are about to delete the following record: S/N = " & rsTestData.Fields("SerialNumber") & " and Test Date = " & rsTestSetup.Fields("Date") & "!  Do you want to continue?", vbCritical Or vbYesNo, "Ready to Delete")
716        If Answer = vbYes Then
717            rsTestSetup.Delete
718            rsTestSetup.Update
719            cmdFindPump_Click
720        End If
' <VB WATCH>
721        Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub cmdEnterPumpData_Click()
           'store the data on the screen to the pump (pumpdata)
' <VB WATCH>
722        On Error GoTo vbwErrHandler
' </VB WATCH>
723        Dim d As Integer
724        Dim sSearch As String
725        Dim ans As Integer
726        Dim boWriteDataWritten As Boolean


           'check for a serial number
727        If LenB(txtSN.Text) = 0 Then
728            MsgBox "You must have a Serial Number to enter data.  Data has not been saved."
729            Exit Sub
730        End If

           'check to make sure most entries are filled in
731        If LenB(txtModelNo.Text) = 0 And optMfr(0).value = True Then
732            MsgBox "You need to enter a MODEL NO before saving data.  Data has not been saved.", vbOKOnly
733            Exit Sub
734        End If
735        If LenB(txtSalesOrderNumber.Text) = 0 Then
736            If InStr(1, txtSN.Text, "-") <> 0 Then
737                txtSalesOrderNumber.Text = Mid$(txtSN.Text, 1, InStr(1, txtSN.Text, "-") - 1)
738            End If
739        End If
740        If LenB(txtSalesOrderNumber.Text) = 0 Then
741            MsgBox "You need to enter a SALES ORDER NUMBER before saving data.  Data has not been saved.", vbOKOnly
742            Exit Sub
743        End If

744        If cmbMotor.ListIndex = -1 And optMfr(0).value = True Then
745            MsgBox "You need to pick a MOTOR before saving data.  Data has not been saved.", vbOKOnly
746            Exit Sub
747        End If

748        If cmbStatorFill.ListIndex = -1 And optMfr(0).value = True Then    'set default
749            cmbStatorFill.ListIndex = 0
750        End If

751        If cmbModel.ListIndex = -1 And optMfr(0).value = True Then
752            MsgBox "You need to pick a MODEL before saving data.  Data has not been saved.", vbOKOnly
753            Exit Sub
754        End If

755        If cmbModelGroup.ListIndex = -1 And optMfr(0).value = True Then
756            MsgBox "You need to pick a MODEL GROUP before saving data.  Data has not been saved.", vbOKOnly
757            Exit Sub
758        End If


759        If cmbDesignPressure.ListIndex = -1 And optMfr(0).value = True Then
760            MsgBox "You need to pick a DESIGN PRESSURE before saving data.  Data has not been saved.", vbOKOnly
761            Exit Sub
762        End If

763        If cmbCirculationPath.ListIndex = -1 And optMfr(0).value = True Then
764            MsgBox "You need to pick a CIRCULATION PATH before saving data.  Data has not been saved.", vbOKOnly
765            Exit Sub
766        End If

767        If cmbRPM.ListIndex = -1 And optMfr(0).value = True Then
768            MsgBox "You need to pick an RPM before saving data.  Data has not been saved.", vbOKOnly
769            Exit Sub
770        End If

       'check TEMC dropdowns

771        If cmbTEMCAdapter.ListIndex = -1 And optMfr(0).value = False Then
772            MsgBox "You need to pick a TEMC ADAPTER before saving data.  Data has not been saved.", vbOKOnly
773            Exit Sub
774        End If

775        If cmbTEMCAdditions.ListIndex = -1 And optMfr(0).value = False Then
776            MsgBox "You need to pick TEMC ADDITIONS before saving data.  Data has not been saved.", vbOKOnly
777            Exit Sub
778        End If

779        If cmbTEMCCirculation.ListIndex = -1 And optMfr(0).value = False Then
780            MsgBox "You need to pick a TEMC CIRCULATION before saving data.  Data has not been saved.", vbOKOnly
781            Exit Sub
782        End If

783        If cmbTEMCDesignPressure.ListIndex = -1 And optMfr(0).value = False Then
784            MsgBox "You need to pick a TEMC DESIGN PRESSURE before saving data.  Data has not been saved.", vbOKOnly
785            Exit Sub
786        End If

787        If cmbTEMCDivisionType.ListIndex = -1 And optMfr(0).value = False Then
788            MsgBox "You need to pick a TEMC DIVISION TYPE before saving data.  Data has not been saved.", vbOKOnly
789            Exit Sub
790        End If

791        If cmbTEMCImpellerType.ListIndex = -1 And optMfr(0).value = False Then
792            MsgBox "You need to pick a TEMC IMPELLER TYPE before saving data.  Data has not been saved.", vbOKOnly
793            Exit Sub
794        End If

795        If cmbTEMCInsulation.ListIndex = -1 And optMfr(0).value = False Then
796            MsgBox "You need to pick a TEMC INSULATION TYPE before saving data.  Data has not been saved.", vbOKOnly
797            Exit Sub
798        End If

799        If cmbTEMCJacketGasket.ListIndex = -1 And optMfr(0).value = False Then
800            MsgBox "You need to pick a TEMC JACKET GASKET before saving data.  Data has not been saved.", vbOKOnly
801            Exit Sub
802        End If

803        If cmbTEMCMaterials.ListIndex = -1 And optMfr(0).value = False Then
804            MsgBox "You need to pick TEMC MATERIALS before saving data.  Data has not been saved.", vbOKOnly
805            Exit Sub
806        End If

807        If cmbTEMCModel.ListIndex = -1 And optMfr(0).value = False Then
808            MsgBox "You need to pick a TEMC MODEL before saving data.  Data has not been saved.", vbOKOnly
809            Exit Sub
810        End If

811        If cmbTEMCNominalImpSize.ListIndex = -1 And optMfr(0).value = False Then
812            MsgBox "You need to pick a TEMC NOMINAL IMPELLER SIZE before saving data.  Data has not been saved.", vbOKOnly
813            Exit Sub
814        End If

815        If cmbTEMCNominalDischargeSize.ListIndex = -1 And optMfr(0).value = False Then
816            MsgBox "You need to pick a TEMC NOMINAL DISCHARGE SIZE before saving data.  Data has not been saved.", vbOKOnly
817            Exit Sub
818        End If

819        If cmbTEMCNominalSuctionSize.ListIndex = -1 And optMfr(0).value = False Then
820            MsgBox "You need to pick a TEMC NOMINAL SUCTION SIZE before saving data.  Data has not been saved.", vbOKOnly
821            Exit Sub
822        End If

823        If cmbTEMCOtherMotor.ListIndex = -1 And optMfr(0).value = False Then
824            MsgBox "You need to pick a TEMC OTHER MOTOR before saving data.  Data has not been saved.", vbOKOnly
825            Exit Sub
826        End If

827        If cmbTEMCPumpStages.ListIndex = -1 And optMfr(0).value = False Then
828            MsgBox "You need to pick TEMC NUMBER OF PUMP STAGES before saving data.  Data has not been saved.", vbOKOnly
829            Exit Sub
830        End If

831        If cmbTEMCTRG.ListIndex = -1 And optMfr(0).value = False Then
832            MsgBox "You need to pick a TEMC TRG before saving data.  Data has not been saved.", vbOKOnly
833            Exit Sub
834        End If

835        If cmbTEMCVoltage.ListIndex = -1 And optMfr(0).value = False Then
836            MsgBox "You need to pick a TEMC VOLTAGE before saving data.  Data has not been saved.", vbOKOnly
837            Exit Sub
838        End If

839        If LenB(txtTEMCFrameNumber.Text) = 0 And optMfr(0).value = False Then
840            MsgBox "You need to enter a TEMC FRAME NUMBER before saving data.  Data has not been saved.", vbOKOnly
841            Exit Sub
842        End If


843        If Not boFoundPump Then     'if we havent found a pump in the database, add it
844            rsPumpData.AddNew
845            boWriteDataWritten = False
846        Else    'else, find the entry
847            sSearch = "Serialnumber = '" & frmPLCData.txtSN.Text & "'"
848            rsPumpData.MoveFirst
849            rsPumpData.Find sSearch, , adSearchForward, 1
850            boWriteDataWritten = True
851        End If

852        If Not IsNull(rsPumpData!DataWritten) Or rsPumpData!DataWritten = True Then
853            ans = MsgBox("You have already entered data for this pump.  Do you want to overwrite the data?", vbDefaultButton2 + vbYesNo, "Overwrite Data?")
854            If ans = vbNo Then
855                rsPumpData!DataWritten = True
856                rsPumpData.Update   'update datawritten
857                Exit Sub
858            End If
859        End If

860        rsPumpData!SerialNumber = frmPLCData.txtSN.Text
861        rsPumpData!ModelNumber = frmPLCData.txtModelNo.Text
862        rsPumpData!SalesOrderNumber = frmPLCData.txtSalesOrderNumber.Text
863        rsPumpData!ShipToCustomer = frmPLCData.txtShpNo.Text
864        rsPumpData!BillToCustomer = frmPLCData.txtBilNo.Text
865        rsPumpData!ApplicationFluid = frmPLCData.txtLiquid
866        rsPumpData!NPSHFile = frmPLCData.txtNPSHFileLocation.Text
867        rsPumpData!RVSPartNo = frmPLCData.txtRVSPartNo.Text
868        rsPumpData!CustPN = frmPLCData.txtXPartNum.Text
869        rsPumpData!CustPO = frmPLCData.txtCustPONum.Text

870        If Len(frmPLCData.txtViscosity) <> 0 Then
871            rsPumpData!ApplicationViscosity = frmPLCData.txtViscosity
872        End If

873        If frmPLCData.chkSuperMarketFeathered.value = Checked Then
874            rsPumpData!Field1 = "Feathered"
875        Else
876            rsPumpData!Field1 = ""
877        End If

878        If LenB(txtSpGr.Text) <> 0 Then
879            If Not IsNumeric(frmPLCData.txtSpGr.Text) Then
880                MsgBox "Specific Gravity must be a number."
881                Exit Sub
882            End If
883            rsPumpData!SpGr = frmPLCData.txtSpGr.Text
884        End If
885        If LenB(txtImpellerDia.Text) <> 0 Then
886            If Not IsNumeric(frmPLCData.txtImpellerDia.Text) Then
887                MsgBox "Impeller Diameter must be a number."
888                Exit Sub
889            End If
890            rsPumpData!impellerdia = frmPLCData.txtImpellerDia.Text
891        End If
892        If LenB(txtDesignFlow.Text) <> 0 Then
893            rsPumpData!designflow = frmPLCData.txtDesignFlow.Text
894        End If
895        If LenB(txtDesignTDH.Text) <> 0 Then
896            rsPumpData!designtdh = frmPLCData.txtDesignTDH.Text
897        End If
898        If LenB(txtRemarks.Text) <> 0 Then
899            rsPumpData!Remarks = txtRemarks.Text
900        End If

901        If optMfr(0).value = True Then
902            d = cmbMotor.ItemData(cmbMotor.ListIndex)
903            rsPumpData!Motor = d
904            d = cmbStatorFill.ItemData(cmbStatorFill.ListIndex)
905            rsPumpData!StatorFill = d
906             d = cmbDesignPressure.ItemData(cmbDesignPressure.ListIndex)
907            rsPumpData!DesignPressure = d
908            d = cmbCirculationPath.ItemData(cmbCirculationPath.ListIndex)
909            rsPumpData!CirculationPath = d
910            d = cmbRPM.ItemData(cmbRPM.ListIndex)
911            rsPumpData!RPM = d
912            d = cmbModel.ItemData(cmbModel.ListIndex)
913            rsPumpData!Model = d
914            d = cmbModelGroup.ItemData(cmbModelGroup.ListIndex)
915            rsPumpData!ModelGroup = d
916        End If
       '   TEMC fields
917        If optMfr(0).value = False Then
918            d = cmbTEMCAdapter.ItemData(cmbTEMCAdapter.ListIndex)
919            rsPumpData!TEMCAdapter = d

920            d = cmbTEMCAdditions.ItemData(cmbTEMCAdditions.ListIndex)
921            rsPumpData!TEMCAdditions = d

922            d = cmbTEMCCirculation.ItemData(cmbTEMCCirculation.ListIndex)
923            rsPumpData!TEMCcirculation = d

924            d = cmbTEMCDesignPressure.ItemData(cmbTEMCDesignPressure.ListIndex)
925            rsPumpData!TEMCDesignpressure = d

926            d = cmbTEMCDivisionType.ItemData(cmbTEMCDivisionType.ListIndex)
927            rsPumpData!TEMCDivisionType = d

928            d = cmbTEMCImpellerType.ItemData(cmbTEMCImpellerType.ListIndex)
929            rsPumpData!TEMCImpellerType = d

930            d = cmbTEMCInsulation.ItemData(cmbTEMCInsulation.ListIndex)
931            rsPumpData!TEMCInsulation = d

932            d = cmbTEMCJacketGasket.ItemData(cmbTEMCJacketGasket.ListIndex)
933            rsPumpData!TEMCJacketGasket = d

934            d = cmbTEMCMaterials.ItemData(cmbTEMCMaterials.ListIndex)
935            rsPumpData!TEMCMaterials = d

936            d = cmbTEMCModel.ItemData(cmbTEMCModel.ListIndex)
937            rsPumpData!TEMCModel = d

938            d = cmbTEMCNominalImpSize.ItemData(cmbTEMCNominalImpSize.ListIndex)
939            rsPumpData!TEMCNominalImpSize = d

940            d = cmbTEMCNominalDischargeSize.ItemData(cmbTEMCNominalDischargeSize.ListIndex)
941            rsPumpData!TEMCNominalDischargeSize = d

942            d = cmbTEMCNominalSuctionSize.ItemData(cmbTEMCNominalSuctionSize.ListIndex)
943            rsPumpData!TEMCNominalSuctionSize = d

944            d = cmbTEMCOtherMotor.ItemData(cmbTEMCOtherMotor.ListIndex)
945            rsPumpData!TEMCOtherMotor = d

946            d = cmbTEMCPumpStages.ItemData(cmbTEMCPumpStages.ListIndex)
947            rsPumpData!TEMCPumpStages = d

948            d = cmbTEMCTRG.ItemData(cmbTEMCTRG.ListIndex)
949            rsPumpData!TEMCTRG = d

950            d = cmbTEMCVoltage.ItemData(cmbTEMCVoltage.ListIndex)
951            rsPumpData!TEMCVoltage = d

952            If LenB(txtTEMCFrameNumber.Text) <> 0 Then
953                rsPumpData!TEMCFrameNumber = frmPLCData.txtTEMCFrameNumber.Text
954            End If
955        End If

956        rsPumpData!ChempumpPump = optMfr(0).value

957        rsPumpData!Approved = False

       'added from TEMC Inspection Report
958        If Len(txtJobNum.Text) <> 0 Then
959            rsPumpData!JobNumber = txtJobNum.Text
960        End If

961        If Len(txtNoPhases.Text) <> 0 Then
962            rsPumpData!Phases = txtNoPhases.Text
963        End If

964        If Len(txtExpClass.Text) <> 0 Then
965            rsPumpData!ExpClass = txtExpClass.Text
966        End If

967        If Len(txtThermalClass.Text) <> 0 Then
968            rsPumpData!ThermalClass = txtThermalClass.Text
969        End If

970        rsPumpData!NPSHr = Val(txtNPSHr.Text)
971        rsPumpData!LiquidTemperature = Val(txtLiquidTemperature.Text)
972        rsPumpData!RatedInputPower = Val(txtRatedInputPower.Text)
973        rsPumpData!FLCurrent = Val(txtAmps.Text)





974        If boWriteDataWritten Then
975            rsPumpData!DataWritten = True
976        Else
977            rsPumpData!DataWritten = False
978        End If

           'write the data into the database
979        rsPumpData.Update
980        boFoundPump = True

           'enter a new test date if it's a new entry
981        If Not boWriteDataWritten Then


982            cmdAddNewTestDate_Click
983        End If
' <VB WATCH>
984        Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub cmdEnterTestData_Click()
           ' save the data on the screen to test data at the selected run
' <VB WATCH>
985        On Error GoTo vbwErrHandler
' </VB WATCH>
986        Dim sSearch As String
987        Dim ans As Integer

           'if we didn't find the test setup, can't enter test data
988        If Not boFoundTestSetup Then
989            MsgBox "You must enter Test Setup Data before entering the Test Data"
990            Exit Sub
991        End If

           'if we don't find data in the test database, add records
992        If boFoundTestData = False Then     'add 8 records for 8 tests
993            AddTestData
994            rsTestData.MoveFirst
995        Else        'find the data in the database
996            sSearch = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"
997            rsTestData.MoveFirst
998            rsTestData.Filter = sSearch
999        End If

           'find the desired record from the form
1000       rsTestData.MoveFirst
1001       rsTestData.Move UpDown1.value - 1

1002       If rsTestData!DataWritten = True Then
1003           ans = MsgBox("You have already entered data for this test.  Do you want to overwrite the data?", vbYesNo + vbDefaultButton2, "Data Already Entered")
1004           If ans = vbNo Then
1005               Exit Sub
1006           End If
1007       End If

1008       rsEff.MoveFirst
1009       rsEff.Move UpDown1.value - 1

1010       If LenB(txtV1.Text) <> 0 Then
1011           rsTestData!VoltageA = Val(txtV1.Text)
1012       End If

1013       If LenB(txtV2.Text) <> 0 Then
1014           rsTestData!VoltageB = Val(txtV2.Text)
1015       End If

1016       If LenB(txtV3.Text) <> 0 Then
1017           rsTestData!VoltageC = Val(txtV3.Text)
1018       End If

1019       If LenB(txtI1.Text) <> 0 Then
1020           rsTestData!CurrentA = Val(txtI1.Text)
1021       End If

1022       If LenB(txtI2.Text) <> 0 Then
1023           rsTestData!CurrentB = Val(txtI2.Text)
1024       End If

1025       If LenB(txtI3.Text) <> 0 Then
1026           rsTestData!CurrentC = Val(txtI3.Text)
1027       End If

1028       If LenB(txtP1.Text) <> 0 Then
1029           rsTestData!PowerA = Val(txtP1.Text)
1030       End If

1031       If LenB(txtP2.Text) <> 0 Then
1032           rsTestData!PowerB = Val(txtP2.Text)
1033       End If

1034       If LenB(txtP3.Text) <> 0 Then
1035           rsTestData!PowerC = Val(txtP3.Text)
1036       End If

1037       If LenB(txtKW.Text) <> 0 Then
1038           rsTestData!TotalPower = Val(txtKW.Text)
1039       End If

1040       rsTestData!Flow = Val(txtFlowDisplay.Text)
1041       rsTestData!DischargePressure = Val(txtDischargeDisplay.Text)
1042       rsTestData!SuctionPressure = Val(txtSuctionDisplay.Text)
1043       rsTestData!TemperatureSuction = Val(txtTemperatureDisplay.Text)

1044       rsTestData!TC1 = Val(txtTC1Display.Text)
1045       rsTestData!TC2 = Val(txtTC2Display.Text)
1046       rsTestData!TC3 = Val(txtTC3Display.Text)
1047       rsTestData!TC4 = Val(txtTC4Display.Text)

1048       rsTestData!CircFlow = Val(txtAI1Display.Text)
1049       rsTestData!RBHTemp = Val(txtAI2Display.Text)
1050       rsTestData!RBHPress = Val(txtAI3Display.Text)
1051       rsTestData!AI4 = Val(txtAI4Display.Text)

1052       rsTestData!ValvePosition = Val(txtValvePosition.Text)
1053       rsTestData!SetPoint = Val(txtSetPoint.Text)

1054       If LenB(txtThrustBal.Text) <> 0 Then
1055           rsTestData!ThrustBalance = txtThrustBal.Text
1056       End If

1057       If LenB(txtVibAx.Text) <> 0 Then
1058           rsTestData!VibrationX = txtVibAx.Text
1059       End If

1060       If LenB(txtVibRad.Text) <> 0 Then
1061           rsTestData!VibrationY = txtVibRad.Text
1062       End If

1063       If LenB(txtTEMCTRGReading.Text) <> 0 Then
1064           rsTestData!TEMCTRG = txtTEMCTRGReading.Text
1065       Else
1066           rsTestData!TEMCTRG = 0
1067       End If

1068       If LenB(txtRPM.Text) <> 0 Then
1069           rsTestData!RPM = txtRPM.Text
1070       End If

1071       If LenB(txtTestRemarks.Text) <> 0 Then
1072           rsTestData!Remarks = txtTestRemarks.Text
1073       Else
1074           rsTestData!Remarks = " "
1075       End If

1076       If LenB(txtTEMCTRGReading.Text) <> 0 Then
1077           rsTestData!TEMCTRG = txtTEMCTRGReading.Text
1078       End If

1079       If LenB(txtTEMCFrontThrust.Text) <> 0 Then
1080           rsTestData!TEMCFrontThrust = txtTEMCFrontThrust.Text
1081       End If

1082       If LenB(txtTEMCRearThrust.Text) <> 0 Then
1083           rsTestData!TEMCRearThrust = txtTEMCRearThrust.Text
1084       End If

1085       If LenB(txtTEMCMomentArm.Text) <> 0 Then
1086           rsTestData!TEMCMomentArm = txtTEMCMomentArm.Text
1087       End If

1088       If LenB(txtTEMCThrustRigPressure.Text) <> 0 Then
1089           rsTestData!TEMCThrustRigPressure = txtTEMCThrustRigPressure.Text
1090       End If

1091       If LenB(txtTEMCViscosity.Text) <> 0 Then
1092           rsTestData!TEMCViscosity = txtTEMCViscosity.Text
1093       End If

1094       If LenB(txtNPSHa.Text) <> 0 Then
1095           rsTestData!NPSHa = txtNPSHa.Text
1096       End If

1097       rsTestData!Approved = False

1098       rsTestData!DataWritten = True

           'update the database
1099       rsTestData.Update

1100       DoEfficiencyCalcs
1101       rsEff.Update

           'update the form
1102       DataGrid1.Refresh
1103       DataGrid2.Refresh

1104       FixPointsToPlot

' <VB WATCH>
1105       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub cmdEnterTestSetupData_Click()
           'save the data on the screen to testsetupdata
' <VB WATCH>
1106       On Error GoTo vbwErrHandler
' </VB WATCH>
1107       Dim I As Integer
1108       Dim d As Integer
1109       Dim sSearch As String
1110       Dim ans As Integer
1111       Dim boWriteDataWritten As Boolean

           'check for a serial number
1112       If LenB(txtSN.Text) = 0 Then
1113           MsgBox "You must have a Serial Number to enter data.", vbOKOnly + vbExclamation, "Cannot Enter Data"
1114           Exit Sub
1115       End If

1116       If Pressed = True Then
1117           If Me.cmbDischDia.ListIndex = -1 Or Me.cmbSuctDia.ListIndex = -1 Or Val(Me.txtSuctHeight.Text) = 0 Or Val(Me.txtDischHeight.Text) = 0 Then
1118               MsgBox "You must have Discharge Diameter AND Suction Diameter AND Suction Height AND Discharge Height entered to enter data.", vbOKOnly + vbExclamation, "Cannot Enter Data"
1119               Exit Sub
1120           End If
1121       End If

1122       Pressed = True
1123       If Not boFoundTestSetup Then    'if we didn't find any test setup, add a record
1124           rsTestSetup.AddNew
1125           cmbTestDate.AddItem Now
1126           cmbTestDate.ListIndex = cmbTestDate.NewIndex
1127           cmdAddNewBalanceHoles.Visible = True
1128           boFoundTestSetup = True
1129           boWriteDataWritten = False
1130           rsTestSetup!DataWritten = False
1131       Else    'find the record and display
1132           sSearch = "SerialNumber = '" & frmPLCData.txtSN.Text & "' AND Date = #" & cmbTestDate.Text & "#"
1133           rsTestSetup.MoveFirst
1134           rsTestSetup.Filter = sSearch
1135           If Not boCanApprove Then
       '            cmdAddNewBalanceHoles.Visible = False
1136           End If
1137           boWriteDataWritten = True
1138       End If

1139       If rsTestSetup!DataWritten = True Then
1140           ans = MsgBox("Data has already been entered for this test date.  Do you want to overwrite it?", vbYesNo + vbDefaultButton2, "Data Exists")
1141           If ans = vbNo Then
1142               Exit Sub
1143           End If
1144       End If

1145       rsTestSetup!SerialNumber = txtSN
1146       rsTestSetup!Date = cmbTestDate.List(cmbTestDate.ListIndex)

1147       I = cmbFlowMeter.ListIndex
1148       If I = -1 Then
1149           d = 1
1150           rsTestSetup!FlowMeterID = d
1151       Else
1152           d = cmbLoopNumber.ItemData(I)
1153           rsTestSetup!FlowMeterID = d
1154       End If

1155       I = cmbSuctionPressureTransducer.ListIndex
1156       If I = -1 Then
1157           d = 1
1158           rsTestSetup!suctionid = d
1159       Else
1160           d = cmbLoopNumber.ItemData(I)
1161           rsTestSetup!suctionid = d
1162       End If

1163       I = cmbDischargePressureTransducer.ListIndex
1164       If I = -1 Then
1165           d = 1
1166           rsTestSetup!dischid = d
1167       Else
1168           d = cmbLoopNumber.ItemData(I)
1169           rsTestSetup!dischid = d
1170       End If

1171       I = cmbTemperatureTransducer.ListIndex
1172       If I = -1 Then
1173           d = 1
1174           rsTestSetup!temperatureid = d
1175       Else
1176           d = cmbLoopNumber.ItemData(I)
1177           rsTestSetup!temperatureid = d
1178       End If

1179       I = Me.cmbCirculationFlowMeter.ListIndex
1180       If I = -1 Or I < 4 Then
1181           d = 5
1182           rsTestSetup!magflowid = d
1183       Else
1184           d = cmbLoopNumber.ItemData(I)
1185           rsTestSetup!magflowid = d
1186       End If


1187       If LenB(txtHDCor.Text) <> 0 Then
1188           rsTestSetup!HDCor = txtHDCor
1189       Else
1190           rsTestSetup!HDCor = 0
1191       End If
1192       If LenB(txtKWMult.Text) <> 0 Then
1193           rsTestSetup!kwmult = txtKWMult
1194       Else
1195           rsTestSetup!kwmult = 1
1196       End If
1197       If LenB(txtWho.Text) <> 0 Then
1198           rsTestSetup!who = txtWho
1199       Else
1200           rsTestSetup!who = vbNullString
1201       End If
1202       If LenB(txtRMA.Text) <> 0 Then
1203           rsTestSetup!RMA = txtRMA
1204       Else
1205           rsTestSetup!RMA = vbNullString
1206       End If
1207       If LenB(frmPLCData.txtDischHeight) <> 0 Then
1208           rsTestSetup!DischargeGageHeight = Val(txtDischHeight)
1209       Else
1210           rsTestSetup!DischargeGageHeight = 0
1211       End If
1212       If LenB(frmPLCData.txtSuctHeight) <> 0 Then
1213           rsTestSetup!SuctionGageHeight = Val(txtSuctHeight)
1214       Else
1215           rsTestSetup!SuctionGageHeight = 0
1216       End If
1217       If LenB(frmPLCData.txtTestSetupRemarks.Text) <> 0 Then
1218           rsTestSetup!Remarks = txtTestSetupRemarks.Text
1219       Else
1220           rsTestSetup!Remarks = vbNullString
1221       End If
1222       If LenB(frmPLCData.txtVFDFreq.Text) <> 0 Then
1223           rsTestSetup!VFDFrequency = txtVFDFreq.Text
1224       Else
1225           rsTestSetup!VFDFrequency = 0
1226       End If

1227       I = cmbOrificeNumber.ListIndex
1228       If I = -1 Then
1229           d = 18      'entry for None
1230       Else
1231           d = cmbOrificeNumber.ItemData(I)
1232       End If
1233       rsTestSetup!orificenumber = d

1234       If LenB(txtEndPlay.Text) <> 0 Then
1235           rsTestSetup!Endplay = Val(frmPLCData.txtEndPlay.Text)
1236       Else
1237           rsTestSetup!Endplay = 0
1238       End If

1239       If LenB(txtGGap.Text) <> 0 Then
1240           rsTestSetup!GGAP = Val(frmPLCData.txtGGap.Text)
1241       Else
1242           rsTestSetup!GGAP = 0
1243       End If

1244       If LenB(txtOtherMods.Text) <> 0 Then
1245           rsTestSetup!OtherMods = txtOtherMods.Text
1246       Else
1247           rsTestSetup!OtherMods = vbNullString
1248       End If

1249       rsTestSetup!Approved = False

1250       I = cmbLoopNumber.ListIndex
1251       If I = -1 Then
1252           d = 1
1253           rsTestSetup!loopnumber = d
1254       Else
1255           d = cmbLoopNumber.ItemData(I)
1256           rsTestSetup!loopnumber = d
1257       End If

1258       I = cmbSuctDia.ListIndex
1259       If I = -1 Then
1260           d = -1
1261       Else
1262           d = cmbSuctDia.ItemData(I)
1263           rsTestSetup!SuctDiam = d
1264       End If

1265       I = cmbDischDia.ListIndex
1266       If I = -1 Then
1267           d = -1
1268       Else
1269           d = cmbDischDia.ItemData(I)
1270           rsTestSetup!DischDiam = d
1271       End If

1272       I = cmbTachID.ListIndex
1273       If I = -1 Then
1274           d = 1
1275           rsTestSetup!tachid = d
1276       Else
1277           d = cmbTachID.ItemData(I)
1278           rsTestSetup!tachid = d
1279       End If

1280       I = cmbAnalyzerNo.ListIndex
1281       If I = -1 Then
1282           d = 1
1283       Else
1284           d = cmbAnalyzerNo.ItemData(I)
1285       End If
1286       rsTestSetup!analyzerno = d

1287       I = cmbTestSpec.ListIndex
1288       If I = -1 Then
1289           d = 1
1290       Else
1291           d = cmbTestSpec.ItemData(I)
1292       End If
1293       rsTestSetup!testspec = d

1294       I = cmbVoltage.ListIndex
1295       If I = -1 Then
1296           d = 1
1297       Else
1298           d = cmbVoltage.ItemData(I)
1299       End If
1300       rsTestSetup!Voltage = d

1301       I = cmbFrequency.ListIndex
1302       If I = -1 Then
1303           d = 1
1304       Else
1305           d = cmbFrequency.ItemData(I)
1306       End If
1307       rsTestSetup!Frequency = d

1308       I = cmbMounting.ListIndex
1309       If I = -1 Then
1310           d = 1
1311       Else
1312           d = cmbMounting.ItemData(I)
1313       End If
1314       rsTestSetup!Mounting = d

1315       I = cmbPLCNo.ListIndex
1316       If I = -1 Then
1317           d = 8
1318       Else
1319           d = cmbPLCNo.ItemData(I)
1320       End If
1321       rsTestSetup!PLCNo = d

1322       rsTestSetup!ImpFeathered = chkFeathered.value

1323       If chkTrimmed.value = 1 Then
1324           rsTestSetup!ImpTrimmed = Val(txtImpTrim)
1325       Else
1326           rsTestSetup!ImpTrimmed = 0
1327       End If
1328       chkTrimmed_Click

1329       If chkOrifice.value = 1 Then
1330           rsTestSetup!PumpDischOrifice = Val(txtOrifice)
1331       Else
1332           rsTestSetup!PumpDischOrifice = 0
1333       End If
1334       chkOrifice_Click

1335       If chkCircOrifice.value = 1 Then
1336           rsTestSetup!CircFlowOrifice = Val(txtCircOrifice)
1337       Else
1338           rsTestSetup!CircFlowOrifice = 0
1339       End If
1340       chkCircOrifice_Click

1341       chkBalanceHoles_Click

1342       If chkNPSH.value = 1 Then
1343           txtNPSHFile.Visible = True
1344           rsTestSetup!NPSHFile = txtNPSHFile
1345       Else
1346           rsTestSetup!NPSHFile = vbNullString
1347           txtNPSHFile.Visible = False
1348       End If

1349       If chkPictures.value = 1 Then
1350           txtPicturesFile.Visible = True
1351           rsTestSetup!PictureFile = txtPicturesFile
1352       Else
1353           rsTestSetup!PictureFile = vbNullString
1354           txtPicturesFile.Visible = False
1355       End If

1356       If chkVibration.value = 1 Then
1357           txtVibrationFile.Visible = True
1358           rsTestSetup!VibrationFile = txtVibrationFile
1359       Else
1360           rsTestSetup!VibrationFile = vbNullString
1361           txtVibrationFile.Visible = False
1362       End If

1363       If boWriteDataWritten Then
1364           rsTestSetup!DataWritten = True
1365       Else
1366           rsTestSetup!DataWritten = False
1367       End If

           'for TEMC Inspection Report
1368       If LenB(frmPLCData.txtTestAndInspection(0).Text) <> 0 Then
1369           rsTestSetup!InsulationMeggerVolts = frmPLCData.txtTestAndInspection(0).Text
1370       Else
1371           rsTestSetup!InsulationMeggerVolts = ""
1372       End If

1373       If LenB(frmPLCData.txtTestAndInspection(1).Text) <> 0 Then
1374           rsTestSetup!InsulationMegOhms = frmPLCData.txtTestAndInspection(1).Text
1375       Else
1376           rsTestSetup!InsulationMegOhms = ""
1377       End If

1378       If LenB(frmPLCData.txtTestAndInspection(2).Text) <> 0 Then
1379           rsTestSetup!DielectricVolts = frmPLCData.txtTestAndInspection(2).Text
1380       Else
1381           rsTestSetup!DielectricVolts = ""
1382       End If

1383       If LenB(frmPLCData.txtTestAndInspection(3).Text) <> 0 Then
1384           rsTestSetup!DielectricTime = frmPLCData.txtTestAndInspection(3).Text
1385       Else
1386           rsTestSetup!DielectricTime = ""
1387       End If

1388       If LenB(frmPLCData.txtTestAndInspection(4).Text) <> 0 Then
1389           rsTestSetup!HydrostaticValue = frmPLCData.txtTestAndInspection(4).Text
1390       Else
1391           rsTestSetup!HydrostaticValue = ""
1392       End If

1393       If LenB(frmPLCData.txtTestAndInspection(5).Text) <> 0 Then
1394           rsTestSetup!HydrostaticTime = frmPLCData.txtTestAndInspection(5).Text
1395       Else
1396           rsTestSetup!HydrostaticTime = ""
1397       End If

1398       If LenB(frmPLCData.txtTestAndInspection(6).Text) <> 0 Then
1399           rsTestSetup!PneumaticValue = frmPLCData.txtTestAndInspection(6).Text
1400       Else
1401           rsTestSetup!PneumaticValue = ""
1402       End If

1403       If LenB(frmPLCData.txtTestAndInspection(7).Text) <> 0 Then
1404           rsTestSetup!PneumaticTime = frmPLCData.txtTestAndInspection(7).Text
1405       Else
1406           rsTestSetup!PneumaticTime = ""
1407       End If

1408       I = cmbTestAndInspection(0).ListIndex
1409       If I = -1 Then
1410           rsTestSetup!HydrostaticUnits = ""
1411       Else
1412           rsTestSetup!HydrostaticUnits = cmbTestAndInspection(0).Text
1413       End If


1414       I = cmbTestAndInspection(1).ListIndex
1415       If I = -1 Then
1416           rsTestSetup!PneumaticUnits = ""
1417       Else
1418           rsTestSetup!PneumaticUnits = cmbTestAndInspection(1).Text
1419       End If

           'use abs to convert from 1 and 0 to boolean
1420       rsTestSetup!insulationgood = Abs(TestAndInspectionGood(0).value)
1421       rsTestSetup!DielectricGood = Abs(TestAndInspectionGood(1).value)
1422       rsTestSetup!HydrostaticGood = Abs(TestAndInspectionGood(2).value)
1423       rsTestSetup!PneumaticGood = Abs(TestAndInspectionGood(3).value)
1424       rsTestSetup!GeneralAppearanceGood = Abs(TestAndInspectionGood(4).value)
1425       rsTestSetup!OutlineDimensionsGood = Abs(TestAndInspectionGood(5).value)
1426       rsTestSetup!MotorNoLoadTestGood = Abs(TestAndInspectionGood(6).value)
1427       rsTestSetup!MotorLockedRotorTestGood = Abs(TestAndInspectionGood(7).value)
1428       rsTestSetup!HydrostaticTestGood = Abs(TestAndInspectionGood(8).value)
1429       rsTestSetup!HydraulicTestGood = Abs(TestAndInspectionGood(9).value)
1430       rsTestSetup!NPSHTestGood = Abs(TestAndInspectionGood(10).value)
1431       rsTestSetup!CleanPurgeSealGood = Abs(TestAndInspectionGood(11).value)
1432       rsTestSetup!PaintCheckGood = Abs(TestAndInspectionGood(12).value)
1433       rsTestSetup!NameplateGood = Abs(TestAndInspectionGood(13).value)
1434       rsTestSetup!SupervisorApproval = Abs(TestAndInspectionGood(14).value)

           'update the database
1435       rsTestSetup.Update

1436       If boFoundTestData = False Then     'add 8 records for 8 tests
1437           AddTestData
1438       End If

1439       rsTestSetup.Filter = vbNullString
' <VB WATCH>
1440       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub cmdExit_Click()
' <VB WATCH>
1441       On Error GoTo vbwErrHandler
' </VB WATCH>
1442       End
' <VB WATCH>
1443       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub cmdFindMagtrols_Click()
' <VB WATCH>
1444       On Error GoTo vbwErrHandler
' </VB WATCH>
1445       FindMagtrols
' <VB WATCH>
1446       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub cmdFindPump_Click()
           ' find the pump whose sn is shown
' <VB WATCH>
1447       On Error GoTo vbwErrHandler
' </VB WATCH>
1448       Dim sAns As String
1449       Dim sSO As String
1450       Dim sParam As String
1451       Dim sName As String

1452       Dim I As Integer

           'clear the data
1453       BlankData

           'set TC and AI labels with default values
1454       txtTitle(0).Text = "TC 1"
1455       txtTitle(1).Text = "(F)"
1456       txtTitle(2).Text = "TC 2"
1457       txtTitle(3).Text = "(F)"
1458       txtTitle(4).Text = "TC 3"
1459       txtTitle(5).Text = "(F)"
1460       txtTitle(6).Text = "TC 4"
1461       txtTitle(7).Text = "(F)"
1462       txtTitle(20).Text = "Circ Flow"
1463       txtTitle(21).Text = "(GPM)"
1464       txtTitle(22).Text = "P1"
1465       txtTitle(23).Text = "(psig)"
1466       txtTitle(24).Text = "P2"
1467       txtTitle(25).Text = "(psig)"
1468       txtTitle(26).Text = "AI 4"
1469       txtTitle(27).Text = ""


1470       For I = 0 To 7
1471           lblAutoMan(I).Caption = "Auto"
1472       Next I

1473       txtFlowDisplay.Enabled = False
1474       txtSuctionDisplay.Enabled = False
1475       txtDischargeDisplay.Enabled = False
1476       txtTemperatureDisplay.Enabled = False
1477       txtAI1Display.Enabled = False
1478       txtAI2Display.Enabled = False
1479       txtAI3Display.Enabled = False
1480       txtAI4Display.Enabled = False


1481       cmdFindPump.Default = False

           'set all found booleans to false
       '    boUsingHP = False
1482       boFoundPump = False
1483       boPumpIsApproved = False
1484       boFoundTestSetup = False
1485       boFoundTestData = False


           'get rid of all test dates in combo box
1486       For I = cmbTestDate.ListCount - 1 To 0 Step -1
1487           cmbTestDate.RemoveItem 0
1488       Next I

1489       rsTestData.Filter = "SerialNumber = ''"

1490       DataGrid2.ClearFields
1491       ClearEff

1492       If rsPumpData.State = adStateOpen Then
1493           If rsPumpData.BOF = False Or rsPumpData.EOF = False Then
1494               rsPumpData.Update
1495           End If
1496           rsPumpData.Close
1497       End If

           'parse the serial number to make sure it is formed correctly
1498       Dim ok As Boolean
1499       ok = UCase(txtSN.Text) Like "[A-Z][A-Z][0-9][0-9][0-9][0-9][A-Z]" Or UCase(txtSN.Text) Like "[A-Z][A-Z][0-9][0-9][0-9][0-9][A-Z]-[0-9]" Or UCase(txtSN.Text) Like "[A-Z][A-Z][0-9][0-9][0-9][0-9][A-Z]-[0-9][0-9]"
1500       If Not ok Then
1501           MsgBox "Serial Number must be 2 letters, 4 numbers, and 1 letter. Please re-enter.", vbOKOnly, "Serial Number not correctly formed."
1502           Exit Sub
1503       End If

           'find the pump listed in the Serial Number text box
1504       qyPumpData.ActiveConnection = cnPumpData
1505       qyPumpData.CommandText = "SELECT * From TempPumpData WHERE (((TempPumpData.SerialNumber)='" & _
                                    txtSN.Text & "'))"
1506       rsPumpData.CursorType = adOpenStatic
1507       rsPumpData.CursorLocation = adUseClient
1508       rsPumpData.Index = "SerialNumber"
1509       rsPumpData.Open qyPumpData
1510       boEpicorFound = False

1511       If rsPumpData.BOF = True And rsPumpData.EOF = True Then
               'if the bof=eof, we have an empty recordset
1512           boFoundPump = False
1513       Else
               'we found it
1514           boFoundPump = True
1515       End If

1516       If boFoundPump = False Then
               'not found in either database, try HP?
1517           sAns = MsgBox("Pump Not Found in the Database.  Look in Epicor?", vbYesNo, "Can't Find Pump")
1518           If sAns = vbNo Then     'new pump - don't get data from HP
1519               boUsingEpicor = False
1520           Else
1521               boUsingEpicor = True
       '            boUsingHP = False
1522           End If
       '        If boUsingEpicor = False Then
       '            sAns = MsgBox("Pump Not Found in the Database.  Look on the HP?", vbYesNo, "Can't Find Pump")
       '            If sAns = vbNo Then     'new pump - don't get data from HP
       '                 boUsingHP = False
       '            Else
       '                boUsingHP = True
       '            End If
       '        End If
1523           EnablePumpDataControls
1524           EnableTestSetupDataControls
1525           EnableTestDataControls
       '        BlankData               'clear any data on the screen
1526           cmdAddNewBalanceHoles.Visible = True

1527       End If

1528       If boFoundPump = True Then    'found the pump
1529           If rsPumpData.Fields("Approved") = True Then
1530               DisablePumpDataControls                         'if it's in the real database, don't allow changes here
1531               boPumpIsApproved = True
1532               lblPumpApproved.Visible = True
1533               If boCanApprove Then
1534                   cmdApprovePump.Caption = "Unapprove this pump"
1535               End If
1536               frmPLCData.cmdApproveTestDate.Enabled = True
1537           Else
1538               EnablePumpDataControls                          'it's in the temp database, allow changes
1539               boPumpIsApproved = False
1540               boTestDateIsApproved = False
1541               lblPumpApproved.Visible = False
1542               If boCanApprove Then
1543                   cmdApprovePump.Caption = "Approve this pump"
1544               End If
1545               cmdApproveTestDate.Caption = "You Must Approve Pump First"
1546               frmPLCData.cmdApproveTestDate.Enabled = False
1547           End If

               'found the pump, show the data
1548           txtModelNo.Text = rsPumpData.Fields("ModelNumber")
1549           frmPLCData.optMfr(0).value = rsPumpData.Fields("ChempumpPump")

1550           If rsPumpData.Fields("ChempumpPump") = True Then
1551               SetCombo cmbMotor, "Motor", rsPumpData
1552               SetCombo cmbDesignPressure, "DesignPressure", rsPumpData
1553               SetCombo cmbRPM, "RPM", rsPumpData
1554               SetCombo cmbCirculationPath, "CirculationPath", rsPumpData
1555               SetCombo cmbStatorFill, "StatorFill", rsPumpData
1556               SetCombo cmbModel, "Model", rsPumpData
1557               SetCombo cmbModelGroup, "ModelGroup", rsPumpData
1558               RatedKW = 999
1559           End If

               'set the TEMC data
1560           If rsPumpData.Fields("ChempumpPump") = False Then
1561               SetCombo cmbTEMCAdapter, "TEMCAdapter", rsPumpData
1562               SetCombo cmbTEMCAdditions, "TEMCAdditions", rsPumpData
1563               SetCombo cmbTEMCCirculation, "TEMCCirculation", rsPumpData
1564               SetCombo cmbTEMCDesignPressure, "TEMCDesignPressure", rsPumpData
1565               SetCombo cmbTEMCNominalDischargeSize, "TEMCNominalDischargeSize", rsPumpData
1566               SetCombo cmbTEMCDivisionType, "TEMCDivisionType", rsPumpData
1567               SetCombo cmbTEMCImpellerType, "TEMCImpellerType", rsPumpData
1568               SetCombo cmbTEMCInsulation, "TEMCInsulation", rsPumpData
1569               SetCombo cmbTEMCJacketGasket, "TEMCJacketGasket", rsPumpData
1570               SetCombo cmbTEMCMaterials, "TEMCMaterials", rsPumpData
1571               SetCombo cmbTEMCModel, "TEMCModel", rsPumpData
1572               SetCombo cmbTEMCNominalImpSize, "TEMCNominalImpSize", rsPumpData
1573               SetCombo cmbTEMCOtherMotor, "TEMCOtherMotor", rsPumpData
1574               SetCombo cmbTEMCPumpStages, "TEMCPumpStages", rsPumpData
1575               SetCombo cmbTEMCNominalSuctionSize, "TEMCNominalSuctionSize", rsPumpData
1576               SetCombo cmbTEMCTRG, "TEMCTRG", rsPumpData
1577               SetCombo cmbTEMCVoltage, "TEMCVoltage", rsPumpData
1578           End If

               'write ship to and bill to info
1579           If Not IsNull(rsPumpData.Fields("ShipToCustomer")) Then
1580               txtShpNo.Text = rsPumpData.Fields("ShipToCustomer")
1581           Else
1582               txtShpNo.Text = vbNullString
1583           End If

1584           If Not IsNull(rsPumpData.Fields("BillToCustomer")) Then
1585               txtBilNo.Text = rsPumpData.Fields("BillToCustomer")
1586           Else
1587               txtBilNo.Text = vbNullString
1588           End If

1589           sName = "ImpellerDia"
1590           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1591               sParam = rsPumpData.Fields(sName)
1592           Else
1593               sParam = vbNullString
1594           End If
1595           txtImpellerDia.Text = sParam

1596           sName = "DesignFlow"
1597           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1598               sParam = rsPumpData.Fields(sName)
1599           Else
1600               sParam = vbNullString
1601           End If
1602           txtDesignFlow.Text = sParam

1603           sName = "DesignTDH"
1604           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1605               sParam = rsPumpData.Fields(sName)
1606           Else
1607               sParam = vbNullString
1608           End If
1609           txtDesignTDH.Text = sParam

1610           sName = "SpGr"
1611           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1612               sParam = rsPumpData.Fields(sName)
1613           Else
1614               sParam = vbNullString
1615           End If
1616           txtSpGr.Text = sParam

1617           sName = "Remarks"
1618           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1619               sParam = rsPumpData.Fields(sName)
1620           Else
1621               sParam = vbNullString
1622           End If
1623           txtRemarks.Text = sParam

1624           sName = "SalesOrderNumber"
1625           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1626               sParam = rsPumpData.Fields(sName)
1627           Else
1628               sParam = vbNullString
1629           End If
1630           txtSalesOrderNumber.Text = sParam

1631           sName = "ApplicationFluid"
1632           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1633               sParam = rsPumpData.Fields(sName)
1634           Else
1635               sParam = vbNullString
1636           End If
1637           txtLiquid.Text = sParam

1638           sName = "NPSHFile"
1639           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1640               sParam = rsPumpData.Fields(sName)
1641           Else
1642               sParam = vbNullString
1643           End If
1644           txtNPSHFileLocation.Text = sParam

1645           sName = "RVSPartNo"
1646           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1647               sParam = rsPumpData.Fields(sName)
1648           Else
1649               sParam = vbNullString
1650           End If
1651           txtRVSPartNo.Text = sParam

1652           sName = "CustPN"
1653           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1654               sParam = rsPumpData.Fields(sName)
1655           Else
1656               sParam = vbNullString
1657           End If
1658           txtXPartNum.Text = sParam

1659           sName = "CustPO"
1660           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1661               sParam = rsPumpData.Fields(sName)
1662           Else
1663               sParam = vbNullString
1664           End If
1665           txtCustPONum.Text = sParam

               'make sure table has custpn - see if last three digits of model no are numeric
       '        sName = "SalesOrderNumber"
       '        If rsPumpData.Fields(sName).ActualSize <> 0 Then
       '            If IsNumeric(Right(rsPumpData.Fields("ModelNumber"), 3)) Then 'no sales order no, must be supermarket
       '                rsPumpData.Fields("CustPN") = rsPumpData.Fields("RVSPartNo")
       '            Else
       '                rsPumpData.Fields("CustPN") = rsPumpData.Fields("ModelNumber")
       '            End If
       '        End If

1666           sName = "ApplicationViscosity"
1667           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1668               sParam = Format(rsPumpData.Fields(sName), "#0.00")
1669           Else
1670               sParam = vbNullString
1671           End If
1672           txtViscosity.Text = sParam

       'added from TEMC Inspection Report
1673           sName = "JobNumber"
1674           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1675               sParam = rsPumpData.Fields(sName)
1676           Else
1677               sParam = ""
1678           End If
1679           txtJobNum.Text = sParam

1680           sName = "Phases"
1681           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1682               sParam = rsPumpData.Fields(sName)
1683           Else
1684               sParam = vbNullString
1685           End If
1686           txtNoPhases.Text = sParam

1687           sName = "ThermalClass"
1688           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1689               sParam = rsPumpData.Fields(sName)
1690           Else
1691               sParam = vbNullString
1692           End If
1693           txtThermalClass.Text = sParam

1694           sName = "ExpClass"
1695           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1696               sParam = rsPumpData.Fields(sName)
1697           Else
1698               sParam = vbNullString
1699           End If
1700           txtExpClass.Text = sParam

1701           sName = "NPSHr"
1702           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1703               sParam = rsPumpData.Fields(sName)
1704           Else
1705               sParam = vbNullString
1706           End If
1707           txtNPSHr.Text = sParam

1708           sName = "LiquidTemperature"
1709           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1710               sParam = rsPumpData.Fields(sName)
1711           Else
1712               sParam = vbNullString
1713           End If
1714           txtLiquidTemperature.Text = sParam

1715           sName = "RatedInputPower"
1716           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1717               sParam = rsPumpData.Fields(sName)
1718           Else
1719               sParam = vbNullString
1720           End If
1721           txtRatedInputPower.Text = sParam

1722           sName = "FLCurrent"
1723           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1724               sParam = rsPumpData.Fields(sName)
1725           Else
1726               sParam = vbNullString
1727           End If
1728           txtAmps.Text = sParam

1729           sName = "TEMCFrameNumber"
1730           If rsPumpData.Fields(sName).ActualSize <> 0 Then
1731               sParam = rsPumpData.Fields(sName)
1732           Else
1733               sParam = vbNullString
1734           End If
1735           txtTEMCFrameNumber.Text = sParam

1736           optMfr(0).value = rsPumpData.Fields("ChempumpPump")
1737           optMfr(1).value = Not optMfr(0).value

1738           If rsPumpData.Fields("Field1") = "Feathered" Then
1739               Me.chkSuperMarketFeathered.value = Checked
1740           Else
1741               Me.chkSuperMarketFeathered.value = Unchecked
1742           End If

               'select the testsetup data
1743           qyTestSetup.ActiveConnection = cnPumpData
1744           qyTestSetup.CommandText = "SELECT * FROM TempTestSetupData WHERE (((TempTestSetupData.SerialNumber)='" & _
                                    txtSN.Text & "')) ORDER BY Date"
       '        qyTestSetup.CommandText = "SELECT * FROM TempTestSetupData WHERE (((TempTestSetupData.SerialNumber)='" & _
'               txtSN.Text & "'))"

1745           With rsTestSetup
1746               If .State = adStateOpen Then
1747                   .Close
1748               End If
1749               .CursorLocation = adUseClient
1750               .CursorType = adOpenStatic
1751               .Index = "FindData"
1752               .Open qyTestSetup
1753           End With


               'add the selection of dates to the Test Date combo box
1754           If rsTestSetup.RecordCount <> 0 Then
1755               For I = 0 To cmbTestDate.ListCount - 1
1756                   cmbTestDate.RemoveItem 0
1757               Next I
1758               rsTestSetup.MoveFirst
1759               For I = 1 To rsTestSetup.RecordCount
1760                   cmbTestDate.AddItem rsTestSetup.Fields("Date")
1761                   rsTestSetup.MoveNext
1762               Next I
1763               rsTestSetup.MoveFirst
1764               boFoundTestSetup = True

1765               If rsTestSetup.Fields("Approved") = True Then
1766                   DisableTestSetupDataControls                         'if it's in the real database, don't allow changes here
1767                   boTestDateIsApproved = True
1768                   lblTestDateApproved.Visible = True
1769                   If boCanApprove Then
1770                       cmdApproveTestDate.Caption = "Unapprove this Test Date"
1771                   End If
1772               Else
1773                   EnableTestSetupDataControls                          'it's in the temp database, allow changes
1774                   lblTestDateApproved.Visible = False
1775                   If boCanApprove Then
1776                       cmdApproveTestDate.Caption = "Approve this Test Date"
1777                   End If
1778               End If
1779               cmbTestDate.ListIndex = 0
1780           Else
1781               MsgBox ("There is no Test Setup Data for Serial Number " & txtSN.Text)
1782               boFoundTestSetup = False        'didn't find any data
1783               boFoundTestData = False
1784               cmbTestDate.AddItem Date        'load with today
1785               cmbTestDate.ListIndex = 0       'show the entry
1786               EnableTestSetupDataControls
1787               txtTestRemarks.Text = ""
1788               txtVibAx.Text = ""
1789               txtVibRad.Text = ""
1790               txtThrustBal.Text = ""
1791               txtTEMCTRGReading.Text = ""
1792               txtTEMCFrontThrust.Text = ""
1793               txtTEMCRearThrust.Text = ""
1794               Exit Sub
1795           End If

1796           If cmbTestDate.ListCount = 1 Then       'if there's only one test date, select it
1797           End If
1798           Exit Sub
1799       End If


1800       Do While boUsingEpicor = True   'need a do loop to exit
1801           If boUsingEpicor = True Then
1802               Dim MyRecord As SNRecord
           '            I = InStr(1, txtSN.Text, "-")
           '            If I > 0 Then
1803                   MyRecord = GetEpicorODBCData(txtSN.Text, EpicorConnectionString)
           '            End If
1804               If MyRecord.SONumber = "" Then
1805                   MsgBox ("Not found in Epicor")
1806                   boUsingEpicor = False
1807                   boEpicorFound = False
1808                   Exit Do
1809               End If

1810               If MyRecord.SONumber = 0 Then
1811                   boEpicorFound = False
1812                   boUsingSupermarketTable = True
1813                   boUsingEpicor = False
1814               Else
1815                   boEpicorFound = True
1816                   boUsingSupermarketTable = False
1817               End If

1818               If boEpicorFound = True Then
1819                   boUsingEpicor = False
       '                boEpicorFound = True
1820                   txtSalesOrderNumber.Text = MyRecord.SONumber
1821                   txtLineNumber.Text = MyRecord.SOLine
1822                   txtBilNo.Text = MyRecord.Customer
1823                   txtXPartNum.Text = MyRecord.XPartNum
1824                   txtCustPONum.Text = MyRecord.CustomerPO

1825                   If MyRecord.ShipTo = "" Then
1826                       txtShpNo.Text = MyRecord.Customer
1827                   Else
1828                       txtShpNo.Text = MyRecord.ShipTo
1829                   End If
1830                   txtModelNo.Text = MyRecord.PartNum
1831                   txtModelNo_Change
1832                   txtDesignTDH.Text = MyRecord.TDH
1833                   txtSpGr.Text = MyRecord.SpGr
1834                   txtImpellerDia.Text = MyRecord.ImpellerDiameter
1835                   txtDesignFlow.Text = MyRecord.Flow
1836                   txtNoPhases.Text = MyRecord.Phases
1837                   txtNPSHr.Text = MyRecord.NPSHr
1838                   txtRatedInputPower.Text = MyRecord.RatedInputPower
1839                   txtAmps.Text = MyRecord.FLCurrent
1840                   txtThermalClass.Text = MyRecord.ThermalClass
1841                   txtViscosity.Text = MyRecord.Viscosity
1842                   txtExpClass.Text = MyRecord.ExpClass
1843                   txtLiquidTemperature.Text = MyRecord.LiquidTemp
1844                   txtLiquid.Text = MyRecord.Fluid
1845                   txtJobNum.Text = MyRecord.JobNumber

1846                   For I = 0 To cmbStatorFill.ListCount - 1
1847                       If InStr(1, UCase$(MyRecord.StatorFill), UCase$(cmbStatorFill.List(I))) <> 0 Then
1848                           cmbStatorFill.ListIndex = I
1849                           Exit For
1850                       End If
1851                   Next I

1852                   For I = 0 To cmbCirculationPath.ListCount - 1
1853                       If InStr(1, UCase$(MyRecord.CirculationPath), UCase$(cmbCirculationPath.List(I))) <> 0 Then
1854                           cmbCirculationPath.ListIndex = I
1855                           Exit For
1856                       End If
1857                   Next I

1858                   For I = 0 To cmbDesignPressure.ListCount - 1
1859                       If InStr(1, MyRecord.DesignPressure, cmbDesignPressure.List(I)) <> 0 Then
1860                           cmbDesignPressure.ListIndex = I
1861                           Exit For
1862                       End If
1863                   Next I

1864                   For I = 0 To cmbVoltage.ListCount - 1
1865                       If InStr(1, MyRecord.Voltage, cmbVoltage.List(I)) <> 0 Then
1866                           cmbVoltage.ListIndex = I
1867                           Exit For
1868                       End If
1869                   Next I

1870                   For I = 0 To cmbFrequency.ListCount - 1
1871                       If InStr(1, MyRecord.Frequency, sName) <> 0 Then
1872                           cmbFrequency.ListIndex = I
1873                           Exit For
1874                       End If
1875                   Next I

1876                   For I = 0 To cmbRPM.ListCount - 1
1877                       If InStr(1, MyRecord.RPM, cmbRPM.List(I)) <> 0 Then
1878                           cmbRPM.ListIndex = I
1879                           Exit For
1880                       End If
1881                   Next I

1882                   For I = 0 To cmbSuctDia.ListCount - 1
1883                       If InStr(1, MyRecord.SuctFlangeSize, cmbSuctDia.List(I)) <> 0 Then
1884                           cmbSuctDia.ListIndex = I
1885                           Exit For
1886                       End If
1887                   Next I

1888                   For I = 0 To cmbDischDia.ListCount - 1
1889                       If InStr(1, MyRecord.DischFlangeSize, cmbDischDia.List(I)) <> 0 Then
1890                           cmbDischDia.ListIndex = I
1891                           Exit For
1892                       End If
1893                   Next I

1894                   For I = 0 To cmbTestSpec.ListCount - 1
1895                       If InStr(1, MyRecord.TestProcedure, cmbTestSpec.List(I)) <> 0 Then
1896                           cmbTestSpec.ListIndex = I
1897                           Exit For
1898                       End If
1899                   Next I

1900                   For I = 0 To cmbMotor.ListCount - 1
1901                       If InStr(1, MyRecord.MotorSize, cmbMotor.List(I)) <> 0 Then
1902                           cmbMotor.ListIndex = I
1903                           Exit For
1904                       End If
1905                   Next I


1906               End If
1907           End If
1908       Loop

1909       If boUsingSupermarketTable = True Then
1910           GetSuperMarketPump MyRecord.PartNum, MyRecord.JobNumber
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
1911       End If
' <VB WATCH>
1912       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub cmdModifyBalanceHoleData_Click()
' <VB WATCH>
1913       On Error GoTo vbwErrHandler
' </VB WATCH>
1914       Dim strInput As String
1915       Dim I As Integer
1916       Dim sNumber As Integer
1917       Dim sDia As String
1918       Dim sBC As String

1919       cmdModifyBalanceHoleData.Visible = False

1920       If dgBalanceHoles.SelBookmarks.Count = 0 Then
1921           cmdModifyBalanceHoleData.Visible = False
1922           Exit Sub
1923       End If

1924       rsBalanceHoles.MoveFirst
1925       rsBalanceHoles.Move dgBalanceHoles.SelBookmarks(0) - dgBalanceHoles.FirstRow

1926       sNumber = rsBalanceHoles!Number
1927       If rsBalanceHoles!diameter = 99 Then
1928           sDia = "Slot"
1929       Else
1930           sDia = str(rsBalanceHoles!diameter)
1931       End If
1932       If rsBalanceHoles!boltcircle = 99 Then
1933           sBC = "Unknown"
1934       Else
1935           sBC = str(rsBalanceHoles!boltcircle)
1936       End If


           'get the data for the balance holes
1937       strInput = InputBox("Enter Number of Holes (0 to delete entry)", , sNumber)
1938       If strInput = "" Then
1939           GoTo DeleteIt
1940       End If
1941       sNumber = CInt(strInput)
1942       If Val(sNumber) = 0 Then
1943           GoTo DeleteIt
1944       End If

1945       strInput = InputBox("Enter Decimal Value of Hole Diameter or 'Slot' (For Example, 0.675) ", , sDia)
1946       If strInput <> "" Then
1947           If UCase(strInput) = "SLOT" Then
1948               strInput = 99
1949           End If
1950           sDia = CSng(strInput)
1951       Else
1952           GoTo CancelPressed
1953       End If

1954       strInput = InputBox("Enter Decimal Value of Bolt Circle or 'Unknown' (For Example, 4.525)", , sBC)
1955       If strInput <> "" Then
1956           If UCase(strInput) = "UNKNOWN" Then
1957               strInput = 99
1958           End If
1959           sBC = CSng(strInput)
1960       Else
1961           GoTo CancelPressed
1962       End If

1963       rsBalanceHoles!Number = sNumber
1964       rsBalanceHoles!diameter = sDia
1965       rsBalanceHoles!boltcircle = sBC

1966       rsBalanceHoles.Update
           'rsBalanceHoles.Filter = "SerialNo = '" & frmPLCData.txtSN.Text & "'"

1967       GetBalanceHoleData txtSN.Text, cmbTestDate.Text
       '    rsBalanceHoles.Requery
1968       rsBalanceHoles.MoveLast
1969       dgBalanceHoles.Refresh
1970       chkBalanceHoles.value = 1
1971       rsBalanceHoles.MoveFirst

1972       Exit Sub

1973   CancelPressed:
1974       MsgBox "No New Balance Hole Data Entered", vbOKOnly

1975   DeleteIt:
1976       If (MsgBox("Do you really want to delete this entry?", vbYesNo, "Deleting Balance Hole Data. . .")) = vbYes Then
1977           rsBalanceHoles.Delete
1978           rsBalanceHoles.Update
1979           GetBalanceHoleData txtSN.Text, cmbTestDate.Text
       '        rsBalanceHoles.Requery
1980           If Not rsBalanceHoles.EOF Then
1981               rsBalanceHoles.MoveLast
1982           End If
1983           dgBalanceHoles.Refresh
1984           chkBalanceHoles.value = 1
1985           If Not rsBalanceHoles.BOF Then
1986               rsBalanceHoles.MoveFirst
1987           End If
1988       End If


' <VB WATCH>
1989       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub cmdReport_Click()
           'view/print a report
' <VB WATCH>
1990       On Error GoTo vbwErrHandler
' </VB WATCH>
1991       Dim I As Integer

1992       ExportToExcel

' <VB WATCH>
1993       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub cmdSearchForPump_Click()
' <VB WATCH>
1994       On Error GoTo vbwErrHandler
' </VB WATCH>
1995       LoadCombo frmSearch.cmbSearchModel, "TEMCHydraulics"

1996       frmSearch.Show
' <VB WATCH>
1997       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub cmdSelectSupermarket_Click()
' <VB WATCH>
1998       On Error GoTo vbwErrHandler
' </VB WATCH>
1999       grpSupermarket.Visible = False
' <VB WATCH>
2000       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub cmdWriteSP_Click()
           'write the sp to the plc
' <VB WATCH>
2001       On Error GoTo vbwErrHandler
' </VB WATCH>
2002       Dim rc As String
2003       Dim S As String

           'write the set point data to the PLC
2004           bWrite = True
2005           S = Right$("0000" & txtWriteSPData, 4)
2006           S = Right$(S, 2) & Left$(S, 2)
2007           rc = StringToByteArray(S, ByteBuffer)

2008           DataLength = HexConvert(ByteBuffer, 2)
2009           DataAddress = StringToHexInt("2005")

2010           rc = GetData

2011           bWrite = False
' <VB WATCH>
2012       Exit Sub
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
    End Select
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
2013       On Error GoTo vbwErrHandler
' </VB WATCH>
2014       Static OriginalColor As Long
2015       If btnRunNPSH.Caption = "Run NPSH" Then
2016           btnRunNPSH.Caption = "Cancel NPSH Run"
2017           OriginalColor = btnRunNPSH.BackColor
2018           tmrNPSHr.Enabled = False
2019           btnRunNPSH.BackColor = vbRed
2020           If boCanApprove Then
2021               txtNPSH(5).Visible = True
2022               lbltab4(5).Visible = True
2023           Else
2024               txtNPSH(5).Visible = False
2025               lbltab4(5).Visible = False
2026           End If
2027           WroteNPSHr = False

2028           frmNPSH.Visible = True
2029           txtNPSH(5).Enabled = True
2030           If Val(txtTDH.Text) <= 10 Then
2031               MsgBox "This test will not work starting with this starting TDH.  Ending test...", vbOKOnly, "Flow is 0"
2032               btnRunNPSH.Caption = "Run NPSH"
2033               btnRunNPSH.BackColor = OriginalColor
2034               frmNPSH.Visible = False
2035               Exit Sub
2036           End If
               'load initial values
2037           If DataGrid2.Row = -1 Then
2038               MsgBox "You must write the normal test data to this row before you run NPSH.", vbOKOnly, "Nothing written for this row"
2039               btnRunNPSH.Caption = "Run NPSH"
2040               btnRunNPSH.BackColor = OriginalColor
2041               frmNPSH.Visible = False
2042               Exit Sub
2043           Else
2044               DataGrid2.Row = UpDown1.value - 1
2045           End If

2046           txtNPSH(0).Text = DataGrid2.Columns("Flow")
2047           txtNPSH(3).Text = DataGrid2.Columns("TDH")
2048           txtNPSH(4) = 0
               'txtNPSH(0).Text = txtFlow.Text
               'txtNPSH(3).Text = txtTDH.Text
2049           txtNPSH(4) = 0
2050       Else
2051           btnRunNPSH.Caption = "Run NPSH"
2052           btnRunNPSH.BackColor = OriginalColor
2053           frmNPSH.Visible = False
2054       End If

           'ReportToExcel
' <VB WATCH>
2055       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

    Private Sub updown1_change()
' <VB WATCH>
2056       On Error GoTo vbwErrHandler
' </VB WATCH>
2057       Dim sName As String

2058       If Not rsTestData.BOF Then
2059           rsTestData.MoveFirst
2060       End If

2061       If Not rsTestData.BOF Or Not rsTestData.EOF Then
2062           rsTestData.Move UpDown1.value - 1
2063       End If

2064       sName = "VibrationX"
2065       If rsTestData.Fields(sName).ActualSize <> 0 Then
2066           txtVibAx.Text = rsTestData.Fields(sName)
2067       Else
       '        txtVibAx.Text = vbNullString
2068       End If

2069       sName = "VibrationY"
2070       If rsTestData.Fields(sName).ActualSize <> 0 Then
2071           txtVibRad.Text = rsTestData.Fields(sName)
2072       Else
       '        txtVibRad.Text = vbNullString
2073       End If

2074       sName = "Remarks"
2075       If rsTestData.Fields(sName).ActualSize <> 0 Then
2076           txtTestRemarks.Text = rsTestData.Fields(sName)
2077       Else
       '        txtTestRemarks.Text = vbNullString
2078       End If

2079       sName = "ThrustBalance"
2080       If rsTestData.Fields(sName).ActualSize <> 0 Then
2081           txtThrustBal.Text = rsTestData.Fields(sName)
2082       Else
       '        txtThrustBal.Text = vbNullString
2083       End If

2084       sName = "TEMCTRG"
2085       If rsTestData.Fields(sName).ActualSize <> 0 Then
2086           txtTEMCTRGReading.Text = rsTestData.Fields(sName)
2087       Else
2088           txtTEMCTRGReading.Text = 0
       '        txtTEMCTRGReading.Text = vbNullString
2089       End If

2090       sName = "TEMCFrontThrust"
2091       If rsTestData.Fields(sName).ActualSize <> 0 Then
2092           txtTEMCFrontThrust.Text = rsTestData.Fields(sName)
2093       Else
       '        txtTEMCFrontThrust.Text = vbNullString
2094       End If

2095       sName = "TEMCRearThrust"
2096       If rsTestData.Fields(sName).ActualSize <> 0 Then
2097           txtTEMCRearThrust.Text = rsTestData.Fields(sName)
2098       Else
       '        txtTEMCRearThrust.Text = vbNullString
2099       End If
2100       sName = "TEMCMomentArm"
2101       If rsTestData.Fields(sName).ActualSize <> 0 Then
2102           txtTEMCMomentArm.Text = rsTestData.Fields(sName)
2103       Else
       '        txtTEMCMomentArm.Text = vbNullString
2104       End If
2105       sName = "TEMCThrustRigPressure"
2106       If rsTestData.Fields(sName).ActualSize <> 0 Then
2107           txtTEMCThrustRigPressure.Text = rsTestData.Fields(sName)
2108       Else
       '        txtTEMCThrustRigPressure.Text = vbNullString
2109       End If
2110       sName = "TEMCViscosity"
2111       If rsTestData.Fields(sName).ActualSize <> 0 And rsTestData.Fields(sName) <> 0 Then
2112           txtTEMCViscosity.Text = rsTestData.Fields(sName)
2113       Else
       '        txtTEMCViscosity.Text = vbNullString
2114       End If

2115       CalculateTEMCForce

2116       rsEff.MoveFirst
2117       rsEff.Move UpDown1.value - 1
' <VB WATCH>
2118       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Sub CalculateTEMCForce()
' <VB WATCH>
2119       On Error GoTo vbwErrHandler
' </VB WATCH>
2120       Dim NoOfPoles As Integer
2121       Dim Frequency As Integer
2122       Dim Additions As String
2123       Dim Frame As String
2124       Dim VOverA As Double
2125       Dim Force As Double
2126       Dim Gravity As Double

2127       If Val(txtSpGr.Text) = 0 Then
2128           Gravity = 1
2129       Else
2130           Gravity = CDbl(Val(txtSpGr.Text))
2131       End If

           'show calculated values
2132       If Val(txtTEMCFrontThrust.Text) = 0 Then
2133           If Val(txtTEMCRearThrust.Text) = 0 Then
               'no thrust entered
2134               lblTEMCFrontRear.Visible = False
2135               txtTEMCCalcForce.Text = " "
2136           Else
                   'rear thrust
2137               txtTEMCCalcForce.Text = Format(Gravity * (Val(txtTEMCRearThrust.Text) * Val(txtTEMCMomentArm.Text) - (Val(txtTEMCThrustRigPressure.Text) / 14.223) * 4.5), "##0.0")
2138               lblTEMCFrontRear.Caption = "REAR"
2139               lblTEMCFrontRear.Visible = True
2140           End If
2141       Else
               'front thrust
2142           txtTEMCCalcForce.Text = Format(Gravity * (Val(txtTEMCFrontThrust.Text) * Val(txtTEMCMomentArm.Text) + (Val(txtTEMCThrustRigPressure.Text) / 14.223) * 4.5), "##0.0")
2143           lblTEMCFrontRear.Caption = "FRONT"
2144           lblTEMCFrontRear.Visible = True
2145       End If

2146       If Val(txtTEMCCalcForce.Text) < 0 Then
2147           txtTEMCCalcForce.Text = -txtTEMCCalcForce
2148           lblTEMCFrontRear.Caption = "FRONT"
2149       End If

           'see how many poles we have, it's the next to last number in the frame size
2150       If Len(txtTEMCFrameNumber) > 2 Then
2151           NoOfPoles = 2 * Val(Left$(Right$(txtTEMCFrameNumber.Text, 2), 1))
2152       End If

2153       If cmbTEMCAdditions.ListIndex <> -1 Then
2154           Additions = Mid$(cmbTEMCAdditions.List(cmbTEMCAdditions.ListIndex), 2, 1)
2155           If Additions = "A" Or Additions = "E" Or Additions = "G" Or Additions = "J" Then
2156               Frequency = 60
2157           ElseIf Additions = "B" Or Additions = "F" Or Additions = "H" Or Additions = "K" Then
2158               Frequency = 50
2159           Else
2160               Frequency = 0
2161           End If
2162       End If

2163       If Len(txtTEMCFrameNumber.Text) = 3 Then
2164           If txtTEMCFrameNumber.Text = "529" Then
2165               Frame = "420"
2166           Else
2167               Frame = Left$(txtTEMCFrameNumber, 2) & "0"
2168           End If
2169       Else
2170           Frame = txtTEMCFrameNumber.Text
2171           If Right$(txtTEMCFrameNumber.Text, 1) = "5" Then
2172               Frame = Frame & Left$(lblTEMCFrontRear.Caption, 1)
2173           Else
2174           End If
2175       End If
2176       Force = DLookupA(3, TEMCForceViscosity, 1, Frame)
2177       If Frequency = 60 Then
2178           Force = Force / 1.2
2179       End If
2180       If Val(txtTEMCViscosity.Text) > 1# Then
2181           If (Val(txtTEMCCalcForce.Text) > 3 * Force) Then
2182               lblTEMCPassFail.Visible = True
2183               lblTEMCPassFail.ForeColor = vbRed
2184               lblTEMCPassFail.Caption = "FAIL"
2185           Else
2186               lblTEMCPassFail.Visible = True
2187               lblTEMCPassFail.ForeColor = vbGreen
2188               lblTEMCPassFail.Caption = "PASS"
2189           End If
2190       End If

2191       If (Val(txtTEMCViscosity.Text) > 0.5) And (Val(txtTEMCViscosity.Text) <= 1#) Then
2192           If (Val(txtTEMCCalcForce.Text) > 2 * Force) Then
2193               lblTEMCPassFail.Visible = True
2194               lblTEMCPassFail.ForeColor = vbRed
2195               lblTEMCPassFail.Caption = "FAIL"
2196           Else
2197               lblTEMCPassFail.Visible = True
2198               lblTEMCPassFail.ForeColor = vbGreen
2199               lblTEMCPassFail.Caption = "PASS"
2200           End If
2201       End If

2202       If (Val(txtTEMCViscosity.Text) > 0.3) And (Val(txtTEMCViscosity.Text) <= 0.5) Then
2203           If (Val(txtTEMCCalcForce.Text) > 1.5 * Force) Then
2204               lblTEMCPassFail.Visible = True
2205               lblTEMCPassFail.ForeColor = vbRed
2206               lblTEMCPassFail.Caption = "FAIL"
2207           Else
2208               lblTEMCPassFail.Visible = True
2209               lblTEMCPassFail.ForeColor = vbGreen
2210               lblTEMCPassFail.Caption = "PASS"
2211           End If
2212       End If

2213       If (Val(txtTEMCViscosity.Text) <= 0.3) Then
2214           If (Val(txtTEMCCalcForce.Text) > 1# * Force) Then
2215               lblTEMCPassFail.Visible = True
2216               lblTEMCPassFail.ForeColor = vbRed
2217               lblTEMCPassFail.Caption = "FAIL"
2218           Else
2219               lblTEMCPassFail.Visible = True
2220               lblTEMCPassFail.ForeColor = vbGreen
2221               lblTEMCPassFail.Caption = "PASS"
2222           End If
2223       End If
2224       If NoOfPoles <> 0 Then
2225           VOverA = (DLookupA(2, TEMCForceViscosity, 1, Frame)) / (NoOfPoles * 30 / Frequency)
2226       End If
       '    If Frequency = 60 Then
       '        VOverA = VOverA * 1.2
       '    End If

2227       txtTEMCPVValue.Text = Format(Val(txtTEMCCalcForce.Text) * VOverA, "##0.0")

2228       If Val(txtTEMCFrontThrust.Text) = 0 And Val(txtTEMCRearThrust.Text) = 0 Then
2229           txtTEMCPVValue.Text = ""
2230           txtTEMCCalcForce.Text = ""
2231           lblTEMCPassFail.Visible = False
2232       End If


           'calculate reverse head
2233       txtRevHead.Text = Format(rsTestData.Fields("RBHPress") - rsTestData.Fields("SuctionPressure") * 2.31, "##0.0")
       '    txtRevHead.Text = Format((CDbl(Val(txtAI3Display.Text)) - CDbl(Val(txtSuctionDisplay.Text))) * 2.31, "##0.0")

' <VB WATCH>
2234       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
    Private Sub updown2_change()
' <VB WATCH>
2235       On Error GoTo vbwErrHandler
' </VB WATCH>
2236       Dim Plothead(1, 7) As Single
2237       Dim HeadPlot(7, 1) As Single

2238       Dim PlotEff() As Single
2239       Dim PlotKW() As Single
2240       Dim PlotAmps() As Single

2241       Dim j As Integer

2242       For j = 0 To UpDown2.value - 1
2243           Plothead(0, j) = HeadFlow(0, j)
2244           Plothead(1, j) = HeadFlow(1, j)
2245           HeadPlot(j, 0) = FlowHead(j, 0)
2246           HeadPlot(j, 1) = FlowHead(j, 1)
       '        ReDim Preserve PlotEff(1, j)
       '        PlotEff(0, j) = EffFlow(0, j)
       '        PlotEff(1, j) = EffFlow(1, j)
       '        ReDim Preserve PlotKW(1, j)
       '        PlotKW(0, j) = KWFlow(0, j)
       '        PlotKW(1, j) = KWFlow(1, j)
       '        ReDim Preserve PlotAmps(1, j)
       '        PlotAmps(0, j) = AmpsFlow(0, j)
       '        PlotAmps(1, j) = AmpsFlow(1, j)
2247       Next j

2248       MSChart1 = HeadPlot

' <VB WATCH>
2249       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub DataGrid1_AfterColUpdate(ByVal ColIndex As Integer)
' <VB WATCH>
2250       On Error GoTo vbwErrHandler
' </VB WATCH>
2251       DoEfficiencyCalcs
' <VB WATCH>
2252       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub dgBalanceHoles_SelChange(Cancel As Integer)
' <VB WATCH>
2253       On Error GoTo vbwErrHandler
' </VB WATCH>
2254       If dgBalanceHoles.SelBookmarks.Count = 0 Then
2255           cmdModifyBalanceHoleData.Visible = False
2256       Else
2257           cmdModifyBalanceHoleData.Visible = True
2258       End If
' <VB WATCH>
2259       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub Form_Activate()
' <VB WATCH>
2260       On Error GoTo vbwErrHandler
' </VB WATCH>
2261       If ProgramEnd = True Then
2262           Unload Me
2263       End If
' <VB WATCH>
2264       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub Form_Load()
' <VB WATCH>
2265       On Error GoTo vbwErrHandler
' </VB WATCH>
2266       Dim RetVal As String
2267       Dim sSendStr As String
2268       Dim I As Integer
2269       Dim j As Integer
2270       Dim sTableName As String
2271       Dim WhichServer As String
2272       Dim WhichDatabase As String

2273       ProgramEnd = False
2274       Dim objWMIService, colProcesses
2275       Set objWMIService = GetObject("winmgmts:")
2276       Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name LIKE 'PolarRundown%'")
       '    Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name LIKE 'Excel%'")
2277       If colProcesses.Count > 1 Then
2278           MsgBox "There is already a copy of Polar Rundown running.  You can only have one copy running at a time", vbOKOnly, "Polar Rundown already running"
2279           Dim f As Form
2280           For Each f In Forms
2281               If f.Name <> Me.Name Then
2282                    Unload f
2283               End If
2284           Next
2285           ProgramEnd = True
2286           Exit Sub
2287       Else
2288       End If
2289       Set objWMIService = Nothing
2290       Set colProcesses = Nothing

2291       debugging = 0   'assume not debugging
2292       WhichServer = "Production"     'change to production server
2293       WhichDatabase = "Production"

2294       If UCase$(Left$(GetMachineName, 5)) = "MROSE" Or UCase$(Left$(GetMachineName, 5)) = "ITTES" Then  'if mickey, see if we want to be in debug
2295           I = MsgBox("Debug?", vbYesNo)
2296           If I = vbYes Then
2297               debugging = 1
2298               WhichServer = "Production"
2299               WhichDatabase = "Production"
2300           Else
2301           End If
2302       End If

2303       If debugging Then
       '        GoTo temp
2304       End If
           'see if the mdb file is where it's supposed to be

2305       Dim developmentDatabase As String
2306       developmentDatabase = GetUNCFromLetter("F:") & sDevelopmentDatabase

2307       If Dir(developmentDatabase) = "" Then
2308           MsgBox "Development.mdb does not exist on F:, Please contact IT.", , "No Development Database"
2309           End
2310       End If

           'get the database info from the new mdb file
2311       Dim cnDevelopment As New ADODB.Connection
2312       Dim qyDevelopment As New ADODB.Command
2313       Dim rsDevelopment As New ADODB.Recordset

2314       On Error GoTo CannotConnect

2315       With cnDevelopment
2316           .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & developmentDatabase & ";Persist Security Info=False; Jet OLEDB:Database Password=Access7277word;"
2317           .ConnectionTimeout = 10
2318           .Open
2319       End With

2320   On Error GoTo vbwErrHandler
2321       GoTo Connected

2322   CannotConnect:
2323       MsgBox "Cannot connect with Development.mdb database.  Please contact IT.", , "Cannot find Connection data."
2324       End

2325   Connected:

           'we're connected, get the data for the Epicor SQL server
2326       qyDevelopment.CommandText = "SELECT * FROM Connections WHERE Connections.WhichServer = '" & WhichServer & "' AND WhichDatabase = '" & WhichDatabase & "'"
2327       qyDevelopment.ActiveConnection = cnDevelopment

2328       rsDevelopment.CursorLocation = adUseClient
2329       rsDevelopment.CursorType = adOpenStatic
2330       rsDevelopment.LockType = adLockOptimistic

2331       On Error GoTo NoServerData

2332       rsDevelopment.Open qyDevelopment

2333   On Error GoTo vbwErrHandler
2334       GoTo GotServerData

2335   NoServerData:

2336       MsgBox "Cannot connect with Development.mdb database.  Please contact IT.", , "Cannot find Connection data."
2337       End

2338   GotServerData:

2339       If rsDevelopment.RecordCount <> 1 Then
2340           GoTo NoServerData
2341       End If

           'construct Epicor connection string
2342       EpicorConnectionString = "Driver={" & rsDevelopment.Fields("ODBCDriver") & "};" & _
                                         "Database=" & rsDevelopment.Fields("DatabaseName") & ";" & _
                                         "Server=" & rsDevelopment.Fields("ServerName") & ";" & _
                                         "UID=" & rsDevelopment.Fields("UserName") & ";" & _
                                         "PWD=" & rsDevelopment.Fields("UserPassword") & ";"


           'make sure we can open the SQL database

2343       On Error GoTo CannotOpenEpicorSQLServer

2344       Dim cnTestEpicor As New ADODB.Connection
2345       cnTestEpicor.ConnectionString = EpicorConnectionString
2346       cnTestEpicor.Open
2347       cnTestEpicor.Close
2348       Set cnTestEpicor = Nothing
2349   On Error GoTo vbwErrHandler

2350       GoTo FoundEpicorSQLServer

2351   CannotOpenEpicorSQLServer:
2352       MsgBox "Cannot connect with the Epicor SQL server specified in Development.mdb.  Please contact IT.", , "Cannot connect with Epicor SQL Server"
2353       End

2354   FoundEpicorSQLServer:
           'get data on rundown database
2355       rsDevelopment.Close
2356       qyDevelopment.CommandText = "SELECT * FROM Connections WHERE Connections.WhichServer = 'PolarRundown'"

2357       On Error GoTo NoRundownDatabase

2358       rsDevelopment.Open qyDevelopment

2359       GoTo FoundRundownDatabase

2360   NoRundownDatabase:
2361       MsgBox "Cannot connect with the Pump Rundown database specified in Development.mdb.  Please contact IT.", , "Cannot connect with Epicor SQL Server"
2362       End

2363   FoundRundownDatabase:
2364       If rsDevelopment.RecordCount <> 1 Then
2365           GoTo NoRundownDatabase
2366           End
2367       End If

2368   temp:

2369       If debugging Then
2370           sDataBaseName = "c:\databases\PolarData.mdb"
2371       Else

2372          sDataBaseName = GetUNCFromLetter("F:") & "\Groups\Shared\databases\PolarData.mdb"

       '        sDataBaseName = rsDevelopment.Fields("ServerName") & rsDevelopment.Fields("DatabaseName")

       '        sDataBaseName = sServerName & "f\groups\shared\databases\PumpData 2k.mdb"
2373       End If

2374       Dim tempFSO
2375       Set tempFSO = CreateObject("Scripting.FileSystemObject")
2376       ParentDirectoryName = tempFSO.getparentfoldername(sDataBaseName)
2377       Set tempFSO = Nothing

           'see if we can open the pump rundown database
2378       On Error GoTo NoRundownDatabase
2379       With cnPumpData
       '        .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataBaseName & ";Persist Security Info=False;Jet OLEDB:Database Password=185TitusAve"
2380           .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & sDataBaseName & ";Persist Security Info=False;"
2381           .ConnectionTimeout = 10
2382           .Open
2383       End With
2384   On Error GoTo vbwErrHandler


2385       If debugging = 0 Then
       '        Printer.Orientation = vbPRORLandscape
2386       End If

2387       lblVersion = "Polar Rundown - Version " & App.Major & "." & App.Minor & "." & App.Revision
2388       frmPLCData.Caption = "Polar Rundown"

2389       boFoundPump = False

2390       Me.Show

2391       MSChart1.Plot.Axis(VtChAxisIdX).AxisTitle = "Flow"
2392       MSChart1.Plot.Axis(VtChAxisIdY).AxisTitle = "TDH"
           'MSChart1.Plot.Axis(VtChAxisIdX).AxisGrid.MajorPen = True
           'MSChart1.Plot.Axis(VtChAxisIdY).AxisGrid.MajorPen = True
2393       MSChart1.Plot.UniformAxis = False
2394       MSChart1.Plot.SeriesCollection.Item(1).SeriesMarker.Auto = False
2395       MSChart1.Plot.SeriesCollection.Item(1).Pen.Width = 5
2396       With MSChart1.Plot.SeriesCollection.Item(1).DataPoints.Item(-1).Marker
2397           .Visible = True
2398           .Size = 50
2399           .Style = VtMarkerStyleCircle
2400           .FillColor.Automatic = False
2401           .FillColor.Set 0, 0, 255
2402       End With

           'assure that the timers are off
2403       frmPLCData.tmrGetDDE.Enabled = False

2404       frmPLCData.tmrStartUp.Enabled = False

           'initialize the PLC network
2405       RetVal = NetWorkInitialize()
2406       If RetVal <> 0 Then
2407           MsgBox ("Can't Initialize Network. Exiting...")
2408           End
2409       End If

2410       If debugging = 0 Then
               'load array of plcs
2411           I = 0
2412           Open rsDevelopment.Fields("ServerName") & "PolarPLCAddresses.txt" For Input As 1
2413           While Not EOF(1)
2414               Input #1, Description(I)
2415               For j = 0 To 125
2416                   Input #1, aDevices(I).Address(j)
2417               Next j
2418               Input #1, j
2419               I = I + 1
2420           Wend
2421           Close #1

2422           DeviceCount = I

2423           If Left$(GetMachineName, 2) = "WV" Then  'if in WV, put MWSC first in loop dropdown
2424               Dim k As Integer
2425               For k = 0 To DeviceCount - 1
2426                   If InStr(Description(k), "MWSC") <> 0 Then
2427                       Exit For
2428                   End If
2429               Next k
2430               Description(DeviceCount) = Description(0)
2431               Description(0) = Description(k)
2432               Description(k) = Description(DeviceCount)

2433               aDevices(DeviceCount) = aDevices(0)
2434               aDevices(0) = aDevices(k)
2435               aDevices(k) = aDevices(DeviceCount)

2436           End If

2437           Dim PLCAddress As String
2438           For I = 0 To DeviceCount - 1
2439               PLCAddress = aDevices(I).Address(4) & "." & aDevices(I).Address(5) & "." & aDevices(I).Address(6) & "." & aDevices(I).Address(7)
2440               RetVal = PingSilent(PLCAddress)
2441               If RetVal <> 0 Then
2442                   frmPLCData.cmbPLCLoop.AddItem Description(I)
2443                   frmPLCData.cmbPLCLoop.ItemData(frmPLCData.cmbPLCLoop.NewIndex) = I
2444               End If
2445           Next I
2446       End If

2447       frmPLCData.cmbPLCLoop.AddItem "Add PLC Data Manually"   'enable the controls for manual entry

           'turn on the PLC led

2448       frmPLCData.cmbPLCLoop.ListIndex = 0
2449       frmPLCData.tmrGetDDE.Enabled = True

           'hook up to the various databases

           'copy the template of the database here
           'see if it exists
2450       Dim fdrive As String
2451       fdrive = GetUNCFromLetter("F:")
2452       If Dir(fdrive & "\groups\shared\databases" & sEffDataBaseName) = "" Then
2453           MsgBox "File does not exist at " & fdrive & "\groups\shared\databases" & sEffDataBaseName & ". Please contact IT", vbOKOnly, "Eff.mdb does not exist"
2454       Else
               'Dim FSO As New FileSystemObject
2455           FileCopy fdrive & "\groups\shared\databases" & sEffDataBaseName, App.Path & sEffDataBaseName
2456       End If


2457       With cnEffData
2458           .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & App.Path & sEffDataBaseName & ";Persist Security Info=False"
2459           .Open
2460       End With

           'open some recordsets
2461       rsPumpData.Index = "SerialNumber"
2462       rsTestSetup.Index = "FindData"
2463       rsTestData.Index = "PrimaryKey"
2464       rsPumpData.Open "TempPumpData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
2465       rsTestSetup.Open "TempTestSetupData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
2466       rsTestData.Filter = "SerialNumber = ''"
2467       rsTestData.CursorLocation = adUseClient
2468       rsTestData.Open "TempTestData", cnPumpData, adOpenStatic, adLockOptimistic, adCmdTable
2469       rsEff.CursorLocation = adUseClient
2470       rsEff.Open "Efficiency", cnEffData, adOpenStatic, adLockOptimistic, adCmdTableDirect
2471       qyBalanceHoles.ActiveConnection = cnPumpData
2472       rsBalanceHoles.CursorLocation = adUseClient
2473       rsBalanceHoles.CursorType = adOpenStatic
2474       rsBalanceHoles.LockType = adLockOptimistic
2475       qyMisc.ActiveConnection = cnPumpData
2476       qyMisc.CommandText = "SELECT MiscParameters.ParameterName, MiscParameters.ParameterValue From MiscParameters WHERE (((MiscParameters.ParameterName)='AllowableTDHVariation'));"
2477       rsMisc.CursorLocation = adUseClient
2478       rsMisc.CursorType = adOpenStatic
2479       rsMisc.LockType = adLockBatchOptimistic
2480       rsMisc.Open qyMisc
2481       txtNPSH(5).Text = rsMisc!ParameterValue

2482       If debugging <> 1 Then
2483           FindMagtrols
2484       Else
2485           cmbMagtrol.AddItem "Add Manually"
2486           cmbMagtrol.ItemData(cmbMagtrol.NewIndex) = 99
2487           cmbMagtrol.ListIndex = 0
2488       End If
2489       optKW(1).value = True
2490       optKW_Click (1)


           'blank out data grid
2491       Set DataGrid1.DataSource = rsTestData

           'load the combo boxes
2492       LoadCombo cmbStatorFill, "StatorFill"
2493       LoadCombo cmbCirculationPath, "CirculationPath"
2494       LoadCombo cmbVoltage, "Voltage"
2495       LoadCombo cmbFrequency, "Frequency"
2496       LoadCombo cmbMotor, "Motor"
2497       LoadCombo cmbDesignPressure, "DesignPressure"
2498       LoadCombo cmbRPM, "RPM"
2499       LoadCombo cmbOrificeNumber, "OrificeNumber"
2500       LoadCombo cmbTestSpec, "TestSpecification"
2501       LoadCombo cmbLoopNumber, "LoopNumber"
2502       LoadCombo cmbSuctDia, "SuctionDiameter"
2503       LoadCombo cmbDischDia, "DischargeDiameter"
2504       LoadCombo cmbTachID, "TachID"
2505       LoadCombo cmbAnalyzerNo, "AnalyzerNo"
2506       LoadCombo cmbModel, "Model"
2507       LoadCombo cmbModelGroup, "ModelGroup"
2508       LoadCombo cmbMounting, "Mounting"
2509       LoadCombo cmbPLCNo, "PLCNo"
2510       LoadCombo cmbFlowMeter, "PumpFlowMeter"
2511       LoadCombo cmbSuctionPressureTransducer, "SuctionPressureTransducer"
2512       LoadCombo cmbDischargePressureTransducer, "DischargePressureTransducer"
2513       LoadCombo cmbTemperatureTransducer, "TemperatureTransducer"
2514       LoadCombo cmbCirculationFlowMeter, "CirculationFlowMeter"
           'LoadCombo cmbSupermarketModel, "SupermarketPumpData"

           'load the TEMC combo boxes, too
2515       LoadCombo cmbTEMCAdapter, "TEMCAdapter"
2516       LoadCombo cmbTEMCAdditions, "TEMCAdditions"
2517       LoadCombo cmbTEMCCirculation, "TEMCCirculation"
2518       LoadCombo cmbTEMCDesignPressure, "TEMCDesignPressure"
2519       LoadCombo cmbTEMCNominalDischargeSize, "TEMCNominalDischargeSize"
2520       LoadCombo cmbTEMCDivisionType, "TEMCDivisionType"
2521       LoadCombo cmbTEMCImpellerType, "TEMCImpellerType"
2522       LoadCombo cmbTEMCInsulation, "TEMCInsulation"
2523       LoadCombo cmbTEMCJacketGasket, "TEMCJacketGasket"
2524       LoadCombo cmbTEMCMaterials, "TEMCMaterials"
2525       LoadCombo cmbTEMCModel, "TEMCModel"
2526       LoadCombo cmbTEMCNominalImpSize, "TEMCNominalImpSize"
2527       LoadCombo cmbTEMCOtherMotor, "TEMCOtherMotor"
2528       LoadCombo cmbTEMCNominalSuctionSize, "TEMCNominalSuctionSize"
2529       LoadCombo cmbTEMCVoltage, "TEMCVoltage"
2530       LoadCombo cmbTEMCPumpStages, "TEMCPumpStages"
2531       LoadCombo cmbTEMCTRG, "TEMCTRG"

           'LoadCombo frmSearch.cmbSearchModel, "Model"

           'fill memory arrays for dlookups
2532       FillArrays

           'choose the first tab
2533       frmPLCData.SSTab1.Tab = 0

           'set the grid column names
2534       Dim c As Column
2535       For Each c In DataGrid1.Columns
2536           Select Case c.DataField
               Case "TestDataID"
2537               c.Visible = False
2538           Case "SerialNumber"
2539               c.Visible = False
2540           Case "Date"
2541               c.Visible = False
2542           Case Else ' Show all other columns.
2543               c.Visible = True
2544               c.Alignment = dbgRight
2545           End Select
2546       Next c

2547       Set dgBalanceHoles.DataSource = rsBalanceHoles

2548       For Each c In dgBalanceHoles.Columns
2549           Select Case c.DataField
               Case "BalanceHoleID"
2550               c.Visible = False
2551           Case "SerialNo"
2552               c.Visible = False
2553           Case "Date"
2554               c.Visible = True
2555               c.Alignment = dbgCenter
2556               c.Width = 2000
2557           Case "Number"
2558               c.Visible = True
2559               c.Alignment = dbgCenter
2560               c.Width = 700
2561           Case "Diameter"
2562               c.Visible = False
2563           Case "Diameter1"
2564               c.Caption = "Diameter"
2565               c.Visible = True
2566               c.Alignment = dbgCenter
2567               c.Width = 700
2568           Case "BoltCircle1"
2569               c.Caption = "Bolt Circle"
2570               c.Visible = True
2571               c.Alignment = dbgCenter
2572               c.Width = 800
2573           Case "BoltCircle"
2574               c.Visible = False
2575           Case "SetNo"
2576               c.Visible = False
2577           Case Else ' Show all other columns.
2578               c.Visible = False
2579           End Select
2580       Next c

2581       BlankData

       '    If debugging <> 1 Then
               'get user initials
2582           frmLogin.Show
       '    End If

2583     optMfr(1).value = True
2584     frmMfr.Visible = False

2585       Pressed = True
' <VB WATCH>
2586       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub Form_Unload(Cancel As Integer)
' <VB WATCH>
2587       On Error GoTo vbwErrHandler
' </VB WATCH>
2588       End
' <VB WATCH>
2589       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub Label15_Click()
' <VB WATCH>
2590       On Error GoTo vbwErrHandler
' </VB WATCH>
2591       frmDiagram.Show
' <VB WATCH>
2592       Exit Sub
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
    End Select
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
2593       On Error GoTo vbwErrHandler
' </VB WATCH>

2594       Dim blnEnabled As Boolean

2595       If lblAutoMan(Index).Caption = "Auto" Then
2596           lblAutoMan(Index).Caption = "Man"
2597           blnEnabled = True
2598       Else
2599           lblAutoMan(Index).Caption = "Auto"
2600           blnEnabled = False
2601       End If

2602       Select Case Index
               Case 0
2603               txtFlowDisplay.Enabled = blnEnabled
2604           Case 1
2605               txtSuctionDisplay.Enabled = blnEnabled
2606           Case 2
2607               txtDischargeDisplay.Enabled = blnEnabled
2608           Case 3
2609               txtTemperatureDisplay.Enabled = blnEnabled
2610           Case 4
2611               txtAI1Display.Enabled = blnEnabled
2612           Case 5
2613               txtAI2Display.Enabled = blnEnabled
2614           Case 6
2615               txtAI3Display.Enabled = blnEnabled
2616           Case 7
2617               txtAI4Display.Enabled = blnEnabled
2618       End Select

' <VB WATCH>
2619       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub tmrNPSHr_Timer()
' <VB WATCH>
2620       On Error GoTo vbwErrHandler
' </VB WATCH>
2621       tmrNPSHr.Enabled = False
2622       If frmNPSH.Visible = True Then
2623           btnRunNPSH_Click    'close test
2624       End If
' <VB WATCH>
2625       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub txtNPSH_Change(Index As Integer)
' <VB WATCH>
2626       On Error GoTo vbwErrHandler
' </VB WATCH>
2627       If Index = 5 Then
2628           If frmNPSH.Visible = True Then
2629               If rsMisc.State = adStateOpen Then
2630                   rsMisc.Close
2631               End If
2632               rsMisc.CursorLocation = adUseClient
2633               rsMisc.Open "Select * from MiscParameters WHERE (ParameterName = 'AllowableTDHVariation');", cnPumpData, adOpenStatic, adLockOptimistic, adCmdText
2634               rsMisc.Fields("ParameterValue").value = txtNPSH(5).Text
2635               rsMisc.Update
2636           End If
2637       End If
' <VB WATCH>
2638       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub txtNPSHFileLocation_Click()
' <VB WATCH>
2639       On Error GoTo vbwErrHandler
' </VB WATCH>
2640       Dim sTempDir As String
2641       On Error Resume Next
2642       sTempDir = CurDir    'Remember the current active directory
2643       CommonDialog2.DialogTitle = "Select a directory" 'titlebar
2644       CommonDialog2.InitDir = "\\tei-main-01\f\en\groups\shared\calibration and rundown\npsh\" 'start dir, might be "C:\" or so also
2645       CommonDialog2.filename = "Select a Directory"  'Something in filenamebox
2646       CommonDialog2.Flags = cdlOFNNoValidate + cdlOFNHideReadOnly
2647       CommonDialog2.Filter = "Directories|*.~#~" 'set files-filter to show dirs only
2648       CommonDialog2.CancelError = True 'allow escape key/cancel
2649       CommonDialog2.ShowSave   'show the dialog screen

2650       If Err <> 32755 Then    ' User didn't chose Cancel.
               'Me.SDir.Text = CurDir
2651       End If

       '    ChDir sTempDir  'restore path to what it was at entering

2652   Me.txtNPSHFileLocation.Text = CommonDialog2.filename

' <VB WATCH>
2653       Exit Sub
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
    End Select
' </VB WATCH>
End Sub





Private Sub txtTitle_LostFocus(Index As Integer)
' <VB WATCH>
2654       On Error GoTo vbwErrHandler
' </VB WATCH>

2655       ChangeTitles Index

' <VB WATCH>
2656       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub ChangeTitles(ChannelNo As Integer)
' <VB WATCH>
2657       On Error GoTo vbwErrHandler
' </VB WATCH>
2658       Dim I As Integer
2659       Dim S As String

2660       If txtTitle(ChannelNo).Locked = True Then
2661           Exit Sub
2662       End If

2663       Dim qy As New ADODB.Command
2664       Dim rs As New ADODB.Recordset

2665       qy.ActiveConnection = cnPumpData

           'see if we have an entry in the table
2666       qy.CommandText = "SELECT * FROM AITitles " & _
                             "WHERE (((AITitles.SerialNo)= '" & txtSN.Text & "') " & _
                             "AND ((AITitles.Date)= #" & cmbTestDate.Text & "#) " & _
                             "AND ((AITitles.Channel)=" & ChannelNo & "));"

2667       With rs     'open the recordset for the query
2668           .CursorLocation = adUseClient
2669           .CursorType = adOpenStatic
2670           .LockType = adLockOptimistic
2671           .Open qy
2672       End With

2673       If (rs.BOF = True And rs.EOF = True) Then  'new record
2674           rs.AddNew
2675           rs.Fields("SerialNo") = txtSN.Text
2676           rs.Fields("Date") = cmbTestDate.Text
2677           rs.Fields("Channel") = CByte(ChannelNo)
2678           rs.Fields("Title") = txtTitle(ChannelNo).Text
2679           rs.Update
2680       Else    'we have an entry, modify it
2681           rs.Fields("SerialNo") = txtSN.Text
2682           rs.Fields("Date") = cmbTestDate.Text
2683           rs.Fields("Channel") = CByte(ChannelNo)
2684           rs.Fields("Title") = txtTitle(ChannelNo).Text
2685           rs.Update
2686       End If

2687       rs.Close
2688       Set rs = Nothing
2689       Set qy = Nothing

' <VB WATCH>
2690       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub optKW_Click(Index As Integer)
' <VB WATCH>
2691       On Error GoTo vbwErrHandler
' </VB WATCH>
2692       Select Case Index
               Case 0  'add 3 powers
2693               txtKW.Enabled = False
2694           Case 1  'enter kw
2695               txtKW.Enabled = True
2696           Case 2  'use analog in 4
2697               txtKW.Enabled = False
2698       End Select
' <VB WATCH>
2699       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub optMfr_Click(Index As Integer)
' <VB WATCH>
2700       On Error GoTo vbwErrHandler
' </VB WATCH>
2701       frmTEMC.Visible = optMfr(1).value
2702       frmChempump.Visible = optMfr(0).value
2703       frmTEMCData.Visible = optMfr(1).value
2704       txtModelNo_Change
' <VB WATCH>
2705       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub tmrGetDDE_Timer()
' <VB WATCH>
2706       On Error GoTo vbwErrHandler
' </VB WATCH>

       'get here every second... get plc and magtrol data

2707       Dim sSendStr As String
2708       Dim I As Integer
2709       Dim VoltMul As Double

2710       If Calibrating Then
2711           Exit Sub
2712       End If

2713       If debugging Then
               'Exit Sub
2714       End If


2715       If boPLCOperating = True Then
2716           frmPLCData.shpGetPLCData.Visible = True    'turn the PLC led on

               'convert the plc data into real numbers
               'the following data are type real
2717           txtFlow.Text = ConvertToReal("4050")
2718           txtSuction.Text = ConvertToReal("4052")
2719           txtDischarge.Text = ConvertToReal("4054")
2720           txtTemperature.Text = ConvertToReal("4056")

2721           txtValvePosition.Text = ConvertToLong("2004")

2722           frmPLCData.txtTC1.Text = ConvertToLong("2200")
2723           frmPLCData.txtTC2.Text = ConvertToLong("2202")
2724           frmPLCData.txtTC3.Text = ConvertToLong("2204")
2725           frmPLCData.txtTC4.Text = ConvertToLong("2206")

2726           frmPLCData.txtAI1.Text = ConvertToReal("4060")
2727           frmPLCData.txtAI2.Text = ConvertToReal("4062")
2728           frmPLCData.txtAI3.Text = ConvertToReal("4064")
2729           frmPLCData.txtAI4.Text = ConvertToReal("4066")

2730           frmPLCData.txtPCoef.Text = ConvertToLong("4036")
2731           frmPLCData.txtICoef.Text = ConvertToLong("4037")
2732           frmPLCData.txtDCoef.Text = ConvertToLong("4040")

2733           frmPLCData.txtSetPoint.Text = ConvertToLong("4035")
2734           frmPLCData.txtInHg.Text = ConvertToLong("1460")


               'modify the data from PLC format to format that we can use
               'and update the screen
2735           If txtFlowDisplay.Enabled = False Then
2736               frmPLCData.txtFlowDisplay = Format$(txtFlow.Text, "###0.00")
2737           End If
2738           If txtSuctionDisplay.Enabled = False Then
2739               frmPLCData.txtSuctionDisplay = Format$((txtSuction.Text) / 10, "##0.00")
2740           End If
2741           If txtDischargeDisplay.Enabled = False Then
2742               frmPLCData.txtDischargeDisplay = Format$(txtDischarge.Text, "##0.00")
2743           End If
2744           If txtTemperatureDisplay.Enabled = False Then
2745               frmPLCData.txtTemperatureDisplay = Format$(txtTemperature.Text, "##0.00")
2746           End If
2747           frmPLCData.txtValvePositionDisplay = (txtValvePosition.Text)

2748           frmPLCData.txtTC1Display = Format$((txtTC1.Text) / 10, "##0.0")
2749           frmPLCData.txtTC2Display = Format$((txtTC2.Text) / 10, "##0.0")
2750           frmPLCData.txtTC3Display = Format$((txtTC3.Text) / 10, "##0.0")
2751           frmPLCData.txtTC4Display = Format$((txtTC4.Text) / 10, "##0.0")

2752           If txtAI1Display.Enabled = False Then
2753               frmPLCData.txtAI1Display = Format$(txtAI1.Text, "##0.00")
2754           End If
2755           If txtAI2Display.Enabled = False Then
2756               frmPLCData.txtAI2Display = Format$(txtAI2.Text, "##0.00")
2757           End If
2758           If txtAI3Display.Enabled = False Then
2759               frmPLCData.txtAI3Display = Format$(txtAI3.Text, "##0.00")
2760           End If
2761           If txtAI4Display.Enabled = False Then
2762               frmPLCData.txtAI4Display = Format$(txtAI4.Text, "##0.00")
2763           End If

2764           frmPLCData.txtSetPointDisplay = (txtSetPoint.Text)

2765           frmPLCData.txtInHgDisplay = Format$(txtInHg.Text / 100, "00.00")

2766           frmPLCData.shpGetPLCData.Visible = False   'turn the PLC led off

2767           frmPLCData.shpGetMagtrolData.Visible = True 'turn the Magtrol led on
2768       End If

2769       If boMagtrolOperating = True Then


               'get the data from the Magtrol
2770           If Right(cmbMagtrol.List(cmbMagtrol.ListIndex), 4) = "5300" Then
2771               sSendStr = vbCrLf
2772               sData = Space$(68)
2773               VoltMul = Sqr(3)
2774           Else
2775               sSendStr = "OT" & vbCrLf
2776               sData = Space$(183)
2777               VoltMul = 1#
2778           End If

2779           On Error GoTo noresponse
2780           If UsingNatInst Then
2781               ibwrt iUD, sSendStr
2782               ibrd iUD, sData

                   'parse the Magrol response
       '            vResponse = CWGPIB1.Tasks("Number Parser").Parse(sData)
2783           Else
                   'Dim Databack As String
2784               sData = TCP.SendGetData("OT")
2785           End If

2786               Dim vSplit() As String
2787               vSplit = Split(Right(sData, Len(sData) - 1), ",")
2788               ReDim vResponse(UBound(vSplit))
2789               For I = 0 To UBound(vSplit) - 1
2790                   vResponse(I) = CDbl(vSplit(I))
2791               Next I

               'format the parsed response
2792           Dim dd As String
2793           dd = "- -"

2794           If Not IsEmpty(vResponse) Then
               '8 entries for 5300 and 12 for the 6530
2795               If UBound(vResponse) = 8 Or UBound(vResponse) = 12 Then
                       'put the responses into the correct text box
2796                   txtV1.Text = Format$(VoltMul * vResponse(1), "###0.0")   'we get back phase voltage and we want line voltage

2797                   Select Case vResponse(0)
                           Case Is < 1
2798                           txtI1.Text = Format$(vResponse(0), "0.0000")
2799                       Case Is < 10
2800                           txtI1.Text = Format$(vResponse(0), "0.000")
2801                       Case Is < 100
2802                           txtI1.Text = Format$(vResponse(0), "00.00")
2803                       Case Else
2804                           txtI1.Text = Format$(vResponse(0), "000.0")
2805                   End Select

2806                   Select Case vResponse(3)
                           Case Is < 1
2807                           txtI2.Text = Format$(vResponse(3), "0.0000")
2808                       Case Is < 10
2809                           txtI2.Text = Format$(vResponse(3), "0.000")
2810                       Case Is < 100
2811                           txtI2.Text = Format$(vResponse(3), "00.00")
2812                       Case Else
2813                           txtI2.Text = Format$(vResponse(3), "000.0")
2814                   End Select

2815                   Select Case vResponse(6)
                           Case Is < 1
2816                           txtI3.Text = Format$(vResponse(6), "0.0000")
2817                       Case Is < 10
2818                           txtI3.Text = Format$(vResponse(6), "0.000")
2819                       Case Is < 100
2820                           txtI3.Text = Format$(vResponse(6), "00.00")
2821                       Case Else
2822                           txtI3.Text = Format$(vResponse(6), "000.0")
2823                   End Select

2824                   txtP1.Text = Format$(vResponse(2) / 1000, "##0.00")     '/ by 1000 to show kW
2825                   txtV2.Text = Format$(VoltMul * vResponse(4), "###0.0")
                       'txtI2.Text = Format$(vResponse(3), "###0.0")
2826                   txtP2.Text = Format$(vResponse(5) / 1000, "##0.00")
2827                   txtV3.Text = Format$(VoltMul * vResponse(7), "###0.0")
                       'txtI3.Text = Format$(vResponse(6), "###0.0")
2828                   txtP3.Text = Format$(vResponse(8) / 1000, "##0.00")
2829                   If (vResponse(0) * vResponse(1) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7)) <> 0 Then
                           'if we have some measured current
                           'pf = sum of power/sum of VA
2830                       If Right(cmbMagtrol.List(cmbMagtrol.ListIndex), 4) = "5300" Then
                               'add kw responses and / by 1000 to get to kW
2831                           txtKW.Text = (vResponse(2) + vResponse(5) + vResponse(8)) / 1000
2832                           txtPF.Text = Format$(100 * (vResponse(2) + vResponse(5) + vResponse(8)) / (vResponse(1) * vResponse(0) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7)), "0.00")
2833                       Else
2834                           txtKW.Text = (vResponse(2) + vResponse(8)) / 1000
2835                           txtPF.Text = Format$(100 * (vResponse(2) + vResponse(8)) / ((Sqr(3) / 3) * (vResponse(1) * vResponse(0) + vResponse(3) * vResponse(4) + vResponse(6) * vResponse(7))), "0.00")
2836                       End If
2837                       Select Case Val(txtKW.Text)
                               Case Is < 1
2838                               txtKW.Text = Format$(txtKW.Text, "0.00000")
2839                           Case Is < 10
2840                               txtKW.Text = Format$(txtKW.Text, "0.0000")
2841                           Case Is < 100
2842                               txtKW.Text = Format$(txtKW.Text, "00.000")
2843                           Case Else
2844                               txtKW.Text = Format$(txtKW.Text, "000.00")
2845                       End Select
2846                   Else
2847                       txtPF = dd
2848                   End If
2849               Else
                       'no response, show all -- in text boxes
2850                   txtV1.Text = dd
2851                   txtI1.Text = dd
2852                   txtP1.Text = dd
2853                   txtV2.Text = dd
2854                   txtI2.Text = dd
2855                   txtP2.Text = dd
2856                   txtV3.Text = dd
2857                   txtI3.Text = dd
2858                   txtP3.Text = dd
2859                   txtPF = dd
2860                   txtKW = dd
2861               End If
2862           End If
2863       Else    'magtrol not operating
2864           Dim dbl As Double

2865           If optKW(0).value = True Then   'add 3 powers
2866               txtKW.Text = Val(txtP1.Text) + Val(txtP2.Text) + Val(txtP3.Text)
2867           End If
2868           If optKW(1).value = True Then   'enter kw
2869               txtP1.Text = Val(txtKW.Text) / 3
2870               txtP2.Text = Val(txtKW.Text) / 3
2871               txtP3.Text = Val(txtKW.Text) / 3
2872           End If
2873           If optKW(2).value = True Then   'use ai4
2874               txtKW.Text = txtAI4Display.Text
2875               txtP1.Text = Val(txtKW.Text) / 3
2876               txtP2.Text = Val(txtKW.Text) / 3
2877               txtP3.Text = Val(txtKW.Text) / 3
2878           End If

2879           dbl = Val(txtV1.Text) * Val(txtI1.Text)
2880           dbl = dbl + Val(txtV2.Text) * Val(txtI2.Text)
2881           dbl = dbl + Val(txtV3.Text) * Val(txtI3.Text)
2882           If dbl <> 0 Then
2883               txtPF.Text = Format$((Val(txtKW.Text) * 1000 * 3 * 100 / (dbl * Sqr(3))), "0.00")
2884           End If
2885       End If

2886   noresponse:
2887   On Error GoTo vbwErrHandler
2888       frmPLCData.shpGetMagtrolData.Visible = False   'turn the Magtrol led off

           'update the little PLC chart
2889       For I = 1 To 99
2890           vPlot(0, I) = vPlot(0, I + 1)
2891           vPlot(1, I) = vPlot(1, I + 1)
2892       Next I
2893       vPlot(0, 100) = txtSetPointDisplay
2894       vPlot(1, 100) = txtFlowDisplay

           'do NPSH stuff
2895       Dim SuctVelHead As Single
2896       Dim DischVelHead As Single
2897       Dim Conversion As Single
2898       Dim SuctionPSIA As Single
2899       Dim DischargePSIA As Single
2900       Dim VaporPress As Single
2901       Dim SpecVolume As Single
2902       Dim NPSHa As Single
2903       Dim NPSHr As Single
2904       Dim TDH As Single
2905       Dim pd As Single


           'velocity head
2906       If cmbSuctDia.ListIndex = -1 Then   'if no suction diameter chosen
2907           SuctVelHead = 0
2908       Else
       '        pd = DLookup("ActualDia", "PipeDiameters", "ID = " & cmbSuctDia.ListIndex + 1)
2909           pd = DLookupA(ActualColNo, PipeDiameters, IDColNo, cmbSuctDia.ItemData(cmbSuctDia.ListIndex) + 1)
2910           SuctVelHead = (0.002592 * Val(txtFlow) ^ 2) / (pd ^ 4)
2911       End If

2912       If cmbDischDia.ListIndex = -1 Then     'if no discharge diameter chosen
2913           DischVelHead = 0
2914       Else
       '        pd = DLookup("ActualDia", "PipeDiameters", "ID = " & cmbDischDia.ListIndex + 1)
2915           pd = DLookupA(ActualColNo, PipeDiameters, IDColNo, cmbDischDia.ItemData(cmbDischDia.ListIndex) + 1)
2916           DischVelHead = (0.002592 * Val(txtFlow) ^ 2) / (pd ^ 4)
2917       End If

           'convert gauges to absolute
2918       If txtInHgDisplay.Text = "" Then
2919           Conversion = 0
2920       Else
2921           Conversion = txtInHgDisplay * 0.491
2922       End If

2923       SuctionPSIA = Val(txtSuctionDisplay) + Conversion
2924       DischargePSIA = Val(txtDischargeDisplay) + Conversion


           'lookup vapor pressure and specific volume in the arrays that we made
           'if temp is out of range, say so and exit
2925       If Val(txtTemperatureDisplay) < 40 Or Val(txtTemperatureDisplay) > 165 Then
2926           txtNPSHa = 0
2927           Exit Sub
2928       Else
2929           I = Val(txtTemperatureDisplay) - 40
       '        VaporPress = DLookup("VaporPressure", "VaporPressure", "ID = " & I)
       '        SpecVolume = DLookup("SpecificVolume", "VaporPressure", "ID= " & I)
2930           VaporPress = DLookupA(VaporPressureColNo, VaporPressure, IDColNo, I)
2931           SpecVolume = DLookupA(SpecificVolumeColNo, VaporPressure, IDColNo, I)
2932       End If

2933       If Not ((txtSuctHeight = "") Or (txtDischHeight = "") Or Not IsNumeric(txtSuctHeight) Or Not IsNumeric(txtDischHeight)) Then
               'NPSHa
2934           NPSHa = (144 * SpecVolume * (SuctionPSIA - VaporPress)) + (txtSuctHeight / 12) + SuctVelHead
       '        NPSHa = CalcTDH(DischargePSIA, SuctionPSIA, 0, DischVelHead, 0, txtTemperature)
2935           txtNPSHa = Format$(NPSHa, "##0.00")

               'tdh
2936           TDH = CalcTDH(DischargePSIA, SuctionPSIA, 0, (DischVelHead - SuctVelHead), (txtDischHeight / 12) - (txtSuctHeight / 12), txtTemperatureDisplay)
2937           txtTDH = Format$(TDH, "##0.00")

2938           If frmNPSH.Visible = True Then
2939               If Val(txtTDH.Text) > 0 Then
2940                   txtNPSH(2).Text = Format(100 * Val(txtTDH.Text) / Val(txtNPSH(3).Text), "##0.00")
2941                   txtNPSH(1).Text = Format(100 * Val(txtFlow.Text) / Val(txtNPSH(0).Text), "##0.00")
                       'check for tdh variation
2942                   If Abs(Val(txtNPSH(1)) - 100) > Val(txtNPSH(5).Text) Then
2943                       MsgBox "The TDH value has varied more than " & txtNPSH(5) & " %. NPSHr data will NOT be written to the data table", vbOKOnly, "TDH variation too large"
2944                       btnRunNPSH_Click
2945                   Else    'tdh variation small
2946                       If Val(txtNPSH(2).Text) <= 97 Then
                               'btnRunNPSH_Click
                               'write the npsh and save
2947                           If WroteNPSHr = False Then
2948                               txtNPSH(4).Text = txtNPSHa.Text
2949                               rsTestData!NPSHr = txtNPSHa.Text
2950                               rsTestData.Update
2951                               rsEff!NPSHr = txtNPSHa.Text
2952                               rsEff.Update
2953                               WroteNPSHr = True
2954                               tmrNPSHr.Interval = 5000
2955                               tmrNPSHr.Enabled = True
2956                           End If
2957                       End If  'val < 97
2958                   End If  'check for tdh variation
2959               End If 'val tdh <=0
2960           Else    'frm not visible
                   'txtNPSHa = Format$(0, "##0.00")
2961           End If  'if frm visible

2962       Else
2963           txtNPSHa = 0
2964       End If
' <VB WATCH>
2965       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub tmrStartUp_Timer()
           'we waited for a while, disable the timer
' <VB WATCH>
2966       On Error GoTo vbwErrHandler
' </VB WATCH>
2967       tmrStartUp.Enabled = False
' <VB WATCH>
2968       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Public Function SetCombo(cmbComboName As ComboBox, sName As String, rs As ADODB.Recordset)
       'set the pump parameter combo box to the right data based upon
       'the number in the database
' <VB WATCH>
2969       On Error GoTo vbwErrHandler
' </VB WATCH>

2970       Dim I As Integer
2971       Dim sParam As String
2972       Dim qy As New ADODB.Command
2973       Dim rs1 As New ADODB.Recordset

2974       If rs.Fields(sName).ActualSize <> 0 Then     'if there's an entry
2975           sParam = rs.Fields(sName)                'get the index number
2976           qy.ActiveConnection = cnPumpData
2977           qy.CommandText = "SELECT * FROM " & sName & " WHERE " & sName & " = " & sParam
2978           Set rs1 = qy.Execute()                                  'get the record for the index number

2979           If rs1.BOF = True And rs1.EOF = True Then
2980               cmbComboName.ListIndex = -1                             'else, remove any pointer
2981               Exit Function
2982           End If

2983           For I = 0 To cmbComboName.ListCount - 1                     'go through the combobox entries
2984               If cmbComboName.ItemData(I) = rs1.Fields(0) Then     'see when we find the desired index number
2985                   cmbComboName.ListIndex = I                                              'if we do, set the combo box
2986                   Exit For                                            'and we're done
2987               End If
2988               cmbComboName.ListIndex = -1                             'else, remove any pointer
2989           Next I
2990       Else
2991           cmbComboName.ListIndex = -1
2992       End If

2993       Exit Function
' <VB WATCH>
2994       Exit Function
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
    End Select
' </VB WATCH>
End Function
Private Function SetComboTestSetup(cmbComboName As ComboBox, sFieldName As String, sTableName As String, rs As ADODB.Recordset)
       'set the pump parameter combo box to the right data based upon
       'the number in the database
' <VB WATCH>
2995       On Error GoTo vbwErrHandler
' </VB WATCH>

       'same as setcombo, except here we also pass in the field name

2996       Dim I As Integer
2997       Dim sParam As String
2998       Dim qy As New ADODB.Command
2999       Dim rs1 As New ADODB.Recordset

3000       If rs.Fields(sFieldName).ActualSize <> 0 Then
               'if plc number, adjust plcaddress id numbers 1 and 2 to plc 8 and 9 respectively
3001           If sTableName = "CirculationFlowMeter" Then
                   'sParam = rs.Fields(sFieldName) + 7
3002               sParam = rs.Fields(sFieldName)
3003               If Val(sParam) < 4 Then
3004                   sParam = str(Val(sParam) + 4)
3005                   rs.Fields(sFieldName) = sParam
3006               End If
3007           Else
3008               sParam = rs.Fields(sFieldName)
3009           End If
3010           qy.ActiveConnection = cnPumpData
3011           qy.CommandText = "SELECT * FROM " & sTableName & " WHERE " & sTableName & " = " & sParam
3012           Set rs1 = qy.Execute()

3013           For I = 0 To cmbComboName.ListCount - 1
3014               If cmbComboName.ItemData(I) = rs1.Fields(0) Then
3015                   cmbComboName.ListIndex = I
3016                   Exit For
3017               End If
3018               cmbComboName.ListIndex = -1
3019           Next I
3020       Else
3021           cmbComboName.ListIndex = -1
3022       End If

3023       Exit Function
' <VB WATCH>
3024       Exit Function
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
    End Select
' </VB WATCH>
End Function

Private Sub DisablePumpDataControls()
           'disable the pump data controls cause we're just showing what we found
' <VB WATCH>
3025       On Error GoTo vbwErrHandler
' </VB WATCH>

3026       txtSalesOrderNumber.Enabled = False
3027       frmMfr.Enabled = False
3028       txtShpNo.Enabled = False
3029       txtBilNo.Enabled = False
3030       txtDesignFlow.Enabled = False
3031       txtDesignTDH.Enabled = False

3032       frmMiscPumpData.Enabled = False

3033       txtModelNo.Enabled = False
3034       txtImpellerDia.Enabled = False

3035       frmTEMC.Enabled = False
3036       frmChempump.Enabled = False

3037       txtRemarks.Enabled = False
3038       Me.cmdAddNewTestDate.Visible = False

3039       cmdEnterPumpData.Enabled = False

' <VB WATCH>
3040       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub DisableTestSetupDataControls()
' <VB WATCH>
3041       On Error GoTo vbwErrHandler
' </VB WATCH>

3042       cmbTestSpec.Enabled = False
3043       txtWho.Enabled = False
3044       txtRMA.Enabled = False

3045       frmLoopAndXducer.Enabled = False
3046       frmElecData.Enabled = False
3047       frmPerfMods.Enabled = False
3048       frmOtherFiles.Enabled = False
3049       frmInstrumentTags.Enabled = False
3050       frmTAndI.Enabled = False
3051       frmThrustBalMods.Enabled = False
3052       txtTestSetupRemarks.Enabled = False

3053       cmdEnterTestSetupData.Enabled = False
3054       cmbPLCNo.Enabled = False
' <VB WATCH>
3055       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub DisableTestDataControls()
' <VB WATCH>
3056       On Error GoTo vbwErrHandler
' </VB WATCH>

3057       cmbPLCLoop.Enabled = False
3058       frmPumpData.Enabled = False
3059       frmThermocouples.Enabled = False
3060       frmAI.Enabled = False
3061       frmMagtrol.Enabled = False
3062       fmrMiscTestData.Enabled = False
3063       frmPLCMisc.Enabled = False
3064       DataGrid1.Enabled = False
3065       DataGrid2.Enabled = False
3066       cmdEnterTestData.Enabled = False

' <VB WATCH>
3067       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub EnableTestSetupDataControls()
' <VB WATCH>
3068       On Error GoTo vbwErrHandler
' </VB WATCH>

3069       cmbTestSpec.Enabled = True
3070       txtWho.Enabled = True
3071       txtRMA.Enabled = True

3072       frmLoopAndXducer.Enabled = True
3073       frmElecData.Enabled = True
3074       frmPerfMods.Enabled = True
3075       frmOtherFiles.Enabled = True
3076       frmInstrumentTags.Enabled = True
3077       frmTAndI.Enabled = True
3078       frmThrustBalMods.Enabled = True
3079       txtTestSetupRemarks.Enabled = True

3080       cmdEnterTestSetupData.Enabled = True
3081       cmbPLCNo.Enabled = True
' <VB WATCH>
3082       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub EnableTestDataControls()
' <VB WATCH>
3083       On Error GoTo vbwErrHandler
' </VB WATCH>

3084       cmbPLCLoop.Enabled = True
3085       frmPumpData.Enabled = True
3086       frmThermocouples.Enabled = True
3087       frmAI.Enabled = True
3088       frmMagtrol.Enabled = True
3089       fmrMiscTestData.Enabled = True
3090       frmPLCMisc.Enabled = True
3091       DataGrid1.Enabled = True
3092       DataGrid2.Enabled = True
3093       cmdEnterTestData.Enabled = True

' <VB WATCH>
3094       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub EnablePumpDataControls()
           'disable the pump data controls cause we're just showing what we found
' <VB WATCH>
3095       On Error GoTo vbwErrHandler
' </VB WATCH>

3096       txtSalesOrderNumber.Enabled = True
3097       frmMfr.Enabled = True
3098       txtShpNo.Enabled = True
3099       txtBilNo.Enabled = True
3100       txtDesignFlow.Enabled = True
3101       txtDesignTDH.Enabled = True

3102       frmMiscPumpData.Enabled = True

3103       txtModelNo.Enabled = True
3104       txtImpellerDia.Enabled = True

3105       frmTEMC.Enabled = True
3106       frmChempump.Enabled = True

3107       txtRemarks.Enabled = True
3108       Me.cmdAddNewTestDate.Visible = True

3109       cmdEnterPumpData.Enabled = True

' <VB WATCH>
3110       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub EnableMagtrolFields()
' <VB WATCH>
3111       On Error GoTo vbwErrHandler
' </VB WATCH>
3112       txtV1.Enabled = True
3113       txtV2.Enabled = True
3114       txtV3.Enabled = True
3115       txtI1.Enabled = True
3116       txtI2.Enabled = True
3117       txtI3.Enabled = True
3118       txtP1.Enabled = True
3119       txtP2.Enabled = True
3120       txtP3.Enabled = True
3121       optKW(0).Visible = True
3122       optKW(1).Visible = True
3123       optKW(2).Visible = True
3124       optKW(1).value = True
3125       optKW_Click (1)
' <VB WATCH>
3126       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub DisableMagtrolFields()
' <VB WATCH>
3127       On Error GoTo vbwErrHandler
' </VB WATCH>
3128       txtV1.Enabled = False
3129       txtV2.Enabled = False
3130       txtV3.Enabled = False
3131       txtI1.Enabled = False
3132       txtI2.Enabled = False
3133       txtI3.Enabled = False
3134       txtP1.Enabled = False
3135       txtP2.Enabled = False
3136       txtP3.Enabled = False
3137       txtKW.Enabled = False
3138       optKW(0).Visible = False
3139       optKW(1).Visible = False
3140       optKW(2).Visible = False
' <VB WATCH>
3141       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub EnablePLCFields()
' <VB WATCH>
3142       On Error GoTo vbwErrHandler
' </VB WATCH>
3143       frmPLCData.txtAI1Display.Enabled = True
3144       frmPLCData.txtAI2Display.Enabled = True
3145       frmPLCData.txtAI3Display.Enabled = True
3146       frmPLCData.txtAI4Display.Enabled = True
3147       frmPLCData.txtTC1Display.Enabled = True
3148       frmPLCData.txtTC2Display.Enabled = True
3149       frmPLCData.txtTC3Display.Enabled = True
3150       frmPLCData.txtTC4Display.Enabled = True
3151       frmPLCData.txtFlowDisplay.Enabled = True
3152       frmPLCData.txtSuctionDisplay.Enabled = True
3153       frmPLCData.txtDischargeDisplay.Enabled = True
3154       frmPLCData.txtTemperatureDisplay.Enabled = True
3155       frmPLCData.txtInHgDisplay.Enabled = True
' <VB WATCH>
3156       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub DisablePLCFields()
' <VB WATCH>
3157       On Error GoTo vbwErrHandler
' </VB WATCH>
3158       frmPLCData.txtAI1Display.Enabled = False
3159       frmPLCData.txtAI2Display.Enabled = False
3160       frmPLCData.txtAI3Display.Enabled = False
3161       frmPLCData.txtAI4Display.Enabled = False
3162       frmPLCData.txtTC1Display.Enabled = False
3163       frmPLCData.txtTC2Display.Enabled = False
3164       frmPLCData.txtTC3Display.Enabled = False
3165       frmPLCData.txtTC4Display.Enabled = False
3166       frmPLCData.txtFlowDisplay.Enabled = False
3167       frmPLCData.txtSuctionDisplay.Enabled = False
3168       frmPLCData.txtDischargeDisplay.Enabled = False
3169       frmPLCData.txtTemperatureDisplay.Enabled = False
3170       frmPLCData.txtInHgDisplay.Enabled = False
' <VB WATCH>
3171       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub BlankData()
' <VB WATCH>
3172       On Error GoTo vbwErrHandler
' </VB WATCH>
3173       txtShpNo.Text = vbNullString
3174       txtBilNo.Text = vbNullString
3175       txtModelNo.Text = vbNullString
3176       cmbMotor.ListIndex = -1
3177       cmbStatorFill.ListIndex = -1
3178       cmbVoltage.ListIndex = -1
3179       cmbDesignPressure.ListIndex = -1
3180       cmbFrequency.ListIndex = -1
3181       cmbCirculationPath.ListIndex = -1
3182       cmbRPM.ListIndex = -1
3183       cmbModel.ListIndex = -1
3184       cmbModelGroup.ListIndex = -1
3185       txtSpGr.Text = vbNullString
3186       txtImpellerDia.Text = vbNullString
3187       txtEndPlay.Text = vbNullString
3188       txtGGap.Text = vbNullString
3189       txtDesignFlow.Text = vbNullString
3190       txtDesignTDH.Text = vbNullString
3191       txtOtherMods.Text = vbNullString
3192       txtRemarks.Text = vbNullString
3193       txtSalesOrderNumber.Text = vbNullString
3194       txtTestSetupRemarks.Text = vbNullString
3195       txtNPSHFile.Text = vbNullString
3196       txtPicturesFile.Text = vbNullString
3197       txtVibrationFile.Text = vbNullString
       '    cmbOrificeNumber.ListIndex = 18
       '    cmbTestSpec.ListIndex = 6       'default = Rev7
3198       cmbLoopNumber.ListIndex = -1
3199       cmbSuctDia.ListIndex = -1
3200       cmbDischDia.ListIndex = -1
3201       cmbTachID.ListIndex = -1
3202       cmbAnalyzerNo.ListIndex = -1
3203       txtTestRemarks.Text = vbNullString
3204       txtHDCor.Text = 0
3205       txtDischHeight.Text = 0
3206       txtSuctHeight.Text = 0
3207       txtKWMult.Text = 1
3208       txtWho.Text = LogInInitials
3209       txtRMA.Text = vbNullString
3210       frmPLCData.chkNPSH.value = 0
3211       frmPLCData.chkPictures.value = 0
3212       frmPLCData.chkVibration.value = 0
3213       cmbFlowMeter.ListIndex = -1
3214       cmbSuctionPressureTransducer.ListIndex = -1
3215       cmbDischargePressureTransducer.ListIndex = -1
3216       cmbTemperatureTransducer.ListIndex = -1
3217       cmbCirculationFlowMeter.ListIndex = -1
3218       frmPLCData.chkBalanceHoles.value = 0
3219       frmPLCData.chkCircOrifice.value = 0
3220       frmPLCData.txtCircOrifice = vbNullString
3221       frmPLCData.txtImpTrim = vbNullString
3222       frmPLCData.txtOrifice = vbNullString
3223       frmPLCData.chkFeathered.value = Unchecked
3224       frmPLCData.chkTrimmed.value = 0
3225       frmPLCData.chkCircOrifice.value = 0
3226       frmPLCData.txtThrustBal = vbNullString
3227       frmPLCData.txtRPM = vbNullString
3228       frmPLCData.txtVibAx = vbNullString
3229       frmPLCData.txtVibRad = vbNullString
3230       frmPLCData.txtTEMCTRGReading = vbNullString
3231       dgBalanceHoles.Visible = False
3232       Me.txtLineNumber.Text = vbNullString
3233       Me.txtNPSHr.Text = vbNullString
3234       Me.txtRatedInputPower.Text = vbNullString
3235       Me.txtAmps.Text = vbNullString
3236       Me.txtThermalClass.Text = vbNullString
3237       Me.txtViscosity.Text = vbNullString
3238       Me.txtExpClass.Text = vbNullString
3239       Me.txtNoPhases.Text = vbNullString
3240       Me.txtLiquidTemperature.Text = vbNullString
3241       Me.txtJobNum.Text = vbNullString
3242       Me.txtTEMCFrameNumber.Text = vbNullString
3243       Me.txtLiquid.Text = vbNullString
3244       Me.chkSuperMarketFeathered.value = Unchecked
3245       Me.txtRVSPartNo.Text = vbNullString
' <VB WATCH>
3246       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub AddTestData()
' <VB WATCH>
3247       On Error GoTo vbwErrHandler
' </VB WATCH>
3248       Dim I As Integer
3249       Dim sFilter As String

3250       ClearEff
3251       rsEff.MoveFirst

3252       For I = 1 To 8
3253           rsTestData.AddNew
3254           rsTestData!SerialNumber = txtSN
3255           rsTestData!Date = cmbTestDate.List(cmbTestDate.ListIndex)
3256           rsTestData!testnumber = I
3257           rsTestData!DataWritten = False
3258           rsTestData.Update
3259           DoEfficiencyCalcs
3260           rsEff.MoveNext
3261           rsTestData.MoveNext
3262       Next I
3263       boFoundTestData = True
           'rsTestData.Update
3264       rsTestData.Requery
3265       rsTestData.Resync

          'select the entries from testdata
3266       sFilter = "SerialNumber='" & txtSN.Text & "' AND Date=#" & cmbTestDate.Text & "#"

3267       rsTestData.Filter = sFilter

3268       Set DataGrid1.DataSource = rsTestData

           ' fix the datagrid

3269       Dim c As Column
3270       For Each c In DataGrid1.Columns
3271          Select Case c.DataField
              Case "TestDataID"
3272             c.Visible = False
3273          Case "SerialNumber"
3274             c.Visible = False
3275          Case "Date"
3276             c.Visible = False
3277          Case Else ' Hide all other columns.
3278             c.Visible = True
3279             c.Alignment = dbgRight
3280          End Select
3281       Next c

3282       rsEff.Requery
3283       DataGrid1.Refresh
3284       DataGrid2.Refresh

' <VB WATCH>
3285       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub DoEfficiencyCalcs()
' <VB WATCH>
3286       On Error GoTo vbwErrHandler
' </VB WATCH>
3287       Dim KW As Single, VI As Single, VITemp As Single
3288       Dim Vave As Single, Iave As Single
3289       Dim I As Integer
3290       Dim j As Integer
3291       Dim HeightDiff As Single

3292       If Not IsNull(rsTestData.Fields("TotalPower")) Then
3293           KW = rsTestData.Fields("TotalPower")
3294       Else
               'if we wrote data with an old version, we will not have written total power
               'if total power = 0 and the three individual powers are not 0, add them

3295           If rsTestData.Fields("PowerA") > 0 Then
3296               If rsTestData.Fields("PowerB") > 0 Then
3297                   If rsTestData.Fields("PowerC") > 0 Then
3298                       KW = rsTestData.Fields("PowerA") + rsTestData.Fields("PowerB") + rsTestData.Fields("PowerC")
3299                   End If
3300               End If
3301           End If
3302      End If

3303       I = 0
3304       Vave = 0
3305       Iave = 0
3306       If Not IsNull(rsTestData.Fields("VoltageA")) And Not IsNull(rsTestData.Fields("CurrentA")) Then
3307           VI = rsTestData.Fields("VoltageA") * rsTestData.Fields("CurrentA")
3308           Vave = rsTestData.Fields("VoltageA")
3309           Iave = rsTestData.Fields("CurrentA")
3310           If VI <> 0 Then
3311               I = I + 1
3312           End If
3313       End If
3314       If Not IsNull(rsTestData.Fields("VoltageB")) And Not IsNull(rsTestData.Fields("CurrentB")) Then
3315           VITemp = rsTestData.Fields("VoltageB") * rsTestData.Fields("CurrentB")
3316           If VITemp <> 0 Then
3317               I = I + 1
3318               VI = VI + VITemp
3319               Vave = Vave + rsTestData.Fields("VoltageB")
3320               Iave = Iave + rsTestData.Fields("CurrentB")
3321           End If
3322       End If
3323       If Not IsNull(rsTestData.Fields("VoltageC")) And Not IsNull(rsTestData.Fields("CurrentC")) Then
3324           VITemp = rsTestData.Fields("VoltageC") * rsTestData.Fields("CurrentC")
3325           If VITemp <> 0 Then
3326               I = I + 1
3327               VI = VI + VITemp
3328               Vave = Vave + rsTestData.Fields("VoltageC")
3329               Iave = Iave + rsTestData.Fields("CurrentC")
3330           End If
3331       End If
3332       If KW = 0 Then
3333           For j = 1 To rsEff.Fields.Count - 1
3334               rsEff.Fields(j) = 0
3335           Next j
       '        Exit Sub
3336       End If
3337       If VI <> 0 Then
3338           rsEff.Fields("Volts") = Vave / I
3339           rsEff.Fields("Amps") = Iave / I
3340           rsEff.Fields("PowerFactor") = 1000 * I * KW / (VI * Sqr(3))
3341           rsEff.Fields("PowerFactor") = 100 * rsEff.Fields("PowerFactor")
3342       Else
3343           rsEff.Fields("PowerFactor") = 0
3344       End If

3345       If optMfr(0).value = True Then
3346           If cmbStatorFill.ListIndex = -1 Then
3347               rsEff.Fields("MotorEfficiency") = Format$(0, "0.00")

3348           Else
3349               rsEff.Fields("Motorefficiency") = Format$(Round(MotorEfficiency(KW, cmbMotor.ItemData(cmbMotor.ListIndex), cmbStatorFill.ItemData(cmbStatorFill.ListIndex)), 1), "00.0")
       '            rsEff.Fields("Motorefficiency") = Format$(Round(MotorEfficiency(KW, cmbMotor.ListIndex, cmbStatorFill.ListIndex), 1), "00.0")
3350           End If
3351       Else
3352           rsEff.Fields("MotorEfficiency") = Format$(Round(TEMCMotorEfficiency(KW, txtTEMCFrameNumber.Text, 460, RatedKW), 1), "00.0")
3353       End If

3354       Dim sHDCor As Single
3355       Dim sDisc As Single
3356       Dim sSuct As Single
3357       If IsNull(rsTestSetup.Fields("HDCor")) Then
3358           sHDCor = 0
3359       Else
3360           sHDCor = rsTestSetup.Fields("HDCor")
3361       End If
3362       If IsNull(rsTestSetup.Fields("DischargeGageHeight")) Then
3363           sDisc = 0
3364       Else
3365           sDisc = rsTestSetup.Fields("DischargeGageHeight")
3366       End If
3367       If IsNull(rsTestSetup.Fields("SuctionGageHeight")) Then
3368           sSuct = 0
3369       Else
3370           sSuct = rsTestSetup.Fields("SuctionGageHeight")
3371       End If
3372       HeightDiff = sHDCor + sDisc / 12 - sSuct / 12
3373       If (cmbDischDia.ListIndex <> -1 And cmbSuctDia.ListIndex <> -1) Then
3374           rsEff.Fields("VelocityHead") = CalcVelHead(rsTestData.Fields("Flow"), cmbDischDia.ItemData(cmbDischDia.ListIndex) + 1, cmbSuctDia.ItemData(cmbSuctDia.ListIndex) + 1)
3375       End If
       '    rsEff.Fields("VelocityHead") = CalcVelHead(rsTestData.Fields("Flow"), cmbDischDia.ListIndex + 1, cmbSuctDia.ListIndex + 1)
3376       rsEff.Fields("TDH") = CalcTDH(rsTestData.Fields("DischargePressure"), rsTestData.Fields("SuctionPressure"), rsTestData.Fields("SuctionInHg"), rsEff.Fields("VelocityHead"), HeightDiff, rsTestData.Fields("TemperatureSuction"))
3377       rsEff.Fields("ElecHP") = 1000 * KW / 746
       '    If (DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))) <> 0 And KW <> 0) Then
3378           If Int(rsTestData.Fields("TemperatureSuction")) >= 40 Then
3379               If (DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))) <> 0 And KW <> 0) Then
           '        rsEff.Fields("LiquidHP") = (rsEff.Fields("TDH") * rsTestData.Fields("Flow") * DLookup("TDHCorr", "TempCorrection", "Temp = 68")) / (3960 * DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))))
3380               rsEff.Fields("LiquidHP") = (rsEff.Fields("TDH") * rsTestData.Fields("Flow") * DLookupA(TDHColNo, TempCorrection, TempColNo, 68)) / (3960 * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))))
           '        rsEff.Fields("OverallEfficiency") = (0.189 * rsTestData.Fields("Flow") * rsEff.Fields("TDH") * DLookup("TDHCorr", "TempCorrection", "Temp = 68")) / (10 * KW * DLookup("TDHCorr", "TempCorrection", "Temp = " & Int(rsTestData.Fields("TemperatureSuction"))))
3381               rsEff.Fields("OverallEfficiency") = (0.189 * rsTestData.Fields("Flow") * rsEff.Fields("TDH") * DLookupA(TDHColNo, TempCorrection, TempColNo, 68)) / (10 * KW * DLookupA(TDHColNo, TempCorrection, TempColNo, Int(rsTestData.Fields("TemperatureSuction"))))
3382               If rsEff.Fields("MotorEfficiency") <> 0 Then
3383                   rsEff.Fields("HydraulicEfficiency") = 100 * rsEff.Fields("OverallEfficiency") / rsEff.Fields("MotorEfficiency")
3384               Else
3385                   rsEff.Fields("HydraulicEfficiency") = 0
3386               End If
3387           Else
3388               rsEff.Fields("LiquidHP") = 0
3389               rsEff.Fields("OverallEfficiency") = 0
3390           End If

3391       Else
3392           rsEff.Fields("LiquidHP") = 0
3393           rsEff.Fields("OverallEfficiency") = 0
3394       End If


3395       I = rsEff.AbsolutePosition
3396       If Not IsNull(rsTestData.Fields("Flow")) Then
3397           rsEff.Fields("Flow") = rsTestData.Fields("Flow")
3398           HeadFlow(0, I - 1) = rsTestData.Fields("Flow")
3399           HeadFlow(1, I - 1) = rsEff.Fields("TDH")
3400           FlowHead(I - 1, 0) = rsTestData.Fields("Flow")
3401           FlowHead(I - 1, 1) = rsEff.Fields("TDH")

       '        EffFlow(0, i - 1) = rsTestData.Fields("Flow")
       '        EffFlow(1, i - 1) = rsEff.Fields("OverallEfficiency")
       '        KWFlow(0, i - 1) = rsTestData.Fields("Flow")
       '        KWFlow(1, i - 1) = KW
       '        AmpsFlow(0, i - 1) = rsTestData.Fields("Flow")
       '        AmpsFlow(1, i - 1) = rsEff.Fields("Amps")
3402       Else
3403           HeadFlow(0, I - 1) = 0
3404           HeadFlow(1, I - 1) = 0
3405           FlowHead(I - 1, 0) = 0
3406           FlowHead(I - 1, 1) = 0

       '        EffFlow(0, i - 1) = 0
       '        EffFlow(1, i - 1) = 0
       '        KWFlow(0, i - 1) = 0
       '        KWFlow(1, i - 1) = 0
       '        AmpsFlow(0, i - 1) = 0
       '        AmpsFlow(1, i - 1) = 0
3407       End If

3408       Dim Plothead(1, 7) As Single
3409       Dim HeadPlot(7, 1) As Single
           'ReDim Preserve Plothead(1, j)
           'ReDim Preserve HeadPlot(j, 1)

       '    Dim PlotEff() As Single
       '    Dim PlotKW() As Single
       '    Dim PlotAmps() As Single
       '    ReDim PlotHead(0, 0)
       '    ReDim PlotEff(0, 0)
       '    ReDim PlotKW(0, 0)
       '
3410       For j = 0 To UpDown2.value - 1
       '        If HeadFlow(1, j) <> 0 Then
       '            ReDim Preserve Plothead(1, j)
       '            ReDim Preserve HeadPlot(j, 1)
3411               Plothead(0, j) = HeadFlow(0, j)
3412               Plothead(1, j) = HeadFlow(1, j)
3413               HeadPlot(j, 0) = FlowHead(j, 0)
3414               HeadPlot(j, 1) = FlowHead(j, 1)
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
3415       Next j




       '    SetGraphMax (Plothead())
       '    If UBound(PlotHead()) <> 0 Then

       'fix 4/29/19

3416           MSChart1.ChartData = HeadPlot

       '    End If

           'copy fields for reports
3417       rsEff.Fields("DischPress") = rsTestData.Fields("Dischargepressure")
3418       rsEff.Fields("SuctPress") = rsTestData.Fields("Suctionpressure")
       '    rsEff.Fields("Volts") = rsTestData.Fields("VoltageA")
       '    rsEff.Fields("Amps") = rsTestData.Fields("CurrentA")
3419       rsEff.Fields("KW") = KW
3420       rsEff.Fields("Freq") = rsTestData.Fields("VFDFrequency")
3421       rsEff.Fields("RPM") = rsTestData.Fields("RPM")
3422       rsEff.Fields("Pos") = rsTestData.Fields("ThrustBalance")
3423       rsEff.Fields("NPSHa") = rsTestData.Fields("NPSHa")
3424       rsEff.Fields("NPSHr") = rsTestData.Fields("NPSHr")
3425       rsEff.Fields("InputPower") = rsTestData.Fields("TotalPower")
3426       rsEff.Fields("Temperature") = rsTestData.Fields("TemperatureSuction")
3427       rsEff.Fields("CircFlow") = rsTestData.Fields("CircFlow")
3428       rsEff.Fields("VibrationX") = rsTestData.Fields("VibrationX")
3429       rsEff.Fields("VibrationY") = rsTestData.Fields("VibrationY")
3430       rsEff.Fields("CurrentA") = rsTestData.Fields("CurrentA")
3431       rsEff.Fields("CurrentB") = rsTestData.Fields("CurrentB")
3432       rsEff.Fields("CurrentC") = rsTestData.Fields("CurrentC")
3433       rsEff.Fields("VoltageA") = rsTestData.Fields("VoltageA")
3434       rsEff.Fields("VoltageB") = rsTestData.Fields("VoltageB")
3435       rsEff.Fields("VoltageC") = rsTestData.Fields("VoltageC")
3436       rsEff.Fields("TC1") = rsTestData.Fields("TC1")
3437       rsEff.Fields("TC2") = rsTestData.Fields("TC2")
3438       rsEff.Fields("TC3") = rsTestData.Fields("TC3")
3439       rsEff.Fields("TC4") = rsTestData.Fields("TC4")
3440       rsEff.Fields("RBHTemp") = rsTestData.Fields("RBHTemp")
3441       rsEff.Fields("RBHPress") = rsTestData.Fields("RBHPress")
3442       rsEff.Fields("AI4") = rsTestData.Fields("AI4")
3443       rsEff.Fields("Remarks") = rsTestData.Fields("Remarks")
3444       rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCFrontThrust")
3445       rsEff.Fields("TEMCRearThrust") = rsTestData.Fields("TEMCRearThrust")
3446       rsEff.Fields("TEMCTRG") = rsTestData.Fields("TEMCTRG")
3447       rsEff.Fields("TEMCThrustRigPressure") = rsTestData.Fields("TEMCThrustRigPressure")
3448       rsEff.Fields("TEMCMomentArm") = rsTestData.Fields("TEMCMomentArm")
3449       rsEff.Fields("TEMCViscosity") = rsTestData.Fields("TEMCViscosity")
3450       If Not IsNull(rsEff.Fields("TEMCFrontThrust")) Then
3451           txtTEMCFrontThrust.Text = rsEff.Fields("TEMCFrontThrust")
3452       End If
3453       If Not IsNull(rsEff.Fields("TEMCREarThrust")) Then
3454           txtTEMCRearThrust.Text = rsEff.Fields("TEMCREarThrust")
3455       End If
3456       If (Not IsNull(rsEff.Fields("TEMCViscosity"))) And (rsEff.Fields("TEMCViscosity") <> 0) Then
3457           txtTEMCViscosity.Text = rsEff.Fields("TEMCViscosity")
3458       End If
3459       If Not IsNull(rsTestData.Fields("TEMCThrustRigPressure")) Then
3460           txtTEMCThrustRigPressure.Text = rsTestData.Fields("TEMCThrustRigPressure")
3461       End If
3462       If Not IsNull(rsTestData.Fields("TEMCMomentArm")) Then
3463           txtTEMCMomentArm.Text = rsTestData.Fields("TEMCMomentArm")
3464       End If

        '   If Not IsNull(Me.txtAI3Display.Text) Then
        '       Me.txtAI3Display = rsTestData.Fields("RBHPress")
        '   End If

3465       CalculateTEMCForce

3466       If Not IsNull(txtTEMCCalcForce.Text) Then
3467           rsEff.Fields("TEMCCalculatedForce") = Val(txtTEMCCalcForce.Text)
3468       Else
3469           rsEff.Fields("TEMCCalculatedForce") = 0
3470       End If

3471       If Not IsNull(txtTEMCPVValue.Text) Then
3472           rsEff.Fields("TEMCPV") = Val(txtTEMCPVValue.Text)
3473       Else
3474           rsEff.Fields("TEMCPV") = 0
3475       End If

3476       If Val(txtTEMCFrontThrust.Text) <> 0 Then
3477           rsEff.Fields("TEMCFR") = "F"
       '        rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCFrontThrust")
3478       Else
3479           If Val(txtTEMCRearThrust.Text) = 0 Then
                   'no thrust
3480               rsEff.Fields("TEMCFR") = " "
3481               rsEff.Fields("TEMCFrontThrust") = 0
3482           Else
3483               rsEff.Fields("TEMCFR") = "R"
       '            rsEff.Fields("TEMCFrontThrust") = rsTestData.Fields("TEMCRearThrust")
3484           End If
3485       End If

3486       rsEff.Fields("TEMCForceDirection") = Left(lblTEMCFrontRear.Caption, 1)

3487       rsEff.Update
3488       DataGrid2.Refresh


' <VB WATCH>
3489       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub ClearEff()
       '    Dim I As Integer, j As Integer
' <VB WATCH>
3490       On Error GoTo vbwErrHandler
' </VB WATCH>
3491       Dim qy As New ADODB.Command

3492       If rsEff.State = adStateOpen Then
3493           If Not (rsEff.BOF = True Or rsEff.EOF = True) Then
3494               rsEff.CancelUpdate
3495           End If
3496           rsEff.Close
3497       End If
3498       qy.ActiveConnection = cnEffData
3499       qy.CommandText = "DROP TABLE Efficiency"
3500       rsEff.Open qy
3501       qy.CommandText = "SELECT EfficiencyOrg.* INTO Efficiency FROM EfficiencyOrg;"
3502       rsEff.Open qy
3503       rsEff.Open "Efficiency", cnEffData, adOpenStatic, adLockOptimistic, adCmdTableDirect

3504       rsEff.Requery
3505       DataGrid2.Refresh

3506       Dim c As Column
3507       For Each c In DataGrid2.Columns
3508           c.Alignment = dbgCenter
3509           c.Width = 750
3510           Select Case c.ColIndex
                   Case 1
3511                   c.Caption = "Flow"
3512                   c.NumberFormat = "###0.00"
3513               Case 2
3514                   c.Caption = "TDH"
3515                   c.NumberFormat = "00.0"
3516               Case 3
3517                   c.Caption = "Overall Eff"
3518                   c.NumberFormat = "00.00"
3519                   c.Width = 850
3520               Case 4
3521                   c.Caption = "PF"
3522                   c.NumberFormat = "00.0"
3523               Case 5
3524                   c.Caption = "Vel Head"
3525                   c.NumberFormat = "00.00"
3526               Case 6
3527                   c.Caption = "Elec HP"
3528                   c.NumberFormat = "#00.0"
3529               Case 7
3530                   c.Caption = "Liq HP"
3531                   c.NumberFormat = "#00.0"
3532               Case Else
3533                   c.Visible = False
3534           End Select
3535       Next c

' <VB WATCH>
3536       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Function JustAlphaNumeric(char As String) As String
' <VB WATCH>
3537       On Error GoTo vbwErrHandler
' </VB WATCH>
3538       Select Case Asc(char)
               Case 42             ' *
3539               JustAlphaNumeric = char
3540           Case 48 To 57       ' 0 - 9
3541               JustAlphaNumeric = char
3542           Case 65 To 90       ' A - Z
3543               JustAlphaNumeric = char
3544           Case 97 To 122      ' a - z
3545               JustAlphaNumeric = UCase(char)
3546           Case Else
3547               JustAlphaNumeric = ""
3548       End Select
' <VB WATCH>
3549       Exit Function
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
    End Select
' </VB WATCH>
End Function



Private Sub txtI1_Change()
' <VB WATCH>
3550       On Error GoTo vbwErrHandler
' </VB WATCH>
3551       txtI2.Text = txtI1.Text
3552       txtI3.Text = txtI1.Text
' <VB WATCH>
3553       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub txtModelNo_Change()
' <VB WATCH>
3554       On Error GoTo vbwErrHandler
' </VB WATCH>
3555       Dim I As Integer
3556       Dim S As String
3557       Dim sFull As String
3558       Dim boDone As Boolean
3559       Dim boRepeat As Boolean

3560       Static bo3Digits As Boolean         '3 digits in frame number
3561       Static bo2Digits As Boolean         '2 digits in stages

3562       If optMfr(0).value = True Then
3563           Exit Sub
3564       End If

3565       cmbTEMCAdapter.ListIndex = -1
3566       cmbTEMCAdditions.ListIndex = -1
3567       cmbTEMCCirculation.ListIndex = -1
3568       cmbTEMCDesignPressure.ListIndex = -1
3569       cmbTEMCNominalDischargeSize.ListIndex = -1
3570       cmbTEMCDivisionType.ListIndex = -1
3571       cmbTEMCImpellerType.ListIndex = -1
3572       cmbTEMCInsulation.ListIndex = -1
3573       cmbTEMCJacketGasket.ListIndex = -1
3574       cmbTEMCMaterials.ListIndex = -1
3575       cmbTEMCModel.ListIndex = -1
3576       cmbTEMCNominalImpSize.ListIndex = -1
3577       cmbTEMCOtherMotor.ListIndex = -1
3578       cmbTEMCPumpStages.ListIndex = -1
3579       cmbTEMCNominalSuctionSize.ListIndex = -1
3580       cmbTEMCTRG.ListIndex = -1
3581       cmbTEMCVoltage.ListIndex = -1


           'first, get rid of spaces, dashes, etc

3582       S = ""
3583       For I = 1 To Len(txtModelNo.Text)
3584           S = S & JustAlphaNumeric(Mid$(txtModelNo.Text, I, 1))
3585       Next I

           'next, fill out the model number to it's max length of 24 characters

3586       boDone = False
3587       boRepeat = False

3588       Do While Not boDone
3589           sFull = ""
3590           For I = 1 To Len(S)
3591               Select Case I
                       Case 1
                           'type
3592                       sFull = sFull & Mid$(S, I, 1)
3593                   Case 2
                           'adapter
3594                       If IsNumeric(Mid$(S, I, 1)) Then
3595                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
3596                           boRepeat = True
3597                           Exit For
3598                       Else
3599                           sFull = sFull & Mid$(S, I, 1)
3600                           boRepeat = False
3601                       End If
3602                   Case 3
                           'materials
3603                       sFull = sFull & Mid$(S, I, 1)
3604                   Case 4
                       'design pressure
3605                       sFull = sFull & Mid$(S, I, 1)
3606                   Case 5
                       'motor frame number - digit 1
3607                       sFull = sFull & Mid$(S, I, 1)
3608                   Case 6
                       'motor frame number - digit 2
3609                       sFull = sFull & Mid$(S, I, 1)
3610                   Case 7
                       'motor frame number - digit 3
3611                       sFull = sFull & Mid$(S, I, 1)
3612                   Case 8
                       'motor frame number - digit 4
3613                       If IsNumeric(Mid$(S, I, 1)) Then
3614                           sFull = sFull & Mid$(S, I, 1)
3615                           boRepeat = False
3616                       Else    '3 digits
       '                        s = Left$(s, i - 1) & "*" & Right$(s, Len(s) - i + 1)
3617                           S = Left$(S, I - 4) & "0" & Right$(S, Len(S) - I + 4)
3618                           boRepeat = True
3619                           Exit For
3620                       End If
3621                   Case 9
                       'insulation
3622                       sFull = sFull & Mid$(S, I, 1)
3623                   Case 10
                       'voltage
3624                       sFull = sFull & Mid$(S, I, 1)
3625                   Case 11
                       'other motor specs
3626                       If Mid$(S, I, 1) = "M" Or Mid$(S, I, 1) = "R" Or Mid$(S, I, 1) = "L" Or Mid$(S, I, 1) = "G" Or Mid$(S, I, 1) = "N" Then
3627                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
3628                           boRepeat = True
3629                           Exit For
3630                       Else
3631                           sFull = sFull & Mid$(S, I, 1)
3632                           boRepeat = False
3633                       End If
3634                   Case 12
                       ' TRG
3635                       sFull = sFull & Mid$(S, I, 1)
3636                   Case 13
                       'Nominal discharge - digit 1
3637                       sFull = sFull & Mid$(S, I, 1)
3638                   Case 14
                       'nominal discharge - digit 2
3639                       sFull = sFull & Mid$(S, I, 1)
3640                   Case 15
                       'nominal suction - digit 1
3641                       sFull = sFull & Mid$(S, I, 1)
3642                   Case 16
                       'nominal suction - digit 2
3643                       sFull = sFull & Mid$(S, I, 1)
3644                   Case 17
                       'nominal impeller size
3645                       sFull = sFull & Mid$(S, I, 1)
3646                   Case 18
                       'impeller type
3647                       If Mid$(S, I, 1) <> "*" Then
3648                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
3649                           boRepeat = True
3650                           Exit For
3651                       Else
3652                           sFull = sFull & Mid$(S, I, 1)
3653                           boRepeat = False
3654                       End If
3655                   Case 19
                       'Division type
3656                       If IsNumeric(Mid$(S, I, 1)) Then
3657                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
3658                           boRepeat = True
3659                           Exit For
3660                       Else
3661                           sFull = sFull & Mid$(S, I, 1)
3662                           boRepeat = False
3663                       End If
3664                   Case 20
                       'pump stages - digit 1
3665                       sFull = sFull & Mid$(S, I, 1)
3666                   Case 21
                       'pump jacket
3667                       If Mid$(S, I, 1) = "A" Or Mid$(S, I, 1) = "B" Or Mid$(S, I, 1) = "E" Or Mid$(S, I, 1) = "F" Or _
                                             Mid$(S, I, 1) = "G" Or Mid$(S, I, 1) = "H" Or Mid$(S, I, 1) = "J" Or Mid$(S, I, 1) = "K" Then
3668                           S = Left$(S, I - 1) & "*" & Right$(S, Len(S) - I + 1)
3669                           boRepeat = True
3670                       Else
3671                           sFull = sFull & Mid$(S, I, 1)
3672                           boRepeat = False
3673                       End If
3674                   Case 22
                       'additions
3675                         sFull = sFull & Mid$(S, I, 1)
3676                   Case 23
                       'circulation
3677                         sFull = sFull & Mid$(S, I, 1)
3678               End Select
3679           Next I
3680           If Not boRepeat Then
3681               boDone = True
3682           End If
3683       Loop

3684       For I = 1 To Len(sFull)
3685           Select Case I
                   Case 1
3686                   ParseTEMCModelNo cmbTEMCModel, Mid$(sFull, I, 1)
3687               Case 2
3688                   ParseTEMCModelNo cmbTEMCAdapter, Mid$(sFull, I, 1)
3689               Case 3
3690                   ParseTEMCModelNo cmbTEMCMaterials, Mid$(sFull, I, 1)
3691               Case 4
3692                   ParseTEMCModelNo cmbTEMCDesignPressure, Mid$(sFull, I, 1)
3693               Case 5
3694                       If Val(Mid$(sFull, I, 1)) = 0 Then
3695                           txtTEMCFrameNumber.Text = Mid$(sFull, 6, 3)
3696                       Else
3697                           txtTEMCFrameNumber.Text = Mid$(sFull, 5, 4)
3698                       End If
3699               Case 9
3700                       ParseTEMCModelNo cmbTEMCInsulation, Mid$(sFull, I, 1)
3701               Case 10
3702                       ParseTEMCModelNo cmbTEMCVoltage, Mid$(sFull, I, 1)
3703               Case 11
3704                       ParseTEMCModelNo cmbTEMCOtherMotor, Mid$(sFull, I, 1)
3705               Case 12
3706                       ParseTEMCModelNo cmbTEMCTRG, Mid$(sFull, I, 1)
3707               Case 13
3708                       ParseTEMCModelNo cmbTEMCNominalDischargeSize, Mid$(sFull, I, 2)
3709               Case 14
3710               Case 15
3711                       ParseTEMCModelNo cmbTEMCNominalSuctionSize, Mid$(sFull, I, 2)
3712               Case 16
3713               Case 17
3714                       ParseTEMCModelNo cmbTEMCNominalImpSize, Mid$(sFull, I, 1)
3715               Case 18
3716                       ParseTEMCModelNo cmbTEMCImpellerType, Mid$(sFull, I, 1)
3717               Case 19
3718                       ParseTEMCModelNo cmbTEMCDivisionType, Mid$(sFull, I, 1)
3719               Case 20
3720                       ParseTEMCModelNo cmbTEMCPumpStages, Mid$(sFull, I, 1)
3721               Case 21
3722                       ParseTEMCModelNo cmbTEMCJacketGasket, Mid$(sFull, I, 1)
3723               Case 22
3724                       ParseTEMCModelNo cmbTEMCAdditions, Mid$(sFull, I, 1)
3725                       ParseTEMCModelNo cmbTEMCCirculation, "*"
3726               Case 23
       '                    ParseTEMCModelNo cmbTEMCCirculation, Mid$(sFull, I, 1)

3727           End Select
3728       Next I

           'give alerts on certain conditions
3729       Dim msg As String
3730       msg = ""
3731       If Left(cmbTEMCVoltage, 3) = "[6]" Then
3732           msg = "Requires Transformer"
3733       End If
3734       If Left(cmbTEMCTRG, 3) = "[L]" Or InStr("X[B][F][H][K]", Left(cmbTEMCAdditions, 3)) > 1 Then
3735           If msg = "" Then
3736               msg = "Requires VFD"
3737           Else
3738               msg = msg & " and " & "Requires VFD"
3739           End If
3740       End If

3741       If msg <> "" Then
3742           frmAlert.txtAlert.Text = msg
3743           frmAlert.Show
3744       End If

' <VB WATCH>
3745       Exit Sub
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
    End Select
' </VB WATCH>
End Sub


Private Sub txtModelNo_Validate(Cancel As Boolean)
' <VB WATCH>
3746       On Error GoTo vbwErrHandler
' </VB WATCH>
3747       Dim I As Integer
3748       Dim S As String

       '    s = txtModelNo.Text
       '    S = Replace(S, "-", "")
       '    S = Replace(S, " ", "")
       '    S = Replace(S, "/", "")

       '    txtModelNo.Text = ""

       '    For i = 1 To Len(s)
       '        txtModelNo.Text = txtModelNo.Text & Mid(s, i, 1)
       '    Next i
3749       txtModelNo_Change

' <VB WATCH>
3750       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub txtNPSHFile_GotFocus()
' <VB WATCH>
3751       On Error GoTo vbwErrHandler
' </VB WATCH>
3752       On Error GoTo FileCancel
3753       If LenB(txtNPSHFile.Text) <> 0 Then
3754           CommonDialog1.filename = txtNPSHFile.Text
3755       End If
3756       CommonDialog1.ShowOpen
3757       txtNPSHFile.Text = CommonDialog1.filename
3758       Exit Sub
3759   FileCancel:
3760   On Error GoTo vbwErrHandler
3761       CommonDialog1.CancelError = False
' <VB WATCH>
3762       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub txtP1_Change()
' <VB WATCH>
3763       On Error GoTo vbwErrHandler
' </VB WATCH>
3764       txtP2.Text = txtP1.Text
3765       txtP3.Text = txtP1.Text
' <VB WATCH>
3766       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub txtPicturesFile_gotfocus()
' <VB WATCH>
3767       On Error GoTo vbwErrHandler
' </VB WATCH>
3768       CommonDialog1.CancelError = True
3769       On Error GoTo FileCancel
3770       If LenB(txtPicturesFile.Text) <> 0 Then
3771           CommonDialog1.filename = txtPicturesFile.Text
3772       End If
3773       CommonDialog1.ShowOpen
3774       txtPicturesFile.Text = CommonDialog1.filename
3775       Exit Sub
3776   FileCancel:
3777   On Error GoTo vbwErrHandler
3778       CommonDialog1.CancelError = False
' <VB WATCH>
3779       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub txtSN_Change()
' <VB WATCH>
3780       On Error GoTo vbwErrHandler
' </VB WATCH>
3781       cmdFindPump.Default = True
' <VB WATCH>
3782       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub txtTEMCFrontThrust_Change()
' <VB WATCH>
3783       On Error GoTo vbwErrHandler
' </VB WATCH>
3784       CalculateTEMCForce
' <VB WATCH>
3785       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub txtTEMCMomentArm_Change()
' <VB WATCH>
3786       On Error GoTo vbwErrHandler
' </VB WATCH>
3787       CalculateTEMCForce
' <VB WATCH>
3788       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub txtTEMCRearThrust_Change()
' <VB WATCH>
3789       On Error GoTo vbwErrHandler
' </VB WATCH>
3790       CalculateTEMCForce
' <VB WATCH>
3791       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub txtTEMCThrustRigPressure_Change()
' <VB WATCH>
3792       On Error GoTo vbwErrHandler
' </VB WATCH>
3793       CalculateTEMCForce
' <VB WATCH>
3794       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub txtTEMCViscosity_Change()
' <VB WATCH>
3795       On Error GoTo vbwErrHandler
' </VB WATCH>
3796       CalculateTEMCForce
' <VB WATCH>
3797       Exit Sub
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
    End Select
' </VB WATCH>
End Sub



Private Sub txtV1_Change()
' <VB WATCH>
3798       On Error GoTo vbwErrHandler
' </VB WATCH>
3799       txtV2.Text = txtV1.Text
3800       txtV3.Text = txtV1.Text
' <VB WATCH>
3801       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Private Sub txtVibrationFile_gotfocus()
' <VB WATCH>
3802       On Error GoTo vbwErrHandler
' </VB WATCH>
3803       On Error GoTo FileCancel
3804       If LenB(txtVibrationFile.Text) <> 0 Then
3805           CommonDialog1.filename = txtVibrationFile.Text
3806       End If
3807       CommonDialog1.ShowOpen
3808       txtVibrationFile.Text = CommonDialog1.filename
3809       Exit Sub
3810   FileCancel:
3811   On Error GoTo vbwErrHandler
3812       CommonDialog1.CancelError = False
' <VB WATCH>
3813       Exit Sub
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
    End Select
' </VB WATCH>
End Sub
Private Sub ExportToExcel()
' <VB WATCH>
3814       On Error GoTo vbwErrHandler
' </VB WATCH>

3815       Dim SaveFileName As String
3816       Dim WorkSheetName As String

3817       Dim I As Integer
3818       Dim iRowNo As Integer
3819       Dim sImp As String
3820       Dim ans As Integer

3821       Dim bCanShowSpeed As Boolean
3822       Dim CantShowReason As String

       'close any running excel processes
3823       Dim objWMIService, colProcesses
3824       Set objWMIService = GetObject("winmgmts:")
3825       Set colProcesses = objWMIService.ExecQuery("Select * from Win32_Process where name LIKE 'Excel%'")
3826       If colProcesses.Count > 0 Then
3827           Set xlApp = Excel.Application
3828       Else
               'use existing copy
       '        Set xlApp = New Excel.Application
3829           Set xlApp = CreateObject("Excel.Application")
3830       End If


3831       CommonDialog1.CancelError = True        'in case the user
3832       On Error GoTo ErrHandler                '  chooses the cancel button

           'set up dialog box
3833       CommonDialog1.DialogTitle = "Open Excel Files"
3834       CommonDialog1.Filter = "Excel Files (*.xls)|*.xls|"  'show Excel files
3835       CommonDialog1.InitDir = App.Path
       '    CommonDialog1.InitDir = "C:\"    'in this directory
3836       CommonDialog1.ShowOpen                              'open the file selection dialog box

3837       If Dir(CommonDialog1.filename) = "" Then            'if the file name does not exist yet
3838           SaveFileName = CommonDialog1.filename           'get the name of the file
3839           If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
3840                xlApp.Workbooks.Close
3841           End If
               ' Create the Excel Workbook Object.
3842   On Error GoTo vbwErrHandler
3843           Set xlBook = xlApp.Workbooks.Add                'add a workbook
3844           WorkSheetName = NewWorkBook                                     'do some stuff for the new workbook
3845           ActiveWorkbook.CheckCompatibility = False
3846           xlApp.ActiveWorkbook.SaveAs filename:=SaveFileName, _
                                 FileFormat:=xlNormal                        'save the file
3847       Else                                                'the file name already exists
3848           SaveFileName = CommonDialog1.filename
               ' Create the Excel Workbook Object.
3849           If Not IsNull(xlApp.Workbooks) Then 'if there's a workbook open, close it
3850                xlApp.Workbooks.Close
3851           End If
3852           Set xlBook = xlApp.Workbooks.Open(SaveFileName)             'get the file name selected
3853           If GetWorksheetTabs(SaveFileName, WorkSheetName) = vbNo Then    'ask the user if he/she wants a new tab.
3854               MsgBox "File not overwritten.", vbOKOnly, "File not Opened"
3855               Exit Sub
3856           Else
3857           End If
3858       End If

3859   On Error GoTo vbwErrHandler

           'see if we can export Speed and SG and if we can, ask user if s/he wants it
           'assume that we can show speed calcs

3860       bCanShowSpeed = False
       'open the template and copy the data from the sheet
       '  excel file resides in ParentDirectoryName + "\Polar SG&Visc Correction5.xls"
           'write the data to the spreadsheet
3861       With xlApp

3862       Dim xlTemplateName As String
3863       xlTemplateName = ParentDirectoryName & sSGandViscSpreadsheetTemplate
3864       Dim xlTemplate As Excel.Workbook
3865       Set xlTemplate = xlApp.Workbooks.Open(xlTemplateName)
3866       Dim TemplateWS As Excel.Worksheet
3867       Dim sheetName As String
3868       sheetName = xlTemplate.Sheets(1).Name
3869       xlTemplate.Sheets(1).Copy After:=xlBook.Sheets(WorkSheetName)

3870       xlTemplate.Close savechanges:=False

3871       Set xlTemplate = Nothing

3872       Application.DisplayAlerts = False
3873       ActiveWorkbook.Worksheets(WorkSheetName).Delete
3874       Application.DisplayAlerts = True
3875       ActiveWorkbook.Worksheets(sheetName).Name = WorkSheetName

           'WorkSheetName = sheetName

           'first see if there is an entry in CalculatedRPM table for this frame size and voltage.
           ' if there is, get the coefficients, else make the coefficients 0

3876           Dim ACoef As Double
3877           Dim BCoef As Double
3878           Dim CCoef As Double

3879           Dim qy As New ADODB.Command
3880           Dim rs As New ADODB.Recordset
3881           qy.ActiveConnection = cnPumpData
3882           Dim VoltageForLookup As Integer
3883           If cmbVoltage.List(cmbVoltage.ListIndex) = "380" And cmbFrequency.List(cmbFrequency.ListIndex) = "50 Hz" Then
3884               VoltageForLookup = 460
3885           ElseIf cmbVoltage.List(cmbVoltage.ListIndex) <> "380" Then
3886               VoltageForLookup = cmbVoltage.List(cmbVoltage.ListIndex)
3887           End If
3888           qy.CommandText = "SELECT * FROM CalculatedRPM WHERE FrameNumber = '" & txtTEMCFrameNumber.Text & _
                          "' AND Voltage = '" & VoltageForLookup & "'"

3889           rs.CursorLocation = adUseClient
3890           rs.CursorType = adOpenStatic

3891           rs.Open qy
3892           If rs.RecordCount = 0 Then
3893               ACoef = 0
3894               BCoef = 0
3895               CCoef = 0
3896               MsgBox ("Cannot find coefficient data for Frame Number " & txtTEMCFrameNumber.Text & _
                          " AND Voltage = " & cmbVoltage.List(cmbVoltage.ListIndex) & _
                          " AND Frequency = " & cmbFrequency.List(cmbFrequency.ListIndex))
3897           Else
3898               ACoef = rs.Fields("A")
3899               BCoef = rs.Fields("B")
3900               CCoef = rs.Fields("C")
3901           End If


           'write header data

3902           .Range("A2").Select
3903           .ActiveCell.FormulaR1C1 = "Serial Number"
3904           .Range("C2").Select
3905           .ActiveCell.FormulaR1C1 = txtSN

3906           .Range("F1").Select
3907           .ActiveCell.FormulaR1C1 = "Customer"
3908           .Range("H1").Select
3909           .ActiveCell.FormulaR1C1 = txtShpNo

3910           .Range("A3").Select
3911           .ActiveCell.FormulaR1C1 = "Model"
3912           .Range("C3").Select
3913           .ActiveCell.FormulaR1C1 = txtModelNo

3914           .Range("F2").Select
3915           .ActiveCell.FormulaR1C1 = "Sales Order"
3916           .Range("H2").Select
3917           .ActiveCell.FormulaR1C1 = txtSalesOrderNumber

3918           .Range("A9").Select
3919           .ActiveCell.FormulaR1C1 = "Design Flow"
3920           .Range("C9").Select
3921           .ActiveCell.FormulaR1C1 = Val(txtDesignFlow)

3922           .Range("A10").Select
3923           .ActiveCell.FormulaR1C1 = "Design Head"
3924           .Range("C10").Select
3925           .ActiveCell.FormulaR1C1 = Val(txtDesignTDH)

3926           .Range("P13").Select
3927           .ActiveCell.FormulaR1C1 = "Barometric Pressure"
3928           .Range("R13").Select
3929           .ActiveCell.FormulaR1C1 = Val(txtInHgDisplay)

3930           .Range("P11").Select
3931           .ActiveCell.FormulaR1C1 = "Suction Gage Height"
3932           .Range("R11").Select
3933           .ActiveCell.FormulaR1C1 = Val(txtSuctHeight)

3934           .Range("P12").Select
3935           .ActiveCell.FormulaR1C1 = "Discharge Gage Height"
3936           .Range("R12").Select
3937           .ActiveCell.FormulaR1C1 = Val(txtDischHeight)

3938           .Range("A1").Select
3939           .ActiveCell.FormulaR1C1 = "Run Date"
3940           .Range("C1").Select
3941           .ActiveCell.FormulaR1C1 = cmbTestDate.List(cmbTestDate.ListIndex)

3942           .Range("D10:E10").Select
3943           With xlApp.Selection
3944               .HorizontalAlignment = xlCenter
3945               .VerticalAlignment = xlBottom
3946               .WrapText = False
3947               .Orientation = 0
3948               .AddIndent = False
3949               .IndentLevel = 0
3950               .ShrinkToFit = False
3951               .ReadingOrder = xlContext
3952               .MergeCells = False
3953           End With
3954           xlApp.Selection.Merge

               'determine rpm

3955           Dim RPMvalue As String
3956           If Mid$(Me.txtTEMCFrameNumber.Text, 2, 1) = "1" Then
               '1 says 2 pole
3957               If Me.cmbFrequency.ListIndex = 0 Then
                       '0 says 50Hz
3958                   RPMvalue = "2900"
3959               ElseIf Me.cmbFrequency.ListIndex = 1 Then
                       ' says 60Hz
3960                   RPMvalue = "3450"
3961               Else
                       'vfd or other, no rpm
3962                   RPMvalue = ""
3963               End If
3964           Else
               '2 says 4 pole
3965               If Me.cmbFrequency.ListIndex = 0 Then
                       '0 says 50Hz
3966                   RPMvalue = "1450"
3967               ElseIf Me.cmbFrequency.ListIndex = 1 Then
                       ' says 60Hz
3968                   RPMvalue = "1750"
3969               Else
                       'vfd or other, no rpm
3970                   RPMvalue = ""
3971               End If
3972           End If

       '        .Range("G1").Select
       '        .ActiveCell.FormulaR1C1 = "RPM"
       '        .Range("I1").Select
       '        .ActiveCell.FormulaR1C1 = RPMvalue

3973           .Range("A5").Select
3974           .ActiveCell.FormulaR1C1 = "Sp Gravity"
3975           .Range("C5").Select
3976           .ActiveCell.FormulaR1C1 = txtSpGr

3977           .Range("A6").Select
3978           .ActiveCell.FormulaR1C1 = "Viscosity"
3979           .Range("C6").Select
3980           .ActiveCell.FormulaR1C1 = txtViscosity

3981           .Range("F4").Select
3982           .ActiveCell.FormulaR1C1 = "Motor"
3983           .Range("H4").Select
3984           .ActiveCell.FormulaR1C1 = txtTEMCFrameNumber.Text

3985           .Range("H12").Select
3986           .ActiveCell.FormulaR1C1 = Me.txtCustPONum.Text

3987           .Range("F5").Select
3988           .ActiveCell.FormulaR1C1 = "Voltage"
3989           .Range("H5").Select
3990           .ActiveCell.FormulaR1C1 = cmbVoltage.List(cmbVoltage.ListIndex)

3991           .Range("K6").Select
3992           .ActiveCell.FormulaR1C1 = "End Play"
3993           .Range("M6").Select
3994           .ActiveCell.FormulaR1C1 = Val(txtEndPlay)

3995           .Range("K7").Select
3996           .ActiveCell.FormulaR1C1 = "G-Gap"
3997           .Range("M7").Select
3998           .ActiveCell.FormulaR1C1 = txtGGap.Text

3999           .Range("A8").Select
4000           .ActiveCell.FormulaR1C1 = "Design Pressure"
4001           .Range("C8").Select
4002           Dim DesPress As String
4003           DesPress = cmbTEMCDesignPressure.List(cmbTEMCDesignPressure.ListIndex)
4004           Dim j As Integer
4005           j = InStrRev(DesPress, "-")
4006           .ActiveCell.FormulaR1C1 = Mid$(DesPress, j + 2)

       '        .Range("G8").Select
       '        .ActiveCell.FormulaR1C1 = "Stator Fill"
       '        .Range("I8").Select
       '        .ActiveCell.FormulaR1C1 = "Dry"

4007           .Range("K4").Select
4008           .ActiveCell.FormulaR1C1 = "Circulation Path"
4009           .Range("M4").Select
4010           .ActiveCell.FormulaR1C1 = cmbTEMCModel.List(cmbTEMCModel.ListIndex)

4011           .Range("M8").Select
4012           .ActiveCell.FormulaR1C1 = txtNPSHr.Text

4013           .Range("K1").Select
4014           .ActiveCell.FormulaR1C1 = "Impeller Dia"
4015           .Range("M1").Select


       '        If LenB(txtImpTrim) <> 0 Then
       '            .ActiveCell.FormulaR1C1 = Val(txtImpTrim)
       '        Else
       '            .ActiveCell.FormulaR1C1 = Val(txtImpellerDia)
       '        End If
       '
4016           If chkTrimmed.value = 1 Then
4017               If Val(txtImpTrim.Text) <> 0 Then
4018                   .ActiveCell.FormulaR1C1 = txtImpTrim
4019               Else
4020                   .ActiveCell.FormulaR1C1 = txtImpellerDia
4021               End If
4022           Else
4023               .ActiveCell.FormulaR1C1 = txtImpellerDia
4024           End If



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

4025           .Range("P9").Select
4026           .ActiveCell.FormulaR1C1 = "Suction Dia"
4027           .Range("R9").Select
4028           .ActiveCell.FormulaR1C1 = cmbSuctDia.List(cmbSuctDia.ListIndex)

4029           .Range("P10").Select
4030           .ActiveCell.FormulaR1C1 = "Discharge Dia"
4031           .Range("R10").Select
4032           .ActiveCell.FormulaR1C1 = cmbDischDia.List(cmbDischDia.ListIndex)

4033           .Range("A11").Select
4034           .ActiveCell.FormulaR1C1 = "Test Spec"
4035           .Range("C11").Select
4036           .ActiveCell.FormulaR1C1 = cmbTestSpec.List(cmbTestSpec.ListIndex)

4037           .Range("K3").Select
4038           .ActiveCell.FormulaR1C1 = "Impeller Feathered"
4039           .Range("M3").Select
4040           If chkFeathered.value = 1 Then
4041               .ActiveCell.FormulaR1C1 = "Yes"
4042           Else
4043               .ActiveCell.FormulaR1C1 = "No"
4044           End If

4045           .Range("K2").Select
4046           .ActiveCell.FormulaR1C1 = "Disch Orifice"
4047           .Range("M2").Select
4048           If chkOrifice.value = 1 Then
4049               .ActiveCell.FormulaR1C1 = Val(txtOrifice)
4050           Else
4051               .ActiveCell.FormulaR1C1 = "None"
4052           End If


4053           .Range("K5").Select
4054           .ActiveCell.FormulaR1C1 = "Circulation Orifice"
4055           .Range("M5").Select
4056           If chkCircOrifice.value = 1 Then
4057               .ActiveCell.FormulaR1C1 = Val(txtCircOrifice)
4058           Else
4059               .ActiveCell.FormulaR1C1 = "None"
4060           End If

4061           .Range("A13").Select
4062           .ActiveCell.FormulaR1C1 = "Other Mods"
4063           .Range("C13").Select
4064           .ActiveCell.FormulaR1C1 = txtOtherMods

4065           .Range("A14").Select
4066           .ActiveCell.FormulaR1C1 = "Remarks"
4067           .Range("C14").Select
4068           .ActiveCell.FormulaR1C1 = txtRemarks

4069           .Range("A15").Select
4070           .ActiveCell.FormulaR1C1 = "Test Setup Remarks"
4071           .Range("C15").Select
4072           .ActiveCell.FormulaR1C1 = txtTestSetupRemarks

4073           .Range("P1").Select
4074           .ActiveCell.FormulaR1C1 = "Suct ID"
4075           .Range("R1").Select
4076           .ActiveCell.FormulaR1C1 = cmbSuctionPressureTransducer.List(cmbSuctionPressureTransducer.ListIndex)

4077           .Range("P2").Select
4078           .ActiveCell.FormulaR1C1 = "Disch ID"
4079           .Range("R2").Select
4080           .ActiveCell.FormulaR1C1 = cmbDischargePressureTransducer.List(cmbDischargePressureTransducer.ListIndex)

4081           .Range("P3").Select
4082           .ActiveCell.FormulaR1C1 = "Temp ID"
4083           .Range("R3").Select
4084           .ActiveCell.FormulaR1C1 = cmbTemperatureTransducer.List(cmbTemperatureTransducer.ListIndex)

4085           .Range("P4").Select
4086           .ActiveCell.FormulaR1C1 = "Circ Flow ID"
4087           .Range("R4").Select
4088           .ActiveCell.FormulaR1C1 = cmbCirculationFlowMeter.List(cmbCirculationFlowMeter.ListIndex)

4089           .Range("P5").Select
4090           .ActiveCell.FormulaR1C1 = "Flow ID"
4091           .Range("R5").Select
4092           .ActiveCell.FormulaR1C1 = cmbFlowMeter.List(cmbFlowMeter.ListIndex)

4093           .Range("P6").Select
4094           .ActiveCell.FormulaR1C1 = "Analyzer ID"
4095           .Range("R6").Select
4096           .ActiveCell.FormulaR1C1 = cmbAnalyzerNo.List(cmbAnalyzerNo.ListIndex)

4097           .Range("P7").Select
4098           .ActiveCell.FormulaR1C1 = "Loop ID"
4099           .Range("R7").Select
4100           .ActiveCell.FormulaR1C1 = cmbLoopNumber.List(cmbLoopNumber.ListIndex)

4101           .Range("A4").Select
4102           .ActiveCell.FormulaR1C1 = "Fluid"
4103           .Range("C4").Select
4104           .ActiveCell.FormulaR1C1 = txtLiquid.Text

4105           .Range("F3").Select
4106           .ActiveCell.FormulaR1C1 = "Cust PN"
4107           .Range("H3").Select
       '        .ActiveCell.FormulaR1C1 = txtRMA.Text
4108           If rsPumpData.Fields("RVSPartNo") <> "" Then
4109               .ActiveCell.FormulaR1C1 = rsPumpData.Fields("RVSPartNo")
4110           End If
4111           If rsPumpData.Fields("CustPN") <> "" Then
4112               .ActiveCell.FormulaR1C1 = rsPumpData.Fields("CustPN")
4113           End If

4114           .Range("A7").Select
4115           .ActiveCell.FormulaR1C1 = "Temperature"
4116           .Range("C7").Select
4117           .ActiveCell.FormulaR1C1 = txtLiquidTemperature.Text

4118           .Range("F6").Select
4119           .ActiveCell.FormulaR1C1 = "Frequency"
4120           .Range("H6").Select
4121           If UCase(cmbFrequency.List(cmbFrequency.ListIndex)) = "VFD" Then
4122               .ActiveCell.FormulaR1C1 = Val(Me.txtVFDFreq)
4123           Else
4124               .ActiveCell.FormulaR1C1 = Val(cmbFrequency.List(cmbFrequency.ListIndex))
4125           End If
       '        .Range("K2").Select
       '        .ActiveCell.FormulaR1C1 = "Disch Orifice"
       '        .Range("M2").Select
       '        .ActiveCell.FormulaR1C1 = txtOrifice.Text

       '        .Range("K12").Select
       '        .ActiveCell.FormulaR1C1 = "Flow Orifice"
       '        .Range("L12").Select
       '        .ActiveCell.FormulaR1C1 = txtCircOrifice.Text

4126           .Range("P8").Select
4127           .ActiveCell.FormulaR1C1 = "PLC No"
4128           .Range("R8").Select
4129           .ActiveCell.FormulaR1C1 = cmbPLCNo.List(cmbPLCNo.ListIndex)

4130           .Range("F7").Select
4131           .ActiveCell.FormulaR1C1 = "Phases"
4132           .Range("H7").Select
4133           .ActiveCell.FormulaR1C1 = txtNoPhases.Text

4134           .Range("F8").Select
4135           .ActiveCell.FormulaR1C1 = "Poles"
4136           .Range("H8").Select
4137           .ActiveCell.FormulaR1C1 = 2 * Val(Left$(Right$(txtTEMCFrameNumber.Text, 2), 1))

4138           .Range("F9").Select
4139           .ActiveCell.FormulaR1C1 = "Rated Current"
4140           .Range("H9").Select
4141           .ActiveCell.FormulaR1C1 = txtAmps.Text

4142           .Range("F10").Select
4143           .ActiveCell.FormulaR1C1 = "Rated Input Power"
4144           .Range("H10").Select
4145           .ActiveCell.FormulaR1C1 = txtRatedInputPower.Text

4146           .Range("F11").Select
4147           .ActiveCell.FormulaR1C1 = "Insulation Class"
4148           .Range("H11").Select
4149           .ActiveCell.FormulaR1C1 = txtThermalClass.Text

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

4150           .Range("A17").Select
4151           .ActiveCell.FormulaR1C1 = "Flow"
4152           .Range("A18").Select
4153           .ActiveCell.FormulaR1C1 = "(GPM)"

4154           .Range("B17").Select
4155           .ActiveCell.FormulaR1C1 = "TDH"
4156           .Range("B18").Select
4157           .ActiveCell.FormulaR1C1 = "(Ft)"

4158           .Range("C17").Select
4159           .ActiveCell.FormulaR1C1 = "KW"

4160           .Range("D17").Select
4161           .ActiveCell.FormulaR1C1 = "Ave"
4162           .Range("D18").Select
4163           .ActiveCell.FormulaR1C1 = "Volts"

4164           .Range("E17").Select
4165           .ActiveCell.FormulaR1C1 = "Ave"
4166           .Range("E18").Select
4167           .ActiveCell.FormulaR1C1 = "Amps"

4168           .Range("F17").Select
4169           .ActiveCell.FormulaR1C1 = "Power"
4170           .Range("F18").Select
4171           .ActiveCell.FormulaR1C1 = "Factor"

4172           .Range("G17").Select
4173           .ActiveCell.FormulaR1C1 = "Overall"
4174           .Range("G18").Select
4175           .ActiveCell.FormulaR1C1 = "Eff"

4176           .Range("H17").Select
4177           .ActiveCell.FormulaR1C1 = "Measured"
4178           .Range("H18").Select
4179           .ActiveCell.FormulaR1C1 = "RPM"

4180           .Range("I17").Select
4181           .ActiveCell.FormulaR1C1 = "Calculated"
4182           .Range("I18").Select
4183           .ActiveCell.FormulaR1C1 = "RPM"

4184           .Range("J17").Select
4185           .ActiveCell.FormulaR1C1 = "Suction"
4186           .Range("J18").Select
4187           .ActiveCell.FormulaR1C1 = "Temp(F)"

4188           .Range("K17").Select
4189           .ActiveCell.FormulaR1C1 = "Disch"
4190           .Range("K18").Select
4191           .ActiveCell.FormulaR1C1 = "Pressure"

4192           .Range("L17").Select
4193           .ActiveCell.FormulaR1C1 = "Suction"
4194           .Range("L18").Select
4195           .ActiveCell.FormulaR1C1 = "Pressure"

4196           .Range("M17").Select
4197           .ActiveCell.FormulaR1C1 = "Vel"
4198           .Range("M18").Select
4199           .ActiveCell.FormulaR1C1 = "Head"

4200           .Range("N17").Select
4201           .ActiveCell.FormulaR1C1 = "Axial"
4202           .Range("N18").Select
4203           .ActiveCell.FormulaR1C1 = "Position"

4204           .Range("O17").Select
4205           .ActiveCell.FormulaR1C1 = "Pct of"
4206           .Range("O18").Select
4207           .ActiveCell.FormulaR1C1 = "End Play"

4208           .Range("P17").Select
4209           .ActiveCell.FormulaR1C1 = "Hydraulic"
4210           .Range("P18").Select
4211           .ActiveCell.FormulaR1C1 = "Efficiency"

       '        .Range("P17").Select
       '        .ActiveCell.FormulaR1C1 = "Circ"
       '        .Range("P18").Select
       '        .ActiveCell.FormulaR1C1 = "Flow"

4212           .Range("Q17").Select
4213           .ActiveCell.FormulaR1C1 = "Motor"
4214           .Range("Q18").Select
4215           .ActiveCell.FormulaR1C1 = "Efficiency"

4216           .Range("S17").Select
4217           .ActiveCell.FormulaR1C1 = "NPSHa"

4218           .Range("T17").Select
4219           .ActiveCell.FormulaR1C1 = "Phase 1"
4220           .Range("T18").Select
4221           .ActiveCell.FormulaR1C1 = "Current"

4222           .Range("U17").Select
4223           .ActiveCell.FormulaR1C1 = "Phase 2"
4224           .Range("U18").Select
4225           .ActiveCell.FormulaR1C1 = "Current"

4226           .Range("V17").Select
4227           .ActiveCell.FormulaR1C1 = "Phase 3"
4228           .Range("V18").Select
4229           .ActiveCell.FormulaR1C1 = "Current"

4230           .Range("W17").Select
4231           .ActiveCell.FormulaR1C1 = "Phase 1"
4232           .Range("W18").Select
4233           .ActiveCell.FormulaR1C1 = "Voltage"

4234           .Range("X17").Select
4235           .ActiveCell.FormulaR1C1 = "Phase 2"
4236           .Range("X18").Select
4237           .ActiveCell.FormulaR1C1 = "Voltage"

4238           .Range("Y17").Select
4239           .ActiveCell.FormulaR1C1 = "Phase 3"
4240           .Range("Y18").Select
4241           .ActiveCell.FormulaR1C1 = "Voltage"

4242           .Range("Z17").Select
4243           .ActiveCell.FormulaR1C1 = "'" & txtTitle(20).Text

4244           .Range("Z18").Select
4245           .ActiveCell.FormulaR1C1 = "'" & txtTitle(21).Text

4246           .Range("AA17").Select
4247           .ActiveCell.FormulaR1C1 = "'" & txtTitle(22).Text

4248           .Range("AA18").Select
4249           .ActiveCell.FormulaR1C1 = "'" & txtTitle(23).Text

4250           .Range("AB17").Select
4251           .ActiveCell.FormulaR1C1 = "'" & txtTitle(24).Text

4252           .Range("AB18").Select
4253           .ActiveCell.FormulaR1C1 = "'" & txtTitle(25).Text

4254           .Range("AC17").Select
4255           .ActiveCell.FormulaR1C1 = "HR"

4256           .Range("AC18").Select
4257           .ActiveCell.FormulaR1C1 = "(ft)"

4258           .Range("AD17").Select
4259           .ActiveCell.FormulaR1C1 = "'" & txtTitle(26).Text

4260           .Range("AD18").Select
4261           .ActiveCell.FormulaR1C1 = "'" & txtTitle(27).Text

4262           .Range("AE17").Select
4263           .ActiveCell.FormulaR1C1 = "TRG"
4264           .Range("AE18").Select
4265           .ActiveCell.FormulaR1C1 = "Position"

4266           .Range("AF17").Select
4267           .ActiveCell.FormulaR1C1 = "Thrust"

4268           .Range("AG17").Select
4269           .ActiveCell.FormulaR1C1 = "F/R"

4270           .Range("AH17").Select
4271           .ActiveCell.FormulaR1C1 = "Moment"
4272           .Range("AH18").Select
4273           .ActiveCell.FormulaR1C1 = "Arm"

4274           .Range("AI17").Select
4275           .ActiveCell.FormulaR1C1 = "Rig"
4276           .Range("AI18").Select
4277           .ActiveCell.FormulaR1C1 = "Pressure"

       '        .Range("AI17").Select
       '        .ActiveCell.FormulaR1C1 = "Viscosity"

4278           .Range("AJ19").Select
4279           .ActiveCell.FormulaR1C1 = "Rear"
4280           .Range("AJ18").Select
4281           .ActiveCell.FormulaR1C1 = "Force"

4282           .Range("AK17").Select
4283           .ActiveCell.FormulaR1C1 = "PV"

4284           .Range("R17").Select
4285           .ActiveCell.FormulaR1C1 = "Shaft"
4286           .Range("R18").Select
4287           .ActiveCell.FormulaR1C1 = "Power"

       '        .Range("AM17").Select
       '        .ActiveCell.FormulaR1C1 = "Pct Full"
       '        .Range("AM18").Select
       '        .ActiveCell.FormulaR1C1 = "Scale"

4288           .Range("AL17").Select
4289           .ActiveCell.FormulaR1C1 = "NPSHr"

4290           .Range("AM17").Select
4291           .ActiveCell.FormulaR1C1 = "Remarks"




               'now output the data

4292           iRowNo = 20

4293           rsEff.MoveFirst
4294           For I = 1 To frmPLCData.UpDown2.value
4295               .Range("A" & iRowNo).Select
4296               .ActiveCell.FormulaR1C1 = rsEff.Fields("Flow")

4297               .Range("B" & iRowNo).Select
4298               .ActiveCell.FormulaR1C1 = rsEff.Fields("TDH")

4299               .Range("C" & iRowNo).Select
4300               .ActiveCell.FormulaR1C1 = rsEff.Fields("KW")

4301               .Range("D" & iRowNo).Select
4302               .ActiveCell.FormulaR1C1 = rsEff.Fields("Volts")

4303               .Range("E" & iRowNo).Select
4304               .ActiveCell.FormulaR1C1 = rsEff.Fields("Amps")

4305               .Range("F" & iRowNo).Select
4306               .ActiveCell.FormulaR1C1 = rsEff.Fields("PowerFactor")

4307               .Range("G" & iRowNo).Select
4308               .ActiveCell.FormulaR1C1 = rsEff.Fields("OverallEfficiency")

4309               .Range("H" & iRowNo).Select
4310               .ActiveCell.FormulaR1C1 = rsEff.Fields("RPM")

4311               .Range("I" & iRowNo).Select
                   'use the coefficients from above to calculate rpm
4312               Dim f As Double
4313               f = .Range("H6").value
4314               .ActiveCell.FormulaR1C1 = (Val(f) / 60) * (ACoef * (rsEff.Fields("KW")) ^ 2 + BCoef * (rsEff.Fields("KW")) + CCoef)

4315               .Range("J" & iRowNo).Select
4316               .ActiveCell.FormulaR1C1 = rsEff.Fields("Temperature")

4317               .Range("K" & iRowNo).Select
4318               .ActiveCell.FormulaR1C1 = rsEff.Fields("DischPress")

4319               .Range("L" & iRowNo).Select
4320               .ActiveCell.FormulaR1C1 = rsEff.Fields("SuctPress")

4321               .Range("M" & iRowNo).Select
4322               .ActiveCell.FormulaR1C1 = rsEff.Fields("VelocityHead")

4323               .Range("N" & iRowNo).Select
4324               .ActiveCell.FormulaR1C1 = rsEff.Fields("Pos")

4325               .Range("O" & iRowNo).Select
4326               .ActiveCell.FormulaR1C1 = 100 * rsEff.Fields("Pos") / Val(txtEndPlay)

4327               .Range("P" & iRowNo).Select
4328               .ActiveCell.FormulaR1C1 = rsEff.Fields("HydraulicEfficiency")

       '            .Range("P" & iRowNo).Select
       '            .ActiveCell.FormulaR1C1 = rsEff.Fields("CircFlow")

4329               .Range("Q" & iRowNo).Select
4330               .ActiveCell.FormulaR1C1 = rsEff.Fields("MotorEfficiency")

4331               .Range("S" & iRowNo).Select
4332               .ActiveCell.FormulaR1C1 = rsEff.Fields("NPSHa")

4333               .Range("T" & iRowNo).Select
4334               .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentA")

4335               .Range("U" & iRowNo).Select
4336               .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentB")

4337               .Range("V" & iRowNo).Select
4338               .ActiveCell.FormulaR1C1 = rsEff.Fields("CurrentC")

4339               .Range("W" & iRowNo).Select
4340               .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageA")

4341               .Range("X" & iRowNo).Select
4342               .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageB")

4343               .Range("Y" & iRowNo).Select
4344               .ActiveCell.FormulaR1C1 = rsEff.Fields("VoltageC")

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

4345               .Range("Z" & iRowNo).Select
4346               .ActiveCell.FormulaR1C1 = rsEff.Fields("CircFlow")

4347               .Range("AA" & iRowNo).Select
4348               .ActiveCell.FormulaR1C1 = rsEff.Fields("RBHTemp")

4349               .Range("AB" & iRowNo).Select
4350               .ActiveCell.FormulaR1C1 = rsEff.Fields("RBHPress")

4351               .Range("AC" & iRowNo).Select
4352               .ActiveCell.FormulaR1C1 = (rsEff.Fields("RBHPress") - rsEff.Fields("SuctPress")) * 2.31

4353               .Range("AD" & iRowNo).Select
4354               .ActiveCell.FormulaR1C1 = rsEff.Fields("AI4")

4355               .Range("AE" & iRowNo).Select
4356               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCTRG")

4357               .Range("AF" & iRowNo).Select
4358               If rsEff.Fields("TEMCFrontThrust") = 0 Then
4359                   If rsEff.Fields("TEMCRearThrust") = 0 Then
4360                       .ActiveCell.FormulaR1C1 = " "
4361                       .Range("AG" & iRowNo).Select
4362                       .ActiveCell.FormulaR1C1 = " "
4363                   Else
4364                       .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCRearThrust")
4365                       .Range("AG" & iRowNo).Select
4366                       .ActiveCell.FormulaR1C1 = "R"
4367                   End If
4368               Else
4369                   .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCFrontThrust")
4370                   .Range("AG" & iRowNo).Select
4371                   .ActiveCell.FormulaR1C1 = "F"
4372               End If

4373               .Range("AH" & iRowNo).Select
4374               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCMomentArm")

4375               .Range("AI" & iRowNo).Select
4376               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCThrustRigPressure")

       '            .Range("AJ" & iRowNo).Select
       '            .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCViscosity")

4377               .Range("AJ" & iRowNo).Select
4378               If rsEff.Fields("TEMCForceDirection") = "F" Then
4379                   .ActiveCell.FormulaR1C1 = -rsEff.Fields("TEMCCalculatedForce")
4380               Else
4381                   .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCCalculatedForce")
4382               End If

4383               .Range("AK" & iRowNo).Select
4384               .ActiveCell.FormulaR1C1 = rsEff.Fields("TEMCPV")

4385               .Range("R" & iRowNo).Select
4386               .ActiveCell.FormulaR1C1 = rsEff.Fields("KW") * rsEff.Fields("MotorEfficiency") / 100

4387               .Range("AL" & iRowNo).Select
4388               .ActiveCell.FormulaR1C1 = rsEff.Fields("NPSHr")

       '            If RatedKW = 999 Then
       '                .ActiveCell.FormulaR1C1 = ""
       '            Else
       '                .ActiveCell.FormulaR1C1 = (rsEff.Fields("KW") * rsEff.Fields("MotorEfficiency")) / (1 * RatedKW)
       '            End If

4389               .Range("AM" & iRowNo).Select
4390               .ActiveCell.FormulaR1C1 = rsEff.Fields("Remarks")


4391               rsEff.MoveNext
4392               iRowNo = iRowNo + 1
4393           Next I

4394           .Range("A20:AS30").Select
4395           .Selection.NumberFormat = "0.00"

           'set up formulas to calculate BEP
           '  first, plot 2nd order polynomial for flow vs hydraulic efficiency
           '  the formulas for doing that are in E68, F68 and G68
           '  only want the formulas to point to the number of points in the test data, so use frmPLCData.CWNumEdit2.value
           '
4396       Dim AColumnRow As String
4397       Dim PColumnRow As String

4398       AColumnRow = "A" & Trim(str(19 + frmPLCData.UpDown2.value))
4399       PColumnRow = "P" & Trim(str(19 + frmPLCData.UpDown2.value))

4400           .Range("E68").Select
4401           .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1)"

4402           .Range("F68").Select
4403           .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1,2)"

4404           .Range("G68").Select
4405           .ActiveCell.Formula = "=INDEX(LINEST(P20:" & PColumnRow & ",A20:" & AColumnRow & "^{1,2}),1,3)"

           'export balance holes
4406       If boGotBalanceHoles Then
4407           If rsBalanceHoles.State = adStateClosed Then
4408               rsBalanceHoles.ActiveConnection = cnPumpData
4409               rsBalanceHoles.Open
4410           End If 'rsBalanceHoles.State = adStateClosed

4411           If rsBalanceHoles.RecordCount <> 0 Then

4412               .Range("K9:N9").Merge
4413               .Range("K9:N9").Formula = "Balance Hole Data"
4414               .Range("K9:N9").HorizontalAlignment = xlCenter

4415               .Range("K10").Select
4416               .ActiveCell.Formula = "Date"

4417               .Range("L10").Select
4418               .ActiveCell.Formula = "Number"

4419               .Range("M10").Select
4420               .ActiveCell.Formula = "Diameter"

4421               .Range("N10").Select
4422               .ActiveCell.Formula = "Bolt Circle"

4423               iRowNo = 11

4424               If rsBalanceHoles.RecordCount > 3 Then
4425                   For I = 1 To rsBalanceHoles.RecordCount - 3
4426                       Rows("13:13").Select
4427                       Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove
4428                   Next I
4429               End If

4430               rsBalanceHoles.MoveFirst
4431               For I = 1 To rsBalanceHoles.RecordCount

4432                   .Range("K" & iRowNo).Select
4433                   .ActiveCell.Formula = rsBalanceHoles.Fields("Date")
4434                   .ActiveCell.NumberFormat = "m/d/yy h:mm AM/PM;@"
4435                   .Range("L" & iRowNo).Select
4436                   .ActiveCell = rsBalanceHoles.Fields("Number")
4437                   .ActiveCell.NumberFormat = "0"
4438                   .Range("M" & iRowNo).Select
4439                   If IsNumeric(rsBalanceHoles.Fields("Diameter1")) Then
4440                       .ActiveCell = Val(rsBalanceHoles.Fields("Diameter1"))
4441                       .ActiveCell.NumberFormat = "0.0000"
4442                   Else
4443                       .ActiveCell = rsBalanceHoles.Fields("Diameter1")
4444                   End If

4445                   .Range("N" & iRowNo).Select
4446                   If IsNumeric(rsBalanceHoles.Fields("BoltCircle1")) Then
4447                       .ActiveCell = Val(rsBalanceHoles.Fields("BoltCircle1"))
4448                       .ActiveCell.NumberFormat = "0.0000"
4449                   Else
4450                       .ActiveCell = rsBalanceHoles.Fields("BoltCircle1")
4451                   End If

4452                   rsBalanceHoles.MoveNext
4453                   iRowNo = iRowNo + 1
4454               Next I
4455               .Range("K10:N" & iRowNo - 1).Select
4456               With .Selection.Interior
4457                   .ColorIndex = 34
4458                   .Pattern = xlSolid
4459               End With
4460           End If 'rsBalanceHoles.RecordCount <> 0
4461       End If ' boGotBalanceHoles

           'plot graphs

4462       Dim SeriesName As String
4463       Dim XVals As String
4464       Dim YVals As String
4465       Dim RowNo As Long
4466       Dim RowStr As String
4467       Dim LastPoint As Integer
4468       Dim LineType As String
4469       Dim AxisGroup As Integer
4470       Dim LabelPos As Integer
4471       Dim LineColor As Long

4472           .ActiveSheet.ChartObjects("HydRepChart").Activate
4473           Dim S As Series
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
4474           Dim aq As Double
4475           Range("AQ56", "AQ71").Select
4476           aq = .Max(Selection)
4477           Dim ax As Double
4478           Range("AX56", "AX71").Select
4479           ax = .Max(Selection)

               'then current (as and az)
4480           Dim at As Double
4481           Range("AS56", "AS71").Select
4482           at = .Max(Selection)
4483           Dim ba As Double
4484           Range("AZ56", "AZ71").Select
4485           ba = .Max(Selection)

4486           Dim CurrentScaleMax As Integer
4487           Dim TDHScaleMax As Integer

4488           Dim MaxTDH As Integer
4489           With Application.WorksheetFunction
4490               If aq > ax Then
4491                   MaxTDH = .Ceiling(aq, 25)
4492               Else
4493                   MaxTDH = .Ceiling(ax, 25)
4494               End If
4495           End With

4496           Dim MaxCurrent As Integer
4497           With Application.WorksheetFunction
4498               If at > ba Then
4499                   Select Case at
                           Case Is <= 5
4500                           CurrentScaleMax = 5

4501                       Case Is <= 10
4502                           CurrentScaleMax = 10

4503                       Case Else
4504                           CurrentScaleMax = 25
4505                   End Select

4506                   MaxCurrent = .Ceiling(at, CurrentScaleMax)
4507               Else
4508                  Select Case ba
                           Case Is <= 5
4509                           CurrentScaleMax = 5

4510                       Case Is <= 10
4511                           CurrentScaleMax = 10

4512                       Case Else
4513                           CurrentScaleMax = 25
4514                   End Select

4515                   MaxCurrent = .Ceiling(ba, CurrentScaleMax)
4516               End If
4517           End With

4518           ActiveSheet.ChartObjects("HydRepChart").Activate
4519            Dim ShtName As String
4520            ShtName = "'" & ActiveSheet.Name & "'"

4521           RowStr = 56 + 15
4522            For I = 1 To 8

4523                Select Case I
                        Case 1
4524                        SeriesName = "=""TDH"""
4525                        XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
4526                        YVals = "=" & ShtName & "!$AQ$56:$AQ$" & RowStr
4527                        LineType = msoLineSolid
4528                        AxisGroup = 1
4529                        LabelPos = xlLabelPositionRight
4530                        LineColor = vbBlue

4531                    Case 2
4532                        SeriesName = "=""Input Power"""
4533                        XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
4534                        YVals = "=" & ShtName & "!$AR$56:$AR$" & RowStr
4535                        LineType = msoLineSolid
4536                        AxisGroup = 2
4537                        LabelPos = xlLabelPositionRight
4538                        LineColor = vbRed

4539                    Case 3
4540                        SeriesName = "=""Current"""
4541                        XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
4542                        YVals = "=" & ShtName & "!$AS$56:$AS$" & RowStr
4543                        LineType = msoLineSolid
4544                        AxisGroup = 2
4545                        LabelPos = xlLabelPositionRight
4546                        LineColor = vbGreen

4547                    Case 4
       '                     SeriesName = "=""Overall Eff"""
       '                     XVals = "=" & ShtName & "!$AP$56:$AP$" & RowStr
       '                     YVals = "=" & ShtName & "!$AT$56:$AT$" & RowStr
       '                     LineType = msoLineSolid
       '                     AxisGroup = 2
       '                     LabelPos = xlLabelPositionRight
       '                     LineColor = vbCyan

4548                    Case 5
4549                        SeriesName = "=""TDH (Adj)"""
4550                        XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
4551                        YVals = "=" & ShtName & "!$AX$56:$AX$" & RowStr
4552                        LineType = msoLineDash
4553                        AxisGroup = 1
4554                        LabelPos = xlLabelPositionBelow
4555                        LineColor = vbBlue

4556                    Case 6
4557                        SeriesName = "=""Input Power (Adj)"""
4558                        XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
4559                        YVals = "=" & ShtName & "!$AY$56:$AY$" & RowStr
4560                        LineType = msoLineDash
4561                        AxisGroup = 2
4562                        LabelPos = xlLabelPositionBelow
4563                        LineColor = vbRed

4564                    Case 7
4565                        SeriesName = "=""Current (Adj)"""
4566                        XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
4567                        YVals = "=" & ShtName & "!$AZ$56:$AZ$" & RowStr
4568                        LineType = msoLineDash
4569                        AxisGroup = 2
4570                        LabelPos = xlLabelPositionBelow
4571                        LineColor = vbGreen

4572                    Case 8
       '                     SeriesName = "=""Overall Eff (Adj)"""
       '                     XVals = "=" & ShtName & "!$AW$56:$AW$" & RowStr
       '                     YVals = "=" & ShtName & "!$BA$56:$BA$" & RowStr
       '                     LineType = msoLineDash
       '                     AxisGroup = 2
       '                     LabelPos = xlLabelPositionBelow
       '                     LineColor = vbCyan

4573               End Select
4574               LastPoint = 16
4575               ActiveChart.SeriesCollection.NewSeries
4576               ActiveChart.SeriesCollection(I).Name = SeriesName
4577               ActiveChart.SeriesCollection(I).XValues = XVals
4578               ActiveChart.SeriesCollection(I).Values = YVals
4579               ActiveChart.SeriesCollection(I).Select
4580               ActiveChart.SeriesCollection(I).Points(LastPoint).Select
4581               ActiveChart.SeriesCollection(I).Points(LastPoint).ApplyDataLabels
4582               ActiveChart.SeriesCollection(I).Points(LastPoint).DataLabel.Select
4583               If I < 5 Then
4584                   Selection.ShowSeriesName = True
4585                   Selection.Position = LabelPos
4586               Else
4587                   Selection.ShowSeriesName = False
4588               End If
4589               Selection.ShowValue = False
4590               ActiveChart.SeriesCollection(I).ChartType = xlXYScatterSmoothNoMarkers
4591               ActiveChart.SeriesCollection(I).Select
4592               With Selection.Format.line
4593                   .Visible = msoTrue
4594                   .DashStyle = LineType
4595                   .ForeColor.RGB = LineColor
4596               End With


4597               ActiveChart.SeriesCollection(I).AxisGroup = AxisGroup
4598               ActiveChart.SeriesCollection(I).DataLabels.Font.Size = 8
4599               ActiveChart.SeriesCollection(I).DataLabels.Font.Name = "Arial"
4600           Next I

               'show design point
4601           SeriesName = "=""Design Point"""
4602           XVals = "=" & ShtName & "!$L$63"
4603           YVals = "=" & ShtName & "!$L$64"
4604           LineType = msoLineSolid
4605           AxisGroup = 1
4606           ActiveChart.SeriesCollection.NewSeries
4607           ActiveChart.SeriesCollection(I).Name = SeriesName
4608           ActiveChart.SeriesCollection(I).XValues = XVals
4609           ActiveChart.SeriesCollection(I).Values = YVals
4610           ActiveChart.SeriesCollection(I).Select

4611           Selection.MarkerStyle = 4
4612           Selection.MarkerSize = 7
4613           With Selection.Format.line
4614               .Visible = msoTrue
4615               .Weight = 2.25
4616               .ForeColor.RGB = vbBlack
4617           End With


4618           ActiveChart.Axes(xlValue).Select
4619           ActiveChart.Axes(xlValue).MinimumScaleIsAuto = True
4620           ActiveChart.Axes(xlValue).MaximumScaleIsAuto = True

4621           ActiveChart.Axes(xlValue).MaximumScale = MaxTDH
4622           ActiveChart.Axes(xlValue).MinimumScale = 0
4623           ActiveChart.Axes(xlValue).MajorUnit = Int(MaxTDH / 5)
4624           Selection.TickLabels.NumberFormat = "0"

4625           ActiveChart.Axes(xlValue, xlSecondary).Select
4626           ActiveChart.Axes(xlValue, xlSecondary).MinimumScaleIsAuto = True
4627           ActiveChart.Axes(xlValue, xlSecondary).MaximumScaleIsAuto = True

4628           ActiveChart.Axes(xlValue, xlSecondary).MaximumScale = MaxCurrent
4629           ActiveChart.Axes(xlValue, xlSecondary).MinimumScale = 0
4630           ActiveChart.Axes(xlValue, xlSecondary).MajorUnit = Int(MaxCurrent / 5)
4631           Selection.TickLabels.NumberFormat = "0"

4632           ActiveChart.Axes(xlValue, xlSecondary).HasTitle = True
4633           ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = "Input Power (kW)-Current (A)"
       '        ActiveChart.Axes(xlValue, xlSecondary).AxisTitle.Characters.Text = "Input Power (kW)-Current (A)-Overall Efficiency (%)"
4634           ActiveChart.SetElement (msoElementSecondaryValueAxisTitleRotated)
               'ActiveSheet.PageSetup.PrintArea = "$CA$1:$CI$50"

4635           Range("A1").Select

               'delete all macros in the excel file

               ' Declare variables to access the macros in the workbook.
4636           Dim objProject As VBIDE.VBProject
4637           Dim objComponent As VBIDE.VBComponent
4638           Dim objCode As VBIDE.CodeModule

               ' Get the project details in the workbook.
4639           Set objProject = xlBook.VBProject

               ' Iterate through each component in the project.
4640           For Each objComponent In objProject.VBComponents

                   ' Delete code modules
4641               Set objCode = objComponent.CodeModule
4642               objCode.DeleteLines 1, objCode.CountOfLines

4643               Set objCode = Nothing
4644               Set objComponent = Nothing
4645           Next

4646           Set objProject = Nothing


4647           xlApp.Visible = True                    'show the sheet

4648           xlApp.VBE.ActiveVBProject.VBComponents.Import ParentDirectoryName & sSaveFileMacroFile
4649           xlApp.Run "AssignButton"
4650       End With

       '    Exit Sub

4651   ErrHandler:
           'User pressed the Cancel button

4652       On Error GoTo notopen
4653       If Not xlApp.ActiveWorkbook Is Nothing Then
4654           ActiveWorkbook.CheckCompatibility = False
4655           xlApp.ActiveWorkbook.Save               'save the workbook
               'xlApp.ActiveWorkbook.Close

4656       End If

4657   notopen:

       '    xlApp.Application.Quit

       '    xlApp.Quit
       '    Set xlApp = Nothing

       '    If CommonDialog1.filename <> "" Then
       '        MsgBox CommonDialog1.filename & " has been written.", vbOKOnly, "File Opened"
       '    End If

4658   On Error GoTo vbwErrHandler

4659       Exit Sub
' <VB WATCH>
4660       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Function GetWorksheetTabs(filename As String, WorkSheetName As String)
' <VB WATCH>
4661       On Error GoTo vbwErrHandler
' </VB WATCH>

           'see what worksheet tabs alread exist in the excel worksheet

4662       Dim intSheets As Integer    'number of sheets in the workbook
4663       Dim I As Integer
4664       Dim S As String
4665       Dim ans
4666       Dim NameOK As Boolean

4667       intSheets = xlApp.Worksheets.Count      'how many sheets are there?

           'define a crlf string
4668       S = vbCrLf

4669       For I = 1 To intSheets
4670           S = S & xlApp.Worksheets(I).Name & vbCrLf   'add in the worksheet name
4671       Next I

           'tell the user the names so far and ask if he/she wants to add another
4672       ans = MsgBox("You have the following Worksheet Names in " & filename & ": " & S & "Do you want to add another sheet to this file?", vbYesNo, "Sheets in Excel File")

           'get the answer
4673       If ans = vbNo Then
4674           GetWorksheetTabs = vbNo     'set up flag for when we return to the calling subroutine
4675           Exit Function
4676       End If

           'get worksheet name from user and check to see that it's not already used

4677       NameOK = False  'start assuming that the name is bad

4678       While Not NameOK    'as long as it's bad, stay in this loop
4679           WorkSheetName = InputBox("Enter Worksheet Name for this run.")  'ask for name

4680           If WorkSheetName = "" Then      'if we get a nul return or user presses cancel
4681               GetWorksheetTabs = vbNo
4682               Exit Function
4683           End If

4684           For I = 1 To xlApp.Worksheets.Count     'go through all of the existing sheets
4685               If WorkSheetName = xlApp.Worksheets(I).Name Then        'if the names are the same
4686                   MsgBox "The name " & WorkSheetName & " already exists for a Worksheet.  Please try again.", vbOKOnly, "Bad Worksheet Name"  'tell the user
4687                   NameOK = False
4688                   Exit For
4689               End If
4690               NameOK = True       'if we make it thru say the name is ok
4691           Next I
4692       Wend

4693       xlApp.Worksheets.Add , xlApp.Worksheets(xlApp.Worksheets.Count)     'add a worksheer
4694       xlApp.Worksheets(xlApp.Worksheets.Count).Name = WorkSheetName       'give it the desired name
4695       GetWorksheetTabs = vbYes                                            'say that the results were ok

' <VB WATCH>
4696       Exit Function
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
    End Select
' </VB WATCH>
End Function
Function NewWorkBook() As String
' <VB WATCH>
4697       On Error GoTo vbwErrHandler
' </VB WATCH>

4698       Dim WorkSheetName As String

           'we've just added a new workbook, delete sheet1, sheet2, etc
4699       xlApp.DisplayAlerts = False
4700       While xlApp.Worksheets.Count > 1
4701           xlApp.Worksheets(1).Delete          'delete the sheet
4702       Wend
4703       xlApp.DisplayAlerts = True

4704       WorkSheetName = InputBox("Enter Title Worksheet Name for this run.")    'get the desired name
4705       xlApp.Worksheets(1).Name = WorkSheetName    'and name the sheet

4706       NewWorkBook = WorkSheetName

' <VB WATCH>
4707       Exit Function
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
    End Select
' </VB WATCH>
End Function

Private Sub CalibrateSoftware()
' <VB WATCH>
4708       On Error GoTo vbwErrHandler
' </VB WATCH>
4709           frmCalibrate.Show
               'Calibrating = True

' <VB WATCH>
4710       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Function ParseTEMCModelNo(cmbComboName As ComboBox, ltr As String)
' <VB WATCH>
4711       On Error GoTo vbwErrHandler
' </VB WATCH>
4712       Dim I As Integer
4713       Dim iStart As Integer
4714       Dim iStop As Integer
4715       Dim strCompare As String

4716       For I = 0 To cmbComboName.ListCount - 1                     'go through the combobox entries
4717           iStart = InStr(1, cmbComboName.List(I), "[")
4718           iStop = InStr(1, cmbComboName.List(I), "]")
4719           strCompare = Mid$(cmbComboName.List(I), iStart + 1, iStop - iStart - 1)
4720           If UCase(strCompare) = UCase(ltr) Then   'see when we find the desired index number
4721               cmbComboName.ListIndex = I                                              'if we do, set the combo box
4722               Exit For                                            'and we're done
4723           End If
       '        cmbComboName.ListIndex = -1                             'else, remove any pointer
4724           cmbComboName.ListIndex = cmbComboName.ListCount - 1                           'else, remove any pointer
4725       Next I

4726       txtModelNo.Text = UCase(txtModelNo.Text)
4727       txtModelNo.SelStart = Len(txtModelNo.Text)
' <VB WATCH>
4728       Exit Function
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
    End Select
' </VB WATCH>
End Function
Public Function LoadCombo(cmbComboName As ComboBox, sTableName As String)
       'load all of the pump parameter combo boxes from the tables on the database
' <VB WATCH>
4729       On Error GoTo vbwErrHandler
' </VB WATCH>

4730       Dim I As Integer
4731       Dim sItem As String
4732       Dim iID As Integer
4733       Dim bUseDropdown As Boolean
4734       Dim qy As New ADODB.Command
4735       Dim rs As New ADODB.Recordset

       '    rsPumpParameters.CursorLocation = adUseClient
       '    If sTableName = "Model" Then
       '        rsPumpParameters.Sort = "Model"
       '    Else
       '        rsPumpParameters.Sort = vbNullString
       '    End If
       '    rsPumpParameters.Open sTableName, cnPumpData, adOpenStatic, adLockOptimistic, adCmdTableDirect

4736       qy.ActiveConnection = cnPumpData
4737       If sTableName = "DischargeDiameter" Or sTableName = "SuctionDiameter" Then
4738           qy.CommandText = "SELECT * FROM " & sTableName & " ORDER BY Val(Description)"
4739       Else
4740           qy.CommandText = "SELECT * FROM " & sTableName & " ORDER BY Description"
4741       End If
4742       If sTableName = "SupermarketPumpData" Then
4743           qy.CommandText = "SELECT ID,Model AS Description FROM " & sTableName
4744       End If
4745       rs.CursorLocation = adUseClient
4746       rs.CursorType = adOpenStatic

4747       rs.Open qy


4748       On Error GoTo NoField
4749       bUseDropdown = True
           'sItem = rsPumpParameters.Fields("UseInDropdown")
       '    If bUseDropdown Then
       '        rsPumpParameters.Sort = "Description"
       '    End If
4750       rs.MoveFirst                                'goto the top
4751       For I = 0 To rs.RecordCount - 1             'go through the whole recordset
4752           sItem = rs.Fields("Description")        'get the description
4753           iID = rs.Fields(0)                      'get the index number - primary key
4754           If bUseDropdown Then
       '            If rsPumpParameters.Fields("UseInDropdown").value = True Then
4755                   cmbComboName.AddItem sItem, I                                   'add the description to the combo box
       '                cmbComboName.AddItem sItem                                   'add the description to the combo box
4756                   cmbComboName.ItemData(cmbComboName.NewIndex) = iID              'add the key number into the item data
       '            End If
4757           End If
4758           rs.MoveNext                             'get the next record
4759       Next I
4760       rs.Close
4761       cmbComboName.ListIndex = -1
4762   On Error GoTo vbwErrHandler
4763       Set rs = Nothing
4764       Set qy = Nothing
4765       Exit Function

4766   NoField:
4767       bUseDropdown = False
4768   On Error GoTo vbwErrHandler
4769       Resume Next

' <VB WATCH>
4770       Exit Function
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
    End Select
' </VB WATCH>
End Function
Function SetGraphMax(Plothead) As Integer
' <VB WATCH>
4771       On Error GoTo vbwErrHandler
' </VB WATCH>

4772       Dim I As Integer
4773       Dim m As Single

4774       m = 0
4775       For I = 0 To UBound(Plothead, 2)
4776           If Plothead(1, I) > m Then
4777               m = Plothead(1, I)
4778           End If
4779       Next I
4780       SetGraphMax = 10 * (Int((m / 10) + 0.5) + 1)
4781       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Auto = False
4782       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Maximum = 10 * (Int((m / 10) + 0.5) + 1)
4783       MSChart1.Plot.Axis(VtChAxisIdY).ValueScale.Minimum = 0

' <VB WATCH>
4784       Exit Function
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
    End Select
' </VB WATCH>
End Function
Public Function CalculateSpeed(CoefSq As Double, CoefLin As Double, CoefConstant As Double, InputHP As Double, SG As Double) As Integer
' <VB WATCH>
4785       On Error GoTo vbwErrHandler
' </VB WATCH>
4786       Dim I As Integer
4787       Dim OldResult As Double
4788       Dim NewResult As Double

4789       CalculateSpeed = 0

4790       If SG > 5 Or SG < 0.01 Then
4791           MsgBox "Bad value for SG...must be between 0.01 and 5.", vbOKOnly, "Bad SG Value"
4792           Exit Function
4793       End If

4794       OldResult = 1000
4795       NewResult = 0

4796       I = 1

4797       Do While Abs(NewResult - OldResult) > 0.1
4798           ReDim Preserve results(I)
4799           Select Case I
                   Case 1
4800                   results(I - 1).HP = InputHP
4801               Case 2
4802                   results(I - 1).HP = results(I - 2).HP * SG
4803               Case Else
4804                   results(I - 1).HP = results(I - 2).HP * (results(I - 2).Speed / results(I - 3).Speed) ^ 3
4805           End Select
4806           OldResult = NewResult
4807           results(I - 1).Speed = CalcPoly(CoefSq, CoefLin, CoefConstant, results(I - 1).HP)
4808           NewResult = results(I - 1).Speed
4809           If I > 15 Then
4810               If I = 0 Or I > 15 Then
4811                   MsgBox "Over 15 calculations and no convergence", vbOKOnly, "Too many iterations"
4812                   Exit Function
4813               End If
4814               Exit Function
4815           End If
4816           I = I + 1
4817       Loop
4818       CalculateSpeed = I - 1
' <VB WATCH>
4819       Exit Function
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
    End Select
' </VB WATCH>
End Function
Public Function CalcPoly(CoefSq As Double, CoefLin As Double, CoefConstant As Double, DataIn As Double) As Double
' <VB WATCH>
4820       On Error GoTo vbwErrHandler
' </VB WATCH>
4821       CalcPoly = CoefSq * DataIn ^ 2 + CoefLin * DataIn + CoefConstant
' <VB WATCH>
4822       Exit Function
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
    End Select
' </VB WATCH>
End Function

Sub GetBalanceHoleData(SerialNumber As String, TestDate As String)
' <VB WATCH>
4823       On Error GoTo vbwErrHandler
' </VB WATCH>
4824       If rsBalanceHoles.State = adStateOpen Then
4825           rsBalanceHoles.Close
4826       End If
4827       qyBalanceHoles.CommandText = "SELECT BalanceHoles.*, " & _
                             "IIf([Diameter]=99, 'Slot', [diameter]) as Diameter1, IIf([BoltCircle]=99, 'Unknown', [BoltCircle]) as BoltCircle1 " & _
                             "FROM BalanceHoles " & _
                             "WHERE [SerialNo] = '" & SerialNumber & "' AND [Date] <= #" & TestDate & "# " & _
                             "ORDER BY [Date], Val([BoltCircle]);"

4828       rsBalanceHoles.Open qyBalanceHoles
4829       rsBalanceHoles.Filter = ""

4830       Set dgBalanceHoles.DataSource = rsBalanceHoles

4831       Dim c As Column
4832       For Each c In dgBalanceHoles.Columns
4833           Select Case c.DataField
               Case "BalanceHoleID"
4834               c.Visible = False
4835           Case "SerialNo"
4836               c.Visible = False
4837           Case "Date"
4838               c.Visible = True
4839               c.Alignment = dbgCenter
4840               c.Width = 2000
4841           Case "Number"
4842               c.Visible = True
4843               c.Alignment = dbgCenter
4844               c.Width = 700
4845           Case "Diameter"
4846               c.Visible = False
4847           Case "Diameter1"
4848               c.Caption = "Diameter"
4849               c.Visible = True
4850               c.Alignment = dbgCenter
4851               c.Width = 700
4852           Case "BoltCircle1"
4853               c.Caption = "Bolt Circle"
4854               c.Visible = True
4855               c.Alignment = dbgCenter
4856               c.Width = 800
4857           Case "BoltCircle"
4858               c.Visible = False
4859           Case Else ' hide all other columns.
4860               c.Visible = False
4861           End Select
4862       Next c

' <VB WATCH>
4863       Exit Sub
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
    End Select
' </VB WATCH>
End Sub

Public Sub FixPointsToPlot()
           'count valid data test entry and set points to plot
' <VB WATCH>
4864       On Error GoTo vbwErrHandler
' </VB WATCH>
4865       If DataGrid2.Row = -1 Then
4866           Exit Sub
4867       End If
4868       Dim PresentGridRow As Integer
4869       PresentGridRow = DataGrid2.Row
4870       Dim GridIndex As Integer
4871       UpDown2.value = 8
4872       If DataGrid2.Row <> -1 Then
4873           For GridIndex = 0 To 7
4874               DataGrid2.Row = GridIndex
4875               If DataGrid2.Columns("Flow") = 0 And DataGrid2.Columns("TDH") = 0 Then
4876                   txtUpDn2.Text = GridIndex
4877                   Exit Sub
4878               End If
4879           Next GridIndex
4880       End If
4881       DataGrid2.Row = PresentGridRow
' <VB WATCH>
4882       Exit Sub
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
    End Select
' </VB WATCH>
End Sub


' <VB WATCH> <VBWATCHFINALPROC>
' Procedure added by VB Watch 'ID
Private Sub Form_Initialize() 'ID
    vbwInitializeProtector ' Initialize VB Watch 'ID
End Sub 'ID
' </VB WATCH>
