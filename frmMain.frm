VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{F4732CE3-9A6C-11D2-8018-0080AD70A386}#3.6#0"; "AresButtonPro.ocx"
Object = "{8E27C92E-1264-101C-8A2F-040224009C02}#7.0#0"; "MSCAL.OCX"
Object = "{648A5603-2C6E-101B-82B6-000000000014}#1.1#0"; "MSCOMM32.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{00025600-0000-0000-C000-000000000046}#5.2#0"; "Crystl32.OCX"
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.3#0"; "ARButton.ocx"
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MSHFLXGD.OCX"
Object = "{67397AA1-7FB1-11D0-B148-00A0C922E820}#6.0#0"; "MSADODC.OCX"
Object = "{C64D70BC-E172-42ED-B119-C0CBE641CCA0}#1.9#0"; "WOWFormer.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "The Ultimate PBX Companion Software"
   ClientHeight    =   6780
   ClientLeft      =   4215
   ClientTop       =   3990
   ClientWidth     =   12075
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   452
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   805
   StartUpPosition =   3  'Windows Default
   Begin WOWFormer_ActiveX.WOWFormer WOWFormer 
      Align           =   1  'Align Top
      Height          =   285
      Left            =   0
      TabIndex        =   40
      Top             =   0
      Width           =   12075
      _ExtentX        =   21299
      _ExtentY        =   503
      PictureLeft     =   "frmMain.frx":080A
      PictureMiddle   =   "frmMain.frx":1274
      PictureRight    =   "frmMain.frx":1312
      PictureLeftWidth=   49
      PictureRightWidth=   105
      FormBorderTop   =   "frmMain.frx":13B0
      FormBorderLeft  =   "frmMain.frx":1412
      FormBorderBottom=   "frmMain.frx":1470
      FormBorderRight =   "frmMain.frx":14D2
      FormBorderLeftWidth=   4
      FormBorderBottomHeight=   4
      FormBorderRightWidth=   4
      FormBackground  =   "frmMain.frx":1530
      FormIcon        =   "frmMain.frx":2182
      CaptionSpacing  =   0
      BeginProperty CaptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CaptionColor    =   128
      PictureMaximize =   "frmMain.frx":299C
      PictureMinimize =   "frmMain.frx":2D2E
      PictureClose    =   "frmMain.frx":30C0
      PictureMinimizeToTray=   "frmMain.frx":3452
      ControlBoxSpacing=   3
      ControlBoxRightPadding=   4
      CaptionPrefix   =   "PCBX Lite 1.0> "
      PictureShrink   =   "frmMain.frx":37E4
      PictureCloseDown=   "frmMain.frx":3B76
      PictureMaximizeDown=   "frmMain.frx":3F08
      PictureMinimizeDown=   "frmMain.frx":429A
      PictureShrinkDown=   "frmMain.frx":462C
      PictureMinimizeToTrayDown=   "frmMain.frx":49BE
      PicturePin      =   "frmMain.frx":4D50
      AllowOnTop      =   0   'False
      PicturePinDown  =   "frmMain.frx":50E2
      PicturePinHover =   "frmMain.frx":5474
      PictureMinimizeToTrayHover=   "frmMain.frx":57C6
      PictureShrinkHover=   "frmMain.frx":5B18
      PictureMinimizeHover=   "frmMain.frx":5E6A
      PictureMaximizeHover=   "frmMain.frx":61BC
      PictureCloseHover=   "frmMain.frx":650E
      TrayTip         =   " The Ultimate PBX Companion Software "
      FormMouseIcon   =   "frmMain.frx":6860
      TrayIcon        =   "frmMain.frx":707A
   End
   Begin MSAdodcLib.Adodc datCall 
      Height          =   330
      Left            =   4080
      Top             =   3120
      Visible         =   0   'False
      Width           =   1935
      _ExtentX        =   3413
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
      Connect         =   ""
      OLEDBString     =   ""
      OLEDBFile       =   ""
      DataSourceName  =   ""
      OtherAttributes =   ""
      UserName        =   ""
      Password        =   ""
      RecordSource    =   ""
      Caption         =   "datCall"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid grdCall 
      Bindings        =   "frmMain.frx":7894
      Height          =   2175
      Left            =   240
      TabIndex        =   39
      Top             =   3120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   3836
      _Version        =   393216
      BackColor       =   16777215
      FixedCols       =   0
      BackColorFixed  =   8388608
      ForeColorFixed  =   16777215
      BackColorSel    =   16640978
      ForeColorSel    =   192
      BackColorBkg    =   16640978
      ScrollTrack     =   -1  'True
      GridLines       =   0
      GridLinesFixed  =   0
      SelectionMode   =   1
      AllowUserResizing=   1
      BorderStyle     =   0
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty FontFixed {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin Crystal.CrystalReport crpCall 
      Left            =   4560
      Top             =   4200
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   348160
      WindowControlBox=   -1  'True
      WindowMaxButton =   -1  'True
      WindowMinButton =   -1  'True
      WindowState     =   2
      PrintFileLinesPerPage=   60
      WindowShowCloseBtn=   -1  'True
      WindowShowPrintSetupBtn=   -1  'True
      WindowShowRefreshBtn=   -1  'True
   End
   Begin MSComDlg.CommonDialog cdgReport 
      Left            =   5040
      Top             =   4200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Open a report..."
      Filter          =   "Crystal report|*.rpt|All files|*.*"
   End
   Begin MSACAL.Calendar cal 
      Height          =   2655
      Left            =   5760
      TabIndex        =   31
      Top             =   3480
      Visible         =   0   'False
      Width           =   4935
      _Version        =   524288
      _ExtentX        =   8705
      _ExtentY        =   4683
      _StockProps     =   1
      BackColor       =   16640978
      Year            =   2003
      Month           =   5
      Day             =   15
      DayLength       =   1
      MonthLength     =   1
      DayFontColor    =   0
      FirstDay        =   7
      GridCellEffect  =   0
      GridFontColor   =   10485760
      GridLinesColor  =   -2147483632
      ShowDateSelectors=   -1  'True
      ShowDays        =   -1  'True
      ShowHorizontalGrid=   -1  'True
      ShowTitle       =   -1  'True
      ShowVerticalGrid=   -1  'True
      TitleFontColor  =   10485760
      ValueIsNull     =   0   'False
      BeginProperty DayFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty GridFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.PictureBox picSearch 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDEBD2&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1695
      Left            =   120
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   777
      TabIndex        =   19
      Top             =   1320
      Width           =   11655
      Begin VB.CheckBox chkDeleted 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDEBD2&
         Caption         =   "Deleted"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   10560
         TabIndex        =   34
         Top             =   900
         Width           =   975
      End
      Begin VB.TextBox txtAccount 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   9000
         TabIndex        =   9
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox txtDateTo 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   7080
         Locked          =   -1  'True
         TabIndex        =   8
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtDateFrom 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5760
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   840
         Width           =   975
      End
      Begin VB.TextBox txtCaller 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   3600
         TabIndex        =   6
         Top             =   840
         Width           =   1455
      End
      Begin VB.TextBox txtDialedNumber 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1200
         TabIndex        =   5
         Top             =   840
         Width           =   1695
      End
      Begin VB.CheckBox chkForwarded 
         Appearance      =   0  'Flat
         BackColor       =   &H00FDEBD2&
         Caption         =   "Forwarded"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   10560
         TabIndex        =   4
         Top             =   420
         Value           =   1  'Checked
         Width           =   1215
      End
      Begin VB.TextBox txtCOLine 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   8520
         TabIndex        =   3
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox txtExtension 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   5520
         TabIndex        =   2
         Top             =   360
         Width           =   2055
      End
      Begin VB.TextBox txtClient 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   2760
         TabIndex        =   1
         Top             =   360
         Width           =   1695
      End
      Begin VB.ComboBox cboGroup 
         Appearance      =   0  'Flat
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   600
         TabIndex        =   0
         Top             =   360
         Width           =   1455
      End
      Begin ARESBUTTONLib.AresButton cmdSearch 
         Height          =   480
         Left            =   9840
         TabIndex        =   23
         Top             =   1320
         Width           =   480
         _Version        =   196614
         FastDraw        =   -1  'True
         EffectFilter    =   19
         PictureURL      =   "F:\Joy\Software\Application\PCBXLite\Graphics\Search.jpg"
         PictureOverURL  =   "F:\Joy\Software\Application\PCBXLite\Graphics\Search_Over.jpg"
         PictureDownURL  =   "F:\Joy\Software\Application\PCBXLite\Graphics\Search_Down.jpg"
         PictureDisableURL=   "F:\Joy\Software\Application\PCBXLite\Graphics\Search_Disable.jpg"
         ToolTipString   =   " Search "
         AutoHandCursor  =   -1  'True
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureRES      =   "frmMain.frx":78AA
         PictureOverRES  =   "frmMain.frx":84FC
         PictureDownRES  =   "frmMain.frx":914E
         PictureDisableRES=   "frmMain.frx":9DA0
         HoldingFlag     =   15
         PrevPointer     =   81033456
         _ExtentX        =   847
         _ExtentY        =   847
         _StockProps     =   80
      End
      Begin ARButtonCtrl.ARButton cmdDelete 
         Height          =   375
         Left            =   0
         TabIndex        =   37
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "Delete"
         ForeColor       =   8388608
         ForeColorOnMouse=   4194304
         ForeColorOnFocus=   12582912
         BackColorOnMouse=   16640978
         BorderColor     =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   2
      End
      Begin ARButtonCtrl.ARButton cmdPurge 
         Height          =   375
         Left            =   975
         TabIndex        =   38
         Top             =   1320
         Width           =   975
         _ExtentX        =   1720
         _ExtentY        =   661
         Caption         =   "Purge"
         ForeColor       =   8388608
         ForeColorOnMouse=   4194304
         ForeColorOnFocus=   12582912
         BackColorOnMouse=   16640978
         BorderColor     =   0
         BackColor       =   16777215
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ShowFocus       =   2
      End
      Begin VB.Label lblAccount 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Account"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   8280
         TabIndex        =   30
         Top             =   900
         Width           =   615
      End
      Begin VB.Label lblDateTo 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "To"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   6840
         TabIndex        =   29
         Top             =   900
         Width           =   180
      End
      Begin VB.Label lblDateFrom 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "From"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   5280
         TabIndex        =   28
         Top             =   900
         Width           =   375
      End
      Begin VB.Label lblCaller 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Caller"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   3120
         TabIndex        =   27
         Top             =   900
         Width           =   420
      End
      Begin VB.Label lblDialedNumber 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Dialed number"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   26
         Top             =   900
         Width           =   1050
      End
      Begin VB.Label lblCOLine 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "CO Line"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   7800
         TabIndex        =   25
         Top             =   420
         Width           =   585
      End
      Begin VB.Label lblExtension 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Extension"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   4680
         TabIndex        =   24
         Top             =   420
         Width           =   735
      End
      Begin VB.Label lblClient 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Client"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   2280
         TabIndex        =   22
         Top             =   420
         Width           =   435
      End
      Begin VB.Label lblSearchTitle 
         AutoSize        =   -1  'True
         BackColor       =   &H00800000&
         Caption         =   "  Search Panel"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   0
         TabIndex        =   21
         Top             =   0
         Width           =   2400
      End
      Begin VB.Label lblGroup 
         AutoSize        =   -1  'True
         BackStyle       =   0  'Transparent
         Caption         =   "Group"
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   240
         Left            =   120
         TabIndex        =   20
         Top             =   420
         Width           =   450
      End
   End
   Begin VB.PictureBox picResizer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   5640
      MousePointer    =   8  'Size NW SE
      Picture         =   "frmMain.frx":A9F2
      ScaleHeight     =   20
      ScaleMode       =   0  'User
      ScaleWidth      =   20
      TabIndex        =   16
      Top             =   6360
      Width           =   240
   End
   Begin VB.PictureBox picToolBar 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDEBD2&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   537
      TabIndex        =   14
      Top             =   360
      Width           =   8055
      Begin ARESBUTTONLib.AresButton cmdSettings 
         Height          =   720
         Left            =   1440
         TabIndex        =   36
         Top             =   0
         Width           =   720
         _Version        =   196614
         Stretch         =   1
         AutoSize        =   0   'False
         FastDraw        =   -1  'True
         EffectFilter    =   19
         PictureURL      =   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\Settings.jpg"
         PictureOverURL  =   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\Settings_Over.jpg"
         PictureDownURL  =   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\Settings_Down.jpg"
         PictureDisableURL=   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\Settings_Disable.jpg"
         ToolTipString   =   " Settings "
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureRES      =   "frmMain.frx":AEE4
         PictureOverRES  =   "frmMain.frx":BB36
         HoldingFlag     =   35
         PrevPointer     =   64853960
         _ExtentX        =   1270
         _ExtentY        =   1270
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmdReport 
         Height          =   720
         Left            =   3120
         TabIndex        =   35
         Top             =   0
         Width           =   720
         _Version        =   196614
         Stretch         =   2
         FastDraw        =   -1  'True
         EffectFilter    =   19
         PictureURL      =   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\Report.jpg"
         PictureOverURL  =   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\Report_Over.jpg"
         PictureDownURL  =   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\Report_Down.jpg"
         PictureDisableURL=   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\Report_Disable.jpg"
         ToolTipString   =   " Report... "
         AutoHandCursor  =   -1  'True
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureRES      =   "frmMain.frx":C788
         PictureOverRES  =   "frmMain.frx":D3DA
         PictureDownRES  =   "frmMain.frx":E02C
         PictureDisableRES=   "frmMain.frx":EC7E
         HoldingFlag     =   15
         PrevPointer     =   87419792
         _ExtentX        =   847
         _ExtentY        =   847
         _StockProps     =   80
      End
      Begin VB.CommandButton cmdTestData 
         Caption         =   "Input test data"
         Height          =   495
         Left            =   5400
         TabIndex        =   33
         Top             =   120
         Visible         =   0   'False
         Width           =   1695
      End
      Begin ARESBUTTONLib.AresButton cmdAbout 
         Height          =   720
         Left            =   3960
         TabIndex        =   13
         Top             =   0
         Width           =   720
         _Version        =   196614
         FastDraw        =   -1  'True
         EffectFilter    =   19
         ToolTipBackColor=   16777215
         ToolTipTextColor=   12582912
         ToolTipGradientColor=   12648447
         PictureURL      =   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\About.jpg"
         PictureOverURL  =   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\About_Over.jpg"
         PictureDownURL  =   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\About_Down.jpg"
         ToolTipString   =   " About... "
         AutoHandCursor  =   -1  'True
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureRES      =   "frmMain.frx":F8D0
         PictureOverRES  =   "frmMain.frx":11422
         PictureDownRES  =   "frmMain.frx":12F74
         PictureDisableRES=   "frmMain.frx":14AC6
         HoldingFlag     =   15
         PrevPointer     =   187611352
         _ExtentX        =   1270
         _ExtentY        =   1270
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmdClient 
         Height          =   720
         Left            =   2280
         TabIndex        =   12
         Top             =   0
         Width           =   720
         _Version        =   196614
         FastDraw        =   -1  'True
         EffectFilter    =   19
         ToolTipBackColor=   16777215
         ToolTipTextColor=   4210816
         ToolTipGradientColor=   16761024
         PictureURL      =   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\Client.jpg"
         PictureOverURL  =   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\Client_Over.jpg"
         PictureDownURL  =   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\Client_Down.jpg"
         PictureDisableURL=   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\Client_Disable.jpg"
         ToolTipString   =   " Client Management"
         AutoHandCursor  =   -1  'True
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureRES      =   "frmMain.frx":14AE2
         PictureOverRES  =   "frmMain.frx":16634
         PictureDownRES  =   "frmMain.frx":18186
         PictureDisableRES=   "frmMain.frx":19CD8
         HoldingFlag     =   15
         PrevPointer     =   187254728
         _ExtentX        =   1270
         _ExtentY        =   1270
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmdSettingDataMap 
         Height          =   720
         Left            =   720
         TabIndex        =   11
         Top             =   0
         Width           =   720
         _Version        =   196614
         FastDraw        =   -1  'True
         EffectFilter    =   19
         ToolTipBackColor=   16777215
         ToolTipTextColor=   4210816
         ToolTipGradientColor=   16761024
         PictureURL      =   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\DataMap.jpg"
         PictureOverURL  =   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\DataMap_Over.jpg"
         PictureDownURL  =   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\DataMap_Down.jpg"
         PictureDisableURL=   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\DataMap_Disable.jpg"
         ToolTipString   =   " Data Mapping Wizard "
         AutoHandCursor  =   -1  'True
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureRES      =   "frmMain.frx":1B82A
         PictureOverRES  =   "frmMain.frx":1D37C
         PictureDownRES  =   "frmMain.frx":1EECE
         PictureDisableRES=   "frmMain.frx":20A20
         HoldingFlag     =   15
         PrevPointer     =   187204544
         _ExtentX        =   1270
         _ExtentY        =   1270
         _StockProps     =   80
      End
      Begin ARESBUTTONLib.AresButton cmdSettingCOM 
         Height          =   720
         Left            =   0
         TabIndex        =   10
         Top             =   0
         Width           =   720
         _Version        =   196614
         FastDraw        =   -1  'True
         EffectFilter    =   19
         ToolTipBackColor=   16777215
         ToolTipTextColor=   4210816
         ToolTipGradientColor=   16761024
         PictureURL      =   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\COM.jpg"
         PictureOverURL  =   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\COM_Over.jpg"
         PictureDownURL  =   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\COM_Down.jpg"
         PictureDisableURL=   "F:\Joy\Software\Application\PCBXLite\Graphics\Toolbar\COM_Disable.jpg"
         ToolTipString   =   " RS232C Serial Port Setting "
         AutoHandCursor  =   -1  'True
         BeginProperty ToolTipFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Trebuchet MS"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         BeginProperty TextFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         PictureRES      =   "frmMain.frx":22572
         PictureOverRES  =   "frmMain.frx":240C4
         PictureDownRES  =   "frmMain.frx":25C16
         PictureDisableRES=   "frmMain.frx":27768
         HoldingFlag     =   79
         PrevPointer     =   187206648
         _ExtentX        =   1270
         _ExtentY        =   1270
         _StockProps     =   80
      End
      Begin VB.Label lblNewData 
         Alignment       =   1  'Right Justify
         AutoSize        =   -1  'True
         BackColor       =   &H00000000&
         BorderStyle     =   1  'Fixed Single
         Caption         =   " ### "
         BeginProperty Font 
            Name            =   "Trebuchet MS"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   330
         Left            =   7320
         TabIndex        =   32
         Top             =   120
         Visible         =   0   'False
         Width           =   540
      End
   End
   Begin MSComctlLib.StatusBar stsBar 
      Height          =   240
      Left            =   120
      TabIndex        =   15
      Top             =   6360
      Width           =   5415
      _ExtentX        =   9551
      _ExtentY        =   423
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   4
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   3651
            MinWidth        =   26
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   2196
            MinWidth        =   26
            Text            =   "Rec. ###/###"
            TextSave        =   "Rec. ###/###"
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   6
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1958
            MinWidth        =   26
            TextSave        =   "09/28/2003"
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Style           =   5
            Alignment       =   1
            AutoSize        =   2
            Object.Width           =   1614
            MinWidth        =   26
            TextSave        =   "09:39 PM"
         EndProperty
      EndProperty
   End
   Begin MSCommLib.MSComm comPBX 
      Left            =   4800
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
      CommPort        =   2
      DTREnable       =   -1  'True
      Handshaking     =   3
      InBufferSize    =   8192
      InputLen        =   82
      OutBufferSize   =   0
      ParityReplace   =   0
      RThreshold      =   82
      InputMode       =   1
   End
   Begin VB.Label lblVResizer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   6210
      Left            =   11880
      MousePointer    =   9  'Size W E
      TabIndex        =   18
      Top             =   360
      Width           =   60
   End
   Begin VB.Label lblHResizer 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      ForeColor       =   &H80000008&
      Height          =   60
      Left            =   120
      MousePointer    =   7  'Size N S
      TabIndex        =   17
      Top             =   6600
      Width           =   8160
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim calDate As TextBox, NewDataAdded As Long

Private Sub LoadSearchList()
With Query("SELECT * FROM tblClientGroup")
    cboGroup.AddItem ""
    While Not .EOF
        cboGroup.AddItem .Fields("ClientGroupName")
        .MoveNext
        DoEvents
    Wend
End With

End Sub

Private Sub cal_Click()
calDate = cal.Value
cal.Visible = False
End Sub

Private Sub cal_LostFocus()
cal.Visible = False
End Sub

Private Sub cboGroup_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSearch_MouseClick
End Sub

Private Sub chkDeleted_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSearch_MouseClick

End Sub

Private Sub chkForwarded_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSearch_MouseClick

End Sub

Private Sub cmdAbout_MouseClick()
frmAbout.Show
End Sub

Private Sub cmdClient_MouseClick()
frmClientList.Show
End Sub

Private Sub cmdDelete_Click()
Dim SQL As String

   On Error GoTo cmdDelete_Click_Error

If cmdDelete.Caption = "Delete" Then
    If datCall.RecordSource = "qryCallList" Then
        MsgBox "Please specify a search criteria before you delete any record.", vbExclamation + vbOKOnly, "No search criteria!"
        Exit Sub
    End If
    
    If MsgBox("Are you sure you want to delete the call records below?", vbQuestion + vbYesNo + vbDefaultButton2, "Delete record?") = vbNo Then Exit Sub
    
    SQL = "UPDATE tblCall LEFT JOIN (tblClientGroup RIGHT JOIN tblClient ON tblClientGroup.ClientGroupID = tblClient.ClientGroupID) ON tblCall.Extension = tblClient.Extension SET tblCall.Deleted = Yes "
Else
    SQL = "UPDATE tblCall LEFT JOIN (tblClientGroup RIGHT JOIN tblClient ON tblClientGroup.ClientGroupID = tblClient.ClientGroupID) ON tblCall.Extension = tblClient.Extension SET tblCall.Deleted = No "
End If
    
SQL = SQL & Mid(datCall.RecordSource, InStr(datCall.RecordSource, "WHERE"))
QueryExec SQL
datCall.Refresh

   On Error GoTo 0
   Exit Sub

cmdDelete_Click_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdDelete_Click of Form frmMain"
End Sub

Private Sub cmdPurge_Click()
If MsgBox("Are you sure you want to purge the deleted records? This will purge all deleted records and this action cannot be undone.", vbQuestion + vbYesNo + vbDefaultButton2, "Purge record?") = vbNo Then Exit Sub

QueryExec "DELETE FROM tblCall WHERE tblCall.Deleted = Yes"
datCall.Refresh

End Sub

Private Sub cmdReport_MouseClick()
On Error Resume Next

If Query(datCall.RecordSource).RecordCount = 0 Then
    MsgBox "Sorry, there is no record to create report for. At least one call record is required to generate a report.", vbExclamation + vbOKOnly, "No record to report"
    Exit Sub
End If

cdgReport.ShowOpen
If Err Then Exit Sub

If datCall.RecordSource = "qryCallList" Then cmdSearch_MouseClick
QueryExec "UPDATE tblCall LEFT JOIN (tblClientGroup RIGHT JOIN tblClient ON tblClientGroup.ClientGroupID=tblClient.ClientGroupID) ON tblCall.Extension=tblClient.Extension SET DoReport = Yes " & Mid(datCall.RecordSource, InStr(datCall.RecordSource, "WHERE"))

crpCall.DataFiles(0) = CheckPath(App.Path, True) & "\PCBXLite.mdb"
crpCall.ReportFileName = cdgReport.FileName
crpCall.DiscardSavedData = True
crpCall.Action = 1

QueryExec "UPDATE tblCall SET DoReport = No"
End Sub

Private Sub cmdSearch_MouseClick()
On Error GoTo cmdSearch_MouseClick_Error

Dim SQL As String

SQL = "SELECT tblClientGroup.ClientGroupName AS [Group], tblClient.LastName + ', ' + tblClient.FirstName AS Client, tblCall.Extension, tblCall.CallDate AS [Date], tblCall.CallTime AS [Time], tblCall.COLine AS [CO Line], tblCall.Forwarding AS Forwarded, tblCall.DialedNumber AS [Dialed number], tblCall.Caller, tblCall.RingDuration AS [Ring duration], tblCall.Duration, tblCall.Account FROM tblCall LEFT JOIN (tblClientGroup RIGHT JOIN tblClient ON tblClientGroup.ClientGroupID = tblClient.ClientGroupID) ON tblCall.Extension = tblClient.Extension WHERE 1=1 "

If cboGroup.Text <> "" Then SQL = SQL & " AND INSTR(LCASE(tblClientGroup.ClientGroupName), '" & LCase(cboGroup.Text) & "')>0 "
If txtClient <> "" Then SQL = SQL & "AND (INSTR(LCASE(tblClient.FirstName), '" & LCase(txtClient) & "')> 0 OR INSTR(LCASE(tblClient.LastName), '" & LCase(txtClient) & "')> 0) "
If txtExtension <> "" Then SQL = SQL & "AND tblCall.Extension IN (" & txtExtension & ") "
If txtCOLine <> "" Then SQL = SQL & "AND tblCall.COLine IN (" & txtCOLine & ") "
If chkForwarded.Value = vbUnchecked Then SQL = SQL & "AND tblCall.Forwarding= 0 "
If txtDialedNumber <> "" Then SQL = SQL & "AND INSTR(tblCall.DialedNumber, '" & txtDialedNumber & "')>0 "
If txtCaller <> "" Then SQL = SQL & "AND INSTR(tblCall.Caller, '" & txtCaller & "')>0 "
If txtDateFrom <> "" Then SQL = SQL & "AND tblCall.CallDate >= #" & txtDateFrom & "# "
If txtDateTo <> "" Then SQL = SQL & "AND tblCall.CallDate <= #" & txtDateTo & "# "
If txtAccount <> "" Then SQL = SQL & "AND INSTR(tblCall.Account, '" & txtAccount & "')>0 "
If chkDeleted.Value = vbChecked Then
    SQL = SQL & "AND tblCall.Deleted= Yes "
    cmdDelete.Caption = "Undelete"
    cmdPurge.Visible = True
Else
    SQL = SQL & "AND tblCall.Deleted= No "
    cmdDelete.Caption = "Delete"
    cmdPurge.Visible = False
End If



datCall.RecordSource = SQL
datCall.Refresh

Dim r As Long, c As Long
grdCall.Col = 6
For r = 1 To grdCall.Rows - 1
    grdCall.Row = r
    If grdCall.Text = "True" Then grdCall.Text = "Yes" Else grdCall.Text = "No"
Next

lblNewData.Visible = False
NewDataAdded = 0

stsBar.Panels(2) = "Record count: " & Query(datCall.RecordSource).RecordCount

   On Error GoTo 0
   Exit Sub

cmdSearch_MouseClick_Error:

    MsgBox "Error " & Err.Number & " (" & Err.Description & ") in procedure cmdSearch_MouseClick of Form frmMain"
End Sub

Private Sub cmdSettingCOM_MouseClick()
frmCOM.Show vbModal
End Sub

Private Sub cmdSettingDataMap_MouseClick()
frmDataMapper.Show vbModal
End Sub

Private Sub cmdSettings_MouseClick()
frmSettings.Show vbModal
End Sub

Private Sub cmdTestData_Click()
InputTestData
End Sub

Private Sub comPBX_OnComm()
    Select Case comPBX.CommEvent
        Case comEvReceive 'Recieve data from port
            Dim Buffer As Variant
            Buffer = comPBX.Input
            COMData = COMData & StrConv(Buffer, vbUnicode)
            COM2DB 'Transfer raw data into the database
        Case comEvCTS 'Other Events
            stsBar.Panels(1) = "'Clear To Send (CTS)' signal detected"
        Case comEvDSR
            stsBar.Panels(1) = "'Data Set Ready (DSR)' signal detected"
        Case comEvCD
            stsBar.Panels(1) = "Change in 'Carrier Detect' line"
        
        Case comEventBreak 'Error Signals
            stsBar.Panels(1) = "Communication Error : Break Signal Recieved"
        Case comEventFrame
            stsBar.Panels(1) = "Communication Error : Framing Error Detected"
        Case comEventOverrun
            stsBar.Panels(1) = "Communication Error : Could not read - data lost"
        Case comEventRxOver
            stsBar.Panels(1) = "Communication Error : Recieve buffer overflow, cannot read rest of the data"
        Case comEventRxParity
            stsBar.Panels(1) = "Communication Error : Parity error detected"
        Case comEventDCB
            stsBar.Panels(1) = "Communication Error : Unexpected error retrieving Device Control Block (DCB) for the port"
    End Select

End Sub

Private Sub Form_Load()
On Error GoTo Error_Form_Load

lblHResizer.BorderStyle = 0
lblVResizer.BorderStyle = 0

'grdCall_.DatabaseName = GlobalADOConnectionString
'grdCall_.RecordSource = "SELECT tblClientGroup.ClientGroupName AS [Group], tblClient.LastName + ', ' + tblClient.FirstName AS Client, tblCall.Extension, tblCall.CallDate AS [Date], tblCall.CallTime AS [Time], tblCall.COLine AS [CO Line], tblCall.Forwarding AS Forwarded, tblCall.DialedNumber AS [Dialed number], tblCall.Caller, tblCall.RingDuration AS [Ring duration], tblCall.Duration, tblCall.Account FROM tblCall LEFT JOIN (tblClientGroup RIGHT JOIN tblClient ON tblClientGroup.ClientGroupID = tblClient.ClientGroupID) ON tblCall.Extension = tblClient.Extension"
'grdCall_.RecordSource = "qryCallList"
'grdCall_.Rebind

LoadSearchList

comPBX.CommPort = COM.Port
comPBX.Settings = COM.BaudRate & "," & Left(COM.Parity, 1) & "," & COM.DataBit & "," & COM.StopBit
comPBX.HandShaking = COM.HandShaking
comPBX.PortOpen = True

WOWFormer.CaptionScrollSpeed = 200

cmdSettingCOM.AutoHandCursor = True
cmdSettingDataMap.AutoHandCursor = True
cmdClient.AutoHandCursor = True
cmdReport.AutoHandCursor = True
cmdAbout.AutoHandCursor = True
cmdSearch.AutoHandCursor = True
cmdSettings.AutoHandCursor = True

datCall.ConnectionString = GlobalADOConnectionString
cmdSearch_MouseClick

Exit Sub

Error_Form_Load:
Select Case Err.Number
Case Else
    MsgBox "Error#" & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "PCBX Lite"
    Resume Next
End Select
End Sub

Private Sub Form_Resize()
On Error Resume Next

picToolBar.Move 4, 19, Me.ScaleWidth - (4 + 4), 48
stsBar.Move 4, Me.ScaleHeight - 4 - stsBar.Height, Me.ScaleWidth - 8 - 16, stsBar.Height
picResizer.Move stsBar.Width + 4, stsBar.Top, 20, 20

lblVResizer.Move Me.ScaleWidth - 4, 19, 4, Me.ScaleHeight - 19 - picResizer.Height
lblHResizer.Move 0, Me.ScaleHeight - 4, Me.ScaleWidth - picResizer.Width, 4

picSearch.Move 4, 19 + picToolBar.Height, Me.ScaleWidth - (4 + 4), 113
grdCall.Move 4, 19 + picToolBar.Height + picSearch.Height, Me.ScaleWidth - (4 + 4), Me.ScaleHeight - 19 - picToolBar.Height - picSearch.Height - stsBar.Height - 4

End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next

comPBX.PortOpen = False

'Unload all other loaded windows so that the application doesn't hang up!
Dim frmObj As Form

For Each frmObj In Forms
    Unload frmObj
Next
End Sub

Private Sub lblHResizer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And Me.WindowState = vbNormal Then
    Me.Height = Me.Height + Y - (WOWFormer.FormBorderBottomHeight - 1) * Screen.TwipsPerPixelY
    WOWFormer.Refresh
End If
End Sub

Private Sub lblVResizer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And Me.WindowState = vbNormal Then
    Me.Width = Me.Width + X - (WOWFormer.FormBorderRightWidth - 1) * Screen.TwipsPerPixelX
End If
End Sub

Private Sub picResizer_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 And Me.WindowState = vbNormal Then
    Me.Width = Me.Width + X * Screen.TwipsPerPixelX - 19 * Screen.TwipsPerPixelX
    Me.Height = Me.Height + Y * Screen.TwipsPerPixelY - 19 * Screen.TwipsPerPixelY
End If
End Sub

Private Sub picSearch_Resize()
lblSearchTitle.Move 0, 0, picSearch.ScaleWidth
cmdSearch.Move picSearch.ScaleWidth - cmdSearch.Width, picSearch.ScaleHeight - cmdSearch.Height
End Sub


Private Sub picToolBar_Resize()
lblNewData.Move picToolBar.ScaleWidth - lblNewData.Width, (picToolBar.ScaleHeight - lblNewData.Height) / 2
End Sub

Private Sub txtAccount_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSearch_MouseClick

End Sub

Private Sub txtCaller_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSearch_MouseClick

End Sub

Private Sub txtClient_GotFocus()
SelectAllText txtClient
End Sub

Private Sub txtClient_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSearch_MouseClick

End Sub

Private Sub txtCOLine_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSearch_MouseClick

End Sub

Private Sub txtDateFrom_GotFocus()
Set calDate = txtDateFrom
cal.Move txtDateFrom.Left, txtDateFrom.Top + txtDateFrom.Height + picSearch.Height
cal.Visible = True
cal.Value = txtDateFrom
cal.SetFocus
End Sub

Private Sub txtDateFrom_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSearch_MouseClick

End Sub

Private Sub txtDateTo_GotFocus()
Set calDate = txtDateTo
cal.Move txtDateTo.Left, txtDateTo.Top + txtDateTo.Height + picSearch.Height
cal.Visible = True
cal.Value = txtDateTo
cal.SetFocus
End Sub

Private Sub txtDateTo_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSearch_MouseClick

End Sub

Private Sub txtDialedNumber_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSearch_MouseClick

End Sub

Private Sub txtExtension_KeyPress(KeyAscii As Integer)
If KeyAscii = 13 Then cmdSearch_MouseClick

End Sub

Private Sub WOWFormer_AfterRestore()
picResizer.Visible = True

End Sub

Private Sub WOWFormer_AfterShrink()
picResizer.Visible = False
End Sub

Private Sub COM2DB()
If Len(COMData) < DataMap.EndMarkerStart + Len(DataMap.EndMarker) - 1 Then 'Insufficient data!
    Exit Sub
Else
    If InStr(COMData, DataMap.StartMarker1) < 1 Then 'No StartMarker1!
        COMData = Mid(COMData, Len(COMData) - DataMap.StartMarker1Start) + 1
        Exit Sub
    Else
        'Remove gurbage at the beginning
        COMData = Mid(COMData, InStr(COMData, DataMap.StartMarker1) - DataMap.StartMarker1Start + 1)
        
        If Mid(COMData, DataMap.StartMarker2Start, Len(DataMap.StartMarker2)) <> DataMap.StartMarker2 Or Mid(COMData, DataMap.EndMarkerStart, Len(DataMap.EndMarker)) <> DataMap.EndMarker Then
            'Invalid data!
            COMData = Mid(COMData, DataMap.StartMarker1Start + 1)
            COM2DB 'Check for data in the rest of the stream
        Else
            'Found at least one valid data
            Dim SQL As String
            SQL = "INSERT INTO tblCall (CallDate, CallTime, Extension, COLine, Forwarding, DialedNumber, Caller, RingDuration, Duration, Account) VALUES ("
            
            If DataMap.SystemDate Then 'Set system date to the CallDate field
                SQL = SQL & "#" & Date & "#, "
            Else 'Extrct CallDate
                SQL = SQL & "#" & Mid(COMData, DataMap.DateStart, DataMap.DateLength) & "#, "
            End If
            
            If DataMap.SystemTime Then 'Set system time to the CallTime field
                SQL = SQL & "#" & Time & "#, "
            Else 'Extract CallTime field
                SQL = SQL & "#" & Mid(COMData, DataMap.TimeStart, DataMap.TimeLength) & "#, "
            End If
            
            SQL = SQL & Mid(COMData, DataMap.ExtensionStart, DataMap.ExtensionLength) & ", " 'Extension
            SQL = SQL & Mid(COMData, DataMap.COLineStart, DataMap.COLineLength) & ", " 'COLine
            
            If Mid(COMData, DataMap.ForwardingFlagStart, Len(DataMap.ForwardingFlag)) = DataMap.ForwardingFlag Then
                SQL = SQL & "1, " 'Call is forwarded
            Else
                SQL = SQL & "0, " 'Call is not forwarded
            End If
            
            SQL = SQL & "'" & Trim(Mid(COMData, DataMap.DialedNumberStart, DataMap.DialedNumberLength)) & "', " 'Dialed number
            
            If DataMap.CallerDontRead Then
                SQL = SQL & "'', " 'Don't read the Caller ID
            Else
                SQL = SQL & "'" & Trim(Mid(COMData, DataMap.CallerStart, DataMap.CallerLength)) & "', " 'Caller ID
            End If
            
            'Check if an incoming call
            If Mid(COMData, DataMap.IncomingMarkerStart, Len(DataMap.IncomingMarker)) = DataMap.IncomingMarker Then
                'Extract the Ring Duration
                SQL = SQL & "'" & Replace(Mid(COMData, DataMap.RingDurationStart, DataMap.RingDurationLength), "'", "''") & "', "
            Else
                SQL = SQL & "'', " 'Not an incoming call, so no Ring Duration
            End If
            
            SQL = SQL & "'" & Replace(Mid(COMData, DataMap.DurationStart, DataMap.DurationLength), "'", "''") & "', " 'Duration
            
            If DataMap.AccountDontRead Then 'Don't read the Account field
                SQL = SQL & "'', "
            Else
                SQL = SQL & "'" & Trim(Mid(COMData, DataMap.AccountStart, DataMap.AccountLength)) & "'"   'Account
            End If
            
            SQL = SQL & ")"
            QueryExec SQL 'Insert the data into the database
            NewDataAdded = NewDataAdded + 1
            
            'Remove the transferred data from the data stream
            COMData = Mid(COMData, DataMap.EndMarkerStart + Len(DataMap.EndMarker))
            
            COM2DB 'Check for any other data in the rest of the stream
        End If
    End If
End If

If NewDataAdded > 0 Then
    lblNewData = " " & NewDataAdded & " "
    lblNewData.Visible = True
End If
End Sub

Sub InputTestData()
Dim strData As String

strData = strData & "05/15/03 17:18 102 01 8825964                        01:02'04  000127" & vbCrLf
strData = strData & "05/15/03 23:23 100 01*INCOMING                 0'23  14:00'44  000127" & vbCrLf
strData = strData & "05/15/03 23:23 104 01 019359319                      00:12'34  000125" & vbCrLf

COMData = strData
COM2DB
End Sub

Public Sub ReloadData()
cmdSearch_MouseClick
End Sub
