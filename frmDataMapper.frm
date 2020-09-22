VERSION 5.00
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.3#0"; "ARButton.ocx"
Object = "{C64D70BC-E172-42ED-B119-C0CBE641CCA0}#1.9#0"; "WOWFormer.ocx"
Begin VB.Form frmDataMapper 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Data mapping wizard"
   ClientHeight    =   8385
   ClientLeft      =   3360
   ClientTop       =   1650
   ClientWidth     =   9255
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmDataMapper.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   559
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   617
   StartUpPosition =   3  'Windows Default
   Begin WOWFormer_ActiveX.WOWFormer WOWFormer 
      Align           =   1  'Align Top
      Height          =   285
      Left            =   0
      TabIndex        =   75
      Top             =   0
      Width           =   9255
      _ExtentX        =   16325
      _ExtentY        =   503
      PictureLeft     =   "frmDataMapper.frx":080A
      PictureMiddle   =   "frmDataMapper.frx":1274
      PictureRight    =   "frmDataMapper.frx":1312
      PictureLeftWidth=   49
      PictureRightWidth=   47
      FormBorderTop   =   "frmDataMapper.frx":13B0
      FormBorderLeft  =   "frmDataMapper.frx":1412
      FormBorderBottom=   "frmDataMapper.frx":1470
      FormBorderRight =   "frmDataMapper.frx":14D2
      FormBorderLeftWidth=   4
      FormBorderBottomHeight=   4
      FormBorderRightWidth=   4
      FormBackground  =   "frmDataMapper.frx":1530
      AllowMaximize   =   0   'False
      FormIcon        =   "frmDataMapper.frx":2182
      AllowMinimize   =   0   'False
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
      PictureMaximize =   "frmDataMapper.frx":299C
      PictureMinimize =   "frmDataMapper.frx":2D2E
      PictureClose    =   "frmDataMapper.frx":30C0
      PictureMinimizeToTray=   "frmDataMapper.frx":3452
      ControlBoxSpacing=   3
      CaptionPrefix   =   "PCBX Lite 1.0> "
      PictureShrink   =   "frmDataMapper.frx":37E4
      MinimizeToTray  =   0   'False
      PictureCloseDown=   "frmDataMapper.frx":3B76
      PictureMaximizeDown=   "frmDataMapper.frx":3F08
      PictureMinimizeDown=   "frmDataMapper.frx":429A
      PictureShrinkDown=   "frmDataMapper.frx":462C
      PictureMinimizeToTrayDown=   "frmDataMapper.frx":49BE
      SnapTolerance   =   700
      PicturePin      =   "frmDataMapper.frx":4D50
      AllowOnTop      =   0   'False
      PicturePinDown  =   "frmDataMapper.frx":50E2
      PicturePinHover =   "frmDataMapper.frx":5474
      PictureMinimizeToTrayHover=   "frmDataMapper.frx":57C6
      PictureShrinkHover=   "frmDataMapper.frx":5B18
      PictureMinimizeHover=   "frmDataMapper.frx":5E6A
      PictureMaximizeHover=   "frmDataMapper.frx":61BC
      PictureCloseHover=   "frmDataMapper.frx":650E
      TrayTip         =   " Data mapping wizard "
      FormMouseIcon   =   "frmDataMapper.frx":6860
      TrayIcon        =   "frmDataMapper.frx":707A
   End
   Begin ARButtonCtrl.ARButton cmdEndMarkerXP 
      Height          =   345
      Left            =   2415
      TabIndex        =   48
      Top             =   7245
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   609
      Caption         =   "End marker"
      ForeColor       =   -2147483630
      ForeColorOnMouse=   8388608
      ForeColorOnFocus=   12582912
      BackColorOnMouse=   16640978
      BackColor       =   16761024
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
      BorderStyle     =   2
   End
   Begin ARButtonCtrl.ARButton cmdAccountXP 
      Height          =   345
      Left            =   2415
      TabIndex        =   47
      Top             =   6900
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   609
      Caption         =   "Account"
      ForeColor       =   -2147483630
      ForeColorOnMouse=   8388608
      ForeColorOnFocus=   12582912
      BackColorOnMouse=   16640978
      BackColor       =   16761024
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
      BorderStyle     =   2
   End
   Begin ARButtonCtrl.ARButton cmdDurationXP 
      Height          =   345
      Left            =   2415
      TabIndex        =   46
      Top             =   6555
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   609
      Caption         =   "Duration"
      ForeColor       =   -2147483630
      ForeColorOnMouse=   8388608
      ForeColorOnFocus=   12582912
      BackColorOnMouse=   16640978
      BackColor       =   16761024
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
      BorderStyle     =   2
   End
   Begin ARButtonCtrl.ARButton cmdRingDurationXP 
      Height          =   345
      Left            =   2415
      TabIndex        =   45
      Top             =   6210
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   609
      Caption         =   "Ring duration"
      ForeColor       =   -2147483630
      ForeColorOnMouse=   8388608
      ForeColorOnFocus=   12582912
      BackColorOnMouse=   16640978
      BackColor       =   16761024
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
      BorderStyle     =   2
   End
   Begin ARButtonCtrl.ARButton cmdCallerXP 
      Height          =   345
      Left            =   2415
      TabIndex        =   44
      Top             =   5865
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   609
      Caption         =   "Caller"
      ForeColor       =   -2147483630
      ForeColorOnMouse=   8388608
      ForeColorOnFocus=   12582912
      BackColorOnMouse=   16640978
      BackColor       =   16761024
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
      BorderStyle     =   2
   End
   Begin ARButtonCtrl.ARButton cmdIncomingMarkerXP 
      Height          =   345
      Left            =   2415
      TabIndex        =   43
      Top             =   5520
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   609
      Caption         =   "Incoming marker"
      ForeColor       =   -2147483630
      ForeColorOnMouse=   8388608
      ForeColorOnFocus=   12582912
      BackColorOnMouse=   16640978
      BackColor       =   16761024
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
      BorderStyle     =   2
   End
   Begin ARButtonCtrl.ARButton cmdDailedNumberXP 
      Height          =   345
      Left            =   2415
      TabIndex        =   42
      Top             =   5175
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   609
      Caption         =   "Dialed number"
      ForeColor       =   -2147483630
      ForeColorOnMouse=   8388608
      ForeColorOnFocus=   12582912
      BackColorOnMouse=   16640978
      BackColor       =   16761024
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
      BorderStyle     =   2
   End
   Begin ARButtonCtrl.ARButton cmdForwardingFlagXP 
      Height          =   345
      Left            =   2415
      TabIndex        =   41
      Top             =   4830
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   609
      Caption         =   "Forwarding flag"
      ForeColor       =   -2147483630
      ForeColorOnMouse=   8388608
      ForeColorOnFocus=   12582912
      BackColorOnMouse=   16640978
      BackColor       =   16761024
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
      BorderStyle     =   2
   End
   Begin ARButtonCtrl.ARButton cmdCOLineXP 
      Height          =   345
      Left            =   2415
      TabIndex        =   40
      Top             =   4485
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   609
      Caption         =   "CO Line"
      ForeColor       =   -2147483630
      ForeColorOnMouse=   8388608
      ForeColorOnFocus=   12582912
      BackColorOnMouse=   16640978
      BackColor       =   16761024
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
      BorderStyle     =   2
   End
   Begin ARButtonCtrl.ARButton cmdExtensionXP 
      Height          =   345
      Left            =   2415
      TabIndex        =   39
      Top             =   4140
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   609
      Caption         =   "Extension"
      ForeColor       =   -2147483630
      ForeColorOnMouse=   8388608
      ForeColorOnFocus=   12582912
      BackColorOnMouse=   16640978
      BackColor       =   16761024
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
      BorderStyle     =   2
   End
   Begin ARButtonCtrl.ARButton cmdTimeXP 
      Height          =   345
      Left            =   2415
      TabIndex        =   38
      Top             =   3795
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   609
      Caption         =   "Time"
      ForeColor       =   -2147483630
      ForeColorOnMouse=   8388608
      ForeColorOnFocus=   12582912
      BackColorOnMouse=   16640978
      BackColor       =   16761024
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
      BorderStyle     =   2
   End
   Begin ARButtonCtrl.ARButton cmdDateXP 
      Height          =   345
      Left            =   2415
      TabIndex        =   37
      Top             =   3450
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   609
      Caption         =   "Date"
      ForeColor       =   -2147483630
      ForeColorOnMouse=   8388608
      ForeColorOnFocus=   12582912
      BackColorOnMouse=   16640978
      BackColor       =   16761024
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
      BorderStyle     =   2
   End
   Begin ARButtonCtrl.ARButton cmdStartMarker1XP 
      Height          =   345
      Left            =   2415
      TabIndex        =   36
      Top             =   3105
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   609
      Caption         =   "Start marker 2"
      ForeColor       =   -2147483630
      ForeColorOnMouse=   8388608
      ForeColorOnFocus=   12582912
      BackColorOnMouse=   16640978
      BackColor       =   16761024
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
      BorderStyle     =   2
   End
   Begin ARButtonCtrl.ARButton cmdStartMarkerXP 
      Height          =   330
      Left            =   2415
      TabIndex        =   35
      Top             =   2775
      Width           =   1545
      _ExtentX        =   2725
      _ExtentY        =   582
      Caption         =   "Start marker 1"
      ForeColor       =   -2147483630
      ForeColorOnMouse=   8388608
      ForeColorOnFocus=   12582912
      BackColorOnMouse=   16640978
      BackColor       =   16761024
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
      BorderStyle     =   2
   End
   Begin VB.CommandButton cmdEndMarker 
      Appearance      =   0  'Flat
      Caption         =   "End marker"
      Height          =   360
      Left            =   720
      TabIndex        =   64
      Top             =   7245
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdAccount 
      Appearance      =   0  'Flat
      Caption         =   "Account"
      Height          =   360
      Left            =   720
      TabIndex        =   63
      Top             =   6900
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdDuration 
      Appearance      =   0  'Flat
      Caption         =   "Duration"
      Height          =   360
      Left            =   720
      TabIndex        =   62
      Top             =   6555
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdRingDuration 
      Appearance      =   0  'Flat
      Caption         =   "Ring duration"
      Height          =   360
      Left            =   720
      TabIndex        =   61
      Top             =   6210
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdCaller 
      Appearance      =   0  'Flat
      Caption         =   "Caller"
      Height          =   360
      Left            =   720
      TabIndex        =   60
      Top             =   5865
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdIncomingMarker 
      Appearance      =   0  'Flat
      Caption         =   "Incoming marker"
      Height          =   360
      Left            =   720
      TabIndex        =   59
      Top             =   5520
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdDialedNumber 
      Appearance      =   0  'Flat
      Caption         =   "Dialed number"
      Height          =   360
      Left            =   720
      TabIndex        =   58
      Top             =   5175
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdForwardingFlag 
      Appearance      =   0  'Flat
      Caption         =   "Forwarding flag"
      Height          =   360
      Left            =   720
      TabIndex        =   57
      Top             =   4830
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdCOLine 
      Appearance      =   0  'Flat
      Caption         =   "CO Line"
      Height          =   360
      Left            =   720
      TabIndex        =   56
      Top             =   4485
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdExtension 
      Appearance      =   0  'Flat
      Caption         =   "Extension"
      Height          =   360
      Left            =   720
      TabIndex        =   55
      Top             =   4140
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdTime 
      Appearance      =   0  'Flat
      Caption         =   "Time"
      Height          =   360
      Left            =   720
      TabIndex        =   53
      Top             =   3795
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdDate 
      Appearance      =   0  'Flat
      Caption         =   "Date"
      Height          =   360
      Left            =   720
      TabIndex        =   51
      Top             =   3450
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.CommandButton cmdStartMarker2 
      Appearance      =   0  'Flat
      Caption         =   "Start marker 2"
      Height          =   360
      Left            =   720
      TabIndex        =   50
      Top             =   3105
      Visible         =   0   'False
      Width           =   1575
   End
   Begin PCBXLite.WizardTitle WizardTitle 
      Height          =   975
      Left            =   60
      Top             =   285
      Width           =   9135
      _ExtentX        =   16113
      _ExtentY        =   1720
      Title           =   " Data Mapping Wizard"
      Description     =   $"frmDataMapper.frx":7894
      TitleForeColor  =   16777215
      TitleBackColor  =   8388608
      DescriptionForeColor=   12632256
      DescriptionBackColor=   8388608
      BeginProperty TitleFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty DescriptionFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CheckBox chkAccount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Check1"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   2.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7440
      TabIndex        =   30
      Top             =   6900
      Width           =   195
   End
   Begin VB.TextBox txtAccountLength 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   5160
      TabIndex        =   29
      Text            =   "Text2"
      Top             =   6900
      Width           =   855
   End
   Begin VB.TextBox txtAccountStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   4080
      TabIndex        =   28
      Text            =   "Text1"
      Top             =   6900
      Width           =   975
   End
   Begin VB.CheckBox chkCaller 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Don't read"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   2.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7440
      TabIndex        =   23
      Top             =   5955
      Width           =   195
   End
   Begin VB.CheckBox chkTime 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Read system time"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7440
      TabIndex        =   10
      Top             =   3855
      Width           =   195
   End
   Begin VB.CheckBox chkDate 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Read system date"
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   7440
      TabIndex        =   7
      Top             =   3510
      Width           =   195
   End
   Begin VB.TextBox txtStartMarker2 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   6120
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   3105
      Width           =   1215
   End
   Begin VB.TextBox txtStartMarker2Start 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   4080
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   3105
      Width           =   975
   End
   Begin VB.TextBox txtEndMarker 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   6120
      TabIndex        =   32
      Text            =   "Text1"
      Top             =   7245
      Width           =   1215
   End
   Begin VB.TextBox txtForwardingFlag 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   6120
      TabIndex        =   16
      Text            =   "Text1"
      Top             =   4830
      Width           =   1215
   End
   Begin VB.TextBox txtIncomingMarker 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   6120
      TabIndex        =   20
      Text            =   "Text1"
      Top             =   5520
      Width           =   1215
   End
   Begin VB.TextBox txtStartMarker1 
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   6120
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2760
      Width           =   1215
   End
   Begin VB.TextBox txtEndMarkerStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   4080
      TabIndex        =   31
      Text            =   "Text1"
      Top             =   7245
      Width           =   975
   End
   Begin VB.TextBox txtDurationLength 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   5160
      TabIndex        =   27
      Text            =   "Text1"
      Top             =   6555
      Width           =   855
   End
   Begin VB.TextBox txtRingDurationLength 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   5160
      TabIndex        =   25
      Text            =   "Text1"
      Top             =   6210
      Width           =   855
   End
   Begin VB.TextBox txtCallerLength 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   5160
      TabIndex        =   22
      Text            =   "Text1"
      Top             =   5865
      Width           =   855
   End
   Begin VB.TextBox txtDialedNumberLength 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   5160
      TabIndex        =   18
      Text            =   "Text1"
      Top             =   5175
      Width           =   855
   End
   Begin VB.TextBox txtCOLineLength 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   5160
      TabIndex        =   14
      Text            =   "Text1"
      Top             =   4485
      Width           =   855
   End
   Begin VB.TextBox txtExtensionLength 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   5160
      TabIndex        =   12
      Text            =   "Text1"
      Top             =   4140
      Width           =   855
   End
   Begin VB.TextBox txtTimeLength 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   5160
      TabIndex        =   9
      Text            =   "Text1"
      Top             =   3795
      Width           =   855
   End
   Begin VB.TextBox txtDateLength 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   5160
      TabIndex        =   6
      Text            =   "Text1"
      Top             =   3450
      Width           =   855
   End
   Begin VB.TextBox txtForwardingFlagStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   4080
      TabIndex        =   15
      Text            =   "Text1"
      Top             =   4830
      Width           =   975
   End
   Begin VB.TextBox txtDurationStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   4080
      TabIndex        =   26
      Text            =   "Text1"
      Top             =   6555
      Width           =   975
   End
   Begin VB.TextBox txtRingDurationStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   4080
      TabIndex        =   24
      Text            =   "Text1"
      Top             =   6210
      Width           =   975
   End
   Begin VB.TextBox txtCallerStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   4080
      TabIndex        =   21
      Text            =   "Text1"
      Top             =   5865
      Width           =   975
   End
   Begin VB.TextBox txtIncomingMarkerStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   4080
      TabIndex        =   19
      Text            =   "Text1"
      Top             =   5520
      Width           =   975
   End
   Begin VB.TextBox txtDialedNumberStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   4080
      TabIndex        =   17
      Text            =   "Text1"
      Top             =   5175
      Width           =   975
   End
   Begin VB.TextBox txtCOLineStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   4080
      TabIndex        =   13
      Text            =   "Text1"
      Top             =   4485
      Width           =   975
   End
   Begin VB.TextBox txtExtensionStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   4080
      TabIndex        =   11
      Text            =   "Text1"
      Top             =   4140
      Width           =   975
   End
   Begin VB.TextBox txtTimeStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   4080
      TabIndex        =   8
      Text            =   "Text1"
      Top             =   3795
      Width           =   975
   End
   Begin VB.TextBox txtDateStart 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   4080
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   3450
      Width           =   975
   End
   Begin VB.TextBox txtStartMarker1Start 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      Height          =   360
      Left            =   4080
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   2760
      Width           =   975
   End
   Begin VB.CommandButton cmdStartMarker1 
      Appearance      =   0  'Flat
      Caption         =   "Start marker 1"
      Height          =   360
      Left            =   720
      TabIndex        =   49
      Top             =   2760
      Visible         =   0   'False
      Width           =   1575
   End
   Begin VB.TextBox txtData 
      Appearance      =   0  'Flat
      BackColor       =   &H00FDEBD2&
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   750
      HideSelection   =   0   'False
      Left            =   2400
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      Top             =   1560
      Width           =   6585
   End
   Begin VB.PictureBox picLeftPicture 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   7080
      Left            =   60
      Picture         =   "frmDataMapper.frx":798B
      ScaleHeight     =   472
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   73
      Top             =   1245
      Width           =   2175
   End
   Begin ARButtonCtrl.ARButton cmdCancel 
      Height          =   375
      Left            =   7800
      TabIndex        =   34
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Cancel"
      ForeColor       =   -2147483630
      ForeColorOnMouse=   12582912
      ForeColorOnFocus=   12582912
      BackColorOnMouse=   16774632
      BorderColor     =   12582912
      BackColor       =   16640978
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
   Begin ARButtonCtrl.ARButton cmdOk 
      Height          =   375
      Left            =   6480
      TabIndex        =   33
      Top             =   7800
      Width           =   1215
      _ExtentX        =   2143
      _ExtentY        =   661
      Caption         =   "Okay"
      ForeColor       =   -2147483630
      ForeColorOnMouse=   12582912
      ForeColorOnFocus=   12582912
      BackColorOnMouse=   16774632
      BorderColor     =   12582912
      BackColor       =   16640978
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
   Begin VB.Label lblOptional 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Optional"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   300
      Left            =   7440
      TabIndex        =   74
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Shape shpFields 
      Height          =   4845
      Left            =   2400
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label lblAccount 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Don't read"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   7680
      TabIndex        =   72
      Top             =   6900
      Width           =   780
   End
   Begin VB.Label lblPasteFromClipboard 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Paste from clipboard"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   240
      Left            =   7440
      MouseIcon       =   "frmDataMapper.frx":DDFA
      MousePointer    =   99  'Custom
      TabIndex        =   65
      Top             =   1320
      Width           =   1545
   End
   Begin VB.Label lblCallerDontRead 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Don't read"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   7680
      TabIndex        =   71
      Top             =   5925
      Width           =   780
   End
   Begin VB.Label lblReadSystemTime 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Read system time"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   7680
      TabIndex        =   54
      Top             =   3825
      Width           =   1305
   End
   Begin VB.Label lblReadSystemDate 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Read system date"
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   7680
      TabIndex        =   52
      Top             =   3480
      Width           =   1320
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Marker"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   300
      Index           =   3
      Left            =   6120
      TabIndex        =   70
      Top             =   2400
      Width           =   1215
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Length"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   300
      Index           =   2
      Left            =   5160
      TabIndex        =   69
      Top             =   2400
      Width           =   855
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Starts at"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   300
      Index           =   1
      Left            =   4080
      TabIndex        =   68
      Top             =   2400
      Width           =   975
   End
   Begin VB.Label lbl 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H00FFC0C0&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Field"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00400000&
      Height          =   300
      Index           =   0
      Left            =   2400
      TabIndex        =   67
      Top             =   2400
      Width           =   1575
   End
   Begin VB.Label lblData 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PBX Data"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2400
      TabIndex        =   66
      Top             =   1320
      Width           =   705
   End
End
Attribute VB_Name = "frmDataMapper"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdAccountXP_Click()
txtAccountStart = txtData.SelStart + 1
txtAccountLength = txtData.SelLength
chkAccount.Value = vbUnchecked
End Sub

Private Sub cmdCallerXP_Click()
txtCallerStart = txtData.SelStart + 1
txtCallerLength = txtData.SelLength
chkCaller.Value = vbUnchecked
End Sub

Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdCOLineXP_Click()
txtCOLineStart = txtData.SelStart + 1
txtCOLineLength = txtData.SelLength
End Sub

Private Sub cmdDateXP_Click()
txtDateStart = txtData.SelStart + 1
txtDateLength = txtData.SelLength
chkDate.Value = vbUnchecked
End Sub

Private Sub cmdDialedNumberXP_Click()
txtDialedNumberStart = txtData.SelStart + 1
txtDialedNumberLength = txtData.SelLength
End Sub

Private Sub cmdDurationXP_Click()
txtDurationStart = txtData.SelStart + 1
txtDurationLength = txtData.SelLength
End Sub

Private Sub cmdEndMarkerXP_Click()
txtEndMarkerStart = txtData.SelStart + 1
txtEndMarker = txtData.SelText
End Sub

Private Sub cmdExtensionXP_Click()
txtExtensionStart = txtData.SelStart + 1
txtExtensionLength = txtData.SelLength
End Sub

Private Sub cmdForwardingFlagXP_Click()
txtForwardingFlagStart = txtData.SelStart + 1
txtForwardingFlag = txtData.SelText
End Sub

Private Sub cmdIncomingMarkerXP_Click()
txtIncomingMarkerStart = txtData.SelStart + 1
txtIncomingMarker = txtData.SelText
End Sub

Private Sub cmdOK_Click()
'Save changed data mapping
DataMap.StartMarker1Start = txtStartMarker1Start
DataMap.StartMarker1 = txtStartMarker1
DataMap.StartMarker2Start = txtStartMarker2Start
DataMap.StartMarker2 = txtStartMarker2
DataMap.DateStart = txtDateStart
DataMap.DateLength = txtDateLength
DataMap.SystemDate = (-1) * chkDate.Value
DataMap.TimeStart = txtTimeStart
DataMap.TimeLength = txtTimeLength
DataMap.SystemTime = (-1) * chkTime.Value
DataMap.ExtensionStart = txtExtensionStart
DataMap.ExtensionLength = txtExtensionLength
DataMap.COLineStart = txtCOLineStart
DataMap.COLineLength = txtCOLineLength
DataMap.ForwardingFlagStart = txtForwardingFlagStart
DataMap.ForwardingFlag = txtForwardingFlag
DataMap.DialedNumberStart = txtDialedNumberStart
DataMap.DialedNumberLength = txtDialedNumberLength
DataMap.IncomingMarkerStart = txtIncomingMarkerStart
DataMap.IncomingMarker = txtIncomingMarker
DataMap.CallerStart = txtCallerStart
DataMap.CallerLength = txtCallerLength
DataMap.CallerDontRead = (-1) * chkCaller.Value
DataMap.RingDurationStart = txtRingDurationStart
DataMap.RingDurationLength = txtRingDurationLength
DataMap.DurationStart = txtDurationStart
DataMap.DurationLength = txtDurationLength
DataMap.AccountStart = txtAccountStart
DataMap.AccountLength = txtAccountLength
DataMap.AccountDontRead = (-1) * chkAccount.Value
DataMap.EndMarkerStart = txtEndMarkerStart
DataMap.EndMarker = txtEndMarker

SaveDataMap

Unload Me
End Sub

Private Sub cmdRingDurationXP_Click()
txtRingDurationStart = txtData.SelStart + 1
txtRingDurationLength = txtData.SelLength
End Sub

Private Sub cmdStartMarker1XP_Click()
txtStartMarker1Start = txtData.SelStart + 1
txtStartMarker1 = txtData.SelText
End Sub

Private Sub cmdStartMarker2XP_Click()
txtStartMarker2Start = txtData.SelStart + 1
txtStartMarker2 = txtData.SelText
End Sub

Private Sub cmdTimeXP_Click()
txtTimeStart = txtData.SelStart + 1
txtTimeLength = txtData.SelLength
chkTime.Value = vbUnchecked
End Sub

Private Sub Form_Load()
'Dump current mapping
txtStartMarker1Start = DataMap.StartMarker1Start
txtStartMarker1 = DataMap.StartMarker1
txtStartMarker2Start = DataMap.StartMarker2Start
txtStartMarker2 = DataMap.StartMarker2
txtDateStart = DataMap.DateStart
txtDateLength = DataMap.DateLength
chkDate.Value = Abs(DataMap.SystemDate)
txtTimeStart = DataMap.TimeStart
txtTimeLength = DataMap.TimeLength
chkTime.Value = Abs(DataMap.SystemTime)
txtExtensionStart = DataMap.ExtensionStart
txtExtensionLength = DataMap.ExtensionLength
txtCOLineStart = DataMap.COLineStart
txtCOLineLength = DataMap.COLineLength
txtForwardingFlagStart = DataMap.ForwardingFlagStart
txtForwardingFlag = DataMap.ForwardingFlag
txtDialedNumberStart = DataMap.DialedNumberStart
txtDialedNumberLength = DataMap.DialedNumberLength
txtIncomingMarkerStart = DataMap.IncomingMarkerStart
txtIncomingMarker = DataMap.IncomingMarker
txtCallerStart = DataMap.CallerStart
txtCallerLength = DataMap.CallerLength
chkCaller.Value = Abs(DataMap.CallerDontRead)
txtRingDurationStart = DataMap.RingDurationStart
txtRingDurationLength = DataMap.RingDurationLength
txtDurationStart = DataMap.DurationStart
txtDurationLength = DataMap.DurationLength
txtAccountStart = DataMap.AccountStart
txtAccountLength = DataMap.AccountLength
chkAccount.Value = Abs(DataMap.AccountDontRead)
txtEndMarkerStart = DataMap.EndMarkerStart
txtEndMarker = DataMap.EndMarker

End Sub

Private Sub lblAccount_Click()
chkAccount = Abs(CInt(Not CBool(chkAccount)))
End Sub

Private Sub lblCallerDontRead_Click()
chkCaller = Abs(CInt(Not CBool(chkCaller)))
End Sub

Private Sub lblPasteFromClipboard_Click()
On Error Resume Next 'Don't crash if the clipboard format is not supported
txtData = Clipboard.GetText
End Sub

Private Sub lblPasteFromClipboard_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then lblPasteFromClipboard.Move lblPasteFromClipboard.Left + 1, lblPasteFromClipboard.Top + 1
End Sub

Private Sub lblPasteFromClipboard_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
If Button = 1 Then lblPasteFromClipboard.Move lblPasteFromClipboard.Left - 1, lblPasteFromClipboard.Top - 1
End Sub

Private Sub lblReadSystemDate_Click()
chkDate = Abs(CInt(Not CBool(chkDate)))
End Sub

Private Sub lblReadSystemTime_Click()
chkTime = Abs(CInt(Not CBool(chkTime)))
End Sub

Private Sub txtAccountLength_GotFocus()
SelectAllText txtAccountLength

End Sub

Private Sub txtAccountStart_GotFocus()
SelectAllText txtAccountStart

End Sub

Private Sub txtCallerLength_GotFocus()
SelectAllText txtCallerLength

End Sub

Private Sub txtCallerStart_GotFocus()
SelectAllText txtCallerStart

End Sub

Private Sub txtCOLineLength_GotFocus()
SelectAllText txtCOLineLength

End Sub

Private Sub txtCOLineStart_GotFocus()
SelectAllText txtCOLineStart

End Sub

Private Sub txtDateLength_GotFocus()
SelectAllText txtDateLength

End Sub

Private Sub txtDateStart_GotFocus()
SelectAllText txtDateStart

End Sub

Private Sub txtDialedNumberLength_GotFocus()
SelectAllText txtDialedNumberLength

End Sub

Private Sub txtDialedNumberStart_GotFocus()
SelectAllText txtDialedNumberStart

End Sub

Private Sub txtDurationLength_GotFocus()
SelectAllText txtDurationLength

End Sub

Private Sub txtDurationStart_GotFocus()
SelectAllText txtDurationStart

End Sub

Private Sub txtEndMarker_GotFocus()
SelectAllText txtEndMarker

End Sub

Private Sub txtEndMarkerStart_GotFocus()
SelectAllText txtEndMarkerStart

End Sub

Private Sub txtExtensionLength_GotFocus()
SelectAllText txtExtensionLength

End Sub

Private Sub txtExtensionStart_GotFocus()
SelectAllText txtExtensionStart

End Sub

Private Sub txtForwardingFlag_GotFocus()
SelectAllText txtForwardingFlag

End Sub

Private Sub txtForwardingFlagStart_GotFocus()
SelectAllText txtForwardingFlagStart

End Sub

Private Sub txtIncomingMarker_GotFocus()
SelectAllText txtIncomingMarker

End Sub

Private Sub txtIncomingMarkerStart_GotFocus()
SelectAllText txtIncomingMarkerStart

End Sub

Private Sub txtRingDurationLength_GotFocus()
SelectAllText txtRingDurationLength

End Sub

Private Sub txtRingDurationStart_GotFocus()
SelectAllText txtRingDurationStart

End Sub

Private Sub txtStartMarker1_GotFocus()
SelectAllText txtStartMarker1
End Sub

Private Sub txtStartMarker1Start_GotFocus()
SelectAllText txtStartMarker1Start
End Sub

Private Sub txtStartMarker2_GotFocus()
SelectAllText txtStartMarker2
End Sub

Private Sub txtStartMarker2Start_GotFocus()
SelectAllText txtStartMarker2Start

End Sub

Private Sub txtTimeLength_GotFocus()
SelectAllText txtTimeLength

End Sub

Private Sub txtTimeStart_GotFocus()
SelectAllText txtTimeStart

End Sub
