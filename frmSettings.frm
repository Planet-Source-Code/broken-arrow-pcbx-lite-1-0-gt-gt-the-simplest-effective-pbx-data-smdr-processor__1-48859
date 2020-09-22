VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.3#0"; "ARButton.ocx"
Object = "{C64D70BC-E172-42ED-B119-C0CBE641CCA0}#1.9#0"; "WOWFormer.ocx"
Begin VB.Form frmSettings 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Settings"
   ClientHeight    =   4290
   ClientLeft      =   4410
   ClientTop       =   4950
   ClientWidth     =   7815
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmSettings.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   286
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   521
   StartUpPosition =   3  'Windows Default
   Begin WOWFormer_ActiveX.WOWFormer WOWFormer 
      Align           =   1  'Align Top
      Height          =   285
      Left            =   0
      TabIndex        =   4
      Top             =   0
      Width           =   7815
      _ExtentX        =   13785
      _ExtentY        =   503
      PictureLeft     =   "frmSettings.frx":030A
      PictureMiddle   =   "frmSettings.frx":0D74
      PictureRight    =   "frmSettings.frx":0E12
      PictureLeftWidth=   49
      PictureRightWidth=   45
      FormBorderTop   =   "frmSettings.frx":0EB0
      FormBorderLeft  =   "frmSettings.frx":0F12
      FormBorderBottom=   "frmSettings.frx":0F70
      FormBorderRight =   "frmSettings.frx":0FD2
      FormBorderLeftWidth=   4
      FormBorderBottomHeight=   4
      FormBorderRightWidth=   4
      FormBackground  =   "frmSettings.frx":1030
      AllowMaximize   =   0   'False
      FormIcon        =   "frmSettings.frx":1C82
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
      PictureMaximize =   "frmSettings.frx":1F9C
      PictureMinimize =   "frmSettings.frx":232E
      PictureClose    =   "frmSettings.frx":26C0
      PictureMinimizeToTray=   "frmSettings.frx":2A52
      ControlBoxSpacing=   3
      ControlBoxRightPadding=   4
      CaptionPrefix   =   "PCBXLite 1.0> "
      PictureShrink   =   "frmSettings.frx":2DE4
      MinimizeToTray  =   0   'False
      PictureCloseDown=   "frmSettings.frx":3176
      PictureMaximizeDown=   "frmSettings.frx":3508
      PictureMinimizeDown=   "frmSettings.frx":389A
      PictureShrinkDown=   "frmSettings.frx":3C2C
      PictureMinimizeToTrayDown=   "frmSettings.frx":3FBE
      SnapTolerance   =   700
      PicturePin      =   "frmSettings.frx":4350
      AllowOnTop      =   0   'False
      PicturePinDown  =   "frmSettings.frx":46E2
      PicturePinHover =   "frmSettings.frx":4A74
      PictureMinimizeToTrayHover=   "frmSettings.frx":4DC6
      PictureShrinkHover=   "frmSettings.frx":5118
      PictureMinimizeHover=   "frmSettings.frx":546A
      PictureMaximizeHover=   "frmSettings.frx":57BC
      PictureCloseHover=   "frmSettings.frx":5B0E
      TrayTip         =   " Settings "
      FormMouseIcon   =   "frmSettings.frx":5E60
      TrayIcon        =   "frmSettings.frx":667A
   End
   Begin MSComDlg.CommonDialog cdlDatabase 
      Left            =   1680
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
      DialogTitle     =   "Select database..."
      Filter          =   "Microsoft Access database|*.mdb|All files|*.*"
   End
   Begin VB.TextBox txtDatabase 
      Height          =   390
      Left            =   1215
      Locked          =   -1  'True
      TabIndex        =   1
      Top             =   480
      Width           =   6375
   End
   Begin ARButtonCtrl.ARButton cmdDatabase 
      Height          =   390
      Left            =   240
      TabIndex        =   0
      Top             =   480
      Width           =   975
      _ExtentX        =   1720
      _ExtentY        =   688
      Caption         =   "Database"
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
   End
   Begin ARButtonCtrl.ARButton cmdCancel 
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   3720
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
      Left            =   5040
      TabIndex        =   3
      Top             =   3720
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
End
Attribute VB_Name = "frmSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdDatabase_Click()
On Error Resume Next

cdlDatabase.ShowOpen
If Err Then Exit Sub 'User chose Cancel

txtDatabase = cdlDatabase.FileName
End Sub

Private Sub cmdOK_Click()
MDBDatabase = txtDatabase
GlobalADOConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & MDBDatabase
frmMain.datCall.ConnectionString = GlobalADOConnectionString

frmMain.ReloadData

SaveAppSetting

Unload Me
End Sub

Private Sub Form_Load()
txtDatabase = MDBDatabase
End Sub
