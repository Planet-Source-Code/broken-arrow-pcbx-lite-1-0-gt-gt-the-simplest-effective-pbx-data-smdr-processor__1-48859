VERSION 5.00
Object = "{C64D70BC-E172-42ED-B119-C0CBE641CCA0}#1.9#0"; "WOWFormer.ocx"
Begin VB.Form frmAbout 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "About PCBX Lite..."
   ClientHeight    =   4230
   ClientLeft      =   6300
   ClientTop       =   2250
   ClientWidth     =   6225
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmAbout.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   282
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   415
   StartUpPosition =   2  'CenterScreen
   Begin WOWFormer_ActiveX.WOWFormer WOWFormer 
      Align           =   1  'Align Top
      Height          =   285
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   6225
      _ExtentX        =   10980
      _ExtentY        =   503
      PictureLeft     =   "frmAbout.frx":030A
      PictureMiddle   =   "frmAbout.frx":0D74
      PictureRight    =   "frmAbout.frx":0E12
      PictureLeftWidth=   49
      PictureRightWidth=   25
      FormBorderTop   =   "frmAbout.frx":0EB0
      FormBorderLeft  =   "frmAbout.frx":0F12
      FormBorderBottom=   "frmAbout.frx":0F70
      FormBorderRight =   "frmAbout.frx":0FD2
      FormBorderLeftWidth=   4
      FormBorderBottomHeight=   4
      FormBorderRightWidth=   4
      FormBackground  =   "frmAbout.frx":1030
      AllowMaximize   =   0   'False
      FormIcon        =   "frmAbout.frx":1C82
      AllowMinimize   =   0   'False
      CaptionSpacing  =   0
      CaptionAlign    =   2
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
      PictureMaximize =   "frmAbout.frx":1F9C
      PictureMinimize =   "frmAbout.frx":232E
      PictureClose    =   "frmAbout.frx":26C0
      PictureMinimizeToTray=   "frmAbout.frx":2A52
      ControlBoxSpacing=   3
      ControlBoxRightPadding=   4
      PictureShrink   =   "frmAbout.frx":2DE4
      AllowShrink     =   0   'False
      MinimizeToTray  =   0   'False
      PictureCloseDown=   "frmAbout.frx":3176
      PictureMaximizeDown=   "frmAbout.frx":3508
      PictureMinimizeDown=   "frmAbout.frx":389A
      PictureShrinkDown=   "frmAbout.frx":3C2C
      PictureMinimizeToTrayDown=   "frmAbout.frx":3FBE
      AutoCloseInterval=   20
      SnapTolerance   =   700
      ControlMenu     =   0   'False
      AlwaysOnTop     =   -1  'True
      Transparency    =   127
      PicturePin      =   "frmAbout.frx":4350
      AllowOnTop      =   0   'False
      PicturePinDown  =   "frmAbout.frx":46E2
      PicturePinHover =   "frmAbout.frx":4A74
      PictureMinimizeToTrayHover=   "frmAbout.frx":4DC6
      PictureShrinkHover=   "frmAbout.frx":5118
      PictureMinimizeHover=   "frmAbout.frx":546A
      PictureMaximizeHover=   "frmAbout.frx":57BC
      PictureCloseHover=   "frmAbout.frx":5B0E
      TrayTip         =   " About PCBX Lite... "
      FormMouseIcon   =   "frmAbout.frx":5E60
      TrayIcon        =   "frmAbout.frx":667A
   End
   Begin PCBXLite.ShadowLabel lblVersion 
      Height          =   495
      Left            =   5040
      Top             =   840
      Width           =   615
      _ExtentX        =   1085
      _ExtentY        =   873
      Title           =   "1.0"
      TitleColor      =   16777215
      ShadowColor     =   0
      BackColor       =   8388608
      ShadowDistance  =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PCBXLite.ShadowLabel lblTitle 
      Height          =   1095
      Left            =   240
      Top             =   240
      Width           =   4815
      _ExtentX        =   8493
      _ExtentY        =   1931
      Title           =   "PCBX Lite"
      TitleColor      =   16761024
      ShadowColor     =   0
      BackColor       =   8388608
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PCBXLite.ShadowLabel lblCopyright 
      Height          =   1245
      Left            =   60
      Top             =   2910
      Width           =   6105
      _ExtentX        =   10769
      _ExtentY        =   2196
      Title           =   $"frmAbout.frx":6994
      TitleColor      =   0
      ShadowColor     =   16777215
      BackColor       =   16512
      ShadowDistance  =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      Padding         =   5
   End
   Begin VB.Label lblInternalVersion 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "##.##.##"
      Height          =   270
      Left            =   5280
      TabIndex        =   0
      Top             =   360
      Width           =   750
   End
   Begin VB.Image imgJoySoftwares 
      Height          =   930
      Left            =   1980
      MouseIcon       =   "frmAbout.frx":6A5E
      MousePointer    =   99  'Custom
      Picture         =   "frmAbout.frx":7328
      Top             =   1380
      Width           =   4185
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
lblInternalVersion = App.Major & "." & App.Minor & "." & App.Revision



End Sub

Private Sub imgJoySoftwares_Click()
Shell "Explorer " & CheckPath(App.Path, True) & "About.htm"
End Sub
