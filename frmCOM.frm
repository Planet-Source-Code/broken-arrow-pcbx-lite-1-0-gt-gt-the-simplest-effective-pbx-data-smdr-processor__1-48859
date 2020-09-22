VERSION 5.00
Object = "{C30897B9-75AC-11D2-94E3-000000000000}#1.3#0"; "ARButton.ocx"
Object = "{C64D70BC-E172-42ED-B119-C0CBE641CCA0}#1.9#0"; "WOWFormer.ocx"
Begin VB.Form frmCOM 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "COM Port Wizard"
   ClientHeight    =   5115
   ClientLeft      =   3465
   ClientTop       =   3285
   ClientWidth     =   8370
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmCOM.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   341
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   558
   StartUpPosition =   3  'Windows Default
   Begin WOWFormer_ActiveX.WOWFormer WOWFormer 
      Align           =   1  'Align Top
      Height          =   285
      Left            =   0
      TabIndex        =   32
      Top             =   0
      Width           =   8370
      _ExtentX        =   14764
      _ExtentY        =   503
      PictureLeft     =   "frmCOM.frx":030A
      PictureMiddle   =   "frmCOM.frx":0D74
      PictureRight    =   "frmCOM.frx":0E12
      PictureLeftWidth=   49
      PictureRightWidth=   47
      FormBorderTop   =   "frmCOM.frx":0EB0
      FormBorderLeft  =   "frmCOM.frx":0F12
      FormBorderBottom=   "frmCOM.frx":0F70
      FormBorderRight =   "frmCOM.frx":0FD2
      FormBorderLeftWidth=   4
      FormBorderBottomHeight=   4
      FormBorderRightWidth=   4
      FormBackground  =   "frmCOM.frx":1030
      AllowMaximize   =   0   'False
      FormIcon        =   "frmCOM.frx":1C82
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
      PictureMaximize =   "frmCOM.frx":1F9C
      PictureMinimize =   "frmCOM.frx":232E
      PictureClose    =   "frmCOM.frx":26C0
      PictureMinimizeToTray=   "frmCOM.frx":2A52
      ControlBoxSpacing=   3
      PictureShrink   =   "frmCOM.frx":2DE4
      MinimizeToTray  =   0   'False
      PictureCloseDown=   "frmCOM.frx":3176
      PictureMaximizeDown=   "frmCOM.frx":3508
      PictureMinimizeDown=   "frmCOM.frx":389A
      PictureShrinkDown=   "frmCOM.frx":3C2C
      PictureMinimizeToTrayDown=   "frmCOM.frx":3FBE
      SnapTolerance   =   700
      PicturePin      =   "frmCOM.frx":4350
      AllowOnTop      =   0   'False
      PicturePinDown  =   "frmCOM.frx":46E2
      PicturePinHover =   "frmCOM.frx":4A74
      PictureMinimizeToTrayHover=   "frmCOM.frx":4DC6
      PictureShrinkHover=   "frmCOM.frx":5118
      PictureMinimizeHover=   "frmCOM.frx":546A
      PictureMaximizeHover=   "frmCOM.frx":57BC
      PictureCloseHover=   "frmCOM.frx":5B0E
      TrayTip         =   " COM Port Wizard "
      FormMouseIcon   =   "frmCOM.frx":5E60
      TrayIcon        =   "frmCOM.frx":667A
   End
   Begin ARButtonCtrl.ARButton cmdCancel 
      Height          =   375
      Left            =   6840
      TabIndex        =   31
      Top             =   4560
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
      Left            =   5520
      TabIndex        =   30
      Top             =   4560
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
   Begin VB.PictureBox picHandShaking 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3840
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   279
      TabIndex        =   25
      Top             =   4080
      Width           =   4215
      Begin VB.OptionButton optHandShaking 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "RTS && X On X Off"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   2580
         TabIndex        =   29
         Top             =   60
         Width           =   1695
      End
      Begin VB.OptionButton optHandShaking 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "RTS"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   1980
         TabIndex        =   28
         Top             =   60
         Width           =   615
      End
      Begin VB.OptionButton optHandShaking 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "X On X Off"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   780
         TabIndex        =   27
         Top             =   60
         Width           =   1215
      End
      Begin VB.OptionButton optHandShaking 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "None"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   26
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.PictureBox picParity 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3840
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   247
      TabIndex        =   19
      Top             =   3555
      Width           =   3735
      Begin VB.OptionButton optParity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Space"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   2940
         TabIndex        =   24
         Top             =   60
         Width           =   855
      End
      Begin VB.OptionButton optParity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Odd"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   2220
         TabIndex        =   23
         Top             =   60
         Width           =   735
      End
      Begin VB.OptionButton optParity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "None"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   1500
         TabIndex        =   22
         Top             =   60
         Width           =   735
      End
      Begin VB.OptionButton optParity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Mark"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   780
         TabIndex        =   21
         Top             =   60
         Width           =   735
      End
      Begin VB.OptionButton optParity 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Even"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   20
         Top             =   60
         Width           =   735
      End
   End
   Begin VB.PictureBox picStopBit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3840
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   151
      TabIndex        =   15
      Top             =   3075
      Width           =   2295
      Begin VB.OptionButton optStopBit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "2 Bit"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   1620
         TabIndex        =   18
         Top             =   60
         Width           =   735
      End
      Begin VB.OptionButton optStopBit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "1.5 Bit"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   780
         TabIndex        =   17
         Top             =   60
         Width           =   855
      End
      Begin VB.OptionButton optStopBit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "1 Bit"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   16
         Top             =   60
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.PictureBox picDataBit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   360
      Left            =   3840
      ScaleHeight     =   22
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   271
      TabIndex        =   9
      Top             =   2595
      Width           =   4095
      Begin VB.OptionButton optDataBit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "4 Bit"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   0
         Left            =   60
         TabIndex        =   14
         Top             =   60
         Width           =   735
      End
      Begin VB.OptionButton optDataBit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "5 Bit"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   1
         Left            =   900
         TabIndex        =   13
         Top             =   60
         Width           =   735
      End
      Begin VB.OptionButton optDataBit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "6 Bit"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   2
         Left            =   1740
         TabIndex        =   12
         Top             =   60
         Width           =   735
      End
      Begin VB.OptionButton optDataBit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "7 Bit"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   3
         Left            =   2580
         TabIndex        =   11
         Top             =   60
         Width           =   735
      End
      Begin VB.OptionButton optDataBit 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "8 Bit"
         ForeColor       =   &H80000008&
         Height          =   240
         Index           =   4
         Left            =   3420
         TabIndex        =   10
         Top             =   60
         Value           =   -1  'True
         Width           =   735
      End
   End
   Begin VB.ComboBox cboBaud 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmCOM.frx":6994
      Left            =   3840
      List            =   "frmCOM.frx":69C8
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   2100
      Width           =   1455
   End
   Begin VB.ComboBox cboPort 
      Appearance      =   0  'Flat
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   390
      ItemData        =   "frmCOM.frx":6A35
      Left            =   3840
      List            =   "frmCOM.frx":6A37
      TabIndex        =   2
      Text            =   "Combo1"
      Top             =   1620
      Width           =   975
   End
   Begin PCBXLite.WizardTitle WizardTitle 
      Height          =   1215
      Left            =   60
      Top             =   285
      Width           =   8250
      _ExtentX        =   14552
      _ExtentY        =   2143
      Title           =   " COM Port Wizard"
      Description     =   $"frmCOM.frx":6A39
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
   Begin VB.PictureBox picLeftPicture 
      Appearance      =   0  'Flat
      BackColor       =   &H00800000&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3570
      Left            =   60
      Picture         =   "frmCOM.frx":6B6E
      ScaleHeight     =   238
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   145
      TabIndex        =   0
      Top             =   1485
      Width           =   2175
   End
   Begin VB.Label lblHandshaking 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Handshake type"
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
      Height          =   270
      Left            =   2400
      TabIndex        =   8
      Top             =   4080
      Width           =   1350
   End
   Begin VB.Label lblParity 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Parity"
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
      Height          =   270
      Left            =   2400
      TabIndex        =   7
      Top             =   3600
      Width           =   510
   End
   Begin VB.Label lblStopBit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Stop bit"
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
      Height          =   270
      Left            =   2400
      TabIndex        =   6
      Top             =   3120
      Width           =   675
   End
   Begin VB.Label lblDataBit 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Data bit"
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
      Height          =   270
      Left            =   2400
      TabIndex        =   5
      Top             =   2640
      Width           =   675
   End
   Begin VB.Label lblBaud 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Baud rate"
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
      Height          =   270
      Left            =   2400
      TabIndex        =   3
      Top             =   2160
      Width           =   810
   End
   Begin VB.Label lblPort 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "COM Port"
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
      Height          =   270
      Left            =   2400
      TabIndex        =   1
      Top             =   1680
      Width           =   795
   End
End
Attribute VB_Name = "frmCOM"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim DataBit As Long, StopBit As Single, Parity As String, HandShaking As Long

Private Sub cboDataBit_Change()

End Sub

Private Sub cmdCancel_Click()
Unload Me

End Sub

Private Sub cmdOK_Click()
On Error GoTo Error_cmdOK_Click

'Save changed COM setting
COM.Port = cboPort
COM.BaudRate = cboBaud
COM.DataBit = DataBit
COM.StopBit = StopBit
COM.Parity = Parity
COM.HandShaking = HandShaking

SaveCOMSetting

With frmMain.comPBX
    .PortOpen = False
    .CommPort = COM.Port
    .Settings = COM.BaudRate & "," & Left(COM.Parity, 1) & "," & COM.DataBit & "," & COM.StopBit
    .HandShaking = COM.HandShaking
    .PortOpen = True
End With

Unload Me

Exit Sub

Error_cmdOK_Click:
Select Case Err.Number
Case Else
    MsgBox "Error#" & Err.Number & ": " & Err.Description, vbCritical + vbOKOnly, "PCBX Lite"
    Resume Next
End Select

End Sub

Private Sub Form_Load()
'Dump current COM setting
cboPort = COM.Port
cboBaud = COM.BaudRate
optDataBit(COM.DataBit - 4).Value = True

Select Case COM.StopBit
Case 1
    optStopBit(0).Value = True
Case 1.5
    optStopBit(1).Value = True
Case 2
    optStopBit(2).Value = True
End Select

Select Case Left(COM.Parity, 1)
Case "E"
    optParity(0).Value = True
Case "M"
    optParity(1).Value = True
Case "N"
    optParity(2).Value = True
Case "O"
    optParity(3).Value = True
Case "S"
    optParity(4).Value = True
End Select

optHandShaking(COM.HandShaking).Value = True

DataBit = COM.DataBit
StopBit = COM.StopBit
Parity = COM.Parity
HandShaking = COM.HandShaking
End Sub

Private Sub optDataBit_Click(Index As Integer)
DataBit = Index + 4
End Sub

Private Sub optHandShaking_Click(Index As Integer)
HandShaking = Index
End Sub

Private Sub optParity_Click(Index As Integer)
Parity = Left(optParity(Index).Caption, 1)
End Sub

Private Sub optStopBit_Click(Index As Integer)
If Index = 0 Then StopBit = 1
If Index = 1 Then StopBit = 1.5
If Index = 2 Then StopBit = 2
End Sub
