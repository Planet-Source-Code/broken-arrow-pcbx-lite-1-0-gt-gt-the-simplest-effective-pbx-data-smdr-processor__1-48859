VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{C64D70BC-E172-42ED-B119-C0CBE641CCA0}#1.9#0"; "WOWFormer.ocx"
Begin VB.Form frmClientList 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Client Management"
   ClientHeight    =   6390
   ClientLeft      =   2700
   ClientTop       =   3150
   ClientWidth     =   10725
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmClientList.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   426
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   715
   StartUpPosition =   3  'Windows Default
   Begin WOWFormer_ActiveX.WOWFormer WOWFormer 
      Align           =   1  'Align Top
      Height          =   285
      Left            =   0
      TabIndex        =   47
      Top             =   0
      Width           =   10725
      _ExtentX        =   18918
      _ExtentY        =   503
      PictureLeft     =   "frmClientList.frx":030A
      PictureMiddle   =   "frmClientList.frx":0D74
      PictureRight    =   "frmClientList.frx":0E12
      PictureLeftWidth=   49
      PictureRightWidth=   65
      FormBorderTop   =   "frmClientList.frx":0EB0
      FormBorderLeft  =   "frmClientList.frx":0F12
      FormBorderBottom=   "frmClientList.frx":0F70
      FormBorderRight =   "frmClientList.frx":0FD2
      FormBorderLeftWidth=   4
      FormBorderBottomHeight=   4
      FormBorderRightWidth=   4
      FormBackground  =   "frmClientList.frx":1030
      AllowMaximize   =   0   'False
      FormIcon        =   "frmClientList.frx":1C82
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
      PictureMaximize =   "frmClientList.frx":1F9C
      PictureMinimize =   "frmClientList.frx":232E
      PictureClose    =   "frmClientList.frx":26C0
      PictureMinimizeToTray=   "frmClientList.frx":2A52
      ControlBoxSpacing=   3
      ControlBoxRightPadding=   4
      CaptionPrefix   =   "PCBX Lite 1.0> "
      PictureShrink   =   "frmClientList.frx":2DE4
      MinimizeToTray  =   0   'False
      PictureCloseDown=   "frmClientList.frx":3176
      PictureMaximizeDown=   "frmClientList.frx":3508
      PictureMinimizeDown=   "frmClientList.frx":389A
      PictureShrinkDown=   "frmClientList.frx":3C2C
      PictureMinimizeToTrayDown=   "frmClientList.frx":3FBE
      SnapTolerance   =   700
      PicturePin      =   "frmClientList.frx":4350
      AllowOnTop      =   0   'False
      PicturePinDown  =   "frmClientList.frx":46E2
      PicturePinHover =   "frmClientList.frx":4A74
      PictureMinimizeToTrayHover=   "frmClientList.frx":4DC6
      PictureShrinkHover=   "frmClientList.frx":5118
      PictureMinimizeHover=   "frmClientList.frx":546A
      PictureMaximizeHover=   "frmClientList.frx":57BC
      PictureCloseHover=   "frmClientList.frx":5B0E
      TrayTip         =   " Client Management "
      FormMouseIcon   =   "frmClientList.frx":5E60
      TrayIcon        =   "frmClientList.frx":667A
   End
   Begin PCBXLite.ShadowLabel lblClientTree 
      Height          =   495
      Left            =   120
      Top             =   360
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   873
      Title           =   "Client Listing"
      TitleColor      =   16761024
      BackColor       =   8388608
      ShadowDistance  =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackStyle       =   1
      Alignment       =   2
   End
   Begin VB.CommandButton cmdClientRemove 
      Appearance      =   0  'Flat
      Caption         =   "Remove"
      Height          =   375
      Left            =   9360
      TabIndex        =   26
      Top             =   5910
      Width           =   1215
   End
   Begin VB.CommandButton cmdClientUpdate 
      Appearance      =   0  'Flat
      Caption         =   "Update"
      Height          =   375
      Left            =   8160
      TabIndex        =   25
      Top             =   5910
      Width           =   1215
   End
   Begin VB.CommandButton cmdClientAddNew 
      Appearance      =   0  'Flat
      Caption         =   "Add new"
      Height          =   375
      Left            =   6960
      TabIndex        =   24
      Top             =   5910
      Width           =   1215
   End
   Begin PCBXLite.ShadowLabel lblClient 
      Height          =   495
      Left            =   3360
      Top             =   2640
      Width           =   1455
      _ExtentX        =   2566
      _ExtentY        =   873
      Title           =   "Client"
      TitleColor      =   16761024
      ShadowDistance  =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin PCBXLite.ShadowLabel lblClientGroup 
      Height          =   615
      Left            =   3360
      Top             =   240
      Width           =   3015
      _ExtentX        =   5318
      _ExtentY        =   1085
      Title           =   "Client Group"
      TitleColor      =   16761024
      ShadowDistance  =   3
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Arial Black"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox txtNote 
      Appearance      =   0  'Flat
      Height          =   630
      Left            =   3960
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   23
      Top             =   5235
      Width           =   6615
   End
   Begin VB.TextBox txtEmail 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   8760
      TabIndex        =   22
      Top             =   4860
      Width           =   1815
   End
   Begin VB.TextBox txtCountry 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   6480
      TabIndex        =   21
      Top             =   4860
      Width           =   1455
   End
   Begin VB.TextBox txtState 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   3960
      TabIndex        =   20
      Top             =   4860
      Width           =   1575
   End
   Begin VB.TextBox txtZIP 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   9600
      TabIndex        =   19
      Top             =   4485
      Width           =   975
   End
   Begin VB.TextBox txtCity 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   6120
      TabIndex        =   18
      Top             =   4485
      Width           =   1815
   End
   Begin VB.TextBox txtStreet 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   3960
      TabIndex        =   17
      Top             =   4485
      Width           =   1575
   End
   Begin VB.TextBox txtDeposit 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   9600
      TabIndex        =   16
      Top             =   4110
      Width           =   975
   End
   Begin VB.TextBox txtOtherCharge 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   7200
      TabIndex        =   15
      Top             =   4110
      Width           =   975
   End
   Begin VB.TextBox txtArear 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   4320
      TabIndex        =   14
      Top             =   4110
      Width           =   975
   End
   Begin VB.TextBox txtExtension 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   9600
      TabIndex        =   13
      Top             =   3735
      Width           =   975
   End
   Begin VB.ComboBox cboGroupName 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   4320
      Sorted          =   -1  'True
      Style           =   2  'Dropdown List
      TabIndex        =   12
      Top             =   3735
      Width           =   3855
   End
   Begin VB.TextBox txtLastName 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   9360
      TabIndex        =   11
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtMiddleName 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   6960
      TabIndex        =   10
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtFirstName 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   4320
      TabIndex        =   9
      Top             =   3360
      Width           =   1215
   End
   Begin VB.TextBox txtTAX 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   5160
      TabIndex        =   3
      Text            =   "Text1"
      Top             =   1710
      Width           =   975
   End
   Begin VB.CommandButton cmdGroupRemove 
      Appearance      =   0  'Flat
      Caption         =   "Remove"
      Height          =   375
      Left            =   9360
      TabIndex        =   8
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdGroupUpdate 
      Appearance      =   0  'Flat
      Caption         =   "Update"
      Height          =   375
      Left            =   8160
      TabIndex        =   7
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton cmdGroupAddNew 
      Appearance      =   0  'Flat
      Caption         =   "Add new"
      Height          =   375
      Left            =   6960
      TabIndex        =   6
      Top             =   2520
      Width           =   1215
   End
   Begin VB.TextBox txtDescription 
      Appearance      =   0  'Flat
      Height          =   795
      Left            =   6360
      MaxLength       =   200
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   5
      Text            =   "frmClientList.frx":6994
      Top             =   1680
      Width           =   4215
   End
   Begin VB.TextBox txtLateFeePercent 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   5160
      TabIndex        =   4
      Text            =   "Text3"
      Top             =   2085
      Width           =   975
   End
   Begin VB.TextBox txtLineRent 
      Alignment       =   1  'Right Justify
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   5160
      TabIndex        =   2
      Text            =   "Text2"
      Top             =   1335
      Width           =   975
   End
   Begin VB.TextBox txtClientGroupName 
      Appearance      =   0  'Flat
      Height          =   390
      Left            =   5160
      MaxLength       =   20
      TabIndex        =   1
      Text            =   "Text1"
      Top             =   960
      Width           =   5415
   End
   Begin MSComctlLib.ImageList ilstClientTree 
      Left            =   1440
      Top             =   3960
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientList.frx":699A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmClientList.frx":7274
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwClient 
      Height          =   5415
      Left            =   120
      TabIndex        =   0
      Top             =   855
      Width           =   3135
      _ExtentX        =   5530
      _ExtentY        =   9551
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   529
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      HotTracking     =   -1  'True
      ImageList       =   "ilstClientTree"
      BorderStyle     =   1
      Appearance      =   0
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
   Begin VB.Label Label22 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Note"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3360
      TabIndex        =   46
      Top             =   5295
      Width           =   405
   End
   Begin VB.Label Label21 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Email"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   8160
      TabIndex        =   45
      Top             =   4920
      Width           =   465
   End
   Begin VB.Label Label20 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Country"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5640
      TabIndex        =   44
      Top             =   4920
      Width           =   675
   End
   Begin VB.Label Label19 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "State"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3360
      TabIndex        =   43
      Top             =   4920
      Width           =   435
   End
   Begin VB.Label Label18 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "ZIP/Postal code"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   8160
      TabIndex        =   42
      Top             =   4545
      Width           =   1305
   End
   Begin VB.Label Label17 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "City"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5640
      TabIndex        =   41
      Top             =   4545
      Width           =   345
   End
   Begin VB.Label Label16 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Street"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3360
      TabIndex        =   40
      Top             =   4545
      Width           =   525
   End
   Begin VB.Label Label15 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Deposit"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   8640
      TabIndex        =   39
      Top             =   4170
      Width           =   645
   End
   Begin VB.Label Label14 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Other charge"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   6000
      TabIndex        =   38
      Top             =   4170
      Width           =   1095
   End
   Begin VB.Label Label13 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Arear"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3360
      TabIndex        =   37
      Top             =   4170
      Width           =   465
   End
   Begin VB.Label Label12 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Extension"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   8640
      TabIndex        =   36
      Top             =   3795
      Width           =   840
   End
   Begin VB.Label Label11 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Group"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3360
      TabIndex        =   35
      Top             =   3795
      Width           =   510
   End
   Begin VB.Label Label10 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Last name"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   8400
      TabIndex        =   34
      Top             =   3420
      Width           =   855
   End
   Begin VB.Label Label9 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Middle name"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   5760
      TabIndex        =   33
      Top             =   3420
      Width           =   1095
   End
   Begin VB.Label Label8 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "First name"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3360
      TabIndex        =   32
      Top             =   3420
      Width           =   900
   End
   Begin VB.Line Line4 
      BorderColor     =   &H00FFFFFF&
      X1              =   224
      X2              =   707
      Y1              =   217
      Y2              =   217
   End
   Begin VB.Line Line3 
      X1              =   224
      X2              =   707
      Y1              =   216
      Y2              =   216
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Description"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   6360
      TabIndex        =   31
      Top             =   1380
      Width           =   975
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Late fee (% of bill)"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3360
      TabIndex        =   30
      Top             =   2145
      Width           =   1590
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TAX"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3360
      TabIndex        =   29
      Top             =   1770
      Width           =   330
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Line rent"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3360
      TabIndex        =   28
      Top             =   1395
      Width           =   795
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Group name"
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3360
      TabIndex        =   27
      Top             =   960
      Width           =   1020
   End
   Begin VB.Line Line2 
      BorderColor     =   &H00FFFFFF&
      X1              =   224
      X2              =   707
      Y1              =   57
      Y2              =   57
   End
   Begin VB.Line Line1 
      X1              =   224
      X2              =   707
      Y1              =   56
      Y2              =   56
   End
End
Attribute VB_Name = "frmClientList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim CurrentClientID As Long

Private Sub LoadClientTree(Optional SelectedNodeKey As String)
Dim GroupNode As Node, ClientNode As Node
Dim TotalGroup As Long, TotalClient As Long

Dim a As Long
For a = tvwClient.Nodes.Count To 1 Step -1
    tvwClient.Nodes.Remove (a)
Next

'Load the client groups
With Query("SELECT * FROM tblClientGroup")
    While Not .EOF
        Set GroupNode = tvwClient.Nodes.Add(, , "Group#" & .Fields("ClientGroupID"), .Fields("ClientGroupName"), 1)
        TotalGroup = TotalGroup + 1
        GroupNode.Expanded = True
        
        'Load the clients
        With Query("SELECT * FROM tblClient WHERE ClientGroupID= " & .Fields("ClientGroupID"))
            TotalClient = 0
            While Not .EOF
                Set ClientNode = tvwClient.Nodes.Add(GroupNode.Key, tvwChild, "Client#" & .Fields("ClientID"), .Fields("Extension") & "- " & .Fields("LastName") & ", " & .Fields("FirstName"), 2)
                TotalClient = TotalClient + 1
                .MoveNext
                DoEvents
            Wend
        End With
        
        cboGroupName.AddItem GroupNode.Text
        GroupNode.Text = GroupNode.Text & " (" & TotalClient & ")"
        .MoveNext
        DoEvents
    Wend
End With

If TotalGroup > 0 Then tvwClient.Nodes(1).Selected = True
tvwClient_Click
If SelectedNodeKey <> "" Then tvwClient.Nodes(SelectedNodeKey).Selected = True
End Sub

Private Sub cmdClientAddNew_Click()
'Check for invalid client entry!
If Trim(txtFirstName) = "" Or Trim(txtLastName) = "" Or Val(txtExtension) = 0 Then
    MsgBox "Sorry, sufficient user input is not provided. You must provide valid firstname, lastname & extension for the client.", vbCritical + vbOKOnly, "Invalid client!"
    Exit Sub
End If

'Check for duplicate extension!
If Query("SELECT * FROM tblClient WHERE Extension= " & Val(txtExtension)).RecordCount > 0 Then
    MsgBox "A client with the specified extension already exists, please provide another extension and try again.", vbCritical + vbOKOnly, "Client already exists!"
    Exit Sub
End If

'Check if the names are identical to any other user!
If Query("SELECT * FROM tblClient WHERE LCASE(TRIM(FirstName))= '" & LCase(Trim(txtFirstName)) & "' AND LCASE(TRIM(LastName))= '" & LCase(Trim(txtLastName)) & "'").RecordCount > 0 Then
    If MsgBox("At least one client found with the same name, are you sure you want to insert the client?", vbYesNo + vbQuestion + vbDefaultButton2, "Duplicate client?") = vbNo Then Exit Sub
End If

'Insert the client
QueryExec "INSERT INTO tblClient (FirstName, MiddleName, LastName, ClientGroupID, Extension, Arear, OtherCharge, Deposit, Street, City, ZIP, State, Country, Email, [Note]) VALUES ('" & txtFirstName & "', '" & txtMiddleName & "', '" & txtLastName & "', " & Query("SELECT * FROM tblClientGroup WHERE ClientGroupName= '" & cboGroupName.Text & "'").Fields("ClientGroupID") & ", " & Val(txtExtension) & ", " & Val(txtArear) & ", " & Val(txtOtherCharge) & ", " & Val(txtDeposit) & ", '" & txtStreet & "', '" & txtCity & "', '" & txtZIP & "', '" & txtState & "', '" & txtCountry & "', '" & txtEmail & "', '" & txtNote & "')"
tvwClient.Nodes.Add "Group#" & Query("SELECT * FROM tblClientGroup WHERE ClientGroupName= '" & cboGroupName.Text & "'").Fields("ClientGroupID"), tvwChild, "Client#" & Query("SELECT * FROM tblClient WHERE Extension= " & Val(txtExtension)).Fields("ClientID"), Val(txtExtension) & "- " & txtLastName & ", " & txtFirstName, 2
End Sub

Private Sub cmdClientRemove_Click()
'Check if there is any call data logged for the client
MsgBox "Please insert the check to prevent the user from removing the client if the client has any call record."

If MsgBox("Are you sure you want to remove the client?", vbQuestion + vbYesNo + vbDefaultButton2, "Remove client?") = vbNo Then Exit Sub

QueryExec "DELETE FROM tblClient WHERE ClientID= " & txtExtension.Tag
tvwClient.Nodes.Remove (tvwClient.SelectedItem.Key)
End Sub

Private Sub cmdClientUpdate_Click()
'Check if a client is selected on the client tree!
If Val(txtExtension) = 0 Then
    MsgBox "Please select a client from the client tree on the left to update the client information.", vbCritical + vbOKOnly, "No client selected!"
    Exit Sub
End If

'Check for invalid client entry!
If Trim(txtFirstName) = "" Or Trim(txtLastName) = "" Or Val(txtExtension) = 0 Then
    MsgBox "Sorry, sufficient user input is not provided. You must provide valid firstname, lastname & extension for the client.", vbCritical + vbOKOnly, "Invalid client!"
    Exit Sub
End If

'Check for duplicate extension!
If Query("SELECT * FROM tblClient WHERE Extension= " & Val(txtExtension) & " AND ClientID <> " & txtExtension.Tag).RecordCount > 0 Then
    MsgBox "A client with the specified extension already exists, please provide another extension and try again.", vbCritical + vbOKOnly, "Client already exists!"
    Exit Sub
End If

'Check if the names are identical to any other user!
If Query("SELECT * FROM tblClient WHERE LCASE(TRIM(FirstName))= '" & LCase(Trim(txtFirstName)) & "' AND LCASE(TRIM(LastName))= '" & LCase(Trim(txtLastName)) & "'" & " AND ClientID <> " & txtExtension.Tag).RecordCount > 0 Then
    If MsgBox("At least one client found with the same name, are you sure you want to insert the client?", vbYesNo + vbQuestion + vbDefaultButton2, "Duplicate client?") = vbNo Then Exit Sub
End If

'Update the client
QueryExec "UPDATE tblClient SET FirstName='" & txtFirstName & "', MiddleName= '" & txtMiddleName & "', LastName= '" & txtLastName & "', ClientGroupID= " & Query("SELECT * FROM tblClientGroup WHERE ClientGroupName= '" & cboGroupName.Text & "'").Fields("ClientGroupID") & ", Extension= " & txtExtension & ", Arear= " & txtArear & ", OtherCharge= " & txtOtherCharge & ", Deposit= " & txtDeposit & ", Street= '" & txtStreet & "', City= '" & txtCity & "', ZIP= '" & txtZIP & "', State= '" & txtState & "', Country= '" & txtCountry & "', Email= '" & txtEmail & "', [Note]= '" & txtNote & "' WHERE ClientID= " & txtExtension.Tag
tvwClient.SelectedItem.Text = Val(txtExtension) & "- " & txtLastName & ", " & txtFirstName
End Sub

Private Sub cmdGroupAddNew_Click()
'Check for null group name!
If Trim(txtClientGroupName) = "" Then
    MsgBox "Sorry, the group name you provided is invalid! Please try with another group name.", vbCritical + vbOKOnly, "Invalid group name!"
    Exit Sub
End If
'Check for duplicate group name
If Query("SELECT * FROM tblClientGroup WHERE ClientGroupName = '" & Trim(txtClientGroupName) & "'").RecordCount > 0 Then
    MsgBox "Sorry, you provided a duplicate group name! Please try with another group name.", vbCritical + vbOKOnly, "Duplicate group name!"
    Exit Sub
End If

QueryExec "INSERT INTO tblClientGroup (ClientGroupName, LineRent, TAX, LateFeePercent, Description) VALUES ( '" & txtClientGroupName & "', " & txtLineRent & ", " & txtTAX & ", " & txtLateFeePercent & ", '" & txtDescription & "')"

Dim GroupNode As Node
Set GroupNode = tvwClient.Nodes.Add(, , "Group#" & Query("SELECT * FROM tblClientGroup WHERE ClientGroupName= '" & txtClientGroupName & "'").Fields("ClientGroupID"), txtClientGroupName & " (0)", 1)
GroupNode.Selected = True

cboGroupName.AddItem txtClientGroupName
End Sub

Private Sub cmdGroupRemove_Click()
If MsgBox("Are you sure you want to remove this group?", vbYesNo + vbDefaultButton2 + vbQuestion, "Remove group?") = vbYes Then
    If Query("SELECT * FROM tblClient WHERE ClientGroupID= " & GetGroupID).RecordCount > 0 Then
        If MsgBox("There are " & Query("SELECT * FROM tblClient WHERE ClientGroupID= " & GetGroupID).RecordCount & " client(s) belong to this group. Removing this group will also remove the clients and this function cannot be reversed. Are you sure you want to remove the group and all of it's clients?", vbYesNo + vbDefaultButton2 + vbQuestion, "Remove group & client?") = vbYes Then
            QueryExec "DELETE FROM tblClient WHERE ClientGroupID= " & GetGroupID
        Else
            Exit Sub
        End If
    End If
    QueryExec "DELETE FROM tblClientGroup WHERE ClientGroupID= " & GetGroupID
    
    'LoadClientTree
    cboGroupName.RemoveItem ItemPosInListObj(cboGroupName, Left(tvwClient.Nodes("Group#" & GetGroupID).Text, Len(tvwClient.Nodes("Group#" & GetGroupID).Text) - 4)) - 1
    tvwClient.Nodes.Remove (tvwClient.SelectedItem.Key)
    tvwClient_Click
Else
    Exit Sub
End If
End Sub

Private Sub cmdGroupUpdate_Click()
'Check for null group name!
If Trim(txtClientGroupName) = "" Then
    MsgBox "Sorry, the group name you provided is invalid! Please try with another group name.", vbCritical + vbOKOnly, "Invalid group name!"
    Exit Sub
End If

'Check for duplicate group name!
If Query("SELECT * FROM tblClientGroup WHERE ClientGroupID <> " & GetGroupID & " AND ClientGroupName= '" & txtClientGroupName & "'").RecordCount > 0 Then
    MsgBox "Sorry, you provided a duplicate group name! Please try with another group name.", vbCritical + vbOKOnly, "Duplicate group name!"
    Exit Sub
End If

QueryExec "UPDATE tblClientGroup SET ClientGroupName = '" & txtClientGroupName & "', LineRent= " & txtLineRent & ", TAX= " & txtTAX & ", LateFeePercent= " & txtLateFeePercent & ", Description= '" & txtDescription & "' WHERE ClientGroupID= " & GetGroupID
cboGroupName.List(ItemPosInListObj(cboGroupName, Left(tvwClient.Nodes("Group#" & GetGroupID).Text, Len(tvwClient.Nodes("Group#" & GetGroupID).Text) - 4)) - 1) = txtClientGroupName
tvwClient.Nodes("Group#" & GetGroupID).Text = txtClientGroupName & " (" & Query("SELECT * FROM tblClient WHERE ClientGroupID= " & GetGroupID).RecordCount & ")"
End Sub

Private Sub Form_Load()
LoadClientTree
End Sub

Private Sub tvwClient_Click()
With Query("SELECT ClientGroupID, ClientGroupName, FORMAT(LineRent, '0.00') AS LineRent, TAX, LateFeePercent, Description FROM tblClientGroup WHERE ClientGroupID= " & GetGroupID)
    txtClientGroupName = .Fields("ClientGroupName")
    txtLineRent = .Fields("LineRent")
    txtTAX = .Fields("TAX")
    txtLateFeePercent = .Fields("LateFeePercent")
    txtDescription = .Fields("Description")
    
'    cboGroupName.ListIndex = ItemPosInListObj(cboGroupName, Left(tvwClient.SelectedItem.Text, Len(tvwClient.SelectedItem.Text) - 4)) - 1
End With

If Not IsGroupNode(tvwClient.SelectedItem) Then 'This is a client node, so populate the client area
    With Query("SELECT tblClient.*, tblClientGroup.* FROM tblClientGroup RIGHT JOIN tblClient ON tblClientGroup.ClientGroupID = tblClient.ClientGroupID WHERE tblClient.ClientID= " & GetClientID)
        txtFirstName = .Fields("FirstName")
        txtMiddleName = .Fields("MiddleName")
        txtLastName = .Fields("LastName")
        cboGroupName = .Fields("ClientGroupName")
        txtExtension = .Fields("Extension")
        txtExtension.Tag = .Fields("ClientID")
        txtArear = .Fields("Arear")
        txtOtherCharge = .Fields("OtherCharge")
        txtDeposit = .Fields("Deposit")
        txtStreet = .Fields("Street")
        txtCity = .Fields("City")
        txtZIP = .Fields("ZIP")
        txtState = .Fields("State")
        txtCountry = .Fields("Country")
        txtEmail = .Fields("Email")
        txtNote = .Fields("Note")
    End With
End If
End Sub

Private Function GetGroupID() As Long
If Left(tvwClient.SelectedItem.Key, 6) = "Group#" Then 'This is a group node
    GetGroupID = Val(Mid(tvwClient.SelectedItem.Key, 7))
Else 'This is a client node
    GetGroupID = Val(Mid(tvwClient.SelectedItem.Parent.Key, 7))
End If
End Function

Private Function GetClientID() As Long
If Not IsGroupNode(tvwClient.SelectedItem) Then
    GetClientID = Mid(tvwClient.SelectedItem.Key, 8)
End If
End Function

Private Function IsGroupNode(NodeToCheck As Node) As Boolean
If Left(NodeToCheck.Key, 6) = "Group#" Then IsGroupNode = True Else IsGroupNode = False
End Function

Private Sub txtArear_GotFocus()
SelectAllText txtArear

End Sub

Private Sub txtCity_GotFocus()
SelectAllText txtCity

End Sub

Private Sub txtClientGroupName_GotFocus()
SelectAllText txtClientGroupName
End Sub

Private Sub txtCountry_GotFocus()
SelectAllText txtCountry

End Sub

Private Sub txtDeposit_GotFocus()
SelectAllText txtDeposit

End Sub

Private Sub txtDescription_GotFocus()
SelectAllText txtDescription

End Sub

Private Sub txtEmail_GotFocus()
SelectAllText txtEmail

End Sub

Private Sub txtExtension_GotFocus()
SelectAllText txtExtension

End Sub

Private Sub txtFirstName_GotFocus()
SelectAllText txtFirstName

End Sub

Private Sub txtLastName_GotFocus()
SelectAllText txtLastName

End Sub

Private Sub txtLateFeePercent_GotFocus()
SelectAllText txtLateFeePercent

End Sub

Private Sub txtLineRent_GotFocus()
SelectAllText txtLineRent

End Sub

Private Sub txtMiddleName_GotFocus()
SelectAllText txtMiddleName

End Sub

Private Sub txtNote_GotFocus()
SelectAllText txtNote

End Sub

Private Sub txtOtherCharge_GotFocus()
SelectAllText txtOtherCharge

End Sub

Private Sub txtState_GotFocus()
SelectAllText txtState

End Sub

Private Sub txtStreet_GotFocus()
SelectAllText txtStreet

End Sub

Private Sub txtTAX_GotFocus()
SelectAllText txtTAX

End Sub

Private Sub txtZIP_GotFocus()
SelectAllText txtZIP

End Sub
