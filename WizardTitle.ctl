VERSION 5.00
Begin VB.UserControl WizardTitle 
   AutoRedraw      =   -1  'True
   CanGetFocus     =   0   'False
   ClientHeight    =   1605
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9.75
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   107
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Label lblDescription 
      Caption         =   "Description"
      Height          =   270
      Left            =   120
      TabIndex        =   1
      Top             =   480
      Width           =   1050
   End
   Begin VB.Label lblTitle 
      AutoSize        =   -1  'True
      Caption         =   "Title"
      Height          =   270
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   420
   End
End
Attribute VB_Name = "WizardTitle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_LightColor = vbWhite
Const m_def_DarkColor = vbBlack
'Property Variables:
Dim m_LightColor As OLE_COLOR
Dim m_DarkColor As OLE_COLOR



Private Sub UserControl_Resize()
Position
End Sub

Private Sub UserControl_Show()
Line (0, UserControl.ScaleHeight - 2)-(UserControl.ScaleWidth - 0, UserControl.ScaleHeight - 2), m_DarkColor
Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth - 0, UserControl.ScaleHeight - 1), m_LightColor

Position
End Sub

Private Sub Position()
lblTitle.Move 0, 0, UserControl.ScaleWidth, lblTitle.Height
lblDescription.Move 0, lblTitle.Height, UserControl.ScaleWidth, UserControl.ScaleHeight - lblTitle.Height - 2
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTitle,lblTitle,-1,Caption
Public Property Get Title() As String
Attribute Title.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Title = lblTitle.Caption
End Property

Public Property Let Title(ByVal New_Title As String)
    lblTitle.Caption() = New_Title
    PropertyChanged "Title"

    Position
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblDescription,lblDescription,-1,Caption
Public Property Get Description() As String
Attribute Description.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Description = lblDescription.Caption
End Property

Public Property Let Description(ByVal New_Description As String)
    lblDescription.Caption() = New_Description
    PropertyChanged "Description"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTitle,lblTitle,-1,ForeColor
Public Property Get TitleForeColor() As OLE_COLOR
Attribute TitleForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    TitleForeColor = lblTitle.ForeColor
End Property

Public Property Let TitleForeColor(ByVal New_TitleForeColor As OLE_COLOR)
    lblTitle.ForeColor() = New_TitleForeColor
    PropertyChanged "TitleForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTitle,lblTitle,-1,BackColor
Public Property Get TitleBackColor() As OLE_COLOR
Attribute TitleBackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    TitleBackColor = lblTitle.BackColor
End Property

Public Property Let TitleBackColor(ByVal New_TitleBackColor As OLE_COLOR)
    lblTitle.BackColor() = New_TitleBackColor
    PropertyChanged "TitleBackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblDescription,lblDescription,-1,ForeColor
Public Property Get DescriptionForeColor() As OLE_COLOR
Attribute DescriptionForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    DescriptionForeColor = lblDescription.ForeColor
End Property

Public Property Let DescriptionForeColor(ByVal New_DescriptionForeColor As OLE_COLOR)
    lblDescription.ForeColor() = New_DescriptionForeColor
    PropertyChanged "DescriptionForeColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblDescription,lblDescription,-1,BackColor
Public Property Get DescriptionBackColor() As OLE_COLOR
Attribute DescriptionBackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    DescriptionBackColor = lblDescription.BackColor
End Property

Public Property Let DescriptionBackColor(ByVal New_DescriptionBackColor As OLE_COLOR)
    lblDescription.BackColor() = New_DescriptionBackColor
    PropertyChanged "DescriptionBackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTitle,lblTitle,-1,Font
Public Property Get TitleFont() As Font
Attribute TitleFont.VB_Description = "Returns a Font object."
    Set TitleFont = lblTitle.Font
End Property

Public Property Set TitleFont(ByVal New_TitleFont As Font)
    Set lblTitle.Font = New_TitleFont
    PropertyChanged "TitleFont"

    Position
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblDescription,lblDescription,-1,Font
Public Property Get DescriptionFont() As Font
Attribute DescriptionFont.VB_Description = "Returns a Font object."
    Set DescriptionFont = lblDescription.Font
End Property

Public Property Set DescriptionFont(ByVal New_DescriptionFont As Font)
    Set lblDescription.Font = New_DescriptionFont
    PropertyChanged "DescriptionFont"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbWhite
Public Property Get LightColor() As OLE_COLOR
    LightColor = m_LightColor
End Property

Public Property Let LightColor(ByVal New_LightColor As OLE_COLOR)
    m_LightColor = New_LightColor
    PropertyChanged "LightColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=10,0,0,vbBlack
Public Property Get DarkColor() As OLE_COLOR
    DarkColor = m_DarkColor
End Property

Public Property Let DarkColor(ByVal New_DarkColor As OLE_COLOR)
    m_DarkColor = New_DarkColor
    PropertyChanged "DarkColor"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_LightColor = m_def_LightColor
    m_DarkColor = m_def_DarkColor
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lblTitle.Caption = PropBag.ReadProperty("Title", "Title")
    lblDescription.Caption = PropBag.ReadProperty("Description", "Description")
    lblTitle.ForeColor = PropBag.ReadProperty("TitleForeColor", &H80000012)
    lblTitle.BackColor = PropBag.ReadProperty("TitleBackColor", &H8000000F)
    lblDescription.ForeColor = PropBag.ReadProperty("DescriptionForeColor", &H80000012)
    lblDescription.BackColor = PropBag.ReadProperty("DescriptionBackColor", &H8000000F)
    Set lblTitle.Font = PropBag.ReadProperty("TitleFont", Ambient.Font)
    Set lblDescription.Font = PropBag.ReadProperty("DescriptionFont", Ambient.Font)
    m_LightColor = PropBag.ReadProperty("LightColor", m_def_LightColor)
    m_DarkColor = PropBag.ReadProperty("DarkColor", m_def_DarkColor)
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Title", lblTitle.Caption, "Title")
    Call PropBag.WriteProperty("Description", lblDescription.Caption, "Description")
    Call PropBag.WriteProperty("TitleForeColor", lblTitle.ForeColor, &H80000012)
    Call PropBag.WriteProperty("TitleBackColor", lblTitle.BackColor, &H8000000F)
    Call PropBag.WriteProperty("DescriptionForeColor", lblDescription.ForeColor, &H80000012)
    Call PropBag.WriteProperty("DescriptionBackColor", lblDescription.BackColor, &H8000000F)
    Call PropBag.WriteProperty("TitleFont", lblTitle.Font, Ambient.Font)
    Call PropBag.WriteProperty("DescriptionFont", lblDescription.Font, Ambient.Font)
    Call PropBag.WriteProperty("LightColor", m_LightColor, m_def_LightColor)
    Call PropBag.WriteProperty("DarkColor", m_DarkColor, m_def_DarkColor)
End Sub

