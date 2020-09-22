VERSION 5.00
Begin VB.UserControl ShadowLabel 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   1215
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2670
   BeginProperty Font 
      Name            =   "Trebuchet MS"
      Size            =   9
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ScaleHeight     =   81
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   178
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   435
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   795
   End
   Begin VB.Label lblShadow 
      BackStyle       =   0  'Transparent
      Caption         =   "Title"
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00808080&
      Height          =   435
      Left            =   480
      TabIndex        =   1
      Top             =   720
      Width           =   795
   End
End
Attribute VB_Name = "ShadowLabel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Default Property Values:
Const m_def_Padding = 0
Const m_def_ShadowDistance = 5
'Property Variables:
Dim m_Padding As Long
Dim m_ShadowDistance As Long
'Event Declarations:
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object."




Private Sub UserControl_Resize()
lblTitle.Move m_Padding, m_Padding, UserControl.ScaleWidth - m_ShadowDistance - (2 * m_Padding), UserControl.ScaleHeight - m_ShadowDistance - m_Padding
lblShadow.Move m_ShadowDistance - 1 + m_Padding, m_ShadowDistance - 1 + m_Padding, UserControl.ScaleWidth - m_ShadowDistance - (2 * m_Padding), UserControl.ScaleHeight - m_ShadowDistance - m_Padding
End Sub
'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTitle,lblTitle,-1,Caption
Public Property Get Title() As String
Attribute Title.VB_Description = "Returns/sets the text displayed in an object's title bar or below an object's icon."
    Title = lblTitle.Caption
End Property

Public Property Let Title(ByVal New_Title As String)
    lblTitle.Caption() = New_Title
    lblShadow.Caption() = New_Title
    PropertyChanged "Title"
    
    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTitle,lblTitle,-1,ForeColor
Public Property Get TitleColor() As OLE_COLOR
Attribute TitleColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    TitleColor = lblTitle.ForeColor
End Property

Public Property Let TitleColor(ByVal New_TitleColor As OLE_COLOR)
    lblTitle.ForeColor() = New_TitleColor
    PropertyChanged "TitleColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblShadow,lblShadow,-1,ForeColor
Public Property Get ShadowColor() As OLE_COLOR
Attribute ShadowColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ShadowColor = lblShadow.ForeColor
End Property

Public Property Let ShadowColor(ByVal New_ShadowColor As OLE_COLOR)
    lblShadow.ForeColor() = New_ShadowColor
    PropertyChanged "ShadowColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,5
Public Property Get ShadowDistance() As Long
    ShadowDistance = m_ShadowDistance
End Property

Public Property Let ShadowDistance(ByVal New_ShadowDistance As Long)
    m_ShadowDistance = New_ShadowDistance
    PropertyChanged "ShadowDistance"

    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTitle,lblTitle,-1,Font
Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = lblTitle.Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set lblTitle.Font = New_Font
    Set lblShadow.Font = New_Font
    PropertyChanged "Font"
    
    UserControl_Resize
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_ShadowDistance = m_def_ShadowDistance
    m_Padding = m_def_Padding
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    lblTitle.Caption = PropBag.ReadProperty("Title", "Title")
    lblShadow.Caption = PropBag.ReadProperty("Title", "Title")
    lblTitle.ForeColor = PropBag.ReadProperty("TitleColor", &HC00000)
    lblShadow.ForeColor = PropBag.ReadProperty("ShadowColor", &H808080)
    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    m_ShadowDistance = PropBag.ReadProperty("ShadowDistance", m_def_ShadowDistance)
    Set lblTitle.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set lblShadow.Font = PropBag.ReadProperty("Font", Ambient.Font)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    UserControl.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    UserControl.BackStyle = PropBag.ReadProperty("BackStyle", 0)
    m_Padding = PropBag.ReadProperty("Padding", m_def_Padding)
    lblTitle.Alignment = PropBag.ReadProperty("Alignment", 0)
    lblShadow.Alignment = PropBag.ReadProperty("Alignment", 0)
End Sub

Private Sub UserControl_Show()
UserControl_Resize
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("Title", lblTitle.Caption, "Title")
    Call PropBag.WriteProperty("TitleColor", lblTitle.ForeColor, &HC00000)
    Call PropBag.WriteProperty("ShadowColor", lblShadow.ForeColor, &H808080)
    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("ShadowDistance", m_ShadowDistance, m_def_ShadowDistance)
    Call PropBag.WriteProperty("Font", lblTitle.Font, Ambient.Font)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", UserControl.MousePointer, 0)
    Call PropBag.WriteProperty("BackStyle", UserControl.BackStyle, 0)
    Call PropBag.WriteProperty("Padding", m_Padding, m_def_Padding)
    Call PropBag.WriteProperty("Alignment", lblTitle.Alignment, 0)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MousePointer
Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = UserControl.MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    UserControl.MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackStyle
Public Property Get BackStyle() As Integer
Attribute BackStyle.VB_Description = "Indicates whether a Label or the background of a Shape is transparent or opaque."
    BackStyle = UserControl.BackStyle
End Property

Public Property Let BackStyle(ByVal New_BackStyle As Integer)
    UserControl.BackStyle() = New_BackStyle
    PropertyChanged "BackStyle"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=8,0,0,0
Public Property Get Padding() As Long
    Padding = m_Padding
End Property

Public Property Let Padding(ByVal New_Padding As Long)
    m_Padding = New_Padding
    PropertyChanged "Padding"

    UserControl_Resize
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=lblTitle,lblTitle,-1,Alignment
Public Property Get Alignment() As Integer
Attribute Alignment.VB_Description = "Returns/sets the alignment of a CheckBox or OptionButton, or a control's text."
    Alignment = lblTitle.Alignment
End Property

Public Property Let Alignment(ByVal New_Alignment As Integer)
    lblTitle.Alignment() = New_Alignment
    lblShadow.Alignment() = New_Alignment
    PropertyChanged "Alignment"
End Property

