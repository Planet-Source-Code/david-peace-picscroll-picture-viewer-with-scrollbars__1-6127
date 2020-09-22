VERSION 5.00
Begin VB.UserControl PicScroll 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3450
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3615
   ScaleHeight     =   3450
   ScaleWidth      =   3615
   ToolboxBitmap   =   "PicScroll.ctx":0000
   Begin VB.PictureBox PicBack 
      AutoRedraw      =   -1  'True
      Height          =   3175
      Left            =   0
      ScaleHeight     =   3120
      ScaleWidth      =   3300
      TabIndex        =   2
      Top             =   0
      Width           =   3355
      Begin VB.PictureBox Pic 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   745
         Left            =   0
         ScaleHeight     =   750
         ScaleWidth      =   750
         TabIndex        =   3
         Top             =   0
         Width           =   745
      End
   End
   Begin VB.HScrollBar ScrollSide 
      Height          =   255
      Left            =   0
      Max             =   0
      MousePointer    =   1  'Arrow
      TabIndex        =   1
      Top             =   3180
      Width           =   3375
   End
   Begin VB.VScrollBar ScrollUp 
      Height          =   3195
      Left            =   3360
      Max             =   0
      MousePointer    =   1  'Arrow
      TabIndex        =   0
      Top             =   0
      Width           =   255
   End
End
Attribute VB_Name = "PicScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Enum BStyle
    None
    Fixed Single
End Enum

Enum MPointer
    Default
    Arrow
    Cross
    I Beam
    Icon
    Size
    Size NE SW
    Size NS
    Size NW SE
    Size WE
    Up Arrow
    Hourglass
    No Drop
    Arrow and Hourglass
    Arrow and Question
    Size All
    Custom
End Enum

Event Click()
Event DblClick()
Event VerticleChange()
Event HorizontalChange()
Event VerticleScroll()
Event HorizontalScroll()
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

Const m_def_VertScrollWidth = 255
Const m_def_HorizScrollHeight = 255
Dim m_VertScrollWidth As Variant
Dim m_HorizScrollHeight As Variant

Private Sub Pic_Click()
RaiseEvent Click
End Sub

Private Sub Pic_DblClick()
RaiseEvent DblClick
End Sub

Private Sub Pic_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub Pic_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub Pic_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Private Sub ScrollSide_Change()
RaiseEvent HorizontalChange
Pic.Left = ScrollSide.Value
End Sub

Private Sub ScrollSide_Scroll()
RaiseEvent HorizontalScroll
Pic.Left = ScrollSide.Value
End Sub

Private Sub ScrollUp_Change()
RaiseEvent VerticleChange
Pic.Top = ScrollUp.Value
End Sub

Private Sub ScrollUp_Scroll()
RaiseEvent VerticleScroll
Pic.Top = ScrollUp.Value
End Sub

Private Sub UserControl_Resize()
Pic.Height = PicHeight
Pic.Width = PicWidth
If Width < 845 + VertScrollWidth Then Width = 845 + VertScrollWidth
If Height < 845 + HorizScrollHeight Then Height = 845 + HorizScrollHeight
PicBack.Move 0, 0, Width - VertScrollWidth - 20, Height - HorizScrollHeight - 20
ScrollUp.Move Width - VertScrollWidth, 0, VertScrollWidth, Height - HorizScrollHeight
ScrollSide.Move 0, Height - HorizScrollHeight, Width - VertScrollWidth, HorizScrollHeight
If UserControl.Enabled Then
    If PicBack.Width > Pic.Width Then ScrollSide.Enabled = False Else ScrollSide.Enabled = True
    If PicBack.Height > Pic.Height Then ScrollUp.Enabled = False Else ScrollUp.Enabled = True
    SetScrollBars Pic, PicBack, ScrollUp, ScrollSide
Else
    ScrollSide.Enabled = False
    ScrollUp.Enabled = False
    UserControl.Enabled = False
End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    PicBack.BorderStyle = PropBag.ReadProperty("BorderStyle", 1)
    PicBack.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    Set Picture = PropBag.ReadProperty("Picture", Nothing)
    Pic.AutoRedraw = PropBag.ReadProperty("AutoRedraw", False)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    PicBack.MousePointer = PropBag.ReadProperty("MousePointer", 0)
    m_VertScrollWidth = PropBag.ReadProperty("VertScrollWidth", m_def_VertScrollWidth)
    m_HorizScrollHeight = PropBag.ReadProperty("HorizScrollHeight", m_def_HorizScrollHeight)
    m_ScrollBars = PropBag.ReadProperty("ScrollBars", m_def_ScrollBars)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Pic.ScaleLeft = PropBag.ReadProperty("PicHeight", 5)
    Pic.ScaleTop = PropBag.ReadProperty("PicWidth", 5)
    Pic.BackColor = PropBag.ReadProperty("PicBackColor", &H8000000F)
    Pic.AutoSize = PropBag.ReadProperty("AutoSize", True)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("BorderStyle", PicBack.BorderStyle, 1)
    Call PropBag.WriteProperty("BackColor", PicBack.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Picture", Picture, Nothing)
    Call PropBag.WriteProperty("AutoRedraw", Pic.AutoRedraw, False)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", PicBack.MousePointer, 0)
    Call PropBag.WriteProperty("VertScrollWidth", m_VertScrollWidth, m_def_VertScrollWidth)
    Call PropBag.WriteProperty("HorizScrollHeight", m_HorizScrollHeight, m_def_HorizScrollHeight)
    Call PropBag.WriteProperty("ScrollBars", m_ScrollBars, m_def_ScrollBars)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("PicHeight", Pic.Height, 5)
    Call PropBag.WriteProperty("PicWidth", Pic.Width, 5)
    Call PropBag.WriteProperty("PicBackColor", Pic.BackColor, &H8000000F)
    Call PropBag.WriteProperty("AutoSize", Pic.AutoSize, True)
End Sub

Public Sub Cls()
    Pic.Cls
End Sub

Public Property Get BorderStyle() As BStyle
Attribute BorderStyle.VB_Description = "Returns/sets the border style of the Frame."
    BorderStyle = PicBack.BorderStyle
End Property

Public Property Let BorderStyle(ByVal New_BorderStyle As BStyle)
    PicBack.BorderStyle() = New_BorderStyle
    PropertyChanged "BorderStyle"
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    BackColor = PicBack.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    PicBack.BackColor() = New_BackColor
    PropertyChanged "BackColor"
End Property

Public Property Get Picture() As Picture
Attribute Picture.VB_Description = "Returns/sets a graphic to be displayed in a control."
    Set Picture = Pic.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set Pic.Picture = New_Picture
    UserControl_Resize
    PropertyChanged "Picture"
End Property

Private Function SetScrollBars(PicFront As PictureBox, PicBack As PictureBox, Vert As VScrollBar, Horz As HScrollBar)
Vert.Max = -(PicFront.Height - PicBack.Height)
Vert.SmallChange = 100
Vert.LargeChange = PicBack.Height / 4
Horz.Max = -(PicFront.Width - PicBack.Width)
Horz.SmallChange = 100
Horz.LargeChange = PicBack.Width / 4
End Function

Public Property Get AutoRedraw() As Boolean
Attribute AutoRedraw.VB_Description = "Returns/sets the output from a graphics method to a persistent bitmap."
    AutoRedraw = Pic.AutoRedraw
End Property

Public Property Let AutoRedraw(ByVal New_AutoRedraw As Boolean)
    Pic.AutoRedraw() = New_AutoRedraw
    PropertyChanged "AutoRedraw"
End Property

Private Sub PicBack_Click()
    RaiseEvent Click
End Sub

Private Sub PicBack_DblClick()
    RaiseEvent DblClick
End Sub

Public Property Get hDC() As Long
Attribute hDC.VB_Description = "Returns a handle (from Microsoft Windows) to the object's device context."
    hDC = Pic.hDC
End Property

Private Sub PicBack_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = PicBack.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    On Error GoTo TheEnd
    Set PicBack.MouseIcon = New_MouseIcon
    Set Pic.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
TheEnd:
End Property

Private Sub PicBack_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Public Property Get MousePointer() As MPointer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    If PicBack.MousePointer = 99 Then
        MousePointer = 16
    Else
        MousePointer = PicBack.MousePointer
    End If
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MPointer)
    If New_MousePointer = 16 Then
        PicBack.MousePointer() = 99
        Pic.MousePointer() = 99
    Else
        PicBack.MousePointer() = New_MousePointer
        Pic.MousePointer() = New_MousePointer
    End If
    PropertyChanged "MousePointer"
End Property

Private Sub PicBack_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

Public Function Point_(X As Single, Y As Single) As Long
Attribute Point_.VB_Description = "Returns, as an integer of type Long, the RGB color of the specified point on a Form or PictureBox object."
    Point_ = Pic.Point(X, Y)
End Function

Public Sub PSet_(X As Single, Y As Single, Color As OLE_COLOR)
    Pic.PSet Step(X, Y), Color
End Sub
Public Property Get VertScrollWidth() As Variant
Attribute VertScrollWidth.VB_Description = "Returns\\sets the Width of the Verticle Scroll Bar."
    VertScrollWidth = m_VertScrollWidth
End Property

Public Property Let VertScrollWidth(ByVal New_VertScrollWidth As Variant)
    If Not IsNumeric(New_VertScrollWidth) Then New_VertScrollWidth = VertScrollWidth
    If New_VertScrollWidth <= 0 Then New_VertScrollWidth = VertScrollWidth
    m_VertScrollWidth = New_VertScrollWidth
    ScrollUp.Width = New_VertScrollWidth
    Width = PicBack.Width + New_VertScrollWidth + 20
    PropertyChanged "VertScrollWidth"
End Property

Public Property Get HorizScrollHeight() As Variant
Attribute HorizScrollHeight.VB_Description = "Returns\\sets the Height of the Horizontal Scroll Bar"
    HorizScrollHeight = m_HorizScrollHeight
End Property

Public Property Let HorizScrollHeight(ByVal New_HorizScrollHeight As Variant)
    If Not IsNumeric(New_HorizScrollHeight) Then New_HorizScrollHeight = HorizScrollHeight
    If New_HorizScrollHeight <= 0 Then New_HorizScrollHeight = HorizScrollHeight
    m_HorizScrollHeight = New_HorizScrollHeight
    ScrollSide.Height = New_HorizScrollHeight
    Height = PicBack.Height + New_HorizScrollHeight + 20
    PropertyChanged "HorizScrollHeight"
End Property

Private Sub UserControl_InitProperties()
    m_VertScrollWidth = m_def_VertScrollWidth
    m_HorizScrollHeight = m_def_HorizScrollHeight
End Sub

Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events."
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    ScrollUp.Enabled() = New_Enabled
    ScrollSide.Enabled() = New_Enabled
    UserControl_Resize
    PropertyChanged "Enabled"
End Property

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object."
    Pic.Refresh
End Sub

Public Property Get PicHeight() As Single
Attribute PicHeight.VB_Description = "Returns/sets the height of the loaded picture."
    PicHeight = Pic.Height
End Property

Public Property Let PicHeight(ByVal New_PicHeight As Single)
    If AutoSize Then New_PicHeight = PicHeight: MsgBox "Read only while Autosize = True", vbCritical
    If Not IsNumeric(New_PicHeight) Then New_PicHeight = PicHeight
    If New_PicHeight <= 1 Then New_PicHeight = PicHeight
    Pic.Height() = New_PicHeight
    UserControl_Resize
    PropertyChanged "PicHeight"
End Property

Public Property Get PicWidth() As Single
Attribute PicWidth.VB_Description = "Returns/sets the width of the loaded picture."
    PicWidth = Pic.Width
End Property

Public Property Let PicWidth(ByVal New_PicWidth As Single)
    If AutoSize Then New_PicWidth = PicWidth: MsgBox "Read only while Autosize = True", vbCritical
    If Not IsNumeric(New_PicWidth) Then New_PicWidth = PicWidth
    If New_PicWidth <= 1 Then New_PicWidth = PicWidth
    Pic.Width() = New_PicWidth
    UserControl_Resize
    PropertyChanged "PicWidth"
End Property

Public Property Get PicBackColor() As OLE_COLOR
Attribute PicBackColor.VB_Description = "Returns/sets the background color used to display text and graphics in an object."
    PicBackColor = Pic.BackColor
End Property

Public Property Let PicBackColor(ByVal New_PicBackColor As OLE_COLOR)
    Pic.BackColor() = New_PicBackColor
    PropertyChanged "PicBackColor"
End Property

Public Property Get AutoSize() As Boolean
Attribute AutoSize.VB_Description = "Determines whether a control is automatically resized to display its entire contents."
    AutoSize = Pic.AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    Pic.AutoSize() = New_AutoSize
    PropertyChanged "AutoSize"
End Property

