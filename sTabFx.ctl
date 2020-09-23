VERSION 5.00
Begin VB.UserControl sTabFx 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000000&
   ClientHeight    =   1620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3150
   ClipControls    =   0   'False
   ControlContainer=   -1  'True
   MouseIcon       =   "sTabFx.ctx":0000
   ScaleHeight     =   108
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   210
   Begin VB.PictureBox TabButton 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   315
      Index           =   0
      Left            =   0
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   69
      TabIndex        =   0
      Top             =   0
      Width           =   1035
   End
End
Attribute VB_Name = "sTabFx"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'DM Tab Control Replacement
' Hi this is a little Tab Control I made about 2 hours ago
' It was just ment to be a simple control but as you start on something you can never stop.
' Features
' Add new tabs

' Show Hottracking
' Change hottracking color
' Show or hide hottracking underline
' Show of hide Highlighted Tabs like in the normal vb on
' Show of hide rect focus
' Chnage the 3D Border thickness
' Change tabs fonts
' Turn on or off selected tab captions in bold
' Chnage the tab style between Tabs or Buttons

'Does not support the removeing of Tabs but I try and add this next time
' if you want to use this in your projects please do so
'all i ask is you remmber who gave it to you.




Private Declare Function ReleaseCapture Lib "user32" () As Long
Private Declare Function GetCapture Lib "user32" () As Long
Private Declare Function SetCapture Lib "user32" (ByVal hwnd As Long) As Long

Private Const m_TabBar_Height = 20
Private Const m_TabBar_Color As Long = vb3DFace

Dim m_ButtonCaption As String
Dim TabIndex As Integer

Dim m_HotTrack As Boolean, m_HotTrackColor As OLE_COLOR, fBold As Boolean
Dim m_TabForeColor As OLE_COLOR, m_TrackUnderline As Boolean, m_Track_Tmp As Boolean
Dim m_HightLight As Boolean, mShowRect As Boolean, mToolTip As Boolean, mShowTrackHand As Boolean

Enum Border3D
    Thin = 1
    Thick = 2
End Enum

Enum tbStyle
    Tabs = 0
    Buttons
End Enum

Private Enum enTabButton
    Selected = 0
    NotSelected = 1
End Enum

Private Type Tabs
    tbCaption As String
    tbKey As String
    tbSelected As Boolean
    tHighlight As Boolean
End Type

Private tmpIdx As Integer

Private TabsX() As Tabs
Private TabCounter As Integer, bIndex As Integer
Dim bStyle As Border3D, tStyle As tbStyle
Event TabMouseMove(index As Integer, Selected As Boolean, Key As String, Caption As String)
Event Click(index As Integer, Key As String, Caption As String)

Private Sub UpDate()
    Redraw
    If TabsX(TabIndex).tbSelected Then
        Call DrawTabs(TabIndex, Selected)
    Else
       Call DrawTabs(TabIndex, NotSelected)
    End If
End Sub

Public Sub Reset()
    Erase TabsX
    ReDim TabsX(0)
    TabsX(0).tbCaption = "Tab 0"
    TabsX(0).tbSelected = True
    TabIndex = 0
    tmpIdx = 0
    
    If TabCounter <> 0 Then
        For x = 1 To TabButton.Count - 1
            Unload TabButton(x)
        Next
    End If
    
    TabCounter = 0
    DrawTabs 0, Selected
    
End Sub

Sub AddTab(Optional Caption As String, Optional Key As String)
    TabCounter = TabCounter + 1
    ReDim Preserve TabsX(TabCounter)
    TabsX(TabCounter).tbSelected = False
    TabsX(TabCounter).tbCaption = Caption
    TabsX(TabCounter).tbKey = Key
    '
    Load TabButton(TabCounter)
    TabButton(TabCounter).Visible = True
    
    DrawTabs TabCounter, NotSelected
End Sub

Sub Redraw()

    UserControl.Cls 'Clear the tab control
    'Draw top Bar of the tab contol
    
    If tStyle = Buttons Then Exit Sub
    UserControl.Line (0, 0)-(UserControl.Width, m_TabBar_Height), m_TabBar_Color, BF
    'Draw 3D Effect around the tab control
    UserControl.Line (0, m_TabBar_Height)-(UserControl.Width, m_TabBar_Height), vbWhite, BF 'Top Line
    UserControl.Line (0, m_TabBar_Height)-(0, UserControl.ScaleHeight - 1), vbWhite 'Left line
    UserControl.Line (UserControl.ScaleWidth - 1, m_TabBar_Height)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight), &H808080 'Right line
    UserControl.Line (0, UserControl.ScaleHeight - 1)-(UserControl.ScaleWidth - 1, UserControl.ScaleHeight - 1), &H808080 'Bottom line
    
End Sub

Private Sub DrawTabs(index As Integer, ButtonOp As enTabButton)
Dim LeftLine As Long, RightLine As Long, TopLine As Long, TabColor As Long
Dim ButObj As Object, TopPos As Integer, TheHeight As Integer, LineBottom As Long
Dim x As Integer, y As Integer, x_pos As Integer, y_pos As Integer, mStep As Byte

On Error Resume Next

    m_ButtonCaption = TabsX(index).tbCaption
    
    If ButtonOp = Selected Then
        If TabStyle = Buttons Then
            TabColor = &HF5F5F5
            TopLine = vb3DShadow
            LeftLine = vb3DShadow
            RightLine = &HF5F5F5
            LineBottom = &HF5F5F5
            TopPos = 0
            TheHeight = (m_TabBar_Height + 1)
        Else
            TabColor = vbButtonFace
            TopLine = &HF5F5F5
            LeftLine = &HFFFFFF
            RightLine = vb3DShadow
            LineBottom = vbButtonFace
            TopPos = 0
            TheHeight = (m_TabBar_Height + 1)
        End If
        TabsX(index).tbSelected = True
        TabButton(index).FontBold = True
        
    ElseIf ButtonOp = NotSelected Then
        If TabStyle = Buttons Then
            TabColor = vbButtonFace
            TopLine = vbWhite
            LeftLine = vbWhite
            RightLine = vb3DShadow
            LineBottom = vb3DShadow
            TopPos = 0
            TheHeight = (m_TabBar_Height + 1)
        Else
            TabColor = &HF5F5F5
            TopLine = &HF5F5F5
            LeftLine = TopLine
            RightLine = TopLine
            LineBottom = TabColor
            TopPos = 2
            TheHeight = (m_TabBar_Height - 2)
        End If
        
        TabsX(index).tbSelected = False
        TabButton(index).FontBold = False
    End If

    If (TabButton(index).FontBold) And (fBold = False) Then TabButton(index).FontBold = False
    
    Set ButObj = TabButton(index)
   'Arrange all the Tab Buttons to correct widths and positions
   On Error Resume Next
    If TabButton.ubound = 0 Then
        TabButton(index).Width = TabButton(index).TextWidth(TabsX(0).tbCaption) + Screen.TwipsPerPixelX
    Else
        For x = 0 To TabCounter
            TabButton(x).Width = TabButton(x).TextWidth(TabsX(x).tbCaption) + Screen.TwipsPerPixelX
            TabButton(x).Left = TabButton(x - 1).Left + TabButton(x - 1).ScaleWidth + 1
        Next
    End If
    
    ButObj.Cls
    ButObj.DrawWidth = UserControl.DrawWidth
    ButObj.DrawMode = vbCopyPen
    
    'Show a highlighted
    If TabsX(index).tbSelected Then 'Check tab selection
        If TabsX(index).tHighlight Then 'Check highlight is enabled
            TabColor = vbHighlight 'set tab backcolor
            ButObj.ForeColor = vbWhite 'set forecolor
            tmpIdx = index 'Keep the old tab index
        Else
            TabButton(tmpIdx).ForeColor = UserControl.ForeColor 'TabButton(0).ForeColor
        End If
    End If
   
    ButObj.BackColor = TabColor
    'vb3DHighlight
    ButObj.Top = TopPos
    ButObj.Height = TheHeight
    
    'Draw Top line
    ButObj.Line (0, 0)-(ButObj.ScaleWidth - 1, 0), TopLine
    'Draw Left Line
    ButObj.Line (0, 0)-(0, ButObj.ScaleHeight), LeftLine
    'Draw Right Line
    ButObj.Line (ButObj.ScaleWidth - 1, 0)-(ButObj.ScaleWidth - 1, ButObj.ScaleHeight), RightLine
    
    ButObj.Line (0, ButObj.ScaleHeight - 1)-(ButObj.ScaleWidth, ButObj.ScaleHeight - 1), LineBottom
    
    
    'This bit of code is used to draw a focus rect around the tab
    ' This is not as good as using API but still not a bad attempt
    If (TabsX(index).tbSelected) And mShowRect Then
        ButObj.DrawWidth = 1
        ButObj.DrawMode = vbInvert
        
        For x = 3 To ButObj.ScaleWidth - 4 Step 3
            ButObj.PSet (x, 3), vbBlack
            ButObj.PSet (x, ButObj.ScaleHeight - 3), vbBlack
        Next
        x = 0
        
        For y = 3 To ButObj.ScaleHeight - 4 Step 3
            ButObj.PSet (2, y), vbBlack
            ButObj.PSet (ButObj.ScaleWidth - 4, y), vbBlack
        Next
        y = 0
        
        ButObj.PSet (2, 3), TabColor
    End If
    
    ButObj.Refresh
    
    'Center the caption on the tabs
    x_pos = (ButObj.ScaleWidth - ButObj.TextWidth(m_ButtonCaption)) \ 2
    y_pos = (ButObj.ScaleHeight - ButObj.TextHeight(m_ButtonCaption)) \ 2
     '
    ButObj.CurrentX = x_pos
    ButObj.CurrentY = y_pos
    ButObj.Print m_ButtonCaption
    
    
End Sub

Private Sub TabButton_Click(index As Integer)
    bIndex = index
    RaiseEvent Click(index, TabsX(index).tbKey, TabsX(index).tbCaption)
    
    If index <> TabIndex Then
        Call DrawTabs(index, Selected)
        Call DrawTabs(TabIndex, NotSelected)
    End If
    
    TabIndex = index

End Sub

Private Sub TabButton_MouseMove(index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)

    RaiseEvent TabMouseMove(index, TabsX(index).tbSelected, TabsX(index).tbKey, TabsX(index).tbCaption)
    
    If mToolTip Then
        TabButton(index).ToolTipText = TabsX(index).tbCaption
    Else
        TabButton(index).ToolTipText = ""
    End If
    
    If m_HotTrack <> True Then Exit Sub
    
    If m_Track_Tmp = True Then m_TrackUnderline = True

    If (x < 0) Or (x > TabButton(index).Width) Or (y < 0) Or (y > TabButton(index).Height) Then
        ReleaseCapture
        TabButton(index).ForeColor = m_TabForeColor
        TabButton(index).FontUnderline = m_Track_Tmp
    ElseIf GetCapture() <> TabButton(index).hwnd Then
        If mShowTrackHand Then
            TabButton(index).MousePointer = vbCustom
        Else
            TabButton(index).MousePointer = vbDefault
        End If
        
        TabButton(index).ForeColor = m_HotTrackColor
        TabButton(index).FontUnderline = m_TrackUnderline
        SetCapture TabButton(index).hwnd
    End If
    
    If TabsX(index).tbSelected Then
        DrawTabs index, Selected
    Else
        DrawTabs index, NotSelected
    End If
    
End Sub

Private Sub UserControl_InitProperties()
    m_HotTrackColor = vbBlue
    m_ButtonSelected = 0
    Set UserControl.Font = Ambient.Font
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_HotTrack = PropBag.ReadProperty("HotTracking", False)
    m_HotTrackColor = PropBag.ReadProperty("TrackingColor", vbBlue)
    fBold = PropBag.ReadProperty("BoldSelection", True)
    bStyle = PropBag.ReadProperty("Border3DStyle", 1)
    UserControl.DrawWidth = PropBag.ReadProperty("Style3D", 1)
    Set TabButton(0).Font = PropBag.ReadProperty("Font", Ambient.Font)
    TabButton(0).ForeColor = PropBag.ReadProperty("ForeColor", &H80000012)
    m_TrackUnderline = PropBag.ReadProperty("TrackUnderLine", False)
    mShowRect = PropBag.ReadProperty("ShowRect", True)
    mToolTip = PropBag.ReadProperty("ShowToolTip", True)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    TabButton(0).MousePointer = PropBag.ReadProperty("MousePointer", 0)
    mShowTrackHand = PropBag.ReadProperty("ShowTrackingHand", True)
    tStyle = PropBag.ReadProperty("TabStyle", 0)
End Sub

Private Sub UserControl_Resize()
    Call Redraw
End Sub

Private Sub UserControl_Show()
    m_TabForeColor = TabButton(0).ForeColor
    UserControl.ForeColor = TabButton(0).ForeColor 'Temp Color holder
    m_Track_Tmp = TabButton(0).Font.Underline
    TabButton(0).MouseIcon = UserControl.MouseIcon
    m_HightLight = vbButtonFace
    TabCounter = 0
    bStyle = Thin
    'tStyle = Tabs
    Call Reset
    Call Redraw
End Sub

Public Sub SelectTab(index As Integer)
    Call TabButton_Click(index)
End Sub

Public Property Get TabCaption(ByVal index As Integer) As String
    TabCaption = TabsX(index).tbCaption
End Property

Public Property Let TabCaption(ByVal index As Integer, ByVal vNewCaption As String)
On Error Resume Next
    TabsX(index).tbCaption = vNewCaption
    Call UpDate
End Property

Public Property Get TabKey(ByVal index As Integer) As String
    TabKey = TabsX(index).tbKey
End Property

Public Property Let TabKey(ByVal index As Integer, ByVal vNewKey As String)
    TabsX(index).tbKey = vNewKey
End Property

Public Property Get TabSelected(ByVal index As Integer) As Boolean
    TabSelected = TabsX(index).tbSelected
End Property

Public Property Get HotTracking() As Boolean
    HotTracking = m_HotTrack
End Property

Public Property Let HotTracking(ByVal vNewValue As Boolean)
    m_HotTrack = vNewValue
    PropertyChanged "HotTracking"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    PropBag.WriteProperty "HotTracking", m_HotTrack, False
    PropBag.WriteProperty "TrackingColor", m_HotTrackColor, vbBlue
    PropBag.WriteProperty "BoldSelection", fBold, True
    PropBag.WriteProperty "Border3DStyle", bStyle, 1
    Call PropBag.WriteProperty("Style3D", UserControl.DrawWidth, 1)
    Call PropBag.WriteProperty("Font", TabButton(0).Font, Ambient.Font)
    Call PropBag.WriteProperty("ForeColor", TabButton(0).ForeColor, &H80000012)
    Call PropBag.WriteProperty("TrackUnderLine", m_TrackUnderline, False)
    Call PropBag.WriteProperty("ShowRect", mShowRect, True)
    Call PropBag.WriteProperty("ShowToolTip", mToolTip, True)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", TabButton(0).MousePointer, 0)
    Call PropBag.WriteProperty("ShowTrackingHand", mShowTrackHand, True)
    Call PropBag.WriteProperty("TabStyle", tStyle, 0)

End Sub

Public Property Get TrackingColor() As OLE_COLOR
    TrackingColor = m_HotTrackColor
End Property

Public Property Let TrackingColor(ByVal vNewValue As OLE_COLOR)
    m_HotTrackColor = vNewValue
    PropertyChanged "TrackingColor"
End Property

Public Property Get TabWidth(index As Integer) As Long
    TabWidth = TabButton(index).Width
End Property

Public Property Get TabHeight(index As Integer) As Long
    TabHeight = TabButton(index).Height
End Property

Public Property Get TabTop(index As Integer) As Long
    TabTop = TabButton(index).Top
End Property

Public Property Get TabLeft(index As Integer) As Long
    TabLeft = TabButton(index).Left
End Property

Public Property Get BoldSelection() As Boolean
    BoldSelection = fBold
End Property

Public Property Let BoldSelection(ByVal vNewBold As Boolean)
    fBold = vNewBold
    PropertyChanged "BoldSelection"
    Call UpDate
End Property

Public Property Get Style3D() As Border3D
Attribute Style3D.VB_Description = "Returns/sets the line width for output from graphics methods."
    Style3D = UserControl.DrawWidth
End Property

Public Property Let Style3D(ByVal New_bStyle As Border3D)
    UserControl.DrawWidth() = New_bStyle
    PropertyChanged "Style3D"
    Call UpDate
    Call Redraw
End Property

Public Property Get Font() As Font
Attribute Font.VB_Description = "Returns a Font object."
Attribute Font.VB_UserMemId = -512
    Set Font = TabButton(0).Font
End Property

Public Property Set Font(ByVal New_Font As Font)
    Set TabButton(0).Font = New_Font
    PropertyChanged "Font"
    m_Track_Tmp = New_Font.Underline
End Property

Public Property Get ForeColor() As OLE_COLOR
Attribute ForeColor.VB_Description = "Returns/sets the foreground color used to display text and graphics in an object."
    ForeColor = TabButton(0).ForeColor
End Property

Public Property Let ForeColor(ByVal New_ForeColor As OLE_COLOR)
    TabButton(0).ForeColor() = New_ForeColor
    UserControl.ForeColor = TabButton(0).ForeColor
    PropertyChanged "ForeColor"
    m_TabForeColor = New_ForeColor
    Call UpDate
End Property

Public Property Get TrackUnderLine() As Boolean
    TrackUnderLine = m_TrackUnderline
End Property

Public Property Let TrackUnderLine(ByVal vNewValue As Boolean)
    m_TrackUnderline = vNewValue
    PropertyChanged "TrackUnderLine"
End Property

Public Property Let HightLight(ByVal index As Integer, ByVal Highlight As Boolean)
    TabsX(index).tHighlight = Highlight
End Property

Public Property Get HightLight(ByVal index As Integer) As Boolean
    HightLight = TabsX(index).tHighlight
End Property

Public Property Get TabCount() As Long
    TabCount = TabCounter + 1
End Property

Public Property Get ShowRect() As Boolean
    ShowRect = mShowRect
End Property

Public Property Let ShowRect(ByVal NewShow As Boolean)
    mShowRect = NewShow
    PropertyChanged "ShowRect"
    Call UpDate
End Property

Public Property Let ShowToolTip(ByVal NewShow As Boolean)
    mToolTip = NewShow
    PropertyChanged "ShowToolTip"
End Property

Public Property Get ShowToolTip() As Boolean
    ShowToolTip = mToolTip
End Property

Public Property Get MouseIcon() As Picture
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon."
    Set MouseIcon = TabButton(0).MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set TabButton(0).MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

Public Property Get MousePointer() As Integer
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object."
    MousePointer = TabButton(0).MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As Integer)
    TabButton(0).MousePointer() = New_MousePointer
    PropertyChanged "MousePointer"
End Property

Public Property Get ShowTrackingHand() As Boolean
    ShowTrackingHand = mShowTrackHand
End Property

Public Property Let ShowTrackingHand(ByVal vNewValue As Boolean)
    mShowTrackHand = vNewValue
    PropertyChanged "ShowTrackingHand"
End Property

Public Property Get TabStyle() As tbStyle
    TabStyle = tStyle
End Property

Public Property Let TabStyle(ByVal vNewValue As tbStyle)
    tStyle = vNewValue
    PropertyChanged "TabStyle"
    Call UpDate
End Property
