VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "DM TabControl"
   ClientHeight    =   6360
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8475
   LinkTopic       =   "Form1"
   ScaleHeight     =   6360
   ScaleWidth      =   8475
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdabout 
      Caption         =   "About"
      Height          =   495
      Left            =   6660
      TabIndex        =   30
      Top             =   3885
      Width           =   1215
   End
   Begin VB.CommandButton cmdexit 
      Caption         =   "Exit"
      Height          =   495
      Left            =   5160
      TabIndex        =   29
      Top             =   3885
      Width           =   1215
   End
   Begin Project1.sTabFx sTabFx3 
      Height          =   1170
      Left            =   195
      TabIndex        =   27
      Top             =   4845
      Width           =   3735
      _ExtentX        =   6588
      _ExtentY        =   2064
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      ShowRect        =   0   'False
      ShowToolTip     =   0   'False
      MouseIcon       =   "demo.frx":0000
      ShowTrackingHand=   0   'False
      TabStyle        =   1
      Begin VB.Label Label3 
         Caption         =   "Tabs with Buttons"
         Height          =   420
         Left            =   255
         TabIndex        =   28
         Top             =   525
         Width           =   3060
      End
   End
   Begin VB.CheckBox chkhand 
      Caption         =   "Show Hotracking Hand Cursor"
      Height          =   255
      Left            =   210
      TabIndex        =   26
      Top             =   3120
      Width           =   3555
   End
   Begin VB.CheckBox chkrect 
      Caption         =   "Show focus rect"
      Height          =   255
      Left            =   210
      TabIndex        =   25
      Top             =   3720
      Width           =   3555
   End
   Begin VB.CommandButton cmdadd 
      Caption         =   "Add new Tabs Below"
      Height          =   360
      Left            =   5160
      TabIndex        =   24
      Top             =   2160
      Width           =   2835
   End
   Begin Project1.sTabFx sTabFx2 
      Height          =   1005
      Left            =   5145
      TabIndex        =   23
      Top             =   2685
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   1984
      HotTracking     =   -1  'True
      TrackingColor   =   16711935
      BoldSelection   =   0   'False
      Style3D         =   2
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      TrackUnderLine  =   -1  'True
      MouseIcon       =   "demo.frx":0162
   End
   Begin VB.CommandButton cmdCap 
      Caption         =   "Set Caption"
      Height          =   360
      Left            =   5115
      TabIndex        =   22
      Top             =   1605
      Width           =   1350
   End
   Begin VB.TextBox txtCaption 
      Height          =   360
      Left            =   6585
      TabIndex        =   21
      Text            =   "Hello"
      Top             =   1605
      Width           =   1470
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select Tab 2"
      Height          =   375
      Left            =   195
      TabIndex        =   18
      Top             =   4140
      Width           =   1710
   End
   Begin VB.CheckBox chkunderline 
      Caption         =   "Show Hot Tracking with underline"
      Height          =   255
      Left            =   210
      TabIndex        =   17
      Top             =   2820
      Width           =   3555
   End
   Begin VB.CheckBox chkBold 
      Caption         =   "Show Selected tab captions in bold"
      Height          =   255
      Left            =   210
      TabIndex        =   16
      Top             =   3435
      Width           =   3555
   End
   Begin VB.CheckBox chkHot 
      Caption         =   "Hot Tracking"
      Height          =   255
      Left            =   210
      TabIndex        =   15
      Top             =   2535
      Width           =   3555
   End
   Begin Project1.sTabFx sTabFx1 
      Height          =   2130
      Left            =   135
      TabIndex        =   1
      Top             =   150
      Width           =   4740
      _ExtentX        =   8361
      _ExtentY        =   3757
      TrackingColor   =   255
      BoldSelection   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Verdana"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   0
      ShowRect        =   0   'False
      MouseIcon       =   "demo.frx":02C4
      ShowTrackingHand=   0   'False
      Begin VB.PictureBox PicTab 
         BackColor       =   &H8000000A&
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   3
         Left            =   4545
         ScaleHeight     =   1335
         ScaleWidth      =   2895
         TabIndex        =   11
         Top             =   1455
         Visible         =   0   'False
         Width           =   2895
         Begin VB.Label Label2 
            AutoSize        =   -1  'True
            Caption         =   "Version 1"
            BeginProperty Font 
               Name            =   "Verdana"
               Size            =   8.25
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   150
            TabIndex        =   14
            Top             =   510
            Width           =   915
         End
         Begin VB.Label Label1 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "DM TabControl Replacement"
            Height          =   195
            Left            =   150
            TabIndex        =   13
            Top             =   225
            Width           =   2070
         End
      End
      Begin VB.PictureBox PicTab 
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   2
         Left            =   4545
         ScaleHeight     =   1335
         ScaleWidth      =   2895
         TabIndex        =   7
         Top             =   1455
         Visible         =   0   'False
         Width           =   2895
         Begin VB.OptionButton Option1 
            Caption         =   "Mouse"
            Height          =   240
            Index           =   8
            Left            =   270
            TabIndex        =   10
            Top             =   210
            Width           =   1425
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Keyboard"
            Height          =   240
            Index           =   7
            Left            =   270
            TabIndex        =   9
            Top             =   510
            Width           =   1425
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Monitor"
            Height          =   240
            Index           =   6
            Left            =   270
            TabIndex        =   8
            Top             =   795
            Width           =   1425
         End
      End
      Begin VB.PictureBox PicTab 
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   1
         Left            =   4545
         ScaleHeight     =   1335
         ScaleWidth      =   2895
         TabIndex        =   6
         Top             =   1455
         Visible         =   0   'False
         Width           =   2895
         Begin VB.DirListBox Dir1 
            Height          =   765
            Left            =   105
            TabIndex        =   12
            Top             =   195
            Width           =   2445
         End
      End
      Begin VB.PictureBox PicTab 
         BorderStyle     =   0  'None
         Height          =   1335
         Index           =   0
         Left            =   4545
         ScaleHeight     =   1335
         ScaleWidth      =   2895
         TabIndex        =   2
         Top             =   1455
         Visible         =   0   'False
         Width           =   2895
         Begin VB.OptionButton Option1 
            Caption         =   "Quake IV"
            Height          =   240
            Index           =   2
            Left            =   270
            TabIndex        =   5
            Top             =   795
            Width           =   1425
         End
         Begin VB.OptionButton Option1 
            Caption         =   "DOOM III"
            Height          =   240
            Index           =   1
            Left            =   270
            TabIndex        =   4
            Top             =   510
            Width           =   1425
         End
         Begin VB.OptionButton Option1 
            Caption         =   "Duke 3D"
            Height          =   240
            Index           =   0
            Left            =   270
            TabIndex        =   3
            Top             =   210
            Width           =   1425
         End
      End
   End
   Begin VB.Label lblIndex 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tab Index:"
      Height          =   195
      Left            =   5295
      TabIndex        =   20
      Top             =   1200
      Width           =   765
   End
   Begin VB.Label lblcap 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tab Caption:"
      Height          =   195
      Left            =   5295
      TabIndex        =   19
      Top             =   900
      Width           =   915
   End
   Begin VB.Label lblkey 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Tab Key:"
      Height          =   195
      Left            =   5295
      TabIndex        =   0
      Top             =   570
      Width           =   645
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim TabIdx As Integer
Dim b As Boolean

Sub ArrangeTabs(index As Integer)
Dim x As Integer

    For x = 0 To PicTab.Count - 1
        PicTab(x).Visible = False
    Next
    
    x = 0
    
    PicTab(index).Visible = True
    PicTab(index).Left = 135
    PicTab(index).Top = 450
    
End Sub

Private Sub chkhand_Click()
    sTabFx1.ShowTrackingHand = chkhand
End Sub

Private Sub chkrect_Click()
    sTabFx1.ShowRect = chkrect
End Sub

Private Sub cmdabout_Click()
    MsgBox "Tab Replacement control by DreamVb.", vbInformation
    
End Sub

Private Sub cmdadd_Click()
    sTabFx2.AddTab "Tab " & sTabFx2.TabCount
    
End Sub

Private Sub cmdCap_Click()
   sTabFx1.TabCaption(TabIdx) = txtCaption.Text
End Sub

Private Sub cmdexit_Click()
    End
End Sub

Private Sub Command1_Click()
    sTabFx1.SelectTab 2
End Sub

Private Sub chkBold_Click()
    sTabFx1.BoldSelection = chkBold
End Sub

Private Sub chkHot_Click()
    sTabFx1.HotTracking = chkHot
End Sub

Private Sub chkunderline_Click()
    sTabFx1.TrackUnderLine = chkunderline
End Sub

Private Sub Form_Load()
    DoEvents
    sTabFx1.TabCaption(0) = "Games"
    sTabFx1.AddTab "Test Tab ", "B"
    sTabFx1.AddTab "Hardware", "C"
    sTabFx1.AddTab "About", "D"
    sTabFx1.SelectTab 1 'Select tab 1
    
    sTabFx3.AddTab "Test"
    sTabFx3.AddTab "Test"
    sTabFx3.AddTab "Test"
            
End Sub

Private Sub sTabFx1_Click(index As Integer, Key As String, Caption As String)
    ArrangeTabs index
    TabIdx = index
    
    lblkey.Caption = "Tab Key: " & Key
    lblcap.Caption = "Tab Caption: " & Caption
    lblIndex.Caption = "Tab Index: " & index
    
    If index = 2 Then sTabFx1.HightLight(index) = True
End Sub

Private Sub sTabFx1_TabMouseMove(index As Integer, Selected As Boolean, Key As String, Caption As String)
    lblkey.Caption = "Tab Key: " & Key
    lblcap.Caption = "Tab Caption: " & Caption
    lblIndex.Caption = "Tab Index: " & index
End Sub
