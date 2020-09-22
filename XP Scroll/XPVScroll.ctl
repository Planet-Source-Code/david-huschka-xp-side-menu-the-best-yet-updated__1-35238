VERSION 5.00
Begin VB.UserControl XPVScroll 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   240
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   FontTransparent =   0   'False
   HasDC           =   0   'False
   LockControls    =   -1  'True
   Picture         =   "XPVScroll.ctx":0000
   ScaleHeight     =   3600
   ScaleWidth      =   240
   ToolboxBitmap   =   "XPVScroll.ctx":1E7C2
   Begin VB.PictureBox imgBarPress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   330
      Left            =   1275
      Picture         =   "XPVScroll.ctx":1EAD4
      ScaleHeight     =   330
      ScaleWidth      =   240
      TabIndex        =   11
      Top             =   1290
      Width           =   240
   End
   Begin VB.PictureBox imgDownPress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   225
      Left            =   885
      Picture         =   "XPVScroll.ctx":29DC6
      ScaleHeight     =   225
      ScaleWidth      =   300
      TabIndex        =   10
      Top             =   3000
      Width           =   300
   End
   Begin VB.PictureBox imgUpPress 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   930
      Picture         =   "XPVScroll.ctx":2A17C
      ScaleHeight     =   210
      ScaleWidth      =   180
      TabIndex        =   9
      Top             =   2595
      Width           =   180
   End
   Begin VB.PictureBox imgBarHot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   1260
      Picture         =   "XPVScroll.ctx":2A532
      ScaleHeight     =   195
      ScaleWidth      =   270
      TabIndex        =   8
      Top             =   2055
      Width           =   270
   End
   Begin VB.PictureBox imgBar 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   240
      Left            =   465
      Picture         =   "XPVScroll.ctx":35824
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   765
      Width           =   240
   End
   Begin VB.PictureBox imgDownHot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   390
      Left            =   735
      Picture         =   "XPVScroll.ctx":40816
      ScaleHeight     =   390
      ScaleWidth      =   330
      TabIndex        =   6
      Top             =   2010
      Width           =   330
   End
   Begin VB.PictureBox imgDown 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   795
      Picture         =   "XPVScroll.ctx":40BCC
      ScaleHeight     =   255
      ScaleWidth      =   240
      TabIndex        =   5
      Top             =   1485
      Width           =   240
   End
   Begin VB.PictureBox imgUpHot 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   285
      Left            =   810
      Picture         =   "XPVScroll.ctx":40F82
      ScaleHeight     =   285
      ScaleWidth      =   285
      TabIndex        =   4
      Top             =   990
      Width           =   285
   End
   Begin VB.PictureBox imgUp 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   210
      Left            =   810
      Picture         =   "XPVScroll.ctx":41338
      ScaleHeight     =   210
      ScaleWidth      =   210
      TabIndex        =   3
      Top             =   510
      Width           =   210
   End
   Begin VB.Timer tScrl 
      Enabled         =   0   'False
      Interval        =   500
      Left            =   -435
      Top             =   1560
   End
   Begin VB.PictureBox Bar 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   0
      Picture         =   "XPVScroll.ctx":416EE
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   2
      Top             =   675
      Width           =   255
      Begin VB.Image Bottom 
         Appearance      =   0  'Flat
         Height          =   45
         Left            =   0
         Picture         =   "XPVScroll.ctx":4C6E0
         Top             =   180
         Width           =   225
      End
      Begin VB.Image Top 
         Appearance      =   0  'Flat
         Height          =   45
         Left            =   0
         Picture         =   "XPVScroll.ctx":4C7B2
         Top             =   0
         Width           =   225
      End
   End
   Begin VB.PictureBox Up 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   -15
      Picture         =   "XPVScroll.ctx":4C884
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   1
      Top             =   0
      Width           =   255
   End
   Begin VB.PictureBox Down 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   255
      Left            =   -15
      Picture         =   "XPVScroll.ctx":4CC3A
      ScaleHeight     =   255
      ScaleWidth      =   255
      TabIndex        =   0
      Top             =   2385
      Width           =   255
   End
End
Attribute VB_Name = "XPVScroll"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private Const DEFAULT_BAR_WIDTH = 225
Private Const MIN_BAR_VALUE = 0
Private Const MAX_BAR_VALUE = 32767
Private Const DEFAULT_LRG_VALUE = 1
Private Const DEFAULT_SML_VALUE = 1
Private Const RIGHT_MOUSE = 1
Private Const DEFAULT_SCROLL = 500
Private Const FAST_SCROLL = 50

Private Enum SDirection
    vbUp = 0
    vbDown = 1
End Enum

Private Enum SSize
    vbSmall = 0
    vbLarge = 1
End Enum

Private bSlideBar As Boolean
Private SlideDirection As SDirection
Private PrevYPosition As Integer
Private bLoaded As Boolean

Private m_Min As Long
Private m_Max As Long
Private m_Value As Long
Private m_Small As Long
Private m_Large As Long
Private m_Real_Height As Integer

Public Event Change()
Public Event Scroll()

Private Sub UserControl_Initialize()

    m_Large = DEFAULT_LRG_VALUE
    m_Min = MIN_BAR_VALUE
    m_Max = MAX_BAR_VALUE
    m_Small = DEFAULT_SML_VALUE
    m_Value = MIN_BAR_VALUE
    InitColors

    tScrl.Interval = DEFAULT_SCROLL
    
    Set Up_ = imgUp
    Set UpHot_ = imgUpHot
    Set UpPres_ = imgUpPress
    Set Down_ = imgDown
    Set DownHot_ = imgDownHot
    Set DownPres_ = imgDownPress
    Set Bar_ = imgBar
    Set BarHot_ = imgBarHot
    Set BarPres_ = imgBarPress
    Subclass Up, 0
    Subclass Down, 1
    Subclass Bar, 2
    
    UserControl_Resize
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
   
    With PropBag
        UserControl.Enabled = .ReadProperty("Enabled", True)
        LargeChange = .ReadProperty("LargeChange", DEFAULT_LRG_VALUE)
        Max = .ReadProperty("Max", MAX_BAR_VALUE)
        Min = .ReadProperty("Min", MIN_BAR_VALUE)
        SmallChange = .ReadProperty("SmallChange", DEFAULT_SML_VALUE)
        Value = .ReadProperty("Value", MIN_BAR_VALUE)
    End With
    bLoaded = True
    UserControl_Resize
End Sub

Private Sub UserControl_Terminate()

    UnSubclass Up
    UnSubclass Down
    UnSubclass Bar
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
   
    With PropBag
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "LargeChange", LargeChange, DEFAULT_LRG_VALUE
        .WriteProperty "Min", Min, MIN_BAR_VALUE
        .WriteProperty "Max", Max, MAX_BAR_VALUE
        .WriteProperty "SmallChange", SmallChange, DEFAULT_SML_VALUE
        .WriteProperty "Value", Value, MIN_BAR_VALUE
    End With
    UserControl_Resize
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        bSlideBar = True
        If Y > Bar.Top + Bar.Height Then
            SlideBar vbDown, vbLarge
            SlideDirection = vbDown
            tScrl.Enabled = True
        Else
            SlideBar vbUp, vbLarge
            SlideDirection = vbUp
            tScrl.Enabled = True
        End If
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    tScrl.Enabled = False
    tScrl.Interval = DEFAULT_SCROLL
    bSlideBar = False
End Sub

Private Sub UserControl_Resize()
    
    If Not bLoaded Then Exit Sub
    
    UserControl.Width = DEFAULT_BAR_WIDTH
    
    If UserControl.Height < Up.Height + Down.Height + BAR_MIN_HEIGHT Then _
        UserControl.Height = Up.Height + Down.Height + BAR_MIN_HEIGHT
    
    SetBarHeight
End Sub

Private Sub Bar_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        bSlideBar = True
        PrevYPosition = Y
    End If
End Sub

Private Sub Bar_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        bSlideBar = False
        If Value > MIN_BAR_VALUE Then
            Bar.Top = (Value * ((UserControl.Height - (Up.Height * 2) - Bar.Height) / m_Max)) + Up.Height
        Else
            Bar.Top = Up.Height
        End If
        
        RaiseEvent Change
    End If
End Sub

Private Sub Bar_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim iIndex As Integer

    If bSlideBar And Button = 1 Then
        If Bar.Top + (Y - PrevYPosition) < Up.Height Then
            Bar.Top = Up.Height
            Value = MIN_BAR_VALUE
        ElseIf Bar.Top + (Y - PrevYPosition) + Bar.Height > UserControl.Height - Down.Height Then
            Bar.Top = UserControl.Height - Bar.Height - Down.Height
            Value = m_Max
        Else
            Bar.Top = Bar.Top + (Y - PrevYPosition)
            Value = (Bar.Top - Up.Height) / ((UserControl.Height - (Up.Height * 2) - Bar.Height) / m_Max)
        End If
        
        RaiseEvent Scroll
    End If
End Sub

Private Sub Up_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        SlideBar vbUp, vbSmall
        SlideDirection = vbUp
        tScrl.Enabled = True
    End If
End Sub

Private Sub Up_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        tScrl.Enabled = False
        tScrl.Interval = DEFAULT_SCROLL
    End If
End Sub

Private Sub Up_DblClick()

    SlideBar vbUp, vbSmall
    SlideDirection = vbUp
    tScrl.Enabled = True
End Sub

Private Sub Down_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        SlideBar vbDown, vbSmall
        SlideDirection = vbDown
        tScrl.Enabled = True
    End If
End Sub

Private Sub Down_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 1 Then
        tScrl.Enabled = False
        tScrl.Interval = DEFAULT_SCROLL
    End If
End Sub

Private Sub Down_DblClick()

    SlideBar vbDown, vbSmall
    SlideDirection = vbDown
    tScrl.Enabled = True
End Sub

Private Sub tScrl_Timer()

    'If Not MouseOver Then Exit Sub
    Select Case SlideDirection
    Case vbUp
        If tScrl.Interval = FAST_SCROLL Then
            If bSlideBar Then
                UserControl_MouseDown RIGHT_MOUSE, 0, 0, 0
            Else
                Up_MouseDown RIGHT_MOUSE, 0, 0, 0
            End If
        End If
    Case vbDown
        If tScrl.Interval = FAST_SCROLL Then
            If bSlideBar Then
                UserControl_MouseDown RIGHT_MOUSE, 0, 0, UserControl.Height
            Else
                Down_MouseDown RIGHT_MOUSE, 0, 0, 0
            End If
        End If
    End Select
    tScrl.Interval = FAST_SCROLL
End Sub

'--------------------------------------------------------------------------------------------------

'       Public Property Declarations

'--------------------------------------------------------------------------------------------------

Public Property Let Value(vVal As Long)

    If vVal >= m_Min And vVal <= m_Max Then
        m_Value = vVal
    ElseIf vVal < m_Min Then
        m_Value = m_Min
    ElseIf vVal > m_Max Then
        m_Value = m_Max
    End If
    
    RaiseEvent Change
    If Not bSlideBar Then SetBarHeight
    
    PropertyChanged "Value"
End Property

Public Property Get Value() As Long

    Value = m_Value
End Property


Public Property Let SmallChange(sVal As Long)

    If sVal >= MIN_BAR_VALUE And sVal <= MAX_BAR_VALUE Then
       m_Small = sVal
    Else
       MsgBox "Invalid property value", vbCritical
       m_Small = DEFAULT_SML_VALUE
    End If
    
    PropertyChanged "SmallChange"
End Property

Public Property Get SmallChange() As Long

    SmallChange = m_Small
End Property

Public Property Let LargeChange(lVal As Long)

    If lVal >= MIN_BAR_VALUE And lVal <= MAX_BAR_VALUE Then
       m_Large = lVal
    Else
       MsgBox "Invalid property value", vbCritical
       m_Large = DEFAULT_LRG_VALUE
    End If
    
    UserControl_Resize
    PropertyChanged "LargeChange"
End Property

Public Property Get LargeChange() As Long

    LargeChange = m_Large
End Property

Public Property Let Min(mVal As Long)

    If mVal < MIN_BAR_VALUE Then mVal = MIN_BAR_VALUE
    If mVal > MAX_BAR_VALUE Then mVal = MIN_BAR_VALUE
    
    m_Min = mVal
    
    Value = IIf(m_Value < m_Min, m_Min, m_Value)
    
    UserControl_Resize
    PropertyChanged "Min"
End Property

Public Property Get Min() As Long

    Min = m_Min
End Property

Public Property Let Max(mVal As Long)
    
    If mVal > MAX_BAR_VALUE Then mVal = MAX_BAR_VALUE
    If mVal < MIN_BAR_VALUE Then mVal = MAX_BAR_VALUE
   
    m_Max = mVal

    Value = IIf(m_Value > m_Max, m_Max, m_Value)
   
    UserControl_Resize
    PropertyChanged "Max"
End Property

Public Property Get Max() As Long

    Max = m_Max
End Property

Public Property Get Enabled() As Boolean

    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal bEnabled As Boolean)

    UserControl.Enabled() = bEnabled

    PropertyChanged "Enabled"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long

    hWnd = UserControl.hWnd
End Property

'------------------------------------------------------------------------------------------------

'       Private Function Declarations

'------------------------------------------------------------------------------------------------

Private Function SlideBar(Direction As SDirection, Size As SSize)
Dim SizeVal As Integer
Dim BarTop As Integer

    Select Case Size
    Case vbSmall
        SizeVal = m_Small
    Case vbLarge
        SizeVal = m_Large
    End Select
    
    Select Case Direction
    Case vbUp
        If Value - SizeVal >= MIN_BAR_VALUE Then
            Value = Value - SizeVal
        Else
            Value = MIN_BAR_VALUE
        End If
        BarTop = ((Value / SizeVal) * (m_Real_Height * IIf(SizeVal = m_Large, 1, (m_Small / m_Large)))) + Up.Height
        If BarTop < Up.Height Then
            Bar.Top = Up.Height
        Else
            Bar.Top = BarTop
        End If
        If BarTop >= UserControl.Height - Down.Height - Bar.Height Then
            Bar.Top = UserControl.Height - Bar.Height - Down.Height
        Else
            Bar.Top = BarTop
        End If
    Case vbDown
        If Value + SizeVal <= m_Max Then
            Value = Value + SizeVal
        Else
            Value = m_Max
        End If
        BarTop = ((Value / SizeVal) * (m_Real_Height * IIf(SizeVal = m_Large, 1, (m_Small / m_Large)))) + Up.Height
        If BarTop >= UserControl.Height - Down.Height - Bar.Height Then
            Bar.Top = UserControl.Height - Bar.Height - Down.Height
        Else
            Bar.Top = BarTop
        End If
    End Select
End Function

Private Function SetBarHeight()

    If m_Large <= m_Max Then
        Bar.Height = (m_Large / (m_Large + m_Max)) * (UserControl.Height - Up.Height - Down.Height)
    Else
        Bar.Height = (1 - (m_Max / (m_Large + m_Max))) * (UserControl.Height - Up.Height - Down.Height)
    End If
    
    m_Real_Height = Bar.Height
    If Bar.Height < BAR_MIN_HEIGHT Then Bar.Height = BAR_MIN_HEIGHT
    
        Down.Top = UserControl.Height - Down.Height
    If Value > MIN_BAR_VALUE Then
        Bar.Top = (Value * ((UserControl.Height - (Up.Height * 2) - Bar.Height) / m_Max)) + Up.Height
    Else
        Bar.Top = Up.Height
    End If
    Bottom.Top = Bar.Height - Bottom.Height
    
    PaintBar Bar
End Function

Private Function InitColors()

    BarMidColor = Array(16234124, 16758412, 16232076, 16234124, 16234124, 16234132)
End Function
