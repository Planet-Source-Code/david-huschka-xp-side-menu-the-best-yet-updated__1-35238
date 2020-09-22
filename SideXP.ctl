VERSION 5.00
Begin VB.UserControl SideXP 
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   4140
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4845
   LockControls    =   -1  'True
   ScaleHeight     =   4140
   ScaleWidth      =   4845
   Begin VB.Timer tmMouse 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2280
      Top             =   2220
   End
   Begin VB.PictureBox pic 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   15
      Index           =   6
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   4845
      TabIndex        =   9
      Top             =   4125
      Width           =   4845
   End
   Begin VB.PictureBox pic 
      Align           =   4  'Align Right
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3570
      Index           =   5
      Left            =   4830
      ScaleHeight     =   3570
      ScaleWidth      =   15
      TabIndex        =   8
      Top             =   555
      Width           =   15
   End
   Begin VB.PictureBox pic 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3570
      Index           =   4
      Left            =   0
      ScaleHeight     =   3570
      ScaleWidth      =   15
      TabIndex        =   7
      Top             =   555
      Width           =   15
   End
   Begin VB.Timer tDrop 
      Left            =   525
      Top             =   2010
   End
   Begin XPSide.ListButton Button 
      Height          =   300
      Index           =   0
      Left            =   315
      TabIndex        =   6
      Top             =   765
      Visible         =   0   'False
      Width           =   1350
      _ExtentX        =   1958
      _ExtentY        =   529
   End
   Begin VB.PictureBox picHead 
      Align           =   1  'Align Top
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   555
      Left            =   0
      MouseIcon       =   "SideXP.ctx":0000
      ScaleHeight     =   555
      ScaleWidth      =   4845
      TabIndex        =   0
      Top             =   0
      Width           =   4845
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Index           =   3
         Left            =   3000
         ScaleHeight     =   30
         ScaleWidth      =   15
         TabIndex        =   5
         Top             =   180
         Width           =   15
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   2
         Left            =   3030
         ScaleHeight     =   15
         ScaleWidth      =   30
         TabIndex        =   4
         Top             =   180
         Width           =   30
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   30
         Index           =   1
         Left            =   0
         ScaleHeight     =   30
         ScaleWidth      =   15
         TabIndex        =   3
         Top             =   180
         Width           =   15
      End
      Begin VB.PictureBox pic 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   15
         Index           =   0
         Left            =   0
         ScaleHeight     =   15
         ScaleWidth      =   30
         TabIndex        =   2
         Top             =   180
         Width           =   30
      End
      Begin VB.Label lblCap 
         Appearance      =   0  'Flat
         AutoSize        =   -1  'True
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         Caption         =   "Caption"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   195
         Left            =   210
         TabIndex        =   1
         Top             =   270
         Width           =   645
      End
      Begin VB.Image imgHead 
         Appearance      =   0  'Flat
         Height          =   375
         Left            =   915
         Top             =   180
         Width           =   1590
      End
      Begin VB.Image Ico 
         Appearance      =   0  'Flat
         Height          =   480
         Left            =   15
         Top             =   30
         Width           =   480
      End
      Begin VB.Shape Mask 
         BorderColor     =   &H00FFFFFF&
         BorderStyle     =   0  'Transparent
         FillColor       =   &H00FFFFFF&
         FillStyle       =   0  'Solid
         Height          =   195
         Left            =   0
         Top             =   0
         Width           =   3750
      End
   End
End
Attribute VB_Name = "SideXP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const M_DEF_AUTO = False
Private Const M_DEF_CAPTION = "SideXP1"
Private Const M_DEF_STATE = vbOpen
Private Const M_DEF_STYLE = vbMain
Private Const M_DEF_HEIGHT = 2400
Private Const M_MIN_WIDTH = 2400

Private m_Auto As Boolean
Private m_Caption As String
Private m_Color As Long
Private m_Height As Integer
Private m_Icon As Picture
Private m_State As P_State
Private m_Style As P_Type
Private m_Button_Count As Integer
Private m_Prev_Tip As Integer
Private m_Drop As Boolean
Private m_Dropn As Boolean
Private m_Pic As Integer

Private tPrevEvent As String
Private lState As StateConstants
Private bMouseDown As Boolean

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Event Action(ButtonId As Integer)
Public Event RollUp(Drop As Integer)
Public Event RollDown(Drop As Integer)
Public Event ToolTipOver(Tip As String)

Private Sub Button_Click(Index As Integer)

    RaiseEvent Action(Index)
End Sub

Private Sub Button_MouseEnter(Index As Integer)

    RaiseEvent ToolTipOver(Button(Index).ToolTip)
    m_Prev_Tip = Index
End Sub

Private Sub Button_MouseExit(Index As Integer)

    If m_Prev_Tip = Index Then RaiseEvent ToolTipOver(vbNullString)
End Sub

Private Sub Ico_Click()

    picHead_Click
End Sub

Private Sub Ico_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    picHead_MouseDown Button, Shift, x, y
End Sub

Private Sub Ico_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    picHead_MouseMove Button, Shift, x, y
End Sub

Private Sub Ico_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    picHead_MouseUp Button, Shift, x, y
End Sub

Private Sub imgHead_Click()

    picHead_Click
End Sub

Private Sub imgHead_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    picHead_MouseDown Button, Shift, x, y
End Sub

Private Sub imgHead_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    picHead_MouseMove Button, Shift, x, y
End Sub

Private Sub imgHead_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    picHead_MouseUp Button, Shift, x, y
End Sub

Private Sub lblCap_Click()

    picHead_Click
End Sub

Private Sub lblCap_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    picHead_MouseDown Button, Shift, x, y
End Sub

Private Sub lblCap_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

    picHead_MouseMove Button, Shift, x, y
End Sub

Private Sub lblCap_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    picHead_MouseUp Button, Shift, x, y
End Sub

Private Sub picHead_Click()
Dim pnt As POINTAPI
    
    GetCursorPos pnt
    ScreenToClient picHead.hWnd, pnt
    
    If pnt.x < 0 Or pnt.y < 12 Or pnt.x > (UserControl.Width / 15) Or pnt.y > 38 Then
        Exit Sub
    Else
        State = IIf(State = vbClose, vbOpen, vbClose)
    End If
End Sub

Private Sub picHead_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = vbLeftButton And y > 180 Then
        bMouseDown = True
        Call DrawButton(btDown)
    End If
End Sub

Private Sub picHead_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
   
    If Not tmMouse.Enabled Then
        tmMouse.Enabled = True
    
    ElseIf Button = 0 Then
        If lState <> btOver Then
            Call DrawButton(btOver)
        End If

    ElseIf Button = vbLeftButton Then
        If lState <> btDown Then
            Call DrawButton(btDown)
        End If
    End If
End Sub

Private Sub picHead_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    bMouseDown = False
    
    If Button = vbLeftButton Then
        Call DrawButton(btUp)
    End If
End Sub

Private Sub UserControl_InitProperties()
Dim iIndex As Integer

    m_Auto = M_DEF_AUTO
    m_Caption = M_DEF_CAPTION
    Set m_Icon = Nothing
    m_State = M_DEF_STATE
    m_Style = M_DEF_STYLE
    m_Button_Count = 0
    m_Height = M_DEF_HEIGHT
    For iIndex = 0 To m_Button_Count - 1
        Button(iIndex).BackColor = UserControl.BackColor
    Next
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("AutoSize", m_Auto, M_DEF_AUTO)
    Call PropBag.WriteProperty("Caption", m_Caption, Nothing)
    Call PropBag.WriteProperty("ColorMask", m_Color, vbWhite)
    Call PropBag.WriteProperty("DropHeight", m_Height, M_DEF_HEIGHT)
    Call PropBag.WriteProperty("Icon", m_Icon, Nothing)
    Call PropBag.WriteProperty("Style", m_Style, M_DEF_STYLE)
    Call PropBag.WriteProperty("State", m_State, M_DEF_STATE)
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Autosize = PropBag.ReadProperty("AutoSize", M_DEF_AUTO)
    Caption = PropBag.ReadProperty("Caption", M_DEF_CAPTION)
    ColorMask = PropBag.ReadProperty("ColorMask", vbWhite)
    DropHeight = PropBag.ReadProperty("DropHeight", M_DEF_HEIGHT)
    Set Icon = PropBag.ReadProperty("Icon", Nothing)
    Style = PropBag.ReadProperty("Style", M_DEF_STYLE)
    State = PropBag.ReadProperty("State", M_DEF_STATE)
End Sub

Public Property Get Autosize() As Boolean

    Autosize = m_Auto
End Property

Public Property Let Autosize(a_Val As Boolean)

    m_Auto = a_Val
    PropertyChanged "AutoSize"
    UserControl_Resize
End Property

Public Property Get Caption() As String

    Caption = m_Caption
End Property

Public Property Let Caption(s_Val As String)

    m_Caption = s_Val
    lblCap = m_Caption
    
    PropertyChanged "Caption"
End Property

Public Property Get DropHeight() As Integer

    DropHeight = m_Height
End Property

Public Property Let DropHeight(d_Value As Integer)

    m_Height = d_Value
    UserControl.Height = m_Height
    
    PropertyChanged "DropHeight"
End Property

Public Property Get Icon() As Picture

    Set Icon = m_Icon
End Property

Public Property Set Icon(ByVal i_Value As Picture)

    Set m_Icon = i_Value
    Set Ico.Picture = m_Icon
       
    If m_Icon Is Nothing Then
        lblCap.Left = 210
    Else
        lblCap.Left = 540
    End If
    
    PropertyChanged "Icon"
    UserControl_Resize
End Property

Public Property Get Style() As P_Type

    Style = m_Style
End Property

Public Property Let Style(p_Val As P_Type)

    m_Style = p_Val
    
    UserControl_Resize
    PropertyChanged "Style"
End Property

Public Property Get State() As P_State

    State = m_State
End Property

Public Property Let State(p_Val As P_State)
Dim iIndex As Integer

    m_State = p_Val
    UserControl_Resize
    For iIndex = 0 To m_Button_Count - 1
        Button(iIndex).BackColor = UserControl.BackColor
    Next
    
    PropertyChanged "State"
End Property

Public Property Get ColorMask() As OLE_COLOR
    
    ColorMask = m_Color
End Property

Public Property Let ColorMask(p_Color As OLE_COLOR)
Dim iIndex As Integer

    If p_Color < 0 Then Exit Property

    m_Color = p_Color
    
    For iIndex = 0 To 3
        pic(iIndex).BackColor = m_Color
    Next
    PropertyChanged "ColorMask"
    Mask.FillColor = m_Color
End Property

Private Sub UserControl_Resize()
Dim iIndex As Integer

    If UserControl.Width < M_MIN_WIDTH Then UserControl.Width = M_MIN_WIDTH
    If m_Height = 0 Then m_Height = UserControl.Height
    
    Select Case m_Style
    Case vbMain
        Select Case m_State
        Case vbOpen
            If Not m_Dropn Then
                imgHead.Picture = LoadResPicture(103, 0)
                m_Pic = 103
            End If
            If UserControl.Height <> m_Height Then
                m_Drop = False
                tDrop.Interval = 1
            End If
        Case vbClose
            If Not m_Dropn Then
                imgHead.Picture = LoadResPicture(104, 0)
                m_Pic = 104
            End If
            If UserControl.Height <> picHead.Height Then
                m_Drop = True
                tDrop.Interval = 1
            End If
        End Select
        If Not m_Dropn Then
            picHead.BackColor = &HC45518
            UserControl.BackColor = RGB(239, 243, 255)
            lblCap.ForeColor = vbWhite
        End If
    Case vbSub
        Select Case m_State
        Case vbOpen
            If Not m_Dropn Then
                imgHead.Picture = LoadResPicture(101, 0)
                m_Pic = 101
            End If
            If UserControl.Height <> m_Height Then
                m_Drop = False
                tDrop.Interval = 1
            End If
        Case vbClose
            If Not m_Dropn Then
                imgHead.Picture = LoadResPicture(102, 0)
                m_Pic = 102
            End If
            If UserControl.Height <> picHead.Height Then
                m_Drop = True
                tDrop.Interval = 1
            End If
        End Select
        If Not m_Dropn Then
            picHead.BackColor = vbWhite
            UserControl.BackColor = RGB(215, 222, 248)
            lblCap.ForeColor = RGB(33, 91, 199)
        End If
    End Select

    imgHead.Left = picHead.Width - imgHead.Width
    pic(2).Left = picHead.Width - 30
    pic(3).Left = picHead.Width - 15
End Sub

Public Function AddButton(Caption As String, Width As Integer, Style As P_Style, _
                          Optional Autosize As Boolean = True, _
                          Optional Enabled As Boolean = True, _
                          Optional Icon As Picture, Optional Tool As String)
                          
    If m_Button_Count > 0 Then
        Load Button(m_Button_Count)
    End If
    Button(m_Button_Count).Left = 150
    Button(m_Button_Count).Top = (m_Button_Count * 300) + 600
    Button(m_Button_Count).Caption = Caption
    Button(m_Button_Count).BackColor = UserControl.BackColor
    Button(m_Button_Count).Width = Width
    Button(m_Button_Count).Autosize = Autosize
    Button(m_Button_Count).Enabled = Enabled
    If Not IsMissing(Icon) Then Set Button(m_Button_Count).Icon = Icon
    If Not IsMissing(Tool) Then Button(m_Button_Count).ToolTip = Tool
    Button(m_Button_Count).Visible = True
    Button(m_Button_Count).Style = Style
    
    m_Button_Count = m_Button_Count + 1
    
    If m_Auto Then
        DropHeight = (m_Button_Count * 300) + 645
    End If
End Function

Private Sub tDrop_Timer()

    If m_Drop Then
        If UserControl.Height - 60 > picHead.Height Then
            UserControl.Height = UserControl.Height - 60
            RaiseEvent RollUp(60)
            m_Dropn = True
        Else
            UserControl.Height = picHead.Height
            RaiseEvent RollUp(60)
            pic(6).Visible = False
            tDrop.Interval = 0
            m_Dropn = False
        End If
    Else
        If UserControl.Height + 60 < m_Height Then
            pic(6).Visible = True
            UserControl.Height = UserControl.Height + 60
            RaiseEvent RollDown(60)
            m_Dropn = True
        Else
            UserControl.Height = m_Height
            RaiseEvent RollDown(60)
            tDrop.Interval = 0
            m_Dropn = False
        End If
    End If
End Sub

Private Sub tmMouse_Timer()
Dim pnt As POINTAPI
    
    GetCursorPos pnt
    ScreenToClient picHead.hWnd, pnt
    
    If pnt.x < 0 Or pnt.y < 12 Or _
            pnt.x > (UserControl.Width / 15) Or _
            pnt.y > 38 Then
        tmMouse.Enabled = False
        
        Call DrawButton(btUp)
    Else
        If bMouseDown Then
            Call DrawButton(btDown)
        Else
            Call DrawButton(btOver)
        End If
    End If
End Sub

Private Function DrawButton(StateVal As StateConstants)

    Select Case StateVal
    Case btOver
        If lblCap.ForeColor <> vbRed Then lblCap.ForeColor = vbRed
        picHead.MousePointer = 99
        Select Case m_Style
        Case vbMain
            Select Case m_State
            Case vbOpen
                If Not m_Dropn Then
                    imgHead.Picture = LoadResPicture(106, 0)
                    m_Pic = 106
                End If
            Case vbClose
                If Not m_Dropn Then
                    imgHead.Picture = LoadResPicture(105, 0)
                    m_Pic = 105
                End If
            End Select
        Case vbSub
            Select Case m_State
            Case vbOpen
                If Not m_Dropn Then
                    imgHead.Picture = LoadResPicture(108, 0)
                    m_Pic = 108
                End If
            Case vbClose
                If Not m_Dropn Then
                    imgHead.Picture = LoadResPicture(107, 0)
                    m_Pic = 107
                End If
            End Select
        End Select
    Case btUp
        picHead.MousePointer = 0
        Select Case m_Style
        Case vbMain
            If lblCap.ForeColor <> vbWhite Then lblCap.ForeColor = vbWhite
            Select Case m_State
            Case vbOpen
                If Not m_Dropn Then
                    imgHead.Picture = LoadResPicture(103, 0)
                    m_Pic = 103
                End If
            Case vbClose
                If Not m_Dropn Then
                    imgHead.Picture = LoadResPicture(104, 0)
                    m_Pic = 104
                End If
            End Select
        Case vbSub
            If lblCap.ForeColor <> RGB(33, 91, 199) Then lblCap.ForeColor = RGB(33, 91, 199)
            Select Case m_State
            Case vbOpen
                If Not m_Dropn Then
                    imgHead.Picture = LoadResPicture(101, 0)
                    m_Pic = 101
                End If
            Case vbClose
                If Not m_Dropn Then
                    imgHead.Picture = LoadResPicture(102, 0)
                    m_Pic = 102
                End If
            End Select
        End Select
    Case btDown
        If lblCap.ForeColor <> vbRed Then lblCap.ForeColor = vbRed
        picHead.MousePointer = 99
        Select Case m_Style
        Case vbMain
            Select Case m_State
            Case vbOpen
                If Not m_Dropn Then
                    imgHead.Picture = LoadResPicture(106, 0)
                    m_Pic = 106
                End If
            Case vbClose
                If Not m_Dropn Then
                    imgHead.Picture = LoadResPicture(105, 0)
                    m_Pic = 105
                End If
            End Select
        Case vbSub
            Select Case m_State
            Case vbOpen
                If Not m_Dropn Then
                    imgHead.Picture = LoadResPicture(108, 0)
                    m_Pic = 108
                End If
            Case vbClose
                If Not m_Dropn Then
                    imgHead.Picture = LoadResPicture(107, 0)
                    m_Pic = 107
                End If
            End Select
        End Select
    End Select
End Function

