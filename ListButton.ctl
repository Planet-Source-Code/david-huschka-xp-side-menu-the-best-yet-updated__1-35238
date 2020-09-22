VERSION 5.00
Begin VB.UserControl ListButton 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   ClientHeight    =   390
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2805
   LockControls    =   -1  'True
   MouseIcon       =   "ListButton.ctx":0000
   ScaleHeight     =   390
   ScaleWidth      =   2805
   Begin VB.Timer tmMouse 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3270
      Top             =   795
   End
   Begin VB.Image img 
      Appearance      =   0  'Flat
      Height          =   240
      Left            =   30
      Picture         =   "ListButton.ctx":030A
      Top             =   30
      Width           =   240
   End
   Begin VB.Label Cap 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BackStyle       =   0  'Transparent
      Caption         =   "Caption"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   480
      TabIndex        =   0
      Top             =   45
      Width           =   555
   End
   Begin VB.Shape Bord 
      Height          =   300
      Left            =   0
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "ListButton"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private tPrevEvent As String
Private lState As StateConstants
Private bMouseDown As Boolean

Private Const M_DEF_AUTO = False
Private Const M_DEF_CAPTION = "ListButton1"
Private Const M_DEF_BACK = vbWhite
Private Const M_DEF_ENABLED = True
Private Const M_DEF_TOOL = vbNullString
Private Const M_DEF_TYPE = xpHyperlink

Private m_AutoSize As Boolean
Private m_Caption As String
Private m_BackColor As Long
Private m_Enabled As Boolean
Private m_Icon As Picture
Private m_Tool As String
Private m_Type As P_Style

Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Event Click()
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseEnter()
Public Event MouseExit()
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Public Event Resize()

Public Property Get Autosize() As Boolean

    Autosize = m_AutoSize
End Property

Public Property Let Autosize(a_Value As Boolean)

    m_AutoSize = a_Value
    
    PropertyChanged "AutoSize"
    UserControl_Resize
End Property

Public Property Get Caption() As String

    Caption = m_Caption
End Property

Public Property Let Caption(ByVal c_Value As String)

    m_Caption = c_Value
    
    Cap = m_Caption
        
    PropertyChanged "Caption"
    UserControl_Resize
End Property

Public Property Get BackColor() As OLE_COLOR

    BackColor = m_BackColor
End Property

Public Property Let BackColor(b_Value As OLE_COLOR)

    m_BackColor = b_Value
    UserControl.BackColor = m_BackColor
    Call DrawButton(btUp)
    
    PropertyChanged "BackColor"
End Property

Public Property Get Enabled() As Boolean

    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal e_Value As Boolean)

    m_Enabled = e_Value
    UserControl.Enabled = m_Enabled
    
    If m_Enabled Then
        lState = btUp
    End If
    Call DrawButton(lState)
    
    PropertyChanged "Enabled"
End Property

Public Property Get Icon() As Picture

    Set Icon = m_Icon
End Property

Public Property Set Icon(ByVal i_Value As Picture)

    Set m_Icon = i_Value
    Set img.Picture = m_Icon
    
    If img.Height > 240 Then
        Set m_Icon = LoadPicture("")
        Set img.Picture = m_Icon
        'MsgBox "16 X 16 Icons Only", vbCritical + vbOKOnly, "Application Error"
    End If
    
    PropertyChanged "Icon"
    UserControl_Resize
End Property

Public Property Get ToolTip() As String
    
    ToolTip = m_Tool
End Property

Public Property Let ToolTip(t_Val As String)

    m_Tool = t_Val
    
    PropertyChanged "ToolTip"
End Property

Public Property Get Style() As P_Style

    Style = m_Type
End Property

Public Property Let Style(s_Val As P_Style)

    m_Type = s_Val
    
    If m_Type = xpHyperlink Then
        UserControl.MousePointer = 99
    Else
        UserControl.MousePointer = 0
    End If
    
    PropertyChanged "Style"
    UserControl_Resize
End Property


Private Sub Cap_Click()

    UserControl_Click
End Sub

Private Sub Cap_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub Cap_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub Cap_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    UserControl_MouseUp Button, Shift, X, Y
End Sub

Private Sub img_Click()

    UserControl_Click
End Sub

Private Sub img_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    UserControl_MouseDown Button, Shift, X, Y
End Sub

Private Sub img_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    UserControl_MouseMove Button, Shift, X, Y
End Sub

Private Sub img_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    UserControl_MouseUp Button, Shift, X, Y
End Sub

Private Sub tmMouse_Timer()
Dim pnt As POINTAPI
    
    GetCursorPos pnt
    ScreenToClient UserControl.hWnd, pnt
    
    If pnt.X < 0 Or pnt.Y < 0 Or _
            pnt.X > (UserControl.Width / 15) Or _
            pnt.Y > 19 Then
        tmMouse.Enabled = False
    
        Call RaiseEventEx("MouseExit")
        
        Call DrawButton(btUp)
    Else
        If bMouseDown Then
            Call DrawButton(btDown)
        Else
            Call DrawButton(btOver)
        End If
    End If
End Sub

Private Sub UserControl_InitProperties()

    Caption = M_DEF_CAPTION
    BackColor = M_DEF_BACK
    Enabled = M_DEF_ENABLED
    Set Icon = LoadPicture("")
    Style = M_DEF_TYPE
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)

    Autosize = PropBag.ReadProperty("AutoSize", M_DEF_AUTO)
    Caption = PropBag.ReadProperty("Caption", M_DEF_CAPTION)
    BackColor = PropBag.ReadProperty("BackColor", M_DEF_BACK)
    Enabled = PropBag.ReadProperty("Enabled", M_DEF_ENABLED)
    Set Icon = PropBag.ReadProperty("Icon", Nothing)
    ToolTip = PropBag.ReadProperty("ToolTip", M_DEF_TOOL)
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("AutoSize", m_AutoSize, M_DEF_AUTO)
    Call PropBag.WriteProperty("Enabled", m_Enabled, M_DEF_ENABLED)
    Call PropBag.WriteProperty("BackColor", m_BackColor, M_DEF_BACK)
    Call PropBag.WriteProperty("Caption", m_Caption, M_DEF_CAPTION)
    Call PropBag.WriteProperty("Icon", m_Icon, Nothing)
    Call PropBag.WriteProperty("ToolTip", m_Tool, M_DEF_TOOL)
End Sub

Private Sub UserControl_Click()

    Call RaiseEventEx("Click")
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = vbLeftButton Then
        bMouseDown = True
        Call DrawButton(btDown)
    End If
    
    Call RaiseEventEx("MouseDown", Button, Shift, X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
   
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

    If X >= 0 And Y >= 0 And X <= UserControl.Width And Y <= UserControl.Height Then
        Call RaiseEventEx("MouseEnter")
        Call RaiseEventEx("MouseMove", Button, Shift, X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY)
    End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    bMouseDown = False
    
    If Button = vbLeftButton Then
        Call DrawButton(btUp)
    End If

    Call RaiseEventEx("MouseUp", Button, Shift, X * Screen.TwipsPerPixelX, Y * Screen.TwipsPerPixelY)
End Sub

Private Sub UserControl_Resize()

    UserControl.Height = 300
    Bord.Width = UserControl.Width
    If m_Icon Is Nothing Then
        Cap.Left = 150
    Else
        Cap.Left = 390
    End If
    
    If UserControl.Width < (Cap.Left + Cap.Width + 150) Or Autosize Or m_Type = xpHyperlink Then _
        UserControl.Width = Cap.Left + Cap.Width + 150
    
    Call DrawButton(btUp)
    Call RaiseEventEx("Resize")
End Sub

Private Function RaiseEventEx(ByVal Name As String, ParamArray Params() As Variant)
        
    Select Case Name
    Case "Click"
        RaiseEvent Click
    Case "MouseDown"
        RaiseEvent MouseDown(CInt(Params(0)), CInt(Params(1)), CSng(Params(2)), CSng(Params(3)))
    Case "MouseMove"
        RaiseEvent MouseMove(CInt(Params(0)), CInt(Params(1)), CSng(Params(2)), CSng(Params(3)))
    Case "MouseUp"
        RaiseEvent MouseUp(CInt(Params(0)), CInt(Params(1)), CSng(Params(2)), CSng(Params(3)))
    Case "MouseExit"
        If tPrevEvent <> "MouseExit" Then
            RaiseEvent MouseExit
        End If
        tPrevEvent = Name
    Case "MouseEnter"
        If tPrevEvent <> "MouseEnter" Then
            RaiseEvent MouseEnter
        End If
        tPrevEvent = Name
        
    Case "Resize"
        RaiseEvent Resize
    End Select
End Function

Private Function DrawButton(StateVal As StateConstants)

    Select Case StateVal
    Case btOver
        Select Case Style
        Case xpHyperlink
            If Not Cap.FontUnderline Then Cap.FontUnderline = True
        Case xpCustom
            UserControl.BackColor = RGB(173, 174, 214)
            Cap.ForeColor = IIf(Enabled, vbWhite, RGB(128, 128, 128))
            Bord.BorderColor = RGB(0, 0, 132)
        End Select
    Case btUp
        If Style = xpHyperlink Then
            If Cap.FontUnderline Then Cap.FontUnderline = False
        End If
        UserControl.BackColor = m_BackColor
        Cap.ForeColor = IIf(Enabled, RGB(33, 93, 198), RGB(128, 128, 128))
        Bord.BorderColor = UserControl.BackColor
    Case btDown
        Select Case Style
        Case xpHyperlink
            Cap.ForeColor = vbRed
            If Not Cap.FontUnderline Then Cap.FontUnderline = True
        Case xpCustom
            UserControl.BackColor = RGB(132, 130, 198)
            Cap.ForeColor = IIf(Enabled, vbBlack, RGB(128, 128, 128))
            Bord.BorderColor = RGB(0, 0, 132)
        End Select
    End Select
End Function
