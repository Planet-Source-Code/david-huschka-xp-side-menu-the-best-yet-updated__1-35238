VERSION 5.00
Begin VB.UserControl XPSideMenu 
   Alignable       =   -1  'True
   Appearance      =   0  'Flat
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000005&
   ClientHeight    =   3255
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3975
   ControlContainer=   -1  'True
   ScaleHeight     =   3255
   ScaleWidth      =   3975
   Begin XPSide.XPVScroll SideScroll 
      Align           =   4  'Align Right
      Height          =   3240
      Left            =   3795
      TabIndex        =   3
      Top             =   0
      Width           =   180
      _ExtentX        =   318
      _ExtentY        =   5715
      LargeChange     =   300
      SmallChange     =   15
   End
   Begin XPSide.SideXP Frame 
      Height          =   1515
      Index           =   0
      Left            =   615
      TabIndex        =   2
      Top             =   30
      Visible         =   0   'False
      Width           =   2820
      _extentx        =   4974
      _extenty        =   2672
      caption         =   "SideXP1"
      dropheight      =   1515
   End
   Begin VB.PictureBox Bord 
      Align           =   3  'Align Left
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   3240
      Index           =   1
      Left            =   0
      ScaleHeight     =   3240
      ScaleWidth      =   15
      TabIndex        =   0
      Top             =   0
      Width           =   15
   End
   Begin VB.PictureBox Bord 
      Align           =   2  'Align Bottom
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      FillStyle       =   0  'Solid
      ForeColor       =   &H80000008&
      Height          =   15
      Index           =   0
      Left            =   0
      ScaleHeight     =   15
      ScaleWidth      =   3975
      TabIndex        =   1
      Top             =   3240
      Width           =   3975
   End
End
Attribute VB_Name = "XPSideMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Private m_Frame_Count As Integer
Private m_Top As Integer
Private m_Scroll As Boolean

Public Event Action(Frame As Integer, Button As Integer)
Public Event ToolTipOver(Tip As String)

Private Sub Frame_Action(Index As Integer, ButtonId As Integer)

    RaiseEvent Action(Index, ButtonId)
End Sub

Private Sub Frame_RollDown(Index As Integer, Drop As Integer)
Dim iIndex As Integer

    For iIndex = Index + 1 To m_Frame_Count - 1
        Frame(iIndex).Top = Frame(iIndex).Top + Drop
    Next
    UserControl_Resize
End Sub

Private Sub Frame_RollUp(Index As Integer, Drop As Integer)
Dim iIndex As Integer

    For iIndex = Index + 1 To m_Frame_Count - 1
        Frame(iIndex).Top = Frame(iIndex).Top - Drop
    Next
    UserControl_Resize
End Sub

Private Sub Frame_ToolTipOver(Index As Integer, Tip As String)

    RaiseEvent ToolTipOver(Tip)
End Sub

Private Sub UserControl_Initialize()

    m_Frame_Count = 0
End Sub

Private Sub UserControl_Resize()
Dim iIndex As Integer
Dim Colr As Long

    DrawGrad 122, 161, 232, 99, 117, 215
    
    If m_Frame_Count > 0 Then
        If ((Frame(m_Frame_Count - 1).Top + Frame(m_Frame_Count - 1).Height + 120) - (Frame(0).Top - 30)) > UserControl.Height Then
            SideScroll.Visible = True
            m_Scroll = True
            SideScroll.Max = (((Frame(m_Frame_Count - 1).Top + Frame(m_Frame_Count - 1).Height + 120) - (Frame(0).Top - 30)) - UserControl.Height) / 15
        Else
            If SideScroll.Value = 0 Then
                SideScroll.Visible = False
                m_Scroll = False
            End If
        End If
    End If
    
    For iIndex = 0 To m_Frame_Count - 1
        Colr = UserControl.Point(Frame(iIndex).Left, Frame(iIndex).Top)
        If Colr > 0 Then Frame(iIndex).ColorMask = Colr
        Frame(iIndex).Width = UserControl.Width - 360 - IIf(m_Scroll, SideScroll.Width, 0)
    Next
End Sub

Private Function DrawGrad(Redval1 As Single, Greenval1 As Single, Blueval1 As Single, _
                      Redval2 As Single, Greenval2 As Single, Blueval2 As Single)
Dim iStep As Integer, iIndex As Integer, iLeft As Integer, iRight As Integer
Dim iRed As Single, iGreen As Single, iBlue As Single

            
    iLeft = 0
    iStep = ((UserControl.Height) / 63)
    iRight = iStep
    
    iRed = (Redval2 - Redval1) / 63
    iGreen = (Greenval2 - Greenval1) / 63
    iBlue = (Blueval2 - Blueval1) / 63
        
    For iIndex = 1 To 63
        UserControl.Line (0, iLeft)-(UserControl.Width, iRight), RGB(Redval1, Greenval1, Blueval1), BF
    
        Redval1 = Redval1 + iRed
        Greenval1 = Greenval1 + iGreen
        Blueval1 = Blueval1 + iBlue
        
        If Redval1 < Redval2 Then Redval1 = Redval2
        If Greenval1 < Greenval2 Then Greenval1 = Greenval2
        If Blueval1 < Blueval2 Then Blueval1 = Blueval2
        
        iLeft = iRight
        iRight = iLeft + iStep
    Next
End Function

Public Function AddFrame(Caption As String, Style As P_Type, State As P_State, Height As Integer, _
                         Optional Autosize As Boolean = False, Optional Icon As Picture)
Dim iIndex As Integer

    With Frame(m_Frame_Count)
    If m_Frame_Count > 0 Then
        Load Frame(m_Frame_Count)
        .Top = Frame(m_Frame_Count - 1).Top + Frame(m_Frame_Count - 1).Height + 90
    End If
        .Left = 180
        .ColorMask = UserControl.Point(.Left, .Top)
        .Caption = Caption
        .Style = Style
        .State = State
        '.Height = Height
        .Width = UserControl.Width - 360 - IIf(SideScroll.Visible, SideScroll.Width, 0)
        .Autosize = Autosize
        If Not IsMissing(Icon) Then Set .Icon = Icon
        .Visible = True
    End With
    
    m_Frame_Count = m_Frame_Count + 1
    UserControl_Resize
End Function

Public Function AddButton(FrameIndex As Integer, Caption As String, Width As Integer, Style As P_Style, _
                          Optional Autosize As Boolean = True, Optional Enabled As Boolean = True, _
                          Optional Icon As Picture, Optional ToolTip As String)
Dim iIndex As Integer

    Frame(FrameIndex).AddButton Caption, Width, Style, Autosize, Enabled, Icon, ToolTip
    If FrameIndex > 0 Then _
        Frame(FrameIndex).Top = Frame(FrameIndex - 1).Top + Frame(FrameIndex - 1).Height + 90
    UserControl_Resize
End Function

Private Sub SideScroll_Scroll()

   SideScroll_Change
End Sub

Private Sub SideScroll_Change()
Dim iIndex As Integer
Dim Colr As Long
Dim iMove As Integer

    If m_Top = SideScroll.Value Then Exit Sub  'Exit if at top or bottom of scroll bar

    If m_Top < SideScroll.Value Then
        iMove = (m_Top - SideScroll.Value) * 15
    Else
        iMove = -(SideScroll.Value - m_Top) * 15
    End If
    
    For iIndex = 0 To m_Frame_Count - 1
        Frame(iIndex).Top = Frame(iIndex).Top + iMove
        Colr = UserControl.Point(Frame(iIndex).Left, Frame(iIndex).Top)
        If Colr > 0 Then Frame(iIndex).ColorMask = Colr
    Next
    
    m_Top = SideScroll.Value
    
    If m_Frame_Count > 0 Then
        With Frame(m_Frame_Count - 1)
            If (.Top + .Height + 120) > UserControl.Height Then
                SideScroll.Visible = True
            Else
                If SideScroll.Value = 0 And SideScroll.Visible Then
                    SideScroll.Visible = False
                    For iIndex = 0 To m_Frame_Count - 1
                        Frame(iIndex).Width = UserControl.Width - 360 - IIf(SideScroll.Visible, SideScroll.Width, 0)
                    Next
                End If
            End If
        End With
    End If
End Sub

