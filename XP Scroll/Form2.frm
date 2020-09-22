VERSION 5.00
Object = "{975C8BC3-2A3C-11D6-8CED-00B0D091BA0C}#7.0#0"; "xpscroll.ocx"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Height          =   2625
      Left            =   705
      ScaleHeight     =   2565
      ScaleWidth      =   2625
      TabIndex        =   1
      Top             =   345
      Width           =   2685
   End
   Begin XPVertScroll.XPVScroll XPVScroll1 
      Align           =   4  'Align Right
      Height          =   3195
      Left            =   4455
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   225
      _ExtentX        =   397
      _ExtentY        =   5636
      LargeChange     =   10
      Max             =   100
      SmallChange     =   10
   End
   Begin VB.Label Label1 
      Appearance      =   0  'Flat
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   195
      Left            =   180
      TabIndex        =   2
      Top             =   90
      Width           =   45
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Resize()


    If ScaleHeight < Picture1.Top + Picture1.Height Then
        XPVScroll1.Visible = True
    Else
        XPVScroll1.Visible = False
    End If
    Label1 = XPVScroll1.Value
End Sub

Private Sub XPVScroll1_Change()

    Label1 = XPVScroll1.Value
End Sub
