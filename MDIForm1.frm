VERSION 5.00
Object = "{1103DCBC-6CBA-11D6-8D03-00B0D091BA0C}#3.0#0"; "xpside.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   6510
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8415
   LinkTopic       =   "MDIForm1"
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   6135
      Width           =   8415
      _ExtentX        =   14843
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5292
            MinWidth        =   5292
         EndProperty
      EndProperty
   End
   Begin XPSide.XPSideMenu XPSideMenu1 
      Align           =   3  'Align Left
      Height          =   6135
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3075
      _ExtentX        =   5424
      _ExtentY        =   10821
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   0
      Top             =   1305
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0B34
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":10CE
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1668
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":1C02
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":219C
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   225
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":2736
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":4440
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":614A
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub MDIForm_Load()
    XPSideMenu1.AddFrame "XP Side Menu 1", vbMain, vbOpen, 1500, True, ImageList1.ListImages(2).Picture
    XPSideMenu1.AddButton 0, "Button 1", 2400, xpHyperlink, False, True, ImageList2.ListImages(1).Picture, "Click For Messagebox"
    XPSideMenu1.AddButton 0, "Open Form1", 2400, xpHyperlink, False, True, ImageList2.ListImages(2).Picture, "Open Form1"
    XPSideMenu1.AddFrame "XP Side Menu 2", vbSub, vbOpen, 1500, True, ImageList1.ListImages(3).Picture
    XPSideMenu1.AddButton 1, "Exit Application", 2400, xpHyperlink, False, True, ImageList2.ListImages(3).Picture, "Click To Exit App"
    XPSideMenu1.AddFrame "XP Side Menu 3", vbSub, vbOpen, 1500, True, ImageList1.ListImages(1).Picture
    XPSideMenu1.AddButton 2, "Custom Button", 2400, xpCustom, False, True, ImageList2.ListImages(4).Picture, "Button Does Nothing"

End Sub

Private Sub XPSideMenu1_Action(Frame As Integer, Button As Integer)

    Select Case Frame
    Case 0
        If Button = 0 Then
            MsgBox "Button 1"
        Else
            Load Form1
        End If
    Case 1
        Unload Me
    End Select
End Sub

Private Sub XPSideMenu1_ToolTipOver(Tip As String)

    StatusBar1.Panels(1).Text = Tip
End Sub
