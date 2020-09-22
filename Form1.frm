VERSION 5.00
Object = "{1103DCBC-6CBA-11D6-8D03-00B0D091BA0C}#3.0#0"; "xpside.ocx"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MsComCtl.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Form1"
   ClientHeight    =   2040
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   2040
   ScaleWidth      =   5820
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   1
      Top             =   1665
      Width           =   5820
      _ExtentX        =   10266
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
      Height          =   1665
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   3285
      _ExtentX        =   5794
      _ExtentY        =   2937
   End
   Begin MSComctlLib.ImageList ImageList2 
      Left            =   4620
      Top             =   3660
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   4845
      Top             =   2355
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
            Picture         =   "Form1.frx":059A
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":22A4
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":3FAE
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()

    XPSideMenu1.AddFrame "Another One", vbMain, vbOpen, 1500, True, ImageList1.ListImages(2).Picture
    XPSideMenu1.AddButton 0, "Button 1", 2400, xpHyperlink, False, True, MDIForm1.ImageList2.ListImages(1).Picture, "Whatever"
    XPSideMenu1.AddButton 0, "Button 2", 2400, xpHyperlink, False, True, MDIForm1.ImageList2.ListImages(2).Picture, "Nothing"
    XPSideMenu1.AddButton 0, "Button 3", 2400, xpHyperlink, False, True, MDIForm1.ImageList2.ListImages(3).Picture, "Nothing"
    XPSideMenu1.AddButton 0, "Button 4", 2400, xpHyperlink, False, True, MDIForm1.ImageList2.ListImages(4).Picture, "Nothing"
    XPSideMenu1.AddButton 0, "Button 5", 2400, xpHyperlink, False, True, MDIForm1.ImageList2.ListImages(5).Picture, "Nothing"
    XPSideMenu1.AddButton 0, "Button 6", 2400, xpHyperlink, False, True, MDIForm1.ImageList2.ListImages(6).Picture, "Nothing"
    XPSideMenu1.AddFrame "Last One", vbSub, vbOpen, 1500, True, ImageList1.ListImages(3).Picture
    XPSideMenu1.AddButton 1, "Button 1", 2400, xpHyperlink, False, True, MDIForm1.ImageList2.ListImages(7).Picture, "Nothing"
End Sub

Private Sub XPSideMenu1_ToolTipOver(Tip As String)

    StatusBar1.Panels(1).Text = Tip
End Sub
