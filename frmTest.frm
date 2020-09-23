VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmTest 
   Caption         =   "Sample"
   ClientHeight    =   4980
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4485
   LinkTopic       =   "Form1"
   ScaleHeight     =   4980
   ScaleWidth      =   4485
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   3060
      Top             =   270
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0000
            Key             =   "drive"
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0452
            Key             =   "computer"
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":08A4
            Key             =   "closed"
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":0CF6
            Key             =   "floppy"
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":1148
            Key             =   "network"
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":159A
            Key             =   "cd"
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmTest.frx":19EC
            Key             =   "open"
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView tvFileTree 
      Height          =   4425
      Left            =   45
      TabIndex        =   0
      Top             =   0
      Width           =   4425
      _ExtentX        =   7805
      _ExtentY        =   7805
      _Version        =   393217
      LabelEdit       =   1
      LineStyle       =   1
      Style           =   7
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   465
      Left            =   45
      TabIndex        =   1
      Top             =   4500
      Width           =   4380
   End
End
Attribute VB_Name = "frmTest"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
    Me.tvFileTree.Move 0, 0, Me.ScaleWidth, Me.ScaleHeight - Me.Label1.Height - 50
    Me.Label1.Width = Me.ScaleWidth
    Me.Label1.Left = 0
    LoadExplorerTreeView Me.tvFileTree
End Sub


Private Sub Form_Resize()
    Me.tvFileTree.Height = Me.ScaleHeight - Me.Label1.Height - 50
    Me.tvFileTree.Width = Me.ScaleWidth
    Me.Label1.Width = Me.ScaleWidth
End Sub


Private Sub tvFileTree_Expand(ByVal Node As MSComctlLib.Node)
    tvHandleClick Node, Me.tvFileTree
End Sub

Private Sub tvFileTree_NodeClick(ByVal Node As MSComctlLib.Node)

    tvHandleClick Node, Me.tvFileTree
    ' show how you can get the path
    Me.Label1.Caption = GetTruePath(Node)
    
End Sub


