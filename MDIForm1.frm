VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.MDIForm MDIForm1 
   BackColor       =   &H8000000C&
   Caption         =   "MDIForm1"
   ClientHeight    =   9225
   ClientLeft      =   120
   ClientTop       =   3405
   ClientWidth     =   11550
   LinkTopic       =   "MDIForm1"
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   10800
      Top             =   720
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      UseMaskColor    =   0   'False
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   2
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "MDIForm1.frx":0C52
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   660
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   11550
      _ExtentX        =   20373
      _ExtentY        =   1164
      ButtonWidth     =   1032
      ButtonHeight    =   1005
      Appearance      =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   2
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "NEWform"
            Object.ToolTipText     =   "创建一个新的窗口（通过新建向导）"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Enabled         =   0   'False
            Key             =   "OPENfile"
            Object.ToolTipText     =   "打开一个文件"
            ImageIndex      =   2
         EndProperty
      EndProperty
   End
   Begin VB.Menu File 
      Caption         =   "文件(&F)"
      Begin VB.Menu NEW 
         Caption         =   "新建"
         Shortcut        =   ^N
      End
      Begin VB.Menu save 
         Caption         =   "保存"
         Enabled         =   0   'False
         Shortcut        =   ^S
      End
      Begin VB.Menu Open 
         Caption         =   "打开"
         Enabled         =   0   'False
         Shortcut        =   ^O
      End
      Begin VB.Menu Nun 
         Caption         =   "--------------"
      End
      Begin VB.Menu exit 
         Caption         =   "退出(&X)"
      End
   End
   Begin VB.Menu Option 
      Caption         =   "选项(&O)"
      Enabled         =   0   'False
      Begin VB.Menu common 
         Caption         =   "通用"
      End
      Begin VB.Menu Advance 
         Caption         =   "高级"
      End
      Begin VB.Menu More 
         Caption         =   "更多"
      End
   End
   Begin VB.Menu Help 
      Caption         =   "帮助(&H)"
      Begin VB.Menu About 
         Caption         =   "关于"
      End
      Begin VB.Menu NEWVER 
         Caption         =   "新版说明"
      End
      Begin VB.Menu feedback 
         Caption         =   "反馈"
         Begin VB.Menu reportBugs 
            Caption         =   "Bug反馈"
         End
         Begin VB.Menu newidea 
            Caption         =   "新点子"
         End
      End
      Begin VB.Menu Update 
         Caption         =   "更新"
         Begin VB.Menu offine 
            Caption         =   "离线更新"
            Enabled         =   0   'False
         End
         Begin VB.Menu online 
            Caption         =   "在线更新"
         End
      End
   End
End
Attribute VB_Name = "MDIForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub About_Click()
frmAbout.Show
End Sub

Private Sub exit_Click()
Unload Me
End Sub

Private Sub MDIForm_Load()
FormNewGuaid.Show
End Sub

Private Sub New_Click()
 FormNewGuaid.Show
End Sub

Private Sub newidea_Click()
frmBrowser2.Show
End Sub

Private Sub NEWVER_Click()
Form2.Show
End Sub

Private Sub online_Click()
frmBrowser3.Show
End Sub

Private Sub reportBugs_Click()
frmBrowser1.Show
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error Resume Next

Select Case Button.Key
    Case "NEWform"
        FormNewGuaid.Show
End Select
End Sub
