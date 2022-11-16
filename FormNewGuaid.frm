VERSION 5.00
Begin VB.Form FormNewGuaid 
   Caption         =   "新建向导"
   ClientHeight    =   5100
   ClientLeft      =   1425
   ClientTop       =   6300
   ClientWidth     =   7215
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   5100
   ScaleWidth      =   7215
   Begin VB.CommandButton Command3 
      Caption         =   "取消"
      Height          =   375
      Left            =   3480
      TabIndex        =   9
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Caption         =   "下一步"
      Height          =   375
      Left            =   4680
      TabIndex        =   8
      Top             =   4440
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   375
      Left            =   5880
      TabIndex        =   7
      Top             =   4440
      Width           =   1095
   End
   Begin VB.ComboBox Combo1 
      Height          =   300
      ItemData        =   "FormNewGuaid.frx":0000
      Left            =   4920
      List            =   "FormNewGuaid.frx":000D
      TabIndex        =   5
      Text            =   " "
      Top             =   3720
      Width           =   2055
   End
   Begin VB.Frame Frame1 
      Caption         =   "框架"
      Height          =   1335
      Left            =   2280
      TabIndex        =   2
      Top             =   2040
      Width           =   4695
      Begin VB.OptionButton Option2 
         Caption         =   "滚动点名窗口"
         Height          =   375
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   3135
      End
      Begin VB.OptionButton Option1 
         Caption         =   "随机数点名窗口"
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   4335
      End
   End
   Begin VB.Label Label3 
      Caption         =   "选项："
      Height          =   255
      Left            =   4320
      TabIndex        =   6
      Top             =   3780
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   120
      X2              =   7080
      Y1              =   4320
      Y2              =   4320
   End
   Begin VB.Label Label2 
      Caption         =   "你想创造什么样的新窗口？"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   1080
      Width           =   4695
   End
   Begin VB.Label Label1 
      Caption         =   "欢迎使用新建向导！"
      Height          =   375
      Left            =   2280
      TabIndex        =   0
      Top             =   120
      Width           =   4695
   End
   Begin VB.Image Image1 
      Height          =   3975
      Left            =   120
      Picture         =   "FormNewGuaid.frx":0023
      Top             =   120
      Width           =   1950
   End
End
Attribute VB_Name = "FormNewGuaid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public A As Boolean
Public B As Boolean

Private Sub NEWFunction()
If Option1.Value = True Then
A = True
ElseIf Option2.Value = True Then
A = False
End If
If A = False Then
FormOrder.Show
ElseIf A = True Then
frmTip.Show
End If
Me.Hide
End Sub


Private Sub Command1_Click()
Call NEWFunction
End Sub

Private Sub Command2_Click()
Call NEWFunction
End Sub

Private Sub Command3_Click()
Unload Me
End Sub
