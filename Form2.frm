VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "ÐÂ°æËµÃ÷"
   ClientHeight    =   3015
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   8700
   LinkTopic       =   "Form2"
   MDIChild        =   -1  'True
   ScaleHeight     =   3015
   ScaleWidth      =   8700
   Begin VB.Label Label1 
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   8415
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Const TIP_FILE = "WhatsNew.txt"
Dim NEWVERtxt(5000) As Variant
Dim i As Variant
Dim Lc As Variant
Dim d As Variant


Private Sub Form_Load()
 Open (App.Path & "\" & TIP_FILE) For Input As #2
Do While Not EOF(2)
Input #2, B
Label1.Caption = Label1.Caption & B
Loop
Close #2
End Sub

