VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Star Draw"
   ClientHeight    =   5760
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   5940
   LinkTopic       =   "Form1"
   ScaleHeight     =   5760
   ScaleWidth      =   5940
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Height          =   375
      Left            =   2603
      TabIndex        =   2
      Text            =   "3"
      Top             =   4560
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Draw Star"
      Height          =   375
      Left            =   2243
      TabIndex        =   1
      Top             =   5160
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      Height          =   3855
      Left            =   203
      ScaleHeight     =   3795
      ScaleWidth      =   5475
      TabIndex        =   0
      Top             =   240
      Width           =   5535
      Begin VB.CommandButton Command3 
         Caption         =   ">"
         Height          =   255
         Left            =   1680
         TabIndex        =   4
         Top             =   120
         Width           =   375
      End
      Begin VB.CommandButton Command2 
         Caption         =   "<"
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   120
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "var _ of _"
         Height          =   255
         Left            =   720
         TabIndex        =   5
         Top             =   120
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim starnum, deginc, varnum, varmax As Double
Const pi = 3.141592

Private Sub initVars()
starnum = 3
Call deginset
varnum = 1
Call varset
End Sub

Private Sub Form_Load()
Call initVars
Call drawstar(starnum, varnum)
End Sub

Private Sub deginset()
deginc = 60 / starnum * pi / 180
End Sub

Private Sub drawstar()
For x = 1 To starnum

Next
End Sub

Private Sub varset()
varmax = starnum \ 2
End Sub

Function modset(ByVal num As Double, ByVal max As Double) As Double
If num < max Then: modset = num
If num > max Then modset = num - modset
End Function
