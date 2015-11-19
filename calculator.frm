VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "calculator"
   ClientHeight    =   8415
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   15360
   LinkTopic       =   "Form1"
   Picture         =   "calculator.frx":0000
   ScaleHeight     =   8415
   ScaleWidth      =   15360
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command7 
      Caption         =   "int division"
      Height          =   375
      Left            =   8400
      TabIndex        =   10
      Top             =   4560
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "mod"
      Height          =   375
      Left            =   6720
      TabIndex        =   9
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command5 
      Caption         =   "raise"
      Height          =   375
      Left            =   5040
      TabIndex        =   8
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton Command4 
      Caption         =   "divide"
      Height          =   375
      Left            =   9000
      TabIndex        =   7
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "multiply"
      Height          =   375
      Left            =   7560
      TabIndex        =   6
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command2 
      Caption         =   "subtract"
      Height          =   375
      Left            =   6120
      TabIndex        =   5
      Top             =   3960
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "add"
      Height          =   375
      Left            =   4680
      TabIndex        =   4
      Top             =   3960
      Width           =   975
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   8520
      TabIndex        =   3
      Top             =   2760
      Width           =   1455
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   4680
      TabIndex        =   1
      Top             =   2760
      Width           =   1575
   End
   Begin VB.Label Label2 
      Height          =   495
      Left            =   4800
      TabIndex        =   2
      Top             =   5400
      Width           =   5175
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Adam Sowders Calculator"
      BeginProperty Font 
         Name            =   "Snap ITC"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   735
      Left            =   4440
      TabIndex        =   0
      Top             =   1080
      Width           =   6375
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim x As Double
Dim y As Double

Private Sub Text1_Change()
On Error GoTo help
Label2.Caption = ""
x = Text1.Text
Exit Sub
help:
MsgBox ("input error try again")
End Sub

Private Sub Text2_Change()
On Error GoTo help:
Label2.Caption = ""
y = Text2.Text
Exit Sub
help:
MsgBox ("input error try again")
End Sub

Private Sub Command1_Click()
Label2.Caption = "answer = " & (x + y)
End Sub

Private Sub Command2_Click()
Label2.Caption = "answer = " & (x - y)
End Sub

Private Sub Command3_Click()
Label2.Caption = "answer = " & (x * y)
End Sub

Private Sub Command4_Click()
Label2.Caption = "answer = " & (x / y)
End Sub

Private Sub Command5_Click()
Label2.Caption = "answer = " & (x ^ y)
End Sub

Private Sub Command6_Click()
Label2.Caption = "answer = " & (x Mod y)
End Sub

Private Sub Command7_Click()
Label2.Caption = "answer = " & (x \ y)
End Sub
