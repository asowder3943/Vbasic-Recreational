VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   16200
   LinkTopic       =   "Form1"
   ScaleHeight     =   8325
   ScaleWidth      =   16200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Calculate"
      Height          =   615
      Left            =   5640
      TabIndex        =   3
      Top             =   5520
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   3360
      TabIndex        =   2
      Text            =   "0"
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text2 
      Height          =   495
      Left            =   5880
      TabIndex        =   1
      Text            =   "0"
      Top             =   4440
      Width           =   1575
   End
   Begin VB.TextBox Text3 
      Height          =   495
      Left            =   8520
      TabIndex        =   0
      Text            =   "0"
      Top             =   4440
      Width           =   1575
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Adam Sowder's Pythyagorean theorem calculator"
      BeginProperty Font 
         Name            =   "Ravie"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   0
      TabIndex        =   12
      Top             =   0
      Width           =   13215
   End
   Begin VB.Line Line1 
      X1              =   6000
      X2              =   9000
      Y1              =   3480
      Y2              =   3480
   End
   Begin VB.Line Line2 
      X1              =   6000
      X2              =   6000
      Y1              =   1080
      Y2              =   3480
   End
   Begin VB.Line Line3 
      X1              =   9000
      X2              =   6000
      Y1              =   3480
      Y2              =   1080
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "a"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   5760
      TabIndex        =   11
      Top             =   2160
      Width           =   255
   End
   Begin VB.Label Label3 
      BackStyle       =   0  'Transparent
      Caption         =   "c"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7560
      TabIndex        =   10
      Top             =   2040
      Width           =   375
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "b"
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7200
      TabIndex        =   9
      Top             =   3480
      Width           =   255
   End
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H8000000B&
      Caption         =   "missing = "
      Height          =   615
      Left            =   5040
      TabIndex        =   8
      Top             =   6960
      Width           =   3375
   End
   Begin VB.Shape Shape1 
      Height          =   255
      Left            =   6000
      Top             =   3240
      Width           =   255
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "A="
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   2880
      TabIndex        =   7
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "C="
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8040
      TabIndex        =   6
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "B="
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   5400
      TabIndex        =   5
      Top             =   4560
      Width           =   375
   End
   Begin VB.Label Label10 
      Caption         =   $"pythagorean theorem.frx":0000
      BeginProperty Font 
         Name            =   "Gill Sans Ultra Bold"
         Size            =   9
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2895
      Left            =   2160
      TabIndex        =   4
      Top             =   1080
      Width           =   2775
   End
   Begin VB.Label Label6 
      Alignment       =   1  'Right Justify
      Caption         =   "Right triangle                      "
      Height          =   2895
      Left            =   5040
      TabIndex        =   13
      Top             =   1080
      Width           =   4815
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim a, b, c As Double

Private Sub Command1_Click()

On Error GoTo help
'defines variables needed for calculations
a = Text1.Text
b = Text2.Text
c = Text3.Text
'makes sure no neg numbers are entered, aswell as stops more than 2 from being 0
If (a < 0 Or b < 0 Or c < 0) Then
MsgBox ("non real triangle")
Else
If (a + b + c = 0 Or (a + b + c = a) Or (a + b + c = b) Or (a + b + c = c)) Then
MsgBox ("not enough information")
Else
'calculates and displays missing value
If (a = 0) Then
Label5.Caption = "missing = " & Sqr(c ^ 2 - b ^ 2)
Else
If (b = 0) Then
Label5.Caption = "missing = " & Sqr(c ^ 2 - a ^ 2)
Else
If (c = 0) Then
Label5.Caption = "missing = " & Sqr(a ^ 2 + b ^ 2)
Else
'displays whether compleated triangle is right or not
If (a ^ 2 + b ^ 2 = c ^ 2) Then
MsgBox ("correct right triangle")
Else
MsgBox ("non-right triangle")
End If
End If
End If
End If
End If
End If
Exit Sub
'incase of input error
help:
MsgBox ("input error read instructions and try again")
End Sub
