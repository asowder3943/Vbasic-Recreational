VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "regpoly solver"
   ClientHeight    =   6330
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12750
   LinkTopic       =   "Form1"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox numt 
      Height          =   285
      Left            =   3480
      TabIndex        =   19
      Text            =   "0"
      Top             =   3840
      Width           =   3015
   End
   Begin VB.CommandButton Command1 
      Caption         =   "calculate"
      Height          =   615
      Left            =   7680
      TabIndex        =   18
      Top             =   4440
      Width           =   1935
   End
   Begin VB.TextBox areat 
      Height          =   285
      Left            =   3480
      TabIndex        =   16
      Text            =   "0"
      Top             =   4200
      Width           =   3015
   End
   Begin VB.TextBox apothemt 
      Height          =   285
      Left            =   3480
      TabIndex        =   15
      Text            =   "0"
      Top             =   5280
      Width           =   3015
   End
   Begin VB.TextBox radiust 
      Height          =   285
      Left            =   3480
      TabIndex        =   14
      Text            =   "0"
      Top             =   5640
      Width           =   3015
   End
   Begin VB.TextBox perimetert 
      Height          =   285
      Left            =   3480
      TabIndex        =   13
      Text            =   "0"
      Top             =   4560
      Width           =   3015
   End
   Begin VB.TextBox sidet 
      Height          =   285
      Left            =   3480
      TabIndex        =   12
      Text            =   "0"
      Top             =   4920
      Width           =   3015
   End
   Begin VB.Label Label13 
      Alignment       =   2  'Center
      Caption         =   $"regpoly.frx":0000
      BeginProperty Font 
         Name            =   "Tiger"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2355
      Left            =   975
      TabIndex        =   17
      Top             =   1320
      Width           =   14145
   End
   Begin VB.Label radiusl 
      Caption         =   "radius length ="
      Height          =   255
      Left            =   10800
      TabIndex        =   11
      Top             =   5640
      Width           =   3255
   End
   Begin VB.Label perimeterl 
      Caption         =   "perimeter ="
      Height          =   255
      Left            =   10800
      TabIndex        =   10
      Top             =   4560
      Width           =   3255
   End
   Begin VB.Label numl 
      Caption         =   "number of sides ="
      Height          =   255
      Left            =   10800
      TabIndex        =   9
      Top             =   3840
      Width           =   3255
   End
   Begin VB.Label sidel 
      Caption         =   "side length ="
      Height          =   255
      Left            =   10800
      TabIndex        =   8
      Top             =   4920
      Width           =   3255
   End
   Begin VB.Label apotheml 
      Caption         =   "apothem length ="
      Height          =   255
      Left            =   10800
      TabIndex        =   7
      Top             =   5280
      Width           =   3255
   End
   Begin VB.Label areal 
      Caption         =   "area ="
      Height          =   255
      Left            =   10800
      TabIndex        =   6
      Top             =   4200
      Width           =   3255
   End
   Begin VB.Label Label9 
      Caption         =   "radius length ="
      Height          =   255
      Left            =   2040
      TabIndex        =   5
      Top             =   5640
      Width           =   1335
   End
   Begin VB.Label Label5 
      Caption         =   "perimeter ="
      Height          =   255
      Left            =   2040
      TabIndex        =   4
      Top             =   4560
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "number of sides ="
      Height          =   255
      Left            =   2040
      TabIndex        =   3
      Top             =   3840
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "side length ="
      Height          =   255
      Left            =   2040
      TabIndex        =   2
      Top             =   4920
      Width           =   1335
   End
   Begin VB.Label Label2 
      Caption         =   "apothem length ="
      Height          =   255
      Left            =   2040
      TabIndex        =   1
      Top             =   5280
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "area ="
      Height          =   255
      Left            =   2040
      TabIndex        =   0
      Top             =   4200
      Width           =   1335
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim num, area, perimeter, side, apothem, radius, same As Double

Private Sub Command1_Click()
num = numt.Text
area = areat.Text
perimeter = perimetert.Text
side = sidel.Text
apothem = apothemt.Text
radius = radiust.Text
numtest = 2

: First
numtest = numtest + 1

If num * area > 0 Then
apothem = area / (n ^ 2 * Tan(180 / n))
perimeter = 2 * area / apothem
side = 4 * area / (n ^ 2 * Tan(180 / n))
radius = Sqr(apothem ^ 2 + side ^ 2)

num1.Caption = num
area1.Caption = area
perimeterl.Caption = perimeter
sidel.Caption = side
apotheml.Caption = apothem
radiusl.Caption = radius
Exit Sub
End If

If num * perimeter > 0 Then
area = num * (perimeter / num) ^ 2 * Tan(180 / num) / 4
GoTo First
End If

If num * side > 0 Then
area = num * side ^ 2 * Tan(180 / num) / 4
GoTo First
End If

If num * apothem > 0 Then
area = num * apothem ^ 2 * Tan(180 / num)
GoTo First
End If

If num * radius > 0 Then
area = num * radius ^ 2 * Sin(180 / num) * Cos(180 / num)
GoTo First
End If

If area * perimeter > 0 And area = numtest * (2 * area / perimeter) ^ 2 * Tan(180 / numtest) Then
num = numtest
GoTo First
End If

If area * side > 0 And area = numtest * side ^ 2 * Tan(180 / numtest) Then
num = numtest
GoTo First
End If

If area * apothem > 0 Then
perimeter = 2 * area / apothem
GoTo First

If area * radius > 0 And area = numtest * radius ^ 2 * Sin(180 / numtest) Then
num = numtest
GoTo First

If perimeter * side > 0 Then
num = perimeter / side
GoTo First

If perimeter * apothem > 0 Then
area = apothem * perimeter / 2
GoTo First

If perimeter * radius > 0 And numtest * radius * 2 / Sin(180 / numtest) Then
num = numtest
GoTo First













End Sub
