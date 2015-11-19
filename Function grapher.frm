VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   9165
   LinkTopic       =   "Form1"
   ScaleHeight     =   6135
   ScaleWidth      =   9165
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   615
      Left            =   7920
      TabIndex        =   5
      Top             =   4920
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   495
      Left            =   2640
      TabIndex        =   4
      Text            =   "Text1"
      Top             =   4920
      Width           =   4695
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   3135
      Left            =   1635
      ScaleHeight     =   3105
      ScaleWidth      =   5865
      TabIndex        =   0
      Top             =   240
      Width           =   5895
   End
   Begin VB.Label Label3 
      Caption         =   "Label3"
      Height          =   735
      Left            =   6600
      TabIndex        =   3
      Top             =   3720
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Label2"
      Height          =   615
      Left            =   4080
      TabIndex        =   2
      Top             =   3840
      Width           =   1575
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   495
      Left            =   1440
      TabIndex        =   1
      Top             =   3840
      Width           =   1695
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Dim Fun(0 To 38) As String
Dim center(1 To 2), xscale, yscale, xInc As Double

Private Sub initVars()
center(1) = Picture1.Width / 2: center(2) = Picture1.Height / 2
xscale = Picture1.Width / 20
yscale = Picture1.Height / 20
xInc = 0.1
Expi
End Sub

Private Sub Command1_Click()
Text1.Text = heal(Text1.Text)
Text1.Text = heal_2(Text1.Text)
End Sub

Private Sub Picture1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
        Label1.Caption = "Cursor: (" & Round(cConvert(CDbl(x), "X"), 1) & "," & Round(cConvert(CDbl(y), "Y"), 1) & ")"
        Label2.Caption = "Cursor: (" & Round(gConvert(cConvert(CDbl(x), "X"), "X"), 1) & "," & Round(gConvert(cConvert(CDbl(y), "Y"), "Y"), 1) & ")"
        Picture1.Cls
        Form_Load
        Picture1.Line (x, 0)-(x, Picture1.Height), vbGreen
        Picture1.Line (0, y)-(Picture1.Width, y), vbGreen
        nx = cConvert(x, "X")
        Picture1.Line (x - 20, funCoord(nx) - 20)-(x + 20, funCoord(nx) + 20), vbRed, BF
        Label3.Caption = "f(" & Round(nx, 1) & ") = " & Round(cConvert(funCoord(nx), "Y"), 1)
    End Sub

Private Sub plotGraph()
    xmod = xscale - center(1) Mod xscale: ymod = yscale - center(2) Mod yscale
    For x = 0 - xmod To Picture1.Width - xmod Step xscale '<-draws vertical grid lines
        Picture1.Line (x, 0)-(x, Picture1.Height), &H8000000F
    Next
    For y = 0 - ymod To Picture1.Height - ymod Step yscale '<-draws horizontal grid lines
        Picture1.Line (0, y)-(Picture1.Width, y), &H8000000F
    Next
    Picture1.Line (Picture1.Width / 2, 0)-(Picture1.Width / 2, Picture1.Height)  '<- y axis
    Picture1.Line (0, Picture1.Height / 2)-(Picture1.Width, Picture1.Height / 2) '<- x axis
End Sub

Private Sub Form_Load()
Call initVars
Call plotGraph
For x = -10 To 10 Step xInc
fcurve (x)
Next
End Sub


Private Sub fcurve(ByVal xval As Double)
    On Error Resume Next
    Picture1.Line (gConvert(xval, "X"), funCoord(xval))-(gConvert(xval + xInc, "X"), funCoord(xval + xInc)), vbBlue
    End Sub

Function cConvert(ByVal Cnum As Double, Ctype As String) As Double
    If StrComp(Ctype, "X") = 0 Then: cConvert = Cnum / xscale - Picture1.Width / (2 * xscale)
    If StrComp(Ctype, "Y") = 0 Then: cConvert = (Cnum / yscale - Picture1.Height / (2 * yscale)) * -1
End Function

Function gConvert(ByVal Cnum As Double, Ctype As String) As Double
    If StrComp(Ctype, "X") = 0 Then: gConvert = (Cnum + Picture1.Width / (2 * xscale)) * xscale
    If StrComp(Ctype, "Y") = 0 Then: gConvert = ((Cnum * -1) + Picture1.Height / (2 * yscale)) * yscale
End Function

Function funCoord(ByVal x As Double) As Double
    funCoord = gConvert(4 * Sin(x), "Y")
End Function





















Public Function heal(equation)                'allows the user to be more slack when entering functions
equation = Replace(equation, "^-", "<!?>")
equation = Replace(equation, " ", "")
equation = "<~>" + equation + "<~>"
For i = 0 To 38
    equation = Replace(equation, Fun(i), "*_" + Str(i) + "_")
Next i
equation = Replace(equation, "[", "(")
equation = Replace(equation, "]", ")")
equation = Replace(equation, "deg", "&&&")
If isForLooks = False Then
    equation = Replace(equation, "^", "```")
    equation = Replace(equation, "&&&", "*(pi/180)")
    equation = Replace(equation, "rad", "*(180/pi)")
    equation = Replace(equation, "x", "z")
    equation = Replace(equation, "Z", "z")
    equation = Replace(equation, "t", "z")
    equation = Replace(equation, "z", "*z^")
    equation = Replace(equation, "A", "*A^")
End If
equation = Replace(equation, "pi", "*pi^")
equation = Replace(equation, "G", "*G^")
equation = Replace(equation, "e", "*e^")
equation = Replace(equation, "lif*e", "*life^")
equation = Replace(equation, "(", "*(")
equation = Replace(equation, ")", ")^")
equation = Replace(equation, "***", "*")
equation = Replace(equation, "**", "*")
equation = Replace(equation, "^^", "^")
equation = Replace(equation, "z^(", "z*(")
equation = Replace(equation, "E*", "E")
equation = Replace(equation, "*E", "E")
equation = Replace(equation, "^-", "-")
equation = Replace(equation, "<!?>", "^-")
equation = Replace(equation, "**", "*")
equation = Replace(equation, "^=", "=")
equation = Replace(equation, "=*", "=")
equation = Replace(equation, "^<", "<")
equation = Replace(equation, "<*", "<")
equation = Replace(equation, "^>", ">")
equation = Replace(equation, ">*", ">")
equation = Replace(equation, "^+", "+")
equation = Replace(equation, "^/", "/")
equation = Replace(equation, "^\", "\")
equation = Replace(equation, "^)", ")")
equation = Replace(equation, "(*", "(")
equation = Replace(equation, "+*", "+")
equation = Replace(equation, "-*", "-")
equation = Replace(equation, "/*", "/")
equation = Replace(equation, "\*", "\")
equation = Replace(equation, "^*", "*")
equation = Replace(equation, "```", "^")
equation = Replace(equation, "^*", "^")
equation = Replace(equation, "^^", "^")
equation = Replace(equation, "*mod", "mod")
equation = Replace(equation, "^mod", "mod")
equation = Replace(equation, "mod*", "mod")
equation = Replace(equation, "&&&", "deg")
equation = Replace(equation, "()", "")
For i = 0 To 38
    equation = Replace(equation, "_" + Str(i) + "_*", Fun(i))
Next i
equation = Replace(equation, "_ 24_", Fun(24))
heal = Replace(equation, "<~>", "")
e_1 = Replace(equation, "(", "")
e_2 = Replace(equation, ")", "")
If Len(e_1) <> Len(e_2) Then MsgBox "One of your expressions does not contain an equal number of opened and closed parenthesis.", , "ParaGraph  -  Formula error!": plotsmooth = False
End Function






Public Function heal_2(equation)                    'second healing step
    equation = Replace(equation, "^-", "<err2>")
    equation = Replace(equation, "/-", "<err4>")
    equation = Replace(equation, "E-", "<err3>")
    equation = Replace(equation, "*-", "<err5>")
    equation = Replace(equation, "-", "-1*")
    equation = Replace(equation, "<err2>", "^-")
    equation = Replace(equation, "<err3>", "E-")
    equation = Replace(equation, "<err4>", "/-")
    equation = Replace(equation, "<err5>", "*-")
    equation = Replace(equation, "rnd", "RandomNumber()")
    equation = Replace(equation, "log", "baseten")
    equation = Replace(equation, "E", "*10^")
    equation = Replace(equation, "**", "*")
    heal_2 = Replace(equation, "sqrt", "sqr")
End Function

Private Sub Expi()
Fun(0) = "asinh"
Fun(1) = "acosh"
Fun(2) = "atanh"
Fun(3) = "acsch"
Fun(4) = "asech"
Fun(5) = "acoth"
Fun(6) = "asin"
Fun(7) = "acos"
Fun(8) = "atan"
Fun(9) = "acsc"
Fun(10) = "asec"
Fun(11) = "acot"
Fun(12) = "sinh"
Fun(13) = "cosh"
Fun(14) = "tanh"
Fun(15) = "csch"
Fun(16) = "sech"
Fun(17) = "coth"
Fun(18) = "sin"
Fun(19) = "cos"
Fun(20) = "tan"
Fun(21) = "csc"
Fun(22) = "sec"
Fun(23) = "cot"
Fun(24) = "rnd"
Fun(25) = "abs"
Fun(26) = "sqrt"
Fun(27) = "sqr"
Fun(28) = "log"
Fun(29) = "ln"
Fun(30) = "exp"
Fun(31) = "round"
Fun(32) = "fix"
Fun(33) = "int"
Fun(34) = "frac"
Fun(35) = "sgn"
Fun(36) = "fact"
Fun(37) = "prime"
Fun(38) = "fib"
End Sub
