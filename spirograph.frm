VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Adam's Digital Spirograph"
   ClientHeight    =   8535
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   7515
   LinkTopic       =   "Form1"
   ScaleHeight     =   8535
   ScaleWidth      =   7515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton speedMore 
      Caption         =   ">"
      Height          =   255
      Left            =   4200
      TabIndex        =   14
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton speedLess 
      Caption         =   "<"
      Enabled         =   0   'False
      Height          =   255
      Left            =   2880
      TabIndex        =   13
      Top             =   360
      Width           =   375
   End
   Begin VB.CommandButton resetAnimation 
      Caption         =   "Reset Animation"
      Height          =   495
      Left            =   5610
      TabIndex        =   12
      Top             =   6480
      Width           =   1575
   End
   Begin VB.CommandButton clearDraw 
      Caption         =   "Clear Drawing"
      Height          =   495
      Left            =   5610
      TabIndex        =   11
      Top             =   5880
      Width           =   1575
   End
   Begin VB.CommandButton Done 
      Caption         =   "Done"
      Height          =   495
      Left            =   5970
      TabIndex        =   10
      Top             =   7200
      Width           =   855
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Epitrochoid"
      Height          =   495
      Left            =   2820
      TabIndex        =   9
      Top             =   7440
      Width           =   1695
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Hypotrochoid"
      Enabled         =   0   'False
      Height          =   495
      Left            =   780
      TabIndex        =   8
      Top             =   7440
      Width           =   1695
   End
   Begin VB.TextBox gllText 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Height          =   285
      Left            =   3120
      TabIndex        =   6
      Text            =   "5"
      Top             =   6840
      Width           =   615
   End
   Begin VB.TextBox rcrText 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Height          =   285
      Left            =   4335
      TabIndex        =   4
      Text            =   "3"
      Top             =   6360
      Width           =   615
   End
   Begin VB.TextBox bcrText 
      Alignment       =   2  'Center
      BackColor       =   &H8000000E&
      Height          =   285
      Left            =   1935
      TabIndex        =   2
      Text            =   "5"
      Top             =   6360
      Width           =   615
   End
   Begin VB.CommandButton showAnimation 
      Caption         =   "Show animation"
      Height          =   495
      Left            =   1770
      TabIndex        =   1
      Top             =   5520
      Width           =   1815
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   5130
      Top             =   7920
   End
   Begin VB.PictureBox Graph 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   3975
      Left            =   390
      ScaleHeight     =   3915
      ScaleWidth      =   6675
      TabIndex        =   0
      Top             =   840
      Width           =   6735
      Begin VB.Label Label7 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         Height          =   375
         Left            =   5160
         TabIndex        =   18
         Top             =   3360
         Width           =   1455
      End
      Begin VB.Label Label6 
         BackColor       =   &H8000000E&
         BackStyle       =   0  'Transparent
         Caption         =   "Label6"
         Height          =   375
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   1095
      End
      Begin VB.Line Line1 
         BorderColor     =   &H0000C000&
         X1              =   3840
         X2              =   6120
         Y1              =   1560
         Y2              =   360
      End
      Begin VB.Shape Shape2 
         BorderColor     =   &H000000FF&
         Height          =   375
         Left            =   3000
         Shape           =   3  'Circle
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Shape Shape1 
         BorderColor     =   &H00FF0000&
         Height          =   855
         Left            =   960
         Shape           =   3  'Circle
         Top             =   1800
         Width           =   1455
      End
   End
   Begin VB.Label errorLabel 
      Caption         =   "Displaying figure representative of input"
      Height          =   375
      Left            =   3810
      TabIndex        =   16
      Top             =   5280
      Width           =   3015
   End
   Begin VB.Label speedLabel 
      Alignment       =   2  'Center
      Caption         =   "Speed"
      Height          =   255
      Left            =   3450
      TabIndex        =   15
      Top             =   360
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Green line length"
      Height          =   255
      Left            =   1560
      TabIndex        =   7
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Label Label2 
      Caption         =   "Red Circle Radius"
      Height          =   255
      Left            =   2775
      TabIndex        =   5
      Top             =   6360
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Blue Circle Radius"
      Height          =   255
      Left            =   330
      TabIndex        =   3
      Top             =   6360
      Width           =   1455
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'<variables>
    Dim center(1 To 2), dest(1 To 2), temp(1 To 2) As Double
    Dim blRad, rdRad, tRad, grLen As Double
    Dim theta, degInc, ltheta, uscale As Double
    Dim hypotrochoid, epitrochoid As Boolean
    Dim thick, z, defNum As Integer
    Const pi = 3.141592654
    
    Private Sub initVars() '<-sets initial values of variables
            center(1) = Graph.Width / 2: center(2) = Graph.Height / 2
            If Graph.Height < Graph.Width Then
                blRad = Graph.Height / 4
            Else
                blRad = Graph.Width / 4
            End If
            defNum = 1
            
            rdRad = blRad / 5 * 3
            tRad = blRad - rdRad
            grLen = blRad
            
            theta = 90
            ltheta = 0
            degInc = -1
            uscale = rdRad / 3
            
            Call Command2_Click
            thick = 2
            
            Label6.Left = 0
            Label6.Top = 0
            Label7.Left = Graph.Width - Label7.Width: Label7.Top = Graph.Height - Label7.Height
            Label6.Caption = "xT = " & vbCrLf & "yT = "
    End Sub
'</variable>

'<object dependant subs>
    Private Sub Form_Load()
        Call initVars
        
        Call bigcirc(blRad)
        Call lilcirc(theta)
        
        Call dline(theta, ltheta)
        Call plotGraph
    End Sub

    Private Sub Graph_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
        Label7.Caption = "Cursor: (" & Round(cConvert(CDbl(X), "X"), 1) & "," & Round(cConvert(CDbl(Y), "Y"), 1) & ")"
    End Sub

    Private Sub Timer1_Timer()
        z = 0
        Do
        theta = theta - degInc: Call lilcirc(theta)
        ltheta = ltheta + degInc: Call dline(theta, ltheta)
        Call fcurve(ltheta)
        z = z + 1
        Loop While z < defNum
    End Sub
    
    '<command buttons>
        Private Sub showAnimation_Click() '<-starts and stops animation
            Timer1.Enabled = Not Timer1.Enabled
            If Timer1.Enabled Then
                showAnimation.Caption = "stop animation"
            Else
                showAnimation.Caption = "show animation"
            End If
        End Sub
        
        Private Sub Command2_Click() '<-sets the function to hypotrochoid
            hypotrochoid = True
            epitrochoid = False
            Command2.Enabled = False
            Command3.Enabled = True
            Call reset
        End Sub
        
        Private Sub Command3_Click() '<-sets the function to epitrochoid
            hypotrochoid = False
            epitrochoid = True
            Command2.Enabled = True
            Command3.Enabled = False
            Call reset
        End Sub
        
        Private Sub Done_Click()
            Unload Me
        End Sub
        
        Private Sub clearLabel_Click()
            Graph.Cls
            Call plotGraph
        End Sub
        
        Private Sub resetAnimation_Click()
            Call reset
        End Sub
        
        '<speed controls>
            Private Sub speedMore_Click()
                speedLess.Enabled = True
                defNum = defNum + 1
                If defNum > 9 Then: speedMore.Enabled = False
            End Sub
        
            Private Sub speedLess_Click()
                speedMore.Enabled = True
                defNum = defNum - 1
                If defNum < 2 Then: speedLess.Enabled = False
            End Sub
        '</speed controls>
    '</command buttons>
    
    '<textbox>
        Private Sub bcrText_Change() '<-changes size of blue circle
            Call valid
            If IsNumeric(bcrText.Text()) Then
                temp(1) = bcrText.Text:
                If temp(1) > 0 Then
                    blRad = bcrText.Text
                    blRad = blRad * uscale
                Else
                    Call invalid
                End If
            Else
                Call invalid
            End If
            Call reset
        End Sub
        
        Private Sub rcrText_Change() '<-changes size of red circle
            Call valid
            If IsNumeric(rcrText.Text()) Then
                temp(1) = rcrText.Text:
                If temp(1) > 0 Then
                    rdRad = rcrText.Text
                    rdRad = rdRad * uscale
                Else
                    Call invalid
                End If
            Else
                Call invalid
            End If
            Call reset
        End Sub
            
        Private Sub gllText_Change() '<-changes size of green line
            Call valid
            If IsNumeric(gllText.Text()) Then
                temp(1) = gllText.Text:
                If temp(1) > 0 Then
                   grLen = gllText.Text
                   grLen = grLen * uscale
                Else
                    Call invalid
                End If
            Else
                Call invalid
            End If
            Call reset
        End Sub
    '</textbox>
'</object dependant subs>

'<dest finder subs>
    Function blCenter(ByVal deg As Double, coordType As String) As Double
        radeg = deg * pi / 180
        If StrComp(coordType, "X") = 0 Then: blCenter = (Sin(radeg) * tRad + center(1))
        If StrComp(coordType, "Y") = 0 Then: blCenter = (Cos(radeg) * tRad + center(2))
    End Function
    
    Function funCoord(ByVal deg As Double, coordType As String) As Double
        radeg = deg * pi / 180
        If hypotrochoid And StrComp(coordType, "X") = 0 Then: funCoord = (tRad * Cos(radeg) + grLen * Cos(tRad / rdRad * radeg) + center(1))
        If epitrochoid And StrComp(coordType, "X") = 0 Then: funCoord = (tRad * Cos(radeg) - grLen * Cos(tRad / rdRad * radeg) + center(1))
        If StrComp(coordType, "Y") = 0 Then: funCoord = Int(tRad * Sin(radeg) - grLen * Sin(tRad / rdRad * radeg) + center(2))
    End Function
'</dest finder subs>

'<draws shapes>
    Private Sub bigcirc(ByVal r As Double) '<-draw blue circle
        Shape1.Width = r * 2
        Shape1.Height = Shape1.Width
        Shape1.Left = center(1) - r
        Shape1.Top = center(2) - r
    End Sub
    
    Private Sub lilcirc(ByVal deg As Double) '<draw red circle
        Shape2.Width = rdRad * 2
        Shape2.Height = Shape2.Width
        Shape2.Left = blCenter(deg, "X") - rdRad
        Shape2.Top = blCenter(deg, "Y") - rdRad
    End Sub
    
    Private Sub dline(ByVal deg1 As Double, ByVal deg2 As Double) '<-draw green line
        Line1.X1 = blCenter(deg1, "X")
        Line1.Y1 = blCenter(deg1, "Y")
        Line1.X2 = funCoord(deg2, "X")
        Line1.Y2 = funCoord(deg2, "Y")
        
    End Sub
'</draws shapes>

'<Plot subs>
    Private Sub plotGraph()
        xmod = uscale - center(1) Mod uscale: ymod = uscale - center(2) Mod uscale
        For X = 0 - xmod To Graph.Width - xmod Step uscale '<-draws vertical grid lines
            Graph.Line (X, 0)-(X, Graph.Height), &H8000000F
        Next
        For Y = 0 - ymod To Graph.Height - ymod Step uscale '<-draws horizontal grid lines
            Graph.Line (0, Y)-(Graph.Width, Y), &H8000000F
        Next
        Graph.Line (Graph.Width / 2, 0)-(Graph.Width / 2, Graph.Height)  '<- y axis
        Graph.Line (0, Graph.Height / 2)-(Graph.Width, Graph.Height / 2) '<- x axis
    End Sub
        
    Private Sub fcurve(ByVal deg As Double)
        Graph.DrawWidth = thick
        Graph.Line (funCoord(deg + degInc, "X"), funCoord(deg + degInc, "Y"))-(funCoord(deg, "X"), funCoord(deg, "Y"))
        Graph.DrawWidth = 1
        Label6.Caption = "xT = " & Round(cConvert(funCoord(deg, "X"), "X"), 1) & vbCrLf & "yT = " & Round(cConvert(funCoord(deg, "Y"), "Y"), 1)
    End Sub
'</Plot subs>

'<helper functions>
    Function cConvert(ByVal Cnum As Double, Ctype As String) As Double
        If StrComp(Ctype, "X") = 0 Then: cConvert = Cnum / uscale - Graph.Width / (2 * uscale)
        If StrComp(Ctype, "Y") = 0 Then: cConvert = (Cnum / uscale - Graph.Height / (2 * uscale)) * -1
    End Function
'</helper functions>

'<other>
    Private Sub reset()
        If Timer1.Enabled Then: Call showAnimation_Click
        theta = 90: ltheta = 0
        If hypotrochoid Then: tRad = blRad - rdRad
        If epitrochoid Then: tRad = blRad + rdRad
        
        Call bigcirc(blRad)
        Call lilcirc(theta)
        Call dline(theta, ltheta)
        
        Graph.Cls
        Call plotGraph
        Label6.Caption = "xT = " & vbCrLf & "yT = "
    End Sub
    
    '<error handling>
        Private Sub invalid()
            errorLabel.Caption = "*invalid entry, Graphic does not represent input"
        End Sub
        
        Private Sub valid()
            errorLabel.Caption = "displaying figure representative of input"
        End Sub
    '</error handling>
'</other>
