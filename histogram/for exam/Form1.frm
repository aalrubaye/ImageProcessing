VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8850
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10515
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   10515
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command10 
      Caption         =   "Errosion"
      Height          =   375
      Left            =   6600
      TabIndex        =   35
      Top             =   6960
      Width           =   1095
   End
   Begin VB.PictureBox Picture9 
      Height          =   1575
      Left            =   4440
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   34
      Top             =   6960
      Width           =   2055
   End
   Begin VB.CommandButton Command9 
      Caption         =   "Dilation"
      Height          =   375
      Left            =   6600
      TabIndex        =   33
      Top             =   5280
      Width           =   1095
   End
   Begin VB.PictureBox Picture8 
      Height          =   1575
      Left            =   4440
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   32
      Top             =   5280
      Width           =   2055
   End
   Begin VB.TextBox Text13 
      Height          =   495
      Left            =   8400
      TabIndex        =   31
      Top             =   2280
      Width           =   1935
   End
   Begin VB.TextBox Text12 
      Height          =   285
      Left            =   6600
      TabIndex        =   30
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox Text11 
      Height          =   285
      Left            =   6600
      TabIndex        =   29
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Command8 
      Caption         =   "ACE"
      Height          =   375
      Left            =   6600
      TabIndex        =   28
      Top             =   3600
      Width           =   855
   End
   Begin VB.PictureBox Picture7 
      Height          =   1575
      Left            =   4440
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   27
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text10 
      Height          =   285
      Left            =   2280
      TabIndex        =   26
      Top             =   7440
      Width           =   855
   End
   Begin VB.TextBox Text9 
      Height          =   285
      Left            =   2280
      TabIndex        =   25
      Top             =   6120
      Width           =   855
   End
   Begin VB.TextBox Text8 
      Height          =   285
      Left            =   2280
      TabIndex        =   24
      Top             =   5760
      Width           =   855
   End
   Begin VB.TextBox Text7 
      Height          =   285
      Left            =   2280
      TabIndex        =   23
      Top             =   4440
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   285
      Left            =   2280
      TabIndex        =   22
      Top             =   4080
      Width           =   855
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Slide"
      Height          =   375
      Left            =   2280
      TabIndex        =   21
      Top             =   6960
      Width           =   975
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Shrink"
      Height          =   375
      Left            =   2280
      TabIndex        =   20
      Top             =   5280
      Width           =   975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Streach"
      Height          =   375
      Left            =   2280
      TabIndex        =   19
      Top             =   3600
      Width           =   975
   End
   Begin VB.PictureBox Picture6 
      Height          =   1575
      Left            =   120
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   18
      Top             =   6960
      Width           =   2055
   End
   Begin VB.PictureBox Picture5 
      Height          =   1575
      Left            =   120
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   17
      Top             =   5280
      Width           =   2055
   End
   Begin VB.PictureBox Picture4 
      Height          =   1575
      Left            =   120
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   16
      Top             =   3600
      Width           =   2055
   End
   Begin VB.TextBox Text5 
      Height          =   405
      Left            =   8400
      TabIndex        =   11
      Top             =   1560
      Width           =   2295
   End
   Begin VB.TextBox Text4 
      Height          =   405
      Left            =   8400
      TabIndex        =   10
      Top             =   1080
      Width           =   2295
   End
   Begin VB.TextBox Text3 
      Height          =   405
      Left            =   8400
      TabIndex        =   9
      Top             =   600
      Width           =   2295
   End
   Begin VB.TextBox Text2 
      Height          =   405
      Left            =   8400
      TabIndex        =   8
      Top             =   120
      Width           =   2295
   End
   Begin VB.CommandButton Command4 
      Caption         =   "His. Fetures"
      Height          =   495
      Left            =   5760
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   2280
      TabIndex        =   6
      Top             =   1920
      Width           =   2055
   End
   Begin VB.PictureBox Picture3 
      Height          =   1575
      Left            =   120
      ScaleHeight     =   300
      ScaleMode       =   0  'User
      ScaleWidth      =   300
      TabIndex        =   5
      Top             =   1920
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Show Histogram"
      Height          =   495
      Left            =   5760
      TabIndex        =   4
      Top             =   120
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   120
      Top             =   2400
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Browse"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Gray"
      Height          =   495
      Left            =   4440
      TabIndex        =   2
      Top             =   720
      Width           =   1215
   End
   Begin VB.PictureBox Picture2 
      Height          =   1575
      Left            =   2280
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   1
      Top             =   120
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   120
      ScaleHeight     =   101
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   133
      TabIndex        =   0
      Top             =   120
      Width           =   2055
   End
   Begin VB.Label Label4 
      Caption         =   "entropy"
      Height          =   255
      Left            =   7800
      TabIndex        =   15
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "energy"
      Height          =   255
      Left            =   7800
      TabIndex        =   14
      Top             =   1200
      Width           =   495
   End
   Begin VB.Label Label2 
      Caption         =   "STD"
      Height          =   255
      Left            =   7800
      TabIndex        =   13
      Top             =   720
      Width           =   495
   End
   Begin VB.Label Label1 
      Caption         =   "Mean"
      Height          =   255
      Left            =   7800
      TabIndex        =   12
      Top             =   240
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p(800, 800) As Long
Dim g(500) As Long
Dim w, h As Long
Dim prob(500) As Double


Private Sub Command1_Click()

w = Picture1.ScaleWidth
h = Picture1.ScaleHeight

For i = 0 To w
   For j = 0 To h
    k = grayfromcolor(Picture1.Point(i, j))
    p(i, j) = k
    Picture2.PSet (i, j), RGB(k, k, k)
   Next
Next

End Sub

Private Sub Command10_Click()
 Dim pb(800, 800) As Long
 Dim pb2(800, 800) As Long
 

 
  For i = 0 To w
    For j = 0 To h
      If (p(i, j) > 127) Then
        pb(i, j) = 255
        pb2(i, j) = 255
    Else
      pb(i, j) = 0
      pb2(i, j) = 0
    End If
    Next j
  Next i
  
    Dim aa(9) As Long
For kkk = 1 To 5
  For i = 1 To w
    For j = 1 To h
  aa(0) = pb(i - 1, j - 1)
  aa(1) = pb(i, j - 1)
  aa(2) = pb(i + 1, j - 1)
  aa(3) = pb(i - 1, j)
  aa(4) = pb(i, j)
  aa(5) = pb(i + 1, j)
  aa(6) = pb(i - 1, j + 1)
  aa(7) = pb(i, j + 1)
  aa(8) = pb(i + 1, j + 1)
  
    
  If (pb(i, j) = 255) Then
    
   If ((aa(0) = 0) Or (aa(1) = 0) Or (aa(2) = 0) Or (aa(3) = 0) Or (aa(5) = 0) Or (aa(6) = 0) Or (aa(7) = 0) Or (aa(8) = 0)) Then
       pb2(i, j) = 0
   End If
   
  End If
  
Next j
Next i

For i = 1 To w
  For j = 1 To h
      ss = pb2(i, j)
      Picture8.PSet (i, j), RGB(ss, ss, ss)
  Next j
Next i

For i = 1 To w
  For j = 1 To h
     pb(i, j) = pb2(i, j)
  Next j
Next i
 Next kkk
End Sub

Private Sub Command2_Click()
cd.ShowOpen
Picture1.Picture = LoadPicture(cd.FileName)

End Sub

Public Function grayfromcolor(color As Long) As Integer
  Dim blue, green, red As Double
  
  blue = Fix((color / 256) / 256)
  green = Fix((color - ((blue * 256) * 256)) / 256)
  red = Fix((color - ((blue * 256) * 256) - (green * 256)))
  
  grayfromcolor = grayfromcolor + blue
  num = num + 1
  grayfromcolor = grayfromcolor + green
  num = num + 1
  grayfromcolor = grayfromcolor + red
  num = num + 1
  grayfromcolor = grayfromcolor / num
  
End Function

Private Sub Command3_Click()
  For i = 0 To 255
    g(i) = 0
  Next i
  
   For i = 0 To w
     For j = 0 To h
      g(p(i, j)) = g(p(i, j)) + 1
    Next j
Next i

For k = 0 To 255
  Text1.Text = Text1.Text & g(k) & "  "
Next k
ph = Picture3.ScaleHeight

For i = 0 To 255
  ss = ph - g(i)
    For j = ph To ss Step -1
       Picture3.PSet (i, j), RGB(100, 100, 100)
    Next j
  
Next i
  
End Sub

Private Sub Command4_Click()

w = Picture1.ScaleWidth
h = Picture1.ScaleWidth

   For i = 0 To 255
      prob(i) = g(i) / (w * h)
   Next i
   
   Mean = 0
   s = 0
   energy = 0
   entropy = 0
   
   For i = 0 To 255
     Mean = Mean + (i * prob(i))
     energy = energy + ((prob(i)) ^ 2)
    If (prob(i) <> 0) Then
     ss = Log(prob(i)) / Log(2)
       entropy = entropy + (prob(i) * ss)
    End If
   Next i
   entropy = -entropy
   
   For i = 0 To 255
     s = s + ((i - Mean) ^ 2) * prob(i)
   Next i
     std = Sqr(s)

  Text2.Text = Mean
  Text3.Text = std
  Text4.Text = energy
  Text5.Text = entropy

End Sub

Private Sub Command5_Click()
  Min = Val(Text6.Text)
  Max = Val(Text7.Text)
   
  ma = p(0, 0)
  mi = p(0, 0)
  
  For i = 0 To w
    For j = 0 To h
      If (p(i, j) > ma) Then
        ma = p(i, j)
      End If
      If (p(i, i) < mi) Then
        mi = p(i, j)
      End If
    Next j
  Next i

For i = 0 To w
  For j = 0 To h
     s = (((p(i, j) - mi) / (ma - mi)) * (Max - Min)) + Min
     Picture4.PSet (i, j), RGB(s, s, s)
  Next j
Next i


End Sub

Private Sub Command6_Click()
  Min = Val(Text8.Text)
  Max = Val(Text9.Text)
  
  ma = p(0, 0)
  mi = p(0, 0)
  
  For i = 0 To w
    For j = 0 To h
      If (p(i, j) > ma) Then
        ma = p(i, j)
      End If
      If (p(i, i) < mi) Then
        mi = p(i, j)
      End If
    Next j
  Next i
  
 For i = 0 To w
   For j = 0 To h
     s = (((Max - Min) / (ma - mi)) * (p(i, j) - mi)) + Min
     Picture5.PSet (i, j), RGB(s, s, s)
   Next j
Next i
  
End Sub

Private Sub Command7_Click()
  slide = Val(Text10.Text)
  
  For i = 0 To w
    For j = 0 To h
        p(i, j) = p(i, j) + slide
        Picture6.PSet (i, j), RGB(p(i, j), p(i, j), p(i, j))
    Next j
  Next i
End Sub

Private Sub Command8_Click()
  k1 = Val(Text11.Text)
  k2 = Val(Text12.Text)
  
 For i = 0 To w
    For j = 0 To h
      GM = GM + p(i, j)
    Next j
  Next i
  
  GM = GM / (w * h)
  Text13.Text = GM
  Dim ar(9) As Long
  
  For i = 1 To w
    For j = 1 To h
  ar(0) = p(i - 1, j - 1)
  ar(1) = p(i, j - 1)
  ar(2) = p(i + 1, j - 1)
  ar(3) = p(i - 1, j)
  ar(4) = p(i, j)
  ar(5) = p(i + 1, j)
  ar(6) = p(i - 1, j + 1)
  ar(7) = p(i, j + 1)
  ar(8) = p(i + 1, j + 1)
   
 ml = 0
 stdl = 0
 
    For ii = 0 To 8
      ml = ml + ar(ii)
    Next ii
    ml = ml / 9
    
   For ii = 0 To 8
    stdl = stdl + ((ar(ii) - ml) ^ 2)
    Next ii
    stdl = stdl / 9
    stdl = Sqr(stdl)
    
    If (stdl = 0) Then
      stdl = 1
    End If
    
 
      ace = Abs((k1 * ((GM / stdl) * (p(i, j) - ml))) + (k2 * (ml)))
      
      Picture7.PSet (i, j), RGB(ace, ace, ace)
      
    Next j
  Next i
  
End Sub

Private Sub Command9_Click()
 Dim pb(800, 800) As Long
 Dim pb2(800, 800) As Long
 

 
  For i = 0 To w
    For j = 0 To h
      If (p(i, j) > 127) Then
        pb(i, j) = 255
        pb2(i, j) = 255
    Else
      pb(i, j) = 0
      pb2(i, j) = 0
    End If
    Next j
  Next i
  
    Dim aa(9) As Long
For kkk = 1 To 5
  For i = 1 To w
    For j = 1 To h
  aa(0) = pb(i - 1, j - 1)
  aa(1) = pb(i, j - 1)
  aa(2) = pb(i + 1, j - 1)
  aa(3) = pb(i - 1, j)
  aa(4) = pb(i, j)
  aa(5) = pb(i + 1, j)
  aa(6) = pb(i - 1, j + 1)
  aa(7) = pb(i, j + 1)
  aa(8) = pb(i + 1, j + 1)
  
    
  If (pb(i, j) = 0) Then
    
   If ((aa(0) = 255) Or (aa(1) = 255) Or (aa(2) = 255) Or (aa(3) = 255) Or (aa(5) = 255) Or (aa(6) = 255) Or (aa(7) = 255) Or (aa(8) = 255)) Then
       pb2(i, j) = 255
   End If
   
  End If
  
Next j
Next i

For i = 1 To w
  For j = 1 To h
      ss = pb2(i, j)
      Picture8.PSet (i, j), RGB(ss, ss, ss)
  Next j
Next i

For i = 1 To w
  For j = 1 To h
     pb(i, j) = pb2(i, j)
  Next j
Next i
 Next kkk
  
End Sub
