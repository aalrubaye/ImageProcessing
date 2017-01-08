VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   8850
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11340
   LinkTopic       =   "Form1"
   ScaleHeight     =   8850
   ScaleWidth      =   11340
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text13 
      Height          =   375
      Left            =   10320
      TabIndex        =   38
      Top             =   5160
      Width           =   855
   End
   Begin VB.TextBox Text12 
      Height          =   375
      Left            =   10320
      TabIndex        =   37
      Top             =   4680
      Width           =   855
   End
   Begin VB.PictureBox Picture10 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   8280
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   35
      Top             =   4680
      Width           =   1995
   End
   Begin VB.TextBox Text11 
      Height          =   375
      Left            =   5280
      TabIndex        =   34
      Top             =   8280
      Width           =   1935
   End
   Begin VB.TextBox Text10 
      Height          =   375
      Left            =   2280
      TabIndex        =   33
      Top             =   8280
      Width           =   1935
   End
   Begin VB.TextBox Text9 
      Height          =   375
      Left            =   120
      TabIndex        =   32
      Top             =   8280
      Width           =   1935
   End
   Begin VB.TextBox Text8 
      Height          =   375
      Left            =   7320
      TabIndex        =   31
      Top             =   4800
      Width           =   855
   End
   Begin VB.TextBox Text7 
      Height          =   375
      Left            =   4320
      TabIndex        =   30
      Top             =   5280
      Width           =   855
   End
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   4320
      TabIndex        =   29
      Top             =   4800
      Width           =   855
   End
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   5280
      ScaleHeight     =   300
      ScaleMode       =   0  'User
      ScaleWidth      =   300
      TabIndex        =   27
      Top             =   6600
      Width           =   1995
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   5280
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   25
      Top             =   4680
      Width           =   1995
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   2280
      ScaleHeight     =   300
      ScaleMode       =   0  'User
      ScaleWidth      =   300
      TabIndex        =   23
      Top             =   6600
      Width           =   1995
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   2280
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   21
      Top             =   4680
      Width           =   1995
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      ScaleHeight     =   300
      ScaleMode       =   0  'User
      ScaleWidth      =   300
      TabIndex        =   19
      Top             =   6600
      Width           =   1995
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   17
      Top             =   4680
      Width           =   1995
   End
   Begin VB.TextBox Text2 
      Height          =   375
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   12
      Top             =   2280
      Width           =   2415
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   11
      Top             =   2760
      Width           =   2415
   End
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   10
      Top             =   3240
      Width           =   2415
   End
   Begin VB.TextBox Text5 
      Height          =   375
      Left            =   3240
      Locked          =   -1  'True
      TabIndex        =   9
      Top             =   3720
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Height          =   375
      Left            =   120
      TabIndex        =   8
      Top             =   3960
      Width           =   1935
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      ScaleHeight     =   300
      ScaleMode       =   0  'User
      ScaleWidth      =   300
      TabIndex        =   6
      Top             =   2280
      Width           =   1995
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   2640
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   3
      Top             =   360
      Width           =   1995
   End
   Begin VB.CommandButton browse 
      Caption         =   "Browse"
      Height          =   400
      Left            =   4680
      TabIndex        =   2
      Top             =   360
      Width           =   975
   End
   Begin VB.CommandButton move 
      Caption         =   "MoveGray"
      Height          =   400
      Left            =   4680
      TabIndex        =   1
      Top             =   840
      Width           =   975
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   5760
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   0
      Top             =   360
      Width           =   1995
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   4920
      Top             =   1440
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label10 
      Caption         =   "ACE"
      Height          =   255
      Left            =   9000
      TabIndex        =   36
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label9 
      Caption         =   "Histogram after Slide"
      Height          =   255
      Left            =   5520
      TabIndex        =   28
      Top             =   6360
      Width           =   1575
   End
   Begin VB.Label Label8 
      Caption         =   "Slide"
      Height          =   255
      Left            =   6000
      TabIndex        =   26
      Top             =   4440
      Width           =   375
   End
   Begin VB.Label Label7 
      Caption         =   "Histogram after Shrink"
      Height          =   255
      Left            =   2400
      TabIndex        =   24
      Top             =   6360
      Width           =   1695
   End
   Begin VB.Label Label6 
      Caption         =   "Shrink"
      Height          =   255
      Left            =   3000
      TabIndex        =   22
      Top             =   4440
      Width           =   495
   End
   Begin VB.Label Label5 
      Caption         =   "Histogram after Streach"
      Height          =   255
      Left            =   240
      TabIndex        =   20
      Top             =   6360
      Width           =   1935
   End
   Begin VB.Label Label4 
      Caption         =   "STREACH"
      Height          =   255
      Left            =   720
      TabIndex        =   18
      Top             =   4440
      Width           =   855
   End
   Begin VB.Label Label27 
      Caption         =   "mean"
      Height          =   255
      Left            =   2520
      TabIndex        =   16
      Top             =   2400
      Width           =   615
   End
   Begin VB.Label Label28 
      Caption         =   "STD"
      Height          =   255
      Left            =   2520
      TabIndex        =   15
      Top             =   2880
      Width           =   615
   End
   Begin VB.Label Label29 
      Caption         =   "energy"
      Height          =   255
      Left            =   2520
      TabIndex        =   14
      Top             =   3360
      Width           =   615
   End
   Begin VB.Label Label30 
      Caption         =   "entropy"
      Height          =   255
      Left            =   2520
      TabIndex        =   13
      Top             =   3840
      Width           =   615
   End
   Begin VB.Label Label3 
      Caption         =   "Histogram"
      Height          =   255
      Left            =   720
      TabIndex        =   7
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "ORIGINAL PICTURE"
      Height          =   255
      Left            =   2880
      TabIndex        =   5
      Top             =   120
      Width           =   1695
   End
   Begin VB.Label Label2 
      Caption         =   "GRAY PICTURE"
      Height          =   255
      Left            =   6120
      TabIndex        =   4
      Top             =   120
      Width           =   1335
   End
   Begin VB.Menu hhisto 
      Caption         =   "Histogram"
      Begin VB.Menu shh 
         Caption         =   "Show Histogram"
         Begin VB.Menu opic 
            Caption         =   "Original Picture"
         End
         Begin VB.Menu astr 
            Caption         =   "after Streach"
         End
         Begin VB.Menu ashrnk 
            Caption         =   "after Shrink"
         End
         Begin VB.Menu asl 
            Caption         =   "after Slide"
         End
      End
      Begin VB.Menu hff 
         Caption         =   "Histogram Features"
      End
      Begin VB.Menu hm 
         Caption         =   "Histogram Modification"
         Begin VB.Menu strch 
            Caption         =   "Streach"
         End
         Begin VB.Menu shrnk 
            Caption         =   "Shrink"
         End
         Begin VB.Menu sld 
            Caption         =   "Slide"
         End
      End
   End
   Begin VB.Menu ace 
      Caption         =   "Adaptive Contrast Enhancement"
      Begin VB.Menu aac 
         Caption         =   "ACE"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim p(800, 800) As Long
Dim p2(800, 800) As Long
Dim p3(800, 800) As Long
Dim p4(800, 800) As Long

Dim w, h, i, j As Long
Dim g(256) As Long
Dim g2(256) As Long
Dim g3(256) As Long
Dim g4(256) As Long

Dim prob(256) As Double

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

Private Sub aac_Click()
k1 = Val(Text12.Text)
k2 = Val(Text13.Text)
globalM = Val(Text2.Text)

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
  
  m = 0
  s = 0
  
  
 For ii = 0 To 8
   m = m + (g(ar(ii)) * prob(g(ar(ii))))
 Next ii
 
 For jj = 0 To 8
   s = s + (((g(ar(jj)) - m) ^ 2) * prob(g(ar(jj))))
 Next jj
 
 s = Sqr(s)
 
 adce = k1 * (globalM / s) * (p(i, j) - m) * (k2 * m)
  
  Picture10.PSet (i, j), RGB(adce, adce, adce)
  
  Next j
Next i

End Sub

Private Sub ashrnk_Click()
  
For k = 0 To 255
  g3(k) = 0
Next k
  
  For i = 0 To w
    For j = 0 To h
      g3(p3(i, j)) = g3(p3(i, j)) + 1
    Next j
Next i

For k = 0 To 255
  Text10.Text = Text10.Text & g3(k) & "  "
Next k
ph = Picture3.ScaleHeight

For i = 0 To 255
  
   ss = ph - g3(i)
   For j = ph To ss Step -1
     Picture7.PSet (i, j), RGB(150, 150, 150)
   Next j
     
Next i

End Sub

Private Sub asl_Click()
  
For k = 0 To 255
  g4(k) = 0
Next k
  
  For i = 0 To w
    For j = 0 To h
      g4(p4(i, j)) = g4(p4(i, j)) + 1
    Next j
Next i

For k = 0 To 255
  Text11.Text = Text11.Text & g4(k) & "  "
Next k
ph = Picture3.ScaleHeight

For i = 0 To 255
  
   ss = ph - g4(i)
   For j = ph To ss Step -1
     Picture9.PSet (i, j), RGB(150, 150, 150)
   Next j
     
Next i

End Sub

Private Sub astr_Click()
  
For k = 0 To 255
  g2(k) = 0
Next k
  
  For i = 0 To w
    For j = 0 To h
      g2(p2(i, j)) = g2(p2(i, j)) + 1
    Next j
Next i

For k = 0 To 255
  Text9.Text = Text9.Text & g2(k) & "  "
Next k
ph = Picture3.ScaleHeight

For i = 0 To 255
  
   ss = ph - g2(i)
   For j = ph To ss Step -1
     Picture5.PSet (i, j), RGB(150, 150, 150)
   Next j
     
Next i

End Sub

Private Sub browse_Click()
cd.ShowOpen
Picture1.Picture = LoadPicture(cd.FileName)
End Sub

Private Sub hff_Click()
    w = Picture3.ScaleWidth
    h = Picture3.ScaleHeight
      
  For i = 0 To 255
    prob(i) = g(i) / (w * h)
  Next i
  
  mean = 0
  std = 0
  energy = 0
  entropy = 0
  
 For i = 0 To 255
   mean = mean + (i * prob(i))
   energy = energy + prob(i) ^ 2
   If (prob(i) <> 0) Then
     ps = Log(prob(i)) / Log(2)
     entropy = entropy + (prob(i) * ps)
    End If
 Next i
 
 For i = 0 To 255
   std = std + (((i - mean) ^ 2) * prob(i))
 Next i
 
 std = Sqr(std)
 entropy = -entropy
 
 Text2.Text = mean
 Text3.Text = std
 Text4.Text = energy
 Text5.Text = entropy
 
End Sub



Private Sub move_Click()

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

Private Sub opic_Click()
  
For k = 0 To 255
  g(k) = 0
Next k
  
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
     Picture3.PSet (i, j), RGB(150, 150, 150)
   Next j
     
Next i


End Sub

Private Sub shrnk_Click()
shrinkmax = Val(Text7.Text)
shrinkmin = Val(Text6.Text)

  maxpixel = p(0, 0)
  minpixel = p(0, 0)
  
  For i = 0 To w
    For j = 0 To h
       
       If (p(i, j) > maxpixel) Then
         maxpixel = p(i, j)
       End If
       If (p(i, j) < minpixel) Then
         minpixel = p(i, j)
       End If
       
    Next j
  Next i
  
  
  For i = 0 To w
    For j = 0 To h
      k = p(i, j)
      ss = ((shrinkmax - shrinkmin) / (maxpixel - minpixel)) * (k - minpixel) + (shrinkmin)
      p3(i, j) = ss
      Picture6.PSet (i, j), RGB(ss, ss, ss)
    Next j
  Next i

End Sub

Private Sub sld_Click()
slide = Val(Text8.Text)
For i = 0 To w
  For j = 0 To h
    k = p(i, j)
    ss = k + slide
    If (ss > 255) Then
    ss = 255
    End If
    If (ss < 0) Then
    ss = 0
    End If
    p4(i, j) = ss
    Picture8.PSet (i, j), RGB(ss, ss, ss)
  Next j
Next i
End Sub

Private Sub strch_Click()
     
  maxpixel = p(0, 0)
  minpixel = p(0, 0)
  
  For i = 0 To w
    For j = 0 To h
       
       If (p(i, j) > maxpixel) Then
         maxpixel = p(i, j)
       End If
       If (p(i, j) < minpixel) Then
         minpixel = p(i, j)
       End If
       
    Next j
  Next i
  
  For i = 0 To w
    For j = 0 To h
          
      k = p(i, j)
      ss = ((k - minpixel) / (maxpixel - minpixel)) * 255 + minpixel
      p2(i, j) = ss
      
      Picture4.PSet (i, j), RGB(ss, ss, ss)

    Next j
 Next i
   
End Sub
