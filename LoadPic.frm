VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form Form1 
   Caption         =   "Filters in Image Processing"
   ClientHeight    =   9795
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   14475
   LinkTopic       =   "Form1"
   ScaleHeight     =   653
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   965
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture9 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   53
      Top             =   8040
      Width           =   1995
   End
   Begin VB.PictureBox Picture8 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   2160
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   52
      Top             =   8040
      Width           =   1995
   End
   Begin VB.PictureBox Picture27 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   8280
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   50
      Top             =   8040
      Width           =   1995
   End
   Begin VB.PictureBox Picture26 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   6240
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   48
      Top             =   8040
      Width           =   1995
   End
   Begin VB.PictureBox Picture25 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   4200
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   46
      Top             =   8040
      Width           =   1995
   End
   Begin VB.PictureBox Picture24 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   12360
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   44
      Top             =   6120
      Width           =   1995
   End
   Begin VB.PictureBox Picture23 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   10320
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   42
      Top             =   6120
      Width           =   1995
   End
   Begin VB.CommandButton Command3 
      Caption         =   "MoveGray"
      Height          =   400
      Left            =   3120
      TabIndex        =   41
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Browse"
      Height          =   400
      Left            =   3120
      TabIndex        =   39
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox Picture22 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   1080
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   38
      Top             =   240
      Width           =   1995
   End
   Begin VB.PictureBox Picture21 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   8280
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   36
      Top             =   6120
      Width           =   1995
   End
   Begin VB.PictureBox Picture20 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   6240
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   34
      Top             =   6120
      Width           =   1995
   End
   Begin VB.PictureBox Picture19 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   4200
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   32
      Top             =   6120
      Width           =   1995
   End
   Begin VB.PictureBox Picture18 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   2160
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   30
      Top             =   6120
      Width           =   1995
   End
   Begin VB.PictureBox Picture17 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   120
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   29
      Top             =   6120
      Width           =   1995
   End
   Begin VB.PictureBox Picture16 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   12360
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   28
      Top             =   4200
      Width           =   1995
   End
   Begin VB.PictureBox Picture15 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   10320
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   26
      Top             =   4200
      Width           =   1995
   End
   Begin VB.PictureBox Picture14 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   8280
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   24
      Top             =   4200
      Width           =   1995
   End
   Begin VB.PictureBox Picture7 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   10320
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   23
      Top             =   2280
      Width           =   1995
   End
   Begin VB.PictureBox Picture6 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   8280
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   22
      Top             =   2280
      Width           =   1995
   End
   Begin VB.PictureBox Picture13 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   6240
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   21
      Top             =   4200
      Width           =   1995
   End
   Begin VB.PictureBox Picture12 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   4200
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   20
      Top             =   4200
      Width           =   1995
   End
   Begin VB.PictureBox Picture11 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   12360
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   19
      Top             =   2280
      Width           =   1995
   End
   Begin VB.PictureBox Picture5 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   6240
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   18
      Top             =   2280
      Width           =   1995
   End
   Begin VB.PictureBox Picture4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   4200
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   17
      Top             =   2280
      Width           =   1995
   End
   Begin VB.PictureBox Picture2 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   7440
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   16
      Top             =   240
      Width           =   1995
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   6600
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox Picture3 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      ScaleHeight     =   231
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   263
      TabIndex        =   5
      Top             =   2280
      Width           =   3975
   End
   Begin VB.CommandButton move 
      Caption         =   "MoveGray"
      Height          =   400
      Left            =   6360
      TabIndex        =   4
      Top             =   720
      Width           =   975
   End
   Begin VB.CommandButton browse 
      Caption         =   "Browse"
      Height          =   400
      Left            =   6360
      TabIndex        =   3
      Top             =   240
      Width           =   975
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   4320
      ScaleHeight     =   103
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   131
      TabIndex        =   0
      Top             =   240
      Width           =   1995
   End
   Begin VB.Label Label26 
      Caption         =   "Mult"
      Height          =   255
      Left            =   840
      TabIndex        =   55
      Top             =   7800
      Width           =   375
   End
   Begin VB.Label Label25 
      Caption         =   "Div"
      Height          =   255
      Left            =   3000
      TabIndex        =   54
      Top             =   7800
      Width           =   255
   End
   Begin VB.Label Label24 
      Caption         =   "Logic NOT"
      Height          =   255
      Left            =   8880
      TabIndex        =   51
      Top             =   7800
      Width           =   975
   End
   Begin VB.Label Label23 
      Caption         =   "Logic OR"
      Height          =   255
      Left            =   6840
      TabIndex        =   49
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label22 
      Caption         =   "Logic And"
      Height          =   255
      Left            =   4800
      TabIndex        =   47
      Top             =   7800
      Width           =   735
   End
   Begin VB.Label Label17 
      Caption         =   "Sub with pic2"
      Height          =   255
      Left            =   12840
      TabIndex        =   45
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label16 
      Caption         =   "Add with pic2"
      Height          =   255
      Left            =   10800
      TabIndex        =   43
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label15 
      Caption         =   "GRAY PICTURE 2"
      Height          =   255
      Left            =   1320
      TabIndex        =   40
      Top             =   0
      Width           =   1455
   End
   Begin VB.Label Label21 
      Caption         =   "Sobel"
      Height          =   255
      Left            =   9000
      TabIndex        =   37
      Top             =   5880
      Width           =   495
   End
   Begin VB.Label Label20 
      Caption         =   "Sobel Column"
      Height          =   255
      Left            =   6720
      TabIndex        =   35
      Top             =   5880
      Width           =   975
   End
   Begin VB.Label Label19 
      Caption         =   "Sobel Row"
      Height          =   255
      Left            =   4800
      TabIndex        =   33
      Top             =   5880
      Width           =   855
   End
   Begin VB.Label Label10 
      Caption         =   "ROBERT"
      Height          =   255
      Left            =   2760
      TabIndex        =   31
      Top             =   5880
      Width           =   735
   End
   Begin VB.Label Label18 
      Caption         =   "BIT REDUCTION (OR)"
      Height          =   255
      Left            =   10440
      TabIndex        =   27
      Top             =   3960
      Width           =   1695
   End
   Begin VB.Label Label14 
      Caption         =   "DIFFERENCE"
      Height          =   255
      Left            =   10800
      TabIndex        =   25
      Top             =   2040
      Width           =   1095
   End
   Begin VB.Label Label13 
      Caption         =   "spatial-cordinates MEDAIN"
      Height          =   255
      Left            =   120
      TabIndex        =   15
      Top             =   5880
      Width           =   1935
   End
   Begin VB.Label Label12 
      Caption         =   "spatial-cordinates AVERAGE"
      Height          =   255
      Left            =   12360
      TabIndex        =   14
      Top             =   3960
      Width           =   2055
   End
   Begin VB.Label Label11 
      Caption         =   "LAPLACE"
      Height          =   255
      Left            =   8880
      TabIndex        =   13
      Top             =   2040
      Width           =   735
   End
   Begin VB.Label Label9 
      Caption         =   "BIT REDUCTION (AND)"
      Height          =   255
      Left            =   8400
      TabIndex        =   12
      Top             =   3960
      Width           =   1935
   End
   Begin VB.Label Label8 
      Caption         =   "THRESHOLDING"
      Height          =   255
      Left            =   6600
      TabIndex        =   11
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label Label7 
      Caption         =   "MEDIAN"
      Height          =   255
      Left            =   4800
      TabIndex        =   10
      Top             =   3960
      Width           =   735
   End
   Begin VB.Label Label6 
      Caption         =   "MEAN"
      Height          =   255
      Left            =   13080
      TabIndex        =   9
      Top             =   2040
      Width           =   615
   End
   Begin VB.Label Label5 
      Caption         =   "CON.ZERO ORD"
      Height          =   255
      Left            =   6600
      TabIndex        =   8
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label4 
      Caption         =   "CON.FIRST ORD"
      Height          =   255
      Left            =   4440
      TabIndex        =   7
      Top             =   2040
      Width           =   1335
   End
   Begin VB.Label Label3 
      Caption         =   "ZOOMING"
      Height          =   255
      Left            =   1560
      TabIndex        =   6
      Top             =   2040
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "GRAY PICTURE"
      Height          =   255
      Left            =   7800
      TabIndex        =   2
      Top             =   0
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "ORIGINAL PICTURE"
      Height          =   255
      Left            =   4560
      TabIndex        =   1
      Top             =   0
      Width           =   1695
   End
   Begin VB.Menu mnufile 
      Caption         =   "&File"
      Begin VB.Menu mnuexit 
         Caption         =   "e&xit"
      End
   End
   Begin VB.Menu mnufilter 
      Caption         =   "Fil&ter"
      Begin VB.Menu mathmnu 
         Caption         =   "Arithmatic &Operation"
         Begin VB.Menu addmnuq 
            Caption         =   "Add with pic"
         End
         Begin VB.Menu submnuq 
            Caption         =   "Sub with pic"
         End
         Begin VB.Menu mulmnu 
            Caption         =   "Mult"
         End
         Begin VB.Menu divmnu 
            Caption         =   "Div"
         End
      End
      Begin VB.Menu mnulogy 
         Caption         =   "Logic Operation"
         Begin VB.Menu mnula 
            Caption         =   "AND"
         End
         Begin VB.Menu mnulor 
            Caption         =   "OR"
         End
         Begin VB.Menu mnulnot 
            Caption         =   "NOT"
         End
      End
      Begin VB.Menu mnuzooming 
         Caption         =   "&Zooming"
         Begin VB.Menu mnurow 
            Caption         =   "Row"
         End
         Begin VB.Menu mnu2dzoom 
            Caption         =   "2d Zoom"
         End
         Begin VB.Menu mnuaverage 
            Caption         =   "Average"
         End
      End
      Begin VB.Menu mnuconvelution 
         Caption         =   "&Convelution"
         Begin VB.Menu mnuforder 
            Caption         =   "First Order"
         End
         Begin VB.Menu mnuzorder 
            Caption         =   "Zero Order"
         End
      End
      Begin VB.Menu mnuspatial 
         Caption         =   "&Spatial"
         Begin VB.Menu mnuenhancment 
            Caption         =   "Enhancment"
            Begin VB.Menu mnulaplace 
               Caption         =   "Laplace"
            End
            Begin VB.Menu mnudifference 
               Caption         =   "Difference"
               Begin VB.Menu mnuhorizontal 
                  Caption         =   "Horizontal"
               End
               Begin VB.Menu mnuvertical 
                  Caption         =   "Vertical"
               End
               Begin VB.Menu mnumaind 
                  Caption         =   "main d"
               End
               Begin VB.Menu munsecondd 
                  Caption         =   "second d"
               End
            End
         End
         Begin VB.Menu mnumean 
            Caption         =   "Mean"
         End
         Begin VB.Menu mnumedain 
            Caption         =   "Medain"
         End
      End
      Begin VB.Menu mnuquantization 
         Caption         =   "&Quantization"
         Begin VB.Menu mnugraylevel 
            Caption         =   "Gray Level"
            Begin VB.Menu mnuthresholding 
               Caption         =   "Thresholding"
            End
            Begin VB.Menu mnubitreduction 
               Caption         =   "Bit Reduction"
               Begin VB.Menu mnuand 
                  Caption         =   "AND"
               End
               Begin VB.Menu mnuor 
                  Caption         =   "OR"
               End
            End
         End
         Begin VB.Menu mnuscordinates 
            Caption         =   "Spatial Cordinates"
            Begin VB.Menu mnusaverage 
               Caption         =   "Average"
            End
            Begin VB.Menu mnusmedain 
               Caption         =   "Medain"
            End
            Begin VB.Menu mnudecination 
               Caption         =   "Decination"
               Enabled         =   0   'False
            End
         End
      End
      Begin VB.Menu edgmnu 
         Caption         =   "&Edge detection"
         Begin VB.Menu robertmnu 
            Caption         =   "Robert"
            Begin VB.Menu bymmnu 
               Caption         =   "by Mask"
            End
            Begin VB.Menu bysmnu 
               Caption         =   "by Sqr"
            End
            Begin VB.Menu byamnu 
               Caption         =   "by Abs"
            End
         End
         Begin VB.Menu sobmnu 
            Caption         =   "Sobel"
            Begin VB.Menu hormnu 
               Caption         =   "Horizontal (Row)"
               Begin VB.Menu bmu 
                  Caption         =   "by Mask"
               End
               Begin VB.Menu bam 
                  Caption         =   "by Abs "
               End
            End
            Begin VB.Menu vermnu 
               Caption         =   "Vertical (Col)"
               Begin VB.Menu bmu2 
                  Caption         =   "by Mask"
               End
               Begin VB.Menu bam2 
                  Caption         =   "by Abs"
               End
            End
            Begin VB.Menu bsmnu 
               Caption         =   "by Sqr"
               Begin VB.Menu bmu3 
                  Caption         =   "by Mask"
               End
               Begin VB.Menu baum 
                  Caption         =   "by Abs"
               End
            End
         End
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
Dim pi2(800, 800) As Long
Dim pi3(800, 800) As Long
Dim cf(800, 800) As Long
Dim mask(8) As Double
Dim mask1(8) As Double
Dim mask2(3) As Long
Dim ss1(7) As String
Dim ss2(7) As String
Dim ss3(7) As String
Dim ar(8) As Long
Dim ar2(3) As Long
Dim ar3(1000) As Long
Dim w, h, i, j As Long

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

Private Sub addmnuq_Click()
For i = 1 To w
  For j = 1 To h
    pi3(i, j) = pi2(i, j) + p(i, j)
    If pi3(i, j) > 255 Then
       pi3(i, j) = 255
    End If
  
  Next j
Next i

For i = 1 To w
  For j = 1 To h
    hh = pi3(i, j)
    Picture23.PSet (i, j), RGB(hh, hh, hh)
  Next j
Next i

End Sub

Private Sub bam_Click()

For i = 1 To w
  For j = 1 To h
   Sum = Fix(Abs(p(i - 1, j + 1) - p(i - 1, j - 1)) + Abs(2 * p(i, j + 1) - 2 * p(i, j - 1)) + Abs(p(i + 1, j + 1) - p(i + 1, j - 1)))
   
   If (Sum > 255) Then
     Sum = 255
    End If
    If (Sum < 0) Then
      Sum = 0
    End If
    cf(i, j) = Sum

  Next j
Next i

For i = 1 To w
  For j = 1 To h
    s = cf(i, j)
    Picture19.PSet (i, j), RGB(s, s, s)
  Next j
Next i
End Sub

Private Sub bam2_Click()

For i = 1 To w
  For j = 1 To h
   Sum = Fix(Abs(p(i + 1, j - 1) - p(i - 1, j - 1)) + Abs(2 * p(i + 1, j) - 2 * p(i - 1, j)) + Abs(p(i + 1, j + 1) - p(i - 1, j + 1)))
   
   If (Sum > 255) Then
     Sum = 255
    End If
    If (Sum < 0) Then
      Sum = 0
    End If
    cf(i, j) = Sum

  Next j
Next i

For i = 1 To w
  For j = 1 To h
    s = cf(i, j)
    Picture20.PSet (i, j), RGB(s, s, s)
  Next j
Next i
End Sub

Private Sub baum_Click()

For i = 1 To w
  For j = 1 To h
   S1 = Fix(Abs(p(i - 1, j + 1) - p(i - 1, j - 1)) + Abs(2 * p(i, j + 1) - 2 * p(i, j - 1)) + Abs(p(i + 1, j + 1) - p(i + 1, j - 1)))
   S2 = Fix(Abs(p(i + 1, j - 1) - p(i - 1, j - 1)) + Abs(2 * p(i + 1, j) - 2 * p(i - 1, j)) + Abs(p(i + 1, j + 1) - p(i - 1, j + 1)))
   
   Sum = Fix(Sqr((S1 ^ 2) + (S2 ^ 2)))
   
   If (Sum > 255) Then
     Sum = 255
    End If
    If (Sum < 0) Then
      Sum = 0
    End If
    cf(i, j) = Sum

  Next j
Next i

For i = 1 To w
  For j = 1 To h
    s = cf(i, j)
    Picture21.PSet (i, j), RGB(s, s, s)
  Next j
Next i
End Sub

Private Sub bmu_Click()
mask(0) = -1
mask(1) = -2
mask(2) = -1
mask(3) = 0
mask(4) = 0
mask(5) = 0
mask(6) = 1
mask(7) = 2
mask(8) = 1

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
  Sum = 0
  For k = 0 To 8
    Sum = Sum + Fix(ar(k) * mask(k))
  Next k
  
   If (Sum > 255) Then
     Sum = 255
    End If
    If (Sum < 0) Then
      Sum = 0
    End If
    cf(i, j) = Sum

  Next j
Next i

For i = 1 To w
  For j = 1 To h
    s = cf(i, j)
    Picture19.PSet (i, j), RGB(s, s, s)
  Next j
Next i

End Sub

Private Sub bmu2_Click()
mask(0) = -1
mask(1) = 0
mask(2) = 1
mask(3) = -2
mask(4) = 0
mask(5) = 2
mask(6) = -1
mask(7) = 0
mask(8) = 1

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
  Sum = 0
  For k = 0 To 8
    Sum = Sum + Fix(ar(k) * mask(k))
  Next k
  
   If (Sum > 255) Then
     Sum = 255
    End If
    If (Sum < 0) Then
      Sum = 0
    End If
    cf(i, j) = Sum

  Next j
Next i

For i = 1 To w
  For j = 1 To h
    s = cf(i, j)
    Picture20.PSet (i, j), RGB(s, s, s)
  Next j
Next i
End Sub

Private Sub bmu3_Click()
mask(0) = -1
mask(1) = 0
mask(2) = 1
mask(3) = -2
mask(4) = 0
mask(5) = 2
mask(6) = -1
mask(7) = 0
mask(8) = 1

mask1(0) = -1
mask1(1) = -2
mask1(2) = -1
mask1(3) = 0
mask1(4) = 0
mask1(5) = 0
mask1(6) = 1
mask1(7) = 2
mask1(8) = 1


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
  
  S1 = 0
  For k = 0 To 8
    S1 = S1 + (ar(k) * mask(k))
  Next k
  
  S2 = 0
  For k = 0 To 8
    S2 = S2 + (ar(k) * mask1(k))
  Next k
  
  mn = Fix(Sqr((S1 ^ 2) + (S2 ^ 2)))
  If (mn > 255) Then
    mn = 255
   End If
   If (mn < 0) Then
    mn = 0
   End If
    cf(i, j) = mn

  Next j
Next i

For i = 1 To w
  For j = 1 To h
    s = cf(i, j)
    Picture21.PSet (i, j), RGB(s, s, s)
  Next j
Next i
End Sub

Private Sub browse_Click()
cd.ShowOpen
Picture1.Picture = LoadPicture(cd.FileName)
End Sub

Private Sub C1_Click()
Picture12.Picture = LoadPicture("")
zz = InputBox("enter picture width value", "WIDTH")
ff = InputBox("enter picture hight value", "HIGHT")
If (zz = "" Or ff = "") Then
MsgBox "Where is the picture Value?", vbCritical, "ERROR!"
Else
  bl = Val(zz)
  bl2 = Val(ff)
  z = (256 / bl)
  F = (256 / bl2)
  
  i = 1
  r = 1
  
  While (i < w)
    j = 1
    s = 1
    While (j < h)
      Sum = 0
      
      For k = i To i + (z - 1)
        For q = j To j + (F - 1)
         Sum = Sum + p(i, j)
        Next q
      Next k
      
      av = Fix(Sum / (z * F))
      cf(r, s) = av
      j = j + F
      s = s + 1
    Wend
    i = i + z
    r = r + 1
  Wend
  rr = r
  ss = s

  For r = 1 To rr
    For s = 1 To ss
     k = cf(r, s)
     Picture12.PSet (r, s), RGB(k, k, k)
    Next s
  Next r
  
End If
End Sub

Private Sub dif_Click()
mask(0) = 0
mask(1) = 1
mask(2) = 0
mask(3) = 0
mask(4) = 1
mask(5) = 0
mask(6) = 0
mask(7) = -1
mask(8) = 0

For i = 1 To w - 1
  For j = 1 To h - 1
  
  ar(0) = p(i - 1, j - 1)
  ar(1) = p(i, j - 1)
  ar(2) = p(i + 1, j - 1)
  ar(3) = p(i - 1, j)
  ar(4) = p(i, j)
  ar(5) = p(i + 1, j)
  ar(6) = p(i - 1, j + 1)
  ar(7) = p(i, j + 1)
  ar(8) = p(i + 1, j + 1)
  
  Sum = 0
  For k = 0 To 8
   Sum = Sum + (ar(k) * mask(k))
  Next k
    
   If (Sum < 0) Then Sum = 0
  cf(i, j) = Sum
  Next j
Next i

For i = 1 To w - 1
  For j = 1 To h - 1
    s = cf(i, j)
    Picture11.PSet (i, j), RGB(s, s, s)
  Next j
Next i

End Sub

Private Sub c2_Click()
Picture13.Picture = LoadPicture("")
zz = InputBox("enter picture width value", "WIDTH")
ff = InputBox("enter picture hight value", "HIGHT")
If (zz = "" Or ff = "") Then
MsgBox "Where is the picture Value?", vbCritical, "ERROR!"
Else
  bl = Val(zz)
  bl2 = Val(ff)
  z = (256 / bl)
  F = (256 / bl2)
  
  i = 1
  r = 1
  
  While (i < w)
    j = 1
    s = 1
    While (j < h)
      mm = 0
      For k = i To i + (z - 1)
        For q = j To j + (F - 1)
         ar3(mm) = p(k, q)
         mm = mm + 1
        Next q
      Next k
   
     For kk = 0 To (z - 1)
       For uu = 0 To (F - 1)
         If (ar3(kk) < ar3(uu)) Then
          temp = ar3(kk)
          ar3(kk) = ar3(uu)
          ar3(uu) = temp
         End If
       Next uu
     Next kk
   
        
       cf(r, s) = Fix(ar3(z / 2))
       
      j = j + F
      s = s + 1
    Wend
    i = i + z
    r = r + 1
  Wend
  rr = r
  ss = s

  For r = 1 To rr
    For s = 1 To ss
     k = cf(r, s)
     Picture13.PSet (r, s), RGB(k, k, k)
    Next s
  Next r
  
End If

End Sub

Private Sub byamnu_Click()
For i = 1 To w
  For j = 1 To h
     rob = Fix(Abs(p(i, j) - p(i - 1, j - 1)) + Abs(p(i, j - 1) - p(i - 1, j)))
     If (rob < 0) Then
     rob = 0
     End If
     If (rob > 255) Then
       rob = 255
    End If
     
     cf(i, j) = rob
     
  Next j
Next i

For i = 1 To w
  For j = 1 To h
    s = cf(i, j)
    Picture18.PSet (i, j), RGB(s, s, s)
  Next j
Next i
End Sub

Private Sub bymmnu_Click()
mask2(0) = -1
mask2(1) = -1
mask2(2) = 1
mask2(3) = 1
For i = 1 To w
  For j = 1 To h
    ar2(0) = p(i, j)
    ar2(1) = p(i + 1, j)
    ar2(2) = p(i, j + 1)
    ar2(3) = p(i + 1, j + 1)
    rob = 0
    For k = 0 To 3
     rob = rob + Fix(mask2(k) * ar2(k))
    Next k
     If (rob < 0) Then
       rob = 0
     End If
     If (rob > 255) Then
       rob = 255
    End If
     
     cf(i, j) = rob

  Next j
Next i

For i = 1 To w
  For j = 1 To h
    s = cf(i, j)
    Picture18.PSet (i, j), RGB(s, s, s)
  Next j
Next i
End Sub

Private Sub bysmnu_Click()
For i = 1 To w
  For j = 1 To h
     rob = Fix(Sqr((p(i, j) - p(i - 1, j - 1)) ^ 2 + (p(i, j - 1) - p(i - 1, j)) ^ 2))
     If (rob < 0) Then
     rob = 0
     End If
     If (rob > 255) Then
       rob = 255
    End If
     
     cf(i, j) = rob
     
  Next j
Next i

For i = 1 To w
  For j = 1 To h
    s = cf(i, j)
    Picture18.PSet (i, j), RGB(s, s, s)
  Next j
Next i
End Sub

Private Sub Command1_Click()
cd.ShowOpen
Picture22.Picture = LoadPicture(cd.FileName)
End Sub

Private Sub Command2_Click()
Picture22.Visible = False
Command1.Visible = False
Command2.Visible = False
Label15.Visible = False
Command3.Visible = False
End Sub


Private Sub Command3_Click()

w = Picture22.ScaleWidth
h = Picture22.ScaleHeight

For i = 0 To w
   For j = 0 To h
    k = grayfromcolor(Picture22.Point(i, j))
    pi2(i, j) = k
    Picture22.PSet (i, j), RGB(k, k, k)
   Next
Next
End Sub


Private Sub divmnu_Click()
x = InputBox("Enter Factor Value")

For i = 1 To w
  For j = 1 To h
    pi3(i, j) = Fix(p(i, j) / x)
    If pi3(i, j) < 0 Then
       pi3(i, j) = 0
    End If
  
  Next j
Next i

For i = 1 To w
  For j = 1 To h
    hh = pi3(i, j)
    Picture8.PSet (i, j), RGB(hh, hh, hh)
  Next j
Next i
End Sub

Private Sub mnu2dzoom_Click()
ss = 1
For i = 1 To w
  For j = 1 To h
    p2(ss, j) = p(i, j)
    p2(ss + 1, j) = p(i, j)
  Next j
  ss = ss + 2
Next i

tt = 1
For i = 1 To w * 2
  For j = 1 To h
    p3(i, tt) = p2(i, j)
    p3(i, tt + 1) = p2(i, j)
    tt = tt + 2
    Next j
    tt = 1
Next i


For i = 1 To w * 2
  For j = 1 To h * 2
    s = p3(i, j)
    Picture3.PSet (i, j), RGB(s, s, s)
  Next j
Next i
End Sub

Private Sub mnuand_Click()
xx = InputBox("enter reduction value", "REDUCTION")
If (xx = "") Then
MsgBox "Where is the Reduction Value?", vbCritical, "ERROR!"
Else
x = Val(xx)

For i = 0 To 7
  If (x = 2 ^ i) Then
    n = i
  End If
Next i

m = 256 - (2 ^ (8 - n))

ii = 7
While (ii >= 0)
 If (m >= 2 ^ ii) Then
   y = m - (2 ^ ii)
   S1 = S1 + "1"
   m = y
  Else
   S1 = S1 + "0"
 End If
   ii = ii - 1
Wend

  For k = 1 To 8
   ss1(k - 1) = Mid(S1, k, 1)
  Next k


shift = 8 - n

For i = 1 To w
  For j = 1 To h
    d = p(i, j)
    ii = 7
    S2 = ""
    While (ii >= 0)
     If (d >= 2 ^ ii) Then
       y = d - (2 ^ ii)
       S2 = S2 + "1"
       d = y
     Else
       S2 = S2 + "0"
     End If
     ii = ii - 1
    Wend
  
  For k = 1 To 8
    ss2(k - 1) = Mid(S2, k, 1)
  Next k
  
  For k = 0 To 7
    If (ss1(k) = "0" Or ss2(k) = "0") Then
      ss3(k) = "0"
    Else
      ss3(k) = "1"
    End If
  Next k
  qp = 0
  k = 0
  le = 7 - shift
  While (le >= 0)
      qp = qp + ss3(k) * (2 ^ le)
      k = k + 1
      le = le - 1
  Wend
   
  cf(i, j) = Fix((qp * 256) / xx)
   
  Next j
Next i


For i = 1 To w
  For j = 1 To h
    ws = cf(i, j)
    Picture14.PSet (i, j), RGB(ws, ws, ws)
    
  Next j
Next i
End If
End Sub

Private Sub mnuaverage_Click()
ss = 1
For i = 1 To w
  For j = 1 To h
   k = Fix(p(i, j) + p(i + 1, j) / 2)
   p2(ss, j) = p(i, j)
   p2(ss + 1, j) = k
  Next j
  ss = ss + 2
Next i

tt = 1
 For i = 1 To w * 2
   For j = 1 To h
    k = Fix(p2(i, j) + p2(i, j + 1) / 2)
    p3(i, tt) = p2(i, j)
    p3(i, tt + 1) = k
    tt = tt + 2
   Next j
 tt = 1
 Next i

For i = 1 To w * 2
  For j = 1 To h * 2
    s = p3(i, j)
    Picture3.PSet (i, j), RGB(s, s, s)
  Next j
Next i
End Sub

Private Sub mnuexit_Click()
End
End Sub

Private Sub mnuforder_Click()
mask(0) = 0.25
mask(1) = 0.5
mask(2) = 0.25
mask(3) = 0.5
mask(4) = 1
mask(5) = 0.5
mask(6) = 0.25
mask(7) = 0.5
mask(8) = 0.25

For i = 1 To w - 1
 For j = 1 To h - 1
  ar(0) = p(i - 1, j - 1)
  ar(1) = p(i, j - 1)
  ar(2) = p(i + 1, j - 1)
  ar(3) = p(i - 1, j)
  ar(4) = p(i, j)
  ar(5) = p(i + 1, j)
  ar(6) = p(i - 1, j + 1)
  ar(7) = p(i, j + 1)
  ar(8) = p(i + 1, j + 1)
  Sum = 0
  For k = 0 To 8
    Sum = Sum + Fix(ar(k) * mask(k))
  Next k
    cf(i, j) = Sum
 Next j
Next i

For i = 1 To w
  For j = 1 To h
    s = cf(i, j)
    Picture4.PSet (i, j), RGB(s, s, s)
  Next j
Next i
End Sub

Private Sub mnuhorizontal_Click()
mask(0) = 0
mask(1) = 0
mask(2) = 0
mask(3) = 1
mask(4) = 1
mask(5) = -1
mask(6) = 0
mask(7) = 0
mask(8) = 0

For i = 1 To w - 1
  For j = 1 To h - 1
  
  ar(0) = p(i - 1, j - 1)
  ar(1) = p(i, j - 1)
  ar(2) = p(i + 1, j - 1)
  ar(3) = p(i - 1, j)
  ar(4) = p(i, j)
  ar(5) = p(i + 1, j)
  ar(6) = p(i - 1, j + 1)
  ar(7) = p(i, j + 1)
  ar(8) = p(i + 1, j + 1)
  
  Sum = 0
  For k = 0 To 8
   Sum = Sum + (ar(k) * mask(k))
  Next k
    
   If (Sum < 0) Then Sum = 0
  cf(i, j) = Sum
  Next j
Next i

For i = 1 To w - 1
  For j = 1 To h - 1
    s = cf(i, j)
    Picture7.PSet (i, j), RGB(s, s, s)
  Next j
Next i
End Sub

Private Sub mnula_Click()
For i = 1 To w
  For j = 1 To h
    k1 = p(i, j)
    k2 = pi2(i, j)
    pi3(i, j) = k1 And k2
  Next
Next

For i = 1 To w
  For j = 1 To h
    sss = pi3(i, j)
    Picture25.PSet (i, j), RGB(sss, sss, sss)
  Next
Next

End Sub

Private Sub mnulaplace_Click()
mask(0) = 0
mask(1) = -1
mask(2) = 0
mask(3) = -1
mask(4) = 4
mask(5) = -1
mask(6) = 0
mask(7) = -1
mask(8) = 0

For i = 1 To w - 1
  For j = 1 To h - 1
  
  ar(0) = p(i - 1, j - 1)
  ar(1) = p(i, j - 1)
  ar(2) = p(i + 1, j - 1)
  ar(3) = p(i - 1, j)
  ar(4) = p(i, j)
  ar(5) = p(i + 1, j)
  ar(6) = p(i - 1, j + 1)
  ar(7) = p(i, j + 1)
  ar(8) = p(i + 1, j + 1)
  
  Sum = 0
  For k = 0 To 8
   Sum = Sum + (ar(k) * mask(k))
  Next k
    
   If (Sum < 0) Then Sum = 0
  cf(i, j) = Sum
  Next j
Next i

For i = 1 To w - 1
  For j = 1 To h - 1
    s = cf(i, j)
    Picture6.PSet (i, j), RGB(s, s, s)
  Next j
Next i
End Sub

Private Sub mnulnot_Click()
Dim S1 As String

For i = 1 To w
  For j = 1 To h
   m = p(i, j)
   S1 = ""
   ii = 7
   While (ii >= 0)
    If (m >= 2 ^ ii) Then
       y = m - (2 ^ ii)
       S1 = S1 + "1"
       m = y
      Else
       S1 = S1 + "0"
    End If
     ii = ii - 1
   Wend
     
  For k = 1 To 8
   ss1(k - 1) = Mid(S1, k, 1)
  Next k
        
    
  For k = 0 To 7
    If (ss1(k) = "1") Then
      ss2(k) = "0"
    Else
      ss2(k) = "1"
    End If
  Next k

 qp = 0
  k = 0
  le = 7
  While (le >= 0)
      qp = qp + (ss2(k) * (2 ^ le))
      k = k + 1
      le = le - 1
  Wend
   
  cf(i, j) = qp
     
  Next
Next

For i = 1 To w
  For j = 1 To h
    sss = cf(i, j)
    Picture27.PSet (i, j), RGB(sss, sss, sss)
  Next
Next


End Sub

Private Sub mnulor_Click()
For i = 1 To w
  For j = 1 To h
    k1 = p(i, j)
    k2 = pi2(i, j)
    pi3(i, j) = k1 Or k2
  Next
Next

For i = 1 To w
  For j = 1 To h
    sss = pi3(i, j)
    Picture26.PSet (i, j), RGB(sss, sss, sss)
  Next
Next
End Sub

Private Sub mnumaind_Click()
mask(0) = 1
mask(1) = 0
mask(2) = 0
mask(3) = 0
mask(4) = 1
mask(5) = 0
mask(6) = 0
mask(7) = 0
mask(8) = -1

For i = 1 To w - 1
  For j = 1 To h - 1
  
  ar(0) = p(i - 1, j - 1)
  ar(1) = p(i, j - 1)
  ar(2) = p(i + 1, j - 1)
  ar(3) = p(i - 1, j)
  ar(4) = p(i, j)
  ar(5) = p(i + 1, j)
  ar(6) = p(i - 1, j + 1)
  ar(7) = p(i, j + 1)
  ar(8) = p(i + 1, j + 1)
  
  Sum = 0
  For k = 0 To 8
   Sum = Sum + (ar(k) * mask(k))
  Next k
    
   If (Sum < 0) Then Sum = 0
  cf(i, j) = Sum
  Next j
Next i

For i = 1 To w - 1
  For j = 1 To h - 1
    s = cf(i, j)
    Picture7.PSet (i, j), RGB(s, s, s)
  Next j
Next i
End Sub

Private Sub mnumean_Click()
For i = 1 To w - 1
  For j = 1 To h - 1
  ar(0) = p(i - 1, j - 1)
  ar(1) = p(i, j - 1)
  ar(2) = p(i + 1, j - 1)
  ar(3) = p(i - 1, j)
  ar(4) = p(i, j)
  ar(5) = p(i + 1, j)
  ar(6) = p(i - 1, j + 1)
  ar(7) = p(i, j + 1)
  ar(8) = p(i + 1, j + 1)
  
  Sum = 0
  For k = 0 To 8
   Sum = Sum + ar(k)
  Next k
    
  Sum = Fix(Sum / 9)
    
  cf(i, j) = Sum
  Next j
Next i

For i = 1 To w - 1
  For j = 1 To h - 1
    s = cf(i, j)
    Picture11.PSet (i, j), RGB(s, s, s)
  Next j
Next i
End Sub

Private Sub mnumedain_Click()
For i = 1 To w - 1
  For j = 1 To h - 1
  ar(0) = p(i - 1, j - 1)
  ar(1) = p(i, j - 1)
  ar(2) = p(i + 1, j - 1)
  ar(3) = p(i - 1, j)
  ar(4) = p(i, j)
  ar(5) = p(i + 1, j)
  ar(6) = p(i - 1, j + 1)
  ar(7) = p(i, j + 1)
  ar(8) = p(i + 1, j + 1)
  
  For k = 0 To 8
    For u = 0 To 8
      If (ar(k) < ar(u)) Then
        temp = ar(k)
        ar(k) = ar(u)
        ar(u) = temp
      End If
    Next u
  Next k
   
  cf(i, j) = ar(4)
   
  Next j
Next i

For i = 1 To w
  For j = 1 To h
    s = cf(i, j)
    Picture12.PSet (i, j), RGB(s, s, s)
  Next j
Next i

End Sub

Private Sub mnuor_Click()
yy = InputBox("enter reduction value", "REDUCTION")
If (yy = "") Then
MsgBox "Where is the Reduction Value?", vbCritical, "ERROR!"
Else
x = Val(yy)

For i = 0 To 7
  If (x = 2 ^ i) Then
    n = i
  End If
Next i

m = (2 ^ (8 - n)) - 1

ii = 7
While (ii >= 0)
 If (m >= 2 ^ ii) Then
   y = m - (2 ^ ii)
   S1 = S1 + "1"
   m = y
  Else
   S1 = S1 + "0"
 End If

   ii = ii - 1
Wend

  For k = 1 To 8
   ss1(k - 1) = Mid(S1, k, 1)
  Next k


shift = 8 - n

For i = 1 To w
  For j = 1 To h
    d = p(i, j)
    ii = 7
    S2 = ""
    While (ii >= 0)
     If (d >= 2 ^ ii) Then
       y = d - (2 ^ ii)
       S2 = S2 + "1"
       d = y
     Else
       S2 = S2 + "0"
     End If
     ii = ii - 1
    Wend
  
  For k = 1 To 8
    ss2(k - 1) = Mid(S2, k, 1)
  Next k
  
  For k = 0 To 7
    If (ss1(k) = "1" Or ss2(k) = "1") Then
      ss3(k) = "1"
    Else
      ss3(k) = "0"
    End If
  Next k
  qp = 0
  k = 0
  le = 7 - shift
  While (le >= 0)
      qp = qp + ss3(k) * (2 ^ le)
      k = k + 1
      le = le - 1
  Wend
   
  cf(i, j) = Fix((qp * 256) / yy)
   
  Next j
Next i


For i = 1 To w
  For j = 1 To h
    ws = cf(i, j)
    Picture15.PSet (i, j), RGB(ws, ws, ws)
    
  Next j
Next i
End If
End Sub

Private Sub mnurow_Click()
tt = 1
For i = 1 To w
  For j = 1 To h
    p2(i, tt) = p(i, j)
    p2(i, tt + 1) = p(i, j)
     tt = tt + 2
    Next j
    tt = 1
Next i

For i = 1 To w
  For j = 1 To h * 2
    s = p2(i, j)
    Picture3.PSet (i, j), RGB(s, s, s)
  Next j
Next i
End Sub

Private Sub mnusaverage_Click()
Picture16.Picture = LoadPicture("")
zz = InputBox("enter picture width value", "WIDTH")
ff = InputBox("enter picture hight value", "HIGHT")
If (zz = "" Or ff = "") Then
MsgBox "Where is the picture Value?", vbCritical, "ERROR!"
Else
  bl = Val(zz)
  bl2 = Val(ff)
  z = (256 / bl)
  F = (256 / bl2)
  
  i = 1
  r = 1
  
  While (i < w)
    j = 1
    s = 1
    While (j < h)
      Sum = 0
      
      For k = i To i + (z - 1)
        For q = j To j + (F - 1)
         Sum = Sum + p(k, q)
        Next q
      Next k
      
      av = Fix(Sum / (z * F))
      cf(r, s) = av
      j = j + F
      s = s + 1
    Wend
    i = i + z
    r = r + 1
  Wend
  rr = r
  ss = s

  For r = 1 To rr
    For s = 1 To ss
     k = cf(r, s)
     Picture16.PSet (r, s), RGB(k, k, k)
    Next s
  Next r
  
End If
zz = 0
ff = 0
End Sub

Private Sub mnusmedain_Click()
Picture17.Picture = LoadPicture("")
zz = InputBox("enter picture width value", "WIDTH")
ff = InputBox("enter picture hight value", "HIGHT")
If (zz = "" Or ff = "") Then
MsgBox "Where is the picture Value?", vbCritical, "ERROR!"
Else
  bl = Val(zz)
  bl2 = Val(ff)
  z = (256 / bl)
  F = (256 / bl2)
  
  i = 1
  r = 1
  
  While (i < w)
    j = 1
    s = 1
    While (j < h)
      mm = 0
      For k = i To i + (z - 1)
        For q = j To j + (F - 1)
         ar3(mm) = p(k, q)
         mm = mm + 1
        Next q
      Next k
   
     For kk = 0 To (z - 1)
       For uu = 0 To (F - 1)
         If (ar3(kk) < ar3(uu)) Then
          temp = ar3(kk)
          ar3(kk) = ar3(uu)
          ar3(uu) = temp
         End If
       Next uu
     Next kk
   
        
       cf(r, s) = Fix(ar3(z / 2))
       
      j = j + F
      s = s + 1
    Wend
    i = i + z
    r = r + 1
  Wend
  rr = r
  ss = s

  For r = 1 To rr
    For s = 1 To ss
     k = cf(r, s)
     Picture17.PSet (r, s), RGB(k, k, k)
    Next s
  Next r
  
End If

End Sub

Private Sub mnuthresholding_Click()
For i = 1 To w
  For j = 1 To h
    If (p(i, j) > 127) Then
      q = 254
    Else
      q = 0
    End If
    
   cf(i, j) = q
    
  Next j
Next i

For i = 1 To w
  For j = 1 To h
    s = cf(i, j)
    Picture13.PSet (i, j), RGB(s, s, s)
  Next j
Next i
End Sub

Private Sub mnuvertical_Click()
mask(0) = 0
mask(1) = 1
mask(2) = 0
mask(3) = 0
mask(4) = 1
mask(5) = 0
mask(6) = 0
mask(7) = -1
mask(8) = 0

For i = 1 To w - 1
  For j = 1 To h - 1
  
  ar(0) = p(i - 1, j - 1)
  ar(1) = p(i, j - 1)
  ar(2) = p(i + 1, j - 1)
  ar(3) = p(i - 1, j)
  ar(4) = p(i, j)
  ar(5) = p(i + 1, j)
  ar(6) = p(i - 1, j + 1)
  ar(7) = p(i, j + 1)
  ar(8) = p(i + 1, j + 1)
  
  Sum = 0
  For k = 0 To 8
   Sum = Sum + (ar(k) * mask(k))
  Next k
    
   If (Sum < 0) Then Sum = 0
  cf(i, j) = Sum
  Next j
Next i

For i = 1 To w - 1
  For j = 1 To h - 1
    s = cf(i, j)
    Picture7.PSet (i, j), RGB(s, s, s)
  Next j
Next i
End Sub

Private Sub mnuzorder_Click()

For i = 0 To w - 1
  For j = 0 To h - 1
   ar2(0) = p(i, j)
   ar2(1) = p(i + 1, j)
   ar2(2) = p(i, j + 1)
   ar2(3) = p(i + 1, j + 1)
   Sum = 0
   
   For k = 0 To 3
    Sum = Sum + ar2(k)
   Next k
   cf(i, j) = Sum
  Next j
Next i

For i = 1 To w
  For j = 1 To h
    s = cf(i, j)
    Picture5.PSet (i, j), RGB(s, s, s)
  Next j
Next i

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

Private Sub mulmnu_Click()
x = InputBox("Enter Factor Value")

For i = 1 To w
  For j = 1 To h
    pi3(i, j) = p(i, j) * x
    If pi3(i, j) > 255 Then
       pi3(i, j) = 255
    End If
  
  Next j
Next i

For i = 1 To w
  For j = 1 To h
    hh = pi3(i, j)
    Picture9.PSet (i, j), RGB(hh, hh, hh)
  Next j
Next i
End Sub

Private Sub munsecondd_Click()
mask(0) = 0
mask(1) = 0
mask(2) = 1
mask(3) = 0
mask(4) = 1
mask(5) = 0
mask(6) = -1
mask(7) = 0
mask(8) = 0

For i = 1 To w - 1
  For j = 1 To h - 1
  
  ar(0) = p(i - 1, j - 1)
  ar(1) = p(i, j - 1)
  ar(2) = p(i + 1, j - 1)
  ar(3) = p(i - 1, j)
  ar(4) = p(i, j)
  ar(5) = p(i + 1, j)
  ar(6) = p(i - 1, j + 1)
  ar(7) = p(i, j + 1)
  ar(8) = p(i + 1, j + 1)
  
  Sum = 0
  For k = 0 To 8
   Sum = Sum + (ar(k) * mask(k))
  Next k
    
   If (Sum < 0) Then Sum = 0
  cf(i, j) = Sum
  Next j
Next i

For i = 1 To w - 1
  For j = 1 To h - 1
    s = cf(i, j)
    Picture7.PSet (i, j), RGB(s, s, s)
  Next j
Next i
End Sub

Private Sub submnuq_Click()
For i = 1 To w
  For j = 1 To h
    pi3(i, j) = p(i, j) - pi2(i, j)
    If pi3(i, j) < 0 Then
       pi3(i, j) = 0
    End If
  
  Next j
Next i

For i = 1 To w
  For j = 1 To h
    hh = pi3(i, j)
    Picture24.PSet (i, j), RGB(hh, hh, hh)
  Next j
Next i
End Sub
