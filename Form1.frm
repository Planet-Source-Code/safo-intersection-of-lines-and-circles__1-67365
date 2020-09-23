VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Intersection"
   ClientHeight    =   5835
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6315
   LinkTopic       =   "Form1"
   ScaleHeight     =   5835
   ScaleWidth      =   6315
   StartUpPosition =   2  'CenterScreen
   Begin VB.CheckBox Check4 
      Caption         =   "Show intersection of lines and circles"
      Height          =   315
      Left            =   345
      TabIndex        =   10
      Top             =   780
      Value           =   1  'Checked
      Width           =   3705
   End
   Begin VB.CheckBox Check3 
      Caption         =   "Show intersection of circles"
      Height          =   315
      Left            =   345
      TabIndex        =   9
      Top             =   450
      Value           =   1  'Checked
      Width           =   3705
   End
   Begin VB.CheckBox Check2 
      Caption         =   "Show intersection of lines"
      Height          =   315
      Left            =   345
      TabIndex        =   8
      Top             =   120
      Value           =   1  'Checked
      Width           =   3705
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Show axes"
      Height          =   315
      Left            =   4200
      TabIndex        =   7
      Top             =   120
      Width           =   1155
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   240
      Top             =   4440
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   4470
      Left            =   345
      ScaleHeight     =   4440
      ScaleWidth      =   5535
      TabIndex        =   0
      Top             =   1185
      Width           =   5565
      Begin VB.PictureBox Point6_Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   3300
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   6
         Top             =   2880
         Width           =   105
      End
      Begin VB.PictureBox Point5_Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   2625
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   5
         Top             =   2040
         Width           =   105
      End
      Begin VB.PictureBox Point4_Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   3945
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   4
         Top             =   4020
         Width           =   105
      End
      Begin VB.PictureBox Point3_Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   3420
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   3
         Top             =   210
         Width           =   105
      End
      Begin VB.PictureBox Point2_Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   4485
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   2
         Top             =   1830
         Width           =   105
      End
      Begin VB.PictureBox Point1_Pic 
         Appearance      =   0  'Flat
         BackColor       =   &H8000000A&
         ForeColor       =   &H80000008&
         Height          =   105
         Left            =   1620
         ScaleHeight     =   75
         ScaleWidth      =   75
         TabIndex        =   1
         Top             =   1050
         Width           =   105
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Coded by : Safo
' Email : safo@zoznam.sk
' Program : Intersection of lines and circles

' This code can find intersection of line and circle, line and line, circle and circle.
' It includes some useful math functions. The code is easy and commented.

Private Type Point_Type
    x As Single
    y As Single
End Type

Private Type Param_Type ' parametric line
    xn As Single
    xt As Single
    yn As Single
    yt As Single
End Type

Private Type Circle_Type ' circle type
    x As Single
    y As Single
    r As Single
End Type

Dim Show_Axes As Boolean
Dim Show_Intersection_C As Boolean
Dim Show_Intersection_L As Boolean
Dim Show_Intersection_CL As Boolean

Private Sub Form_Load()
    Show_Axes = False
    Show_Intersection_L = True
    Show_Intersection_C = True
    Show_Intersection_CL = True
End Sub

Private Sub Check1_Click()
    Show_Axes = Check1.Value
End Sub

Private Sub Check2_Click()
    Show_Intersection_L = Check2.Value
End Sub

Private Sub Check3_Click()
    Show_Intersection_C = Check3.Value
End Sub

Private Sub Check4_Click()
    Show_Intersection_CL = Check4.Value
End Sub

Private Sub Point1_Pic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Movement Point1_Pic, Button
    End If
End Sub

Private Sub Point2_Pic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Movement Point2_Pic, Button
    End If
End Sub

Private Sub Point3_Pic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Movement Point3_Pic, Button
    End If
End Sub

Private Sub Point4_Pic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Movement Point4_Pic, Button
    End If
End Sub

Private Sub Point5_Pic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Movement Point5_Pic, Button
    End If
End Sub

Private Sub Point6_Pic_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 1 Then
        Movement Point6_Pic, Button
    End If
End Sub

Private Function Main()

Dim a As Point_Type
Dim b As Point_Type
Dim C As Point_Type
Dim d As Point_Type
Dim p As Param_Type
Dim q As Param_Type
Dim k As Circle_Type
Dim l As Circle_Type
Dim n As Integer  ' number of points of intersection

Picture1.Cls

a.x = Point1_Pic.Left / 15 + 3
a.y = Point1_Pic.Top / 15 + 3
b.x = Point2_Pic.Left / 15 + 3
b.y = Point2_Pic.Top / 15 + 3
C.x = Point3_Pic.Left / 15 + 3
C.y = Point3_Pic.Top / 15 + 3
d.x = Point4_Pic.Left / 15 + 3
d.y = Point4_Pic.Top / 15 + 3
p.xn = a.x  ' parametric line p
p.xt = b.x - a.x
p.yn = a.y
p.yt = b.y - a.y
q.xn = C.x  ' parametric line q
q.xt = d.x - C.x
q.yn = C.y
q.yt = d.y - C.y
k.x = Point5_Pic.Left / 15 + 3 ' circle k
k.y = Point5_Pic.Top / 15 + 3
k.r = 30
l.x = Point6_Pic.Left / 15 + 3 ' circle l
l.y = Point6_Pic.Top / 15 + 3
l.r = 50

' Draw x axis and y axis
If Show_Axes Then
    Picture1.Line (Picture1.Width / 2, 0)-(Picture1.Width / 2, Picture1.Height), RGB(200, 200, 200)
    Picture1.Line (0, Picture1.Height / 2)-(Picture1.Width, Picture1.Height / 2), RGB(200, 200, 200)
End If

If Show_Intersection_L Then

    ' intersection of line p and line q
    If Line_Line(p, q, desx, desy) Then
        Picture1.Circle (desx * 15, desy * 15), 75, vbGreen
    End If

End If

If Show_Intersection_CL Then
   
    ' intersection of line p and circle k
    n = Line_Circle(p, k, desx1, desy1, desx2, desy2)
    
    If n >= 1 Then Picture1.Circle (desx1 * 15, desy1 * 15), 75, vbRed
    If n = 2 Then Picture1.Circle (desx2 * 15, desy2 * 15), 75, vbRed
    desx1 = 0: desy1 = 0: desx2 = 0: desy2 = 0

    ' intersection of line q and circle k
    n = Line_Circle(q, k, desx1, desy1, desx2, desy2)
        
    If n >= 1 Then Picture1.Circle (desx1 * 15, desy1 * 15), 75, vbRed
    If n = 2 Then Picture1.Circle (desx2 * 15, desy2 * 15), 75, vbRed
    desx1 = 0: desy1 = 0: desx2 = 0: desy2 = 0
    
    ' intersection of line q and circle l
    n = Line_Circle(p, l, desx1, desy1, desx2, desy2)
        
    If n >= 1 Then Picture1.Circle (desx1 * 15, desy1 * 15), 75, vbBlue
    If n = 2 Then Picture1.Circle (desx2 * 15, desy2 * 15), 75, vbBlue
    desx1 = 0: desy1 = 0: desx2 = 0: desy2 = 0
    
    ' intersection of line q and circle l
    n = Line_Circle(q, l, desx1, desy1, desx2, desy2)
        
    If n >= 1 Then Picture1.Circle (desx1 * 15, desy1 * 15), 75, vbBlue
    If n = 2 Then Picture1.Circle (desx2 * 15, desy2 * 15), 75, vbBlue
    desx1 = 0: desy1 = 0: desx2 = 0: desy2 = 0

End If

If Show_Intersection_C Then

    ' intersection of circle k and circle l
    n = Circle_Circle(k, l, desx1, desy1, desx2, desy2)
    If n >= 1 Then Picture1.Circle (desx1 * 15, desy1 * 15), 75, vbMagenta
    If n = 2 Then Picture1.Circle (desx2 * 15, desy2 * 15), 75, vbMagenta

End If

' Draw line p and line q
Picture1.Line (a.x * 15, a.y * 15)-(b.x * 15, b.y * 15), vbBlack
Picture1.Line (C.x * 15, C.y * 15)-(d.x * 15, d.y * 15), vbBlack

' Draw circle k and circle l
Picture1.Circle (k.x * 15, k.y * 15), k.r * 15, vbBlack
Picture1.Circle (l.x * 15, l.y * 15), l.r * 15, vbBlack

End Function

Private Function Line_Circle(p As Param_Type, k As Circle_Type, desx1, desy1, desx2, desy2) As Integer
    Dim a As Single
    Dim x1 As Single
    Dim x2 As Single
    Dim y1 As Single
    Dim y2 As Single
    Dim point1_exist As Boolean
    Dim point2_exist As Boolean
    Dim points As Integer ' points of intersection

    x1 = p.xn - k.x
    x2 = (p.xn + p.xt) - k.x
    y1 = p.yn - k.y
    y2 = (p.yn + p.yt) - k.y
    
    dx = x2 - x1
    dy = y2 - y1
    dr = Sqr(dx ^ 2 + dy ^ 2)
    d = x1 * y2 - x2 * y1
    
    If (k.r ^ 2 * dr ^ 2 - d ^ 2) >= 0 Then intersection = True
    
    If intersection = True Then
    
        If (k.r ^ 2 * dr ^ 2 - d ^ 2) < 0 Then
            a = 0
        Else
            a = Sqr(k.r ^ 2 * dr ^ 2 - d ^ 2)
        End If
        
        desx1 = (d * dy + My_Sgn(dy) * dx * a) / dr ^ 2 + k.x
        desy1 = (-d * dx + Abs(dy) * a) / dr ^ 2 + k.y
        desx2 = (d * dy - My_Sgn(dy) * dx * a) / dr ^ 2 + k.x
        desy2 = (-d * dx - Abs(dy) * a) / dr ^ 2 + k.y
        
    End If
    
    point1_exist = Point_Line(desx1, desy1, p)
    point2_exist = Point_Line(desx2, desy2, p)
    
    If point1_exist And point2_exist Then
        points = 2
    Else
        points = 0
        If point1_exist Then points = 1
        If point2_exist Then points = 1: desx1 = desx2: desy1 = desy2
    End If
    
    Line_Circle = points

End Function

Private Function Circle_Circle(k As Circle_Type, l As Circle_Type, desx1, desy1, desx2, desy2) As Integer
    Dim d As Single
    Dim x As Single
    Dim y As Single
    Dim points As Integer
    
    ' R = k.r
    ' r = l.r
    
    dx = k.x - l.x
    dy = k.y - l.y
    
    d = Sqr(dx ^ 2 + dy ^ 2)
    
    If d <> 0 Then
    
        x = (d ^ 2 - l.r ^ 2 + k.r ^ 2) / (2 * d)
        If (k.r ^ 2 - x ^ 2) > 0 Then y = Sqr(k.r ^ 2 - x ^ 2)
    
        desx1 = k.x - (x * (dx / d) + y * (dy / d))
        desy1 = k.y - (-y * (dx / d) + x * (dy / d))
        desx2 = k.x - (x * (dx / d) - y * (dy / d))
        desy2 = k.y - (y * (dx / d) + x * (dy / d))
    
    End If
    
    If d < (k.r + l.r) And d > Abs(k.r - l.r) Then points = 2
    If d = (k.r + l.r) Then points = 1
    
    Circle_Circle = points

End Function

Private Function Line_Line(a As Param_Type, b As Param_Type, desx, desy) As Boolean
    Dim a1 As Single
    Dim a2 As Single
    Dim b1 As Single
    Dim b2 As Single
    Dim c1 As Single
    Dim c2 As Single
    Dim f As Boolean

    a1 = a.xt
    a2 = a.yt
    b1 = -b.xt
    b2 = -b.yt
    c1 = a.xn - b.xn
    c2 = a.yn - b.yn
    t = 0
    s = 0
    
    If b1 = 0 Then
    
        If a1 <> 0 Then t = -c1 / a1
    
    Else
        
        t = ((b2 * c1) / b1 - c2) / (a2 - (b2 * a1) / b1)
    
    End If
        
        desx = a.xn + a.xt * t
        desy = a.yn + a.yt * t
            
    If Point_Line(desx, desy, a) And Point_Line(desx, desy, b) Then f = True
    
    Line_Line = f

End Function

Private Function Point_Line(x, y, p As Param_Type) As Boolean
    Dim t1 As Single
    Dim t2 As Single
    Dim op As Boolean
    
    If p.xt = 0 Then
        t = (y - p.yn) / p.yt
        If t <= 1 And t >= 0 And x = p.xn Then op = True
    End If
    
    If p.yt = 0 Then
        t = (x - p.xn) / p.xt
        If t <= 1 And t >= 0 And y = p.yn Then op = True
    End If
    
    If p.xt <> 0 And p.yt <> 0 Then
        t1 = (x - p.xn) / p.xt
        t2 = (y - p.yn) / p.yt
        If Abs(t1 - t2) <= 0.001 And t1 <= 1 And t1 >= 0 Then op = True
    End If
    
    Point_Line = op
    
End Function

Private Function My_Sgn(x) As Integer

    If x < 0 Then
        My_Sgn = -1
    Else
        My_Sgn = 1
    End If
    
End Function

Private Sub Timer1_Timer()

Main

End Sub
