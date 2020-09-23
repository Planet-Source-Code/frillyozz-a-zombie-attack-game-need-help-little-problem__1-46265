VERSION 5.00
Begin VB.Form frmLevel1 
   BackColor       =   &H00008000&
   BorderStyle     =   0  'None
   Caption         =   "Z o m b i E   A t t a c K"
   ClientHeight    =   5865
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   8805
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MousePointer    =   2  'Cross
   ScaleHeight     =   391
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   587
   StartUpPosition =   2  'CenterScreen
   Begin VB.Timer Timer3 
      Left            =   1080
      Top             =   120
   End
   Begin VB.PictureBox Picture15 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   1320
      Picture         =   "Form1.frx":0000
      ScaleHeight     =   720
      ScaleWidth      =   390
      TabIndex        =   16
      Top             =   4080
      Width           =   390
   End
   Begin VB.PictureBox Picture12 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   1320
      Picture         =   "Form1.frx":0F42
      ScaleHeight     =   720
      ScaleWidth      =   390
      TabIndex        =   10
      Top             =   3240
      Width           =   390
   End
   Begin VB.PictureBox Picture9 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   1320
      Picture         =   "Form1.frx":1E84
      ScaleHeight     =   720
      ScaleWidth      =   390
      TabIndex        =   7
      Top             =   2400
      Width           =   390
   End
   Begin VB.PictureBox Picture6 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   1320
      Picture         =   "Form1.frx":2DC6
      ScaleHeight     =   720
      ScaleWidth      =   390
      TabIndex        =   4
      Top             =   1560
      Width           =   390
   End
   Begin VB.PictureBox Picture3 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   1320
      Picture         =   "Form1.frx":3D08
      ScaleHeight     =   720
      ScaleWidth      =   390
      TabIndex        =   3
      Top             =   720
      Width           =   390
   End
   Begin VB.Timer Timer2 
      Interval        =   2000
      Left            =   600
      Top             =   120
   End
   Begin VB.PictureBox Picture14 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   720
      Picture         =   "Form1.frx":4C4A
      ScaleHeight     =   720
      ScaleWidth      =   390
      TabIndex        =   15
      Top             =   4080
      Width           =   390
   End
   Begin VB.PictureBox Picture13 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   120
      Picture         =   "Form1.frx":5B8C
      ScaleHeight     =   720
      ScaleWidth      =   405
      TabIndex        =   14
      Top             =   4080
      Width           =   405
   End
   Begin VB.PictureBox Picture10 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   120
      Picture         =   "Form1.frx":6B8E
      ScaleHeight     =   720
      ScaleWidth      =   405
      TabIndex        =   12
      Top             =   3240
      Width           =   405
   End
   Begin VB.PictureBox Picture11 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   720
      Picture         =   "Form1.frx":7B90
      ScaleHeight     =   720
      ScaleWidth      =   390
      TabIndex        =   11
      Top             =   3240
      Width           =   390
   End
   Begin VB.PictureBox Picture7 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   120
      Picture         =   "Form1.frx":8AD2
      ScaleHeight     =   720
      ScaleWidth      =   405
      TabIndex        =   9
      Top             =   2400
      Width           =   405
   End
   Begin VB.PictureBox Picture8 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   720
      Picture         =   "Form1.frx":9AD4
      ScaleHeight     =   720
      ScaleWidth      =   390
      TabIndex        =   8
      Top             =   2400
      Width           =   390
   End
   Begin VB.PictureBox Picture4 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   120
      Picture         =   "Form1.frx":AA16
      ScaleHeight     =   720
      ScaleWidth      =   405
      TabIndex        =   6
      Top             =   1560
      Width           =   405
   End
   Begin VB.PictureBox Picture5 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   720
      Picture         =   "Form1.frx":BA18
      ScaleHeight     =   720
      ScaleWidth      =   390
      TabIndex        =   5
      Top             =   1560
      Width           =   390
   End
   Begin VB.Timer Timer1 
      Left            =   120
      Top             =   120
   End
   Begin VB.PictureBox Picture2 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   720
      Picture         =   "Form1.frx":C95A
      ScaleHeight     =   720
      ScaleWidth      =   390
      TabIndex        =   1
      Top             =   720
      Width           =   390
   End
   Begin VB.PictureBox Picture1 
      AutoSize        =   -1  'True
      BackColor       =   &H00008000&
      BorderStyle     =   0  'None
      Height          =   720
      Left            =   120
      Picture         =   "Form1.frx":D89C
      ScaleHeight     =   720
      ScaleWidth      =   405
      TabIndex        =   0
      Top             =   720
      Width           =   405
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Game Over"
      BeginProperty Font 
         Name            =   "Secret Service Typewriter"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   547
      TabIndex        =   18
      Top             =   1920
      Visible         =   0   'False
      Width           =   7710
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Level 1"
      BeginProperty Font 
         Name            =   "Secret Service Typewriter"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1575
      Left            =   1395
      TabIndex        =   17
      Top             =   1920
      Width           =   6000
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   240
      Left            =   1200
      TabIndex        =   13
      Top             =   5520
      Width           =   105
   End
   Begin VB.Label labBullets 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "BULLETS"
      BeginProperty Font 
         Name            =   "Secret Service Typewriter"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   5520
      Width           =   960
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H000000C0&
      FillColor       =   &H000000C0&
      FillStyle       =   0  'Solid
      Height          =   495
      Left            =   0
      Top             =   5400
      Width           =   8895
   End
End
Attribute VB_Name = "frmLevel1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim z1 As Integer
Dim zr1 As Integer
Dim z2 As Integer
Dim zr2 As Integer
Dim z3 As Integer
Dim zr3 As Integer
Dim z4 As Integer
Dim zr4 As Integer
Dim z5 As Integer
Dim zr5 As Integer
Dim bullets As Integer
Dim b As Integer
Dim speed As Double
Dim done3 As Integer
Dim done1 As Integer
Dim done2 As Integer
Dim done4 As Integer
Dim done5 As Integer


Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyR Then
        bullets = 6
        Label1.Caption = 6
        sndPlaySound App.Path + "\sfx\gunreload1.WAV", SND_NODEFAULT + SND_ASYNC
        labBullets.ForeColor = vbBlack
    End If

End Sub

Private Sub Form_Load()
    
    z1 = 1
    z2 = 1
    z3 = 1
    z4 = 1
    z5 = 1
    b = 1
    bullets = 6
    
    done1 = 0
    done2 = 0
    done3 = 0
    done4 = 0
    done5 = 0
    
    Picture1.Top = -50
    Picture4.Top = -60
    Picture7.Top = -200
    Picture10.Top = -100
    Picture13.Top = -200
    speed = 5
    
    Picture1.Left = 500
    Picture4.Left = ((Rnd * 200) + 100)
    Picture7.Left = ((Rnd * 300) + 170)
    Picture10.Left = Minute(Time) * ((Rnd * 5) + 1)
    Picture13.Left = ((Rnd * 350) + 150)
    
    
    Picture2.Visible = False
    Picture3.Visible = False
    
    Picture5.Visible = False
    Picture6.Visible = False
    
    Picture8.Visible = False
    Picture9.Visible = False
    
    Picture11.Visible = False
    Picture12.Visible = False
    
    Picture14.Visible = False
    Picture15.Visible = False

End Sub


Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bullets > 0 Then
    sndPlaySound App.Path + "\sfx\gunshot1.WAV", SND_NODEFAULT + SND_ASYNC
    bullets = bullets - 1
    Label1.Caption = Val(bullets)
    End If

End Sub


Private Sub Picture1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bullets > 0 Then
        Picture3.Visible = True
        Timer3.Interval = 400
        z1 = 3
        sndPlaySound App.Path + "\sfx\gunshot1.WAV", SND_NODEFAULT + SND_ASYNC
        bullets = bullets - 1
        Label1.Caption = Val(bullets)
        zr1 = zr1 + 1
        
        If zr1 < 8 Then
            Picture1.Top = -100
            Picture1.Left = ((220 * Rnd) + 150)
            Picture1.Visible = True
            z1 = 1
        End If
        
        If zr1 >= 8 Then
            done1 = 1
        End If
    
    End If

End Sub


Private Sub Picture2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bullets > 0 Then
        Picture3.Visible = True
        Timer3.Interval = 400
        z1 = 3
        sndPlaySound App.Path + "\sfx\gunshot1.WAV", SND_NODEFAULT + SND_ASYNC
        bullets = bullets - 1
        Label1.Caption = Val(bullets)
        zr1 = zr1 + 1
        
        If zr1 < 8 Then
            Picture1.Top = -100
            Picture1.Left = ((Rnd * 300) + 50)
            Picture1.Visible = True
            z1 = 1
            
        If zr1 >= 8 Then
            done1 = 1
        End If

        End If

    End If

End Sub


Private Sub Picture10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bullets > 0 Then
        Picture12.Visible = True
        Timer3.Interval = 400
        z4 = 3
        sndPlaySound App.Path + "\sfx\gunshot1.WAV", SND_NODEFAULT + SND_ASYNC
        bullets = bullets - 1
        Label1.Caption = Val(bullets)
        zr4 = zr4 + 1
        
        If zr4 < 8 Then
            Picture10.Top = -100
            Picture10.Left = ((Rnd * 300) + 1)
            Picture10.Visible = True
            z4 = 1
        End If
        
        If zr4 >= 8 Then
            done4 = 1
        End If
        
    End If

End Sub


Private Sub Picture11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bullets > 0 Then
        Picture12.Visible = True
        Timer3.Interval = 400
        z4 = 3
        sndPlaySound App.Path + "\sfx\gunshot1.WAV", SND_NODEFAULT + SND_ASYNC
        bullets = bullets - 1
        Label1.Caption = Val(bullets)
        zr4 = zr4 + 1
        
        If zr4 < 8 Then
            Picture10.Top = -50
            Picture10.Left = Minute(Time)
            Picture10.Visible = True
            z4 = 1
        End If
        
        If zr4 >= 8 Then
            done4 = 1
        End If
        
    End If

End Sub


Private Sub Picture13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bullets > 0 Then
        Picture15.Visible = True
        Timer3.Interval = 400
        z5 = 3
        sndPlaySound App.Path + "\sfx\gunshot1.WAV", SND_NODEFAULT + SND_ASYNC
        bullets = bullets - 1
        Label1.Caption = Val(bullets)
        zr5 = zr5 + 1
        
        If zr5 < 8 Then
            Picture13.Top = -100
            Picture13.Left = ((Rnd * 250) + 100)
            Picture13.Visible = True
            z5 = 1
        End If
        
        If zr5 >= 8 Then
            done5 = 1
        End If
        
    End If

End Sub


Private Sub Picture14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bullets > 0 Then
        Picture15.Visible = True
        Timer3.Interval = 400
        z5 = 3
        sndPlaySound App.Path + "\sfx\gunshot1.WAV", SND_NODEFAULT + SND_ASYNC
        bullets = bullets - 1
        Label1.Caption = Val(bullets)
        zr5 = zr5 + 1
        
        If zr5 < 8 Then
            Picture13.Top = -100
            Picture13.Left = ((Rnd * 300) + 75)
            Picture13.Visible = True
            z5 = 1
        End If
        
        If zr5 >= 8 Then
            done5 = 1
        End If
        
    End If

End Sub



Private Sub Picture4_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bullets > 0 Then
        Picture6.Visible = True
        Timer3.Interval = 400
        z2 = 3
        sndPlaySound App.Path + "\sfx\gunshot1.WAV", SND_NODEFAULT + SND_ASYNC
        bullets = bullets - 1
        Label1.Caption = Val(bullets)
        zr2 = zr2 + 1
        
        If zr2 < 8 Then
            Picture4.Top = -100
            Picture4.Left = ((Rnd * 300) + 50)
            Picture4.Visible = True
            z2 = 1
        End If
        
        If zr2 >= 8 Then
            done2 = 1
        End If
        
    End If

End Sub

Private Sub Picture5_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    
    If bullets > 0 Then
        Picture6.Visible = True
        Timer3.Interval = 400
        z2 = 3
        sndPlaySound App.Path + "\sfx\gunshot1.WAV", SND_NODEFAULT + SND_ASYNC
        bullets = bullets - 1
        Label1.Caption = Val(bullets)
        zr2 = zr2 + 1
        
        If zr2 < 8 Then
            Picture4.Top = -100
            Picture4.Left = ((Rnd * 140) + 50)
            Picture4.Visible = True
            z2 = 1
        End If
        
        If zr2 >= 8 Then
            done2 = 1
        End If
        
    End If
    
End Sub

Private Sub Picture7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bullets > 0 Then
        Picture9.Visible = True
        Timer3.Interval = 400
        z3 = 3
        sndPlaySound App.Path + "\sfx\gunshot1.WAV", SND_NODEFAULT + SND_ASYNC
        bullets = bullets - 1
        Label1.Caption = Val(bullets)
        zr3 = zr3 + 1
        
        If zr3 < 8 Then
            Picture7.Top = -100
            Picture7.Left = ((Rnd) + 100)
            Picture7.Visible = True
            z3 = 1
        End If
        
        If zr3 >= 8 Then
            done3 = 1
        End If
        
    End If

End Sub

Private Sub Picture8_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If bullets > 0 Then
        Picture9.Visible = True
        Timer3.Interval = 400
        z3 = 3
        sndPlaySound App.Path + "\sfx\gunshot1.WAV", SND_NODEFAULT + SND_ASYNC
        bullets = bullets - 1
        Label1.Caption = Val(bullets)
        zr3 = zr3 + 1
        
        If zr3 < 8 Then
            Picture7.Top = -100
            Picture7.Left = ((Rnd * 270) + 60)
            Picture7.Visible = True
            z3 = 1
        End If
        
        If zr3 >= 8 Then
            done3 = 1
        End If
        
    End If

End Sub

Private Sub Timer1_Timer()


If Picture1.Top + Picture1.Height >= 360 And done1 = 0 Then
    Label3.Visible = True
End If

If Picture4.Top + Picture4.Height >= 360 And done2 = 0 Then
    Label3.Visible = True
End If

If Picture7.Top + Picture7.Height >= 360 And done3 = 0 Then
    Label3.Visible = True
End If

If Picture10.Top + Picture10.Height >= 360 And done4 = 0 Then
    Label3.Visible = True
End If

If Picture13.Top + Picture13.Height >= 360 And done5 = 0 Then
    Label3.Visible = True
End If



    If bullets = 0 Then
    labBullets.ForeColor = vbYellow
    End If

'All for zombie one
    If z1 = 3 Then
        Picture1.Visible = False
        Picture2.Visible = False
    End If

    If z1 = 1 Then
        Picture2.Visible = True
        Picture1.Visible = False
        z1 = z1 - 1
    End If
    
    
    If z1 = 2 Then
        Picture1.Visible = True
        Picture2.Visible = False
        z1 = z1 - 3
    End If
    
    z1 = z1 + 2
    
    Picture1.Top = Picture1.Top + speed
    Picture2.Top = Picture1.Top
    Picture2.Left = Picture1.Left
    Picture3.Top = Picture1.Top
    Picture3.Left = Picture1.Left
    
    
    
'All for zombie two
    If z2 = 3 Then
        Picture4.Visible = False
        Picture5.Visible = False
    End If

    If z2 = 1 Then
        Picture5.Visible = True
        Picture4.Visible = False
        z2 = z2 - 1
    End If
    
    
    If z2 = 2 Then
        Picture4.Visible = True
        Picture5.Visible = False
        z2 = z2 - 3
    End If
    
        z2 = z2 + 2
    
    Picture4.Top = Picture4.Top + speed
    Picture5.Top = Picture4.Top
    Picture5.Left = Picture4.Left
    Picture6.Top = Picture4.Top
    Picture6.Left = Picture4.Left
    
    
    
    
    'All for zombie three
    If z3 = 3 Then
        Picture7.Visible = False
        Picture8.Visible = False
    End If

    If z3 = 1 Then
        Picture8.Visible = True
        Picture7.Visible = False
        z3 = z3 - 1
    End If
    
    
    If z3 = 2 Then
        Picture7.Visible = True
        Picture8.Visible = False
        z3 = z3 - 3
    End If
    
        z3 = z3 + 2
    
    Picture7.Top = Picture7.Top + speed
    Picture8.Top = Picture7.Top
    Picture8.Left = Picture7.Left
    Picture9.Top = Picture7.Top
    Picture9.Left = Picture7.Left
    
    
    
    
        
    'All for zombie four
    If z4 = 3 Then
        Picture10.Visible = False
        Picture11.Visible = False
    End If

    If z4 = 1 Then
        Picture11.Visible = True
        Picture10.Visible = False
        z4 = z4 - 1
    End If
    
    
    If z4 = 2 Then
        Picture10.Visible = True
        Picture11.Visible = False
        z4 = z4 - 3
    End If
    
        z4 = z4 + 2
    
    Picture10.Top = Picture10.Top + speed
    Picture11.Top = Picture10.Top
    Picture11.Left = Picture10.Left
    Picture12.Top = Picture10.Top
    Picture12.Left = Picture10.Left
    
    
    
    
    'All for zombie five
    If z5 = 3 Then
        Picture13.Visible = False
        Picture14.Visible = False
    End If

    If z5 = 1 Then
        Picture14.Visible = True
        Picture13.Visible = False
        z5 = z5 - 1
    End If
    
    
    If z5 = 2 Then
        Picture13.Visible = True
        Picture14.Visible = False
        z5 = z5 - 3
    End If
    
        z5 = z5 + 2
    
    Picture13.Top = Picture13.Top + speed
    Picture14.Top = Picture13.Top
    Picture14.Left = Picture13.Left
    Picture15.Top = Picture13.Top
    Picture15.Left = Picture13.Left
    
End Sub


Private Sub Timer2_Timer()

Label2.Visible = False
Timer1.Interval = 100

    If Picture1.Visible = False And Picture4.Visible = False And Picture7.Visible = False And Picture10.Visible = False And Picture13.Visible = False Then
        If Picture2.Visible = False And Picture5.Visible = False And Picture8.Visible = False And Picture11.Visible = False And Picture13.Visible = False Then
            frmLevel2.Visible = True
            frmLevel1.Visible = False
        End If
    End If
    

End Sub

Private Sub Timer3_Timer()

    speed = speed + 0.3

    Picture3.Visible = False
    Picture6.Visible = False
    Picture9.Visible = False
    Picture12.Visible = False
    Picture15.Visible = False
    
    Timer3.Interval = 0

End Sub

