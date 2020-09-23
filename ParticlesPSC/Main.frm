VERSION 5.00
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   " Particle Effects Demo"
   ClientHeight    =   9855
   ClientLeft      =   150
   ClientTop       =   450
   ClientWidth     =   11445
   ForeColor       =   &H00000000&
   Icon            =   "Main.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   657
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   763
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox picCTRL 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H0080C0FF&
      ForeColor       =   &H80000008&
      Height          =   780
      Left            =   0
      ScaleHeight     =   50
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   750
      TabIndex        =   2
      Top             =   6780
      Width           =   11280
      Begin VB.CommandButton cmdSound 
         BackColor       =   &H0080C0FF&
         Caption         =   "Sound ON"
         Height          =   510
         Left            =   1020
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   105
         Width           =   945
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H0080C0FF&
         Caption         =   "Wiper"
         Height          =   225
         Index           =   7
         Left            =   9585
         TabIndex        =   18
         Top             =   90
         Width           =   900
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H0080C0FF&
         Caption         =   "Expand"
         Height          =   225
         Index           =   6
         Left            =   8610
         TabIndex        =   17
         Top             =   90
         Width           =   900
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H0080C0FF&
         Caption         =   "Spirals"
         Height          =   225
         Index           =   5
         Left            =   7725
         TabIndex        =   16
         Top             =   90
         Width           =   900
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H0080C0FF&
         Caption         =   "Wavy"
         Height          =   225
         Index           =   4
         Left            =   6615
         TabIndex        =   15
         Top             =   90
         Width           =   900
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H0080C0FF&
         Caption         =   "Spray"
         Height          =   225
         Index           =   3
         Left            =   9585
         TabIndex        =   14
         Top             =   360
         Width           =   900
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H0080C0FF&
         Caption         =   "Spurt"
         Height          =   225
         Index           =   2
         Left            =   8610
         TabIndex        =   13
         Top             =   360
         Width           =   900
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H0080C0FF&
         Caption         =   "Hot"
         Height          =   225
         Index           =   1
         Left            =   7740
         TabIndex        =   12
         Top             =   360
         Width           =   615
      End
      Begin VB.OptionButton optType 
         BackColor       =   &H0080C0FF&
         Caption         =   "Fountain"
         Height          =   225
         Index           =   0
         Left            =   6615
         TabIndex        =   11
         Top             =   360
         Value           =   -1  'True
         Width           =   1050
      End
      Begin VB.HScrollBar scrSpeed 
         Height          =   195
         Left            =   4440
         Max             =   90
         Min             =   10
         TabIndex        =   7
         TabStop         =   0   'False
         Top             =   390
         Value           =   10
         Width           =   1800
      End
      Begin VB.HScrollBar scrAngle 
         Height          =   195
         Left            =   2520
         Max             =   110
         Min             =   70
         TabIndex        =   5
         TabStop         =   0   'False
         Top             =   390
         Value           =   70
         Width           =   1800
      End
      Begin VB.CommandButton cmdStart 
         BackColor       =   &H0080C0FF&
         Caption         =   "Start"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   510
         Left            =   60
         Style           =   1  'Graphical
         TabIndex        =   3
         Top             =   105
         Width           =   720
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Speed"
         Height          =   195
         Index           =   1
         Left            =   4680
         TabIndex        =   10
         Top             =   150
         Width           =   645
      End
      Begin VB.Label Label2 
         BackColor       =   &H0080C0FF&
         Caption         =   "Angle"
         Height          =   195
         Index           =   0
         Left            =   2790
         TabIndex        =   9
         Top             =   150
         Width           =   510
      End
      Begin VB.Label LabSpeed 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   5460
         TabIndex        =   8
         Top             =   90
         Width           =   555
      End
      Begin VB.Label LabAngle 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   3615
         TabIndex        =   6
         Top             =   90
         Width           =   555
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         BorderStyle     =   1  'Fixed Single
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   10485
         TabIndex        =   4
         Top             =   390
         Width           =   765
      End
   End
   Begin VB.PictureBox picIN 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1395
      Left            =   11760
      ScaleHeight     =   93
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   1
      Top             =   900
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.PictureBox PIC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H00C0C0C0&
      ForeColor       =   &H80000008&
      Height          =   6750
      Left            =   0
      ScaleHeight     =   448
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   750
      TabIndex        =   0
      Top             =   0
      Width           =   11280
      Begin VB.Image imEmit 
         Height          =   330
         Left            =   5415
         Picture         =   "Main.frx":18BA
         Top             =   6360
         Width           =   420
      End
   End
   Begin VB.Menu mnuOpen 
      Caption         =   "&Open"
   End
   Begin VB.Menu mnuSave 
      Caption         =   "&Save As"
   End
   Begin VB.Menu mnuExit 
      Caption         =   "E&xit"
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' Particle Effects Demo  by  Robert Rayment
' June 2008

' Update 10 June 08

'1. Minor update Sound ON/OFF caption changed.

Option Explicit

' Loaded image size
Private picwidth As Long
Private picheight  As Long
' Display image size
Private W As Long, H As Long  ' Image width & hweight

Private FPS As Long  ' Frames Per Second
Private aDone As Boolean   ' Loop exit

' Particles
Private xp() As Long, yp() As Long  ' Pixel centre coords
Private NumParticles As Long        ' Number of particles
Private Breadth As Long             ' Random Breadth of particles
' Colors
Private CCen As Long
Private CTop As Long, CLef As Long, CRit As Long, CBot As Long

Private CenR As Byte, CenG As Byte, CenB As Byte
Private TopR As Byte, TopG As Byte, TopB As Byte
Private LefR As Byte, LefG As Byte, LefB As Byte
Private RitR As Byte, RitG As Byte, RitB As Byte
Private BotR As Byte, BotG As Byte, BotB As Byte

Private Red As Byte, Green As Byte, Blue As Byte

' Emitter
Private imx As Single, imy As Single
Private imMouseDown As Boolean
Private OldX As Long
Private OldY As Long
Private EmitType As Long

' Parabolic parameters & scrollbars
Private MaxSpeed As Single
Private Grav As Single
Private Angle As Single
Private sAngle As Single
Private aScroll As Boolean

' Timing
Private TT As Long
Private TimeElapsed As Single         ' ms
Private TimeLng As Long               ' Int ms
Private ScaledTime As Single
Private MaxScaledTime As Single
Private ST As Single
Private STDiv As Long

' File
Private PathSpec$, CurrentPath$, FileSpec$
Private SavePath$, SaveSpec$

Dim CommonDialog1 As cOSDialog

Dim tmr As CTiming   ' For cTimLNG.cls

' picDATAORG(x,y)
' picDATA(x,y)
' For picDATA(x,y) color: use RGB(Blue,Green,Red) ie Red & Blue swapped
' y = H-1
'  ^
'  |
' y = 0

Private Sub cmdStart_Click()
Dim k As Long
Dim Speed As Single  ' Varing speed (Pressure, Spread)
Dim B As Single      ' Varying breadth as (yp())
Dim S As Single      ' Variable
Dim rad As Single    ' Variable

' imEmit
Dim imx As Single
Dim imxStart As Single

   If aDone = False Then
      aDone = True
      If aSound Then StopPlay
      cmdStart.Caption = "Start"
      Exit Sub
   End If
   cmdStart.Caption = "Stop"
   
   aDone = False
   
   FPS = 0
   ' T&E
   Breadth = 25
   Grav = -9.8 / 300
   ScaledTime = 0
   Speed = 0
   
   If EmitType = 4 Then ' ie Wavy start in middle
      imEmit.Left = W / 2
   End If
   
   ' Sound 101,102,,,108
   Play CInt(EmitType)  ' = 0,1,2,3,4,5,6,7
   
   imxStart = imEmit.Left
   Do
      imx = imEmit.Left + imEmit.Width \ 2
      tmr.Reset
      ' Reset to background image
      picDATA() = picDATAORG()
      
      For k = 0 To NumParticles - 1
         Angle = sAngle * d2r# + (Rnd - 0.5) / 10
         ' EmitType = 0 Fountain, 1 Hot, 2 Spurt, 3 Spray
         '            4 Wavy, 5 Spirals, 6 Expand, 7 Wiper
         ' STDiv = 15,45,45,45,45,45,45,45 for EmitType 0,1,2,3,4,5,6,7
         ST = ScaledTime + (k / STDiv)
         
         Select Case EmitType
         Case 0, 1, 2   ' Fountain,Hot,Spurt
            ' Parabolic
            yp(k) = 0.5 * Grav * (ST * ST) + Speed * Sin(Angle) * ST
            xp(k) = MaxSpeed * Cos(Angle) * ST
            xp(k) = xp(k) + imx
            B = yp(k) * Breadth / (H - 1)
         Case 3   ' Spray
            rad = k * NumParticles / H
            yp(k) = rad * Sin(Angle) / (ST + 60)
            If Speed > 1 Then
               B = yp(k) * Breadth / (2 * H - 1)
            Else
               B = yp(k) * 10 / (2 * H - 1)
            End If
               xp(k) = MaxSpeed * rad * Cos(Angle) * Speed / 45
               xp(k) = xp(k) + imx
         Case 4   ' Wavy
            rad = k * NumParticles / H
            yp(k) = MaxSpeed / 2 * rad * Sin(Angle) / (ST + 60)
            xp(k) = (yp(k) / 2) * Sin(2 * pi# / (H / 5) * ST)
            ' Oscillate emitter
            S = 4 * pi# / MaxScaledTime * ScaledTime
            imEmit.Left = imxStart * (1 + 0.25 * Sin(S))
            xp(k) = xp(k) + imEmit.Left + imEmit.Width \ 2
            B = 20 * Sin((sAngle - 90) * d2r#)
         Case 5   ' Spirals
            rad = k * Speed * NumParticles / H
            S = 0.25 * ST * k * MaxSpeed
            yp(k) = S * Sin(rad) / NumParticles
            xp(k) = S * Cos(rad) / NumParticles
            xp(k) = xp(k) + imEmit.Left + imEmit.Width \ 2
            B = 20 * Sin((sAngle - 90) * d2r#)
         Case 6   ' Expand
            rad = k * NumParticles / H
            rad = rad * rad * Speed
            S = 0.25 * ST * k * MaxSpeed
            yp(k) = S * Sin(rad) / NumParticles
            xp(k) = S * Cos(rad) / NumParticles
            xp(k) = xp(k) + imEmit.Left + imEmit.Width \ 2
            B = 20 * Sin((sAngle - 90) * d2r#)
         Case 7   ' Wiper
            rad = k * MaxSpeed * NumParticles / H
            yp(k) = k * ST * Tan(rad) / NumParticles
            xp(k) = k * ST * Tan(rad) / NumParticles
            S = sAngle / 10 - 90
            S = S * d2r# * ScaledTime
            xp(k) = xp(k) * Cos(S)
            xp(k) = xp(k) + imEmit.Left + imEmit.Width \ 2
            B = MaxSpeed - 0.45
            xp(k) = xp(k) + (Rnd - 0.5) * B * 10
         End Select
         
         yp(k) = yp(k) + (Rnd - 0.5) * B * 10  ' B*10 gives spread
         yp(k) = yp(k) + (H - imEmit.Top)
      
         ' Test boundaries
         If xp(k) < 0 Then
            xp(k) = 0
         ElseIf xp(k) > W - 2 Then
            xp(k) = W - 2
         End If
         Select Case EmitType
         Case 0, 1, 2   ' Fountain,Hot
            If yp(k) < H - (imEmit.Top + imEmit.Height) + 1 Then
               yp(k) = H - (imEmit.Top + imEmit.Height) + 1
            ElseIf yp(k) > H - 2 Then
               yp(k) = H - 2
            End If
         Case Else
            If yp(k) < 1 Then
               yp(k) = 1
            ElseIf yp(k) > H - 2 Then
               yp(k) = H - 2
            End If
         End Select
         
         ' Color particle k Colors set for each EmitType
         ' 0 1 0
         ' 1 1 1
         ' 0 1 0
         ' Centre
         picDATA(xp(k), yp(k)) = CCen
         If xp(k) > 1 Then
            ' Left
            picDATA(xp(k) - 1, yp(k)) = CLef
            ' Right
            picDATA(xp(k) + 1, yp(k)) = CRit
         End If
         
         If yp(k) > H - (imEmit.Top + imEmit.Height) + 1 Then
            ' Top Displays above yp(k)-1
            picDATA(xp(k), yp(k) + 1) = CTop
            ' Bottom
            picDATA(xp(k), yp(k) - 1) = CBot
         End If
      Next k
      
      ' 0 Fountain, 1 Hot, 2 Spurt, 3 Spray,
      ' 4 Wavy, 5 Spirals, 6 Expand, 7 Wiper
      Select Case EmitType
      Case 0, 2, 5, 6, 7
      '
      Case 1, 3, 4 ' BLEND
         For k = 0 To NumParticles - 1
            If xp(k) > 1 Then
               ' Left
               'picDATA(xp(k) - 1, yp(k)) = CLef
               LngToRGB picDATAORG(xp(k) - 2, yp(k)), Blue, Green, Red
               picDATA(xp(k) - 2, yp(k)) = _
                  RGB((1& * Blue + LefB) \ 2, (1& * Green + LefG) \ 2, (1& * Red + LefR) \ 2)
            End If
         
            If yp(k) < H - 2 Then
               ' Top Displays above yp(k)-1
               'picDATA(xp(k), yp(k) + 1) = CTop
               LngToRGB picDATAORG(xp(k), yp(k) + 2), Blue, Green, Red
               picDATA(xp(k), yp(k) + 2) = _
                  RGB((1& * Blue + TopB) \ 2, (1& * Green + TopG) \ 2, (1& * Red + TopR) \ 2)
            End If
         
            If xp(k) > 1 And yp(k) < H - 2 Then
               ' Top-Left
               picDATA(xp(k) - 1, yp(k) + 1) = picDATAORG(xp(k), yp(k) + 2)
            End If
            
            If xp(k) < W - 2 And yp(k) < H - 2 Then
               ' Top-Right
               picDATA(xp(k) + 1, yp(k) + 1) = picDATAORG(xp(k), yp(k) + 2)
            End If
         Next k
      End Select
      
      DISPLAY PIC, picDATA()  ' NB Inline DISPLAY not worth it!
      
      TT = CLng(tmr.Elapsed)
      ' Delayer
      Do
         TimeElapsed = CLng(tmr.Elapsed)
         TimeLng = CLng(TimeElapsed / 10)
         TimeLng = 10 * TimeLng
         If TimeLng <> 0 And (TimeLng Mod 2) = 0 Then
            Exit Do
         End If
      Loop
      
      ' Increase Speed up to scrSpeed value ie MaxSpeed
      ' Grows from Emitter MaxSpeed = .5 to 4.5
      Speed = Speed + 0.05
      If Speed > MaxSpeed Then Speed = MaxSpeed

      FPS = 1000 / TT 'TimeLng
      Label1 = "fps =" & Str$(FPS)
      ScaledTime = ScaledTime + 4.5
      ' MaxScaledTime = 30,16,350,30,350,350,350,350  EmitType 0,1,2,3,4,5,6,7
      If ScaledTime > MaxScaledTime Then ScaledTime = 0
      DoEvents
   Loop Until aDone
   StopPlay
End Sub

Private Sub cmdSound_Click()
   aDone = True
   aSound = Not aSound
   If Not aSound Then
      StopPlay
      cmdSound.Caption = "Sound ON"
   Else
      cmdSound.Caption = "Sound OFF"
   End If
   PIC.SetFocus
   cmdStart_Click
End Sub

Private Sub Form_Load()
Dim FormW As Long
Dim FormH As Long
Dim BorderW As Long
Dim BorderH As Long
Dim CapH As Long
Dim MenuH As Long

'Dim mHDC  As Long
'Dim mBMPold As Long
   
   PathSpec$ = App.Path
   If Right$(PathSpec$, 1) <> "\" Then PathSpec$ = PathSpec$ & "\"
   CurrentPath$ = PathSpec$
   
   STX = Screen.TwipsPerPixelX
   STY = Screen.TwipsPerPixelY
   ' Display size
   W = 752
   H = 424 + 20
   
   PIC.Move 0, 2, W, H
   picCTRL.Move 0, PIC.Top + PIC.Height, W, 52
   imEmit.Top = PIC.Top + PIC.Height - imEmit.Height - 1
   
   ' Size Form
   CapH = GetSystemMetrics(SM_CYCAPTION)
   MenuH = GetSystemMetrics(SM_CYMENU)
   BorderW = GetSystemMetrics(SM_CXBORDER)
   BorderH = GetSystemMetrics(SM_CYBORDER)
   FormW = (W + 2 * BorderW + 4) * STX
   FormH = (H + picCTRL.Height) * STY
   FormH = FormH + (CapH + MenuH + 2 * BorderH) * STY
   Form1.Width = FormW
   Form1.Height = FormH
   
   picIN.AutoSize = True
   FileSpec$ = PathSpec$ & "Pics/Ancient.jpg"
   INITPIC
   
   Set tmr = New CTiming
   imEmit.Left = W \ 2
   OldX = imEmit.Left
   
   aScroll = False
   scrAngle.Value = 90  ' Min/Max 70 - 110
   sAngle = 90
   LabAngle = 0
   
   scrSpeed.Value = 90  ' Min/Max 10 - 90
   MaxSpeed = 4.5
   LabSpeed = Str$(4.5)
   optType_Click 0
   aScroll = True
   
   LoadWavs
   aSound = False
   
   Show

   NumParticles = 5000  ' 25,000 pixels
   ReDim xp(0 To NumParticles - 1)
   ReDim yp(0 To NumParticles - 1)
   
End Sub

Private Sub INITPIC()
' Private FileSpec$
Dim mHDC  As Long
Dim mBMPold As Long
   picIN.Picture = LoadPicture(FileSpec$)
   GetObject picIN.Image, Len(PicInfo), PicInfo
   picwidth = PicInfo.bmWidth
   picheight = PicInfo.bmHeight
   
   'Stretch whole Image from picIN to PIC
   SetStretchBltMode PIC.hdc, HALFTONE
   StretchBlt PIC.hdc, 0, 0, W, H, _
      picIN.hdc, 0, 0, picwidth, picheight, SRCCOPY
   PIC.Refresh
   With picIN
      .Picture = LoadPicture
      .Width = 4
      .Height = 4
   End With
   ReDim picDATAORG(0 To W - 1, 0 To H - 1)
   ReDim picDATA(0 To W - 1, 0 To H - 1)
   ' Public BHI As BITMAPINFOHEADER
   With BHI
      .biSize = 40
      .biPlanes = 1
      .biWidth = W
      .biHeight = H
      .biBitCount = 32
   End With
   mHDC = CreateCompatibleDC(0)
   mBMPold = SelectObject(mHDC, PIC.Image)
   If GetDIBits(mHDC, PIC.Image, 0, H, picDATAORG(0, 0), BHI, 0) = 0 Then
      MsgBox "DIB ERROR", vbCritical, "Fountain"
      Stop
      Exit Sub
   End If
   SelectObject mHDC, mBMPold
   DeleteDC mHDC
   picDATA() = picDATAORG()

End Sub

Private Sub mnuOpen_Click()
Dim Title$, Filt$, Indir$
Dim FIndex As Long
   
   aDone = True
   If aSound Then StopPlay
   aSound = False
   cmdSound.Caption = "Sound OFF"
   cmdStart.Caption = "Start"
      
   Title$ = "Open a picture file"
   Filt$ = "Pics bmp,jpg,gif|*.bmp;*.jpg;*.gif"
   FileSpec$ = ""
   Indir$ = CurrentPath$ 'Pathspec$
   Set CommonDialog1 = New cOSDialog
   CommonDialog1.ShowOpen FileSpec$, Title$, Filt$, Indir$, "", Me.hWnd, FIndex
   Set CommonDialog1 = Nothing
   
   If Len(FileSpec$) = 0 Then Exit Sub
   CurrentPath$ = FileSpec$
   
   INITPIC

End Sub

Private Sub mnuSave_Click()
Dim Title$, Filt$, Indir$
Dim FIndex As Long
   
   aDone = True
   If aSound Then StopPlay
   aSound = False
   cmdSound.Caption = "Sound OFF"
   cmdStart.Caption = "Start"
      
   Title$ = "Save Displayed Image"
   Filt$ = "Pics bmp|*.bmp"
   SaveSpec$ = ""
   Indir$ = SavePath$
   Set CommonDialog1 = New cOSDialog
   CommonDialog1.ShowSave SaveSpec$, Title$, Filt$, Indir$, "", Me.hWnd, FIndex
   Set CommonDialog1 = Nothing
   
   If Len(SaveSpec$) = 0 Then Exit Sub
   FixExtension SaveSpec$, ".bmp"
   SavePath$ = SaveSpec$
   SavePicture PIC.Image, SaveSpec$
End Sub

Private Sub FixExtension(FSpec$, Ext$)
' In: SaveSpec$ & Ext$ (".xxx")
Dim p As Long
   If Len(FSpec$) = 0 Then Exit Sub
   Ext$ = LCase$(Ext$)
   p = InStr(1, FSpec$, ".")
   If p = 0 Then
      FSpec$ = FSpec$ & Ext$
   Else
      FSpec$ = Mid$(FSpec$, 1, p - 1) & Ext$
   End If
End Sub


' Color Types
Private Sub optType_Click(Index As Integer)
'NB RGB (B,G,R)  ie B & R reversed
   aDone = True
   StopPlay
   ' Reset to background image
   picDATA() = picDATAORG()
   DISPLAY PIC, picDATA()  ' NB Inline DISPLAY not worth it!
   EmitType = Index
   Label2(0) = "Angle"
   Label2(1) = "Pressure"
   Select Case Index
   Case 0   ' Fountain
      CCen = RGB(255, 255, 255)
      CTop = RGB(255, 255, 255)
      CLef = RGB(255, 255, 255)
      CRit = RGB(255, 255, 0)
      CBot = RGB(255, 0, 0)    ' ie Blue
      STDiv = 15
      MaxScaledTime = 30
   Case 1   ' Hot
      CenR = 255: CenG = 255: CenB = 0
      CCen = RGB(CenB, CenG, CenR)
      
      TopR = 255: TopG = 0: TopB = 0
      CTop = RGB(TopB, TopG, TopR)
      
      LefR = 255: LefG = 255: LefB = 0
      CLef = RGB(LefB, LefG, LefR)
      
      RitR = 255: RitG = 255: RitB = 0
      CRit = RGB(RitB, RitG, RitB)
      
      BotR = 255: BotG = 255: BotB = 0
      CBot = RGB(BotB, BotG, BotR)
      STDiv = 45
      MaxScaledTime = 16
   Case 2   ' Spurt
      CenR = 255: CenG = 255: CenB = 255
      CCen = RGB(CenB, CenG, CenR)
      
      TopR = 200: TopG = 200: TopB = 200
      CTop = RGB(TopB, TopG, TopR)
      
      LefR = 128: LefG = 128: LefB = 128
      CLef = RGB(LefB, LefG, LefR)
      
      RitR = 0: RitG = 0: RitB = 0
      CRit = RGB(RitB, RitG, RitB)
      
      BotR = 64: BotG = 64: BotB = 64
      CBot = RGB(BotB, BotG, BotR)
      STDiv = 45
      MaxScaledTime = 350
   Case 3   ' Spray
      Label2(1) = "Spread"
      CenR = 128: CenG = 128: CenB = 128
      CCen = RGB(CenB, CenG, CenR)
      
      TopR = 255: TopG = 255: TopB = 255
      CTop = RGB(TopB, TopG, TopR)
      
      LefR = 255: LefG = 255: LefB = 255
      CLef = RGB(LefB, LefG, LefR)
      
      RitR = 0: RitG = 0: RitB = 0
      CRit = RGB(RitB, RitG, RitB)
      
      BotR = 64: BotG = 64: BotB = 64
      CBot = RGB(BotB, BotG, BotR)
      STDiv = 45
      MaxScaledTime = 30
   Case 4   ' Wavy
      Label2(0) = "Spread"
      CenR = 0: CenG = 128: CenB = 255
      CCen = RGB(CenB, CenG, CenR)
      
      TopR = 200: TopG = 200: TopB = 200
      CTop = RGB(TopB, TopG, TopR)
      
      LefR = 128: LefG = 128: LefB = 128
      CLef = RGB(LefB, LefG, LefR)
      
      RitR = 128: RitG = 128: RitB = 128
      CRit = RGB(RitB, RitG, RitB)
      
      BotR = 64: BotG = 64: BotB = 64
      CBot = RGB(BotB, BotG, BotR)
      STDiv = 45
      MaxScaledTime = 350
      imEmit.Left = W / 2
   Case 5   ' Spirals
      Label2(0) = "Spread"
      CenR = 0: CenG = 128: CenB = 255
      CCen = RGB(CenB, CenG, CenR)
      
      TopR = 200: TopG = 200: TopB = 200
      CTop = RGB(TopB, TopG, TopR)
      
      LefR = 128: LefG = 128: LefB = 128
      CLef = RGB(LefB, LefG, LefR)
      
      RitR = 128: RitG = 128: RitB = 128
      CRit = RGB(RitB, RitG, RitB)
      
      BotR = 64: BotG = 64: BotB = 64
      CBot = RGB(BotB, BotG, BotR)
      STDiv = 45
      MaxScaledTime = 350
      imEmit.Left = W / 2
   Case 6   ' Expand
      Label2(0) = "Spread"
      CenR = 255: CenG = 0: CenB = 0
      CCen = RGB(CenB, CenG, CenR)
      
      TopR = 255: TopG = 0: TopB = 0
      CTop = RGB(TopB, TopG, TopR)
      
      LefR = 255: LefG = 255: LefB = 0
      CLef = RGB(LefB, LefG, LefR)
      
      RitR = 128: RitG = 128: RitB = 128
      CRit = RGB(RitB, RitG, RitB)
      
      BotR = 64: BotG = 64: BotB = 64
      CBot = RGB(BotB, BotG, BotR)
      STDiv = 45
      MaxScaledTime = 350
      imEmit.Left = W / 2
   Case 7   ' Wiper
      Label2(0) = "Speed"
      Label2(1) = "Spread"
      CenR = 255: CenG = 0: CenB = 0
      CCen = RGB(CenB, CenG, CenR)
      
      TopR = 255: TopG = 0: TopB = 0
      CTop = RGB(TopB, TopG, TopR)
      
      LefR = 255: LefG = 255: LefB = 0
      CLef = RGB(LefB, LefG, LefR)
      
      RitR = 128: RitG = 128: RitB = 128
      CRit = RGB(RitB, RitG, RitB)
      
      BotR = 64: BotG = 64: BotB = 64
      CBot = RGB(BotB, BotG, BotR)
      STDiv = 45
      MaxScaledTime = 350
      imEmit.Left = W / 2
   End Select
   If aScroll Then cmdStart_Click
End Sub

' Control Angle
Private Sub scrAngle_Scroll()
   Call scrAngle_Change
End Sub
Private Sub scrAngle_Change()
' 70 - 110
Dim S As Single
   If Not aScroll Then Exit Sub
   S = scrAngle.Value
   S = S - 90  ' -20 to +20
   LabAngle = Str$(S)
   LabAngle.Refresh
   S = -S
   sAngle = S + 90
End Sub

' Control Speed
Private Sub scrSpeed_Scroll()
   Call scrSpeed_Change
End Sub
Private Sub scrSpeed_Change()
' 10 -> 90
Dim i As Single
   If Not aScroll Then Exit Sub
   i = scrSpeed.Value / 20  ' .5 -> 4.5
   LabSpeed = Str$(i)
   LabSpeed.Refresh
   MaxSpeed = i
End Sub

' Emitter
Private Sub imEmit_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
   imx = X
   imy = Y
   imMouseDown = True
End Sub

Private Sub imEmit_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Dim NewX As Long
Dim NewY As Long
   If imMouseDown Then
      ' Calculate new position
      NewX = OldX + (X - imx) \ STX
      If NewX < 0 Then
         NewX = 0
      ElseIf NewX > W - imEmit.Width Then
         NewX = W - imEmit.Width
      End If
      imEmit.Left = NewX
      OldX = NewX
      ' Calculate new position
      NewY = OldY + (Y - imy) \ STY
      If NewY < 10 Then
         NewY = 10
      ElseIf NewY > H - imEmit.Height Then
         NewY = H - imEmit.Height
      End If
      imEmit.Top = NewY
      OldY = NewY
   End If
End Sub

Private Sub imEmit_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
   imMouseDown = False
End Sub

' Exit
Private Sub mnuExit_Click()
   aDone = True
   StopPlay
   Set tmr = Nothing
   Unload Me
   End
End Sub

Private Sub Form_Unload(Cancel As Integer)
   aDone = True
   StopPlay
   Set tmr = Nothing
   Unload Me
   End
End Sub


