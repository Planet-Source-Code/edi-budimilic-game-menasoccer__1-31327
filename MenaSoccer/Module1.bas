Attribute VB_Name = "Module1"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Public Declare Function SetPixel Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function SetPixelV Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Public Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nPICWidth As Long, ByVal nPICHeight As Long, ByVal dwRop As Long) As Long
Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal lpString As String, ByVal nCount As Long) As Long

Public Const SRCCOPY = &HCC0020
Public Const SRCPAINT = &HEE0086
Public Const SRCAND = &H8800C6
Public Const SRCINVERT = &H660046
Public Const SRCERASE = &H440328
Public Const SND_SYNC = &H0
Public Const SND_ASYNC = &H1

Public ShootBall As Integer, PStamina As Integer
Public XX As Integer, YY As Integer, FPS As Integer, Pi
Public Ii, Ii2 ' Frame limiter
Public Dis As Integer, Ang As Integer ' Dis=Distance - Ang=Kut od A i B tocke
Public BallX, BallY, BallZ, BallSpeed, BallAngle As Integer ' X,Y Lopte, Njezina Brzina
Public PlayerX, PlayerY, PlayerSpeed As Integer, HoldingLong As Integer  ' X,Y Igraca, Njegova Brzina,koliko dugo drzi loptu uza se
Sub Init(hwnd As Long)
On Error GoTo CrapThingDontWork

'creates direct draw
Set DX_Draw = DX.DirectDrawCreate("")

DX_Draw.SetCooperativeLevel hwnd, DDSCL_FULLSCREEN Or DDSCL_ALLOWMODEX Or DDSCL_EXCLUSIVE
Exit Sub

CrapThingDontWork:
MsgBox "Cannot activate DirectX 7! Make sure you have it installed correctly!", vbExclamation, "Menasoccer Error!"
End
End Sub
Public Sub CameraControl()
GetAngle 340, 200, BallX, BallY
If Dis > 10 Then
XX = XX + (Cos(Ang * Pi) * (Dis - 90))
YY = YY + (Sin(Ang * Pi) * (Dis - 90))
BallX = BallX - (Cos(Ang * Pi) * (Dis - 90))
BallY = BallY - (Sin(Ang * Pi) * (Dis - 90))
End If
Exit Sub 'maknuti ovo za skoro normalnu kameru
If Dis > 80 And Dis < 90 Then
XX = XX + (Cos(Ang * Pi) * (3))
YY = YY + (Sin(Ang * Pi) * (3))
BallX = BallX - (Cos(Ang * Pi) * (3))
BallY = BallY - (Sin(Ang * Pi) * (3))
End If
Call TextOut(Form1.Pic.hdc, 400, 0, Dis & "          ", 10)

'STARO STARO STARO - MALO NESTABILNO TRZAJUCA KAMERA
'Dim PLAS As Integer
'PLAS = (Dis \ 30)
'If PLAS > 4 Then PLAS = 4
'Call TextOut(Form1.Pic.HDC, 400, 0, PLAS & "          ", 10)
'If BallY > 260 Then YY = YY + PLAS: BallY = BallY - PLAS: GoTo 1
'If BallY < 240 Then YY = YY - PLAS: BallY = BallY + PLAS: GoTo 1
'1
'If BallX > 300 Then XX = XX + PLAS: BallX = BallX - PLAS
'If BallX < 280 Then XX = XX - PLAS: BallX = BallX + PLAS: Exit Sub

End Sub

Public Sub Draw()
1
DoEvents
DetectKeyboard
Form1.Pic.Cls
'Crta Nogometni teren
Call BitBlt(Form1.Pic.hdc, 0, 0, Form1.Pic1.Width, Form1.Pic1.Height, Form1.Pic1.hdc, XX + 0, YY + 0, SRCCOPY)

CameraControl ' Kamera prai metu

'Crta Ball
Call BitBlt(Form1.Pic.hdc, BallX, BallY, 12, 12, Form1.PicBallM.hdc, 0, 0, SRCAND)
Call BitBlt(Form1.Pic.hdc, BallX, BallY, 12, 12, Form1.PicBall.hdc, 0, 0, SRCPAINT)

'Da nacrta gdje je Player ako sam van dometa ekrana - ??
GetAngle (BallX) + 6, (BallY) + 6, (PlayerX - XX + Form1.PicPlayer.Width / 2), (PlayerY - YY + Form1.PicPlayer.Height / 2)
If Dis > 186 + (Dis / 7) Then
SetPixel Form1.Pic.hdc, BallX + (Cos(Ang * Pi)) * (180 + (Dis / 7)), BallY + (Sin(Ang * Pi)) * (180 + (Dis / 7)), vbYellow
SetPixel Form1.Pic.hdc, BallX + (Cos(Ang * Pi)) * (182 + (Dis / 8)), BallY + (Sin(Ang * Pi)) * (182 + (Dis / 8)), vbYellow
SetPixel Form1.Pic.hdc, BallX + (Cos(Ang * Pi)) * (184 + (Dis / 9)), BallY + (Sin(Ang * Pi)) * (184 + (Dis / 9)), vbYellow
SetPixel Form1.Pic.hdc, BallX + (Cos(Ang * Pi)) * (186 + (Dis / 10)), BallY + (Sin(Ang * Pi)) * (186 + (Dis / 10)), vbYellow
End If

Form1.Pic.Circle ((PlayerX - XX) + Form1.PicPlayer.Width / 2 - 2, (PlayerY - YY) + Form1.PicPlayer.Height / 2 - 2), 10, vbYellow ' Crta krug oko Playera
'Crta Player
Call BitBlt(Form1.Pic.hdc, PlayerX - XX, PlayerY - YY, 21, 21, Form1.PicPlayerM.hdc, 0, 0, SRCAND)
Call BitBlt(Form1.Pic.hdc, PlayerX - XX, PlayerY - YY, 21, 21, Form1.PicPlayer.hdc, 0, 0, SRCPAINT)

Collide 'Detektira Player s loptom

'Crta Statusbar Stamina
Call BitBlt(Form1.Pic.hdc, Form1.Pic.Width \ 2 - 170, Form1.Pic.Height - 20, 100, 10, Form1.PicStatusShoot.hdc, 0, 0, SRCCOPY)
Call BitBlt(Form1.Pic.hdc, Form1.Pic.Width \ 2 - 170, Form1.Pic.Height - 20, (100 / 7000) * PStamina, 10, Form1.PicStatusbarStamina.hdc, 0, 0, SRCCOPY)
If Not PStamina > 6500 Then PStamina = PStamina * 1 + 1


FrameRate ' Limiter i Text na ekran
GoTo 1
End Sub
Public Sub FrameRate()
For Ii = 0 To Ii2 'Frame limiter ovih par redova
Next Ii
If FPS > 150 Then Ii2 = Ii2 + 1000
If FPS > 120 Then Ii2 = Ii2 + 1000
If FPS > 90 Then Ii2 = Ii2 + 1000
If FPS > 70 Then Ii2 = Ii2 + 1000
If FPS > 60 Then Ii2 = Ii2 + 1000
If FPS < 60 Then Ii2 = Ii2 - 20
'If Ii2 < 1 Then Ii2 = "0"
Call TextOut(Form1.Pic.hdc, 5, 0, "Fps= " & Form1.Caption & "    FpsSkiped= " & Ii2 \ 100 & "         ", 34) ' TEXT da izbaci framerate
FPS = FPS + 1
End Sub

Public Sub Collide()
GetAngle (PlayerX - XX + Form1.PicPlayer.Width / 2), (PlayerY - YY + Form1.PicPlayer.Height / 2), (BallX) + 6, (BallY) + 6

If Dis < 40 Then ' Napucavanje lopte -  40=daljina kada ce poceti puniti snagu da pukne loptu
If LastMove.RShift Then
ShootBall = ShootBall + 3: If ShootBall > 90 Then ShootBall = 90 Else If ShootBall < 0 Then ShootBall = 0
Else
If Dis < 20 Then If ShootBall > 30 Then BallZ = 1: BallSpeed = PlayerSpeed + ShootBall:  SKick1 -(Rnd * 500): PStamina = PStamina - ShootBall: ShootBall = 0   ' Ovo predzadnje Zvuk udrca u loptu / napucavanje
End If
Else
If Dis > 40 Then ShootBall = 0
End If
'Crta StatusShooting ball
Call BitBlt(Form1.Pic.hdc, Form1.Pic.Width \ 2 - 50, Form1.Pic.Height - 20, 100, 10, Form1.PicStatusShoot.hdc, 0, 0, SRCCOPY)
Call BitBlt(Form1.Pic.hdc, Form1.Pic.Width \ 2 - 50, Form1.Pic.Height - 20, (100 / 90) * ShootBall, 10, Form1.PicStatusShootC.hdc, 0, 0, SRCCOPY)

HoldingLong = HoldingLong + 1
If Dis < 16 And Not BallZ = 1 Then
If HoldingLong < 90 Then GuraLoptuSta PlayerSpeed + 3 Else GuraLoptu PlayerSpeed + 10  'za sve strane
Else
If Dis > 25 Then If HoldingLong > 89 Then HoldingLong = 0
End If


Curb ' Felsanje lopte

'Da se odbije od rubova slike/pozadine/terena
If BallX < 0 - XX Then BallAngle = BallAngle - BallAngle
If BallY < 0 - YY Then BallAngle = -BallAngle
If BallX + Form1.PicBall.Width > Form1.Pic1.Width - XX Then BallAngle = -BallAngle - 180
If BallY + Form1.PicBall.Height > Form1.Pic1.Height - YY Then BallAngle = -BallAngle
'Da poravna stupnjeve od naredbi koje odbijaju od ruba ekrana
If BallAngle > 359 Then BallAngle = BallAngle - 360
If BallAngle < 0 Then BallAngle = BallAngle * 1 + 360

If Not BallSpeed < 1 Then BallSpeed = BallSpeed - 1  ' Usporava loptu
        'Pomice loptu na zeljenu poziciju kuta...
        BallX = BallX + (Cos(BallAngle * Pi) * (BallSpeed / 10))
        BallY = BallY + (Sin(BallAngle * Pi) * (BallSpeed / 10))
End Sub
Public Sub Curb() ' Felsanje Lopte
If Dis > 12 And Dis < 250 And BallZ > 0 Then ' Ako je dalje od mene i u zraku da :
If BallAngle > 225 And BallAngle < 315 Then ' Felsanje lopte ljevodesno ako idem prema gore
If LastMove.Left Then BallAngle = BallAngle - 1
If LastMove.Right Then BallAngle = BallAngle + 1
End If
If BallAngle > 45 And BallAngle < 135 Then ' Felsanje lopte ljevodesno ako idem prema dole
If LastMove.Left Then BallAngle = BallAngle + 1
If LastMove.Right Then BallAngle = BallAngle - 1
End If
If BallAngle > 135 And BallAngle < 225 Then ' Felsanje lopte goredole ako idem prema ljevo
If LastMove.Up Then BallAngle = BallAngle + 1
If LastMove.Down Then BallAngle = BallAngle - 1
End If
If BallAngle > 315 Or BallAngle < 45 Then ' Felsanje lopte goredole ako idem prema desno
If LastMove.Up Then BallAngle = BallAngle - 1
If LastMove.Down Then BallAngle = BallAngle + 1
End If
If BallSpeed < 1 Then BallZ = 0 ' Ako je pala brzina lopte na 0 da oznaci da je lopta na podu
End If
End Sub
Public Sub SKick1(Volume) ' Zvuk gurkanja lopte
On Error Resume Next
dsKick1.SetVolume Volume
dsKick1.SetCurrentPosition Rnd * 100 'Namjesi br.Buffer-a koji ce koristiti
dsKick1.Play DSBPLAY_DEFAULT   'Ispusta zvuk
End Sub
Public Function GetAngle(X2, Y2, x, y) 'XY=Item   X2Y2=0
If y = Y2 And x > X2 Then ' Desno
   Ang = 0
ElseIf y > Y2 And x = X2 Then ' Dole
   Ang = 90
ElseIf y = Y2 And x < X2 Then ' Ljevo
   Ang = 180
ElseIf y < Y2 And x = X2 Then ' Gore
   Ang = 270
Else
   Ang = Abs(Atn((x - X2) / (y - Y2)) * (4 * Atn(1)) * 18)
   If y > Y2 And x < X2 Then ' Dole Ljevo
      Ang = Ang + 90
   ElseIf y > Y2 And x > X2 Then ' Dole Desno
      Ang = -Ang + 90
   ElseIf y < Y2 And x < X2 Then ' Gore Ljevo
      Ang = -Ang + 270
   ElseIf y < Y2 And x > X2 Then ' Gore Desno
      Ang = Ang + 270
   End If
End If
'Ang je odgovor na Kut
Dis = Sqr(((x - X2) ^ 2) + ((y - Y2) ^ 2)) 'Dis je odgovor na daljinu izbedju objekata
End Function
Public Sub GuraLoptu(BallSpd As Integer)
BallSpeed = BallSpd
BallAngle = Ang
SKick1 -(Int(Rnd * 2000) + 1500) ' Zvuk udrca u loptu / dodavanje

If LastMove.Up And LastMove.Left Then 'Gore ljevo
If BallAngle > 180 And BallAngle < 220 Then BallAngle = BallAngle + (230 - BallAngle)
If BallAngle > 230 And BallAngle < 270 Then BallAngle = BallAngle - (BallAngle - 220)
BallSpeed = BallSpeed + 10
Exit Sub
End If
If LastMove.Up And LastMove.Right Then ' Gore desno
If BallAngle > 270 And BallAngle < 310 Then BallAngle = BallAngle + (320 - BallAngle)
If BallAngle > 320 And BallAngle < 360 Then BallAngle = BallAngle - (BallAngle - 310)
BallSpeed = BallSpeed + 10
Exit Sub
End If
If LastMove.Down And LastMove.Left Then ' Dole ljevo
If BallAngle > 90 And BallAngle < 130 Then BallAngle = BallAngle + (140 - BallAngle)
If BallAngle > 140 And BallAngle < 180 Then BallAngle = BallAngle - (BallAngle - 130)
BallSpeed = BallSpeed + 10
Exit Sub
End If
If LastMove.Down And LastMove.Right Then ' Dole right
If BallAngle > 0 And BallAngle < 40 Then BallAngle = BallAngle + (50 - BallAngle)
If BallAngle > 50 And BallAngle < 90 Then BallAngle = BallAngle - (BallAngle - 40)
BallSpeed = BallSpeed + 10
Exit Sub
End If

'Osnovni smjerovi
If LastMove.Up Then
If BallAngle > 180 And BallAngle < 265 Then BallAngle = BallAngle + (275 - BallAngle)
If BallAngle > 275 And BallAngle < 360 Then BallAngle = BallAngle - (BallAngle - 265)
End If
If LastMove.Down Then
If BallAngle > 0 And BallAngle < 85 Then BallAngle = BallAngle + (95 - BallAngle)
If BallAngle > 95 And BallAngle < 180 Then BallAngle = BallAngle - (BallAngle - 85)
End If
If LastMove.Left Then
If BallAngle > 90 And BallAngle < 175 Then BallAngle = BallAngle + (185 - BallAngle)
If BallAngle > 185 And BallAngle < 270 Then BallAngle = BallAngle - (BallAngle - 175)
End If
If LastMove.Right Then
If BallAngle > 270 And BallAngle < 355 Then BallAngle = BallAngle + (5 - BallAngle)
If BallAngle > 5 And BallAngle < 90 Then BallAngle = BallAngle - (BallAngle - 355)
End If


End Sub
Public Sub GuraLoptuSta(BallSpd As Integer)
BallSpeed = BallSpd
BallAngle = Ang
SKick1 -(Int(Rnd * 2000) + 1500) ' Zvuk udrca u loptu / dodavanje

If PStamina < 1 Then PStamina = "7000"
PStamina = PStamina - 5

If LastMove.Up And LastMove.Left Then 'Gore ljevo
If BallAngle > 180 And BallAngle < 220 Then BallAngle = BallAngle + (230 - BallAngle)
If BallAngle > 230 And BallAngle < 270 Then BallAngle = BallAngle - (BallAngle - 220)
BallSpeed = BallSpeed + 10
BallX = PlayerX - XX + 2
BallY = PlayerY - YY + 2
Exit Sub
End If
If LastMove.Up And LastMove.Right Then ' Gore desno
If BallAngle > 270 And BallAngle < 310 Then BallAngle = BallAngle + (320 - BallAngle)
If BallAngle > 320 And BallAngle < 360 Then BallAngle = BallAngle - (BallAngle - 310)
BallSpeed = BallSpeed + 10
BallY = PlayerY - YY + 2
BallX = 10 + PlayerX - XX + 2
Exit Sub
End If
If LastMove.Down And LastMove.Left Then ' Dole ljevo
If BallAngle > 90 And BallAngle < 130 Then BallAngle = BallAngle + (140 - BallAngle)
If BallAngle > 140 And BallAngle < 180 Then BallAngle = BallAngle - (BallAngle - 130)
BallSpeed = BallSpeed + 10
BallY = 10 + PlayerY - YY + 2
BallX = PlayerX - XX + 2
Exit Sub
End If
If LastMove.Down And LastMove.Right Then ' Dole right
If BallAngle > 0 And BallAngle < 40 Then BallAngle = BallAngle + (50 - BallAngle)
If BallAngle > 50 And BallAngle < 90 Then BallAngle = BallAngle - (BallAngle - 40)
BallSpeed = BallSpeed + 10
BallY = 10 + PlayerY - YY + 2
BallX = 10 + PlayerX - XX + 2
Exit Sub
End If

'Osnovni smjerovi
If LastMove.Up Then
If BallAngle > 180 And BallAngle < 265 Then BallAngle = BallAngle + (275 - BallAngle)
If BallAngle > 275 And BallAngle < 360 Then BallAngle = BallAngle - (BallAngle - 265)
BallY = PlayerY - YY + 2
End If
If LastMove.Down Then
If BallAngle > 0 And BallAngle < 85 Then BallAngle = BallAngle + (95 - BallAngle)
If BallAngle > 95 And BallAngle < 180 Then BallAngle = BallAngle - (BallAngle - 85)
BallY = 10 + PlayerY - YY + 2
End If
If LastMove.Left Then
If BallAngle > 90 And BallAngle < 175 Then BallAngle = BallAngle + (185 - BallAngle)
If BallAngle > 185 And BallAngle < 270 Then BallAngle = BallAngle - (BallAngle - 175)
BallX = PlayerX - XX + 2
End If
If LastMove.Right Then
If BallAngle > 270 And BallAngle < 355 Then BallAngle = BallAngle + (5 - BallAngle)
If BallAngle > 5 And BallAngle < 90 Then BallAngle = BallAngle - (BallAngle - 355)
BallX = 10 + PlayerX - XX + 2
End If
End Sub

