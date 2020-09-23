VERSION 5.00
Begin VB.Form Form1 
   Appearance      =   0  'Flat
   BackColor       =   &H80000004&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MegaSoccer"
   ClientHeight    =   9300
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9600
   Icon            =   "Form1.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   620
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox PicStatusbarStamina 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1440
      Picture         =   "Form1.frx":1782
      ScaleHeight     =   150
      ScaleWidth      =   1500
      TabIndex        =   11
      Top             =   7920
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.PictureBox PicStatusShootC 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1440
      Picture         =   "Form1.frx":3221
      ScaleHeight     =   150
      ScaleWidth      =   1500
      TabIndex        =   10
      Top             =   7800
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.PictureBox PicStatusShoot 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   1440
      Picture         =   "Form1.frx":49A5
      ScaleHeight     =   150
      ScaleWidth      =   1500
      TabIndex        =   9
      Top             =   7680
      Visible         =   0   'False
      Width           =   1530
   End
   Begin VB.PictureBox PicPlayerM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   360
      Picture         =   "Form1.frx":5F1A
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   8
      Top             =   7920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PicPlayer 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   0
      Picture         =   "Form1.frx":734F
      ScaleHeight     =   345
      ScaleWidth      =   345
      TabIndex        =   7
      Top             =   7920
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.PictureBox PicBallM 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   240
      Picture         =   "Form1.frx":8842
      ScaleHeight     =   150
      ScaleWidth      =   150
      TabIndex        =   6
      Top             =   7680
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.PictureBox PicBall 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   180
      Left            =   0
      Picture         =   "Form1.frx":9B29
      ScaleHeight     =   150
      ScaleWidth      =   150
      TabIndex        =   5
      Top             =   7680
      Visible         =   0   'False
      Width           =   180
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Sound 2"
      Height          =   375
      Left            =   2760
      TabIndex        =   2
      Top             =   7200
      Width           =   1095
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load big map"
      Height          =   375
      Left            =   0
      TabIndex        =   4
      Top             =   7200
      Width           =   1575
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Sound 1"
      Height          =   375
      Left            =   1680
      TabIndex        =   3
      Top             =   7200
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   1000
      Left            =   9000
      Top             =   7320
   End
   Begin VB.PictureBox Pic 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H00000000&
      BorderStyle     =   0  'None
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   7200
      Left            =   0
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   640
      TabIndex        =   0
      Top             =   0
      Width           =   9600
   End
   Begin VB.PictureBox Pic1 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   19260
      Left            =   0
      Picture         =   "Form1.frx":AE2C
      ScaleHeight     =   19200
      ScaleWidth      =   11520
      TabIndex        =   1
      Top             =   8400
      Visible         =   0   'False
      Width           =   11580
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
dsVidaExtra.SetVolume Volume
dsVidaExtra.SetCurrentPosition Rnd * 500 'Namjesi br.Buffer-a koji ce koristiti
dsVidaExtra.Play DSBPLAY_DEFAULT 'Ispusta zvuk

'Dim Dis As Single
'Dis = ((X1 - X2) ^ 2) + ((Y1 - Y2) ^ 2)
'Distance = Sqr(Dis) 'za detektiranje objekta u krugu oko X1
End Sub

Private Sub Command2_Click()
dsKick1.SetVolume Volume
dsKick1.SetCurrentPosition Rnd * 500 'Namjesi br.Buffer-a koji ce koristiti
dsKick1.Play DSBPLAY_DEFAULT 'Ispusta zvuk
End Sub

Private Sub Command3_Click()
Pic1.Picture = LoadPicture(App.Path & "\teren na grubo2.jpg")
End Sub

Private Sub Form_Activate()
Pi = (4 * Atn(1)) / 180

'Init Form1.hwnd  'It prepares DirectDraw for Resolution changing
'DX_Draw.SetDisplayMode 640, 480, 16, 0, DDSDM_DEFAULT

InitializeDS 'It prepares DirectSound to be used
InitializeDI 'It prepares DirectInput to be used

BallX = (Pic.Width / 2) - (PicBall.Width / 2) + 50
BallY = (Pic.Height / 2) - (PicBall.Height / 2) + 50
PlayerX = (Pic.Width / 2) - (PicPlayer.Width / 2)
PlayerY = (Pic.Height / 2) - (PicPlayer.Height / 2)

Draw
End Sub
Private Sub Form_Unload(Cancel As Integer)
Unload Form1
End
End Sub
Private Sub Timer1_Timer()
Form1.Caption = FPS
FPS = 0
End Sub
