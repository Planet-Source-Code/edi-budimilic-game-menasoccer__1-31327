Attribute VB_Name = "DirectInputMod"
Option Explicit

Public Type LastMovePT
 Up As String
 Down As String
 Left As String
 Right As String
 RShift As String
 RShiftUp As String
 Slash As String
End Type
Public LastMove As LastMovePT

Public DI As DirectInput
Public DIKeyboard As DirectInputDevice
Public DIMouse As DirectInputDevice

Public Sub DetectKeyboard()
Dim KeyboardState(0 To 255) As Byte
DIKeyboard.Acquire
DIKeyboard.GetDeviceState 256, KeyboardState(0)
    
'It give you a state of key: 1 - KeyDown, 2 - Hold, 3 - KeyUp
If (KeyboardState(DIK_ESCAPE)) <> 0 Then End
If (KeyboardState(DIK_UP)) <> 0 Then LastMove.Up = "1" Else LastMove.Up = "0"    'Micanje Igraca Gore itd...
If (KeyboardState(DIK_DOWN)) <> 0 Then LastMove.Down = "1" Else LastMove.Down = "0"
If (KeyboardState(DIK_LEFT)) <> 0 Then LastMove.Left = "1" Else LastMove.Left = "0"
If (KeyboardState(DIK_RIGHT)) <> 0 Then LastMove.Right = "1" Else LastMove.Right = "0"
If (KeyboardState(DIK_RSHIFT)) <> 0 Then LastMove.RShift = "1" Else LastMove.RShift = "0"


If (KeyboardState(DIK_UP)) <> 0 Or (KeyboardState(DIK_DOWN)) <> 0 Or (KeyboardState(DIK_LEFT)) <> 0 Or (KeyboardState(DIK_RIGHT)) <> 0 Then PlayerSpeed = PlayerSpeed + 2
If PlayerSpeed > "31" Then PlayerSpeed = PlayerSpeed - 1
If Not PlayerSpeed < "1" Then PlayerSpeed = PlayerSpeed - 1
If LastMove.Up = 1 Then PlayerY = PlayerY - (PlayerSpeed / 10)
If LastMove.Down = 1 Then PlayerY = PlayerY + (PlayerSpeed / 10)
If LastMove.Left = 1 Then PlayerX = PlayerX - (PlayerSpeed / 10)
If LastMove.Right = 1 Then PlayerX = PlayerX + (PlayerSpeed / 10)

End Sub

Public Sub Destroy_DI()
    Set DIKeyboard = Nothing
    Set DIMouse = Nothing
    Set DI = Nothing
End Sub
Public Sub InitializeDI()
Set DI = DX.DirectInputCreate()
Set DIKeyboard = DI.CreateDevice("GUID_SysKeyboard")
    

DIKeyboard.SetCommonDataFormat DIFORMAT_KEYBOARD
' Za gdje ce se koristiti Tastatura - nonexclusive je svugdje
DIKeyboard.SetCooperativeLevel Form1.hwnd, DISCL_BACKGROUND Or DISCL_NONEXCLUSIVE

DIKeyboard.Acquire

Set DIMouse = DI.CreateDevice("GUID_SysMouse")

DIMouse.SetCommonDataFormat DIFORMAT_MOUSE

DIMouse.SetCooperativeLevel Form1.hwnd, DISCL_FOREGROUND Or DISCL_EXCLUSIVE
End Sub

