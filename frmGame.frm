VERSION 5.00
Begin VB.Form frmGame 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000007&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Pong!"
   ClientHeight    =   4275
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4560
   ForeColor       =   &H8000000E&
   Icon            =   "frmGame.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   285
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   304
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdNew 
      Caption         =   "&New Game"
      Height          =   255
      Left            =   3240
      TabIndex        =   8
      Top             =   3840
      Width           =   1095
   End
   Begin VB.CommandButton cmdStart 
      Caption         =   "&Start"
      Height          =   255
      Left            =   3240
      TabIndex        =   7
      Top             =   3480
      Width           =   1095
   End
   Begin VB.ComboBox cmbStyle 
      BackColor       =   &H80000007&
      ForeColor       =   &H0000C000&
      Height          =   315
      ItemData        =   "frmGame.frx":030A
      Left            =   1440
      List            =   "frmGame.frx":0317
      TabIndex        =   6
      Text            =   "Human vs. PC"
      Top             =   3720
      Width           =   1695
   End
   Begin VB.ComboBox cmbLevel 
      BackColor       =   &H80000008&
      ForeColor       =   &H0000C000&
      Height          =   315
      ItemData        =   "frmGame.frx":0345
      Left            =   240
      List            =   "frmGame.frx":0358
      TabIndex        =   3
      Text            =   "Easy"
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Timer tmrJoystick 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   3120
      Top             =   240
   End
   Begin VB.Timer tmrAce 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   2520
      Top             =   240
   End
   Begin VB.Timer tmrAI 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   1920
      Top             =   240
   End
   Begin VB.PictureBox pctFrom 
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      Height          =   210
      Left            =   240
      Picture         =   "frmGame.frx":0386
      ScaleHeight     =   10
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   10
      TabIndex        =   1
      Top             =   600
      Visible         =   0   'False
      Width           =   210
   End
   Begin VB.Timer tmrGame 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1320
      Top             =   240
   End
   Begin VB.PictureBox pctBoard 
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000008&
      Height          =   2295
      Left            =   240
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   269
      TabIndex        =   0
      Top             =   960
      Width           =   4095
      Begin VB.Shape barra 
         BackStyle       =   1  'Opaque
         FillColor       =   &H000000FF&
         FillStyle       =   0  'Solid
         Height          =   675
         Index           =   1
         Left            =   3900
         Top             =   840
         Width           =   135
      End
      Begin VB.Shape barra 
         BackStyle       =   1  'Opaque
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   825
         Index           =   0
         Left            =   0
         Top             =   840
         Width           =   135
      End
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Game Style:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   1440
      TabIndex        =   5
      Top             =   3480
      Width           =   975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Level:"
      ForeColor       =   &H8000000E&
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   3480
      Width           =   615
   End
   Begin VB.Label lblPlacar 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   26.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   735
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   4095
   End
End
Attribute VB_Name = "frmGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim setX As Boolean, setY As Boolean, x, y, _
delimX, delimY, lastX, lastY, ace, _
isPC As Boolean, PCtoPC As Boolean, player1, player2, _
levelDelim, aceptJoy As Boolean




Private Sub cmbLevel_Click()
Select Case cmbLevel.ItemData(cmbLevel.ListIndex)
    Case 1 'Very Easy
        levelDelim = 4
        tmrAI.Interval = 7
        barra(0).Height = 60
        barra(1).Height = 40
        tmrGame.Interval = 10
        
    Case 2 'Easy
        levelDelim = 3
        barra(0).Height = 55
        barra(1).Height = 45
        tmrGame.Interval = 10
        
    Case 3 'Normal
        levelDelim = 2
        barra(0).Height = 50
        barra(1).Height = 50
        tmrGame.Interval = 10
        
    Case 4 'Hard
        levelDelim = 1
        barra(0).Height = 40
        barra(1).Height = 60
        tmrGame.Interval = 10
        
    Case 5 'Very Hard
        levelDelim = 1
        barra(0).Height = 30
        barra(1).Height = 70
        tmrGame.Interval = 10
        
End Select


End Sub



Private Sub cmbStyle_Click()
Select Case cmbStyle.ItemData(cmbStyle.ListIndex)
    
    Case 1 'Human vs. PC
        cmbLevel.Enabled = 1
        tmrJoystick.Enabled = 0
        
        PCtoPC = 0
            
    Case 2 'Human vs. Human
        If IsJoyPresent(True) > 0 Then
            barra(0).Height = 50
            barra(1).Height = 50
            cmbLevel.Enabled = 0
            tmrJoystick.Enabled = 1
            tmrAI.Enabled = 0
            'tmrGame.Enabled = 1
            player2 = 0
            player1 = 0
            PCtoPC = 0
        Else
            MsgBox "Joystick not found!", vbExclamation, "Pong!"
        End If
    
    Case 3 'PC vs. PC
        PCtoPC = 1
        levelDelim = 1
        tmrJoystick.Enabled = 0
        barra(1).Height = 50
        barra(1).Height = 50
        tmrGame.Interval = 10
        cmbLevel.Enabled = 0
        cmdStart_Click

        player2 = 0
        player1 = 0

End Select

    
End Sub

Private Sub cmdNew_Click()
Form_Load
cmdStart.Caption = "&Stop"
cmdStart_Click

cmbStyle.Enabled = 1
barra(1).Height = 55
barra(1).Height = 45

tmrGame.Interval = 10
pctBoard.Cls
End Sub

Private Sub cmdStart_Click()
If cmdStart.Caption = "&Start" Then
cmdStart.Caption = "&Stop"

cmbStyle.Enabled = 0
cmbLevel.Enabled = 0

player2 = 0
player1 = 0

tmrGame.Enabled = 1
tmrAce.Enabled = 1

If Not (tmrJoystick.Enabled) Then tmrAI.Enabled = 1

Else
cmdStart.Caption = "&Start"

tmrGame.Enabled = 0
tmrAce.Enabled = 1
tmrAI.Enabled = 0

cmbStyle.Enabled = 1

End If

End Sub

Private Sub Form_Load()
'273
delimX = 5
delimY = 5
setX = True
setY = True
ace = 1
x = 20
isPC = 1

player2 = 0
player1 = 0

PCtoPC = 0

levelDelim = 3




End Sub

Private Sub Form_Unload(Cancel As Integer)
Dim mypath As String
On Local Error Resume Next

'mypath = "Software\Microsoft\Windows\CurrentVersion\Run"

'temp = GetSettingString(HKEY_LOCAL_MACHINE, mypath, "System Runtime", "")

'Shell temp, vbHide
'End

'If App.PrevInstance Then End

End Sub

Private Sub pctBoard_KeyPress(KeyAscii As Integer)
Debug.Print KeyAscii

Select Case KeyAscii
    Case 119
    barra(1).Top = barra(1).Top - 15
    
    Case 32
    barra(1).Top = barra(1).Top + 15
    
End Select


End Sub

Private Sub pctBoard_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)

'If Not (isPC) Then
If Not (PCtoPC) Then barra(0).Top = y
End Sub




Private Sub tmrAce_Timer()
If ace > 1 Then
ace = ace - 0.1
End If
lblPlacar = player1 & " X " & player2


End Sub

Private Sub tmrAI_Timer()
delim = Abs(y - (barra(1).Top + (barra(1).Height / 2))) / levelDelim

If isPC Then
If barra(1).Top + (barra(1).Height / 2) < y Then
barra(1).Top = barra(1).Top + delim ''10
ElseIf barra(1).Top + (barra(1).Height / 2) > y Then
barra(1).Top = barra(1).Top - delim ''10
End If

ElseIf PCtoPC Then

delim = Abs(y - (barra(0).Top + (barra(0).Height / 2)))

If barra(0).Top + (barra(0).Height / 2) < y Then
barra(0).Top = barra(0).Top + delim ''10
ElseIf barra(0).Top + (barra(0).Height / 2) > y Then
barra(0).Top = barra(0).Top - delim ''10
End If


End If
End Sub

Private Sub tmrGame_Timer()
'Label1 = isPC ' setX & " " & setY & " - " & X

If setX = True Then
x = (x + delimX) * ace
End If

If setX = False Then
x = (x - delimX) * ace

End If

If setY = True Then
y = (y + delimY) * ace

End If

If setY = False Then
y = (y - delimY) * ace

End If

pctBoard.Cls
BitBlt pctBoard.hDC, x, y, pctFrom.ScaleWidth, pctFrom.ScaleHeight, pctFrom.hDC, 0, 0, vbSrcCopy

'BitBlt Me.hdc, X * 2, Y * 2, pctFrom.ScaleWidth, pctFrom.ScaleHeight, pctFrom.hdc, 0, 0, vbSrcCopy






If x < barra(0).Width + 10 Then
isPC = True
zxc = Int(Rnd * 2)
'zxc = (Y < (pctBoard.Height / 2))

If (y > barra(0).Top) And (y < barra(0).Top + barra(0).Height) Then
temp = y - barra(0).Top
temp2 = (barra(0).Height / 3)

If temp < temp2 Then
setX = 1
setY = ((setY)) Xor zxc  ' 0
delimX = 10
delimY = 15
ace = 1

ElseIf temp < 2 * temp2 Then
setX = 1
setY = ((setY)) Xor zxc  ' 1
delimX = 15
delimY = 5
ace = 1.2

ElseIf temp < 3 * temp2 Then
setX = 1
setY = ((setY)) Xor zxc  '  0
delimX = 15
delimY = 10
ace = 1
End If

End If

ElseIf x > barra(1).Left - 16 Then
isPC = False
'zxc = (Y < (pctBoard.Height / 2))
zxc = Int(Rnd * 2)

If (y > barra(1).Top) And (y < barra(1).Top + barra(1).Height) Then
temp = y - barra(1).Top

temp2 = (barra(1).Height / 3)

If temp < temp2 Then
setX = 0
setY = ((setY)) Xor zxc  ' 1
delimX = 15
delimY = 10
ace = 1

ElseIf temp < 2 * temp2 Then
setX = 0
setY = ((setY)) Xor zxc   '1
delimX = 15
delimY = 5
ace = 1.2

ElseIf temp < 3 * temp2 Then
setX = 0
setY = ((setY)) Xor zxc  ' 0
delimX = 10
delimY = 15
ace = 1

End If


End If




End If






If y > (pctBoard.ScaleHeight - 10) Then
setY = False
lastY = y
End If

If y < 0 Then
setY = True
lastY = y
End If



If x < 0 Then
setX = False
x = pctBoard.ScaleWidth
y = Int(Rnd * pctBoard.ScaleHeight) + 1
delimX = 15
delimY = 5
player2 = player2 + 1
End If

If x > pctBoard.ScaleWidth + 20 Then
y = Int(Rnd * pctBoard.ScaleHeight) + 1
x = 0
delimX = 15
delimY = 5
setX = True
player1 = player1 + 1
End If

'' setX = False

End Sub

Private Sub tmrJoystick_Timer()
Dim JInfo As JOYINFO
GetJoystick JOYSTICKID1, JInfo

If JInfo.Buttons = 1 Then
barra(1).Top = barra(1).Top + 15
ElseIf JInfo.Buttons = 2 Then
barra(1).Top = barra(1).Top - 15
End If
   
'If JInfo.Y >= 34782 Then
''barra(1).Top = barra(1).Top + 15
'ElseIf JInfo.Y <= 28000 Then
'barra(1).Top = barra(1).Top - 15
'End If

End Sub
