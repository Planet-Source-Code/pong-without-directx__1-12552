Attribute VB_Name = "modGame"
Public Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal XSrc As Long, ByVal YSrc As Long, ByVal dwRop As Long) As Long

Const JOY_BUTTON1 = &H1
Const JOY_BUTTON2 = &H2
Const JOY_BUTTON3 = &H4
Const JOY_BUTTON4 = &H8
Const JOYERR_BASE = 160
Const JOYERR_NOERROR = (0)
Const JOYERR_NOCANDO = (JOYERR_BASE + 6)
Const JOYERR_PARMS = (JOYERR_BASE + 5)
Const JOYERR_UNPLUGGED = (JOYERR_BASE + 7)
Const MAXPNAMELEN = 32
Const JOYSTICKID1 = 0
Const JOYSTICKID2 = 1

Public Type JOYINFO
   x As Long
   y As Long
   z As Long
   Buttons As Long
End Type
Private Type JOYCAPS
   wMid As Integer
   wPid As Integer
   szPname As String * MAXPNAMELEN
   wXmin As Long
   wXmax As Long
   wYmin As Long
   wYmax As Long
   wZmin As Long
   wZmax As Long
   wNumButtons As Long
   wPeriodMin As Long
   wPeriodMax As Long
 End Type

Private Declare Function joyGetDevCaps Lib "winmm.dll" Alias "joyGetDevCapsA" (ByVal id As Long, lpCaps As JOYCAPS, ByVal uSize As Long) As Long
Private Declare Function joyGetNumDevs Lib "winmm.dll" () As Long
Private Declare Function joyGetPos Lib "winmm.dll" (ByVal uJoyID As Long, pji As JOYINFO) As Long
Public Function GetJoystick(ByVal joy As Integer, JI As JOYINFO) As Boolean
   If joyGetPos(joy, JI) <> JOYERR_NOERROR Then
      GetJoystick = False
   Else
      GetJoystick = True
   End If
End Function
'  If IsConnected is False then it returns the number of
'  joysticks the driver supports. (But may not be connected)
'
'  If IsConnected is True the it returns the number of
'  joysticks present and connected.
'
'  IsConnected is true by default.
Public Function IsJoyPresent(Optional IsConnected As Variant) As Long
   Dim ic As Boolean
   Dim i As Long
   Dim j As Long
   Dim ret As Long
   Dim JI As JOYINFO
   
   ic = IIf(IsMissing(IsConnected), True, CBool(IsConnected))

   i = joyGetNumDevs
   
   If ic Then
      j = 0
      Do While i > 0
         i = i - 1   'Joysticks id's are 0 and 1
         If joyGetPos(i, JI) = JOYERR_NOERROR Then
            j = j + 1
         End If
      Loop
   
      IsJoyPresent = j
   Else
      IsJoyPresent = i
   End If
   
End Function
'  Fills the ji structure with the minimum x, y, and z
'  coordinates.  Buttons is filled with the number of
'  buttons.
Public Function GetJoyMin(ByVal joy As Integer, JI As JOYINFO) As Boolean
   Dim jc As JOYCAPS
   
   If joyGetDevCaps(joy, jc, Len(jc)) <> JOYERR_NOERROR Then
      GetJoyMin = False
      
   Else
      JI.x = jc.wXmin
      JI.y = jc.wYmin
      JI.z = jc.wZmin
      JI.Buttons = jc.wNumButtons
   
      GetJoyMin = True
   End If
End Function
'  Fills the ji structure with the maximum x, y, and z
'  coordinates.  Buttons is filled with the number of
'  buttons.
Public Function GetJoyMax(ByVal joy As Integer, JI As JOYINFO) As Boolean
   Dim jc As JOYCAPS
   If joyGetDevCaps(joy, jc, Len(jc)) <> JOYERR_NOERROR Then
      GetJoyMax = False
   Else
      JI.x = jc.wXmax
      JI.y = jc.wYmax
      JI.z = jc.wZmax
      JI.Buttons = jc.wNumButtons
      GetJoyMax = True
   End If
End Function

Public Sub none()

    'KPD-Team 1999
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    Dim JInfo As JOYINFO
    'Clear the form
    'frmGame.Cls
    'Print the information to the form
    'frmGame.Print "Number of joysticks the driver supports:" + Str$(IsJoyPresent(False))
    'frmGame.Print "Number of connected joysticks:" + Str$(IsJoyPresent(True))
    GetJoystick JOYSTICKID1, JInfo
    'frmGame.Print "Number of buttons:" + Str$(JInfo.Buttons)
    
    Select Case JInfo.Buttons
    
        Case 0
        
        Case 1
        
    End Select
    
        
    'GetJoyMax JOYSTICKID1, JInfo
    'frmGame.Print "Max X:" + Str$(JInfo.X)
    'frmGame.Print "Max Y:" + Str$(JInfo.Y)
    'frmGame.Print "Max Z:" + Str$(JInfo.Z)
    'GetJoyMin JOYSTICKID1, JInfo
    'frmGame.Print "Min X:" + Str$(JInfo.X)
    'frmGame.Print "Min Y:" + Str$(JInfo.Y)
    'frmGame.Print "Min Z:" + Str$(JInfo.Z)
    
End Sub

    
