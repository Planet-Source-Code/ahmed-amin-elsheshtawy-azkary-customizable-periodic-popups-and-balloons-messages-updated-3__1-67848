Attribute VB_Name = "mFadeForm"
Option Explicit

Const LWA_COLORKEY = &H1
Const LWA_ALPHA = &H2
Const GWL_EXSTYLE = (-20)
'Const WS_EX_LAYERED = &H80000

'WS_EX_LAYERED - Tells Windows to combine this window with other windows beneath it using the specified alpha functions.
'WS_EX_TRANSPARENT - Makes the Window transparent to the mouse.
'WS_EX_TOOLWINDOW - Ensure it doesn't appear in the taskbar.
'WS_EX_TOPMOST - Display window on top.
Const WS_EX_TOPMOST As Long = &H8&
Const WS_EX_TRANSPARENT  As Long = &H20&
Const WS_EX_TOOLWINDOW As Long = &H80&
Const WS_EX_LAYERED As Long = &H80000
Const WS_POPUP = &H80000000
Const WS_VISIBLE = &H10000000
Const SPI_GETSELECTIONFADE As Long = &H1014&

Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long

Public Sub FadeForm(formHwnd As Long, Amount As Long)
    
    'Fad In/Out W2K, XP:
    'KPD-Team 2000
    'URL: http://www.allapi.net/
    'E-Mail: KPDTeam@Allapi.net
    Dim Ret As Long
    
    On Error GoTo NotSupported
    
    'Set the window style to 'Layered'
    Ret = GetWindowLong(formHwnd, GWL_EXSTYLE)
    
    Ret = Ret Or WS_EX_LAYERED Or WS_EX_TRANSPARENT Or WS_EX_TOOLWINDOW
    
    SetWindowLong formHwnd, GWL_EXSTYLE, Ret
    
    'Set the opacity of the layered window to 128
    SetLayeredWindowAttributes formHwnd, 0, Amount, LWA_ALPHA
    Exit Sub
NotSupported:
End Sub

