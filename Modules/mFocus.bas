Attribute VB_Name = "mFocus"
Option Explicit

Private Const SW_SHOWNORMAL = 1
Private Const SW_SHOWMINIMIZED = 2
Private Const SW_SHOWMAXIMIZED = 3
Private Const SW_SHOWNOACTIVATE = 4
Private Const SW_SHOW = 5

Private Type RECT
   Left    As Long
   Top     As Long
   Right   As Long
   Bottom  As Long
End Type

Private Type POINTAPI
   x       As Long
   y       As Long
End Type

Private Type WINDOWPLACEMENT
   length            As Long
   flags             As Long
   showCmd           As Long
   ptMinPosition     As POINTAPI
   ptMaxPosition     As POINTAPI
   rcNormalPosition  As RECT
End Type

Private Declare Function GetWindowPlacement Lib "user32" _
   (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Private Declare Function SetWindowPlacement Lib "user32" _
   (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long

Public Sub PreventFocus(hWndToActivate As Long)

    Dim currRect As RECT
    Dim currWinP As WINDOWPLACEMENT
  
    With currWinP
       .length = Len(currWinP)
       Call GetWindowPlacement(hWndToActivate, currWinP)

       .length = Len(currWinP)
       .flags = 0&
       .showCmd = SW_SHOWNOACTIVATE
    End With
    
    Call SetWindowPlacement(hWndToActivate, currWinP)
  
End Sub


