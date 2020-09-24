Attribute VB_Name = "mTransparent"
Option Explicit

Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRECT As RECT) As Long
Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRECT As RECT) As Long
Declare Function CombineRgn Lib "gdi32" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Declare Function ScreenToClient Lib "user32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Boolean) As Long
    Public Const RGN_AND = 1
    Public Const RGN_COPY = 5
    Public Const RGN_DIFF = 4
    Public Const RGN_OR = 2
    Public Const RGN_XOR = 3

Type POINTAPI
    x As Long
    Y As Long
    End Type
Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    End Type

Public Sub MakeTransparent(frm As Form)

    'This code was takin from a AOL Visual Basic
    'Message Board. It was submited by: SOOPRcow
    
    Dim rctClient As RECT, rctFrame As RECT
    Dim hClient As Long, hFrame As Long
    
    '// Grab client area and frame area
    GetWindowRect frm.hWnd, rctFrame
    GetClientRect frm.hWnd, rctClient
    
    '// Convert client coordinates to screen
    '     coordinates
    Dim lpTL As POINTAPI, lpBR As POINTAPI
    
    lpTL.x = rctFrame.Left
    lpTL.Y = rctFrame.Top
    lpBR.x = rctFrame.Right
    lpBR.Y = rctFrame.Bottom
    
    ScreenToClient frm.hWnd, lpTL
    ScreenToClient frm.hWnd, lpBR
    
    rctFrame.Left = lpTL.x
    rctFrame.Top = lpTL.Y
    rctFrame.Right = lpBR.x
    rctFrame.Bottom = lpBR.Y
    
    rctClient.Left = Abs(rctFrame.Left)
    rctClient.Top = Abs(rctFrame.Top)
    rctClient.Right = rctClient.Right + Abs(rctFrame.Left)
    rctClient.Bottom = rctClient.Bottom + Abs(rctFrame.Top)
    
    rctFrame.Right = rctFrame.Right + Abs(rctFrame.Left)
    rctFrame.Bottom = rctFrame.Bottom + Abs(rctFrame.Top)
    rctFrame.Left = 0
    rctFrame.Top = 0
    
    '// Convert RECT structures to region handles
    
    hClient = CreateRectRgn(rctClient.Left, rctClient.Top, rctClient.Right, rctClient.Bottom)
    hFrame = CreateRectRgn(rctFrame.Left, rctFrame.Top, rctFrame.Right, rctFrame.Bottom)
    
    '// Create the new "Transparent" region
    CombineRgn hFrame, hClient, hFrame, RGN_XOR
    
    '// Now lock the window's area to this created region
    SetWindowRgn frm.hWnd, hFrame, True
    
End Sub


