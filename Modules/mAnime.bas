Attribute VB_Name = "mdlAnime"

Option Explicit

'----------------------------------------------------------------------------------------------------------------------------------------------------------
' Source Code   : AnimateForm
' Auther        : Jim Jose
' eMail         : jimjosev33@yahoo.com
' Purpose       : Cooool flash style animations in Vb
' Comment       : Function contains 13 effects, each have
'               : reverse effect too. So total 26 animations in one function
'               : Completly error checked and free from memory leaks
' Copyright Jim Jose, Gtech Creations - 2005
'----------------------------------------------------------------------------------------------------------------------------------------------------------

'[APIs]
Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function CombineRgn Lib "gdi32.dll" (ByVal hDestRgn As Long, ByVal hSrcRgn1 As Long, ByVal hSrcRgn2 As Long, ByVal nCombineMode As Long) As Long
Private Declare Sub SleepAPI Lib "kernel32.dll" (ByVal dwMilliseconds As Long)

Public Declare Function GetTickCount Lib "kernel32.dll" () As Long

'[Event Enum]
Public Enum AnimeEventEnum
    aUnload = 0
    aload = 1
End Enum

'[Effect Enum]
Public Enum AnimeEffectEnum
    eAppearFromLeft = 0
    eAppearFromRight = 1
    eAppearFromTop = 2
    eAppearFromBottom = 3
    eGenerateLeftTop = 4
    eGenerateLeftBottom = 5
    eGenerateRightTop = 6
    eGenerateRightBottom = 7
    eStrechHorizontally = 8
    eStrechVertically = 9
    eZoomOut = 10
    eFoldOut = 11
    eCurtonHorizontal = 12
    eCurtonVertical = 13
End Enum

'[Constants]
Private Const RGN_AND As Long = 1
Private Const RGN_OR As Long = 2
Private Const RGN_XOR As Long = 3
Private Const RGN_COPY As Long = 5
Private Const RGN_DIFF As Long = 4

Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Const SWP_NOMOVE = &H2
Private Const SWP_NOSIZE = &H1
Private Const SWP_NOACTIVATE = &H10
Private Const SWP_SHOWWINDOW = &H40
Private Const TOPMOST_FLAGS = SWP_NOMOVE Or SWP_NOSIZE

Private Const SWP_FLAGS = SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE


Private Declare Function SetWindowPos Lib "user32" _
                                    (ByVal hwnd As Long, _
                                     ByVal hWndInsertAfter As Long, _
                                     ByVal x As Long, y, _
                                     ByVal cx As Long, _
                                     ByVal cy As Long, _
                                     ByVal wFlags As Long) As Long

' Sub to sleep x milli seconds
Public Sub Sleep(lngSleep As Long)
   Dim lngSleepEnd As Long
   'GetTickCount: Retrieves the number of milliseconds that have elapsed since the system was started, up to 49.7 days.
   lngSleepEnd = GetTickCount + lngSleep '* 1000
   While GetTickCount <= lngSleepEnd
      DoEvents
   Wend
End Sub


'-------------------------------------------------------------------------
' Procedure  : AnimateForm
' Auther     : Jim Jose
' Input      : AnimeObject, Event , Effect + Time/frame values
' OutPut     : None
' Purpose    : Cooool flash style animations in Vb
'-------------------------------------------------------------------------

Public Function AnimateForm(hwndObject As Object, ByVal aEvent As AnimeEventEnum, _
                            Optional ByVal aEffect As AnimeEffectEnum = 11, _
                            Optional ByVal FrameTime As Long = 1, _
                            Optional ByVal FrameCount As Long = 33) As Boolean
On Error GoTo Handle
Dim X1 As Long, Y1 As Long
Dim hRgn As Long, tmpRgn As Long
Dim XValue As Long, YValue As Long
Dim XIncr As Double, YIncr As Double
Dim wHeight As Long, wWidth As Long

    wWidth = hwndObject.Width / Screen.TwipsPerPixelX
    wHeight = hwndObject.Height / Screen.TwipsPerPixelY
    
    hwndObject.Visible = True
    
    Select Case aEffect
    
        Case eAppearFromLeft
        
            XIncr = wWidth / FrameCount
            For X1 = 0 To FrameCount
            
                ' Define the size of current frame/Create it
                XValue = X1 * XIncr
                hRgn = CreateRectRgn(0, 0, XValue, wHeight)
                
                ' If unload then take the reverse(anti) region
                If aEvent = aUnload Then
                    tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
                    CombineRgn hRgn, hRgn, tmpRgn, RGN_XOR
                    DeleteObject tmpRgn
                End If
                                
                ' Set the new region for the window
                SetWindowRgn hwndObject.hwnd, hRgn, True
                DoEvents
                                               
                Sleep FrameTime
               
            Next X1
            
        Case eAppearFromRight
        
            XIncr = wWidth / FrameCount
            For X1 = 0 To FrameCount
                
                ' Define the size of current frame/Create it
                XValue = wWidth - X1 * XIncr
                hRgn = CreateRectRgn(XValue, 0, wWidth, wHeight)
                
                ' If unload then take the reverse(anti) region
                If aEvent = aUnload Then
                    tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
                    CombineRgn hRgn, hRgn, tmpRgn, RGN_XOR
                    DeleteObject tmpRgn
                End If
                
                ' Set the new region for the window
                SetWindowRgn hwndObject.hwnd, hRgn, True:  DoEvents
                Sleep FrameTime
                
            Next X1
            
        Case eAppearFromTop
        
            YIncr = wHeight / FrameCount
            For Y1 = 0 To FrameCount
            
                ' Define the size of current frame/Create it
                YValue = Y1 * YIncr
                hRgn = CreateRectRgn(0, 0, wWidth, YValue)
                
                ' If unload then take the reverse(anti) region
                If aEvent = aUnload Then
                    tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
                    CombineRgn hRgn, hRgn, tmpRgn, RGN_XOR
                    DeleteObject tmpRgn
                End If
                
                ' Set the new region for the window
                SetWindowRgn hwndObject.hwnd, hRgn, True:   DoEvents
                Sleep FrameTime
                
            Next Y1
            
        Case eAppearFromBottom
        
            YIncr = wHeight / FrameCount
            For Y1 = 0 To FrameCount
            
                ' Define the size of current frame/Create it
                YValue = wHeight - Y1 * YIncr
                hRgn = CreateRectRgn(0, YValue, wWidth, wHeight)
                
                ' If unload then take the reverse(anti) region
                If aEvent = aUnload Then
                    tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
                    CombineRgn hRgn, hRgn, tmpRgn, RGN_XOR
                    DeleteObject tmpRgn
                End If
                
                ' Set the new region for the window
                SetWindowRgn hwndObject.hwnd, hRgn, True: DoEvents
                Sleep FrameTime
                
            Next Y1
            
        Case eGenerateLeftTop
        
            XIncr = wWidth / FrameCount: YIncr = wHeight / FrameCount
            For X1 = 0 To FrameCount
                
                ' Define / Create Region
                If aEvent = aload Then XValue = X1 * XIncr: YValue = X1 * YIncr Else XValue = wWidth - X1 * XIncr: YValue = wHeight - X1 * YIncr
                hRgn = CreateRectRgn(0, 0, XValue, YValue)

                ' Set the new region for the window
                SetWindowRgn hwndObject.hwnd, hRgn, True:   DoEvents
                Sleep FrameTime
            
            Next X1
            
        Case eGenerateLeftBottom
        
            XIncr = wWidth / FrameCount: YIncr = wHeight / FrameCount
            For X1 = 0 To FrameCount
            
                ' Define / Create Region
                If aEvent = aload Then XValue = X1 * XIncr: YValue = wHeight - X1 * YIncr Else XValue = wWidth - X1 * XIncr: YValue = X1 * YIncr
                hRgn = CreateRectRgn(0, wHeight, XValue, YValue)
                
                ' Set the new region for the window
                SetWindowRgn hwndObject.hwnd, hRgn, True: DoEvents
                Sleep FrameTime
                
            Next X1
            
        Case eGenerateRightTop
        
            XIncr = wWidth / FrameCount: YIncr = wHeight / FrameCount
            For X1 = 0 To FrameCount
            
                ' Define / Create Region
                If aEvent = aload Then XValue = wWidth - X1 * XIncr: YValue = X1 * YIncr Else XValue = X1 * XIncr: YValue = wHeight - X1 * YIncr
                hRgn = CreateRectRgn(XValue, YValue, wWidth, 0)
                
                ' Set the new region for the window
                SetWindowRgn hwndObject.hwnd, hRgn, True: DoEvents
                Sleep FrameTime
                
            Next X1
            
        Case eGenerateRightBottom
        
            XIncr = wWidth / FrameCount: YIncr = wHeight / FrameCount
            For X1 = 0 To FrameCount
                
                ' Define / Create Region
                If aEvent = aload Then XValue = wWidth - X1 * XIncr: YValue = wHeight - X1 * YIncr Else XValue = X1 * XIncr: YValue = X1 * YIncr
                hRgn = CreateRectRgn(XValue, YValue, wWidth, wHeight)
                
                ' Set the new region for the window
                SetWindowRgn hwndObject.hwnd, hRgn, True: DoEvents
                Sleep FrameTime
                
            Next X1
            
        Case eStrechHorizontally
        
            XIncr = wWidth / FrameCount
            For X1 = 0 To FrameCount
            
                ' Define / Create Region
                If aEvent = aload Then XValue = wWidth - X1 * XIncr Else XValue = X1 * XIncr
                hRgn = CreateRectRgn(XValue / 2, 0, wWidth - XValue / 2, wHeight)
                
                ' Set the new region for the window
                SetWindowRgn hwndObject.hwnd, hRgn, True: DoEvents
                Sleep FrameTime
                
            Next X1
            
        Case eStrechVertically
        
            YIncr = wHeight / FrameCount
            For Y1 = 0 To FrameCount
            
                ' Define / Create Region
                If aEvent = aload Then YValue = Y1 * YIncr Else YValue = wHeight - Y1 * YIncr
                hRgn = CreateRectRgn(0, wHeight / 2 - YValue / 2, wWidth, wHeight / 2 + YValue / 2)
                
                ' Set the new region for the window
                SetWindowRgn hwndObject.hwnd, hRgn, True: DoEvents
                Sleep FrameTime
                
            Next Y1
            
        Case eZoomOut
        
            XIncr = wWidth / FrameCount: YIncr = wHeight / FrameCount
            For X1 = 0 To FrameCount
            
                ' Define / Create Region
                If aEvent = aload Then XValue = X1 * XIncr: YValue = X1 * YIncr Else XValue = wWidth - X1 * XIncr: YValue = wHeight - X1 * YIncr
                hRgn = CreateRectRgn((wWidth - XValue) / 2, (wHeight - YValue) / 2, (wWidth + XValue) / 2, (wHeight + YValue) / 2)
                
                ' Set the new region for the window
                SetWindowRgn hwndObject.hwnd, hRgn, True: DoEvents
                Sleep FrameTime
                
            Next X1
            
        Case eFoldOut
        
            If hwndObject.Width >= hwndObject.Height Then XIncr = wWidth / FrameCount: YIncr = wWidth / FrameCount Else XIncr = wHeight / FrameCount: YIncr = wHeight / FrameCount
            For X1 = 0 To FrameCount
            
                ' Define / Create Region
                If aEvent = aload Then XValue = X1 * XIncr: YValue = X1 * YIncr Else XValue = wWidth - X1 * XIncr: YValue = wHeight - X1 * YIncr
                hRgn = CreateRectRgn((wWidth - XValue) / 2, (wHeight - YValue) / 2, (wWidth + XValue) / 2, (wHeight + YValue) / 2)
                
                ' Set the new region for the window
                SetWindowRgn hwndObject.hwnd, hRgn, True: DoEvents
                Sleep FrameTime
                
            Next X1
            
        Case eCurtonHorizontal
        
            Dim ScanWidth As Long
            ScanWidth = FrameCount / 2
            For Y1 = 0 To FrameCount / 2
                
                ' Initiate region
                hRgn = CreateRectRgn(0, 0, 0, 0)
                For X1 = 0 To wHeight / FrameCount * 2
                    ' Create each curton region
                    tmpRgn = CreateRectRgn(0, X1 * ScanWidth, wWidth, X1 * ScanWidth + Y1)
                    CombineRgn hRgn, hRgn, tmpRgn, RGN_OR
                    DeleteObject tmpRgn
                Next X1
                
                ' If unload then take the reverse(anti) region
                If aEvent = aUnload Then
                    tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
                    CombineRgn hRgn, hRgn, tmpRgn, RGN_XOR
                    DeleteObject tmpRgn
                End If
                
                ' Set the new region for the window
                SetWindowRgn hwndObject.hwnd, hRgn, True: DoEvents
                Sleep FrameTime
                
            Next Y1
            
        Case eCurtonVertical
        
            ScanWidth = FrameCount / 2
            For X1 = 0 To FrameCount / 2
            
                ' Initiate Region
                hRgn = CreateRectRgn(0, 0, 0, 0)
                For Y1 = 0 To wWidth / FrameCount * 2
                    ' Create each curton region
                    tmpRgn = CreateRectRgn(Y1 * ScanWidth, 0, Y1 * ScanWidth + X1, wHeight)
                    CombineRgn hRgn, hRgn, tmpRgn, RGN_OR
                    DeleteObject tmpRgn
                Next Y1
                
                ' If unload then take the reverse(anti) region
                If aEvent = aUnload Then
                    tmpRgn = CreateRectRgn(0, 0, wWidth, wHeight)
                    CombineRgn hRgn, hRgn, tmpRgn, RGN_XOR
                    DeleteObject tmpRgn
                End If
                
                ' Set the new region for the window
                SetWindowRgn hwndObject.hwnd, hRgn, True: DoEvents
                Sleep FrameTime
            Next X1
            
    End Select

    AnimateForm = True
    
Exit Function
Handle:
    AnimateForm = False
End Function

Public Sub MakeTopMostNoFocus(hwnd As Long)
    'SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, TOPMOST_FLAGS
    SetWindowPos hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_FLAGS
End Sub

