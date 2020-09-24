Attribute VB_Name = "mAnimateForms"
Option Explicit


'**************************************
'Windows API/Global Declarations for :Zo
'     om Open/Close with DrawAnimatedRects
'**************************************
'**************************************
' Name: Zoom Open/Close with DrawAnimate
'     dRects
' Description:Using DrawAnimatedRects, w
'     hich appears to be the same function Win


'     dows uses when Minimizing or Restoring a
    '     window, you can zoom the title bar of a


    '     form to or from any of the four corners
        '     of the screen, to or from an object, or
        '     to or from any Rectangle. The effect is
        '     that the form appears to originate from
        '     a corner, from an object, or from any de
        '     fined rectangle, and returns there when
        '     dismissed. Plug this module into any pro
        '     ject, and then you can zoom with one lin
        '     e of code.
' By: James Greene
'
'
' Inputs:3 functions called in much the
'     same way


'for any of the 3, you'll need to provide
'    1) the handle of the window you want to zoom
'    2) a second rectangle, which can be derived by specifying either a corner of the screen, the hWnd of an object, or directly, by providing a RECT structure
'    3)the direction you'd like to zoom, either from the Form's rectangle to the rectangle derived from argument 2, or vice-versa
'
' Returns:All Functions return True on S
'     uccess
'
'Assumes:1)you should set StartUpPositio
'     n to Manual for your form, and 2)use Mov
'     e to center it on the screen
'3)works better if you set Visible = False
'then manually set Visible to True after calling it
'4)because VB doesn't let you have control over sizing or placing of MDI children, this is not recommended for them
'
'Side Effects:no side effects, but see a
'     bove
'This code is copyrighted and has limite
'     d warranties.
'Please see http://www.Planet-Source-Cod
'     e.com/xq/ASP/txtCodeId.5757/lngWId.1/qx/
'     vb/scripts/ShowCode.htm
'for details.
'**************************************

' Author:
' James Greene
' jgreene@esper.com
' 1/28/2000
' this module is a wrapper around DrawAn
'     imatedRects
' gives the illusion that an opening or
'     closing window
' is coming out of, or returning to,
' one of the four corners of the screen
' for another good example of using this
'     stuff
' go to http://www.zonecorp.com/VB5/
' (Ramon Guerrero's VB5 Net Stop)
' and have a look at the office assistan
'     t sample
' call ZoomFrm on Form_Load or _Unload,
'     or with Hide or Show
' set bClose = True when Unloading or Hi
'     ding
' see more notes in ZoomFrm
' NOTE: I've set these declares to publi
'     c for maximum flexibility
' for maximum efficiency, set them to pr
'     ivate


Public Declare Function DrawAnimatedRects Lib "user32" _
    (ByVal hWnd As Long, ByVal idAni As Long, lprcFrom As RECT, _
    lprcTo As RECT) As Long


Public Declare Function GetWindowRect Lib "user32" _
    (ByVal hWnd As Long, lpRect As RECT) As Long
    ' idAni Constants
    Public Const IDANI_OPEN = &H1
    Public Const IDANI_CLOSE = &H2
    Public Const IDANI_CAPTION = &H3


Public Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
    End Type
    ' leave this public
    ' specify the corner of the screen you w
    '     ant the window to come from/ go to


Public Enum AniZoomConstants
    TopLeftZoom = 0 ' TopLeft
    TopRightZoom = 1 ' TopRight
    BottomLeftZoom = 2 ' BottomLeft
    BottomRightZoom = 3 ' BottomRight
End Enum


Public Sub PositionRect(RC As RECT, lLeft&, lTop&, lHeight&, lWidth&, Optional iScaleMode As ScaleModeConstants = vbPixels)
    ' fills out the RECT structure RC
    ' last two args are Height and Width, no
    '     t Bottom and Top
    ' so it's much like using the VB Move fu
    '     nction...
    ' IMPORTANT: this is in pixels, unless y
    '     ou set iScaleMode to vbTwips
    ' if scalemode is twips, divide all para
    '     maters by TwipsPerPixelX or Y


    If iScaleMode = vbTwips Then
        lLeft = lLeft / Screen.TwipsPerPixelX
        lWidth = lWidth / Screen.TwipsPerPixelX
        lTop = lTop / Screen.TwipsPerPixelY
        lHeight = lHeight / Screen.TwipsPerPixelY
    End If
    RC.Left = lLeft: RC.Top = lTop
    RC.Bottom = RC.Top + lHeight: RC.Right = RC.Left + lWidth
End Sub


Public Function ZoomFrm(frmhWnd&, _
    Optional ScreenCorner As AniZoomConstants = TopLeftZoom, _
    Optional bClose As Boolean = False) As Boolean
    Dim rcFrom As RECT, rcTo As RECT
    Dim Ret& ' return value
    Dim AniFlags& ' for OR ' ing the flags
    Dim lLeft&, lTop& ' left and top of a rect
    Dim lScreenHeight&, lScreenWidth& ' / TwipsPerPixelX, Y
    ' makes the zoom-open or zoom-closed ani
    '     mated effect
    ' set bClose = True if you're closing th
    '     e Form!!!!
    ' call this last thing from FORM_Load an
    '     d/or FORM_Unload (bClose = True)
    ' if it doesn't look smooth, set your fo
    '     rm's Visible Property to False before ca
    '     lling this...
    ' when OPENING the form, set Visible = T
    '     rue afterwards
    ' this will work for MDI child windows,
    '     but there's a slight catch
    ' as the window is pulled into position.
    '     ...
    If frmhWnd = 0 Then Exit Function
    ' convert Screen.Height and Width to pix
    '     els (and subtract 1 from each)
    lScreenWidth = (Screen.Width / Screen.TwipsPerPixelX) - 1
    lScreenHeight = (Screen.Height / Screen.TwipsPerPixelY) - 1
    ' set the left and the top point, so we
    '     can position the rectangle below
    ' coulda used PointAPI here just as well
    '     ...


    Select Case ScreenCorner
        Case TopLeftZoom
        lLeft = 0: lTop = 0
        Case TopRightZoom
        lLeft = lScreenWidth: lTop = 0
        Case BottomLeftZoom
        lLeft = 0: lTop = lScreenHeight
        Case BottomRightZoom
        lLeft = lScreenWidth: lTop = lScreenHeight
    End Select


If Not bClose Then
    ' opening the form
    PositionRect rcFrom, lLeft, lTop, 1, 1
    GetWindowRect frmhWnd, rcTo
    AniFlags = IDANI_OPEN
Else
    ' closing the form
    PositionRect rcTo, lLeft, lTop, 1, 1
    GetWindowRect frmhWnd, rcFrom
    AniFlags = IDANI_CLOSE
End If
' haven't got this to work without the I
'     DANI_CAPTION bit set!
AniFlags = AniFlags Or IDANI_CAPTION
Ret = DrawAnimatedRects(frmhWnd, AniFlags, rcFrom, rcTo)
' returns True on success
If Ret <> 0 Then ZoomFrm = True
End Function


Public Function ZoomOBJ(frmhWnd&, ObjhWnd&, _
    Optional bZoomToObject As Boolean = True) As Boolean
    ' zoom form to/from any object with an h
    '     Wnd property
    ' bZoomToObject = True: Zoom from form's
    '     Rectangle to the object's rectangle
    ' bZoomToObject = False: Zoom from objec
    '     t's Rectangle to the form's rectangle
    Dim rcFrom As RECT, rcTo As RECT
    Dim Ret& ' return value
    Dim AniFlags& ' for OR ' ing the flags
    Dim lLeft&, lTop& ' left and top of a rect
    ' set bClose = True if you're closing th
    '     e Form!!!!
    ' gotta have an hWnd
    If frmhWnd = 0 Or ObjhWnd = 0 Then Exit Function
    ' in this case, no difference between ID
    '     ANI_OPEN or IDANI_CLOSE
    ' and the IDANI_CAPTION bit has to be se
    '     t, it seems
    AniFlags = IDANI_OPEN Or IDANI_CAPTION


    If Not bZoomToObject Then
        ' zooming from the object to the window
        GetWindowRect ObjhWnd, rcFrom
        GetWindowRect frmhWnd, rcTo
    Else
        ' zooming from the window to the object
        GetWindowRect frmhWnd, rcFrom
        GetWindowRect ObjhWnd, rcTo
    End If
    Ret = DrawAnimatedRects(frmhWnd, AniFlags, rcFrom, rcTo)
    ' returns True on success
    If Ret <> 0 Then ZoomOBJ = True
End Function


Public Function ZoomRECT(frmhWnd&, RC As RECT, _
    Optional bZoomToRECT As Boolean = True) As Boolean
    ' zoom to/from any rectangle
    ' use PositionRECT, if you want, to fill
    '     out the RECT structure
    ' or use GetWindowRect and pass it an hW
    '     nd and a Rect Structure
    ' bZoomToRECT = True ' zoom from form's
    '     rect to the rect RC
    ' bZoomToRECT = False ' zoom from the re
    '     ctangle RC to the form's rectangle
    ' gotta have an hWnd
    If frmhWnd = 0 Then Exit Function
    Dim AniFlags&, Ret&
    Dim rcFORM As RECT
    ' in this case, no difference between ID
    '     ANI_OPEN or IDANI_CLOSE
    ' and the IDANI_CAPTION bit has to be se
    '     t, it seems
    AniFlags = IDANI_OPEN Or IDANI_CAPTION
    GetWindowRect frmhWnd, rcFORM


    If Not bZoomToRECT Then
        ' zooming from the rect to the window
        Ret = DrawAnimatedRects(frmhWnd, AniFlags, RC, rcFORM)
    Else
        ' zooming from the window to the rect
        Ret = DrawAnimatedRects(frmhWnd, AniFlags, rcFORM, RC)
    End If
    ' returns True on success
    If Ret <> 0 Then ZoomRECT = True
End Function

        


