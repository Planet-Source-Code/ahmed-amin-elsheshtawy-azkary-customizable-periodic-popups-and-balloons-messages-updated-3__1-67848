Attribute VB_Name = "mMain"
Option Explicit

Public fMainForm As frmMain

'====================================================================
'====================================================================
Public Declare Function ShellExecute Lib "shell32.dll" _
            Alias "ShellExecuteA" _
            (ByVal hwnd As Long, _
            ByVal lpOperation As String, _
            ByVal lpFile As String, _
            ByVal lpParameters As String, _
            ByVal lpDirectory As String, _
            ByVal nShowCmd As Long) As Long
            

Public Declare Function LoadImage Lib "user32" Alias "LoadImageA" (ByVal hInst _
    As Long, ByVal lpsz As String, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 _
    As Long, ByVal un2 As Long) As Long
    
Public Declare Function GetActiveWindow Lib "user32" () As Long
Public Declare Function SetFocusAPI Lib "user32" Alias "SetFocus" (ByVal hwnd As Long) As Long

Public Declare Function ShowWindow Lib "user32" _
            (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

Public Const SW_SHOWNOACTIVATE = 4


Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Const SM_CXFULLSCREEN = 16   ' Width of window client area
Public Const SM_CYFULLSCREEN = 17   ' Height of window client area

'====================================================================
Public AzkarFile As String
Public AzkarsList() As String
Public AzkarCount As Long

Public Type Azkars
    ID As Long
    Zekr As String
    Period As Double
    PopupWidth As Long
    PopupHeight As Long
    PopupPosition As Long
    Enabled As Long
    ShowPopup As Long
    ShowBalloon As Long
    
    BalloonColor As Long
    BalloonBackColor As Long
    
    PlaySound As Long
    Animation As Long
    AnimationTime As Long
    ShowDelay As Long
    Transparency As Long
    Theme As String
    Alignment As Long
    SoundFile As String
    
    FontName As String
    FontSize  As Long
    FontColor  As Long
    FontBold  As Long
    FontItalic  As Long
    FontUnderline  As Long
End Type

Public Azkars() As Azkars
Public PopupsCollection As New Collection
Public PopupsCount As Long
Public Periods() As Long
Public ScreenWidth As Long
Public ScreenHeight As Long
Public LanguageFile As String
Public Languages As String
Public DefaultLanguage  As String
Public LanguageDirection As String
Public bLoaded As Boolean
Public NewSelectedImage As String
Public RunningStatus As Boolean
Public MenuTrayEnable As String
Public MenuTrayDisable As String
Public MenuTrayShow As String
Public MenuTrayHide As String
Public MenuTrayExit As String
Public MenuTrayAzkar As String
Public ExitMessage As String
Public AlignmentLeft As String
Public AlignmentRight As String
Public AlignmentCenter As String
Public AzkarStatusLabel  As String
Public MinutesLabel As String
Public RotateAzkarLabel As String
Public RotationTimeLabel As String
Public LabelsColorLabel As String
Public LabelsColor As String
Public RotateAzkar As Long
Public RotateAzkarPeriod As Double
Public RotateAzkarIndex As Long

Public AzkarStatus As Long

Public CounterLabel As String
Public ResetCounter As String
Public CounterUp As String
    
Public Counter100Label As String
Public CounterReset100 As String
Public CounterUp100 As String
Public AzkarIndex As String
Public AzkarClose  As String
Public CommandOK As String
Public CommandCancel As String
Public AutoStartButton As String
Public AutoStopButton As String
Public AutoStartMessage As String

Public AutoStartUp As Long
'====================================================================
'====================================================================
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_CREATEDIBSECTION = &H2000
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADTRANSPARENT = &H20
Public Const LR_COPYRETURNORG = &H4
Public Const IMAGE_BITMAP = 0
Public Const IMAGE_ICON = 1

Private Const ILC_COLOR = &H0
Private Const ILC_MASK = &H1
Public Const ILC_COLOR4 = &H4
Public Const ILC_COLOR8 = &H8
Public Const ILC_COLOR16 = &H10
Public Const ILC_COLOR24 = &H18
Public Const ILC_COLOR32 = &H20
Public Const ILD_NORMAL = 0


Public Const ID_Left_Pane As Long = 1
Public Const ID_Right_Pane As Long = 2
'Public Const ID_Bottom_Pane As Long = 3
Public Const ID_Status_Pane As Long = 3
Public Const ID_Settings_Pane As Long = 4
Public Const ID_Options_Pane As Long = 5
Public Const ID_Tools_Pane As Long = 6
Public Const ID_Registeration_Pane As Long = 7
Public Const ID_Help_Pane As Long = 8
'====================================================================
'Global Declaration
Global Const gAppName = "Azkary"
Public Const AppRegPath = "Islamware\Azkary"
Public Const AppRegSettingsSection = "Settings"

'====================================================================

'====================================================================
'====================================================================
Sub Main()
    Set fMainForm = New frmMain
    fMainForm.Show
End Sub
'====================================================================
'====================================================================

'====================================================================
